"""
openai-batch/poll_batches.py

GitHub Actions에서 실행. 카드뉴스 번역 파이프라인
(bonus_model_interfaces.submit_translation_batch)이 제출해둔 OpenAI Batch API
작업 중 완료된 것을 찾아 결과를 각 채널 워크북의 TRANSLATE_DATA 탭에 채워넣는다.

트리거는 두 가지: (1) submit_translation_batch가 제출 직후 배치 완료를 직접
감지해 즉시 쏘는 workflow_dispatch(평소엔 이걸로 2~3분 내 처리), (2) 1시간
간격 schedule cron - (1)이 놓친 경우(GUI 종료 등)를 위한 저빈도 백업(2026-09-11).

- 대상 판별: status=="completed" AND metadata.contents_sheet_id가 있는 배치만
  (이 프로젝트가 제출한 배치는 전부 submit_translation_batch가 이 키를 채워서 만듦 -
  같은 OpenAI 계정의 다른 용도 배치와 섞이지 않게 하는 마커).
- 최근 3일 이내 생성된 배치만 훑는다(Batch API completion_window가 최대 24h라
  그보다 오래된 건 이미 completed거나 expired 상태로 끝났을 것 - 매 실행마다 전체
  이력을 다시 훑지 않기 위한 상한).
- 멱등성: TRANSLATE_DATA의 대상 셀이 이미 채워져 있으면 건너뛴다. 실제로 새로 쓴 값이
  하나도 없으면 시트 업로드 자체를 생략(불필요한 Drive 쓰기 방지).
"""
import json
import os
import time
import uuid

import requests
from openpyxl import load_workbook
from openai import OpenAI

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
REQUEST_TIMEOUT = 30
HERE = os.path.dirname(os.path.abspath(__file__))
MAX_BATCH_AGE_SEC = 3 * 24 * 3600

GOOGLE_CLIENT_ID = os.environ["GOOGLE_CLIENT_ID"]
GOOGLE_CLIENT_SECRET = os.environ["GOOGLE_CLIENT_SECRET"]
GOOGLE_REFRESH_TOKEN = os.environ["GOOGLE_REFRESH_TOKEN"]

_access_token = None


def _get_access_token(force_refresh=False):
    global _access_token
    if _access_token and not force_refresh:
        return _access_token
    r = requests.post("https://oauth2.googleapis.com/token", data={
        "client_id": GOOGLE_CLIENT_ID,
        "client_secret": GOOGLE_CLIENT_SECRET,
        "refresh_token": GOOGLE_REFRESH_TOKEN,
        "grant_type": "refresh_token",
    }, timeout=REQUEST_TIMEOUT)
    r.raise_for_status()
    _access_token = r.json()["access_token"]
    return _access_token


def _request_with_retry(method, url, **kwargs):
    kwargs.setdefault("timeout", REQUEST_TIMEOUT)
    headers = kwargs.pop("headers", {})
    headers["Authorization"] = f"Bearer {_get_access_token()}"

    for attempt in range(1, 4):
        try:
            resp = requests.request(method, url, headers=headers, **kwargs)
        except (requests.exceptions.ConnectionError, requests.exceptions.SSLError, requests.exceptions.Timeout):
            if attempt == 3:
                raise
            time.sleep(2 * attempt)
            continue
        if resp.status_code == 401:
            headers["Authorization"] = f"Bearer {_get_access_token(force_refresh=True)}"
            continue
        if resp.status_code in (502, 503, 504) and attempt < 3:
            time.sleep(2 * attempt)
            continue
        resp.raise_for_status()
        return resp
    resp.raise_for_status()
    return resp


def download_workbook(file_id):
    r = _request_with_retry(
        "GET", f"https://www.googleapis.com/drive/v3/files/{file_id}/export",
        params={"mimeType": XLSX_MIME},
    )
    xlsx_path = os.path.join(HERE, f"_poll_io_{file_id}_{uuid.uuid4().hex}.xlsx")
    with open(xlsx_path, "wb") as f:
        f.write(r.content)
    return load_workbook(xlsx_path), xlsx_path


def upload_workbook(wb, xlsx_path, file_id):
    wb.save(xlsx_path)
    try:
        with open(xlsx_path, "rb") as f:
            content = f.read()
        _request_with_retry(
            "PATCH", f"https://www.googleapis.com/upload/drive/v3/files/{file_id}?uploadType=media",
            headers={"Content-Type": XLSX_MIME}, data=content,
        )
    finally:
        os.remove(xlsx_path)


def ws_read_dicts(ws):
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return [], []
    headers = list(rows[0])
    out = []
    for row in rows[1:]:
        if all(v is None for v in row):
            continue
        out.append({headers[i]: row[i] for i in range(len(headers)) if i < len(row)})
    return headers, out


def ws_write_dicts(ws, headers, dict_rows):
    if ws.max_row >= 2:
        ws.delete_rows(2, ws.max_row - 1)
    for row in dict_rows:
        ws.append([row.get(h, "") for h in headers])


def upsert_translate_data(contents_sheet_id, updates):
    """updates: [(file_id, column, value), ...]. pipeline/bonus_ui.py의
    persist_fields_batch와 동일한 upsert 규칙(행이 없으면 새로 만든다) +
    이미 채워진 셀은 덮어쓰지 않는 멱등성 체크. Returns: written_count."""
    wb, xlsx_path = download_workbook(contents_sheet_id)
    ws = wb["TRANSLATE_DATA"]
    headers, rows = ws_read_dicts(ws)
    by_file_id = {r.get("file_id"): r for r in rows}
    written = 0
    for file_id, column, value in updates:
        r = by_file_id.get(file_id)
        if r is None:
            r = {h: "" for h in headers}
            r["file_id"] = file_id
            rows.append(r)
            by_file_id[file_id] = r
        if (r.get(column) or "").strip():
            continue  # 이미 채워져 있음 - 덮어쓰지 않음(멱등성)
        r[column] = value
        written += 1
    if written:
        ws_write_dicts(ws, headers, rows)
        upload_workbook(wb, xlsx_path, contents_sheet_id)
    else:
        os.remove(xlsx_path)
    return written


def main():
    client = OpenAI(api_key=os.environ["OPENAI_API_KEY"])
    now = time.time()

    batches = [
        b for b in client.batches.list(limit=100)
        if now - b.created_at <= MAX_BATCH_AGE_SEC
    ]
    completed = [
        b for b in batches
        if b.status == "completed" and (b.metadata or {}).get("contents_sheet_id") and b.output_file_id
    ]
    print(f"최근 {MAX_BATCH_AGE_SEC // 3600}h 배치 {len(batches)}개 중 처리대상 {len(completed)}개")

    by_sheet = {}
    for b in completed:
        by_sheet.setdefault(b.metadata["contents_sheet_id"], []).append(b)

    for contents_sheet_id, sheet_batches in by_sheet.items():
        updates = []
        for b in sheet_batches:
            field = b.metadata.get("field")
            target_lang = b.metadata.get("target_lang")
            if not field or not target_lang:
                continue
            column = f"{target_lang}_{field}"
            content = client.files.content(b.output_file_id).text
            for line in content.splitlines():
                if not line.strip():
                    continue
                row = json.loads(line)
                custom_id = row.get("custom_id")
                resp = row.get("response") or {}
                if resp.get("status_code") != 200:
                    print(f"  경고: batch={b.id} custom_id={custom_id} 응답 실패, 건너뜀")
                    continue
                try:
                    text = resp["body"]["choices"][0]["message"]["content"].strip()
                except (KeyError, IndexError, TypeError):
                    continue
                if custom_id and text:
                    updates.append((custom_id, column, text))
        if updates:
            written = upsert_translate_data(contents_sheet_id, updates)
            print(f"{contents_sheet_id}: {len(updates)}건 확인, {written}건 신규 반영")


if __name__ == "__main__":
    main()
