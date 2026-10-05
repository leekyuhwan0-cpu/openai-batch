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

Gemini QA 검증 레이어(2026-09-26 확정, project_번역검증_gemini_파이프라인_확정 참고):
GPT 배치 번역 결과를 TRANSLATE_DATA에 쓰기 전에 gemini-3.8-flash로 (category,
target_lang, field) 그룹 단위 배치검증 -> FAIL만 gpt-6-luna reasoning_effort=medium로
재번역 -> 재검증 -> 그래도 FAIL이면 해당 칸에 "[번역실패]" 표식(2026-10-05 개정, Vertex 사용).
poll_batches.py는 별도 레포라 bonus_model_interfaces.py/bonus_sheet_io.py를 그대로
import할 수 없음(파일 자체가 없고, bonus_sheet_io.py는 top-level `import msvcrt`라
Linux 러너에서 죽음) - 그래서 TRANSLATE_PROMPTS/SOURCE_ACCOUNTS 조회도 이 파일 안의
download_workbook/ws_read_dicts를 그대로 재사용해 MAIN_SHEET_ID를 직접 읽는다(OAuth
자격증명은 bonus_sheet_io.py의 하드코딩 값과 동일해 기존 GOOGLE_* 시크릿으로 충분).
"""
import json
import os
import re
import time
import uuid

import requests
from openpyxl import load_workbook
from openai import OpenAI

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
REQUEST_TIMEOUT = 30
HERE = os.path.dirname(os.path.abspath(__file__))
MAX_BATCH_AGE_SEC = 3 * 24 * 3600
MAIN_SHEET_ID = "1IK7MbiTFJiV_ofeco9FwYs7UFY_f4QHHIDZJb7T_4-o"

GEMINI_MODEL = "gemini-3.8-flash"
VERTEX_API_KEY = os.environ.get("VERTEX_API_KEY", "")
VERTEX_URL = f"https://aiplatform.googleapis.com/v1/publishers/google/models/{GEMINI_MODEL}:generateContent"
VERIFY_WORKERS = 4
XHIGH_MODEL = "gpt-6-luna"  # project_번역모델_확정 참고 - 후킹/캡션 둘다 이 모델로 확정운영중

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


SHEETS_API_BASE = "https://sheets.googleapis.com/v4/spreadsheets"


def _col_letter(idx0):
    n = idx0 + 1
    out = ""
    while n:
        n, rem = divmod(n - 1, 26)
        out = chr(65 + rem) + out
    return out


def _sheet_values(contents_sheet_id, tab):
    """탭 전체를 Sheets API로 지금 시점 그대로 읽는다. -> (headers, [(시트행번호, dict), ...])"""
    from urllib.parse import quote
    r = _request_with_retry("GET", f"{SHEETS_API_BASE}/{contents_sheet_id}/values/{quote(tab)}")
    values = r.json().get("values") or []
    if not values:
        return [], []
    headers = values[0]
    rows = []
    for i, row in enumerate(values[1:], start=2):
        if not any(str(v).strip() for v in row):
            continue
        rows.append((i, {headers[j]: row[j] for j in range(len(headers)) if j < len(row)}))
    return headers, rows


def upsert_translate_data(wb, xlsx_path, contents_sheet_id, updates):
    """updates: [(file_id, column, value), ...]. Returns: written_count.

    2026-10-04: 워크북 통째 업로드 방식 폐기 -> 셀 단위 Sheets API. 예전 방식은 (1) 배치
    완료~Gemini 검증이 끝날 때까지 들고 있던 오래된 사본으로 TRANSLATE_DATA 탭을 통째로
    덮어써서 그 사이 앱에서 삭제한 콘텐츠(CONTENTS/ASSETS 행 포함)를 되살릴 수 있었고,
    (2) 삭제된 콘텐츠의 늦게 도착한 결과가 TRANSLATE_DATA 맨 아래에 고아 행으로 새로
    생겼다(603번). 이제 쓰기 직전에 TRANSLATE_DATA/ASSETS를 새로 읽어 (a) ASSETS에
    실존하는 file_id만 쓰고 (b) 기존 행은 비어 있는 셀만 채우고 (c) 새 행은 append한다.
    wb/xlsx_path는 호출부 시그니처 호환용 - 더 이상 업로드하지 않고 임시파일만 지운다."""
    try:
        os.remove(xlsx_path)
    except OSError:
        pass
    headers, rows = _sheet_values(contents_sheet_id, "TRANSLATE_DATA")
    _, asset_rows = _sheet_values(contents_sheet_id, "ASSETS")
    alive = {str(r.get("local_file_id")) for _, r in asset_rows}
    col_idx = {h: i for i, h in enumerate(headers)}
    by_file_id = {str(r.get("file_id")): (rn, r) for rn, r in rows}

    cell_updates = []
    new_rows = {}
    written = 0
    dropped = 0
    for file_id, column, value in updates:
        if str(file_id) not in alive:
            dropped += 1
            continue
        if column not in col_idx:
            raise ValueError(f"TRANSLATE_DATA 헤더에 '{column}' 없음")
        hit = by_file_id.get(str(file_id))
        if hit is not None:
            rn, r = hit
            if (r.get(column) or "").strip():
                continue  # 이미 채워져 있음 - 덮어쓰지 않음(멱등성)
            cell_updates.append((f"TRANSLATE_DATA!{_col_letter(col_idx[column])}{rn}", value))
            r[column] = value
            written += 1
        else:
            nr = new_rows.setdefault(str(file_id), {h: "" for h in headers})
            nr["file_id"] = str(file_id)
            if not (nr.get(column) or "").strip():
                nr[column] = value
                written += 1
    if dropped:
        print(f"  경고: ASSETS에 없는(삭제된) file_id {dropped}건은 쓰지 않음")
    if cell_updates:
        _request_with_retry(
            "POST", f"{SHEETS_API_BASE}/{contents_sheet_id}/values:batchUpdate",
            json={"valueInputOption": "RAW",
                  "data": [{"range": a1, "values": [[v]]} for a1, v in cell_updates]},
        )
    if new_rows:
        _request_with_retry(
            "POST", f"{SHEETS_API_BASE}/{contents_sheet_id}/values/TRANSLATE_DATA:append",
            params={"valueInputOption": "RAW", "insertDataOption": "INSERT_ROWS"},
            json={"values": [[r.get(h, "") for h in headers] for r in new_rows.values()]},
        )
    return written


_main_lookup_cache = None


def _load_main_lookups():
    """MAIN_SHEET_ID(TRANSLATE_PROMPTS/SOURCE_ACCOUNTS가 있는 마스터 시트)를 실행당
    1회만 다운로드해 캐싱. category는 CONTENTS/ASSETS가 아니라 SOURCE_ACCOUNTS 행에
    있고(bonus_job_processor.py의 ch_data["sa"].get("category") 확인), SOURCE_ACCOUNTS는
    channel_id 단위지 contents_sheet_id 단위가 아니므로 contents_sheet_id ->
    category 매핑을 SOURCE_ACCOUNTS.contents_sheet_id 컬럼으로 만들어둔다."""
    global _main_lookup_cache
    if _main_lookup_cache is None:
        wb, xlsx_path = download_workbook(MAIN_SHEET_ID)
        try:
            _, tp_rows = ws_read_dicts(wb["TRANSLATE_PROMPTS"])
            _, sa_rows = ws_read_dicts(wb["SOURCE_ACCOUNTS"])
        finally:
            os.remove(xlsx_path)
        category_by_sheet_id = {r.get("contents_sheet_id"): r.get("category") for r in sa_rows}
        _main_lookup_cache = (tp_rows, category_by_sheet_id)
    return _main_lookup_cache


def _get_translate_prompt_row(tp_rows, category, lang):
    for row in tp_rows:
        if row.get("translate_category") == category and row.get("translate_lang") == lang:
            return row
    return None


def build_source_lookup(wb):
    """같은 contents_sheet_id 워크북 안의 CONTENTS/ASSETS 탭에서 custom_id
    (=ASSETS.local_file_id, bonus_ui.py의 submit_translation_batch 호출부 확인)로
    원문을 조회할 수 있는 딕셔너리를 만든다."""
    _, c_rows = ws_read_dicts(wb["CONTENTS"])
    _, a_rows = ws_read_dicts(wb["ASSETS"])
    content_by_id = {c.get("content_id"): c for c in c_rows}
    assets_by_local_file_id = {a.get("local_file_id"): a for a in a_rows if a.get("local_file_id")}
    return content_by_id, assets_by_local_file_id


def get_source_text(field, custom_id, content_by_id, assets_by_local_file_id):
    """field="hook"이면 그 asset 자신의 source_hook, "caption"이면 그 asset의
    content_id로 CONTENTS를 찾아 source_caption(캡션 배치의 custom_id는 main_asset의
    local_file_id라 이렇게 한 단계 거쳐야 함)."""
    asset = assets_by_local_file_id.get(custom_id)
    if asset is None:
        return None
    if field == "hook":
        return (asset.get("source_hook") or "").strip()
    content = content_by_id.get(asset.get("content_id"))
    return (content.get("source_caption") or "").strip() if content else None


def build_chunk_system_prompt(tp_row, field):
    """앱(bonus_model_interfaces._build_chunk_system_prompt)과 동일 규칙: 후킹=gpt_prompt, 캡션=
    gpt_caption_prompt, {EXCHANGE_RATE} 치환 + {TEXT} 줄 제거. 없으면 ""."""
    col = "gpt_prompt" if field == "hook" else "gpt_caption_prompt"
    prompt = ((tp_row or {}).get(col) or "").strip()
    if not prompt:
        return ""
    prompt = prompt.replace("{EXCHANGE_RATE}", ((tp_row or {}).get("exchange_rate") or "").strip())
    return "\n".join(ln for ln in prompt.split("\n") if "{TEXT}" not in ln).strip()


_CHUNK_OUTPUT_INSTRUCTION = (
    'Return only a JSON object of the form {"results": [{"id": "<item id>", "text": "<translation>"}]} '
    "with exactly one result per input item, using the same id. Keep line breaks inside text as \n."
)


def _chunk_response_format(ids):
    return {"type": "json_schema", "json_schema": {
        "name": "translations", "strict": True,
        "schema": {
            "type": "object", "additionalProperties": False, "required": ["results"],
            "properties": {"results": {"type": "array", "items": {
                "type": "object", "additionalProperties": False, "required": ["id", "text"],
                "properties": {"id": {"type": "string", "enum": list(ids)}, "text": {"type": "string"}},
            }}},
        },
    }}


def _build_single_user_message(field, text, context_caption):
    item = {"id": "single", "text": text}
    payload = {"items": [item]}
    if field == "hook" and context_caption:
        item = {"id": "single", "cid": "c", "text": text}
        payload = {"contexts": {"c": context_caption}, "items": [item]}
    return json.dumps(payload, ensure_ascii=False) + "\n\n" + _CHUNK_OUTPUT_INSTRUCTION


def parse_chunk_response(content, expected_ids):
    """앱의 parse_chunk_response와 동일 규칙: JSON 실패 -> {}, 요청에 없던 id 무시, 중복 id는 첫 값,
    빈 text는 누락 취급."""
    try:
        results = json.loads(content)["results"]
    except Exception:
        return {}
    expected = set(expected_ids)
    out = {}
    for r in results if isinstance(results, list) else []:
        try:
            rid, text = r["id"], (r["text"] or "").strip()
        except Exception:
            continue
        if rid in expected and rid not in out and text:
            out[rid] = text
    return out


def translate_single(client, system_prompt, field, text, context_caption, reasoning_effort="low"):
    """아이템 1개짜리 청크로 개별 재호출(청크와 같은 프롬프트/스키마/파서). 실패 시 None."""
    try:
        resp = client.chat.completions.create(
            model=XHIGH_MODEL,
            messages=[{"role": "system", "content": system_prompt},
                      {"role": "user", "content": _build_single_user_message(field, text, context_caption)}],
            response_format=_chunk_response_format(["single"]),
            reasoning_effort=reasoning_effort,
        )
        return parse_chunk_response(resp.choices[0].message.content or "", ["single"]).get("single")
    except Exception as e:
        print(f"    개별 재호출 실패: {e}")
        return None


def get_context_caption(custom_id, content_by_id, assets_by_local_file_id):
    asset = assets_by_local_file_id.get(custom_id)
    content = content_by_id.get(asset.get("content_id")) if asset else None
    return (content.get("source_caption") or "").strip() if content else ""


def chunk_expected_ids(client, batch):
    """청크 배치의 input 파일에서 청크별 요청 id(response_format enum)를 복원. -> {chunk custom_id: [ids]}"""
    out = {}
    text = client.files.content(batch.input_file_id).text
    for line in text.splitlines():
        if not line.strip():
            continue
        req = json.loads(line)
        try:
            enum = req["body"]["response_format"]["json_schema"]["schema"]["properties"]["results"]["items"]["properties"]["id"]["enum"]
        except (KeyError, TypeError):
            enum = []
        out[req.get("custom_id")] = list(enum)
    return out


def _gemini_call(payload):
    """Vertex express 호출. 429/500/503/연결오류는 백오프 재시도(최대 4회). 전부 실패하면 None
    (호출부가 GPT 원본을 그대로 채택하고 로그만 남김)."""
    if not VERTEX_API_KEY:
        print("  경고: VERTEX_API_KEY 없음 - 검증 없이 원본 채택")
        return None
    for attempt in range(1, 5):
        try:
            resp = requests.post(VERTEX_URL, params={"key": VERTEX_API_KEY}, json=payload, timeout=180)
        except (requests.exceptions.ConnectionError, requests.exceptions.Timeout):
            time.sleep(4 * attempt)
            continue
        if resp.status_code == 200:
            return resp.json()
        if resp.status_code in (429, 500, 503):
            time.sleep(4 * attempt)
            continue
        print(f"  Vertex HTTP {resp.status_code}: {resp.text[:200]}")
        return None
    print("  경고: Vertex 호출 재시도 소진 - 해당 청크는 검증 없이 원본 채택")
    return None


# 2026-10-05: 검증 재설계 - 스타일/품질은 보지 않고 (1)금액 환산 (2)숫자 오류 (3)원본통화 잔존
# (4)미번역만 본다. 지침은 코드가 아니라 TRANSLATE_PROMPTS.verify_prompt(시트)에서 읽는다.
VERIFY_CHUNK = 20          # GPT 번역 청크(20개)와 동일 단위
FAIL_MARK = "[번역실패]"    # 재번역까지 실패한 칸 앞에 붙이는 표식 "[번역실패] 실패한번역문"(앱이 빨간색으로 표시)

LANG_NAMES = {
    "kr": "한국어", "ja": "일본어", "de": "독일어", "tr": "터키어", "en": "영어", "es": "스페인어",
    "fr": "프랑스어", "it": "이탈리아어", "nl": "네덜란드어", "se": "스웨덴어", "dk": "덴마크어", "no": "노르웨이어",
}

_VERIFY_SCHEMA = {
    "type": "ARRAY",
    "items": {
        "type": "OBJECT",
        "properties": {"id": {"type": "INTEGER"}, "reason": {"type": "STRING"}},
        "required": ["id", "reason"],
    },
}

_TARGET_SCRIPT = {
    "kr": re.compile(r"[가-힣]"),
    "ja": re.compile(r"[぀-ヿ一-鿿]"),
}

def looks_untranslated(source, translated, target_lang):
    """코드 미번역 판별(Gemini와 2중 필터) - kr/ja 전용. 타겟 문자가 없거나 라틴 문자가 타겟
    문자보다 많으면 미번역. 다른 언어는 Gemini에만 맡기므로 항상 False."""
    pat = _TARGET_SCRIPT.get(target_lang)
    if pat is None:
        return False
    t = (translated or "").strip()
    if not t:
        return True
    n_target = len(pat.findall(t))
    n_latin = len(re.findall(r"[A-Za-zÀ-ÿĞğİıŞş]", t))
    return n_target == 0 or n_latin > n_target


def build_verify_prompt(vp_row, target_lang):
    """verify_prompt(시트) + {TARGET_LANG}/{EXCHANGE_RATE} 치환. 없으면 ""."""
    prompt = ((vp_row or {}).get("verify_prompt") or "").strip()
    if not prompt:
        return ""
    return (prompt.replace("{TARGET_LANG}", LANG_NAMES.get(target_lang, target_lang))
                  .replace("{EXCHANGE_RATE}", ((vp_row or {}).get("exchange_rate") or "").strip()))


def _gemini_verify(verify_prompt, items):
    """items: [{"id": int, "content_id", "field", "source", "translated"}, ...] (<=VERIFY_CHUNK).
    Returns: {id: reason} (FAIL 항목만, 전부 PASS면 {}) 또는 호출/파싱 실패 시 None."""
    lines = [f"[id={it['id']}] (콘텐츠={it.get('content_id')}, 종류={'후킹' if it['field'] == 'hook' else '캡션'})\n"
             f"원문: {it['source']}\n번역: {it['translated']}" for it in items]
    prompt = f"{verify_prompt}\n\n===\n\n[검수 항목 {len(items)}개]\n\n" + "\n\n".join(lines)
    data = _gemini_call({
        "contents": [{"role": "user", "parts": [{"text": prompt}]}],
        "generationConfig": {"responseMimeType": "application/json", "responseSchema": _VERIFY_SCHEMA,
                             "thinkingConfig": {"thinkingBudget": 0}},
    })
    if data is None:
        return None
    try:
        text = "".join(p.get("text", "") for p in data["candidates"][0]["content"]["parts"])
        return {p["id"]: p.get("reason", "") for p in json.loads(text)}
    except (KeyError, IndexError, TypeError, json.JSONDecodeError) as e:
        print(f"  경고: Gemini 검증 응답 파싱 실패({e})")
        return None


def _verify_in_chunks(verify_prompt, items):
    """VERIFY_CHUNK개씩 끊어 VERIFY_WORKERS개 병렬 검증. Returns: (fails {id: reason}, unverified_ids set)."""
    from concurrent.futures import ThreadPoolExecutor
    chunks = [items[i:i + VERIFY_CHUNK] for i in range(0, len(items), VERIFY_CHUNK)]
    fails, unverified = {}, set()
    if not chunks:
        return fails, unverified
    with ThreadPoolExecutor(VERIFY_WORKERS) as ex:
        results = list(ex.map(lambda c: _gemini_verify(verify_prompt, c), chunks))
    for chunk, res in zip(chunks, results):
        if res is None:
            unverified.update(it["id"] for it in chunk)
            continue
        ids = {it["id"] for it in chunk}
        fails.update({k: v for k, v in res.items() if k in ids})
    return fails, unverified


def _gpt_retry_medium(client, system_prompt, field, source_text, context_caption=""):
    """FAIL난 항목을 같은 시스템 프롬프트/스키마(아이템 1개 청크)로 medium 재번역. 실패 시 예외."""
    resp = client.chat.completions.create(
        model=XHIGH_MODEL,
        messages=[{"role": "system", "content": system_prompt},
                  {"role": "user", "content": _build_single_user_message(field, source_text, context_caption)}],
        response_format=_chunk_response_format(["single"]),
        reasoning_effort="medium",
    )
    text = parse_chunk_response(resp.choices[0].message.content or "", ["single"]).get("single")
    if not text:
        raise RuntimeError("medium 응답 파싱 실패/빈 결과")
    return text


def verify_lang_group(client, tp_row, target_lang, items, events=None):
    """한 타겟 언어의 번역 항목 전체(후킹+캡션)를 검증/교정한다.
    흐름: Gemini 검증(+kr/ja 코드 미번역 판별) -> FAIL은 GPT medium 재번역 -> 재검증 ->
    통과면 적용, 또 FAIL이면 FAIL_MARK. Gemini 호출 자체가 실패한 항목은 GPT 원본 유지(로그만).
    items: [{"custom_id","field","content_id","source","translated","context"}, ...]
    Returns: [(custom_id, field, value), ...]. events: 테스트/로그용 (custom_id, field, 종류, 상세) 리스트."""
    def ev(it, kind, detail=""):
        print(f"    [{kind}] {it['custom_id']} {it['field']} {detail[:120]}")
        if events is not None:
            events.append((it["custom_id"], it["field"], kind, detail))

    sys_prompts = {f: build_chunk_system_prompt(tp_row, f) for f in ("hook", "caption")}
    verify_prompt = build_verify_prompt(tp_row, target_lang)
    if not verify_prompt:
        print(f"  경고: lang={target_lang!r} verify_prompt 없음 - Gemini 검증 스킵(kr/ja 코드검사만 수행)")

    final = {(it["custom_id"], it["field"]): it["translated"] for it in items}

    # 원문이 비었거나 "-"인 항목은 검증 대상 아님
    targets = [it for it in items if (it["source"] or "").strip() not in ("", "-")]
    targets.sort(key=lambda x: (str(x.get("content_id")), x["custom_id"], x["field"]))
    for n, it in enumerate(targets):
        it["id"] = n

    # 1) Gemini 검증 + 코드 미번역 판별
    if verify_prompt and targets:
        fails, unverified = _verify_in_chunks(verify_prompt, targets)
    else:
        fails, unverified = {}, set()
    for it in targets:
        if it["id"] in unverified:
            ev(it, "검증불가", "Vertex 실패 - GPT 원본 유지")
        if it["id"] in fails:
            ev(it, "FAIL", fails[it["id"]])
        if looks_untranslated(it["source"], it["translated"], target_lang):
            if it["id"] not in fails:
                ev(it, "미번역(코드)", it["translated"][:40])
            fails[it["id"]] = fails.get(it["id"]) or "미번역(코드 판별)"

    # 2) FAIL -> medium 재번역
    retried = []
    for it in (t for t in targets if t["id"] in fails):
        try:
            t = _gpt_retry_medium(client, sys_prompts[it["field"]], it["field"], it["source"], it["context"])
        except Exception as e:
            print(f"    medium 재번역 실패(custom_id={it['custom_id']}): {e}")
            ev(it, "번역실패", "재번역 호출 실패")
            final[(it["custom_id"], it["field"])] = f"{FAIL_MARK} {it['translated']}"
            continue
        retried.append(dict(it, translated=t))

    # 3) 재검증 -> 통과면 적용, FAIL이면 표식
    for n, it in enumerate(retried):
        it["id"] = n
    if verify_prompt and retried:
        refails, reunv = _verify_in_chunks(verify_prompt, retried)
    else:
        refails, reunv = {}, set()
    for it in retried:
        key = (it["custom_id"], it["field"])
        code_bad = looks_untranslated(it["source"], it["translated"], target_lang)
        if it["id"] in refails or code_bad:
            ev(it, "번역실패", refails.get(it["id"]) or "재번역도 미번역(코드)")
            final[key] = f"{FAIL_MARK} {it['translated']}"
        else:
            if it["id"] in reunv:
                ev(it, "재검증불가", "재번역본 미검증 적용")
            ev(it, "재번역 통과", it["translated"][:60])
            final[key] = it["translated"]
    return [(cid, field, v) for (cid, field), v in final.items()]


def verify_and_correct(client, contents_sheet_id, wb, raw_updates):
    """raw_updates: [(custom_id, field, target_lang, text), ...] (GPT 배치 원본 결과).
    언어별로 후킹+캡션을 함께 verify_lang_group으로 검증/교정. Returns: [(custom_id, column,
    value), ...] (upsert_translate_data에 바로 넘길 최종 형태)."""
    content_by_id, assets_by_local_file_id = build_source_lookup(wb)
    tp_rows, category_by_sheet_id = _load_main_lookups()
    category = category_by_sheet_id.get(contents_sheet_id)

    by_lang = {}
    final_updates = []
    for custom_id, field, target_lang, text in raw_updates:
        source_text = get_source_text(field, custom_id, content_by_id, assets_by_local_file_id)
        if not source_text:
            print(f"  경고: custom_id={custom_id} 원문 조회 실패 - 검증 없이 원본 채택")
            final_updates.append((custom_id, f"{target_lang}_{field}", text))
            continue
        asset = assets_by_local_file_id.get(custom_id) or {}
        by_lang.setdefault(target_lang, []).append({
            "custom_id": custom_id, "field": field, "content_id": asset.get("content_id"),
            "source": source_text, "translated": text,
            "context": get_context_caption(custom_id, content_by_id, assets_by_local_file_id)})

    for target_lang, items in by_lang.items():
        tp_row = _get_translate_prompt_row(tp_rows, category, target_lang)
        results = verify_lang_group(client, tp_row, target_lang, items)
        n_fail = sum(1 for _, _, v in results if v.startswith(FAIL_MARK))
        print(f"  category={category} lang={target_lang}: {len(items)}건 검증 완료 (번역실패 {n_fail}건)")
        final_updates.extend((cid, f"{target_lang}_{field}", v) for cid, field, v in results)
    return final_updates


def main():
    client = OpenAI(api_key=os.environ["OPENAI_API_KEY"])
    dispatch_batch_id = (os.environ.get("DISPATCH_BATCH_ID") or "").strip()

    if dispatch_batch_id:
        # 즉시-트리거 경로: submit_translation_batch가 완료를 직접 감지해 이 배치
        # ID 하나만 콕 집어 깨웠다 - 72h 전체 스캔/재검증 없이 이 배치 세트만 처리.
        b = client.batches.retrieve(dispatch_batch_id)
        if b.status != "completed" or not b.output_file_id or not (b.metadata or {}).get("contents_sheet_id"):
            print(f"batch_id={dispatch_batch_id}: 아직 처리 불가(status={b.status}) - 종료")
            return
        completed = [b]
        print(f"[즉시 처리] batch_id={dispatch_batch_id} 단건")
    else:
        # 백업 스캔 경로: 즉시-트리거를 놓쳤을 때만 여기로 옴(GUI 종료 등, 드묾).
        now = time.time()
        batches = [
            b for b in client.batches.list(limit=100)
            if now - b.created_at <= MAX_BATCH_AGE_SEC
        ]
        completed = [
            b for b in batches
            if b.status == "completed" and (b.metadata or {}).get("contents_sheet_id") and b.output_file_id
        ]
        print(f"[백업 스캔] 최근 {MAX_BATCH_AGE_SEC // 3600}h 배치 {len(batches)}개 중 처리대상 {len(completed)}개")

    by_sheet = {}
    for b in completed:
        by_sheet.setdefault(b.metadata["contents_sheet_id"], []).append(b)

    for contents_sheet_id, sheet_batches in by_sheet.items():
        raw_updates = []  # (custom_id, field, target_lang, text) - 검증 전 GPT 원본
        missing = []  # 청크 응답에서 누락/파싱실패한 (id, field, target_lang, category) - 개별 재호출 대상
        for b in sheet_batches:
            field = b.metadata.get("field")
            target_lang = b.metadata.get("target_lang")
            if not field or not target_lang:
                continue
            is_chunk = b.metadata.get("mode") == "chunk"
            expected_by_chunk = chunk_expected_ids(client, b) if is_chunk else {}
            content = client.files.content(b.output_file_id).text
            seen_chunks = set()
            for line in content.splitlines():
                if not line.strip():
                    continue
                row = json.loads(line)
                custom_id = row.get("custom_id")
                resp = row.get("response") or {}
                if is_chunk:
                    seen_chunks.add(custom_id)
                    expected = expected_by_chunk.get(custom_id, [])
                    msg = ""
                    if resp.get("status_code") == 200:
                        try:
                            msg = resp["body"]["choices"][0]["message"]["content"] or ""
                        except (KeyError, IndexError, TypeError):
                            msg = ""
                    parsed = parse_chunk_response(msg, expected)
                    for rid in expected:
                        if rid in parsed:
                            raw_updates.append((rid, field, target_lang, parsed[rid]))
                        else:
                            missing.append((rid, field, target_lang, b.metadata.get("category") or ""))
                    continue
                if resp.get("status_code") != 200:
                    print(f"  경고: batch={b.id} custom_id={custom_id} 응답 실패, 건너뜀")
                    continue
                try:
                    text = resp["body"]["choices"][0]["message"]["content"].strip()
                except (KeyError, IndexError, TypeError):
                    continue
                if custom_id and text:
                    raw_updates.append((custom_id, field, target_lang, text))
            if is_chunk:
                # output에 아예 없는 청크(만료/에러파일행)도 누락 처리
                for cid_, ids in expected_by_chunk.items():
                    if cid_ not in seen_chunks:
                        missing.extend((rid, field, target_lang, b.metadata.get("category") or "") for rid in ids)

        if not raw_updates and not missing:
            continue


        wb, xlsx_path = download_workbook(contents_sheet_id)

        # 안전망: TRANSLATE_DATA에 이미 값이 채워진 (file_id, column)은 검증 전에
        # 제외 - 백업 스캔 경로에서 이미 처리된 배치가 섞여 들어와도 Gemini를
        # 다시 태우지 않는다(즉시-트리거 경로는 항상 방금 완료된 배치 1개뿐이라
        # 보통 걸릴 게 없지만, 동일하게 적용해도 무해함).
        _, td_rows = ws_read_dicts(wb["TRANSLATE_DATA"])
        td_by_id = {r.get("file_id"): r for r in td_rows}

        if missing:
            content_by_id, assets_by_local_file_id = build_source_lookup(wb)
            tp_rows, category_by_sheet_id = _load_main_lookups()
            print(f"  {contents_sheet_id}: 청크 응답 누락/실패 {len(missing)}건 - 개별 재호출")
            for rid, field, target_lang, cat in missing:
                if ((td_by_id.get(rid) or {}).get(f"{target_lang}_{field}") or "").strip():
                    continue  # 이미 채워져 있음
                tp_row = _get_translate_prompt_row(tp_rows, cat or category_by_sheet_id.get(contents_sheet_id), target_lang)
                system_prompt = build_chunk_system_prompt(tp_row, field)
                src = get_source_text(field, rid, content_by_id, assets_by_local_file_id)
                if not system_prompt or not src:
                    print(f"    id={rid}: 프롬프트/원문 없음 - 쓰지 않음(최종저장에서 차단됨)")
                    continue
                ctx = get_context_caption(rid, content_by_id, assets_by_local_file_id) if field == "hook" else ""
                text = translate_single(client, system_prompt, field, src, ctx)
                if text:
                    raw_updates.append((rid, field, target_lang, text))
                else:
                    print(f"    id={rid}: 개별 재호출도 실패 - 쓰지 않음(최종저장에서 차단됨)")
        if not raw_updates:
            os.remove(xlsx_path)
            continue
        new_updates = [
            (cid, field, target_lang, text)
            for (cid, field, target_lang, text) in raw_updates
            if not ((td_by_id.get(cid) or {}).get(f"{target_lang}_{field}") or "").strip()
        ]
        skipped = len(raw_updates) - len(new_updates)
        if skipped:
            print(f"  {contents_sheet_id}: {len(raw_updates)}건 중 {skipped}건 이미 처리됨 - {len(new_updates)}건만 검증")
        if not new_updates:
            os.remove(xlsx_path)
            continue

        final_updates = verify_and_correct(client, contents_sheet_id, wb, new_updates)
        if final_updates:
            written = upsert_translate_data(wb, xlsx_path, contents_sheet_id, final_updates)
            print(f"{contents_sheet_id}: {len(final_updates)}건 확인, {written}건 신규 반영")
        else:
            os.remove(xlsx_path)


if __name__ == "__main__":
    main()
