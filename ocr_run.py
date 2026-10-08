"""OCR 사전제출 서버 처리(2026-10-09). bonus_ui의 PreRunWorker(kind=ocr)를 대체.

GUI가 workflow_dispatch로 (시트ID, 번호범위)만 넘기고 끈다. 여기서:
  ASSETS에서 source_hook이 빈 이미지(영상 V 제외)를 Drive에서 받아 Gemini(Vertex express)로
  OCR -> 청크(CHUNK개)마다 즉시 ASSETS.source_hook에 기록(슬라이드 1이고 CONTENTS.keyword가
  비었으면 keyword도). 중간에 끊겨도 쓴 분량은 보존되고, 실패 건은 빈칸으로 남아 재실행 시 재포함.
프롬프트/모델/스레드 수는 전부 메인시트 v2_MODEL_REGISTRY(hook_ocr, keyword 행) params에서 읽는다
(hook_keyword.py와 동일 로직 - 코드 하드코딩 금지 원칙)."""
import base64
import concurrent.futures
import json
import os
import re
import sys
import time

import requests

from poll_batches import (MAIN_SHEET_ID, SHEETS_API_BASE, VERTEX_API_KEY, _col_letter,
                          _request_with_retry, _sheet_values)

CHUNK = 40
MAX_RUN_SEC = 5 * 3600 + 1800  # 6시간 제한 전에 스스로 멈춤(남은 건 재실행으로 이어감)
_VERTEX_URL = "https://aiplatform.googleapis.com/v1/publishers/google/models/{model}:generateContent"
_HANGUL_WORD = re.compile(r"^[가-힣]+$")
DEFAULT_HOOK_PROMPT = (
    "이 이미지 하단에 오버레이된 캡션 텍스트를 보이는 그대로(원문 언어 그대로) 정확히 옮겨 적어. "
    "로고/워터마크/화살표 안내문구(예: 채널명, KAYDIRIN 등)는 제외하고 캡션 본문 문장만 적어. "
    "텍스트가 전혀 안 보이면 '-' 한 글자만 출력해. 번역이나 설명 없이 원문 텍스트만 출력해."
)
DEFAULT_KEYWORD_RULE = (
    "You extract exactly ONE Korean keyword (a single word, no particles) that best represents "
    "the core subject of the given news script, for use as an internal filename tag. "
    "It must be written entirely in Korean (한국어) and must be exactly one word with no spaces."
)


class TransientError(Exception):
    pass


def _to_int(v):
    try:
        return int(str(v).strip())
    except Exception:
        return 0


def _mime(b):
    if b[:4] == b"\x89PNG":
        return "image/png"
    if b[:4] == b"RIFF":
        return "image/webp"
    return "image/jpeg"


def _call_gemini(model, thinking, prompt, image_bytes, timeout, attempts=3):
    payload = {"contents": [{"role": "user", "parts": [
        {"inline_data": {"mime_type": _mime(image_bytes), "data": base64.b64encode(image_bytes).decode()}},
        {"text": prompt}]}]}
    if thinking:
        payload["generationConfig"] = {"thinkingConfig": thinking}
    url = _VERTEX_URL.format(model=model)
    last = "unknown"
    for a in range(1, attempts + 1):
        try:
            r = requests.post(url, params={"key": VERTEX_API_KEY}, json=payload, timeout=timeout)
        except requests.RequestException as e:
            last = type(e).__name__  # URL(키)이 섞일 수 있어 클래스명만
            time.sleep(2 * a)
            continue
        if r.status_code == 200:
            try:
                parts = r.json()["candidates"][0]["content"]["parts"]
                return "".join(p.get("text", "") for p in parts).strip()
            except Exception:
                return ""
        if r.status_code in (429, 500, 503):
            last = f"HTTP {r.status_code}"
            time.sleep(4 * a)
            continue
        raise RuntimeError(f"Gemini HTTP {r.status_code}: {r.text[:200].replace(VERTEX_API_KEY, '***')}")
    raise TransientError(last)


def _registry_params(rows, task):
    for r in rows:
        if r.get("task_category") == task:
            try:
                return json.loads(r.get("params") or "{}")
            except Exception:
                return {}
    return {}


def load_cfg():
    _, rows = _sheet_values(MAIN_SHEET_ID, "v2_MODEL_REGISTRY")
    rows = [r for _, r in rows]
    h, k = _registry_params(rows, "hook_ocr"), _registry_params(rows, "keyword")
    return {
        "model": h.get("model", "gemini-3.1-flash-lite"),
        "thinking": h.get("thinking", {"thinkingBudget": 0}),
        "delimiter": h.get("delimiter", "###KEYWORD###"),
        "placeholder": h.get("empty_placeholder", "-"),
        "hook_prompt": h.get("prompt") or DEFAULT_HOOK_PROMPT,
        "workers": int(h.get("workers", 8)),
        "fallback_workers": int(h.get("fallback_workers", 4)),
        "timeout": int(h.get("timeout_sec", 120)),
        "parse_retries": int(h.get("parse_retries", 2)),
        "kw_rule": k.get("system_prompt") or DEFAULT_KEYWORD_RULE,
        "kw_retry_model": k.get("retry_model", "gemini-3.8-flash"),
        "kw_retry_thinking": k.get("retry_thinking", {"thinkingBudget": 0}),
        "kw_retries": int(k.get("max_retries", 2)),
    }


def _combined_prompt(cfg):
    return (cfg["hook_prompt"] + "\n\n그 다음, 아래 규칙으로 키워드를 뽑아.\n" + cfg["kw_rule"] + "\n"
            "캡션이 '-'(텍스트 없음)이면 키워드도 '-'로 해.\n\n"
            f"출력 형식(정확히 이 형식만): 첫 줄들에 캡션 원문, 그 다음 줄에 {cfg['delimiter']} 라는 구분자, "
            "그 다음 줄에 키워드 한 단어. 다른 설명 금지.")


def _keyword_only_prompt(cfg):
    return ("이 이미지 하단 캡션의 핵심 주제를 대표하는 키워드를 뽑아.\n" + cfg["kw_rule"] + "\n"
            "캡션이 없으면 이미지 전체의 핵심 피사체로 뽑아. 키워드 한 단어만 출력하고 다른 설명은 금지.")


def extract_one(image_bytes, want_keyword, cfg):
    if not want_keyword:
        raw = _call_gemini(cfg["model"], cfg["thinking"], cfg["hook_prompt"], image_bytes, cfg["timeout"])
        return {"hook": raw.strip() or cfg["placeholder"], "keyword": "", "note": ""}
    hook, kw, note = None, "", ""
    prompt = _combined_prompt(cfg)
    for _ in range(1 + cfg["parse_retries"]):
        raw = _call_gemini(cfg["model"], cfg["thinking"], prompt, image_bytes, cfg["timeout"])
        if cfg["delimiter"] in (raw or ""):
            c, k = raw.split(cfg["delimiter"], 1)
            hook, kw = c.strip(), k.strip()
            break
    if hook is None:
        hook, kw, note = cfg["placeholder"], "", "구분자 파싱 실패"
    hook = hook or cfg["placeholder"]
    tries = 0
    while not _HANGUL_WORD.match(kw or "") and tries < cfg["kw_retries"]:
        tries += 1
        raw = _call_gemini(cfg["kw_retry_model"], cfg["kw_retry_thinking"], _keyword_only_prompt(cfg),
                           image_bytes, cfg["timeout"])
        kw = raw.strip().splitlines()[0].strip() if raw.strip() else ""
    if not _HANGUL_WORD.match(kw or ""):
        kw, note = "", (note + " 키워드 실패").strip()
    return {"hook": hook, "keyword": kw, "note": note}


def drive_file_id(url):
    for marker in ("/file/d/", "/folders/", "/d/"):
        if marker in (url or ""):
            return url.split(marker, 1)[1].split("/", 1)[0].split("?", 1)[0]
    return url if url and "/" not in url else None


def drive_download(url):
    fid = drive_file_id(url)
    if not fid:
        raise RuntimeError("input_url에서 파일ID 추출 실패")
    return _request_with_retry("GET", f"https://www.googleapis.com/drive/v3/files/{fid}",
                               params={"alt": "media"}, timeout=60).content


def process(item, cfg):
    data = drive_download(item["a"]["input_url"])
    return extract_one(data, item["want_kw"], cfg)


def run_stage(items, workers, cfg):
    """반환: (results dict idx->결과, 통신성 실패 idx 리스트)"""
    out, transient = {}, []
    with concurrent.futures.ThreadPoolExecutor(max_workers=max(1, workers)) as ex:
        futs = {ex.submit(process, it, cfg): i for i, it in enumerate(items)}
        for fut in concurrent.futures.as_completed(futs):
            i = futs[fut]
            try:
                out[i] = fut.result()
            except (TransientError, requests.RequestException) as e:
                transient.append(i)
                out[i] = {"hook": "", "keyword": "", "note": f"통신 실패: {type(e).__name__}"}
            except Exception as e:
                out[i] = {"hook": "", "keyword": "", "note": f"오류: {type(e).__name__}: {str(e)[:150]}"}
    return out, transient


def write_cells(sheet_id, updates):
    if not updates:
        return
    body = {"valueInputOption": "RAW",
            "data": [{"range": a1, "values": [[v]]} for a1, v in updates]}
    _request_with_retry("POST", f"{SHEETS_API_BASE}/{sheet_id}/values:batchUpdate", json=body)


def main():
    sheet_id = os.environ["SHEET_ID"].strip()
    lo, hi = _to_int(os.environ.get("ID_MIN")), _to_int(os.environ.get("ID_MAX"))
    id_range = (lo, hi) if lo and hi else None
    t0 = time.time()
    cfg = load_cfg()
    print(f"설정: model={cfg['model']} workers={cfg['workers']} 범위={id_range or '빈칸 전체'}")

    total_saved = total_kw = 0
    failures = []
    skipped_ids = set()  # 이번 실행에서 실패해 건너뛴 asset(무한루프 방지)
    first = True
    while True:
        if time.time() - t0 > MAX_RUN_SEC:
            print("시간 제한 임박 - 중단(남은 건 재실행하면 이어서 처리)")
            break
        # 매 청크마다 시트를 새로 읽는다: 행 번호가 바뀌었거나 GUI/다른 실행이 이미 채운 건 건드리지 않기 위함
        a_head, a_rows = _sheet_values(sheet_id, "ASSETS")
        c_head, c_rows = _sheet_values(sheet_id, "CONTENTS")
        c_by_id = {c.get("content_id"): (rn, c) for rn, c in c_rows}
        pending = []
        for rn, a in a_rows:
            if a.get("asset_id") in skipped_ids or not a.get("asset_id") or not a.get("input_url"):
                continue
            if (a.get("source_hook") or "").strip():
                continue
            if (a.get("processing_mode") or "").strip().upper() == "V":
                continue
            cc = c_by_id.get(a.get("content_id"))
            if not cc:
                continue
            if id_range and not (id_range[0] <= _to_int(cc[1].get("local_file_id")) <= id_range[1]):
                continue
            pending.append((rn, a, cc))
        if not pending:
            break
        if first:
            print(f"OCR 대상 {len(pending)}건")
            first = False
        chunk = pending[:CHUNK]
        kw_asked, items = set(), []
        for rn, a, (crn, c) in chunk:
            want = (_to_int(a.get("slide_order")) == 1 and not (c.get("keyword") or "").strip()
                    and c.get("content_id") not in kw_asked)
            if want:
                kw_asked.add(c.get("content_id"))
            items.append({"a": a, "want_kw": want})
        res, transient = run_stage(items, cfg["workers"], cfg)
        if transient:
            r2, _ = run_stage([items[i] for i in transient], cfg["fallback_workers"], cfg)
            for j, i in enumerate(transient):
                res[i] = r2[j]
        col_hook = _col_letter(a_head.index("source_hook"))
        col_kw = _col_letter(c_head.index("keyword")) if "keyword" in c_head else None
        updates = []
        for i, (rn, a, (crn, c)) in enumerate(chunk):
            r = res[i]
            if r.get("hook"):
                updates.append((f"ASSETS!{col_hook}{rn}", r["hook"]))
                total_saved += 1
            else:
                skipped_ids.add(a["asset_id"])
                failures.append(f"{a['asset_id']}: {r.get('note')}")
            if r.get("keyword") and col_kw and _to_int(a.get("slide_order")) == 1 \
                    and not (c.get("keyword") or "").strip():
                updates.append((f"CONTENTS!{col_kw}{crn}", r["keyword"]))
                total_kw += 1
        write_cells(sheet_id, updates)
        print(f"청크 저장: 누적 {total_saved}건, 키워드 {total_kw}건, 실패 {len(failures)}건 "
              f"(남은 대상 약 {len(pending) - len(chunk)})")
    print(f"완료: 저장 {total_saved}건, 키워드 {total_kw}건, 실패 {len(failures)}건")
    for f in failures[:30]:
        print("  실패:", f)
    if failures:
        print("실패 건은 빈칸으로 남음 - 재실행 시 다시 포함")


if __name__ == "__main__":
    sys.exit(main())
