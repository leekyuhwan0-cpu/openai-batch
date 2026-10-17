"""A/R(실사/AI생성) 판정 사전실행 서버 처리(2026-10-17). 로컬 mode_detect_ui.py(GPU/CPU 서브프로세스)를 대체.

ASSETS에서 processing_mode가 R/A/V가 아니고 ar_pred가 빈 이미지를 Drive에서 받아 HF Space(CPU,
community-forensics-frontier)의 /detect_batch로 판정 -> 배치마다 즉시 ASSETS.ar_pred에 "A"/"R"을 기록.
processing_mode(확정값)는 건드리지 않는다: 메인 UI가 ar_pred를 "미저장" 배지로 보여주고, 사용자가
확인 후 "모드판별 저장"을 눌러야 processing_mode로 들어간다. 중간에 끊겨도 쓴 분량은 보존되고
실패 건은 빈칸으로 남아 재실행 시 재포함된다.

Space URL 목록/배치크기는 메인시트 v2_MODEL_REGISTRY(mode_detect_cloud 행) params에서 읽는다.
Space 하나는 한 번에 한 배치만 처리하므로(CPU 2코어) 병렬화는 space_urls를 늘려서 한다."""
import concurrent.futures
import os
import sys
import tempfile
import time

from poll_batches import (MAIN_SHEET_ID, SHEETS_API_BASE, _col_letter, _request_with_retry,
                          _sheet_values)
from ocr_run import _to_int, drive_download, write_cells

import json

MAX_RUN_SEC = 5 * 3600 + 1800  # 6시간 제한 전에 스스로 멈춤(남은 건 재실행으로 이어감)
TASK = "mode_detect_cloud"
PRED_COL = "ar_pred"
DOWNLOAD_WORKERS = 8


def load_cfg():
    _, rows = _sheet_values(MAIN_SHEET_ID, "v2_MODEL_REGISTRY")
    for _, r in rows:
        if r.get("task_category") == TASK and str(r.get("is_active")).strip().upper() == "TRUE":
            p = json.loads(r.get("params") or "{}")
            urls = p.get("space_urls") or ([r["space_url"]] if r.get("space_url") else [])
            return {"spaces": [u.rsplit("huggingface.co/spaces/", 1)[-1].strip("/") for u in urls],
                    "batch": int(p.get("batch_size", 30)),
                    "api_name": p.get("api_name", "/detect_batch"),
                    "timeout": int(p.get("timeout_sec", 900))}
    raise SystemExit(f"v2_MODEL_REGISTRY에 활성 {TASK} 행이 없음")


def ensure_pred_column(sheet_id, a_head):
    """ASSETS에 ar_pred 헤더가 없으면 맨 오른쪽 빈 칸에 추가(기존 컬럼/데이터는 건드리지 않음)."""
    if PRED_COL in a_head:
        return a_head
    write_cells(sheet_id, [(f"ASSETS!{_col_letter(len(a_head))}1", PRED_COL)])
    print(f"ASSETS에 {PRED_COL} 컬럼 추가({_col_letter(len(a_head))})")
    return a_head + [PRED_COL]


def pending_assets(sheet_id, id_range, skipped):
    a_head, a_rows = _sheet_values(sheet_id, "ASSETS")
    a_head = ensure_pred_column(sheet_id, a_head)
    _, c_rows = _sheet_values(sheet_id, "CONTENTS")
    c_by_id = {c.get("content_id"): c for _, c in c_rows}
    out = []
    for rn, a in a_rows:
        aid = a.get("asset_id")
        if not aid or aid in skipped or not a.get("input_url"):
            continue
        if (a.get("processing_mode") or "").strip().upper() in ("R", "A", "V"):
            continue
        if (a.get(PRED_COL) or "").strip():
            continue
        c = c_by_id.get(a.get("content_id"))
        if id_range and not (c and id_range[0] <= _to_int(c.get("local_file_id")) <= id_range[1]):
            continue
        out.append((rn, a))
    return a_head, out


def _ext(b):
    return ".png" if b[:4] == b"\x89PNG" else ".webp" if b[:4] == b"RIFF" else ".jpg"


def run_batch(client, items, tmpdir, api_name):
    """items: [(rn, asset)] -> [(rn, asset, 'A'|'R'|None, 사유)]. 다운로드 실패 건은 빼고 보낸다."""
    from gradio_client import handle_file
    paths, ok = [], []

    def fetch(item):
        rn, a = item
        data = drive_download(a["input_url"])
        p = os.path.join(tmpdir, f"{a['asset_id']}{_ext(data)}")
        with open(p, "wb") as f:
            f.write(data)
        return rn, a, p

    with concurrent.futures.ThreadPoolExecutor(max_workers=DOWNLOAD_WORKERS) as dl:
        futs = [(it, dl.submit(fetch, it)) for it in items]
        for (rn, a), fut in futs:  # items 순서 유지
            try:
                paths.append(fut.result())
            except Exception as e:
                ok.append((rn, a, None, f"다운로드 실패: {type(e).__name__}"))
    res = []
    if paths:
        r = client.predict([handle_file(p) for _, _, p in paths], api_name=api_name)
        if not isinstance(r, list) or len(r) != len(paths):
            raise RuntimeError(f"Space 응답 개수 불일치: {len(paths)}장 -> {type(r).__name__} {len(r) if hasattr(r, '__len__') else ''}")
        for (rn, a, _), x in zip(paths, r):
            lab = x.get("label")
            res.append((rn, a, {"ai": "A", "real": "R"}.get(lab), None if lab in ("ai", "real") else f"판정 오류: {x.get('error', lab)}"))
    return ok + res


def main():
    from gradio_client import Client
    sheet_id = os.environ["SHEET_ID"].strip()
    lo, hi = _to_int(os.environ.get("ID_MIN")), _to_int(os.environ.get("ID_MAX"))
    id_range = (lo, hi) if lo and hi else None
    cfg = load_cfg()
    token = os.environ.get("HF_TOKEN") or None
    clients = [Client(s, token=token) for s in cfg["spaces"]]
    print(f"설정: spaces={cfg['spaces']} batch={cfg['batch']} 범위={id_range or '빈칸 전체'}")

    t0 = time.time()
    total, failures, skipped = 0, [], set()
    first = True
    with tempfile.TemporaryDirectory() as tmpdir, \
            concurrent.futures.ThreadPoolExecutor(max_workers=len(clients)) as ex:
        while time.time() - t0 < MAX_RUN_SEC:
            # 매 라운드마다 시트를 새로 읽는다: 다른 실행/GUI가 이미 채운 건 건드리지 않기 위함
            a_head, pending = pending_assets(sheet_id, id_range, skipped)
            if not pending:
                break
            if first:
                print(f"판정 대상 {len(pending)}건")
                first = False
            # Space 하나당 한 배치씩 동시에(Space는 한 번에 한 배치만 처리)
            take = pending[: cfg["batch"] * len(clients)]
            batches = [take[i:i + cfg["batch"]] for i in range(0, len(take), cfg["batch"])]
            futs = {ex.submit(run_batch, clients[i], b, tmpdir, cfg["api_name"]): b
                    for i, b in enumerate(batches)}
            col = _col_letter(a_head.index(PRED_COL))
            for fut in concurrent.futures.as_completed(futs):
                b = futs[fut]
                try:
                    out = fut.result()
                except Exception as e:
                    for _, a in b:
                        skipped.add(a["asset_id"])
                        failures.append(f"{a['asset_id']}: {type(e).__name__}: {str(e)[:120]}")
                    print(f"배치 실패({len(b)}장): {type(e).__name__}: {str(e)[:150]}")
                    continue
                updates = []
                for rn, a, lab, why in out:
                    if lab:
                        updates.append((f"ASSETS!{col}{rn}", lab))
                        total += 1
                    else:
                        skipped.add(a["asset_id"])
                        failures.append(f"{a['asset_id']}: {why}")
                write_cells(sheet_id, updates)
                print(f"배치 저장 {len(updates)}장: 누적 {total}건, 실패 {len(failures)}건, "
                      f"경과 {time.time() - t0:.0f}s (남은 대상 약 {len(pending) - len(take)})")
    print(f"완료: 저장 {total}건, 실패 {len(failures)}건, {time.time() - t0:.0f}s")
    for f in failures[:30]:
        print("  실패:", f)
    if failures:
        print("실패 건은 빈칸으로 남음 - 재실행 시 다시 포함")


if __name__ == "__main__":
    sys.exit(main())
