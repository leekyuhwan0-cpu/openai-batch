"""매일 1회: 선택된 인스타 채널의 '어제(KST)' 게시물 중 좋아요 >= 기준 이상을 VIRAL_DAILY 탭에 기록하고
이미지/영상을 Drive 채널별 폴더에 저장. 공유수는 취득 불가라 빈칸.
채널 선택: 결과시트 VIRAL_CHANNELS 탭의 '활성' 열에 O 표시한 채널만 수집.
"""
import os, sys, time, json, datetime as dt
from urllib.parse import quote
import requests

import poll_batches as pb  # 구글 토큰/시트 헬퍼 재사용

RESULT_SHEET = "1Kdh0gTJYCfYvsUCa5Wqo6a9ITX140q9ovOi9n4HvyGU"
MAIN_SHEET = "1IK7MbiTFJiV_ofeco9FwYs7UFY_f4QHHIDZJb7T_4-o"  # SOURCE_ACCOUNTS(카테고리/언어 원본)
DAILY_TAB = "VIRAL_DAILY"
CONTENTS_TAB = "CONTENTS"
ASSETS_TAB = "ASSETS"
CHANNEL_TAB = "VIRAL_CHANNELS"
SC_KEY = os.environ["SCRAPECREATORS_API_KEY"]
SC_URL = "https://api.scrapecreators.com/v2/instagram/user/posts"
MIN_LIKES = int(os.environ.get("VIRAL_MIN_LIKES", "1000"))
DRIVE_ROOT_ID = os.environ.get("VIRAL_DRIVE_ROOT_ID", "")
DRIVE_ROOT_NAME = "바이럴수집"
MAX_PAGES = 3
PIN_SLOTS = 4  # 상단 고정글 자리(ac미디어는 맨 위 4개가 옛날 글이었음)
KST = dt.timezone(dt.timedelta(hours=9))
DRY = os.environ.get("VIRAL_DRY") == "1"  # 1이면 시트 기록/Drive 업로드 없이 로그만
DRIVE = "https://www.googleapis.com/drive/v3/files"


# ---- 파일명 규칙 (확정 전 임시. 바꿀 때 여기만 수정) ----
def file_name(channel, date_str, code, idx, ext):
    suffix = "" if idx == 0 else f"-{idx}"
    return f"{channel}_{date_str}_{code}{suffix}.{ext}"


def _sheets(method, path, **kw):
    return pb._request_with_retry(method, f"{pb.SHEETS_API_BASE}/{RESULT_SHEET}{path}", **kw)


def write_cells(updates):
    """updates: [(A1범위, [[값]])]"""
    if DRY or not updates:
        return
    _sheets("POST", "/values:batchUpdate", json={
        "valueInputOption": "RAW",
        "data": [{"range": r, "values": v} for r, v in updates]})


# ---- ScrapeCreators ----
def fetch_posts(handle, y_start, y_end):
    """어제(KST) 게시물. 상단 고정 최대 3개 때문에 index>=3에서 어제 이전 글이 나오면 중단."""
    out, cursor = [], None
    for page in range(MAX_PAGES):
        params = {"handle": handle}
        if cursor:
            params["cursor"] = cursor
        r = requests.get(SC_URL, headers={"x-api-key": SC_KEY}, params=params, timeout=60)
        if r.status_code != 200:
            raise RuntimeError(f"HTTP {r.status_code} {r.text[:150]}")
        d = r.json()
        items = d.get("items") or []
        stop, run = False, 0
        for i, it in enumerate(items):
            t = it.get("taken_at") or 0
            if y_start <= t < y_end:
                out.append(it)
                run = 0
            elif t < y_start and (page > 0 or i >= PIN_SLOTS):
                run += 1  # 어제 이전 글이 연속 2개 나오면 더 볼 필요 없음
                if run >= 2:
                    stop = True
            else:
                run = 0
        cursor = d.get("profile_grid_items_cursor")
        if stop or not d.get("more_available") or not cursor or not items:
            break
        if (items[-1].get("taken_at") or 0) < y_start:
            break
    return out


def media_urls(it):
    """-> (유형, 이미지URL들, 영상URL들)"""
    def img(m):
        c = (m.get("image_versions2") or {}).get("candidates") or []
        return c[0]["url"] if c else None

    mt = it.get("media_type")
    imgs, vids = [], []
    if mt == 8:
        for m in it.get("carousel_media") or []:
            u = img(m)
            if u:
                imgs.append(u)
            if m.get("video_url"):
                vids.append(m["video_url"])
        kind = "카드뉴스"
    elif mt == 2:
        u = img(it)
        if u:
            imgs.append(u)
        if it.get("video_url"):
            vids.append(it["video_url"])
        kind = "영상"
    else:
        u = img(it)
        if u:
            imgs.append(u)
        kind = "사진"
    return kind, imgs, vids


# ---- Drive ----
def find_or_create_folder(name, parent):
    esc = name.replace("\\", "\\\\").replace("'", "\\'")
    q = f"name='{esc}' and mimeType='application/vnd.google-apps.folder' and trashed=false"
    if parent:
        q += f" and '{parent}' in parents"
    r = pb._request_with_retry("GET", DRIVE, params={"q": q, "fields": "files(id)", "pageSize": 1}).json()
    if r.get("files"):
        return r["files"][0]["id"]
    body = {"name": name, "mimeType": "application/vnd.google-apps.folder"}
    if parent:
        body["parents"] = [parent]
    return pb._request_with_retry("POST", DRIVE, json=body).json()["id"]


def upload_bytes(folder_id, name, data, mime):
    meta = {"name": name, "parents": [folder_id]}
    boundary = "vrl%d" % int(time.time() * 1000)
    body = (f"--{boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n"
            f"{json.dumps(meta)}\r\n--{boundary}\r\nContent-Type: {mime}\r\n\r\n").encode() \
        + data + f"\r\n--{boundary}--".encode()
    r = pb._request_with_retry(
        "POST", "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id",
        headers={"Content-Type": f"multipart/related; boundary={boundary}"}, data=body, timeout=300)
    return r.json()["id"]


def drive_link(file_id):
    return f"https://drive.google.com/file/d/{file_id}/view"


def _to_int(v):
    try:
        return int(float(str(v).strip()))
    except (ValueError, TypeError):
        return 0


def next_contents_slot():
    """CONTENTS에서 (다음 local_file_id, 채울 첫 빈 행 번호). content_id가 빈 행(사전 서식 행)은 재사용."""
    _, rows = pb._sheet_values(RESULT_SHEET, CONTENTS_TAB)
    max_num, free_row, last_row = 0, None, 1
    for rn, d in rows:
        last_row = max(last_row, rn)
        if str(d.get("content_id", "")).strip():
            max_num = max(max_num, _to_int(d.get("local_file_id")))
        elif free_row is None:
            free_row = rn
    return max_num + 1, (free_row or last_row + 1)


def append_asset_rows(rows):
    if DRY or not rows:
        return
    _sheets("POST", f"/values/{quote(ASSETS_TAB)}!A1:append",
            params={"valueInputOption": "RAW", "insertDataOption": "INSERT_ROWS"},
            json={"values": rows})


def download(url):
    r = requests.get(url, timeout=120)
    r.raise_for_status()
    return r.content


def main():
    now = dt.datetime.now(KST)
    today = now.date()
    yday = today - dt.timedelta(days=1)
    y_start = int(dt.datetime(yday.year, yday.month, yday.day, tzinfo=KST).timestamp())
    y_end = y_start + 86400
    date_str = yday.strftime("%Y%m%d")
    print(f"대상일(KST) {yday}  좋아요>={MIN_LIKES}  DRY={DRY}")

    headers, ch_rows = pb._sheet_values(RESULT_SHEET, CHANNEL_TAB)
    col = {h: i for i, h in enumerate(headers)}
    active = [(r, d) for r, d in ch_rows if str(d.get("활성", "")).strip().upper() in ("O", "TRUE", "1")]
    print(f"활성 채널 {len(active)}개")
    if not active:
        return

    # 카테고리/언어는 메인 시트 SOURCE_ACCOUNTS(채널명=channel_id)에서 조회 - 한 곳에서만 관리
    sa_info = {}
    for _, sd in pb._sheet_values(MAIN_SHEET, "SOURCE_ACCOUNTS")[1]:
        sa_info[str(sd.get("channel_id", "")).strip()] = (sd.get("category", ""), sd.get("language", ""))

    seen = {d.get("게시물링크") for _, d in pb._sheet_values(RESULT_SHEET, DAILY_TAB)[1]}
    root = DRIVE_ROOT_ID or (None if DRY else find_or_create_folder(DRIVE_ROOT_NAME, None))

    next_num, content_row = (1, 2) if DRY else next_contents_slot()
    new_rows, asset_rows, errors = [], [], 0
    for r, d in active:
        name, handle = d.get("채널명") or d.get("핸들"), d.get("핸들")
        try:
            posts = fetch_posts(handle, y_start, y_end)
            hits = [p for p in posts if (p.get("like_count") or 0) >= MIN_LIKES]
            folder = None
            for p in hits:
                code = p["code"]
                link = f"https://www.instagram.com/p/{code}/"
                if link in seen:
                    continue
                kind, imgs, vids = media_urls(p)
                saved, img_links = [], []
                if not DRY:
                    folder = folder or find_or_create_folder(name, root)
                    for i, u in enumerate(imgs):
                        fn = file_name(name, date_str, code, i, "jpg")
                        fid = upload_bytes(folder, fn, download(u), "image/jpeg")
                        saved.append(fn)
                        img_links.append((fn, drive_link(fid)))
                    for i, u in enumerate(vids):
                        fn = file_name(name, date_str, code, i, "mp4")
                        upload_bytes(folder, fn, download(u), "video/mp4")
                        saved.append(fn)
                taken = dt.datetime.fromtimestamp(p["taken_at"], KST).strftime("%Y-%m-%d %H:%M")
                new_rows.append([
                    today.strftime("%Y-%m-%d"), name, sa_info.get(name, ("", ""))[1] or d.get("언어", ""), taken,
                    p.get("like_count") or 0, p.get("comment_count") or 0, p.get("play_count") or "",
                    kind, len(imgs), link, ((p.get("caption") or {}).get("text") or "")[:500],
                    "\n".join(imgs), "\n".join(vids), "O", "O", ", ".join(saved)])
                seen.add(link)
                if not DRY and img_links:
                    # CONTENTS 1행 + ASSETS 슬라이드별 N행 (v2 파이프라인 포맷, 채널 무관 전역 일련번호)
                    cid = f"{name}_{next_num}"
                    caption = ((p.get("caption") or {}).get("text") or "")
                    write_cells([
                        (f"{CONTENTS_TAB}!A{content_row}:B{content_row}", [[
                            sa_info.get(name, ("", ""))[1] or d.get("언어", ""), sa_info.get(name, ("", ""))[0]]]),
                        (f"{CONTENTS_TAB}!C{content_row}:H{content_row}", [[
                            name, cid, next_num, caption, today.strftime("%Y-%m-%d"), "NEW"]])])
                    for i, (fn, url) in enumerate(img_links):
                        asset_rows.append([f"{cid}_{i + 1}", cid, fn, i + 1, url])
                    next_num += 1
                    content_row += 1
            status = f"OK 어제{len(posts)}/통과{len(hits)}"
        except Exception as e:
            errors += 1
            status = f"ERR {str(e)[:120]}"
        print(name, status)
        write_cells([(f"{CHANNEL_TAB}!{pb._col_letter(col['마지막수집'])}{r}:{pb._col_letter(col['마지막상태'])}{r}",
                      [[today.strftime("%Y-%m-%d"), status]])])

    append_asset_rows(asset_rows)
    if new_rows and not DRY:
        _sheets("POST", f"/values/{quote(DAILY_TAB)}!A1:append",
                params={"valueInputOption": "RAW", "insertDataOption": "INSERT_ROWS"},
                json={"values": new_rows})
    print(f"신규 {len(new_rows)}건, 채널오류 {errors}건")
    if errors and errors == len(active):
        sys.exit(1)


if __name__ == "__main__":
    main()
