"""
GitHub Actions에서 주기적으로 실행되는 진입점.
5개 구글 스프레드시트(CS예약/인플루언서/콘텐츠/바이럴포스팅/체험단) 전체를
구글 시트 API로 읽어 build_data.py와 동일한 집계 로직으로 data/summary.json을
통째로 재생성한다 (로컬 xlsx 전체 재생성 경로인 build()와 동일한 결과가 나와야 함).

원본 시트 내용은 이 프로세스 메모리 안에서만 다뤄지고, 집계 결과 외에는
아무것도 파일로 저장하거나 커밋하지 않는다.
"""
import datetime
import json
import os
import sys

from google.oauth2 import service_account
from googleapiclient.discovery import build as build_service

sys.path.insert(0, os.path.dirname(__file__))
from build_data import (  # noqa: E402
    COLS,
    CONTENT_SHEET,
    EXPERIENCE_SHEETS,
    INFLUENCER_SHEETS,
    MEGA_ROI_SHEET,
    OUTPUT_PATH,
    SHEET_NAME,
    VIRAL_SHEETS,
    build_content_performance,
    build_cs_bookings,
    build_experience_bookings,
    build_influencer_funnel,
    build_mega_roi,
    build_viral_posting,
    find_header_row,
)

# 구글 시트 API를 valueRenderOption=UNFORMATTED_VALUE로 호출하면 날짜/시간 셀이
# (openpyxl처럼 datetime 객체가 아니라) 1899-12-30 기준 일련번호로 온다.
# 각 빌더는 openpyxl 타입(datetime/time)을 기대하므로, 빌더에 넘기기 전에
# "날짜/시간으로 알려진 컬럼"만 골라 되돌려 놓는다. 그 외 숫자 컬럼(팔로워 수,
# 견적 등)은 UNFORMATTED_VALUE 그대로 두어야 parse_numeric류가 정상 동작한다.
GSHEET_EPOCH = datetime.datetime(1899, 12, 30)


def serial_to_datetime(v):
    if isinstance(v, bool) or not isinstance(v, (int, float)):
        return v
    days = int(v)
    dt = GSHEET_EPOCH + datetime.timedelta(days=days)
    frac = v - days
    if frac:
        dt += datetime.timedelta(seconds=round(frac * 86400))
    return dt


def serial_to_time(v):
    if isinstance(v, bool) or not isinstance(v, (int, float)):
        return v
    seconds = round((v % 1) * 86400)
    return (datetime.datetime.min + datetime.timedelta(seconds=seconds)).time()


def fix_date_time_columns(rows, header_idx, date_idx=(), time_idx=()):
    if not date_idx and not time_idx:
        return rows
    fixed = []
    for i, r in enumerate(rows):
        if i <= header_idx:
            fixed.append(r)
            continue
        r = list(r)
        for p in date_idx:
            if p < len(r):
                r[p] = serial_to_datetime(r[p])
        for p in time_idx:
            if p < len(r):
                r[p] = serial_to_time(r[p])
        fixed.append(r)
    return fixed


def fix_named_date_time_columns(rows, header_marker, date_names=(), time_names=()):
    header_idx = find_header_row(rows, header_marker)
    if header_idx >= len(rows):
        return rows
    header = rows[header_idx]
    col_idx = {}
    for i, v in enumerate(header):
        if v and str(v).strip() not in col_idx:
            col_idx[str(v).strip()] = i
    date_idx = [col_idx[n] for n in date_names if n in col_idx]
    time_idx = [col_idx[n] for n in time_names if n in col_idx]
    return fix_date_time_columns(rows, header_idx, date_idx, time_idx)

SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
KEY_PATH = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS")

# 로컬 xlsx 스냅샷과 라이브 시트의 탭 이름이 다른 경우만 매핑 (라이브 시트 메타데이터로 확인됨,
# 2026-08-30). CS/콘텐츠/바이럴/체험단은 전부 로컬 이름과 정확히 일치해서 매핑 불필요.
LIVE_SHEET_NAME = {
    "중국 인풀루언서체험단": "중국 인풀루언서/체험단",
    "대만 인풀루언서체험단": "대만 인풀루언서/체험단",
    "일본 인풀루언서체험단": "일본 인풀루언서/체험단",
    "영미 인풀루언서체험단": "영미 인풀루언서/체험단",
}

SPREADSHEETS = {
    "cs": {"id": "1OH1krRpZ6dDnAqoJFWrdMInOXJPa2gSuePbm_dsPEZk", "sheets": [SHEET_NAME]},
    "influencer": {
        "id": "1yEWyVVPW5cYQkjJ-EAlx37veDfZFLiCQG3nILkcgmH8",
        "sheets": INFLUENCER_SHEETS + [MEGA_ROI_SHEET],
    },
    "content": {"id": "1N2ls94wtqTzRJdPprDb-7KgGXTIDOz6W9aAuoZ5PETI", "sheets": [CONTENT_SHEET]},
    "viral": {"id": "1kNf-4A90R2QJ8vjh-_UNQ3G5rcXrLeNqO2BKli4fhCA", "sheets": VIRAL_SHEETS},
    "experience": {
        "id": "14nyN_DPfNYSM2oCLBszuCWywLXcP_X_eSebNl3w4hlQ",
        "sheets": EXPERIENCE_SHEETS,
    },
}


def fetch_all_rows(values_api):
    all_rows = {}
    for group in SPREADSHEETS.values():
        for local_name in group["sheets"]:
            live_name = LIVE_SHEET_NAME.get(local_name, local_name)
            result = values_api.get(
                spreadsheetId=group["id"],
                range=f"'{live_name}'!A:AZ",
                valueRenderOption="UNFORMATTED_VALUE",
            ).execute()
            all_rows[local_name] = result.get("values", [])

    # CS: 타이틀행(0)+헤더행(1) 뒤부터 데이터, 날짜/시간 컬럼은 고정 인덱스(COLS)로 알려져 있음
    all_rows[SHEET_NAME] = fix_date_time_columns(
        all_rows[SHEET_NAME],
        header_idx=1,
        date_idx=[COLS["인입일자"], COLS["예약일"]],
        time_idx=[COLS["예약시간"]],
    )
    # 체험단: 헤더 위치가 시트마다 다름(find_header_row로 찾음), 컬럼명 기준
    for name in EXPERIENCE_SHEETS:
        all_rows[name] = fix_named_date_time_columns(
            all_rows[name], "체험단국적", date_names=["예약 날짜"], time_names=["예약 시간"]
        )
    # 바이럴 포스팅: 마찬가지로 헤더 위치가 시트마다 다름
    for name in VIRAL_SHEETS:
        all_rows[name] = fix_named_date_time_columns(
            all_rows[name], "순번", date_names=["업로드 날짜"]
        )

    return all_rows


def sync():
    if not KEY_PATH:
        raise SystemExit("GOOGLE_APPLICATION_CREDENTIALS 환경변수가 설정되어 있지 않음")

    creds = service_account.Credentials.from_service_account_file(KEY_PATH, scopes=SCOPES)
    service = build_service("sheets", "v4", credentials=creds)
    rows_by_sheet = fetch_all_rows(service.spreadsheets().values())

    daily, staff, procedures, slots, n_total, n_used = build_cs_bookings(rows_by_sheet[SHEET_NAME])
    influencer, inf_total, inf_used = build_influencer_funnel(rows_by_sheet)
    mega_roi, roi_total, roi_used = build_mega_roi(rows_by_sheet)
    content_uploads, content_engagement, content_total, content_used = build_content_performance(
        rows_by_sheet
    )
    experience_bookings, experience_slots, exp_total, exp_used = build_experience_bookings(
        rows_by_sheet
    )
    viral_uploads, viral_engagement, viral_total, viral_used = build_viral_posting(rows_by_sheet)

    new_output = {
        "sourceRowCount": n_total,
        "usedRowCount": n_used,
        "daily": daily,
        "staffMonthly": staff,
        "procedureMonthly": procedures,
        "bookingSlots": slots,
        "influencerSourceRowCount": inf_total,
        "influencerUsedRowCount": inf_used,
        "influencerFunnel": influencer,
        "megaRoiSourceRowCount": roi_total,
        "megaRoiUsedRowCount": roi_used,
        "megaRoi": mega_roi,
        "contentSourceRowCount": content_total,
        "contentUsedRowCount": content_used,
        "contentUploads": content_uploads,
        "contentEngagement": content_engagement,
        "experienceSourceRowCount": exp_total,
        "experienceUsedRowCount": exp_used,
        "experienceBookings": experience_bookings,
        "experienceSlots": experience_slots,
        "viralPostingSourceRowCount": viral_total,
        "viralPostingUsedRowCount": viral_used,
        "viralPostingUploads": viral_uploads,
        "viralPostingEngagement": viral_engagement,
    }

    # 안전장치: build_data.py의 build()와 동일한 PII 검사를 그대로 적용
    blob = json.dumps(new_output, ensure_ascii=False)
    forbidden_keys = [
        "이름", "성함", "차트번호", "전화번호", "생년월일", "위챗 ID",
        "아이디", "인적사항", "비고", "콘텐츠 링크", "계정 링크",
        "콘텐츠 업로드 링크", "Links", "xhslink", "instagram.com",
        "tiktok.com", "xiaohongshu.com", "왓으앱", "카톡",
        "포스팅 링크", "threads.com", "xhslink.com",
    ]
    for k in forbidden_keys:
        assert k not in blob, f"PII 관련 키가 출력에 포함됨: {k}"

    with open(OUTPUT_PATH, encoding="utf-8") as f:
        old_output = json.load(f)
    old_output.pop("generatedAt", None)

    if old_output == new_output:
        print("변경 없음 - summary.json을 다시 쓰지 않음")
        return

    new_output["generatedAt"] = datetime.datetime.now().isoformat(timespec="seconds")
    with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
        json.dump(new_output, f, ensure_ascii=False, indent=2)

    print(f"CS 예약: 원본 {n_total}, 사용 {n_used}")
    print(f"인플루언서: 원본 {inf_total}, 사용 {inf_used}")
    print(f"메가ROI: 원본 {roi_total}, 사용 {roi_used}")
    print(f"콘텐츠: 원본 {content_total}, 사용 {content_used}")
    print(f"체험단: 원본 {exp_total}, 사용 {exp_used}")
    print(f"바이럴포스팅: 원본 {viral_total}, 사용 {viral_used}")
    print(f"갱신됨: {OUTPUT_PATH}")


if __name__ == "__main__":
    sync()
