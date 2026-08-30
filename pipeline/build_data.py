"""
원본 CS 상담 로그(개인정보 포함)를 읽어, 개인 식별이 불가능한 수준까지
집계(aggregate)한 결과만 data/summary.json 으로 출력한다.

절대 하지 않는 것: 원본 행 단위 레코드 출력, 이름/차트번호/전화번호/생년월일 등
PII 컬럼을 읽거나 저장하는 것.
"""
import json
import re
import datetime
import collections
import os

import openpyxl

SOURCE_PATH = os.path.expanduser(
    "~/Desktop/Dashboard_data/중화권 예약 &DB리스트.xlsx"
)
SHEET_NAME = "예약문의통합"

INFLUENCER_PATH = os.path.expanduser(
    "~/Desktop/Dashboard_data/마케팅팀 인플루언서 및 2차동의 관리시트.xlsx"
)
INFLUENCER_SHEETS = [
    "중국 인풀루언서체험단",
    "대만 인풀루언서체험단",
    "일본 인풀루언서체험단",
    "영미 인풀루언서체험단",
    "줄기세포 시딩 인풀루언서",
]
MEGA_ROI_SHEET = "메가인플루언서 협업ROI"

CONTENT_PATH = os.path.expanduser(
    "~/Desktop/Dashboard_data/HB_ 콘텐츠 업로드 관리.xlsx"
)
CONTENT_SHEET = "전체통합"

EXPERIENCE_PATH = os.path.expanduser(
    "~/Desktop/Dashboard_data/힐링브리즈 체험단 정리.xlsx"
)
EXPERIENCE_SHEETS = ["중국 체험단 예약", "대만체험단 예약", "일본 체험단 예약", "영미 체험단 예약"]
EXPERIENCE_STATUS_MAP = {
    "방문완": "방문완료", "방문완료": "방문완료",
    "예약완": "예약완료", "예약완료": "예약완료",
    "취소": "취소",
    "예약요청": "예약요청",
}

VIRAL_PATH = os.path.expanduser(
    "~/Desktop/Dashboard_data/HB_중화권 바이럴 포스팅.xlsx"
)
VIRAL_SHEETS = ["샤오홍슈 막계정", "스레드 막계정"]  # '루후 막계정' 시트는 비어 있어 제외
STAGE_MAP = {
    "답변없음": "답변없음",
    "협찬거절": "협찬거절",
    "컨택중": "협의중",
    "스케줄조율": "협의중",
    "예약완료": "예약완료",
    "방문완료": "방문완료",
    "업로드완료": "업로드완료",
}

OUTPUT_PATH = os.path.join(os.path.dirname(__file__), "..", "data", "summary.json")

MIN_VALID_DATE = datetime.datetime(2024, 1, 1)

# 읽을 컬럼만 명시적으로 화이트리스트 (PII 컬럼은 여기 없음 -> 애초에 접근 안 함)
COLS = {
    "인입일자": 2,
    "담당자": 3,
    "인입경로": 4,
    "초재진": 5,
    "문의시술명": 8,
    "예약일": 9,
    "예약시간": 10,
    "상담유형": 11,
}

BOOKED_STATUS = "예약완료"

WEEKDAY_KO = ["월", "화", "수", "목", "금", "토", "일"]


def clean_str(v):
    if v is None:
        return None
    s = str(v).strip()
    return s or None


def normalize_stage(raw):
    if raw is None:
        return None
    key = re.sub(r"\s+", "", str(raw).strip())
    if not key:
        return None
    return STAGE_MAP.get(key, "기타")


def follower_tier(v):
    if not isinstance(v, (int, float)):
        return "미상"
    if v < 10000:
        return "1만 미만"
    if v < 50000:
        return "1만~5만"
    if v < 100000:
        return "5만~10만"
    return "10만 이상"


def parse_follower_count(v):
    """'75.1만' 같은 한국어 단위 문자열 또는 이미 숫자인 값을 원 단위 숫자로 변환."""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    if not s:
        return None
    m = re.match(r"^([\d.]+)\s*만$", s)
    if m:
        try:
            return float(m.group(1)) * 10000
        except ValueError:
            return None
    try:
        return float(s)
    except ValueError:
        return None


def parse_roi_rate(v):
    """ROI달성률: 정상 값은 float(예: 0.85=85%, 5.5=550%). '#VALUE!'/'#DIV/0!' 등 에러 문자열은 제외."""
    if isinstance(v, (int, float)):
        return float(v)
    return None


def parse_numeric(v):
    if isinstance(v, (int, float)):
        return float(v)
    return None


def parse_upload_date(v):
    """'2025.09.27' / '26.06.29' / '2026. 6. 24' 등 다양한 표기의 업로드날짜 문자열을 date로 변환."""
    if v is None:
        return None
    s = str(v).strip()
    if not s:
        return None
    parts = [p.strip() for p in s.split(".")]
    if len(parts) != 3 or not all(parts):
        return None
    y, m, d = parts
    if len(y) == 2:
        y = "20" + y
    try:
        return datetime.date(int(y), int(m), int(d))
    except ValueError:
        return None


def parse_experience_date(v):
    """'2025.11.13(목)' 같은 요일 괄호 포함 문자열이나 datetime 값을 date로 변환."""
    if isinstance(v, datetime.datetime):
        return v.date()
    if v is None:
        return None
    s = re.sub(r"\(.*?\)", "", str(v)).strip()
    if not s:
        return None
    parts = [p.strip() for p in s.split(".") if p.strip()]
    if len(parts) != 3:
        return None
    y, m, d = parts
    if len(y) == 2:
        y = "20" + y
    try:
        return datetime.date(int(y), int(m), int(d))
    except ValueError:
        return None


def normalize_experience_status(v):
    """'전화번호'/'비고' 칸에 섞여 들어간 상태 마킹만 골라내고, 그 외 자유 텍스트(시술명·메모 등)는 버린다."""
    key = re.sub(r"\s+", "", str(v).strip()) if v is not None else ""
    return EXPERIENCE_STATUS_MAP.get(key)


def find_header_row(rows, marker):
    for i, r in enumerate(rows[:5]):
        if marker in r:
            return i
    return 0


def load_local_workbook_rows(path, sheet_names):
    """로컬 xlsx에서 지정된 시트들을 읽어 {시트명: 행리스트} 형태로 반환."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    return {name: list(wb[name].iter_rows(values_only=True)) for name in sheet_names}


def build_influencer_funnel(rows_by_sheet):
    """rows_by_sheet: {시트명: 행리스트}. 로컬 xlsx든 구글 시트 API든 이 형태로만 넘기면 됨."""
    counter = collections.Counter()
    n_total = 0
    n_used = 0

    for sheet_name in INFLUENCER_SHEETS:
        rows = rows_by_sheet[sheet_name]
        header = rows[0]
        col_idx = {}
        for i, v in enumerate(header):
            if v and v not in col_idx:
                col_idx[v] = i

        def get(r, name, col_idx=col_idx):
            idx = col_idx.get(name)
            return r[idx] if idx is not None and idx < len(r) else None

        for r in rows[1:]:
            if not any(v is not None for v in r):
                continue
            n_total += 1
            stage = normalize_stage(get(r, "진행상태"))
            if stage is None:
                continue
            n_used += 1
            country = clean_str(get(r, "국적")) or "미상"
            staff = clean_str(get(r, "루트")) or "미상"
            tier = follower_tier(get(r, "팔로우 수"))
            counter[(country, staff, tier, stage)] += 1

    influencer = [
        {"country": c, "staff": s, "followerTier": t, "stage": g, "count": cnt}
        for (c, s, t, g), cnt in sorted(counter.items())
    ]
    return influencer, n_total, n_used


def build_mega_roi(rows_by_sheet):
    """마케팅팀 인플루언서 시트의 '메가인플루언서 협업ROI' 탭을 집계.
    개인 식별 컬럼(차트번호/아이디/계정 링크/콘텐츠 업로드 링크)은 애초에 읽지 않는다.
    rows_by_sheet: {시트명: 행리스트}. 로컬 xlsx든 구글 시트 API든 이 형태로만 넘기면 됨.
    """
    rows = rows_by_sheet[MEGA_ROI_SHEET]
    header = rows[1]  # 0: 타이틀행('방문전'), 1: 헤더행
    col_idx = {}
    for i, v in enumerate(header):
        if v and v not in col_idx:
            col_idx[v] = i

    def get(r, name):
        idx = col_idx.get(name)
        return r[idx] if idx is not None and idx < len(r) else None

    n_total = 0
    n_used = 0
    acc = collections.defaultdict(lambda: {
        "count": 0,
        "estimateSum": 0.0, "estimateCount": 0,
        "cpvSum": 0.0, "cpvCount": 0,
        "roiSum": 0.0, "roiCount": 0,
    })

    for r in rows[2:]:
        if not any(v is not None for v in r):
            continue
        n_total += 1

        staff = clean_str(get(r, "담당자")) or "미상"
        platform = clean_str(get(r, "플랫폼 채널")) or "미상"
        nationality = clean_str(get(r, "국적")) or "미상"
        follower_num = parse_follower_count(get(r, "팔로워 수"))
        tier = follower_tier(follower_num)

        estimate = parse_numeric(get(r, "견적(단위:만원)"))
        cpv = parse_numeric(get(r, "CPV(예측)"))
        roi = parse_roi_rate(get(r, "ROI달성률"))

        n_used += 1
        key = (staff, platform, nationality, tier)
        bucket = acc[key]
        bucket["count"] += 1
        if estimate is not None:
            bucket["estimateSum"] += estimate
            bucket["estimateCount"] += 1
        if cpv is not None:
            bucket["cpvSum"] += cpv
            bucket["cpvCount"] += 1
        if roi is not None:
            bucket["roiSum"] += roi
            bucket["roiCount"] += 1

    mega_roi = [
        {
            "staff": staff, "platform": platform, "nationality": nationality,
            "followerTier": tier, "count": b["count"],
            "estimateSum": round(b["estimateSum"], 4), "estimateCount": b["estimateCount"],
            "cpvSum": round(b["cpvSum"], 6), "cpvCount": b["cpvCount"],
            "roiSum": round(b["roiSum"], 6), "roiCount": b["roiCount"],
        }
        for (staff, platform, nationality, tier), b in sorted(acc.items())
    ]
    return mega_roi, n_total, n_used


def build_experience_bookings(rows_by_sheet):
    """체험단 예약 시트(중국/대만/일본/영미)를 월/국가/상태 단위로 집계.
    성함/생년월일/전화번호/왓츠앱·카톡/차트번호/비고 원문은 절대 읽어 내보내지 않는다.
    '전화번호'·'비고' 칸에 상태 문구(방문 완/예약 완/취소 등)가 섞여 들어간 경우만
    normalize_experience_status()로 걸러 코드화하고, 그 외 자유 텍스트(시술명 메모 등)는 버린다.
    중국/대만 시트는 소스에 중복 행이 섞여 있어(대만 탭에 중국 데이터가 일부 복사돼 들어감)
    두 시트를 합쳐 완전-일치 행 기준으로 중복 제거한 뒤, 실제 '체험단국적' 값으로 재배정한다.
    """
    combined_cn_tw = list(dict.fromkeys(
        rows_by_sheet["중국 체험단 예약"] + rows_by_sheet["대만체험단 예약"]
    ))
    sheet_row_groups = [combined_cn_tw, rows_by_sheet["일본 체험단 예약"], rows_by_sheet["영미 체험단 예약"]]

    month_counter = collections.Counter()  # (month, country, status)
    slot_counter = collections.Counter()  # (weekday, hour)
    n_total = 0
    n_used = 0

    for rows in sheet_row_groups:
        header_idx = find_header_row(rows, "체험단국적")
        header = rows[header_idx]
        col_idx = {}
        for i, v in enumerate(header):
            if v and v not in col_idx:
                col_idx[str(v).strip()] = i

        def get(r, name, col_idx=col_idx):
            idx = col_idx.get(name)
            return r[idx] if idx is not None and idx < len(r) else None

        for r in rows[header_idx + 1:]:
            if not any(v is not None for v in r):
                continue
            n_total += 1
            country = clean_str(get(r, "체험단국적"))
            booking_date = parse_experience_date(get(r, "예약 날짜"))
            if country is None or booking_date is None:
                continue
            n_used += 1

            status = (
                normalize_experience_status(get(r, "전화번호"))
                or normalize_experience_status(get(r, "비고"))
                or normalize_experience_status(get(r, "왓으앱/카톡"))
                or "미표시"
            )
            month_key = booking_date.strftime("%Y-%m")
            month_counter[(month_key, country, status)] += 1

            booking_time = get(r, "예약 시간")
            if isinstance(booking_time, datetime.time):
                weekday = WEEKDAY_KO[booking_date.weekday()]
                slot_counter[(weekday, booking_time.hour)] += 1

    bookings = [
        {"month": m, "country": c, "status": s, "count": cnt}
        for (m, c, s), cnt in sorted(month_counter.items())
    ]
    slots = [
        {"weekday": w, "hour": h, "count": c}
        for (w, h), c in sorted(slot_counter.items(), key=lambda kv: (WEEKDAY_KO.index(kv[0][0]), kv[0][1]))
    ]
    return bookings, slots, n_total, n_used


def build_cs_bookings(rows):
    """예약문의통합 시트(rows[0]: 타이틀행, rows[1]: 헤더행)를 일자/담당자/시술/슬롯 단위로 집계.
    rows: 시트 전체 행리스트. 로컬 xlsx든 구글 시트 API든 이 형태로만 넘기면 됨.
    """
    daily_counter = collections.Counter()  # (date_iso, 채널, 초재진, 상담유형)
    staff_counter = collections.Counter()  # (month, 담당자, booked_bool)
    procedure_counter = collections.Counter()  # (month, 시술명)
    slot_counter = collections.Counter()  # (요일, 시간대)

    n_total = 0
    n_used = 0

    for r in rows[2:]:  # 0: 타이틀행, 1: 헤더행
        n_total += 1
        inquiry_date = r[COLS["인입일자"]]
        if not isinstance(inquiry_date, datetime.datetime):
            continue
        if inquiry_date < MIN_VALID_DATE:
            continue  # 이상치(예: 2020년 데이터 오입력) 제외

        channel = clean_str(r[COLS["인입경로"]])
        visit_type = clean_str(r[COLS["초재진"]])
        status = clean_str(r[COLS["상담유형"]])
        staff = clean_str(r[COLS["담당자"]])
        procedure = clean_str(r[COLS["문의시술명"]])

        n_used += 1
        date_iso = inquiry_date.date().isoformat()
        month_key = inquiry_date.strftime("%Y-%m")

        daily_counter[(date_iso, channel or "미상", visit_type or "미상", status or "미상")] += 1

        if staff:
            booked = status == BOOKED_STATUS
            staff_counter[(month_key, staff, booked)] += 1

        if procedure:
            procedure_counter[(month_key, procedure)] += 1

        booking_date = r[COLS["예약일"]]
        booking_time = r[COLS["예약시간"]]
        if isinstance(booking_date, datetime.datetime) and isinstance(
            booking_time, datetime.time
        ):
            weekday = WEEKDAY_KO[booking_date.weekday()]
            hour = booking_time.hour
            slot_counter[(weekday, hour)] += 1

    daily = [
        {"date": d, "channel": ch, "visitType": vt, "status": st, "count": c}
        for (d, ch, vt, st), c in sorted(daily_counter.items())
    ]
    staff = [
        {"month": m, "staff": s, "booked": b, "count": c}
        for (m, s, b), c in sorted(staff_counter.items())
    ]
    procedures = [
        {"month": m, "procedure": p, "count": c}
        for (m, p), c in sorted(procedure_counter.items())
    ]
    slots = [
        {"weekday": w, "hour": h, "count": c}
        for (w, h), c in sorted(slot_counter.items(), key=lambda kv: (WEEKDAY_KO.index(kv[0][0]), kv[0][1]))
    ]
    return daily, staff, procedures, slots, n_total, n_used


def build_content_performance(rows_by_sheet):
    """콘텐츠 업로드 관리 시트(전체통합)를 월/국가/플랫폼/상태 단위로 집계.
    Links(원본 게시물 링크), 인사이트취합일은 읽지 않는다.
    rows_by_sheet: {시트명: 행리스트}. 로컬 xlsx든 구글 시트 API든 이 형태로만 넘기면 됨.
    """
    rows = rows_by_sheet[CONTENT_SHEET]
    header = rows[0]
    col_idx = {}
    for i, v in enumerate(header):
        if v and v not in col_idx:
            col_idx[v] = i

    def get(r, name):
        idx = col_idx.get(name)
        return r[idx] if idx is not None and idx < len(r) else None

    n_total = 0
    n_used = 0
    upload_counter = collections.Counter()  # (month, country, platform, status)
    engagement_acc = collections.defaultdict(lambda: {
        "likesSum": 0.0, "likesCount": 0,
        "commentsSum": 0.0, "commentsCount": 0,
        "sharesSum": 0.0, "sharesCount": 0,
        "savesSum": 0.0, "savesCount": 0,
        "viewsSum": 0.0, "viewsCount": 0,
        "pairedLikesSum": 0.0, "pairedViewsSum": 0.0, "pairedCount": 0,
    })

    for r in rows[1:]:
        if not any(v is not None for v in r):
            continue
        n_total += 1

        country = clean_str(get(r, "국가"))
        upload_date = parse_upload_date(get(r, "업로드날짜"))
        if country is None or upload_date is None:
            continue  # 국가·날짜 둘 다 없는 빈 행(참조용 행) 제외

        n_used += 1
        platform = clean_str(get(r, "플랫폼")) or "미상"
        status = clean_str(get(r, "상태")) or "미상"
        month_key = upload_date.strftime("%Y-%m")

        upload_counter[(month_key, country, platform, status)] += 1

        likes = parse_numeric(get(r, "좋아요"))
        comments = parse_numeric(get(r, "댓글"))
        shares = parse_numeric(get(r, "공유(DM)"))
        saves = parse_numeric(get(r, "저장"))
        views = parse_numeric(get(r, "조회수"))

        eng = engagement_acc[(month_key, platform)]
        if likes is not None:
            eng["likesSum"] += likes
            eng["likesCount"] += 1
        if comments is not None:
            eng["commentsSum"] += comments
            eng["commentsCount"] += 1
        if shares is not None:
            eng["sharesSum"] += shares
            eng["sharesCount"] += 1
        if saves is not None:
            eng["savesSum"] += saves
            eng["savesCount"] += 1
        if views is not None:
            eng["viewsSum"] += views
            eng["viewsCount"] += 1
        if likes is not None and views is not None:
            eng["pairedLikesSum"] += likes
            eng["pairedViewsSum"] += views
            eng["pairedCount"] += 1

    uploads = [
        {"month": m, "country": c, "platform": p, "status": s, "count": cnt}
        for (m, c, p, s), cnt in sorted(upload_counter.items())
    ]
    engagement = [
        {"month": m, "platform": p, **{k: (round(v, 4) if isinstance(v, float) else v) for k, v in b.items()}}
        for (m, p), b in sorted(engagement_acc.items())
    ]
    return uploads, engagement, n_total, n_used


def build_viral_posting(rows_by_sheet):
    """바이럴 포스팅 막계정 시트(샤오홍슈/스레드)를 월/플랫폼 단위로 집계.
    포스팅 링크(원본 게시물 URL)는 애초에 읽지 않는다.
    '비고 (업로드 내용)' 칸은 자유 텍스트(캠페인 태그 등)라 원문을 절대 읽어 내보내지 않고,
    trim한 값이 정확히 '삭제'인지 여부만 골라내 (month, platform)별 삭제 건수로만 남긴다.
    rows_by_sheet: {시트명: 행리스트}. 로컬 xlsx든 구글 시트 API든 이 형태로만 넘기면 됨.
    """
    upload_counter = collections.Counter()  # (month, platform)
    deleted_counter = collections.Counter()  # (month, platform)
    engagement_acc = collections.defaultdict(lambda: {
        "likesSum": 0.0, "likesCount": 0,
        "commentsSum": 0.0, "commentsCount": 0,
        "viewsSum": 0.0, "viewsCount": 0,
        "pairedLikesSum": 0.0, "pairedViewsSum": 0.0, "pairedCount": 0,
    })

    n_total = 0
    n_used = 0

    for sheet_name in VIRAL_SHEETS:
        rows = rows_by_sheet[sheet_name]
        header_idx = find_header_row(rows, "업로드 날짜")
        header = rows[header_idx]
        col_idx = {}
        for i, v in enumerate(header):
            if v and v not in col_idx:
                col_idx[str(v).strip()] = i

        def get(r, name, col_idx=col_idx):
            idx = col_idx.get(name)
            return r[idx] if idx is not None and idx < len(r) else None

        for r in rows[header_idx + 1:]:
            if not any(v is not None for v in r):
                continue
            n_total += 1

            upload_date = get(r, "업로드 날짜")
            if not isinstance(upload_date, datetime.datetime):
                continue  # 이 시트의 업로드 날짜는 실제 datetime 값 (문자열 아님)
            n_used += 1

            platform = clean_str(get(r, "업로드 플랫폼")) or "미상"
            month_key = upload_date.strftime("%Y-%m")
            upload_counter[(month_key, platform)] += 1

            memo = get(r, "비고 (업로드 내용)")
            if memo is not None and str(memo).strip() == "삭제":
                deleted_counter[(month_key, platform)] += 1

            likes = parse_numeric(get(r, "좋아요 수"))
            comments = parse_numeric(get(r, "댓글 수"))
            views = parse_numeric(get(r, "조회수"))

            eng = engagement_acc[(month_key, platform)]
            if likes is not None:
                eng["likesSum"] += likes
                eng["likesCount"] += 1
            if comments is not None:
                eng["commentsSum"] += comments
                eng["commentsCount"] += 1
            if views is not None:
                eng["viewsSum"] += views
                eng["viewsCount"] += 1
            if likes is not None and views is not None:
                eng["pairedLikesSum"] += likes
                eng["pairedViewsSum"] += views
                eng["pairedCount"] += 1

    uploads = [
        {"month": m, "platform": p, "count": cnt, "deletedCount": deleted_counter.get((m, p), 0)}
        for (m, p), cnt in sorted(upload_counter.items())
    ]
    engagement = [
        {"month": m, "platform": p, **{k: (round(v, 4) if isinstance(v, float) else v) for k, v in b.items()}}
        for (m, p), b in sorted(engagement_acc.items())
    ]
    return uploads, engagement, n_total, n_used


def build():
    cs_rows_by_sheet = load_local_workbook_rows(SOURCE_PATH, [SHEET_NAME])
    daily, staff, procedures, slots, n_total, n_used = build_cs_bookings(
        cs_rows_by_sheet[SHEET_NAME]
    )

    influencer_rows_by_sheet = load_local_workbook_rows(
        INFLUENCER_PATH, INFLUENCER_SHEETS + [MEGA_ROI_SHEET]
    )
    influencer, inf_total, inf_used = build_influencer_funnel(influencer_rows_by_sheet)
    mega_roi, roi_total, roi_used = build_mega_roi(influencer_rows_by_sheet)

    content_rows_by_sheet = load_local_workbook_rows(CONTENT_PATH, [CONTENT_SHEET])
    content_uploads, content_engagement, content_total, content_used = build_content_performance(
        content_rows_by_sheet
    )

    experience_rows_by_sheet = load_local_workbook_rows(EXPERIENCE_PATH, EXPERIENCE_SHEETS)
    experience_bookings, experience_slots, exp_total, exp_used = build_experience_bookings(
        experience_rows_by_sheet
    )

    viral_rows_by_sheet = load_local_workbook_rows(VIRAL_PATH, VIRAL_SHEETS)
    viral_uploads, viral_engagement, viral_total, viral_used = build_viral_posting(
        viral_rows_by_sheet
    )

    output = {
        "generatedAt": datetime.datetime.now().isoformat(timespec="seconds"),
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

    # 안전장치: 출력에 PII 흔적이 없는지 확인 (이름/전화번호/차트번호/소셜 아이디/링크 패턴)
    blob = json.dumps(output, ensure_ascii=False)
    forbidden_keys = [
        "이름", "성함", "차트번호", "전화번호", "생년월일", "위챗 ID",
        "아이디", "인적사항", "비고", "콘텐츠 링크", "계정 링크",
        "콘텐츠 업로드 링크", "Links", "xhslink", "instagram.com",
        "tiktok.com", "xiaohongshu.com", "왓으앱", "카톡",
        "포스팅 링크", "threads.com", "xhslink.com",
    ]
    for k in forbidden_keys:
        assert k not in blob, f"PII 관련 키가 출력에 포함됨: {k}"

    os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)
    with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)

    print(f"원본 행수: {n_total}, 사용된 행수(유효 날짜): {n_used}")
    print(f"daily 집계행: {len(daily)}, staffMonthly: {len(staff)}, "
          f"procedureMonthly: {len(procedures)}, bookingSlots: {len(slots)}")
    print(f"인플루언서 원본 행수: {inf_total}, 사용된 행수: {inf_used}, "
          f"influencerFunnel 집계행: {len(influencer)}")
    print(f"메가인플루언서 ROI 원본 행수: {roi_total}, 사용된 행수: {roi_used}, "
          f"megaRoi 집계행: {len(mega_roi)}")
    print(f"콘텐츠 업로드 원본 행수: {content_total}, 사용된 행수: {content_used}, "
          f"contentUploads 집계행: {len(content_uploads)}, contentEngagement 집계행: {len(content_engagement)}")
    print(f"체험단 예약 원본 행수: {exp_total}, 사용된 행수: {exp_used}, "
          f"experienceBookings 집계행: {len(experience_bookings)}, experienceSlots 집계행: {len(experience_slots)}")
    print(f"바이럴 포스팅 원본 행수: {viral_total}, 사용된 행수(유효 날짜): {viral_used}, "
          f"viralPostingUploads 집계행: {len(viral_uploads)}, viralPostingEngagement 집계행: {len(viral_engagement)}")
    print(f"출력: {OUTPUT_PATH}")

    stage_totals = collections.Counter()
    for row in influencer:
        stage_totals[row["stage"]] += row["count"]
    print("인플루언서 단계별 합계:", dict(stage_totals))

    # 검증용: 상담유형 총합이 원본과 일치하는지 확인할 수 있도록 콘솔에 요약 출력
    status_totals = collections.Counter()
    for row in daily:
        status_totals[row["status"]] += row["count"]
    print("상담유형별 합계:", dict(status_totals))

    # 검증용: 콘텐츠 국가/플랫폼/상태별 합계 (원본 값과 대조)
    country_totals = collections.Counter()
    platform_totals = collections.Counter()
    status_totals2 = collections.Counter()
    for row in content_uploads:
        country_totals[row["country"]] += row["count"]
        platform_totals[row["platform"]] += row["count"]
        status_totals2[row["status"]] += row["count"]
    print("콘텐츠 국가별 합계:", dict(country_totals))
    print("콘텐츠 플랫폼별 합계:", dict(platform_totals))
    print("콘텐츠 상태별 합계:", dict(status_totals2))

    exp_country_totals = collections.Counter()
    exp_status_totals = collections.Counter()
    for row in experience_bookings:
        exp_country_totals[row["country"]] += row["count"]
        exp_status_totals[row["status"]] += row["count"]
    print("체험단 국가별 합계:", dict(exp_country_totals))
    print("체험단 상태별 합계:", dict(exp_status_totals))

    # 검증용: 바이럴 포스팅 플랫폼별/삭제 건수 합계
    viral_platform_totals = collections.Counter()
    viral_deleted_totals = collections.Counter()
    for row in viral_uploads:
        viral_platform_totals[row["platform"]] += row["count"]
        viral_deleted_totals[row["platform"]] += row["deletedCount"]
    print("바이럴 포스팅 플랫폼별 합계:", dict(viral_platform_totals))
    print("바이럴 포스팅 플랫폼별 삭제 건수:", dict(viral_deleted_totals))


if __name__ == "__main__":
    build()
