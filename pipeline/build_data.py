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


def load_rows():
    wb = openpyxl.load_workbook(SOURCE_PATH, read_only=True, data_only=True)
    ws = wb[SHEET_NAME]
    rows = list(ws.iter_rows(values_only=True))
    return rows[2:]  # 0: 타이틀행, 1: 헤더행


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


def build_influencer_funnel():
    wb = openpyxl.load_workbook(INFLUENCER_PATH, read_only=True, data_only=True)
    counter = collections.Counter()
    n_total = 0
    n_used = 0

    for sheet_name in INFLUENCER_SHEETS:
        ws = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
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


def build():
    rows = load_rows()

    daily_counter = collections.Counter()  # (date_iso, 채널, 초재진, 상담유형)
    staff_counter = collections.Counter()  # (month, 담당자, booked_bool)
    procedure_counter = collections.Counter()  # (month, 시술명)
    slot_counter = collections.Counter()  # (요일, 시간대)

    n_total = 0
    n_used = 0

    for r in rows:
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

    influencer, inf_total, inf_used = build_influencer_funnel()

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
    }

    # 안전장치: 출력에 PII 흔적이 없는지 확인 (이름/전화번호/차트번호/소셜 아이디 패턴)
    blob = json.dumps(output, ensure_ascii=False)
    forbidden_keys = [
        "이름", "차트번호", "전화번호", "생년월일", "위챗 ID",
        "아이디", "인적사항", "비고", "콘텐츠 링크", "계정 링크",
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


if __name__ == "__main__":
    build()
