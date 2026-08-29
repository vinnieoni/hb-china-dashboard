"""
원본 CS 상담 로그(개인정보 포함)를 읽어, 개인 식별이 불가능한 수준까지
집계(aggregate)한 결과만 data/summary.json 으로 출력한다.

절대 하지 않는 것: 원본 행 단위 레코드 출력, 이름/차트번호/전화번호/생년월일 등
PII 컬럼을 읽거나 저장하는 것.
"""
import json
import datetime
import collections
import os

import openpyxl

SOURCE_PATH = os.path.expanduser(
    "~/Desktop/Dashboard_data/중화권 예약 &DB리스트.xlsx"
)
SHEET_NAME = "예약문의통합"
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

    output = {
        "generatedAt": datetime.datetime.now().isoformat(timespec="seconds"),
        "sourceRowCount": n_total,
        "usedRowCount": n_used,
        "daily": daily,
        "staffMonthly": staff,
        "procedureMonthly": procedures,
        "bookingSlots": slots,
    }

    # 안전장치: 출력에 PII 흔적이 없는지 확인 (이름/전화번호/차트번호 패턴)
    blob = json.dumps(output, ensure_ascii=False)
    forbidden_keys = ["이름", "차트번호", "전화번호", "생년월일", "위챗 ID"]
    for k in forbidden_keys:
        assert k not in blob, f"PII 관련 키가 출력에 포함됨: {k}"

    os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)
    with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)

    print(f"원본 행수: {n_total}, 사용된 행수(유효 날짜): {n_used}")
    print(f"daily 집계행: {len(daily)}, staffMonthly: {len(staff)}, "
          f"procedureMonthly: {len(procedures)}, bookingSlots: {len(slots)}")
    print(f"출력: {OUTPUT_PATH}")

    # 검증용: 상담유형 총합이 원본과 일치하는지 확인할 수 있도록 콘솔에 요약 출력
    status_totals = collections.Counter()
    for row in daily:
        status_totals[row["status"]] += row["count"]
    print("상담유형별 합계:", dict(status_totals))


if __name__ == "__main__":
    build()
