"""
GitHub Actions에서 주기적으로 실행되는 진입점.
구글 시트 API로 인플루언서/메가ROI 6개 탭만 읽어 재집계하고,
data/summary.json의 해당 키만 갱신한다 (나머지 섹션은 그대로 유지).

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
    INFLUENCER_SHEETS,
    MEGA_ROI_SHEET,
    OUTPUT_PATH,
    build_influencer_funnel,
    build_mega_roi,
)

SPREADSHEET_ID = "1yEWyVVPW5cYQkjJ-EAlx37veDfZFLiCQG3nILkcgmH8"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
KEY_PATH = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS")

# 로컬 xlsx 스냅샷과 라이브 시트의 탭 이름이 다른 경우만 매핑 (라이브 시트 메타데이터로 확인됨,
# 2026-08-30). 매핑에 없는 이름은 그대로 사용.
LIVE_SHEET_NAME = {
    "중국 인풀루언서체험단": "중국 인풀루언서/체험단",
    "대만 인풀루언서체험단": "대만 인풀루언서/체험단",
    "일본 인풀루언서체험단": "일본 인풀루언서/체험단",
    "영미 인풀루언서체험단": "영미 인풀루언서/체험단",
}


def fetch_rows_by_sheet():
    creds = service_account.Credentials.from_service_account_file(KEY_PATH, scopes=SCOPES)
    service = build_service("sheets", "v4", credentials=creds)
    values_api = service.spreadsheets().values()

    rows_by_sheet = {}
    for local_name in INFLUENCER_SHEETS + [MEGA_ROI_SHEET]:
        live_name = LIVE_SHEET_NAME.get(local_name, local_name)
        result = values_api.get(
            spreadsheetId=SPREADSHEET_ID,
            range=f"'{live_name}'!A:AZ",
            valueRenderOption="UNFORMATTED_VALUE",
        ).execute()
        rows_by_sheet[local_name] = result.get("values", [])
    return rows_by_sheet


def sync():
    if not KEY_PATH:
        raise SystemExit("GOOGLE_APPLICATION_CREDENTIALS 환경변수가 설정되어 있지 않음")

    rows_by_sheet = fetch_rows_by_sheet()
    influencer, inf_total, inf_used = build_influencer_funnel(rows_by_sheet)
    mega_roi, roi_total, roi_used = build_mega_roi(rows_by_sheet)

    with open(OUTPUT_PATH, encoding="utf-8") as f:
        output = json.load(f)

    unchanged = (
        output.get("influencerSourceRowCount") == inf_total
        and output.get("influencerUsedRowCount") == inf_used
        and output.get("influencerFunnel") == influencer
        and output.get("megaRoiSourceRowCount") == roi_total
        and output.get("megaRoiUsedRowCount") == roi_used
        and output.get("megaRoi") == mega_roi
    )
    if unchanged:
        print("변경 없음 - summary.json을 다시 쓰지 않음 (generatedAt도 그대로 유지)")
        return

    output["influencerSourceRowCount"] = inf_total
    output["influencerUsedRowCount"] = inf_used
    output["influencerFunnel"] = influencer
    output["megaRoiSourceRowCount"] = roi_total
    output["megaRoiUsedRowCount"] = roi_used
    output["megaRoi"] = mega_roi
    output["generatedAt"] = datetime.datetime.now().isoformat(timespec="seconds")

    # 안전장치: build_data.py의 build()와 동일한 PII 검사를 병합된 전체 blob에 다시 적용
    blob = json.dumps(output, ensure_ascii=False)
    forbidden_keys = [
        "이름", "차트번호", "전화번호", "생년월일", "위챗 ID",
        "아이디", "인적사항", "비고", "콘텐츠 링크", "계정 링크",
        "콘텐츠 업로드 링크", "Links", "xhslink", "instagram.com",
        "tiktok.com", "xiaohongshu.com",
    ]
    for k in forbidden_keys:
        assert k not in blob, f"PII 관련 키가 출력에 포함됨: {k}"

    with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)

    print(f"인플루언서: 원본 {inf_total}행, 사용 {inf_used}행, 집계 {len(influencer)}행")
    print(f"메가ROI: 원본 {roi_total}행, 사용 {roi_used}행, 집계 {len(mega_roi)}행")
    print(f"갱신됨: {OUTPUT_PATH}")


if __name__ == "__main__":
    sync()
