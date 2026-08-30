"""임시 진단 스크립트: 5개 스프레드시트의 실제 탭 이름/gid를 나열한다.
서비스 계정이 아직 공유되지 않은 시트는 403으로 표시된다. 확인 후 삭제 예정."""
import os

from google.oauth2 import service_account
from googleapiclient.discovery import build as build_service

SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
KEY_PATH = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS")

SHEETS = {
    "인플루언서(연결됨)": "1yEWyVVPW5cYQkjJ-EAlx37veDfZFLiCQG3nILkcgmH8",
    "CS예약": "1OH1krRpZ6dDnAqoJFWrdMInOXJPa2gSuePbm_dsPEZk",
    "콘텐츠업로드": "1N2ls94wtqTzRJdPprDb-7KgGXTIDOz6W9aAuoZ5PETI",
    "바이럴포스팅": "1kNf-4A90R2QJ8vjh-_UNQ3G5rcXrLeNqO2BKli4fhCA",
    "체험단정리": "14nyN_DPfNYSM2oCLBszuCWywLXcP_X_eSebNl3w4hlQ",
}


def main():
    creds = service_account.Credentials.from_service_account_file(KEY_PATH, scopes=SCOPES)
    service = build_service("sheets", "v4", credentials=creds)
    for label, sid in SHEETS.items():
        print(f"\n=== {label} ({sid}) ===")
        try:
            meta = service.spreadsheets().get(spreadsheetId=sid).execute()
            print("제목:", meta["properties"]["title"])
            for s in meta["sheets"]:
                p = s["properties"]
                print(f"  gid={p['sheetId']}  이름={p['title']!r}")
        except Exception as e:
            print("에러:", e)


if __name__ == "__main__":
    main()
