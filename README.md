# 중화권 CS 예약 대시보드

`중화권 예약 &DB리스트.xlsx`의 `예약문의통합` 시트를 날짜×채널×유형 단위로
집계해, 개인정보(이름/차트번호/전화번호/생년월일 등) 없이 통계 패턴만
보여주는 정적 대시보드입니다.

## 구조

- `pipeline/build_data.py` — 원본 xlsx(로컬의 `~/Desktop/Dashboard_data`)를
  읽어 `data/summary.json`에 집계 결과만 출력합니다. 원본 파일은 이 저장소에
  절대 포함되지 않습니다 (`.gitignore` 참고).
- `index.html`, `styles.css`, `app.js` — `data/summary.json`만 읽어 렌더링하는
  순수 정적 사이트. 빌드 과정 없이 GitHub Pages에 그대로 배포 가능합니다.

## 데이터 갱신

```bash
python3 pipeline/build_data.py
```

## 로컬 확인

```bash
python3 -m http.server 8123
# http://localhost:8123 접속
```
