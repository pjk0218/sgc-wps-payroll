# SGC Arabia · WPS Payroll Files

HR 급여 엑셀을 회계팀 은행 양식(BSF-WPS Conversion Tool, .xls + Export 매크로)으로 가공하는 웹앱.
서버 없이 브라우저에서만 동작하며, 급여 자료는 브라우저와 팀 공유 폴더 밖으로 나가지 않는다.

## 규칙
- 금액(급여액·기본급·주택·기타·공제): HR 파일 값. 빈칸은 0.
- 이름·은행코드·IBAN: 마스터 값. 마스터에 없으면 HR 값으로 만들고 알림.
- Payment Description: `Salary for <Month> <Year>` 로 통일.
- Value Date: 집행일 다음 날. 목요일이면 다음 주 일요일.
- 항상 별도 파일로 빼는 사람(D360 은행 이용자, PDH trainee 등)은 공유 폴더 `settings.json`에서 관리 (코드에 사번을 넣지 않음).
- 양식의 매크로·버튼·서식·수식은 그대로 두고 셀 값만 교체. 만든 파일은 다시 읽어 자체 검증.

## 파일
- `index.html`, `app.js` — 화면, 저장(공유 폴더 / 브라우저)
- `wps.js` — 가공 규칙과 점검
- `xlsbin.js` — .xls(BIFF8) 읽기·부분 수정, .xlsx 읽기 (외부 라이브러리 없음)
- `tests/` — 10월 정답지 대조(regress), 인원 증감(cross), 화면 E2E(e2e). 급여 자료가 있어야 실행됨.

## 공유 폴더 구조
```
templates/<출력>.xls     은행 양식 (회계팀 최종본)
master.json              마스터
settings.json            별도 파일 대상자 (사번 → 출력)
2026-11/HR/              HR 원본
2026-11/Accounting/      만든 은행 파일
2026-11/run.json         기록 (인원·합계·알림, IBAN 제외)
```
