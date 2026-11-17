# 국내 터미널 재고표 취합 — Claude 클라우드 Routine 파일럿

기존 Windows 작업 스케줄러(`KMTC-StockReport-Daily`)로 로컬 실행하던 재고표 취합을
Claude 클라우드 스케줄 에이전트(routine)로 이전하기 위한 패키지.

## 아키텍처

- **에이전트(클라우드)**: Gmail 커넥터로 9개 터미널 메일 검색 → 첨부 다운로드 → 저장 → 취합본 발송
- **Python(`merge_stock_cloud.py`)**: 다운로드된 파일들을 하나의 엑셀로 취합 (win32com 無, 리눅스 호환)
- 사내 Oracle DB 미사용 → 클라우드 이전 가능한 유일 후보

### 로컬 버전 대비 변경점
| 항목 | 로컬 | 클라우드 |
|------|------|----------|
| Gmail | OAuth 로컬 토큰 | **Gmail 커넥터(MCP)** |
| HJNC 수식 | Excel COM(win32com) | **openpyxl data_only(캐시값)** |
| .xls 처리 | xlrd | **xlrd==1.2.0 핀** |
| Google Drive 업로드 | 있음 | **제거**(필요 시 Drive 커넥터로 별도 추가) |

## 선행 조건

1. **Gmail 커넥터 연결**: https://claude.ai/customize/connectors
2. 이 repo를 **private GitHub repo**로 push
3. routine 생성 시 Gmail 커넥터 attach + 이 repo를 source로 지정

## 수신자 / 제목

- 수신: `eqc@ekmtc.com`, `sbjung@ekmtc.com`
- 제목: `{YYYYMMDD}_국내 터미널 재고표 취합`

## Routine 프롬프트 (자기완결형)

아래 프롬프트를 routine의 `events[].message.content`에 넣는다.

```
당신은 '국내 터미널 재고표 취합' 자동화 에이전트다. 오늘(Asia/Seoul) 기준으로 아래를 수행하라.

[0] 날짜: 오늘 날짜를 KST로 구한다. 토/일이면 즉시 종료(평일만 수행).
    변수: YYYYMMDD, YYMMDD, MMDD, "YYYY/MM/DD", "MM월 DD일", 전일/익일(YYYY/MM/DD).

[1] 작업폴더 생성: ./work  (다운로드 첨부 저장용)

[2] Gmail 커넥터로 아래 9개 터미널 메일을 검색하고, 규칙에 맞는 첨부 1개를
    ./work/{터미널}.xls 또는 .xlsx (원본 확장자 유지)로 저장한다.
    메일이 없거나 적합 첨부가 없으면 그 터미널은 '미수신'으로 기록하고 건너뛴다.

  1) HBGT  q: subject:({YYMMDD}_KMTC_HBGT OR {YYYYMMDD}_KMTC_HBGT) after:{전일} before:{익일}
           pick: 확장자 제외 파일명이 "{YYMMDD}_KMTC_HBGT"와 정확히 일치. 없으면
                 식별자 "kmtc_hbgt"로 끝나는 파일이 유일할 때만 사용(_RF/_SOC 등 배제)
  2) BPTS  q: subject:{YYMMDD}_KMTC_BPTS            pick: 유일한 첨부
  3) BPTG  q: subject:{YYMMDD}_KMTC_BPTG            pick: 유일한 엑셀 첨부
  4) HJIT  q: subject:{YYMMDD}_KMTC_HJIT            pick: 파일명에 "Stock" 포함 엑셀
  5) ICT   q: subject:{YYMMDD} KMTC_ICT             pick: 파일명에 {YYYYMMDD} 포함 엑셀
  6) SNCT  q: subject:SNCT subject:"KMD Daily Stock Report" after:{전일} before:{익일}
           pick: 유일한 첨부
  7) HJNC  q: subject:"HJNC {MMDD} KMD Empty status"  pick: 유일한 첨부
  8) PNCT  q: subject:"KMTC_DPW Busan Daily Stock Report" after:{전일} before:{익일}
           pick: 유일한 첨부(가장 최근 1통)
  9) DGT   q: subject:"Dongwon Global Terminal Stock Report" subject:"{MM월 DD일}"
           pick: 유일한 첨부

  * 한 터미널에 여러 통이 오면 기준일 수신분을 쓴다.
  * 첨부는 반드시 바이너리 원본 그대로 저장한다(엑셀 손상 방지).

[3] 취합 실행:
    pip install -r requirements.txt
    python merge_stock_cloud.py {YYYYMMDD} ./work "./work/{YYYYMMDD}_국내 터미널 재고표 취합.xlsx"
    - stdout에 "FAILED:AAA,BBB" 줄이 있으면 그 터미널은 '취합실패(파일손상)'로 기록.

[4] 메일 발송 (Gmail 커넥터):
    수신: eqc@ekmtc.com, sbjung@ekmtc.com
    제목: {YYYYMMDD}_국내 터미널 재고표 취합
    첨부: 위 취합 xlsx (발송 전 openpyxl로 열어 무결성 확인, 1KB 미만이면 중단)
    본문:
      {YYYYMMDD} 기준 국내 터미널 재고표 취합 파일을 송부드립니다.

      [취합 대상 터미널 - 총 9개]
      1. PUS/HBGT
      2. PUS/BPTS
      3. PUS/BPTG
      4. INC/HJIT : https://www.hjit.co.kr/kor/Main.do
      5. INC/ICT : https://service.psa-ict.co.kr/
      6. INC/SNCT : https://snct.sun-kwang.co.kr/infoservice/webpage/main/mainPage.jsp
      7. PNC/HJNC : https://www.hjnc.co.kr/esvc
      8. PNC/PNCT : https://svc.pncport.com/info/Main.do
      9. PNC/DGT : https://info.dgtbusan.com/DGT/esvc

      감사합니다.
    - 미수신 터미널은 해당 줄 끝에 " (취합파일없음)" 표시.
    - 취합실패(파일손상) 터미널은 " (취합실패-파일손상)" 표시 + 본문 하단에 안내문.

[5] 완료 보고: 다운로드/미수신/실패 터미널 목록과 발송 결과를 요약 출력.
```

## 참고 / 리스크

- claude.ai Gmail 커넥터의 **바이너리 첨부 다운로드·첨부 발송 지원 여부**는 실제 연결 후
  `RemoteTrigger action:run`으로 1회 테스트하여 확정해야 한다(이 파일럿 최대 리스크).
- HJNC Report는 발신처가 저장한 캐시값에 의존 → 값이 비면 발신처 파일에 캐시가 없는 것.
- cron 최소 주기 1시간. 평일 09:00 KST = 00:00 UTC → `0 0 * * 1-5`.
