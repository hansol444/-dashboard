# Topic: 회의록 정리

## 개요
녹취록 TXT 파일을 Claude로 구조화 요약하고, Notion에 등록하는 업무.
업무 지시/방향성 변화를 자동 추출해 대시보드 연동까지 수행.

## 실행 방법
```bash
cd meeting-notes
python summarize.py input/회의록.txt           # 요약만
python summarize.py input/회의록.txt --notion   # Notion에도 등록
```

## 출력물
- 구조화된 요약 (안건별, 결정사항, 액션 아이템)
- 업무 지시 추출 → `output/pending_actions.json` (대시보드 연동)
- 방향성 변화 추출 → Team Context 업데이트 재료

## 프로세스
1. 녹취록 TXT 파일을 `meeting-notes/input/`에 넣기
2. `summarize.py` 실행
3. 생성된 요약 검토 (output/ 폴더)
4. `--notion` 옵션으로 Notion Meeting Notes DB에 등록
5. `pending_actions.json`의 업무 지시 항목 대시보드 반영 확인

## 주의사항
- 녹취록 파일은 TXT 형식 (STT 결과물)
- 발화자 구분이 명확할수록 요약 품질 높아짐
- Notion 등록 시 페이지 제목 = 회의 날짜 + 참석자

## 담당 연락처
- 창준님 — 회의록 원본 및 등록 확인
