# Architecture

> 제품 범위는 [prd.md](./prd.md), 고객 맥락은 [customer.md](./customer.md), 운영 원칙은 [AGENTS.md](./AGENTS.md)를 참조한다.

## 문서 목적
이 문서는 Data Audit Agent가 무엇을 입력받고, 어떻게 처리하고, 무엇을 출력해야 하는지 정리한다.

## 1. 시스템 목표
검증 대상 폴더를 입력받아 보고서 초안과 참고 자료를 비교하고, 검토형 Excel 패키지를 만든다.

## 2. 입력
- 검증 폴더 경로
- 보고서 초안 PDF/DOCX
- 참고 자료 PDF/XLSX/DOCX
- 선택 입력: 중요 수치, 검증 범위, 사용자 메모
- 샘플 기준: `report-draft.md`, `r-one-extract.md`

## 3. 처리 흐름
1. 폴더를 스캔해 파일 목록을 만든다.
2. 파일을 보고서 후보, 참고 자료, 보류로 분류한다.
3. 보고서와 참고 자료를 파싱한다.
4. 검증 블록을 추출한다.
5. 중요 수치를 표시한다.
6. 출처 후보를 찾는다.
7. 값, 기간, 기준, 단위를 비교한다.
8. `일치 / 불일치 / 확인 필요`를 판정한다.
9. 근거와 재현 메모를 기록한다.
10. Excel 감사 패키지를 생성한다.
11. `expected-results.md` 수준의 결과 요약을 남긴다.

## 4. 핵심 컴포넌트
- Folder Intake: 폴더와 파일 메타데이터 수집
- File Role Classifier: 보고서/참고자료/보류 분류
- Report Parser: 페이지, 표, 문단, 그림 구조 추출
- Source Parser: 출처 값과 위치 추출
- Verification Block Extractor: 검증 대상 숫자 블록 추출
- Priority Marker: 중요 수치 표시
- Source Matcher: 블록과 출처 연결
- Verification Engine: 판정 생성
- Evidence Recorder: 근거 저장
- Output Builder: Excel 결과물 생성

## 5. 데이터 계약
- `file_inventory`: 파일명, 경로, 형식, 역할 후보
- `verification_block`: 블록 유형, 순서, 페이지, 값, 단위, 기간, 중요 수치 여부
- `source_evidence`: 출처 파일명, 위치, 값, 단위, 기간
- `verification_result`: 판정, 차이 설명, 근거, 사람 검토 필요 여부

## 5-1. 샘플 출력
- `report-draft.md`: 보고서 초안
- `r-one-extract.md`: 출처 발췌
- `expected-results.md`: 기대 결과 요약

## 6. 사람 승인 지점
- 검증 폴더 선택
- 중요 수치 확인
- 역할이 애매한 파일 분류
- `확인 필요` 항목 판단
- `불일치` 반영 여부 결정
- 최종 승인

## 7. 실패 / 안전장치
- 출처를 못 찾으면 `확인 필요`로 남긴다.
- OCR이 불안정하면 보수적으로 판정한다.
- 복수 출처 후보는 임의 확정하지 않는다.
- 자동 수정은 제안 수준으로만 둔다.
- 모든 결과는 출처 파일과 위치를 따라가게 남긴다.
