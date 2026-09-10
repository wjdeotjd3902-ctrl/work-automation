---
name: outlook-shortpath-attachments
description: Classic Outlook 첨부 저장이나 ZIP 해제 시 긴 경로 문제가 발생하면 짧은 경로에 저장하고 원본 정보를 매니페스트로 보존한다.
---

# 첨부파일 저장

## 역할

대상 메일의 첨부파일을 짧은 경로 규칙에 따라 저장하여 파일과 원본 대응 기록을 생성한다.

## 입력값

| 항목 | 내용 | 필수 여부 |
|---|---|---|
| 메일 위치 | 메일함과 Outlook 폴더 경로 | 필수 |
| 조회 기간 | 시작일과 종료일, 종료일 포함 | 필수 |
| 저장 위치 | 짧은 로컬 출력 루트 | 필수 |
| 범주 필터 | 접수·제출 등, 미지정 시 제한 없음 | 선택 |
| 추가 처리 | MSG·인라인 이미지 포함 여부, ZIP 내부 확인 여부 | 선택 |

## 수행 방법

### 1. 저장 계획 확인

입력값은 기존 요청에서 확인하고 추론할 수 없는 필수 항목만 질문한다. `C:\OA\project` 또는 얕은 작업 폴더를 사용하며 Outlook 폴더 트리를 재현하지 않는다.

`scripts/export_outlook_attachments.ps1`을 `-Apply` 없이 실행하여 대상 메일·첨부·제외·실패 건수를 확인한다. 예상 범위와 다르거나 읽기에 실패하면 원인을 해결한다.

### 2. 첨부 저장 및 ZIP 해제

실제 저장이 승인된 범위에서 같은 명령에 `-Apply`를 추가한다. 저장 규칙은 다음과 같다.

| 대상 | 규칙 |
|---|---|
| 메일 폴더 | `yyyyMMdd_HHmmss_<EntryID hash>` |
| 첨부파일 | `a001_<original-name hash>.<ext>` |
| 경로 길이 | 실제 출력 루트 120자 미만, 저장 경로 240자 미만 |
| 원본 대응 기록 | 제목, EntryID, 원래 파일명, 수신 시각, 저장 경로, 크기, 결과를 매니페스트에 기록 |

ZIP 내부가 필요하면 `scripts/extract_archives_short.py <output-root>`의 계획을 검토한 후 `--apply`로 실행한다. 번호 기반 경로로 해제하고 원래 압축파일·멤버 이름은 `archive_manifest.json`에 기록한다.

### 3. 결과 검증

저장·건너뜀·실패 건수, 매니페스트, 파일 존재 및 대표 파일 시그니처를 확인하고 실제 출력 경로를 보고한다.

## 출력값

| 항목 | 형식 | 내용 |
|---|---|---|
| 첨부파일 | 원본 확장자 | 짧은 경로에 저장된 파일 |
| 원본 대응 기록 | JSON·CSV | `manifest.json`, `manifest.csv` |
| ZIP 대응 기록 | JSON | 해제 시 `archive_manifest.json` |
| 처리 결과 | 요약 | 저장·건너뜀·실패 건수와 출력 경로 |

## 주의사항

- 기본은 Dry Run이며 메일·범주·Flag·폴더를 변경하지 않는다.
- MSG는 기본 제외한다. 이미지는 Outlook 속성상 숨김 또는 인라인일 때만 제외하며 문서·ZIP에 이 기준을 적용하지 않는다.
- 재실행 시 예상 크기가 같은 기존 파일은 건너뛰고 무관한 파일을 삭제·덮어쓰지 않는다.
- 읽기 실패를 대상 0건으로 보고하지 않는다.
- 같은 원인으로 세 번 실패하면 중단하고 의심되는 전제를 명시한다.

## 예시

### 입력값

스킬 폴더에서 실행하는 Dry Run 예시다.

```powershell
pwsh -File scripts/export_outlook_attachments.ps1 -Mailbox 'user@example.com' -FolderPath '받은 편지함\프로젝트\도서관리' -StartDate '2026-09-01' -EndDate '2026-09-10' -Categories '접수' -Destination 'C:\OA\Project'
```

승인된 실제 저장에는 `-Apply`를 추가한다. `-IncludeMsg`, `-IncludeInlineImages`는 요청된 경우만 사용한다.

### 출력값

```text
<출력 루트>/
  <메일별 짧은 폴더>/<짧은 첨부파일명>
  manifest.json
  manifest.csv
```
