# 한글·Word PDF 변환 스킬 예제

AI 대화에서 지정한 한글·Word 문서를 PDF로 일괄 변환하는 Windows용 스킬 예제입니다. AI는 요청 범위와 결과 보고를 맡고, 포함된 PowerShell 스크립트가 설치된 한글·Word의 PDF 저장 기능을 사용합니다.

> 초기 예제입니다. 실제 한글·Word 변환 검증은 아직 완료되지 않았습니다. 먼저 대표 문서 한 개로 PDF 생성과 내용·서식을 확인하세요.

## 필요한 환경

- Windows
- HWP·HWPX 변환: 한컴오피스 한글
- DOC·DOCX 변환: Microsoft Word
- 대상 PC의 파일과 프로그램 실행 기능을 사용할 수 있는 AI 환경

규칙 파일만 첨부하면 PC 접근 권한이 생기는 것은 아닙니다. 상시 폴더 감시와 예약 실행은 포함하지 않습니다.

## 파일 구조

```text
skills/pdf-batch-convert/
├── SKILL.md                    # 요청 해석과 변환·보호·보고 규칙
├── agents/openai.yaml          # 스킬 표시 정보
└── scripts/Convert-ToPDF.ps1    # 실제 변환 스크립트
```

## 스킬로 사용하기

1. 저장소를 내려받습니다.
2. `skills/pdf-batch-convert` 폴더 전체를 자신의 Codex 스킬 폴더로 복사합니다. 기본 위치는 사용자 홈의 `.codex/skills`이며 `CODEX_HOME`을 따로 지정했다면 그 아래 `skills`를 사용합니다.
3. 스킬을 인식하는 새 대화에서 아래처럼 요청합니다. 경로는 실제 위치로 바꾸세요.

```text
$pdf-batch-convert C:\업무\제출서류 폴더의 한글·Word 파일을 PDF로 변환해줘. 하위 폴더는 제외해줘.
```

## 직접 실행하기

Windows PowerShell에서 저장소 폴더를 기준으로 실행합니다.

```powershell
powershell.exe -NoProfile -STA -File .\skills\pdf-batch-convert\scripts\Convert-ToPDF.ps1
```

파일 선택창에서 여러 문서를 선택할 수 있습니다. 실행 정책으로 차단되면 조직·PC의 권한 절차를 따르세요.

AI가 파일 목록을 전달하는 방법은 [SKILL.md](skills/pdf-batch-convert/SKILL.md)에 있습니다.

## 기본 규칙

- 지원 형식: `.hwp`, `.hwpx`, `.doc`, `.docx`
- 원본을 수정하거나 삭제하지 않습니다.
- 원본 폴더 안의 `PDF 결과`에 `계약서.hwp.pdf`처럼 저장합니다.
- 같은 이름의 PDF가 있으면 건너뜁니다. 원본 변경 여부를 비교하는 기능은 없습니다.
- 상세 CSV 기록은 스크립트 옆의 `변환 기록`에 저장합니다.
- 외부 변환 서비스로 문서를 업로드하지 않습니다.

## 한계와 검증 상태

스킬 형식과 기존 결과 건너뛰기·원본 보존 동작을 확인했습니다. 실제 변환 검증은 완료하지 못했습니다. PDF 시작 표시 `%PDF-` 확인은 서식·내용의 정확성을 보장하지 않습니다.

한글 접근 확인창, 암호, 복구 안내 등에 사용자 응답이 필요할 수 있습니다. 파일별 자동 시간 제한은 없으며 프로그램이 대기하면 작업도 대기합니다. 여러 작업을 동시에 실행하지 마세요.

개인 문서와 생성된 PDF·로그는 이 예제에 포함하지 않습니다.
