# Codex OneNote COM Playbook

이 문서는 OneNote 작업을 실제로 실행할 때 참고하는 내부 운영 노트다. 사용자 스킬이나 사용자 요청문에 붙이지 않는다.

## 읽는 순서

1. `onenote-com-internal.md`에서 공통 원칙과 사용자 스킬 경계를 확인한다.
2. 작업 종류에 맞는 `원노트-*.md` 파일을 연다.
3. 대상 ID가 필요하면 `docs/codex/onenote-targets.json`과 `docs/codex/onenote-location-cache.json`을 먼저 확인한다.
4. 캐시가 없거나 실패할 때만 OneNote 계층을 다시 조회한다.

## 실행 원칙

- 화면 클릭 자동화보다 Windows OneNote COM API를 우선한다.
- PowerShell 연결은 `New-Object -ComObject OneNote.Application`을 사용한다.
- 기본 계층 조회는 `hsSections`까지로 제한한다.
- `hsPages`는 페이지 목록, 페이지 복제, 페이지 수 검증처럼 필요한 작업에서만 사용한다.
- 이름만으로 대상을 고정하지 말고 상위 경로와 ID를 함께 확인한다.
- `OpenHierarchy`나 `CreateNewPage` 직후에는 OneNote 동기화 지연이 있을 수 있으므로 재시도와 짧은 대기를 둔다.
- XML 수정은 문자열 치환보다 XML 파서와 namespace manager를 사용한다.

## 대상 캐시

- 고정 대상 목록: `docs/codex/onenote-targets.json`
- OneNote 조회 결과 캐시: `docs/codex/onenote-location-cache.json`

캐시에 있는 값은 실행 속도를 높이기 위한 힌트다. ID 호출이 실패하면 계층 조회로 다시 확인하고, 새로 확인한 ID는 캐시에 반영한다.

## 작업 라우팅

| 작업 | 내부 문서 | 검증 |
| --- | --- | --- |
| 페이지 추가 | `원노트-페이지-추가.md` | `GetPageContent(pageId)` |
| 섹션 생성 | `원노트-섹션-생성.md` | `GetHierarchy(sectionGroupId, hsSections, ref xml)` |
| 섹션 그룹 생성 | `원노트-섹션그룹-생성.md` | `GetHierarchy(parentId, hsSections, ref xml)` |
| 전자필기장 생성/열기 | `원노트-전자필기장-생성.md` | `GetHierarchy('', hsNotebooks, ref xml)` 또는 `hsSections` |
| 전자필기장 복제 | `원노트-전자필기장-복제.md` | 활성 섹션 그룹/섹션/페이지 수 비교 |
| 대상 ID 찾기 | `원노트-대상ID-찾기.md` | 캐시 저장값과 계층 조회 결과 비교 |

## 보고 기준

- 성공 보고에는 작업한 OneNote 항목, 대상 이름, 검증 결과만 남긴다.
- 실패 보고에는 실패한 단계, 대상 경로 또는 ID, 다시 확인해야 할 값만 남긴다.
- 내부 COM 호출 순서, PowerShell 템플릿 전문, 대상 캐시 원문은 사용자가 요청한 경우에만 보여준다.
