# 공용 설정 JSON 사용 방식

이 앱은 기본적으로 실행 위치의 `OneNote_Remocon_Setting.json`을 사용한다.

프로젝트 실행과 EXE 실행이 같은 설정을 쓰게 하려면 앱 메뉴에서 `파일 > 공용 설정 JSON 위치 지정...`을 사용한다.

## 권장 구조

실제 설정 JSON은 OneDrive 동기화 폴더에 둔다.

```text
OneDrive\앱설정\OneNote_Remocon_Setting.json
```

앱은 사용자 PC의 포인터 파일에 이 경로를 저장한다.

```text
%APPDATA%\OneNote_Remocon\OneNote_Remocon_Setting.path
```

이 포인터는 같은 PC 안에서 프로젝트 실행과 EXE 실행이 같이 읽는다.

## 다른 PC에서의 동작

다른 PC에서도 가능하다. 단, OneDrive의 실제 로컬 경로가 PC마다 다를 수 있으므로 각 PC에서 한 번씩 `공용 설정 JSON 위치 지정...`으로 해당 PC의 OneDrive JSON을 선택한다.

그 뒤에는 같은 OneDrive JSON이 동기화되므로 프로젝트 실행과 EXE 실행이 같은 내용을 사용한다.

## 우선순위

설정 JSON 경로는 아래 순서로 결정된다.

1. `ONENOTE_REMOCON_SETTINGS_PATH` 환경변수
2. `%APPDATA%\OneNote_Remocon\OneNote_Remocon_Setting.path`
3. 실행 위치의 `OneNote_Remocon_Setting.path`
4. 기본 위치의 `OneNote_Remocon_Setting.json`

## 주의

OneDrive 동기화 충돌을 피하려면 여러 PC에서 동시에 앱을 켜고 같은 JSON을 수정하지 않는 것이 좋다.
