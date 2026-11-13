# interx-deck

INTERX 회사 양식 PPT 를 만드는 Claude Code 플러그인.

## 설치

```sh
npx --allow-git=all muramasaDev/interx-deck-skill
```

설치가 끝나면 Claude Code 를 다시 시작한다.

- Python, 패키지, 글꼴이 없으면 함께 설치된다
- 업데이트: 같은 명령
- 제거: 명령 끝에 `uninstall`
- 저장소 읽기 권한이 필요하다
- Windows 와 Mac. 렌더 확인과 PDF 는 PowerPoint (Mac 은 PowerPoint 가 없으면 LibreOffice, 줄 바꿈이 조금 다를 수 있음)
- Mac 은 명령의 `python` 을 `python3` 로

## 사용

Claude Code 에서 말로 요청한다. 덱마다 작성자, 파일명, 버전을 확인한다.

| 명령 | 용도 | 요청 예 |
|---|---|---|
| `/interx-deck:report` | 보고 자료, 고객 보고, 주간보고 | "○○ 결과 보고 덱 만들어 줘" |
| `/interx-deck:design` | 설계 문서, 구성도 | "수집 흐름 구성도 그려 줘" |
| `/interx-deck:okr` | 월간 Individual OKR | "이번 달 OKR 써 줘" |
| `/interx-deck:using` | 무엇을 쓸지 모를 때 | |

- 결과: 작업 폴더(또는 말한 폴더)에 pptx 와 PDF 가 같은 이름으로 생긴다
  - 이름: `날짜_파일명_작성자_버전`
- 고칠 때: "3장 표를 아래로" 처럼 말로 요청하거나 PowerPoint 에서 직접 고친다
- OKR 은 지난달 덱 파일을 함께 준다

## 주의

- 회사 DRM 때문에 PowerPoint 로 한 번 연 파일은 Claude 가 다시 읽지 못한다
- 보낼 때는 PDF 를 쓴다

## 라이선스

INTERX 사내 전용. 번들 글꼴 Noto Sans KR 은 SIL OFL 1.1.
