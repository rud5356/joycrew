# Joycrew Follow Checker

쪼이크루 팔로우/맞팔 상태를 확인하는 정적 웹 도구입니다. 인스타그램에서 내려받은 `followers_1.json`, `following.json` 파일을 브라우저에 업로드하면 쪼이크루 목록, 운영진, 언팔리스트를 기준으로 확인할 항목을 정리해 줍니다.

배포 주소:

```text
https://rud5356.github.io/joycrew/
```

## 주요 기능

- 인스타그램 팔로워/팔로잉 JSON 파일 업로드
- 쪼이크루 본계정과 부계정 맞팔 상태 확인
- 운영진 예외 처리
- 언팔리스트 전체 확인 및 언팔 필요 여부 표시
- 모든 분석을 브라우저 안에서만 처리

## 사용 방법

1. 인스타그램에서 내 정보 다운로드를 요청합니다.
2. 다운로드 형식은 `JSON`으로 선택합니다.
3. 압축을 푼 뒤 아래 파일을 준비합니다.

```text
connections/followers_and_following/followers_1.json
connections/followers_and_following/following.json
```

4. 홈페이지에서 두 파일을 업로드합니다.
5. `분석하기`를 눌러 결과를 확인합니다.

## joycrew_list.csv 형식

`joycrew_list.csv`는 헤더 없이 4개 열을 사용합니다.

| 열 | 의미 | 체크 규칙 |
|---:|---|---|
| 1열 | 조이크루 본계정 | 맞팔 체크 대상 |
| 2열 | 부계정 | 맞팔 체크 대상 |
| 3열 | 운영진 | 내가 팔로우하면 정상, 운영진이 나를 맞팔하지 않아도 정상 |
| 4열 | 언팔리스트 | 전체 목록 표시, 내가 팔로우 중이면 언팔필수 |

예시:

```csv
main_account,sub_account,staff_account,unfollow_account
```

계정명은 `@` 없이 쓰는 것을 권장합니다. 빈칸은 있어도 됩니다.

## 결과 상태

- `맞팔`: 서로 팔로우 중
- `맞팔 안 함`: 내가 팔로우 중이지만 상대가 나를 팔로우하지 않음
- `팔로우 안 함`: 상대가 나를 팔로우하지만 내가 팔로우하지 않음
- `팔로우 필요`: 운영진인데 내가 팔로우하지 않음
- `운영진 OK`: 운영진을 내가 팔로우 중
- `언팔필수`: 언팔리스트 계정을 내가 팔로우 중
- `언팔 완료`: 언팔리스트 계정을 내가 팔로우하지 않음

## 로컬에서 확인하기

`fetch()`로 `joycrew_list.csv`를 읽기 때문에 HTML 파일을 직접 더블클릭해서 여는 것보다 로컬 서버로 확인하는 것이 안정적입니다.

```powershell
cd C:\repos\joycrew
python -m http.server 8000
```

브라우저에서 아래 주소를 엽니다.

```text
http://127.0.0.1:8000/
```

## 파일 구조

```text
index.html             GitHub Pages 기본 진입 파일
joycrew_checker.html   체크기 HTML
joycrew_list.csv       체크 기준 목록
joycrew.csv            기존 1열 목록 파일
```

## 배포

GitHub Pages 설정은 다음 기준을 사용합니다.

```text
Source: Deploy from a branch
Branch: master
Folder: / (root)
```

`master` 브랜치에 push하면 GitHub Pages가 자동으로 갱신됩니다. 반영에는 몇 분 정도 걸릴 수 있습니다.

## 개인정보 주의

`followers_1.json`, `following.json`은 개인 인스타그램 데이터입니다. 저장소에 커밋하지 말고 로컬에서만 사용하세요.
