# [엔지니어 가이드] 개발 기초 환경 구축 입문서

안녕하세요! 이제 막 개발에 입문하시는 분들이 가장 효율적이고 안정적으로 개발 환경을 구축할 수 있도록 이 가이드를 작성했습니다.

특히 진행할 프로젝트는 **React, Python, PostgreSQL**이라는 현대적이고 AI 활용도가 매우 높은 기술 스택을 사용합니다. 아래 순서대로 차근차근 따라오시면 최상의 개발 준비를 마치실 수 있습니다.

---

## 1. 개발 프로그램 대분류 및 다운로드 링크

가장 먼저 필요한 도구들을 종류별로 정리했습니다. 표의 링크를 통해 설치 파일을 미리 다운로드해 주세요.

| 구분 | 프로그램명 | 용도 | 다운로드 링크 | 비고 |
| :--- | :--- | :--- | :--- | :--- |
| **AI Agent** | **Antigravity** | **개발 환경 구축 및 코딩 가이드** | [다운로드](https://edgedl.me.gvt1.com/edgedl/release2/j0qc3/antigravity/stable/1.23.2-4781536860569600/windows-x64/Antigravity.exe)  | **가장 먼저 설치 권장** |
| **Editor** | **Visual Studio Code** | 코드 작성, 수정, 디버깅 | [다운로드](https://code.visualstudio.com/) | AI 코딩 도구의 표준 |
| **Frontend** | **Node.js (LTS)** | React, Vite 실행 환경 | [다운로드](https://nodejs.org/en) | 20.x(LTS) 버전 권장 |
| **Backend** | **Python 3.11+** | FastAPI 서버 실행 | [다운로드](https://www.python.org/downloads/) | 최신 안정 버전 추천 |
| **Database** | **PostgreSQL 16** | 맛집 데이터 저장소(DB) | [다운로드](https://www.postgresql.org/download/windows/) | 서버용 DB 엔진 |
| **DB Tools** | **DBeaver** | DB 관리용 GUI 도구 | [다운로드](https://dbeaver.io/download/) | 무료 오픈소스 도구 |
| **VCS** | **Git** | 코드 버전 및 협업 관리 | [다운로드](https://git-scm.com/download/win) | 필수 기본 도구 |

---

## 2. 권장 설치 순서 (Step-by-Step)

프로그램 간의 의존성과 편의성을 고려하여 다음 순서로 설치하는 것을 강력히 권장합니다.

### Step 0: AI 가이드 준비 (Antigravity)
*   **Antigravity**를 가장 먼저 설치하세요. 설치 후 "환경 구축 시작할게, 도와줘"라고 말하면, 이후 모든 단계에서 발생하는 문제나 궁금증을 실시간으로 해결해 드립니다.

### Step 1: 기본 도구 (Editor & Git)
1. **Visual Studio Code**: 설치 시 'Code로 열기' 메뉴를 탐색기에 추가하는 옵션을 체크하세요.
2. **Git**: 기본 옵션으로 설치하되, 기본 에디터를 VS Code로 선택하면 편리합니다.

### Step 2: 실행 환경 (Node.js & Python)
1. **Node.js**: 설치 완료 후 터미널에서 `node -v` 명령어로 확인이 필요합니다.
2. **Python**: **중요!** 설치 시작 화면에서 **[Add Python to PATH]** 체크박스를 반드시 선택해야 합니다. 그렇지 않으면 터미널에서 실행되지 않습니다.

### Step 3: 데이터베이스 (PostgreSQL & DBeaver)
1. **PostgreSQL**: 설치 과정에서 설정하는 `postgres` 사용자의 비밀번호를 반드시 메모해 두세요. (본 프로젝트 연결 시 필요)
2. **DBeaver**: 설치 후 PostgreSQL에 접속하여 데이터가 잘 들어가는지 확인하는 용도로 사용합니다.

---

## 3. 설치 시 주의사항 및 엔지니어 팁

### 📂 작업 폴더(Workspace) 설정
*   **경로에 공백 금지**: `C:\Users\홍 길동\Project` 처럼 사용자 이름에 공백이 있으면 일부 도구가 오작동할 수 있습니다.
*   **권장 경로**: `D:\Dev\Project\aaa` 또는 `C:\Dev` 처럼 짧고 공백이 없는 영문 경로를 사용하세요.

### 🐍 파이썬 가상 환경 (Virtual Environment)
*   백엔드 개발 시에는 반드시 가상 환경(`.venv`)을 만들어 사용해야 합니다. 이는 프로젝트별로 필요한 도구들이 서로 엉키지 않게 보호해 주는 역할을 합니다.

### 🤖 AI 코딩 에이전트 Antigravity 활용법
*   본 프로젝트 개발에는 강력한 AI 에이전트인 **Antigravity**가 함께합니다. 단순히 코드를 짜주는 것을 넘어, 복잡한 프로젝트 구조를 분석하고 안전한 구현 계획을 제안합니다.
*   **리서치 요청**: "이 기능이 어떻게 동작하는지 분석해줘"라고 요청하면 `research.md`에 상세 내용을 정리해 줍니다.
*   **계획 수립**: 코드를 작성하기 전 반드시 `plan.md`를 통해 구현 방향을 먼저 제시하므로, 초보자도 실수를 줄이고 전문적인 개발 방식을 배울 수 있습니다.
*   **질문하기**: 개발 중 막히는 부분이 생기면 언제든 "이 에러 왜 발생하는 거야?"라고 물어보세요.

---

## 4. 최종 설치 확인 (체크리스트)

설치가 완료되었다면, 터미널(PowerShell 또는 CMD)을 열고 아래 명령어들을 입력하여 버전을 확인해 보세요.

```bash
# 1. 노드 버전 확인
node -v

# 2. 파이썬 버전 확인
python --version

# 3. 깃 버전 확인
git -v
```

모든 명령어에서 버전 정보가 출력된다면, 이제 개발을 시작할 준비가 모두 끝났습니다! 환영합니다! 🚀
