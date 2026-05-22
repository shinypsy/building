import sys

dial_content = """- 생성 시간: 2026-05-21T09:11:23+09:00
- 요청: git 보류
- 답변: 네, Jay. Git 업로드는 보류하겠습니다. 다음 지시를 기다리겠습니다.
"""

with open(r'd:\Dev\Project\sample\doc\dial_260521.md', 'a', encoding='utf-8') as f:
    f.write(dial_content)
