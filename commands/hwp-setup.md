---
description: HWP Studio 사용 환경을 진단합니다. Python, pyhwpx, 한글 프로그램의 설치 상태를 확인하고 문제가 있으면 해결 방법을 안내합니다.
allowed-tools:
  - "mcp__hwp-studio__hwp_check_setup"
---

# HWP 환경 진단

hwp_check_setup 도구를 호출하여 환경을 진단하세요.

결과에 따라 사용자에게 안내하세요:

## 모두 정상 + 한글 실행 중
"환경이 준비되었습니다! `/hwp-help`로 사용 가능한 기능을 확인하세요."

## Python 미설치
"Python이 설치되어 있지 않습니다."
→ https://www.python.org/downloads/ 에서 설치
→ 설치 시 "Add Python to PATH" 반드시 체크
→ Microsoft Store 버전이 아닌 python.org 공식 버전 권장
→ 설치 후 터미널을 새로 열고 다시 `/hwp-setup` 실행

## Microsoft Store Python 감지
"Microsoft Store Python이 감지되었습니다. pyhwpx가 정상 동작하지 않을 수 있습니다."
→ python.org 공식 버전으로 재설치를 권장
→ 또는 PYTHON_PATH 환경변수를 설정하세요

## pyhwpx 미설치
"pyhwpx가 설치되어 있지 않습니다."
→ 터미널에서 실행: `pip install pyhwpx pywin32`
→ 설치 후 다시 `/hwp-setup` 실행

## 한글 미설치
"한글(HWP) 프로그램이 설치되어 있지 않습니다."
→ 한컴오피스 한글 2014 이상 설치 필요
→ 설치 후 한글을 한번 실행하여 초기 설정 완료

## 한글 미실행
"한글(HWP) 프로그램을 실행하고 빈 문서를 열어주세요."
→ 문서 작업 전 한글이 실행 중이어야 합니다
