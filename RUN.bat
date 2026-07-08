@echo off
chcp 949 >nul
title 교육결과보고서_자동생성기_v3 (파이썬으로 바로 실행)
pushd "%~dp0"

set "PYCMD="
py -3 --version >nul 2>&1 && set "PYCMD=py -3"
if not defined PYCMD (
  python --version >nul 2>&1 && set "PYCMD=python"
)
if not defined PYCMD (
  echo [오류] 파이썬이 설치되어 있지 않습니다.
  echo        https://www.python.org/downloads/ 에서 설치하세요.
  echo        ("Add python.exe to PATH" 체크 필수)
  pause
  popd
  exit /b 1
)

echo 프로그램을 준비합니다. 처음 실행 시 필요한 모듈을 설치할 수 있습니다...
%PYCMD% -c "import pptx, openpyxl" >nul 2>&1 || %PYCMD% -m pip install python-pptx openpyxl
%PYCMD% -c "import tkinterdnd2" >nul 2>&1 || %PYCMD% -m pip install tkinterdnd2

%PYCMD% "report_generator.py"
popd
