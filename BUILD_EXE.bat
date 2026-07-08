@echo off
chcp 949 >nul
title EXE 빌드 - 교육결과보고서_자동생성기_v3
setlocal enabledelayedexpansion
pushd "%~dp0"

echo ============================================================
echo   [교육결과보고서_자동생성기_v3] 실행파일(EXE) 만들기
echo   이 창이 끝나면 dist 폴더에 exe가 생성됩니다.
echo   (파이썬이 설치된 PC에서 한 번만 실행하면 됩니다)
echo ============================================================
echo.

rem --- 파이썬 자동 감지 (py 런처 우선, 없으면 python) ---
set "PYCMD="
py -3 --version >nul 2>&1 && set "PYCMD=py -3"
if not defined PYCMD (
  python --version >nul 2>&1 && set "PYCMD=python"
)
if not defined PYCMD (
  echo [오류] 파이썬을 찾을 수 없습니다.
  echo        https://www.python.org/downloads/ 에서 설치하세요.
  echo        설치 첫 화면에서 "Add python.exe to PATH" 를 반드시 체크하세요.
  echo.
  pause
  popd
  exit /b 1
)
echo [확인] 파이썬 명령: %PYCMD%
%PYCMD% --version
echo.

echo [1/3] 필요한 패키지를 설치합니다...
%PYCMD% -m pip install --upgrade pip
%PYCMD% -m pip install --upgrade pyinstaller python-pptx openpyxl tkinterdnd2 lxml
if errorlevel 1 goto ERR

echo.
echo [2/3] EXE 를 만듭니다... (수 분 정도 걸릴 수 있습니다)
%PYCMD% -m PyInstaller --onefile --noconsole --clean --noconfirm ^
  --name "교육결과보고서_자동생성기_v3" ^
  --add-data "assets;assets" ^
  --collect-all tkinterdnd2 ^
  --collect-all pptx ^
  --collect-all openpyxl ^
  --hidden-import lxml._elementpath ^
  "report_generator.py"
if errorlevel 1 goto ERR

echo.
echo [3/3] 설정/템플릿 파일을 dist 폴더로 복사합니다...
if exist "설정.xlsx"            copy /Y "설정.xlsx" "dist\" >nul
if exist "템플릿_결보표양.pptx"  copy /Y "템플릿_결보표양.pptx" "dist\" >nul
if exist "빠른시작.txt"          copy /Y "빠른시작.txt" "dist\" >nul
if exist "사용설명서.md"         copy /Y "사용설명서.md" "dist\" >nul
if exist "새기능_직접입력.txt"    copy /Y "새기능_직접입력.txt" "dist\" >nul

echo.
echo ============================================================
echo   완료! 아래 dist 폴더 안의 내용을 통째로 배포하세요.
echo     - 교육결과보고서_자동생성기_v3.exe   (프로그램 본체)
echo     - 설정.xlsx                 (교육마다 이 값만 수정)
echo     - 템플릿_결보표양.pptx        (서식, 그대로 두기)
echo ============================================================
echo.
if exist "dist" start "" "dist"
pause
popd
exit /b 0

:ERR
echo.
echo [오류] 빌드 중 문제가 발생했습니다.
echo   1) 인터넷 연결을 확인하세요(패키지 다운로드 필요).
echo   2) 백신이 pyinstaller 를 차단하면 잠시 예외로 등록하세요.
echo   3) 위에 표시된 마지막 오류 메시지를 그대로 담당자에게 전달하세요.
echo.
pause
popd
exit /b 1
