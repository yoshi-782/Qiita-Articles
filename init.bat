@echo off
setlocal

set NAME=%1

if "%NAME%"=="" (
  echo 記事名を指定してください
  echo 例: templates comfyui-inpaint
  exit /b 1
)

set DEST=public\%NAME%

if "%DEST%"=="" (
  echo DESTが指定されていません
  exit /b 1
)

if exist "%DEST%" (
  echo 既に存在します: %DEST%
  exit /b 1
)

mkdir "%DEST%"
cd "%DEST%"

(
    echo ---
    echo title: %NAME%
    echo tags:
    echo - ''
    echo private: false
    echo updated_at: ''
    echo id: null
    echo organization_url_name: null
    echo slide: false
    echo ignorePublish: false
    echo ---
    echo # new article body
    echo 
) > article.md

mkdir images
mkdir src

echo.
echo 作成完了:
echo   %DEST%
echo.
endlocal
