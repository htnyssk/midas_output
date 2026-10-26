@echo off
rem 配布用 exe を作成します（PyInstaller 使用）。結果は dist\midas_output フォルダ。
chcp 65001 >nul
cd /d "%~dp0"
python -m pip install -r requirements.txt pyinstaller || goto :error
python -m PyInstaller --noconfirm --clean --onedir --console --name midas_output app.py || goto :error

rem 設定ファイルは exe と同じ場所に置く（利用者が編集できるよう exe には含めない）
copy /y config.yaml dist\midas_output\config.yaml >nul
if not exist dist\midas_output\images mkdir dist\midas_output\images
echo.
echo 完了: dist\midas_output\midas_output.exe
echo 配布するときは dist\midas_output フォルダごと渡してください。
pause
exit /b 0

:error
echo ビルドに失敗しました。
pause
exit /b 1
