@echo off
echo @echo off > cmd001.bat
echo fc /b %1.txt %2.txt >> cmd001.bat
REM ***********************************************
REM コマンドの先頭の @ は、そのコマンドの表示を抑制
REM echo off は以降のコマンドの表示を抑制
REM ***********************************************

REM ***********************************************
REM サブルーチンの呼び出し
REM ***********************************************
call :GetFC
if errorlevel 1 ( Call :CheckOk ) else ( Call :CheckErr )

REM ***********************************************
REM 処理の終わり
REM ***********************************************
goto end



REM ***********************************************
REM サブルーチン
REM ***********************************************
:GetFC
for /F "delims=: tokens=2" %%i in ('call cmd001.bat') do (
    if "%%i"==" 相違点は検出されませんでした" (
        exit /B 1
    )
)
exit /B 0

:CheckOk
echo ファイルは一致しました
echo 作業を続けて下さい
exit /B

:CheckErr
echo ファイルは一致しませんでした
echo 作業を中止して下さい
exit /B


REM ***********************************************
REM 記述の終わり
REM ***********************************************
:end