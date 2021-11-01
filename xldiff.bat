@ECHO OFF
set keyword=xls
ECHO %1 | find %keyword% >NUL
IF %ERRORLEVEL% ==1 (
 REM Create paths to original and current spreadsheets to store in tmp
 set path1=%1
 set path2=%2
 REM Change forward slash to back slash on all paths for the Excel tool
 set path1=%path1:/=\%
 set path2=%path2:/=\%
 REM Clean up paths and write to tmp file as spreadsheetcompare requires input to be passed inside of a file.
 dir %path1% /B /S > tmp.txt
 dir %path2% /B /S >> tmp.txt

 "C:\Program Files (x86)\Microsoft Office\root\Office16\DCF\spreadsheetcompare" tmp.txt
) ELSE (
 "C:\Program Files\WinMerge\WinMergeU.exe" %1 %2
)