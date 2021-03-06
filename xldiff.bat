set path1=%1
set path2=%2
REM Create paths to original and current spreadsheets to store in tmp
REM Change forward slash to back slash on all paths for the Excel tool
set path1=%path1:/=\%
set path2=%path2:/=\%
if not "%path1:xls=%" == "%path1%" (
 REM Clean up paths and write to tmp file as spreadsheetcompare requires input to be passed inside of a file.
 dir %path1% /B /S > tmp.txt
 dir %path2% /B /S >> tmp.txt
 "C:\Program Files (x86)\Microsoft Office\root\Office16\DCF\spreadsheetcompare" tmp.txt
) else if not "%path1:smp=%" == "%path1%" (
 if not "%~f1" == "%~dp0before.SMP" (
  copy /B "%~f1" "%~dp0\before.txt"
 )
 copy /B "%~f2" "%~dp0\after.txt"
 find <%~dp0\before.txt /V "hogehogehoge" >%~dp0\before.smp
 find <%~dp0\after.txt /V "hogehogehoge" >%~dp0\after.smp
 Start "" %~dp0\before.smp
 Start "" %~dp0\after.smp
) else (
 "C:\Program Files\WinMerge\WinMergeU.exe" %1 %2
)