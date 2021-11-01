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
 START "" "C:\Program Files (x86)\tsuzuki\Sma4Win\sma4win.exe" %1
 START "" "C:\Program Files (x86)\tsuzuki\Sma4Win\sma4win.exe" %2 
) else (
 "C:\Program Files\WinMerge\WinMergeU.exe" %1 %2
)