REM SET D=S:
SET D=\\nymain-fs01\abs-files
SET P=INT
SET F=bin\x86\Release

SET PROJECTS=ABSC ABSCS ABSX AP AR AS BM CC DP ED FA GL IC PO RS SA SO SP TA WH WHC TAC
REM PAUSE
XCOPY ABS\%F%\*.EXE bin /Y /D /C /I
XCOPY ABS\%F%\*.ISL bin /Y /D /C /I
DEL Temp\*.* /q
REM PAUSE
for %%i in (%PROJECTS%) DO (
DEL %%i\Datasets\*.* /q
XCOPY %%i\%F%\*.DLL bin /Y /D /C /I
)
ECHO NOW COPYING BINARIES TO %D%\%P%\bin
PAUSE
XCOPY bin %D%\%P%\bin /Y /E /C /D /I >> deploy.log
REM XCOPY Reports %D%\%P%\Reports /Y /E /C /D /I >> deploy.log
REM XCOPY Images %D%\%P%\Images /Y /E /C /D /I >> deploy.log
PAUSE

