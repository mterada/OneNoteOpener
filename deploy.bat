
set TARG=OneNoteOpener.exe


cd bin\Release
call essign %TARG%

copy %TARG% C:\tools
echo "copied to C:\tools\."

pause
