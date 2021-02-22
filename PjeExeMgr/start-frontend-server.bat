@echo off

TITLE FRONTEND-SERVER

call config.cmd

cd %EXEPJE_PATH%\exe-frontend

ng serve --host 0.0.0.0 --disable-host-check

pause