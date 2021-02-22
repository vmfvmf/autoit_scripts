@echo off
pje-exed1-a1
TITLE BACKEND-SERVER

call config.cmd

set "MODULE_OPTS=-jaxpmodule javax.xml.jaxp-provider"
cd %WILDF%
call standalone.bat
pause