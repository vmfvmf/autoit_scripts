@echo off

CALL config.cmd

SET var="%~1"

if "%~1"=="" SET var="pa"
if %var%=="pa" (
TITLE BUILD PJE-ASSINATURA
cd %EXEPJE_PATH%\pje-assinatura
call %MAVN% clean install -DskipTests
)

if "%~1"=="" SET var="pb"
if %var%=="pb" (
TITLE BUILD PJE-BACKEND
cd %EXEPJE_PATH%\pje-backend
call %MAVN% clean install -DskipTests
)

if "%~1"=="" SET var="pcm"
if %var%=="pcm" (
TITLE BUILD PJE-CENTRALMANDADOS
cd %EXEPJE_PATH%\pje-centralmandados
call %MAVN% clean install -DskipTests
)

if "%~1"=="" SET var="pcmf"
if %var%=="pcmf" (
TITLE BUILD PJE-CENTRALMANDADOS-FRONTEND
cd %EXEPJE_PATH%\pje-centralmandados-frontend
REM npm i
)

if "%~1"=="" SET var="pco"
if %var%=="pco" (
TITLE BUILD PJE-CORPORATIVO
cd %EXEPJE_PATH%\pje-corporativo
call %MAVN% clean install -DskipTests
)

if "%~1"=="" SET var="et"
if %var%=="et" (
TITLE BUILD PJE-ETIQUETAS
cd %EXEPJE_PATH%\pje-etiquetas-backend
call %MAVN% clean install -DskipTests
)

if "%~1"=="" SET var="pi"
if %var%=="pi" (
TITLE BUILD PJE-INTEGRACAO
cd %EXEPJE_PATH%\pje-integracao
call %MAVN% clean install -DskipTests
)

if "%~1"=="" SET var="ps"
if %var%=="ps" (
TITLE BUILD PJE-SEGURANCA
cd %EXEPJE_PATH%\pje-seguranca
call %MAVN% clean install -DskipTests
)

if "%~1"=="" SET var="eb"
if %var%=="eb" (
TITLE BUILD EXE-BACKEND
cd %EXEPJE_PATH%\exe-backend
call %MAVN% clean install -DskipTests
)

if "%~1"=="" SET var="ef"
if %var%=="ef" (
TITLE BUILD EXE-FRONTEND
cd %EXEPJE_PATH%\exe-frontend
npm i
)

pause
exit