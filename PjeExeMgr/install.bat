@echo off

CALL config.cmd

SET var="%~1"

if not exist %EXEPJE_PATH%\exe-pje mkdir %EXEPJE_PATH%\exe-pje

if "%~1"=="" SET var="ef"
if %var%=="ef" (
title INSTALL exe-frontend
rmdir %EXEPJE_PATH%\exe-frontend /S /Q
cd %EXEPJE_PATH%
git clone https://git.pje.csjt.jus.br/pje-satelites/exe/exe-frontend.git
)

if "%~1"=="" SET var="eb"
if %var%=="eb" (
title INSTALL exe-backend
rmdir %EXEPJE_PATH%\exe-backend /S /Q
cd %EXEPJE_PATH%
git clone https://git.pje.csjt.jus.br/pje-satelites/exe/exe-backend.git
)

REM if "%~1"=="" SET var="ps"
REM if %var%=="ps" (
REM title UPDATE pje-seguranca
REM cd %EXEPJE_PATH%
REM rmdir %EXEPJE_PATH%\pje-seguranca /S /Q
REM git clone https://git.pje.csjt.jus.br/pje2/pje-seguranca.git
REM )


if "%~1"=="" SET var="pb"
if %var%=="pb" (
title INSTALL pje-backend
cd %EXEPJE_PATH%
rmdir %EXEPJE_PATH%\pje-backend /S /Q
git clone https://git.pje.csjt.jus.br/pje2/pje-backend.git
)

if "%~1"=="" SET var="pi"
if %var%=="pi" (
title INSTALL pje-integracao
rmdir %EXEPJE_PATH%\pje-integracao /S /Q
cd %EXEPJE_PATH%\pje24\integracao
git clone https://git.pje.csjt.jus.br/pje2/pje-integracao.git
)

pause
