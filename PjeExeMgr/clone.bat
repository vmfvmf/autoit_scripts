@echo off

CALL config.cmd

SET var="%~1"

if not exist %EXEPJE_PATH% mkdir %EXEPJE_PATH%

if "%~1"=="" SET var="pa"
if %var%=="pa" (
TITLE CLONE PJE-ASSINATURA
if exist %EXEPJE_PATH%\pje-assinatura rmdir %EXEPJE_PATH%\pje-assinatura /S /Q
cd %EXEPJE_PATH%
git clone https://git.pje.csjt.jus.br/pje2/pje-assinatura 
)

if "%~1"=="" SET var="pb"
if %var%=="pb" (
TITLE CLONE PJE-BACKEND
cd %EXEPJE_PATH%\pje-backend
if exist %EXEPJE_PATH%\pje-backend rmdir %EXEPJE_PATH%\pje-backend /S /Q
cd %EXEPJE_PATH%
git clone https://git.pje.csjt.jus.br/pje2/pje-backend
)

if "%~1"=="" SET var="pcm"
if %var%=="pcm" (
TITLE CLONE PJE-CENTRALMANDADOS
if exist %EXEPJE_PATH%\pje-centralmandados rmdir %EXEPJE_PATH%\pje-centralmandados /S /Q
cd %EXEPJE_PATH%
git clone https://git.pje.csjt.jus.br/pje2/pje-centralmandados
)

if "%~1"=="" SET var="pcmf"
if %var%=="pcmf" (
TITLE CLONE PJE-CENTRALMANDADOS-FRONTEND
if exist %EXEPJE_PATH%\pje-centralmandados-frontend rmdir %EXEPJE_PATH%\pje-centralmandados-frontend /S /Q
cd %EXEPJE_PATH%
git clone https://git.pje.csjt.jus.br/pje2/pje-centralmandados-frontend
)

if "%~1"=="" SET var="pco"
if %var%=="pco" (
TITLE CLONE PJE-CORPORATIVO
if exist %EXEPJE_PATH%\pje-corporativo rmdir %EXEPJE_PATH%\pje-corporativo /S /Q
cd %EXEPJE_PATH%
git clone https://git.pje.csjt.jus.br/pje2/pje-corporativo
)

if "%~1"=="" SET var="pet"
if %var%=="pet" (
TITLE CLONE PJE-ETIQUETAS
if exist %EXEPJE_PATH%\pje-etiquetas-backend rmdir %EXEPJE_PATH%\pje-etiquetas-backend /S /Q
cd %EXEPJE_PATH%
git clone https://git.pje.csjt.jus.br/pje2/etiquetas/pje-etiquetas-backend
)

if "%~1"=="" SET var="pi"
if %var%=="pi" (
TITLE CLONE PJE-INTEGRACAO
if exist %EXEPJE_PATH%\pje-integracao rmdir %EXEPJE_PATH%\pje-integracao /S /Q
cd %EXEPJE_PATH%
git clone https://git.pje.csjt.jus.br/pje2/pje-integracao
)

if "%~1"=="" SET var="ps"
if %var%=="ps" (
TITLE CLONE PJE-SEGURANCA
if exist %EXEPJE_PATH%\pje-seguranca rmdir %EXEPJE_PATH%\pje-seguranca /S /Q
cd %EXEPJE_PATH%
git clone https://git.pje.csjt.jus.br/pje2/pje-seguranca
)

if "%~1"=="" SET var="eb"
if %var%=="eb" (
TITLE CLONE EXE-BACKEND
if exist %EXEPJE_PATH%\exe-backend rmdir %EXEPJE_PATH%\exe-backend /S /Q
cd %EXEPJE_PATH%
git clone https://git.pje.csjt.jus.br/pje-satelites/exe/exe-backend
)

if "%~1"=="" SET var="ef"
if %var%=="ef" (
TITLE CLONE EXE-FRONTEND
if exist %EXEPJE_PATH%\exe-frontend rmdir %EXEPJE_PATH%\exe-frontend /S /Q
cd %EXEPJE_PATH%
git clone https://git.pje.csjt.jus.br/pje-satelites/exe/exe-frontend
)

pause
exit