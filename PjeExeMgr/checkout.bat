@echo off

call config.cmd

SET var="%~1"

if "%~1"=="" SET var="eb"
if %var%=="eb" (
title UPDATE EXE-BACKEND
cd %EXEPJE_PATH%\exe-backend
git checkout %BNCH_EXE_B%
)

if "%~1"=="" SET var="ef"
if %var%=="ef" (
title UPDATE EXE-FRONTEND
cd %EXEPJE_PATH%\exe-frontend
git checkout %BNCH_EXE_F%
)

if "%~1"=="" SET var="pa"
if %var%=="pa" (
title UPDATE PJE-ASSINATURA
cd %EXEPJE_PATH%\pje-assinatura
git checkout %BNCH_PJE_A%
)

if "%~1"=="" SET var="pb"
if %var%=="pb" (
title UPDATE PJE-BACKEND
cd %EXEPJE_PATH%\pje-backend
git checkout %BNCH_PJE_B%
)

if "%~1"=="" SET var="pcm"
if %var%=="pcm" (
title UPDATE PJE-CENTRAL MANDADOS
cd %EXEPJE_PATH%\pje-centralmandados
git checkout %BNCH_PJE_CM%
)

if "%~1"=="" SET var="pcmf"
if %var%=="pcmf" (
title UPDATE PJE-CENTRAL MANDADOS FRONTEND
cd %EXEPJE_PATH%\pje-centralmandados-frontend
git checkout %BNCH_PJE_CMF%
)

if "%~1"=="" SET var="pco"
if %var%=="pco" (
title UPDATE PJE-CORPORATIVO
cd %EXEPJE_PATH%\pje-corporativo
git checkout %BNCH_PJE_CO%
)

if "%~1"=="" SET var="pet"
if %var%=="pet" (
title UPDATE PJE-ETIQUETAS-BACKEND
cd %EXEPJE_PATH%\pje-etiquetas-backend
git checkout %BNCH_PJE_ET_B%
)

if "%~1"=="" SET var="pi"
if %var%=="pi" (
title UPDATE PJE-INTEGRACAO
cd %EXEPJE_PATH%\pje-integracao
git checkout %BNCH_PJE_I%
)

if "%~1"=="" SET var="ps"
if %var%=="ps" (
title UPDATE PJE-SEGURANCA
cd %EXEPJE_PATH%\pje-seguranca
git checkout %BNCH_PJE_S%
)

if "%~1"=="" SET var="psc"
if %var%=="psc" (
title UPDATE PJE-SCRIPTS
cd %EXEPJE_PATH%\pje-scripts
git checkout %BNCH_PJE_SC%
)

pause
exit
