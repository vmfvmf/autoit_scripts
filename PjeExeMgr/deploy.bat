@echo off

CALL config.cmd

SET var="%~1" 

IF NOT EXIST %WILDF%\jboss-cli.bat (
	echo "WILDFLY jboss-cli.bat nao encontrado. Verifique o caminho no arquivo config.cmd. Encerrendo."
	pause
	exit
)

if "%~1"=="" SET var="pb"
if %var%=="pb" (
	title DEPLOY PJE-BACKEND
	call %WILDF%\jboss-cli.bat --connect --command="deploy --force %EXEPJE_PATH%\pje-backend\pje-comum-api\target\pje-comum-api.war"
)

if "%~1"=="" SET var="pa"
if %var%=="pa" (
	title DEPLOY PJE-ASSINATURA
	call %WILDF%\jboss-cli.bat --connect --command="deploy --force %EXEPJE_PATH%\pje-assinatura\pje-assinatura-api\target\pje-assinatura-api.war"
)

if "%~1"=="" SET var="pet"
if %var%=="pet" (
	title DEPLOY PJE-ETIQUETAS
	call %WILDF%\jboss-cli.bat --connect --command="deploy --force %EXEPJE_PATH%\pje-etiquetas-backend\pje-etiquetas-api\target\pje-etiquetas-api.war"
)

if "%~1"=="" SET var="pco"
if %var%=="pco" (
	title DEPLOY PJE-CORPORATIVO
	call %WILDF%\jboss-cli.bat --connect --command="deploy --force %EXEPJE_PATH%\pje-corporativo\pje-corporativo-api\target\pje-corporativo-api.war"
)

if "%~1"=="" SET var="pcm"
if %var%=="pcm" (
	title DEPLOY PJE-CENTRAL MANDADOS
	call %WILDF%\jboss-cli.bat --connect --command="deploy --force %EXEPJE_PATH%\pje-centralmandados\pje-centralmandados-api\target\pje-centralmandados-api.war"
)

if "%~1"=="" SET var="pi"
if %var%=="pi" (
	title DEPLOY PJE-INTEGRACAO
	call %WILDF%\jboss-cli.bat --connect --command="deploy --force %EXEPJE_PATH%\pje-integracao\pje-integracao-api\target\pje-integracao-api.war"
)


if "%~1"=="" SET var="eb"
if %var%=="eb" (
	title DEPLOY EXE-BACKEND
	call %WILDF%\jboss-cli.bat --connect --command="deploy --force %EXEPJE_PATH%\exe-backend\exe-backend-api\target\exe-backend-api.war"
)

if "%~1"=="" SET var="ej"
if %var%=="ej" if %var%=="eb" (
	title DEPLOY EXE-BACKEND JOBS
	call %WILDF%\jboss-cli.bat --connect --command="deploy --force %EXEPJE_PATH%\exe-backend\exe-backend-jobs\target\exe-backend-jobs.war"
)

if "%~1"=="" SET var="ps"
if %var%=="ps" (
	title DEPLOY PJE-SEGURANCA
	call %WILDF%\jboss-cli.bat --connect --command="deploy --force %EXEPJE_PATH%\pje-seguranca\pje-seguranca-api\target\pje-seguranca-api.war"
)
exit
