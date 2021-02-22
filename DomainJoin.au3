#RequireAdmin
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=installer.ico
#AutoIt3Wrapper_Res_Fileversion=1.0.1.5
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.2
 Author:         Vinicius M Ferraz

 Script Function: 20/9/16
	DomainJoin

#ce ----------------------------------------------------------------------------

#include <MsgBoxConstants.au3>
#include <ButtonConstants.au3>
#include <GuiButton.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <GuiEdit.au3>
#include <FontConstants.au3>
;#include <Logator.au3>
;#include <ScriptTempoDiretivaInicializacao.au3>

;Constants
Const $JOIN_DOMAIN = 1
Const $ACCT_CREATE = 2

;User / Domain Data
Local $sUser = "root"
Local $sPassword = "XP.smb.15"
Local $sDomain = "TRT15-DOM"
Local $sOU = "";"<ou=Computers,dc=DOMAIN,dc=com>";<ou=myou,dc=mydomain,dc=com>
Local $UI = true, $reboot = True

;altera_tempo_diretiva_inicializacao() ;Chama o script que muda o tempo de expiração da LDAP.

Local $nTentativas = 5

If $cmdline[0] > 0 Then
	For $i = 1 To $cmdline[0]
		If $cmdline[$i] == "-noui" Then
			;LogInfo("ModoSemUI","ModoSemUI")
			$UI = False
		EndIf
		If StringInStr($cmdline[$i],"-t:") > 0 Then
			Local $vars = StringSplit($cmdline[$i],":")
			$nTentativas = $vars[0]>1 ? $vars[2] : 5
			If $nTentativas == 0 Then $UI = False
		ElseIf StringInStr($cmdline[$i],"-noreboot") > 0 Then
				$reboot = False
		EndIf
	Next

EndIf

If $UI Then GUICreate("TRT15 - Apoio ao Usuário", 800, 300)
;;LogInfo("Ingressando no Dominio TRT15-Dom","Ingressando no Dominio TRT15-Dom")
If $UI Then Local $hTitle = GUICtrlCreateLabel("Ingressando Dominio TRT15-DOM", 10, 2, 780, 50, $SS_CENTER, $SS_BLACKFRAME)
If $UI Then GUICtrlSetFont ($hTitle, 26, $FW_BOLD  )
If $UI Then Local $hEdit = GUICtrlCreateEdit("", 10, 60, 780, 220, BitOR($WS_VSCROLL,$ES_READONLY))
If $UI Then GUISetState(@SW_SHOW)

; VERIFICA NOME DA MAQUINA ^[Nn][Tt]-\d{6}$
;If Not StringRegExp(@ComputerName, "^[Nn][Tt]-\d{6}$") Then
;	MsgBox($IDOK + $MB_ICONERROR,"ERRO", "O nome da máquina está errado, o padrão é NT-xxxxxx, é mandatório a correção para inclusão ao dominío.", 30)
;	Exit
;EndIf

; VERIFICA SE O REGISTRO ESTA OK
Local $keyName = "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\LanmanWorkstation\Parameters", _
	$vName1 = "DomainCompatibilityMode", $vName2 = "DNSNameResolutionRequired"

If $UI Then _GUICtrlEdit_AppendText($hEdit, "Verficando se registro está configurado para ingressar em dominios... " & @CRLF)
If LogRegRead("Verifica"&$vName1,$keyName,$vName1) <> 1 Then
	If $UI Then _GUICtrlEdit_AppendText($hEdit, "Corrigindo registro... " & @CRLF)
	LogRegWrite("Escreve"&$vName1,$keyName, $vName1, "REG_DWORD", 1)
EndIf
If LogRegRead("Verifica"&$vName2,$keyName,$vName2) <> 0 Then
	LogRegWrite("Escreve"&$vName2,$keyName, $vName2, "REG_DWORD", 0)
EndIf

#CS Verificação do NetLogon
Local $oWMI = ObjGet("winmgmts:{impersonationLevel=Impersonate}!\\" & @ComputerName & "\root\cimv2")
If StringInStr(@ScriptFullPath,"dominio") > 0 Then
	If $UI Then _GUICtrlEdit_AppendText($hEdit, "Testando NetLogon" & @CRLF)
	Do
		Local $services = $oWMI.ExecQuery('Select * from Win32_Service Where Name = "netlogon"')
		For $svc In $services
			$services = $svc
			;LogInfo("EstadoNetLogo", $svc.State)
			ExitLoop
		Next
		If $svc.State = "Running" Then ExitLoop
	Until False
	If $UI Then _GUICtrlEdit_AppendText($hEdit, "NetLogon Pronto" & @CRLF)
EndIf
#CE
Local $tries = 1

While $tries < $nTentativas
	If $UI Then _GUICtrlEdit_AppendText($hEdit, "Registrando-se com servidor do dominio tentativa " & $tries & @CRLF)
	Local $joined = JoinDomain()
	$tries = $tries + ( $joined ? 10 : 1 )
	If Not $joined And $tries < $nTentativas Then
		If $UI Then _GUICtrlEdit_AppendText($hEdit, "Aguardando 30 segundos para nova tentativa..." & @CRLF)
		Sleep(30000)
	EndIf
WEnd

Local $djDir = "c:\sw_util_trt\DomainJoin" 	, _
	  $dj  = $djDir & "\DomainJoin.exe"	, _
	  $ssc   = "start /min c:\sw_util_trt\causerver\starter\ssc.exe ", _
	  $linkdj = @DesktopCommonDir & "\Ingressar no Dominio TRT.lnk"

If $tries < 10 And StringInStr(@ScriptFullPath,"dominio") > 0 Then ; Falhou na instalação
	If $UI Then _GUICtrlEdit_AppendText($hEdit, @CRLF & @CRLF & "O ingresso ao dominio falhou, será necessário executar 'Ingressar Dominio TRT' após as configurações serem finalizadas." & @CRLF)
	If $UI Then WinActivate("TRT15 - Apoio ao Usuário")
	If Not FileExists($djDir) Then DirCreate($djDir)
	;LogInfo("Atalho", "Criando atalho para ingresso manual ao dominio")
	;LogFileCopy("Copia DomainJoin", "C:\sw_util_trt\dominio\domainjoin.exe", $dj)

	; CRIA BAT
	FileWriteLine($djDir & "\domainjoin.bat", ' @ECHO OFF') ;
	FileWriteLine($djDir & "\domainjoin.bat", ' tasklist /FI "IMAGENAME eq starterserver.exe" 2>NUL | find /I /N "starterserver.exe">NUL') ;
	FileWriteLine($djDir & "\domainjoin.bat",' IF "%ERRORLEVEL%"=="0" ( ') ;
	FileWriteLine($djDir & "\domainjoin.bat", $ssc & $dj) ;
	FileWriteLine($djDir & "\domainjoin.bat", ') ELSE (') ;
	FileWriteLine($djDir & "\domainjoin.bat", ' ECHO STARTERSERVER OFFLINE ERROR ') ;
	FileWriteLine($djDir & "\domainjoin.bat", ' PAUSE') ;
	FileWriteLine($djDir & "\domainjoin.bat", ') ') ;

	FileCreateShortcut($djDir & "\domainjoin.bat", $linkdj, @WindowsDir, "", _
				"DomainJoin: Software para registrar computador no dominio TRT15-Dom",@SystemDir & "\netshell.dll","",1)
	Sleep(10000)
ElseIf $joined And FileExists($linkdj) Or $joined And FileExists($dj) Then ; INTSTALOU PELO ATALHO DO DESKTOP
	;LogFileDelete("DeletaAtalho",$linkdj)
	Uninstall() ; REMOVE O EXECUTAVEL
	If $reboot Then
		If MsgBox($MB_ICONWARNING + $MB_YESNO, "Atenção!", "É preciso reiniciar o computador para terminar a configuração, deseja fazer isso agora?", 15  ) <> $IDNO Then RunWait("shutdown -r -t 00")
	EndIf
ElseIf Not $joined And FileExists($linkdj) Then ; NÃO INSTALOU, ERRO
	MsgBox($MB_ICONERROR + $MB_OKCANCEL, "Erro","O ingresso no dominio TRT15-Dom falhou.",15)
EndIf

Func JoinDomain()
	;LogInfo("JoinDomain","Executando comando para ingressar do TRT15-Dom.")
	Dim $oWMI = ObjGet("winmgmts:{impersonationLevel=Impersonate}!\\" & @ComputerName & "\root\cimv2:Win32_ComputerSystem.Name='" & @ComputerName & "'")
	;LogInfo("IngressandoRede","Executando Método WMI JoinDomainOrWorkGroup")
	Local $retorno
	;Join to Domain

	; JOINDOMAINORWORKGROUP GERA LOG EM C:\WINDOWS\DEBUG\NETSETUP.LOG
	; JOINDOMAINORWORKGROUP É DEPENDENTE DO SERVIÇO NETLOGON (DEVE ESTAR RUNNING)
	; PORÉM DURANTE A EXECUÇÃO DO PROVISIONAMENTO SOMENTE EXECUTANDO O JOINDOMAINWORKGROUP QUE O NETLOGO FICA COM STATE RUNNING
	$retorno = $oWMI.JoinDomainOrWorkGroup($sDomain, $sPassword, $sDomain & "\" & $sUser, $sOU, $JOIN_DOMAIN+$ACCT_CREATE)

	If $retorno <> 0 And $retorno <> 2691 Then
		;_GUICtrlEdit_AppendText($hEdit, "Erro ao registrar-se. Erro: " & $retorno & @CRLF & "Fechando script." )
		Local $erro
		Switch $retorno
			Case 5
				$erro = "Acesso negado"
			Case 87
				$erro =  "Parâmetro incorreto"
			Case 110
				$erro = "O sistema não pode abrir o objeto especificado"
			Case 1323
				$erro = "Não foi possível atualizar a senha"
			case 1326
				$erro = "Falha de Login. Usuário desconhecido ou senha incorreta"
			case 1355
				$erro = "O dominio espedificado não existe ou não foi possível encontrá-lo"
			case 2224
				$erro = "A conta já existe"
			case 2692
				$erro = "O computador não é membro de dominio"
			case 2697
				$erro = "O nome deste computador não está cadastrado na LDAP"
			case Else
				$erro = "Erro não especificado"
		EndSwitch

		If $UI Then _GUICtrlEdit_AppendText($hEdit, "Ingresso no Dominio falhou ( " & $retorno & ": " & $erro & " )." & @CRLF)
		;LogInfo("JoinDomain","Ingresso no Dominio falhou ( " & $retorno & ": " & $erro & " ).")

		Return False
	Else
		If $retorno = 2691 Then
			If $UI Then _GUICtrlEdit_AppendText($hEdit, "Este computador já estava registrado no domínio.")
			;LogInfo("JoinDomain","Sucesso! Este computador foi registrado no domínio.")
		Else
			If $UI Then _GUICtrlEdit_AppendText($hEdit, "Sucesso! Este computador foi registrado no domínio.")
			;LogInfo("JoinDomain","Sucesso! Este computador foi registrado no domínio.")
		EndIf
		Sleep(5000)
		Return True
	EndIf
EndFunc

Func Uninstall()
	;delete all you files here
	Local $s_RemoveBat = @TempDir & "\remove.bat"
	Local $h_RemoveBat = FileOpen($s_RemoveBat, 2)
	FileWrite($h_RemoveBat, ":start" & @CRLF & 'rmdir /S /Q "c:\sw_util_trt\domainjoin"' & @CRLF & 'IF EXIST "' & @ScriptFullPath & '" goto start' & @CRLF & 'del "' & $s_RemoveBat & '"')
	FileClose($h_RemoveBat)
Run($s_RemoveBat, "", @SW_HIDE)
EndFunc

