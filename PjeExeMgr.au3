; ----------------------------------------------------------------------------
; ----------------------------------------------------------------------------
;
; Author:
;   Malu05 aka. Mads Hagbart Lund <mads@madx.dk>
;
; Script Function:
;   Login Script
;
; Notes:
;
; ----------------------------------------------------------------------------
; ----------------------------------------------------------------------------

;;================================================================================
;;INCLUDES
;;================================================================================

#include <GUIConstantsEx.au3>
#include <EditConstants.au3>
#include <ButtonConstants.au3>
#include <WindowsConstants.au3>
#include <File.au3>
;#RequireAdmin

Local $prjFolder = "C:\GIT\EXEPje"
Local $scriptsFolder = @ScriptDir & "\PjeExeMgr"
Local $pjeexemgr =  $scriptsFolder & "\config.cmd"

Local $frontendFolder = $prjFolder & "\exe-frontend"
Local $backendFolder = $prjFolder & "\exe-backendend"
Local $pjeBackendFolder = $prjFolder & "\pje24\backend"
Local $pjeIntegracaoFolder = $prjFolder & "\pje24\integracao"
Local $pjeSegurancaFolder = $prjFolder & "\pje24\seguranca"

Local $installScript = $scriptsFolder & "\install.bat"
Local $updateScript = $scriptsFolder & "\update.bat"
Local $buildScript = $scriptsFolder & "\build.bat"
Local $deployScript = $scriptsFolder & "\deploy.bat"
Local $startBackendScript = $scriptsFolder & "\start-backend-server.bat"
Local $startFrontendScript = $scriptsFolder & "\start-frontend-server.bat"

Local $exebbchField
Local $exefbchField
Local $pjebbchField
Local $pjeetbbchField
Local $pjeibchField
Local $pjesbchField

Local $mvn = "mvn"
Local $wildf
Local $BNCH_PJE_B=17196
Local $BNCH_PJE_I=17172
Local $BNCH_PJE_ET_B=17712
Local $BNCH_PJE_S=9862
Local $BNCH_EXE_B=17196
Local $BNCH_EXE_F=17200

If FileExists($pjeexemgr) Then
	$content = StringSplit(FileRead($pjeexemgr), @CRLF)
	;_ArrayDisplay($content)
	For $i = 0 To $content[0]
		$var = StringSplit($content[$i],"=")
		;_ArrayDisplay($var)
		Switch $var[1]
			Case "SET EXEPJE_PATH"
				$prjFolder = $var[2]
			Case "SET MAVN"
				$mvn = $var[2]
			Case "SET WILDF"
				$wildf = $var[2]
			Case "SET BNCH_PJE_B"
				$BNCH_PJE_B = $var[2]
			Case "SET BNCH_PJE_ET_B"
				$BNCH_PJE_ET_B = $var[2]
			Case "SET BNCH_PJE_I"
				$BNCH_PJE_I = $var[2]
			Case "SET BNCH_PJE_S"
				$BNCH_PJE_S = $var[2]
			Case "SET BNCH_EXE_B"
				$BNCH_EXE_B = $var[2]
			Case "SET BNCH_EXE_F"
				$BNCH_EXE_F = $var[2]
		EndSwitch
	Next
EndIf

janela()


Func janela()
	;;================================================================================
	;;GUI CREATION
	;;================================================================================

	$Login = GUICreate("PJE-EXE Manager",340,240)
; commands
	GUICtrlCreateGroup("Comandos", 4, 4, 120, 150, BitOr($WS_THICKFRAME, $BS_CENTER)) ; I want to cnter "Group 1"
	$ckBoxInstall = GUICtrlCreateCheckbox("Re-Install", 20, 26)
	$ckBoxUpdate = GUICtrlCreateCheckbox("Update", 20, 56)
	$ckBoxBuild = GUICtrlCreateCheckbox("Build", 20, 86)
	$ckBoxDeploy = GUICtrlCreateCheckbox("Deploy", 20, 116)

; applications
	GUICtrlCreateGroup("Applications", 140, 4, 180, 220, BitOr($WS_THICKFRAME, $BS_CENTER)) ; I want to cnter "Group 1"
	$ckBoxExeBackend = GUICtrlCreateCheckbox("Exe Backend", 156, 26)
	$exebbchField = GUICtrlCreateInput($BNCH_EXE_B,256,28,40)
	$ckBoxExeFrontend = GUICtrlCreateCheckbox("Exe Frontend", 156, 56)
	$exefbchField = GUICtrlCreateInput($BNCH_EXE_F,256,58,40)

	$ckBoxPjeBackend = GUICtrlCreateCheckbox("PJE Backend", 156, 96)
	$pjebbchField = GUICtrlCreateInput($BNCH_PJE_B,256,98,40)
	$ckBoxPjeEtiquetasBackend = GUICtrlCreateCheckbox("PJE Etiquetas B.", 156, 126)
	$pjeetbbchField = GUICtrlCreateInput($BNCH_PJE_ET_B,256,128,40)
	$ckBoxPjeIntegracao = GUICtrlCreateCheckbox("PJE Integração", 156, 156)
	$pjeibchField = GUICtrlCreateInput($BNCH_PJE_I,256,158,40)
	$ckBoxPjeSeguranca = GUICtrlCreateCheckbox("PJE Segurança", 156, 186)
	$pjesbchField = GUICtrlCreateInput($BNCH_PJE_S,256,188,40)


	$startBackendButton = GUICtrlCreateButton("R. BE",10,200,44)
	$startFrontendButton = GUICtrlCreateButton("R.FE",54,200,44)
	$GoButton = GUICtrlCreateButton("Go",4,160,60)
	GUISetState()

	;;================================================================================
	;;LOGIN LOOP
	;;================================================================================
	While 1                                                                                             ; //Initialize loop
		$msg = GUIGetMsg()                                                                              ; //Recive Input
		Select
			Case $msg = $GUI_EVENT_CLOSE                                                ; //If the Exit or Close button is clicked, close the app.                                                 ; //Say goodbye
					Exit                     		; //Exit the Application
			Case $msg = $startBackendButton
				Run(@ComSpec & " /c " & $startBackendScript, $scriptsFolder, @SW_SHOW)
			Case $msg = $startFrontendButton
				Run(@ComSpec & " /c " & $startFrontendScript, $scriptsFolder, @SW_SHOW)
			Case $msg = $GoButton
				gravaDados()
				;update
				If GUICtrlRead($ckBoxInstall) = 1 Then
					If GUICtrlRead($ckBoxExeBackend) = 1 Or GUICtrlRead($ckBoxExeFrontend) = 1 Or GUICtrlRead($ckBoxPjeBackend) = 1 Or GUICtrlRead($ckBoxPjeIntegracao) = 1 Or GUICtrlRead($ckBoxPjeSeguranca) = 1 Then
						If GUICtrlRead($ckBoxExeFrontend) = 1 Then
							RunWait(@ComSpec & " /c " & $installScript & " ef", $scriptsFolder, @SW_SHOW)
						EndIf

						If GUICtrlRead($ckBoxExeBackend) = 1  Then
							RunWait(@ComSpec & " /c " & $installScript & " eb", $scriptsFolder, @SW_SHOW)
						EndIf

						If GUICtrlRead($ckBoxPjeBackend) = 1  Then
							RunWait(@ComSpec & " /c " & $installScript & " pb", $scriptsFolder, @SW_SHOW)
						 EndIf
						 If GUICtrlRead($ckBoxPjeEtiqquetasBackend) = 1  Then
							RunWait(@ComSpec & " /c " & $installScript & " et", $scriptsFolder, @SW_SHOW)
						EndIf
						If GUICtrlRead($ckBoxPjeIntegracao) = 1  Then
							RunWait(@ComSpec & " /c " & $installScript & " pi", $scriptsFolder, @SW_SHOW)
						EndIf
						If GUICtrlRead($ckBoxPjeSeguranca) = 1  Then
							RunWait(@ComSpec & " /c " & $installScript & " ps", $scriptsFolder, @SW_SHOW)
						EndIf
					Else
						RunWait(@ComSpec & " /c " & $installScript, $scriptsFolder, @SW_SHOW)
					EndIf
				EndIf
				If GUICtrlRead($ckBoxUpdate) = 1 Then
					If GUICtrlRead($ckBoxExeBackend) = 1 Or GUICtrlRead($ckBoxExeFrontend) = 1 Or GUICtrlRead($ckBoxPjeBackend) = 1 Or GUICtrlRead($ckBoxPjeIntegracao) = 1 Or GUICtrlRead($ckBoxPjeSeguranca) = 1 Then
						If GUICtrlRead($ckBoxExeBackend) = 1  Then
							RunWait(@ComSpec & " /c " & $updateScript & " eb", $scriptsFolder, @SW_SHOW)
						EndIf
						If GUICtrlRead($ckBoxExeFrontend) = 1  Then
							RunWait(@ComSpec & " /c " & $updateScript & " ef", $scriptsFolder, @SW_SHOW)
						EndIf

						If GUICtrlRead($ckBoxPjeBackend) = 1  Then
							RunWait(@ComSpec & " /c " & $updateScript & " pb", $scriptsFolder, @SW_SHOW)
						 EndIf
						 If GUICtrlRead($ckBoxPjeEtiquetasBackend) = 1  Then
							RunWait(@ComSpec & " /c " & $updateScript & " et", $scriptsFolder, @SW_SHOW)
						EndIf
						If GUICtrlRead($ckBoxPjeIntegracao) = 1  Then
							RunWait(@ComSpec & " /c " & $updateScript & " pi", $scriptsFolder, @SW_SHOW)
						EndIf
						If GUICtrlRead($ckBoxPjeSeguranca) = 1  Then
							RunWait(@ComSpec & " /c " & $updateScript & " ps", $scriptsFolder, @SW_SHOW)
						EndIf
					Else
						RunWait(@ComSpec & " /c " & $updateScript, $scriptsFolder, @SW_SHOW)
					EndIf
				EndIf

				If GUICtrlRead($ckBoxBuild) = 1 Then
					If GUICtrlRead($ckBoxExeBackend) = 1 Or GUICtrlRead($ckBoxExeFrontend) = 1 Or GUICtrlRead($ckBoxPjeBackend) = 1 Or GUICtrlRead($ckBoxPjeIntegracao) = 1 Or GUICtrlRead($ckBoxPjeSeguranca) = 1 Then

						If GUICtrlRead($ckBoxPjeBackend) = 1  Then
							RunWait(@ComSpec & " /c " & $buildScript & " pb", $scriptsFolder, @SW_SHOW)
						 EndIf
						 If GUICtrlRead($ckBoxPjeEtiquetasBackend) = 1  Then
							RunWait(@ComSpec & " /c " & $buildScript & " et", $scriptsFolder, @SW_SHOW)
						EndIf
						If GUICtrlRead($ckBoxPjeIntegracao) = 1  Then
							RunWait(@ComSpec & " /c " & $buildScript & " pi", $scriptsFolder, @SW_SHOW)
						EndIf
						If GUICtrlRead($ckBoxPjeSeguranca) = 1  Then
							RunWait(@ComSpec & " /c " & $buildScript & " ps", $scriptsFolder, @SW_SHOW)
						EndIf

						If GUICtrlRead($ckBoxExeBackend) = 1  Then
							RunWait(@ComSpec & " /c " & $buildScript & " eb", $scriptsFolder, @SW_SHOW)
						EndIf
						If GUICtrlRead($ckBoxExeFrontend) = 1  Then
							RunWait(@ComSpec & " /c " & $buildScript & " ef", $scriptsFolder, @SW_SHOW)
						EndIf
					Else
						RunWait(@ComSpec & " /c " & $buildScript, $scriptsFolder, @SW_SHOW)
					EndIf
				EndIf

				If GUICtrlRead($ckBoxDeploy) = 1 Then
					If WinExists("BACKEND-SERVER") AND ( GUICtrlRead($ckBoxExeBackend) = 1 Or GUICtrlRead($ckBoxExeFrontend) = 1 Or GUICtrlRead($ckBoxPjeBackend) = 1 Or GUICtrlRead($ckBoxPjeIntegracao) = 1 Or GUICtrlRead($ckBoxPjeSeguranca) = 1) Then
						If GUICtrlRead($ckBoxExeBackend) = 1  Then
							Run(@ComSpec & " /c " & $deployScript & " eb", $scriptsFolder, @SW_SHOW)
						EndIf
						If GUICtrlRead($ckBoxExeFrontend) = 1  Then
							Run(@ComSpec & " /c " & $deployScript & " ef", $scriptsFolder, @SW_SHOW)
						EndIf

						If GUICtrlRead($ckBoxPjeBackend) = 1  Then
							Run(@ComSpec & " /c " & $deployScript & " pb", $scriptsFolder, @SW_SHOW)
						 EndIf
						 If GUICtrlRead($ckBoxPjeEtiquetasBackend) = 1  Then
							Run(@ComSpec & " /c " & $deployScript & " et", $scriptsFolder, @SW_SHOW)
						EndIf
						If GUICtrlRead($ckBoxPjeIntegracao) = 1  Then
							Run(@ComSpec & " /c " & $deployScript & " pi", $scriptsFolder, @SW_SHOW)
						EndIf
						If GUICtrlRead($ckBoxPjeSeguranca) = 1  Then
							Run(@ComSpec & " /c " & $deployScript & " ps", $scriptsFolder, @SW_SHOW)
						EndIf
					ElseIf WinExists("BACKEND-SERVER") Then
						Run(@ComSpec & " /c " & $deployScript, $scriptsFolder, @SW_SHOW)
					Else
						MsgBox(0,'Erro', 'Para deploy o wildfly deve estar funcionando!')
					EndIf
				EndIf
		EndSelect
	WEnd
EndFunc


Func gravaDados()
	FileDelete($pjeexemgr)
	FileWrite($pjeexemgr,  "SET EXEPJE_PATH=" & $prjFolder & @CRLF & "SET MAVN=" & $mvn & @CRLF _
	& "SET WILDF=" & $wildf & @CRLF  _
	& "SET BNCH_PJE_B=" & GUICtrlRead($pjebbchField) & @CRLF _
	& "SET BNCH_PJE_ET_B=" & GUICtrlRead($pjeetbbchField) & @CRLF _
	& "SET BNCH_PJE_I=" & GUICtrlRead($pjeibchField) & @CRLF _
	& "SET BNCH_PJE_S=" & GUICtrlRead($pjesbchField) & @CRLF _
	& "SET BNCH_EXE_B=" & GUICtrlRead($exebbchField) & @CRLF _
	& "SET BNCH_EXE_F=" & GUICtrlRead($exefbchField) )
EndFunc
