; ----------------------------------------------------------------------------
; ----------------------------------------------------------------------------
;
; Author:
;   Vinicius Ferraz
;
; Script Function:
;  Executador Inteligente de Testes Automatizados
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
#include <MsgBoxConstants.au3>
#include <GUIListBox.au3>
#include <File.au3>
;#RequireAdmin

Local $wdBtn, $label, $scripts, $update, $exe, $selItems, $report
janela()

Func janela()
	;;================================================================================
	;;GUI CREATION
	;;================================================================================

	$Login = GUICreate("EITA - Executor Inteligente de Testes Automatizados",340,160)
; WEBDRIVER
;	GUICtrlCreateGroup("WEBDRIVER", 4, 4, 120, 100, BitOr($WS_THICKFRAME, $BS_CENTER)) ; I want to cnter "Group 1"
;	 GUICtrlCreateLabel("status", 20, 36)
;	 $label = GUICtrlCreateLabel("", 50, 36, 39, 15)
;	 $wdBtn = GUICtrlCreateButton("Start",40,60,44)

; scripts
	GUICtrlCreateGroup("Scripts", 4, 4, 325, 100, BitOr($WS_THICKFRAME, $BS_CENTER)) ; I want to cnter "Group 1"
   $scripts = GUICtrlCreateList("", 18, 28, 300, 70)
   GUICtrlSetLimit(-1, 200) ; to limit horizontal scrolling
   $update = GUICtrlCreateButton("UPDATE WD",4,110,80, 44)
   $exe = GUICtrlCreateButton("EXECUTE",90,110,160, 44)
   $report = GUICtrlCreateButton("REPORT",255,110,75, 44)
	GUISetState()
	updateList ()

	;;================================================================================
	;;LOGIN LOOP
	;;================================================================================

	While 1
	  updateWD()
		$msg = GUIGetMsg()
		Select
			Case $msg = $GUI_EVENT_CLOSE                                                ; //If the Exit or Close button is clicked, close the app.                                                 ; //Say goodbye
			   Exit
			;Case $msg = $wdBtn
			;   wdbtnclick()
			Case $msg = $update
			   Run("node ../../node_modules/protractor/bin/webdriver-manager update")
			Case $msg = $exe
			    Run("RunTests.bat tmp/tests/" & StringReplace(GUICtrlRead($scripts), ".ts", ".js"))
			 Case $msg = $report
				ShellExecute(@ScriptDir & "/reports/dashboardReport/index.html")
		EndSelect
	WEnd
EndFunc

Func updateList ()
   Local $lst = _FileListToArray ("tests")
   GUICtrlSetData($scripts, "")
   If @error = 1 Then
         MsgBox($MB_SYSTEMMODAL, "", "Path was invalid.")
         Exit
    EndIf
    If @error = 4 Then
         MsgBox($MB_SYSTEMMODAL, "", "No file(s) were found.")
         Exit
    EndIf
    If $lst[0] > 0 Then
 	  For $i = 1 To ($lst[0] )
 		 GUICtrlSetData($scripts, $lst[$i])
 	  Next
    EndIf
EndFunc

Func updateWD()
   If Not WinExists("WEBDRIVER") Then
		 GUICtrlSetData($wdBtn, "start")
		 GUICtrlSetData($label, "stopped")
		 GUICtrlSetColor($label, 0xFF0000) ; Red
		 GUICtrlSetBkColor($label, 0x000000)
   Else
		GUICtrlSetData($wdBtn, "stop")
		GUICtrlSetData($label, "running")
		GUICtrlSetColor($label, 0x00FF00) ; Red
		 GUICtrlSetBkColor($label, 0xFFFFFF)
	  EndIf
   EndFunc

Func wdbtnclick()
   If WinExists("WEBDRIVER") Then
	  ;WinClose("WEBDRIVER)
	  If WinClose("WEBDRIVER") Then
		 GUICtrlSetData($wdBtn, "start")
		 GUICtrlSetData($label, "stopped")
		 GUICtrlSetColor($label, 0xFF0000) ; Red
		 GUICtrlSetBkColor($label, 0x000000)
		 EndIf
    Else
        Run("webdriver-start.bat")
		GUICtrlSetData($wdBtn, "start")
		GUICtrlSetData($label, "running")
		GUICtrlSetColor($label, 0x00FF00) ; Red
		 GUICtrlSetBkColor($label, 0xFFFFFF)
    EndIf
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
