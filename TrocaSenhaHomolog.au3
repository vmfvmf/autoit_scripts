; ----------------------------------------------------------------------------
; ----------------------------------------------------------------------------
;
; Author:
;   Vinicius Ferraz
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
#include "Mycon.au3"
#include <GUIConstantsEx.au3>
#include <EditConstants.au3>

Local $usr = ""
Local $psw = ""
Local $defaultPsw = "Senha123?"
Local $tshconfig =  "tsh.config"

If FileExists($tshconfig) Then
	$content = StringSplit(FileRead($tshconfig), @CRLF)
	$usr = $content[1]
	$psw = $content[3]
	$defaultPsw = $content[5]
	_login()
Else
	janelaLogin()
EndIf

Func janelaLogin()
	;;================================================================================
	;;GUI CREATION
	;;================================================================================
	$Login = GUICreate("Login",220,100)                                                                 ; //Create the login window
	$User = GUICtrlCreateInput($usr,10,10)
	GUICtrlSetTip( -1, "Usuário")	; //Create the Username Input
	$Pass = GUICtrlCreateInput($psw,10,40,-1,-1,$ES_PASSWORD)                                     ; //Create the Password Input
	GUICtrlSetTip( -1, "Senha")
	$OKbutton = GUICtrlCreateButton("Login",110,70,100)                                                     ; //Create the Exit Button
	GUISetState()                                                                                       ; //Display the GUI

	;;================================================================================
	;;LOGIN LOOP
	;;================================================================================
	While 1                                                                                             ; //Initialize loop
		$msg = GUIGetMsg()                                                                              ; //Recive Input
		Select
			Case $msg = $GUI_EVENT_CLOSE                                                ; //If the Exit or Close button is clicked, close the app.                                                 ; //Say goodbye
					Exit                     		; //Exit the Application
			Case $msg = $OKbutton                                                                       ; //If the Login button is clicked goto the login function
				$usr = GUICtrlRead($User)                                                                 ; //Get the username from the Input
				$psw = GUICtrlRead($Pass)
				_login()
				;WinSetState($Login, '', '@SW_HIDE')
		EndSelect
	WEnd
EndFunc

;;================================================================================
;;_login()
;;================================================================================
func _login()
if $usr = "" or $psw = "" Then                                                  ; //Check if the user forgot to put in username and/or password_
    MsgBox(0,"ERROR","Digite o usuário e a senha")                                        ;       -> Then Print this error
Else
    loginOracleDB($usr,$psw)
	gravaDados()
	janelaTrocaSenhas()
EndIf
EndFunc ;--> _login()

Func janelaTrocaSenhas()
	;;================================================================================
	;;GUI CREATION
	;;================================================================================
	$trocasenha = GUICreate("Troca Senha Homologação",220,250)                                                                 ; //Create the login window
	GUICtrlCreateLabel("Senha",10,10)
	$defPsw = GUICtrlCreateInput($defaultPsw,45,8,164,18)
	$Users = GUICtrlCreateEdit("",10,30, 200, 180, BitOR($GUI_SS_DEFAULT_EDIT,$ES_MULTILINE))                                                        ; //Create the Username Input
	GUICtrlSetTip( -1, "Digite o(s) nome(s) de usuário(s), 1 por linha.")
	$OKbutton = GUICtrlCreateButton("Trocar Senha(s)",110,220,100)
		  ; //Display the GUI
	  GUISetState()
	;WinSetState($trocasenha, '', '@SW_SHOW')
	While 1
		; //Initialize loop
		$msg = GUIGetMsg()                                                                              ; //Recive Input
		Select
			Case $msg = $GUI_EVENT_CLOSE                                                ; //If the Exit or Close button is clicked, close the app.                                                 ; //Say goodbye
				Exit                                                                                    ; //Exit the Application
			Case $msg = $OKbutton  ; 																	//If the Login button is clicked goto the login function
				$usrNames = StringSplit(GUICtrlRead($Users), @CRLF)
				For $i = 1 to $usrNames[0]
					If(Not($usrNames[$i] = "")) Then
						$result = executaSelect("select pacldap.Alterasenha('" & $usrNames[$i] & "', '" & $defaultPsw & "') from dual")[0][0]
						If(StringLen($result) > 0) Then
							MsgBox(-1,$result, "A senha do usuário " & $usrNames[$i] & " foi alterada com sucesso para " & $defaultPsw)
							gravaDados()
						Else
							MsgBox(-1,"Erro", "Não foi possível alterar a senha do usuário " & $usrNames[$i])
						EndIf
					EndIf
				Next
				GUICtrlSetData($Users,"")
				;
				;Break
		EndSelect
	WEnd
EndFunc

Func gravaDados()
	FileDelete($tshconfig)
	FileWrite($tshconfig, $usr & @CRLF & $psw & @CRLF & $defaultPsw)
EndFunc
