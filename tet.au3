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
#include "Mycon.au3"
#include <INet.au3>
#include <GUIConstantsEx.au3>
#include <EditConstants.au3>

;;================================================================================
;;GUI CREATION
;;================================================================================
$Login = GUICreate("Login",220,100)                                                                 ; //Create the login window
$User = GUICtrlCreateInput("Username",10,10)                                                        ; //Create the Username Input
$Pass = GUICtrlCreateInput("Password",10,40,-1,-1,$ES_PASSWORD)                                     ; //Create the Password Input
$OKbutton = GUICtrlCreateButton("Login",110,70,100)                                                 ; //Create the Login Button
$Exit = GUICtrlCreateButton("Exit",10,70,100)                                                       ; //Create the Exit Button
GUISetState()                                                                                       ; //Display the GUI

;;================================================================================
;;LOGIN LOOP
;;================================================================================
While 1                                                                                             ; //Initialize loop
    $msg = GUIGetMsg()                                                                              ; //Recive Input
    Select
        Case $msg = $GUI_EVENT_CLOSE or $msg = $Exit                                                ; //If the Exit or Close button is clicked, close the app.
            MsgBox(0, "Saindo", "Saindo.")                                                         ; //Say goodbye
            Exit                                                                                    ; //Exit the Application
        Case $msg = $OKbutton                                                                       ; //If the Login button is clicked goto the login function
            _login()
    EndSelect
WEnd

;;================================================================================
;;_login()
;;================================================================================
func _login()
$UsernameInput = GUICtrlRead($User)                                                                 ; //Get the username from the Input
$PasswordInput = GUICtrlRead($Pass)                                                                 ; //Get the password from the Input
if $UsernameInput = "" or $PasswordInput = "" Then                                                  ; //Check if the user forgot to put in username and/or password_
    MsgBox(0,"ERROR","Please Enter a Username and Password")                                        ;       -> Then Print this error
Else
    loginOracleDB($UsernameInput,$PasswordInput)
	WinSetState($Login, '', '@SW_HIDE')
EndIf
EndFunc ;--> _login()