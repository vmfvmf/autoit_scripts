#Include <Array.au3>
#Include <Chrome.au3>
#include <IE.au3>



#include <IE.au3>


#cs
$sURL = "http://localhost:8080/AutoAtendimentoInterno/f/n/relferiasservidor"
$oIE = _IECreate()
$oIE = _IENavigate($oIE,"http://localhost:8080/AutoAtendimentoInterno/f/n/relferiasservidor")

$colForms = _IEFormGetCollection($oIE) ; get all forms
For $oForm In $colForms ; loop over form collection

    ConsoleWrite("---- FORM " & $oForm.name & " --------------------" & @CRLF)
    $oFormElements = _IEFormElementGetCollection($oForm) ; get all elements
    For $oFormElement In $oFormElements ; loop over element collection

        If StringLower($oFormElement.tagName) == 'input' Then ; it is an input
            ConsoleWrite("> input." & $oFormElement.type & " " & $oFormElement.name & @CRLF)
            _IEFormElementSetValue($oFormElement, "Found You", 0) ; set value of the field
        ElseIf StringLower($oFormElement.tagName) == 'textarea' Then ; it is a textarea
            ConsoleWrite("> textarea " & $oFormElement.name & @CRLF)
            _IEFormElementSetValue($oFormElement, "Found You", 0) ; set value of the field
        EndIf

    Next

Next
#ce
Local $oIE = _IECreate()
_IENavigate($oIE,"http://localhost:8080/AutoAtendimentoInterno/f/n/relferiasservidor")
;_IENavigate($oIE,"javascript:alert('teste');")

$jsUser = "document.getElementById('id_j_username').value='davidbasto';"
$jsPsw = "document.getElementById('id_j_password').value='senha';"
_IELoadWait($oIE)
Sleep(2000)

Local $oForm = _IEFormGetCollection($oIE, 0)
Local $oQuery = _IEFormElementGetCollection($oForm, 1)
_IEFormElementSetValue($oQuery, "AutoIt IE.au3")
#ce
#cs
Local $oForm = _IEGetObjByName($oIE, "id_j_username")
_IEPropertySet($oForm, "value", "teste")
MsgBox($MB_SYSTEMMODAL, "id_j_username", _IEPropertyGet($oForm, "value") & @CRLF)

;_IENavigate($oIE,"javascript:" & $jsUser)
;_IENavigate($oIE,"javascript:" & $jsPsw)

ConsoleWrite("ERRO!"&@error)
;Local $oIE = _IECreate("localhost:8080/AutoAtendimentoInterno/f/n/relferiasservidor"")
#ce
#cs
$gPath = "C:\Users\viniciusferraz\AppData\Local\Google\Chrome SxS\Application\chrome.exe"

; Close any existing Chrome browser
_ChromeShutdown()

; Start Chrome with the URL "http://www.december.com/html/demo/form.html"
_ChromeStartup("http://localhost:8080/AutoAtendimentoInterno/f/n/relferiasservidor", $gPath) ; /f/n/relferiasservidor
;_ChromeDocWaitForReadyStateCompleted()

;_ChromeDocWaitForExistenceByTitle("Autenticação", 2)
Do
	Sleep(5000)
Until( StringInStr(_ChromeDocGetTitle(),  "Autenticação") > 0 Or StringInStr(_ChromeDocGetTitle(),  "Autoatendimento Interno") > 0 )

	IF StringInStr(_ChromeDocGetTitle(),  "Autenticação") > 0 Then
		ConsoleWrite( _ChromeDocGetTitle())
		_ChromeObjSetValueByName("j_username", "davidbasto")
		_ChromeObjSetValueByName("j_password", "senha")
		_ChromeInputClickByType("submit")
		Sleep(1000)
	EndIf

	do
		Sleep(500)
	until(StringInStr(_ChromeDocGetTitle(),  "Autoatendimento Interno") > 0)

ConsoleWrite( _ChromeDocGetTitle())
_ChromeObjSetValueByName("corpo:formulario:servidor", "12215")
_ChromeObjSetValueByName("corpo:formulario:servidorficha", "1")

;sleep(2000)
;ConsoleWrite(_ChromeElementFireOnChangeEvnt("corpo:formulario:servidor')") & @error)
;autoRecuperacaoVinculado('corpo:formulario:servidor');return false;
;corpo:formulario:situacoes:0:indSelecao
; document.getElementById("corpo:formulario:servidor")

ConsoleWrite(_ChromeEval("autoRecuperacaoVinculado('corpo:formulario:servidor')",7))
ConsoleWrite(_ChromeEval('document.getElementById("corpo:formulario:servidor").change()', 7))

_ChromeInputClickByName("corpo:formulario:situacoes:0:indSelecao")
;document.getElementsByClassName("plc-corpo-acao-t")[0].click()
_ChromeEval('document.getElementsByClassName("plc-corpo-acao-t")[0].click()', 7)
Sleep(1000)
ConsoleWrite(@AppDataDir & "\AutoIt3\Chrome Native Messaging Host")

#ce