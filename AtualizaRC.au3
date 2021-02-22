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
#include <File.au3>
;#RequireAdmin

Local $numRC = "1.0.0.0"
Local $prjFolder = "C:\GIT\prj"
Local $rcconfig =  "arc.config"
Local $pomFolder = "C:\GIT\prj\RC1.0.0.0\prj_parent"
Local $repPath = "https://svn.trt15.jus.br/svn/repositorioSVN/TRT15/"
Local $pomName = "pom.xml"


If FileExists($rcconfig) Then
	$content = StringSplit(FileRead($rcconfig), @CRLF)
	;_ArrayDisplay($content)
	$numRC = $content[1]
	$prjFolder = $content[3]
	$pomFolder = $content[5]
	$repPath = $content[7]
	$pomName = $content[9]
EndIf

janelaAtualizaRC()


Func janelaAtualizaRC()
	;;================================================================================
	;;GUI CREATION
	;;================================================================================

	$Login = GUICreate("Atualiza RC",280,200)

	GUICtrlCreateLabel("Num. RC",10,12)
	$numRCField = GUICtrlCreateInput($numRC,60,10,80)
	GUICtrlSetTip( -1, "Apenas o número da RC, ex: 1.0.0.0")

	GUICtrlCreateLabel("Pasta Prj",10,42)
	$prjFolderField = GUICtrlCreateInput($prjFolder,60,40,200)
	GUICtrlSetTip( -1, "A pasta do projeto onde será copiada a pasta da RC")

	GUICtrlCreateLabel("POM Prj",10,72)
	$pomFolderField = GUICtrlCreateInput($pomFolder,60,70,200)
	GUICtrlSetTip( -1, "A pasta do checkout onde fica o arquivo pom, geralmente na pasta parent. Obs: Este caminho tem o número da RC, ex: C:\GIT\prj\RC%n%\prj_parent -> observe que o número da RC é substituído por %n% para não precisar atualizar este campo a cada RC")

	GUICtrlCreateLabel("Rep path",10,102)
	$repPathField = GUICtrlCreateInput($repPath,60,100,200)
	GUICtrlSetTip( -1, "O caminho do repositório do projeto, ex: https://svn.trt15.jus.br/svn/repositorioSVN/TRT15/prj/tags/RC -> observe que RC é a pasta onde estão as RCs")

	GUICtrlCreateLabel("POM name",6,132)
	$pomNameField = GUICtrlCreateInput($pomName,60,130,200)
	GUICtrlSetTip( -1, "O nome do arquivo pom que fica geralmente na pasta parent da RC (o nome desse arquivo pode sofrer variações em projetos)")

	; GUICtrlSetTip( -1, "Usuário")	; //Create the Username Input

	$ckBoxDelRcs = GUICtrlCreateCheckbox("Del RCs Locais", 10, 160)
	$ckBoxCout = GUICtrlCreateCheckbox("Checkout", 110, 160)
	$ckBoxinstala = GUICtrlCreateCheckbox("Instala", 10, 180)


	$OKbutton = GUICtrlCreateButton("Go",205,170,60)                                                     ; //Create the Exit Button
	GUISetState()                                                                                       ; //Display the GUI

	;;================================================================================
	;;LOGIN LOOP
	;;================================================================================
	While 1                                                                                             ; //Initialize loop
		$msg = GUIGetMsg()                                                                              ; //Recive Input
		Select
			Case $msg = $GUI_EVENT_CLOSE                                                ; //If the Exit or Close button is clicked, close the app.                                                 ; //Say goodbye
					Exit                     		; //Exit the Application
			Case $msg = $OKbutton
				$numRC = GUICtrlRead($numRCField)
				$prjFolder = GUICtrlRead($prjFolderField)
				$pomFolder = GUICtrlRead($pomFolderField)
				$repPath = GUICtrlRead($repPathField)
				$pomName = GUICtrlRead($pomNameField)
				gravaDados()

				If GUICtrlRead($ckBoxDelRcs) = 1  Then
					$pastas = _FileListToArray($prjFolder, "RC*")
					For $i = 1 To $pastas[0]
						DirRemove ($prjFolder & "/" & $pastas[$i], 1)
					Next
				EndIf

				If GUICtrlRead($ckBoxCout) = 1 Then
					;$p = StringReplace($repPath, "%n%", $numRC)
					RunWait(@ComSpec & " /c " & "svn checkout " & $repPath & "/RC" & $numRC, $prjFolder, @SW_SHOW)
				EndIf

				If GUICtrlRead($ckBoxinstala) = 1 Then
					$p = StringReplace($pomFolder, "%n%", $numRC)
					;MsgBox(-1,"", "mvn -f " & $pomName & " clean install plc:deploy-completo -Dambiente=desenv -Dmaven.test.skip=true::mvn install plc:deploy-completo" & $p)
					RunWait(@ComSpec & " /c " & "mvn -f " & $pomName & " clean install plc:deploy-completo -Dambiente=desenv -Dmaven.test.skip=true::mvn install plc:deploy-completo", $p, @SW_SHOW)
				EndIf


		EndSelect
	WEnd
EndFunc


Func gravaDados()
	FileDelete($rcconfig)
	FileWrite($rcconfig, $numRC & @CRLF & $prjFolder & @CRLF & $pomFolder & @CRLF & $repPath & @CRLF & $pomName)
EndFunc
