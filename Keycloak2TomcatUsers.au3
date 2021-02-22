#RequireAdmin
#include <File.au3>
#include <Array.au3>
#include <GUIConstantsEx.au3>
#include <EditConstants.au3>

Local $k2tuconfig =  "k2tu.config"
;Local $tomcat = "%CATALINA_HOME%/"
Local $app_nome = "AutoAtendimentoExterno"
Local $tomcat = "C:\apache-tomcat-6\"

Local $confCtx = $tomcat & "conf\Catalina\localhost\" & $app_nome & ".xml"
Local $mtCtx   = $tomcat & "webapps\" & $app_nome & "\META-INF\context.xml"
Local $webWeb  = $tomcat & "webapps\" & $app_nome & "\WEB-INF\web.xml"

Local $valve_kc = '<Valve className="org.keycloak.adapters.tomcat.KeycloakAuthenticatorValve"/>'
Local $valve_j  = '<!--Valve className="org.keycloak.adapters.tomcat.KeycloakAuthenticatorValve"/-->'

Local $uma_auth_kc = '<role-name>uma_authorization</role-name>'
Local $uma_auth_j =  '<!--role-name>uma_authorization</role-name-->'

Local $login_config_form_auth_method_kc = '<auth-method>BASIC</auth-method>'
Local $login_config_form_auth_method_j  = '<auth-method>FORM</auth-method>'

Local $login_config_form_realm_name_kc = '<realm-name>TRT15</realm-name>'
Local $login_config_form_realm_name_j  = '<realm-name>default</realm-name><form-login-config><form-login-page>/res-plc/login/loginPlc.html</form-login-page><form-error-page>/res-plc/login/loginErroPlc.html</form-error-page></form-login-config>'

If FileExists($k2tuconfig) Then
	$content = StringSplit(FileRead($k2tuconfig), @CRLF)
	$app_nome = $content[1]
	$tomcat = $content[3]
EndIf

janela()

Func janela()
	;;================================================================================
	;;GUI CREATION
	;;================================================================================
	$trocasenha = GUICreate("Troca Keycloak/Tomcat-Users",350,100)
	GUICtrlCreateLabel("App_Name",10,10)
	$app_nome_tf = GUICtrlCreateInput($app_nome,85,8,164,18)
	GUICtrlCreateLabel("Tomcat",10,40)
	$tomcat_tf = GUICtrlCreateInput($tomcat,85,38,164,18)
	GUICtrlSetTip( -1, "Digite o(s) nome(s) de usuário(s), 1 por linha.")
	$OKbutton = GUICtrlCreateButton("Trocar",85,70,100)
   GUISetState()
	While 1
		; //Initialize loop
		$msg = GUIGetMsg()
		Select
			Case $msg = $GUI_EVENT_CLOSE
				Exit
			 Case $msg = $OKbutton
				$app_nome = GUICtrlRead($app_nome_tf)
				$tomcat = GUICtrlRead($tomcat_tf)
				alterna_keycloak2tomcat()
				gravaDados()

		EndSelect
	WEnd
 EndFunc

 Func gravaDados()
	FileDelete($k2tuconfig)
	FileWrite($k2tuconfig, $app_nome & @CRLF & $tomcat)
EndFunc

Func alterna_keycloak2tomcat()
   Local $kc2j = CheckStringInFile($confCtx, $valve_kc)

   If $kc2j Then
	  ReplaceInFile($confCtx, $valve_kc, $valve_j)

	  ReplaceInFile($mtCtx, $valve_kc, $valve_j)

	  ReplaceInFile($webWeb, $uma_auth_kc, $uma_auth_j)
	  ReplaceInFile($webWeb, $login_config_form_auth_method_kc, $login_config_form_auth_method_j)
	  ReplaceInFile($webWeb, $login_config_form_realm_name_kc, $login_config_form_realm_name_j)
	  ComentParentTag($webWeb, $uma_auth_kc, 'security-role')
   Else
	  ReplaceInFile($confCtx, $valve_j, $valve_kc)

	  ReplaceInFile($mtCtx, $valve_j, $valve_kc)

	  UncomentParentTag($webWeb, $uma_auth_kc, 'security-role')
	  ReplaceInFile($webWeb, $uma_auth_j, $uma_auth_kc)
	  ReplaceInFile($webWeb, $login_config_form_auth_method_j, $login_config_form_auth_method_kc)
	  ReplaceInFile($webWeb, $login_config_form_realm_name_j, $login_config_form_realm_name_kc)
   EndIf
EndFunc

Func CheckStringInFile($fileP, $string)
   Local $file
   _FileReadToArray($fileP, $file)
   For $a = 1 To $file[0]
      If StringInStr($file[$a], $string) Then
		 Return True
      EndIf
   Next
   Return False
EndFunc

Func ReplaceInFile($filePath, $originalString, $replaceString)
   Local $file, $sText, $a = 1
   _FileReadToArray($filePath, $file)
   For $a = 1 To $file[0]
      If StringInStr($file[$a], $originalString) Then
         If  Not _FileWriteToLine($filePath, $a, $replaceString, True) Then
			MsgBox(0, "Erro", "Houve um erro: "& @error)
		 EndIf
		 Return
      EndIf
   Next
EndFunc

Func ComentParentTag($filePath, $childTag2Find, $parentTagName)
   Local $file, $sText, $a = 1
   _FileReadToArray($filePath, $file)
   For $a = 1 To $file[0]
      If StringInStr($file[$a], $childTag2Find) Then
		 $b = $a
		 Do
			$b = $b -1
		 Until StringInStr($file[$b], $parentTagName)
         If  Not _FileWriteToLine($filePath, $b, StringReplace($file[$b], "<", "<!--"), True) Then
			MsgBox(0, "Erro", "Houve um erro: "& @error)
		 EndIf
		 $closingTagName = "/" & $parentTagName
		 $b = $a
		 Do
			$b = $b +1
		 Until StringInStr($file[$b], $closingTagName)
         If  Not _FileWriteToLine($filePath, $b, StringReplace($file[$b], ">", "-->"), True) Then
			MsgBox(0, "Erro", "Houve um erro: "& @error)
		 EndIf
		 Return
      EndIf
   Next
EndFunc

Func UncomentParentTag($filePath, $childTag2Find, $parentTagName)
   Local $file, $sText, $a = 1
   _FileReadToArray($filePath, $file)
   For $a = 1 To $file[0]
      If StringInStr($file[$a], $childTag2Find) Then
		 $b = $a
		 Do
			$b = $b -1
		 Until StringInStr($file[$b], $parentTagName)
         If  Not _FileWriteToLine($filePath, $b, StringReplace($file[$b], "<!--", "<"), True) Then
			MsgBox(0, "Erro", "Houve um erro: "& @error)
		 EndIf
		 $closingTag = "/" & $parentTagName
		 $b = $a
		 Do
			$b = $b +1
		 Until StringInStr($file[$b], $closingTag)
         If  Not _FileWriteToLine($filePath, $b, StringReplace($file[$b], "-->", ">"), True) Then
			MsgBox(0, "Erro", "Houve um erro: "& @error)
		 EndIf
		 Return
      EndIf
   Next
EndFunc
