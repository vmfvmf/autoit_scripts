#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=archive.ico
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <MsgBoxConstants.au3>
#RequireAdmin

DirCopy(@ScriptDir & "\Selenium", "C:\powerlogic\Selenium", 1)
DirCopy(@ScriptDir & "\NetBeansProjects", @MyDocumentsDir & "\NetBeansProjects", 1)
DirCopy(@ScriptDir & "\db6lxkz3.selenium", @AppDataDir & "\Mozilla\Firefox\Profiles\db6lxkz3.selenium", 1)

$key = "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Environment"
$val = "PATH"
$PATH = RegRead($key, $val)

$sAddThisPath = "C:\powerlogic\Selenium\"
$PATH = $PATH & ";" & $sAddThisPath

RegWrite($key,$val,"REG_EXPAND_SZ",$PATH)
EnvUpdate()

If MsgBox ($MB_YESNO + $MB_ICONQUESTION, "Atenção!", "Deseja instalar o NetBeans?") = $IDNO Then
	Exit
EndIf

run(@ScriptDir & "\jdk-8u111-nb-8_2-windows-x64.exe")
Sleep(10000)
WinActivate("Instalador do Java SE Development Kit e dos NetBeans IDEs")
WinWaitActive("Instalador do Java SE Development Kit e dos NetBeans IDEs")
Sleep(2000)
Send("{ENTER}")
Sleep(200)
Send("{ENTER}")
Sleep(200)
Send("{ENTER}")
Sleep(200)
Send("{ENTER}")
Sleep(200)
