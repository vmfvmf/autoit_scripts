#include <GUIConstants.au3>
#include <Array.au3>
;#RequireAdmin
Dim $oMyError

; Initializes COM handler
$oMyError = ObjEvent("AutoIt.Error","MyErrFunc")

Global $ado = ObjCreate( "ADODB.Connection" )    ; Create a COM ADODB Object  with the Beta version

; loginOracleDB("viniciusferraz","viniciusferraz")

Func loginOracleDB($usr, $psw)
	With $ado
		; 'Set data source - for OLEDB this is a tns alias, for ODBC it can be 'either a tns alias or a DSN.
		; If "provider" is used this means that the ODBC connections is used via DSN.
		; if Driver is used = "Driver={Microsoft ODBC for Oracle};Server=TNSnames_ora;Uid=demo;Pwd=demo;" then this is a DSN Less connector
		; More Info for Oracle MS KB Q193332
		.ConnectionString =("Driver={Microsoft ODBC for Oracle};Server=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.15.228.190)(PORT=1521))(CONNECT_DATA=(SID=orac)));Uid=" & $usr & ";Pwd=" & $psw & ";")
		.Open
	EndWith
EndFunc


Func executaQry($qry)
	$Rows = 0
	$ado.execute($qry, $Rows, 1)
	return $Rows
EndFunc


Func executaSelect($qry)
	$adors = ObjCreate( "ADODB.RecordSet" )    ; Create a Record Set to handles SQL Records

	With $adors
			.ActiveConnection = $ado
			;.CursorLocation = "adUseClient"
			;.LockType = "adLockReadOnly" ; Set ODBC connection read only
			.Source = $qry
			.Open
	EndWith

	Local $array = $adors.GetRows()
	;While not $adors.EOF
	;	Msgbox( 0, "teste",$adors.Fields( 0 ).Value  )    ; Columns in the AutoIt console use Column Name or Index
	;	$adors.MoveNext                                                ; Go to the next record
	;WEnd
	return $array
EndFunc

; This COM error Handler
Func MyErrFunc()
  $HexNumber=hex($oMyError.number,8)
  Msgbox(0,"AutoItCOM Test","We intercepted a COM Error !"       & @CRLF  & @CRLF & _
             "err.description is: "    & @TAB & $oMyError.description    & @CRLF & _
             "err.windescription:"     & @TAB & $oMyError.windescription & @CRLF & _
             "err.number is: "         & @TAB & $HexNumber              & @CRLF & _
             "err.lastdllerror is: "   & @TAB & $oMyError.lastdllerror   & @CRLF & _
             "err.scriptline is: "     & @TAB & $oMyError.scriptline     & @CRLF & _
             "err.source is: "         & @TAB & $oMyError.source         & @CRLF & _
             "err.helpfile is: "       & @TAB & $oMyError.helpfile       & @CRLF & _
             "err.helpcontext is: "    & @TAB & $oMyError.helpcontext _
            )
  SetError(1)  ; to check for after this function returns
Endfunc