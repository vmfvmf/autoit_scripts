#include-once
#Region ADO.au3 - Option, Includes, Setup
#Tidy_Parameters=/sort_funcs /reel
#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w- 4 -w 5 -w 6 -w 7
#include-once
#include "ADO_CONSTANTS.au3"
#include <Array.au3>
#include <Date.au3>
#include <StringConstants.au3>
#include <AutoItConstants.au3>

#EndRegion  ADO.au3 - Option, Includes, Setup

#Region ADO.au3 - UDF Header
; #INDEX# ========================================================================
; Title .........: ADO.au3
; AutoIt Version : 3.3.10.2++
; Language ......: English
; Description ...: A collection of Function for use with an ADO database like MS SQL, MS Access ...
; Author ........: Chris Lambert, mLipok
; Modified ......: eltorro, Elias Assad Neto, CarlH
; Version .......: 2.1.5 BETA - Work in progress 2016/02/24
; ================================================================================

#CS
	2015/08/18
	.	new collection of Functions for EVENT handling - grouped in #Region ADO.au3 - Functions - Event's Handling

	2015/08/24
	.	using ADO_CONSTANTS.au3

	2015/09/02
	.	removed $oConnection = -1, currently all function use ByRef $oConnection

	2015/09/15
	.	Renamed: $_eSQL_RESULT_ >> $ADOSQL_RESULT_ - mLipok
	.	Renamed: $_eSQL_ERROR_ >> $ADOSQL_ERROR_ - mLipok

	2015/10/04 >> 2015/11/06
	.	Renamed: Enums: $ADOSQL_RESULT_ >> $ADO_RET_- mLipok
	.	Renamed: Enums: $ADOSQL_ERROR_ >> $ADO_ERR_- mLipok
	.	Renamed: Enums: $ADO_RET_ERROR >> $ADO_RET_FAILURE- mLipok
	.	Renamed: Enums: $ADO_RET_OK >> $ADO_RET_SUCCESS- mLipok
	.	Renamed: Enums: $ADO_ERR_PARAMETERS >> $ADO_ERR_INVALIDPARAMETERTYPE - mLipok
	.	Renamed: Enums: $ADO_ERR_OK >> $ADO_ERR_SUCCESS - mLipok
	.	Renamed: Function: _SQLVerison >> _ADO_Version - mLipok
	.	Renamed: Function: _SQL_Close >> _ADO_Connection_Close - mLipok
	.	Renamed: Function: _SQL_Startup >> _ADO_Connection_Create - mLipok
	.	Renamed: Function: _SQL_Execute >> _ADO_Execute - mLipok
	.	Renamed: Function: __SQL_EVENT >> __ADO_EVENT - mLipok
	.	New: Function: __ADO_IsValidObjectType - mLipok
	.	New: Enums: $ADO_EXT_INTERNALFUNCTION - mLipok
	.	New: Function: _ADO_Recordset_ToArray - mLipok
	.	Refactored: _SQL_GetTable2D - mLipok
	.	Remove: $ADO_ERR_OTHER >> $ADO_ERR_GENERAL - mLipok
	.	Changed: Function: _SQL_FetchNames : Parameter $oRecordset is now ByRef - mLipok
	.	Added: Function: Parameter: _ADO_Recordset_ToArray >> $bFieldNamesInFirstRow = True - mLipok
	.		this was a speed issue as the entire table was moved step by step
	.	Added: Enums: $ADO_RS_ARRAY_* for use with Return form _ADO_Recordset_ToArray when $bFieldNamesInFirstRow was used - mLipok
	.	Added: Function: _ADO_Recordset_Display - mLipok
	.	Added: Function: __ADO_RecordsetArray_Display - mLipok
	.	Renamed: Variable: $oADODB_Connection >> $oConnection - mLipok
	.	Added: Function: _ADO_Execute: Validation for $oConnection - mLipok
	.	New: Function: __ADO_Command_IsValid - mLipok
	.	New: Function: __ADO_Connection_IsValid - mLipok
	.	New: Function: __ADO_Recordset_IsValid - mLipok
	.	New: Enums: $ADO_ERR_NOCURRENTRECORD - mLipok
	.	Renamed: $ADO_* >> $ADO_* - mLipok
	.	Renamed: _SQL_CommandTimeout >> _ADO_Connection_CommandTimeout - mLipok
	.
	.
	2015/11/06 >>
	.	Removed: Function: _SQL_GetErrMsg() - mLipok
	.	Removed: Variable: $g__sSQL_ErrorDescription - mLipok
	.	Renamed: Function: _SQL_PROVIDER_VERSION >> _ADO_MSSQL_GetProviderVersion - mLipok
	.	Renamed: Function: _SQL_DRIVER_VERSION >> _ADO_MSSQL_GetDriverVersion - mLipok
	.	Removed: Parameter: Function: _SQL_FetchData() $aRow - mLipok
	.	Removed: Parameter: Function: _SQL_FetchNames() $aRow - mLipok
	.	Refactored: _SQL_FetchNames - mLipok
	.	Refactored: _SQL_FetchData - mLipok
	.	Changed: _ADO_Recordset_Display - parameters order - $iAlternateColors <> $bFieldNamesInFirstRow mLipok
	.	Changed: _ADO_Recordset_Display - $bFieldNamesInFirstRow  now default is = False  - mLipok
	.	Added: Function: _ADO_RecordsetArray_IsValid - mLipok
	.	Refactored: Function: __ADO_RecordsetArray_Display - added _ADO_RecordsetArray_IsValid  - mLipok
	.	New: Function: _ADO_RecordsetArray_GetContent - mLipok
	.	New: Function: _ADO_RecordsetArray_GetFieldNames - mLipok
	.
	.
	2016/02/24
	.	Removed: Function: $__sSQL_Last_ConnectionString - mLipok
	.	Removed: Function: _SQL_QuerySingleRowAsString - mLipok
	.	Removed: Function: _SQL_QuerySingleRow - mLipok
	.	Removed: Function: _SQL_GetTable - mLipok
	.	Removed: Function: _SQL_GetTableAsString - mLipok
	.	Removed: Function: _ADO_SQLConnection_DBName - mLipok
	.	Removed: Function: _SQL_RegisterErrorHandler - mLipok
	.	Removed: Function: _SQL_UnRegisterErrorHandler - mLipok
	.	Removed: Function: _SQL_GetTable2D --> look in _ADO_Execute --> third parameter $bReturnAsArray - mLipok
	.	Added: 	Parameter in function: $bReturnAsArray - mLipok
	.
	.	Changed: Function: _ADO_Recordset_ToArray - Parameter - $bFieldNamesInFirstRow is not optional any more - mLipok
	.		(This is first step to change Behavior)
	.	Renamed: Function: _ADO_RecordsetArray_IsValid >> __ADO_RecordsetArray_IsValid - is now INTERNAL - mLipok
	.	Renamed: Function: _SQL_AccessConnect >> _ADO_Connection_OpenAccess - mLipok
	.	Renamed: Function: _SQL_ExcelConnect >> _ADO_Connection_OpenExcel - mLipok
	.	Renamed: Function: _ADO_Connection_OpenJet >> _ADO_Connection_OpenJet - mLipok
	.	Renamed: Function: _ADO_SQLConnectionOpen >> _ADO_Connection_OpenMSSQL - mLipok
	.	Refactored:	_ADO_Connection_OpenMSSQL : $sAPPNAME - mLipok
	.	Change:	_ADO_Connection_OpenMSSQL : parameter : reordering - mLipok
	.	Added:	_ADO_Connection_OpenMSSQL : parameter : $sWSID - mLipok
	.	Added:	_ADO_Connection_OpenMSSQL : parameter : $bUseProviderInsteadDriver - mLipok
	.	Change:	__ADO_MSSQL_CONNECTION_STRING_SQLAuth : parameter : reordering - mLipok
	.	Added:	__ADO_MSSQL_CONNECTION_STRING_SQLAuth : parameter : $sAPPNAME - mLipok
	.	Added:	Function: _ADO_Connection_PropertiesToArray - mLipok
	.
	2016/02/24 FIRST PUBLIC RELEASE
	.
	.
	.
	.
	.
	@LAST
	.
	.
	; 2015-09-01 _ADO_EVENTS_SetUp() not working properly
	.
	.
	. TODO:  Descripition to check:  On Success  - Returns $ADO_RET_SUCCESS

#CE

#EndRegion  ADO.au3 - UDF Header

#Region ADO.au3 - Variable Declaration

Global $__g_fnFetchProgress = Null
; #VARIABLES# ====================================================================
Global Const $__sSQL_UDFVersion = "2.1.5 BETA" ;   2016/02/24

Global Enum _
		$ADO_ERR_SUCCESS, _ ;			 	No Error
		$ADO_ERR_GENERAL, _ ;   			General - some ADO Error - Not classified type of error
		$ADO_ERR_COMERROR, _ ;   			COM Error - check your COM Error Handler
		$ADO_ERR_COMHANDLER, _ ;   			COM Error Handler Registration
		$ADO_ERR_CONNECTION, _ ;   			$oConection.Open 	- Opening error
		$ADO_ERR_ISNOTOBJECT, _ ;			Function Parameters error - Expected/Required Object Type
		$ADO_ERR_INVALIDOBJECTTYPE, _ ;		Function Parameters error - Expected/Required Object Type is returned in Return Value as string
		$ADO_ERR_INVALIDPARAMETERTYPE, _ ;  Function Parameters error - Invalid Variable type passed to the function
		$ADO_ERR_INVALIDPARAMETERVALUE, _ ; Function Parameters error - Invalid value passed to the function
		$ADO_ERR_INVALIDARRAY, _ ;			Function Parameters error - The Recordset is Empty
		$ADO_ERR_RECORDSETEMPTY, _ ;		The Recordset is Empty
		$ADO_ERR_NOCURRENTRECORD, _ ;		The Recordset has no current record
		$ADO_ERR_ENUMCOUNTER ;-------------	just for testing

Global Enum _
		$ADO_EXT_DEFAULT, _ ;				default Extended Value
		$ADO_EXT_PARAM1, _ ;				Error Occurs in 1-Parameter
		$ADO_EXT_PARAM2, _ ;				Error Occurs in 2-Parameter
		$ADO_EXT_PARAM3, _ ;				Error Occurs in 3-Parameter
		$ADO_EXT_PARAM4, _ ;				Error Occurs in 4-Parameter
		$ADO_EXT_PARAM5, _ ;				Error Occurs in 5-Parameter
		$ADO_EXT_PARAM6, _ ;				Error Occurs in 6-Parameter
		$ADO_EXT_INTERNALFUNCTION, _ ;		Error Related to internal Function - should not happend - UDF Dev make something wrong ???
		$ADO_EXT_ENUMCOUNTER ;-------------	just for testing

Global Enum _
		$ADO_RET_FAILURE = -1, _ ;			Failure result
		$ADO_RET_SUCCESS = 1, _ ;			Successful result
		$ADO_RET_ENUMCOUNTER ;-------------	just for testing

Global Enum _
		$ADO_RS_ARRAY_GUID, _ ;				Array GUID
		$ADO_RS_ARRAY_FIELDNAMES, _ ;		Array index for inner FileNames Array
		$ADO_RS_ARRAY_RSCONTENT, _ ;		Array index for inner Recordset Array
		$ADO_RS_ARRAY_ENUMCOUNTR ;--------- just for testing

Global Const $ADO_RS_GUID = '{2399DBEE-2450-462D-B102-9094A9EB5D02}'
#EndRegion  ADO.au3 - Variable Declaration

#Region ADO.au3 - Functions

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __ADO_RecordsetArray_Display
; Description ...:
; Syntax ........: __ADO_RecordsetArray_Display(Byref $aRocordset[, $sTitle = ''[, $iAlternateColors = Default]])
; Parameters ....: $aRocordset          - [in/out] an array of unknowns.
;                  $sTitle              - [optional] a string value. Default is ''.
;                  $iAlternateColors    - [optional] an integer value. Default is Default.
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __ADO_RecordsetArray_Display(ByRef $aRocordset, $sTitle = '', $iAlternateColors = Default)
	If __ADO_RecordsetArray_IsValid($aRocordset) Then
		Local $sArrayHeader = _ArrayToString($aRocordset[$ADO_RS_ARRAY_FIELDNAMES], '|')
		Local $aSelect = _ADO_RecordsetArray_GetContent($aRocordset)
		_ArrayDisplay($aSelect, $sTitle, "", 0, '|', $sArrayHeader, Default, $iAlternateColors)
		If @error Then
			Return SetError($ADO_ERR_GENERAL, $ADO_EXT_DEFAULT, $ADO_RET_FAILURE)
		Else
			Return SetError($ADO_ERR_SUCCESS, $ADO_EXT_DEFAULT, $ADO_RET_SUCCESS)
		EndIf
	ElseIf UBound($aRocordset) Then
		_ArrayDisplay($aRocordset, $sTitle, "", 0, Default, Default, Default, $iAlternateColors)
		If @error Then
			Return SetError($ADO_ERR_GENERAL, $ADO_EXT_DEFAULT, $ADO_RET_FAILURE)
		Else
			Return SetError($ADO_ERR_SUCCESS, $ADO_EXT_DEFAULT, $ADO_RET_SUCCESS)
		EndIf
	Else
		Return SetError($ADO_ERR_INVALIDPARAMETERVALUE, $ADO_EXT_DEFAULT, $ADO_RET_FAILURE)
	EndIf
EndFunc    ;==>__ADO_RecordsetArray_Display

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __ADO_RecordsetArray_IsValid
; Description ...:
; Syntax ........: __ADO_RecordsetArray_IsValid(Byref $aRocordset)
; Parameters ....: $aRocordset          - [in/out] an array of unknowns.
; Return values .: True/False
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __ADO_RecordsetArray_IsValid(ByRef $aRocordset)
	If _
			UBound($aRocordset, $UBOUND_DIMENSIONS) = 1 _
			And UBound($aRocordset, $UBOUND_ROWS) = $ADO_RS_ARRAY_ENUMCOUNTR _
			And $aRocordset[$ADO_RS_ARRAY_GUID] = $ADO_RS_GUID _
			 Then
		Return SetError($ADO_ERR_SUCCESS, $ADO_EXT_DEFAULT, True)
	EndIf
	Return SetError($ADO_ERR_INVALIDARRAY, $ADO_EXT_DEFAULT, False)
EndFunc    ;==>__ADO_RecordsetArray_IsValid

; #FUNCTION# ====================================================================================================================
; Name ..........: _ADO_Recordset_Display
; Description ...:
; Syntax ........: _ADO_Recordset_Display(Byref $vRocordset[, $sTitle = ''[, $iAlternateColors = Default[,
;                  $bFieldNamesInFirstRow = False]]])
; Parameters ....: $vRocordset          - [in/out] a variant value.
;                  $sTitle              - [optional] a string value. Default is ''.
;                  $iAlternateColors    - [optional] an integer value. Default is Default.
;                  $bFieldNamesInFirstRow- [optional] a boolean value. Default is False.
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _ADO_Recordset_Display(ByRef $vRocordset, $sTitle = '', $iAlternateColors = Default, $bFieldNamesInFirstRow = False)
	Local $vResult = $ADO_RET_FAILURE
	If UBound($vRocordset) Then
		$vResult = __ADO_RecordsetArray_Display($vRocordset, $sTitle)
		Return SetError(@error, @extended, $vResult)
	ElseIf __ADO_Recordset_IsNotEmpty($vRocordset) = $ADO_RET_SUCCESS Then
		Local $aRecordset_GetRowsResult = _ADO_Recordset_ToArray($vRocordset, $bFieldNamesInFirstRow)
		$vResult = __ADO_RecordsetArray_Display($aRecordset_GetRowsResult, $sTitle, $iAlternateColors)
		Return SetError(@error, @extended, $vResult)
	Else
		Return SetError(@error, @extended, $ADO_RET_FAILURE) ; @error and @extended returned from __ADO_Recordset_IsValid
	EndIf
EndFunc    ;==>_ADO_Recordset_Display

; #FUNCTION# ====================================================================================================================
; Name ..........: _ADO_Recordset_Find
; Description ...: Searches a Recordset for the row that satisfies the specified criteria.
; Syntax ........: _ADO_Recordset_Find(Byref $oRecordset, $Criteria[, $SkipRows = 0[, $SearchDirection = $ADO_adSearchForward[, $Start = $ADO_adBookmarkCurrent]]])
; Parameters ....: $oRecordset          - [in/out] An unknown value.
;                  $Criteria            - An unknown value.
;                  $SkipRows            - [optional] An unknown value. Default is 0.
;                  $SearchDirection     - [optional] An unknown value. Default is $ADO_adSearchForward.
;                  $Start               - [optional] An unknown value. Default is $ADO_adBookmarkCurrent.
; Return values .: None - see remarks
; Author ........: mLipok
; Modified ......:
; Remarks .......: If the criteria is met, the current row position is set on the found record; otherwise, the position is set to the end (or start) of the Recordset.
; Related .......:
; Link ..........: http://msdn.microsoft.com/en-us/library/windows/desktop/ms676117(v=vs.85).aspx
; Example .......: No
; ===============================================================================================================================
Func _ADO_Recordset_Find(ByRef $oRecordset, $Criteria, $SkipRows = 0, $SearchDirection = $ADO_adSearchForward, $Start = $ADO_adBookmarkCurrent)
	; Error handler, automatic cleanup at end of function
	Local $oADO_COM_ErrorHandler = ObjEvent("AutoIt.Error", __ADO_ComErrorHandler_InternalFunction)
	If @error Then Return SetError($ADO_ERR_COMHANDLER, @error, $ADO_RET_FAILURE)
	#forceref $oADO_COM_ErrorHandler
	__ADO_Recordset_IsNotEmpty($oRecordset)
	If @error Then Return SetError(@error, @extended, $ADO_RET_FAILURE)

	$oRecordset.Find($Criteria, $SkipRows, $SearchDirection, $Start)
	If @error Then Return SetError($ADO_ERR_COMERROR, @error, $ADO_RET_FAILURE)

	Return SetError($ADO_ERR_SUCCESS, $ADO_EXT_DEFAULT, $ADO_RET_SUCCESS)

EndFunc    ;==>_ADO_Recordset_Find

; #FUNCTION# ====================================================================================================================
; Name ..........: _ADO_Recordset_ToArray
; Description ...:
; Syntax ........: _ADO_Recordset_ToArray(Byref $oRecordset, $bFieldNamesInFirstRow)
; Parameters ....: $oRecordset          - [in/out] an object.
;                  $bFieldNamesInFirstRow- a boolean value.
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _ADO_Recordset_ToArray(ByRef $oRecordset, $bFieldNamesInFirstRow)
	; Error handler, automatic cleanup at end of function
	Local $oADO_COM_ErrorHandler = ObjEvent("AutoIt.Error", __ADO_ComErrorHandler_InternalFunction)
	If @error Then Return SetError($ADO_ERR_COMHANDLER, @error, $ADO_RET_FAILURE)
	#forceref $oADO_COM_ErrorHandler

	__ADO_Recordset_IsNotEmpty($oRecordset)
	If @error Then Return SetError(@error, @extended, $ADO_RET_FAILURE)

	; save current Recordset rows postion to $oRecordset_Bookmark
	Local $oRecordset_Bookmark = Null
	If $oRecordset.Supports($ADO_adBookmark) Then $oRecordset_Bookmark = $oRecordset.Bookmark

	Local $aRecordset_GetRowsResult = $oRecordset.GetRows()
	If @error Then ; Trap COM error, report and return
		Return SetError($ADO_ERR_COMERROR, @error, $ADO_RET_FAILURE)
	ElseIf UBound($aRecordset_GetRowsResult) Then
		Local $aResult[0]

		; Restore Recordset row position from stored $oRecordset_Bookmark
		If $oRecordset_Bookmark = Null Then
			$oRecordset.moveFirst()
		Else
			$oRecordset.Bookmark = $oRecordset_Bookmark
		EndIf

		Local $iColumns_count = UBound($aRecordset_GetRowsResult, $UBOUND_COLUMNS)
		Local $iRows_count = UBound($aRecordset_GetRowsResult)

		If $bFieldNamesInFirstRow Then
			; Adjust the array to fit the column names and move all data down 1 row
			ReDim $aRecordset_GetRowsResult[$iRows_count + 1][$iColumns_count]

			; Move all records down
			For $iRow_idx = $iRows_count To 1 Step -1
				For $y = 0 To $iColumns_count - 1
					$aRecordset_GetRowsResult[$iRow_idx][$y] = $aRecordset_GetRowsResult[$iRow_idx - 1][$y]
				Next
			Next

			; Add the coloumn names
			For $iCol_idx = 0 To $iColumns_count - 1 ;get the column names and put into 0 array element
				$aRecordset_GetRowsResult[0][$iCol_idx] = $oRecordset.Fields($iCol_idx).Name
			Next
			$aResult = $aRecordset_GetRowsResult
			Return SetError($ADO_ERR_SUCCESS, $iRows_count + 1, $aResult)
		Else
			ReDim $aResult[$ADO_RS_ARRAY_ENUMCOUNTR]
			Local $aFiledNames_Temp[$iColumns_count]

			For $iCol_idx = 0 To $iColumns_count - 1 ;get the column names and put into 0 array element
				$aFiledNames_Temp[$iCol_idx] = $oRecordset.Fields($iCol_idx).Name
			Next
			$aResult[$ADO_RS_ARRAY_GUID] = $ADO_RS_GUID
			$aResult[$ADO_RS_ARRAY_FIELDNAMES] = $aFiledNames_Temp
			$aResult[$ADO_RS_ARRAY_RSCONTENT] = $aRecordset_GetRowsResult
			Return SetError($ADO_ERR_SUCCESS, $iRows_count, $aResult)
		EndIf
	EndIf

	Return SetError($ADO_ERR_RECORDSETEMPTY, $ADO_EXT_DEFAULT, $ADO_RET_FAILURE)
EndFunc    ;==>_ADO_Recordset_ToArray

; #FUNCTION# ====================================================================================================================
; Name ..........: _ADO_Recordset_ToString
; Description ...:
; Syntax ........: _ADO_Recordset_ToString(Byref $oRecordset[, $sDelim = "|"[, $bReturnColumnNames = True]])
; Parameters ....: $oRecordset          - [in/out] an object.
;                  $sDelim              - [optional] a string value. Default is "|".
;                  $bReturnColumnNames  - [optional] a boolean value. Default is True.
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _ADO_Recordset_ToString(ByRef $oRecordset, $sDelim = "|", $bReturnColumnNames = True)
	; Error handler, automatic cleanup at end of function
	Local $oADO_COM_ErrorHandler = ObjEvent("AutoIt.Error", __ADO_ComErrorHandler_InternalFunction)
	If @error Then Return SetError($ADO_ERR_COMHANDLER, @error, $ADO_RET_FAILURE)
	#forceref $oADO_COM_ErrorHandler

	__ADO_Recordset_IsNotEmpty($oRecordset)
	If @error Then Return SetError(@error, @extended, $ADO_RET_FAILURE)

	#forceref $bReturnColumnNames ; no yet implemented

	; save current Recordset rows postion to $oRecordset_Bookmark
	Local $oRecordset_Bookmark = Null
	If $oRecordset.Supports($ADO_adBookmark) Then $oRecordset_Bookmark = $oRecordset.Bookmark

	; GetString Method (ADO)
	; https://msdn.microsoft.com/en-us/library/ms676975(v=vs.85).aspx
	Local $sString = $oRecordset.GetString($ADO_adClipString, $oRecordset.RecordCount, $sDelim, @CR, 'Null')
	If @error Then ; Trap COM error, report and return
		Return SetError($ADO_ERR_COMERROR, @error, $ADO_RET_FAILURE)
	ElseIf IsString($sString) Then
		; Restore Recordset row position from stored $oRecordset_Bookmark
		If $oRecordset_Bookmark = Null Then
			$oRecordset.moveFirst()
		Else
			$oRecordset.Bookmark = $oRecordset_Bookmark
		EndIf

		Return SetError($ADO_ERR_SUCCESS, $oRecordset.RecordCount, $sString)
	Else
		Return SetError($ADO_ERR_RECORDSETEMPTY, $ADO_EXT_DEFAULT, $ADO_RET_FAILURE)
	EndIf
EndFunc    ;==>_ADO_Recordset_ToString

; #FUNCTION# ====================================================================================================================
; Name ..........: _ADO_RecordsetArray_GetContent
; Description ...:
; Syntax ........: _ADO_RecordsetArray_GetContent(Byref $aRocordset)
; Parameters ....: $aRocordset          - [in/out] an array of unknowns.
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _ADO_RecordsetArray_GetContent(ByRef $aRocordset)
	__ADO_RecordsetArray_IsValid($aRocordset)
	If @error Then Return SetError(@error, @extended, Null)

	Return SetError($ADO_ERR_SUCCESS, $ADO_EXT_DEFAULT, $aRocordset[$ADO_RS_ARRAY_RSCONTENT])
EndFunc    ;==>_ADO_RecordsetArray_GetContent

; #FUNCTION# ====================================================================================================================
; Name ..........: _ADO_RecordsetArray_GetFieldNames
; Description ...:
; Syntax ........: _ADO_RecordsetArray_GetFieldNames(Byref $aRocordset)
; Parameters ....: $aRocordset          - [in/out] an array of unknowns.
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _ADO_RecordsetArray_GetFieldNames(ByRef $aRocordset)
	__ADO_RecordsetArray_IsValid($aRocordset)
	If @error Then Return SetError(@error, @extended, Null)

	Return SetError($ADO_ERR_SUCCESS, $ADO_EXT_DEFAULT, $aRocordset[$ADO_RS_ARRAY_FIELDNAMES])
EndFunc    ;==>_ADO_RecordsetArray_GetFieldNames
#EndRegion  ADO.au3 - Functions

#Region ADO.au3 - Functions - Connection & Management

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __ADO_Command_IsValid
; Description ...:
; Syntax ........: __ADO_Command_IsValid(Byref $oCommand)
; Parameters ....: $oCommand            - [in/out] an object.
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __ADO_Command_IsValid(ByRef $oCommand)
	Local $iValidationResult = __ADO_IsValidObjectType($oCommand, 'ADODB.Command')
	Return SetError(@error, @extended, $iValidationResult)
EndFunc    ;==>__ADO_Command_IsValid

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __ADO_Connection_IsValid
; Description ...:
; Syntax ........: __ADO_Connection_IsValid(Byref $oConnection)
; Parameters ....: $oConnection         - [in/out] an object. ADODB.Connection object.
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __ADO_Connection_IsValid(ByRef $oConnection)
	Local $iValidationResult = __ADO_IsValidObjectType($oConnection, 'ADODB.Connection')
	Return SetError(@error, @extended, $iValidationResult)
EndFunc    ;==>__ADO_Connection_IsValid

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __ADO_IsValidObjectType
; Description ...:
; Syntax ........: __ADO_IsValidObjectType(Byref $oObjectToCheck, $sRequiredProgID)
; Parameters ....: $oObjectToCheck      - [in/out] an object.
;                  $sRequiredProgID     - a string value.
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __ADO_IsValidObjectType(ByRef $oObjectToCheck, $sRequiredProgID)
	If Not IsString($sRequiredProgID) Then
		Return SetError($ADO_ERR_INVALIDPARAMETERTYPE, $ADO_EXT_INTERNALFUNCTION, $ADO_RET_FAILURE)
	ElseIf $sRequiredProgID = '' Then
		Return SetError($ADO_ERR_INVALIDPARAMETERVALUE, $ADO_EXT_INTERNALFUNCTION, $ADO_RET_FAILURE)
	ElseIf Not IsObj($oObjectToCheck) Then
		Return SetError($ADO_ERR_ISNOTOBJECT, $ADO_EXT_DEFAULT, $ADO_RET_FAILURE)
	ElseIf StringInStr(ObjName($oObjectToCheck, $OBJ_PROGID), $sRequiredProgID) = 0 Then
		Return SetError($ADO_ERR_INVALIDOBJECTTYPE, $ADO_EXT_DEFAULT, $ADO_RET_FAILURE)
	Else
		Return SetError($ADO_ERR_SUCCESS, $ADO_EXT_DEFAULT, $ADO_RET_SUCCESS)
	EndIf
EndFunc    ;==>__ADO_IsValidObjectType

; #FUNCTION# ====================================================================================================================
; Name ..........: __ADO_MSSQL_CONNECTION_STRING_SQLAuth
; Description ...:
; Syntax ........: __ADO_MSSQL_CONNECTION_STRING_SQLAuth($sServer, $sDataBase, $sUserName, $sPassword[, $bUseProviderInsteadDriver = True])
; Parameters ....: $sServer             - A string value.
;                  $sDataBase           - A string value.
;                  $sUserName           - A string value.
;                  $sPassword           - A string value.
;                  $bUseProviderInsteadDriver- [optional] A binary value. Default is True.
; Return values .: $sConnectionString
; Author ........: mLipok
; Modified ......:
; Remarks .......: TODO - description
; Related .......:
; Link ..........: https://msdn.microsoft.com/pl-pl/library/ms130822(v=sql.110).aspx
; Example .......: No
; ===============================================================================================================================
Func __ADO_MSSQL_CONNECTION_STRING_SQLAuth($sServer, $sDataBase, $sUserName, $sPassword, $sAppName = Default, $bUseProviderInsteadDriver = True)
	Local Static $sConnectionString = ''

	Local Static $sLastParameters = Default
	Local $sNewParameters = $sServer & $sDataBase & $sUserName & $sPassword & $sAppName & $bUseProviderInsteadDriver

	If $sLastParameters <> $sNewParameters Then
		If $bUseProviderInsteadDriver Then
			$sConnectionString = "PROVIDER=" & _ADO_MSSQL_GetProviderVersion() & ";SERVER=" & $sServer & ";DATABASE=" & $sDataBase & ";UID=" & $sUserName & ";PWD=" & $sPassword & ";"
			If $sAppName <> Default And $sAppName <> '' Then $sConnectionString &= 'Application Name=' & $sAppName & ';'
		Else
			$sConnectionString = "DRIVER={" & _ADO_MSSQL_GetDriverVersion() & "};SERVER=" & $sServer & ";DATABASE=" & $sDataBase & ";UID=" & $sUserName & ";PWD=" & $sPassword & ";"
			If $sAppName <> Default And $sAppName <> '' Then $sConnectionString &= 'APPNAME=' & $sAppName & ';'
		EndIf
		$sLastParameters = $sNewParameters
	EndIf

	Return SetError($ADO_ERR_SUCCESS, $ADO_EXT_DEFAULT, $sConnectionString)

EndFunc    ;==>__ADO_MSSQL_CONNECTION_STRING_SQLAuth

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __ADO_Recordset_IsNotEmpty
; Description ...:
; Syntax ........: __ADO_Recordset_IsNotEmpty(Byref $oRecordset)
; Parameters ....: $oRecordset          - [in/out] an object.
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __ADO_Recordset_IsNotEmpty(ByRef $oRecordset)
	; Error handler, automatic cleanup at end of function
	Local $oADO_COM_ErrorHandler = ObjEvent("AutoIt.Error", __ADO_ComErrorHandler_InternalFunction)
	If @error Then Return SetError($ADO_ERR_COMHANDLER, @error, $ADO_RET_FAILURE)
	#forceref $oADO_COM_ErrorHandler

	__ADO_Recordset_IsValid($oRecordset)
	If @error Then
		Return SetError(@error, @extended, $ADO_RET_FAILURE)
	ElseIf $oRecordset.bof = -1 And $oRecordset.eof = True Then ; no current record
		Return SetError($ADO_ERR_NOCURRENTRECORD, $ADO_EXT_DEFAULT, $ADO_RET_FAILURE)
	ElseIf $oRecordset.RecordCount = 0 Then
		Return SetError($ADO_ERR_RECORDSETEMPTY, $ADO_EXT_DEFAULT, $ADO_RET_FAILURE)
	Else
		Return SetError($ADO_ERR_SUCCESS, $ADO_EXT_DEFAULT, $ADO_RET_SUCCESS)
	EndIf
EndFunc    ;==>__ADO_Recordset_IsNotEmpty

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __ADO_Recordset_IsValid
; Description ...:
; Syntax ........: __ADO_Recordset_IsValid(Byref $oRecordset)
; Parameters ....: $oRecordset          - [in/out] an object.
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __ADO_Recordset_IsValid(ByRef $oRecordset)
	Local $iValidationResult = __ADO_IsValidObjectType($oRecordset, 'ADODB.Recordset')
	Return SetError(@error, @extended, $iValidationResult)
EndFunc    ;==>__ADO_Recordset_IsValid

; #FUNCTION# ====================================================================================================================
; Name ..........: _ADO_Command
; Description ...:
; Syntax ........: _ADO_Command(Byref $oConnection[, $sQuery = ''[, $iCommandType = $ADO_adCmdText]])
; Parameters ....: $oConnection         - [in/out] an object. ADODB.Connection object.
;                  $sQuery              - [optional] a string value. Default is ''.
;                  $iCommandType        - [optional] an integer value. Default is $ADO_adCmdText.
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _ADO_Command(ByRef $oConnection, $sQuery = '', $iCommandType = $ADO_adCmdText)
	; Error handler, automatic cleanup at end of function
	Local $oADO_COM_ErrorHandler = ObjEvent("AutoIt.Error", __ADO_ComErrorHandler_InternalFunction)
	If @error Then Return SetError($ADO_ERR_COMHANDLER, @error, $ADO_RET_FAILURE)
	#forceref $oADO_COM_ErrorHandler

	__ADO_Connection_IsValid($oConnection)
	If @error Then
		Return SetError(@error, @extended, $ADO_RET_FAILURE)
	ElseIf Not IsString($sQuery) Then
		Return SetError($ADO_ERR_INVALIDPARAMETERTYPE, $ADO_EXT_DEFAULT, $ADO_RET_FAILURE)
	EndIf

	; How To Determine Number of Records Affected by an ADO UPDATE
	; https://support.microsoft.com/en-us/kb/195048
	; Use the command object to perform an UPDATE and return the count of affected records.
	Local $iRecordsAffected = -1
	Local $oCommand = ObjCreate("ADODB.Command")
	If @error Then Return SetError(@error, @extended, $ADO_RET_FAILURE)
	With $oCommand
		.CommandType = $iCommandType
		.ActiveConnection = $oConnection
		.CommandText = $sQuery
		.Execute($iRecordsAffected)
	EndWith
	If @error Then
		Return SetError(@error, @extended, $ADO_RET_FAILURE)
	Else
		Return SetError($ADO_ERR_SUCCESS, $iRecordsAffected, $ADO_RET_SUCCESS)
	EndIf
EndFunc    ;==>_ADO_Command

; #FUNCTION# ===================================================================
; Name ..........: _ADO_Connection_Close
; Description ...: Closes an open ADODB.Connection
; Syntax.........:  _ADO_Connection_Close (ByRef $oConnection)
; Parameters ....: $oConnection         - [in/out] an object. ADODB.Connection object.
; Return values .: On Success - Returns $ADO_RET_SUCCESS
;                  On Failure - Returns $ADO_RET_FAILURE and set @error to $ADO_ERR_*
; Author ........: Chris Lambert
; Modified ......: mLipok
; Remarks .......:
; Related .......:
; Link ..........;
; Example .......; no
; ==============================================================================
Func _ADO_Connection_Close(ByRef $oConnection)
	__ADO_Connection_IsValid($oConnection)
	If @error Then Return SetError(@error, @extended, $ADO_RET_FAILURE)

	$oConnection.Close
	Return SetError($ADO_ERR_SUCCESS, $ADO_EXT_DEFAULT, $ADO_RET_SUCCESS)

EndFunc    ;==>_ADO_Connection_Close

; #FUNCTION# ===================================================================
; Name ..........: _ADO_Connection_CommandTimeout
; Description ...: Sets and retrieves SQL CommandTimeout
; Syntax.........:  _ADO_Connection_CommandTimeout(ByRef $oConnection,$iTimeout)
; Parameters ....: $oConnection         - [in/out] an object. ADODB.Connection object.
;                  $iTimeout   			- The timeout period to set if left blank the current value will be retrieved
; Return values .: On Success - Returns SQL Command timeout period
;                  On Failure - Returns $ADO_RET_FAILURE and set @error to $ADO_ERR_*
; Author ........: Chris Lambert
; Modified ......: mLipok
; Remarks .......:
; Related .......:
; Link ..........;
; Example .......; no
; ==============================================================================
Func _ADO_Connection_CommandTimeout(ByRef $oConnection, $iTimeOut = Default)
	; Error handler, automatic cleanup at end of function
	Local $oADO_COM_ErrorHandler = ObjEvent("AutoIt.Error", __ADO_ComErrorHandler_InternalFunction)
	If @error Then Return SetError($ADO_ERR_COMHANDLER, @error, $ADO_RET_FAILURE)
	#forceref $oADO_COM_ErrorHandler

	__ADO_Connection_IsValid($oConnection)
	If @error Then
		Return SetError(@error, @extended, $ADO_RET_FAILURE)
	ElseIf $iTimeOut = Default Then
		Return SetError($ADO_ERR_SUCCESS, $ADO_EXT_DEFAULT, $oConnection.CommandTimeout)
	ElseIf Not IsInt($iTimeOut) Then
		Return SetError($ADO_ERR_INVALIDPARAMETERTYPE, $ADO_EXT_DEFAULT, $ADO_RET_FAILURE)
	Else
		$oConnection.CommandTimeout = $iTimeOut
		Return SetError($ADO_ERR_SUCCESS, $ADO_EXT_DEFAULT, $oConnection.CommandTimeout)
	EndIf
EndFunc    ;==>_ADO_Connection_CommandTimeout

; #FUNCTION# ===================================================================
; Name ..........: _ADO_Connection_Create
; Description ...: Creates ADODB.Connection object
; Syntax.........:  _ADO_Connection_Create()
; Parameters ....: None
; Return values .: On Success - Returns $oConnection Object
;                  On Failure - Returns $ADO_RET_FAILURE and set @error to $ADO_ERR_*
; Author ........: Chris Lambert
; Modified ......: mLipok
; Remarks .......:
; Related .......:
; Link ..........;
; Example .......; no
; ==============================================================================
Func _ADO_Connection_Create()
	Local $oConnection = ObjCreate("ADODB.Connection")
	If @error Then Return SetError($ADO_ERR_COMERROR, @error, $ADO_RET_FAILURE)

	Return SetError($ADO_ERR_SUCCESS, $ADO_EXT_DEFAULT, $oConnection)
EndFunc    ;==>_ADO_Connection_Create

; #FUNCTION# ====================================================================================================================
; Name ..........: _ADO_Connection_OpenConString
; Description ...: Open Connection based on Connection String passed to the function
; Syntax ........: _ADO_Connection_OpenConString(Byref $oConnection, $sConnectionString)
; Parameters ....: $oConnection         - [in/out] an object. ADODB.Connection object.
;                  $sConnectionString   - a string value.
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......: Description TODO
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _ADO_Connection_OpenConString(ByRef $oConnection, $sConnectionString)
	__ADO_Connection_IsValid($oConnection)
	If @error Then Return SetError(@error, @extended, $ADO_RET_FAILURE)

	$oConnection.Open($sConnectionString)
	If @error Then Return SetError($ADO_ERR_CONNECTION, @error, $ADO_RET_FAILURE)

	Return SetError($ADO_ERR_SUCCESS, $ADO_EXT_DEFAULT, $ADO_RET_SUCCESS)

EndFunc    ;==>_ADO_Connection_OpenConString

; #FUNCTION# ===================================================================
; Name ..........: _ADO_Connection_OpenJet
; Description ...: Starts a Database Connection to a Jet Database
; Syntax.........:  _ADO_Connection_OpenJet(ByRef $oConnection, $sFilePath1)
; Parameters ....: $oConnection         - [in/out] an object. ADODB.Connection object.
;                  $sFilePath1  		- Path to Jet Database file
; Return values .: On Success - Returns $ADO_RET_SUCCESS
;                  On Failure - Returns $ADO_RET_FAILURE and set @error to $ADO_ERR_*
; Author ........: Chris Lambert
; Modified ......: mLipok
; Remarks .......:
; Related .......:
; Link ..........;
; Example .......; no
; ==============================================================================
Func _ADO_Connection_OpenJet(ByRef $oConnection, $sDataBase_FileFullPath)
	__ADO_Connection_IsValid($oConnection)
	If @error Then Return SetError(@error, @extended, $ADO_RET_FAILURE)

	$oConnection.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & $sDataBase_FileFullPath & ";")
	If @error Then Return SetError($ADO_ERR_COMERROR, @error, $ADO_RET_FAILURE)

	Return SetError($ADO_ERR_SUCCESS, $ADO_EXT_DEFAULT, $ADO_RET_SUCCESS)

EndFunc    ;==>_ADO_Connection_OpenJet

; #FUNCTION# ====================================================================================================================
; Name ..........: _ADO_Connection_OpenMSSQL
; Description ...: Starts a Database Connection
; Syntax ........: _ADO_Connection_OpenMSSQL(Byref $oConnection, $sServer, $sDBName, $sUserName, $sPassword[, $sAppName = Default[,
;                  $sWSID = Default[, $bSQLAuth = True [, $bUseProviderInsteadDriver = True]]]])
; Parameters ....: $oConnection         - [in/out] an object. ADODB.Connection object.
;                  $sServer             - a string value. The server to connect to.
;                  $sDBName             - a string value. The database name to open.
;                  $sUserName           - a string value. Username for database access.
;                  $sPassword           - a string value. Password for database user.
;                  $sAppName            - [optional] a string value. Default is Default.
;                  $sWSID               - [optional] a string value. Default is Default.
;                  $bSQLAuth            - [optional] a boolean value. Default is True.
;                  $bUseProviderInsteadDriver- [optional] a boolean value. Default is True.
; Return values .: On Success - Returns $ADO_RET_SUCCESS
;                  On Failure - Returns $ADO_RET_FAILURE and set @error to $ADO_ERR_*
; Author ........: Chris Lambert
; Modified ......: mLipok
; Remarks .......:
; Related .......:
; Link ..........: https://msdn.microsoft.com/pl-pl/library/ms130822(v=sql.110).aspx
; Example .......: No
; ===============================================================================================================================
Func _ADO_Connection_OpenMSSQL(ByRef $oConnection, $sServer, $sDBName, $sUserName, $sPassword, $sAppName = Default, $sWSID = Default, $bSQLAuth = True, $bUseProviderInsteadDriver = True)
	__ADO_Connection_IsValid($oConnection)
	If @error Then Return SetError(@error, @extended, $ADO_RET_FAILURE)

	If $oConnection.State = $ADO_adStateOpen Then
		Return SetError($ADO_ERR_SUCCESS, $ADO_EXT_DEFAULT, $ADO_RET_SUCCESS)
	EndIf

	Local $sConnectionString = ''
	If $bSQLAuth = True Then
		$sConnectionString = __ADO_MSSQL_CONNECTION_STRING_SQLAuth($sServer, $sDBName, $sUserName, $sPassword, $sAppName, $bUseProviderInsteadDriver)
	Else
		$oConnection.Properties("Integrated Security").Value = "SSPI"
		$oConnection.Properties("User ID") = $sUserName
		$oConnection.Properties("Password") = $sPassword
		$sConnectionString = "DRIVER={SQL Server};SERVER=" & $sServer & ";DATABASE=" & $sDBName & ";"
		$sConnectionString = "APP=" & $sAppName & ";"
	EndIf

	If $sWSID <> Default And $sWSID <> "" Then $sConnectionString &= "WSID=" & $sWSID & ";"

	$oConnection.Open($sConnectionString)
	If @error Then Return SetError($ADO_ERR_CONNECTION, @error, $ADO_RET_FAILURE)

	Local $vSQLOpenError_state = @error
	While Sleep(10)
		If Not $vSQLOpenError_state Or $oConnection.State = $ADO_adStateOpen Then
			__ADO_EVENTS_INIT($oConnection)
			Return SetError($ADO_ERR_SUCCESS, $ADO_EXT_DEFAULT, $ADO_RET_SUCCESS)
		Else
			Return SetError($ADO_ERR_CONNECTION, $ADO_EXT_DEFAULT, $ADO_RET_FAILURE)
		EndIf
	WEnd

EndFunc    ;==>_ADO_Connection_OpenMSSQL

; #FUNCTION# ====================================================================================================================
; Name ..........: _ADO_Connection_Timeout
; Description ...: Sets and retrieves SQL ConnectionTimeout
; Syntax ........: _ADO_Connection_Timeout(Byref $oConnection[, $iTimeOut = Default])
; Parameters ....: $oConnection         - [in/out] an object. ADODB.Connection object.
;                  $iTimeOut            - [optional] an integer value. Default is Default. The timeout period to set if left blank the current value will be retrieved
; Return values .: On Success - Returns Connection timeout period
;                  On Failure - Returns $ADO_RET_FAILURE and set @error to $ADO_ERR_*
; Author ........: Chris Lambert
; Modified ......: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _ADO_Connection_Timeout(ByRef $oConnection, $iTimeOut = Default)
	__ADO_Connection_IsValid($oConnection)
	If @error Then
		Return SetError(@error, @extended, $ADO_RET_FAILURE)
	ElseIf $iTimeOut = Default Then
		Return SetError($ADO_ERR_SUCCESS, $ADO_EXT_DEFAULT, $oConnection.ConnectionTimeout)
	ElseIf Not IsInt($iTimeOut) Then
		Return SetError($ADO_ERR_INVALIDPARAMETERTYPE, $ADO_EXT_DEFAULT, $ADO_RET_FAILURE)
	Else
		$oConnection.Close
		$oConnection.ConnectionTimeout = $iTimeOut
		Return SetError($ADO_ERR_SUCCESS, $ADO_EXT_DEFAULT, $ADO_RET_SUCCESS)
	EndIf

EndFunc    ;==>_ADO_Connection_Timeout

; #FUNCTION# ====================================================================================================================
; Name ..........: _ADO_Execute
; Description ...: Executes an SQL Query
; Syntax ........: _ADO_Execute(Byref $oConnection, $sQuery[, $bReturnAsArray = False[, $bFieldNamesInFirstRow = False]])
; Parameters ....: $oConnection         - [in/out] an object. ADODB.Connection object.
;                  $sQuery              - a string value. SQL Statement to be executed.
;                  $bReturnAsArray      - [optional] a boolean value. Default is False.
;                  $bFieldNamesInFirstRow- [optional] a boolean value. Default is False.
; Return values .: On Success - Returns $oRecordset object or $aRecordsetAsArray
;                  On Failure - Returns $ADO_RET_FAILURE and set @error to $ADO_ERR_*
; Author ........: Chris Lambert
; Modified ......: mLipok
; Remarks .......:
; Related .......:
; Link ..........;
; Example .......; no
; ===============================================================================================================================
Func _ADO_Execute(ByRef $oConnection, $sQuery, $bReturnAsArray = False, $bFieldNamesInFirstRow = False)
	__ADO_Connection_IsValid($oConnection)
	If @error Then
		Return SetError(@error, @extended, $ADO_RET_FAILURE)
	ElseIf Not IsString($sQuery) Then
		Return SetError($ADO_ERR_INVALIDPARAMETERTYPE, $ADO_EXT_PARAM2, $ADO_RET_FAILURE)
	ElseIf $sQuery = '' Then
		Return SetError($ADO_ERR_INVALIDPARAMETERVALUE, $ADO_EXT_PARAM2, $ADO_RET_FAILURE)
	ElseIf Not IsBool($bReturnAsArray) Then
		Return SetError($ADO_ERR_INVALIDPARAMETERTYPE, $ADO_EXT_PARAM3, $ADO_RET_FAILURE)
	ElseIf Not IsBool($bFieldNamesInFirstRow) Then
		Return SetError($ADO_ERR_INVALIDPARAMETERTYPE, $ADO_EXT_PARAM4, $ADO_RET_FAILURE)
	EndIf

	Local $oRecordset = $oConnection.Execute($sQuery)
	If @error Then Return SetError($ADO_ERR_COMERROR, @error, $ADO_RET_FAILURE)

	If $bReturnAsArray Then
		Local $aRecordsetAsArray = _ADO_Recordset_ToArray($oRecordset, $bFieldNamesInFirstRow)
		Return SetError(@error, @extended, $aRecordsetAsArray)
	EndIf

	Return SetError($ADO_ERR_SUCCESS, $ADO_EXT_DEFAULT, $oRecordset)

EndFunc    ;==>_ADO_Execute

; #FUNCTION# ====================================================================================================================
; Name ..........: _ADO_MSSQL_GetDriverVersion
; Description ...: check for newer DRIVER parameter for CONNECTIONSTRING
; Syntax ........: _ADO_MSSQL_GetDriverVersion()
; Parameters ....: none.
; Return values .: $s_ADO_MSSQL_GetDriverVersion
; Author ........: mLipok
; Modified ......:
; Remarks .......: TODO - description
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _ADO_MSSQL_GetDriverVersion()
	Local Static $s_ADO_MSSQL_GetDriverVersion = Default
	If $s_ADO_MSSQL_GetDriverVersion = Default Then
;~ 		Local  $sSQL_NCLI_2014 = RegRead('HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Microsoft SQL Server Native Client 11.0\CurrentVersion', 'Version') ; For SQL Server 2008/SQL Server 2008 R2
		Local $sSQL_NCLI_2012 = RegRead('HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Microsoft SQL Server Native Client 11.0\CurrentVersion', 'Version') ; For SQL Server 2008/SQL Server 2008 R2
		Local $sSQL_NCLI_2008 = RegRead('HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Microsoft SQL Server Native Client 10.0\CurrentVersion', 'Version') ; For SQL Server 2008/SQL Server 2008 R2
		Local $sSQL_NCLI_2005 = RegRead('HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Microsoft SQL Native Client\CurrentVersion', 'Version') ; For SQL Server 2005
		Select
;~ 			Case  $sSQL_NCLI_2014 <> ''
;~ 				$s_ADO_MSSQL_GetDriverVersion = 'SQL Server Native Client 11.0'
			Case $sSQL_NCLI_2012 <> ''
				$s_ADO_MSSQL_GetDriverVersion = 'SQL Server Native Client 11.0'
			Case $sSQL_NCLI_2008 <> ''
				$s_ADO_MSSQL_GetDriverVersion = 'SQL Server Native Client 10.0'
			Case $sSQL_NCLI_2005 <> ''
				$s_ADO_MSSQL_GetDriverVersion = 'SQL Native Client'
			Case Else
				$s_ADO_MSSQL_GetDriverVersion = 'SQL Server'
		EndSelect
	EndIf
	Return $s_ADO_MSSQL_GetDriverVersion

EndFunc    ;==>_ADO_MSSQL_GetDriverVersion

; #FUNCTION# ====================================================================================================================
; Name ..........: _ADO_MSSQL_GetProviderVersion
; Description ...: check for newer PROVIDER parameter for CONNECTIONSTRING
; Syntax ........: _ADO_MSSQL_GetProviderVersion()
; Parameters ....: none.
; Return values .: $s_ADO_MSSQL_GetProviderVersion
; Author ........: mLipok
; Modified ......:
; Remarks .......: TODO - description
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _ADO_MSSQL_GetProviderVersion()
	Local Static $s_ADO_MSSQL_GetProviderVersion = Default
	If $s_ADO_MSSQL_GetProviderVersion = Default Then
;~ 		Local  $sSQL_NCLI_2014 = RegRead('HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Microsoft SQL Server Native Client 11.0\CurrentVersion', 'Version') ; For SQL Server 2008/SQL Server 2008 R2
		Local $sSQL_NCLI_2012 = RegRead('HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Microsoft SQL Server Native Client 11.0\CurrentVersion', 'Version') ; For SQL Server 2008/SQL Server 2008 R2
		Local $sSQL_NCLI_2008 = RegRead('HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Microsoft SQL Server Native Client 10.0\CurrentVersion', 'Version') ; For SQL Server 2008/SQL Server 2008 R2
		Local $sSQL_NCLI_2005 = RegRead('HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Microsoft SQL Native Client\CurrentVersion', 'Version') ; For SQL Server 2005
		Select
;~ 			Case  $sSQL_NCLI_2014 <> ''
;~ 				$s_ADO_MSSQL_GetProviderVersion = 'SQL Server Native Client 11.0'
			Case $sSQL_NCLI_2012 <> ''
				$s_ADO_MSSQL_GetProviderVersion = 'SQLNCLI11'
			Case $sSQL_NCLI_2008 <> ''
				$s_ADO_MSSQL_GetProviderVersion = 'SQLNCLI10'
			Case $sSQL_NCLI_2005 <> ''
				$s_ADO_MSSQL_GetProviderVersion = 'SQLNCLI'
			Case Else
				$s_ADO_MSSQL_GetProviderVersion = 'sqloledb'
		EndSelect
	EndIf
	Return $s_ADO_MSSQL_GetProviderVersion

EndFunc    ;==>_ADO_MSSQL_GetProviderVersion

; #FUNCTION# ===================================================================
; Name ..........: _ADO_Recordset_Create
; Description ...: Creates ADODB.Recordset object
; Syntax.........:  _ADO_Recordset_Create()
; Parameters ....: None
; Return values .: On Success - Returns $oRecordset Object
;                  On Failure - Returns $ADO_RET_FAILURE and set @error to $ADO_ERR_*
; Author ........: mLipok
; Modified ......: mLipok
; Remarks .......:
; Related .......:
; Link ..........;
; Example .......; no
; ==============================================================================
Func _ADO_Recordset_Create()
	; Error handler, automatic cleanup at end of function
	Local $oADO_COM_ErrorHandler = ObjEvent("AutoIt.Error", __ADO_ComErrorHandler_InternalFunction)
	If @error Then Return SetError($ADO_ERR_COMHANDLER, @error, $ADO_RET_FAILURE)
	#forceref $oADO_COM_ErrorHandler

	Local $oRecordset = ObjCreate("ADODB.Recordset")
	If @error Then Return SetError($ADO_ERR_COMERROR, @error, $ADO_RET_FAILURE)

	Return SetError($ADO_ERR_SUCCESS, $ADO_EXT_DEFAULT, $oRecordset)
EndFunc    ;==>_ADO_Recordset_Create

; #FUNCTION# ====================================================================================================================
; Name ..........: _ADO_Version
; Description ...:
; Syntax ........: _ADO_Version([ByRef $oConnection])
; Parameters ....: $oConnection         - [in/out] an object. ADODB.Connection object.
; Return values .: None
; Author ........: Chris Lambert
; Modified ......: mLipok
; Remarks .......: TODO - description
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _ADO_Version(ByRef $oConnection)
	; Error handler, automatic cleanup at end of function
	Local $oADO_COM_ErrorHandler = ObjEvent("AutoIt.Error", __ADO_ComErrorHandler_InternalFunction)
	If @error Then Return SetError($ADO_ERR_COMHANDLER, @error, $ADO_RET_FAILURE)
	#forceref $oADO_COM_ErrorHandler

	__ADO_Connection_IsValid($oConnection)
	If @error Then
		Return SetError(@error, @extended, $ADO_RET_FAILURE)
	Else
		Return SetError($ADO_ERR_SUCCESS, $ADO_EXT_DEFAULT, $oConnection.Version)
	EndIf
EndFunc    ;==>_ADO_Version
#EndRegion  ADO.au3 - Functions - Connection & Management

#Region ADO.au3 - Functions - ADDON - COM ERROR HANDLER
; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __ADO_ComErrorHandler_InternalFunction
; Description ...:
; Syntax ........: __ADO_ComErrorHandler_InternalFunction($oCOMError)
; Parameters ....: $oCOMError           - an object.
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __ADO_ComErrorHandler_InternalFunction($oCOMError)
	; Do nothing special, just check @error after suspect functions.
	#forceref $oCOMError
	Local $sUserFunction = _ADO_COMErrorHandler_UserFunction()
	If IsFunc($sUserFunction) Then $sUserFunction($oCOMError)
EndFunc    ;==>__ADO_ComErrorHandler_InternalFunction

; #FUNCTION# ===================================================================
; Name ..........: _ADO_COMErrorHandler
; Description ...: Autoit COM Error handler function
; Syntax ........: _ADO_COMErrorHandler()
; Parameters ....: None.
; Return values .:
; Author ........: Chris Lambert
; Modified ......: mLipok
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: no
; ================================================================================
Func _ADO_COMErrorHandler($oADO_Error)
	; Error Object
	; https://msdn.microsoft.com/en-us/library/windows/desktop/ms677507(v=vs.85).aspx
	; Error Object Properties, Methods, and Events
	; https://msdn.microsoft.com/en-us/library/windows/desktop/ms678396(v=vs.85).aspx

	Local $HexNumber = Hex($oADO_Error.number, 8)
	Local $g__sSQL_ComErrorDescription = ''
	$g__sSQL_ComErrorDescription &= "ADO.au3 (" & $oADO_Error.scriptline & ") : ==> COM Error intercepted !" & @CRLF
	$g__sSQL_ComErrorDescription &= "$oADO_Error.description is: " & @TAB & $oADO_Error.description & @CRLF
	$g__sSQL_ComErrorDescription &= "$oADO_Error.windescription: " & @TAB & $oADO_Error.windescription & @CRLF
	$g__sSQL_ComErrorDescription &= "$oADO_Error.number is: " & @TAB & $HexNumber & @CRLF
	$g__sSQL_ComErrorDescription &= "$oADO_Error.lastdllerror is: " & @TAB & $oADO_Error.lastdllerror & @CRLF
	$g__sSQL_ComErrorDescription &= "$oADO_Error.scriptline is: " & @TAB & $oADO_Error.scriptline & @CRLF

	; Source Property (ADO Error)
	; https://msdn.microsoft.com/en-us/library/windows/desktop/ms675830(v=vs.85).aspx
	$g__sSQL_ComErrorDescription &= "$oADO_Error.source is: " & @TAB & $oADO_Error.source & @CRLF
	$g__sSQL_ComErrorDescription &= "$oADO_Error.helpfile is: " & @TAB & $oADO_Error.helpfile & @CRLF
	$g__sSQL_ComErrorDescription &= "$oADO_Error.helpcontext is: " & @TAB & $oADO_Error.helpcontext & @CRLF

	#CS
		; NativeError Property (ADO)
		; https://msdn.microsoft.com/en-us/library/windows/desktop/ms678049(v=vs.85).aspx
		$g__sSQL_ComErrorDescription &= "$oADO_Error.NativeError is: " & @TAB & $oADO_Error.NativeError & @CRLF

		; SQLState Property
		; https://msdn.microsoft.com/en-us/library/windows/desktop/ms681570(v=vs.85).aspx
		$g__sSQL_ComErrorDescription &= "$oADO_Error.SQLState is: " & @TAB & $oADO_Error.SQLState & @CRLF

	#CE

	ConsoleWrite("###############################" & @CRLF & $g__sSQL_ComErrorDescription & "###############################" & @CRLF)
	SetError($ADO_ERR_GENERAL, $ADO_EXT_DEFAULT, $g__sSQL_ComErrorDescription)
EndFunc    ;==>_ADO_COMErrorHandler

; #FUNCTION# ====================================================================================================================
; Name ..........: _ADO_COMErrorHandler_UserFunction
; Description ...: Set up user function to get COM Error Handler outside ADO.au3 UDF
; Syntax ........: _ADO_COMErrorHandler_UserFunction([$fnUserFunction = Default])
; Parameters ....: $fnUserFunction      - [optional] a floating point value. Default is Default.
; Return values .: On Success - $fnUserFunction_Static
;                  On Failure - Returns $ADO_RET_FAILURE and set @error to $ADO_ERR_*
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _ADO_COMErrorHandler_UserFunction($fnUserFunction = Default)
	; in case when user do not set his own function UDF must use internal function to avoid AutoItError
	Local Static $fnUserFunction_Static = ''

	If $fnUserFunction = Default Then
		; just return stored static variable
		Return $fnUserFunction_Static
	ElseIf IsFunc($fnUserFunction) Then
		; set and return static variable
		$fnUserFunction_Static = $fnUserFunction
		Return $fnUserFunction_Static
	Else
		; reset static variable
		$fnUserFunction_Static = ''
		Return SetError($ADO_ERR_INVALIDPARAMETERTYPE, $ADO_EXT_DEFAULT, $fnUserFunction_Static)
	EndIf
EndFunc    ;==>_ADO_COMErrorHandler_UserFunction
#EndRegion  ADO.au3 - Functions - ADDON - COM ERROR HANDLER

#Region ADO.au3 - Functions - ADDON - COM EVENT HANDLER
; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __ADO_EVENT__BeginTransComplete
; Description ...: BeginTransComplete is called after the BeginTrans operation
; Syntax ........: __ADO_EVENT__BeginTransComplete($iTransactionLevel, Byref $oError, $i_adStatus, Byref $oConnection)
; Parameters ....: $iTransactionLevel   - an integer value. A Long value that contains the new transaction level of the BeginTrans that caused this event.
;                  $oError              - [in/out] an object. An Error object. It describes the error that occurred if the value of EventStatusEnum is adStatusErrorsOccurred; otherwise it is not set.
;                  $i_adStatus          - an integer value. An EventStatusEnum status value. When any of these events is called, this parameter is set to adStatusOK if the operation that caused the event was successful, or to adStatusErrorsOccurred if the operation failed.
;											These events can prevent subsequent notifications by setting this parameter to adStatusUnwantedEvent before the event returns.
;                  $oConnection         - [in/out] an object. The Connection object for which this event occurred.
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: https://msdn.microsoft.com/en-us/library/windows/desktop/ms681493%28v=vs.85%29.aspx
; Example .......: No
; ===============================================================================================================================
Func __ADO_EVENT__BeginTransComplete($iTransactionLevel, ByRef $oError, $i_adStatus, ByRef $oConnection)
;~ 	https://msdn.microsoft.com/en-us/library/ms681493(v=vs.85).aspx
	If Not _ADO_EVENTS_SetUp() Then Return

	__ADO_ConsoleWrite_Blue(" ADO EVENT fired function: __ADO_EVENT__BeginTransComplete:")
	__ADO_ConsoleWrite_Blue("   $iTransactionLevel=" & $iTransactionLevel)
	__ADO_ConsoleWrite_Blue("   $i_adStatus=" & $i_adStatus)
	#forceref $oError, $oConnection
EndFunc    ;==>__ADO_EVENT__BeginTransComplete

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __ADO_EVENT__CommitTransComplete
; Description ...: CommitTransComplete is called after the CommitTrans operation.
; Syntax ........: __ADO_EVENT__CommitTransComplete(Byref $oError, $i_adStatus, Byref $oConnection)
; Parameters ....: $oError              - [in/out] an object.
;                  $i_adStatus          - an integer value.
;                  $oConnection         - [in/out] an object.
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: https://msdn.microsoft.com/en-us/library/windows/desktop/ms681493%28v=vs.85%29.aspx
; Example .......: No
; ===============================================================================================================================
Func __ADO_EVENT__CommitTransComplete(ByRef $oError, $i_adStatus, ByRef $oConnection)
	If Not _ADO_EVENTS_SetUp() Then Return
	__ADO_ConsoleWrite_Blue(" ADO EVENT fired function: __ADO_EVENT__CommitTransComplete:")
	__ADO_ConsoleWrite_Blue("   $i_adStatus=" & $i_adStatus)
	If $i_adStatus = $ADO_adStatusErrorsOccurred Then
		__ADO_ConsoleWrite_Red("   $i_adStatus=$ADO_adStatusErrorsOccurred=" & $i_adStatus)
		__ADO_ConsoleWrite_Red("   STARTING:  $oConnection.RollbackTrans")
		$oConnection.RollbackTrans
	EndIf
	#forceref $oError, $oConnection
EndFunc    ;==>__ADO_EVENT__CommitTransComplete

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __ADO_EVENT__ConnectComplete
; Description ...: ConnectComplete Events (ADO)
; Syntax ........: __ADO_EVENT__ConnectComplete(Byref $oError, $i_adStatus, Byref $oConnection)
; Parameters ....: $oError              - [in/out] an object.
;                  $i_adStatus          - an integer value.
;                  $oConnection         - [in/out] an object.
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: https://msdn.microsoft.com/en-us/library/windows/desktop/ms676126(v=vs.85).aspx
; Example .......: No
; ===============================================================================================================================
Func __ADO_EVENT__ConnectComplete(ByRef $oError, $i_adStatus, ByRef $oConnection)
	If Not _ADO_EVENTS_SetUp() Then Return
	__ADO_ConsoleWrite_Blue(" ADO EVENT fired function: __ADO_EVENT__ConnectComplete:")
	__ADO_ConsoleWrite_Blue("   $i_adStatus=" & $i_adStatus)
	#forceref $oError, $i_adStatus, $oConnection
EndFunc    ;==>__ADO_EVENT__ConnectComplete

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __ADO_EVENT__Disconnect
; Description ...: Disconnect Events (ADO)
; Syntax ........: __ADO_EVENT__Disconnect($i_adStatus, Byref $oConnection)
; Parameters ....: $i_adStatus          - an integer value.
;                  $oConnection         - [in/out] an object.
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: https://msdn.microsoft.com/en-us/library/windows/desktop/ms676126(v=vs.85).aspx
; Example .......: No
; ===============================================================================================================================
Func __ADO_EVENT__Disconnect($i_adStatus, ByRef $oConnection)
	If Not _ADO_EVENTS_SetUp() Then Return
	__ADO_ConsoleWrite_Blue(" ADO EVENT fired function: __ADO_EVENT__Disconnect:")
	__ADO_ConsoleWrite_Blue("   $i_adStatus=" & $i_adStatus)
	#forceref $i_adStatus, $oConnection
EndFunc    ;==>__ADO_EVENT__Disconnect

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __ADO_EVENT__FetchComplete
; Description ...: FetchComplete Event (ADO)
; Syntax ........: __ADO_EVENT__FetchComplete(Byref $oError, $i_adStatus, Byref $oRecordset)
; Parameters ....: $oError              - [in/out] an object.
;                  $i_adStatus          - an integer value.
;                  $oRecordset          - [in/out] an object.
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: https://msdn.microsoft.com/en-us/library/windows/desktop/ms677512(v=vs.85).aspx
; Example .......: No
; ===============================================================================================================================
Func __ADO_EVENT__FetchComplete(ByRef $oError, $i_adStatus, ByRef $oRecordset)
	If Not _ADO_EVENTS_SetUp() Then Return
	__ADO_ConsoleWrite_Blue(" ADO EVENT fired function: __ADO_EVENT__FEtchComplete:")
	__ADO_ConsoleWrite_Blue("   $i_adStatus=" & $i_adStatus)
	#forceref $oError, $oRecordset
EndFunc    ;==>__ADO_EVENT__FetchComplete

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __ADO_EVENT__FetchProgress
; Description ...: FetchProgress Event (ADO)
; Syntax ........: __ADO_EVENT__FetchProgress($iProgress, $iMaxProgress, $i_adStatus, Byref $oRecordset)
; Parameters ....: $iProgress           - an integer value.
;                  $iMaxProgress        - an integer value.
;                  $i_adStatus          - an integer value.
;                  $oRecordset          - [in/out] an object.
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: https://msdn.microsoft.com/en-us/library/windows/desktop/ms675535(v=vs.85).aspx
; Example .......: No
; ===============================================================================================================================
Func __ADO_EVENT__FetchProgress($iProgress, $iMaxProgress, $i_adStatus, ByRef $oRecordset)
	If Not _ADO_EVENTS_SetUp() Then Return
	__ADO_ConsoleWrite_Blue(" ADO EVENT fired function: __ADO_EVENT__FetchProgress:")
	__ADO_ConsoleWrite_Blue("   $iProgress=" & $iProgress)
	__ADO_ConsoleWrite_Blue("   $iMaxProgress=" & $iMaxProgress)
	__ADO_ConsoleWrite_Blue("   $i_adStatus=" & $i_adStatus)
	If IsFunc($__g_fnFetchProgress) Then
		$__g_fnFetchProgress($iProgress, $iMaxProgress, $i_adStatus, $oRecordset)
	EndIf
	#forceref $oRecordset
EndFunc    ;==>__ADO_EVENT__FetchProgress

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __ADO_EVENT__InfoMessage
; Description ...:
; Syntax ........: __ADO_EVENT__InfoMessage(Byref $oError, $i_adStatus, Byref $oConnection)
; Parameters ....: $oError              - [in/out] an object.
;                  $i_adStatus          - an integer value.
;                  $oConnection         - [in/out] an object.
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: https://msdn.microsoft.com/en-us/library/windows/desktop/ms675859(v=vs.85).aspx
; Example .......: No
; ===============================================================================================================================
Func __ADO_EVENT__InfoMessage(ByRef $oError, $i_adStatus, ByRef $oConnection)
	If Not _ADO_EVENTS_SetUp() Then Return
	__ADO_ConsoleWrite_Blue(" ADO EVENT fired function: __ADO_EVENT__InfoMessage:")
	__ADO_ConsoleWrite_Blue("   $i_adStatus=" & $i_adStatus)
	#forceref $oError, $i_adStatus, $oConnection
EndFunc    ;==>__ADO_EVENT__InfoMessage

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __ADO_EVENT__RollbackTransComplete
; Description ...: RollbackTransComplete is called after the RollbackTrans operation.
; Syntax ........: __ADO_EVENT__RollbackTransComplete(Byref $oError, $i_adStatus, Byref $oConnection)
; Parameters ....: $oError              - [in/out] an object.
;                  $i_adStatus          - an integer value.
;                  $oConnection         - [in/out] an object.
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: https://msdn.microsoft.com/en-us/library/windows/desktop/ms681493%28v=vs.85%29.aspx
; Example .......: No
; ===============================================================================================================================
Func __ADO_EVENT__RollbackTransComplete(ByRef $oError, $i_adStatus, ByRef $oConnection)
	If Not _ADO_EVENTS_SetUp() Then Return
	__ADO_ConsoleWrite_Blue(" ADO EVENT fired function: __ADO_EVENT__RollbackTransComplete:")
	__ADO_ConsoleWrite_Blue("   $i_adStatus=" & $i_adStatus)
	#forceref $oError, $oConnection
EndFunc    ;==>__ADO_EVENT__RollbackTransComplete

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __ADO_EVENT__WillConnect
; Description ...: WillConnect Event (ADO)
; Syntax ........: __ADO_EVENT__WillConnect($sConnection_String, $sUserID, $sPassword, $iOptions, $i_adStatus, Byref $oConnection)
; Parameters ....: $sConnection_String   - a string value.
;                  $sUserID             - a string value.
;                  $sPassword           - a string value.
;                  $iOptions            - an integer value.
;                  $i_adStatus          - an integer value.
;                  $oConnection         - [in/out] an object.
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: https://msdn.microsoft.com/en-us/library/windows/desktop/ms680962(v=vs.85).aspx
; Example .......: No
; ===============================================================================================================================
Func __ADO_EVENT__WillConnect($sConnection_String, $sUserID, $sPassword, $iOptions, $i_adStatus, ByRef $oConnection)
	If Not _ADO_EVENTS_SetUp() Then Return
	__ADO_ConsoleWrite_Blue(" ADO EVENT fired function: __ADO_EVENT__WillConnect:")
	__ADO_ConsoleWrite_Blue("   $sConnection_String=" & $sConnection_String)
	__ADO_ConsoleWrite_Blue("   $sUserID=" & $sUserID)
	__ADO_ConsoleWrite_Blue("   $sPassword=" & $sPassword)
	__ADO_ConsoleWrite_Blue("   $iOptions=" & $iOptions)
	__ADO_ConsoleWrite_Blue("   $i_adStatus=" & $i_adStatus)
	#forceref $oConnection
EndFunc    ;==>__ADO_EVENT__WillConnect

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __ADO_EVENT__WillExecute
; Description ...: WillExecute Event (ADO)
; Syntax ........: __ADO_EVENT__WillExecute($sSource, $iCursorType, $iLockType, $iOptions, $i_adStatus, Byref $oCommand,
;                  Byref $oRecordset, Byref $oConnection)
; Parameters ....: $sSource             - a string value.
;                  $iCursorType         - an integer value.
;                  $iLockType           - an integer value.
;                  $iOptions            - an integer value.
;                  $i_adStatus          - an integer value.
;                  $oCommand            - [in/out] an object.
;                  $oRecordset          - [in/out] an object.
;                  $oConnection         - [in/out] an object.
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: https://msdn.microsoft.com/en-us/library/windows/desktop/ms680993(v=vs.85).aspx
; Example .......: No
; ===============================================================================================================================
Func __ADO_EVENT__WillExecute($sSource, $iCursorType, $iLockType, $iOptions, $i_adStatus, ByRef $oCommand, ByRef $oRecordset, ByRef $oConnection)
	If Not _ADO_EVENTS_SetUp() Then Return
	__ADO_ConsoleWrite_Blue(" ADO EVENT fired function: __ADO_EVENT__WillExecute:")
	__ADO_ConsoleWrite_Blue("   $sSource=" & StringRegExpReplace($sSource, '\R', ' '))
	__ADO_ConsoleWrite_Blue("   $iCursorType=" & $iCursorType)
	__ADO_ConsoleWrite_Blue("   $iLockType=" & $iLockType)
	__ADO_ConsoleWrite_Blue("   $iOptions=" & $iOptions)
	__ADO_ConsoleWrite_Blue("   $i_adStatus=" & $i_adStatus)
	#forceref $oCommand, $oRecordset, $oConnection
EndFunc    ;==>__ADO_EVENT__WillExecute

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __ADO_EVENTS_INIT
; Description ...: Function to initialize ADO EVENTs handling
; Syntax ........: __ADO_EVENTS_INIT(Byref $oConnection)
; Parameters ....: $oConnection         - [in/out] an object. ADODB.Connection object.
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __ADO_EVENTS_INIT(ByRef $oConnection)
	__ADO_Connection_IsValid($oConnection)
	If @error Then
		Return SetError(@error, @extended, $ADO_RET_FAILURE)
	EndIf
	Local Static $oADO_EventHandler = ''
	If $oADO_EventHandler = '' Then
		$oADO_EventHandler = ObjEvent($oConnection, "__ADO_EVENT__", "ConnectionEvents") ; @TODO check with #Au3Stripper_Parameters=/TL /debug
	Else
		$oADO_EventHandler = ''
	EndIf
EndFunc    ;==>__ADO_EVENTS_INIT

; #FUNCTION# ====================================================================================================================
; Name ..........: _ADO_EVENTS_SetUp
; Description ...: Enable/Disable/Get - ADO EVENTs handling status
; Syntax ........: _ADO_EVENTS_SetUp([$bInitializeEvents = Default])
; Parameters ....: $bInitializeEvents   - [optional] a boolean value. Default is Default.
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _ADO_EVENTS_SetUp($bInitializeEvents = Default)
	Local Static $bInitializeEvents_static = True

	If $bInitializeEvents = Default Then
		Return SetError($ADO_ERR_SUCCESS, $ADO_EXT_DEFAULT, $bInitializeEvents_static)
	ElseIf IsBool($bInitializeEvents) Then
		$bInitializeEvents_static = $bInitializeEvents
		Return SetError($ADO_ERR_SUCCESS, 1, $bInitializeEvents_static)
	Else
		Return SetError($ADO_ERR_INVALIDPARAMETERTYPE, 2, $bInitializeEvents_static)
	EndIf
EndFunc    ;==>_ADO_EVENTS_SetUp
#EndRegion  ADO.au3 - Functions - ADDON - COM EVENT HANDLER

#Region ADO.au3 - Functions - MISC
; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __ADO_ConsoleWrite_Blue
; Description ...:
; Syntax ........: __ADO_ConsoleWrite_Blue($sText)
; Parameters ....: $sText               - a string value.
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __ADO_ConsoleWrite_Blue($sText)
	ConsoleWrite(BinaryToString(StringToBinary('>>' & $sText & @CRLF, 4), 1))
EndFunc    ;==>__ADO_ConsoleWrite_Blue

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __ADO_ConsoleWrite_Red
; Description ...:
; Syntax ........: __ADO_ConsoleWrite_Red($sText)
; Parameters ....: $sText               - a string value.
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __ADO_ConsoleWrite_Red($sText)
	ConsoleWrite(BinaryToString(StringToBinary('!!!!!!!!!' & $sText & @CRLF, 4), 1))
EndFunc    ;==>__ADO_ConsoleWrite_Red

; #FUNCTION# ====================================================================================================================
; Name ..........: _Au3Date_to_SQLDate
; Description ...:
; Syntax ........: _Au3Date_to_SQLDate($sAu3Date)
; Parameters ....: $sAu3Date            - a string value.
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _Au3Date_to_SQLDate($sAu3Date)
	; IN:  1970/01/01 12:30:15
	; OUT: 1970-01-01T12:30:15.000

	If Not _DateIsValid($sAu3Date) Then
		Return SetError($ADO_ERR_GENERAL, $ADO_EXT_PARAM1, $ADO_RET_FAILURE)
	EndIf

	; if only date then add time
	If StringRegExpReplace($sAu3Date, '(\d{4}\/\d{2}\/\d{2})', '') = '' Then $sAu3Date &= ' 00:00:00'
	; replace "/" to "-"    and add miliseconds
	Local $sSQLDate = StringReplace($sAu3Date, '/', '-') & '.000'
	; change the space (separator for date and time) for SQL equivalent T char
	$sSQLDate = StringReplace($sSQLDate, ' ', 'T')

	Return SetError($ADO_ERR_SUCCESS, $ADO_EXT_DEFAULT, $sSQLDate)
EndFunc    ;==>_Au3Date_to_SQLDate

; #FUNCTION# ====================================================================================================================
; Name ..........: _SQLDate_to_Au3Date
; Description ...:
; Syntax ........: _SQLDate_to_Au3Date($sDate[, $bOnlyYMD = False])
; Parameters ....: $sDate               - a string value.
;                  $bOnlyYMD            - [optional] a boolean value. Default is False.
; Return values .: Au3Date
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _SQLDate_to_Au3Date($sDate, $bOnlyYMD = False)
	Local $sParam = ($bOnlyYMD = True) ? '$1\/$2\/$3' : '$1\/$2\/$3\ $4:$5:$6'
	Return StringRegExpReplace($sDate, '(\d{4})(\d{2})(\d{2})(\d{2})(\d{2})(\d{2})', $sParam)

EndFunc    ;==>_SQLDate_to_Au3Date
#EndRegion  ADO.au3 - Functions - MISC

#Region ADO.au3 - Functions - Getting Data
; #FUNCTION# ====================================================================================================================
; Name ..........: _ADO_ExecuteQueryToArray
; Description ...: Passes a Array Containing result of Executed Query
; Syntax ........: _ADO_ExecuteQueryToArray(Byref $oConnection, $sQuery[, $bFieldNamesInFirstRow = True])
; Parameters ....: $oConnection         - [in/out] an object. ADODB.Connection object.
;                  $sQuery              - a string value.
;                  $bFieldNamesInFirstRow- [optional] a boolean value. Default is True.
; Return values .: On Success - Returns $aResult
;                  On Failure - Returns $ADO_RET_FAILURE and set @error to $ADO_ERR_*
; Author ........: Chris Lambert
; Modified ......: Stephen Podhajecki (eltorro), mLipok
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _ADO_ExecuteQueryToArray(ByRef $oConnection, $sQuery, $bFieldNamesInFirstRow)
	Local $oRecordset = _ADO_Execute($oConnection, $sQuery)
	If @error Then Return SetError(@error, @extended, $ADO_RET_FAILURE)

	Local $aResult = _ADO_Recordset_ToArray($oRecordset, $bFieldNamesInFirstRow)
	Return SetError(@error, @extended, $aResult)

EndFunc    ;==>_ADO_ExecuteQueryToArray

; #FUNCTION# ===================================================================
; Name ..........: _SQL_FetchData()
; Description ...: Fetches 1 Row of Data from an $oRecordset object returned from _ADO_Execute()
; Syntax.........:  _SQL_FetchData($oRecordset,ByRef $aRow)
; Parameters ....: $oRecordset    - object passed out by _ADO_Execute()
; Return values .: On Success - Returns $aRow
;                  On Failure - Returns $ADO_RET_FAILURE and set @error to $ADO_ERR_*
; Author ........: Chris Lambert
; Modified ......: mLipok
; Remarks .......:
; Related .......:
; Link ..........;
; Example .......; no
; ==============================================================================
Func _SQL_FetchData($oRecordset, $sSQL_Delim = "¬&~")
	; Error handler, automatic cleanup at end of function
	Local $oADO_COM_ErrorHandler = ObjEvent("AutoIt.Error", __ADO_ComErrorHandler_InternalFunction)
	If @error Then Return SetError($ADO_ERR_COMHANDLER, @error, $ADO_RET_FAILURE)
	#forceref $oADO_COM_ErrorHandler

	__ADO_Recordset_IsNotEmpty($oRecordset)
	If @error Then
		Return SetError(@error, @extended, $ADO_RET_FAILURE)
	ElseIf $oRecordset.EOF Then
		Return SetError($ADO_ERR_GENERAL, $ADO_EXT_DEFAULT, $ADO_RET_FAILURE)
	EndIf

	Local $sCurrentRow = ""
	Local $oFields_coll = $oRecordset.Fields
	If @error Then Return SetError($ADO_ERR_COMERROR, @error, $ADO_RET_FAILURE)

	; get the Row content
	For $oField_enum In $oFields_coll
		$sCurrentRow &= $oField_enum.Value & $sSQL_Delim
	Next

	; Move to next row
	$oRecordset.MoveNext

	Local $iDelimLen = StringLen($sSQL_Delim)
	If StringRight($sCurrentRow, $iDelimLen) = $sSQL_Delim Then $sCurrentRow = StringTrimRight($sCurrentRow, $iDelimLen)

	Local $aRow = StringSplit($sCurrentRow, $sSQL_Delim, $STR_ENTIRESPLIT + $STR_NOCOUNT)
	Return SetError($ADO_ERR_SUCCESS, $ADO_EXT_DEFAULT, $aRow)

EndFunc    ;==>_SQL_FetchData

; #FUNCTION# ===================================================================
; Name ..........: _SQL_FetchNames()
; Description ...: Read out the Tablenames of a _SQL_Query() based query
; Syntax.........:  _SQL_FetchNames($oRecordset, ByRef $aNames)
; Parameters ....: $oRecordset    - object passed out by _ADO_Execute()
; Return values .: On Success - Returns $aFieldNames
;                  On Failure - Returns $ADO_RET_FAILURE and set @error to $ADO_ERR_*
; Author ........: Chris Lambert
; Modified ......: mLipok
; Remarks .......:
; Related .......:
; Link ..........;
; Example .......; no
; ==============================================================================
Func _SQL_FetchNames(ByRef $oRecordset, $sSQL_Delim = "¬&~")
	; Error handler, automatic cleanup at end of function
	Local $oADO_COM_ErrorHandler = ObjEvent("AutoIt.Error", __ADO_ComErrorHandler_InternalFunction)
	If @error Then Return SetError($ADO_ERR_COMHANDLER, @error, $ADO_RET_FAILURE)
	#forceref $oADO_COM_ErrorHandler

	__ADO_Recordset_IsValid($oRecordset)
	If @error Then
		Return SetError(@error, @extended, $ADO_RET_FAILURE)
	EndIf

	Local $sFieldNames = ""
	Local $oFields_coll = $oRecordset.Fields
	If @error Then Return SetError($ADO_ERR_COMERROR, @error, $ADO_RET_FAILURE)

	; get the column names and put into 0 array element
	For $oField_enum In $oFields_coll
		$sFieldNames &= $oField_enum.Name & $sSQL_Delim
	Next

	If $sFieldNames <> "" Then $sFieldNames = StringTrimRight($sFieldNames, StringLen($sSQL_Delim))

	Local $aFieldNames = StringSplit($sFieldNames, $sSQL_Delim, $STR_ENTIRESPLIT + $STR_NOCOUNT)
	Return SetError($ADO_ERR_SUCCESS, $ADO_EXT_DEFAULT, $aFieldNames)

EndFunc    ;==>_SQL_FetchNames

; #FUNCTION# ===================================================================
; Name ..........: _SQL_GetTableName()
; Description ...: Get Table List Of Open Data Base
; Syntax.........:  _SQL_GetTableName(ByRef $oConnection[,$sType = "TABLE" ])
; Parameters ....: $oConnection         - [in/out] an object. ADODB.Connection object.
;                   $sType 				- Table Type  "TABLE" (Default), "VIEW", "SYSTEM TABLE", "ACCESS TABLE"
;                   					$sType = "*" - Return All Tables in a  Array2D  $aTable[n][2]  $aTable[n][0] = Table Name $aTable[n][1] = Table  Type
; Return values .: On Success - Returns a 1D Array Of Table Names / 2D Array is $sType = "*"
;                  On Failure - Returns $ADO_RET_FAILURE and set @error to $ADO_ERR_*
; Author ........: Elias Assad Neto
; Modified ......: ChrisL, mLipok
; Remarks .......:
; Related .......:
; Link ..........;
; Example .......; no
; ==============================================================================
Func _SQL_GetTableName(ByRef $oConnection, $sType = "TABLE")
	; Error handler, automatic cleanup at end of function
	Local $oADO_COM_ErrorHandler = ObjEvent("AutoIt.Error", __ADO_ComErrorHandler_InternalFunction)
	If @error Then Return SetError($ADO_ERR_COMHANDLER, @error, $ADO_RET_FAILURE)
	#forceref $oADO_COM_ErrorHandler

	__ADO_Connection_IsValid($oConnection)
	If @error Then
		Return SetError(@error, @extended, $ADO_RET_FAILURE)
	EndIf

	Local $rs = $oConnection.OpenSchema($ADO_adSchemaTables)
	If @error Then
		Return SetError($ADO_ERR_COMERROR, @error, $ADO_RET_FAILURE)
	EndIf

	Local $oField = $rs.Fields("TABLE_NAME")
	Local $aTable

	If $sType = "*" Then ; All Table
		Do ;Check for a user table object
			If UBound($aTable) = 0 Then
				Dim $aTable[1][2]
			Else
				ReDim $aTable[UBound($aTable) + 1][2]
			EndIf
			$aTable[UBound($aTable) - 1][0] = $oField.Value
			$aTable[UBound($aTable) - 1][1] = $rs.Fields("TABLE_TYPE").Value
			$rs.MoveNext
		Until $rs.EOF
	Else ; Selected Table
		Do ;Check for a user table object
			If $rs.Fields("TABLE_TYPE").Value = $sType Then
				If UBound($aTable) = 0 Then
					Dim $aTable[1]
				Else
					ReDim $aTable[UBound($aTable) + 1]
				EndIf
				$aTable[UBound($aTable) - 1] = $oField.Value
			EndIf
			$rs.MoveNext
		Until $rs.EOF
	EndIf

	If UBound($aTable) = 0 Then
		Return SetError($ADO_ERR_GENERAL, $ADO_EXT_DEFAULT, $ADO_RET_FAILURE) ; Table Not Found
	EndIf

	Return SetError($ADO_ERR_SUCCESS, $ADO_EXT_DEFAULT, $aTable)

EndFunc    ;==>_SQL_GetTableName
#EndRegion  ADO.au3 - Functions - Getting Data

#Region ADO.au3 - MySQL
#EndRegion  ADO.au3 - MySQL

#Region ADO.au3 - PostgreSQL
#EndRegion  ADO.au3 - PostgreSQL

#Region ADO.au3 - MS Access

; #FUNCTION# ===================================================================
; Name ..........: _ADO_Connection_OpenAccess
; Description ...: Starts a Database Connection to an Access Database
; Syntax ........: _ADO_Connection_OpenAccess(ByRef $oConnection, $sMDB_FileFullPath)
; Parameters ....: $oConnection         - [in/out] an object. ADODB.Connection object.
;                  $sMDB_FileFullPath	- Path to an Access Database file
; Return values .: On Success - Returns $ADO_RET_SUCCESS
;                  On Failure - Returns $ADO_RET_FAILURE and set @error to $ADO_ERR_*
; Author ........: Chris Lambert
; Modified ......: mLipok
; Remarks .......:
; Related .......:
; Link ..........;
; Example .......: no
; ================================================================================
Func _ADO_Connection_OpenAccess(ByRef $oConnection, $sMDB_FileFullPath = "")
	; Error handler, automatic cleanup at end of function
	Local $oADO_COM_ErrorHandler = ObjEvent("AutoIt.Error", __ADO_ComErrorHandler_InternalFunction)
	If @error Then Return SetError($ADO_ERR_COMHANDLER, @error, $ADO_RET_FAILURE)
	#forceref $oADO_COM_ErrorHandler

	__ADO_Connection_IsValid($oConnection)
	If @error Then
		Return SetError(@error, @extended, $ADO_RET_FAILURE)
	EndIf

	$oConnection.Open("Driver={Microsoft Access Driver (*.mdb)};Dbq=" & $sMDB_FileFullPath & ";")
	If @error Then
		Return SetError($ADO_ERR_COMERROR, @error, $ADO_RET_FAILURE)
	Else
		Return SetError($ADO_ERR_SUCCESS, $ADO_EXT_DEFAULT, $ADO_RET_SUCCESS)
	EndIf

EndFunc    ;==>_ADO_Connection_OpenAccess
#EndRegion  ADO.au3 - MS Access

#Region ADO.au3 - MS Excel

; #FUNCTION# ===================================================================
; Name ..........: _ADO_Connection_OpenExcel
; Description ...: Starts a Database Connection to an Excel WorkBook
; Syntax ........: _ADO_Connection_OpenAccess(ByRef $oConnection,$sXLS_FileFullPath)
; Parameters ....: $oConnection         - [in/out] an object. ADODB.Connection object.
;                  $sXLS_FileFullPath	- Path to an Excel file
; Return values .: On Success - Returns $ADO_RET_SUCCESS
;                  On Failure - Returns $ADO_RET_FAILURE and set @error to $ADO_ERR_*
; Author ........: Chris Lambert
; Modified ......: CarlH, mLipok
; Remarks .......:
; Related .......:
; Link ..........;
; Example .......: no
; ================================================================================
Func _ADO_Connection_OpenExcel(ByRef $oConnection, $sXLS_FileFullPath = "", $HDR = "Yes")
	; Error handler, automatic cleanup at end of function
	Local $oADO_COM_ErrorHandler = ObjEvent("AutoIt.Error", __ADO_ComErrorHandler_InternalFunction)
	If @error Then Return SetError($ADO_ERR_COMHANDLER, @error, $ADO_RET_FAILURE)
	#forceref $oADO_COM_ErrorHandler

	__ADO_Connection_IsValid($oConnection)
	If @error Then
		Return SetError(@error, @extended, $ADO_RET_FAILURE)
	EndIf

	Local $sConnection_String = _
			"Provider=Microsoft.Jet.OLEDB.4.0;" & _
			"Data Source=" & $sXLS_FileFullPath & ";" & _
			"Extended Properties='Excel 8.0;HDR=" & $HDR & "';"

	$oConnection.Open($sConnection_String)
	If @error Then
		Return SetError($ADO_ERR_COMERROR, @error, $ADO_RET_FAILURE)
	Else
		Return SetError($ADO_ERR_SUCCESS, $ADO_EXT_DEFAULT, $ADO_RET_SUCCESS)
	EndIf

EndFunc    ;==>_ADO_Connection_OpenExcel
#EndRegion  ADO.au3 - MS Excel

#Region ADO.au3 - TODO and Help/Docs

#CS
	SQLState Property
	https://msdn.microsoft.com/en-us/library/windows/desktop/ms681570(v=vs.85).aspx

	NativeError Property (ADO)
	https://msdn.microsoft.com/en-us/library/windows/desktop/ms678049(v=vs.85).aspx

	Programming ADO SQL Server Applications
	https://technet.microsoft.com/en-us/library/aa905875(v=sql.80).aspx
	https://technet.microsoft.com/en-us/library/aa214053(v=sql.80).aspx

	ADO API Reference
	https://msdn.microsoft.com/en-us/library/windows/desktop/ms678086(v=vs.85).aspx

	ADO Code Examples
	https://msdn.microsoft.com/en-us/library/windows/desktop/ms681484(v=vs.85).aspx

	Microsoft OLE DB Provider for SQL Server
	https://msdn.microsoft.com/en-us/library/windows/desktop/ms677227(v=vs.85).aspx

	OpenSchema Method Example (VB)
	https://msdn.microsoft.com/en-us/library/windows/desktop/ms675853(v=vs.85).aspx

	Errors Collection Properties, Methods, and Events
	https://msdn.microsoft.com/en-us/library/windows/desktop/ms676176(v=vs.85).aspx

	ErrorValueEnum
	https://msdn.microsoft.com/en-us/library/windows/desktop/ms677004(v=vs.85).aspx

	ADO Code Examples VBScript
	https://msdn.microsoft.com/en-us/library/ms676589(v=vs.85).aspx

	ADO Code Examples in Visual Basic
	https://msdn.microsoft.com/en-us/library/ms675104(v=vs.85).aspx
#CE

#CS ADO Events some reference

	Handling ADO Events
	https://msdn.microsoft.com/en-us/library/windows/desktop/ms681467(v=vs.85).aspx

	ADO Event Handler Summary
	https://msdn.microsoft.com/en-us/library/ms677579(v=vs.85).aspx

	Handling Errors and Messages in ADO
	https://technet.microsoft.com/en-us/library/aa905919(v=sql.80).aspx

	ExecuteComplete Event (ADO)
	https://msdn.microsoft.com/en-us/library/windows/desktop/ms676183(v=vs.85).aspx


	ADO Error Reference
	https://msdn.microsoft.com/en-us/library/ms681549(v=vs.85).aspx

	ADO Collections
	https://msdn.microsoft.com/en-us/library/ms677591(v=vs.85).aspx

	WillChangeRecordset and RecordsetChangeComplete Events (ADO)
	https://msdn.microsoft.com/en-us/library/ms680919(v=vs.85).aspx


	Handling Errors and Messages in ADO
	https://technet.microsoft.com/en-us/library/aa905919(v=sql.80).aspx

	Performing Transactions in ADO
	https://technet.microsoft.com/en-us/library/aa905921(v=sql.80).aspx

	An ADO Transaction
	https://msdn.microsoft.com/en-us/library/aa227162(v=vs.60).aspx

	ADO BeginTrans, CommitTrans, and RollbackTrans Methods
	http://www.w3schools.com/asp/met_conn_begintrans.asp

	BeginTrans, CommitTrans, and RollbackTrans Methods Example (VB)
	https://msdn.microsoft.com/en-us/library/windows/desktop/ms677538%28v=vs.85%29.aspx

#CE

#CS
	View Object (ADOX)
	https://msdn.microsoft.com/en-us/library/ms676503(v=vs.85).aspx

	Views Collection (ADOX)
	https://msdn.microsoft.com/en-us/library/ms677523(v=vs.85).aspx

	Views Collection, CommandText Property Example (VB)
	https://msdn.microsoft.com/en-us/library/ms677503(v=vs.85).aspx

	Views and Fields Collections Example (VB)
	https://msdn.microsoft.com/en-us/library/ms680939(v=vs.85).aspx


	How To Determine Number of Records Affected by an ADO UPDATE
	https://support.microsoft.com/en-us/kb/195048

	ADO Programmer's Guide
	https://msdn.microsoft.com/en-us/library/ms681025(v=vs.85).aspx

	ADO Programmer's Reference
	https://msdn.microsoft.com/en-us/library/ms676539(v=vs.85).aspx

	ADO Objects and Interfaces
	https://msdn.microsoft.com/en-us/library/ms679836(v=vs.85).aspx

	ADOX Programming Code Examples
	http://allenbrowne.com/func-adox.html

	ADO Programming Code Examples
	http://allenbrowne.com/func-ADO.html

	Driver Specification Subkeys
	https://msdn.microsoft.com/en-us/library/ms714538(v=vs.85).aspx
	Local $key = "HKEY_LOCAL_MACHINE\SOFTWARE\ODBC\ODBCINST.INI\ODBC Drivers"

	The SQL Server Native Client...
	https://msdn.microsoft.com/pl-pl/sqlserver/aa937733.aspx

	Get Started Developing with the SQL Server Native Client
	https://msdn.microsoft.com/pl-pl/sqlserver/ff658533

	Building Applications with SQL Server Native Client
	https://msdn.microsoft.com/en-us/library/ms130904.aspx

	When to Use SQL Server Native Client
	https://msdn.microsoft.com/en-us/library/ms130828.aspx

	What's New in SQL Server Native Client
	https://msdn.microsoft.com/en-us/library/cc280510.aspx

	SQL Server Native Client Features
	https://msdn.microsoft.com/en-us/library/ms131456.aspx

	SQL Server Native Client Programming
	https://msdn.microsoft.com/en-us/library/ms130892.aspx

	Native API for SQL Server FAQ
	https://msdn.microsoft.com/en-us/sqlserver/aa937707.aspx

#CE
#EndRegion  ADO.au3 - TODO and Help/Docs

#Region ADO.au3 - NEW WIP

; #FUNCTION# ====================================================================================================================
; Name ..........: _ADO_Connection_PropertiesToArray
; Description ...: List all Connection Properties
; Syntax ........: _ADO_Connection_PropertiesToArray(Byref $oConnection)
; Parameters ....: $oConnection         - [in/out] an object. ADODB.Connection object.
; Return values .: On Success - Returns $aProperties
;                  On Failure - Returns $ADO_RET_FAILURE and set @error to $ADO_ERR_*
; Author ........: water
; Modified ......: mLipok
; Remarks .......:
; Related .......:
; Link ..........: https://www.autoitscript.com/wiki/ADO_Tools
; Example .......: No
; ===============================================================================================================================
Func _ADO_Connection_PropertiesToArray(ByRef $oConnection)
	; Error handler, automatic cleanup at end of function
	Local $oADO_COM_ErrorHandler = ObjEvent("AutoIt.Error", __ADO_ComErrorHandler_InternalFunction)
	If @error Then Return SetError($ADO_ERR_COMHANDLER, @error, $ADO_RET_FAILURE)
	#forceref $oADO_COM_ErrorHandler

	__ADO_Connection_IsValid($oConnection)
	If @error Then Return SetError(@error, @extended, $ADO_RET_FAILURE)

	; Property Object (ADO)
	; https://msdn.microsoft.com/en-us/library/windows/desktop/ms677577(v=vs.85).aspx
	Local $oProperties_coll = $oConnection.Properties
	Local $aProperties[$oProperties_coll.count][4]
	Local $iIndex = 0

	For $oProperty_enum In $oProperties_coll
		$aProperties[$iIndex][0] = $oProperty_enum.Name
		$aProperties[$iIndex][1] = $oProperty_enum.Type
		$aProperties[$iIndex][2] = $oProperty_enum.Value
		$aProperties[$iIndex][3] = $oProperty_enum.Attributes
		$iIndex += 1
	Next

	$oProperties_coll = Null
	Return SetError($ADO_ERR_SUCCESS, $ADO_EXT_DEFAULT, $aProperties)

EndFunc    ;==>_ADO_Connection_PropertiesToArray
#EndRegion  ADO.au3 - NEW WIP
