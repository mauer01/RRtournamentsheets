#AutoIt3Wrapper_AU3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6

#include-once
#include "OOoCalcConstants.au3"

; #INDEX# =======================================================================================================================
; Title .........: OOoCalc
; AutoIt Version : 3.3.7.20++
; Language ......: Functions that assist in the automation of OpenOffice/LibreOffice Calc
; Author(s) .....: Leagnus, Andy G, GMK, mLipok
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
;_OOoCalc_BookNew
;_OOoCalc_BookOpen
;_OOoCalc_BookAttach
;_OOoCalc_BookSave
;_OOoCalc_BookSaveAs
;_OOoCalc_BookClose
;_OOoCalc_WriteCell
;_OOoCalc_WriteFormula
;_OOoCalc_WriteFromArray
;_OOoCalc_HyperlinkInsert
;_OOoCalc_RangeMoveOrCopy
;_OOoCalc_RangeSort
;_OOoCalc_RangeClearContents
;_OOoCalc_CreateBorders
;_OOoCalc_NumberFormat
;_OOoCalc_ReadCell
;_OOoCalc_ReadSheetToArray
;_OOoCalc_RowDelete
;_OOoCalc_ColumnDelete
;_OOoCalc_RowInsert
;_OOoCalc_ColumnInsert
;_OOoCalc_SheetAddNew
;_OOoCalc_SheetDelete
;_OOoCalc_SheetNameGet
;_OOoCalc_SheetNameSet
;_OOoCalc_SheetList
;_OOoCalc_SheetActivate
;_OOoCalc_SheetSetVisibility
;_OOoCalc_SheetMove
;_OOoCalc_SheetPrint
;_OOoCalc_HorizontalAlignSet
;_OOoCalc_FontSetProperties
;_OOoCalc_CellSetColors
;_OOoCalc_RowSetColors
;_OOoCalc_ColumnSetColors
;_OOoCalc_RowSetProperties
;_OOoCalc_ColumnSetProperties
;_OOoCalc_FindInRange
;_OOoCalc_ReplaceInRange
; ===============================================================================================================================

; #INTERNAL_USE_ONLY# ===========================================================================================================
;__OOoCalc_CellA1ToRC
;__OOoCalc_CellRCToA1
;__OOoCalc_CreateBorderLine
;__OOoCalc_FileToURL
;__OOoCalc_GetCell
;__OOoCalc_GetRange
;__OOoCalc_GetSheet
;__OOoCalc_GetUsedRange
;__OOoCalc_RangeIsValid
;__OOoCalc_SetPropertyValue
;__OOoCalc_WorksheetIsValid
;__OOoCalc_ComErrorHandler_InternalFunction
;_OOoCalc_ComErrorHandler_UserFunction
; ===============================================================================================================================

#cs #CHANGES ====================================================================================================================

	2016/02/12
	Moved constants and enums to separate file
	Script-Breaking Change: Renamed all functions to start with _OOoCalc_ or __OOoCalc_
	Added internal error handler for all functions
	Added __OOoCalc_ComErrorHandler_InternalFunction and _OOoCalc_ComErrorHandler_UserFunction
	Removed error notification, along with __OOoCalcErrorNotify function
	_OOoCalc_RowSetProperties and _OOoCalc_ColumnSetProperties: Added Return SetError($_OOoCalcStatus_Success, 0, 1)
	Removed _OOoCalc_ErrorHandlerRegister and _OOoCalc_ErrorHandlerDeRegister
	_OOoCalc_SheetMove: Converted If...Then...Else code to use ternary operators
	_OOoCalc_FindInRange: Added local enum variables $eAddress, $eRow, $eColumn, $eString

	2016/08/05
	Potential Script-Breaking Change: Corrected number constants not matching documentation
	Removed unused global variables from OOoCalcConstants.au3
	Corrected header for _OOoCalc_NumberFormat
	Pulled _DateTimeSplit and _DateIsValid from Date.au3 to avoid using the include
	_OOoCalc_WriteCell sets cell formula for dates (if value is a valid date)

	2016/08/17
	Fixed _OOoCalc_BookAttach to work when more than one Calc spreadsheet is open

	2016/11/14
	Added ByRef to parameter in __OOoCalc_ComErrorHandler_InternalFunction (Thanks to mLipok for catching that)
	Removed functions imported from Date.au3, as they were no longer used in _OOoCalc_WriteCell
	Implemented changes to __OOoCalc_WorksheetIsValid by mLipok (thanks, mLipok!)
	Finally fixed _OOoCalc_RangeSort
	Removed __OOoCalc_SetSortField, as it is no longer needed

	2016/11/15
	Made all error checks for $oObj consistent to return $_OOoCalcStatus_InvalidDataType with @extended = 1 (Thanks to mLipok)
	Revised _OOoCalc_BookAttach error checking
	Revised _OOoCalc_BookSave error checking
	Added error checking to _OOoCalc_RangeSort
	Added error check for $iRow in _OOoCalc_RowDelete and _OOoCalc_RowInsert (Thanks to mLipok)
	Added error check for $iCol in _OOoCalc_ColumnDelete and _OOoCalc_ColumnInsert (Thanks to mLipok)
	Revised _OOoCalc_SheetAddNew error checking (Thanks to mLipok)
	Revised _OOoCalc_SheetDelete error checking
	Added error checking to _OOoCalc_SheetSetVisibility
	Corrected @extended return for _OOoCalc_CellSetColors sheet not found
	Corrected @extended return for _OOoCalc_RowSetProperties sheet not found
	Corrected @extended return for _OOoCalc_ColumnSetProperties sheet not found
	Corrected @extended return for _OOoCalc_FindInRange sheet not found
	Corrected @extended freutn for _OOoCalc_ReplaceInRange sheet not found
	Simplified __OOoCalc_CellA1ToRC
	Added error checking to __OOoCalc_CreateBorderLine
	Added error checking to __OOoCalc_FileToURL
	Added error checking to __OOoCalc_GetCell
	Revised __OOoCalc_GetRange
	Revised __OOoCalc_GetSheet
	Simplified __OOoCalc_RangeIsValid
	Added error checking to __OOoCalc_SetPropertyValue
	Modified _OOoCalc_WriteFromArray to put some variable declarations after error checking
	Replaced $iLineWidth with $nLineWidth

	2016/11/16
	Added ByRef to object parameters

#ce =============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _OOoCalc_BookNew
; Description ...: Creates a new workbook and returns its object identifier.
; Syntax ........: _OOoCalc_BookNew([$bHidden = False])
;                  $bHidden   - [optional] Flag, whether to open the workbook as visible (True or False)
;                                          Default is True.
; Return values .: On Success - Returns an object variable pointing to active com.sun.star.frame.Desktop Component object
;                  On Failure - Returns 0 and sets @error:
;                  |@error    - 0 ($_OOoCalcStatus_Success)      = No Error
;                  |          - 1 ($_OOoCalcStatus_GeneralError) = General Error
;                  |          - 3 ($_OOoCalcStatus_InvalidDataType) = Invalid Data Type
; Author ........: Leagnus, GMK
; Modified ......:
; Remarks .......: None
; Related .......: _OOoCalc_BookAttach, _OOoCalc_BookOpen
; Link ..........: http://www.openoffice.org/api/docs/common/ref/com/sun/star/frame/XComponentLoader.html#loadComponentFromURL
; Example .......: No
; ===============================================================================================================================
Func _OOoCalc_BookNew($bHidden = False)
	Local $oOOoCalc_COM_ErrorHandler = ObjEvent("AutoIt.Error", __OOoCalc_ComErrorHandler_InternalFunction)
    #forceref $oOOoCalc_COM_ErrorHandler
	If Not IsBool($bHidden) Then Return SetError($_OOoCalcStatus_InvalidDataType, 1, 0)
	Local $oSM = ObjCreate("com.sun.star.ServiceManager")
	If Not IsObj($oSM) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	Local $oDesktop = $oSM.createInstance("com.sun.star.frame.Desktop") ; Create a desktop object
	If Not IsObj($oDesktop) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	Local $aoProperties[1] = [__OOoCalc_SetPropertyValue('ReadOnly', False)]
	If Not IsObj($aoProperties[0]) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	If $bHidden Then
		ReDim $aoProperties[2]
		$aoProperties[1] = __OOoCalc_SetPropertyValue('Hidden', True)
		If Not IsObj($aoProperties[1]) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	EndIf
	Local $oReturn = $oDesktop.loadComponentFromURL('private:factory/scalc', '_default', 0, $aoProperties)
	If Not IsObj($oReturn) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	Return SetError($_OOoCalcStatus_Success, 0, $oReturn)
EndFunc   ;==>_OOoCalc_BookNew

; #FUNCTION# ====================================================================================================================
; Name ..........: _OOoCalc_BookOpen
; Description ...: Opens an existing workbook and returns its object identifier.
; Syntax ........: _OOoCalc_BookOpen($sFileName[, $bHidden = True[, $bReadOnly = True[, $sPassword = '']])
; Parameters ....: $sFileName - Path and filename of the file to be opened
;                  $bHidden   - [optional] Flag, whether to open the workbook as visible (True or False)
;                                          Default is True.
;                  $bReadOnly - [optional] Flag, whether to open the workbook as read-only (True or False)
;                                          Default is False.
;                  $sPassword - [optional] The password that was used to read-protect the workbook, if any
;                                          Default is ''.
; Return values .: On Success - Returns an object variable pointing to active com.sun.star.frame.Desktop Component object
;                  On Failure - Returns 0 and sets @error:
;                  |@error    - 0 ($_OOoCalcStatus_Success)         = No Error
;                  |          - 1 ($_OOoCalcStatus_GeneralError)    = General Error
;                  |          - 3 ($_OOoCalcStatus_InvalidDataType) = Invalid Data Type
;                  |          - 5 ($_OOoCalcStatus_NoMatch)         = No Match
;                  |@extended - Contains Invalid Parameter Number
; Author ........: Leagnus, GMK
; Modified ......:
; Remarks .......: None
; Related .......: _OOoCalc_BookAttach, _OOoCalc_BookNew
; Link ..........: http://www.openoffice.org/api/docs/common/ref/com/sun/star/frame/XComponentLoader.html#loadComponentFromURL
; Example .......: No
; ===============================================================================================================================
Func _OOoCalc_BookOpen($sFileName, $bHidden = False, $bReadOnly = False, $sPassword = '')
	Local $oOOoCalc_COM_ErrorHandler = ObjEvent("AutoIt.Error", __OOoCalc_ComErrorHandler_InternalFunction)
    #forceref $oOOoCalc_COM_ErrorHandler
	If Not IsString($sFileName) Then Return SetError($_OOoCalcStatus_InvalidDataType, 1, 0)
	If Not FileExists($sFileName) Then Return SetError($_OOoCalcStatus_NoMatch, 1, 0)
	If Not IsBool($bHidden) Then Return SetError($_OOoCalcStatus_InvalidDataType, 2, 0)
	If Not IsBool($bReadOnly) Then Return SetError($_OOoCalcStatus_InvalidDataType, 3, 0)
	If Not IsString($sPassword) Then Return SetError($_OOoCalcStatus_InvalidDataType, 4, 0)
	Local $oSM = ObjCreate("com.sun.star.ServiceManager")
	If Not IsObj($oSM) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	Local $oDesktop = $oSM.createInstance("com.sun.star.frame.Desktop")
	If Not IsObj($oDesktop) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	Local $sURL = __OOoCalc_FileToURL($sFileName)
	Local $aoProperties[2] = [__OOoCalc_SetPropertyValue('Hidden', $bHidden)]
	If Not IsObj($aoProperties[0]) Then Return SetError($_OOoCalcStatus_GeneralError, 2, 0)
	$aoProperties[1] = __OOoCalc_SetPropertyValue('ReadOnly', $bReadOnly)
	If Not IsObj($aoProperties[1]) Then Return SetError($_OOoCalcStatus_GeneralError, 3, 0)
	If $sPassword <> '' Then
		ReDim $aoProperties[2]
		$aoProperties[1] = __OOoCalc_SetPropertyValue('Password', $sPassword)
		If Not IsObj($aoProperties[1]) Then Return SetError($_OOoCalcStatus_GeneralError, 4, 0)
	EndIf
	Local $oReturn = $oDesktop.loadComponentFromURL($sURL, '_default', 0, $aoProperties)
	If Not IsObj($oReturn) Then Return SetError($_OOoCalcStatus_GeneralError, 1, 0)
	Return SetError($_OOoCalcStatus_Success, 0, $oReturn)
EndFunc   ;==>_OOoCalc_BookOpen

; #FUNCTION# ====================================================================================================================
; Name ..........: _OOoCalc_BookAttach
; Description ...: Attach to the first existing instance of OOo Calc
; Syntax ........: _OOoCalc_BookAttach($sFileName)
; Parameters ....: $sFileName - FileName of the open workbook with extension and without full UNC path
; Return values .: On Success - Returns an object variable pointing to active com.sun.star.frame.Desktop Component object
;                  On Failure - Returns 0 and sets @error:
;                  |@error    - 0 ($_OOoCalcStatus_Success)      = No Error
;                  |          - 1 ($_OOoCalcStatus_GeneralError) = General Error
;                  |          - 5 ($_OOoCalcStatus_NoMatch)      = No Match
; Author ........: Leagnus. Thanks to ms777; GMK
; Modified ......:
; Remarks .......: None
; Related .......: _OOoCalc_BookNew, _OOoCalc_BookOpen
; Link ..........: http://www.openoffice.org/api/docs/common/ref/com/sun/star/frame/XDesktop.html#getCurrentComponent
; Example .......: No
; ===============================================================================================================================
Func _OOoCalc_BookAttach($sFileName)
	Local $oOOoCalc_COM_ErrorHandler = ObjEvent("AutoIt.Error", __OOoCalc_ComErrorHandler_InternalFunction)
    #forceref $oOOoCalc_COM_ErrorHandler
	If Not IsString($sFileName) Then Return SetError($_OOoCalcStatus_InvalidDataType, 1, 0)
	If Not FileExists($sFileName) Then Return SetError($_OOoCalcStatus_NoMatch, 1, 0)
	Local $oSM = ObjCreate("com.sun.star.ServiceManager")
	If Not IsObj($oSM) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	Local $oDesktop = $oSM.createInstance("com.sun.star.frame.Desktop")
	If Not IsObj($oDesktop) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	Local $iFrameNum = $oDesktop.Frames.Count
	Local $iInstance = 1
	If $iFrameNum > 1 Then
		Local $asWinList = WinList('[CLASS:SALFRAME]', '')
		Local $asCurWin
		For $iCounter = 1 To $asWinList[0][0]
			$asCurWin = StringRegExp($asWinList[$iCounter][0], '(.*?\.\w{3}).*Calc', 1)
			If Not @error And StringInStr($asCurWin[0], $sFileName) Then
				$iInstance = $iCounter
				ExitLoop
			EndIf
		Next
	EndIf
	Local $sWinTitle = WinGetTitle('[CLASS:SALFRAME; INSTANCE:' & $iInstance & ']')
	If StringInStr($sWinTitle, $sFileName) <> 1 Then Return SetError($_OOoCalcStatus_NoMatch, 0, 0)
	WinActivate('[CLASS:SALFRAME; INSTANCE:' & $iInstance & ']')
	Local $oReturn = $oDesktop.getCurrentComponent
	If Not IsObj($oReturn) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	Return SetError($_OOoCalcStatus_Success, 0, $oReturn)
EndFunc   ;==>_OOoCalc_BookAttach

; #FUNCTION# ====================================================================================================================
; Name ..........: _OOoCalc_BookSave
; Description ...: Saves (stores) the currently opened workbook of the specified Calc object.
; Syntax ........: _OOoCalc_BookSave(ByRef $oObj)
; Parameters ....: $oObj     - Calc object opened by a preceding call to _OOoCalc_BookOpen, _OOoCalc_BookNew, or
;                               _OOoCalc_BookAttach
; Return values .: On Success - Returns 1 and saves the workbook
;                  On Failure - Returns 0 and sets @error:
;                  |@error    - 0 ($_OOoCalcStatus_Success)         = No Error
;                  |          - 1 ($_OOoCalcStatus_GeneralError)    = General Error
;                  |          - 3 ($_OOoCalcStatus_InvalidDataType) = Invalid Data Type
; Author ........: GMK
; Modified ......:
; Remarks .......: None
; Related .......: _OOoCalc_BookSaveAs
; Link ..........: http://www.openoffice.org/api/docs/common/ref/com/sun/star/frame/XStorable.html#store
; Example .......: No
; ===============================================================================================================================
Func _OOoCalc_BookSave(ByRef $oObj)
	Local $oOOoCalc_COM_ErrorHandler = ObjEvent("AutoIt.Error", __OOoCalc_ComErrorHandler_InternalFunction)
    #forceref $oOOoCalc_COM_ErrorHandler
	If Not IsObj($oObj) Then Return SetError($_OOoCalcStatus_InvalidDataType, 1, 0)
	If Not $oObj.hasLocation Or $oObj.isReadOnly Then Return SetError($_OOoCalcStatus_GeneralError, 0, 1)
	$oObj.store()
	Return SetError($_OOoCalcStatus_Success, 0, 1)
EndFunc   ;==>_OOoCalc_BookSave

; #FUNCTION# ====================================================================================================================
; Name ..........: _OOoCalc_BookSaveAs
; Description ...: Saves Calc workbook with the specified file name of the specified Calc object.
; Syntax ........: _OOoCalc_BookSaveAs(ByRef $oObj, $sFilePath[, $sFilterName = ''[, $bOverwrite = False]])
; Parameters ....: $oObj           - Calc object opened by a preceding call to _OOoCalc_BookOpen, _OOoCalc_BookNew, or
;                                    _OOoCalc_BookAttach
;                  $sFilePath      - File name
;                  $sFilterName    - [optional] Filter name.
;                                               Default is ''.
;                  $bOverwrite     - [optional] Overwrite flag
;                                               Default is False.
; Return values .: On Success      - Returns 1
;                  On Failure      - Returns 0 and sets @error:
;                  |@error         - 0 ($_OOoCalcStatus_Success)         = No Error
;                  |               - 1 ($_OOoCalcStatus_GeneralError) = General Error
;                  |               - 3 ($_OOoCalcStatus_InvalidDataType) = Invalid Data Type
;                  |@extended      - Contains Invalid Parameter Number
; Author ........: Andy G, GMK
; Modified ......:
; Remarks .......: None
; Related .......: _OOoCalc_BookSave
; Link ..........: http://www.openoffice.org/api/docs/common/ref/com/sun/star/frame/XStorable.html#storeAsURL
; Example .......: No
; ===============================================================================================================================
Func _OOoCalc_BookSaveAs(ByRef $oObj, $sFilePath, $sFilterName = '', $bOverwrite = False)
	Local $oOOoCalc_COM_ErrorHandler = ObjEvent("AutoIt.Error", __OOoCalc_ComErrorHandler_InternalFunction)
    #forceref $oOOoCalc_COM_ErrorHandler
	If Not IsObj($oObj) Then Return SetError($_OOoCalcStatus_InvalidDataType, 1, 0)
	If Not IsString($sFilePath) Then Return SetError($_OOoCalcStatus_InvalidDataType, 2, 0)
	If Not IsString($sFilterName) Then Return SetError($_OOoCalcStatus_InvalidDataType, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($_OOoCalcStatus_InvalidDataType, 4, 0)
	$sFilePath = __OOoCalc_FileToURL($sFilePath)
	Local $avOptions[2] = [__OOoCalc_SetPropertyValue('FilterName', $sFilterName), __OOoCalc_SetPropertyValue('Overwrite', $bOverwrite)]
	$oObj.storeToURL($sFilePath, $avOptions)
	If @error Then Return SetError($_OOoCalcStatus_ComError, 0, 0)
	Return SetError($_OOoCalcStatus_Success, 0, 1)
EndFunc   ;==>_OOoCalc_BookSaveAs

; #FUNCTION# ====================================================================================================================
; Name ..........: _OOoCalc_BookClose
; Description ...: Closes an existing workbook of the specified Calc object.
; Syntax ........: _OOoCalc_BookClose(ByRef $oObj[, $bDeliverOwnership = True])
; Parameters ....: $oObj              - Calc object opened by a preceding call to _OOoCalc_BookOpen, _OOoCalc_BookNew, or
;                                       _OOoCalc_BookAttach
;                  $bDeliverOwnership - [optional] Deliver ownership?
;                                                  Default is True.
; Return values .: On Success         - Returns 1
;                  On Failure         - Returns 0 and sets @error:
;                  |@error            - 0 ($_OOoCalcStatus_Success)         = No Error
;                  |                  - 3 ($_OOoCalcStatus_InvalidDataType) = Invalid Data Type
;                  |@extended         - Contains Invalid Parameter Number
; Author ........: Leagnus, GMK
; Modified ......:
; Remarks .......: Document will not be saved automatically before closing.
; Related .......: None
; Link ..........: http://www.openoffice.org/api/docs/common/ref/com/sun/star/util/XCloseable.html#close
; Example .......: No
; ===============================================================================================================================
Func _OOoCalc_BookClose(ByRef $oObj, $bDeliverOwnership = True)
	Local $oOOoCalc_COM_ErrorHandler = ObjEvent("AutoIt.Error", __OOoCalc_ComErrorHandler_InternalFunction)
    #forceref $oOOoCalc_COM_ErrorHandler
	If Not IsObj($oObj) Then Return SetError($_OOoCalcStatus_InvalidDataType, 1, 0)
	If Not IsBool($bDeliverOwnership) Then Return SetError($_OOoCalcStatus_InvalidDataType, 2, 0)
	$oObj.Close($bDeliverOwnership)
	Return SetError($_OOoCalcStatus_Success, 0, 1)
EndFunc   ;==>_OOoCalc_BookClose

; #FUNCTION# ====================================================================================================================
; Name ..........: _OOoCalc_WriteCell
; Description ...: Write a string or value to the specified worksheet cell in a workbook of the specified Calc object.
; Syntax ........: _OOoCalc_WriteCell(ByRef $oObj, $vValue, $vRangeOrRow,[ $iCol = -1,[ $vSheet = -1]])
; Parameters ....: $oObj        - Calc object opened by a preceding call to _OOoCalc_BookOpen, _OOoCalc_BookNew, or
;                                 _OOoCalc_BookAttach
;                  $vValue      - Value or string to be written
;                  $vRangeOrRow - Either an A1 range, or index (0-based) of the row to which to write if using RC.
;                  $iCol        - [optional] Index (0-based) of the column to which to write if using RC.
;                                            Default is -1.
;                  $vSheet      - [optional] Worksheet, either by index (0-based) or name.
;                                            Default is -1, which would use the active worksheet.
; Return values .: On Success   - Returns 1
;                  On Failure   - Returns 0 and sets @error:
;                  |@error      - 0 ($_OOoCalcStatus_Success)         = No Error
;                  |            - 1 ($_OOoCalcStatus_GeneralError     = General Error
;                  |            - 3 ($_OOoCalcStatus_InvalidDataType) = Invalid Data Type
;                  |            - 4 ($_OOoCalcStatus_InvalidValue)    = Invalid Value
;                  |            - 5 ($_OOoCalcStatus_NoMatch)         = No Match
;                  |@extended   - Contains Invalid Parameter Number
; Author ........: Leagnus, GMK
; Modified ......:
; Remarks .......: None
; Related .......: _OOoCalc_ReadCell, _OOoCalc_WriteFormula
; Link ..........: http://www.openoffice.org/api/docs/common/ref/com/sun/star/table/XCell.html
; Example .......: No
; ===============================================================================================================================
Func _OOoCalc_WriteCell(ByRef $oObj, $vValue, $vRangeOrRow, $iCol = -1, $vSheet = -1)
	Local $oOOoCalc_COM_ErrorHandler = ObjEvent("AutoIt.Error", __OOoCalc_ComErrorHandler_InternalFunction)
    #forceref $oOOoCalc_COM_ErrorHandler
	If Not IsObj($oObj) Then Return SetError($_OOoCalcStatus_InvalidDataType, 1, 0)
	If Not IsInt($vRangeOrRow) And Not IsString($vRangeOrRow) Then Return SetError($_OOoCalcStatus_InvalidDataType, 3, 0)
	If Not IsInt($vRangeOrRow) And Not __OOoCalc_RangeIsValid($vRangeOrRow) Then Return SetError($_OOoCalcStatus_InvalidValue, 3, 0)
	If Not IsInt($iCol) Then Return SetError($_OOoCalcStatus_InvalidDataType, 4, 0)
	If IsInt($vRangeOrRow) And $iCol < 0 Then Return SetError($_OOoCalcStatus_InvalidValue, 4, 0)
	If Not IsInt($vSheet) And Not IsString($vSheet) Then Return SetError($_OOoCalcStatus_InvalidDataType, 5, 0)
	If IsString($vSheet) And Not __OOoCalc_WorksheetIsValid($oObj, $vSheet) Then Return SetError($_OOoCalcStatus_NoMatch, 5, 0)
	Local $oSheet = __OOoCalc_GetSheet($oObj, $vSheet)
	If Not IsObj($oSheet) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	Local $oCell = __OOoCalc_GetCell($oSheet, $vRangeOrRow, $iCol)
	If Not IsObj($oCell) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	$oCell.setFormula($vValue)
	Return SetError($_OOoCalcStatus_Success, 0, 1)
EndFunc   ;==>_OOoCalc_WriteCell

; #FUNCTION# ====================================================================================================================
; Name ..........: _OOoCalc_WriteFormula
; Description ...: Write a formula to the specified worksheet cell in a workbook of the specified Calc object.
; Syntax ........: _OOoCalc_WriteFormula(ByRef $oObj, $sFormula, $vRangeOrRow[, $iCol = -1[, $vSheet = -1]])
; Parameters ....: $oObj        - Calc object opened by a preceding call to _OOoCalc_BookOpen, _OOoCalc_BookNew, or
;                                 _OOoCalc_BookAttach
;                  $sFormula    - Formula to be written
;                  $vRangeOrRow - Either an A1 range, or index (0-based) of the row to which to write if using RC.
;                  $iCol        - [optional] Index (0-based) of the column to which to write if using RC.
;                                            Default is -1.
;                  $vSheet      - [optional] Worksheet, either by index (0-based) or name.
;                                            Default is -1, which would use the active worksheet.
; Return values .: On Success   - Returns 1
;                  On Failure   - Returns 0 and sets @error:
;                  |@error      - 0 ($_OOoCalcStatus_Success)         = No Error
;                  |            - 3 ($_OOoCalcStatus_InvalidDataType) = Invalid Data Type
;                  |            - 4 ($_OOoCalcStatus_InvalidValue)    = Invalid Value
;                  |            - 5 ($_OOoCalcStatus_NoMatch)         = No Match
;                  |@extended   - Contains Invalid Parameter Number
; Author ........: Leagnus, GMK
; Modified ......:
; Remarks .......: None
; Related .......: _OOoCalc_WriteCell
; Link ..........: http://www.openoffice.org/api/docs/common/ref/com/sun/star/table/XCell.html#setFormula
; Example .......: No
; ===============================================================================================================================
Func _OOoCalc_WriteFormula(ByRef $oObj, $sFormula, $vRangeOrRow, $iCol = -1, $vSheet = -1)
	Local $oOOoCalc_COM_ErrorHandler = ObjEvent("AutoIt.Error", __OOoCalc_ComErrorHandler_InternalFunction)
    #forceref $oOOoCalc_COM_ErrorHandler
	If Not IsObj($oObj) Then Return SetError($_OOoCalcStatus_InvalidDataType, 1, 0)
	If Not IsString($sFormula) Then Return SetError($_OOoCalcStatus_InvalidDataType, 2, 0)
	If Not IsInt($vRangeOrRow) And Not __OOoCalc_RangeIsValid($vRangeOrRow) Then Return SetError($_OOoCalcStatus_InvalidValue, 3, 0)
	If Not IsInt($iCol) Then Return SetError($_OOoCalcStatus_InvalidDataType, 4, 0)
	If IsInt($vRangeOrRow) And $iCol < 0 Then Return SetError($_OOoCalcStatus_InvalidValue, 4, 0)
	If Not IsInt($vSheet) And Not IsString($vSheet) Then Return SetError($_OOoCalcStatus_InvalidDataType, 5, 0)
	If $vSheet > -1 And Not __OOoCalc_WorksheetIsValid($oObj, $vSheet) Then Return SetError($_OOoCalcStatus_NoMatch, 5, 0)
	Local $oSheet = __OOoCalc_GetSheet($oObj, $vSheet)
	If Not IsObj($oSheet) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	Local $oCell = __OOoCalc_GetCell($oSheet, $vRangeOrRow, $iCol)
	If Not IsObj($oCell) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	$oCell.setFormula($sFormula)
	Return SetError($_OOoCalcStatus_Success, 0, 1)
EndFunc   ;==>_OOoCalc_WriteFormula

; #FUNCTION# ====================================================================================================================
; Name ..........: _OOoCalc_WriteFromArray
; Description ...: Write an array to a row or column on the specified worksheet of the specified Calc object.
; Syntax ........: _OOoCalc_WriteFromArray(ByRef $oObj, Byref $avArray, $vRangeOrRow[, $iCol = -1[, $vSheet = -1[, $bTranspose = False]]])
; Parameters ....: $oObj         - Calc object opened by a preceding call to _OOoCalc_BookOpen, _OOoCalc_BookNew, or
;                                  _OOoCalc_BookAttach
;                  $avArray      - The array (ByRef) to write
;                  $vRangeOrRow  - Either an A1 range, or index (0-based) of the row to which to write if using RC.
;                  $iCol         - [optional] Index (0-based) of the column to which to write if using RC.
;                                             Default is -1.
;                  $vSheet       - [optional] Worksheet, either by index (0-based) or name.
;                                             Default is -1, which would use the active worksheet.
;                  $bTranspose   - [optional] Transpose array
;                                             Default is False.
; Return values .: On Success    - Returns 1
;                  On Failure    - Returns 0 and sets @error:
;                  |@error       - 0 ($_OOoCalcStatus_Success)         = No Error
;                  |             - 3 ($_OOoCalcStatus_InvalidDataType) = Invalid Data Type
;                  |             - 4 ($_OOoCalcStatus_InvalidValue)    = Invalid Value
;                  |             - 5 ($_OOoCalcStatus_NoMatch)         = No Match
;                  |@extended    - Contains Invalid Parameter Number
; Author ........: Andy G, GMK
; Modified ......:
; Remarks .......: None
; Related .......: _OOoCalc_WriteCell, _OOoCalc_ReadArray
; Link ..........: http://www.openoffice.org/api/docs/common/ref/com/sun/star/sheet/XCellRangeData.html#setDataArray
; Example .......: No
; ===============================================================================================================================
Func _OOoCalc_WriteFromArray(ByRef $oObj, ByRef $avArray, $vRangeOrRow, $iCol = -1, $vSheet = -1, $bTranspose = False)
	Local $oOOoCalc_COM_ErrorHandler = ObjEvent("AutoIt.Error", __OOoCalc_ComErrorHandler_InternalFunction)
    #forceref $oOOoCalc_COM_ErrorHandler
	If Not IsObj($oObj) Then Return SetError($_OOoCalcStatus_InvalidDataType, 1, 0)
	If Not IsArray($avArray) Then Return SetError($_OOoCalcStatus_InvalidDataType, 2, 0)
	Local $iDim = UBound($avArray, 0)
	If $iDim > 2 Then Return SetError($_OOoCalcStatus_InvalidValue, 2, 0)
	If Not IsInt($vRangeOrRow) And Not __OOoCalc_RangeIsValid($vRangeOrRow) Then Return SetError($_OOoCalcStatus_InvalidValue, 3, 0)
	If Not IsInt($iCol) Then Return SetError($_OOoCalcStatus_InvalidDataType, 4, 0)
	If IsInt($vRangeOrRow) And $iCol < 0 Then Return SetError($_OOoCalcStatus_InvalidValue, 4, 0)
	If Not IsInt($vSheet) And Not IsString($vSheet) Then Return SetError($_OOoCalcStatus_InvalidDataType, 5, 0)
	If $vSheet > -1 And Not __OOoCalc_WorksheetIsValid($oObj, $vSheet) Then Return SetError($_OOoCalcStatus_NoMatch, 5, 0)
	If Not IsBool($bTranspose) Then Return SetError($_OOoCalcStatus_InvalidDataType, 6, 0)
	Local $iRows = UBound($avArray, 1)
	Local $iLastRow = $iRows - 1
	Local $iCols = UBound($avArray, 2)
	Local $iLastCol = $iCols - 1
	Local $avTempArray[1], $avOOoArray[1]
	If $bTranspose Then
		Switch $iDim
			Case 1
				ReDim $avTempArray[1][$iRows]
				For $iCol = 0 To $iLastRow
					$avTempArray[0][$iCol] = $avArray[$iCol]
				Next
			Case Else
				ReDim $avTempArray[$iCols][$iRows]
				For $iRow = 0 To $iLastRow
					For $iCol = 0 To $iLastCol
						$avTempArray[$iLastCol - $iCol][$iRow] = $avArray[$iRow][$iCol]
					Next
				Next
		EndSwitch
		$avArray = $avTempArray
		$iDim = UBound($avArray, 0)
		$iRows = UBound($avArray, 1)
		$iLastRow = $iRows - 1
		$iCols = UBound($avArray, 2)
		$iLastCol = $iCols - 1
	EndIf
	Local $oSheet = __OOoCalc_GetSheet($oObj, $vSheet)
	If Not IsObj($oSheet) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	Local $iEndRow, $iEndCol, $oRange, $aiRC
	If IsInt($vRangeOrRow) Then
		$iEndRow = $vRangeOrRow + $iRows - 1
		$iEndCol = $iCol + $iCols - 1
		$oRange = __OOoCalc_GetRange($oSheet, $vRangeOrRow, $iCol, $iEndRow, $iEndCol)
	Else
		$aiRC = __OOoCalc_CellA1ToRC($vRangeOrRow)
		$iEndRow = $aiRC[0] + $iRows - 1
		$iEndCol = $aiRC[1] + $iCols - 1
		$oRange = __OOoCalc_GetRange($oSheet, $aiRC[0], $aiRC[1], $iEndRow, $iEndCol)
	EndIf
	If Not IsObj($oRange) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	If $iDim = 1 Then
		$avOOoArray[0] = $avArray
	Else
		ReDim $avOOoArray[$iRows]
		Local $avRows[$iCols]
		For $iRow = 0 To $iLastRow
			For $iCol = 0 To $iLastCol
				$avRows[$iCol] = $avArray[$iRow][$iCol]
			Next
			$avOOoArray[$iRow] = $avRows
		Next
	EndIf
	$oRange.setDataArray($avOOoArray)
	Return SetError($_OOoCalcStatus_Success, 0, 1)
EndFunc   ;==>_OOoCalc_WriteFromArray

; #FUNCTION# ====================================================================================================================
; Name ..........: _OOoCalc_HyperlinkInsert
; Description ...: Inserts a hyperlink into the specified sheet.
; Syntax ........: _OOoCalc_HyperlinkInsert(ByRef $oObj, $sLinkText, $sAddress, $vRangeOrRow[, $iCol = -1[, $vSheet = -1]])
; Parameters ....: $oObj        - Calc object opened by a preceding call to _OOoCalc_BookOpen, _OOoCalc_BookNew, or
;                                 _OOoCalc_BookAttach
;                  $sLinkText   - Link text
;                  $sAddress    - Valid URL
;                  $vRangeOrRow - Either an A1 range, or index (0-based) of row to which to write if using RC.
;                  $iCol        - [optional] The index (0-based) of the column to which to write if using RC.
;                                            Default is -1.
;                  $vSheet      - [optional] Worksheet, either by index (0-based) or name.
;                                            Default is -1, which would use the active worksheet.
; Return values .: On Success   - Returns 1
;                  On Failure   - Returns 0 and sets @error:
;                  |@error      - 0 ($_OOoCalcStatus_Success)         = No Error
;                  |            - 3 ($_OOoCalcStatus_InvalidDataType) = Invalid Data Type
;                  |            - 4 ($_OOoCalcStatus_InvalidValue)    = Invalid Value
;                  |            - 5 ($_OOoCalcStatus_NoMatch)         = No Match
;                  |@extended   - Contains Invalid Parameter Number
; Author ........: GMK
; Modified ......:
; Remarks .......: None
; Related .......: None
; Link ..........: http://www.openoffice.org/api/docs/common/ref/com/sun/star/text/textfield/URL.html
; Example .......: No
; ===============================================================================================================================
Func _OOoCalc_HyperlinkInsert(ByRef $oObj, $sLinkText, $sAddress, $vRangeOrRow, $iCol = -1, $vSheet = -1)
	Local $oOOoCalc_COM_ErrorHandler = ObjEvent("AutoIt.Error", __OOoCalc_ComErrorHandler_InternalFunction)
    #forceref $oOOoCalc_COM_ErrorHandler
	If Not IsObj($oObj) Then Return SetError($_OOoCalcStatus_InvalidDataType, 1, 0)
	If Not IsString($sLinkText) Then Return SetError($_OOoCalcStatus_InvalidDataType, 2, 0)
	If Not IsString($sAddress) Then Return SetError($_OOoCalcStatus_InvalidDataType, 3, 0)
	If Not IsInt($vRangeOrRow) And Not __OOoCalc_RangeIsValid($vRangeOrRow) Then Return SetError($_OOoCalcStatus_InvalidValue, 4, 0)
	If Not IsInt($iCol) Then Return SetError($_OOoCalcStatus_InvalidDataType, 5, 0)
	If IsInt($vRangeOrRow) And $iCol < 0 Then Return SetError($_OOoCalcStatus_InvalidValue, 5, 0)
	If Not IsInt($vSheet) And Not IsString($vSheet) Then Return SetError($_OOoCalcStatus_InvalidDataType, 6, 0)
	If $vSheet > -1 And Not __OOoCalc_WorksheetIsValid($oObj, $vSheet) Then Return SetError($_OOoCalcStatus_NoMatch, 5, 0)
	Local $oSheet = __OOoCalc_GetSheet($oObj, $vSheet)
	If Not IsObj($oSheet) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	Local $oCell = __OOoCalc_GetCell($oSheet, $vRangeOrRow, $iCol)
	If Not IsObj($oCell) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	Local $oText = $oCell.getText
	If Not IsObj($oText) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	Local $oLink = $oObj.createInstance('com.sun.star.text.TextField.URL')
	If Not IsObj($oLink) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	$oLink.URL = $sAddress
	$oLink.Representation = $sLinkText
	$oCell.insertTextContent($oText.createTextCursor, $oLink, True)
	Return SetError($_OOoCalcStatus_Success, 0, 1)
EndFunc   ;==>_OOoCalc_HyperlinkInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _OOoCalc_RangeMoveOrCopy
; Description ...: Move or copy a specified range to another
; Syntax ........: _OOoCalc_RangeMoveOrCopy(ByRef $oObj, $sSrcRange, $sDestCell, $iFlag[, $vSrcSheet = -1[, $vDestSheet = -1]])
; Parameters ....: $oObj       - Calc object opened by a preceding call to _OOoCalc_BookOpen, _OOoCalc_BookNew, or
;                                _OOoCalc_BookAttach
;                  $sSrcRange  - Source range
;                  $sDestCell  - Destination cell
;                  $iFlag      - Move/copy flag:
;                                0 = Move
;                                1 = Copy
;                  $vSrcSheet  - [optional] Source worksheet, either by index (0-based) or name.
;                                           Default is -1, which would use the active worksheet.
;                  $vDestSheet - [optional] Destination worksheet, either by index (0-based) or name.
;                                           Default is -1, which would use the active worksheet.
; Return values .: On Success  - Returns 1
;                  On Failure  - Returns 0 and sets @error:
;                  |@error     - 0 ($_OOoCalcStatus_Success)         = No Error
;                  |           - 3 ($_OOoCalcStatus_InvalidDataType) = Invalid Data Type
;                  |           - 4 ($_OOoCalcStatus_InvalidValue)    = Invalid Value
;                  |           - 5 ($_OOoCalcStatus_NoMatch)         = No Match
;                  |@extended  - Contains Invalid Parameter Number
; Author ........: Andy G, GMK
; Modified ......:
; Remarks .......: None
; Related .......: None
; Link ..........: http://www.openoffice.org/api/docs/common/ref/com/sun/star/sheet/XCellRangeMovement.html
; Example .......: No
; ===============================================================================================================================
Func _OOoCalc_RangeMoveOrCopy(ByRef $oObj, $sSrcRange, $sDestCell, $iFlag, $vSrcSheet = -1, $vDestSheet = -1)
	Local $oOOoCalc_COM_ErrorHandler = ObjEvent("AutoIt.Error", __OOoCalc_ComErrorHandler_InternalFunction)
    #forceref $oOOoCalc_COM_ErrorHandler
	If Not IsObj($oObj) Then Return SetError($_OOoCalcStatus_InvalidDataType, 1, 0)
	If Not __OOoCalc_RangeIsValid($sSrcRange) Then Return SetError($_OOoCalcStatus_InvalidValue, 2, 0)
	If Not __OOoCalc_RangeIsValid($sDestCell) Then Return SetError($_OOoCalcStatus_InvalidValue, 3, 0)
	If $iFlag <> 0 And $iFlag <> 1 Then Return SetError($_OOoCalcStatus_InvalidValue, 4, 0)
	If Not IsInt($vSrcSheet) And Not IsString($vSrcSheet) Then Return SetError($_OOoCalcStatus_InvalidDataType, 5, 0)
	If $vSrcSheet > -1 And Not __OOoCalc_WorksheetIsValid($oObj, $vSrcSheet) Then Return SetError($_OOoCalcStatus_NoMatch, 5, 0)
	If Not IsInt($vDestSheet) And Not IsString($vDestSheet) Then Return SetError($_OOoCalcStatus_InvalidDataType, 6, 0)
	If $vDestSheet > -1 And Not __OOoCalc_WorksheetIsValid($oObj, $vDestSheet) Then Return SetError($_OOoCalcStatus_NoMatch, 6, 0)
	Local $oSrcSheet = __OOoCalc_GetSheet($oObj, $vSrcSheet)
	If Not IsObj($oSrcSheet) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	Local $oDestSheet = __OOoCalc_GetSheet($oObj, $vDestSheet)
	If Not IsObj($oDestSheet) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	Local $oSrcRange = $oSrcSheet.getCellRangeByName($sSrcRange).RangeAddress
	If Not IsObj($oSrcRange) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	Local $oDestRange = $oDestSheet.getCellRangeByname($sDestCell).CellAddress
	If Not IsObj($oDestRange) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	If $iFlag = 0 Then
		$oSrcSheet.MoveRange($oDestRange, $oSrcRange)
	Else
		$oSrcSheet.CopyRange($oDestRange, $oSrcRange)
	EndIf
	Return SetError($_OOoCalcStatus_Success, 0, 1)
EndFunc   ;==>_OOoCalc_RangeMoveOrCopy

; #FUNCTION# ====================================================================================================================
; Name ..........: _OOoCalc_RangeSort
; Description ...: Sort a range
; Syntax ........: _OOoCalc_RangeSort(ByRef $oObj, $vSheet, $sRange[, $iSortField1 = 0[, $bIsAscending1 = True[, $bHasHeader = False[, $bCaseSensitive = False[, $bByRows = True[, $iSortField2 = -1[, $bIsAscending2 = True[, $iSortField3 = -1[, $bIsAscending3 = True]]]]]]]]])
; Parameters ....: $oObj             - Calc object opened by a preceding call to _OOoCalcBookOpen, _OOoCalcBookNew, or
;                                      _OOoCalcBookAttach
;                  $vSheet           - Worksheet, either by index (0-based) or name.
;                                      -1 would use the active worksheet.
;                  $sRange           - An A1 range.
;                  $iSortField1      - [optional] Sort field 1 (column index (0-based)).
;                                                 Default is 0.
;                  $bIsAscending1    - [optional] Sort field 1 as ascending.
;                                                 Default is True. If False, sort is descending.
;                  $bHasHeader       - [optional] Range has header.
;                                                 Default is False. If True, sort with header.
;                  $bCaseSensitive   - [optional] Range is Case Sensitive.
;                                                 Default is False.
;                  $bByRows          - [optional] Sort By Rows.
;                                                 Default is True. If False, sort by columns.
;                  $iSortField2      - [optional] Sort field 2 (column index (0-based)).
;                                                 Default is -1.
;                  $bIsAscending2    - [optional] Sort field 2 as ascending.
;                                                 Default is True. If False, sort is descending.
;                  $iSortField3      - [optional] Sort field 3 (column index (0-based)).
;                                                 Default is -1.
;                  $bIsAscending3    - [optional] Sort field 3 as ascending.
;                                                 Default is True. If False, sort is descending.
; Return values .: On Success        - Returns 1
;                  On Failure        - Returns 0 and sets @error:
;                  |@error           - 0 ($_OOoCalcStatus_Success)         = No Error
;                  |                 - 3 ($_OOoCalcStatus_InvalidDataType) = Invalid Data Type
;                  |                 - 4 ($_OOoCalcStatus_InvalidValue)    = Invalid Value
;                  |                 - 5 ($_OOoCalcStatus_NoMatch)         = No Match
;                  |@extended        - Contains Invalid Parameter Number
; Author ........: GMK
; Modified ......:
; Remarks .......: None
; Related .......: None
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _OOoCalc_RangeSort(ByRef $oObj, $vSheet, $sRange, $iSortField1 = 0, $bIsAscending1 = True, $bHasHeader = False, $bCaseSensitive = False, $bByRows = True, $iSortField2 = -1, $bIsAscending2 = True, $iSortField3 = -1, $bIsAscending3 = True)
	If Not IsObj($oObj) Then Return SetError($_OOoCalcStatus_InvalidDataType, 1, 0)
	If $vSheet > -1 And Not __OOoCalc_WorksheetIsValid($oObj, $vSheet) Then Return SetError($_OOoCalcStatus_NoMatch, 2, 0)
	If Not IsInt($vSheet) And Not IsString($vSheet) Then Return SetError($_OOoCalcStatus_InvalidDataType, 2, 0)
	If Not __OOoCalc_RangeIsValid($sRange) Then Return SetError($_OOoCalcStatus_InvalidValue, 3, 0)
	If Not IsInt($iSortField1) Then Return SetError($_OOoCalcStatus_InvalidDataType, 4, 0)
	If Not IsBool($bIsAscending1) Then Return SetError($_OOoCalcStatus_InvalidDataType, 5, 0)
	If Not IsBool($bHasHeader) Then Return SetError($_OOoCalcStatus_InvalidDataType, 6, 0)
	If Not IsBool($bCaseSensitive) Then Return SetError($_OOoCalcStatus_InvalidDataType, 7, 0)
	If Not IsBool($bByRows) Then Return SetError($_OOoCalcStatus_InvalidDataType, 8, 0)
	If Not IsInt($iSortField2) Then Return SetError($_OOoCalcStatus_InvalidDataType, 9, 0)
	If Not IsBool($bIsAscending2) Then Return SetError($_OOoCalcStatus_InvalidDataType, 10, 0)
	If Not IsInt($iSortField3) Then Return SetError($_OOoCalcStatus_InvalidDataType, 11, 0)
	If Not IsBool($bIsAscending3) Then Return SetError($_OOoCalcStatus_InvalidDataType, 12, 0)
	If $vSheet <> Default Then _OOoCalc_SheetActivate($oObj, $vSheet)
	Local $oSM = ObjCreate("com.sun.star.ServiceManager")
	If Not IsObj($oSM) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	Local $oDocument = $oObj.CurrentController.Frame
	If Not IsObj($oDocument) Then Return SetError($_OOoCalcStatus_GeneralError, 1, 0)
	Local $oDispatcher = $oSM.createInstance('com.sun.star.frame.DispatchHelper') ; Create a desktop object
	If Not IsObj($oDispatcher) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	Local $oSheet = __OOoCalc_GetSheet($oObj, $vSheet)
	If @error Then Return SetError(@error, 2, 0)
	Local $oRange = __OOoCalc_GetRange($oSheet, $sRange)
	If @error Then Return SetError(@error, 0, 0)
	$oObj.getCurrentController.select($oRange)
	Local $avArgs2[7] = [ _
			__OOoCalc_SetPropertyValue('ByRows', $bByRows), _
			__OOoCalc_SetPropertyValue('HasHeader', $bHasHeader), _
			__OOoCalc_SetPropertyValue('CaseSensitive', $bCaseSensitive), _
			__OOoCalc_SetPropertyValue('IncludeAttribs', true), _
			__OOoCalc_SetPropertyValue('UserDefIndex', 0), _
			__OOoCalc_SetPropertyValue('Col1', $iSortField1), _
			__OOoCalc_SetPropertyValue('Ascending1', $bIsAscending1)]
	For $i = 2 To 3
		If Eval('iSortField' & $i) > -1 Then
			Local $iUBound = UBound($avArgs2)
			ReDim $avArgs2[$iUBound + 2]
			$iUBound = UBound($avArgs2)
			$avArgs2[$iUBound - 2] = __OOoCalc_SetPropertyValue('Col' & $i, Eval('iSortField' & $i))
			$avArgs2[$iUBound - 1] = __OOoCalc_SetPropertyValue('Ascending' & $i, Eval('bIsAscending' & $i))
		EndIf
	Next
	$oDispatcher.executeDispatch($oDocument, '.uno:DataSort', '', 0, $avArgs2)
	Return SetError($_OOoCalcStatus_Success, 0, 1)
EndFunc   ;==>_OOoCalc_RangeSort

; #FUNCTION# ====================================================================================================================
; Name ..........: _OOoCalc_RangeClearContents
; Description ...: Clears contents of given range
; Syntax ........: _OOoCalc_RangeClearContents(ByRef $oObj[, $vRangeOrRowStart = -1[, $iColStart = -1[, $iRowEnd = -1[, $iColEnd = -1[, $vSheet = -1[, $iCellFlag = 31]]]]]])
; Parameters ....: $oObj             - Calc object opened by a preceding call to _OOoCalc_BookOpen, _OOoCalc_BookNew, or
;                                      _OOoCalc_BookAttach
;                  $vRangeOrRowStart - [optional] Either an A1 range, or index (0-based) of row to which to write if using RC.
;                                                 Default is -1 (used range).
;                  $iColStart        - [optional] The index (0-based) of the column to which to write if using RC.
;                                                 Default is -1.
;                  $iRowEnd          - [optional] Index (0-based) of ending row to which to write if using RC.
;                                                 Default is -1.
;                  $iColEnd          - [optional] Index (0-based) of ending column to which to write if using RC.
;                                                 Default is -1.
;                  $vSheet           - [optional] Worksheet, either by index (0-based) or name.
;                                                 Default is -1, which would use the active worksheet.
;                  $iCellFlag        - [optional] Type of cell contents to clear. (May be added.) Default is 31.
;                                                 $VALUE = 1
;                                                 $DATETIME = 2
;                                                 $STRING = 4
;                                                 $ANNOTATION = 8
;                                                 $FORMULA = 16
;                                                 $HARDATTR = 32
;                                                 $STYLES = 64
;                                                 $OBJECTS = 128
;                                                 $EDITATTR = 256
;                                                 $FORMATTED = 512
; Return values .: On Success        - Returns 1
;                  On Failure        - Returns 0 and sets @error:
;                  |@error           - 0 ($_OOoCalcStatus_Success)         = No Error
;                  |                 - 3 ($_OOoCalcStatus_InvalidDataType) = Invalid Data Type
;                  |                 - 5 ($_OOoCalcStatus_NoMatch)         = No Match
;                  |@extended        - Contains Invalid Parameter Number
; Author ........: GMK
; Modified ......:
; Remarks .......: None
; Related .......: None
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _OOoCalc_RangeClearContents(ByRef $oObj, $vRangeOrRowStart = -1, $iColStart = -1, $iRowEnd = -1, $iColEnd = -1, $vSheet = -1, $iCellFlag = 31)
	Local $oOOoCalc_COM_ErrorHandler = ObjEvent("AutoIt.Error", __OOoCalc_ComErrorHandler_InternalFunction)
    #forceref $oOOoCalc_COM_ErrorHandler
	If Not IsObj($oObj) Then Return SetError($_OOoCalcStatus_InvalidDataType, 1, 0)
	If Not IsInt($vRangeOrRowStart) And Not __OOoCalc_RangeIsValid($vRangeOrRowStart) Then Return SetError($_OOoCalcStatus_InvalidValue, 2, 0)
	If Not IsInt($iColStart) Then Return SetError($_OOoCalcStatus_InvalidDataType, 3, 0)
	If $vRangeOrRowStart < -1 And IsInt($vRangeOrRowStart) And $iColStart < 0 Then Return SetError($_OOoCalcStatus_InvalidValue, 3, 0)
	If Not IsInt($iRowEnd) Then Return SetError($_OOoCalcStatus_InvalidDataType, 4, 0)
	If Not IsInt($iColEnd) Then Return SetError($_OOoCalcStatus_InvalidDataType, 5, 0)
	If Not IsInt($vSheet) And Not IsString($vSheet) Then Return SetError($_OOoCalcStatus_InvalidDataType, 6, 0)
	If $vSheet > -1 And Not __OOoCalc_WorksheetIsValid($oObj, $vSheet) Then Return SetError($_OOoCalcStatus_NoMatch, 6, 0)
	Local $oSheet = __OOoCalc_GetSheet($oObj, $vSheet)
	If Not IsObj($oSheet) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	Local $oRange = __OOoCalc_GetRange($oSheet, $vRangeOrRowStart, $iColStart, $iRowEnd, $iColEnd)
	If Not IsObj($oRange) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	$oRange.clearContents($iCellFlag)
	Return SetError($_OOoCalcStatus_Success, 0, 1)
EndFunc   ;==>_OOoCalc_RangeClearContents

; #FUNCTION# ====================================================================================================================
; Name ..........: _OOoCalc_CreateBorders
; Description ...: Create borders in a given range of a Calc object.
; Syntax ........: _OOoCalc_CreateBorders(ByRef $oObj, $vRangeOrRowStart[, $iColStart = -1[, $iRowEnd = -1[, $iColEnd = -1[, $vSheet = -1[, $bTop = True[, $bRight = True[, $bBottom = True[, $bLeft = True[, $bHorizontal = False[, $bVertical = False[, $bDouble = False[, $nLineWidth = 10.5]]]]]]]]]]]])
; Parameters ....: $oObj             - Calc object opened by a preceding call to _OOoCalc_BookOpen, _OOoCalc_BookNew, or
;                                      _OOoCalc_BookAttach
;                  $vRangeOrRowStart - Either an A1 range, or index (0-based) of row to which to write if using RC
;                  $iColStart        - [optional] The index (0-based) of the column to which to write if using RC.
;                                                 Default is -1.
;                  $iRowEnd          - [optional] Index (0-based) of ending row to which to write if using RC.
;                                                 Default is -1.
;                  $iColEnd          - [optional] Index (0-based) of ending column to which to write if using RC.
;                                                 Default is -1.
;                  $vSheet           - [optional] Worksheet, either by index (0-based) or name.
;                                                 Default is -1, which would use the active worksheet.
;                  $bTop             - [optional] Turn on top border.
;                                                 Default is True.
;                  $bRight           - [optional] Turn on right border.
;                                                 Default is True.
;                  $bBottom          - [optional] Turn on bottom border.
;                                                 Default is True.
;                  $bLeft            - [optional] Turn on left border.
;                                                 Default is True.
;                  $bHorizontal      - [optional] Turn on inner horizontal border.
;                                                 Default is False.
;                  $bVertical        - [optional] Turn on inner vertical border.
;                                                 Default is False.
;                  $bDouble          - [optional] Turn on double line.
;                                                 Default is False.
;                  $nLineWidth       - [optional] Line width in 1/100 mm.
;                                                 Default is 10.5 (0.3 pt).
; Return values .: On Success        - Returns 1
;                  On Failure        - Returns 0 and sets @error:
;                  |@error           - 0 ($_OOoCalcStatus_Success)         = No Error
;                  |                 - 3 ($_OOoCalcStatus_InvalidDataType) = Invalid Data Type
;                  |                 - 4 ($_OOoCalcStatus_InvalidValue)    = Invalid Value
;                  |                 - 5 ($_OOoCalcStatus_NoMatch)         = No Match
;                  |@extended        - Contains Invalid Parameter Number
; Author ........: GMK
; Modified ......:
; Remarks .......: None
; Related .......: None
; Link ..........: http://www.openoffice.org/api/docs/common/ref/com/sun/star/table/TableBorder.html
; Example .......: No
; ===============================================================================================================================
Func _OOoCalc_CreateBorders(ByRef $oObj, $vRangeOrRowStart, $iColStart = -1, $iRowEnd = -1, $iColEnd = -1, $vSheet = -1, $bTop = True, $bRight = True, $bBottom = True, $bLeft = True, $bHorizontal = False, $bVertical = False, $bDouble = False, $nLineWidth = 10.5)
	Local $oOOoCalc_COM_ErrorHandler = ObjEvent("AutoIt.Error", __OOoCalc_ComErrorHandler_InternalFunction)
    #forceref $oOOoCalc_COM_ErrorHandler
	If Not IsObj($oObj) Then Return SetError($_OOoCalcStatus_InvalidDataType, 1, 0)
	If Not IsInt($vRangeOrRowStart) And Not IsString($vRangeOrRowStart) Then Return SetError($_OOoCalcStatus_InvalidDataType, 2, 0)
	If IsString($vRangeOrRowStart) And Not __OOoCalc_RangeIsValid($vRangeOrRowStart) Then Return SetError($_OOoCalcStatus_InvalidValue, 2, 0)
	If Not IsInt($iColStart) Then Return SetError($_OOoCalcStatus_InvalidDataType, 3, 0)
	If IsInt($vRangeOrRowStart) And $iColStart < 1 Then Return SetError($_OOoCalcStatus_InvalidValue, 3, 0)
	If Not IsInt($iRowEnd) Then Return SetError($_OOoCalcStatus_InvalidDataType, 4, 0)
	If Not IsInt($iColEnd) Then Return SetError($_OOoCalcStatus_InvalidDataType, 5, 0)
	If Not IsInt($vSheet) And Not IsString($vSheet) Then Return SetError($_OOoCalcStatus_InvalidDataType, 6, 0)
	If IsString($vSheet) And Not __OOoCalc_WorksheetIsValid($oObj, $vSheet) Then Return SetError($_OOoCalcStatus_NoMatch, 6, 0)
	If Not IsBool($bTop) Then Return SetError($_OOoCalcStatus_InvalidDataType, 7, 0)
	If Not IsBool($bRight) Then Return SetError($_OOoCalcStatus_InvalidDataType, 8, 0)
	If Not IsBool($bBottom) Then Return SetError($_OOoCalcStatus_InvalidDataType, 9, 0)
	If Not IsBool($bBottom) Then Return SetError($_OOoCalcStatus_InvalidDataType, 10, 0)
	If Not IsBool($bHorizontal) Then Return SetError($_OOoCalcStatus_InvalidDataType, 11, 0)
	If Not IsBool($bVertical) Then Return SetError($_OOoCalcStatus_InvalidDataType, 12, 0)
	If Not IsBool($bDouble) Then Return SetError($_OOoCalcStatus_InvalidDataType, 13, 0)
	If Not IsNumber($nLineWidth) Then Return SetError($_OOoCalcStatus_InvalidDataType, 14, 0)
	Local $oSheet = __OOoCalc_GetSheet($oObj, $vSheet)
	If Not IsObj($oSheet) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	Local $oRange = __OOoCalc_GetRange($oSheet, $vRangeOrRowStart, $iColStart, $iRowEnd, $iColEnd)
	If Not IsObj($oRange) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	Local $oTableBorder = $oRange.TableBorder
	If Not IsObj($oTableBorder) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	Local $oTopLine
	If $bTop Then
		$oTopLine = __OOoCalc_CreateBorderLine($nLineWidth, $bDouble)
		If Not IsObj($oTopLine) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
		$oTableBorder.IsTopLineValid = 1
		$oTableBorder.TopLine = $oTopLine
	EndIf
	Local $oRightLine
	If $bRight Then
		$oRightLine = __OOoCalc_CreateBorderLine($nLineWidth, $bDouble)
		If Not IsObj($oRightLine) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
		$oTableBorder.IsRightLineValid = 1
		$oTableBorder.RightLine = $oRightLine
	EndIf
	Local $oBottomLine
	If $bBottom Then
		$oBottomLine = __OOoCalc_CreateBorderLine($nLineWidth, $bDouble)
		If Not IsObj($oBottomLine) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
		$oTableBorder.IsBottomLineValid = 1
		$oTableBorder.BottomLine = $oBottomLine
	EndIf
	Local $oLeftLine
	If $bLeft Then
		$oLeftLine = __OOoCalc_CreateBorderLine($nLineWidth, $bDouble)
		If Not IsObj($oLeftLine) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
		$oTableBorder.IsLeftLineValid = 1
		$oTableBorder.LeftLine = $oLeftLine
	EndIf
	Local $oHorizontalLine
	If $bHorizontal Then
		$oHorizontalLine = __OOoCalc_CreateBorderLine($nLineWidth, $bDouble)
		If Not IsObj($oHorizontalLine) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
		$oTableBorder.IsHorizontalLineValid = 1
		$oTableBorder.HorizontalLine = $oHorizontalLine
	EndIf
	Local $oVerticalLine
	If $bVertical Then
		$oVerticalLine = __OOoCalc_CreateBorderLine($nLineWidth, $bDouble)
		If Not IsObj($oVerticalLine) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
		$oTableBorder.IsVerticalLineValid = 1
		$oTableBorder.VerticalLine = $oVerticalLine
	EndIf
	If Not IsObj($oTableBorder) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	$oRange.TableBorder = $oTableBorder
	Return SetError($_OOoCalcStatus_Success, 0, 1)
EndFunc   ;==>_OOoCalc_CreateBorders

; #FUNCTION# ====================================================================================================================
; Name ..........: _OOoCalc_NumberFormat
; Description ...: Applies the specified formatting to the cells in the specified RC Range.
; Syntax ........: _OOoCalc_NumberFormat(ByRef $oObj, $iFormat, $vRangeOrRowStart[, $iColStart = -1[, $iRowEnd = -1[, $iColEnd = -1[, $vSheet = -1]]]])
; Parameters ....: $oObj             - Calc object opened by a preceding call to _OOoCalc_BookOpen, _OOoCalc_BookNew, or
;                                      _OOoCalc_BookAttach
;                  $iFormat          - The formatting string to apply to the specified range:
;                                       0 ($NUMBER_STANDARD)
;                                       1 ($NUMBER_INT)
;                                       2 ($NUMBER_DEC2)
;                                       3 ($NUMBER_1000INT)
;                                       4 ($NUMBER_1000DEC2)
;                                       5 ($NUMBER_SYSTEM)
;                                       6 ($SCIENTIFIC_000E000)
;                                       7 ($SCIENTIFIC_000E00)
;                                      10 ($PERCENT_INT)
;                                      11 ($PERCENT_DEC2)
;                                      12 ($FRACTION_1)
;                                      13 ($FRACTION_2)
;                                      20 ($CURRENCY_1000INT)
;                                      21 ($CURRENCY_1000DEC2)
;                                      22 ($CURRENCY_1000INT_RED)
;                                      23 ($CURRENCY_1000DEC2_RED)
;                                      24 ($CURRENCY_1000DEC2_CCC)
;                                      25 ($CURRENCY_1000DEC2_DASHED)
;                                      30 ($DATE_SYSTEM_SHORT)
;                                      31 ($DATE_DEF_NNDDMMMYY)
;                                      32 ($DATE_SYS_MMYY)
;                                      33 ($DATE_SYS_DDMMM)
;                                      34 ($DATE_MMMM)
;                                      35 ($DATE_QQJJ)
;                                      36 ($DATE_SYS_DDMMYYYY)
;                                      37 ($DATE_SYS_DDMMYY)
;                                      38 ($DATE_SYS_NNNNDMMMMYYYY)
;                                      39 ($DATE_SYS_DMMMYY)
;                                      40 ($TIME_HHMM)
;                                      41 ($TIME_HHMMSS)
;                                      42 ($TIME_HHMMAMPM)
;                                      43 ($TIME_HHMMSSAMPM)
;                                      44 ($TIME_HH_MMSS)
;                                      45 ($TIME_MMSS00)
;                                      46 ($TIME_HH_MMSS00)
;                                      50 ($DATETIME_SYSTEM_SHORT_HHMM)
;                  $vRangeOrRowStart - Either an A1 range, or index (0-based) of row to which to write if using RC
;                  $iColStart        - [optional] The index (0-based) of the column to which to write if using RC.
;                                                 Default is -1.
;                  $iRowEnd          - [optional] Index (0-based) of ending row to which to write if using RC.
;                                                 Default is -1.
;                  $iColEnd          - [optional] Index (0-based) of ending column to which to write if using RC.
;                                                 Default is -1.
;                  $vSheet           - [optional] Worksheet, either by index (0-based) or name.
;                                                 Default is -1, which would use the active worksheet.
; Return values .: On Success        - Returns 1
;                  On Failure        - Returns 0 and sets @error:
;                  |@error           - 0 ($_OOoCalcStatus_Success)         = No Error
;                  |                 - 3 ($_OOoCalcStatus_InvalidDataType) = Invalid Data Type
;                  |                 - 4 ($_OOoCalcStatus_InvalidValue)    = Invalid Value
;                  |                 - 5 ($_OOoCalcStatus_NoMatch)         = No Match
;                  |@extended        - Contains Invalid Parameter Number
; Author ........: GMK
; Modified ......:
; Remarks .......: None
; Related .......: None
; Link ..........: http://www.openoffice.org/api/docs/common/ref/com/sun/star/text/CellRange.html#NumberFormat
; Example .......: No
; ===============================================================================================================================
Func _OOoCalc_NumberFormat(ByRef $oObj, $iFormat, $vRangeOrRowStart, $iColStart = -1, $iRowEnd = -1, $iColEnd = -1, $vSheet = -1)
	Local $oOOoCalc_COM_ErrorHandler = ObjEvent("AutoIt.Error", __OOoCalc_ComErrorHandler_InternalFunction)
    #forceref $oOOoCalc_COM_ErrorHandler
	If Not IsObj($oObj) Then Return SetError($_OOoCalcStatus_InvalidDataType, 1, 0)
	If Not IsInt($iFormat) Then Return SetError($_OOoCalcStatus_InvalidDataType, 2, 0)
	If $iFormat < 0 Or $iFormat > 50 Then Return SetError($_OOoCalcStatus_InvalidValue, 2, 0)
	If Not IsInt($vRangeOrRowStart) And Not __OOoCalc_RangeIsValid($vRangeOrRowStart) Then Return SetError($_OOoCalcStatus_InvalidValue, 3, 0)
	If Not IsInt($iColStart) Then Return SetError($_OOoCalcStatus_InvalidDataType, 4, 0)
	If IsInt($vRangeOrRowStart) And $iColStart < 1 Then Return SetError($_OOoCalcStatus_InvalidValue, 4, 0)
	If Not IsInt($iRowEnd) Then Return SetError($_OOoCalcStatus_InvalidDataType, 5, 0)
	If Not IsInt($iColEnd) Then Return SetError($_OOoCalcStatus_InvalidDataType, 6, 0)
	If Not IsInt($vSheet) And Not IsString($vSheet) Then Return SetError($_OOoCalcStatus_InvalidDataType, 7, 0)
	If $vSheet > -1 And Not __OOoCalc_WorksheetIsValid($oObj, $vSheet) Then Return SetError($_OOoCalcStatus_NoMatch, 7, 0)
	Local $oSheet = __OOoCalc_GetSheet($oObj, $vSheet)
	If Not IsObj($oSheet) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	Local $oRange = __OOoCalc_GetRange($oSheet, $vRangeOrRowStart, $iColStart, $iRowEnd, $iColEnd)
	If Not IsObj($oRange) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	$oRange.NumberFormat = $iFormat
	Return SetError($_OOoCalcStatus_Success, 0, 1)
EndFunc   ;==>_OOoCalc_NumberFormat

; #FUNCTION# ====================================================================================================================
; Name ..........: _OOoCalc_ReadCell
; Description ...: Read information from the specified worksheet cell
; Syntax ........: _OOoCalc_ReadCell(ByRef $oObj, $vRangeOrRow[, $iCol = -1[, $vSheet = -1]])
; Parameters ....: $oObj        - Calc object opened by a preceding call to _OOoCalc_BookOpen, _OOoCalc_BookNew, or
;                                 _OOoCalc_BookAttach
;                  $vRangeOrRow - Either an A1 range, or index (0-based) of row to which to write if using RC
;                  $iCol        - [optional] The index (0-based) of the column to which to write if using RC.
;                                            Default is -1.
;                  $vSheet      - [optional] Worksheet, either by index (0-based) or name.
;                                            Default is -1, which would use the active worksheet.
; Return values .: On Success   - Returns the data from the specified cell
;                  On Failure   - Returns 0 and sets @error:
;                  |@error      - 0 ($_OOoCalcStatus_Success)         = No Error
;                  |            - 3 ($_OOoCalcStatus_InvalidDataType) = Invalid Data Type
;                  |            - 4 ($_OOoCalcStatus_InvalidValue)    = Invalid Value
;                  |            - 5 ($_OOoCalcStatus_NoMatch)         = No Match
;                  |@extended   - Contains Invalid Parameter Number
; Author ........: Leagnus, GMK
; Modified ......:
; Remarks .......: None
; Related .......: _OOoCalc_WriteCell
; Link ..........: http://www.openoffice.org/api/docs/common/ref/com/sun/star/table/XCell.html#getValue
; Example .......: No
; ===============================================================================================================================
Func _OOoCalc_ReadCell(ByRef $oObj, $vRangeOrRow, $iCol = -1, $vSheet = -1, $iFormulaOrValue = 1)
	Local $oOOoCalc_COM_ErrorHandler = ObjEvent("AutoIt.Error", __OOoCalc_ComErrorHandler_InternalFunction)
    #forceref $oOOoCalc_COM_ErrorHandler
	If Not IsObj($oObj) Then Return SetError($_OOoCalcStatus_InvalidDataType, 1, 0)
	If Not IsInt($vRangeOrRow) And Not __OOoCalc_RangeIsValid($vRangeOrRow) Then Return SetError($_OOoCalcStatus_InvalidValue, 2, 0)
	If Not IsInt($iCol) Then Return SetError($_OOoCalcStatus_InvalidDataType, 3, 0)
	If IsInt($vRangeOrRow) And $iCol < 0 Then Return SetError($_OOoCalcStatus_InvalidValue, 3, 0)
	If Not IsInt($vSheet) And Not IsString($vSheet) Then Return SetError($_OOoCalcStatus_InvalidDataType, 4, 0)
	If $vSheet > -1 And Not __OOoCalc_WorksheetIsValid($oObj, $vSheet) Then Return SetError($_OOoCalcStatus_NoMatch, 4, 0)
	Local $oSheet = __OOoCalc_GetSheet($oObj, $vSheet)
	If Not IsObj($oSheet) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	Local $oCell = __OOoCalc_GetCell($oSheet, $vRangeOrRow, $iCol)
	If Not IsObj($oCell) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	Local $vReturn
	Switch $iFormulaOrValue
		Case 0
			$vReturn = $oCell.getFormula
		Case 1
			$vReturn = $oCell.getstring
	EndSwitch
	Return SetError($_OOoCalcStatus_Success, 0, $vReturn)
EndFunc   ;==>_OOoCalc_ReadCell

; #FUNCTION# ====================================================================================================================
; Name ..........: _OOoCalc_ReadSheetToArray
; Description ...: Creates an array from specified range of specified  worksheet.
; Syntax ........: _OOoCalc_ReadSheetToArray(ByRef $oObj[, $vRangeOrRowStart = -1[, $iColStart = -1[, $iRowEnd = -1[, $iColEnd = -1[, $vSheet = -1]]]]])
; Parameters ....: $oObj             - Calc object opened by a preceding call to _OOoCalc_BookOpen, _OOoCalc_BookNew, or
;                                      _OOoCalc_BookAttach
;                  $vRangeOrRowStart - [optional] Range to which to write array. If default value of -1 is used, the used range will be
;                                      used.
;                  $iColStart        - [optional] Index (0-based) of the starting column, if using RC.
;                                                 Default is -1.
;                  $iRowEnd          - [optional] Ending row, if using RC.
;                                                 Default is -1.
;                  $iColEnd          - [optional] Index (0-based) of the ending column, if using RC.
;                                                 Default is -1.
;                  $vSheet           - [optional] Worksheet, either by index (0-based) or name.
;                                                 Default is -1, which would use the active worksheet.
; Return values .: On Success        - Returns 2D array
;                  On Failure        - Returns 0 and sets @error:
;                  |@error           - 0 ($_OOoCalcStatus_Success)         = No Error
;                  |                 - 3 ($_OOoCalcStatus_InvalidDataType) = Invalid Data Type
;                  |                 - 4 ($_OOoCalcStatus_InvalidValue)    = Invalid Value
;                  |                 - 5 ($_OOoCalcStatus_NoMatch)         = No Match
;                  |@extended        - Contains Invalid Parameter Number
; Author ........: Andy G, GMK
; Modified ......:
; Remarks .......: None
; Related .......: _OOoCalc_CellRead, _OOoCalc_WriteFromArray
; Link ..........: http://www.openoffice.org/api/docs/common/ref/com/sun/star/sheet/XCellRangeData.html#getDataArray
; Example .......: No
; ===============================================================================================================================
Func _OOoCalc_ReadSheetToArray(ByRef $oObj, $vRangeOrRowStart = -1, $iColStart = -1, $iRowEnd = -1, $iColEnd = -1, $vSheet = -1)
	Local $oOOoCalc_COM_ErrorHandler = ObjEvent("AutoIt.Error", __OOoCalc_ComErrorHandler_InternalFunction)
    #forceref $oOOoCalc_COM_ErrorHandler
	If Not IsObj($oObj) Then Return SetError($_OOoCalcStatus_InvalidDataType, 1, 0)
	If Not IsInt($vRangeOrRowStart) And Not __OOoCalc_RangeIsValid($vRangeOrRowStart) Then Return SetError($_OOoCalcStatus_InvalidValue, 2, 0)
	If Not IsInt($iColStart) Then Return SetError($_OOoCalcStatus_InvalidDataType, 3, 0)
	If $vRangeOrRowStart < -1 And IsInt($vRangeOrRowStart) And $iColStart < 0 Then Return SetError($_OOoCalcStatus_InvalidValue, 3, 0)
	If Not IsInt($iRowEnd) Then Return SetError($_OOoCalcStatus_InvalidDataType, 4, 0)
	If Not IsInt($iColEnd) Then Return SetError($_OOoCalcStatus_InvalidDataType, 5, 0)
	If Not IsInt($vSheet) And Not IsString($vSheet) Then Return SetError($_OOoCalcStatus_InvalidDataType, 6, 0)
	If $vSheet > -1 And Not __OOoCalc_WorksheetIsValid($oObj, $vSheet) Then Return SetError($_OOoCalcStatus_NoMatch, 6, 0)
	Local $oSheet = __OOoCalc_GetSheet($oObj, $vSheet)
	If Not IsObj($oSheet) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	Local $oRange = __OOoCalc_GetRange($oSheet, $vRangeOrRowStart, $iColStart, $iRowEnd, $iColEnd)
	If Not IsObj($oRange) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	Local $avData = $oRange.getDataArray
	Local $iRows = UBound($avData)
	Local $iLastRow = $iRows - 1
	Local $iCols = UBound($avData[0])
	Local $iLastCol = $iCols - 1
	Local $avReturn[$iRows][$iCols]
	Local $aRow
	For $iRow = 0 To $iLastRow
		$aRow = $avData[$iRow]
		$iCols = UBound($aRow)
		$iLastCol = $iCols - 1
		For $iCol = 0 To $iLastCol
			$avReturn[$iRow][$iCol] = $aRow[$iCol]
		Next
	Next
	Return SetError($_OOoCalcStatus_Success, 0, $avReturn)
EndFunc   ;==>_OOoCalc_ReadSheetToArray

; #FUNCTION# ====================================================================================================================
; Name ..........: _OOoCalc_RowDelete
; Description ...: Delete a number of rows from the specified worksheet.
; Syntax ........: _OOoCalc_RowDelete(ByRef $oObj, $iRow[, $iNumRows = 1[, $vSheet = -1]])
; Parameters ....: $oObj      - Calc object opened by a preceding call to _OOoCalc_BookOpen, _OOoCalc_BookNew, or
;                               _OOoCalc_BookAttach
;                  $iRow      - The index (0-based) of the row to delete
;                  $iNumRows  - [optional] The number of rows to delete.
;                                          Default is 1.
;                  $vSheet    - [optional] Worksheet, either by index (0-based) or name.
;                                          Default is -1, which would use the active worksheet.
; Return values .: On Success - Returns 1
;                  On Failure - Returns 0 and sets @error:
;                  |@error    - 0 ($_OOoCalcStatus_Success)         = No Error
;                  |          - 3 ($_OOoCalcStatus_InvalidDataType) = Invalid Data Type
;                  |          - 5 ($_OOoCalcStatus_NoMatch)         = No Match
;                  |@extended - Contains Invalid Parameter Number
; Author ........: Andy G, GMK, mLipok
; Modified ......:
; Remarks .......: None
; Related .......: _OOoCalc_RowInsert
; Link ..........: http://www.openoffice.org/api/docs/common/ref/com/sun/star/table/XTableRows.html#removeByIndex
; Example .......: No
; ===============================================================================================================================
Func _OOoCalc_RowDelete(ByRef $oObj, $iRow, $iNumRows = 1, $vSheet = -1)
	Local $oOOoCalc_COM_ErrorHandler = ObjEvent("AutoIt.Error", __OOoCalc_ComErrorHandler_InternalFunction)
    #forceref $oOOoCalc_COM_ErrorHandler
	If Not IsObj($oObj) Then Return SetError($_OOoCalcStatus_InvalidDataType, 1, 0)
	If Not IsInt($iRow) Then Return SetError($_OOoCalcStatus_InvalidDataType, 2, 0)
	If Not IsInt($iNumRows) Then Return SetError($_OOoCalcStatus_InvalidDataType, 3, 0)
	If Not IsInt($vSheet) And Not IsString($vSheet) Then Return SetError($_OOoCalcStatus_InvalidDataType, 4, 0)
	If $vSheet > -1 And Not __OOoCalc_WorksheetIsValid($oObj, $vSheet) Then Return SetError($_OOoCalcStatus_NoMatch, 4, 0)
	Local $oSheet = __OOoCalc_GetSheet($oObj, $vSheet)
	If Not IsObj($oSheet) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	If $iRow < 0 Then Return SetError($_OOoCalcStatus_InvalidValue, 2, 0)
	$oSheet.getRows.removeByIndex($iRow, $iNumRows)
	Return SetError($_OOoCalcStatus_Success, 0, 1)
EndFunc   ;==>_OOoCalc_RowDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _OOoCalc_ColumnDelete
; Description ...: Delete a number of Columns from the specified worksheet.
; Syntax ........: _OOoCalc_ColumnDelete(ByRef $oObj, $iCol[, $iNumColumns = 1[, $vSheet = -1]])
; Parameters ....: $oObj        - Calc object opened by a preceding call to _OOoCalc_BookOpen, _OOoCalc_BookNew, or
;                                 _OOoCalc_BookAttach
;                  $iCol        - The index (0-based) of the column to delete
;                  $iNumColumns - [optional] The number of columns to delete.
;                                            Default is 1.
;                  $vSheet      - [optional] Worksheet, either by index (0-based) or name.
;                                            Default is -1, which would use the active worksheet.
; Return values .: On Success   - Returns 1
;                  On Failure   - Returns 0 and sets @error:
;                  |@error      - 0 ($_OOoCalcStatus_Success)         = No Error
;                  |            - 3 ($_OOoCalcStatus_InvalidDataType) = Invalid Data Type
;                  |            - 5 ($_OOoCalcStatus_NoMatch)         = No Match
;                  |@extended   - Contains Invalid Parameter Number
; Author ........: Andy G, GMK
; Modified ......:
; Remarks .......: None
; Related .......: _OOoCalc_ColumnInsert
; Link ..........: http://www.openoffice.org/api/docs/common/ref/com/sun/star/table/XTableColumns.html#removeByIndex
; Example .......: No
; ===============================================================================================================================
Func _OOoCalc_ColumnDelete(ByRef $oObj, $iCol, $iNumColumns = 1, $vSheet = -1)
	Local $oOOoCalc_COM_ErrorHandler = ObjEvent("AutoIt.Error", __OOoCalc_ComErrorHandler_InternalFunction)
    #forceref $oOOoCalc_COM_ErrorHandler
	If Not IsObj($oObj) Then Return SetError($_OOoCalcStatus_InvalidDataType, 1, 0)
	If Not IsInt($iCol) Then Return SetError($_OOoCalcStatus_InvalidDataType, 2, 0)
	If Not IsInt($iNumColumns) Then Return SetError($_OOoCalcStatus_InvalidDataType, 3, 0)
	If Not IsInt($vSheet) And Not IsString($vSheet) Then Return SetError($_OOoCalcStatus_InvalidDataType, 4, 0)
	If $vSheet > -1 And Not __OOoCalc_WorksheetIsValid($oObj, $vSheet) Then Return SetError($_OOoCalcStatus_NoMatch, 4, 0)
	Local $oSheet = __OOoCalc_GetSheet($oObj, $vSheet)
	If Not IsObj($oSheet) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	If $iCol < 0 Then Return SetError($_OOoCalcStatus_InvalidValue, 2, 0)
	$oSheet.getColumns.removeByIndex($iCol, $iNumColumns)
	Return SetError($_OOoCalcStatus_Success, 0, 1)
EndFunc   ;==>_OOoCalc_ColumnDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _OOoCalc_RowInsert
; Description ...: Insert a number of rows from the specified worksheet.
; Syntax ........: _OOoCalc_RowInsert(ByRef $oObj, $iRow[, $iNumRows = 1[, $vSheet = -1]])
; Parameters ....: $oObj      - Calc object opened by a preceding call to _OOoCalc_BookOpen, _OOoCalc_BookNew, or
;                               _OOoCalc_BookAttach
;                  $iRow      - The index (0-based) of the row to insert
;                  $iNumRows  - [optional] The number of rows to insert.
;                                          Default is 1.
;                  $vSheet    - [optional] Worksheet, either by index (0-based) or name.
;                                          Default is -1, which would use the active worksheet.
; Return values .: On Success - Returns 1
;                  On Failure - Returns 0 and sets @error:
;                  |@error    - 0 ($_OOoCalcStatus_Success)         = No Error
;                  |          - 3 ($_OOoCalcStatus_InvalidDataType) = Invalid Data Type
;                  |          - 5 ($_OOoCalcStatus_NoMatch)         = No Match
;                  |@extended - Contains Invalid Parameter Number
; Author ........: Andy G, GMK, mLipok
; Modified ......:
; Remarks .......: None
; Related .......: _OOoCalc_RowInsert
; Link ..........: http://www.openoffice.org/api/docs/common/ref/com/sun/star/table/XTableRows.html#removeByIndex
; Example .......: No
; ===============================================================================================================================
Func _OOoCalc_RowInsert(ByRef $oObj, $iRow, $iNumRows = 1, $vSheet = -1)
	Local $oOOoCalc_COM_ErrorHandler = ObjEvent("AutoIt.Error", __OOoCalc_ComErrorHandler_InternalFunction)
    #forceref $oOOoCalc_COM_ErrorHandler
	If Not IsObj($oObj) Then Return SetError($_OOoCalcStatus_InvalidDataType, 1, 0)
	If Not IsInt($iRow) Then Return SetError($_OOoCalcStatus_InvalidDataType, 2, 0)
	If Not IsInt($iNumRows) Then Return SetError($_OOoCalcStatus_InvalidDataType, 3, 0)
	If Not IsInt($vSheet) And Not IsString($vSheet) Then Return SetError($_OOoCalcStatus_InvalidDataType, 4, 0)
	If $vSheet > -1 And Not __OOoCalc_WorksheetIsValid($oObj, $vSheet) Then Return SetError($_OOoCalcStatus_NoMatch, 4, 0)
	Local $oSheet = __OOoCalc_GetSheet($oObj, $vSheet)
	If Not IsObj($oSheet) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	If $iRow < 0 Then Return SetError($_OOoCalcStatus_InvalidValue, 2, 0)
	$oSheet.getRows.insertByIndex($iRow, $iNumRows)
	Return SetError($_OOoCalcStatus_Success, 0, 1)
EndFunc   ;==>_OOoCalc_RowInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _OOoCalc_ColumnInsert
; Description ...: Insert a number of Columns from the specified worksheet.
; Syntax ........: _OOoCalc_ColumnInsert(ByRef $oObj, $iCol[, $iNumColumns = 1[, $vSheet = -1]])
; Parameters ....: $oObj        - Calc object opened by a preceding call to _OOoCalc_BookOpen, _OOoCalc_BookNew, or
;                                 _OOoCalc_BookAttach
;                  $iCol        - The index (0-based) of the column to insert
;                  $iNumColumns - [optional] The number of Columns to insert.
;                                            Default is 1.
;                  $vSheet      - [optional] Worksheet, either by index (0-based) or name.
;                                            Default is -1, which would use the active worksheet.
; Return values .: On Success   - Returns 1
;                  On Failure   - Returns 0 and sets @error:
;                  |@error      - 0 ($_OOoCalcStatus_Success)         = No Error
;                  |            - 3 ($_OOoCalcStatus_InvalidDataType) = Invalid Data Type
;                  |            - 5 ($_OOoCalcStatus_NoMatch)         = No Match
;                  |@extended   - Contains Invalid Parameter Number
; Author ........: Andy G, GMK, mLipok
; Modified ......:
; Remarks .......: None
; Related .......: _OOoCalc_ColumnInsert
; Link ..........: http://www.openoffice.org/api/docs/common/ref/com/sun/star/table/XTableColumns.html#removeByIndex
; Example .......: No
; ===============================================================================================================================
Func _OOoCalc_ColumnInsert(ByRef $oObj, $iCol, $iNumColumns = 1, $vSheet = -1)
	Local $oOOoCalc_COM_ErrorHandler = ObjEvent("AutoIt.Error", __OOoCalc_ComErrorHandler_InternalFunction)
    #forceref $oOOoCalc_COM_ErrorHandler
	If Not IsObj($oObj) Then Return SetError($_OOoCalcStatus_InvalidDataType, 1, 0)
	If Not IsInt($iCol) Then Return SetError($_OOoCalcStatus_InvalidDataType, 2, 0)
	If Not IsInt($iNumColumns) Then Return SetError($_OOoCalcStatus_InvalidDataType, 3, 0)
	If Not IsInt($vSheet) And Not IsString($vSheet) Then Return SetError($_OOoCalcStatus_InvalidDataType, 4, 0)
	If $vSheet > -1 And Not __OOoCalc_WorksheetIsValid($oObj, $vSheet) Then Return SetError($_OOoCalcStatus_NoMatch, 4, 0)
	Local $oSheet = __OOoCalc_GetSheet($oObj, $vSheet)
	If Not IsObj($oSheet) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	If $iCol < 0 Then Return SetError($_OOoCalcStatus_InvalidValue, 2, 0)
	$oSheet.getColumns.insertByIndex($iCol, $iNumColumns)
	Return SetError($_OOoCalcStatus_Success, 0, 1)
EndFunc   ;==>_OOoCalc_ColumnInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _OOoCalc_SheetAddNew
; Description ...: Add new sheet to workbook - optionally with a name.
; Syntax ........: _OOoCalc_SheetAddNew(ByRef $oObj[, $sName = ''[, $iPos = '']])
; Parameters ....: $oObj      - Calc object opened by a preceding call to _OOoCalc_BookOpen, _OOoCalc_BookNew, or
;                                _OOoCalc_BookAttach
;                  $sSheet    - [optional] Name of the sheet to be added.
;                                          Default is ''.
;                  $iPos      - [optional] Number position (0-based index) of the sheet to be added.
;                                          Default is -1, which will put the sheet at the end.
; Return values .: On Success - Returns 1
;                  On Failure - Returns 0 and sets @error:
;                  |@error    - 0 ($_OOoCalcStatus_Success)         = No Error
;                  |          - 3 ($_OOoCalcStatus_InvalidDataType) = Invalid Data Type
;                  |          - 4 ($_OOoCalcStatus_InvalidValue)    = Invalid Value
;                  |@extended - Contains Invalid Parameter Number
; Author ........: Andy G, GMK, mLipok
; Modified ......:
; Remarks .......: If no name is specified, a generic name is provided, such as 'Sheet4'
; Related .......: _OOoCalc_SheetDelete
; Link ..........: http://www.openoffice.org/api/docs/common/ref/com/sun/star/container/XNameContainer.html#insertByName
; Example .......: No
; ===============================================================================================================================
Func _OOoCalc_SheetAddNew(ByRef $oObj, $sSheet = '', $iPos = -1)
	Local $oOOoCalc_COM_ErrorHandler = ObjEvent("AutoIt.Error", __OOoCalc_ComErrorHandler_InternalFunction)
    #forceref $oOOoCalc_COM_ErrorHandler
	If Not IsObj($oObj) Then Return SetError($_OOoCalcStatus_InvalidDataType, 1, 0)
	If Not IsString($sSheet) Then Return SetError($_OOoCalcStatus_InvalidDataType, 2, 0)
	If Not IsInt($iPos) Then Return SetError($_OOoCalcStatus_InvalidDataType, 3, 0)
	Local $iCount = $oObj.GetSheets.count
	If $sSheet = '' Then $sSheet = 'Sheet' & $iCount + 1
	If $iPos < 0 Then $iPos = $iCount
	Local $oSheets = $oObj.getSheets.createEnumeration
	If Not IsObj($oSheets) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	Local $bFlag = False, $oElement
	While $oSheets.hasMoreElements
		$oElement = $oSheets.nextElement
		If $oElement.Name = $sSheet Then
			$bFlag = True
			ExitLoop
		EndIf
	WEnd
    If $bFlag Then Return SetError($_OOoCalcStatus_InvalidValue, 2, 0)
    $oObj.Sheets.insertNewByName($sSheet, $iPos)
    Return SetError($_OOoCalcStatus_Success, 0, 1)
EndFunc   ;==>_OOoCalc_SheetAddNew

; #FUNCTION# ====================================================================================================================
; Name ..........: _OOoCalc_SheetDelete
; Description ...: Delete the specified sheet by string name or by index.
; Syntax ........: _OOoCalc_SheetDelete(ByRef $oObj, $vSheet)
; Parameters ....: $oObj      - Calc object opened by a preceding call to _OOoCalc_BookOpen, _OOoCalc_BookNew, or
;                               _OOoCalc_BookAttach
;                  $vSheet    - Worksheet by index (0-based) or name
; Return values .: On Success - Returns 1
;                  On Failure - Returns 0 and sets @error:
;                  |@error    - 0 ($_OOoCalcStatus_Success)         = No Error
;                  |          - 3 ($_OOoCalcStatus_InvalidDataType) = Invalid Data Type
;                  |          - 5 ($_OOoCalcStatus_NoMatch)         = No Match
;                  |@extended - Contains Invalid Parameter Number
; Author ........: GMK
; Modified ......:
; Remarks .......: None
; Related .......: _OOoCalc_SheetAddNew, _OOoCalc_SheetActivate
; Link ..........: http://www.openoffice.org/api/docs/common/ref/com/sun/star/container/XNameContainer.html#removeByName
; Example .......: No
; ===============================================================================================================================
Func _OOoCalc_SheetDelete(ByRef $oObj, $vSheet)
	Local $oOOoCalc_COM_ErrorHandler = ObjEvent("AutoIt.Error", __OOoCalc_ComErrorHandler_InternalFunction)
    #forceref $oOOoCalc_COM_ErrorHandler
	If Not IsObj($oObj) Then Return SetError($_OOoCalcStatus_InvalidDataType, 1, 0)
	If Not IsInt($vSheet) And Not IsString($vSheet) Then Return SetError($_OOoCalcStatus_InvalidDataType, 2, 0)
	If Not __OOoCalc_WorksheetIsValid($oObj, $vSheet) Then Return SetError($_OOoCalcStatus_NoMatch, 2, 0)
	If IsInt($vSheet) Then $vSheet = $oObj.getSheets.getByIndex($vSheet - 1).Name
	$oObj.getSheets.removeByName($vSheet)
	Return SetError($_OOoCalcStatus_Success, 0, 1)
EndFunc   ;==>_OOoCalc_SheetDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _OOoCalc_SheetNameGet
; Description ...: Return the name of the active sheet.
; Syntax ........: _OOoCalc_SheetNameGet(ByRef $oObj)
; Parameters ....: $oObj     - Calc object opened by a preceding call to _OOoCalc_BookOpen, _OOoCalc_BookNew, or
;                               _OOoCalc_BookAttach
; Return values .: On Success - Returns the name of the active sheet (string)
;                  On Failure - Returns 0 and sets @error:
;                  |@error    - 0 ($_OOoCalcStatus_Success)         = No Error
;                  |          - 3 ($_OOoCalcStatus_InvalidDataType) = Invalid Data Type
; Author ........: Andy G
; Modified ......:
; Remarks .......: None
; Related .......: _OOoCalc_SheetNameSet, _OOoCalc_SheetList
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _OOoCalc_SheetNameGet(ByRef $oObj)
	Local $oOOoCalc_COM_ErrorHandler = ObjEvent("AutoIt.Error", __OOoCalc_ComErrorHandler_InternalFunction)
    #forceref $oOOoCalc_COM_ErrorHandler
	If Not IsObj($oObj) Then Return SetError($_OOoCalcStatus_InvalidDataType, 1, 0)
	Local $sReturn = $oObj.CurrentController.ActiveSheet.name
	Return SetError($_OOoCalcStatus_Success, 0, $sReturn)
EndFunc   ;==>_OOoCalc_SheetNameGet

; #FUNCTION# ====================================================================================================================
; Name ..........: _OOoCalc_SheetNameSet
; Description ...: Set the name of the active sheet.
; Syntax ........: _OOoCalc_SheetNameSet(ByRef $oObj, $sSheetName)
; Parameters ....: $oObj       - Calc object opened by a preceding call to _OOoCalc_BookOpen, _OOoCalc_BookNew, or
;                                _OOoCalc_BookAttach
;                  $sSheetName - Sheet name
; Return values .: On Success  - Returns the name of the active sheet (string)
;                  On Failure  - Returns 0 and sets @error:
;                  |@error     - 0 ($_OOoCalcStatus_Success)         = No Error
;                  |           - 3 ($_OOoCalcStatus_InvalidDataType) = Invalid Data Type
;                  |@extended  - Contains Invalid Parameter Number
; Author ........: Andy G
; Modified ......:
; Remarks .......: None
; Related .......: _OOoCalc_SheetNameGet, _OOoCalc_SheetList
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _OOoCalc_SheetNameSet(ByRef $oObj, $sSheetName)
	Local $oOOoCalc_COM_ErrorHandler = ObjEvent("AutoIt.Error", __OOoCalc_ComErrorHandler_InternalFunction)
    #forceref $oOOoCalc_COM_ErrorHandler
	If Not IsObj($oObj) Then Return SetError($_OOoCalcStatus_InvalidDataType, 1, 0)
	If Not IsString($sSheetName) Then Return SetError($_OOoCalcStatus_InvalidDataType, 2, 0)
	$oObj.CurrentController.ActiveSheet.name = $sSheetName
	SetError($_OOoCalcStatus_Success, 0, 1)
EndFunc   ;==>_OOoCalc_SheetNameSet

; #FUNCTION# ====================================================================================================================
; Name ..........: _OOoCalc_SheetList
; Description ...: Returns a list of all sheets in workbook, by name, as an array.
; Syntax ........: _OOoCalc_SheetList(ByRef $oObj)
; Parameters ....: $oObj      - Calc object opened by a preceding call to _OOoCalc_BookOpen, _OOoCalc_BookNew, or
;                               _OOoCalc_BookAttach
; Return values .: On Success - Returns an array of the sheet names in the workbook with the first entry of the array indicating
;                               the array size
;                  On Failure - Returns 0 and sets @error:
;                  |@error    - 0 ($_OOoCalcStatus_Success)         = No Error
;                  |          - 3 ($_OOoCalcStatus_InvalidDataType) = Invalid Data Type
; Author ........: Leagnus, Kris, GMK
; Modified ......:
; Remarks .......: None
; Related .......: None
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _OOoCalc_SheetList(ByRef $oObj)
	Local $oOOoCalc_COM_ErrorHandler = ObjEvent("AutoIt.Error", __OOoCalc_ComErrorHandler_InternalFunction)
    #forceref $oOOoCalc_COM_ErrorHandler
	If Not IsObj($oObj) Then Return SetError($_OOoCalcStatus_InvalidDataType, 1, 0)
	Local $asReturn[1] = [$oObj.getSheets.count]
	Local $asWorksheets = $oObj.getSheets.ElementNames
	Local $iSheets = UBound($asWorksheets)
	ReDim $asReturn[$iSheets + 1]
	Local $iLastRow = UBound($asReturn) - 1
	For $iRow = 1 To $iLastRow
		$asReturn[$iRow] = $asWorksheets[$iRow - 1]
	Next
	Return SetError($_OOoCalcStatus_Success, 0, $asReturn)
EndFunc   ;==>_OOoCalc_SheetList

; #FUNCTION# ====================================================================================================================
; Name ..........: _OOoCalc_SheetActivate
; Description ...: Activate the specified sheet by string name
; Syntax ........: _OOoCalc_SheetActivate(ByRef $oObj, $vSheet)
; Parameters ....: $oObj     - Calc object opened by a preceding call to _OOoCalc_BookOpen, _OOoCalc_BookNew, or
;                               _OOoCalc_BookAttach
;                  $vSheet    - Worksheet, either by index (0-based) or name
; Return values .: On Success - Returns 1
;                  On Failure - Returns 0 and sets @error:
;                  |@error    - 0 ($_OOoCalcStatus_Success)         = No Error
;                  |          - 3 ($_OOoCalcStatus_InvalidDataType) = Invalid Data Type
;                  |          - 5 ($_OOoCalcStatus_NoMatch)         = No Match
;                  |@extended - Contains Invalid Parameter Number
; Author ........: Leagnus, Kris, GMK
; Modified ......:
; Remarks .......: Made the selection procedure more elegant (nothing gets selected).  This function is only for visual aid for
;                  user and is useless for other funcs: _OOoCalc_FindInRange, _OOoCalc_ReadCell, _OOoCalc_WriteCell, etc.
; Related .......: None
; Link ..........: http://www.openoffice.org/api/docs/common/ref/com/sun/star/sheet/XSpreadsheetView.html#setActiveSheet
; Example .......: No
; ===============================================================================================================================
Func _OOoCalc_SheetActivate(ByRef $oObj, $vSheet)
	Local $oOOoCalc_COM_ErrorHandler = ObjEvent("AutoIt.Error", __OOoCalc_ComErrorHandler_InternalFunction)
    #forceref $oOOoCalc_COM_ErrorHandler
	If Not IsObj($oObj) Then Return SetError($_OOoCalcStatus_InvalidDataType, 1, 0)
	If Not IsInt($vSheet) And Not IsString($vSheet) Then Return SetError($_OOoCalcStatus_InvalidDataType, 2, 0)
	If $vSheet > -1 And Not __OOoCalc_WorksheetIsValid($oObj, $vSheet) Then Return SetError($_OOoCalcStatus_NoMatch, 2, 0)
	Local $oSheet = __OOoCalc_GetSheet($oObj, $vSheet)
	If Not IsObj($oSheet) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	$oObj.getCurrentController.setActiveSheet($oSheet)
	Return SetError($_OOoCalcStatus_Success, 0, 1)
EndFunc   ;==>_OOoCalc_SheetActivate

; #FUNCTION# ====================================================================================================================
; Name ..........: _OOoCalc_SheetSetVisibility
; Description ...: Activate the specified sheet by string name
; Syntax ........: _OOoCalc_SheetSetVisibility(ByRef $oObj, $vSheet[, $bVisibility = True])
; Parameters ....: $oObj        - Calc object opened by a preceding call to _OOoCalc_BookOpen, _OOoCalc_BookNew, or
;                                 _OOoCalc_BookAttach
;                  $vSheet      - Worksheet, either by index (0-based) or name
;                  $iVisibility - Flag for visibility (1 = True or 0 = False). Default is 1.
; Return values .: On Success   - Returns 1
;                  On Failure   - Returns 0 and sets @error:
;                  |@error      - 0 ($_OOoCalcStatus_Success)         = No Error
;                  |            - 3 ($_OOoCalcStatus_InvalidDataType) = Invalid Data Type
;                  |            - 5 ($_OOoCalcStatus_NoMatch)         = No Match
;                  |@extended   - Contains Invalid Parameter Number
; Author ........: GMK
; Modified ......:
; Remarks .......:
; Related .......: None
; Link ..........: None
; Example .......: No
; ===============================================================================================================================
Func _OOoCalc_SheetSetVisibility(ByRef $oObj, $vSheet, $iVisibility = 1)
	Local $oOOoCalc_COM_ErrorHandler = ObjEvent("AutoIt.Error", __OOoCalc_ComErrorHandler_InternalFunction)
    #forceref $oOOoCalc_COM_ErrorHandler
	If Not IsObj($oObj) Then Return SetError($_OOoCalcStatus_InvalidDataType, 1, 0)
	If Not IsInt($vSheet) And Not IsString($vSheet) Then Return SetError($_OOoCalcStatus_InvalidDataType, 2, 0)
	If $vSheet > -1 And Not __OOoCalc_WorksheetIsValid($oObj, $vSheet) Then Return SetError($_OOoCalcStatus_NoMatch, 2, 0)
	If Not IsInt($iVisibility) Then Return SetError($_OOoCalcStatus_InvalidDataType, 3, 0)
	If $iVisibility <> 0 And $iVisibility <> 1 Then Return SetError($_OOoCalcStatus_InvalidValue, 3, 0)
	Local $oSheet = __OOoCalc_GetSheet($oObj, $vSheet)
	If Not IsObj($oSheet) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	$oSheet.IsVisible = $iVisibility
	Return SetError($_OOoCalcStatus_Success, 0, 1)
EndFunc   ;==>_OOoCalc_SheetSetVisibility

; #FUNCTION# ====================================================================================================================
; Name ..........: _OOoCalc_SheetMove
; Description ...: Move the specified sheet before another specified position.
; Syntax ........: _OOoCalc_SheetMove(ByRef $oObj, $sSheet, $iPosition)
; Parameters ....: $oObj      - Calc object opened by a preceding call to _OOoCalc_BookOpen, _OOoCalc_BookNew, or
;                               _OOoCalc_BookAttach
;                  $iPosition - Position in which to move sheet.
;                  $vSheet    - [optional] Worksheet, either by index (0-based) or name.
;                                          Default is -1, which would use the active worksheet.
; Return values .: On Success - Returns 1
;                  On Failure - Returns 0 and sets @error:
;                  |@error    - 0 ($_OOoCalcStatus_Success)         = No Error
;                  |          - 3 ($_OOoCalcStatus_InvalidDataType) = Invalid Data Type
;                  |          - 5 ($_OOoCalcStatus_NoMatch)         = No Match
;                  |@extended - Contains Invalid Parameter Number
; Author ........: GMK
; Modified ......:
; Remarks .......: None
; Related .......: _OOoCalc_SheetAddNew, _OOoCalc_SheetDelete, _OOoCalc_SheetNameGet, _OOoCalc_SheetNameSet, _OOoCalc_SheetList, _OOoCalc_SheetActivate
; Link ..........: http://www.openoffice.org/api/docs/common/ref/com/sun/star/sheet/XSpreadsheets.html#moveByName
; Example .......: No
; ===============================================================================================================================
Func _OOoCalc_SheetMove(ByRef $oObj, $iPosition, $vSheet = -1)
	Local $oOOoCalc_COM_ErrorHandler = ObjEvent("AutoIt.Error", __OOoCalc_ComErrorHandler_InternalFunction)
    #forceref $oOOoCalc_COM_ErrorHandler
	If Not IsObj($oObj) Then Return SetError($_OOoCalcStatus_InvalidDataType, 1, 0)
	If Not IsInt($iPosition) Then Return SetError($_OOoCalcStatus_InvalidDataType, 2, 0)
	If Not IsInt($vSheet) And Not IsString($vSheet) Then Return SetError($_OOoCalcStatus_InvalidDataType, 3, 0)
	If $vSheet > -1 And Not __OOoCalc_WorksheetIsValid($oObj, $vSheet) Then Return SetError($_OOoCalcStatus_NoMatch, 3, 0)
	If IsInt($vSheet) Then $vSheet = ($vSheet < 0) ? (_OOoCalc_SheetNameGet($oObj)) : ($oObj.getSheets.getByIndex($vSheet).name)
	$oObj.Sheets.moveByName($vSheet, $iPosition)
	Return SetError($_OOoCalcStatus_Success, 0, 1)
EndFunc   ;==>_OOoCalc_SheetMove

; #FUNCTION# ====================================================================================================================
; Name ..........: _OOoCalc_SheetPrint
; Description ...: Prints the specified sheet from the specified Calc object.
; Syntax ........: _OOoCalc_SheetPrint(ByRef $oObj, $vSheet[, $sPrinter = ''[, $iCopyCount = 1[, $sFileName = ''[, $bCollate = True[, $vPages = 'ALL'[, $bWait = False[, $iDuplexMode = $OFF]]]]]]])
; Parameters ....: $oObj        - Calc object opened by a preceding call to _OOoCalc_BookOpen, _OOoCalc_BookNew, or
;                                 _OOoCalc_BookAttach.
;                  $vSheet      - [optional] Worksheet, either by index (0-based) or name.
;                                            Default is -1, which would use the active worksheet.
;                  $sPrinter    - [optional] Printer name. If left blank or printer name is not found, default printer is used.
;                  $iCopyCount  - [optional] Specifies the number of copies to print.
;                                            Default is 1.
;                  $sFileName   - [optional] Specifies the name of a file to which to print.
;                                            Default is a blank string.
;                  $bCollate    - [optional] Advises the printer to collate the pages of the copies.
;                                            Default is True.
;                  $vPages      - [optional] Specifies which pages to print. This range is given as at the user interface.
;                                            For example: '1-4;10' to print the pages 1 to 4 and 10.
;                                            Default is 'ALL'.
;                  $bWait       - [optional] If set to True, the corresponding print request will be executed synchronous.
;                                            Default is the asynchronous print mode. ATTENTION: Setting this field to True is
;                                            highly recommended. Otherwhise following actions (as e.g. closing the corresponding
;                                            model) can fail.
;                  $iDuplexMode - [optional] Determines the duplex mode for the print job.
;                                 0 ($UNKNOWN)
;                                 1 ($OFF) [Default]
;                                 2 ($LONGEDGE)
;                                 3 ($SHORTEDGE)
; Return values .: On Success   - Returns 1
;                  On Failure   - Returns 0 and sets @error:
;                  |@error      - 0 ($_OOoCalcStatus_Success)         = No Error
;                  |            - 3 ($_OOoCalcStatus_InvalidDataType) = Invalid Data Type
;                  |            - 4 ($_OOoCalcStatus_InvalidValue)    = Invalid Value
;                  |            - 5 ($_OOoCalcStatus_NoMatch)         = No Match
;                  |@extended   - Contains Invalid Parameter Number
; Author ........: GMK
; Modified ......:
; Remarks .......: None
; Related .......: None
; Link ..........: http://www.openoffice.org/api/docs/common/ref/com/sun/star/view/XPrintable.html#print
; Example .......: No
; ===============================================================================================================================
Func _OOoCalc_SheetPrint(ByRef $oObj, $vSheet = -1, $sPrinter = '', $iCopyCount = 1, $sFileName = '', $bCollate = True, $vPages = 'ALL', $bWait = True, $iDuplexMode = $OFF)
	Local $oOOoCalc_COM_ErrorHandler = ObjEvent("AutoIt.Error", __OOoCalc_ComErrorHandler_InternalFunction)
    #forceref $oOOoCalc_COM_ErrorHandler
	If Not IsObj($oObj) Then Return SetError($_OOoCalcStatus_InvalidDataType, 1, 0)
	If Not IsInt($vSheet) And Not IsString($vSheet) Then Return SetError($_OOoCalcStatus_InvalidDataType, 2, 0)
	If $vSheet > -1 And Not __OOoCalc_WorksheetIsValid($oObj, $vSheet) Then Return SetError($_OOoCalcStatus_NoMatch, 2, 0)
	If Not IsString($sPrinter) Then Return SetError($_OOoCalcStatus_InvalidDataType, 3, 0)
	If Not IsInt($iCopyCount) Then Return SetError($_OOoCalcStatus_InvalidDataType, 4, 0)
	If Not IsString($sFileName) Then Return SetError($_OOoCalcStatus_InvalidDataType, 5, 0)
	If Not IsBool($bCollate) Then Return SetError($_OOoCalcStatus_InvalidDataType, 6, 0)
	If Not IsInt($vPages) And Not IsString($vPages) Then Return SetError($_OOoCalcStatus_InvalidDataType, 7, 0)
	If Not IsBool($bWait) Then Return SetError($_OOoCalcStatus_InvalidDataType, 8, 0)
	If $iDuplexMode < 0 Or $iDuplexMode > 3 Then Return SetError($_OOoCalcStatus_InvalidValue, 9, 0)
	Local $sOldSheet = _OOoCalc_SheetNameGet($oObj)
	_OOoCalc_SheetActivate($oObj, $vSheet)
	If $sPrinter <> '' Then
		Local $avSetPrinterOpt[1] = [__OOoCalc_SetPropertyValue('Name', $sPrinter)]
		$oObj.setPrinter($avSetPrinterOpt)
	EndIf
	Local $avPrintOpt[4] = [__OOoCalc_SetPropertyValue('CopyCount', $iCopyCount), _
			__OOoCalc_SetPropertyValue('Collate', $bCollate), _
			__OOoCalc_SetPropertyValue('Wait', $bWait), _
			__OOoCalc_SetPropertyValue('DuplexMode', $iDuplexMode)]
	If $vPages <> 'ALL' Then
		ReDim $avPrintOpt[UBound($avPrintOpt) + 1]
		$avPrintOpt[UBound($avPrintOpt) - 1] = __OOoCalc_SetPropertyValue('Pages', $vPages)
	EndIf
	If $sFileName <> '' Then
		ReDim $avPrintOpt[UBound($avPrintOpt) + 1]
		$avPrintOpt[UBound($avPrintOpt) - 1] = __OOoCalc_SetPropertyValue('FileName', $sFileName)
	EndIf
	$oObj.print($avPrintOpt)
	_OOoCalc_SheetActivate($oObj, $sOldSheet)
	Return SetError($_OOoCalcStatus_Success, 0, 1)
EndFunc   ;==>_OOoCalc_SheetPrint

; #FUNCTION# ====================================================================================================================
; Name ..........: _OOoCalc_HorizontalAlignSet
; Description ...: Set the horizontal alignment of each cell in a range.
; Syntax ........: _OOoCalc_HorizontalAlignSet(ByRef $oObj, $iHorizAlign, $vRangeOrRowStart[, $iColStart = -1[, $iRowEnd = -1[, $iColEnd = -1[, $vSheet = -1]]]])
; Parameters ....: $oObj             - Calc object opened by a preceding call to _OOoCalc_BookOpen, _OOoCalc_BookNew, or
;                                      _OOoCalc_BookAttach
;                  $iHorizAlign      - [optional] Horizontal alignment.
;                                                 Default is 1 ($LEFT).
;                                      0 ($STANDARD)
;                                      1 ($LEFT)
;                                      2 ($CENTER)
;                                      3 ($RIGHT)
;                                      4 ($BLOCK)
;                                      5 ($REPEAT)
;                  $vRangeOrRowStart - Either an A1 range, or index (0-based) of row to set alignment if using RC
;                  $iColStart        - [optional] The index (0-based) of the starting column to set alignment if using RC.
;                                                 Default is -1.
;                  $iRowEnd          - [optional] Index (0-based) of ending row to set alignment if using RC.
;                                                 Default is -1.
;                  $iColEnd          - [optional] Index (0-based) of ending column to set alignment if using RC.
;                                                 Default is -1.
;                  $vSheet           - [optional] Worksheet, either by index (0-based) or name.
;                                                 Default is -1, which would use the active worksheet.
; Return values .: On Success        - Returns 1
;                  On Failure        - Returns 0 and sets @error:
;                  |@error           - 0 ($_OOoCalcStatus_Success)         = No Error
;                  |                 - 3 ($_OOoCalcStatus_InvalidDataType) = Invalid Data Type
;                  |                 - 4 ($_OOoCalcStatus_InvalidValue)    = Invalid Value
;                  |                 - 5 ($_OOoCalcStatus_NoMatch)         = No Match
;                  |@extended        - Contains Invalid Parameter Number
; Author ........: GMK
; Modified ......:
; Remarks .......: None
; Related .......: None
; Link ..........: http://www.openoffice.org/api/docs/common/ref/com/sun/star/table/CellHoriJustify.html
; Example .......: No
; ===============================================================================================================================
Func _OOoCalc_HorizontalAlignSet(ByRef $oObj, $iHorizAlign, $vRangeOrRowStart, $iColStart = -1, $iRowEnd = -1, $iColEnd = -1, $vSheet = -1)
	Local $oOOoCalc_COM_ErrorHandler = ObjEvent("AutoIt.Error", __OOoCalc_ComErrorHandler_InternalFunction)
    #forceref $oOOoCalc_COM_ErrorHandler
	If Not IsObj($oObj) Then Return SetError($_OOoCalcStatus_InvalidDataType, 1, 0)
	If Not IsInt($iHorizAlign) Then Return SetError($_OOoCalcStatus_InvalidDataType, 2, 0)
	If $iHorizAlign < 0 Or $iHorizAlign > 5 Then Return SetError($_OOoCalcStatus_InvalidValue, 2, 0)
	If Not IsInt($vRangeOrRowStart) And Not __OOoCalc_RangeIsValid($vRangeOrRowStart) Then Return SetError($_OOoCalcStatus_InvalidValue, 3, 0)
	If Not IsInt($iColStart) Then Return SetError($_OOoCalcStatus_InvalidDataType, 4, 0)
	If IsInt($vRangeOrRowStart) And $iColStart < 0 Then Return SetError($_OOoCalcStatus_InvalidValue, 4, 0)
	If Not IsInt($iRowEnd) Then Return SetError($_OOoCalcStatus_InvalidDataType, 5, 0)
	If Not IsInt($iColEnd) Then Return SetError($_OOoCalcStatus_InvalidDataType, 6, 0)
	If Not IsInt($vSheet) And Not IsString($vSheet) Then Return SetError($_OOoCalcStatus_InvalidDataType, 7, 0)
	If $vSheet > -1 And Not __OOoCalc_WorksheetIsValid($oObj, $vSheet) Then Return SetError($_OOoCalcStatus_NoMatch, 7, 0)
	Local $oSheet = __OOoCalc_GetSheet($oObj, $vSheet)
	If Not IsObj($oSheet) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	Local $oRange = __OOoCalc_GetRange($oSheet, $vRangeOrRowStart, $iColStart, $iRowEnd, $iColEnd)
	If Not IsObj($oRange) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	$oRange.HoriJustify = $iHorizAlign
	Return SetError($_OOoCalcStatus_Success, 0, 1)
EndFunc   ;==>_OOoCalc_HorizontalAlignSet

; #FUNCTION# ====================================================================================================================
; Name ..........: _OOoCalc_FontSetProperties
; Description ...: Set the bold, italic, and underline font properties of a range in a Calc object.
; Syntax ........: _OOoCalc_FontSetProperties(ByRef $oObj, $vRangeOrRowStart[, $iColStart = -1[, $iRowEnd = -1[, $iColEnd = -1[, $vSheet = -1[, $bBold = False[, $bItalic = False[, $bUnderline = False]]]]]]])
; Parameters ....: $oObj             - Calc object opened by a preceding call to _OOoCalc_BookOpen, _OOoCalc_BookNew, or
;                                      _OOoCalc_BookAttach
;                  $vRangeOrRowStart - Either an A1 range, or index (0-based) of row to set font properties if using RC
;                  $iColStart        - [optional] The column to set font properties if using RC.
;                                                 Default is -1.
;                  $iRowEnd          - [optional] Index (0-based) of ending row to set font properties if using RC.
;                                                 Default is -1.
;                  $iColEnd          - [optional] Index (0-based) of ending column to set font properties if using RC.
;                                                 Default is -1.
;                  $vSheet           - [optional] Worksheet, either by index (0-based) or name.
;                                                 Default is -1, which would use the active worksheet.
;                  $bBold            - [optional] Turn Bold on.
;                                                 Default is False.
;                  $bItalic          - [optional] Turn Italic on.
;                                                 Default is False.
;                  $bUnderline       - [optional] Turn Italic on.
;                                                 Default is False.
; Return values .: On Success        - Returns 1
;                  On Failure        - Returns 0 and sets @error:
;                  |@error           - 0 ($_OOoCalcStatus_Success)         = No Error
;                  |                 - 3 ($_OOoCalcStatus_InvalidDataType) = Invalid Data Type
;                  |                 - 4 ($_OOoCalcStatus_InvalidValue)    = Invalid Value
;                  |                 - 5 ($_OOoCalcStatus_NoMatch)         = No Match
;                  |@extended        - Contains Invalid Parameter Number
; Author ........: GMK
; Modified ......:
; Remarks .......: None
; Related .......: None
; Link ..........: http://www.openoffice.org/api/docs/common/ref/com/sun/star/style/CharacterProperties.html
; Example .......: No
; ===============================================================================================================================
Func _OOoCalc_FontSetProperties(ByRef $oObj, $vRangeOrRowStart, $iColStart = -1, $iRowEnd = -1, $iColEnd = -1, $vSheet = -1, $bBold = False, $bItalic = False, $bUnderline = False)
	Local $oOOoCalc_COM_ErrorHandler = ObjEvent("AutoIt.Error", __OOoCalc_ComErrorHandler_InternalFunction)
    #forceref $oOOoCalc_COM_ErrorHandler
	If Not IsObj($oObj) Then Return SetError($_OOoCalcStatus_InvalidDataType, 1, 0)
	If Not IsInt($vRangeOrRowStart) And Not __OOoCalc_RangeIsValid($vRangeOrRowStart) Then Return SetError($_OOoCalcStatus_InvalidValue, 2, 0)
	If Not IsInt($iColStart) Then Return SetError($_OOoCalcStatus_InvalidDataType, 3, 0)
	If IsInt($vRangeOrRowStart) And $iColStart < 0 Then Return SetError($_OOoCalcStatus_InvalidValue, 3, 0)
	If Not IsInt($iRowEnd) Then Return SetError($_OOoCalcStatus_InvalidDataType, 4, 0)
	If Not IsInt($iColEnd) Then Return SetError($_OOoCalcStatus_InvalidDataType, 5, 0)
	If Not IsInt($vSheet) And Not IsString($vSheet) Then Return SetError($_OOoCalcStatus_InvalidDataType, 6, 0)
	If $vSheet > -1 And Not __OOoCalc_WorksheetIsValid($oObj, $vSheet) Then Return SetError($_OOoCalcStatus_NoMatch, 6, 0)
	If Not IsBool($bBold) Then Return SetError($_OOoCalcStatus_InvalidDataType, 7, 0)
	If Not IsBool($bItalic) Then Return SetError($_OOoCalcStatus_InvalidDataType, 8, 0)
	If Not IsBool($bUnderline) Then Return SetError($_OOoCalcStatus_InvalidDataType, 9, 0)
	Local $oSheet = __OOoCalc_GetSheet($oObj, $vSheet)
	If Not IsObj($oSheet) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	Local $oRange = __OOoCalc_GetRange($oSheet, $vRangeOrRowStart, $iColStart, $iRowEnd, $iColEnd)
	If Not IsObj($oRange) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	If $bBold Then $oRange.CharWeight = 150
	If $bItalic Then $oRange.CharPosture = 2
	If $bUnderline Then $oRange.CharUnderline = 1
	Return SetError($_OOoCalcStatus_Success, 0, 1)
EndFunc   ;==>_OOoCalc_FontSetProperties

; #FUNCTION# ====================================================================================================================
; Name ..........: _OOoCalc_CellSetColors
; Description ...: Set the foreground and background cell colors of a range in a Calc object.
; Syntax ........: _OOoCalc_CellSetColors(ByRef $oObj, $nForeColor, $nBackColor, $vRangeOrRowStart[, $iColStart = -1[, $iRowEnd = -1[, $iColEnd = -1[, $vSheet = -1]]]])
; Parameters ....: $oObj             - Calc object opened by a preceding call to _OOoCalc_BookOpen, _OOoCalc_BookNew, or
;                                      _OOoCalc_BookAttach
;                  $nForeColor       - Cell foreground (font) color
;                  $nBackColor       - Cell (background) color
;                  $vRangeOrRowStart - Either an A1 range, or index (0-based) of row to set cell colors if using RC
;                  $iColStart        - [optional] The column to set Cell Color if using RC.
;                                                 Default is -1.
;                  $iRowEnd          - [optional] Index (0-based) of ending row to set Cell Color if using RC.
;                                                 Default is -1.
;                  $iColEnd          - [optional] Index (0-based) of ending column to set Cell Color if using RC.
;                                                 Default is -1.
;                  $vSheet           - [optional] Worksheet, either by index (0-based) or name.
;                                                 Default is -1, which would use the active worksheet.
; Return values .: On Success        - Returns 1
;                  On Failure        - Returns 0 and sets @error:
;                  |@error           - 0 ($_OOoCalcStatus_Success)         = No Error
;                  |                 - 3 ($_OOoCalcStatus_InvalidDataType) = Invalid Data Type
;                  |                 - 4 ($_OOoCalcStatus_InvalidValue)    = Invalid Value
;                  |                 - 5 ($_OOoCalcStatus_NoMatch)         = No Match
;                  |@extended        - Contains Invalid Parameter Number
; Author ........: Andy G, GMK
; Modified ......:
; Remarks .......: None
; Related .......: _OOoCalc_RowSetColors, _OOoCalc_ColSetColors
; Link ..........: http://www.openoffice.org/api/docs/common/ref/com/sun/star/table/CellProperties.html#CellBackColor
; Example .......: No
; ===============================================================================================================================
Func _OOoCalc_CellSetColors(ByRef $oObj, $nForeColor, $nBackColor, $vRangeOrRowStart, $iColStart = -1, $iRowEnd = -1, $iColEnd = -1, $vSheet = -1)
	Local $oOOoCalc_COM_ErrorHandler = ObjEvent("AutoIt.Error", __OOoCalc_ComErrorHandler_InternalFunction)
    #forceref $oOOoCalc_COM_ErrorHandler
	If Not IsObj($oObj) Then Return SetError($_OOoCalcStatus_InvalidDataType, 1, 0)
	If Not IsNumber($nForeColor) Then Return SetError($_OOoCalcStatus_InvalidDataType, 2, 0)
	If Not IsNumber($nBackColor) Then Return SetError($_OOoCalcStatus_InvalidDataType, 3, 0)
	If Not IsInt($vRangeOrRowStart) And Not __OOoCalc_RangeIsValid($vRangeOrRowStart) Then Return SetError($_OOoCalcStatus_InvalidValue, 4, 0)
	If Not IsInt($iColStart) Then Return SetError($_OOoCalcStatus_InvalidDataType, 5, 0)
	If IsInt($vRangeOrRowStart) And $iColStart < 0 Then Return SetError($_OOoCalcStatus_InvalidValue, 5, 0)
	If Not IsInt($iRowEnd) Then Return SetError($_OOoCalcStatus_InvalidDataType, 6, 0)
	If Not IsInt($iColEnd) Then Return SetError($_OOoCalcStatus_InvalidDataType, 7, 0)
	If Not IsInt($vSheet) And Not IsString($vSheet) Then Return SetError($_OOoCalcStatus_InvalidDataType, 8, 0)
	If $vSheet > -1 And Not __OOoCalc_WorksheetIsValid($oObj, $vSheet) Then Return SetError($_OOoCalcStatus_NoMatch, 8, 0)
	Local $oSheet = __OOoCalc_GetSheet($oObj, $vSheet)
	If Not IsObj($oSheet) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	Local $oRange = __OOoCalc_GetRange($oSheet, $vRangeOrRowStart, $iColStart, $iRowEnd, $iColEnd)
	If Not IsObj($oRange) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	$oRange.CharColor = $nForeColor
	$oRange.CellBackColor = $nBackColor
	Return SetError($_OOoCalcStatus_Success, 0, 1)
EndFunc   ;==>_OOoCalc_CellSetColors

; #FUNCTION# ====================================================================================================================
; Name ..........: _OOoCalc_RowSetColors
; Description ...: Set the foreground and background cell colors of a range in a Calc object.
; Syntax ........: _OOoCalc_RowSetColors(ByRef $oObj, $nForeColor, $nBackColor, $iRow[, $vSheet = -1])
; Parameters ....: $oObj       - Calc object opened by a preceding call to _OOoCalc_BookOpen, _OOoCalc_BookNew, or
;                                _OOoCalc_BookAttach
;                  $nForeColor - Row foreground (font) color
;                  $nBackColor - Row (background) color
;                  $iRow       - Row in which to set color
;                  $vSheet     - [optional] Worksheet, either by index (0-based) or name.
;                                           Default is -1, which would use the active worksheet.
; Return values .: On Success  - Returns 1
;                  On Failure  - Returns 0 and sets @error:
;                  |@error     - 0 ($_OOoCalcStatus_Success)         = No Error
;                  |           - 3 ($_OOoCalcStatus_InvalidDataType) = Invalid Data Type
;                  |           - 4 ($_OOoCalcStatus_InvalidValue)    = Invalid Value
;                  |           - 5 ($_OOoCalcStatus_NoMatch)         = No Match
;                  |@extended  - Contains Invalid Parameter Number
; Author ........: Andy G, GMK
; Modified ......:
; Remarks .......: None
; Related .......: _OOoCalc_CellSetColors, _OOoCalc_ColumnSetColors
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _OOoCalc_RowSetColors(ByRef $oObj, $nForeColor, $nBackColor, $iRow, $vSheet = -1) ;RRGGBB
	Local $oOOoCalc_COM_ErrorHandler = ObjEvent("AutoIt.Error", __OOoCalc_ComErrorHandler_InternalFunction)
    #forceref $oOOoCalc_COM_ErrorHandler
	If Not IsObj($oObj) Then Return SetError($_OOoCalcStatus_InvalidDataType, 1, 0)
	If Not IsNumber($nForeColor) Then Return SetError($_OOoCalcStatus_InvalidDataType, 2, 0)
	If Not IsNumber($nBackColor) Then Return SetError($_OOoCalcStatus_InvalidDataType, 3, 0)
	If Not IsInt($iRow) Then Return SetError($_OOoCalcStatus_InvalidValue, 4, 0)
	If Not IsInt($vSheet) And Not IsString($vSheet) Then Return SetError($_OOoCalcStatus_InvalidDataType, 5, 0)
	If $vSheet > -1 And Not __OOoCalc_WorksheetIsValid($oObj, $vSheet) Then Return SetError($_OOoCalcStatus_NoMatch, 5, 0)
	Local $oSheet = __OOoCalc_GetSheet($oObj, $vSheet)
	If Not IsObj($oSheet) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	Local $oRowComponent = $oSheet.getRows.getByIndex($iRow)
	If Not IsObj($oRowComponent) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	$oRowComponent.CharColor = $nForeColor
	$oRowComponent.CellBackColor = $nBackColor
	Return SetError($_OOoCalcStatus_Success, 0, 1)
EndFunc   ;==>_OOoCalc_RowSetColors

; #FUNCTION# ====================================================================================================================
; Name ..........: _OOoCalc_ColumnSetColors
; Description ...: Set the foreground and background cell colors of a range in a Calc object.
; Syntax ........: _OOoCalc_ColumnSetColors(ByRef $oObj, $nForeColor, $nBackColor, $iCol[, $vSheet = -1])
; Parameters ....: $oObj       - Calc object opened by a preceding call to _OOoCalc_BookOpen, _OOoCalc_BookNew, or
;                                _OOoCalc_BookAttach
;                  $nForeColor - Column foreground (font) color
;                  $nBackColor - Column (background) color
;                  $iCol       - Column in which to set color
;                  $vSheet     - [optional] Worksheet, either by index (0-based) or name.
;                                           Default is -1, which would use the active worksheet.
; Return values .: On Success  - Returns 1
;                  On Failure  - Returns 0 and sets @error:
;                  |@error     - 0 ($_OOoCalcStatus_Success)         = No Error
;                  |           - 3 ($_OOoCalcStatus_InvalidDataType) = Invalid Data Type
;                  |           - 4 ($_OOoCalcStatus_InvalidValue)    = Invalid Value
;                  |           - 5 ($_OOoCalcStatus_NoMatch)         = No Match
;                  |@extended  - Contains Invalid Parameter Number
; Author ........: Andy G, GMK
; Modified ......:
; Remarks .......: None
; Related .......: _OOoCalc_CellSetColors, _OOoCalc_ColumnSetColors
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _OOoCalc_ColumnSetColors(ByRef $oObj, $nForeColor, $nBackColor, $iCol, $vSheet = -1) ;RRGGBB
	Local $oOOoCalc_COM_ErrorHandler = ObjEvent("AutoIt.Error", __OOoCalc_ComErrorHandler_InternalFunction)
    #forceref $oOOoCalc_COM_ErrorHandler
	If Not IsObj($oObj) Then Return SetError($_OOoCalcStatus_InvalidDataType, 1, 0)
	If Not IsNumber($nForeColor) Then Return SetError($_OOoCalcStatus_InvalidDataType, 2, 0)
	If Not IsNumber($nBackColor) Then Return SetError($_OOoCalcStatus_InvalidDataType, 3, 0)
	If Not IsInt($iCol) Then Return SetError($_OOoCalcStatus_InvalidValue, 4, 0)
	If Not IsInt($vSheet) And Not IsString($vSheet) Then Return SetError($_OOoCalcStatus_InvalidDataType, 5, 0)
	If $vSheet > -1 And Not __OOoCalc_WorksheetIsValid($oObj, $vSheet) Then Return SetError($_OOoCalcStatus_NoMatch, 5, 0)
	Local $oSheet = __OOoCalc_GetSheet($oObj, $vSheet)
	If Not IsObj($oSheet) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	Local $oColumnComponent = $oSheet.getColumns.getByIndex($iCol)
	If Not IsObj($oColumnComponent) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	$oColumnComponent.CharColor = $nForeColor
	$oColumnComponent.CellBackColor = $nBackColor
	Return SetError($_OOoCalcStatus_Success, 0, 1)
EndFunc   ;==>_OOoCalc_ColumnSetColors

; #FUNCTION# ====================================================================================================================
; Name ..........: _OOoCalc_RowSetProperties
; Description ...: Set row properties, including optimal height (autosize), visibility (hidden or shown) and new page (page break).
; Syntax ........: _OOoCalc_RowSetProperties(ByRef $oObj, $iRow, $nHeight[, $bOptHeight = True[, $bVisible = True[, $bNewPage = False[, $vSheet = -1]]]])
; Parameters ....: $oObj       - Calc object opened by a preceding call to _OOoCalc_BookOpen, _OOoCalc_BookNew, or
;                                _OOoCalc_BookAttach
;                  $iRow       - Row in which to set properties
;                  $nHeight    - Row height, in 100ths mm
;                  $bOptHeight - [optional] Set optimal height (autosize).
;                                           Default is True.
;                  $bVisible   - [optional] Set row to visible (unhidden).
;                                           Default is True.
;                  $bNewPage   - [optional] Start new page on row (page break).
;                                           Default is False.
;                  $vSheet     - [optional] Worksheet, either by index (0-based) or name.
;                                           Default is -1, which would use the active worksheet.
; Return values .: On Success  - Returns 1
;                  On Failure  - Returns 0 and sets @error:
;                  |@error     - 0 ($_OOoCalcStatus_Success)         = No Error
;                  |           - 3 ($_OOoCalcStatus_InvalidDataType) = Invalid Data Type
;                  |           - 4 ($_OOoCalcStatus_InvalidValue)    = Invalid Value
;                  |           - 5 ($_OOoCalcStatus_NoMatch)         = No Match
;                  |@extended  - Contains Invalid Parameter Number
; Author ........: Andy G, GMK
; Modified ......:
; Remarks .......: None
; Related .......: _OOoCalc_ColumnSetProperties
; Link ..........: http://www.openoffice.org/api/docs/common/ref/com/sun/star/table/TableRow.html
; Example .......: No
; ===============================================================================================================================
Func _OOoCalc_RowSetProperties(ByRef $oObj, $iRow, $nHeight, $bOptHeight = True, $bVisible = True, $bNewPage = False, $vSheet = -1)
	Local $oOOoCalc_COM_ErrorHandler = ObjEvent("AutoIt.Error", __OOoCalc_ComErrorHandler_InternalFunction)
    #forceref $oOOoCalc_COM_ErrorHandler
	If Not IsObj($oObj) Then Return SetError($_OOoCalcStatus_InvalidDataType, 1, 0)
	If Not IsInt($iRow) Then Return SetError($_OOoCalcStatus_InvalidDataType, 2, 0)
	If Not IsNumber($nHeight) Then Return SetError($_OOoCalcStatus_InvalidDataType, 3, 0)
	If Not IsBool($bOptHeight) Then Return SetError($_OOoCalcStatus_InvalidDataType, 4, 0)
	If Not IsBool($bVisible) Then Return SetError($_OOoCalcStatus_InvalidDataType, 5, 0)
	If Not IsBool($bNewPage) Then Return SetError($_OOoCalcStatus_InvalidDataType, 6, 0)
	If Not IsInt($vSheet) And Not IsString($vSheet) Then Return SetError($_OOoCalcStatus_InvalidDataType, 7, 0)
	If $vSheet > -1 And Not __OOoCalc_WorksheetIsValid($oObj, $vSheet) Then Return SetError($_OOoCalcStatus_NoMatch, 7, 0)
	Local $oSheet = __OOoCalc_GetSheet($oObj, $vSheet)
	If Not IsObj($oSheet) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	Local $oRowComponent = $oSheet.getRows.getByIndex($iRow)
	If Not IsObj($oRowComponent) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	If Not $bOptHeight Then $oRowComponent.Height = $nHeight ;column height (in 100ths of mm)
	With $oRowComponent
		.OptimalHeight = Number($bOptHeight)
		.IsVisible = Number($bVisible)
		.IsStartOfNewPage = Number($bNewPage)
	EndWith
	Return SetError($_OOoCalcStatus_Success, 0, 1)
EndFunc   ;==>_OOoCalc_RowSetProperties

; #FUNCTION# ====================================================================================================================
; Name ..........: _OOoCalc_ColumnSetProperties
; Description ...: Set Col properties, including optimal Width (autosize), visibility (hidden or shown) and new page (page break).
; Syntax ........: _OOoCalc_ColumnSetProperties(ByRef $oObj, $iCol, $nWidth[, $bOptWidth = True[, $bVisible = True[, $bNewPage = False[, $vSheet = -1]]]])
; Parameters ....: $oObj      - Calc object opened by a preceding call to _OOoCalc_BookOpen, _OOoCalc_BookNew, or
;                               _OOoCalc_BookAttach
;                  $iCol      - Col in which to set properties
;                  $nWidth    - Col width, in 1/100 mm
;                  $bOptWidth - [optional] Set optimal Width (autosize).
;                                          Default is True
;                  $bVisible  - [optional] Set Col to visible (unhidden).
;                                          Default is True.
;                  $bNewPage  - [optional] Start new page on Col (page break).
;                                          Default is False.
;                  $vSheet    - [optional] Worksheet, either by index (0-based) or name.
;                                          Default is -1, which would use the active worksheet.
; Return values .: On Success - Returns 1
;                  On Failure - Returns 0 and sets @error:
;                  |@error    - 0 ($_OOoCalcStatus_Success)         = No Error
;                  |          - 3 ($_OOoCalcStatus_InvalidDataType) = Invalid Data Type
;                  |          - 4 ($_OOoCalcStatus_InvalidValue)    = Invalid Value
;                  |          - 5 ($_OOoCalcStatus_NoMatch)         = No Match
;                  |@extended - Contains Invalid Parameter Number
; Author ........: Andy G, GMK
; Modified ......:
; Remarks .......: None
; Related .......: _OOoCalc_RowSetProperties
; Link ..........: http://www.openoffice.org/api/docs/common/ref/com/sun/star/table/TableCol.html
; Example .......: No
; ===============================================================================================================================
Func _OOoCalc_ColumnSetProperties(ByRef $oObj, $iCol, $nWidth, $bOptWidth = True, $bVisible = True, $bNewPage = False, $vSheet = -1)
	Local $oOOoCalc_COM_ErrorHandler = ObjEvent("AutoIt.Error", __OOoCalc_ComErrorHandler_InternalFunction)
    #forceref $oOOoCalc_COM_ErrorHandler
	If Not IsObj($oObj) Then Return SetError($_OOoCalcStatus_InvalidDataType, 1, 0)
	If Not IsInt($iCol) Then Return SetError($_OOoCalcStatus_InvalidDataType, 2, 0)
	If Not IsNumber($nWidth) Then Return SetError($_OOoCalcStatus_InvalidDataType, 3, 0)
	If Not IsBool($bOptWidth) Then Return SetError($_OOoCalcStatus_InvalidDataType, 4, 0)
	If Not IsBool($bVisible) Then Return SetError($_OOoCalcStatus_InvalidDataType, 5, 0)
	If Not IsBool($bNewPage) Then Return SetError($_OOoCalcStatus_InvalidDataType, 6, 0)
	If Not IsInt($vSheet) And Not IsString($vSheet) Then Return SetError($_OOoCalcStatus_InvalidDataType, 7, 0)
	If $vSheet > -1 And Not __OOoCalc_WorksheetIsValid($oObj, $vSheet) Then Return SetError($_OOoCalcStatus_NoMatch, 7, 0)
	Local $oSheet = __OOoCalc_GetSheet($oObj, $vSheet)
	If Not IsObj($oSheet) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	Local $oColComponent = $oSheet.getColumns.getByIndex($iCol)
	If Not IsObj($oColComponent) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	If IsNumber($nWidth) And Not $bOptWidth Then $oColComponent.Width = $nWidth ;column Width (in 100ths of mm)
	With $oColComponent
		.OptimalWidth = Number($bOptWidth)
		.IsVisible = Number($bVisible)
		.IsStartOfNewPage = Number($bNewPage)
	EndWith
	Return SetError($_OOoCalcStatus_Success, 0, 1)
EndFunc   ;==>_OOoCalc_ColumnSetProperties

; #FUNCTION# ====================================================================================================================
; Name ..........: _OOoCalc_FindInRange
; Description ...: Finds all instances of a string in a range and returns their addresses as a two-dimensional array.
; Syntax ........: _OOoCalc_FindInRange(ByRef $oObj, $sString[, $vRangeOrRowStart = -1[, $iColStart = -1[, $iRowEnd = -1[, $iColEnd = -1[, $vSheet = -1[, $bWholeWord = False[, $bMatchCase = False[, $bRegExp = False]]]]]]]])
; Parameters ....: $oObj             - Calc object opened by a preceding call to _OOoCalc_BookOpen, _OOoCalc_BookNew, or
;                                      _OOoCalc_BookAttach
;                  $sString          - The string for which to search
;                  $vRangeOrRowStart - [optional] Either an A1 range, or index (0-based) of row to search if using RC. If
;                                      default value of -1 is used, the used range will be used.
;                  $iColStart        - [optional] The column to search if using RC.
;                                                 Default is -1.
;                  $iRowEnd          - [optional] Index (0-based) of ending row to search if using RC.
;                                                 Default is -1.
;                  $iColEnd          - [optional] Index (0-based) of ending column to search if using RC.
;                                                 Default is -1.
;                  $vSheet           - [optional] Worksheet, either by index (0-based) or name.
;                                                 Default is -1, which would use the activevworksheet.
;                  $bWholeWord       - [optional] If True, only complete words will be found.  If False, partial match is possible.
;                                                 Default is False.
;                  $bMatchCase       - [optional] Specify whether case should match in search (True or False).
;                                                 Default is False.
;                  $bRegExp          - [optional] If True, the search string is evaluated as a regular expression.
; Return values .: On Success        - Returns a two dimensional array with addresses of matching cells. If no matches found, returns null string.
;                  |                   UBound($avArray) - The number of found cells
;                  |                   $avArray[x][0] - The address of found cell
;                  |                   $avArray[x][1] - The index of the row of found cell (0-based)
;                  |                   $avArray[x][2] - The index of the column of found cell (0-based)
;                  |                   $avArray[x][3] - The value of found cell
;                  On Failure        - Returns 0 and sets @error:
;                  |@error           - 0 ($_OOoCalcStatus_Success)         = No Error
;                  |                 - 1 ($_OOoCalcStatus_GeneralError)    = General Error
;                  |                 - 3 ($_OOoCalcStatus_InvalidDataType) = Invalid Data Type
;                  |                 - 4 ($_OOoCalcStatus_InvalidValue)    = Invalid Value
;                  |                 - 5 ($_OOoCalcStatus_NoMatch)         = No Match
;                  |@extended        - Contains Invalid Parameter Number
; Author ........: Leagnus, GMK
; Modified ......:
; Remarks .......: None
; Related .......: None
; Link ..........: http://www.openoffice.org/api/docs/common/ref/com/sun/star/util/SearchDescriptor.html
; Example .......: No
; ===============================================================================================================================
Func _OOoCalc_FindInRange(ByRef $oObj, $sString, $vRangeOrRowStart = -1, $iColStart = -1, $iRowEnd = -1, $iColEnd = -1, $vSheet = -1, $bWholeWord = False, $bMatchCase = False, $bRegExp = False)
	Local $oOOoCalc_COM_ErrorHandler = ObjEvent("AutoIt.Error", __OOoCalc_ComErrorHandler_InternalFunction)
    #forceref $oOOoCalc_COM_ErrorHandler
	Local Enum $eAddress, $eRow, $eColumn, $eString
	If Not IsObj($oObj) Then Return SetError($_OOoCalcStatus_InvalidDataType, 1, 0)
	If Not IsString($sString) Then Return SetError($_OOoCalcStatus_InvalidDataType, 2, 0)
	If Not IsInt($vRangeOrRowStart) And Not __OOoCalc_RangeIsValid($vRangeOrRowStart) Then Return SetError($_OOoCalcStatus_InvalidValue, 3, 0)
	If Not IsInt($iColStart) Then Return SetError($_OOoCalcStatus_InvalidDataType, 4, 0)
	If $vRangeOrRowStart > -1 And IsInt($vRangeOrRowStart) And $iColStart < 0 Then Return SetError($_OOoCalcStatus_InvalidValue, 4, 0)
	If Not IsInt($iRowEnd) Then Return SetError($_OOoCalcStatus_InvalidDataType, 5, 0)
	If Not IsInt($iColEnd) Then Return SetError($_OOoCalcStatus_InvalidDataType, 6, 0)
	If Not IsInt($vSheet) And Not IsString($vSheet) Then Return SetError($_OOoCalcStatus_InvalidDataType, 7, 0)
	If $vSheet > -1 And Not __OOoCalc_WorksheetIsValid($oObj, $vSheet) Then Return SetError($_OOoCalcStatus_NoMatch, 7, 0)
	If Not IsBool($bWholeWord) Then Return SetError($_OOoCalcStatus_InvalidDataType, 8, 0)
	If Not IsBool($bMatchCase) Then Return SetError($_OOoCalcStatus_InvalidDataType, 9, 0)
	If Not IsBool($bRegExp) Then Return SetError($_OOoCalcStatus_InvalidDataType, 10, 0)
	Local $oSheet = __OOoCalc_GetSheet($oObj, $vSheet)
	If Not IsObj($oSheet) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	Local $oRange = __OOoCalc_GetRange($oSheet, $vRangeOrRowStart, $iColStart, $iRowEnd, $iColEnd)
	If Not IsObj($oRange) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	Local $iCount = 0 ; row index starts with 0
	; 	Create a descriptor from a searchable document:
	Local $oSearchDescriptor = $oRange.createSearchDescriptor
	If Not IsObj($oSearchDescriptor) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	; 	Set the text for which to search and other
	; 	http://api.openoffice.org/docs/common/ref/com/sun/star/util/SearchDescriptor.html
	With $oSearchDescriptor
		.SearchString = $sString
		;    	SearchWords forces the entire cell to contain only the search string:
		.SearchWords = $bWholeWord
		.SearchCaseSensitive = $bMatchCase
		.SearchRegularExpression = $bRegExp
	EndWith
	Local $avReturn[$iCount + 1][4] ; 1 row 4 columns
	; 	Find the first one:
	Local $oCell = $oSheet.findFirst($oSearchDescriptor)
	If Not IsObj($oCell) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	$avReturn[0][$eAddress] = __OOoCalc_CellRCToA1($oCell.CellAddress.Row, $oCell.CellAddress.Column)
	$avReturn[0][$eRow] = $oCell.CellAddress.Row
	$avReturn[0][$eColumn] = $oCell.CellAddress.Column
	$avReturn[0][$eString] = $oCell.getString
	; 	Find all next instances:
	While IsObj($oCell)
		$oCell = $oSheet.findNext($oCell, $oSearchDescriptor)
		If IsObj($oCell) Then
			$iCount += 1
			ReDim $avReturn[$iCount + 1][3] ; +1 because ReDim starts not with 0 like indices
			$avReturn[0][$eAddress] = __OOoCalc_CellRCToA1($oCell.CellAddress.Row, $oCell.CellAddress.Column)
			$avReturn[0][$eRow] = $oCell.CellAddress.Row
			$avReturn[0][$eColumn] = $oCell.CellAddress.Column
			$avReturn[0][$eString] = $oCell.getString
		EndIf
	WEnd
	Return SetError($_OOoCalcStatus_Success, 0, $avReturn)
EndFunc   ;==>_OOoCalc_FindInRange

; #FUNCTION# ====================================================================================================================
; Name ..........: _OOoCalc_ReplaceInRange
; Description ...: Finds all instances of a string in a range and replace them with the replace string.
; Syntax ........: _OOoCalc_ReplaceInRange(ByRef $oObj, $sSearchString, $sReplaceString[, $vRangeOrRowStart = -1[, $iColStart = -1[, $iRowEnd = -1[, $iColEnd = -1[, $vSheet = -1[, $bWholeWord = False[, $bMatchCase = False[, $bRegExp = False]]]]]]]])
; Parameters ....: $oObj             - Calc object opened by a preceding call to _OOoCalc_BookOpen, _OOoCalc_BookNew, or
;                                      _OOoCalc_BookAttach
;                  $sSearchString    - Search string.
;                  $sReplaceString   - Replacement string.
;                  $vSheet           - [optional] Worksheet, either by index (0-based) or name.
;                                                 Default is -1, which would use the active worksheet.
;                  $vRangeOrRowStart - [optional] Either an A1 range, or index (0-based) of row to search if using RC. If
;                                      default value of -1 is used, the used range will be used.
;                  $iColStart        - [optional] The column to search if using RC.
;                                                 Default is -1.
;                  $iRowEnd          - [optional] Index (0-based) of ending row to search if using RC.
;                                                 Default is -1.
;                  $iColEnd          - [optional] Index (0-based) of ending column to search if using RC.
;                                                 Default is -1.
;                  $bWholeWord       - [optional] Match whole word.
;                                                 Default is False.
;                  $bMatchCase       - [optional] Match case.
;                                                 Default is False.
;                  $bRegExp          - [optional] Use regular expression.
;                                                 Default is False.
; Return values .: On Success        - Returns 1
;                  On Failure        - Returns 0 and sets @error:
;                  |@error           - 0 ($_OOoCalcStatus_Success)         = No Error
;                  |                 - 3 ($_OOoCalcStatus_InvalidDataType) = Invalid Data Type
;                  |                 - 4 ($_OOoCalcStatus_InvalidValue)    = Invalid Value
;                  |                 - 5 ($_OOoCalcStatus_NoMatch)         = No Match
;                  |@extended        - Contains Invalid Parameter Number
; Author ........: GMK
; Modified ......:
; Remarks .......: None
; Related .......: _OOoCalc_FindInRange
; Link ..........: http://www.openoffice.org/api/docs/common/ref/com/sun/star/util/XReplaceable.html
; Example .......: No
; ===============================================================================================================================
Func _OOoCalc_ReplaceInRange(ByRef $oObj, $sSearchString, $sReplaceString, $vRangeOrRowStart = -1, $iColStart = -1, $iRowEnd = -1, $iColEnd = -1, $vSheet = -1, $bWholeWord = False, $bMatchCase = False, $bRegExp = False)
	Local $oOOoCalc_COM_ErrorHandler = ObjEvent("AutoIt.Error", __OOoCalc_ComErrorHandler_InternalFunction)
    #forceref $oOOoCalc_COM_ErrorHandler
	If Not IsObj($oObj) Then Return SetError($_OOoCalcStatus_InvalidDataType, 1, 0)
	If Not IsString($sSearchString) Then Return SetError($_OOoCalcStatus_InvalidDataType, 2, 0)
	If Not IsString($sReplaceString) Then Return SetError($_OOoCalcStatus_InvalidDataType, 3, 0)
	If Not IsInt($vRangeOrRowStart) And Not __OOoCalc_RangeIsValid($vRangeOrRowStart) Then Return SetError($_OOoCalcStatus_InvalidValue, 4, 0)
	If Not IsInt($iColStart) Then Return SetError($_OOoCalcStatus_InvalidDataType, 5, 0)
	If IsInt($vRangeOrRowStart) And $vRangeOrRowStart > -1 And $iColStart < 0 Then Return SetError($_OOoCalcStatus_InvalidValue, 5, 0)
	If Not IsInt($iRowEnd) Then Return SetError($_OOoCalcStatus_InvalidDataType, 6, 0)
	If Not IsInt($iColEnd) Then Return SetError($_OOoCalcStatus_InvalidDataType, 7, 0)
	If Not IsInt($vSheet) And Not IsString($vSheet) Then Return SetError($_OOoCalcStatus_InvalidDataType, 8, 0)
	If $vSheet > -1 And Not __OOoCalc_WorksheetIsValid($oObj, $vSheet) Then Return SetError($_OOoCalcStatus_NoMatch, 8, 0)
	If Not IsBool($bWholeWord) Then Return SetError($_OOoCalcStatus_InvalidDataType, 9, 0)
	If Not IsBool($bMatchCase) Then Return SetError($_OOoCalcStatus_InvalidDataType, 10, 0)
	If Not IsBool($bRegExp) Then Return SetError($_OOoCalcStatus_InvalidDataType, 11, 0)
	Local $oSheet = __OOoCalc_GetSheet($oObj, $vSheet)
	If Not IsObj($oSheet) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	Local $oRange = __OOoCalc_GetRange($oSheet, $vRangeOrRowStart, $iColStart, $iRowEnd, $iColEnd)
	If Not IsObj($oRange) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	Local $oReplaceDescriptor = $oRange.createReplaceDescriptor
	If Not IsObj($oReplaceDescriptor) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	With $oReplaceDescriptor
		.SearchString = $sSearchString
		.ReplaceString = $sReplaceString
		.SearchWords = $bWholeWord
		.SearchCaseSensitive = $bMatchCase
		.SearchRegularExpression = $bRegExp
	EndWith
	$oRange.ReplaceAll($oReplaceDescriptor)
	Return SetError($_OOoCalcStatus_Success, 0, 1)
EndFunc   ;==>_OOoCalc_ReplaceInRange

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __OOoCalc_CellA1ToRC
; Description ...: Extracts 0-based index of row and column from given cell address
; Syntax ........: __OOoCalc_CellA1ToRC($sCellName)
; Parameters ....: $sCellName - A string value.
; Return values .: An array containing the row and column (0-based index) of the given cell address
;                  |[0] = Row (0-based index)
;                  |[1] = Column (0-based index)
; Author ........: Leagnus, GMK
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func __OOoCalc_CellA1ToRC($sCellName)
	Local $aiReturn[2]
	$sCellName = StringUpper($sCellName)
	Local $aiRow = StringRegExp($sCellName, '\d+', 1)
	$aiReturn[0] = $aiRow[0] - 1
	Local $asColumn = StringRegExp($sCellName, '[[:alpha:]]{0,2}', 1)
	Local $iColumn = (Asc(StringMid($asColumn[0], 1, 1)) - 65)
	If StringLen($asColumn[0]) = 2 Then $iColumn = (($iColumn + 1) * 26) + (Asc(StringMid($asColumn[0], 2, 1)) - 65)
	$aiReturn[1] = $iColumn
	Return $aiReturn
EndFunc   ;==>__OOoCalc_CellA1ToRC

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __OOoCalc_CellRCToA1
; Description ...: Convert row and column to cell address
; Syntax ........: __OOoCalc_CellRCToA1($iRow, $iCol)
; Parameters ....: $iRow - Index (0-based) of row
;                  $iCol - Index (0-based) of column
; Return values .: Address of cell
; Author ........: Leagnus
; Modified ......: GMK, mLipok
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func __OOoCalc_CellRCToA1($iRow, $iCol)
    Local $iQuotient = Int($iCol / 26)
    Local $iRemainder = Mod($iCol, 26)
    Local $sChar = Chr(65 + $iRemainder)
    If $iQuotient <> 0 Then $sChar = Chr(65 + $iQuotient) & $sChar
    Return $sChar & String($iRow + 1)
EndFunc   ;==>__OOoCalc_CellRCToA1

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __OOoCalc_CreateBorderLine
; Description ...: Create a border line for use in _OOoCalc_CreateBorders
; Syntax ........: __OOoCalc_CreateBorderLine($nLineWidth[, $bDouble = False])
; Parameters ....: $nLineWidth - Line width in 1/100 mm
;                  $bDouble    - [optional] Double line. Default is False.
; Return values .: Returns border line object
; Author ........: GMK
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func __OOoCalc_CreateBorderLine($nLineWidth, $bDouble = False)
	Local $oOOoCalc_COM_ErrorHandler = ObjEvent("AutoIt.Error", __OOoCalc_ComErrorHandler_InternalFunction)
    #forceref $oOOoCalc_COM_ErrorHandler
	If Not IsNumber($nLineWidth) Then Return SetError($_OOoCalcStatus_InvalidDataType, 1, 0)
	If Not IsBool($bDouble) Then Return SetError($_OOoCalcStatus_InvalidDataType, 2, 0)
	Local $oSM = ObjCreate("com.sun.star.ServiceManager")
	If Not IsObj($oSM) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	Local $oReturn = $oSM.Bridge_GetStruct('com.sun.star.table.BorderLine')
	If Not IsObj($oReturn) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	$oReturn.OuterLineWidth = $nLineWidth
	If $bDouble Then
		$oReturn.LineDistance = $nLineWidth
		$oReturn.InnerLineWidth = $nLineWidth
	Else
		$oReturn.InnerLineWidth = 0
	EndIf
	Return SetError($_OOoCalcStatus_Success, 0, $oReturn)
EndFunc   ;==>__OOoCalc_CreateBorderLine

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __OOoCalc_FileToURL
; Description ...: Convert filename to URL
; Syntax ........: __OOoCalc_FileToURL($sFileName)
; Parameters ....: $sFileName - Full path of file name
; Return values .: URL of the filename
; Author ........: Leagnus, GMK
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func __OOoCalc_FileToURL($sFileName)
	If Not IsString($sFileName) Then Return SetError($_OOoCalcStatus_InvalidDataType, 1, 0)
	$sFileName = StringReplace($sFileName, ':', '|')
	$sFileName = StringReplace($sFileName, ' ', '%20')
	Local $sReturn = 'file:///' & StringReplace($sFileName, '\', '/')
	Return SetError($_OOoCalcStatus_Success, 0, $sReturn)
EndFunc   ;==>__OOoCalc_FileToURL

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __OOoCalc_GetCell
; Description ...: Get cell object from given sheet and address or row and column
; Syntax ........: __OOoCalc_GetCell(ByRef $oSheet, $vRangeOrRow[, $iCol = -1])
; Parameters ....: $oSheet      - Object returned by __OOoCalc_GetSheet
;                  $vRangeOrRow - Either an A1 range, or index (0-based) of row of cell if using RC.
;                  $iCol        - [optional] Index (0-based) of column of cell if using RC. Default is -1.
; Return values .: Returns a cell object
; Author ........: GMK
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func __OOoCalc_GetCell(ByRef $oSheet, $vRangeOrRow, $iCol = -1)
	Local $oOOoCalc_COM_ErrorHandler = ObjEvent("AutoIt.Error", __OOoCalc_ComErrorHandler_InternalFunction)
    #forceref $oOOoCalc_COM_ErrorHandler
	If Not IsObj($oSheet) Then Return SetError($_OOoCalcStatus_InvalidDataType, 1, 0)
	If Not IsString($vRangeOrRow) And Not IsInt($vRangeOrRow) Then Return SetError($_OOoCalcStatus_InvalidDataType, 2, 0)
	If IsString($vRangeOrRow) And Not __OOoCalc_RangeIsValid($vRangeOrRow) Then Return SetError($_OOoCalcStatus_InvalidValue, 2, 0)
	If Not IsInt($iCol) Then Return SetError($_OOoCalcStatus_InvalidDataType, 3, 0)
	Local $oReturn
	Select
		Case IsString($vRangeOrRow)
			$oReturn = $oSheet.getCellRangeByName($vRangeOrRow)
		Case IsInt($vRangeOrRow)
			$oReturn = $oSheet.getCellByPosition($iCol, $vRangeOrRow)
	EndSelect
	If Not IsObj($oReturn) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	Return SetError($_OOoCalcStatus_Success, 0, $oReturn)
EndFunc   ;==>__OOoCalc_GetCell

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __OOoCalc_GetRange
; Description ...: Get range object from given sheet and address range or rows and columns
; Syntax ........: __OOoCalc_GetRange(ByRef $oSheet, $vRangeOrRowStart[, $iColStart = -1[, $iRowEnd = -1[, $iColEnd = -1]]])
; Parameters ....: $oSheet           - Object returned by __OOoCalc_GetSheet
;                  $vRangeOrRowStart - Either an A1 range, or index (0-based) of starting row of range if using RC.
;                                      A value of -1 will return the used range.
;                  $iColStart        - [optional] The index (0-based) of the starting column of range if using RC.
;                                                 Default is -1.
;                  $iRowEnd          - [optional] Index (0-based) of ending row of range if using RC.
;                                                 Default is -1.
;                  $iColEnd          - [optional] Index (0-based) of ending column of range if using RC.
;                                                 Default is -1.
; Return values .: Returns a range object
; Author ........: GMK
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func __OOoCalc_GetRange(ByRef $oSheet, $vRangeOrRowStart, $iColStart = -1, $iRowEnd = -1, $iColEnd = -1)
	Local $oOOoCalc_COM_ErrorHandler = ObjEvent("AutoIt.Error", __OOoCalc_ComErrorHandler_InternalFunction)
    #forceref $oOOoCalc_COM_ErrorHandler
	If Not IsObj($oSheet) Then Return SetError($_OOoCalcStatus_InvalidDataType, 1, 0)
	If Not IsInt($vRangeOrRowStart) And Not IsString($vRangeOrRowStart) Then Return SetError($_OOoCalcStatus_InvalidDataType, 2, 0)
	If IsString($vRangeOrRowStart) And Not __OOoCalc_RangeIsValid($vRangeOrRowStart) Then Return SetError($_OOoCalcStatus_InvalidValue, 2, 0)
	If Not IsInt($iColStart) Then Return SetError($_OOoCalcStatus_InvalidDataType, 3, 0)
	If Not IsInt($iRowEnd) Then Return SetError($_OOoCalcStatus_InvalidDataType, 4, 0)
	If Not IsInt($iColEnd) Then Return SetError($_OOoCalcStatus_InvalidDataType, 5, 0)
	Local $oReturn = __OOoCalc_GetUsedRange($oSheet)
	Select
		Case IsString($vRangeOrRowStart)
			$oReturn = $oSheet.getCellRangeByName($vRangeOrRowStart)
		Case IsInt($vRangeOrRowStart) And $vRangeOrRowStart > -1
			If $iColEnd < 0 Then $iColEnd = $iColStart
			If $iRowEnd < 0 Then $iRowEnd = $vRangeOrRowStart
			$oReturn = $oSheet.getCellRangeByPosition($iColStart, $vRangeOrRowStart, $iColEnd, $iRowEnd)
	EndSelect
	If Not IsObj($oReturn) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	Return SetError($_OOoCalcStatus_Success, 0, $oReturn)
EndFunc   ;==>__OOoCalc_GetRange

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __OOoCalc_GetSheet
; Description ...: Get sheet object from given Calc object and sheet name or index
; Syntax ........: __OOoCalc_GetSheet(ByRef $oObj, $vSheet)
; Parameters ....: $oObj   - Calc object opened by a preceding call to _OOoCalc_BookOpen, _OOoCalc_BookNew, or _OOoCalc_BookAttach
;                  $vSheet - Worksheet, either by index (0-based) or name. -1 returns active sheet.
; Return values .: Returns a sheet object
; Author ........: GMK
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func __OOoCalc_GetSheet(ByRef $oObj, $vSheet)
	Local $oOOoCalc_COM_ErrorHandler = ObjEvent("AutoIt.Error", __OOoCalc_ComErrorHandler_InternalFunction)
    #forceref $oOOoCalc_COM_ErrorHandler
	If Not IsObj($oObj) Then Return SetError($_OOoCalcStatus_InvalidDataType, 1, 0)
	If Not IsInt($vSheet) And Not IsString($vSheet) Then Return SetError($_OOoCalcStatus_InvalidDataType, 2, 0)
	If IsString($vSheet) And Not __OOoCalc_WorksheetIsValid($oObj, $vSheet) Then Return SetError($_OOoCalcStatus_NoMatch, 2, 0)
	Local $oSheets = $oObj.getSheets
	If Not IsObj($oSheets) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	Local $oReturn = $oObj.CurrentController.ActiveSheet
	Select
		Case IsString($vSheet)
			$oReturn = $oSheets.getByName($vSheet)
		Case IsInt($vSheet) And $vSheet > -1
			$oReturn = $oSheets.getByIndex($vSheet)
	EndSelect
	If Not IsObj($oReturn) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	Return SetError($_OOoCalcStatus_Success, 0, $oReturn)
EndFunc   ;==>__OOoCalc_GetSheet

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __OOoCalc_GetUsedRange
; Description ...: Get used range object of given Calc object
; Syntax ........: __OOoCalc_GetUsedRange(ByRef $oObj)
; Parameters ....: $oObj - Calc object opened by a preceding call to _OOoCalc_BookOpen, _OOoCalc_BookNew, or _OOoCalc_BookAttach
; Return values .: Returns a range object
; Author ........: Leagnus, GMK
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func __OOoCalc_GetUsedRange(ByRef $oObj)
	Local $oOOoCalc_COM_ErrorHandler = ObjEvent("AutoIt.Error", __OOoCalc_ComErrorHandler_InternalFunction)
    #forceref $oOOoCalc_COM_ErrorHandler
	If Not IsObj($oObj) Then Return SetError($_OOoCalcStatus_InvalidDataType, 1, 0)
	Local $oCursor = $oObj.createCursor
	If Not IsObj($oCursor) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	$oCursor.gotoStartOfUsedArea(False)
	Local $oStart = $oCursor.getRangeAddress
	If Not IsObj($oStart) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	$oCursor.gotoEndOfUsedArea(False)
	Local $oEnd = $oCursor.getRangeAddress
	If Not IsObj($oEnd) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	Local $oReturn = $oObj.getCellRangeByPosition($oStart.EndColumn, $oStart.EndRow, $oEnd.EndColumn, $oEnd.EndRow)
	If Not IsObj($oReturn) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	Return SetError($_OOoCalcStatus_Success, 0, $oReturn)
EndFunc   ;==>__OOoCalc_GetUsedRange

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __OOoCalc_RangeIsValid
; Description ...: Check if cell address or range is valid
; Syntax ........: __OOoCalc_RangeIsValid($sRange)
; Parameters ....: $sRange - Cell address or range.
; Return values .: Returns True if range is valid, False if invalid.
; Author ........: GMK
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func __OOoCalc_RangeIsValid($sRange)
	If StringRegExp($sRange, '[A-Z,a-z]+[0-9]+:[A-Z,a-z]+[0-9]+') Or (Not StringInStr($sRange, ':') And StringRegExp($sRange, '[A-Z,a-z]+[0-9]+')) Then Return SetError($_OOoCalcStatus_Success, 0, True)
	Return SetError($_OOoCalcStatus_Success, 0, False)
EndFunc   ;==>__OOoCalc_RangeIsValid

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __OOoCalc_SetPropertyValue
; Description ...: Creates a property value struct object.
; Syntax ........: __OOoCalc_SetPropertyValue($sName, $vValue)
; Parameters ....: $sName  - Property name.
;                  $vValue - Property value.
; Return values .: Returns the PropertyValue object
; Author ........: Leagnus, GMK
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func __OOoCalc_SetPropertyValue($sName, $vValue)
	Local $oOOoCalc_COM_ErrorHandler = ObjEvent("AutoIt.Error", __OOoCalc_ComErrorHandler_InternalFunction)
    #forceref $oOOoCalc_COM_ErrorHandler
	If Not IsString($sName) Then Return SetError($_OOoCalcStatus_InvalidDataType, 1, 0)
	Local $oSM = ObjCreate("com.sun.star.ServiceManager")
	If Not IsObj($oSM) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	Local $oReturn = $oSM.Bridge_GetStruct('com.sun.star.beans.PropertyValue')
	If Not IsObj($oReturn) Then Return SetError($_OOoCalcStatus_GeneralError, 0, 0)
	$oReturn.Name = $sName
	$oReturn.Value = $vValue
	Return SetError($_OOoCalcStatus_Success, 0, $oReturn)
EndFunc   ;==>__OOoCalc_SetPropertyValue

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __OOoCalc_WorksheetIsValid
; Description ...: Tests whether a sheet is valid, either by name or index.
; Syntax ........: __OOoCalc_WorksheetIsValid(ByRef $oObj, $vSheet)
; Parameters ....: $oObj   - Calc object opened by a preceding call to _OOoCalc_BookOpen, _OOoCalc_BookNew, or _OOoCalc_BookAttach
;                  $vSheet - Worksheet, either by index (0-based) or name.
; Return values .: Returns True if the worksheet is valid, False if it is invalid
; Author ........: GMK, mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func __OOoCalc_WorksheetIsValid(ByRef $oObj, $vSheet)
    Local $oOOoCalc_COM_ErrorHandler = ObjEvent("AutoIt.Error", __OOoCalc_ComErrorHandler_InternalFunction)
    #forceref $oOOoCalc_COM_ErrorHandler
    If Not IsObj($oObj) Then Return SetError($_OOoCalcStatus_InvalidDataType, 1, 0)
	If IsInt($vSheet) Then
        If $vSheet >= 0 And $vSheet <= $oObj.getSheets.count - 1 Then Return SetError($_OOoCalcStatus_Success, 0, True)
        Return SetError($_OOoCalcStatus_NoMatch, 0, False)
    ElseIf IsString($vSheet) And $vSheet <> '' Then
        Local $asWorksheets = _OOoCalc_SheetList($oObj)
        For $iSheet = 1 To $asWorksheets[0]
            If $asWorksheets[$iSheet] == $vSheet Then Return SetError($_OOoCalcStatus_Success, 0, True)
        Next
        Return SetError($_OOoCalcStatus_NoMatch, 0, False)
    EndIf
    Return SetError($_OOoCalcStatus_InvalidDataType, 0, False)
EndFunc   ;==>__OOoCalc_WorksheetIsValid

; #INTERNAL_USE_ONLY#==========================================================
; Name ..........: __OOoCalc_ComErrorHandler_InternalFunction
; Description ...: A COM error handling routine.
; Syntax.........: __OOoCalc_ComErrorHandler_InternalFunction()
; Parameters ....: None
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........;
; Example .......; [yes/no]
; ==============================================================================
Func __OOoCalc_ComErrorHandler_InternalFunction(ByRef $oCOMError)
	; If not defined ComErrorHandler_UserFunction then this function do nothing special
	; In that case you only can check @error / @extended after suspect functions
	Local $sUserFunction = _OOoCalc_ComErrorHandler_UserFunction()
	If IsFunc($sUserFunction) Then $sUserFunction($oCOMError)
EndFunc    ;==>__OOoCalc_ComErrorHandler_InternalFunction

; #FUNCTION# ====================================================================================================================
; Name ..........: __OOoCalc_ComErrorHandler_UserFunction
; Description ...: Set a UserFunctionWrapper to move the Fired COM Error Error outside UDF
; Syntax ........: _OOoCalc_ComErrorHandler_UserFunction([$fnUserFunction = Default])
; Parameters ....: $fnUserFunction- [optional] a Function. Default value is Default.
; Return values .: ErrorHandler Function
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _OOoCalc_ComErrorHandler_UserFunction($fnUserFunction = Default)
	; in case when user does not set function UDF, must use internal function to avoid AutoItError
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
		; return error as incorrect parameter was passed to this function
		Return SetError($_OOoCalcStatus_InvalidValue, 0, $fnUserFunction_Static)
	EndIf
EndFunc    ;==>_OOoCalc_ComErrorHandler_UserFunction
