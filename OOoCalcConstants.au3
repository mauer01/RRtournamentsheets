#include-once

#region ;**** Variables ****
; #ENUMS# =======================================================================================================================
Global Enum _; Error Status Types
		$_OOoCalcStatus_Success = 0, _
		$_OOoCalcStatus_GeneralError, _
		$_OOoCalcStatus_ComError, _
		$_OOoCalcStatus_InvalidDataType, _
		$_OOoCalcStatus_InvalidValue, _
		$_OOoCalcStatus_NoMatch
Global Enum Step * 2 _; NotificationLevel
		$_OOoCalcNotifyLevel_None = 0, _
		$_OOoCalcNotifyNotifyLevel_Warning = 1, _
		$_OOoCalcNotifyNotifyLevel_Error, _
		$_OOoCalcNotifyNotifyLevel_ComError
Global Enum Step * 2 _; NotificationMethod
		$_OOoCalcNotifyMethod_Silent = 0, _
		$_OOoCalcNotifyMethod_Console = 1, _
		$_OOoCalcNotifyMethod_ToolTip, _
		$_OOoCalcNotifyMethod_MsgBox
; ===============================================================================================================================

; #CONSTANTS# ===================================================================================================================
Global Enum Step * 2 _ ; Cell flag constants derived from http://www.openoffice.org/api/docs/common/ref/com/sun/star/sheet/CellFlags.html
		$VALUE = 1, _
		$DATETIME, _
		$STRING, _
		$ANNOTATION, _
		$FORMULA, _
		$HARDATTR, _
		$STYLES, _
		$OBJECTS, _
		$EDITATTR, _
		$FORMATTED

Global Enum _
		$NUMBER_STANDARD = 0, _
		$NUMBER_INT, _
		$NUMBER_DEC2, _
		$NUMBER_1000INT, _
		$NUMBER_1000DEC2, _
		$NUMBER_SYSTEM, _
		$SCIENTIFIC_000E000, _
		$SCIENTIFIC_000E00, _
		$PERCENT_INT = 10, _
		$PERCENT_DEC2, _
		$FRACTION_1, _
		$FRACTION_2, _
		$CURRENCY_1000INT = 20, _
		$CURRENCY_1000DEC2, _
		$CURRENCY_1000INT_RED, _
		$CURRENCY_1000DEC2_RED, _
		$CURRENCY_1000DEC2_CCC, _
		$CURRENCY_1000DEC2_DASHED, _
		$DATE_SYSTEM_SHORT = 30, _
		$DATE_DEF_NNDDMMMYY, _
		$DATE_SYS_MMYY, _
		$DATE_SYS_DDMMM, _
		$DATE_MMMM, _
		$DATE_QQJJ, _
		$DATE_SYS_DDMMYYYY, _
		$DATE_SYS_DDMMYY, _
		$DATE_SYS_NNNNDMMMMYYYY, _
		$DATE_SYS_DMMMYY, _
		$TIME_HHMM, _
		$TIME_HHMMSS, _
		$TIME_HHMMAMPM, _
		$TIME_HHMMSSAMPM, _
		$TIME_HH_MMSS, _
		$TIME_MMSS00, _
		$TIME_HH_MMSS00, _
		$DATETIME_SYSTEM_SHORT_HHMM = 50


Global Enum _ ; CellHoriJustify enumerations taken from http://www.openoffice.org/api/docs/common/ref/com/sun/star/table/CellHoriJustify.html
		$STANDARD = 0, _
		$LEFT, _
		$CENTER, _
		$RIGHT, _
		$BLOCK, _
		$REPEAT

Global Enum _ ; DuplexMode enumerations taken from http://www.openoffice.org/api/docs/common/ref/com/sun/star/view/DuplexMode.html
		$UNKNOWN = 0, _
		$OFF, _
		$LONGEDGE, _
		$SHORTEDGE
; ===============================================================================================================================
#endregion ;**** Variables ****