#include "xlsxNative.au3"

; create 2D array
Global $A[][] = [[1, 2, 3, 4, 5], ["", "6", 7, "", "8"], [], [9, "", "10", 11, True]]
_ArrayDisplay($A, "source array")

; convert the Array into a xlsx-file:
_xlsx_WriteFromArray(@ScriptDir & "\Text.xlsx", $A)

; read this xlsx-file into a 2D-Array:
$aSheet = _xlsx_2Array(@ScriptDir & "\Text.xlsx")
_ArrayDisplay($aSheet, "imported data from xlsx")

;  ; determine the worksheets of a file
;  Local $aSheets = _xlsx_getWorkSheets(@ScriptDir & "\Text.xlsx")
;  _ArrayDisplay($aSheets, "Sheet List", "", 64 + 32 , Default, "ID|Name|sheetID")



;  #include "xlsxNative.au3"

;  Local $sFile = @ScriptDir & "\Test.xlsx"

;  ; determine the worksheets of a file
;  Local $aSheets = _xlsx_getWorkSheets($sFile)
;  _ArrayDisplay($aSheets, "Sheet List", "", 64 + 32 , Default, "ID|Name")