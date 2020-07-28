#include <Date.au3>

; #INDEX# =======================================================================================================================
; Title .........: xlsxNative
; Version .......: 0.1
; AutoIt Version : 3.3.14.5
; Language ......: English
; Description ...: Functions to read data from Excel-xlsx files without the need of having excel installed
; Author(s) .....: AspirinJunkie
; ===============================================================================================================================



; #FUNCTION# ======================================================================================
; Name ..........: _xlsx_2Array
; Description ...: reads single worksheets of an Excel xlsx-file into an array
; Syntax ........: _xlsx_2Array(Const $sFile [, Const $sSheetNr = 1])
; Parameters ....: $sFile      - path-string of the xlsx-file
;                  $sSheetNr   - number (1-based) of the target worksheet
; Return values .: Success - Return 2D-Array with the worksheet content
;                  Failure - Return False and set @error to:
;        				@error = 1 - error importing shared-string list (@extended = @error of __xlsx_readSharedStrings)
;                              = 2 - error reading cell-values (@extended = @error of __xlsx_readCells)
;                              = 3 - error unpacking the xlsx-file (@extended = @error of __unzip)
; Author ........: AspirinJunkie
; Last changed ..: 2020-07-27
; =================================================================================================
Func _xlsx_2Array(Const $sFile, Const $sSheetNr = 1)
	Local $pthWorkDir = @TempDir & "\xlsxWork\"
	Local $pthStrings = $pthWorkDir & "xl\sharedStrings.xml"

	; unpack xlsx-file
	__unzip($sFile, $pthWorkDir, "shared*.xml sheet*.xml")
	If @error Then Return SetError(3, @error, False)

	Local $pthSheet = FileExists($pthWorkDir & "xl\worksheets\sheet.xml") ? $pthWorkDir & "xl\worksheets\sheet.xml" : $pthWorkDir & "xl\worksheets\sheet" & $sSheetNr & ".xml"

	; read strings into an 1D-array
	Local $aStrings = __xlsx_readSharedStrings($pthStrings)
	If @error Then Return SetError(1, @error, False)

	; read all cells into an 2D-array
	Local $aCells = __xlsx_readCells($pthSheet, $aStrings)
	If @error Then Return SetError(2, @error, False)

	; remove temporary data
	DirRemove($pthWorkDir, 1)

	Return $aCells
EndFunc   ;==>_xlsx_2Array


#Region xlsx specific helper functions

; #FUNCTION# ======================================================================================
; Name ..........: __xlsx_readCells
; Description ...: import xlsx worksheet values
; Syntax ........: __xlsx_readCells(Const $sFile, ByRef $aStrings)
; Parameters ....: $sFile      - path-string of the worksheet-xml
;                  $aStrings   - array with shared strings (return value of __xlsx_readSharedStrings)
; Return values .: Success - Return 2D-Array with the worksheet content
;                  Failure - Return False and set @error to:
;        				@error = 1 - cannot create XMLDOM-Object
;                              = 2 - cannot open worksheet file
;                              = 3 - cannot extract cell objects out of the xml-structure
;                              = 4 - cannot determine worksheet dimensions
; Author ........: AspirinJunkie
; Last changed ..: 2020-07-27
; =================================================================================================
Func __xlsx_readCells(Const $sFile, ByRef $aStrings)
	; Sheet einlesen
	Local $oXML = ObjCreate("Microsoft.XMLDOM")
	If Not IsObj($oXML) Then Return SetError(1, 0, False)
	$oXML.Async = False
	$oXML.resolveExternals = False
	$oXML.validateOnParse = False

	If Not $oXML.load($sFile) Then Return SetError(2, 0, False)

	If IsObj($oXML.selectSingleNode("/x:worksheet")) Then
		$sXPath = '/x:worksheet/x:sheetData/x:row/x:c'
		$sXPathValue = "x:v"
	Else
		$sXPath = '/worksheet/sheetData/row/c'
		$sXPathValue = "v"
	EndIf

	Local $oCells = $oXML.selectNodes($sXPath)
	If Not IsObj($oCells) Then Return SetError(3, 0, False)

	Local $sS, $sR, $sT, $aCoords

	; determine dimensions:
	Local $dColumnMax = 1, $dRowMax = 1

	Local $oDim = $oXML.selectSingleNode('/worksheet/dimension')
	If IsObj($oDim) Then ; we can use the range attribute
		Local $aDim = StringRegExp($oDim.GetAttribute("ref"), '([A-Z]+\d+)$', 1)
		If @error Then Return SetError(4, @error, False)

		$aDim = __xlsx_CellstringToRowColumn($aDim[0])
		$dColumnMax = $aDim[0]
		$dRowMax = $aDim[1]
	Else ; we have to determine the range ourself
		For $oCell In $oCells
			$sR = $oCell.GetAttribute("r")
			$aCoords = __xlsx_CellstringToRowColumn($sR)

			If $aCoords[0] > $dColumnMax Then $dColumnMax = $aCoords[0]
			If $aCoords[1] > $dRowMax Then $dRowMax = $aCoords[1]
		Next
	EndIf

	; create output array
	Local $aRet[$dRowMax][$dColumnMax]

	; read cell values
	Local $i = 0
	For $oCell In $oCells
		$i += 1

;~ 		$sS = $oCell.GetAttribute("s")	; style id
		$sR = $oCell.GetAttribute("r")  ; cell-coordinate
		$sT = $oCell.GetAttribute("t")  ; type of cell-value

		If $sT = "s" Then    ; value = shared string-id
			$sValue = $aStrings[Int(__xmlSingleText($oCell, $sXPathValue))]
		Else ; normal value
			$sValue = __xmlSingleText($oCell, $sXPathValue)
			If StringRegExp($sValue, '(?i)\A(?|0x\d+|[-+]?(?>\d+)(?>\.\d+)?(?:e[-+]?\d+)?)\Z') Then $sValue = Number($sValue) ; if number then convert to number type
		EndIf
		$aCoords = __xlsx_CellstringToRowColumn($sR)

		$aRet[$aCoords[1] - 1][$aCoords[0] - 1] = $sValue
	Next

	Return $aRet
EndFunc   ;==>__xlsx_readCells


; #FUNCTION# ======================================================================================
; Name ..........: __xlsx_readSharedStrings
; Description ...: import the shared string list from an xml file inside an xlsx-file
; Syntax ........: __xlsx_readSharedStrings(Const $sFile)
; Parameters ....: $sFile      - path-string of the shared-string-xml
; Return values .: Success - Return 2D-Array with the worksheet content
;                  Failure - Return False and set @error to:
;        				@error = 1 - cannot create XMLDOM-Object
;                              = 2 - cannot open worksheet file
;                              = 3 - cannot extract shared string objects out of the xml-structure
; Author ........: AspirinJunkie
; Last changed ..: 2020-07-27
; =================================================================================================
Func __xlsx_readSharedStrings(Const $sFile)
	Local $oXML = ObjCreate("Microsoft.XMLDOM")
	If Not IsObj($oXML) Then Return SetError(1, 0, False)
	$oXML.Async = False
	$oXML.resolveExternals = False
	$oXML.validateOnParse = False

	If Not $oXML.load($sFile) Then Return SetError(2, 0, False)

	Local $sXPath = IsObj($oXML.selectSingleNode("/x:sst")) ? '/x:sst/x:si/x:t' : '/sst/si/t'

	Local $oStrings = $oXML.selectNodes($sXPath)
	If Not IsObj($oStrings) Then Return SetError(3, 0, False)

	Local $aRet[$oStrings.length], $i = 0

	For $oText In $oStrings
		$aRet[$i] = $oText.text
		$i += 1
	Next

	Return $aRet
EndFunc   ;==>__xlsx_readSharedStrings

; converts excel formatted cell coordinate to array [column, row]
Func __xlsx_CellstringToRowColumn($sID)
	Local $aSplit = StringRegExp($sID, '^([A-Z]+)(\d+)$', 1)
	If @error Then Return SetError(1, @error, False)

	Local $aChars = StringSplit($aSplit[0], '', 3)
	Local $j, $dV, $aRet[2] = [0, Int($aSplit[1])]

	For $i = 0 To UBound($aChars) - 1
		$j = UBound($aChars) - $i - 1
		$aRet[0] += (Asc($aChars[$i]) - 64) * (26 ^ $j)
	Next
	Return $aRet
EndFunc   ;==>__xlsx_CellstringToRowColumn

; converts excel date value to array [year, month, day, hour, minute, second]
Func __xlsxExcel2Date($dExcelDate)
	Local $aRet[6]
	Local $fTimeRaw = $dExcelDate - Int($dExcelDate)

	_DayValueToDate(2415018.5 + Int($dExcelDate), $aRet[0], $aRet[1], $aRet[2])

	; process the time
	$aRet[3] = Floor($fTimeRaw * 24)
	$fTimeRaw -= $aRet[3] / 24    ; = Mod($fTimeRaw, 1/24)
	$aRet[4] = Floor($fTimeRaw * 1440)
	$fTimeRaw -= $aRet[4] / 1440
	$aRet[5] = Floor($fTimeRaw * 86400)

	Return $aRet
EndFunc   ;==>__xlsxExcel2Date

#EndRegion xlsx specific helper functions



#Region general helper functions

Func __unzip($sInput, $sOutput, Const $sPattern = "")
	If Not FileExists($sInput) Then Return SetError(1, @error, False)
	$sOutput = StringRegExpReplace($sOutput, '(\\*)$', '')
	If Not StringInStr(FileGetAttrib($sOutput), 'D', 1) Then
		If Not DirCreate($sOutput) Then Return SetError(1, @error, False)
	EndIf

	If FileExists("7za.exe") Then
		Local $dRet = RunWait(StringFormat('7za.exe x "%s" -o"%s" %s -r -tzip -bd -bb0 -aoa', $sInput, $sOutput, $sPattern), "", @SW_HIDE)
		Return $dRet = 0 ? True : SetError(2, $dRet, False)

	Else ; much slower
		FileCopy($sInput, @TempDir & "\tmp.zip")

		Local Static $oShell = ObjCreate("Shell.Application")
		$oShell.Namespace($sOutput).CopyHere($oShell.Namespace(@TempDir & "\tmp.zip").Items)
		FileDelete(@TempDir & "\tmp.zip")
	EndIf
	Return 1
EndFunc   ;==>__unzip

; returns single from a xml-dom-object and handles errors
Func __xmlSingleText(ByRef $oXML, Const $sXPath)
	Local $oTmp = $oXML.selectSingleNode($sXPath)
	Return IsObj($oTmp) ? $oTmp.text : SetError(1, 0, "")
EndFunc   ;==>__xmlSingleText

#Region general helper functions



