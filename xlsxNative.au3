#include <Date.au3>

; #INDEX# =======================================================================================================================
; Title .........: xlsxNative
; Version .......: 0.2.1
; AutoIt Version : 3.3.14.5
; Language ......: English
; Description ...: Functions to read data from Excel-xlsx files without the need of having excel installed
; Author(s) .....: AspirinJunkie
; Last changed ..: 2020-07.31
; ===============================================================================================================================



; #FUNCTION# ======================================================================================
; Name ..........: _xlsx_2Array
; Description ...: reads single worksheets of an Excel xlsx-file into an array
; Syntax ........: _xlsx_2Array(Const $sFile [, Const $sSheetNr = 1 [, $dRowFrom = 1 [, $dRowTo = Default]]])
; Parameters ....: $sFile      - path-string of the xlsx-file
;                  $sSheetNr   - number (1-based) of the target worksheet
;                  $dRowFrom   - row number (1-based) where to start the extraction
;                  $dRowTo   - row number (1-based) where to stop the extraction
; Return values .: Success - Return 2D-Array with the worksheet content
;                  Failure - Return False and set @error to:
;        				@error = 1 - error importing shared-string list (@extended = @error of __xlsx_readSharedStrings)
;                              = 2 - error reading cell-values (@extended = @error of __xlsx_readCells)
;                              = 3 - error unpacking the xlsx-file (@extended = @error of __unzip)
; Author ........: AspirinJunkie
; Last changed ..: 2020-07-27
; =================================================================================================
Func _xlsx_2Array(Const $sFile, Const $sSheetNr = 1, $dRowFrom = 1, $dRowTo = Default)
	Local $pthWorkDir = @TempDir & "\xlsxWork\"
	Local $pthStrings = $pthWorkDir & "xl\sharedStrings.xml"

	; correct wrong values for  $dRowFrom and $dRowTo
	If $dRowFrom < 1 Or Not IsInt($dRowFrom) Then $dRowFrom = 1
	If $dRowFrom > $dRowTo Then
		$dRowFrom = 1
		$dRowTo = Default
	EndIf

	; unpack xlsx-file
	__unzip($sFile, $pthWorkDir, "shared*.xml sheet.xml sheet" & $sSheetNr & ".xml")
	If @error Then Return SetError(3, @error, False)

	Local $pthSheet = FileExists($pthWorkDir & "xl\worksheets\sheet.xml") ? $pthWorkDir & "xl\worksheets\sheet.xml" : $pthWorkDir & "xl\worksheets\sheet" & $sSheetNr & ".xml"

	; read strings into an 1D-array
	Local $aStrings = __xlsx_readSharedStrings($pthStrings)
	If @error Then Return SetError(1, @error, False)

	; read all cells into an 2D-array
	Local $aCells = __xlsx_readCells($pthSheet, $aStrings, $dRowFrom, $dRowTo)
	If @error Then Return SetError(2, @error, False)

	; remove temporary data
	DirRemove($pthWorkDir, 1)

	Return $aCells
EndFunc   ;==>_xlsx_2Array


#Region xlsx specific helper functions

; #FUNCTION# ======================================================================================
; Name ..........: __xlsx_readCells
; Description ...: import xlsx worksheet values
; Syntax ........: __xlsx_readCells(Const $sFile, ByRef $aStrings [, $dRowFrom = 1 [, $dRowTo = Default]])
; Parameters ....: $sFile      - path-string of the worksheet-xml
;                  $aStrings   - array with shared strings (return value of __xlsx_readSharedStrings)
;                  $dRowFrom   - row number (1-based) where to start the extraction
;                  $dRowTo   - row number (1-based) where to stop the extraction
; Return values .: Success - Return 2D-Array with the worksheet content
;                  Failure - Return False and set @error to:
;        				@error = 1 - cannot create XMLDOM-Object
;                              = 2 - cannot open worksheet file
;                              = 3 - cannot extract cell objects out of the xml-structure
;                              = 4 - cannot determine worksheet dimensions
;                              = 5 - wrong string id in shared-string value
; Author ........: AspirinJunkie
; Last changed ..: 2020-07-27
; =================================================================================================
Func __xlsx_readCells(Const $sFile, ByRef $aStrings, Const $dRowFrom = 1, $dRowTo = Default)
	Local $oXML = __xlsx_getXMLObject()

	If Not $oXML.load($sFile) Then Return SetError(2, 0, False)

	; determine the namespace prefix:
	Local $sPre = $oXML.documentElement.prefix
	If $sPre <> "" Then $sPre &= ":"

; select the cell-nodes (but only if they have a value-child)
	Local $oCells = $oXML.selectNodes('/' & $sPre & 'worksheet/' & $sPre & 'sheetData/' & $sPre & 'row/' & $sPre & 'c[' & $sPre & 'v]')
	If Not IsObj($oCells) Then Return SetError(3, 0, False)

	; determine dimensions:
	Local $dColumnMax = 1, $dRowMax = 1, $sR, $aCoords
	For $oCell In $oCells
		$sR = $oCell.GetAttribute("r")
		$aCoords = __xlsx_CellstringToRowColumn($sR)
		$oCell.SetAttribute("zeile", $aCoords[1])
		$oCell.SetAttribute("spalte", $aCoords[0])
		If $aCoords[0] > $dColumnMax Then $dColumnMax = $aCoords[0]
		If $aCoords[1] > $dRowMax Then $dRowMax = $aCoords[1]
	Next

	; create output array
	If $dRowTo <> Default Then $dRowMax = $dRowTo > $dRowMax ? $dRowMax : $dRowTo
	Local $aRet[$dRowMax - $dRowFrom + 1][$dColumnMax]

	; read cell values
	Local $i = 0, $sTmp
	For $oCell In $oCells
		$i += 1

		Local $dRow = $oCell.GetAttribute("zeile")
		If $dRow < $dRowFrom Or $dRow > $dRowMax Then ContinueLoop

		If $oCell.GetAttribute("t") = "s" Then    ; value = shared string-id
			$sTmp = Int(__xmlSingleText($oCell, $sPre & 'v'))
			If $sTmp > UBound($aStrings) Then Return SetError(5, $sTmp, False)
			$sValue = $aStrings[$sTmp]
		Else ; normal value
			$sValue = __xmlSingleText($oCell, $sPre & 'v')
			If StringRegExp($sValue, '(?i)\A(?|0x\d+|[-+]?(?>\d+)(?>\.\d+)?(?:e[-+]?\d+)?)\Z') Then $sValue = Number($sValue) ; if number then convert to number type
		EndIf
		$aRet[$oCell.GetAttribute("zeile") - $dRowFrom][$oCell.GetAttribute("spalte") - 1] = $sValue
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
	Local $oXML = __xlsx_getXMLObject()

	If Not $oXML.load($sFile) Then Return SetError(2, 0, False)

	Local $sPre = $oXML.documentElement.prefix
	If $sPre <> "" Then $sPre &= ":"

	Local $oStrings = $oXML.selectNodes('/' & $sPre & 'sst/' & $sPre & 'si')
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

; function to share one single xmldom-object over the functions but without beeing a global variable
Func __xlsx_getXMLObject()
	Local Static $c = 0
	Local Static $oX = ObjCreate("Microsoft.XMLDOM")

	If $c = 0 Then
		With $oX
			.Async = False
			.resolveExternals = False
			.validateOnParse = False
			.setProperty("ForcedResync", False)
		EndWith
		$c = 1
	EndIf
	Return $oX
EndFunc   ;==>__xlsx_getXMLObject

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
		$oShell.Namespace($sOutput).CopyHere($oShell.Namespace(@TempDir & "\tmp.zip").Items, 4 + 16)
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



