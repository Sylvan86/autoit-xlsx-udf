#include <Date.au3>
#include <Array.au3>

; #INDEX# =======================================================================================================================
; Title .........: xlsxNative
; Version .......: 0.7
; AutoIt Version : 3.3.14.5
; Language ......: English
; Description ...: Functions to read/write data from/to Excel-xlsx files without the need of having excel installed
; Author(s) .....: AspirinJunkie
; Last changed ..: 2023-01-26
; License .......: This work is free.
;                  You can redistribute it and/or modify it under the terms of the Do What The Fuck You Want To Public License, Version 2,
;                  as published by Sam Hocevar.
;                  See http://www.wtfpl.net/ for more details.
; ===============================================================================================================================


; #FUNCTION# ======================================================================================
; Name ..........: _xlsx_2Array
; Description ...: reads single worksheets of an Excel xlsx-file into an array
; Syntax ........: _xlsx_2Array(Const $sFile [, Const $iSheetNr = 1 [, $dRowFrom = 1 [, $dRowTo = Default]]])
; Parameters ....: $sFile      - path-string of the xlsx-file
;                  $iSheetNr   - id (1-based) of the target worksheet (like determined with _xlsx_getWorkSheets())
;                  $dRowFrom   - row number (1-based) where to start the extraction
;                  $dRowTo     - row number (1-based) where to stop the extraction
;                  $dColFrom   - column number (1-based) where to start the extraction
;                  $dColTo     - column number (1-based) where to stop the extraction
; Return values .: Success - Return 2D-Array with the worksheet content
;                  Failure - Return False and set @error to:
;        				@error = 1 - error importing shared-string list (@extended = @error of __xlsx_readSharedStrings)
;                              = 2 - error reading cell-values (@extended = @error of __xlsx_readCells)
;                              = 3 - error unpacking the xlsx-file (@extended = @error of __xlsx_unzip)
;                              = 4 - worksheet # doesn't exists
;                              = 5 - wrong filepath for sheet-document
; Author ........: AspirinJunkie
; Last changed ..: 2023-01-26
; =================================================================================================
Func _xlsx_2Array(Const $sFile, Const $iSheetNr = 1, $dRowFrom = 1, $dRowTo = Default, $dColFrom = 1, $dColTo = Default)
	Local $pthWorkDir = @TempDir & "\xlsxWork\"

	; correct wrong values for  $dRowFrom and $dRowTo
	If $dRowFrom < 1 Or Not IsInt($dRowFrom) Then $dRowFrom = 1
	If $dRowFrom > $dRowTo Then
		$dRowFrom = 1
		$dRowTo = Default
	EndIf

	; correct wrong values for  $dColFrom and $dColTo
	If $dColFrom < 1 Or Not IsInt($dColFrom) Then $dColFrom = 1
	If $dColFrom > $dColTo Then
		$dColFrom = 1
		$dColTo = Default
	EndIf

	; unpack xlsx-file
	__xlsx_unzip($sFile, $pthWorkDir, "xl\*.xml _rels\.rels xl\_rels\*.rels")
	If @error Then
		DirRemove($pthWorkDir, 1)
		Return SetError(3, @error, False)
	EndIf

	; determine file paths:
	Local $mFiles = __xlsx_getSubFiles($pthWorkDir)
	If @error Then
		DirRemove($pthWorkDir, 1)
		Return SetError(1, @error, False)
	EndIf

	; determine sheet file
	Local $mSheets = $mFiles["Worksheets"]
	If Not MapExists($mSheets, $iSheetNr - 1) Then
		DirRemove($pthWorkDir, 1)
		Return SetError(4, UBound($mSheets), False)
	EndIf
	Local $pthSheet = ($mSheets[$iSheetNr - 1])["File"]
	If Not FileExists($pthSheet) Then
		DirRemove($pthWorkDir, 1)
		Return SetError(5, 0, False)
	EndIf

	; read shared strings into an 1D-array
	If (Not MapExists($mFiles, "SharedStringsFile")) Or (Not FileExists($mFiles["SharedStringsFile"])) Then
		Local $aStrings[0]
	Else
		Local $aStrings = __xlsx_readSharedStrings($mFiles["SharedStringsFile"])
		If @error Then Local $aStrings[0]
	EndIf

	; read all cells into an 2D-array
	Local $aCells = __xlsx_readCells($pthSheet, $aStrings, $dRowFrom, $dRowTo, $dColFrom, $dColTo)
	If @error Then Return SetError(2, @error, False)

	; remove temporary data
	DirRemove($pthWorkDir, 1)

	Return $aCells
EndFunc   ;==>_xlsx_2Array


; #FUNCTION# ======================================================================================
; Name ..........: _xlsx_WriteFromArray
; Description ...: export a array into a xlsx-file
; Syntax ........: _xlsx_WriteFromArray(Const $sFile, ByRef $aArray)
; Parameters ....: $sFile      - output path and name for the result xlsx-file
;                  $aArray     - 1D/2D-Array
; Return values .: Success - create a xlsx-file with the content of the array
;                  Failure - Return False and set @error to:
;        				@error = 1 - $aArray is not a array
;                              = 2 - error during DirCreate (see @extended for which exactly)
;                              = 3 - error during FileWrite (see @extended for which exactly)
;                              = 4 - error zipping the file
; Author ........: AspirinJunkie
; Last changed ..: 2021-03-29
; =================================================================================================
Func _xlsx_WriteFromArray(Const $sFile, ByRef $aArray)
	Local $pthWorkDir = @TempDir & "\xlsxWork\", $dSuccess

	; convert 1D Array to 2D-Array if needed:
	If UBound($aArray, 0) = 1 Then
		Local $aA[UBound($aArray)][1]
		For $i = 0 To UBound($aA) - 1
			$aA[$i][0] = $aArray[$i]
		Next
	Else
		If Not IsArray($aArray) Then Return SetError(1, 0, False)
		Local $aA = $aArray
	EndIf

	; [Content_Types].xml
	$dSuccess = DirCreate($pthWorkDir)
	If $dSuccess = 0 Then Return SetError(2, 1, False)
	$dSuccess = FileWrite($pthWorkDir & '[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings" /><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" /><Default Extension="xml" ContentType="application/xml" /><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" /><Override PartName="/xl/worksheets/sheet.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" /></Types>')
	If $dSuccess = 0 Then Return SetError(3, 1, False)

	; .rels
	$dSuccess = DirCreate($pthWorkDir & '_rels')
	If $dSuccess = 0 Then Return SetError(2, 2, False)
	$dSuccess = FileWrite($pthWorkDir & '_rels\.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
	If $dSuccess = 0 Then Return SetError(3, 2, False)

	; workbook.xml
	$dSuccess = DirCreate($pthWorkDir & 'xl')
	If $dSuccess = 0 Then Return SetError(2, 3, False)
	$dSuccess = FileWrite($pthWorkDir & 'xl\workbook.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="1" sheetId="1" r:id="rId1" /></sheets></workbook>')
	If $dSuccess = 0 Then Return SetError(3, 3, False)

	; workbook.xml.rels
	$dSuccess = DirCreate($pthWorkDir & 'xl\_rels')
	If $dSuccess = 0 Then Return SetError(2, 4, False)
	$dSuccess = FileWrite($pthWorkDir & 'xl\_rels\workbook.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet.xml"/></Relationships>')
	If $dSuccess = 0 Then Return SetError(3, 4, False)

	; sheet.xml
	$dSuccess = DirCreate($pthWorkDir & 'xl\worksheets')
	If $dSuccess = 0 Then Return SetError(2, 5, False)
	Local $sSheet = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData>'
	For $r = 0 To UBound($aA) - 1
		$sSheet &= '<row>'
		For $c = 0 To UBound($aA, 2) - 1
			; empty cell:
			If $aA[$r][$c] = "" Then
				$sSheet &= '<c />'
			Else
				Switch VarGetType($aA[$r][$c])
					Case "Double", "Float", "Int32", "Int64" ; a number
						$sSheet &= '<c t="n"><v>' & String($aA[$r][$c]) & '</v></c>'
					Case "Bool"
						$sSheet &= '<c t="b"><v>' & Int($aA[$r][$c]) & '</v></c>'
					Case Else ; especially a string
						$sSheet &= '<c t="inlineStr"><is><t>' & __xlsx_escape4xml($aA[$r][$c]) & '</t></is></c>'
				EndSwitch
			EndIf
		Next
		$sSheet &= '</row>'
	Next
	$sSheet &= '</sheetData></worksheet>'
	$dSuccess = FileWrite($pthWorkDir & 'xl\worksheets\sheet.xml', $sSheet)
	If $dSuccess = 0 Then Return SetError(3, 5, False)

	; zip to xlsx
	FileDelete($sFile)
	__xlsx_zip($pthWorkDir & "*", $sFile)
	If @error Then Return SetError(4, @error, False)

	; remove temporary data
	DirRemove($pthWorkDir, 1)

	Return True
EndFunc   ;==>_xlsx_WriteFromArray

; #FUNCTION# ======================================================================================
; Name ..........: _xlsx_getWorkSheets
; Description ...: return the sheets ids and names of a xlsx file
; Syntax ........: _xlsx_getWorkSheets(Const $sFile)
; Parameters ....: $sFile      - path-string of the xlsx-file
; Return values .: Success - Return 2D-Array with the worksheet list where [...][0] = id and [...][1] = sheet name
;                  Failure - Return False and set @error to:
;        				@error = 1 - error unzipping the xlsx file
;                              = 2 - error loading the .rels file
;                              = 3 - error determine the workbook.xml file path
;                              = 4 - error loading the workbook.xml
;                              = 5 - error determine the sheets node inside the workbook.xml
; Author ........: AspirinJunkie
; Last changed ..: 2023-01-26
; =================================================================================================
Func _xlsx_getWorkSheets(Const $sFile)
	Local $pthWorkDir = @TempDir & "\xlsxWork\"

	__xlsx_unzip($sFile, $pthWorkDir, "xl\*.xml _rels\.rels xl\_rels\*.rels")
	If @error Then
		DirRemove($pthWorkDir, 1)
		Return SetError(3, @error, False)
	EndIf

	Local $mFiles = __xlsx_getSubFiles($pthWorkDir)
	If @error Then
		DirRemove($pthWorkDir, 1)
		Return SetError(1, @error, Null)
	EndIf

	Local $mSheets = $mFiles["Worksheets"]
	Local $aRet[UBound($mSheets)][2]

	For $i = 0 To UBound($aRet) - 1
		$mSheet = $mSheets[$i]
		$aRet[$i][0] = $i + 1
		$aRet[$i][1] = $mSheet["Name"]
	Next

	DirRemove($pthWorkDir, 1)
	Return $aRet
EndFunc   ;==>_xlsx_2Array

#Region xlsx specific helper functions

; #FUNCTION# ======================================================================================
; Name ..........: __xlsx_readCells
; Description ...: import xlsx worksheet values
; Syntax ........: __xlsx_readCells(Const $sFile, ByRef $aStrings [, $dRowFrom = 1 [, $dRowTo = Default [, $dColFrom = 1 [, $dColTo = Default]]]])
; Parameters ....: $sFile      - path-string of the worksheet-xml
;                  $aStrings   - array with shared strings (return value of __xlsx_readSharedStrings)
;                  $dRowFrom   - row number (1-based) where to start the extraction
;                  $dRowTo     - row number (1-based) where to stop the extraction
;                  $dColFrom   - column number (1-based) where to start the extraction
;                  $dColTo     - column number (1-based) where to stop the extraction
; Return values .: Success - Return 2D-Array with the worksheet content
;                  Failure - Return False and set @error to:
;        				@error = 1 - cannot read worksheet file
;                              = 2 - cannot find any <row>-elements
;                              = 3 - wrong string id in shared-string value
; Author ........: AspirinJunkie
; Last changed ..: 2023-02-27
; =================================================================================================
Func __xlsx_readCells($sFilePath, ByRef $aStrings, Const $dRowFrom = 1, $dRowTo = Default, $dColFrom = 1, $dColTo = Default)
	Local $sFileRaw = FileRead($sFilePath)
	If @error Then Return SetError(1,@error, Null)

	; determine if xml-elements have a prefix
	; if so, then change the regex-pattern
	; it's no problem to use a universal pattern but due to performance reasons the distinction is better
	If StringRegExp(StringLeft($sFileRaw, 1000), '<\w+:worksheet') Then
		Local $patRows = '(?s)<(?>\w+:)?row(?|\s+\/>|(?>\s+([^>]*))?>\s*(.+?)\s*<\/(?>\w+:)?row)'
		Local $patCells = '(?s)<(?>\w+:)?c\s*(?|\/>|([^>]*)>\s*(.*?)\s*<\/(?>\w+:)?c\b)'
		Local $patValue = '(?s)<(?>\w+:)?v>\s*([^<]*?)\s*<\/(?>\w+:)?v'
		Local $patText = '(?s)<(?>\w+:)?t>\s*([^<]*?)\s*<\/(?>\w+:)?t'
		Local $patRichText = '(?s)<(?>\w+:)?r>\s*([^<]*?)\s*<\/(?>\w+:)?r'
	Else
		Local $patRows = '(?s)<row(?|\s+\/>|(?>\s+([^>]*))?>\s*(.+?)\s*<\/row)'
		Local $patCells = '(?s)<c\s*(?|\/>|([^>]*)>\s*(.*?)\s*<\/c\b)'
		Local $patValue = '(?s)<v>\s*([^<]*?)\s*<\/v'
		Local $patText = '(?s)<t>\s*([^<]*?)\s*<\/t'
		Local $patRichText = '(?s)<r>\s*([^<]*?)\s*<\/r'
	EndIf

	; read rows first:
	Local $aRows = StringRegExp($sFileRaw, $patRows, 4, 3)
	If @error Then Return SetError(2, @error, Null)

	; pre-dimension the return array
	Local $aReturn[(IsInt($dRowTo) ? $dRowTo - $dRowFrom + 1 : Ubound($aRows) - $dRowFrom + 1)][(IsInt($dColTo) ? $dColTo - $dColFrom + 1 : 1)]

	; iterate over rows:
	Local $aRETmp, $aRowCol
	Local $iRowCount = 0, $iRow, $iRowMax = 0, $iColMax = 0
	For $aRow In $aRows

		; empty row element: <row />
		If UBound($aRow) < 2 Then
			$iRowCount += 1
			ContinueLoop
		EndIf

		If $aRow[1] <> "" Then ; row attributes
			; row number as xml-attribute of <row>
			$aRETmp = StringRegExp($aRow[1], '\br="(.*?)"', 3)
			If Not @error Then $iRow = Int($aRETmp[0]) - 1
		Else
			$iRow = $iRowCount
		EndIf

		; row range check:
		If $iRow < ($dRowFrom - 1) Then
			$iRowCount += 1
			ContinueLoop
		EndIf
		If IsInt($dRowTo) And $iRow >= $dRowTo Then ContinueLoop

		; read the cells of the row
		Local $aCells = StringRegExp($aRow[2], $patCells, 4)
		If @error Then
			$iRowCount += 1
			ContinueLoop ; skip if empty
		EndIf

		; redim output array if number of cols is bigger
		If (IsKeyword($dColTo) = 1) And UBound($aCells) > UBound($aReturn, 2) Then Redim $aReturn[UBound($aReturn)][UBound($aCells) - $dColFrom + 1]

		; iterate over cells of this row:
		Local $iColCount = 0, $iCol
		For $aCell in $aCells

			If UBound($aCell) < 3 Then
				$iColCount += 1
				ContinueLoop
			EndIf
			Local $sType = ""

			If $aCell[1] <> "" Then ; cell attributes
				$aRETmp = StringRegExp($aCell[1], '\bt="(.*?)"', 3)
				If Not @error Then $sType = $aRETmp[0]

				; check if cell has a cell coordinate attribute
				$aRETmp = StringRegExp($aCell[1], '\br="(.*?)"', 3)
				If @error Then
					$iCol = $iColCount
					If IsInt($dColTo) And $iCol >= $dColTo Then
						$iRowCount += 1
						ContinueLoop 2
					EndIf
				Else ; it has a cell coordinate attribute
					$aRowCol = __xlsx_CellstringToRowColumn($aRETmp[0])
					$iCol = $aRowCol[0] - 1

					; column range check
					If $iCol < ($dColFrom - 1) Then
						$iColCount += 1
						ContinueLoop
					EndIf
					If IsInt($dColTo) And $iCol >= $dColTo Then ContinueLoop

					; add colums to output array if cell has bigger column coordinate (possible if coordinate attribute is used)
					If ($iCol - $dColFrom + 1) >= UBound($aReturn, 2) Then Redim $aReturn[UBound($aReturn, 1)][$iCol - $dColFrom + 2]

					$iRow = $aRowCol[1] - 1
					If $iRow < ($dRowFrom - 1) Then ContinueLoop

					; add rows to output array if cell has bigger row coordinate (possible if coordinate attribute is used)
					If ($iRow - $dRowFrom + 1) > UBound($aReturn, 1) Then Redim $aReturn[$iRow - $dRowFrom + 1][UBound($aReturn, 2)]
				EndIf
			Else
				$iCol = $iColCount

				; skip rest of current row complete if column is bigger then the user wants
				If IsInt($dColTo) And $iCol >= $dColTo Then
					$iRowCount += 1
					ContinueLoop 2
				EndIf
			EndIf

			; column range check
			If $iCol < ($dColFrom - 1) Then
				$iColCount += 1
				ContinueLoop
			EndIf

			; treat according to cell type
			$aRETmp = StringRegExp($aCell[2], $patValue, 3)
			Local $sValue = @error ? "" : $aRETmp[0]
			Switch $sType
				Case "s" ; value = shared string-id
					$sValue = Int($sValue)
					If $sValue > UBound($aStrings) Then Return SetError(3, $sValue, False)
					$sValue = $aStrings[$sValue]
				Case "inlineStr" ; inline string
					; The value is not in <v> but in <is></is> where the structure is the same as in a <si> element of sharedStrings.xml So the value is either in <t></t> for normal text or in <r></r> for rich text
					$aRETmp = StringRegExp($aCell[2], '<t>\K[^<]*', 3)
					If @error Then ContinueLoop
					If UBound($aRETmp) = 1 Then ; single line value
						$sValue = $aRETmp[0]
					Else ; multiline value
						For $j = 0 To UBound($aRETmp) - 1
							$sValue &= $aRETmp[$j] & @CRLF
						Next
						$sValue = StringTrimRight($sValue, 2)
					EndIf
				Case "str" ; formula
					; here the formula itself is in an [optional] <f></f> while the last calculated value is normally in <v></v> - so nothing to do because $sValue has already value of <v>...</v>
				Case "n" ; number (integers, floats, dates, times)
					$sValue = Number($sValue)
				; Case "e" ; error
				Case "b" ; boolean
					$sValue = $sValue = True
				Case Else ; normal value
					If StringRegExp($sValue, '(?i)\A(?|0x\d+|[-+]?(?>\d+)(?>\.\d+)?(?:e[-+]?\d+)?)\Z') Then $sValue = Number($sValue) ; if number then convert to number type
			EndSwitch

			; determine empty rows and columns
			If $sValue <> "" Then
				If $iRow > $iRowMax Then $iRowMax = $iRow
				If $iCol > $iColMax Then $iColMax = $iCol
			EndIf

			; handle value
			$aReturn[$iRow - $dRowFrom + 1][$iCol - $dColFrom + 1] = $sValue

			$iColCount += 1
		Next

		$iRowCount += 1
	Next

	; cut off empty rows
	If UBound($aReturn, 1) > ($iRowMax - $dRowFrom + 2) Then Redim $aReturn[$iRowMax - $dRowFrom + 2][UBound($aReturn, 2)]

	Return $aReturn
EndFunc


; #FUNCTION# ======================================================================================
; Name ..........: __xlsx_getSubFiles
; Description ...: determine the file paths of the sheet files and the sharedStrings file
; Syntax ........: __xlsx_getSubFiles([$pthWorkDir = @TempDir & "\xlsxWork\"])
; Parameters ....: $pthWorkDir - root folder of extracted xlsx file
; Return values .: Success - Return 2D-Array with the worksheet content
;                  Failure - Return False and set @error to:
;        				@error = 1 - no global .rels file found
;                              = 2 - no workbook file determined
;                              = 3 - no .rel-file connected to the workbook found
;                              = 4 - error opening the workbook file
;                              = 5 - error creating the xml-object
; Author ........: AspirinJunkie
; Last changed ..: 2023-01-26
; =================================================================================================
Func __xlsx_getSubFiles($pthWorkDir = @TempDir & "\xlsxWork\")
	If StringRight($pthWorkDir, 1) <> "\" Then $pthWorkDir &= "\"

	; unpack xlsx-file
	Local $oXML = __xlsx_getXMLObject()
	If @error Then Return SetError(5, @error, 0)

	; Map for paths
	Local $mRet[]

	;  determine main workbook file and their path
	Local $sSubPath = $pthWorkDir, $sWorkbook
	If Not $oXML.load($pthWorkDir & "\_rels\.rels") Then Return SetError(1, 0, False)
	Local $sPre = $oXML.documentElement.prefix
	If $sPre <> "" Then $sPre &= ":"
	For $oRS In $oXML.selectNodes('//' & $sPre & 'Relationship')
		If StringRegExp($oRS.getAttribute("Type"), 'officeDocument$') Then
			$mRet["WorkbookFileName"] = StringRegExpReplace($oRS.getAttribute("Target"), '^.+\/', '')
			$sSubPath = StringRegExpReplace($oRS.getAttribute("Target"), '\/?(.+?)\/[^\/]+$', '$1') & "\"
			$mRet["WorkbookPath"] = $pthWorkDir & $sSubPath
			ExitLoop
		EndIf
	Next
	If $mRet["WorkbookFileName"] = "" Then Return SetError(2, 0, False)

	; determine the relevant files
	If Not $oXML.load($pthWorkDir & $sSubPath & "_rels\" & $mRet.WorkbookFileName & ".rels") Then Return SetError(3, 0, False)
	Local $sPre = $oXML.documentElement.prefix
	If $sPre <> "" Then $sPre &= ":"
	Local $sTarget, $sId, $mSheetsByID[]
	For $oRS In $oXML.selectNodes('//' & $sPre & 'Relationship')
		$sId = $oRS.getAttribute("Id")
		$sTarget = $oRS.getAttribute("Target")

		; differentiate by type
		Switch StringRegExpReplace($oRS.getAttribute("Type"), '.*?([^\/]+)$', '$1', 1)
			Case "worksheet"
				$mSheetsByID[$sId] = __xlsx_Target2AbsolutePath($sTarget, $pthWorkDir, $sSubPath)
			Case "sharedStrings"
				$mRet["SharedStringsFile"] = __xlsx_Target2AbsolutePath($sTarget, $pthWorkDir, $sSubPath)
			Case Else
				ContinueLoop
		EndSwitch
	Next

	; determine the the sheet names and their order
	If Not $oXML.load($mRet["WorkbookPath"] & $mRet.WorkbookFileName) Then Return SetError(4, 0, False)
	Local $sPre = $oXML.documentElement.prefix
	If $sPre <> "" Then $sPre &= ":"
	Local  $mSheets[]
	For $oSheet In $oXML.selectNodes('//' & $sPre & 'sheets/' & $sPre & 'sheet')
		Local $mSheet[]

		$mSheet["Name"] = $oSheet.getAttribute("name")
		$mSheet["ID"] = $oSheet.getAttribute("sheetId") ; senseless information?
		$mSheet["File"] = MapExists($mSheetsByID, $oSheet.getAttribute("r:id")) ? $mSheetsByID[$oSheet.getAttribute("r:id")] : Null

		MapAppend($mSheets, $mSheet)
	Next
	$mRet["Worksheets"] = $mSheets

	Return SetExtended(UBound($mSheets), $mRet)
EndFunc   ;==>__xlsx_getSubFiles


; #FUNCTION# ======================================================================================
; Name ..........: __xlsx_readSharedStrings
; Description ...: import the shared string list from an xml file inside an xlsx-file
; Syntax ........: __xlsx_readSharedStrings(Const $sFile)
; Parameters ....: $sFile      - path-string of the shared-string-xml
; Return values .: Success - Return 2D-Array with the worksheet content
;                  Failure - Return False and set @error to:
;        				@error = 1 - error reading $sFile
;                              = 2 - no <si>-elements found = no shared strings
; Author ........: AspirinJunkie
; Last changed ..: 2023-02-27
; =================================================================================================
Func __xlsx_readSharedStrings(Const $sFile)
	Local $sFileRaw = FileRead($sFile)
	IF @error THen Return SetError(1, @error, Null)

	Local $aSI = StringRegExp($sFileRaw, '(?s)<si>(.+?)<\/si>', 3)
	If @error THen Return SetError(2, @error, Null)

	; dimension output array
	Local $aRet[UBound($aSI)]

	Local $aText
	For $i = 0 To UBound($aSI) - 1
		$aText = StringRegExp($aSI[$i], '<t>\K[^<]*', 3)
		If @error Then ContinueLoop
		If UBound($aText) = 1 Then
			$aRet[$i] = $aText[0]
		Else
			For $j = 0 To UBound($aText) - 1
				$aRet[$i] &= $aText[$j] & @CRLF
			Next
			$aRet[$i] = StringTrimRight($aRet[$i], 2)
		EndIf
	Next

	Return $aRet
EndFunc   ;==>__xlsx_readSharedStrings

; converts excel formatted cell coordinate to array [column, row]
Func __xlsx_CellstringToRowColumn($sID)
	Local $aSplit = StringRegExp($sID, "^([A-Z]+)(\d+)$", 1)
	If @error Then Return SetError(1, @error, False)
	Local $iCol = 0
	For $i = 1 To StringLen($aSplit[0])
		$iCol = $iCol * 26 + Asc(StringMid($aSplit[0], $i, 1)) - 64
	Next
	Local $aRet = [$iCol, Int($aSplit[1])]
	Return $aRet
EndFunc   ;==>__xlsx_CellstringToRowColumn


; #FUNCTION# ======================================================================================
; Name ..........: __xlsxExcel2Date
; Description ...: convert a excel date-value into an editable form (array of components or formatted string)
;                  with standard parameters the local date and time format is used
; Syntax ........: __xlsxExcel2Date($dExcelDate[, Const $sType = Default[, Const $iFlags = 0x01[, Const $sFormat = ""[, Const $iFlagsTime = 0[, Const $sFormatTime = ""]]]]])
; Parameters ....: $dExcelDate     - excel date value as number or string
;                  $sType          - Default: an array[6]: [year, month, day, hour, minute, second]
;                                    "date": string with formatted date
;                                    "time": string with formatted time
;                                    "datetime": string with formatted date and time
;                  $iFlags         - parameter $iflags of _WinAPI_GetDateFormat()
;                  $sFormat        - parameter $sFormat of _WinAPI_GetDateFormat()
;                  $iFlagsTime     - parameter $iflags of _WinAPI_GetTimeFormat()
;                  $sFormatTime    - parameter $sFormat of _WinAPI_GetTimeFormat()
; Return values .: Success - Return array or string (depending on the selected $sType)
;                  Failure - Return False and set @error to:
;        				@error = 1 - invalid value for $sType
; Author ........: AspirinJunkie
; Last changed ..: 2020-08-07
; =================================================================================================
Func __xlsxExcel2Date($dExcelDate, Const $sType = Default, Const $iFlags = 0x01, Const $sFormat = "", Const $iFlagsTime = 0, Const $sFormatTime = "")
	Switch $sType
		Case Default
			Local $aRet[6], $fTimeRaw = $dExcelDate - Int($dExcelDate)

			_DayValueToDate(2415018.5 + Int($dExcelDate), $aRet[0], $aRet[1], $aRet[2])

			; process the time
			$aRet[3] = Floor($fTimeRaw * 24)
			$fTimeRaw -= $aRet[3] / 24 ; = Mod($fTimeRaw, 1/24)
			$aRet[4] = Floor($fTimeRaw * 1440)
			$fTimeRaw -= $aRet[4] / 1440
			$aRet[5] = Floor($fTimeRaw * 86400)

			Return $aRet
		Case "date"
			Local $y, $m, $d
			_DayValueToDate(2415018.5 + Int($dExcelDate), $y, $m, $d)
			Local $tDateTime = _Date_Time_EncodeSystemTime($m, $d, $y)
			Return _WinAPI_GetDateFormat(0x0400, $tDateTime, $iFlags, $sFormat)
		Case "time"
			Local $h, $min, $s, $fTimeRaw = $dExcelDate - Int($dExcelDate)
			; process the time
			$h = Floor($fTimeRaw * 24)
			$fTimeRaw -= $h / 24 ; = Mod($fTimeRaw, 1/24)
			$min = Floor($fTimeRaw * 1440)
			$fTimeRaw -= $min / 1440
			$s = Floor($fTimeRaw * 86400)
			Return StringFormat("%02d:%02d:%02d", $h, $min, $s)
		Case "datetime"
			Local $y, $m, $d, $h, $min, $s, $fTimeRaw = $dExcelDate - Int($dExcelDate)
			; process the time
			$h = Floor($fTimeRaw * 24)
			$fTimeRaw -= $h / 24 ; = Mod($fTimeRaw, 1/24)
			$min = Floor($fTimeRaw * 1440)
			$fTimeRaw -= $min / 1440
			$s = Floor($fTimeRaw * 86400)

			_DayValueToDate(2415018.5 + Int($dExcelDate), $y, $m, $d)
			Local $tDateTime = _Date_Time_EncodeSystemTime($m, $d, $y, $h, $min, $s)
			Return _WinAPI_GetDateFormat(0x0400, $tDateTime, $iFlags, $sFormat) & " " & _WinAPI_GetTimeFormat(0, $tDateTime, $iFlagsTime, $sFormatTime)
		Case Else
			Return SetError(1, 0, False)
	EndSwitch
EndFunc   ;==>__xlsxExcel2Date

; function to share one single xmldom-object over the functions but without beeing a global variable
Func __xlsx_getXMLObject()
	Local Static $c = 0
	Local Static $oX = ObjCreate("Microsoft.XMLDOM")
	If @error Then Return SetError(@error, @extended, 0)

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

; convert the target-attributes into an absolute file path
Func __xlsx_Target2AbsolutePath($sTarget, $pthWorkDir = @TempDir & "\xlsxWork\", $sSubPath = "xl")
	If StringRight($pthWorkDir, 1) <> "\" Then $pthWorkDir &= "\"

	$sTarget = StringRegExpReplace($sTarget, '\/?(.+?)\/([^\/]+)$', '$1/$2')
	If Not FileExists($pthWorkDir & $sTarget) Then $sTarget = $sSubPath & "/" & $sTarget

	Return $pthWorkDir & $sTarget
EndFunc

#EndRegion xlsx specific helper functions



#Region general helper functions

Func __xlsx_unzip($sInput, $sOutput, Const $sPattern = "")
	; TODO: maybe powershell "	" or jar is faster than shell.application
	; RunWait(StringFormat('powershell.exe -Command "Expand-Archive -LiteralPath ''%s'' -DestinationPath ''%s'' -Force"', $sFile, $sOutPath), "", @SW_HIDE)
	; but!: expand-archive can only handle files with zip-ending. so still have to rename or copy to a zip-file

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
EndFunc   ;==>__xlsx_unzip

; compress a folder content into a zip-file (file extension doesn't have to be .zip)
Func __xlsx_zip($sInput, $sOutput)
	Local $iExitCode

	;  $iExitCode = RunWait(StringFormat('zip.exe -qr -9 "%s" .'', $sInput, $sOutput & ".zip"), "", @SW_HIDE)
	;  If Not @error And $iExitCode = 0 Then

	; "jar -cMf targetArchive.zip sourceDirectory"

	$iExitCode = RunWait(StringFormat('7za.exe a -mm=Deflate -mfb=258 -mpass=15 -r "%s" "%s"', $sOutput, $sInput), "", @SW_HIDE)
	If @error Or $iExitCode <> 0 Then

		$iExitCode = RunWait(StringFormat('powershell.exe -Command "Compress-Archive ''%s'' ''%s''"', $sInput, $sOutput & ".zip"), "", @SW_HIDE)
		If Not @error And $iExitCode = 0 Then
			FileMove($sOutput & ".zip", $sOutput)
		EndIf
	EndIf
	Return 1
EndFunc   ;==>__xlsx_zip

; escape special xml characters
Func __xlsx_escape4xml($sString)
	$sString = StringReplace($sString, '&', '&amp;', 0, 1)
	$sString = StringReplace($sString, '"', '&quot;', 0, 1)
	$sString = StringReplace($sString, "'", '&apos;', 0, 1)
	$sString = StringReplace($sString, '<', '&lt;', 0, 1)
	$sString = StringReplace($sString, '>', '&gt;', 0, 1)
	Return $sString
EndFunc

; returns value of a single xml-dom-node and handles errors
Func __xmlSingleText(ByRef $oXML, Const $sXPath)
	Local $oTmp = $oXML.selectSingleNode($sXPath)
	Return IsObj($oTmp) ? $oTmp.text : SetError(1, 0, "")
EndFunc   ;==>__xmlSingleText

#Region general helper functions



