#include-once
#include <Date.au3>
#include <Array.au3>

; #INDEX# =======================================================================================================================
; Title .........: xlsxNative
; Version .......: 0.8.0
; AutoIt Version : 3.3.18.0
; Language ......: English
; Description ...: Functions to read/write data from/to Excel-xlsx files without the need of having excel installed
; Author(s) .....: AspirinJunkie
; Last changed ..: 2025-10-09
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
	Local $aStrings[0], $iErr

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
		$iErr = @error
		DirRemove($pthWorkDir, 1)
		Return SetError(3, $iErr, False)
	EndIf

	; determine file paths:
	Local $mFiles = __xlsx_getSubFiles($pthWorkDir)
	If @error Then
		$iErr = @error
		DirRemove($pthWorkDir, 1)
		Return SetError(1, $iErr, False)
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
	If (MapExists($mFiles, "SharedStringsFile")) And (FileExists($mFiles["SharedStringsFile"])) Then
		$aStrings = __xlsx_readSharedStrings($mFiles["SharedStringsFile"])
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
	Local Const $patDATETIME = '(?:^(?<date>(?>19|20)\d\d-[01]\d-(?>[012]\d|3[01]))[T ]?(?<time>\d\d\:\d\d(?>\:\d\d(?>[\.,]\d+))?)$|\g<date>|\g<time>)'
	Local $pthWorkDir = @TempDir & "\xlsxWork\", $dSuccess
	Local $vVal, $bDates = False, $aRE

	; convert 1D Array to 2D-Array if needed:
	Local $aA[0]
	If UBound($aArray, 0) = 1 Then
		Redim $aA[UBound($aArray)][1]
		For $i = 0 To UBound($aA) - 1
			$aA[$i][0] = $aArray[$i]
		Next
	Else
		If Not IsArray($aArray) Then Return SetError(1, 0, False)
		$aA = $aArray
	EndIf

	; determine infos about the position of empty cells (helps to reduce file size)
	Local $mArrayInfos = __xlsx_determineEmptyArrayElements($aA)
	If @error Then Return SetError(5, @error, False)
	Local $aLastIDs = $mArrayInfos.aLastIDs, $iFirstWrittenRow = $mArrayInfos.iFirstWrittenRow, $iLastWrittenRow = $mArrayInfos.iLastWrittenRow

	; sheet.xml
	$dSuccess = DirCreate($pthWorkDir & 'xl')
	If $dSuccess = 0 Then Return SetError(2, 3, False)
	$dSuccess = DirCreate($pthWorkDir & 'xl\w')
	If $dSuccess = 0 Then Return SetError(2, 5, False)
	Local $sSheet = '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>'
	Local $iLastNonEmptyRow = -1, $iLastNonEmptyCol, $sCellElemString
	For $r = $iFirstWrittenRow To $iLastWrittenRow

		; don't process empty lines
		If $aLastIDs[$r] = 0 Then ContinueLoop

		; if there have been empty lines before - use the "r" attribute
		If $iLastNonEmptyRow = $r - 1 Then
			$sSheet &= '<row>'
		Else
			$sSheet &= '<row r="' & $r + 1 & '">'
		EndIf

		; variable to determine that this line was not empty
		$iLastNonEmptyRow = $r
		$iLastNonEmptyCol = -1 ; reset

		For $c = 0 To $aLastIDs[$r] - 1
			$vVal = $aA[$r][$c]

			; empty or NULL cell:
			If $vVal = "" Or (IsKeyword($vVal) = 2) Then ContinueLoop

			; handle empty cells before the current cell
			$sCellElemString = ($iLastNonEmptyCol = $c - 1) ? "c" : 'c r="' & __xlsx_createCellString($r+1, $c+1) & '"'
			$iLastNonEmptyCol = $c
				
			Switch VarGetType($vVal)
				Case "Double", "Float", "Int32", "Int64" ; a number (t="n" is default so to save space leave it)
					$sSheet &= '<' & $sCellElemString & '><v>' & String($vVal) & '</v></c>'

				Case "Bool"
					$sSheet &= '<' & $sCellElemString & ' t="b"><v>' & Int($vVal) & '</v></c>'

				Case Else ; especially a string or a function

					; maybe a function
					If StringLeft($vVal, 1) = "=" Then
						If StringMid($vVal, 2, 1) = "=" Then ; escape leading "=" through doubling --> handle as normal string value
							$sSheet &= '<' & $sCellElemString & ' t="inlineStr"><is><t>' & __xlsx_escape4xml(StringTrimLeft($vVal, 1)) & '</t></is></c>'

						Else ; handle as Excel function value
							$sSheet &= '<' & $sCellElemString & '><f>' & __xlsx_escape4xml(StringTrimLeft($vVal, 1)) & '</f></c>'

						EndIf

						ContinueLoop
					EndIf

					; check if string contains a date/time
					$aRE = StringRegExp($vVal, $patDATETIME, 3)
					If @error Then 
						$sSheet &= '<' & $sCellElemString & ' t="inlineStr"><is><t>' & __xlsx_escape4xml($vVal) & '</t></is></c>' ; normal string
						ContinueLoop
					EndIf

					Switch UBound($aRE, 1)
						Case 1 ; a date or time only
							If StringLen($aRE[0]) < 6 Then ; time only
								$sSheet &= StringFormat('<' & $sCellElemString & ' s="2"><v>%f</v></c>', StringLeft($vVal, 2) / 24.0 + StringMid($vVal, 4, 2) / 1440.0)

							Else ; date value
								$sSheet &= '<' & $sCellElemString & ' t="d" s="1"><v>' & $vVal & '</v></c>' ; date value

							EndIf
							$bDates = True

						Case 2 ; date + time 
							$sSheet &= '<' & $sCellElemString & ' t="d" s="3"><v>' & $vVal & '</v></c>'
							$bDates = True

					EndSwitch
			EndSwitch
		Next
		$sSheet &= '</row>'
	Next
	$sSheet &= '</sheetData></worksheet>'

	$dSuccess = FileWrite($pthWorkDir & 'xl\w\s.xml', $sSheet)
	If $dSuccess = 0 Then Return SetError(3, 5, False)
	
	; [Content_Types].xml
	$dSuccess = DirCreate($pthWorkDir)
	If $dSuccess = 0 Then Return SetError(2, 1, False)
	$dSuccess = FileWrite($pthWorkDir & '[Content_Types].xml', StringFormat( _
		'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" /><Override PartName="/xl/w.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" /><Override PartName="/xl/w/s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" />%s</Types>', _ 
		$bDates ? '<Override PartName="/xl/st.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' : ''))
	If $dSuccess = 0 Then Return SetError(3, 1, False)

	; .rels
	$dSuccess = DirCreate($pthWorkDir & '_rels')
	If $dSuccess = 0 Then Return SetError(2, 2, False)
	$dSuccess = FileWrite($pthWorkDir & '_rels\.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/w.xml"/></Relationships>')
	If $dSuccess = 0 Then Return SetError(3, 2, False)

	; workbook.xml
	$dSuccess = FileWrite($pthWorkDir & 'xl\w.xml', '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="1" sheetId="1" r:id="rId1" /></sheets></workbook>')
	If $dSuccess = 0 Then Return SetError(3, 3, False)

	; workbook.xml.rels
	$dSuccess = DirCreate($pthWorkDir & 'xl\_rels')
	If $dSuccess = 0 Then Return SetError(2, 4, False)
	$dSuccess = FileWrite($pthWorkDir & 'xl\_rels\w.xml.rels', StringFormat( _
		'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="w/s.xml"/>%s</Relationships>', _ 
		$bDates ? '<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="st.xml"/>' : ''))
	If $dSuccess = 0 Then Return SetError(3, 4, False)

	; styles.xml
	If $bDates Then 
		$dSuccess = FileWrite($pthWorkDir & 'xl\st.xml', '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font/></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="4"><xf xfId="0"/><xf xfId="0" numFmtId="14" applyNumberFormat="1"/><xf xfId="0" numFmtId="20" applyNumberFormat="1"/><xf xfId="0" numFmtId="22" applyNumberFormat="1"/></cellXfs></styleSheet>')
		If $dSuccess = 0 Then Return SetError(3, 5, False)
	EndIf

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

	Local $mSheets = $mFiles["Worksheets"], $mSheet
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
; Last changed ..: 2023-03-13
; =================================================================================================
Func __xlsx_readCells($sFilePath, ByRef $aStrings, Const $dRowFrom = 1, $dRowTo = Default, $dColFrom = 1, $dColTo = Default)
	Local $sFileRaw = FileRead($sFilePath)
	If @error Then Return SetError(1,@error, Null)

	; determine if xml-elements have a prefix
	; if so, then change the regex-pattern
	; it's no problem to use a universal pattern but due to performance reasons the distinction is better
	Local $patRows, $patCells, $patValue
	If StringRegExp(StringLeft($sFileRaw, 1000), '<\w+:worksheet') Then
		$patRows = '(?s)<(?>\w+:)?row(?|\s+\/>|(?>\s+([^>]*))?>\s*(.+?)\s*<\/(?>\w+:)?row)'
		$patCells = '(?s)<(?>\w+:)?c\s*(?|([^>]*)\/>|([^>]*)>\s*(.*?)\s*<\/(?>\w+:)?c\b)'
		$patValue = '(?s)<(?>\w+:)?v>\s*([^<]*?)\s*<\/(?>\w+:)?v'
		;~ $patText = '(?s)<(?>\w+:)?t>\s*([^<]*?)\s*<\/(?>\w+:)?t'
		;~ $patRichText = '(?s)<(?>\w+:)?r>\s*([^<]*?)\s*<\/(?>\w+:)?r'
	Else
		$patRows = '(?s)<row(?|\s+\/>|(?>\s+([^>]*))?>\s*(.+?)\s*<\/row)'
		$patCells = '(?s)<c\s*(?|([^>]*)\/>|([^>]*)>\s*(.*?)\s*<\/c\b)'
		$patValue = '(?s)<v>\s*([^<]*?)\s*<\/v'
		;~ $patText = '(?s)<t>\s*([^<]*?)\s*<\/t'
		;~ $patRichText = '(?s)<r>\s*([^<]*?)\s*<\/r'
	EndIf

	; read rows first:
	Local $aRows = StringRegExp($sFileRaw, $patRows, 4, 3)
	If @error Then Return SetError(2, @error, Null)

	; pre-dimension the return array
	Local $aReturn[(IsInt($dRowTo) ? $dRowTo - $dRowFrom + 1 : Ubound($aRows) - $dRowFrom + 1)][(IsInt($dColTo) ? $dColTo - $dColFrom + 1 : 1)]

	; iterate over rows:
	Local $aRETmp, $aRowCol
	Local $iRow = -1, $iRowMax = 0, $iCol, $iColMax = 0
	For $aRow In $aRows

		; empty row element: <row />
		If UBound($aRow) < 2 Then
			$iRow += 1
			ContinueLoop
		EndIf

		If $aRow[1] <> "" Then ; row attributes
			; row number as xml-attribute of <row>
			$aRETmp = StringRegExp($aRow[1], '\br="(.*?)"', 3)
			If Not @error Then $iRow = Int($aRETmp[0]) - 1
		Else
			$iRow += 1
		EndIf

		; row range check (ignore rows outside the user specified range):
		If $iRow < ($dRowFrom - 1) Then ContinueLoop
		If IsInt($dRowTo) And $iRow >= $dRowTo Then ContinueLoop

		; ReDim return array (needful if "r"-attribute is used in <row>)
		If UBound($aReturn, 1) <= $iRow Then Redim $aReturn[$iRow + 1][UBound($aReturn, 2)]

		; read the cells of the row
		Local $aCells = StringRegExp($aRow[2], $patCells, 4)
		If @error Then ContinueLoop ; skip if empty

		; redim output array if number of cols is bigger
		If (IsKeyword($dColTo) = 1) And UBound($aCells) > UBound($aReturn, 2) Then Redim $aReturn[UBound($aReturn)][UBound($aCells) - $dColFrom + 1]


		; iterate over cells of this row:
		$iCol      = -1
		For $aCell in $aCells

			If UBound($aCell) < 3 Then
				$iCol += 1
				ContinueLoop
			EndIf
			Local $sType = ""

			If $aCell[1] <> "" Then ; cell attributes
				$aRETmp = StringRegExp($aCell[1], '\bt="(.*?)"', 3)
				If Not @error Then $sType = $aRETmp[0]

				; check if cell has a cell coordinate attribute
				$aRETmp = StringRegExp($aCell[1], '\br="(.*?)"', 3)
				If @error Then ; no cell coordinate attribute
					$iCol += 1
					
					; ignore columns outside the user specified range
					If IsInt($dColTo) And $iCol >= $dColTo Then	ContinueLoop 2

				Else ; it has a cell coordinate attribute
					$aRowCol = __xlsx_CellstringToRowColumn($aRETmp[0])
					$iCol = $aRowCol[0] - 1

					; ignore columns outside the user specified range
					If $iCol < ($dColFrom - 1) Then ContinueLoop
					If IsInt($dColTo) And $iCol >= $dColTo Then ContinueLoop

					; add colums to output array if cell has bigger column coordinate (possible if coordinate attribute is used)
					If ($iCol - $dColFrom + 1) >= UBound($aReturn, 2) Then Redim $aReturn[UBound($aReturn, 1)][$iCol - $dColFrom + 2]

					; read the row of the cell (normally should be the same as the row itself)
					If ($aRowCol[1] - 1) <> $iRow Then
						$iRow = $aRowCol[1] - 1
					
						; row range check (ignore rows outside the user specified range):
						If $iRow < ($dRowFrom - 1) Then ContinueLoop
						If IsInt($dRowTo) And $iRow >= $dRowTo Then ContinueLoop

						; ReDim return array (needful if "r"-attribute is used in <row>)
						If UBound($aReturn, 1) <= $iRow Then Redim $aReturn[$iRow + 1][UBound($aReturn, 2)]
					EndIf
				EndIf

			Else ; no cell attributes 
				$iCol += 1
					
				; skip rest of current row complete if column is bigger then the user wants
				If IsInt($dColTo) And $iCol >= $dColTo Then ContinueLoop 2

			EndIf

			; column range check
			If $iCol < ($dColFrom - 1) Then ContinueLoop

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

				Case Else ; normal value / defaults to "n"
					If StringRegExp($sValue, '(?i)\A(?|0x\d+|[-+]?(?>\d+)(?>\.\d+)?(?:e[-+]?\d+)?)\Z') Then $sValue = Number($sValue) ; if number then convert to number type

			EndSwitch

			; determine empty rows and columns
			If $sValue <> "" Then
				If $iRow > $iRowMax Then $iRowMax = $iRow
				If $iCol > $iColMax Then $iColMax = $iCol
			EndIf

			; add value to output array
			$aReturn[$iRow - $dRowFrom + 1][$iCol - $dColFrom + 1] = $sValue

		Next

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
	Local $sSubPath = $pthWorkDir
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
	$sPre = $oXML.documentElement.prefix
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
	$sPre = $oXML.documentElement.prefix
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
	If @error THen Return SetError(1, @error, Null)

	Local $aSI = StringRegExp($sFileRaw, '(?s)<(?>\w:)?si>(.+?)<\/(?>\w:)?si>', 3)
	If @error THen Return SetError(2, @error, Null)

	; dimension output array
	Local $aRet[UBound($aSI)]

	Local $aText
	For $i = 0 To UBound($aSI) - 1
		$aText = StringRegExp($aSI[$i], '<(?>\w:)?t[^>]*>\K[^<]*', 3)
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

; convert row number and column number into Excel´s cell coordinate syntax
; $iRow and $iCol are 1-based
Func __xlsx_createCellString($iRow, $iCol)
    If $iRow < 1 Then Return SetError(1, $iRow, 0)
    If $iCol < 1 Then Return SetError(2, $iCol, 0)

    Local $sRet = "", $iMod 

    While $iCol > 0
        $iCol -= 1   ; because it`s 1-based                          
        $iMod  = Mod($iCol, 26)     
        $sRet = Chr(65 + $iMod) & $sRet
        $iCol = Int($iCol / 26)                 
    WEnd

    Return $sRet & $iRow 
EndFunc


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
	Local $fTimeRaw, $y, $m, $d, $h, $min, $s, $tDateTime
	Switch $sType
		Case Default
			Local $aRet[6]
			$fTimeRaw = $dExcelDate - Int($dExcelDate)

			_DayValueToDate(2415018.5 + Int($dExcelDate), $aRet[0], $aRet[1], $aRet[2])

			; process the time
			$aRet[3] = Floor($fTimeRaw * 24)
			$fTimeRaw -= $aRet[3] / 24 ; = Mod($fTimeRaw, 1/24)
			$aRet[4] = Floor($fTimeRaw * 1440)
			$fTimeRaw -= $aRet[4] / 1440
			$aRet[5] = Floor($fTimeRaw * 86400)

			Return $aRet
		Case "date"
			_DayValueToDate(2415018.5 + Int($dExcelDate), $y, $m, $d)
			$tDateTime = _Date_Time_EncodeSystemTime($m, $d, $y)
			Return _WinAPI_GetDateFormat(0x0400, $tDateTime, $iFlags, $sFormat)
		Case "time"
			$fTimeRaw = $dExcelDate - Int($dExcelDate)
			; process the time
			$h = Floor($fTimeRaw * 24)
			$fTimeRaw -= $h / 24 ; = Mod($fTimeRaw, 1/24)
			$min = Floor($fTimeRaw * 1440)
			$fTimeRaw -= $min / 1440
			$s = Floor($fTimeRaw * 86400)
			Return StringFormat("%02d:%02d:%02d", $h, $min, $s)
		Case "datetime"
			$fTimeRaw = $dExcelDate - Int($dExcelDate)
			; process the time
			$h = Floor($fTimeRaw * 24)
			$fTimeRaw -= $h / 24 ; = Mod($fTimeRaw, 1/24)
			$min = Floor($fTimeRaw * 1440)
			$fTimeRaw -= $min / 1440
			$s = Floor($fTimeRaw * 86400)

			_DayValueToDate(2415018.5 + Int($dExcelDate), $y, $m, $d)
			$tDateTime = _Date_Time_EncodeSystemTime($m, $d, $y, $h, $min, $s)
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
	$iExitCode = RunWait(StringFormat('7za.exe a -tzip -mm=Deflate -mx=9 -mfb=258 -mpass=15 -mtc=off -mtm=off -mta=off "%s" "%s"', $sOutput, $sInput), "", @SW_HIDE)
	If @error Or $iExitCode <> 0 Then

		$iExitCode = RunWait(StringFormat('powershell.exe -Command "Compress-Archive ''%s'' ''%s''"', $sInput, $sOutput & ".zip"), "", @SW_HIDE)
		If Not @error And $iExitCode = 0 Then
			FileMove($sOutput & ".zip", $sOutput)
		EndIf
	EndIf
	Return 1
EndFunc   ;==>__xlsx_zip


; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __xlsx_escape4xml
; Description ...: escape special xml characters
; Syntax ........: __xlsx_escape4xml($sString)
; Parameters ....: $sString             - string with unescaped xml special chars
; Return values .: the string with escaped special chars
; Author ........: AspirinJunkie
; ===============================================================================================================================
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

; determine the last written element column for every array row 
; and the number of empty lines at the beginning 
; and the last written line in the array 
Func __xlsx_determineEmptyArrayElements(ByRef $aArray)
    If UBound($aArray, 0) <> 2 Then Return SetError(1, UBound($aArray, 0), Null)

    Local $iRows = UBound($aArray, 1), $iCols = UBound($aArray, 2), $aNumElements[$iRows]
    Local $iC, $i, $j, $iEmptyPre = 0, $iLastDataRow = 0
    For $i = 0 To $iRows -1
        $iC = 0

        ; go through current row and determine the column of the last element
        For $j = 0 To $iCols - 1
            If $aArray[$i][$j] <> "" And IsKeyword($aArray[$i][$j]) <> 2 Then $iC = $j + 1
        Next

        ; determine number of empty lines at the beginning
        If $iC = 0 And $iEmptyPre = $i Then $iEmptyPre = $i + 1
        
        ; determine the last written row number
        If $iC <> 0 Then $iLastDataRow = $i

        $aNumElements[$i] = $iC
    Next

    Local $mRet[]
    $mRet.aLastIDs = $aNumElements
    $mRet.iFirstWrittenRow = $iEmptyPre
    $mRet.iLastWrittenRow = $iLastDataRow

    Return $mRet
EndFunc

#Region general helper functions



