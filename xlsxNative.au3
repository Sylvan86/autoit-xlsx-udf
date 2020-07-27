#include <Array.au3>


$iT = TimerInit()
$aZellen = _xlsx_2Array(@ScriptDir & "\Umsätze_2020-07-22.xlsx", 1)
$iT = TimerDiff($iT)

ConsoleWrite($iT & @CRLF)
_ArrayDisplay($aZellen)




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
	__unzip($sFile, $pthWorkDir)
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
	If Not IsObj($oXML) Then Return SetError(1,0,False)
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

#EndRegion xlsx specific helper functions



#Region general helper functions
Func __unzip($sInput, $sOutput, Const $SOpts = "")
	If Not FileExists($sInput) Then Return SetError(1, @error, False)
	$sOutput = StringRegExpReplace($sOutput, '(\\*)$', '')
	If Not StringInStr(FileGetAttrib($sOutput), 'D', 1) Then
		If Not DirCreate($sOutput) Then Return SetError(1, @error, False)
	EndIf

	Return __RunCmd("unzip", $SOpts & ' ' & __QiN($sInput) & " -d " & __QiN($sOutput))
EndFunc   ;==>__unzip

; adds quotes if white spaces and escapes tabs and vertical space
Func __QiN($s_String, Const $s_Quote = '"', Const $s_QEscape = '""', Const $s_TabEscape = Default, Const $s_LineEscape = Default)
	; by AspirinJunkie

	; escape line breaks
	If $s_LineEscape <> Default Then $s_String = StringRegExpReplace($s_String, '(\v)+', $s_LineEscape)

	; escape tabs
	If $s_TabEscape <> Default Then $s_String = StringRegExpReplace($s_String, '(\t)+', $s_TabEscape)

	; quotes around if necessary
	If StringInStr($s_String, ' ', 1) Then Return $s_Quote & StringReplace($s_String, $s_Quote, $s_QEscape, 0, 1) & $s_Quote

	Return $s_String
EndFunc   ;==>__QiN

; #FUNCTION# ======================================================================================
; Name ..........: __RunCmd()
; Description ...: runs commandline programs or cmd-command and return their output
; Syntax ........: __RunCmd($s_Cmd, [$sParameter = '', [$b_CmdSpec = False, [$WorkDir = @WorkingDir]]])
; Parameters ....: $s_Cmd        - the command which should be executed (can be full command or without parameters)
;                  $sParameter   - additional parameters for the command
;                  $b_CmdSpec    - If true the command is interpreted as a command for cmd.exe
;                  $WorkDir      - the working directory
; Return values .: Success: returns a string with the output
;                  Failure: set @error and returns a debug-string
; Author ........: AspirinJunkie
; =================================================================================================
Func __RunCmd($s_Cmd, $sParameter = '', $b_CmdSpec = False, $WorkDir = @WorkingDir)
	Local Static $h_User32DLL = DllOpen('user32.dll')
	If @error Then Return SetError(1, @error, "")

	If $b_CmdSpec Then $s_Cmd = @ComSpec & " /c " & $s_Cmd

	Local $s_Ret, $s_Line, $s_Err

	If $sParameter <> '' Then $sParameter = ' ' & $sParameter
	Local $iPID = Run($s_Cmd & $sParameter, $WorkDir, @SW_HIDE, 0x2 + 0x4)
	If @error Then Return SetError(2, @error, "")

	Do
		Sleep(10)
		$s_Line = StdoutRead($iPID)
		If @extended > 0 Then $s_Ret &= $s_Line
	Until @error

	Do
		Sleep(10)
		$s_Line = StderrRead($iPID)
		If @extended > 0 Then $s_Err &= $s_Line
	Until @error

	$s_Ret = DllCall($h_User32DLL, 'BOOL', 'OemToChar', 'str', $s_Ret, 'str', '')[2]

	If $s_Err <> "" Then
		If $s_Ret <> "" Then Return SetError(3, 0, "------- StdOut -----------" & @CRLF & $s_Ret & @CRLF & @CRLF & "------- StdErr -----------" & @CRLF & $s_Err)
		Return SetError(3, 0, $s_Err)
	EndIf
	Return $s_Ret
EndFunc   ;==>__RunCmd

; returns single from a xml-dom-object and handles errors
Func __xmlSingleText(ByRef $oXML, Const $sXPath)
	Local $oTmp = $oXML.selectSingleNode($sXPath)
	Return IsObj($oTmp) ? $oTmp.text : SetError(1, 0, "")
EndFunc   ;==>__xmlSingleText

#Region general helper functions



