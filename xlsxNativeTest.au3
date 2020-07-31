#include <Array.au3>
#include "xlsxNative.au3"

; Microsoft.XMLDom-Objekt: https://docs.microsoft.com/en-us/previous-versions/windows/desktop/ms764730(v=vs.85)
; Node-Objekt: https://docs.microsoft.com/en-us/previous-versions/windows/desktop/ms761386(v=vs.85)


; https://support.microsoft.com/de-de/help/294797/how-to-specify-namespace-when-querying-the-dom-with-xpath
; https://www.w3schools.com/xml/xml_namespaces.asp
; https://docs.microsoft.com/de-de/dotnet/standard/data/xml/namespace-support-in-the-dom
; https://support.microsoft.com/de-de/help/288147/how-to-use-xpath-to-query-against-a-user-defined-default-namespace

; Hier auch verwalten und entfernen von namespaces:
; https://docs.microsoft.com/de-de/dotnet/standard/data/xml/managing-namespaces-in-an-xml-document

; attribut "xmlns" auslesen - evtl. auch direkte Funktion hierfür

; Namespace entfernen?:
; https://docs.microsoft.com/en-us/previous-versions/windows/desktop/ms765383(v=vs.85)
;https://docs.microsoft.com/en-us/previous-versions/windows/desktop/ms756048(v=vs.85)
Local $oErrorHandler = ObjEvent("AutoIt.Error", "_ErrFunc")

;~ $sXMLPath = 'Test2.xlsx'
;~ $sXMLPath = 'Abteilungsausflug.xlsx'
;~ $sXMLPath = 'Umsätze_2020-07-22.xlsx'
$sXMLPath = 'test3.xlsx'

$aWorksheet = _xlsx_2Array($sXMLPath, 1)

_ArrayDisplay($aWorksheet)

; User's COM error function. Will be called if COM error occurs
Func _ErrFunc($oError)
    ; Do anything here.
    ConsoleWrite(@ScriptName & " (" & $oError.scriptline & ") : ==> COM Error intercepted !" & @CRLF & _
            @TAB & "err.number is: " & @TAB & @TAB & "0x" & Hex($oError.number) & @CRLF & _
            @TAB & "err.windescription:" & @TAB & $oError.windescription & @CRLF & _
            @TAB & "err.description is: " & @TAB & $oError.description & @CRLF & _
            @TAB & "err.source is: " & @TAB & @TAB & $oError.source & @CRLF & _
            @TAB & "err.helpfile is: " & @TAB & $oError.helpfile & @CRLF & _
            @TAB & "err.helpcontext is: " & @TAB & $oError.helpcontext & @CRLF & _
            @TAB & "err.lastdllerror is: " & @TAB & $oError.lastdllerror & @CRLF & _
            @TAB & "err.scriptline is: " & @TAB & $oError.scriptline & @CRLF & _
            @TAB & "err.retcode is: " & @TAB & "0x" & Hex($oError.retcode) & @CRLF & @CRLF)
EndFunc   ;==>_ErrFunc