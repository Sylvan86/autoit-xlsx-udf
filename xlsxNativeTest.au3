#include <Array.au3>

#include "xlsxNative.au3"


;~ Global Const $sPathXlsxFile = 'H:\Administratives\Zeitenrechner.xlsm'
;~ Global Const $sPathXlsxFile = @ScriptDir & "\Abteilungsausflug.xlsx"
Global Const $sPathXlsxFile = @ScriptDir & "\Test.xlsx"

$aZellen = _xlsx_2Array($sPathXlsxFile, 2)

_ArrayDisplay($aZellen)

