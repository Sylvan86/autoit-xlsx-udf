#include <Array.au3>
#include "xlsxNative.au3"


$sXMLPath = 'Test.xlsx'
;~ $sXMLPath = 'Abteilungsausflug.xlsx'
;~ $sXMLPath = 'Umsätze_2020-07-22.xlsx'

$aWorksheet = _xlsx_2Array($sXMLPath, 1, 14, 15)

_ArrayDisplay($aWorksheet)





