This UDF provides 2 functions to read data directly from xlsx files or to output data as xlsx file.

Only the cell contents are considered - no cell formatting and the like.
It is therefore explicitly not a full replacement for the Excel UDF, since its scope goes well beyond that.
But to quickly read in data or to work with xlsx files without having Excel installed, the UDF can be quite useful. 

There may also be specially formatted xlsx files which I have not yet encountered during testing and which may cause problems. In this case it is best to make a message about it here and upload the file.

Note: xlsx files must be unpacked for reading. To make this as fast as possible it is recommended to put a [>>7za.exe<<](https://7-zip.org/a/7z2201-extra.7z) file into the script directory, otherwise a slow alternative will be used.

Otherwise an example says more than 1000 words:

```AutoIt
#include "xlsxNative.au3"

; create 2D array
Global $A[][] = [[1, 2, 3, 4, 5], ["", "6", 7, "", "8"], [], [9, "", "10", 11, True]]
_ArrayDisplay($A, "source array")

; convert the Array into a xlsx-file:
_xlsx_WriteFromArray(@ScriptDir & "\Text.xlsx", $A)

; read this xlsx-file into a 2D-Array:
$aSheet = _xlsx_2Array(@ScriptDir & "\Text.xlsx")
_ArrayDisplay($aSheet, "imported data from xlsx")
```