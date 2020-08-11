#include <array.au3>
#include <_S7_HW_Config.au3>

#cs
Beispiel für _S7_HW_Config / _S7_COM.

Erstellt in einem Projekt eine neue Hardwarekonfiguration.

Muss vorhanden sein:
Projektname: HW_Config_Test
Station: "SIMATIC 300(1)

#ce

; S7-Object
Local $oS7 = _S7_Simatic_ObjCreate()
_S7_Simatic_AutomaticSave($oS7, False) ; Automatisches Speichern ausschalten - wegen Geschwindigkeit
If @error Then Exit

; Projekt-Object
$oSt = $oS7.Projects("HW_Config_Test").Stations("SIMATIC 300(1)")

; Rack in Project einfügen
$oRack = _S7_HWConfig_Add_Rack($oSt, "rack1")

; CPU in Rack einfügen
$oCPU = _S7_HWConfig_Add_CPU($oRack, "CPU 317-2PN/DP", "6ES7 317-2EK14-0AB0")
$oSub = _S7_HWConfig_Add_SubSystem($oCPU, "Kreis 1") ; SubSystem-Verbindung mit CPU

; Konfiguration für CPU-Module - werden in Rack eingfüg
Dim $aSlave[2][5]
$aSlave[0][0] = "AO8x12Bit"
$aSlave[0][1] = "6ES7 332-5HF00-0AB0"
$aSlave[0][2] = -1
$aSlave[0][3] = 0

$aSlave[1][0] = "DI32xDC24V"
$aSlave[1][1] = "6ES7 321-1BL00-0AA0"
$aSlave[1][2] = -1
$aSlave[1][3] = 0
_S7_HWConfig_Add_CPU_Moduls($oRack, $aSlave)

; verschiedene andere Geräte - werden in SubSystem eingefügt
_S7_HWConfig_AddFestoPP($oSub, "PP2", 13, 14)
_S7_HWConfig_AddFestoPP($oSub, "PP3", 12, 12)
_S7_HWConfig_AddFestoPP($oSub, "PP4", 20, 58)
_S7_HWConfig_AddFestoPP($oSub, "PP21", 34, 92)
_S7_HWConfig_Add_DP_Koppler($oSub, "PP-Koppler", 90)

; Konfiguration für eine ET200S - wird in SubSystem eingefügt
Dim $aSlave[5][5]
$aSlave[0][0] = "PM-E DC24V"
$aSlave[0][1] = "6ES7 138-4CA01-0AA0"
$aSlave[0][2] = -1
$aSlave[0][3] = -1

$aSlave[1][0] = "4DI DC24V ST"
$aSlave[1][1] = "6ES7 131-4BD01-0AA0"
$aSlave[1][2] = 98
$aSlave[1][3] = 0

$aSlave[2][0] = "4DI DC24V ST"
$aSlave[2][1] = "6ES7 131-4BD01-0AA0"
$aSlave[2][2] = 98
$aSlave[2][3] = 4

$aSlave[3][0] = "4DI DC24V ST"
$aSlave[3][1] = "6ES7 131-4BD01-0AA0"
$aSlave[3][2] = 99
$aSlave[3][3] = 0

$aSlave[4][0] = "4DI DC24V ST"
$aSlave[4][1] = "6ES7 131-4BD01-0AA0"
$aSlave[4][2] = 99
$aSlave[4][3] = 4
$_S7_Line = 61
_S7_HWConfig_Add_ET200S($oSub, "Test_ET200S", 37, $aSlave)

; Konfiguration für eine IM153 - wird in SubSystem eingefügt
Dim $aSlave[2][5]
$aSlave[0][0] = "DI32xDC24V"
$aSlave[0][1] = "6ES7 321-1BL00-0AA0"
$aSlave[0][2] = 19
$aSlave[0][3] = 0

$aSlave[1][0] = "DI32xDC24V"
$aSlave[1][1] = "6ES7 321-1BL00-0AA0"
$aSlave[1][2] = 20
$aSlave[1][3] = 0
_S7_HWConfig_Add_IM153($oSub, "Test_IM153", 38, $aSlave)

; Projekt speichern
_S7_Simatic_AutomaticSave($oS7, True)
