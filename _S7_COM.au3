#include <array.au3>
#include <excel.au3>
#Region Description
; ==============================================================================
; UDF ...........: FF.au3
Global Const $_S7_AU3VERSION = "0.5.9"
; Description ...: An UDF for Simatic STEP autmation.
; Requirement ...: Simatic Step7 > V5.3 SP3 / MS-Excel (for the function: _S7_SymbolTable_ExportToExcel)
; Author(s) .....: Thorsten Willert
; Date ..........: 10.12.2019
; AutoIt Version : v3.3.8.1
; ==============================================================================
#cs
	Simatic
	Simatic.Simatic.1
	OBJ 	_S7_Simatic_ObjCreate / Simatic = Simatic.Simatic.1
	BOOL 	_S7_Simatic_AutomaticSave / Simatic.AutomaticSave (Read / Write)
	BSTR 	_S7_Simatic_VerbLogFile (read write)
	_S7_Simatic_SetPGInterface / Simatic.SetPGInterface (Opens Dialog)
	_S7_Simatic_UnattendedServerMode !!!
	LONG 	_S7_Simatic_MsgAssignmentType / Simatic.MsgAssignmentType (Read / Write)
	BOOL 	_S7_Simatic_IsSilentMode (read)
	_S7_Simatic_Save / Simatic.Save (void)
	HRESULT _S7_Simatic_GetSTEP7Language / Simatic.GetSTEP7Language (Read) ??? funktioniert nicht mit AutoIt

	Projects
	Simatic.Projects
	OBJ 	_S7_Projects_GetProject
	BOOL 	_S7_Projects_Exists
	Array 	_S7_Projects_GetList
	Int 	_S7_Projects_Count
	BOOL 	_S7_Projects_Add

	Project
	Simatic.Projects.Project
	Array 	_S7_Project_GetInfo
	_S7_Project_Name
	String	_S7_Project_Creator (Read / Write)
	String	_S7_Project_Comment (Read / Write)
	Bool 	_S7_Project_Remove

	Stations
	Simatic.Projects.Project.Stations
	OBJ 	_S7_Stations_GetStation
	Bool	_S7_Stations_Exists
	Array	_S7_Stations_GetList
	Int	_S7_Stations_Count
	Bool	_S7_Stations_Import
	Bool	_S7_Stations_Add
	Bool	_S7_Stations_Remove

	Programs
	Simatic.Projects.Project.Programs
	Array 	_S7_Programs_GetList

	Blocks
	Simatic.Projects.Project.Programs.Next("Blocks")
	OBJ	_S7_Blocks_GetBlock

	Source Files
	Simatic.Projects.Project.Programs.Next("Source Files")
	OBJ	_S7_SourceFiles_GetSource
	Bool	_S7_SourceFiles_Export
	Bool	_S7_SourceFiles_Add
	Bool	_S7_SourceFiles_Compile
	Array	_S7_SourceFiles_GetInfo

#ce
#Region Description

Global $_S7_MultiUser = True ; if True you can't disable AutomaticSave with _S7_Simatic_AutomaticSave() !!!

Global Enum _
		$_S7_Success, _
		$_S7_GeneralError, _
		$_S7_ComError, _
		$_S7_InvalidDataType, _
		$_S7_InvalidObjectType, _
		$_S7_InvalidValue

Global Const $_S7_Any = -1
Global Const $_S7Project = 1122305, $_S7Library = 1122306; S7ProjectType
Global Const $_S7 = 1327361, $_M7 =1327367 ; S7ProgramType
Global Enum $_S7Run, $_S7Stop, $_S7Halt, $_S7Defect, $_S7Startup ; S7ModState
Global Enum $_S7SymImportInsert, $_S7SymImportOverwriteNameLeading, $_S7SymImportOverwriteOperandLeading ; S7SymImportFlags
Global Enum $_S7300Station = 30, $_S7400Station, $_S7HStation, $_S7PCStation ; S7StationType
Global Enum $_S7Block, $_S7Container = 63, $_S7Source, $_S7Plan ; S7SWObjType ???

Global Enum  $_S7FB = 1138945, $_S7FC, $_S7DB, $_S7OB, $_S7SDB, $_S7SDBs = 0, $_S7UDT = 1138950, $_S7SFC, $_S7SFB, $_S7VAT ; S7BlockType
	#cs
	  FB = 1138945
	  FC = 1138946

	  DB = 1138947
	  OB = 1138948
	  DB = 1138949

	  UDT = 1138950
	  SFC = 1138951
	  SFB = 1138952

	  VAT = 1138953
   #ce

Global Const $_S7BlockContainer = 1138689, $_S7SourceContainer = 1122308, $_S7PlanContainer = -1; S7ContainerType ???
Global Enum $_S7AWL, $_S7SCL, $_S7GR7, $_S7SCLMake, $_S7GG, $_S7ZG, $_S7NET ; S7SourceType;
;Global Enum $_S7SymImportInsert, $_S7SymImportOverwriteNameLeading, $_S7SymImportOverwriteOperandLeading ; S7SymImportFlags;

; V5.1 SP2
Global Enum $_MPI, $_PROFIBUS, $_INDUSTRIAL_ETHERNET, $_PTP ; S7SubnetType
Global Enum $_NO_DATA, $_DLG_SYMB, $_SUBNET_DATA, $_CONN_DATA, $_SUBNET_CONN_DATA ; NetDataFlags
;Global Enum $_INVALID, $_S7_CONNECTION, $_S7_CONNECTION_REDUNDANT, $_POINTTOPOINT ; S7ConnType

; V5.3 SP3
Global Enum $_INVALID, $_S7_CONNECTION, $_S7_CONNECTION_REDUNDANT, $_POINTTOPOINT, $_ISO, $_UDP ; S7ConnType;

;==============================================================================
Global $o_S7_ErrorHandler = ObjEvent("AutoIt.Error", "__S7_ErrFunc")
Global $_S7_Function ; Name of the current function / error handling
;==============================================================================

#region Simatic
; #FUNCTION# ===================================================================
; Name ..........: _S7_Simatic_ObjCreate
; Description ...: Creates a simatic object
; AutoIt Version : V3.3.8.1
; Requirement(s).: Step7 V5.5 SP3
; Syntax ........: _S7_Simatic_ObjCreate([$sVerbLogFile = ""[, $bUnattendedServerMode = False]])
; Parameter(s): .: $sVerbLogFile - Optional: (Default = "") :
;                  $bUnattendedServerMode - Optional: (Default = False) :
; Return Value ..: Success      -
;                  Failure      -
;                  @ERROR       -
;                  @EXTENDED    -
; Author(s) .....: Thorsten Willert
; Date ..........: Sun Jan 12 03:24:12 CET 2014
; Version .......: 1.1
; ==============================================================================
Func _S7_Simatic_ObjCreate($sVerbLogFile = "", $bUnattendedServerMode = False)
	; Simatic.Simatic.1
	$_S7_Function = "_S7_Simatic_ObjCreate"

	Local $oS7 = ObjCreate("Simatic.Simatic.1")
	If Not IsObj($oS7) Then Return SetError($_S7_GeneralError, 1, "")

	If @error = $_S7_ComError Then
		Return SetError($_S7_ComError, 2, "")
	Else
		_S7_Simatic_VerbLogFile($oS7, $sVerbLogFile)
		If @error Then Return SetError(@error, 3, "")

		_S7_Simatic_UnattendedServerMode($oS7, $bUnattendedServerMode)
		If @error Then Return SetError(@error, 4, "")

		Return SetError($_S7_Success, 0, $oS7)
	EndIf
EndFunc   ;==>_S7_Simatic_ObjCreate

; #FUNCTION# ===================================================================
; Name ..........: _S7_Simatic_AutomaticSave
; Description ...:
; AutoIt Version : V3.3.8.1
; Requirement(s).: Step7 V5.5 SP3
; Syntax ........: _S7_Simatic_AutomaticSave(ByRef $oS7[, $bSave = -1])
; Parameter(s): .: $oS7         -
;                  $bSave       - Optional: (Default = -1) :
; Return Value ..: Success      - True Or Current value
;                  Failure      - False
;                  @ERROR       -
;                  @EXTENDED    -
; Author(s) .....: Thorsten Willert
; Date ..........: Sun Jan 12 03:25:29 CET 2014
; Version .......: 1.2
; ==============================================================================
#cs
	Hinweis
	Die Kommando-Schnittstelle speichert auch bei AutomaticSave = False in folgenden Situationen:
	·	Beim Beenden des Programms, das die Kommando-Schnittstelle verwendet
	·	Beim Erzeugen oder Löschen von Projekten
	·	Beim Erzeugen (auch Importieren) von Stationen
	·	Vor Aufruf der Methoden Edit oder Compile.
	Wenn mehrere Anwendungen (STEP 7-Anwendungen oder über Kommando-Schnittstelle) gleichzeitig auf
	demselben Projekt arbeiten, darf AutomaticSave nicht ausgeschaltet werden.
#ce
Func _S7_Simatic_AutomaticSave(ByRef $oS7, $bSave = -1)
	$_S7_Function = "_S7_Simatic_AutomaticSave"
	;Simatic.AutomaticSave
	If $_S7_MultiUser = True Then Return ; !!!!

	If Not IsObj($oS7) Then Return SetError($_S7_InvalidDataType)
	Switch $bSave
		Case 1
			$oS7.AutomaticSave = True
		Case 0
			$oS7.AutomaticSave = False
		Case -1
			Return $oS7.AutomaticSave
	EndSwitch
	Return SetError(0)
EndFunc   ;==>_S7_Simatic_AutomaticSave
; ==============================================================================
Func _S7_Simatic_GetSTEP7Language(ByRef $oS7)
	$_S7_Function = "_S7_Simatic_GetSTEP7Language"
	If Not IsObj($oS7) Then Return SetError($_S7_InvalidDataType, 1, "")
	Local $tmp
	$oS7.GetSTEP7Language($tmp)
	Return SetError($_S7_Success, 0, $tmp) ; ???
EndFunc   ;==>_S7_Simatic_GetSTEP7Language
;==============================================================================
Func _S7_Simatic_Save(ByRef $oS7)
	;Simatic.Save
	$_S7_Function = "_S7_Simatic_Save"
	If Not IsObj($oS7) Then Return SetError($_S7_InvalidDataType, 1, 0)
	$oS7.Save
	Return SetError($_S7_Success, 0, 1)
EndFunc   ;==>_S7_Simatic_Save
;==============================================================================
Func _S7_Simatic_SetPGInterface(ByRef $oS7)
	; Simatic.SetPGInterface
	$_S7_Function = "_S7_Simatic_SetPGInterface"
	If Not IsObj($oS7) Then Return SetError($_S7_InvalidDataType, 1, 0)
	Local $h = $oS7.SetPGInterface
	If IsHWnd($h) Then
		Return SetError($_S7_Success, 0, 1)
	Else
		Return SetError($_S7_GeneralError, 0, 0)
	EndIf
EndFunc   ;==>_S7_Simatic_SetPGInterface
;==============================================================================
Func _S7_Simatic_MsgAssignmentType(ByRef $oS7, $iMode)
	;Simatic.MsgAssignmentType
	$_S7_Function = "_S7_Simatic_MsgAssignmentType"
	If Not IsObj($oS7) Then Return SetError($_S7_InvalidDataType, 1, 0)
	$iMode = Int($iMode)
	Switch $iMode
		Case 0, 1, 2 ; Dialog immer anzeigen ; CPU-weite Einstellung ; Projektweite Einstellung
			$oS7.MsgAssignmentType = $iMode
			Return SetError($_S7_Success, 0, 1)
		Case Else
			Return SetError($_S7_InvalidValue, 0, 0)
	EndSwitch
EndFunc   ;==>_S7_Simatic_MsgAssignmentType
;==============================================================================
#cs
	UnattendedServerMode

	Wenn die Eigenschaft gesetzt ist (=TRUE), dann werden alle Meldungen unterdrückt,
	also auch die, für die Sie keine automatische Antwort in der Registry hinterlegt haben.
	In diesem Fall wird die Meldung automatisch mit der vorselektierten Schaltfläche, z. B. "Nein", quittiert.
	Die Folge ist, dass die Meldung falsch quittiert werden kann. Sie dürfen die Eigenschaft
	nur dann auf "TRUE" setzen, wenn das Programm auf keinen Fall "hängen" bleiben darf und
	dafür eine falsche Quittierung in Kauf genommen werden kann.
#ce
Func _S7_Simatic_UnattendedServerMode(ByRef $oS7, $bMode = False)
	$_S7_Function = "_S7_Simatic_UnattendedServerMode"
	If Not IsObj($oS7) Then Return SetError($_S7_InvalidDataType, 1, 0)
	If $bMode Then
		$oS7.UnattendedServerMode = True
	Else
		$oS7.UnattendedServerMode = False
	EndIf
	Return SetError($_S7_Success, 0, 1)
EndFunc   ;==>_S7_Simatic_UnattendedServerMode
;==============================================================================
#cs
	VerbLogFile

	Pfadname einer Protokolldatei (Log-Datei). In dieser Datei werden die während der Compile-Aufrufe
	anfallenden Meldungen (die z. B. im Fehlerfenster des Editors angezeigt würden) protokolliert.
	Wenn für VerbLogFile ein String (d. h. ein gültiger Pfad) eingesetzt ist, wird implizit der Silent Mode
	angefordert. Im Silent-Mode erscheint bei Operationen wie "Compile" kein sichtbares, vom Anwender
	manuell zu schließendes Applikationsfenster mehr.
	Durch Setzen auf den Leerstring wird der Silent Mode wieder deaktiviert.

	Wenn die Dienst-erbringende Komponente diesen Mechanismus nicht unterstützt, wird wie gewohnt das
	Fehlerfenster des Editors aufgerufen.
#ce
Func _S7_Simatic_VerbLogFile(ByRef $oS7, $sFile = "")
	$_S7_Function = "_S7_Simatic_VerbLogFile"
	If Not IsObj($oS7) Then Return SetError($_S7_InvalidDataType, 1, 0)
	If ($sFile <> "") And (Not FileExists($sFile)) Then
		Local $h = FileOpen($sFile, 9)
		If @error Then
			Return SetError($_S7_InvalidValue, 2, 0)
		Else
			FileWrite($h, "")
			FileClose($h)
			$oS7.VerbLogFile($sFile)
			Return SetError($_S7_Success, 0, 1)
		EndIf
	Else
		$oS7.VerbLogFile = ""
		Return SetError(0, 0, "")
	EndIf

	Return SetError($_S7_GeneralError, 0, 0)
EndFunc   ;==>_S7_Simatic_VerbLogFile
;==============================================================================
Func _S7_Simatic_IsSilentMode(ByRef $oS7)
	$_S7_Function = "_S7_Simatic_IsSilentMode"
	If Not IsObj($oS7) Then Return SetError($_S7_InvalidDataType, 1, False)
	Local $sFile = $oS7.VerbLogFile
	If $sFile = "" Then Return SetError($_S7_Success, 0, False)
	If FileExists($sFile) Then Return SetError($_S7_Success, 0, True)
EndFunc   ;==>_S7_Simatic_IsSilentMode
;==============================================================================
#endregion Simatic

#region S7Projects

;==============================================================================
Func _S7_Projects_GetProject(ByRef $oS7, $vProject)
	$_S7_Function = "_S7_Projects_GetProject"
	Local $o = $oS7.Projects
	Local $ret = __S7_GetObject($o, $vProject)
	Return SetError(@error, @extended, $ret)
EndFunc   ;==>_S7_Projects_GetProject
;==============================================================================
Func _S7_Projects_Exists(ByRef $oS7, $sProjectName)
	$_S7_Function = "_S7_Projects_Exists"
	Local $o = $oS7.Projects
	Local $ret = __S7_ObjExists($o, $sProjectName)
	Return SetError(@error, @extended, $ret)
EndFunc   ;==>_S7_Projects_Exists
;==============================================================================
Func _S7_Projects_GetList(ByRef $oS7, $iS7ProjectType = Default)
	$_S7_Function = "_S7_Projects_GetList"
	Local $o = $os7.Projects
	Local $ret = __S7_GetList($o, $iS7ProjectType, $_S7_Any)
	Return SetError(@error, @extended, $ret)
EndFunc   ;==>_S7_Projects_GetList
;==============================================================================
Func _S7_Projects_Count(ByRef $oS7)
	$_S7_Function = "_S7_Projects_Count"
	If Not IsObj($oS7) Then Return SetError($_S7_InvalidDataType, 1, 0)
	Return SetError($_S7_Success, 0, $oS7.Projects.Count)
EndFunc   ;==>_S7_Projects_Count
;==============================================================================
Func _S7_Projects_Add(ByRef $oS7, $sName, $sProjectRootDir = "", $S7ProjectType = Default)
	$_S7_Function = "_S7_Projects_Add"
	If Not IsObj($oS7) Then Return SetError($_S7_InvalidDataType, 1, 0)
	If $sName = "" Then Return SetError($_S7_InvalidValue, 2, 0)
	If ($sProjectRootDir <> "") And (Not FileExists($sProjectRootDir)) Then Return SetError($_S7_InvalidValue, 3, 0)
	If $S7ProjectType = Default Then $S7ProjectType = $_S7Project

	$oS7.Projects.Add($sName, $sProjectRootDir, $_S7Project)
	If @error = $_S7_ComError Then
		Return SetError($_S7_ComError, 0, 0)
	Else
		Return SetError($_S7_Success, 0, 1)
	EndIf
EndFunc   ;==>_S7_Projects_Add
;==============================================================================
#endregion S7Projects

#region Project
;==============================================================================
Func _S7_Project_GetInfo(ByRef $oProject)
	$_S7_Function = "_S7_Project_GetInfo"
	Local $aRet[1][8]
	If Not IsObj($oProject) Then Return SetError($_S7_InvalidDataType, 1, $aRet)

   With $oProject
		_DataFill($aRet, 1, _
			   .Name, _ 	; 0
			   .LogPath, _ 	; 1
			   .Creator, _ 	; 2
			   .Comment, _ 	; 3
			   .Created, _ 	; 4
			   .Modified, _ 	; 5
			   .Version, _ 	; 6
			   .Type) ; 7
		If @error Then Return SetError($_S7_GeneralError, @extended, $aRet)
   EndWith

	Return SetError($_S7_Success, 0, $aRet)
EndFunc   ;==>_S7_Project_GetInfo
;==============================================================================
Func _S7_Project_Name(ByRef $oProject, $sName)
	$_S7_Function = "_S7_Project_Name"
	If Not IsObj($oProject) Then Return SetError($_S7_InvalidDataType, 1, 0)

	$oProject.Name = String($sName)

	If @error = $_S7_ComError Then
		Return SetError($_S7_ComError, 0, 0)
	Else
		Return SetError($_S7_Success, 0, 1)
	EndIf
EndFunc   ;==>_S7_Project_Name
;==============================================================================
Func _S7_Project_Creator(ByRef $oProject, $sCreator = "")
	$_S7_Function = "_S7_Project_Creator"
	If Not IsObj($oProject) Then Return SetError($_S7_InvalidDataType, 1, 0)

	Switch $sCreator
		Case ""
			Return SetError($_S7_Success, 0, $oProject.Creator)
		Case Else
			Return SetError($_S7_Success, 0, $oProject.Creator = String($sCreator))
	EndSwitch

	Return SetError($_S7_GeneralError, 0, "")
EndFunc   ;==>_S7_Project_Creator
;==============================================================================
Func _S7_Project_Comment(ByRef $oProject, $sComment = "")
	$_S7_Function = "_S7_Project_Comment"
	If Not IsObj($oProject) Then Return SetError($_S7_InvalidDataType)

	Switch $sComment
		Case ""
			Return SetError($_S7_Success, 0, $oProject.Comment)
		Case Else
			Return SetError($_S7_Success, 0, $oProject.Comment = String($sComment))
	EndSwitch

	Return SetError($_S7_GeneralError, 0, "")
EndFunc   ;==>_S7_Project_Comment
;==============================================================================
Func _S7_Project_Remove(ByRef $oProject)
	$_S7_Function = "_S7_Project_Remove"
	If Not IsObj($oProject) Then Return SetError($_S7_InvalidDataType, 1, 0)

	$oProject.Remove
	If @error = $_S7_ComError Then
		Return SetError($_S7_ComError, 0, 0)
	Else
		Return SetError($_S7_Success, 0, 1)
	EndIf
EndFunc   ;==>_S7_Project_Remove
;==============================================================================
#endregion Project

#region Stations
;==============================================================================
Func _S7_Stations_GetStation(ByRef $oProject, $vStation)
	$_S7_Function = "_S7_Stations_GetStation"
	; Simatic.Projects.Project.Stations
	$o = $oProject.Stations
	Local $ret = __S7_GetObject($oProject, $vStation)
	Return SetError(@error, @extended, $ret)
EndFunc   ;==>_S7_Stations_GetStation
;==============================================================================
Func _S7_Stations_Exists(ByRef $oStation, $sStationName)
	$_S7_Function = "_S7_Stations_Exists"
	; Simatic.Projects.Project.Stations
	Local $ret = __S7_ObjExists($oStation, $sStationName)
	Return SetError(@error, @extended, $ret)
EndFunc   ;==>_S7_Stations_Exists
;==============================================================================
Func _S7_Stations_GetList(ByRef $oProject, $S7StationType = Default)
	$_S7_Function = "_S7_Stations_GetList"
	; $_S7300Station, $_S7400Station, $_S7HStation, $_S7PCStation ; S7StationType
	; Simatic.Projects.Project.Stations
	Local $o = $oProject.Stations
	Local $ret = __S7_GetList($o, $S7StationType, $_S7_Any)
	Return SetError(@error, @extended, $ret)
EndFunc   ;==>_S7_Stations_GetList
;==============================================================================
Func _S7_Stations_Count(ByRef $oStations)
	$_S7_Function = "_S7_Stations_Count"
	; Simatic.Projects.Project.Stations
	If Not IsObj($oStations) Then Return SetError($_S7_InvalidDataType, 1, 0)
	Return SetError($_S7_Success, 0, $oStations.Count)
EndFunc   ;==>_S7_Stations_Count
;==============================================================================
Func _S7_Stations_Import(ByRef $oStations, $sFile)
	$_S7_Function = "_S7_Stations_Import"
	; Simatic.Projects.Project.Stations
	If Not IsObj($oStations) Then Return SetError($_S7_InvalidDataType, 1, "")
	If Not FileExists($sFile) Then Return SetError($_S7_InvalidValue, 2, "")

	Local $oStation = $oStations.Import($sFile)
	If @error = $_S7_ComError Then
		Return SetError($_S7_ComError, 0, "")
	Else
		Return SetError($_S7_Success, 0, $oStation)
	EndIf
EndFunc   ;==>_S7_Stations_Import
;==============================================================================
; $_S7300Station, $_S7400Station, $_S7HStation, $_S7PCStation ; S7StationType
Func _S7_Stations_Add(ByRef $oStations, $sStationName, $S7StationType)
	$_S7_Function = "_S7_Stations_Add"
	; Simatic.Projects.Project.Stations
	If Not IsObj($oStations) Then Return SetError($_S7_InvalidDataType, 1, 0)
	If $sStationName = "" Then Return SetError($_S7_InvalidValue, 2, 0)
	If _S7_Stations_Exists($oStations, $sStationName) Then Return SetError($_S7_InvalidValue, 2, 0)
	Switch $S7StationType
		Case $_S7300Station, $_S7400Station, $_S7HStation, $_S7PCStation
		Case Else
			Return SetError($_S7_InvalidValue, 3, 0)
	EndSwitch

	$oStations.Add($sStationName, $S7StationType)
	If @error = $_S7_ComError Then
		Return SetError($_S7_ComError, 0, 0)
	Else
		Return SetError($_S7_Success, 0, 1)
	EndIf
EndFunc   ;==>_S7_Stations_Add
;==============================================================================
Func _S7_Stations_Remove(ByRef $oProject, $vStation)
	$_S7_Function = "_S7_Stations_Remove"
	; Simatic.Projects.Project.Stations
	If Not IsObj($oProject) Then Return SetError($_S7_InvalidDataType, 1, 0)
	Switch VarGetType($vStation)
		Case "Int32"
			If $vStation < 1 Or $vStation > $oProject.Stations.Count Then Return SetError($_S7_InvalidValue, 2, 0)
		Case "String"
			If Not _S7_Stations_Exists($oProject, $vStation) Then Return SetError($_S7_InvalidValue, 2, 0)
		Case Else
			Return SetError($_S7_InvalidDataType, 2, 0)
	EndSwitch

	$oProject.Stations.Remove($vStation)
	If @error = $_S7_ComError Then
		Return SetError($_S7_ComError, 0, 0)
	Else
		Return SetError($_S7_Success, 0, 1)
	EndIf
EndFunc   ;==>_S7_Stations_Remove
;==============================================================================
#endregion Stations

#region S7Programs
;==============================================================================
;==============================================================================
Func _S7_Programs_GetList(ByRef $oS7Project, $iS7ProgramType = Default) ;  $_S7, $_M7 ; S7ProgramType
	$_S7_Function = "_S7_Programs_GetList"
	; Simatic.Projects.Project.Programs
	Local $o = $oS7Project.Programs
	Local $ret = __S7_GetList($o, $iS7ProgramType, $_S7_Any)
	Return SetError(@error, @extended, $ret)
EndFunc   ;==>_S7_Programs_GetList
;==============================================================================
 Func _S7_Programs_GenerateSource(ByRef $oProgram, $sFile, $bClipPut = False)
   $_S7_Function = "_S7_Programs_GenerateSource"
   If Not IsObj($oProgram) Then Return SetError($_S7_InvalidDataType, 1, 0)
   If $sFile = "" Then Return SetError($_S7_InvalidValue, 2, 0)

   $oProgram.GenerateSource($sFile)
   If @error = $_S7_ComError Then
	  Return SetError($_S7_ComError, 0, 0)
   Else
	  If $bClipPut Then ClipPut( FileRead($sFile) )
	  Return SetError($_S7_Success, 0, 1)
   EndIf
EndFunc   ;==>_S7_Programs_GenerateSource
;==============================================================================
#endregion S7Programs

#region S7SymbolTable
Func _S7_SymbolTable_Export($oProgram, $sType = "asc")
   $_S7_Function = "_S7_SymbolTable_Export"
   If Not IsObj($oProgram) Then Return SetError($_S7_InvalidDataType, 1, 0)
   Switch $sType
	   Case "asc"
   EndSwitch
EndFunc
;==============================================================================
Func _S7_SymbolTable_ExportToExcel($oProgram, $sExcelFile, $vTab, $iCol, $iRow)
	$_S7_Function = "_S7_SymbolTable_ExportToExcel"
	Return
EndFunc

#endregion S7SymbolTable

#region "Blocks"
;==============================================================================
Func _S7_Blocks_GetBlock(ByRef $oProgram, $vBlock)
	$_S7_Function = "_S7_Blocks_GetBlock"
	If Not IsObj($oProgram) Then Return SetError($_S7_InvalidDataType, 1, 0)

	Local $oBlock = $oProgram.Next("Blocks").Next($vBlock)
	If @error = $_S7_ComError Then
		Return SetError($_S7_ComError, 0, "")
	Else
		Return SetError($_S7_Success, 0, $oBlock)
	EndIf
EndFunc   ;==>_S7_Blocks_GetBlock
;===============================================================================
Func _S7_Blocks_GetInfo(ByRef $oBlock)
	$_S7_Function = "_S7_Blocks_GetInfo"
	Local $aRet[1][14]
	If Not IsObj($oBlock) Then Return SetError($_S7_InvalidDataType, 1, $aRet)


   ;With $oBlock
		_DataFill($aRet, 1, _
				$oBlock.Name, _ 	; 0
				$oBlock.LogPath, _ 	; 1
				$oBlock.Creator, _ 	; 2
				$oBlock.Comment, _ 	; 3
				$oBlock.Created, _ 	; 4
				$oBlock.Modified, _ 	; 5
				$oBlock.Version, _ 	; 6
			    $oBlock.Size, _	; 7
				$oBlock.SymbolicName, _ ; 8
				$oBlock.Language, _ ; 9
				$oBlock.HeaderName, _ ; 10
				$oBlock.HeaderVersion, _; 11
				$oBlock.Type, _	; 12
				$oBlock.ConcreteType) ; 13


		If @error Then Return SetError($_S7_GeneralError, @extended, $aRet)
   ;EndWith

	Return SetError($_S7_Success, 0, $aRet)
EndFunc   ;==>_S7_Blocks_GetInfo
;==============================================================================
#endregion "Blocks"

#region "Source Files"
;==============================================================================
Func _S7_SourceFiles_GetSource(ByRef $oProgram, $vSrc)
	$_S7_Function = "_S7_SourceFiles_GetSource"
	If Not IsObj($oProgram) Then Return SetError($_S7_InvalidDataType, 1, 0)

	Local $oSrc = $oProgram.Next("Source Files").Next($vSrc)
	If @error = $_S7_ComError Then
		Return SetError($_S7_ComError, 0, "")
	Else
		Return SetError($_S7_Success, 0, $oSrc)
	EndIf
EndFunc   ;==>_S7_SourceFiles_GetSource
;==============================================================================
Func _S7_SourceFiles_Export(ByRef $oSource, $sFile)
	$_S7_Function = "_S7_SourceFiles_Export"
	If Not IsObj($oSource) Then Return SetError($_S7_InvalidDataType, 1, 0)
	If $sFile = "" Then Return SetError($_S7_InvalidValue, 2, 0)

	$oSource.Export($sFile)
	If @error = $_S7_ComError Then
		Return SetError($_S7_ComError, 0, 0)
	Else
		Return SetError($_S7_Success, 0, 1)
	EndIf
EndFunc   ;==>_S7_SourceFiles_Export
;==============================================================================
Func _S7_SourceFiles_Add(ByRef $oProgram, $sSourceName, $S7SWObjType = $_S7Source, $sFile = "")
	$_S7_Function = "_S7_SourceFiles_Add"
	If Not IsObj($oProgram) Then Return SetError($_S7_InvalidDataType, 1, 0)
	If $sSourceName = "" Then Return SetError($_S7_InvalidValue, 2, 0)
	If Not IsInt($S7SWObjType) Then Return SetError($_S7_InvalidValue, 3, 0)
	If $sFile = "" Or Not FileExists($sFile) Then Return SetError($_S7_InvalidValue, 4, 0)

	$oProgram.Next("Source Files").Next.Add($sSourceName, $S7SWObjType, $sFile)
	If @error = $_S7_ComError Then
		Return SetError($_S7_ComError, 0, 0)
	Else
		Return SetError($_S7_Success, 0, 1)
	EndIf
EndFunc   ;==>_S7_SourceFiles_Add
;==============================================================================
Func _S7_SourceFiles_Compile(ByRef $oSource)
	$_S7_Function = "_S7_SourceFiles_Compile"
	If Not IsObj($oSource) Then Return SetError($_S7_InvalidDataType, 1, 0)

	$oSource.Compile()
	If @error = $_S7_ComError Then
		Return SetError($_S7_ComError, 0, 0)
	Else
		Return SetError($_S7_Success, 0, 1)
	EndIf
EndFunc   ;==>_S7_SourceFiles_Compile
;==============================================================================
Func _S7_SourceFiles_GetInfo(ByRef $oSource)
	$_S7_Function = "_S7_SourceFiles_GetInfo"
	Local $aRet[1][10]
	If Not IsObj($oSource) Then Return SetError($_S7_InvalidDataType, 1, $aRet)

	;With $oSource
		_DataFill($aRet, 1, _
				$oSource.Name, _ 	; 0
				$oSource.LogPath, _ 	; 1
				$oSource.Creator, _ 	; 2
				$oSource.Comment, _ 	; 3
				$oSource.Created, _ 	; 4
				$oSource.Modified, _ 	; 5
				$oSource.Version, _ 	; 6
				$oSource.Size, _	; 7
				$oSource.Type, _	; 8
				$oSource.ConcreteType) ; 9
		If @error Then Return SetError($_S7_GeneralError, @extended, $aRet)
	;EndWith

	Return SetError($_S7_Success, 0, $aRet)
EndFunc   ;==>_S7_SourceFiles_GetInfo
#endregion "Source Files"

;==============================================================================
Func _DataFill(ByRef $a, $iMode, $0 = "", $1 = "", $2 = "", $3 = "", $4 = "", $5 = "", $6 = "", $7 = "", $8 = "", $9 = "", $10 = "", $11 = "", $12 = "", $13 = "", $14 = "", $15 = "", $16 = "", $17 = "", $18 = "", $19 = "")
	Local $s = AutoItSetOption("GUIDataSeparatorChar")
	Switch $iMode
		Case 1
			Local $iE = UBound($a, 2) - 1
			;If $iE-2 <= @NumParams Then Return SetError(1, @NumParams, False)
			For $i = 0 To $iE
				$a[0][$i] = Eval($i)
			Next
		Case 2
			For $i = 0 To @NumParams - 1
				$a &= Eval($i) & $s
			Next
			$a &= Eval(@NumParams)
	EndSwitch

	Return SetError(0, 0, True)
EndFunc   ;==>_DataFill
;==============================================================================
Func __S7_ObjExists(ByRef $oObj, ByRef $sName)
	If Not IsObj($oObj) Then Return SetError($_S7_InvalidDataType, 1, 0)
	If $sName = "" Then Return SetError($_S7_InvalidValue, 2, 0)

	For $o In $oObj
		If $o.Name = $sName Then Return SetError($_S7_Success, 0, 1)
	Next

	Return SetError($_S7_GeneralError, 0, 0)
EndFunc   ;==>__S7_ObjExists
;==============================================================================
Func __S7_GetObject(ByRef $oObj, ByRef $vObject)
	If Not IsObj($oObj) Then Return SetError($_S7_InvalidDataType, 1, "")

	Switch VarGetType($vObject)
		Case "String"
			For $o In $oObj
				If $o.Name = $vObject Then Return SetError($_S7_Success, 0, $o)
			Next
		Case "Int32"
			Local $i = 1
			For $o In $oObj
				If $i = $vObject Then Return SetError($_S7_Success, 0, $o)
				$i += 1
			Next
		Case Else
			Return SetError($_S7_InvalidDataType, 2, "")
	EndSwitch
EndFunc   ;==>__S7_GetObject
;==============================================================================
Func __S7_GetList(ByRef $oObj, ByRef $Type, Const ByRef $Default)
	Local $aRet[1]
	If Not IsObj($oObj) Then Return SetError($_S7_InvalidDataType, 1, $aRet)
	If $Type = Default Then $Type = $Default
	ReDim $aRet[1000]
	Local $i = 0

	For $o In $oObj
		ConsoleWrite("__S7_GetList: " & $o.Name & @CRLF)
		ConsoleWrite("__S7_GetList: " & $o.Type & @CRLF)
		;ConsoleWrite("__S7_GetList: " & $o.Parent & @CRLF)
		If ($Type > $_S7_Any) And ($o.Type = $Type) Then
			$aRet[$i] = $o.Name
			$i += 1
		ElseIf $Type = 0 Then
			$aRet[$i] = $o.Name
			$i += 1
		EndIf
	Next

	If $i > 0 Then
		ReDim $aRet[$i]
	Else
		ReDim $aRet[1]
	EndIf

	Return SetError($_S7_Success, $i, $aRet)
EndFunc   ;==>__S7_GetList
;==============================================================================
Func __S7_GetTypedContainer($oObject, $S7ContainerType)
	If Not IsObj($oObject) Then Return SetError($_S7_InvalidDataType, 1, "")
	For $o In $oObject
		ConsoleWrite("__S7_GetTypedContainer Name: " & $o.Name & @CRLF)
		ConsoleWrite("__S7_GetTypedContainer LogPath: " & $o.LogPath & @CRLF)
		ConsoleWrite("__S7_GetTypedContainer Creator: " & $o.Creator & @CRLF)
		ConsoleWrite("__S7_GetTypedContainer Comment: " & $o.Comment & @CRLF)
		ConsoleWrite("__S7_GetTypedContainer Created: " & $o.Created & @CRLF)
		ConsoleWrite("__S7_GetTypedContainer Modified: " & $o.Modified & @CRLF)
		ConsoleWrite("__S7_GetTypedContainer Version: " & $o.Version & @CRLF)
		ConsoleWrite("__S7_GetTypedContainer Size: " & $o.Size & @CRLF)
		ConsoleWrite("__S7_GetTypedContainer Type: " & $o.Type & @CRLF)
		ConsoleWrite("__S7_GetTypedContainer ConcreteType: " & $o.ConcreteType & @CRLF)
		If $o.Type = $_S7Container Then
			If $o.ConcreteType = $S7ContainerType Then Return SetError(0, 0, $o)
		EndIf
	Next
	Return SetError($_S7_GeneralError, 0, "")
EndFunc   ;==>__S7_GetTypedContainer
;==============================================================================
Func __S7_ErrFunc($oError)
	Local Const $aErrDisc[10] = [ _
			"err.number is", _
			"err.windescription", _
			"err.description", _
			"err.windescription", _
			"err.source", _
			"err.helpfile", _
			"err.helpcontext", _
			"err.lastdllerror", _
			"err.scriptline", _
			"err.retcode"]
	Local $aErr[10]
	 $aErr[0] = $oError.number
	 $aErr[1] = $oError.windescription
	 $aErr[2] = $oError.description
	 $aErr[3] = $oError.source
	 $aErr[4] = $oError.helpfile
	 $aErr[5] = $oError.helpcontext
	 $aErr[6] = $oError.lastdllerror
	 $aErr[7] = $oError.scriptline
	 $aErr[8] = $oError.retcode

	ConsoleWrite("Error in function " & $_S7_Function & ":" & @CRLF)
	For $i = 0 To 9
		If StringStripWS($aErrDisc[$i], 7) <> "" Then ConsoleWrite($aErrDisc[$i] & ": " & $aErr[$i] & @CRLF)
	Next
	 $_S7_Function = "MAIN"

	Return SetError($_S7_ComError, 0, $oError)
EndFunc   ;==>__S7_ErrFunc
