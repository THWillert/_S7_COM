# _S7_COM
AutoIt Bibliothek für Siemens Step 7


### Simatic
Simatic.Simatic.1
-	OBJ 	_S7_Simatic_ObjCreate / Simatic = Simatic.Simatic.1
-	BOOL 	_S7_Simatic_AutomaticSave / Simatic.AutomaticSave (Read / Write)
-	BSTR 	_S7_Simatic_VerbLogFile (read write)
-	_S7_Simatic_SetPGInterface / Simatic.SetPGInterface (Opens Dialog)
-	_S7_Simatic_UnattendedServerMode !!!
-	LONG 	_S7_Simatic_MsgAssignmentType / Simatic.MsgAssignmentType (Read / Write)
-	BOOL 	_S7_Simatic_IsSilentMode (read)
-	_S7_Simatic_Save / Simatic.Save (void)
-	HRESULT _S7_Simatic_GetSTEP7Language / Simatic.GetSTEP7Language (Read) ??? funktioniert nicht mit AutoIt

### Projects
Simatic.Projects
-	OBJ 	_S7_Projects_GetProject
-	BOOL 	_S7_Projects_Exists
-	Array 	_S7_Projects_GetList
-	Int 	_S7_Projects_Count
-	BOOL 	_S7_Projects_Add

### Project
Simatic.Projects.Project
-	Array 	_S7_Project_GetInfo
-	_S7_Project_Name
-	String	_S7_Project_Creator (Read / Write)
-	String	_S7_Project_Comment (Read / Write)
-	Bool 	_S7_Project_Remove

### Stations
Simatic.Projects.Project.Stations
-	OBJ 	_S7_Stations_GetStation
-	Bool	_S7_Stations_Exists
-	Array	_S7_Stations_GetList
-	Int	_S7_Stations_Count
-	Bool	_S7_Stations_Import
-	Bool	_S7_Stations_Add
- Bool	_S7_Stations_Remove

### Programs
Simatic.Projects.Project.Programs
-	Array 	_S7_Programs_GetList

### Blocks
Simatic.Projects.Project.Programs.Next("Blocks")
-	OBJ	_S7_Blocks_GetBlock

### Source Files
Simatic.Projects.Project.Programs.Next("Source Files")
-	OBJ	_S7_SourceFiles_GetSource
-	Bool	_S7_SourceFiles_Export
-	Bool	_S7_SourceFiles_Add
-	Bool	_S7_SourceFiles_Compile
-	Array	_S7_SourceFiles_GetInfo
