'***************************************************************************************
'
'  Copyright (c) 2007-2012 Northrop Grumman Corporation
'
'  Licensed by Tricare Management Activity under license from the Copyright owner.
'
'  This text file must be included in all Derivative Works of the licensed Source Code.
'
'***************************************************************************************

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ModuleSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const BUFFER_LENGTH As Long = 254

Private Const CMD_SWITCH As String = "commandlineswitch"
Private Const APP_NCID As String = "appncid"
Private Const SERVER_NAME As String = "servername"
Private Const APP_TYPE As String = "apptype"
Private Const MENU_TXT As String = "menutext"
Private Const BNT_TXT As String = "buttontext"
Private Const BNT_BMP As String = "buttonbitmap"
Private Const T_TIP As String = "tooltip"
Private Const PAT_REQ As String = "patientrequired"
Private Const ACC_PRIV As String = "accesspriv"
Private Const SUMMARY As String = "summaryitem"
Private Const PARENT_ENTRY As String = "folderparententry"
Private Const FOLDER_SEC As String = "foldersection"

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private mConn As CHCSII_CONN.Conn
Private mMode As CHCSII_CONN.eBackend

'Determines if the External Server has a command line switch associated with it.  If it does then check to see if the
'switch is set.  If the the switch is set then add the external server to the list.
Private Sub AddExternalConfigInfo(ByRef CommandSwitches As Object, ByRef Params As iParameters, ByRef List As iList)
Dim Valid As Boolean

  Valid = Params.Exists(CMD_SWITCH)
  
  If Not Valid Then
    If Len(Params.Item(CMD_SWITCH)) > 0 Then
      If Not (CommandSwitches Is Nothing) Then
        Valid = CommandSwitches(UCase$(Params.Item(CMD_SWITCH))) = UCase$(Params.Item(CMD_SWITCH))
      Else
        Valid = False
      End If
    Else
      Valid = True
    End If
  End If
  
  If Valid Then List.AddRow Params

End Sub

Private Function BuildFullFileName(ByVal Path As String, ByVal File As String) As String

  If Right$(Path, 1) = "\" Then
    BuildFullFileName = Path & File
  Else
    BuildFullFileName = Path & "\" & File
  End If
  
End Function

'Determine the last ordinal number set in the collection of folder items.
Private Function DetermineNextOrdinal(ByRef Items As FolderItems) As Long
Dim NextOrdinal As Long
Dim Item As FolderItem

  For Each Item In Items
    If Item.Ordinal > NextOrdinal Then NextOrdinal = Item.Ordinal
  Next

  DetermineNextOrdinal = NextOrdinal + 1
  
End Function

'Take the Properties string and build a collection of parameters from it (key value pairs).
'Then add the the parameter as a row to a List object.
Private Sub ExtractExternalConfigInfo(ByVal AppName As String, ByVal Properties As String, ByVal INIFileName As String, _
                                      ByRef CommandSwitches As Object, ByRef List As iList)
Dim AppProperties As Variant
Dim i As Long
Dim INIValue As String
Dim rc As Long
Dim Params As iParameters

  On Error GoTo ErrHandler
  
  AppProperties = Split(Properties, Chr(0))
  
  Set Params = New ParametersGeneric
  
  For i = 0 To UBound(AppProperties)
    'Get the value of each property for this app
    If Trim$(AppProperties(i)) <> vbNullString Then
      INIValue = Space$(BUFFER_LENGTH)
            
      rc = GetPrivateProfileString(AppName, CStr(AppProperties(i)), vbNullString, INIValue, BUFFER_LENGTH, INIFileName)
            
      Params.Add LCase$(AppProperties(i)), Left$(INIValue, rc)
    End If
  Next
  
  AddExternalConfigInfo CommandSwitches, Params, List
  
  Exit Sub
    
ErrHandler:
  Err.Raise Err.Number, "MappperCHCSII.ExtractExternalConfigInfo", Err.Description

End Sub

'Reads the CHCSII_APPS.INI file and looks for the following format:
'    [Additional Applications]
'    App1=SampleApp_DontLoad     <-- "App" & seqential-number and name of section below
'
'    [SampleApp_DontLoad]        <-- this must be exactly the same text as listed in [Additional Applications]
'    CommandLineSwitch=GEMS      <-- Any command line switch for this app.
'    appNCID=1999985             <-- the actual application NCID
'    ServerName =TheaterSecurity.TheaterSec  <-- the server name
'    AppType =TOOLS_APP          <-- one of: TOOLS_APP, CLINICAL_APP, SUMMARY_APP, PATIENT_ID_APP, PATIENT_SEARCH_APP, SUMMARY_APP
'    MenuText =&Theater Security <-- text to display on the menu
'    ButtonText =Theater Security    <-- text to display on the button and in the folder tree
'    ButtonBitmap =Theater_Sec.bmp   <-- name of the bitmap file located in the $\images directory
'    ToolTip =Theater Security   <-- text to display as a tool tip
'    PatientRequired =False      <-- Is a patient required before displaying this?  One of: Yes, No, Optional
'    Enabled =False              <-- Is this to be enabled on startup?  One of True, False
'    ButtonSeparatorNeeded =False    <-- should there be a separator to the left of this button?
'    LaunchSequence =            <-- for children of Encounter and A&P and Dental.
'    SummaryItem =False          <-- Does this app support summary view?
'    FolderParentEntry=Tools     <-- the name of the parent folder entry.  Currently one of:Desktop, Tools, Reports, Health History (case sensitive, sorry)
'    FolderSection=Inbox         <-- Which section of the folder tree?  One of Inbox, Chart (case sensitive, sorry)
'    AccessPriv=                 <-- User priv. required to access module
Private Function GetExternalApps(ByRef CommandSwitches As Object) As iList
Dim INIFileName As String
Dim INIValue As String
Dim ExternalApps As Variant
Dim AppName As String
Dim i As Long
Dim CoreShared As CWShared
Dim List As iList
Dim rc As Long

Const APPSECTION As String = "Additional Applications"

  On Error GoTo ErrHandler
  
  Set List = New ListCollection
    
  Set CoreShared = New CWShared
  INIFileName = BuildFullFileName(CoreShared.INIPath, "CHCSII_APPS.INI")
  Set CoreShared = Nothing
  
  If Len(Dir(INIFileName)) = 0 Then
    Set GetExternalApps = List
    Exit Function
  End If
  
  INIValue = Space$(BUFFER_LENGTH)
      
  'get list of applications
  GetPrivateProfileString APPSECTION, 0&, vbNullString, INIValue, BUFFER_LENGTH, INIFileName
  ExternalApps = Split(INIValue, Chr(0))
  
  For i = 0 To UBound(ExternalApps)
    If Len(ExternalApps(i)) = 0 Then Exit For
    
    INIValue = Space$(BUFFER_LENGTH)
    
    'get name of each application
    rc = GetPrivateProfileString(APPSECTION, CStr(ExternalApps(i)), vbNullString, INIValue, BUFFER_LENGTH, INIFileName)
    AppName = Left$(INIValue, rc)
        
    INIValue = Space$(BUFFER_LENGTH)
    
    'Get a list of properites for this app
    GetPrivateProfileString AppName, 0&, vbNullString, INIValue, BUFFER_LENGTH, INIFileName
        
    ExtractExternalConfigInfo AppName, INIValue, INIFileName, CommandSwitches, List
  Next
      
  Set GetExternalApps = List
      
  Exit Function
    
ErrHandler:
  Set GetExternalApps = List
  Err.Raise Err.Number, "MappperCHCSII.GetExternalApps", Err.Description

End Function

'Load OLE Server configurations from the CHCSII_APPS.INI file.
Private Sub GetExternalConfigs(ByRef Configs As ModuleConfigs, ByRef CommandSwitches As Object)
Dim List As iList
Dim Config As ModuleConfig

  Set List = GetExternalApps(CommandSwitches)
  
  With List
    .MoveFirst
    
    Do While Not .EOF
      Set Config = Configs.Add(.Fields.Item(APP_NCID).Value)
      
      Config.AccessPriv = .Fields.Item(ACC_PRIV).Value
      Config.ApplicationNCID = .Fields.Item(APP_NCID).Value
      Config.ApplicationType = .Fields.Item(APP_TYPE).Value
      Config.ButtonCaption = .Fields.Item(BNT_TXT).Value
      Config.CommandLineSwitch = .Fields.Item(CMD_SWITCH).Value
      Config.HasPatientSummary = CBool(.Fields.Item(SUMMARY).Value)
      Config.ImageName = .Fields.Item(BNT_BMP).Value
      Config.MenuCaption = .Fields.Item(MENU_TXT).Value
      
      Select Case LCase$(.Fields.Item(PAT_REQ).Value)
        Case "yes"
          Config.PatientRequired = PATIENT_REQUIRED_YES
          
        Case "no"
          Config.PatientRequired = PATIENT_REQUIRED_NO
          
        Case "optional"
          Config.PatientRequired = PATIENT_REQUIRED_OPTIONAL
      End Select

      Config.Selected = True
      Config.ServerName = .Fields.Item(SERVER_NAME).Value
      Config.ToolTip = .Fields.Item(T_TIP).Value
      
      .MoveNext
    Loop
  End With
  
  Exit Sub
  
ErrHandler:
  Err.Raise Err.Number, "MappperCHCSII.GetExternalConfigs", Err.Description
  
End Sub

'Load Folder Items from the CHCSII_APPS.INI file.
Private Sub GetExternalFolders(ByRef Items As FolderItems)
Dim List As iList
Dim Item As FolderItem
Dim NextOrdinal As Long

  Set List = GetExternalApps(Nothing)
  
  NextOrdinal = DetermineNextOrdinal(Items)
  
  With List
    Do While Not .EOF
      .MoveFirst
    
      Set Item = Items.Add(NextOrdinal)
      
      Item.ApplicationNCID = .Fields.Item(APP_NCID).Value
      Item.ImageName = "closedFolder"
      Item.Key = .Fields.Item(BNT_TXT).Value
      Item.NodeText = .Fields.Item(BNT_TXT).Value
      Item.Section = .Fields.Item(FOLDER_SEC).Value
      Item.Ordinal = NextOrdinal
      Item.ParentOrdinal = .Fields.Item(PARENT_ENTRY).Value
      
      NextOrdinal = NextOrdinal + 1
      .MoveNext
    Loop
  End With
    
  If List.RecordCount > 0 Then
    SetExternalFoldersParentOrdinal Items
  End If
  
  Exit Sub
  
ErrHandler:
  Err.Raise Err.Number, "MappperCHCSII.GetExternalFolders", Err.Description

End Sub

Public Sub Init(ByRef Conn As CHCSII_CONN.Conn)
  Set mConn = Conn
End Sub

'Retrieves OLE Server (modules) configurations from Osh_Ini_Settings table.
Private Function LoadConfigs(ByRef CommandSwitches As DL_Support.iParameters, ByVal ReturnAll As Boolean) As ModuleConfigs
Dim SQLOps As ICHCSII_SQL.ISqlOps
Dim Configs As ModuleConfigs
Dim Config As ModuleConfig
Dim ConfigValues As Variant
Dim Helper As ModuleConfigHelper
Dim ValidConfig As Boolean

Const CONFIG_SETTINGS_SQL As String = "SELECT ApKey, ApValue FROM OSH_INI_SETTINGS WHERE Application_NCID = '1801' AND ApSection = 'cw applications'"
Const OSH_APP_NCID As Integer = 6

  On Error GoTo ErrHandler
    
  Set Configs = New ModuleConfigs
  Set Helper = New ModuleConfigHelper
     
  Set SQLOps = LoadSQLOps(CONFIG_SETTINGS_SQL)                              'Load Config Settings
  
  Do While Not SQLOps.EOF                                                   'Add ModuleConfig objects from List object to ModuleConfigs object
    ConfigValues = Split(SQLOps.Value("apvalue"), ";")                      'Settings are in a ; separated string
  
    Set Config = Configs.Add(ConfigValues(OSH_APP_NCID))

    Config.ServerName = SQLOps("apkey")

    ValidConfig = Helper.ExtractConfigInfo(Config, ConfigValues, ReturnAll, CommandSwitches)
    
    If Not ValidConfig Then
      Configs.Remove ConfigValues(OSH_APP_NCID)
    End If
    
    Set Config = Nothing
    
    SQLOps.MoveNext
  Loop
  
  GetExternalConfigs Configs, CommandSwitches
  
  Set LoadConfigs = Configs

  Exit Function

ErrHandler:
  Set LoadConfigs = Configs
  Err.Raise Err.Number, "MappperCHCSII.LoadConfigs", Err.Description

End Function

Private Function LoadSQLOps(ByVal SelectStmt As String) As ISqlOps
Dim SQLOps As ICHCSII_SQL.ISqlOps

  Set SQLOps = mConn.CHCSII_SQLOPS(mMode)
  SQLOps.Execute SelectStmt

  Set LoadSQLOps = SQLOps
  
End Function

'Currently only 2 modules are loaded for a patient (user with Role of Patient).
Private Function LoadUserPatientConfigs() As ModuleConfigs
Dim Configs As ModuleConfigs
Dim Config As ModuleConfig

Const NCID_HEAR_APPLICATION As String = "184918"
Const NCID_PATIENT_QUESTIONNAIRES As String = "1999996"
Const CLINICAL_APP_TYPE As String = "CLINICAL_APP"

  On Error GoTo ErrHandler
  
  Set Configs = New ModuleConfigs
  
  Set Config = Configs.Add(NCID_HEAR_APPLICATION)
  
  With Config
    .ApplicationNCID = NCID_HEAR_APPLICATION
    .ApplicationType = CLINICAL_APP_TYPE
    .MenuCaption = "Hears"
    .ButtonCaption = "Hears"
    .ImageName = "hear.bmp"
    .ToolTip = "Complete HEARS Questionaire"
    .PatientRequired = PATIENT_REQUIRED_NO
    .Selected = True
    .ServerName = "CouplerLauncher.Coupler"
    .HasPatientSummary = False
    .HasPrint = False
    .HasPrintPreview = False
    .HasPropertyView = False
  End With
  
  Set Config = Nothing
  Set Config = Configs.Add(NCID_PATIENT_QUESTIONNAIRES)
  
  With Config
      .ApplicationNCID = NCID_PATIENT_QUESTIONNAIRES
      .ServerName = "CHCSII_PQuest.PQuest"
      .ApplicationType = CLINICAL_APP_TYPE
      .MenuCaption = "&Patient Questionnaires"
      .ButtonCaption = "Questionnaires"
      .ImageName = "Health_History.bmp"
      .ToolTip = "Answer Patient Questionnaires"
      .PatientRequired = PATIENT_REQUIRED_NO
      .Selected = False
      .HasPrint = False
      .HasPrintPreview = False
      .HasPatientSummary = False
      .HasPropertyView = False
  End With
  
  Set LoadUserPatientConfigs = Configs
  
  Exit Function
    
ErrHandler:
  Set LoadUserPatientConfigs = Configs
  Err.Raise Err.Number, "MappperCHCSII.LoadUserPatientConfigs", Err.Description

End Function

'Gets the abbr rank description for a given Rank NCID.
Public Function RetrieveAbbrRank(ByVal RankNCID As String) As String
Dim SQLOps As ICHCSII_SQL.ISqlOps
Dim SQL As String

Const RANK_SQL As String = "SELECT abbrev_rank from rank_abbrev WHERE ncid = '%NCID'"
Const FLD_AB_RANK As String = "abbrev_rank"

  SQL = Replace(RANK_SQL, "%NCID", RankNCID)
  
  Set SQLOps = LoadSQLOps(SQL)
  
  If Not SQLOps.EOF Then
    RetrieveAbbrRank = SQLOps.Value(FLD_AB_RANK)
  End If
  
End Function

'Retrieves Folder Items from Folder_Items table.
Public Function RetrieveFolderItems() As DTOs.FolderItems
Dim SQLOps As ICHCSII_SQL.ISqlOps
Dim Items As FolderItems
Dim Item As FolderItem

Const FOLDERS_SQL As String = "SELECT * FROM Folder_Items ORDER BY Ordinal"
  
  On Error GoTo ErrHandler
      
  Set Items = New FolderItems
          
  Set SQLOps = LoadSQLOps(FOLDERS_SQL)                              'Load Folder Items
    
  Do While Not SQLOps.EOF                                           'Add FolderItem objects from List object to FolderItems object
    Set Item = Items.Add(CStr(SQLOps("ordinal")))
    
    With Item
      .ApplicationNCID = SQLOps("appncid")
      .ParentOrdinal = SQLOps("parententry")
      .Key = SQLOps("appname")
      .ImageName = "closedFolder"
      .NodeText = SQLOps("buttontext") & vbNullString
      .Ordinal = SQLOps("ordinal")
      .Section = SQLOps("section")
    End With
        
    Set Item = Nothing
    
    SQLOps.MoveNext
  Loop
    
  Set SQLOps = Nothing
        
  GetExternalFolders Items
  
  Set RetrieveFolderItems = Items
  
  Exit Function
  
ErrHandler:
  Err.Raise Err.Number, "ModuleSetup.RetrieveFolderItems", Err.Description

End Function

'Returns the configuration for modules.
'If UserIsPatient is True then modules used by a patient are loaded only.
'If ReturnAll is true then all module configurations are returned even ones that are not selected.
Public Function RetrieveModuleConfigurations(ByRef CommandSwitches As DL_Support.iParameters, Optional ByVal UserIsPatient As Boolean = False, Optional ByVal ReturnAll As Boolean = False) As DTOs.ModuleConfigs

  If Not UserIsPatient Then
    Set RetrieveModuleConfigurations = LoadConfigs(CommandSwitches, ReturnAll)
  Else
    Set RetrieveModuleConfigurations = LoadUserPatientConfigs
  End If

End Function

'External servers will store the name of the parent in the parent ordinal property.
'So we need to find the parent ordinal and replace the value with the numeric value.
Private Sub SetExternalFoldersParentOrdinal(ByRef Items As FolderItems)
Dim i As Long
Dim j As Long

  For i = 1 To Items.Count
    If Val(Items.Item(i).ParentOrdinal) = 0 Then
    
      For j = 1 To Items.Count
        If UCase$(Items.Item(j).NodeText) = UCase$(Items.Item(i).ParentOrdinal) Then
          Items.Item(i).ParentOrdinal = Items.Item(j).Ordinal
          
          Exit For
        End If
      Next
      
      If Val(Items.Item(i).ParentOrdinal) = 0 Then
        Items.Item(i).ParentOrdinal = 1
      End If
    End If
  Next

End Sub

