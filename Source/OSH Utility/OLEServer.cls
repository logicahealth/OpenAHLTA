VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OLEServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


Public lProgId                      As Long
Public bInitial                    As Boolean

Private msApplicationName           As String
Private msExecutable                As String
Private msAppNCID                   As String
Private msAppType                   As String
Private msMenuString                As String
Private mlMenuId                    As Long
Private mlItemSpecificHelpId        As Long
Private msIconString                As String
Private msButtonBitmap              As String
Private msToolTip                   As String
Private mbEnabled                   As Boolean
Private mbButtonSeparatorNeeded     As Boolean
Private miPatientRequired           As Integer  '-
Private mbRegularViewRunning        As Boolean
Private mbSummaryViewRunning        As Boolean   '- added as part of 'move summary'
Private mbUserHasAccess             As Boolean
Private miEncounterDocServer        As Integer   '- is this svr shown under Encounters? 0-No, 1-Yes only, 2-Yes also
Private msLaunchSequence            As String    '- Used To launch submodules under Current Encounters
Private miTreeOrder                 As Integer   '- Used to keep Go Menu, ToolBar and TreeView ordering in sync
Private msOKToChangePatient         As String    '- 'Y' or 'N'
Private msOKToCloseView             As String    '- 'Y' or 'N'
Private mbDisablePropertyView       As Boolean   '- True if module doesn't support property view  - check each time it comes up?
Private mbSupportsPatSummary        As Boolean
Private mbPrint                     As Boolean
Private mbPrintPreview              As Boolean
Private mbExternal                  As Boolean   '- is this an "external" app?
Private mFolderParentEntry          As String    '- the name of the parent folder entry.  Currently one of: Tools, Reports, Health History (case sensitive, sorry)
Private mFolderSection              As String    '- Which section of the folder tree?  One of Inbox, Chart (case sensitive, sorry)
Private msAccessPriv                As String    '- the string for "has privilege" function
Private msCommandLine               As String
Private mbSelected                  As Boolean

'_________________ internal copy set on initial load, used to see if data changed
Private zlProgId                    As Long
Private zhWnd                       As Long
Private zsApplicationName           As String
Private zsExecutable                As String
Private zsAppNCID                   As String
Private zsAppType                   As String
Private zsMenuString                As String
Private zlMenuId                    As Long
Private zlItemSpecificHelpId        As Long
Private zsIconString                As String
Private zsButtonBitmap              As String
Private zsToolTip                   As String
Private zbEnabled                   As Boolean
Private zbButtonSeparatorNeeded     As Boolean
Private ziPatientRequired           As Integer  '-
Private zbRegularViewRunning        As Boolean
Private zbSummaryViewRunning        As Boolean   '- added as part of 'move summary'
Private zbUserHasAccess             As Boolean
Private ziEncounterDocServer        As Integer   '- is this svr shown under Encounters? 0-No, 1-Yes only, 2-Yes also
Private zsLaunchSequence            As String    '- Used To launch submodules under Current Encounters
Private ziTreeOrder                 As Integer   '- Used to keep Go Menu, ToolBar and TreeView ordering in sync
Private zsOKToChangePatient         As String    '- 'Y' or 'N'
Private zsOKToCloseView             As String    '- 'Y' or 'N'
Private zbDisablePropertyView       As Boolean   '- True if module doesn't support property view  - check each time it comes up?
Private zbSupportsPatSummary        As Boolean
Private zbPrint                     As Boolean
Private zbPrintPreview              As Boolean

Private zbExternal                  As Boolean   '- is this an "external" app?
Private zFolderParentEntry          As String    '- the name of the parent folder entry.  Currently one of: Tools, Reports, Health History (case sensitive, sorry)
Private zFolderSection              As String    '- Which section of the folder tree?  One of Inbox, Chart (case sensitive, sorry)
Private zsAccessPriv                As String    '- the string for "has privilege" function
Private zsCommandLine               As String
Private zbSelected                  As Boolean


'____________________________________________
Public bNew                       As Boolean
Public bObjectUpdated             As Boolean
Public sOriginalData              As String


Public Property Get bChanged() As Boolean

    bChanged = (zsApplicationName = msApplicationName)
    bChanged = bChanged And (zsExecutable = msExecutable)
    bChanged = bChanged And (zsAppNCID = msAppNCID)
    bChanged = bChanged And (zsAppType = msAppType)
    bChanged = bChanged And (zsMenuString = msMenuString)
    bChanged = bChanged And (zlMenuId = mlMenuId)
    bChanged = bChanged And (zlItemSpecificHelpId = mlItemSpecificHelpId)
    bChanged = bChanged And (zsIconString = msIconString)
    bChanged = bChanged And (zsButtonBitmap = msButtonBitmap)
    bChanged = bChanged And (zsToolTip = msToolTip)
    bChanged = bChanged And (zbEnabled = mbEnabled)
    bChanged = bChanged And (zbButtonSeparatorNeeded = mbButtonSeparatorNeeded)
    bChanged = bChanged And (ziPatientRequired = miPatientRequired)
    bChanged = bChanged And (zbRegularViewRunning = mbRegularViewRunning)
    bChanged = bChanged And (zbSummaryViewRunning = mbSummaryViewRunning)
    bChanged = bChanged And (zbUserHasAccess = mbUserHasAccess)
    bChanged = bChanged And (ziEncounterDocServer = miEncounterDocServer)
    bChanged = bChanged And (zsLaunchSequence = msLaunchSequence)
    bChanged = bChanged And (ziTreeOrder = miTreeOrder)
    bChanged = bChanged And (zsOKToChangePatient = msOKToChangePatient)
    bChanged = bChanged And (zsOKToCloseView = msOKToCloseView)
    bChanged = bChanged And (zbDisablePropertyView = mbDisablePropertyView)
    bChanged = bChanged And (zbSupportsPatSummary = mbSupportsPatSummary)
    bChanged = bChanged And (zbPrint = mbPrint)
    bChanged = bChanged And (zbPrintPreview = mbPrintPreview)
    bChanged = bChanged And (zbExternal = mbExternal)
    bChanged = bChanged And (zFolderParentEntry = mFolderParentEntry)
    bChanged = bChanged And (zFolderSection = mFolderSection)
    bChanged = bChanged And (zsAccessPriv = msAccessPriv)
    bChanged = bChanged And (zsCommandLine = msCommandLine)
    bChanged = bChanged And (zbSelected = mbSelected)
    bChanged = Not bChanged
End Property

Public Function ResetData() As Boolean
    msApplicationName = zsApplicationName
    msExecutable = zsExecutable
    msAppNCID = zsAppNCID
    msAppType = zsAppType
    msMenuString = zsMenuString
    mlMenuId = zlMenuId
    mlItemSpecificHelpId = zlItemSpecificHelpId
    msIconString = zsIconString
    msButtonBitmap = zsButtonBitmap
    msToolTip = zsToolTip
    mbEnabled = zbEnabled
    mbButtonSeparatorNeeded = zbButtonSeparatorNeeded
    miPatientRequired = ziPatientRequired
    mbRegularViewRunning = zbRegularViewRunning
    mbSummaryViewRunning = zbSummaryViewRunning
    mbUserHasAccess = zbUserHasAccess
    miEncounterDocServer = ziEncounterDocServer
    msLaunchSequence = zsLaunchSequence
    miTreeOrder = ziTreeOrder
    msOKToChangePatient = zsOKToChangePatient
    msOKToCloseView = zsOKToCloseView
    mbDisablePropertyView = zbDisablePropertyView
    mbSupportsPatSummary = zbSupportsPatSummary
    mbPrint = zbPrint
    mbPrintPreview = zbPrintPreview
    mbExternal = zbExternal
    mFolderParentEntry = zFolderParentEntry
    mFolderSection = zFolderSection
    msAccessPriv = zsAccessPriv
    msCommandLine = zsCommandLine
    mbSelected = zbSelected

End Function
Public Function FieldChanged(ByVal sField As String) As Boolean
Select Case sField
Case "ApplicationName":                 FieldChanged = Not (zsApplicationName = msApplicationName)
Case "Executable":                      FieldChanged = Not (zsExecutable = msExecutable)
Case "AppNCID":                         FieldChanged = Not (zsAppNCID = msAppNCID)
Case "AppType":                         FieldChanged = Not (zsAppType = msAppType)
Case "MenuString":                      FieldChanged = Not (zsMenuString = msMenuString)
Case "MenuId":                          FieldChanged = Not (zlMenuId = mlMenuId)
Case "ItemSpecificHelpId":              FieldChanged = Not (zlItemSpecificHelpId = mlItemSpecificHelpId)
Case "IconString":                      FieldChanged = Not (zsIconString = msIconString)
Case "ButtonBitmap":                    FieldChanged = Not (zsButtonBitmap = msButtonBitmap)
Case "ToolTip":                         FieldChanged = Not (zsToolTip = msToolTip)
Case "Enabled":                         FieldChanged = Not (zbEnabled = mbEnabled)
Case "ButtonSeparatorNeeded":           FieldChanged = Not (zbButtonSeparatorNeeded = mbButtonSeparatorNeeded)
Case "PatientRequired":                 FieldChanged = Not (ziPatientRequired = miPatientRequired)
Case "RegularViewRunning":              FieldChanged = Not (zbRegularViewRunning = mbRegularViewRunning)
Case "SummaryViewRunning":              FieldChanged = Not (zbSummaryViewRunning = mbSummaryViewRunning)
Case "UserHasAccess":                   FieldChanged = Not (zbUserHasAccess = mbUserHasAccess)
Case "EncounterDocServer":              FieldChanged = Not (ziEncounterDocServer = miEncounterDocServer)
Case "LaunchSequence":                  FieldChanged = Not (zsLaunchSequence = msLaunchSequence)
Case "TreeOrder":                       FieldChanged = Not (ziTreeOrder = miTreeOrder)
Case "OKToChangePatient":               FieldChanged = Not (zsOKToChangePatient = msOKToChangePatient)
Case "OKToCloseView":                   FieldChanged = Not (zsOKToCloseView = msOKToCloseView)
Case "DisablePropertyView":             FieldChanged = Not (zbDisablePropertyView = mbDisablePropertyView)
Case "SupportsPatSummary":              FieldChanged = Not (zbSupportsPatSummary = mbSupportsPatSummary)
Case "Print":                           FieldChanged = Not (zbPrint = mbPrint)
Case "PrintPreview":                    FieldChanged = Not (zbPrintPreview = mbPrintPreview)
Case "External":                        FieldChanged = Not (zbExternal = mbExternal)
Case "olderParentEntry":                FieldChanged = Not (zFolderParentEntry = mFolderParentEntry)
Case "olderSection":                    FieldChanged = Not (zFolderSection = mFolderSection)
Case "AccessPriv":                      FieldChanged = Not (zsAccessPriv = msAccessPriv)
Case "CommandLine":                     FieldChanged = Not (zsCommandLine = msCommandLine)
Case "Selected":                        FieldChanged = Not (zbSelected = mbSelected)
End Select

End Function



Public Property Get sApplicationName() As String:               sApplicationName = msApplicationName: End Property
Public Property Get sExecutable() As String:                    sExecutable = msExecutable: End Property
Public Property Get sAppNCID() As String:                       sAppNCID = msAppNCID: End Property
Public Property Get sAppType() As String:                       sAppType = msAppType: End Property
Public Property Get sMenuString() As String:                    sMenuString = msMenuString: End Property
Public Property Get lMenuId() As Long:                          lMenuId = mlMenuId: End Property
Public Property Get lItemSpecificHelpId() As Long:              lItemSpecificHelpId = mlItemSpecificHelpId: End Property
Public Property Get sIconString() As String:                    sIconString = msIconString: End Property
Public Property Get sButtonBitmap() As String:                  sButtonBitmap = msButtonBitmap: End Property
Public Property Get sToolTip() As String:                       sToolTip = msToolTip: End Property
Public Property Get bEnabled() As Boolean:                      bEnabled = mbEnabled: End Property
Public Property Get bButtonSeparatorNeeded() As Boolean:        bButtonSeparatorNeeded = mbButtonSeparatorNeeded: End Property
Public Property Get iPatientRequired() As Integer:              iPatientRequired = miPatientRequired: End Property
Public Property Get bRegularViewRunning() As Boolean:           bRegularViewRunning = mbRegularViewRunning: End Property
Public Property Get bSummaryViewRunning() As Boolean:           bSummaryViewRunning = mbSummaryViewRunning: End Property
Public Property Get bUserHasAccess() As Boolean:                bUserHasAccess = mbUserHasAccess: End Property
Public Property Get iEncounterDocServer() As Integer:           iEncounterDocServer = miEncounterDocServer: End Property
Public Property Get sLaunchSequence() As String:                sLaunchSequence = msLaunchSequence: End Property
Public Property Get iTreeOrder() As Integer:                    iTreeOrder = miTreeOrder: End Property
Public Property Get sOKToChangePatient() As String:             sOKToChangePatient = msOKToChangePatient: End Property
Public Property Get sOKToCloseView() As String:                 sOKToCloseView = msOKToCloseView: End Property
Public Property Get bDisablePropertyView() As Boolean:          bDisablePropertyView = mbDisablePropertyView: End Property
Public Property Get bSupportsPatSummary() As Boolean:           bSupportsPatSummary = mbSupportsPatSummary: End Property
Public Property Get bPrint() As Boolean:                        bPrint = mbPrint: End Property
Public Property Get bPrintPreview() As Boolean:                 bPrintPreview = mbPrintPreview: End Property
Public Property Get bExternal() As Boolean:                     bExternal = mbExternal: End Property
Public Property Get FolderParentEntry() As String:              FolderParentEntry = mFolderParentEntry: End Property
Public Property Get FolderSection() As String:                  FolderSection = mFolderSection: End Property
Public Property Get sAccessPriv() As String:                    sAccessPriv = msAccessPriv: End Property
Public Property Get sCommandLine() As String:                   sCommandLine = msCommandLine: End Property
Public Property Get bSelected() As Boolean:                     bSelected = mbSelected: End Property

Public Property Let sApplicationName(RHV As String)
    msApplicationName = RHV
    If bInitial Then zsApplicationName = RHV
    End Property
Public Property Let sExecutable(RHV As String)
    msExecutable = RHV
    If bInitial Then zsExecutable = RHV
    End Property
Public Property Let sAppNCID(RHV As String)
    msAppNCID = RHV
    If bInitial Then zsAppNCID = RHV
    End Property
Public Property Let sAppType(RHV As String)
    msAppType = RHV
    If bInitial Then zsAppType = RHV
    End Property
Public Property Let sMenuString(RHV As String)
    msMenuString = RHV
    If bInitial Then zsMenuString = RHV
    End Property
Public Property Let lMenuId(RHV As Long)
    mlMenuId = RHV
    If bInitial Then zlMenuId = RHV
    End Property
Public Property Let lItemSpecificHelpId(RHV As Long)
    mlItemSpecificHelpId = RHV
    If bInitial Then zlItemSpecificHelpId = RHV
    End Property
Public Property Let sIconString(RHV As String)
    msIconString = RHV
    If bInitial Then zsIconString = RHV
    End Property
Public Property Let sButtonBitmap(RHV As String)
    msButtonBitmap = RHV
    If bInitial Then zsButtonBitmap = RHV
    End Property
Public Property Let sToolTip(RHV As String)
    msToolTip = RHV
    If bInitial Then zsToolTip = RHV
    End Property
Public Property Let bEnabled(RHV As Boolean)
    mbEnabled = RHV
    If bInitial Then zbEnabled = RHV
    End Property
Public Property Let bButtonSeparatorNeeded(RHV As Boolean)
    mbButtonSeparatorNeeded = RHV
    If bInitial Then zbButtonSeparatorNeeded = RHV
    End Property
Public Property Let iPatientRequired(RHV As Integer)
    miPatientRequired = RHV
    If bInitial Then ziPatientRequired = RHV
    End Property
Public Property Let bRegularViewRunning(RHV As Boolean)
    mbRegularViewRunning = RHV
    If bInitial Then zbRegularViewRunning = RHV
    End Property
Public Property Let bSummaryViewRunning(RHV As Boolean)
    mbSummaryViewRunning = RHV
    If bInitial Then zbSummaryViewRunning = RHV
    End Property
Public Property Let bUserHasAccess(RHV As Boolean)
    mbUserHasAccess = RHV
    If bInitial Then zbUserHasAccess = RHV
    End Property
Public Property Let iEncounterDocServer(RHV As Integer)
    miEncounterDocServer = RHV
    If bInitial Then ziEncounterDocServer = RHV
    End Property
Public Property Let sLaunchSequence(RHV As String)
    msLaunchSequence = RHV
    If bInitial Then zsLaunchSequence = RHV
    End Property
Public Property Let iTreeOrder(RHV As Integer)
    miTreeOrder = RHV
    If bInitial Then ziTreeOrder = RHV
    End Property
Public Property Let sOKToChangePatient(RHV As String)
    msOKToChangePatient = RHV
    If bInitial Then zsOKToChangePatient = RHV
    End Property
Public Property Let sOKToCloseView(RHV As String)
    msOKToCloseView = RHV
    If bInitial Then zsOKToCloseView = RHV
    End Property
Public Property Let bDisablePropertyView(RHV As Boolean)
    mbDisablePropertyView = RHV
    If bInitial Then zbDisablePropertyView = RHV
    End Property
Public Property Let bSupportsPatSummary(RHV As Boolean)
    mbSupportsPatSummary = RHV
    If bInitial Then zbSupportsPatSummary = RHV
    End Property
Public Property Let bPrint(RHV As Boolean)
    mbPrint = RHV
    If bInitial Then zbPrint = RHV
    End Property
Public Property Let bPrintPreview(RHV As Boolean)
    mbPrintPreview = RHV
    If bInitial Then zbPrintPreview = RHV
    End Property
Public Property Let bExternal(RHV As Boolean)
    mbExternal = RHV
    If bInitial Then zbExternal = RHV
    End Property
Public Property Let FolderParentEntry(RHV As String)
    mFolderParentEntry = RHV
    If bInitial Then zFolderParentEntry = RHV
    End Property
Public Property Let FolderSection(RHV As String)
    mFolderSection = RHV
    If bInitial Then zFolderSection = RHV
    End Property
Public Property Let sAccessPriv(RHV As String)
    msAccessPriv = RHV
    If bInitial Then zsAccessPriv = RHV
    End Property
Public Property Let sCommandLine(RHV As String)
    msCommandLine = RHV
    If bInitial Then zsCommandLine = RHV
    End Property
Public Property Let bSelected(RHV As Boolean)
    mbSelected = RHV
    If bInitial Then zbSelected = RHV
    End Property


