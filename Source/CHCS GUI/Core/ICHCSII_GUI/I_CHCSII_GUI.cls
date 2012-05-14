VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "I_CHCSII_Gui"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum eControl
    ecClientArea = 1
    ecToolBars
    ecShortcut
    ecTreeView
    ecPicTitleBox
    ecServerTitle
    ecPatientIDLine
    ecActiveBar
End Enum

Public Enum eProperty
    bHoldVisualIndicators = 1
    bHoldUpdateAppItems
    lFileMenuHandle
    lPrintMenuHandle
    lPrintPreviewMenuHandle
    lEditMenuHandle
    lViewMenuHandle
    lToolBarMenuHandle
    lActionBarMenuHandle
    lGoMenuHandle
    lToolsMenuHandle
    lActionsMenuHandle
    lHelpMenuHandle
    bPatientLineVisible
End Enum
Public Property Get WinHandle() As Long: End Property
Public Property Get Control(ByVal sName As eControl) As Object: End Property
Public Property Get Property(ByVal sName As eProperty) As Variant: End Property
Public Property Let Property(ByVal sName As eProperty, ByVal vValue As Variant): End Property

Public Function ShowMsg(ByVal sText As String, _
               Optional ByVal lButtons As Long, _
               Optional ByVal sTitle As String) As VbMsgBoxResult: End Function

Public Function ChangeUser() As Boolean:                        End Function
Public Function GetVisualIndicators() As Scripting.Dictionary:  End Function

Public Sub UpdateAppItems(Optional ByVal vbForceRebuildToolbar As Boolean = False): End Sub
Public Sub UpdateVisualIndicators():                         End Sub
Public Sub DisableWindowHooks(Optional Reset As Boolean):    End Sub
Public Sub UpdateLastSelectedPatInfo(robjPatient As Object): End Sub
Public Sub PatientChanged():    End Sub
Public Sub DisableGUI():        End Sub
Public Sub EnableGUI():         End Sub
Public Sub Display():           End Sub

