VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ICHCSII_CoreAPI1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'StartUp Order
'Initialize
'IsOkToOpenView
'Display

Public Enum enumViewType
  vtAll = 0
  vtClinicalSummary = 1
  vtEncounterSummary = 2
  vtPatientSummary = 3
  vtRegular = 4
  vtAnP = 5
  vtPreviousEncounter = 6
  vtProperty = 7
  vtPrint = 8
  vtPrintPreview = 9
End Enum

Public Sub About()
End Sub

Public Function CloseView(ByVal ViewType As enumViewType) As Boolean
End Function

Public Function Display(ByVal ViewType As enumViewType, _
                   ByVal ParenthWnd As Long, _
                   ByRef MainForm As Long, _
                   ByRef MyMenus As iMenus, _
                   ByVal Left As Long, _
                   ByVal Top As Long, _
                   ByVal Width As Long, _
                   ByVal Height As Long) As Boolean
End Function

Public Function Initialize(ByVal MyProgId As Long, _
                      ByRef Comm As Object, _
                      ByRef LogonEnvironment As Object, _
                      ByRef Config As Object, _
                      ByRef Patient As Object, _
                      ByRef User As Object) As Boolean
End Function

Public Function IsOkToChangePatient() As Boolean
End Function

Public Function IsOkToCloseView() As Boolean
End Function

Public Function IsOKToOpenView(ByVal ViewType As enumViewType) As Boolean
End Function

Public Sub MenuItemSelected(ByVal Caption As String, ByVal MenuID As Long)
End Sub

Public Function Message(ByVal MessageType As Long, _
                        ByVal SourceNCID As String, _
                        Optional ByRef Value As Variant = vbNullString) As Boolean
End Function

Public Sub PatientSelected()

End Sub

Public Sub Properties(ByRef LogonEnvironment As Object, _
                      ByRef Config As Object, _
                      ByVal OwnerNCID As String, _
                      ByVal CalledByDesigner As Boolean)
End Sub

Public Function QueryView(ByVal ViewType As enumViewType) As Boolean
End Function

Public Sub ResizeView(ByVal ViewType As enumViewType, _
                      ByVal Left As Long, _
                      ByVal Top As Long, _
                      ByVal Width As Long, _
                      ByVal Height As Long)
End Sub


