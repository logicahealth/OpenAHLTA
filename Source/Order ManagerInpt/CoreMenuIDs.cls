VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCoreMenuID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mcolCoreMenuIDs As Collection
Private Const mlMISSING As Long = -1
'--- ENABLE FOR NEXT ENUMERATOR
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = mcolCoreMenuIDs.[_NewEnum]
End Function

Public Function Item(enmData As orderFilters) As Long
Attribute Item.VB_UserMemId = 0
  '--- PROC ATTRIBUTES: DEFAULT PROPERTY
  Dim sKey As String
  On Error Resume Next
    
  sKey = CStr(enmData)
  Item = mlMISSING
  Item = mcolCoreMenuIDs(sKey)
  
End Function

Public Function MenuItem(enmData As orderFilters) As Long
  Dim sKey As String
  On Error Resume Next
    
  sKey = CStr(enmData)
  MenuItem = mlMISSING
  MenuItem = mcolCoreMenuIDs(sKey)

End Function

Public Function Add(ByVal lCoreMenuID As Long, ByVal enmApptKey As orderFilters) As Boolean
  Dim lTestCoreID As Long
  Dim sKey As String
  On Error Resume Next

  lTestCoreID = mlMISSING
  sKey = CStr(enmApptKey)
  
  lTestCoreID = mcolCoreMenuIDs(sKey)
    
  If lTestCoreID <> mlMISSING Then
    Exit Function
  Else
    mcolCoreMenuIDs.Add lCoreMenuID, sKey
  End If

End Function

Public Sub Remove(enmAppt As orderFilters)
  On Error Resume Next
  Dim sKey As String
  On Error Resume Next
  
  sKey = CStr(enmAppt)
  mcolCoreMenuIDs.Remove sKey
  
End Sub

Private Sub Class_Initialize()
  Set mcolCoreMenuIDs = New Collection
End Sub

Private Sub Class_Terminate()
  Set mcolCoreMenuIDs = Nothing
End Sub
