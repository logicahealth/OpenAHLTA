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
Attribute VB_Name = "cCoreMenuID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--- SCR-18843 MSolano 11/16/01
'--- ADDED CLASS TO MAP KEY VALUE PAIRS
'--- OF APPOINTMENTS MENUS W/ CORE MENUS

Option Explicit

Private mcolCoreMenuIDs As Collection
Const mlMISSING As Long = -1
'--- ENABLE FOR NEXT ENUMERATOR
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  '--- PROC ATTRIBUTES:
  '---  - PROC ID: -4
  '---  - HIDDEN
  Set NewEnum = mcolCoreMenuIDs.[_NewEnum]
End Function

Public Function Item(enmData As AppointmentFilters) As Long
Attribute Item.VB_UserMemId = 0
  '--- PROC ATTRIBUTES: DEFAULT PROPERTY
  Dim sKey As String
  On Error Resume Next
  
  'Dbug.Assert 1 = 2  'removable: temp stop won't get compiled; stays across crashing IDE
  
  sKey = CStr(enmData)
  Item = mlMISSING
  Item = mcolCoreMenuIDs(sKey)
  
End Function

Public Function MenuItem(enmData As AppointmentMenuActions) As Long
  '--- PROC ATTRIBUTES: DEFAULT PROPERTY
  Dim sKey As String
  On Error Resume Next
  
  'Dbug.Assert 1 = 2  'removable: temp stop won't get compiled; stays across crashing IDE
  
  sKey = CStr(enmData)
  MenuItem = mlMISSING
  MenuItem = mcolCoreMenuIDs(sKey)

End Function

Public Function Add( _
                    ByVal lCoreMenuID As Long, _
                    ByVal enmApptKey As AppointmentFilters _
                  ) As Boolean
  Dim lTestCoreID As Long
  Dim sKey As String
  On Error Resume Next

  'dbug.Assert 1 = 2  'removable: temp stop won't get compiled; stays across crashing IDE

  lTestCoreID = mlMISSING
  sKey = CStr(enmApptKey)
  
  '--- test if present
  lTestCoreID = mcolCoreMenuIDs(sKey)
  
  'Debug.Assert lTestCoreID = mlMISSING
  
  If lTestCoreID <> mlMISSING Then
    '--- already in collection
    Exit Function
  Else
    mcolCoreMenuIDs.Add lCoreMenuID, sKey
  End If

End Function

Public Sub Remove(enmAppt As AppointmentFilters)
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
