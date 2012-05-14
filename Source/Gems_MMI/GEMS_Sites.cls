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
Attribute VB_Name = "GEMS_Sites"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'local variable to hold collection
Private mCol As Collection
Private mCurr As GEMS_Site
Private mvarIdx As Long                '- current addr in mcol
Private mvarEOF As Boolean

Public Function AddSite(oSite As GEMS_Site) As Boolean
    On Error GoTo errHandler    '- just in case there's already one there
    If Not oSite Is Nothing Then
        mCol.Add oSite
        AddSite = True
        Set mCurr = oSite
        mvarEOF = False
        mvarIdx = mCol.Count
    End If
      
errHandler:
End Function

Public Function Add(sUpdateFlag As String, sFacilityNCID As String, _
            sMedRecNum As String, sRadiologyNum As String, _
            Optional sKey As String) As GEMS_Site
    'create a new object
    Dim objNewMember As GEMS_Site
    On Error GoTo errHandler
    
    Set objNewMember = New GEMS_Site

    'set the properties passed into the method
    objNewMember.UpdateFlag = sUpdateFlag
    objNewMember.FacilityNCID = sFacilityNCID
    objNewMember.MedRecNum = sMedRecNum
    objNewMember.RadiologyNum = sRadiologyNum
    If Len(sKey) > 0 Then
      mCol.Add objNewMember, sKey
    Else
      mCol.Add objNewMember
    End If

    'return the object created
    Set Add = objNewMember
    Set mCurr = objNewMember
    mvarEOF = False
    mvarIdx = mCol.Count
    Set objNewMember = Nothing
errHandler:

End Function

Public Property Get Item(vntIndexKey As Variant) As GEMS_Site
Attribute Item.VB_UserMemId = 0
  
  On Error Resume Next
  Set Item = mCol(vntIndexKey)
  Set mCurr = mCol(vntIndexKey)

End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    On Error Resume Next
    mCol.Remove vntIndexKey
    Set mCurr = New GEMS_Site
    mvarIdx = 0
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    Set mCol = New Collection
    Set mCurr = New GEMS_Site
    mvarIdx = 0
End Sub


Private Sub Class_Terminate()
    Set mCol = Nothing
    Set mCurr = Nothing
End Sub

Public Function MoveNext() As Boolean
    mvarIdx = mvarIdx + 1
    If mvarIdx <= mCol.Count Then
        Set mCurr = mCol.mvarIdx
        mvarEOF = False
        MoveNext = True
    Else
        Set mCurr = New GEMS_Site
        mvarEOF = True
        mvarIdx = 0
        MoveNext = False
    End If
End Function

Public Function MoveFirst() As Boolean
    mvarIdx = 1
    If mvarIdx <= mCol.Count Then
        Set mCurr = mCol.mvarIdx
        mvarEOF = False
        MoveFirst = True
    Else
        Set mCurr = New GEMS_Site
        mvarEOF = True
        mvarIdx = 0
        MoveFirst = False
    End If
 
End Function
'=============================== mcurr stuff ====================
Public Property Get UpdateFlag() As String
    UpdateFlag = mCurr.UpdateFlag
End Property
Public Property Let UpdateFlag(ByVal RHV As String)
    mCurr.UpdateFlag = RHV
End Property

Public Property Get FacilityNCID() As String
    FacilityNCID = mCurr.FacilityNCID
End Property
Public Property Let FacilityNCID(ByVal RHV As String)
    mCurr.FacilityNCID = RHV
End Property

Public Property Get MedRecNum() As String
    MedRecNum = mCurr.MedRecNum
End Property
Public Property Let MedRecNum(ByVal RHV As String)
    mCurr.MedRecNum = RHV
End Property

Public Property Get RadiologyNum() As String
    RadiologyNum = mCurr.RadiologyNum
End Property
Public Property Let RadiologyNum(ByVal RHV As String)
    mCurr.RadiologyNum = RHV
End Property


