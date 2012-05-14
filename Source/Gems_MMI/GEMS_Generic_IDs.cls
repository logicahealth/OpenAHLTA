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
Attribute VB_Name = "GEMS_Generic_IDs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mCol As Collection

Public Function AddGenericID(oGenID As GEMS_Generic_Id)
  On Error Resume Next
  mCol.Add oGenID, CStr(oGenID.id_type_ncid)
  Debug.Print "AddGenericID: Return - " & Err.Number; " adding - " & CStr(oGenID.id_type_ncid)
End Function

'*-- Coleman 061201: Added sUpdateFlag parameter
Public Function Add(sUpdateFlag As String, assigning_authority As Long, id_value As String, id_type_ncid As Long, Optional sKey As String) As GEMS_Generic_Id
    'create a new object
    Dim objNewMember As GEMS_Generic_Id
    Dim bDoingRetry As Boolean
    
    Set objNewMember = New GEMS_Generic_Id

    'set the properties passed into the method
    objNewMember.UpdateFlag = sUpdateFlag   '*-- Coleman 061201
    objNewMember.assigning_authority = assigning_authority
    objNewMember.id_value = id_value
    objNewMember.id_type_ncid = id_type_ncid
    
    On Error GoTo ReplaceOldObject
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing
    Exit Function
    
ReplaceOldObject:
    '*-- Coleman 081301: allow one retry only.
    If (Err.Number = 457) And Not bDoingRetry Then  '*-- 457 = duplicate key
        mCol.Remove sKey
        Resume
    End If
    Set objNewMember = Nothing
    
End Function

Public Property Get Item(vntIndexKey As Variant) As GEMS_Generic_Id
Attribute Item.VB_UserMemId = 0

  On Error Resume Next
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
  On Error Resume Next
    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub



