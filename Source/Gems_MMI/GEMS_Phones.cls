VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GEMS_Phones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"GEMS_Phone"
Attribute VB_Ext_KEY = "Member0" ,"GEMS_Phone"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable to hold collection
Private mCol As Collection

Public Function AddPhone(oPhone As GEMS_Phone) As Boolean
  On Error GoTo errHandler
  If Not oPhone Is Nothing Then
    If Not oPhone.PhoneType = "" Then
      mCol.Add oPhone, CStr(oPhone.PhoneType)
    Else
      mCol.Add oPhone
    End If
    AddPhone = True
  End If
errHandler:
End Function

Public Function Add(UpdateFlag As String, Extension As String, Phone As String, AreaCode As String, CountryCode As String, PhoneType As String) As GEMS_Phone
    'create a new object
    Dim objNewMember As GEMS_Phone
    Dim bAlreadyDidRetry As Boolean
    
    On Error GoTo CheckForDuplicate
    
    Set objNewMember = New GEMS_Phone

    'set the properties passed into the method
    objNewMember.UpdateFlag = UpdateFlag
    objNewMember.Extension = Extension
    objNewMember.Phone = Phone
    objNewMember.AreaCode = AreaCode
    objNewMember.CountryCode = CountryCode
    objNewMember.PhoneType = PhoneType
     
     If Not PhoneType = 0 Then
        mCol.Add objNewMember, CStr(PhoneType)
      Else
        mCol.Add objNewMember
      End If
    
    Set Add = objNewMember
    Set objNewMember = Nothing
    Exit Function
    
CheckForDuplicate:
    '*-- Check to see if record of this type is already there, if so remove it.
    If (Err.Number = 457) And (Not bAlreadyDidRetry) Then  '*-- Key already in use.
        mCol.Remove CStr(PhoneType)
        bAlreadyDidRetry = True         '*-- Only do this once
        Resume                          '*-- Go back and try again
    End If
    Exit Function
    Resume

End Function

Public Property Get Item(vntIndexKey As Variant) As GEMS_Phone   ', Optional Cnt As Long) As GEMS_Phone
Attribute Item.VB_UserMemId = 0
  'Cnt = mCol.Count
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

