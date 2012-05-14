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
Attribute VB_Name = "Immunizations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"Immunization"
Attribute VB_Ext_KEY = "Member0" ,"Immunization"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable to hold collection
Private mCol As Collection

Public Function Add(Optional sFMP As String, Optional sSSN As String, Optional sKey As String) As Immunization ', Optional sRouteNCID As String, Optional sSiteNCID As String, Optional fDose As Single, Optional sResult As String, Optional dNextDueDate As Date, Optional sStatusName As String, Optional dAdminDate As Date, Optional sEventID As String, Optional sManufacturer As String, Optional sLot As String, Optional dExpirationDate As Date, Optional sAdminLocationName As String, Optional sSeqNumberName As String, Optional sVaccineNCID As String, Optional sUnitsNCID As String, Optional sClinicianSSN As String, Optional sComment As String, Optional sICD As String, Optional lCPT As Long, Optional sKey As String) As Immunization
    'create a new object
    Dim objNewMember As Immunization
    Set objNewMember = New Immunization


    'set the properties passed into the method
    objNewMember.sFMP = sFMP
    objNewMember.sSSN = sSSN
'    objNewMember.sRouteNCID = sRouteNCID
'    objNewMember.sSiteNCID = sSiteNCID
'    objNewMember.fDose = fDose
'    objNewMember.sResult = sResult
'    objNewMember.dNextDueDate = dNextDueDate
'    objNewMember.sStatusNCID = sStatusNCID
'    objNewMember.sStatusName = sStatusName
'    objNewMember.dAdminDate = dAdminDate
'    objNewMember.sEventID = sEventID
'    objNewMember.sManufacturerName = sManufacturerName'
'    objNewMember.sLot = sLot
'    objNewMember.dExpirationDate = dExpirationDate
'    objNewMember.sAdminLocationName = sAdminLocationName
'    objNewMember.sSeqNumberName = sSeqNumberName
'    objNewMember.sSeqNumberNCID = sSeqNumberNCID
'    objNewMember.sVaccineNCID = sVaccineNCID
'    objNewMember.sUnitsNCID = sUnitsNCID
'    objNewMember.sClinicianSSN = sClinicianSSN
'    objNewMember.sComment = sComment
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As Immunization
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

