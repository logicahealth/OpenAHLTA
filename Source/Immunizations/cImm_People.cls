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
Attribute VB_Name = "cImm_People"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"Imm_People"
Attribute VB_Ext_KEY = "Member0" ,"Imm_People"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mCol As Collection

Public Function Add(Key As String, SPON_SSAN_FMP As String, PT_SSAN As String, PT_LAST_NAME As String, PT_FIRST_NAME As String, PT_MIDDLE_NAME As String, PT_SUFFIX As String, PT_RANK As String, PT_DOB As Date, PT_GENDER As String, PT_STATUS As String, PT_SERVICE As String, PT_UNIT As String, Pt_OFFICE_SYMBOL As String, DataSource As String, LastEdited As Date, Current As Boolean, DEERS_DupID As String, DEERS_Date As Date, ExemptionExpirationDate As Date, ExemptionCode As Integer, PersonSelected As Boolean, Comments As String, Optional sKey As String) As Imm_People
  Dim objNewMember As Imm_People
  Set objNewMember = New Imm_People


  objNewMember.Key = Key
  objNewMember.SPON_SSAN_FMP = SPON_SSAN_FMP
  objNewMember.PT_SSAN = PT_SSAN
  objNewMember.PT_LAST_NAME = PT_LAST_NAME
  objNewMember.PT_FIRST_NAME = PT_FIRST_NAME
  objNewMember.PT_MIDDLE_NAME = PT_MIDDLE_NAME
  objNewMember.PT_SUFFIX = PT_SUFFIX
  objNewMember.PT_RANK = PT_RANK
  objNewMember.PT_DOB = PT_DOB
  objNewMember.PT_GENDER = PT_GENDER
  objNewMember.PT_STATUS = PT_STATUS
  objNewMember.PT_SERVICE = PT_SERVICE
  objNewMember.PT_UNIT = PT_UNIT
  objNewMember.Pt_OFFICE_SYMBOL = Pt_OFFICE_SYMBOL
  objNewMember.DataSource = DataSource
  objNewMember.LastEdited = LastEdited
  objNewMember.Current = Current
  objNewMember.DEERS_DupID = DEERS_DupID
  objNewMember.DEERS_Date = DEERS_Date
  objNewMember.ExemptionExpirationDate = ExemptionExpirationDate
  objNewMember.ExemptionCode = ExemptionCode
  objNewMember.PersonSelected = PersonSelected
  objNewMember.Comments = Comments
  If Len(sKey) = 0 Then
      mCol.Add objNewMember
  Else
      mCol.Add objNewMember, sKey
  End If


  Set Add = objNewMember
  Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As Imm_People
Attribute Item.VB_UserMemId = 0
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
  Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
  mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
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

