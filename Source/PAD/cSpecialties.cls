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
Attribute VB_Name = "cSpecialties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"cSpecialty"
Attribute VB_Ext_KEY = "Member0" ,"cSpecialty"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mcol As Collection

Public Function AddSpecialty(ByVal Spec As cSpecialty) As Boolean
    On Error GoTo ErrHandler

    mcol.Add Spec, Spec.SpecName
    AddSpecialty = True
    Exit Function
ErrHandler:
    AddSpecialty = False
End Function
Public Function GetSpecialty(ByVal SpecNCID As String) As cSpecialty
    For Each GetSpecialty In mcol
        If GetSpecialty.SpecNCID = SpecNCID Then
            Exit Function
        End If
    Next GetSpecialty
End Function
Public Property Get Item(vntIndexKey As Variant) As cSpecialty
Attribute Item.VB_UserMemId = 0
    If IsNumeric(vntIndexKey) Then
        vntIndexKey = CStr(vntIndexKey)
    End If
  Set Item = mcol(vntIndexKey)
End Property
Public Property Get Count() As Long
    Count = mcol.Count
End Property
Public Sub Remove(vntIndexKey As Variant)
    mcol.Remove vntIndexKey
End Sub
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mcol.[_NewEnum]
End Property
Private Sub Class_Initialize()
    Set mcol = New Collection
End Sub
Private Sub Class_Terminate()
    Set mcol = Nothing
End Sub

Public Function LoadSpecialties()
    Dim oSpec As cSpecialty
    Dim oRS As ADODB.Recordset
    Dim oDAL As cDAL
    Set oDAL = New cDAL
    Set mcol = New Collection
    
    Set oRS = oDAL.GetPADData(edtp_Specialty)
    
    On Error Resume Next
    Do While Not oRS.EOF
        Set oSpec = New cSpecialty
        oSpec.SpecNCID = vbNullString & oRS("NCID")
        oSpec.SpecName = vbNullString & oRS("description")
        oRS.MoveNext
        mcol.Add oSpec, oSpec.SpecNCID
    Loop

End Function
