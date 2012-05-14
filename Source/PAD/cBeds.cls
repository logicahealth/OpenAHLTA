VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBeds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"cBed"
Option Explicit

Private mcol        As Collection   '__ of cBed
Private mWard       As cWard
Private mvarcBed    As cBed

Public Property Set Ward(RHV As cWard): Set mWard = RHV: End Property

Public Function AddBed(Bed As cBed) As Boolean
'    Set Bed.Ward = mWard
    mcol.Add Bed, Bed.BedNumber
End Function

Public Function RemoveBed(Bed As cBed) As Boolean
    If Bed.Patient Is Nothing _
    Or Bed.Status = "INACTIVE" Then
        mcol.Remove Bed.BedNumber
        RemoveBed = True
        Exit Function
    ElseIf Val(Bed.Patient.Unit_Number) = 0 Then
        mcol.Remove Bed.BedNumber
        RemoveBed = True
        Exit Function
    Else
        '__ we have a patient so this is not allowed
    End If
End Function


Public Property Get Item(vntIndexKey As Variant) As cBed
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

