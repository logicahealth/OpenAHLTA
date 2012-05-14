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
Attribute VB_Name = "cWards"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"cWard"
Option Explicit

Private mcol        As Collection           '-- of cwards

Public Function AddWard(objNewMember As cWard) As Boolean
On Error GoTo ErrHandler
    mcol.Add objNewMember, CStr(objNewMember.WardID)
    AddWard = True
    Exit Function
ErrHandler:
    AddWard = False
End Function

Public Property Get Item(vntIndexKey As Variant) As cWard
Attribute Item.VB_UserMemId = 0
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

Public Function LoadWards() As Boolean
    Set mcol = New Collection
    Dim oRSWards As ADODB.Recordset
    Dim orsbeds As ADODB.Recordset
    Dim oWard As cWard
    Dim oBed As cBed
    Dim BedNum As Integer
    
    Dim oDal As cDAL
    Set oDal = New cDAL
    
    Set oRSWards = oDal.GetPADData(edtp_Ward_facility, gobjLogon.FacilityNCID)
    Do While Not oRSWards.EOF
        '__for each ward
        Set oWard = New cWard
        Call oWard.DeserializeFromRecordset(oRSWards)
        AddWard oWard
        oRSWards.MoveNext
    Loop
    
    '__ now go back and get all the beds
    For Each oWard In mcol
        Set orsbeds = oDal.GetPADData(edtp_Ward_Bed, oWard.WardID)
        
        Do While Not orsbeds.EOF
            Set oBed = New cBed
            Call oBed.DeserializedFromRecordset(orsbeds)
            Set oBed.Ward = oWard
            oWard.Beds.AddBed oBed
            orsbeds.MoveNext
        Loop
    Next oWard
    
    LoadWards = True
    
    Exit Function
    Resume
End Function
 
 
Public Function SaveWard(oWard As cWard) As Boolean
    On Error GoTo ErrHandler
    Dim oDal As cDAL
    Set oDal = New cDAL
    Dim oRSWards As ADODB.Recordset
    Dim orsbeds As ADODB.Recordset
   
    Dim oBed As cBed
    Dim BedNum As Integer
    
    '__ validate info
    If oWard.LocationNCID = vbNullString Or oWard.Specialty = vbNullString Or oWard.Beds.Count = 0 Then
        MsgBxARMd "The specialty is not sufficiently described", vbInformation, "Unable to Save Setup"
        Exit Function
    End If
    
    
    Set oRSWards = oDal.GetPADData(edtp_Ward_One, oWard.WardID)
    
    If oRSWards.EOF Then
        oRSWards.AddNew
        oWard.WardID = oDal.GetNextID(esMAINSEQ)
    End If
    
    If oWard.FacilityNCID = vbNullString Then
        oWard.FacilityNCID = gobjLogon.FacilityNCID
    End If
    
    Call oWard.SerializeToRecordset(oRSWards)
    oRSWards.Update
    
    Set orsbeds = oDal.GetPADData(edtp_Ward_Bed, oWard.WardID)
    For Each oBed In oWard.Beds
        orsbeds.Filter = "bed_id = " & oBed.BedNumber
        If orsbeds.EOF Then
            orsbeds.AddNew
        End If
        If oBed.Ward Is Nothing Then
            Set oBed.Ward = oWard
        End If
        Call oBed.SerializeToRecordset(orsbeds)
        orsbeds.Update
    Next oBed
    
    oWard.objectstate = osClean
    SaveWard = True

Exit Function
ErrHandler:
    MsgBxARMd "Application Error: " & Err.Description & " Occurred in CHCSII_PAD.cWards.SaveWard."
    Exit Function
    Resume
End Function


Public Function GetWard(ByVal SpecialtyNCID As String, ByVal LocationNCID As String) As cWard
    Dim oWard As cWard
    On Error Resume Next
    For Each oWard In mcol
        If oWard.Specialty = SpecialtyNCID And oWard.LocationNCID = LocationNCID Then
            Set GetWard = oWard
            Exit Function
        End If
    Next oWard
End Function

 Public Function RecomputeCapacity(Optional WardID As Long = 0) As Boolean
    Dim oWard As cWard
    On Error GoTo ErrHandler
    If WardID <> 0 Then  '64119
        Set oWard = mcol(CStr(WardID))
        If Not oWard Is Nothing Then
            RecomputeCapacity = oWard.ComputeCapacity
        Else
            RecomputeCapacity = False
        End If
    Else
        '__ do them all
        For Each oWard In mcol
            oWard.ComputeCapacity
        Next oWard
        RecomputeCapacity = True
    End If
    Exit Function
ErrHandler:
MsgBox "Application Error: " & Err.Description & " Occurred in PAD.cWards.RecomputeCapacity"

 End Function
 
 
 
'__ this is currently an orphaned process
'Public Function LoadpatientsIntoWards(oInpatients As cInPatients)
'    Dim oWard As cWard
'    Dim oBed As cBed
'
'    For Each oWard In mCol
'         'oward.FacilityNCID =
'        For Each oBed In oWard.Beds
'            Set oBed.Patient = oInpatients.getInPatientByBed(oBed)
'        Next oBed
'    Next oWard
' End Function
