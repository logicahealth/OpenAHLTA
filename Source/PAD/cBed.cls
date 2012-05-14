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
Attribute VB_Name = "cBed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Private mvarWard        As cWard
Private mvarBedNumber   As String
Private mvarPatient     As cInpatient
Private mvarObjectState As ObjectState_
Private mvarStatus      As String

Public Property Let objectstate(RHV As ObjectState_):       mvarObjectState = RHV:          End Property
Public Property Get objectstate() As ObjectState_:          objectstate = mvarObjectState:  End Property
Public Property Set Patient(ByVal vData As cInpatient): Set mvarPatient = vData:            End Property
Public Property Get Patient() As cInpatient:            Set Patient = mvarPatient:          End Property
Public Property Let BedNumber(ByVal vData As String):       mvarBedNumber = vData:          End Property
Public Property Get BedNumber() As String:                  BedNumber = mvarBedNumber:      End Property
Public Property Get Status() As String:                     Status = mvarStatus:            End Property
Public Property Let Status(RHV As String):                  mvarStatus = RHV:               End Property
Public Property Set Ward(ByVal vData As cWard):         Set mvarWard = vData:               End Property
Public Property Get Ward() As cWard:                    Set Ward = mvarWard:                End Property
Private Sub Class_Initialize()
'__ do NOT init the wards and pateints here, that causes a cyclic init process.
    Set mvarWard = New cWard
    mvarWard.WardID = -1
End Sub

Public Function DeserializedFromRecordset(oRS As ADODB.Recordset) As Boolean
   
    Dim i As Long
    Dim KVP As KeyValuePair
    If oRS.EOF Then Exit Function
    
    mvarBedNumber = vbNullString & oRS("Bed_id")
    Set mvarWard = gWards(CStr(oRS("ward_ncid")))
    mvarStatus = vbNullString & oRS("bed_status")
    
'    Debug.Assert CBool(CStr(oRS("ward_ncid")) <> "100021563")
    '-- it's ok, we'll get this when we load the patients
    If Not gInPatients Is Nothing Then
        If Val(vbNullString & oRS("inpatient_id")) > 0 Then
            '__get patient
            Set mvarPatient = gInPatients(vbNullString & oRS("inpatient_id"))
        End If
    End If
End Function

Public Function SerializeToRecordset(oRS As ADODB.Recordset) As Boolean

    On Error GoTo ErrHnd
    
    Dim oBed As cBed
    
    If oRS.EOF Then Exit Function

    oRS("WARD_NCID") = mvarWard.WardID
    oRS("BED_ID") = mvarBedNumber
    If Not mvarPatient Is Nothing Then
        oRS("INPATIENT_ID") = mvarPatient.ID
    Else 'SCR 72841 The inpatient was not being cleared from the table when released from the bed.
        oRS("INPATIENT_ID") = Null
    End If
    oRS("BED_STATUS") = mvarStatus
    oRS("FACILITY_NCID") = mvarWard.FacilityNCID
    
Exit Function

ErrHnd:
    MsgBxARMd "Application Error: " & Err.Description & " Occurred in CHCSII_PAD.cBed.SerializeToRecordset."
    
End Function


Public Function CloneBed() As cBed
    Set CloneBed = New cBed
    CloneBed.BedNumber = mvarBedNumber
    CloneBed.objectstate = mvarObjectState
    Set CloneBed.Patient = mvarPatient
    CloneBed.Status = mvarStatus
    Set CloneBed.Ward = mvarWard
End Function
