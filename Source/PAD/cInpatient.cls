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
Attribute VB_Name = "cInpatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit


Private mvarAdmissionDate           As Date
Private mvarAdmissionType           As String
Private mvarAdmittingProviderName   As String
Private mvarAdmittingProviderNCID   As String
Private mvarAttendingProviderName   As String
Private mvarAttendingProviderNCID   As String
Private mvarBed                     As cBed
Private mvarcomments                As String
Private mcolDiagnoses               As Collection 'of cdiagnosis
Private mvarDirty                   As Boolean      '-- has the data changed?

'This date means that the patient is discharged. It is saved in the inpatient_admission table.
Private mvarDischargedate           As Date

'Disposition updates this date and it is read from the frmConfirmDischarge form, this does not mean the patient is discharged.
'The date is saved in the inpatient_discharge table
Private mvarPotentialDischargeDate As Date

Private mvarDischargeDiagnosis      As String
Private mvarDisposition             As String
Private mvarID                      As String
Private mvarLastTransferDate        As Date
Private mvarObjectState             As ObjectState_
Private mvarPrimaryServiceNCID               As String
Private mcolTransfers               As Collection   '-- of ctransfer
Private mvarUnit_Number             As String
Private mvarWardID                  As String
Private mvarTimeBias                As Long
'__ these are read-only
Private mvarPatientName             As String
Private mvarBranch                  As String
Private mvarUnit                    As String
Private mvarFmpSsn                  As String
Private mvarRank                    As String
Private mvarStatus                  As String   '-- one of 'Inpatient' or 'Discharged'
Private mvarApptID                  As String
Private mvarAssignedBedNumber       As String
Private mvarPostDischargeNote       As String
Private mvarDischargeNote           As String
Private mvarAdmittedFromEncNum      As String

Public CellFontBold            As Boolean  '-- bold for inpatients and not-bold for discharged patients
Public CellForeColor           As Long     '-- vbblack if all data there, vbred if missing some data


Public Property Let AdmittedFromEncNum(vData As String)
    
    mvarAdmittedFromEncNum = vData
    
End Property

Public Property Get AdmittedFromEncNum() As String
    
    AdmittedFromEncNum = mvarAdmittedFromEncNum
    
End Property

Public Property Get DischargeNote() As String:                  DischargeNote = mvarDischargeNote:              End Property
Public Property Let DischargeNote(ByVal vData As String):       mvarDischargeNote = vData:                      End Property
Public Property Let ID(ByVal vData As String)

    mvarID = vData

End Property

Public Property Get PotentialDischargeDate() As Date
    
    PotentialDischargeDate = mvarPotentialDischargeDate
    
End Property

Public Property Let PotentialDischargeDate(vData As Date)
    
    mvarPotentialDischargeDate = vData
    
End Property

Public Property Get TimeBias() As Long:                         TimeBias = mvarTimeBias:                        End Property
Public Property Let TimeBias(ByVal vData As Long):              mvarTimeBias = vData:                           End Property
Public Property Get PostDischargeNote() As String:              PostDischargeNote = mvarPostDischargeNote:      End Property
Public Property Let PostDischargeNote(ByVal vData As String):   mvarPostDischargeNote = vData:                  End Property
Public Property Get AssignedBedNumber() As String:              AssignedBedNumber = mvarAssignedBedNumber:      End Property
Public Property Let AssignedBedNumber(RHV As String):           mvarAssignedBedNumber = RHV:                    End Property
Public Property Get AppointmentID() As String:                  AppointmentID = mvarApptID:                     End Property
Public Property Let AppointmentID(ByVal vData As String):       mvarApptID = vData:                             End Property
Public Property Get AdmissionDate() As Date:                    AdmissionDate = mvarAdmissionDate:              End Property
Public Property Let AdmissionDate(ByVal vData As Date):         mvarAdmissionDate = vData:                      End Property
Public Property Get AdmissionType() As String:                  AdmissionType = mvarAdmissionType:              End Property
Public Property Let AdmissionType(ByVal vData As String):       mvarAdmissionType = vData:                      End Property
Public Property Get AdmittingProviderName() As String:          AdmittingProviderName = mvarAdmittingProviderName:  End Property
Public Property Let AdmittingProviderName(ByVal RHV As String): mvarAdmittingProviderName = RHV:                    End Property
Public Property Get AdmittingProviderNCID() As String:          AdmittingProviderNCID = mvarAdmittingProviderNCID:  End Property
Public Property Get AttendingProviderName() As String:          AttendingProviderName = mvarAttendingProviderName:  End Property
Public Property Let AttendingProviderName(ByVal RHV As String): mvarAttendingProviderName = RHV:                    End Property
Public Property Get AttendingProviderNCID() As String:          AttendingProviderNCID = mvarAttendingProviderNCID:  End Property
Public Property Get comments() As String:                       comments = mvarcomments:                            End Property
Public Property Let comments(ByVal RHV As String):              mvarcomments = RHV:                                 End Property
Public Property Get Bed() As cBed:                          Set Bed = mvarBed:                                  End Property
Public Property Get Diagnoses() As Collection:              Set Diagnoses = mcolDiagnoses:                      End Property
Public Property Set Diagnoses(RHV As Collection):               Set mcolDiagnoses = RHV:                        End Property
Public Property Get DischargeDiagnosis() As String:             DischargeDiagnosis = mvarDischargeDiagnosis:    End Property
Public Property Let DischargeDiagnosis(RHV As String):          mvarDischargeDiagnosis = RHV:                   End Property

Public Property Get Disposition() As String:                    Disposition = mvarDisposition:              End Property
Public Property Let Disposition(RHV As String):                 mvarDisposition = RHV:                      End Property
Public Property Let DispositionNCID(RHV As String)
    mvarDisposition = Disposition_NCID_to_Text(RHV)
End Property

Public Property Get ID() As String
    ID = mvarID
End Property
Public Property Get objectstate() As ObjectState_:              objectstate = mvarObjectState:                          End Property
Public Property Let objectstate(RHV As ObjectState_):           mvarObjectState = RHV:                                  End Property
Public Property Get Unit_Number() As String:                    Unit_Number = mvarUnit_Number:                      End Property
Public Property Get PatientName() As String:                    PatientName = mvarPatientName:                      End Property
Public Property Get Branch() As String:                         Branch = mvarBranch:              End Property
Public Property Get PrimaryServiceNCID() As String:             PrimaryServiceNCID = mvarPrimaryServiceNCID:          End Property
Public Property Let PrimaryServiceNCID(RHV As String):          mvarPrimaryServiceNCID = RHV:                End Property
Public Property Get Status() As String:                         Status = mvarStatus:                End Property
Public Property Get Unit() As String:                           Unit = mvarUnit:                      End Property
Public Property Get FMPSSN() As String:                         FMPSSN = mvarFmpSsn:                  End Property
Public Property Get Rank() As String:                           Rank = mvarRank:                      End Property
Public Property Get Ward_ID() As String:                        Ward_ID = mvarWardID:          End Property
Public Property Let Ward_ID(RHV As String):                     mvarWardID = RHV:                End Property

Public Property Let AdmittingProviderNCID(ByVal RHV As String)
    Dim oProv As CHCSII_ClinicClient.Provider
    mvarAdmittingProviderNCID = RHV
    If RHV <> vbNullString Then
        For Each oProv In gColProviders
            If oProv.ClinicianNCID = mvarAdmittingProviderNCID Then
                mvarAdmittingProviderName = oProv.ProviderName
                Exit For
            End If
        Next oProv
    Else
        mvarAdmittingProviderName = vbNullString
    End If
End Property

Public Property Let AttendingProviderNCID(ByVal RHV As String)
    Dim oProv As CHCSII_ClinicClient.Provider
    mvarAttendingProviderNCID = RHV
    If RHV <> vbNullString Then
        For Each oProv In gColProviders
            If oProv.ClinicianNCID = mvarAttendingProviderNCID Then
                mvarAttendingProviderName = oProv.ProviderName
                Exit For
            End If
        Next oProv
    Else
        mvarAttendingProviderName = vbNullString
    End If
End Property
Public Property Set Bed(ByVal vData As cBed)
    Set mvarBed = vData
    mvarAssignedBedNumber = mvarBed.BedNumber
    
    'SF 7-5-05 SCR 72841 Set the Ward_ID
    mvarWardID = mvarBed.Ward.WardID
End Property

Public Property Get Diagnosis() As String
    If mcolDiagnoses.Count > 0 Then
        Diagnosis = mcolDiagnoses.Item(1).Description
    Else
        Diagnosis = "(None Recorded)"
    End If
End Property
Public Property Get DischargeDate() As Date:                    DischargeDate = mvarDischargedate:                  End Property

Public Property Let DischargeDate(ByVal RHV As Date)
    mvarDischargedate = RHV
    If CDbl(RHV) > 0 Then
        mvarStatus = "DISCHARGED"
        CellFontBold = False
    Else
        mvarStatus = "INPATIENT"
        CellFontBold = True
    End If
    
End Property

Public Property Get LastTransferDate() As String
    LastTransferDate = mvarLastTransferDate
End Property
Public Property Let LastTransferDate(RHV As String)
    mvarLastTransferDate = RHV
End Property
Public Property Let Unit_Number(ByVal vData As String)
    If mvarUnit_Number = vData Then Exit Property
    mvarUnit_Number = vData
    Call GetPatientData(mvarUnit_Number, mvarPatientName, mvarBranch, mvarUnit, mvarFmpSsn, mvarRank)
End Property

Private Sub Class_Initialize()
    Set mcolDiagnoses = New Collection
    Set mvarBed = New cBed
    mvarStatus = "INPATIENT"
    CellFontBold = True
    CellForeColor = vbRed
    Set mcolTransfers = New Collection
End Sub

Private Sub Class_Terminate()
    Set mcolDiagnoses = Nothing
    Set mvarBed = Nothing
    Set mcolTransfers = Nothing
End Sub

Public Function AddDiagnoses(RHV As cDiagnosis) As Boolean
'<< scr 49182 Only allow a single instance of a medcin id (inpatient_diagnosis Primary Key is inpatient ID and Medcin ID)
    Dim oDiag As cDiagnosis
    For Each oDiag In mcolDiagnoses
        If RHV.SnoID = oDiag.SnoID Then
            '__ don't add it
            Exit Function
        End If
    Next oDiag
    mcolDiagnoses.Add RHV
    AddDiagnoses = True
End Function

Public Function DeserializeFromRecordset(oRS As ADODB.Recordset)

On Error GoTo ErrHnd
    
    If oRS.EOF Then Exit Function
    Dim oWard As cWard
    Dim oBed As cBed
    Dim oProv As CHCSII_ClinicClient.Provider
    Dim oTr As cTransfer
    
    mvarID = vbNullString & oRS("INPATIENT_ID")
    
    mvarUnit_Number = vbNullString & oRS("UNIT_NUMBER")
    mvarAdmissionType = vbNullString & oRS("ADMISSION_TYPE_CODE")
    mvarAdmissionDate = vbNullString & oRS("ADMISSION_DATE")
    mvarApptID = vbNullString & oRS("Appt_ID")
    Dim DateStr As String
    DateStr = vbNullString & oRS("DISCHARGE_DATE")
    If IsDate(DateStr) Then
        DischargeDate = CDate(DateStr)   '__ use public var to fire the side effects
    Else
        DischargeDate = 0
    End If

    AdmittingProviderNCID = vbNullString & oRS("ADMITTING_PROVIDER_NCID")
    AttendingProviderNCID = vbNullString & oRS("ATTENDING_PROVIDER_NCID")
    mvarPrimaryServiceNCID = vbNullString & oRS("PRIMARY_SERVICE_NCID")
    mvarcomments = vbNullString & oRS("COMMENTS_TEXT")
    mvarTimeBias = Val(vbNullString & oRS("TimeBias"))
    mvarAdmittedFromEncNum = oRS("ADMITTED_FROM_ENC_NUM") & ""
    
    '__ go through the wards and find this guy's bed.  then assign him to it & it to him
    '__ depending on state, the wards may not have been loaded yet
    '   but that's ok, we'll do this when we load the wards, too
    '-- also, the ward may no longer exist for the current setup
    '
    mvarWardID = vbNullString & oRS("WARD_NCID")
    mvarAssignedBedNumber = vbNullString & oRS("ASSIGNED_BED_NUMBER")
    '-- only assign to bed if not discharged
    
    If mvarStatus = "INPATIENT" And Not gWards Is Nothing Then
        For Each oWard In gWards
            If oWard.WardID = mvarWardID Then
                For Each oBed In oWard.Beds
                    If oBed.BedNumber = mvarAssignedBedNumber Then
                        Set mvarBed = oBed
                        Set oBed.Patient = Me
                        Exit For
                    End If
                Next oBed
            End If
        Next oWard
    End If
    '__ get the rest of his data
    Call GetPatientData(mvarUnit_Number, mvarPatientName, mvarBranch, mvarUnit, mvarFmpSsn, mvarRank)
    
    Dim oDiag As cDiagnosis
    Dim oDal As cDAL
    Set oDal = New cDAL
    Dim oRSD As ADODB.Recordset
    
    Set oRSD = oDal.GetInpatientData(edti_Diagnoses, mvarID, "ADMISSION")
    Do While Not oRSD.EOF
        Set oDiag = New cDiagnosis
        Call oDiag.DeserializeFromRecordset(oRSD)
        mcolDiagnoses.Add oDiag
        oRSD.MoveNext
    Loop
    '__ get the transfers for this guy
    oRSD.Close
    Set oRSD = oDal.GetInpatientData(edti_Transfer, mvarID)
    Do While Not oRSD.EOF
        Set oTr = New cTransfer
        Call oTr.DeserializeFromRecordset(oRSD)
        mcolTransfers.Add oTr
        oRSD.MoveNext
    Loop
    oRSD.Close
    If mcolTransfers.Count > 0 Then
        'SCR 101232 change from 1 to mcolTransfers.Count
        LastTransferDate = mcolTransfers(mcolTransfers.Count).TransDate
    End If
    
'-- get discharge comments
'    If CDbl(mvarDischargedate) > 0 Then
        Set oRSD = oDal.GetInpatientData(edti_DischargeData, mvarID)
        If Not oRSD.EOF Then
            If Not IsNull(oRSD("DischargeDate")) Then
                mvarPotentialDischargeDate = oRSD("DischargeDate")
            End If
            
            If CDbl(mvarDischargedate) > 0 Then
                mvarPostDischargeNote = vbNullString & oRSD("postdischargetext")
                mvarDischargeNote = vbNullString & oRSD("dischargetext")
                mvarDischargeDiagnosis = vbNullString & oRSD("discharge_Diagnosis")
                DispositionNCID = vbNullString & oRSD("dispositionncid")
            End If
        End If
        oRSD.Close
'    End If
    
    mvarDirty = False
 
Exit Function

ErrHnd:
    MsgBxARMd "Application Error: " & Err.Description & " Occurred in CHCSII_PAD.cInpatient.DeserializeFromRecordset."
    Exit Function
    Resume
 End Function


Public Function SerializeToRecordset(oRS As ADODB.Recordset)
    On Error GoTo ErrHandler
    
    oRS("INPATIENT_ID") = mvarID
    oRS("UNIT_NUMBER") = mvarUnit_Number
    oRS("ADMISSION_TYPE_CODE") = mvarAdmissionType
    oRS("ADMISSION_DATE") = mvarAdmissionDate
    If Not CDbl(mvarDischargedate) = 0 Then
        oRS("DISCHARGE_DATE") = mvarDischargedate
    End If
    oRS("TimeBias") = mvarTimeBias
    oRS("COMMENTS_TEXT") = mvarcomments
'<< scr 45073 attend/admittin provider ncids are now comming back as -1, so change the <> 0 to > 0
    If Val(mvarAdmittingProviderNCID) > 0 Then
        oRS("ADMITTING_PROVIDER_NCID") = mvarAdmittingProviderNCID
    End If
    If Val(mvarAttendingProviderNCID) > 0 Then
        oRS("ATTENDING_PROVIDER_NCID") = mvarAttendingProviderNCID
    End If
    If Not mvarBed Is Nothing Then
        If Not Val(mvarBed.BedNumber) = 0 Then
            oRS("ASSIGNED_BED_NUMBER") = mvarBed.BedNumber
        Else
            oRS("ASSIGNED_BED_NUMBER") = Null
        End If
        If Not mvarBed.Ward Is Nothing Then
                                            'SF 3/10/05 SCR 67239
            If Val(mvarBed.BedNumber) = 0 And mvarWardID <> "" Then
                oRS("WARD_NCID") = mvarWardID
            Else
                oRS("WARD_NCID") = mvarBed.Ward.WardID
            End If
        End If
    End If
    oRS("facility_ncid") = gobjLogon.FacilityNCID
    If mvarPrimaryServiceNCID <> vbNullString Then
        oRS("PRIMARY_SERVICE_NCID") = mvarPrimaryServiceNCID
    End If
    'appt id can not be null.  so just to be sure, make it at least a zero
    If Val(mvarApptID) = 0 Then
        mvarApptID = "0"
    End If
    oRS("Appt_ID") = mvarApptID
    'RQT 61021
    oRS("ADMITTED_FROM_ENC_NUM") = Val(mvarAdmittedFromEncNum)
    
    Exit Function
    
ErrHandler:
    MsgBxARMd "Application Error: " & Err.Description & " Occurred in CHCSII_PAD.cInpatient.SerializeToRecordset."
    Exit Function
    Resume
End Function

