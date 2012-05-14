VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cInPatients"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"cInpatient"
Option Explicit


'local variable to hold collection
Private mcol            As Collection
Private mvarcInpatient  As cInpatient '-- the 'current' inpatient object


Public Function GetInpatientFromDB(InptAdmissionID As String, FacilityNCID As String)
    
    'This loads an inpatient object from the DB
    'Called from the EncounterPrevious.Documents class to get discharged patients.
    
On Error GoTo ErrHnd
    
    Dim oDal As cDAL
    Dim oRS As Recordset
    Dim oPat As cInpatient
    
    'Don't load the patient from the db if the patient already exists in the collection.
    'SCR 67667, 68131
    Set GetInpatientFromDB = getInPatientByID(InptAdmissionID)
    If Not GetInpatientFromDB Is Nothing Then
        Exit Function
    End If
    
    Set oDal = New cDAL
    Set oRS = oDal.GetInpatientData(edti_Patient_one, InptAdmissionID, FacilityNCID)
    If Not oRS.EOF Then
        Set oPat = New cInpatient
        oPat.DeserializeFromRecordset oRS
        Set GetInpatientFromDB = oPat
    End If
        
    'SCR 67621 Add the inpatient to the internal collection so it can be set as the selected inpatient.
    mcol.Add oPat, CStr(oPat.ID)
    
    Exit Function
    
ErrHnd:
    If Err.Number = 457 Then
        'SF 4-4-05 SCR 68494 The patient is already added
        Resume Next
    Else
        Err.Raise Err.Number, Err.Source, "Error in CHCSII_PAD.cInPatients.GetInpatientFromDB: " & Err.Description
    End If
    
    Exit Function
    Resume
End Function


Public Property Get Inpatient() As cInpatient
    If mvarcInpatient Is Nothing Then
        Set mvarcInpatient = New cInpatient
    End If

    Set Inpatient = mvarcInpatient
End Property

Public Property Set Inpatient(vData As cInpatient)
    '-- only set this if the object is in the collection
    '  if this is a 'selected pat list' object the vdata may not be in the collection
    '  and the same is true of the 'global inpatient list'
    
    'SCR 67541
    If vData Is Nothing Then
        Set mvarcInpatient = vData
        Exit Property
    End If
    
    'If vData Is Nothing Then Exit Property
    
    Dim oPat As cInpatient
    
    'RQT 62107 When the patient is cleared, it wasn't clearing the selected outpatient
    If vData.ID = "" Then
        Set mvarcInpatient = vData
        Exit Property
    End If
    
    For Each oPat In mcol
        If oPat.ID = vData.ID Then
            Set mvarcInpatient = oPat
            Exit For
        End If
    Next oPat
    
End Property


Public Function AddPatient(opatient As cInpatient) As Boolean
    On Error GoTo ErrHandler
    
    'SF 6-24-05 Performance opt.
    If opatient.ID <> Empty Then
        mcol.Add opatient, CStr(opatient.ID)
    Else
        mcol.Add opatient
    End If
    AddPatient = True
    Exit Function
    
ErrHandler:
    MsgBxARMd "Application Error: " & Err.Description & " Occurred in CHCSII_PAD.cInpatients.AddPatient. Could not add a patient."

End Function

Public Function getInPatientByUnitNumber(ByVal UnitNumber As String) As cInpatient
    Dim opatient As cInpatient
    For Each opatient In mcol
        If opatient.Unit_Number = UnitNumber Then
            Set getInPatientByUnitNumber = opatient
            Exit Function
        End If
    Next opatient
    
    'Not used yet.
'    'If the patient is not found, the load it from the DB
'    Set getInPatientByUnitNumber = GetInpatientFromDB(UnitNumber, gobjLogon.FacilityNCID)
    
End Function

Public Function getInPatientByID(ByVal ID As Long) As cInpatient
    Dim opatient As cInpatient
    For Each opatient In mcol
        If opatient.ID = ID Then
            Set getInPatientByID = opatient
            Exit Function
        End If
    Next opatient
End Function


Public Function getInPatientByBed(Bed As cBed) As cInpatient
    Dim opatient As cInpatient
    For Each opatient In mcol
        If opatient.Bed.BedNumber = Bed.BedNumber _
        And opatient.Bed.Ward = Bed.Ward Then
            Set getInPatientByBed = opatient
            Exit Function
        End If
    Next opatient
End Function


Public Property Get Item(vntIndexKey As Variant) As cInpatient
Attribute Item.VB_UserMemId = 0
    
    On Error GoTo ErrHnd
    If IsNumeric(vntIndexKey) Then
        vntIndexKey = CStr(vntIndexKey)
    End If
    Set Item = mcol(vntIndexKey)
    
Exit Property

ErrHnd:
    If Err.Number = 5 Then
        Set Item = Nothing
    Else
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
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
  Set mvarcInpatient = Nothing
    Set mcol = Nothing
End Sub

Public Function LoadInPatients(CurrOptions As cOptions) As Boolean

On Error GoTo ErrHnd

    Dim oDal As cDAL
    Set oDal = New cDAL
    Dim oRSPatients As ADODB.Recordset
    Dim opatient As cInpatient
    Dim tmpdate As Date '_ used to flip the dates if fromdate is greater than todate
    
    Dim oWard As cWard
    Dim oBed As cBed
    Dim BedNum As Integer
    Dim sSql As String
    Dim vItem As Variant
    
    
    
    If CurrOptions Is Nothing Then
        Set CurrOptions = New cOptions
    End If
    
    Set oRSPatients = oDal.GetPatientList(CurrOptions)
    Set mvarcInpatient = New cInpatient
    
    Set mcol = New Collection
    
    'SF 6-24-05 Performance opt.
    'Since all the inpatients "should" be in the global gInPatients collection, get it from there.
    Do Until oRSPatients.EOF
        Set opatient = gInPatients.Item(CStr(oRSPatients("Inpatient_ID")))
        If Not opatient Is Nothing Then
            AddPatient opatient
        Else
            Set opatient = New cInpatient
            Call opatient.DeserializeFromRecordset(oRSPatients)
            AddPatient opatient
        End If
        oRSPatients.MoveNext
    Loop
    
'    Do While Not oRSPatients.EOF
'        Set opatient = New cInpatient
'        Call opatient.DeserializeFromRecordset(oRSPatients)
'        AddPatient opatient
'        oRSPatients.MoveNext
'    Loop

Exit Function

ErrHnd:
    MsgBxARMd "Application Error: " & Err.Description & " Occurred in CHCSII_PAD.cInPatients.LoadInPatients."
    
End Function

Public Function Save(Optional oPat As cInpatient) As Boolean
    On Error GoTo ErrHnd
    
    If oPat Is Nothing Then
    '-- save the default patient
        If Not mvarcInpatient Is Nothing Then
            Save = SaveInpatient(mvarcInpatient)
            mvarcInpatient.objectstate = osClean
        End If
    Else
        Save = SaveInpatient(oPat)
        oPat.objectstate = osClean
    End If

Exit Function

ErrHnd:
    MsgBxARMd "Application Error: " & Err.Description & " Occurred in CHCSII_PAD.cInPatients.Save."
    
End Function

Private Function SaveInpatient(Optional oPat As cInpatient) As Boolean
    Dim oDal As cDAL
    Set oDal = New cDAL
    Dim oRS As ADODB.Recordset
    
    On Error GoTo ErrHandler
    
    If oPat.ID = 0 Then
        oPat.ID = oDal.GetNextID(esJDOE_NUMBER)
    End If
    Set oRS = oDal.GetInpatientData(edti_Patient_one, oPat.ID, gobjLogon.FacilityNCID, adOpenDynamic, adLockOptimistic)
    
    If oRS.EOF Then
        oRS.AddNew
    End If
    Call oPat.SerializeToRecordset(oRS)
    oRS.Update
    
'<< scr 46480 have to save the patient before we can add diagnoses
    
    Dim oDiag As cDiagnosis
    Dim oRSD As ADODB.Recordset
'<< scr 48606  have to include the final param
    Set oRSD = oDal.SetInpatientData(edti_Diagnoses, oPat.ID, "ADMISSION")
    
    For Each oDiag In oPat.Diagnoses
        oRSD.AddNew
        Call oDiag.SerializeToRecordset(oRSD)
        oRSD("diagnosis_type") = "ADMISSION"        '<< scr 49187,18766 this change wasn't fully integrated
        oRSD.Update
    Next oDiag
    If oRSD.State = adStateOpen Then
        oRSD.Close
    End If
    
    oRS.Close
    
    If Val(oPat.Ward_ID) <> 0 Then
        Call gWards.SaveWard(gWards(oPat.Ward_ID))
    End If
    
    SaveInpatient = True
    Exit Function
ErrHandler:
    If Err.Number = -2147217873 Then
        'ORA-00001: unique constraint (GEMS.INPATIENT_DIAGNOSIS_PK) violated
        'The provider entered multiple Mecdin NCIDS that are the same.
        'Only save the first.
        'SF 6/1/05 comment this out
        'Resume Next
    ElseIf Err.Number = 3219 Then
        'Operation is not allowed in this context.
        'SF 6/1/05 comment this out
        'Resume Next
    
    Else
        SaveInpatient = False
        gobjShared.ShowVBError Err.Number, Err.Description, "Inpatient.cls - Save Inpatient Record", "Error Saving Inpatient Record", vbExclamation
    End If
    
    Exit Function
    Resume
End Function

Public Function CreateInpatient(Patient As MMMHISPatient.Patient, _
                                SnoConvert As IConvert, _
                  Optional ByVal AdmittingProviderNCID As String, _
                  Optional ByVal AttendingProviderNCID As String, _
                  Optional ByVal AdmissionDate As Date, _
                  Optional ByVal EncounterNumber As String = vbNullString, _
                  Optional ByVal AdmittedFromEncNum As String = vbNullString) As cInpatient
              '__ facility ncid is assumed to be the logon facility.
    
    '__ this is called from outside as a 'remote' patient admission method (directly from disposition)
    '   We need to build up an inpatient record and then show it in 'edit' mode on the new admission form
    '   then save it, and if the form is up, add it to the grid
    Dim oDal As cDAL
    Dim oRS As ADODB.Recordset
    Dim oDiag As cDiagnosis
    
'build an inpatient record
    Set CreateInpatient = New cInpatient
    With CreateInpatient
    
        If Val(EncounterNumber) > 0 Then
            Set oDal = New cDAL
            Set oRS = oDal.GetInpatientData(edti_EncDiagnoses, EncounterNumber, gobjLogon.FacilityNCID)
            Do While Not oRS.EOF
                Set oDiag = New cDiagnosis
                oDiag.NCID = vbNullString & oRS("DIAGNOSISNCID")
                oDiag.SnoID = vbNullString & oRS("SnoID")
                oDiag.ComplaintType = vbNullString & oRS("NEWFOLLOWUP")
                oDiag.comment = vbNullString & oRS("DIAGNOSESCOMMENT")
                'oDiag.MedcinPrefix = vbNullString & oRS("PREFIX")
                oDiag.Description = SnoConvert.SnoDescription(oDiag.SnoID)

                .AddDiagnoses oDiag
                oRS.MoveNext
            Loop
        End If
        .AdmissionDate = gobjComm.LocalToGmt(AdmissionDate)
        .TimeBias = gobjComm.GetTotalTimeBias '105144
        .AdmissionType = "DIRECT"
        .AdmittingProviderNCID = AdmittingProviderNCID
        .AttendingProviderNCID = AttendingProviderNCID
        .objectstate = osNew
        .Unit_Number = Patient.UnitNumber
        .AdmittedFromEncNum = AdmittedFromEncNum
        
    End With

End Function

Public Function ReloadPatient(oPat As cInpatient) As Boolean

On Error GoTo ErrHandler
    Dim obj As cInpatient
    Dim oRS As ADODB.Recordset
    Dim i As Long
    Dim oDal As New cDAL
    
    If oPat Is Nothing Then
        ReloadPatient = False
        Exit Function
    End If
    '-- remove this patient from the collection and reload from db  (like after a transfer or other update in the selectedpatients object)
    For i = 1 To mcol.Count
        Set obj = mcol(i)
        If obj.ID = oPat.ID Then
            Set oRS = oDal.GetInpatientData(edti_Patient_one, oPat.ID, gobjLogon.FacilityNCID)
            If Not oRS.EOF Then
                '-- reload the info from db
                Call obj.DeserializeFromRecordset(oRS)
                ReloadPatient = True
                Exit Function
            Else
                ReloadPatient = False
                Exit Function
            End If
            Exit For
        End If
    Next i
    
    Exit Function
    
ErrHandler:
    ReloadPatient = False
End Function
