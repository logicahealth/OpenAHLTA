VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Vaccine_Rec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private msVaccine As String
Private msMinAge As String
Private msMaxAge As String
Private msVaccineID As String
Private msEffectsVacID As String
Private msHL7_ID As String
Private msNbrInSeries As String
Private msGenericName As String
Private msCPT As String
Private msICD As String
Private msBooster As String
Private msMPF_DIN As String
Private msChildHood As String
Private msInactive As String
Private msHistKey As String
Private msGenNameID As String
Private msImmRelNcid As String
Private msVIS_Version As String
Private msSchedule As String
Private SeriesData As String
Private ExpertData As String
Public Function ModVaccineData() As Boolean

  ' This Routine modifies an existing Record in the Vaccine DataBase

 On Local Error GoTo haveerror
 
  Dim xSet As String
  Dim sql As String
  Dim rs As ADODB.Recordset
  Dim NewNcid As Boolean
  Dim VacRec() As String
  Dim Count As Integer
  Dim vacidrec As String
  Dim sepchar As String
  Dim LastNcid As Long
  
  '...Do we need to Update Immunization History
    If NewHistKey <> OldHistKey Then
        FrmStatus.LabStatus = FrmStatus.LabStatus & vbCrLf & "Adding to VacKey"
        FrmStatus.Refresh
        Set rs = gobjSQL.OpenRecordset("Select ImmRelNCID from HistVacKey where HistKey = '" & NewHistKey & "'")
        If rs.EOF Then
             If gobjShared.UseMySQLDB Then
                gobjSQL.ExecuteSQL "UPDATE seq_imm_ref SET id=LAST_INSERT_ID(id+1)"
                Set rs = gobjSQL.OpenRecordset("SELECT LAST_INSERT_ID()")
             ElseIf gobjShared.UseLocalDB Then
                Set rs = gobjSQL.OpenRecordset("Select IMM_REF.Nextval from DUAL")
             End If
             NewNcid = True
        End If
        ImmRelNcid = rs(0).Value
        If rs.State <> 0 Then rs.Close
        Set rs = Nothing
        If NewNcid Then
            Set rs = gobjSQL.OpenRecordset("Select ImmRelNcid from HistVacKey where HistKey = '" & OldHistKey & "'")
            LastNcid = rs(0).Value
            If rs.State <> 0 Then rs.Close
            Set rs = Nothing
            sql = "Insert INTO HistVacKey (VacIDs, HistKey, PriorNCID, ImmRelNcid) Values ('" & HistVacIDs & "', '" & NewHistKey & "', " & CStr(LastNcid) & ", " & CStr(ImmRelNcid) & ")"
            gobjSQL.ExecuteSQL sql
        End If
        FrmStatus.LabStatus = FrmStatus.LabStatus & vbCrLf & "Successful"
        FrmStatus.Refresh
    End If
  
 If mVaccineDirty Then
    
   FrmStatus.LabStatus = FrmStatus.LabStatus & vbCrLf & "Modifying Vaccine File"
   FrmStatus.Refresh
 
        '<< SCR #25155;   Developer: Brian Mowbray 06/27/2002 02:02 PM
        xSet = "Vaccine = '" & msVaccine & "', " & _
               "ChildHood = " & msChildHood & ", " & _
               "EffectsVacID = '" & msEffectsVacID & "', " & _
               "NbrInSeries = " & msNbrInSeries & ", " & _
               "MPF_DIN = '" & msMPF_DIN & "', " & _
               "GenericName = '" & msGenericName & "', " & _
               "GenNameID = " & msGenNameID & ", " & _
               "Max_Age = '" & msMaxAge & "', " & _
               "Min_Age = '" & msMinAge & "', " & _
               "Booster = " & msBooster & ", " & _
               "CPT = '" & msCPT & "', " & _
               "ICD = '" & msICD & "', " & _
               "HL7_ID = '" & msHL7_ID & "', " & _
               "HistKey = '" & RemoveInvalidChar(NewHistKey) & "', " & _
               "ImmRelNCID = " & CStr(ImmRelNcid) & ", " & _
               "InActive = " & msInactive & ", " & _
               "VIS_Version = '" & ValidateDBChar(msVIS_Version) & "'"
     
             If Not gobjShared.UseLocalDB Then
               xSet = xSet & ", " & _
              "UpdatedOn = " & ConvertDateToIOType(Now) & ", " & _
              "UpdatedBy = " & gobjLogon.UserNCID & " "
             End If
             
      sql = "Update" & Name_Of_Vaccine_Table & "SET " & xSet & _
            "WHERE VaccineID = " & msVaccineID
      
      gobjSQL.ExecuteSQL sql
  
    FrmStatus.LabStatus = FrmStatus.LabStatus & vbCrLf & "Successful"
    FrmStatus.Refresh
 
  End If
  
  FrmStatus.LabStatus = FrmStatus.LabStatus & vbCrLf & "Adjusting Generic Names"
  FrmStatus.Refresh
  If PriorGenericName <> VaccineForm.CmbGenNames.Text Then
  
   '...Remove VacID from Prior Generic Name Dictionary
   IMM_Tables.ReMove_VacID_From_GenericName msVaccineID, PriorGenericName
   
   PriorGenericName = VaccineForm.CmbGenNames.Text
   
   '...Add Vacid to New Generic Name Dictionary
   IMM_Tables.Add_VacID_To_Generic_Name msVaccineID, PriorGenericName
   
  End If
  FrmStatus.LabStatus = FrmStatus.LabStatus & vbCrLf & "Successful"
  FrmStatus.Refresh
  
  '...See if we have to make mass modifications
  '   to the Vaccine and people's History file
  If NewHistKey <> OldHistKey And WhereVacIdsUpdate <> "" Then
  
     FrmStatus.LabStatus = FrmStatus.LabStatus & vbCrLf & "The Following will take awhile."
     FrmStatus.LabStatus = FrmStatus.LabStatus & vbCrLf & "Adjusting Related Vaccines"
     FrmStatus.Refresh
     
     '...Find any Previous records that might be related to the parent VacID
     If InStrRev(OldHistKey, PrefixInitVacidHF) <> 1 Then
        WhereVacIdsUpdate = Replace(WhereVacIdsUpdate, "=", "<>")
        WhereVacIdsUpdate = Replace(WhereVacIdsUpdate, " OR ", " AND ")
        sql = "Select VaccineID from" & Name_Of_Vaccine_Table & "WHERE " & WhereVacIdsUpdate & " AND  HistKey LIKE " & Left$(OldHistKey, InStrRev(OldHistKey, PrefixInitVacidHF) - 1) & "%"
        Set rs = gobjSQL.ExecuteSQL(sql)
        If rs.RecordCount > 0 Then
           With rs
             ReDim VacRec(.RecordCount - 1)
             Count = 0
             Do While Not .EOF
               VacRec(Count) = .Fields(0).Value
               .MoveNext
             Loop
           End With
           WhereVacIdsUpdate = Replace(WhereVacIdsUpdate, "<>", "=")
           WhereVacIdsUpdate = Replace(WhereVacIdsUpdate, " AND ", " OR ")
           For Count = 0 To UBound(VacRec)
             WhereVacIdsUpdate = WhereVacIdsUpdate & " OR VaccineID = " & VacRec(Count)
           Next Count
        End If
        If rs.State <> 0 Then rs.Close
        Set rs = Nothing
     End If
     '...Make all related Vaccines utilize the same ImmRelNcid
     sql = "Update" & Name_Of_Vaccine_Table & _
           "SET ImmRelNcid = " & CStr(ImmRelNcid) & ", HistKey = '" & NewHistKey & "' + Histkey" & _
           "WHERE " & WhereVacIdsUpdate
     gobjSQL.ExecuteSQL sql
     FrmStatus.LabStatus = FrmStatus.LabStatus & vbCrLf & "Successful"
     
     FrmStatus.LabStatus = FrmStatus.LabStatus & vbCrLf & "Adjusting Patient History File"
     FrmStatus.Refresh
     '...Adjust grouping keys
     sql = "Update " & "IMM_People_Vaccine" & _
           " SET ImmRelNcid = " & CStr(ImmRelNcid) & _
           " WHERE " & WhereVacIdsUpdate
     gobjSQL.ExecuteSQL sql
     FrmStatus.LabStatus = FrmStatus.LabStatus & vbCrLf & "Successful"
     FrmStatus.Refresh
     
     OldHistKey = NewHistKey
   
  End If
  
  '...Does the series need to be modified for this vaccine
  If mVaccineSeriesDirty Then
  
    FrmStatus.LabStatus = FrmStatus.LabStatus & vbCrLf & "Adjusting Immunization Schedule"
    FrmStatus.Refresh
    
    '...Need this in case no delete records found
    On Error Resume Next
    
    'CURRENTLY THE SCHEDULE NAME ACIP IS HARD CODED IN THE TRANSACTION
    'ADDING A SCHEDULE NAME REQUIRES SUBSTANTIAL MODIFICATION TO THE TABLES
    'AND OTHER RELATED FUNCTIONS.
    
    If msSchedule = "" Then msSchedule = DefaultVaccineSchedule
    
    '...Remove Schedule from Table
    sql = "Delete From TBLSchedule where VaccineID = " & msVaccineID & " AND ScheduleName = '" & msSchedule & "'"
    gobjSQL.ExecuteSQL sql
    
    '...Remove Schedule from Dictionary
    SeriesData = ""
    
    For Count = 1 To ModifiedVaccineSeries.Count
    
      VacRec = Split(ModifiedVaccineSeries.Item(CStr(Count)), FieldSepChar)
      sql = "Insert into TBLSchedule (ScheduleName, VaccineID, DoseNumber, Age_From, " & _
            "Age_To, Min_Age, Min_Interval, Min_Count, Recm_Age, Recm_Interval, Skip_Age, " & _
            "CreatedBy, CreatedOn) VALUES (" & _
            "'" & msSchedule & "', " & _
            VacRec(VacDictGrp2Fields.VaccineID) & ", " & _
            VacRec(VacDictGrp2Fields.DoseNumber) & ", '" & _
            VacRec(VacDictGrp2Fields.AgeFrom) & "', '" & _
            VacRec(VacDictGrp2Fields.AgeTo) & "', '" & _
            VacRec(VacDictGrp2Fields.MinAge) & "', '" & _
            VacRec(VacDictGrp2Fields.MinInterval) & "', '" & _
            VacRec(VacDictGrp2Fields.MinCount) & "', '" & _
            VacRec(VacDictGrp2Fields.RecmAge) & "', '" & _
            VacRec(VacDictGrp2Fields.RecmInterval) & "', '" & _
            VacRec(VacDictGrp2Fields.SkipAge) & "', " & _
            gobjLogon.UserNCID & ", " & _
            ConvertDateToIOType(Now) & _
            ")"
        
       gobjSQL.ExecuteSQL sql
       AddASeries VacRec
    Next Count
    
    FrmStatus.LabStatus = FrmStatus.LabStatus & vbCrLf & "Successful"
    FrmStatus.Refresh
     
  End If
  
  ModVaccineData = True
  
exitRoutine:

 Exit Function
 
haveerror:
' Debug.Print Err.Number, Err.Description
' Resume  ' TJNXXX
 ErrMsgBox "GenVaccineInfo_ModVaccineData: ", Error$, Err
 Resume exitRoutine
 Resume 'Debug
End Function

Public Function AddNewVaccineData() As Boolean

 ' Add a Record to the Vaccine DataBase (TBLVACCINES)

 On Local Error GoTo haveerror
 Dim sql As String
 Dim rs As ADODB.Recordset
 Dim VaccineID As String
 Dim vacidrec As String
 Dim GenNameRec() As String
 Dim RelatedIDs As String
 Dim vacids() As String
 
 FrmStatus.LabStatus = FrmStatus.LabStatus & vbCrLf & "Adding Vaccine to VaccineFile"
 FrmStatus.Refresh
 
     '<< SCR #25155;   Developer: Brian Mowbray 06/27/2002 02:02 PM
     '<< SCR #33656;   Developer: Brian Mowbray 03/14/2003 09:30 AM
     sql = "INSERT INTO " & Name_Of_Vaccine_Table & _
           " (VaccineID, Vaccine, Childhood, InActive, EffectsVacID, NbrInSeries, MPF_DIN, " & _
           "GenericName, GenNameID, Max_Age, Min_Age, booster, CPT, ICD, HL7_ID, HistKey, ImmRelNCID, VIS_Version, CreatedBy, CreatedOn) " & _
           "VALUES (" & msVaccineID & _
           ", '" & msVaccine & _
           "', " & msChildHood & _
           ", " & msInactive & _
           ", '" & msEffectsVacID & _
           "', '" & msNbrInSeries & _
           "', '" & msMPF_DIN & _
           "', '" & msGenericName & _
           "', " & msImmRelNcid & _
           ", '" & msMaxAge & _
           "', '" & msMinAge & _
           "', " & msBooster & _
           ", '" & msCPT & _
           "', '" & msICD & _
           "', '" & msHL7_ID & _
           "', '" & msHistKey & _
           "', " & msImmRelNcid & _
           ", '" & ValidateDBChar(msVIS_Version) & _
           "', " & gobjLogon.UserNCID & ", " & ConvertDateToIOType(Now) & " )"
           
 '...Add Record to Vaccine Table
 gobjSQL.ExecuteSQL sql
 
 FrmStatus.LabStatus = FrmStatus.LabStatus & vbCrLf & "Succesful"
 FrmStatus.Refresh
 
 '...Update New Generic Name Dictionary
 IMM_Tables.Insert_VacID_To_GenName msVaccineID, msGenericName
 
 '...Set Old Data
 PriorGenericName = msGenericName
 
 FrmStatus.LabStatus = FrmStatus.LabStatus & vbCrLf & "Updating VacKey"
 FrmStatus.Refresh
 '...See if HistVacKey is Properly Initialized in case we are
 '   Using a New Generic Name
 sql = "Select VacIds from HistVacKey where ImmRelNCID = " & msImmRelNcid
 Set rs = gobjSQL.OpenRecordset(sql)
 With rs
   If Not .EOF Then
    RelatedIDs = .Fields(0).Value
    If RelatedIDs = "" Then
      '<< Begin:SCR #36395;   Developer: Brian Mowbray 05/21/2003 05:05 PM
      gobjSQL.ExecuteSQL ("Update HistVacKey Set VacIDs = '" & WithinFieldSepChar & msVaccineID & WithinFieldSepChar & "' Where ImmRelNCID = " & msImmRelNcid)
    '>> End: SCR #36395;
    End If
    'See if this was a Combination drug
    GenNameRec = IMM_Tables.Get_A_Generic_Name_Record(msGenericName)
    If GenNameRec(GenVacDictFields.Combination) = "1" Then
      '...Have a Combination drug
      If RelatedIDs <> "" Then
        vacids = Split(Mid$(RelatedIDs, 2, Len(RelatedIDs) - 2), WithinFieldSepChar)
        '...Get Common Combination drugs
        If UBound(vacids) > 0 Then
         '...Get Vacids that do the same thing
         vacids = Split(vacids(UBound(vacids)), MultSameGenComb)
         '...Make new entry have same immrelkey and histkey and effectsvacid
         gobjSQL.ExecuteSQL ("Update" & Name_Of_Vaccine_Table & "N Set N.immrelkey = O.immrelkey, N.HistKey = O.HistKey, N.EffectsVacID = O.EffectsVacID" & msVaccineID & " Where O.VaccineID = (Select * From" & Name_Of_Vaccine_Table & "O Where O.VaccineID = " & vacids(0) & ")")
        End If
      End If
    End If
    FrmStatus.LabStatus = FrmStatus.LabStatus & vbCrLf & "Successful"
    FrmStatus.Refresh
   Else
     MsgBxARMd "Invald ImmRelNCID in HistVacKey Table" & vbCrLf & vbCrLf & _
            "Vaccine was added", vbExclamation + vbOKOnly, "Generic Name: " & msGenericName
   End If
 End With
 
 rs.Close
 Set rs = Nothing
 
 AddNewVaccineData = True
  
  
exitRoutine:

 Exit Function
 
haveerror:
'Debug.Print Err.Number, Err.Description
'Resume         ' TJNXXX
 ErrMsgBox "GenVaccineInfo_AddNewVaccineData:", Error$, Err
 Resume exitRoutine
 Resume
End Function

Public Property Get Vaccine() As String
 Vaccine = msVaccine
End Property

Public Property Let Vaccine(ByVal vNewValue As String)
 msVaccine = RemoveInvalidChar(vNewValue)
End Property

Public Property Get MinAge() As String
MinAge = msMinAge
End Property

Public Property Let MinAge(ByVal vNewValue As String)
 msMinAge = UCase$(vNewValue)
End Property

Public Property Get MaxAge() As String
 MaxAge = msMaxAge
End Property

Public Property Let MaxAge(ByVal vNewValue As String)
  msMaxAge = UCase$(vNewValue)
End Property

Public Property Get VaccineID() As String
  VaccineID = msVaccineID
End Property

Public Property Let VaccineID(ByVal vNewValue As String)
  msVaccineID = vNewValue
End Property

Public Property Get EffectsVacID() As String
  EffectsVacID = msEffectsVacID
End Property

Public Property Let EffectsVacID(ByVal vNewValue As String)
  msEffectsVacID = vNewValue
End Property

Public Property Get HL7_ID() As String
  HL7_ID = msHL7_ID
End Property

Public Property Let HL7_ID(ByVal vNewValue As String)
  msHL7_ID = RemoveInvalidChar(vNewValue)
End Property

Public Property Get NbrInSeries() As String
  NbrInSeries = msNbrInSeries
End Property

Public Property Let NbrInSeries(ByVal vNewValue As String)
  msNbrInSeries = RemoveInvalidChar(vNewValue)
End Property

Public Property Get GenericName() As String
  GenericName = msGenericName
End Property

Public Property Let GenericName(ByVal vNewValue As String)
  msGenericName = vNewValue
End Property

Public Property Get CPT() As String
  CPT = msCPT
End Property

Public Property Let CPT(ByVal vNewValue As String)
  msCPT = RemoveInvalidChar(vNewValue)
End Property

Public Property Get ICD() As String
  ICD = msICD
End Property

Public Property Let ICD(ByVal vNewValue As String)
  msICD = RemoveInvalidChar(vNewValue)
End Property

Public Property Get Booster() As String
  Booster = msBooster
End Property

Public Property Let Booster(ByVal vNewValue As String)
  msBooster = vNewValue
End Property

Public Property Get MPF_DIN() As String
  MPF_DIN = msMPF_DIN
End Property

Public Property Let MPF_DIN(ByVal vNewValue As String)
  msMPF_DIN = RemoveInvalidChar(vNewValue)
End Property

Public Property Get InActive() As String
  InActive = msInactive
End Property
Public Property Get ChildHood() As String
  ChildHood = msChildHood
End Property

Public Property Let InActive(ByVal vNewValue As String)
  msInactive = vNewValue
End Property

Public Property Let ChildHood(ByVal vNewValue As String)
  msChildHood = vNewValue
End Property


Public Property Get HistKey() As String
  HistKey = msHistKey
End Property

Public Property Let HistKey(ByVal vNewValue As String)
  msHistKey = vNewValue
End Property

Public Property Get GenNameID() As String
  GenNameID = msGenNameID
End Property

Public Property Let GenNameID(ByVal vNewValue As String)
  msGenNameID = vNewValue
End Property

Public Property Get ImmRelNcid() As String
  ImmRelNcid = msImmRelNcid
End Property

Public Property Let ImmRelNcid(ByVal vNewValue As String)
  msImmRelNcid = vNewValue
End Property

Public Function Record() As String

    ' Build Vaccine Record

    Dim xfield As Integer
    Dim field As String
    Dim sepchar As String
    Dim DelRecord As String
    
    For xfield = 0 To VacDictGrp1Fields.NumFlds - 1
    
         Select Case xfield
            Case VacDictGrp1Fields.Vaccine
               field = msVaccine
               
            Case VacDictGrp1Fields.MinAge
               field = msMinAge
               
            Case VacDictGrp1Fields.MaxAge
               field = msMaxAge
               
            Case VacDictGrp1Fields.VaccineID
               field = msVaccineID
               
            Case VacDictGrp1Fields.EffectsVacID
               field = msEffectsVacID
               
            Case VacDictGrp1Fields.HL7_ID
               field = msHL7_ID
               
            Case VacDictGrp1Fields.NbrInSeries
               field = msNbrInSeries
               
            Case VacDictGrp1Fields.GenericName
               field = msGenericName
               
            Case VacDictGrp1Fields.CPT
               field = msCPT
               
            Case VacDictGrp1Fields.ICD
               field = msICD
               
            Case VacDictGrp1Fields.Booster
               field = msBooster
               
            '<< SCR #25155;   Developer: Brian Mowbray 06/27/2002 02:02 PM
            Case VacDictGrp1Fields.VIS_Version
               field = msVIS_Version
          
            Case VacDictGrp1Fields.MPF_DIN
               field = msMPF_DIN
               
            Case VacDictGrp1Fields.ChildHood
               field = msChildHood
               
            Case VacDictGrp1Fields.InActive
               field = msInactive
               
            Case VacDictGrp1Fields.HistKey
               field = msHistKey
               
            Case VacDictGrp1Fields.GenNameID
               field = msGenNameID
               
            Case VacDictGrp1Fields.ImmRel_Ncid
               field = msImmRelNcid
               
            Case VacDictGrp1Fields.InActive
               field = msSchedule
        End Select
        
        DelRecord = DelRecord & sepchar & field
        sepchar = FieldSepChar
        
    Next xfield
    
    Record = DelRecord

End Function

Public Function SeriesArray() As String
  SeriesArray = SeriesData
End Function

Public Function ExpertArray() As String
  ExpertArray = ExpertData
End Function

Public Function AddASeries(ASeries() As String) As Boolean

   Dim Count As Integer
   Dim swork As String
   Dim sepchar As String
   Dim Record() As String
   Dim Series As String
   
      '...Invalid Data
      If UBound(ASeries) <> VacDictGrp2Fields.NumFlds - 1 Then Exit Function
      
      Record = Split(SeriesData, SeriesSepChar)
      
      If UBound(Record) = -1 Then
       Series = "1"
      Else
       Series = CStr(UBound(Record) + 2)
      End If
                  swork = ""
                  sepchar = ""
                  For Count = 0 To VacDictGrp2Fields.NumFlds - 1
                     swork = swork & sepchar & ASeries(Count)
                     sepchar = FieldSepChar
                  Next Count
                  
                  SeriesData = SeriesData & IIf(SeriesData = "", "", SeriesSepChar) & swork
                  
                  AddASeries = True
End Function

Public Property Let SaveASchedule(ByVal vNewValue As String)
 '...Set an Entire immunization schedule
 SeriesData = vNewValue
End Property

Public Property Get Schedule() As String
 Schedule = msSchedule
End Property

Public Property Let Schedule(ByVal vNewValue As String)
 msSchedule = vNewValue
End Property

Private Sub Class_Initialize()
    '--- COMMON TO MULTIPLE SCR-ARMs: MSolano 9/13/01 SCR-ARM-16110, 16130
    '--- Init routine should be called from first top level creatable coclass
    '--- It can be called multiple times, but will actually run only once.
    '--- Instrument all creatable classes in case dll's multiple entry
    '--- points are actually used
    Call ARM_InitAllTransactionsForImmunizations
End Sub

'<< SCR #25155;   Developer: Brian Mowbray 06/27/2002 02:02 PM
Public Property Get VIS_Version() As String
  VIS_Version = msVIS_Version
End Property

'<< SCR #25155;   Developer: Brian Mowbray 06/27/2002 02:03 PM
Public Property Let VIS_Version(ByVal vNewValue As String)
 msVIS_Version = vNewValue
End Property
