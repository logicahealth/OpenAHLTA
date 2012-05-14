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
Attribute VB_Name = "Patient_Imm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'...Error Handling Variables
Private lLastError As Long
Private sLastErrorDesc As String

'...Holds all Group Names Used by a patient
Private obd_PatientsGroups As Dictionary

'...Holds all Generic Names Used by a patient
Private obd_PatientsGenNames As Dictionary

'...Patients's (obd_PatientsGenNames) Fields
Private Enum PatientGenNames
  AgeIsAFactor
  GenericName
  VaccineID
  Vacc_Series
  Vacc_Date
  vacc_Nextseries
  Vacc_NextDue
  VAcc_Current
  Vacc_Exempt
  ExemptionExpirationDate
  result
  CurrentStatus
  idcontrol                 ' see DefIDControl for definition
  NumFlds = idcontrol       ' Always last and set = to preceeding field
End Enum

'...Dictionary of Patients Last History record/GenericName
Private obd_PatientsHistory As Dictionary
'...Dictionary of Patients Vaccination
Private obd_PatientsVaccines As Dictionary

Private DOB As Variant
Private SSN As String
Private Unit As String
Private Status As String
Private Service As String
Private PatVaccineHistory As ADODB.Recordset

Private Type Combination
  VaxId As String                ' VaxID of Combination vaccine
  VacEffected As String          ' All VacId's that can be effected
  ChosenGenName As String        ' Generic name to associate with
End Type

Private Potion As Combination    ' Combination Data
Public Sub dump_obd_Patientsvaccines()

   On Local Error GoTo haveerror

    ' Dump Patient Vaccine Dictionary    TJN_REMOVE

     Dim Count As Integer
     Dim gn As String
     Dim VacKey As Variant
     Dim sGeneric As String
     Dim VacFields() As String
     Dim newage As Variant
     Dim sSQL As String
     Dim VaccinesToAnalyze As Dictionary
     Dim displayed As Boolean
     
     
     Set VaccinesToAnalyze = New Dictionary
     
     For Each VacKey In obd_PatientsVaccines
      VacFields = GetPatientsLastVaccineRecord(VacKey)
      For Each newage In VacFields
       If newage = "" Then newage = "Void"
       Select Case Count
         Case vacrsfields.VaccineID
           sSQL = newage
           VaccinesToAnalyze.Add sSQL, sSQL
       End Select
       Count = Count + 1
      Next newage
     Next VacKey
     
     Dim Group1() As String
     Dim Group2() As String
     Dim Group3() As String
     Debug.Print "Vaccines To Analyze"
     Debug.Print
     For Each VacKey In VaccinesToAnalyze
      VacFields = IMM_Tables.Vaccine_Record(CStr(VacKey))
      '...Display General Data
      Group1 = Split(VacFields(0), FieldSepChar)
      Debug.Print
      Debug.Print "Next Vaccine"
      Debug.Print
      Debug.Print "Vac Description", "MinAge", "MaxAge", "VaccineID", "EffectsVacID", "HL7_ID", "NBRInSeries", "GenericName", "    CPT", "ICD", "Booster", "MPF_DIN", "ChildHood", "    HistKey", "          GenNameID", "IMMRel_Ncid", "    Inactive"
      Count = 0
      gn = ""
      For Each newage In Group1
        If newage = "" Then newage = "Void"
        Select Case Count
           Case VacDictGrp1Fields.Vaccine
            sGeneric = ""
           Case VacDictGrp1Fields.MinAge
            If Len(sSQL) >= 28 Then
             sGeneric = Left$(sSQL, 23) & Space$(5)
            Else
             sGeneric = Space$(28 - Len(sSQL))
            End If
            '<< SCR #37932;   Developer: Brian Mowbray 06/18/2003 01:32 PM
            'Changed so that when this information is being written it will not error based on the length of the data
           Case VacDictGrp1Fields.MaxAge
            sGeneric = Space$(Len(sSQL) + 5)
           Case VacDictGrp1Fields.VaccineID
            sGeneric = Space$(Len(sSQL) + 5)
           Case VacDictGrp1Fields.EffectsVacID
            sGeneric = Space$(Len(sSQL) + 5)
           Case VacDictGrp1Fields.HL7_ID
            sGeneric = Space$(Len(sSQL) + 5)
           Case VacDictGrp1Fields.NbrInSeries
            sGeneric = Space$(Len(sSQL) + 5)
           Case VacDictGrp1Fields.GenericName
            sGeneric = Space$(Len(sSQL) + 5)
           Case VacDictGrp1Fields.CPT
            sGeneric = Space$(Len(sSQL) + 5)
           Case VacDictGrp1Fields.ICD
            sGeneric = Space$(Len(sSQL) + 5)
           Case VacDictGrp1Fields.Booster
            sGeneric = Space$(Len(sSQL) + 5)
           Case VacDictGrp1Fields.MPF_DIN
            sGeneric = Space$(Len(sSQL) + 5)
           Case VacDictGrp1Fields.ChildHood
            sGeneric = Space$(Len(sSQL) + 5)
           Case VacDictGrp1Fields.InActive
            sGeneric = Space$(Len(sSQL) + 5)
           Case VacDictGrp1Fields.HistKey
            sGeneric = Space$(Len(sSQL) + 5)
           Case VacDictGrp1Fields.GenericName
            sGeneric = Space$(Len(sSQL) + 5)
           Case VacDictGrp1Fields.ImmRel_Ncid
            sGeneric = Space$(Len(sSQL) + 5)
           Case VacDictGrp1Fields.InActive
            sGeneric = Space$(Len(sSQL) + 5)
        End Select
       gn = gn & sGeneric & newage
       Count = Count + 1
       sSQL = newage
      Next newage
      Debug.Print gn
      '...Display Series
      Group2 = Split(VacFields(1), SeriesSepChar)
      Dim SubGroup() As String
      displayed = False
      
      For Each newage In Group2
       SubGroup = Split(newage, FieldSepChar)
       gn = ""
       For Count = 0 To UBound(SubGroup)
         If Not displayed Then
            Debug.Print
            Debug.Print "Display Series for this Vaccination"
            Debug.Print
            Debug.Print "VaccineID", "DoseNumber", "AgeFrom", "AgeTO", "MinAge", "MinInterval", "MinCount", "RecmAge", "RecmInterval", "SkipAge"
            displayed = True
         End If
         newage = SubGroup(Count)
         If newage = "" Then newage = "Void"
         Select Case Count
           Case VacDictGrp2Fields.VaccineID
            sGeneric = ""
           Case VacDictGrp2Fields.DoseNumber
            sGeneric = Space$(Len(sSQL) + 5)
           Case VacDictGrp2Fields.AgeFrom
            sGeneric = Space$(Len(sSQL) + 5)
           Case VacDictGrp2Fields.AgeTo
            sGeneric = Space$(Len(sSQL) + 5)
           Case VacDictGrp2Fields.MinAge
            sGeneric = Space$(Len(sSQL) + 5)
           Case VacDictGrp2Fields.MinInterval
            sGeneric = Space$(Len(sSQL) + 5)
           Case VacDictGrp2Fields.MinCount
            sGeneric = Space$(Len(sSQL) + 5)
           Case VacDictGrp2Fields.RecmAge
            sGeneric = Space$(Len(sSQL) + 5)
           Case VacDictGrp2Fields.RecmInterval
            sGeneric = Space$(Len(sSQL) + 5)
           Case VacDictGrp2Fields.SkipAge
            sGeneric = Space$(Len(sSQL) + 5)
        End Select
        gn = gn & sGeneric & newage
        sSQL = newage
       Next Count
       Debug.Print gn
      Next newage
      If Not displayed Then
         Debug.Print
         Debug.Print "No series for this Vaccination"
      End If
      Debug.Print
      Debug.Print "Expert System not implimented"
      Debug.Print "for this vaccination."
      Debug.Print
      '...Display Additional Conditions
      Group3 = Split(VacFields(2), SeriesSepChar)
      displayed = False
     Next VacKey
     
     Debug.Print
     Debug.Print "Patients_Vaccines   Dictionary    For SSN:  " & SSN
     Debug.Print
     Debug.Print "GenericName", "MinAge", "MaxAge", "IgnoreAge", "VaccineID", "Vacc_Series", "Vacc_Date", "Next_Series", "Vacc_NextDue", "Vacc_Current", "Vacc_Exempt", "SolvedByVacID", "       ExExpDate", "Result", "Beg-Status", "End-Status", "ID-Control"
     For Each VacKey In obd_PatientsVaccines
      VacFields = GetPatientsLastVaccineRecord(VacKey)
      gn = ""
      Count = 0
      For Each newage In VacFields
       If newage = "" Then newage = "Void"
       Select Case Count
        Case vacrsfields.GenericName
         sGeneric = ""
         newage = Left$(newage, 12)
        Case vacrsfields.MinAge
         sGeneric = Space$(Len(sSQL) + 5)
         newage = Left$(newage, 12)
        Case vacrsfields.MaxAge
         sGeneric = Space$(Len(sSQL) + 5)
         newage = Left$(newage, 12)
        Case vacrsfields.IgnoreAge
         sGeneric = Space$(Len(sSQL) + 5)
         newage = Left$(newage, 12)
        Case vacrsfields.VaccineID
         sGeneric = Space$(Len(sSQL) + 5)
         newage = Left$(newage, 12)
        Case vacrsfields.Vacc_Series
         sGeneric = Space$(Len(sSQL) + 5)
         newage = Left$(newage, 12)
        Case vacrsfields.Vacc_Date
         sGeneric = Space$(Len(sSQL) + 5)
         newage = Left$(newage, 12)
        Case vacrsfields.vacc_Nextseries
         sGeneric = Space$(Len(sSQL) + 5)
         newage = Left$(newage, 12)
        Case vacrsfields.Vacc_NextDue
         sGeneric = Space$(Len(sSQL) + 5)
         newage = Left$(newage, 12)
        Case vacrsfields.VAcc_Current
         sGeneric = Space$(Len(sSQL) + 5)
         newage = Left$(newage, 12)
        Case vacrsfields.Vacc_Exempt
         sGeneric = Space$(Len(sSQL) + 5)
         newage = Left$(newage, 12)
        Case vacrsfields.SolvedByVacID
         sGeneric = Space$(Len(sSQL) + 5)
         newage = Left$(newage, 12)
        Case vacrsfields.ExemptionExpirationDate
         sGeneric = Space$(Len(sSQL) + 5)
         newage = Left$(newage, 12)
        Case vacrsfields.result
         sGeneric = Space$(Len(sSQL) + 5)
         newage = Left$(newage, 12)
        Case vacrsfields.begstatus
         sGeneric = Space$(Len(sSQL) + 5)
         newage = Left$(newage, 12)
        Case vacrsfields.EndStatus
         sGeneric = Space$(Len(sSQL) + 5)
         newage = Left$(newage, 12)
        Case vacrsfields.idcontrol
         sGeneric = Space$(Len(sSQL) + 5)
         newage = Left$(newage, 12)
       End Select
       gn = gn & sGeneric & newage
       sSQL = newage
       Count = Count + 1
      Next newage
      Debug.Print gn
     Next VacKey
     
exitRoutine:
  Exit Sub
haveerror:
 ErrMsgBox "Patient_Imm_Dump_obd_Patientsvaccines", Error$, Err
 Resume exitRoutine
 Resume
End Sub
Public Function ImmRecordExist(ByVal lVacID As Long, ByVal vDate As Variant) As Boolean
  Dim VacRec() As String

  On Error GoTo Func_Err
  
  ImmRecordExist = False

  'Check to see if vaccine exists in dictionary
  If obd_PatientsVaccines.Exists(CStr(lVacID)) Then
  
    
    If utNullToZero(vDate) <> 0 Then
      
      'Grab the last vaccine for this vaccine id and check to see if the date matches
      VacRec = PatientIMM.GetPatientsLastVaccineRecord(lVacID, False)
      If UBound(VacRec) <> -1 Then
        If Len(VacRec(vacrsfields.Vacc_Date)) > 0 Then
          If CDate(VacRec(vacrsfields.Vacc_Date)) = CDate(vDate) Then
            ImmRecordExist = True
          End If
        End If
      End If
    Else
      'Set to true because an Immunization does exist for this vaccine id
      ImmRecordExist = True
    End If
  End If
 
Func_Exit:
  On Error Resume Next
  Exit Function

Func_Err:
 ErrMsgBox "Patient_Imm.ImmRecordExist", Error$, Err
 Resume Func_Exit
 Resume 'Debug
End Function

Private Sub PatientGenericIMMSchedule(Optional ByVal AddedGenNames As String = vbNullString)

  ' Loop through all Generic Names required by a Patient
  ' Pull all Vacines Defined in a Generic Name (Inverted Dictionary)
  ' Address those vaccines in PatientsVaccine Dictionary
  
  Dim gn As Variant
  Dim AddedGenericNames() As String
  Dim AnalyzeGenName As Boolean
  Dim AddedGenericName As Variant
  
  On Error GoTo Sub_Err
  
  '...Is this a Modification to Patient
  If LenB(AddedGenNames) <> 0 Then
     '  YES
     AddedGenericNames = Split(AddedGenNames, WithinFieldSepChar)
  Else
     AnalyzeGenName = True
  End If
  
  '...Get Next Generic Name for this patient
  For Each gn In obd_PatientsGenNames
'    Debug.Assert LCase(CStr(gn)) <> "ippd"

    ' Are we Processing Specific Generic Names
    If LenB(AddedGenNames) <> 0 Then
      '  YES
        AnalyzeGenName = False
        For Each AddedGenericName In AddedGenericNames
          If gn = AddedGenericName Then
             AnalyzeGenName = True
             Exit For
          End If
        Next AddedGenericName
    End If
    
    If AnalyzeGenName Then UpdateVactoGenForPatient gn
    
  Next gn

Sub_exit:
  On Error Resume Next
  Exit Sub

Sub_Err:
 ErrMsgBox "Patient_Imm.PatientGenericIMMSchedule", Error$, Err
 Resume Sub_exit
 Resume 'Debug


End Sub
Private Sub CalculateNextDueDate(ByVal EvaluationDate As Variant, Optional ByVal AddedVaccines As String = vbNullString)
 
   ' This Routine Calucaltes the Next Due Date
   ' for a vaccine.  Input is a dictionary of
   ' vaccines to be considered for a patient.
   ' This routine only calucaltes the next due
   ' date and sets an appropriate status.
   
   On Local Error GoTo haveerror
   
   Dim VacKey As Variant
   Dim AnalyzeVaccine As Boolean
   
   '...Is This a Patient Modification
   If LenB(AddedVaccines) <> 0 Then
     ' YES
        Dim AddedVacKeys() As String
        Dim AddedVacKey As Variant
        AddedVacKeys = Split(AddedVaccines, WithinFieldSepChar)
   Else
        AnalyzeVaccine = True
   End If
   
   If Not IsDate(EvaluationDate) Then EvaluationDate = Now
   
   For Each VacKey In obd_PatientsVaccines
   
     If LenB(AddedVaccines) <> 0 Then
        AnalyzeVaccine = False
        For Each AddedVacKey In AddedVacKeys
           If AddedVacKey = VacKey Then
              AnalyzeVaccine = True
              Exit For
           End If
        Next AddedVacKey
     End If
   
     If AnalyzeVaccine Then ADueDate CStr(VacKey), EvaluationDate, False
     
   Next VacKey
   
exitRoutine:
 Exit Sub
haveerror:
 ErrMsgBox "Patient_Imm_CalculateNextDueDate", Error$, Err
 Resume exitRoutine
 Resume
End Sub

Public Function GetPatientsVaccineSchedule(ByVal Refresh As Boolean, _
                                           Optional ByVal sspon_SSAN_FMP As String, _
                                           Optional ByVal sUnit As String = "", _
                                           Optional ByVal vDOB As Variant = vbNullString, _
                                           Optional ByVal vEvalDate As Variant = vbNullString, _
                                           Optional ByVal sStatus As String = vbNullString, _
                                           Optional ByVal sService As String = vbNullString) As ADODB.Recordset
  
  On Error GoTo haveerror
  
  ' Establish Patient's basic Demographic Data
  If Not Refresh Then
    DOB = vDOB
    SSN = sspon_SSAN_FMP
    Unit = sUnit
    Status = sStatus
    Service = sService
  Else
    '...Reset Patient's ditionaries
    Reset
  End If
  
  If utNullToZero(vEvalDate) = 0 Then
    vEvalDate = Date
  End If
  
  '...Establish Demis
  If gsMTF_ID = vbNullString Then SayMTF_ID

  '...Load Unit Group Requirements if necessary
  '...including Service, Status codes and Age
  IMM_Tables.LoadRequiredUnitGroups Unit, Service, Status, DOB
  '...Populate Patient's Groups
  IMM_Tables.GetUnitGroups obd_PatientsGroups
  
  '...Get User Defined Groups
  LoadUserDefinedGroups
  
  '...Populate Patient with required Generic Vaccine Names
      ' bookmark
  LoadPatientGenericNames CDate(vEvalDate)
  
  '...Populate Patients Vaccine Table
  LoadPatientsVaccines vEvalDate
    
  '...Load Patient's Vaccination History
  LoadPatientsIMMHistory

  
  '...Calculate Next due date base on Series and cycles only
  CalculateNextDueDate vEvalDate
  
  '-----------------------------------------------------------------------------------
     ' Code goes here for removing Vaccines due to
     ' Conditions (Expert System) entered through GUI
  '-----------------------------------------------------------------------------------
  
  '...Check for Combination Vaccines and adjust status of vaccines
  '...effected by a combination drug
  CheckForCombinationVaccines
  
  '...Build Patient Generic IMM Schedule base on Patient's Vaccine Table
  PatientGenericIMMSchedule
  
  '-----------------------------------------------------------------------------------
     ' Debug Print out Dictionaries
     
     Call dump_obd_Patientsvaccines
     
  '-----------------------------------------------------------------------------------
  
exitRoutine:
  
  Exit Function
  
haveerror:
'  Debug.Print Err.Number, Err.Description
'  Resume                        ' TJNXXX
  lLastError = Err.Number
  sLastErrorDesc = "Patient_IMM: GetPatientsVaccineSchedule: " & Err.Description
  Err.Raise lLastError, "Patient_IMM: GetPatientsVaccineSchedule: ", sLastErrorDesc
End Function


Public Sub LoadUserDefinedGroups()

 On Local Error GoTo haveerror

 ' This Builds a Group Record Dictionary indexed by group
 ' This Builds a Group for a Patient inverted file indexed by Patient

 Dim rs As ADODB.Recordset
 Dim Group As ADODB.field
 Dim sql As String
 
 sql = "SELECT Group_Name,MTF_ID from IMM_People_Group where SPON_SSAN_FMP = '" & SSN & "' AND MTF_ID = '" & gsMTF_ID & "'"
 
 '...Read IMM_People_Group Table
 
 Set rs = gobjSQL.OpenRecordset(sql, adOpenForwardOnly, adLockReadOnly)
 With rs
   
   If Not .EOF Then
   
    '...establish fast Field access
    Set Group = .Fields("Group_Name")
    
    Do While Not .EOF
      AddGroupToPatient Group
      .MoveNext
    Loop
   End If
   .Close
 End With
 
 Set rs = Nothing
 
 
exitRoutine:

 Exit Sub
 
haveerror:
 ErrMsgBox "Patient_Imm_LoadUserDefinedGroups", Error$, Err
 Resume exitRoutine
 Resume 'Debug
End Sub
Public Function IsUserDefinedGroup(ByVal GroupName As String) As Boolean
  
  Dim GrpRec() As String
  
  On Error GoTo Func_Err
  
  If obd_PatientsGroups.Exists(GroupName) Then
   GrpRec = Split(obd_PatientsGroups.Item(GroupName), FieldSepChar)
   IsUserDefinedGroup = IIf((Val(GrpRec(grouprec.idcontrol)) And DefIDControl.UserGroup) <> 0, True, False)
  End If

Func_Exit:
  On Error Resume Next
  Exit Function

Func_Err:
 ErrMsgBox "Patient_Imm.IsUserDefinedGroup", Error$, Err
 Resume Func_Exit
 Resume 'Debug

End Function

Private Sub Class_Initialize()
    '--- COMMON TO MULTIPLE SCR-ARMs: MSolano 9/13/01 SCR-ARM-16110, 16130
    '--- Init routine should be called from first top level creatable coclass
    '--- It can be called multiple times, but will actually run only once.
    '--- Instrument all creatable classes in case dll's multiple entry
    '--- points are actually used
    Call ARM_InitAllTransactionsForImmunizations
End Sub

Private Sub Class_Terminate()
  On Error Resume Next
  If Not PatVaccineHistory Is Nothing Then
    If PatVaccineHistory.State <> 0 Then PatVaccineHistory.Close
  End If
  Set PatVaccineHistory = Nothing
  Set obd_PatientsHistory = Nothing
  Set obd_PatientsGenNames = Nothing
  Set obd_PatientsVaccines = Nothing
  Set obd_PatientsGroups = Nothing
End Sub


Public Function LoadPatientGenericNames(Optional ByVal vDate As Date = #1/1/100#, Optional GroupName As String = vbNullString) As String

 On Local Error GoTo haveerror

 '  This routine loads Generic Names for a patient
 '  Based on Groups already assigned to that patient
 '  and age critieria controlled by groups

 Dim Group As Variant
 Dim Generic As Variant
 Dim GrpRecord() As String
 Dim UseRecord As Boolean
 Dim newage As Variant
 Dim AgeIsAFactor As String
 Dim GenNames() As String
 Dim GenRec() As String
 Dim GenNamesAdded As String
 Dim AddedSepChar As String
 
  If vDate = #1/1/100# Then vDate = Date
  
  '...Establish Generic Name Record
  ReDim GenRec(PatientGenNames.NumFlds)
  
  '...Look at all the groups assigned to this patient
  For Each Group In obd_PatientsGroups
'    Debug.Assert Trim(CStr(group)) <> ""
    '...Default to use this group
    If GroupName = vbNullString Or Group = GroupName Then
        UseRecord = True
        
        '...Load a Group Record assigned to this patient
        GrpRecord = Split(obd_PatientsGroups.Item(Group), FieldSepChar)
        
        '...do we have a valid date of birth
        '...If not then age won't be a factor
        If IsDate(DOB) Then
        
         '...Is this group controlled by age
         AgeIsAFactor = IIf(GrpRecord(grouprec.ConsiderAge) = "1", "Y", "N")
         If AgeIsAFactor = "Y" Then
         
           '...Yes Make sure patient is old enough
           newage = CalcAgeOrTime(DOB, GrpRecord(grouprec.MinAge))
           
           '...If no age was entered in Group definition
           '...then assume age is ok
           If IsDate(newage) Then
           
            '...Set Don't use record if patient is not old enough
            If newage > vDate Then UseRecord = False
            
           End If
           
           '...Is record still valid
           If UseRecord Then
           
            '...Yes Make sure patient is not to old
            newage = CalcAgeOrTime(DOB, GrpRecord(grouprec.MaxAge))
            
            '...If no age was entered in Group definition
            '...then assume age is ok
            If IsDate(newage) Then
            
             '...Set Don't use record if patient is too old
             If newage < vDate Then UseRecord = False
             
            End If
           End If
         End If
         
        Else
        
         '...Since no Age was passed for this patient
         '...Indicate Age is not a factor
         AgeIsAFactor = "N"
         
        End If
        
        '...Does this record meet the neccesary conditions
        If UseRecord Then
        
            '...Yes load generic names associated with this record
            GenNames = Split(GrpRecord(grouprec.GenericNames), InvertedTblSepChar)
            
            '...If Generic name is not already defined to patient
            '...then put into patients generic name table
            '...along with age dependency information
            For Each Generic In GenNames
            
    '          Debug.Assert Trim$(CStr(Generic)) <> ""
            
              If Not obd_PatientsGenNames.Exists(Generic) Then
                '...Add Generic Name
                Debug.Assert Generic <> "ANTHRAX"
                GenRec(PatientGenNames.AgeIsAFactor) = AgeIsAFactor
                GenRec(PatientGenNames.GenericName) = CStr(Generic)
                GenRec(PatientGenNames.idcontrol) = CStr(Val(GenRec(PatientGenNames.idcontrol)) Or GrpRecord(grouprec.idcontrol))
                obd_PatientsGenNames.Add Generic, Join(GenRec, FieldSepChar)
                GenNamesAdded = GenNamesAdded & AddedSepChar & Generic
                AddedSepChar = WithinFieldSepChar
              End If
            Next Generic
        End If
    End If
  Next Group
  
  '...Return Generic Names added
  '...Not used when first initializing Patient
  LoadPatientGenericNames = GenNamesAdded
  
exitRoutine:
  Exit Function
haveerror:
 ErrMsgBox "Patient_Imm_LoadPatientGenericNames", Error$, Err
 Resume exitRoutine
End Function

Public Function LoadPatientsVaccines(ByVal vEvalDate As Variant, Optional ByVal GenNamesAdded As String = vbNullString) As String

  On Local Error GoTo haveerror

  Dim VacFields() As String
  Dim VacKey As Variant
  Dim VacID() As String
  Dim vackeydata() As String
  Dim gn As String
  Dim gnkeydata() As String
  Dim ValidAge As Boolean
  Dim newage As Variant
  Dim GenName As Variant
  Dim PatGenRec() As String
  Dim AgeIsAFactor As String
  Dim VacTable As Dictionary
  Dim ID As String
  Dim GenNamesInput() As String
  Dim MaxNamesAdded As Integer
  Dim GenNameInput As Variant
  Dim AnalyzeName As Boolean
  Dim VaccinesAdded As String
  Dim VaccinesAddedSepChar As String
  Dim ExistingKeys() As Variant
  
  Set VacTable = New Dictionary
  
  '...Is this a patient modification
  If LenB(GenNamesAdded) <> 0 Then
  
    ' YES
    '...Get vaccines used by patient
    ExistingKeys = obd_PatientsVaccines.Keys
    '...Populate all Vaccines already in Patient Dictionary
    For Each VacKey In ExistingKeys
     VacTable.Add VacKey, VacKey
    Next VacKey
    
    '...Yes get vaccine generic names added
    GenNamesInput = Split(GenNamesAdded, WithinFieldSepChar)
    MaxNamesAdded = UBound(GenNamesInput)
    
  Else
    ' Process all Generic Names
    AnalyzeName = True
  End If
  
  '...Establish Fields for Vaccine Record
  ReDim VacFields(vacrsfields.numFields - 1)
  
  For Each GenName In obd_PatientsGenNames
  
    gn = CStr(GenName)
    '...Is this a patient update
    If LenB(GenNamesAdded) <> 0 Then
       ' YES
       AnalyzeName = False
       For Each GenNameInput In GenNamesInput
         If GenName = GenNameInput Then
           '...Only analye those new generic names added
           AnalyzeName = True
           Exit For
         End If
       Next GenNameInput
    End If
    
    If AnalyzeName Then
    
        '...Load Patients generic Record
        PatGenRec = Split(obd_PatientsGenNames.Item(gn), FieldSepChar)
        ID = PatGenRec(PatientGenNames.idcontrol)
        
        '...Load Age Criteria for this generic name
        '...as Determined in routine LoadPatientGenericNames
        AgeIsAFactor = PatGenRec(PatientGenNames.AgeIsAFactor)
        
        '...Load Generic Name
        gnkeydata = IMM_Tables.Get_A_Generic_Name_Record(gn)
        
        '...Assume Valid Age
        ValidAge = True
        
        '...Make Sure Patient falls within valid limits
        '...of the the Generic Vaccine Name
        If AgeIsAFactor = "Y" And UBound(gnkeydata) <> -1 Then
        
            '...Yes Age is a factor
            '...First check age against Generic Name
             newage = CalcAgeOrTime(DOB, gnkeydata(GenVacDictFields.MinAge))
             
             If IsDate(newage) Then
              If newage > vEvalDate Then ValidAge = False
             End If
             If ValidAge Then
              newage = CalcAgeOrTime(DOB, gnkeydata(GenVacDictFields.MaxAge))
              If IsDate(newage) Then
               If newage < vEvalDate Then ValidAge = False
              End If
             End If
                
        End If
        
        If ValidAge Then
              
            '...Populate Patient with all Vaccines to consider
            '...for this generic name
            VacID = IMM_Tables.VacIDsFromGenName(gn)
            
            For Each VacKey In VacID
            
              '...Check to see if this vaccine already looked at
              If Not VacTable.Exists(VacKey) Then
              
                  '...Indicate this vaccine already analyzed
                  '...by age criteria
                  VacTable.Add VacKey, VacKey
                  
                  '...Get Minimal Vaccine Data
                  vackeydata = IMM_Tables.Vaccine_BasicData(VacKey)
              
                  '...Do we need to consider age  (DETERMINED WHEN LOADING GROUPS)
                  If AgeIsAFactor = "Y" Then
                  
                   '...Need to Re-establish valid age indicator
                   ValidAge = True
                  
                   '...If Valid Age Make sure Patient's age is within Vaccine's Age Limit
                   If ValidAge Then
                      '.....Incase vaccine's age range is a tighter bound then Generic Limits
                      newage = CalcAgeOrTime(DOB, vackeydata(VacDictGrp1Fields.MinAge))
                      
                      If IsDate(newage) Then
                       If newage > vEvalDate Then ValidAge = False
                      End If
                      If ValidAge Then
                       newage = CalcAgeOrTime(DOB, vackeydata(VacDictGrp1Fields.MaxAge))
                       If IsDate(newage) Then
                        If newage < vEvalDate Then ValidAge = False
                       End If
                      End If
                      
                   End If
                  End If
                  
                  '...Is Patient's Age Valid
                  If ValidAge Then
                      '...YES   Include Vaccine
                      VacFields(vacrsfields.GenericName) = gn
                      VacFields(vacrsfields.VaccineID) = VacKey
                      VacFields(vacrsfields.begstatus) = "N"       ' indicate New
                      VacFields(vacrsfields.idcontrol) = ID
                      obd_PatientsVaccines.Add VacKey, Join(VacFields, FieldSepChar)
                      VaccinesAdded = VaccinesAdded & VaccinesAddedSepChar & VacKey
                      VaccinesAddedSepChar = WithinFieldSepChar
                  End If
                  
              End If
            Next VacKey
            
        End If
        
    End If
  Next GenName
  LoadPatientsVaccines = VaccinesAdded
  Set VacTable = Nothing
exitRoutine:
  Exit Function
haveerror:
 ErrMsgBox "Patient_Imm_LoadPatientsVaccines", Error$, Err
 Resume exitRoutine
 Debug.Print Err.Number, Err.Description
 Resume

End Function

Public Function BuildGenericRecordSet(ByVal RequiredOnly As Boolean) As ADODB.Recordset
 On Local Error GoTo haveerror
 Dim gn As Variant
 Dim rst As ADODB.Recordset
 Dim GenNameRec() As String
 Dim ProcessData As Boolean
'  Set rst = New ADODB.Recordset
'  If rst.Fields.Count = 0 Then
'    rst.Fields.Append "GenericName", adVarChar, 50, adFldKeyColumn
'    rst.Fields.Append "VaccineID", adInteger, , adFldIsNullable
'    rst.Fields.Append "Vacc_Series", adInteger, , adFldIsNullable
'    rst.Fields.Append "Vacc_Date", adDate, , adFldIsNullable
'    rst.Fields.Append "Vacc_NextDue", adDate, , adFldIsNullable
'    rst.Fields.Append "Vacc_Current", adBoolean, , False
'    rst.Fields.Append "Vacc_Exempt", adInteger, , adFldIsNullable
'    rst.Fields.Append "ExemptionExpirationDate", adDate, , adFldIsNullable
'    rst.Fields.Append "Result", adChar, 1, adFldIsNullable
'  End If
'  rst.Open , , adOpenDynamic, adLockOptimistic
  
  '<< Begin:SCR #37777;   Developer: Brian Mowbray 06/16/2003 04:58 PM
  CreateRst Generic_rst, rst
  '>> End: SCR #37777;
  For Each gn In obd_PatientsGenNames
    ' Kluge Kluge Fix to get problem resolved
    If Not IsEmpty(gn) Then
        GenNameRec = Split(obd_PatientsGenNames.Item(gn), FieldSepChar)
    '    Debug.Assert LCase(gn) <> "mmr"
        Select Case GenNameRec(PatientGenNames.CurrentStatus)
            Case "V", "C"
            
              ProcessData = True
              If RequiredOnly And (Val(GenNameRec(PatientGenNames.idcontrol)) And DefIDControl.AllGroups) = 0 Then ProcessData = False
               
               If ProcessData Then
               
                With rst
                  .AddNew
                    .Fields("GenericName") = GenNameRec(PatientGenNames.GenericName) & vbNullString
                    .Fields("VaccineID") = GenNameRec(PatientGenNames.VaccineID)
                    If GenNameRec(PatientGenNames.vacc_Nextseries) <> "" Then
                     .Fields("Vacc_Series") = GenNameRec(PatientGenNames.vacc_Nextseries)
                    ElseIf GenNameRec(PatientGenNames.Vacc_Series) <> "" Then
                     .Fields("Vacc_Series") = GenNameRec(PatientGenNames.Vacc_Series)
                    End If
                    If IsDate(GenNameRec(PatientGenNames.Vacc_Date)) Then .Fields("Vacc_Date") = GenNameRec(PatientGenNames.Vacc_Date)
                    If IsDate(GenNameRec(PatientGenNames.Vacc_NextDue)) Then .Fields("Vacc_NextDue") = GenNameRec(PatientGenNames.Vacc_NextDue)
                    If Not GenNameRec(PatientGenNames.VAcc_Current) = "" Then .Fields("Vacc_Current") = IIf(UCase$(GenNameRec(PatientGenNames.VAcc_Current)) = "FALSE", False, True)
                    If Not GenNameRec(PatientGenNames.Vacc_Exempt) = "" Then .Fields("Vacc_Exempt") = GenNameRec(PatientGenNames.Vacc_Exempt)
                    If IsDate(GenNameRec(PatientGenNames.ExemptionExpirationDate)) Then .Fields("ExemptionExpirationDate") = GenNameRec(PatientGenNames.ExemptionExpirationDate)
                    If Not GenNameRec(PatientGenNames.result) = "" Then .Fields("Result") = GenNameRec(PatientGenNames.result) & vbNullString
                  .Update
                End With
                
               End If
              
            Case Else
               '  Skip this Generic Name
        End Select
    End If
  Next gn
  If Not rst.BOF Or Not rst.EOF Then rst.MoveFirst
  Set BuildGenericRecordSet = rst
exitRoutine:
  Exit Function
haveerror:
 ErrMsgBox "Patient_Imm_BuildGenericReocrdSet", Error$, Err
 Resume exitRoutine
 Resume
End Function

Private Function BuildFinalGenericRecord(ByVal gn As String, VacPatRec() As String)

  '...Build final Generic Record
  Dim GenNameRec() As String
  
  On Error GoTo Func_Err
  
  GenNameRec = Split(obd_PatientsGenNames.Item(gn), FieldSepChar)
  
          GenNameRec(PatientGenNames.GenericName) = CStr(gn)
          GenNameRec(PatientGenNames.VaccineID) = VacPatRec(vacrsfields.VaccineID)
          GenNameRec(PatientGenNames.Vacc_Series) = VacPatRec(vacrsfields.Vacc_Series)
          GenNameRec(PatientGenNames.Vacc_Date) = VacPatRec(vacrsfields.Vacc_Date)
          GenNameRec(PatientGenNames.vacc_Nextseries) = VacPatRec(vacrsfields.vacc_Nextseries)
          GenNameRec(PatientGenNames.Vacc_NextDue) = IIf(IsDate(VacPatRec(vacrsfields.Vacc_NextDue)), VacPatRec(vacrsfields.Vacc_NextDue), vbNullString)
          GenNameRec(PatientGenNames.VAcc_Current) = VacPatRec(vacrsfields.VAcc_Current)
          GenNameRec(PatientGenNames.Vacc_Exempt) = VacPatRec(vacrsfields.Vacc_Exempt)
          GenNameRec(PatientGenNames.ExemptionExpirationDate) = VacPatRec(vacrsfields.ExemptionExpirationDate)
          GenNameRec(PatientGenNames.result) = VacPatRec(vacrsfields.result)
          GenNameRec(PatientGenNames.CurrentStatus) = VacPatRec(vacrsfields.EndStatus)
          
   obd_PatientsGenNames.Item(gn) = Join(GenNameRec, FieldSepChar)

Func_Exit:
  On Error Resume Next
  Exit Function

Func_Err:
 ErrMsgBox "Patient_Imm.BuildFinalGenericRecord", Error$, Err
 Resume Func_Exit
 Resume 'Debug

End Function

Private Sub LoadPatientsIMMHistory()

  '...Load Patient's Vaccination History
  
  On Local Error GoTo haveerror
  
  Dim sSQL As String
  Dim GenRec() As String
  Dim HistoryGN As Dictionary
  Dim vacidfld As field
  Dim VacIDSForGenName() As String
  Dim VacID As Variant
  Dim VacKeyStr As String
  Dim gn As String
  Dim VacFields() As String
  
  sSQL = "SELECT IMMRelNCID, Vacc_Date, " & _
         "VaccineID, Vacc_Series,  " & _
         "Vacc_Exempt, ExemptionExpirationDate, " & _
         "RXN_Size, Result, " & _
         "Mfg_Code, Lot_Nbr "
  sSQL = sSQL & "FROM Imm_People_Vaccine "
  sSQL = sSQL & "WHERE SPON_SSAN_FMP ='" & SSN & "' ORDER BY IMMRelNCID DESC, Vacc_Date DESC, VaccineID DESC"
'  Set PatVaccineHistory = gobjSQL.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  Set PatVaccineHistory = gobjSQL.OpenRecordset(sSQL, adOpenStatic, adLockReadOnly)
  
  '...Establish Generic Name Record
  '...for any generic name not defined already
  '...for this patient
  ReDim GenRec(PatientGenNames.NumFlds)
  
  '...If a generic Name does not exists for a patient
  '...Then that immunization was not defined in any
  '...of the groups associated with unit/user
  '...FOR LIABILITY REASONS I will force this
  '...generic name to be age dependent
  GenRec(PatientGenNames.AgeIsAFactor) = "Y"
  
  If Not PatVaccineHistory.EOF Then
    Set HistoryGN = New Dictionary
    Set vacidfld = PatVaccineHistory.Fields("VaccineID")
  End If
  '...Save/Replace Generic Vacination Names associated
  '...with this Patient's Vacination History
  '...Goes into Patient's Generic Vaciantaion Dictionary
  Do Until PatVaccineHistory.EOF
   'Take First Generic Name in each Generic Group
   VacKeyStr = CStr(vacidfld.Value)
   gn = IMM_Tables.Get_A_Generic_Name(VacKeyStr)
   
   '------------------------------------------------
   
      ' If you want to trap for a specific generic
      ' name.....DO IT HERE
   
'   Debug.Assert LCase(gn) <> "ippd"

   '------------------------------------------------

   '...Have we already processed this Generic Name
   
   '---Potential Problem:  SCR 18678
   
   '   If Two vaccinations are given with the same generic name and
   '   both generic names have the same series but different VacID's
   '   then the most recent VacId given will be saved and the older VacID
   '   will be deleted from the inverted dictionary obd_PatientsVaccines.
   '   The fact that they have the same series is not reaaly the issue.  The
   '   fact that they have different VacID's is the issue.
   
   '  ___Solution:  Concantenate VacID to GN in dictionary HistoryGN. Then at
   '     the end of PatVaccineHistory.EOF go through HistoryGN deleteing only
   '     the ones from obd_PatientsVaccines that don't appear in HistoryGN.
   
   '  Result:  When doing edits on Vaccinations through the History Screen an
   '           invalid Data Message stating the VacID does not exist will go away.
   
   If Not HistoryGN.Exists(gn) Then
        '...No
        HistoryGN.Add gn, gn
        If obd_PatientsVaccines.Exists(VacKeyStr) Then
         VacFields = GetPatientsLastVaccineRecord(VacKeyStr)
         
    '     '...No Inlude vaccines in History file that are not defined in Groups/Units
        Else
         ReDim VacFields(vacrsfields.numFields - 1)
         VacFields(vacrsfields.GenericName) = gn
         VacFields(vacrsfields.VaccineID) = VacKeyStr
         VacFields(vacrsfields.idcontrol) = CStr(DefIDControl.history)
         obd_PatientsVaccines.Add VacKeyStr, Join(VacFields, FieldSepChar)
         
         If Not obd_PatientsGenNames.Exists(gn) Then
           Debug.Assert gn <> "ANTHRAX"
           GenRec(PatientGenNames.GenericName) = gn
           GenRec(PatientGenNames.idcontrol) = CStr(Val(GenRec(PatientGenNames.idcontrol)) Or DefIDControl.history)    ' Indicate defined by History file
           obd_PatientsGenNames.Add gn, Join(GenRec, FieldSepChar)
         End If
         
        End If
        '...indicate old or in progress
        VacFields(vacrsfields.begstatus) = "H"
        With PatVaccineHistory
          VacFields(vacrsfields.Vacc_Date) = CStr(.Fields("Vacc_Date").Value & vbNullString)
          VacFields(vacrsfields.Vacc_Series) = CStr(.Fields("Vacc_Series").Value & vbNullString)
          VacFields(vacrsfields.Vacc_Exempt) = CStr(.Fields("Vacc_Exempt").Value & vbNullString)
          VacFields(vacrsfields.ExemptionExpirationDate) = CStr(.Fields("ExemptionExpirationDate") & vbNullString)
          VacFields(vacrsfields.result) = .Fields("Result") & vbNullString
        End With
        obd_PatientsVaccines.Item(VacKeyStr) = Join(VacFields, FieldSepChar)
        obd_PatientsHistory.Add VacKeyStr, gn
        
        '...Get all VacIDs for this Generic Name
        VacIDSForGenName = IMM_Tables.VacIDsFromGenName(gn)
        
        If UBound(VacIDSForGenName) <> -1 Then
            '...Remove All other same Generic Vaccines from PatientsHistory
            For Each VacID In VacIDSForGenName
              '...If same Vac ID skip it
              If VacID <> VacKeyStr Then
                If obd_PatientsVaccines.Exists(CStr(VacID)) Then obd_PatientsVaccines.Remove CStr(VacID)
              End If
            
            Next VacID
        Else
          '  VacId Exists in History File but Generic Name is not
          '  Defined in Generic name table
          MsgBxARMd Disclaimer & _
                 "       VacID:  " & IIf(Len(VacKeyStr) = 0, "VOID", VacKeyStr) & vbCrLf & _
                 "Generic Name:  " & IIf(Len(gn) = 0, "VOID", gn) & vbCrLf & vbCrLf & _
                 "Is not defined in the" & Name_Of_Generic_Vac_Names_Table & "File.", vbExclamation + vbOKOnly, "Invalid Generic Name"
        End If
        
   End If
   PatVaccineHistory.MoveNext
  Loop
  Set PatVaccineHistory.ActiveConnection = Nothing
exitRoutine:
  Exit Sub
  
haveerror:
 ErrMsgBox "Patient_Imm_LoadPatientsIMMHistory", Error$, Err
 Resume exitRoutine
 Resume
End Sub

Public Function CheckForExempt(NextDate As Variant, PatVacData() As String, ByVal EvaluationDate As Date, VacSeries() As String) As String
  Dim Exemptflag As String
  Dim ASeries() As String
  Dim TempDate As Date
  ' SCR 16494  T.Noren  09/06/01
  Dim NoMoreSeries As Boolean
  
  On Error GoTo Func_Err
  
  Select Case PatVacData(vacrsfields.Vacc_Exempt)
     Case "0", "5", "99"
     Case Else
      Exemptflag = SayExemptionText(PatVacData(vacrsfields.Vacc_Exempt))
  End Select
  
  If LenB(Exemptflag) <> 0 Then
      ' SCR 16494  T.Noren  09/06/01
      If Val(PatVacData(vacrsfields.Vacc_Series)) > UBound(VacSeries) Then NoMoreSeries = True
      If IsDate(PatVacData(vacrsfields.ExemptionExpirationDate)) Then
       '...Yes...Are we past the exemption date
       If EvaluationDate > PatVacData(vacrsfields.ExemptionExpirationDate) Or PatVacData(vacrsfields.ExemptionExpirationDate) = #12/31/9999# Then
        ' SCR 16494  T.Noren  09/06/01
        If NoMoreSeries Then
         NextDate = -1
        Else
          '...Yes caluclate earliest possible time
          '...Assumption that the date for VacRSFields.vacc_Date is a Valid Date
          '...If it is not DO NOT KNOW HOW TO CALCULATE NEXT DUE DATE
          ASeries = Split(VacSeries(Val(PatVacData(vacrsfields.Vacc_Series))), FieldSepChar)
          NextDate = CalcAgeOrTime(PatVacData(vacrsfields.Vacc_Date), ASeries(VacDictGrp2Fields.MinInterval))
        End If
        If PatVacData(vacrsfields.ExemptionExpirationDate) <> #12/31/9999# Then
         '...Calculate date based on 1 day past exempt period
         TempDate = DateAdd("d", 1, PatVacData(vacrsfields.ExemptionExpirationDate))
         '...Make sure that the correct date is used
         If TempDate < NextDate Then NextDate = TempDate
        End If
       End If
      Else
        ' SCR 16494  T.Noren  09/06/01
        If NoMoreSeries Then
         NextDate = -1
        Else
          '...Assumption that the date for VacRSFields.vacc_Date is a Valid Date
          '...If it is not DO NOT KNOW HOW TO CALCULATE NEXT DUE DATE
          ASeries = Split(VacSeries(Val(PatVacData(vacrsfields.Vacc_Series))), FieldSepChar)
          NextDate = CalcAgeOrTime(PatVacData(vacrsfields.Vacc_Date), ASeries(VacDictGrp2Fields.MinInterval))
        End If
      End If
  End If
                
  CheckForExempt = Exemptflag
    
Func_Exit:
  On Error Resume Next
  Exit Function

Func_Err:
 ErrMsgBox "Patient_Imm.CheckForExempt", Error$, Err
 Resume Func_Exit
 Resume 'Debug
    
End Function

Public Function Recalc()

  On Error GoTo haveerror
  
  '...Load Patient's Vaccination History...
  '...Indicate get all vaccines in History File
  LoadPatientsIMMHistory

  
  '...Calculate Next due date base on Series and cycles only
  CalculateNextDueDate Now
  
  '-----------------------------------------------------------------------------------
     ' Code goes here for removing Vaccines due to
     ' Conditions (Expert System) entered through GUI
  '-----------------------------------------------------------------------------------
  
  '...Build Patient Generic IMM Schedule base on Patient's Vaccine Table
  PatientGenericIMMSchedule
  
  '-----------------------------------------------------------------------------------
     ' Debug Print out Dictionaries
     
     Call dump_obd_Patientsvaccines
     
  '-----------------------------------------------------------------------------------
  
  
  '...Build Record set
  Set Recalc = BuildGenericRecordSet(True)
  
exitRoutine:
  
  Exit Function
  
haveerror:
'  Debug.Print Err.Number, Err.Description
'  Resume                        ' TJNXXX
  lLastError = Err.Number
  sLastErrorDesc = "Patient_IMM: Recalc: " & Err.Description
  Err.Raise lLastError, "Patient_IMM: Recalc: ", sLastErrorDesc
End Function

Public Sub Reset()
  
  '  Clear Previuos Patients Dictionary
  
  Set obd_PatientsHistory = New Dictionary
  Set obd_PatientsGenNames = New Dictionary
  Set obd_PatientsVaccines = New Dictionary
  Set obd_PatientsGroups = New Dictionary
  obd_PatientsGroups.CompareMode = TextCompare
  obd_PatientsGenNames.CompareMode = TextCompare
  obd_PatientsVaccines.CompareMode = TextCompare
  obd_PatientsHistory.CompareMode = TextCompare
End Sub

Public Function VaxNextDueADD(ByVal xVaxID As String, _
                           Optional ByVal vSeries As Variant = vbNull, _
                           Optional ByVal vDate As Variant = vbNullString, _
                           Optional ByVal iExempt As Integer = 0, _
                           Optional ByVal vExemptDate As Variant = vbNullString, _
                           Optional ByVal AgeDependent As Boolean = False, _
                           Optional ByVal NextDateType As Long = NextDateTypeCalc.StatusAsOf, _
                           Optional ByVal SaveNewRecord As Boolean = False) As Variant

 ' Given a Vaccines ID this routine will calculate the
 ' Next due date.
 
 ' It can perform 2 types of calculations:
 
 '     1. If NextDateType is StatusAsOf then _
          NextDue Date will be based on _
          prior immunizations.
          
 '     2. If NextDateType is AddASeries then _
          NextDue Date will be based on _
          vSeries + 1.  If the vSeries is the _
          end of the series NextDue Date will _
          be Null
          
'      3. If NextDateType is DeleteASeries then _
          this functionality needs to be added.
          
' This routine will also invoke some artifical intelligence:

'     If the VaxId does not exists in the Patients Record it _
      will add it.
      
'     It will determine if the VaxId is in a group that can _
      alter the series through Combination immunizations.  If _
      so it will determine the proper series number.
          
On Local Error GoTo haveerror

 Dim PatGenRec() As String
 Dim GenRec() As String
 Dim VacFields() As String
 Dim gn As String
 Dim CreateNewRecord As Boolean
 Dim CreatedNewVacRecord As Boolean
 Dim CreatedNewGenRecord As Boolean
 Dim NextKey As Integer
 Dim HoldVac() As String
 Dim HoldGen() As String
 Dim answer As VbMsgBoxResult
 Dim VaxId As String
 
  '...Get Generic Name of VacID
  If Potion.VaxId = "" Then
   gn = IMM_Tables.Get_A_Generic_Name(xVaxID)
  Else
   '...The Following line was commented out so that
   '...We would use the generic name of the potion
'   gn = Potion.ChosenGenName
   If gn = "" Then gn = IMM_Tables.Get_A_Generic_Name(VaxId)
  End If
  
  '...Determine Last VacID
  NextKey = GetHighNewVacKey(xVaxID)
  If NextKey = 0 Then
    VaxId = xVaxID
  Else
   VaxId = CreateVacID(xVaxID, NextKey)
  End If
  
 '...Does this Vaccine exists in Patients Data
 If Not obd_PatientsVaccines.Exists(VaxId) Then
  ' No
  '...Add the record to the Patient
  
    '...Establish Fields for Vaccine Record
    ReDim VacFields(vacrsfields.numFields - 1)
    GenRec = IMM_Tables.Get_A_Generic_Name_Record(gn)
    VacFields(vacrsfields.GenericName) = gn
    VacFields(vacrsfields.VaccineID) = VaxId
    '...IF generic record doesn't exist then this is a New Entry
    '...Else This is an add on to a series and consequently prior series exist on
    '...History file
    VacFields(vacrsfields.begstatus) = IIf(obd_PatientsGenNames.Exists(gn), "H", "N")
    
    VacFields(vacrsfields.Vacc_Exempt) = CStr(iExempt)     ' set Exemption
    VacFields(vacrsfields.ExemptionExpirationDate) = CStr(vExemptDate)
    VacFields(vacrsfields.IgnoreAge) = IIf(AgeDependent, "N", "Y")
  '...Is patient already using a common generic name
  If obd_PatientsGenNames.Exists(gn) Then
    ' Yes
    VacFields(vacrsfields.idcontrol) = DefIDControl.Alternate
    PatGenRec = Split(obd_PatientsGenNames.Item(gn), FieldSepChar)
    HoldGen = PatGenRec
    'Populate Immunization data with Data from Last Immunization
    VacFields(vacrsfields.ExemptionExpirationDate) = PatGenRec(PatientGenNames.ExemptionExpirationDate)
    VacFields(vacrsfields.Vacc_Series) = PatGenRec(PatientGenNames.Vacc_Series)
    VacFields(vacrsfields.Vacc_Exempt) = PatGenRec(PatientGenNames.Vacc_Exempt)
    VacFields(vacrsfields.Vacc_Date) = PatGenRec(PatientGenNames.Vacc_Date)
  Else
    VacFields(vacrsfields.idcontrol) = DefIDControl.NextSeries
    ReDim PatGenRec(PatientGenNames.NumFlds)
    PatGenRec(PatientGenNames.AgeIsAFactor) = "N"
    PatGenRec(PatientGenNames.GenericName) = gn
    PatGenRec(PatientGenNames.VaccineID) = VaxId
    PatGenRec(PatientGenNames.idcontrol) = CStr(Val(PatGenRec(PatientGenNames.idcontrol)) Or DefIDControl.NextSeries)
    Debug.Assert gn <> "ANTHRAX"
    obd_PatientsGenNames.Add gn, Join(PatGenRec, FieldSepChar)
    CreatedNewGenRecord = True
  End If
  obd_PatientsVaccines.Add VaxId, Join(VacFields, FieldSepChar)
  CreatedNewVacRecord = True
 Else
  ' Patients VacID record exists
  PatGenRec = Split(obd_PatientsGenNames.Item(gn), FieldSepChar)
  ' Save original Generic Data
  HoldGen = PatGenRec
  VacFields = GetPatientsLastVaccineRecord(VaxId)
  ' Save Original Vaccine Data
  HoldVac = VacFields
  VacFields(vacrsfields.idcontrol) = DefIDControl.NextSeries
  If PatGenRec(PatientGenNames.Vacc_Series) = vbNullString Then
    VacFields(vacrsfields.Vacc_Series) = "0"
  Else
    VacFields(vacrsfields.Vacc_Series) = PatGenRec(PatientGenNames.Vacc_Series)
  End If
  VacFields(vacrsfields.begstatus) = "H"
  VacFields(vacrsfields.Vacc_Exempt) = CStr(iExempt)     ' set Exemption
  VacFields(vacrsfields.ExemptionExpirationDate) = CStr(vExemptDate)
  VacFields(vacrsfields.Vacc_Date) = CStr(vDate)
  
'SCR 18898 - Brian Mowbray 11/18/2001
'This  allows the nextdue date to calculate correctly when adding an IPPD
  VacFields(vacrsfields.result) = vbNullString
'**

  If PatGenRec(PatientGenNames.Vacc_Series) <> vSeries And (PatGenRec(PatientGenNames.Vacc_Series) <> "" And vSeries <> 0) Then
    answer = MsgBxARMd("Last Generic Series #: " & CStr(PatGenRec(PatientGenNames.Vacc_Series)) & vbCrLf & _
           "New Series added was : " & CStr(vSeries) & vbCrLf & vbCrLf & _
           "This is not a proper entry." & vbCrLf & vbCrLf & _
           "Do you want to continue?", vbQuestion + vbYesNo, "Verifying Input")
    If answer = vbNo Then Exit Function
  End If
  VacFields(vacrsfields.Vacc_Series) = vSeries
  
 End If
 
    If IsDate(vDate) Then
    
      If IsDate(VacFields(vacrsfields.Vacc_Date)) Then
       VacFields(vacrsfields.Vacc_Date) = vDate
       If Not CreatedNewVacRecord Then CreateNewRecord = True
      Else
       VacFields(vacrsfields.Vacc_Date) = vDate
      End If
      
    Else
      MsgBxARMd "Missing a Date", vbExclamation + vbOKOnly, "Invalid Input"
      If CreatedNewGenRecord Then obd_PatientsGenNames.Remove gn
      If CreatedNewVacRecord Then obd_PatientsVaccines.Remove VaxId
      Exit Function
      
    End If
    
    '...Check for Last Series Input
    If IsNumeric(vSeries) Then
    
     If IsNumeric(VacFields(vacrsfields.Vacc_Series)) _
          Or _
        (VacFields(vacrsfields.Vacc_Series) = "" And CInt(vSeries) = 0) Then
        
      If (CInt(vSeries) = CInt("0" & VacFields(vacrsfields.Vacc_Series))) _
           And VacFields(vacrsfields.begstatus) = "H" _
            Or _
        ( _
          (CInt(vSeries) = 0 And VacFields(vacrsfields.Vacc_Series) = vbNullString) _
          And VacFields(vacrsfields.begstatus) = "N" _
        ) _
            Then
'        VacFields(vacrsfields.Vacc_Series) = CInt(vSeries)
        If Not CreatedNewVacRecord Then
          CreateNewRecord = True
          VacFields(vacrsfields.Vacc_Series) = VacFields(vacrsfields.Vacc_Series) + 1
        End If
      Else
        MsgBxARMd "Series Input: " & CStr(vSeries) & vbCrLf & _
               "is not greater than last series given (" & VacFields(vacrsfields.Vacc_Series) & ") on record." & vbCrLf, _
               vbExclamation + vbOKOnly, "Invalid Input"
        If CreatedNewGenRecord Then obd_PatientsGenNames.Remove gn
        If CreatedNewVacRecord Then obd_PatientsVaccines.Remove VaxId
        Exit Function
      
      End If
      
     Else
        MsgBxARMd "Invalid Series #: " & CStr(vSeries) & vbCrLf & _
               "Vac Rec Series #: " & VacFields(vacrsfields.Vacc_Series) & vbCrLf & _
               "Beginning Status: " & VacFields(vacrsfields.begstatus) * vbCrLf & _
               "Last Vac Date: " & VacFields(vacrsfields.Vacc_Date), vbExclamation + vbOKOnly, "Invalid Input:"
        Exit Function
     End If
     
    Else
    
        MsgBxARMd "Series Input: " & CStr(vSeries) & vbCrLf & _
               "Last Series (" & VacFields(vacrsfields.Vacc_Series) & ") on record." & vbCrLf, _
               vbExclamation + vbOKOnly, "Invalid Input"
        If CreatedNewGenRecord Then obd_PatientsGenNames.Remove gn
        If CreatedNewVacRecord Then obd_PatientsVaccines.Remove VaxId
        Exit Function
     
    End If
    
    '...How do we update Vaccine status
    If Not CreatedNewVacRecord And CreateNewRecord Then
      '...Create New Record
      '...Develop New Key
      NextKey = GetHighNewVacKey(VaxId)
      
      '...Build New Key  Format   VaxID | NEW | <Index>
      VaxId = CreateVacID(VaxId, NextKey + 1)
      obd_PatientsVaccines.Add VaxId, Join(VacFields, FieldSepChar)
      CreatedNewVacRecord = True
    Else
      obd_PatientsVaccines.Item(VaxId) = Join(VacFields, FieldSepChar)
    End If
 
 VaxNextDueADD = ADueDate(VaxId, vDate, NextDateType)
 
 '...do we get rid of newly created records
 If Not SaveNewRecord Then
 
    If CreatedNewVacRecord Then
      obd_PatientsVaccines.Remove VaxId
    Else
      '...Restore Original Values
      obd_PatientsVaccines.Item(VaxId) = Join(HoldVac, FieldSepChar)
    End If
    If CreatedNewGenRecord Then
      obd_PatientsGenNames.Remove gn
    Else
      obd_PatientsGenNames.Item(gn) = Join(HoldGen, FieldSepChar)
    End If
 Else
 
    VacFields = GetPatientsLastVaccineRecord(VaxId)
    '...This is now History
    VacFields(vacrsfields.begstatus) = "H"
    VacFields(vacrsfields.Vacc_Series) = "1"
    VacFields(vacrsfields.vacc_Nextseries) = "2"
    obd_PatientsVaccines.Item(VaxId) = Join(VacFields, FieldSepChar)
    '...See If this effects other Vaccines
    VaccinesEffected True, VaxId
    '...Show in Generic Data for Patient
    UpdateVactoGenForPatient gn

 End If
 
exitRoutine:
 Exit Function
haveerror:
 Resume exitRoutine
 Debug.Print Err.Number, Err.Description
 Resume   ' TJNXXX
End Function
Public Function VaxNextDueEDIT(ByVal VaxId As String, _
                           Optional ByVal vSeries As Variant = vbNull, _
                           Optional ByVal vDate As Variant = vbNullString, _
                           Optional ByVal iExempt As Integer = 0, _
                           Optional ByVal vExemptDate As Variant = vbNullString, _
                           Optional ByVal result As String = vbNullString, _
                           Optional ByVal AgeDependent As Boolean = False, _
                           Optional ByVal NextDateType As Long = NextDateTypeCalc.StatusAsOf, _
                           Optional ByVal SaveNewRecord As Boolean = False) As Variant

 ' Given a Vaccines ID this routine will calculate the
 ' Next due date.
 
 ' It can perform 2 types of calculations:
 
 '     1. If NextDateType is StatusAsOf then _
          NextDue Date will be based on _
          prior immunizations.
          
 '     2. If NextDateType is AddASeries then _
          NextDue Date will be based on _
          vSeries + 1.  If the vSeries is the _
          end of the series NextDue Date will _
          be Null
          
'      3. If NextDateType is DeleteASeries then _
          this functionality needs to be added.
          
' This routine will also invoke some artifical intelligence:

'     If the VaxId does not exists in the Patients Record it _
      will add it.
      
'     It will determine if the VaxId is in a group that can _
      alter the series through Combination immunizations.  If _
      so it will determine the proper series number.
          
On Local Error GoTo haveerror

 Dim PatGenRec() As String
 Dim VacFields() As String
 Dim gn As String
 Dim HoldVac() As String
 Dim HoldGen() As String
 
  '...Get Generic Name of VacID
  gn = IMM_Tables.Get_A_Generic_Name(VaxId)
 
 '...Does this Vaccine exists in Patients Data
 If Not obd_PatientsVaccines.Exists(VaxId) Then
 
   '...Invalid situation
   '...Not editing last VacID for this generic name
   ' To solve SCR 18678 then following message box was initially commented out
   ' to ignore editing historical data (not most recent vaccine in series)
   ' However a flag HistoricalRecord in frmvacc_single is being used to filter
   ' Historical edits.  Consequently the following code should never be executed
   ' unless some other error exists.  Do not comment out the MsgBxARMd but solve the
   ' problem elsewhere where the problem exists.
   
   MsgBxARMd "Vaccine ID:  " & VaxId & vbCrLf & vbCrLf & _
          "COULD NOT BE LOCATED", vbExclamation + vbOKOnly, "Invalid Situation"
          VaxNextDueEDIT = vbNullString
          Exit Function
 Else
 
  PatGenRec = Split(obd_PatientsGenNames.Item(gn), FieldSepChar)
  HoldGen = PatGenRec
  VacFields = GetPatientsLastVaccineRecord(VaxId)
  HoldVac = VacFields
  VacFields(vacrsfields.Vacc_Exempt) = CStr(iExempt)     ' set Exemption
  VacFields(vacrsfields.ExemptionExpirationDate) = CStr(vExemptDate)
  VacFields(vacrsfields.Vacc_Date) = CStr(vDate)
  VacFields(vacrsfields.Vacc_Series) = vSeries
  VacFields(vacrsfields.result) = result
  obd_PatientsVaccines.Item(VaxId) = Join(VacFields, FieldSepChar)
  
 End If
 
 
 VaxNextDueEDIT = ADueDate(VaxId, Now, NextDateType)
 
 '...do we get rid of newly created records
 If Not SaveNewRecord Then
 
    '...Restore Original Values
    obd_PatientsVaccines.Item(VaxId) = Join(HoldVac, FieldSepChar)
    obd_PatientsGenNames.Item(gn) = Join(HoldGen, FieldSepChar)
    
 Else
 
    '...Show in Generic Data for Patient
    UpdateVactoGenForPatient gn

 End If
 
exitRoutine:
 Exit Function
haveerror:
 Resume exitRoutine
 Debug.Print Err.Number, Err.Description
 Resume   ' TJNXXX
End Function

Public Function VaxNextDue(ByVal VaxId As String, _
                           Optional ByVal vSeries As Variant = vbNull, _
                           Optional ByVal vDate As Variant = vbNullString, _
                           Optional ByVal iExempt As Integer = 0, _
                           Optional ByVal vExemptDate As Variant = vbNullString, _
                           Optional ByVal AgeDependent As Boolean = False, _
                           Optional ByVal NextDateType As Long = NextDateTypeCalc.StatusAsOf) As Variant

 ' Given a Vaccines ID this routine will calculate the
 ' Next due date.
 
 ' It can perform 2 types of calculations:
 
 '     1. If NextDateType is StatusAsOf then _
          NextDue Date will be based on _
          prior immunizations.
          
On Local Error GoTo haveerror

 If IsNull(vDate) Then vDate = Now
 
 VaxNextDue = ADueDate(VaxId, vDate, NextDateType)
 
 
exitRoutine:
 Exit Function
haveerror:
 Resume exitRoutine
 Debug.Print Err.Number, Err.Description
 Resume   ' TJNXXX
End Function

Public Function ADueDate(ByVal oVacKey As String, _
                         ByVal xEvalDate As Variant, _
                         ByVal NextSeries As Long) As Variant

   On Local Error GoTo haveerror
   
   Dim VacSeries() As String
   Dim ASeries() As String
   Dim BSeries() As String
   Dim Process As Boolean
   Dim NextDate As Variant
   Dim gn As String
   Dim GenData() As String
   Dim Status As String
   Dim ExemptReason As String
   Dim LastSeriesNumber As Integer
   Dim PatVacRec() As String
   Dim VacKey As String
   Dim evaldate As Date
   Dim PatGenRec() As String
   Dim DrugType As GenericTYpeDrug
   
   evaldate = Format$(CDate(xEvalDate), "mm/dd/yy")
   
     PatVacRec = GetPatientsLastVaccineRecord(oVacKey)
     
     '...oVacKey
     '   Format    Vackey & WithinFieldSepChar & AddedSeries
     '             WithinFieldSepChar & AddedSeries is only
     '             Used when a Vaccine Series is added to a
     '             patient's History of Vaccines....
     '             Otherwise VacKey = OVacKey = VacID
     
     If InStr(1, oVacKey, WithinFieldSepChar, vbTextCompare) <= 0 Then
        VacKey = oVacKey
     Else
        VacKey = Mid$(oVacKey, 1, InStr(1, oVacKey, WithinFieldSepChar, vbTextCompare) - 1)
     End If
     gn = PatVacRec(vacrsfields.GenericName)
'     Debug.Assert LCase(gn) <> "mmr"       ' TJNASSERT
     Process = True
     Select Case PatVacRec(vacrsfields.begstatus)
       Case "N"   ' Automatic go with todays date/or Date Entered in Add function
        LastSeriesNumber = 1
        If PatVacRec(vacrsfields.idcontrol) = DefIDControl.NextSeries Then
         '...Adding a New VaxID from the ADD vaccine function
         PatVacRec(vacrsfields.Vacc_Date) = Format$(xEvalDate, "mm/dd/yy")
         PatVacRec(vacrsfields.Vacc_Series) = 1 ' This reflects above date administered
         GoTo Historical         ' This is now looked upon as an Historical record
        Else
         NextDate = Format$(Now, "mm/dd/yy")
        End If
        
        '...KLUGE   KLUGE   KLUGE
        
        '...THIS IS TEMPERARY UNTIL EXPERT SYSTEM CAN BE IMPLIMENTED
        '...IF IMMUNIZATION IS CYCLICAL DON'T SCHEDULE UNLESS IN PROPER PERIOD
        
         GenData = IMM_Tables.Get_A_Generic_Name_Record(gn)
         If GenData(GenVacDictFields.Cycle) <> "" Then
          If GenData(GenVacDictFields.Range) <> "" Then
           Process = GetCycleRange(GenData(GenVacDictFields.Range), evaldate)
          End If
         End If
       Case "H"   ' Analyze Status
        'Are WE Adding a New VaxID from the ADD vaccine function with same Generic Name
        If PatVacRec(vacrsfields.idcontrol) = DefIDControl.Alternate Then
          ' Yes
          '...Load Last Series for this Generic Name
          PatGenRec = Split(obd_PatientsGenNames.Item(gn), FieldSepChar)
          '...Save last Series Number Used
          PatVacRec(vacrsfields.Vacc_Series) = PatGenRec(PatientGenNames.Vacc_Series)
        End If
Historical:
        VacSeries = IMM_Tables.Vaccine_Series(VacKey)
        DrugType = IMM_Tables.GetVaccineType(VacKey)
        '...Do we have a valid Series
        If UBound(VacSeries) <> -1 Then
        
            '...Load last series # administered
            LastSeriesNumber = CInt(PatVacRec(vacrsfields.Vacc_Series))
            
            '...Calculate Index into a Series
            LastSeriesNumber = LastSeriesNumber - 1
            
            '...Is this an invaid entry because
            '...series # is > then lines in Schedule
            If LastSeriesNumber > UBound(VacSeries) Then
              ' YES
'              MsgBxARMd "Current Series #: " & CStr(LastSeriesNumber + 1) & " exceeds Series Defined", vbExclamation + vbOKOnly, "Invalid Series"
'              Exit Function
              ' Assume Vaccine is current
              Status = "C"
              PatVacRec(vacrsfields.VAcc_Current) = "True"
              Process = False
            '...Have we completed this Series
            ElseIf UBound(VacSeries) >= LastSeriesNumber Then
                '...No
                '...Do we have to worry about something being exempted
                ExemptReason = CheckForExempt(NextDate, PatVacRec, evaldate, VacSeries)
                
                '...Are we Exempt
                If LenB(ExemptReason) = 0 Then
                 ' No
                 '...Assumption that the date for VacRSFields.vacc_Date is a Valid Date
                 '...If it is not DO NOT KNOW HOW TO CALCULATE NEXT DUE DATE
                  ASeries = Split(VacSeries(LastSeriesNumber), FieldSepChar)
                  
                  '<< Begin:SCR #44875;   Developer: Brian Mowbray 12/15/2003 01:47 PM
                  BSeries = Split(VacSeries(IIf(IMM_Tables.IsVaccineASeries(VacKey), LastSeriesNumber + 1, LastSeriesNumber)), FieldSepChar)
                  
                  If IsVaccineExemptByAge(DOB, BSeries(VacDictGrp2Fields.AgeFrom), BSeries(VacDictGrp2Fields.AgeTo)) Then
                      NextDate = Null
                  Else
                    NextDate = CalcAgeOrTime(PatVacRec(vacrsfields.Vacc_Date), ASeries(VacDictGrp2Fields.MinInterval))
                  End If
                  '>> End: SCR #44875;
                  
                
                  'Checking to see if vaccine is a test (IPPD, TB test) - Brian Mowbray 10/22/2001
                  If IMM_Tables.IsVaccineATest(VacKey) Then
                    
                    'If result is empty set the next due to null because test is complete
                    'otherwise use the nextdue date calculated from above
                    If Len(PatVacRec(vacrsfields.result) & vbNullString) > 0 Then
                      NextDate = Null
                    End If
                    
                    
                  End If
                
                Else
                
                  '******************************************************************
                  'SCR 70710 - NextDate should be NULL for any Exemption besides 'None'
                  '1-June-2005 - B. Fote
                  NextDate = Null
                  '******************************************************************
                
                End If
                
            '...Completed Series
            Else
              '...Set New Status
              '...Yes Completed Series Look to see if cyclical
              Status = "C"
              GoTo Cycle
            End If
        Else
          
          '...Check for Series
          If IMM_Tables.IsVaccineASeries(VacKey) Then
            Status = "C"
            PatVacRec(vacrsfields.VAcc_Current) = "True"
            GoTo Cycle
          Else
          
            '...Invalid series
            Process = False
            'The user should not get this warning if the vaccine is inactive. 12/13/2001 B. Mowbray
            If Not IMM_Tables.IsVaccineInActive(VacKey) Then
              MsgBxARMd Disclaimer & _
                     "       VacID:  " & IIf(Len(VacKey) = 0, "VOID", VacKey) & vbCrLf & _
                     "Generic Name:  " & IIf(Len(gn) = 0, "VOID", gn) & vbCrLf & _
                     "To correct this problem please have the Enterprise Administrator either add a schedule for this series vaccine or change the vaccine to a single dose or make this vaccine inactive.", vbExclamation + vbOKOnly, "No Series Definition"
            
            
            
            End If
          End If
          
        End If
       Case "C"   ' Vacination Complete
Cycle:
         '...Indicate Patient is Current
         PatVacRec(vacrsfields.VAcc_Current) = "True"
         '...Load Generic Information for this Vaccine
         GenData = IMM_Tables.Get_A_Generic_Name_Record(gn)
         If GenData(GenVacDictFields.Cycle) <> "" Then
          If GenData(GenVacDictFields.Range) <> "" Then
           Process = GetCycleRange(GenData(GenVacDictFields.Range), evaldate)
          End If
          If Process Then
            '...Do we have to worry about something being exempted
            ExemptReason = CheckForExempt(NextDate, PatVacRec, evaldate, VacSeries)
            If LenB(ExemptReason) = 0 Then
               NextDate = CalcAgeOrTime(PatVacRec(vacrsfields.Vacc_Date), GenData(GenVacDictFields.Cycle))
            End If
          End If
         Else
          '...This is not cyclical so don't process again
          Process = False
         End If
       Case Else
         Process = False
         If Mid$(PatVacRec(vacrsfields.EndStatus), 1, 1) <> "I" Then
            MsgBxARMd "Invalid PatientVaccine Status", vbInformation + vbOKOnly, "Processing will Continue"
         End If
     End Select
     '...Are Conditions right to schedule
     If Process Then
       '...suppose to have a valid date
       If IsDate(NextDate) Then
        PatVacRec(vacrsfields.Vacc_NextDue) = CDate(NextDate)
        '...Are We Updateing to next series
        If NextSeries = NextDateTypeCalc.AddASeries Then
         ' Yes
         ' Set Status as Current
             PatVacRec(vacrsfields.vacc_Nextseries) = LastSeriesNumber + 1
             '...Indicate Complete or Current
             Status = "C"
             PatVacRec(vacrsfields.VAcc_Current) = "True"
        Else
         ' No
         '...Set Status Etc. based on Current situation
            If NextDate <= evaldate Then
             If PatVacRec(vacrsfields.begstatus) = "N" Then PatVacRec(vacrsfields.vacc_Nextseries) = LastSeriesNumber
             Status = "S"
            Else
             '...Indicate Complete or Current
             Status = "C"
             PatVacRec(vacrsfields.VAcc_Current) = "True"
            End If
        End If
       Else
         '...See if we are exempt
         If LenB(ExemptReason) = 0 Then
           ' No
           If NextDate <> -1 Then
            Status = "D"
           Else
            Status = "C"
            PatVacRec(vacrsfields.VAcc_Current) = "True"
           End If
         Else
           ' Yes
           ' Show Status as current and put reason in nextdue date
           Status = "C"
           PatVacRec(vacrsfields.VAcc_Current) = "True"
           PatVacRec(vacrsfields.Vacc_NextDue) = ExemptReason
           
         End If
       End If
     Else
        '...Incase complete but not cyclical
        '...or  next series in a Historical entry
        
        If NextSeries = NextDateTypeCalc.AddASeries Then
         '...Indicate current
         If Status = "" Then
           Status = "C"
           PatVacRec(vacrsfields.VAcc_Current) = "True"
         End If
        Else
          '...Indicate Don't Schedule
          If Status = "" Then Status = "D"
        End If
     End If
     If Status = "D" Then PatVacRec(vacrsfields.VAcc_Current) = "False"
     PatVacRec(vacrsfields.EndStatus) = Status
     '...Write data Record back
     obd_PatientsVaccines.Item(oVacKey) = Join(PatVacRec, FieldSepChar)
     ADueDate = NextDate
exitRoutine:
     Exit Function
haveerror:
     Resume exitRoutine
     Debug.Print Err.Number, Err.Description
     Resume     ' TJNXXX
End Function

Public Sub UpdateVactoGenForPatient(ByVal gn As Variant)

  On Local Error GoTo haveerror

  Dim VacPatRec() As String
  Dim VacPatRec2() As String
  Dim vacids() As String
  Dim VacID As String
  Dim VacIDIndex As Integer
  Dim GeneralIndex As Integer
  Dim GeneralID As String
  Dim VacNewIDs() As String
  
'    Debug.Assert InStr(gn, "Anthrax") = 0
    '...Get all vaccines for this Generic Name
    vacids = IMM_Tables.VacIDsFromGenName(CStr(gn))
    
    For VacIDIndex = 0 To UBound(vacids)
    
     VacID = vacids(VacIDIndex)
     
     '...Does ID still exist
     If obd_PatientsVaccines.Exists(VacID) Then
     
       VacNewIDs = GetThisVacIDKeys(VacID)
        '...Load most recent Vaccine
        VacID = VacNewIDs(UBound(VacNewIDs))
        VacPatRec = GetPatientsLastVaccineRecord(VacID)
        
        '<< Begin:SCR #52240;   Developer: Brian Mowbray 04/14/2004 09:35 PM
        'Needed to add check here
        If UBound(VacPatRec) <> -1 Then
      
          Select Case VacPatRec(vacrsfields.EndStatus)
          
              Case "S"      ' Schedule this Vaccine
              
               '...Is this the one to schedule
               If IsDate(VacPatRec(vacrsfields.Vacc_NextDue)) Then
               
                     '...Modify status to...Needs to have Vaccine
                     VacPatRec(vacrsfields.EndStatus) = "V"
                     obd_PatientsVaccines.Item(VacID) = Join(VacPatRec, FieldSepChar)
                     
                     '...Process all other related vaccines (Same Generic Name)
                     For GeneralIndex = VacIDIndex + 1 To UBound(vacids)
                      GeneralID = vacids(GeneralIndex)
                      If obd_PatientsVaccines.Exists(GeneralID) Then
                        VacPatRec2 = GetPatientsLastVaccineRecord(GeneralID)
                        If VacPatRec2(vacrsfields.EndStatus) = "S" Then
                          VacPatRec2(vacrsfields.EndStatus) = "v_"
                          VacPatRec2(vacrsfields.SolvedByVacID) = VacID
                          obd_PatientsVaccines.Item(GeneralID) = Join(VacPatRec2, FieldSepChar)
                        End If
                      End If
                     Next GeneralIndex
                     BuildFinalGenericRecord CStr(gn), VacPatRec
                     Exit For
                     
               End If
               
              Case "D"      ' don't Schedule
                 ' Do Nothing
                 
              Case "C"
                 BuildFinalGenericRecord CStr(gn), VacPatRec
                 
              Case Else
                ' Vaccines Solved by another vaccine
               
          End Select
      End If
         
     End If
     
    Next VacIDIndex
    
exitRoutine:
  Exit Sub
haveerror:
 ErrMsgBox "Patient_Imm_UpdateVactoGenForPatient", Error$, Err
 Resume exitRoutine

End Sub

Public Sub DeletePatVaccine(ByVal Vac As Long)

  '...Delete A Vaccine from a patient's dictionary
  Dim VacID As String
  Dim gn As String
  Dim GenRec() As String
  
  On Error GoTo Sub_Err
  
  VacID = CStr(Vac)

  If obd_PatientsVaccines.Exists(VacID) Then
  
     gn = IMM_Tables.Get_A_Generic_Name(VacID)
     GenRec = Split(obd_PatientsGenNames.Item(gn), FieldSepChar)
     Select Case Val(GenRec(PatientGenNames.idcontrol)) And DefIDControl.AllGroups
        Case Is <> 0
          GetPatientsVaccineSchedule True
        Case Else
          '...If not user Defined or Unit Defined then Delete it
          obd_PatientsGenNames.Remove gn
          obd_PatientsVaccines.Remove VacID
          GetPatientsVaccineSchedule True
     End Select
  
  End If


Sub_exit:
  On Error Resume Next
  Exit Sub

Sub_Err:
 ErrMsgBox "Patient_Imm.DeletePatVaccine", Error$, Err
 Resume Sub_exit
 Resume 'Debug

End Sub

Public Function GetNextGenSeries(ByVal GenName As String) As String

 '...Loads Last Series Given for a generic Series

 Dim VacRec() As String
 
 On Error GoTo Func_Err
 
 If obd_PatientsGenNames.Exists(GenName) Then
 
   VacRec = Split(obd_PatientsGenNames.Item(GenName), FieldSepChar)
   GetNextGenSeries = VacRec(PatientGenNames.vacc_Nextseries)
 End If

Func_Exit:
  On Error Resume Next
  Exit Function

Func_Err:
 ErrMsgBox "Patient_Imm.GetNextGenSeries", Error$, Err
 Resume Func_Exit
 Resume 'Debug

End Function
Public Function GetLastGenSeries(ByVal GenName As String) As Integer

 '...Loads Last Series Given for a generic Series

  Dim VacRec() As String
 
  On Error GoTo Func_Err
  
  If obd_PatientsGenNames.Exists(GenName) Then
  
    VacRec = Split(obd_PatientsGenNames.Item(GenName), FieldSepChar)
    GetLastGenSeries = CInt("0" & VacRec(PatientGenNames.Vacc_Series))
  End If

Func_Exit:
  On Error Resume Next
  Exit Function

Func_Err:
 ErrMsgBox "Patient_Imm.GetLastGenSeries", Error$, Err
 Resume Func_Exit
 Resume 'Debug

End Function

Public Function GetLastVacSeries(ByVal VacID As String) As String

 '...Loads Last Series Given

  Dim VacRec() As String
  Dim newkey As Integer
 
  On Error GoTo Func_Err

  If obd_PatientsVaccines.Exists(VacID) Then
    '...Determine if multiple Vaccines for this VaxID
    newkey = GetHighNewVacKey(VacID)
    
    '...If So create new VacID
    If newkey <> 0 Then VacID = CreateVacID(VacID, newkey)
    ' Load Vac Record
    VacRec = GetPatientsLastVaccineRecord(VacID)
    
    '...Get Last Series Administered
    GetLastVacSeries = "0" & VacRec(vacrsfields.Vacc_Series)
  Else
   GetLastVacSeries = "0"
  End If
  
Func_Exit:
  On Error Resume Next
  Exit Function

Func_Err:
 ErrMsgBox "Patient_Imm.GetLastVacSeries", Error$, Err
 Resume Func_Exit
 Resume 'Debug


End Function

Friend Function GetPatientsLastVaccineRecord(ByVal VacID As String, Optional ByVal GetNewKey As Boolean = True) As String()

  '...Get Patients Last Vaccine Record
 
 Dim newkey As Integer
    
  On Error GoTo Func_Err
   
   If GetNewKey Then
     newkey = GetHighNewVacKey(VacID)
    
    '...If So create new VacID
    If newkey <> 0 Then VacID = CreateVacID(VacID, newkey)
    
   Else
     If Not obd_PatientsVaccines.Exists(VacID) Then
      GetPatientsLastVaccineRecord = SetUboundToMinusOne
      Exit Function
     End If
     
   End If
   
   GetPatientsLastVaccineRecord = Split(obd_PatientsVaccines.Item(VacID), FieldSepChar)
   
Func_Exit:
  On Error Resume Next
  Exit Function

Func_Err:
 ErrMsgBox "Patient_Imm.GetPatientsLastVaccineRecord", Error$, Err
 Resume Func_Exit
 Resume 'Debug
   
End Function

Friend Function GetHighNewVacKey(ByVal VaxId As String) As Integer
 
  '  Find Highest Key in VaxID
  
  '... Format   VaxID | NEW | <Index>

 On Error GoTo Func_Err
 
 Dim vacids() As String
 Dim vKey As Variant
 Dim xloc As Long
 Dim NextKey As Integer

      '...Find all records with this VaxID
      vacids = GetThisVacIDKeys(VaxId)
      
      '...find Highest index
      For Each vKey In vacids
      
       xloc = InStrRev(CStr(vKey), "NEW", , vbTextCompare)
       If xloc <> 0 Then
         xloc = Val(Mid$(CStr(vKey), xloc + 4))
         If NextKey < xloc Then NextKey = xloc
       End If
       
      Next vKey
      
  '...Return Highest Index
  GetHighNewVacKey = NextKey
  
Func_Exit:
  On Error Resume Next
  Exit Function

Func_Err:
 ErrMsgBox "Patient_Imm.GetHighNewVacKey", Error$, Err
 Resume Func_Exit
 Resume 'Debug

End Function

Public Function CreateVacID(ByVal VaxId As String, ByVal NewID As Integer) As String

   '  Build new Key for additional vaccines added to obd_PatientsVaccines Dictionary
   
   Dim xloc As Integer
   
   On Error GoTo Func_Err
   
   xloc = InStr(VaxId, WithinFieldSepChar)
   If xloc <> 0 Then VaxId = Left(VaxId, xloc - 1)
   CreateVacID = VaxId & WithinFieldSepChar & "NEW" & WithinFieldSepChar & CStr(NewID)
   
Func_Exit:
  On Error Resume Next
  Exit Function

Func_Err:
 ErrMsgBox "Patient_Imm.CreateVacID", Error$, Err
 Resume Func_Exit
 Resume 'Debug
   
End Function

Public Function CalcNextSeries(ByVal VaxId As String) As String

  ' This function tries to determine the next series number
  ' given a VaxID

 Dim gn As String
 Dim GenRec() As String
 Dim VacRec() As String
 Dim relrec() As String
 Dim Count As Integer
 Dim gnstr As String
 Dim gn2 As String
 Dim GenNames() As String
 
 On Error GoTo Func_Err
 
 '...Get Generic Name of VacID
 gn = IMM_Tables.Get_A_Generic_Name(VaxId)
 
 '...Generic Name for patient
 GenRec = IMM_Tables.Get_A_Generic_Name_Record(gn)
  
 '...Is this a combination Vaccine
 If GenRec(GenVacDictFields.Combination) = "1" Then
  ' Yes
  '...Indicate combination being used
  Potion.VaxId = VaxId
  '...Load Vaccine Record
  VacRec = IMM_Tables.Vaccine_BasicData(VaxId)
  Potion.VacEffected = VacRec(VacDictGrp1Fields.EffectsVacID)
  '<< Begin:SCR #36665;   Developer: Brian Mowbray 05/30/2003 04:36 PM
  If Len(Potion.VacEffected) > 0 Then
    '...Load all Vaccines this Vaccine relates too
    relrec = Split(Mid$(VacRec(VacDictGrp1Fields.EffectsVacID), 2, Len(VacRec(VacDictGrp1Fields.EffectsVacID)) - 2), WithinFieldSepChar)
    '...Determine the number of different Generic Names
    For Count = 0 To UBound(relrec)
      gn2 = IMM_Tables.Get_A_Generic_Name(relrec(Count))
      If InStr(gnstr, gn2) = 0 Then
        gnstr = gnstr & WithinFieldSepChar & gn2
      End If
    Next Count
    '...Different Generic Names Stored in GenNames
    GenNames = Split(Mid$(gnstr, 2), WithinFieldSepChar)
  End If
 '>> End: SCR #36665;
 Else
  '...Indicate no combination being used
  Potion.VaxId = ""
 End If
 
 If obd_PatientsGenNames.Exists(gn) Then
   ' Patient has had this Generic Name before
   GenRec = Split(obd_PatientsGenNames.Item(gn), FieldSepChar)

  'SCR 18898 - Brian Mowbray 11/18/2001
  'This  allows the nextdue date to calculate correctly when adding an IPPD
  
  '*********************************************************
  'SCR 79484 - The Series for Influenza should always be "1"
  'The VaccineID for Influenza is 16, 15, 88
  '*********************************************************
  
   If IMM_Tables.IsVaccineATest(VaxId) Or VaxId = "15" Or VaxId = "16" Or VaxId = "88" Then
    CalcNextSeries = "1"
   Else
    CalcNextSeries = CInt("0" & GenRec(PatientGenNames.Vacc_Series)) + 1
   End If
  '**
 ElseIf Potion.VacEffected <> "" Then
''---------------------------------------------------------------------------------
   '...Is there only one generic name
'   If UBound(GenNames) = 0 Then
'    ' Yes
'    '...Store generic name effected
'    Potion.ChosenGenName = GenNames(0)
'    If obd_PatientsGenNames.Exists(GenNames(0)) Then
'      ' Patient has had this Generic Name before
'      GenRec = Split(obd_PatientsGenNames.Item(GenNames(0)), FieldSepChar)
'      CalcNextSeries = GenRec(PatientGenNames.Vacc_Series) + 1
'    Else
'      '...Assume first in Series
'      CalcNextSeries = "1"
'    End If
'   Else
'    ' No
'    ' Need to show all generic names and have user select the one this will be
'    ' used for.....FOR LOOK FOR First ONE THAT ALREADY EXISTS
'    GenKeys = obd_PatientsGenNames.Keys
'    For Count = 0 To UBound(GenKeys)
'      For Count2 = 0 To UBound(GenNames)
'        If GenNames(Count2) = GenKeys(Count) Then
'         Potion.ChosenGenName = GenKeys(Count)
'         foundit = True
'         Exit For
'        End If
'      Next Count2
'      If foundit Then Exit For
'    Next Count
'    If foundit Then
'      GenRec = Split(obd_PatientsGenNames.Item(Potion.ChosenGenName), FieldSepChar)
'      CalcNextSeries = GenRec(PatientGenNames.Vacc_Series) + 1
'    End If
'   End If
''--------------------------------------------------------------------------------
  ' No
  ' First in Series
  CalcNextSeries = "1"
 Else
  ' No
  ' First in Series
  CalcNextSeries = "1"
 End If
 
Func_Exit:
  On Error Resume Next
  Exit Function

Func_Err:
 ErrMsgBox "Patient_Imm.CalcNextSeries", Error$, Err
 Resume Func_Exit
 Resume 'Debug
End Function

Public Function GetThisVacIDKeys(ByVal VacID As String) As String()

   ' this finds base vacid and all newly added vacid
   
   '  I.E.   vacid,vacid | New | 1,vacid | New | 2,   ....etc
 
  Dim vackeys() As Variant
  Dim VacNewIDs() As String
  Dim Count As Integer
  Dim lenv As Integer
  Dim hd As Dictionary
  Dim newkey As String
  Dim thiskey As String
  Dim xkey As Variant
  
  
  On Error GoTo Func_Err
    
    Set hd = New Dictionary
  
        ' Get all the vaccine keys
        vackeys = obd_PatientsVaccines.Keys
        
        ' Get the specific keys for this exact VacID
        VacNewIDs = Filter(vackeys, VacID, True, vbTextCompare)
        
        newkey = VacID & WithinFieldSepChar
        
        lenv = Len(VacID)
        
        For Count = 0 To UBound(VacNewIDs)
        
         thiskey = VacNewIDs(Count)
         
         If Len(thiskey) > lenv Then
          If newkey = Left$(thiskey, lenv + 1) Then hd.Add thiskey, thiskey
          
         ElseIf Len(thiskey) = lenv Then
           hd.Add thiskey, thiskey
           
         End If
        
        Next Count
        
        If hd.Count > 0 Then ReDim VacNewIDs(hd.Count - 1)
        
        Count = 0
        
        For Each xkey In hd
          VacNewIDs(Count) = xkey
          Count = Count + 1
        Next xkey
        
        GetThisVacIDKeys = VacNewIDs
        

Func_Exit:
  On Error Resume Next
  Set hd = Nothing
  Exit Function

Func_Err:
 ErrMsgBox "Patient_Imm.GetThisVacIDKeys", Error$, Err
 Resume Func_Exit
 Resume 'Debug
End Function

Public Sub AddExemption(ByVal VacID As String, ByVal iVacc_Series As Integer, ByVal iExempt As Integer, ByVal vDate As Variant)
  
  ' Add an Exemption to Required Records
  
  Dim VacRec() As String
  Dim gn As String
  
  On Error GoTo Sub_Err
  
  If obd_PatientsVaccines.Exists(VacID) Then
  
     gn = IMM_Tables.Get_A_Generic_Name(VacID)
     
     VacRec = GetPatientsLastVaccineRecord(VacID)
     
     VacRec(vacrsfields.ExemptionExpirationDate) = IIf(IsNull(vDate), vbNullString, vDate)
     VacRec(vacrsfields.Vacc_Exempt) = CStr(iExempt)
     VacRec(vacrsfields.Vacc_Series) = CStr(iVacc_Series)
     VacRec(vacrsfields.Vacc_Date) = Format$(Date, "mm/dd/yy")
     VacRec(vacrsfields.VAcc_Current) = "True"
     obd_PatientsVaccines.Add CreateVacID(VacID, GetHighNewVacKey(VacID) + 1), Join(VacRec, FieldSepChar)
     
     BuildFinalGenericRecord gn, VacRec
     
  End If

Sub_exit:
  On Error Resume Next
  Exit Sub

Sub_Err:
 ErrMsgBox "Patient_Imm.GetThisVacIDKeys", Error$, Err
 Resume Sub_exit
 Resume 'Debug

End Sub

Public Function VaccinesEffected(ByVal Solves As Boolean, ByVal ThisVacID As String) As Boolean

   ' This routine does One of the Following:
   
   '                Solves  = True     VacID will solve ThisVacID
   '                          False    VacID is Solved By ThisVacID
 On Local Error GoTo haveerror
 Dim vacids() As String
 Dim VacID As Variant
 Dim ThisVacIDRec() As String
 Dim vacidrec() As String
 Dim ThisVacIDDate As Date
 Dim VacIDDate As Date
 Dim foundit As Boolean
 Dim change As Boolean
 Dim TheVaxID As String
 Dim xloc As Integer
 
   
   If obd_PatientsVaccines.Exists(ThisVacID) Then
     ThisVacIDRec = Split(obd_PatientsVaccines.Item(ThisVacID), FieldSepChar)
     ThisVacIDDate = CDate(ThisVacIDRec(vacrsfields.Vacc_Date))
   Else
    Exit Function
   End If
      
   '...Get Base VacID
   xloc = InStr(ThisVacID, WithinFieldSepChar)
   If xloc <> 0 Then
    TheVaxID = Left(ThisVacID, xloc - 1)
   Else
    TheVaxID = ThisVacID
   End If
   
   Select Case Solves
   
       Case False
       
         vacids = IMM_Tables.ThisVaccineIsSolvedBy(TheVaxID)
         If UBound(vacids) <> -1 Then
         
            For Each VacID In vacids
              change = False
              '...is this VacId to be considered
              If obd_PatientsVaccines.Exists(CStr(VacID)) Then
                 ' YES
                 ' Load VacID record
                 vacidrec = Split(obd_PatientsVaccines.Item(VacID), FieldSepChar)
                 '...Is it Current
                 If vacidrec(vacrsfields.VAcc_Current) = "True" Then
                   ' YES
                   ThisVacIDRec(vacrsfields.VAcc_Current) = "True"
                   'Indicate Current
                   ThisVacIDRec(vacrsfields.EndStatus) = "C"
                   ThisVacIDRec(vacrsfields.Vacc_NextDue) = vbNullString
                   foundit = True
                   change = True
                 End If
                 ' Is there a valid Date
                 If LenB(vacidrec(vacrsfields.Vacc_Date)) <> 0 Then
                    ' YES
                    VacIDDate = CDate(vacidrec(vacrsfields.Vacc_Date))
                    ' VacID date administered more current
                    If VacIDDate > ThisVacIDDate Then
                      ' YES
                      ' Reflect the more current date
                      ThisVacIDRec(vacrsfields.Vacc_Date) = vacidrec(vacrsfields.Vacc_Date)
                      ThisVacIDRec(vacrsfields.SolvedByVacID) = TheVaxID
                      change = True
                    End If
                 End If
                 ' Does ThisVacID record need to be updated
                 If change Then
                   ' YES
                   obd_PatientsVaccines.Item(ThisVacID) = Join(ThisVacIDRec, FieldSepChar)
                   ' Update generic record
                   UpdateVactoGenForPatient IMM_Tables.Get_A_Generic_Name(ThisVacID)
                 End If
              End If
              ' Even though we may have found or changed ThisVacID record we still need to
              ' look at all the VacID's in case one is more current then the one found
            Next VacID
            
            VaccinesEffected = foundit
         
         End If
       
       Case True
       
         vacids = IMM_Tables.ThisVaccineSolves(TheVaxID)
         If UBound(vacids) <> -1 Then
         
          '  Are we Current
          If ThisVacIDRec(vacrsfields.VAcc_Current) = "True" Then
            ' YES
            For Each VacID In vacids
              ' is VacId required
              If obd_PatientsVaccines.Exists(CStr(VacID)) Then
                 ' YES
                 vacidrec = Split(obd_PatientsVaccines.Item(VacID), FieldSepChar)
                    change = False
                    ' is VacID already current
                    If vacidrec(vacrsfields.VAcc_Current) = vbNullString _
                       Or _
                       vacidrec(vacrsfields.VAcc_Current) = "False" Then
                       ' NO
                       ' Make VacID Current
                       vacidrec(vacrsfields.VAcc_Current) = "True"
                       'Indicate Current
                       vacidrec(vacrsfields.EndStatus) = "C"
                       vacidrec(vacrsfields.Vacc_NextDue) = vbNullString
                       ' Does VacID have a valid date
                       If LenB(vacidrec(vacrsfields.Vacc_Date)) = 0 Then
                         ' No
                         ' Use ThisVacID's date then
                         vacidrec(vacrsfields.Vacc_Date) = ThisVacIDRec(vacrsfields.Vacc_Date)
                       Else
                        VacIDDate = CDate(vacidrec(vacrsfields.Vacc_Date))
                        ' Use ThisVacID's date if more current
                        If VacIDDate < ThisVacIDDate Then vacidrec(vacrsfields.Vacc_Date) = ThisVacIDRec(vacrsfields.Vacc_Date)
                       End If
                       vacidrec(vacrsfields.SolvedByVacID) = TheVaxID
                       change = True
                       foundit = True
                    Else
                       ' Does VacID have a valid Date
                       If LenB(vacidrec(vacrsfields.Vacc_Date)) = 0 Then
                         ' NO
                         ' Use ThisVacID's date
                         vacidrec(vacrsfields.Vacc_Date) = ThisVacIDRec(vacrsfields.Vacc_Date)
                         change = True
                       Else
                        ' Determine more current date between VacID and ThisVacID
                        VacIDDate = CDate(vacidrec(vacrsfields.Vacc_Date))
                        If VacIDDate < ThisVacIDDate Then
                          vacidrec(vacrsfields.Vacc_Date) = ThisVacIDRec(vacrsfields.Vacc_Date)
                          change = True
                        End If
                       End If
                    End If
                    ' Does VacID need to be updated
                    If change Then
                      ' Yes
                      obd_PatientsVaccines.Item(VacID) = Join(vacidrec, FieldSepChar)
                      ' Update Generic Record
                      UpdateVactoGenForPatient IMM_Tables.Get_A_Generic_Name(VacID)
                    End If
              End If
              
            Next VacID
            
            VaccinesEffected = foundit
            
          End If
         
         End If
       
   End Select
exitRoutine:
   Exit Function
haveerror:
   Debug.Print Err.Number, Err.Description
   Resume exitRoutine
   Resume
End Function

Private Sub CheckForCombinationVaccines()

  '...See if Vaccines effected other Vaccines

  Dim ComboVac() As String
  Dim Combo As Variant
  
  On Error GoTo Sub_Err
  
  ComboVac = IMM_Tables.GetCombinationVaccines
  
  If UBound(ComboVac) <> -1 Then
  
    For Each Combo In ComboVac
    
'     Debug.Assert CStr(Combo) <> "51"
      '...See If this effects other Vaccines
      VaccinesEffected True, CStr(Combo)
    
    Next Combo
  
  End If

Sub_exit:
  On Error Resume Next
  Exit Sub

Sub_Err:
 ErrMsgBox "Patient_Imm.CheckForCombinationVaccines", Error$, Err
 Resume Sub_exit
 Resume 'Debug
  
End Sub

Public Function GetPatientsGroups(ByVal OwnerShip As Integer) As Dictionary
  Dim Group As Variant
  Dim Groups As Dictionary
  Dim grouprecord() As String
  Dim GroupName As String
  Dim xloc As Integer
  
  On Error GoTo Func_Err
  
  Set Groups = New Dictionary
  
  For Each Group In obd_PatientsGroups
  
   grouprecord = Split(obd_PatientsGroups.Item(Group), FieldSepChar)
   
   If (grouprecord(grouprec.idcontrol) And OwnerShip) Then
   
     xloc = InStr(Group, WithinFieldSepChar)
     
     If xloc = 0 Then
       GroupName = CStr(Group)
     Else
       GroupName = Mid$(Group, xloc + 1)
     End If
     
     If Not Groups.Exists(GroupName) Then Groups.Add GroupName, grouprecord
     
   End If
   
  Next Group
  
  Set GetPatientsGroups = Groups
  
Func_Exit:
  On Error Resume Next
  Set Groups = Nothing
  Exit Function

Func_Err:
 ErrMsgBox "Patient_Imm.GetPatientsGroups", Error$, Err
 Resume Func_Exit
 Resume 'Debug
  
  
End Function

Public Sub AddGroupToPatient(ByVal GroupName As String, Optional UpdateVaccine As Boolean = False)

 '  Adds a Group to obd_PagientGroups if it does not already Exist
 '  The record added to the group is controlled by the Following priority
   
 '    1.  Unit Group Names
 '    2.  User Defined Group Names
 
 On Local Error GoTo haveerror
 
 Dim Record As String
 Dim DefPrefix As String
 Dim ProcessVaccine As Boolean
 
 DefPrefix = "DEFAULT"
 
      '...Build Groups to work on
      '...If group already exists it came from the Unit Required Groups
      '...therefore the entry is not over written
      ' bookmark
      If Not obd_PatientsGroups.Exists(GroupName) Then
        '...Look for Default record First
        Record = IMM_Tables.UserGroupRecord(DefPrefix & WithinFieldSepChar & GroupName)
        If LenB(Record) <> 0 Then
         obd_PatientsGroups.Add GroupName, Record
         ProcessVaccine = True
        Else
         '...Look for userDemis defined record second
         Record = IMM_Tables.UserGroupRecord(gsMTF_ID & WithinFieldSepChar & GroupName)
         If LenB(Record) <> 0 Then
          obd_PatientsGroups.Add GroupName, Record
          ProcessVaccine = True
         Else
            MsgBxARMd Disclaimer & _
                   "Patient Social: " & IIf(Len(SSN) = 0, "VOID", SSN) & vbCrLf & _
                   "        DmisID: " & IIf(Len(gsMTF_ID) = 0, "VOID", gsMTF_ID) & vbCrLf & _
                   "    Group Name: " & IIf(Len(GroupName) = 0, "VOID", GroupName) & vbCrLf & vbCrLf & _
                   "Does not have a Group Name entry in Table User_Vaccine_Groups", vbInformation + vbOKOnly, "Invalid Record"
         End If
        End If
        
        '...Do we need to process vaccines
        If ProcessVaccine And UpdateVaccine Then
           ' YES
           ' Add required vaccines etc. Here
            Dim AddedInformation As String
            Dim AddedGenNames As String
               AddedGenNames = LoadPatientGenericNames(, GroupName)
               If LenB(AddedGenNames) <> 0 Then
                  AddedInformation = LoadPatientsVaccines(Now, AddedGenNames)
                  If LenB(AddedInformation) <> 0 Then
                     CalculateNextDueDate Now, AddedInformation
                     CheckForCombinationVaccines
                     PatientGenericIMMSchedule AddedGenNames
                  End If
               End If
           
        End If
        
      End If
exitRoutine:
   Exit Sub
haveerror:
   Debug.Print Err.Number, Err.Description
   Resume exitRoutine
End Sub

Public Sub RemoveGroupFromPatient(ByVal GroupName As String)
  Dim GenericNames() As String
  Dim OtherGroupGenNames() As String
  Dim GenericNameCount() As Boolean
  Dim Record() As String
  Dim Group As Variant
  Dim OldGenName As Variant
  Dim OtherGenName As Variant
  Dim Count As Integer
  Dim Historical As Boolean
  Dim HistVac As Variant
  Dim GenNameToDelete As String
  
  On Error GoTo Sub_Err
  
  '...Retrieve group record being removed
  Record = Split(obd_PatientsGroups.Item(GroupName), FieldSepChar)
  '...Load Generic Vaccine Names associated with that group
  GenericNames = Split(Record(grouprec.GenericNames), InvertedTblSepChar)
  '...Establish counter array
  '...If genericnamecount is TRUE then generic Vaccine Name is also in another group
  '...and consequently that generic vaccine name cannot be removed
  ReDim GenericNameCount(UBound(GenericNames))
  For Each Group In obd_PatientsGroups
    If Group <> GroupName Then
      Record = Split(obd_PatientsGroups.Item(Group), FieldSepChar)
      OtherGroupGenNames = Split(Record(grouprec.GenericNames), InvertedTblSepChar)
      Count = 0
      For Each OldGenName In GenericNames
        For Each OtherGenName In OtherGroupGenNames
         If OldGenName = OtherGenName Then
           GenericNameCount(Count) = True
           Exit For
         End If
        Next OtherGenName
        Count = Count + 1
      Next OldGenName
    End If
  Next Group
  '...See If certain Generic Vaccine Names should not be removed
  For Count = 0 To UBound(GenericNameCount)
    If Not GenericNameCount(Count) Then
    
      GenNameToDelete = GenericNames(Count)
      
      ' See if we need to keep because Patient has an Historical Vaccine
      Historical = False
      
      For Each HistVac In obd_PatientsHistory
      
      ' T.Noren  10/29/01 remove 'Force to upper case
'       If UCase$(obd_PatientsHistory.Item(HistVac)) = GenNameToDelete Then
       If obd_PatientsHistory.Item(HistVac) = GenNameToDelete Then
        Historical = True
        Exit For
       End If
       
      Next HistVac
      ' Remove Generic Name
      If Not Historical Then
        For Each HistVac In obd_PatientsVaccines
          '<< Begin:SCR #39335;   Developer: Brian Mowbray 07/30/2003 01:49 PM
          If obd_PatientsVaccines.Item(HistVac) <> Empty Then
            Record = Split(obd_PatientsVaccines.Item(HistVac), FieldSepChar)
            ' T.Noren  10/29/01 remove 'Force to upper case
  '          If UCase$(Record(vacrsfields.GenericName)) = GenNameToDelete Then obd_PatientsVaccines.Remove HistVac
            If Record(vacrsfields.GenericName) = GenNameToDelete Then obd_PatientsVaccines.Remove HistVac
          End If
        '>> End: SCR #39335;
        Next HistVac
        
        ' SCR 18378   T.Noren   10/05/01
        ' obd_PatientsGenNames.Remove GenNameToDelete
        If obd_PatientsGenNames.Exists(GenNameToDelete) Then obd_PatientsGenNames.Remove GenNameToDelete
        
      End If
    End If
  Next Count
  obd_PatientsGroups.Remove GroupName

Sub_exit:
  On Error Resume Next
  Exit Sub

Sub_Err:
 ErrMsgBox "Patient_Imm.RemoveGroupFromPatient", Error$, Err
 Resume Sub_exit
 Resume 'Debug
End Sub
