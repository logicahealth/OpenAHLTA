VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EncPregnancy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const MODULE_NAME As String = "EncPregnancy"

'Declare Object State Variables
Private mbDirty As Boolean
Private meCurrentState As enumRecordState

'Declare Data Variables
Private mbPregnancyIndicator As Boolean
Private mdLastPeriod As Date
Private mdEstDOB As Date

'Declare Data Variables - Details Mode
Private mbMenopauseIndicator As Boolean
Private mbHysterectomyIndicator As Boolean
Private mlTotalAbortus As Long
Private mlSpontaneousAbortions As Long
Private mlElectiveAbortions As Long
Private mlEctopicAbortions As Long
Private mlGravida As Long
Private mlPara As Long
Private mlLivingChildren As Long
Private moBirthControlMethods As Dictionary
Private moSavedBCMethods As Dictionary              'Used to track BC methods that are in the database

'Declare Ancillary Variables - These are not persisted
Private mlPatientAge As Long        'Age in years
Private msPatientSex As String      'Sex Code 'F' for female
Private moPatient As MMMHISPatient.Patient

Private Sub Class_Initialize()
   meCurrentState = rsNew
   Clear
End Sub

Private Sub Class_Terminate()
   Set moPatient = Nothing
End Sub

''''''''''''''''''''''''''''''''''''''''''''''
' Property Values
''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let PregnancyIndicator(ByVal bNewValue As Boolean)
   If mbPregnancyIndicator <> bNewValue Then
      mbPregnancyIndicator = bNewValue
      If Not mbPregnancyIndicator Then Clear
      mbDirty = True
   End If
End Property

Public Property Get PregnancyIndicator() As Boolean
   PregnancyIndicator = mbPregnancyIndicator
End Property

Public Property Let LastMenstrualPeriod(ByVal dNewValue As Date)
   If mdLastPeriod <> dNewValue Then
      mdLastPeriod = dNewValue
      mbDirty = True
   End If
End Property

Public Property Get LastMenstrualPeriod() As Date
   LastMenstrualPeriod = mdLastPeriod
End Property

Public Property Let EstimatedDOB(ByVal dNewValue As Date)
   If mdEstDOB <> dNewValue Then
      mdEstDOB = dNewValue
      mbDirty = True
   End If
End Property

Public Property Get EstimatedDOB() As Date
   EstimatedDOB = mdEstDOB
End Property

Public Property Let PatientAge(ByVal bNewValue As Long)
   mlPatientAge = bNewValue
End Property

Public Property Get PatientAge() As Long
   PatientAge = mlPatientAge
End Property

Public Property Let PatientSex(ByVal sNewValue As String)
   msPatientSex = sNewValue
End Property

Public Property Get PatientSex() As String
   PatientSex = msPatientSex
End Property

Public Property Let MenopauseIndicator(ByVal bNewValue As Boolean)
   If mbMenopauseIndicator <> bNewValue Then
      mbMenopauseIndicator = bNewValue
      mbDirty = True
   End If
End Property

Public Property Get MenopauseIndicator() As Boolean
   MenopauseIndicator = mbMenopauseIndicator
End Property

Public Property Let HysterectomyIndicator(ByVal bNewValue As Boolean)
   If mbHysterectomyIndicator <> bNewValue Then
      mbHysterectomyIndicator = bNewValue
      mbDirty = True
   End If
End Property

Public Property Get HysterectomyIndicator() As Boolean
   HysterectomyIndicator = mbHysterectomyIndicator
End Property

Public Property Get Completed() As Boolean
   'Completed if pregnant and LMP and EstDOB are filled in
   Completed = (PregnancyIndicator And mdLastPeriod <> NULL_DATE And mdEstDOB <> NULL_DATE)
End Property

Public Property Let TotalAbortus(ByVal lNewValue As Long)
   If mlTotalAbortus <> lNewValue Then
      mlTotalAbortus = lNewValue
      mbDirty = True
   End If
End Property

Public Property Get TotalAbortus() As Long
   TotalAbortus = mlTotalAbortus
End Property

Public Property Let LivingChildren(ByVal lNewValue As Long)
   If mlLivingChildren <> lNewValue Then
      mlLivingChildren = lNewValue
      mbDirty = True
   End If
End Property

Public Property Get LivingChildren() As Long
   LivingChildren = mlLivingChildren
End Property

Public Property Let Para(ByVal lNewValue As Long)
   If mlPara <> lNewValue Then
      mlPara = lNewValue
      mbDirty = True
   End If
End Property

Public Property Get Para() As Long
   Para = mlPara
End Property

Public Property Let Gravida(ByVal lNewValue As Long)
   If mlGravida <> lNewValue Then
      mlGravida = lNewValue
      mbDirty = True
   End If
End Property

Public Property Get Gravida() As Long
   Gravida = mlGravida
End Property

Public Property Let SpontaneousAbortions(ByVal lNewValue As Long)
   If mlSpontaneousAbortions <> lNewValue Then
      mlSpontaneousAbortions = lNewValue
      mbDirty = True
   End If
End Property

Public Property Get SpontaneousAbortions() As Long
   SpontaneousAbortions = mlSpontaneousAbortions
End Property

Public Property Let ElectiveAbortions(ByVal lNewValue As Long)
   If mlElectiveAbortions <> lNewValue Then
      mlElectiveAbortions = lNewValue
      mbDirty = True
   End If
End Property

Public Property Get ElectiveAbortions() As Long
   ElectiveAbortions = mlElectiveAbortions
End Property

Public Property Let EctopicAbortions(ByVal lNewValue As Long)
   If mlEctopicAbortions <> lNewValue Then
      mlEctopicAbortions = lNewValue
      mbDirty = True
   End If
End Property

Public Property Get EctopicAbortions() As Long
   EctopicAbortions = mlEctopicAbortions
End Property

''''''''''''''''''''''''''''''''''''''''''''''
' Data Methods
''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GetSQLFields() As String
   GetSQLFields = " PREGNANCY_INDICATOR, LAST_PERIOD_DATE, ESTIMATED_BIRTH_DATE, "
End Function

'Initialize data object from SQL object
Friend Sub InitFromSQL(ByRef oSQLEx As ICHCSII_SQL.ISqlOps, ByRef oPatient As MMMHISPatient.Patient, _
      Optional ByVal bExtendedFields As Boolean = False)
      
   PregnancyIndicator = (oSQLEx("PREGNANCY_INDICATOR") & "" = 1)
   LastMenstrualPeriod = SQLToDate(oSQLEx("LAST_PERIOD_DATE"))
   EstimatedDOB = SQLToDate(oSQLEx("ESTIMATED_BIRTH_DATE"))
   PatientSex = oPatient.Sex
   PatientAge = ConvertAge(oPatient.Age)
   
   'Load extended fields set for CDR mode
   If bExtendedFields Then
      MenopauseIndicator = (oSQLEx("POST_MENOPAUSE_INDICATOR") & "" = 1)
      HysterectomyIndicator = (oSQLEx("POST_HYSTERECTOMY_INDICATOR") & "" = 1)
      Gravida = IIf(oSQLEx("GRAVIDA_VALUE") = vbNullString, -1, oSQLEx("GRAVIDA_VALUE"))
      Para = IIf(oSQLEx("PARA_VALUE") = vbNullString, -1, oSQLEx("PARA_VALUE"))
      TotalAbortus = IIf(oSQLEx("ABORTUS_VALUE") = vbNullString, -1, oSQLEx("ABORTUS_VALUE"))
      SpontaneousAbortions = IIf(oSQLEx("SPONTANEOUS_ABORTIONS_VALUE") = vbNullString, -1, oSQLEx("SPONTANEOUS_ABORTIONS_VALUE"))
      ElectiveAbortions = IIf(oSQLEx("ELECTIVE_ABORTIONS_VALUE") = vbNullString, -1, oSQLEx("ELECTIVE_ABORTIONS_VALUE"))
      EctopicAbortions = IIf(oSQLEx("ECTOPIC_ABORTIONS_VALUE") = vbNullString, -1, oSQLEx("ECTOPIC_ABORTIONS_VALUE"))
      LivingChildren = IIf(oSQLEx("LIVING_CHILDREN_VALUE") = vbNullString, -1, oSQLEx("LIVING_CHILDREN_VALUE"))
   End If
   
   mbDirty = False
   Set moPatient = oPatient
   
End Sub

Friend Sub InitFromPatient(ByRef oPatient As MMMHISPatient.Patient, Optional ByVal bExtended As Boolean = False)

On Error GoTo ErrorHandler

Dim lIdx As Long
Dim sProperty As String
Dim sValue As String
   
   If bExtended Then
   
      With oPatient.PregnancyProperties
      
         For lIdx = 0 To .Count - 1
            sProperty = .Keys(lIdx)
            sValue = .Items(lIdx)
            
            If Left(sProperty, 8) = "BCMethod" Then
               BirthControlMethods.Add sValue, sValue
            Else
               CallByName Me, sProperty, VbLet, sValue
            End If
            
         Next
         
      End With
        
   Else
      'Load from simple properties
      PregnancyIndicator = oPatient.PregnancyIndicator
      LastMenstrualPeriod = oPatient.LastMenstrualPeriod
      EstimatedDOB = oPatient.EstimatedDOB
   End If
   
   PatientSex = oPatient.Sex
   PatientAge = ConvertAge(oPatient.Age)
   mbDirty = False
   meCurrentState = rsNew
   Set moPatient = oPatient
   
ErrorHandler:
   If Err.Number = 438 Then
      Err.Clear
      Resume Next
   ElseIf CBool(Err) Then
      HandleErrors Err, MODULE_NAME, "CDRPlusLoad"
   End If
   
End Sub

'Convert the age as it is in the patient object to a number
'Just set to 2 if less than 2 years old since internal business rules won't be affected
Private Function ConvertAge(ByVal sAge As String) As Long
   
   If InStr(1, sAge, "yo", vbTextCompare) > 0 Then
      sAge = Replace(sAge, "yo", "", 1, -1, vbTextCompare)
      If IsNumeric(sAge) Then
         ConvertAge = Val(sAge)
      Else
         ConvertAge = 2
      End If
   Else
      ConvertAge = 2
   End If
   
End Function

Friend Function GetSQLUpdate() As String
   If PregnancyIndicator Then
      GetSQLUpdate = ", PREGNANCY_INDICATOR = 1" & _
                     ", LAST_PERIOD_DATE = " & IIf(mdLastPeriod = NULL_DATE, "NULL", gobjCHCSIIConn.SQLDate(mdLastPeriod)) & _
                     ", ESTIMATED_BIRTH_DATE = " & IIf(mdEstDOB = NULL_DATE, "NULL", gobjCHCSIIConn.SQLDate(mdEstDOB))
   Else
      GetSQLUpdate = ", PREGNANCY_INDICATOR = 0" & _
                     ", LAST_PERIOD_DATE = NULL" & _
                     ", ESTIMATED_BIRTH_DATE = NULL"
   End If
End Function

Friend Function GetSQLValues() As String
   If PregnancyIndicator Then
      GetSQLValues = " 1, " & _
                     IIf(mdLastPeriod = NULL_DATE, "NULL", gobjCHCSIIConn.SQLDate(mdLastPeriod)) & ", " & _
                     IIf(mdEstDOB = NULL_DATE, "NULL", gobjCHCSIIConn.SQLDate(mdEstDOB)) & ", "
   Else
      GetSQLValues = " 0, " & IIf(mdLastPeriod = NULL_DATE, "NULL", gobjCHCSIIConn.SQLDate(mdLastPeriod)) & ", NULL, "
   End If
End Function

'SCR 46951 - Perform validation
Public Function Validate(ByVal bSigning As Boolean, ByRef sWarning As String) As Boolean

On Error GoTo ErrorHandler
      
   If mbPregnancyIndicator And mdLastPeriod <> NULL_DATE And mdLastPeriod > Date Then
      sWarning = "Last Menstrual Period should not be in the future."
   ElseIf mbPregnancyIndicator And mdEstDOB <> NULL_DATE And mdEstDOB < Date Then
      sWarning = "Estimated Date of Birth should not be in the past."
   'ElseIf mbPregnancyIndicator And (DateDiff("d", mdLastPeriod) > 320 Or DateDiff("d", mdLastPeriod) < 27) Then
   '   sWarning = "The LMP date appears incorrect if the patient is pregnant." & vbCrLf & "Do you wish to continue?"
   ElseIf bSigning And mbPregnancyIndicator And (mdEstDOB = NULL_DATE Or mdLastPeriod = NULL_DATE) Then
      sWarning = "Last Menstrual Period and Estimated Date of Birth are required when the patient is pregnant."
   Else
      Validate = True
      sWarning = vbNullString
   End If

ErrorHandler:
   If CBool(Err) Then HandleErrors Err, MODULE_NAME, "Validate"
   
End Function

Friend Sub CopyToWam(ByRef objWamData As WAMData.EncPregnancy)
   objWamData.PregnancyIndicator = Me.PregnancyIndicator
   If Me.PregnancyIndicator Then
      objWamData.LastMenstrualPeriod = Me.LastMenstrualPeriod
      objWamData.EstimatedDOB = Me.EstimatedDOB
   Else
      objWamData.LastMenstrualPeriod = NULL_DATE
      objWamData.EstimatedDOB = NULL_DATE
   End If
End Sub

Public Property Get IsDirty() As Boolean
   IsDirty = mbDirty
End Property

Private Function CDRPlusSaveContraceptions(ByVal sFacilityNCID As String, ByVal sEncNumber As String) As Boolean

On Error GoTo ErrorHandler

Dim sDelNCIDs As String
Dim sSql As String
Dim oSQL As ICHCSII_SQL.ISqlOps
Dim lIdx As Long

   'Initialize the connection object
   Set oSQL = gobjCHCSIIConn.CHCSII_SQLOPS(Auto)
   
   'Build list of BC methods that need to be removed from the DB
   For lIdx = 0 To SavedBCMethods.Count - 1
      If Not moBirthControlMethods.Exists(moSavedBCMethods.Keys(lIdx)) Then
         sDelNCIDs = sDelNCIDs & moSavedBCMethods.Keys(lIdx) & ", "
      End If
   Next lIdx
   
   'If items were found remove then from the database
   If Len(sDelNCIDs) > 0 Then
      sDelNCIDs = Mid$(sDelNCIDs, 1, Len(sDelNCIDs) - 2)
      sSql = "DELETE FROM encounter_contraception " & _
                  " WHERE Facility_NCID = " & sFacilityNCID & _
                     " AND Encounter_Number = " & sEncNumber & _
                     " AND Contraception_Type_NCID in ( " & sDelNCIDs & ")"
      
      oSQL.Execute sSql
   End If
   
   'Insert newly added BC methods into the database
   For lIdx = 0 To BirthControlMethods.Count - 1
      If Not moSavedBCMethods.Exists(moBirthControlMethods.Keys(lIdx)) Then
         'This item doesn't already exist so insert it into the DB
         sSql = "INSERT INTO encounter_contraception (" & _
                  " Facility_NCID, Encounter_Number, Contraception_Type_NCID " & _
                  " ) VALUES ( " & sFacilityNCID & ", " & sEncNumber & ", " & _
                  moBirthControlMethods.Keys(lIdx) & " )"
      
         oSQL.Execute sSql
      End If
   
   Next lIdx
   
   'Get the saved list back in sync with actual items in DB
   SyncBCMethods
   
   CDRPlusSaveContraceptions = True
   
ErrorHandler:
   Set oSQL = Nothing

   If CBool(Err) Then HandleErrors Err, MODULE_NAME, "CDRPlusSave"
      
End Function

Public Function CDRPlusSave(ByVal sFacilityNCID As String, ByVal sEncNumber As String, ByVal eEncStatus As EncStatus) As Boolean

On Error GoTo ErrorHandler

Dim sSql As String
Dim oSQL As ICHCSII_SQL.ISqlOps
Dim oShared As CWShared
Dim bSuccess As Boolean
Dim sNullValue As String

   'SCR 52248 - Don't save for non female patients
   If msPatientSex <> "F" Then
      CDRPlusSave = True
      Exit Function
   End If
   
   'Determine what value to use for nulls - ITT DAS object handles nulls differently than the CDR Mode DAS
   Set oShared = New CWShared
   If oShared.IsAppMode(modeITT) Then
      sNullValue = "-1"
   Else
      sNullValue = "NULL"
   End If
   
   'Validate SQL Parameters
   ValidateSqlParams sFacilityNCID, sEncNumber
    
   If meCurrentState = rsNew Then
      
      sSql = "INSERT INTO Encounter_Pregnancy ( " & _
               " Facility_NCID, Encounter_Number, Pregnancy_Indicator, " & _
               " Last_Period_Date, Estimated_Birth_Date, Post_Menopause_Indicator, " & _
               " Post_Hysterectomy_Indicator, Abortus_Value, Spontaneous_Abortions_Value, " & _
               " Elective_Abortions_Value, Ectopic_Abortions_Value, Gravida_Value, " & _
               " Para_Value, Living_Children_Value " & _
             ") VALUES ( " & sFacilityNCID & ", " & sEncNumber & ", " & GetSQLValues & _
                  IIf(mbMenopauseIndicator, 1, 0) & ", " & _
                  IIf(mbHysterectomyIndicator, 1, 0) & ", " & _
                  IIf(mlTotalAbortus >= 0, mlTotalAbortus, sNullValue) & ", " & _
                  IIf(mlSpontaneousAbortions >= 0, mlSpontaneousAbortions, sNullValue) & ", " & _
                  IIf(mlElectiveAbortions >= 0, mlElectiveAbortions, sNullValue) & ", " & _
                  IIf(mlEctopicAbortions >= 0, mlEctopicAbortions, sNullValue) & ", " & _
                  IIf(mlGravida >= 0, mlGravida, sNullValue) & ", " & _
                  IIf(mlPara >= 0, mlPara, sNullValue) & ", " & _
                  IIf(mlLivingChildren >= 0, mlLivingChildren, sNullValue) & " )"
                  
   ElseIf (meCurrentState = rsSaved And IsDirty) Or meCurrentState = rsModified Then
      
      sSql = "UPDATE Encounter_Pregnancy " & _
               " SET Pregnancy_Indicator = " & IIf(mbPregnancyIndicator, 1, 0) & ", " & _
               " Last_Period_Date = " & IIf(mdLastPeriod = NULL_DATE, "NULL", gobjCHCSIIConn.SQLDate(mdLastPeriod)) & ", " & _
               " Estimated_Birth_Date = " & IIf(mdEstDOB = NULL_DATE, "NULL", gobjCHCSIIConn.SQLDate(mdEstDOB)) & ", " & _
               " Post_Menopause_Indicator = " & IIf(mbMenopauseIndicator, 1, 0) & ", " & _
               " Post_Hysterectomy_Indicator = " & IIf(mbHysterectomyIndicator, 1, 0) & ", " & _
               " Abortus_Value = " & IIf(mlTotalAbortus >= 0, mlTotalAbortus, sNullValue) & ", " & _
               " Spontaneous_Abortions_Value = " & IIf(mlSpontaneousAbortions >= 0, mlSpontaneousAbortions, sNullValue) & ", " & _
               " Elective_Abortions_Value = " & IIf(mlElectiveAbortions >= 0, mlElectiveAbortions, sNullValue) & ", " & _
               " Ectopic_Abortions_Value = " & IIf(mlEctopicAbortions >= 0, mlEctopicAbortions, sNullValue) & ", " & _
               " Gravida_Value = " & IIf(mlGravida >= 0, mlGravida, sNullValue) & ", " & _
               " Para_Value = " & IIf(mlPara >= 0, mlPara, sNullValue) & ", " & _
               " Living_Children_Value = " & IIf(mlLivingChildren >= 0, mlLivingChildren, sNullValue) & _
             " WHERE Facility_NCID = " & sFacilityNCID & _
                  " AND Encounter_Number = " & sEncNumber

   End If
   
   If Len(sSql) > 0 Then
      Set oSQL = gobjCHCSIIConn.CHCSII_SQLOPS(Auto)
      oSQL.Execute sSql
            
      If meCurrentState = rsDelete Then
         meCurrentState = rsRemove
      Else
         meCurrentState = rsSaved
      End If
                 
   End If
   
   'Now save contraceptions and update the patient data
   bSuccess = True
   If mbDirty Then bSuccess = bSuccess And CDRPlusSaveContraceptions(sFacilityNCID, sEncNumber)
   If bSuccess And mbDirty Then bSuccess = UpdatePatient(sFacilityNCID, sEncNumber, eEncStatus)
   
   mbDirty = False
   
   CDRPlusSave = bSuccess

ErrorHandler:
   Set oSQL = Nothing
   Set oShared = Nothing
   
   If CBool(Err) Then HandleErrors Err, MODULE_NAME, "CDRPlusSave"
      
End Function

'Update the pregnancy data in the database
Public Function GemsSave(ByVal sFacilityNCID As String, ByVal sEncNumber As String, ByVal sEncStatus As EncStatus) As Boolean

On Error GoTo ErrorHandler

Dim sSql As String
Dim oDas As ICHCSII_DAS
Dim bSuccess As Boolean

   'Validate SQL Parameters
   ValidateSqlParams sFacilityNCID, sEncNumber
   
   'Build Update SQL
   sSql = "UPDATE Encounters " & _
            " SET " & Mid(GetSQLUpdate, 2) & _
            " WHERE FacilityNCID = " & sFacilityNCID & _
            "  AND EncounterNumber = " & sEncNumber
   
   Set oDas = gobjCHCSIIConn.CHCSII_DAS(Auto)
   oDas.ExecuteSQL sSql
   
   If sEncStatus = Updating Or sEncStatus = Updated Then
      If Not GemsCheckForNewerEncounter(sFacilityNCID, sEncNumber) Then
         'SCR 46950 - Only update patient for ammended encounters when there are no encounters that have been completed since this one
         bSuccess = moPatient.UpdatePregnancyStatus(PregnancyIndicator, LastMenstrualPeriod, EstimatedDOB)
      Else
         'Don't update patient Information
         bSuccess = True
      End If
   Else
      'Allways update the patient level data for all other statuses
      bSuccess = moPatient.UpdatePregnancyStatus(PregnancyIndicator, LastMenstrualPeriod, EstimatedDOB)
   End If
   
   If Not PregnancyIndicator Then Clear
   
   mbDirty = False
   
   GemsSave = bSuccess
   
ErrorHandler:
   Set oDas = Nothing

   If CBool(Err) Then HandleErrors Err, MODULE_NAME, "GemsSave"
   
End Function

'Determine if any encounters have been signed (completed) since this encounter was signed
Private Function GemsCheckForNewerEncounter(ByVal sFacilityNCID As String, ByVal sEncounterID As String) As Boolean

On Error GoTo ErrorHandler

Dim sSql As String
Dim oDas As ICHCSII_DAS
Dim oRS As ADODB.Recordset

   'Initialize Return Value
   GemsCheckForNewerEncounter = False
   
   'Build the SQL
   sSql = "SELECT Count(*) " & _
          " FROM (   SELECT Unit_Number, WhenSigned " & _
                     " FROM Encounters " & _
                     " WHERE FacilityNCID = " & sFacilityNCID & " AND " & _
                        " EncounterNumber = " & sEncounterID & " ) CurEnc, " & _
           " Encounters PatEnc " & _
           " WHERE PatEnc.Unit_number = CurEnc.Unit_Number AND " & _
               " PatEnc.WhenSigned > CurEnc.WhenSigned"
   
   'Initialize the connection object
   Set oDas = gobjCHCSIIConn.CHCSII_DAS(Auto)
   
   Set oRS = oDas.OpenRecordset(sSql, adOpenForwardOnly, adLockReadOnly)
   
   If Not oRS Is Nothing Then
      If Not oRS.BOF And Not oRS.EOF Then
         GemsCheckForNewerEncounter = CBool(oRS(0) > 0)
      End If
   End If
   
ErrorHandler:
   Set oDas = Nothing

   If CBool(Err) Then HandleErrors Err, MODULE_NAME, "GemsCheckForNewerEncounter"
   
End Function

Public Sub Clear()
   'Set Default Values
   mbPregnancyIndicator = False
   mdLastPeriod = NULL_DATE
   mdEstDOB = NULL_DATE
   mbMenopauseIndicator = False
   mbHysterectomyIndicator = False
   mlGravida = -1
   mlPara = -1
   mlTotalAbortus = -1
   mlLivingChildren = -1
   mlSpontaneousAbortions = -1
   mlElectiveAbortions = -1
   mlEctopicAbortions = -1
   mbDirty = False
End Sub

Public Sub AddBCMethod(ByVal sTypeNCID As String, sTypeName As String)
   If Not BirthControlMethods.Exists(sTypeNCID) Then
      moBirthControlMethods.Add sTypeNCID, sTypeName
      mbDirty = True
   End If
End Sub

Public Sub RemoveBCMethod(ByVal sTypeNCID As String)
   If BirthControlMethods.Exists(sTypeNCID) Then
      moBirthControlMethods.Remove sTypeNCID
      mbDirty = True
   End If
End Sub

Public Property Get BirthControlMethods() As Dictionary
   If moBirthControlMethods Is Nothing Then
      Set moBirthControlMethods = New Dictionary
   End If
   Set BirthControlMethods = moBirthControlMethods
End Property

Private Property Get SavedBCMethods() As Dictionary
   If moSavedBCMethods Is Nothing Then
      Set moSavedBCMethods = New Dictionary
   End If
   Set SavedBCMethods = moSavedBCMethods
End Property

'Load Pregnancy data from the CDRPlus Schema
Public Function CDRPlusLoad(ByVal sFacilityNCID As String, ByVal sEncNumber As String, ByRef oPatient As MMMHISPatient.Patient) As Boolean
   
On Error GoTo ErrorHandler

Dim sSql As String
Dim oSQL As ICHCSII_SQL.ISqlOps
Dim bSuccess As Boolean
Dim lRecordCount As Long

   'Validate SQL Parameters
   ValidateSqlParams sFacilityNCID, sEncNumber
   
   'Initialize the connection
   Set oSQL = gobjCHCSIIConn.CHCSII_SQLOPS(Auto)
   
   'SCR 52248 - Retreive pregnancy data for females only
   If oPatient.Sex = "F" Then
      sSql = "SELECT * FROM Encounter_Pregnancy " & _
             " WHERE Facility_NCID = " & sFacilityNCID & _
               " AND Encounter_Number = " & sEncNumber
      oSQL.Execute sSql
      lRecordCount = oSQL.RecordCount
   End If
   
   If lRecordCount = 0 Then
      'No data found so initialize internal parameters only
      Set moPatient = oPatient
      mlPatientAge = ConvertAge(moPatient.Age)
      msPatientSex = moPatient.Sex
      meCurrentState = rsNew
   Else
   
      'Set Parameter Values
      oSQL.MoveFirst
      InitFromSQL oSQL, oPatient, True
   
      'Retreive birth control data
      sSql = "SELECT * FROM Encounter_Contraception" & _
             " WHERE Facility_NCID = " & sFacilityNCID & _
               " AND Encounter_Number = " & sEncNumber
      oSQL.Execute sSql
      
      Set moBirthControlMethods = New Dictionary
      Do While Not oSQL.EOF
         moBirthControlMethods.Add CStr(oSQL("CONTRACEPTION_TYPE_NCID")), CStr(oSQL("CONTRACEPTION_TYPE_NCID"))
         oSQL.MoveNext
      Loop
      SyncBCMethods
      
      meCurrentState = rsSaved
      
   End If
   
   mbDirty = False

   CDRPlusLoad = True
   
ErrorHandler:
   Set oSQL = Nothing

   If CBool(Err) Then HandleErrors Err, MODULE_NAME, "CDRPlusLoad"
            
End Function

Private Sub ValidateSqlParams(ByVal sFacilityNCID As String, ByVal sEncNumber As String)
   If sEncNumber = vbNullString Then
      Err.Raise 911, , "Missing Encounter Id"
   ElseIf sFacilityNCID = vbNullString Then
      Err.Raise 911, , "Missing Facility NCID"
   End If
End Sub

'Update the saved BC methods list with the latest current copy
Private Sub SyncBCMethods()
Dim lIdx As Long
   Set moSavedBCMethods = Nothing
   Set moSavedBCMethods = New Dictionary
   For lIdx = 0 To moBirthControlMethods.Count - 1
      moSavedBCMethods.Add moBirthControlMethods.Keys(lIdx), moBirthControlMethods.Items(lIdx)
   Next lIdx
End Sub

Private Function UpdatePatient(ByVal sFacilityNCID As String, ByVal sEncNumber As String, ByVal sEncStatus As EncStatus) As Boolean

On Error GoTo ErrorHandler

Dim sPatientData As String
Dim lIdx As Long

   'Build Patient Persistance String
   sPatientData = "PregnancyIndicator=" & IIf(mbPregnancyIndicator, 1, 0) & ", " & _
      "LastMenstrualPeriod=" & mdLastPeriod & ", " & _
      "EstimatedDOB=" & mdEstDOB & ", " & _
      "MenopauseIndicator=" & IIf(mbMenopauseIndicator, 1, 0) & ", " & _
      "HysterectomyIndicator=" & IIf(mbHysterectomyIndicator, 1, 0) & _
      IIf(mlTotalAbortus >= 0, ", TotalAbortus=" & mlTotalAbortus, vbNullString) & _
      IIf(mlSpontaneousAbortions >= 0, ", SpontaneousAbortions=" & mlSpontaneousAbortions, vbNullString) & _
      IIf(mlElectiveAbortions >= 0, ", ElectiveAbortions=" & mlElectiveAbortions, vbNullString) & _
      IIf(mlEctopicAbortions >= 0, ", EctopicAbortions=" & mlEctopicAbortions, vbNullString) & _
      IIf(mlGravida >= 0, ", Gravida=" & mlGravida, vbNullString) & _
      IIf(mlPara >= 0, ", Para=" & mlPara, vbNullString) & _
      IIf(mlLivingChildren >= 0, ", LivingChildren=" & mlLivingChildren, vbNullString)

   For lIdx = 0 To BirthControlMethods.Count - 1
      sPatientData = sPatientData & ", BCMethod" & lIdx + 1 & "=" & BirthControlMethods.Keys(lIdx)
   Next
   
   If sEncStatus = Updating Or sEncStatus = Updated Then
      If Not CDRPlusCheckForNewerEncounter(sFacilityNCID, sEncNumber, moPatient.UnitNumber) Then
         'Only update patient for ammended encounters when there are no encounters that have been completed since this one
         UpdatePatient = moPatient.UpdatePregnancyStatus(PregnancyIndicator, LastMenstrualPeriod, EstimatedDOB, sPatientData)
      Else
         'Don't update patient Information
         UpdatePatient = True
      End If
   Else
      'Allways update the patient level data for all other statuses
      UpdatePatient = moPatient.UpdatePregnancyStatus(PregnancyIndicator, LastMenstrualPeriod, EstimatedDOB, sPatientData)
   End If
      
ErrorHandler:

   If CBool(Err) Then HandleErrors Err, MODULE_NAME, "UpdatePatient"
            
End Function


'Determine if any encounters have been signed (completed) since this encounter was signed
Private Function CDRPlusCheckForNewerEncounter(ByVal sFacilityNCID As String, ByVal sEncounterID As String, _
      ByVal sPatientUnitNumber As String) As Boolean

On Error GoTo ErrorHandler

Dim sSql As String
Dim oDas As ICHCSII_DAS
Dim oRS As ADODB.Recordset

   'Initialize Return Value
   CDRPlusCheckForNewerEncounter = False
   
   'Build the SQL
   sSql = "SELECT Count(*) " & _
          " FROM (   SELECT Patient_Unit_Number, Date_Modified, Date_Created " & _
                     " FROM Appointment " & _
                     " WHERE Facility_NCID = " & sFacilityNCID & _
                        " AND Encounter_Number = " & SQLQuote(sEncounterID) & _
                        " AND Patient_Unit_Number = " & sPatientUnitNumber & " ) CurAppt, " & _
           " Appointment PatAppt " & _
           " WHERE PatAppt.Patient_Unit_Number = CurAppt.Patient_Unit_Number " & _
               " AND PatAppt.Encounter_Status = 4 " & _
               " AND ( PatAppt.Date_Created > CurAppt.Date_Modified " & _
               " or PatAppt.Date_Modified > CurAppt.Date_Modified )"
   
   'Initialize the connection object
   Set oDas = gobjCHCSIIConn.CHCSII_DAS(Auto)
   
   Set oRS = oDas.OpenRecordset(sSql, adOpenForwardOnly, adLockReadOnly)
   
   If Not oRS Is Nothing Then
      If Not oRS.BOF And Not oRS.EOF Then
         CDRPlusCheckForNewerEncounter = CBool(oRS(0) > 0)
      End If
   End If
   
ErrorHandler:
   Set oDas = Nothing

   If CBool(Err) Then HandleErrors Err, MODULE_NAME, "CDRPlusCheckForNewerEncounter"
   
End Function

Public Property Get RTF(ByVal bExtended As Boolean) As String

On Error GoTo ErrorHandler

Dim sRTF As String
Dim lIdx As Long

   If bExtended Then
      
      If PregnancyIndicator Or HysterectomyIndicator Or MenopauseIndicator Or BirthControlMethods.Count > 0 Or _
         Gravida >= 0 Or Para >= 0 Or TotalAbortus >= 0 Or LivingChildren >= 0 Then
              
         'There is Female Only data so build RTF
         sRTF = "{\rtf1\ansi\ucl\deff0\deflang1033\deflangfe1033" & _
                  "{\fonttbl{\f0\fswiss\fcharset0\fprq2 Arial;}}" & _
                  "{\colortbl;\red0\green0\blue0;}" & _
                  "\pard\f0\fs16" & _
                  " G" & IIf(Gravida >= 0, Gravida, vbNullString) & _
                  " P" & IIf(Para >= 0, Para, vbNullString) & _
                  " A" & IIf(TotalAbortus >= 0, TotalAbortus, vbNullString) & _
                  " LC" & IIf(LivingChildren >= 0, LivingChildren, vbNullString) & "."
                  
         If PregnancyIndicator Then sRTF = sRTF & "  Pregnant."
         If MenopauseIndicator Then sRTF = sRTF & "  Post-menopause."
         If HysterectomyIndicator Then sRTF = sRTF & "  Post-hyserectomy."
                  
         If LastMenstrualPeriod <> NULL_DATE Then sRTF = sRTF & "  LMP: " & Format(LastMenstrualPeriod, gsDefaultDate) & "."
         If EstimatedDOB <> NULL_DATE Then sRTF = sRTF & "  EDC: " & Format(EstimatedDOB, gsDefaultDate) & "."
            
         If BirthControlMethods.Count > 0 Then
            sRTF = sRTF & "  Birth Control Method:  " & GetBCMethodDesc(BirthControlMethods.Keys(0))
            For lIdx = 1 To BirthControlMethods.Count - 1
               sRTF = sRTF & ", " & GetBCMethodDesc(BirthControlMethods.Keys(lIdx))
            Next
            sRTF = sRTF & "."
         End If
         sRTF = sRTF & "\par\pard\tx3690\tx6480\pard }}"
         
      Else
         sRTF = vbNullString
      End If
      
   Else
      
      If PregnancyIndicator Then
         sRTF = "{\pard\plain\f0\fs16\b Pregnant.\b0 " & _
               IIf(LastMenstrualPeriod <> NULL_DATE, "  LMP: " & Format(LastMenstrualPeriod, gsDefaultDate) & ".", vbNullString) & _
               IIf(EstimatedDOB <> NULL_DATE, "  EDC: " & Format(EstimatedDOB, gsDefaultDate) & ".", vbNullString) & "\par }"
      Else
         sRTF = vbNullString
      End If
         
   End If
   
   RTF = sRTF
   
ErrorHandler:
   If CBool(Err) Then HandleErrors Err, MODULE_NAME, "RTF"
   
End Property

Private Function GetBCMethodDesc(ByVal sBCMethodNCID As String) As String
   Select Case sBCMethodNCID
      Case "0":            GetBCMethodDesc = "None"
      Case "24455":        GetBCMethodDesc = "Abstinence"
      Case "92466":        GetBCMethodDesc = "Birth Control Pill"
      Case "177213":       GetBCMethodDesc = "Condom"
      Case "3000257516":   GetBCMethodDesc = "Diaphragm"
      Case "111123":       GetBCMethodDesc = "Foam"
      Case "3000257521":   GetBCMethodDesc = "Intrauterine Device (IUD)"
      Case "14518433":     GetBCMethodDesc = "Intramuscular Injection (e.g. Depoprovera)"
      Case "111111":       GetBCMethodDesc = "Natural Family Planning Method"
      Case "14503055":     GetBCMethodDesc = "Sterilization (Tubal Ligation)"
      Case "14514091":     GetBCMethodDesc = "Sterilization (Current Partner Has Vasectomy)"
      Case "14501453":     GetBCMethodDesc = "Subdermal Implant (e.g. Norplant)"
      Case "34882":        GetBCMethodDesc = "Other"
      Case "1016299":      GetBCMethodDesc = "NOT APPLICABLE (HYSTERECTOMY)"
      Case Else:           GetBCMethodDesc = "Unknown Method"
   End Select
End Function

Friend Function GetSigningWarningText() As String

On Error GoTo ErrorHandler

Dim sTemp As String
   
   GetSigningWarningText = vbNullString
   
   If PregnancyIndicator And Not Completed Then
      If Not Validate(True, sTemp) Then
         GetSigningWarningText = " - Female Only Data" & vbCrLf & Space(6) & sTemp & vbCrLf
      End If
   End If
   
ErrorHandler:
   If CBool(Err) Then HandleErrors Err, MODULE_NAME, "GetSigningWarningText"
   
End Function

Public Sub InitFromCopy(ByVal oSource As EncPregnancy)
'Initialize the object from a copy but don't change any of the persistence data

Dim lIdx As Long

   mbPregnancyIndicator = oSource.PregnancyIndicator
   mdLastPeriod = oSource.LastMenstrualPeriod
   mdEstDOB = oSource.EstimatedDOB
   mbMenopauseIndicator = oSource.MenopauseIndicator
   mbHysterectomyIndicator = oSource.HysterectomyIndicator
   mlTotalAbortus = oSource.TotalAbortus
   mlSpontaneousAbortions = oSource.SpontaneousAbortions
   mlElectiveAbortions = oSource.ElectiveAbortions
   mlEctopicAbortions = oSource.EctopicAbortions
   mlGravida = oSource.Gravida
   mlPara = oSource.Para
   mlLivingChildren = oSource.LivingChildren
   
   'Initialize Birth Control Properties
   Set moBirthControlMethods = New Dictionary
   For lIdx = 0 To oSource.BirthControlMethods.Count - 1
      moBirthControlMethods.Add oSource.BirthControlMethods.Keys(lIdx), oSource.BirthControlMethods.Items(lIdx)
   Next
   
   'Initialize if this is a new object instance, otherwise we want to retain info about what exists in the database
   If moSavedBCMethods Is Nothing Then Set moSavedBCMethods = New Dictionary

   'Initialize Patient Object Properties
   Set moPatient = oSource.Patient
   mlPatientAge = ConvertAge(moPatient.Age)
   msPatientSex = moPatient.Sex
   
   mbDirty = mbDirty Or oSource.IsDirty

End Sub

Friend Property Get Patient() As MMMHISPatient.Patient
   'Used for copy function
   Set Patient = moPatient
End Property
