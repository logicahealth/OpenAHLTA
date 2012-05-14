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
Attribute VB_Name = "clsLabCase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''================================================================================================
''  Class Name: clsLabCase
''
''  Description: This class is used by Dental Lab component to describe,
''               get and store lab case related data
''
''  Properties:
''     msPatientUnitID         to store    PatientUnitID
''     msEncounterFacilityNCID to store    EncounterFacilityNCID
''     msEncounterNumber       to store    EncounterNumber
''     msDentalEncounterID     to store    DentalEncounterID
''     msProviderID            to store    ProviderID
''     msProviderName          to store    ProviderName
''
''     msLabCaseID             to store    LabCaseID
''     msLabCaseNo             to store    LabCaseNo
''     msDateCaseCompleted     to store    DateCaseCompleted
''     msStatus                to store    Status
''     msComments              to store    Comments
''     msCreateDateTime        to store    CreateDateTime
''     msCaseLinkedTo          to store    CaseLinkedTo
''
''     mobjLabTXs              link to LabTXs object
''
''  Methods:
''    ClassInitial           - to initialize class members
''    FillFromRecordset      - to get values from recordset's fields
''    FillToRecordset        - to store values in recordset's fields
''    CheckLabCaseNo         - to verify lab case number for newly created lab case
''    GenerateCaseNo         - to generate new lab case number
''    GetLabCaseChildMembers - to retrieve all tx and tasks data related to lab case
''================================================================================================
Option Explicit

#If DEBUGON Then
'-- Private DR As DebugRec
'-- Public DL As cDebugLog
'-- Public Const Exe = "OrderSets"
Private Const Module = "Dental.cls"
#End If

'////////////////////  C L A S S   P R O P E R T I E S    S E C T I O N  ////////////////////
'Encounter info
Private msPatientUnitID         As String    'PatientUnitID
Private msEncounterFacilityNCID As String    'EncounterFacilityNCID
Private msEncounterNumber       As String    'EncounterNumber
Private msDentalEncounterID     As String    'DentalEncounterID
Private msProviderID            As String    'ProviderID
Private msProviderName          As String    'ProviderName

'LabCase info
Private msLabCaseID             As String    'LabCaseID
Private msLabCaseNo             As String    'LabCaseNo
Private msDateCaseCompleted     As String    'DateCaseCompleted
Private msStatus                As String    'Status
Private msComments              As String    'Comments
Private msCreateDateTime        As String    'CreateDateTime
Private msCaseLinkedTo          As String    'CaseLinkedTo

'link to LabTXs object which has private LabTX collection
Private mobjLabTXs              As clsLabTXs 'LabTXs object

'link to the object LabDentSpec
Private mobjLabDentSpec As clsLabDentSpec

' Storage for the debug ID
Private mlngDebugID As Long

Private Sub Class_Initialize()

On Error GoTo ErrHandler

'--- initialize all data members
    msPatientUnitID = "0"
    msEncounterFacilityNCID = "0"
    msEncounterNumber = "0"
    msProviderID = ""
    msProviderName = ""
    msDentalEncounterID = ""

    msLabCaseID = ""
    msLabCaseNo = ""
    msDateCaseCompleted = ""
    msStatus = ""
    msComments = ""
    msCreateDateTime = ""
    msCaseLinkedTo = "0"

    Set mobjLabTXs = Nothing
   
   'abstract class
   Set mobjLabDentSpec = Nothing

   mlngDebugID = DebugSerial
   ' Add a string entry to the global collection
   gcolDebug.Add "LabCase Initialize; DebugID =" & DebugID, CStr(DebugID)

Exit Sub

ErrHandler:
   Call gobjShared.ShowVBError(Err.Number, Err.Description, _
                "clsLabCase.Initialize", "Labs", vbCritical)
End Sub

Private Sub Class_Terminate()
On Error GoTo ErrHandler

    Set mobjLabTXs = Nothing
    
    Set mobjLabDentSpec = Nothing
    
    'Remove the string entry to make sure that the object is not around any more
    gcolDebug.Remove CStr(DebugID)

Exit Sub

ErrHandler:
   Call gobjShared.ShowVBError(Err.Number, Err.Description, _
                "clsLabCase.Terminate", "Labs", vbCritical)

End Sub

Public Sub ClassInitial()

'On Error GoTo ErrHandler
    
    msPatientUnitID = guCurEncInfo.PatientUnitID
    'SCR 44705 RWalker 11/07/2003
    'Use the current user's info if a provider otherwise use the encounter provider info
    If Sign_CoSign_Check(cwobjLogonEnviron.UserNCID) Then
        msProviderName = cwobjLogonEnviron.UserName
        msProviderID = cwobjLogonEnviron.UserNCID
    Else
        msProviderID = guCurEncInfo.ProviderID
        msProviderName = guCurEncInfo.ProviderFullName
    End If
    
    msEncounterFacilityNCID = guCurEncInfo.EncounterFacilityNCID
    msEncounterNumber = guCurEncInfo.EncounterNumber
    msDentalEncounterID = guCurEncInfo.DentalEncounterID
    
    Set mobjLabTXs = New clsLabTXs
    Set mobjLabDentSpec = New clsLabDentSpec
    
End Sub
Public Function Sign_CoSign_Check(ByRef UserNCID As String) As Boolean
    
'SCR 44705 RWalker 11/07/2003 New code to verify the logged on user's signature ability

    Dim sLogin      As String
    Dim sMsg        As String
    Dim CanSign     As CHCSII_PrivilegeType
    Dim NeedCosign  As CHCSII_PrivilegeType
    Dim bCansign    As Boolean
    Dim bNeedCosign As Boolean

    On Error GoTo ErrHandler

    sLogin = UserNCID
    CanSign = Priv_Sign
    NeedCosign = Priv_Need_Cosign

    If sLogin = "" Then
        Debug.Print "No sLogin found for selected provider."
        Exit Function
    End If

'SCR 45346 RWalker 11/24/2003
    Sign_CoSign_Check = False
    If Not gobjShared.IsAppMode(modeITT) Then
        sMsg = "HasPrivilege|" & sLogin & "|" & Priv_Current_Encounter & "|" & CanSign
        Call cwobjCWComm.Message(cwiGENERIC, sMsg, NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION)
        If sMsg = "TRUE" Then
            bCansign = True
        End If
        sMsg = "HasPrivilege|" & sLogin & "|" & Priv_Current_Encounter & "|" & NeedCosign
        Call cwobjCWComm.Message(cwiGENERIC, sMsg, NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION)
        If sMsg = "TRUE" Then
            bNeedCosign = True
        End If
    Else
        bCansign = True
    End If

    If bNeedCosign = True Or bCansign = True Then
        Sign_CoSign_Check = True
    End If

#If DEBUGON Then
   Dim DR  As DebugRec
   If Not DL Is Nothing Then DL.DebugLog "CHCSII_Dental_Labs", "clsLabCase", "Sign_CoSign_Check", "Labs", DR
#End If

    Exit Function

ErrHandler:
'SCR 45346 RWalker 11/24/2003
'   Call gobjShared.ShowVBError(Err.Number, Err.Description, _
                "clsLabCase.Sign_CoSign_Check", "Labs", vbCritical)

End Function


Public Property Get LabTXs() As clsLabTXs
    Set LabTXs = mobjLabTXs
End Property

Public Property Set LabTXs(robjValue As clsLabTXs)
    Set mobjLabTXs = robjValue
End Property

Public Property Let LabCaseNo(ByVal vsValue As String)
    msLabCaseNo = vsValue
End Property

Public Property Get LabCaseNo() As String
    LabCaseNo = msLabCaseNo
End Property

Public Property Let LabCaseID(ByVal vsValue As String)
    msLabCaseID = vsValue
End Property

Public Property Get LabCaseID() As String
    LabCaseID = msLabCaseID
End Property

Public Property Let CreateDateTime(ByVal vsValue As String)
    msCreateDateTime = vsValue
End Property

Public Property Get CreateDateTime() As String
    CreateDateTime = msCreateDateTime
End Property

Public Property Let DateCaseCompleted(ByVal vsValue As String)
    msDateCaseCompleted = vsValue
End Property

Public Property Get DateCaseCompleted() As String
    DateCaseCompleted = msDateCaseCompleted
End Property

Public Property Let CaseLinkedTo(ByVal vsValue As String)
    CaseLinkedTo = vsValue
End Property

Public Property Get CaseLinkedTo() As String
    CaseLinkedTo = msCaseLinkedTo
End Property
'
Public Property Set LabDentSpec(robjValue As clsLabDentSpec)
    Set mobjLabDentSpec = robjValue
End Property

Public Property Get LabDentSpec() As clsLabDentSpec
    Set LabDentSpec = mobjLabDentSpec
End Property

Public Property Let Comments(ByVal vsValue As String)
    msComments = vsValue
End Property

Public Property Get Comments() As String
    Comments = msComments
End Property

Public Property Let DentalEncounterID(ByVal vsValue As String)
    msDentalEncounterID = vsValue
End Property

Public Property Get DentalEncounterID() As String
    DentalEncounterID = msDentalEncounterID
End Property

Public Property Let PatientUnitID(ByVal vsValue As String)
    msPatientUnitID = vsValue
End Property

Public Property Get PatientUnitID() As String
    PatientUnitID = msPatientUnitID
End Property

Public Property Let EncounterFacilityNCID(ByVal vsValue As String)
    msEncounterFacilityNCID = vsValue
End Property

Public Property Get EncounterFacilityNCID() As String
    EncounterFacilityNCID = msEncounterFacilityNCID
End Property

Public Property Let Status(ByVal vsValue As String)
    msStatus = vsValue
End Property

Public Property Get Status() As String
    Status = msStatus
End Property

Public Property Let ProviderID(ByVal vsValue As String)
    msProviderID = vsValue
End Property

Public Property Get ProviderID() As String
    ProviderID = msProviderID
End Property

Public Property Let ProviderName(ByVal vsValue As String)
    msProviderName = vsValue
End Property

Public Property Get ProviderName() As String
    ProviderName = msProviderName
End Property

Public Property Let EncounterNumber(ByVal vsValue As String)
    msEncounterNumber = vsValue
End Property

Public Property Get EncounterNumber() As String
    EncounterNumber = msEncounterNumber
End Property

Property Get DebugID() As Long
   DebugID = mlngDebugID
End Property

Public Sub FillFromRecordset(roRs As Recordset)
''===========================================================================================
''  Description:  This method populates LabCase object's members with data
''                from recordset fields
''
''  Called By: This method is called by LoadLabCase in modLabs module
''
''  Argument(s):
''     roRs - Contains recordset
''
''  Returns:
''
''  Notes:
''
''  Programmer:    Alek Samedov
''===========================================================================================

#If DEBUGON Then
   Dim DR  As DebugRec

   If Not DL Is Nothing Then DL.DebugLog Exe, Module, "FillFromRecordset", "Begin", DR
#End If

'--- error handling ---
On Error GoTo ErrHandler

    '--- fill in the properties from the LAB_CASE table fields ---
    msLabCaseID = roRs("Lab_Case_ID") & ""
    msLabCaseNo = roRs("LAB_CASE_NUMBER") & ""
    msDateCaseCompleted = roRs("Date_CASE_COMPLETED") & ""
    msStatus = roRs("Status") & ""
    msComments = roRs("Comments") & ""
    msProviderID = roRs("ProviderID") & ""
    msCreateDateTime = roRs("CreateDateTime") & ""
    msDentalEncounterID = roRs("DENTAL_ENCOUNTER_ID") & ""
    msCaseLinkedTo = roRs("LINKED") & ""

#If DEBUGON Then
   If Not DL Is Nothing Then DL.DebugLog Exe, Module, "FillFromRecordSet", "End", DR
#End If

Exit Sub

ErrHandler:
Resume Next
   Call gobjShared.ShowVBError(Err.Number, Err.Description, _
                "clsLabCase.FillFromRecordset", "Labs", vbCritical)
End Sub

Public Sub FillToRecordset(roRs As Recordset)
''===========================================================================================
''  Description: This method populates fields of the recordset with data from properties
''
''  Called By: this method is called by ... in ... code
''
''  Argument(s):
''     roRs - Contains recordset
''
''  Returns:
''
''  Notes:
''
''  Programmer:    Alek Samedov
''===========================================================================================

#If DEBUGON Then
   Dim DR  As DebugRec

   If Not DL Is Nothing Then DL.DebugLog Exe, Module, "FillToRecordset", "Begin", DR
#End If


'--- error handling ---
On Error GoTo ErrHandler

If giLabCaseStatus = gcCreateNewLabCase Then
    'create new lab case number
   
   '<<< temp solution
    msLabCaseNo = msLabCaseNo & "-0001"
    '--- fill in the LAB_CASE table fields from the properties
    roRs("LAB_CASE_NUMBER") = msLabCaseNo & ""
    roRs("Date_CASE_COMPLETED") = msDateCaseCompleted
    roRs("Status") = msStatus
    roRs("Comments") = msComments
    roRs("ProviderID") = msProviderID
    roRs("CreateDateTime") = msCreateDateTime
    roRs("DENTAL_ENCOUNTER_ID") = msEncounterNumber
    roRs("LINKED") = msCaseLinkedTo
    
Else
    'update existing record
    roRs("LAB_CASE_NUMBER") = msLabCaseNo
    roRs("Date_CASE_COMPLETED") = msDateCaseCompleted
    roRs("Status") = msStatus
    roRs("Comments") = msComments
    roRs("ProviderID") = msProviderID
    roRs("CreateDateTime") = msCreateDateTime
    roRs("DENTAL_ENCOUNTER_ID") = msEncounterNumber
    roRs("LINKED") = msCaseLinkedTo
End If
    

#If DEBUGON Then
   If Not DL Is Nothing Then DL.DebugLog Exe, Module, "FillToRecordSet", "End", DR
#End If
    
    
Exit Sub

ErrHandler:
    Resume Next
   Call gobjShared.ShowVBError(Err.Number, Err.Description, _
                "clsLabCase.FillToRecordset", "Labs", vbCritical)
End Sub

Public Function GenerateCaseNo() As String
''===========================================================================================
''  Description: This method generates new lab case number for newly created lab case
''
''  Called By: this method is called by LoadLabCase procedure in modLabs code
''
''  Argument(s):
''
''  Returns:   string value to present LabCaseNo properly formatted
''
''  Notes:
''
''  Programmer:    Alek Samedov
''===========================================================================================

On Error GoTo ErrHandler

'KDunne 11/7/2002
'SCR 28782 - Getting dates from Encounter's Apptointment date instead of CDR's
'            sysdate function.
'GenerateCaseNo = Format$(GetSysDateTime(gsDATE_FORMAT_VB), "ddmmmyyyy") & "-" & GetMaxVal("LAB_CASE", "LAB_CASE_ID")
GenerateCaseNo = Format$(gobjEnc.CurrentEncounter.StartDTS, "ddmmmyyyy") & "-" & GetMaxVal("LAB_CASE", "LAB_CASE_ID")
Exit Function

ErrHandler:
   Call gobjShared.ShowVBError(Err.Number, Err.Description, _
                "clsLabCase.GenerateCaseNo", "Labs", vbCritical)
   
Err.Raise Err.Number, Err.Source, Err.Description

End Function

Public Function GetLabCaseChildMembers() As Boolean
''===========================================================================================
''  Description: This method gets the recordset of the lab related treatments
''               and creates the collection of the lab treatments
''
''  Called By: this method is called by LoadLabCase in modLabs module
''
''  Argument(s):
''
''  Returns:
''         TRUE  - if all lab case related data successfully retrieved from database
''         FALSE - if failed
''
''  Notes:
''
''  Programmer:    Alek Samedov
''===========================================================================================

Dim sTableName As String
Dim sCondLabFlag As String
Dim colCond As New Collection

Dim oCurLabTX As clsLabTX

Dim oRs As ADODB.Recordset

Dim i

Dim iChartType As Integer
Dim bFoundAtLeastOnce As Boolean

Dim oLabTX As clsLabTX

Dim oLabTxRs As ADODB.Recordset
Dim sTXCodeID As String

'--- error handling ---
On Error GoTo ErrHandler

'--- set initial return value
GetLabCaseChildMembers = True

bFoundAtLeastOnce = False

'KDunne 11/7/2002
'SCR 28782 - Getting dates from Encounter's Appointment date instead of CDR's
'            sysdate function.
gsDefaultTXDateOrdered = gobjEnc.CurrentEncounter.StartDTS

'====09/02/01===========================
i = 1
For iChartType = gcGraphChart To gcNonGraphChart

   'to get all treatments data from db
   If gobjLabCase.LabTXs.GetLabTXs(oRs, iChartType) Then
      'i.e. number of TXs records is not 0 (gcRowsSelected/"N Rows Selected")
   
      If Not bFoundAtLeastOnce Then bFoundAtLeastOnce = True
      
      oRs.MoveFirst
      'i = 1
      Do While Not oRs.EOF
         gobjLabCase.LabTXs.Add oRs, iChartType
         Set oCurLabTX = gobjLabCase.LabTXs.Item(i)
         
         oCurLabTX.ChartType = iChartType
                  
         sTXCodeID = CStr(oCurLabTX.TxCodeID)
         'for existing (found) TXs the following LabTX properties - LabTxID, TX_DATE_ORDERED, etc.-
         'should be updated with values from LAB_TX_SPECIFICATIONS table
         If FoundLabTX(oLabTxRs, sTXCodeID) Then
            'oLabTxRs recordset contains 1 record from LAB_TX_SPECIFICATIONS table
            'some LabTX properties should be updated with existing data
            oCurLabTX.UpdateFromRecordset oLabTxRs
            
            oCurLabTX.RecordStatus = gcStatusQuo
            
         Else
            'it means that
            'new tx record
            oCurLabTX.RecordStatus = gcINSERT
         End If
         
         'to create task collection for each lab related treatment
         If oCurLabTX.ChartType = iChartType Then
            If Not oCurLabTX.GetLabTXChildMembers(oCurLabTX, iChartType) Then
               GetLabCaseChildMembers = False
               Exit Function
            End If
         End If
         
         Set oCurLabTX = Nothing
         oRs.MoveNext
         i = i + 1
      Loop
      
   End If
   
   If Not oRs Is Nothing Then Set oRs = Nothing
   
Next iChartType

If Not bFoundAtLeastOnce Then
   GetLabCaseChildMembers = False
   Exit Function
End If

Exit Function

ErrHandler:
GetLabCaseChildMembers = False
If Not oCurLabTX Is Nothing Then Set oCurLabTX = Nothing

   Call gobjShared.ShowVBError(Err.Number, Err.Description, _
                "clsLabCase.GetLabCaseChildMembers", "Labs", vbCritical)
    
End Function

