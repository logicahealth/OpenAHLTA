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
Attribute VB_Name = "ExamsBL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mcolExams As Collection
Private mobjDas As ICHCSII_DataAccess.ICHCSII_DAS
'
Private mblnUseLocalDB As Boolean
Private mobjDentalParent As CHCSII_Dental.Dental
Private mrsExams As ADODB.Recordset
Private mrsMeasurements As ADODB.Recordset
Private mlngPatientId As Long
Private mblnUseStoredProcs As Boolean
Private mobjComm As Comm
'

'Exams were loaded in descending order (by Date)
Public Property Get LastExam() As ExamBL
    If Not mcolExams Is Nothing Then
        Set LastExam = mcolExams(1)
    End If
End Property
Public Property Get FirstExam() As ExamBL
    On Error Resume Next
    Set FirstExam = mcolExams(mcolExams.Count)
    On Error GoTo 0
End Property

Public Property Get PatientID() As Long
    PatientID = mlngPatientId
End Property
Public Property Let PatientID(lngPatientID As Long)
    mlngPatientId = lngPatientID
End Property

Friend Property Get ExamsRS() As ADODB.Recordset
    Set ExamsRS = mrsExams
End Property
Friend Property Set ExamsRS(rs As ADODB.Recordset)
    Set mrsExams = rs
End Property
Friend Property Get MeasurementsRS() As ADODB.Recordset
    Set MeasurementsRS = mrsMeasurements
End Property
Friend Property Set MeasurementsRS(rs As ADODB.Recordset)
    Set mrsMeasurements = rs
End Property

Public Property Get DentalParent() As CHCSII_Dental.Dental
    Set DentalParent = mobjDentalParent
End Property
Public Property Set DentalParent(objDentalParent As CHCSII_Dental.Dental)
Dim objExam As ExamBL

    Set mobjDentalParent = objDentalParent
    
    'Set the patient ID
    If Not objDentalParent Is Nothing Then
        mlngPatientId = objDentalParent.sPtnID
    End If
    
    For Each objExam In mcolExams
        Set objExam.DentalParent = objDentalParent
    Next
    
End Property

Public Property Get UseLocalDB() As Boolean
    UseLocalDB = mblnUseLocalDB
End Property
Public Property Let UseLocalDB(blnUseLocalDB As Boolean)
    mblnUseLocalDB = blnUseLocalDB
End Property
Public Property Get CommObject() As Comm
   Set CommObject = mobjComm
End Property
Public Property Set CommObject(objComm As Comm)
   Set mobjComm = objComm
End Property


Public Function Item(lngIndex As Long) As ExamBL
    Dim strKey As String
    Dim objExam As ExamBL
    
    strKey = GetKey(lngIndex)
    
    On Error Resume Next  'don't return an item if the requested exam doesn't exist
    Set objExam = mcolExams(strKey)
    On Error GoTo 0
    
    Set Item = objExam
End Function
Public Function Count() As Integer
    Count = mcolExams.Count
End Function
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mcolExams.[_NewEnum]
End Function
Friend Function Add(objExam As ExamBL)
    mcolExams.Add objExam
End Function

Private Sub Class_Initialize()
    Set mcolExams = New Collection
    
End Sub

Private Sub Class_Terminate()
    Set mcolExams = Nothing
End Sub

Public Sub UpdateCurrentNotesPrognosis()
Dim objDL As PerioDL
Dim objExam As ExamBL

    Set objDL = GetPerioDL
    
    Set objExam = mcolExams(GetKey(DentalParent.psDentalEncounterID))
    objDL.LoadNotesPrognosis objExam
End Sub

Public Function LoadMeasurements(objExam As ExamBL)
Dim objDL As PerioDL

    If Not objExam.MeasurementsLoaded Then
        Set objDL = GetPerioDL
        
        objDL.LoadPerioPaths objExam, mrsExams
        objDL.LoadPerioMeasurements objExam, mrsMeasurements
        objDL.LoadOralState objExam
        objExam.AssessChartability
        objExam.MeasurementsLoaded = True
        Set objDL = Nothing
    End If
    
End Function

Public Function Load(lngPatientUnitID As Long) As Boolean
Dim objDL As PerioDL
Dim objRS As Recordset
Dim colExams As ExamsBL
Dim objExam As ExamBL

    'retrieve exam data
    'DummyLoad
    
    Set objDL = GetPerioDL
    
    'get the existing perio exams
    Set colExams = objDL.GetPerioExams(lngPatientUnitID)
    
    Set Me.ExamsRS = colExams.ExamsRS
    Set Me.MeasurementsRS = colExams.MeasurementsRS
    
    For Each objExam In colExams
        Set objExam.DentalParent = mobjDentalParent
        objExam.PatientID = lngPatientUnitID
       mcolExams.Add objExam, GetKey(objExam.DentalEncounterID)
    Next
    
    Set objDL = Nothing
    Set objRS = Nothing
    Set colExams = Nothing
    Set objExam = Nothing
    
    Load = True
End Function

Public Function CreateNewExam(lngDentEncID As Long, dtDentEncDate As Date) As ExamBL
Dim objDL As PerioDL
Dim objExam As ExamBL
Dim strTeethPath As String
Dim i As Integer
    
    'init teeth path to all 32 secondary teeth
    
    For i = 1 To 31
        strTeethPath = strTeethPath & i & ","
    Next i
    strTeethPath = strTeethPath & 32
    
    Set objDL = GetPerioDL
    
    Set objExam = objDL.InsertNewExam(lngDentEncID, perInitial, strTeethPath)
    objExam.InitialLoad = True
    objExam.ExamDate = dtDentEncDate
    Set objExam.DentalParent = mobjDentalParent
    
    mcolExams.Add objExam, GetKey(objExam.DentalEncounterID)
    
    'Create the measurement entries
    Set objExam.Mouth.PerioDL = objDL
    objExam.Mouth.CreateNewMouth objExam.DentalEncounterID
    
    'No longer need to load oral state at this point, it will be done later
    'objDL.LoadOralState objExam
    'objExam.InitialLoad = False
    Set CreateNewExam = objExam
End Function

Public Function SaveExam(lngDentEncID As Long) As Boolean
Dim objDL As PerioDL
Dim objExam As ExamBL


    Set objDL = GetPerioDL
    Set objExam = mcolExams(GetKey(lngDentEncID))
    objDL.SaveExam objExam
    
    SaveExam = True
    
End Function

Private Function DummyLoad() As Boolean
Dim objExam As ExamBL
Dim i As Integer


    Set mcolExams = New Collection

    Set objExam = New ExamBL
    objExam.ExamDate = "3/1/2001"
    mcolExams.Add objExam, GetKey(0)
    
    Set objExam = New ExamBL
    objExam.ExamDate = "9/1/2001"
    mcolExams.Add objExam, GetKey(1)

    Set objExam = New ExamBL
    objExam.ExamDate = "11/1/2001"
    mcolExams.Add objExam, GetKey(2)

End Function

Private Function GetKey(lngIndex As Long)
    GetKey = "Exam" & lngIndex
End Function

Private Function GetPerioDL() As PerioDL
Static objDL As PerioDL
    If objDL Is Nothing Then
        Set objDL = New PerioDL
        objDL.UseLocalDB = mblnUseLocalDB
        objDL.UseStoredProcs = mblnUseStoredProcs
        'SCR 33622
        If mobjComm Is Nothing Then
            Set mobjComm = New Comm
        End If
        Set objDL.CommObject = mobjComm
    End If
    Set GetPerioDL = objDL
End Function

Public Sub Dispose()
Dim objExam As ExamBL
Dim i As Integer

    For Each objExam In mcolExams
        objExam.Dispose
    Next objExam
    
    For i = mcolExams.Count To 1 Step -1
        mcolExams.Remove i
    Next i
    Set mcolExams = Nothing
End Sub

Public Property Get UseStoredProcs() As Boolean
   UseStoredProcs = mblnUseStoredProcs
End Property
Public Property Let UseStoredProcs(blnUse As Boolean)
   mblnUseStoredProcs = blnUse
End Property
