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
Attribute VB_Name = "FormCPGPOC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private moCollRecSet As Collection
Private moCollFormulas As Collection

Private moPatientDemographics As PatientDemographics

Private Const NCID_PATIENT_PLANOFCARE As String = "1999992" 'NEED TO GET A REAL ONE FROM 3M

Private Const gsMainDelim As String = "|"
Private Const msEscChr As String = "\"

Private Function InCPGS(sUnitNumber As String) As Boolean

On Error GoTo ErrHandler

InCPGS = False

Dim sStmt As String


sStmt = "select 'x' from CPGREGISTRY where unit_number = " & sUnitNumber

goDBSqlOps.Execute sStmt

If Not goDBSqlOps.EOF Then
    InCPGS = True
End If

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "FormCPGPOC.InCPGS", "CHCSII_Forms", vbExclamation
CleanUp:
    Exit Function
End Function


Private Function xBuildMgmtString(colMgmt As Collection, _
                                 sSelfMgmtPlanText As String) As Boolean

On Error GoTo ErrHandler

xBuildMgmtString = False

Dim sSelfMgmtPlan As String
Dim vMgmtTemp As Variant
Dim sMgmt As String
Dim sTempRTF As String
Dim sSearch As String
Dim sHeader As String

    
sSelfMgmtPlan = ""
sSearch = "{\colortbl\red0\green0\blue0;}"

sHeader = "{\rtf1\ansi\deff0\deftab720{\fonttbl{\f0\fswiss MS Sans Serif;}{\f1\froman\fcharset2 Symbol;}{\f2\fswiss MS Sans Serif;}}"
sHeader = sHeader & vbCrLf & "{\colortbl\red0\green0\blue0;}" & vbCrLf

For Each vMgmtTemp In colMgmt
    sMgmt = CStr(vMgmtTemp)
    
    If Not Trim(sMgmt) = "" Then
        sTempRTF = sMgmt
        If InStr(1, sTempRTF, sSearch, vbTextCompare) > 0 Then
           If Not Len(sTempRTF) <= InStr(1, sTempRTF, sSearch, vbTextCompare) + Len(sSearch) Then
              sTempRTF = right(sTempRTF, Len(sTempRTF) - (InStr(1, sTempRTF, sSearch, vbTextCompare) + Len(sSearch) + 1))
           End If
        End If
        Do:
            If Trim(sTempRTF) = "" Then Exit Do
            If Not (Asc(right(sTempRTF, 1)) = 10 Or Asc(right(sTempRTF, 1)) = 13 Or right(sTempRTF, 1) = "}") Then Exit Do
            sTempRTF = left(sTempRTF, Len(sTempRTF) - 1)
        Loop
        If Not Trim(sTempRTF) = "" Then
           sSelfMgmtPlan = sSelfMgmtPlan & sTempRTF
           If InStr(1, sTempRTF, "\par", vbTextCompare) = 0 Then sSelfMgmtPlan = sSelfMgmtPlan & "\par "
        End If
     End If
Next

If Not Trim(sSelfMgmtPlan) = "" Then sSelfMgmtPlan = sHeader & sSelfMgmtPlan & "}"

sSelfMgmtPlanText = sSelfMgmtPlan

xBuildMgmtString = True
GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "FormCPGPOC.xBuildMgmtString", "CHCSII_Forms", vbExclamation
CleanUp:

End Function
Private Function BuildMgmtString(colMgmt As Collection, _
                                 sSelfMgmtPlanText As String) As Boolean

On Error GoTo ErrHandler

BuildMgmtString = False

Dim sSelfMgmtPlan As String
Dim vMgmtTemp As Variant
Dim sMgmt As String
Dim sTempRTF As String
Dim sSearch As String
Dim sHeader As String

    
sSelfMgmtPlan = ""
    sSearch = "{\colortbl\red0\green0\blue0;}"
    sSearch = "{\colortbl\red0\green0\blue0;}"
    
    sHeader = "{\rtf1 " '\ansi\deff0\deftab720{\fonttbl{\f0\fswiss MS Sans Serif;}{\f1\froman\fcharset2 Symbol;}{\f2\fswiss MS Sans Serif;}}"
    'sHeader = sHeader & vbCrLf & "{\colortbl\red0\green0\blue0;}" & vbCrLf
 
For Each vMgmtTemp In colMgmt
    sMgmt = CStr(vMgmtTemp)
    
    If Not Trim(sMgmt) = "" Then
        sTempRTF = sMgmt
        'If InStr(1, sTempRTF, sSearch, vbTextCompare) > 0 Then
        '   If Not Len(sTempRTF) <= InStr(1, sTempRTF, sSearch, vbTextCompare) + Len(sSearch) Then
        If left(sTempRTF, 6) = "{\rtf1" Then
              'sTempRTF = Right(sTempRTF, Len(sTempRTF) - (InStr(1, sTempRTF, sSearch, vbTextCompare) + Len(sSearch) + 1))
              sTempRTF = right(sTempRTF, Len(sTempRTF) - 6)
        End If
        Do:
            If Trim(sTempRTF) = "" Then Exit Do
            If Not (Asc(right(sTempRTF, 1)) = 10 Or Asc(right(sTempRTF, 1)) = 13 Or right(sTempRTF, 1) = "}") Then Exit Do
            sTempRTF = left(sTempRTF, Len(sTempRTF) - 1)
        Loop
        If Not Trim(sTempRTF) = "" Then
           sSelfMgmtPlan = sSelfMgmtPlan & sTempRTF
           If InStr(1, sTempRTF, "\par", vbTextCompare) = 0 Then sSelfMgmtPlan = sSelfMgmtPlan & "\par "
        End If
     End If
Next

'If Not Trim(sSelfMgmtPlan) = "" Then sSelfMgmtPlan = sHeader & sSelfMgmtPlan & "}"
If Not Trim(sSelfMgmtPlan) = "" Then sSelfMgmtPlan = sHeader & sSelfMgmtPlan & "}"


sSelfMgmtPlanText = sSelfMgmtPlan

BuildMgmtString = True
GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "FormCPGPOC.BuildMgmtString", "CHCSII_Forms", vbExclamation
CleanUp:

End Function

Public Function DeserializeString(rStr As String, rsDelim As String) As Collection
    Dim cResult As Collection
    Dim x As Integer
    Dim sField As String
    Dim sChr As String
    Set cResult = New Collection
    For x = 1 To Len(rStr)
        sChr = Mid(rStr, x, 1)
        Select Case sChr
        Case msEscChr
            sField = sField & Mid(rStr, x + 1, 1)
            x = x + 1
        Case rsDelim
            cResult.Add sField
            sField = ""
        Case Else
            sField = sField & sChr
        End Select
    Next x
    Set DeserializeString = cResult
End Function
Private Function GetPOCGoals(sUnitNumber As String, _
                             oRec As ADODB.Recordset) As Boolean

    Dim sRTF As String
    GetPOCGoals = False
    
    Dim oPatientPOC As Object
    
    Set oPatientPOC = cwobjCWComm.InitializeOLEServer(NCID_PATIENT_PLANOFCARE)
    sRTF = oPatientPOC.GetEncounterContRTF

    GetPOCGoals = True
    GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "FormCPGPOC.GetPOCGoals", "CHCSII_Forms", vbExclamation
CleanUp:
    Set oPatientPOC = Nothing
    
End Function



Private Function xGetSelfMgmtText(sUnitNumber As String, _
                                 oRec As ADODB.Recordset) As Boolean

On Error GoTo ErrHandler
xGetSelfMgmtText = False

Dim sStmt As String
Dim i As Integer
Dim colMgmt As Collection
Dim colParams As Collection
Dim sSelfMgmtPlanText As String
Dim oForm As frmSF600
Dim sSelfMgmtText As String

Set colMgmt = New Collection
'-- Get Patient Management Plan Goals Text
'-----------------------------------------------------------------------------------
sStmt = "select mgmt_plan from cpg_patient_goals where unit_number = " & sUnitNumber

goDBSqlOps.Execute sStmt

Do While Not goDBSqlOps.EOF
    colMgmt.Add goDBSqlOps.Value("mgmt_plan")
Loop
    
'-- Get POC Goals Text
'-----------------------------------------------------------------------------------
sStmt = "select d.PARAMS  " & vbNewLine & _
        "from CPGREGISTRY r, POCDEF d " & vbNewLine & _
        "where d.CPG_ID = r.CPG_ID and r.UNIT_NUMBER = " & sUnitNumber
    
goDBSqlOps.Execute sStmt

If Not (goDBSqlOps.BOF And goDBSqlOps.EOF) Then
   goDBSqlOps.MoveFirst
   Do Until goDBSqlOps.EOF
        
        Set colParams = DeserializeString(goDBSqlOps.Value("PARAMS"), gsMainDelim)
        
        If colParams.Count > 0 Then ' the first item is the action plan, rest are criteria
           colMgmt.Add Trim(colParams(1))
        End If
        
        goDBSqlOps.MoveNext
   Loop
End If

'-- Get RTF String
'-----------------------------------------------------------------------------------
If Not BuildMgmtString(colMgmt, sSelfMgmtPlanText) Then GoTo CleanUp
DoEvents
'Set oForm = New frmSF600

Load frmRTF
DoEvents
'Load oForm
DoEvents
On Error Resume Next
frmRTF.RichTextBox1.TextRTF = sSelfMgmtPlanText

'oForm.txBody.RTFSelText = sSelfMgmtPlanText
DoEvents
sSelfMgmtText = frmRTF.RichTextBox1.Text
DoEvents
Unload frmRTF
'-- Build Recordset
'-----------------------------------------------------------------------------------
Set oRec = New ADODB.Recordset
oRec.CursorLocation = adUseClient

'oRec.Fields.Append "Unit_Number", adLongVarWChar, 10
oRec.Fields.Append "mgmt_text", adLongVarWChar, 8000
oRec.Open

oRec.AddNew
'oRec.Fields("Unit_Number")0.Value = sUnitNumber & Chr(0)
oRec.Fields("mgmt_text").Value = sSelfMgmtText & Chr(0)
oRec.Update

xGetSelfMgmtText = True
GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "FormCPGPOC.xGetSelfMgmtText", "CHCSII_Forms", vbExclamation
CleanUp:
    Set oForm = Nothing
End Function

Private Function GetSelfMgmtText(sUnitNumber As String, _
                                 oRec As ADODB.Recordset) As Boolean

On Error GoTo ErrHandler
GetSelfMgmtText = False

Dim sStmt As String
Dim i As Integer
Dim colMgmt As Collection
Dim colParams As Collection
Dim sSelfMgmtPlanText As String
Dim oForm As frmSF600
Dim sSelfMgmtText As String

Set colMgmt = New Collection
'-- Get Patient Management Plan Goals Text
'-----------------------------------------------------------------------------------
sStmt = "select mgmt_plan from cpg_patient_goals where unit_number = " & sUnitNumber
goDBSqlOps.Execute sStmt
If Not (goDBSqlOps.BOF And goDBSqlOps.EOF) Then
   goDBSqlOps.MoveFirst
   Do Until goDBSqlOps.EOF
       colMgmt.Add goDBSqlOps.Value("mgmt_plan")
       goDBSqlOps.MoveNext
   Loop
End If

'-- Get POC Goals Text
'-----------------------------------------------------------------------------------
sStmt = "select d.PARAMS  " & vbNewLine & _
        "from CPGREGISTRY r, POCDEF d " & vbNewLine & _
        "where d.CPG_ID = r.CPG_ID and r.UNIT_NUMBER = " & sUnitNumber
    
goDBSqlOps.Execute sStmt
If Not (goDBSqlOps.BOF And goDBSqlOps.EOF) Then
   goDBSqlOps.MoveFirst
   Do Until goDBSqlOps.EOF
       colMgmt.Add Trim(goDBSqlOps.Value("PARAMS"))
       goDBSqlOps.MoveNext
   Loop
End If

'-- Get RTF String
'-----------------------------------------------------------------------------------
If Not BuildMgmtString(colMgmt, sSelfMgmtPlanText) Then GoTo CleanUp
DoEvents
'Set oForm = New frmSF600

Load frmRTF
DoEvents
'Load oForm
DoEvents
On Error Resume Next
frmRTF.RichTextBox1.TextRTF = sSelfMgmtPlanText

'oForm.txBody.RTFSelText = sSelfMgmtPlanText
DoEvents
sSelfMgmtText = frmRTF.RichTextBox1.Text
DoEvents
Unload frmRTF
'-- Build Recordset
'-----------------------------------------------------------------------------------
Set oRec = New ADODB.Recordset
oRec.CursorLocation = adUseClient

'oRec.Fields.Append "Unit_Number", adLongVarWChar, 10
oRec.Fields.Append "mgmt_text", adLongVarWChar, 8000
oRec.Open

oRec.AddNew
'oRec.Fields("Unit_Number")0.Value = sUnitNumber & Chr(0)
oRec.Fields("mgmt_text").Value = sSelfMgmtText & Chr(0)
oRec.Update

GetSelfMgmtText = True
GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "FormCPGPOC.GetSelfMgmtText", "CHCSII_Forms", vbExclamation
CleanUp:
    Set oForm = Nothing
End Function

Private Sub Class_Initialize()
    Set moPatientDemographics = New PatientDemographics
End Sub

Private Sub Class_Terminate()
    Set moCollRecSet = Nothing
    Set moCollFormulas = Nothing
    Set moPatientDemographics = Nothing
End Sub

Public Function Build(sPatientUnitNumber As String, _
                      oReport As CRAXDRT.Report, _
                      Optional robjComm As Object) As Boolean

    On Error GoTo ErrHandler
    Build = False

    Dim oCrystal As Crystal
    Dim oRec As ADODB.Recordset
    Dim sSponsorUnitNumber As String
    Dim sUnitNumber As String
    Dim oGoals As Object
    Dim sSelfMgmtText As String
    DoEvents


    If InCPGS(sPatientUnitNumber) Then
        '-- Get Data
        '----------------------------------------------------------------------------
        Set moCollRecSet = Nothing
        Set moCollRecSet = New Collection
    
        '-- Patient Demographics
        '----------------------------------------------------------
        If Not moPatientDemographics.RecordGetDemographics(sPatientUnitNumber, oRec) Then GoTo CleanUp
        moCollRecSet.Add oRec, "demographics"
    
        '-- Sponsor Demographics
        '----------------------------------------------------------
        If oRec.Fields.Item("FMP").Value = "20" Then
            sSponsorUnitNumber = sPatientUnitNumber
        Else
            sSponsorUnitNumber = moPatientDemographics.GetUnitNumber(oRec.Fields.Item("Sponsor_SSN").Value)
        End If
        sSponsorUnitNumber = moPatientDemographics.GetUnitNumber(oRec.Fields.Item("Sponsor_SSN").Value)
        If Not moPatientDemographics.RecordGetDemographics(sSponsorUnitNumber, oRec) Then GoTo CleanUp
        
        moCollRecSet.Add oRec, "sponsor"
        
        '-- Get Patient Plan of Care
        '----------------------------------------------------------
        If Not GetPOCGoals(sPatientUnitNumber, oRec) Then GoTo CleanUp
        moCollRecSet.Add oRec, "goals"
        
        '-- Get Patient Self Management Text
        '----------------------------------------------------------
        If Not GetSelfMgmtText(sPatientUnitNumber, oRec) Then GoTo CleanUp
        moCollRecSet.Add oRec, "selfmgmt"
        
    ''    Dim MyPlans As ADODB.Recordset
    ''
    ''    Set oGoals = robjComm.InitializeOLEServer(NCID_PATIENT_PLANOFCARE)
    ''
    ''    Set MyPlans = oGoals.Plans
        
    ''    Dim MyGoals As ADODB.Recordset
    ''    Dim MyResults As ADODB.Recordset
    ''
    ''    Set MyGoals = oGoals.goals
    ''    Set MyResults = oGoals.Results
    
    
    '-- Set Formulas
    '-----------------------------------------------------------------------------
        Set moCollFormulas = Nothing
        Set moCollFormulas = New Collection
        moCollFormulas.Add "'" & Format(Now, gsDefaultFormat & "  HHMM") & "'", "TodaysDateTime"
        
    
    '-- Create Report Object
    '-----------------------------------------------------------------------------
        Set oCrystal = New Crystal
        Set oReport = Nothing
        If Not oCrystal.ReportOpen("CPG_POC.rpt", oReport) Then GoTo CleanUp
        If Not oCrystal.AssignRecordSets(moCollRecSet, oReport) Then GoTo CleanUp
        If Not oCrystal.AssignFormulas(moCollFormulas, oReport) Then GoTo CleanUp
    Else
        MsgBxARMd "Cannot print the plan of care since the patient is not on the registry.", vbOKOnly + vbInformation
        GoTo CleanUp
    End If
    
    Build = True
    GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "FormCPGPOC.Build", "CHCSII_Forms", vbExclamation
    Resume Next
CleanUp:
    Set oCrystal = Nothing
    Set oRec = Nothing
    
End Function


