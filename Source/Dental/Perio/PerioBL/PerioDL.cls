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
Attribute VB_Name = "PerioDL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mobjDas As ICHCSII_DataAccess.ICHCSII_DAS
Private mobjSqlOps As ICHCSII_SQL.ISqlOps
Private mblnUseLocalDB As Boolean
Private mblnUseStoredProcs As Boolean
Private mobjComm As Comm
'

Public Property Get UseLocalDB() As Boolean
    UseLocalDB = mblnUseLocalDB
End Property
Public Property Let UseLocalDB(blnUseLocalDB As Boolean)
    mblnUseLocalDB = blnUseLocalDB
End Property

Private Function GetDasObj() As ICHCSII_DAS

'SF 11/17/03
If goCHCSIIConn Is Nothing Then
    Set goCHCSIIConn = New CHCSII_CONN.Conn
End If
If mobjDas Is Nothing Then
    Set mobjDas = goCHCSIIConn.CHCSII_DAS
End If

'If mobjDas Is Nothing Then
'   If mblnUseLocalDB Then
'      Set mobjDas = New GEMS_DAS
'   Else
'      Set mobjDas = New CHCSII_DAS
'   End If
'End If

Set GetDasObj = mobjDas

End Function

Private Function GetSQLOpsObj() As ICHCSII_SQL.ISqlOps


If mobjSqlOps Is Nothing Then
    If goCHCSIIConn Is Nothing Then
        Set goCHCSIIConn = New CHCSII_CONN.Conn
    End If
    Set mobjSqlOps = goCHCSIIConn.CHCSII_SQLOPS
End If

'If mobjSqlOps Is Nothing Then
'   If mblnUseLocalDB Then
'      Set mobjSqlOps = New GEMS_SQL.SqlOps
'   Else
'      Set mobjSqlOps = New CHCSII_SQL.SqlOps
'   End If
'End If
Set GetSQLOpsObj = mobjSqlOps

End Function

Public Function SaveNarrativeSummary(objExam As ExamBL) As Boolean

Dim objNarrSumm As clsNarrSumm
Dim objCurrSection As clsISection
   
    Set objNarrSumm = New clsNarrSumm
    Set objCurrSection = objNarrSumm
    objCurrSection.LoadSection objExam.DentalParent.sPtnID, objExam.DentalEncounterID
   With objNarrSumm
        'SF - 10-16-02 Req 8554, .ExamPerioTabNote is not the correct property.
        .PerioChartingNote = objExam.StatisticsNote
      '.ExamPerioTabNote = objExam.StatisticsNote
   End With
   objCurrSection.SaveSection
End Function
Public Function SaveNotesPrognosis(objExam As ExamBL) As Boolean
'Prognosis is saved as a pipe delimited string of ToothIDs (1-52), no certain order, so
'we can just tack on ToothId if not found

    Dim objPrognosis As clsPrognosis
    Dim objSection As CHCSII_Dental_NotesPlus.clsISection
    Dim enumProg As ePrognosis
    Dim objTooth As ToothBL
    Dim intToothId As Integer
    Dim strExcellent As String
    Dim strGood As String
    Dim strFair As String
    Dim strQuestionable As String
    Dim strPoor As String
    Dim strHopeless As String
    Dim strPrognosis As String
    Dim strLeft As String
    Dim strRight As String
    Dim strNoteToothName As String
    Dim intStart As Integer
    Dim intFound As Integer
    Dim blnToothFound As Boolean
    Dim blnSaveProg As Boolean

    Set objPrognosis = New clsPrognosis
    Set objSection = objPrognosis
    objSection.LoadSection objExam.DentalParent.sPtnID, objExam.DentalEncounterID
      With objPrognosis
         strExcellent = .ExcellentTeeth
         strGood = .GoodTeeth
         strFair = .FairTeeth
         strQuestionable = .QuestionableTeeth
         strPoor = .PoorTeeth
         strHopeless = .HopelessTeeth
      End With

    For Each objTooth In objExam.Mouth
        enumProg = objTooth.Prognosis
        Select Case enumProg
            Case perUnknown
                blnSaveProg = False
            Case perExcellent
                strPrognosis = strExcellent
                blnSaveProg = True
            Case perGood
                strPrognosis = strGood
                blnSaveProg = True
            Case perFair
                strPrognosis = strFair
                blnSaveProg = True
            Case perQuestionable
                strPrognosis = strQuestionable
                blnSaveProg = True
            Case perPoor
                strPrognosis = strPoor
                blnSaveProg = True
            Case perHopeless
                strPrognosis = strHopeless
                blnSaveProg = True
        End Select

        If blnSaveProg Then
            intStart = 1
            'put an end delimiter to find the last value
            If Len(strPrognosis) <> 0 Then
                strPrognosis = strPrognosis & "|"
            End If
            intFound = InStr(intStart, strPrognosis, "|", vbTextCompare)

            blnToothFound = False
            Do While intFound <> 0
                intToothId = Mid(strPrognosis, intStart, intFound - intStart)

                'if current ToothId is found, nothing to do
                If intToothId = objTooth.ToothId Then
                    blnToothFound = True
                    Exit Do
                End If

                intStart = intFound + 1
                intFound = InStr(intStart, strPrognosis, "|", vbTextCompare)
            Loop

            If Not blnToothFound Then 'Tooth not found, append to end of string
                strPrognosis = strPrognosis & objTooth.ToothId
            
                Select Case enumProg
                    Case perUnknown
                    Case perExcellent
                        strExcellent = strPrognosis
                    Case perGood
                        strGood = strPrognosis
                    Case perFair
                        strFair = strPrognosis
                    Case perQuestionable
                        strQuestionable = strPrognosis
                    Case perPoor
                        strPoor = strPrognosis
                    Case perHopeless
                        strHopeless = strPrognosis
                End Select
            
            End If
        End If

    Next objTooth

   With objPrognosis
      .ExcellentTeeth = strExcellent
      .GoodTeeth = strGood
      .FairTeeth = strFair
      .HopelessTeeth = strHopeless
      .PoorTeeth = strPoor
      .QuestionableTeeth = strQuestionable
   End With
   objSection.SaveSection

End Function

Public Function LoadNotesPrognosis(objExam As ExamBL) As Boolean
    Dim objPrognosis As clsPrognosis
    Dim objSection As CHCSII_Dental_NotesPlus.clsISection
    Dim strPrognosis As String
    Dim intStart As Integer
    Dim intFound As Integer
    Dim intToothId As Integer
    Dim ePrognosis As ePrognosis
    Dim objTooth As ToothBL
    Dim intProgCount As Integer
    
    'Initialize the prognosis
    For Each objTooth In objExam.Mouth
        objTooth.Prognosis = perUnknown
    Next objTooth
    
    Set objPrognosis = New clsPrognosis
    Set objSection = objPrognosis
    objSection.LoadSection objExam.PatientID, objExam.DentalEncounterID
    
    For intProgCount = 1 To 6
        'cycle from best to worse so end up showing the worst prognosis
        'Prognosis is at pipe delimited string of ToothId's 1-52)
        Select Case intProgCount
            Case 1
               strPrognosis = objPrognosis.ExcellentTeeth
               ePrognosis = perExcellent
            Case 2
               strPrognosis = objPrognosis.GoodTeeth
               ePrognosis = perGood
            Case 3
               strPrognosis = objPrognosis.FairTeeth
               ePrognosis = perFair
            Case 4
               strPrognosis = objPrognosis.QuestionableTeeth
               ePrognosis = perQuestionable
            Case 5
               strPrognosis = objPrognosis.PoorTeeth
               ePrognosis = perPoor
            Case 6
               strPrognosis = objPrognosis.HopelessTeeth
               ePrognosis = perHopeless
        End Select
        
        intStart = 1
        'put an end delimiter to find the last value
        If Len(strPrognosis) <> 0 Then
            strPrognosis = strPrognosis & "|"
        End If
        intFound = InStr(intStart, strPrognosis, "|", vbTextCompare)
        Do While intFound <> 0
            intToothId = Mid(strPrognosis, intStart, intFound - intStart)
            Select Case intToothId
                Case 1 To 32
                    objExam.Mouth.Tooth(intToothId).Prognosis = ePrognosis
                Case 33 To 42
                    objExam.Mouth.Tooth(intToothId - 29).Prognosis = ePrognosis
                Case 43 To 52
                    objExam.Mouth.Tooth(intToothId - 23).Prognosis = ePrognosis
            End Select
            
            intStart = intFound + 1
            intFound = InStr(intStart, strPrognosis, "|", vbTextCompare)
        Loop
    Next intProgCount
    
End Function

Public Function LoadOralState(objExam As ExamBL) As Boolean

Dim objSVGSingleton As srvSVGComm.SingletonFactory
Dim objSVGComm As srvSVGComm.clsSVGComm
Dim objGroups As CHCSII_Dental_DXTX.colGroups
'Dim objCharts As CHCSII_Dental_DXTX.clsCharts
Dim objNotes As CHCSII_Dental_DXTX.clsNotes
Dim strTeethPath As String
Dim objGenNotes As clsNotes

    'if exam is part of current encounter then use SVGComm teeth path
    If LoadOralFromServers(objExam) Then
        
        'Teeth Path Now comes from the SVG Manage library
        strTeethPath = GetTeethPath

        objExam.TeethPath = strTeethPath
        
        'Clear any previously loaded pontic/unerupted/supernumerary info
        objExam.ResetPonticInfo
        objExam.ResetSupernumeraryInfo
        objExam.ResetUneruptedInfo
        objExam.ResetImplantInfo
        
        'Get pontic/unerupted/supernumerary info from Dental DataLayer
        LoadToothDxTx objExam
        objExam.InitialLoad = False
    'else use default teeth path (coming from either new default, or from stored db value)
    End If
    
    LoadToothPathData objExam.TeethPath, objExam
    LoadNotesPrognosis objExam
End Function

Private Function GetTeethPath() As String

Dim objSVGMng As CHCSII_Dental_SVG_Manager.clsManage
Dim objSVGSingleton As CHCSII_Dental_SVG_Manager.SF
Dim oGCTooth As clsGCTooth
Dim strPath As String

    Set objSVGSingleton = New CHCSII_Dental_SVG_Manager.SF
    Set objSVGMng = objSVGSingleton.Singleton

    strPath = ""
    
    For Each oGCTooth In objSVGMng.mcolTeeth
        If oGCTooth.DateTimeWholeToothRemoved = CDate(0) And objSVGMng.Positions.GetPositionByToothNum(oGCTooth.ToothId).ToothErupted = True Then
            If oGCTooth.ToothVisible Then
                strPath = strPath & oGCTooth.ToothId & ","
            End If
        End If
    Next oGCTooth
    
    strPath = Trim(strPath) 'get rid of trailing spaces
    If Len(strPath) <> 0 Then 'get rid of trailing commas
        strPath = Left(strPath, Len(strPath) - 1)
    End If
    
    Set objSVGSingleton = Nothing
    Set objSVGMng = Nothing
    
    GetTeethPath = strPath
    
End Function


Public Function LoadOralFromServers(objExam As ExamBL) As Boolean
Dim blnRet As Boolean

    'If objExam.DentalEncounterID = objExam.DentalParent.psDentalEncounterID Then
    If objExam.InitialLoad Then
        blnRet = True
    Else
        blnRet = False
    End If
    LoadOralFromServers = blnRet
End Function

Private Function LoadToothDxTx(objExam As ExamBL)

Dim objSVGMng As CHCSII_Dental_SVG_Manager.clsManage
Dim objSVGSingleton As CHCSII_Dental_SVG_Manager.SF
Dim oGCTooth As clsGCTooth
Dim dtRecentStatus As Date
Dim strToothStatus As String
Dim dtStatus As Date

    Set objSVGSingleton = New CHCSII_Dental_SVG_Manager.SF
    Set objSVGMng = objSVGSingleton.Singleton

    For Each oGCTooth In objSVGMng.mcolTeeth
        strToothStatus = ""
        dtRecentStatus = CDate(0)
        
        'Having an implant overrides everything else, so check for that first
        dtStatus = oGCTooth.DateTimeImplant
        If dtStatus > dtRecentStatus Then
            dtRecentStatus = dtStatus
            strToothStatus = "Implant"
        Else
            'Do everything else by timestamp
            dtStatus = oGCTooth.DateTimeCrownRemoved
            If dtStatus > dtRecentStatus Then
                dtRecentStatus = dtStatus
                strToothStatus = "CrownRemoved"
            End If
            
            dtStatus = oGCTooth.DateTimeErupted
            If dtStatus > dtRecentStatus Then
                dtRecentStatus = dtStatus
                strToothStatus = "Erupted"
            End If
            
            dtStatus = oGCTooth.DateTimeRootRemoved
            If dtStatus > dtRecentStatus Then
                dtRecentStatus = dtStatus
                strToothStatus = "RootRemoved"
            End If
            
            dtStatus = oGCTooth.DateTimeSupernumary
            If dtStatus > dtRecentStatus Then
                dtRecentStatus = dtStatus
                strToothStatus = "Supernumary"
            End If
    
            dtStatus = oGCTooth.DateTimeUnErupted
            If dtStatus > dtRecentStatus Then
                dtRecentStatus = dtStatus
                strToothStatus = "UnErupted"
            End If
    
            'Whole Tooth removed already checked in GetToothPath
            'dtStatus = oGCTooth.dDateTimeWholeToothRemoved
            'If dtStatus > dtRecentStatus Then
            '    dtRecentStatus = dtStatus
            '    strToothStatus = "WholeToothRemoved"
            'End If
        
        End If
        
    
        
        Select Case strToothStatus
            Case "CrownRemoved"
                '???
            Case "Erupted"
                '???
            Case "Implant"
                objExam.Mouth.ToothById(oGCTooth.ToothId).IsImplant = True
            Case "RootRemoved"
                objExam.Mouth.ToothById(oGCTooth.ToothId).IsPontic = True
            Case "Supernumary"
                objExam.Mouth.ToothById(oGCTooth.ToothId).IsSupernumerary = True
            Case "UnErupted"
                objExam.Mouth.ToothById(oGCTooth.ToothId).IsUnerupted = True
            'Case "WholeToothRemoved"
            '    objExam.Mouth.ToothByName(oGCTooth.ToothId).IsMissing = True
        End Select
    Next oGCTooth
    
    Set objSVGSingleton = Nothing
    Set objSVGMng = Nothing
    
    
End Function




Private Function LoadToothPathData(strTeethPath As String, objExam As ExamBL)
Dim blnPresentTeeth(1 To 52) As Long
Dim intStart As Integer
Dim intFound As Integer
Dim intToothName As Integer
Dim objTeeth As MouthBL
Dim objTooth As ToothBL

    'Init all teeth to missing
    Set objTeeth = objExam.Mouth
    For Each objTooth In objTeeth
      'Implants are chartable even though they don't appear in tooth path
      If Not objTooth.IsImplant Then
        objTooth.IsMissing = True
      End If
    Next objTooth
    
    'append a comma to the end so the last tooth is found
    If Len(strTeethPath) <> 0 Then
        strTeethPath = strTeethPath & ","
    End If
    
    intStart = 1
    intFound = InStr(intStart, strTeethPath, ",", vbTextCompare)
    Do While intFound <> 0
        intToothName = Mid(strTeethPath, intStart, intFound - intStart)
        Select Case intToothName
            Case 1 To 32
                objExam.Mouth.Tooth(intToothName).IsMissing = False
                objExam.Mouth.Tooth(intToothName).IsPrimary = False
            Case 33 To 42
                objExam.Mouth.Tooth(intToothName - 29).IsMissing = False
                objExam.Mouth.Tooth(intToothName - 29).IsPrimary = True
            Case 43 To 52
                objExam.Mouth.Tooth(intToothName - 23).IsMissing = False
                objExam.Mouth.Tooth(intToothName - 23).IsPrimary = True
        End Select
        
        intStart = intFound + 1
        intFound = InStr(intStart, strTeethPath, ",", vbTextCompare)
    Loop
    
    Set objTooth = Nothing
    Set objTeeth = Nothing
    
End Function


Public Function GetPerioExams(lngPatientUnitID As Long) As ExamsBL
    'Dim objDas As ICHCSII_DAS
    Dim objExamsRS As Recordset
    Dim objMeasRS As Recordset
    Dim objExams As ExamsBL
    Dim objExam As ExamBL
    Dim fldEval As ADODB.Field
    Dim fldEncID As ADODB.Field
    Dim fldEncDt As ADODB.Field
    
    'Dim objRec As Record
    
    'Dim strSQL As String
    
    'strSQL = GetPatientExamsSQL(lngPatientUnitID)
    
    'Set objDas = GetDasObj
    Set objExams = New ExamsBL
    
    'Set objExamsRS = objDas.OpenRecordset(strSQL)
    Set objExamsRS = GetExamRS(lngPatientUnitID)
    Set objMeasRS = GetMeasRS(lngPatientUnitID)

    With objExamsRS
      If Not .BOF And Not .EOF Then
          .MoveFirst
      End If
      
      'Save the recordsets for on-demand loading by the Exams object
      Set objExams.ExamsRS = objExamsRS
      Set objExams.MeasurementsRS = objMeasRS
      
      If Not .EOF Then
         Set fldEval = .Fields("Evaluation_Type_ID")
         Set fldEncID = .Fields("Dental_Encounter_ID")
         Set fldEncDt = .Fields("Encounter_Date")
      End If
      
      Do While Not objExamsRS.EOF
          Set objExam = New ExamBL
          'Load the exam data
          With objExam
             .EvalType = fldEval.Value
             .DentalEncounterID = fldEncID.Value
             
             'Using the stored procs is returning GMT, so have to convert
             If UseStoredProcs Then
               .ExamDate = mobjComm.GmtToLocal(fldEncDt.Value)
             Else
               .ExamDate = fldEncDt.Value
             End If
          End With
          
          objExams.Add objExam
          .MoveNext
      Loop
    End With
    
    Set objExamsRS = Nothing
    'Set objDas = Nothing
    
    Set GetPerioExams = objExams
    Set objExams = Nothing
    
End Function

Friend Function LoadPerioPaths(objExam As ExamBL, objExamsRS As Recordset)
Dim fldTeethPath As ADODB.Field
Dim fldPontic As ADODB.Field
Dim fldUnerupted As ADODB.Field
Dim fldSuper As ADODB.Field
Dim fldImplant As ADODB.Field

    If Not (objExamsRS.BOF And objExamsRS.EOF) Then
        objExamsRS.Filter = "Dental_Encounter_ID = " & objExam.DentalEncounterID
    End If
    If Not objExamsRS.BOF And Not objExamsRS.EOF Then
        objExamsRS.MoveFirst
    End If
    
    If Not objExamsRS.EOF Then
      Set fldTeethPath = objExamsRS.Fields("Teeth_Path")
      Set fldPontic = objExamsRS.Fields("Pontic")
      Set fldUnerupted = objExamsRS.Fields("Unerupted")
      Set fldSuper = objExamsRS.Fields("Supernumerary")
      Set fldImplant = objExamsRS.Fields("Implant")
    End If
    
    If Not objExamsRS.EOF Then
        'Get TeethPath
        If Not IsNull(fldTeethPath.Value) And Len(Trim(fldTeethPath.Value)) <> 0 Then
            objExam.TeethPath = fldTeethPath.Value
        End If
    
        'Load DB Pontic, Unerupted, and Supernumerary info
        If Not IsNull(fldPontic.Value) And Len(Trim(fldPontic.Value)) <> 0 Then
            objExam.PonticPath = fldPontic.Value
        End If
    
        If Not IsNull(fldUnerupted.Value) And Len(Trim(fldUnerupted.Value)) <> 0 Then
            objExam.UneruptedPath = fldUnerupted.Value
        End If
    
        If Not IsNull(fldSuper.Value) And Len(Trim(fldSuper.Value)) <> 0 Then
            objExam.SupernumeraryPath = fldSuper.Value
        End If
        
        If Not IsNull(fldImplant.Value) And Len(Trim(fldImplant.Value)) <> 0 Then
            objExam.ImplantPath = fldImplant.Value
        End If
        
    End If
    objExamsRS.Filter = ""
End Function

Friend Function LoadPerioMeasurements(objExam As ExamBL, objRS As Recordset)
Dim lngToothID As Integer
Dim objTooth As ToothBL
Dim strProb() As String
Dim strGM() As String
Dim strCA(6) As String
Dim i As Integer

Dim fldTooth As ADODB.Field
Dim fldProb As ADODB.Field
Dim fldGM As ADODB.Field
Dim fldMGJ As ADODB.Field
Dim fldFurc As ADODB.Field
Dim fldBone As ADODB.Field
Dim fldBleed As ADODB.Field
Dim fldSupp As ADODB.Field
Dim fldPlaq As ADODB.Field
Dim fldMob As ADODB.Field
   

   'For some reason having the with block first doesn't want to compile????
   
   If Not (objRS.BOF And objRS.EOF) Then
      objRS.Filter = "Dental_Encounter_ID = " & objExam.DentalEncounterID
   
   With objRS
      If Not .BOF And Not .EOF Then
          .MoveFirst
      End If

      If Not .EOF Then
         Set fldTooth = .Fields("Tooth_ID")
         Set fldProb = .Fields("ProbDepth")
         Set fldGM = .Fields("GM")
         Set fldMGJ = .Fields("MGJ")
         Set fldFurc = .Fields("Furcation")
         Set fldBone = .Fields("BoneLevel")
         Set fldBleed = .Fields("Bleeding")
         Set fldSupp = .Fields("Suppuration")
         Set fldPlaq = .Fields("Plaque")
         Set fldMob = .Fields("Mobility")
      End If
   End With
    
    Do While Not objRS.EOF
        lngToothID = fldTooth.Value
        Set objTooth = objExam.Mouth.Tooth(lngToothID)
        
        LoadSixValueDBString fldProb.Value, objTooth, perProbDepth
        LoadSixValueDBString fldGM.Value, objTooth, perGM
        
        strProb = Split(fldProb.Value, ",")
        strGM = Split(fldGM.Value, ",")
        For i = 0 To 5
            If Len(Trim(strProb(i))) = 0 Then
                strProb(i) = " "
            End If
            If Len(Trim(strGM(i))) = 0 Then
                strGM(i) = " "
            End If
            
            If strProb(i) = " " And strGM(i) = " " Then
                strCA(i) = " "
            ElseIf strProb(i) = " " Then
                If strGM(i) = "0" Then
                    strCA(i) = " "
                Else
                    strCA(i) = strGM(i)
                End If
            ElseIf strGM(i) = " " Then
                strCA(i) = strProb(i)
            Else
                strCA(i) = CStr(CInt(strProb(i)) + CInt(strGM(i)))
            End If
        Next i
        'LoadSixValueDBString objRS.Fields("CA"), objTooth, perCA
        LoadSixValueDBString Join(strCA, ","), objTooth, perCA
        
        
        LoadSixValueDBString fldMGJ.Value, objTooth, perMGJ
        LoadFurcationDbString fldFurc.Value, objTooth
        'LoadSixValueDBString objRS.Fields("Furcation"), objtooth, perFurcation
        'LoadSixValueDBString objRS.Fields("BoneLevel"), objtooth, perBoneLoss
        LoadBoneLossDbString fldBone.Value, objTooth
        LoadSixValueDBString fldBleed.Value, objTooth, perBleeding
        LoadSixValueDBString fldSupp.Value, objTooth, perSuppuration
        LoadSixValueDBString fldPlaq.Value, objTooth, perPlaque
        
        LoadOneValueDBString fldMob.Value, objTooth, perMobility
        'LoadOneValueDBString objRS.Fields("Prognosis"), objTooth, perPrognosis
        
        objRS.MoveNext
    Loop
    
    'clear the filter
    objRS.Filter = ""
    End If
End Function
Private Function LoadFurcationDbString(strDbString As String, objTooth As ToothBL)
Dim j As Long
Dim eSide As eToothSideType
Dim eLoc As eMeasurementLocation
Dim eMeasType As eMeasurementType
Dim strValues(1 To 6) As String
Dim strTemp As String
Dim intCommaLoc As Integer
Dim k As Long
Dim strValue

    strValue = strDbString
    
    'Parse the string into individual values
    For k = 1 To 3
        intCommaLoc = InStr(strValue, ",")
        If intCommaLoc <> 0 Then
            strValues(k) = Left(strValue, intCommaLoc - 1)
            strValue = Right(strValue, Len(strValue) - intCommaLoc)
        Else
            strValues(k) = strValue
        End If
    Next k
    
    
    eSide = perSideLingual
    eMeasType = perFurcation
    For j = 1 To 3
        Select Case j
            Case 1
                eLoc = perLocOne
            Case 2
                eLoc = perLocTwo
            Case 3
                eLoc = perLocThree
        End Select
        
        strTemp = strValues(j)
        If Len(Trim(strTemp)) = 0 Then
            strTemp = perUnknownValue
        End If
        objTooth.Value(eMeasType, eSide, eLoc) = strTemp
    Next j
    
    LoadFurcationDbString = True
End Function
Private Function LoadBoneLossDbString(strDbString As String, objTooth As ToothBL)
Dim j As Long
Dim eSide As eToothSideType
Dim eLoc As eMeasurementLocation
Dim eMeasType As eMeasurementType
Dim strValues(1 To 6) As String
Dim strTemp As String
Dim intCommaLoc As Integer
Dim k As Long
Dim strValue As String

    strValue = strDbString
    
    'Parse the string into individual values
    For k = 1 To 2
        intCommaLoc = InStr(strValue, ",")
        If intCommaLoc <> 0 Then
            strValues(k) = Left(strValue, intCommaLoc - 1)
            strValue = Right(strValue, Len(strValue) - intCommaLoc)
        Else
            strValues(k) = strValue
        End If
    Next k
    
    
    eSide = perSideLingual
    eMeasType = perBoneLoss
    For j = 1 To 2
        Select Case j
            Case 1
                eLoc = perLocOne
            Case 2
                eLoc = perLocTwo
        End Select
        
        strTemp = strValues(j)
        If Len(Trim(strTemp)) = 0 Then
            strTemp = perUnknownValue
        End If
        objTooth.Value(eMeasType, eSide, eLoc) = strTemp
    Next j
    
    LoadBoneLossDbString = True
End Function

Private Function LoadMGJDbString(strDbString As String, objTooth As ToothBL)
'Dim j As Long
'Dim eSide As eToothSideType
'Dim eLoc As eMeasurementLocation
'Dim eMeasType As eMeasurementType
'Dim strValues(1 To 6) As String
'Dim strTemp As String
'Dim intCommaLoc As Integer
'Dim k As Long
'Dim strValue As String
'
'    strValue = strDbString
'
'    'Parse the string into individual values
'    For k = 1 To 2
'        intCommaLoc = InStr(strValue, ",")
'        If intCommaLoc <> 0 Then
'            strValues(k) = Left(strValue, intCommaLoc - 1)
'            strValue = Right(strValue, Len(strValue) - intCommaLoc)
'        Else
'            strValues(k) = strValue
'        End If
'    Next k
'
'    eLoc = perLocOne
'    eMeasType = perMGJ
'    For j = 1 To 2
'        Select Case j
'            Case 1
'                eSide = perSideBuccal
'            Case 2
'                eSide = perSideLingual
'        End Select
'
'        strTemp = strValues(j)
'        If Len(Trim(strTemp)) = 0 Then
'            strTemp = perUnknownValue
'        End If
'        objtooth.Value(eMeasType, eSide, eLoc) = strTemp
'    Next j
'
'    LoadMGJDbString = True
End Function


'Private Function LoadPerioMeasurements(objExam As ExamBL, objRS As Recordset)
'Dim lngToothID As Long
'Dim eLocationType As eMeasurementLocation
'Dim eMeasurementType As eMeasurementType
'Dim strMeasurement As String
'Dim eToothSide As eGridType
'
'    objRS.Filter = "Dental_Encounter_ID = " & objExam.DentalEncounterID
'    If Not objRS.BOF And Not objRS.EOF Then
'        objRS.MoveFirst
'    End If
'
'    Do While Not objRS.EOF
'        lngToothID = objRS.Fields("Tooth_ID")
'        eLocationType = objRS.Fields("Location_ID")
'        eMeasurementType = objRS.Fields("Measurement_Type_ID")
'        strMeasurement = objRS.Fields("Measurement")
'        eToothSide = objRS.Fields("Tooth_Side_ID")
'        objExam.Mouth.Tooth(CInt(lngToothID)).Value(eMeasurementType, eToothSide, eLocationType) = strMeasurement
'        objRS.MoveNext
'    Loop
'
'End Function

Private Function GetExamRS(lngPatientUnitID As Long) As Recordset
Dim strSQL As String
Dim objRS As Recordset
Dim objDas As ICHCSII_DAS
'Dim sngStart As Single
'Dim sngEnd As Single

  ' sngStart = Timer
    Set objDas = GetDasObj
    If goCHCSIIConn.UseSQLServer Or Not UseStoredProcs Then
      strSQL = GetPatientExamsSQL(lngPatientUnitID)
      'strSQL = "SELECT pe.* FROM Perio_Measurements pe, Dental_Encounter de " & _
               "WHERE de.Patient_Unit_ID = " & lngPatientUnitID & _
               " AND pe.Dental_Encounter_ID = de.Dental_Encounter_ID "
      Set objRS = objDas.OpenRecordset(strSQL)
    Else
      Dim objCmd As CHCSII_Command.CHCSII_Cmd
'      Dim objSPDas As ICHCSII_DAS
'
'      Set objSPDas = New CHCSII_DASsP
      Set objCmd = New CHCSII_Cmd
      With objCmd
         .SPName = "Perio_Exams_Load"
            
         .AddParam False, False, "textString", lngPatientUnitID
         .AddParam True, True, "", ""
      End With
      
      
      Set objRS = objDas.OpenRecordsetSP(objCmd)


      Set objCmd = Nothing
      'Set objSPDas = Nothing
    End If
    Set GetExamRS = objRS
    Set objRS = Nothing
   '   sngEnd = Timer
   '   Debug.Print "GetExamRS: " & (sngEnd - sngStart) & " sec"
End Function

Private Function GetMeasRS(lngPatientUnitID As Long) As Recordset
Dim strSQL As String
Dim objRS As Recordset
Dim objDas As ICHCSII_DAS
'Dim sngStart As Single
'Dim sngEnd As Single

 '     sngStart = Timer
    Set objDas = GetDasObj
    If TypeOf objDas Is GEMS_DAS Or Not UseStoredProcs Then
      strSQL = "SELECT pe.* FROM Perio_Measurements pe, Dental_Encounter de " & _
               "WHERE de.Patient_Unit_ID = " & lngPatientUnitID & _
               " AND pe.Dental_Encounter_ID = de.Dental_Encounter_ID "
      Set objRS = objDas.OpenRecordset(strSQL)
    Else
      Dim objCmd As CHCSII_Command.CHCSII_Cmd
      'Dim objSPDas As ICHCSII_DAS
      
      'Set objSPDas = New CHCSII_DASsP
      Set objCmd = New CHCSII_Cmd
      With objCmd
         .SPName = "Perio_Measurements_Load"
            
         .AddParam False, False, "textString", lngPatientUnitID
         .AddParam True, True, "", ""
      End With

      Set objRS = objDas.OpenRecordsetSP(objCmd)
      Set objCmd = Nothing
      'Set objSPDas = Nothing
    End If
    Set GetMeasRS = objRS
    Set objRS = Nothing
  '    sngEnd = Timer
  '    Debug.Print "GetMeasRS: " & (sngEnd - sngStart) & " sec"
End Function

Private Function GetPatientExamsSQL(lngPatientID As Long) As String
Dim strSQL As String

    'Version to get Teeth Path
    strSQL = "SELECT pe.Evaluation_Type_ID, pe.Teeth_Path, pe.Dental_Encounter_ID, de.Encounter_Date, pe.Pontic, pe.Unerupted, pe.Supernumerary, pe.Implant FROM Perio_Exams pe, Dental_Encounter de " & _
             "WHERE de.Patient_Unit_ID = " & lngPatientID & _
             " AND pe.Dental_Encounter_ID = de.Dental_Encounter_ID" & _
             " ORDER BY de.Encounter_Date DESC"
             

    GetPatientExamsSQL = strSQL
End Function

Private Function GetDentalEncounterDate(lngDentEncID As Long) As Date
    Dim objDas As ICHCSII_DAS
    Dim objRS As Recordset
    Dim strSQL As String
    Dim dtDentEnc As Date
    
    strSQL = GetDentalEncounterDateSQL(lngDentEncID)
    
    Set objDas = GetDasObj
    
    Set objRS = objDas.OpenRecordset(strSQL)
    
    If Not objRS.BOF And Not objRS.EOF Then
        objRS.MoveFirst
    End If
    
    dtDentEnc = DateValue(objRS.Fields("Encounter_Date"))
    
    Set objRS = Nothing
    Set objDas = Nothing
    
    GetDentalEncounterDate = dtDentEnc
End Function

Private Function GetDentalEncounterDateSQL(lngDentEncID As Long) As String
Dim strSQL As String

    strSQL = "SELECT Encounter_Date FROM Dental_Encounter " & _
             " WHERE Dental_Encounter_ID = " & lngDentEncID
    
    GetDentalEncounterDateSQL = strSQL
    
End Function

'Add public InsertNewExam(EncID, EvalType)
Public Function InsertNewExam(lngDentalEncID As Long, enumEvalType As eEvalType, strTeethPath As String) As ExamBL

    Dim objDas As ICHCSII_DAS
    Dim objExam As ExamBL
    Dim strSQL As String
    
    Set objDas = GetDasObj
    
    If TypeOf objDas Is GEMS_DAS Or Not UseStoredProcs Then
       strSQL = InsertExamSQL(lngDentalEncID, enumEvalType, strTeethPath)
   
       objDas.ExecuteSQL strSQL
    Else
      Dim objCmd As CHCSII_Command.CHCSII_Cmd
      'Dim objSPDas As ICHCSII_DAS
      
      'Set objSPDas = New CHCSII_DASsP
      Set objCmd = New CHCSII_Cmd
      objCmd.SPName = "Perio_Insert_Exam"
         
      objCmd.AddParam False, False, "textString", lngDentalEncID
      objCmd.AddParam False, False, "textString", enumEvalType
      objCmd.AddParam False, False, "textString", strTeethPath
      
      objDas.ExecuteSP objCmd
      Set objCmd = Nothing
      'Set objSPDas = Nothing
   End If
    
    
    Set objExam = New ExamBL
    objExam.DentalEncounterID = lngDentalEncID
    objExam.EvalType = enumEvalType
    objExam.ExamDate = GetDentalEncounterDate(lngDentalEncID)
    
    Set InsertNewExam = objExam
    
    Set objExam = Nothing
    Set objDas = Nothing
End Function

'Add private InsertExamSQL(EncID, EvalType)
Private Function InsertExamSQL(lngDetnalEncID As Long, enumEvalType As eEvalType, strTeethPath As String)
    Dim strSQL As String
    
    strSQL = "INSERT INTO Perio_Exams(Dental_Encounter_ID, Evaluation_Type_ID, Teeth_Path) " & _
             "VALUES(" & lngDetnalEncID & ", " & enumEvalType & ", '" & strTeethPath & "')"
    
    InsertExamSQL = strSQL
End Function

Public Function SaveExam(objExam As ExamBL) As Boolean
Dim objDas As ICHCSII_DAS
Dim strSQL As String
    
    strSQL = SaveExamSQL(objExam)
    
    Set objDas = GetDasObj
    objDas.ExecuteSQL strSQL
    
    SaveMeasurements objExam
    SaveNotesPrognosis objExam
    SaveNarrativeSummary objExam
    
    SaveExam = True
End Function

Public Function SaveMeasurements(objExam As ExamBL) As Boolean
Dim strSQL As String
Dim objTooth As ToothBL
'Dim objDas As ICHCSII_DAS
Dim lngDentEncID As Long
Dim objSQL As ISqlOps
Dim colSQL As Collection

    'Set objDas = GetDasObj
    Set objSQL = GetSQLOpsObj
    Set colSQL = New Collection
    
    lngDentEncID = objExam.DentalEncounterID
    
    For Each objTooth In objExam.Mouth
        strSQL = UpdateTeethMeasurementSQL(lngDentEncID, objTooth)
        'objDas.ExecuteSQL strSQL
        colSQL.Add strSQL
    Next objTooth
    objSQL.ExecuteMulti colSQL
End Function
'Private Function SaveMeasurements(objExam As ExamBL) As Boolean
'    Dim lngToothID As Long
'    Dim objTooth As ToothBL
'    Dim lngDentEncID As Long
'
'    lngDentEncID = objExam.DentalEncounterID
'
'    For Each objTooth In objExam.Mouth
'        '6 value per tooth measurements
'        UpdateProbDepths lngDentEncID, objTooth
'        UpdateCA lngDentEncID, objTooth
'        UpdateGM lngDentEncID, objTooth
'        UpdateBleeding lngDentEncID, objTooth
'        UpdatePlaque lngDentEncID, objTooth
'        UpdateSuppuration lngDentEncID, objTooth
'        UpdateMGJ lngDentEncID, objTooth
'        UpdateFurcation lngDentEncID, objTooth
'        UpdateBoneLoss lngDentEncID, objTooth
'
'        'the rest of the measurements
'        UpdateMobility lngDentEncID, objTooth
'        UpdatePrognosis lngDentEncID, objTooth
'
'    Next objTooth
'
'    SaveMeasurements = True
'
'    Set objTooth = Nothing
'
'End Function

Private Function SaveExamSQL(objExam As ExamBL) As String
Dim strSQL As String
    
    strSQL = "UPDATE Perio_Exams SET Evaluation_Type_ID = " & objExam.EvalType & _
             ", Teeth_Path = '" & objExam.TeethPath & "'" & _
             ", Pontic = '" & objExam.PonticPath & "'" & _
             ", Unerupted = '" & objExam.UneruptedPath & "'" & _
             ", Supernumerary = '" & objExam.SupernumeraryPath & "'" & _
             ", Implant = '" & objExam.ImplantPath & "'" & _
             " WHERE Dental_Encounter_ID = " & objExam.DentalEncounterID
    
    SaveExamSQL = strSQL
End Function

Public Function InsertNewTeethMeasurements(lngDentEncID As Long) As Collection
Dim objTeeth As Collection
Dim strSQL As String
Dim lngToothID As Long
Dim objTooth As ToothBL
Dim colSQL As Collection
Dim objSQL As ISqlOps

    Set objTeeth = New Collection

    Set objSQL = GetSQLOpsObj
    Set colSQL = New Collection

    For lngToothID = 1 To 32
        Set objTooth = New ToothBL
        objTooth.ToothNumber = lngToothID
        objTeeth.Add objTooth

        strSQL = InsertNewTeethMeasurementsSQL(lngDentEncID, objTooth)
        colSQL.Add strSQL

    Next lngToothID

    objSQL.ExecuteMulti colSQL
    Set InsertNewTeethMeasurements = objTeeth
End Function


'Public Function InsertNewTeethMeasurements(lngDentEncID As Long) As Collection
'    Dim objTeeth  As Collection
'    Dim strSQL As String
'    Dim lngToothID As Long
'    Dim objtooth As ToothBL
'
'    Set objTeeth = New Collection
'
'    For lngToothID = 1 To 32
'        Set objtooth = New ToothBL
'        objtooth.ToothNumber = lngToothID
'
'        InsertNewProbDepths lngDentEncID, objtooth
'        InsertNewCA lngDentEncID, objtooth
'        InsertNewGM lngDentEncID, objtooth
'        InsertNewBleeding lngDentEncID, objtooth
'        InsertNewPlaque lngDentEncID, objtooth
'        InsertNewSuppuration lngDentEncID, objtooth
'        InsertNewMobility lngDentEncID, objtooth
'        InsertNewPrognosis lngDentEncID, objtooth
'        InsertNewMGJ lngDentEncID, objtooth
'        InsertNewFurcation lngDentEncID, objtooth
'        InsertNewBoneLoss lngDentEncID, objtooth
'
'        objTeeth.Add objtooth
'    Next lngToothID
'
'    Set InsertNewTeethMeasurements = objTeeth
'
'    Set objTeeth = Nothing
'End Function

'Private Function InsertNewOneValueMeasurement(lngDentEncID As Long, lngMeasurementType As eMeasurementType, objTooth As ToothBL)
'Dim objDas As ICHCSII_DAS
'Dim strSQL As String
'Dim strValue As String
'Dim eSide As eToothSideType
'Dim eLoc As eMeasurementLocation
'
'    Set objDas = GetDasObj
'
'    strValue = perUnknownValue
'    eSide = perSideWholeTooth
'    eLoc = perLocWholeTooth
'
'    strSQL = InsertNewTeethMeasurementsSQL(lngDentEncID, objTooth.ToothNumber, eSide, eLoc, lngMeasurementType, strValue)
'    objTooth.Value(lngMeasurementType, eSide, eLoc) = strValue
'    objDas.ExecuteSQL strSQL
'
'End Function

'Private Function InsertNewSixValueMeasurements(lngDentEncID As Long, lngMeasurementType As eMeasurementType, objTooth As ToothBL)
'Dim objDas As ICHCSII_DAS
'Dim strSQL As String
'Dim strValue As String
'Dim eSide As eGridType
'Dim eLoc As eMeasurementLocation
'Dim i As Integer
'Dim j As Integer
'
'    Set objDas = GetDasObj
'
'    'strValue = perunknownvalue
'
'    'for each side
'    For i = 1 To 2
'        Select Case i
'            Case 1
'                eSide = perBuccal
'            Case 2
'                eSide = perLingual
'        End Select
'
'        'for each location per side
'        For j = 1 To 3
'            Select Case j
'                Case 1
'                    eLoc = perLocOne
'                Case 2
'                    eLoc = perLocTwo
'                Case 3
'                    eLoc = perLocThree
'            End Select
'
'            strValue = objTooth.Value(lngMeasurementType, eSide, eLoc)
'            If Len(Trim(strValue)) = 0 Then
'                strValue = perUnknownValue
'            End If
'
'            strSQL = InsertNewTeethMeasurementsSQL(lngDentEncID, objTooth.ToothNumber, eSide, eLoc, lngMeasurementType, strValue)
'            objTooth.Value(lngMeasurementType, eSide, eLoc) = strValue
'            objDas.ExecuteSQL strSQL
'
'        Next j
'    Next i
'
'End Function

'Private Function InsertNewProbDepths(lngDentEncID As Long, objTooth As ToothBL)
'    InsertNewSixValueMeasurements lngDentEncID, perProbDepth, objTooth
'End Function
'Private Function InsertNewCA(lngDentEncID As Long, objTooth As ToothBL)
'    InsertNewSixValueMeasurements lngDentEncID, perCA, objTooth
'End Function
'Private Function InsertNewGM(lngDentEncID As Long, objTooth As ToothBL)
'    InsertNewSixValueMeasurements lngDentEncID, perGM, objTooth
'End Function
'Private Function InsertNewBleeding(lngDentEncID As Long, objTooth As ToothBL)
'    InsertNewSixValueMeasurements lngDentEncID, perBleeding, objTooth
'End Function
'Private Function InsertNewPlaque(lngDentEncID As Long, objTooth As ToothBL)
'    InsertNewSixValueMeasurements lngDentEncID, perPlaque, objTooth
'End Function
'Private Function InsertNewSuppuration(lngDentEncID As Long, objTooth As ToothBL)
'    InsertNewSixValueMeasurements lngDentEncID, perSuppuration, objTooth
'End Function
'Private Function InsertNewMGJ(lngDentEncID As Long, objTooth As ToothBL)
'    InsertNewSixValueMeasurements lngDentEncID, perMGJ, objTooth
'End Function
'Private Function InsertNewFurcation(lngDentEncID As Long, objTooth As ToothBL)
'    InsertNewSixValueMeasurements lngDentEncID, perFurcation, objTooth
'End Function
'Private Function InsertNewBoneLoss(lngDentEncID As Long, objTooth As ToothBL)
'    InsertNewSixValueMeasurements lngDentEncID, perBoneLoss, objTooth
'End Function
'
'
'Private Function InsertNewMobility(lngDentEncID As Long, objTooth As ToothBL)
'    InsertNewOneValueMeasurement lngDentEncID, perMobility, objTooth
'End Function
'Private Function InsertNewPrognosis(lngDentEncID As Long, objTooth As ToothBL)
'    InsertNewOneValueMeasurement lngDentEncID, perPrognosis, objTooth
'End Function


'Private Function UpdateProbDepths(lngDentEncID As Long, objTooth As ToothBL)
'    UpdateSixValueMeasurements lngDentEncID, perProbDepth, objTooth
'End Function
'Private Function UpdateCA(lngDentEncID As Long, objTooth As ToothBL)
'    UpdateSixValueMeasurements lngDentEncID, perCA, objTooth
'End Function
'Private Function UpdateGM(lngDentEncID As Long, objTooth As ToothBL)
'    UpdateSixValueMeasurements lngDentEncID, perGM, objTooth
'End Function
'Private Function UpdateBleeding(lngDentEncID As Long, objTooth As ToothBL)
'    UpdateSixValueMeasurements lngDentEncID, perBleeding, objTooth
'End Function
'Private Function UpdatePlaque(lngDentEncID As Long, objTooth As ToothBL)
'    UpdateSixValueMeasurements lngDentEncID, perPlaque, objTooth
'End Function
'Private Function UpdateSuppuration(lngDentEncID As Long, objTooth As ToothBL)
'    UpdateSixValueMeasurements lngDentEncID, perSuppuration, objTooth
'End Function
'Private Function UpdateMGJ(lngDentEncID As Long, objTooth As ToothBL)
'    UpdateSixValueMeasurements lngDentEncID, perMGJ, objTooth
'End Function
'Private Function UpdateFurcation(lngDentEncID As Long, objTooth As ToothBL)
'    UpdateSixValueMeasurements lngDentEncID, perFurcation, objTooth
'End Function
'Private Function UpdateBoneLoss(lngDentEncID As Long, objTooth As ToothBL)
'    UpdateSixValueMeasurements lngDentEncID, perBoneLoss, objTooth
'End Function
'
'Private Function UpdateMobility(lngDentEncID As Long, objTooth As ToothBL)
'    UpdateOneValueMeasurement lngDentEncID, perMobility, objTooth
'End Function
'Private Function UpdatePrognosis(lngDentEncID As Long, objTooth As ToothBL)
'    UpdateOneValueMeasurement lngDentEncID, perPrognosis, objTooth
'End Function
'
'Private Function UpdateSixValueMeasurements(lngDentEncID As Long, lngMeasurementType As eMeasurementType, objTooth As ToothBL) As Boolean
'Dim objDas As ICHCSII_DAS
'Dim strSQL  As String
'Dim strValue As String
'Dim eSide As eGridType
'Dim eLoc As eMeasurementLocation
'Dim i As Integer
'Dim j As Integer
'
'    Set objDas = GetDasObj
'
'    For i = 1 To 2
'        Select Case i
'            Case 1
'                eSide = perBuccal
'            Case 2
'                eSide = perLingual
'        End Select
'
'        For j = 1 To 3
'            Select Case j
'                Case 1
'                    eLoc = perLocOne
'                Case 2
'                    eLoc = perLocTwo
'                Case 3
'                    eLoc = perLocThree
'            End Select
'
'            strValue = objTooth.Value(lngMeasurementType, eSide, eLoc)
'            If Len(Trim(strValue)) = 0 Then
'                strValue = perUnknownValue
'            End If
'            strSQL = UpdateTeethMeasurementSQL(lngDentEncID, objTooth.ToothNumber, eSide, eLoc, lngMeasurementType, strValue)
'            objDas.ExecuteSQL strSQL
'
'        Next j
'    Next i
'
'    UpdateSixValueMeasurements = True
'End Function

'Private Function UpdateOneValueMeasurement(lngDentEncID As Long, lngMeasurementType As eMeasurementType, objTooth As ToothBL) As Boolean
'Dim objDas As ICHCSII_DAS
'Dim strSQL  As String
'Dim strValue As String
'Dim eSide As eToothSideType
'Dim eLoc As eMeasurementLocation
'
'    Set objDas = GetDasObj
'
'    eSide = perSideWholeTooth
'    eLoc = perLocWholeTooth
'
'    strValue = objTooth.Value(lngMeasurementType, eSide, eLoc)
'    If Len(Trim(strValue)) = 0 Then
'        strValue = perUnknownValue
'    End If
'    strSQL = UpdateTeethMeasurementSQL(lngDentEncID, objTooth.ToothNumber, eSide, eLoc, lngMeasurementType, strValue)
'    objDas.ExecuteSQL strSQL
'
'    UpdateOneValueMeasurement = True
'End Function

Private Function InsertNewTeethMeasurementsSQL(lngDentEncID As Long, objTooth As ToothBL) As String
Dim strSQL As String
Dim strProbDepth As String
Dim strGM As String
Dim strCA As String
Dim strMGJ As String
Dim strMobility As String
Dim strFurcation As String
Dim strPrognosis As String
Dim strBoneLevel As String
Dim strBleeding As String
Dim strSuppuration As String
Dim strPlaque As String

    'Get the values to Insert
    strProbDepth = GetProbDepthDBString(objTooth)
    strGM = GetGmDbString(objTooth)
    strCA = GetCaDbSTring(objTooth)
    strMGJ = GetMgjDbString(objTooth)
    strFurcation = GetFurcationDbString(objTooth)
    strBoneLevel = GetBoneLossDbString(objTooth)
    strBleeding = GetBleedingDbString(objTooth)
    strSuppuration = GetSuppurationDbString(objTooth)
    strPlaque = GetPlaqueDbString(objTooth)
    strMobility = GetMobilityDbString(objTooth)
    strPrognosis = GetPrognosisDbString(objTooth)
        
    strSQL = "INSERT INTO Perio_Measurements(Dental_Encounter_ID, Tooth_ID, ProbDepth, CA, " & _
            "GM, MGJ, Mobility, Furcation, Prognosis, BoneLevel, Bleeding, Suppuration, Plaque) " & _
            "VALUES(" & lngDentEncID & ", " & objTooth.ToothNumber & ", " & _
            "'" & strProbDepth & "', " & _
            "'" & strCA & "', " & _
            "'" & strGM & "', " & _
            "'" & strMGJ & "', " & _
            "'" & strMobility & "', " & _
            "'" & strFurcation & "', " & _
            "'" & strPrognosis & "', " & _
            "'" & strBoneLevel & "', " & _
            "'" & strBleeding & "', " & _
            "'" & strSuppuration & "', " & _
            "'" & strPlaque & "' " & _
            ")"
        
    InsertNewTeethMeasurementsSQL = strSQL
End Function

Private Function GetOneValueDBString(objTooth As ToothBL, eMeasType As eMeasurementType) As String
Dim strValue As String
Dim eSide As eToothSideType
Dim eLoc As eMeasurementLocation

    eSide = perSideWholeTooth
    eLoc = perLocWholeTooth
    strValue = objTooth.Value(eMeasType, eSide, eLoc)

    GetOneValueDBString = strValue

End Function
Private Function GetProbDepthDBString(objTooth As ToothBL) As String
    GetProbDepthDBString = GetSixValueDBString(objTooth, perProbDepth)
End Function
Private Function GetCaDbSTring(objTooth As ToothBL) As String
    GetCaDbSTring = GetSixValueDBString(objTooth, perProbDepth)
End Function
Private Function GetGmDbString(objTooth As ToothBL) As String
    GetGmDbString = GetSixValueDBString(objTooth, perGM)
End Function
Private Function GetMgjDbString(objTooth As ToothBL) As String
Dim strTemp As String
Dim strValues() As String

    strTemp = GetSixValueDBString(objTooth, perMGJ)
    strValues = GetDbStringArray(strTemp)
    
    GetMgjDbString = strValues(1) & "," & strValues(4)
End Function
Private Function GetDbStringArray(strDbString As String) As String()
Dim strTemp As String
Dim intEnd As Integer
Dim intStart As Integer
Dim strValues(1 To 6) As String
Dim i As Long

    strTemp = strDbString
    
    intStart = 1
    intEnd = 0
    For i = 1 To 6
        intEnd = InStr(strTemp, ",")
        If intEnd <> 0 Then
            strValues(i) = Mid(strTemp, intStart, intEnd - intStart)
            strTemp = Right(strTemp, Len(strTemp) - intEnd)
        Else 'no delimiter found, at the end
            strValues(i) = strTemp
            Exit For
        End If
    Next i
    
    GetDbStringArray = strValues
End Function
Private Function GetMobilityDbString(objTooth As ToothBL) As String
    GetMobilityDbString = GetOneValueDBString(objTooth, perMobility)
End Function
Private Function GetFurcationDbString(objTooth As ToothBL) As String
Dim strTemp As String
Dim intLoc As Integer
Dim strValues() As String

    strTemp = GetSixValueDBString(objTooth, perFurcation)
    
    strValues = GetDbStringArray(strTemp)
    
    GetFurcationDbString = strValues(4) & "," & strValues(5) & "," & strValues(6)
'    'Only use the first 3 values
'    intLoc = InStr(strTemp, ",")
'    intLoc = InStr(intLoc + 1, strTemp, ",")
'    intLoc = InStr(intLoc + 1, strTemp, ",")
'
'    strTemp = Left(strTemp, intLoc - 1)
'
'    GetFurcationDbString = strTemp
End Function
Private Function GetPrognosisDbString(objTooth As ToothBL) As String
Dim strTemp As String
Dim eProg As ePrognosis

    eProg = GetOneValueDBString(objTooth, perPrognosis)
    Select Case eProg
        Case perUnknown
            strTemp = ""
        Case perExcellent
            strTemp = "Excellent"
        Case perGood
            strTemp = "Good"
        Case perFair
            strTemp = "Fair"
        Case perQuestionable
            strTemp = "Questionable"
        Case perPoor
            strTemp = "Poor"
        Case perHopeless
            strTemp = "Hopeless"
    End Select
    GetPrognosisDbString = strTemp
End Function
Private Function GetBoneLossDbString(objTooth As ToothBL) As String
Dim strTemp As String
Dim intLoc As Integer
Dim strValues() As String

    strTemp = GetSixValueDBString(objTooth, perBoneLoss)
    strValues = GetDbStringArray(strTemp)
    
    GetBoneLossDbString = strValues(4) & "," & strValues(5)
'    'Only use the first 2 values
'    intLoc = InStr(strTemp, ",")
'    intLoc = InStr(intLoc + 1, strTemp, ",")
'
'    strTemp = Left(strTemp, intLoc - 1)
'
'    GetBoneLossDbString = strTemp

End Function
Private Function GetBleedingDbString(objTooth As ToothBL) As String
    GetBleedingDbString = GetSixValueDBString(objTooth, perBleeding)
End Function
Private Function GetSuppurationDbString(objTooth As ToothBL) As String
    GetSuppurationDbString = GetSixValueDBString(objTooth, perSuppuration)
End Function
Private Function GetPlaqueDbString(objTooth As ToothBL) As String
    GetPlaqueDbString = GetSixValueDBString(objTooth, perPlaque)
End Function


Private Function GetSixValueDBString(objTooth As ToothBL, eMeasType As eMeasurementType) As String
Dim strValue As String
Dim strTemp As String
Dim i As Long
Dim j As Long
Dim eSide As eToothSideType
Dim eLoc As eMeasurementLocation

    For i = 1 To 2
        Select Case i
            Case 1
                eSide = perSideBuccal
            Case 2
                eSide = perSideLingual
        End Select
        
        For j = 1 To 3
            Select Case j
                Case 1
                    eLoc = perLocOne
                Case 2
                    eLoc = perLocTwo
                Case 3
                    eLoc = perLocThree
            End Select
            
            strTemp = UCase(objTooth.Value(eMeasType, eSide, eLoc))
            If strTemp = "FALSE" Or strTemp = "TRUE" Then
                strTemp = Left(strTemp, 1)
            End If
            strValue = strValue & strTemp & ","
        Next j
    Next i

    'strip off the last comma
    strValue = Left(strValue, Len(strValue) - 1)
    GetSixValueDBString = strValue
End Function
Private Function LoadOneValueDBString(ByVal strValue As String, objTooth As ToothBL, eMeasType As eMeasurementType) As Boolean
Dim blnRet As Boolean
Dim lngValue As Long

    blnRet = True
    
    Select Case eMeasType
        
        Case perMobility
            objTooth.Mobility = strValue
        
        Case perPrognosis
            Select Case Trim(strValue)
                Case ""
                    lngValue = perUnknownValue
                Case "Excellent"
                    lngValue = perExcellent
                Case "Good"
                    lngValue = perGood
                Case "Fair"
                    lngValue = perFair
                Case "Questionable"
                    lngValue = perQuestionable
                Case "Poor"
                    lngValue = perPoor
                Case "Hopeless"
                    lngValue = perHopeless
            End Select
            objTooth.Prognosis = lngValue
        
        Case Else
            blnRet = False
    End Select
    
    LoadOneValueDBString = blnRet

End Function
Private Function LoadSixValueDBString(ByVal strValue As String, objTooth As ToothBL, eMeasType As eMeasurementType) As Boolean
Dim i As Long
Dim j As Long
Dim eSide As eToothSideType
Dim eLoc As eMeasurementLocation
Dim strValues(1 To 6) As String
Dim strTemp As String
Dim intCommaLoc As Integer
Dim k As Long
    
    'Parse the string into individual values
    For k = 1 To 6
        intCommaLoc = InStr(strValue, ",")
        If intCommaLoc <> 0 Then
            strValues(k) = Left(strValue, intCommaLoc - 1)
            strValue = Right(strValue, Len(strValue) - intCommaLoc)
        Else
            strValues(k) = strValue
        End If
        
        Select Case UCase(strValues(k))
            Case "T"
                strValues(k) = "TRUE"
            Case "F"
                strValues(k) = "FALSE"
            Case Else
        End Select
    Next k
    
    
    For i = 0 To 1
        Select Case i
            Case 0
                eSide = perSideBuccal
            Case 1
                eSide = perSideLingual
        End Select
        
        For j = 1 To 3
            Select Case j
                Case 1
                    eLoc = perLocOne
                Case 2
                    eLoc = perLocTwo
                Case 3
                    eLoc = perLocThree
            End Select
            
            strTemp = strValues(i * 3 + j)
            If Len(Trim(strTemp)) = 0 Then
                strTemp = perUnknownValue
            End If
            objTooth.Value(eMeasType, eSide, eLoc) = strTemp
        Next j
    Next i
    LoadSixValueDBString = True
End Function

'Private Function InsertNewTeethMeasurementsSQL(lngDentEncID As Long, lngToothID As Long, lngSideID As Long, lngLocationID As Long, lngMeasurementTypeID As eMeasurementType, strMeasurement As String) As String
'Dim strSQL As String
'    strSQL = "INSERT INTO Perio_Measurements(Dental_Encounter_ID, Tooth_ID, " & _
'             "Tooth_Side_ID, Location_ID, Measurement_Type_ID, Measurement) " & _
'             "VALUES(" & lngDentEncID & ", " & lngToothID & ", " & lngSideID & ", " & lngLocationID & _
'             ", " & lngMeasurementTypeID & ", '" & strMeasurement & " ')"
'    InsertNewTeethMeasurementsSQL = strSQL
'End Function
'Private Function UpdateTeethMeasurementSQL(lngDentEncID As Long, lngToothID As Long, lngSideID As Long, lngLocationID As Long, lngMeasurementTypeID As eMeasurementType, strMeasurement As String) As String
'Dim strSQL As String
'
'    strSQL = "UPDATE Perio_Measurements Set Measurement = '" & strMeasurement & "' " & _
'             " WHERE Dental_Encounter_ID = " & lngDentEncID & " AND Tooth_ID = " & lngToothID & _
'             " AND Tooth_Side_ID = " & lngSideID & " AND Location_ID = " & lngLocationID & _
'             " AND Measurement_Type_ID = " & lngMeasurementTypeID
'    UpdateTeethMeasurementSQL = strSQL
'End Function
Private Function UpdateTeethMeasurementSQL(lngDentEncID As Long, objTooth As ToothBL) As String
Dim strSQL As String
Dim strProbDepth As String
Dim strGM As String
Dim strCA As String
Dim strMGJ As String
Dim strMobility As String
Dim strFurcation As String
Dim strPrognosis As String
Dim strBoneLevel As String
Dim strBleeding As String
Dim strSuppuration As String
Dim strPlaque As String
Dim lngToothID As Long

    'Get the values to Insert
    lngToothID = objTooth.ToothNumber
    
    strProbDepth = GetProbDepthDBString(objTooth)
    strGM = GetGmDbString(objTooth)
    strCA = GetCaDbSTring(objTooth)
    
    strMGJ = GetMgjDbString(objTooth)
    strFurcation = GetFurcationDbString(objTooth)
    strBoneLevel = GetBoneLossDbString(objTooth)
    strBleeding = GetBleedingDbString(objTooth)
    strSuppuration = GetSuppurationDbString(objTooth)
    strPlaque = GetPlaqueDbString(objTooth)
    
    strMobility = GetMobilityDbString(objTooth)
    strPrognosis = GetPrognosisDbString(objTooth)
        
    strSQL = "UPDATE Perio_Measurements SET " & _
             "ProbDepth = '" & strProbDepth & "', " & _
             "CA = '" & strCA & "', " & _
             "GM = '" & strGM & "', " & _
             "MGJ = '" & strMGJ & "', " & _
             "Mobility = '" & strMobility & "', " & _
             "Furcation = '" & strFurcation & "', " & _
             "Prognosis = '" & strPrognosis & "', " & _
             "BoneLevel = '" & strBoneLevel & "', " & _
             "Bleeding = '" & strBleeding & "', " & _
             "Suppuration = '" & strSuppuration & "', " & _
             "Plaque = '" & strPlaque & "' " & _
             "WHERE Dental_Encounter_ID = " & lngDentEncID & " " & _
             "AND Tooth_ID = " & lngToothID
        
    UpdateTeethMeasurementSQL = strSQL
End Function


Public Property Get UseStoredProcs() As Boolean
   UseStoredProcs = mblnUseStoredProcs
End Property
Public Property Let UseStoredProcs(blnUse As Boolean)
   mblnUseStoredProcs = blnUse
End Property
Public Property Get CommObject() As Comm
   Set CommObject = mobjComm
End Property
Public Property Set CommObject(objComm As Comm)
   Set mobjComm = objComm
End Property

