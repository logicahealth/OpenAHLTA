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
Attribute VB_Name = "ExamBL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Enum examPath
    examPontic
    examUnerupted
    examSupernumerary
    examImplant
End Enum

Public EvalType As eEvalType
Public ExamDate As Date
Public Comments As String
Public PatientID As Long
Public DentalEncounterID As Long
Public Mouth As MouthBL
Private mstrTeethPath As String
Private mstrPonticPath As String
Private mstrUneruptedPath As String
Private mstrSupernumeraryPath As String
Private mstrImplantPath As String

Private mobjExamsParent As ExamsBL
Private mobjDentalParent As CHCSII_Dental.Dental
Private mblnInitialLoad As Boolean
Private mblnMeasurementsLoaded As Boolean
Private mblnIsChartable As Boolean  'at least one chartable tooth exists

Friend Property Get MeasurementsLoaded() As Boolean
    MeasurementsLoaded = mblnMeasurementsLoaded
End Property
Friend Property Let MeasurementsLoaded(blnLoaded As Boolean)
    mblnMeasurementsLoaded = blnLoaded
End Property

Friend Property Get InitialLoad() As Boolean
    InitialLoad = mblnInitialLoad
End Property
Friend Property Let InitialLoad(blnInitialLoad As Boolean)
    mblnInitialLoad = blnInitialLoad
End Property

Public Property Get DentalParent() As CHCSII_Dental.Dental
    Set DentalParent = mobjDentalParent
End Property
Public Property Set DentalParent(objDentalParent As CHCSII_Dental.Dental)
    Set mobjDentalParent = objDentalParent
End Property


Public Property Get ExamsParent() As ExamsBL
    Set ExamsParent = mobjExamsParent
End Property
Public Property Set ExamsParent(objExams As ExamsBL)
    Set mobjExamsParent = objExams
End Property

Private Sub Class_Initialize()
Dim i As Integer
Dim strTeethPath As String

    Set Mouth = New MouthBL
    
    'init teeth path to all 32 secondary teeth
    For i = 1 To 31
        strTeethPath = strTeethPath & i & ","
    Next i
    strTeethPath = strTeethPath & "32"
    
    Me.TeethPath = strTeethPath
    Me.InitialLoad = False
    
End Sub

Public Property Get TeethPath() As String
    TeethPath = mstrTeethPath
End Property
Public Property Let TeethPath(strTeethPath As String)
    mstrTeethPath = strTeethPath
End Property

Friend Property Get ImplantPath() As String
Dim objTooth As ToothBL
Dim strPath As String

    For Each objTooth In Mouth
        If objTooth.IsImplant Then
            strPath = strPath & "," & objTooth.ToothName
        End If
    Next objTooth
    
    'strip off the leading comma
    If Len(strPath) <> 0 Then
        strPath = Right(strPath, Len(strPath) - 1)
    End If
    ImplantPath = strPath
End Property
Friend Property Let ImplantPath(strImplant As String)
    mstrImplantPath = strImplant
    ParseImplant
End Property


Friend Property Get PonticPath() As String
Dim objTooth As ToothBL
Dim strPath As String

    For Each objTooth In Mouth
        If objTooth.IsPontic Then
            strPath = strPath & "," & objTooth.ToothName
        End If
    Next objTooth
    
    'strip off the leading comma
    If Len(strPath) <> 0 Then
        strPath = Right(strPath, Len(strPath) - 1)
    End If
    PonticPath = strPath
End Property
Friend Property Let PonticPath(strPontic As String)
    mstrPonticPath = strPontic
    ParsePontic
End Property

Friend Property Get UneruptedPath() As String
Dim objTooth As ToothBL
Dim strPath As String

    For Each objTooth In Mouth
        If objTooth.IsUnerupted Then
            strPath = strPath & "," & objTooth.ToothName
        End If
    Next objTooth
    
    'strip off the leading comma
    If Len(strPath) <> 0 Then
        strPath = Right(strPath, Len(strPath) - 1)
    End If
    UneruptedPath = strPath
End Property
Friend Property Let UneruptedPath(strUnerupted As String)
    mstrUneruptedPath = strUnerupted
    ParseUnerupted
End Property
Friend Property Get SupernumeraryPath() As String
Dim objTooth As ToothBL
Dim strPath As String

    For Each objTooth In Mouth
        If objTooth.IsSupernumerary Then
            strPath = strPath & "," & objTooth.ToothName
        End If
    Next objTooth
    
    'strip off the leading comma
    If Len(strPath) <> 0 Then
        strPath = Right(strPath, Len(strPath) - 1)
    End If
    
    SupernumeraryPath = strPath
End Property
Friend Property Let SupernumeraryPath(strSupernumerary As String)
    mstrSupernumeraryPath = strSupernumerary
    ParseSupernumerary
End Property

Private Sub Class_Terminate()
    Set mobjDentalParent = Nothing
    Set mobjExamsParent = Nothing
End Sub

Private Sub ParseImplant()
    ParseToothProperty mstrImplantPath, examImplant
End Sub
Private Sub ParsePontic()
    ParseToothProperty mstrPonticPath, examPontic
End Sub
Private Sub ParseUnerupted()
    ParseToothProperty mstrUneruptedPath, examUnerupted
End Sub
Private Sub ParseSupernumerary()
    ParseToothProperty mstrSupernumeraryPath, examSupernumerary
End Sub

Private Sub ParseToothProperty(strTeeth As String, ePath As examPath)
Dim strTeethList() As String
Dim i As Integer
Dim objTooth As ToothBL

    Select Case ePath
        Case examPath.examPontic
            ResetPonticInfo
        Case examPath.examSupernumerary
            ResetSupernumeraryInfo
        Case examPath.examUnerupted
            ResetUneruptedInfo
        Case examPath.examImplant
            ResetImplantInfo
    End Select

    strTeethList = Split(strTeeth, ",")
    For i = LBound(strTeethList) To UBound(strTeethList)
        Set objTooth = Mouth.ToothByName(strTeethList(i))
            
        Select Case ePath
            Case examPath.examPontic
                objTooth.IsPontic = True
            Case examPath.examSupernumerary
                objTooth.IsSupernumerary = True
            Case examPath.examUnerupted
                objTooth.IsUnerupted = True
            Case examPath.examImplant
                objTooth.IsImplant = True
        End Select
    Next i
    
End Sub

Public Sub ResetImplantInfo()
Dim objTooth As ToothBL

    For Each objTooth In Mouth
        objTooth.IsImplant = False
    Next objTooth
End Sub
Public Sub ResetPonticInfo()
Dim objTooth As ToothBL

    For Each objTooth In Mouth
        objTooth.IsPontic = False
    Next objTooth
End Sub

Public Sub ResetUneruptedInfo()
Dim objTooth As ToothBL
    
    For Each objTooth In Mouth
        objTooth.IsUnerupted = False
    Next objTooth
End Sub

Public Sub ResetSupernumeraryInfo()
Dim objTooth As ToothBL

    For Each objTooth In Mouth
        objTooth.IsSupernumerary = False
    Next objTooth
End Sub

Friend Sub AssessChartability()
'Sets the ExamBL.IsChartable property that states if any teeth are chartable
'i.e. IsChartable is False only if all teeth are not chartable
Dim objTooth As ToothBL
    
    On Error GoTo ErrHandler
    
    mblnIsChartable = False
    For Each objTooth In Mouth
        With objTooth
            'if pontic, missing or unerupted, can't chart on it
            If .IsChartable Then
                'a chartable tooth exists
                mblnIsChartable = True
                Exit For
            End If
        End With
    Next objTooth
    
    Exit Sub
    
ErrHandler:
    Err.Raise Err.Number, "ExamBL.AssessChartability", Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Public Property Get IsChartable() As Boolean
    IsChartable = mblnIsChartable
End Property

Friend Sub Dispose()
    Set mobjExamsParent = Nothing
    Set mobjDentalParent = Nothing
End Sub

Public Function StatisticsNote() As String
Dim strNote As String
Dim strLabel As String

   With Mouth
       .CalculateStatistics
       'strNote = crtfStartNote & crtfStartTemplateName & "Perio Charting" & crtfEndTemplateName
       strNote = crtfStartTemplateName & "Perio Charting" & crtfEndTemplateName
       strNote = strNote & StatisticSectionRtf("Missing Teeth", .MissingTeethList)
       strNote = strNote & StatisticSectionRtf("Furcation Teeth", .FurcationTeethList)
       strNote = strNote & StatisticSectionRtf("Mobility Teeth", .MobilityTeethList)
       strNote = strNote & StatisticSectionRtf("Bleeding Teeth", .BleedingTeethList)
       
       strNote = strNote & StatisticSectionRtf("Number of Teeth/Sites Examined", .ExaminedTeethCount & " / " & .ExaminedSitesCount)
       strNote = strNote & StatisticSectionRtf("Number of Teeth/Sites Examined that Bled", .BleedingExaminedTeethCount & " / " & .BleedingExaminedSitesCount)
       strNote = strNote & StatisticSectionRtf("Percentage of Sites Examined that Bled", .BleedingSitesPercentage)
       strNote = strNote & StatisticSectionRtf("Number of Teeth/Sites Examined with Plaque", .PlaqueExaminedTeethCount & " / " & .PlaqueExaminedSitesCount)
       strNote = strNote & StatisticSectionRtf("Percentage of Sites Examined with Plaque", .PlaqueSitesPercentage)
       strNote = strNote & StatisticSectionRtf("Number of Teeth in Probing Depth Ranges 1-3/4-5/6+", _
         .ProbDepthTeethRange1_3Count & " / " & .ProbDepthTeethRange4_5Count & " / " & .ProbDepthTeethRange6_20Count)
       strNote = strNote & StatisticSectionRtf("Number of Sites in Probing Depth Ranges 1-3/4-5/6+", _
         .ProbDepthSitesRange1_3Count & " / " & .ProbDepthSitesRange4_5Count & " / " & .ProbDepthSitesRange6_20Count)
                                                                                                     
                                                                                                           
      'Build the Probing Depth Range string
      strLabel = ""
      If .ProbDepthMinValid Then
          strLabel = .ProbDepthMin
      Else
          strLabel = "na"
      End If
      strLabel = strLabel & " / "
      If .ProbDepthMaxValid Then
          strLabel = strLabel & .ProbDepthMax
      Else
          strLabel = strLabel & "na"
      End If
      strLabel = strLabel & " / "
      If .ProbDepthAvgValid Then
          strLabel = strLabel & FormatNumber(.ProbDepthAvg, 2)
      Else
          strLabel = strLabel & "na"
      End If
                                                                                                     
       strNote = strNote & StatisticSectionRtf("Probing Depth Min/Max/Avg", strLabel)
       strNote = strNote & StatisticSectionRtf("Number of Teeth in CA Ranges 1-3/4-5/6+", _
         .CATeethRange1_3Count & " / " & .CATeethRange4_5Count & " / " & .CATeethRange6_20Count)
       strNote = strNote & StatisticSectionRtf("Number of Sites in CA Ranges 1-3/4-5/6+", _
         .CASitesRange1_3Count & " / " & .CASitesRange4_5Count & " / " & .CASitesRange6_20Count)
       
       
      'Build the CA Range
      strLabel = ""
      If .CAMinValid Then
          strLabel = .CAMin
      Else
          strLabel = "na"
      End If
      strLabel = strLabel & " / "
      If .CAMaxValid Then
          strLabel = strLabel & .CAMax
      Else
          strLabel = strLabel & "na"
      End If
      strLabel = strLabel & " / "
      If .CAAvgValid Then
          strLabel = strLabel & FormatNumber(.CAAvg, 2)
      Else
          strLabel = strLabel & "na"
      End If
       strNote = strNote & StatisticSectionRtf("CA Min/Max/Avg", strLabel)
       
   End With
   StatisticsNote = strNote
End Function

Private Function StatisticSectionRtf(strHeader As String, strDetails As String) As String
Dim strNote As String

   If Len(strDetails) <> 0 Then
      'Bold header for each detail
      'strNote = crtfStartTemplateName & strHeader & ":  " & crtfEndTemplateName
      'strNote = strNote & strDetails & crtfEndSection
      
      'header and detail in plain text on same line
      strNote = strHeader & ":  " & strDetails & crtfEndSection
   End If
   StatisticSectionRtf = strNote
End Function
