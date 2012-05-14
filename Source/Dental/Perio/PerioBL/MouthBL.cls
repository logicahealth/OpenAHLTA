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
Attribute VB_Name = "MouthBL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mcolTeeth As Collection
Private mobjPerioDL As PerioDL


'Statistics property variables
Private mstrMissingTeethList  As String
Private mintMissingTeethCount As Integer
Private mstrFurcationTeethList As String
Private mintFurcationTeethCount As Integer
Private mstrMobilityTeethList As String
Private mintMobilityTeethCount As Integer
Private mstrBleedingTeethList As String
Private mintExaminedTeethCount As Integer
Private mintExaminedSitesCount As Integer
Private mintBleedingTeethCount As Integer
Private mintBleedingExaminedTeethCount As Integer
Private mintBleedingSitesCount As Integer
Private mintBleedingExaminedSitesCount As Integer
Private mintPlaqueTeethCount As Integer
Private mintPlaqueExaminedTeethCount As Integer
Private mintPlaqueSitesCount As Integer
Private mintPlaqueExaminedSitesCount As Integer
Private mintProbDepthTeethRangeCount(1 To 3) As Integer
Private mintProbDepthSitesRangeCount(1 To 3) As Integer
Private mintProbDepthMin As Integer
Private mintProbDepthMax As Integer
Private mintProbDepthSum As Integer

Private mintCATeethRangeCount(1 To 3) As Integer
Private mintCASitesRangeCount(1 To 3) As Integer
Private mintCAMin As Integer
Private mintCAMax As Integer
Private mintCASum As Integer
Private mintCACount As Integer

'End Statistics property variables

Friend Property Set PerioDL(objPerioDL As PerioDL)
    Set mobjPerioDL = objPerioDL
End Property

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mcolTeeth.[_NewEnum]
End Function

Private Function LoadTeeth() As Boolean
Dim i As Integer
Dim objTooth As ToothBL

    For i = 1 To 32
        Set objTooth = New ToothBL
        objTooth.ToothNumber = i
        mcolTeeth.Add objTooth, GetKey(i)
    Next i
    LoadTeeth = True
End Function

Public Function Tooth(i As Integer) As ToothBL
    Set Tooth = mcolTeeth(GetKey(i))
End Function

Public Function ToothByName(ByVal strToothName As String) As ToothBL
Dim intToothPos As Integer 'Tooth position 1-32
    
    strToothName = Trim(strToothName)
    If IsNumeric(strToothName) Then
        intToothPos = CInt(strToothName)
    Else
        Select Case Asc(strToothName)
            Case Asc("A") To Asc("J")
                intToothPos = (Asc(strToothName) - Asc("A")) + 4
            Case Asc("K") To Asc("T")
                intToothPos = (Asc(strToothName) - Asc("K")) + 20
        End Select
    End If
            
    Set ToothByName = mcolTeeth(GetKey(intToothPos))
        
End Function

Public Function ToothById(ByVal intToothId As Integer) As ToothBL
Dim intToothPos As Integer 'Tooth position 1-32

    If intToothId >= 1 And intToothId <= 32 Then
        intToothPos = intToothId
    ElseIf intToothId >= 33 And intToothId <= 42 Then
        intToothPos = intToothId - 29
    ElseIf intToothId >= 43 And intToothId <= 52 Then
        intToothPos = intToothId - 23
    End If
        
    Set ToothById = mcolTeeth(GetKey(intToothPos))

End Function

Private Function GetKey(i As Integer) As String
    GetKey = "Tooth" & i
End Function


Private Sub Class_Initialize()
    Set mcolTeeth = New Collection
    LoadTeeth
End Sub

Private Sub Class_Terminate()
    Set mcolTeeth = Nothing
End Sub


Public Function CreateNewMouth(lngDentEncID As Long) As Boolean
Dim objDL As PerioDL
Dim objMouth As MouthBL
Dim colTeeth As Collection
Dim objTooth As ToothBL

    Set objDL = mobjPerioDL
    Set colTeeth = objDL.InsertNewTeethMeasurements(lngDentEncID)
    
    Set mcolTeeth = Nothing
    Set mcolTeeth = New Collection
    
    For Each objTooth In colTeeth
        
        mcolTeeth.Add objTooth, GetKey(objTooth.ToothNumber)
    Next objTooth
    
    CreateNewMouth = True
End Function

Public Sub CalculateStatistics()
Dim objTooth As ToothBL
Dim eSide As eToothSideType
Dim eLoc As eMeasurementLocation
Dim i As Long
Dim j As Long
Dim blnFurcationTooth As Boolean
Dim blnMobilityTooth As Boolean
Dim blnBleedingTooth As Boolean
Dim blnExaminedTooth As Boolean
Dim blnExaminedSite As Boolean
Dim blnPlaqueTooth As Boolean
Dim blnProbDepthToothRange(1 To 3) As Boolean
Dim blnCAToothRange(1 To 3) As Boolean
Dim varProbDepth As Variant
Dim varCA As Variant

    'General inits
    mstrMissingTeethList = ""
    mintMissingTeethCount = 0
    mstrFurcationTeethList = ""
    mintFurcationTeethCount = 0
    mstrMobilityTeethList = ""
    mintMobilityTeethCount = 0
    mstrBleedingTeethList = ""
    mintExaminedTeethCount = 0
    mintExaminedSitesCount = 0
    mintBleedingTeethCount = 0
    mintBleedingExaminedTeethCount = 0
    mintBleedingSitesCount = 0
    mintBleedingExaminedSitesCount = 0
    
    mintPlaqueTeethCount = 0
    mintPlaqueExaminedTeethCount = 0
    mintPlaqueSitesCount = 0
    mintPlaqueExaminedSitesCount = 0
    
    
    mintProbDepthTeethRangeCount(1) = 0
    mintProbDepthTeethRangeCount(2) = 0
    mintProbDepthTeethRangeCount(3) = 0
    mintProbDepthSitesRangeCount(1) = 0
    mintProbDepthSitesRangeCount(2) = 0
    mintProbDepthSitesRangeCount(3) = 0
    mintProbDepthMin = 100
    mintProbDepthMax = -100
    mintProbDepthSum = 0
    
    mintCATeethRangeCount(1) = 0
    mintCATeethRangeCount(2) = 0
    mintCATeethRangeCount(3) = 0
    mintCASitesRangeCount(1) = 0
    mintCASitesRangeCount(2) = 0
    mintCASitesRangeCount(3) = 0
    mintCAMin = 100
    mintCAMax = -100
    mintCASum = 0
    mintCACount = 0
    
    
    For Each objTooth In mcolTeeth
        'Tooth specific inits
        blnExaminedTooth = False
        blnFurcationTooth = False
        blnMobilityTooth = False
        blnBleedingTooth = False
        blnPlaqueTooth = False
        blnProbDepthToothRange(1) = False
        blnProbDepthToothRange(2) = False
        blnProbDepthToothRange(3) = False
        
        blnCAToothRange(1) = False
        blnCAToothRange(2) = False
        blnCAToothRange(3) = False

        'see if tooth is missing
        If objTooth.IsMissing Then
            If Len(objTooth.ToothName) Then
                mstrMissingTeethList = mstrMissingTeethList & " " & objTooth.ToothName
                mintMissingTeethCount = mintMissingTeethCount + 1
            End If
        End If
    
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
            
                'Site specific inits
                blnExaminedSite = False
            
                If objTooth.Furcation(eSide, eLoc) <> " " Then
                    If objTooth.Furcation(eSide, perLocOne) > 0 Then
                        blnFurcationTooth = True
                    End If
                End If
                
                'Probing Depth Statistics
                varProbDepth = objTooth.ProbingDepth(eSide, eLoc)
                If varProbDepth <> " " Then
                    blnExaminedTooth = True
                    blnExaminedSite = True
                    mintExaminedSitesCount = mintExaminedSitesCount + 1
                    
                    Select Case varProbDepth
                        Case 1, 2, 3
                            blnProbDepthToothRange(1) = True
                            mintProbDepthSitesRangeCount(1) = mintProbDepthSitesRangeCount(1) + 1
                        Case 4, 5
                            blnProbDepthToothRange(2) = True
                            mintProbDepthSitesRangeCount(2) = mintProbDepthSitesRangeCount(2) + 1
                        Case Else
                            blnProbDepthToothRange(3) = True
                            mintProbDepthSitesRangeCount(3) = mintProbDepthSitesRangeCount(3) + 1
                    End Select
                    
                    If varProbDepth > mintProbDepthMax Then
                        mintProbDepthMax = varProbDepth
                    End If
                    If varProbDepth < mintProbDepthMin Then
                        mintProbDepthMin = varProbDepth
                    End If
                    
                    mintProbDepthSum = mintProbDepthSum + varProbDepth
                End If
                
                'CA statistics
                varCA = objTooth.CA(eSide, eLoc)
                If varCA <> " " Then
                    Select Case varCA
                        Case 0
                            'do nothing
                        Case 1, 2, 3
                            blnCAToothRange(1) = True
                            mintCASitesRangeCount(1) = mintCASitesRangeCount(1) + 1
                        Case 4, 5
                            blnCAToothRange(2) = True
                            mintCASitesRangeCount(2) = mintCASitesRangeCount(2) + 1
                        Case Else
                            blnCAToothRange(3) = True
                            mintCASitesRangeCount(3) = mintCASitesRangeCount(3) + 1
                    End Select
                    
                    If varCA > mintCAMax Then
                        mintCAMax = varCA
                    End If
                    If varCA < mintCAMin Then
                        mintCAMin = varCA
                    End If
                    
                    mintCASum = mintCASum + varCA
                    mintCACount = mintCACount + 1
                End If
            
                If objTooth.Bleeding(eSide, eLoc) Then
                    blnBleedingTooth = True
                    mintBleedingSitesCount = mintBleedingSitesCount + 1
                    
                    If blnExaminedSite Then
                        mintBleedingExaminedSitesCount = mintBleedingExaminedSitesCount + 1
                    End If

                End If
            
            
                If objTooth.Plaque(eSide, eLoc) Then
                    blnPlaqueTooth = True
                    mintPlaqueSitesCount = mintPlaqueSitesCount + 1
                    
                    If blnExaminedSite Then
                        mintPlaqueExaminedSitesCount = mintPlaqueExaminedSitesCount + 1
                    End If
                End If
            
            Next j
        Next i
        
        If objTooth.Mobility <> " " Then
            If objTooth.Mobility <> perUnknownValue Then
                blnMobilityTooth = True
            End If
        End If

        If blnFurcationTooth Then
            mstrFurcationTeethList = mstrFurcationTeethList & " " & objTooth.ToothName
            mintFurcationTeethCount = mintFurcationTeethCount + 1
        End If
        If blnMobilityTooth Then
            mstrMobilityTeethList = mstrMobilityTeethList & " " & objTooth.ToothName
            mintMobilityTeethCount = mintMobilityTeethCount + 1
        End If
        If blnBleedingTooth Then
            mstrBleedingTeethList = mstrBleedingTeethList & " " & objTooth.ToothName
            mintBleedingTeethCount = mintBleedingTeethCount + 1
            If blnExaminedTooth Then
                mintBleedingExaminedTeethCount = mintBleedingExaminedTeethCount + 1
            End If
        End If
        If blnExaminedTooth Then
            mintExaminedTeethCount = mintExaminedTeethCount + 1
        End If
        If blnPlaqueTooth Then
            mintPlaqueTeethCount = mintPlaqueTeethCount + 1
            If blnExaminedTooth Then
                mintPlaqueExaminedTeethCount = mintPlaqueExaminedTeethCount + 1
            End If
        End If
        If blnProbDepthToothRange(1) Then
            mintProbDepthTeethRangeCount(1) = mintProbDepthTeethRangeCount(1) + 1
        End If
        If blnProbDepthToothRange(2) Then
            mintProbDepthTeethRangeCount(2) = mintProbDepthTeethRangeCount(2) + 1
        End If
        If blnProbDepthToothRange(3) Then
            mintProbDepthTeethRangeCount(3) = mintProbDepthTeethRangeCount(3) + 1
        End If
        
        If blnCAToothRange(1) Then
            mintCATeethRangeCount(1) = mintCATeethRangeCount(1) + 1
        End If
        If blnCAToothRange(2) Then
            mintCATeethRangeCount(2) = mintCATeethRangeCount(2) + 1
        End If
        If blnCAToothRange(3) Then
            mintCATeethRangeCount(3) = mintCATeethRangeCount(3) + 1
        End If
        
        
    Next objTooth

End Sub
Public Property Get MissingTeethList() As String
    MissingTeethList = FormatList(mstrMissingTeethList)
End Property
Public Property Get MissingTeethCount() As Integer
    MissingTeethCount = mintMissingTeethCount
End Property
Public Property Get FurcationTeethList() As String
    FurcationTeethList = FormatList(mstrFurcationTeethList)
End Property
Public Property Get FurcationTeethCount() As Integer
    FurcationTeethCount = mintFurcationTeethCount
End Property
Public Property Get MobilityTeethList() As String
    MobilityTeethList = FormatList(mstrMobilityTeethList)
End Property
Public Property Get MobilityTeethCount() As Integer
    MobilityTeethCount = mintMobilityTeethCount
End Property
Public Property Get BleedingTeethList() As String
    BleedingTeethList = FormatList(mstrBleedingTeethList)
End Property
Public Property Get ExaminedTeethCount() As Integer
    ExaminedTeethCount = mintExaminedTeethCount
End Property
Public Property Get ExaminedSitesCount() As Integer
    ExaminedSitesCount = mintExaminedSitesCount
End Property
Public Property Get BleedingTeethCount() As Integer
    BleedingTeethCount = mintBleedingTeethCount
End Property
Public Property Get BleedingExaminedTeethCount() As Integer
    BleedingExaminedTeethCount = mintBleedingExaminedTeethCount
End Property
Public Property Get BleedingSitescount() As Integer
    BleedingSitescount = mintBleedingSitesCount
End Property
Public Property Get BleedingExaminedSitesCount() As Integer
    BleedingExaminedSitesCount = mintBleedingExaminedSitesCount
End Property
Public Property Get BleedingSitesPercentage() As Single
    If ExaminedSitesCount <> 0 Then
        BleedingSitesPercentage = 100 * (BleedingExaminedSitesCount / ExaminedSitesCount)
    Else
        BleedingSitesPercentage = 0
    End If
End Property

Public Property Get PlaqueTeethCount() As Integer
    PlaqueTeethCount = mintPlaqueTeethCount
End Property
Public Property Get PlaqueExaminedTeethCount() As Integer
    PlaqueExaminedTeethCount = mintPlaqueExaminedTeethCount
End Property
Public Property Get PlaqueSitesCount() As Integer
    PlaqueSitesCount = mintPlaqueSitesCount
End Property
Public Property Get PlaqueExaminedSitesCount() As Integer
    PlaqueExaminedSitesCount = mintPlaqueExaminedSitesCount
End Property
Public Property Get PlaqueSitesPercentage() As Single
    If ExaminedSitesCount <> 0 Then
        PlaqueSitesPercentage = 100 * (PlaqueExaminedSitesCount / ExaminedSitesCount)
    Else
        PlaqueSitesPercentage = 0
    End If
End Property

Public Property Get ProbDepthTeethRange1_3Count() As Integer
    ProbDepthTeethRange1_3Count = mintProbDepthTeethRangeCount(1)
End Property
Public Property Get ProbDepthTeethRange4_5Count() As Integer
    ProbDepthTeethRange4_5Count = mintProbDepthTeethRangeCount(2)
End Property
Public Property Get ProbDepthTeethRange6_20Count() As Integer
    ProbDepthTeethRange6_20Count = mintProbDepthTeethRangeCount(3)
End Property
Public Property Get ProbDepthSitesRange1_3Count() As Integer
    ProbDepthSitesRange1_3Count = mintProbDepthSitesRangeCount(1)
End Property
Public Property Get ProbDepthSitesRange4_5Count() As Integer
    ProbDepthSitesRange4_5Count = mintProbDepthSitesRangeCount(2)
End Property
Public Property Get ProbDepthSitesRange6_20Count() As Integer
    ProbDepthSitesRange6_20Count = mintProbDepthSitesRangeCount(3)
End Property
Public Property Get ProbDepthMin() As Integer
    ProbDepthMin = mintProbDepthMin
End Property
Public Property Get ProbDepthMinValid() As Boolean
    If mintProbDepthMin <> 100 Then
        ProbDepthMinValid = True
    Else
        ProbDepthMinValid = False
    End If
End Property
Public Property Get ProbDepthMax() As Integer
    ProbDepthMax = mintProbDepthMax
End Property
Public Property Get ProbDepthMaxValid() As Boolean
    If mintProbDepthMax <> -100 Then
        ProbDepthMaxValid = True
    Else
        ProbDepthMaxValid = False
    End If
End Property
Public Property Get ProbDepthAvg() As Single
    If mintExaminedSitesCount <> 0 Then
        ProbDepthAvg = mintProbDepthSum / mintExaminedSitesCount
    Else
        ProbDepthAvg = 0
    End If
End Property
Public Property Get ProbDepthAvgValid() As Boolean
    If mintExaminedSitesCount <> 0 Then
        ProbDepthAvgValid = True
    Else
        ProbDepthAvgValid = False
    End If
End Property
Public Property Get CATeethRange1_3Count() As Integer
    CATeethRange1_3Count = mintCATeethRangeCount(1)
End Property
Public Property Get CATeethRange4_5Count() As Integer
    CATeethRange4_5Count = mintCATeethRangeCount(2)
End Property
Public Property Get CATeethRange6_20Count() As Integer
    CATeethRange6_20Count = mintCATeethRangeCount(3)
End Property
Public Property Get CASitesRange1_3Count() As Integer
    CASitesRange1_3Count = mintCASitesRangeCount(1)
End Property
Public Property Get CASitesRange4_5Count() As Integer
    CASitesRange4_5Count = mintCASitesRangeCount(2)
End Property
Public Property Get CASitesRange6_20Count() As Integer
    CASitesRange6_20Count = mintCASitesRangeCount(3)
End Property
Public Property Get CAMin() As Integer
    CAMin = mintCAMin
End Property
Public Property Get CAMinValid() As Boolean
    If mintCAMin <> 100 Then
        CAMinValid = True
    Else
        CAMinValid = False
    End If
End Property
Public Property Get CAMax() As Integer
    CAMax = mintCAMax
End Property
Public Property Get CAMaxValid() As Boolean
    If mintCAMax <> -100 Then
        CAMaxValid = True
    Else
        CAMaxValid = False
    End If
End Property
Public Property Get CAAvg() As Single
    If mintCACount <> 0 Then
        CAAvg = mintCASum / mintCACount
    Else
        CAAvg = 0
    End If
End Property
Public Property Get CAAvgValid() As Boolean
    If mintCACount <> 0 Then
        CAAvgValid = True
    Else
        CAAvgValid = False
    End If
End Property
Private Function FormatList(ByVal strValue As String) As String
    strValue = Trim(strValue)
    strValue = Replace(strValue, " ", ", ")
    FormatList = strValue
End Function
