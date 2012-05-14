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
Attribute VB_Name = "ToothBL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mobjProbDepth As Dictionary
Private mobjGM As Dictionary
Private mobjCA As Dictionary
Private mobjBleeding As Dictionary
Private mobjSuppuration As Dictionary
Private mobjPlaque As Dictionary
Private mobjMGJ As Dictionary
Private mintMobility As Integer
Private mobjFurcation As Dictionary
Private mobjBoneLoss As Dictionary
Private menumPrognosis As ePrognosis
Private mlngToothNumber As Long 'Tooth Position
Private mblnPrimaryTooth As Boolean 'indicates primary or secondary
Private mblnMissingTooth As Boolean 'indicates tooth is missing
Private mblnPonticTooth As Boolean 'indicates tooth is a pontic
Private mblnUneruptedTooth As Boolean 'indicates tooth is unerupted
Private mblnSupernumeraryTooth As Boolean 'indicates tooth is supernumerary
Private mblnImplantTooth As Boolean

Public Property Get ToothId() As Long
    If mblnPrimaryTooth Then
        Select Case mlngToothNumber
            Case 4 To 13
                ToothId = mlngToothNumber + 29
            Case 20 To 29
                ToothId = mlngToothNumber + 23
            Case Else
                ToothId = mlngToothNumber
        End Select
    Else
        ToothId = mlngToothNumber
    End If
End Property
Public Property Get ToothNumber() As Long
    ToothNumber = mlngToothNumber
End Property
Public Property Let ToothNumber(lngToothNumber As Long)
    mlngToothNumber = lngToothNumber
End Property

Public Property Get ToothName() As String
        If mblnPrimaryTooth Then
            Select Case mlngToothNumber
                Case 4 To 13
                    ToothName = Chr(Asc("A") + (mlngToothNumber - 4))
                Case 20 To 29
                    ToothName = Chr(Asc("K") + (mlngToothNumber - 20))
                Case Else
                    ToothName = ""
            End Select
        Else
            ToothName = mlngToothNumber
        End If
End Property

Public Property Get IsSupernumerary() As Boolean
    IsSupernumerary = mblnSupernumeraryTooth
End Property
Public Property Let IsSupernumerary(blnSupernumerary As Boolean)
    mblnSupernumeraryTooth = blnSupernumerary
End Property

Public Property Get IsPrimary() As Boolean
    IsPrimary = mblnPrimaryTooth
End Property
Public Property Let IsPrimary(blnPrimaryTooth As Boolean)
    mblnPrimaryTooth = blnPrimaryTooth
    
    'Set ToothName
    If blnPrimaryTooth Then 'if primary
    Else 'secondary
    End If
End Property

Public Property Get IsMissing() As Boolean
    IsMissing = mblnMissingTooth
End Property
Public Property Let IsMissing(blnMissing As Boolean)
    mblnMissingTooth = blnMissing
End Property
Public Property Get IsPontic() As Boolean
    IsPontic = mblnPonticTooth
End Property
Public Property Let IsPontic(blnPontic As Boolean)
    mblnPonticTooth = blnPontic
End Property
Public Property Get IsUnerupted() As Boolean
    IsUnerupted = mblnUneruptedTooth
End Property
Public Property Let IsUnerupted(blnUnerupted As Boolean)
    mblnUneruptedTooth = blnUnerupted
End Property
Public Property Get IsImplant() As Boolean
    IsImplant = mblnImplantTooth
End Property
Public Property Get IsChartable() As Boolean
    On Error GoTo ErrHandler
    
    IsChartable = Not (IsMissing Or IsPontic Or IsUnerupted)
    
    Exit Property
ErrHandler:
    Err.Raise Err.Number, "ToothBL.IsChartable", Err.Description, Err.HelpFile, Err.HelpContext
End Property
Public Property Let IsImplant(blnImplant As Boolean)
    mblnImplantTooth = blnImplant
End Property

Private Function GetKey(Side As eGridType, Location As eMeasurementLocation) As String
    GetKey = "S" & Side & "L" & Location
End Function

Public Property Get Value(Measurement As eMeasurementType, Optional Side As eGridType, Optional Location As eMeasurementLocation) As Variant
    Select Case Measurement
        Case eMeasurementType.perBleeding
            Value = Bleeding(Side, Location)
        Case eMeasurementType.perBoneLoss
            Value = BoneLoss(Side, Location)
        Case eMeasurementType.perCA
            Value = CA(Side, Location)
        Case eMeasurementType.perFurcation
            Value = Furcation(Side, Location)
        Case eMeasurementType.perGM
            Value = GM(Side, Location)
        Case eMeasurementType.perMGJ
            Value = MGJ(Side, Location)
        Case eMeasurementType.perMobility
            Value = Mobility
        Case eMeasurementType.perPlaque
            Value = Plaque(Side, Location)
        Case eMeasurementType.perProbDepth
            Value = ProbingDepth(Side, Location)
        Case eMeasurementType.perPrognosis
            Value = Prognosis
        Case eMeasurementType.perSuppuration
            Value = Suppuration(Side, Location)
    End Select
End Property

Public Property Let Value(Measurement As eMeasurementType, Side As eGridType, Location As eMeasurementLocation, NewValue As Variant)
    Select Case Measurement
        Case eMeasurementType.perBleeding
            Bleeding(Side, Location) = NewValue
        Case eMeasurementType.perBoneLoss
            BoneLoss(Side, Location) = NewValue
        Case eMeasurementType.perCA
            CA(Side, Location) = NewValue
        Case eMeasurementType.perFurcation
            Furcation(Side, Location) = NewValue
        Case eMeasurementType.perGM
            GM(Side, Location) = NewValue
        Case eMeasurementType.perMGJ
            MGJ(Side, Location) = NewValue
        Case eMeasurementType.perMobility
            Mobility = NewValue
        Case eMeasurementType.perPlaque
            Plaque(Side, Location) = NewValue
        Case eMeasurementType.perProbDepth
            ProbingDepth(Side, Location) = NewValue
        Case eMeasurementType.perPrognosis
            Prognosis = NewValue
        Case eMeasurementType.perSuppuration
            Suppuration(Side, Location) = NewValue
    End Select
End Property


Public Property Get ProbingDepth(Side As eGridType, Location As eMeasurementLocation) As Variant
Dim intvalue As Integer
    intvalue = mobjProbDepth(GetKey(Side, Location))
    If intvalue = perUnknownValue Then
        ProbingDepth = " "
    Else
        ProbingDepth = intvalue
    End If
End Property
Public Property Let ProbingDepth(Side As eGridType, Location As eMeasurementLocation, Value)
    mobjProbDepth(GetKey(Side, Location)) = Value
End Property

Public Property Get GM(Side As eGridType, Location As eMeasurementLocation)
Dim intvalue As Integer
    intvalue = mobjGM(GetKey(Side, Location))
    If intvalue = perUnknownValue Then
        GM = " "
    Else
        GM = intvalue
    End If
End Property
Public Property Let GM(Side As eGridType, Location As eMeasurementLocation, Value)
    mobjGM(GetKey(Side, Location)) = Value
End Property

Public Property Get CA(Side As eGridType, Location As eMeasurementLocation)
Dim intvalue As Integer
    intvalue = mobjCA(GetKey(Side, Location))
    If intvalue = perUnknownValue Then
        CA = " "
    Else
        CA = intvalue
    End If
End Property
Public Property Let CA(Side As eGridType, Location As eMeasurementLocation, Value)
    mobjCA(GetKey(Side, Location)) = Value
End Property

Public Property Get Bleeding(Side As eGridType, Location As eMeasurementLocation) As Boolean
    Bleeding = mobjBleeding(GetKey(Side, Location))
End Property
Public Property Let Bleeding(Side As eGridType, Location As eMeasurementLocation, Value As Boolean)
    mobjBleeding(GetKey(Side, Location)) = Value
End Property

Public Property Get Suppuration(Side As eGridType, Location As eMeasurementLocation) As Boolean
    Suppuration = mobjSuppuration(GetKey(Side, Location))
End Property
Public Property Let Suppuration(Side As eGridType, Location As eMeasurementLocation, Value As Boolean)
    mobjSuppuration(GetKey(Side, Location)) = Value
End Property

Public Property Get Plaque(Side As eGridType, Location As eMeasurementLocation) As Boolean
    Plaque = mobjPlaque(GetKey(Side, Location))
End Property
Public Property Let Plaque(Side As eGridType, Location As eMeasurementLocation, Value As Boolean)
    mobjPlaque(GetKey(Side, Location)) = Value
End Property

Public Property Get MGJ(Side As eGridType, Location As eMeasurementLocation)
Dim intvalue As Integer
    intvalue = mobjMGJ(GetKey(Side, Location))
    If intvalue = perUnknownValue Then
        MGJ = " "
    Else
        MGJ = intvalue
    End If
End Property
Public Property Let MGJ(Side As eGridType, Location As eMeasurementLocation, Value)
    mobjMGJ(GetKey(Side, Location)) = Value
End Property

Public Property Get Furcation(Side As eGridType, Location As eMeasurementLocation)
Dim intvalue As Integer
    intvalue = mobjFurcation(GetKey(Side, Location))
    If intvalue = perUnknownValue Then
        Furcation = " "
    Else
        Furcation = intvalue
    End If
    
End Property
Public Property Let Furcation(Side As eGridType, Location As eMeasurementLocation, Value)
    mobjFurcation(GetKey(Side, Location)) = Value
End Property

Public Property Get BoneLoss(Side As eGridType, Location As eMeasurementLocation)
Dim intvalue As Integer
    intvalue = mobjBoneLoss(GetKey(Side, Location))
    If intvalue = perUnknownValue Then
        BoneLoss = " "
    Else
        BoneLoss = intvalue
    End If
End Property
Public Property Let BoneLoss(Side As eGridType, Location As eMeasurementLocation, ByVal Value)
    If Len(Trim(Value)) = 0 Then 'empty string, set to unknown value
        Value = perUnknownValue
    End If
    mobjBoneLoss(GetKey(Side, Location)) = Value
End Property

Public Property Get Mobility()
    If mintMobility = perUnknownValue Then
        Mobility = " "
    Else
        Mobility = mintMobility
    End If
End Property
Public Property Let Mobility(intvalue)
    If Len(Trim(CStr(intvalue))) = 0 Then
        mintMobility = perUnknownValue
    Else
        mintMobility = intvalue
    End If
End Property

Public Property Get Prognosis() As ePrognosis
    Prognosis = menumPrognosis
End Property
Public Property Let Prognosis(eValue As ePrognosis)
    menumPrognosis = eValue
End Property

Private Sub Class_Initialize()
Dim i As Long
Dim j As Long
Dim eSide As eToothSideType
Dim eLoc As eMeasurementLocation

    Set mobjProbDepth = New Dictionary
    Set mobjGM = New Dictionary
    Set mobjCA = New Dictionary
    Set mobjBleeding = New Dictionary
    Set mobjSuppuration = New Dictionary
    Set mobjPlaque = New Dictionary
    Set mobjMGJ = New Dictionary
    Set mobjFurcation = New Dictionary
    Set mobjBoneLoss = New Dictionary
    
    'init the values
    Prognosis = ePrognosis.perUnknown
    Mobility = perUnknownValue
    
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
            
            ProbingDepth(eSide, eLoc) = perUnknownValue
            GM(eSide, eLoc) = 0
            CA(eSide, eLoc) = perUnknownValue
            Bleeding(eSide, eLoc) = False
            Suppuration(eSide, eLoc) = False
            Plaque(eSide, eLoc) = False
            MGJ(eSide, eLoc) = perUnknownValue
            Furcation(eSide, eLoc) = perUnknownValue
            BoneLoss(eSide, eLoc) = perUnknownValue
            
        Next j
    Next i
    
End Sub

Private Sub Class_Terminate()
    Set mobjProbDepth = Nothing
    Set mobjGM = Nothing
    Set mobjCA = Nothing
    Set mobjBleeding = Nothing
    Set mobjSuppuration = Nothing
    Set mobjPlaque = Nothing
    Set mobjMGJ = Nothing
    Set mobjFurcation = Nothing
    Set mobjBoneLoss = Nothing
End Sub
