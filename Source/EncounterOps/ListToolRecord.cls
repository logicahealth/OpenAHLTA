VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ListToolRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private msSnoID As String
Private msPrefix As String
Private msModifier As String
Private msResult As String
Private msStatus As String
Private msOnset As String
Private msDuration As String
Private msValue As String
Private msUnit As String
Private miRangeScale As Integer
Private mnRangeNormalLow As Double
Private mnRangeNormalHigh As Double
Private mlReferid As Long
Private miFlag As Integer
Private msNote As String
Private msEpisode As String
Private msQualifierLink As String
Private miQualifierLinkSequence As Integer
Private msDataId As Long
Private msPKCText As String

Public Function Clone() As ListToolRecord
  Dim newObject As New ListToolRecord
  With newObject
    .Duration = Me.Duration
    .Episode = Me.Episode
    .Flag = Me.Flag
    .SnoID = Me.SnoID
    .Modifier = Me.Modifier
    .Note = Me.Note
    .Onset = Me.Onset
    .prefix = Me.prefix
    .QualifierLink = Me.QualifierLink
    .QualifierLinkSequence = Me.QualifierLinkSequence
    .RangeNormalHigh = Me.RangeNormalHigh
    .RangeNormalLow = Me.RangeNormalLow
    .RangeScale = Me.RangeScale
    .Referid = Me.Referid
    .result = Me.result
    .Status = Me.Status
    .Unit = Me.Unit
    .value = Me.value
    
  End With
  Set Clone = newObject
  Set newObject = Nothing
End Function

Friend Sub ReadFromClinObs(ByRef objClinObs As CLINTYPEATLLib.ClinicalObservation)
    Dim objClinObsMod As ClinicalObservation
    Dim i As Long

    'Medcin ID may be saved as string value SCR-40931
    If StrComp(objClinObs.ObsValue.value.InfoType, "basicString", vbTextCompare) = 0 Then
        msSnoID = objClinObs.ObsValue.value.BasicString
    Else
        msSnoID = CStr(objClinObs.ObsValue.value.ULong)
    End If
    ' Work around 3M memory leak in their _newenum
    For i = 1 To objClinObs.ObsMods.Count
    Set objClinObsMod = objClinObs.ObsMods.Item(i)
'    For Each objClinObsMod In objClinObs.ObsMods
      If Not objClinObsMod.ObsValue.IsEmpty Then
        Select Case objClinObsMod.ObsId.value.CodedWOSform.ValueConcept
        Case NCID_MEDCIN_DURATIONOBS
            msDuration = objClinObsMod.ObsValue.value.BasicString
        Case NCID_MEDCIN_MODIFIEROBS
            msModifier = objClinObsMod.ObsValue.value.BasicString
        Case NCID_MEDCIN_ONSETOBS
            msOnset = objClinObsMod.ObsValue.value.BasicString
        Case NCID_MEDCIN_PREFIXOBS
            msPrefix = objClinObsMod.ObsValue.value.BasicString
        Case NCID_MEDCIN_QLINKOBS
            msQualifierLink = objClinObsMod.ObsValue.value.BasicString
        Case NCID_MEDCIN_QLINKSEQOBS
            miQualifierLinkSequence = objClinObsMod.ObsValue.value.ULong
        Case NCID_MEDCIN_RANGENORMALHIGHOBS
            mnRangeNormalHigh = objClinObsMod.ObsValue.value.SDecimal.Decimal
        Case NCID_MEDCIN_RANGENORMALLOWOBS
            mnRangeNormalLow = objClinObsMod.ObsValue.value.SDecimal.Decimal
        Case NCID_MEDCIN_RANGESCALEOBS
            miRangeScale = objClinObsMod.ObsValue.value.ULong
        Case NCID_MEDCIN_REFERIDOBS
            mlReferid = objClinObsMod.ObsValue.value.ULong
        Case NCID_MEDCIN_RESULTOBS
            msResult = objClinObsMod.ObsValue.value.BasicString
        Case NCID_MEDCIN_STATUSOBS
            msStatus = objClinObsMod.ObsValue.value.BasicString
        Case NCID_MEDCIN_UNITOBS
            msUnit = objClinObsMod.ObsValue.value.BasicString
        Case NCID_MEDCIN_VALUEOBS
            msValue = objClinObsMod.ObsValue.value.BasicString
        Case NCID_MEDCIN_CHARTFLAGOBS
            miFlag = objClinObsMod.ObsValue.value.ULong
        Case NCID_MEDCIN_NOTEOBS
            msNote = objClinObsMod.ObsValue.value.BasicString
        End Select
      End If
'    Next objClinObsMod
    Next i
End Sub

Friend Sub WriteToClinObs(ByRef objClinEvent As CLINTYPEATLLib.ClinicalEvent, Optional bSnoIDOnly As Boolean = False)
    Dim objClinObs As ClinicalObservation
    'Save as String value SCR-40931
    Set objClinObs = AddClinObs(objClinEvent, NCID_MEDCIN_OBS, msSnoID, "", "basicString")
    
    If Not bSnoIDOnly Then 'only store SnoIDs to save space in DB SCR-16307
        Call AddClinObsMod(objClinObs, NCID_MEDCIN_DURATIONOBS, msDuration, "", "basicString")
        Call AddClinObsMod(objClinObs, NCID_MEDCIN_MODIFIEROBS, msModifier, "", "basicString")
        Call AddClinObsMod(objClinObs, NCID_MEDCIN_ONSETOBS, msOnset, "", "basicString")
        Call AddClinObsMod(objClinObs, NCID_MEDCIN_PREFIXOBS, msPrefix, "", "basicString")
        Call AddClinObsMod(objClinObs, NCID_MEDCIN_QLINKOBS, msQualifierLink, "", "basicString")
        Call AddClinObsMod(objClinObs, NCID_MEDCIN_QLINKSEQOBS, miQualifierLinkSequence, "", "uLong")
        Call AddClinObsMod(objClinObs, NCID_MEDCIN_RANGENORMALHIGHOBS, mnRangeNormalHigh, "", "sDecimal")
        Call AddClinObsMod(objClinObs, NCID_MEDCIN_RANGENORMALLOWOBS, mnRangeNormalLow, "", "sDecimal")
        Call AddClinObsMod(objClinObs, NCID_MEDCIN_RANGESCALEOBS, miRangeScale, "", "uLong")
        Call AddClinObsMod(objClinObs, NCID_MEDCIN_REFERIDOBS, mlReferid, "", "uLong")
        Call AddClinObsMod(objClinObs, NCID_MEDCIN_RESULTOBS, msResult, "", "basicString")
        Call AddClinObsMod(objClinObs, NCID_MEDCIN_STATUSOBS, msStatus, "", "basicString")
        Call AddClinObsMod(objClinObs, NCID_MEDCIN_UNITOBS, msUnit, "", "basicString")
        Call AddClinObsMod(objClinObs, NCID_MEDCIN_VALUEOBS, msValue, "", "basicString")
        Call AddClinObsMod(objClinObs, NCID_MEDCIN_CHARTFLAGOBS, miFlag, "", "uLong")
        Call AddClinObsMod(objClinObs, NCID_MEDCIN_NOTEOBS, msNote, "", "basicString")
    End If
    Set objClinObs = Nothing
End Sub

Friend Sub GEMS_ReadFromClinObs(ByRef oSQL As ICHCSII_SQL.ISqlOps)
    On Error Resume Next
    msDataId = oSQL!DataId
    msDuration = "" & oSQL!Duration
    msModifier = "" & oSQL!Modifier
    msOnset = "" & oSQL!Onset
    msPrefix = "" & oSQL!prefix
    msQualifierLink = "" & oSQL!qlink
    miQualifierLinkSequence = Val("" & oSQL!qlinkseq)
    mnRangeNormalHigh = Val("" & oSQL!RangeNormalHigh)
    mnRangeNormalLow = Val("" & oSQL!RangeNormalLow)
    miRangeScale = Val("" & oSQL!RangeScale)
    mlReferid = Val("" & oSQL!Referid)
    msResult = "" & oSQL!result
    msStatus = "" & oSQL!Status
    msUnit = "" & oSQL!Unit
    msValue = "" & oSQL!value
    msNote = "" & oSQL!Note
    msSnoID = CStr("" & oSQL!SnoID)
    miFlag = Val("" & oSQL!ChartFlag)  'SCR-21653, 21648
End Sub

Friend Sub GEMS_WriteToClinObs(ByRef objISQL As ICHCSII_SQL.ISqlOpsEx)
    
    With objISQL
      !SnoID = msSnoID
      !Duration = msDuration
      !Modifier = msModifier
      !Onset = msOnset
      If msPrefix = "" Then msPrefix = " "
      !prefix = msPrefix
      !qlink = msQualifierLink
      !qlinkseq = miQualifierLinkSequence
      !RangeNormalHigh = mnRangeNormalHigh
      !RangeNormalLow = mnRangeNormalLow
      !RangeScale = miRangeScale
      !Referid = mlReferid
      !result = msResult
      !Status = msStatus
      !Unit = msUnit
      !value = msValue
      !Note = msNote
      !ChartFlag = miFlag  'SCR-21653, 21648 Was not saving flag which determines whether HPI or ROS
   End With
   
End Sub

'----------------------------------------
' Initialize the ListToolRecord object
'----------------------------------------
Private Sub Class_Initialize()
  msSnoID = ""
  msPrefix = " " 'PREFIX_BLANK
  msModifier = " "
  msResult = " " 'RESULT_BLANK
  msStatus = " "
  msOnset = ""
  msDuration = ""
  msValue = ""
  msUnit = ""
  miRangeScale = 0
  mnRangeNormalLow = 0
  mnRangeNormalHigh = 0
  mlReferid = 0
  '--Set msRx = Nothing
  miFlag = 0
  msNote = ""
  msEpisode = ""
  msQualifierLink = ""
  miQualifierLinkSequence = 0
End Sub

Public Property Get Duration() As String:                          Duration = msDuration:                             End Property
Public Property Get Episode() As String:                           Episode = msEpisode:                               End Property
Public Property Get Flag() As Integer:                             Flag = miFlag:                                     End Property
Public Property Get SnoID() As String:                          SnoID = msSnoID:                             End Property
Public Property Get Modifier() As String:                          Modifier = msModifier:                             End Property
Public Property Get Note() As String:                              Note = msNote:                                     End Property
Public Property Get Onset() As String:                             Onset = msOnset:                                   End Property
Public Property Get prefix() As String:                            prefix = msPrefix:                                 End Property
Public Property Get QualifierLink() As String:                     QualifierLink = msQualifierLink:                   End Property
Public Property Get QualifierLinkSequence() As Integer:            QualifierLinkSequence = miQualifierLinkSequence:   End Property
Public Property Get RangeNormalHigh() As Double:                   RangeNormalHigh = mnRangeNormalHigh:               End Property
Public Property Get RangeNormalLow() As Double:                    RangeNormalLow = mnRangeNormalLow:                 End Property
Public Property Get RangeScale() As Integer:                       RangeScale = miRangeScale:                         End Property
Public Property Get Referid() As Long:                             Referid = mlReferid:                               End Property
Public Property Get result() As String:                            result = msResult:                                 End Property
Public Property Get Status() As String:                            Status = msStatus:                                 End Property
Public Property Get Unit() As String:                              Unit = msUnit:                                     End Property
Public Property Get value() As String:                             value = msValue:                                   End Property
Public Property Get DataId() As Long:                              DataId = msDataId:                                 End Property
Public Property Get PKCText() As String:                           PKCText = msPKCText:                               End Property

Public Property Let Duration(ByVal vData As String):               msDuration = vData:                                End Property
Public Property Let Episode(ByVal vData As String):                msEpisode = vData:                                 End Property
Public Property Let Flag(ByVal vData As Integer):                  miFlag = vData:                                    End Property
Public Property Let SnoID(ByVal vData As String):               msSnoID = vData:                                End Property
Public Property Let Modifier(ByVal vData As String):               msModifier = vData:                                End Property
Public Property Let Note(ByVal vData As String):                   msNote = vData:                                    End Property
Public Property Let Onset(ByVal vData As String):                  msOnset = vData:                                   End Property
Public Property Let prefix(ByVal vData As String):                 msPrefix = vData:                                  End Property
Public Property Let QualifierLink(ByVal vData As String):          msQualifierLink = vData:                           End Property
Public Property Let QualifierLinkSequence(ByVal vData As Integer): miQualifierLinkSequence = vData:                   End Property
Public Property Let RangeNormalHigh(ByVal vData As Double):        mnRangeNormalHigh = vData:                         End Property
Public Property Let RangeNormalLow(ByVal vData As Double):         mnRangeNormalLow = vData:                          End Property
Public Property Let RangeScale(ByVal vData As Integer):            miRangeScale = vData:                              End Property
Public Property Let Referid(ByVal vData As Long):                  mlReferid = vData:                                 End Property
Public Property Let result(ByVal vData As String):                 msResult = vData:                                  End Property
Public Property Let Status(ByVal vData As String):                 msStatus = vData:                                  End Property
Public Property Let Unit(ByVal vData As String):                   msUnit = vData:                                    End Property
Public Property Let value(ByVal vData As String):                  msValue = vData:                                   End Property
Public Property Let DataId(ByVal vData As Long):                   msDataId = vData:                                  End Property
Public Property Let PKCText(ByVal vData As String):                msPKCText = vData:                                 End Property
