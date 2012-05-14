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
Attribute VB_Name = "clsExProsEsth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Implements clsISection

Private mvarlPatientID As Long
Private mvarlDentalEncounterID As Long

Private dfDirtyFlag As DirtyFlagEnum

'local variable(s) to hold property value(s)
Private mvarPatientPrimaryComplaint As String 'local copy
Private mvarBaseOfNose As Double ''''local copy
Private mvarBorderOfChin As Double ''''local copy
Private mvarLipLength As Double ''''local copy
Private mvarLipDiscrepancy As String 'local copy
Private mvarFacialMidline As Double ''''local copy
Private mvarFacialMidlineRightLeft As ExamProsEsthMidlineEnum 'local copy
Private mvarCentralExposedRepose As Double 'local copy
Private mvarDistalExtentOfSmile As Double ''''local copy
Private mvarIncisalEdges As String 'local copy
Private mvarToothColor As String 'local copy
Private mvarToothAlignment As String 'local copy
Private mvarPerioConsiderations As String 'local copy
Private mvarBuccalCorridors As String 'local copy
Private mvarBruxism As String 'local copy
Private mvarBruxismType As String 'local copy
Private mvarCEJLocated As String 'local copy
Private mvarSoundBone As String 'local copy
Private mvarSulcularDepth As String 'local copy
Private mvarAdequateTissue As String 'local copy
Private mvarPosteriorOcclusalPlane As String 'local copy
Private mvarPostOcclusalPlanePos As String
Private mvarCentralExposed As String ''''local copy
Private mvarUpperLip As String 'local copy
Private mvarClinicalCrownTeethNumbers As String 'local copy
Private mvarNaturallyShortTeethNumbers As String 'local copy
Private mvarIncisalTeethNumbers As String 'local copy
Private mvarExcessiveGingivalTeethNumbers As String 'local copy
Private mvarPassiveEruption As String 'local copy
Private mvarDentoalveolarExtrusion As String 'local copy
Private mvarVerticalMaxExcess As String 'local copy

'local copy of previous value
Private mvarPrevValue As Double

'Constants used to define the rows in the grid
Private Const cExProsEstBuccalRow = 1
Private Const cExProsEstBruxismRow = 2
Private Const cExProsEstBruxismTypeRow = 3
Private Const cExProsEstCEJRow = 4
Private Const cExProsEstSoundRow = 5
Private Const cExProsEstSulcularRow = 6
Private Const cExProsEstAdequateRow = 7
Private Const cExProsEstPosteriorRow = 8
Private Const cExProsEstPosteriorPosRow = 9
Private Const cExProsEstCentralRow = 10
Private Const cExProsEstUpperRow = 11
Private Const cExProsEstShortClinicalCrownRow = 12
Private Const cExProsEstNaturalRow = 13
Private Const cExProsEstIncisalRow = 14
Private Const cExProsEstExcessRow = 15
Private Const cExProsEstAlteredRow = 16
Private Const cExProsEstDentoRow = 17
Private Const cExProsEstVerticalRow = 18

'Constants used to define options in the Exam-Pros-Esthetic grid
Private Const cEstBuccalValues As String = " |Acceptable|Deficient|Full"
Private Const cEstBruxismValues As String = " |Y|N"
Private Const cEstBruxismTypeValues As String = " |Slight|Moderate|Severe"
Private Const cEstCEJLocatedValues As String = " |Y|N"
Private Const cEstAdequateTissueValues As String = " |Y|N"
Private Const cEstPosteriorValues As String = " |Regular|Irregular"
Private Const cEstPosteriorPosValues As String = " |High|Low|Properly"
Private Const cEstCentralExposedValues As String = " |-5mm|-4mm|-3mm|-2mm|-1mm|0mm|+1mm|+2mm|+3mm" _
                                                 & "|+4mm|+5mm|+6mm|+7mm|+8mm|+9mm|+10mm|+11mm|+12mm|>12mm"
Private Const cEstUpperLipValues As String = " |Short|Long|Hyperactive"
Private Const cEstClinicalCrownTeethNumberValues As String = " |Y|N"
Private Const cEstExcessValues As String = " |Y|N"
Private Const cEstAlteredPassiveValues As String = " |Y|N"
Private Const cEstDentoExtrusionValues As String = " |Y|N"
Private Const cEstVerticalExcessValues As String = " |Y|N"

Private Const cSQL_Insert As String = "PatientPrimaryComplaint,BaseOfNose,BorderOfChin,LipLength," _
& "LipDiscrepancy,FacialMidline,FacialMidlineRightLeft,CentralExposed,DistalExtentOfSmile," _
& "IncisalEdges,ToothColor,ToothAlignment," _
& "PerioConsiderations,BuccalCorridors,Bruxism,BruxismType,CEJLocated," _
& "SoundBone," _
& "SulcularDepth,AdequateTissue,PosteriorOcclusalPlane,PostOcclusalPlanePos," _
& "CentralExposedFullSmile,UpperLip,ClinicalCrownTeethNumbers,NaturallyShortTeethNumbers," _
& "IncisalTeethNumbers,ExcessiveGingivalTeethNumbers,PassiveEruption,DentoalveolarExtrusion,VerticalMaxExcess"

Private Const cDblMaxLen As Integer = 5
Private Const cTxtMaxLen As Integer = 255
Private Const cNotesMaxLen As Integer = 2000
 
Public Enum ExamProsEsthMidlineEnum
    exProsEsth_Not_Addressed = 0
    exProsEsth_Right = 1
    exProsEsth_Left = 2
End Enum

Public Property Let VerticalMaxExcess(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.VerticalMaxExcess = 5
    mvarVerticalMaxExcess = vData
    UpdateDirtyFlag
End Property


Public Property Get VerticalMaxExcess() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.VerticalMaxExcess
    VerticalMaxExcess = mvarVerticalMaxExcess
End Property



Public Property Let DentoalveolarExtrusion(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DentoalveolarExtrusion = 5
    mvarDentoalveolarExtrusion = vData
    UpdateDirtyFlag
End Property


Public Property Get DentoalveolarExtrusion() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DentoalveolarExtrusion
    DentoalveolarExtrusion = mvarDentoalveolarExtrusion
End Property



Public Property Let PassiveEruption(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PassiveEruption = 5
    mvarPassiveEruption = vData
    UpdateDirtyFlag
End Property


Public Property Get PassiveEruption() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PassiveEruption
    PassiveEruption = mvarPassiveEruption
End Property



Public Property Let ExcessiveGingivalTeethNumbers(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ExcessiveGingivalTeethNumbers = 5
    mvarExcessiveGingivalTeethNumbers = vData
    UpdateDirtyFlag
End Property


Public Property Get ExcessiveGingivalTeethNumbers() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ExcessiveGingivalTeethNumbers
    ExcessiveGingivalTeethNumbers = mvarExcessiveGingivalTeethNumbers
End Property



Public Property Let IncisalTeethNumbers(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IncisalTeethNumbers = 5
    mvarIncisalTeethNumbers = vData
    UpdateDirtyFlag
End Property


Public Property Get IncisalTeethNumbers() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IncisalTeethNumbers
    IncisalTeethNumbers = mvarIncisalTeethNumbers
End Property



Public Property Let NaturallyShortTeethNumbers(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ShortTeethNumbers = 5
    mvarNaturallyShortTeethNumbers = vData
    UpdateDirtyFlag
End Property


Public Property Get NaturallyShortTeethNumbers() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ShortTeethNumbers
    NaturallyShortTeethNumbers = mvarNaturallyShortTeethNumbers
End Property



Public Property Let ClinicalCrownTeethNumbers(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ClinicalCrownTeethNumbers = 5
    mvarClinicalCrownTeethNumbers = vData
    UpdateDirtyFlag
End Property


Public Property Get ClinicalCrownTeethNumbers() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ClinicalCrownTeethNumbers
    ClinicalCrownTeethNumbers = mvarClinicalCrownTeethNumbers
End Property



Public Property Let UpperLip(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UpperLip = 5
    mvarUpperLip = vData
    UpdateDirtyFlag
End Property


Public Property Get UpperLip() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UpperLip
    UpperLip = mvarUpperLip
End Property



Public Property Let CentralExposedRepose(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CentralExposedRepose = 5
    mvarCentralExposedRepose = vData
    UpdateDirtyFlag
End Property


Public Property Get CentralExposedRepose() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CentralExposedRepose
    CentralExposedRepose = mvarCentralExposedRepose
End Property



Public Property Let PosteriorOcclusalPlane(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PosteriorOcclusalPlane = 5
    mvarPosteriorOcclusalPlane = vData
    UpdateDirtyFlag
End Property


Public Property Get PosteriorOcclusalPlane() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PosteriorOcclusalPlane
    PosteriorOcclusalPlane = mvarPosteriorOcclusalPlane
End Property

Public Property Let PostOcclusalPlanePos(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PosteriorOcclusalPlane = 5
    mvarPostOcclusalPlanePos = vData
    UpdateDirtyFlag
End Property


Public Property Get PostOcclusalPlanePos() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PosteriorOcclusalPlane
    PostOcclusalPlanePos = mvarPostOcclusalPlanePos
End Property

Public Property Let AdequateTissue(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AdequateTissue = 5
    mvarAdequateTissue = vData
    UpdateDirtyFlag
End Property


Public Property Get AdequateTissue() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AdequateTissue
    AdequateTissue = mvarAdequateTissue
End Property



Public Property Let SulcularDepth(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SulcularDepth = 5
    mvarSulcularDepth = Left(vData, cTxtMaxLen)
    UpdateDirtyFlag
End Property


Public Property Get SulcularDepth() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SulcularDepth
    SulcularDepth = mvarSulcularDepth
End Property



Public Property Let SoundBone(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SoundBone = 5
    mvarSoundBone = Left(vData, cTxtMaxLen)
    UpdateDirtyFlag
End Property


Public Property Get SoundBone() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SoundBone
    SoundBone = mvarSoundBone
End Property



Public Property Let CEJLocated(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CEJLocated = 5
    mvarCEJLocated = vData
    UpdateDirtyFlag
End Property


Public Property Get CEJLocated() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CEJLocated
    CEJLocated = mvarCEJLocated
End Property



Public Property Let BruxismType(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BruxismType = 5
    mvarBruxismType = vData
    UpdateDirtyFlag
End Property


Public Property Get BruxismType() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BruxismType
    BruxismType = mvarBruxismType
End Property



Public Property Let Bruxism(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Bruxism = 5
    mvarBruxism = vData
    UpdateDirtyFlag
End Property


Public Property Get Bruxism() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Bruxism
    Bruxism = mvarBruxism
End Property



Public Property Let BuccalCorridors(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BuccalCorridors = 5
    mvarBuccalCorridors = vData
    UpdateDirtyFlag
End Property


Public Property Get BuccalCorridors() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BuccalCorridors
    BuccalCorridors = mvarBuccalCorridors
End Property



Public Property Let PerioConsiderations(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PerioConsiderations = 5
    If Len(vData) > 2000 Then
        goCWShared.ShowVBError 1000, "Text string too long", "NotesPlus.clsExProsthEsth", "String too long", vbInformation
    End If
    mvarPerioConsiderations = Left(vData, cNotesMaxLen)
    UpdateDirtyFlag
End Property


Public Property Get PerioConsiderations() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PerioConsiderations
    PerioConsiderations = mvarPerioConsiderations
End Property



Public Property Let ToothAlignment(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ToothAlignment = 5
    If Len(vData) > 255 Then
        goCWShared.ShowVBError 1000, "Text string too long", "NotesPlus.clsExProsthEsth", "String too long", vbInformation
    End If
    mvarToothAlignment = Left(vData, cTxtMaxLen)
    UpdateDirtyFlag
End Property


Public Property Get ToothAlignment() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ToothAlignment
    ToothAlignment = mvarToothAlignment
End Property



Public Property Let ToothColor(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ToothColor = 5
    If Len(vData) > 255 Then
        goCWShared.ShowVBError 1000, "Text string too long", "NotesPlus.clsExProsthEsth", "String too long", vbInformation
    End If
    mvarToothColor = Left(vData, cTxtMaxLen)
    UpdateDirtyFlag
End Property


Public Property Get ToothColor() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ToothColor
    ToothColor = mvarToothColor
End Property



Public Property Let IncisalEdges(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IncisalEdges = 5
    If Len(vData) > 255 Then
        goCWShared.ShowVBError 1000, "Text string too long", "NotesPlus.clsExProsthEsth", "String too long", vbInformation
    End If
    mvarIncisalEdges = Left(vData, cTxtMaxLen)
    UpdateDirtyFlag
End Property


Public Property Get IncisalEdges() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IncisalEdges
    IncisalEdges = mvarIncisalEdges
End Property



Public Property Let DistalExtentOfSmile(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DistalExtentOfSmile = 5
    mvarDistalExtentOfSmile = vData
    UpdateDirtyFlag
End Property


Public Property Get DistalExtentOfSmile() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DistalExtentOfSmile
    DistalExtentOfSmile = mvarDistalExtentOfSmile
End Property



Public Property Let CentralExposed(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CentralExposed = 5
    mvarCentralExposed = vData
    UpdateDirtyFlag
End Property


Public Property Get CentralExposed() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CentralExposed
    CentralExposed = mvarCentralExposed
End Property



Public Property Let FacialMidlineRightLeft(ByVal vData As ExamProsEsthMidlineEnum)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FacialMidlineRightLeft = 5
    mvarFacialMidlineRightLeft = vData
    UpdateDirtyFlag
End Property


Public Property Get FacialMidlineRightLeft() As ExamProsEsthMidlineEnum
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FacialMidlineRightLeft
    FacialMidlineRightLeft = mvarFacialMidlineRightLeft
End Property



Public Property Let FacialMidline(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FacialMidline = 5
    mvarFacialMidline = vData
    UpdateDirtyFlag
End Property


Public Property Get FacialMidline() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FacialMidline
    FacialMidline = mvarFacialMidline
End Property



Public Property Let LipDiscrepancy(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LipDiscrepancy = 5
    If Len(vData) > 255 Then
        goCWShared.ShowVBError 1000, "Text string too long", "NotesPlus.clsExProsthEsth", "String too long", vbInformation
    End If
    mvarLipDiscrepancy = Left(vData, cTxtMaxLen)
    UpdateDirtyFlag
End Property


Public Property Get LipDiscrepancy() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LipDiscrepancy
    LipDiscrepancy = mvarLipDiscrepancy
End Property



Public Property Let LipLength(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LipLength = 5
    mvarLipLength = vData
    UpdateDirtyFlag
End Property


Public Property Get LipLength() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LipLength
    LipLength = mvarLipLength
End Property



Public Property Let BorderOfChin(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BorderOfChin = 5
    mvarBorderOfChin = vData
    UpdateDirtyFlag
End Property


Public Property Get BorderOfChin() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BorderOfChin
    BorderOfChin = mvarBorderOfChin
End Property



Public Property Let BaseOfNose(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BaseOfNose = 5
    mvarBaseOfNose = vData
    UpdateDirtyFlag
End Property


Public Property Get BaseOfNose() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BaseOfNose
    BaseOfNose = mvarBaseOfNose
End Property



Public Property Let PatientPrimaryComplaint(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PatientPrimaryComplaint = 5
    If Len(vData) > 255 Then
        goCWShared.ShowVBError 1000, "Text string too long", "NotesPlus.clsExProsthEsth", "String too long", vbInformation
    End If
    mvarPatientPrimaryComplaint = Left(vData, cTxtMaxLen)
    UpdateDirtyFlag
End Property


Public Property Get PatientPrimaryComplaint() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PatientPrimaryComplaint
    PatientPrimaryComplaint = mvarPatientPrimaryComplaint
End Property

Public Property Get BuccalValues() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BuccalCorridors
    BuccalValues = cEstBuccalValues
End Property

Public Property Get RowBuccal() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BuccalCorridors
    RowBuccal = cExProsEstBuccalRow
End Property

Public Property Get BruxismValues() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BuccalCorridors
    BruxismValues = cEstBruxismValues
End Property

Public Property Get RowBruxism() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BuccalCorridors
    RowBruxism = cExProsEstBruxismRow
End Property

Public Property Get BruxismTypeValues() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BuccalCorridors
    BruxismTypeValues = cEstBruxismTypeValues
End Property

Public Property Get RowBruxismType() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BuccalCorridors
    RowBruxismType = cExProsEstBruxismTypeRow
End Property

Public Property Get PosteriorOccusalPlaneValues() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BuccalCorridors
    PosteriorOccusalPlaneValues = cEstPosteriorValues
End Property

Public Property Get PostOccusalPlanePosValues() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BuccalCorridors
    PostOccusalPlanePosValues = cEstPosteriorPosValues
End Property

Public Property Get RowPosteriorOcclusal() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BuccalCorridors
    RowPosteriorOcclusal = cExProsEstPosteriorRow
End Property

Public Property Get RowPostOcclusalPos() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BuccalCorridors
    RowPostOcclusalPos = cExProsEstPosteriorPosRow
End Property

Public Property Get RowCEJLocated() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BuccalCorridors
    RowCEJLocated = cExProsEstCEJRow
End Property

Public Property Get CEJLocatedValues() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BuccalCorridors
    CEJLocatedValues = cEstCEJLocatedValues
End Property

Public Property Get RowAdequateTissue() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BuccalCorridors
    RowAdequateTissue = cExProsEstAdequateRow
End Property

Public Property Get AdequateTissueValues() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BuccalCorridors
    AdequateTissueValues = cEstAdequateTissueValues
End Property

Public Property Get RowCentralExposed() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BuccalCorridors
    RowCentralExposed = cExProsEstCentralRow
End Property

Public Property Get CentralExposedValues() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BuccalCorridors
    CentralExposedValues = cEstCentralExposedValues
End Property

Public Property Get RowShortClinicalCrown() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BuccalCorridors
    RowShortClinicalCrown = cExProsEstShortClinicalCrownRow
End Property

Public Property Get ClinicalCrownTeethNumbersValues() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BuccalCorridors
    ClinicalCrownTeethNumbersValues = cEstClinicalCrownTeethNumberValues
End Property

Public Property Get RowExcess() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BuccalCorridors
    RowExcess = cExProsEstExcessRow
End Property

Public Property Get ExcessValues() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BuccalCorridors
    ExcessValues = cEstExcessValues
End Property

Public Property Get RowAlteredPassive() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BuccalCorridors
    RowAlteredPassive = cExProsEstAlteredRow
End Property

Public Property Get AlteredPassiveValues() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BuccalCorridors
    AlteredPassiveValues = cEstAlteredPassiveValues
End Property

Public Property Get RowDentoExtrusion() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BuccalCorridors
    RowDentoExtrusion = cExProsEstDentoRow
End Property

Public Property Get DentoExtrusionValues() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BuccalCorridors
    DentoExtrusionValues = cEstDentoExtrusionValues
End Property

Public Property Get RowVerticalExcess() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BuccalCorridors
    RowVerticalExcess = cExProsEstVerticalRow
End Property

Public Property Get VerticalExcessValues() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BuccalCorridors
    VerticalExcessValues = cEstVerticalExcessValues
End Property

Public Property Get RowSoundBone() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BuccalCorridors
    RowSoundBone = cExProsEstSoundRow
End Property

Public Property Get RowSulcularDepth() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BuccalCorridors
    RowSulcularDepth = cExProsEstSulcularRow
End Property

Public Property Get RowUpperLip() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BuccalCorridors
    RowUpperLip = cExProsEstUpperRow
End Property

Public Property Get UpperLipValues() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BuccalCorridors
    UpperLipValues = cEstUpperLipValues
End Property

Public Property Get RowNaturallyShortTeeth() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BuccalCorridors
    RowNaturallyShortTeeth = cExProsEstNaturalRow
End Property

Public Property Get RowIncisalWear() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BuccalCorridors
    RowIncisalWear = cExProsEstIncisalRow
End Property

Public Property Let PrevValue(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BaseOfNose = 5
    mvarPrevValue = vData
    'UpdateDirtyFlag
End Property


Public Property Get PrevValue() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BaseOfNose
    PrevValue = mvarPrevValue
End Property

'***********************************************************************
'Sub:LoadEstheticFindingsGrid
'Author:KDunne
'Date: 12/13/2001
'Purpose: Loads Esthetic Findings grid - initial settings.
'***********************************************************************
Public Sub LoadEstheticFindingsGrid(ByVal objFlexGrid As VSFlex7Ctl.vsFlexGrid)
    
Const sPROC_NAME = "LoadEstheticFindingsGrid"

On Error GoTo ErrHandler

    With objFlexGrid
        .Cols = 2
        .Rows = 19
        .Width = 5000 '4800
        .ColWidth(1) = 2000
        .ColWidth(0) = 2700 '2500
        .Cell(flexcpText, 0, 1) = "Findings:"

        .Cell(flexcpText, cExProsEstBuccalRow, 0) = "Buccal corridor:"
        .Cell(flexcpText, cExProsEstBruxismRow, 0) = "Bruxism exists:"
        .Cell(flexcpText, cExProsEstBruxismTypeRow, 0) = "Bruxism degree:"
        .Cell(flexcpText, cExProsEstCEJRow, 0) = "CEJ located?"
        .Cell(flexcpText, cExProsEstSoundRow, 0) = "Sound bone:"
        .Cell(flexcpText, cExProsEstSulcularRow, 0) = "Sulcular depth:"
        .Cell(flexcpText, cExProsEstAdequateRow, 0) = "Adequate keratinized tissue?"
        .Cell(flexcpText, cExProsEstPosteriorRow, 0) = "Posterior occlusal plane:"
        .Cell(flexcpText, cExProsEstPosteriorPosRow, 0) = "Posterior occlusal plane positioned:"
        .Cell(flexcpText, cExProsEstCentralRow, 0) = "Central exposed:"
        .Cell(flexcpText, cExProsEstUpperRow, 0) = "Upper lip:"
        .Cell(flexcpText, cExProsEstShortClinicalCrownRow, 0) = "Short clinical crown teeth #'s:"
        .Cell(flexcpText, cExProsEstNaturalRow, 0) = "Naturally short teeth #'s:"
        .Cell(flexcpText, cExProsEstIncisalRow, 0) = "Incisal wear - #'s:"
        .Cell(flexcpText, cExProsEstExcessRow, 0) = "Excessive gingival coverage - #'s:"
        .Cell(flexcpText, cExProsEstAlteredRow, 0) = "Altered passive eruption?"
        .Cell(flexcpText, cExProsEstDentoRow, 0) = "Dentoalveolar extrusion?"
        .Cell(flexcpText, cExProsEstVerticalRow, 0) = "Vertical maxillary excess?"
    
        .Editable = flexEDKbdMouse
        .ScrollBars = flexScrollBarBoth
    
    End With ' objFlexGrid
    
Exit Sub
    
SubExit:
    Set objFlexGrid = Nothing
    Exit Sub
Exit Sub

ErrHandler:
    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & "." & sPROC_NAME, App.Title, vbCritical
    GoTo SubExit
End Sub
'***********************************************************************
'Sub:RefreshEstheticFindingsGrid
'Author:ASamedov
'Date: 12/16/2001
'Purpose: Updates Esthetic findings grid with selected values.
'***********************************************************************
Public Sub RefreshEstheticFindingsGrid(ByVal objExProsEsth As clsExProsEsth, _
                                       ByVal objFlexGrid As VSFlex7Ctl.vsFlexGrid)
Dim objToothCtrl As CToothControl
On Error GoTo ErrHandler

Const sPROC_NAME = "RefreshEstheticFindingsGrid"
    Set objToothCtrl = New CToothControl
    With objFlexGrid
        .Cell(flexcpText, cExProsEstBuccalRow, 1) = objExProsEsth.BuccalCorridors
        .Cell(flexcpText, cExProsEstBruxismRow, 1) = objExProsEsth.Bruxism
        .Cell(flexcpText, cExProsEstBruxismTypeRow, 1) = objExProsEsth.BruxismType
        .Cell(flexcpText, cExProsEstCEJRow, 1) = objExProsEsth.CEJLocated
        .Cell(flexcpText, cExProsEstSoundRow, 1) = objExProsEsth.SoundBone
        .Cell(flexcpText, cExProsEstSulcularRow, 1) = objExProsEsth.SulcularDepth
        .Cell(flexcpText, cExProsEstAdequateRow, 1) = objExProsEsth.AdequateTissue
        .Cell(flexcpText, cExProsEstPosteriorRow, 1) = objExProsEsth.PosteriorOcclusalPlane
        .Cell(flexcpText, cExProsEstPosteriorPosRow, 1) = objExProsEsth.PostOcclusalPlanePos
        .Cell(flexcpText, cExProsEstCentralRow, 1) = objExProsEsth.CentralExposed
        .Cell(flexcpText, cExProsEstUpperRow, 1) = objExProsEsth.UpperLip
        .Cell(flexcpText, cExProsEstShortClinicalCrownRow, 1) = objToothCtrl.BuildDisplayString(objExProsEsth.ClinicalCrownTeethNumbers)
        .Cell(flexcpText, cExProsEstNaturalRow, 1) = objToothCtrl.BuildDisplayString(objExProsEsth.NaturallyShortTeethNumbers)
        .Cell(flexcpText, cExProsEstIncisalRow, 1) = objToothCtrl.BuildDisplayString(objExProsEsth.IncisalTeethNumbers)
        .Cell(flexcpText, cExProsEstExcessRow, 1) = objToothCtrl.BuildDisplayString(objExProsEsth.ExcessiveGingivalTeethNumbers)
        .Cell(flexcpText, cExProsEstAlteredRow, 1) = objExProsEsth.PassiveEruption
        .Cell(flexcpText, cExProsEstDentoRow, 1) = objExProsEsth.DentoalveolarExtrusion
        .Cell(flexcpText, cExProsEstVerticalRow, 1) = objExProsEsth.VerticalMaxExcess
        
    End With

    Set objToothCtrl = Nothing
Exit Sub

ErrHandler:
    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & "." & sPROC_NAME, App.Title, vbCritical

End Sub
 

Private Sub Class_Initialize()
    dfDirtyFlag = dfNoAction
End Sub

Private Property Get clsISection_DentalEncounterID() As Long

    clsISection_DentalEncounterID = mvarlDentalEncounterID

End Property

Private Function clsISection_GetNote() As String
Dim sTemp As String, sTemp1 As String, sTemp2 As String, sTemp3 As String
Dim objToothCtrl As CToothControl

Const sPROC_NAME = "GetNote"

On Error GoTo ErrHnd
    
    Set objToothCtrl = New CToothControl
    
    'Esthetic Evaluation section
    If mvarPatientPrimaryComplaint <> "" Then
        sTemp1 = sTemp1 & vbTab & "Patient's primary cosmetic complaint: " & mvarPatientPrimaryComplaint & vbCrLf
    End If
    
    If mvarBaseOfNose <> 0 Then
        sTemp1 = sTemp1 & vbTab & "Glabella/base of nose (in mm): " & mvarBaseOfNose & vbCrLf
    End If
    
    If mvarBorderOfChin <> 0 Then
        sTemp1 = sTemp1 & vbTab & "Base of nose/border of chin (in mm): " & mvarBorderOfChin & vbCrLf
    End If
       
    If mvarLipLength <> 0 Then
        sTemp1 = sTemp1 & vbTab & "Lip length (in mm): " & mvarLipLength & vbCrLf
    End If
       
    If mvarLipDiscrepancy <> "" Then
        sTemp1 = sTemp1 & vbTab & "Lip discrepancy: " & mvarLipDiscrepancy & vbCrLf
    End If
       
    If mvarFacialMidline <> 0 Then
        sTemp1 = sTemp1 & vbTab & "Dental/Facial midline (in mm): " & mvarFacialMidline & vbCrLf
    End If
       
    If mvarFacialMidlineRightLeft <> ExamProsEsthMidlineEnum.exProsEsth_Not_Addressed Then
        sTemp1 = sTemp1 & vbTab & "Dental/Facial midline: " & GetMidLineEnumDesc(mvarFacialMidlineRightLeft) & vbCrLf
    End If
       
    If mvarCentralExposedRepose <> 0 Then
        sTemp1 = sTemp1 & vbTab & "Central exposed in repose (in mm): " & mvarCentralExposedRepose & vbCrLf
    End If
       
    If mvarDistalExtentOfSmile <> 0 Then
        sTemp1 = sTemp1 & vbTab & "Distal extent of smile (tooth #): " & mvarDistalExtentOfSmile & vbCrLf
    End If
       
    If mvarIncisalEdges <> "" Then
        sTemp1 = sTemp1 & vbTab & "Incisal edges to lower lip in smile: " & mvarIncisalEdges & vbCrLf
    End If
       
     If mvarToothColor <> "" Then
        sTemp1 = sTemp1 & vbTab & "Tooth color: " & mvarToothColor & vbCrLf
    End If
      
    If mvarToothAlignment <> "" Then
        sTemp1 = sTemp1 & vbTab & "Tooth alignment: " & mvarToothAlignment & vbCrLf
    End If
      
    If mvarPerioConsiderations <> "" Then
        sTemp1 = sTemp1 & vbTab & "Periodontal considerations: " & mvarPerioConsiderations & vbCrLf
    End If
          
    'Esthetic Evaluation Header
    If sTemp1 <> "" Then
        sTemp1 = vbCrLf & "Esthetic Evaluation: " & vbCrLf & sTemp1 & vbCrLf
    End If
       
    'Findings section
    If mvarBuccalCorridors <> "" Then
        sTemp2 = sTemp2 & vbTab & "Buccal corridor: " & mvarBuccalCorridors & vbCrLf
    End If

    If mvarBruxism <> "" Then
        sTemp2 = sTemp2 & vbTab & "Bruxism exists: " & mvarBruxism & vbCrLf
    End If

    If mvarBruxismType <> "" Then
        sTemp2 = sTemp2 & vbTab & "Bruxism degree: " & mvarBruxismType & vbCrLf
    End If

    If mvarCEJLocated <> "" Then
        sTemp2 = sTemp2 & vbTab & "CEJ located?: " & mvarCEJLocated & vbCrLf
    End If

    If mvarSoundBone <> "" Then
        sTemp2 = sTemp2 & vbTab & "Sound bone: " & mvarSoundBone & vbCrLf
    End If

    If mvarSulcularDepth <> "" Then
        sTemp2 = sTemp2 & vbTab & "Sulcular depth: " & mvarSulcularDepth & vbCrLf
    End If

    If mvarAdequateTissue <> "" Then
        sTemp2 = sTemp2 & vbTab & "Adequate keratinized tissue?: " & mvarAdequateTissue & vbCrLf
    End If

    If mvarPosteriorOcclusalPlane <> "" Then
        sTemp2 = sTemp2 & vbTab & "Posterior occlusal plane: " & mvarPosteriorOcclusalPlane & vbCrLf
    End If

    If mvarPostOcclusalPlanePos <> "" Then
        sTemp2 = sTemp2 & vbTab & "Posterior occlusal plane positioned: " & mvarPostOcclusalPlanePos & vbCrLf
    End If

    If mvarCentralExposed <> "" Then
        sTemp2 = sTemp2 & vbTab & "Central exposed: " & mvarCentralExposed & vbCrLf
    End If

    If Trim(mvarUpperLip) <> "" Then
        sTemp2 = sTemp2 & vbTab & "Upper lip: " & mvarUpperLip & vbCrLf
    End If

    If mvarClinicalCrownTeethNumbers <> "" Then
        sTemp2 = sTemp2 & vbTab & "Short clinical crown teeth #'s: " & objToothCtrl.BuildDisplayString(mvarClinicalCrownTeethNumbers) & vbCrLf
    End If

    If mvarNaturallyShortTeethNumbers <> "" Then
        sTemp2 = sTemp2 & vbTab & "Naturally short teeth #'s: " & objToothCtrl.BuildDisplayString(mvarNaturallyShortTeethNumbers) & vbCrLf
    End If

    If mvarIncisalTeethNumbers <> "" Then
        sTemp2 = sTemp2 & vbTab & "Incisal wear - #'s: " & objToothCtrl.BuildDisplayString(mvarIncisalTeethNumbers) & vbCrLf
    End If
    
    If mvarExcessiveGingivalTeethNumbers <> "" Then
        sTemp2 = sTemp2 & vbTab & "Excessive gingival coverage - #'s: " & objToothCtrl.BuildDisplayString(mvarExcessiveGingivalTeethNumbers) & vbCrLf
    End If
    
    If mvarPassiveEruption <> "" Then
        sTemp2 = sTemp2 & vbTab & "Altered passive eruption?: " & mvarPassiveEruption & vbCrLf
    End If
    
    If mvarDentoalveolarExtrusion <> "" Then
        sTemp2 = sTemp2 & vbTab & "Dentoalveolar extrusion?: " & mvarDentoalveolarExtrusion & vbCrLf
    End If
    
    If mvarVerticalMaxExcess <> "" Then
        sTemp2 = sTemp2 & vbTab & "Vertical maxillary excess?: " & mvarVerticalMaxExcess & vbCrLf
    End If

    'Findings Header
    If sTemp2 <> "" Then
        sTemp2 = vbCrLf & "Findings: " & vbCrLf & sTemp2 & vbCrLf & vbCrLf
    End If

    'Summary
    sTemp = sTemp1 & " " & sTemp2
       
    clsISection_GetNote = sTemp
    
FuncExit:
    Set objToothCtrl = Nothing
Exit Function

ErrHnd:
    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & "." & sPROC_NAME, App.Title, vbCritical
    GoTo FuncExit
End Function

Private Function clsISection_GetRTFNote() As String
Dim sTemp As String, sTemp1 As String, sTemp2 As String, sTemp3 As String
Dim objToothCtrl As CToothControl

Const sPROC_NAME = "GetRTFNote"

On Error GoTo ErrHnd
    
    Set objToothCtrl = New CToothControl
    
    'Esthetic Evaluation section
    If Trim(mvarPatientPrimaryComplaint) <> "" Then
        sTemp1 = sTemp1 & "Patient's primary cosmetic complaint: " & mvarPatientPrimaryComplaint & "; "
    End If
    
    If mvarBaseOfNose <> 0 Then
        sTemp1 = sTemp1 & "Glabella/base of nose (in mm): " & mvarBaseOfNose & "; "
    End If
    
    If mvarBorderOfChin <> 0 Then
        sTemp1 = sTemp1 & "Base of nose/border of chin (in mm): " & mvarBorderOfChin & "; "
    End If
       
    If mvarLipLength <> 0 Then
        sTemp1 = sTemp1 & "Lip length (in mm): " & mvarLipLength & "; "
    End If
       
    If Trim(mvarLipDiscrepancy) <> "" Then
        sTemp1 = sTemp1 & "Lip discrepancy: " & mvarLipDiscrepancy & "; "
    End If
       
    If mvarFacialMidline <> 0 Then
        sTemp1 = sTemp1 & "Dental/Facial midline (in mm): " & mvarFacialMidline & "; "
    End If
       
    If mvarFacialMidlineRightLeft <> ExamProsEsthMidlineEnum.exProsEsth_Not_Addressed Then
        sTemp1 = sTemp1 & "Dental/Facial midline: " & GetMidLineEnumDesc(mvarFacialMidlineRightLeft) & "; "
    End If
       
    If mvarCentralExposedRepose <> 0 Then
        sTemp1 = sTemp1 & "Central exposed in repose (in mm): " & mvarCentralExposedRepose & "; "
    End If
       
    If mvarDistalExtentOfSmile <> 0 Then
        sTemp1 = sTemp1 & "Distal extent of smile (tooth #): " & mvarDistalExtentOfSmile & "; "
    End If
       
    If Trim(mvarIncisalEdges) <> "" Then
        sTemp1 = sTemp1 & "Incisal edges to lower lip in smile: " & mvarIncisalEdges & "; "
    End If
       
     If Trim(mvarToothColor) <> "" Then
        sTemp1 = sTemp1 & "Tooth color: " & mvarToothColor & "; "
    End If
      
    If Trim(mvarToothAlignment) <> "" Then
        sTemp1 = sTemp1 & "Tooth alignment: " & mvarToothAlignment & "; "
    End If
      
    If Trim(mvarPerioConsiderations) <> "" Then
        sTemp1 = sTemp1 & "Periodontal considerations: " & mvarPerioConsiderations & "; "
    End If
          
    'Esthetic Evaluation Header
    If Trim(sTemp1) <> "" Then
        sTemp1 = crtfStartControlName & "Esthetic Evaluation: " & crtfEndControlName & sTemp1
    End If
       
    'Findings section
    If Trim(mvarBuccalCorridors) <> "" Then
        sTemp2 = sTemp2 & "Buccal corridor: " & mvarBuccalCorridors & "; "
    End If

    If Trim(mvarBruxism) <> "" Then
        sTemp2 = sTemp2 & "Bruxism exists: " & mvarBruxism & "; "
    End If

    If Trim(mvarBruxismType) <> "" Then
        sTemp2 = sTemp2 & "Bruxism degree: " & mvarBruxismType & "; "
    End If

    If Trim(mvarCEJLocated) <> "" Then
        sTemp2 = sTemp2 & "CEJ located?: " & mvarCEJLocated & "; "
    End If

    If Trim(mvarSoundBone) <> "" Then
        sTemp2 = sTemp2 & "Sound bone: " & mvarSoundBone & "; "
    End If

    If Trim(mvarSulcularDepth) <> "" Then
        sTemp2 = sTemp2 & "Sulcular depth: " & mvarSulcularDepth & "; "
    End If

    If Trim(mvarAdequateTissue) <> "" Then
        sTemp2 = sTemp2 & "Adequate keratinized tissue?: " & mvarAdequateTissue & "; "
    End If

    If Trim(mvarPosteriorOcclusalPlane) <> "" Then
        sTemp2 = sTemp2 & "Posterior occlusal plane: " & mvarPosteriorOcclusalPlane & "; "
    End If

    If Trim(mvarPostOcclusalPlanePos) <> "" Then
        sTemp2 = sTemp2 & "Posterior occlusal plane positioned: " & mvarPostOcclusalPlanePos & "; "
    End If

    If Trim(mvarCentralExposed) <> "" Then
        sTemp2 = sTemp2 & "Central exposed: " & mvarCentralExposed & "; "
    End If

    If Trim(mvarUpperLip) <> "" Then
        sTemp2 = sTemp2 & "Upper lip: " & mvarUpperLip & "; "
    End If

    If Trim(mvarClinicalCrownTeethNumbers) <> "" Then
        sTemp2 = sTemp2 & "Short clinical crown teeth #'s: " & objToothCtrl.BuildDisplayString(mvarClinicalCrownTeethNumbers) & "; "
    End If

    If Trim(mvarNaturallyShortTeethNumbers) <> "" Then
        sTemp2 = sTemp2 & "Naturally short teeth #'s: " & objToothCtrl.BuildDisplayString(mvarNaturallyShortTeethNumbers) & "; "
    End If

    If Trim(mvarIncisalTeethNumbers) <> "" Then
        sTemp2 = sTemp2 & "Incisal wear - #'s: " & objToothCtrl.BuildDisplayString(mvarIncisalTeethNumbers) & "; "
    End If
    
    If Trim(mvarExcessiveGingivalTeethNumbers) <> "" Then
        sTemp2 = sTemp2 & "Excessive gingival coverage - #'s: " & objToothCtrl.BuildDisplayString(mvarExcessiveGingivalTeethNumbers) & "; "
    End If
    
    If Trim(mvarPassiveEruption) <> "" Then
        sTemp2 = sTemp2 & "Altered passive eruption?: " & mvarPassiveEruption & "; "
    End If
    
    If Trim(mvarDentoalveolarExtrusion) <> "" Then
        sTemp2 = sTemp2 & "Dentoalveolar extrusion?: " & mvarDentoalveolarExtrusion & "; "
    End If
    
    If Trim(mvarVerticalMaxExcess) <> "" Then
        sTemp2 = sTemp2 & "Vertical maxillary excess?: " & mvarVerticalMaxExcess & "; "
    End If

    'Findings Header
    If sTemp2 <> "" Then
        sTemp2 = crtfStartControlName & "Findings: " & crtfEndControlName & sTemp2
    End If

    'Summary
    sTemp = sTemp1 & " " & sTemp2
       
    clsISection_GetRTFNote = sTemp
    
FuncExit:
    Set objToothCtrl = Nothing
Exit Function

ErrHnd:
    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & "." & sPROC_NAME, App.Title, vbCritical
    GoTo FuncExit

End Function

'*******************************************************************
'Function: LoadSection
'Author: ASamedov
'Date: 12/19/01
'Returns: True/False
'Purpose: This loads the data from the db
'*******************************************************************
Private Function clsISection_LoadSection(ByVal PatientID As Long, ByVal DentalEncounterID As Long) As Boolean

Dim sSQL As String, rs As Recordset, bReturnResult As Boolean
Const sPROC_NAME = "LoadSection"

On Error GoTo ErrHnd
    
    bReturnResult = True
    
    mvarlPatientID = PatientID
    mvarlDentalEncounterID = DentalEncounterID
    
    'Load the values
    
    sSQL = "SELECT "
    sSQL = sSQL & "PatientPrimaryComplaint,BaseOfNose,BorderOfChin,LipLength,LipDiscrepancy,"
    sSQL = sSQL & "FacialMidline,FacialMidlineRightLeft,CentralExposed,DistalExtentOfSmile,"
    sSQL = sSQL & "IncisalEdges,ToothColor,ToothAlignment,PerioConsiderations,BuccalCorridors,"
    sSQL = sSQL & "Bruxism,BruxismType,CEJLocated,SoundBone,SulcularDepth,AdequateTissue,PosteriorOcclusalPlane,PostOcclusalPlanePos,"
    sSQL = sSQL & "CentralExposedFullSmile,UpperLip,ClinicalCrownTeethNumbers,NaturallyShortTeethNumbers,"
    sSQL = sSQL & "IncisalTeethNumbers,ExcessiveGingivalTeethNumbers,PassiveEruption,DentoalveolarExtrusion,VerticalMaxExcess"

    sSQL = sSQL & " FROM TEMPL_EXAM_PROS_ESTHETIC"
    sSQL = sSQL & " WHERE TEMPL_EXAM_PROS_ESTHETIC.DENTAL_ENCOUNTER_ID = " & mvarlDentalEncounterID
        
    Set rs = goconn.OpenRecordset(sSQL)

    If rs.EOF Then
        dfDirtyFlag = dfAdd
        'Init the values here to "" ("False" is default for boolean)
        mvarPatientPrimaryComplaint = ""
        mvarBaseOfNose = 0
        mvarBorderOfChin = 0
        mvarLipLength = 0
        mvarLipDiscrepancy = ""
        mvarFacialMidline = 0
        'mvarFacialMidlineRightLeft = 0
        mvarCentralExposedRepose = 0
        mvarDistalExtentOfSmile = 0
        mvarIncisalEdges = ""
        mvarToothColor = ""
        mvarToothAlignment = ""
        mvarPerioConsiderations = ""
        mvarBuccalCorridors = ""
        mvarBruxism = ""
        mvarBruxismType = ""
        mvarCEJLocated = ""
        mvarSoundBone = ""
        mvarSulcularDepth = ""
        mvarAdequateTissue = ""
        mvarPosteriorOcclusalPlane = ""
        mvarPostOcclusalPlanePos = ""
        mvarCentralExposed = ""
        mvarUpperLip = ""
        mvarClinicalCrownTeethNumbers = ""
        mvarNaturallyShortTeethNumbers = ""
        mvarIncisalTeethNumbers = ""
        mvarExcessiveGingivalTeethNumbers = ""
        mvarPassiveEruption = ""
        mvarDentoalveolarExtrusion = ""
        mvarVerticalMaxExcess = ""
    Else
        dfDirtyFlag = dfNoAction

        'Set the values from the db including "" to allow for nulls
        'mvarlWorkloadID = Val(rs("Workload_ID") & "")

        mvarPatientPrimaryComplaint = rs("PatientPrimaryComplaint") & ""
        mvarBaseOfNose = Val(rs("BaseOfNose") & "")
        mvarBorderOfChin = Val(rs("BorderOfChin") & "")
        mvarLipLength = Val(rs("LipLength") & "")
        mvarLipDiscrepancy = rs("LipDiscrepancy") & ""
        mvarFacialMidline = Val(rs("FacialMidline") & "")
        mvarFacialMidlineRightLeft = Val(rs("FacialMidlineRightLeft") & "")
        mvarCentralExposedRepose = Val(rs("CentralExposed") & "")
        mvarDistalExtentOfSmile = Val(rs("DistalExtentOfSmile") & "")
        mvarIncisalEdges = rs("IncisalEdges") & ""
        mvarToothColor = rs("ToothColor") & ""
        mvarToothAlignment = rs("ToothAlignment") & ""
        mvarPerioConsiderations = rs("PerioConsiderations") & ""
        mvarBuccalCorridors = rs("BuccalCorridors") & ""
        mvarBruxism = rs("Bruxism") & ""
        mvarBruxismType = rs("BruxismType") & ""
        mvarCEJLocated = rs("CEJLocated") & ""
        mvarSoundBone = rs("SoundBone") & ""
        mvarSulcularDepth = rs("SulcularDepth") & ""
        mvarAdequateTissue = rs("AdequateTissue") & ""
        mvarPosteriorOcclusalPlane = rs("PosteriorOcclusalPlane") & ""
        mvarPostOcclusalPlanePos = rs("PostOcclusalPlanePos") & ""
        mvarCentralExposed = rs("CentralExposedFullSmile") & ""
        mvarUpperLip = rs("UpperLip") & ""
        mvarClinicalCrownTeethNumbers = rs("ClinicalCrownTeethNumbers") & ""
        mvarNaturallyShortTeethNumbers = rs("NaturallyShortTeethNumbers") & ""
        mvarIncisalTeethNumbers = rs("IncisalTeethNumbers") & ""
        mvarExcessiveGingivalTeethNumbers = rs("ExcessiveGingivalTeethNumbers") & ""
        mvarPassiveEruption = rs("PassiveEruption") & ""
        mvarDentoalveolarExtrusion = rs("DentoalveolarExtrusion") & ""
        mvarVerticalMaxExcess = rs("VerticalMaxExcess") & ""
        
    End If

    rs.Close

    Set rs = Nothing
   
    clsISection_LoadSection = bReturnResult
    
Exit Function

ErrHnd:
    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & "." & sPROC_NAME, App.Title, vbCritical

End Function

Private Property Get clsISection_PatientID() As Long
    
    clsISection_PatientID = mvarlPatientID

End Property

'*******************************************************************
'Function: SaveSection
'Author: Your name here
'Date:
'Returns: True/False
'Purpose: Saves the data to the table
'*******************************************************************
Private Function clsISection_SaveSection() As Boolean

Dim sSQL As String
Dim iLoop As Integer

Const sPROC_NAME As String = "SaveSection"

On Error GoTo ErrHnd

    sSQL = ""

    If dfDirtyFlag = dfNoAction Then
        clsISection_SaveSection = True
        Exit Function
    Else
        'Update the workload
                
    End If
        
    If dfDirtyFlag = dfAdd Then
        'The record does not exist, so create an insert statement
        sSQL = "INSERT INTO TEMPL_EXAM_PROS_ESTHETIC ("
        sSQL = sSQL & "Dental_Encounter_ID,"
        sSQL = sSQL & cSQL_Insert & " "

        sSQL = sSQL & ") VALUES ("

        sSQL = sSQL & CLng(mvarlDentalEncounterID) & ","
        sSQL = sSQL & "'" & RQ(mvarPatientPrimaryComplaint) & "'" & ","
        sSQL = sSQL & CDbl(mvarBaseOfNose) & ","
        sSQL = sSQL & CDbl(mvarBorderOfChin) & ","
        sSQL = sSQL & CDbl(mvarLipLength) & ","
        sSQL = sSQL & "'" & RQ(mvarLipDiscrepancy) & "'" & ","
        sSQL = sSQL & CDbl(mvarFacialMidline) & ","
        sSQL = sSQL & CInt(mvarFacialMidlineRightLeft) & ","
        sSQL = sSQL & CDbl(mvarCentralExposedRepose) & ","
        sSQL = sSQL & CDbl(mvarDistalExtentOfSmile) & ","
        sSQL = sSQL & "'" & RQ(mvarIncisalEdges) & "'" & ","
        sSQL = sSQL & "'" & RQ(mvarToothColor) & "'" & ","
        sSQL = sSQL & "'" & RQ(mvarToothAlignment) & "'" & ","
        sSQL = sSQL & "'" & RQ(mvarPerioConsiderations) & "'" & ","
        
        sSQL = sSQL & "'" & mvarBuccalCorridors & "'" & ","
        sSQL = sSQL & "'" & mvarBruxism & "'" & ","
        sSQL = sSQL & "'" & mvarBruxismType & "'" & ","
        sSQL = sSQL & "'" & mvarCEJLocated & "'" & ","
        sSQL = sSQL & "'" & RQ(mvarSoundBone) & "'" & ","
        sSQL = sSQL & "'" & RQ(mvarSulcularDepth) & "'" & ","
        sSQL = sSQL & "'" & mvarAdequateTissue & "'" & ","
        sSQL = sSQL & "'" & mvarPosteriorOcclusalPlane & "'" & ","
        sSQL = sSQL & "'" & mvarPostOcclusalPlanePos & "'" & ","
        sSQL = sSQL & "'" & mvarCentralExposed & "'" & ","
        sSQL = sSQL & "'" & mvarUpperLip & "'" & ","
        sSQL = sSQL & "'" & RQ(mvarClinicalCrownTeethNumbers) & "'" & ","
        sSQL = sSQL & "'" & RQ(mvarNaturallyShortTeethNumbers) & "'" & ","
        sSQL = sSQL & "'" & RQ(mvarIncisalTeethNumbers) & "'" & ","
        sSQL = sSQL & "'" & RQ(mvarExcessiveGingivalTeethNumbers) & "'" & ","
        sSQL = sSQL & "'" & RQ(mvarPassiveEruption) & "'" & ","
        sSQL = sSQL & "'" & RQ(mvarDentoalveolarExtrusion) & "'" & ","
        sSQL = sSQL & "'" & RQ(mvarVerticalMaxExcess) & "'"
        
        sSQL = sSQL & ")"

    ElseIf dfDirtyFlag = dfUpdate Then
        'The record exists, so create an update statement
        sSQL = ""

        sSQL = "UPDATE TEMPL_EXAM_PROS_ESTHETIC SET "
        
        sSQL = sSQL & "PATIENTPRIMARYCOMPLAINT = " & "'" & RQ(mvarPatientPrimaryComplaint) & "'" & ","
        sSQL = sSQL & "BASEOFNOSE = " & CDbl(mvarBaseOfNose) & ","
        sSQL = sSQL & "BORDEROFCHIN = " & CDbl(mvarBorderOfChin) & ","
        sSQL = sSQL & "LIPLENGTH = " & CDbl(mvarLipLength) & ","
        sSQL = sSQL & "LIPDISCREPANCY = " & "'" & RQ(mvarLipDiscrepancy) & "'" & ","
        sSQL = sSQL & "FACIALMIDLINE = " & CDbl(mvarFacialMidline) & ","
        sSQL = sSQL & "FACIALMIDLINERIGHTLEFT = " & CInt(mvarFacialMidlineRightLeft) & ","
        sSQL = sSQL & "CENTRALEXPOSED = " & CDbl(mvarCentralExposedRepose) & ","
        sSQL = sSQL & "DISTALEXTENTOFSMILE = " & CDbl(mvarDistalExtentOfSmile) & ","
        sSQL = sSQL & "INCISALEDGES = " & "'" & RQ(mvarIncisalEdges) & "'" & ","
        sSQL = sSQL & "TOOTHCOLOR = " & "'" & RQ(mvarToothColor) & "'" & ","
        sSQL = sSQL & "TOOTHALIGNMENT = " & "'" & RQ(mvarToothAlignment) & "'" & ","
        sSQL = sSQL & "PERIOCONSIDERATIONS = " & "'" & RQ(mvarPerioConsiderations) & "'" & ","
        
        sSQL = sSQL & "BUCCALCORRIDORS = " & "'" & mvarBuccalCorridors & "'" & ","
        sSQL = sSQL & "BRUXISM = " & "'" & mvarBruxism & "'" & ","
        sSQL = sSQL & "BRUXISMTYPE = " & "'" & mvarBruxismType & "'" & ","
        sSQL = sSQL & "CEJLOCATED = " & "'" & mvarCEJLocated & "'" & ","
        sSQL = sSQL & "SOUNDBONE = " & "'" & RQ(mvarSoundBone) & "'" & ","
        sSQL = sSQL & "SULCULARDEPTH = " & "'" & RQ(mvarSulcularDepth) & "'" & ","
        sSQL = sSQL & "ADEQUATETISSUE = " & "'" & mvarAdequateTissue & "'" & ","
        sSQL = sSQL & "POSTERIOROCCLUSALPLANE = " & "'" & mvarPosteriorOcclusalPlane & "'" & ","
        sSQL = sSQL & "POSTOCCLUSALPLANEPOS = " & "'" & mvarPostOcclusalPlanePos & "'" & ","
        sSQL = sSQL & "CENTRALEXPOSEDFULLSMILE = " & "'" & mvarCentralExposed & "'" & ","
        sSQL = sSQL & "UPPERLIP = " & "'" & mvarUpperLip & "'" & ","
        sSQL = sSQL & "CLINICALCROWNTEETHNUMBERS = " & "'" & RQ(mvarClinicalCrownTeethNumbers) & "'" & ","
        sSQL = sSQL & "NATURALLYSHORTTEETHNUMBERS = " & "'" & RQ(mvarNaturallyShortTeethNumbers) & "'" & ","
        sSQL = sSQL & "INCISALTEETHNUMBERS = " & "'" & RQ(mvarIncisalTeethNumbers) & "'" & ","
        sSQL = sSQL & "EXCESSIVEGINGIVALTEETHNUMBERS = " & "'" & RQ(mvarExcessiveGingivalTeethNumbers) & "'" & ","
        sSQL = sSQL & "PASSIVEERUPTION = " & "'" & RQ(mvarPassiveEruption) & "'" & ","
        sSQL = sSQL & "DENTOALVEOLAREXTRUSION = " & "'" & RQ(mvarDentoalveolarExtrusion) & "'" & ","
        sSQL = sSQL & "VERTICALMAXEXCESS = " & "'" & RQ(mvarVerticalMaxExcess) & "' "

        sSQL = sSQL & "WHERE (DENTAL_ENCOUNTER_ID = " & mvarlDentalEncounterID & ")"

    ElseIf dfDirtyFlag = dfdelete Then
        'Mark the record as deleted
        sSQL = ""
    End If

    goconn.ExecuteSQL sSQL
    
    dfDirtyFlag = dfNoAction
    
    clsISection_SaveSection = True

Exit Function

ErrHnd:
    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & "." & sPROC_NAME, App.Title, vbCritical
    
    'Delete the workload credit if it didn't update correctly
    'UpdateWorkload mvarlWLID_Nutritional, dfdelete, mvarlDentalEncounterID, "", 1

End Function

Private Function UpdateDirtyFlag()

    If dfDirtyFlag = dfNoAction Then
        dfDirtyFlag = dfUpdate
    End If
    
End Function

'******************************************************************************************
'Function:GetMidLineDescription
'Author:KDunne
'Date:4/25/2002
'Purpose: Accepts a ExamProsEsthMidlineEnum value and returns the appropriate description.
'******************************************************************************************
Public Function GetMidLineEnumDesc(ByVal lngValue As ExamProsEsthMidlineEnum) As String
    Dim strReturn As String
    
    Select Case lngValue
        Case ExamProsEsthMidlineEnum.exProsEsth_Right
            strReturn = "Right"
        Case ExamProsEsthMidlineEnum.exProsEsth_Left
            strReturn = "Left"
        Case ExamProsEsthMidlineEnum.exProsEsth_Not_Addressed
            strReturn = Empty
        Case Else
            strReturn = Empty
    End Select
    
    GetMidLineEnumDesc = strReturn
End Function
