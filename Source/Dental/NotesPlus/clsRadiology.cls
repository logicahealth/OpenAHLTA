VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRadiology"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Implements clsISection

Private mvarlPatientID As Long
Private mvarlDentalEncounterID As Long

Private dfDirtyFlag As DirtyFlagEnum

'local variable(s) to hold property value(s)
Private mvarbNoRadTaken As Boolean 'local copy
Private mvarbPrevRadRev As Boolean
Private mvarbSingleBWXRTaken As Boolean
Private mvarb2BWXRTaken As Boolean
Private mvarb4BWXRTaken As Boolean
Private mvarbPeriapicalRadTaken As Boolean
Private mvarbPANOTaken As Boolean
Private mvarbCompleteSeriesBWTaken As Boolean
Private mvarbCephalometricFilmTaken As Boolean
Private mvarbOcclusalFilmMax As Boolean
Private mvarbOcclusalFilmMan As Boolean
Private mvarbOtherFilms As Boolean

Private mvarbExtraoralInit As Boolean
Private mvarbExtraoralAdd As Boolean
Private mvarbPALateralSkull As Boolean
Private mvarbSialography As Boolean
Private mvarbTMJArthrogram As Boolean
Private mvarbOtherTMJ As Boolean
Private mvarbTomographic As Boolean

Private mvarsPrevRadRevFindString As String 'local copy
Private mvarsSingleBWXRFindString As String 'local copy
Private mvars2BWXRFindString As String 'local copy
Private mvars4BWXRFindString As String 'local copy
Private mvarsPANOFindString As String 'local copy
Private mvarsCompleteSeriesFindString As String 'local copy
Private mvarsCephalometricFilmFindString As String 'local copy
Private mvarsOcclusalFilmMaxFindString As String 'local copy
Private mvarsOcclusalFilmManFindString As String 'local copy
Private mvarsOtherFindString As String 'local copy

Private mvarsExtraoralInitFindString As String
Private mvarsExtraoralAddFindString As String
Private mvarsPALateralSkullFindString As String
Private mvarsSialographyFindString As String
Private mvarsTMJArthrogramFindString As String
Private mvarsOtherTMJFindString As String
Private mvarsTomographicFindString As String

Private mvariNumberPA As Integer

'combobox
Private mvarsOtherFilmsComboBox As String

Private mvarsRadNotes As String 'local copy

Private mvarcolPrevRadRevFindings As Collection
Private mvarcolSingleBWXRTakenFindings As Collection
Private mvarcol2BWXRTakenFindings As Collection
Private mvarcol4BWXRTakenFindings As Collection
Private mvarcolPANOTakenFindings As Collection
Private mvarcolCompleteSeriesBWTakenFindings As Collection
Private mvarcolCephalometricFilmTakenFindings As Collection
Private mvarcolOcclusalFilmMaxFindings As Collection
Private mvarcolOcclusalFilmManFindings As Collection
Private mvarcolOtherFindings As Collection

Private mvarcolExtraoralInitFindings As Collection
Private mvarcolExtraoralAddFindings As Collection
Private mvarcolPALateralSkullFindings As Collection
Private mvarcolSialographyFindings As Collection
Private mvarcolTMJArthrogramFindings As Collection
Private mvarcolOtherTMJFindings As Collection
Private mvarcolTomographicFindings As Collection

'local variable(s) to hold property value(s)
Private mvarWLID_SingleBWXRTaken As Long 'local copy
Private mvarWLID_BWXR2Taken As Long 'local copy
Private mvarWLID_BWXR4Taken As Long 'local copy
Private mvarWLID_PATakenFirstTooth As Long 'local copy
Private mvarWLID_PATakenAdditionalTooth As Long 'local copy
Private mvarWLID_PANOTaken As Long 'local copy
Private mvarWLID_CompleteSeriesTaken As Long 'local copy
Private mvarWLID_CephalometricFilmTaken As Long 'local copy
Private mvarWLID_OcculsalFilmMax As Long 'local copy
Private mvarWLID_OcculsalFilmMan As Long 'local copy
Private mvarWLID_ExtraoralInit As Long 'local copy
Private mvarWLID_ExtraoralAdd As Long 'local copy
Private mvarWLID_PALateralSkull As Long 'local copy
Private mvarWLID_Sialography As Long 'local copy
Private mvarWLID_TMJArth As Long 'local copy
Private mvarWLID_OtherTMJ As Long 'local copy
Private mvarWLID_Tomographic As Long 'local copy

Private Type WorkLoadRecord
    WorkloadId As Long          'TX_ADA_CODE_ID
    WorkLoadDef As String       'ADA_Definition
    WorkloadCode As String      'ADA_Code
End Type

Private SingleBWXRTaken_WL As WorkLoadRecord
Private BWXR2Taken_WL As WorkLoadRecord
Private BWXR4Taken_WL As WorkLoadRecord
Private PATaken_WL As WorkLoadRecord
Private PATakenAddl_WL As WorkLoadRecord
Private PANOTaken_WL As WorkLoadRecord
Private CompleteSeriesTaken_WL As WorkLoadRecord
Private CephalometricFilmTaken_WL As WorkLoadRecord
Private OcculsalFilmMax_WL As WorkLoadRecord
Private OcculsalFilmMan_WL As WorkLoadRecord
Private ExtraoralInit_WL As WorkLoadRecord
Private ExtraoralAdd_WL As WorkLoadRecord
Private PALateralSkull_WL As WorkLoadRecord
Private Sialography_WL As WorkLoadRecord
Private TMJArth_WL As WorkLoadRecord
Private OtherTMJ_WL As WorkLoadRecord
Private Tomographic_WL As WorkLoadRecord

Public Enum WL_TypeEnum
    enumSingleBWXRTaken = 0
    enum2BWXRTaken = 1
    enum4BWXRTaken = 2
    enumPATaken = 3
    enumPATakenAddl = 4
    enumPANOTaken = 5
    enumCompleteSeriesTaken = 6
    enumCephalometricTaken = 7
    enumOcclusalFilmMax = 8
    enumOcclusalFilmMan = 9
    enumExtraoralInit = 10
    enumExtraoralAdd = 11
    enumPALateralSkull = 12
    enumSialography = 13
    enumTMJArth = 14
    enumOtherTMJ = 15
    enumTomographic = 16
    enumWorkLoadTypeMax = 17
End Enum


Private mcolRadFindings As colRadFindings

Const sSQL_Select1 As String = "NoRadTaken,PreviousRadReviewed,SingleBWXRTaken,BWXR2Taken,BWXR4Taken," _
                            & "PeriapicalRadTaken,PANOTaken,CompleteSeriesBWTaken,CephalometricFilmTaken," _
                            & "OcclusalFilmMax,OcclusalFilmMan,OtherFilms,NumberPA," _
                            & "ExtraoralInit,ExtraoralAdd,PALateralSkull,Sialography,TMJArthrogram," _
                            & "OtherTMJ,Tomographic," _
                            & "WID_SingleBWXRTaken,WID_BWXR2Taken,WID_BWXR4Taken,WID_PATakenFirstTooth," _
                            & "WID_PATakenAdditionalTooth,WID_PANOTaken,WID_CompleteSeriesTaken," _
                            & "WID_CephalometricFilmTaken,WID_OcculsalFilmMax,WID_OcculsalFilmMan," _
                            & "WID_ExtraoralInit,WID_ExtraoralAdd,WID_PALateralSkull," _
                            & "WID_Sialography,WID_TMJArth,WID_OtherTMJ,WID_Tomographic"

Const sSQL_Select2 As String = "PrevRadFindings,SingleBWXRFindings,BWXR2Findings," _
                            & "BWXR4Findings,PANOFindings,CompleteSeriesBWFindings," _
                            & "CephalometricFilmFindings"
                            
Const sSQL_Select3 As String = "OcclusalFilmMaxFindings,OcclusalFilmManFindings,ExtraoralInitFindings," _
                            & "ExtraoralAddFindings,PALateralSkullFindings,SialographyFindings," _
                            & "TMJArthrogramFindings,OtherTMJFindings,TomographicFindings"
                            
Const sSQL_Select4 As String = "RadiologyNotes"
                            
Const sSQL_Insert As String = "Dental_Encounter_ID," _
                            & "NoRadTaken,PreviousRadReviewed,SingleBWXRTaken,BWXR2Taken,BWXR4Taken," _
                            & "PeriapicalRadTaken,PANOTaken,CompleteSeriesBWTaken,CephalometricFilmTaken," _
                            & "OcclusalFilmMax,OcclusalFilmMan,OtherFilms,NumberPA," _
                            & "ExtraoralInit,ExtraoralAdd,PALateralSkull,Sialography,TMJArthrogram," _
                            & "OtherTMJ,Tomographic," _
                            & "WID_SingleBWXRTaken,WID_BWXR2Taken,WID_BWXR4Taken,WID_PATakenFirstTooth," _
                            & "WID_PATakenAdditionalTooth,WID_PANOTaken,WID_CompleteSeriesTaken," _
                            & "WID_CephalometricFilmTaken,WID_OcculsalFilmMax,WID_OcculsalFilmMan," _
                            & "WID_ExtraoralInit,WID_ExtraoralAdd,WID_PALateralSkull," _
                            & "WID_Sialography,WID_TMJArth,WID_OtherTMJ,WID_Tomographic," _
                            & "RadiologyNotes," _
                            & "PrevRadFindings,SingleBWXRFindings,BWXR2Findings," _
                            & "BWXR4Findings,PANOFindings,CompleteSeriesBWFindings," _
                            & "CephalometricFilmFindings," _
                            & "OcclusalFilmMaxFindings,OcclusalFilmManFindings,ExtraoralInitFindings," _
                            & "ExtraoralAddFindings,PALateralSkullFindings,SialographyFindings," _
                            & "TMJArthrogramFindings,OtherTMJFindings,TomographicFindings"


Public Property Set RadFindings(ByVal vData As colRadFindings)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RadNotes = 5
    Set mcolRadFindings = vData
    'UpdateDirtyFlag
End Property

Public Property Get RadFindings() As colRadFindings
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PrevRadRevFind

    Set RadFindings = mcolRadFindings
    
End Property

Public Property Let WLID_Tomographic(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WLID_Tomographic = 5
    mvarWLID_Tomographic = vData
    UpdateDirtyFlag
    
End Property


Public Property Get WLID_Tomographic() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WLID_Tomographic
    WLID_Tomographic = mvarWLID_Tomographic
End Property



Public Property Let WLID_OtherTMJ(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WLID_OtherTMJ = 5
    mvarWLID_OtherTMJ = vData
    UpdateDirtyFlag
    
End Property


Public Property Get WLID_OtherTMJ() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WLID_OtherTMJ
    WLID_OtherTMJ = mvarWLID_OtherTMJ
End Property



Public Property Let WLID_TMJArth(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WLID_TMJArth = 5
    mvarWLID_TMJArth = vData
    UpdateDirtyFlag
    
End Property


Public Property Get WLID_TMJArth() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WLID_TMJArth
    WLID_TMJArth = mvarWLID_TMJArth
End Property



Public Property Let WLID_Sialography(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WLID_Sialography = 5
    mvarWLID_Sialography = vData
    UpdateDirtyFlag
    
End Property


Public Property Get WLID_Sialography() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WLID_Sialography
    WLID_Sialography = mvarWLID_Sialography
End Property



Public Property Let WLID_PALateralSkull(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WLID_PALateralSkull = 5
    mvarWLID_PALateralSkull = vData
    UpdateDirtyFlag
    
End Property


Public Property Get WLID_PALateralSkull() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WLID_PALateralSkull
    WLID_PALateralSkull = mvarWLID_PALateralSkull
End Property



Public Property Let WLID_ExtraoralAdd(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WLID_ExtraoralAdd = 5
    mvarWLID_ExtraoralAdd = vData
    UpdateDirtyFlag
    
End Property


Public Property Get WLID_ExtraoralAdd() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WLID_ExtraoralAdd
    WLID_ExtraoralAdd = mvarWLID_ExtraoralAdd
End Property



Public Property Let WLID_ExtraoralInit(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WLID_ExtraoralInit = 5
    mvarWLID_ExtraoralInit = vData
    UpdateDirtyFlag
    
End Property


Public Property Get WLID_ExtraoralInit() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WLID_ExtraoralInit
    WLID_ExtraoralInit = mvarWLID_ExtraoralInit
End Property



Public Property Let WLID_OcculsalFilmMan(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WLID_OcculsalFilmMan = 5
    mvarWLID_OcculsalFilmMan = vData
    UpdateDirtyFlag
    
End Property


Public Property Get WLID_OcculsalFilmMan() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WLID_OcculsalFilmMan
    WLID_OcculsalFilmMan = mvarWLID_OcculsalFilmMan
End Property



Public Property Let WLID_OcculsalFilmMax(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WLID_OcculsalFilmMax = 5
    mvarWLID_OcculsalFilmMax = vData
    UpdateDirtyFlag
    
End Property


Public Property Get WLID_OcculsalFilmMax() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WLID_OcculsalFilmMax
    WLID_OcculsalFilmMax = mvarWLID_OcculsalFilmMax
End Property



Public Property Let WLID_CephalometricFilmTaken(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WLID_CephalometricFilmTaken = 5
    mvarWLID_CephalometricFilmTaken = vData
    UpdateDirtyFlag
    
End Property


Public Property Get WLID_CephalometricFilmTaken() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WLID_CephalometricFilmTaken
    WLID_CephalometricFilmTaken = mvarWLID_CephalometricFilmTaken
End Property



Public Property Let WLID_CompleteSeriesTaken(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WLID_CompleteSeriesTaken = 5
    mvarWLID_CompleteSeriesTaken = vData
    UpdateDirtyFlag
    
End Property


Public Property Get WLID_CompleteSeriesTaken() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WLID_CompleteSeriesTaken
    WLID_CompleteSeriesTaken = mvarWLID_CompleteSeriesTaken
End Property



Public Property Let WLID_PANOTaken(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WLID_PANOTaken = 5
    mvarWLID_PANOTaken = vData
    UpdateDirtyFlag
    
End Property


Public Property Get WLID_PANOTaken() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WLID_PANOTaken
    WLID_PANOTaken = mvarWLID_PANOTaken
End Property



Public Property Let WLID_PATakenAdditionalTooth(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WLID_PATakenAdditionalTooth = 5
    mvarWLID_PATakenAdditionalTooth = vData
    UpdateDirtyFlag
    
End Property


Public Property Get WLID_PATakenAdditionalTooth() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WLID_PATakenAdditionalTooth
    WLID_PATakenAdditionalTooth = mvarWLID_PATakenAdditionalTooth
End Property



Public Property Let WLID_PATakenFirstTooth(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WLID_PATakenFirstTooth = 5
    mvarWLID_PATakenFirstTooth = vData
    UpdateDirtyFlag
    
End Property


Public Property Get WLID_PATakenFirstTooth() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WLID_PATakenFirstTooth
    WLID_PATakenFirstTooth = mvarWLID_PATakenFirstTooth
End Property



Public Property Let WLID_BWXR4Taken(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WLID_BWXR4Taken = 5
    mvarWLID_BWXR4Taken = vData
    UpdateDirtyFlag
    
End Property


Public Property Get WLID_BWXR4Taken() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WLID_BWXR4Taken
    WLID_BWXR4Taken = mvarWLID_BWXR4Taken
End Property



Public Property Let WLID_BWXR2Taken(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WLID_BWXR2Taken = 5
    mvarWLID_BWXR2Taken = vData
    UpdateDirtyFlag
    
End Property


Public Property Get WLID_BWXR2Taken() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WLID_BWXR2Taken
    WLID_BWXR2Taken = mvarWLID_BWXR2Taken
End Property



Public Property Let WLID_SingleBWXRTaken(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WLID_SingleBWXRTaken = 5
    mvarWLID_SingleBWXRTaken = vData
    UpdateDirtyFlag
    
End Property


Public Property Get WLID_SingleBWXRTaken() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WLID_SingleBWXRTaken
    WLID_SingleBWXRTaken = mvarWLID_SingleBWXRTaken
End Property


                          
Public Property Let PrevRadRevFindString(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PrevRadRevFind = 5
    If mvarsPrevRadRevFindString <> vData Then
        mvarsPrevRadRevFindString = vData
        UpdateDirtyFlag
    End If
End Property


Public Property Get PrevRadRevFindString() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PrevRadRevFind

    If mvarsPrevRadRevFindString = "" Then
        mvarsPrevRadRevFindString = gcFindings
    End If

    PrevRadRevFindString = mvarsPrevRadRevFindString
    
End Property


Public Property Let SingleBWXRFindString(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PrevRadRevFind = 5
    If mvarsSingleBWXRFindString <> vData Then
        mvarsSingleBWXRFindString = vData
        UpdateDirtyFlag
    End If
End Property


Public Property Get SingleBWXRFindString() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PrevRadRevFind

    If mvarsSingleBWXRFindString = "" Then
        mvarsSingleBWXRFindString = gcFindings
    End If

    SingleBWXRFindString = mvarsSingleBWXRFindString
    
End Property


Public Property Let BWXR2FindString(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PrevRadRevFind = 5
    If mvars2BWXRFindString <> vData Then
        mvars2BWXRFindString = vData
        UpdateDirtyFlag
    End If
    
End Property


Public Property Get BWXR2FindString() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PrevRadRevFind

    If mvars2BWXRFindString = "" Then
        mvars2BWXRFindString = gcFindings
    End If

    BWXR2FindString = mvars2BWXRFindString
    
End Property

Public Property Let BWXR4FindString(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PrevRadRevFind = 5
    mvars4BWXRFindString = vData
    UpdateDirtyFlag
End Property


Public Property Get BWXR4FindString() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PrevRadRevFind

    If mvars4BWXRFindString = "" Then
        mvars4BWXRFindString = gcFindings
    End If

    BWXR4FindString = mvars4BWXRFindString
    
End Property

Public Property Let PANOFindString(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PrevRadRevFind = 5
    mvarsPANOFindString = vData
    UpdateDirtyFlag
End Property


Public Property Get PANOFindString() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PrevRadRevFind

    If mvarsPANOFindString = "" Then
        mvarsPANOFindString = gcFindings
    End If

    PANOFindString = mvarsPANOFindString
    
End Property


Public Property Let CompleteSeriesFindString(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PrevRadRevFind = 5
    mvarsCompleteSeriesFindString = vData
    UpdateDirtyFlag
End Property


Public Property Get CompleteSeriesFindString() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PrevRadRevFind

    If mvarsCompleteSeriesFindString = "" Then
        mvarsCompleteSeriesFindString = gcFindings
    End If

    CompleteSeriesFindString = mvarsCompleteSeriesFindString
    
End Property


Public Property Let CephalometricFilmFindString(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PrevRadRevFind = 5
    mvarsCephalometricFilmFindString = vData
    UpdateDirtyFlag
End Property


Public Property Get CephalometricFilmFindString() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PrevRadRevFind

    If mvarsCephalometricFilmFindString = "" Then
        mvarsCephalometricFilmFindString = gcFindings
    End If

    CephalometricFilmFindString = mvarsCephalometricFilmFindString
    
End Property


Public Property Let OcclusalFilmMaxFindString(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PrevRadRevFind = 5
    mvarsOcclusalFilmMaxFindString = vData
    UpdateDirtyFlag
End Property


Public Property Get OcclusalFilmMaxFindString() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PrevRadRevFind

    If mvarsOcclusalFilmMaxFindString = "" Then
        mvarsOcclusalFilmMaxFindString = gcFindings
    End If

    OcclusalFilmMaxFindString = mvarsOcclusalFilmMaxFindString
    
End Property

Public Property Let OcclusalFilmManFindString(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PrevRadRevFind = 5
    mvarsOcclusalFilmManFindString = vData
    UpdateDirtyFlag
End Property


Public Property Get OcclusalFilmManFindString() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PrevRadRevFind

    If mvarsOcclusalFilmManFindString = "" Then
        mvarsOcclusalFilmManFindString = gcFindings
    End If

    OcclusalFilmManFindString = mvarsOcclusalFilmManFindString
    
End Property


Public Property Let OtherFindString(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PrevRadRevFind = 5
    mvarsOtherFindString = vData
    UpdateDirtyFlag
End Property


Public Property Get OtherFindString() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PrevRadRevFind

    If mvarsOtherFindString = "" Then
        mvarsOtherFindString = gcFindings
    End If

    OtherFindString = mvarsOtherFindString
    
End Property


Public Property Let ExtraoralInitFindString(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PrevRadRevFind = 5
    mvarsExtraoralInitFindString = vData
    UpdateDirtyFlag
End Property


Public Property Get ExtraoralInitFindString() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PrevRadRevFind

    If mvarsExtraoralInitFindString = "" Then
        mvarsExtraoralInitFindString = gcFindings
    End If

    ExtraoralInitFindString = mvarsExtraoralInitFindString
    
End Property


Public Property Let ExtraoralAddFindString(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PrevRadRevFind = 5
    mvarsExtraoralAddFindString = vData
    UpdateDirtyFlag
End Property


Public Property Get ExtraoralAddFindString() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PrevRadRevFind

    If mvarsExtraoralAddFindString = "" Then
        mvarsExtraoralAddFindString = gcFindings
    End If

    ExtraoralAddFindString = mvarsExtraoralAddFindString
    
End Property


Public Property Let PALateralSkullFindString(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PrevRadRevFind = 5
    mvarsPALateralSkullFindString = vData
    UpdateDirtyFlag
End Property


Public Property Get PALateralSkullFindString() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PrevRadRevFind

    If mvarsPALateralSkullFindString = "" Then
        mvarsPALateralSkullFindString = gcFindings
    End If

    PALateralSkullFindString = mvarsPALateralSkullFindString
    
End Property

Public Property Let SialographyFindString(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PrevRadRevFind = 5
    mvarsSialographyFindString = vData
    UpdateDirtyFlag
End Property


Public Property Get SialographyFindString() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PrevRadRevFind

    If mvarsSialographyFindString = "" Then
        mvarsSialographyFindString = gcFindings
    End If

    SialographyFindString = mvarsSialographyFindString
    
End Property

Public Property Let TMJArthrogramFindString(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PrevRadRevFind = 5
    mvarsTMJArthrogramFindString = vData
    UpdateDirtyFlag
End Property


Public Property Get TMJArthrogramFindString() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PrevRadRevFind

    If mvarsTMJArthrogramFindString = "" Then
        mvarsTMJArthrogramFindString = gcFindings
    End If

    TMJArthrogramFindString = mvarsTMJArthrogramFindString
    
End Property

Public Property Let OtherTMJFindString(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PrevRadRevFind = 5
    mvarsOtherTMJFindString = vData
    UpdateDirtyFlag
End Property


Public Property Get OtherTMJFindString() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PrevRadRevFind

    If mvarsOtherTMJFindString = "" Then
        mvarsOtherTMJFindString = gcFindings
    End If

    OtherTMJFindString = mvarsOtherTMJFindString
    
End Property


Public Property Let TomographicFindString(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PrevRadRevFind = 5
    mvarsTomographicFindString = vData
    UpdateDirtyFlag
End Property


Public Property Get TomographicFindString() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PrevRadRevFind

    If mvarsTomographicFindString = "" Then
        mvarsTomographicFindString = gcFindings
    End If

    TomographicFindString = mvarsTomographicFindString
    
End Property

Public Property Let RadNotes(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RadNotes = 5
    mvarsRadNotes = vData
    UpdateDirtyFlag
End Property


Public Property Get RadNotes() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RadNotes
    RadNotes = mvarsRadNotes
End Property

Public Property Let NoRadTaken(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NoRadTaken = 5
    mvarbNoRadTaken = vData
    UpdateDirtyFlag
End Property


Public Property Get NoRadTaken() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NoRadTaken
    NoRadTaken = mvarbNoRadTaken
End Property


Public Property Let PrevRadRev(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NoRadTaken = 5
    mvarbPrevRadRev = vData
    UpdateDirtyFlag
End Property


Public Property Get PrevRadRev() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NoRadTaken
    PrevRadRev = mvarbPrevRadRev
End Property

Public Property Let SingleBWXRTaken(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NoRadTaken = 5
    mvarbSingleBWXRTaken = vData
    UpdateDirtyFlag
End Property


Public Property Get SingleBWXRTaken() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NoRadTaken
    SingleBWXRTaken = mvarbSingleBWXRTaken
End Property

Public Property Let BWXRTaken2(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NoRadTaken = 5
    mvarb2BWXRTaken = vData
    UpdateDirtyFlag
End Property

Public Property Get BWXRTaken2() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NoRadTaken
    BWXRTaken2 = mvarb2BWXRTaken
End Property


Public Property Let BWXRTaken4(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NoRadTaken = 5
    mvarb4BWXRTaken = vData
    UpdateDirtyFlag
End Property

Public Property Get BWXRTaken4() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NoRadTaken
    BWXRTaken4 = mvarb4BWXRTaken
End Property

Public Property Let PeriapicalRadTaken(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NoRadTaken = 5
    If mvarbPeriapicalRadTaken <> vData Then
        mvarbPeriapicalRadTaken = vData
        UpdateDirtyFlag
    End If
End Property

Public Property Get PeriapicalRadTaken() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NoRadTaken
    PeriapicalRadTaken = mvarbPeriapicalRadTaken
End Property


Public Property Let PANOTaken(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NoRadTaken = 5
    mvarbPANOTaken = vData
    UpdateDirtyFlag
End Property

Public Property Get PANOTaken() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NoRadTaken
    PANOTaken = mvarbPANOTaken
End Property

Public Property Let CompleteSeriesBWTaken(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NoRadTaken = 5
    mvarbCompleteSeriesBWTaken = vData
    UpdateDirtyFlag
End Property

Public Property Get CompleteSeriesBWTaken() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NoRadTaken
    CompleteSeriesBWTaken = mvarbCompleteSeriesBWTaken
End Property

Public Property Let CephalometricFilmTaken(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NoRadTaken = 5
    mvarbCephalometricFilmTaken = vData
    UpdateDirtyFlag
End Property

Public Property Get CephalometricFilmTaken() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NoRadTaken
    CephalometricFilmTaken = mvarbCephalometricFilmTaken
End Property

Public Property Let OcclusalFilmMax(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NoRadTaken = 5
    mvarbOcclusalFilmMax = vData
    UpdateDirtyFlag
End Property

Public Property Get OcclusalFilmMax() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NoRadTaken
    OcclusalFilmMax = mvarbOcclusalFilmMax
End Property

Public Property Let OcclusalFilmMan(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NoRadTaken = 5
    mvarbOcclusalFilmMan = vData
    UpdateDirtyFlag
End Property

Public Property Get OcclusalFilmMan() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NoRadTaken
    OcclusalFilmMan = mvarbOcclusalFilmMan
End Property

Public Property Let OtherFilms(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NoRadTaken = 5
    mvarbOtherFilms = vData
    UpdateDirtyFlag
End Property

Public Property Get OtherFilms() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NoRadTaken
    OtherFilms = mvarbOtherFilms
End Property


Public Property Let OtherFilmsComboBox(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NoRadTaken = 5
    mvarsOtherFilmsComboBox = vData
    'UpdateDirtyFlag
End Property

Public Property Get OtherFilmsComboBox() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NoRadTaken
    OtherFilmsComboBox = mvarsOtherFilmsComboBox
End Property

Public Property Let NumberPA(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NoRadTaken = 5
    If mvariNumberPA <> vData Then
        mvariNumberPA = vData
        UpdateDirtyFlag
    End If
    
End Property

Public Property Get NumberPA() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NoRadTaken
    NumberPA = mvariNumberPA
End Property

Public Property Let ExtraoralInit(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NoRadTaken = 5
    mvarbExtraoralInit = vData
    UpdateDirtyFlag
End Property


Public Property Get ExtraoralInit() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NoRadTaken
    ExtraoralInit = mvarbExtraoralInit
End Property

Public Property Let ExtraoralAdd(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NoRadTaken = 5
    mvarbExtraoralAdd = vData
    UpdateDirtyFlag
End Property


Public Property Get ExtraoralAdd() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NoRadTaken
    ExtraoralAdd = mvarbExtraoralAdd
End Property

Public Property Let PALateralSkull(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NoRadTaken = 5
    mvarbPALateralSkull = vData
    UpdateDirtyFlag
End Property


Public Property Get PALateralSkull() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NoRadTaken
    PALateralSkull = mvarbPALateralSkull
End Property

Public Property Let Sialography(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NoRadTaken = 5
    mvarbSialography = vData
    UpdateDirtyFlag
End Property


Public Property Get Sialography() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NoRadTaken
    Sialography = mvarbSialography
End Property

Public Property Let TMJArthrogram(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NoRadTaken = 5
    mvarbTMJArthrogram = vData
    UpdateDirtyFlag
End Property


Public Property Get TMJArthrogram() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NoRadTaken
    TMJArthrogram = mvarbTMJArthrogram
End Property

Public Property Let OtherTMJ(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NoRadTaken = 5
    mvarbOtherTMJ = vData
    UpdateDirtyFlag
End Property


Public Property Get OtherTMJ() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NoRadTaken
    OtherTMJ = mvarbOtherTMJ
End Property

Public Property Let Tomographic(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NoRadTaken = 5
    mvarbTomographic = vData
    UpdateDirtyFlag
End Property


Public Property Get Tomographic() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NoRadTaken
    Tomographic = mvarbTomographic
End Property

Public Property Get colPrevRadRevFindings() As Collection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PrevRadRevFind

    Set colPrevRadRevFindings = mvarcolPrevRadRevFindings
    
End Property


Public Property Set colPrevRadRevFindings(ByVal vData As Collection)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RadNotes = 5
    Set mvarcolPrevRadRevFindings = vData
    'UpdateDirtyFlag
End Property

Public Property Get colSingleBWXRTakenFindings() As Collection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PrevRadRevFind

    Set colSingleBWXRTakenFindings = mvarcolSingleBWXRTakenFindings
    
End Property


Public Property Set colSingleBWXRTakenFindings(ByVal vData As Collection)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RadNotes = 5
    Set mvarcolSingleBWXRTakenFindings = vData
    'UpdateDirtyFlag
End Property

Public Property Get col2BWXRTakenFindings() As Collection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PrevRadRevFind

    Set col2BWXRTakenFindings = mvarcol2BWXRTakenFindings
    
End Property

Public Property Set col2BWXRTakenFindings(ByVal vData As Collection)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RadNotes = 5
    Set mvarcol2BWXRTakenFindings = vData
End Property

Public Property Get col4BWXRTakenFindings() As Collection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PrevRadRevFind

    Set col4BWXRTakenFindings = mvarcol4BWXRTakenFindings
    
End Property

Public Property Set col4BWXRTakenFindings(ByVal vData As Collection)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RadNotes = 5
    Set mvarcol4BWXRTakenFindings = vData
End Property

Public Property Get colCephalometricFilmTakenFindings() As Collection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PrevRadRevFind

    Set colCephalometricFilmTakenFindings = mvarcolCephalometricFilmTakenFindings
    
End Property

Public Property Set colCephalometricFilmTakenFindings(ByVal vData As Collection)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RadNotes = 5
    Set mvarcolCephalometricFilmTakenFindings = vData
End Property
Public Property Get colPANOTakenFindings() As Collection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PrevRadRevFind

    Set colPANOTakenFindings = mvarcolPANOTakenFindings
    
End Property

Public Property Set colPANOTakenFindings(ByVal vData As Collection)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RadNotes = 5
    Set mvarcolPANOTakenFindings = vData
End Property

Public Property Get colCompleteSeriesBWTakenFindings() As Collection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PrevRadRevFind

    Set colCompleteSeriesBWTakenFindings = mvarcolCompleteSeriesBWTakenFindings
    
End Property

Public Property Set colCompleteSeriesBWTakenFindings(ByVal vData As Collection)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RadNotes = 5
    Set mvarcolCompleteSeriesBWTakenFindings = vData
End Property

Public Property Get colOcclusalFilmMaxFindings() As Collection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PrevRadRevFind

    Set colOcclusalFilmMaxFindings = mvarcolOcclusalFilmMaxFindings
    
End Property

Public Property Set colOcclusalFilmMaxFindings(ByVal vData As Collection)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RadNotes = 5
    Set mvarcolOcclusalFilmMaxFindings = vData
End Property

Public Property Get colOcclusalFilmManFindings() As Collection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PrevRadRevFind

    Set colOcclusalFilmManFindings = mvarcolOcclusalFilmManFindings
    
End Property

Public Property Set colOcclusalFilmManFindings(ByVal vData As Collection)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RadNotes = 5
    Set mvarcolOcclusalFilmManFindings = vData
End Property

Public Property Get colOtherFindings() As Collection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PrevRadRevFind

    Set colOtherFindings = mvarcolOtherFindings
    
End Property

Public Property Set colOtherFindings(ByVal vData As Collection)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RadNotes = 5
    Set mvarcolOtherFindings = vData
End Property


Public Property Get colExtraoralInitFindings() As Collection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PrevRadRevFind

    Set colExtraoralInitFindings = mvarcolExtraoralInitFindings
    
End Property

Public Property Set colExtraoralInitFindings(ByVal vData As Collection)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RadNotes = 5
    Set mvarcolExtraoralInitFindings = vData
End Property

Public Property Get colExtraoralAddFindings() As Collection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PrevRadRevFind

    Set colExtraoralAddFindings = mvarcolExtraoralAddFindings
    
End Property

Public Property Set colExtraoralAddFindings(ByVal vData As Collection)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RadNotes = 5
    Set mvarcolExtraoralAddFindings = vData
End Property


Public Property Get colPALateralSkullFindings() As Collection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PrevRadRevFind

    Set colPALateralSkullFindings = mvarcolPALateralSkullFindings
    
End Property

Public Property Set colPALateralSkullFindings(ByVal vData As Collection)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RadNotes = 5
    Set mvarcolPALateralSkullFindings = vData
End Property


Public Property Get colSialographyFindings() As Collection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PrevRadRevFind

    Set colSialographyFindings = mvarcolSialographyFindings
    
End Property

Public Property Set colSialographyFindings(ByVal vData As Collection)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RadNotes = 5
    Set mvarcolSialographyFindings = vData
End Property


Public Property Get colTMJArthrogramFindings() As Collection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PrevRadRevFind

    Set colTMJArthrogramFindings = mvarcolTMJArthrogramFindings
    
End Property

Public Property Set colTMJArthrogramFindings(ByVal vData As Collection)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RadNotes = 5
    Set mvarcolTMJArthrogramFindings = vData
End Property


Public Property Get colOtherTMJFindings() As Collection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PrevRadRevFind

    Set colOtherTMJFindings = mvarcolOtherTMJFindings
    
End Property

Public Property Set colOtherTMJFindings(ByVal vData As Collection)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RadNotes = 5
    Set mvarcolOtherTMJFindings = vData
End Property


Public Property Get colTomographicFindings() As Collection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PrevRadRevFind

    Set colTomographicFindings = mvarcolTomographicFindings
    
End Property

Public Property Set colTomographicFindings(ByVal vData As Collection)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RadNotes = 5
    Set mvarcolTomographicFindings = vData
End Property

Private Function UpdateDirtyFlag()

    If dfDirtyFlag = dfNoAction Then
        dfDirtyFlag = dfUpdate
    End If
    
End Function


Private Sub Class_Initialize()

    dfDirtyFlag = dfNoAction
    
    mvarsPrevRadRevFindString = gcFindings
    mvarsSingleBWXRFindString = gcFindings
    mvars2BWXRFindString = gcFindings
    mvars4BWXRFindString = gcFindings
    mvarsPANOFindString = gcFindings
    mvarsCompleteSeriesFindString = gcFindings
    mvarsCephalometricFilmFindString = gcFindings
    mvarsOcclusalFilmMaxFindString = gcFindings
    mvarsOcclusalFilmManFindString = gcFindings
    mvarsOtherFindString = gcFindings
    
    mvarsExtraoralInitFindString = gcFindings
    mvarsExtraoralAddFindString = gcFindings
    mvarsPALateralSkullFindString = gcFindings
    mvarsSialographyFindString = gcFindings
    mvarsTMJArthrogramFindString = gcFindings
    mvarsOtherTMJFindString = gcFindings
    mvarsTomographicFindString = gcFindings
    
    SingleBWXRTaken_WL.WorkloadId = 48
    SingleBWXRTaken_WL.WorkloadCode = "D0270"
    SingleBWXRTaken_WL.WorkLoadDef = "Single BWX"

    BWXR2Taken_WL.WorkloadId = 50
    BWXR2Taken_WL.WorkloadCode = "D0272"
    BWXR2Taken_WL.WorkLoadDef = "BWX x2"
    
    BWXR4Taken_WL.WorkloadId = 49
    BWXR4Taken_WL.WorkloadCode = "D0274"
    BWXR4Taken_WL.WorkLoadDef = "BWX x4"
    
    PATaken_WL.WorkloadId = 204
    PATaken_WL.WorkloadCode = "D0220"
    PATaken_WL.WorkLoadDef = "PAX"
    
    PATakenAddl_WL.WorkloadId = 203
    PATakenAddl_WL.WorkloadCode = "D0230"
    PATakenAddl_WL.WorkLoadDef = "PAX additional"
    
    PANOTaken_WL.WorkloadId = 308
    PANOTaken_WL.WorkloadCode = "D0330"
    PANOTaken_WL.WorkLoadDef = "Panoramic film"
    
    CompleteSeriesTaken_WL.WorkloadId = 201
    CompleteSeriesTaken_WL.WorkloadCode = "D0210"
    CompleteSeriesTaken_WL.WorkLoadDef = "Intraoral - complete x-ray series"
    
    CephalometricFilmTaken_WL.WorkloadId = 59
    CephalometricFilmTaken_WL.WorkloadCode = "D0340"
    CephalometricFilmTaken_WL.WorkLoadDef = "Cephalometric film"
    
    OcculsalFilmMax_WL.WorkloadId = 202
    OcculsalFilmMax_WL.WorkloadCode = "D0240"
    OcculsalFilmMax_WL.WorkLoadDef = "Occlusal radiograph"
    
    OcculsalFilmMan_WL.WorkloadId = 202
    OcculsalFilmMan_WL.WorkloadCode = "D0240"
    OcculsalFilmMan_WL.WorkLoadDef = "Occlusal radiograph"
    
    ExtraoralInit_WL.WorkloadId = 135
    ExtraoralInit_WL.WorkloadCode = "D0250"
    ExtraoralInit_WL.WorkLoadDef = "Extraoral radiograph"
    
    ExtraoralAdd_WL.WorkloadId = 134
    ExtraoralAdd_WL.WorkloadCode = "D0260"
    ExtraoralAdd_WL.WorkLoadDef = "Extraoral radiograph, additional"
    
    PALateralSkull_WL.WorkloadId = 328
    PALateralSkull_WL.WorkloadCode = "D0290"
    PALateralSkull_WL.WorkLoadDef = "PA or Lat skull"
    
    Sialography_WL.WorkloadId = 419
    Sialography_WL.WorkloadCode = "D0310"
    Sialography_WL.WorkLoadDef = "Sialography"
    
    TMJArth_WL.WorkloadId = 458
    TMJArth_WL.WorkloadCode = "D0320"
    TMJArth_WL.WorkLoadDef = "TMJ arthogram"
    
    OtherTMJ_WL.WorkloadId = 299
    OtherTMJ_WL.WorkloadCode = "D0321"
    OtherTMJ_WL.WorkLoadDef = "TMJ film"
    
    Tomographic_WL.WorkloadId = 468
    Tomographic_WL.WorkloadCode = "D0322"
    Tomographic_WL.WorkLoadDef = "Tomographic survey"
    
End Sub

Private Sub Class_Terminate()
    Set mcolRadFindings = Nothing
End Sub

Private Property Get clsISection_DentalEncounterID() As Long

    clsISection_DentalEncounterID = mvarlDentalEncounterID

End Property

Private Function clsISection_GetNote() As String

Dim sTemp As String
Dim sStringDisplay As String
Dim oRadFindItem As clsRadFindItem
    
    If mvarbNoRadTaken = True Then
        sTemp = sTemp & vbTab & "No radiographs taken." & vbCrLf
    End If
       
    If mvarbPrevRadRev = True Then
        sTemp = sTemp & vbTab & "Previous radiographs reviewed."
    End If
       
    sStringDisplay = FindingsDisplay(mvarcolPrevRadRevFindings)
    If sStringDisplay <> "" Then
        sTemp = sTemp & " Findings: " & sStringDisplay & vbCrLf
    End If
    
    If mvarbSingleBWXRTaken = True Then
        sTemp = sTemp & vbTab & "Single BWXR taken."
    End If
       
    sStringDisplay = FindingsDisplay(mvarcolSingleBWXRTakenFindings)
    If sStringDisplay <> "" Then
        sTemp = sTemp & " Findings: " & sStringDisplay & vbCrLf
    End If
           
    If mvarb2BWXRTaken = True Then
        sTemp = sTemp & vbTab & "2 BWXR taken."
    End If
       
    sStringDisplay = FindingsDisplay(mvarcol2BWXRTakenFindings)
    If sStringDisplay <> "" Then
        sTemp = sTemp & " Findings: " & sStringDisplay & vbCrLf
    End If
        
    If mvarb4BWXRTaken = True Then
        sTemp = sTemp & vbTab & "4 BWXR taken."
    End If
       
    sStringDisplay = FindingsDisplay(mvarcol4BWXRTakenFindings)
    If sStringDisplay <> "" Then
        sTemp = sTemp & " Findings: " & sStringDisplay & vbCrLf
    End If
    
    If mvarbPeriapicalRadTaken = True Then
        sTemp = sTemp & vbTab & "Periapical radiograph taken. "
    End If
    
    If mvariNumberPA <> 0 Then
        sTemp = sTemp & mvariNumberPA & " Number of PA's taken." & vbCrLf
    End If
    
    If mvarbPANOTaken = True Then
        sTemp = sTemp & vbTab & "PANO taken."
    End If
       
    sStringDisplay = FindingsDisplay(mvarcolPANOTakenFindings)
    If sStringDisplay <> "" Then
        sTemp = sTemp & " Findings: " & sStringDisplay & vbCrLf
    End If
    
    If mvarbCompleteSeriesBWTaken = True Then
        sTemp = sTemp & vbTab & "Complete series with BW taken."
    End If
       
    sStringDisplay = FindingsDisplay(mvarcolCompleteSeriesBWTakenFindings)
    If sStringDisplay <> "" Then
        sTemp = sTemp & " Findings: " & sStringDisplay & vbCrLf
    End If
   
    If mvarbCephalometricFilmTaken = True Then
        sTemp = sTemp & vbTab & "Cephalometric film taken."
    End If
       
    sStringDisplay = FindingsDisplay(mvarcolCephalometricFilmTakenFindings)
    If sStringDisplay <> "" Then
        sTemp = sTemp & " Findings: " & sStringDisplay & vbCrLf
    End If
    
    If mvarbOcclusalFilmMax = True Then
        sTemp = sTemp & vbTab & "Occlusal film-max."
    End If
       
    sStringDisplay = FindingsDisplay(mvarcolOcclusalFilmMaxFindings)
    If sStringDisplay <> "" Then
        sTemp = sTemp & " Findings: " & sStringDisplay & vbCrLf
    End If
    
    If mvarbOcclusalFilmMan = True Then
        sTemp = sTemp & vbTab & "Occlusal film-man."
    End If
       
    sStringDisplay = FindingsDisplay(mvarcolOcclusalFilmManFindings)
    If sStringDisplay <> "" Then
        sTemp = sTemp & " Findings: " & sStringDisplay & vbCrLf
    End If
        
    If mvarbOtherFilms = True Then
        sTemp = sTemp & vbTab & "Other films: " & vbCrLf
    End If
          
    If mvarbExtraoralInit = True Then
        sTemp = sTemp & vbTab & vbTab & "Extraoral, initial. "
    End If
     
    sStringDisplay = FindingsDisplay(mvarcolExtraoralInitFindings)
    If sStringDisplay <> "" Then
        sTemp = sTemp & " Findings: " & sStringDisplay & vbCrLf
    End If
     
    If mvarbExtraoralAdd = True Then
        sTemp = sTemp & vbTab & vbTab & "Extraoral, additional. "
    End If
     
    sStringDisplay = FindingsDisplay(mvarcolExtraoralAddFindings)
    If sStringDisplay <> "" Then
        sTemp = sTemp & " Findings: " & sStringDisplay & vbCrLf
    End If
     
    If mvarbPALateralSkull = True Then
        sTemp = sTemp & vbTab & vbTab & "PA, lateral skull. "
    End If
     
    sStringDisplay = FindingsDisplay(mvarcolPALateralSkullFindings)
    If sStringDisplay <> "" Then
        sTemp = sTemp & " Findings: " & sStringDisplay & vbCrLf
    End If
          
     If mvarbSialography = True Then
        sTemp = sTemp & vbTab & vbTab & "Sialography. "
    End If
     
    sStringDisplay = FindingsDisplay(mvarcolSialographyFindings)
    If sStringDisplay <> "" Then
        sTemp = sTemp & " Findings: " & sStringDisplay & vbCrLf
    End If
         
    If mvarbTMJArthrogram = True Then
        sTemp = sTemp & vbTab & vbTab & "TMJ arthrogram. "
    End If
     
    sStringDisplay = FindingsDisplay(mvarcolTMJArthrogramFindings)
    If sStringDisplay <> "" Then
        sTemp = sTemp & " Findings: " & sStringDisplay & vbCrLf
    End If
          
    If mvarbOtherTMJ = True Then
        sTemp = sTemp & vbTab & vbTab & "Other TMJ. "
    End If
     
    sStringDisplay = FindingsDisplay(mvarcolOtherTMJFindings)
    If sStringDisplay <> "" Then
        sTemp = sTemp & " Findings: " & sStringDisplay & vbCrLf
    End If
          
    If mvarbTomographic = True Then
        sTemp = sTemp & vbTab & vbTab & "Tomographic. "
    End If
     
    sStringDisplay = FindingsDisplay(mvarcolTomographicFindings)
    If sStringDisplay <> "" Then
        sTemp = sTemp & " Findings: " & sStringDisplay & vbCrLf
    End If
                         
    sTemp = sTemp & mcolRadFindings.GetNote
        
    If mvarsRadNotes <> "" Then
        sTemp = sTemp & vbTab & "Comments: " & mvarsRadNotes & vbCrLf
    End If
        
    If sTemp <> "" Then
        sTemp = vbCrLf & "Radiology:" & vbCrLf & sTemp
    End If
        
    clsISection_GetNote = sTemp
    
End Function

Private Function clsISection_GetRTFNote() As String

Dim sTemp As String
Dim sStringDisplay As String
Dim oRadFindItem As clsRadFindItem
    
    If mvarbNoRadTaken = True Then
        sTemp = sTemp & "; " & "No radiographs taken."
    End If
       
    If mvarbPrevRadRev = True Then
        sTemp = sTemp & "; " & "Previous radiographs reviewed."
    End If
       
    sStringDisplay = FindingsDisplay(mvarcolPrevRadRevFindings)
    If sStringDisplay <> "" Then
        sTemp = sTemp & "; " & "Findings: " & sStringDisplay
    End If
    
    If mvarbSingleBWXRTaken = True Then
        sTemp = sTemp & "; " & "Single BWXR taken."
    End If
       
    sStringDisplay = FindingsDisplay(mvarcolSingleBWXRTakenFindings)
    If sStringDisplay <> "" Then
        sTemp = sTemp & "; " & "Findings: " & sStringDisplay
    End If
           
    If mvarb2BWXRTaken = True Then
        sTemp = sTemp & "; " & "2 BWXR taken."
    End If
       
    sStringDisplay = FindingsDisplay(mvarcol2BWXRTakenFindings)
    If sStringDisplay <> "" Then
        sTemp = sTemp & "; " & "Findings: " & sStringDisplay
    End If
        
    If mvarb4BWXRTaken = True Then
        sTemp = sTemp & "; " & "4 BWXR taken."
    End If
       
    sStringDisplay = FindingsDisplay(mvarcol4BWXRTakenFindings)
    If sStringDisplay <> "" Then
        sTemp = sTemp & "; " & "Findings: " & sStringDisplay
    End If
    
    If mvarbPeriapicalRadTaken = True Then
        sTemp = sTemp & "; " & "Periapical radiograph taken."
    End If
    
    If mvariNumberPA <> 0 Then
        sTemp = sTemp & "; " & mvariNumberPA & " Number of PA's taken."
    End If
    
    If mvarbPANOTaken = True Then
        sTemp = sTemp & "; " & "PANO taken."
    End If
       
    sStringDisplay = FindingsDisplay(mvarcolPANOTakenFindings)
    If sStringDisplay <> "" Then
        sTemp = sTemp & "; " & "Findings: " & sStringDisplay
    End If
    
    If mvarbCompleteSeriesBWTaken = True Then
        sTemp = sTemp & "; " & "Complete series with BW taken."
    End If
       
    sStringDisplay = FindingsDisplay(mvarcolCompleteSeriesBWTakenFindings)
    If sStringDisplay <> "" Then
        sTemp = sTemp & "; " & "Findings: " & sStringDisplay
    End If
   
    If mvarbCephalometricFilmTaken = True Then
        sTemp = sTemp & "; " & "Cephalometric film taken."
    End If
       
    sStringDisplay = FindingsDisplay(mvarcolCephalometricFilmTakenFindings)
    If sStringDisplay <> "" Then
        sTemp = sTemp & "; " & "Findings: " & sStringDisplay
    End If
    
    If mvarbOcclusalFilmMax = True Then
        sTemp = sTemp & "; " & "Occlusal film-max."
    End If
       
    sStringDisplay = FindingsDisplay(mvarcolOcclusalFilmMaxFindings)
    If sStringDisplay <> "" Then
        sTemp = sTemp & "; " & "Findings: " & sStringDisplay
    End If
    
    If mvarbOcclusalFilmMan = True Then
        sTemp = sTemp & "; " & "Occlusal film-man."
    End If
       
    sStringDisplay = FindingsDisplay(mvarcolOcclusalFilmManFindings)
    If sStringDisplay <> "" Then
        sTemp = sTemp & "; " & "Findings: " & sStringDisplay
    End If
        
    If mvarbOtherFilms = True Then
        sTemp = sTemp & "; " & "Other films: "
    End If
          
    If mvarbExtraoralInit = True Then
        sTemp = sTemp & "; " & "Extraoral, initial. "
    End If
     
    sStringDisplay = FindingsDisplay(mvarcolExtraoralInitFindings)
    If sStringDisplay <> "" Then
        sTemp = sTemp & "; " & "Findings: " & sStringDisplay
    End If
     
    If mvarbExtraoralAdd = True Then
        sTemp = sTemp & "; " & "Extraoral, additional. "
    End If
     
    sStringDisplay = FindingsDisplay(mvarcolExtraoralAddFindings)
    If sStringDisplay <> "" Then
        sTemp = sTemp & "; " & "Findings: " & sStringDisplay
    End If
     
    If mvarbPALateralSkull = True Then
        sTemp = sTemp & "; " & "PA, lateral skull. "
    End If
     
    sStringDisplay = FindingsDisplay(mvarcolPALateralSkullFindings)
    If sStringDisplay <> "" Then
        sTemp = sTemp & "; " & "Findings: " & sStringDisplay
    End If
          
     If mvarbSialography = True Then
        sTemp = sTemp & "; " & "Sialography. "
    End If
     
    sStringDisplay = FindingsDisplay(mvarcolSialographyFindings)
    If sStringDisplay <> "" Then
        sTemp = sTemp & "; " & "Findings: " & sStringDisplay
    End If
         
    If mvarbTMJArthrogram = True Then
        sTemp = sTemp & "; " & "TMJ arthrogram. "
    End If
     
    sStringDisplay = FindingsDisplay(mvarcolTMJArthrogramFindings)
    If sStringDisplay <> "" Then
        sTemp = sTemp & "; " & "Findings: " & sStringDisplay
    End If
          
    If mvarbOtherTMJ = True Then
        sTemp = sTemp & "; " & "Other TMJ. "
    End If
     
    sStringDisplay = FindingsDisplay(mvarcolOtherTMJFindings)
    If sStringDisplay <> "" Then
        sTemp = sTemp & "; " & "Findings: " & sStringDisplay
    End If
          
    If mvarbTomographic = True Then
        sTemp = sTemp & "; " & "Tomographic. "
    End If
     
    sStringDisplay = FindingsDisplay(mvarcolTomographicFindings)
    If sStringDisplay <> "" Then
        sTemp = sTemp & "; " & "Findings: " & sStringDisplay
    End If
                         
    sTemp = sTemp & mcolRadFindings.GetNote
        
    If mvarsRadNotes <> "" Then
        sTemp = sTemp & "; " & "Comments: " & mvarsRadNotes
    End If
        
    If sTemp <> "" Then
        'Remove the first "; "
        sTemp = Mid(sTemp, Len("; ") + 1)
        'sTemp = crtfStartControlName & "Radiology:" & sTemp & crtfEndControlName
        sTemp = crtfStartControlName & sTemp & crtfEndControlName
    End If
        
    clsISection_GetRTFNote = sTemp
    
End Function

Private Function clsISection_LoadSection(ByVal PatientID As Long, ByVal DentalEncounterID As Long) As Boolean
Dim sSQL As String, rs As Recordset, rsPrev As Recordset
Const sPROC_NAME = "LoadSection"

On Error GoTo ErrHnd

    mvarlDentalEncounterID = DentalEncounterID
    
    'Load Radiology flex grid data
    'to determine whether data exist in the grid
    'to meet PA Taken requirements
    'this textbox should be pre-populated with number of rows
    'from the grid, if any, and should be disabled
    'otherwise it contains the value entered by the user
    Set mcolRadFindings = New colRadFindings
    mcolRadFindings.LoadSection mvarlPatientID, mvarlDentalEncounterID
    
    'Load the values
    sSQL = "Select "
    sSQL = sSQL & sSQL_Select1
    sSQL = sSQL & " From TEMPL_RADIOLOGY "
    sSQL = sSQL & " Where DENTAL_ENCOUNTER_ID = " & mvarlDentalEncounterID

    Set rs = goconn.OpenRecordset(sSQL)
    
    If rs.EOF Then
        dfDirtyFlag = dfAdd
        'Init the values here to ""
                
        mvarbNoRadTaken = False
        mvarbPrevRadRev = False
        mvarbSingleBWXRTaken = False
        mvarb2BWXRTaken = False
        mvarb4BWXRTaken = False
        mvarbPeriapicalRadTaken = False
        mvarbPANOTaken = False
        mvarbCompleteSeriesBWTaken = False
        mvarbCephalometricFilmTaken = False
        mvarbOcclusalFilmMax = False
        mvarbOcclusalFilmMan = False
        mvarbOtherFilms = False

        mvariNumberPA = 0

        mvarbExtraoralInit = False
        mvarbExtraoralAdd = False
        mvarbPALateralSkull = False
        mvarbSialography = False
        mvarbTMJArthrogram = False
        mvarbOtherTMJ = False
        mvarbTomographic = False

        mvarWLID_SingleBWXRTaken = 0
        mvarWLID_BWXR2Taken = 0
        mvarWLID_BWXR4Taken = 0
        mvarWLID_PATakenFirstTooth = 0
        mvarWLID_PATakenAdditionalTooth = 0
        mvarWLID_PANOTaken = 0
        mvarWLID_CompleteSeriesTaken = 0
        mvarWLID_CephalometricFilmTaken = 0
        mvarWLID_OcculsalFilmMax = 0
        mvarWLID_OcculsalFilmMan = 0
        mvarWLID_ExtraoralInit = 0
        mvarWLID_ExtraoralAdd = 0
        mvarWLID_PALateralSkull = 0
        mvarWLID_Sialography = 0
        mvarWLID_TMJArth = 0
        mvarWLID_OtherTMJ = 0
        mvarWLID_Tomographic = 0
                
    Else
        dfDirtyFlag = dfNoAction
        'Set the values from the db including "" to allow for nulls
        mvarbNoRadTaken = CBool(Val(rs("NoRadTaken") & ""))
        mvarbPrevRadRev = CBool(Val(rs("PreviousRadReviewed") & ""))
        mvarbSingleBWXRTaken = CBool(Val(rs("SingleBWXRTaken") & ""))
        mvarb2BWXRTaken = CBool(Val(rs("BWXR2Taken") & ""))
        mvarb4BWXRTaken = CBool(Val(rs("BWXR4Taken") & ""))
        mvarbPeriapicalRadTaken = CBool(Val(rs("PeriapicalRadTaken") & ""))
        mvarbPANOTaken = CBool(Val(rs("PANOTaken") & ""))
        mvarbCompleteSeriesBWTaken = CBool(Val(rs("CompleteSeriesBWTaken") & ""))
        mvarbCephalometricFilmTaken = CBool(Val(rs("CephalometricFilmTaken") & ""))
        mvarbOcclusalFilmMax = CBool(Val(rs("OcclusalFilmMax") & ""))
        mvarbOcclusalFilmMan = CBool(Val(rs("OcclusalFilmMan") & ""))
        mvarbOtherFilms = CBool(Val(rs("OtherFilms") & ""))

        mvariNumberPA = Val(rs("NumberPA") & "")

        mvarbExtraoralInit = CBool(Val(rs("ExtraoralInit") & ""))
        mvarbExtraoralAdd = CBool(Val(rs("ExtraoralAdd") & ""))
        mvarbPALateralSkull = CBool(Val(rs("PALateralSkull") & ""))
        mvarbSialography = CBool(Val(rs("Sialography") & ""))
        mvarbTMJArthrogram = CBool(Val(rs("TMJArthrogram") & ""))
        mvarbOtherTMJ = CBool(Val(rs("OtherTMJ") & ""))
        mvarbTomographic = CBool(Val(rs("Tomographic") & ""))

        mvarWLID_SingleBWXRTaken = Val(rs("WID_SingleBWXRTaken") & "")
        mvarWLID_BWXR2Taken = Val(rs("WID_BWXR2Taken") & "")
        mvarWLID_BWXR4Taken = Val(rs("WID_BWXR4Taken") & "")
        mvarWLID_PATakenFirstTooth = Val(rs("WID_PATakenFirstTooth") & "")
        mvarWLID_PATakenAdditionalTooth = Val(rs("WID_PATakenAdditionalTooth") & "")
        mvarWLID_PANOTaken = Val(rs("WID_PANOTaken") & "")
        mvarWLID_CompleteSeriesTaken = Val(rs("WID_CompleteSeriesTaken") & "")
        mvarWLID_CephalometricFilmTaken = Val(rs("WID_CephalometricFilmTaken") & "")
        mvarWLID_OcculsalFilmMax = Val(rs("WID_OcculsalFilmMax") & "")
        mvarWLID_OcculsalFilmMan = Val(rs("WID_OcculsalFilmMan") & "")
        mvarWLID_ExtraoralInit = Val(rs("WID_ExtraoralInit") & "")
        mvarWLID_ExtraoralAdd = Val(rs("WID_ExtraoralAdd") & "")
        mvarWLID_PALateralSkull = Val(rs("WID_PALateralSkull") & "")
        mvarWLID_Sialography = Val(rs("WID_Sialography") & "")
        mvarWLID_TMJArth = Val(rs("WID_TMJArth") & "")
        mvarWLID_OtherTMJ = Val(rs("WID_OtherTMJ") & "")
        mvarWLID_Tomographic = Val(rs("WID_Tomographic") & "")
        
    End If
    
    rs.Close
    
    'Load the values
    sSQL = "Select "
    sSQL = sSQL & sSQL_Select2
    sSQL = sSQL & " From TEMPL_RADIOLOGY "
    sSQL = sSQL & " Where DENTAL_ENCOUNTER_ID = " & mvarlDentalEncounterID

    Set rs = goconn.OpenRecordset(sSQL)
       
    If rs.EOF Then
        dfDirtyFlag = dfAdd
        
        mvarsPrevRadRevFindString = ""
        mvarsSingleBWXRFindString = ""
        mvars2BWXRFindString = ""
        mvars4BWXRFindString = ""
        mvarsPANOFindString = ""
        mvarsCompleteSeriesFindString = ""
        mvarsCephalometricFilmFindString = ""
                
    Else
        dfDirtyFlag = dfNoAction
        'Set the values from the db including "" to allow for nulls
        mvarsPrevRadRevFindString = rs("PrevRadFindings") & ""
        mvarsSingleBWXRFindString = rs("SingleBWXRFindings") & ""
        mvars2BWXRFindString = rs("BWXR2Findings") & ""
        mvars4BWXRFindString = rs("BWXR4Findings") & ""
        mvarsPANOFindString = rs("PANOFindings") & ""
        mvarsCompleteSeriesFindString = rs("CompleteSeriesBWFindings") & ""
        mvarsCephalometricFilmFindString = rs("CephalometricFilmFindings") & ""
        
    End If
    
    rs.Close
    
    'Load the values
    sSQL = "Select "
    sSQL = sSQL & sSQL_Select3
    sSQL = sSQL & " From TEMPL_RADIOLOGY "
    sSQL = sSQL & " Where DENTAL_ENCOUNTER_ID = " & mvarlDentalEncounterID

    Set rs = goconn.OpenRecordset(sSQL)
       
    If rs.EOF Then
        dfDirtyFlag = dfAdd
        
        mvarsOcclusalFilmMaxFindString = ""
        mvarsOcclusalFilmManFindString = ""
        mvarsOtherFindString = ""
        
        mvarsExtraoralInitFindString = ""
        mvarsExtraoralAddFindString = ""
        mvarsPALateralSkullFindString = ""
        mvarsSialographyFindString = ""
        mvarsTMJArthrogramFindString = ""
        mvarsOtherTMJFindString = ""
        mvarsTomographicFindString = ""
                
    Else
        dfDirtyFlag = dfNoAction
        'Set the values from the db including "" to allow for nulls
        
        mvarsOcclusalFilmMaxFindString = rs("OcclusalFilmMaxFindings") & ""
        mvarsOcclusalFilmManFindString = rs("OcclusalFilmManFindings") & ""
        
        mvarsExtraoralInitFindString = rs("ExtraoralInitFindings") & ""
        mvarsExtraoralAddFindString = rs("ExtraoralAddFindings") & ""
        mvarsPALateralSkullFindString = rs("PALateralSkullFindings") & ""
        mvarsSialographyFindString = rs("SialographyFindings") & ""
        mvarsTMJArthrogramFindString = rs("TMJArthrogramFindings") & ""
        mvarsOtherTMJFindString = rs("OtherTMJFindings") & ""
        mvarsTomographicFindString = rs("TomographicFindings") & ""
        
    End If
    
    rs.Close
    
    'Load the values
    sSQL = "Select "
    sSQL = sSQL & sSQL_Select4
    sSQL = sSQL & " From TEMPL_RADIOLOGY "
    sSQL = sSQL & " Where DENTAL_ENCOUNTER_ID = " & mvarlDentalEncounterID

    Set rs = goconn.OpenRecordset(sSQL)
       
    If rs.EOF Then
        dfDirtyFlag = dfAdd
                
        mvarsRadNotes = ""
        
    Else
        dfDirtyFlag = dfNoAction
        'Set the values from the db including "" to allow for nulls
        mvarsRadNotes = rs("RadiologyNotes") & ""
    End If
    
    rs.Close

    Set rs = Nothing
    
    clsISection_LoadSection = True
    
Exit Function

ErrHnd:
    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & "." & sPROC_NAME, App.Title, vbCritical
End Function

Private Property Get clsISection_PatientID() As Long

    clsISection_PatientID = mvarlPatientID

End Property

Private Function clsISection_SaveSection() As Boolean
Dim sSQL As String

Dim iLoop As Integer

Const sPROC_NAME As String = "SaveSection"


On Error GoTo ErrHnd

    mcolRadFindings.SaveSection

    If dfDirtyFlag = dfNoAction Then
        clsISection_SaveSection = True
        Exit Function
    Else
        'Update the workload
        For iLoop = 0 To enumWorkLoadTypeMax - 1
            DoWorkLoadBizRules (iLoop)
        Next iLoop
    End If

    If dfDirtyFlag = dfNoAction Then
        clsISection_SaveSection = True
        Exit Function
        
    ElseIf dfDirtyFlag = dfAdd Then
        'The record does not exist, so create an insert statement
        sSQL = "Insert into TEMPL_RADIOLOGY ("
        sSQL = sSQL & sSQL_Insert
        sSQL = sSQL & ")"
        sSQL = sSQL & " Values ("
        sSQL = sSQL & CLng(mvarlDentalEncounterID) & ","
        sSQL = sSQL & CInt(mvarbNoRadTaken) & ","
        sSQL = sSQL & CInt(mvarbPrevRadRev) & ","
        sSQL = sSQL & CInt(mvarbSingleBWXRTaken) & ","
        sSQL = sSQL & CInt(mvarb2BWXRTaken) & ","
        sSQL = sSQL & CInt(mvarb4BWXRTaken) & ","
        sSQL = sSQL & CInt(mvarbPeriapicalRadTaken) & ","
        sSQL = sSQL & CInt(mvarbPANOTaken) & ","
        sSQL = sSQL & CInt(mvarbCompleteSeriesBWTaken) & ","
        sSQL = sSQL & CInt(mvarbCephalometricFilmTaken) & ","
        sSQL = sSQL & CInt(mvarbOcclusalFilmMax) & ","
        sSQL = sSQL & CInt(mvarbOcclusalFilmMan) & ","
        sSQL = sSQL & CInt(mvarbOtherFilms) & ","

        sSQL = sSQL & CInt(mvariNumberPA) & ","

        sSQL = sSQL & CInt(mvarbExtraoralInit) & ","
        sSQL = sSQL & CInt(mvarbExtraoralAdd) & ","
        sSQL = sSQL & CInt(mvarbPALateralSkull) & ","
        sSQL = sSQL & CInt(mvarbSialography) & ","
        sSQL = sSQL & CInt(mvarbTMJArthrogram) & ","
        sSQL = sSQL & CInt(mvarbOtherTMJ) & ","
        sSQL = sSQL & CInt(mvarbTomographic) & ","

        sSQL = sSQL & CLng(mvarWLID_SingleBWXRTaken) & ","
        sSQL = sSQL & CLng(mvarWLID_BWXR2Taken) & ","
        sSQL = sSQL & CLng(mvarWLID_BWXR4Taken) & ","
        sSQL = sSQL & CLng(mvarWLID_PATakenFirstTooth) & ","
        sSQL = sSQL & CLng(mvarWLID_PATakenAdditionalTooth) & ","
        sSQL = sSQL & CLng(mvarWLID_PANOTaken) & ","
        sSQL = sSQL & CLng(mvarWLID_CompleteSeriesTaken) & ","
        sSQL = sSQL & CLng(mvarWLID_CephalometricFilmTaken) & ","
        sSQL = sSQL & CLng(mvarWLID_OcculsalFilmMax) & ","
        sSQL = sSQL & CLng(mvarWLID_OcculsalFilmMan) & ","
        sSQL = sSQL & CLng(mvarWLID_ExtraoralInit) & ","
        sSQL = sSQL & CLng(mvarWLID_ExtraoralAdd) & ","
        sSQL = sSQL & CLng(mvarWLID_PALateralSkull) & ","
        sSQL = sSQL & CLng(mvarWLID_Sialography) & ","
        sSQL = sSQL & CLng(mvarWLID_TMJArth) & ","
        sSQL = sSQL & CLng(mvarWLID_OtherTMJ) & ","
        sSQL = sSQL & CLng(mvarWLID_Tomographic) & ","
        
        'KDunne 4/02/2002 SCR 22630
        'mvarsRadNotes was out of order in the value section.
        sSQL = sSQL & "'" & RQ(mvarsRadNotes) & "'" & ","
        'End scr 22630
        
        sSQL = sSQL & "'" & (mvarsPrevRadRevFindString) & "'" & ","
        sSQL = sSQL & "'" & (mvarsSingleBWXRFindString) & "'" & ","
        sSQL = sSQL & "'" & (mvars2BWXRFindString) & "'" & ","
        sSQL = sSQL & "'" & (mvars4BWXRFindString) & "'" & ","
        sSQL = sSQL & "'" & (mvarsPANOFindString) & "'" & ","
        sSQL = sSQL & "'" & (mvarsCompleteSeriesFindString) & "'" & ","
        sSQL = sSQL & "'" & (mvarsCephalometricFilmFindString) & "'" & ","
        sSQL = sSQL & "'" & (mvarsOcclusalFilmMaxFindString) & "'" & ","
        sSQL = sSQL & "'" & (mvarsOcclusalFilmManFindString) & "'" & ","
        
        sSQL = sSQL & "'" & (mvarsExtraoralInitFindString) & "'" & ","
        sSQL = sSQL & "'" & (mvarsExtraoralAddFindString) & "'" & ","
        sSQL = sSQL & "'" & (mvarsPALateralSkullFindString) & "'" & ","
        sSQL = sSQL & "'" & (mvarsSialographyFindString) & "'" & ","
        sSQL = sSQL & "'" & (mvarsTMJArthrogramFindString) & "'" & ","
        sSQL = sSQL & "'" & (mvarsOtherTMJFindString) & "'" & ","
        sSQL = sSQL & "'" & (mvarsTomographicFindString) & "'"
                
        sSQL = sSQL & ")"

    ElseIf dfDirtyFlag = dfUpdate Then
        'The record exists, so create an update statement
        sSQL = "Update TEMPL_RADIOLOGY Set" _
            & " NoRadTaken = " & CInt(mvarbNoRadTaken) & ", " _
            & " PreviousRadReviewed = " & CInt(mvarbPrevRadRev) & ", " _
            & " SingleBWXRTaken = " & CInt(mvarbSingleBWXRTaken) & ", " _
            & " BWXR2Taken = " & CInt(mvarb2BWXRTaken) & ", " _
            & " BWXR4Taken = " & CInt(mvarb4BWXRTaken) & ", " _
            & " PeriapicalRadTaken = " & CInt(mvarbPeriapicalRadTaken) & ", " _
            & " PANOTaken = " & CInt(mvarbPANOTaken) & ", "
            
        sSQL = sSQL & " CompleteSeriesBWTaken = " & CInt(mvarbCompleteSeriesBWTaken) & ", " _
            & " CephalometricFilmTaken = " & CInt(mvarbCephalometricFilmTaken) & ", " _
            & " OcclusalFilmMax = " & CInt(mvarbOcclusalFilmMax) & ", " _
            & " OcclusalFilmMan = " & CInt(mvarbOcclusalFilmMan) & ", " _
            & " OtherFilms = " & CInt(mvarbOtherFilms) & ", " _
            & " NumberPA = " & CInt(mvariNumberPA) & ", "
            
        sSQL = sSQL & " ExtraoralInit = " & CInt(mvarbExtraoralInit) & ", " _
            & " ExtraoralAdd = " & CInt(mvarbExtraoralAdd) & ", " _
            & " PALateralSkull = " & CInt(mvarbPALateralSkull) & ", " _
            & " Sialography = " & CInt(mvarbSialography) & ", " _
            & " TMJArthrogram = " & CInt(mvarbTMJArthrogram) & ", " _
            & " OtherTMJ = " & CInt(mvarbOtherTMJ) & ", " _
            & " Tomographic = " & CInt(mvarbTomographic) & ", "

        sSQL = sSQL & " WID_SingleBWXRTaken = " & CInt(mvarWLID_SingleBWXRTaken) & ", " _
            & " WID_BWXR2Taken = " & CInt(mvarWLID_BWXR2Taken) & ", " _
            & " WID_BWXR4Taken = " & CInt(mvarWLID_BWXR4Taken) & ", " _
            & " WID_PATakenFirstTooth = " & CInt(mvarWLID_PATakenFirstTooth) & ", " _
            & " WID_PATakenAdditionalTooth = " & CInt(mvarWLID_PATakenAdditionalTooth) & ", " _
            & " WID_PANOTaken = " & CInt(mvarWLID_PANOTaken) & ", " _
            & " WID_CompleteSeriesTaken = " & CInt(mvarWLID_CompleteSeriesTaken) & ", "

        sSQL = sSQL & " WID_CephalometricFilmTaken = " & CInt(mvarWLID_CephalometricFilmTaken) & ", " _
            & " WID_OcculsalFilmMax = " & CInt(mvarWLID_OcculsalFilmMax) & ", " _
            & " WID_OcculsalFilmMan = " & CInt(mvarWLID_OcculsalFilmMan) & ", " _
            & " WID_ExtraoralInit = " & CInt(mvarWLID_ExtraoralInit) & ", " _
            & " WID_ExtraoralAdd = " & CInt(mvarWLID_ExtraoralAdd) & ", " _
            & " WID_PALateralSkull = " & CInt(mvarWLID_PALateralSkull) & ", " _
            & " WID_Sialography = " & CInt(mvarWLID_Sialography) & ", " _
            & " WID_TMJArth = " & CInt(mvarWLID_TMJArth) & ", " _
            & " WID_OtherTMJ = " & CInt(mvarWLID_OtherTMJ) & ", " _
            & " WID_Tomographic = " & CInt(mvarWLID_Tomographic) & ", "

        sSQL = sSQL & " PrevRadFindings = " & "'" & (mvarsPrevRadRevFindString) & "'" & ", " _
            & " SingleBWXRFindings = " & "'" & (mvarsSingleBWXRFindString) & "'" & ", " _
            & " BWXR2Findings = " & "'" & (mvars2BWXRFindString) & "'" & ", " _
            & " BWXR4Findings = " & "'" & (mvars4BWXRFindString) & "'" & ", " _
            & " PANOFindings = " & "'" & (mvarsPANOFindString) & "'" & ", " _
            & " CompleteSeriesBWFindings = " & "'" & (mvarsCompleteSeriesFindString) & "'" & ", " _
            & " CephalometricFilmFindings = " & "'" & (mvarsCephalometricFilmFindString) & "'" & ", " _
            & " OcclusalFilmMaxFindings = " & "'" & (mvarsOcclusalFilmMaxFindString) & "'" & ", " _
            & " OcclusalFilmManFindings = " & "'" & (mvarsOcclusalFilmManFindString) & "'" & ", "
            
        sSQL = sSQL & " ExtraoralInitFindings = " & "'" & (mvarsExtraoralInitFindString) & "'" & ", " _
            & " ExtraoralAddFindings = " & "'" & (mvarsExtraoralAddFindString) & "'" & ", " _
            & " PALateralSkullFindings = " & "'" & (mvarsPALateralSkullFindString) & "'" & ", " _
            & " SialographyFindings = " & "'" & (mvarsSialographyFindString) & "'" & ", " _
            & " TMJArthrogramFindings = " & "'" & (mvarsTMJArthrogramFindString) & "'" & ", " _
            & " OtherTMJFindings = " & "'" & (mvarsOtherTMJFindString) & "'" & ", " _
            & " TomographicFindings = " & "'" & (mvarsTomographicFindString) & "'" & ", " _
            & " RadiologyNotes = '" & RQ(mvarsRadNotes) & "' " _
            & " Where DENTAL_ENCOUNTER_ID = " & mvarlDentalEncounterID
            
    End If
    
    'This does not work with the new sql serv
    'Debug.Assert False
    
    goconn.ExecuteSQL sSQL
    
    dfDirtyFlag = dfNoAction
    
    clsISection_SaveSection = True

Exit Function

ErrHnd:
    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & "." & sPROC_NAME, App.Title, vbCritical
End Function


Public Function ValueStringDisplay(ByVal colFindings As Collection) As String

    Dim objItem As clsRadCBItem
    
    On Error GoTo ErrHandler
    
    For Each objItem In colFindings
        If objItem.Checked = True Then
            ValueStringDisplay = ValueStringDisplay & ", " & objItem.Description
        End If
    Next
        
    'Remove the first ", "
    If ValueStringDisplay <> "" Then
        ValueStringDisplay = Mid(ValueStringDisplay, Len(", ") + 1)
    End If
    
    Set objItem = Nothing
    
Exit Function
    
ErrHandler:

End Function

Public Function FindingsDisplay(ByVal colFind As Collection) As String

    FindingsDisplay = ValueStringDisplay(colFind)

End Function

Private Function GetWorkloadCode(ByVal vlWorkLoadType As Long, _
                                 ByRef viMultiplier As Integer) As String

    'Add code here to figure out the workload code for this instance and return it.
    'Used in SaveSection

Dim iMultiCount As Integer

    GetWorkloadCode = ""
    viMultiplier = 0

    Select Case vlWorkLoadType
        Case WL_TypeEnum.enumSingleBWXRTaken
            If mvarbSingleBWXRTaken Then
                GetWorkloadCode = SingleBWXRTaken_WL.WorkloadId
                viMultiplier = 1
            End If

        Case WL_TypeEnum.enum2BWXRTaken
            If mvarb2BWXRTaken Then
                GetWorkloadCode = BWXR2Taken_WL.WorkloadId
                viMultiplier = 1
            End If

        Case WL_TypeEnum.enum4BWXRTaken
            If mvarb4BWXRTaken Then
                GetWorkloadCode = BWXR4Taken_WL.WorkloadId
                viMultiplier = 1
            End If

        Case WL_TypeEnum.enumPATaken
            'If mvarbPeriapicalRadTaken Then
            If mvariNumberPA > 0 Then
                GetWorkloadCode = PATaken_WL.WorkloadId
                viMultiplier = 1
            End If

        Case WL_TypeEnum.enumPATakenAddl
            If mvariNumberPA > 1 Then
                GetWorkloadCode = PATakenAddl_WL.WorkloadId
                viMultiplier = mvariNumberPA - 1
            End If

        Case WL_TypeEnum.enumPANOTaken
            If mvarbPANOTaken Then
                GetWorkloadCode = PANOTaken_WL.WorkloadId
                viMultiplier = 1
            End If

        Case WL_TypeEnum.enumCompleteSeriesTaken
            If mvarbCompleteSeriesBWTaken Then
                GetWorkloadCode = CompleteSeriesTaken_WL.WorkloadId
                viMultiplier = 1 'SCR 22314
            End If

        Case WL_TypeEnum.enumCephalometricTaken
            If mvarbCephalometricFilmTaken Then
                GetWorkloadCode = CephalometricFilmTaken_WL.WorkloadId
                viMultiplier = 1
            End If

        Case WL_TypeEnum.enumOcclusalFilmMax
            If mvarbOcclusalFilmMax Then
                GetWorkloadCode = OcculsalFilmMax_WL.WorkloadId
                viMultiplier = 1
            End If

        Case WL_TypeEnum.enumOcclusalFilmMan
            If mvarbOcclusalFilmMan Then
                GetWorkloadCode = OcculsalFilmMan_WL.WorkloadId
                viMultiplier = 1 'SCR 22314
            End If

        Case WL_TypeEnum.enumExtraoralInit
            If mvarbExtraoralInit Then
                GetWorkloadCode = ExtraoralInit_WL.WorkloadId
                viMultiplier = 1
            End If

        Case WL_TypeEnum.enumExtraoralAdd
            If mvarbExtraoralAdd Then
                GetWorkloadCode = ExtraoralAdd_WL.WorkloadId
                viMultiplier = 1
            End If

        Case WL_TypeEnum.enumPALateralSkull
            If mvarbPALateralSkull Then
                GetWorkloadCode = PALateralSkull_WL.WorkloadId
                viMultiplier = 1
            End If

        Case WL_TypeEnum.enumSialography
            If mvarbSialography Then
                GetWorkloadCode = Sialography_WL.WorkloadId
                viMultiplier = 1
            End If

        Case WL_TypeEnum.enumTMJArth
            If mvarbTMJArthrogram Then
                GetWorkloadCode = TMJArth_WL.WorkloadId
                viMultiplier = 1
            End If

        Case WL_TypeEnum.enumOtherTMJ
            If mvarbOtherTMJ Then
                GetWorkloadCode = OtherTMJ_WL.WorkloadId
                viMultiplier = 1
            End If

        Case WL_TypeEnum.enumTomographic
            If mvarbTomographic Then
                GetWorkloadCode = Tomographic_WL.WorkloadId
                viMultiplier = 1
            End If

    End Select

End Function

Public Sub DoWorkLoadBizRules(ByVal vlWorkLoadType As WL_TypeEnum)

Dim lItemType As Long
Dim dfFlag As DirtyFlagEnum
Dim iMultiplier As Integer
Dim sWorkLoadCode As String

Const sPROC_NAME = "DoWorkLoadBizRules"
    
On Error GoTo ErrHandler
           
    If vlWorkLoadType = enumSingleBWXRTaken Then lItemType = mvarWLID_SingleBWXRTaken
    If vlWorkLoadType = enum2BWXRTaken Then lItemType = mvarWLID_BWXR2Taken
    If vlWorkLoadType = enum4BWXRTaken Then lItemType = mvarWLID_BWXR4Taken
    If vlWorkLoadType = enumPATaken Then lItemType = mvarWLID_PATakenFirstTooth
    If vlWorkLoadType = enumPATakenAddl Then lItemType = mvarWLID_PATakenAdditionalTooth
    If vlWorkLoadType = enumPANOTaken Then lItemType = mvarWLID_PANOTaken
    If vlWorkLoadType = enumCompleteSeriesTaken Then lItemType = mvarWLID_CompleteSeriesTaken
    If vlWorkLoadType = enumCephalometricTaken Then lItemType = mvarWLID_CephalometricFilmTaken
    If vlWorkLoadType = enumOcclusalFilmMax Then lItemType = mvarWLID_OcculsalFilmMax
    If vlWorkLoadType = enumOcclusalFilmMan Then lItemType = mvarWLID_OcculsalFilmMan
    If vlWorkLoadType = enumExtraoralInit Then lItemType = mvarWLID_ExtraoralInit
    If vlWorkLoadType = enumExtraoralAdd Then lItemType = mvarWLID_ExtraoralAdd
    If vlWorkLoadType = enumPALateralSkull Then lItemType = mvarWLID_PALateralSkull
    If vlWorkLoadType = enumSialography Then lItemType = mvarWLID_Sialography
    If vlWorkLoadType = enumTMJArth Then lItemType = mvarWLID_TMJArth
    If vlWorkLoadType = enumOtherTMJ Then lItemType = mvarWLID_OtherTMJ
    If vlWorkLoadType = enumTomographic Then lItemType = mvarWLID_Tomographic
       
    sWorkLoadCode = GetWorkloadCode(vlWorkLoadType, iMultiplier)
    If sWorkLoadCode <> "" Then
        If lItemType = 0 Then
            dfFlag = dfAdd
        Else
            dfFlag = dfUpdate
        End If
        If Not UpdateWorkload(lItemType, dfFlag, mvarlDentalEncounterID, sWorkLoadCode, iMultiplier) Then
            Exit Sub
        End If
    Else
        If lItemType <> 0 Then
            'Delete the workload credit
            UpdateWorkload lItemType, dfdelete, mvarlDentalEncounterID, "", iMultiplier
        End If
    End If
    
    If vlWorkLoadType = enumSingleBWXRTaken Then mvarWLID_SingleBWXRTaken = lItemType
    If vlWorkLoadType = enum2BWXRTaken Then mvarWLID_BWXR2Taken = lItemType
    If vlWorkLoadType = enum4BWXRTaken Then mvarWLID_BWXR4Taken = lItemType
    If vlWorkLoadType = enumPATaken Then mvarWLID_PATakenFirstTooth = lItemType
    If vlWorkLoadType = enumPATakenAddl Then mvarWLID_PATakenAdditionalTooth = lItemType
    If vlWorkLoadType = enumPANOTaken Then mvarWLID_PANOTaken = lItemType
    If vlWorkLoadType = enumCompleteSeriesTaken Then mvarWLID_CompleteSeriesTaken = lItemType
    If vlWorkLoadType = enumCephalometricTaken Then mvarWLID_CephalometricFilmTaken = lItemType
    If vlWorkLoadType = enumOcclusalFilmMax Then mvarWLID_OcculsalFilmMax = lItemType
    If vlWorkLoadType = enumOcclusalFilmMan Then mvarWLID_OcculsalFilmMan = lItemType
    If vlWorkLoadType = enumExtraoralInit Then mvarWLID_ExtraoralInit = lItemType
    If vlWorkLoadType = enumExtraoralAdd Then mvarWLID_ExtraoralAdd = lItemType
    If vlWorkLoadType = enumPALateralSkull Then mvarWLID_PALateralSkull = lItemType
    If vlWorkLoadType = enumSialography Then mvarWLID_Sialography = lItemType
    If vlWorkLoadType = enumTMJArth Then mvarWLID_TMJArth = lItemType
    If vlWorkLoadType = enumOtherTMJ Then mvarWLID_OtherTMJ = lItemType
    If vlWorkLoadType = enumTomographic Then mvarWLID_Tomographic = lItemType

Exit Sub

ErrHandler:
    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & "." & sPROC_NAME, App.Title, vbCritical
    'Delete the workload credit if it didn't update correctly
    UpdateWorkload lItemType, dfdelete, mvarlDentalEncounterID, "", 1
End Sub

