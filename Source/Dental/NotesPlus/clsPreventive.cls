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
Attribute VB_Name = "clsPreventive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Implements clsISection

'Used to determine if the data should be inserted or updated when modified.
Private bSavedToDB As Boolean

Private mvarlPatientID As Long
Private mvarlDentalEncounterID As Long

Private dfDirtyFlag As DirtyFlagEnum
'local variable(s) to hold Preventive (PRV) property value(s)
'Oral hygiene (OH) demonstration
Private mvarbPRV_OH_Brushing             As Boolean     'local copy
Private mvarbPRV_OH_Flossing             As Boolean     'local copy
Private mvarsPRV_OH_Note                 As String      'local copy
'Wellness Counseling (WC)
Private mvarbPRV_WC_Nutritional          As Boolean     'local copy
Private mvarbPRV_WC_Tobacco              As Boolean     'local copy
Private mvarsPRV_WC_Tobacco_Note         As String      'local copy
Private mvarsPRV_WC_Nutritional_Note     As String      'local copy
'Delivery (DL) of
Private mvarbPRV_DL_MaxBleaching         As Boolean     'local copy
Private mvarbPRV_DL_MaxAthletic          As Boolean     'local copy
Private mvarbPRV_DL_MaxFluoride          As Boolean     'local copy
Private mvarbPRV_DL_ManBleaching         As Boolean     'local copy
Private mvarbPRV_DL_ManAthletic          As Boolean     'local copy
Private mvarbPRV_DL_ManFluoride          As Boolean     'local copy
'Impressions made (IM) of
Private mvarbPRV_IM_MaxMouthGuard        As Boolean     'local copy
Private mvarbPRV_IM_MaxFluoride          As Boolean     'local copy
Private mvarbPRV_IM_ManMouthGuard        As Boolean     'local copy
Private mvarbPRV_IM_ManFluoride          As Boolean     'local copy
Private mvarbPRV_IM_MaxBleach            As Boolean     'local copy
Private mvarbPRV_IM_ManBleach            As Boolean     'local copy
Private mvarbPRV_IM_MaxDiag              As Boolean     'local copy
Private mvarbPRV_IM_ManDiag              As Boolean     'local copy
Private mvarsPRV_Preventive_Comments     As String      'local copy
'local variable(s) to hold Prophylaxis/Fluoride property value(s)
'Dental Prophylaxis (PH) /Topical Fluoride
Private mvarbPRH_PROPH_Complete          As Boolean     'local copy
Private mvarbPRH_PROPH_Incomplete        As Boolean     'local copy
'Dental Prophylaxis - Fluoride used
Private mvarbPRH_PROPH_Fluoride_SNF2     As Boolean     'local copy
Private mvarbPRH_PROPH_Fluoride_APF      As Boolean     'local copy
Private mvarbPRH_PROPH_Fluoride_NAF      As Boolean     'local copy
Private mvarbPRH_PROPH_Fluoride_Varnish  As Boolean     'local copy
'Supragingival (SUP) scaling
Private mvarbPRH_SUP_FullMouth           As Boolean     'local copy
Private mvarbPRH_SUP_MaxRight            As Boolean     'local copy
Private mvarbPRH_SUP_ManRight            As Boolean     'local copy
Private mvarbPRH_SUP_MaxAnt              As Boolean     'local copy
Private mvarbPRH_Sup_ManAnt              As Boolean     'local copy
Private mvarbPRH_SUP_MaxLeft             As Boolean     'local copy
Private mvarbPRH_SUP_ManLeft             As Boolean     'local copy
'Polish (POL)
Private mvarbPRH_POL_FullMouth           As Boolean     'local copy
Private mvarbPRH_POL_MaxRight            As Boolean     'local copy
Private mvarbPRH_POL_MaxAnt              As Boolean     'local copy
Private mvarbPRH_POL_MaxLeft             As Boolean     'local copy
Private mvarbPRH_POL_ManLeft             As Boolean     'local copy
Private mvarbPRH_POL_ManRight            As Boolean     'local copy
Private mvarbPRH_POL_ManAnt              As Boolean     'local copy
Private mvarsPRH_Prophylaxis_Comments    As String      'local copy

'local variable(s) to hold Hygiene (HG) property value(s)
'Scaling and Root Planing
Private mvarbHG_Scaling_ManLeft_1To3          As Boolean     'local copy
Private mvarbHG_Scaling_ManLeft_4Plus          As Boolean     'local copy

Private mvarbHG_Scaling_ManRight_1To3         As Boolean     'local copy
Private mvarbHG_Scaling_ManRight_4Plus         As Boolean     'local copy

Private mvarbHG_Scaling_MaxLeft_1To3          As Boolean     'local copy
Private mvarbHG_Scaling_MaxLeft_4Plus          As Boolean     'local copy

Private mvarbHG_Scaling_MaxRight_1To3         As Boolean     'local copy
Private mvarbHG_Scaling_MaxRight_4Plus        As Boolean

'Private mvarbHG_Scaling_AllQuadr           As Boolean     'local copy

'Debridement (DM) and Maintenance
Private mvarbHG_DM_FullMouth             As Boolean     'local copy
Private mvarbHG_DM_Perio                 As Boolean     'local copy
Private mvarbHG_DM_Implant               As Boolean     'local copy
Private mvarsHG_DM_Note                  As String      'local copy
'Medicaments (MED) used
Private mvarbHG_MED_Actisite             As Boolean     'local copy
Private mvarbHG_MED_Atridox              As Boolean     'local copy
Private mvarbHG_MED_Chlorhexidine        As Boolean     'local copy
Private mvarbHG_MED_Doxycycline          As Boolean     'local copy
Private mvarbHG_MED_Fluoride_Sodium      As Boolean     'local copy
Private mvarbHG_MED_Fluoride_Stannous    As Boolean     'local copy
Private mvarbHG_MED_Fluoride_Varnish     As Boolean     'local copy
Private mvarbHG_MED_Hydrogen             As Boolean     'local copy
Private mvarbHG_MED_Painfree             As Boolean     'local copy
Private mvarbHG_MED_Perio                As Boolean     'local copy
Private mvarbHG_MED_Potassium            As Boolean     'local copy
Private mvarbHG_MED_Saline               As Boolean     'local copy
Private mvarbHG_MED_Tetracycline         As Boolean     'local copy
Private mvarbHG_MED_Topical              As Boolean     'local copy
Private mvarbHG_MED_Duraphet             As Boolean     'local copy
Private mvarsHG_MED_Note                 As String      'local copy
Private mblnHG_MED_Potassium_Ox          As Boolean


'Workload
Private mvarlWLID_Nutritional            As Long         'local copy
Private mvarlWLID_Tobacco                As Long         'local copy
Private mvarlWLID_OHDemo                 As Long         'local copy
Private mvarlWLID_Impress                As Long         'local copy
Private mvarlWLID_Proph                  As Long         'local copy
Private mvarlWLID_Bleaching              As Long         'local copy
Private mvarlWLID_Athletic               As Long         'local copy
Private mvarlWLID_Fluoride               As Long         'local copy
Private mvarlWLID_Scaling_1To3           As Long         'local copy
Private mvarlWLID_Scaling_4Plus          As Long         'local copy
Private mvarlWLID_FullMouth              As Long         'local copy
Private mvarlWLID_Perio                  As Long         'local copy
Private mvarlWLID_Implant                As Long         'local copy
Private mvarlWLID_SiteSpecific           As Long         'local copy
Private mvarlWLID_DAgent                 As Long         'local copy

Public Enum NotesCommandTypeEnum
   enumPRV_OH_Note = 0              ' Oral Hygiene Note
   enumPRV_WC_Nutritional_Note = 1  ' Wellness Counseling Nutritional Note
   enumPRV_WC_Tobacco_Note = 2      ' Wellness Counseling Tobacco Note
   enumPRV_Preventive_Comments = 3  ' Preventive Procedure Comments
   enumPRH_Prophylaxis_Comments = 4 ' Prophylaxis Comments
   enumHG_DM_Note = 5               ' Hygiene Debridement and Maintenance
   enumHG_MED_Note = 6              ' Hygiene Medicaments used
End Enum

Public Enum WorkLoadTypeEnum
    enumPRV_WC_Nutritional = 0
    enumPRV_WC_Tobacco = 1
    enumPRV_OH_Demo = 2
    enumPRV_Impressions = 3
    enumPRV_DL_Bleaching = 4
    enumPRV_DL_Athletic = 5
    enumPRV_DL_Fluoride = 6
    enumPROPH_Fluoride = 7
    enumHG_Scaling_1to3 = 8
    enumHG_DM_FullMouth = 9
    enumHG_DM_Perio = 10
    enumHG_DM_Implant = 11
    enumHG_MED_SiteSpecific = 12
    enumHG_MED_Agent = 13
    enumHG_Scaling_4Plus = 14
    enumWorkLoadTypeMax = 15
End Enum

Private Const cSQL_Select1 As String = "PRV_OH_Brushing,PRV_OH_Flossing,PRV_WC_Nutritional,PRV_WC_Tobacco," _
& "PRV_DL_MaxBleaching,PRV_DL_MaxAthletic,PRV_DL_MaxFluoride,PRV_DL_ManBleaching,PRV_DL_ManAthletic," _
& "PRV_DL_ManFluoride,PRV_IM_MaxMouthGuard,PRV_IM_MaxFluoride," _
& "PRV_IM_ManMouthGuard,PRV_IM_ManFluoride,PRV_IM_MaxBleach,PRV_IM_ManBleach,PRV_IM_MaxDiag," _
& "PRV_IM_ManDiag," _
& "PRH_Proph_Complete,PRH_Proph_Incomplete,PRH_Proph_Fluoride_SNF2," _
& "PRH_Proph_Fluoride_APF,PRH_Proph_Fluoride_NAF,PRH_Proph_Fluoride_Varnish,PRH_Sup_FullMouth," _
& "PRH_Sup_MaxRight,PRH_Sup_ManRight,PRH_Sup_MaxAnt,PRH_Sup_ManAnt,PRH_Sup_MaxLeft,PRH_Sup_ManLeft," _
& "PRH_Pol_FullMouth,PRH_Pol_MaxRight,PRH_Pol_MaxAnt,PRH_Pol_MaxLeft,PRH_Pol_ManLeft," _
& "PRH_Pol_ManRight,PRH_Pol_ManAnt"

Private Const cSQL_Select2 As String = "HG_Scaling_ManLeft_1To3,HG_Scaling_ManLeft_4Plus,HG_Scaling_ManRight_1To3,HG_Scaling_ManRight_4Plus,HG_Scaling_MaxLeft_1To3,HG_Scaling_MaxLeft_4Plus," _
& "HG_Scaling_MaxRight_1To3,HG_Scaling_MaxRight_4Plus,HG_DM_FullMouth,HG_DM_Perio,HG_DM_Implant,HG_Med_Actisite," _
& "HG_Med_Atridox,HG_Med_Chlorhexidine,HG_Med_Doxycycline,HG_Med_Fluoride_Sodium," _
& "HG_Med_Fluoride_Stannous,HG_Med_Fluoride_Varnish,HG_Med_Hydrogen,HG_MED_Painfree,HG_Med_Perio," _
& "HG_Med_Potassium, HG_MED_POtassium_Ox, HG_Med_Saline,HG_Med_Tetracycline,HG_Med_Topical,HG_MED_Duraphet"

Private Const cSQL_Insert As String = "PRV_OH_Brushing,PRV_OH_Flossing,PRV_WC_Nutritional,PRV_WC_Tobacco," _
& "PRV_DL_MaxBleaching,PRV_DL_MaxAthletic,PRV_DL_MaxFluoride,PRV_DL_ManBleaching,PRV_DL_ManAthletic," _
& "PRV_DL_ManFluoride,PRV_IM_MaxMouthGuard,PRV_IM_MaxFluoride," _
& "PRV_IM_ManMouthGuard,PRV_IM_ManFluoride,PRV_IM_MaxBleach,PRV_IM_ManBleach,PRV_IM_MaxDiag," _
& "PRV_IM_ManDiag," _
& "PRH_Proph_Complete,PRH_Proph_Incomplete,PRH_Proph_Fluoride_SNF2," _
& "PRH_Proph_Fluoride_APF,PRH_Proph_Fluoride_NAF,PRH_Proph_Fluoride_Varnish,PRH_Sup_FullMouth," _
& "PRH_Sup_MaxRight,PRH_Sup_ManRight,PRH_Sup_MaxAnt,PRH_Sup_ManAnt,PRH_Sup_MaxLeft,PRH_Sup_ManLeft," _
& "PRH_Pol_FullMouth,PRH_Pol_MaxRight,PRH_Pol_MaxAnt,PRH_Pol_MaxLeft,PRH_Pol_ManLeft," _
& "PRH_Pol_ManRight,PRH_Pol_ManAnt," _
& "HG_Scaling_ManLeft_1To3,HG_Scaling_ManLeft_4Plus,HG_Scaling_ManRight_1To3,HG_Scaling_ManRight_4Plus,HG_Scaling_MaxLeft_1To3,HG_Scaling_MaxLeft_4Plus," _
& "HG_Scaling_MaxRight_1To3,HG_Scaling_MaxRight_4Plus,HG_DM_FullMouth,HG_DM_Perio,HG_DM_Implant,HG_Med_Actisite," _
& "HG_Med_Atridox,HG_Med_Chlorhexidine,HG_Med_Doxycycline,HG_Med_Fluoride_Sodium," _
& "HG_Med_Fluoride_Stannous,HG_Med_Fluoride_Varnish,HG_Med_Hydrogen,HG_MED_Painfree,HG_Med_Perio," _
& "HG_Med_Potassium, HG_Med_Potassium_Ox, HG_Med_Saline,HG_Med_Tetracycline,HG_Med_Topical, HG_MED_Duraphet"

Private Type WorkLoadRecord
    WorkloadId As Long          'TX_ADA_CODE_ID
    WorkLoadDef As String       'ADA_Definition
    WorkloadCode As String      'ADA_Code
End Type

Private Nutritional_WL As WorkLoadRecord
Private Tobacco_WL As WorkLoadRecord
Private OHDemo_WL As WorkLoadRecord
Private Impress_WL As WorkLoadRecord

Private Bleaching_WL As WorkLoadRecord
Private Athletic_WL As WorkLoadRecord
Private Fluoride_WL As WorkLoadRecord
Private Proph_Fluoride_WL As WorkLoadRecord

Private Scaling_WL_1To3 As WorkLoadRecord
Private Scaling_WL_4Plus As WorkLoadRecord


Private FullMouth_WL As WorkLoadRecord
Private Perio_WL As WorkLoadRecord
Private Implant_WL As WorkLoadRecord
Private SiteSpecific As WorkLoadRecord
Private DAgent As WorkLoadRecord


Public Enum ProphWorkLoadType
    enumPROPH_CompIncompAdult = 0
    enumPROPH_FlorideUsedAdult = 1
    enumPROPH_BothAdult = 2
    enumPROPH_CompIncompChild = 3
    enumPROPH_FlorideUsedChild = 4
    enumPROPH_BothChild = 5
End Enum

Private ProphFluorideSpecific_WL As ProphWorkLoadType

Private ProphCompleteIncompleteAdult_WL As WorkLoadRecord
Private FluorideUsedAdult_WL As WorkLoadRecord
Private ProphFluorideBothAdult_WL As WorkLoadRecord
Private ProphCompleteIncompleteChild_WL As WorkLoadRecord
Private FluorideUsedChild_WL As WorkLoadRecord
Private ProphFluorideBothChild_WL As WorkLoadRecord
'local variable(s) to hold property value(s)
Private mvarbPatientIsAdult As Boolean 'local copy


Public Property Let PatientIsAdult(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PatientIsAdult = 5
    mvarbPatientIsAdult = vData
End Property


Public Property Get PatientIsAdult() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PatientIsAdult
    PatientIsAdult = mvarbPatientIsAdult
End Property





Public Property Let HG_MED_Duraphet(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.HG_MED_Duraphet = 5
    mvarbHG_MED_Duraphet = vData
End Property


Public Property Get HG_MED_Duraphet() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HG_MED_Duraphet
    HG_MED_Duraphet = mvarbHG_MED_Duraphet
End Property



Public Property Let WLID_DAgent(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WLID_DAgent = 5
    mvarlWLID_DAgent = vData
End Property


Public Property Get WLID_DAgent() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WLID_DAgent
    WLID_DAgent = mvarlWLID_DAgent
End Property



Public Property Let WLID_SiteSpecific(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WLID_SiteSpecific = 5
    mvarlWLID_SiteSpecific = vData
End Property


Public Property Get WLID_SiteSpecific() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WLID_SiteSpecific
    WLID_SiteSpecific = mvarlWLID_SiteSpecific
End Property



Public Property Let WLID_Implant(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WLID_Implant = 5
    mvarlWLID_Implant = vData
End Property


Public Property Get WLID_Implant() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WLID_Implant
    WLID_Implant = mvarlWLID_Implant
End Property



Public Property Let WLID_Perio(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WLID_Perio = 5
    mvarlWLID_Perio = vData
End Property


Public Property Get WLID_Perio() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WLID_Perio
    WLID_Perio = mvarlWLID_Perio
End Property



Public Property Let WLID_FullMouth(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WLID_FullMouth = 5
    mvarlWLID_FullMouth = vData
End Property


Public Property Get WLID_FullMouth() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WLID_FullMouth
    WLID_FullMouth = mvarlWLID_FullMouth
End Property



Public Property Let WLID_Fluoride(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WLID_Fluoride = 5
    mvarlWLID_Fluoride = vData
End Property


Public Property Get WLID_Fluoride() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WLID_Fluoride
    WLID_Fluoride = mvarlWLID_Fluoride
End Property



Public Property Let WLID_Athletic(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WLID_Athletic = 5
    mvarlWLID_Athletic = vData
End Property


Public Property Get WLID_Athletic() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WLID_Athletic
    WLID_Athletic = mvarlWLID_Athletic
End Property



Public Property Let WLID_Bleaching(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WLID_Bleaching = 5
    mvarlWLID_Bleaching = vData
End Property


Public Property Get WLID_Bleaching() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WLID_Bleaching
    WLID_Bleaching = mvarlWLID_Bleaching
End Property



Public Property Let WLID_Proph(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WLID_Proph = 5
    mvarlWLID_Proph = vData
End Property


Public Property Get WLID_Proph() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WLID_Proph
    WLID_Proph = mvarlWLID_Proph
End Property



Public Property Let WLID_Impress(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WLID_Impress = 5
    mvarlWLID_Impress = vData
End Property


Public Property Get WLID_Impress() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WLID_Impress
    WLID_Impress = mvarlWLID_Impress
End Property



Public Property Let WLID_OHDemo(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WLID_OHDemo = 5
    mvarlWLID_OHDemo = vData
End Property


Public Property Get WLID_OHDemo() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WLID_OHDemo
    WLID_OHDemo = mvarlWLID_OHDemo
End Property




Public Property Let WLID_Tobacco(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WLID_Tobacco = 5
    mvarlWLID_Tobacco = vData
End Property


Public Property Get WLID_Tobacco() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WLID_Tobacco
    WLID_Tobacco = mvarlWLID_Tobacco
End Property




Public Property Let PRV_WC_Tobacco_Note(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PRV_WC_Tobacco_Note = 5

    mvarsPRV_WC_Tobacco_Note = vData
    UpdateDirtyFlag
    
End Property

Public Property Get PRV_WC_Tobacco_Note() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PRV_WC_Tobacco_Note
    PRV_WC_Tobacco_Note = mvarsPRV_WC_Tobacco_Note
End Property

Public Property Let HG_MED_Note(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Med_Note = 5

    mvarsHG_MED_Note = vData
    UpdateDirtyFlag
    
End Property

Public Property Get HG_MED_Note() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Med_Note
    HG_MED_Note = mvarsHG_MED_Note
End Property

Public Property Let HG_MED_Topical(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Med_Topical = 5
    mvarbHG_MED_Topical = vData
    UpdateDirtyFlag
End Property

Public Property Get HG_MED_Topical() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Med_Topical
    HG_MED_Topical = mvarbHG_MED_Topical
End Property

Public Property Let HG_MED_Tetracycline(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Med_Tetracycline = 5
    mvarbHG_MED_Tetracycline = vData
    UpdateDirtyFlag
End Property

Public Property Get HG_MED_Tetracycline() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Med_Tetracycline
    HG_MED_Tetracycline = mvarbHG_MED_Tetracycline
End Property

Public Property Let HG_MED_Saline(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Med_Saline = 5
    mvarbHG_MED_Saline = vData
    UpdateDirtyFlag
End Property

Public Property Get HG_MED_Saline() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Med_Saline
    HG_MED_Saline = mvarbHG_MED_Saline
End Property

Public Property Let HG_MED_Potassium(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Med_Potassium = 5
    mvarbHG_MED_Potassium = vData
    UpdateDirtyFlag
End Property

Public Property Get HG_MED_Potassium() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Med_Potassium
    HG_MED_Potassium = mvarbHG_MED_Potassium
End Property

Public Property Let HG_MED_Potassium_Ox(ByVal blnData As Boolean)
    mblnHG_MED_Potassium_Ox = blnData
    UpdateDirtyFlag
End Property
Public Property Get HG_MED_Potassium_Ox() As Boolean
    HG_MED_Potassium_Ox = mblnHG_MED_Potassium_Ox
End Property

Public Property Let HG_MED_Perio(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Med_Perio = 5
    mvarbHG_MED_Perio = vData
    UpdateDirtyFlag
End Property

Public Property Get HG_MED_Perio() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Med_Perio
    HG_MED_Perio = mvarbHG_MED_Perio
End Property

Public Property Let HG_MED_Painfree(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Med_Paintree = 5
    mvarbHG_MED_Painfree = vData
    UpdateDirtyFlag
End Property

Public Property Get HG_MED_Painfree() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Med_Paintree
    HG_MED_Painfree = mvarbHG_MED_Painfree
End Property

Public Property Let HG_MED_Hydrogen(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Med_Hydrogen = 5
    mvarbHG_MED_Hydrogen = vData
    UpdateDirtyFlag
End Property

Public Property Get HG_MED_Hydrogen() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Med_Hydrogen
    HG_MED_Hydrogen = mvarbHG_MED_Hydrogen
End Property

Public Property Let HG_MED_Fluoride_Varnish(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Med_Fluoride_Varnish = 5
    mvarbHG_MED_Fluoride_Varnish = vData
    UpdateDirtyFlag
End Property

Public Property Get HG_MED_Fluoride_Varnish() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Med_Fluoride_Varnish
    HG_MED_Fluoride_Varnish = mvarbHG_MED_Fluoride_Varnish
End Property

Public Property Let HG_MED_Fluoride_Stannous(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Med_Fluoride_Stannous = 5
    mvarbHG_MED_Fluoride_Stannous = vData
    UpdateDirtyFlag
End Property

Public Property Get HG_MED_Fluoride_Stannous() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Med_Fluoride_Stannous
    HG_MED_Fluoride_Stannous = mvarbHG_MED_Fluoride_Stannous
End Property

Public Property Let HG_MED_Fluoride_Sodium(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Med_Fluoride_Sodium = 5
    mvarbHG_MED_Fluoride_Sodium = vData
    UpdateDirtyFlag
End Property

Public Property Get HG_MED_Fluoride_Sodium() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Med_Fluoride_Sodium
    HG_MED_Fluoride_Sodium = mvarbHG_MED_Fluoride_Sodium
End Property

Public Property Let HG_MED_Doxycycline(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Med_Doxycycline = 5
    mvarbHG_MED_Doxycycline = vData
    UpdateDirtyFlag
End Property

Public Property Get HG_MED_Doxycycline() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Med_Doxycycline
    HG_MED_Doxycycline = mvarbHG_MED_Doxycycline
End Property

Public Property Let HG_MED_Chlorhexidine(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Med_Chlorhexidine = 5
    mvarbHG_MED_Chlorhexidine = vData
    UpdateDirtyFlag
End Property

Public Property Get HG_MED_Chlorhexidine() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Med_Chlorhexidine
    HG_MED_Chlorhexidine = mvarbHG_MED_Chlorhexidine
End Property

Public Property Let HG_MED_Atridox(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Med_Atridox = 5
    mvarbHG_MED_Atridox = vData
    UpdateDirtyFlag
End Property

Public Property Get HG_MED_Atridox() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Med_Atridox
    HG_MED_Atridox = mvarbHG_MED_Atridox
End Property

Public Property Let HG_MED_Actisite(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Med_Actisite = 5
    mvarbHG_MED_Actisite = vData
    UpdateDirtyFlag
End Property

Public Property Get HG_MED_Actisite() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Med_Actisite
    HG_MED_Actisite = mvarbHG_MED_Actisite
End Property

Public Property Let HG_DM_Note(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DM_Note = 5
    mvarsHG_DM_Note = vData
    UpdateDirtyFlag
End Property

Public Property Get HG_DM_Note() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DM_Note
    HG_DM_Note = mvarsHG_DM_Note
End Property

Public Property Let HG_DM_Implant(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DM_Implant = 5
    mvarbHG_DM_Implant = vData
    UpdateDirtyFlag
End Property

Public Property Get HG_DM_Implant() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DM_Implant
    HG_DM_Implant = mvarbHG_DM_Implant
End Property

Public Property Let HG_DM_Perio(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DM_Perio = 5
    mvarbHG_DM_Perio = vData
    UpdateDirtyFlag
End Property

Public Property Get HG_DM_Perio() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DM_Perio
    HG_DM_Perio = mvarbHG_DM_Perio
End Property

Public Property Let HG_DM_FullMouth(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DM_FullMouth = 5
    mvarbHG_DM_FullMouth = vData
    UpdateDirtyFlag
End Property

Public Property Get HG_DM_FullMouth() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DM_FullMouth
    HG_DM_FullMouth = mvarbHG_DM_FullMouth
End Property

Public Property Let HG_Scaling_MaxRight_1To3(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.HScaling_MaxRight = 5
    mvarbHG_Scaling_MaxRight_1To3 = vData
    
    If vData = True Then
        mvarbHG_Scaling_MaxRight_4Plus = False
    End If
    
    UpdateDirtyFlag
End Property

Public Property Get HG_Scaling_MaxRight_1To3() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HScaling_MaxRight
    HG_Scaling_MaxRight_1To3 = mvarbHG_Scaling_MaxRight_1To3
End Property
Public Property Let HG_Scaling_MaxRight_4Plus(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.HScaling_MaxRight = 5
    mvarbHG_Scaling_MaxRight_4Plus = vData
    
    If vData = True Then
        mvarbHG_Scaling_MaxRight_1To3 = False
    End If
    
    UpdateDirtyFlag
End Property

Public Property Get HG_Scaling_MaxRight_4Plus() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HScaling_MaxRight
    HG_Scaling_MaxRight_4Plus = mvarbHG_Scaling_MaxRight_4Plus
End Property

Public Property Let HG_Scaling_MaxLeft_1To3(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.HScaling_MaxLeft = 5
    mvarbHG_Scaling_MaxLeft_1To3 = vData
    
    If vData = True Then
        mvarbHG_Scaling_MaxLeft_4Plus = False
    End If
    
    UpdateDirtyFlag
End Property

Public Property Get HG_Scaling_MaxLeft_1To3() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HScaling_MaxLeft
    HG_Scaling_MaxLeft_1To3 = mvarbHG_Scaling_MaxLeft_1To3
End Property
Public Property Let HG_Scaling_MaxLeft_4Plus(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.HScaling_MaxLeft = 5
    mvarbHG_Scaling_MaxLeft_4Plus = vData
    
    If vData = True Then
        mvarbHG_Scaling_MaxLeft_1To3 = False
    End If
    
    UpdateDirtyFlag
End Property

Public Property Get HG_Scaling_MaxLeft_4Plus() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HScaling_MaxLeft
    HG_Scaling_MaxLeft_4Plus = mvarbHG_Scaling_MaxLeft_4Plus
End Property

Public Property Let HG_Scaling_ManRight_1To3(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.HScaling_ManRight = 5
    mvarbHG_Scaling_ManRight_1To3 = vData
    
    If vData = True Then
        mvarbHG_Scaling_ManRight_4Plus = False
    End If
    
    UpdateDirtyFlag
End Property

Public Property Get HG_Scaling_ManRight_1To3() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HScaling_ManRight
    HG_Scaling_ManRight_1To3 = mvarbHG_Scaling_ManRight_1To3
End Property
Public Property Let HG_Scaling_ManRight_4Plus(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.HScaling_ManRight = 5
    mvarbHG_Scaling_ManRight_4Plus = vData
    
    If vData = True Then
        mvarbHG_Scaling_ManRight_1To3 = False
    End If
    
    UpdateDirtyFlag
End Property

Public Property Get HG_Scaling_ManRight_4Plus() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HScaling_ManRight
    HG_Scaling_ManRight_4Plus = mvarbHG_Scaling_ManRight_4Plus
End Property

Public Property Let HG_Scaling_ManLeft_1To3(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.HScaling_ManLeft = 5
    mvarbHG_Scaling_ManLeft_1To3 = vData
    
    If vData = True Then
        mvarbHG_Scaling_ManLeft_4Plus = False
    End If
    
    UpdateDirtyFlag
End Property

Public Property Get HG_Scaling_ManLeft_1To3() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HScaling_ManLeft
    HG_Scaling_ManLeft_1To3 = mvarbHG_Scaling_ManLeft_1To3
End Property

Public Property Let HG_Scaling_ManLeft_4Plus(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.HScaling_ManLeft = 5
    mvarbHG_Scaling_ManLeft_4Plus = vData
    
    If vData = True Then
        mvarbHG_Scaling_ManLeft_1To3 = False
    End If
    
    UpdateDirtyFlag
End Property

Public Property Get HG_Scaling_ManLeft_4Plus() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HScaling_ManLeft
    HG_Scaling_ManLeft_4Plus = mvarbHG_Scaling_ManLeft_4Plus
End Property
Public Property Let PRH_Prophylaxis_Comments(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Prophylaxis_Comments = 5
    mvarsPRH_Prophylaxis_Comments = vData
    UpdateDirtyFlag
End Property

Public Property Get PRH_Prophylaxis_Comments() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Prophylaxis_Comments
    PRH_Prophylaxis_Comments = mvarsPRH_Prophylaxis_Comments
End Property

Public Property Let PRH_POL_ManAnt(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Pol_ManAnt = 5
    mvarbPRH_POL_ManAnt = vData
    UpdateDirtyFlag
End Property

Public Property Get PRH_POL_ManAnt() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Pol_ManAnt
    PRH_POL_ManAnt = mvarbPRH_POL_ManAnt
End Property

Public Property Let PRH_POL_ManRight(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Pol_ManRight = 5
    mvarbPRH_POL_ManRight = vData
    UpdateDirtyFlag
End Property

Public Property Get PRH_POL_ManRight() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Pol_ManRight
    PRH_POL_ManRight = mvarbPRH_POL_ManRight
End Property

Public Property Let PRH_POL_ManLeft(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Pol_ManLeft = 5
    mvarbPRH_POL_ManLeft = vData
    UpdateDirtyFlag
End Property

Public Property Get PRH_POL_ManLeft() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Pol_ManLeft
    PRH_POL_ManLeft = mvarbPRH_POL_ManLeft
End Property

Public Property Let PRH_POL_MaxLeft(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Pol_MaxLeft = 5
    mvarbPRH_POL_MaxLeft = vData
    UpdateDirtyFlag
End Property

Public Property Get PRH_POL_MaxLeft() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Pol_MaxLeft
    PRH_POL_MaxLeft = mvarbPRH_POL_MaxLeft
End Property

Public Property Let PRH_POL_MaxAnt(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Pol_MaxAnt = 5
    mvarbPRH_POL_MaxAnt = vData
    UpdateDirtyFlag
End Property

Public Property Get PRH_POL_MaxAnt() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Pol_MaxAnt
    PRH_POL_MaxAnt = mvarbPRH_POL_MaxAnt
End Property

Public Property Let PRH_POL_MaxRight(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Pol_MaxRight = 5
    mvarbPRH_POL_MaxRight = vData
    UpdateDirtyFlag
End Property

Public Property Get PRH_POL_MaxRight() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Pol_MaxRight
    PRH_POL_MaxRight = mvarbPRH_POL_MaxRight
End Property

Public Property Let PRH_POL_FullMouth(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Pol_FullMouth = 5
    mvarbPRH_POL_FullMouth = vData
    DoGUIBizRules
    UpdateDirtyFlag
End Property

Public Property Get PRH_POL_FullMouth() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Pol_FullMouth
    PRH_POL_FullMouth = mvarbPRH_POL_FullMouth
End Property

Public Property Let PRH_SUP_ManLeft(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Sup_ManLeft = 5
    mvarbPRH_SUP_ManLeft = vData
    UpdateDirtyFlag
End Property

Public Property Get PRH_SUP_ManLeft() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Sup_ManLeft
    PRH_SUP_ManLeft = mvarbPRH_SUP_ManLeft
End Property

Public Property Let PRH_SUP_MaxLeft(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Sup_MaxLeft = 5
    mvarbPRH_SUP_MaxLeft = vData
    UpdateDirtyFlag
End Property

Public Property Get PRH_SUP_MaxLeft() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Sup_MaxLeft
    PRH_SUP_MaxLeft = mvarbPRH_SUP_MaxLeft
End Property

Public Property Let PRH_SUP_ManAnt(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Sup_ManAnt = 5
    mvarbPRH_Sup_ManAnt = vData
    UpdateDirtyFlag
End Property

Public Property Get PRH_SUP_ManAnt() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Sup_ManAnt
    PRH_SUP_ManAnt = mvarbPRH_Sup_ManAnt
End Property

Public Property Let PRH_SUP_MaxAnt(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Sup_MaxAnt = 5
    mvarbPRH_SUP_MaxAnt = vData
    UpdateDirtyFlag
End Property

Public Property Get PRH_SUP_MaxAnt() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Sup_MaxAnt
    PRH_SUP_MaxAnt = mvarbPRH_SUP_MaxAnt
End Property

Public Property Let PRH_SUP_ManRight(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Sup_ManRight = 5
    mvarbPRH_SUP_ManRight = vData
    UpdateDirtyFlag
End Property

Public Property Get PRH_SUP_ManRight() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Sup_ManRight
    PRH_SUP_ManRight = mvarbPRH_SUP_ManRight
End Property

Public Property Let PRH_SUP_MaxRight(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Sup_MaxRight = 5
    mvarbPRH_SUP_MaxRight = vData
    UpdateDirtyFlag
End Property

Public Property Get PRH_SUP_MaxRight() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Sup_MaxRight
    PRH_SUP_MaxRight = mvarbPRH_SUP_MaxRight
End Property

Public Property Let PRH_SUP_FullMouth(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Sup_FullMouth = 5
    mvarbPRH_SUP_FullMouth = vData
    DoGUIBizRules
    UpdateDirtyFlag
End Property

Public Property Get PRH_SUP_FullMouth() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Sup_FullMouth
    PRH_SUP_FullMouth = mvarbPRH_SUP_FullMouth
End Property

Public Property Let PRH_PROPH_Fluoride_Varnish(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Proph_Fluoride_Varnish = 5
    mvarbPRH_PROPH_Fluoride_Varnish = vData
    UpdateDirtyFlag
End Property

Public Property Get PRH_PROPH_Fluoride_Varnish() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Proph_Fluoride_Varnish
    PRH_PROPH_Fluoride_Varnish = mvarbPRH_PROPH_Fluoride_Varnish
End Property

Public Property Let PRH_PROPH_Fluoride_NAF(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Proph_Fluoride_NAF = 5
    mvarbPRH_PROPH_Fluoride_NAF = vData
    UpdateDirtyFlag
End Property

Public Property Get PRH_PROPH_Fluoride_NAF() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Proph_Fluoride_NAF
    PRH_PROPH_Fluoride_NAF = mvarbPRH_PROPH_Fluoride_NAF
End Property

Public Property Let PRH_PROPH_Fluoride_APF(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Proph_Fluoride_APF = 5
    mvarbPRH_PROPH_Fluoride_APF = vData
    UpdateDirtyFlag
End Property

Public Property Get PRH_PROPH_Fluoride_APF() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Proph_Fluoride_APF
    PRH_PROPH_Fluoride_APF = mvarbPRH_PROPH_Fluoride_APF
End Property

Public Property Let PRH_PROPH_Fluoride_SNF2(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Proph_Fluoride_SNF2 = 5
    mvarbPRH_PROPH_Fluoride_SNF2 = vData
    UpdateDirtyFlag
End Property

Public Property Get PRH_PROPH_Fluoride_SNF2() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Proph_Fluoride_SNF2
    PRH_PROPH_Fluoride_SNF2 = mvarbPRH_PROPH_Fluoride_SNF2
End Property

Public Property Let PRH_PROPH_Incomplete(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Proph_Incomplete = 5
    mvarbPRH_PROPH_Incomplete = vData
    UpdateDirtyFlag
End Property

Public Property Get PRH_PROPH_Incomplete() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Proph_Incomplete
    PRH_PROPH_Incomplete = mvarbPRH_PROPH_Incomplete
End Property

Public Property Let PRH_PROPH_Complete(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Proph_Complete = 5
    mvarbPRH_PROPH_Complete = vData
    UpdateDirtyFlag
End Property

Public Property Get PRH_PROPH_Complete() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Proph_Complete
    PRH_PROPH_Complete = mvarbPRH_PROPH_Complete
End Property

Public Property Let PRV_Preventive_Comments(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Preventive_Comments = 5
    mvarsPRV_Preventive_Comments = vData
    UpdateDirtyFlag
End Property

Public Property Get PRV_Preventive_Comments() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Preventive_Comments
    PRV_Preventive_Comments = mvarsPRV_Preventive_Comments
End Property

Public Property Let PRV_IM_ManDiag(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IM_ManDiag = 5
    mvarbPRV_IM_ManDiag = vData
    UpdateDirtyFlag
End Property

Public Property Get PRV_IM_ManDiag() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IM_ManDiag
    PRV_IM_ManDiag = mvarbPRV_IM_ManDiag
End Property

Public Property Let PRV_IM_MaxDiag(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IM_MaxDiag = 5
    mvarbPRV_IM_MaxDiag = vData
    UpdateDirtyFlag
End Property

Public Property Get PRV_IM_MaxDiag() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IM_MaxDiag
    PRV_IM_MaxDiag = mvarbPRV_IM_MaxDiag
End Property

Public Property Let PRV_IM_ManBleach(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IM_ManBleach = 5
    mvarbPRV_IM_ManBleach = vData
    UpdateDirtyFlag
End Property

Public Property Get PRV_IM_ManBleach() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IM_ManBleach
    PRV_IM_ManBleach = mvarbPRV_IM_ManBleach
End Property

Public Property Let PRV_IM_MaxBleach(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IM_MaxBleach = 5
    mvarbPRV_IM_MaxBleach = vData
    UpdateDirtyFlag
End Property

Public Property Get PRV_IM_MaxBleach() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IM_MaxBleach
    PRV_IM_MaxBleach = mvarbPRV_IM_MaxBleach
End Property

Public Property Let PRV_IM_ManFluoride(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IM_ManFluoride = 5
    mvarbPRV_IM_ManFluoride = vData
    UpdateDirtyFlag
End Property

Public Property Get PRV_IM_ManFluoride() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IM_ManFluoride
    PRV_IM_ManFluoride = mvarbPRV_IM_ManFluoride
End Property


Public Property Let PRV_DL_ManFluoride(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.D_Fluoride = 5
    mvarbPRV_DL_ManFluoride = vData
    UpdateDirtyFlag
End Property

Public Property Get PRV_DL_ManFluoride() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.D_Fluoride
    PRV_DL_ManFluoride = mvarbPRV_DL_ManFluoride
End Property

Public Property Let PRV_DL_ManAthletic(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.D_Athletic = 5
    mvarbPRV_DL_ManAthletic = vData
    UpdateDirtyFlag
End Property

Public Property Get PRV_DL_ManAthletic() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.D_Athletic
    PRV_DL_ManAthletic = mvarbPRV_DL_ManAthletic
End Property

Public Property Let PRV_DL_ManBleaching(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.D_Bleaching = 5
    mvarbPRV_DL_ManBleaching = vData
    UpdateDirtyFlag
End Property

Public Property Get PRV_DL_ManBleaching() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.D_Bleaching
    PRV_DL_ManBleaching = mvarbPRV_DL_ManBleaching
End Property

Public Property Let PRV_IM_ManMouthGuard(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IM_ManMouthGuard = 5
    mvarbPRV_IM_ManMouthGuard = vData
    UpdateDirtyFlag
End Property

Public Property Get PRV_IM_ManMouthGuard() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IM_ManMouthGuard
    PRV_IM_ManMouthGuard = mvarbPRV_IM_ManMouthGuard
End Property

Public Property Let PRV_IM_MaxFluoride(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IM_MaxFluoride = 5
    mvarbPRV_IM_MaxFluoride = vData
    UpdateDirtyFlag
End Property

Public Property Get PRV_IM_MaxFluoride() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IM_MaxFluoride
    PRV_IM_MaxFluoride = mvarbPRV_IM_MaxFluoride
End Property

Public Property Let PRV_IM_MaxMouthGuard(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IM_MaxMouthGuard = 5
    mvarbPRV_IM_MaxMouthGuard = vData
    UpdateDirtyFlag
End Property

Public Property Get PRV_IM_MaxMouthGuard() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IM_MaxMouthGuard
    PRV_IM_MaxMouthGuard = mvarbPRV_IM_MaxMouthGuard
End Property

Public Property Let PRV_DL_MaxFluoride(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.D_Fluoride = 5
    mvarbPRV_DL_MaxFluoride = vData
    UpdateDirtyFlag
End Property

Public Property Get PRV_DL_MaxFluoride() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.D_Fluoride
    PRV_DL_MaxFluoride = mvarbPRV_DL_MaxFluoride
End Property

Public Property Let PRV_DL_MaxAthletic(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.D_Athletic = 5
    mvarbPRV_DL_MaxAthletic = vData
    UpdateDirtyFlag
End Property

Public Property Get PRV_DL_MaxAthletic() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.D_Athletic
    PRV_DL_MaxAthletic = mvarbPRV_DL_MaxAthletic
End Property

Public Property Let PRV_DL_MaxBleaching(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.D_Bleaching = 5
    mvarbPRV_DL_MaxBleaching = vData
    UpdateDirtyFlag
End Property

Public Property Get PRV_DL_MaxBleaching() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.D_Bleaching
    PRV_DL_MaxBleaching = mvarbPRV_DL_MaxBleaching
End Property

Public Property Let PRV_WC_Nutritional_Note(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WC_Nutritional_Note = 5
    mvarsPRV_WC_Nutritional_Note = vData
    UpdateDirtyFlag
    
End Property

Public Property Get PRV_WC_Nutritional_Note() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WC_Nutritional_Note
    PRV_WC_Nutritional_Note = mvarsPRV_WC_Nutritional_Note
End Property

Public Property Let PRV_WC_Tobacco(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WC_Tobacco = 5
    mvarbPRV_WC_Tobacco = vData
    UpdateDirtyFlag
End Property

Public Property Get PRV_WC_Tobacco() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WC_Tobacco
    PRV_WC_Tobacco = mvarbPRV_WC_Tobacco
End Property

Public Property Let PRV_WC_Nutritional(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WC_Nutritional = 5
    mvarbPRV_WC_Nutritional = vData
    UpdateDirtyFlag
End Property

Public Property Get PRV_WC_Nutritional() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WC_Nutritional
    PRV_WC_Nutritional = mvarbPRV_WC_Nutritional
End Property

Public Property Let PRV_OH_Note(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.OHNote = 5
    mvarsPRV_OH_Note = vData
    UpdateDirtyFlag
End Property

Public Property Get PRV_OH_Note() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.OHNote
    PRV_OH_Note = mvarsPRV_OH_Note
End Property

Public Property Let PRV_OH_Flossing(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.OHFlossing = 5
    mvarbPRV_OH_Flossing = vData
    UpdateDirtyFlag
End Property

Public Property Get PRV_OH_Flossing() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.OHFlossing
    PRV_OH_Flossing = mvarbPRV_OH_Flossing
End Property

Public Property Let PRV_OH_Brushing(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.OHBrushing = 5
    mvarbPRV_OH_Brushing = vData
    UpdateDirtyFlag
End Property

Public Property Get PRV_OH_Brushing() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.OHBrushing
    PRV_OH_Brushing = mvarbPRV_OH_Brushing
End Property

Public Property Let WLID_Nutritional(vData As Long)
    
    mvarlWLID_Nutritional = vData
    UpdateDirtyFlag
    
End Property

Public Property Get WLID_Nutritional() As Long
    
    WLID_Nutritional = mvarlWLID_Nutritional
    
End Property

Private Property Get clsISection_DentalEncounterID() As Long
    
    clsISection_DentalEncounterID = mvarlDentalEncounterID

End Property

Private Function GetWorkloadCode(ByVal vlWorkLoadType As Long, _
                                 ByRef viMultiplier As Integer) As String

    'Add code here to figure out the workload code for this instance and return it.
    'Used in SaveSection
    
Dim iMultiCount As Integer
    
    GetWorkloadCode = ""
    viMultiplier = 0
    
    Select Case vlWorkLoadType
        Case WorkLoadTypeEnum.enumPRV_WC_Nutritional
            If mvarbPRV_WC_Nutritional Then
                GetWorkloadCode = Nutritional_WL.WorkloadId
                viMultiplier = 1
            End If
            
        Case WorkLoadTypeEnum.enumPRV_WC_Tobacco
            If mvarbPRV_WC_Tobacco Then
                GetWorkloadCode = Tobacco_WL.WorkloadId
                viMultiplier = 1
            End If
                                
        Case WorkLoadTypeEnum.enumPRV_OH_Demo
            If (mvarbPRV_OH_Brushing Or mvarbPRV_OH_Flossing) Then
                GetWorkloadCode = OHDemo_WL.WorkloadId
                viMultiplier = 1
            End If
                    
        Case WorkLoadTypeEnum.enumPRV_Impressions
            If (mvarbPRV_IM_MaxDiag And mvarbPRV_IM_ManDiag) Then
                GetWorkloadCode = Impress_WL.WorkloadId
                viMultiplier = 2
            ElseIf (mvarbPRV_IM_MaxDiag Or mvarbPRV_IM_ManDiag) Then
                GetWorkloadCode = Impress_WL.WorkloadId
                viMultiplier = 1
            End If
        
        Case WorkLoadTypeEnum.enumPRV_DL_Bleaching
            If (mvarbPRV_DL_MaxBleaching And mvarbPRV_DL_ManBleaching) Then
                GetWorkloadCode = Bleaching_WL.WorkloadId
                viMultiplier = 2
            ElseIf (mvarbPRV_DL_MaxBleaching Or mvarbPRV_DL_ManBleaching) Then
                GetWorkloadCode = Bleaching_WL.WorkloadId
                viMultiplier = 1
            End If
                        
        Case WorkLoadTypeEnum.enumPRV_DL_Athletic
            If (mvarbPRV_DL_MaxAthletic And mvarbPRV_DL_ManAthletic) Then
                GetWorkloadCode = Athletic_WL.WorkloadId
                viMultiplier = 2
            ElseIf (mvarbPRV_DL_MaxAthletic Or mvarbPRV_DL_ManAthletic) Then
                GetWorkloadCode = Athletic_WL.WorkloadId
                viMultiplier = 1
            End If
            
        Case WorkLoadTypeEnum.enumPRV_DL_Fluoride
            If (mvarbPRV_DL_MaxFluoride And mvarbPRV_DL_ManFluoride) Then
                GetWorkloadCode = Fluoride_WL.WorkloadId
                viMultiplier = 2
            ElseIf (mvarbPRV_DL_MaxFluoride Or mvarbPRV_DL_ManFluoride) Then
                GetWorkloadCode = Fluoride_WL.WorkloadId
                viMultiplier = 1
            End If
            
        Case WorkLoadTypeEnum.enumPROPH_Fluoride
        
        If mvarbPatientIsAdult Then
            'Proph
            If (mvarbPRH_PROPH_Complete Or mvarbPRH_PROPH_Incomplete) And _
               Not (mvarbPRH_PROPH_Fluoride_SNF2 Or mvarbPRH_PROPH_Fluoride_APF Or _
                mvarbPRH_PROPH_Fluoride_APF Or mvarbPRH_PROPH_Fluoride_NAF Or _
                mvarbPRH_PROPH_Fluoride_Varnish) Then
                GetWorkloadCode = ProphCompleteIncompleteAdult_WL.WorkloadId
                viMultiplier = 1
            End If
            
            'Fluoride
            If (mvarbPRH_PROPH_Fluoride_SNF2 Or mvarbPRH_PROPH_Fluoride_APF Or _
                mvarbPRH_PROPH_Fluoride_APF Or mvarbPRH_PROPH_Fluoride_NAF Or _
                mvarbPRH_PROPH_Fluoride_Varnish) And _
                Not (mvarbPRH_PROPH_Complete Or mvarbPRH_PROPH_Incomplete) Then
                GetWorkloadCode = FluorideUsedAdult_WL.WorkloadId
                viMultiplier = 1
            End If
            
            'Both
            If (mvarbPRH_PROPH_Complete Or mvarbPRH_PROPH_Incomplete) And _
               (mvarbPRH_PROPH_Fluoride_SNF2 Or mvarbPRH_PROPH_Fluoride_APF Or _
                mvarbPRH_PROPH_Fluoride_APF Or mvarbPRH_PROPH_Fluoride_NAF Or _
                mvarbPRH_PROPH_Fluoride_Varnish) Then
                GetWorkloadCode = ProphFluorideBothAdult_WL.WorkloadId
                viMultiplier = 1
            End If
            
        Else
            'Proph
            If (mvarbPRH_PROPH_Complete Or mvarbPRH_PROPH_Incomplete) And _
               Not (mvarbPRH_PROPH_Fluoride_SNF2 Or mvarbPRH_PROPH_Fluoride_APF Or _
                mvarbPRH_PROPH_Fluoride_APF Or mvarbPRH_PROPH_Fluoride_NAF Or _
                mvarbPRH_PROPH_Fluoride_Varnish) Then
                GetWorkloadCode = ProphCompleteIncompleteChild_WL.WorkloadId
                viMultiplier = 1
            End If
            'Fluoride
            If (mvarbPRH_PROPH_Fluoride_SNF2 Or mvarbPRH_PROPH_Fluoride_APF Or _
                mvarbPRH_PROPH_Fluoride_APF Or mvarbPRH_PROPH_Fluoride_NAF Or _
                mvarbPRH_PROPH_Fluoride_Varnish) And _
                Not (mvarbPRH_PROPH_Complete Or mvarbPRH_PROPH_Incomplete) Then
                GetWorkloadCode = FluorideUsedChild_WL.WorkloadId
                viMultiplier = 1
            End If
            'Both
            If (mvarbPRH_PROPH_Complete Or mvarbPRH_PROPH_Incomplete) And _
               (mvarbPRH_PROPH_Fluoride_SNF2 Or mvarbPRH_PROPH_Fluoride_APF Or _
                mvarbPRH_PROPH_Fluoride_APF Or mvarbPRH_PROPH_Fluoride_NAF Or _
                mvarbPRH_PROPH_Fluoride_Varnish) Then
                GetWorkloadCode = ProphFluorideBothChild_WL.WorkloadId
                viMultiplier = 1
            End If
        End If
        
        'Scaling 1-3 teeth
        Case WorkLoadTypeEnum.enumHG_Scaling_1to3

            If mvarbHG_Scaling_ManLeft_1To3 Then
                iMultiCount = iMultiCount + 1
            End If
            If mvarbHG_Scaling_ManRight_1To3 Then
                iMultiCount = iMultiCount + 1
            End If
            If mvarbHG_Scaling_MaxLeft_1To3 Then
                iMultiCount = iMultiCount + 1
            End If
            If mvarbHG_Scaling_MaxRight_1To3 Then
                iMultiCount = iMultiCount + 1
            End If
            
            If iMultiCount > 0 Then
                GetWorkloadCode = Scaling_WL_1To3.WorkloadId
                viMultiplier = iMultiCount
            End If

        'Scaling 4+ teeth
        Case WorkLoadTypeEnum.enumHG_Scaling_4Plus
            If mvarbHG_Scaling_MaxRight_4Plus Then
                iMultiCount = iMultiCount + 1
            End If
            If HG_Scaling_ManRight_4Plus Then
                iMultiCount = iMultiCount + 1
            End If
            If mvarbHG_Scaling_MaxLeft_4Plus Then
                iMultiCount = iMultiCount + 1
            End If
            If mvarbHG_Scaling_ManLeft_4Plus Then
                iMultiCount = iMultiCount + 1
            End If
            
            If iMultiCount > 0 Then
                GetWorkloadCode = Scaling_WL_4Plus.WorkloadId
                viMultiplier = iMultiCount
            End If

        Case WorkLoadTypeEnum.enumHG_DM_FullMouth
            If mvarbHG_DM_FullMouth Then
                GetWorkloadCode = FullMouth_WL.WorkloadId
                viMultiplier = 1
            End If
                
        Case WorkLoadTypeEnum.enumHG_DM_Perio
            If mvarbHG_DM_Perio Then
                GetWorkloadCode = Perio_WL.WorkloadId
                viMultiplier = 1
            End If
        
        Case WorkLoadTypeEnum.enumHG_DM_Implant
'SCR 11118
'            If mvarbHG_DM_Implant Then
'                GetWorkloadCode = Implant_WL.WorkloadId
'                viMultiplier = 1
'            End If
        
        Case WorkLoadTypeEnum.enumHG_MED_SiteSpecific
            If (mvarbHG_MED_Actisite Or mvarbHG_MED_Atridox Or mvarbHG_MED_Perio) Then
                GetWorkloadCode = SiteSpecific.WorkloadId
                viMultiplier = 1
            End If
        
        Case WorkLoadTypeEnum.enumHG_MED_Agent
            If (mvarbHG_MED_Fluoride_Varnish Or mvarbHG_MED_Potassium Or mvarbHG_MED_Duraphet Or mblnHG_MED_Potassium_Ox) Then
                GetWorkloadCode = DAgent.WorkloadId
                viMultiplier = 1
            End If
                
    End Select

End Function

Private Property Get clsISection_PatientID() As Long
    
    clsISection_PatientID = mvarlPatientID

End Property

Private Function clsISection_GetNote() As String

    clsISection_GetNote = "*****GetNote Method no longer valid.*****"

End Function

'*******************************************************************
'Function: GetNote
'Author: ASamedov
'Date:
'Returns: A string that is displayed in the narrative summary for this encounter
'Purpose: To compare the values from this encounter and the previous encounter and
'return a text note.
'*******************************************************************
Private Function clsISection_GetRTFNote() As String

Dim sTemp As String, sTemp1 As String, sTemp2 As String, sTemp3 As String
Dim sNote As String
Const sPROC_NAME = "clsISection_GetRTFNote"

On Error GoTo ErrHnd
    
    'Preventive section
    

    'Oral hygiene demo
    If mvarbPRV_OH_Brushing Or mvarbPRV_OH_Flossing Or mvarsPRV_OH_Note <> "" Then
        sNote = sNote + crtfStartControlName & "Oral hygiene demonstration: "
        If mvarbPRV_OH_Brushing Then
            sNote = sNote & "Brushing; "
        End If
        If mvarbPRV_OH_Flossing Then
            sNote = sNote & "Flossing; "
        End If
        'Preventive Comments/Notes
        If mvarsPRV_OH_Note <> "" Then
            sNote = sNote & "Oral Hygiene Comment: " & mvarsPRV_OH_Note + "; "
        End If
        'Remove the last "; "
        sNote = Left(sNote, Len(sNote) - Len("; "))
        sNote = sNote & crtfEndControlName
    End If
    
       
    'Wellness counseling        *********************************************************
    If mvarbPRV_WC_Nutritional Or mvarbPRV_WC_Tobacco Or _
        mvarsPRV_WC_Nutritional_Note <> "" Or mvarsPRV_WC_Tobacco_Note <> "" Then
        sNote = sNote + crtfStartControlName & "Wellness counseling: "
        If mvarbPRV_WC_Nutritional Then
            sNote = sNote & "Dental Nutritional Counseling; "
        End If
           
        If mvarbPRV_WC_Tobacco Then
            sNote = sNote & "Tobacco counseling performed; "
        End If
    
        If mvarsPRV_WC_Nutritional_Note <> "" Then
            sNote = sNote & "Wellness Counseling Nutritional Comment: " & mvarsPRV_WC_Nutritional_Note + "; "
        End If
    
        If mvarsPRV_WC_Tobacco_Note <> "" Then
            sNote = sNote & "Wellness Counseling Tobacco Comment: " & mvarsPRV_WC_Tobacco_Note + "; "
        End If
                  
        'Remove the last "; "
        sNote = Left(sNote, Len(sNote) - Len("; "))
        sNote = sNote & crtfEndControlName
    End If
              
              
    'Delivery of *********************************************************************
    
    If mvarbPRV_DL_MaxBleaching Then
        sTemp1 = sTemp1 & "Max Bleaching tray; "
    End If
       
    If mvarbPRV_DL_MaxAthletic Then
        sTemp1 = sTemp1 & "Max Athletic mouthguard; "
    End If
       
    If mvarbPRV_DL_MaxFluoride Then
        sTemp1 = sTemp1 & "Max Fluoride tray; "
    End If
       
    If mvarbPRV_DL_ManBleaching Then
        sTemp1 = sTemp1 & "Man Bleaching tray; "
    End If
       
    If mvarbPRV_DL_ManAthletic Then
        sTemp1 = sTemp1 & "Man Athletic mouthguard; "
    End If
       
    If mvarbPRV_DL_ManFluoride Then
        sTemp1 = sTemp1 & "Man Fluoride tray; "
    End If
              
    'Delivery of Header
    If sTemp1 <> "" Then
        sNote = sNote + crtfStartControlName & "Delivery of: "
        sNote = sNote & sTemp1
        'Remove the last "; "
        sNote = Left(sNote, Len(sNote) - Len("; "))
        sNote = sNote & crtfEndControlName
        sTemp1 = ""
    End If
              
    'Impressions made for *****************************************************************
    If mvarbPRV_IM_MaxMouthGuard Then
        sTemp1 = sTemp1 & "Max Mouthguard; "
    End If
       
    If mvarbPRV_IM_MaxFluoride Then
        sTemp1 = sTemp1 & "Max Fluoride Tray; "
    End If
       
    If mvarbPRV_IM_MaxBleach Then
        sTemp1 = sTemp1 & "Max Bleach Tray; "
    End If
       
    If mvarbPRV_IM_MaxDiag Then
        sTemp1 = sTemp1 & "Max Diag. Casts; "
    End If
       
    If mvarbPRV_IM_ManMouthGuard Then
        sTemp1 = sTemp1 & "Man Mouthguard; "
    End If
       
    If mvarbPRV_IM_ManFluoride Then
        sTemp1 = sTemp1 & "Man Fluoride Tray; "
    End If
       
    If mvarbPRV_IM_ManBleach Then
        sTemp1 = sTemp1 & "Man Bleach Tray; "
    End If
       
    If mvarbPRV_IM_ManDiag Then
        sTemp1 = sTemp1 & "Man Diag. Casts; "
    End If
                            
    'Impressions made for Header
    If sTemp1 <> "" Then
        sNote = sNote + crtfStartControlName & "Impressions made for: "
        sNote = sNote & sTemp1
        'Remove the last "; "
        sNote = Left(sNote, Len(sNote) - Len("; "))
        sNote = sNote & crtfEndControlName
        sTemp1 = ""
    End If

    If mvarsPRV_Preventive_Comments <> "" Then
        sNote = sNote + crtfStartControlName & "Preventive Comments: "
        sNote = sNote & mvarsPRV_Preventive_Comments
        sNote = sNote & crtfEndControlName

    End If
       
    clsISection_GetRTFNote = sNote
    
Exit Function

ErrHnd:
    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & "." & sPROC_NAME, App.Title, vbCritical
    
End Function

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
        For iLoop = 0 To enumWorkLoadTypeMax - 1
            DoWorkLoadBizRules (iLoop)
        Next iLoop
                
    End If
        
    If dfDirtyFlag = dfAdd Then
        'The record does not exist, so create an insert statement
        sSQL = "INSERT INTO TEMPL_PREVENTIVE ("
        sSQL = sSQL & "Dental_Encounter_ID,"
        sSQL = sSQL & cSQL_Insert & ","            'all fields except for notes and workloads
        sSQL = sSQL & "PRV_OH_Note,"
        sSQL = sSQL & "PRV_WC_Tobacco_Note,"
        sSQL = sSQL & "PRV_WC_Nutritional_Note,"
        sSQL = sSQL & "PRV_Preventive_Comments,"
        sSQL = sSQL & "PRH_Prophylaxis_Comments,"
        sSQL = sSQL & "HG_DM_Note,"
        sSQL = sSQL & "HG_Med_Note,"
        sSQL = sSQL & "WID_Nutritional,"
        sSQL = sSQL & "WID_Tobacco,"
        sSQL = sSQL & "WID_OHDemo,"
        sSQL = sSQL & "WID_Impress,"
        sSQL = sSQL & "WID_Proph,"
        sSQL = sSQL & "WID_Bleaching,"
        sSQL = sSQL & "WID_Athletic,"
        sSQL = sSQL & "WID_Fluoride,"
        sSQL = sSQL & "WID_Scaling_1To3,"
        sSQL = sSQL & "WID_Scaling_4Plus,"
        sSQL = sSQL & "WID_FullMouth,"
        sSQL = sSQL & "WID_Perio,"
        sSQL = sSQL & "WID_Implant,"
        sSQL = sSQL & "WID_SiteSpecific,"
        sSQL = sSQL & "WID_DAgent"
        
        sSQL = sSQL & ") VALUES ("
        
        sSQL = sSQL & CLng(mvarlDentalEncounterID) & ","
        sSQL = sSQL & CInt(mvarbPRV_OH_Brushing) & ","
        sSQL = sSQL & CInt(mvarbPRV_OH_Flossing) & ","
        sSQL = sSQL & CInt(mvarbPRV_WC_Nutritional) & ","
        sSQL = sSQL & CInt(mvarbPRV_WC_Tobacco) & ","
        sSQL = sSQL & CInt(mvarbPRV_DL_MaxBleaching) & ","
        sSQL = sSQL & CInt(mvarbPRV_DL_MaxAthletic) & ","
        sSQL = sSQL & CInt(mvarbPRV_DL_MaxFluoride) & ","
        sSQL = sSQL & CInt(mvarbPRV_DL_ManBleaching) & ","
        sSQL = sSQL & CInt(mvarbPRV_DL_ManAthletic) & ","
        sSQL = sSQL & CInt(mvarbPRV_DL_ManFluoride) & ","
        sSQL = sSQL & CInt(mvarbPRV_IM_MaxMouthGuard) & ","
        sSQL = sSQL & CInt(mvarbPRV_IM_MaxFluoride) & ","
        sSQL = sSQL & CInt(mvarbPRV_IM_ManMouthGuard) & ","
        sSQL = sSQL & CInt(mvarbPRV_IM_ManFluoride) & ","
        sSQL = sSQL & CInt(mvarbPRV_IM_MaxBleach) & ","
        sSQL = sSQL & CInt(mvarbPRV_IM_ManBleach) & ","
        sSQL = sSQL & CInt(mvarbPRV_IM_MaxDiag) & ","
        sSQL = sSQL & CInt(mvarbPRV_IM_ManDiag) & ","
        sSQL = sSQL & CInt(mvarbPRH_PROPH_Complete) & ","
        sSQL = sSQL & CInt(mvarbPRH_PROPH_Incomplete) & ","
        sSQL = sSQL & CInt(mvarbPRH_PROPH_Fluoride_SNF2) & ","
        sSQL = sSQL & CInt(mvarbPRH_PROPH_Fluoride_APF) & ","
        sSQL = sSQL & CInt(mvarbPRH_PROPH_Fluoride_NAF) & ","
        sSQL = sSQL & CInt(mvarbPRH_PROPH_Fluoride_Varnish) & ","
        sSQL = sSQL & CInt(mvarbPRH_SUP_FullMouth) & ","
        sSQL = sSQL & CInt(mvarbPRH_SUP_MaxRight) & ","
        sSQL = sSQL & CInt(mvarbPRH_SUP_ManRight) & ","
        sSQL = sSQL & CInt(mvarbPRH_SUP_MaxAnt) & ","
        sSQL = sSQL & CInt(mvarbPRH_Sup_ManAnt) & ","
        sSQL = sSQL & CInt(mvarbPRH_SUP_MaxLeft) & ","
        sSQL = sSQL & CInt(mvarbPRH_SUP_ManLeft) & ","
        sSQL = sSQL & CInt(mvarbPRH_POL_FullMouth) & ","
        sSQL = sSQL & CInt(mvarbPRH_POL_MaxRight) & ","
        sSQL = sSQL & CInt(mvarbPRH_POL_MaxAnt) & ","
        sSQL = sSQL & CInt(mvarbPRH_POL_MaxLeft) & ","
        sSQL = sSQL & CInt(mvarbPRH_POL_ManLeft) & ","
        sSQL = sSQL & CInt(mvarbPRH_POL_ManRight) & ","
        sSQL = sSQL & CInt(mvarbPRH_POL_ManAnt) & ","
        sSQL = sSQL & CInt(mvarbHG_Scaling_ManLeft_1To3) & ","
        sSQL = sSQL & CInt(mvarbHG_Scaling_ManLeft_4Plus) & ","
        sSQL = sSQL & CInt(mvarbHG_Scaling_ManRight_1To3) & ","
        sSQL = sSQL & CInt(mvarbHG_Scaling_ManRight_4Plus) & ","
        sSQL = sSQL & CInt(mvarbHG_Scaling_MaxLeft_1To3) & ","
        sSQL = sSQL & CInt(mvarbHG_Scaling_MaxLeft_4Plus) & ","
        sSQL = sSQL & CInt(mvarbHG_Scaling_MaxRight_1To3) & ","
        sSQL = sSQL & CInt(mvarbHG_Scaling_MaxRight_4Plus) & ","
        sSQL = sSQL & CInt(mvarbHG_DM_FullMouth) & ","
        sSQL = sSQL & CInt(mvarbHG_DM_Perio) & ","
        sSQL = sSQL & CInt(mvarbHG_DM_Implant) & ","
        sSQL = sSQL & CInt(mvarbHG_MED_Actisite) & ","
        sSQL = sSQL & CInt(mvarbHG_MED_Atridox) & ","
        sSQL = sSQL & CInt(mvarbHG_MED_Chlorhexidine) & ","
        sSQL = sSQL & CInt(mvarbHG_MED_Doxycycline) & ","
        sSQL = sSQL & CInt(mvarbHG_MED_Fluoride_Sodium) & ","
        sSQL = sSQL & CInt(mvarbHG_MED_Fluoride_Stannous) & ","
        sSQL = sSQL & CInt(mvarbHG_MED_Fluoride_Varnish) & ","
        sSQL = sSQL & CInt(mvarbHG_MED_Hydrogen) & ","
        sSQL = sSQL & CInt(mvarbHG_MED_Painfree) & ","
        sSQL = sSQL & CInt(mvarbHG_MED_Perio) & ","
        sSQL = sSQL & CInt(mvarbHG_MED_Potassium) & ","
        'KDunne 4/25/2002
        'SCR 23373
        sSQL = sSQL & CInt(mblnHG_MED_Potassium_Ox) & ","
        'End scr 23373
        
        sSQL = sSQL & CInt(mvarbHG_MED_Saline) & ","
        sSQL = sSQL & CInt(mvarbHG_MED_Tetracycline) & ","
        sSQL = sSQL & CInt(mvarbHG_MED_Topical) & ","
        sSQL = sSQL & CInt(mvarbHG_MED_Duraphet) & ","
        sSQL = sSQL & "'" & RQ(mvarsPRV_OH_Note) & "'" & ","
        sSQL = sSQL & "'" & RQ(mvarsPRV_WC_Tobacco_Note) & "'" & ","
        sSQL = sSQL & "'" & RQ(mvarsPRV_WC_Nutritional_Note) & "'" & ","
        sSQL = sSQL & "'" & RQ(mvarsPRV_Preventive_Comments) & "'" & ","
        sSQL = sSQL & "'" & RQ(mvarsPRH_Prophylaxis_Comments) & "'" & ","
        sSQL = sSQL & "'" & RQ(mvarsHG_DM_Note) & "'" & ","
        sSQL = sSQL & "'" & RQ(mvarsHG_MED_Note) & "'" & ","
        sSQL = sSQL & CInt(mvarlWLID_Nutritional) & ","
        sSQL = sSQL & CInt(mvarlWLID_Tobacco) & ","
        
        sSQL = sSQL & CInt(mvarlWLID_OHDemo) & ","
        sSQL = sSQL & CInt(mvarlWLID_Impress) & ","
        sSQL = sSQL & CInt(mvarlWLID_Proph) & ","
        sSQL = sSQL & CInt(mvarlWLID_Bleaching) & ","
        sSQL = sSQL & CInt(mvarlWLID_Athletic) & ","
        sSQL = sSQL & CInt(mvarlWLID_Fluoride) & ","
        sSQL = sSQL & CInt(mvarlWLID_Scaling_1To3) & ","
        sSQL = sSQL & CInt(mvarlWLID_Scaling_4Plus) & ","
        sSQL = sSQL & CInt(mvarlWLID_FullMouth) & ","
        sSQL = sSQL & CInt(mvarlWLID_Perio) & ","
        sSQL = sSQL & CInt(mvarlWLID_Implant) & ","
        sSQL = sSQL & CInt(mvarlWLID_SiteSpecific) & ","
        sSQL = sSQL & CInt(mvarlWLID_DAgent)
        
        sSQL = sSQL & ")"

    ElseIf dfDirtyFlag = dfUpdate Then
        'The record exists, so create an update statement
        sSQL = ""
        
        sSQL = "UPDATE TEMPL_PREVENTIVE SET "
        sSQL = sSQL & "PRV_OH_Brushing = " & CInt(mvarbPRV_OH_Brushing) & ", "
        sSQL = sSQL & "PRV_OH_Flossing = " & CInt(mvarbPRV_OH_Flossing) & ", "
        sSQL = sSQL & "PRV_OH_Note = " & "'" & RQ(mvarsPRV_OH_Note) & "'" & ", "
        sSQL = sSQL & "PRV_WC_Nutritional = " & CInt(mvarbPRV_WC_Nutritional) & ", "
        sSQL = sSQL & "PRV_WC_Tobacco = " & CInt(mvarbPRV_WC_Tobacco) & ", "
        sSQL = sSQL & "PRV_WC_Tobacco_Note = " & "'" & RQ(mvarsPRV_WC_Tobacco_Note) & "'" & ", "
        sSQL = sSQL & "PRV_WC_Nutritional_Note = " & "'" & RQ(mvarsPRV_WC_Nutritional_Note) & "'" & ", "
        sSQL = sSQL & "PRV_DL_MaxBleaching = " & CInt(mvarbPRV_DL_MaxBleaching) & ", "
        sSQL = sSQL & "PRV_DL_MaxAthletic = " & CInt(mvarbPRV_DL_MaxAthletic) & ", "
        sSQL = sSQL & "PRV_DL_MaxFluoride = " & CInt(mvarbPRV_DL_MaxFluoride) & ", "
        sSQL = sSQL & "PRV_DL_ManBleaching = " & CInt(mvarbPRV_DL_ManBleaching) & ", "
        sSQL = sSQL & "PRV_DL_ManAthletic = " & CInt(mvarbPRV_DL_ManAthletic) & ", "
        sSQL = sSQL & "PRV_DL_ManFluoride = " & CInt(mvarbPRV_DL_ManFluoride) & ", "
        sSQL = sSQL & "PRV_IM_MaxMouthGuard = " & CInt(mvarbPRV_IM_MaxMouthGuard) & ", "
        sSQL = sSQL & "PRV_IM_MaxFluoride = " & CInt(mvarbPRV_IM_MaxFluoride) & ", "
        sSQL = sSQL & "PRV_IM_ManMouthGuard = " & CInt(mvarbPRV_IM_ManMouthGuard) & ", "
        sSQL = sSQL & "PRV_IM_ManFluoride = " & CInt(mvarbPRV_IM_ManFluoride) & ", "
        sSQL = sSQL & "PRV_IM_MaxBleach = " & CInt(mvarbPRV_IM_MaxBleach) & ", "
        sSQL = sSQL & "PRV_IM_ManBleach = " & CInt(mvarbPRV_IM_ManBleach) & ", "
        sSQL = sSQL & "PRV_IM_MaxDiag = " & CInt(mvarbPRV_IM_MaxDiag) & ", "
        sSQL = sSQL & "PRV_IM_ManDiag = " & CInt(mvarbPRV_IM_ManDiag) & ", "
        sSQL = sSQL & "PRV_Preventive_Comments = " & "'" & RQ(mvarsPRV_Preventive_Comments) & "'" & ", "
        sSQL = sSQL & "PRH_Proph_Complete = " & CInt(mvarbPRH_PROPH_Complete) & ", "
        sSQL = sSQL & "PRH_Proph_Incomplete = " & CInt(mvarbPRH_PROPH_Incomplete) & ", "
        sSQL = sSQL & "PRH_Proph_Fluoride_SNF2 = " & CInt(mvarbPRH_PROPH_Fluoride_SNF2) & ", "
        sSQL = sSQL & "PRH_Proph_Fluoride_APF = " & CInt(mvarbPRH_PROPH_Fluoride_APF) & ", "
        sSQL = sSQL & "PRH_Proph_Fluoride_NAF = " & CInt(mvarbPRH_PROPH_Fluoride_NAF) & ", "
        sSQL = sSQL & "PRH_Proph_Fluoride_Varnish = " & CInt(mvarbPRH_PROPH_Fluoride_Varnish) & ", "
        sSQL = sSQL & "PRH_Sup_FullMouth = " & CInt(mvarbPRH_SUP_FullMouth) & ", "
        sSQL = sSQL & "PRH_Sup_MaxRight = " & CInt(mvarbPRH_SUP_MaxRight) & ", "
        sSQL = sSQL & "PRH_Sup_ManRight = " & CInt(mvarbPRH_SUP_ManRight) & ", "
        sSQL = sSQL & "PRH_Sup_MaxAnt = " & CInt(mvarbPRH_SUP_MaxAnt) & ", "
        sSQL = sSQL & "PRH_Sup_ManAnt = " & CInt(mvarbPRH_Sup_ManAnt) & ", "
        sSQL = sSQL & "PRH_Sup_MaxLeft = " & CInt(mvarbPRH_SUP_MaxLeft) & ", "
        sSQL = sSQL & "PRH_Sup_ManLeft = " & CInt(mvarbPRH_SUP_ManLeft) & ", "
        sSQL = sSQL & "PRH_Pol_FullMouth = " & CInt(mvarbPRH_POL_FullMouth) & ", "
        sSQL = sSQL & "PRH_Pol_MaxRight = " & CInt(mvarbPRH_POL_MaxRight) & ", "
        sSQL = sSQL & "PRH_Pol_MaxAnt = " & CInt(mvarbPRH_POL_MaxAnt) & ", "
        sSQL = sSQL & "PRH_Pol_MaxLeft = " & CInt(mvarbPRH_POL_MaxLeft) & ", "
        sSQL = sSQL & "PRH_Pol_ManLeft = " & CInt(mvarbPRH_POL_ManLeft) & ", "
        sSQL = sSQL & "PRH_Pol_ManRight = " & CInt(mvarbPRH_POL_ManRight) & ", "
        sSQL = sSQL & "PRH_Pol_ManAnt = " & CInt(mvarbPRH_POL_ManAnt) & ", "
        sSQL = sSQL & "PRH_Prophylaxis_Comments = " & "'" & RQ(mvarsPRH_Prophylaxis_Comments) & "'" & ", "
        sSQL = sSQL & "HG_Scaling_ManLeft_1To3 = " & CInt(mvarbHG_Scaling_ManLeft_1To3) & ", "
        sSQL = sSQL & "HG_Scaling_ManLeft_4Plus = " & CInt(mvarbHG_Scaling_ManLeft_4Plus) & ", "
        sSQL = sSQL & "HG_Scaling_ManRight_1To3 = " & CInt(mvarbHG_Scaling_ManRight_1To3) & ", "
        sSQL = sSQL & "HG_Scaling_ManRight_4Plus = " & CInt(mvarbHG_Scaling_ManRight_4Plus) & ", "
        sSQL = sSQL & "HG_Scaling_MaxLeft_1To3 = " & CInt(mvarbHG_Scaling_MaxLeft_1To3) & ", "
        sSQL = sSQL & "HG_Scaling_MaxLeft_4Plus = " & CInt(mvarbHG_Scaling_MaxLeft_4Plus) & ", "
        sSQL = sSQL & "HG_Scaling_MaxRight_1To3 = " & CInt(mvarbHG_Scaling_MaxRight_1To3) & ", "
        sSQL = sSQL & "HG_Scaling_MaxRight_4Plus = " & CInt(mvarbHG_Scaling_MaxRight_4Plus) & ", "
        sSQL = sSQL & "HG_DM_FullMouth = " & CInt(mvarbHG_DM_FullMouth) & ", "
        sSQL = sSQL & "HG_DM_Perio = " & CInt(mvarbHG_DM_Perio) & ", "
        sSQL = sSQL & "HG_DM_Implant = " & CInt(mvarbHG_DM_Implant) & ", "
        sSQL = sSQL & "HG_DM_Note = " & "'" & RQ(mvarsHG_DM_Note) & "'" & ", "
        sSQL = sSQL & "HG_Med_Actisite = " & CInt(mvarbHG_MED_Actisite) & ", "
        sSQL = sSQL & "HG_Med_Atridox = " & CInt(mvarbHG_MED_Atridox) & ", "
        sSQL = sSQL & "HG_Med_Chlorhexidine = " & CInt(mvarbHG_MED_Chlorhexidine) & ", "
        sSQL = sSQL & "HG_Med_Doxycycline = " & CInt(mvarbHG_MED_Doxycycline) & ", "
        sSQL = sSQL & "HG_Med_Fluoride_Sodium = " & CInt(mvarbHG_MED_Fluoride_Sodium) & ", "
        sSQL = sSQL & "HG_Med_Fluoride_Stannous = " & CInt(mvarbHG_MED_Fluoride_Stannous) & ", "
        sSQL = sSQL & "HG_Med_Fluoride_Varnish = " & CInt(mvarbHG_MED_Fluoride_Varnish) & ", "
        sSQL = sSQL & "HG_Med_Hydrogen = " & CInt(mvarbHG_MED_Hydrogen) & ", "
        sSQL = sSQL & "HG_MED_Painfree = " & CInt(mvarbHG_MED_Painfree) & ", "
        sSQL = sSQL & "HG_Med_Perio = " & CInt(mvarbHG_MED_Perio) & ", "
        sSQL = sSQL & "HG_Med_Potassium = " & CInt(mvarbHG_MED_Potassium) & ", "
        'KDunne 4/25/2002
        'SCR 23373
        sSQL = sSQL & "HG_MED_Potassium_Ox = " & CInt(mblnHG_MED_Potassium_Ox) & ", "
        'End Scr 23373
        
        sSQL = sSQL & "HG_Med_Saline = " & CInt(mvarbHG_MED_Saline) & ", "
        sSQL = sSQL & "HG_Med_Tetracycline = " & CInt(mvarbHG_MED_Tetracycline) & ", "
        sSQL = sSQL & "HG_Med_Topical = " & CInt(mvarbHG_MED_Topical) & ", "
        sSQL = sSQL & "HG_Med_Duraphet = " & CInt(mvarbHG_MED_Duraphet) & ", "
        sSQL = sSQL & "HG_Med_Note = '" & RQ(mvarsHG_MED_Note) & "' " & ", "
        
        sSQL = sSQL & "WID_Nutritional = " & CInt(mvarlWLID_Nutritional) & ", "
        sSQL = sSQL & "WID_Tobacco = " & CInt(mvarlWLID_Tobacco) & ", "
        
        sSQL = sSQL & "WID_OHDemo = " & CInt(mvarlWLID_OHDemo) & ", "
        sSQL = sSQL & "WID_Impress = " & CInt(mvarlWLID_Impress) & ", "
        sSQL = sSQL & "WID_Proph = " & CInt(mvarlWLID_Proph) & ", "
        sSQL = sSQL & "WID_Bleaching = " & CInt(mvarlWLID_Bleaching) & ", "
        sSQL = sSQL & "WID_Athletic = " & CInt(mvarlWLID_Athletic) & ", "
        sSQL = sSQL & "WID_Fluoride = " & CInt(mvarlWLID_Fluoride) & ", "
        sSQL = sSQL & "WID_Scaling_1TO3 = " & CInt(mvarlWLID_Scaling_1To3) & ", "
        sSQL = sSQL & "WID_Scaling_4Plus = " & CInt(mvarlWLID_Scaling_4Plus) & ", "
        sSQL = sSQL & "WID_FullMouth = " & CInt(mvarlWLID_FullMouth) & ", "
        sSQL = sSQL & "WID_Perio = " & CInt(mvarlWLID_Perio) & ", "
        sSQL = sSQL & "WID_Implant = " & CInt(mvarlWLID_Implant) & ", "
        sSQL = sSQL & "WID_SiteSpecific = " & CInt(mvarlWLID_SiteSpecific) & ", "
        sSQL = sSQL & "WID_DAgent = " & CInt(mvarlWLID_DAgent) & " "
        
        sSQL = sSQL & "WHERE (DENTAL_ENCOUNTER_ID = " & mvarlDentalEncounterID & ")"
    
    ElseIf dfDirtyFlag = dfdelete Then
        'Mark the record as deleted
        sSQL = ""
    End If
    
   ' Debug.Assert False
    
    goconn.ExecuteSQL sSQL
    
    bSavedToDB = True
    dfDirtyFlag = dfNoAction
    
    clsISection_SaveSection = True

Exit Function

ErrHnd:
    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & "." & sPROC_NAME, App.Title, vbCritical
    
    'Delete the workload credit if it didn't update correctly
    'UpdateWorkload mvarlWLID_Nutritional, dfdelete, mvarlDentalEncounterID, "", 1

End Function


'*******************************************************************
'Function: LoadSection
'Author: ASamedov
'Date: 11/30/01
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
    
    If gbUseStoredProc = True Then
        Dim oCmd As New CHCSII_Command.CHCSII_Cmd
        
        oCmd.SPName = "Dent_Notes_Preventive_Load"

        Call oCmd.AddParam(False, False, "textString", mvarlDentalEncounterID)
        Call oCmd.AddParam(True, True, "", "")
        
        Set rs = goconn.OpenRecordsetSP(oCmd)
        
        'Preventive and Prophylaxis/Fluoride sections
        If rs.EOF Then
            bSavedToDB = False
            dfDirtyFlag = dfAdd
            'Init the values here to "" ("False" is default for boolean)
            mvarlWLID_Nutritional = 0
            mvarlWLID_Tobacco = 0
            mvarlWLID_OHDemo = 0
            mvarlWLID_Impress = 0
            mvarlWLID_Proph = 0
            mvarlWLID_Bleaching = 0
            mvarlWLID_Athletic = 0
            mvarlWLID_Fluoride = 0
            mvarlWLID_Scaling_1To3 = 0
            mvarlWLID_Scaling_4Plus = 0
            mvarlWLID_FullMouth = 0
            mvarlWLID_Perio = 0
            mvarlWLID_Implant = 0
            mvarlWLID_SiteSpecific = 0
            mvarlWLID_DAgent = 0
            
            mvarsPRV_OH_Note = ""
            mvarsPRV_WC_Tobacco_Note = ""
            mvarsPRV_WC_Nutritional_Note = ""
            mvarsPRV_Preventive_Comments = ""
            mvarsPRH_Prophylaxis_Comments = ""
            mvarsHG_DM_Note = ""
            mvarsHG_MED_Note = ""
        Else
            bSavedToDB = True
            dfDirtyFlag = dfNoAction
            
            'Set the values from the db including "" to allow for nulls
            'mvarlWorkloadID = Val(rs("Workload_ID") & "")
            
            mvarbPRV_OH_Brushing = CBool(Val(rs("PRV_OH_Brushing") & ""))
            mvarbPRV_OH_Flossing = CBool(Val(rs("PRV_OH_Flossing") & ""))
            mvarbPRV_WC_Nutritional = CBool(Val(rs("PRV_WC_Nutritional") & ""))
            mvarbPRV_WC_Tobacco = CBool(Val(rs("PRV_WC_Tobacco") & ""))
            mvarbPRV_DL_MaxBleaching = CBool(Val(rs("PRV_DL_MaxBleaching") & ""))
            mvarbPRV_DL_MaxAthletic = CBool(Val(rs("PRV_DL_MaxAthletic") & ""))
            mvarbPRV_DL_MaxFluoride = CBool(Val(rs("PRV_DL_MaxFluoride") & ""))
            mvarbPRV_DL_ManBleaching = CBool(Val(rs("PRV_DL_ManBleaching") & ""))
            mvarbPRV_DL_ManAthletic = CBool(Val(rs("PRV_DL_ManAthletic") & ""))
            mvarbPRV_DL_ManFluoride = CBool(Val(rs("PRV_DL_ManFluoride") & ""))
            mvarbPRV_IM_MaxMouthGuard = CBool(Val(rs("PRV_IM_MaxMouthGuard") & ""))
            mvarbPRV_IM_MaxFluoride = CBool(Val(rs("PRV_IM_MaxFluoride") & ""))
            mvarbPRV_IM_ManMouthGuard = CBool(Val(rs("PRV_IM_ManMouthGuard") & ""))
            mvarbPRV_IM_ManFluoride = CBool(Val(rs("PRV_IM_ManFluoride") & ""))
            mvarbPRV_IM_MaxBleach = CBool(Val(rs("PRV_IM_MaxBleach") & ""))
            mvarbPRV_IM_ManBleach = CBool(Val(rs("PRV_IM_ManBleach") & ""))
            mvarbPRV_IM_MaxDiag = CBool(Val(rs("PRV_IM_MaxDiag") & ""))
            mvarbPRV_IM_ManDiag = CBool(Val(rs("PRV_IM_ManDiag") & ""))
            mvarbPRH_PROPH_Complete = CBool(Val(rs("PRH_Proph_Complete") & ""))
            mvarbPRH_PROPH_Incomplete = CBool(Val(rs("PRH_Proph_Incomplete") & ""))
            mvarbPRH_PROPH_Fluoride_SNF2 = CBool(Val(rs("PRH_Proph_Fluoride_SNF2") & ""))
            mvarbPRH_PROPH_Fluoride_APF = CBool(Val(rs("PRH_Proph_Fluoride_APF") & ""))
            mvarbPRH_PROPH_Fluoride_NAF = CBool(Val(rs("PRH_Proph_Fluoride_NAF") & ""))
            mvarbPRH_PROPH_Fluoride_Varnish = CBool(Val(rs("PRH_Proph_Fluoride_Varnish") & ""))
            mvarbPRH_SUP_FullMouth = CBool(Val(rs("PRH_Sup_FullMouth") & ""))
            mvarbPRH_SUP_MaxRight = CBool(Val(rs("PRH_Sup_MaxRight") & ""))
            mvarbPRH_SUP_ManRight = CBool(Val(rs("PRH_Sup_ManRight") & ""))
            mvarbPRH_SUP_MaxAnt = CBool(Val(rs("PRH_Sup_MaxAnt") & ""))
            mvarbPRH_Sup_ManAnt = CBool(Val(rs("PRH_Sup_ManAnt") & ""))
            mvarbPRH_SUP_MaxLeft = CBool(Val(rs("PRH_Sup_MaxLeft") & ""))
            mvarbPRH_SUP_ManLeft = CBool(Val(rs("PRH_Sup_ManLeft") & ""))
            mvarbPRH_POL_FullMouth = CBool(Val(rs("PRH_Pol_FullMouth") & ""))
            mvarbPRH_POL_MaxRight = CBool(Val(rs("PRH_Pol_MaxRight") & ""))
            mvarbPRH_POL_MaxAnt = CBool(Val(rs("PRH_Pol_MaxAnt") & ""))
            mvarbPRH_POL_MaxLeft = CBool(Val(rs("PRH_Pol_MaxLeft") & ""))
            mvarbPRH_POL_ManLeft = CBool(Val(rs("PRH_Pol_ManLeft") & ""))
            mvarbPRH_POL_ManRight = CBool(Val(rs("PRH_Pol_ManRight") & ""))
            mvarbPRH_POL_ManAnt = CBool(Val(rs("PRH_Pol_ManAnt") & ""))
            mvarsPRV_OH_Note = rs("PRV_OH_Note") & ""
        End If
        rs.Close
        
        Set oCmd = New CHCSII_Command.CHCSII_Cmd
        
        oCmd.SPName = "Dent_Notes_Preventive2_Load"

        Call oCmd.AddParam(False, False, "textString", mvarlDentalEncounterID)
        Call oCmd.AddParam(True, True, "", "")
        
        Set rs = goconn.OpenRecordsetSP(oCmd)
        
        If Not rs.EOF Then
            mvarbHG_Scaling_ManLeft_1To3 = CBool(Val(rs("HG_Scaling_ManLeft_1To3") & ""))
            mvarbHG_Scaling_ManLeft_4Plus = CBool(Val(rs("HG_Scaling_ManLeft_4Plus") & ""))
            mvarbHG_Scaling_ManRight_1To3 = CBool(Val(rs("HG_Scaling_ManRight_1To3") & ""))
            mvarbHG_Scaling_ManRight_4Plus = CBool(Val(rs("HG_Scaling_ManRight_4Plus") & ""))
            mvarbHG_Scaling_MaxLeft_1To3 = CBool(Val(rs("HG_Scaling_MaxLeft_1To3") & ""))
            mvarbHG_Scaling_MaxLeft_4Plus = CBool(Val(rs("HG_Scaling_MaxLeft_4Plus") & ""))
            mvarbHG_Scaling_MaxRight_1To3 = CBool(Val(rs("HG_Scaling_MaxRight_1To3") & ""))
            mvarbHG_Scaling_MaxRight_4Plus = CBool(Val(rs("HG_Scaling_MaxRight_4plus") & ""))
            mvarbHG_DM_FullMouth = CBool(Val(rs("HG_DM_FullMouth") & ""))
            mvarbHG_DM_Perio = CBool(Val(rs("HG_DM_Perio") & ""))
            mvarbHG_DM_Implant = CBool(Val(rs("HG_DM_Implant") & ""))
            mvarbHG_MED_Actisite = CBool(Val(rs("HG_Med_Actisite") & ""))
            mvarbHG_MED_Atridox = CBool(Val(rs("HG_Med_Atridox") & ""))
            mvarbHG_MED_Chlorhexidine = CBool(Val(rs("HG_Med_Chlorhexidine") & ""))
            mvarbHG_MED_Doxycycline = CBool(Val(rs("HG_Med_Doxycycline") & ""))
            mvarbHG_MED_Fluoride_Sodium = CBool(Val(rs("HG_Med_Fluoride_Sodium") & ""))
            mvarbHG_MED_Fluoride_Stannous = CBool(Val(rs("HG_Med_Fluoride_Stannous") & ""))
            mvarbHG_MED_Fluoride_Varnish = CBool(Val(rs("HG_Med_Fluoride_Varnish") & ""))
            mvarbHG_MED_Hydrogen = CBool(Val(rs("HG_Med_Hydrogen") & ""))
            mvarbHG_MED_Painfree = CBool(Val(rs("HG_MED_Painfree") & ""))
            mvarbHG_MED_Perio = CBool(Val(rs("HG_Med_Perio") & ""))
            mvarbHG_MED_Potassium = CBool(Val(rs("HG_Med_Potassium") & ""))
            mvarbHG_MED_Saline = CBool(Val(rs("HG_Med_Saline") & ""))
            mvarbHG_MED_Tetracycline = CBool(Val(rs("HG_Med_Tetracycline") & ""))
            mvarbHG_MED_Topical = CBool(Val(rs("HG_Med_Topical") & ""))
            mvarbHG_MED_Duraphet = CBool(Val(rs("HG_Med_Duraphet") & ""))
            'KDunne 4/25/2002
            'SCR 23373
            mblnHG_MED_Potassium_Ox = CBool(Val(rs("HG_MED_POTASSIUM_OX") & ""))
        
            mvarsPRV_WC_Tobacco_Note = rs("PRV_WC_Tobacco_Note") & ""
            mvarsPRV_WC_Nutritional_Note = rs("PRV_WC_Nutritional_Note") & ""
        
            mvarsPRV_Preventive_Comments = rs("PRV_Preventive_Comments") & ""
            mvarsPRH_Prophylaxis_Comments = rs("PRH_Prophylaxis_Comments") & ""
        
            mvarsHG_DM_Note = rs("HG_DM_Note") & ""
            mvarsHG_MED_Note = rs("HG_Med_Note") & ""
        
            mvarlWLID_Nutritional = Val(rs("WID_NUTRITIONAL") & "")
            mvarlWLID_Tobacco = Val(rs("WID_TOBACCO") & "")
            mvarlWLID_OHDemo = Val(rs("WID_OHDEMO") & "")
            mvarlWLID_Impress = Val(rs("WID_IMPRESS") & "")
            mvarlWLID_Proph = Val(rs("WID_PROPH") & "")
            mvarlWLID_Bleaching = Val(rs("WID_BLEACHING") & "")
            mvarlWLID_Athletic = Val(rs("WID_ATHLETIC") & "")
            mvarlWLID_Fluoride = Val(rs("WID_FLUORIDE") & "")
            mvarlWLID_Scaling_1To3 = Val(rs("WID_SCALING_1To3") & "")
            mvarlWLID_Scaling_4Plus = Val(rs("WID_SCALING_4PLUS") & "")
            mvarlWLID_FullMouth = Val(rs("WID_FULLMOUTH") & "")
            mvarlWLID_Perio = Val(rs("WID_PERIO") & "")
            mvarlWLID_Implant = Val(rs("WID_IMPLANT") & "")
            mvarlWLID_SiteSpecific = Val(rs("WID_SITESPECIFIC") & "")
            mvarlWLID_DAgent = Val(rs("WID_DAGENT") & "")
        End If
        
        rs.Close
        'End of Workload section
        
    Else
    
        'Preventive and Prophylaxis/Fluoride sections
        sSQL = "SELECT "
        sSQL = sSQL & cSQL_Select1
        sSQL = sSQL & ",PRV_OH_Note"
        sSQL = sSQL & " FROM TEMPL_PREVENTIVE"
        sSQL = sSQL & " WHERE TEMPL_PREVENTIVE.DENTAL_ENCOUNTER_ID = " & mvarlDentalEncounterID
            
        Set rs = goconn.OpenRecordset(sSQL)
        
        If rs.EOF Then
            bSavedToDB = False
            dfDirtyFlag = dfAdd
            'Init the values here to "" ("False" is default for boolean)
            mvarlWLID_Nutritional = 0
            mvarlWLID_Tobacco = 0
            mvarlWLID_OHDemo = 0
            mvarlWLID_Impress = 0
            mvarlWLID_Proph = 0
            mvarlWLID_Bleaching = 0
            mvarlWLID_Athletic = 0
            mvarlWLID_Fluoride = 0
            mvarlWLID_Scaling_1To3 = 0
            mvarlWLID_Scaling_4Plus = 0
            mvarlWLID_FullMouth = 0
            mvarlWLID_Perio = 0
            mvarlWLID_Implant = 0
            mvarlWLID_SiteSpecific = 0
            mvarlWLID_DAgent = 0
            
            mvarsPRV_OH_Note = ""
            mvarsPRV_WC_Tobacco_Note = ""
            mvarsPRV_WC_Nutritional_Note = ""
            mvarsPRV_Preventive_Comments = ""
            mvarsPRH_Prophylaxis_Comments = ""
            mvarsHG_DM_Note = ""
            mvarsHG_MED_Note = ""
        Else
            bSavedToDB = True
            dfDirtyFlag = dfNoAction
            
            'Set the values from the db including "" to allow for nulls
            'mvarlWorkloadID = Val(rs("Workload_ID") & "")
            
            mvarbPRV_OH_Brushing = CBool(Val(rs("PRV_OH_Brushing") & ""))
            mvarbPRV_OH_Flossing = CBool(Val(rs("PRV_OH_Flossing") & ""))
            mvarbPRV_WC_Nutritional = CBool(Val(rs("PRV_WC_Nutritional") & ""))
            mvarbPRV_WC_Tobacco = CBool(Val(rs("PRV_WC_Tobacco") & ""))
            mvarbPRV_DL_MaxBleaching = CBool(Val(rs("PRV_DL_MaxBleaching") & ""))
            mvarbPRV_DL_MaxAthletic = CBool(Val(rs("PRV_DL_MaxAthletic") & ""))
            mvarbPRV_DL_MaxFluoride = CBool(Val(rs("PRV_DL_MaxFluoride") & ""))
            mvarbPRV_DL_ManBleaching = CBool(Val(rs("PRV_DL_ManBleaching") & ""))
            mvarbPRV_DL_ManAthletic = CBool(Val(rs("PRV_DL_ManAthletic") & ""))
            mvarbPRV_DL_ManFluoride = CBool(Val(rs("PRV_DL_ManFluoride") & ""))
            mvarbPRV_IM_MaxMouthGuard = CBool(Val(rs("PRV_IM_MaxMouthGuard") & ""))
            mvarbPRV_IM_MaxFluoride = CBool(Val(rs("PRV_IM_MaxFluoride") & ""))
            mvarbPRV_IM_ManMouthGuard = CBool(Val(rs("PRV_IM_ManMouthGuard") & ""))
            mvarbPRV_IM_ManFluoride = CBool(Val(rs("PRV_IM_ManFluoride") & ""))
            mvarbPRV_IM_MaxBleach = CBool(Val(rs("PRV_IM_MaxBleach") & ""))
            mvarbPRV_IM_ManBleach = CBool(Val(rs("PRV_IM_ManBleach") & ""))
            mvarbPRV_IM_MaxDiag = CBool(Val(rs("PRV_IM_MaxDiag") & ""))
            mvarbPRV_IM_ManDiag = CBool(Val(rs("PRV_IM_ManDiag") & ""))
            mvarbPRH_PROPH_Complete = CBool(Val(rs("PRH_Proph_Complete") & ""))
            mvarbPRH_PROPH_Incomplete = CBool(Val(rs("PRH_Proph_Incomplete") & ""))
            mvarbPRH_PROPH_Fluoride_SNF2 = CBool(Val(rs("PRH_Proph_Fluoride_SNF2") & ""))
            mvarbPRH_PROPH_Fluoride_APF = CBool(Val(rs("PRH_Proph_Fluoride_APF") & ""))
            mvarbPRH_PROPH_Fluoride_NAF = CBool(Val(rs("PRH_Proph_Fluoride_NAF") & ""))
            mvarbPRH_PROPH_Fluoride_Varnish = CBool(Val(rs("PRH_Proph_Fluoride_Varnish") & ""))
            mvarbPRH_SUP_FullMouth = CBool(Val(rs("PRH_Sup_FullMouth") & ""))
            mvarbPRH_SUP_MaxRight = CBool(Val(rs("PRH_Sup_MaxRight") & ""))
            mvarbPRH_SUP_ManRight = CBool(Val(rs("PRH_Sup_ManRight") & ""))
            mvarbPRH_SUP_MaxAnt = CBool(Val(rs("PRH_Sup_MaxAnt") & ""))
            mvarbPRH_Sup_ManAnt = CBool(Val(rs("PRH_Sup_ManAnt") & ""))
            mvarbPRH_SUP_MaxLeft = CBool(Val(rs("PRH_Sup_MaxLeft") & ""))
            mvarbPRH_SUP_ManLeft = CBool(Val(rs("PRH_Sup_ManLeft") & ""))
            mvarbPRH_POL_FullMouth = CBool(Val(rs("PRH_Pol_FullMouth") & ""))
            mvarbPRH_POL_MaxRight = CBool(Val(rs("PRH_Pol_MaxRight") & ""))
            mvarbPRH_POL_MaxAnt = CBool(Val(rs("PRH_Pol_MaxAnt") & ""))
            mvarbPRH_POL_MaxLeft = CBool(Val(rs("PRH_Pol_MaxLeft") & ""))
            mvarbPRH_POL_ManLeft = CBool(Val(rs("PRH_Pol_ManLeft") & ""))
            mvarbPRH_POL_ManRight = CBool(Val(rs("PRH_Pol_ManRight") & ""))
            mvarbPRH_POL_ManAnt = CBool(Val(rs("PRH_Pol_ManAnt") & ""))
            mvarsPRV_OH_Note = rs("PRV_OH_Note") & ""
        End If
            
        rs.Close
            
        'Hygiene section
        sSQL = "SELECT "
        sSQL = sSQL & cSQL_Select2
        
        sSQL = sSQL & " FROM TEMPL_PREVENTIVE"
        sSQL = sSQL & " WHERE TEMPL_PREVENTIVE.DENTAL_ENCOUNTER_ID = " & mvarlDentalEncounterID
            
        Set rs = goconn.OpenRecordset(sSQL)
        If rs.EOF Then
            dfDirtyFlag = dfAdd
        Else
            dfDirtyFlag = dfNoAction
            mvarbHG_Scaling_ManLeft_1To3 = CBool(Val(rs("HG_Scaling_ManLeft_1To3") & ""))
            mvarbHG_Scaling_ManLeft_4Plus = CBool(Val(rs("HG_Scaling_ManLeft_4Plus") & ""))
            mvarbHG_Scaling_ManRight_1To3 = CBool(Val(rs("HG_Scaling_ManRight_1To3") & ""))
            mvarbHG_Scaling_ManRight_4Plus = CBool(Val(rs("HG_Scaling_ManRight_4Plus") & ""))
            mvarbHG_Scaling_MaxLeft_1To3 = CBool(Val(rs("HG_Scaling_MaxLeft_1To3") & ""))
            mvarbHG_Scaling_MaxLeft_4Plus = CBool(Val(rs("HG_Scaling_MaxLeft_4Plus") & ""))
            mvarbHG_Scaling_MaxRight_1To3 = CBool(Val(rs("HG_Scaling_MaxRight_1To3") & ""))
            mvarbHG_Scaling_MaxRight_4Plus = CBool(Val(rs("HG_Scaling_MaxRight_4Plus") & ""))
            mvarbHG_DM_FullMouth = CBool(Val(rs("HG_DM_FullMouth") & ""))
            mvarbHG_DM_Perio = CBool(Val(rs("HG_DM_Perio") & ""))
            mvarbHG_DM_Implant = CBool(Val(rs("HG_DM_Implant") & ""))
            mvarbHG_MED_Actisite = CBool(Val(rs("HG_Med_Actisite") & ""))
            mvarbHG_MED_Atridox = CBool(Val(rs("HG_Med_Atridox") & ""))
            mvarbHG_MED_Chlorhexidine = CBool(Val(rs("HG_Med_Chlorhexidine") & ""))
            mvarbHG_MED_Doxycycline = CBool(Val(rs("HG_Med_Doxycycline") & ""))
            mvarbHG_MED_Fluoride_Sodium = CBool(Val(rs("HG_Med_Fluoride_Sodium") & ""))
            mvarbHG_MED_Fluoride_Stannous = CBool(Val(rs("HG_Med_Fluoride_Stannous") & ""))
            mvarbHG_MED_Fluoride_Varnish = CBool(Val(rs("HG_Med_Fluoride_Varnish") & ""))
            mvarbHG_MED_Hydrogen = CBool(Val(rs("HG_Med_Hydrogen") & ""))
            mvarbHG_MED_Painfree = CBool(Val(rs("HG_MED_Painfree") & ""))
            mvarbHG_MED_Perio = CBool(Val(rs("HG_Med_Perio") & ""))
            mvarbHG_MED_Potassium = CBool(Val(rs("HG_Med_Potassium") & ""))
            mvarbHG_MED_Saline = CBool(Val(rs("HG_Med_Saline") & ""))
            mvarbHG_MED_Tetracycline = CBool(Val(rs("HG_Med_Tetracycline") & ""))
            mvarbHG_MED_Topical = CBool(Val(rs("HG_Med_Topical") & ""))
            mvarbHG_MED_Duraphet = CBool(Val(rs("HG_Med_Duraphet") & ""))
            'KDunne 4/25/2002
            'SCR 23373
            mblnHG_MED_Potassium_Ox = CBool(Val(rs("HG_MED_POTASSIUM_OX") & ""))
            
        End If
        
        rs.Close
        
        'Notes section
        sSQL = "SELECT "
        sSQL = sSQL & "PRV_WC_Tobacco_Note,"
        sSQL = sSQL & "PRV_WC_Nutritional_Note"
        sSQL = sSQL & " FROM TEMPL_PREVENTIVE"
        sSQL = sSQL & " WHERE TEMPL_PREVENTIVE.DENTAL_ENCOUNTER_ID = " & mvarlDentalEncounterID
            
        Set rs = goconn.OpenRecordset(sSQL)
        If rs.EOF Then
            dfDirtyFlag = dfAdd
        Else
            dfDirtyFlag = dfNoAction
            mvarsPRV_WC_Tobacco_Note = rs("PRV_WC_Tobacco_Note") & ""
            mvarsPRV_WC_Nutritional_Note = rs("PRV_WC_Nutritional_Note") & ""
        End If
        
        rs.Close
        
        sSQL = "SELECT "
        sSQL = sSQL & "PRV_Preventive_Comments,"
        sSQL = sSQL & "PRH_Prophylaxis_Comments"
        sSQL = sSQL & " FROM TEMPL_PREVENTIVE"
        sSQL = sSQL & " WHERE TEMPL_PREVENTIVE.DENTAL_ENCOUNTER_ID = " & mvarlDentalEncounterID
        
        Set rs = goconn.OpenRecordset(sSQL)
        If rs.EOF Then
            dfDirtyFlag = dfAdd
        Else
            dfDirtyFlag = dfNoAction
            mvarsPRV_Preventive_Comments = rs("PRV_Preventive_Comments") & ""
            mvarsPRH_Prophylaxis_Comments = rs("PRH_Prophylaxis_Comments") & ""
        End If
        
        rs.Close
            
        sSQL = "SELECT "
        sSQL = sSQL & "HG_DM_Note,"
        sSQL = sSQL & "HG_Med_Note"
        sSQL = sSQL & " FROM TEMPL_PREVENTIVE"
        sSQL = sSQL & " WHERE TEMPL_PREVENTIVE.DENTAL_ENCOUNTER_ID = " & mvarlDentalEncounterID
            
        Set rs = goconn.OpenRecordset(sSQL)
        If rs.EOF Then
            dfDirtyFlag = dfAdd
        Else
            dfDirtyFlag = dfNoAction
            mvarsHG_DM_Note = rs("HG_DM_Note") & ""
            mvarsHG_MED_Note = rs("HG_Med_Note") & ""
        End If
        
        rs.Close
        'End of Notes section
            
        'Workload section
        sSQL = "SELECT "
        sSQL = sSQL & "WID_NUTRITIONAL,"
        sSQL = sSQL & "WID_TOBACCO,"
        sSQL = sSQL & "WID_OHDEMO,"
        sSQL = sSQL & "WID_IMPRESS,"
        sSQL = sSQL & "WID_PROPH,"
        sSQL = sSQL & "WID_BLEACHING,"
        sSQL = sSQL & "WID_ATHLETIC,"
        sSQL = sSQL & "WID_FLUORIDE,"
        sSQL = sSQL & "WID_SCALING_1to3,"
        sSQL = sSQL & "WID_SCALING_4Plus,"
        sSQL = sSQL & "WID_FULLMOUTH,"
        sSQL = sSQL & "WID_PERIO,"
        sSQL = sSQL & "WID_IMPLANT,"
        sSQL = sSQL & "WID_SITESPECIFIC,"
        sSQL = sSQL & "WID_DAGENT"
        
        sSQL = sSQL & " FROM TEMPL_PREVENTIVE"
        sSQL = sSQL & " WHERE TEMPL_PREVENTIVE.DENTAL_ENCOUNTER_ID = " & mvarlDentalEncounterID
            
        Set rs = goconn.OpenRecordset(sSQL)
        If rs.EOF Then
            dfDirtyFlag = dfAdd
        Else
            dfDirtyFlag = dfNoAction
            mvarlWLID_Nutritional = Val(rs("WID_NUTRITIONAL") & "")
            mvarlWLID_Tobacco = Val(rs("WID_TOBACCO") & "")
            mvarlWLID_OHDemo = Val(rs("WID_OHDEMO") & "")
            mvarlWLID_Impress = Val(rs("WID_IMPRESS") & "")
            mvarlWLID_Proph = Val(rs("WID_PROPH") & "")
            mvarlWLID_Bleaching = Val(rs("WID_BLEACHING") & "")
            mvarlWLID_Athletic = Val(rs("WID_ATHLETIC") & "")
            mvarlWLID_Fluoride = Val(rs("WID_FLUORIDE") & "")
            mvarlWLID_Scaling_1To3 = Val(rs("WID_SCALING_1to3") & "")
            mvarlWLID_Scaling_4Plus = Val(rs("WID_SCALING_4Plus") & "")
            mvarlWLID_FullMouth = Val(rs("WID_FULLMOUTH") & "")
            mvarlWLID_Perio = Val(rs("WID_PERIO") & "")
            mvarlWLID_Implant = Val(rs("WID_IMPLANT") & "")
            mvarlWLID_SiteSpecific = Val(rs("WID_SITESPECIFIC") & "")
            mvarlWLID_DAgent = Val(rs("WID_DAGENT") & "")
              
        End If
        rs.Close
        'End of Workload section
    End If
    
    Set rs = Nothing
   
    clsISection_LoadSection = bReturnResult
    
Exit Function

ErrHnd:
    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & "." & sPROC_NAME, App.Title, vbCritical

End Function

Private Function UpdateDirtyFlag()

    If bSavedToDB = False Then
        dfDirtyFlag = dfAdd
    ElseIf dfDirtyFlag = dfNoAction Then
        dfDirtyFlag = dfUpdate
    End If
      
End Function

Private Sub Class_Initialize()

    dfDirtyFlag = dfNoAction
    
    Nutritional_WL.WorkloadId = 257
    Nutritional_WL.WorkloadCode = "D1310"
    Nutritional_WL.WorkLoadDef = "Nutritional counseling"
        
    Tobacco_WL.WorkloadId = 465
    Tobacco_WL.WorkloadCode = "D1320"
    Tobacco_WL.WorkLoadDef = "Tobacco cessation counseling"
    
    OHDemo_WL.WorkloadId = 283
    OHDemo_WL.WorkloadCode = "D1330"
    OHDemo_WL.WorkLoadDef = "Oral hygiene instruct"
   
    Impress_WL.WorkloadId = 117
    Impress_WL.WorkloadCode = "D0470"
    Impress_WL.WorkLoadDef = "Diagnostics casts"
      
      
      
    ProphCompleteIncompleteAdult_WL.WorkloadId = 340
    ProphCompleteIncompleteAdult_WL.WorkloadCode = "D1110"
    ProphCompleteIncompleteAdult_WL.WorkLoadDef = "Adult prophy"
    
    FluorideUsedAdult_WL.WorkloadId = 473
    FluorideUsedAdult_WL.WorkloadCode = "D1204"
    FluorideUsedAdult_WL.WorkLoadDef = "Adult topical fluoride w/o prophy"
    
    ProphFluorideBothAdult_WL.WorkloadId = 471
    ProphFluorideBothAdult_WL.WorkloadCode = "D1205"
    ProphFluorideBothAdult_WL.WorkLoadDef = "Adult prophy and fluoride"
       
    ProphCompleteIncompleteChild_WL.WorkloadId = 341
    ProphCompleteIncompleteChild_WL.WorkloadCode = "D1120"
    ProphCompleteIncompleteChild_WL.WorkLoadDef = "Child prophy"
    
    FluorideUsedChild_WL.WorkloadId = 474
    FluorideUsedChild_WL.WorkloadCode = "D1203"
    FluorideUsedChild_WL.WorkLoadDef = "Child topical fluoride w/o prophy"
    
    ProphFluorideBothChild_WL.WorkloadId = 472
    ProphFluorideBothChild_WL.WorkloadCode = "D1201"
    ProphFluorideBothChild_WL.WorkLoadDef = "Child prophy and fluoride"
   
   
   
   
    Bleaching_WL.WorkloadId = 505
    Bleaching_WL.WorkloadCode = "D9972"
    Bleaching_WL.WorkLoadDef = "External bleacing - per arch"
    
    Athletic_WL.WorkloadId = 136
    Athletic_WL.WorkloadCode = "D9941"
    Athletic_WL.WorkLoadDef = "Athletic mouthguard"
    
    Fluoride_WL.WorkloadId = 150
    Fluoride_WL.WorkloadCode = "D5986"
    Fluoride_WL.WorkLoadDef = "Fluoride tray"
    
    Scaling_WL_1To3.WorkloadId = 719
    Scaling_WL_1To3.WorkloadCode = "D4342"
    Scaling_WL_1To3.WorkLoadDef = "Scaling and root planing, one to three teeth, per quad"
    
    Scaling_WL_4Plus.WorkloadId = 316
    Scaling_WL_4Plus.WorkloadCode = "D4341"
    Scaling_WL_4Plus.WorkLoadDef = "Scaling and root planing"
    
    FullMouth_WL.WorkloadId = 155
    FullMouth_WL.WorkloadCode = "D4355"
    FullMouth_WL.WorkLoadDef = "Full mouth debridement"

    Perio_WL.WorkloadId = 315
    Perio_WL.WorkloadCode = "D4910"
    Perio_WL.WorkLoadDef = "Perio maintenance"
    
    Implant_WL.WorkloadId = 175
    Implant_WL.WorkloadCode = "D6080"
    Implant_WL.WorkLoadDef = "Implant maintenance"
   
    SiteSpecific.WorkloadId = 224
    SiteSpecific.WorkloadCode = "D4381"
    SiteSpecific.WorkLoadDef = "Controlled release chemotherapeutic agent"
    
    DAgent.WorkloadId = 29
    DAgent.WorkloadCode = "D9910"
    DAgent.WorkLoadDef = "Application of desensitizing medicament"
        
End Sub

Public Sub DoGUIBizRules()

Const sPROC_NAME = "DoGUIBizRules"
    
On Error GoTo ErrHandler:
        
    If mvarbPRH_SUP_FullMouth Then
        mvarbPRH_SUP_MaxRight = False
        mvarbPRH_SUP_ManRight = False
        mvarbPRH_SUP_MaxAnt = False
        mvarbPRH_Sup_ManAnt = False
        mvarbPRH_SUP_MaxLeft = False
        mvarbPRH_SUP_ManLeft = False
    End If
    
    If mvarbPRH_POL_FullMouth Then
        mvarbPRH_POL_MaxRight = False
        mvarbPRH_POL_ManRight = False
        mvarbPRH_POL_MaxAnt = False
        mvarbPRH_POL_ManAnt = False
        mvarbPRH_POL_MaxLeft = False
        mvarbPRH_POL_ManLeft = False
    End If

Exit Sub

ErrHandler:
    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & "." & sPROC_NAME, App.Title, vbCritical
    
End Sub

Public Sub DoWorkLoadBizRules(ByVal vlWorkLoadType As WorkLoadTypeEnum)

Dim lTempWLID As Long
Dim dfFlag As DirtyFlagEnum
Dim iMultiplier As Integer
Dim sWorkLoadCode As String
Dim blnCariesWL As Boolean
Dim blnDelete As Boolean

blnDelete = True
Const sPROC_NAME = "DoWorkLoadBizRules"
    
blnDelete = True

On Error GoTo ErrHandler
    sWorkLoadCode = GetWorkloadCode(vlWorkLoadType, iMultiplier)
    
    If vlWorkLoadType = enumPRV_WC_Nutritional Then
    
        'KDunne 7/18/2002
        'SCR 25851 Dental Nutritonal workload credit may also be received in
        '           caries risk protocol. Make sure that we haven't already
        '           received it there.
        If lTempWLID = 0 And sWorkLoadCode <> Empty Then
            lTempWLID = FindWorkloadId(Val(sWorkLoadCode), mvarlDentalEncounterID)

        End If
        
        If CheckCariesNutrionalWL(mvarlDentalEncounterID) And sWorkLoadCode = Empty Then
            blnDelete = False
       
        End If
        
        lTempWLID = mvarlWLID_Nutritional
'        If mvarlWLID_Nutritional > 0 And lTempWLID = 0 And sWorkLoadCode = Empty Then
'            lTempWLID = mvarlWLID_Nutritional
'        End If
    End If
    If vlWorkLoadType = enumPRV_WC_Tobacco Then lTempWLID = mvarlWLID_Tobacco
    If vlWorkLoadType = enumPRV_OH_Demo Then
        lTempWLID = mvarlWLID_OHDemo
        
        'KDunne 7/18/2002
        'SCR 25853 Oral Hygiene demo workload credit may also be received in
        '           caries risk protocol. Make sure that we haven't already
        '           received it there.
        If lTempWLID = 0 Then
            lTempWLID = FindWorkloadId(Val(sWorkLoadCode), mvarlDentalEncounterID)
        End If
    End If
    If vlWorkLoadType = enumPRV_Impressions Then lTempWLID = mvarlWLID_Impress
    If vlWorkLoadType = enumPRV_DL_Bleaching Then lTempWLID = mvarlWLID_Bleaching
    If vlWorkLoadType = enumPRV_DL_Athletic Then lTempWLID = mvarlWLID_Athletic
    If vlWorkLoadType = enumPRV_DL_Fluoride Then lTempWLID = mvarlWLID_Fluoride
    If vlWorkLoadType = enumPROPH_Fluoride Then lTempWLID = mvarlWLID_Proph
    If vlWorkLoadType = enumHG_Scaling_1to3 Then lTempWLID = mvarlWLID_Scaling_1To3
    If vlWorkLoadType = enumHG_Scaling_4Plus Then lTempWLID = mvarlWLID_Scaling_4Plus
    If vlWorkLoadType = enumHG_DM_FullMouth Then lTempWLID = mvarlWLID_FullMouth
    If vlWorkLoadType = enumHG_DM_Perio Then lTempWLID = mvarlWLID_Perio
'SCR 11118
'    If vlWorkLoadType = enumHG_DM_Implant Then lTempWLID = mvarlWLID_Implant
    If vlWorkLoadType = enumHG_MED_SiteSpecific Then lTempWLID = mvarlWLID_SiteSpecific
    If vlWorkLoadType = enumHG_MED_Agent Then lTempWLID = mvarlWLID_DAgent
       
    If sWorkLoadCode <> "" Then
        If lTempWLID = 0 Then
            dfFlag = dfAdd
        Else
            dfFlag = dfUpdate
        End If
        If Not UpdateWorkload(lTempWLID, dfFlag, mvarlDentalEncounterID, sWorkLoadCode, iMultiplier) Then
            Exit Sub
        End If
    Else
        If lTempWLID <> 0 And blnDelete Then
            'Delete the workload credit
            UpdateWorkload lTempWLID, dfdelete, mvarlDentalEncounterID, "", iMultiplier
        End If
    End If
    
    If vlWorkLoadType = enumPRV_WC_Nutritional Then mvarlWLID_Nutritional = lTempWLID
    If vlWorkLoadType = enumPRV_WC_Tobacco Then mvarlWLID_Tobacco = lTempWLID
    If vlWorkLoadType = enumPRV_OH_Demo Then mvarlWLID_OHDemo = lTempWLID
    If vlWorkLoadType = enumPRV_Impressions Then mvarlWLID_Impress = lTempWLID
    If vlWorkLoadType = enumPRV_DL_Bleaching Then mvarlWLID_Bleaching = lTempWLID
    If vlWorkLoadType = enumPRV_DL_Athletic Then mvarlWLID_Athletic = lTempWLID
    If vlWorkLoadType = enumPRV_DL_Fluoride Then mvarlWLID_Fluoride = lTempWLID
    If vlWorkLoadType = enumPROPH_Fluoride Then mvarlWLID_Proph = lTempWLID
    If vlWorkLoadType = enumHG_Scaling_1to3 Then mvarlWLID_Scaling_1To3 = lTempWLID
    If vlWorkLoadType = enumHG_Scaling_4Plus Then mvarlWLID_Scaling_4Plus = lTempWLID
    If vlWorkLoadType = enumHG_DM_FullMouth Then mvarlWLID_FullMouth = lTempWLID
    If vlWorkLoadType = enumHG_DM_Perio Then mvarlWLID_Perio = lTempWLID
'SCR 11118
'    If vlWorkLoadType = enumHG_DM_Implant Then mvarlWLID_Implant = lTempWLID
    If vlWorkLoadType = enumHG_MED_SiteSpecific Then mvarlWLID_SiteSpecific = lTempWLID
    If vlWorkLoadType = enumHG_MED_Agent Then mvarlWLID_DAgent = lTempWLID

Exit Sub

ErrHandler:
    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & "." & sPROC_NAME, App.Title, vbCritical
    'Delete the workload credit if it didn't update correctly
    UpdateWorkload lTempWLID, dfdelete, mvarlDentalEncounterID, "", 1
End Sub

Private Function CheckCariesNutrionalWL(ByVal lngDentEncId As Long) As Boolean
    Dim objRS As ADODB.Recordset
    Dim strSql As String
    Dim blnReturn As Boolean
    
    blnReturn = False
    
    strSql = "Select DENTAL_NUTRIT_COUNSEL from TEMPL_CARIES_RISK_PROTOCOL WHERE DENTAL_ENCOUNTER_ID = " & lngDentEncId

    Set objRS = goconn.OpenRecordset(strSql)
    
    If Not objRS Is Nothing Then
        
        
        If objRS.BOF = False And objRS.EOF = False Then
            objRS.MoveFirst
            blnReturn = CBool(objRS("DENTAL_NUTRIT_COUNSEL"))
        End If
    End If
    
    CheckCariesNutrionalWL = blnReturn
End Function
