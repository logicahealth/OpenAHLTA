VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExProsDAFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Implements clsISection

Private mlngDentalEncounterId As Long
Private mlngPatUnitId As Long
Private mlngDirtyFlag As DirtyFlagEnum
Private mlngId As Long

Private mlngProfile As DentAssess_ProfileFindEnum 'local copy
Private mlngMuscle_Tone As DentAssess_MuscleToneEnum 'local copy
Private mlngMuscleDev As DentAssess_MuscleDevelopmentEnum 'local copy
Private mlngLip_Contour As DentAssess_LipContourEnum 'local copy
Private mlngLip_Mobility As DentAssess_LipMobilityEnum 'local copy
Private mlngLip_Length As DentAssess_LipLengthEnum 'local copy
Private mlngSpeech As DentAssess_SpeechEnum 'local copy
Private mlngCoordination As DentAssess_CoordinationEnum 'local copy
Private mlngArch_Size As DentAssess_ArchSizeEnum 'local copy
Private mlngArch_Form As DentAssess_ArchFormEnum 'local copy
Private mlngRidge_Form_Max As DentAssess_RidgeFormMax 'local copy
Private mlngRidge_Form_Man As DentAssess_RidgeFormMan 'local copy
Private mlngTori_Max As DentAssess_ToriEnum 'local copy
Private mlngTori_Man As DentAssess_ToriEnum 'local copy
Private mlngInterarch_Space As DentAssess_InterarchSpaceEnum 'local copy
Private mlngRidge_Parallelism As DentAssess_RidgeParrellEnum 'local copy
Private mlngRidge_Relationship As DentAssess_RidgeRelationsEnum 'local copy
Private mlngBone_Quality As DentAssess_BoneQualityEnum 'local copy
Private mlngBone_Quantity As DentAssess_BoneQuantityEnum 'local copy
Private mlngThroat_Form As DentAssess_ThroatFormEnum 'local copy
Private mlngPalatal_Throat_Form As DentAssess_PalatalThroatFormEnum 'local copy
Private mlngPalatal_Sensitivity As DentAssess_PalatalSensitiveEnum 'local copy
Private mlngMucosa_Thickness As DentAssess_MucosaThickEnum 'local copy
Private mlngMucosa_Conditioning As DentAssess_MucosaCondEnum 'local copy
Private mlngBorder_Attachment_Max As DentAssess_BorderAttachEnum 'local copy
Private mlngBorder_Attachment_Man As DentAssess_BorderAttachEnum 'local copy
Private mlngFrenum_Attachment_Max As DentAssess_FrenumAttachEnum 'local copy
Private mlngFrenum_Attachment_Man As DentAssess_FrenumAttachEnum 'local copy
Private mlngSaliva As DentAssess_SalivaEnum 'local copy
Private mlngTongue_Size As DentAssess_TongueSizeEnum 'local copy
Private mlngTongue_Position As DentAssess_TonguePosEnum 'local copy
Private mlngExisting_Denture_Esthetics As DentAssess_GoodFairPoorEnum 'local copy
'local variable(s) to hold property value(s)
Private mlngExisting_Denture_Phonetics As DentAssess_GoodFairPoorEnum 'local copy
Private mlngExisting_Denture_Retention As DentAssess_GoodFairPoorEnum 'local copy
Private mlngExisting_Denture_Stability As DentAssess_GoodFairPoorEnum 'local copy
Private mlngExisting_Denture_Extension As DentAssess_GoodFairPoorEnum 'local copy
Private mlngExisting_Denture_Contours As DentAssess_GoodFairPoorEnum 'local copy
Private mlngCentric_Relation As DentAssess_AcceptablityEnum 'local copy
Private mlngVDO As DentAssess_VDOEnum 'local copy
Private mlngCrossbite As DentAssess_CrossbiteEnum 'local copy
Private mlngCharacterization As DentAssess_CharacterizationEnum 'local copy
Private mlngHygiene As DentAssess_GoodFairPoorEnum 'local copy
Private mlngWear As DentAssess_DentureWearEnum 'local copy
Private mblnPostDamMaxUnaccpt As Boolean 'local copy
Private mblnPostDamManUnaccpt As Boolean 'local copy
Private mblnBaseAdaptMaxUnaccpt As Boolean 'local copy
Private mblnBaseAdaptManUnaccpt As Boolean 'local copy
Private mblnMidLineMaxUnaccpt As Boolean 'local copy
Private mblnMidLineManUnaccpt As Boolean 'local copy
Private mblnComfortManUnaccept As Boolean 'local copy
Private mblnComfortMaxUnaccept As Boolean 'local copy
Private mlngDentOcclPlane As Long

Private Const const_ModuleName = "clsExProsDAFind"

Public Enum DentAssess_CentricRelationEnum
    daCR_Not_Addressed
    daCR_CR1
    daCR_CR2
End Enum

Public Enum DentAssess_DentureOcclPlaneEnum
    daDOP_Not_Addressed
    daDOP_Adequate
    daDOP_High
    daDOP_Low
End Enum

Public Enum DentAssess_ProfileFindEnum
    daPF_Not_Addressed
    daPF_Normal
    daPF_Rethrognathic
    daPF_Prognathic
End Enum

Public Enum DentAssess_MuscleToneEnum
    daMT_Not_Addressed
    daMT_Normal
    daMT_SlightlyImpaired
    daMT_Impaired
End Enum

Public Enum DentAssess_MuscleDevelopmentEnum
    daMD_Not_Addressed
    daMD_HeavyFacialMusclalture
    daMD_Medium
    daMD_Light
End Enum

Public Enum DentAssess_LipContourEnum
    daLC_Not_Addressed
    daLC_Adequate
    daLC_Inadequate
    daLC_Not_Applicable
    daLC_Excessive
End Enum

Public Enum DentAssess_LipMobilityEnum
    daLM_Not_Addressed
    daLM_Normal
    daLM_ReducedMobility
    daLM_Paralysis
End Enum
Public Enum DentAssess_LipLengthEnum
    daLL_Not_Addressed
    daLL_LongLength
    daLL_NormalLength
    daLL_ShortLength
    
End Enum

Public Enum DentAssess_SpeechEnum
    daS_Not_Addressed
    daS_Normal
    daS_Affected
End Enum

Public Enum DentAssess_CoordinationEnum
    daCoord_Not_Addressed
    daCoord_Excellent
    daCoord_Fair
    daCoord_Poor
End Enum

Public Enum DentAssess_ArchSizeEnum
    daAS_Not_Addressed
    daAS_Large
    daAS_Medium
    daAS_Small
End Enum

Public Enum DentAssess_ArchFormEnum
    daAF_Not_Addressed
    daAF_Square
    daAF_Tapering
    daAF_Ovoid
End Enum

Public Enum DentAssess_RidgeFormMax
    daRFM_Not_Addressed
    daRFM_Square
    daRFM_Tapering
    daRFM_Flat
End Enum

Public Enum DentAssess_RidgeFormMan
    daRFM_Not_Addressed
    daRFM_favorable
    daRFM_LessFavorable
    daRFM_Unfavorable
End Enum

Public Enum DentAssess_ToriEnum
    daT_Not_Addressed
    daT_Absent
    daT_Moderate
    daT_Large
End Enum

Public Enum DentAssess_InterarchSpaceEnum
    daIS_Not_Addressed
    daIS_Ideal
    daIS_Excessive
    daIS_Insufficient
End Enum

Public Enum DentAssess_RidgeParrellEnum
    daRP_Not_Addressed
    daRP_Parallel
    daRP_ManDiverges
    daRP_BothDiverge
End Enum

Public Enum DentAssess_RidgeRelationsEnum
    daRR_Not_Addressed
    daRR_Normal
    daRR_Retrognathic
    daRR_Prognathic
End Enum

Public Enum DentAssess_BoneQuantityEnum
    daBQuan_Not_Addressed
    daBQuan_A
    daBQuan_B
    daBQuan_C
    daBQuan_D
    daBQuan_E
End Enum

Public Enum DentAssess_BoneQualityEnum
    daBQual_Not_Addressed
    daBQual_1
    daBQual_2
    daBQual_3
    daBQual_4

End Enum

Public Enum DentAssess_ThroatFormEnum
    daTF_Not_Addressed
    daTF_RightAngle
    daTF_SlightAngle
    daTF_ModerateAngle
End Enum

Public Enum DentAssess_PalatalThroatFormEnum
    daPTF_Not_Addressed
    daPTF_Large
    daPTF_Medium
    daPTF_Small
End Enum

Public Enum DentAssess_PalatalSensitiveEnum
    daPS_Not_Addressed
    daPS_Normal
    daPS_Sensitive
    daPS_HyperSensitive
End Enum

Public Enum DentAssess_MucosaThickEnum
    daMT_Not_Addressed
    daMT_Normal
    daMT_Thick
    daMT_ExcessiveThick
End Enum

Public Enum DentAssess_MucosaCondEnum
    daMC_Not_Addressed
    daMC_Healthy
    daMC_Irritated
    daMC_Pathologic
End Enum

Public Enum DentAssess_BorderAttachEnum
    daBA_Not_Addressed
    daBA_Attach1
    daBA_Attach2
    daBA_Attach3

End Enum

Public Enum DentAssess_FrenumAttachEnum
    daFA_Not_Addressed
    daFA_High
    daFA_Medium
    daFA_FreniEncroach

End Enum

Public Enum DentAssess_SalivaEnum
    daS_Not_Addressed
    daS_Normal
    daS_Excessive
    daS_Xerostomia
End Enum

Public Enum DentAssess_TongueSizeEnum
    daTS_Not_Addressed
    daTS_Normal
    daTS_TeethAbsent
    daTS_Large
End Enum

Public Enum DentAssess_TonguePosEnum
    daTP_Not_Addressed
    daTP_Normal
    daTP_Class1Retracted
    daTP_Class2Retracted
End Enum

Public Enum DentAssess_GoodFairPoorEnum
    daGFP_Not_Addressed
    daGFP_Good
    daGFP_Fair
    daGFP_Poor
End Enum

Public Enum DentAssess_AcceptablityEnum
    daAcc_Not_Addressed
    daAcc_Acceptable
    daAcc_Unacceptable
End Enum

Public Enum DentAssess_VDOEnum
    daVDO_Not_Addressed
    daVDO_Adequate
    daVDO_Inadequate
    daVDO_Excessive
End Enum

Public Enum DentAssess_CrossbiteEnum
    daCB_Not_Addressed
    daCB_None
    daCB_Unilateral
    daCB_Bilateral
    daCB_Interior
End Enum

Public Enum DentAssess_CharacterizationEnum
    daChar_Not_Addressed
    daChar_Characterized
    daChar_Uncharacterized
End Enum

Public Enum DentAssess_DentureWearEnum
    daDW_Not_Addressed
    daDW_Minimal
    daDW_Moderate
    daDW_Severe
End Enum

Public Property Let Muscle_Development(ByVal vData As DentAssess_MuscleDevelopmentEnum)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Muscle_Development = 5
    mlngMuscleDev = vData
    UpdateDirtyFlag
End Property


Public Property Get Muscle_Development() As DentAssess_MuscleDevelopmentEnum
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Muscle_Development
    Muscle_Development = mlngMuscleDev
End Property




Public Property Let Comfort_Max_Unacceptable(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Comfort_Man_Unacceptable = 5
    mblnComfortMaxUnaccept = vData
    UpdateDirtyFlag
End Property




Public Property Let Comfort_Man_Unacceptable(ByVal blnData As Boolean)
      mblnComfortManUnaccept = blnData
      UpdateDirtyFlag
End Property
Public Property Get Comfort_Man_Unacceptable() As Boolean
    Comfort_Man_Unacceptable = mblnComfortManUnaccept
End Property

Public Property Get Comfort_Max_Unacceptable() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Comfort_Max_Unacceptable
    Comfort_Max_Unacceptable = mblnComfortMaxUnaccept
End Property


Public Property Get Midline_Deviation_Man_Unacceptable() As Boolean
    Midline_Deviation_Man_Unacceptable = mblnMidLineManUnaccpt
End Property
Public Property Let Midline_Deviation_Man_Unacceptable(ByVal blnData As Boolean)
    mblnMidLineManUnaccpt = blnData
End Property

Public Property Let Midline_Deviation_Max_Unacceptable(ByVal vData As Boolean)

    mblnMidLineMaxUnaccpt = vData
    UpdateDirtyFlag
End Property


Public Property Get Midline_Deviation_Max_Unacceptable() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Midline_Deviation_Max_Unacceptable
    Midline_Deviation_Max_Unacceptable = mblnMidLineMaxUnaccpt
End Property



Public Property Let Base_Adapt_Man_Unacceptable(ByVal vData As Boolean)
    mblnBaseAdaptManUnaccpt = vData
    UpdateDirtyFlag
End Property


Public Property Get Base_Adapt_Man_Unacceptable() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Base_Adapt_Man_Unacceptable
    Base_Adapt_Man_Unacceptable = mblnBaseAdaptManUnaccpt
End Property



Public Property Let Base_Adapt_Max_Unacceptable(ByVal vData As Boolean)
    UpdateDirtyFlag
    mblnBaseAdaptMaxUnaccpt = vData
End Property


Public Property Get Base_Adapt_Max_Unacceptable() As Boolean
    Base_Adapt_Max_Unacceptable = mblnBaseAdaptMaxUnaccpt
End Property



Public Property Let Post_Dam_Man_Unacceptable(ByVal vData As Boolean)
    mblnPostDamManUnaccpt = vData
    UpdateDirtyFlag
End Property


Public Property Get Post_Dam_Man_Unacceptable() As Boolean
    Post_Dam_Man_Unacceptable = mblnPostDamManUnaccpt
End Property



Public Property Let Post_Dam_Max_Unacceptable(ByVal vData As Boolean)
    mblnPostDamMaxUnaccpt = vData
    UpdateDirtyFlag
End Property


Public Property Get Post_Dam_Max_Unacceptable() As Boolean
    Post_Dam_Max_Unacceptable = mblnPostDamMaxUnaccpt
End Property

Public Property Let Wear(ByVal vData As DentAssess_DentureWearEnum)
    mlngWear = vData
    UpdateDirtyFlag
End Property


Public Property Get Wear() As DentAssess_DentureWearEnum
    Wear = mlngWear
End Property

Public Property Let Hygiene(ByVal vData As DentAssess_GoodFairPoorEnum)
    mlngHygiene = vData
    UpdateDirtyFlag
End Property


Public Property Get Hygiene() As DentAssess_GoodFairPoorEnum
    Hygiene = mlngHygiene
End Property



Public Property Let Characterization(ByVal vData As DentAssess_CharacterizationEnum)
    mlngCharacterization = vData
    UpdateDirtyFlag
End Property


Public Property Get Characterization() As DentAssess_CharacterizationEnum
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Characterization
    Characterization = mlngCharacterization
End Property



Public Property Let Crossbite(ByVal vData As DentAssess_CrossbiteEnum)

    mlngCrossbite = vData
    UpdateDirtyFlag
End Property


Public Property Get Crossbite() As DentAssess_CrossbiteEnum
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Crossbite
    Crossbite = mlngCrossbite
End Property



Public Property Let VDO(ByVal vData As DentAssess_VDOEnum)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.VDO = 5
    mlngVDO = vData
    UpdateDirtyFlag
End Property


Public Property Get VDO() As DentAssess_VDOEnum
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.VDO
     VDO = mlngVDO
End Property



Public Property Let Centric_Relation(ByVal vData As DentAssess_AcceptablityEnum)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Centric_Relation = 5
    mlngCentric_Relation = vData
    UpdateDirtyFlag
End Property


Public Property Get Centric_Relation() As DentAssess_AcceptablityEnum
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Centric_Relation
     Centric_Relation = mlngCentric_Relation
End Property



Public Property Let Existing_Denture_Contours(ByVal vData As DentAssess_GoodFairPoorEnum)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Existing_Denture_Contours = 5
    mlngExisting_Denture_Contours = vData
    UpdateDirtyFlag
End Property


Public Property Get Existing_Denture_Contours() As DentAssess_GoodFairPoorEnum
    Existing_Denture_Contours = mlngExisting_Denture_Contours
End Property

Public Property Let Existing_Denture_Extension(ByVal vData As DentAssess_GoodFairPoorEnum)
    mlngExisting_Denture_Extension = vData
    UpdateDirtyFlag
End Property


Public Property Get Existing_Denture_Extension() As DentAssess_GoodFairPoorEnum
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Existing_Denture_Extension
    Existing_Denture_Extension = mlngExisting_Denture_Extension
End Property



Public Property Let Existing_Denture_Stability(ByVal vData As DentAssess_GoodFairPoorEnum)

    mlngExisting_Denture_Stability = vData
    UpdateDirtyFlag
End Property


Public Property Get Existing_Denture_Stability() As DentAssess_GoodFairPoorEnum
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Existing_Denture_Stability
    Existing_Denture_Stability = mlngExisting_Denture_Stability
End Property



Public Property Let Existing_Denture_Retention(ByVal vData As DentAssess_GoodFairPoorEnum)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Existing_Denture_Retention = 5
    mlngExisting_Denture_Retention = vData
    UpdateDirtyFlag
End Property


Public Property Get Existing_Denture_Retention() As DentAssess_GoodFairPoorEnum
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Existing_Denture_Retention
    Existing_Denture_Retention = mlngExisting_Denture_Retention
End Property



Public Property Let Existing_Denture_Phonetics(ByVal vData As DentAssess_GoodFairPoorEnum)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Existing_Denture_Phonetics = 5
    mlngExisting_Denture_Phonetics = vData
    UpdateDirtyFlag
End Property


Public Property Get Existing_Denture_Phonetics() As DentAssess_GoodFairPoorEnum
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Existing_Denture_Phonetics
    Existing_Denture_Phonetics = mlngExisting_Denture_Phonetics
End Property



Public Property Let Existing_Denture_Esthetics(ByVal vData As DentAssess_GoodFairPoorEnum)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Existing_Denture_Esthetics = 5
    mlngExisting_Denture_Esthetics = vData
    UpdateDirtyFlag
End Property


Public Property Get Existing_Denture_Esthetics() As DentAssess_GoodFairPoorEnum
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Existing_Denture_Esthetics
    Existing_Denture_Esthetics = mlngExisting_Denture_Esthetics
End Property



Public Property Let Tongue_Position(ByVal vData As DentAssess_TonguePosEnum)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Tongue_Position = 5
    mlngTongue_Position = vData
    UpdateDirtyFlag
End Property


Public Property Get Tongue_Position() As DentAssess_TonguePosEnum
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Tongue_Position
    Tongue_Position = mlngTongue_Position
End Property



Public Property Let Tongue_Size(ByVal vData As DentAssess_TongueSizeEnum)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Tongue_Size = 5
    mlngTongue_Size = vData
    UpdateDirtyFlag
End Property


Public Property Get Tongue_Size() As DentAssess_TongueSizeEnum
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Tongue_Size
    Tongue_Size = mlngTongue_Size
End Property



Public Property Let Saliva(ByVal vData As DentAssess_SalivaEnum)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Saliva = 5
    mlngSaliva = vData
    UpdateDirtyFlag
End Property


Public Property Get Saliva() As DentAssess_SalivaEnum
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Saliva
    Saliva = mlngSaliva
End Property



Public Property Let Frenum_Attachment_Man(ByVal vData As DentAssess_FrenumAttachEnum)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Frenum_Attachment_Man = 5
    mlngFrenum_Attachment_Man = vData
    UpdateDirtyFlag
End Property


Public Property Get Frenum_Attachment_Man() As DentAssess_FrenumAttachEnum
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Frenum_Attachment_Man
    Frenum_Attachment_Man = mlngFrenum_Attachment_Man
End Property



Public Property Let Frenum_Attachment_Max(ByVal vData As DentAssess_FrenumAttachEnum)
     UpdateDirtyFlag
'Syntax: X.Frenum_Attachment_Max = 5
    mlngFrenum_Attachment_Max = vData
End Property


Public Property Get Frenum_Attachment_Max() As DentAssess_FrenumAttachEnum
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Frenum_Attachment_Max
    Frenum_Attachment_Max = mlngFrenum_Attachment_Max
End Property



Public Property Let Border_Attachment_Man(ByVal vData As DentAssess_BorderAttachEnum)
    UpdateDirtyFlag
'Syntax: X.Border_Attachment_Man = 5
    mlngBorder_Attachment_Man = vData
End Property


Public Property Get Border_Attachment_Man() As DentAssess_BorderAttachEnum
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Border_Attachment_Man
    Border_Attachment_Man = mlngBorder_Attachment_Man
End Property



Public Property Let Border_Attachment_Max(ByVal vData As DentAssess_BorderAttachEnum)
    UpdateDirtyFlag
'Syntax: X.Border_Attachment_Max = 5
    mlngBorder_Attachment_Max = vData
End Property


Public Property Get Border_Attachment_Max() As DentAssess_BorderAttachEnum
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Border_Attachment_Max
    Border_Attachment_Max = mlngBorder_Attachment_Max
End Property



Public Property Let Mucosa_Conditioning(ByVal vData As DentAssess_MucosaCondEnum)
    UpdateDirtyFlag
'Syntax: X.Mucosa_Conditioning = 5
    mlngMucosa_Conditioning = vData
End Property


Public Property Get Mucosa_Conditioning() As DentAssess_MucosaCondEnum
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Mucosa_Conditioning
    Mucosa_Conditioning = mlngMucosa_Conditioning
End Property



Public Property Let Mucosa_Thickness(ByVal vData As DentAssess_MucosaThickEnum)
    UpdateDirtyFlag
'Syntax: X.Mucosa_Thickness = 5
    mlngMucosa_Thickness = vData
End Property


Public Property Get Mucosa_Thickness() As DentAssess_MucosaThickEnum
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Mucosa_Thickness
    Mucosa_Thickness = mlngMucosa_Thickness
End Property



Public Property Let Palatal_Sensitivity(ByVal vData As DentAssess_PalatalSensitiveEnum)
    UpdateDirtyFlag
'Syntax: X.Palatal_Sensitivity = 5
    mlngPalatal_Sensitivity = vData
End Property


Public Property Get Palatal_Sensitivity() As DentAssess_PalatalSensitiveEnum
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Palatal_Sensitivity
    Palatal_Sensitivity = mlngPalatal_Sensitivity
End Property

Public Property Let Denture_Occl_Plane(ByVal lngValue As DentAssess_DentureOcclPlaneEnum)
    mlngDentOcclPlane = lngValue
    UpdateDirtyFlag
End Property
Public Property Get Denture_Occl_Plane() As DentAssess_DentureOcclPlaneEnum
    Denture_Occl_Plane = mlngDentOcclPlane
End Property

Public Property Let Palatal_Throat_Form(ByVal vData As DentAssess_PalatalThroatFormEnum)
    UpdateDirtyFlag
'Syntax: X.Palatal_Throat_Form = 5
    mlngPalatal_Throat_Form = vData
End Property


Public Property Get Palatal_Throat_Form() As DentAssess_PalatalThroatFormEnum
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Palatal_Throat_Form
    Palatal_Throat_Form = mlngPalatal_Throat_Form
End Property



Public Property Let Throat_Form(ByVal vData As DentAssess_ThroatFormEnum)
    UpdateDirtyFlag
'Syntax: X.Throat_Form = 5
    mlngThroat_Form = vData
End Property


Public Property Get Throat_Form() As DentAssess_ThroatFormEnum
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Throat_Form
    Throat_Form = mlngThroat_Form
End Property



Public Property Let Bone_Quantity(ByVal vData As DentAssess_BoneQuantityEnum)
    UpdateDirtyFlag
'Syntax: X.Bone_Quantity = 5
    mlngBone_Quantity = vData
End Property


Public Property Get Bone_Quantity() As DentAssess_BoneQuantityEnum
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Bone_Quantity
    Bone_Quantity = mlngBone_Quantity
End Property



Public Property Let Bone_Quality(ByVal vData As DentAssess_BoneQualityEnum)
    UpdateDirtyFlag
'Syntax: X.Bone_Quality = 5
    mlngBone_Quality = vData
End Property


Public Property Get Bone_Quality() As DentAssess_BoneQualityEnum
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Bone_Quality
    Bone_Quality = mlngBone_Quality
End Property



Public Property Let Ridge_Relationship(ByVal vData As DentAssess_RidgeRelationsEnum)
    UpdateDirtyFlag
'Syntax: X.Ridge_Relationship = 5
    mlngRidge_Relationship = vData
End Property


Public Property Get Ridge_Relationship() As DentAssess_RidgeRelationsEnum
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Ridge_Relationship
    Ridge_Relationship = mlngRidge_Relationship
End Property



Public Property Let Ridge_Parallelism(ByVal vData As DentAssess_RidgeParrellEnum)
    UpdateDirtyFlag
'Syntax: X.Ridge_Parallelism = 5
    mlngRidge_Parallelism = vData
End Property


Public Property Get Ridge_Parallelism() As DentAssess_RidgeParrellEnum
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Ridge_Parallelism
    Ridge_Parallelism = mlngRidge_Parallelism
End Property



Public Property Let Interarch_Space(ByVal vData As DentAssess_InterarchSpaceEnum)
    UpdateDirtyFlag
'Syntax: X.Interarch_Space = 5
    mlngInterarch_Space = vData
End Property


Public Property Get Interarch_Space() As DentAssess_InterarchSpaceEnum
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Interarch_Space
    Interarch_Space = mlngInterarch_Space
End Property



Public Property Let Tori_Man(ByVal vData As DentAssess_ToriEnum)
    UpdateDirtyFlag
'Syntax: X.Tori_Man = 5
    mlngTori_Man = vData
End Property


Public Property Get Tori_Man() As DentAssess_ToriEnum
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Tori_Man
    Tori_Man = mlngTori_Man
End Property



Public Property Let Tori_Max(ByVal vData As DentAssess_ToriEnum)
    UpdateDirtyFlag
'Syntax: X.Tori_Max = 5
    mlngTori_Max = vData
End Property


Public Property Get Tori_Max() As DentAssess_ToriEnum
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Tori_Max
    Tori_Max = mlngTori_Max
End Property



Public Property Let Ridge_Form_Man(ByVal vData As DentAssess_RidgeFormMan)
    UpdateDirtyFlag
'Syntax: X.Ridge_Form_Man = 5
    mlngRidge_Form_Man = vData
End Property


Public Property Get Ridge_Form_Man() As DentAssess_RidgeFormMan
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Ridge_Form_Man
    Ridge_Form_Man = mlngRidge_Form_Man
End Property



Public Property Let Ridge_Form_Max(ByVal vData As DentAssess_RidgeFormMax)
    UpdateDirtyFlag
'Syntax: X.Ridge_Form_Max = 5
    mlngRidge_Form_Max = vData
End Property


Public Property Get Ridge_Form_Max() As DentAssess_RidgeFormMax
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Ridge_Form_Max
    Ridge_Form_Max = mlngRidge_Form_Max
End Property



Public Property Let Arch_Form(ByVal vData As DentAssess_ArchFormEnum)
    UpdateDirtyFlag
'Syntax: X.Arch_Form = 5
    mlngArch_Form = vData
End Property


Public Property Get Arch_Form() As DentAssess_ArchFormEnum
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Arch_Form
    Arch_Form = mlngArch_Form
End Property



Public Property Let Arch_Size(ByVal vData As DentAssess_ArchSizeEnum)
    UpdateDirtyFlag
'Syntax: X.Arch_Size = 5
    mlngArch_Size = vData
End Property


Public Property Get Arch_Size() As DentAssess_ArchSizeEnum
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Arch_Size
    Arch_Size = mlngArch_Size
End Property



Public Property Let Coordination(ByVal vData As DentAssess_CoordinationEnum)
    UpdateDirtyFlag
'Syntax: X.Coordination = 5
    mlngCoordination = vData
End Property


Public Property Get Coordination() As DentAssess_CoordinationEnum
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Coordination
    Coordination = mlngCoordination
End Property



Public Property Let Speech(ByVal vData As DentAssess_SpeechEnum)
    UpdateDirtyFlag
'Syntax: X.Speech = 5
    mlngSpeech = vData
End Property


Public Property Get Speech() As DentAssess_SpeechEnum
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Speech
    Speech = mlngSpeech
End Property



Public Property Let Lip_Length(ByVal vData As DentAssess_LipLengthEnum)
    UpdateDirtyFlag
'Syntax: X.Lip_Length = 5
    mlngLip_Length = vData
End Property


Public Property Get Lip_Length() As DentAssess_LipLengthEnum
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Lip_Length
    Lip_Length = mlngLip_Length
End Property



Public Property Let Lip_Mobility(ByVal vData As DentAssess_LipMobilityEnum)
    UpdateDirtyFlag
'Syntax: X.Lip_Mobility = 5
    mlngLip_Mobility = vData
End Property


Public Property Get Lip_Mobility() As DentAssess_LipMobilityEnum
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Lip_Mobility
    Lip_Mobility = mlngLip_Mobility
End Property



Public Property Let Lip_Contour(ByVal vData As DentAssess_LipContourEnum)
    UpdateDirtyFlag
'Syntax: X.Lip_Contour = 5
    mlngLip_Contour = vData
End Property


Public Property Get Lip_Contour() As DentAssess_LipContourEnum
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Lip_Contour
    Lip_Contour = mlngLip_Contour
End Property



Public Property Let Muscle_Tone(ByVal vData As DentAssess_MuscleToneEnum)
    UpdateDirtyFlag
'Syntax: X.Muscle_Tone = 5
    mlngMuscle_Tone = vData
End Property


Public Property Get Muscle_Tone() As DentAssess_MuscleToneEnum
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Muscle_Tone
    Muscle_Tone = mlngMuscle_Tone
End Property



Public Property Let Profile(ByVal vData As DentAssess_ProfileFindEnum)
    UpdateDirtyFlag
'Syntax: X.Profile = 5
    mlngProfile = vData
End Property


Public Property Get Profile() As DentAssess_ProfileFindEnum
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Profile
    Profile = mlngProfile
End Property




Private Property Get clsISection_DentalEncounterID() As Long
    clsISection_DentalEncounterID = mlngDentalEncounterId
End Property

Private Function clsISection_GetNote() As String
    Dim strNote As String
    
    strNote = Empty
    
        If mlngProfile <> DentAssess_ProfileFindEnum.daPF_Not_Addressed Then
            strNote = strNote & "Profile: " & _
                GetProfileEnumDesc(mlngProfile) & "; "
        End If
        
        If mlngMuscle_Tone <> DentAssess_MuscleToneEnum.daMT_Not_Addressed Then
            strNote = strNote & "Muscle Tone: " & _
                        GetMuscleToneEnumDesc(mlngMuscle_Tone) & "; "
        End If
        
        If mlngMuscleDev <> DentAssess_MuscleDevelopmentEnum.daMD_Not_Addressed Then
            strNote = strNote & "Muscle Development: " & _
                GetMuscleDevEnumDesc(mlngMuscleDev) & "; "
        End If
        
        If mlngLip_Contour <> DentAssess_LipContourEnum.daLC_Not_Addressed Then
            strNote = strNote & "Lip Contour: " & _
                GetLipContourEnumDesc(mlngLip_Contour) & "; "
        End If
        
        If mlngLip_Mobility <> DentAssess_LipMobilityEnum.daLM_Not_Addressed Then
            strNote = strNote & "Lip Mobility: " & _
                    GetLipMobilityEnumDesc(mlngLip_Mobility) & "; "
        End If
        
        If mlngLip_Length <> DentAssess_LipLengthEnum.daLL_Not_Addressed Then
            strNote = strNote & "Lip Length: " & _
                GetLipLengthEnumDesc(mlngLip_Length) & "; "
        End If
        
        If mlngSpeech <> DentAssess_SpeechEnum.daS_Not_Addressed Then
            strNote = strNote & "Speech: " & _
                GetSpeechEnumDesc(mlngSpeech) & "; "
        End If
        If mlngCoordination <> DentAssess_CoordinationEnum.daCoord_Not_Addressed Then
            strNote = strNote & "Coordination: " & _
                GetCoordinationEnumDesc(mlngCoordination) & "; "
        End If

        If mlngArch_Size <> DentAssess_ArchSizeEnum.daAS_Not_Addressed Then
              strNote = strNote & "Arch Size: " & _
                GetArchSizeEnumDesc(mlngArch_Size) & "; "
        End If
        If mlngArch_Form <> DentAssess_ArchFormEnum.daAF_Not_Addressed Then
              strNote = strNote & "Arch Form: " & _
                GetArchFormEnumDesc(mlngArch_Form) & "; "
        End If
        If mlngRidge_Form_Max <> DentAssess_RidgeFormMax.daRFM_Not_Addressed Then
              strNote = strNote & "Ridge Form (Max): " & _
                GetRidgeFormMaxEnumDesc(mlngRidge_Form_Max) & "; "
        End If
        If mlngRidge_Form_Man <> DentAssess_RidgeFormMan.daRFM_Not_Addressed Then
              strNote = strNote & "Ridge Form (Man): " & _
                GetRidgeFormManEnumDesc(mlngRidge_Form_Man) & "; "
        End If
        If mlngTori_Max <> DentAssess_ToriEnum.daT_Not_Addressed Then
              strNote = strNote & "Tori (Max): " & _
                GetToriEnumDesc(mlngTori_Max) & "; "
        End If
        If mlngTori_Man <> DentAssess_ToriEnum.daT_Not_Addressed Then
              strNote = strNote & "Tori (Man): " & _
                GetToriEnumDesc(mlngTori_Man) & "; "
        End If
        If mlngInterarch_Space <> DentAssess_InterarchSpaceEnum.daIS_Not_Addressed Then
              strNote = strNote & "Interarch Space: " & _
                GetInterarchSpaceEnumDesc(mlngInterarch_Space) & "; "
        End If
        If mlngRidge_Parallelism <> DentAssess_RidgeParrellEnum.daRP_Not_Addressed Then
              strNote = strNote & "Ridge Parallelism: " & _
                GetRidgeParallelismEnumDesc(mlngRidge_Parallelism) & "; "
        End If
        If mlngRidge_Relationship <> DentAssess_RidgeRelationsEnum.daRR_Not_Addressed Then
              strNote = strNote & "Ridge Relationship (Angle): " & _
                GetRidgeRelationsEnumDesc(mlngRidge_Relationship) & "; "
        End If
        If mlngBone_Quality <> DentAssess_BoneQualityEnum.daBQual_Not_Addressed Then
              strNote = strNote & "Bone Quality (Branemark): " & _
                GetBoneQualityEnumDesc(mlngBone_Quality) & "; "
        End If
        If mlngBone_Quantity <> DentAssess_BoneQuantityEnum.daBQuan_Not_Addressed Then
              strNote = strNote & "Bone Quantity (Branemark): " & _
                GetBoneQuantityEnumDesc(mlngBone_Quantity) & "; "
        End If
        If mlngThroat_Form <> DentAssess_ThroatFormEnum.daTF_Not_Addressed Then
              strNote = strNote & "Throat Form (Neil) : " & _
                GetThroatFormEnumDesc(mlngThroat_Form) & "; "
        End If
        If mlngPalatal_Throat_Form <> DentAssess_PalatalThroatFormEnum.daPTF_Not_Addressed Then
              strNote = strNote & "Palatal Throat Form (House): " & _
                GetPalatalThroatFormEnumDesc(mlngPalatal_Throat_Form) & "; "
        End If
        If mlngPalatal_Sensitivity <> DentAssess_PalatalSensitiveEnum.daPS_Not_Addressed Then
              strNote = strNote & "Palatal Sensitivity (House): " & _
                GetPalatalSensitivityEnumDesc(mlngPalatal_Sensitivity) & "; "
        End If
        If mlngMucosa_Thickness <> DentAssess_MucosaThickEnum.daMT_Not_Addressed Then
              strNote = strNote & "Mucosa Thickness (House): " & _
                GetMucosaThicknessEnumDesc(mlngMucosa_Thickness) & "; "
        End If
        If mlngMucosa_Conditioning <> DentAssess_MucosaCondEnum.daMC_Not_Addressed Then
              strNote = strNote & "Mucosa Condition (House): " & _
                GetMucosaConditionEnumDesc(mlngMucosa_Conditioning) & "; "
        End If
        If mlngBorder_Attachment_Max <> DentAssess_BorderAttachEnum.daBA_Not_Addressed Then
              strNote = strNote & "Border Attachment (Max) - House: " & _
                GetBorderAttachEnumDesc(mlngBorder_Attachment_Max) & "; "
        End If
        If mlngBorder_Attachment_Man <> DentAssess_BorderAttachEnum.daBA_Not_Addressed Then
              strNote = strNote & "Border Attachment (Man) - House: " & _
                GetBorderAttachEnumDesc(mlngBorder_Attachment_Man) & "; "
        End If
        If mlngFrenum_Attachment_Max <> DentAssess_FrenumAttachEnum.daFA_Not_Addressed Then
              strNote = strNote & "Frenum Attachment (Max) - House: " & _
                GetFrenumAttachEnumDesc(mlngFrenum_Attachment_Max) & "; "
        End If
        If mlngFrenum_Attachment_Man <> DentAssess_FrenumAttachEnum.daFA_Not_Addressed Then
              strNote = strNote & "Frenum Attachment (Man) - House: " & _
                GetFrenumAttachEnumDesc(mlngFrenum_Attachment_Man) & "; "
        End If
        If mlngSaliva <> DentAssess_SalivaEnum.daS_Not_Addressed Then
              strNote = strNote & "Saliva: " & _
                GetSalivaEnumDesc(mlngSaliva) & "; "
        End If
        If mlngTongue_Size <> DentAssess_TongueSizeEnum.daTS_Not_Addressed Then
              strNote = strNote & "Tongue Size (House) : " & _
                GetTongueSizeEnumDesc(mlngTongue_Size) & "; "
        End If
        If mlngTongue_Position <> DentAssess_TonguePosEnum.daTP_Not_Addressed Then
              strNote = strNote & "Tongue Position (Wright): " & _
                GetTonguePosEnumDesc(mlngTongue_Position) & "; "
        End If
        If mlngExisting_Denture_Esthetics <> DentAssess_GoodFairPoorEnum.daGFP_Not_Addressed Then
              strNote = strNote & "Existing Denture Esthetics: " & _
                GetGoodFairPoorEnumDesc(mlngExisting_Denture_Esthetics) & "; "
        End If
        If mlngExisting_Denture_Phonetics <> DentAssess_GoodFairPoorEnum.daGFP_Not_Addressed Then
              strNote = strNote & "Existing Denture Phonetics: " & _
                GetGoodFairPoorEnumDesc(mlngExisting_Denture_Phonetics) & "; "
        End If
        If mlngExisting_Denture_Retention <> DentAssess_GoodFairPoorEnum.daGFP_Not_Addressed Then
              strNote = strNote & "Existing Denture Retention: " & _
                GetGoodFairPoorEnumDesc(mlngExisting_Denture_Retention) & "; "
        End If
        If mlngExisting_Denture_Stability <> DentAssess_GoodFairPoorEnum.daGFP_Not_Addressed Then
              strNote = strNote & "Existing Denture Stability: " & _
                GetGoodFairPoorEnumDesc(mlngExisting_Denture_Stability) & "; "
        End If
        If mlngExisting_Denture_Extension <> DentAssess_GoodFairPoorEnum.daGFP_Not_Addressed Then
              strNote = strNote & "Existing Denture Extension: " & _
                GetGoodFairPoorEnumDesc(mlngExisting_Denture_Extension) & "; "
        End If
        If mlngExisting_Denture_Contours <> DentAssess_GoodFairPoorEnum.daGFP_Not_Addressed Then
              strNote = strNote & "Existing Denture Contours: " & _
                GetGoodFairPoorEnumDesc(mlngExisting_Denture_Contours) & "; "
        End If
        If mlngCentric_Relation <> DentAssess_CentricRelationEnum.daCR_Not_Addressed Then
              strNote = strNote & "Centric Relation: " & _
                GetCentricRelationEnumDesc(mlngCentric_Relation) & "; "
        End If
        If mlngDentOcclPlane <> DentAssess_DentureOcclPlaneEnum.daDOP_Not_Addressed Then
            strNote = strNote & "Denture Occlusal Plane: " & _
                GetDentureOcclPlaneEnumDesc(mlngDentOcclPlane) & "; "
        End If
        If mlngVDO <> DentAssess_VDOEnum.daVDO_Not_Addressed Then
              strNote = strNote & "VDO: " & _
              GetVDOEnumDesc(mlngVDO) & "; "
        End If
        If mlngCrossbite <> DentAssess_CrossbiteEnum.daCB_Not_Addressed Then
              strNote = strNote & "Crossbite: " & _
                GetCrossbiteEnumDesc(mlngCrossbite) & "; "
        End If
        If mlngCharacterization <> DentAssess_CharacterizationEnum.daChar_Not_Addressed Then
              strNote = strNote & "Characterization: " & _
                GetCharacterizationEnumDesc(mlngCharacterization) & "; "
        End If
        If mlngHygiene <> DentAssess_GoodFairPoorEnum.daGFP_Not_Addressed Then
              strNote = strNote & "Hygiene: " & _
                GetGoodFairPoorEnumDesc(mlngHygiene) & "; "
        End If
        If mlngWear <> DentAssess_DentureWearEnum.daDW_Not_Addressed Then
              strNote = strNote & "Wear of denture teeth: " & _
                GetDentureWearEnumDesc(mlngWear) & "; "
        End If
        If mblnBaseAdaptManUnaccpt <> False Then
              strNote = strNote & "Base Adapt. (Man) is unacceptable." & "; "
        End If
        If mblnBaseAdaptMaxUnaccpt <> False Then
              strNote = strNote & "Base Adapt. (Max) is unacceptable." & "; "
        End If
        If mblnComfortManUnaccept <> False Then
              strNote = strNote & "Comfort (Man) is unacceptable." & "; "
        End If
        If mblnComfortMaxUnaccept <> False Then
              strNote = strNote & "Comfort (Max) is unacceptable." & "; "
        End If
        If mblnMidLineManUnaccpt <> False Then
              strNote = strNote & "Midline Deviation (Man) is unacceptable." & "; "
        End If
        If mblnMidLineMaxUnaccpt <> False Then
              strNote = strNote & "Midline Deviation (Max) is unacceptable." & "; "
        End If
        If mblnPostDamManUnaccpt <> False Then
              strNote = strNote & "Post Dam (Man) is unacceptable." & "; "
        End If
        If mblnPostDamMaxUnaccpt <> False Then
              strNote = strNote & "Post Dam (Max) is unacceptable." & "; "
        End If
            
        If strNote <> "" Then
            strNote = "Denture Assessement Findings: " & strNote
        End If
        clsISection_GetNote = strNote
End Function
Private Function clsISection_GetRTFNote() As String
    Dim strNote As String
    
    strNote = Empty
    
        If mlngProfile > DentAssess_ProfileFindEnum.daPF_Not_Addressed Then
            strNote = strNote & "Profile: " & _
                GetProfileEnumDesc(mlngProfile) & "; "
        End If
        
        If mlngMuscle_Tone > DentAssess_MuscleToneEnum.daMT_Not_Addressed Then
            strNote = strNote & "Muscle Tone: " & _
                        GetMuscleToneEnumDesc(mlngMuscle_Tone) & "; "
        End If
        
        If mlngMuscleDev > DentAssess_MuscleDevelopmentEnum.daMD_Not_Addressed Then
            strNote = strNote & "Muscle Development: " & _
                GetMuscleDevEnumDesc(mlngMuscleDev) & "; "
        End If
        
        If mlngLip_Contour > DentAssess_LipContourEnum.daLC_Not_Addressed Then
            strNote = strNote & "Lip Contour: " & _
                GetLipContourEnumDesc(mlngLip_Contour) & "; "
        End If
        
        If mlngLip_Mobility > DentAssess_LipMobilityEnum.daLM_Not_Addressed Then
            strNote = strNote & "Lip Mobility: " & _
                    GetLipMobilityEnumDesc(mlngLip_Mobility) & "; "
        End If
        
        If mlngLip_Length > DentAssess_LipLengthEnum.daLL_Not_Addressed Then
            strNote = strNote & "Lip Length: " & _
                GetLipLengthEnumDesc(mlngLip_Length) & "; "
        End If
        
        If mlngSpeech > DentAssess_SpeechEnum.daS_Not_Addressed Then
            strNote = strNote & "Speech: " & _
                GetSpeechEnumDesc(mlngSpeech) & "; "
        End If
        If mlngCoordination > DentAssess_CoordinationEnum.daCoord_Not_Addressed Then
            strNote = strNote & "Coordination: " & _
                GetCoordinationEnumDesc(mlngCoordination) & "; "
        End If

        If mlngArch_Size > DentAssess_ArchSizeEnum.daAS_Not_Addressed Then
              strNote = strNote & "Arch Size: " & _
                GetArchSizeEnumDesc(mlngArch_Size) & "; "
        End If
        If mlngArch_Form > DentAssess_ArchFormEnum.daAF_Not_Addressed Then
              strNote = strNote & "Arch Form: " & _
                GetArchFormEnumDesc(mlngArch_Form) & "; "
        End If
        If mlngRidge_Form_Max > DentAssess_RidgeFormMax.daRFM_Not_Addressed Then
              strNote = strNote & "Ridge Form (Max): " & _
                GetRidgeFormMaxEnumDesc(mlngRidge_Form_Max) & "; "
        End If
        If mlngRidge_Form_Man > DentAssess_RidgeFormMan.daRFM_Not_Addressed Then
              strNote = strNote & "Ridge Form (Man): " & _
                GetRidgeFormManEnumDesc(mlngRidge_Form_Man) & "; "
        End If
        If mlngTori_Max > DentAssess_ToriEnum.daT_Not_Addressed Then
              strNote = strNote & "Tori (Max): " & _
                GetToriEnumDesc(mlngTori_Max) & "; "
        End If
        If mlngTori_Man > DentAssess_ToriEnum.daT_Not_Addressed Then
              strNote = strNote & "Tori (Man): " & _
                GetToriEnumDesc(mlngTori_Man) & "; "
        End If
        If mlngInterarch_Space > DentAssess_InterarchSpaceEnum.daIS_Not_Addressed Then
              strNote = strNote & "Interarch Space: " & _
                GetInterarchSpaceEnumDesc(mlngInterarch_Space) & "; "
        End If
        If mlngRidge_Parallelism > DentAssess_RidgeParrellEnum.daRP_Not_Addressed Then
              strNote = strNote & "Ridge Parallelism: " & _
                GetRidgeParallelismEnumDesc(mlngRidge_Parallelism) & "; "
        End If
        If mlngRidge_Relationship > DentAssess_RidgeRelationsEnum.daRR_Not_Addressed Then
              strNote = strNote & "Ridge Relationship (Angle): " & _
                GetRidgeRelationsEnumDesc(mlngRidge_Relationship) & "; "
        End If
        If mlngBone_Quality > DentAssess_BoneQualityEnum.daBQual_Not_Addressed Then
              strNote = strNote & "Bone Quality (Branemark): " & _
                GetBoneQualityEnumDesc(mlngBone_Quality) & "; "
        End If
        If mlngBone_Quantity > DentAssess_BoneQuantityEnum.daBQuan_Not_Addressed Then
              strNote = strNote & "Bone Quantity (Branemark): " & _
                GetBoneQuantityEnumDesc(mlngBone_Quantity) & "; "
        End If
        If mlngThroat_Form > DentAssess_ThroatFormEnum.daTF_Not_Addressed Then
              strNote = strNote & "Throat Form (Neil) : " & _
                GetThroatFormEnumDesc(mlngThroat_Form) & "; "
        End If
        If mlngPalatal_Throat_Form > DentAssess_PalatalThroatFormEnum.daPTF_Not_Addressed Then
              strNote = strNote & "Palatal Throat Form (House): " & _
                GetPalatalThroatFormEnumDesc(mlngPalatal_Throat_Form) & "; "
        End If
        If mlngPalatal_Sensitivity > DentAssess_PalatalSensitiveEnum.daPS_Not_Addressed Then
              strNote = strNote & "Palatal Sensitivity (House): " & _
                GetPalatalSensitivityEnumDesc(mlngPalatal_Sensitivity) & "; "
        End If
        If mlngMucosa_Thickness > DentAssess_MucosaThickEnum.daMT_Not_Addressed Then
              strNote = strNote & "Mucosa Thickness (House): " & _
                GetMucosaThicknessEnumDesc(mlngMucosa_Thickness) & "; "
        End If
        If mlngMucosa_Conditioning > DentAssess_MucosaCondEnum.daMC_Not_Addressed Then
              strNote = strNote & "Mucosa Condition (House): " & _
                GetMucosaConditionEnumDesc(mlngMucosa_Conditioning) & "; "
        End If
        If mlngBorder_Attachment_Max > DentAssess_BorderAttachEnum.daBA_Not_Addressed Then
              strNote = strNote & "Border Attachment (Max) - House: " & _
                GetBorderAttachEnumDesc(mlngBorder_Attachment_Max) & "; "
        End If
        If mlngBorder_Attachment_Man > DentAssess_BorderAttachEnum.daBA_Not_Addressed Then
              strNote = strNote & "Border Attachment (Man) - House: " & _
                GetBorderAttachEnumDesc(mlngBorder_Attachment_Man) & "; "
        End If
        If mlngFrenum_Attachment_Max > DentAssess_FrenumAttachEnum.daFA_Not_Addressed Then
              strNote = strNote & "Frenum Attachment (Max) - House: " & _
                GetFrenumAttachEnumDesc(mlngFrenum_Attachment_Max) & "; "
        End If
        If mlngFrenum_Attachment_Man > DentAssess_FrenumAttachEnum.daFA_Not_Addressed Then
              strNote = strNote & "Frenum Attachment (Man) - House: " & _
                GetFrenumAttachEnumDesc(mlngFrenum_Attachment_Man) & "; "
        End If
        If mlngSaliva > DentAssess_SalivaEnum.daS_Not_Addressed Then
              strNote = strNote & "Saliva: " & _
                GetSalivaEnumDesc(mlngSaliva) & "; "
        End If
        If mlngTongue_Size > DentAssess_TongueSizeEnum.daTS_Not_Addressed Then
              strNote = strNote & "Tongue Size (House) : " & _
                GetTongueSizeEnumDesc(mlngTongue_Size) & "; "
        End If
        If mlngTongue_Position > DentAssess_TonguePosEnum.daTP_Not_Addressed Then
              strNote = strNote & "Tongue Position (Wright): " & _
                GetTonguePosEnumDesc(mlngTongue_Position) & "; "
        End If
        If mlngExisting_Denture_Esthetics > DentAssess_GoodFairPoorEnum.daGFP_Not_Addressed Then
              strNote = strNote & "Existing Denture Esthetics: " & _
                GetGoodFairPoorEnumDesc(mlngExisting_Denture_Esthetics) & "; "
        End If
        If mlngExisting_Denture_Phonetics > DentAssess_GoodFairPoorEnum.daGFP_Not_Addressed Then
              strNote = strNote & "Existing Denture Phonetics: " & _
                GetGoodFairPoorEnumDesc(mlngExisting_Denture_Phonetics) & "; "
        End If
        If mlngExisting_Denture_Retention > DentAssess_GoodFairPoorEnum.daGFP_Not_Addressed Then
              strNote = strNote & "Existing Denture Retention: " & _
                GetGoodFairPoorEnumDesc(mlngExisting_Denture_Retention) & "; "
        End If
        If mlngExisting_Denture_Stability > DentAssess_GoodFairPoorEnum.daGFP_Not_Addressed Then
              strNote = strNote & "Existing Denture Stability: " & _
                GetGoodFairPoorEnumDesc(mlngExisting_Denture_Stability) & "; "
        End If
        If mlngExisting_Denture_Extension > DentAssess_GoodFairPoorEnum.daGFP_Not_Addressed Then
              strNote = strNote & "Existing Denture Extension: " & _
                GetGoodFairPoorEnumDesc(mlngExisting_Denture_Extension) & "; "
        End If
        If mlngExisting_Denture_Contours > DentAssess_GoodFairPoorEnum.daGFP_Not_Addressed Then
              strNote = strNote & "Existing Denture Contours: " & _
                GetGoodFairPoorEnumDesc(mlngExisting_Denture_Contours) & "; "
        End If
        If mlngCentric_Relation > DentAssess_CentricRelationEnum.daCR_Not_Addressed Then
              strNote = strNote & "Centric Relation: " & _
                GetCentricRelationEnumDesc(mlngCentric_Relation) & "; "
        End If
        If mlngDentOcclPlane > DentAssess_DentureOcclPlaneEnum.daDOP_Not_Addressed Then
            strNote = strNote & "Denture Occlusal Plane: " & _
                GetDentureOcclPlaneEnumDesc(mlngDentOcclPlane) & "; "
        End If
        If mlngVDO > DentAssess_VDOEnum.daVDO_Not_Addressed Then
              strNote = strNote & "VDO: " & _
              GetVDOEnumDesc(mlngVDO) & "; "
        End If
        If mlngCrossbite > DentAssess_CrossbiteEnum.daCB_Not_Addressed Then
              strNote = strNote & "Crossbite: " & _
                GetCrossbiteEnumDesc(mlngCrossbite) & "; "
        End If
        If mlngCharacterization > DentAssess_CharacterizationEnum.daChar_Not_Addressed Then
              strNote = strNote & "Characterization: " & _
                GetCharacterizationEnumDesc(mlngCharacterization) & "; "
        End If
        If mlngHygiene > DentAssess_GoodFairPoorEnum.daGFP_Not_Addressed Then
              strNote = strNote & "Hygiene: " & _
                GetGoodFairPoorEnumDesc(mlngHygiene) & "; "
        End If
        If mlngWear > DentAssess_DentureWearEnum.daDW_Not_Addressed Then
              strNote = strNote & "Wear of denture teeth: " & _
                GetDentureWearEnumDesc(mlngWear) & "; "
        End If
        If mblnBaseAdaptManUnaccpt = True Then
              strNote = strNote & "Base Adapt. (Man) is unacceptable." & "; "
        End If
        If mblnBaseAdaptMaxUnaccpt = True Then
              strNote = strNote & "Base Adapt. (Max) is unacceptable." & "; "
        End If
        If mblnComfortManUnaccept = True Then
              strNote = strNote & "Comfort (Man) is unacceptable." & "; "
        End If
        If mblnComfortMaxUnaccept = True Then
              strNote = strNote & "Comfort (Max) is unacceptable." & "; "
        End If
        If mblnMidLineManUnaccpt = True Then
              strNote = strNote & "Midline Deviation (Man) is unacceptable." & "; "
        End If
        If mblnMidLineMaxUnaccpt = True Then
              strNote = strNote & "Midline Deviation (Max) is unacceptable." & "; "
        End If
        If mblnPostDamManUnaccpt = True Then
              strNote = strNote & "Post Dam (Man) is unacceptable." & "; "
        End If
        If mblnPostDamMaxUnaccpt = True Then
              strNote = strNote & "Post Dam (Max) is unacceptable." & "; "
        End If
            
        If Trim(strNote) <> "" Then
            strNote = crtfStartControlName & "Denture Assessement Findings: " & _
                crtfEndControlName & strNote
        End If
        clsISection_GetRTFNote = strNote
End Function
Private Function clsISection_LoadSection(ByVal PatientID As Long, ByVal DentalEncounterID As Long) As Boolean
    Dim blnReturn As Boolean
    Dim objRS As ADODB.Recordset
    
    On Error GoTo ErrHandler
    
    mlngPatUnitId = PatientID
    mlngDentalEncounterId = DentalEncounterID
    
    Set objRS = DB_LoadDentureAssessmentFindings(mlngDentalEncounterId)
    With objRS
        If .EOF Then
            mlngDirtyFlag = dfAdd
            'reproperty values here
            mlngProfile = DentAssess_ProfileFindEnum.daPF_Not_Addressed
            mlngMuscle_Tone = DentAssess_MuscleToneEnum.daMT_Not_Addressed
            mlngMuscleDev = DentAssess_MuscleToneEnum.daMT_Not_Addressed
            mlngLip_Contour = DentAssess_LipContourEnum.daLC_Not_Addressed
            mlngLip_Mobility = DentAssess_LipMobilityEnum.daLM_Not_Addressed
            mlngLip_Length = DentAssess_LipLengthEnum.daLL_Not_Addressed
            mlngSpeech = DentAssess_SpeechEnum.daS_Not_Addressed
            mlngCoordination = DentAssess_CoordinationEnum.daCoord_Not_Addressed
            mlngArch_Size = DentAssess_ArchSizeEnum.daAS_Not_Addressed
            mlngArch_Form = DentAssess_ArchFormEnum.daAF_Not_Addressed
            mlngRidge_Form_Max = DentAssess_RidgeFormMax.daRFM_Not_Addressed
            mlngRidge_Form_Man = DentAssess_RidgeFormMax.daRFM_Not_Addressed
            mlngTori_Max = DentAssess_ToriEnum.daT_Not_Addressed
            mlngTori_Man = DentAssess_ToriEnum.daT_Not_Addressed
            mlngInterarch_Space = DentAssess_InterarchSpaceEnum.daIS_Not_Addressed
            mlngRidge_Parallelism = DentAssess_RidgeParrellEnum.daRP_Not_Addressed
            mlngRidge_Relationship = DentAssess_RidgeRelationsEnum.daRR_Not_Addressed
            mlngBone_Quality = DentAssess_BoneQualityEnum.daBQual_Not_Addressed
            mlngBone_Quantity = DentAssess_BoneQuantityEnum.daBQuan_Not_Addressed
            mlngThroat_Form = DentAssess_ThroatFormEnum.daTF_Not_Addressed
            mlngPalatal_Throat_Form = DentAssess_PalatalThroatFormEnum.daPTF_Not_Addressed
            mlngPalatal_Sensitivity = DentAssess_PalatalSensitiveEnum.daPS_Not_Addressed
            mlngMucosa_Thickness = DentAssess_MucosaThickEnum.daMT_Not_Addressed
            mlngMucosa_Conditioning = DentAssess_MucosaCondEnum.daMC_Not_Addressed
            mlngBorder_Attachment_Max = DentAssess_BorderAttachEnum.daBA_Not_Addressed
            mlngBorder_Attachment_Man = DentAssess_BorderAttachEnum.daBA_Not_Addressed
            mlngFrenum_Attachment_Max = DentAssess_FrenumAttachEnum.daFA_Not_Addressed
            mlngFrenum_Attachment_Man = DentAssess_FrenumAttachEnum.daFA_Not_Addressed
            mlngSaliva = DentAssess_SalivaEnum.daS_Not_Addressed
            mlngTongue_Size = DentAssess_TongueSizeEnum.daTS_Not_Addressed
            mlngTongue_Position = DentAssess_TonguePosEnum.daTP_Not_Addressed
            mlngExisting_Denture_Esthetics = DentAssess_GoodFairPoorEnum.daGFP_Not_Addressed
            mlngExisting_Denture_Phonetics = DentAssess_GoodFairPoorEnum.daGFP_Not_Addressed
            mlngExisting_Denture_Retention = DentAssess_GoodFairPoorEnum.daGFP_Not_Addressed
            mlngExisting_Denture_Stability = DentAssess_GoodFairPoorEnum.daGFP_Not_Addressed
            mlngExisting_Denture_Extension = DentAssess_GoodFairPoorEnum.daGFP_Not_Addressed
            mlngExisting_Denture_Contours = DentAssess_GoodFairPoorEnum.daGFP_Not_Addressed
            mlngCentric_Relation = DentAssess_AcceptablityEnum.daAcc_Not_Addressed
            mlngVDO = DentAssess_VDOEnum.daVDO_Not_Addressed
            mlngCrossbite = DentAssess_CrossbiteEnum.daCB_Not_Addressed
            mlngCharacterization = DentAssess_CharacterizationEnum.daChar_Not_Addressed
            mlngHygiene = DentAssess_GoodFairPoorEnum.daGFP_Not_Addressed
            mlngWear = DentAssess_DentureWearEnum.daDW_Not_Addressed
            mlngDentOcclPlane = DentAssess_DentureOcclPlaneEnum.daDOP_Not_Addressed
            
            mblnBaseAdaptManUnaccpt = False
            mblnBaseAdaptMaxUnaccpt = False
            mblnComfortMaxUnaccept = False
            mblnComfortManUnaccept = False
            mblnMidLineManUnaccpt = False
            mblnMidLineMaxUnaccpt = False
            mblnPostDamManUnaccpt = False
            mblnPostDamMaxUnaccpt = False
        Else
            mlngDirtyFlag = dfUpdate
            'Set property values here
            
            mlngId = .Fields("EP_PA_FIND_ID")
            mlngProfile = .Fields("PROFILE")
            mlngMuscle_Tone = .Fields("MUSCLE_TONE")
            mlngMuscleDev = .Fields("MUSCLE_DEVELOPMENT")
            mlngLip_Contour = .Fields("LIP_CONTOUR")
            mlngLip_Mobility = .Fields("LIP_MOBILITY")
            
            mlngLip_Length = .Fields("LIP_LENGTH")
            
            mlngSpeech = .Fields("SPEECH")
            mlngCoordination = .Fields("COORDINATION")
            mlngArch_Size = .Fields("ARCH_SIZE")
            mlngArch_Form = .Fields("ARCH_FORM")
            mlngRidge_Form_Max = .Fields("RIDGE_FORM_MAX")
            mlngRidge_Form_Man = .Fields("RIDGE_FORM_MAN")
            mlngTori_Max = .Fields("TORI_MAX")
            mlngTori_Man = .Fields("TORI_MAN")
            mlngInterarch_Space = .Fields("INTERARCH_SPACE")
            mlngRidge_Parallelism = .Fields("RIDGE_PARRALLELISM")
            
            mlngRidge_Relationship = .Fields("Ridge_Relations")
            
            mlngBone_Quality = .Fields("BONE_QUALITY")
            mlngBone_Quantity = .Fields("BONE_QUANTITY")
            mlngThroat_Form = .Fields("THROAT_FORM")
            mlngPalatal_Throat_Form = .Fields("PALATAL_THROAT_FORM")
            mlngPalatal_Sensitivity = .Fields("PALATAL_SENSITIVITY")
            mlngMucosa_Thickness = .Fields("MUCOSA_THICKNESS")
            mlngMucosa_Conditioning = .Fields("MUCOSA_CONDITION")
            mlngBorder_Attachment_Max = .Fields("BORDER_ATTACHMENT_MAX")
            mlngBorder_Attachment_Man = .Fields("BORDER_ATTACHMENT_MAN")
            mlngFrenum_Attachment_Max = .Fields("FRENUM_ATTACH_MAX")
            mlngFrenum_Attachment_Man = .Fields("FRENUM_ATTACH_MAN")
            mlngSaliva = .Fields("SALIVA")
            mlngTongue_Size = .Fields("TONGUE_SIZE")
            mlngTongue_Position = .Fields("TONGUE_POSITION")
            mlngExisting_Denture_Esthetics = .Fields("EXIST_DENTURE_ESTH")
            mlngExisting_Denture_Phonetics = .Fields("EXIST_DENTURE_PHON")
            mlngExisting_Denture_Retention = .Fields("EXIST_DENTURE_RETEN")
            mlngExisting_Denture_Stability = .Fields("EXIST_DENTURE_STAB")
            mlngExisting_Denture_Extension = .Fields("EXIST_DENTURE_EXTEN")
            mlngExisting_Denture_Contours = .Fields("EXIST_DENTURE_CONT")
            mlngCentric_Relation = .Fields("CENTRIC_RELATION")
            mlngVDO = .Fields("VDO")
            mlngCrossbite = .Fields("CROSSBITE")
            mlngHygiene = .Fields("HYGIENE")
            mlngCharacterization = .Fields("CHARACTER")
            mlngWear = .Fields("WEAR")
            mlngDentOcclPlane = .Fields("DENTURE_OCCL_PLANE")
            mblnBaseAdaptManUnaccpt = CBool(.Fields("BASE_ADAPT_MAN"))
            mblnBaseAdaptMaxUnaccpt = CBool(.Fields("BASE_ADAPT_MAX"))
            mblnComfortManUnaccept = CBool(.Fields("COMFORT_MAN"))
            mblnComfortMaxUnaccept = CBool(.Fields("COMFORT_MAX"))
            mblnMidLineManUnaccpt = CBool(.Fields("MID_DEV_MAN"))
            mblnMidLineMaxUnaccpt = CBool(.Fields("MID_DEV_MAX"))
            mblnPostDamManUnaccpt = CBool(.Fields("POST_DAM_MAN_UNACCEPT"))
            mblnPostDamMaxUnaccpt = CBool(.Fields("POST_DAM_MAX_UNACCEPT"))
            
        End If
    End With
    
FuncExit:
    clsISection_LoadSection = blnReturn
    Set objRS = Nothing
    Exit Function
ErrHandler:
   blnReturn = False
   goCWShared.ShowVBError Err.Number, Err.Description, const_ModuleName & ".clsISection_LoadSection()", "", vbCritical
   GoTo FuncExit
End Function

Private Property Get clsISection_PatientID() As Long
    clsISection_PatientID = mlngPatUnitId
End Property

Private Function clsISection_SaveSection() As Boolean
    Dim strSql As String
    Dim blnReturn As Boolean
    
    On Error GoTo ErrHandler
    
    strSql = Empty
    Select Case mlngDirtyFlag
        Case dfNoAction
            blnReturn = True
        Case dfAdd
            
            mlngId = goconn.GetNextID(esTEMPL_EXAM_PROS_DENTURE_F_SEQ)
'            strSql = "SELECT TEMPL_EXAM_PROS_DENTURE_F_Seq.NEXTVAL From Dual"
'            Set objRS = goconn.OpenRecordset(strSql)
'            mlngId = objRS.Fields("NEXTVAL")

            strSql = "INSERT INTO TEMPL_EXAM_PROS_DENTURE_FIND " & _
             "(" & _
             "EP_PA_FIND_ID, DENTAL_ENCOUNTER_ID, PROFILE, " & _
             "MUSCLE_TONE, MUSCLE_DEVELOPMENT, LIP_CONTOUR, LIP_LENGTH, " & _
             "LIP_MOBILITY, SPEECH, COORDINATION, ARCH_SIZE, " & _
             "ARCH_FORM, RIDGE_FORM_MAX, RIDGE_FORM_MAN, " & _
             "TORI_MAX, TORI_MAN, INTERARCH_SPACE, " & _
             "RIDGE_PARRALLELISM, Ridge_Relations, BONE_QUANTITY, BONE_QUALITY, " & _
             "THROAT_FORM, PALATAL_THROAT_FORM, " & _
             "MUCOSA_THICKNESS, PALATAL_SENSITIVITY, " & _
             "MUCOSA_CONDITION, BORDER_ATTACHMENT_MAX, " & _
             "BORDER_ATTACHMENT_MAN, FRENUM_ATTACH_MAX, " & _
             "FRENUM_ATTACH_MAN, SALIVA, TONGUE_SIZE, " & _
             "TONGUE_POSITION, EXIST_DENTURE_ESTH, " & _
             "EXIST_DENTURE_PHON, EXIST_DENTURE_RETEN, " & _
             "EXIST_DENTURE_STAB, EXIST_DENTURE_EXTEN, " & _
             "EXIST_DENTURE_CONT, CENTRIC_RELATION, " & _
             "DENTURE_OCCL_PLANE, VDO, CROSSBITE, CHARACTER, " & _
             "HYGIENE, WEAR, POST_DAM_MAX_UNACCEPT, " & _
             "POST_DAM_MAN_UNACCEPT, BASE_ADAPT_MAX, " & _
             "BASE_ADAPT_MAN, MID_DEV_MAX, MID_DEV_MAN, " & _
             "COMFORT_MAX , COMFORT_MAN )" & _
             "VALUES (" & mlngId & ", " & mlngDentalEncounterId & ", " & mlngProfile & _
             ", " & mlngMuscle_Tone & ", " & mlngMuscleDev & ", " & mlngLip_Contour & _
             ", " & mlngLip_Length & ", " & mlngLip_Mobility & ", " & mlngSpeech & ", "
    strSql = strSql & _
             mlngCoordination & ", " & mlngArch_Size & ", " & mlngArch_Form & ", " & _
             mlngRidge_Form_Max & ", " & mlngRidge_Form_Man & ", " & mlngTori_Max & ", " & _
             mlngTori_Man & ", " & mlngInterarch_Space & ", " & mlngRidge_Parallelism & _
             ", " & mlngRidge_Relationship & ", " & mlngBone_Quantity & ", " & mlngBone_Quality & _
             ", " & mlngThroat_Form & ", " & mlngPalatal_Throat_Form & ", " & mlngMucosa_Thickness & _
             ", " & mlngPalatal_Sensitivity & ", " & mlngMucosa_Conditioning & ", " & _
             mlngBorder_Attachment_Max & ", " & mlngBorder_Attachment_Man & ", " & _
             mlngFrenum_Attachment_Max & ", " & mlngFrenum_Attachment_Man & ", " & _
             mlngSaliva & ", " & mlngTongue_Size & ", " & mlngTongue_Position & ", " & _
             mlngExisting_Denture_Esthetics & ", " & mlngExisting_Denture_Phonetics & ", " & _
             mlngExisting_Denture_Retention & ", " & mlngExisting_Denture_Stability & ", " & _
             mlngExisting_Denture_Extension & ", " & mlngExisting_Denture_Contours & ", " & _
             mlngCentric_Relation & ", " & mlngDentOcclPlane & ", " & mlngVDO & ", " & mlngCrossbite & ", " & _
             mlngCharacterization & ", " & mlngHygiene & ", " & mlngWear & ", " & _
             CInt(mblnPostDamMaxUnaccpt) & ", " & CInt(mblnPostDamManUnaccpt) & ", " & _
             CInt(mblnBaseAdaptMaxUnaccpt) & ", " & CInt(mblnBaseAdaptManUnaccpt) & ", " & _
             CInt(mblnMidLineMaxUnaccpt) & ", " & CInt(mblnMidLineManUnaccpt) & ", " & _
             CInt(mblnComfortMaxUnaccept) & ", " & CInt(mblnComfortManUnaccept) & ")"
        Case dfUpdate
            strSql = "UPDATE TEMPL_EXAM_PROS_DENTURE_FIND " & _
               " SET DENTAL_ENCOUNTER_ID = " & mlngDentalEncounterId & ", " & _
                "PROFILE = " & mlngProfile & ", MUSCLE_TONE = " & mlngMuscle_Tone & _
                ", MUSCLE_DEVELOPMENT = " & mlngMuscleDev & ", LIP_CONTOUR = " & mlngLip_Contour & _
                ", LIP_MOBILITY = " & mlngLip_Mobility & ", SPEECH = " & mlngSpeech & _
                ", COORDINATION = " & mlngCoordination & ", ARCH_SIZE = " & mlngArch_Size & _
                ", ARCH_FORM = " & mlngArch_Form & ", RIDGE_FORM_MAX = " & mlngRidge_Form_Max & _
                ", RIDGE_FORM_MAN = " & mlngRidge_Form_Man & ", TORI_MAX =  " & mlngTori_Max & _
                ", TORI_MAN = " & mlngTori_Man & ", INTERARCH_SPACE = " & mlngInterarch_Space & _
                ",  RIDGE_PARRALLELISM = " & mlngRidge_Parallelism & ", BONE_QUANTITY = " & _
                mlngBone_Quantity & ", BONE_QUALITY = " & mlngBone_Quality & ", THROAT_FORM = " & _
                mlngThroat_Form & ", PALATAL_THROAT_FORM = " & mlngPalatal_Throat_Form & _
                ", MUCOSA_THICKNESS = " & mlngMucosa_Thickness & ", PALATAL_SENSITIVITY = " & _
                mlngPalatal_Sensitivity & ", BORDER_ATTACHMENT_MAX = " & mlngBorder_Attachment_Max & _
                ", MUCOSA_CONDITION = " & mlngMucosa_Conditioning & ", BORDER_ATTACHMENT_MAN = " & mlngBorder_Attachment_Man & _
                ",FRENUM_ATTACH_MAX = " & mlngFrenum_Attachment_Max & ", FRENUM_ATTACH_MAN = " & _
                mlngFrenum_Attachment_Man & ", SALIVA = " & mlngSaliva & ", TONGUE_SIZE = " & _
                mlngTongue_Size & ", TONGUE_POSITION = " & mlngTongue_Position & _
                ", EXIST_DENTURE_ESTH = " & mlngExisting_Denture_Esthetics & ", EXIST_DENTURE_PHON = " & _
                mlngExisting_Denture_Phonetics & ", EXIST_DENTURE_RETEN = " & mlngExisting_Denture_Retention & _
                ", EXIST_DENTURE_STAB = " & mlngExisting_Denture_Stability & ", EXIST_DENTURE_EXTEN = " & _
                mlngExisting_Denture_Extension & ", EXIST_DENTURE_CONT = " & mlngExisting_Denture_Contours & _
                ", CENTRIC_RELATION = " & mlngCentric_Relation & _
                ", VDO = " & mlngVDO & ", CROSSBITE = " & mlngCrossbite & ", CHARACTER = " & _
                mlngCharacterization & ", HYGIENE = " & mlngHygiene & ", WEAR = " & mlngWear
            strSql = strSql & ", POST_DAM_MAX_UNACCEPT = " & CInt(mblnPostDamMaxUnaccpt) & _
                ", POST_DAM_MAN_UNACCEPT = " & CInt(mblnPostDamManUnaccpt) & _
                ", BASE_ADAPT_MAX = " & CInt(mblnBaseAdaptMaxUnaccpt) & _
                ", BASE_ADAPT_MAN = " & CInt(mblnBaseAdaptManUnaccpt) & _
                ", MID_DEV_MAX = " & CInt(mblnMidLineMaxUnaccpt) & _
                ", MID_DEV_MAN = " & CInt(mblnMidLineManUnaccpt) & _
                ", COMFORT_MAX = " & CInt(mblnComfortMaxUnaccept) & _
                ", COMFORT_MAN = " & CInt(mblnComfortManUnaccept) & _
                ", LIP_LENGTH = " & mlngLip_Length & _
                ", RIDGE_RELATIONS = " & mlngRidge_Relationship & _
                ", DENTURE_OCCL_PLANE = " & mlngDentOcclPlane & _
                " Where EP_PA_FIND_ID = " & mlngId
        Case dfdelete
            strSql = "DELETE FROM TEMPL_EXAM_PROS_DENTURE_FIND WHERE " & _
                     " EP_PA_FIND_ID = " & mlngId
    End Select
    
    If strSql <> "" Then
        goconn.ExecuteSQL (strSql)
    End If
    
    mlngDirtyFlag = dfNoAction
    
    blnReturn = True
FuncExit:
    clsISection_SaveSection = blnReturn
    Exit Function
ErrHandler:
    blnReturn = False
    goCWShared.ShowVBError Err.Number, Err.Description, const_ModuleName & ".clsISection_SaveSection()", "", vbCritical
    GoTo FuncExit
End Function

'****************************************************
'Sub: UpdateDirtyFlag
'Author: KDunne
'Date: 11/5/2001
'Purpose: Updates module level dirty flag field
'****************************************************
Private Sub UpdateDirtyFlag()

    If mlngDirtyFlag = dfNoAction Then
        mlngDirtyFlag = dfUpdate
    End If
    
End Sub

'*******************************************************************************
'Function: DB_LoadDentureAssessmentFindings
'Author:KDunne
'Date:12/12/2001
'Purpose: Loads Info from the TEMPL_EXAM_PROS_DENTURE_FIND table
'*******************************************************************************
Private Function DB_LoadDentureAssessmentFindings(ByVal lngDentalEncounterId As Long) As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo ErrHandler:
    
    strSql = "Select EP_PA_FIND_ID, DENTAL_ENCOUNTER_ID, PROFILE, " & _
             "MUSCLE_TONE, MUSCLE_DEVELOPMENT, LIP_CONTOUR, " & _
             "LIP_MOBILITY, SPEECH, COORDINATION, ARCH_SIZE, " & _
             "ARCH_FORM, RIDGE_FORM_MAX, RIDGE_FORM_MAN, " & _
             "TORI_MAX, TORI_MAN, INTERARCH_SPACE, " & _
             "RIDGE_PARRALLELISM, BONE_QUANTITY, BONE_QUALITY, " & _
             "THROAT_FORM, PALATAL_THROAT_FORM, " & _
             "MUCOSA_THICKNESS, PALATAL_SENSITIVITY, " & _
             "MUCOSA_CONDITION, BORDER_ATTACHMENT_MAX, " & _
             "BORDER_ATTACHMENT_MAN, FRENUM_ATTACH_MAX, " & _
             "FRENUM_ATTACH_MAN, SALIVA, TONGUE_SIZE, " & _
             "TONGUE_POSITION, EXIST_DENTURE_ESTH, " & _
             "EXIST_DENTURE_PHON, EXIST_DENTURE_RETEN, " & _
             "EXIST_DENTURE_STAB, EXIST_DENTURE_EXTEN, " & _
             "EXIST_DENTURE_CONT, CENTRIC_RELATION, " & _
             "DENTURE_OCCL_PLANE, VDO, CROSSBITE, CHARACTER, " & _
             "HYGIENE, WEAR, POST_DAM_MAX_UNACCEPT, " & _
             "POST_DAM_MAN_UNACCEPT, BASE_ADAPT_MAX, " & _
             "BASE_ADAPT_MAN, MID_DEV_MAX, MID_DEV_MAN, " & _
             "COMFORT_MAX, COMFORT_MAN, LIP_LENGTH, Ridge_Relations " & _
             "FROM TEMPL_EXAM_PROS_DENTURE_FIND " & _
             "WHERE DENTAL_ENCOUNTER_ID = " & lngDentalEncounterId
             
    Set DB_LoadDentureAssessmentFindings = goconn.OpenRecordset(strSql)
    Exit Function
    
ErrHandler:
    goCWShared.ShowVBError Err.Number, Err.Description, const_ModuleName & ".DB_LoadDentureAssessmentFindings()", "", vbCritical
    Set DB_LoadDentureAssessmentFindings = Nothing

End Function

'*******************************************************************
'Function: GetProfileEnumDesc
'Author: KDunne
'Date: 12/11/2001
'Purpose. Returns the description for a given DentAssess_ProfileFindEnum value.
'*******************************************************************
Public Function GetProfileEnumDesc(ByVal lngValue As _
        DentAssess_ProfileFindEnum) As String
    Dim strReturn As String
    
    Select Case lngValue
        Case DentAssess_ProfileFindEnum.daPF_Normal
            strReturn = "Normal -1"
        Case DentAssess_ProfileFindEnum.daPF_Prognathic
            strReturn = "Prognathic - 3"
        Case DentAssess_ProfileFindEnum.daPF_Rethrognathic
            strReturn = "Retrognathic - 2"
        Case DentAssess_ProfileFindEnum.daPF_Not_Addressed
            strReturn = Empty
        Case Else
            strReturn = Empty
    End Select
    
    GetProfileEnumDesc = strReturn
End Function
'******************************************************************
'Function:
'Author: KDunne
'Date: 12/14/2001
'Purpose: Builds a list of DentAssess_ProfileFindEnum 's
'******************************************************************
Public Function BuildProfileEnumList() As String
    Dim X As Integer
    Dim strReturn As String
    
    For X = 0 To 3
        strReturn = strReturn & GetProfileEnumDesc(X) & "|"
    Next
    
    BuildProfileEnumList = " |" & strReturn
End Function

'*******************************************************************
'Function: GetMuscleToneEnumDesc
'Author: KDunne
'Date: 12/11/2001
'Purpose. Returns the description for a given DentAssess_ProfileFindEnum value.
'*******************************************************************
Public Function GetMuscleToneEnumDesc(ByVal lngValue As _
        DentAssess_MuscleToneEnum) As String
    Dim strReturn As String
    
    Select Case lngValue
        Case DentAssess_MuscleToneEnum.daMT_Impaired
            strReturn = "Impaired - 3"
        Case DentAssess_MuscleToneEnum.daMT_Normal
            strReturn = "Normal - 1"
        Case DentAssess_MuscleToneEnum.daMT_SlightlyImpaired
            strReturn = "Slightly impaired - 2"
        Case DentAssess_MuscleToneEnum.daMT_Not_Addressed
            strReturn = Empty
        Case Else
            strReturn = Empty
    End Select
    
    GetMuscleToneEnumDesc = strReturn
End Function
'******************************************************************
'Function: BuildMuscleToneEnumList
'Author: KDunne
'Date: 12/14/2001
'Purpose: Builds a list of DentAssess_ProfileFindEnum 's
'******************************************************************
Public Function BuildMuscleToneEnumList() As String
    Dim X As Integer
    Dim strReturn As String
    
    For X = 0 To 3
        strReturn = strReturn & GetMuscleToneEnumDesc(X) & "|"
    Next
    
    BuildMuscleToneEnumList = " |" & strReturn
End Function
'*******************************************************************
'Function: GetMuscleDevEnumDesc
'Author: KDunne
'Date: 12/11/2001
'Purpose. Returns the description for a given DentAssess_MuscleDevelopmentEnum value.
'*******************************************************************
Public Function GetMuscleDevEnumDesc(ByVal lngValue As _
        DentAssess_MuscleDevelopmentEnum) As String
    Dim strReturn As String
    
    Select Case lngValue
        Case DentAssess_MuscleDevelopmentEnum.daMD_HeavyFacialMusclalture
            strReturn = "Heavy facial musculature - 1"
        Case DentAssess_MuscleDevelopmentEnum.daMD_Light
            strReturn = "Light - 3"
        Case DentAssess_MuscleDevelopmentEnum.daMD_Medium
            strReturn = "Medium - 2"
        Case DentAssess_MuscleDevelopmentEnum.daMD_Not_Addressed
            strReturn = Empty
        Case Else
            strReturn = Empty
    End Select
    
    GetMuscleDevEnumDesc = strReturn
End Function
'******************************************************************
'Function: BuildMuscleDevEnumList
'Author: KDunne
'Date: 12/14/2001
'Purpose: Builds a list of DentAssess_MuscleDevelopmentEnum 's
'******************************************************************
Public Function BuildMuscleDevEnumList() As String
    Dim X As Integer
    Dim strReturn As String
    
    For X = 0 To 3
        strReturn = strReturn & GetMuscleDevEnumDesc(X) & "|"
    Next
    
    BuildMuscleDevEnumList = " |" & strReturn
End Function
'*******************************************************************
'Function: GetLipContourEnumDesc
'Author: KDunne
'Date: 12/11/2001
'Purpose. Returns the description for a given DentAssess_LipContourEnum value.
'*******************************************************************
Public Function GetLipContourEnumDesc(ByVal lngValue As _
        DentAssess_LipContourEnum) As String
    Dim strReturn As String
    
    Select Case lngValue
        Case DentAssess_LipContourEnum.daLC_Adequate
            strReturn = "Adequately supported"
        Case DentAssess_LipContourEnum.daLC_Excessive
            strReturn = "Excessive"
        Case DentAssess_LipContourEnum.daLC_Inadequate
            strReturn = "Inadequately supported"
        Case DentAssess_LipContourEnum.daLC_Not_Applicable
            strReturn = "Not applicable"
        Case DentAssess_LipContourEnum.daLC_Not_Addressed
            strReturn = Empty
        Case Else
            strReturn = Empty
    End Select
    
    GetLipContourEnumDesc = strReturn
End Function
'******************************************************************
'Function: BuildLipContourEnumList
'Author: KDunne
'Date: 12/14/2001
'Purpose: Builds a list of DentAssess_LipContourEnum 's
'******************************************************************
Public Function BuildLipContourEnumList() As String
    Dim X As Integer
    Dim strReturn As String
    
    For X = 0 To 4
        strReturn = strReturn & GetLipContourEnumDesc(X) & "|"
    Next
    
    BuildLipContourEnumList = " |" & strReturn
End Function
'*******************************************************************
'Function: GetLipMobilityEnumDesc
'Author: KDunne
'Date: 12/11/2001
'Purpose. Returns the description for a given DentAssess_LipContourEnum value.
'*******************************************************************
Public Function GetLipMobilityEnumDesc(ByVal lngValue As _
        DentAssess_LipMobilityEnum) As String
    Dim strReturn As String
    
    Select Case lngValue
        Case DentAssess_LipMobilityEnum.daLM_Normal
            strReturn = "Normal - 1"
        Case DentAssess_LipMobilityEnum.daLM_Paralysis
            strReturn = "Paralysis - 3"
        Case DentAssess_LipMobilityEnum.daLM_ReducedMobility
            strReturn = "Reduced mobility - 2"
        Case DentAssess_LipMobilityEnum.daLM_Not_Addressed
            strReturn = Empty
        Case Else
            strReturn = Empty
    End Select
    
    GetLipMobilityEnumDesc = strReturn
End Function
'******************************************************************
'Function: BuildLipMobilityEnumList
'Author: KDunne
'Date: 12/14/2001
'Purpose: Builds a list of DentAssess_LipMobilityEnum 's
'******************************************************************
Public Function BuildLipMobilityEnumList() As String
    Dim X As Integer
    Dim strReturn As String
    
    For X = 0 To 3
        strReturn = strReturn & GetLipMobilityEnumDesc(X) & "|"
    Next
    
    BuildLipMobilityEnumList = " |" & strReturn
End Function
'*******************************************************************
'Function: GetLipLengthEnumDesc
'Author: KDunne
'Date: 12/11/2001
'Purpose. Returns the description for a given DentAssess_LipLengthEnum value.
'*******************************************************************
Public Function GetLipLengthEnumDesc(ByVal lngValue As _
        DentAssess_LipLengthEnum) As String
    Dim strReturn As String
    
    Select Case lngValue
        Case DentAssess_LipLengthEnum.daLL_NormalLength
            strReturn = "Normal"
        Case DentAssess_LipLengthEnum.daLL_LongLength
            strReturn = "Long"
        Case DentAssess_LipLengthEnum.daLL_ShortLength
            strReturn = "Short"
        Case DentAssess_LipLengthEnum.daLL_Not_Addressed
            strReturn = Empty
        Case Else
            strReturn = Empty
    End Select
    
    GetLipLengthEnumDesc = strReturn
End Function
'******************************************************************
'Function: BuildLipLengthEnumList
'Author: KDunne
'Date: 12/14/2001
'Purpose: Builds a list of DentAssess_LipLengthEnum 's
'******************************************************************
Public Function BuildLipLengthEnumList() As String
    Dim X As Integer
    Dim strReturn As String
    
    For X = 0 To 3
        strReturn = strReturn & GetLipLengthEnumDesc(X) & "|"
    Next
    
    BuildLipLengthEnumList = " |" & strReturn
End Function
'*******************************************************************
'Function: GetSpeechEnumDesc
'Author: KDunne
'Date: 12/11/2001
'Purpose. Returns the description for a given DentAssess_SpeechEnum value.
'*******************************************************************
Public Function GetSpeechEnumDesc(ByVal lngValue As _
        DentAssess_SpeechEnum) As String
    Dim strReturn As String
    
    Select Case lngValue
        Case DentAssess_SpeechEnum.daS_Normal
            strReturn = "Normal"
        Case DentAssess_SpeechEnum.daS_Affected
            strReturn = "Affected"
        Case DentAssess_SpeechEnum.daS_Not_Addressed
            strReturn = Empty
        Case Else
            strReturn = Empty
    End Select
    
    GetSpeechEnumDesc = strReturn
End Function
'******************************************************************
'Function: BuildSpeechEnumList
'Author: KDunne
'Date: 12/14/2001
'Purpose: Builds a list of DentAssess_SpeechEnum 's
'******************************************************************
Public Function BuildSpeechEnumList() As String
    Dim X As Integer
    Dim strReturn As String
    
    For X = 0 To 3
        strReturn = strReturn & GetSpeechEnumDesc(X) & "|"
    Next
    
    BuildSpeechEnumList = " |" & strReturn
End Function
'*******************************************************************
'Function: GetCoordinationEnumDesc
'Author: KDunne
'Date: 12/11/2001
'Purpose. Returns the description for a given DentAssess_CoordinationEnum value.
'*******************************************************************
Public Function GetCoordinationEnumDesc(ByVal lngValue As _
        DentAssess_CoordinationEnum) As String
    Dim strReturn As String
    
    Select Case lngValue
        Case DentAssess_CoordinationEnum.daCoord_Excellent
            strReturn = "Excellent - 1"
        Case DentAssess_CoordinationEnum.daCoord_Fair
            strReturn = "Fair - 2"
        Case DentAssess_CoordinationEnum.daCoord_Poor
            strReturn = "Poor - 3"
        Case DentAssess_CoordinationEnum.daCoord_Not_Addressed
            strReturn = Empty
        Case Else
            strReturn = Empty
    End Select
    
    GetCoordinationEnumDesc = strReturn
End Function
'******************************************************************
'Function: BuildCoordinationEnumList
'Author: KDunne
'Date: 12/14/2001
'Purpose: Builds a list of DentAssess_CoordinationEnum 's
'******************************************************************
Public Function BuildCoordinationEnumList() As String
    Dim X As Integer
    Dim strReturn As String
    
    For X = 0 To 3
        strReturn = strReturn & GetCoordinationEnumDesc(X) & "|"
    Next
    
    BuildCoordinationEnumList = " |" & strReturn
End Function
'*******************************************************************
'Function: GetArchSizeEnumDesc
'Author: KDunne
'Date: 12/11/2001
'Purpose. Returns the description for a given DentAssess_ArchSizeEnum value.
'*******************************************************************
Public Function GetArchSizeEnumDesc(ByVal lngValue As _
        DentAssess_ArchSizeEnum) As String
    Dim strReturn As String
    
    Select Case lngValue
        Case DentAssess_ArchSizeEnum.daAS_Large
            strReturn = "Large - 1"
        Case DentAssess_ArchSizeEnum.daAS_Medium
            strReturn = "Medium - 2"
        Case DentAssess_ArchSizeEnum.daAS_Small
            strReturn = "Small - 3"
        Case DentAssess_ArchSizeEnum.daAS_Not_Addressed
            strReturn = Empty
        Case Else
            strReturn = Empty
    End Select
    
    GetArchSizeEnumDesc = strReturn
End Function
'******************************************************************
'Function: BuildArchSizeEnumList
'Author: KDunne
'Date: 12/14/2001
'Purpose: Builds a list of DentAssess_ArchSizeEnum 's
'******************************************************************
Public Function BuildArchSizeEnumList() As String
    Dim X As Integer
    Dim strReturn As String
    
    For X = 0 To 3
        strReturn = strReturn & GetArchSizeEnumDesc(X) & "|"
    Next
    
    BuildArchSizeEnumList = " |" & strReturn
End Function

'*******************************************************************
'Function: GetArchFormEnumDesc
'Author: KDunne
'Date: 12/11/2001
'Purpose. Returns the description for a given DentAssess_ArchFormEnum value.
'*******************************************************************
Public Function GetArchFormEnumDesc(ByVal lngValue As _
        DentAssess_ArchFormEnum) As String
    Dim strReturn As String
    
    Select Case lngValue
        Case DentAssess_ArchFormEnum.daAF_Square
            strReturn = "Square - 1"
        Case DentAssess_ArchFormEnum.daAF_Ovoid
            strReturn = "Ovoid - 3"
        Case DentAssess_ArchFormEnum.daAF_Tapering
            strReturn = "Tapering - 2"
        Case DentAssess_ArchFormEnum.daAF_Not_Addressed
            strReturn = Empty
        Case Else
            strReturn = Empty
    End Select
    
    GetArchFormEnumDesc = strReturn
End Function
'******************************************************************
'Function: BuildArchFormEnumList
'Author: KDunne
'Date: 12/14/2001
'Purpose: Builds a list of DentAssess_ArchFormEnum 's
'******************************************************************
Public Function BuildArchFormEnumList() As String
    Dim X As Integer
    Dim strReturn As String
    
    For X = 0 To 3
        strReturn = strReturn & GetArchFormEnumDesc(X) & "|"
    Next
    
    BuildArchFormEnumList = " |" & strReturn
End Function
'*******************************************************************
'Function: GetRidgeFormMaxEnumDesc
'Author: KDunne
'Date: 12/11/2001
'Purpose. Returns the description for a given DentAssess_RidgeFormMax value.
'*******************************************************************
Public Function GetRidgeFormMaxEnumDesc(ByVal lngValue As _
        DentAssess_RidgeFormMax) As String
    Dim strReturn As String
    
    Select Case lngValue
        Case DentAssess_RidgeFormMax.daRFM_Square
            strReturn = "Square - 1"
        Case DentAssess_RidgeFormMax.daRFM_Tapering
            strReturn = "Tapering - 2"
        Case DentAssess_RidgeFormMax.daRFM_Flat
            strReturn = "Flat - 3"
        Case DentAssess_RidgeFormMax.daRFM_Not_Addressed
            strReturn = Empty
        Case Else
            strReturn = Empty
    End Select
    
    GetRidgeFormMaxEnumDesc = strReturn
End Function
'******************************************************************
'Function: BuildRidgeFormMaxEnumList
'Author: KDunne
'Date: 12/14/2001
'Purpose: Builds a list of DentAssess_RidgeFormMax 's
'******************************************************************
Public Function BuildRidgeFormMaxEnumList() As String
    Dim X As Integer
    Dim strReturn As String
    
    For X = 0 To 3
        strReturn = strReturn & GetRidgeFormMaxEnumDesc(X) & "|"
    Next
    
    BuildRidgeFormMaxEnumList = " |" & strReturn
End Function

'*******************************************************************
'Function: GetRidgeFormManEnumDesc
'Author: KDunne
'Date: 12/11/2001
'Purpose. Returns the description for a given DentAssess_RidgeFormMan value.
'*******************************************************************
Public Function GetRidgeFormManEnumDesc(ByVal lngValue As _
        DentAssess_RidgeFormMan) As String
    Dim strReturn As String
    
    Select Case lngValue
        Case DentAssess_RidgeFormMan.daRFM_LessFavorable
            strReturn = "Less favorable - 2"
        Case DentAssess_RidgeFormMan.daRFM_favorable
            strReturn = "Favorable - 1"
        Case DentAssess_RidgeFormMan.daRFM_Unfavorable
            strReturn = "Unfavorable - 3"
        Case DentAssess_RidgeFormMan.daRFM_Not_Addressed
            strReturn = Empty
        Case Else
            strReturn = Empty
    End Select
    
    GetRidgeFormManEnumDesc = strReturn
End Function
'******************************************************************
'Function: BuildRidgeFormManEnumList
'Author: KDunne
'Date: 12/14/2001
'Purpose: Builds a list of DentAssess_RidgeFormMan 's
'******************************************************************
Public Function BuildRidgeFormManEnumList() As String
    Dim X As Integer
    Dim strReturn As String
    
    For X = 0 To 3
        strReturn = strReturn & GetRidgeFormManEnumDesc(X) & "|"
    Next
    
    BuildRidgeFormManEnumList = " |" & strReturn
End Function
'*******************************************************************
'Function: GetToriEnumDesc
'Author: KDunne
'Date: 12/11/2001
'Purpose. Returns the description for a given DentAssess_ToriEnum value.
'*******************************************************************
Public Function GetToriEnumDesc(ByVal lngValue As _
        DentAssess_ToriEnum) As String
    Dim strReturn As String
    
    Select Case lngValue
        Case DentAssess_ToriEnum.daT_Absent
            strReturn = "Absent - 1"
        Case DentAssess_ToriEnum.daT_Moderate
            strReturn = "Moderate - 2"
        Case DentAssess_ToriEnum.daT_Large
            strReturn = "Large - 3"
        Case DentAssess_ToriEnum.daT_Not_Addressed
            strReturn = Empty
        Case Else
            strReturn = Empty
    End Select
    
    GetToriEnumDesc = strReturn
End Function
'******************************************************************
'Function: BuildToriEnumList
'Author: KDunne
'Date: 12/14/2001
'Purpose: Builds a list of DentAssess_ToriEnum 's
'******************************************************************
Public Function BuildToriEnumList() As String
    Dim X As Integer
    Dim strReturn As String
    
    For X = 0 To 3
        strReturn = strReturn & GetToriEnumDesc(X) & "|"
    Next
    
    BuildToriEnumList = " |" & strReturn
End Function
'*******************************************************************
'Function: GetInterarchSpaceEnumDesc
'Author: KDunne
'Date: 12/11/2001
'Purpose. Returns the description for a given DentAssess_InterarchSpaceEnum value.
'*******************************************************************
Public Function GetInterarchSpaceEnumDesc(ByVal lngValue As _
        DentAssess_InterarchSpaceEnum) As String
    Dim strReturn As String
    
    Select Case lngValue
    
        Case DentAssess_InterarchSpaceEnum.daIS_Ideal
            strReturn = "Ideal/Accomodates the artificial teeth - 1"
        Case DentAssess_InterarchSpaceEnum.daIS_Excessive
            strReturn = "Excessive interarch space - 2"
        Case DentAssess_InterarchSpaceEnum.daIS_Insufficient
            strReturn = "Insufficient space - 3"
        Case DentAssess_InterarchSpaceEnum.daIS_Not_Addressed
            strReturn = Empty
        Case Else
            strReturn = Empty
    End Select
    
    GetInterarchSpaceEnumDesc = strReturn
End Function
'******************************************************************
'Function: BuildInterArchEnumList
'Author: KDunne
'Date: 12/14/2001
'Purpose: Builds a list of DentAssess_InterarchSpaceEnum 's
'******************************************************************
Public Function BuildInterArchEnumList() As String
    Dim X As Integer
    Dim strReturn As String
    
    For X = 0 To 3
        strReturn = strReturn & GetInterarchSpaceEnumDesc(X) & "|"
    Next
    
    BuildInterArchEnumList = " |" & strReturn
End Function
'*******************************************************************
'Function: GetRidgeParallelismEnumDesc
'Author: KDunne
'Date: 12/11/2001
'Purpose. Returns the description for a given DentAssess_RidgeParrellEnum value.
'*******************************************************************
Public Function GetRidgeParallelismEnumDesc(ByVal lngValue As _
        DentAssess_RidgeParrellEnum) As String
    Dim strReturn As String
    
    Select Case lngValue
    
        Case DentAssess_RidgeParrellEnum.daRP_BothDiverge
            strReturn = "Both ridges diverge anteriorly - 3"
        Case DentAssess_RidgeParrellEnum.daRP_ManDiverges
            strReturn = "Man. is divergent, anteriorly - 2"
        Case DentAssess_RidgeParrellEnum.daRP_Parallel
            strReturn = "Parallel - 1"
        Case DentAssess_RidgeParrellEnum.daRP_Not_Addressed
            strReturn = Empty
        Case Else
            strReturn = Empty
    End Select
    
    GetRidgeParallelismEnumDesc = strReturn
End Function
'******************************************************************
'Function: BuildRidgeParallelismEnumList
'Author: KDunne
'Date: 12/14/2001
'Purpose: Builds a list of DentAssess_RidgeParrellEnum 's
'******************************************************************
Public Function BuildRidgeParallelismEnumList() As String
    Dim X As Integer
    Dim strReturn As String
    
    For X = 0 To 3
        strReturn = strReturn & GetRidgeParallelismEnumDesc(X) & "|"
    Next
    
    BuildRidgeParallelismEnumList = " |" & strReturn
End Function
'*******************************************************************
'Function: GetRidgeRelationsEnumDesc
'Author: KDunne
'Date: 12/11/2001
'Purpose. Returns the description for a given DentAssess_RidgeRelationsEnum value.
'*******************************************************************
Public Function GetRidgeRelationsEnumDesc(ByVal lngValue As _
        DentAssess_RidgeRelationsEnum) As String
    Dim strReturn As String
    
    Select Case lngValue
    
        Case DentAssess_RidgeRelationsEnum.daRR_Normal
            strReturn = "Normal - 1"
        Case DentAssess_RidgeRelationsEnum.daRR_Prognathic
            strReturn = "Prognathic - 3"
        Case DentAssess_RidgeRelationsEnum.daRR_Retrognathic
            strReturn = "Retrognathic - 2"
        Case DentAssess_RidgeRelationsEnum.daRR_Not_Addressed
            strReturn = Empty
        Case Else
            strReturn = Empty
    End Select
    
    GetRidgeRelationsEnumDesc = strReturn
End Function
'******************************************************************
'Function: BuildRidgeRelationsEnumList
'Author: KDunne
'Date: 12/14/2001
'Purpose: Builds a list of DentAssess_RidgeRelationsEnum 's
'******************************************************************
Public Function BuildRidgeRelationsEnumList() As String
    Dim X As Integer
    Dim strReturn As String
    
    For X = 0 To 3
        strReturn = strReturn & GetRidgeRelationsEnumDesc(X) & "|"
    Next
    
    BuildRidgeRelationsEnumList = " |" & strReturn
End Function
'*******************************************************************
'Function: GetBoneQualityEnumDesc
'Author: KDunne
'Date: 12/11/2001
'Purpose. Returns the description for a given DentAssess_RidgeRelationsEnum value.
'*******************************************************************
Public Function GetBoneQualityEnumDesc(ByVal lngValue As _
        DentAssess_BoneQualityEnum) As String
    Dim strReturn As String
    
    Select Case lngValue
    
        Case DentAssess_BoneQualityEnum.daBQual_1
            strReturn = "1"
        Case DentAssess_BoneQualityEnum.daBQual_2
            strReturn = "2"
        Case DentAssess_BoneQualityEnum.daBQual_3
            strReturn = "3"
        Case DentAssess_BoneQualityEnum.daBQual_4
            strReturn = "4"
        Case Else
            strReturn = Empty
    End Select
    
    GetBoneQualityEnumDesc = strReturn
End Function
'******************************************************************
'Function: BuildBoneQualityEnumList
'Author: KDunne
'Date: 12/14/2001
'Purpose: Builds a list of DentAssess_BoneQualityEnum 's
'******************************************************************
Public Function BuildBoneQualityEnumList() As String
    Dim X As Integer
    Dim strReturn As String
    
    For X = 0 To 4
        strReturn = strReturn & GetBoneQualityEnumDesc(X) & "|"
    Next
    
    BuildBoneQualityEnumList = " |" & strReturn
End Function
'*******************************************************************
'Function: GetBoneQuantityEnumDesc
'Author: KDunne
'Date: 12/11/2001
'Purpose. Returns the description for a given DentAssess_BoneQuantityEnum value.
'*******************************************************************
Public Function GetBoneQuantityEnumDesc(ByVal lngValue As _
        DentAssess_BoneQuantityEnum) As String
    Dim strReturn As String
    
    Select Case lngValue
    
        Case DentAssess_BoneQuantityEnum.daBQuan_A
            strReturn = "A"
        Case DentAssess_BoneQuantityEnum.daBQuan_B
            strReturn = "B"
        Case DentAssess_BoneQuantityEnum.daBQuan_C
            strReturn = "C"
        Case DentAssess_BoneQuantityEnum.daBQuan_D
            strReturn = "D"
        Case DentAssess_BoneQuantityEnum.daBQuan_E
            strReturn = "E"
        Case Else
            strReturn = Empty
    End Select
    
    GetBoneQuantityEnumDesc = strReturn
End Function

'******************************************************************
'Function: BuildBoneQuantityEnumList
'Author: KDunne
'Date: 12/14/2001
'Purpose: Builds a list of DentAssess_BoneQuantityEnum 's
'******************************************************************
Public Function BuildBoneQuantityEnumList() As String
    Dim X As Integer
    Dim strReturn As String
    
    For X = 0 To 5
        strReturn = strReturn & GetBoneQuantityEnumDesc(X) & "|"
    Next
    
    BuildBoneQuantityEnumList = " |" & strReturn
End Function
'*******************************************************************
'Function: GetThroatFormEnumDesc
'Author: KDunne
'Date: 12/11/2001
'Purpose. Returns the description for a given DentAssess_ThroatFormEnum value.
'*******************************************************************
Public Function GetThroatFormEnumDesc(ByVal lngValue As _
        DentAssess_ThroatFormEnum) As String
    Dim strReturn As String
    
    Select Case lngValue
    
        Case DentAssess_ThroatFormEnum.daTF_ModerateAngle
            strReturn = "moderate sloping angle - 3"
        Case DentAssess_ThroatFormEnum.daTF_RightAngle
            strReturn = "Approx. a right angle in the posterior lingual vestibule - 1"
        Case DentAssess_ThroatFormEnum.daTF_SlightAngle
            strReturn = "slight sloping angle - 2"
        Case DentAssess_ThroatFormEnum.daTF_Not_Addressed
            strReturn = Empty
        Case Else
            strReturn = Empty
    End Select
    
    GetThroatFormEnumDesc = strReturn
End Function
'******************************************************************
'Function: BuildThroatFormEnumList
'Author: KDunne
'Date: 12/14/2001
'Purpose: Builds a list of DentAssess_ThroatFormEnum 's
'******************************************************************
Public Function BuildThroatFormEnumList() As String
    Dim X As Integer
    Dim strReturn As String
    
    For X = 0 To 3
        strReturn = strReturn & GetThroatFormEnumDesc(X) & "|"
    Next
    
    BuildThroatFormEnumList = " |" & strReturn
End Function
'*******************************************************************
'Function: GetPalatalThroatFormEnumDesc
'Author: KDunne
'Date: 12/11/2001
'Purpose. Returns the description for a given DentAssess_PalatalThroatFormEnum value.
'*******************************************************************
Public Function GetPalatalThroatFormEnumDesc(ByVal lngValue As _
        DentAssess_PalatalThroatFormEnum) As String
    Dim strReturn As String
    
    Select Case lngValue
    
        Case DentAssess_PalatalThroatFormEnum.daPTF_Large
            strReturn = "Large and normal in form, immovable band of resilient " & _
                            "tissue 5-12 mm D to line across the D of tuberosity - 1"
        Case DentAssess_PalatalThroatFormEnum.daPTF_Medium
            strReturn = "Medium in form, band 3-5 mm D of Tuberosity - 2"
        Case DentAssess_PalatalThroatFormEnum.daPTF_Small
            strReturn = "Usually a small maxilla, curtain 3-5 mm anterior to D of tuberosity - 3"
        Case DentAssess_PalatalThroatFormEnum.daPTF_Not_Addressed
            strReturn = Empty
        Case Else
            strReturn = Empty
    End Select
    
    GetPalatalThroatFormEnumDesc = strReturn
End Function
'******************************************************************
'Function: BuildPalatalThroatFormEnumList
'Author: KDunne
'Date: 12/14/2001
'Purpose: Builds a list of DentAssess_PalatalThroatFormEnum 's
'******************************************************************
Public Function BuildPalatalThroatFormEnumList() As String
    Dim X As Integer
    Dim strReturn As String
    
    For X = 0 To 3
        strReturn = strReturn & GetPalatalThroatFormEnumDesc(X) & "|"
    Next
    
    BuildPalatalThroatFormEnumList = " |" & strReturn
End Function
'*******************************************************************
'Function: GetPalatalSensitivityEnumDesc
'Author: KDunne
'Date: 12/11/2001
'Purpose. Returns the description for a given DentAssess_PalatalSensitiveEnum value.
'*******************************************************************
Public Function GetPalatalSensitivityEnumDesc(ByVal lngValue As _
        DentAssess_PalatalSensitiveEnum) As String
    Dim strReturn As String
    
    Select Case lngValue
    
        Case DentAssess_PalatalSensitiveEnum.daPS_Normal
            strReturn = "Normal - 1"
        Case DentAssess_PalatalSensitiveEnum.daPS_Sensitive
            strReturn = "Sensitive - 2"
        Case DentAssess_PalatalSensitiveEnum.daPS_HyperSensitive
            strReturn = "Hypersensitive - 3"
        Case DentAssess_PalatalSensitiveEnum.daPS_Not_Addressed
            strReturn = Empty
        Case Else
            strReturn = Empty
    End Select
    
    GetPalatalSensitivityEnumDesc = strReturn
End Function
'******************************************************************
'Function: BuildPalatalSensitivityEnumList
'Author: KDunne
'Date: 12/14/2001
'Purpose: Builds a list of DentAssess_PalatalSensitiveEnum 's
'******************************************************************
Public Function BuildPalatalSensitivityEnumList() As String
    Dim X As Integer
    Dim strReturn As String
    
    For X = 0 To 3
        strReturn = strReturn & GetPalatalSensitivityEnumDesc(X) & "|"
    Next
    
    BuildPalatalSensitivityEnumList = " |" & strReturn
End Function
'*******************************************************************
'Function: GetMucosaThicknessEnumDesc
'Author: KDunne
'Date: 12/11/2001
'Purpose. Returns the description for a given DentAssess_MucosaThickEnum value.
'*******************************************************************
Public Function GetMucosaThicknessEnumDesc(ByVal lngValue As _
        DentAssess_MucosaThickEnum) As String
    Dim strReturn As String
    
    Select Case lngValue
    
        Case DentAssess_MucosaThickEnum.daMT_Normal
            strReturn = "Normal, 1 mm thick - 1"
        Case DentAssess_MucosaThickEnum.daMT_Thick
            strReturn = "Thin membrane or 2x normal thickness - 2"
        Case DentAssess_MucosaThickEnum.daMT_ExcessiveThick
            strReturn = "Excessively thick w/redundant tissue - 3"
        Case DentAssess_MucosaThickEnum.daMT_Not_Addressed
            strReturn = Empty
        Case Else
            strReturn = Empty
    End Select
    
    GetMucosaThicknessEnumDesc = strReturn
End Function
'******************************************************************
'Function: BuildMucosaThicknessEnumList
'Author: KDunne
'Date: 12/14/2001
'Purpose: Builds a list of DentAssess_MucosaThickEnum 's
'******************************************************************
Public Function BuildMucosaThicknessEnumList() As String
    Dim X As Integer
    Dim strReturn As String
    
    For X = 0 To 3
        strReturn = strReturn & GetMucosaThicknessEnumDesc(X) & "|"
    Next
    
    BuildMucosaThicknessEnumList = " |" & strReturn
End Function
'*******************************************************************
'Function: GetMucosaConditionEnumDesc
'Author: KDunne
'Date: 12/11/2001
'Purpose. Returns the description for a given DentAssess_MucosaCondEnum value.
'*******************************************************************
Public Function GetMucosaConditionEnumDesc(ByVal lngValue As _
        DentAssess_MucosaCondEnum) As String
    Dim strReturn As String
    
    Select Case lngValue
    
        Case DentAssess_MucosaCondEnum.daMC_Healthy
            strReturn = "Healthy - 1"
        Case DentAssess_MucosaCondEnum.daMC_Irritated
            strReturn = "Irritated - 2"
        Case DentAssess_MucosaCondEnum.daMC_Pathologic
            strReturn = "Pathologic - 3"
        Case DentAssess_MucosaCondEnum.daMC_Not_Addressed
            strReturn = Empty
        Case Else
            strReturn = Empty
    End Select
    
    GetMucosaConditionEnumDesc = strReturn
End Function

'******************************************************************
'Function: BuildMucosaConditionEnumList
'Author: KDunne
'Date: 12/14/2001
'Purpose: Builds a list of DentAssess_MucosaCondEnum 's
'******************************************************************
Public Function BuildMucosaConditionEnumList() As String
    Dim X As Integer
    Dim strReturn As String
    
    For X = 0 To 3
        strReturn = strReturn & GetMucosaConditionEnumDesc(X) & "|"
    Next
    
    BuildMucosaConditionEnumList = " |" & strReturn
End Function
'*******************************************************************
'Function: GetBorderAttachEnumDesc
'Author: KDunne
'Date: 12/12/2001
'Purpose. Returns the description for a given DentAssess_BorderAttachEnum value.
'*******************************************************************
Public Function GetBorderAttachEnumDesc(ByVal lngValue As _
        DentAssess_BorderAttachEnum) As String
    Dim strReturn As String
    
    Select Case lngValue
    
        Case DentAssess_BorderAttachEnum.daBA_Attach1
            strReturn = "Attachments are high in maxilla or low in mandible with " & _
                        "relation to ridge crest (0.5 inches or more between level " & _
                        "of attachment and crest of ridge) - 1"
        Case DentAssess_BorderAttachEnum.daBA_Attach2
            strReturn = "Attachment height in relation to the crest of the ridge " & _
                        "is between 0.25 and 0.50 inches - 2"
        Case DentAssess_BorderAttachEnum.daBA_Attach3
            strReturn = "Attachment height is less than 0.25 inches from the " & _
                        "ridge crest - 3"
        Case DentAssess_BorderAttachEnum.daBA_Not_Addressed
            strReturn = Empty
        Case Else
            strReturn = Empty
    End Select
    
    GetBorderAttachEnumDesc = strReturn
End Function
'******************************************************************
'Function: BuildBorderAttachmentEnumList
'Author: KDunne
'Date: 12/14/2001
'Purpose: Builds a list of DentAssess_BorderAttachEnum 's
'******************************************************************
Public Function BuildBorderAttachmentEnumList() As String
    Dim X As Integer
    Dim strReturn As String
    
    For X = 0 To 3
        strReturn = strReturn & GetBorderAttachEnumDesc(X) & "|"
    Next
    
    BuildBorderAttachmentEnumList = " |" & strReturn
End Function
'*******************************************************************
'Function: GetFrenumAttachEnumDesc
'Author: KDunne
'Date: 12/12/2001
'Purpose. Returns the description for a given DentAssess_FrenumAttachEnum value.
'*******************************************************************
Public Function GetFrenumAttachEnumDesc(ByVal lngValue As _
        DentAssess_FrenumAttachEnum) As String
    Dim strReturn As String
    
    Select Case lngValue
    
        Case DentAssess_FrenumAttachEnum.daFA_High
            strReturn = "High in the maxilla or low in the mandible with " & _
                        "respect to the crest of the ridge - 1"
        Case DentAssess_FrenumAttachEnum.daFA_Medium
            strReturn = "Medium - 2"
        Case DentAssess_FrenumAttachEnum.daFA_FreniEncroach
            strReturn = "Freni encroach on the crest of the ridge and may " & _
                        "interfere with the denture seal.  Surgical correction " & _
                        "may be required - 3"
        Case DentAssess_FrenumAttachEnum.daFA_Not_Addressed
            strReturn = Empty
        Case Else
            strReturn = Empty
    End Select
    
    GetFrenumAttachEnumDesc = strReturn
End Function
'******************************************************************
'Function: BuildFrenumAttachmentEnumList
'Author: KDunne
'Date: 12/14/2001
'Purpose: Builds a list of DentAssess_FrenumAttachEnum 's
'******************************************************************
Public Function BuildFrenumAttachmentEnumList() As String
    Dim X As Integer
    Dim strReturn As String
    
    For X = 0 To 3
        strReturn = strReturn & GetFrenumAttachEnumDesc(X) & "|"
    Next
    
    BuildFrenumAttachmentEnumList = " |" & strReturn
End Function
'*******************************************************************
'Function: GetSalivaEnumDesc
'Author: KDunne
'Date: 12/12/2001
'Purpose. Returns the description for a given DentAssess_SalivaEnum value.
'*******************************************************************
Public Function GetSalivaEnumDesc(ByVal lngValue As _
        DentAssess_SalivaEnum) As String
    Dim strReturn As String
    
    Select Case lngValue
    
        Case DentAssess_SalivaEnum.daS_Normal
            strReturn = "Normal - 1"
        Case DentAssess_SalivaEnum.daS_Excessive
            strReturn = "Excessive - 2"
        Case DentAssess_SalivaEnum.daS_Xerostomia
            strReturn = "Xerostomia - 3"
        Case DentAssess_SalivaEnum.daS_Not_Addressed
            strReturn = Empty
        Case Else
            strReturn = Empty
    End Select
    
    GetSalivaEnumDesc = strReturn
End Function

'******************************************************************
'Function: BuildSalivaEnumList
'Author: KDunne
'Date: 12/14/2001
'Purpose: Builds a list of DentAssess_SalivaEnum 's
'******************************************************************
Public Function BuildSalivaEnumList() As String
    Dim X As Integer
    Dim strReturn As String
    
    For X = 0 To 3
        strReturn = strReturn & GetSalivaEnumDesc(X) & "|"
    Next
    
    BuildSalivaEnumList = " |" & strReturn
End Function

'*******************************************************************
'Function: GetTongueSizeEnumDesc
'Author: KDunne
'Date: 12/12/2001
'Purpose. Returns the description for a given DentAssess_TongueSizeEnum value.
'*******************************************************************
Public Function GetTongueSizeEnumDesc(ByVal lngValue As _
        DentAssess_TongueSizeEnum) As String
    Dim strReturn As String
    
    Select Case lngValue
    
        Case DentAssess_TongueSizeEnum.daTS_Normal
            strReturn = "Normal in size, development and function (pt. w/teeth) - 1"
        Case DentAssess_TongueSizeEnum.daTS_TeethAbsent
            strReturn = "Teeth absent for long time, tongue form and function " & _
                        "change - 2"
        Case DentAssess_TongueSizeEnum.daTS_Large
            strReturn = "Excessively large tongue - 3"
        Case DentAssess_TongueSizeEnum.daTS_Not_Addressed
            strReturn = Empty
        Case Else
            strReturn = Empty
    End Select
    
    GetTongueSizeEnumDesc = strReturn
End Function
'******************************************************************
'Function: BuildTongueSizeEnumList
'Author: KDunne
'Date: 12/14/2001
'Purpose: Builds a list of DentAssess_TongueSizeEnum 's
'******************************************************************
Public Function BuildTongueSizeEnumList() As String
    Dim X As Integer
    Dim strReturn As String
    
    For X = 0 To 3
        strReturn = strReturn & GetTongueSizeEnumDesc(X) & "|"
    Next
    
    BuildTongueSizeEnumList = " |" & strReturn
End Function
'*******************************************************************
'Function: GetTonguePosEnumDesc
'Author: KDunne
'Date: 12/12/2001
'Purpose. Returns the description for a given DentAssess_TonguePosEnum value.
'*******************************************************************
Public Function GetTonguePosEnumDesc(ByVal lngValue As _
        DentAssess_TonguePosEnum) As String
    Dim strReturn As String
    
    Select Case lngValue
    
        Case DentAssess_TonguePosEnum.daTP_Normal
            strReturn = "Normal = fills floor of mouth, lateral border on occl " & _
                        "plane, apex on incisal edges - 1"
        Case DentAssess_TonguePosEnum.daTP_Class1Retracted
            strReturn = "Class 1 retracted = Floor or mouth pulled downward, " & _
                        "exposed back to 1st molar, borders above occl plane, " & _
                        "apex down to floor - 2"
        Case DentAssess_TonguePosEnum.daTP_Class2Retracted
            strReturn = "Class 2 retracted = Tense, backward, upward, apex " & _
                        "into body of tongue - 3"
        Case DentAssess_TonguePosEnum.daTP_Not_Addressed
            strReturn = Empty
        Case Else
            strReturn = Empty
    End Select
    
    GetTonguePosEnumDesc = strReturn
End Function
'******************************************************************
'Function: BuildTonguePosEnumList
'Author: KDunne
'Date: 12/14/2001
'Purpose: Builds a list of DentAssess_TonguePosEnum 's
'******************************************************************
Public Function BuildTonguePosEnumList() As String
    Dim X As Integer
    Dim strReturn As String
    
    For X = 0 To 3
        strReturn = strReturn & GetTonguePosEnumDesc(X) & "|"
    Next
    
    BuildTonguePosEnumList = " |" & strReturn
End Function
'*******************************************************************
'Function: GetGoodFairPoorEnumDesc
'Author: KDunne
'Date: 12/12/2001
'Purpose. Returns the description for a given DentAssess_GoodFairPoorEnum value.
'*******************************************************************
Public Function GetGoodFairPoorEnumDesc(ByVal lngValue As _
        DentAssess_GoodFairPoorEnum) As String
    Dim strReturn As String
    
    Select Case lngValue
    
        Case DentAssess_GoodFairPoorEnum.daGFP_Good
            strReturn = "Good - 1"
        Case DentAssess_GoodFairPoorEnum.daGFP_Fair
            strReturn = "Fair - 2"
        Case DentAssess_GoodFairPoorEnum.daGFP_Poor
            strReturn = "Poor - 3"
        Case DentAssess_GoodFairPoorEnum.daGFP_Not_Addressed
            strReturn = Empty
        Case Else
            strReturn = Empty
    End Select
    
    GetGoodFairPoorEnumDesc = strReturn
End Function
'******************************************************************
'Function: BuildGoodFairPoorEnumList
'Author: KDunne
'Date: 12/14/2001
'Purpose: Builds a list of DentAssess_GoodFairPoorEnum 's
'******************************************************************
Public Function BuildGoodFairPoorEnumList() As String
    Dim X As Integer
    Dim strReturn As String
    
    For X = 0 To 3
        strReturn = strReturn & GetGoodFairPoorEnumDesc(X) & "|"
    Next
    
    BuildGoodFairPoorEnumList = " |" & strReturn
End Function
'*******************************************************************
'Function: GetAcceptablityEnumDesc
'Author: KDunne
'Date: 12/12/2001
'Purpose. Returns the description for a given DentAssess_AcceptablityEnum value.
'*******************************************************************
Public Function GetAcceptablityEnumDesc(ByVal lngValue As _
        DentAssess_AcceptablityEnum) As String
    Dim strReturn As String
    
    Select Case lngValue
    
        Case DentAssess_AcceptablityEnum.daAcc_Acceptable
            strReturn = "Acceptable"
        Case DentAssess_AcceptablityEnum.daAcc_Unacceptable
            strReturn = "Unacceptable"
        Case DentAssess_AcceptablityEnum.daAcc_Not_Addressed
            strReturn = Empty
        Case Else
            strReturn = Empty
    End Select
    
    GetAcceptablityEnumDesc = strReturn
End Function
'******************************************************************
'Function: BuildAcceptabilityEnumList
'Author: KDunne
'Date: 12/14/2001
'Purpose: Builds a list of DentAssess_AcceptablityEnum 's
'******************************************************************
Public Function BuildAcceptabilityEnumList() As String
    Dim X As Integer
    Dim strReturn As String
    
    For X = 0 To 3
        strReturn = strReturn & GetAcceptablityEnumDesc(X) & "|"
    Next
    
    BuildAcceptabilityEnumList = " |" & strReturn
End Function
'*******************************************************************
'Function: GetVDOEnumDesc
'Author: KDunne
'Date: 12/12/2001
'Purpose. Returns the description for a given DentAssess_VDOEnum value.
'*******************************************************************
Public Function GetVDOEnumDesc(ByVal lngValue As _
        DentAssess_VDOEnum) As String
    Dim strReturn As String
    
    Select Case lngValue
        Case DentAssess_VDOEnum.daVDO_Adequate
            strReturn = "Adequate"
        Case DentAssess_VDOEnum.daVDO_Inadequate
            strReturn = "Inadequate"
        Case DentAssess_VDOEnum.daVDO_Excessive
            strReturn = "Excessive"
        Case DentAssess_VDOEnum.daVDO_Not_Addressed
            strReturn = Empty
        Case Else
            strReturn = Empty
    End Select
    
    GetVDOEnumDesc = strReturn
End Function
'******************************************************************
'Function: BuildVDOEnumList
'Author: KDunne
'Date: 12/14/2001
'Purpose: Builds a list of DentAssess_VDOEnum 's
'******************************************************************
Public Function BuildVDOEnumList() As String
    Dim X As Integer
    Dim strReturn As String
    
    For X = 0 To 3
        strReturn = strReturn & GetVDOEnumDesc(X) & "|"
    Next
    
    BuildVDOEnumList = " |" & strReturn
End Function
'*******************************************************************
'Function: GetCrossbiteEnumDesc
'Author: KDunne
'Date: 12/12/2001
'Purpose. Returns the description for a given DentAssess_CrossbiteEnum value.
'*******************************************************************
Public Function GetCrossbiteEnumDesc(ByVal lngValue As _
        DentAssess_CrossbiteEnum) As String
    Dim strReturn As String
    
    Select Case lngValue
    
        Case DentAssess_CrossbiteEnum.daCB_None
            strReturn = "None"
        Case DentAssess_CrossbiteEnum.daCB_Unilateral
            strReturn = "Unilateral"
        Case DentAssess_CrossbiteEnum.daCB_Bilateral
            strReturn = "Bilateral"
        Case DentAssess_CrossbiteEnum.daCB_Interior
            strReturn = "Interior"
        Case DentAssess_CrossbiteEnum.daCB_Not_Addressed
            strReturn = Empty
        Case Else
            strReturn = Empty
    End Select
    
    GetCrossbiteEnumDesc = strReturn
End Function
'******************************************************************
'Function: BuildCrossbiteEnumList
'Author: KDunne
'Date: 12/14/2001
'Purpose: Builds a list of DentAssess_CrossbiteEnum 's
'******************************************************************
Public Function BuildCrossbiteEnumList() As String
    Dim X As Integer
    Dim strReturn As String
    
    For X = 0 To 4
        strReturn = strReturn & GetCrossbiteEnumDesc(X) & "|"
    Next
    
    BuildCrossbiteEnumList = " |" & strReturn
End Function
'*******************************************************************
'Function: GetCharacterizationEnumDesc
'Author: KDunne
'Date: 12/12/2001
'Purpose. Returns the description for a given DentAssess_CharacterizationEnum value.
'*******************************************************************
Public Function GetCharacterizationEnumDesc(ByVal lngValue As _
        DentAssess_CharacterizationEnum) As String
    Dim strReturn As String
    
    Select Case lngValue
    
        Case DentAssess_CharacterizationEnum.daChar_Characterized
            strReturn = "Characterized"
        Case DentAssess_CharacterizationEnum.daChar_Uncharacterized
            strReturn = "Uncharacterized"
        Case DentAssess_CharacterizationEnum.daChar_Not_Addressed
            strReturn = Empty
        Case Else
            strReturn = Empty
    End Select
    
    GetCharacterizationEnumDesc = strReturn
End Function

'******************************************************************
'Function: BuildCharacterizationEnumList
'Author: KDunne
'Date: 12/14/2001
'Purpose: Builds a list of DentAssess_CharacterizationEnum 's
'******************************************************************
Public Function BuildCharacterizationEnumList() As String
    Dim X As Integer
    Dim strReturn As String
    
    For X = 0 To 4
        strReturn = strReturn & GetCharacterizationEnumDesc(X) & "|"
    Next
    
    BuildCharacterizationEnumList = " |" & strReturn
End Function

'*******************************************************************
'Function: GetCentricRelationEnumDesc
'Author: KDunne
'Date: 12/18/2001
'Purpose. Returns the description for a given DentAssess_CentricRelationEnum value.
'*******************************************************************
Public Function GetCentricRelationEnumDesc(ByVal lngValue As _
    DentAssess_CentricRelationEnum) As String
        
    Dim strReturn As String
    
    Select Case lngValue
        Case DentAssess_CentricRelationEnum.daCR_CR1
            strReturn = "CR=MI"
        Case DentAssess_CentricRelationEnum.daCR_CR2
            strReturn = "CR*MI"
        Case DentAssess_CentricRelationEnum.daCR_Not_Addressed
            strReturn = Empty
        Case Else
            strReturn = Empty
    End Select

    GetCentricRelationEnumDesc = strReturn
End Function
'*******************************************************************
'Function: GetDentureOcclPlaneEnumDesc
'Author: KDunne
'Date: 12/18/2001
'Purpose. Returns the description for a given DentAssess_DentureOcclEnum value.
'*******************************************************************
Public Function GetDentureOcclPlaneEnumDesc(ByVal lngValue As _
        DentAssess_DentureOcclPlaneEnum) As String
    Dim strReturn As String
    
    Select Case lngValue
        Case DentAssess_DentureOcclPlaneEnum.daDOP_Adequate
            strReturn = "Adequate"
        Case DentAssess_DentureOcclPlaneEnum.daDOP_High
            strReturn = "High"
        Case DentAssess_DentureOcclPlaneEnum.daDOP_Low
            strReturn = "Low"
        Case DentAssess_DentureOcclPlaneEnum.daDOP_Not_Addressed
            strReturn = Empty
        Case Else
            strReturn = Empty
    End Select
    
    GetDentureOcclPlaneEnumDesc = strReturn
End Function
        
'*******************************************************************
'Function: GetDentureWearEnumDesc
'Author: KDunne
'Date: 12/12/2001
'Purpose. Returns the description for a given DentAssess_DentureWearEnum value.
'History:
'       4/2/2002 - KDunne, scr 22623
'*******************************************************************
Public Function GetDentureWearEnumDesc(ByVal lngValue As _
        DentAssess_DentureWearEnum) As String
    Dim strReturn As String
    
    Select Case lngValue
    
        Case DentAssess_DentureWearEnum.daDW_Minimal
            strReturn = "Minimum - 1"
        Case DentAssess_DentureWearEnum.daDW_Moderate
            strReturn = "Moderate - 2"
        Case DentAssess_DentureWearEnum.daDW_Severe
            strReturn = "Severe - 3"
        Case DentAssess_DentureWearEnum.daDW_Not_Addressed
            strReturn = Empty
        Case Else
            strReturn = Empty
    End Select
    
    GetDentureWearEnumDesc = strReturn
End Function
'*****************************************************************************
'Function:
'Author: KDunne
'Date: 12/18/2001
'Purpose: Builds a list of DentAssess_CentricRelationsEnum 's
'******************************************************************
Public Function BuildCentricRelationsEnumList() As String
    Dim X As Integer
    Dim strReturn As String
    
    For X = 0 To 2
        strReturn = strReturn & GetCentricRelationEnumDesc(X) & "|"
    Next
    
    BuildCentricRelationsEnumList = " |" & strReturn
End Function
'******************************************************************
'Function: BuildDentureOcclPlaneEnumList
'Author: KDunne
'Date: 12/18/2001
'Purpose: Builds a list of DentAssess_DentureOcclPlaneEnum 's
'******************************************************************
Public Function BuildDentureOcclPlaneEnumList() As String
    Dim X As Integer
    Dim strReturn As String
    
    For X = 0 To 3
        strReturn = strReturn & GetDentureOcclPlaneEnumDesc(X) & "|"
    Next
    
    BuildDentureOcclPlaneEnumList = " |" & strReturn
End Function
'******************************************************************
'Function: BuildDentureWearEnumList
'Author: KDunne
'Date: 12/14/2001
'Purpose: Builds a list of DentAssess_DentureWearEnum 's
'******************************************************************
Public Function BuildDentureWearEnumList() As String
    Dim X As Integer
    Dim strReturn As String
    
    For X = 0 To 3
        strReturn = strReturn & GetDentureWearEnumDesc(X) & "|"
    Next
    
    BuildDentureWearEnumList = " |" & strReturn
End Function

