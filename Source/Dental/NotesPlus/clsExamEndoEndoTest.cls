VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExamEndoTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mlngTestId As Long
Private mlngDentalEncounterId As Long
Private mlngProviderId As Long
Private mlngPatientID As Long

Private mstrToothNumber As String
Private mlngEPT As ExamEndoEPTEnum
Private mlngColdResponse As ExamEndoTestResponseEnum
Private mlngColdLinger As ExamEndoYesNoEnum
Private mlngColdResponseTest As ExamEndoColdResponseTestEnum
Private mlngHeatResponseTest As ExamEndoHeatResponseTestEnum
Private mlngHeatResponse As ExamEndoTestResponseEnum
Private mlngHeatLinger As ExamEndoYesNoEnum
Private mlngPercussion As ExamEndoTestResponseEnum
Private mlngPainPalpation As ExamEndoYesNoEnum
Private mstrPalpationLoc As String
Private mlngSinusTract As ExamEndoYesNoEnum
Private mstrSTTracedTo As String
Private mstrSTComments As String
Private mlngSTTermination As ExamEndoSTTermination
Private mstrMagnify As String
Private mstrTransIllum As String
Private mlngMobility As ExamEndoMobilityEnum
Private mstrProbing As String
Private mlngPainOcclusion As ExamEndoYesNoEnum
Private mstrPainOcclusionLoc As String
Private mlngPainRelease As ExamEndoYesNoEnum
Private mstrPainReleaseLoc As String
Private mlngCaries As ExamEndoYesNoEnum
Private mlngRadiographFind As ExamEndoYesNoEnum
Private mstrRadiographFindCmnt As String
Private mlngTrauma As ExamEndoTraumaEnum
Private mstrAnatomicVar As String
Private mstrRestoreHx As String
Private mlngMicroleak As ExamEndoYesNoEnum
Private mlngBleachHx As ExamEndoYesNoEnum
Private mstrHxComments As String
Private mblnUnRestoredCaries As Boolean
Private mobjSvgComm As srvSVGComm.clsSVGComm
Private mobjSvgFac As New srvSVGComm.SingletonFactory
Private mlngDirtyFlag As DirtyFlagEnum
Private mblnWroteColdTest As Boolean
Private mblnHeatTestWritten As Boolean

'Workload additions
Private mlWorkloadID As Long

Private mobjDentalNotes As CHCSII_Dental_DXTX.clsNotes
Private mobjDx As CHCSII_Dental_DXTX.clsDx
Private mobjGroup As CHCSII_Dental_DXTX.clsGroup

Public Enum ExamEndoYesNoEnum
    Not_Addressed = 0
    Yes = 1
    No = 2
End Enum
Public Enum ExamEndoEPTEnum
    exEndoEPT_Not_Addressed
    Test0
    Test1
    Test2
    Test3
    Test4
    Test5
    Test6
    Test7
    Test8
    Test9
    Test10
    Test11
    Test12
    Test13
    Test14
    Test15
    Test16
    Test17
    Test18
    Test19
    Test20
    Test21
    Test22
    Test23
    Test24
    Test25
    Test26
    Test27
    Test28
    Test29
    Test30
    Test31
    Test32
    Test33
    Test34
    Test35
    Test36
    Test37
    Test38
    Test39
    Test40
    Test41
    Test42
    Test43
    Test44
    Test45
    Test46
    Test47
    Test48
    Test49
    Test50
    Test51
    Test52
    Test53
    Test54
    Test55
    Test56
    Test57
    Test58
    Test59
    Test60
    Test61
    Test62
    Test63
    Test64
    Test65
    Test66
    Test67
    Test68
    Test69
    Test70
    Test71
    Test72
    Test73
    Test74
    Test75
    Test76
    Test77
    Test78
    Test79
    Test80
    NoResponse
End Enum

Public Enum ExamEndoTestResponseEnum
    Not_Addressed = 0
    Minus = 1
    Plus = 2
    PlusPlus = 3
    PlusPlusPlus = 4
    Delayed = 5
   
End Enum
'KDunne 8/12/2002
'scr 17239
Public Enum ExamEndoPercussionEnum
    Not_Addressed = 0
    Minus = 1
    Plus = 2
    PlusPlus = 3
    PlusPlusPlus = 4
End Enum

Public Enum ExamEndoColdResponseTestEnum
    Not_Addressed = 0
    CarbonDioxide = 1
    EthylChloride = 2
    Ice = 3
End Enum

Public Enum ExamEndoHeatResponseTestEnum
    Not_Addressed = 0
    RD_HotFluid = 1
    Friction = 2
    HotGuttaPercha = 3
End Enum

Public Enum ExamEndoSTTermination
    Not_Addressed = 0
    OralCavity = 1
    Skin = 2
    Nasal = 3
    MaxilaryAntrim = 4
End Enum
Public Enum ExamEndoMobilityEnum
    Not_Addressed = 0
    Mobility1 = 1
    Mobility2 = 2
    Mobility3 = 3
    WNL = 4
    ANK = 5
End Enum
Public Enum ExamEndoTraumaEnum
    Not_Addressed = 0
    Yes = 1
    No = 2
    Concussion = 3
    Subluxation = 4
    LateralLuxation = 5
    ExtrusiveLuxation = 6
    IntrusiveLuxation = 7
    Avulsion = 8
    HorizontalRootFracture = 9
    UncomplicatedFracture = 10
    ComplicatedFracture = 11
End Enum
Public Enum DxTxMobilityICDEnums
    Not_Addressed = 0
    MillerMobility1 = 1634
    MillerMobility2 = 1625
    MillerMobility3 = 1623
    Ankylosis = 895
End Enum

Public Property Let WorkloadId(vData As Long)

    mlWorkloadID = vData
    UpdateDirtyFlag
    
End Property

Public Property Get WorkloadId() As Long

    WorkloadId = mlWorkloadID
    'UpdateDirtyFlag
    
End Property

Private Sub Class_Initialize()
    Initialize_Values
    Set mobjSvgComm = mobjSvgFac.Singleton
End Sub
Public Property Get Dirty_Flag() As DirtyFlagEnum
    Dirty_Flag = mlngDirtyFlag
End Property
Public Property Let Dirty_Flag(ByVal lngValue As DirtyFlagEnum)
    mlngDirtyFlag = lngValue
End Property
Public Property Get Test_Id() As Long
    Test_Id = mlngTestId
End Property

Public Property Let Test_Id(ByVal lngValue As Long)
    mlngTestId = lngValue
End Property

Public Property Get Dental_EncounterId() As Long
    Dental_EncounterId = mlngDentalEncounterId
End Property

Public Property Let Dental_EncounterId(ByVal lngValue As Long)
    mlngDentalEncounterId = lngValue
End Property
Public Property Let PatientID(ByVal lngValue As Long)
    mlngPatientID = lngValue
End Property
Public Property Get PatientID() As Long
    PatientID = mlngPatientID
End Property
Public Property Get ProviderId() As Long
    ProviderId = mlngProviderId
End Property
Public Property Let ProviderId(ByVal lngValue As Long)
    mlngProviderId = lngValue
End Property
Public Property Get Tooth_Number() As String
    Tooth_Number = mstrToothNumber
End Property

Public Property Let Tooth_Number(ByVal strValue As String)
    mstrToothNumber = strValue
    UpdateDirtyFlag
    If Trim(strValue) <> "" Then
        CheckForUnRestoredCaries CLng(strValue)
        CheckForMobility CLng(strValue)
    End If
    
End Property

Public Property Get EPT() As ExamEndoEPTEnum
    If mlngEPT < 0 Then
        mlngEPT = ExamEndoEPTEnum.exEndoEPT_Not_Addressed
    End If
    EPT = mlngEPT
End Property

Public Property Let EPT(ByVal lngValue As ExamEndoEPTEnum)
      mlngEPT = lngValue
    UpdateDirtyFlag
End Property
Public Property Get Cold_Response_Test() As ExamEndoColdResponseTestEnum
    Cold_Response_Test = mlngColdResponseTest
End Property

Public Property Let Cold_Response_Test(ByVal lngValue As ExamEndoColdResponseTestEnum)
    mlngColdResponseTest = lngValue
    UpdateDirtyFlag
End Property

Public Property Get Cold_Response() As ExamEndoTestResponseEnum
    Cold_Response = mlngColdResponse
End Property

Public Property Let Cold_Response(ByVal lngValue As ExamEndoTestResponseEnum)
    mlngColdResponse = lngValue
    UpdateDirtyFlag
End Property

Public Property Get Cold_Linger() As ExamEndoYesNoEnum
    Cold_Linger = mlngColdLinger
End Property

Public Property Let Cold_Linger(ByVal lngValue As ExamEndoYesNoEnum)
    mlngColdLinger = lngValue
    UpdateDirtyFlag
End Property

Public Property Get Heat_Response() As ExamEndoTestResponseEnum
    Heat_Response = mlngHeatResponse
End Property

Public Property Let Heat_Response(ByVal lngValue As ExamEndoTestResponseEnum)
    mlngHeatResponse = lngValue
    UpdateDirtyFlag
End Property

Public Property Get Heat_Linger() As ExamEndoYesNoEnum
    Heat_Linger = mlngHeatLinger
End Property

Public Property Let Heat_Linger(ByVal lngValue As ExamEndoYesNoEnum)
    mlngHeatLinger = lngValue
    UpdateDirtyFlag
End Property

Public Property Get Percussion() As ExamEndoPercussionEnum
    Percussion = mlngPercussion
End Property

Public Property Let Percussion(ByVal lngValue As ExamEndoPercussionEnum)
    mlngPercussion = lngValue
    UpdateDirtyFlag
End Property

Public Property Get Pain_On_Palpation() As ExamEndoYesNoEnum
    Pain_On_Palpation = mlngPainPalpation
End Property

Public Property Let Pain_On_Palpation(ByVal lngValue As ExamEndoYesNoEnum)
    mlngPainPalpation = lngValue
    UpdateDirtyFlag
End Property

Public Property Get Palpation_Location() As String
    Palpation_Location = mstrPalpationLoc
End Property

Public Property Let Palpation_Location(ByVal strValue As String)
    If Len(strValue) > 2000 Then
        goCWShared.ShowVBError 1000, "Palpation_Location string too long", "Notes+.clseExamEndoEndoTest", "String too long", vbInformation
    Else
        mstrPalpationLoc = strValue
        UpdateDirtyFlag
    End If
End Property

Public Property Get Sinus_Tract() As ExamEndoYesNoEnum
    Sinus_Tract = mlngSinusTract
End Property

Public Property Let Sinus_Tract(ByVal lngValue As ExamEndoYesNoEnum)
    mlngSinusTract = lngValue
    UpdateDirtyFlag
End Property

Public Property Get ST_Traced_To() As String
    ST_Traced_To = mstrSTTracedTo
End Property

Public Property Let ST_Traced_To(ByVal strValue As String)
    If Len(strValue) > 2000 Then
        goCWShared.ShowVBError 1000, "ST_Traced_To string too long", "Notes+.clseExamEndoEndoTest", "String too long", vbInformation
    Else
        mstrSTTracedTo = strValue
        UpdateDirtyFlag
    End If
End Property

Public Property Get ST_Comments() As String
    ST_Comments = mstrSTComments
End Property

Public Property Let ST_Comments(ByVal strValue As String)
    If Len(strValue) > 2000 Then
        goCWShared.ShowVBError 1000, "ST_Comments string too long", "Notes+.clseExamEndoEndoTest", "String too long", vbInformation
    Else
        mstrSTComments = strValue
        UpdateDirtyFlag
    End If
End Property

Public Property Get ST_Termination() As ExamEndoSTTermination
    ST_Termination = mlngSTTermination
End Property

Public Property Let ST_Termination(ByVal lngValue As ExamEndoSTTermination)
    mlngSTTermination = lngValue
    UpdateDirtyFlag
End Property

Public Property Get Magnification_Findings() As String
    Magnification_Findings = mstrMagnify
End Property

Public Property Let Magnification_Findings(ByVal strValue As String)
    If Len(strValue) > 2000 Then
        goCWShared.ShowVBError 1000, "Magnification_Findings string too long", "Notes+.clseExamEndoEndoTest", "String too long", vbInformation
    Else
        mstrMagnify = strValue
        UpdateDirtyFlag
    End If
End Property

Public Property Get Transillumination() As String
    Transillumination = mstrTransIllum
End Property

Public Property Let Transillumination(ByVal strValue As String)
    If Len(strValue) > 1000 Then
        goCWShared.ShowVBError 1000, "Transillumination string too long", "Notes+.clseExamEndoEndoTest", "String too long", vbInformation
    Else
        mstrTransIllum = strValue
        UpdateDirtyFlag
    End If
End Property

Public Property Get Mobility() As ExamEndoMobilityEnum
    'First, refresh mobility info from the grid.
    CheckForMobility CLng(mstrToothNumber)
    Mobility = mlngMobility
End Property

Public Property Let Mobility(ByVal lngValue As ExamEndoMobilityEnum)
    
    Dim lngDXCodeId As Long
    Dim objSS As clsSurfaceSet
    Dim objTooth As CHCSII_Dental_DXTX.clsTooth
    Dim intMenuItemId As Integer
    Dim lngGroupCnt As Long
    
    If lngValue = -1 Then Exit Property
    
    lngDXCodeId = GetMobilityDxCode(lngValue)
    intMenuItemId = GetMobilityMenuItemId(lngValue)
    
    'This must be set before we add/delete just incase it was removed in GC
    'SF 9-16-03
    CheckForMobility mstrToothNumber
    
    If mobjDx Is Nothing = False Then
        'Within Big Tooth, whenever mobility is changed, the
        'provider must first delete the existing Dx and then
        'add a new dx with the new mobility. Here, we do that
        'automatically. Ok Per Paul. KDunne 3/7/2002
        mobjDx.StatusID = dxNonActiveDeleted

        UpdateSVGGraphic mobjSvgComm, mobjDx, Nothing, Nothing, CInt(mstrToothNumber), "777", True
        
        Set mobjDx = Nothing
        DoEvents
    End If
    If lngValue <> ExamEndoMobilityEnum.WNL Then

            'Need to check to make sure that user didn't set a new
            'dx to blank. Can't add a blank Dx to the Datalayer.
            If lngValue <> ExamEndoMobilityEnum.Not_Addressed Then
                'Add Dx to the data layer.
                If mobjGroup Is Nothing Then
'                    lngGroupCnt = mobjDentalNotes.Groups.Count
'                    If lngGroupCnt > 0 Then
'                        Set mobjGroup = mobjDentalNotes.Groups(lngGroupCnt)
'                    Else
                        Set mobjGroup = mobjDentalNotes.Groups.Add
'                    End If
                End If
                Set mobjDx = mobjGroup.Dxs.Add(sNotes, lngDXCodeId, mlngProviderId)
                mobjDx.ProviderId = mlngProviderId
                mobjDx.DentalEncounterID = mlngDentalEncounterId
                Set objSS = mobjDx.SurfaceSets.Add(0)
                
                mobjDx.Chartable = True
                
                objSS.SurfaceSet = "777"
                objSS.State = False
                
                objSS.SurfaceAbbr = " "
                objSS.SymbolID = GetMobilityMenuItemId(lngValue)
                Set objTooth = objSS.Teeth.Add(mlngDentalEncounterId)
                objTooth.ToothID = CLng(mstrToothNumber)
                
                'Must update SVG Com object in order to get the tooth name.
'                If lngValue = ExamEndoMobilityEnum.ANK Then
'                    Call mobjSvgComm.ToothObject.ToothChanged(CInt(mstrToothNumber))
'                Else
                    UpdateSVGGraphic mobjSvgComm, mobjDx, Nothing, objSS, CInt(mstrToothNumber), "777", False
'                End If
                objTooth.ToothName = mobjSvgComm.ToothObject.getToothName
                
                mobjDx.Chartable = mobjDx.Chartable
                objSS.SurfaceSet = "777"
            
            End If

    End If 'lngValue
    
      If Not mobjDx Is Nothing And lngValue = ExamEndoMobilityEnum.WNL Then

        UpdateSVGGraphic mobjSvgComm, mobjDx, Nothing, objSS, CInt(mstrToothNumber), "777", False

      End If
      
    Set objTooth = Nothing
    Set objSS = Nothing

    mlngMobility = lngValue
    UpdateDirtyFlag
End Property

'************************************************************************
'Function:GetMobilityMenuItemId
'Author:KDunne
'Date: 3/7/2002
'Purpose: Returns the menu item id for Mobility 1, 2 or 3. Returns 0 for
'         all  other Moblititys.
'*************************************************************************
Private Function GetMobilityMenuItemId(ByVal lngMobility As ExamEndoMobilityEnum) As Integer
    Dim intReturn As Integer
    
    Select Case lngMobility
        Case ExamEndoMobilityEnum.Mobility1
            intReturn = 142
        Case ExamEndoMobilityEnum.Mobility2
            intReturn = 143
        Case ExamEndoMobilityEnum.Mobility3
            intReturn = 145
        'SFoster 2-25-03 SCR 32956, if it draws, it needs a MenuItemID
        Case ExamEndoMobilityEnum.ANK
            intReturn = 1141
        Case Else
            intReturn = 0
    End Select
    
    GetMobilityMenuItemId = intReturn
End Function
Public Property Get Probing() As String
    Probing = mstrProbing
End Property

Public Property Let Probing(ByVal strValue As String)
    If Len(strValue) > 1000 Then
        goCWShared.ShowVBError 1000, "Probing string too long", "Notes+.clseExamEndoEndoTest", "String too long", vbInformation
    Else
        mstrProbing = strValue
        UpdateDirtyFlag
    End If
End Property
 Public Property Get Pain_On_Occulusion() As ExamEndoYesNoEnum
    Pain_On_Occulusion = mlngPainOcclusion
 End Property

Public Property Let Pain_On_Occulusion(ByVal lngValue As ExamEndoYesNoEnum)
    mlngPainOcclusion = lngValue
    UpdateDirtyFlag
End Property

Public Property Get Pain_On_Occulusion_Location() As String
    Pain_On_Occulusion_Location = mstrPainOcclusionLoc
End Property

Public Property Let Pain_On_Occulusion_Location(ByVal strString As String)
    If Len(strString) > 1000 Then
        goCWShared.ShowVBError 1000, "Pain_On_Occulusion_Location string too long", "Notes+.clseExamEndoEndoTest", "String too long", vbInformation
    Else
        mstrPainOcclusionLoc = strString
        UpdateDirtyFlag
    End If
End Property

Public Property Get Pain_On_Release() As ExamEndoYesNoEnum
    Pain_On_Release = mlngPainRelease
End Property

Public Property Let Pain_On_Release(ByVal lngValue As ExamEndoYesNoEnum)
    mlngPainRelease = lngValue
    UpdateDirtyFlag
End Property

Public Property Get Pain_On_Release_Location() As String
    Pain_On_Release_Location = mstrPainReleaseLoc
End Property


Public Property Let Pain_On_Release_Location(ByVal strValue As String)

    If Len(strValue) > 1000 Then
        goCWShared.ShowVBError 1000, "Pain_On_Release_Location string too long", "Notes+.clseExamEndoEndoTest", "String too long", vbInformation
    Else
        mstrPainReleaseLoc = strValue
        UpdateDirtyFlag
    End If
End Property

Public Property Get Caries() As ExamEndoYesNoEnum
    Caries = mlngCaries
End Property

Public Property Let Caries(ByVal lngValue As ExamEndoYesNoEnum)
    mlngCaries = lngValue
    UpdateDirtyFlag
End Property

'Read only
Public Property Get Unrestored_Caries() As Boolean
    Unrestored_Caries = mblnUnRestoredCaries
End Property

Public Property Get Radiograph_Findings() As ExamEndoYesNoEnum
   Radiograph_Findings = mlngRadiographFind
End Property

Public Property Let Radiograph_Findings(ByVal lngValue As ExamEndoYesNoEnum)
    mlngRadiographFind = lngValue
    UpdateDirtyFlag
End Property

Public Property Get Radiograph_Findings_Comments() As String
    Radiograph_Findings_Comments = mstrRadiographFindCmnt
End Property

Public Property Let Radiograph_Findings_Comments(ByVal strValue As String)
    If Len(strValue) > 2000 Then
        goCWShared.ShowVBError 1000, "Radiograph_Findings_Comments string too long", "Notes+.clseExamEndoEndoTest", "String too long", vbInformation
    Else
        mstrRadiographFindCmnt = strValue
        UpdateDirtyFlag
    End If
End Property

Public Property Get Trauma() As ExamEndoTraumaEnum
    Trauma = mlngTrauma
End Property

Public Property Let Trauma(ByVal lngValue As ExamEndoTraumaEnum)
    mlngTrauma = lngValue
    UpdateDirtyFlag
End Property

Public Property Get Anatomic_Variations() As String
    Anatomic_Variations = mstrAnatomicVar
End Property
Public Property Let Anatomic_Variations(ByVal strString As String)

    If Len(strString) > 2000 Then
        goCWShared.ShowVBError 1000, "Anatomic_Variations string too long", "Notes+.clseExamEndoEndoTest", "String too long", vbInformation
    Else
        mstrAnatomicVar = strString
        UpdateDirtyFlag
    End If
    
End Property

Public Property Get Restorative_Hx() As String
    Restorative_Hx = mstrRestoreHx
End Property

Public Property Let Restorative_Hx(ByVal strString As String)
    If Len(strString) > 2000 Then
        goCWShared.ShowVBError 1000, "Restorative_Hx string too long", "Notes+.clseExamEndoEndoTest", "String too long", vbInformation
    Else
        mstrRestoreHx = strString
        UpdateDirtyFlag
    End If
End Property

Public Property Get Hx_Bleaching() As ExamEndoYesNoEnum
    Hx_Bleaching = mlngBleachHx
End Property

Public Property Let Hx_Bleaching(ByVal lngValue As ExamEndoYesNoEnum)
    mlngBleachHx = lngValue
    UpdateDirtyFlag
End Property

Public Property Get Microleakage() As ExamEndoYesNoEnum
    Microleakage = mlngMicroleak
End Property

Public Property Let Microleakage(ByVal lngValue As ExamEndoYesNoEnum)
    mlngMicroleak = lngValue
    UpdateDirtyFlag
End Property

Public Property Get Hx_Comments() As String
    Hx_Comments = mstrHxComments
End Property

Public Property Let Hx_Comments(ByVal strString As String)
    If Len(strString) > 2000 Then
        goCWShared.ShowVBError 1000, "Hx_Comments string too long", "Notes+.clseExamEndoEndoTest", "String too long", vbInformation
    Else
        mstrHxComments = strString
        UpdateDirtyFlag
    End If
End Property

Public Property Get ColdTestWritten() As Boolean
    ColdTestWritten = mblnWroteColdTest
End Property
Public Property Let ColdTestWritten(ByVal blnValue As Boolean)
    mblnWroteColdTest = blnValue
End Property
Public Property Set DentalNotes(ByVal objObject As CHCSII_Dental_DXTX.clsNotes)
    Set mobjDentalNotes = objObject
End Property
Public Property Get DentalNotes() As CHCSII_Dental_DXTX.clsNotes
    Set DentalNotes = mobjDentalNotes
End Property
Public Property Get Heat_Response_Test() As ExamEndoHeatResponseTestEnum
    Heat_Response_Test = mlngHeatResponseTest
End Property
Public Property Let Heat_Response_Test(ByVal lngValue As ExamEndoHeatResponseTestEnum)
    mlngHeatResponseTest = lngValue
    UpdateDirtyFlag
End Property
Public Property Get HeatTestWritten() As Boolean
    HeatTestWritten = mblnHeatTestWritten
End Property
Public Property Let HeatTestWritten(ByVal blnValue As Boolean)
    mblnHeatTestWritten = blnValue
End Property

Public Sub ClearValues(ByVal intVal As Integer)
    Select Case intVal
        Case 0
            mstrPainOcclusionLoc = "MF (mesiofacial)|ML (mesiolingual)|DF (distofacial)|DL (distolingual)|" & _
                "F (facial)|L (lingual)|P (palatal)|"
        Case 1
            mstrSTComments = Empty
            mstrSTTracedTo = Empty
        Case 2
            'SCR 34288 Added spaces
            mstrPalpationLoc = "F (facial)|L (lingual)|MF (mesiofacial)|" & _
                            "ML (mesiolingual)|DF (distofacial)|DL (distolingual)|P (palatal)"
        Case 3
            mstrPainReleaseLoc = "MF (mesiofacial)|ML (mesiolingual)|DF (distofacial)|DL (distolingual)|" & _
                "F (facial)|L (lingual)|P (palatal)|"
    End Select
End Sub
'***********************************************************
'Sub: Initialize_Values
'Author: KDunne
'Date: 11/12/2001
'Purpose: Sets module level variables to their intial values.
'************************************************************
Private Sub Initialize_Values()
    mlngTestId = 0
    mlngDentalEncounterId = 0
    mstrToothNumber = Empty
    mlngEPT = ExamEndoEPTEnum.NoResponse
    mlngColdResponse = ExamEndoTestResponseEnum.Not_Addressed
    mlngColdLinger = ExamEndoYesNoEnum.Not_Addressed
    mlngEPT = ExamEndoEPTEnum.exEndoEPT_Not_Addressed
    mlngColdResponseTest = ExamEndoColdResponseTestEnum.Not_Addressed
    mlngHeatResponse = ExamEndoTestResponseEnum.Not_Addressed
    mlngHeatLinger = ExamEndoYesNoEnum.Not_Addressed
    mlngPercussion = ExamEndoPercussionEnum.Not_Addressed
    mlngPainPalpation = ExamEndoYesNoEnum.Not_Addressed
    
'SCR 43167 RWalker 10/21/2003 Initialize drop down
    ClearValues 2
    
    ClearValues 0
    mlngSinusTract = ExamEndoYesNoEnum.Not_Addressed
    ClearValues 1
    mlngSTTermination = ExamEndoSTTermination.Not_Addressed
    mstrMagnify = "normal|craze line|fracture|discoloration|caries|previously undetected canal|" & _
        "isthmus|obstruction of root canal|separated instrument|"
    mstrTransIllum = "normal|craze line|fracture|discoloration|caries|previously undetected canal|"
    mlngMobility = ExamEndoMobilityEnum.Not_Addressed
    mstrProbing = Empty
    mlngPainOcclusion = ExamEndoYesNoEnum.Not_Addressed
    mlngPainRelease = ExamEndoYesNoEnum.Not_Addressed
    ClearValues 3
    mlngCaries = ExamEndoYesNoEnum.Not_Addressed
    mlngRadiographFind = ExamEndoYesNoEnum.Not_Addressed
    mstrRadiographFindCmnt = Empty
    mlngTrauma = ExamEndoYesNoEnum.Not_Addressed
    mstrAnatomicVar = Empty
    mlngMicroleak = ExamEndoYesNoEnum.Not_Addressed
    mstrRestoreHx = Empty
    mlngBleachHx = ExamEndoYesNoEnum.Not_Addressed
    mstrHxComments = Empty
End Sub

'****************************************************
'Sub: UpdateDirtyFlag
'Author: KDunne
'Date: 11/5/2001
'Purpose: Updates module level dirty flag field
'****************************************************
Public Sub UpdateDirtyFlag()
    If mlngDirtyFlag = dfNoAction Then
        If mlngTestId = 0 Then
            mlngDirtyFlag = dfAdd
        Else
            mlngDirtyFlag = dfUpdate
        End If
    End If

End Sub

'*******************************************************************
'Function: SaveSection
'Author: Kdunne
'Date: 11/13/2001
'Returns: True/False
'Purpose: Saves the data to the table
'*******************************************************************
Public Function SaveSection() As Boolean
    Dim strSql As String
    Dim blnReturn As Boolean
    Dim objRS As ADODB.Recordset
    Dim strOnset As String
    
    On Error GoTo ErrHandler
    
    Select Case mlngDirtyFlag
        Case dfNoAction
            blnReturn = True
            GoTo FuncExit
        Case dfAdd
            'Get the primary key
            mlngTestId = goconn.GetNextID(esTEMPL_EXAM_ENDO_ENDONTIC_T_SEQ)
'            strSql = "SELECT templ_exam_endo_endontic_t_seq.NEXTVAL From Dual"
'            Set objRS = goconn.OpenRecordset(strSql)
'            mlngTestId = objRS.Fields("NEXTVAL")
            
            strSql = "INSERT INTO TEMPL_EXAM_ENDO_ENDONTIC_TEST " & _
                    " (EX_ENDO_TEST_ID, DENTAL_ENCOUNTER_ID, TOOTH_NUMBER, EPT, COLD_RESPONSE, " & _
                    " COLD_RESPONSE_TEST, HEAT_RESPONSE, COLD_LINGER, " & _
                    " HEAT_LINGER, PERCUSSION, PALPATION, " & _
                    " PALPATION_LOCATION, SINUS_TRACT, ST_TRACED_TO, " & _
                    " ST_COMMENTS, ST_TERMINATION, MAGNIFICATION, " & _
                    " TRANSILLUMINATION, MOBLITIY, PROBING, " & _
                    " PAIN_ON_OCCLUSION, PAIN_ON_OCCLUSION_LOC,  " & _
                    " PAIN_ON_RELEASE, PAIN_ON_RELEASE_LOC, CARIES,  " & _
                    " RADIOGRAPH_FINDINGS, " & _
                    " RADIOGRAPH_FINDINGS_COMMENTS, TRAUMA, " & _
                    " ANATOMIC_VARIATIONS, RESTORATIVE_HX, MICROLEAKAGE, " & _
                    " BLEACHING_HX , Hx_Comments, HEAT_TEST, Workload_ID) " & _
                    " VALUES (" & mlngTestId & ", " & mlngDentalEncounterId & ", '" & mstrToothNumber & _
                    "', " & mlngEPT & ", " & mlngColdResponse & ", " & mlngColdResponseTest & _
                    ", " & mlngHeatResponse & ", " & mlngColdLinger & ", " & _
                    mlngHeatLinger & ", " & mlngPercussion & ", " & mlngPainPalpation & ", '" & _
                    mstrPalpationLoc & "', " & mlngSinusTract & ", '" & mstrSTTracedTo & "', '" & _
                    Scrub_Apostrophe(mstrSTComments) & "', " & mlngSTTermination & ", '" & mstrMagnify & "', '" & _
                    mstrTransIllum & "', " & mlngMobility & ", '" & mstrProbing & "', " & _
                    mlngPainOcclusion & ", '" & mstrPainOcclusionLoc & "'," & mlngPainRelease & ", '" & _
                    mstrPainReleaseLoc & "', " & mlngCaries & ", " & mlngRadiographFind & ", '" & _
                    Scrub_Apostrophe(mstrRadiographFindCmnt) & "', " & mlngTrauma & ", '" & Scrub_Apostrophe(mstrAnatomicVar) & "', '" & _
                    Scrub_Apostrophe(mstrRestoreHx) & "', " & mlngMicroleak & ", " & mlngBleachHx & ", '" & _
                    Scrub_Apostrophe(mstrHxComments) & "', " & mlngHeatResponseTest & ", " & mlWorkloadID & ")"
                    
            If ValidateSave() = False Then
                strSql = ""
            End If
        Case dfUpdate
            strSql = "Update TEMPL_EXAM_ENDO_ENDONTIC_TEST SET " & _
                "TOOTH_NUMBER = '" & mstrToothNumber & "', EPT = " & mlngEPT & _
                ", COLD_RESPONSE = " & mlngColdResponse & ", COLD_RESPONSE_TEST = " & _
                mlngColdResponseTest & ", HEAT_RESPONSE = " & mlngHeatResponse & ", COLD_LINGER = " & _
                mlngColdLinger & ", HEAT_LINGER = " & mlngHeatLinger & ", PERCUSSION = " & _
                mlngPercussion & ", PALPATION = " & mlngPainPalpation & ", PALPATION_LOCATION = '" & _
                mstrPalpationLoc & "', SINUS_TRACT = " & mlngSinusTract & ", ST_TRACED_TO = '" & _
                mstrSTTracedTo & "', ST_COMMENTS = '" & Scrub_Apostrophe(mstrSTComments) & " ', ST_TERMINATION = " & _
                mlngSTTermination & ", MAGNIFICATION = '" & mstrMagnify & "', TRANSILLUMINATION = '" & _
                mstrTransIllum & "', MOBLITIY = " & mlngMobility & ", PROBING = '" & mstrProbing & _
                "', PAIN_ON_OCCLUSION = " & mlngPainOcclusion & ", PAIN_ON_OCCLUSION_LOC = '" & _
                mstrPainOcclusionLoc & "', PAIN_ON_RELEASE = " & mlngPainRelease & ",PAIN_ON_RELEASE_LOC = '" & _
                mstrPainReleaseLoc & "', CARIES = " & mlngCaries & ", RADIOGRAPH_FINDINGS = " & _
                mlngRadiographFind & ", RADIOGRAPH_FINDINGS_COMMENTS = '" & Scrub_Apostrophe(mstrRadiographFindCmnt) & _
                "', TRAUMA = " & mlngTrauma & ", ANATOMIC_VARIATIONS = '" & Scrub_Apostrophe(mstrAnatomicVar) & "', RESTORATIVE_HX = '" & _
                Scrub_Apostrophe(mstrRestoreHx) & "', MICROLEAKAGE = " & mlngMicroleak & ", BLEACHING_HX = " & _
                mlngBleachHx & ", HX_COMMENTS = '" & Scrub_Apostrophe(mstrHxComments) & "'," & _
                " HEAT_TEST = " & mlngHeatResponseTest & ", Workload_ID = " & mlWorkloadID & " Where EX_ENDO_TEST_ID = " & mlngTestId
            If ValidateSave() = False Then
                blnReturn = False
                GoTo FuncExit
                
            End If
        Case dfdelete
            strSql = "DELETE FROM TEMPL_EXAM_ENDO_ENDONTIC_TEST WHERE " & _
                     " EX_ENDO_TEST_ID = " & mlngTestId
    End Select
        
    If strSql <> "" Then
        goconn.ExecuteSQL (strSql)
        blnReturn = True
    End If
    mlngDirtyFlag = dfNoAction

    
FuncExit:
    Set objRS = Nothing
    SaveSection = blnReturn
    Exit Function
ErrHandler:
    SaveSection = False
    'goCWShared.ShowVBError Err.Number, Err.Description, Err.Source, App.Title, vbCritical


End Function

'*******************************************************************
'Function: GetNote
'Author: KDunne
'Date: 11/13/01
'Returns: A string that is displayed in the narrative summary for this encounter
'Purpose: To compare the values from this encounter and the previous encounter and
'return a text note.
'*******************************************************************
Public Function GetNote() As String
    Dim strReturn As String
    Dim strTemp As String
    Dim objProbing As CProbingLocations
    Dim objToothCtrl As CToothControl
        
    If mstrToothNumber = "" Then
        GetNote = Empty
        Exit Function
    Else
        Set objToothCtrl = New CToothControl
        strReturn = "Selected Tooth = " & objToothCtrl.Get_Tooth_Name(mstrToothNumber) & ";"
    End If
    
    'Read Mobility Info from Datalayer
    CheckForMobility CLng(mstrToothNumber)
    
    strTemp = Trim(ExamEndo_GetEPTEnumDesc(mlngEPT))
    
    If strTemp <> "" Then
        strReturn = strReturn & " EPT = " & strTemp & ";"
    End If
    
    If mblnWroteColdTest = False Then
        strTemp = Trim(ExamEndo_GetColdResponseTestEnumDesc(mlngColdResponseTest))
        
        If strTemp <> "" Then
            strReturn = strReturn & " Cold Response Test = " & strTemp & ";"
            mblnWroteColdTest = True
        End If
        
    End If
    strTemp = Trim(ExamEndo_GetTestResponseEnumDesc(mlngColdResponse))
    
    If strTemp <> "" Then
        strReturn = strReturn & " Cold Response = " & strTemp & ";"
    End If
    
    strTemp = Trim(ExamEndo_GetYesNoEnumDesc(mlngColdLinger))
    
    If strTemp <> "" Then
        strReturn = strReturn & " Cold Linger = " & strTemp & ";"
    End If
    
    If mblnHeatTestWritten = False Then
        strTemp = ExamEndo_GetHeatResponseTestEnumDesc(mlngHeatResponseTest)
        
        If strTemp <> "" Then
            strReturn = strReturn & " Heat Response Test = " & strTemp & ";"
            mblnHeatTestWritten = True
        End If
    
    End If
    strTemp = Trim(ExamEndo_GetTestResponseEnumDesc(mlngHeatResponse))
    
    If strTemp <> "" Then
        strReturn = strReturn & " Heat Response = " & strTemp & ";"
    End If
    
    strTemp = Trim(ExamEndo_GetYesNoEnumDesc(mlngHeatLinger))
    
    If strTemp <> "" Then
        strReturn = strReturn & " Heat Linger = " & strTemp & ";"
    End If
    
    strTemp = ExamEndo_GetPercussionEnumDesc(mlngPercussion)
    
    If strTemp <> "" Then
        strReturn = strReturn & " Percussion = " & strTemp & ";"
    End If
        
    strTemp = ExamEndo_GetYesNoEnumDesc(mlngPainPalpation)
    
    If strTemp <> "" Then
        strReturn = strReturn & " Pain on Palpation = " & strTemp & ";"
    End If
   
    strTemp = Get_MultiSelect_DisplayString(mstrPalpationLoc)
    If strTemp <> "" Then
         strReturn = strReturn & " Palpation Location(s) = " & strTemp & ";"
    End If
    
    strTemp = ExamEndo_GetYesNoEnumDesc(mlngSinusTract)
    
    If strTemp <> "" Then
        strReturn = strReturn & " Sinus Tract Present = " & strTemp & ";"
    End If
    
    If Trim(mstrSTTracedTo) <> "" Then
        'SCR 47455 Added objToothCtrl.Get_Tooth_Name(....)
        strReturn = strReturn & " Sinus Tract Traced to " & mstrSTTracedTo & ";"
    End If
    
    strTemp = ExamEndo_GetSTTerminationEnumDesc(mlngSTTermination)
    If strTemp <> "" Then
        strReturn = strReturn & " Sinus Tract Termination = " & strTemp & ";"
    End If
    
    If Trim(mstrSTComments) <> "" Then
        strReturn = strReturn & " Sinus Tract Comments: " & mstrSTComments & ";"
    End If
    
    strTemp = Get_MultiSelect_DisplayString(mstrMagnify)
    If strTemp <> "" Then
        strReturn = strReturn & " Magnification Findings  = " & strTemp & ";"
    End If
    
    strTemp = Get_MultiSelect_DisplayString(mstrTransIllum)
    If strTemp <> "" Then
        strReturn = strReturn & "Transillumination = " & strTemp & ";"
    End If
    
    strTemp = ExamEndo_GetMobilityEnumDesc(mlngMobility)
    
    If strTemp <> "" Then
        strReturn = strReturn & " Mobility = " & strTemp & ";"
    End If
    
    If Trim(mstrProbing) <> "" Then
        Set objProbing = New CProbingLocations
        
        strReturn = strReturn & " Probing Documentation = " & objProbing.Build_Display_String(mstrProbing) & ";"
    End If
    
    strTemp = Trim(ExamEndo_GetYesNoEnumDesc(mlngPainOcclusion))
    
    If strTemp <> "" Then
        strReturn = strReturn & " Pain on Occlusion = " & strTemp & ";"
    End If
    
    strTemp = Trim(Get_MultiSelect_DisplayString(mstrPainOcclusionLoc))
    If strTemp <> "" Then
        strReturn = strReturn & " Pain on Occlusion Location = " & strTemp & ";"
    End If
    
    strTemp = ExamEndo_GetYesNoEnumDesc(mlngPainRelease)
    
    If strTemp <> "" Then
        strReturn = strReturn & " Pain on Release = " & strTemp & ";"
    End If
    
    strTemp = Get_MultiSelect_DisplayString(mstrPainReleaseLoc)
    If strTemp <> "" Then
        strReturn = strReturn & " Pain on Release Location(s) = " & strTemp & ";"
    End If
    
    strTemp = ExamEndo_GetYesNoEnumDesc(mlngCaries)
    
    If strTemp <> "" Then
        strReturn = strReturn & " Caries = " & strTemp & ";"
    End If
    
    strTemp = ExamEndo_GetYesNoEnumDesc(mlngRadiographFind)
    
    If strTemp <> "" Then
        strReturn = strReturn & " Radiograph Findings = " & strTemp & ";"
    End If
    
    If Trim(mstrRadiographFindCmnt) <> "" Then
        strReturn = strReturn & " Radiograph Findings Comments: " & mstrRadiographFindCmnt & ";"
    End If
    
    strTemp = ExamEndo_GetTraumaEnumDesc(mlngTrauma)
    
    If strTemp <> "" Then
        strReturn = strReturn & " Trauma = " & strTemp & ";"
    End If
    
    If Trim(mstrAnatomicVar) <> "" Then
        strReturn = strReturn & " Anatomic Variations = " & mstrAnatomicVar & ";"
    End If
    
    If Trim(mstrRestoreHx) <> "" Then
        strReturn = strReturn & " Restorative Hx = " & mstrRestoreHx & ";"
    End If
    
    strTemp = ExamEndo_GetYesNoEnumDesc(mlngMicroleak)
    
    If strTemp <> "" Then
        strReturn = strReturn & " Microleakage = " & strTemp & ";"
    End If
    
    strTemp = ExamEndo_GetYesNoEnumDesc(mlngBleachHx)
    
    If strTemp <> "" Then
        strReturn = strReturn & " Hx of Bleaching = " & strTemp & ";"
    End If
    
    If Trim(mstrHxComments) <> "" Then
        strReturn = strReturn & " Comments = " & mstrHxComments & ";"
    End If
    
    GetNote = strReturn
End Function



'******************************************************************************
'Function: ValidateSave
'Author: KDunne
'Returns: boolean. True if any of the validation criteria are met.
'Purpose: Validates the test for a single tooth. Tooth number plus one other piece
'         of info must be selected in order to save to the db.
'******************************************************************************
Public Function ValidateSave() As Boolean
    Dim blnSave As Boolean
    
    On Error GoTo ErrHandler
    blnSave = False
    
    If mstrToothNumber = "" Then
        ValidateSave = False
        Exit Function
    End If
        
    If mlngEPT <> ExamEndoEPTEnum.NoResponse Then
        blnSave = True
    End If
    If mlngColdResponse <> ExamEndoTestResponseEnum.Not_Addressed Then
        blnSave = True
    End If
    If mlngColdLinger <> ExamEndoYesNoEnum.Not_Addressed Then
        blnSave = True
    End If
    
    If mlngColdResponseTest <> ExamEndoColdResponseTestEnum.Not_Addressed Then
        blnSave = True
    End If
    
    If mlngHeatResponse <> ExamEndoTestResponseEnum.Not_Addressed Then
        blnSave = True
    End If
    
    If mlngHeatLinger <> ExamEndoYesNoEnum.Not_Addressed Then
        blnSave = True
    End If
    If mlngPercussion <> ExamEndoPercussionEnum.Not_Addressed Then
        blnSave = True
    End If
    If mlngPainPalpation <> ExamEndoYesNoEnum.Not_Addressed Then
        blnSave = True
    End If
    
    If Get_MultiSelect_DisplayString(mstrPalpationLoc) <> "" Then
        blnSave = True
    End If
    
    If mlngSinusTract <> ExamEndoYesNoEnum.Not_Addressed Then
        blnSave = True
    End If
    
    If Get_MultiSelect_DisplayString(mstrSTTracedTo) <> "" Then
        blnSave = True
    End If
    
    If mstrSTComments <> "" Then
        blnSave = True
    End If
    
    If mlngSTTermination <> ExamEndoSTTermination.Not_Addressed Then
        blnSave = True
    End If
    
    If Get_MultiSelect_DisplayString(mstrMagnify) <> "" Then
        blnSave = True
    End If
    
    If Get_MultiSelect_DisplayString(mstrTransIllum) <> "" Then
        blnSave = True
    End If
    
    If mlngMobility <> ExamEndoMobilityEnum.Not_Addressed Then
        blnSave = True
    End If
        
    If mlngPainOcclusion <> ExamEndoYesNoEnum.Not_Addressed Then
        blnSave = True
    End If
    
    If Get_MultiSelect_DisplayString(mstrPainOcclusionLoc) <> "" Then
        blnSave = True
    End If
    
    If mlngPainRelease <> ExamEndoYesNoEnum.Not_Addressed Then
        blnSave = True
    End If
    
    If Get_MultiSelect_DisplayString(mstrPainReleaseLoc) <> "" Then
        blnSave = True
    End If
    
    If mlngCaries <> ExamEndoYesNoEnum.Not_Addressed Then
        blnSave = True
    End If
    
    If mlngRadiographFind <> ExamEndoYesNoEnum.Not_Addressed Then
        blnSave = True
    End If
    
    If mstrRadiographFindCmnt <> "" Then
        blnSave = True
    End If
    
    If mlngTrauma <> ExamEndoTraumaEnum.Not_Addressed Then
        blnSave = True
    End If
    
    If Get_MultiSelect_DisplayString(mstrAnatomicVar) <> "" Then
        blnSave = True
    End If
    
    If mstrRestoreHx <> "" Then
        blnSave = True
    End If
    
    If mlngMicroleak <> ExamEndoYesNoEnum.Not_Addressed Then
        blnSave = True
    End If
    If mlngBleachHx <> ExamEndoYesNoEnum.Not_Addressed Then
        blnSave = True
    End If
    
    If mstrHxComments <> "" Then
        blnSave = True
    End If
    
    If mlngColdResponse <> ExamEndoTestResponseEnum.Not_Addressed Then
        If mlngColdResponseTest = ExamEndoColdResponseTestEnum.Not_Addressed Then
            blnSave = False
'            MsgBxARMd "Cold Response Test must be selected"
        End If
    End If
    
    If mlngColdResponseTest <> ExamEndoColdResponseTestEnum.Not_Addressed Then
        If mlngColdResponse = ExamEndoTestResponseEnum.Not_Addressed Then
            blnSave = False
'            MsgBxARMd "Cold Response must be selected"
        End If
    End If
    
    If mlngHeatResponse <> ExamEndoTestResponseEnum.Not_Addressed Then
        If mlngHeatResponseTest = ExamEndoHeatResponseTestEnum.Not_Addressed Then
            blnSave = False
'            MsgBxARMd "Heat Response Test must be selected"
        End If
    End If
    
    If mlngHeatResponseTest <> ExamEndoHeatResponseTestEnum.Not_Addressed Then
        If mlngHeatResponse = ExamEndoTestResponseEnum.Not_Addressed Then
            blnSave = False
'            MsgBxARMd "Heat Response must be selected"
        End If
    End If
    ValidateSave = blnSave
    
    Exit Function
ErrHandler:
    ValidateSave = False
    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source, App.Title, vbCritical

End Function

'*********************************************************************************
'Function: DB_LOAD_UNRESTORED_CARIES
'Author: KDunne
'Date: 11/20/2001
'Purpose: Loads Unrestored caries diagnoses for a given tooth.
'*********************************************************************************
Private Function DB_LOAD_UNRESTORED_CARIES(ByVal lngToothid As Long, _
    ByVal lngDentalEncounterId As Long) As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo ErrHandler
    
'SCR 42724 RWalker 10/20/2003
'Also check for new code cICDCodeK02_0 ("Caries, with minimal extension into dentin") 1418
    strSql = "SELECT DXS.DX_CODE_ID,DX.STATUS_ID " & _
        " FROM DENTAL_ENCOUNTER_DX DX, " & _
        " DENTAL_ENCOUNTER_DXS_SS DXS, " & _
        " DENTAL_ENC_SURFACE_SET DSS, " & _
        " DENTAL_ENCOUNTER_TEETH DT " & _
        " WHERE (DT.TOOTH_ID = " & lngToothid & ") AND " & _
        " (DT.DENTAL_ENCOUNTER_ID = " & lngDentalEncounterId & ") AND " & _
        " DSS.DENTAL_ENCOUNTER_TEETH_ID = DT.DENTAL_ENCOUNTER_TEETH_ID " & _
        " AND DXS.SURFACE_SET_ID = DSS.SURFACE_SET_ID AND " & _
        " DX.DX_CODE_ID = DXS.DX_CODE_ID AND " & _
        " DX.STATUS_ID IN (" & cDxStatusNew & ", " & cDxStatusOnGoing & ") AND " & _
        " DXS.DX_CODE_ID IN (" & cICDCodeK02_2 & ", " & cICDCodeK02_12 & ", " & _
        cICDCodeK02_13 & ", " & cICDCodeK02_14 & ", " & cICDCodeK02_3 & ", " & _
        cICDCodeK02_4 & ", " & cICDCodeK02_5 & ", " & cICDCodeK02_51 & ", " & _
        cICDCodeK02_0 & ")"
   
    Set DB_LOAD_UNRESTORED_CARIES = goconn.OpenRecordset(strSql)
    
    Exit Function
ErrHandler:

    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & "." & "DB_LOAD_UNRESTORED_CARIES", App.Title, vbCritical
        
    Set DB_LOAD_UNRESTORED_CARIES = Nothing
End Function 'DB_LOAD_UNRESTORED_CARIES

'**************************************************************
'Function: CheckForUnRestoredCaries
'Author: KDunne
'Date:11/20/2001
'Purpose: Checks a given tooth for any unrestored caries.
'Notes: Called in the property let statement for Tooth_Number
'**************************************************************
Private Sub CheckForUnRestoredCaries(ByVal lngToothid As Long)
'    Dim objRS As ADODB.Recordset
'
'
'    Set objRS = DB_LOAD_UNRESTORED_CARIES(lngToothid, mlngDentalEncounterId)
'
'    If objRS Is Nothing Then
'        'We encountered an error trying to read from the db.
'        mblnUnRestoredCaries = False
'    Else
'        If objRS.EOF Then
'            'No unrestored caries for this tooth.
'            mblnUnRestoredCaries = False
'        Else
'            mblnUnRestoredCaries = True
'        End If
'    End If
'
'    Set objRS = Nothing
    Dim objGroups As CHCSII_Dental_DXTX.colGroups
    Dim objGroup As CHCSII_Dental_DXTX.clsGroup
    Dim objDxs As CHCSII_Dental_DXTX.colDxs
    Dim objDx As CHCSII_Dental_DXTX.clsDx
    Dim objTeeth As CHCSII_Dental_DXTX.colTeeth
    Dim objTooth As CHCSII_Dental_DXTX.clsTooth
    Dim objSSs As CHCSII_Dental_DXTX.colSurfaceSets
    Dim objSS As CHCSII_Dental_DXTX.clsSurfaceSet
    Dim intIndex As Integer
    Dim intCount As Integer
    
    On Error GoTo ErrHandler:
    
    If mobjDentalNotes Is Nothing Then
        Exit Sub
    End If
    mblnUnRestoredCaries = False
    
    Set objGroups = mobjDentalNotes.Groups
    
    If objGroups Is Nothing Then
        Exit Sub
    End If
    
    For Each objGroup In objGroups
        'In each Group, we have a collection of DX and TX's.
        Set objDxs = objGroup.Dxs
        
        'Loop through the Diagnosis'
        For Each objDx In objDxs
            'SCR 42724 RWalker 10/20/2003
            'Also check for codes:
            'cICDCodeK02_0 ("Caries, with minimal extension into dentin") 1418
            'cICDCodeK02_13 ("Caries, moderate to advanced into dentin ") 1420
            If (objDx.DxICDCodeID = cICDCodeK02_2 Or objDx.DxICDCodeID = cICDCodeK02_12 Or _
                objDx.DxICDCodeID = cICDCodeK02_14 Or objDx.DxICDCodeID = cICDCodeK02_3 Or _
                objDx.DxICDCodeID = cICDCodeK02_4 Or objDx.DxICDCodeID = cICDCodeK02_5 Or _
                objDx.DxICDCodeID = cICDCodeK02_51 Or _
                objDx.DxICDCodeID = cICDCodeK02_0 Or objDx.DxICDCodeID = cICDCodeK02_13) And _
                (objDx.StatusID = dxActiveNew Or objDx.StatusID = dxActiveOngoing) Then
                'Diagnosis is for mobility, and it was not deleted.
                Set objSSs = objDx.SurfaceSets
                
                'Each Diagnosis contains a collection of surface sets.
                'Loop through each Surface Set in the Surface Set collection
                For Each objSS In objSSs
                    'Each Surface set can contain multiple teeth. Currently each
                    'Surface set only contains a single tooth, but could potentially
                    'contain multiple teeth.
                    Set objTeeth = objSS.Teeth
                    For Each objTooth In objTeeth
                        If objTooth.ToothID = lngToothid Then
                            mblnUnRestoredCaries = True
                          
                        End If
                    Next
                Next
                
            End If
        Next
    Next

    Exit Sub
ErrHandler:
    goCWShared.ShowVBError Err.Number, Err.Description, "clsExamEndoTest.CheckForUnRestoredCaries()", "", vbCritical
    
End Sub
'**************************************************************
'Function: CheckForMobility
'Author: KDunne
'Date:11/20/2001
'Purpose: Checks a given tooth for Mobility Info.
'Notes: Called in the property let statement for Tooth_Number
'**************************************************************

Private Sub CheckForMobility(ByVal lngToothid As Long)
    Dim objGroups As CHCSII_Dental_DXTX.colGroups
    Dim objGroup As CHCSII_Dental_DXTX.clsGroup
    Dim objDxs As CHCSII_Dental_DXTX.colDxs
    Dim objDx As CHCSII_Dental_DXTX.clsDx
    Dim objTeeth As CHCSII_Dental_DXTX.colTeeth
    Dim objTooth As CHCSII_Dental_DXTX.clsTooth
    Dim objSSs As CHCSII_Dental_DXTX.colSurfaceSets
    Dim objSS As CHCSII_Dental_DXTX.clsSurfaceSet
    Dim intIndex As Integer
    Dim intCount As Integer
    
    On Error GoTo ErrHandler:
    
    If mlngMobility <> ExamEndoMobilityEnum.WNL Then
        'Clear out any previous mobilty setting
        mlngMobility = ExamEndoMobilityEnum.Not_Addressed
    End If
    
    If mobjDentalNotes Is Nothing Then
        Exit Sub
    End If
    Set objGroups = mobjDentalNotes.Groups
    
    If objGroups Is Nothing Then
        Exit Sub
    End If
    
    'Clear out the previous values
    'SF 9-16-03
    Set mobjDx = Nothing
    Set mobjGroup = Nothing

    For Each objGroup In objGroups
        'In each Group, we have a collection of DX and TX's.
        Set objDxs = objGroup.Dxs
        
        'Loop through the Diagnosis'
        For Each objDx In objDxs
            If (objDx.DxICDCodeID = DxTxMobilityICDEnums.Ankylosis Or _
                objDx.DxICDCodeID = DxTxMobilityICDEnums.MillerMobility1 Or objDx.DxICDCodeID = _
                DxTxMobilityICDEnums.MillerMobility2 Or objDx.DxICDCodeID = DxTxMobilityICDEnums.MillerMobility3) _
                And objDx.StatusID <> dxNonActiveDeleted Then
                'Diagnosis is for mobility, and it was not deleted.
                Set objSSs = objDx.SurfaceSets
                
                'Each Diagnosis contains a collection of surface sets.
                'Loop through each Surface Set in the Surface Set collection
                For Each objSS In objSSs
                    'Each Surface set can contain multiple teeth. Currently each
                    'Surface set only contains a single tooth, but could potentially
                    'contain multiple teeth.
                    Set objTeeth = objSS.Teeth
                    For Each objTooth In objTeeth
                        If objTooth.ToothID = lngToothid Then
                            Select Case objDx.DxICDCodeID
                                Case DxTxMobilityICDEnums.MillerMobility1
                                    mlngMobility = ExamEndoMobilityEnum.Mobility1
                                Case DxTxMobilityICDEnums.MillerMobility2
                                    mlngMobility = ExamEndoMobilityEnum.Mobility2
                                Case DxTxMobilityICDEnums.MillerMobility3
                                    mlngMobility = ExamEndoMobilityEnum.Mobility3
                                Case DxTxMobilityICDEnums.Ankylosis
                                    mlngMobility = ExamEndoMobilityEnum.ANK
                                Case Else
                                    If mlngMobility <> ExamEndoMobilityEnum.WNL Then
                                        mlngMobility = ExamEndoMobilityEnum.Not_Addressed
                                    End If
                            End Select
                            
                            Set mobjDx = objDx
                            Set mobjGroup = objGroup
'                            Exit Sub
                        
                        End If
                    Next
                Next
                
            End If
        Next
    Next

    Exit Sub
ErrHandler:
    goCWShared.ShowVBError Err.Number, Err.Description, "clsExamEndoTest.CheckForMoblity()", "", vbCritical
End Sub
'*********************************************************************************
'Function: DB_LOAD_TOOTH_MOBILITY_INFO
'Author: KDunne
'Date: 11/20/2001
'Purpose: Loads mobility Info for a given tooth.
'*********************************************************************************
Private Function DB_LOAD_TOOTH_MOBILITY_INFO(ByVal lngToothid As Long, _
    ByVal lngDentalEncId As Long) As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo ErrHandler
    
    strSql = "SELECT DXS.DX_CODE_ID,DX.STATUS_ID " & _
        " FROM DENTAL_ENCOUNTER_DX DX, " & _
        " DENTAL_ENCOUNTER_DXS_SS DXS, " & _
        " DENTAL_ENC_SURFACE_SET DSS, " & _
        " DENTAL_ENCOUNTER_TEETH DT " & _
        " WHERE (DT.TOOTH_ID = " & lngToothid & ") AND " & _
        " (DT.DENTAL_ENCOUNTER_ID = " & lngDentalEncId & ") AND " & _
        " DSS.DENTAL_ENCOUNTER_TEETH_ID = DT.DENTAL_ENCOUNTER_TEETH_ID " & _
        " AND DXS.SURFACE_SET_ID = DSS.SURFACE_SET_ID AND " & _
        " DX.DX_CODE_ID = DXS.DX_CODE_ID AND " & _
        " DXS.DX_CODE_ID IN (" & DxTxMobilityICDEnums.MillerMobility1 & ", " & _
        DxTxMobilityICDEnums.MillerMobility2 & ", " & DxTxMobilityICDEnums.MillerMobility3 & _
        ", " & DxTxMobilityICDEnums.Ankylosis & ")"
    
    Set DB_LOAD_TOOTH_MOBILITY_INFO = goconn.OpenRecordset(strSql)
    
    Exit Function
ErrHandler:

    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & "." & "DB_LOAD_UNRESTORED_CARIES", App.Title, vbCritical
        
    Set DB_LOAD_TOOTH_MOBILITY_INFO = Nothing
End Function 'DB_LOAD_TOOTH_MOBILITY_INFO

Private Sub Class_Terminate()
    Set mobjDentalNotes = Nothing
    Set mobjDx = Nothing
    Set mobjSvgComm = Nothing
    Set mobjSvgFac = Nothing
End Sub
'**************************************************************************
'Function: GetMobilityDxCode
'Author: Kdunne
'Date:2/28/2002
'Purpose: Given a user mobility selection, function returns the appropriate
'          Dx Code Id.
'****************************************************************************
Private Function GetMobilityDxCode(ByVal lngMobility As ExamEndoMobilityEnum) As Long
    Dim lngReturn As DxTxMobilityICDEnums

    Select Case lngMobility
        Case ExamEndoMobilityEnum.ANK
            lngReturn = DxTxMobilityICDEnums.Ankylosis
        Case ExamEndoMobilityEnum.Mobility1
            lngReturn = DxTxMobilityICDEnums.MillerMobility1
        Case ExamEndoMobilityEnum.Mobility2
            lngReturn = DxTxMobilityICDEnums.MillerMobility2
        Case ExamEndoMobilityEnum.Mobility3
            lngReturn = DxTxMobilityICDEnums.MillerMobility3
        Case ExamEndoMobilityEnum.WNL
            'WNL entries are not populated in the Dx/Tx grid
            lngReturn = DxTxMobilityICDEnums.Not_Addressed
        Case Else
            lngReturn = DxTxMobilityICDEnums.Not_Addressed
    
    End Select
    
    GetMobilityDxCode = lngReturn
End Function


'************************************************************************
'Function: UpdateSVGGraphic
'Author: KDunne
'Date: 3/8/2002
'Purpose:UPdates the SVG Graphic in Big Tooth
'************************************************************************
Private Function UpdateSVGGraphic(objSVGComm As srvSVGComm.clsSVGComm, objDx As CHCSII_Dental_DXTX.clsDx, _
                    objTx As CHCSII_Dental_DXTX.clsTx, objSS As CHCSII_Dental_DXTX.clsSurfaceSet, _
                    ByVal intToothId As Integer, ByVal strSurfaceSet As String, _
                    Optional ByVal blnRemoveItem As Boolean = False) As Boolean
    Dim blnReturn As Boolean
    
    On Error GoTo ErrHandler
    
    DoEvents
    'Refresh the grid after I draw the tooth
    objSVGComm.TeethObject.ClickTeeth intToothId, True
    objSVGComm.ToothObject.ToothChanged intToothId, False
    objSVGComm.ToothObject.selectedSurfaces = strSurfaceSet
    
    If blnRemoveItem Then
        objSVGComm.updateSVG objDx, objTx
    Else
        objSVGComm.drawMenuItem intToothId, strSurfaceSet, objSS.SymbolID, objDx, objTx
        
        objSVGComm.UpdateTooth intToothId, 1, False
        objSVGComm.UpdateTooth intToothId, 2, False
        objSVGComm.UpdateTooth intToothId, 0, False
    End If
    
    objSVGComm.TeethObject.setTeethPath mlngPatientID
    
    blnReturn = True
    
    
FuncExit:
    DoEvents
    UpdateSVGGraphic = blnReturn
    Exit Function
    
ErrHandler:
    blnReturn = False
    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source, App.Title, vbCritical
    GoTo FuncExit
End Function


