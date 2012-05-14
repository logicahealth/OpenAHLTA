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
Attribute VB_Name = "CToothControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IToothCtrlParent

Private mobjTeeth As CTeeth
Private mstrFormCaption As String
Private mstrValueString As String
Private mblnCancelled As Boolean
Private mobjDAS As ICHCSII_DAS
Private marrValues() As String
Private marrMissingTeeth() As String
Private mintMissingToothCnt As Integer
Private mobjDentalNOtes As CHCSII_Dental_DXTX.clsNotes
Private mlngControlType As ToothControlType
Private mlngState As ToothControlState
Private mstrDataString As String
Private mstrDisplayString As String
Private mlngToothCanal As ToothControlCanalsEnum
Private mblnMissingTeethSelectable As Boolean
Private mblnSingleSelect As Boolean
Private marrRegions() As String
Private mobjSVGTeeth As srvSVGComm.clsTeeth
Private mobjSVGMgr As CHCSII_Dental_SVG_Manager.clsManage
Private mlngPatientAge As Long
Private mstrImplantPath As String
Private mstrPonticPath As String

Private Const const_MODULE_NAME = "CToothControl"

Public Enum ToothControlType
    Standard
    Radiographic
    Region
    CanalMap
End Enum

Public Enum ToothControlState
    Radiopacity = 0
    Radiolucency = 1
    Calculus = 2
    BiologicalWidth = 3
    CrRootRatio = 4
    DefectiveRest = 5
    RootProximity = 6
    SinusProximity = 7
    WidenedPDL = 8
    LaminaDura = 9
    FurcaInvasion = 10
    HorizBoneLoss = 11
    VertBoneLoss = 12
    OsseousDefects = 13
    Generalized = 14
    Localized = 15
    PeriapicalLesion = 16
    Overhangs = 17
    Caries = 18
End Enum

Public Enum RegionCodeEnum
    UpperRightSextant = 100
    UpperAnteriorSextant = 101
    UpperLeftSextant = 102
    LowerLeftSextant = 103
    LowerAnteriorSextant = 104
    LowerRightSextant = 105
    RightTMJ = 106
    LeftTMJ = 107
    Alveolar_mucosa = 108
    Alveolar_ridge = 109
    Buccal_mucosa = 110
    Commissure = 111
    Floor_of_mouth = 112
    Frenum = 113
    Gingiva = 114
    Hamular_notch = 115
    Hard_palate = 116
    Lower_lip = 117
    Mandible = 118
    Maxilla = 119
    Nasopharynx = 120
    Neck = 121
    Oral_Cavity = 122
    palate = 123
    Parotid_gland = 124
    Premaxilla = 125
    Retromolar_area = 126
    Sinus = 127
    Skin = 128
    Soft_palate = 129
    Sublingual_gland = 130
    Submandibular_gland = 131
    Tempromandibular_joint = 132
    Tongue = 133
    Tonsillar_area = 134
    Tuberosity = 135
    Upper_lip = 136
    Uvula = 137
    Vestibule = 138

End Enum

Public Enum ToothControlCanalsEnum
    None
    Mesial
    Mesiobuccal
    Mesiobuccal2
    Mesiobuccal3
    Mesiolingual
    Distal
    DistalBuccal
    Distallingual
    Facial
    Palatal
    Palatal2
    Lingual
    NOS
End Enum


Public Function ExternalInit(CWShared As CWShared, SVGMgr As CHCSII_Dental_SVG_Manager.clsManage)

'This is used to manually init the the object when it is used outside the process of chcsii

    Set mobjSVGMgr = SVGMgr
    'Set mobjSVGTeeth = SVGComm.TeethObject
    Set gobjCWShared = CWShared
    
    Class_Initialize
    
End Function


Private Sub Class_Initialize()
'    Dim blnTheater As Boolean
    Dim oSVGComm As srvSVGComm.clsSVGComm
    
    Set mobjTeeth = New CTeeth
    
'    blnTheater = gobjCWShared.UseLocalDB
    
    Set goChcsiiConn = New CHCSII_CONN.Conn
    Set mobjDAS = goChcsiiConn.CHCSII_DAS
'    If blnTheater Then
'        'Theater Mode
'        Set mobjDAS = New GEMS_DAS
'    Else
'        Set mobjDAS = New CHCSII_DAS
'    End If
       
    'Intialize SVG objects
    'mobjSVGTeeh IS NOT USED?!?!?!?!?
    If mobjSVGTeeth Is Nothing Then
        Set oSVGComm = gobjSingleton.Singleton
        If oSVGComm Is Nothing = False Then
            Set mobjSVGTeeth = oSVGComm.TeethObject
        Else
            Set mobjSVGTeeth = New srvSVGComm.clsTeeth
        End If
    End If
    
    If mobjSVGMgr Is Nothing Then
        Set mobjSVGMgr = gobjSVGMgrSF.Singleton
    End If
    
    mblnMissingTeethSelectable = True
    mblnSingleSelect = False
    
    Set oSVGComm = Nothing
End Sub

Private Sub Class_Terminate()
    Set mobjTeeth = Nothing
    Set mobjDAS = Nothing
    Set mobjDentalNOtes = Nothing
    Set mobjSVGTeeth = Nothing
    Set mobjSVGMgr = Nothing
End Sub

Public Property Let ControlType(ByVal lngType As ToothControlType)
    mlngControlType = lngType
End Property
Public Property Get ControlType() As ToothControlType
    ControlType = mlngControlType
End Property
Public Property Let FormCaption(ByVal strCaption As String)
    mstrFormCaption = strCaption
End Property
Public Property Get FormCaption() As String
    FormCaption = mstrFormCaption
End Property
Public Property Set DentalNotes(ByVal objNotes As CHCSII_Dental_DXTX.clsNotes)
    Set mobjDentalNOtes = objNotes
End Property

Public Property Get DentalNotes() As CHCSII_Dental_DXTX.clsNotes
    Set DentalNotes = mobjDentalNOtes
End Property

Public Property Get Teeth() As CTeeth
    Set Teeth = mobjTeeth
End Property

Friend Property Set Teeth(ByVal objTeeth As CTeeth)
    Set mobjTeeth = objTeeth
End Property

Public Property Get ValueString() As String
    ValueString = mstrValueString
End Property
Public Property Get DisplayString() As String
    DisplayString = mstrDisplayString
End Property
Public Property Get MissingTeethSelectable() As Boolean
    MissingTeethSelectable = mblnMissingTeethSelectable
End Property
Public Property Let MissingTeethSelectable(ByVal blnValue As Boolean)
    mblnMissingTeethSelectable = blnValue
End Property
Public Property Get SingleSelect() As Boolean
    SingleSelect = mblnSingleSelect
End Property
Public Property Let SingleSelect(ByVal blnValue As Boolean)
    mblnSingleSelect = blnValue
End Property
Public Property Get ToothCanal() As ToothControlCanalsEnum
    ToothCanal = mlngToothCanal
End Property
Public Property Let ToothCanal(ByVal lngCanal As ToothControlCanalsEnum)
    mlngToothCanal = lngCanal
    
End Property
Public Property Get ToothCanalDesc() As String
    ToothCanalDesc = Get_Canal_Description(mlngToothCanal)
End Property
'*******************************************************************
'Sub: InitializeString
'Author: KDunne
'Date: 10/24/2001
'Purpose: Converts Value string passed in from Load_Teeth_Info.. func
'         to a module level array
'*******************************************************************
Private Sub InitializeString(ByVal strValues As String)
    marrValues = Split(strValues, const_ITEM_DELIMETER)
End Sub

'*******************************************************************
'Function: CheckToothPressed
'Author: KDunne
'Date: 10/24/2001
'Purpose: Checks to see if a given button was defaulted to pressed.
'*******************************************************************
Private Function CheckToothPressed(ByVal lngToothId As Long) As Boolean
    Dim arrValues() As String
    Dim intIndex As Integer
    Dim blnReturn As Boolean
    Dim intCount As Integer
    

    blnReturn = False
    
    For intIndex = 0 To UBound(marrValues)
        arrValues = Split(marrValues(intIndex), ",")
        If Get_Tooth_Number(arrValues(0)) = lngToothId Then
        'We've found the tooth in the value string. Now check to see if its
        'pressed or not. There are two formats for the value string, one is
        'just a list of tooth id's that have been pressed. The other is a
        'list of all the teeth in the mouth, and whether or not they were
        'pressed.
            If UBound(arrValues) = 0 Then
                'value string is the first format. Its presence in the string
                'means that it was pressed.
                blnReturn = True
            Else
                blnReturn = CBool(arrValues(1))
            End If
        End If
    Next
    
    CheckToothPressed = blnReturn
    
End Function
'*******************************************************************
'Function: CheckToothMissing
'Author: KDunne
'Date: 10/24/2001
'Purpose: Checks to see if a given tooth is missing. Checks the Data
'         Layer object.
'Notes: The DentalNotes object property of Must be set prior to checking
'*******************************************************************
Private Function CheckToothMissing(ByVal lngToothId As Long) As Boolean
    Dim blnReturn As Boolean
    Dim intIndex As Integer
    Dim intCount As Integer
    
    blnReturn = False
    
    
    If mintMissingToothCnt = 0 Then
        'No missing teeth.
        blnReturn = False
    Else
        For intIndex = 0 To mintMissingToothCnt
            If CLng(marrMissingTeeth(intIndex)) = lngToothId Then
                blnReturn = True
                Exit For
            End If
        Next
    End If
    CheckToothMissing = blnReturn
End Function
'*******************************************************************
'Sub: LoadMissingTeeth
'Author: KDunne
'Date: 10/24/2001
'Purpose: Loops through the Data Layer object and builds a list of
'         Missing teeth, and converts them into a module level array
'         that can be checked by CheckToothMissing() function
'Notes: The DentalNotes object property of Must be set prior to checking
'*******************************************************************

Private Sub LoadMissingTeeth()
    Dim objGroups As CHCSII_Dental_DXTX.colGroups
    Dim objGroup As CHCSII_Dental_DXTX.clsGroup
    Dim objDXs As CHCSII_Dental_DXTX.colDxs
    Dim objDx As CHCSII_Dental_DXTX.clsDx
    Dim objTXs As CHCSII_Dental_DXTX.colTxs
    Dim objTx As CHCSII_Dental_DXTX.clsTx
    Dim objTeeth As CHCSII_Dental_DXTX.colTeeth
    Dim GCTooth As CHCSII_Dental_DXTX.clsTooth
    Dim objSSs As CHCSII_Dental_DXTX.colSurfaceSets
    Dim objSS As CHCSII_Dental_DXTX.clsSurfaceSet
    Dim arrMissingTeeth() As String
    Dim intIndex As Integer
    Dim intIndexB As Integer
    Dim intCount As Integer
    Dim strMissingTeeth As String
    Dim blnDuplicate As Boolean
    
    strMissingTeeth = Empty
    
    If mobjDentalNOtes Is Nothing Then
        Exit Sub
    End If
    Set objGroups = mobjDentalNOtes.Groups
    
    If objGroups Is Nothing Then
        Exit Sub
    End If
    
    For Each objGroup In objGroups
        'In each Group, we have a collection of DX and TX's.
        Set objDXs = objGroup.Dxs
        Set objTXs = objGroup.Txs
        
        'Loop through the Diagnosis'
        For Each objDx In objDXs
            If objDx.DxICDCodeID = const_MISSING_TOOTH_DX And objDx.StatusID <> dxNonActiveDeleted Then
                'Diagnosis is for a missing tooth, and it was not deleted.
                Set objSSs = objDx.SurfaceSets
                
                'Each Diagnosis contains a collection of surface sets.
                'Loop through each Surface Set in the Surface Set collection
                For Each objSS In objSSs
                    'Each Surface set can contain multiple teeth. Currently each
                    'Surface set only contains a single tooth, but could potentially
                    'contain multiple teeth.
                    If objSS.SymbolID = 151 And InStr(objSS.SurfaceSet, ",") > 0 Then
                        'Process multiple missing teeth
                        strMissingTeeth = strMissingTeeth & Replace(objSS.SurfaceSet, ",", const_ITEM_DELIMETER) & const_ITEM_DELIMETER
                    Else
                        Set objTeeth = objSS.Teeth
                        For Each GCTooth In objTeeth
                            strMissingTeeth = strMissingTeeth & GCTooth.ToothID & const_ITEM_DELIMETER
                        Next
                    End If
                Next
                
            End If
        Next
        
        'Now look in the Treatments
        For Each objTx In objTXs
        
            With objTx
                'Inside the treatment, we check the txADACode. There are two extraction
                'codes that we need to look for.
                If .TxADACodeID = const_EXTRACTION_TOOTH_SNG_TX Or .TxADACodeID = const_EXTRACTION_TOOTH_MULTI_TX Then
                    'Its one of the Extraction codes. Make sure its Completed, InProgess,
                    'or pre exisiting
                    If .StatusID = txCompleted Or .StatusID = txInProgress Or .StatusID = txPreExisting Then
                        Set objSSs = .SurfaceSets
                        
                        'Loop through all of the Surface sets
                        For Each objSS In objSSs
                        
                            Set objTeeth = objSS.Teeth
                            
                            'Loop through all of the Teeth in each surface set.
                            For Each GCTooth In objTeeth
                                strMissingTeeth = strMissingTeeth & GCTooth.ToothID & const_ITEM_DELIMETER
                            
                            Next 'GCTooth
                            
                        Next 'objSS
                    End If
                End If
            End With
        Next 'objTx
        
    Next 'objGroup
    
    arrMissingTeeth = Split(strMissingTeeth, const_ITEM_DELIMETER)
    
    'We're going to search this array many times. Loop through and remove
    'any duplicate Tooth Id's.
    
    strMissingTeeth = Empty
    
    intCount = UBound(arrMissingTeeth) - 1
    
    For intIndex = 0 To UBound(arrMissingTeeth) - 1
    
        blnDuplicate = False
        
        'Loop through all of the Items already added and check to see if we've
        'already added this tooth.
        For intIndexB = (intIndex - 1) To 0 Step -1
            If arrMissingTeeth(intIndex) = arrMissingTeeth(intIndexB) Then
                blnDuplicate = True
                Exit For
            End If
            
        Next 'intIndexB
        
        If blnDuplicate = False Then
            strMissingTeeth = strMissingTeeth & arrMissingTeeth(intIndex) & const_ITEM_DELIMETER
        End If
    Next
    
    marrMissingTeeth = Split(strMissingTeeth, const_ITEM_DELIMETER)
    mintMissingToothCnt = UBound(marrMissingTeeth) - 1
    
    Set objGroups = Nothing
    Set objDXs = Nothing
    Set objTXs = Nothing
    Set objGroup = Nothing
    Set objDx = Nothing
    Set objTx = Nothing
    Set objTeeth = Nothing
    Set GCTooth = Nothing
    Set objSSs = Nothing
    Set objSS = Nothing
    
    Exit Sub
    
End Sub
'*******************************************************************
'Function: Get_User_Input
'Author: KDunne
'Date: 10/08/2001
'Returns: boolean. True if successful, False if error occurs or
'         User cancels.
'Purpose: Prompts user for Input. Displays buttons representing each
'         tooth, 1-32. An * next to the number indicates that the tooth
'         is missing. After user selects "OK", a | deliminated value
'         string is generated.
'*******************************************************************
Public Function Get_User_Input(Optional Parent As IToothCtrlParent = Nothing) As Boolean
    
On Error GoTo ErrHnd

    Dim blnReturn As Boolean
    Dim objForm As Object
    
    Select Case mlngControlType
        Case Radiographic
            Set objForm = frmToothControlRadio
            objForm.LoadStateInfo (mstrValueString)
        Case Region
            Set objForm = frmToothControlRegion
        Case Else
            Set objForm = frmToothControl
    End Select
    
    If Parent Is Nothing Then
        Set objForm.Parent = Me
    Else
        Set objForm.Parent = Parent
    End If
    
    Set objForm.Teeth = mobjTeeth
    
    If mstrFormCaption <> "" Then
        objForm.Caption = mstrFormCaption
    End If
    

    objForm.MissingTeethSelectable = mblnMissingTeethSelectable
    objForm.SingleSelect = mblnSingleSelect
    Load objForm
    
    blnReturn = objForm.Populate_Tool_Bar
    
    If blnReturn Then
        'This is not carried over when used in the image viewer because it is out of process
        If Not cwobjCWComm Is Nothing Then
            cwobjCWComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, cwlMyProgId
        End If
        
        objForm.Show vbModal
        
        If Not cwobjCWComm Is Nothing Then
            cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, cwlMyProgId
        End If
        
        If objForm.UserCancelled Then
            blnReturn = False
        Else
            Select Case mlngControlType
                Case Standard
                    mstrValueString = mobjTeeth.BuildValueString
                    mstrDisplayString = mobjTeeth.BuildDisplayString
                Case Region
                    If objForm.Tooth_Selected Or objForm.Region = 0 Then
                        mstrValueString = mobjTeeth.BuildValueString
                        mstrDisplayString = mobjTeeth.BuildNameString
                    Else
                        mstrValueString = objForm.Region
                        mstrDisplayString = Get_Region_Name(objForm.Region)
                    End If
                Case CanalMap
                    mstrValueString = mobjTeeth.BuildValueString
                    mstrDisplayString = mobjTeeth.BuildDisplayString
                    mlngToothCanal = objForm.Canal

            End Select

            blnReturn = True
        End If
        
        Unload objForm
        
    End If
    
    Get_User_Input = blnReturn
Exit Function

ErrHnd:
    MsgBox "Error in CToothControl.Get_User_Input: " & ": Error: " & Err.Description
    
End Function

'*******************************************************************
'Function: Get_Tooth_Name
'Author: KDunne
'Date: 10/24/2001
'Returns: string.
'Purpose: Given a tooth id (1-52), function returns the appropriate tooth name
'*******************************************************************

Public Function Get_Tooth_Name(ByVal lngToothId As Long) As String
    Dim strReturn As String
    
    Select Case lngToothId
        Case Is > 52
            strReturn = ""
        Case Is > 32
            'Tooth is secondary tooth in the 20 - 29 range
            strReturn = UCase(Chr(lngToothId - 32 + 64))
        Case Is < 0
            'Invalid tooth Number
            strReturn = Empty
        Case Else
            'Tooth is secondary
            strReturn = lngToothId
    End Select
     
    Get_Tooth_Name = Trim(strReturn)
End Function
'*****************************************************************
'Function:
'Author:KDunne
'Date:1/22/2002
'Purpose: Given a tooth, (1-52) or (A-T), function returns the
'           tooth's position (1-32)
'*****************************************************************
Public Function Get_Tooth_Position(ByVal strTooth As String) As Integer
    Dim intToothNum As Integer
    Dim intReturn As Integer
    
    If IsNumeric(strTooth) Then
        intToothNum = CInt(strTooth)
        
        Select Case intToothNum
            Case Is > 52
                intReturn = 0
            Case Is > 42
                intReturn = intToothNum - 23
            Case Is > 32
                intReturn = intToothNum - 29
            Case Else
                intReturn = intToothNum
        End Select
            
    Else
        intReturn = Get_Tooth_Number(strTooth)
    End If

    Get_Tooth_Position = intReturn
End Function
'*******************************************************************
'Function: Get_Tooth_Number
'Author: KDunne
'Date: 11/26/2001
'Returns: Number.
'Purpose: Given a tooth Name(A-T), function returns the appropriate tooth number
'*******************************************************************
Public Function Get_Tooth_Number(ByVal strToothName As String) As Long
    Dim lngReturn As Long
    

    If IsNumeric(strToothName) Then
        lngReturn = CLng(strToothName)
    Else
        lngReturn = Asc(strToothName) + 32 - 64
    End If
    
    If lngReturn <= 52 Then
        Get_Tooth_Number = lngReturn
    End If
End Function
'*******************************************************************
'Function: DB_Load_Teeth
'Author: KDunne
'Date: 10/24/2001
'Returns: ADO Recordset.
'Purpose: Given a Patient Unit Id, function returns all of the teeth in
'         that patient's mouth.
'*******************************************************************
Public Function DB_Load_Teeth(ByVal lngPatientUnitId As Long) As ADODB.Recordset
    Dim strSQL As String

    
'    strSQL = "SELECT Distinct DENTAL_ENC_VIEW_TEETH.TOOTH_VIEW, " & _
'              " DENTAL_ENC_VIEW_TEETH.TOOTH_ID, DENTAL_ENCOUNTER_TEETH.TOOTH_NAME " & _
'              " From DENTAL_ENC_VIEW_TEETH, DENTAL_ENCOUNTER_TEETH, DENTAL_ENCOUNTER " & _
'              " WHERE (DENTAL_ENC_VIEW_TEETH.PATIENT_UNIT_ID = " & lngPatientUnitId & ") AND (TOOTH_VIEW <> 0) AND " & _
'              " (DENTAL_ENCOUNTER.PATIENT_UNIT_ID = " & lngPatientUnitId & ") AND " & _
'              " (DENTAL_ENCOUNTER_TEETH.DENTAL_ENCOUNTER_ID = DENTAL_ENCOUNTER.DENTAL_ENCOUNTER_ID) AND " & _
'              " (DENTAL_ENCOUNTER_TEETH.TOOTH_ID = DENTAL_ENC_VIEW_TEETH.TOOTH_VIEW) "

    strSQL = "SELECT Distinct DENTAL_ENC_VIEW_TEETH.TOOTH_VIEW, " & _
              " DENTAL_ENC_VIEW_TEETH.TOOTH_ID " & _
              " From DENTAL_ENC_VIEW_TEETH " & _
              " WHERE (DENTAL_ENC_VIEW_TEETH.PATIENT_UNIT_ID = " & lngPatientUnitId & ") AND (TOOTH_VIEW > 0)"
    
    Set DB_Load_Teeth = mobjDAS.OpenRecordset(strSQL)
    
End Function
'Caller must set a reference to Dental Data Layer Notes object before
'this can be called.

'*******************************************************************
'Function: Load_Tooth_Information
'Author: KDunne
'Date: 10/24/2001
'Returns: boolean
'Purpose: Given a Patient Unit Id, function returns all of the teeth in
'         that patient's mouth.
'Notes:   Caller must set a reference to Dental Data Layer Notes object before
'         this can be called.
'*******************************************************************
Public Function Load_Tooth_Information(ByVal lngPatUnitId As Long, Optional ByVal strValues As String) As Boolean
    Dim strSQL As String
    Dim strName As String
    Dim objRS As ADODB.Recordset
    Dim blnMissing As Boolean
    Dim blnChecked As Boolean
    Dim lngToothView As Long
    Dim intIndex As Long
    Dim arrValues() As String
    Dim arrSVGTeeth() As String
    Dim intCounter As Integer
    Dim arrDefaultTeeth() As String
    'Dim oPositions As colPositions
    Dim blnImplant As Boolean
    Dim blnPontic As Boolean
    Dim strSVGPath As String
    Dim lngToothId As Long
    
    On Error GoTo ErrHandler
    
    Set mobjTeeth = New CTeeth
    
    strValues = Trim(strValues)
    
    If mlngControlType = Standard Or mlngControlType = Region Or mlngControlType = CanalMap Then
        If strValues <> "" Then
            InitializeString strValues
        End If
    End If

    mstrValueString = strValues
    'Debug.Assert False
    arrDefaultTeeth = Split(mobjSVGMgr.GetTeethPathFromRc(gintPatientAge), ",")
    'arrDefaultTeeth = Split(mobjSVGMgr.GetTeethPath(gintPatientAge), ",")
    
    'KDUnne 5/3/3003
    'SCR 23572 -The method of storing patient's current oral state has changed again.
    strSVGPath = GetTeethPath(mobjSVGMgr)
    arrSVGTeeth = Split(strSVGPath, const_ITEM_DELIMETER)
    
    Set mobjTeeth = New CTeeth
    LoadMissingTeeth
    
    For intCounter = 0 To UBound(arrSVGTeeth)
        If CheckToothMissing(arrSVGTeeth(intCounter)) Then
            blnMissing = True
        Else
            blnMissing = False
        End If

        blnImplant = CheckToothPath(arrSVGTeeth(intCounter), mstrImplantPath)
        blnPontic = CheckToothPath(arrSVGTeeth(intCounter), mstrPonticPath)
        
        If mlngControlType = Standard Or mlngControlType = Region Or mlngControlType = CanalMap Then
            'Check to see if user passed in values of buttons already pressed.
            'If they didn't, we'll never have a button pressed. Otherwise, check
            'to see if the current Tooth Id (TOOTH_VIEW) has been pressed.
            If strValues = "" Then
                blnChecked = False
            Else
                blnChecked = CheckToothPressed(arrSVGTeeth(intCounter))
            End If
        End If

        mobjTeeth.Add_Tooth arrSVGTeeth(intCounter), Get_Tooth_Position(arrSVGTeeth(intCounter)), _
            blnChecked, Get_Tooth_Name(arrSVGTeeth(intCounter)), blnMissing, blnImplant, blnPontic
    Next

    
    'KDunne 1/13/2004
    'SCRs 47764, 47766 - Need to pass in the original svg teeth path, and check to see
    '                    if the pre-selected values are not part of the patient's current
    '                    oral state.
    mobjTeeth.TeethPath = strSVGPath
    
    If mlngControlType = Standard Or mlngControlType = Region Or mlngControlType = CanalMap Then
    
        If strValues <> "" Then
            mobjTeeth.Initialize strValues 'SCRs 47764, 47766
            
            'Make sure that all of the pre-selected teeth are in our teeth collection.
            'Loop through array of values, check to see if its in the Teeth Collection,
            'if not then add it.
            For intIndex = 0 To UBound(marrValues)
                arrValues = Split(marrValues(intIndex), const_VALUE_DELIMETER)
                    
                lngToothId = Get_Tooth_Number(arrValues(0))
                If mobjTeeth.Load_Tooth_By_Number(lngToothId) = False Then
                    'Check to see if the tooth is missing
                    blnMissing = CheckToothMissing(lngToothId)
                    blnImplant = CheckToothPath(lngToothId, mstrImplantPath)
                    blnPontic = CheckToothPath(lngToothId, mstrPonticPath)
                    
                    mobjTeeth.Add_Tooth lngToothId, Get_Tooth_Position(lngToothId), True, Get_Tooth_Name(lngToothId), _
                        blnMissing, blnImplant, blnPontic
                End If
            Next
        End If
    End If
    
    'Make Sure all of our missing teeth are in the teeth collection
    For intIndex = 0 To UBound(marrMissingTeeth)
        If marrMissingTeeth(intIndex) <> "" Then
            If mobjTeeth.Load_Tooth_By_Number(marrMissingTeeth(intIndex)) = False Then
                blnImplant = CheckToothPath(marrMissingTeeth(intIndex), mstrImplantPath)
                blnPontic = CheckToothPath(marrMissingTeeth(intIndex), mstrPonticPath)
                
                mobjTeeth.Add_Tooth marrMissingTeeth(intIndex), Get_Tooth_Position(marrMissingTeeth(intIndex)), _
                    False, Get_Tooth_Name(marrMissingTeeth(intIndex)), True, blnImplant, blnPontic
                        
            End If
        
        End If
    Next
    
    Load_Tooth_Information = True
    Exit Function
    
ErrHandler:
    Load_Tooth_Information = False
    gobjCWShared.ShowVBError Err.Number, Err.Description, "CToothControl.Load_Tooth_Information", "", vbCritical
    
End Function
Public Function Load_Region_Information(ByVal lngPatUnitId As Long, Optional ByVal strValues As String = "") As Boolean
    Dim arrValues() As String
    Dim blnReturn As Boolean
    Dim objRS As ADODB.Recordset
    Dim lngLocId As Long
    Dim strTemp As String
    Dim lCounter As Long
    
    'SF 11-03-03 Clear out the combobox
    frmToothControlRegion.cmbLocations.Clear
    
    'SF 11-03-03 44279 The flaw in this logic is that it does not use the ID from the Note_Location table,
    'it assumes the ids in the Note_Location table start at 1 and are sequential.
    
    frmToothControlRegion.cmbLocations.AddItem ""
    Set objRS = DB_Load_Locations

    With objRS
        strTemp = Empty
        Do While Not .EOF
            strTemp = strTemp & .Fields("LOCATION_NAME") & const_ITEM_DELIMETER
            
            frmToothControlRegion.cmbLocations.AddItem .Fields("LOCATION_NAME")
            
            'Add the item data
            frmToothControlRegion.cmbLocations.ItemData(frmToothControlRegion.cmbLocations.ListCount - 1) = .Fields("LOCATION_ID")
            
            .MoveNext
        Loop
    End With
    marrRegions = Split(strTemp, const_ITEM_DELIMETER)
     
    If strValues <> "" Then
        arrValues = Split(strValues, const_ITEM_DELIMETER)
        
        If CInt(arrValues(0)) >= 100 Then
            'lngLocId = arrValues(0) - 100
            blnReturn = Load_Tooth_Information(lngPatUnitId)
            
            'SF 11-03-03 Set the user's selection
            For lCounter = 0 To frmToothControlRegion.cmbLocations.ListCount - 1
                If frmToothControlRegion.cmbLocations.ItemData(lCounter) = (arrValues(0) - 100) Then
                    frmToothControlRegion.cmbLocations.ListIndex = lCounter
                    Exit For
                End If
            Next
            
        Else
            blnReturn = Load_Tooth_Information(lngPatUnitId, strValues)
        
        End If
    Else
        blnReturn = Load_Tooth_Information(lngPatUnitId)
    End If
'    blnReturn = Load_Tooth_Information(lngPatUnitId)
    Load_Region_Information = blnReturn

End Function
Private Sub IToothCtrlParent_SetValueString(ByVal strValues As String, ByVal strDisplayString As String)
    mstrValueString = strValues
    mstrDisplayString = strDisplayString
End Sub

Private Sub IToothCtrlParent_UserCancelled()
    mblnCancelled = True
End Sub

'*******************************************************************
'Function: Get_Region_Name
'Author: KDunne
'Date: 11/04/2001
'Purpose: Takes a Region Code and returns the region name.
'*******************************************************************
Public Function Get_Region_Name(ByVal lngRegionCode As Long) As String
    Dim strReturn As String
    Dim strSQL As String
    Dim objRS As ADODB.Recordset
    
    If lngRegionCode > 100 Then
        strSQL = "Select LOCATION_NAME From NOTE_LOCATION WHERE LOCATION_ID = " & (lngRegionCode - 100)
    
        Set objRS = mobjDAS.OpenRecordset(strSQL)
        
        strReturn = objRS.Fields("LOCATION_NAME")
    Else
        strReturn = Empty
    End If


'    Select Case lngRegionCode
'
'    Case UpperRightSextant
'        strReturn = "Upper Right Sextant"
'    Case UpperAnteriorSextant
'     strReturn = "Upper Anterior Sextant"
'    Case UpperLeftSextant
'     strReturn = "Upper Left Sextant"
'    Case LowerLeftSextant
'     strReturn = "Lower Left Sextant"
'    Case LowerAnteriorSextant
'     strReturn = "Lower Anterior Sextant"
'    Case LowerRightSextant
'     strReturn = "Lower Right Sextant"
'    Case RightTMJ
'     strReturn = "Right TMJ"
'    Case LeftTMJ
'     strReturn = "Left TMJ"
'    Case Alveolar_mucosa
'     strReturn = "Alveolar mucosa"
'    Case Alveolar_ridge
'     strReturn = "Alveolar ridge"
'    Case Buccal_mucosa
'     strReturn = "Buccal mucosa"
'    Case Commissure
'     strReturn = "Commissure"
'    Case Floor_of_mouth
'     strReturn = "Floor of mouth"
'    Case Frenum
'    strReturn = "Frenum"
'    Case Gingiva
'     strReturn = "Gingiva"
'    Case Hamular_notch
'     strReturn = "Hamular notch"
'    Case Hard_palate
'     strReturn = "Hard palate"
'    Case Lower_lip
'     strReturn = "Lower lip"
'    Case Mandible
'     strReturn = "Mandible"
'    Case Maxilla
'     strReturn = "Maxilla"
'    Case Nasopharynx
'     strReturn = "Nasopharynx"
'    Case Neck
'     strReturn = "Neck"
'    Case Oral_Cavity
'     strReturn = "Oral Cavity"
'    Case palate
'     strReturn = "Palate"
'    Case Parotid_gland
'     strReturn = "Parotid gland"
'    Case Premaxilla
'     strReturn = "Premaxilla"
'    Case Retromolar_area
'     strReturn = "Retromolar area"
'    Case Sinus
'     strReturn = "Sinus"
'    Case Skin
'     strReturn = "Skin"
'    Case Soft_palate
'     strReturn = "Soft palate"
'    Case Sublingual_gland
'     strReturn = "Sublingual gland"
'    Case Submandibular_gland
'     strReturn = "Submandibular gland"
'    Case Tempromandibular_joint
'     strReturn = "Tempromandibular joint"
'    Case Tongue
'     strReturn = "Tongue"
'    Case Tonsillar_area
'     strReturn = "Tonsillar area"
'    Case Tuberosity
'     strReturn = "Tuberosity"
'    Case Upper_lip
'     strReturn = "Upper lip"
'    Case Uvula
'     strReturn = "Uvula"
'    Case Vestibule
'     strReturn = "Vestibule"
'End Select

Get_Region_Name = strReturn


End Function

'************************************************************
'Function DB_LOAD_LOCATIONS
'Author:KDunne
'Date 11/16/2001
'Purpose: Loads Location/Region Information from Note_Location table
'************************************************************
Private Function DB_Load_Locations() As ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select LOCATION_ID, LOCATION_NAME From NOTE_LOCATION ORDER BY LOCATION_ID"
    
    Set DB_Load_Locations = mobjDAS.OpenRecordset(strSQL)
    
End Function

'****************************************************************
'Function: Get_Canal_Description
'Author: KDunne
'Date: 11/26/2001
'Purpose: Given a long value, returns string description for a canal
'****************************************************************
Public Function Get_Canal_Description(ByVal lngValue As ToothControlCanalsEnum) As String
    Dim strReturn As String
    
    
    Select Case lngValue
        Case ToothControlCanalsEnum.Mesial
            strReturn = "Mesial(M)"
        Case ToothControlCanalsEnum.Mesiobuccal
            strReturn = "Mesiobuccal(MB)"
        Case ToothControlCanalsEnum.Mesiobuccal2
            strReturn = "Mesiobuccal2(MB2)"
        Case ToothControlCanalsEnum.Mesiobuccal3
            strReturn = "Mesiobuccal3(MB3)"
        Case ToothControlCanalsEnum.Mesiolingual
            strReturn = "Mesiolingual(ML)"
        Case ToothControlCanalsEnum.Distal
            strReturn = "Distal(D)"
        Case ToothControlCanalsEnum.DistalBuccal
            strReturn = "Distobuccal(DB)"
        Case ToothControlCanalsEnum.Distallingual
            strReturn = "Distolingual(DL)"
        Case ToothControlCanalsEnum.Facial
            strReturn = "Facial(F)"
        Case ToothControlCanalsEnum.Palatal
            strReturn = "Palatal(P)"
        Case ToothControlCanalsEnum.Palatal2
            strReturn = "Palatal2(P2)"
        Case ToothControlCanalsEnum.Lingual
            strReturn = "Lingual(L)"
        Case ToothControlCanalsEnum.NOS
            strReturn = "Not Otherwise Specified(NOS)"
        Case ToothControlCanalsEnum.None
            strReturn = Empty
        Case Else
            strReturn = Empty
    
    End Select
    
    Get_Canal_Description = strReturn
End Function
'**************************************************************************
'Sub: LoadMissingTeeth2
'Author:KDunne
'Date:1/24/2002
'Purpose: Compares the Default tooth path and the age path from SVG Comm server.
'         if a tooth is in the default path, but missing from the
'         age path, then that tooth is missing.
'***************************************************************************
Private Sub LoadMissingTeeth2(ByVal strDefaultPath As String, ByVal strAgePath As String)
    Dim strMissingTeeth As String
    Dim arrDefault() As String
    Dim arrAge() As String
    Dim blnMissing As Boolean
    Dim intDefaultTooth As Integer
    Dim x As Integer
    Dim y As Integer
    
    strMissingTeeth = Empty
    mintMissingToothCnt = 0
    
    arrDefault = Split(strDefaultPath, ",")
    arrAge = Split(strAgePath, ",")
    
    For x = 0 To UBound(arrDefault)
        intDefaultTooth = arrDefault(x)
        blnMissing = True
        'Look for the selected tooth within the age path
        For y = 0 To UBound(arrAge)
            If CInt(arrAge(y)) = intDefaultTooth Then
                blnMissing = False
                Exit For
            End If
        Next
    
        If blnMissing Then
            'Tooth is missing. Add it to our missing teeth array.
            strMissingTeeth = strMissingTeeth & intDefaultTooth & ","
            mintMissingToothCnt = mintMissingToothCnt + 1
        End If
    Next
    
    marrMissingTeeth = Split(strMissingTeeth, ",")
    

End Sub

'***********************************************************************
'Function:BuildDisplayString
'Author:KDunne
'Date:3/1/2002
'Purpose: Accepts a value string, then returns the appropriate display
'           string
'***********************************************************************
Public Function BuildDisplayString(ByVal strValues As String) As String
    Dim strReturn As String
    Dim arrValues() As String
    Dim x As Integer
    
    arrValues = Split(strValues, const_ITEM_DELIMETER)
    strReturn = Empty
    
    For x = 0 To UBound(arrValues)
        If IsNumeric(arrValues(x)) Then
            strReturn = strReturn & ", " & Get_Tooth_Name(arrValues(x))
        End If
    Next
    
    strReturn = Replace(strReturn, ", ", "", , 1)
       
    BuildDisplayString = strReturn

End Function

'*******************************************************************************
'Function:GetTeethPath
'Author:KDunne
'Date:5/3/2002
'Purpose: Loads the teeth path from SVGComm.
'********************************************************************************
Private Function GetTeethPath(ByRef objSVGMng As CHCSII_Dental_SVG_Manager.clsManage) As String
    Dim strPath As String
    Dim strImplantPath As String
    Dim strPonticPath As String
    
    Dim GCTooth As CHCSII_Dental_SVG_Manager.clsGCTooth
    
    On Error GoTo ErrHandler
    
    mstrImplantPath = Empty
    mstrPonticPath = Empty
    
    For Each GCTooth In objSVGMng.mColTeeth
        'We're only concerned with visible teeth.
        'Debug.Assert False
        If GCTooth.ToothVisible And GCTooth.IsToothPlaceHolder = False Then
            'Make sure that the tooth isn't missing
            If GCTooth.DateTimeWholeToothRemoved = CDate(0) Then
            
                If GCTooth.DateTimeRootRemoved <> CDate(0) Then
                    strPonticPath = strPonticPath & GCTooth.ToothID & const_ITEM_DELIMETER
                End If
                'SF 3-12-04 Replaced "," with const_ITEM_DELIMETER
                strPath = strPath & GCTooth.ToothID & const_ITEM_DELIMETER

            'KDunne 6/21/2002
            'SCR 25042 - Implant teeth are treated as a seperate case. They're loaded into
            '            both the teeth path and a module level implant path.
            ElseIf GCTooth.DateTimeImplant <> CDate(0) Then
                'Tooth is an implant
                strImplantPath = strImplantPath & GCTooth.ToothID & const_ITEM_DELIMETER
                'SF 3-12-04 Replaced "," with const_ITEM_DELIMETER
                strPath = strPath & GCTooth.ToothID & const_ITEM_DELIMETER
            End If 'dDateTimeToothRemoved
            
        End If 'bToothVisible
        'End scr 25042
    Next 'GCTooth
    
    strPath = ScrubPath(strPath)
    mstrImplantPath = ScrubPath(strImplantPath)
    mstrPonticPath = ScrubPath(strPonticPath)
    
FuncExit:
    Set GCTooth = Nothing
    GetTeethPath = strPath
    Exit Function

ErrHandler:
    gobjCWShared.ShowVBError Err.Number, Err.Description, "CToothControl.GetTeethPath", "", vbCritical
    mstrImplantPath = Empty
    mstrPonticPath = Empty
    strPath = Empty
    GoTo FuncExit
End Function

'******************************************************************
'Function:ScrubPath
'Author:KDunne
'Date:6/21/2002
'Purpose:Cleans a path string
'*******************************************************************
Private Function ScrubPath(ByVal strPath As String) As String
    
    strPath = Trim(strPath) 'get rid of trailing spaces
    If Len(strPath) <> 0 Then 'get rid of trailing commas
        strPath = Left(strPath, Len(strPath) - 1)
    End If
    
    ScrubPath = strPath
End Function
'*******************************************************************
'Function: CheckToothPath
'Author: KDunne
'Date: 10/24/2001
'Purpose: Checks to see if a given tooth is in the supplied tooth path
'*******************************************************************
Private Function CheckToothPath(ByVal lngToothId As Long, ByVal strPath As String) As Boolean
    Dim blnReturn As Boolean
    Dim intIndex As Integer
    Dim intCount As Integer
    Dim arrPath() As String
    
    blnReturn = False
    
    
    If strPath = Empty Then
        blnReturn = False
        GoTo FuncExit
    End If
    
    arrPath = Split(strPath, const_ITEM_DELIMETER)
    
    intCount = UBound(arrPath)
    
    For intIndex = 0 To intCount
        If arrPath(intIndex) = lngToothId Then
            blnReturn = True
            GoTo FuncExit
        End If
    Next 'intIndex
    
FuncExit:
    CheckToothPath = blnReturn
    
End Function
