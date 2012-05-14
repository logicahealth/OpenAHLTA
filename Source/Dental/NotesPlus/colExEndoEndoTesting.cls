VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colExEnTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements clsISection

Private mlngDentalEncounterId As Long
Private mlngPatUnitId As Long
Private mlngProviderId As Long
Private mlngDirtyFlag As DirtyFlagEnum
Private mcolItems As Collection
Private mobjDataLayer As CHCSII_Dental_DXTX.colDxs
Private mobjDentalNotes As CHCSII_Dental_DXTX.clsNotes
Private mblnCancel As Boolean

'D0460
Private Const mcPulpVitalityWorkload As Long = 346
Public Property Get Cancel() As Boolean
    Cancel = mblnCancel
End Property
Private Sub DetermineWorkload()

    'Workload vars
    Dim oTest As clsExamEndoTest
    Dim oWorkloadTest As clsExamEndoTest
    Dim bAppliedWorkload As Boolean
    Dim bApplyWorkload As Boolean
    Dim lWorkloadID As Long
    
    '6/23/03 - 36297
    'Assign workload once per encounter (once per all instances of oTest)

    For Each oTest In mcolItems
        
        If oTest.WorkloadId <> 0 Then
            'Workload has been applied to this object, so remember it (in case we need to delete it) and remove it from this object
            bAppliedWorkload = True
            lWorkloadID = oTest.WorkloadId
            oTest.WorkloadId = 0
        End If
        
        If (oTest.EPT <> ExamEndoEPTEnum.exEndoEPT_Not_Addressed _
            Or oTest.Cold_Response_Test <> ExamEndoColdResponseTestEnum.Not_Addressed _
            Or oTest.Cold_Response <> ExamEndoColdResponseTestEnum.Not_Addressed _
            Or oTest.Cold_Linger <> ExamEndoColdResponseTestEnum.Not_Addressed _
            Or oTest.Heat_Response_Test <> ExamEndoHeatResponseTestEnum.Not_Addressed _
            Or oTest.Heat_Response <> ExamEndoHeatResponseTestEnum.Not_Addressed _
            Or oTest.Heat_Linger <> ExamEndoHeatResponseTestEnum.Not_Addressed) And oTest.Dirty_Flag <> dfdelete Then
            
            'Workload should be applied to an object
            bApplyWorkload = True
            Set oWorkloadTest = oTest

        End If
    Next
    
    If bApplyWorkload = True And bAppliedWorkload = False Then
        'Add the workload
        UpdateWorkload lWorkloadID, dfAdd, mlngDentalEncounterId, mcPulpVitalityWorkload, 1
        oWorkloadTest.WorkloadId = lWorkloadID
        
    ElseIf bApplyWorkload = True And bAppliedWorkload = True Then
        'Update the workload
        UpdateWorkload lWorkloadID, dfUpdate, mlngDentalEncounterId, mcPulpVitalityWorkload, 1
        oWorkloadTest.WorkloadId = lWorkloadID
    
    ElseIf bApplyWorkload = False And bAppliedWorkload = True Then
        'Delete the workload
        UpdateWorkload lWorkloadID, dfdelete, mlngDentalEncounterId, mcPulpVitalityWorkload, 1
    End If
    
    Set oWorkloadTest = Nothing
    
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = mcolItems.[_NewEnum]
End Property
Public Property Get Item(vntIndexKey As Variant) As clsTreatPGraft
  Set Item = mcolItems(vntIndexKey)
End Property
Public Property Get Count() As Long
    Count = mcolItems.Count
End Property
Public Property Get ProviderId() As Long
    ProviderId = mlngProviderId
End Property
Public Property Let ProviderId(ByVal lngValue As Long)
    mlngProviderId = lngValue
End Property

Public Property Get Endontic_Test(ByVal intIndex As Integer) As clsExamEndoTest
    Dim objTest As clsExamEndoTest
    
    If intIndex > mcolItems.Count Then
        Set objTest = New clsExamEndoTest
        objTest.PatientID = mlngPatUnitId
        objTest.Dental_EncounterId = mlngDentalEncounterId
        objTest.ProviderId = mlngProviderId
        Set objTest.DentalNotes = mobjDentalNotes
        mcolItems.Add objTest
    End If
    
    Set Endontic_Test = mcolItems(intIndex)
        
    Set objTest = Nothing
End Property

'*************************************************************************
'Sub:Clean_Tests
'Author:KDunne
'Date: 2/1/2001
'Purpose:LOops through all of the tests in mcolItems and removes tests
'        with no Tooth_Number
'*************************************************************************
Public Sub Clean_Tests()
    Dim X As Integer
    Dim intCount As Integer
    Dim objTest As clsExamEndoTest
    
    intCount = mcolItems.Count
    For X = 1 To intCount
        Set objTest = mcolItems(X)
        
        With objTest
            If IsNumeric(.Tooth_Number) Then
                If Val(.Tooth_Number) <= 0 Then
                    Remove_Test X
                    intCount = mcolItems.Count
                End If '.Tooth_Number <= 0
            Else
                Remove_Test X
                intCount = mcolItems.Count
            End If 'IsNumeric
        End With 'objTest
    Next
    
    Set objTest = Nothing
End Sub
Public Sub Remove_Test(ByVal intIndex As Integer)
    On Error GoTo ErrHandler
    Dim objTest As clsExamEndoTest
    
    'Need to Delete Item from DB first.
    Set objTest = mcolItems.Item(intIndex)
    
    objTest.Dirty_Flag = dfdelete
    
    DetermineWorkload
    
    objTest.SaveSection
    
    mcolItems.Remove (intIndex)
    
FuncExit:
    Set objTest = Nothing
    
    Exit Sub
ErrHandler:
    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & "." & " Property Set DataLayer", App.Title, vbCritical
    GoTo FuncExit
End Sub
Public Property Set DataLayer(ByVal objDataLayer As CHCSII_Dental_DXTX.colDxs)
    On Error GoTo ErrHandler
    
    Set mobjDataLayer = objDataLayer
    
    Exit Property
    
ErrHandler:
    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & "." & " Property Set DataLayer", App.Title, vbCritical
    Set mobjDataLayer = Nothing
    
End Property
Public Property Get DentalNotes() As CHCSII_Dental_DXTX.clsNotes
    Set DentalNotes = mobjDentalNotes
End Property
Public Property Set DentalNotes(ByVal objObject As CHCSII_Dental_DXTX.clsNotes)
    On Error GoTo ErrHandler
    
    Set mobjDentalNotes = objObject
    
    Exit Property
ErrHandler:
    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & "." & " Property Set DentalNotes", App.Title, vbCritical
    Set mobjDentalNotes = Nothing
End Property
'*******************************************************************
'Sub: SortTests
'AUthor: KDunne
'Date: 11/15/2001
'Purpose: SOrts tests based on tooth number.
'Notes: Sorting is based on Tooth number. Primary Teeth(A-T) come after
'       secondary Teeth (1-32).
'*******************************************************************
Public Sub SortTests()
    Dim objTest As clsExamEndoTest
    Dim colSort As Collection
    Dim colUnsorted As Collection
    Dim strCurrTooth As String
    Dim lngCurrIndex As Long
    Dim lngIndex As Long
    
    If mcolItems Is Nothing Then
        Exit Sub
    End If
    
    Set colSort = New Collection
    Set colUnsorted = mcolItems
    
    strCurrTooth = Empty
    lngIndex = 0
    
    Do While colUnsorted.Count > 0
        strCurrTooth = Empty
        lngIndex = 0
        For Each objTest In colUnsorted
            lngIndex = lngIndex + 1
            If strCurrTooth = "" Then
                strCurrTooth = objTest.Tooth_Number
                lngCurrIndex = lngIndex
            ElseIf IsNumeric(strCurrTooth) Then
                'Previous tooth was numeric. The only way the current
                'tooth can replace him is if he is a lower number.
                If IsNumeric(objTest.Tooth_Number) Then

                    If CInt(strCurrTooth) > CInt(objTest.Tooth_Number) Then
                        strCurrTooth = objTest.Tooth_Number
                        lngCurrIndex = lngIndex
                    End If
                End If
            Else
                'Previous tooth is a primary tooth (A-T). If
                'Current tooth is a number, it goes first. Else,
                'If current tooth is also a letter, see which one is lower.
                If IsNumeric(objTest.Tooth_Number) Then
                    strCurrTooth = objTest.Tooth_Number
                    lngCurrIndex = lngIndex
                Else
                    If Asc(objTest.Tooth_Number) < Asc(strCurrTooth) Then
                        strCurrTooth = objTest.Tooth_Number
                        lngCurrIndex = lngIndex
                    End If
                End If
            
            End If
        Next
        
        colSort.Add colUnsorted.Item(lngCurrIndex)
        colUnsorted.Remove lngCurrIndex
    Loop
    
    Set mcolItems = colSort
    
    Set colSort = Nothing
    Set colUnsorted = Nothing
    Set objTest = Nothing
    
End Sub
Private Property Get clsISection_DentalEncounterID() As Long
    clsISection_DentalEncounterID = mlngDentalEncounterId
End Property

Private Function clsISection_GetNote() As String
    Dim strNote As String
    Dim objTest As clsExamEndoTest
    Dim blnTestWritten As Boolean
    Dim blnHeatTestWritten As Boolean
    
    strNote = Empty
    blnTestWritten = False
    blnHeatTestWritten = False
    
    If mcolItems Is Nothing = False Then
        For Each objTest In mcolItems
            objTest.HeatTestWritten = blnHeatTestWritten
            objTest.ColdTestWritten = blnTestWritten
            strNote = strNote & objTest.GetNote
            blnTestWritten = objTest.ColdTestWritten
            blnHeatTestWritten = objTest.HeatTestWritten
        Next
    End If
    
    If strNote <> "" Then
        strNote = "Endodontics Testing: " & strNote
    End If
    clsISection_GetNote = strNote
End Function

Private Function clsISection_GetRTFNote() As String
    Dim strNote As String
    Dim objTest As clsExamEndoTest
    Dim blnTestWritten As Boolean
    Dim blnHeatTestWritten As Boolean
    Dim lngTempColdTest As ExamEndoColdResponseTestEnum
    Dim lngTempHeatTest As ExamEndoHeatResponseTestEnum
    
    strNote = Empty
    blnTestWritten = False
    blnHeatTestWritten = False
    
    If mcolItems Is Nothing = False Then
        For Each objTest In mcolItems
            'KDunne 8/16/2002
            'SCR 21561 - Need to document multiple Heat/Cold Tests
            If objTest.Cold_Response_Test = lngTempColdTest Then
                objTest.ColdTestWritten = blnTestWritten
            Else
                objTest.ColdTestWritten = False
            End If
            
            If objTest.Heat_Response_Test = lngTempHeatTest Then
                objTest.HeatTestWritten = blnHeatTestWritten
            Else
                objTest.HeatTestWritten = False
            End If
            lngTempColdTest = objTest.Cold_Response_Test
            lngTempHeatTest = objTest.Heat_Response_Test
            'End SCR 21561
            strNote = strNote & objTest.GetNote
            blnTestWritten = objTest.ColdTestWritten
            blnHeatTestWritten = objTest.HeatTestWritten
        Next
    End If
    
    If strNote <> "" Then
        strNote = crtfStartControlName & "Endodontics Testing: " & strNote & _
            crtfEndControlName
    End If
    clsISection_GetRTFNote = strNote
End Function

Private Function clsISection_LoadSection(ByVal PatientID As Long, ByVal DentalEncounterID As Long) As Boolean
    Dim blnReturn As Boolean
    Dim objRS As ADODB.Recordset
    Dim objTest As clsExamEndoTest
    
    On Error GoTo ErrHandler
    
    mlngPatUnitId = PatientID
    mlngDentalEncounterId = DentalEncounterID
    
    Set objRS = DB_Load_Endontic_Testing(mlngDentalEncounterId)
    Set mcolItems = New Collection
    
    With objRS
        If .EOF Then
            mlngDirtyFlag = dfAdd
            'reset property values here
            
        Else
            mlngDirtyFlag = dfUpdate
            With objRS
                On Error Resume Next
                Do While .EOF = False
                    Set objTest = New clsExamEndoTest
                    Set objTest.DentalNotes = mobjDentalNotes
                    objTest.ProviderId = mlngProviderId
                    objTest.PatientID = mlngPatUnitId
                    objTest.Tooth_Number = .Fields("TOOTH_NUMBER") & ""
                    objTest.Anatomic_Variations = .Fields("ANATOMIC_VARIATIONS") & ""
                    objTest.Caries = .Fields("CARIES")
                    objTest.Cold_Linger = .Fields("COLD_LINGER")
                    objTest.Cold_Response = .Fields("COLD_RESPONSE")
                    objTest.Cold_Response_Test = .Fields("COLD_RESPONSE_TEST")
                    objTest.Dental_EncounterId = mlngDentalEncounterId
                    objTest.EPT = .Fields("EPT")
                    objTest.Heat_Response_Test = .Fields("HEAT_TEST")
                    objTest.Heat_Linger = .Fields("HEAT_LINGER")
                    objTest.Heat_Response = .Fields("HEAT_RESPONSE")
                    objTest.Hx_Bleaching = .Fields("BLEACHING_HX")
                    objTest.Hx_Comments = .Fields("HX_COMMENTS") & ""
                    objTest.Magnification_Findings = .Fields("MAGNIFICATION")
                    objTest.Microleakage = .Fields("MICROLEAKAGE")
                    If .Fields("MOBLITIY") = 4 Then
                        objTest.Mobility = .Fields("MOBLITIY")
                    End If
                    objTest.Pain_On_Occulusion = .Fields("PAIN_ON_OCCLUSION")
                    objTest.Pain_On_Occulusion_Location = .Fields("PAIN_ON_OCCLUSION_LOC") & ""
                    objTest.Pain_On_Release = .Fields("PAIN_ON_RELEASE")
                    objTest.Pain_On_Release_Location = .Fields("PAIN_ON_RELEASE_LOC") & ""
                    objTest.Pain_On_Palpation = .Fields("PALPATION")
                    objTest.Palpation_Location = .Fields("PALPATION_LOCATION")
                    objTest.Percussion = .Fields("PERCUSSION")
                    objTest.Probing = .Fields("PROBING") & ""
                    objTest.Radiograph_Findings = .Fields("RADIOGRAPH_FINDINGS")
                    objTest.Radiograph_Findings_Comments = .Fields("RADIOGRAPH_FINDINGS_COMMENTS") & ""
                    objTest.Restorative_Hx = .Fields("RESTORATIVE_HX") & ""
                    objTest.Sinus_Tract = .Fields("SINUS_TRACT")
                    objTest.ST_Comments = .Fields("ST_COMMENTS") & ""
                    objTest.ST_Termination = .Fields("ST_TERMINATION")
                    objTest.ST_Traced_To = .Fields("ST_TRACED_TO")
                    objTest.Test_Id = .Fields("EX_ENDO_TEST_ID")
                    objTest.Transillumination = .Fields("TRANSILLUMINATION") & ""
                    objTest.Trauma = .Fields("TRAUMA")
                    objTest.WorkloadId = .Fields("Workload_ID") & ""
                    objTest.Dirty_Flag = dfNoAction
                        
                    mcolItems.Add objTest
                
                    .MoveNext
                Loop
            
            End With
            'Set property values here
        End If
    End With
    
FuncExit:
    clsISection_LoadSection = blnReturn
    Set objRS = Nothing
    Exit Function
ErrHandler:
   blnReturn = False
    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & "." & " LoadSection", App.Title, vbCritical

   GoTo FuncExit
End Function

Private Property Get clsISection_PatientID() As Long
    clsISection_PatientID = mlngPatUnitId
End Property

Private Function clsISection_SaveSection() As Boolean
    Dim strSql As String
    Dim blnReturn As Boolean
    Dim objRS As ADODB.Recordset
    Dim objTest As clsExamEndoTest
    Dim blnFound As Boolean
    
    'KDunne
    'Collection class contains no properties of its own, and there isn't a corresponding
    'table associated with it.
   
   DetermineWorkload
   
   mlngDirtyFlag = dfNoAction
    
    'If there's a collection, loop through it here.
    blnReturn = True
    'mblnCancel = False
    If mcolItems Is Nothing = False Then
        For Each objTest In mcolItems
            'blnFound = False
            objTest.Dental_EncounterId = mlngDentalEncounterId
            
            If objTest.SaveSection Then
                blnReturn = True

                
'                If objTest.Mobility <> ExamEndoMobilityEnum.Not_Addressed And objTest.Mobility <> ExamEndoMobilityEnum.WNL Then
'                    'KDunne We need to wait for Datalayer to be updated. The tooth is a chartable
'                    'location. Currently, the notes datalayer object doesn't support
'                    'chartable locations.
'                    blnReturn = UpdateDataLayer(objTest.Mobility, objTest.Tooth_Number)
'                End If
            Else
                blnReturn = False
                'mblnCancel = True
            End If
        Next
    End If
    
    clsISection_SaveSection = blnReturn
    Set objRS = Nothing
End Function

Private Function UpdateDataLayer(ByVal lngMobility As ExamEndoMobilityEnum, ByVal lngToothNumber As Long) As Boolean
    Dim lngICDCode As DxTxMobilityICDEnums
    Dim objGroups As CHCSII_Dental_DXTX.colGroups
    Dim objGroup As CHCSII_Dental_DXTX.clsGroup
    Dim objDx As CHCSII_Dental_DXTX.clsDx
    Dim objDxs As CHCSII_Dental_DXTX.colDxs
    Dim objSS As CHCSII_Dental_DXTX.clsSurfaceSet
    Dim objTooth As CHCSII_Dental_DXTX.clsTooth
    Dim objLocation As CHCSII_Dental_DXTX.clsLocation
    Dim objToothCtrl As CToothControl
    Dim blnFound As Boolean
    Dim blnFoundTooth As Boolean
    Dim blnFoundSS As Boolean
    Dim blnReturn As Boolean
    
    If mobjDentalNotes Is Nothing Then
        UpdateDataLayer = True
        Exit Function
    End If
    
    Select Case lngMobility
        Case ExamEndoMobilityEnum.Mobility1
            lngICDCode = MillerMobility1
        Case ExamEndoMobilityEnum.Mobility2
            lngICDCode = MillerMobility2
        Case ExamEndoMobilityEnum.Mobility3
            lngICDCode = MillerMobility3
        Case ExamEndoMobilityEnum.ANK
            lngICDCode = Ankylosis
        Case ExamEndoMobilityEnum.WNL, ExamEndoMobilityEnum.Not_Addressed
            lngICDCode = 0
    End Select
    
    If lngICDCode > 0 Then
        'User has selected a mobility value.
        Set objGroups = mobjDentalNotes.Groups
        
        If objGroups Is Nothing Then
            UpdateDataLayer = True
            Exit Function
        End If
        
        Set objToothCtrl = New CToothControl
        
        For Each objGroup In objGroups
        
            blnFound = False
            'See if we already have our ICD Code.
            Set objDxs = objGroup.Dxs
            For Each objDx In objDxs
                If objDx.DxICDCodeID = lngICDCode Then
                    blnFound = True
                    Exit For
                End If
            Next 'objDX
            
            If blnFound Then
                blnFoundSS = False
                blnFoundTooth = False
                
                For Each objSS In objDx.SurfaceSets
                    If objSS.SurfaceSet = 1 Then
                        blnFoundSS = True
                        For Each objTooth In objSS.Teeth
                            If objTooth.ToothID = lngToothNumber Then
                                blnFoundTooth = True
                                GoTo SaveSection
                            End If
                        Next 'ObjTooth
                        
                        GoTo SaveSection
                    End If
                Next 'ObjSS
            End If
        Next 'objGroup
        
        
SaveSection:
        If blnFoundSS Then
            If blnFoundTooth Then
                'Do Nothing. Our Surface set and tooth already contain the diagnosis
            Else
                'We're just missing the tooth from our surface set. Add it here.
                Set objTooth = objSS.Teeth.Add(mlngDentalEncounterId)
                objTooth.ToothID = lngToothNumber
                objTooth.ToothName = objToothCtrl.Get_Tooth_Name(lngToothNumber)
                objSS.DirtyFlag = dfUpdate
                objDx.DirtyFlag = dfUpdate
                
            End If
        Else
            If Not blnFound Then
                Set objGroup = objGroups.Add
                Set objDx = objGroup.Dxs.Add(sNotes, lngICDCode, mlngProviderId)
            End If

            objDx.Chartable = True
            Set objSS = objDx.SurfaceSets.Add()
            objSS.SurfaceSet = 1
            objSS.SurfaceAbbr = 1
            Set objTooth = objSS.Teeth.Add(mlngDentalEncounterId)
            objTooth.ToothID = lngToothNumber
            objTooth.ToothName = objToothCtrl.Get_Tooth_Name(lngToothNumber)
            
        End If
        
        objDx.DirtyFlag = dfUpdate

        mobjDentalNotes.SaveGroups
        
    Else
        'User hasn't selected a mobility value
        blnReturn = True
    End If
    
    UpdateDataLayer = blnReturn
    
    Set objToothCtrl = Nothing
    Set objGroups = Nothing
    Set objGroup = Nothing
    Set objDxs = Nothing
    Set objDx = Nothing
    Set objSS = Nothing
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

Public Function DB_Load_Endontic_Testing(ByVal lngDentalEncId As Long) As ADODB.Recordset
    Dim strSql As String
    
    strSql = "Select EX_ENDO_TEST_ID, TOOTH_NUMBER, EPT, COLD_RESPONSE, " & _
                    " COLD_RESPONSE_TEST, HEAT_RESPONSE, COLD_LINGER, " & _
                    " HEAT_LINGER, PERCUSSION, PALPATION, HEAT_TEST," & _
                    " PALPATION_LOCATION, SINUS_TRACT, ST_TRACED_TO, " & _
                    " ST_COMMENTS, ST_TERMINATION, MAGNIFICATION, " & _
                    " TRANSILLUMINATION, MOBLITIY, PROBING, " & _
                    " PAIN_ON_OCCLUSION, PAIN_ON_OCCLUSION_LOC,  " & _
                    " PAIN_ON_RELEASE, PAIN_ON_RELEASE_LOC, CARIES,  " & _
                    " RADIOGRAPH_FINDINGS, " & _
                    " RADIOGRAPH_FINDINGS_COMMENTS, TRAUMA, " & _
                    " ANATOMIC_VARIATIONS, RESTORATIVE_HX, MICROLEAKAGE, " & _
                    " BLEACHING_HX , Hx_Comments, Workload_ID " & _
                    " From TEMPL_EXAM_ENDO_ENDONTIC_TEST WHERE DENTAL_ENCOUNTER_ID = " & lngDentalEncId


    Set DB_Load_Endontic_Testing = goconn.OpenRecordset(strSql)
End Function

