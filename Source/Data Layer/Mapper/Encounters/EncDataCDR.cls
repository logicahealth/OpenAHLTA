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
Attribute VB_Name = "EncDataCDR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements iEncDataMapper

Private mConn As CHCSII_CONN.Conn
Private mCache As CacheMapper

Private mMode As CHCSII_CONN.eBackend       'Data Access mode used for DAS object
Private mUseCache As Boolean

Private mEncounter As Encounter


Private Function iEncDataMapper_EncounterList(Params As DL_Support.iParameters) As DTOs.Encounter
    Dim param As DM_Encounters.ParamsEncList
    
    Set param = Params
    
End Function

Private Sub iEncDataMapper_Init(Conn As CHCSII_CONN.Conn, Cache As DL_Support.CacheMapper, ByVal DataAccessMode As CHCSII_CONN.eBackend, Optional ByVal CacheConfigItems As Boolean = False)
  Set mConn = Conn
  Set mCache = Cache
  
  mMode = DataAccessMode
  mUseCache = CacheConfigItems
End Sub

Private Function iEncDataMapper_Retrieve(Params As DL_Support.iParameters) As DTOs.Encounter

    Dim param As DM_Encounters.ParamsEncRetrieve
    Dim colDataStrings As Collection
    
    Set param = Params
    
    Set colDataStrings = GetDataStrings(param.PatientUnitNumber, param.FacilityNCID, param.EncounterNumber)
    
    Set iEncDataMapper_Retrieve = LoadDataStrings(colDataStrings, param.IncludeChildModuleData)
    
End Function

Private Function iEncDataMapper_Save(Encounter As DTOs.Encounter) As Boolean

   Dim bResult          As Boolean
    Dim i                As Integer
    Dim objData          As DataString
    Dim objEvent         As ClinicalEvent
    Dim objHeader        As ClinObsHeader
    Dim objClinObs       As ClinicalObservation
    'Dim objSectionParent As SectionParent
    Dim oEmh             As EmhCtrl
    Dim objProvider      As EncProvider
    Dim objEncRTF       As Document
    Dim bRaiseEvent      As Boolean
    Dim lDocumentNumber  As Long
    Dim msRTF            As String
    Dim State            As iState
    Dim bSuccess         As Boolean

    On Error GoTo ErrHandler

    'If UpdateLock is set and wasn't just set, then don't perform the save
    With Encounter
        If .UpdateLock Then
            iEncDataMapper_Save = True
            Exit Function
        End If
    
        If StartSectionUpdate(eHeader, 0) Then
            'RASIE ERROR
             '--- SCR-ARM-25468 RWalker 4/10/03 Must manually stop a correlated tx
            'goARM.StopCmplxTx gsARM_Tx_CorrUT_SignFamily, enmARMTxStatusFailed, enmARMStop_MeOnly
            GoTo Cleanup
        End If
    
        If .Status = complete Or .Status = NeedsCoSignature Or .Status = Updated Then 'SCR-34229
            'If there is no E&M Code then SADR Status has to be NO_COUNT SCR-26152
            If Len(.Sections(eDisp).mobjSection.EandMNCID) = 0 And (.SADRStatus = NCID_SADR_NEEDS_NEW_PROCESSING Or .SADRStatus = NCID_SADR_NEEDS_UPDATE_PROCESSING) Then
                .SADRStatus = NCID_SADR_NO_COUNT
            End If
            'Also
            'If there is no Primary Dx then SADR Status has to be NO_COUNT SCR-26152
            If .PrimaryDxSnoID = 0 And (.SADRStatus = NCID_SADR_NEEDS_NEW_PROCESSING Or .SADRStatus = NCID_SADR_NEEDS_UPDATE_PROCESSING) Then
                .SADRStatus = NCID_SADR_NO_COUNT
            End If
        End If
    
        'Reset msRTF
        msRTF = ""
    
        Set objData = InitializeDataString(eHeader, .DataID, .CreatedDate, .WhenSignedDate, Encounter)
        Set objEvent = objData.PatientData
        Set objHeader = objEvent.Header
        If .Providers.Count > 0 Then
          objData.Clinician.ValueConcept.UniqueId = .Providers.Item(1).NCID
        Else
          objData.Clinician.ValueConcept.UniqueId = 0
        End If
        If CInt(.Status) <> 0 Then
            objHeader.TestStatus.InfoType = "att"
            objHeader.TestStatus.Value.InfoType = "codedWOSform"
            objHeader.TestStatus.Value.CodedWOSform.ValueConcept.UniqueId = StatusEnumToNCID(.Status)
            bRaiseEvent = True
        End If
        Call AddCommentValuePair(objHeader, "WASUPDATED", .WasUpdated)
        If .WhenSignedDate > 0 Then
            Call AddCommentValuePair(objHeader, "WHENSIGNED", .WhenSignedDate)
            Call AddCommentValuePair(objHeader, "WHOSIGNED", .WhoSigned.NCID)
            Call AddCommentValuePair(objHeader, "SIGNERLINE2", .WhoSigned.SigLine2)
            Call AddCommentValuePair(objHeader, "SIGNERLINE3", .WhoSigned.SigLine3)
        End If
        If .WhenCosigned > 0 Then
            Call AddCommentValuePair(objHeader, "WHENCOSIGNED", .WhenCosigned)
            Call AddCommentValuePair(objHeader, "WHOCOSIGNED", .WhoCosigned.NCID)
            Call AddCommentValuePair(objHeader, "COSIGNERLINE2", .WhoCosigned.SigLine2)
            Call AddCommentValuePair(objHeader, "COSIGNERLINE3", .WhoCosigned.SigLine3)
        End If
        Call AddCommentValuePair(objHeader, "ENCTYPE", .ENCTYPE)
        Call AddCommentValuePair(objHeader, "ENCCAT", .Appointment.TypeAppt)
        'Call AddCommentValuePair(objHeader, "TIMEZONE",     )
        'Call AddCommentValuePair(objHeader, "EANDMCODEREVIEW", msEandMCodeReview)   'DEM -- 05/04/00
        'Call AddCommentValuePair(objHeader, "ALLERGYVERIFIEDBYDATE", .Sections(RFV).AllergyVerifiedDate)
        'Call AddCommentValuePair(objHeader, "ALLERGYVERIFIEDBYNCID", .Sections(RFV).AllergyVerifiedByNCID)
        Call AddCommentValuePair(objHeader, "VERSION", App.Major & "." & App.Minor & "." & App.Revision)
        Call AddCommentValuePair(objHeader, "APPTDTS", .Appointment.DateTime)
        Call AddCommentValuePair(objHeader, "UPDATELOCK", CInt(.UpdateLock))
        'Call AddCommentValuePair(objHeader, "DXSnoID", .)
        'Call AddCommentValuePair(objHeader, "DXMEDCINPREFIX", msPrimaryDxMedcinPrefix)
        'Call AddCommentValuePair(objHeader, "APSTATUSCOMMENT", msAPStatusComment)  'SCR-17942
        Call AddCommentValuePair(objHeader, "ENCNOCOUNT", .Class)
        Call AddCommentValuePair(objHeader, "WORKLOADOPTIONS", .Appointment.WorkloadTypeOption)
        Call AddCommentValuePair(objHeader, "WORKLOAD", .Appointment.WorkLoad)
        Call AddCommentValuePair(objHeader, "MEPRSCODES", .Appointment)
        Call AddCommentValuePair(objHeader, "PATIENTSTATUS", .Appointment.Classification)
        Call AddCommentValuePair(objHeader, "APPTIEN", .Appointment.IEN)
        Call AddCommentValuePair(objHeader, "TEMPLATEID", .TemplatedID) 'SCR-44706 RBELL 11/3/03  Save data in a text item in db
    
        'HIPAA Rqmt 'SCR#43283
    '    If Len(mobjShared.CmdLineSwitches("HIPAA837")) Then
    '        Call AddCommentValuePair(objHeader, "HIPAA_CODE", .)
    '    End If
    
    
        If .Sensitivity Then 'SCR-35365
            Call AddCommentValuePair(objHeader, "SENSITIVE", .Sensitivity)
        End If
    
        If .Appointment.ResultedConsultOrder Then
            Call AddCommentValuePair(objHeader, "RESULTEDCONSULTORDER", "T")
        End If
    
        If .Appointment.ArrivedByAmbulance Then
            Call AddCommentValuePair(objHeader, "ARRIVEDBYAMBULANCE", "T")
        End If
    
        If Not .Cosigner Is Nothing Then
            Call AddClinObs(objEvent, NCID_NEEDS_COSIGN, .Cosigner.NCID)
        End If
        If .Providers.Count > 0 Then
            Call AddCommentValuePair(objHeader, "PRIMPROV", .Providers(1).NCID)
        End If
        If .ENCTYPE = TeleConsult Then
            Call AddClinObs(objEvent, NCID_APPOINTMENT_TYPE, NCID_TELECON_APPT_TYPE)
        End If
    '    If Val(.PrimaryDxNCID) > 0 Then
    '        Call AddClinObs(objEvent, NCID_DIAGNOSIS, .PrimaryDxNCID)
    '    End If
    
        If Val(.SADRStatus) > 0 Then
            Call AddClinObs(objEvent, NCID_SADR_STATUS, .SADRStatus)
        End If 'val sadrstatus
        If Val(.Appointment.Id) > 0 Then
            Call AddClinObs(objEvent, NCID_APPOINTMENT_ID, .Appointment.Id, "codedWOSform", "basicString")
        End If
        
        For i = 2 To .Providers.Count
            Set objProvider = .Providers(i)
            If Val(objProvider.NCID) > 0 Then
                Set objClinObs = AddClinObs(objEvent, NCID_ADDITIONALPROVIDER, objProvider.NCID)
                If Val(objProvider.RoleNcid) > 0 Then
                    Call AddClinObsMod(objClinObs, NCID_PROVIDER_ROLE, objProvider.RoleNcid)
                End If
            End If
            Set objProvider = Nothing
        Next i
    
        'DEM - added this for E & M Coding Review 05/04/00
    '    If EandMReviewed And EandMReviewDate <> NULL_DATE And Val(EandMReviewedBy) > 0 Then
    '        If mobjLocking.LockSection(eDisp, 0, SectionLock) Then
    '            Set objSectionParent = mcolSecParents(CStr(eDisp))
    '            Call objSectionParent.OpenEncounter(Me)
    '            If Not mcolSecParents(CStr(eDisp)).mobjSection.Save Then
    '                Err.Raise vbObjectError, "Encounter.Save", "Error saving E & M Coding Information"
    '            End If
    '            Call mobjLocking.UnlockSection(eDisp, 0)
    '        End If
    '    End If
        'DEM - end add
        lDocumentNumber = 0
        For Each objEncRTF In .PrevSignedRTFs
            Dim objInfo As EventActionInfo
            lDocumentNumber = lDocumentNumber + 1
            Set objClinObs = AddRTF2(objEvent, objEncRTF.body, lDocumentNumber, objEncRTF.OwnerNCID, objEncRTF.LastModifiedDate) 'SCR#42412
            'Set objClinObs = AddRTF(objEvent, objEncRTF.msRTF)
            Set objInfo = objClinObs.ActionsInfo.Add
            objInfo.ActionId.ValueConcept.UniqueId = NCID_VERIFIED
            objInfo.Clinician.ValueConcept.UniqueId = objEncRTF.OwnerNCID
            objInfo.EffectiveTime = objEncRTF.LastModifiedDate
            Set objInfo = Nothing
            Set objClinObs = Nothing
        Next objEncRTF
        
        Call AddClinObs(objEvent, NCID_ENC_TEMPLATE_ID, .TemplateData, , "multiMedia")
        
        On Error Resume Next
        Err.Clear
        bResult = objData.Save
        If Err.Number <> 0 Then
            Set oEmh = New EmhCtrl
            'RAISE ERROR
    '        Err.RAISE 100000, "Encounters DL",  "Error:  " & oEmh.Number & Chr$(13) & Chr$(10) & _
    '                "oEmh.Number = " & oEmh.Number & Chr$(13) & Chr$(10) & _
    '                "oEmh.Entity = " & oEmh.Entity & Chr$(13) & Chr$(10) & _
    '                "oEmh.Orgin = " & oEmh.Origin & Chr$(13) & Chr$(10) & _
    '                "oEmh.ReporterId = " & oEmh.ReporterId & Chr$(13) & Chr$(10) & _
    '                "oEmh.ReporterSubId = " & oEmh.ReporterSubId & Chr$(13) & Chr$(10) & _
    '                "oEmh.Message = " & oEmh.Message & Chr$(13) & Chr$(10) & _
    '                "oEmh.SysMsgString = " & oEmh.SysMsgString & Chr$(13) & Chr$(10) & _
    '                "oEmh.usermsgstring = " & oEmh.UserMsgString
            #If debugon Then
                If Not DL Is Nothing Then DL.DebugLog Exe, Module, "Save", "End", DR, "Failed to Save Enc", oEmh.Message
            #End If
            Set oEmh = Nothing
            iEncDataMapper_Save = False
            'mbDirty = True
        ElseIf bResult Then
            Encounter.DataID = objData.DataID
            If bRaiseEvent Then
                'RaiseEvent when Enc Status is saved SCR-21322
                'ZZZ
                'RaiseEvent EncStatusChanged(.Status)
            End If
            iEncDataMapper_Save = True
            'mbDirty = False
        Else
            'SCR-38260, 39408
            'RASIE ERROR
            'DisplayDetailedError "Encounter.Save", ERROR_MESSAGE_SAVE_FAILURE
            'Err.RAISE 100000, "Encounters DL",  ERROR_MESSAGE_SAVE_FAILURE & vbCrLf & "Encounter.Save(" & PatientID & " - " & FacilityNCID & " - " & EncounterID & ")", vbCritical 'SCR-39485, 38953
            iEncDataMapper_Save = False
            'mbDirty = True
        End If
    
    End With
    GoTo Cleanup
ErrHandler:
    'RAISE ERROR
    'ZZZ
   'Call mobjShared.ShowVBError(Err.Number, Err.Description, "Enctr.Save ", "Enc", vbCritical)
Cleanup:
    Call EndSectionUpdate(eHeader, 0, True)
    #If debugon Then
        If Not DL Is Nothing Then DL.DebugLog Exe, Module, "Save", "End", DR
    #End If
    Set objEvent = Nothing
    Set objClinObs = Nothing
    Set objHeader = Nothing
    Set objData = Nothing
    Set objEncRTF = Nothing
    Set objProvider = Nothing
    Set oEmh = Nothing

End Function



Private Function GetDataStrings(ByVal PatientUnitNumber As String, _
                               ByVal FacilityNCID As String, _
                               ByVal EncounterNumber As String) As Collection
    Dim objDAS As DasCtrl
    Dim colDataStrings As Collection
    
    If Not DebugMode Then On Error GoTo ErrHandler
    
    #If debugon Then
        Dim DR As DebugRec
        If Not DL Is Nothing Then DL.DebugLog Exe, Module, "GetDataStrings", "Begin", DR
    #End If
    
    Set objDAS = New DasCtrl
    objDAS.ClearAll
    objDAS.GeneralCriteria.ClearAll
    objDAS.GeneralCriteria.UnitNumber = PatientUnitNumber
    objDAS.GeneralCriteria.Facility.ValueConcept = FacilityNCID
    objDAS.GeneralCriteria.EncounterNumber = EncounterNumber
    objDAS.GeneralCriteria.QueryQualifier = "EVENT_TYPE_NCID in (" & NCID_ENCEVENT & "," & NCID_DAS_CLINICAL_TEXT & ")"
    objDAS.GeneralCriteria.ReturnMultiMedia = True

    objDAS.DoSearch


    If objDAS.EOF Then
        GoTo Cleanup
    End If
    Set colDataStrings = New Collection
    Do While Not objDAS.EOF
        colDataStrings.Add objDAS.Events
        objDAS.GoToNext
    Loop
    Set GetDataStrings = colDataStrings
    GoTo Cleanup
ErrHandler:
    'err.Raise (Err.Number, Err.Description, "Encounter.GetDataStrings ", "EncounterOps", vbCritical)
Cleanup:
    Set objDAS = Nothing
    Set colDataStrings = Nothing
End Function


Friend Function LoadDataStrings(ByRef colDataStrings As Collection, Optional ByVal Lite As Boolean = False, Optional ByVal StatusOnly As Boolean = False) As DTOs.Encounter
    Dim objData         As DataString
    Dim objEvent        As ClinicalEvent
    Dim objCommentGroup As CommentGroup
    Dim objCommentItem  As CommentItem
    Dim objObs          As ClinicalObservation
    Dim objObsMod       As ClinicalObservation
    Dim objProvider     As EncProvider
    Dim objEncRTF       As Document
    Dim objConcept      As ConceptCtrl
    Dim sName           As String
    Dim sValue          As String
    Dim lDocumentNumber As Long
    Dim i As Long
    Dim j As Long
    Dim Encounter       As Encounter
    Dim mnUpdateCount   As Long

    On Error GoTo ErrHandler
    '
    ' Find the datastring for header information and load it
    'Using the For...Each method with 3M objects causes memory leak
    For i = 1 To colDataStrings.Count
    'For Each objData In colDataStrings
        Set objData = colDataStrings(i)
        If objData.DataTypeNCID = NCID_ENCEVENT Then
            Set objEvent = objData.PatientData
            If objEvent.Header.ObsBatId.Value.Coded.ValueConcept.UniqueId = SectionEnumToNCID(eHeader) Then
                Exit For
            End If
        End If
    Next
    If objData Is Nothing Then
        Err.Raise 100000, "LoadDataStrings", "Encounter data does not contain primary event."
        GoTo Cleanup
    End If
    
    With Encounter
    
        If StatusOnly Then
            .Status = StatusNCIDToEnum(objEvent.Header.TestStatus.Value.CodedWOSform.ValueConcept)
            Exit Function
        End If
        
        Set objConcept = New ConceptCtrl 'Not called in PG mode
        
        .WasUpdated = False
        .WhenSignedDate = 0
        .WhenCosigned = 0
        Set .Cosigner = Nothing
        Set .WhoSigned = Nothing
        Set .WhoCosigned = Nothing
        'Set mcolEncRTFs = New Collection
        Set .Providers = New EncProviders
        lDocumentNumber = 0
        
        .DataID = objData.DataID
        .Id = objData.EncounterNumber
        If .Id = 0 Then
            'MsgBxARMd "Encounter data contains a zero encounter id.", vbExclamation, ExeModule
            GoTo Cleanup
        End If
        .PatientUnitNum = objData.UnitNumber
        .FacilityNCID = objData.EncounterFacility.ValueConcept
        .FacilityName = objData.EncounterFacility.ValueConcept.PrefRep("2000").Representation
        .ClinicNCID = objData.CreatePointOfCare.PhysicalLocation.NursingDivision.ValueConcept
        .ClinicName = objData.CreatePointOfCare.PhysicalLocation.NursingDivision.ValueConcept.PrefRep("2000").Representation
    '   Get PP from Comments SCR-43638
    '    Set objProvider = New EncProvider
    '    If objData.ModifyClinician Is Nothing Then
    '        objProvider.NCID = objData.CreateClinician.ValueConcept.UniqueId
    '        objProvider.FullName = objData.CreateClinician.ValueConcept.PrefRep("2000").Representation
    '    Else
    '        objProvider.NCID = objData.ModifyClinician.ValueConcept.UniqueId
    '        objProvider.FullName = objData.ModifyClinician.ValueConcept.PrefRep("2000").Representation
    '    End If
    '    mcolProviders.Add objProvider, CStr(objProvider.NCID)
    '    Set objProvider = Nothing
        mnUpdateCount = objData.ActionsInfo.Count
        .UpdateDate = objData.ActionsInfo.Item(mnUpdateCount).EnteredTime
        .UpdateClinicianNCID = objData.ActionsInfo.Item(mnUpdateCount).EnteredBy.ValueConcept.UniqueId
    End With
    
    With objEvent.Header
        Encounter.CreatedDate = .ObsDateTime.StartTime.Value.DateTime
        Encounter.WhenSignedDate = .ObsDateTime.EndTime.Value.DateTime
        If CDbl(Encounter.WhenSignedDate) < 1000 Then
            Encounter.WhenSignedDate = 0
        End If
        '.TestStatus.value
        
        Encounter.Status = StatusNCIDToEnum(.TestStatus.Value.CodedWOSform.ValueConcept)
        'Using the For...Each method with 3M objects causes memory leak
        For i = 1 To .Comments.Count
        'For Each objCommentGroup In .Comments
            For j = 1 To .Comments(i).Comments.Count
            'For Each objCommentItem In objCommentGroup.Comments
                Set objCommentItem = .Comments(i).Comments(j)
                Call ParseCommentItem(objCommentItem, sName, sValue)
                Select Case sName
                Case "WASUPDATED"
                    Encounter.WasUpdated = sValue
                Case "COSIGNER"
'jf todo: take this out after a while, as correct way of storing cosigner ncid takes hold
                    Set Encounter.Cosigner = New EncProvider
                    Encounter.Cosigner.NCID = sValue
                    objConcept.UniqueId = sValue
                    Encounter.Cosigner.FullName = objConcept.PrefRep("2000").Representation
                Case "WHENCOSIGNED"
                    Encounter.WhenCosigned = sValue
                Case "WHOCOSIGNED"
                    Set Encounter.WhoCosigned = New EncProvider
                    Encounter.WhoCosigned.NCID = sValue
                    objConcept.UniqueId = sValue
                    Encounter.WhoCosigned.FullName = objConcept.PrefRep("2000").Representation
                Case "COSIGNERLINE2"
                    If Not Encounter.WhoCosigned Is Nothing Then
                        Encounter.WhoCosigned.SigLine2 = sValue
                    End If
                Case "COSIGNERLINE3"
                    If Not Encounter.WhoCosigned Is Nothing Then
                        Encounter.WhoCosigned.SigLine3 = sValue
                    End If
                Case "WHENSIGNED"
                    Encounter.WhenSignedDate = CDate(sValue)
                Case "WHOSIGNED"
                    Set Encounter.WhoSigned = New EncProvider
                    Encounter.WhoSigned.NCID = sValue
                    objConcept.UniqueId = sValue
                    Encounter.WhoSigned.FullName = objConcept.PrefRep("2000").Representation
                Case "SIGNERLINE2"
                    If Not Encounter.WhoSigned Is Nothing Then
                        Encounter.WhoSigned.SigLine2 = sValue
                    End If
                Case "SIGNERLINE3"
                    If Not Encounter.WhoSigned Is Nothing Then
                        Encounter.WhoSigned.SigLine3 = sValue
                    End If
                Case "ENCCAT"
                    'EncounterCategory = sValue
                Case "TIMEZONE"
                    'msTimeZone = sValue
                Case "ENCTYPE"
                    Encounter.ENCTYPE = Val(sValue)
                Case "ENCNOCOUNT"
'                    ec_Followup = 1
'                    ec_New = 2
'                    ec_NoCount = 3
                    Select Case Val(sValue)
                        Case 1 'ec_Followup = 1
                            Encounter.Class = ec_Followup
                        Case 2 'ec_New = 2
                            Encounter.Class = ec_New
                        Case 3 'ec_NoCount = 3
                            Encounter.Class = ec_NoCount
                            '<< SCR 10086
                            Encounter.SADRStatus = NCID_SADR_NO_COUNT
                        Case Else
                    End Select
                Case "EANDMCODEREVIEW"          'DEM - 05/04/00
                    'msEandMCodeReview = sValue
                Case "ALLERGYVERIFIEDBYDATE"
                    Encounter.Sections("RFV").AllergyVerifiedByDate = sValue
                Case "ALLERGYVERIFIEDBYNCID"
                    Encounter.Sections("RFV").AllergyVerifiedByNCID = sValue
                Case "SENSITIVE" 'SCR-35365
                    If StrComp(sValue, "T") = 0 Then
                        'mbSensitivity = True
                        'Encounter.Sensitivity = esrStandard
                    ElseIf IsNumeric(sValue) Then
                        Encounter.Sensitivity = CLng(sValue)
                    End If
                    
                Case "ARRIVEDBYAMBULANCE"
                    If StrComp(sValue, "T") = 0 Then
                        Encounter.Appointment.ArrivedByAmbulance = True
                    End If
                Case "APPTDTS"
                    Encounter.Appointment.DateTime = CDate(sValue)
                Case "UPDATELOCK"
                    Encounter.UpdateLock = CBool(sValue)
                Case "DXSnoID"
                    Encounter.PrimaryDxSnoID = CLng(sValue)
                Case "DXMEDCINPREFIX"
                    'msPrimaryDxMedcinPrefix = sValue
                Case "RESULTEDCONSULTORDER"
                    If StrComp(sValue, "T") = 0 Then
                        Encounter.Appointment.ResultedConsultOrder = True
                    End If
                Case "APSTATUSCOMMENT"  'SCR-17942
                    Encounter.Sections("AP").StatusComment = sValue
                Case "WORKLOADOPTIONS"
                    Encounter.Appointment.WorkloadTypeOption = sValue
                Case "WORKLOAD"
                    Encounter.Appointment.WorkLoad = sValue
                Case "PATIENTSTATUS"
                    'Encounter.Appointment.PatientStatus = sValue
                Case "MEPRSCODES"
                    'Encounter.Appointment.MEPRSCodes = sValue
                Case "APPTIEN"
                    Encounter.Appointment.IEN = sValue
                Case "TEMPLATEID"
                    Encounter.TemplatedID = sValue 'SCR-44706 RBELL 11/3/03 retrieve template id data from db text object
                Case "PRIMPROV"
                    'Modify Clinician is unreliable SCR-43638
                    Set objProvider = New EncProvider
                    objProvider.NCID = sValue
                    objConcept.UniqueId = sValue
                    objProvider.FullName = objConcept.PrefRep("2000").Representation
                    objProvider.RoleNcid = NCID_Attending_Provider      'Default value when loading encounters saved with previous version so old encounters will work with HIPAA
                    Encounter.Providers.Add objProvider, CStr(objProvider.NCID)
                    Set objProvider = Nothing
                Case "HIPAA_CODE" 'SCR#43283
                    'Encounter.ProvHIPAACode = sValue
                End Select
                
                
            Next 'objCommentItem
        Next 'objCommentGroup
    End With
    
    lDocumentNumber = 1
    
    With Encounter
        For i = 1 To objEvent.ClinObs.Count
        'For Each objObs In objEvent.ClinObs
            Set objObs = objEvent.ClinObs(i)
            
            Select Case objObs.ObsId.Value.CodedWOSform.ValueConcept
            Case NCID_NEEDS_COSIGN
                Set .Cosigner = New EncProvider
                .Cosigner.NCID = objObs.ObsValue.Value.CodedWOSform.ValueConcept
                .Cosigner.FullName = objObs.ObsValue.Value.CodedWOSform.ValueConcept.PrefRep("2000").Representation
            Case NCID_ENCOUNTERRTF
    '            If Not Lite Then
                    Dim objInfo As EventActionInfo
                    
                    If objObs.ObsMods.Count > 2 Then
                        'Encounter RTF may be saved in sections 'SCR#42412
                        'added to build 11/4/03
                        'Checks to make sure RTF is stored in this ClinObs
                        If StrComp(objObs.ObsMods(3).ObsValue.Value.InfoType, "basicString", vbTextCompare) = 0 Then
                            If StrComp(objObs.ObsMods(3).ObsValue.Value.BasicString, "RTF", vbTextCompare) = 0 Then
                                'Only need to extract RTF once per document
                                If objObs.ObsMods(1).ObsValue.Value.ULong = lDocumentNumber Then
                                    Set objEncRTF = New Document
                                    objEncRTF.body = ExtractRTF2(objEvent, lDocumentNumber)
    
                                    lDocumentNumber = lDocumentNumber + 1
                                End If
                            End If
                        End If
                    Else
                        'Encounter old way
                        Set objEncRTF = New Document
                        objEncRTF.body = ExtractRTF(objObs)
                    End If
                    
                    If Not objEncRTF Is Nothing Then
                        If objObs.ActionsInfo.Count <> 0 Then
                        Set objInfo = objObs.ActionsInfo.Item(1)
                        objEncRTF.LastModifiedDate = objObs.ActionsInfo.Item(1).EnteredTime
                        objEncRTF.OwnerNCID = objInfo.Clinician.ValueConcept.UniqueId
                        objEncRTF.LastModifiedDate = objInfo.EffectiveTime
                        
                        Set objInfo = Nothing
                        'Encounter.EncRTFs.Add objEncRTF ' let these remain in descending date order
                        End If
                        Set objEncRTF = Nothing
                    End If
    '            End If
            Case NCID_ADDITIONALPROVIDER
                If Len(gobjShared.CmdLineSwitches("HIPAA837")) > 0 Then
                   Set objProvider = New EncProvider
                   'objProvider.CDRLoad objObs
                Else
                   Set objProvider = New EncProvider
                   objProvider.NCID = objObs.ObsValue.Value.CodedWOSform.ValueConcept
                   objProvider.FullName = objObs.ObsValue.Value.CodedWOSform.ValueConcept.PrefRep("2000").Representation
                   For j = 1 To objObs.ObsMods.Count
                   'For Each objObsMod In objObs.ObsMods
                       Set objObsMod = objObs.ObsMods(j)
                       Select Case objObsMod.ObsId.Value.CodedWOSform.ValueConcept
                       Case NCID_PROVIDER_ROLE
                           objProvider.RoleNcid = objObsMod.ObsValue.Value.CodedWOSform.ValueConcept
                           objProvider.Role = objObsMod.ObsValue.Value.CodedWOSform.ValueConcept.PrefRep("2000").Representation
                       End Select
                   Next 'objObsMod
                End If
                
                Dim bAddProvider As Boolean
    
                bAddProvider = True
    
                'Check if provider alredy exists in collection SCR-25124
                For j = 1 To .Providers.Count
                    If .Providers(j).NCID = objProvider.NCID Then
                        bAddProvider = False
                    End If
                Next
                
                If bAddProvider Then
                    .Providers.Add objProvider, CStr(objProvider.NCID)
                ElseIf Encounter.Providers.Count = 1 And Len(gobjShared.CmdLineSwitches("HIPAA837")) > 0 Then
                   'when HIPAA is enabled check to see if this is the same NCID as the primary provider
                   'if it is we should replace the one loaded with the provider retrieved from the new CDR storage location
                   .Providers.Remove 1
                   .Providers.Add objProvider, CStr(objProvider.NCID)
                End If
                
            Case NCID_DIAGNOSIS
                'PrimaryDiagnosisNCID = objObs.ObsValue.Value.CodedWOSform.ValueConcept.UniqueId
            Case NCID_SADR_STATUS
                .SADRStatus = objObs.ObsValue.Value.CodedWOSform.ValueConcept
                            
                If .SADRStatus = NCID_SADR_NO_COUNT Then 'SCR-21981
                    .Class = ec_NoCount
                End If
                
            Case NCID_APPOINTMENT_ID
                .Appointment.Id = objObs.ObsValue.Value.BasicString
            Case NCID_APPOINTMENT_TYPE:
            Case NCID_SECURITY_LABEL:
            Case NCID_ENC_CATEGORY:
            Case NCID_ENC_TEMPLATE_ID:
                .TemplateData = StrConv(objObs.ObsValue.Value.MultiMedia.Data, vbUnicode)
            End Select
        Next 'objObs
    End With
    If Not Lite Then
'        Set mobjLocking = New Locking
'        mobjLocking.msFacilityNCID = FacilityNCID
'        mobjLocking.msEncounterNumber = mnEncounterID
'        mobjLocking.msUserNCID = mobjLogon.UserNCID
'        Call CreateSections(colDataStrings)
'        mbFullyLoaded = True
    End If

    LoadDataStrings = True
    GoTo Cleanup
ErrHandler:
    'Call mobjShared.ShowVBError(Err.Number, Err.Description, "Encounter.Load ", "EncounterOps", vbCritical)
Cleanup:
'    mbLoading = False
    'Resume
 '   mbDirty = False
    Set objData = Nothing
    Set objEvent = Nothing
    Set objCommentGroup = Nothing
    Set objCommentItem = Nothing
    Set objObs = Nothing
    Set objObsMod = Nothing
    Set objProvider = Nothing
    Set objConcept = Nothing
    Set objEncRTF = Nothing
End Function
