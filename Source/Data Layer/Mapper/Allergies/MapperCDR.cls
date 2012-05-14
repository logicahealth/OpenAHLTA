VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MapperCDR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements DL_Interfaces.iMapperAllergies

Private Const MODULE As String = "DM_Allergies.MapperCDR"

Private mConn As CHCSII_CONN.Conn    'CHCSII Connection
Private mCache As CacheMapper        'Core Cache manager

Private Function iMapperAllergies_DeleteAllergy(ByRef DelAllergy As DTOs.Allergy) As Boolean
    'when calling from frmAllergy, set mbDataLoading = true before this call, false afterwards
    'mbDeleteCDR should be set false before call, then true after
    'call AddAllergyTransaction(Synchronize) after deletion
    
    'the follwing will need to be set in the DTO... they previously used globals
    '   .FacilityNCID = cwobjLogonEnviron.FacilityNCID
    '   .EncounterFacilityNCID = frmAllergy.msEncounterFacilityNCID
    '   .Clinician = lblClinician.Caption
    '   .EncounterNumber = frmAllergy.msEncounterNumber
    '   .UnitNumber = glCurPat
    '   .DataID (was parameter to mbDeleteCDR)
    
    Dim objDas              As IDasCtrl         'data access service, used to access the database
    Dim objDatastring       As IDataString      'datastring object, used to access the patient data
    Dim objAllergy          As IAllergyInfo     'allergy info object, used to access the patient's allergy information
    Dim ccTemp             As ICodedCtrl           'catches the return for AllergenIds.Add
    Dim sTemp               As String           'temporary string, used for concatenations
    
    On Error GoTo ErrHandler

    Set objDas = New DasCtrl
    With objDas
        .ClearAll
        .GeneralCriteria.ClearAll
        .GeneralCriteria.UnitNumber = DelAllergy.UnitNumber
        .GeneralCriteria.EventId = DelAllergy.DataId
        .DoSearch

        Set objDatastring = .Events
        Set objAllergy = objDatastring.PatientData
            
        If Not objDatastring Is Nothing Then 'SCR-4134 ADDED
           With objDatastring
                 .Clinician.ValueConcept = DelAllergy.Clinician
                 .EncounterFacility.ValueConcept = DelAllergy.EncounterFacilityNCID
                 .EncounterNumber = DelAllergy.EncounterNumber
                 .Application.ValueConcept.UniqueID = NCID_ALLERGIES_APPLICATION
                 .Facility.ValueConcept.UniqueID = DelAllergy.FacilityNCID
           End With
        Else
           Exit Function
        End If 'SCR-4134 ADDED
        
        sTemp = "<CHCS IEN|DESC>:" & IEN_NKA & "|" & "***DELETE***"
        objAllergy.AllergyType.ValueConcept.UniqueID = gNKA_TYPE
        
        Set ccTemp = objAllergy.AllergenIds.Add
        ccTemp.LText = sTemp
        objAllergy.CompositeSubstance.LText = sTemp
    End With
    
    objDatastring.Save
    
    Set objAllergy = Nothing
    Set objDatastring = Nothing
    Set ccTemp = Nothing
    Set objDas = Nothing
    
    iMapperAllergies_DeleteAllergy = True

    Exit Function

ErrHandler:
    Set objAllergy = Nothing
    Set objDatastring = Nothing
    Set objDas = Nothing
    
    Err.Raise Err.Number, MODULE & ".iMapperCDR_DeleteAllergy", Err.Description
End Function

Private Sub iMapperAllergies_Init(ByRef Support As DL_Support.iParameters, ByRef Cache As DL_Support.CacheMapper)

    On Error GoTo ErrHandler

    'Save objects given to us by DataAccess Factory
    Set mConn = Support.Item("DBConn").Value
    Set mCache = Cache
  
    Exit Sub
  
ErrHandler:
    Err.Raise Err.Number, "MapperGeneric.Init", Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Private Function iMapperAllergies_RetrieveAllergies(ByRef params As DL_Support.iParameters) As DTOs.Allergies
    'these were originally passed in as parameters:
    '                 mbNKA_Rec As Boolean (set to false if any allergies had a value other
    '                                       than 999001 or 999000)
    '                       TODO: write a function to replace this logic
    
    '                 NKA_DataId As String (had no discernable purpose)
    
    'these were globals:
    '   param.PatientUnitNumber =  glCurPat
    '   This.TimeSearchType = miTimeSearchType  (comes from gtDefaultTimePeriod.iTimePeriodType ?)
    
    'set frmAllergy.mbDataLoading = True before calling, false afterwards
    'set frmAllergy.mbUpdatedAllergyInfo = false before calling
    
    'will need to set the following Params from the global time object:
    '   TimeObject_TimeSearchType (this will be -999 if not set, and the other TimeObject parameters will be ignored)
    '   TimeObject_LastN
    '   TimeObject_CalcSlidingTime = gobjTime.CalcSlidingTime(gobjTime(SlidingTime)
    '   TimeObject_BeginTime
    '   TimeObject_EndTime
    '   TimeObject_EncounterNumber
    '   TimeObject_EncounterFacilityNCID

    
    Dim objDas          As IDasCtrl                 'data access service, used to access the database
    Dim objString       As IDataString              'datastring object, used to access the patient data
    Dim objAllergy      As IAllergyInfo             'allergy info object, used to access the patient's allergy information
    Dim objReaction     As IReactions               'reactions object, used to access the reactions for an allergy
    Dim objComments     As ITextComments            'text comments object, used to access the comments for an allergy
    Dim j               As Integer                  'loop variable/counter
    Dim tmpAllergy As DTOs.Allergy                  'Allergy DTO, used while building each allergy
    Dim returnAllergies As DTOs.Allergies           'Allergies DTO, return value of this function
    Dim param As DM_Allergies.AllergySearchParam    'AllergySearchParam, used to typecast the params argument
    Dim tmpReaction As DTOs.AllergyReaction         'AllergyReaction DTO, used while building each Allergy DTO
    
    On Error GoTo ErrHandler
    
    Set param = params
    Set returnAllergies = New DTOs.Allergies
    Set objDas = New DasCtrl
    
    DoAllergySearch objDas, param
    
    Do While Not objDas.EOF
        Set tmpAllergy = New DTOs.Allergy
    
        If (objDas.Events.DataType = gALLERGY_DATA) Then
            Set objString = objDas.Events
                
            With objString
                'Application that stored the string
                tmpAllergy.Application = .Application.ValueConcept.UniqueID
                            
                'POC
                If Not (.ModifyPointOfCare Is Nothing) Then
                    tmpAllergy.PointOfCareFacilityNCID = .ModifyPointOfCare.PhysicalLocation.Facility.ValueConcept
                    tmpAllergy.FacilityNCID = .ModifyPointOfCare.PhysicalLocation.Facility.ValueConcept
                ElseIf Not (.CreatePointOfCare Is Nothing) Then
                    tmpAllergy.PointOfCareFacilityNCID = .CreatePointOfCare.PhysicalLocation.Facility.ValueConcept
                    tmpAllergy.FacilityNCID = .CreatePointOfCare.PhysicalLocation.Facility.ValueConcept
                End If
        
                'EncounterFacility - this is the facility controlling the billing
                tmpAllergy.EncounterFacilityNCID = .EncounterFacility.ValueConcept
                        
                'Encounter Number - billing number
                tmpAllergy.EncounterNumber = .EncounterNumber
                
                'DataId - unique number of the data string
                tmpAllergy.DataId = .DataId
                                
                Set objAllergy = .PatientData
                
            End With 'objString
            
            'Noted date
            tmpAllergy.NotedDate = Format$(objAllergy.IdentificationDate, gsDefaultDate)
            
            'SCR #42771- SL- 10-06-03
            '   formatting with gsDefaultDateTime is throwing an error in Theater cutover... changing to default date
            tmpAllergy.NotedDate = Format$(objAllergy.IdentificationDate, gsDefaultDate)
            
            SetAllergyType objAllergy, tmpAllergy
            
            
            
            If tmpAllergy.Allergen = "***DELETE***" Then
                GoTo NextDasRecord
            ElseIf Err <> 0 Then
                GoTo ErrHandler
            Else
                On Error GoTo ErrHandler
                
                'Allergy type
                If Not objAllergy.AllergyType.IsEmpty Then
                    tmpAllergy.AllergyTypeNCID = objAllergy.AllergyType.ValueConcept
                    Select Case tmpAllergy.AllergyTypeNCID
                        Case gNKA_TYPE
                            If objString.DeleteFlag = 0 Then
                               tmpAllergy.AllergyType = "NKA"
                            End If
                        Case Else
                            If tmpAllergy.AllergenNCID = IEN_NKA Or _
                                tmpAllergy.AllergenNCID = IEN_NKA_2 Or _
                                tmpAllergy.AllergenNCID = IEN_NKDA Or _
                                tmpAllergy.AllergenNCID = IEN_NKDA_2 Or _
                                tmpAllergy.AllergenNCID = IEN_NKA_NONE Then 'ADDED SCR-4985, SCR-9738
                                
                                 'gSelectedAllergen = "***DELETE***"
                                 'gnCurrentAllergyIndex = iTotalStrings
                                 'gSelectedAllergenNCID = tmpAllergy.AllergenNCID
                                 'mbDoNotRefresh = True
                                 
                                 '--- mbSaveStartedElsewhere - ARM SCR 19546 MSolano 1/7/02
                                 '--- Delete's are save's w/ a delete flag - don't count as a save
                                 
                                 
                                 'TODO: need to figure out what was happening here
                                 '      (in frmAllergy)
                                 'mbSaveStartedElsewhere = True
                                 'Call cmdSave_Click
                                 'mbSaveStartedElsewhere = False
                                 
                                 'TODO: moving this down... add the NKA record at the end
                                 '      is this really necessary?
                                 'If Not CheckAllergies("NKA") Then
                                 '   Call add_nka_record
                                 '   mbAddNKARecord = False
                                 '
                                 '   'Check to make sure nka record is added, if not try again outside this procedure
                                 '   If Not CheckAllergies("NKA") Then
                                 '       mbAddNKARecord = True
                                 '   End If
                                 'End If
                             End If
                    End Select
                Else
                    tmpAllergy.AllergyTypeNCID = vbNullString
                    tmpAllergy.AllergyType = vbNullString
                End If
            End If
            On Error GoTo ErrHandler
            
            'Reactions
            Set objReaction = objAllergy.Reactions

            'Info source
            If Not objAllergy.SourceOfInfo.IsEmpty Then
                tmpAllergy.InfoSourceNCID = objAllergy.SourceOfInfo.ValueConcept
            End If

            'use the acquisition of the clinician name to invoke the lookup of the
            'rest of the preferred representations.  using the clinician should avoid
            'two if then statements.
            
            CopyClinicanAndPointOfCare objString, tmpAllergy
            
            
            tmpAllergy.EncounterFacility = _
               objString.EncounterFacility.Value(NCID_3M_DEFAULT_CONTEXT)

            
            With objReaction
                '---make room to store the reactions---
                If .count > 0 Then
                    For j = 1 To .count
                        Set tmpReaction = New DTOs.AllergyReaction
                        If .Item(j).IsText Then
                            tmpReaction.NCID = vbNullString
                            tmpReaction.Name = "Other: " & .Item(j).Value(NCID_3M_DEFAULT_CONTEXT)
                        Else
                            tmpReaction.NCID = .Item(j).ValueConcept
                            tmpReaction.Name = .Item(j).Value(NCID_3M_DEFAULT_CONTEXT)
                        End If
                        tmpAllergy.AllergyReactions.Add tmpReaction
                    Next j
                End If
            End With
            
            tmpAllergy.InfoSource = objAllergy.SourceOfInfo.Value(NCID_3M_DEFAULT_CONTEXT)
            
            'Comments
            If Not objAllergy.TextComments.IsEmpty Then
               SetComments objAllergy.TextComments, tmpAllergy
            Else
                tmpAllergy.Comment = vbNullString
            End If
            
            'Set the unit number
            tmpAllergy.UnitNumber = params.Item("patientunitnumber").Value
            
            returnAllergies.Add tmpAllergy

            Set objComments = Nothing
            Set objReaction = Nothing
            Set objAllergy = Nothing
            Set objString = Nothing
            
NextDasRecord:
        End If
        
        objDas.GoToNext
    Loop
    
    Set iMapperAllergies_RetrieveAllergies = returnAllergies

    Set objDas = Nothing
    
CleanupAndExit:
    On Error Resume Next
    Set objComments = Nothing
    Set objReaction = Nothing
    Set objAllergy = Nothing
    Set objString = Nothing
    Set objDas = Nothing
    Exit Function

ErrHandler:
    Err.Raise Err.Number, MODULE & ".iMapperAllergies_RetrieveAllergies", Err.Description
    
End Function

Private Function iMapperAllergies_RetrieveAllergens(ByRef params As DL_Support.iParameters) As DTOs.Allergens
Dim SQLHelper As AllergiesSQL

    Set SQLHelper = New AllergiesSQL
    SQLHelper.Init mConn, CDR
    
    Set iMapperAllergies_RetrieveAllergens = SQLHelper.RetrieveAllergens(params)
    Set SQLHelper = Nothing

End Function

Private Function iMapperAllergies_SaveAllergy(ByRef SavAllergy As DTOs.Allergy) As Boolean

    'set .DataID to 0 if adding a new allergy
    
    'when called from frmAllergy, set mbDataLoading = false before call,
    '   then true after, followed by frmAllergy.AddAllergyTransaction(Synchronize)
    
    
    'the following were handled with globals, and will now need to be set before calling save
    '   .UnitNumber = glCurPat
    '   .EncounterNumber = objEncounter.CurrentEncounter.EncounterID
    '   .FacilityNCID = objEncounter.CurrentEncounter.FacilityNCID
    '                   or = .PointOfCareFacilityNCID if .FacilityNCID is empty
    '   .Allergen (?) = gSelectedAllergen
    '           when TransferToCDR is true, the .Allergen value will change on return from this function
    '           and should be re-stored in gSelectedAllergen
        
    'Transfer2CDR global is gone... put "<CDR>" in the front of the .Allergen field to trigger the
    '   same code... that designator will be removed in this function
    
    '.AllergyReaction.Name should be limited to 60 characters outside of this function
    '       (used to be implemented, with message, here)
    
    Dim objDas            As IDasCtrl           'data access service, used to access the database
    Dim objDatastring     As IDataString        'datastring object, used to access the patient data
    Dim ccTemp           As ICodedCtrl             'used to catch the return from AllergenIds.Add
    Dim objAllergy        As IAllergyInfo       'allergy object, used to write data to CDR
    Dim objDomain         As IDomainToRange     'domaintorange object
    Dim i                 As Long               'loop variable/counter
    Dim sEncounter        As String             'encounter number
    Dim sFacility         As String             'facility ncid
    
    On Error GoTo ErrHandler
    
    If SavAllergy.DataId = 0 Then 'this is a new string
            
        Set objDatastring = New DataString
        With objDatastring
             .UnitNumber = SavAllergy.UnitNumber
             .DataType = gALLERGY_DATA
             Set objAllergy = .PatientData
        End With
        
        If (SavAllergy.FacilityNCID = 0) Or (SavAllergy.FacilityNCID = vbNullString) Then
            SavAllergy.FacilityNCID = SavAllergy.PointOfCareFacilityNCID
        End If
    
    Else
        Set objDas = New DasCtrl
        
        With objDas
            .ClearAll
            
            .GeneralCriteria.ClearAll
            .GeneralCriteria.UnitNumber = SavAllergy.UnitNumber
            .GeneralCriteria.EventId = SavAllergy.DataId
            
            .DoSearch
        
            If Not .EOF Then
                 Set objAllergy = .Events.PatientData
                 Set objDatastring = .Events
                 
                 'Remove collections, add back as needed, then simply reassign over existing values
                 With objAllergy
                     
                     
                    'clear AllergenIds, Reactions and Comments
                     With .AllergenIds
                         For i = .count To 1 Step -1
                             .Remove i
                         Next i
                     End With '.AllergenIds
                     With .Reactions
                         For i = .count To 1 Step -1
                             .Remove i
                         Next i
                     End With
                     With .TextComments
                         For i = .count To 1 Step -1
                             .Remove i
                         Next i
                     End With '.TextComments
                 End With 'objAllergy
                 
                 sEncounter = .Events.EncounterNumber
                 sFacility = .Events.Facility.ValueConcept.UniqueID
            
            Else 'SCR-4134 ADDED
                
                GoTo CleanupAndExit
            End If
        End With 'objDas
    End If
    
    'must set every time for a save
    If Not objDatastring Is Nothing Then 'SCR-4134 ADDED
    
        With objDatastring
            .Clinician.ValueConcept = SavAllergy.ClinicianNCID
            
            'SCR #52409- clinician name was not being stored
            .Clinician.LText = SavAllergy.Clinician
            
            .EncounterFacility.ValueConcept = sFacility 'SavAllergy.EncounterFacilityNCID
            .EncounterNumber = sEncounter 'SavAllergy.EncounterNumber
            With .PointOfCare.PhysicalLocation
                .InfoType = "facility"
                .Facility.ValueConcept = sFacility
                    'where sFacility = SavAllergy.PointOfCareFacilityNCID
            End With
            
            .Application.ValueConcept.UniqueID = NCID_ALLERGIES_APPLICATION
            
            'fix?
            .Facility.ValueConcept.UniqueID = SavAllergy.FacilityNCID
        End With
        
    End If 'SCR-4134 ADDED
    
    SetAllergyObjectInfo objAllergy, SavAllergy
    

    Set ccTemp = Nothing
    
    '---store reactions---
    
    
    StoreReactions SavAllergy, objAllergy
    StoreCommentsCHCSIDFacilities SavAllergy, objAllergy
    
    
    DebugPrintAllergy SavAllergy
    
    '---save/update the string---
    If objDatastring.Save Then 'ADDED SCR-4134
        SavAllergy.DataId = objDatastring.DataId 'SCR-4134 ADDED
        
        'TODO: AddAllergyTransaction???
    
    End If
        
    If Err Then
        GoTo ErrHandler
    End If
    
    #If DebugOn Then
        DL.DebugLog Exe, MODULE, _
                    "mSaveCDR", "P-After mSaveCDR", DR
    #End If
    GoTo CleanupAndExit
    
ErrHandler:
    'Resume Next
    
    Err.Raise Err.Number, MODULE & ".iMapperAllergies_SaveAllerrgy", Err.Description

CleanupAndExit:
    '-----------------------------------------------------------
    '---release reference to objects---
    '-----------------------------------------------------------
    'Resume Next
    Set objDomain = Nothing
    Set ccTemp = Nothing
    Set objAllergy = Nothing
    Set objDatastring = Nothing
    Set objDas = Nothing
    
    'mbDataLoading = False   'SCR-6328 Used to prevent Checkbox chk_VerifiedThisEncounter"'s Click() event-handler from executing code.
End Function

Private Function iMapperAllergies_RetrieveAllergyReactions(ByRef params As iParameters) As DTOs.AllergyReactions
    On Error GoTo ErrHandler
    
    Dim bMaskSearch As Boolean                  'boolean, true if search keyword contains a mask (%)
    Dim sMask As String                         'mask for searching for reactions
    Dim iIndex As Long                          'loop variable/counter
    Dim tmpReaction As DTOs.AllergyReaction     'AllergyReaction DTO used to build AllergyReactions DTO
    Dim param As ReactionSearchParams           'object used to typecast incoming param argument
    Dim result As AllergyReactions              'AllergyReactions DTO returned from this function
    
    Set param = params
    
    Set result = New AllergyReactions
    
    Dim oPickListCtrl As PickListCtrl
    Set oPickListCtrl = New PickListCtrl
    
    oPickListCtrl.Domain = gREACTION_DOMAIN
    oPickListCtrl.Load
    
    If InStr(param.PrefixKeyword, "%") <> 0 Then
        bMaskSearch = True
        sMask = Left$(param.PrefixKeyword, InStr(param.PrefixKeyword, "%") - 1)
    Else
        sMask = param.PrefixKeyword
    End If
    
    sMask = UCase$(Trim$(sMask))
    
    For iIndex = 1 To oPickListCtrl.ListItemCnt
        With oPickListCtrl.ListItem(iIndex)
            If bMaskSearch Then
                If (InStr(UCase$(.Representation), sMask) = 1) Then
                    Set tmpReaction = New DTOs.AllergyReaction
                    tmpReaction.NCID = .Concept
                    tmpReaction.Name = UCase$(.Representation)
                    result.Add tmpReaction
                End If
            Else
                If (UCase$(.Representation) = sMask) Or (sMask = vbNullString) Then
                    Set tmpReaction = New DTOs.AllergyReaction
                    tmpReaction.NCID = .Concept
                    tmpReaction.Name = UCase$(.Representation)
                    result.Add tmpReaction
                End If
            End If
        End With
    Next
                
    Set iMapperAllergies_RetrieveAllergyReactions = result
                
    GoTo CleanUp
    
ErrHandler:
    Err.Raise Err.Number, MODULE & ".iMapperAllergies_RetrieveAllergyReactions", Err.Description

CleanUp:
    On Error Resume Next
    Set oPickListCtrl = Nothing
End Function


Public Sub DebugPrintAllergy(ByRef dtoAllergy As DTOs.Allergy)
    'this code isn't used; but, can be very helpful in debugging
    On Error GoTo ErrHandler

    Debug.Print "Application: " & dtoAllergy.Application    '
    Debug.Print "NotedDate: " & dtoAllergy.NotedDate      '
    Debug.Print "Clinician: " & dtoAllergy.Clinician      '
    Debug.Print "ClinicianNCID: " & dtoAllergy.ClinicianNCID      '
    Debug.Print "PointOfCareFacility: " & dtoAllergy.PointOfCareFacility    '
    Debug.Print "PointOfCareFacilityNCID: " & dtoAllergy.PointOfCareFacilityNCID    '
    Debug.Print "EncounterFacility: " & dtoAllergy.EncounterFacility      '
    Debug.Print "EncounterFacilityNCID: " & dtoAllergy.EncounterFacilityNCID      '
    Debug.Print "EncounterNumber: " & dtoAllergy.EncounterNumber    '
    Debug.Print "AllergyType: " & dtoAllergy.AllergyType    '
    Debug.Print "AllergyTypeNCID: " & dtoAllergy.AllergyTypeNCID    '
    Debug.Print "Allergen: " & dtoAllergy.Allergen   '
    Debug.Print "AllergenNCID: " & dtoAllergy.AllergenNCID   '
    Debug.Print "AllergyReactions: " & "<OBJECT>"
    Debug.Print "InfoSource: " & dtoAllergy.InfoSource     '
    Debug.Print "InfoSourceNCID: " & dtoAllergy.InfoSourceNCID     '
    Debug.Print "Comment: " & dtoAllergy.Comment    '
    Debug.Print "CreateUser: " & dtoAllergy.CreateUser     '
    Debug.Print "CreateUserNCID: " & dtoAllergy.CreateUserNCID     '
    Debug.Print "CreateTime: " & dtoAllergy.CreateTime     '
    Debug.Print "Modified: " & CStr(dtoAllergy.Modified)
    Debug.Print "ModifyUser: " & dtoAllergy.ModifyUser     '
    Debug.Print "ModifyUserNCID: " & dtoAllergy.ModifyUserNCID     '
    Debug.Print "ModifyTime: " & dtoAllergy.ModifyTime     '
    Debug.Print "Facility: " & dtoAllergy.Facility   '
    Debug.Print "FacilityNCID: " & dtoAllergy.FacilityNCID   '
    Debug.Print "DataID: " & dtoAllergy.DataId     '
    Debug.Print "NewRec: " & dtoAllergy.NewRec
    Debug.Print "ModsFlg: " & dtoAllergy.ModsFlg
    Debug.Print "CHCS_ID: " & dtoAllergy.CHCS_ID    '
    Debug.Print "CHCSII_ENTRY: " & dtoAllergy.CHCSII_ENTRY   '
    Debug.Print "CHCS_FACILITIES: " & "<ARRAY>"
    Debug.Print "CDRDataID: " & dtoAllergy.CDRDataID      '
    Debug.Print "UnitNumber: " & dtoAllergy.UnitNumber     '
    Debug.Print "UpdateFlag: " & dtoAllergy.UpDateFlag     '
    Debug.Print "Tmip_Status: " & dtoAllergy.Tmip_Status    '
    
ErrHandler:

End Sub


Private Function iMapperAllergies_SaveAllergyTransaction(ByRef AllergyTrans As DTOs.AllergyTransaction) As Boolean
    iMapperAllergies_SaveAllergyTransaction = SaveAllergyTransaction_SP(AllergyTrans.UnitNumber)
End Function

Private Function iMapperAllergies_RetrieveAllergyTransactions(ByRef param As iParameters) As DTOs.AllergyTransactions
    On Error GoTo ErrHandler
    
    'note: if param.UnitNumber is specified, this performs a join with the MMI_Site table, and
    '   returns fewer records.  If UnitNumber is not specified, the .IENNumber must be.
    
    Dim dtoAllergyTransaction As DTOs.AllergyTransaction    'AllergyTransaction DTO used to build the return AllergyTransactions DTO
    Dim dtoAllergyTransactions As DTOs.AllergyTransactions  'AllergyTransactions DTO to be returned from this function
    Dim params As DM_Allergies.TransactionSearchParams      'TransactionSearchParams to typecast param argument
    Dim sTempDate As String                                 'string used to rebuild the returned transaction date
    Dim rs As ADODB.Recordset                               'recordset to catch query results
    
    Set params = param
    
    Set rs = RetrieveAllergyTransactions_SP(params.UnitNumber, CLng(params.TransType))
    
    'Dim sSQL As String
    'If params.IncludeMMIData Then
    '    sSQL = "SELECT at.allergy_trans_id, ms.medical_record_number, ms.facility_cid, at.update_status, at.trans_date " _
    '            & " FROM mmi_site ms, allergy_transaction at" _
    '            & " WHERE ms.unit_number = " & params.UnitNumber _
    '            & " AND to_number(ms.medical_record_number) = at.patient_ien(+)" _
    '            & " AND ms.facility_cid = at.facility_ncid(+)" _
    '            & " AND " & params.TransType & " = at.trans_type(+)"
    '
    '    If params.Synchronize Then
    '        sSQL = sSQL & " and 0 = at.update_status(+) "
    '    End If
    '
    ' Else
    '     sSQL = "Select allergy_trans_id, update_status, trans_date, facility_ncid, patient_ien from allergy_transaction where " _
    '             & params.TransType & " = trans_type and unit_number = " & params.UnitNumber
    'End If
    
    'Set rs = mConn.CHCSII_DAS.OpenRecordsetSP(sSQL)

    
    Set dtoAllergyTransactions = New DTOs.AllergyTransactions
    Do While Not rs.EOF
        Set dtoAllergyTransaction = New DTOs.AllergyTransaction
        
        With dtoAllergyTransaction
            
            'SCRs 50392, 50394 - unable to recreate the problem; but, this value can come back null
            '   in the performance optimization environment
            
            If (IsNull(rs.Fields("allergy_trans_id"))) Then
                .AllergyTransID = 0
            Else
                .AllergyTransID = rs.Fields("allergy_Trans_Id")
            End If
                
            .PatientIEN = rs.Fields("patient_ien")
            .FacilityNCID = rs.Fields("facility_ncid")
            
            .UpdateStatus = rs.Fields("update_status")
            
            sTempDate = rs.Fields("trans_date")
            If sTempDate <> vbNullString Then
                sTempDate = Mid$(sTempDate, 9, 2) & ":" & Mid$(sTempDate, 11, 2) & ":" & Mid$(sTempDate, 13, 2) & " " & Mid$(sTempDate, 5, 2) & "/" & Mid$(sTempDate, 7, 2) & "/" & Mid$(sTempDate, 1, 4)
                .TransDate = CDate(sTempDate)
            Else
                .TransDate = 0
            End If
            
        End With
        dtoAllergyTransactions.Add dtoAllergyTransaction
        
        rs.MoveNext
    Loop
    
    Set iMapperAllergies_RetrieveAllergyTransactions = dtoAllergyTransactions
    
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, MODULE & ".iMapperAllergies_RetrieveAllergyTransactions", Err.Description

End Function


Private Function iMapperAllergies_RecordsetTransferToLDDB(ByRef col As Collection) As Long
    Err.Raise 12345, MODULE & ".iMapperAllergies_RetrieveAllergyTransactions", "This method has not been implemented in this data mode."
End Function

Private Function iMapperAllergies_SaveAllergies(ByRef SavAllergies As DTOs.Allergies) As Boolean
    Err.Raise 12345, MODULE & ".iMapperAllergies_SaveAllergies", "This method has not been implemented in this data mode."
End Function

Private Function SaveAllergyTransaction_SP(ByVal sPatientUnitNum As String) As Boolean
    Dim objCmd As CHCSII_Command.CHCSII_Cmd         'CHCSII_Cmd object to call stored procedure

    On Error GoTo ErrHandler

    Set objCmd = New CHCSII_Command.CHCSII_Cmd
    
    objCmd.SPName = "MT_ALLERGY_PKG" + ".PRC_REQUEST_ALLERGY_SYNC_CW"
    
    objCmd.AddParam False, False, "textString", sPatientUnitNum
    objCmd.AddParam True, True, vbNullString, vbNullString
    
    mConn.CHCSII_DAS(CDR).ExecuteSP objCmd
    
    SaveAllergyTransaction_SP = True
CleanUp:
    Set objCmd = Nothing

    Exit Function
ErrHandler:
    SaveAllergyTransaction_SP = False
    Err.Raise Err.Number, MODULE & ".AddAllergyTransaction_SP", Err.Description
    GoTo CleanUp
End Function

Private Function RetrieveAllergyTransactions_SP(ByVal sPatientUnitNum As String, ByVal lTransType As Long) As ADODB.Recordset
    Dim objCmd As CHCSII_Command.CHCSII_Cmd         'CHCSII_Cmd object to call stored procedure

    On Error GoTo ErrHandler

    Set objCmd = New CHCSII_Command.CHCSII_Cmd
    
    objCmd.SPName = "MT_ALLERGY_PKG.PRC_REFRESH_ALLERGY_TRANS"
    
    objCmd.AddParam False, False, "textString", sPatientUnitNum
    objCmd.AddParam False, False, "textString", lTransType
    objCmd.AddParam True, True, vbNullString, vbNullString
    
    Set RetrieveAllergyTransactions_SP = mConn.CHCSII_DAS(CDR).OpenRecordsetSP(objCmd)
CleanUp:
    Set objCmd = Nothing

    Exit Function
ErrHandler:
    Err.Raise Err.Number, MODULE & ".RetrieveAllergyTransactions_SP", Err.Description
    GoTo CleanUp
End Function

Private Sub SetComments(ByRef objComments As ITextComments, ByRef tmpAllergy As DTOs.Allergy)
    Dim j As Long
    Dim k As Long
    Dim iArrayCount As Long
    
    On Error GoTo ErrHandler

    tmpAllergy.Comment = vbNullString
    For j = 1 To objComments.count
        With objComments.Item(j).Comments
            For k = 1 To .count
                If Left$(.Item(k).Text, 18) = "ENTERED IN CHCSII:" Then
                    tmpAllergy.CHCSII_ENTRY = Right$(.Item(k).Text, 1)
                    If tmpAllergy.CHCSII_ENTRY = "N" Then
                        'Handle Theater Providers 'SCR-25768
                        If Left$(tmpAllergy.EncounterFacility, 7) = "THEATER" Then
                            tmpAllergy.Clinician = "(Unknown Provider)"
                        Else
                            tmpAllergy.Clinician = "CHCS"
                        End If
                    End If
                ElseIf Left$(.Item(k).Text, 9) = "FACILITY:" Then
                    iArrayCount = tmpAllergy.CHCS_FACILITIES_Count
                    tmpAllergy.CHCS_FACILITIES(CLng(iArrayCount)) = Right$(.Item(k).Text, Len(.Item(k).Text) - 9)
                ElseIf Left$(.Item(k).Text, 8) <> "CHCS ID:" Then
                    tmpAllergy.Comment = tmpAllergy.Comment & .Item(k).Text
                Else
                    tmpAllergy.CHCS_ID = .Item(k).Text
                End If
            Next k
        End With
    Next j
    
    Exit Sub

ErrHandler:
    Err.Raise Err.Number, MODULE & ".SetComments", Err.Description

End Sub


Private Sub StoreReactions(ByRef SavAllergy As DTOs.Allergy, ByRef objAllergy As IAllergyInfo)
    Dim sTemp As String
    Dim ccTemp As ICodedCtrl
    Dim i As Long
    Dim sReactionNCID As String
    
    On Error GoTo ErrHandler

    With SavAllergy
        If InStr(1, .Allergen, "DELETE", 1) = 0 Then 'ADDED SCR-4134
            If .AllergyReactions.count > 0 Then
                For i = 1 To .AllergyReactions.count
                    sReactionNCID = .AllergyReactions.Item(i).NCID
                    
                    Set ccTemp = objAllergy.Reactions.Add
                    
                    If Len(sReactionNCID) > 0 Then      '--coded--
                        ccTemp.ValueConcept.UniqueID = sReactionNCID$
                    Else    '--free-text--
                        sTemp = .AllergyReactions.Item(i).Name
                        
                        'Reaction
                        If InStr(1, sTemp, "Other: ") > 0 Then
                           '---parse out "Other: "---
                           sTemp = Right$(sTemp, Len(sTemp) - 7)
                        End If
                        
                        'limit reaction to 60 chars
                        If (Len(sTemp) > 60) Then
                            sTemp = Left$(sTemp, 60)
                            .AllergyReactions.Item(i).Name = sTemp
                        End If
                        
                        ccTemp.LText = sTemp
                    End If
                Next
            End If
       End If
    End With 'SavAllergy
    
    Set ccTemp = Nothing
    
    Exit Sub
    
ErrHandler:
    Err.Raise Err.Number, MODULE & ".StoreReactions", Err.Description

End Sub

Private Sub SetAllergyType(ByRef objAllergy As IAllergyInfo, ByRef tmpAllergy As DTOs.Allergy)
    Dim sTokens() As String
    Dim sSubTokens() As String
    Dim i As Integer
    
    On Error GoTo ErrHandler
    
    With objAllergy.CompositeSubstance
        If Not .IsEmpty Then
            If .IsText Then
                tmpAllergy.AllergenNCID = "Text"
                If InStr(1, .LText, "<CHCS IEN|DESC>:") = 1 Then
                    sTokens = Split(.LText, ":")
                    If (sTokens(1) = "Other") Then
                        sTokens = Split(sTokens(2), "|")
                    Else
                        sTokens = Split(sTokens(1), "|")
                    End If
                    If UBound(sTokens) = 0 Then
                        'This should not happen, but we will
                        'assume if it does that this is the Allergen.
                        '--------------------------------------------
                        tmpAllergy.Allergen = sTokens(0)
                    ElseIf UBound(sTokens) = 1 Then
                        tmpAllergy.Allergen = sTokens(1)
                        tmpAllergy.AllergenNCID = sTokens(0)
                    Else
                        'This should not happen either, but if it
                        'does we will reassemble tokens 1 through n.
                        '-----------------------------------
                        ReDim sSubTokens(0 To UBound(sTokens) - 1)
                        For i = 1 To UBound(sTokens)
                            sSubTokens(i - 1) = sTokens(i)
                        Next
                        tmpAllergy.Allergen = Join(sSubTokens, "|")
                        tmpAllergy.AllergenNCID = sTokens(0)
                    End If
                Else
                    If objAllergy.AllergenIds.count = 0 Then
                       tmpAllergy.Allergen = "Other: " & .LText
                    Else
                       tmpAllergy.Allergen = _
                          "Other: " & objAllergy.AllergenIds(1).LText
                    End If
                End If
            Else
                tmpAllergy.AllergenNCID = .ValueConcept
                tmpAllergy.Allergen = .Value(NCID_3M_DEFAULT_CONTEXT)
            End If
        End If
    
    End With 'objAllergy.CompositeSubstance
    
    Exit Sub

ErrHandler:
    Err.Raise Err.Number, MODULE & ".SetAllergyType", Err.Description

End Sub

Private Sub SetAllergyObjectInfo(ByRef objAllergy As IAllergyInfo, ByRef SavAllergy As DTOs.Allergy)
    Dim ccTemp As ICodedCtrl
    Dim objDomain As DomainToRange
    Dim i As Long

    On Error GoTo ErrHandler

    With objAllergy
        '---fill in properties of the object---
        If Not SavAllergy.NotedDate = vbNullString Then
            If Format$(SavAllergy.NotedDate, gsDefaultDate) <> SavAllergy.NotedDate Then
                 .IdentificationDate = CDate(Left$(SavAllergy.NotedDate, InStrRev(SavAllergy.NotedDate, " ") - 1))
            Else
                .IdentificationDate = SavAllergy.NotedDate
            End If
        Else
           .IdentificationDate = Now
        End If
        
        .AllergyType.ValueConcept.UniqueID = SavAllergy.AllergyTypeNCID
        
        If SavAllergy.AllergenNCID = "Text" Then
            .CompositeSubstance.LText = SavAllergy.Allergen
            
            Set ccTemp = .AllergenIds.Add
            ccTemp.LText = SavAllergy.Allergen
        ElseIf Left$(SavAllergy.Allergen, 5) = "<CDR>" Then
            'Format to CDR SCR-17652
            
            'hacking this to replace what was once the Transfer2CDR global variable (in frmAllergy)
            '   put "<CDR>" in front of the allergen to trigger this
        
            SavAllergy.Allergen = Right$(SavAllergy.Allergen, Len(SavAllergy.Allergen) - 5)
            SavAllergy.Allergen = "<CHCS IEN|DESC>:" & SavAllergy.AllergenNCID & "|" & SavAllergy.Allergen
            .CompositeSubstance.LText = SavAllergy.Allergen
            .AllergenIds.Add.LText = SavAllergy.Allergen
        Else
            .CompositeSubstance.ValueConcept.UniqueID = SavAllergy.AllergenNCID
            
            'now store out the stuff for alert group
            If SavAllergy.AllergyTypeNCID = gDRUG_TYPE Then
                Set objDomain = New DomainToRange
                
                If objDomain Is Nothing Then
                    Err.Raise 123456, "MapperCDR.iMapperAllergies_SaveAllergy", "Failed to obtain a new DomainToRange object."
                Else
                    With objDomain
                        .Domain = SavAllergy.AllergenNCID
                        .Relationship = gHAS_INGREDIENT
                        .Load
                    End With 'objDomain

                    If objDomain.RangeItemCnt > 0 Then
                        For i = 1 To objDomain.RangeItemCnt
                            Set ccTemp = .AllergenIds.Add
                            ccTemp.ValueConcept.UniqueID = objDomain.RangeItem(i)
                        Next i
                    Else
                        Set ccTemp = .AllergenIds.Add
                        ccTemp.ValueConcept.UniqueID = SavAllergy.AllergenNCID
                    End If
                    
                    Set objDomain = Nothing
                End If
                
            Else    'non drug ncid
                Set ccTemp = .AllergenIds.Add
                ccTemp.ValueConcept.UniqueID = SavAllergy.AllergenNCID
            End If
            
        End If
        
        'TRUNCATE COMPOSITE SUBSTANCE TEXT TO 80 CHARS:
        With .CompositeSubstance
            If Len(.LText) > 80 Then
               .LText = Left$(.LText, 80)
            End If
        End With '.CompositeSubstance
        
    End With 'objAllergy
    
    Exit Sub
    
ErrHandler:
    Err.Raise Err.Number, MODULE & ".SetAllergyObjectInfo", Err.Description
End Sub

Public Sub DoAllergySearch(ByRef objDas As IDasCtrl, ByRef param As DM_Allergies.AllergySearchParam)
    On Error GoTo ErrHandler

    If Not objDas Is Nothing Then
        objDas.ClearAll
        objDas.GeneralCriteria.ClearAll
        objDas.GeneralCriteria.UnitNumber = param.PatientUnitNumber
        
        objDas.GeneralCriteria.EventType.ValueConcept = NCID_PATIENT_ALLERGIES
        objDas.GeneralCriteria.QueryName.ValueConcept = 69409
    
        '---set corresponding "time" period---
        Select Case param.TimeObject_TimeSearchType
            Case cwiLASTN_TIMES
                objDas.GeneralCriteria.QueryCount = param.TimeObject_LastN
            Case cwiSLIDING_TIMES
                objDas.GeneralCriteria.EventEndGMTime = Format$(Now, "General Date")
                objDas.GeneralCriteria.EventStartGMTime = Format$(param.TimeObject_CalcSlidingTime, "General Date")
            Case cwiFIXED_TIMES
                objDas.GeneralCriteria.EventStartGMTime = Format$(param.TimeObject_BeginTime, "General Date")
                objDas.GeneralCriteria.EventEndGMTime = Format$(param.TimeObject_EndTime, "General Date")
            Case cwiENCOUNTER_NUMBER_TIMES
                objDas.GeneralCriteria.EncounterNumber = param.TimeObject_EncounterNumber
                objDas.GeneralCriteria.Facility.ValueConcept.UniqueID = param.TimeObject_EncounterFacilityNCID
            Case Else   '--do nothing--
        End Select
        
        '---build the data---
        objDas.DoSearch
    End If
    
    Exit Sub

ErrHandler:
    Err.Raise Err.Number, MODULE & ".DoAllergySearch", Err.Description

End Sub


Public Sub CopyClinicanAndPointOfCare(ByRef objString As DataString, ByRef tmpAllergy As DTOs.Allergy)
    On Error GoTo ErrHandler
    
    'Clinician
    If Not (objString.ModifyClinician Is Nothing) Then
        tmpAllergy.ClinicianNCID = objString.ModifyClinician.ValueConcept
        tmpAllergy.Clinician = _
           objString.ModifyClinician.Value(NCID_USER_FULL_NAME_CONTEXT)
    ElseIf Not (objString.CreateClinician Is Nothing) Then
        tmpAllergy.ClinicianNCID = objString.CreateClinician.ValueConcept
        tmpAllergy.Clinician = _
           objString.CreateClinician.Value(NCID_USER_FULL_NAME_CONTEXT)
    End If

    'POC
    If Not (objString.ModifyPointOfCare Is Nothing) Then
        tmpAllergy.PointOfCareFacility = _
           objString.ModifyPointOfCare.PhysicalLocation.Facility.Value(NCID_3M_DEFAULT_CONTEXT)
    ElseIf Not (objString.CreatePointOfCare Is Nothing) Then
        tmpAllergy.PointOfCareFacility = _
           objString.CreatePointOfCare.PhysicalLocation.Facility.Value(NCID_3M_DEFAULT_CONTEXT)
    End If
    Exit Sub

ErrHandler:
    Err.Raise Err.Number, MODULE & ".CopyClinicianAndPointOfCare", Err.Description

End Sub


Private Sub StoreCommentsCHCSIDFacilities(ByRef SavAllergy As DTOs.Allergy, ByRef objAllergy As AllergyInfo)
    Dim tcTemp As ITextComment
    Dim stTemp As IStringText
    
    Dim i As Long
    
    On Error GoTo ErrHandler

    If SavAllergy.InfoSource <> vbNullString Then
       If SavAllergy.InfoSourceNCID <> vbNullString Then
          objAllergy.SourceOfInfo.ValueConcept.UniqueID = SavAllergy.InfoSourceNCID
       Else
          objAllergy.SourceOfInfo.LText = Mid$(SavAllergy.InfoSource, 8)
       End If
    End If

    If SavAllergy.Comment <> vbNullString Then
        Set tcTemp = objAllergy.TextComments.Add 'objComments
        Set stTemp = tcTemp.Comments.Add
        With stTemp
            .TextType = "englishText"
            .Text.Text = SavAllergy.Comment
        End With
    End If
    
    If SavAllergy.CHCS_ID <> vbNullString Then
        Set tcTemp = objAllergy.TextComments.Add 'objComments
        Set stTemp = tcTemp.Comments.Add
        With stTemp
            .TextType = "englishText"
            .Text.Text = SavAllergy.CHCS_ID
        End With
    End If
    
    For i = 0 To SavAllergy.CHCS_FACILITIES_Count - 1
         Set tcTemp = objAllergy.TextComments.Add 'objComments
         Set stTemp = tcTemp.Comments.Add
         With stTemp
             .TextType = "englishText"
             If SavAllergy.ModsFlg Then
                If Right$(SavAllergy.CHCS_FACILITIES(i), 1) = "U" Then
                   .Text.Text = "FACILITY:" & SavAllergy.CHCS_FACILITIES(i)
                Else
                   .Text.Text = "FACILITY:" & SavAllergy.CHCS_FACILITIES(i) & "U"
                End If
             Else
                .Text.Text = "FACILITY:" & SavAllergy.CHCS_FACILITIES(i)
             End If
         End With
    Next
    
    Set tcTemp = objAllergy.TextComments.Add 'objComments
    Set stTemp = tcTemp.Comments.Add
    With stTemp
        .TextType = "englishText"
        If SavAllergy.CHCSII_ENTRY = "Y" Then
           .Text.Text = "ENTERED IN CHCSII:Y"
        Else
           .Text.Text = "ENTERED IN CHCSII:N"
        End If
    End With

    Exit Sub

ErrHandler:
    Err.Raise Err.Number, MODULE & ".StoreCommentsCHCSIDFacilities", Err.Description


End Sub

