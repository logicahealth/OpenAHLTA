VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AllergiesSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const MODULE As String = "DM_Allergies.AllergiesSQL"
Private Const ERROR_NUMBER As Long = 12345

Private mConn As CHCSII_CONN.Conn
Private mMode As CHCSII_CONN.eBackend       'Data Access mode used for DAS object

Public Sub Init(ByRef Conn As CHCSII_CONN.Conn, ByVal DataAccessMode As CHCSII_CONN.eBackend)
  Set mConn = Conn
  mMode = DataAccessMode
End Sub

Private Sub CopyAllergyInfoFromRSToDTO(ByRef rs As ADODB.Recordset, ByRef dtoAllergy As DTOs.Allergy, ByRef objConcept As GEMS_ConceptCtrl)
    On Error GoTo ErrHandler
    
    With dtoAllergy
        .CreateUserNCID = IsNullValue(rs.Fields("CreatedBy"))
        .CreateTime = IsNullValue(rs.Fields("CreatedOn"))
        .UnitNumber = IsNullValue(rs.Fields("Unit_Number"))
        .PointOfCareFacility = IsNullValue(rs.Fields("PointOfCareFacilityNCID"), 0)
        .DataId = IsNullValue(rs.Fields("DataID"), 0)
        .CDRDataID = IsNullValue(rs.Fields("CDR_DATAID"), 0)
        
        'SCR #47869... NotedDate can be null
        .NotedDate = IsNullValue(Format$(rs.Fields("NotedDate"), gsDefaultDate), Format$(Now, gsDefaultDate))
        
        .AllergenNCID = IsNullValue(rs.Fields("AllergenNCID"), 0)
        .Allergen = IsNullValue(rs.Fields("Allergen"), 0)
        If Not IsNull(rs.Fields("AllergyTypeNCID")) Then
            .AllergyTypeNCID = rs.Fields("AllergyTypeNCID")
            Select Case .AllergyTypeNCID
                Case gNKA_TYPE
                    .AllergyType = "NKA"
                Case Else
                    .AllergyType = "Med"
            End Select
        Else
            .AllergyTypeNCID = vbNullString
            .AllergyType = vbNullString
        End If
        .InfoSourceNCID = IsNullValue(rs.Fields("InfoSourceNCID"), 0)
        .InfoSource = ReturnRepFromNCID(gINFOSOURCE_DOMAIN, CLng(.InfoSourceNCID))
        If Not IsNull(rs.Fields("ModifyUserNCID")) Then
            .ModifyUserNCID = rs.Fields("ModifyUserNCID")
        End If
        objConcept.UniqueID = .ModifyUserNCID
        .ModifyUser = Trim$(objConcept.PrefRep("2000").Representation)
        
        'SCR #50335 - typo had CreatedBy being assigned to ClinicianNCID
        If Not IsNull(rs.Fields("ClinicianNCID")) Then
            .ClinicianNCID = rs.Fields("ClinicianNCID")
        End If
        objConcept.UniqueID = .ClinicianNCID
        .Clinician = Trim$(objConcept.PrefRep("2000").Representation)
    End With
    
    Exit Sub

ErrHandler:
    Err.Raise Err.Number, MODULE & ".CopyAllergyInfoFromRSToDTO", Err.Description

End Sub

Public Function DeleteAllergy(ByRef DelAllergy As DTOs.Allergy) As Boolean
    On Error GoTo ErrHandler
    
    DeleteAllergy = False
    Dim das As ICHCSII_DAS
    Set das = mConn.CHCSII_DAS(mMode)
    
    'if the DataID is "0", it's an NKA record, and actually gets "deleted"
    If DelAllergy.DataId = "0" Then
        das.ExecuteSQL "delete from ALLERGIES where UNIT_NUMBER = " & DelAllergy.UnitNumber & " and DATAID = " & DelAllergy.DataId
    
    'otherwise, we just set the updateflag to "D"
    Else
        das.ExecuteSQL "update ALLERGIES set UPDATEFLAG = 'D' where UNIT_NUMBER = " & DelAllergy.UnitNumber & " and DATAID = " & DelAllergy.DataId
    
    End If
    
    DeleteAllergy = True
    
    Exit Function
ErrHandler:
    Err.Raise Err.Number, MODULE & ".DeleteAllergy", Err.Description

End Function

Private Function GetAllergyInsertSQL(ByRef SavAllergy As DTOs.Allergy, ByVal NewFlag As Boolean) As String
    Dim sFieldsSQL As String
    Dim sValuesSQL As String
    Dim sUpdateSQL As String
    
    On Error GoTo ErrHandler
    
    With SavAllergy

        sFieldsSQL = "insert into ALLERGIES (UNIT_NUMBER, DATAID, ALLERGEN, ALLERGENNCID, ALLERGYTYPENCID, CLINICIANNCID, CREATETIME, CREATEUSERNCID, MODIFYUSERNCID " _
                    & ", FACILITY_NCID, INFOSOURCENCID, MODIFIED, NOTEDDATE, POINTOFCAREFACILITYNCID, UPDATEFLAG) "
        
        If NewFlag Then 'adding new
            sValuesSQL = "values (" & CStr(.UnitNumber) & ", " & .DataId & ", " & mConn.SQLQuote(.Allergen) & ", " & vbCrLf & .AllergenNCID & ", " & .AllergyTypeNCID & ", " & .ClinicianNCID & vbCrLf & ", " _
                        & mConn.SQLDate(Now) & ", " & .CreateUserNCID & ", " & .ModifyUserNCID & ", " & .FacilityNCID & ", " & .InfoSourceNCID & ", " & .Modified & ", " _
                        & mConn.SQLDate(.NotedDate) & ", " & .PointOfCareFacilityNCID & ", 'N')"
                         
        ElseIf InStr(.Allergen, "***DELETE***") <> 0 Then  'not sure why deletions are being handled here; but, keeping
                                                            ' the logic for compatibility
                                                            
            sUpdateSQL = "UPDATE ALLERGIES SET UPDATEFLAG = 'D' "
            .UpDateFlag = "D"
        
        Else    'update
            sUpdateSQL = "update ALLERGIES set ALLERGEN = '" & .Allergen & "', ALLERGENNCID = " & .AllergenNCID & ", "
            
            If Not Trim$(.ClinicianNCID) = vbNullString Then
               sUpdateSQL = sUpdateSQL & "CLINICIANNCID = " & .ClinicianNCID & ", "
            End If
                
            If Not Trim$(.InfoSourceNCID) = vbNullString Then
               sUpdateSQL = sUpdateSQL & "INFOSOURCENCID = " & .InfoSourceNCID & ", "
            End If
                
            sUpdateSQL = sUpdateSQL & "MODIFIED = " & .Modified & ", MODIFYTIME = " & mConn.SQLDate(Now) & ", " & vbCrLf
            sUpdateSQL = sUpdateSQL & "ModifyUserNCID = " & .ModifyUserNCID & ", "
            
            'SCR #50265- for some reason, code was setting sepchar field (does not exist) instead
            '       of CDR_DataID
            If .CDRDataID <> vbNullString Then
                sUpdateSQL = sUpdateSQL & "CDR_DataID = " & .CDRDataID & ", "
            End If
            
            If Not Trim$(.InfoSourceNCID) = vbNullString Then
               sUpdateSQL = sUpdateSQL & "NotedDate = " & mConn.SQLDate(.NotedDate) & ", "
            End If
            
            sUpdateSQL = sUpdateSQL & "UPDATEFLAG = 'U'"
        End If
            
        If Not NewFlag Then
          sUpdateSQL = sUpdateSQL & " WHERE UNIT_NUMBER = " & .UnitNumber & " AND DATAID = " & .DataId
          GetAllergyInsertSQL = sUpdateSQL
        Else
           GetAllergyInsertSQL = sFieldsSQL & " " & sValuesSQL
        End If
    End With

    Exit Function

ErrHandler:
    Err.Raise Err.Number, MODULE & ".GetAllergyInsertSQL", Err.Description

End Function

Public Function RetrieveAllergies(ByRef params As iParameters) As DTOs.Allergies
    On Error GoTo ErrHandler
    
    'AllergySearchParameters:
        'IncludeReactions As Boolean
        'IncludeComment As Boolean
        'PatientUnitNumber As String    'required
        'PatientIENNumber As String     'only supported under CHCS
    
    Dim dtoReturn As DTOs.Allergies
    Dim dtoAllergy As DTOs.Allergy
    Dim rs As ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Dim objConcept As GEMS_ConceptCtrl
    Dim sSQL As String
    
    Dim param As AllergySearchParam
    Set param = params
    
    Set objConcept = New GEMS_ConceptCtrl
    
    'pull the allergies for the patient
    sSQL = "select * from ALLERGIES where UNIT_NUMBER = " & param.PatientUnitNumber _
                & " and (UPDATEFLAG <> 'D' or UPDATEFLAG is NULL)"
    
    Set rs = mConn.CHCSII_DAS(mMode).OpenRecordset(sSQL)
    
    Set dtoReturn = New DTOs.Allergies
    
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        Do Until rs.EOF
            Set dtoAllergy = New DTOs.Allergy
            
            CopyAllergyInfoFromRSToDTO rs, dtoAllergy, objConcept
            
            With dtoAllergy
                
                'AllergyReactions
                SetAllergyReactions dtoAllergy, param, objConcept
                
AllergenComments:
                'Comments
                If param.IncludeComment Then
                    SetAllergyComments dtoAllergy, param
                End If
SkipComments:
                rs.MoveNext
            End With
            dtoReturn.Add dtoAllergy
        Loop
    
        #If DebugOn Then
            DL.DebugLog Exe, MODULE, " ", "P-After mRetrieve", DR
        #End If
    End If
    
    Set RetrieveAllergies = dtoReturn
    GoTo CleanUp

ErrHandler:
    Err.Raise Err.Number, MODULE & ".RetrieveAllergies", Err.Description
    Set RetrieveAllergies = Nothing
    
CleanUp:
    On Error Resume Next
    If rs.State = adStateOpen Then
        rs.Close
    End If
    Set rs = Nothing
    If rs2.State = adStateOpen Then
        rs2.Close
    End If
    Set rs2 = Nothing
    Set objConcept = Nothing

End Function

Public Function RetrieveAllergyReactions(ByRef params As iParameters) As DTOs.AllergyReactions
    On Error GoTo ErrHandler
    
    Dim bMaskSearch As Boolean
    Dim sMask As String
    Dim iIndex As Long
    Dim tmpReaction As DTOs.AllergyReaction
    
    Dim param As ReactionSearchParams
    Set param = params
    
    Dim result As AllergyReactions
    Set result = New AllergyReactions
    
    Dim oPickListCtrl As GEMS_PickListCtrl
    Set oPickListCtrl = New GEMS_PickListCtrl
    
    oPickListCtrl.Domain = gREACTION_DOMAIN
    oPickListCtrl.Load
    
    param.PrefixKeyword = UCase$(param.PrefixKeyword)
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
                    tmpReaction.NCID = .NCID
                    tmpReaction.Name = UCase$(.Representation)
                    result.Add tmpReaction
                End If
            Else
                If (UCase$(.Representation) = sMask) Or (sMask = vbNullString) Then
                    Set tmpReaction = New DTOs.AllergyReaction
                    tmpReaction.NCID = .NCID
                    tmpReaction.Name = UCase$(.Representation)
                    result.Add tmpReaction
                End If
            End If
        End With
    Next
                
    Set RetrieveAllergyReactions = result
                
    GoTo CleanUp
    
ErrHandler:
    Err.Raise Err.Number, MODULE & ".RetrieveAllergyReactions", Err.Description

CleanUp:
    On Error Resume Next
    Set oPickListCtrl = Nothing
    
End Function

Public Function RetrieveAllergens(ByRef params As DL_Support.iParameters) As DTOs.Allergens
    'currently used by MapperCDR and MapperTheater
    
    On Error GoTo ErrHandler
    
    Dim param As AllergenSearchParam
    Set param = params
    
    Dim ReturnAllergens As DTOs.Allergens
    Set ReturnAllergens = New DTOs.Allergens
    
    Dim tmpAllergen As DTOs.Allergen
    
    Dim rs As ADODB.Recordset
    Dim sSQL As String
    Dim DataIO As ICHCSII_DAS
    Set DataIO = mConn.CHCSII_DAS(mMode)
    
    On Error GoTo ErrHandler
    If param.PrefixKeyword = vbNullString Then
        sSQL = "select IEN, BG_GNN, ALLTYPE from CHCS_ALLERGIES where OBSOLETE_DATE is NULL"
    Else
        If Right$(Trim$(param.PrefixKeyword), 1) = "*" Then
            sSQL = "select IEN, BN_GNN, GNN, ALLTYPE from CHCS_ALLERGIES where BN_GNN like " & mConn.SQLQuote(UCase$(param.PrefixKeyword)) & " AND OBSOLETE_DATE is NULL"
        Else
            sSQL = "select IEN, BN_GNN, GNN, ALLTYPE from CHCS_ALLERGIES where BN_GNN like " & mConn.SQLQuote(UCase$(param.PrefixKeyword) & "%") & " and OBSOLETE_DATE is NULL"
        End If
    End If
    
    Set rs = DataIO.OpenRecordset(sSQL)
    
    If Not rs Is Nothing Then
        Do Until rs.EOF
            Set tmpAllergen = New DTOs.Allergen
            
            With tmpAllergen
                .NCID = CStr(rs.Fields("ien").Value)
                .Name = rs.Fields("bn_gnn").Value
                If rs.Fields("gnn").Value <> vbNullString Then
                    .Name = .Name & " (" & rs.Fields("gnn").Value & ")"
                Else
                    .Name = .Name & " {" & rs.Fields("alltype").Value & "}"
                End If
            End With
                
            ReturnAllergens.Add tmpAllergen
            rs.MoveNext
        Loop
    End If

    GoTo CleanUp
        
ErrHandler:
    Err.Raise Err.Number, MODULE & ".RetrieveAllergens", Err.Description
        
CleanUp:
    Set RetrieveAllergens = ReturnAllergens
    If rs.State = adStateOpen Then
        rs.Close
    End If
    Set rs = Nothing

End Function

Public Function SaveAllergies(ByRef SavAllergies As DTOs.Allergies) As Boolean
    On Error GoTo ErrHandler
    'this will wipe out any existing allergies
    
    Dim i As Long, reactIdx As Long
    Dim oSQLUpdater As ICHCSII_SQL.ISqlOpsEx
      
    Set oSQLUpdater = mConn.CHCSII_SQLOPS_EX(mMode) '<SCR 36921
   
    If IsNumeric(SavAllergies.Item(1).EncounterFacilityNCID) And IsNumeric(SavAllergies.Item(1).ClinicianNCID) Then
        oSQLUpdater.ValidateFacility SavAllergies.Item(1).EncounterFacilityNCID, True
        'oSQLUpdater.ValidateProvider cwobjLogonEnviron.UserNCID, True, SavAllergies.Item(1).EncounterFacilityNCID, SavAllergies.Item(1).ClinicianNCID
        oSQLUpdater.ValidateProvider SavAllergies.Item(1).ClinicianNCID, True, SavAllergies.Item(1).EncounterFacilityNCID, SavAllergies.Item(1).ClinicianNCID
    ElseIf IsNumeric(SavAllergies.Item(1).EncounterFacilityNCID) Then
        oSQLUpdater.ValidateFacility SavAllergies.Item(1).EncounterFacilityNCID, True
        'oSQLUpdater.ValidateProvider cwobjLogonEnviron.UserNCID, True, SavAllergies.Item(1).EncounterFacilityNCID
        oSQLUpdater.ValidateProvider SavAllergies.Item(1).ClinicianNCID, True, SavAllergies.Item(1).EncounterFacilityNCID
    Else
        'oSQLUpdater.ValidateProvider cwobjLogonEnviron.UserNCID, True
        oSQLUpdater.ValidateProvider SavAllergies.Item(1).ClinicianNCID, True
    End If
   
    oSQLUpdater.Execute "select count(*) from Allergies where UNIT_NUMBER=" & SavAllergies.Item(1).UnitNumber
    If CLng(oSQLUpdater.Value(0)) > 0 Then
        oSQLUpdater.Execute "delete from allergyreaction where UNIT_NUMBER=" & SavAllergies.Item(1).UnitNumber
        oSQLUpdater.Execute "delete from allergycomment  where UNIT_NUMBER=" & SavAllergies.Item(1).UnitNumber
        oSQLUpdater.Execute "delete from allergies       where UNIT_NUMBER=" & SavAllergies.Item(1).UnitNumber
    End If
   
    'figure out if update, insert or delete
    For i = 1 To SavAllergies.count
      
        If Not SetAllergyInfo(SavAllergies.Item(i), oSQLUpdater) Then
            GoTo ExitRoutine
        End If
        
        For reactIdx = 1 To SavAllergies.Item(i).AllergyReactions.count
            With oSQLUpdater
                .Execute "select AllergenNCID, ReactionNCID, Reaction, unit_number, UpDateFlag from allergyreaction       where UNIT_NUMBER=" & SavAllergies.Item(1).UnitNumber
            
                'Allergy can have multiple reactions  SCR-26259
                SetUpdateflag SavAllergies.Item(i), oSQLUpdater
                .AddNew
            
                .Value("AllergenNCID") = SavAllergies.Item(i).AllergenNCID
                .Value("ReactionNCID") = SavAllergies.Item(i).AllergyReactions.Item(reactIdx).NCID
                .Value("Reaction") = SavAllergies.Item(i).AllergyReactions.Item(reactIdx).Name
                
                .Value("unit_number") = SavAllergies.Item(i).UnitNumber
                
                .Value("UpDateFlag") = SavAllergies.Item(i).UpDateFlag
            
                If Not .Update Then
                    GoTo ExitRoutine
                End If
            End With
        Next reactIdx
      
        If Len(SavAllergies.Item(i).Comment) Then
            With oSQLUpdater
                .Execute "select AllergenNCID, unit_number, allergyComment, updateflag from allergycomment       where UNIT_NUMBER=" & SavAllergies.Item(1).UnitNumber
                If SetUpdateflag(SavAllergies.Item(i), oSQLUpdater) = "I" Then
                    .AddNew
                End If
            
                .Value("AllergenNCID") = SavAllergies.Item(i).AllergenNCID
                
                If (Len(SavAllergies.Item(i).UnitNumber) <> 0) Then
                    .Value("unit_number") = SavAllergies.Item(i).UnitNumber
                Else
                    .Value("unit_number") = SavAllergies.Item(1).UnitNumber
                End If
                
                .Value("allergyComment") = SavAllergies.Item(i).Comment
                .Value("updateflag") = SavAllergies.Item(i).UpDateFlag
                If Not .Update Then
                    GoTo ExitRoutine
                End If
            End With
        End If

   Next i
   
   SaveAllergies = True
   GoTo ExitRoutine

ErrHandler:
    Err.Raise Err.Number, MODULE & ".SaveAllergies", Err.Description

ExitRoutine:
   Set oSQLUpdater = Nothing

Exit Function
Resume
End Function

Public Function SaveAllergy(ByRef SavAllergy As DTOs.Allergy) As Boolean

    '****************************************************************
    '.DataID property should be set to 0 for a new addition
    '
    '
    'this logic originally got the following information from globals:
    '   .UnitNumber = glCurPat
    '   .Allergen, .AllergenNCID = gSelectedAllergen
    'will now expect it to be set before this method is called
    '****************************************************************

    On Error GoTo ErrHandler
    
    Dim sDeleteSQL As String
    Dim sFieldsSQL As String
    Dim sValuesSQL As String
    Dim objICHCSII_DAS As ICHCSII_DAS
    Dim oSeqId As ICHCSII_SQL.ISqlOps
    Dim count As Long
    Dim NewFlag As Boolean
    Dim sSQL As String
    
    With SavAllergy
        Set objICHCSII_DAS = mConn.CHCSII_DAS(mMode)
        Set oSeqId = mConn.CHCSII_SQLOPS(mMode)
        
        NewFlag = (.DataId = 0)
        If NewFlag Then
            .DataId = CStr(oSeqId.GetNextID)
        End If
        
        If InStr(1, .Allergen, "<CHCS IEN|DESC>:") = 1 Then
            Dim sTokens() As String
            Dim i As Integer
            sTokens = Split(.Allergen, ":")
            If (sTokens(1) = "Other") Then
                sTokens = Split(sTokens(2), "|")
            Else
                sTokens = Split(sTokens(1), "|")
            End If
            If UBound(sTokens) = 0 Then
                'This should not happen, but we will
                'assume if it does that this is the Allergen.
                '--------------------------------------------
                .Allergen = sTokens(0)
            ElseIf UBound(sTokens) = 1 Then
                .Allergen = sTokens(1)
                .AllergenNCID = sTokens(0)
            Else
                'This should not happen either, but if it
                'does we will reassemble tokens 1 through n.
                '-----------------------------------
                Dim sSubTokens() As String
                ReDim sSubTokens(0 To UBound(sTokens) - 1)
                For i = 1 To UBound(sTokens)
                    sSubTokens(i - 1) = sTokens(i)
                Next
                .Allergen = Join(sSubTokens, "|")
                .AllergenNCID = sTokens(0)
            End If
        Else
          .Allergen = "Other: " & .Allergen
        End If
        Set oSeqId = Nothing
        
        '!!
        objICHCSII_DAS.BeginTrans
        
        If .DataId <> "0" Then             ' T.Noren  03/16/01
            
            sSQL = GetAllergyInsertSQL(SavAllergy, NewFlag)
            objICHCSII_DAS.ExecuteSQL sSQL
        
AllergenReactions:
                
            sDeleteSQL = "delete from ALLERGYREACTION where UNIT_NUMBER = " & .UnitNumber & " and ALLERGENNCID = " & .AllergenNCID
            objICHCSII_DAS.ExecuteSQL sDeleteSQL
            
            For count = .AllergyReactions.count To 1 Step -1
                If InStr(.Allergen, "***DELETE***") <> 0 Then
                    
                    .AllergyReactions.Remove (count)
                
                Else
                
                    If .AllergyReactions.Item(count).NCID = vbNullString Then
                        .AllergyReactions.Item(count).NCID = "0"
                    End If
                     
                     'SCR 18836 WR 2/20/2002 issue with 2 reaction fields in SQL string
                    If .AllergyReactions.Item(count).NCID = "76281" Then
                        If InStr(.AllergyReactions.Item(count).Name, "Other: ") <> 0 Then
                            .AllergyReactions.Item(count).Name = Mid$(.AllergyReactions.Item(count).Name, Len("Other: ") + 1)
                        End If
                    End If
                    sFieldsSQL = "insert into ALLERGYREACTION (UNIT_NUMBER, REACTION, REACTIONNCID, ALLERGENNCID, "
                     
                    sValuesSQL = "values (" & CStr(.UnitNumber) & ", " & mConn.SQLQuote(.AllergyReactions.Item(count).Name) & ", " & .AllergyReactions.Item(count).NCID & ", " & .AllergenNCID
                     
                    sFieldsSQL = sFieldsSQL & "UPDATEFLAG) "
                    sValuesSQL = sValuesSQL & ", 'N') "
                    objICHCSII_DAS.ExecuteSQL sFieldsSQL & " " & sValuesSQL
                End If
            Next count
                
AllergenComments:
                
            sDeleteSQL = "delete from ALLERGYCOMMENT where UNIT_NUMBER = " & .UnitNumber & " and ALLERGENNCID = " & .AllergenNCID
            objICHCSII_DAS.ExecuteSQL sDeleteSQL
            sFieldsSQL = "insert into ALLERGYCOMMENT (UNIT_NUMBER, ALLERGENNCID, ALLERGYCOMMENT, UPDATEFLAG) " _
                   & "values (" & CStr(.UnitNumber) & ", " & .AllergenNCID & ", " & mConn.SQLQuote(.Comment) & ", 'N') "
            objICHCSII_DAS.ExecuteSQL sFieldsSQL
            
            '!!
            objICHCSII_DAS.CommitTrans
            
        Else 'Invalid Sequence ID
            objICHCSII_DAS.RollbackTrans
            Err.Raise ERROR_NUMBER, MODULE & ".SaveAllergy", "Unable to obtain next sequence ID"
          
        End If
    End With
    GoTo CleanUp

ErrHandler:
    objICHCSII_DAS.RollbackTrans
    Err.Raise Err.Number, MODULE & ".SaveAllergy", Err.Description

CleanUp:
    Exit Function

End Function

Private Function SetAllergyInfo(ByRef SavAllergy As DTOs.Allergy, ByRef oSQLUpdater As ICHCSII_SQL.ISqlOpsEx) As Boolean
    On Error GoTo ErrHandler
    
    SetAllergyInfo = False
    
    With oSQLUpdater
    
          '**** PORTED FROM THEATER - 8/6/03
          ' Begin:SCR #29715;   Developer: MAG 11/25/2002 11:06 AM
          'validate for every allergy SCR-26289
          If IsNumeric(SavAllergy.EncounterFacilityNCID) And IsNumeric(SavAllergy.ClinicianNCID) Then
              .ValidateFacility SavAllergy.EncounterFacilityNCID, True
              .ValidateProvider SavAllergy.ClinicianNCID, True, SavAllergy.EncounterFacilityNCID, SavAllergy.ClinicianNCID
          ElseIf IsNumeric(SavAllergy.EncounterFacilityNCID) Then
              .ValidateFacility SavAllergy.EncounterFacilityNCID, True
              '.ValidateProvider cwobjLogonEnviron.UserNCID, True, SavAllergy.EncounterFacilityNCID
              .ValidateProvider SavAllergy.ClinicianNCID, True, SavAllergy.EncounterFacilityNCID
          Else
              '.ValidateProvider cwobjLogonEnviron.UserNCID, True
              .ValidateProvider SavAllergy.ClinicianNCID, True
          End If
        
          'Call .ValidateFacility(SavAllergy.EncounterFacilityNCID, True)
          'Call .ValidateProvider(SavAllergy.ClinicianNCID, True, SavAllergy.EncounterFacilityNCID, SavAllergy.ClinicianNCID)
          ' End: SCR #29715;
          '**** END PORT
      
          .ValidateFacility SavAllergy.EncounterFacilityNCID, True
        
          'SL- 12/08/03-  SCR #46537
          '   there's no ClinicianNCID for NKA, so this was throwing a type mismatch error
          If (SavAllergy.ClinicianNCID <> vbNullString) And (Not SavAllergy.AllergyType <> "NKA") Then
              .ValidateProvider SavAllergy.ClinicianNCID, True, SavAllergy.EncounterFacilityNCID, SavAllergy.ClinicianNCID
          End If
            
          .Execute "select * from allergies       where UNIT_NUMBER=" & SavAllergy.UnitNumber
          If SetUpdateflag(SavAllergy, oSQLUpdater) = "I" Then
              .AddNew
          End If
         
          .Value("Allergen") = SavAllergy.Allergen
          .Value("AllergenNCID") = SavAllergy.AllergenNCID
          .Value("AllergyTypeNCID") = SavAllergy.AllergyTypeNCID
          'swap DataIds
          .Value("CDR_DATAID") = SavAllergy.DataId
          .Value("ClinicianNCID") = SavAllergy.ClinicianNCID
         
         .Value("CreateTime") = SavAllergy.CreateTime
         
         
         'SF cannot comment this out or it violates ALLERGIES_FK2
         If SavAllergy.CreateUserNCID = Empty Then
             .Value("CreateUserNCID") = SavAllergy.ClinicianNCID
         Else
             .Value("CreateUserNCID") = SavAllergy.CreateUserNCID
         End If
         
         'SF cannot comment this out or it violates ALLERGIES_FK3
         If SavAllergy.ModifyUserNCID = Empty Then
           .Value("MODIFYUSERNCID") = SavAllergy.ClinicianNCID
         Else
           .Value("MODIFYUSERNCID") = SavAllergy.ModifyUserNCID
         End If
        
         .Value("DataId") = oSQLUpdater.GetNextID(esMAINSEQ)     'SavAllergy.DataId
         .Value("FACILITY_NCID") = SavAllergy.EncounterFacilityNCID
         
         'SCR 19356
         'Allergies.infosourcencid type number(22). If empty, assume Patient is the info source
          If IsNumeric(SavAllergy.InfoSourceNCID) Then
              .Value("InfoSourceNCID") = SavAllergy.InfoSourceNCID
          Else
              .Value("InfoSourceNCID") = NCID_INFO_SOURCE_PATIENT
          End If
          
          .Value("CREATEDBY") = SavAllergy.CreateUserNCID
          
          .Value("NotedDate") = SavAllergy.NotedDate
          .Value("PointOfCareFacilityNCID") = SavAllergy.PointOfCareFacilityNCID
          .Value("Tmip_Status") = CLng("0" & SavAllergy.Tmip_Status)
          
          .Value("unit_number") = SavAllergy.UnitNumber
              
          .Value("updateflag") = SavAllergy.UpDateFlag
          
          SetAllergyInfo = .Update
          
    End With
    
    Exit Function

ErrHandler:
    Err.Raise Err.Number, MODULE & ".SetAllergyInfo", Err.Description
    Exit Function
    Resume
End Function

Public Function SaveAllergyTransaction(ByRef AllergyTrans As DTOs.AllergyTransaction) As Boolean
    'currently used by MapperCDR and MapperTheater
    
    On Error GoTo ErrHandler
    
    Dim rs As ADODB.Recordset
    Dim sSQL As String

    With AllergyTrans
        If .AllergyTransID = "0" Or .AllergyTransID = vbNullString Then
            sSQL = "SELECT cdrplus.allergy_trans_id_sequence.nextval FROM dual"
            Set rs = mConn.CHCSII_DAS(mMode).OpenRecordset(sSQL)
            .AllergyTransID = rs.Fields(0)
        
            sSQL = "INSERT INTO allergy_transaction (allergy_trans_id, unit_number, patient_ien, " _
                     & "facility_ncid, trans_type, trans_date, update_status) " _
                     & "VALUES (" & .AllergyTransID & ", " & .UnitNumber & ", " _
                     & .PatientIEN & ", " _
                     & .FacilityNCID & ", " _
                     & .TransType & ", '" & Format$(.TransDate, "YYYYMMDDHHNNSS") & "', 0)"
        Else
            sSQL = "UPDATE allergy_transaction SET trans_date = '" & Format$(.TransDate, "YYYYMMDDHHNNSS") _
                    + "', update_status = 0 WHERE allergy_trans_id = " & .AllergyTransID
        End If
    End With
    
    mConn.CHCSII_DAS(mMode).ExecuteSQL sSQL
    
    SaveAllergyTransaction = True
    GoTo CleanUp
ErrHandler:
    Err.Raise Err.Number, MODULE & ".SaveAllergyTransaction", Err.Description
    
CleanUp:
    On Error Resume Next
    SaveAllergyTransaction = False
    Set rs = Nothing
End Function

Private Sub SetAllergyComments(ByRef dtoAllergy As DTOs.Allergy, ByRef param As DM_Allergies.AllergySearchParam)
    Dim rs As ADODB.Recordset
    Dim sSQL As String
    Dim sTemp As String
    Dim sComments() As String
    Dim i As Long
    
    On Error GoTo ErrHandler
    
    With dtoAllergy
        sSQL = "select ALLERGYCOMMENT from ALLERGYCOMMENT where UNIT_NUMBER = " & param.PatientUnitNumber _
            & " and ALLERGENNCID = " & .AllergenNCID & " and (UPDATEFLAG <> 'D' or UPDATEFLAG is NULL)"
        Set rs = mConn.CHCSII_DAS(mMode).OpenRecordset(sSQL)
        If rs.RecordCount > 0 Then
            rs.MoveFirst
            
            Do Until rs.EOF
                sTemp = IsNullValue(rs.Fields("ALLERGYCOMMENT"), vbNullString)
                sComments = Split(sTemp, "|")
                For i = 0 To UBound(sComments, 1)
                    If Left$(sComments(i), 18) = "ENTERED IN CHCSII:" Then
                        
                        .CHCSII_ENTRY = Right$(sComments(i), 1)
                        If .CHCSII_ENTRY = "N" Then
                           .Clinician = "CHCS"
                        End If
                       
                    ElseIf Left$(sComments(i), 8) = "CHCS ID:" Then
                        .CHCS_ID = sComments(i)
                    Else
                        .Comment = .Comment & sComments(i)
                    End If
                Next
                rs.MoveNext
            Loop
        Else
            .Comment = vbNullString
        End If
    End With
    
    GoTo CleanUp
    
ErrHandler:
    Err.Raise Err.Number, MODULE & ".SetAllergyComments", Err.Description

CleanUp:
    On Error Resume Next
    rs.Close
    Set rs = Nothing
End Sub

Private Sub SetAllergyReactions(ByRef dtoAllergy As DTOs.Allergy, ByRef param As DM_Allergies.AllergySearchParam, ByRef objConcept As GEMS_ConceptCtrl)
    On Error GoTo ErrHandler
    
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    Dim tmpReaction As DTOs.AllergyReaction
    
    With dtoAllergy
        If param.IncludeReactions Then
            Set .AllergyReactions = New AllergyReactions
            sSQL = "select REACTION, REACTIONNCID from ALLERGYREACTION where UNIT_NUMBER = " & param.PatientUnitNumber _
                    & " and ALLERGENNCID = " & .AllergenNCID & " and (UPDATEFLAG <> 'D' or UPDATEFLAG is NULL)"
                    
            Set rs = mConn.CHCSII_DAS(mMode).OpenRecordset(sSQL)
            If rs.RecordCount > 0 Then
            
                rs.MoveFirst
                Set tmpReaction = New AllergyReaction
                
                Do Until rs.EOF
                    Set tmpReaction = New AllergyReaction
                    tmpReaction.NCID = rs.Fields("ReactionNCID")
                    tmpReaction.Name = IsNullValue(rs.Fields("Reaction"))
                    If (tmpReaction.NCID = "0") Or (tmpReaction.NCID = 76281) Then
                        'SCR #52651... once "Other" and "Unknown" have their NCIDs reset, they
                        '      can't be saved again due to table uniqueness constraints.
                        'tmpReaction.NCID = vbNullString
                    Else
                        objConcept.UniqueID = tmpReaction.NCID
                        tmpReaction.Name = Trim$(objConcept.PrefRep("2000").Representation)
                    End If
                    .AllergyReactions.Add tmpReaction
                    rs.MoveNext
                Loop
            End If
        End If
    End With
    
    GoTo CleanUp
    
ErrHandler:
    Err.Raise Err.Number, MODULE & ".SetAllergyReactions", Err.Description

CleanUp:
    On Error Resume Next
    Set rs = Nothing

End Sub

Private Function SetUpdateflag(ByRef Data As DTOs.Allergy, ByRef rs As ICHCSII_SQL.ISqlOpsEx) As String
    On Error GoTo ErrHandler

    If rs.RecordCount > 0 Then
        rs.MoveFirst
        Do While Not rs.EOF
            If rs!AllergenNCID = Data.AllergenNCID Then
                Data.UpDateFlag = "U"
                SetUpdateflag = Data.UpDateFlag
                Exit Function
            End If
         
            rs.MoveNext
        Loop
    End If
   
    Data.UpDateFlag = "I"
    SetUpdateflag = Data.UpDateFlag
    
    Exit Function

ErrHandler:
    Err.Raise Err.Number, MODULE & ".SetUpdateFlag", Err.Description

End Function

