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
Attribute VB_Name = "MapperTheater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements DL_Interfaces.iMapperAllergies

Private Const MODULE As String = "DM_Allergies.MapperTheater"  'name of this module for error messages
Private Const ERROR_NUMBER As Long = 123456                    'arbitrary error number for "method not implemented errors

Private mConn As CHCSII_CONN.Conn               'CHCSII Connection

Private Function iMapperAllergies_DeleteAllergy(ByRef DelAllergy As DTOs.Allergy) As Boolean
Dim SQLHelper As AllergiesTheaterSQL

    Set SQLHelper = New AllergiesTheaterSQL
    SQLHelper.Init mConn
    
    iMapperAllergies_DeleteAllergy = SQLHelper.DeleteAllergy(DelAllergy)
    Set SQLHelper = Nothing

End Function

Private Sub iMapperAllergies_Init(ByRef Support As DL_Support.iParameters, ByRef Cache As DL_Support.CacheMapper)
    On Error GoTo ErrHandler

    'Save objects given to us by DataAccess Factory
    Set mConn = Support.Item("DBConn").Value
    
    Exit Sub

ErrHandler:
    Err.Raise Err.Number, MODULE & ".iMapperAllergies_Init", Err.Description

End Sub

Private Function iMapperAllergies_RetrieveAllergies(ByRef params As iParameters) As DTOs.Allergies
Dim SQLHelper As AllergiesTheaterSQL

    Set SQLHelper = New AllergiesTheaterSQL
    SQLHelper.Init mConn
    
    Set iMapperAllergies_RetrieveAllergies = SQLHelper.RetrieveAllergies(params)
    Set SQLHelper = Nothing

End Function

Private Function iMapperAllergies_RetrieveAllergens(ByRef params As DL_Support.iParameters) As DTOs.Allergens
Dim SQLHelper As AllergiesTheaterSQL

    Set SQLHelper = New AllergiesTheaterSQL
    SQLHelper.Init mConn
    
    Set iMapperAllergies_RetrieveAllergens = SQLHelper.RetrieveAllergens(params)
    Set SQLHelper = Nothing

End Function

Private Function iMapperAllergies_SaveAllergies(ByRef SavAllergies As DTOs.Allergies) As Boolean
Dim SQLHelper As AllergiesTheaterSQL

    Set SQLHelper = New AllergiesTheaterSQL
    SQLHelper.Init mConn
    
    iMapperAllergies_SaveAllergies = SQLHelper.SaveAllergies(SavAllergies)
    Set SQLHelper = Nothing

End Function

Private Function iMapperAllergies_SaveAllergy(ByRef SavAllergy As DTOs.Allergy) As Boolean
Dim SQLHelper As AllergiesTheaterSQL

    Set SQLHelper = New AllergiesTheaterSQL
    SQLHelper.Init mConn
    
    iMapperAllergies_SaveAllergy = SQLHelper.SaveAllergy(SavAllergy)
    Set SQLHelper = Nothing

End Function

Private Function iMapperAllergies_RetrieveAllergyReactions(ByRef params As iParameters) As DTOs.AllergyReactions
Dim SQLHelper As AllergiesTheaterSQL

    Set SQLHelper = New AllergiesTheaterSQL
    SQLHelper.Init mConn
    
    Set iMapperAllergies_RetrieveAllergyReactions = SQLHelper.RetrieveAllergyReactions(params)
    Set SQLHelper = Nothing

End Function

Private Function iMapperAllergies_SaveAllergyTransaction(ByRef AllergyTrans As DTOs.AllergyTransaction) As Boolean
Dim SQLHelper As AllergiesTheaterSQL

    Set SQLHelper = New AllergiesTheaterSQL
    SQLHelper.Init mConn
    
    iMapperAllergies_SaveAllergyTransaction = SQLHelper.SaveAllergyTransaction(AllergyTrans)
    Set SQLHelper = Nothing

End Function

Private Function iMapperAllergies_RetrieveAllergyTransactions(ByRef params As iParameters) As DTOs.AllergyTransactions
    Err.Raise ERROR_NUMBER, MODULE & ".iMapperAllergies_RetrieveAllergyTransactions", "This method has not been implemented in this data mode."
End Function

Private Function ValidateFacility(ByRef FacilityNCID As String, ByRef FacilityName As String, ByVal CreateIfNotFound As Boolean, _
                        Optional ByRef ClinicNCID As Long = 0, _
                        Optional ByRef ClinicName As String = vbNullString) As Boolean


    '-- see if the facility is in the facility table
    Dim bNeedFacility As Boolean        'flag... facility not in database
    Dim bNeedClinic As Boolean          'flag... clinic not in database
    Dim bSuccess As Boolean             'flag... becomes the return value of this function
    
    On Error GoTo ErrHandler
    '- see what's out there

    bNeedFacility = Not FindFacility(FacilityNCID)
    '642 Patch 2 changes made to synchronize baseline. Brian Mowbray 7/9/2002
    '    bNeedClinic = Not FindClinic(FacilityNCID, ClinicNCID)
    
    
    'If Clinic is not 0, then we need to find the Clinic
    bNeedClinic = True
    If ClinicNCID = 0 Then
      bNeedClinic = True
    Else
      bNeedClinic = Not FindClinic(FacilityNCID, ClinicNCID)
    End If

    If Not bNeedFacility And Not bNeedClinic Then
        ValidateFacility = True
        Exit Function
    End If
    
    If Not CreateIfNotFound Then
        ValidateFacility = Not bNeedFacility And Not bNeedClinic
        Exit Function
    End If
    
    '- so we shouldn't create what we don't have
    bSuccess = True

    If bNeedFacility Then
        bSuccess = BackfillFacility(FacilityNCID, FacilityName)
    End If
    
    
    '642 Patch 2 changes made to synchronize baseline. Brian Mowbray 7/9/2002
    If ClinicNCID <> 0 Then
        If bNeedClinic Then
            bSuccess = bSuccess And BackfillClinic(FacilityNCID, ClinicNCID, ClinicName)
        End If
    End If
    
    ValidateFacility = bSuccess
    
    Exit Function

ErrHandler:

    Err.Raise Err.Number, MODULE & ".ValidateFacility", Err.Description

End Function

Private Function FindFacility(ByVal FacilityNCID As Long) As Boolean
    '_by ncid only
    Dim rs As ADODB.Recordset   'recordset to catch query results
    
    On Error GoTo Func_Err
    
    Set rs = mConn.CHCSII_DAS(Auto).OpenRecordset("Select facility_ncid from facility_ncid_correlation where facility_ncid = " & FacilityNCID)
    
    If Not rs.EOF Then
        FindFacility = True
    Else
        FindFacility = False
    End If
    
Func_Exit:
    On Error Resume Next
    rs.Close
    Set rs = Nothing
    Exit Function

Func_Err:
    Err.Raise Err.Number, MODULE & ".FindFacility", Err.Description
    Resume Func_Exit

End Function

Private Function BackfillFacility(ByVal FacilityNCID As String, ByVal FacilityName As String) As Boolean
    Dim rs As ADODB.Recordset   'recordset to catch query results

    If FacilityNCID = 0 Then
        Exit Function
    End If
    On Error GoTo ErrHandler
    
    If FacilityName = vbNullString Or Len(FacilityName) > 10 Then
        FacilityName = "CHCS T ADT"
    End If
    
    With mConn.CHCSII_DAS(Auto)
        .ExecuteSQL "insert into facility_ncid_correlation (facility_ncid,  type) values (" & FacilityNCID & ",'" & FacilityName & "')"
        Set rs = .OpenRecordset("select ncid from ncids where ncid= " & FacilityNCID & " and domainncid =119")
        If rs.EOF Then
            .ExecuteSQL "insert into ncids (ncid, domainncid, description) values (" & FacilityNCID & ",119,'" & FacilityName & "')"
        End If
        rs.Close
        Set rs = .OpenRecordset("select ncid from ncids where ncid= " & FacilityNCID & " and domainncid =33962")
        If rs.EOF Then
            .ExecuteSQL "insert into ncids (ncid, domainncid, description) values (" & FacilityNCID & ",33962,'" & FacilityName & "')"
        End If
    End With
    
    BackfillFacility = True
   
Func_Exit:
    On Error Resume Next
    rs.Close
    Set rs = Nothing
    Exit Function
   
ErrHandler:
    Err.Raise Err.Number, MODULE & ".BackfillFacility", Err.Description
    Resume Func_Exit
    
End Function


Private Function BackfillClinic(ByVal FacilityNCID As String, _
                                ByVal ClinicNCID As String, _
                                ByVal ClinicName As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim meprs_code As String     'gets the next value from dual
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    
    If ClinicNCID = 0 Then
        Exit Function
    End If
    
    meprs_code = GetNextID()
    
    If ClinicName = vbNullString Or Len(ClinicName) > 40 Then
        ClinicName = "CHCSII-T ADT"
    End If
    
    'We are going to Search the Clinic Table for a Clinic NCID.
      'on Facility_ncid and Name
    sSQL = "SELECT NCID " & _
            "FROM Clinic " & _
            "WHERE Facility_NCID = " & FacilityNCID & " " & _
            "AND Name = '" & ClinicName & "'"
            
    Set rs = mConn.CHCSII_DAS(Auto).OpenRecordset(sSQL)
    
    If Not (rs.EOF And rs.BOF) Then
      rs.MoveFirst
      ClinicNCID = rs.Fields(0)
      BackfillClinic = True
      Exit Function
    End If
    
    With mConn.CHCSII_DAS(Auto)
        .ExecuteSQL "insert into clinic (ncid, facility_ncid, location_type,meprs_code, meprs_ien, name) values (" _
                & ClinicNCID & "," & FacilityNCID & ",'CLINIC'," & meprs_code & "," & ClinicNCID & ",'" & ClinicName & "')"
    
        .ExecuteSQL "insert into ncids (ncid, domainncid, description) values (" _
                & ClinicNCID & ",1003,'" & ClinicName & "')"

        .ExecuteSQL "insert into meprs_code(facility_ncid, ien, code, description) values (" _
                & FacilityNCID & "," & ClinicNCID & "," & meprs_code & ",'" & ClinicName & "')"
    End With
                
    BackfillClinic = True

    Exit Function
ErrHandler:

    Err.Raise Err.Number, MODULE & ".BackfillClinic", Err.Description
  
End Function


Private Function FindClinic(ByVal FacilityNCID As Long, ByVal ClinicNCID As Long) As Boolean
    '_ by ncid only

    Dim rs As ADODB.Recordset   'recordset to catch query results
    
    On Error GoTo Func_Err
    
    If ClinicNCID = 0 Then
        FindClinic = True
        Exit Function
    End If
    
    Set rs = mConn.CHCSII_DAS(Auto).OpenRecordset("Select ncid from clinic where facility_ncid = " & FacilityNCID _
           & " and ncid = " & ClinicNCID)
    If Not rs.EOF Then
        FindClinic = True
    Else
        FindClinic = False
    End If
    
Func_Exit:
    On Error Resume Next
    rs.Close
    Set rs = Nothing
    Exit Function

Func_Err:
    Err.Raise Err.Number, MODULE & ".FindClinic", Err.Description
    Resume Func_Exit
End Function


Private Function GetNextID(Optional ByVal eSeqName As ICHCSII_SQL.eSeq = 0&) As Variant
'-- pulls the next value from the MainSeq sequence
    On Error GoTo ErrHandler
    
    Dim rs As ADODB.Recordset       'recordset to catch query results
    Dim sSeqName As String          'sequence name used when getting nextval from dual
    
    sSeqName = SeqEnumToString(eSeqName)
    
    Set rs = mConn.CHCSII_DAS(Auto).OpenRecordset("Select " & sSeqName & ".nextval from dual")

    If Not rs.EOF Then
        GetNextID = rs(0)
    End If
    
Func_Exit:
    On Error Resume Next
    rs.Close
    Set rs = Nothing
    Exit Function

ErrHandler:
    '<< SCR #44407;   Developer: Brian Mowbray 01/13/2004 02:10 PM
    Err.Raise Err.Number, MODULE & ".GetNextID", Err.Description
    Resume Func_Exit

End Function

Private Function SeqEnumToString(ByVal eiSeq As eSeq) As String
    On Error GoTo ErrHandler
    
    Select Case eiSeq
        Case esMAINSEQ
            SeqEnumToString = "MainSeq"
        Case esAPPT_ID
            SeqEnumToString = "APPT_ID"
        Case esAPPT_TRANS_ID
            SeqEnumToString = "APPT_TRANS_ID"
        Case esENC_TEMPLATE_ID
            SeqEnumToString = "ENC_TEMPLATE_ID"
        Case esENCOUNTER_ID
            SeqEnumToString = "ENCOUNTER_ID"
        Case esIEN
            SeqEnumToString = "IEN"
        Case esPCM_CODE
            SeqEnumToString = "PCM_CODE"
        Case esPI_DATA_ID
            SeqEnumToString = "PI_DATA_ID"
        Case esUNIT_NUMBER
            SeqEnumToString = "UNIT_NUMBER"
        Case esUSER_NUMBER
            SeqEnumToString = "USER_NUMBER"
        Case esJDOE_NUMBER
            SeqEnumToString = "JDOE_NUMBER"
        Case esTEMP_SSN
            SeqEnumToString = "TEMP_SSN"
    End Select
    
    Exit Function
    
ErrHandler:
    Err.Raise Err.Description, MODULE & ".SeqEnumToString", Err.Description

End Function


'**************************************************
'SCR 71415
'This Function is being used from DataManager
' When importing Allergies data from the EIC file.
'**************************************************
Private Function iMapperAllergies_RecordsetTransferToLDDB(ByRef col As Collection) As Long
    Dim lReturn As Long                          'return value for this function
    Dim RSObject    As idmbe.I_RSObjects         'DataManagerBacked RSObject used during transfer
    Dim RSAllergies As idmbe.I_RSObjects         'DataManagerBacked RSObject used during transfer
    Dim RSAllergyReaction As idmbe.I_RSObjects   'DataManagerBacked RSObject used during transfer
    Dim RSAllergyComment As idmbe.I_RSObjects     'DataManagerBacked RSObject used during transfer
    Dim dtoAllergy As DTOs.Allergy               'Allergy DTO used to iterate through Allergies DTO
    Dim dtoAllergies As DTOs.Allergies           'Allergies DTO used to retrieve data
    Dim bAddAllergy As Boolean                   'flag... should the allergy be added?
    Dim param As DM_Allergies.AllergySearchParam 'AllergySearchParam used to retrieve data
    Dim i As Long                                'loop index/counter
    
    Dim sFacilityNCID As String
    Dim sFacilityName As String
    Dim lClinicNCID As Long
    Dim sClinicName As String
    Dim sClinicianNCID As String
    
    On Error GoTo ErrHandler
    
    lReturn = 0
    
    For Each RSObject In col
        With RSObject
            If .TableName = "ALLERGIES" Then
                Set RSAllergies = RSObject
            ElseIf .TableName = "ALLERGYREACTION" Then
                Set RSAllergyReaction = RSObject
            ElseIf .TableName = "ALLERGYCOMMENT" Then
                Set RSAllergyComment = RSObject
            End If
        End With
    Next
    
    If Not RSAllergies Is Nothing Then
      RSAllergies.DBRecordSet.MoveFirst
    End If
    
    If Not RSAllergyReaction Is Nothing Then
      RSAllergyReaction.DBRecordSet.MoveFirst
    End If
    
    If Not RSAllergyComment Is Nothing Then
      RSAllergyComment.DBRecordSet.MoveFirst
    End If
    
    
    If Not RSAllergies Is Nothing Then
        'Validate Facility
                
        sFacilityNCID = RSAllergies.DBRecordSet.Fields("FACILITY_NCID") & vbNullString
        sFacilityName = vbNullString
        lClinicNCID = 0
        sClinicName = vbNullString
        
        Call ValidateFacility(sFacilityNCID, sFacilityName, True, lClinicNCID, sClinicName)
        
        sFacilityNCID = RSAllergies.DBRecordSet.Fields("FACILITY_NCID")
        
        sClinicianNCID = RSAllergies.DBRecordSet.Fields("CLINICIANNCID")
        
        'Get LDDB Allergies
        
        Set param = New DM_Allergies.AllergySearchParam
        With param
            .IncludeComment = True
            .IncludeReactions = True
            .PatientUnitNumber = RSAllergies.DBRecordSet.Fields("UNIT_NUMBER")
        End With
        
        Set dtoAllergies = iMapperAllergies_RetrieveAllergies(param)
    End If
        
    Do Until RSAllergies.DBRecordSet.EOF
        
        With RSAllergies
               
            bAddAllergy = True
            'See if Allergy Exists in LDDB
            For Each dtoAllergy In dtoAllergies
                If dtoAllergy.AllergenNCID = .DBRecordSet.Fields("ALLERGENNCID") Then
                    'Allergy exists, check create time
                    If dtoAllergy.CreateTime > .DBRecordSet.Fields("CREATETIME") Then
                        'LDDB Allergy is newer
                        bAddAllergy = False
                    ElseIf dtoAllergy.CreateTime = .DBRecordSet.Fields("CREATETIME") Then
                      
                      'LDDB Allergy has been updated more recently
                      If dtoAllergy.ModifyTime >= .DBRecordSet.Fields("MODIFYTIME") & vbNullString Then
                        bAddAllergy = False
                      Else
                        'SCR 136325 Update the Allergy if the EIC version is newer
                        CopyAllergyInfo .DBRecordSet, dtoAllergy
                        GetReactionsAndComments RSAllergies, RSAllergyReaction, RSAllergyComment, dtoAllergy
                        
                        dtoAllergy.UnitNumber = param.PatientUnitNumber
                        
                        If dtoAllergy.FacilityNCID = vbNullString Then
                          dtoAllergy.FacilityNCID = sFacilityNCID
                        End If
                        
                        If dtoAllergy.ClinicianNCID = vbNullString Then
                          dtoAllergy.ClinicianNCID = sClinicianNCID
                        End If
                        
                        If dtoAllergy.DataId = vbNullString Then
                          dtoAllergy.DataId = 0
                        End If
                        Call iMapperAllergies_SaveAllergy(dtoAllergy)
                        bAddAllergy = False
                      End If

                    End If
                End If
            Next
            
            Set dtoAllergy = Nothing
            If bAddAllergy Then
            
                Set dtoAllergy = New DTOs.Allergy
                
                CopyAllergyInfo .DBRecordSet, dtoAllergy
                
                GetReactionsAndComments RSAllergies, RSAllergyReaction, RSAllergyComment, dtoAllergy
                
                If dtoAllergies Is Nothing Then
                    Set dtoAllergies = New DTOs.Allergies
                End If
                dtoAllergies.Add dtoAllergy
            End If
            
            .DBRecordSet.MoveNext
        End With
        
        If Not dtoAllergies Is Nothing Then
            For i = 1 To dtoAllergies.count
                
                dtoAllergies.Item(i).UnitNumber = param.PatientUnitNumber
                
                If dtoAllergies.Item(i).FacilityNCID = vbNullString Then
                  dtoAllergies.Item(i).FacilityNCID = sFacilityNCID
                End If
                
                If dtoAllergies.Item(i).ClinicianNCID = vbNullString Then
                  dtoAllergies.Item(i).ClinicianNCID = sClinicianNCID
                End If
                
                If dtoAllergies.Item(i).DataId = vbNullString Then
                  dtoAllergies.Item(i).DataId = 0
                End If
                
                
                If iMapperAllergies_SaveAllergy(dtoAllergies.Item(i)) = False Then
                End If
            Next
        End If
    
    Loop
    
    iMapperAllergies_RecordsetTransferToLDDB = lReturn
    
    Exit Function
    
ErrHandler:

   '---show the VB error---
    Err.Raise Err.Description, MODULE & ".iMapperAllergies_RecordsetTransferToLDDB", Err.Description
    iMapperAllergies_RecordsetTransferToLDDB = -1
End Function

Private Sub CopyAllergyInfo(ByRef FromRecordset As ADODB.Recordset, ByRef ToDTO As DTOs.Allergy)
    On Error GoTo ErrHandler
    
    ToDTO.Allergen = FromRecordset("ALLERGEN") & vbNullString
    ToDTO.AllergenNCID = FromRecordset("ALLERGENNCID") & vbNullString
    ToDTO.AllergyTypeNCID = FromRecordset("ALLERGYTYPENCID") & vbNullString
    ToDTO.CDRDataID = FromRecordset("CDR_DATAID") & vbNullString
    ToDTO.ClinicianNCID = FromRecordset("CLINICIANNCID") & vbNullString
    ToDTO.CreateTime = FromRecordset("CREATETIME") & vbNullString
    
    'SCR 163958
    ToDTO.CreateUserNCID = FromRecordset("createuserncid") & vbNullString
    'ToDTO.CreateUserNCID = FromRecordset("CREATEDBY") & vbNullString
    
    ToDTO.EncounterFacilityNCID = FromRecordset("FACILITY_NCID") & vbNullString
    ToDTO.InfoSourceNCID = FromRecordset("INFOSOURCENCID") & vbNullString
    ToDTO.Modified = FromRecordset("MODIFIED") & vbNullString
    ToDTO.ModifyTime = FromRecordset("MODIFYTIME") & vbNullString
    ToDTO.ModifyUserNCID = FromRecordset("MODIFYUSERNCID") & vbNullString
    ToDTO.NotedDate = FromRecordset("NOTEDDATE") & vbNullString
    ToDTO.PointOfCareFacilityNCID = FromRecordset("PointOfCareFacilityNCID") & vbNullString
    ToDTO.UpDateFlag = FromRecordset("UPDATEFLAG") & vbNullString
    Exit Sub

ErrHandler:

    Err.Raise Err.Description, MODULE & ".CopyAllergyInfo", Err.Description

End Sub


Private Sub GetReactionsAndComments(ByRef RSAllergies As idmbe.I_RSObjects, ByRef RSAllergyReaction As idmbe.I_RSObjects, ByRef RSAllergyComment As idmbe.I_RSObjects, ByRef dtoAllergy As DTOs.Allergy)
    On Error GoTo ErrHandler
    Dim tmpReaction As DTOs.AllergyReaction
    Dim count As Integer
    'Get Reaction
    RSAllergyReaction.DBRecordSet.MoveFirst
    
   'SCR 136325 remove any existing reactions
   For count = dtoAllergy.AllergyReactions.count To 1 Step -1
      dtoAllergy.AllergyReactions.Remove (count)
   Next count
   
    Do Until RSAllergyReaction.DBRecordSet.EOF
        If RSAllergies.DBRecordSet("ALLERGENNCID") = RSAllergyReaction.DBRecordSet("ALLERGENNCID") Then
            Set tmpReaction = New DTOs.AllergyReaction
            
            tmpReaction.Name = RSAllergyReaction.DBRecordSet("REACTION")
            tmpReaction.NCID = RSAllergyReaction.DBRecordSet("REACTIONNCID")
            dtoAllergy.AllergyReactions.Add tmpReaction
            'Allow multiple reactions to be read in
'            Exit Do
        End If
    
        RSAllergyReaction.DBRecordSet.MoveNext
    Loop
    
    'Get Comment
    
    If Not RSAllergyComment Is Nothing Then
    
        RSAllergyComment.DBRecordSet.MoveFirst
        
        Do Until RSAllergyComment.DBRecordSet.EOF
        
            If RSAllergies.DBRecordSet("ALLERGENNCID") = RSAllergyComment.DBRecordSet("ALLERGENNCID") Then
                dtoAllergy.Comment = RSAllergyComment.DBRecordSet("ALLERGYCOMMENT")
                Exit Do
            End If
        
            RSAllergyComment.DBRecordSet.MoveNext
        Loop
    
    End If
    
    Exit Sub
ErrHandler:

    Err.Raise Err.Description, MODULE & ".GetReactionsAndComments", Err.Description

End Sub

