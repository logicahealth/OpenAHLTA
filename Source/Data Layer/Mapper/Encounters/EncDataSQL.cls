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
Attribute VB_Name = "EncDataSQL"
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

Private Function iEncDataMapper_EncounterList(Params As DL_Support.iParameters) As DTOs.Encounter
    Dim colEncounters   As Collection
    Dim objEncounter    As Encounter         '- an encounter object
    Dim colDataStrings  As Collection
    Dim objDAS          As ICHCSII_DAS
    Dim sSql            As String
    Dim objRS           As ADODB.Recordset
    Dim oEncProv        As EncProvider
    Dim sMsgText        As String
    Dim sApptIEN        As String
    Dim bFound          As Boolean
    Dim sbuf()          As String
    Dim sClinicianNCID  As String
    Dim objAppts        As Object
    Dim objClinicOps    As Object
    Dim param           As DM_Encounters.ParamsEncList
    Dim RetrieveParam   As DM_Encounters.ParamsEncRetrieve
    Dim Lite            As Boolean
    
    On Error GoTo ErrHandler

    Set param = Params
    
    'date:01/09/2004 SCR #:47870 developer: jrm Description: Added for ITT mode
    'Auto is needed to determine ITT, or Theatre

    Set objDAS = mConn.CHCSII_DAS(Auto) '''<SCR 36493
    
    If param.Grouping = ePrev Then
        'PERF Return Prev Enocunters only SCR-40046
        ' Statuses needed by Prev Encounter
        'Case complete, NeedsCoSignature, Updated, PatientLeftWithoutBeingSeen, _
        '     CancelByFacility, CancelByPatient, NoShow, Updating, CancelByProvider
        sSql = "select EncounterNumber, FacilityNCID from Encounters where unit_number = " & param.PatientUnitNumber _
                  & " AND STATUS IN ( 14510306, 14510309, 205657, 14510310, 14508309, 14505856, 14613990, 665, 160595 ) and TYPE <> " & InpatientNote
    ElseIf param.Grouping = eInpatientNotes Then
        sSql = "select EncounterNumber, FacilityNCID from Encounters where unit_number = " & param.PatientUnitNumber & " and TYPE = " & InpatientNote
    Else
        'Return all encounters SCR#43002
        sSql = "select EncounterNumber, FacilityNCID from Encounters where unit_number = " & param.PatientUnitNumber
    End If
    
    If mMode = LGS Then
        'PERF do not fully load encounters in PGUI mode SCR-40046
        Lite = True
    End If
    
    Set objRS = objDAS.OpenRecordset(sSql)
    
    Set colEncounters = New Collection
    Do While Not objRS.EOF
        'Set objEncounter = New Encounter
        'objEncounter.Init mobjUser, mobjComm, mobjShared, mobjLogon, mobjPatient, mobjProperties
        Set RetrieveParam = New DM_Encounters.ParamsEncRetrieve
        
        RetrieveParam.EncounterNumber = objRS("EncounterNumber")
        RetrieveParam.FacilityNCID = objRS("facilityncid")
        RetrieveParam.IncludeChildModuleData = Lite
        RetrieveParam.PatientUnitNumber = param.PatientUnitNumber
        Set objEncounter = iEncDataMapper_Retrieve(RetrieveParam)
        If Not objEncounter Is Nothing Then
            On Error Resume Next
            colEncounters.Add objEncounter, CStr(objEncounter.id)
            On Error GoTo ErrHandler
        End If
        
        objRS.MoveNext
    Loop   ' while not objsql.eof
    
    Exit Function
    
ErrHandler:

End Function

Private Sub iEncDataMapper_Init(Conn As CHCSII_CONN.Conn, Cache As DL_Support.CacheMapper, ByVal DataAccessMode As CHCSII_CONN.eBackend, Optional ByVal CacheConfigItems As Boolean = False)
  Set mConn = Conn
  Set mCache = Cache
  
  mMode = DataAccessMode
  mUseCache = CacheConfigItems
End Sub

Private Function iEncDataMapper_Retrieve(Params As DL_Support.iParameters) As DTOs.Encounter
    Dim objDAS As ICHCSII_DAS
    Dim param As DM_Encounters.ParamsEncRetrieve
    
    Dim objRS As ADODB.Recordset
    Dim sSql As String
    Dim objProvider     As EncProvider

    Dim objConcept      As GEMSConceptCtrl.GEMS_ConceptCtrl
    Dim sName           As String
    Dim sValue          As String
    
    Dim oSQLEx As ICHCSII_SQL.ISqlOps
    Dim oSQLEx2 As ICHCSII_SQL.ISqlOpsEx
    
    Dim Encounter As Encounter
    
    Set Encounter = New Encounter
    
    Set param = Params
    
    'date:01/14/2004 SCR #:47870 developer: jrm Description: Added for ITT mode
    'Auto is needed to determine ITT, or Theatre; Mode needed is theatre logic gate

    Set objDAS = mConn.CHCSII_DAS(Auto)

    'date:01/14/2004 SCR #:47870 developer: jrm Description: Added for ITT mode
    'Auto is needed to determine ITT, or Theatre; Mode needed is theatre logic gate

    Set oSQLEx = mConn.CHCSII_SQLOPS(Auto) '''<SCR 36493

    On Error GoTo ErrHandler
    
    Set objConcept = New GEMSConceptCtrl.GEMS_ConceptCtrl
    
    With Encounter

   'Build the encounters select statement
      sSql = "select E.* from Encounters E where E.EncounterNumber = " & param.EncounterNumber _
            & " and E.FacilityNCID = " & param.FacilityNCID & " and E.Unit_Number = " & param.PatientUnitNumber
   
    oSQLEx.Execute sSql
    If oSQLEx.EOF Then
        If mMode = GEMS Or mMode = MYSQL Then
            'RAISE ERROR
            'ZZZ
            'Err.RAISE 100000, "Encounters DL",  "Encounter data does not contain primary event.", vbExclamation, ExeModule
        End If
        GoTo Cleanup
    End If
    
    .DataID = oSQLEx("EncounterNumber")
    .id = oSQLEx("EncounterNumber")
    If .id = 0 Then
        'RAISE ERROR
        'ZZZ
        'Err.RAISE 100000, "Encounters DL",  "Encounter data contains a zero encounter id.", vbExclamation, ExeModule
        GoTo Cleanup
    End If

    .PatientUnitNum = "" & oSQLEx("Unit_Number")
    .FacilityNCID = "" & oSQLEx("FacilityNCID")
    objConcept.uniqueID = .FacilityNCID
    
    .FacilityName = objConcept.PrefRep("2000").Representation
    
    .ClinicNCID = "" & oSQLEx("ClinicNCID")
    objConcept.uniqueID = .ClinicNCID
    
    .ClinicName = objConcept.PrefRep("2000").Representation
      
    Set objProvider = New EncProvider
    objProvider.NCID = "" & oSQLEx("primaryproviderncid")
    objConcept.uniqueID = objProvider.NCID
    
    objProvider.FullName = objConcept.PrefRep("2000").Representation
    objProvider.RoleNCID = rtnAttendingProvider             'Set default provider role in case this encounter was created with a non HIPPA version
    .Providers.Add objProvider, CStr(objProvider.NCID)
   
    If "" & oSQLEx("UpdatedBy") = "" Then
        Set objProvider = Nothing
        Set objProvider = New EncProvider
        objProvider.NCID = "" & oSQLEx("CreatedBy")
        objConcept.uniqueID = objProvider.NCID
        'objProvider.FullName = objConcept.PrefRep("2000").Representation
    Else
        Set objProvider = Nothing
        Set objProvider = New EncProvider
        objProvider.NCID = "" & oSQLEx("UpdatedBy")
        objConcept.uniqueID = objProvider.NCID
        'objProvider.FullName = objConcept.PrefRep("2000").Representation
    End If

    If objProvider.NCID <> "" & oSQLEx("primaryproviderncid") And objProvider.NCID <> 0 Then
    
        objProvider.FullName = objConcept.PrefRep("2000").Representation
        
        .Providers.Add objProvider, CStr(objProvider.NCID)
    End If
    Set objProvider = Nothing


    '<< tfb change to handle null value 07-11-2001
    If oSQLEx("UpdatedOn") <> "" Then
        .UpdateDate = "" & oSQLEx("UpdatedOn")  'note:  msUpdateDate is a date data type
    Else
        .UpdateDate = CDate(0)
    End If

    If Not oSQLEx("UpdatedBy") = 0 Then
        .UpdateClinicianNCID = "" & oSQLEx("UpdatedBy")
    Else
        .UpdateClinicianNCID = vbNullString
    End If

    If Not oSQLEx("CreatedBy") = 0 Then 'SCR#42996
        .CreateClinicianNCID = "" & oSQLEx("CreatedBy")
    Else
        .CreateClinicianNCID = vbNullString
    End If
    
    If oSQLEx("StartDTS") <> "" Then
        .CreatedDate = oSQLEx("StartDTS")
    Else
       .CreatedDate = CDate(0)
    End If

    If oSQLEx("EndDTS") <> "" Then
        .WhenSignedDate = oSQLEx("EndDTS")
    Else
        .WhenSignedDate = CDate(0)
    End If
    
    If oSQLEx("APPTDTS") <> "" Then
        .Appointment.DateTime = oSQLEx("APPTDTS")
    Else
        .Appointment.DateTime = CDate(0)
    End If

'<< SCR  10086 moved from below to have this set before status is updated
    If Val("" & oSQLEx("NOCOUNT")) > 0 Then
        .Class = ec_NoCount
    End If
'>>
    .Status = StatusNCIDToEnum("" & oSQLEx("Status"))
    .WasUpdated = (oSQLEx("WasUpdated") = "Y")
    'Amending does not clear who signed NCID so do not load instead
    'If meStatus <> Updating Then
        If oSQLEx("WHENCOSIGNED") <> "" Then
            .WhenCosigned = oSQLEx("WHENCOSIGNED")
        End If
        Set .WhoCosigned = New EncProvider
        .WhoCosigned.NCID = "" & oSQLEx("WHOCOSIGNEDNCID")
        If Not .WhoCosigned.NCID = 0 Then
            objConcept.uniqueID = .WhoCosigned.NCID
            .WhoCosigned.FullName = objConcept.PrefRep("2000").Representation
            .WhoCosigned.SigLine2 = "" & oSQLEx("COSIGNERLINE2")
            .WhoCosigned.SigLine3 = "" & oSQLEx("COSIGNERLINE3")
        End If
        .WhenSignedDate = CDate("" & oSQLEx("WhenSigned"))
        Set .WhoSigned = New EncProvider
        .WhoSigned.NCID = "" & oSQLEx("WHOSIGNEDNCID")
        If Not .WhoSigned.NCID = 0 Then
            objConcept.uniqueID = .WhoSigned.NCID
            .WhoSigned.FullName = objConcept.PrefRep("2000").Representation
            .WhoSigned.SigLine2 = "" & oSQLEx("SIGNERLINE2")
            .WhoSigned.SigLine3 = "" & oSQLEx("SIGNERLINE3")
        Else
            Set .WhoSigned = Nothing
        End If
    'End If
    '.Category = "" & oSQLEx("Category")
    '.TimeZone = "" & oSQLEx("TIMEZONE")
    .ENCTYPE = "" & oSQLEx("TYPE")
'    If IsDate(oSQLEx("ALLERGYVERIFIEDBYDATE")) Then
'        mdAllergyVerifiedByDate = oSQLEx("ALLERGYVERIFIEDBYDATE")
'    End If
'    msAllergyVerifiedByNCID = "" & oSQLEx("ALLERGYVERIFIEDBYNCID")

    If Not oSQLEx("CoSignerNCID") = 0 Then
    '- if the one who is to cosign has cosigned, don't add him again to the list
        If Not "" & oSQLEx("CoSignerNCID") = "" & oSQLEx("WHOCOSIGNEDNCID") Then
            Set .Cosigner = New EncProvider
            .Cosigner.NCID = "" & oSQLEx("CoSignerNCID")
            objConcept.uniqueID = .Cosigner.NCID
            .Cosigner.FullName = objConcept.PrefRep("2000").Representation
        End If
    End If

    If Not oSQLEx("AdditionalProvider1NCID") = 0 Then
        Set objProvider = New EncProvider
        objProvider.NCID = "" & oSQLEx("AdditionalProvider1NCID")
        objProvider.RoleNCID = "" & oSQLEx("APROLE1NCID") 'SCR-19027 'Wrong column name oSQLEx("APROLE1")
        
        If mMode <> LGS Then
            objConcept.uniqueID = objProvider.NCID
            objProvider.FullName = objConcept.PrefRep("2000").Representation
            objConcept.uniqueID = objProvider.RoleNCID
            objProvider.Role = objConcept.PrefRep("2000").Representation
        Else
            objConcept.uniqueID = objProvider.NCID
            objProvider.FullName = objConcept.PrefRep("2000").Representation
            
            Select Case Trim$(objProvider.RoleNCID)
                Case "4"
                    objProvider.Role = "Nurse"
                Case "2"
                    objProvider.Role = "Assisting"
                Case "3"
                    objProvider.Role = "Supervising"
                Case "5"
                    objProvider.Role = "Para-professional"
                Case "1"          '
                    objProvider.Role = "Attending"
            End Select
        End If

        .Providers.Add objProvider, CStr(objProvider.NCID)
    End If
    If Not oSQLEx("AdditionalProvider2NCID") = 0 Then
        Set objProvider = New EncProvider
        objProvider.NCID = "" & oSQLEx("AdditionalProvider2NCID")
        objProvider.RoleNCID = "" & oSQLEx("APROLE2NCID") 'SCR-19027 'Wrong column name oSQLEx("APROLE2")

        If mMode <> LGS Then
            objConcept.uniqueID = objProvider.NCID
            objProvider.FullName = objConcept.PrefRep("2000").Representation
            objConcept.uniqueID = objProvider.RoleNCID
            objProvider.Role = objConcept.PrefRep("2000").Representation
        Else
            objConcept.uniqueID = objProvider.NCID
            objProvider.FullName = objConcept.PrefRep("2000").Representation
            Select Case Trim$(objProvider.RoleNCID)
                Case "4"
                    objProvider.Role = "Nurse"
                Case "2"
                    objProvider.Role = "Assisting"
                Case "3"
                    objProvider.Role = "Supervising"
                Case "5"
                    objProvider.Role = "Para-professional"
                Case "1"          '
                    objProvider.Role = "Attending"
            End Select
        End If
        .Providers.Add objProvider, CStr(objProvider.NCID)
    End If

    '.PrimaryDiagnosisNCID = "" & oSQLEx("PrimarydiagnosisNCID")
    .PrimaryDxSnoID = "" & oSQLEx("DXSnoID")   ' Load Medcin ID & Prefix SCR-19435
    'msPrimaryDxMedcinPrefix = "" & oSQLEx("DXMEDCINPREFIX")
    .SADRStatus = "" & oSQLEx("SADRStatus")
    .Appointment.id = "" & oSQLEx("APPTID")
    

    'Get Template Data
    If mMode = LGS Then
      .TemplateData = gobjShared.CHCSConnection.ReadChunk("ENCOUNTERS", "ENCOUNTERNUMBER = " & .id, "TEMPLATEDATA")
      .Appointment.IEN = .Appointment.id
    Else
        'Theater and ITT mode SCR#49315
        Set oSQLEx2 = mConn.CHCSII_SQLOPS_EX(Auto)
          
        oSQLEx2.Execute "Select TemplateData from Encounters " & vbCrLf & _
        "where ENCOUNTERNUMBER = " & CStr(.id) & vbCrLf & _
        "AND FACILITYNCID = " & .FacilityNCID
        
        If Not (oSQLEx2.BOF And oSQLEx2.EOF) Then
            oSQLEx2.GetAsChunk "TemplateData", .TemplateData
        End If
        
        Set oSQLEx2 = Nothing
    End If
    
    If mMode = CDR Then
      'Application is in CDR mode, set GEMS encounter object UpdateLock value = false
      .UpdateLock = False
    Else
      .UpdateLock = CBool(oSQLEx("UpdateLock"))
    End If
    
    .Appointment.Classification = "" & oSQLEx("Patient_Status")
    
    If mMode = GEMS Or mMode = MYSQL Then
        .InpatientService = "" & oSQLEx("INPATIENT_SERVICE")
    End If
    
    
 '==========================
 '--RTF
   Call GEMS_LoadRTFs(.FacilityNCID, _
                      .id)
 '==========
    End With
    If param.IncludeChildModuleData Then
        'Set mobjLocking = New Locking
        'mobjLocking.msFacilityNCID = FacilityNCID
        'mobjLocking.msEncounterNumber = mnEncounterID
        'mobjLocking.msUserNCID = mobjLogon.UserNCID
'        Call GEMS_CreateSections(param.PatientUnitNumber, _
'                                 param.FacilityNCID, _
'                                 param.EncounterNumber, _
'                                 objDAS)
        'mbFullyLoaded = True
    End If
    
    'GEMS_LoadEncounter = True
    GoTo Cleanup
    
ErrHandler:
    'RASIE ERROR
    'ZZZ
    'Call mobjShared.ShowVBError(Err.Number, Err.Description, "Encounter.GEMS_LoadEncounter ", "Encounter", vbCritical)

Cleanup:
'    mbLoading = False
'    mbDirty = False
    Set objProvider = Nothing
    Set objConcept = Nothing
   ' If objRS.State > 0 Then objRS.Close
    Set objRS = Nothing
    Set objDAS = Nothing
End Function

Private Function iEncDataMapper_Save(Encounter As DTOs.Encounter) As Boolean

   Dim bResult          As Boolean
    Dim i                As Integer
    'Dim objSectionParent As SectionParent
    Dim objEncRTF        As Document
    Dim objDAS           As ICHCSII_DAS
    Dim sql              As String
    Dim sql2             As String
    Dim sTemp            As String
    Dim oSQLEx As ICHCSII_SQL.ISqlOpsEx
    
    On Error GoTo ErrHandler

    
    If Not StartSectionUpdate(eHeader, 0) Then
        'RAISE ERROR
        'Err.RAISE 100000, "Encounters DL",  "Your lock on this encounter was broken by another user, and your changes cannot be saved.", vbInformation, "Encounter"
        GoTo Cleanup
    End If
    
    'date:01/14/2004 SCR #:47870 developer: jrm Description: Added for ITT mode
    'Auto is needed to determine ITT, or Theatre; Mode needed is theatre logic gate
    
    Set objDAS = mConn.CHCSII_DAS(Auto) '''<SCR 36493

    With Encounter

    If mMode = GEMS Or mMode = MYSQL Or mMode = ITT Then
        
    'date:01/14/2004 SCR #:47870 developer: jrm Description: Added for ITT mode
    'Auto is needed to determine ITT, or Theatre; Mode needed is theatre logic gate
        Set oSQLEx = mConn.CHCSII_SQLOPS(Auto)
        
        oSQLEx.Execute "select * from encounters where encounternumber = " & .id
        
        If Not oSQLEx.EOF Then
            oSQLEx("UNIT_NUMBER") = .PatientUnitNum
            oSQLEx("FACILITYNCID") = .FacilityNCID
            oSQLEx("ENCOUNTERNUMBER") = .id
            oSQLEx("CLINICNCID") = .ClinicNCID
            oSQLEx("STATUS") = StatusEnumToNCID(.Status)
            oSQLEx("CATEGORY") = .Appointment.Classification
            oSQLEx("TYPE") = .ENCTYPE
            ' not used SCR-25084
            'oSQLEx("ENDDTS") = mdEndDTS
            oSQLEx("APPTID") = .Appointment.id
            oSQLEx("STARTDTS") = .CreatedDate
            oSQLEx("SADRSTATUS") = .SADRStatus
            oSQLEx.AppendAsChunk "TEMPLATEDATA", .TemplateData
            'oSQLEx("TIMEZONE") = msTimeZone
            oSQLEx("UPDATELOCK") = CInt(.UpdateLock)
            oSQLEx("APPTDTS") = .Appointment.DateTime
        
'            If Trim$(msPrimaryDiagnosisNCID) <> vbNullString Then
'                oSQLEx("PRIMARYDIAGNOSISNCID") = msPrimaryDiagnosisNCID
'            End If
            
            ' Save Medcin ID & Prefix SCR-19435
            oSQLEx("DXSnoID") = .PrimaryDxSnoID
            
'            If Trim$(msPrimaryDxMedcinPrefix) <> vbNullString Then
'                oSQLEx("DXMEDCINPREFIX") = msPrimaryDxMedcinPrefix
'            End If

            If .Providers.Count > 0 Then
                oSQLEx("PRIMARYPROVIDERNCID") = .Providers(1).NCID
                If .Providers.Count > 1 Then
                    oSQLEx("ADDITIONALPROVIDER1NCID") = .Providers(2).NCID
                    oSQLEx("APROLE1NCID") = .Providers(2).RoleNCID
                    If .Providers.Count > 2 Then
                        oSQLEx("ADDITIONALPROVIDER2NCID") = .Providers(3).NCID
                        oSQLEx("APROLE2NCID") = .Providers(3).RoleNCID
                    End If
                End If
        
                If Not .Cosigner Is Nothing Then
                    oSQLEx("COSIGNERNCID") = .Cosigner.NCID
                        oSQLEx("WHENCOSIGNED") = .WhenCosigned
        
                        If Not .WhoCosigned Is Nothing Then
                            oSQLEx("WHOCOSIGNEDNCID") = .WhoCosigned.NCID
                            oSQLEx("COSIGNERLINE2") = .WhoCosigned.SigLine2
                            oSQLEx("COSIGNERLINE3") = .WhoCosigned.SigLine3
                        End If
                Else
                    'Need to clear NCID in case we are amending
                    'oSQLEx("COSIGNERNCID") = vbEmpty
                End If
        
                If Not .WhoSigned Is Nothing Then
                    If .WhenSignedDate <> 0 Then  'SCR-21188
                        oSQLEx("WHENSIGNED") = .WhenSignedDate  ' mConn.SQLdate(mdWhenSigned)
                    End If
                    oSQLEx("WHOSIGNEDNCID") = .WhoSigned.NCID
                    oSQLEx("SIGNERLINE2") = .WhoSigned.SigLine2
                    oSQLEx("SIGNERLINE3") = .WhoSigned.SigLine3
                Else
                    'Need to clear NCID in case we are amending
                    'oSQLEx("WHOSIGNEDNCID") = Null
                End If
                
                If Not Val(.Sections("RFV").AllergyVerifiedByNCID) = 0 Then
                    oSQLEx("ALLERGYVERIFIEDBYNCID") = .Sections("RFV").AllergyVerifiedByNCID   'WRogers removed the ticks, field is number not character string
                End If
                oSQLEx("ALLERGYVERIFIEDBYDATE") = .Sections("RFV").AllergyVerifiedByDate  ' mConn.SQLdate(mdAllergyVerifiedByDate)
            Else
                'RASIE ERROR
                'ZZZ
                'Err.RAISE 100000, "Encounters DL",  "Primary Provider data missing in GEMS_Save for Encounter Number: " & mnEncounterID, vbCritical, "EncounterOps"
                GoTo Cleanup
            End If
            
            oSQLEx("INPATIENT_SERVICE") = .InpatientService
          
            oSQLEx.Update
            
            'ZZZ
            'GEMS_Save = SaveHIPAAData(False)
    
        End If
    Else
        'PGUI MODE
        sql = "update Encounters set " _
            & "UNIT_NUMBER= " & .PatientUnitNum _
            & ",FACILITYNCID= " & .FacilityNCID _
            & ",ENCOUNTERNUMBER= " & .id _
            & ",CLINICNCID= " & .ClinicNCID _
            & ",STATUS= " & StatusEnumToNCID(.Status) _
            & ",CATEGORY= '" & .Appointment.Classification & "'" _
            & ",TYPE= " & .ENCTYPE _
            & ",ENDDTS=" & mConn.SQLDate(.WhenSignedDate) _
            & ",APPTID= '" & .Appointment.id & "'" _
            & ",STARTDTS=" & mConn.SQLDate(.CreatedDate) _
            & ",SADRSTATUS= '" & .SADRStatus & "'" _
            & ",TIMEZONE= '" & .TemplateData & "' " _
            & ",UPDATELOCK = " & CInt(.UpdateLock) & " "  'SCR 17781 ed clarke
        
        sql = sql & ",PATIENT_STATUS= " & mConn.SQLQuote(.Appointment)
            If Trim(.Sections("AnP").Diagnoses(1).NCID) <> vbNullString Then
            End If
    
        If .Providers.Count > 0 Then
            sql = sql & ",PRIMARYPROVIDERNCID= " & .Providers(1).NCID
            If .Providers.Count > 1 Then
                sql = sql & ",ADDITIONALPROVIDER1NCID= " & .Providers(2).NCID _
                          & ",APROLE1NCID= " & .Providers(2).RoleNCID
                If .Providers.Count > 2 Then
                    sql = sql & ",ADDITIONALPROVIDER2NCID= " & .Providers(3).NCID _
                              & ",APROLE2NCID= " & .Providers(3).RoleNCID
                Else
                    'Clear Add Prov 2 SCR#44276
                    sql = sql & ",ADDITIONALPROVIDER2NCID= 0" _
                              & ",APROLE2NCID= 0"
                End If
            Else
                'Clear Addl Prov 1 and 2 SCR#44276
                sql = sql & ",ADDITIONALPROVIDER1NCID= 0" _
                          & ",APROLE1NCID= 0"
                          
                sql = sql & ",ADDITIONALPROVIDER2NCID= 0" _
                          & ",APROLE2NCID= 0"
            End If
    
            'Use SQL Quote for Sig lines SCR#48788
            If Not .Cosigner Is Nothing Then  'tfb 07/06/2001
                sql = sql & ",COSIGNERNCID= '" & .Cosigner.NCID & "'" _
                    & ",WHENCOSIGNED= " & mConn.SQLDate(.WhenCosigned)
    
                    If Not .WhoCosigned Is Nothing Then
                        sql = sql & ",WHOCOSIGNEDNCID= '" & .WhoCosigned.NCID & "'" _
                            & ",COSIGNERLINE2= " & mConn.SQLQuote(.WhoCosigned.SigLine2) & "" _
                            & ",COSIGNERLINE3= " & mConn.SQLQuote(.WhoCosigned.SigLine3) & ""
                    End If
            End If
    
            If Not .WhoSigned Is Nothing Then
                sql = sql & ",WHENSIGNED= " & mConn.SQLDate(.WhenSignedDate) _
                    & ",WHOSIGNEDNCID= " & .WhoSigned.NCID _
                    & ",SIGNERLINE2= " & mConn.SQLQuote(.WhoSigned.SigLine2) & "" _
                    & ",SIGNERLINE3= " & mConn.SQLQuote(.WhoSigned.SigLine3) & ""
            End If
            
            'ITT PROVIDERGUI UPDATES SCR-42759
            If mMode = ITT + LGS And .Sections("RFV").AllergyVerifiedByNCID = "0" Then
                .Sections("RFV").AllergyVerifiedByNCID = ""
            End If
            
'<SCR <DEV
            sql = sql & ",ALLERGYVERIFIEDBYNCID= " & M_sSqlVal(.Sections("RFV").AllergyVerifiedByNCID, mMode <> ITT)
            sql = sql & ",ALLERGYVERIFIEDBYDATE= " & mConn.SQLDate(.Sections("RFV").AllergyVerifiedByDate) _
                    & " Where EncounterNumber = " & .id  'Use Enc # SCR-40830
'>SCR
        Else
            'SCR#48295
            'RASIE ERROR
            'Err.RAISE 100000, "Encounters DL",  "Primary Provider data missing in GEMS_Save for Encounter Number: " & .Id, vbCritical, "EncounterOps"
            GoTo Cleanup
        End If
        objDAS.ExecuteSQL sql
                              
        sTemp = .TemplateData 'Using local variable prevent msTemplateData from being overwritten by AppendChunk method
        
        Set oSQLEx = mConn.CHCSII_SQLOPS(Auto)

        If Len(sTemp) > 0 Then
          oSQLEx.Execute "Select TemplateData from Encounters " & vbCrLf & _
            "where ENCOUNTERNUMBER = " & CStr(.id) & vbCrLf & _
            "AND FACILITYNCID = " & .FacilityNCID
          If Not (oSQLEx.BOF And oSQLEx.EOF) Then
            oSQLEx.AppendAsChunk "TemplateData", sTemp
            oSQLEx.Update
          End If
        End If
        
        Set oSQLEx = Nothing
        
    End If
    
    Dim oSQLOps   As ICHCSII_SQL.ISqlOps
    
    'date:01/14/2004 SCR #:47870 developer: jrm Description: Added for ITT mode
    'Auto is needed to determine ITT, or Theatre; Mode needed is theatre logic gate

    Set oSQLOps = mConn.CHCSII_SQLOPS(Auto) '''<SCR 36493
    
    Dim rtf_index As Long
    Dim objNote   As Document
    Dim objNoteSection As iSection
    
    objDAS.ExecuteSQL "Delete from enc_rtfs where " _
                    & " ENCOUNTERNUMBER = " & .id _
                    & "  and FACILITYNCID = " & .FacilityNCID _
                    & " and rtfType = 'RTF'"
                    
    For Each objEncRTF In .PrevSignedRTFs
        rtf_index = oSQLOps.GetNextID()
        sql2 = ""
        sql2 = "insert into enc_RTFs (FACILITYNCID,ENCOUNTERNUMBER,rtftype,DTS,ENC_RTFSINDEX ) "
        sql2 = sql2 & " values ( " & .FacilityNCID & ", " & .id & ", "

        If mMode = ITT Then 'ITT Change
            sql2 = sql2 & " 'RTF', " & mConn.SQLDate(objEncRTF.LastModifiedDate) & "," & rtf_index & " )"
        Else
            'Format does not work when transferring LDDB->CDR SCR-22939
            sql2 = sql2 & " 'RTF', " & mConn.SQLDate(objEncRTF.LastModifiedDate) & "," & rtf_index & " )"
            'Removed single quotes; SCR 25196
            'sql2 = sql2 & " 'RTF', '" & mConn.SQLdate(objEncRTF.mdDTS) & "'," & rtf_index & " )"
        End If
        
        
        Call objDAS.ExecuteSQL(sql2)
        
 ' ZZZ  SAVE TEXT NOTE
'        Set objNote = New Document
'        Set objNoteSection = objNote
'        objNoteSection.DataId = rtf_index
'        objNote.OwnerName = mobjLogon.UserName
'        objNote.OwnerNCID = objEncRTF.msUserNCID
'        objNote.LastModifiedDate = objEncRTF.mdDTS
'        objNote.body = objEncRTF.msRTF
'
'        objNote.GEMS_WriteToDataString
'        Set objNote = Nothing
        
    Next objEncRTF
    
    'RaiseEvent when Enc Status is saved SCR-21322
    'Do not raise event for Dental encounters SCR-23470, 24189
    'ZZZ RAISE IN APP
'    If .Type <> Dental Then
'        RaiseEvent EncStatusChanged(meStatus)
'    End If
    
    Set oSQLOps = Nothing
'    GEMS_Save = True
 '   mbDirty = False
    
    End With
    GoTo Cleanup
ErrHandler:
    'RASIE ERROR
   'Call mobjShared.ShowVBError(Err.Number, Err.Description, "EnctrGems.Save ", "Enc", vbCritical)
Cleanup:
    Call EndSectionUpdate(eHeader, 0)
    
    #If debugon Then
        If Not DL Is Nothing Then DL.DebugLog Exe, Module, "Save", "End", DR
    #End If
    
   
    'If GEMS_Save Then
    '    mobjComm.SetStatusText "Encounter " & mnEncounterID & " was saved."
    'Else
    '    mobjComm.SetStatusText ""
    'End If
    
    Set objDAS = Nothing



End Function


Friend Function GEMS_LoadProviders(ByVal lFacilityNCID As String, _
                               ByVal lEncounterNumber As String) As Boolean
                               
On Error GoTo ErrorHandler

Dim sSql As String
Dim oSQL As ICHCSII_SQL.ISqlOps
Dim oProvider As EncProvider
Dim oProcedure As EncProvProcedure
Dim cProviders As EncProviders

   'Initialize new collection
   Set cProviders = New EncProviders

   'Build the SQL - This recordset will be the combination of providers and procedures
   If gobjShared Is Nothing Then Set gobjShared = New CWShared
   If gobjShared.UseSQLServer Then
      'Build SQL using SQLServer Syntax
      sSql = "SELECT P.Name, EP.Provider_NCID, P.IEN as Provider_IEN, EP.Role_Type_NCID, PP.DataId as ProcedureDataId, " & _
            " PR.SnoID, RT.Role_Type_Description as Role_Desc, EP.Priority_Indicator " & _
            " FROM Encounter_Provider EP " & _
            " INNER JOIN Provider_Role_Type RT On EP.Role_Type_NCID = RT.Role_Type_NCID " & _
            " INNER JOIN Provider P ON EP.Provider_NCID = P.NCID " & _
            " LEFT OUTER JOIN Encounter_Procedure PP on EP.EncounterNumber = PP.EncounterNumber " & _
               " AND EP.FacilityNCID = PP.FacilityNCID AND EP.Provider_NCID = PP.Provider_NCID " & _
            " LEFT OUTER JOIN Procedures PR on PP.DataId = PR.DataId " & _
            " WHERE EP.EncounterNumber = " & lEncounterNumber & " AND EP.FacilityNCID = " & lFacilityNCID & _
            " ORDER BY EP.Priority_Indicator ASC"
   Else
      sSql = "SELECT P.Name, EP.Provider_NCID, P.IEN as Provider_IEN, EP.Role_Type_NCID, PP.DataId as ProcedureDataId, " & _
            " PR.SnoID, RT.Role_Type_Description as Role_Desc, EP.Priority_Indicator " & _
            " FROM Encounter_Provider EP, Encounter_Procedure PP, Provider_Role_Type RT, Provider P, Procedures PR " & _
            " WHERE EP.EncounterNumber = " & lEncounterNumber & " AND EP.FacilityNCID = " & lFacilityNCID & _
            " AND EP.EncounterNumber = PP.EncounterNumber(+) AND EP.FacilityNCID = PP.FacilityNCID(+) " & _
            " AND EP.Provider_NCID = PP.Provider_NCID(+) " & _
            " AND EP.Role_Type_NCID = RT.Role_Type_NCID " & _
            " AND EP.Provider_ncid = P.ncid " & _
            " AND PP.DataId = PR.DataId(+) " & _
            " ORDER BY Priority_Indicator ASC "
   End If
      
   Set oSQL = mConn.CHCSII_SQLOPS(Auto)
   oSQL.Execute sSql
    
   Do While Not oSQL.EOF
   
      'Create and load the provider
      Set oProvider = New EncProvider
      
      'Load the provider
      oProvider.PrevNCID = oSQL("Provider_NCID") & ""
      oProvider.NCID = oSQL("Provider_NCID") & ""
      oProvider.IEN = oSQL("Provider_IEN") & ""
      oProvider.FullName = oSQL("Name") & ""
      oProvider.RoleNCID = oSQL("Role_Type_NCID") & ""
      oProvider.Role = oSQL("Role_Desc") & ""
      'oProvider.State =
      
      'Create and load the procedure associations
      Do
         If Not IsNull(oSQL("ProcedureDataId")) Then
            'Add a procedure to the provider - make sure we have a valid data id since ITT mode a null DB value will be returned as a zero
            If Val(oSQL("ProcedureDataId")) > 0 Then
               Set oProcedure = New EncProvProcedure
               oProcedure.DataID = oSQL("ProcedureDataId")
               oProcedure.SnoID = oSQL("SnoID")
               oProvider.Procedures.Add oProcedure
               Set oProcedure = Nothing
            End If
         End If
         oSQL.MoveNext
         If oSQL.EOF Then Exit Do
      Loop Until oSQL("Provider_NCID") & "" <> CStr(oProvider.NCID)
      
      'Add the provider to the collection
      cProviders.Add oProvider, CStr(oProvider.NCID)
            
      Set oProvider = Nothing
      
   Loop
      
   If cProviders.Count > 0 Then
      'For backwards compatibility - only update encounter collection if we found provider information in new structure
      'ZZZ
      'Set .Providers = Nothing
      'Set .Providers = cProviders
   End If
   
ErrorHandler:
   Set oSQL = Nothing
   Set oProvider = Nothing
   Set oProcedure = Nothing
   Set cProviders = Nothing
   
   'If CBool(Err) Then HandleErrors Err, MODULE_NAME, "GEMS_LoadProviders"
           
End Function
Private Function GEMS_LoadRTFs(ByVal FacilityNCID As String, _
                               ByVal EncounterNumber As String) As Boolean
    Dim objEncRTF As Note
    Dim objSQL    As ICHCSII_SQL.ISqlOps
    Dim oEncRTFSec As iSection

    'date:01/14/2004 SCR #:47870 developer: jrm Description: Added for ITT mode
    'Auto is needed to determine ITT, or Theatre; Mode needed is theatre logic gate

    Set objSQL = mConn.CHCSII_SQLOPS(Auto) '''<SCR 36493
    
    Dim sSql      As String
    Dim aDoc      As Variant
    
    Dim oShared   As CWShared
    
    Set oShared = New CWShared
    
    sSql = "Select * from Enc_RTFs where EncounterNumber = " & EncounterNumber _
        & " and FacilityNCID = " & FacilityNCID _
        & " and rtftype =  'RTF' " _
        & " Order by DTS DESC"
    objSQL.Execute sSql
    
    Set objSQL = Nothing
    
End Function



