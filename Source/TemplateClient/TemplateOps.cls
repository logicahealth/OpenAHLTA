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
Attribute VB_Name = "TemplateOps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'<< GEMS    '*-- Coleman 060401: Add Theater support
'Private mobjSQL As CHCSII_SQL.SqlOps
Private mobjSQL As ICHCSII_SQL.ISqlOps
Private moConn As CHCSII_CONN.Conn
Private moMTFs As ADODB.Recordset

Private mobjConcept_GEMS As GEMSConceptCtrl.GEMS_ConceptCtrl
'GEMS >>
Private mobjConcept As Object 'SCR-53524 5/10/2004 Need to be generic for theater

Public Enum TemplateTypeEnum
    EncounterTemplateType = 0
    ListTemplateType = 1
    FormTemplateType = 2  'emc Forms support
    AllTemplateTypes = 3
    Both_Encounter_And_List_Type
    Both_Encounter_And_Form_Type
    Both_Form_And_List_Type
End Enum

'
'--- Tracker # 10695 - Add a parameter call to get individual sections of the template.
'--- Added the Enum Type 'TemplateSectionEnum' to the calls for 'GetTemplate'
'--- and 'NewTemplate' - RP
'
Public Enum TemplateSectionEnum
    ComplaintListTemplateSection = 0
    AppointmentListTemplateSection
    ProblemListTemplateSection
    AutociteListTemplateSection
    ClipartListTemplateSection
    NotesListTemplateSection
    DiagnosisListTemplateSection
    OrderListTemplateSection
    ProcedureListTemplateSection
    InstructionsTemplateSection
    FormsTemplateSection 'emc Forms support
    InvalidTemplateSection = 99
End Enum

Public Enum TemplateOwnerTypeEnum
    NoOwnerType = 0
    PersonalOwnerType = 1
    ClinicOwnerType = 2
    MTFOwnerType = 3
    EnterpriseOwnerType = 4
End Enum

Public Enum TemplateNameSearchEnum
    BeginsWithNameSearch = 0
    ContainsNameSearch = 1
End Enum

Private msSearchID As String
Private msSearchName As String
Private meSearchNameType As TemplateNameSearchEnum
Private mbSearchNameCaseInsensitive As Boolean
Private mbSearchShared As Boolean
Private meSearchOwnerType As TemplateOwnerTypeEnum
Private mcSearchOwners As Collection
Private mcSearchTemplateType As Collection 'SCR-49343 RBELL 2/9/4 Added template type support
Private msSearchSpecialtyID As String
Private moSearchComplaints As TemplateData
Private moSearchAppointmentTypes As TemplateData
Private moSearchProblems As TemplateData
Private mbSearchFavorites As Boolean
Private mbSearchAdvanced As Boolean
Private moTemplateHxOps As TemplateHx

Private mbEnterpriseReadAccess As Boolean
Private mbMTFReadAccess As Boolean
Private mbClinicReadAccess As Boolean
Private mbPersonalReadAccess As Boolean

Private oEMCategoryDictionary As Scripting.Dictionary
Private oSpecialtyDictionary As Scripting.Dictionary
Private moComplaintList As TemplateData
Private moAppointmentList As TemplateData
Private moProblemList As TemplateData
Private moAutociteList As TemplateData
Private moNotesList As TemplateData
Private moClipartList As TemplateData
Private moDiagnosisList As TemplateData
Private moOrderList As TemplateData
Private moProcedureList As TemplateData
Private moInstructions As TemplateData
Private moForms As TemplateData 'emc Forms support

Public Function AuditTemplateHistory(pTemplateID As String) As Boolean
    'SCR-49441 RBELL 2/11/2004 Added public interface to call global template audit function
    On Error GoTo ErrHandler
    AuditTemplateHistory = gAuditTemplateHistory(pTemplateID)
    Exit Function

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.AuditTemplateHistory", Err.Description

End Function

Private Function BuildNameSearchString(SearchType As TemplateTypeEnum) As String
'SCR-49343 RBELL 2/9/4 Added template type support

    Dim sNameSearch As String
    Dim sNameString As String
    Dim sNamePattern As String
    Dim sNamePrefix As String
    
    On Error GoTo ErrHandler
    Select Case SearchType
        Case EncounterTemplateType
            sNamePrefix = EncTemplatePrefix
        Case ListTemplateType
            sNamePrefix = ListTemplatePrefix
        Case FormTemplateType
            sNamePrefix = FormTemplatePrefix
        Case Else
            sNamePrefix = vbNullString
    End Select
    
    If msSearchName = vbNullString Then
        sNameString = " name like "
        sNamePattern = vbNullString
    Else
        If mbSearchNameCaseInsensitive Then
            If gobjShared.UseSQLServer And Not gobjShared.AppMode = modeCHCSI_GUI Then
                sNameString = " name like "
            Else
                sNameString = " upper(name) like "
            End If
            sNamePattern = UCase(msSearchName)
        Else
            sNameString = " name like "
            sNamePattern = msSearchName
        End If
        If meSearchNameType = ContainsNameSearch Then
            sNamePattern = "%" & sNamePattern
        End If
    End If
    
    Select Case SearchType
        Case EncounterTemplateType
            sNameSearch = " and " & sNameString & SQLQuote(EncTemplatePrefix & sNamePattern & "%") & vbCrLf
        Case ListTemplateType
            sNameSearch = " and " & sNameString & SQLQuote(ListTemplatePrefix & sNamePattern & "%") & vbCrLf
        Case FormTemplateType
            sNameSearch = " and " & sNameString & SQLQuote(FormTemplatePrefix & sNamePattern & "%") & vbCrLf
        Case Both_Encounter_And_List_Type
            sNameSearch = " and (" & sNameString & SQLQuote(EncTemplatePrefix & sNamePattern & "%") & " OR " & sNameString & SQLQuote(ListTemplatePrefix & sNamePattern & "%") & ")" & vbCrLf
        Case Both_Encounter_And_Form_Type
            sNameSearch = " and (" & sNameString & SQLQuote(EncTemplatePrefix & sNamePattern & "%") & " OR " & sNameString & SQLQuote(FormTemplatePrefix & sNamePattern & "%") & ")" & vbCrLf
        Case Both_Form_And_List_Type
            sNameSearch = " and (" & sNameString & SQLQuote(ListTemplatePrefix & sNamePattern & "%") & " OR " & sNameString & SQLQuote(FormTemplatePrefix & sNamePattern & "%") & ")" & vbCrLf
        Case Else
            sNameSearch = " and (" & sNameString & SQLQuote(EncTemplatePrefix & sNamePattern & "%") & " OR " & _
                                    sNameString & SQLQuote(ListTemplatePrefix & sNamePattern & "%") & " OR " & _
                                    sNameString & SQLQuote(FormTemplatePrefix & sNamePattern & "%") & ")" & vbCrLf
    End Select

    BuildNameSearchString = sNameSearch
    Exit Function

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.BuildNameSearchString", Err.Description
    
End Function

Public Function CHCS_Get_Legacy_Templates(pType As TemplateOwnerTypeEnum, pOwner As String) As Collection

    Dim rsTemplates As ADODB.Recordset
    Dim oTemplate As Template
    Dim eOwnerType As TemplateOwnerTypeEnum
    Dim oCHCSInterface_Template As ICHCSTemplates
    Dim pResults As Collection
    
    Dim cTemplates As Collection
    Dim cData As Collection
    
    Dim sData As String
    
    On Error GoTo ErrHandler
    If gobjShared Is Nothing Then Set gobjShared = New CWShared
    Set oCHCSInterface_Template = gobjShared.CHCSConnection
    Set pResults = New Collection
    Set cTemplates = New Collection
    
    Select Case pType
        Case TemplateOwnerTypeEnum.MTFOwnerType
            eOwnerType = eTemplateType_MTF
        Case TemplateOwnerTypeEnum.ClinicOwnerType
            eOwnerType = eTemplateType_Clinic
        Case TemplateOwnerTypeEnum.PersonalOwnerType
            eOwnerType = eTemplateType_userncid
        Case Else
            eOwnerType = eTemplateType_userncid
    End Select

    Set rsTemplates = oCHCSInterface_Template.RetrieveTemplatesCHCS(eOwnerType, pOwner)
    If Not rsTemplates Is Nothing Then
'      For Each oField In rsTemplates.Fields
'        Debug.Print oField.Name
'      Next
      With rsTemplates
             .Sort = "OR_SET_NAME ASC"
             If Not (.BOF And .EOF) Then
                .MoveFirst
                Do Until .EOF
                        On Error Resume Next
                        Set oTemplate = cTemplates(.Fields("OR_SET_NAME"))
                        If Err = 5 Then
                           Err.Clear
                           Set oTemplate = NewTemplate(EncounterTemplateType)
                        ElseIf Err Then
                           GoTo ErrHandler
                        End If
                        
                        On Error GoTo ErrHandler
                        oTemplate.TemplateID = vbNullString
                        oTemplate.TemplateName = .Fields("OR_SET_NAME")
                        If eOwnerType = eTemplateType_Clinic Then
                            oTemplate.OwnerType = ClinicOwnerType
                            oTemplate.OwnerNCID = cwobjLogonEnviron.NursingDivisionNCID
                        Else
                            oTemplate.OwnerType = PersonalOwnerType
                            oTemplate.OwnerNCID = cwobjLogonEnviron.UserNCID
                        End If
                        oTemplate.TemplateFacilityNCID = cwobjLogonEnviron.FacilityNCID
                        
                        With oTemplate
                               Set cData = New Collection
                               cData.Add rsTemplates.Fields("OR_SET_NAME") & vbNullString
                                cData.Add rsTemplates.Fields("OR_SET_FULL_NAME") & vbNullString
                                cData.Add rsTemplates.Fields("CREATED_FOR_IEN") & vbNullString
                                cData.Add rsTemplates.Fields("CREATED_FOR") & vbNullString
                                cData.Add rsTemplates.Fields("CLINIC_IEN") & vbNullString
                                cData.Add rsTemplates.Fields("CLINIC") & vbNullString
                                cData.Add rsTemplates.Fields("ORDER_TYPE_IEN") & vbNullString
                                cData.Add rsTemplates.Fields("ORDER_TYPE") & vbNullString
                                cData.Add rsTemplates.Fields("PROCEDURE_IEN") & vbNullString
                                cData.Add rsTemplates.Fields("ORDER_COMMENT") & vbNullString
                                cData.Add rsTemplates.Fields("COMMENT_DISPLAY") & vbNullString
                                cData.Add rsTemplates.Fields("TIMING_DISPLAY") & vbNullString
                                cData.Add rsTemplates.Fields("ORDER_DISPLAY") & vbNullString
                                cData.Add rsTemplates.Fields("PRIORITY") & vbNullString
                                cData.Add rsTemplates.Fields("FREQUENCY") & vbNullString
                                cData.Add rsTemplates.Fields("SCHEDULE_TYPE") & vbNullString
                                cData.Add rsTemplates.Fields("DURATION") & vbNullString
                                cData.Add rsTemplates.Fields("TIMING_TEXT") & vbNullString
                                cData.Add rsTemplates.Fields("TOTAL_DOSE") & vbNullString
                                cData.Add rsTemplates.Fields("LAB_TEST_IEN") & vbNullString
                                cData.Add rsTemplates.Fields("LAB_TEST") & vbNullString
                                cData.Add rsTemplates.Fields("COLLECTION_SAMPLE") & vbNullString
                                cData.Add rsTemplates.Fields("COLLECTION_METHOD") & vbNullString
                                cData.Add rsTemplates.Fields("SPECIMEN_SITE") & vbNullString
                                cData.Add rsTemplates.Fields("LAB_PROCESSING_PRIORITY") & vbNullString
                                cData.Add rsTemplates.Fields("MEDICATION_IEN") & vbNullString
                                cData.Add rsTemplates.Fields("MEDICATION") & vbNullString
                                cData.Add rsTemplates.Fields("SIG") & vbNullString
                                cData.Add rsTemplates.Fields("QUANTITY") & vbNullString
                                cData.Add rsTemplates.Fields("REFILLS") & vbNullString
                                cData.Add rsTemplates.Fields("CHILD_RESISTANT_CONTAINER") & vbNullString
                                cData.Add rsTemplates.Fields("MED_COMMENT") & vbNullString
                                cData.Add rsTemplates.Fields("RAD_IEN") & vbNullString
                                cData.Add rsTemplates.Fields("RAD") & vbNullString
                                cData.Add rsTemplates.Fields("IMMEDIATE_READING") & vbNullString
                                cData.Add rsTemplates.Fields("CLINICAL_IMPRESSION") & vbNullString
                                cData.Add rsTemplates.Fields("TYPE") & vbNullString
                                cData.Add rsTemplates.Fields("TUBE_TOP_COLOR") & vbNullString
                                cData.Add rsTemplates.Fields("TYPE_IEN") & vbNullString
                                cData.Add rsTemplates.Fields("COLLECTION_SAMPLE_IEN") & vbNullString
                                sData = SerializeCollection(cData, gsListDelim)
                               .ImportCHCSOrders sData
                        End With
                    .MoveNext
                    On Error Resume Next
                    cTemplates.Add oTemplate, oTemplate.TemplateName
                    If Err = 457 Then
                       cTemplates.Remove oTemplate.TemplateName
                       cTemplates.Add oTemplate, oTemplate.TemplateName
                    End If
                    On Error GoTo ErrHandler
                Loop
             End If
      End With
    End If

    GoTo CleanUpAndExit

ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "TemplateOps.CHCS_Get_Legacy_Templates", "TemplateClient", vbCritical)
    'Resume Next
CleanUpAndExit:
    Set CHCS_Get_Legacy_Templates = cTemplates
    Set oTemplate = Nothing
    Set rsTemplates = Nothing
    Set oCHCSInterface_Template = Nothing


End Function

Public Function CheckForTemplateReplacement(oTemplate As Object) As Integer
    'SCR-44247 RBELL Check whether the template exists and return an indicator to calling routine
    
    Dim oConn As CHCSII_CONN.Conn
    Dim oSQL As ISqlOps
    Dim sPrefix As String
    
    On Error GoTo ErrHandler
    Set oConn = New CHCSII_CONN.Conn
    Set oSQL = oConn.CHCSII_SQLOPS(Auto)
    CheckForTemplateReplacement = -1
    
    If oTemplate Is Nothing Then Exit Function
    Select Case oTemplate.TemplateType
        Case EncounterTemplateType
            sPrefix = EncTemplatePrefix
        Case ListTemplateType
            sPrefix = ListTemplatePrefix
        Case FormTemplateType
            sPrefix = FormTemplatePrefix
        Case Else
            Exit Function
    End Select
    
    If oTemplate.TemplateID = 0 Then
        oSQL.Execute "Select id, null as FAVORITES_USER, Creator_NCID from template_info where name = " & SQLQuote(sPrefix & oTemplate.TemplateName) _
            & " and owner = " & oTemplate.OwnerNCID
    Else
        oSQL.Execute "Select id, null as FAVORITES_USER, Creator_NCID  from template_info where name = " & SQLQuote(sPrefix & oTemplate.TemplateName) _
            & " and id <> " & oTemplate.TemplateID & " and owner = " & oTemplate.OwnerNCID
    End If
    
    If oSQL.RecordCount > 0 Then
        CheckForTemplateReplacement = frmReplace.QueryReplace(oTemplate.TemplateName, Me)
        If CheckForTemplateReplacement = ReplaceTemplate Then
           Set oTemplate = GetTemplate(oSQL("ID"))
        End If
    End If
    
    GoTo CleanUpAndExit
    
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.CheckForTemplateReplacement", Err.Description

CleanUpAndExit:
    Set oConn = Nothing
    Set oSQL = Nothing

End Function

Private Function DoSearchMTFs() As ADODB.Recordset
           
    Dim sQuery As String
    Dim oConn As CHCSII_CONN.Conn
    Dim oDas As ICHCSII_DAS
    
    On Error GoTo ErrHandler
    'SCR-52426 RBELL 4/15/2004 Only execute query in CDR mode
    
    If gobjShared.IsAppMode(modeCDR) And Not gobjShared.IsAppMode(modeITT) Then
        Set oConn = New CHCSII_CONN.Conn
        Set oDas = oConn.CHCSII_DAS(Auto)
        If mbSearchAdvanced = True Or gcwobjUser.HasPrivilegeEx(gsMTFAccess, Priv_Read) Then
           sQuery = "select ncid, UPPER(representation) as name" & vbCrLf & _
                "From concept_relation, rsform, rsform_context" & vbCrLf & _
                "Where rsform.rsform_id = rsform_context.rsform_id" & vbCrLf & _
                "and context_ncid = 2000" & vbCrLf & _
                "and concept_relation_ncid = 33962" & vbCrLf & _
                "and relationship_ncid = 363" & vbCrLf & _
                "and ncid = concept_ncid" & vbCrLf & _
                "order by name"
        Else
           sQuery = "select ncid, UPPER(representation) as name" & vbCrLf & _
                "From concept_relation, rsform, rsform_context" & vbCrLf & _
                "Where ncid in (SELECT facility_ncid from user_clinic u where user_ncid = " & Logon.UserNCID & ") and " & vbCrLf & _
                "rsform.rsform_id = rsform_context.rsform_id" & vbCrLf & _
                "and context_ncid = 2000" & vbCrLf & _
                "and concept_relation_ncid = 33962" & vbCrLf & _
                "and relationship_ncid = 363" & vbCrLf & _
                "and ncid = concept_ncid" & vbCrLf & _
                "order by name"
        End If
    Else
        Set DoSearchMTFs = Nothing
        Exit Function
    End If
    
    Set DoSearchMTFs = oDas.OpenRecordset(sQuery)
    GoTo CleanUpAndExit

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.DoSearchMTFs", Err.Description
CleanUpAndExit:
    Set oConn = Nothing
    Set oDas = Nothing

End Function

Public Function GetMTFs() As ADODB.Recordset
    
    On Error GoTo ErrHandler
    If moMTFs Is Nothing Then Set moMTFs = DoSearchMTFs
    Set GetMTFs = moMTFs
    
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.GetMTFs", Err.Description
    
End Function

Public Sub UpdateFavorite(rsTemplateID As String, bfavorite As Boolean)
    
    Dim objSQL    As ICHCSII_SQL.ISqlOps
    Dim moConn As CHCSII_CONN.Conn
    Dim sSQL As String
                              
    On Error GoTo ErrHandler
    Set moConn = New CHCSII_CONN.Conn
    Set objSQL = moConn.CHCSII_SQLOPS(Auto)
          
     If bfavorite Then
       sSQL = "update template_info set ISActive='Y' where id = " & rsTemplateID & vbNullString
    Else
       sSQL = "update template_info set ISActive='N' where id = " & rsTemplateID & vbNullString
    End If
    
    Call objSQL.Execute(sSQL)
        
    GoTo CleanUpAndExit

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.UpdateFavorite", Err.Description

CleanUpAndExit:
    Set moConn = Nothing
    Set objSQL = Nothing

End Sub

Public Sub UpdateShared(rsTemplateID As String, bShared As Boolean)
    'SCR-44203 10/27/03 RBELL New routine added to support updating the shared field for the template.
    
    Dim objSQL    As ICHCSII_SQL.ISqlOps
    Dim moConn As CHCSII_CONN.Conn
    Dim sSQL As String
                              
    On Error GoTo ErrHandler
    Set moConn = New CHCSII_CONN.Conn
    Set objSQL = moConn.CHCSII_SQLOPS(Auto)
          
     If bShared Then
       sSQL = "update template_info set IsShared= 'Y' where id = " & rsTemplateID & vbNullString
    Else
       sSQL = "update template_info set IsShared= 'N' where id = " & rsTemplateID & vbNullString
    End If
    
    Call objSQL.Execute(sSQL)
    GoTo CleanUpAndExit

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.UpdateShared", Err.Description

CleanUpAndExit:
        
    Set moConn = Nothing
    Set objSQL = Nothing

End Sub


Public Property Get ComplaintList() As TemplateData
    
    On Error GoTo ErrHandler
    Set ComplaintList = moComplaintList
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.ComplaintList", Err.Description
    
End Property

Public Property Set ComplaintList(ByVal RHV As TemplateData)
    
    On Error GoTo ErrHandler
    Set moComplaintList = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.ComplaintList", Err.Description
    
End Property

Public Property Get AppointmentList() As TemplateData
    
    On Error GoTo ErrHandler
    Set AppointmentList = moAppointmentList
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.AppointmentList", Err.Description
    
End Property

Public Property Set AppointmentList(ByVal RHV As TemplateData)
    
    On Error GoTo ErrHandler
    Set moAppointmentList = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.AppointmentList", Err.Description
    
End Property

Friend Property Get Concept() As Object
    
    On Error GoTo ErrHandler
    If Not gobjShared.IsAppMode(modeCHCSI_GUI) Then   '<SCR 36493
        If mobjConcept Is Nothing Then
            'SCR-53524 5/10/2004 Set object based on appmode
            If gobjShared.IsAppMode(modeCDR) And Not gobjShared.IsAppMode(modeITT) Then
                Set mobjConcept = New ConceptCtrl
            Else
                Set mobjConcept = New GEMS_ConceptCtrl
            End If
        End If
        Set Concept = mobjConcept
    End If
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.Concept", Err.Description
    
End Property
Public Property Get ProblemList() As TemplateData
    
    On Error GoTo ErrHandler
    Set ProblemList = moProblemList
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.ProblemList", Err.Description
    
End Property

Public Property Set ProblemList(ByVal RHV As TemplateData)
    
    On Error GoTo ErrHandler
    Set moProblemList = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.ProblemList", Err.Description
    
End Property

Public Property Get AutociteList() As TemplateData
    
    On Error GoTo ErrHandler
    Set AutociteList = moAutociteList
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.AutociteList", Err.Description
    
End Property

Public Property Set AutociteList(ByVal RHV As TemplateData)
    
    On Error GoTo ErrHandler
    Set moAutociteList = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.AutociteList", Err.Description
    
End Property

Public Property Get NotesList() As TemplateData
    
    On Error GoTo ErrHandler
    Set NotesList = moNotesList
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.NotesList", Err.Description
    
End Property

Public Property Set NotesList(ByVal RHV As TemplateData)
    
    On Error GoTo ErrHandler
    Set moNotesList = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.NotesList", Err.Description
    
End Property

Public Property Get ClipartList() As TemplateData
    
    On Error GoTo ErrHandler
    Set ClipartList = moClipartList
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.ClipartList", Err.Description
    
End Property

Public Property Set ClipartList(ByVal RHV As TemplateData)
    
    On Error GoTo ErrHandler
    Set moClipartList = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.ClipartList", Err.Description
    
End Property

Public Property Get DiagnosisList() As TemplateData
    
    On Error GoTo ErrHandler
    Set DiagnosisList = moDiagnosisList
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.DiagnosisList", Err.Description
    
End Property

Public Property Set DiagnosisList(ByVal RHV As TemplateData)
    
    On Error GoTo ErrHandler
    Set moDiagnosisList = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.DiagnosisList", Err.Description
    
End Property

Public Property Get OrderList() As TemplateData
    
    On Error GoTo ErrHandler
    Set OrderList = moOrderList
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.OrderList", Err.Description
    
End Property

Public Property Set OrderList(ByVal RHV As TemplateData)
    
    On Error GoTo ErrHandler
    Set moOrderList = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.OrderList", Err.Description
    
End Property

Public Property Get ProcedureList() As TemplateData
    
    On Error GoTo ErrHandler
    Set ProcedureList = moProcedureList
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.ProcedureList", Err.Description
    
End Property

Public Property Set ProcedureList(ByVal RHV As TemplateData)
    
    On Error GoTo ErrHandler
    Set moProcedureList = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.ProcedureList", Err.Description
    
End Property

Public Property Get Instructions() As TemplateData
    
    On Error GoTo ErrHandler
    Set Instructions = moInstructions
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.Instructions", Err.Description
    
End Property

Public Property Set Instructions(ByVal RHV As TemplateData)
    
    On Error GoTo ErrHandler
    Set moInstructions = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.Instructions", Err.Description
    
End Property

'emc Forms Support
Public Property Get Forms() As TemplateData
   
    On Error GoTo ErrHandler
    Set Forms = moForms
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.Forms", Err.Description
    
End Property
Public Property Set Forms(ByVal RHV As TemplateData)
   
    On Error GoTo ErrHandler
    Set moForms = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.Forms", Err.Description
    
End Property

Public Property Get EMCategoryDictionary() As Scripting.Dictionary
    
    On Error GoTo ErrHandler
    Set EMCategoryDictionary = oEMCategoryDictionary
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.EMCategoryDictionary", Err.Description
    
End Property

Public Property Set EMCategoryDictionary(ByVal RHV As Scripting.Dictionary)
    
    On Error GoTo ErrHandler
    Set oEMCategoryDictionary = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.EMCategoryDictionary", Err.Description
    
End Property

Public Property Get SpecialtyDictionary() As Scripting.Dictionary
    
    On Error GoTo ErrHandler
    Set SpecialtyDictionary = oSpecialtyDictionary
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.SpecialtyDictionary", Err.Description
    
End Property

Public Property Set SpecialtyDictionary(ByVal RHV As Scripting.Dictionary)
    
    On Error GoTo ErrHandler
    Set oSpecialtyDictionary = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.SpecialtyDictionary", Err.Description
    
End Property

Public Function GetCPGItems() As Collection
    
    Dim sSQL As String
    Dim moConn As CHCSII_CONN.Conn
    Dim oSQL As ICHCSII_SQL.ISqlOps
    
    Set moConn = New CHCSII_CONN.Conn
    Set oSQL = moConn.CHCSII_SQLOPS(Auto)
    'GEMS >>
    
    On Error GoTo ErrHandler
    
    Set GetCPGItems = New Collection
    If gobjShared.IsAppMode(modeCHCSI_GUI) Or gobjShared.IsAppMode(modeITT) Or gobjShared.IsAppMode(modeTheater) Then Exit Function
    sSQL = "SELECT * FROM CPGDEF"
    oSQL.Execute sSQL
    
    If Not (oSQL.BOF And oSQL.EOF) Then
       oSQL.MoveFirst
       Do Until oSQL.EOF
           GetCPGItems.Add oSQL("CPG_ID") & "|" & oSQL("NAME")
           oSQL.MoveNext
       Loop
    End If
    
    GoTo CleanUpAndExit
    
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "TemplateOps.GetCPGItems", "TemplateClient", vbCritical)

CleanUpAndExit:
    Set oSQL = Nothing
    Set moConn = Nothing

End Function
Friend Property Get Logon() As MMMHISLogon.Logon
    
    On Error GoTo ErrHandler
    Set Logon = cwobjLogonEnviron
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.Logon", Err.Description
    
End Property

Public Sub Initialize(ByVal vlProgId As Long, ByRef robjComm As Object, ByRef robjLogon As MMMHISLogon.Logon, ByRef robjConfig As Object, ByRef robjPatient As Object, ByRef robjUser As Object)
    
    On Error GoTo ErrHandler
    Set cwobjCWComm = robjComm
    Set cwobjPatient = robjPatient
    Set gcwobjUser = robjUser
    Set cwobjLogonEnviron = robjLogon
    Set gobjComm = robjComm
    Set gobjPatient = robjPatient
    Set cwobjConfig = robjConfig
    
    mbEnterpriseReadAccess = gcwobjUser.HasPrivilegeEx(Priv_Templates_Enterprise_Template_Guidelines, Priv_Read) Or gcwobjUser.HasPrivilegeEx(Priv_Order_Sets_Enterprise_Order_Set, Priv_Read)
    mbMTFReadAccess = gcwobjUser.HasPrivilegeEx(Priv_Templates_MTF_Template_Guidelines, Priv_Read) Or gcwobjUser.HasPrivilegeEx(Priv_Order_Sets_MTF_Order_Set, Priv_Read)
    mbClinicReadAccess = gcwobjUser.HasPrivilegeEx(Priv_Templates_Clinic_Templates, Priv_Read) Or gcwobjUser.HasPrivilegeEx(Priv_Order_Sets_Clinic_Order_Set, Priv_Read)
    mbPersonalReadAccess = gcwobjUser.HasPrivilegeEx(Priv_Templates_Personal_Template, Priv_Read) Or gcwobjUser.HasPrivilegeEx(Priv_Order_Sets_Personal_Order_Sets, Priv_Read)
    
    If goOEGetSessionMgr Is Nothing Then
       Set goOEGetSessionMgr = New GEMS_OE.OE_GetSessionManager
    End If
    
    If Not goOEGetSessionMgr Is Nothing Then
       Set goOESession = goOEGetSessionMgr.GetSessionManager
    End If
    
    Exit Sub

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.Initialize", Err.Description
       
End Sub

Public Sub ClearSearchCriteria()
    
    On Error GoTo ErrHandler
    If gcwobjUser Is Nothing Then
        Err.Raise vbObjectError + 1, "CHCSII_TemplateClient.TemplateOps", "Initialize(...) must be called before using TemplateOps."
        Exit Sub
    End If
    
    msSearchID = vbNullString
    msSearchName = vbNullString
    meSearchNameType = BeginsWithNameSearch
    mbSearchNameCaseInsensitive = False
    SearchShared = False
    meSearchOwnerType = 0
    Set mcSearchOwners = New Collection
    Set mcSearchTemplateType = New Collection
    mbSearchFavorites = False
    msSearchSpecialtyID = vbNullString
    Set moSearchComplaints = New Complaints
    Set moSearchAppointmentTypes = New AppointmentTypes
    Set moSearchProblems = New Problems
    Exit Sub

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.ClearSearchCriteria", Err.Description
    
End Sub

Public Function CheckForDuplicate(ByVal ID As String, ByVal Name As String, ByVal OwnerNCID As Long, ByVal TemplateType As TemplateTypeEnum) As Boolean
    
    On Error GoTo ErrHandler
    If TemplateType = EncounterTemplateType Then
        Name = EncTemplatePrefix & Name
    ElseIf TemplateType = ListTemplateType Then 'emc added for forms support
        Name = ListTemplatePrefix & Name
    ElseIf TemplateType = FormTemplateType Then 'emc Forms support
        Name = FormTemplatePrefix & Name
    End If
    
    'SCR-50858 RBELL 3/8/2004 using 'x' returned a record even if none existed changed to an actual field name
    If ID = vbNullString Then
        If gobjShared.UseSQLServer And Not gobjShared.AppMode = modeCHCSI_GUI Then
            mobjSQL.Execute "Select id from template_info where name = " & SQLQuote(Name) _
                & " and owner = " & OwnerNCID
        Else
            mobjSQL.Execute "Select id from template_info where UPPER(name) = " & UCase(SQLQuote(Name)) _
                & " and owner = " & OwnerNCID
        End If
    Else
        If gobjShared.UseSQLServer And Not gobjShared.AppMode = modeCHCSI_GUI Then
            mobjSQL.Execute "Select id from template_info where name = " & SQLQuote(Name) _
                & " and id <> " & ID & " and owner = " & OwnerNCID
        Else
            mobjSQL.Execute "Select id from template_info where UPPER(name) = " & UCase(SQLQuote(Name)) _
                & " and id <> " & ID & " and owner = " & OwnerNCID
        End If
    End If
    If mobjSQL.RecordCount > 0 Then
        CheckForDuplicate = True
    End If
    Exit Function

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.CheckForDuplicate", Err.Description
    
End Function

Public Function CheckPrivilege(ByVal vOwnerNCID As String, _
                               ByVal vOwnerType As CHCSII_TemplateClient.TemplateOwnerTypeEnum, _
                               ByVal vPrivilege As CHCSII_ClientSecurity.CHCSII_PrivilegeType, _
                               Optional ByVal vSilent As Boolean = False) As Boolean
    
    On Error GoTo ErrHandler
    CheckPrivilege = CheckUserPrivilege(gcwobjUser, vOwnerNCID, cwobjLogonEnviron.UserNCID, vOwnerType, vPrivilege, vSilent)
    Exit Function

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.CheckPrivilege", Err.Description
    
End Function

Public Function DoInteractiveSearch(ByVal SearchType As TemplateTypeEnum, _
                        Optional ByRef ExistingCollection As Collection = Nothing) As Collection
'SCR-49343 RBELL 2/9/4 Added template type support 2 params

    Dim objTemplate As Object
    Dim lPointer As Long
    
    On Error GoTo ErrHandler
    frmSearch.TemplateOps = Me
    If cwobjCWComm.CmdLineSwitch("YELLOW") = vbNullString Then
       frmSearch.TemplateOps.SearchAdvanced = False
    Else
       frmSearch.TemplateOps.SearchAdvanced = True
    End If
    
    frmSearch.SearchType = SearchType
    
    If Not ExistingCollection Is Nothing Then
        frmSearch.MergeResults = True
    Else
        frmSearch.MergeResults = False
    End If
        
    cwobjCWComm.Message cwiSHOWING_MODAL_FORM, vbNullString, NCID_CW_APPLICATION, NCID_TEMPLATE_MANAGEMENT
    frmSearch.Show vbModal
    cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, vbNullString, NCID_CW_APPLICATION, NCID_TEMPLATE_MANAGEMENT
    If Not frmSearch.Cancelled Then
        lPointer = Screen.MousePointer
        Screen.MousePointer = vbHourglass
        DoEvents
        Set DoInteractiveSearch = DoSearchLite(SearchType)
        If frmSearch.MergeResults And Not ExistingCollection Is Nothing Then
            For Each objTemplate In ExistingCollection
                DoInteractiveSearch.Add objTemplate, objTemplate.TemplateID
            Next objTemplate
        End If
        Screen.MousePointer = lPointer
    Else
        'Set to nothing is user cancels SCR-11654
        Set DoInteractiveSearch = Nothing
    End If
    
    Exit Function

ErrHandler:
    If Err = 457 Then
       If Not objTemplate Is Nothing Then
          DoInteractiveSearch.Remove objTemplate.TemplateID
          Resume
       End If
    End If
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.DoInteractiveSearch", Err.Description
    'Resume Next
    
End Function

Private Function DoQuery(ByVal SearchType As TemplateTypeEnum, ByVal WithXRefType As Boolean, Optional sFolderID As String = vbNullString) As Boolean
'SCR-48609 RBELL 1/23/04 Add optional parameter for processing
'SCR-44219, scr-44220, scr-48192 RBELL 1/23/04 Modified search options based on advanced search options
    Dim i As Integer
    Dim sQuery As String
    Dim sStr As String
    Dim sInfoWhere As String
    Dim sXRefWhere As String
    Dim bFavorites As Boolean
    Dim bSearchOwn As Boolean
    Dim bVirtTable As Boolean 'SCR-11652
    
    If gcwobjUser Is Nothing Then
        Err.Raise vbObjectError + 1, "CHCSII_TemplateClient.TemplateOps", "Initialize(...) must be called before using TemplateOps."
        Exit Function
    End If
    
    If gobjShared Is Nothing Then Set gobjShared = New CWShared
    
    On Error GoTo ErrHandler
    
    If Not msSearchID = vbNullString Then
        sInfoWhere = sInfoWhere & " and id = " & msSearchID & vbCrLf
    End If
        
    'SCR-49343 RBELL 2/9/4 Added template type support
    sInfoWhere = sInfoWhere & BuildNameSearchString(SearchType)

    Select Case meSearchOwnerType
    
        Case PersonalOwnerType
            sInfoWhere = sInfoWhere & " and type = " & tt_user & vbCrLf
            'SCR-44220 RBELL 1/15/04 Only add owner criteria if not advanced type
            If mcSearchOwners.Count = 0 And Not mbSearchAdvanced Then
                sInfoWhere = sInfoWhere & " and owner = " & cwobjLogonEnviron.UserNCID & vbCrLf
                If SearchShared Then
                    sInfoWhere = sInfoWhere & " and isshared = 'Y'" & vbCrLf
                End If
            ElseIf mcSearchOwners.Count = 0 And mbSearchAdvanced Then
                'SCR-48736 RBELL 1/27/2004 Need to modify SQL for shared templates
                sInfoWhere = sInfoWhere & " and ((owner = " & cwobjLogonEnviron.UserNCID
                If SearchShared Then sInfoWhere = sInfoWhere & " and isshared = 'Y')" & vbCrLf
                sInfoWhere = sInfoWhere & ")" & vbCrLf
                sInfoWhere = sInfoWhere & " OR (owner <> " & cwobjLogonEnviron.UserNCID & " and isshared = 'Y'))" & vbCrLf
            ElseIf mcSearchOwners.Count > 0 Then
                For i = 1 To mcSearchOwners.Count
                    If CStr(mcSearchOwners(i)) = cwobjLogonEnviron.UserNCID Then
                        bSearchOwn = True
                        mcSearchOwners.Remove i
                        sInfoWhere = sInfoWhere & "and (owner = " & cwobjLogonEnviron.UserNCID
                        Exit For
                    End If
                Next i
                If mcSearchOwners.Count > 0 Then
                    If bSearchOwn Then
                        sInfoWhere = sInfoWhere & " or ("
                    Else
                        sInfoWhere = sInfoWhere & " and "
                    End If
                    sInfoWhere = sInfoWhere & "isshared = 'Y'" & " and owner in(" & mcSearchOwners(1)
                    For i = 2 To mcSearchOwners.Count
                        sInfoWhere = sInfoWhere & ", " & mcSearchOwners(i)
                    Next i
                    sInfoWhere = sInfoWhere & ")" & vbCrLf
                End If
                If bSearchOwn Then
                    If mcSearchOwners.Count > 0 Then
                        mcSearchOwners.Add cwobjLogonEnviron.UserNCID, , 1
                        sInfoWhere = sInfoWhere & "))"
                    Else
                        mcSearchOwners.Add cwobjLogonEnviron.UserNCID
                        sInfoWhere = sInfoWhere & ")"
                    End If
                End If
            End If
        Case ClinicOwnerType
            sInfoWhere = sInfoWhere & " and type = " & tt_facility & " and isshared = 'Y'" & vbCrLf
            'SCR-44220 RBELL 1/15/04 Only add owner criteria if not advanced type
            If mcSearchOwners.Count = 0 And Not mbSearchAdvanced Then
                sInfoWhere = sInfoWhere & " and owner in (select clinic_ncid from user_clinic" _
                        & " where user_ncid = " & cwobjLogonEnviron.UserNCID & ")" & vbCrLf
            ElseIf mcSearchOwners.Count = 0 And mbSearchAdvanced Then
                'SCR-44220 RBELL 1/15/04 If no owner provided limit templates to entries from the clinic table
                sInfoWhere = sInfoWhere & " and owner in (select distinct ncid from clinic)" & vbCrLf
            ElseIf mcSearchOwners.Count > 0 Then
                sInfoWhere = sInfoWhere & " and owner in(" & mcSearchOwners(1)
                For i = 2 To mcSearchOwners.Count
                    sInfoWhere = sInfoWhere & ", " & mcSearchOwners(i)
                Next i
                sInfoWhere = sInfoWhere & ")" & vbCrLf
            End If
        Case MTFOwnerType
            sInfoWhere = sInfoWhere & " and type = " & tt_facility & vbCrLf
            'SCR-44220 RBELL 1/15/04 Only add owner criteria if not advanced type
            If mcSearchOwners.Count = 0 And Not mbSearchAdvanced Then
               sInfoWhere = sInfoWhere & " and owner = " & cwobjLogonEnviron.FacilityNCID & vbCrLf
            ElseIf mcSearchOwners.Count = 0 And mbSearchAdvanced Then
               sInfoWhere = sInfoWhere & " and owner in (SELECT distinct p.NCID FROM rsform p, concept_relation cr WHERE  p.ncid = cr.concept_ncid and cr.CONCEPT_RELATION_NCID = 33962)"
            ElseIf mcSearchOwners.Count > 0 Then
                sInfoWhere = sInfoWhere & " and owner in(" & mcSearchOwners(1)
                For i = 2 To mcSearchOwners.Count
                    sInfoWhere = sInfoWhere & ", " & mcSearchOwners(i)
                Next i
                sInfoWhere = sInfoWhere & ")" & vbCrLf
            End If
        Case EnterpriseOwnerType
            sInfoWhere = sInfoWhere & " and type = " & tt_enterprise & vbCrLf
        Case Else
            ' By default, a user can see anything except other users' non-shared templates
            ' and clinic templates for clinics he is not assigned to
            'SCR-44220 RBELL 1/15/04 Only add owner criteria if not advanced type
            If Not mbSearchAdvanced Then
                If gobjShared.AppMode = modeCHCSI_GUI Then
                    'SCR-49546 RBELL 2/13/2004 Subquery in statement in PGUI mode caused lockup. Replaced with a call to sub that built clinic list and removed MTF option from query
                    'SCR-51367 RBELL 3/25/2004 Backed 49546 out to see if speed improvements have been made on backend
                    sInfoWhere = sInfoWhere & " and (" & vbCrLf _
                        & "(type = " & tt_user & " and (isshared = 'Y' or owner = " & cwobjLogonEnviron.UserNCID & "))" & vbCrLf _
                        & " or (type = " & tt_facility & " and owner in (select clinic_ncid from user_clinic where user_ncid = " & cwobjLogonEnviron.UserNCID & ")))" & vbCrLf
                Else
                    sInfoWhere = sInfoWhere & " and (" & vbCrLf _
                        & "(type = " & tt_user & " and (isshared = 'Y' or owner = " & cwobjLogonEnviron.UserNCID & "))" & vbCrLf _
                        & " or (type = " & tt_facility & " and owner in (select clinic_ncid from user_clinic where user_ncid = " & cwobjLogonEnviron.UserNCID & "))" & vbCrLf _
                        & " or (type = " & tt_facility & " and owner = " & cwobjLogonEnviron.FacilityNCID & ")" & vbCrLf _
                        & " or (type = " & tt_enterprise & " and owner = " & cwobjLogonEnviron.EnterpriseNCID & "))" & vbCrLf
                End If
            End If
    End Select
    sXRefWhere = vbNullString
    sStr = moSearchComplaints.GetQueryCriteria
    If sStr <> vbNullString Then
        If sXRefWhere = vbNullString Then
            sXRefWhere = sStr & vbCrLf
        Else
            sXRefWhere = sXRefWhere & " or " & sStr & vbCrLf
        End If
    End If
    sStr = moSearchAppointmentTypes.GetQueryCriteria
    If sStr <> vbNullString Then
        If sXRefWhere = vbNullString Then
            sXRefWhere = sStr & vbCrLf
        Else
            sXRefWhere = sXRefWhere & " or " & sStr & vbCrLf
        End If
    End If
    sStr = moSearchProblems.GetQueryCriteria
    If sStr <> vbNullString Then
        If sXRefWhere = vbNullString Then
            sXRefWhere = sStr & vbCrLf
        Else
            sXRefWhere = sXRefWhere & " or " & sStr & vbCrLf
        End If
    End If
    bVirtTable = False
    If msSearchSpecialtyID <> vbNullString Then
        sStr = "(type_ncid = " & SpecialtyNCID & " and value_ncid = " & msSearchSpecialtyID & ")"
        If sXRefWhere = vbNullString Then
            sXRefWhere = sStr & vbCrLf
            
        Else
            bVirtTable = True
            'sXRefWhere = sXRefWhere & " or " & sStr & vbCrLf
        End If
    End If
    If mbSearchFavorites Then
        bFavorites = False
    Else
        bFavorites = True
    End If
    
    sQuery = "select distinct info.id, info.owner, info.isactive, info.type, info.isshared, info.name, " & _
                    "favorites.user_ncid as favorites_user, favorites.folder_id, facility.template_facility" & vbCrLf
    
    If sXRefWhere = vbNullString And Not bVirtTable Then
        If bFavorites = False Then
            sQuery = sQuery & " from template_info info JOIN template_favorites favorites ON info.id = favorites.template_id"
            If Me.SearchFavorites Then
                sQuery = sQuery & " and favorites.user_ncid = " & cwobjLogonEnviron.UserNCID & vbCrLf
            ElseIf meSearchOwnerType = EnterpriseOwnerType And sFolderID = "0" Then
                  sQuery = sQuery & " and favorites.folder_id = 0 and favorites.template_id not in (SELECT template_id from template_favorites where folder_id > 0 )"
            End If
            sQuery = sQuery & " LEFT JOIN template_facility facility ON info.id = facility.id " & vbCrLf
        Else
            sQuery = sQuery & " from template_info info LEFT JOIN template_favorites favorites ON info.id = favorites.template_id"
            If Me.SearchFavorites Then
                sQuery = sQuery & " and favorites.user_ncid = " & cwobjLogonEnviron.UserNCID & vbCrLf
            ElseIf meSearchOwnerType = EnterpriseOwnerType And sFolderID = "0" Then
                  sQuery = sQuery & " and favorites.folder_id = 0 and favorites.template_id not in (SELECT template_id from template_favorites where folder_id > 0 )"
            End If
            sQuery = sQuery & " LEFT JOIN template_facility facility ON info.id = facility.id " & vbCrLf
        End If
    Else
        If WithXRefType Then
            sQuery = sQuery & ", type_ncid"
        End If
        
        'If we want to look at a subset of a table use virtual table as subset
        'Ex. look for a Dx in a Speciality SCR-11652
        If bVirtTable Then
            If bFavorites = False Then
                sQuery = sQuery & " from template_info info JOIN template_favorites favorites ON info.id = favorites.template_id "
                If Me.SearchFavorites Then
                    sQuery = sQuery & " and favorites.user_ncid = " & cwobjLogonEnviron.UserNCID & vbCrLf
                ElseIf meSearchOwnerType = EnterpriseOwnerType And sFolderID = "0" Then
                      sQuery = sQuery & " and favorites.folder_id = 0 and favorites.template_id not in (SELECT template_id from template_favorites where folder_id > 0 )"
                End If
                sQuery = sQuery & " LEFT JOIN template_facility facility ON info.id = facility.id " & vbCrLf & _
                    "JOIN (select id, type_ncid, value_ncid from template_xref where id in " & vbCrLf & _
                    "(select id from template_xref where type_ncid = " & SpecialtyNCID & " and value_ncid = " & msSearchSpecialtyID & "))" & " xref ON xref.id = info.id " & vbCrLf
            Else
                sQuery = sQuery & " from template_info info LEFT JOIN template_favorites favorites ON info.id = favorites.template_id "
                If Me.SearchFavorites Then
                    sQuery = sQuery & " and favorites.user_ncid = " & cwobjLogonEnviron.UserNCID & vbCrLf
                ElseIf meSearchOwnerType = EnterpriseOwnerType And sFolderID = "0" Then
                      sQuery = sQuery & " and favorites.folder_id = 0 and favorites.template_id not in (SELECT template_id from template_favorites where folder_id > 0 )"
                End If
                sQuery = sQuery & " LEFT JOIN template_facility facility ON info.id = facility.id " & vbCrLf & _
                    "JOIN (select id, type_ncid, value_ncid from template_xref where id in " & vbCrLf & _
                    "(select id from template_xref where type_ncid = " & SpecialtyNCID & " and value_ncid = " & msSearchSpecialtyID & "))" & " xref ON xref.id = info.id " & vbCrLf
            End If
        Else
            If bFavorites = False Then
                sQuery = sQuery & " from template_info info JOIN template_favorites favorites ON info.id = favorites.template_id "
                If Me.SearchFavorites Then
                    sQuery = sQuery & " and favorites.user_ncid = " & cwobjLogonEnviron.UserNCID & vbCrLf
                ElseIf meSearchOwnerType = EnterpriseOwnerType And sFolderID = "0" Then
                      sQuery = sQuery & " and favorites.folder_id = 0 and favorites.template_id not in (SELECT template_id from template_favorites where folder_id > 0 )"
                End If
                sQuery = sQuery & " LEFT JOIN template_facility facility ON info.id = facility.id JOIN template_xref xref ON xref.id = info.id " & vbCrLf
            Else
                sQuery = sQuery & " from template_info info LEFT JOIN template_favorites favorites ON info.id = favorites.template_id "
                If Me.SearchFavorites Then
                    sQuery = sQuery & " and favorites.user_ncid = " & cwobjLogonEnviron.UserNCID & vbCrLf
                ElseIf meSearchOwnerType = EnterpriseOwnerType And sFolderID = "0" Then
                      sQuery = sQuery & " and favorites.folder_id = 0 and favorites.template_id not in (SELECT template_id from template_favorites where folder_id > 0 )"
                End If
                sQuery = sQuery & " LEFT JOIN template_facility facility ON info.id = facility.id JOIN template_xref xref ON xref.id = info.id " & vbCrLf
            End If
        End If
        
    End If
    
    If sXRefWhere <> vbNullString Then
        sQuery = sQuery & " and (" & sXRefWhere & ")"
    End If
    
    sQuery = sQuery & sInfoWhere
        
    sQuery = sQuery & " order by lower(info.name), info.id"

    Debug.Print sQuery
    mobjSQL.Execute sQuery
    DoQuery = True
    Exit Function

ErrHandler:
    'Resume Next
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "TemplateOps.DoQuery", "TemplateClient", vbCritical)
    Exit Function
    Resume
End Function
    
Private Function DoQueryWithoutFavorites(ByVal SearchType As TemplateTypeEnum, ByVal WithXRefType As Boolean) As Boolean
    
    Dim i As Integer
    Dim sQuery As String
    Dim sStr As String
    Dim sInfoWhere As String
    Dim sXRefWhere As String
    Dim sFavorites As String
    Dim sNamePrefix As String
    Dim sNamePattern As String
    Dim sClinicOwner As String
    Dim bSearchOwn As Boolean
    Dim bVirtTable As Boolean 'SCR-11652
    
    If gcwobjUser Is Nothing Then
        Err.Raise vbObjectError + 1, "CHCSII_TemplateClient.TemplateOps", "Initialize(...) must be called before using TemplateOps."
        Exit Function
    End If
    
    If gobjShared Is Nothing Then Set gobjShared = New CWShared
    
    On Error GoTo ErrHandler
    
    If msSearchID <> vbNullString Then sInfoWhere = sInfoWhere & " and id = " & msSearchID & vbCrLf
    sInfoWhere = sInfoWhere & BuildNameSearchString(SearchType)
    
    Select Case meSearchOwnerType
        Case PersonalOwnerType
            sInfoWhere = sInfoWhere & " and type = " & tt_user & vbCrLf
            If mcSearchOwners.Count = 0 Then
                sInfoWhere = sInfoWhere & " and owner = " & cwobjLogonEnviron.UserNCID & vbCrLf
                If SearchShared Then
                    sInfoWhere = sInfoWhere & " and isshared = 'Y'" & vbCrLf
                End If
            Else
                For i = 1 To mcSearchOwners.Count
                    If CStr(mcSearchOwners(i)) = cwobjLogonEnviron.UserNCID Then
                        bSearchOwn = True
                        mcSearchOwners.Remove i
                        sInfoWhere = sInfoWhere & "and (owner = " & cwobjLogonEnviron.UserNCID
                        Exit For
                    End If
                Next i
                If mcSearchOwners.Count > 0 Then
                    If bSearchOwn Then
                        sInfoWhere = sInfoWhere & " or ("
                    Else
                        sInfoWhere = sInfoWhere & " and "
                    End If
                    sInfoWhere = sInfoWhere & "isshared = 'Y'" & " and owner in(" & mcSearchOwners(1)
                    For i = 2 To mcSearchOwners.Count
                        sInfoWhere = sInfoWhere & ", " & mcSearchOwners(i)
                    Next i
                    sInfoWhere = sInfoWhere & ")" & vbCrLf
                End If
                If bSearchOwn Then
                    If mcSearchOwners.Count > 0 Then
                        mcSearchOwners.Add cwobjLogonEnviron.UserNCID, , 1
                        sInfoWhere = sInfoWhere & "))"
                    Else
                        mcSearchOwners.Add cwobjLogonEnviron.UserNCID
                        sInfoWhere = sInfoWhere & ")"
                    End If
                End If
            End If
        Case ClinicOwnerType
            sInfoWhere = sInfoWhere & " and type = " & tt_facility & " and isshared = 'Y'" & vbCrLf
            If mcSearchOwners.Count = 0 Then
                sInfoWhere = sInfoWhere & " and owner in (select clinic_ncid from user_clinic" & vbCrLf _
                        & "               where user_ncid = " & cwobjLogonEnviron.UserNCID & ")" & vbCrLf
            Else
                'SCR-47593 RBELL 1/13/04  Added processing to strip extra character if searching on a DMIS ID folder
                If Not IsNumeric(Left(mcSearchOwners(1), 1)) Then
                    sInfoWhere = sInfoWhere & " and owner in(" & Mid(mcSearchOwners(1), 2)
                    sClinicOwner = sClinicOwner & Mid(mcSearchOwners(1), 2)
                Else
                    sInfoWhere = sInfoWhere & " and owner in(" & mcSearchOwners(1)
                    sClinicOwner = sClinicOwner & mcSearchOwners(1)
                End If
                For i = 2 To mcSearchOwners.Count
                    sInfoWhere = sInfoWhere & ", " & mcSearchOwners(i)
                    sClinicOwner = sClinicOwner & ", " & mcSearchOwners(i)
                Next i
                sInfoWhere = sInfoWhere & ")" & vbCrLf
            End If
        Case MTFOwnerType
            sInfoWhere = sInfoWhere & " and type = " & tt_facility & vbCrLf
            If mcSearchOwners.Count = 0 Then
                sInfoWhere = sInfoWhere & " and owner = " & cwobjLogonEnviron.FacilityNCID & vbCrLf
            Else
                sInfoWhere = sInfoWhere & " and owner in(" & mcSearchOwners(1)
                For i = 2 To mcSearchOwners.Count
                    sInfoWhere = sInfoWhere & ", " & mcSearchOwners(i)
                Next i
                sInfoWhere = sInfoWhere & ")" & vbCrLf
            End If
        Case EnterpriseOwnerType
            sInfoWhere = sInfoWhere & " and type = " & tt_enterprise & vbCrLf
        Case Else
            ' By default, a user can see anything except other users' non-shared templates
            ' and clinic templates for clinics he is not assigned to
            If gobjShared.AppMode = modeCHCSI_GUI Then
                sInfoWhere = sInfoWhere & " and (" & vbCrLf _
                    & "(type = " & tt_user & " and (isshared = 'Y' or owner = " & cwobjLogonEnviron.UserNCID & "))" & vbCrLf _
                    & " or (type = " & tt_facility & " and owner in (select clinic_ncid from user_clinic where user_ncid = " & cwobjLogonEnviron.UserNCID & "))" & vbCrLf _
                    & " or (type = " & tt_facility & " and owner = " & cwobjLogonEnviron.FacilityNCID & "))" & vbCrLf
            Else
                sInfoWhere = sInfoWhere & " and (" & vbCrLf _
                    & "(type = " & tt_user & " and (isshared = 'Y' or owner = " & cwobjLogonEnviron.UserNCID & "))" & vbCrLf _
                    & " or (type = " & tt_facility & " and owner in (select clinic_ncid from user_clinic where user_ncid = " & cwobjLogonEnviron.UserNCID & "))" & vbCrLf _
                    & " or (type = " & tt_facility & " and owner = " & cwobjLogonEnviron.FacilityNCID & ")" & vbCrLf _
                    & " or (type = " & tt_enterprise & " and owner = " & cwobjLogonEnviron.EnterpriseNCID & "))" & vbCrLf
            End If
    End Select
    sXRefWhere = vbNullString
    sStr = moSearchComplaints.GetQueryCriteria
    If sStr <> vbNullString Then
        If sXRefWhere = vbNullString Then
            sXRefWhere = sStr & vbCrLf
        Else
            sXRefWhere = sXRefWhere & " or " & sStr & vbCrLf
        End If
    End If
    sStr = moSearchAppointmentTypes.GetQueryCriteria
    If sStr <> vbNullString Then
        If sXRefWhere = vbNullString Then
            sXRefWhere = sStr & vbCrLf
        Else
            sXRefWhere = sXRefWhere & " or " & sStr & vbCrLf
        End If
    End If
    sStr = moSearchProblems.GetQueryCriteria
    If sStr <> vbNullString Then
        If sXRefWhere = vbNullString Then
            sXRefWhere = sStr & vbCrLf
        Else
            sXRefWhere = sXRefWhere & " or " & sStr & vbCrLf
        End If
    End If
    bVirtTable = False
    If msSearchSpecialtyID <> vbNullString Then
        sStr = "(type_ncid = " & SpecialtyNCID & " and value_ncid = " & msSearchSpecialtyID & ")"
        If sXRefWhere = vbNullString Then
            sXRefWhere = sStr & vbCrLf
            
        Else
            bVirtTable = True
            'sXRefWhere = sXRefWhere & " or " & sStr & vbCrLf
        End If
    End If
    If mbSearchFavorites Then
        sFavorites = vbNullString
    Else
        If gobjShared.UseSQLServer And Not gobjShared.AppMode = modeCHCSI_GUI Then
            sFavorites = "*"
        Else
            sFavorites = "(+)"
        End If
    End If
    
    sQuery = "select distinct info.id, info.owner, info.isactive, info.type, info.isshared, info.name, " & _
                    "'0' as favorites_user, '0' as folder_id, facility.template_facility" & vbCrLf
    
    If sXRefWhere = vbNullString And Not bVirtTable Then
        sQuery = sQuery & " from template_info info, template_facility facility" & vbCrLf
    Else
        If WithXRefType Then
            sQuery = sQuery & ", type_ncid"
        End If
        
        'If we want to look at a subset of a table use virtual table as subset
        'Ex. look for a Dx in a Speciality SCR-11652
        If bVirtTable Then
            sStr = "type_ncid = " & SpecialtyNCID & " and value_ncid = " & msSearchSpecialtyID
            sQuery = sQuery & " from template_info info, template_facility facility," & vbCrLf & _
                "(select id, type_ncid, value_ncid from template_xref where id in " & vbCrLf & _
                "(select id from template_xref where " & sStr & "))" & " xref" & vbCrLf
            
        Else
            sQuery = sQuery & " from template_info info, template_facility facility, template_xref xref" & vbCrLf
        End If
        
    End If
    'Do outer join with Template_Facility to avoid having to call GetTemplateFacility method
    If gobjShared.UseSQLServer And Not gobjShared.AppMode = modeCHCSI_GUI Then
        sQuery = sQuery & " where info.id not in (select template_id from template_favorites where folder_id > 0) and info.id *= facility.id" _
            & sInfoWhere
    Else
        If meSearchOwnerType = ClinicOwnerType Then
            sQuery = sQuery & " where info.id not in (select template_id from template_favorites where folder_id > 0 and folder_id not in (" & _
                    sClinicOwner & ")) and info.id = facility.id(+)" & sInfoWhere
        Else
            sQuery = sQuery & " where info.id not in (select template_id from template_favorites where folder_id > 0) and info.id = facility.id(+)" _
                    & sInfoWhere
        End If
    End If
    
    If sXRefWhere <> vbNullString Then
        sQuery = sQuery & " and xref.id = info.id and (" & sXRefWhere & ")"
    End If
    
    If Not gobjShared.UseSQLServer And Not gobjShared.AppMode = modeCHCSI_GUI Then
       sQuery = sQuery & " order by lower(info.name), info.id"  '--- SCR 18151 Wang
    Else
       sQuery = sQuery & " order by info.name, info.id"  '--- SCR 18151 Wang
    End If

    'Debug.Print sQuery
    mobjSQL.Execute sQuery
    DoQueryWithoutFavorites = True
    Exit Function
ErrHandler:
    'Resume Next
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "TemplateOps.DoQueryWithoutFavorites", "TemplateClient", vbCritical)

End Function
    

Public Function DoSearch(Optional ByVal SearchType As TemplateTypeEnum = EncounterTemplateType) As Collection
    Dim objTemplate As Object
    On Error GoTo ErrHandler
    Set DoSearch = New Collection
    If DoQuery(SearchType, False) Then
        Do While Not mobjSQL.EOF
            Set objTemplate = NewTemplate(SearchType)
            objTemplate.Deserialize mobjSQL
            DoSearch.Add objTemplate, objTemplate.TemplateID
            mobjSQL.MoveNext
        Loop
    End If
    GoTo Cleanup
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "TemplateOps.DoSearch", "TemplateClient", vbCritical)
Cleanup:
    Set objTemplate = Nothing
End Function

Public Function DoSearchLite(Optional ByVal SearchType As TemplateTypeEnum = EncounterTemplateType, _
                                            Optional sFolderID As String = vbNullString) As Collection
                                            
    Dim objTemplate As TemplateLite
    Dim bResult As Boolean
    
    On Error GoTo ErrHandler
    Set DoSearchLite = New Collection
    If gobjShared Is Nothing Then Set gobjShared = New CWShared
    If moMTFs Is Nothing Then Set moMTFs = DoSearchMTFs
    'SCR-44176 RBELL 10/22/03 Used a Val function to check for empty strings or zeroes passed in a strings
    If val(sFolderID) = 0 Then
        'SCR-48609 RBELL 1/23/04 Pass in new parameter
        bResult = DoQuery(SearchType, True, sFolderID)
    Else
        bResult = DoFolderQuery(SearchType, True, sFolderID)
    End If
    If bResult Then
        Do While Not mobjSQL.EOF
            If Not objTemplate Is Nothing Then
                If objTemplate.TemplateID = mobjSQL("ID") Then
                    objTemplate.IsFavoriteTemplate mobjSQL("FAVORITES_USER")
                    GoTo NextSQL
                End If
            End If
            Set objTemplate = New TemplateLite
            Set objTemplate.SpecialtyDictionary = SpecialtyDictionary
            objTemplate.Deserialize mobjSQL, moMTFs
            
            If Not IsNull(mobjSQL("FOLDER_ID")) Then objTemplate.FavoriteFolderID = mobjSQL("FOLDER_ID")
            DoSearchLite.Add objTemplate, CStr(objTemplate.TemplateID)
NextSQL:
            If Not IsNull(mobjSQL("TEMPLATE_FACILITY")) Then objTemplate.TemplateFacilityNCID = mobjSQL("TEMPLATE_FACILITY")
            mobjSQL.MoveNext
        Loop
    End If
    GoTo Cleanup

ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "TemplateOps.DoSearchLite", "TemplateClient", vbCritical)
    'Resume Next

Cleanup:
    Set objTemplate = Nothing

End Function

Public Function DoStandaloneOrderSearchLite() As Collection
'scr-44211, scr-44212 RBELL 2/11/2004 Returns a collection of templates for the current user standalone order set
                                            
    Dim oConn As CHCSII_CONN.Conn
    Dim oSQL As ICHCSII_SQL.ISqlOps
    Dim objTemplate As TemplateLite
    Dim sSQL As String
    
    On Error GoTo ErrHandler
    Set DoStandaloneOrderSearchLite = New Collection
    Set oConn = New CHCSII_CONN.Conn
    Set oSQL = oConn.CHCSII_SQLOPS(Auto)
    If moMTFs Is Nothing Then Set moMTFs = DoSearchMTFs
    
    If gobjShared Is Nothing Then Set gobjShared = New CWShared
    If gobjShared.UseSQLServer Then
        'SCR-46910 5/17/2004 RBELL Added SQL Server specific query
        sSQL = "select info.id, info.owner, info.isactive, info.type, info.isshared, info.name, favorites.user_ncid as favorites_user, favorites.folder_id, facility.template_facility" & vbCrLf & _
            "from template_info info, template_favorites favorites, template_order_set orderset, template_facility facility" & vbCrLf & _
            "where favorites.template_id =* orderset.template_id" & vbCrLf & _
            "and favorites.user_ncid =* orderset.user_ncid" & vbCrLf & _
            "and info.id *= facility.id and  name like 'E%'" & vbCrLf & _
            "and info.id = orderset.template_id and" & vbCrLf & _
            "orderset.user_ncid = " & cwobjLogonEnviron.UserNCID & vbCrLf & _
            "order by lower(info.name), info.id"
    Else
        sSQL = "select distinct info.id, info.owner, info.isactive, info.type, info.isshared, info.name, favorites.user_ncid as favorites_user, favorites.folder_id, facility.template_facility" & vbCrLf & _
            "from template_info info, template_favorites favorites, template_order_set orderset, template_facility facility" & vbCrLf & _
            "where favorites.template_id (+) = orderset.template_id" & vbCrLf & _
            "and favorites.user_ncid (+) = orderset.user_ncid" & vbCrLf & _
            "and info.id = facility.id(+) and  name like 'E%'" & vbCrLf & _
            "and info.id = orderset.template_id and" & vbCrLf & _
            "orderset.user_ncid = " & cwobjLogonEnviron.UserNCID & vbCrLf & _
            "order by lower(info.name), info.id"
    End If
    
    oSQL.Execute sSQL
    If oSQL.BOF And oSQL.EOF Then Exit Function
    
    Do While Not oSQL.EOF
          If Not objTemplate Is Nothing Then
             If objTemplate.TemplateID = oSQL("ID") Then
                objTemplate.IsFavoriteTemplate oSQL("FAVORITES_USER")
                Call objTemplate.AdjustSource(oSQL)
                GoTo NextSQL
             End If
          End If
          
          Set objTemplate = New TemplateLite
          Set objTemplate.SpecialtyDictionary = SpecialtyDictionary
          objTemplate.Deserialize oSQL, moMTFs
            
          If Not IsNull(oSQL("FOLDER_ID")) Then objTemplate.FavoriteFolderID = oSQL("FOLDER_ID")
          DoStandaloneOrderSearchLite.Add objTemplate, CStr(objTemplate.TemplateID)
NextSQL:
          
          If Not IsNull(oSQL("TEMPLATE_FACILITY")) Then objTemplate.TemplateFacilityNCID = oSQL("TEMPLATE_FACILITY")
          oSQL.MoveNext
    Loop
    
    GoTo Cleanup
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "TemplateOps.DoStandaloneOrderSearchLite", "TemplateClient", vbCritical)
    'Resume Next
Cleanup:
    Set objTemplate = Nothing
    
End Function
Public Function DoLinkedTemplatesSearchLite(pTemplateID As String) As Collection
'scr- RBELL 2/13/2004 Returns a collection of templates that are linked to the template id passed in
                                            
    Dim oConn As CHCSII_CONN.Conn
    Dim oSQL As ICHCSII_SQL.ISqlOps
    Dim objTemplate As TemplateLite
    Dim sSQL As String
    
    On Error GoTo ErrHandler
    Set DoLinkedTemplatesSearchLite = New Collection
    If val(pTemplateID) = 0 Then GoTo Cleanup
    Set oConn = New CHCSII_CONN.Conn
    Set oSQL = oConn.CHCSII_SQLOPS(Auto)
    If moMTFs Is Nothing Then Set moMTFs = DoSearchMTFs
    
    'TABLE: TEMPLATE_LINKS
    'COLUMNS:
    'template_id Number(20)
    'linked_template_id Number(20)

    If gobjShared Is Nothing Then Set gobjShared = New CWShared
    If gobjShared.UseSQLServer Then
        sSQL = "select info.id, info.owner, info.isactive, info.type, info.isshared, info.name, favorites.user_ncid as favorites_user, favorites.folder_id, facility.template_facility" & vbCrLf & _
            "from template_info info, template_favorites favorites, template_links links, template_facility facility" & vbCrLf & _
            "where favorites.template_id =* links.template_id" & vbCrLf & _
            "and favorites.user_ncid = " & cwobjLogonEnviron.UserNCID & vbCrLf & _
            "and info.id *= facility.id " & vbCrLf & _
            "and info.id = links.linked_template_id and" & vbCrLf & _
            "links.template_id = " & pTemplateID & vbCrLf & _
            "order by lower(info.name), info.id"
    Else
        sSQL = "select distinct info.id, info.owner, info.isactive, info.type, info.isshared, info.name, favorites.user_ncid as favorites_user, favorites.folder_id, facility.template_facility" & vbCrLf & _
            "from template_info info, template_favorites favorites, template_links links, template_facility facility" & vbCrLf & _
            "where favorites.template_id (+) = links.template_id" & vbCrLf & _
            "and favorites.user_ncid (+) = " & cwobjLogonEnviron.UserNCID & vbCrLf & _
            "and info.id = facility.id(+)" & vbCrLf & _
            "and info.id = links.linked_template_id and" & vbCrLf & _
            "links.template_id = " & pTemplateID & vbCrLf & _
            "order by lower(info.name), info.id"
    End If
        
    oSQL.Execute sSQL
    If oSQL.BOF And oSQL.EOF Then Exit Function
    
    Do While Not oSQL.EOF
          If Not objTemplate Is Nothing Then
             If objTemplate.TemplateID = oSQL("ID") Then
                objTemplate.IsFavoriteTemplate oSQL("FAVORITES_USER")
                Call objTemplate.AdjustSource(oSQL)
                GoTo NextSQL
             End If
          End If
          
          Set objTemplate = New TemplateLite
          Set objTemplate.SpecialtyDictionary = SpecialtyDictionary
          objTemplate.Deserialize oSQL, moMTFs
            
          If Not IsNull(oSQL("FOLDER_ID")) Then objTemplate.FavoriteFolderID = oSQL("FOLDER_ID")
          DoLinkedTemplatesSearchLite.Add objTemplate, CStr(objTemplate.TemplateID)
NextSQL:
          
          If Not IsNull(oSQL("TEMPLATE_FACILITY")) Then objTemplate.TemplateFacilityNCID = oSQL("TEMPLATE_FACILITY")
          oSQL.MoveNext
    Loop
    
    GoTo Cleanup
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.DoLinkedTemplatesSearchLite", Err.Description
    'Resume Next
Cleanup:
    Set objTemplate = Nothing
    
End Function

Public Function DoSearchLite_Selection(Optional ByVal SearchType As TemplateTypeEnum = EncounterTemplateType, _
                                            Optional sFolderID As String = vbNullString) As Collection
                                            
    Dim objTemplate As TemplateLite
    Dim bResult As Boolean
    
    On Error GoTo ErrHandler
    
    Set DoSearchLite_Selection = New Collection
    If gobjShared Is Nothing Then Set gobjShared = New CWShared
    If moMTFs Is Nothing Then Set moMTFs = DoSearchMTFs
    
    If val(sFolderID) = 0 Then
        bResult = DoQueryWithoutFavorites(SearchType, True)
    Else
        bResult = DoFolderQuery(SearchType, True, sFolderID)
    End If
    If bResult Then
        Do While Not mobjSQL.EOF
            If Not objTemplate Is Nothing Then
                If objTemplate.TemplateID = mobjSQL("ID") Then
                    Call objTemplate.AdjustSource(mobjSQL)
                    GoTo NextSQL
                End If
            End If
            Set objTemplate = New TemplateLite
            Set objTemplate.SpecialtyDictionary = SpecialtyDictionary
            objTemplate.Deserialize mobjSQL, moMTFs
            
            If Not IsNull(mobjSQL("FOLDER_ID")) Then objTemplate.FavoriteFolderID = mobjSQL("FOLDER_ID")
            DoSearchLite_Selection.Add objTemplate, CStr(objTemplate.TemplateID)
NextSQL:
            'objTemplate.TemplateFacilityNCID = GetTemplateFacility(objTemplate.TemplateID)
            'Template Facility now returned by SQL query
            If Not IsNull(mobjSQL("TEMPLATE_FACILITY")) Then objTemplate.TemplateFacilityNCID = mobjSQL("TEMPLATE_FACILITY")
            mobjSQL.MoveNext
        Loop
    End If
    GoTo Cleanup
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "TemplateOps.DoSearchLite_Selection", "TemplateClient", vbCritical)
    'Resume Next
Cleanup:
    Set objTemplate = Nothing
End Function

Public Function BuildClinicList() As String
'scr-49546 RBELL 2/13/2004 Returns a collection of templates that are linked to the template id passed in
                                            
    Dim oConn As CHCSII_CONN.Conn
    Dim oSQL As ICHCSII_SQL.ISqlOps
    Dim sSQL As String
    Dim sTemp As String
    
    On Error GoTo ErrHandler
    BuildClinicList = "0" 'Default so SQL will not fail
    Set oConn = New CHCSII_CONN.Conn
    Set oSQL = oConn.CHCSII_SQLOPS(Auto)
    
    sSQL = "select clinic_ncid from user_clinic where user_ncid = " & cwobjLogonEnviron.UserNCID
    oSQL.Execute sSQL
    If oSQL.BOF And oSQL.EOF Then Exit Function
    
    Do While Not oSQL.EOF
          If Len(sTemp) = 0 Then
             sTemp = oSQL("clinic_ncid")
          Else
             sTemp = sTemp & ", " & oSQL("clinic_ncid")
          End If
          oSQL.MoveNext
    Loop
    
    If Len(sTemp) = 0 Then
       BuildClinicList = "0"
    Else
       BuildClinicList = sTemp
    End If
    
    GoTo Cleanup
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "TemplateOps.BuildClinicList", "TemplateClient", vbCritical)
    'Resume Next
Cleanup:
    Set oConn = Nothing
    Set oSQL = Nothing
    
End Function

Public Function GetFavoriteTemplates(bAllTemplateFavs As Boolean, Optional bIncludeDefault As Boolean = False) As Collection
    
    Dim oTemplateLite As TemplateLite
    
    Dim sBuf As String
    Dim bDefaultInFavorites As Boolean
    
    On Error GoTo ErrHandler
    ClearSearchCriteria
    mbSearchFavorites = True
    
    'SCR-51367 RBELL 3/25/2004 Reduced code by using new parameter to get all templates
    If bAllTemplateFavs = True Then
        Set GetFavoriteTemplates = DoSearchLite(AllTemplateTypes)
    Else
        Set GetFavoriteTemplates = DoSearchLite(EncounterTemplateType)
    End If
    
    If bIncludeDefault Then
       cwobjConfig.getValue NCID_TEMPLATE_MANAGEMENT, "Default", "DefaultTemplate", sBuf
       If Not sBuf = vbNullString Then
          bDefaultInFavorites = False
          For Each oTemplateLite In GetFavoriteTemplates
                If oTemplateLite.TemplateID = sBuf Then
                    bDefaultInFavorites = True
                    Exit For
                End If
          Next
          If Not bDefaultInFavorites Then
             Set oTemplateLite = GetTemplateLite(sBuf)
             If Not oTemplateLite Is Nothing Then GetFavoriteTemplates.Add oTemplateLite, CStr(oTemplateLite.TemplateID)
          End If
       End If
    End If
    
    If Not moTemplateHxOps Is Nothing Then moTemplateHxOps.IsCurrentEncounterTemplateValid = EncounterTemplateLoaded
    GoTo Cleanup
    
ErrHandler:
    'Resume Next
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "TemplateOps.GetFavoriteTemplates", "TemplateClient", vbCritical)
Cleanup:
    Set oTemplateLite = Nothing

End Function
Public Function GetCurrentEncounterTemplate(Optional ByRef eTemplateSection As TemplateSectionEnum = InvalidTemplateSection) As Template
    
    Dim objEncounter As CHCSIIEncounterOps.Encounter
    Dim objEncParent As CHCSIIEncounterCurrent.EncounterParent
    Dim objEncTemplate As Template
    
    On Error GoTo ErrHandler
    Set objEncParent = cwobjCWComm.InitializeOLEServer(NCID_ENCOUNTER_APPLICATION)
    Set objEncounter = objEncParent.CurrentEncounter
    If objEncounter Is Nothing Then
        MsgBxARMd "Could not get current encounter object.", , "Template Client"
        GoTo Cleanup
    End If
    Set objEncTemplate = NewTemplate(EncounterTemplateType)
    If objEncounter.TemplateData = vbNullString Then
        objEncTemplate.TemplateName = "No Encounter Template Loaded"
    Else
        objEncTemplate.Import objEncounter.TemplateData, eTemplateSection, True
        'SCR-51442 RBELL 3/24/2004 Set template name by default
        'scr-52896 rbell 5/11/2004 Do not rename copy forward template
        If Not UCase(Left(objEncTemplate.TemplateName, 12)) = "COPY FORWARD" Then objEncTemplate.TemplateName = "Current Encounter Template"
    End If
    Set GetCurrentEncounterTemplate = objEncTemplate
    GoTo Cleanup

ErrHandler:
    'Resume Next
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "TemplateOps.GetCurrentEncounterTemplate", "TemplateClient", vbCritical)

Cleanup:
    Set objEncParent = Nothing
    Set objEncounter = Nothing
    Set objEncTemplate = Nothing
End Function

Public Function GetDefaultTemplate() As Template
    
    Dim sBuf As String
    Dim sSQL As String
    
    Dim oConn As CHCSII_CONN.Conn
    Dim oSQL As ICHCSII_SQL.ISqlOps
    
    On Error GoTo ErrHandler
    
    Set oConn = New CHCSII_CONN.Conn
    Set oSQL = oConn.CHCSII_SQLOPS(Auto)
    
    'SCR 48941 - Used SQL instead of config objcet to only get the users defualt not the enterprise default
    sSQL = "SELECT Application_NCID, APSection, APKey, APValue" & vbCrLf & _
        "FROM OSH_INI_SETTINGS o" & vbCrLf & _
        "WHERE Application_NCID = 184913 AND" & vbCrLf & _
        "ApOwnerNCID = " & cwobjLogonEnviron.UserNCID & " AND" & vbCrLf & _
        "UPPER(APSection) = 'DEFAULT' AND" & vbCrLf & _
        "UPPER(APKey) = 'DEFAULTTEMPLATE'"
    
    oSQL.Execute sSQL
    If oSQL.BOF And oSQL.EOF Then
        GoTo CleanUpAndExit
    Else
        oSQL.MoveFirst
        'cwobjConfig.getValue NCID_TEMPLATE_MANAGEMENT, "Default", "DefaultTemplate", sBuf
        sBuf = oSQL("APValue")
    End If
    If Len(sBuf) = 0 Then Exit Function
    Set GetDefaultTemplate = GetTemplate(sBuf)
    
    GoTo CleanUpAndExit
    
ErrHandler:
    'Resume Next
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "TemplateOps.GetDefaultTemplate", "TemplateClient", vbCritical)
CleanUpAndExit:
    Set oConn = Nothing
    Set oSQL = Nothing

End Function


Public Function GetCurrentEncounterCPGExclusions() As Collection
        
    On Error GoTo ErrHandler
    
    Set GetCurrentEncounterCPGExclusions = New Collection
    
    GoTo Cleanup

ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "TemplateOps.GetCurrentEncounterCPGExclusions", "TemplateClient", vbCritical)
    'Resume Next
    
Cleanup:
    
End Function

Public Function GetTemplate(rsTemplateID As String, _
                                            Optional eTemplateSection As TemplateSectionEnum = InvalidTemplateSection, _
                                            Optional LoadLinkedTemplates As Boolean = True) As Object
'
'--- Tracker # 10695 - Add a parameter call to get individual sections of the template.
'--- Added the Enum Type 'TemplateSectionEnum' to the calls for 'GetTemplate'
'--- and 'NewTemplate' - RP
' SCR-51810 RBELL 3/30/2004 The new parameter LoadLinkedTemplates combined with the YT2DISPLAY switch to add the linked tempalte items to the template being retrieved
    Dim sQuery As String
    Dim eType As TemplateTypeEnum
    
    On Error GoTo ErrHandler
    
    If gobjShared Is Nothing Then Set gobjShared = New CWShared
    If rsTemplateID = "-1" Then
        Set GetTemplate = GetCurrentEncounterTemplate(eTemplateSection)
        Exit Function
    End If
    
    If moMTFs Is Nothing Then Set moMTFs = DoSearchMTFs
        sQuery = "select tf.template_facility, info.id, info.name, favorites.user_ncid as favorites_user, favorites.folder_id, " _
            & " info.isshared, info.isactive, info.type, info.owner, info.creator_ncid, info.modifior_ncid, favorites.folder_id" & vbCrLf _
            & " from template_facility tf LEFT JOIN template_info info ON tf.id = info.id LEFT JOIN template_favorites favorites ON favorites.template_id = info.id" & vbCrLf _
            & " " _
            & " " & vbCrLf _
            & " " & vbCrLf _
            & " Where info.id = " & rsTemplateID
    mobjSQL.Execute sQuery
    If Not mobjSQL.EOF Then
        If Left$(mobjSQL("NAME"), 1) = ListTemplatePrefix Then
            eType = ListTemplateType
        ElseIf Left$(mobjSQL("NAME"), 1) = EncTemplatePrefix Then
            eType = EncounterTemplateType
        ElseIf Left$(mobjSQL("NAME"), 1) = FormTemplatePrefix Then 'emc Forms support
            eType = FormTemplateType
        End If
        Set GetTemplate = NewTemplate(eType)
        GetTemplate.Deserialize mobjSQL, moMTFs
    End If
    
    If Not cwobjCWComm.CmdLineSwitch("YT2DISPLAY") = vbNullString Then
        If GetTemplate Is Nothing Then Exit Function
        If TypeOf GetTemplate Is Template Then
            If LoadLinkedTemplates Then GetTemplate.Merge_Template MergeLinkedTemplate(rsTemplateID)
        End If
    End If
    Exit Function
    
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "TemplateOps.GetTemplate", "TemplateClient", vbCritical)
    'Resume Next
    Exit Function
    Resume
End Function
Private Function MergeLinkedTemplate(rsTemplateID As String) As Object
' SCR-51810 RBELL 3/30/2004 The new function merges the linked template items to the template id being passed in
    
    Dim sQuery As String
    Dim sTemplateExport As String
    Dim oTemplate As Template
    Dim oMergedTemplate As Template
    Dim oListTemplate As ListTemplate
    
    On Error GoTo ErrHandler
    
    If rsTemplateID = "-1" Then Exit Function
    If gobjShared.UseSQLServer And Not gobjShared.AppMode = modeCHCSI_GUI Then
        sQuery = "select tf.template_facility, info.id, info.name, favorites.user_ncid as favorites_user, favorites.folder_id, " _
            & "info.isshared, info.isactive, info.type, info.owner, info.creator_ncid, info.modifior_ncid, favorites.folder_id" & vbCrLf _
            & "from template_facility tf,  template_info info, template_favorites favorites, template_links links" & vbCrLf _
            & "where tf.id =* info.id and " & vbCrLf _
            & "favorites.template_id =* info.id" & vbCrLf _
            & "and favorites.user_ncid =* info.owner " & vbCrLf _
            & "and info.id = links.linked_template_id " & vbCrLf _
            & "and links.template_id = " & rsTemplateID
    Else
        sQuery = "select tf.template_facility, info.id, info.name, favorites.user_ncid as favorites_user, favorites.folder_id, " _
            & " info.isshared, info.isactive, info.type, info.owner, info.creator_ncid, info.modifior_ncid, favorites.folder_id" & vbCrLf _
            & "from template_facility tf,  template_info info, template_favorites favorites, template_links links" & vbCrLf _
            & "where tf.id(+) = info.id and " & vbCrLf _
            & "favorites.template_id(+) = info.id" & vbCrLf _
            & "and favorites.user_ncid(+) = info.owner " & vbCrLf _
            & "and info.id = links.linked_template_id " & vbCrLf _
            & "and links.template_id = " & rsTemplateID
    End If
    mobjSQL.Execute sQuery
    If Not mobjSQL.EOF Then
       Set oMergedTemplate = New Template
       Set oMergedTemplate.EMCategoryDictionary = EMCategoryDictionary
       Set oMergedTemplate.SpecialtyDictionary = SpecialtyDictionary
       oMergedTemplate.TemplateFacilityNCID = cwobjLogonEnviron.FacilityNCID 'SCR-51810 5/7/2004 RBELL Set facility for orders to correctly merge
       Do Until mobjSQL.EOF
            Set oTemplate = New Template
            If Left$(mobjSQL("NAME"), 1) = ListTemplatePrefix Then
                Set oListTemplate = New ListTemplate
                oListTemplate.Deserialize mobjSQL
                With oTemplate
                    .TemplateID = oListTemplate.TemplateID
                    .TemplateName = vbNullString 'oListTemplate.TemplateName
                    .OwnerType = oListTemplate.OwnerType
                    .OwnerNCID = oListTemplate.OwnerNCID
                    .IsFavorite = oListTemplate.IsFavorite
                    .IsActive = oListTemplate.IsActive
                    .IsShared = oListTemplate.IsShared
                    sTemplateExport = oListTemplate.Export & gsListDelim
                    If Not Len(sTemplateExport) = 0 Then
                        If Not Left(sTemplateExport, 2) = "L," Then sTemplateExport = "L," & oListTemplate.Export & gsListDelim
                    End If
                    .NotesList.Import sTemplateExport
                End With
            ElseIf Left$(mobjSQL("NAME"), 1) = EncTemplatePrefix Then
                oTemplate.Deserialize mobjSQL
            ElseIf Left$(mobjSQL("NAME"), 1) = FormTemplatePrefix Then 'emc Forms support
                'eType = FormTemplateType
            End If
            If Not oTemplate Is Nothing Then
                oTemplate.TemplateFacilityNCID = cwobjLogonEnviron.FacilityNCID
                oMergedTemplate.Merge_Template oTemplate
            End If
            mobjSQL.MoveNext
        Loop
    End If
    
    Set MergeLinkedTemplate = oMergedTemplate
    GoTo Cleanup
    
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "TemplateOps.MergeLinkedTemplate", "TemplateClient", vbCritical)
    'Resume Next
Cleanup:
    Set oTemplate = Nothing

End Function

Public Function GetTemplateLite(rsTemplateID As String) As TemplateLite
    
    Dim sQuery As String
    
    On Error GoTo ErrHandler
    If gobjShared Is Nothing Then Set gobjShared = New CWShared
    If moMTFs Is Nothing Then Set moMTFs = DoSearchMTFs
    'SCR-46115 RBELL 12/2/3 If TemplateID not passed in return an empty TemplateLite object
    If val(rsTemplateID) = 0 Then
        Set GetTemplateLite = New TemplateLite
        Exit Function
    End If
    
    'SCR-46790 RBELL 12/16/03 Added additional fields to result set for deserialize routine
    If gobjShared.UseSQLServer Then
        sQuery = "select tf.template_facility, info.id, info.name, favorites.user_ncid as favorites_user," _
            & " info.isshared, info.isactive, info.type, info.owner, favorites.folder_id" & vbCrLf _
            & " from template_facility tf,  template_info info, template_favorites favorites" & vbCrLf _
            & " where tf.id =* info.id and " & vbCrLf _
            & "   favorites.template_id =* info.id" & vbCrLf _
            & "   and favorites.user_ncid = " & cwobjLogonEnviron.UserNCID & vbCrLf _
            & "   and info.id = " & rsTemplateID
    Else
        sQuery = "select tf.template_facility, info.id, info.name, favorites.user_ncid as favorites_user," _
            & " info.isshared, info.isactive, info.type, info.owner, favorites.folder_id" & vbCrLf _
            & " from template_facility tf,  template_info info, template_favorites favorites" & vbCrLf _
            & " where " _
            & "   tf.id(+) = info.id and " & vbCrLf _
            & "   favorites.template_id(+) = info.id" & vbCrLf _
            & "   and favorites.user_ncid(+) = " & cwobjLogonEnviron.UserNCID & vbCrLf _
            & "   and info.id = " & rsTemplateID
    End If
    mobjSQL.Execute sQuery
    If Not mobjSQL.EOF Then
'        If Left$(mobjSQL("NAME"), 1) = ListTemplatePrefix Then
'            eType = ListTemplateType
'        ElseIf Left$(mobjSQL("NAME"), 1) = EncTemplatePrefix Then
'            eType = EncounterTemplateType
'        ElseIf Left$(mobjSQL("NAME"), 1) = FormTemplatePrefix Then
'            eType = FormTemplateType 'emc Forms support
'        End If
        Set GetTemplateLite = New TemplateLite
        GetTemplateLite.Deserialize mobjSQL, moMTFs
    End If
    GoTo Cleanup
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "TemplateOps.GetTemplateLite", "TemplateClient", vbCritical)
    'Resume Next
Cleanup:
End Function

Public Function NewTemplate(ByVal TemplateType As TemplateTypeEnum) As Object
'
'--- Tracker # 10695 - Add a parameter call to get individual sections of the template.
'--- Added the Optional Enum Type 'TemplateSectionEnum' to the calls for 'GetTemplate'
'--- and 'NewTemplate' - RP
'
    Dim objEncTemplate As Template
    Dim objListTemplate As ListTemplate
    Dim objFormTemplate As FormTemplate 'emc Forms support
    
    On Error GoTo ErrHandler
    If gcwobjUser Is Nothing Then
        Err.Raise vbObjectError + 1, "CHCSII_TemplateClient.TemplateOps", "Initialize(...) must be called before using TemplateOps."
        Exit Function
    End If
    
    If TemplateType = EncounterTemplateType Then
        Set objEncTemplate = New Template
        Set objEncTemplate.EMCategoryDictionary = EMCategoryDictionary
        Set objEncTemplate.SpecialtyDictionary = SpecialtyDictionary
        Set NewTemplate = objEncTemplate
    ElseIf TemplateType = ListTemplateType Then
        Set objListTemplate = New ListTemplate
        Set objListTemplate.SpecialtyDictionary = SpecialtyDictionary
        Set NewTemplate = objListTemplate
'    ElseIf TemplateType = FormTemplateType Then 'emc Forms Support
'       Set objFormTemplate = New FormTemplate
'       Set objFormTemplate.SpecialtyDictionary = SpecialtyDictionary
'       Set NewTemplate = objFormTemplate
    End If
    
    Set objEncTemplate = Nothing
    Set objListTemplate = Nothing
    Set objFormTemplate = Nothing 'emc Forms support
    Exit Function

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.NewTemplate", Err.Description
       
End Function

Public Function AddToFavorites(rsTemplateID As String, rsUserNCID As String) As Boolean
    On Error GoTo ErrHandler
    mobjSQL.Execute "insert into template_favorites(user_ncid, template_id,folder_id) values (" _
        & SQLQuote(rsUserNCID) & ", " & SQLQuote(rsTemplateID) & ", 0)"
    Exit Function

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.AddToFavorites", Err.Description

End Function
Public Function AddToStandaloneOrderSets(rsTemplateID As String, rsUserNCID As String) As Boolean
'scr-44211, scr-44212 RBELL 2/11/2004 Adds an entry to the standalone order set collection
    On Error GoTo ErrHandler
    mobjSQL.Execute "insert into TEMPLATE_ORDER_SET(user_ncid, template_id) values (" _
        & SQLQuote(rsUserNCID) & ", " & SQLQuote(rsTemplateID) & ")"
    Exit Function

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.AddToStandaloneOrderSets", Err.Description

End Function
Public Function AddLinkedTemplate(rsTemplateID As String, rsLinkedTemplateID As String) As Boolean
'scr- RBELL 2/13/2004 Adds an entry to the linked template list
    
    Dim oConn As CHCSII_CONN.Conn
    Dim oSQL As ICHCSII_SQL.ISqlOps
    Dim sSQL As String
    
    On Error GoTo ErrHandler
    AddLinkedTemplate = False
    Set oConn = New CHCSII_CONN.Conn
    Set oSQL = oConn.CHCSII_SQLOPS(Auto)
    
    If gobjShared Is Nothing Then Set gobjShared = New CWShared
    sSQL = "select * " & vbCrLf & _
        "from template_links links" & vbCrLf & _
        "where links.linked_template_id = " & rsLinkedTemplateID & " and" & vbCrLf & _
        "links.template_id = " & rsTemplateID
    
    oSQL.Execute sSQL
    If Not (oSQL.BOF And oSQL.EOF) Then Exit Function
    
    mobjSQL.Execute "insert into TEMPLATE_LINKS(template_id, linked_template_id) values (" _
        & rsTemplateID & ", " & rsLinkedTemplateID & ")"
    AddLinkedTemplate = True
    GoTo Cleanup

ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "TemplateOps.AddLinkedTemplate", "TemplateClient", vbExclamation)
Cleanup:
    Set oSQL = Nothing
    Set oConn = Nothing

End Function

Public Function RemoveFromFavorites(rsTemplateID As String, rsUserNCID As String) As Boolean
    
    On Error GoTo ErrHandler
    mobjSQL.Execute "delete from template_favorites where" _
        & " user_ncid = " & SQLQuote(rsUserNCID) & " and template_id = " & SQLQuote(rsTemplateID)
    RemoveFromFavorites = (Err.Number = 0)
    Exit Function

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.RemoveFromFavorites", Err.Description
       
End Function
Public Function RemoveFromStandaloneOrderSets(rsTemplateID As String, rsUserNCID As String) As Boolean
'scr-44211, scr-44212 RBELL 2/11/2004 Removes an entry from the stndalone order sets based on the parameters passed in
    
    On Error GoTo ErrHandler
    mobjSQL.Execute "delete from TEMPLATE_ORDER_SET where" _
        & " user_ncid = " & SQLQuote(rsUserNCID) & " and template_id = " & SQLQuote(rsTemplateID)
    RemoveFromStandaloneOrderSets = (Err.Number = 0)
    Exit Function

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.RemoveFromStandaloneOrderSets", Err.Description
       
End Function
Public Function RemoveFromLinked(rsTemplateID As String, rsLinkedTemplateID As String) As Boolean
'scr- RBELL 2/13/2004 Removes an entry from the linked template list
    
    On Error GoTo ErrHandler
    mobjSQL.Execute "delete from TEMPLATE_LINKS where" _
        & " template_id = " & rsTemplateID & " and linked_template_id = " & rsLinkedTemplateID
    RemoveFromLinked = (Err.Number = 0)
    Exit Function

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.RemoveFromLinked", Err.Description
       
End Function

Public Function RemoveAllLinked(rsTemplateID As String) As Boolean
'scr- RBELL 2/13/2004 Removes an entry from the linked template list
    On Error GoTo ErrHandler
    mobjSQL.Execute "delete from TEMPLATE_LINKS where template_id = " & rsTemplateID
    RemoveAllLinked = (Err.Number = 0)
    Exit Function

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.RemoveAllLinked", Err.Description
       
End Function

Private Sub Class_Initialize()
    
    Set gobjShared = New CWShared
    Set moConn = New CHCSII_CONN.Conn
    Set mobjSQL = moConn.CHCSII_SQLOPS(Auto)
    
    Set mcSearchOwners = New Collection
    Set mcSearchTemplateType = New Collection 'SCR-49343 RBELL 2/9/4 Added template type support
    Set moSearchComplaints = New Complaints
    Set moSearchAppointmentTypes = New AppointmentTypes
    Set moSearchProblems = New Problems
    Set SpecialtyDictionary = New Scripting.Dictionary
    Set moTemplateHxOps = New TemplateHx
        
    SpecialtyDictionary.Add "1", "Aerospace/Aviation Medicine"
    SpecialtyDictionary.Add "2", "Allergy/Immunology"
    SpecialtyDictionary.Add "3", "Cardiology"
    SpecialtyDictionary.Add "4", "Cardiothoracic Surgery"
    SpecialtyDictionary.Add "5", "Dermatology"
    SpecialtyDictionary.Add "6", "Endocrinology"
    SpecialtyDictionary.Add "7", "Family Medicine"
    SpecialtyDictionary.Add "8", "Gastroenterology"
    SpecialtyDictionary.Add "9", "General Surgery"
    SpecialtyDictionary.Add "10", "Gynecology"
    SpecialtyDictionary.Add "11", "Hematology"
    SpecialtyDictionary.Add "12", "Hyperbaric/Undersea Medicine"
    SpecialtyDictionary.Add "13", "Infectious Disease"
    SpecialtyDictionary.Add "30", "Internal Medicine"
    SpecialtyDictionary.Add "14", "Nephrology"
    SpecialtyDictionary.Add "15", "Neurology"
    SpecialtyDictionary.Add "16", "Neurosurgery"
    SpecialtyDictionary.Add "17", "Obstetrics/Gynecology"
    SpecialtyDictionary.Add "18", "Occupational and Environmental Medicine"
    SpecialtyDictionary.Add "19", "Oncology"
    SpecialtyDictionary.Add "20", "Ophthalmology"
    SpecialtyDictionary.Add "21", "Orthopedics"
    SpecialtyDictionary.Add "22", "Otorhinolaryngology"
    SpecialtyDictionary.Add "23", "Pediatrics"
    SpecialtyDictionary.Add "24", "Preventive Medicine"
    SpecialtyDictionary.Add "25", "Psychiatric Disorders"
    SpecialtyDictionary.Add "26", "Pulmonary Medicine"
    SpecialtyDictionary.Add "27", "Rheumatology"
    SpecialtyDictionary.Add "28", "Tropical Medicine"
    SpecialtyDictionary.Add "29", "Urology"
    
    Set EMCategoryDictionary = New Scripting.Dictionary
    EMCategoryDictionary.Add "1", "TELEPHONE CONSULTS"
    EMCategoryDictionary.Add "2", "OFFICE / OUTPATIENT VISIT, EST"
    EMCategoryDictionary.Add "3", "OFFICE / OUTPATIENT VISIT, NEW"
    EMCategoryDictionary.Add "4", "OFFICE CONSULTATION"
    EMCategoryDictionary.Add "5", "CARE PLAN OVERSIGHT SERVICES"
    EMCategoryDictionary.Add "6", "CASE MANAGEMENT SERVICES (Team Conferences)"
    EMCategoryDictionary.Add "7", "EMERGENCY CARE"
    EMCategoryDictionary.Add "8", "OBSERVATION CARE"
    EMCategoryDictionary.Add "9", "OTHER E/M SERVICES"
    EMCategoryDictionary.Add "10", "PREVENTIVE MEDICINE SERVICES, NEW"
    EMCategoryDictionary.Add "11", "PREVENTIVE MEDICINE SERVICES, EST"
    EMCategoryDictionary.Add "12", "PREVENTIVE MEDICINE SERVICES, OTHER"
    EMCategoryDictionary.Add "13", "SPECIAL E/M SERVICES"
    EMCategoryDictionary.Add "14", "STANDBY SERVICES"
    
    '---set the help file for this application---
    App.HelpFile = gobjShared.CWHelpFile

End Sub

Private Sub Class_Terminate()
    
    'SCR 20618 Phil Crowder 1/17/02
    'Global reference no longer used.
    'Unload frmMedcin
    Set moTemplateHxOps = Nothing
    Set oEMCategoryDictionary = Nothing
    Set oSpecialtyDictionary = Nothing
    Set mobjSQL = Nothing
    Set mcSearchOwners = Nothing
    Set mcSearchTemplateType = Nothing 'SCR-49343 RBELL 2/9/4 Added template type support
    Set moSearchComplaints = Nothing
    Set moSearchAppointmentTypes = Nothing
    Set moSearchProblems = Nothing
    Set mobjConcept = Nothing
    
    Set goOESession = Nothing
    Set goOEGetSessionMgr = Nothing
    'Set cwobjCWComm = Nothing 'Have to leave initialized to prevent run-time error
    'Set cwobjPatient = Nothing 'Have to leave initialized to prevent run-time error
    'Set cwobjLogonEnviron = Nothing 'Have to leave initialized to prevent run-time error

    Set moComplaintList = Nothing
    Set moAppointmentList = Nothing
    Set moProblemList = Nothing
    Set moAutociteList = Nothing
    Set moNotesList = Nothing
    Set moClipartList = Nothing
    Set moDiagnosisList = Nothing
    Set moOrderList = Nothing
    Set moProcedureList = Nothing
    Set moInstructions = Nothing
    Set moForms = Nothing 'emc Forms Support

    Set gobjSnoConvert = Nothing
    Set moMTFs = Nothing
    
End Sub

Public Property Get EncounterTemplateLoaded() As Boolean
    
    Dim oEnc As Object
    Dim oEncounter As Object
    
    On Error GoTo ErrHandler
    EncounterTemplateLoaded = False
    
    Set oEnc = cwobjCWComm.InitializeOLEServer(NCID_ENCOUNTER_APPLICATION)
    If Not oEnc Is Nothing Then
        If Not oEnc.CurrentEncounter Is Nothing Then
           Set oEncounter = oEnc.CurrentEncounter
           If Len(oEncounter.TemplateData) > 0 Then EncounterTemplateLoaded = True
        End If
    End If
    
    GoTo Cleanup

ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "TemplateOps.EncounterTemplateLoaded", "TemplateClient", vbExclamation)
Cleanup:

    Set oEnc = Nothing
    Set oEncounter = Nothing

End Property

Private Function DoFolderQuery(ByVal SearchType As TemplateTypeEnum, _
                                               ByVal WithXRefType As Boolean, _
                                               sFolder_ID As String) As Boolean
'modification of DoQuery, this ones gets all templates with a specific folder ID
'optional param replaces the userNCID with the Enterprise NCID for Shared folders
'///scr34275 pjs 4.8.2003

    Dim sQuery As String
    Dim sStr As String
    Dim sWhere As String
    Dim sXRefWhere As String
    Dim sFavorites As String
    Dim sNamePrefix As String
    Dim sNamePattern As String
    Dim bVirtTable As Boolean 'SCR-11652
    
    If gcwobjUser Is Nothing Then
        Err.Raise vbObjectError + 1, "CHCSII_TemplateClient.TemplateOps", "Initialize(...) must be called before using TemplateOps."
        Exit Function
    End If
    If gobjShared Is Nothing Then Set gobjShared = New CWShared
    
    On Error GoTo ErrHandler
    
    If msSearchID <> vbNullString Then
        sWhere = " and id = " & msSearchID & vbCrLf
    End If
    
    Select Case SearchType
        Case EncounterTemplateType
            sNamePrefix = EncTemplatePrefix
        Case ListTemplateType
            sNamePrefix = ListTemplatePrefix
        Case FormTemplateType
            sNamePrefix = FormTemplatePrefix
    End Select
    
    If msSearchName = vbNullString Then
        sWhere = sWhere & " and name like "
        sNamePattern = vbNullString
    Else
        If mbSearchNameCaseInsensitive Then
            If gobjShared.UseSQLServer Then
                sWhere = sWhere & " and name like "
            Else
                sWhere = sWhere & " and upper(name) like "
            End If
            sNamePattern = UCase(msSearchName)
        Else
            sWhere = sWhere & " and name like "
            sNamePattern = msSearchName
        End If
        If meSearchNameType = ContainsNameSearch Then
            sNamePattern = "%" & sNamePattern
        End If
    End If
    
    'build the LIKE clause for the SQL
    sWhere = sWhere & SQLQuote(sNamePrefix & sNamePattern & "%") & vbCrLf
    
    'sWhere = sWhere & " and type = " & tt_enterprise & vbCrLf
    
    sXRefWhere = vbNullString
    sStr = moSearchComplaints.GetQueryCriteria
    If sStr <> vbNullString Then
        If sXRefWhere = vbNullString Then
            sXRefWhere = sStr & vbCrLf
        Else
            sXRefWhere = sXRefWhere & " or " & sStr & vbCrLf
        End If
    End If
    sStr = moSearchAppointmentTypes.GetQueryCriteria
    If sStr <> vbNullString Then
        If sXRefWhere = vbNullString Then
            sXRefWhere = sStr & vbCrLf
        Else
            sXRefWhere = sXRefWhere & " or " & sStr & vbCrLf
        End If
    End If
    sStr = moSearchProblems.GetQueryCriteria
    If sStr <> vbNullString Then
        If sXRefWhere = vbNullString Then
            sXRefWhere = sStr & vbCrLf
        Else
            sXRefWhere = sXRefWhere & " or " & sStr & vbCrLf
        End If
    End If
    bVirtTable = False
    If msSearchSpecialtyID <> vbNullString Then
        sStr = "(type_ncid = " & SpecialtyNCID & " and value_ncid = " & msSearchSpecialtyID & ")"
        If sXRefWhere = vbNullString Then
            sXRefWhere = sStr & vbCrLf
            
        Else
            bVirtTable = True
            'sXRefWhere = sXRefWhere & " or " & sStr & vbCrLf
        End If
    End If
    If mbSearchFavorites Then
        sFavorites = vbNullString
    Else
        If gobjShared.UseSQLServer Then
            sFavorites = "*"
        Else
            sFavorites = "(+)"
        End If
    End If
    sQuery = "select distinct info.id, info.owner, info.isactive, info.type, info.isshared, info.name, " & vbCrLf
    sQuery = sQuery & "favorites.user_ncid as favorites_user, favorites.folder_id, facility.template_facility" & vbCrLf
                
    If sXRefWhere = vbNullString And Not bVirtTable Then
        sQuery = sQuery & " from template_info info, template_favorites favorites, template_facility facility" & vbCrLf
    Else
        If WithXRefType Then
            sQuery = sQuery & ", type_ncid"
        End If
        
        'If we want to look at a subset of a table use virtual table as subset
        'Ex. look for a Dx in a Speciality SCR-11652
        If bVirtTable Then
            sStr = "type_ncid = " & SpecialtyNCID & " and value_ncid = " & msSearchSpecialtyID
            sQuery = sQuery & " from template_info info, template_favorites favorites, template_facility facility," & vbCrLf & _
                "(select id, type_ncid, value_ncid from template_xref where id in " & vbCrLf & _
                "(select id from template_xref where " & sStr & "))" & " xref" & vbCrLf
            
        Else
            sQuery = sQuery & " from template_info info, template_favorites favorites, template_facility facility, template_xref xref" & vbCrLf
        End If
        
    End If
    'Do outer join with Template_Facility to avoid having to call GetTemplateFacility method
    If gobjShared.UseSQLServer Then
        'SCR-50481 RBELL 3/3/2004 Removed outer join as unecessary and caused errors
        sQuery = sQuery & " where info.id = favorites.template_id" & vbCrLf
        sQuery = sQuery & "   and favorites.user_ncid = " & " info.Owner " & vbCrLf
        sQuery = sQuery & "   and favorites.folder_id" & " =" & sFolder_ID & vbCrLf
        sQuery = sQuery & "   and info.id *= facility.id" & sWhere
    Else
        sQuery = sQuery & " where info.id = favorites.template_id" & sFavorites & vbCrLf
        'SCR-44255 RBELL 3/16/2003 Used a joined table value rather than a logon object derivative to complete the join
        sQuery = sQuery & "   and favorites.user_ncid" & sFavorites & " = info.Owner " & vbCrLf
        sQuery = sQuery & "   and favorites.folder_id" & " =" & sFolder_ID & vbCrLf
        sQuery = sQuery & "   and info.id = facility.id(+)" & sWhere
    End If
    If sXRefWhere <> vbNullString Then
        sQuery = sQuery & " and xref.id = info.id and (" & sXRefWhere & ")"
    End If
    
    If Not gobjShared.UseSQLServer Then
       sQuery = sQuery & " order by lower(info.name), info.id"  '--- SCR 18151 Wang
    Else
       sQuery = sQuery & " order by info.name, info.id"  '--- SCR 18151 Wang
    End If
'Debug.Print sQuery
    mobjSQL.Execute sQuery
    DoFolderQuery = True
    Exit Function

ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "TemplateOps.DoFolderQuery", "TemplateClient", vbCritical)

End Function



Public Property Get SearchID() As String
    
    On Error GoTo ErrHandler
    SearchID = msSearchID
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.SearchID", Err.Description
    
End Property

Public Property Let SearchID(ByVal RHV As String)
    
    On Error GoTo ErrHandler
    msSearchID = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.SearchID", Err.Description
    
End Property

Public Property Get SearchName() As String
    
    On Error GoTo ErrHandler
    SearchName = msSearchName
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.SearchName", Err.Description
    
End Property

Public Property Let SearchName(ByVal RHV As String)
    
    On Error GoTo ErrHandler
    msSearchName = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.SearchName", Err.Description
    
End Property

Public Property Get SearchNameType() As TemplateNameSearchEnum
    
    On Error GoTo ErrHandler
    SearchNameType = meSearchNameType
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.SearchNameType", Err.Description
    
End Property

Public Property Let SearchNameType(ByVal RHV As TemplateNameSearchEnum)
    
    On Error GoTo ErrHandler
    meSearchNameType = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.SearchNameType", Err.Description
    
End Property

Public Property Get SearchNameCaseInsensitive() As Boolean
    
    On Error GoTo ErrHandler
    SearchNameCaseInsensitive = mbSearchNameCaseInsensitive
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.SearchNameCaseInsensitive", Err.Description
    
End Property

Public Property Let SearchNameCaseInsensitive(ByVal RHV As Boolean)
    
    On Error GoTo ErrHandler
    mbSearchNameCaseInsensitive = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.SearchNameCaseInsensitive", Err.Description
    
End Property

Public Property Get SearchShared() As Boolean
    
    On Error GoTo ErrHandler
    SearchShared = mbSearchShared
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.SearchShared", Err.Description
    
End Property

Public Property Let SearchShared(ByVal RHV As Boolean)
    
    On Error GoTo ErrHandler
    mbSearchShared = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.SearchShared", Err.Description
    
End Property

Public Property Get SearchOwnerType() As TemplateOwnerTypeEnum
    
    On Error GoTo ErrHandler
    SearchOwnerType = meSearchOwnerType
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.SearchOwnerType", Err.Description
    
End Property

Public Property Let SearchOwnerType(ByVal RHV As TemplateOwnerTypeEnum)
    
    On Error GoTo ErrHandler
    meSearchOwnerType = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.SearchOwnerType", Err.Description
    
End Property

Public Property Get SearchOwners() As Collection
    
    On Error GoTo ErrHandler
    Set SearchOwners = mcSearchOwners
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.SearchOwners", Err.Description
    
End Property

Public Property Let SearchOwners(ByVal RHV As Collection)
    
    On Error GoTo ErrHandler
    Set mcSearchOwners = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.SearchOwners", Err.Description
    
End Property

Public Property Get SearchTemplateType() As Collection
    
    On Error GoTo ErrHandler
    Set SearchTemplateType = mcSearchTemplateType
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.SearchTemplateType", Err.Description
    
End Property

Public Property Let SearchTemplateType(ByVal RHV As Collection)
    
    On Error GoTo ErrHandler
    Set mcSearchTemplateType = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.SearchTemplateType", Err.Description
    
End Property

Public Property Get SearchSpecialtyID() As String
    
    On Error GoTo ErrHandler
    SearchSpecialtyID = msSearchSpecialtyID
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.SearchSpecialtyID", Err.Description
    
End Property

Public Property Let SearchSpecialtyID(ByVal RHV As String)
    
    On Error GoTo ErrHandler
    msSearchSpecialtyID = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.SearchSpecialtyID", Err.Description
    
End Property

Public Property Get SearchComplaints() As TemplateData
    
    On Error GoTo ErrHandler
    Set SearchComplaints = moSearchComplaints
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.SearchComplaints", Err.Description
    
End Property

Public Property Let SearchComplaints(ByVal RHV As TemplateData)
    
    On Error GoTo ErrHandler
    Set moSearchComplaints = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.SearchComplaints", Err.Description
    
End Property

Public Property Get SearchAppointmentTypes() As TemplateData
    
    On Error GoTo ErrHandler
    Set SearchAppointmentTypes = moSearchAppointmentTypes
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.SearchAppointmentTypes", Err.Description
    
End Property

Public Property Let SearchAppointmentTypes(ByVal RHV As TemplateData)
    
    On Error GoTo ErrHandler
    Set moSearchAppointmentTypes = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.SearchAppointmentTypes", Err.Description
    
End Property

Public Property Get SearchProblems() As TemplateData
    
    On Error GoTo ErrHandler
    Set SearchProblems = moSearchProblems
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.SearchProblems", Err.Description
    
End Property

Public Property Let SearchProblems(ByVal RHV As TemplateData)
    
    On Error GoTo ErrHandler
    Set moSearchProblems = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.SearchProblems", Err.Description
    
End Property

Public Property Get SearchFavorites() As Boolean
    
    On Error GoTo ErrHandler
    SearchFavorites = mbSearchFavorites
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.SearchFavorites", Err.Description
    
End Property

Public Property Let SearchFavorites(ByVal RHV As Boolean)
    
    On Error GoTo ErrHandler
    mbSearchFavorites = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.SearchFavorites", Err.Description
    
End Property

Public Property Get SearchAdvanced() As Boolean
    
    On Error GoTo ErrHandler
    SearchAdvanced = mbSearchAdvanced
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.SearchAdvanced", Err.Description
    
End Property

Public Property Let SearchAdvanced(ByVal RHV As Boolean)
    
    On Error GoTo ErrHandler
    mbSearchAdvanced = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.SearchAdvanced", Err.Description
    
End Property

Public Property Get TemplateHxOps() As TemplateHx
    
    On Error GoTo ErrHandler
    Set TemplateHxOps = moTemplateHxOps
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.TemplateHxOps", Err.Description
    
End Property

Public Property Let TemplateHxOps(ByVal RHV As TemplateHx)
    
    On Error GoTo ErrHandler
    Set moTemplateHxOps = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.TemplateHxOps", Err.Description
    
End Property

Public Property Get EnterpriseReadAccess() As Boolean

    On Error GoTo ErrHandler
    EnterpriseReadAccess = mbEnterpriseReadAccess
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.EnterpriseReadAccess", Err.Description
    
End Property

Public Property Let EnterpriseReadAccess(ByVal RHV As Boolean)
    
    On Error GoTo ErrHandler
    mbEnterpriseReadAccess = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.EnterpriseReadAccess", Err.Description
    
End Property

Public Property Get MTFReadAccess() As Boolean
    
    On Error GoTo ErrHandler
    MTFReadAccess = mbMTFReadAccess
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.MTFReadAccess", Err.Description
    
End Property

Public Property Let MTFReadAccess(ByVal RHV As Boolean)
    
    On Error GoTo ErrHandler
    mbMTFReadAccess = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.MTFReadAccess", Err.Description
    
End Property

Public Property Get ClinicReadAccess() As Boolean
    
    On Error GoTo ErrHandler
    ClinicReadAccess = mbClinicReadAccess
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.ClinicReadAccess", Err.Description
    
End Property

Public Property Let ClinicReadAccess(ByVal RHV As Boolean)
    
    On Error GoTo ErrHandler
    mbClinicReadAccess = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.ClinicReadAccess", Err.Description
    
End Property

Public Property Get PersonalReadAccess() As Boolean
    
    On Error GoTo ErrHandler
    PersonalReadAccess = mbPersonalReadAccess
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.PersonalReadAccess", Err.Description
    
End Property

Public Property Let PersonalReadAccess(ByVal RHV As Boolean)
    
    On Error GoTo ErrHandler
    mbPersonalReadAccess = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateOps.PersonalReadAccess", Err.Description
    
End Property
