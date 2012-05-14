VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Template"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private oEMCategoryDictionary As Scripting.Dictionary
Private oSpecialtyDictionary As Scripting.Dictionary

Private Const msSerialVersion As String = "V3"

Private mobjEncTmplt As ENCTMPLTATLLib.EncounterTemplate
'GEMS    '*-- Coleman 060401: Add support for Theater
Private mobjEncTmplt_GEMS As GEMS_EncTmplt.EncounterTemplate
Private msName As String
Private msOwnerNCID As String
Private msOwnerName As String
Private mbDirty As Boolean

Private mbIsFavorite As Boolean
Private mbIsShared As Boolean
Private mbIsActive As Boolean
Private meOwnerType As TemplateOwnerTypeEnum
Private msTemplateID As String
Private msSpecialtyID As String
Private msEMCategoryID As String
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
Private msTemplateFacilityNCID As String
Private msCPG_ID As String
Private msFavoriteFolderID As String
Private msCreateUser As String    'SCR-47048 RBELL 12/17/03  ADDED
Private msModifyUser As String    'SCR-47048 RBELL 12/17/03  ADDED

Private mcLinkedTemplates As Collection
Public Property Get FavoriteFolderID() As String
    On Error GoTo ErrHandler
    FavoriteFolderID = msFavoriteFolderID
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.FavoriteFolderID", Err.Description

End Property
Public Property Let FavoriteFolderID(ByVal RHV As String)
    On Error GoTo ErrHandler
    msFavoriteFolderID = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.FavoriteFolderID", Err.Description

End Property

Public Property Get EMCategoryDictionary() As Scripting.Dictionary
    On Error GoTo ErrHandler
    Set EMCategoryDictionary = oEMCategoryDictionary
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.EMCategoryDictionary", Err.Description

End Property

Public Property Set EMCategoryDictionary(ByVal RHV As Scripting.Dictionary)
    On Error GoTo ErrHandler
    Set oEMCategoryDictionary = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.EMCategoryDictionary", Err.Description

End Property

Public Sub EnableCopyForward(pState As Boolean)
    
    On Error GoTo ErrHandler
    gCopyForwardNote = pState
    Exit Sub
    
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.EnableCopyForward", Err.Description
    
End Sub

'Public Function MaintainOrders() As Boolean
'
'    On Error GoTo ErrHandler
'    Screen.MousePointer = vbHourglass
'    Set frmOrderSets.mobjTemplate = Me
'    cwobjCWComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_TEMPLATE_MANAGEMENT
'    frmOrderSets.Display
'    cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_TEMPLATE_MANAGEMENT
'    Screen.MousePointer = vbDefault
'    Exit Function
'
'ErrHandler:
'    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.MaintainOrders", Err.Description
'
'End Function

Public Sub ResetAutoCiteList()
    
    On Error GoTo ErrHandler
    Set AutociteList = New AutoCites
    Exit Sub
    
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.ResetAutoCiteList", Err.Description

End Sub

Public Property Get SpecialtyDictionary() As Scripting.Dictionary
    
    On Error GoTo ErrHandler
    Set SpecialtyDictionary = oSpecialtyDictionary
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.SpecialtyDictionary", Err.Description

End Property

Public Property Set SpecialtyDictionary(ByVal RHV As Scripting.Dictionary)
    
    On Error GoTo ErrHandler
    Set oSpecialtyDictionary = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.SpecialtyDictionary", Err.Description

End Property

Public Property Get IsNew() As Boolean
    
    On Error GoTo ErrHandler
    If val(TemplateID) = 0 Then IsNew = True
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.IsNew", Err.Description

End Property

Public Function RemoveTemplateCPGXRef(pTemplateID As String) As Boolean

    Dim sSQL As String
    Dim sInsertSQL As String
    
    Dim iReason As Integer
    Dim sComment As String
    
    Dim objEncounter As CHCSIIEncounterOps.Encounter
    Dim objEncParent As CHCSIIEncounterCurrent.EncounterParent
    Dim moConn As CHCSII_CONN.Conn
    Dim oSQL As ICHCSII_SQL.ISqlOps
    
    On Error GoTo ErrHandler
    Set moConn = New CHCSII_CONN.Conn
    Set oSQL = moConn.CHCSII_SQLOPS(Auto)
    
    If cwobjCWComm.CmdLineSwitch("CPG") = "" Then Exit Function
    If gobjShared.IsAppMode(modeCHCSI_GUI) Then   '<SCR 36493
       RemoveTemplateCPGXRef = True
       Exit Function
    End If
    
    RemoveTemplateCPGXRef = False
    
    Set objEncParent = cwobjCWComm.InitializeOLEServer(NCID_ENCOUNTER_APPLICATION)
    Set objEncounter = objEncParent.CurrentEncounter
    If objEncounter Is Nothing Then
       MsgBxARMd "Could not get current encounter object.", , "Template Client"
       GoTo CleanUpAndExit
    End If
    
    iReason = 0
    sComment = ""
    
    sSQL = "SELECT * FROM CPG_TEMPLATE_REASON "
    sSQL = sSQL & "WHERE UNIT_NUMBER = '" & objEncounter.PatientID & "' AND "
    sSQL = sSQL & "ENCOUNTER_NUMBER = '" & objEncounter.EncounterID & "' AND "
    sSQL = sSQL & "FACILITY_NCID = '" & objEncounter.FacilityNCID & "' AND "
    sSQL = sSQL & "TEMPLATE_ID = '" & pTemplateID & "' "
    
    If Not (oSQL.BOF And oSQL.EOF) Then
       oSQL.MoveFirst
       iReason = oSQL("REASON_REMOVED")
       sComment = oSQL("COMMENTS")
    End If
    
    cwobjCWComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_TEMPLATE_MANAGEMENT
    Load frmCPGDeselectionReason
    frmCPGDeselectionReason.LoadReason iReason, sComment
    frmCPGDeselectionReason.Show vbModal
    If frmCPGDeselectionReason.Reason = vbNullString Then
       iReason = 0
    Else
       iReason = CInt(frmCPGDeselectionReason.Reason)
    End If
    sComment = frmCPGDeselectionReason.Comment
    Unload frmCPGDeselectionReason
    cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_TEMPLATE_MANAGEMENT
    
    If Not frmCPGDeselectionReason.Cancelled Then
       sSQL = "DELETE CPG_TEMPLATE_REASON "
       sSQL = sSQL & "WHERE UNIT_NUMBER = '" & objEncounter.PatientID & "' AND "
       sSQL = sSQL & "ENCOUNTER_NUMBER = '" & objEncounter.EncounterID & "' AND "
       sSQL = sSQL & "FACILITY_NCID = '" & objEncounter.FacilityNCID & "' AND "
       sSQL = sSQL & "TEMPLATE_ID = '" & pTemplateID & "' "
       
       sInsertSQL = "insert into CPG_TEMPLATE_REASON(UNIT_NUMBER, ENCOUNTER_NUMBER, FACILITY_NCID, TEMPLATE_ID, " _
                        & "REASON_REMOVED, COMMENTS, CREATOR_NCID, CREATE_TIME ) " & _
                        "values (" & SQLQuote(objEncounter.PatientID) & ", " & SQLQuote(objEncounter.EncounterID) & ", " & _
                        SQLQuote(objEncounter.FacilityNCID) & ", " & SQLQuote(pTemplateID) & ", " & SQLQuote(CStr(iReason)) & _
                        ", " & SQLQuote(sComment) & ", " & SQLQuote(cwobjLogonEnviron.UserNCID) & ", " & SQLQuote(Format(Now, gsDefaultDate)) & " )"
                                
        oSQL.Execute sSQL
        oSQL.Execute sInsertSQL
    Else
        GoTo CleanUpAndExit
    End If
    
    RemoveTemplateCPGXRef = True
    GoTo CleanUpAndExit

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.RemoveTemplateCPGXRef", Err.Description
    'Resume Next
CleanUpAndExit:
    Set objEncounter = Nothing
    Set objEncParent = Nothing
    Set oSQL = Nothing
    Set moConn = Nothing
    
End Function


Private Sub SaveTemplateCPGXref()

    Dim oSQL As ICHCSII_SQL.ISqlOps
    Dim moConn As CHCSII_CONN.Conn
    
    On Error GoTo ErrHandler
    Set moConn = New CHCSII_CONN.Conn
    Set oSQL = moConn.CHCSII_SQLOPS(Auto)
    
    oSQL.Execute "delete from cpg_template_xref where template_id = " & TemplateID
    oSQL.Execute "insert into cpg_template_xref(template_id, cpg_id) values (" _
                    & TemplateID & ", " & SQLQuote(CPG_ID) & ")"
    GoTo CleanUpAndExit
    
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.SaveTemplateCPGXref", Err.Description

CleanUpAndExit:
    Set oSQL = Nothing
    Set moConn = Nothing
    
End Sub

Private Sub SaveTemplateFacility()

    Dim oSQL As ICHCSII_SQL.ISqlOps
    Dim moConn As CHCSII_CONN.Conn
    
    Set moConn = New CHCSII_CONN.Conn
    Set oSQL = moConn.CHCSII_SQLOPS(Auto)
    
    On Error Resume Next    '*-- Ignore error on delete - no record there
    oSQL.Execute "delete from TEMPLATE_FACILITY where ID = " & TemplateID
    If Err.Number <> 0 Then
        Debug.Print Err.Number & ": " & Err.Description
    End If
    On Error GoTo ErrHandler
    oSQL.Execute "insert into TEMPLATE_FACILITY(ID, TEMPLATE_FACILITY) values (" _
                    & TemplateID & ", " & CStr(cwobjLogonEnviron.FacilityNCID) & ")"
    GoTo CleanUpAndExit
    
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.SaveTemplateFacility", Err.Description

CleanUpAndExit:
    Set oSQL = Nothing
    
End Sub

Public Property Get TemplateName() As String
    
    On Error GoTo ErrHandler
    If Len(msName) <= 1 Then Exit Property
    TemplateName = Mid(msName, 2)
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.TemplateName", Err.Description

End Property

Public Property Let TemplateName(ByRef rsName As String)
    
    On Error GoTo ErrHandler
    msName = EncTemplatePrefix & rsName
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.TemplateName", Err.Description

End Property

Public Property Get TemplateType() As TemplateTypeEnum
    
    On Error GoTo ErrHandler
    TemplateType = EncounterTemplateType
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.TemplateType", Err.Description

End Property
Friend Sub Initialize(robjLogon As MMMHISLogon.Logon, _
                        robjUser As CHCSII_ClientSecurity.ICHCSII_User, _
                        robjShared As MMMHISShared.CWShared, _
                        robjConcept As Object, _
                        robjPatient As MMMHISPatient.Patient, _
                        robjComm As Object, _
                        vlProgId)

End Sub

Public Property Get isDirty() As Boolean
    On Error GoTo ErrHandler
    If mbDirty Then
        isDirty = True
    ElseIf ComplaintList.isDirty Then
        isDirty = True
    ElseIf AppointmentList.isDirty Then
        isDirty = True
    ElseIf ProblemList.isDirty Then
        isDirty = True
    ElseIf AutociteList.isDirty Then
        isDirty = True
    ElseIf NotesList.isDirty Then
        isDirty = True
    ElseIf ClipartList.isDirty Then
        isDirty = True
    ElseIf DiagnosisList.isDirty Then
        isDirty = True
    ElseIf OrderList.isDirty Then
        isDirty = True
    ElseIf ProcedureList.isDirty Then
        isDirty = True
    ElseIf Instructions.isDirty Then
        isDirty = True
    End If
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.isDirty", Err.Description

End Property

Public Property Get EMCategoryName() As String
    
    On Error GoTo ErrHandler
    EMCategoryName = EMCategoryDictionary(EMCategoryID)
    Exit Property

ErrHandler:
    If Err = 5 Then
       EMCategoryName = vbNullString
       Exit Property
    End If
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.EMCategoryName", Err.Description

End Property

Public Property Get OwnerNCID() As String
    
    On Error GoTo ErrHandler
    OwnerNCID = msOwnerNCID
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.OwnerNCID", Err.Description

End Property

Public Property Let OwnerNCID(NCID As String)
    
    On Error GoTo ErrHandler
    msOwnerNCID = NCID
    msOwnerName = ""
    mbDirty = True
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.OwnerNCID", Err.Description

End Property

Public Property Get OwnerName() As String
    
    Dim objClinicOps As CHCSII_ClinicClient.ClinicOpsEx
    Dim objClinic As DTOs.Clinic
    
    On Error GoTo ErrHandler
    
    If msOwnerNCID <> "" And msOwnerName = "" Then
        If OwnerType = ClinicOwnerType Then
            On Error GoTo Cleanup
            Set objClinicOps = New CHCSII_ClinicClient.ClinicOpsEx
            Set objClinic = objClinicOps.GetClinic(msOwnerNCID)
            msOwnerName = objClinic.Name & " " & objClinic.MEPRSCode
        End If
    End If
    OwnerName = msOwnerName
    
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.OwnerName", Err.Description

Cleanup:
    Set objClinicOps = Nothing
    Set objClinic = Nothing
End Property
'<< GEMS
Public Property Get GEMS_OwnerName() As String
    
    Dim objConcept As GEMSConceptCtrl.GEMS_ConceptCtrl
    Dim objClinicOps As CHCSII_ClinicClient.ClinicOpsEx
    Dim objClinic As DTOs.Clinic
    
    On Error GoTo ErrHandler
    If msOwnerNCID <> "" And msOwnerName = "" Then
        If OwnerType = EnterpriseOwnerType Then
           msOwnerName = cwobjLogonEnviron.EnterpriseName
        ElseIf OwnerNCID = cwobjLogonEnviron.FacilityNCID And OwnerType = MTFOwnerType Then
            msOwnerName = cwobjLogonEnviron.FacilityName
        ElseIf OwnerNCID = cwobjLogonEnviron.NursingDivisionNCID And OwnerType = ClinicOwnerType Then
            msOwnerName = cwobjLogonEnviron.NursingDivisionName
        ElseIf OwnerType = ClinicOwnerType Then
            On Error GoTo Cleanup
            Set objClinicOps = New CHCSII_ClinicClient.ClinicOpsEx
            Set objClinic = objClinicOps.GetClinic(msOwnerNCID)
            msOwnerName = objClinic.Name & " " & objClinic.MEPRSCode
        Else
            Set objConcept = New GEMSConceptCtrl.GEMS_ConceptCtrl
            objConcept.UniqueId = msOwnerNCID
            If Not val(msOwnerNCID) = 0 Then msOwnerName = objConcept.PrefRep("2000").Representation
        End If
    End If
    GEMS_OwnerName = msOwnerName
    
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.GEMS_OwnerName", Err.Description

Cleanup:
    Set objClinicOps = Nothing
    Set objClinic = Nothing
    Set objConcept = Nothing

End Property

Public Property Get SpecialtyName() As String
    
    On Error GoTo ErrHandler
    SpecialtyName = SpecialtyDictionary(SpecialtyID)
    Exit Property

ErrHandler:
    If Err = 5 Then
       SpecialtyName = vbNullString
       Exit Property
    End If
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.SpecialtyName", Err.Description

End Property

Public Function LoadFromEncounter(ByRef robjEncounter As CHCSIIEncounterOps.Encounter) As Boolean
    
    
    On Error GoTo ErrHandler
    If gCopyForwardNote Then
        TemplateName = "Copy Forward " & robjEncounter.EncounterID
        TemplateFacilityNCID = robjEncounter.FacilityNCID
    Else
        TemplateName = "Created from Encounter " & robjEncounter.EncounterID
    End If
    
    Call AppointmentList.GetFromEncounter(robjEncounter)
    Call ComplaintList.GetFromEncounter(robjEncounter)
    Call ProblemList.GetFromEncounter(robjEncounter)
    Call ClipartList.GetFromEncounter(robjEncounter)
    Call NotesList.GetFromEncounter(robjEncounter)
    Call DiagnosisList.GetFromEncounter(robjEncounter)
    Call OrderList.GetFromEncounter(robjEncounter)
    Call ProcedureList.GetFromEncounter(robjEncounter)
    Call AutociteList.GetFromEncounter(robjEncounter)
    Call Instructions.GetFromEncounter(robjEncounter)
    mbDirty = True
    LoadFromEncounter = True
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.LoadFromEncounter", Err.Description

End Function

Public Function Merge(rsTemplateID As String) As Boolean
    
    On Error GoTo ErrHandler
        
    Merge = GEMS_Merge(rsTemplateID)
    
    GoTo Cleanup

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.Merge", Err.Description
    
Cleanup:

End Function
Public Function Merge_Template(oTemplate As Template) As Boolean
    
    On Error GoTo ErrHandler
    If oTemplate Is Nothing Then Exit Function
    Import_Section oTemplate.AppointmentList.Export, AppointmentListTemplateSection
    Import_Section oTemplate.ComplaintList.Export, ComplaintListTemplateSection
    Import_Section oTemplate.ProblemList.Export, ProblemListTemplateSection
    Import_Section oTemplate.AutociteList.Export, AutociteListTemplateSection
    Import_Section oTemplate.ClipartList.Export, ClipartListTemplateSection
    'SCR-51810 RBELL 4/8/2004 Must pass SO export strings in concatenated
    If Left(moNotesList.Export, 1) = FormTemplatePrefix Then
        Import_Section oTemplate.NotesList.Export & moNotesList.Export, NotesListTemplateSection
    Else
        Import_Section moNotesList.Export & oTemplate.NotesList.Export, NotesListTemplateSection
    End If
    Import_Section oTemplate.DiagnosisList.Export, DiagnosisListTemplateSection
    Import_Section oTemplate.ProcedureList.Export, ProcedureListTemplateSection
    Import_Section oTemplate.Instructions.Export, InstructionsTemplateSection
    If TemplateFacilityNCID = oTemplate.TemplateFacilityNCID Then
       Import_Section oTemplate.OrderList.Export, OrderListTemplateSection
    End If
    
    Exit Function

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.Merge_Template", Err.Description
    'Resume Next
    
End Function

Public Function Merge_Form_Template(oTemplate As FormTemplate) As Boolean
    
    On Error GoTo ErrHandler
    If oTemplate Is Nothing Then Exit Function
    Import_Section moNotesList.Export & FormTemplatePrefix & oTemplate.TemplateID & gsListDelim, NotesListTemplateSection
    Exit Function

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.Merge_Form_Template", Err.Description
    'Resume Next
    
End Function

'<< GEMS    '*-- Coleman 060601: Adding Theater support
Public Function GEMS_Merge(rsTemplateID As String) As Boolean
    
    Dim objEncTmplt As GEMS_EncTmplt.EncounterTemplate
    
    On Error GoTo ErrHandler
    Set objEncTmplt = New GEMS_EncTmplt.EncounterTemplate
    objEncTmplt.TemplateID = rsTemplateID
    objEncTmplt.Read
    
    msTemplateID = rsTemplateID
    msTemplateFacilityNCID = GetTemplateFacility(TemplateID)

    GEMS_DeserializeChildren objEncTmplt
    mbDirty = True
    GEMS_Merge = True
    GoTo Cleanup

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.GEMS_Merge", Err.Description

Cleanup:
    'Resume Next
    Set objEncTmplt = Nothing
End Function
' GEMS >>

Public Function Delete() As Boolean
    
    Dim mobjSQL As ICHCSII_SQL.ISqlOps
    Dim moConn As CHCSII_CONN.Conn
    
    Dim oCHCSInterface_Template As ICHCSTemplates
    Set oCHCSInterface_Template = gobjShared.CHCSConnection
    
    On Error GoTo ErrHandler
    
    If mobjEncTmplt_GEMS Is Nothing Then
       MsgBxARMd "Error: attempt to delete Encounter Template not properly initialized."
    ElseIf CheckUserPrivilege(gcwobjUser, msOwnerNCID, cwobjLogonEnviron.UserNCID, OwnerType, Priv_Delete, True) Then
       mobjEncTmplt_GEMS.Delete
       Set moConn = New CHCSII_CONN.Conn
       Set mobjSQL = moConn.CHCSII_SQLOPS(Auto)
         
        mobjSQL.Execute "select * from encounters where templateid = " & TemplateID
        If Not (mobjSQL.EOF = True And mobjSQL.BOF) Then
          Delete = False
          Exit Function
        End If
        mobjSQL.Execute "select * from template_DATA where id = " & TemplateID
        If Not mobjSQL.EOF = True Then
          mobjSQL.Execute "delete from template_DATA where id = " & TemplateID
        End If
                
        mobjSQL.Execute "select * from TEMPLATE_XREF where ID = " & TemplateID
        If Not mobjSQL.EOF = True Then
          mobjSQL.Execute "delete from TEMPLATE_XREF where ID = " & TemplateID
        End If
            
        mobjSQL.Execute "select * from TEMPLATE_FAVORITES where TEMPLATE_ID =" & TemplateID
        If Not mobjSQL.EOF = True Then
            mobjSQL.Execute "delete from TEMPLATE_FAVORITES where TEMPLATE_ID = " & TemplateID
        End If
            
        mobjSQL.Execute "select * from template_info where id = " & TemplateID
        If Not mobjSQL.EOF = True Then
           mobjSQL.Execute "delete from template_info where id = " & TemplateID
        End If
    
       Set mobjSQL = Nothing
       Delete = True
    End If
    
    Set mobjSQL = Nothing
    Exit Function
    
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "ListTemplate.Delete()", "TemplateClient", vbExclamation)

End Function

Public Function Deserialize(robjSQL As ICHCSII_SQL.ISqlOps, Optional pMTFs As ADODB.Recordset = Nothing) As Boolean
    
    Dim iFieldIndex As Integer
    
    On Error GoTo ErrHandler
    Deserialize = GEMS_Deserialize(robjSQL)
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.Deserialize", Err.Description

End Function
'<< GEMS    '*-- Coleman 060401: Add Theater support
Public Function GEMS_Deserialize(robjSQL As ICHCSII_SQL.ISqlOps) As Boolean
    
    Dim iFieldIndex As Integer
    
    On Error GoTo ErrHandler
    Set mobjEncTmplt_GEMS = New GEMS_EncTmplt.EncounterTemplate
    mobjEncTmplt_GEMS.TemplateID = robjSQL("ID")
    mobjEncTmplt_GEMS.Read
    TemplateID = mobjEncTmplt_GEMS.TemplateID
    
    Select Case UCase(TypeName(robjSQL("FAVORITES_USER")))
        Case "STRING"
            If CStr(robjSQL("FAVORITES_USER")) <> cwobjLogonEnviron.UserNCID Then
                IsFavorite = True
            Else
                IsFavorite = False
            End If
        Case "DECIMAL", "INTEGER", "LONG"
            If CLng(robjSQL("FAVORITES_USER")) > 0 Then
                IsFavorite = True
            Else
                IsFavorite = False
            End If
            
        Case Else
            IsFavorite = False
    End Select
    
    OwnerNCID = mobjEncTmplt_GEMS.owner
    IsActive = mobjEncTmplt_GEMS.IsActive
    Select Case mobjEncTmplt_GEMS.eType
    Case tt_facility
        If mobjEncTmplt_GEMS.IsShared Then
            OwnerType = ClinicOwnerType
        Else
            OwnerType = MTFOwnerType
        End If
        IsShared = True
    Case tt_enterprise
        IsShared = True
        OwnerType = EnterpriseOwnerType
    Case Else
        IsShared = mobjEncTmplt_GEMS.IsShared
        OwnerType = PersonalOwnerType
    End Select
    msName = mobjEncTmplt_GEMS.Name
            
    On Error Resume Next
    iFieldIndex = robjSQL.ColumnIndex("TEMPLATE_FACILITY")
    If Err = FieldNotInQueryResults Then
      msTemplateFacilityNCID = GetTemplateFacility(TemplateID)
    Else
      msTemplateFacilityNCID = robjSQL("TEMPLATE_FACILITY")
    End If
        
    msFavoriteFolderID = robjSQL("FOLDER_ID") & vbNullString
    msCreateUser = CStr(robjSQL("CREATOR_NCID") & vbNullString)  'SCR-47431 RBELL 12/23/03 Deserialize new field
    msModifyUser = CStr(robjSQL("MODIFIOR_NCID") & vbNullString) 'SCR-47431 RBELL 12/23/03 Deserialize new field
    '<< SCR 11975 'Save As Errors' - Fix: Use cwobjCWComm instead of mobjComm.
    '<< It seems that mobjComm is being set to Nothing before we're done with it.
    '<< To Do: Look into doing away with the local CW object variables (like mobjComm)
    '<< and always use the global CW ones instead.  This type of problem has popped up
    '<< several times now.
    CPG_ID = "0"
    
    GEMS_DeserializeChildren mobjEncTmplt_GEMS
    mbDirty = False
    GEMS_Deserialize = True
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.GEMS_Deserialize", Err.Description

End Function
' GEMS >>

Private Function DeserializeChildren(robjEncTmplt As ENCTMPLTATLLib.EncounterTemplate) As Boolean
    
    Dim objRef As ENCTMPLTATLLib.ETXref
    Dim colList As Collection
    Dim i As Integer
    
    On Error GoTo ErrHandler
    ' deserialize the cross-reference types
    Call AppointmentList.Deserialize(robjEncTmplt)
    Call ComplaintList.Deserialize(robjEncTmplt)
    Call ProblemList.Deserialize(robjEncTmplt)
    For Each objRef In robjEncTmplt.References
        If objRef.Type = SpecialtyNCID Then
            SpecialtyID = objRef.Value
            Exit For
        End If
    Next objRef
    Set objRef = Nothing
    
    'deserialize the encounter data
    For i = 1 To robjEncTmplt.Items.Count
        If robjEncTmplt.Items(i).Type = dt_blob Then
            Set colList = DeserializeString(robjEncTmplt.Items(i).dataBlob, gsMainDelim)
            Exit For
        End If
    Next i
    If Not colList Is Nothing Then
        If colList.Count > 0 Then
            For i = colList.Count + 1 To 9
                colList.Add ""
            Next i
            Select Case colList(1) ' this is the version stamp
            Case msSerialVersion
                EMCategoryID = colList(2)
                Call AutociteList.Deserialize(colList(3))
                Call ClipartList.Deserialize(colList(4))
                Call NotesList.Deserialize(colList(5))
                Call DiagnosisList.Deserialize(colList(6))
                Call ProcedureList.Deserialize(colList(7))
                If msTemplateFacilityNCID = cwobjLogonEnviron.FacilityNCID Then
                   Call OrderList.Deserialize(colList(8))
                Else
                   Call OrderList.Deserialize("")
                End If
                Call Instructions.Deserialize(colList(9))
            Case "V2"
                EMCategoryID = colList(2)
                Call AutociteList.Deserialize(colList(3))
                Call ClipartList.Deserialize(colList(4))
                Call NotesList.Deserialize(colList(5))
                Call DiagnosisList.Deserialize(colList(6))
                Call ProcedureList.Deserialize(colList(7))
                If msTemplateFacilityNCID = cwobjLogonEnviron.FacilityNCID Then
                   Call OrderList.Deserialize(colList(8))
                Else
                   Call OrderList.Deserialize("")
                End If
            Case Else
                MsgBxARMd "Unexpected template serialization version '" & colList(1) & "'."
            End Select
        End If
        Set colList = Nothing
    End If
    
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.DeserializeChildren", Err.Description

End Function
'<< GEMS    '*-- Coleman 060601: Adding Theater support
Private Function GEMS_DeserializeChildren(robjEncTmplt As _
                            GEMS_EncTmplt.EncounterTemplate) As Boolean
    
    Dim objRef As GEMS_EncTmplt.ETXref
    Dim colList As Collection
    Dim i As Long
    
    On Error GoTo ErrHandler
    ' deserialize the cross-reference types
    Call AppointmentList.Deserialize(robjEncTmplt)
    Call ComplaintList.Deserialize(robjEncTmplt)
    Call ProblemList.Deserialize(robjEncTmplt)
    For Each objRef In robjEncTmplt.References
        If objRef.eType = SpecialtyNCID Then
            SpecialtyID = objRef.Value
            Exit For
        End If
    Next objRef
    Set objRef = Nothing
    
    'deserialize the encounter data
    For i = 1 To robjEncTmplt.Items.Count
        If robjEncTmplt.Items(i).eType = dt_blob Then
            Set colList = DeserializeString(robjEncTmplt.Items(i).dataBlob, _
                                            gsMainDelim)
            Exit For
        End If
    Next i
    If Not colList Is Nothing Then
        If colList.Count > 0 Then
            For i = colList.Count + 1 To 9
                colList.Add ""
            Next i
            Select Case colList(1) ' this is the version stamp
            Case msSerialVersion
                EMCategoryID = colList(2)
                Call AutociteList.Deserialize(colList(3))
                Call ClipartList.Deserialize(colList(4))
                Call NotesList.Deserialize(colList(5))
                Call DiagnosisList.Deserialize(colList(6))
                Call ProcedureList.Deserialize(colList(7))
                If msTemplateFacilityNCID = cwobjLogonEnviron.FacilityNCID Then
                   Call OrderList.Deserialize(colList(8))
                Else
                   Call OrderList.Deserialize("")
                End If
                Call Instructions.Deserialize(colList(9))
            Case "V2"
                EMCategoryID = colList(2)
                Call AutociteList.Deserialize(colList(3))
                Call ClipartList.Deserialize(colList(4))
                Call NotesList.Deserialize(colList(5))
                Call DiagnosisList.Deserialize(colList(6))
                Call ProcedureList.Deserialize(colList(7))
                If msTemplateFacilityNCID = cwobjLogonEnviron.FacilityNCID Then
                   Call OrderList.Deserialize(colList(8))
                Else
                   Call OrderList.Deserialize("")
                End If
            Case Else
                MsgBxARMd "Unexpected template serialization version '" & _
                        colList(1) & "'."
            End Select
        End If
        Set colList = Nothing
    End If
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.GEMS_DeserializeChildren", Err.Description

End Function
' GEMS >>

Public Function Export() As String
    
    Dim colList As Collection
    Dim colHeader As Collection
    
    On Error GoTo ErrHandler
    Set colHeader = New Collection
    Set colList = New Collection
    colHeader.Add msSerialVersion
    colHeader.Add msName
    colHeader.Add OwnerType
    colHeader.Add IsShared
    colHeader.Add SpecialtyID
    colHeader.Add EMCategoryID
    If TemplateFacilityNCID = "" Then
        TemplateFacilityNCID = "0"
    End If
    colHeader.Add TemplateFacilityNCID
    colHeader.Add TemplateID
    colList.Add SerializeCollection(colHeader, gsListDelim)
    colList.Add AppointmentList.Export()
    colList.Add ComplaintList.Export()
    colList.Add ProblemList.Export()
    colList.Add AutociteList.Export()
    colList.Add ClipartList.Export()
    colList.Add NotesList.Export()
    colList.Add DiagnosisList.Export()
    colList.Add ProcedureList.Export()
    colList.Add OrderList.Export()
    colList.Add Instructions.Export()
    Export = SerializeCollection(colList, gsMainDelim)
    
    GoTo Cleanup

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.Export", Err.Description

Cleanup:
    'Resume Next
    Set colList = Nothing
    Set colHeader = Nothing

End Function

Public Function Import(rsString As String, Optional ByRef eTemplateSection As TemplateSectionEnum = InvalidTemplateSection, Optional SupressMsgs As Boolean = False) As Boolean
    
    Dim colList As Collection
    Dim colHeader As Collection
    
    On Error GoTo ErrHandler
    If gobjShared Is Nothing Then Set gobjShared = New CWShared
    
    If rsString = "" Then Exit Function
    Set colList = DeserializeString(rsString, gsMainDelim)
    If colList.Count = 0 Then
       GoTo Cleanup
    End If
    
    Set colHeader = DeserializeString(colList(1), gsListDelim)
    If colHeader.Count = 0 Then
       GoTo Cleanup
    ElseIf colHeader(1) <> msSerialVersion Then
       If Not SupressMsgs Then MsgBxARMd "Imported string version ('" & colHeader(1) & "') is not valid."
       GoTo Cleanup
    End If
    
    msName = colHeader(2)
    OwnerType = colHeader(3)
    IsShared = colHeader(4)
    SpecialtyID = colHeader(5)
    EMCategoryID = colHeader(6)
    If colHeader.Count >= 7 Then TemplateFacilityNCID = colHeader(7)
    'SCR-44889 RBELL 11/10/03 COMMENTED OUT LINE Line was added to support Ed Clarke no longer needed as Ed changed his design
    ' this line also causes an issue on importing templates
    'If colHeader.Count >= 8 Then TemplateID = colHeader(8)
        
    If eTemplateSection = AppointmentListTemplateSection Or eTemplateSection = InvalidTemplateSection Then AppointmentList.Import colList(2)
    If eTemplateSection = ComplaintListTemplateSection Or eTemplateSection = InvalidTemplateSection Then ComplaintList.Import colList(3)
    If eTemplateSection = ProblemListTemplateSection Or eTemplateSection = InvalidTemplateSection Then ProblemList.Import colList(4)
    If eTemplateSection = AutociteListTemplateSection Or eTemplateSection = InvalidTemplateSection Then AutociteList.Import colList(5)
    If eTemplateSection = ClipartListTemplateSection Or eTemplateSection = InvalidTemplateSection Then ClipartList.Import colList(6)
    If eTemplateSection = NotesListTemplateSection Or eTemplateSection = InvalidTemplateSection Then NotesList.Import colList(7)
    If eTemplateSection = DiagnosisListTemplateSection Or eTemplateSection = InvalidTemplateSection Then DiagnosisList.Import colList(8)
    If eTemplateSection = ProcedureListTemplateSection Or eTemplateSection = InvalidTemplateSection Then ProcedureList.Import colList(9)
    If eTemplateSection = InstructionsTemplateSection Or eTemplateSection = InvalidTemplateSection Then Instructions.Import colList(11)
    
    If Not TemplateID = "" Then TemplateFacilityNCID = GetTemplateFacility(TemplateID)
    'SCR-50853 RBELL 3/8/2004 Changed from GEMSDB function to IsAppMode function
    'If GemsDB Then TemplateFacilityNCID = cwobjLogonEnviron.FacilityNCID  '<SCR 37645 Alberstein
    If gobjShared.IsAppMode(modeTheater) Then TemplateFacilityNCID = cwobjLogonEnviron.FacilityNCID  '<SCR 37645 Alberstein
    If (eTemplateSection = OrderListTemplateSection Or eTemplateSection = InvalidTemplateSection) Then
       If colHeader.Count < 7 Then
          If Not colList(10) = vbNullString Then
             If Not SupressMsgs Then MsgBxARMd "Cannot import order sets from the specified template, the template does not specify the MTF the template originated with.", vbOKOnly + vbInformation
          End If
       'SCR-50853 RBELL 3/8/2004 Changed from GEMSDB function to IsAppMode function
       'ElseIf colHeader(7) = cwobjLogonEnviron.FacilityNCID Or GemsDB Then '<SCR 37645 Alberstein
       ElseIf colHeader(7) = cwobjLogonEnviron.FacilityNCID Or gobjShared.IsAppMode(modeTheater) Then '<SCR 37645 Alberstein
          If eTemplateSection = OrderListTemplateSection Or eTemplateSection = InvalidTemplateSection Then OrderList.Import colList(10)
       ElseIf Not colList(10) = "" Then
          If Not (TemplateID = "" And TemplateFacilityNCID = "") Then
             If Not SupressMsgs Then MsgBxARMd "Cannot import order sets from another MTF, the order sets contained within this template will be discarded.", vbOKOnly + vbInformation
          End If
       End If
    End If
    
    Import = True
    mbDirty = True
    GoTo Cleanup
    
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.Import", Err.Description
    'Resume Next

Cleanup:
    Set colList = Nothing
    Set colHeader = Nothing
End Function
Public Function ImportCHCSOrders(pOrderData As String) As Boolean
    
    Dim oOrders As Orders
    
    On Error GoTo ErrHandler
    If gobjShared Is Nothing Then Set gobjShared = New CWShared
    
    If Len(pOrderData) = 0 Then Exit Function
    Set oOrders = OrderList
    oOrders.DeserializeCHCSOrders pOrderData
    'OrderList.Import colList(10)
    
    ImportCHCSOrders = True
    Exit Function

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.ImportCHCSOrders", Err.Description
    'Resume Next
    
End Function

Public Function Import_Section(rsString As String, ByRef eTemplateSection As TemplateSectionEnum) As Boolean
    
    On Error GoTo ErrHandler
    If gobjShared Is Nothing Then Set gobjShared = New CWShared
    
    If rsString = "" Then Exit Function
    
    If eTemplateSection = AppointmentListTemplateSection Then AppointmentList.Import rsString
    If eTemplateSection = ComplaintListTemplateSection Then ComplaintList.Import rsString
    If eTemplateSection = ProblemListTemplateSection Then ProblemList.Import rsString
    If eTemplateSection = AutociteListTemplateSection Then AutociteList.Import rsString
    If eTemplateSection = ClipartListTemplateSection Then ClipartList.Import rsString
    If eTemplateSection = NotesListTemplateSection Then NotesList.Import rsString
    If eTemplateSection = DiagnosisListTemplateSection Then DiagnosisList.Import rsString
    If eTemplateSection = ProcedureListTemplateSection Then ProcedureList.Import rsString
    If eTemplateSection = InstructionsTemplateSection Then Instructions.Import rsString
    If eTemplateSection = OrderListTemplateSection Or eTemplateSection = InvalidTemplateSection Then OrderList.Import rsString
    
    Import_Section = True
    GoTo Cleanup
    
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.Import_Section", Err.Description

Cleanup:
    'Resume Next
    mbDirty = True

End Function

Public Function Save() As Boolean
    
    On Error GoTo ErrHandler
    
    If gobjShared Is Nothing Then Set gobjShared = New CWShared
    
    Save = GEMS_Save
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.Save", Err.Description

End Function
Private Function GetTemplateFacility(pTemplateID As String) As String
    
    Dim oSQL As ICHCSII_SQL.ISqlOps
    Dim moConn As CHCSII_CONN.Conn
    Dim sSQL As String
            
    On Error GoTo ErrHandler
    GetTemplateFacility = vbNullString
    Set moConn = New CHCSII_CONN.Conn
    Set oSQL = moConn.CHCSII_SQLOPS(Auto)
    
    sSQL = "SELECT TEMPLATE_FACILITY FROM TEMPLATE_FACILITY WHERE ID = " & pTemplateID
    oSQL.Execute sSQL
    
    If Not (oSQL.BOF And oSQL.EOF) Then
       If Not oSQL("TEMPLATE_FACILITY") = "" Then
          GetTemplateFacility = oSQL("TEMPLATE_FACILITY")
       End If
    End If
    
    GoTo CleanUpAndExit
    
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.GetTemplateFacility", Err.Description

CleanUpAndExit:
    Set oSQL = Nothing
    Set moConn = Nothing
    
End Function


' << GEMS    '*-- Coleman 060401:
Public Function GEMS_Save() As Boolean
    
    Dim colList As Collection
    Dim objRef As GEMS_EncTmplt.ETXref
    Dim objSQL As ICHCSII_SQL.ISqlOps '<< GEMS    '*-- Coleman 060401: Add Theater support
    Dim moConn As CHCSII_CONN.Conn
    
    On Error GoTo ErrHandler
    If Not CheckUserPrivilege(gcwobjUser, msOwnerNCID, _
                cwobjLogonEnviron.UserNCID, OwnerType, Priv_Write, True) Then
        Exit Function
    End If
    If mobjEncTmplt_GEMS Is Nothing Then
        Set mobjEncTmplt_GEMS = New GEMS_EncTmplt.EncounterTemplate
        mobjEncTmplt_GEMS.ClearAll
    Else
        mobjEncTmplt_GEMS.Items.ClearAll
        mobjEncTmplt_GEMS.References.ClearAll
    End If
    With mobjEncTmplt_GEMS
        'SCR-45317 RBELL 11/24/03 ADDED VAL function so that the create_user gets set properly
        If val(TemplateID) = 0 Then
            .TemplateID = 0
            .create_user = cwobjLogonEnviron.UserNCID 'SCR-44085 RBELL 10/14/03 Set to establish favorites correctly for Theater
        Else
            .TemplateID = TemplateID
            .create_user = msCreateUser 'SCR-47048 RBELL 12/17/03  ADDED
            .modify_user = msModifyUser 'SCR-47048 RBELL 12/17/03  ADDED
        End If
        .Name = msName
        .owner = msOwnerNCID
        .IsActive = IsActive
        .IsFavorite = IsFavorite
        .FavoriteFolderID = msFavoriteFolderID
        
        Select Case OwnerType
        Case ClinicOwnerType
            .eType = tt_facility        '*-- Coleman 060601: field name change
            .IsShared = True
        Case MTFOwnerType
            .eType = tt_facility        '*-- Coleman 060601: field name change
            .IsShared = False
        Case EnterpriseOwnerType
            .eType = tt_enterprise      '*-- Coleman 060601: field name change
            .IsShared = False
        Case Else
            .eType = tt_user            '*-- Coleman 060601: field name change
            .IsShared = IsShared
        End Select
    End With
    ' specialty is a singleton cross reference
    If SpecialtyID <> "" Then
        Set objRef = mobjEncTmplt_GEMS.References.Add
        objRef.eType = SpecialtyNCID
        objRef.Value = SpecialtyID
        Set objRef = Nothing
    End If
    'serialize the cross-reference type lists
    Call AppointmentList.Serialize(mobjEncTmplt_GEMS)
    Call ComplaintList.Serialize(mobjEncTmplt_GEMS)
    Call ProblemList.Serialize(mobjEncTmplt_GEMS)
    'serialize the encounter data types
    Set colList = New Collection
    colList.Add msSerialVersion
    colList.Add EMCategoryID
    Call AutociteList.Serialize(colList)
    Call ClipartList.Serialize(colList)
    Call NotesList.Serialize(colList)
    Call DiagnosisList.Serialize(colList)
    Call ProcedureList.Serialize(colList)
    Call OrderList.Serialize(colList)
    Call Instructions.Serialize(colList)
    mobjEncTmplt_GEMS.Items.Add
    mobjEncTmplt_GEMS.Items(mobjEncTmplt_GEMS.Items.Count).eType = dt_blob
    mobjEncTmplt_GEMS.Items(mobjEncTmplt_GEMS.Items.Count).dataBlob = SerializeCollection(colList, gsMainDelim)
    mobjEncTmplt_GEMS.Save
    TemplateID = mobjEncTmplt_GEMS.TemplateID
    GEMS_Save = True
    SaveTemplateFacility
    
    
    Set moConn = New CHCSII_CONN.Conn
    Set objSQL = moConn.CHCSII_SQLOPS(Auto)
    If IsFavorite Then
        'Added folder_id to insert stmt SCR-17931
'        bSaveFavorites = True
        'SCR-51169 RBELL 3/16/2004 Added a delete to clear the favorite record in case it already exists
        objSQL.Execute "DELETE FROM template_favorites where template_id = " & TemplateID & " and user_ncid = " & cwobjLogonEnviron.UserNCID
        If meOwnerType = PersonalOwnerType Then
            objSQL.Execute "insert into template_favorites(user_ncid, template_id, folder_id) values (" _
                & SQLQuote(cwobjLogonEnviron.UserNCID) & ", " & SQLQuote(TemplateID) & ", " & CStr(val(msFavoriteFolderID)) & ")"
        Else
            objSQL.Execute "insert into template_favorites(user_ncid, template_id, folder_id) values (" _
                & SQLQuote(cwobjLogonEnviron.UserNCID) & ", " & SQLQuote(TemplateID) & ", 0)"
        End If
    Else
        objSQL.Execute "delete from template_favorites where" _
            & " user_ncid = " & SQLQuote(cwobjLogonEnviron.UserNCID) & " and template_id = " & SQLQuote(msTemplateID)
    End If
    
    'SCR-50692 RBELL 3/3/2004 Changed condition based on moving to a base folder
    If Not meOwnerType = PersonalOwnerType Then
        'If template is not a personal template then an additional template_favorites entry needs to be made
        'SCR-44249 RBELL 12/23/03 Refined where caluse on DELETE to limit to the owner of template
        objSQL.Execute "delete from template_favorites where template_id = " & SQLQuote(msTemplateID) & " AND user_ncid = " & SQLQuote(msOwnerNCID)
        'SCR-47564 RBELL 6/3/04 Changed favorite informaiton based on owner
        If Not msOwnerNCID = msFavoriteFolderID Then
            objSQL.Execute "insert into template_favorites(user_ncid, template_id, folder_id) values (" _
                    & SQLQuote(msOwnerNCID) & ", " & SQLQuote(TemplateID) & ", " & CStr(val(msFavoriteFolderID)) & ")"
        Else
            objSQL.Execute "insert into template_favorites(user_ncid, template_id, folder_id) values (" _
                    & SQLQuote(msOwnerNCID) & ", " & SQLQuote(TemplateID) & ", 0)"
        End If
    End If
    
    mbDirty = False
    GoTo CleanUpAndExit
    
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "ListTemplate.GEMS_Save()", "TemplateClient", vbExclamation)
    'Resume Next

CleanUpAndExit:
    Set moConn = Nothing
    Set objSQL = Nothing
    
    Exit Function
    Resume
End Function

'
Public Function SaveAs() As Boolean
    
    On Error GoTo ErrHandler
    frmSaveListTemplate.Template = Me
    frmSaveListTemplate.SaveAs = True
    cwobjCWComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_TEMPLATE_MANAGEMENT
    frmSaveListTemplate.Show vbModal
    cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_TEMPLATE_MANAGEMENT
    If frmSaveListTemplate.SaveSelected Then SaveAs = Save
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.SaveAs", Err.Description
    
End Function
Public Function SaveDefaultTemplateAs() As Boolean
    
    On Error GoTo ErrHandler
    frmSaveListTemplate.SaveAs = True
    cwobjCWComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_TEMPLATE_MANAGEMENT
    Do:
       'SCR-48022 RBELL 1/14/04 Copied message before save dialog this message will redisplay after dialog closes if the user doesn't pick a personal type template
       MsgBxARMd "Please save this template as a personal template to set it as your default template.", vbOKOnly + vbInformation, "Save Default Template"
       frmSaveListTemplate.Template = Me
        frmSaveListTemplate.Show vbModal
        If frmSaveListTemplate.Cancelled Then Exit Do
        If Not meOwnerType = PersonalOwnerType Then
           MsgBxARMd "Please save this template as a personal template to set it as your default template.", vbOKOnly + vbInformation, "Save Default Template"
        ElseIf frmSaveListTemplate.SaveSelected Then
           SaveDefaultTemplateAs = Save
           Exit Do
        End If
    Loop
    cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_TEMPLATE_MANAGEMENT
    
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.SaveDefaultTemplateAs", Err.Description

End Function

Private Sub Class_Initialize()
    On Error Resume Next
    If gobjShared Is Nothing Then Set gobjShared = New CWShared
    Set ComplaintList = New Complaints
    Set AppointmentList = New AppointmentTypes
    Set ProblemList = New Problems
    Set AutociteList = New AutoCites
    Set NotesList = New NoteTemplates
    Set ClipartList = New ClipArts
    Set DiagnosisList = New Diagnoses
    Set OrderList = New Orders
    Set ProcedureList = New Procedures
    Set Instructions = New Instructions
    IsActive = True
End Sub

Private Sub Class_Terminate()
    
    On Error Resume Next
    Set ComplaintList = Nothing
    Set AppointmentList = Nothing
    Set ProblemList = Nothing
    Set AutociteList = Nothing
    Set NotesList = Nothing
    Set ClipartList = Nothing
    Set DiagnosisList = Nothing
    Set OrderList = Nothing
    Set ProcedureList = Nothing
    Set Instructions = Nothing
    If Not mobjEncTmplt Is Nothing Then mobjEncTmplt.ClearAll
    Set mobjEncTmplt = Nothing
    If Not mobjEncTmplt_GEMS Is Nothing Then mobjEncTmplt_GEMS.ClearAll
    Set mobjEncTmplt_GEMS = Nothing
    Set SpecialtyDictionary = Nothing
    Set EMCategoryDictionary = Nothing
    Set gobjSnoConvert = Nothing
    
End Sub


Public Property Get IsFavorite() As Boolean
    On Error GoTo ErrHandler
    IsFavorite = mbIsFavorite
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.IsFavorite", Err.Description

End Property

Public Property Let IsFavorite(ByVal RHV As Boolean)
    On Error GoTo ErrHandler
    mbIsFavorite = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.IsFavorite", Err.Description

End Property

Public Property Get IsShared() As Boolean
    On Error GoTo ErrHandler
    IsShared = mbIsShared
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.IsShared", Err.Description

End Property

Public Property Let IsShared(ByVal RHV As Boolean)
    
    On Error GoTo ErrHandler
    mbIsShared = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.IsShared", Err.Description

End Property

Public Property Get IsActive() As Boolean
    On Error GoTo ErrHandler
    IsActive = mbIsActive
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.IsActive", Err.Description

End Property

Public Property Let IsActive(ByVal RHV As Boolean)
    On Error GoTo ErrHandler
    mbIsActive = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.IsActive", Err.Description

End Property

Public Property Get OwnerType() As TemplateOwnerTypeEnum
    
    On Error GoTo ErrHandler
    OwnerType = meOwnerType
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.OwnerType", Err.Description

End Property

Public Property Let OwnerType(ByVal RHV As TemplateOwnerTypeEnum)
    
    On Error GoTo ErrHandler
    meOwnerType = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.OwnerType", Err.Description

End Property

Public Property Get TemplateID() As String
    
    On Error GoTo ErrHandler
    TemplateID = msTemplateID
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.TemplateID", Err.Description

End Property

Public Property Let TemplateID(ByVal RHV As String)
    
    On Error GoTo ErrHandler
    msTemplateID = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.TemplateID", Err.Description

End Property

Public Property Get SpecialtyID() As String
    
    On Error GoTo ErrHandler
    SpecialtyID = msSpecialtyID
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.SpecialtyID", Err.Description

End Property

Public Property Let SpecialtyID(ByVal RHV As String)
    
    On Error GoTo ErrHandler
    msSpecialtyID = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.SpecialtyID", Err.Description

End Property

Public Property Get EMCategoryID() As String
    On Error GoTo ErrHandler
    EMCategoryID = msEMCategoryID
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.EMCategoryID", Err.Description

End Property

Public Property Let EMCategoryID(ByVal RHV As String)
    On Error GoTo ErrHandler
    msEMCategoryID = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.EMCategoryID", Err.Description

End Property

Public Property Get ComplaintList() As TemplateData
    On Error GoTo ErrHandler
    Set ComplaintList = moComplaintList
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.ComplaintList", Err.Description

End Property

Public Property Set ComplaintList(ByVal RHV As TemplateData)
    On Error GoTo ErrHandler
    Set moComplaintList = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.ComplaintList", Err.Description

End Property

Public Property Get AppointmentList() As TemplateData
    
    On Error GoTo ErrHandler
    Set AppointmentList = moAppointmentList
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.AppointmentList", Err.Description
    
End Property

Public Property Set AppointmentList(ByVal RHV As TemplateData)
    
    On Error GoTo ErrHandler
    Set moAppointmentList = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.AppointmentList", Err.Description

End Property

Public Property Get ProblemList() As TemplateData
    
    On Error GoTo ErrHandler
    Set ProblemList = moProblemList
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.ProblemList", Err.Description

End Property

Public Property Set ProblemList(ByVal RHV As TemplateData)
    
    On Error GoTo ErrHandler
    Set moProblemList = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.ProblemList", Err.Description

End Property

Public Property Get AutociteList() As TemplateData
    On Error GoTo ErrHandler
    Set AutociteList = moAutociteList
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.AutociteList", Err.Description

End Property
Public Property Set AutociteList(ByVal RHV As TemplateData)
    On Error GoTo ErrHandler
    Set moAutociteList = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.AutociteList", Err.Description

End Property

Public Property Get NotesList() As TemplateData
    
    On Error GoTo ErrHandler
    Set NotesList = moNotesList
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.NotesList", Err.Description

End Property

Public Property Set NotesList(ByVal RHV As TemplateData)
    
    On Error GoTo ErrHandler
    Set moNotesList = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.NotesList", Err.Description

End Property

Public Property Get ClipartList() As TemplateData
    On Error GoTo ErrHandler
    Set ClipartList = moClipartList
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.ClipartList", Err.Description

End Property

Public Property Set ClipartList(ByVal RHV As TemplateData)
    On Error GoTo ErrHandler
    Set moClipartList = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.ClipartList", Err.Description

End Property

Public Property Get DiagnosisList() As TemplateData
    On Error GoTo ErrHandler
    Set DiagnosisList = moDiagnosisList
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.DiagnosisList", Err.Description

End Property

Public Property Set DiagnosisList(ByVal RHV As TemplateData)
    On Error GoTo ErrHandler
    Set moDiagnosisList = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.DiagnosisList", Err.Description

End Property

Public Property Get OrderList() As TemplateData
    
    On Error GoTo ErrHandler
    Set OrderList = moOrderList
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.OrderList", Err.Description

End Property

Public Property Set OrderList(ByVal RHV As TemplateData)
    
    On Error GoTo ErrHandler
    Set moOrderList = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.OrderList", Err.Description

End Property

Public Property Get ProcedureList() As TemplateData
    
    On Error GoTo ErrHandler
    Set ProcedureList = moProcedureList
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.ProcedureList", Err.Description

End Property

Public Property Set ProcedureList(ByVal RHV As TemplateData)
    
    On Error GoTo ErrHandler
    Set moProcedureList = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.ProcedureList", Err.Description

End Property

Public Property Get Instructions() As TemplateData
    On Error GoTo ErrHandler
    Set Instructions = moInstructions
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.Instructions", Err.Description

End Property

Public Property Set Instructions(ByVal RHV As TemplateData)
    On Error GoTo ErrHandler
    Set moInstructions = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.Instructions", Err.Description

End Property

Public Property Get CPG_ID() As String
    On Error GoTo ErrHandler
    CPG_ID = msCPG_ID
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.CPG_ID", Err.Description

End Property

Public Property Let CPG_ID(ByVal RHV As String)
    On Error GoTo ErrHandler
    msCPG_ID = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.CPG_ID", Err.Description

End Property

Public Property Get TemplateFacilityNCID() As String
    
    On Error GoTo ErrHandler
    TemplateFacilityNCID = msTemplateFacilityNCID
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.TemplateFacilityNCID", Err.Description

End Property

Public Property Let TemplateFacilityNCID(ByVal RHV As String)
    
    On Error GoTo ErrHandler
    msTemplateFacilityNCID = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.TemplateFacilityNCID", Err.Description

End Property



Public Property Get CreateUser() As String
    'SCR-47048 RBELL 12/17/03  ADDED
    On Error GoTo ErrHandler
    CreateUser = msCreateUser
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.CreateUser", Err.Description

End Property

Public Property Let CreateUser(ByVal RHV As String)
    'SCR-47048 RBELL 12/17/03  ADDED
    On Error GoTo ErrHandler
    msCreateUser = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.CreateUser", Err.Description

End Property

Public Property Get ModifyUser() As String
    'SCR-47048 RBELL 12/17/03  ADDED
    On Error GoTo ErrHandler
    ModifyUser = msModifyUser
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.ModifyUser", Err.Description

End Property

Public Property Let ModifyUser(ByVal RHV As String)
    'SCR-47048 RBELL 12/17/03  ADDED
    On Error GoTo ErrHandler
    msModifyUser = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.ModifyUser", Err.Description

End Property

Public Property Get StandaloneOrderSet() As Boolean
    
    'scr-44211, scr-44212 RBELL 2/11/2004 Returns boolean indicating if the template object represents a standalone order set
    On Error GoTo ErrHandler
    StandaloneOrderSet = IsStandaloneOrderSet(msTemplateID)
    Exit Property

ErrHandler:
    If Err = 5 Then
       StandaloneOrderSet = False
       Exit Property
    End If
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.StandaloneOrderSet", Err.Description

End Property

Public Property Get LinkedTemplates() As Collection
    On Error GoTo ErrHandler
    If mcLinkedTemplates Is Nothing Then Set mcLinkedTemplates = New Collection
    Set LinkedTemplates = mcLinkedTemplates
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.LinkedTemplates", Err.Description

End Property

Public Property Let LinkedTemplates(ByVal RHV As Collection)
    On Error GoTo ErrHandler
    Set mcLinkedTemplates = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.Template.LinkedTemplates", Err.Description

End Property
