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
Attribute VB_Name = "ListTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private oSpecialtyDictionary As Scripting.Dictionary

Private Const msSerialVersion As String = "V3"

Private mobjEncTmplt As ENCTMPLTATLLib.EncounterTemplate
'*-- Coleman 060601: GEMS - Adding Theater support
Private mobjEncTmplt_GEMS As GEMS_EncTmplt.EncounterTemplate
Private msName As String
Private msOwnerNCID As String
Private msOwnerName As String

Private mbIsFavorite As Boolean
Private mbIsShared As Boolean
Private mbIsActive As Boolean
Private meOwnerType As TemplateOwnerTypeEnum
Private msTemplateID As String
Private msSpecialtyID As String
Private mcListNoteItems As Collection ' of ListItem objects

Public Property Get SpecialtyDictionary() As Scripting.Dictionary
    On Error GoTo ErrHandler
    Set SpecialtyDictionary = oSpecialtyDictionary
    Exit Property
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.ListTemplate_Specialty", Err.Description
End Property

Public Property Set SpecialtyDictionary(ByVal RHV As Scripting.Dictionary)
    On Error GoTo ErrHandler
    Set oSpecialtyDictionary = RHV
    Exit Property
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.ListTemplate_Specialty", Err.Description
End Property

Public Property Get TemplateName() As String
    On Error GoTo ErrHandler
    TemplateName = Mid(msName, 2)
    Exit Property
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.ListTemplate_TemplateName", Err.Description
End Property

Public Property Let TemplateName(ByRef rsName As String)
    On Error GoTo ErrHandler
    msName = ListTemplatePrefix & rsName
    Exit Property
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.ListTemplate_TemplateName", Err.Description
End Property

Public Property Get TemplateType() As TemplateTypeEnum
    On Error GoTo ErrHandler
    TemplateType = ListTemplateType
    Exit Property
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.ListTemplate_TemplateType", Err.Description
End Property

Public Property Get OwnerNCID() As String
    On Error GoTo ErrHandler
    OwnerNCID = msOwnerNCID
    Exit Property
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.ListTemplate_OwnerNCID", Err.Description
End Property

Public Property Let OwnerNCID(NCID As String)
    On Error GoTo ErrHandler
    msOwnerNCID = NCID
    msOwnerName = ""
    Exit Property
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.ListTemplate_OwnerNCID", Err.Description
End Property

Public Property Get OwnerName() As String
    
    Dim objConcept As CONCEPTATLLib.ConceptCtrl
    Dim objClinicOps As CHCSII_ClinicClient.ClinicOpsEx
    Dim objClinic As DTOs.Clinic
    
    On Error GoTo ErrHandler
    If GemsDB Then '<SCR 37645 Alberstein
        OwnerName = GEMS_OwnerName
    Else    ' GEMS >>
        If msOwnerNCID <> vbNullString And msOwnerName = vbNullString Then
            If OwnerType = ClinicOwnerType Then
                On Error GoTo Cleanup
                Set objClinicOps = New CHCSII_ClinicClient.ClinicOpsEx
                Set objClinic = objClinicOps.GetClinic(msOwnerNCID)
                msOwnerName = objClinic.Name & " " & objClinic.MEPRSCode
            ElseIf Not gobjShared.IsAppMode(modeCHCSI_GUI) Then   '<SCR 36493
                Set objConcept = New CONCEPTATLLib.ConceptCtrl
                objConcept.UniqueId = msOwnerNCID
                msOwnerName = objConcept.PrefRep("2000").Representation
            End If
        End If
        OwnerName = msOwnerName
    End If
    GoTo Cleanup
    
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.ListTemplate_OwnerName", Err.Description
    
Cleanup:
    Set objClinicOps = Nothing
    Set objClinic = Nothing
    Set objConcept = Nothing

End Property
'<< GEMS        '*-- Coleman 072401:
Public Property Get GEMS_OwnerName() As String
    
    Dim objConcept As GEMSConceptCtrl.GEMS_ConceptCtrl
    Dim objClinicOps As CHCSII_ClinicClient.ClinicOpsEx
    Dim objClinic As DTOs.Clinic
    
    On Error GoTo ErrHandler
    If msOwnerNCID <> "" And msOwnerName = "" Then
        If OwnerType = ClinicOwnerType Then
            On Error GoTo Cleanup
            Set objClinicOps = New CHCSII_ClinicClient.ClinicOpsEx
            Set objClinic = objClinicOps.GetClinic(msOwnerNCID)
            msOwnerName = objClinic.Name & " " & objClinic.MEPRSCode
        Else
            Set objConcept = New GEMSConceptCtrl.GEMS_ConceptCtrl
            objConcept.UniqueId = msOwnerNCID
            msOwnerName = objConcept.PrefRep("2000").Representation
        End If
    End If
    GEMS_OwnerName = msOwnerName
    GoTo Cleanup
    
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.ListTemplate_GEMS_OwnerName", Err.Description
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
    Else
        Err.Raise Err.Number, "CHCSII_TemplateClient.ListTemplate_SpecialtyName", Err.Description
    End If
End Property

Public Function LoadFromEncounter(ByRef robjEncounter As CHCSIIEncounterOps.Encounter) As Boolean
    
    On Error GoTo ErrHandler
    TemplateName = "Created from Encounter " & robjEncounter.EncounterID
    MsgBxARMd "Loading a List Note Template from an encounter is not yet implemented."
    LoadFromEncounter = True
    Exit Function
    
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "ListTemplate.LoadFromEncounter()", "TemplateClient", vbExclamation)

End Function

Public Function Merge(rsTemplateID As String) As Boolean
    
    Dim objEncTmplt As ENCTMPLTATLLib.EncounterTemplate
    
    On Error GoTo ErrHandler
    If GemsDB Then '<SCR 37645 Alberstein
        Merge = GEMS_Merge(rsTemplateID)
    Else  ' GEMS >>
        Set objEncTmplt = New ENCTMPLTATLLib.EncounterTemplate
        objEncTmplt.TemplateID = rsTemplateID
        If Not val(rsTemplateID) = 0 Then objEncTmplt.Read
        DeserializeChildren objEncTmplt
        Merge = True
    End If
    GoTo Cleanup
    
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "ListTemplate.Merge()", "TemplateClient", vbExclamation)
Cleanup:
    Set objEncTmplt = Nothing

End Function
Public Function GEMS_Merge(rsTemplateID As String) As Boolean
    
    Dim objEncTmplt As GEMS_EncTmplt.EncounterTemplate
    
    On Error GoTo ErrHandler
    Set objEncTmplt = New GEMS_EncTmplt.EncounterTemplate
    If rsTemplateID = vbNullString Then Exit Function
    
    objEncTmplt.TemplateID = rsTemplateID
    objEncTmplt.Read
    GEMS_DeserializeChildren objEncTmplt
    GEMS_Merge = True
    GoTo Cleanup
    
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "ListTemplate.GEMS_Merge", "TemplateClient", vbExclamation)
Cleanup:
    Set objEncTmplt = Nothing

End Function

Public Function Delete() As Boolean
    
    Dim mobjSQL As ICHCSII_SQL.ISqlOps
    Dim moConn As CHCSII_CONN.Conn
        
    On Error GoTo ErrHandler
    Set moConn = New CHCSII_CONN.Conn
    Set mobjSQL = moConn.CHCSII_SQLOPS(Auto)
    
    If GemsDB Then '<SCR 37645 Alberstein
      If mobjEncTmplt_GEMS Is Nothing Then
        MsgBxARMd "Error: attempt to delete List Note Template not properly initialized."
      ElseIf CheckUserPrivilege(gcwobjUser, msOwnerNCID, cwobjLogonEnviron.UserNCID, OwnerType, Priv_Delete, False) Then
        mobjEncTmplt_GEMS.Delete
        mobjSQL.Execute "delete from template_favorites where template_id = " & TemplateID
        Delete = True
      End If
    Else
      If mobjEncTmplt Is Nothing Then
        MsgBxARMd "Error: attempt to delete List Note Template not properly initialized."
      ElseIf CheckUserPrivilege(gcwobjUser, msOwnerNCID, cwobjLogonEnviron.UserNCID, OwnerType, Priv_Delete, False) Then
        mobjEncTmplt.Delete
        mobjSQL.Execute "delete from template_favorites where template_id = " & TemplateID
        Delete = True
      End If
    End If
    GoTo CleanUpAndExit
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "ListTemplate.Delete", "TemplateClient", vbExclamation)
CleanUpAndExit:
    Set mobjSQL = Nothing
End Function

Public Function Load(ByRef rsTemplateID As String) As Boolean
    
    On Error GoTo ErrHandler
    If gobjShared Is Nothing Then Set gobjShared = New CWShared
    
    If GemsDB Then '<SCR 37645 Alberstein
        Load = GEMS_Load(rsTemplateID)
    Else    ' GEMS >>
        Set mobjEncTmplt = New ENCTMPLTATLLib.EncounterTemplate
        mobjEncTmplt.TemplateID = rsTemplateID
        If rsTemplateID = vbNullString Then Exit Function
        mobjEncTmplt.Read
        TemplateID = mobjEncTmplt.TemplateID
        OwnerNCID = mobjEncTmplt.owner
        IsActive = mobjEncTmplt.IsActive
        Select Case mobjEncTmplt.Type
        Case tt_facility
            If mobjEncTmplt.IsShared Then
                OwnerType = ClinicOwnerType
            Else
                OwnerType = MTFOwnerType
            End If
            IsShared = False
        Case tt_enterprise
            IsShared = False
            OwnerType = EnterpriseOwnerType
        Case Else
            IsShared = mobjEncTmplt.IsShared
            OwnerType = PersonalOwnerType
        End Select
        msName = mobjEncTmplt.Name
        DeserializeChildren mobjEncTmplt
        Load = True
    End If
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.ListTemplate.Load", Err.Description

End Function
'<< GEMS     '*-- Coleman 060601: Adding Theater support
Public Function GEMS_Load(ByRef rsTemplateID As String) As Boolean
    
    On Error GoTo ErrHandler
    Set mobjEncTmplt_GEMS = New GEMS_EncTmplt.EncounterTemplate
    If rsTemplateID = vbNullString Then Exit Function
    mobjEncTmplt_GEMS.TemplateID = rsTemplateID
    mobjEncTmplt_GEMS.Read
    TemplateID = mobjEncTmplt_GEMS.TemplateID
    OwnerNCID = mobjEncTmplt_GEMS.owner
    IsActive = mobjEncTmplt_GEMS.IsActive
    
    Select Case mobjEncTmplt_GEMS.eType
        Case tt_facility
            If mobjEncTmplt_GEMS.IsShared Then
                OwnerType = ClinicOwnerType
            Else
                OwnerType = MTFOwnerType
            End If
            IsShared = False
        Case tt_enterprise
            IsShared = False
            OwnerType = EnterpriseOwnerType
        Case Else
            IsShared = mobjEncTmplt_GEMS.IsShared
            OwnerType = PersonalOwnerType
    End Select
    
    msName = mobjEncTmplt_GEMS.Name
    GEMS_DeserializeChildren mobjEncTmplt_GEMS
    GEMS_Load = True
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.ListTemplate.GEMS_Load", Err.Description

End Function

Public Function LoadListToolRecords(mcToolRecords As Collection) As Boolean
    
    Dim objItem As ListNoteItem
    Dim oToolRecord As ListToolRecord
    
    On Error GoTo ErrHandler
    If mcToolRecords Is Nothing Then Exit Function
    If mcToolRecords.Count = 0 Then Exit Function
    
    If OwnerType = NoOwnerType Then
        OwnerType = PersonalOwnerType
        IsShared = True
    End If
    
    For Each oToolRecord In mcToolRecords
        Set objItem = New ListNoteItem
        objItem.SnoID = oToolRecord.SnoID
        objItem.Prefix = oToolRecord.Prefix
        objItem.Flag = oToolRecord.Flag
        objItem.Note = oToolRecord.Note
        objItem.Result = oToolRecord.Result
        ListNoteItems.Add objItem
    Next
    
    LoadListToolRecords = True
    GoTo Cleanup

ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "ListTemplate.LoadListToolRecords", "TemplateClient", vbExclamation)
Cleanup:
    Set objItem = Nothing
    Set oToolRecord = Nothing

End Function

Public Function Deserialize(robjSQL As ICHCSII_SQL.ISqlOps, Optional pMTFs As ADODB.Recordset = Nothing) As Boolean
    
    On Error GoTo ErrHandler
    If robjSQL Is Nothing Then Exit Function
    If robjSQL.BOF And robjSQL.EOF Then Exit Function
    
    Deserialize = Load(robjSQL("ID"))
    mbIsFavorite = GetUserFavoriteEntry(msTemplateID)
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.ListTemplate_Deserialize", Err.Description

End Function

Private Function DeserializeChildren(robjEncTmplt As ENCTMPLTATLLib.EncounterTemplate) As Boolean
    
    Dim objRef As ENCTMPLTATLLib.ETXref
    Dim objItem As ListNoteItem
    Dim colList As Collection
    Dim colItem As Collection
    Dim i As Integer

    On Error GoTo ErrHandler
    If robjEncTmplt Is Nothing Then Exit Function
    For Each objRef In robjEncTmplt.References
        If objRef.Type = SpecialtyNCID Then
            SpecialtyID = objRef.Value
            Exit For
        End If
    Next objRef
    
    Set objRef = Nothing
    For i = 1 To robjEncTmplt.Items.Count
        If robjEncTmplt.Items(i).Type = dt_blob Then
            Set colList = DeserializeString(robjEncTmplt.Items(i).dataBlob, gsMainDelim)
            Exit For
        End If
    Next i
    
    If Not colList Is Nothing Then
        If colList.Count > 0 Then
            If colList(1) = msSerialVersion Then
                For i = 2 To colList.Count
                    Set objItem = New ListNoteItem
                    If Not colList(i) = "" Then
                        Set colItem = DeserializeString(colList(i), gsRecordDelim)
                        If colItem.Count > 0 Then
                            objItem.SnoID = colItem(1)
                            objItem.Prefix = colItem(2)
                            objItem.Flag = colItem(3)
                            objItem.Note = colItem(4)
                            If colItem.Count >= 5 Then objItem.Result = colItem(5)
                            ListNoteItems.Add objItem
                        End If
                    End If
                Next i
            Else
                MsgBxARMd "Unexpected template serialization version '" & colList(1) & "'."
            End If
        End If
        Set colList = Nothing
        Set colItem = Nothing
        Set objItem = Nothing
    End If

    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.ListTemplate_DeserializeChildren", Err.Description

End Function
Private Function GEMS_DeserializeChildren(robjEncTmplt As Object) As Boolean
    
    Dim objRef As Object
    Dim objItem As ListNoteItem
    Dim colList As Collection
    Dim colItem As Collection
    Dim i As Integer
    
    On Error GoTo ErrHandler
    If robjEncTmplt Is Nothing Then Exit Function
    For Each objRef In robjEncTmplt.References
        If objRef.eType = SpecialtyNCID Then
            SpecialtyID = objRef.Value
            Exit For
        End If
    Next objRef
    Set objRef = Nothing
    For i = 1 To robjEncTmplt.Items.Count
        If robjEncTmplt.Items(CStr(i)).eType = dt_blob Then
            Set colList = DeserializeString(robjEncTmplt.Items(CStr(i)).dataBlob, gsMainDelim)
            Exit For
        End If
    Next i
    If Not colList Is Nothing Then
        If colList.Count > 0 Then
            If colList(1) = msSerialVersion Then
                For i = 2 To colList.Count
                    Set objItem = New ListNoteItem
                    Set colItem = DeserializeString(colList(i), gsRecordDelim)
                    objItem.SnoID = colItem(1)
                    objItem.Prefix = colItem(2)
                    objItem.Flag = colItem(3)
                    objItem.Note = colItem(4)
                    If colItem.Count >= 5 Then objItem.Result = colItem(5)
                    ListNoteItems.Add objItem
                Next i
            Else
                MsgBxARMd "Unexpected template serialization version '" & colList(1) & "'."
            End If
        End If
    End If
    GoTo CleanUpAndExit
    
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.ListTemplate.GEMS_DeserializeChildren", Err.Description
CleanUpAndExit:
    Set colList = Nothing
    Set colItem = Nothing
    Set objItem = Nothing

End Function

Public Function Export() As String
    
    Dim colList As Collection
    Dim colItem As Collection
    Dim objItem As ListNoteItem
    
    On Error GoTo ErrHandler
    Set colItem = New Collection
    Set colList = New Collection
    
    colItem.Add msSerialVersion
    colItem.Add msName
    colItem.Add OwnerType
    colItem.Add IsShared
    colItem.Add SpecialtyID
    colList.Add SerializeCollection(colItem, gsRecordDelim)
    For Each objItem In ListNoteItems
        Set colItem = New Collection
        colItem.Add objItem.SnoID
        colItem.Add objItem.Prefix
        colItem.Add objItem.Flag
        colItem.Add objItem.Note
        If gCopyForwardNote Then colItem.Add objItem.Result
        colList.Add SerializeCollection(colItem, gsRecordDelim)
    Next objItem
    Export = SerializeCollection(colList, gsMainDelim)
    GoTo Cleanup

ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "ListTemplate.Export()", "TemplateClient", vbExclamation)
Cleanup:
    Set colList = Nothing
    Set colItem = Nothing
    Set objItem = Nothing

End Function

Public Function Import(rsString As String) As Boolean
    
    Dim colList As Collection
    Dim colItem As Collection
    Dim objItem As ListNoteItem
    Dim i As Integer
    
    On Error GoTo ErrHandler
    If rsString = vbNullString Then Exit Function
    If Left(rsString, 2) = "L," Then rsString = Mid(rsString, 3)
    Set colList = DeserializeString(rsString, gsMainDelim)
    Set colItem = DeserializeString(colList(1), gsRecordDelim)
    If colItem(1) <> msSerialVersion Then
        MsgBxARMd "Imported string version ('" & colItem(1) & "') is not valid."
        GoTo Cleanup
    End If
    msName = colItem(2)
    If OwnerType = 0 Then
        OwnerType = colItem(3)
        IsShared = colItem(4)
    End If
    SpecialtyID = colItem(5)
    For i = 2 To colList.Count
        Set objItem = New ListNoteItem
        Set colItem = DeserializeString(colList(i), gsRecordDelim)
        objItem.SnoID = colItem(1)
        objItem.Prefix = colItem(2)
        objItem.Flag = colItem(3)
        objItem.Note = colItem(4)
        If colItem.Count >= 5 Then objItem.Result = colItem(5)
        ListNoteItems.Add objItem
    Next i
    Import = True
    GoTo Cleanup
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "ListTemplate.Import()", "TemplateClient", vbExclamation)
    'Resume Next
Cleanup:
    Set colList = Nothing
    Set colItem = Nothing
    Set objItem = Nothing
End Function

Public Function SaveInteractive(Optional ByVal SaveAs As Boolean = False, _
                                           Optional ByRef Parent As Object, _
                                           Optional ByVal ShowSharedCheckbox As Boolean = False, _
                                           Optional ByVal HideAddToButton As Boolean = False) As Boolean
    
    On Error GoTo ErrHandler
    frmSaveListTemplate.Template = Me
    frmSaveListTemplate.SaveAs = SaveAs
    frmSaveListTemplate.Cancelled = True
    frmSaveListTemplate.HideAddTo = HideAddToButton
    
    'COMMENTED-OUT SCR-15966 frmSaveListTemplate.chkShared.Visible = ShowSharedCheckbox
    cwobjCWComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_TEMPLATE_MANAGEMENT
    frmSaveListTemplate.Show vbModal, Parent
    cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_TEMPLATE_MANAGEMENT
    If Not frmSaveListTemplate.Cancelled Then
        SaveInteractive = Save()
    End If

    'SCR-49306 RBELL 2/9/4 Reset flag value to prevent buttons from displaying incorrectly
    frmSaveListTemplate.HideAddTo = False
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.ListTemplate.SaveInteractive", Err.Description
    
End Function

Public Function Save() As Boolean
    
    Dim colList As Collection
    Dim colItem As Collection
    Dim objItem As ListNoteItem
    Dim objRef As ENCTMPLTATLLib.ETXref
    Dim objSQL As ICHCSII_SQL.ISqlOps
    Dim moConn As CHCSII_CONN.Conn
    
    On Error GoTo ErrHandler
    
    If GemsDB Then '<SCR 37645 Alberstein
        Save = GEMS_Save
        Exit Function
    End If
    
    If Not CheckUserPrivilege(gcwobjUser, msOwnerNCID, cwobjLogonEnviron.UserNCID, OwnerType, Priv_Write, True) Then
        Exit Function
    End If
    If mobjEncTmplt Is Nothing Then
        Set mobjEncTmplt = New ENCTMPLTATLLib.EncounterTemplate
        If Not val(TemplateID) = 0 Then
            mobjEncTmplt.TemplateID = TemplateID
            mobjEncTmplt.Read
        Else
            mobjEncTmplt.ClearAll
        End If
    End If
    mobjEncTmplt.Items.ClearAll
    mobjEncTmplt.References.ClearAll
    With mobjEncTmplt
        If TemplateID = "" Then
            .TemplateID = 0
        Else
            .TemplateID = TemplateID
        End If
        .Name = msName
        .owner = msOwnerNCID
        .IsActive = IsActive
        Select Case OwnerType
        Case ClinicOwnerType
            .Type = tt_facility
            .IsShared = True
        Case MTFOwnerType
            .Type = tt_facility
            .IsShared = False
        Case EnterpriseOwnerType
            .Type = tt_enterprise
            .IsShared = False
        Case Else
            .Type = tt_user
            .IsShared = IsShared
        End Select
    End With
    ' specialty is a singleton cross reference
    If SpecialtyID <> "" Then
        Set objRef = mobjEncTmplt.References.Add
        objRef.Type = SpecialtyNCID
        objRef.Value = SpecialtyID
        Set objRef = Nothing
    End If
    Set colList = New Collection
    colList.Add msSerialVersion
    For Each objItem In ListNoteItems
        Set colItem = New Collection
        colItem.Add objItem.SnoID
        colItem.Add objItem.Prefix
        colItem.Add objItem.Flag
        colItem.Add objItem.Note
        If gCopyForwardNote Then colItem.Add objItem.Result
        colList.Add SerializeCollection(colItem, gsRecordDelim)
    Next objItem
    mobjEncTmplt.Items.Add
    mobjEncTmplt.Items(mobjEncTmplt.Items.Count).Type = dt_blob
    mobjEncTmplt.Items(mobjEncTmplt.Items.Count).dataBlob = SerializeCollection(colList, gsMainDelim)
    mobjEncTmplt.Save
    msTemplateID = CStr(mobjEncTmplt.TemplateID)
    Save = True
    On Error Resume Next
    Set moConn = New CHCSII_CONN.Conn
    Set objSQL = moConn.CHCSII_SQLOPS(Auto)
    If IsFavorite Then
        'Added folder_id to insert stmt SCR-17931
        objSQL.Execute "insert into template_favorites(user_ncid, template_id, folder_id) values (" _
            & SQLQuote(cwobjLogonEnviron.UserNCID) & ", " & mobjEncTmplt.TemplateID & ", 0)"
    Else
        objSQL.Execute "delete from template_favorites where" _
            & " user_ncid = " & SQLQuote(cwobjLogonEnviron.UserNCID) & " and template_id = " & mobjEncTmplt.TemplateID
    End If
    Err.Clear
    GoTo Cleanup
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "ListTemplate.Save()", "TemplateClient", vbExclamation)
    'Resume Next
Cleanup:
    Set colList = Nothing
    Set colItem = Nothing
    Set objItem = Nothing
    Set objRef = Nothing
    Set objSQL = Nothing
    Set moConn = Nothing

End Function
'<< GEMS
Public Function GEMS_Save() As Boolean
    
    Dim colList As Collection
    Dim colItem As Collection
    Dim objItem As ListNoteItem
    Dim objRef As GEMS_EncTmplt.ETXref
    Dim objSQL As ICHCSII_SQL.ISqlOps
    Dim moConn As CHCSII_CONN.Conn
        
    On Error GoTo ErrHandler
    If Not CheckUserPrivilege(gcwobjUser, msOwnerNCID, cwobjLogonEnviron.UserNCID, OwnerType, Priv_Write, True) Then
        Exit Function
    End If
    If mobjEncTmplt_GEMS Is Nothing Then
        Set mobjEncTmplt_GEMS = New GEMS_EncTmplt.EncounterTemplate
        If TemplateID <> "" Then
            mobjEncTmplt_GEMS.TemplateID = TemplateID
            'SCR 173161 - Do Not reload Template beofre saving it
            'mobjEncTmplt_GEMS.Read
        Else
            mobjEncTmplt_GEMS.ClearAll
        End If
    End If
    mobjEncTmplt_GEMS.Items.ClearAll
    mobjEncTmplt_GEMS.References.ClearAll
    With mobjEncTmplt_GEMS
        If TemplateID = "" Then
            .TemplateID = 0
        Else
            .TemplateID = TemplateID
        End If
        .Name = msName
        .owner = msOwnerNCID
        .IsActive = IsActive
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
    Set colList = New Collection
    colList.Add msSerialVersion
    For Each objItem In ListNoteItems
        Set colItem = New Collection
        colItem.Add objItem.SnoID
        colItem.Add objItem.Prefix
        colItem.Add objItem.Flag
        colItem.Add objItem.Note
        colList.Add SerializeCollection(colItem, gsRecordDelim)
    Next objItem
    mobjEncTmplt_GEMS.Items.Add
    mobjEncTmplt_GEMS.Items(mobjEncTmplt_GEMS.Items.Count).eType = dt_blob
    mobjEncTmplt_GEMS.Items(mobjEncTmplt_GEMS.Items.Count).dataBlob = SerializeCollection(colList, gsMainDelim)
    mobjEncTmplt_GEMS.Save
    GEMS_Save = True
    On Error Resume Next
    Set moConn = New CHCSII_CONN.Conn
    Set objSQL = moConn.CHCSII_SQLOPS(Auto)
    If objSQL Is Nothing Then
        Call gobjShared.ShowVBError(0, "Error: Failed to create SQL object.", _
                "ListTemplate.GEMS_Save", "TemplateClient Error", vbExclamation)
    ElseIf IsFavorite Then
        objSQL.Execute "insert into template_favorites(user_ncid, template_id, folder_id) values (" _
            & SQLQuote(cwobjLogonEnviron.UserNCID) & ", " & mobjEncTmplt_GEMS.TemplateID & ", 0)"
    Else
        objSQL.Execute "delete from template_favorites where" _
            & " user_ncid = " & SQLQuote(cwobjLogonEnviron.UserNCID) & " and template_id = " & mobjEncTmplt_GEMS.TemplateID
    End If
    Err.Clear
    GoTo Cleanup
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "ListTemplate.GEMS_Save()", "TemplateClient", vbExclamation)
Cleanup:
    Set colList = Nothing
    Set colItem = Nothing
    Set objItem = Nothing
    Set objRef = Nothing
    Set objSQL = Nothing
End Function
' GEMS >>
Private Sub Class_Initialize()
    On Error Resume Next
    Set ListNoteItems = New Collection
    IsActive = True
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Set ListNoteItems = Nothing
    If Not mobjEncTmplt Is Nothing Then
        mobjEncTmplt.ClearAll
        Set mobjEncTmplt = Nothing
    End If
    If Not mobjEncTmplt_GEMS Is Nothing Then
        mobjEncTmplt_GEMS.ClearAll
        Set mobjEncTmplt_GEMS = Nothing
    End If
    Set SpecialtyDictionary = Nothing
End Sub


Public Property Get IsFavorite() As Boolean
    On Error GoTo ErrHandler
    IsFavorite = mbIsFavorite
    Exit Property
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.ListTemplate.IsFavorite", Err.Description
End Property

Public Property Let IsFavorite(ByVal RHV As Boolean)
    On Error GoTo ErrHandler
    mbIsFavorite = RHV
    Exit Property
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.ListTemplate.IsFavorite", Err.Description
End Property

Public Property Get IsShared() As Boolean
    On Error GoTo ErrHandler
    IsShared = mbIsShared
    Exit Property
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.ListTemplate.IsShared", Err.Description
End Property

Public Property Let IsShared(ByVal RHV As Boolean)
    On Error GoTo ErrHandler
    mbIsShared = RHV
    Exit Property
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.ListTemplate.IsShared", Err.Description
End Property

Public Property Get IsActive() As Boolean
    On Error GoTo ErrHandler
    IsActive = mbIsActive
    Exit Property
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.ListTemplate.IsActive", Err.Description
End Property

Public Property Let IsActive(ByVal RHV As Boolean)
    On Error GoTo ErrHandler
    mbIsActive = RHV
    Exit Property
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.ListTemplate.IsActive", Err.Description
End Property

Public Property Get OwnerType() As TemplateOwnerTypeEnum
    On Error GoTo ErrHandler
    OwnerType = meOwnerType
    Exit Property
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.ListTemplate_OwnerType", Err.Description
End Property

Public Property Let OwnerType(ByVal RHV As TemplateOwnerTypeEnum)
    On Error GoTo ErrHandler
    meOwnerType = RHV
    Exit Property
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.ListTemplate_OwnerType", Err.Description
End Property

Public Property Get TemplateID() As String
    On Error GoTo ErrHandler
    TemplateID = msTemplateID
    Exit Property
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.ListTemplate_TemplateID", Err.Description
End Property

Public Property Let TemplateID(ByVal RHV As String)
    On Error GoTo ErrHandler
    msTemplateID = RHV
    Exit Property
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.ListTemplate_TemplateID", Err.Description
End Property

Public Property Get SpecialtyID() As String
    On Error GoTo ErrHandler
    SpecialtyID = msSpecialtyID
    Exit Property
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.ListTemplate_SpecialtyID", Err.Description
End Property

Public Property Let SpecialtyID(ByVal RHV As String)
    On Error GoTo ErrHandler
    msSpecialtyID = RHV
    Exit Property
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.ListTemplate_SpecialtyID", Err.Description
End Property

Public Property Get ListNoteItems() As Collection
    On Error GoTo ErrHandler
    Set ListNoteItems = mcListNoteItems
    Exit Property
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.ListTemplate.ListNoteItems", Err.Description
End Property

Public Property Set ListNoteItems(ByVal RHV As Collection)
    On Error GoTo ErrHandler
    Set mcListNoteItems = RHV
    Exit Property
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.ListTemplate.ListNoteItems", Err.Description
End Property
