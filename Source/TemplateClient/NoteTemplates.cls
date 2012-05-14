VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NoteTemplates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements TemplateData

Private Const MyLabel As String = "Notes Templates"

Private mListBox As VB.ListBox
Private mItems As Dictionary
Private mbDirty As Boolean
Private mobjConcept As CONCEPTATLLib.ConceptCtrl

Private mbIncludeResult As Boolean

' GEMS
Private mobjConcept_GEMS As GEMSConceptCtrl.GEMS_ConceptCtrl
Private mobjEncListTemplate As ListTemplate ' this is a virtual template created directly from an encounter

Public Property Get IncludeResult() As Boolean
    On Error GoTo ErrHandler
    IncludeResult = mbIncludeResult
    Exit Property
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.NoteTemplates.IncludeResult", Err.Description
End Property

Public Property Let IncludeResult(ByVal RHV As Boolean)
    On Error GoTo ErrHandler
    mbIncludeResult = RHV
    Exit Property
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.NoteTemplates.IncludeResult", Err.Description
End Property

Private Function GetIDFromName(pName As String) As Long
    
    Dim objSQL As ICHCSII_SQL.ISqlOps
    Dim moConn As CHCSII_CONN.Conn
    Dim sSQL As String
    
    On Error GoTo ErrHandler
    GetIDFromName = 0
    If gobjShared.IsAppMode(modeCHCSI_GUI) Then Exit Function '<SCR 36493
    Set moConn = New CHCSII_CONN.Conn
    Set objSQL = moConn.CHCSII_SQLOPS(Auto)
    
    sSQL = "SELECT ID FROM TEMPLATE_INFO WHERE NAME = " & SQLQuote(ListTemplatePrefix & pName)
    With objSQL
        .Execute sSQL
        If .BOF And .EOF Then GoTo Cleanup
        .MoveFirst
        GetIDFromName = .Value("ID")
    End With
    
    GoTo Cleanup
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "NoteTemplates.GetIDFromName", "TemplateClient", vbCritical)
    'Resume Next
Cleanup:
    Set objSQL = Nothing
    Set moConn = Nothing

End Function

Public Property Get ListTemplates() As Collection
    
    Dim objListTemplate As ListTemplate
    Dim sValue As Variant
    Dim i As Integer
    
    On Error GoTo ErrHandler
    Set ListTemplates = New Collection
    
    For i = 0 To mItems.Count - 1
        sValue = mItems.Keys(i)
        If sValue = ListTemplatePrefix Then
            ' this is the virtual template
            ListTemplates.Add mobjEncListTemplate
        ElseIf Left$(sValue, 1) = ListTemplatePrefix Then
            Set objListTemplate = New ListTemplate
            If objListTemplate.Load(Mid$(sValue, 2)) Then
                ListTemplates.Add objListTemplate
            End If
            Set objListTemplate = Nothing
        End If
    Next i
    
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.NoteTemplates.ListTemplates", Err.Description

End Property

Public Property Get FormTemplates() As Collection

'MedcinForms support added 7/17/2003 EMC
   Dim objFormTemplate As FormTemplate
   Dim sValue As Variant
   Dim i As Integer
   
    On Error GoTo ErrHandler
   Set FormTemplates = New Collection
   For i = 0 To mItems.Count - 1
      sValue = mItems.Keys(i)
      If sValue = FormTemplatePrefix Then
         'this is the virtual template
         FormTemplates.Add mobjEncListTemplate
      ElseIf Left(sValue, 1) = FormTemplatePrefix Then
         Set objFormTemplate = New FormTemplate
         If objFormTemplate.Load(Mid(sValue, 2)) Then
            FormTemplates.Add objFormTemplate
         End If
         Set objFormTemplate = Nothing
      End If
   Next i
    
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.NoteTemplates.FormTemplates", Err.Description

End Property

Public Property Get VirtualListTemplate() As ListTemplate
    On Error GoTo ErrHandler
    Set VirtualListTemplate = mobjEncListTemplate
    Exit Property
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.NoteTemplates.VirtualListTemplate", Err.Description
End Property

Public Sub AddItem(ByRef rsID As String, ByRef rsDesc As String)
    
    On Error GoTo ErrHandler
    mItems.Add rsID, rsDesc
    mbDirty = True
    mDisplayItems
    Exit Sub
    
ErrHandler:
    If Err = 457 Then
       mItems.Remove rsID
       Resume
    End If
    Err.Raise Err.Number, "CHCSII_TemplateClient.NoteTemplates.AddItem", Err.Description
    
End Sub

Private Sub TemplateData_DoMoveDown()
    
    Dim sNCIDFrom As String
    Dim sDescFrom As String
    Dim sNCIDTo As String
    Dim sDescTo As String
    Dim i As Integer
    
    On Error GoTo ErrHandler
    For i = mListBox.ListCount - 1 To 0 Step -1
        If mListBox.Selected(i) Then
            sNCIDFrom = mItems.Keys(i)
            sDescFrom = mListBox.List(i)
            sNCIDTo = mItems.Keys(i + 1)
            sDescTo = mListBox.List(i + 1)
            If sNCIDFrom = "0" Or sNCIDTo = "0" Then Exit Sub 'SCR-53144 5/4/2004 RBELL Fix data related error
            
            mItems.Key(sNCIDFrom) = "0"
            mItems.Key(sNCIDTo) = sNCIDFrom
            mItems(sNCIDFrom) = sDescFrom
            mItems.Key("0") = sNCIDTo
            mItems(sNCIDTo) = sDescTo
            mDisplayItems
            mListBox.Selected(i + 1) = True
            mbDirty = True
        End If
    Next i
    Exit Sub
    
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.NoteTemplates.TemplateData_DoMoveDown", Err.Description

End Sub

Private Sub TemplateData_DoMoveUp()
    
    Dim sNCIDFrom As String
    Dim sDescFrom As String
    Dim sNCIDTo As String
    Dim sDescTo As String
    Dim i As Integer
    
    On Error GoTo ErrHandler
    For i = mListBox.ListCount - 1 To 0 Step -1
        If mListBox.Selected(i) Then
            sNCIDFrom = mItems.Keys(i)
            sDescFrom = mListBox.List(i)
            sNCIDTo = mItems.Keys(i - 1)
            sDescTo = mListBox.List(i - 1)
            If sNCIDFrom = "0" Or sNCIDTo = "0" Then Exit Sub 'SCR-53144 5/4/2004 RBELL Fix data related error
            
            mItems.Key(sNCIDFrom) = "0"
            mItems.Key(sNCIDTo) = sNCIDFrom
            mItems(sNCIDFrom) = sDescFrom
            mItems.Key("0") = sNCIDTo
            mItems(sNCIDTo) = sDescTo
            mDisplayItems
            mListBox.Selected(i - 1) = True
            mbDirty = True
            Exit For
        End If
    Next i
    Exit Sub
    
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.NoteTemplates.TemplateData_DoMoveUp", Err.Description

End Sub


Private Function TemplateData_Export() As String
    
    Dim objVirtTemplate As ListTemplate
    Dim colItems As Collection
    Dim i As Integer
    Dim sValue As Variant
    
    On Error GoTo ErrHandler
    Set objVirtTemplate = New ListTemplate
    Set colItems = New Collection
    
    For i = 0 To mItems.Count - 1
        sValue = mItems.Keys(i)
        If Trim$(mItems.Items(i)) <> "[List]" Then 'Ignore unusable templates 'SCR-7847
            If sValue = ListTemplatePrefix Then ' this is the virtual template
                Call objVirtTemplate.Import(mobjEncListTemplate.Export())
                colItems.Add ListTemplatePrefix & gsRecordDelim & objVirtTemplate.Export()
            ElseIf Left$(sValue, 1) = ListTemplatePrefix Then
                Set objVirtTemplate = New ListTemplate
                Call objVirtTemplate.Load(Mid$(sValue, 2))
                colItems.Add objVirtTemplate.Export()
            Else
                colItems.Add sValue
            End If
        End If
    Next i
    
    Set objVirtTemplate = Nothing
    TemplateData_Export = SerializeCollection(colItems, gsListDelim)
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.NoteTemplates.TemplateData_Export", Err.Description

End Function

Private Function TemplateData_Import(rsString As String) As Boolean
    'SCR-34348 4/8/03 RBELL DISABLED IMPORT OF THIS ITEM FOR PGUI MODES
    On Error GoTo ErrHandler
    If gobjShared Is Nothing Then Set gobjShared = New CWShared
    If gobjShared.IsAppMode(modeCHCSI_GUI) Then Exit Function '<SCR 36493
    
    TemplateData_Import = TemplateData_DeSerialize(rsString)
    Exit Function

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.NoteTemplates.TemplateData_Serialize", Err.Description

End Function

Private Function TemplateData_GetFromEncounter(objEncounter As CHCSIIEncounterOps.Encounter) As Boolean
    
    Dim objListRecord As CHCSIIEncounterOps.ListToolRecord
    Dim objCheckRecord As ListNoteItem
    
    Dim objTemplateOps As CHCSII_TemplateClient.TemplateOps
    Dim oListTemplate As ListTemplate
    Dim colRecords As Collection
    Dim colCopyForwardRecords As Collection
    Dim lRecord As Long
    Dim lCheckRecord As Long
    
    Dim objItem As ListNoteItem
    Dim oMedcin As Object
    
    On Error GoTo ErrHandler
    Set colRecords = objEncounter.AllListNoteRecords()
    If colRecords.Count > 0 Then
        'SCR-52758 RBELL 4/22/2004 Moved 2 object set statements up in code ni case the template is a copy forward
        Set objTemplateOps = New CHCSII_TemplateClient.TemplateOps
        Set oListTemplate = objTemplateOps.NewTemplate(ListTemplateType)
        If gCopyForwardNote Then
           Set colCopyForwardRecords = New Collection
           For lRecord = colRecords.Count To 1 Step -1
                Set objListRecord = colRecords(lRecord)
                Set objItem = New ListNoteItem
                objItem.SnoID = objListRecord.SnoID
                objItem.Flag = objListRecord.Flag
                objItem.Prefix = objListRecord.Prefix
                objItem.Note = objListRecord.Note
                objItem.Result = objListRecord.Result
                For lCheckRecord = 1 To colCopyForwardRecords.Count
                    Set objCheckRecord = colCopyForwardRecords(lCheckRecord)
                    If objCheckRecord.SnoID = CStr(objListRecord.SnoID) Then
                       objItem.Note = objListRecord.Note & " " & objCheckRecord.Note
                       If Not objCheckRecord.Result = objListRecord.Result Then objItem.Result = ""
                       colCopyForwardRecords.Remove lCheckRecord
                       Exit For
                    End If
                Next
                colCopyForwardRecords.Add objItem
           Next
           Set oListTemplate.ListNoteItems = colCopyForwardRecords
           Set mobjEncListTemplate.ListNoteItems = colCopyForwardRecords
           mobjEncListTemplate.TemplateName = "Copy Forward " & objEncounter.EncounterID
           oListTemplate.TemplateName = "Copy Forward " & objEncounter.EncounterID
        Else
             For Each objListRecord In colRecords
                 Set objItem = New ListNoteItem
                 objItem.SnoID = objListRecord.SnoID
                 objItem.Flag = objListRecord.Flag
                 objItem.Prefix = objListRecord.Prefix
                 objItem.Note = objListRecord.Note
                 objItem.Result = objListRecord.Result
                 oListTemplate.ListNoteItems.Add objItem
                 mobjEncListTemplate.ListNoteItems.Add objItem
             Next objListRecord
             If oListTemplate.SaveInteractive(True, Nothing, True, True) Then
                mobjEncListTemplate.TemplateName = oListTemplate.TemplateName
             Else
                mobjEncListTemplate.TemplateName = "Created from Encounter " & objEncounter.EncounterID
             End If
        End If
        Set oMedcin = cwobjCWComm.InitializeOLEServer(NCID_MEDCIN_APPLICATION)
        oMedcin.RefreshTemplates
    End If
    
    If mobjEncListTemplate.ListNoteItems.Count > 0 Then
        mItems.Add ListTemplatePrefix, "[List] " & mobjEncListTemplate.TemplateName
    End If
    GoTo Cleanup
    
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.NoteTemplates.TemplateData_GetFromEncounter", Err.Description
    'Resume Next

Cleanup:
    Set oMedcin = Nothing
    Set colRecords = Nothing
    Set objItem = Nothing
    Set objListRecord = Nothing
    Set objCheckRecord = Nothing
    Set colCopyForwardRecords = Nothing
    
End Function

Private Function TemplateData_GetQueryCriteria() As String

    On Error GoTo ErrHandler
    Exit Function

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.NoteTemplates.TemplateData_GetQueryCrieteria", Err.Description
    
End Function

Private Sub mDisplayItems()
    
    Dim i As Integer
    
    On Error GoTo ErrHandler
    mListBox.Clear
    For i = 0 To mItems.Count - 1
        mListBox.AddItem mItems.Items(i)
        mListBox.ItemData(mListBox.NewIndex) = i
    Next i
    Exit Sub
    
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.AutoCites.mDisplayItems", Err.Description
    
End Sub

Private Sub Class_Initialize()
    On Error Resume Next
    Set mItems = New Dictionary
    
    If gobjShared Is Nothing Then Set gobjShared = New CWShared

    Set mobjConcept_GEMS = New GEMSConceptCtrl.GEMS_ConceptCtrl
    
    Set mobjEncListTemplate = New ListTemplate
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Set mItems = Nothing
    Set mobjEncListTemplate = Nothing
    Set mobjConcept = Nothing
    Set mListBox = Nothing
End Sub

Private Property Get TemplateData_CanEdit() As Boolean
    On Error GoTo ErrHandler
    TemplateData_CanEdit = False
    Exit Property
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.NoteTemplates.CanEdit", Err.Description
End Property

Private Function TemplateData_DeSerialize(rvData As Variant) As Boolean
'Medcin Forms support added 7/17/2003 EMC
    Dim colItems As Collection
    Dim sValue As Variant
    Dim colFields As Collection
    Dim objItem As ListNoteItem
    Dim lTemplateID As Long
    
    Dim objSQL As ICHCSII_SQL.ISqlOps
    Dim moConn As CHCSII_CONN.Conn
        
    Dim strPrefix As String
    Dim strDesc As String
              
    On Error GoTo ErrHandler
    Set colItems = DeserializeString(CStr(rvData), gsListDelim)
    Set mobjEncListTemplate = New ListTemplate 'SCR-51810 RBELL 4/8/2004 Move object set outside for loop so it only occurs once
    Set mItems = New Dictionary

    For Each sValue In colItems
         'if Medcin forms supported, determine if template is form or list template
         If Not cwobjCWComm.CmdLineSwitch("MEDCINFORMS") = "" Then
            'Is this a form or list template
            If Left(sValue, 1) = FormTemplatePrefix Then
               strPrefix = FormTemplatePrefix
               strDesc = "[Form] "
            Else 'If Left(sValue, 1) = ListTemplatePrefix Then
               'SCR-51338 5/6/2004 RBELL Default to List templates if identifier is not a form template
               strPrefix = ListTemplatePrefix
               strDesc = "[List] "
            End If
         Else 'otherwise, assume it's a list (once switch goes away can always perform check above
            strPrefix = ListTemplatePrefix
            strDesc = "[List] "
         End If
         
         ' this is a medcin template
         If Mid$(sValue, 2, 1) = gsRecordDelim Or Left(sValue, 2) = "V3" Then
             'this is a virtual medcin template created directly from an encounter
             Set colFields = DeserializeString(CStr(sValue), gsRecordDelim)
             If Left(sValue, 2) = "V3" Then
                mobjEncListTemplate.Import CStr(sValue)
             Else
                mobjEncListTemplate.Import Mid$(sValue, 3)
             End If
             
             If (Not mobjEncListTemplate.TemplateName = "From Export/Import" And Not mobjEncListTemplate.TemplateName = vbNullString) Then
               lTemplateID = GetIDFromName(mobjEncListTemplate.TemplateName)
               If Not lTemplateID = 0 Then
                  If InStr(1, mobjEncListTemplate.TemplateName, "Copy Forward", vbTextCompare) = 0 Then
                     If Not mobjEncListTemplate.TemplateName = "" Then mItems.Add strPrefix & CStr(lTemplateID), strDesc & mobjEncListTemplate.TemplateName
                  Else
                     If Not mobjEncListTemplate.TemplateName = "" Then mItems.Add strPrefix, strDesc & mobjEncListTemplate.TemplateName
                  End If
               Else
                  mobjEncListTemplate.OwnerNCID = cwobjLogonEnviron.UserNCID
                  '*****************************************************************************************************************
                  'SCR-39547 COMMENTED-OUT CREATING COPY FORWARD TEMPLATES UNITENTIONALLY AND NOT NEEDED
                  'If Not gobjShared.IsAppMode(modeCHCSI_GUI) Then mobjEncListTemplate.Save  '<SCR 36493
                  '*****************************************************************************************************************
                  'SCR-52416 RBELL 4/13/2004 VB error 13 on empty template id value
                  lTemplateID = val(mobjEncListTemplate.TemplateID)
                  If Not mobjEncListTemplate.TemplateName = "" Then mItems.Add strPrefix, strDesc & mobjEncListTemplate.TemplateName
               End If
             ElseIf mobjEncListTemplate.TemplateName = vbNullString Then
                'Used by template mgmt to refresh the note section of the template data
                mItems.Add strPrefix, vbNullString
             End If
             On Error Resume Next
         Else
             ' this is a regular medcin template--we have a template id
             If Not sValue = vbNullString Then
                Set moConn = New CHCSII_CONN.Conn
                Set objSQL = moConn.CHCSII_SQLOPS(Auto)
                objSQL.Execute "select name from template_info where id = " & Mid$(sValue, 2)
                On Error Resume Next
                If Not Trim(objSQL("NAME")) = "" Then mItems.Add CStr(sValue), strDesc & Mid$(objSQL("NAME"), 2)
                On Error GoTo ErrHandler
             End If
         End If
    Next sValue
    If Not mListBox Is Nothing Then
        Call mDisplayItems
    End If
    mbDirty = False
    TemplateData_DeSerialize = True
    GoTo Cleanup
ErrHandler:
    If Err = 457 Then
       mItems.Remove strPrefix & CStr(lTemplateID)
       Resume
    End If
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "NoteTemplates.Deserialize", "TemplateClient", vbCritical)
    'Resume Next
Cleanup:
    Set objSQL = Nothing
    Set objItem = Nothing
    Set colItems = Nothing
    Set colFields = Nothing
    Set moConn = Nothing

End Function

Private Sub TemplateData_DoAdd()
    
    Dim objTemplateOps As CHCSII_TemplateClient.TemplateOps
        
    On Error GoTo ErrHandler
    ' frmNoteTemplates needs its own templateops object to do list template searches
    Set objTemplateOps = New CHCSII_TemplateClient.TemplateOps
    cwobjCWComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_TEMPLATE_MANAGEMENT
    frmNoteTemplates.DoAdd Me, objTemplateOps
    cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_TEMPLATE_MANAGEMENT
    Set objTemplateOps = Nothing
    Exit Sub
    
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.NoteTemplates.TemplateData_DoAdd", Err.Description

End Sub

Private Sub TemplateData_DoEdit()

    On Error GoTo ErrHandler
    Exit Sub
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.NoteTemplates.TemplateData_DoEdit", Err.Description

End Sub

Private Sub TemplateData_DoRemove()
    
    Dim i As Integer
    Dim j As Integer
    Dim nRemove As Integer
    
    On Error GoTo ErrHandler
    For i = mListBox.ListCount - 1 To 0 Step -1
        If mListBox.Selected(i) Then
            nRemove = mListBox.ItemData(i)
            If mItems.Keys(nRemove) = ListTemplatePrefix Then
                Set mobjEncListTemplate = New ListTemplate
                'mobjEncListTemplate.Initialize cwobjlogonenviron, gcwobjUser, gobjShared, mobjConcept, cwobjPatient, cwobjcwComm, mlProgId
            End If
            mItems.Remove mItems.Keys(nRemove)
            For j = 0 To mListBox.ListCount - 1
                If mListBox.ItemData(j) > nRemove Then
                    mListBox.ItemData(j) = mListBox.ItemData(j) - 1
                End If
            Next j
            mListBox.RemoveItem i
            mbDirty = True
        End If
    Next i
    Exit Sub
    
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.NoteTemplates.TemplateData_DoRemove", Err.Description

End Sub

Private Property Let TemplateData_isDirty(RHS As Boolean)
    On Error GoTo ErrHandler
    mbDirty = RHS
    Exit Property
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.NoteTemplates.IsDirty", Err.Description
End Property

Private Property Get TemplateData_isDirty() As Boolean
    On Error GoTo ErrHandler
    TemplateData_isDirty = mbDirty
    Exit Property
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.NoteTemplates.IsDirty", Err.Description
End Property

Private Property Get TemplateData_Items() As Scripting.IDictionary
    On Error GoTo ErrHandler
    Set TemplateData_Items = mItems
    Exit Property
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.NoteTemplates.Items", Err.Description
End Property

Private Property Get TemplateData_Label() As String
    On Error GoTo ErrHandler
    TemplateData_Label = MyLabel
    Exit Property
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.NoteTemplates.Label", Err.Description
End Property

Private Property Set TemplateData_ListBox(RHS As Object)
    
    On Error GoTo ErrHandler
    Set mListBox = RHS
    Call mDisplayItems
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.NoteTemplates.ListBox", Err.Description

End Property

Private Property Get TemplateData_ListBox() As Object

    On Error GoTo ErrHandler
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.NoteTemplates.ListBox", Err.Description
    
End Property

Private Function TemplateData_Serialize(rvData As Variant) As Boolean
    
    Dim colItems As Collection
    Dim i As Integer
    Dim sValue As Variant
    
    On Error GoTo ErrHandler
    Set colItems = New Collection
    For i = 0 To mItems.Count - 1
        sValue = mItems.Keys(i)
        If sValue = ListTemplatePrefix Then
            colItems.Add ListTemplatePrefix & gsRecordDelim & mobjEncListTemplate.Export() ' this is the virtual template
        ElseIf sValue = vbNullString Then
            colItems.Add ListTemplatePrefix & gsRecordDelim & mobjEncListTemplate.Export() ' this is the virtual template used by template mgmt
        Else
            colItems.Add sValue
        End If
    Next i
    rvData.Add SerializeCollection(colItems, gsListDelim)
    mbDirty = False
    Exit Function
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.NoteTemplates.TemplateData_Serialize", Err.Description

End Function
