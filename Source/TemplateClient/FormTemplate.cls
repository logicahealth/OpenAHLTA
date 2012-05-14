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
Attribute VB_Name = "FormTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private oSpecialtyDictionary As Scripting.Dictionary

Private Const msSerialVersion As String = "V1.0"

Private mobjEncTmplt As ENCTMPLTATLLib.EncounterTemplate
'*-- Coleman 060601: GEMS - Adding Theater support
Private mobjEncTmplt_GEMS As GEMS_EncTmplt.EncounterTemplate
Private mnVersion As Integer
Private msName As String
Private msOwnerNCID As String
Private msOwnerName As String

Private mbIsFavorite As Boolean
Private mbIsShared As Boolean
Private mbIsActive As Boolean
Private meOwnerType As TemplateOwnerTypeEnum
Private msTemplateID As String
Private msSpecialtyID As String
Private mcFormItems As Collection ' of FormItem objects  'emc Forms support
Private msFavoriteFolderID As String
Public Property Get FavoriteFolderID() As String
    
    On Error GoTo ErrHandler
    FavoriteFolderID = msFavoriteFolderID
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.FormTemplate.FavoriteFolderID", Err.Description

End Property
Public Property Let FavoriteFolderID(ByVal RHV As String)
    On Error GoTo ErrHandler
    msFavoriteFolderID = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.FormTemplate.FavoriteFolderID", Err.Description

End Property

Public Property Get SpecialtyDictionary() As Scripting.Dictionary
    Set SpecialtyDictionary = oSpecialtyDictionary
End Property

Public Property Set SpecialtyDictionary(ByVal RHV As Scripting.Dictionary)
    Set oSpecialtyDictionary = RHV
End Property

Public Property Get TemplateName() As String
    TemplateName = Mid(msName, 2)
End Property

Public Property Let TemplateName(ByRef rsName As String)
    msName = FormTemplatePrefix & rsName
End Property

Public Property Get TemplateType() As TemplateTypeEnum
    TemplateType = FormTemplateType
End Property

Public Property Get OwnerNCID() As String
    OwnerNCID = msOwnerNCID
End Property

Public Property Let OwnerNCID(NCID As String)
    msOwnerNCID = NCID
    msOwnerName = ""
End Property

Public Property Get OwnerName() As String
    Dim objConcept As CONCEPTATLLib.ConceptCtrl
    Dim objClinicOps As CHCSII_ClinicClient.ClinicOpsEx
    Dim objClinic As DTOs.Clinic
    
    '<< GEMS        '*-- Coleman 072401:
    If GemsDB Then '<SCR 37645 Alberstein
        OwnerName = GEMS_OwnerName
    Else    ' GEMS >>
        If msOwnerNCID <> "" And msOwnerName = "" Then
            If OwnerType = ClinicOwnerType Then
                On Error GoTo Cleanup
                Set objClinicOps = New CHCSII_ClinicClient.ClinicOpsEx
                Set objClinic = objClinicOps.GetClinic(msOwnerNCID)
                msOwnerName = objClinic.Name & " " & objClinic.MEPRSCode
            ElseIf gobjShared.IsAppMode(modeCDR) And Not gobjShared.IsAppMode(modeITT) Then
                Set objConcept = New CONCEPTATLLib.ConceptCtrl
                objConcept.UniqueId = msOwnerNCID
                msOwnerName = objConcept.PrefRep("2000").Representation
            End If
        End If
        OwnerName = msOwnerName
    End If
    
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
Cleanup:
    Set objClinicOps = Nothing
    Set objClinic = Nothing
    Set objConcept = Nothing
End Property
' GEMS >>

Public Property Get SpecialtyName() As String
    On Error Resume Next
    SpecialtyName = SpecialtyDictionary(SpecialtyID)
End Property

Public Function LoadFromEncounter(ByRef robjEncounter As CHCSIIEncounterOps.Encounter) As Boolean
    TemplateName = "Created from Encounter " & robjEncounter.EncounterID
    MsgBxARMd "Loading a Form Template from an encounter is not yet implemented."
    LoadFromEncounter = True
    GoTo Cleanup
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "FormTemplate.LoadFromEncounter()", "TemplateClient", vbExclamation)
Cleanup:
End Function

Public Function Merge(rsTemplateID As String) As Boolean
    On Error GoTo ErrHandler
    Dim objEncTmplt As ENCTMPLTATLLib.EncounterTemplate
    
    'emc Forms Support
    MsgBxARMd "Merging a Form Template is not yet implemented."
    Merge = True
    
    GoTo Cleanup
    
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "FormTemplate.Merge()", "TemplateClient", vbExclamation)
Cleanup:
    Set objEncTmplt = Nothing
End Function
'<< GEMS    '*-- Coleman 060601: Adding Theater support
Public Function GEMS_Merge(rsTemplateID As String) As Boolean
'    On Error GoTo ErrHandler
'    Dim objEncTmplt As GEMS_EncTmplt.EncounterTemplate
'    Set objEncTmplt = New GEMS_EncTmplt.EncounterTemplate
'
'    objEncTmplt.TemplateID = rsTemplateID
'    objEncTmplt.Read
'    GEMS_DeserializeChildren objEncTmplt
'    GEMS_Merge = True
'    GoTo Cleanup
'
'ErrHandler:
'    Call gobjShared.ShowVBError(Err.Number, Err.Description, "FormTemplate.GEMS_Merge()", "TemplateClient", vbExclamation)
'Cleanup:
'    Set objEncTmplt = Nothing
End Function
' GEMS >>

Public Function Delete() As Boolean
'<DEAD
    Dim mobjSQL As ICHCSII_SQL.ISqlOps
    Dim moConn As CHCSII_CONN.Conn
        
    On Error GoTo ErrHandler
    Set moConn = New CHCSII_CONN.Conn
    Set mobjSQL = moConn.CHCSII_SQLOPS(Auto)
    
    If GemsDB Then '<SCR 37645 Alberstein
      If mobjEncTmplt_GEMS Is Nothing Then
        MsgBxARMd "Error: attempt to delete Form Template not properly initialized."
      ElseIf CheckUserPrivilege(gcwobjUser, msOwnerNCID, cwobjLogonEnviron.UserNCID, OwnerType, Priv_Delete, False) Then
        mobjEncTmplt_GEMS.Delete
        mobjSQL.Execute "delete from template_favorites where template_id = " & TemplateID
        mobjSQL.Execute "delete from template_form_data where template_id = " & TemplateID
        Delete = True
      End If
    Else
      If mobjEncTmplt Is Nothing Then
        MsgBxARMd "Error: attempt to delete Form Template not properly initialized."
      ElseIf CheckUserPrivilege(gcwobjUser, msOwnerNCID, cwobjLogonEnviron.UserNCID, OwnerType, Priv_Delete, False) Then
        mobjEncTmplt.Delete
        mobjSQL.Execute "delete from template_favorites where template_id = " & TemplateID
        mobjSQL.Execute "delete from template_form_data where template_id = " & TemplateID
        Delete = True
      End If
    End If
    GoTo CleanUpAndExit
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "FormTemplate.Delete()", "TemplateClient", vbExclamation)
CleanUpAndExit:
    Set mobjSQL = Nothing
End Function

Public Function Load(ByRef rsTemplateID As String) As Boolean
    '<< GEMS     '*-- Coleman 060601: Adding Theater support
    'SCR 18334; S.McAvoy 10/31/01
    If gobjShared Is Nothing Then
        Set gobjShared = New CWShared
    End If
    
    
    If GemsDB Then '<SCR 37645 Alberstein
        Load = GEMS_Load(rsTemplateID)
    Else    ' GEMS >>
        Set mobjEncTmplt = New ENCTMPLTATLLib.EncounterTemplate
        mobjEncTmplt.TemplateID = rsTemplateID
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
End Function
'<< GEMS     '*-- Coleman 060601: Adding Theater support
Public Function GEMS_Load(ByRef rsTemplateID As String) As Boolean
    Set mobjEncTmplt_GEMS = New GEMS_EncTmplt.EncounterTemplate
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
End Function
' GEMS >>

Public Function LoadListToolRecords(mcToolRecords As Collection) As Boolean
    'emc Forms support
    MsgBxARMd "LoadListToolRecords not supported for Forms templates"
    
'    Dim objItem As ListNoteItem
'    Dim oToolRecord As ListToolRecord
'
'    On Error GoTo ErrHandler
'    If mcToolRecords Is Nothing Then Exit Function
'    If mcToolRecords.Count = 0 Then Exit Function
'
'    If OwnerType = NoOwnerType Then
'        OwnerType = PersonalOwnerType
'        IsShared = True
'    End If
'
'
'    For Each oToolRecord In mcToolRecords
'        Set objItem = New ListNoteItem
'        objItem.SnoID = oToolRecord.SnoID
'        objItem.Prefix = oToolRecord.Prefix
'        objItem.Flag = oToolRecord.Flag
'        objItem.Note = oToolRecord.Note
'        objItem.Result = oToolRecord.Result
'        ListNoteItems.Add objItem
'    Next
'
'    LoadListToolRecords = True
'    GoTo Cleanup
'
'ErrHandler:
'    Call gobjShared.ShowVBError(Err.Number, Err.Description, "FormTemplate.LoadListToolRecords", "TemplateClient", vbExclamation)
'Cleanup:
'    Set objItem = Nothing
'    Set oToolRecord = Nothing
'
End Function

Public Function Deserialize(robjSQL As ICHCSII_SQL.ISqlOps, Optional pMTFs As ADODB.Recordset = Nothing) As Boolean
    
    On Error GoTo ErrHandler
    Deserialize = Load(robjSQL("ID"))
    Select Case UCase(TypeName(robjSQL("FAVORITES_USER")))
        Case "STRING"
            If CStr(robjSQL("FAVORITES_USER")) & vbNullString = cwobjLogonEnviron.UserNCID Then
                IsFavorite = True
            Else
                IsFavorite = False
            End If
        Case "DECIMAL", "INTEGER", "LONG"
            If val(robjSQL("FAVORITES_USER")) = val(cwobjLogonEnviron.UserNCID) Then
                IsFavorite = True
            Else
                IsFavorite = False
            End If
            
        Case Else
            IsFavorite = False
    End Select
    
    msFavoriteFolderID = robjSQL("FOLDER_ID") & vbNullString
    Exit Function
    
ErrHandler:
    If Err = FieldNotInQueryResults Then Resume Next
    Err.Raise Err.Number, "CHCSII_TemplateClient.FormTemplate.Deserialize", Err.Description
    
End Function

Private Function DeserializeChildren(robjEncTmplt As ENCTMPLTATLLib.EncounterTemplate) As Boolean
    
    Dim objRef As ENCTMPLTATLLib.ETXref
    Dim objItem As FormItem
    Dim colForm As Collection
    Dim colItem As Collection
    Dim i As Integer

    On Error GoTo ErrHandler
    For Each objRef In robjEncTmplt.References
        If objRef.Type = SpecialtyNCID Then
            SpecialtyID = objRef.Value
            Exit For
        End If
    Next objRef
    Set objRef = Nothing
    Set colForm = DeserializeString(GetFormDataBlob, gsMainDelim)
    
    If Not colForm Is Nothing Then
        If colForm.Count > 0 Then
            If colForm(1) = msSerialVersion Then
                For i = 2 To colForm.Count
                    Set objItem = New FormItem
                    Set colItem = DeserializeString(colForm(i), gsRecordDelim)
                    objItem.Page = colItem(1)
                    objItem.Left = colItem(2)
                    objItem.Top = colItem(3)
                    objItem.Right = colItem(4)
                    objItem.Bottom = colItem(5)
                    objItem.SnoID = colItem(6)
                    objItem.Flags = colItem(7)
                    objItem.Prefix = colItem(8)
                    objItem.ItemData = colItem(9)
                    objItem.Description = colItem(10)
                    FormItems.Add objItem
                Next i
            Else
                MsgBxARMd "Unexpected template serialization version '" & colForm(1) & "'."
            End If
        End If
        Set colForm = Nothing
        Set colItem = Nothing
        Set objItem = Nothing
    End If
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.FormTemplate.DeserializeChildren", Err.Description

End Function
Private Function GEMS_DeserializeChildren(robjEncTmplt As Object) As Boolean
    
    Dim objRef As Object
    Dim objItem As FormItem
    Dim colForm As Collection
    Dim colItem As Collection
    Dim i As Integer
    For Each objRef In robjEncTmplt.References
        If objRef.eType = SpecialtyNCID Then
            SpecialtyID = objRef.Value
            Exit For
        End If
    Next objRef
    
    'R. Foster SCR #48802;   Developer: Robert Foster 01/29/2004 11:08 AM
    'R. Foster Note: Function Changed: GEMS_Save, GEMS_DeserializeChildren      Impact of Change:  Enable the Template CLient to work in ITT mode.
    'ITT CONVERGENCE
    Set objRef = Nothing
    Set colForm = DeserializeString(GetFormDataBlob, gsMainDelim)
    If Not colForm Is Nothing Then
        If colForm.Count > 0 Then
            If colForm(1) = msSerialVersion Then
                For i = 2 To colForm.Count
                    Set objItem = New FormItem
                    Set colItem = DeserializeString(colForm(i), gsRecordDelim)
                    objItem.Page = colItem(1)
                    objItem.Left = colItem(2)
                    objItem.Top = colItem(3)
                    objItem.Right = colItem(4)
                    objItem.Bottom = colItem(5)
                    objItem.SnoID = colItem(6)
                    objItem.Flags = colItem(7)
                    objItem.Prefix = colItem(8)
                    objItem.ItemData = colItem(9)
                    objItem.Description = colItem(10)
                    FormItems.Add objItem
                    Debug.Print objItem.ItemData
                Next i
            Else
                MsgBxARMd "Unexpected template serialization version '" & colForm(1) & "'."
            End If
        End If
        Set colForm = Nothing
        Set colItem = Nothing
        Set objItem = Nothing
    End If

End Function

Public Function Export() As String
    'Dim colList As Collection
    Dim colForm As Collection
    Dim colItem As Collection
    Dim objItem As FormItem
    Dim strTemp As String
    
    On Error GoTo ErrHandler
    Set colItem = New Collection
    Set colForm = New Collection
    colItem.Add """MedcinForm-" & msSerialVersion & """"
    strTemp = SerializeCollection(colItem, gsRecordDelim)
    colForm.Add Left(strTemp, Len(strTemp) - 1) 'Drop trailing comma
    Set colItem = New Collection
    
    colItem.Add """" & Me.TemplateName & """"
    colItem.Add """CHCSII"""
    colItem.Add """System"""
    strTemp = SerializeCollection(colItem, gsRecordDelim)
    colForm.Add Left(strTemp, Len(strTemp) - 1) 'Drop trailing comma
    Set colItem = New Collection
    
    For Each objItem In FormItems
        Set colItem = New Collection
        
        colItem.Add objItem.Page
        colItem.Add objItem.Left
        colItem.Add objItem.Top
        colItem.Add objItem.Right
        colItem.Add objItem.Bottom
        colItem.Add objItem.SnoID
        colItem.Add objItem.Flags
        'SCR 173081 removed additional quotes
        colItem.Add objItem.Prefix
        colItem.Add objItem.ItemData
        colItem.Add objItem.Description
        
        strTemp = SerializeCollection(colItem, gsRecordDelim)
        colForm.Add Left(strTemp, Len(strTemp) - 1)
    Next objItem
    strTemp = SerializeCollection(colForm, vbCrLf)
    Export = Left(strTemp, Len(strTemp) - 1)
    GoTo Cleanup
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "FormTemplate.Export()", "TemplateClient", vbExclamation)
Cleanup:
    Set colForm = Nothing
    Set colItem = Nothing
    Set objItem = Nothing
End Function

Public Function Import(rsString As String) As Boolean
    'Dim colList As Collection
    Dim colForm As Collection
    Dim colItem As Collection
    Dim objItem As FormItem
    Dim i As Integer
    On Error GoTo ErrHandler
    If rsString = "" Then
        Exit Function
    End If
    
    If Split(rsString, vbCrLf)(0) = """MedcinForm-V1.0""" Then
      'Import = ImportMedicompForm(rsString)
      Import = ImportFormFileContents(rsString)
    Else
      Set colForm = DeserializeString(rsString, gsMainDelim)
      Set colItem = DeserializeString(colForm(1), gsRecordDelim)
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
      For i = 2 To colForm.Count
          Set objItem = New FormItem
          Set colItem = DeserializeString(colForm(i), gsRecordDelim)
           objItem.Page = colItem(1)
           objItem.Left = colItem(2)
           objItem.Top = colItem(3)
           objItem.Right = colItem(4)
           objItem.Bottom = colItem(5)
           objItem.SnoID = colItem(6)
           objItem.Flags = colItem(7)
           objItem.Prefix = colItem(8)
           objItem.ItemData = colItem(9)
           objItem.Description = colItem(10)
          'If colItem.Count >= 5 Then objItem.Result = colItem(5)
          FormItems.Add objItem
      Next i
      Import = True
    End If
    
    GoTo Cleanup
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "FormTemplate.Import()", "TemplateClient", vbExclamation)
Cleanup:
    Set colForm = Nothing
    Set colItem = Nothing
    Set objItem = Nothing
End Function

Private Function ImportMedicompForm(rsString As String) As Boolean
'Medicomp export format using old style VB file I/O.  Since we're only receiving
'the contents of the file and not the filename, need to "manually" parse the contents
'Comma delimited where 1st seven entries are numeric.  Last 3 are strings with double-quotes.
'Just need make sure we handle case where a string contains a comma

   Dim strItems() As String
   Dim strFields() As String
   Dim i As Integer
   Dim objItem As FormItem
   Dim j As Integer
   Dim k As Integer
   
   strItems = Split(rsString, vbCrLf)

   For k = 1 To UBound(strItems)
      strFields = Split(strItems(k), ",")
      
      If k = 1 Then
         'msName = strFields(0)
         
               For i = 0 To UBound(strFields)
                  msName = msName & strFields(i)
                  If Left(msName, 1) = """" And Right(msName, 1) = """" Then
                     'j = i + 1
                     msName = Mid(msName, 2, Len(msName) - 2)
                     msName = FormTemplatePrefix & msName
                     Exit For
                  End If
               Next i
               
               'Limit to 50 characters
               If Len(msName) > 50 Then
                msName = Left(msName, 50)
               End If
               
         OwnerType = PersonalOwnerType
         IsShared = False
      Else
         If Len(strItems(k)) <> 0 Then
            Set objItem = New FormItem
            If UBound(strFields) = 9 Then  'everything split out nicely
               With objItem
                  If IsNumeric(strFields(0)) Then .Page = strFields(0)
                  If IsNumeric(strFields(1)) Then .Left = strFields(1)
                  If IsNumeric(strFields(2)) Then .Top = strFields(2)
                  If IsNumeric(strFields(3)) Then .Right = strFields(3)
                  If IsNumeric(strFields(4)) Then .Bottom = strFields(4)
                  If IsNumeric(strFields(5)) Then .SnoID = strFields(5)
                  If IsNumeric(strFields(6)) Then .Flags = strFields(6)
                .Prefix = StripString(strFields(7))
                  .ItemData = StripString(Replace(strFields(8), "^t", vbTab))
                  .Description = StripString(strFields(9))
               End With
            ElseIf UBound(strFields) > 1 Then
               With objItem
                  'Numeric fields
                  If IsNumeric(strFields(0)) Then .Page = strFields(0)
                  If IsNumeric(strFields(1)) Then .Left = strFields(1)
                  If IsNumeric(strFields(2)) Then .Top = strFields(2)
                  If IsNumeric(strFields(3)) Then .Right = strFields(3)
                  If IsNumeric(strFields(4)) Then .Bottom = strFields(4)
                  If IsNumeric(strFields(5)) Then .SnoID = strFields(5)
                  If IsNumeric(strFields(6)) Then .Flags = strFields(6)
                  
                  'String Fields
                  For i = 7 To UBound(strFields)
                     .Prefix = .Prefix & strFields(i)
                     If Left(.Prefix, 1) = """" And Right(.Prefix, 1) = """" Then
                        .Prefix = StripString(.Prefix)
                        j = i + 1
                        Exit For
                     End If
                  Next i
                  
                  For i = j To UBound(strFields)
                     .ItemData = .ItemData & Replace(strFields(i), "^t", vbTab)
                     If Left(.ItemData, 1) = """" And Right(.ItemData, 1) = """" Then
                        .ItemData = StripString(.ItemData)
                        j = i + 1
                        Exit For
                     End If
                  Next i
                  
                  For i = j To UBound(strFields)
                     .Description = .Description & strFields(i)
                     If Left(.Description, 1) = """" And Right(.Description, 1) = """" Then
                        .Description = StripString(.Description)
                        j = i + 1
                        Exit For
                     End If
                  Next i
   
               End With
            End If

                FormItems.Add objItem

         End If

         
      End If
   Next k

   ImportMedicompForm = True
End Function
Private Function StripString(strData As String) As String
'Removes 1st and last character of a string
'with intent of removing leading and trailing double-quotes
   If Len(strData) = 2 Then
      StripString = ""
   Else
      StripString = Mid(strData, 2, Len(strData) - 2)
   End If
End Function

Private Function ReadLong(strInput As String) As Long
   
End Function


Public Function SaveInteractive(Optional ByVal SaveAs As Boolean = False, _
                                           Optional ByRef Parent As Object, _
                                           Optional ByVal ShowSharedCheckbox As Boolean = False, _
                                           Optional ByVal HideAddToButton As Boolean = False) As Boolean
    On Error Resume Next
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

End Function

Public Function Save() As Boolean
    
    Dim colForm As Collection
    Dim colItem As Collection
    Dim objItem As FormItem
    Dim objRef As ENCTMPLTATLLib.ETXref
    Dim objSQL As ICHCSII_SQL.ISqlOps
    Dim moConn As CHCSII_CONN.Conn
    Dim strBlob As String
    Dim bSaveFavorites As Boolean
    
    On Error GoTo ErrHandler
    bSaveFavorites = False
    
    ' << GEMS    '*-- Coleman 081001:
    If GemsDB Then '<SCR 37645 Alberstein
        Save = GEMS_Save
        Exit Function
    End If
    ' GEMS >>
    
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
    Set colForm = New Collection
    colForm.Add msSerialVersion
    For Each objItem In FormItems
        Set colItem = New Collection
        
        colItem.Add objItem.Page
        colItem.Add objItem.Left
        colItem.Add objItem.Top
        colItem.Add objItem.Right
        colItem.Add objItem.Bottom
        colItem.Add objItem.SnoID
        colItem.Add objItem.Flags
        colItem.Add objItem.Prefix
        colItem.Add objItem.ItemData
        colItem.Add objItem.Description
        
        'If gCopyForwardNote Then colItem.Add objItem.Result
        colForm.Add SerializeCollection(colItem, gsRecordDelim)
    Next objItem
    mobjEncTmplt.Items.Add
    mobjEncTmplt.Items(mobjEncTmplt.Items.Count).Type = dt_blob
    strBlob = SerializeCollection(colForm, gsMainDelim)
    'mobjEncTmplt.Items(mobjEncTmplt.Items.Count).dataBlob = SerializeCollection(colForm, gsMainDelim)
    mobjEncTmplt.Save
    msTemplateID = CStr(mobjEncTmplt.TemplateID)
    Call SaveFormDataBlob(strBlob) 'we need the templateID, so perform the "main" save 1st
    Save = True
'    On Error Resume Next
    Set moConn = New CHCSII_CONN.Conn
    Set objSQL = moConn.CHCSII_SQLOPS(Auto)
    
'    If IsFavorite Then
'        'Added folder_id to insert stmt SCR-17931
'        objSQL.Execute "insert into template_favorites(user_ncid, template_id, folder_id) values (" _
'            & SQLQuote(cwobjLogonEnviron.UserNCID) & ", " & mobjEncTmplt.TemplateID & ", 0)"
'    Else
'        objSQL.Execute "delete from template_favorites where" _
'            & " user_ncid = " & SQLQuote(cwobjLogonEnviron.UserNCID) & " and template_id = " & mobjEncTmplt.TemplateID
'    End If
    If IsFavorite Then
        bSaveFavorites = True
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
    
    If Not meOwnerType = PersonalOwnerType Then
        'If template is not a personal template then an additional template_favorites entry needs to be made
        objSQL.Execute "delete from template_favorites where template_id = " & SQLQuote(msTemplateID) & " AND user_ncid = " & SQLQuote(msOwnerNCID)
        objSQL.Execute "insert into template_favorites(user_ncid, template_id, folder_id) values (" _
                & SQLQuote(msOwnerNCID) & ", " & SQLQuote(TemplateID) & ", " & CStr(val(msFavoriteFolderID)) & ")"
    End If
    
    bSaveFavorites = False
    GoTo Cleanup

ErrHandler:
    If Err = FieldNotInQueryResults And bSaveFavorites Then Resume Next 'Insert of favorite violated key disregard as entry already exist for this user
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "FormTemplate.Save()", "TemplateClient", vbExclamation)
    'Resume Next
Cleanup:
    Set colForm = Nothing
    Set colItem = Nothing
    Set objItem = Nothing
    Set objRef = Nothing
    Set objSQL = Nothing
    Set moConn = Nothing

End Function
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
    Err.Raise Err.Number, "CHCSII_TemplateClient.FormTemplate.SaveAs", Err.Description
    
End Function

'<< GEMS
Public Function GEMS_Save() As Boolean
    
    Dim colForm As Collection
    Dim colItem As Collection
    Dim objItem As FormItem
    Dim objRef As GEMS_EncTmplt.ETXref
    Dim objSQL As ICHCSII_SQL.ISqlOps
    Dim moConn As CHCSII_CONN.Conn
    Dim bSaveFavorites As Boolean
    
    'R. Foster SCR #48802;   Developer: Robert Foster 01/29/2004 11:08 AM
    'R. Foster Note: Function Changed: GEMS_Save, GEMS_DeserializeChildren      Impact of Change:  Enable the Template CLient to work in ITT mode.
    'ITT CONVERGENCE
    Dim strBlob As String
        
    On Error GoTo ErrHandler
    bSaveFavorites = False
    
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
    Set colForm = New Collection
    colForm.Add msSerialVersion
    For Each objItem In FormItems
        Set colItem = New Collection
        
        colItem.Add objItem.Page
        colItem.Add objItem.Left
        colItem.Add objItem.Top
        colItem.Add objItem.Right
        colItem.Add objItem.Bottom
        colItem.Add objItem.SnoID
        colItem.Add objItem.Flags
        colItem.Add objItem.Prefix
        colItem.Add objItem.ItemData
        colItem.Add objItem.Description
        Debug.Print objItem.ItemData
        colForm.Add SerializeCollection(colItem, gsRecordDelim)
    Next objItem
    mobjEncTmplt_GEMS.Items.Add
    mobjEncTmplt_GEMS.Items(mobjEncTmplt_GEMS.Items.Count).eType = dt_blob
    mobjEncTmplt_GEMS.Items(mobjEncTmplt_GEMS.Items.Count).dataBlob = SerializeCollection(colForm, gsMainDelim)
    'R. Foster SCR #48802;   Developer: Robert Foster 01/29/2004 11:08 AM
    'R. Foster Note: Function Changed: GEMS_Save, GEMS_DeserializeChildren      Impact of Change:  Enable the Template CLient to work in ITT mode.
    'ITT CONVERGENCE
    strBlob = SerializeCollection(colForm, gsMainDelim)
    mobjEncTmplt_GEMS.Save
    'R. Foster SCR #48802;   Developer: Robert Foster 01/29/2004 11:08 AM
    'R. Foster Note: Function Changed: GEMS_Save, GEMS_DeserializeChildren      Impact of Change:  Enable the Template CLient to work in ITT mode.
    'ITT CONVERGENCE
    msTemplateID = CStr(mobjEncTmplt_GEMS.TemplateID)
    Call SaveFormDataBlob(strBlob) 'we need the templateID, so perform the "main" save 1st
    GEMS_Save = True
    On Error Resume Next
    Set moConn = New CHCSII_CONN.Conn
    Set objSQL = moConn.CHCSII_SQLOPS(Auto)
    If objSQL Is Nothing Then
        Call gobjShared.ShowVBError(0, "Error: Failed to create SQL object.", _
                "FormTemplate.GEMS_Save", "TemplateClient Error", vbExclamation)
        Exit Function
    End If
    
    If IsFavorite Then
        bSaveFavorites = True
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
    
    If Not meOwnerType = PersonalOwnerType Then
        'If template is not a personal template then an additional template_favorites entry needs to be made
        objSQL.Execute "delete from template_favorites where template_id = " & SQLQuote(msTemplateID) & " AND user_ncid = " & SQLQuote(msOwnerNCID)
        objSQL.Execute "insert into template_favorites(user_ncid, template_id, folder_id) values (" _
                & SQLQuote(msOwnerNCID) & ", " & SQLQuote(TemplateID) & ", " & CStr(val(msFavoriteFolderID)) & ")"
    End If
    
    bSaveFavorites = False
    GoTo Cleanup

ErrHandler:
    If Err = FieldNotInQueryResults And bSaveFavorites Then Resume Next 'Insert of favorite violated key disregard as entry already exist for this user
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "FormTemplate.GEMS_Save()", "TemplateClient", vbExclamation)
    
Cleanup:
    Set colForm = Nothing
    Set colItem = Nothing
    Set objItem = Nothing
    Set objRef = Nothing
    Set objSQL = Nothing
    
End Function
' GEMS >>
Private Sub Class_Initialize()
    Set FormItems = New Collection
    IsActive = True
End Sub

Private Sub Class_Terminate()
   'SCR 20618 Phil Crowder 1/17/02
   'Global reference not longer used.
   ' On Error Resume Next
   ' Unload frmMedcin
   ' If Err Then Err.Clear
    On Error GoTo 0
    Set FormItems = Nothing
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
    IsFavorite = mbIsFavorite
End Property

Public Property Let IsFavorite(ByVal RHV As Boolean)
    mbIsFavorite = RHV
End Property

Public Property Get IsShared() As Boolean
    IsShared = mbIsShared
End Property

Public Property Let IsShared(ByVal RHV As Boolean)
    mbIsShared = RHV
End Property

Public Property Get IsActive() As Boolean
    IsActive = mbIsActive
End Property

Public Property Let IsActive(ByVal RHV As Boolean)
    mbIsActive = RHV
End Property

Public Property Get OwnerType() As TemplateOwnerTypeEnum
    OwnerType = meOwnerType
End Property

Public Property Let OwnerType(ByVal RHV As TemplateOwnerTypeEnum)
    meOwnerType = RHV
End Property

Public Property Get TemplateID() As String
    TemplateID = msTemplateID
End Property

Public Property Let TemplateID(ByVal RHV As String)
    msTemplateID = RHV
End Property

Public Property Get SpecialtyID() As String
    SpecialtyID = msSpecialtyID
End Property

Public Property Let SpecialtyID(ByVal RHV As String)
    msSpecialtyID = RHV
End Property

Public Property Get FormItems() As Collection
    Set FormItems = mcFormItems
End Property

Public Property Set FormItems(ByVal RHV As Collection)
    Set mcFormItems = RHV
End Property


Public Function SaveFormDataBlob(strBlob As String) As Boolean

    Dim mobjSQL As ICHCSII_SQL.ISqlOps
    Dim moConn As CHCSII_CONN.Conn
    Dim strSQL As String
    Dim strChunks() As String
    Dim intChunks As Integer
    Dim i As Integer
    Dim strChunk As String
    
    On Error GoTo ErrHandler
    Set moConn = New CHCSII_CONN.Conn
    Set mobjSQL = moConn.CHCSII_SQLOPS(Auto)
    
    'Remove any existing entries
    If Len(Trim(TemplateID)) <> 0 Then
        strSQL = "Delete FROM Template_Form_Data where Template_Id = " & TemplateID
        mobjSQL.Execute strSQL
    End If
    
    'Determine how many chunks are needed
    intChunks = CInt(Len(strBlob) \ 4000) + 1 'Notice the backslash to only return the integer portion, no rounding
    
    ReDim strChunks(1 To intChunks)
    
    For i = 1 To intChunks

        strChunk = Mid(strBlob, CLng(4000) * (i - 1) + 1, 4000)  'Need CLng() to force Long calculations
        strChunk = Replace(strChunk, "'", "''") 'Handle single quotes in SQL

        strSQL = "Insert INTO Template_Form_Data(Template_Id, Seq, Form_Data) Values(" & TemplateID & ", " & i & ", '" & strChunk & "')"
        mobjSQL.Execute strSQL
    Next i
    
    GoTo CleanUpAndExit
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "FormTemplate.SaveFormDataBlob()", "TemplateClient", vbExclamation)
CleanUpAndExit:
    Set mobjSQL = Nothing
End Function

Private Function GetFormDataBlob() As String

    Dim mobjSQL As ICHCSII_SQL.ISqlOps
    Dim moConn As CHCSII_CONN.Conn
    Dim strSQL As String
    Dim strChunks() As String
    Dim intChunks As Integer
    Dim i As Integer
    
    On Error GoTo ErrHandler
    Set moConn = New CHCSII_CONN.Conn
    Set mobjSQL = moConn.CHCSII_SQLOPS(Auto)
    
    'Retrieve the entries
    strSQL = "select * from Template_Form_Data WHERE Template_ID = " & TemplateID & " ORDER BY SEQ"
    mobjSQL.Execute strSQL
    
    mobjSQL.MoveFirst
    If mobjSQL.RecordCount > 0 Then
        ReDim strChunks(1 To mobjSQL.RecordCount)
        
        For i = 1 To mobjSQL.RecordCount
            strChunks(i) = mobjSQL.Value("Form_Data")
            mobjSQL.MoveNext
        Next i
        GetFormDataBlob = Join(strChunks, "")
    End If
    
    GoTo CleanUpAndExit
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "FormTemplate.Delete()", "TemplateClient", vbExclamation)
CleanUpAndExit:
    Set mobjSQL = Nothing
End Function
Private Function ImportFormFileContents(rsString As String) As Boolean
'Medicomp export format using old style VB file I/O.  Since we're only receiving
'the contents of the file and not the filename, need to save the data off as a temp file
'and then access it using old school vb
    
    Dim fs As Scripting.FileSystemObject
    Dim fileName As String
    Dim ts As Scripting.TextStream
    
    Set fs = New Scripting.FileSystemObject
    fileName = fs.BuildPath(gobjShared.RTDataPath, fs.GetTempName)
    Set ts = fs.OpenTextFile(fileName, ForWriting, True)
    ts.Write rsString
    ts.Close
    
    ImportFormFileContents = ImportMedicompFormFile(fileName)
    
    fs.DeleteFile fileName
End Function

Private Function ImportMedicompFormFile(fileName As String) As Boolean
'Given a file name, read the file as a Medicomp Form file
    Dim fileNumber As Integer
    Dim signature As String
    Dim Name As String
    Dim group As String
    Dim owner As String
    Dim Page As Integer
    Dim leftPos As Integer
    Dim topPos As Integer
    Dim rightPos As Integer
    Dim bottomPos As Integer
    Dim SnoID As String
    Dim Flags As Long
    Dim Prefix As String
    Dim ItemData As String
    Dim Description As String
    Dim objItem As FormItem
    Dim sLineItem As String
    Dim sBuf() As String
    '--- Fix the problem with itemdata contains comma
    Dim sTemp As String
    Dim iPosition As Integer
    
    '--- Fix the problem with itemdata contains line feed
    Dim sFinalLineItem As String
    
    On Error GoTo ErrHandler
    
    fileNumber = FreeFile
    Open fileName For Input As #fileNumber
    
    Line Input #fileNumber, signature
    
    Line Input #fileNumber, sLineItem
    
    sBuf = Split(sLineItem, ",", , vbTextCompare)
    Name = Replace(sBuf(0), Chr(34), vbNullString)
    group = Replace(sBuf(1), Chr(34), vbNullString)
    owner = Replace(sBuf(2), Chr(34), vbNullString)
        
    msName = FormTemplatePrefix & Name
               
    If Len(msName) > 50 Then
        msName = Left$(msName, 50)
    End If
    
    OwnerType = PersonalOwnerType
    IsShared = False
    Do While Not EOF(fileNumber)
    
        'Use Line Input to read file to make sure all data is valid SCR #90031
        
        '--- Fix the problem with itemdata contains line feed
        sFinalLineItem = vbNullString
        Do While Right$(sFinalLineItem, 1) <> Chr(34)
            '--- if the last line is just tab key, ignore the line
            If LenB(Replace(sLineItem, Chr(9), vbNullString)) = 0 Then
                Exit Do
            Else
                Line Input #fileNumber, sLineItem
                sFinalLineItem = sFinalLineItem & Chr(10) & Chr(13) & sLineItem
           End If
        Loop
        'break only into 10 substrings (exactly what input command would have done)
        sBuf = Split(sFinalLineItem, ",", 10, vbTextCompare)
                
        If UBound(sBuf) = 9 Then
            
            If Right(sBuf(8), 1) <> Chr(34) And Left(sBuf(9), 1) <> Chr(34) Then
                sTemp = sBuf(8) & gsRecordDelim & sBuf(9)
                iPosition = InStr(sTemp, Chr(34) & gsRecordDelim)
                sBuf(8) = Left$(sTemp, iPosition)
                sBuf(9) = Mid$(sTemp, iPosition + 2)
            End If
            
            ItemData = ReplaceRecordDelimInText(sBuf(8))
            
            'Page , leftPos, topPos, rightPos, bottomPos, SnoID, Flags, Prefix, ItemData, Description
            ItemData = Replace(sBuf(8), "^t", vbTab) 'Done in Medicomp code
        
            Set objItem = New FormItem
            With objItem
               .Page = sBuf(0)
               .Left = sBuf(1)
               .Top = sBuf(2)
               .Right = sBuf(3)
               .Bottom = sBuf(4)
               .SnoID = sBuf(5)
               .Flags = sBuf(6)
               .Prefix = sBuf(7)
               .ItemData = ItemData
               .Description = ReplaceRecordDelimInText(sBuf(9))
            End With
            FormItems.Add objItem
        End If
    Loop
    Close #fileNumber
    
    ImportMedicompFormFile = True
    
    Exit Function
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "FormTemplate.ImportMedicompFormFile()", "TemplateClient", vbExclamation)
End Function

Public Function ReplaceRecordDelimInText(ByVal sText As String) As String
    ReplaceRecordDelimInText = Replace(sText, gsRecordDelim, "~!~")
End Function

Public Function RestoreRecordDelimInText(ByVal sText As String) As String
  RestoreRecordDelimInText = Replace(sText, "~!~", gsRecordDelim)
End Function
