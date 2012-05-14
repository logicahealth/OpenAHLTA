VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TextNote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mbHasItemBeenSaved As Boolean
Private mbIsComplete As Boolean
Private mbIsDirty As Boolean
Private mbIsNew As Boolean
Private mbSensitive As Boolean

Private mdDateCreated As Date
Private mdDateModified As Date
Private mdDTS As Date

Private mlEncounterID As Long
Private mlModifiedBy As Long
Private mlOwnerNCID As Long

Private msCategory As String
Private msEventID As String
Private msOwnerName As String
Private msRTF As String
Private msTitle As String

Private mobjDataString As DataString
Private Const COMPRESSED_DATA_ID As String = "COMPRESSED_DATA_ID="

'****************************************************************************************
'  Function Name: DeleteDataString
'
'  Description: This function deletes the datastring specified by the passed in
'               Event Id.
'
'  Paramters:
'    vsEventId - Contains the "EventId" of the string to delete.
'
'  Return Value:
'    Returns TRUE if the datastring was deleted successfully.
'    Returns FALSE if an error occurred in deleting the datastring.
'****************************************************************************************
Private Function DeleteDataString(ByVal vsEventId As String) As Boolean
    
    Dim sOwnerAppNCID As String
    Dim bSuccess As Boolean
    Dim objDataString As DataString
    
    On Error GoTo ErrHandler
    DeleteDataString = False
        
    gobjDas.ClearAll '---clear out previous "data event" strings---
    gobjDas.GeneralCriteria.ClearAll '---clear previous search criteria---
    gobjDas.GeneralCriteria.EventID = vsEventId '---set search criteria---
    gobjDas.DoSearch '---get the data to delete---
    If gobjDas.EOF Then     '--if data does not exist--
       Beep
       MsgBox "Unable to find the datastring for EventId: " & vsEventId & " to delete.", vbExclamation, "Clinical Notes"
       Exit Function
    Else
        Set objDataString = gobjDas.Events '---point to the datastring---
    End If
    
    sOwnerAppNCID = objDataString.Application.ValueConcept.UniqueId
    objDataString.DeleteFlag = True '---set the "delete" flag---

    '------------------------------------------------------------------------------------------------------
    '  Each time you "modify" any datastring, certain properties are "cleared out".  You MUST reset these
    '    properties BEFORE you perform a "save".
    '------------------------------------------------------------------------------------------------------
    objDataString.Application.ValueConcept.UniqueId = sOwnerAppNCID
    objDataString.Facility.ValueConcept.UniqueId = cwobjLogonEnviron.FacilityNCID

    '---delete the clinical note---
    DeleteDataString = objDataString.Save

    If DeleteDataString = False Then
       Beep
       MsgBox "Unable to delete the datastring for EventId: " & vsEventId & ".", vbExclamation, "Text Notes"
       Exit Function
    End If
    
    GoTo CleanUpAndExit
    
ErrHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "frmClinicalNotes.mDeleteDataString", "Clinical Notes", vbExclamation)
    
CleanUpAndExit:
    Set objDataString = Nothing '---release reference to object---

End Function


Public Sub LoadNote(NoteItem As uListItem)
    
    Dim sSQL As String
    Dim sDate As String
    
    On Error GoTo ErrHandler
    
    With NoteItem
        sDate = Format(.Date_Created, gDefaultDateTime)
        sSQL = "SELECT * FROM GENERIC_APP_NOTES WHERE UNIT_NUMBER = " & cwobjPatient.UnitNumber & " AND " & vbLf & _
                     "DATE_CREATED = '" & sDate & "' AND CREATED_BY = " & CStr(.Entered_By_NCID)
    End With
    
    gSQL.Execute sSQL
    If Not (gSQL.BOF And gSQL.EOF) Then
       gSQL.MoveFirst
       msCategory = gSQL("CATEGORY")
       msTitle = gSQL("TITLE")
       OwnerNCID = Val(gSQL("CREATED_BY"))
       mlEncounterID = gSQL("ENCOUNTER_NUMBER")
       msEventID = gSQL("DATA_ID")
       mdDateCreated = gSQL("DATE_CREATED")
       mdDateModified = gSQL("DATE_MODIFIED")
       ModifiedBy = Val(gSQL("MODIFIED_BY"))
       mbIsNew = False
       mbIsDirty = False
       SetUpDataString
    End If
    
    mbHasItemBeenSaved = False
    Exit Sub
    
ErrHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "TextNote.LoadNote", "Note", vbExclamation)
    'Resume Next
    
End Sub
Public Function DeleteNote(NoteItem As uListItem) As Boolean
    
    Dim sSQL As String
    Dim sDate As String
    
    On Error GoTo ErrHandler
    DeleteNote = False
    
    With NoteItem
        sDate = Format(.Date_Created, gDefaultDateTime)
        sSQL = "SELECT * FROM GENERIC_APP_NOTES WHERE UNIT_NUMBER = " & cwobjPatient.UnitNumber & " AND " & vbLf & _
                     "DATE_CREATED = '" & sDate & "' AND CREATED_BY = " & CStr(.Entered_By_NCID)
    End With
    
    gSQL.Execute sSQL
    If Not (gSQL.BOF And gSQL.EOF) Then
       gSQL.MoveFirst
       msCategory = gSQL("CATEGORY")
       msTitle = gSQL("TITLE")
       mlOwnerNCID = gSQL("CREATED_BY")
       mlEncounterID = gSQL("ENCOUNTER_NUMBER")
       msEventID = CStr(Val(gSQL("DATA_ID")))
       mdDateCreated = Val(gSQL("DATE_CREATED"))
       mdDateModified = gSQL("DATE_MODIFIED")
       mlModifiedBy = Val(gSQL("MODIFIED_BY"))
       mbIsNew = False
       mbIsDirty = False
    End If
    
    If DeleteDataString(CStr(msEventID)) Then
       sSQL = "DELETE FROM GENERIC_APP_NOTES WHERE UNIT_NUMBER = " & cwobjPatient.UnitNumber & " AND " & vbLf & _
                     "DATE_CREATED = '" & sDate & "' AND CREATED_BY = " & CStr(NoteItem.Entered_By_NCID)
       gSQL.Execute sSQL
    End If
    
    DeleteNote = True
    Exit Function
    
ErrHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "TextNote.DeleteNote", "Note", vbExclamation)
    'Resume Next
    
End Function


Private Function SaveGEMSTextNote() As Long
    
    On Error GoTo ErrHandler
    SaveGEMSTextNote = 0
    
    GoTo CleanUpAndExit
    
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "NoteTool.SaveGEMSTextNote", "Note Tool", vbExclamation)
    'Resume Next
    
CleanUpAndExit:


End Function

Private Function SaveTextNote() As Long
        
    Dim sHTMLString As String
    Dim bSuccess As Boolean
    Dim sBuf As String
    Dim objClinicalText As ClinicalText
    Dim sOwnerAppNCID As String
    Dim bSlientMode As Boolean

    Dim aBytes() As Byte
    Dim objZlib As New ClinNoteZLib
    Dim nOriginalSize As Long
    Dim bCompressed As Boolean
    Dim objStrings As New DataString
    Dim objClinObsHeader As ClinObsHeader
    Dim objClinicalEvent As ClinicalEvent
    Dim objClinicalObservation As ClinicalObservation
    
    Dim dtToday As Date
    dtToday = Now
    
    On Error GoTo ErrHandler
    SaveTextNote = 0
    
    Set objClinicalText = mobjDataString.PatientData
    
    objClinicalText.EventDate = Now
    mobjDataString.EncounterFacility.ValueConcept.UniqueId = cwobjLogonEnviron.FacilityNCID
    mobjDataString.EncounterNumber = CStr(mlEncounterID)
    mobjDataString.Clinician.ValueConcept.UniqueId = gobjClinician.NCID
    '---only "set" the "PhysicalLocation.InfoType" if its not set already---
    If mobjDataString.PointOfCare.PhysicalLocation.InfoType <> "facility" Then
       mobjDataString.PointOfCare.PhysicalLocation.InfoType = "facility"
    End If
    mobjDataString.PointOfCare.PhysicalLocation.Facility.ValueConcept.UniqueId = cwobjLogonEnviron.FacilityNCID
    objClinicalText.ContextInfo.DictatingClinician.ValueConcept.UniqueId = gobjClinician.NCID
    objClinicalText.PointOfCare.PhysicalLocation.Facility.ValueConcept.UniqueId = cwobjLogonEnviron.FacilityNCID

    sHTMLString = msRTF
    nOriginalSize = Len(sHTMLString)
        
    'compress data
    aBytes = StrConv(Replace(sHTMLString, Chr(0), ""), vbFromUnicode)
            
    Set objZlib = New ClinNoteZLib
    Call objZlib.CompressData(aBytes)
    bCompressed = True
            
    Set objZlib = Nothing
            
    'Create ClinicalEvent to store compressed data
    'Set up DataString & ClinObsHeader
    objStrings.UnitNumber = cwobjPatient.UnitNumber
    objStrings.DataType = 5
            
    Set objClinicalEvent = objStrings.PatientData
    Set objClinObsHeader = objClinicalEvent.Header
                
    With objClinObsHeader
           .ObsBatId.InfoType = "att"
           .ObsBatId.Value.InfoType = "codedWOSform"
           .ObsBatId.Value.CodedWOSform.ValueConcept.UniqueId = "209643"
           .TestStatus.InfoType = "att"
           .TestStatus.Value.InfoType = "codedWOSform"
           .ObsDateTime.InfoType = "timeIntAtt"
           .ObsDateTime.StartTime.InfoType = "att"
           .ObsDateTime.StartTime.Value.InfoType = "dateTime"
           .ObsDateTime.EndTime.InfoType = "att"
           .ObsDateTime.EndTime.Value.InfoType = "dateTime"
    End With
            
    With objStrings
           If gParentNCID = "" Then
              .Application.ValueConcept = NCID_CLINICAL_NOTES_APPLICATION
           Else
              .Application.ValueConcept = gParentNCID
           End If
           .Clinician.ValueConcept = cwobjLogonEnviron.UserNCID
           .EncounterFacility.ValueConcept = cwobjLogonEnviron.FacilityNCID
           .EncounterNumber = mlEncounterID
           .Facility.ValueConcept = cwobjLogonEnviron.FacilityNCID
           .PointOfCare.PhysicalLocation.InfoType = "facility"
           .PointOfCare.PhysicalLocation.Facility.ValueConcept = cwobjLogonEnviron.FacilityNCID
    End With
        
    objClinObsHeader.ObsDateTime.StartTime.Value.DateTime = dtToday
    objClinObsHeader.ObsDateTime.EndTime.Value.DateTime = dtToday
    Set objClinicalObservation = objClinicalEvent.ClinObs.Add
    
    With objClinicalObservation
           .ObsTypeNcid = "119294"
           .ObsId.InfoType = "att"
           .ObsId.Value.InfoType = "codedWOSform"
           .ObsId.Value.CodedWOSform.ValueConcept.UniqueId = "209643"
           '---value---
           .ObsValue.InfoType = "att"
           .ObsValue.Value.InfoType = "multiMedia"
           .ObsValue.Value.MultiMedia.MediaType.ValueConcept.UniqueId = "162798"
           .ObsValue.Value.MultiMedia.MediaSubType.ValueConcept.UniqueId = "162811"
           .ObsValue.Value.MultiMedia.Data = aBytes
           If bCompressed Then
              .ObsValue.Value.MultiMedia.UncompressedLength = CStr(nOriginalSize)
              .ObsValue.Value.MultiMedia.CompressionScheme.ValueConcept.UniqueId = "98765"
           End If
    End With
        
    objStrings.Save
    objClinicalText.ReportText.TextObject.Text.Text = COMPRESSED_DATA_ID & CStr(objStrings.DataID)
            
    Set objStrings = Nothing
    Set objClinObsHeader = Nothing
    Set objClinicalEvent = Nothing
    Set objClinicalObservation = Nothing
        
    '--------------------------------------------------------------------------------------------------------------------
    '  Each time you "modify" any clinicial text string, certain properties are "cleared out".  You MUST reset these
    '    properties BEFORE you perform a "save".
    '
    '  NOTE: If the user is saving a "new" string, "mark" it as being owned by the Clinical Notes Application.  If the
    '        user is saving an existing string, restore original "owner".  This is needed because of earlier versions of
    '        CW which "tagged" clinical notes as "owned" by Clinical Workstation and NOT Clinical Notes.
    '--------------------------------------------------------------------------------------------------------------------
    mobjDataString.PatientData.ReportType.ValueConcept.UniqueId = 69587
    If gParentNCID = "" Then
       sOwnerAppNCID = NCID_CLINICAL_NOTES_APPLICATION
    Else
       sOwnerAppNCID = gParentNCID
    End If
    mobjDataString.Application.ValueConcept.UniqueId = sOwnerAppNCID
    mobjDataString.Facility.ValueConcept.UniqueId = cwobjLogonEnviron.FacilityNCID
    mobjDataString.Save
    SaveTextNote = mobjDataString.DataID
    GoTo CleanUpAndExit
    
ErrHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "NoteTool.SaveTextNote", "Note Tool", vbExclamation)
    'Resume Next
    
CleanUpAndExit:
    Set objClinicalText = Nothing
    Set objZlib = Nothing
    Set objStrings = Nothing
    Set objClinObsHeader = Nothing
    Set objClinicalEvent = Nothing
    Set objClinicalObservation = Nothing
        
End Function
'''''''''''''''''''''''''''''''''''''''''''''''
' SCR #10622 Provide means of setting a       '
' Note as sensitve.                           '
'''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get Sensitive() As Boolean        '
    Sensitive = mbSensitive                      '''''''''''''''''''
End Property                                                        '
Public Property Let Sensitive(ByVal p_bSensitive As Boolean)        '
    mbSensitive = p_bSensitive                                     '
End Property                                                        '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get DTS() As Date
    If CDbl(mdDTS) = 0 Then mdDTS = Now
    DTS = mdDTS
End Property

Public Property Let DTS(ByVal RHV As Date)
    mdDTS = RHV
End Property

Public Property Get OwnerNCID() As Long
    OwnerNCID = mlOwnerNCID
End Property

Public Property Let OwnerNCID(ByVal RHV As Long)
    mlOwnerNCID = RHV
    If mlOwnerNCID = 0 Then
        OwnerName = ""
    Else
        gConceptGems.UniqueId = mlOwnerNCID
        msOwnerName = gConceptGems.PrefRep("2000").Representation
    End If
End Property

Public Property Get OwnerName() As String
    OwnerName = msOwnerName
End Property

Public Property Let OwnerName(ByVal RHV As String)
    msOwnerName = RHV
End Property

Public Property Get Category() As String
    Category = msCategory
End Property

Public Property Let Category(ByVal RHV As String)
    msCategory = RHV
End Property

Friend Sub SetUpDataString()
    
    Const iCLINICAL_TEXT_DATA As Integer = 1
    Const sTRANSCRIPTION_INFO = "transcriptionInfo"
    Dim objClinicalText As ClinicalText
    Dim objClinicalEvent As ClinicalEvent
    Dim objClinOb As ClinicalObservation
    Dim objClinObs As ClinicalObservations
    Dim bImageFound As Boolean
    Dim DataID As Long
    
    Dim aBytes() As Byte
    Dim objZlib As New ClinNoteZLib
    Dim nOriginalSize As Long
        
    On Error GoTo ErrHandler
    If gobjDas Is Nothing Then Set gobjDas = New DasCtrl
    
    If mbIsNew Then
       gobjClinician.ClearClinician
       '-----------------------------------------------------------------------------
       '  On a "new" note, always default to the corresponding "Storing..." fields
       '    in the logon environment object.
       '-----------------------------------------------------------------------------
       gobjClinician.Name = cwobjLogonEnviron.StoringUserName
       gobjClinician.NCID = cwobjLogonEnviron.StoringUserNCID
       gobjClinician.RoleNCID = cwobjLogonEnviron.StoringUserRoleNCID
       gobjClinician.FacilityNCID = cwobjLogonEnviron.StoringUserFacilityNCID
       '---create a new datastring object---
       Set mobjDataString = New DataString
    
       '---set "predefined" fields of the "datastring" object that don't change, but need to be set in order to store the datastring---
       mobjDataString.DataType = iCLINICAL_TEXT_DATA
        
       If gParentNCID = "" Then
          mobjDataString.Application.ValueConcept.UniqueId = NCID_CLINICAL_NOTES_APPLICATION
       Else
          mobjDataString.Application.ValueConcept.UniqueId = gParentNCID
       End If
        
       mobjDataString.UnitNumber = cwobjPatient.UnitNumber
       mobjDataString.Clinician.ValueConcept.UniqueId = gobjClinician.NCID
       mobjDataString.PointOfCare.PhysicalLocation.InfoType = "facility"
       mobjDataString.PointOfCare.PhysicalLocation.Facility.ValueConcept.UniqueId = cwobjLogonEnviron.FacilityNCID
       mobjDataString.Facility.ValueConcept.UniqueId = cwobjLogonEnviron.FacilityNCID
       mobjDataString.EncounterFacility.ValueConcept.UniqueId = cwobjLogonEnviron.FacilityNCID
       mobjDataString.EncounterNumber = CStr(mlEncounterID)
    
       '---set "predefined" fields of the "patient data" object that don't change, but need to be set in order to store the datastring---
       Set objClinicalText = mobjDataString.PatientData    '--point to the "clinical text" part of the datastring--
       objClinicalText.ReportStatus.ValueConcept.UniqueId = NCID_RESULT_STATUS_FINAL
       objClinicalText.CompressionScheme.ValueConcept.UniqueId = NCID_UNCOMPRESSED_SCHEME
       objClinicalText.ReportText.TextType = "formatted"
       objClinicalText.ReportText.TextObject.TextType = "htmlText"
       objClinicalText.PointOfCare.PhysicalLocation.InfoType = "facility"
       objClinicalText.ContextInfo.InfoType = sTRANSCRIPTION_INFO
       objClinicalText.ContextInfo.DictatingClinician.ValueConcept.UniqueId = gobjClinician.NCID
       objClinicalText.EventDate = Now
       objClinicalText.PointOfCare.PhysicalLocation.Facility.ValueConcept.UniqueId = cwobjLogonEnviron.FacilityNCID
    Else
       '---clear out previous "data event" strings---
       gobjDas.ClearAll
       gobjDas.GeneralCriteria.ClearAll '---clear previous search criteria---
       gobjDas.GeneralCriteria.EventID = msEventID '---set search criteria---
       gobjDas.GeneralCriteria.ReturnMultiMedia = True
       gobjDas.GeneralCriteria.UnitNumber = cwobjPatient.UnitNumber
       gobjDas.DoSearch '---get the data to edit---
       If gobjDas.EOF Then     '--if data does not exist--
          MsgBox "Unable to find the datastring for EventId: " & msEventID & " to edit.", vbExclamation, "Clinical Notes"
          GoTo CleanUpAndExit
       Else
          Set mobjDataString = gobjDas.Events '---point to this datastring---
          If mobjDataString.IsEmpty Then GoTo CleanUpAndExit
          
          Set objClinicalText = mobjDataString.PatientData
          '---extract the HTML text---
          If Not (Left$(objClinicalText.ReportText.TextObject.Text.Text, Len(COMPRESSED_DATA_ID)) = COMPRESSED_DATA_ID) Then
             msRTF = objClinicalText.ReportText.TextObject.Text.Text
          Else
             DataID = Int(Right$(objClinicalText.ReportText.TextObject.Text.Text, Len(objClinicalText.ReportText.TextObject.Text.Text) - Len(COMPRESSED_DATA_ID)))
             msRTF = GetDocument(DataID)
          End If
       End If
    End If
    
    GoTo CleanUpAndExit
    
ErrHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "TextNote.SetUpDataString", "Text Notes", vbExclamation)
    'Resume Next
    
CleanUpAndExit:
    Set objClinicalText = Nothing
    Set objClinicalEvent = Nothing
    Set objClinOb = Nothing
    Set objClinObs = Nothing
    Set objZlib = Nothing
    
End Sub

Public Function GetDocument(DataID As Long) As String
   
   Dim iCount As Integer
   Dim objDataString As DataString
   Dim objClinicalText As ClinicalText
   Dim objClinicalEvent As ClinicalEvent
   Dim objClinOb As ClinicalObservation
   Dim objClinObs As ClinicalObservations
   Dim bImageFound As Boolean
   Dim objDas As New DasCtrl
   
    Dim aBytes() As Byte
    Dim objZlib As New ClinNoteZLib
    Dim nOriginalSize As Long
    
   On Error GoTo ErrHandler
   
   bImageFound = False
      
   '---Clear out previous "data event" strings.---
   objDas.ClearAll
   
   '---Clear previous search criteria.---
   objDas.GeneralCriteria.ClearAll
   
   '---Set search criteria common for the specific Image."---
   
   objDas.GeneralCriteria.EventID = DataID
   objDas.GeneralCriteria.ReturnMultiMedia = True

   objDas.DoSearch
   
   While (Not objDas.EOF And Not bImageFound)
      '--data exists--
      Set objDataString = objDas.Events
      If Not objDataString.IsEmpty Then
         Set objClinicalEvent = objDataString.PatientData
         If Not objClinicalEvent.IsEmpty Then
            Set objClinObs = objClinicalEvent.ClinObs
            If Not objClinObs.IsEmpty Then
                Set objClinOb = objClinObs.Item(1)
                If Not objClinOb.IsEmpty Then
                   If objClinOb.ObsValue.InfoType = "att" Then
                      If objClinOb.ObsValue.Value.InfoType = "multiMedia" Then
                         If Not objClinOb.ObsValue.Value.MultiMedia.IsEmpty Then
                         
                              aBytes = objClinOb.ObsValue.Value.MultiMedia.Data
                              Set objZlib = New ClinNoteZLib
                              nOriginalSize = objClinOb.ObsValue.Value.MultiMedia.UncompressedLength
                              Call objZlib.DecompressData(aBytes, nOriginalSize)
                              Set objZlib = Nothing
                              
                              GetDocument = StrConv(aBytes, vbUnicode)
                     
                              bImageFound = True
                         End If
                      End If
                   End If
                End If
            End If
         End If
      End If
      objDas.GoToNext
   Wend
   
   If Not bImageFound Then
      Beep
      MsgBox "No Image data found for Patient Unit# " & cwobjPatient.UnitNumber & ", Event Id " & DataID, vbExclamation & "."
   End If
   
   Exit Function

ErrHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description & vbCrLf & "Patient Unit# " & cwobjPatient.UnitNumber & vbCrLf & "Event Id " & DataID, _
        "TextNote.GetDocument ", "", vbExclamation)

    '---set return value---
    Exit Function

End Function

Public Property Get Title() As String
    Title = msTitle
End Property

Public Property Let Title(ByVal RHV As String)
    msTitle = RHV
End Property

Public Property Get IsComplete() As Boolean
    IsComplete = mbIsComplete
End Property

Public Property Let IsComplete(ByVal RHV As Boolean)
    mbIsComplete = RHV
End Property

Public Property Get IsSensitive() As Boolean
    IsSensitive = mbSensitive
End Property

Public Property Let IsSensitive(ByVal RHV As Boolean)
    mbSensitive = RHV
End Property


Public Property Get IsDirty() As Boolean
    IsDirty = mbIsDirty
End Property

Public Property Let IsDirty(ByVal RHV As Boolean)
    mbIsDirty = RHV
End Property

Public Property Get IsNew() As Boolean
    IsNew = mbIsNew
End Property

Public Property Let IsNew(ByVal RHV As Boolean)
    mbIsNew = RHV
End Property
Public Function SaveNote() As Boolean
    
    Dim sSQL As String
    Dim sDate As String
    Dim lEventID As Long
    Dim oConn As New CHCSII_CONN.Conn
    
    On Error GoTo ErrHandler
    Screen.MousePointer = vbHourglass
    SaveNote = False
    
    lEventID = SaveGEMSTextNote
    
    If lEventID = 0 Then Exit Function
    
    If mbIsNew Then
       sDate = Format(DTS, gDefaultDateTime)
       sSQL = "INSERT INTO GENERIC_APP_NOTES  FIELDS (UNIT_NUMBER, FACILITY_NCID, DATE_CREATED, CREATED_BY, ENTRY_TOOL_NCID, "
       sSQL = sSQL & "CATEGORY, TITLE, ENCOUNTER_NUMBER, DATA_ID)"
       sSQL = sSQL & " VALUES (" & CStr(cwobjPatient.UnitNumber) & ", " & oConn.SQLQuote(cwobjLogonEnviron.FacilityNCID) & ", '" & sDate
       sSQL = sSQL & "', " & cwobjLogonEnviron.UserNCID & ", " & oConn.SQLQuote(gParentNCID) & ", " & oConn.SQLQuote(msCategory) & ", " & oConn.SQLQuote(msTitle)
       sSQL = sSQL & ", " & oConn.SQLQuote(CStr(mlEncounterID)) & ", " & CStr(lEventID) & ") "
       gSQL.Execute sSQL
       mdDateCreated = DTS
       OwnerNCID = cwobjLogonEnviron.UserNCID
    Else
       sDate = Format(Now, gDefaultDateTime)
       sSQL = "UPDATE GENERIC_APP_NOTES " & vbLf
       sSQL = sSQL & "SET CATEGORY = " & oConn.SQLQuote(msCategory) & vbLf
       sSQL = sSQL & ", TITLE = " & oConn.SQLQuote(msTitle) & vbLf
       sSQL = sSQL & ", MODIFIED_BY = " & CStr(cwobjLogonEnviron.UserNCID) & vbLf
       sSQL = sSQL & ", DATE_MODIFIED = '" & sDate & "'" & vbLf
       sSQL = sSQL & ", DATA_ID = " & CStr(lEventID) & vbLf
       sSQL = sSQL & " WHERE UNIT_NUMBER = " & CStr(cwobjPatient.UnitNumber) & vbLf & _
                            " AND DATE_CREATED = '" & Format(mdDateCreated, gDefaultDateTime) & _
                            "' AND CREATED_BY = " & cwobjLogonEnviron.UserNCID
        gSQL.Execute sSQL
        mdDateModified = Now
        mlModifiedBy = cwobjLogonEnviron.UserNCID
    End If
    
    msEventID = CStr(lEventID)
    
    mbHasItemBeenSaved = True
    mbIsNew = False
    mbIsDirty = False
    
    SaveNote = True
    GoTo CleanUpAndExit
    
ErrHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "TextNote.SaveNote", "Note", vbExclamation)
    Resume Next
    
CleanUpAndExit:
    Screen.MousePointer = vbDefault

End Function

Public Property Get RTF() As String
    RTF = msRTF
End Property

Public Property Let RTF(ByVal RHV As String)
    msRTF = RHV
End Property

Public Property Get EncounterID() As Long
    EncounterID = mlEncounterID
End Property

Public Property Let EncounterID(ByVal RHV As Long)
    mlEncounterID = RHV
End Property

Public Property Get DateCreated() As Date
    DateCreated = mdDateCreated
End Property

Public Property Let DateCreated(ByVal RHV As Date)
    mdDateCreated = RHV
End Property

Public Property Get DateModified() As Date
    DateModified = mdDateModified
End Property

Public Property Let DateModified(ByVal RHV As Date)
    mdDateModified = RHV
End Property

Public Property Get ModifiedBy() As Long
    ModifiedBy = mlModifiedBy
End Property

Public Property Let ModifiedBy(ByVal RHV As Long)
    mlModifiedBy = RHV
    If Not mlModifiedBy = 0 Then
        gConceptGems.UniqueId = mlOwnerNCID
        msOwnerName = gConceptGems.PrefRep("2000").Representation
    End If
End Property

Private Sub Class_Terminate()
    Set mobjDataString = Nothing
End Sub



Public Property Get EventID() As String
    EventID = msEventID
End Property

Public Property Let EventID(ByVal RHV As String)
    msEventID = RHV
End Property

Public Property Get HasItemBeenSaved() As Boolean
    HasItemBeenSaved = mbHasItemBeenSaved
End Property
