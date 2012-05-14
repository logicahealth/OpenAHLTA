VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RFV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mobjEncounter As Encounter
Private mobjParent As SectionParent
Private msRTF As String
Private mbDirty As Boolean
Private mnDataID As Long
Private mobjModule As Object
Private mobjIChild As IChild
Private mobjICoreAPI3 As CHCSII_CoreAPI3.ICHCSII_CoreAPI3
Private msProgID As String
Private mbRegularViewRunning As Boolean
Private msOwnerNCID As String
Private msOwnerName As String
Private mdDTS As Date
Private mbUpdated As Boolean 'SCR-39592

Private msWSRFV             As String      '- Screening work Status

Friend Property Get DataID() As Long
    DataID = mnDataID
End Property
Friend Property Let DataID(lngDataID As Long)
    mnDataID = lngDataID
End Property
Friend Property Get Dirty() As Boolean
    Dirty = mbDirty
End Property
Friend Property Let Dirty(blnDirty As Boolean)
    mbDirty = blnDirty
End Property

Public Property Get WorkStatusScreening() As String
    WorkStatusScreening = msWSRFV
End Property

Public Property Let WorkStatusScreening(RHV As String)
    If msWSRFV <> RHV Then
        msWSRFV = RHV
        mbDirty = True
    End If
End Property

Public Property Get RTF() As String

    On Error GoTo ErrHandler
    
    Dim oShared As CWShared
    
    Set oShared = New CWShared
    
    'If mbUpdated Then 'SCR-39592
    '    RTF = FormatRtfDisplay(msRTF, "Screening", "Last updated", msOwnerName, mdDTS)
    'Else
        RTF = FormatRtfDisplay(msRTF, "Screening", "Written", msOwnerName, mdDTS)
        
        If Len(oShared.CmdLineSwitches("HIPAA837")) Then 'SCR#53033
            If Not mobjEncounter Is Nothing Then
                If Not mobjEncounter.PregnancyData Is Nothing Then
                    RTF = RTF & mobjEncounter.PregnancyData.RTF(True)
                End If
            End If
        End If
         
    'End If
    
    Set oShared = Nothing
    
    Exit Property
    
ErrHandler:
    Call mobjEncounter.CWShared.ShowVBError(Err.Number, Err.Description, "Enctr.Get RTF", "Enc", vbCritical)
    
End Property

Public Sub SetRTF(ByRef RTF As String)
    'If StrComp(RTF, msRTF, vbTextCompare) Then
        'Change in RFV Section
        
        'If owner is diff then add change history
        If msOwnerNCID <> "" And msOwnerNCID <> "0" _
        And (mobjEncounter.Status = Updating Or msOwnerNCID <> mobjEncounter.Logon.UserNCID) Then
            If Not mobjEncounter.AddHistoryItem(Overwritten, "RFV Note", msRTF, msOwnerName, mdDTS, "Screening") Then
                MsgBxARMd "Unable to copy section to change history.", vbCritical, "Encounter"
                Exit Sub
            End If
            'mbUpdated = True
        End If
        
        'Set RTF
        msRTF = RTF
        
        mbDirty = True
    'End If
End Sub

Public Sub Load(ByRef objParent As SectionParent, ByRef colDataStrings As Collection)
    Dim objData As DATASTRATLLib.DataString
    Dim objEvent As CLINTYPEATLLib.ClinicalEvent
    Dim objCommentGroup As CLINTYPEATLLib.CommentGroup
    Dim objCommentItem As CLINTYPEATLLib.CommentItem
    Dim sName As String
    Dim sValue As String
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    msWSRFV = ""
    msRTF = ""
    mbDirty = False
    mnDataID = 0
    
    'Using the For...Each method with 3M objects causes memory leak
    For i = 1 To colDataStrings.Count
    'For Each objData In colDataStrings
        Set objData = colDataStrings(i)
        If objData.DataTypeNCID = NCID_ENCEVENT Then
            Set objEvent = objData.PatientData
            If objEvent.Header.ObsBatId.value.Coded.ValueConcept.UniqueId = SectionEnumToNCID(eRFV) Then
                mnDataID = objData.DataID
                If Not objData.ModifyClinician Is Nothing Then
                    msOwnerNCID = objData.ModifyClinician.ValueConcept.UniqueId
                    msOwnerName = objData.ModifyClinician.ValueConcept.PrefRep("2000").Representation
                ElseIf Not objData.CreateClinician Is Nothing Then
                    msOwnerNCID = objData.CreateClinician.ValueConcept.UniqueId
                    msOwnerName = objData.CreateClinician.ValueConcept.PrefRep("2000").Representation
                End If
                mdDTS = objEvent.Header.ObsDateTime.StartTime.value.DateTime
                If objEvent.ClinObs.Count > 0 Then
                    msRTF = ExtractRTF(objEvent.ClinObs.Item(1))
                    For j = 1 To objEvent.Header.Comments.Count
                    'For Each objCommentGroup In objEvent.Header.Comments
                        Set objCommentGroup = objEvent.Header.Comments(j)
                        For k = 1 To objCommentGroup.Comments.Count
                        'For Each objCommentItem In objCommentGroup.Comments
                            Set objCommentItem = objCommentGroup.Comments(k)
                            Call ParseCommentItem(objCommentItem, sName, sValue)
                            Select Case sName
                            Case "WSRFV"
                                msWSRFV = sValue
                            Case "UPDATED"
                                mbUpdated = CBool(sValue)
                            End Select
                        Next 'objCommentItem
                    Next 'objCommentGroup
                End If
                Exit For
            End If
        End If
    Next 'objData
End Sub

Public Sub MenuItemSelected(ByVal MenuID As Long)
    On Error Resume Next
    Call mobjModule.MenuItemSelected(MenuID)
End Sub

Public Sub OpenEncounter(ByRef objParent As SectionParent, ByRef objEncounter As Encounter)
    Set mobjParent = objParent
    Set mobjEncounter = objEncounter
End Sub

Public Sub CloseEncounter()
    Set mobjParent = Nothing
    Set mobjEncounter = Nothing
    Set mobjModule = Nothing
    Set mobjIChild = Nothing
    Set mobjICoreAPI3 = Nothing
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' SCR #6890 Refresh issues.                              '
' Purpose:                                               '
' Refresh the windows of a section already active.       '
' Parameters:                                            '
' None.                                                  '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Public Function RefreshSection() As Boolean

'
'    On Error Resume Next
'    If Not mbRegularViewRunning Then
'        RefreshSection = False '999
'        Exit Function
'    End If
'
'    If Not mobjEncounter.LockingObject.LockSection(eRFV, 0, SectionLock) Then
'        Exit Function
'    End If
'
'    mobjIChild.SetParent mobjParent
'    mbRegularViewRunning = mobjEncounter.EncounterParent.RefreshChildView(mobjModule)
'
'    If mbRegularViewRunning Then
'        RefreshSection = True
'    Else
'        mobjIChild.SetParent Nothing
'        mobjEncounter.LockingObject.UnlockSection eRFV, 0
'    End If
'
'End Function
Public Function OpenSection() As Boolean
    On Error Resume Next
    If mbRegularViewRunning Then
        Exit Function
    End If
    If Not mobjEncounter.LockingObject.LockSection(eRFV, 0, SectionLock) Then
        Exit Function
    End If
'<<SCR 10673,10675(dup) Encounter data refresh before children load
'   Dropping this line to stop the refresh
'   mobjEncounter.RefreshDisplay
'>>
    '''''''''''''''''''''''''''''''''''''''''
    ' SCR #9984 Error Verbage for Telcons.  '
    '''''''''''''''''''''''''''''''''''''''''
    If Not mobjEncounter.SectionIsEditable(eRFV) Then
        If mobjEncounter.EncounterType = TeleConsult Then 'SCR#46973
            MsgBxARMd "The Screening Module is not available when accessed from a Telcon Encounter.", vbInformation, "Encounter"
        Else
            MsgBxARMd "You do not have the required privileges to open the RFV Module.", vbInformation, "Encounter"
            'MsgBxARMd "This encounter has been saved by another user and you can no longer edit the RFV section.", vbInformation, "Encounter"
            GoTo CleanUp
        End If
    End If
    
    If mobjEncounter.EncounterType = InpatientNote Then 'SCR#42959
        MsgBxARMd "The Screening Module is not available when accessed from an Inpatient Note Encounter.", vbInformation, "Encounter"
        GoTo CleanUp
    End If
    
    If msOwnerNCID <> "" _
    And (mobjEncounter.Status = Updating Or msOwnerNCID <> mobjEncounter.Logon.UserNCID) Then
'        If Not mobjEncounter.Status = Updating Then
            ' a user other than the owner is opening this section
            ' if they really meant to do this, then save a copy to the change history and
            ' save the current data as belonging to the new owner
            'Answer is always Yes SCR-26954
'            If MsgBxARMd("The Screening section of this encounter is currently owned by " & msOwnerName & "." & vbCrLf & vbCrLf _
'            & "Do you want to save a copy in change history and take over the section?", vbYesNo, "Encounter Section Ownership") = vbNo Then
'                GoTo Cleanup
'            End If
'        End If
        If Not mobjEncounter.AddHistoryItem(Overwritten, "Screening Note", RTF) Then
            MsgBxARMd "Unable to copy section to change history.", vbCritical, "Encounter"
            GoTo CleanUp
        End If
        mbDirty = True
        If Not Save Then
            GoTo CleanUp
        End If
    End If
    If mobjModule Is Nothing Then
        Set mobjModule = mobjEncounter.Comm.InitializeOLEServer(SectionEnumToNCID(eRFV))
        If mobjModule Is Nothing Then
            GoTo CleanUp
        End If
        Set mobjIChild = mobjModule
        If mobjIChild Is Nothing Then
            Set mobjModule = Nothing
            MsgBxARMd "The Screening application module does not support the EncounterOps.IChild interface.", vbInformation, "Encounter"
            GoTo CleanUp
        End If
        On Error Resume Next
        Set mobjICoreAPI3 = mobjModule
        Call mobjEncounter.Comm.Message(cwiGET_APPLICATION_PROGID, msProgID, NCID_CW_APPLICATION, SectionEnumToNCID(eRFV))
    End If
    Call mobjIChild.SetParent(mobjParent)
    mbRegularViewRunning = mobjEncounter.EncounterParent.OpenChildView(mobjModule)
CleanUp:
    If mbRegularViewRunning Then
        OpenSection = True
    Else
        Call mobjIChild.SetParent(Nothing)
        Call mobjEncounter.LockingObject.UnlockSection(eRFV, 0)
    End If
End Function

Public Function CloseSection() As Boolean
'''''''''''''''''''''''''''''''''
' Testing Refresh               '
' SCR 11399  Unwanted Refresh.  '
'''''''''''''''''''''''''''''''''
On Error Resume Next
If mbRegularViewRunning Then
        If mobjICoreAPI3 Is Nothing Then
            Call mobjModule.CloseView(cwiREGULAR_VIEW)
            CloseSection = True
        Else
            CloseSection = mobjICoreAPI3.CloseView2(cwiREGULAR_VIEW)
        End If
        If CloseSection Then
            Call mobjIChild.SetParent(Nothing)
            'Removed duplicate unlock call.  Calls made in ClosedView or Form_Unload of module SCR-16591
            'Call mobjEncounter.LockingObject.UnlockSection(eRFV, 0)
            mbRegularViewRunning = False
        End If
        
        'CloseSection = Not CBool(Err)
        'Err.Clear
    End If
End Function

Public Sub ClosedView() ' called by child through SectionParent
    mbRegularViewRunning = False
    Call mobjIChild.SetParent(Nothing)
    Call mobjEncounter.EncounterParent.hidechild
    Call mobjEncounter.LockingObject.UnlockSection(eRFV, 0)
End Sub

Public Function Save() As Boolean
' Assume section already locked for update before getting here
    
    On Error GoTo ErrHandler
    'SCR 50480 - make sure we save if the HIPAA data has changed
    If Not mbDirty And Not mobjEncounter.AccidentData.IsDirty And Not mobjEncounter.PregnancyData.IsDirty Then
        Save = True
        Exit Function
    End If
        
    Save = GEMS_Save
    
    GoTo CleanUp
ErrHandler:
    Call mobjEncounter.CWShared.ShowVBError(Err.Number, Err.Description, "Enctr.Save ", "Enc", vbCritical)
CleanUp:
    If Save Then
        mobjEncounter.Comm.SetStatusText "Encounter " & mobjEncounter.EncounterID & " " & SectionEnumToName(mobjParent.meSection) & " section was saved."
    Else
        mobjEncounter.Comm.SetStatusText ""
    End If
End Function

Public Sub GEMS_Load(ByRef objParent As SectionParent, objRS As ADODB.Recordset)
    Dim oConcept As New GEMS_ConceptCtrl
    
    msWSRFV = ""
    msRTF = ""
    mbDirty = False
    mnDataID = 0
    On Error Resume Next
    If objRS.EOF Then Exit Sub
     
    mnDataID = Val("" & objRS("DataID"))
    msOwnerNCID = "" & objRS("ownerncid")
    If Not msOwnerNCID = "" Then
        oConcept.UniqueId = msOwnerNCID
        msOwnerName = oConcept.PrefRep("2000").Representation
    End If
    If IsDate(objRS("DTS")) Then
        mdDTS = objRS("DTS")
    End If
    
    'Decompress if compressed SCR-24683
    If CVar(objRS("OriginalSize")) > 0 Then
        Dim objZlib As EncZLib
        Set objZlib = New EncZLib
        Dim aBytes() As Byte
        
        aBytes = objRS("Doc")
        Call objZlib.DecompressData(aBytes, objRS("OriginalSize"))
        msRTF = StrConv(aBytes, vbUnicode)
        'msRTF = DecompressRTF("" & objRS("Doc"), objRS("OriginalSize"))
    End If
    'msRTF = "" & objRS("Doc")
    msWSRFV = "" & objRS("WSRFV")

    Set oConcept = Nothing
End Sub


Public Function GEMS_Save() As Boolean
' Assume section already locked for update before getting here
    Dim mobjDAS     As ICHCSII_DAS
    Dim mobjSQL     As ICHCSII_SQL.ISqlOpsEx
    Dim oRS         As ADODB.Recordset
    Dim sql As String
    
    On Error GoTo ErrHandler
    If Not mbDirty And Not mobjEncounter.AccidentData.IsDirty And Not mobjEncounter.PregnancyData.IsDirty Then
        GEMS_Save = True
        Exit Function
    End If
    
    'Check if encounter owner is Dr. Unassigned SCR-24349
    Call mobjEncounter.CheckEncounterOwnership
    
    mnDataID = mobjEncounter.DataID
    msOwnerNCID = mobjEncounter.Logon.UserNCID
    msOwnerName = mobjEncounter.Logon.UserName 'Set owner name SCR-26083
    mdDTS = Now
    
    'WRogers 8/2001
    sql = "select * " & _
    " From enc_sections" & _
    " Where " & _
    " encounternumber =" & mobjEncounter.EncounterID & _
    " and enc_sectionsindex = " & SectionEnumToNCID(eRFV)
    
    Set mobjSQL = gobjCHCSIIConn.CHCSII_SQLOPS(Auto)  '''<SCR 36493
    
    mobjSQL.Execute sql
     
    If mobjSQL.EOF And mobjSQL.BOF Then
'- new record do an insert
        mnDataID = mobjSQL.GetNextID(exMAINSEQ) '  GEMS_Seq_NextVal("MainSEQ")
        mobjSQL.AddNew
        mobjSQL.value("DataID") = mnDataID  'WRogers SCR 15167
    End If
    
    'WRogers DataID is only supposed to be updated for new records
    'mobjSQL.Value("DataID") = mnDataID
    mobjSQL.value("facilityncid") = mobjEncounter.FacilityNCID
    mobjSQL.value("encounternumber") = mobjEncounter.EncounterID
    mobjSQL.value("enc_sectionsindex") = Val(SectionEnumToNCID(eRFV))
    mobjSQL.value("ownerncid") = msOwnerNCID
    mobjSQL.value("ownername") = mobjEncounter.Logon.UserName
    mobjSQL.value("DTS") = mdDTS
    
    mobjSQL.value("OriginalSize") = Len(msRTF) 'SCR-24683
  
    'Don't compress if RTF does not exists SCR-25712
    If Len(msRTF) Then
        Dim objZlib As EncZLib
        Dim aBytes() As Byte
        Set objZlib = New EncZLib
        
        aBytes = StrConv(Replace(msRTF, Chr(0), ""), vbFromUnicode)
        objZlib.CompressData aBytes
    End If
    
    mobjSQL.value("Doc") = aBytes
    'mobjSQL.value("Doc") = CompressRTF("" & msRTF)
    mobjSQL.value("WSRFV") = "" & msWSRFV

    If mobjSQL.Update Then
        GEMS_Save = mobjEncounter.SaveHIPAAData()
        mbDirty = False
    End If
    
    GoTo CleanUp
ErrHandler:
    Call mobjEncounter.CWShared.ShowVBError(Err.Number, Err.Description, "Enctr.gems_Save ", "Enc", vbCritical)
CleanUp:
    Set mobjSQL = Nothing
    If GEMS_Save Then
        mobjEncounter.Comm.SetStatusText "Encounter " & mobjEncounter.EncounterID & " " & SectionEnumToName(mobjParent.meSection) & " section was saved."
    Else
        mobjEncounter.Comm.SetStatusText ""
    End If
End Function

