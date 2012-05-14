VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Dental"
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

Public Property Get RTF() As String
    RTF = FormatRtfDisplay(msRTF, "Dental", "Written", msOwnerName, mdDTS)
End Property

Public Sub SetRTF(ByRef RTF As String)
    If RTF <> msRTF Then
        msRTF = RTF
        mbDirty = True
    End If
End Sub


Public Sub Load(ByRef objParent As SectionParent, ByRef colDataStrings As Collection)
    Dim objData As DATASTRATLLib.DataString
    Dim objEvent As CLINTYPEATLLib.ClinicalEvent
    Dim i As Long

    Dim oShared As New CWShared
    
    If oShared.IsAppMode(modeCHCSI_GUI) Then '<SCR SCR 36493
        Exit Sub
    End If
    
    msRTF = ""
    mbDirty = False
    mnDataID = 0
    
    'Using the For...Each method with 3M objects causes memory leak
    For i = 1 To colDataStrings.Count
    'For Each objData In colDataStrings
        Set objData = colDataStrings(i)
        If objData.DataTypeNCID = NCID_ENCEVENT Then
            Set objEvent = objData.PatientData
            If objEvent.Header.ObsBatId.value.Coded.ValueConcept.UniqueId = SectionEnumToNCID(eDental) Then
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
                End If
                Exit For
            End If
        End If
    Next 'objData
End Sub

Public Sub GEMS_Load(ByRef objParent As SectionParent, objRS As ADODB.Recordset)
   
   '<bjr082401 theater
    Dim oConcept As New GEMS_ConceptCtrl

    msRTF = ""
    mbDirty = False
    mnDataID = 0
    
    On Error Resume Next
    If objRS.EOF Then Exit Sub
    
    msOwnerNCID = "" & objRS("ownerncid")
    oConcept.UniqueId = msOwnerNCID
    msOwnerName = oConcept.PrefRep("2000").Representation
    mdDTS = objRS("DTS")
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

    Set oConcept = Nothing
                    
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
'    If Not mobjEncounter.LockingObject.LockSection(eDental, 0, SectionLock) Then
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
'        mobjEncounter.LockingObject.UnlockSection eDental, 0
'    End If
'
'End Function
Public Function OpenSection() As Boolean
    On Error Resume Next
    If mbRegularViewRunning Then
        Exit Function
    End If
    If Not mobjEncounter.LockingObject.LockSection(eDental, 0, SectionLock) Then
        Exit Function
    End If
'<<SCR 10673,10675(dup) Encounter data refresh before children load
'   Dropping this line to stop the refresh
'   mobjEncounter.RefreshDisplay
'>>
    If msOwnerNCID <> "" _
    And (mobjEncounter.Status = Updating Or msOwnerNCID <> mobjEncounter.Logon.UserNCID) Then
        If Not mobjEncounter.Status = Updating Then
            ' a user other than the owner is opening this section
            ' if they really meant to do this, then save a copy to the change history and
            ' save the current data as belonging to the new owner
            If MsgBxARMd("The Dental section of this encounter is currently owned by " & msOwnerName & "." & vbCrLf & vbCrLf _
            & "Do you want to save a copy in change history and take over the section?", vbYesNo, "Encounter Section Ownership") = vbNo Then
                GoTo CleanUp
            End If
        End If
        If Not mobjEncounter.AddHistoryItem(Overwritten, "Dental Note", RTF) Then
            MsgBxARMd "Unable to copy section to change history.", vbCritical, "Encounter"
            GoTo CleanUp
        End If
        mbDirty = True
        If Not Save Then
            GoTo CleanUp
        End If
    End If
    If mobjModule Is Nothing Then
        Set mobjModule = mobjEncounter.Comm.InitializeOLEServer(SectionEnumToNCID(eDental))
        If mobjModule Is Nothing Then
            GoTo CleanUp
        End If
        Set mobjIChild = mobjModule
        If mobjIChild Is Nothing Then
            Set mobjModule = Nothing
            MsgBxARMd "The Dental application module does not support the EncounterOps.IChild interface.", vbInformation, "Encounter"
            GoTo CleanUp
        End If
        On Error Resume Next
        Set mobjICoreAPI3 = mobjModule
        Call mobjEncounter.Comm.Message(cwiGET_APPLICATION_PROGID, msProgID, NCID_CW_APPLICATION, SectionEnumToNCID(eDental))
    End If
    Call mobjIChild.SetParent(mobjParent)
    mbRegularViewRunning = mobjEncounter.EncounterParent.OpenChildView(mobjModule)
CleanUp:
    If mbRegularViewRunning Then
        OpenSection = True
    Else
        Call mobjIChild.SetParent(Nothing)
        Call mobjEncounter.LockingObject.UnlockSection(eDental, 0)
    End If
End Function

Public Function CloseSection() As Boolean
'''''''''''''''''''''''''''''''''
' Testing Refresh.              '
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
            Call mobjEncounter.LockingObject.UnlockSection(eDental, 0)
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
    Call mobjEncounter.LockingObject.UnlockSection(eDental, 0)
End Sub

Public Function Save() As Boolean
' Assume section already locked for update before getting here
   Dim objData As DataString
   On Error GoTo ErrHandler

   If Not mbDirty Then
      Save = True
      Exit Function
   End If
   
    Save = GEMS_Save
    Exit Function

ErrHandler:
   Call mobjEncounter.CWShared.ShowVBError(Err.Number, Err.Description, "Enctr.Save ", "Enc", vbCritical)

End Function

Public Function GEMS_Save() As Boolean
   '<bjr082801 theater - derived from RFV.GEMS_Save
   ' Assume section already locked for update before getting here
   'Dim objData As DataString '<bjr082801
   'Dim mobjDAS     As ICHCSII_DAS
   Dim mobjSQL     As ICHCSII_SQL.ISqlOpsEx
   'Dim oRS         As ADODB.Recordset
   Dim sql As String

   On Error GoTo ErrHandler
   If Not mbDirty Then
      GEMS_Save = True
      Exit Function
   End If
   
   'Check if encounter owner is Dr. Unassigned SCR-24349
   Call mobjEncounter.CheckEncounterOwnership
   
   mnDataID = mobjEncounter.DataID   'theater
   msOwnerNCID = mobjEncounter.Logon.UserNCID
   msOwnerName = mobjEncounter.Logon.UserName
   mdDTS = Now
   
   '<bjr082801 removed for theater
   'Set objData = mobjEncounter.InitializeDataString(mobjParent.meSection, mnDataID, mdDTS, mdDTS)
   'Call AddRTF(objData.PatientData, msRTF)
  
   sql = "select * " & _
   " From enc_sections" & _
   " Where " & _
   " encounternumber =" & mobjEncounter.EncounterID & _
   " and enc_sectionsindex = " & SectionEnumToNCID(eDental)
   
   Set mobjSQL = gobjCHCSIIConn.CHCSII_SQLOPS(Auto)  '''<SCR 36493
   mobjSQL.Execute sql
   If mobjSQL.EOF And mobjSQL.BOF Then
      '- new record do an insert
      mnDataID = mobjSQL.GetNextID(exMAINSEQ)  'GEMS_Seq_NextVal("MainSEQ")
      mobjSQL.AddNew
      mobjSQL.value("DataID") = mnDataID
   End If
    
    'DataID is only supposed to be updated for new records
    mobjSQL.value("facilityncid") = mobjEncounter.FacilityNCID
    mobjSQL.value("encounternumber") = mobjEncounter.EncounterID
    mobjSQL.value("enc_sectionsindex") = Val(SectionEnumToNCID(eDental))
    mobjSQL.value("ownerncid") = msOwnerNCID
    mobjSQL.value("DTS") = mdDTS
    mobjSQL.value("OriginalSize") = Len(msRTF) 'SCR-24683
    Dim objZlib As EncZLib
    Dim aBytes() As Byte
    Set objZlib = New EncZLib
    
    aBytes = StrConv(Replace(msRTF, Chr(0), ""), vbFromUnicode)
    objZlib.CompressData aBytes
    mobjSQL.value("Doc") = aBytes
    'mobjSQL.value("Doc") = CompressRTF(msRTF)
    'mobjSQL.Value("WSRFV") = "" & msWSRFV
    
    'If objData.Save Then
    '   mnDataID = objData.DataId
    '   Save = True
    '   mbDirty = False
    'End If
    If mobjSQL.Update Then
       GEMS_Save = True
       'mnDataID = objData.DataId
       mbDirty = False
    End If
   
   GoTo CleanUp
ErrHandler:
    Call mobjEncounter.CWShared.ShowVBError(Err.Number, Err.Description, "Enctr.GEMS_Save ", "Enc", vbCritical)
CleanUp:
    Set mobjSQL = Nothing
    If Save Then
        mobjEncounter.Comm.SetStatusText "Encounter " & mobjEncounter.EncounterID & " " & SectionEnumToName(mobjParent.meSection) & " section was saved."
    Else
        mobjEncounter.Comm.SetStatusText ""
    End If
End Function
Public Sub NotifyComplete()
    Dim bCreatedObject As Boolean
    On Error Resume Next
    If mobjModule Is Nothing Then
        Set mobjModule = mobjEncounter.Comm.InitializeOLEServer(SectionEnumToNCID(eDental))
        If mobjModule Is Nothing Then
            Exit Sub
        End If
        Set mobjIChild = mobjModule
        If mobjIChild Is Nothing Then
            Set mobjModule = Nothing
            Exit Sub
        End If
        bCreatedObject = True
    End If
    Call mobjIChild.SetParent(mobjParent)
    Call mobjModule.CompletedEncounter
    Call mobjIChild.SetParent(Nothing)
    
    If bCreatedObject Then
        Set mobjModule = Nothing
        Set mobjIChild = Nothing
    End If
End Sub

Public Sub PreEncounterSign()
'HWCC 1/30/02 -
'SCRs 17424(also reference SCR 17422) and SCR 17133 requires
' Dental Parent to launch 2 pop up modal forms before provider
'  sign the encounter. Per Dental SA Erin O'Donohue, those pop
'   up modal forms are to be display before the
'    "Encounter is not Complete" form. The code below call Dental Parent
'     so Dental Parent can pop up those forms required by the SCRs
'CBA 3/12/2
'SCR-22060 - Initialize Dental module if needed
'            Error will be caught by EncSignOps.SignEncounter so only one error message will be shown
    Dim bCreatedObject As Boolean
    
    If mobjModule Is Nothing Then
        Set mobjModule = mobjEncounter.Comm.InitializeOLEServer(SectionEnumToNCID(eDental))
        If mobjModule Is Nothing Then
            Exit Sub
        End If
        Set mobjIChild = mobjModule
        If mobjIChild Is Nothing Then
            Set mobjModule = Nothing
            Exit Sub
        End If
        bCreatedObject = True
    End If
    Call mobjIChild.SetParent(mobjParent)
    
    Call mobjModule.PreEncounterSign   'Calling Dental Parent
    
    Call mobjIChild.SetParent(Nothing)
    
    If bCreatedObject Then
        Set mobjModule = Nothing
        Set mobjIChild = Nothing
    End If
    
End Sub
