VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Vitals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const mstrDegreeTag     As String = "<%%degree tag%%>"
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

Private mcolSnoIDs As Collection

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

Public Property Get SnoIDs() As Collection
    Set SnoIDs = mcolSnoIDs
End Property

'Add NewSnoIDs to colSnoIDs
Public Property Set SnoIDs(ByRef colNewSnoID As Collection)
    Dim i As Integer
    Dim j As Integer
    Dim objLTR As ListToolRecord
    Dim bDuplicate As Boolean
    
    On Error Resume Next
    If colNewSnoID Is Nothing Then
        Set mcolSnoIDs = Nothing  'Delete Collection 'SCR-26003
        Exit Property
    End If

    If mcolSnoIDs Is Nothing Then
        Set mcolSnoIDs = New Collection
    End If
    
    For i = 1 To colNewSnoID.Count
    
        'Check for duplicates
        For j = 1 To mcolSnoIDs.Count
            Set objLTR = mcolSnoIDs.Item(j)
            If objLTR.SnoID = colNewSnoID.Item(i) Then
                bDuplicate = True
            End If
        Next
        
        If Not bDuplicate Then
            Set objLTR = New ListToolRecord
            objLTR.SnoID = colNewSnoID.Item(i)
            mcolSnoIDs.Add objLTR
        End If

        bDuplicate = False
    Next i
    Set objLTR = Nothing
End Property

Public Property Get RTF() As String
    RTF = msRTF
    

End Property

Public Sub SetRTF(ByRef strRTF As String)
    If strRTF <> msRTF Then
        msRTF = strRTF
        mbDirty = True
    End If
End Sub

Public Sub Load(ByRef objParent As SectionParent, ByRef colDataStrings As Collection)
    Dim objData As DATASTRATLLib.DataString
    Dim objEvent As CLINTYPEATLLib.ClinicalEvent
    Dim objClinObs As CLINTYPEATLLib.ClinicalObservation
    Dim objLTR As ListToolRecord
    Dim i As Long
    Dim j As Long
    
    Set mcolSnoIDs = Nothing
    Set mcolSnoIDs = New Collection
    
    msRTF = ""
    mbDirty = False
    mnDataID = 0
    
    'Using the For...Each method with 3M objects causes memory leak
    For i = 1 To colDataStrings.Count
    'For Each objData In colDataStrings
        Set objData = colDataStrings(i)
        If objData.DataTypeNCID = NCID_ENCEVENT Then
            Set objEvent = objData.PatientData
            If objEvent.Header.ObsBatId.value.Coded.ValueConcept.UniqueId = SectionEnumToNCID(eVitals) Then
            
                mnDataID = objData.DataID
                
                ' Medcin Obs are now stored in Vitals SCR-11863
                For j = 1 To objEvent.ClinObs.Count
                    Set objClinObs = objEvent.ClinObs.Item(j)
                    Select Case objClinObs.ObsId.value.CodedWOSform.ValueConcept.UniqueId
                    Case NCID_ENCOUNTERRTF
                        msRTF = ExtractRTF(objClinObs)
                    Case NCID_MEDCIN_OBS
                        Set objLTR = New ListToolRecord
                        Call objLTR.ReadFromClinObs(objClinObs)
                        mcolSnoIDs.Add objLTR
                        Set objLTR = Nothing
                    End Select
                Next j
                
'                If objEvent.ClinObs.Count > 0 Then
'                    msRTF = ExtractRTF(objEvent.ClinObs.Item(1))
'                End If
'                Exit For
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
'    If Not mobjEncounter.LockingObject.LockSection(eVitals, 0, SectionLock) Then
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
'        mobjEncounter.LockingObject.UnlockSection eVitals, 0
'    End If
'
'End Function
Public Function OpenSection() As Boolean
    On Error Resume Next
    If mbRegularViewRunning Then
        Exit Function
    End If
    If Not mobjEncounter.LockingObject.LockSection(eVitals, 0, SectionLock) Then
        Exit Function
    End If
'<<SCR 10673,10675(dup) Encounter data refresh before children load
'   Dropping this line to stop the refresh
'   mobjEncounter.RefreshDisplay
'>>
    If Not mobjEncounter.SectionIsEditable(eVitals) Then
        'MsgBxARMd "This encounter has been saved by another user and you can no longer edit the Vitals section.", vbInformation, "Encounter"
        If mobjEncounter.EncounterType = TeleConsult Then 'SCR#46973
            MsgBxARMd "The Vitals Module is not available when accessed from a Telcon Encounter.", vbInformation, "Encounter"
        Else
            MsgBxARMd "You do not have the required privileges to open the Vitals Module.", vbInformation, "Encounter"
        End If
        GoTo CleanUp
    End If
    If mobjEncounter.Status = Updating Then
        If Not mobjEncounter.AddHistoryItem(Overwritten, "Vitals", RTF) Then
            MsgBxARMd "Unable to copy section to change history.", vbCritical, "Encounter"
            GoTo CleanUp
        End If
        mbDirty = True
        If Not Save Then
            GoTo CleanUp
        End If
    End If
    If mobjModule Is Nothing Then
        Set mobjModule = mobjEncounter.Comm.InitializeOLEServer(SectionEnumToNCID(eVitals))
        If mobjModule Is Nothing Then
            GoTo CleanUp
        End If
        Set mobjIChild = mobjModule
        If mobjIChild Is Nothing Then
            Set mobjModule = Nothing
            MsgBxARMd "The Vitals application module does not support the EncounterOps.IChild interface.", vbInformation, "Encounter"
            GoTo CleanUp
        End If
        On Error Resume Next
        Set mobjICoreAPI3 = mobjModule
        Call mobjEncounter.Comm.Message(cwiGET_APPLICATION_PROGID, msProgID, NCID_CW_APPLICATION, SectionEnumToNCID(eVitals))
    End If
    Call mobjIChild.SetParent(mobjParent)
    mbRegularViewRunning = mobjEncounter.EncounterParent.OpenChildView(mobjModule)
CleanUp:
    If mbRegularViewRunning Then
        OpenSection = True
    Else
        Call mobjIChild.SetParent(Nothing)
        Call mobjEncounter.LockingObject.UnlockSection(eVitals, 0)
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
            'Removed duplicate unlock call.  Calls made in ClosedView or Form_Unload of module SCR-16591
            'Call mobjEncounter.LockingObject.UnlockSection(eVitals, 0)
            mbRegularViewRunning = False
            'CloseSection = True
        End If
        
        'CloseSection = Not CBool(Err)
        'Err.Clear
    End If
End Function

Public Sub ClosedView() ' called by child through SectionParent
    mbRegularViewRunning = False
    Call mobjIChild.SetParent(Nothing)
    Call mobjEncounter.EncounterParent.hidechild
    Call mobjEncounter.LockingObject.UnlockSection(eVitals, 0)
End Sub

Public Function Save() As Boolean
' Assume section already locked for update before getting here
    Dim objData As DataString
    Dim objEvent As ClinicalEvent
    Dim objLTR As ListToolRecord
    
    On Error GoTo ErrHandler
    
    Save = GEMS_Save
    mbDirty = False              ' T.Noren...06/14/10...indicate data clean
    
    GoTo CleanUp
ErrHandler:
    Call mobjEncounter.CWShared.ShowVBError(Err.Number, Err.Description, "Enctr.Save ", "Enc", vbCritical)
CleanUp:
    Set objData = Nothing
    Set objEvent = Nothing
    Set objLTR = Nothing
    
    If Save Then
        mobjEncounter.Comm.SetStatusText "Encounter " & mobjEncounter.EncounterID & " " & SectionEnumToName(mobjParent.meSection) & " section was saved."
    Else
        mobjEncounter.Comm.SetStatusText ""
    End If
End Function

Public Function GEMS_Load(ByRef objParent As SectionParent, objRS As ADODB.Recordset)
    '*********************************************
    ' TODO: NEED TO SAVE LIST TOOL RECORDS FOR VITALS
    '       use A/P as guide
    Dim sVitals As String
    
    msRTF = ""
    mbDirty = False
    mnDataID = 0
    Do While Not objRS.EOF
        mnDataID = Val("" & objRS("DataID"))
        
        'Decompress if compressed SCR-24683
        'If Original Size = 0 then dont load any RTF SCR-854 ITT
        If CVar(objRS("OriginalSize")) > 0 Then
            Dim objZlib As EncZLib
            Set objZlib = New EncZLib
            Dim aBytes() As Byte
            
            aBytes = objRS("Doc")
            Call objZlib.DecompressData(aBytes, objRS("OriginalSize"))
            sVitals = StrConv(aBytes, vbUnicode)
            'sVitals = DecompressRTF("" & objRS("Doc"), objRS("OriginalSize"))
            msRTF = Replace(sVitals, mstrDegreeTag, "°")
        End If

        objRS.MoveNext
    Loop
    
End Function

Private Function GEMS_Save() As Boolean
' Assume section already locked for update before getting here

    '*********************************************
    ' TODO: NEED TO SAVE LIST TOOL RECORDS FOR VITALS
    '       use A/P as guide
    
    On Error GoTo Proc_Error
    
    Dim objSQL      As ICHCSII_SQL.ISqlOpsEx
    Dim strSQL      As String
    Dim sVitals     As String
    
    GEMS_Save = False
    
    If Not mbDirty Then
        GEMS_Save = True
        Exit Function
    End If
    
    'Check if encounter owner is Dr. Unassigned SCR-24349
    Call mobjEncounter.CheckEncounterOwnership

    Set objSQL = gobjCHCSIIConn.CHCSII_SQLOPS(Auto)  '''<SCR 36493

    strSQL = "SELECT DataID, FacilityNCID, EncounterNumber, Enc_SectionsIndex, Doc, OwnerName, OwnerNCID, OriginalSize " & _
             "FROM Enc_Sections " & _
             "WHERE EncounterNumber = " & Str(mobjEncounter.EncounterID) & " " & _
             "AND Enc_SectionsIndex = " & Str(SectionEnumToNCID(eVitals)) & " " & _
             "AND FacilityNCID = " & Str(mobjEncounter.FacilityNCID)
            
    objSQL.Execute strSQL
    
    If (objSQL.BOF And objSQL.EOF) Then
        objSQL.AddNew
        mnDataID = objSQL.GetNextID(exMAINSEQ)  'GEMS_Seq_NextVal("MainSEQ")
        objSQL("DataID") = mnDataID
        objSQL("FacilityNCID") = mobjEncounter.FacilityNCID
        objSQL("EncounterNumber") = mobjEncounter.EncounterID
        objSQL("OwnerName") = mobjEncounter.Logon.UserName 'SCR-22999
        objSQL("OwnerNCID") = mobjEncounter.Logon.UserNCID 'SCR-22999
        objSQL("enc_sectionsindex") = SectionEnumToNCID(eVitals)
    End If
    
    'Compress RTF SCR-24683
    Dim objZlib As EncZLib
    Dim aBytes() As Byte
    Set objZlib = New EncZLib
    
    If Len(msRTF) <> 0 Then
        sVitals = Replace(msRTF, "°", mstrDegreeTag)
        aBytes = StrConv(Replace(sVitals, Chr(0), ""), vbFromUnicode)
        objZlib.CompressData aBytes
        objSQL.value("Doc") = aBytes
    Else
        'Fixed SCR #34615 - Jane Shen 4/18/2003
        objSQL("Doc") = ""
    End If

    'Moved out here just in case deleting Vitals SCR-854 ITT
    objSQL.value("Doc") = aBytes
    
    objSQL("OriginalSize") = Len(Replace(sVitals, "°", mstrDegreeTag))
    'objSQL("Doc") = CompressRTF()
    objSQL.Update
    
    GEMS_Save = True

Proc_Exit:
    Set objSQL = Nothing
    mobjEncounter.Comm.SetStatusText "Encounter " & mobjEncounter.EncounterID & " " & SectionEnumToName(mobjParent.meSection) & " section was saved."
    Exit Function
    
Proc_Error:
    GEMS_Save = False
    mobjEncounter.Comm.SetStatusText ""
    Call mobjEncounter.CWShared.ShowVBError(Err.Number, Err.Description, "Enctr.gems_Save ", "Enc", vbCritical)
    GoTo Proc_Exit
End Function

Private Sub Class_Initialize()
    Set mcolSnoIDs = New Collection
End Sub
