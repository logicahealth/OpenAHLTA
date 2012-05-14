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
Attribute VB_Name = "AP"
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

Private mcolRelatedDiagnoses As Collection
Private meAPStatus As EncStatus
Private mcolSnoIDs As Collection

Public Property Get DTS() As Date
    DTS = mdDTS
End Property

Public Property Get RTF() As String

    If mbUpdated Then 'SCR-39592
        RTF = FormatRtfDisplay(msRTF, "A/P", "Last updated", msOwnerName, mdDTS)
    Else
        RTF = FormatRtfDisplay(msRTF, "A/P", "Written", msOwnerName, mdDTS)
    End If
End Property

Public Property Set RelatedDiagnoses(ByRef RHV As Collection)
    Dim objDiag As Object
    Dim i As Integer
    On Error Resume Next
    If RHV Is Nothing Then
        If mcolRelatedDiagnoses.Count = 0 Then
            Exit Property
        End If
        GoTo SetData
    End If
    If RHV.Count <> mcolRelatedDiagnoses.Count Then
        GoTo SetData
    End If
    For i = 1 To RHV.Count
        Set objDiag = RHV.Item(i)
        If objDiag.DataId <> mcolRelatedDiagnoses.Item(i) Then
            GoTo SetData
        End If
    Next i
    Set objDiag = Nothing
    Exit Property
SetData:
    mbDirty = True
    Set mcolRelatedDiagnoses = New Collection
    For Each objDiag In RHV
        mcolRelatedDiagnoses.Add objDiag.DataId
    Next objDiag
End Property

Public Property Get Status() As EncStatus
    Status = meAPStatus
End Property

Public Property Let APStatus(ByVal RHV As EncStatus)
    If meAPStatus <> RHV Then
        mbDirty = True
        meAPStatus = RHV
    End If
End Property

Public Property Get Dirty() As Boolean
    Dirty = mbDirty
End Property
Public Property Let Dirty(blnDirty As Boolean)
    mbDirty = blnDirty
End Property
Friend Property Get DataId() As Long
    DataId = mnDataID
End Property
Friend Property Let DataId(lngDataID As Long)
    mnDataID = lngDataID
End Property

Friend Property Get SnoIDs() As Collection
    Set SnoIDs = mcolSnoIDs
End Property

Public Property Set SnoIDs(ByRef RHV As Collection)
    'SCR 104309 changed code to check for multiple SnoID
    Dim i As Integer
    Dim j As Integer
    Dim objLTR As ListToolRecord
    Dim bDuplicate As Boolean
    
    On Error Resume Next
    If RHV Is Nothing Then
        Set mcolSnoIDs = Nothing
        Exit Property
    End If

    If mcolSnoIDs Is Nothing Then
        Set mcolSnoIDs = New Collection
    End If
    
    For i = 1 To RHV.Count
    
        'Check for duplicates
        For j = 1 To mcolSnoIDs.Count
            Set objLTR = mcolSnoIDs.Item(j)
            If objLTR.SnoID = RHV.Item(i).SnoID Then
                bDuplicate = True
            End If
        Next
        
        If Not bDuplicate Then
            Set objLTR = RHV.Item(i)
            mcolSnoIDs.Add objLTR
            mbDirty = True
        End If

        bDuplicate = False
    Next i
    Set objLTR = Nothing
    
    'A/P will now pass collection of ListToolRecords SCR#49601
    'So conversion is no longer needed
'    Set mcolSnoIDs = RHV
'    mbDirty = True
    
'    Dim i As Integer
'    Dim objLTR As ListToolRecord
'    On Error Resume Next
'    If RHV Is Nothing Then
'        If mcolSnoIDs.Count = 0 Then
'            Exit Property
'        End If
'        GoTo SetData
'    End If
'    If RHV.Count <> mcolSnoIDs.Count Then
'        GoTo SetData
'    End If
'    For i = 1 To RHV.Count
'        Set objLTR = mcolSnoIDs.Item(i)
'        If Val(objLTR.SnoID) <> mcolSnoIDs.Item(i) Then
'            GoTo SetData
'        End If
'    Next i
'    Set objLTR = Nothing
'    Exit Property
'SetData:
'    mbDirty = True
'    Set mcolSnoIDs = New Collection
'    For i = 1 To RHV.Count
'        Set objLTR = New ListToolRecord
'        objLTR.SnoID = RHV.Item(i)
'        mcolSnoIDs.Add objLTR
'    Next i
End Property

Private Sub Class_Initialize()
    Set mcolRelatedDiagnoses = New Collection
    Set mcolSnoIDs = New Collection
End Sub

Private Sub Class_Terminate()
    Set mcolRelatedDiagnoses = Nothing
    Set mcolSnoIDs = Nothing
End Sub
'AP will call SetRTF only if information has been updated
Public Sub SetRTF(ByRef RTF As String, Optional Updated As Boolean)
        
    'If owner is diff and updated then add change history SCR#39592
    If msOwnerNCID <> "" And msOwnerNCID <> "0" _
    And (mobjEncounter.Status = Updating Or msOwnerNCID <> mobjEncounter.Logon.UserNCID) Then
        If Not mobjEncounter.AddHistoryItem(Overwritten, "A/P Note", msRTF, msOwnerName, mdDTS, "A/P") Then
            MsgBxARMd "Unable to copy section to change history.", vbCritical, "Encounter"
            Exit Sub
        End If
        mbUpdated = True
    End If
    
    'Set RTF
    msRTF = RTF
    
    'Reset DTS when udpating RTF in Provider GUI mode SCR-31811
    If mobjEncounter.CWShared.IsAppMode(modeCHCSI_GUI) Then '<SCR SCR 36493
        mdDTS = Now
    End If
    
    mbDirty = True
End Sub

Public Sub Load(ByRef objParent As SectionParent, ByRef colDataStrings As Collection)
    Dim objData As DATASTRATLLib.DataString
    Dim objEvent As CLINTYPEATLLib.ClinicalEvent
    Dim objClinObs As CLINTYPEATLLib.ClinicalObservation

    Dim objCommentGroup As CLINTYPEATLLib.CommentGroup

    Dim objCommentItem As CLINTYPEATLLib.CommentItem
    Dim objLTR As ListToolRecord
    Dim sName As String
    Dim sValue As String
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim l As Long
    Dim oShared As New CWShared
'    Dim AnPRTF As WAMData.RTF
    
    Set mcolRelatedDiagnoses = New Collection
    Set mcolSnoIDs = New Collection
    meAPStatus = Incomplete
    
    If oShared.IsAppMode(modeCHCSI_GUI) Then '<SCR 36493
'        Set AnPRTF = colDataStrings.Item(CStr(eap))
'
'        mdDTS = AnPRTF.OnsetDate
'        msOwnerName = AnPRTF.WrittenBy
'        msRTF = AnPRTF.RTF
'        meAPStatus = colDataStrings.Item("AnPStatus")
        Exit Sub
    End If
    
    msRTF = ""
    mbDirty = False
    mnDataID = 0
    'Using the For...Each method with 3M objects causes memory leak
    For j = 1 To colDataStrings.Count
    'For Each objData In colDataStrings
        Set objData = colDataStrings(j)
        If objData.DataTypeNCID = NCID_ENCEVENT Then
            Set objEvent = objData.PatientData
            If objEvent.Header.ObsBatId.value.Coded.ValueConcept.UniqueId = SectionEnumToNCID(eap) Then
                mnDataID = objData.DataId
                If Not objData.ModifyClinician Is Nothing Then
                    msOwnerNCID = objData.ModifyClinician.ValueConcept.UniqueId
                    msOwnerName = objData.ModifyClinician.ValueConcept.PrefRep("2000").Representation
                ElseIf Not objData.CreateClinician Is Nothing Then
                    msOwnerNCID = objData.CreateClinician.ValueConcept.UniqueId
                    msOwnerName = objData.CreateClinician.ValueConcept.PrefRep("2000").Representation
                End If
                mdDTS = objEvent.Header.ObsDateTime.StartTime.value.DateTime
                meAPStatus = StatusNCIDToEnum(objEvent.Header.TestStatus.value.CodedWOSform.ValueConcept)
'                For Each objCommentGroup In objEvent.Header.Comments
'                    For Each objCommentItem In objCommentGroup.Comments
'                        Call ParseCommentItem(objCommentItem, sName, sValue)
'                        Select Case sName
'                        End Select
'                    Next objCommentItem
'                Next objCommentGroup
                For i = 1 To objEvent.ClinObs.Count
                    Set objClinObs = objEvent.ClinObs.Item(i)
                    Select Case objClinObs.ObsId.value.CodedWOSform.ValueConcept.UniqueId
                    Case NCID_ENCOUNTERRTF
                        msRTF = ExtractRTF(objClinObs)
                    Case NCID_MEDCIN_OBS
                        Set objLTR = New ListToolRecord
                        Call objLTR.ReadFromClinObs(objClinObs)
                        mcolSnoIDs.Add objLTR
                        Set objLTR = Nothing
                    End Select
                Next i
                For i = 1 To objData.SemanticLinks.Count
                    mcolRelatedDiagnoses.Add objData.SemanticLinks.Item(i).ObjectPointer.StringKey.StringId
                Next i
                
                For i = 1 To objEvent.Header.Comments.Count
                    Set objCommentGroup = objEvent.Header.Comments(i)
                    For k = 1 To objCommentGroup.Comments.Count
                        Set objCommentItem = objCommentGroup.Comments(k)
                        Call ParseCommentItem(objCommentItem, sName, sValue)
                        Select Case sName
                        Case "UPDATED"
                            mbUpdated = CBool(sValue)
                        End Select
                    Next 'k objCommentItem
                Next 'j objCommentGroup
                
                
                Exit For
            End If
        End If
    Next 'j (objData)
End Sub

Public Sub GEMS_Load(ByRef objParent As SectionParent, objRS As ADODB.Recordset)

    Dim objLTR As ListToolRecord
    Dim sName As String
    Dim sValue As String
    Dim i As Integer
    Dim EncDataId As Long
    Dim LTRDataID As Long
    Dim sql As String
    Dim objZlib As EncZLib
    Dim aBytes() As Byte
    Dim dDTS As Date
    Dim sOwnerNCID As String
    
    
    Dim oShared As CWShared
    
    Set oShared = New CWShared
    
'    If colDataStrings Is Nothing Then Exit Sub
'    If colDataStrings.Count = 0 Then Exit Sub
'
    'EncDataId = colDataStrings(1)
    Set mcolRelatedDiagnoses = New Collection
    Set mcolSnoIDs = New Collection
    meAPStatus = Incomplete
    'msRTF = ""
    mbDirty = False
    mnDataID = 0
    
    Dim oSQL As ICHCSII_SQL.ISqlOps
    Set oSQL = gobjCHCSIIConn.CHCSII_SQLOPS(Auto)  '''<SCR 36493
    
    If objRS.EOF Then
        '- make one
        mnDataID = oSQL.GetNextID()
        sql = "insert into enc_sections " & _
              "(dataid, encounternumber, facilityncid, enc_sectionsindex) values (" & _
              " " & mnDataID & ", " & _
              " " & mobjEncounter.EncounterID & ", " & _
              " " & mobjEncounter.FacilityNCID & ", " & _
              " " & SectionEnumToNCID(eap) & ") "
              
        oSQL.Execute sql
        
    Else
        
        
        If Not IsNull(objRS!DTS) Then
            dDTS = CDate(objRS!DTS)
        End If
        
        sOwnerNCID = "" & objRS!OwnerNCID
        
        msOwnerName = "" & objRS!OwnerName
        mnDataID = Val("" & objRS!DataId)
        meAPStatus = StatusNCIDToEnum(Val("" & objRS!Status))
        
        If Not IsNull(objRS!UPDATED_FLAG) Then
            mbUpdated = CBool(objRS!UPDATED_FLAG)
        End If
        
        If oShared.IsAppMode(modeCHCSI_GUI) Then
        
            'Default to Incomplete 'SCR#48472
            If meAPStatus = 0 Then meAPStatus = Incomplete
            
            If Len(sOwnerNCID) And sOwnerNCID <> "0" Then
                'Check if data has been updated SCR#43993
                If Not (dDTS = mdDTS And StrComp(sOwnerNCID, msOwnerNCID, vbTextCompare) = 0) Then
                    'msRTF = oShared.CHCSConnection.ReadChunk("ENC_SECTIONS", "ENCOUNTERNUMBER = " & objRS("ENCOUNTERNUMBER") & " AND enc_sectionsindex = " & SectionEnumToNCID(eap), "DOC")
                    msRTF = ReadChunkCHCSConn("ENC_SECTIONS", "ENCOUNTERNUMBER = " & objRS("ENCOUNTERNUMBER") & " AND enc_sectionsindex = " & SectionEnumToNCID(eap), "DOC", meAPStatus = complete)
                End If
            End If
        Else
            'Decompress if compressed SCR-24683
            If CVar(objRS!OriginalSize) > 0 Then
                Set objZlib = New EncZLib
                aBytes = objRS!Doc
                Call objZlib.DecompressData(aBytes, objRS!OriginalSize)
                msRTF = StrConv(aBytes, vbUnicode)
                'msRTF = DecompressRTF("" & objRS!Doc, objRS!OriginalSize)
            End If
        End If
        
        msOwnerNCID = sOwnerNCID
        mdDTS = dDTS
        'SCR 107019
        Set mobjParent = objParent
        CheckForMedChanges
        
    End If

     '- the dataid in listtoolrecordS = the dataid in the parent listtoolrecorD
    sql = "select lr.dataid as dataid, lrs.* from listtoolrecord lr, listtoolrecords lrs " & _
          " where encounternumber = " & mobjEncounter.EncounterID & _
          " and facilityncid = " & mobjEncounter.FacilityNCID & _
          " and SourceNCID = " & SectionEnumToNCID(eap)
          
    If mobjEncounter.CWShared.UseMySQLDB Then
        sql = "select lr.dataid as dataid, lrs.* from listtoolrecord lr LEFT OUTER JOIN listtoolrecords lrs ON (lr.dataid = lrs.dataid)" & _
        " where encounternumber = " & mobjEncounter.EncounterID & _
          " and facilityncid = " & mobjEncounter.FacilityNCID & _
          " and SourceNCID = " & SectionEnumToNCID(eap)
    ElseIf mobjEncounter.CWShared.UseSQLServer Then
        sql = sql & " and lr.dataid *= lrs.dataid "
    Else
        sql = sql & " and lr.dataid = lrs.dataid (+)"
    End If
    
    oSQL.Execute sql
      
    Do While Not oSQL.EOF
        Set objLTR = New ListToolRecord
        Call objLTR.GEMS_ReadFromClinObs(oSQL)
        mcolSnoIDs.Add objLTR
        Set objLTR = Nothing
        oSQL.MoveNext
    Loop
     
    oSQL.Execute "select diagnosisncid from enc_diagnosis where encounternumber = " & mobjEncounter.EncounterID & " and facilityncid = " & mobjEncounter.FacilityNCID
    Do While Not oSQL.EOF
        mcolRelatedDiagnoses.Add oSQL!DiagnosisNCID
        oSQL.MoveNext
    Loop


  Set oSQL = Nothing
End Sub

Private Sub CheckForMedChanges()
'SCR 107019
    On Error GoTo ErrHandler
    
    Dim sSql As String
    Dim Rs As Recordset
    Dim Connection As CHCSII_CONN.Conn
    Dim oDas As ICHCSII_DAS
    Dim oShared As CWShared
    Dim bHasTemplate As Boolean
   
    Set oShared = New CWShared

    Set Connection = New CHCSII_CONN.Conn
    Set oDas = Connection.CHCSII_DAS
  
    sSql = "SELECT * FROM oe_rx WHERE enc_fac_ncid = " & mobjEncounter.FacilityNCID & _
                      " and enc_num = '" & mobjEncounter.EncounterID & "' " & " and CHILD_MED_IEN is NULL" & _
                      " and order_date > " & gobjCHCSIIConn.SQLDate(mdDTS)
    Debug.Print sSql
    Set Rs = oDas.OpenRecordset(sSql)
    If Rs.EOF = False Then
        'Medications have been changed in the Order Management Module for the encounter
        If mobjModule Is Nothing Then
            Set mobjModule = mobjEncounter.Comm.InitializeOLEServer(SectionEnumToNCID(eap))
            If mobjModule Is Nothing Then
                GoTo CleanUp
            End If
            Set mobjIChild = mobjModule
            If mobjIChild Is Nothing Then
                Set mobjModule = Nothing
                MsgBxARMd "The A/P application module does not support the EncounterOps.IChild interface.", vbInformation, "Encounter"
                GoTo CleanUp
            End If
            On Error Resume Next
            Set mobjICoreAPI3 = mobjModule
            Call mobjEncounter.Comm.Message(cwiGET_APPLICATION_PROGID, msProgID, NCID_CW_APPLICATION, SectionEnumToNCID(eap))
        End If
        
        
        'Determine if template exists for encounter
        If Len(Trim$(mobjEncounter.TemplateData)) > 0 Then
            bHasTemplate = True
        Else
            bHasTemplate = False
        End If
    
        mobjParent.OpenEncounter mobjEncounter
        Call mobjIChild.SetParent(mobjParent)
'SCR 162028
'        Call mobjModule.InitializeAnp(, bHasTemplate)
        GEMS_Save
CleanUp:
        On Error Resume Next
        If Not mbRegularViewRunning Then
            Call mobjIChild.SetParent(Nothing)
            Call mobjEncounter.LockingObject.UnlockSection(eap, 0)
            Call mobjEncounter.LockingObject.UnlockSection(eHeader, 0)    'SCR 51459 - Release lock on the providers info
        End If
    End If
  
    Exit Sub
  
ErrHandler:
    On Error Resume Next
    oShared.ShowVBError Err.Number, Err.Description, "ProxyCDRPlus.AddNewPI", "A/P", vbExclamation
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
'    If Not mobjEncounter.LockingObject.LockSection(eAP, 0, SectionLock) Then
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
'        mobjEncounter.LockingObject.UnlockSection eAP, 0
'    End If
'
'End Function
Public Function OpenSection() As Boolean
    On Error Resume Next
    
    Dim bHasTemplate As Boolean
    
    If mbRegularViewRunning Then
        Exit Function
    End If
    If Not mobjEncounter.LockingObject.LockSection(eap, 0, SectionLock) Then
        Exit Function
    End If
    
    'SCR 51459 - Obtain a lock on the providers info when entering A/P
    If Not mobjEncounter.LockingObject.LockSection(eHeader, 0, SectionLock, "Providers") Then
      Exit Function
    End If
   
'<<SCR 10673,10675(dup) Encounter data refresh before children load
'   Dropping this line to stop the refresh
'   mobjEncounter.RefreshDisplay
'>>
    If Not mobjEncounter.SectionIsEditable(eap) Then
        MsgBxARMd "This encounter has been saved by another user and you can no longer edit the AP section.", vbInformation, "Encounter"
        GoTo CleanUp
    End If
'    Moved to SetRTF section
'    If msOwnerNCID <> "" And msOwnerNCID <> "0" _
'    And (mobjEncounter.Status = Updating Or msOwnerNCID <> mobjEncounter.Logon.UserNCID) Then
''        If Not mobjEncounter.Status = Updating Then
'            ' a user other than the owner is opening this section
'            ' if they really meant to do this, then save a copy to the change history and
'            ' save the current data as belonging to the new owner
'            'Answer is always Yes SCR-26954
''            If MsgBxARMd("The AP section of this encounter is currently owned by " & msOwnerName & "." & vbCrLf & vbCrLf _
''            & "Do you want to save a copy in change history and take over the section?", vbYesNo, "Encounter Section Ownership") = vbNo Then
''                GoTo Cleanup
''            End If
''        End If
'        If Not mobjEncounter.AddHistoryItem(Overwritten, "A/P Note", RTF) Then
'            MsgBxARMd "Unable to copy section to change history.", vbCritical, "Encounter"
'            GoTo CleanUp
'        End If
'        mbDirty = True
'        If Not Save Then
'            GoTo CleanUp
'        End If
'    End If
    If mobjModule Is Nothing Then
        Set mobjModule = mobjEncounter.Comm.InitializeOLEServer(SectionEnumToNCID(eap))
        If mobjModule Is Nothing Then
            GoTo CleanUp
        End If
        Set mobjIChild = mobjModule
        If mobjIChild Is Nothing Then
            Set mobjModule = Nothing
            MsgBxARMd "The A/P application module does not support the EncounterOps.IChild interface.", vbInformation, "Encounter"
            GoTo CleanUp
        End If
        On Error Resume Next
        Set mobjICoreAPI3 = mobjModule
        Call mobjEncounter.Comm.Message(cwiGET_APPLICATION_PROGID, msProgID, NCID_CW_APPLICATION, SectionEnumToNCID(eap))
    End If
    
    
    'Determine if template exists for encounter
    If Len(Trim$(mobjEncounter.TemplateData)) > 0 Then
        bHasTemplate = True
    Else
        bHasTemplate = False
    End If

    Call mobjIChild.SetParent(mobjParent)
    Call mobjModule.InitializeAnp(, bHasTemplate)
    
    mbRegularViewRunning = mobjEncounter.EncounterParent.OpenChildView(mobjModule)
    
'    AP does not use WAM when loading AP data
'    If mobjEncounter.CWShared.IsAppMode(modeCHCSI_GUI) Then
'        'Reset WAM object so data is refreshed in case another user added WAM info SCR#44687
'        mobjEncounter.ResetWAM
'    End If
    
CleanUp:
    If mbRegularViewRunning Then
        OpenSection = True
    Else
        Call mobjIChild.SetParent(Nothing)
        Call mobjEncounter.LockingObject.UnlockSection(eap, 0)
        Call mobjEncounter.LockingObject.UnlockSection(eHeader, 0)    'SCR 51459 - Release lock on the providers info
    End If
End Function

Public Function CloseSection() As Boolean
'''''''''''''''''''''''''''''''''
' Testing Refresh.              '
' SCR 11399  Unwanted Refresh.  '
'''''''''''''''''''''''''''''''''
On Error Resume Next
    Screen.MousePointer = vbHourglass
    If mbRegularViewRunning Then
        If mobjICoreAPI3 Is Nothing Then
            Call mobjModule.CloseView(cwiREGULAR_VIEW)
            CloseSection = (Not mbRegularViewRunning)
        Else
            CloseSection = mobjICoreAPI3.CloseView2(cwiREGULAR_VIEW)
        End If
        If CloseSection Then
            Call mobjIChild.SetParent(Nothing)
            'Removed duplicate unlock call.  Calls made in ClosedView or Form_Unload of module SCR-16591
            'Call mobjEncounter.LockingObject.UnlockSection(eap, 0)
            'call only if CloseSection succeeds SCR-32630
            ' PERF CloseView also called by AP module SCR-38428
            ' Private Sub CloseAnP()
            'ClosedView
        End If
                
        
        'CloseSection = Not CBool(Err)
        'Err.Clear
    End If
    Screen.MousePointer = vbDefault
End Function

Public Sub ClosedView() ' also called by child through SectionParent
    mbRegularViewRunning = False
    Call mobjIChild.SetParent(Nothing)
    Call mobjEncounter.EncounterParent.hidechild
    Call mobjEncounter.LockingObject.UnlockSection(eap, 0)
    Call mobjEncounter.LockingObject.UnlockSection(eHeader, 0)    'SCR 51459 - Release lock on the providers info when closing A/P
End Sub

Public Function Save() As Boolean
' Assume section already locked for update before getting here
    Dim objData As DataString
    Dim objEvent As ClinicalEvent
    Dim objHeader As ClinObsHeader
    Dim objLink As SemanticLink
    Dim objLTR As ListToolRecord
    Dim sDiagnosis As Variant
    On Error GoTo ErrHandler
        
    If Not mbDirty Then
        Save = True
        Exit Function
    End If
    
    Save = GEMS_Save
    Exit Function

ErrHandler:
    Call mobjEncounter.CWShared.ShowVBError(Err.Number, Err.Description, "EnctrAP.Save ", "Enc", vbCritical)

End Function

Public Function GEMS_Save() As Boolean
' Assume section already locked for update before getting here
    Dim objLTR     As ListToolRecord
    Dim sDiagnosis As Variant
    Dim sql        As String
    Dim DataId     As String
    Dim oSQLEx       As ICHCSII_SQL.ISqlOpsEx
    Dim oSQL       As ICHCSII_SQL.ISqlOps
    Dim szSQL       As String
    Dim oSQLGen     As SQLGen
    Dim lStatus     As Long
    
    Set oSQL = gobjCHCSIIConn.CHCSII_SQLOPS(Auto)  '''<SCR 36493
   
    On Error GoTo ErrHandler
    
     'Preserve Status since CheckEncounterOwnership may refresh data and reset Status SCR#49641
     lStatus = meAPStatus
     
     'Check if encounter owner is Dr. Unassigned SCR-24349
     Call mobjEncounter.CheckEncounterOwnership
    
     meAPStatus = lStatus
   
    msOwnerNCID = mobjEncounter.Logon.UserNCID
    msOwnerName = mobjEncounter.Logon.UserName
    mdDTS = Now
    If meAPStatus = 0 Then meAPStatus = Incomplete

        Set oSQLEx = gobjCHCSIIConn.CHCSII_SQLOPS(Auto)
        
        If mnDataID > 0 Then
            '- update it
            oSQLEx.Execute "select * from enc_sections where dataid = " & mnDataID
            'Compress RTF SCR-24683
            Dim objZlib As EncZLib
            Dim aBytes() As Byte
            Set objZlib = New EncZLib
            
            aBytes = StrConv(Replace(msRTF, Chr(0), ""), vbFromUnicode)
            objZlib.CompressData aBytes
            With oSQLEx
                .value("Doc") = aBytes
            'oSQL.value("doc") = CompressRTF(msRTF)
                .value("originalsize") = Len(msRTF)
                .value("DTS") = mdDTS
                .value("OwnerNCID") = msOwnerNCID
                .value("OwnerName") = msOwnerName
                .value("Status") = StatusEnumToNCID(meAPStatus)
                .value("UPDATED_FLAG") = CStr(mbUpdated)
                .Update
            End With
                     
        Else
            '- insert it - this should not happen since all needed new ones are created on the load event
        End If
    
    '-- we don't store the dataid for list tool records. so go get it
    oSQL.Execute "select dataid from listToolRecord where encounternumber = " & mobjEncounter.EncounterID & " and facilityncid = " & mobjEncounter.FacilityNCID _
      & " and SourceNCID = " & SectionEnumToNCID(eap)
    If Not oSQL.EOF Then
        DataId = oSQL.value(0)
        'DataId = oSQL("dataid")
    Else
    '-- create the parent record in listtoolrecord table
       DataId = oSQL.GetNextID()
        oSQL.Execute "Insert into listToolRecord (dataid,EncounterNumber, FacilityNCID, ClinicNCID, SourceNCID, ClinicianNCID) values (" _
             & DataId & "," & mobjEncounter.EncounterID & "," & mobjEncounter.FacilityNCID & "," _
             & mobjEncounter.ClinicNCID & "," & SectionEnumToNCID(eap) & ", " & msOwnerNCID & ")"
    End If
    
    oSQL.Execute "delete from listToolRecords where DataID = " & DataId

    oSQL.Execute ("SELECT * FROM ListToolRecords WHERE DataID = -1")
    'oSQL.Execute "listToolRecords" 'This method doesn't work for SQL Server SCR-609 ITT
    
    For Each objLTR In mcolSnoIDs
        
        If Val(objLTR.SnoID) > 0 Then 'SCR-40393
            Set oSQLGen = New SQLGen
            oSQLGen.SetTransType = eSQLInsert
        
            oSQLGen.AddToSQL "DataID", DataId, eSQLNumber
            oSQLGen.AddToSQL "updateflag", "N", eSQLString
            oSQLGen.AddToSQL "SnoID", objLTR.SnoID, eSQLString
            oSQLGen.AddToSQL "Duration", objLTR.Duration, eSQLString
            oSQLGen.AddToSQL "Modifier", objLTR.Modifier, eSQLString
            oSQLGen.AddToSQL "Onset", objLTR.Onset, eSQLString
            oSQLGen.AddToSQL "prefix", objLTR.prefix, eSQLString
            oSQLGen.AddToSQL "qlink", objLTR.QualifierLink, eSQLString
            oSQLGen.AddToSQL "qlinkseq", objLTR.QualifierLinkSequence, eSQLNumber
            oSQLGen.AddToSQL "RangeNormalHigh", objLTR.RangeNormalHigh, eSQLNumber
            oSQLGen.AddToSQL "RangeNormalLow", objLTR.RangeNormalLow, eSQLNumber
            oSQLGen.AddToSQL "RangeScale", objLTR.RangeScale, eSQLNumber
            oSQLGen.AddToSQL "Referid", objLTR.Referid, eSQLNumber
            oSQLGen.AddToSQL "result", objLTR.result, eSQLString
            oSQLGen.AddToSQL "Status", objLTR.Status, eSQLString
            oSQLGen.AddToSQL "Unit", objLTR.Unit, eSQLString
            oSQLGen.AddToSQL "Value", objLTR.value, eSQLString
            
            szSQL = oSQLGen.GenerateSQL("ListToolRecords", "")
            oSQL.Execute szSQL
            
            Set oSQLGen = Nothing
        End If
        
    Next objLTR
    
    GEMS_Save = True
    Set oSQL = Nothing
    Set oSQLEx = Nothing
    mbDirty = False

    Exit Function
    
ErrHandler:
    Call mobjEncounter.CWShared.ShowVBError(Err.Number, Err.Description, "EnctrGems.Save ", "Enc", vbCritical)
    Exit Function
    Resume
End Function
Public Sub NotifyComplete()
    Dim bCreatedObject As Boolean
    On Error GoTo ErrHnd
    
    'On Error Resume Next
    
    'If mobjEncounter.CWShared.IsAppMode(modeCHCSI_GUI) Then Exit Sub 'SCR-40055
    
    If mobjModule Is Nothing Then
        Set mobjModule = mobjEncounter.Comm.InitializeOLEServer(SectionEnumToNCID(eap))
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
    mobjModule.InitializeAnp True
    Call mobjModule.CompletedEncounter
    Call mobjIChild.SetParent(Nothing)
    
    If bCreatedObject Then
        Set mobjModule = Nothing
        Set mobjIChild = Nothing
    End If

Exit Sub
ErrHnd:
    MsgBox "Application error: " & Err.Description & " Occurred in CHCSIIEncounterOps.AP.NotifyComplete."
    
    
End Sub


