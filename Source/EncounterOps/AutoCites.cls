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
Attribute VB_Name = "AutoCites"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mobjEncounter As Encounter
Private mobjParent As SectionParent
Private msRTF As String
Private msLabRTF As String
Private msRadRTF As String
Private msQuestionnaireRTF As String
Private mbDirty As Boolean
Private mnDataID As Long
Private mbOpening As Boolean
Private msOwnerNCID As String
Private msOwnerName As String
Private mdDTS As Date
Private oSQL As ICHCSII_SQL.ISqlOpsEx
Private msTableRTF As String

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
    'msRTF = ""
    ' mobjEncounter may not be set SCR-31581
    If Not mobjEncounter Is Nothing Then
        If mobjEncounter.CWShared.IsAppMode(modeCHCSI_GUI) Then '<SCR SCR 36493
            If mdDTS = 0 Then
                mdDTS = Now
            End If
            'Use user instead of PP SCR-35537
            RTF = FormatRtfDisplay(msRTF, "AutoCites", "Refreshed", mobjEncounter.Logon.UserName, mdDTS)
        End If
    End If
    
    If Len(RTF) = 0 Then
        RTF = FormatRtfDisplay(msRTF, "AutoCites", "Refreshed", msOwnerName, mdDTS)
    End If
    
End Property

Public Property Get LabRTF() As String
    If Not msLabRTF = "" Then
        LabRTF = FormatRtfDisplay(msLabRTF, "Lab AutoCites", "Refreshed", msOwnerName, mdDTS)
    End If
End Property

Public Property Get RadRTF() As String
    If Not msRadRTF = "" Then
        RadRTF = FormatRtfDisplay(msRadRTF, "Rad AutoCites", "Refreshed", msOwnerName, mdDTS)
    End If
End Property
Public Property Get QuestionnaireRTF() As String
    If Not msQuestionnaireRTF = "" Then
       QuestionnaireRTF = FormatRtfDisplay(msQuestionnaireRTF, "Questionnaire AutoCites", "Refreshed", msOwnerName, mdDTS)
    End If
End Property

Public Sub Load(ByRef objParent As SectionParent, ByRef colDataStrings As Collection)
    Dim objData As DATASTRATLLib.DataString
    Dim objEvent As CLINTYPEATLLib.ClinicalEvent
    Dim i As Long
    Dim oShared As New CWShared
    
    If oShared.IsAppMode(modeCHCSI_GUI) Then '<SCR SCR 36493
        Exit Sub
    End If
'
    If Not mobjEncounter Is Nothing Then
        If Not mobjEncounter.CWShared.IsAppMode(modeCHCSI_GUI) Then '<SCR SCR 36493
            msRTF = ""
        End If
    End If
    mbDirty = False
    mnDataID = 0
    'Using the For...Each method with 3M objects causes memory leak
    For i = 1 To colDataStrings.Count
    'For Each objData In colDataStrings
        Set objData = colDataStrings(i)
        If objData.DataTypeNCID = NCID_ENCEVENT Then
            Set objEvent = objData.PatientData
            If objEvent.Header.ObsBatId.value.Coded.ValueConcept.UniqueId = SectionEnumToNCID(eAutoCite) Then
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
                    If objEvent.ClinObs.Count > 1 Then
                        msLabRTF = ExtractRTF(objEvent.ClinObs.Item(2))
                        If objEvent.ClinObs.Count > 2 Then
                            msRadRTF = ExtractRTF(objEvent.ClinObs.Item(3))
                        End If
                        If objEvent.ClinObs.Count > 3 Then
                            msQuestionnaireRTF = ExtractRTF(objEvent.ClinObs.Item(4))
                        End If
                    End If
                End If
                Exit For
            End If
        End If
    Next 'objData
End Sub

Public Sub OpenEncounter(ByRef objParent As SectionParent, ByRef objEncounter As Encounter)
    
On Error GoTo ErrHandler
    
    mbOpening = True
    Set mobjParent = objParent
    Set mobjEncounter = objEncounter
    'Refresh Autocites whenever any user opens Encounter SCR-37839
    'Refresh Autocites whenever PP opens Encounter SCR-34292
    'Refresh autocite for newly opened encounter by primary provider SCR# 57244, 57245, 57246, 57247
    If mobjEncounter.UserIsPrimaryProvider And msRTF = "" Then
        'Dont refresh if encounter in one of these statuses
        'Occurs when amending an encounter SCR-34372
        'added NeedsCosignature SCR#50616
        If mobjEncounter.Status = complete _
        Or mobjEncounter.Status = PatientLeftWithoutBeingSeen _
        Or mobjEncounter.Status = CancelByFacility _
        Or mobjEncounter.Status = CancelByPatient _
        Or mobjEncounter.Status = CancelByProvider _
        Or mobjEncounter.Status = NoShow _
        Or mobjEncounter.Status = Updated _
        Or mobjEncounter.Status = NeedsCoSignature Then
            Exit Sub
        End If
        Call Refresh
    End If
    mbOpening = False
    
    Exit Sub
    
ErrHandler:

    MsgBxARMd "EncounterOps " & Err & " Autocites Open Encounter"
End Sub

Public Sub CloseEncounter()
    Set mobjParent = Nothing
    Set mobjEncounter = Nothing
End Sub

Public Function OpenSection() As Boolean
    'Autocite turned on for all users SCR-10731
    'If mobjEncounter.UserIsPrimaryProvider Then
    Screen.MousePointer = vbHourglass 'SCR-23678
    Call Refresh
    Screen.MousePointer = vbHourglass
    Call mobjEncounter.RefreshDisplay
    OpenSection = True
    'End If
End Function

Public Sub Refresh()
    
    Dim objComm As MMMHISComm.Comm
    Dim colAC As Collection
    Dim objAC As AutoCiteProp
    Dim sRTFAC As String
    Dim bOwnLock As Boolean
    Dim tEncounterInformation As Encounter_Key_Information
    Dim bAddTable As Boolean
    
    On Error GoTo ErrHandler
    
    LoadTableRTF

    Screen.MousePointer = vbHourglass 'SCR-23678
    If mobjEncounter Is Nothing Then
        GoTo CleanUp
    End If

    If Not mobjEncounter.LockingObject.SectionIsLocked(eAutoCite, 0) Then
        bOwnLock = mobjEncounter.LockingObject.LockSection(eAutoCite, 0, SectionLock)
        If Not bOwnLock Then
            GoTo CleanUp
        End If
    End If
    'PERF - Do not refresh for newly opened encounters SCR-31509
    If Len(msRTF) Then
        mobjEncounter.Refresh
    End If
    If Not mobjEncounter.SectionIsEditable(eAutoCite) Then
        MsgBxARMd "This encounter has been saved by another user and you can no longer refresh the AutoCite section.", vbInformation, "Encounter"
        GoTo CleanUp
    End If
    If mobjEncounter.Status = Updating Then
        If Not mobjEncounter.AddHistoryItem(Overwritten, "AutoCites", RTF) Then
            MsgBxARMd "Unable to copy section to change history.", vbCritical, "Encounter"
            GoTo CleanUp
        End If
        mbDirty = True
        Call Save
    End If

    mobjEncounter.Comm.SetStatusText "Refreshing Encounter " & mobjEncounter.EncounterID & " " & SectionEnumToName(mobjParent.meSection) & " section."
    msRTF = ""
    msLabRTF = ""
    msRadRTF = ""
    msQuestionnaireRTF = ""
    
    Set objComm = mobjEncounter.EncounterParent.CWComm
    
    If mobjEncounter.TemplateData = "" Then
       Set colAC = mobjEncounter.EncounterParent.EncounterOps.EncProperties.AutoCites
    Else
       With mobjEncounter.EncounterParent.EncounterOps.EncProperties
              'In PG just parse the template data SCR-35162
              If Not mobjEncounter.CWShared.IsAppMode(modeCHCSI_GUI) Then '<SCR SCR 36493
                tEncounterInformation.ID = CStr(mobjEncounter.EncounterID)
                tEncounterInformation.Facility = mobjEncounter.FacilityNCID
                tEncounterInformation.Unit_NUmber = mobjEncounter.PatientID
                .RetrieveTemplateProperties colAC, tEncounterInformation
              End If
              If Not .TemplatePropertiesExist Then Set colAC = .ParseTemplateData(mobjEncounter.TemplateData)
       End With
     End If
    
    For Each objAC In colAC
    
        Screen.MousePointer = vbHourglass 'SCR-23678
        
        If objAC.Selected Then
            
            sRTFAC = GetSectionAutociteRTF(objAC)
                    
            AddSectionAutociteRTF objAC.AppNCID, objAC.AppName, sRTFAC, bAddTable
                     
            sRTFAC = vbNullString 'Reset Autocite RTF String SCR-467 ITT

        End If
    Next objAC
    If Len(msRTF) = 0 Then
        msRTF = " "
    End If
    
    If Len(objComm.CmdLineSwitch("GREEN")) Then
        If bAddTable Then
            'Replace placeholder with empty string in case they were no autocited SCR-28658
            msTableRTF = Replace(msTableRTF, "<GOALSUMMARY_RTF>", vbNullString) '"{\pard\plain\f0\fs16\b }")
            msTableRTF = Replace(msTableRTF, "<FAMILYHX_RTF>", " \par\par ")
            msTableRTF = Replace(msTableRTF, "<ALLERGIES_RTF>", vbNullString)
            msTableRTF = Replace(msTableRTF, "<PROBLEMS_RTF>", vbNullString)
            'Add Table RTF to RTF string only if sections in table were autocited SCR-28658
            msRTF = msTableRTF & msRTF
        End If
    End If
   
   If mobjEncounter.CWShared.IsAppMode(modeCHCSI_GUI) And Not mobjEncounter.CWShared.IsAppMode(modeITT) And Len(objComm.CmdLineSwitch("HIPAA837")) > 0 Then
      'SCR 46932 - Display Pregnancy Data in AutoCites
      With mobjEncounter.PregnancyData
         'SCR 46933 - don't display if patient isn't indicated as pregnant
         If .PregnancyIndicator Then
            msRTF = msRTF & "{\pard\plain\f0\fs16\b Pregnant.\b0 " & _
                  IIf(.LastMenstrualPeriod <> NULL_DATE, "  LMP: " & Format(.LastMenstrualPeriod, gsDefaultDate) & ".", vbNullString) & _
                  IIf(.EstimatedDOB <> NULL_DATE, "  EDC: " & Format(.EstimatedDOB, gsDefaultDate) & ".", vbNullString) & "\par }"
         End If
      End With
      msRTF = msRTF & mobjEncounter.AccidentData.RTF(mobjEncounter.FacilityNCID)
   End If
   
    mbDirty = True
    
    Call Save(True) 'SCR-26446
    
CleanUp:
    If bOwnLock Then
        Call mobjEncounter.LockingObject.UnlockSection(eAutoCite, 0)
    End If
    Set objComm = Nothing
    Set colAC = Nothing
    Set objAC = Nothing
    mobjEncounter.Comm.SetStatusText ""
    
    Exit Sub
    
ErrHandler:
   Call mobjEncounter.CWShared.ShowVBError(Err.Number, Err.Description, "Autocites.Refresh ", "Encounters", vbCritical)
    Exit Sub
    Resume
End Sub
Friend Function Save(Optional bLocked As Boolean = False) As Boolean
' Assume section already locked before getting here, but that we need to do the start/end update
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

Public Sub GEMS_Load(ByRef objParent As SectionParent, objRS As ADODB.Recordset)
    Dim oConcept As New GEMS_ConceptCtrl
    Dim lsize       As Long
    Dim loffset     As Long
    Dim varChunk    As Variant
    Dim Index       As Integer
    Dim ChunkSize   As Long
    Dim lResult     As Long
    Dim sTxt        As String
    Dim vText As Variant
    Dim dDTS As Date
    Dim sOwnerNCID As String
    Dim oShared     As CWShared
    
    Set oShared = New CWShared
      
    
    'msRTF = ""
    mbDirty = False
    mnDataID = 0
    On Error Resume Next   '- for the date fields below
    If Not objRS.EOF Then
        mnDataID = Val("" & objRS("DataID"))
        sOwnerNCID = "" & objRS("ownerncid")
        If Not msOwnerNCID = "" Then
            oConcept.UniqueId = sOwnerNCID
            msOwnerName = oConcept.PrefRep("2000").Representation
        End If
        'Fix SCR 18679
        If Not IsDate(objRS("DTS")) Then
            dDTS = Now
        Else
            dDTS = objRS("DTS")
        End If
        'msRTF = "" & objRS.("Doc")
      
        msLabRTF = "" & objRS("LabRTF")
        msRadRTF = "" & objRS("RadRTF")
        
        If oShared.IsAppMode(modeCHCSI_GUI) Then
            If Len(sOwnerNCID) And sOwnerNCID <> "0" Then
                'Check if data has been updated SCR#43993
                If Not (dDTS = mdDTS And StrComp(sOwnerNCID, msOwnerNCID, vbTextCompare) = 0) Then
                    'msRTF = oShared.CHCSConnection.ReadChunk("ENC_SECTIONS", "ENCOUNTERNUMBER = " & objRS("ENCOUNTERNUMBER") & " AND enc_sectionsindex = " & SectionEnumToNCID(eAutoCite), "DOC")
                    msRTF = ReadChunkCHCSConn("ENC_SECTIONS", "ENCOUNTERNUMBER = " & objRS("ENCOUNTERNUMBER") & " AND enc_sectionsindex = " & SectionEnumToNCID(eAutoCite), "DOC", False)
                End If
            End If
        Else
            'Decompress if compressed SCR-24683
            If CVar(objRS.Fields("OriginalSize")) > 0 Then
                Dim objZlib As EncZLib
                Set objZlib = New EncZLib
                Dim aBytes() As Byte
                
                aBytes = objRS.Fields("Doc")
                Call objZlib.DecompressData(aBytes, objRS.Fields("OriginalSize"))
                msRTF = StrConv(aBytes, vbUnicode)
                'msRTF = DecompressRTF("" & objRS.Fields("Doc"), objRS.Fields("OriginalSize"))
            Else
    
    '            '********************************
    '            'Use GetChunk to get Long data over 4K 'SCR-18343
    '            lsize = objRS.Fields("Doc").ActualSize
    '
    '            loffset = 0 ' Reset offset
    '            vText = ""
    '
    '            '-- set chunksize
    '            ChunkSize = 2000
    '            If lsize > 10000 Then
    '                ChunkSize = lsize \ 10
    '                '- Make it EVEN (2-bytes per char, remember?)
    '                If ChunkSize \ 2 <> ChunkSize / 2 Then ChunkSize = ChunkSize + 1
    '                If ChunkSize > 20000 Then ChunkSize = 20000
    '            End If
    '
    '            Do While loffset < lsize
    '               varChunk = objRS.Fields("Doc").GetChunk(ChunkSize)
    '               vText = vText & varChunk
    '               loffset = loffset + ChunkSize
    '            Loop
    '
    '            msRTF = CStr(vText)
                ' END CODE CHANGE 'SCR-18343
                '*************************************
            End If
        End If
        msOwnerNCID = sOwnerNCID
        mdDTS = dDTS
        
    Else
        mdDTS = Now
    End If
    Set oConcept = Nothing
                    
End Sub

' Use ICHCSII_SQL to do inserts and updates to utilize AppendChunk functionality
' to save over 4K of data in type Long column SCR-18343
Friend Function GEMS_Save() As Boolean
    Dim mobjDAS As ICHCSII_DAS
    Dim lRecCnt As Long     '- number of records affected
    Dim msql As ICHCSII_SQL.ISqlOpsEx
    'Set msql = GetSQLOps
    Dim oConcept As GEMS_ConceptCtrl

    Dim oDas As ICHCSII_DAS
    
    Set oDas = gobjCHCSIIConn.CHCSII_DAS(Auto) '''<SCR 36493
    Set msql = gobjCHCSIIConn.CHCSII_SQLOPS(Auto)

' Assume section already locked before getting here, but that we need to do the start/end update
    On Error GoTo ErrHandler
    
    If Not mbDirty Then
        GEMS_Save = True
        Exit Function
    End If
    Dim objDAS As ICHCSII_DAS
    'Set objDAS = gobjCHCSIIConn.CHCSII_DAS(Auto) '''<SCR 36493
    Dim sql As String
    
    msOwnerNCID = mobjEncounter.Logon.UserNCID
    
    'jrm:  Setting to new before using. Dim as .. as new .. was failing.
    Set oConcept = New GEMS_ConceptCtrl
    
    'Set ownername scr-23851
    If Not msOwnerNCID = "" Then
        oConcept.UniqueId = msOwnerNCID
        msOwnerName = oConcept.PrefRep("2000").Representation
    End If
    
    mdDTS = Now
    
    If Not mobjEncounter.LockingObject.StartSectionUpdate(mobjParent.meSection, 0) Then
        MsgBxARMd "Your lock on the " & SectionEnumToName(mobjParent.meSection) & " section was broken by another user, and your changes cannot be saved.", vbInformation, "Encounter"
        GoTo CleanUp
    End If
        
    If mnDataID > 0 Then
        Call msql.Execute("Select * from enc_sections where dataid = " & mnDataID)
        If msql.EOF Then   '-- didn't find it-- make a new one

            mnDataID = msql.GetNextID

            'SF 84601 Added where clause to not open the entire table.
            Call msql.Execute("select * from enc_sections where dataid=0")
            msql.AddNew
            msql("DataID") = mnDataID
        End If
    Else
    
        mnDataID = msql.GetNextID

        Call msql.Execute("SELECT DATAID, FACILITYNCID, ENC_SECTIONSINDEX," _
            & " ENCOUNTERNUMBER, DTS, DOC," _
            & " OWNERNCID, OWNERNAME, SENSITIVITYLEVEL," _
            & " STATUS, ORIGINALSIZE, LABRTF," _
            & " RADRTF, CREATEDBY, CREATEDON," _
            & " UPDATEDBY, UPDATEDON, COMPLETE," _
            & " CATEGORY, TITLE, WSRFV," _
            & " DATATYPENCID, UPDATED_FLAG FROM ENC_SECTIONS" _
            & " where dataid=0")

        msql.AddNew
        msql("DataID") = mnDataID
        
    End If
    
    msql("FacilityNCID") = mobjEncounter.FacilityNCID
    msql("EncounterNumber") = mobjEncounter.EncounterID
    msql("enc_sectionsindex") = SectionEnumToNCID(eAutoCite)
    msql("ownerncid") = mobjEncounter.Logon.UserNCID 'scr-23851 REMOVE sqlquote
    msql("DTS") = mdDTS 'SQLDate(mdDTS) SCR-24331

    'Compress RTF SCR-24683
    Dim objZlib As EncZLib
    Dim aBytes() As Byte
    Set objZlib = New EncZLib
            
    If Len(msRTF) <> 0 Then
        aBytes = StrConv(Replace(msRTF, Chr(0), ""), vbFromUnicode)
        objZlib.CompressData aBytes
        msql("Doc") = aBytes
    Else
        'Fixed SCR #34615 - Jane Shen 4/18/2003
        'Fix Also applied to Autocites SCR-36826
        msql("Doc") = ""
    End If

    'msql.AppendAsChunk "Doc", msRTF
    msql("OriginalSize") = Len(msRTF)

    'Dont need sqlquote here 'SCR-737, 739
    msql("LabRTF") = msLabRTF
    msql("RadRTF") = msRadRTF
    msql.Update
            
    mbDirty = False
    GoTo CleanUp
ErrHandler:
   Call mobjEncounter.CWShared.ShowVBError(Err.Number, Err.Description, "Enctr.GEMS_Save ", "Enc", vbCritical)
CleanUp:
    Call mobjEncounter.LockingObject.EndSectionUpdate(mobjParent.meSection, 0)
    
    If GEMS_Save And Not mbOpening Then
        mobjEncounter.Comm.SetStatusText "Encounter " & mobjEncounter.EncounterID & " " & SectionEnumToName(mobjParent.meSection) & " section was saved."
    Else
        mobjEncounter.Comm.SetStatusText ""
    End If
    Set objDAS = Nothing
    Set oConcept = Nothing
End Function


Private Function GetSQLOps() As ICHCSII_SQL.ISqlOpsEx
    If oSQL Is Nothing Then
        Set oSQL = gobjCHCSIIConn.CHCSII_SQLOPS(Auto)  '''<SCR 36493
    End If
    Set GetSQLOps = oSQL
End Function

Private Function GetSectionAutociteRTF(ByVal objAC As AutoCiteProp) As String

    Dim sRTFAC As String
    Dim objModule As IChild
    Dim objModule2 As IChild2
    Dim objComm As MMMHISComm.Comm
    
    On Error GoTo ErrHandler
    
    Set objComm = mobjEncounter.EncounterParent.CWComm
    
    Select Case objAC.AppNCID
    Case NCID_FAMILY_HISTORY_APPLICATION, NCID_PROBLEMS_APPLICATION
        'Problems implement IChild2 for FamilyHx
        Set objModule2 = objComm.InitializeOLEServer(NCID_PROBLEMS_APPLICATION)
        
        If Not objModule2 Is Nothing Then
            sRTFAC = objModule2.GetAutoCiteNarrative("", 0, 0, CLng(objAC.AppNCID), 0, 0, 0)
        End If
    Case NCID_EXPIRED_MEDICATIONS 'SCR-20884
        Set objModule = objComm.InitializeOLEServer(NCID_MEDICATIONS_APPLICATION)
        
        If Not objModule Is Nothing Then
            sRTFAC = objModule.GetAutoCiteNarrative("EXPIRED", objAC.NumberPeriods, objAC.Period)
        End If
        
    'SCR 44005
    Case NCID_INPATIENT_MEDICATIONS
        
        If mobjEncounter.Category Like APPT_TYPE_RNDS Then 'SCR#45422
            Set objModule = objComm.InitializeOLEServer(NCID_MEDICATIONS_APPLICATION)
            
            If Not objModule Is Nothing Then
                sRTFAC = objModule.GetAutoCiteNarrative("INP", objAC.NumberPeriods, objAC.Period)
            End If
        End If
      
    Case Else
        Set objModule = objComm.InitializeOLEServer(objAC.AppNCID)
        
        If Not objModule Is Nothing Then
            sRTFAC = objModule.GetAutoCiteNarrative("", objAC.NumberPeriods, objAC.Period)
        End If
    End Select
    
    If Len(sRTFAC) = 0 Then
            sRTFAC = "No " & objAC.AppName & " Found."
    End If
    
    '9/12/2001 WRogers this line is problematic when no error handler is used
    'If Err Or (objModule Is Nothing And objModule2 Is Nothing) Then
    'SCR 16497
    If (objModule Is Nothing And objModule2 Is Nothing) Then
        sRTFAC = "Could not retrieve autocites for " & objAC.AppName
    End If
    
    GetSectionAutociteRTF = sRTFAC

    Set objModule = Nothing
    Set objModule2 = Nothing
    Set objComm = Nothing
    
    Exit Function
    
ErrHandler:
   Call mobjEncounter.CWShared.ShowVBError(Err.Number, Err.Description, "Enctr.GetSectionAutociteRTF ", "Enc", vbCritical)
    
End Function

Private Sub LoadTableRTF()

    On Error GoTo ErrHandler
    
    'Define table SCR-28658
    'Table RTF with <PROBLEMS_RTF>, <FAMILYHX_RTF>, <ALLERGIES_RTF> and <GOALSUMMARY_RTF> as placeholders to be replaceed by autocite section RTFS
    msTableRTF = "\trowd \trgaph115\trleft-115\trkeep\tpvpara\tposy0\tdfrmtxtLeft180\tdfrmtxtRight180\trftsWidth1\trftsWidthB3\trftsWidthA3\trpaddl115\trpaddr115\trpaddfl3\trpaddfr3\tabsnoovrlp1 \clvertalt\clbrdrt\brdrnone \clbrdrl\brdrnone \clbrdrb\brdrnone \clbrdrr" _
    & "\brdrnone \cltxlrtb\clftsWidth3\clwWidth3192 \cellx3077\clvertalt\clbrdrt\brdrnone \clbrdrl\brdrnone \clbrdrb\brdrnone \clbrdrr\brdrnone \cltxlrtb\clftsWidth3\clwWidth3192 \cellx6269\clvertalt\clbrdrt\brdrnone \clbrdrl\brdrnone \clbrdrb\brdrnone \clbrdrr" _
    & "\brdrnone \cltxlrtb\clftsWidth3\clwWidth3192 \cellx9461\pard\plain \s21\ql \fi-240\li240\ri0\widctlpar\intbl\aspalpha\aspnum\faauto\adjustright\rin0\lin240 \fs24\lang1033\langfe1033\cgrid\langnp1033\langfenp1033 {<PROBLEMS_RTF>\cell <FAMILYHX_RTF>" _
    & "}\pard\plain \ql \li0\ri0\widctlpar\intbl\aspalpha\aspnum\faauto\adjustright\rin0\lin0 \fs24\lang1033\langfe1033\cgrid\langnp1033\langfenp1033 {<GOALSUMMARY_RTF>\cell }\pard\plain" _
    & "\s21\ql \fi-240\li240\ri0\widctlpar\intbl\aspalpha\aspnum\faauto\adjustright\rin0\lin240 \fs24\lang1033\langfe1033\cgrid\langnp1033\langfenp1033 {<ALLERGIES_RTF>\cell }\pard\plain \ql \li0\ri0\widctlpar\intbl\aspalpha\aspnum\faauto\adjustright\rin0\lin0" _
    & "\fs24\lang1033\langfe1033\cgrid\langnp1033\langfenp1033 {\trowd \trgaph115\trleft-115\trkeep\tpvpara\tposy0\tdfrmtxtLeft180\tdfrmtxtRight180\trftsWidth1\trftsWidthB3\trftsWidthA3\trpaddl115\trpaddr115\trpaddfl3\trpaddfr3\tabsnoovrlp1 \clvertalt\clbrdrt" _
    & "\brdrnone \clbrdrl\brdrnone \clbrdrb\brdrnone \clbrdrr\brdrnone \cltxlrtb\clftsWidth3\clwWidth3192 \cellx3077\clvertalt\clbrdrt\brdrnone \clbrdrl\brdrnone \clbrdrb\brdrnone \clbrdrr\brdrnone \cltxlrtb\clftsWidth3\clwWidth3192 \cellx6269\clvertalt\clbrdrt" _
    & "\brdrnone \clbrdrl\brdrnone \clbrdrb\brdrnone \clbrdrr\brdrnone \cltxlrtb\clftsWidth3\clwWidth3192 \cellx9461\row }\pard\plain \s20\ql \li0\ri0\widctlpar\aspalpha\aspnum\faauto\adjustright\rin0\lin0\itap0" _
    & "\fs24\lang1033\langfe1033\cgrid\langnp1033\langfenp1033"

    Exit Sub
    
ErrHandler:
   Call mobjEncounter.CWShared.ShowVBError(Err.Number, Err.Description, "Enctr.LoadTableRTF ", "Enc", vbCritical)

End Sub

Private Sub AddSectionAutociteRTF(ByVal AppNCID As String, ByVal AppName As String, ByVal sRTFAC As String, ByRef bAddTable As Boolean)

    On Error GoTo ErrHandler
    
   ' If Not mobjEncounter.CWShared.IsAppMode(modeCHCSI_GUI) Then '<SCR SCR 36493
    If Len(mobjEncounter.CWShared.CmdLineSwitches("GREEN")) Then
        Select Case AppNCID
            Case SectionEnumToNCID(eLabs)
                msLabRTF = "{\pard\plain\f0\fs16\b " & AppName & "\par\b0 " & sRTFAC & "\par }"
            Case SectionEnumToNCID(eRads)
                msRadRTF = "{\pard\plain\f0\fs16\b " & AppName & "\par\b0 " & sRTFAC & "\par }"
            Case SectionEnumToNCID(eQuestionnaire)
                msQuestionnaireRTF = "{\pard\plain\f0\fs16\b " & AppName & "\par\b0 " & sRTFAC & "\par }"
            'Place these RTFs in Table, replacing placeholders SCR-28658
            Case NCID_ALLERGIES_APPLICATION '\pard\plain \s2\fi-720\li720\fs16\ri2880
                msTableRTF = Replace(msTableRTF, "<ALLERGIES_RTF>", "{\pard\plain\f0\fs16\b " & AppName & "\par\b0 \pard\plain \s2\fi-720\li320\fs16 " & sRTFAC & "}")
                bAddTable = True
            Case NCID_PROBLEMS_APPLICATION
                msTableRTF = Replace(msTableRTF, "<PROBLEMS_RTF>", "{\pard\plain\f0\fs16\b " & AppName & "\par\b0 \pard\plain \s2\fi-720\li320\fs16 " & sRTFAC & "}")
                bAddTable = True
            Case NCID_FAMILY_HISTORY_APPLICATION
                msTableRTF = Replace(msTableRTF, "<FAMILYHX_RTF>", "{\pard\plain\f0\fs16\b " & AppName & "\par\b0 \pard\plain \s2\fi-720\li320\fs16 " & sRTFAC & "}")
                bAddTable = True
            Case NCID_PATIENT_PLANOFCARE
                msTableRTF = Replace(msTableRTF, "<GOALSUMMARY_RTF>", "{\pard\plain\f0\fs16\b " & AppName & "\par\b0 \pard\plain \s2\fi-720\li320\fs16 " & sRTFAC & "}")
                bAddTable = True
            'Coupler RTF gets placed in Add Note section
            Case NCID_HEAR_APPLICATION
                mobjEncounter.AddTextNote "Coupler Information", msRTF
            Case NCID_EXPIRED_MEDICATIONS
                'Do not display expired meds for RNDS encounters SCR#44005
                If Not (mobjEncounter.Category Like APPT_TYPE_RNDS) Then
                    msRTF = msRTF & "{\pard\plain\f0\fs16\b " & AppName & "\par\b0 " & sRTFAC & "\par }"
                End If
            Case NCID_INPATIENT_MEDICATIONS
                If mobjEncounter.Category Like APPT_TYPE_RNDS Then 'SCR#45422, SCR#51918
                    msRTF = msRTF & "{\pard\plain\f0\fs16\b " & AppName & "\par\b0 " & sRTFAC & "\par }"
                End If
            Case Else
                msRTF = msRTF & "{\pard\plain\f0\fs16\b " & AppName & "\par\b0 " & sRTFAC & "\par }"
        End Select
    Else
        Select Case AppNCID
            Case SectionEnumToNCID(eLabs)
                msLabRTF = "{\pard\plain\f0\fs16\b " & AppName & "\par\b0 " & sRTFAC & "\par }"
            Case SectionEnumToNCID(eRads)
                msRadRTF = "{\pard\plain\f0\fs16\b " & AppName & "\par\b0 " & sRTFAC & "\par }"
            Case SectionEnumToNCID(eQuestionnaire)
                msQuestionnaireRTF = "{\pard\plain\f0\fs16\b " & AppName & "\par\b0 " & sRTFAC & "\par }"
            'Coupler RTF gets placed in Add Note section
            Case NCID_HEAR_APPLICATION
                mobjEncounter.AddTextNote "Coupler Information", msRTF
            Case Else
             ' SCR 106643 9/29/06 Quan SCR 107415 1/2/07 Quan
                If AppName = "Vitals" And sRTFAC <> "No Vitals Found." Then
                    msRTF = msRTF & "{\pard\plain\f0\fs16\b\par\b0 " & sRTFAC & "\par }"
                Else
                    msRTF = msRTF & "{\pard\plain\f0\fs16\b " & AppName & "\par\b0 " & sRTFAC & "\par }"
                End If
        End Select
    End If
'     Else
'         'Apply to PG SCR-35747
'         If Len(objComm.CmdLineSwitch("GREEN")) Then
'             Select Case AppNCID
'                 'Place these RTFs in Table, replacing placeholders SCR-28658
'                 Case NCID_ALLERGIES_APPLICATION '\pard\plain \s2\fi-720\li720\fs16\ri2880
'                     msTableRTF = Replace(msTableRTF, "<ALLERGIES_RTF>", "{\pard\plain\f0\fs16\b " & AppName & "\par\b0 \pard\plain \s2\fi-720\li320\fs16 " & sRTFAC & "}")
'                     bAddTable = True
'                 Case NCID_PROBLEMS_APPLICATION
'                     msTableRTF = Replace(msTableRTF, "<PROBLEMS_RTF>", "{\pard\plain\f0\fs16\b " & AppName & "\par\b0 \pard\plain \s2\fi-720\li320\fs16 " & sRTFAC & "}")
'                     bAddTable = True
'                 Case NCID_FAMILY_HISTORY_APPLICATION
'                     msTableRTF = Replace(msTableRTF, "<FAMILYHX_RTF>", "{\pard\plain\f0\fs16\b " & AppName & "\par\b0 \pard\plain \s2\fi-720\li320\fs16 " & sRTFAC & "}")
'                     bAddTable = True
'                 Case NCID_PATIENT_PLANOFCARE
'                     msTableRTF = Replace(msTableRTF, "<GOALSUMMARY_RTF>", "{\pard\plain\f0\fs16\b " & AppName & "\par\b0 \pard\plain \s2\fi-720\li320\fs16 " & sRTFAC & "}")
'                     bAddTable = True
'                 Case NCID_INPATIENT_MEDICATIONS
'                     If mobjEncounter.Category Like APPT_TYPE_RNDS Then 'SCR#45422
'                         msRTF = msRTF & "{\pard\plain\f0\fs16\b " & AppName & "\par\b0 " & sRTFAC & "\par }"
'                     End If
'                 Case Else
'                     msRTF = msRTF & "{\pard\plain\f0\fs16\b " & AppName & "\par\b0 " & sRTFAC & "\par }"
'             End Select
'         Else
'             'In Medcin Lite mode all autocite will be displayed in Autocites section since SO section does not exist
'             msRTF = msRTF & "{\pard\plain\f0\fs16\b " & AppName & "\par\b0 " & sRTFAC & "\par }"
'         End If
'
'     End If
       
    Exit Sub
    
ErrHandler:
   Call mobjEncounter.CWShared.ShowVBError(Err.Number, Err.Description, "Enctr.AddSectionAutociteRTF ", "Enc", vbCritical)
    
End Sub
