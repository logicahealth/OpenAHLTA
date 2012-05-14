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
Attribute VB_Name = "Disposition"
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

Private msDentalDispStatus As String
Private msDentalDispStatusDesc As String
Private msDentalDispComments As String
Private msDentalDispOptions As String

Private meDispStatus As EncStatus
Private msDispositionText As String
Private msDispositionNCID As String
Private msDispAdminOption As String
Private msEandMNCID As String
Private msEandMCalc As String
Private msEMCalcOverrideNote As String
Private msWSDisp As String
Private msFollowupTimeFrame As String
Private msFollowupComments As String
Private msDiscussedComments As String
Private msItemsDiscussed As String
Private msDBNICategory As String
Private msDBNICause As String

'Extra properties for E & M Code Review
Private mbEandMReviewed     As Boolean
Private msEandMReviewedBy   As String
Private mdtEandMReviewDate  As Date
Private Const EMPTY_CODE As String = "0"
Const FLD_METADATA = "DispMetaData"
Private m_szDispMetadata As String
Const NCID_METADATA = "4161962"
Public Property Get DTS() As Date
    DTS = mdDTS
End Property
Public Property Get DispMetadata() As String
    DispMetadata = m_szDispMetadata
End Property

Public Property Let DispMetadata(ByVal vNewValue As String)
    If m_szDispMetadata <> vNewValue Then
        Me.Dirty = True
    End If
    
    m_szDispMetadata = vNewValue
End Property

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

Public Property Get EandMReviewed() As Boolean
    EandMReviewed = mbEandMReviewed
End Property

Public Property Let EandMReviewed(RHV As Boolean)
    If mbEandMReviewed <> RHV Then
        mbEandMReviewed = RHV
        mbDirty = True
    End If
End Property

Public Property Get EandMReviewedBy() As String
    EandMReviewedBy = msEandMReviewedBy
End Property

Public Property Let EandMReviewedBy(RHV As String)
    If msEandMReviewedBy <> RHV Then
        msEandMReviewedBy = RHV
        mbDirty = True
    End If
End Property

Public Property Get EandMReviewDate() As Date
    EandMReviewDate = mdtEandMReviewDate
End Property

Public Property Let EandMReviewDate(RHV As Date)
    If mdtEandMReviewDate <> RHV Then
        mdtEandMReviewDate = RHV
        mbDirty = True
    End If
End Property

Public Property Get DispositionText() As String
    DispositionText = msDispositionText
End Property

Public Property Let DispositionText(RHV As String)
    If msDispositionText <> RHV Then
        msDispositionText = RHV
        mbDirty = True
    End If
End Property

Public Property Get DispositionNCID() As String
    DispositionNCID = msDispositionNCID
End Property

Public Property Let DispositionNCID(RHV As String)
    If msDispositionNCID <> RHV Then
        msDispositionNCID = RHV
        mbDirty = True
    End If
End Property

Public Property Get DispAdminOption() As String
    DispAdminOption = msDispAdminOption
End Property

Public Property Let DispAdminOption(RHV As String)
    If msDispAdminOption <> RHV Then
        msDispAdminOption = RHV
        mbDirty = True
    End If
End Property

Public Property Get EandMNCID() As String
    
    EandMNCID = msEandMNCID
End Property

Public Property Let EandMNCID(RHV As String)
    If msEandMNCID <> RHV Then
        msEandMNCID = RHV
        mbDirty = True
    End If
End Property

Public Property Get EandMCalc() As String
    EandMCalc = msEandMCalc
End Property

Public Property Let EandMCalc(RHV As String)
    If msEandMCalc <> RHV Then
        msEandMCalc = RHV
        mbDirty = True
    End If
End Property

Public Property Get WorkStatusDisposition() As String
    WorkStatusDisposition = msWSDisp
End Property

Public Property Let WorkStatusDisposition(RHV As String)
    If msWSDisp <> RHV Then
        msWSDisp = RHV
        mbDirty = True
    End If
End Property

Public Property Get FollowupTimeFrame() As String
    FollowupTimeFrame = msFollowupTimeFrame
End Property

Public Property Let FollowupTimeFrame(RHV As String)
    If msFollowupTimeFrame <> RHV Then
        msFollowupTimeFrame = RHV
        mbDirty = True
    End If
End Property

Public Property Get FollowupComments() As String
    FollowupComments = msFollowupComments
End Property

Public Property Let FollowupComments(RHV As String)
    If msFollowupComments <> RHV Then
        msFollowupComments = RHV
        mbDirty = True
    End If
End Property

Public Property Get ItemsDiscussed() As String
    ItemsDiscussed = msItemsDiscussed
End Property

Public Property Let ItemsDiscussed(RHV As String)
    If msItemsDiscussed <> RHV Then
        msItemsDiscussed = RHV
        mbDirty = True
    End If
End Property

Public Property Get Status() As EncStatus
    Status = meDispStatus
End Property

Public Sub SetStatus(ByVal RHV As EncStatus)
    If meDispStatus <> RHV Then
        meDispStatus = RHV
        mbDirty = True
    End If
End Sub

Public Property Get RTF() As String
    Dim colSectionStart As New Collection
    Dim colSectionReturned As Collection
    Dim oShared As New CWShared
    
    colSectionStart.Add "Comments :"
    colSectionStart.Add "\par"
    
    Set colSectionReturned = GetTextSections(msRTF, colSectionStart, vbTextCompare)
    
    If Not colSectionReturned Is Nothing Then
        With colSectionReturned
            Select Case .Count
            Case 1
                msRTF = .Item(1)
            Case 2
                msRTF = .Item(1) & SplitLongWords(.Item(2), 60, "\pard")
            Case 3
                msRTF = .Item(1) & SplitLongWords(.Item(2), 60, "\par\pard") & .Item(3)
            'else leave msRTF alone
            End Select
        End With
    End If
        
    If mbUpdated Then 'SCR-39592
        RTF = FormatRtfDisplay(msRTF, "Disposition", "Last updated", msOwnerName, mdDTS)
    Else
        RTF = FormatRtfDisplay(msRTF, "Disposition", "Written", msOwnerName, mdDTS)
    End If

    
End Property

Public Sub SetRTF(ByRef RTF As String, Optional Updated As Boolean)
        
    'Also compare RTFs since comments are not stored in meta data SCR#43409
    If StrComp(msRTF, RTF, vbTextCompare) <> 0 Then
        Updated = True
    End If
        
    'If owner is diff and updated then add change history  SCR#39592
    If msOwnerNCID <> "" And msOwnerNCID <> "0" _
    And (mobjEncounter.Status = Updating Or msOwnerNCID <> mobjEncounter.Logon.UserNCID) And Updated Then
        If Not mobjEncounter.AddHistoryItem(Overwritten, "Disposition Note", msRTF, msOwnerName, mdDTS, "Disposition") Then
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
'    If Not mobjEncounter.LockingObject.LockSection(eDisp, 0, SectionLock) Then
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
'        mobjEncounter.LockingObject.UnlockSection eDisp, 0
'    End If
'
'End Function
Public Function OpenSection() As Boolean
    
    Dim oPAD As Object
    Dim dLastTransferDate As Date
    
    On Error Resume Next
    If mbRegularViewRunning Then
        Exit Function
    End If
    If Not mobjEncounter.LockingObject.LockSection(eDisp, 0, SectionLock) Then
        Exit Function
    End If
    
    Screen.MousePointer = vbHourglass 'SCR#45075
'<<SCR 10673,10675(dup) Encounter data refresh before children load
'   Dropping this line to stop the refresh
'   mobjEncounter.RefreshDisplay
'>>
                                                       'SCR 65426
    If mobjEncounter.EncounterType = InpatientNote And mobjEncounter.InptNoteTypeNCID <> NCID_Discharge_Note Then 'SCR#42959
        MsgBxARMd "The Disposition Module is not available when accessed from an Inpatient Note Encounter.", vbInformation, "Encounter"
        GoTo CleanUp
    End If
    
    If Not mobjEncounter.SectionIsEditable(eDisp) Then
        MsgBxARMd "This encounter has been saved by another user and you can no longer edit the Disposition section.", vbInformation, "Encounter"
        GoTo CleanUp
    End If
'    If msOwnerNCID <> "" _
'    And (mobjEncounter.Status = Updating Or msOwnerNCID <> mobjEncounter.Logon.UserNCID) Then
''        If Not mobjEncounter.Status = Updating Then
'            ' a user other than the owner is opening this section
'            ' if they really meant to do this, then save a copy to the change history and
'            ' save the current data as belonging to the new owner
'            'Answer is always Yes SCR-26954
''            If MsgBxARMd("The Disposition section of this encounter is currently owned by " & msOwnerName & "." & vbCrLf & vbCrLf _
''            & "Do you want to save a copy in change history and take over the section?", vbYesNo, "Encounter Section Ownership") = vbNo Then
''                GoTo Cleanup
''            End If
''        End If
'        If Not mobjEncounter.AddHistoryItem(Overwritten, "Disposition Note", RTF) Then
'            MsgBxARMd "Unable to copy section to change history.", vbCritical, "Encounter"
'            GoTo CleanUp
'        End If
'        mbDirty = True
''      Unnecessary save since save will be done on close.  Also resets DTS which affect APAfterDisp Call. SCR-40868
''        If Not Save Then
''            GoTo CleanUp
''        End If
'    End If
    If mobjModule Is Nothing Then
        Set mobjModule = mobjEncounter.Comm.InitializeOLEServer(SectionEnumToNCID(eDisp))
        If mobjModule Is Nothing Then
            GoTo CleanUp
        End If
        Set mobjIChild = mobjModule
        If mobjIChild Is Nothing Then
            Set mobjModule = Nothing
            MsgBxARMd "The Disposition application module does not support the EncounterOps.IChild interface.", vbInformation, "Encounter"
            GoTo CleanUp
        End If
        On Error Resume Next
        Set mobjICoreAPI3 = mobjModule
        Call mobjEncounter.Comm.Message(cwiGET_APPLICATION_PROGID, msProgID, NCID_CW_APPLICATION, SectionEnumToNCID(eDisp))
    End If
    Call mobjIChild.SetParent(mobjParent)
    
    'SCR 65426
    If mobjEncounter.EncounterType <> InpatientNote And mobjEncounter.InptNoteTypeNCID <> NCID_Discharge_Note Then
        'Display disposition in regular mode.
        mobjModule.InitDisplayMode 1  'default  SCR#47795
    Else
        'Display disposition in discharge mode.
        '65426 SF 2-17-05 Similiar logic to PAD.Inpatient.DischargePatient
        Set oPAD = mobjEncounter.Comm.InitializeOLEServer(NCID_INPATIENT_APPLICATION)
        With oPAD.SelectedPatients.Inpatient
            If .LastTransferDate = "12:00:00 AM" Then
                dLastTransferDate = Empty
            Else
                dLastTransferDate = mobjEncounter.Comm.GmtToLocal(.LastTransferDate)
            End If
            mobjModule.InitDisplayMode 2, .ID, mobjEncounter.Comm.GmtToLocal(.AdmissionDate), dLastTransferDate
            'mobjEncounter.Comm.Message cwiSTART_OLE_SERVER, NCID_DISPOSITION_APPLICATION, NCID_CW_APPLICATION, NCID_INPATIENT_APPLICATION
        End With
    End If
    
        mbRegularViewRunning = mobjEncounter.EncounterParent.OpenChildView(mobjModule)
    
CleanUp:
    If mbRegularViewRunning Then
        OpenSection = True
    Else
        Call mobjIChild.SetParent(Nothing)
        Call mobjEncounter.LockingObject.UnlockSection(eDisp, 0)
    End If
    
    Screen.MousePointer = vbDefault
    Exit Function
    Resume
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
            'Call mobjEncounter.LockingObject.UnlockSection(eDisp, 0)
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
    Call mobjEncounter.LockingObject.UnlockSection(eDisp, 0)
End Sub

Public Function Save() As Boolean
' Assume section already locked for update before getting here
    Dim objData As DataString
    Dim objEvent As ClinicalEvent
    Dim objClinObs As ClinicalObservation
    Dim objObsMod As ClinicalObservation
    Dim objActionInfo As EventActionInfo
    Dim DispAdmin() As String
    Dim i As Integer
    
    On Error GoTo ErrHandler
    
    Save = GEMS_Save
    Exit Function

ErrHandler:
    Call mobjEncounter.CWShared.ShowVBError(Err.Number, Err.Description, "Enctr.Save ", "Enc", vbCritical)

End Function

Public Sub GEMS_Load(ByRef objParent As SectionParent, objRS As ADODB.Recordset)
    Dim i As Integer
    Dim sName As String
    Dim sValue As String
    Dim oConcept As New GEMS_ConceptCtrl
    Dim oShared As CWShared
    Dim sSql As String
    Dim aBytes() As Byte
    Dim objZlib As EncZLib
    Dim dDTS As Date
    Dim sOwnerNCID As String
    
    Set oShared = New CWShared
    
    'msRTF = ""
    mbDirty = False
    mnDataID = 0
    'msOwnerNCID = ""
    msOwnerName = ""
    'mdDTS = 0
    meDispStatus = Incomplete
    msDispositionText = ""
    msDispositionNCID = ""
    msDispAdminOption = ""
    msEandMNCID = ""
    msEandMCalc = ""
    msWSDisp = ""
    msFollowupTimeFrame = ""
    msFollowupComments = ""
    msDiscussedComments = ""
    msItemsDiscussed = ""
    msDBNICategory = ""
    msDBNICause = ""
    
    On Error Resume Next
    If objRS.EOF Then Exit Sub
    
    mnDataID = Val("" & objRS("DataID"))
    sOwnerNCID = "" & objRS("ownerncid")
    If Not sOwnerNCID = "" Then
        oConcept.UniqueId = sOwnerNCID
        msOwnerName = oConcept.PrefRep("2000").Representation
    End If
    If IsDate(objRS("DTS")) Then
        dDTS = objRS("DTS")
    End If
    
    meDispStatus = StatusNCIDToEnum(objRS("Status"))
    
    If Not IsNull(objRS!UPDATED_FLAG) Then
        mbUpdated = CBool(objRS!UPDATED_FLAG)
    End If
        
    If oShared.IsAppMode(modeCHCSI_GUI) Then
        If Len(sOwnerNCID) And sOwnerNCID <> "0" Then
            'Check if data has been updated SCR#43993
            If Not (dDTS = mdDTS And StrComp(sOwnerNCID, msOwnerNCID, vbTextCompare) = 0) Then
                'msRTF = oShared.CHCSConnection.ReadChunk("ENC_SECTIONS", "ENCOUNTERNUMBER = " & objRS("ENCOUNTERNUMBER") & " AND enc_sectionsindex = " & SectionEnumToNCID(eDisp), "DOC")
                msRTF = ReadChunkCHCSConn("ENC_SECTIONS", "ENCOUNTERNUMBER = " & objRS("ENCOUNTERNUMBER") & " AND enc_sectionsindex = " & SectionEnumToNCID(eDisp), "DOC", meDispStatus = complete)
                
                'Get Meta Data
                m_szDispMetadata = oShared.CHCSConnection.ReadChunk("enc_disposition", "ENCOUNTERNUMBER = " & objRS("ENCOUNTERNUMBER"), "DISPMETADATA")
            End If
        End If
    Else
        'Decompress if compressed SCR-24683
        If CVar(objRS("OriginalSize")) > 0 Then
            Set objZlib = New EncZLib
            
            aBytes = objRS("Doc")
            Call objZlib.DecompressData(aBytes, objRS("OriginalSize"))
            msRTF = StrConv(aBytes, vbUnicode)
            'msRTF = DecompressRTF("" & objRS("Doc"), objRS("OriginalSize"))
        End If
    End If
    
    msOwnerNCID = sOwnerNCID
    mdDTS = dDTS
    
    Dim oSQL As ICHCSII_SQL.ISqlOps
    Set oSQL = gobjCHCSIIConn.CHCSII_SQLOPS(Auto)  '''<SCR 36493
    
    'SCR 18618 WR 11/29/2001
    sSql = "select * from Enc_Disposition " _
            & "where encounternumber = " & objRS!EncounterNumber _
            & "     and facilityNCID = " & objRS!FacilityNCID
    oSQL.Execute sSql
    
    On Error Resume Next
    If Not oSQL.EOF Then
        msWSDisp = "" & oSQL("WSDisp")
        msEandMCalc = "" & oSQL("EANDMCALC")
        msFollowupTimeFrame = "" & oSQL("FOLLOWUPTIMEFRAME")
        msFollowupComments = "" & oSQL("FOLLOWUUPCOMMENTS")
        msDiscussedComments = "" & oSQL("DISCUSSEDCOMMENT")
        msItemsDiscussed = "" & oSQL("ITEMSDISCUSSED")
        'msDispositionText = "" & oSQL("SUPPLDISPDATA")
        msEandMNCID = "" & oSQL("NCID")
        mbEandMReviewed = ("" & oSQL("EandMReviewed") = "Y")
        msEandMReviewedBy = "" & oSQL("EandMReviewedBy")
        mdtEandMReviewDate = oSQL("EandMReviewedDate")
        msDispositionNCID = "" & oSQL("DispositionNCID")
        msDispositionText = "" & oSQL("DispositionText")
        msDispAdminOption = "" & oSQL("DispAdminOption")
        msDBNICategory = "" & oSQL("INJ_ILL_CATEGORY")
        msDBNICause = "" & oSQL("INJ_ILL_CAUSE")
        msDentalDispStatus = oSQL("DentalStatus")
        msDentalDispOptions = oSQL("DentalOptions")
        msDentalDispStatusDesc = oSQL("DentalStatusDesc")
        msDiscussedComments = oSQL("DentalDispComments")
        If Not oShared.IsAppMode(modeCHCSI_GUI) Then
            aBytes = oSQL("DispMetaData")
            m_szDispMetadata = StrConv(aBytes, vbUnicode)
        End If
    End If
    Set oSQL = Nothing
    Set oConcept = Nothing
End Sub
'Use GEMS_SQL.SqlOps instead of GEMS_DAS SCR-27772
Public Function GEMS_Save() As Boolean
' Assume section already locked for update before getting here
    Dim DispAdmin()     As String
    Dim i               As Integer
    Dim sSql            As String
    Dim objSQLEx        As ICHCSII_SQL.ISqlOpsEx
    Dim objSQL          As ICHCSII_SQL.ISqlOps
    Dim objZlib         As EncZLib
    Dim aBytes()        As Byte
    Dim oSQLGen         As SQLGen
    Dim szSQL           As String
    Dim sbuf            As String
    Dim bRefreshAutoCites   As Boolean
    
    On Error GoTo ErrHandler
    If Not mbDirty Then
        GEMS_Save = True
        Exit Function
    End If
    
    'Check if encounter owner is Dr. Unassigned SCR-24349
    Call mobjEncounter.CheckEncounterOwnership
    
    msOwnerNCID = mobjEncounter.Logon.UserNCID
    msOwnerName = mobjEncounter.Logon.UserName
    mdDTS = Now
    
    'WRogers 9/17/2001
    'SCR 16204
    'exit gracefully if no updates are present and view not open
    If Len(msEandMNCID & msEandMCalc & msWSDisp & msDispositionNCID & msDispositionText & msDispAdminOption & msFollowupTimeFrame & msFollowupComments & msDiscussedComments & msItemsDiscussed) = 0 And _
        Not mbRegularViewRunning Then
        mbDirty = False
        GEMS_Save = True
        GoTo CleanUp
    End If
    
        Set objSQLEx = gobjCHCSIIConn.CHCSII_SQLOPS(Auto)  '''<SCR 36493
        
        'Set oGems_das = New GEMS_DAS
        sSql = "select * from enc_sections where EncounterNumber = " & mobjEncounter.DataID & _
            " and enc_sectionsindex =" & SectionEnumToNCID(eDisp)
        'Rs.Open "select * from enc_sections where DataID = " & mnDataID, cnn, adOpenKeyset, adLockPessimistic, adCmdText
        'Set Rs = oGems_das.OpenRecordset(sSQL, adOpenKeyset, adLockPessimistic)
        
        Call objSQLEx.Execute(sSql)
        
        With objSQLEx
            If .EOF Then
                'mnDataID = GEMS_Seq_NextVal("MainSEQ")
                'rs.Open "enc_sections", Cnn, adOpenKeyset, adLockPessimistic, adCmdTable
                'add new record
                .AddNew
                !enc_sectionsindex = SectionEnumToNCID(eDisp)
                !DataID = .GetNextID(esMAINSEQ)   'GEMS_Seq_NextVal("MainSEQ")
                !EncounterNumber = mobjEncounter.EncounterID
                !FacilityNCID = mobjEncounter.FacilityNCID
            End If
         
            !OriginalSize = Len(msRTF) 'SCR-24683
    
            If Len(msRTF) <> 0 Then 'ITT Change
                Set objZlib = New EncZLib
                aBytes = StrConv(Replace(msRTF, Chr(0), ""), vbFromUnicode)
                objZlib.CompressData aBytes
            End If
            
            !Doc = aBytes
            '!Doc = CompressRTF(msRTF)
            !OwnerNCID = msOwnerNCID
            !OwnerName = msOwnerName
            !DTS = mdDTS
            !Status = StatusEnumToNCID(meDispStatus)
            !UPDATED_FLAG = CStr(mbUpdated)
            .Update
            '.Close
        End With
    
        sSql = "select * from enc_disposition " & _
                "where EncounterNumber = " & mobjEncounter.EncounterID & _
                " and FacilityNCID = " & mobjEncounter.FacilityNCID
                
        'Rs.Open "select * from encounters where dataid = " & mobjEncounter.DataId, Cnn, adOpenKeyset, adLockOptimistic, adCmdText
        'Set Rs = oGems_das.OpenRecordset(sSQL, adOpenKeyset, adLockOptimistic)
            
        Call objSQLEx.Execute(sSql)
        
        With objSQLEx
            'WRogers 9/17/2001
            'SCR 16204
            If .EOF Then
                .AddNew
                If Len(msEandMNCID) > 0 Then 'SCR-22881
                    If StrComp(msEandMNCID, EMPTY_CODE) Then 'ITT Change
                        !NCID = msEandMNCID
                    End If
                End If
            End If
            
            !EandMCalc = msEandMCalc
            !WSDisp = msWSDisp
            !EncounterNumber = mobjEncounter.EncounterID
            If Len(msDispositionNCID) > 0 Then 'ITT Change
                !DispositionNCID = msDispositionNCID
            End If
            !DispositionText = msDispositionText
            !DispAdminOption = msDispAdminOption
            !FollowupTimeFrame = msFollowupTimeFrame
            !FollowuupComments = msFollowupComments ' check that spelling
            !DiscussedComment = msDiscussedComments
            !ItemsDiscussed = msItemsDiscussed
            !FacilityNCID = mobjEncounter.FacilityNCID
            !INJ_ILL_CATEGORY = msDBNICategory
            !INJ_ILL_CAUSE = msDBNICause
            'added by #28830 to support Dental requirements
            !DentalStatus = msDentalDispStatus
            !DentalOptions = msDentalDispOptions
            !DentalStatusDesc = msDentalDispStatusDesc
            !DentalDispComments = msDiscussedComments
            aBytes = StrConv(Replace(m_szDispMetadata, Chr(0), ""), vbFromUnicode)
            !DispMetadata = aBytes
            .Update
            'Rs.Close
        End With
        
        Call mobjEncounter.SaveHIPAAData
        
        'May need to refresh data on Appointment screen SCR#49319
        mobjEncounter.Comm.Message cwiREFRESH_DATA, "APPTID|" & mobjEncounter.AppointmentID, NCID_CLINIC_SCHEDULE_APPLICATION, NCID_ENCOUNTER_APPLICATION

    
    GEMS_Save = True
    mbDirty = False
    GoTo CleanUp
ErrHandler:
    Call mobjEncounter.CWShared.ShowVBError(Err.Number, Err.Description, "EncDisp.GEMS_Save ", "Enc", vbCritical)
CleanUp:
    Set objSQL = Nothing
    Set objSQLEx = Nothing
    If GEMS_Save Then
        mobjEncounter.Comm.SetStatusText "Encounter " & mobjEncounter.EncounterID & " " & SectionEnumToName(mobjParent.meSection) & " section was saved."
    Else
        mobjEncounter.Comm.SetStatusText ""
    End If
End Function

Public Property Get EMCalcOverrideNote() As String
    EMCalcOverrideNote = msEMCalcOverrideNote
End Property

Public Property Let EMCalcOverrideNote(ByVal RHV As String)
    If msEMCalcOverrideNote <> RHV Then
       msEMCalcOverrideNote = RHV
       mbDirty = True
    End If
End Property

Public Property Get DBNICategory() As String
    DBNICategory = msDBNICategory
End Property

Public Property Let DBNICategory(RHV As String)
    msDBNICategory = RHV
    mbDirty = True
End Property

Public Property Get DBNICause() As String
    DBNICause = msDBNICause
End Property

Public Property Let DBNICause(RHV As String)
    msDBNICause = RHV
    mbDirty = True
End Property

Public Property Get DentalDispStatus() As String
    DentalDispStatus = msDentalDispStatus
End Property

Public Property Let DentalDispStatus(ByVal RHV As String)
    msDentalDispStatus = RHV
    mbDirty = True
End Property

Public Property Get DentalDispStatusDesc() As String
    DentalDispStatusDesc = msDentalDispStatusDesc
End Property

Public Property Let DentalDispStatusDesc(ByVal RHV As String)
    msDentalDispStatusDesc = RHV
    mbDirty = True
End Property

Public Property Get DentalDispComments() As String
    DentalDispComments = msDentalDispComments
End Property

Public Property Let DentalDispComments(ByVal RHV As String)
    msDentalDispComments = RHV
    mbDirty = True
End Property


Public Property Get DentalDispOptions() As String
    DentalDispOptions = msDentalDispOptions
End Property

Public Property Let DentalDispOptions(ByVal RHV As String)
    msDentalDispOptions = RHV
    mbDirty = True
End Property

'MAG 05/15/01
Public Function GetTextSections(sText As String, colSearchText As Collection, Optional CompareType As VbCompareMethod = vbBinaryCompare) As Collection
'This Function splits sText string into separate parts based on string values in the colSearchText
' and returns a collection of these string parts
On Error GoTo ErrHandler

Dim nPos As Integer
Dim nNew As Integer
Dim sSearchString As Variant
Dim sPreviousSearch As String

    Set GetTextSections = New Collection
    
    If (colSearchText Is Nothing) Or (colSearchText.Count = 0) Then
        GoTo ErrHandler
    End If
    
    nPos = 1
    sPreviousSearch = ""
    For Each sSearchString In colSearchText
        If sPreviousSearch = sSearchString Then
            nNew = InStr(nPos + 1, sText, sSearchString, CompareType)
        Else
            nNew = InStr(nPos, sText, sSearchString, CompareType)
        End If
        
        If nNew = 0 Then
            Exit For
        End If
        GetTextSections.Add Mid(sText, nPos, (nNew - nPos))
        nPos = nNew
        sPreviousSearch = sSearchString
    Next
    
    GetTextSections.Add Mid(sText, nPos, (Len(sText) - nPos + 1))
Exit Function

ErrHandler:
    Set GetTextSections = Nothing
End Function


Public Property Get DiscussedComments() As String
    DiscussedComments = msDiscussedComments
End Property

Public Property Let DiscussedComments(ByVal RHV As String)
    If msDiscussedComments <> RHV Then
        msDiscussedComments = RHV
        mbDirty = True
    End If
End Property


Public Function NotifyComplete() As Boolean

    Dim oDOM As New DOMDocument30
    Dim mobjSel As IXMLDOMSelection
    Dim objND As IXMLDOMNode
    Dim objNDA As IXMLDOMAttribute
    Dim oDom2 As New DOMDocument30
    Dim oDomS As IXMLDOMSelection
    Dim dStartDate As Date
    Dim dEndDate As Date
    Dim sComment As String
    Dim sDesc As String
    Dim sSnoID As String
    Dim sPrefix As String
    Dim sMsg As String
    Dim i As Integer
    
    On Error GoTo ErrHandler
    
    'Save done by Disposition SCR#50905
    If DispositionNCID <> NCID_ReleaseWDutyLimitations Then
        sMsg = "Remove Profile Data from DB"
        NotifyComplete = RemoveProfileData()
    End If

    Exit Function
    
    'Following code may be implemented at a later time
    sMsg = "Load Meta Data"
    'Meta data may not exist for older encounters and telcons SCR-38136
    If Len(DispMetadata) Then
        oDOM.loadXML DispMetadata
        Set mobjSel = oDOM.documentElement.selectNodes("//MetaData")
        For i = 0 To mobjSel.Context.Attributes.length - 1
            If mobjSel.Context.Attributes.Item(i).baseName = "Disposition" Then
                Set objNDA = mobjSel.Context.Attributes.Item(i)
                Exit For
            End If
        Next
        
        sMsg = "Load Profile XML"
        oDom2.loadXML objNDA.Text
        Set oDomS = oDom2.documentElement.selectNodes("//Profile")
        
        sMsg = "Get Profile info from XML"
        For i = 0 To oDomS.Item(0).Attributes.length - 1
            Select Case oDomS.Item(0).Attributes(i).baseName
                Case "StartDate"
                    dStartDate = CDate(oDomS.Item(0).Attributes(i).Text)
                Case "PREFIX"
                    sPrefix = oDomS.Item(0).Attributes(i).Text
                Case "SnoID"
                    sSnoID = oDomS.Item(0).Attributes(i).Text
                Case "EndDate"
                    dEndDate = CDate(oDomS.Item(0).Attributes(i).Text)
                Case "Description"
                    sDesc = oDomS.Item(0).Attributes(i).Text
                Case "Comment"
                    sComment = oDomS.Item(0).Attributes(i).Text
            End Select
        Next
        
        If DispositionNCID = NCID_ReleaseWDutyLimitations Then
            sMsg = "Save Profile Data to DB"
            ' Profile Data is saved in the Disposition module
            'NotifyComplete = SaveProfileData(sComment, sDesc, sSnoID, sPrefix, dStartDate, dEndDate)
            NotifyComplete = True
            
        Else
            sMsg = "Remove Profile Data from DB"
            NotifyComplete = RemoveProfileData()
        End If
    End If
    
    Exit Function
        
ErrHandler:
    Call mobjEncounter.CWShared.ShowVBError(Err.Number, sMsg & " - " & Err.Description, "Disp.NotifyComplete ", "Enc", vbCritical)

End Function

' Use SQLCurrDate instead of "Sysdate" SCR-42863
' Use SQLQuote for desc and comments SCR-35146
' Reformat SQL when Readiness data already exists SCR#49302
Private Function SaveProfileData(ByVal sComment As String, ByVal sDescription As String, ByVal sSnoID As String, _
                                    ByVal sPrefix As String, ByVal dStartDate As Date, ByVal dEndDate As Date) As Boolean

    Dim szSQL As String
    Dim szValues As String
    Dim Rs As ADODB.Recordset
    Dim objDAS As ICHCSII_DAS
    Dim F As ADODB.Field
    Dim szFlds As String
    Dim szVals As String
    Dim bAddField As Boolean
    
    On Error GoTo errXIT
    
    Set objDAS = gobjCHCSIIConn.CHCSII_DAS(Auto) '''<SCR 36493
    
    szSQL = "SELECT * FROM Readiness "
    szSQL = szSQL & "WHERE Patient_Unit_Number = '%1' "
    szSQL = szSQL & "AND Last_Readiness_Dt IN "
    szSQL = szSQL & "   (SELECT Max(Last_Readiness_Dt) "
    szSQL = szSQL & "   FROM Readiness "
    szSQL = szSQL & "   WHERE Patient_Unit_Number = '%1' )"

    szSQL = Replace(szSQL, "%1", mobjEncounter.Patient.UnitNumber)
    
    'Copy Last Row Forward
    Set Rs = objDAS.OpenRecordset(szSQL)
    
    If Rs.RecordCount Then
        
        For Each F In Rs.Fields
            bAddField = True
            Debug.Print F.Name & " - " & F.value
            Select Case F.Name
            Case "PATIENT_UNIT_NUMBER"
                szVals = szVals & "'" & mobjEncounter.Patient.UnitNumber & "', "
            Case "ENCOUNTER_NUMBER"
                szVals = szVals & "'" & mobjEncounter.EncounterID & "', "
            Case "TEMP_PROFILE" 'SCR#48704
                szVals = szVals & "'Y', "
            Case "TEMP_PROFILE_START_DT"
                szVals = szVals & gobjCHCSIIConn.SQLDate(dStartDate) & ", "
            Case "TEMP_PROFILE_END_DT"
                szVals = szVals & gobjCHCSIIConn.SQLDate(dEndDate) & ", "
            Case "TEMP_PROFILE_DIGNOS_CD"
                szVals = szVals & "'" & sSnoID & "', "
            Case "TEMP_PROFILE_DIGNOS_DSC"
                 szVals = szVals & SQLQuote(sDescription) & ", "
            Case "TEMP_PROFILE_LIMITATION"
                szVals = szVals & SQLQuote(sComment) & ", "
            Case "LAST_READINESS_DT"
                szVals = szVals & gobjCHCSIIConn.SQLDate(Now) & ", "
            Case Else
                'R. Foster SCR #48177  Developer: Robert Foster 01/20/2004 06:14 PM
                'R. Foster Note: Function Changed: CodedValues.cls - EMCodeNCID (Let); Profile.cls - Save;  modControls.bas - GetEMTable;      Impact of Change:  Modifiy code to account for ITT mode
                'ITT CONVERGENCE MODIFICATION
                'If null then don't add SCR-42863
                If Not IsNull(F.value) Then
                    If F.Type = adDBTimeStamp Or F.Type = adDate Then 'SCR#48704
                        If F.value <> #12/31/9999# Then
                            szVals = szVals & gobjCHCSIIConn.SQLDate(F.value) & ", "
                        Else
                            bAddField = False
                        End If
                    Else
                        If F.Type = adLongVarChar Or F.Type = adChar Or F.Type = adVarChar Or F.Type = adLongVarWChar Then  'SCR#48704
                            If Len(Trim$(F.value)) Then
                                szVals = szVals & "'" & F.value & "', "
                            Else
                                bAddField = False
                            End If
                        Else
                            szVals = szVals & F.value & ", "
                        End If
                    End If
                Else
                    bAddField = False
                End If
            End Select
            
            If bAddField Then
                szFlds = szFlds & F.Name & ", " & vbCrLf
            End If
        Next
        
        szFlds = Left(szFlds, Len(szFlds) - 4)
        szVals = Left(szVals, Len(szVals) - 2)
        
        szSQL = "INSERT INTO READINESS (" & szFlds
        szSQL = szSQL & ") Values (" & szVals & ")"
    Else
        'Make Virgin Row
        szSQL = "INSERT INTO READINESS " & vbCrLf
        szSQL = szSQL & "( LAST_READINESS_DT, PATIENT_UNIT_NUMBER, " & vbCrLf
        szSQL = szSQL & " ENCOUNTER_NUMBER, TEMP_PROFILE, " & vbCrLf
        szSQL = szSQL & " TEMP_PROFILE_START_DT, TEMP_PROFILE_END_DT, " & vbCrLf
        szSQL = szSQL & " TEMP_PROFILE_DIGNOS_CD, TEMP_PROFILE_DIGNOS_DSC, " & vbCrLf
        szSQL = szSQL & " TEMP_PROFILE_LIMITATION ) VALUES " & vbCrLf
        szSQL = szSQL & "( %LAST_READINESS_DT, '%PATIENT_UNIT_NUMBER', '%ENCOUNTER_NUMBER', " & vbCrLf
        szSQL = szSQL & "'%TEMP_PROFILE', %TEMP_PROFILE_START_DT, %TEMP_PROFILE_END_DT," & vbCrLf
        szSQL = szSQL & "'%TEMP_PROFILE_DIGNOS_CD', %TEMP_PROFILE_DIGNOS_DSC, %TEMP_PROFILE_LIMITATION )" & vbCrLf
        
        szSQL = Replace(szSQL, "%LAST_READINESS_DT", SQLCurrDate)
        szSQL = Replace(szSQL, "%PATIENT_UNIT_NUMBER", mobjEncounter.Patient.UnitNumber)
        szSQL = Replace(szSQL, "%ENCOUNTER_NUMBER", mobjEncounter.EncounterID)
        szSQL = Replace(szSQL, "%TEMP_PROFILE_START_DT", gobjCHCSIIConn.SQLDate(dStartDate))
        szSQL = Replace(szSQL, "%TEMP_PROFILE_END_DT", gobjCHCSIIConn.SQLDate(dEndDate))
        szSQL = Replace(szSQL, "%TEMP_PROFILE_DIGNOS_CD", sSnoID)
        szSQL = Replace(szSQL, "%TEMP_PROFILE_DIGNOS_DSC", SQLQuote(sDescription))
        szSQL = Replace(szSQL, "%TEMP_PROFILE_LIMITATION", SQLQuote(sComment))
        szSQL = Replace(szSQL, "%TEMP_PROFILE", "Y")
        
    End If
        
    objDAS.ExecuteSQL szSQL
    
XIT:
    SaveProfileData = True
    
    Exit Function
    
errXIT:
    Call mobjEncounter.CWShared.ShowVBError(Err.Number & " : " & szSQL, Err.Description, "Disp.SaveProfileData ", "Enc", vbCritical)

End Function


'SCR-36337
Private Function RemoveProfileData() As Boolean

    Dim szSQL As String
    Dim szValues As String
    Dim Rs As ADODB.Recordset
    Dim objDAS As ICHCSII_DAS
    Dim F As ADODB.Field
    Dim szFlds As String
    Dim szVals As String
    
    On Error GoTo errXIT
    
    Set objDAS = gobjCHCSIIConn.CHCSII_DAS(Auto) '''<SCR 36493
    
    szSQL = "SELECT * FROM Readiness "
    szSQL = szSQL & "WHERE Patient_Unit_Number = '%1' "
    szSQL = szSQL & "AND ENCOUNTER_NUMBER = " & "'" & mobjEncounter.EncounterID & "' "
    szSQL = szSQL & "AND TEMP_PROFILE = 'Y'"
    
    szSQL = Replace(szSQL, "%1", mobjEncounter.Patient.UnitNumber)
    
    'Copy Last Row Forward
    Set Rs = objDAS.OpenRecordset(szSQL)
    
    If Rs.RecordCount Then
        szSQL = "DELETE FROM READINESS "
        szSQL = szSQL & "WHERE Patient_Unit_Number = '%1' "
        szSQL = szSQL & "AND ENCOUNTER_NUMBER = " & "'" & mobjEncounter.EncounterID & "' "
        szSQL = szSQL & "AND TEMP_PROFILE = 'Y'"
        
        szSQL = Replace(szSQL, "%1", mobjEncounter.Patient.UnitNumber)
        
        objDAS.ExecuteSQL szSQL
    End If
    
XIT:
    RemoveProfileData = True
    
    Exit Function
    
errXIT:
    Call mobjEncounter.CWShared.ShowVBError(Err.Number & " : " & szSQL, Err.Description, "Disp.RemoveProfileData ", "Enc", vbCritical)

End Function
