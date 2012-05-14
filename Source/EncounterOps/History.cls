VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "History"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mobjEncounter As Encounter
Private mobjParent As SectionParent

Private mcolItems As Collection

Friend Function AddHistoryItem(ByVal Reason As ChangeHistoryEnum, ByRef ItemName As String, ByRef RTF As String) As Boolean
    Dim objItem As TextNote
    Set objItem = New TextNote
    
    With objItem
        .mbIsComplete = True
        .msCategory = Reason
        .msTitle = ItemName
        .msRTF = RTF
        .msOwnerNCID = mobjEncounter.Logon.UserNCID
        .msOwnerName = mobjEncounter.Logon.UserName
        .mdDTS = Now
    End With
    AddHistoryItem = Save(objItem)
    Set objItem = Nothing
End Function

'Clean up Change Histry RTF SCR-28667
Public Property Get RTF() As String
    Dim objItem As TextNote
    Dim eReason As ChangeHistoryEnum
    Dim sWhat As String
    Dim sAction As String
    Dim sHeader As String
    Dim sVerb As String
    
    On Error GoTo ErrHandler
    
    If Not mcolItems Is Nothing Then
        If mcolItems.Count < 1 Then
            Exit Property
        End If
    Else
        Exit Property
    End If
    
    For Each objItem In mcolItems
        If RTF = "" Then
            If Len(mobjEncounter.Comm.CmdLineSwitch("GREEN")) = 0 Then
                'Vitals and Allergies use \f2 so define \f2 in header SCR-39958
                RTF = "{\rtf1\ansi\ucl\deff0\deflang1033\deflangfe1033" _
                    & "{\fonttbl{\f0\fswiss\fcharset0\fprq2 Arial;}{\f2\fswiss\fcharset0\fprq2 Arial;}}" _
                    & "{\colortbl;\red0\green0\blue0;}" _
                    & "\pard\plain\f0\fs16\b\par CHANGE HISTORY \par\par"
                sHeader = RTF
            Else
                RTF = "{\rtf1\ansi\ucl\deff0\deflang1033\deflangfe1033" _
                    & "{\fonttbl{\f0\fswiss\fcharset0\fprq2 Arial;}}" _
                    & "{\colortbl;\red0\green0\blue0;}" _
                    & "\pard\plain\f0\fs16\b "
                    
                'Extended line
                RTF = RTF & "_________________________________________________________________________________________________________ \par "
                
                'RTF = RTF & "\pard\brdrt\brdrs\brdrw20\brsp20\brdrb\brdrs\brdrw20\brsp20\brdrbtw\brdrs\brdrw20\brsp20\plain "

                RTF = RTF & "\pard\plain\f0\fs16\b CHANGE HISTORY\par "
                sHeader = RTF
            End If
        Else
            If Len(mobjEncounter.Comm.CmdLineSwitch("GREEN")) = 0 Then
                RTF = RTF & " \par \pard\plain\f0\fs16 \par "
            Else
                RTF = RTF & "\pard\plain\f0\fs16 "
            End If
        End If
        
        If StrComp(objItem.msTitle, "AutoCites", vbTextCompare) = 0 Then 'SCR#43297
            sVerb = "Were"
        Else
            sVerb = "Was"
        End If
        
        eReason = objItem.msCategory
        Select Case eReason
        Case Deletion
            sAction = sVerb & " Deleted"
        Case LockedOut
            sAction = sVerb & " Saved to History (After Lockout)"
        Case Overwritten
            sAction = sVerb & " Overwritten"
        Case EncAmended
            sAction = "No Longer Applies because this Encounter Was Opened for Amendment"
            objItem.msTitle = "Signature(s)"
        Case Else
            sAction = "" 'Reset Action string SCR-16247
        End Select
        If objItem.msTitle = "" Then
            sWhat = "Entry"
        Else
            sWhat = objItem.msTitle
        End If
        
        'Removed duplicate Change History Header SCR-35042
        If Len(mobjEncounter.Comm.CmdLineSwitch("GREEN")) = 0 Then
            RTF = RTF & "\pard\plain\f0\fs16\b\i\ul The following " & sWhat & " " & sAction & " \b0 by " _
                & objItem.msOwnerName & " @ " & FormatDTG(objItem.mdDTS) & ":\par" _
                & " \pard\plain \widctlpar\adjustright \cgrid" & objItem.msRTF & "}"
        Else
            'Removed close bracket & "}" 'SCR#39592
            RTF = RTF & "\pard\plain\f0\fs16\b\i\ul The following " & sWhat & " " & sAction & " \b0 by " _
                & objItem.msOwnerName & " @ " & FormatDTG(objItem.mdDTS) & ":\par" _
                & " \pard\plain \widctlpar\adjustright \cgrid" & Replace(objItem.msRTF, "\fs16", "\fs12")
        End If
    Next objItem
    If RTF <> "" Then
        RTF = RTF & "}"
    End If
    
    Set objItem = Nothing
    
    Exit Property
    
ErrHandler:
    Call mobjEncounter.CWShared.ShowVBError(Err.Number, Err.Description, "History.RTF (Get)", "EncounterOps", vbCritical)
    Exit Property
    
    Resume
End Property

Public Sub Load(ByRef objParent As SectionParent, ByRef colDataStrings As Collection)
    Dim objData As DATASTRATLLib.DataString
    Dim objEvent As CLINTYPEATLLib.ClinicalEvent
    Dim objItem As TextNote
    
    If Not mobjEncounter Is Nothing Then
        GEMS_Load objParent, colDataStrings
        Exit Sub
    End If
    
End Sub

Public Sub OpenEncounter(ByRef objParent As SectionParent, ByRef objEncounter As Encounter)
    Set mobjParent = objParent
    Set mobjEncounter = objEncounter
End Sub

Public Sub CloseEncounter()
    Set mobjParent = Nothing
    Set mobjEncounter = Nothing
End Sub

Private Function Save(ByRef objItem As TextNote) As Boolean
' Assume section NOT already locked for update before getting here
    Dim objData As DataString
    Dim bWasNew As Boolean
    On Error GoTo ErrHandler
    
    Save = GEMS_Save(objItem)
    Exit Function

ErrHandler:
    Call mobjEncounter.CWShared.ShowVBError(Err.Number, Err.Description, "TextNote.Save ", "EncounterOps", vbCritical)

End Function

Public Function GEMS_Load(ByRef objParent As SectionParent, ByVal objRS As ADODB.Recordset)
    
    Dim objItem As TextNote
    Set mcolItems = New Collection
    
    Do While Not objRS.EOF
    
        Set objItem = New TextNote
        Call objItem.GEMS_ReadFromRecordSet(objRS)
        
        If mcolItems.Count > 0 Then
            mcolItems.Add objItem, CStr(objItem.mnDataID), 1
        Else
            mcolItems.Add objItem, CStr(objItem.mnDataID)
        End If
        Set objItem = Nothing
        objRS.MoveNext
    Loop
   
End Function

Private Function GEMS_Save(ByRef objItem As TextNote) As Boolean
' Assume section NOT already locked for update before getting here

    Dim bWasNew As Boolean
    On Error GoTo ErrHandler
    
    If objItem.mnDataID = 0 Then
        bWasNew = True
    Else
        If Not mobjEncounter.LockingObject.StartSectionUpdate(eHistory, objItem.mnDataID) Then
            GoTo CleanUp
        End If
    End If

    Dim sql As String
    Dim oSQLOps As ICHCSII_SQL.ISqlOps
    Set oSQLOps = gobjCHCSIIConn.CHCSII_SQLOPS(Auto)  '''<SCR 36493
    Dim rtf_index As Long
    
    If objItem.mnDataID = 0 Then
        rtf_index = oSQLOps.GetNextID()
        sql = "insert into enc_rtfs (ENCOUNTERNUMBER,FACILITYNCID,DTS, rtftype, ENC_RTFSINDEX ) "
        sql = sql & " values (" & mobjEncounter.EncounterID & "," & mobjEncounter.FacilityNCID & ","
        sql = sql & gobjCHCSIIConn.SQLDate(objItem.mdDTS) & ", 'TextNote', " & rtf_index & ")"
        
        oSQLOps.Execute sql
        objItem.mnDataID = rtf_index
    End If
    
    Call objItem.GEMS_WriteToDataString

    GEMS_Save = True
    
    GoTo CleanUp
    
ErrHandler:
    Call mobjEncounter.CWShared.ShowVBError(Err.Number, Err.Description, "TextNote.Save ", "EncounterOps", vbCritical)
    
CleanUp:
    Set oSQLOps = Nothing
    On Error Resume Next
    If bWasNew Then
        mcolItems.Add objItem, CStr(objItem.mnDataID)
    Else
        Call mobjEncounter.LockingObject.EndSectionUpdate(eHistory, objItem.mnDataID)
    End If
    
    If GEMS_Save Then
        mobjEncounter.Comm.SetStatusText "Encounter " & mobjEncounter.EncounterID & " change history was saved."
    Else
        mobjEncounter.Comm.SetStatusText ""
    End If
End Function

