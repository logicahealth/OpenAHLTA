VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Note"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const msCLASS_NAME$ = "Note"
Private Const msMODULE_NAME$ = msCLASS_NAME$

Private mobjNote As TextNote
Private mNoteList As Collection

Public Type uListItem
    EventID As String
    Note_Title As String
    Note_Collection_Key As String
    Note_Category As String
    Note_Text As String
    Entry_Tool_NCID As String
    Date_Created As Date
    Entered_By As String
    Entered_By_NCID As Long
    Date_Modified As Date
    Modified_By As String
    Modified_By_NCID As Long
End Type

#If DEBUGON Then
  Private DR            As DebugRec
  Private DL            As cDebugLog
  Private Const msEXE_NAME$ = "Note"     'Could just use App.Title for this.
#End If

Private Sub AddNewItemToNotesList()
    
    Dim vListItem As uListItem
    
    vListItem.EventID = mobjNote.EventID
    vListItem.Note_Title = mobjNote.Title
    vListItem.Note_Category = mobjNote.Category
    vListItem.Note_Text = mobjNote.RTF
    vListItem.Entry_Tool_NCID = gParentNCID
    vListItem.Date_Created = mobjNote.DateCreated
    vListItem.Entered_By = mobjNote.OwnerName
    vListItem.Entered_By_NCID = mobjNote.OwnerNCID
    vListItem.Date_Modified = mobjNote.DateModified
    vListItem.Modified_By = mobjNote.ModifiedBy
    vListItem.Modified_By_NCID = mobjNote.ModifiedBy
    vListItem.Note_Collection_Key = CStr(cwobjPatient.UnitNumber) & "|" & CStr(vListItem.Date_Created) & "|" & CStr(vListItem.Entered_By_NCID)
    mNoteList.Add vListItem, vListItem.Note_Collection_Key
    
End Sub


Public Function AddNote(Optional pEncounterID As Long) As Boolean
    
    If Not pEncounterID = 0 Then
       If MsgBox("Adding a note to the encounter using the free text option " & _
                "may cause the auto calculated E&M code to be incorrect. Do you wish to continue?", vbYesNo + vbQuestion) = vbNo Then
          Exit Function
       End If
    End If
    
    Set mobjNote = New TextNote
    mobjNote.IsNew = True
    mobjNote.EncounterID = pEncounterID
    mobjNote.OwnerNCID = cwobjLogonEnviron.UserNCID
    mobjNote.SetUpDataString
    frmTextNote.ParentObject = Me
    frmTextNote.Display
    
    AddNote = Not mobjNote.IsNew
    If AddNote = True Then AddNewItemToNotesList
    
End Function

Public Function EditNote(NoteItem As uListItem) As Boolean
    
    Set mobjNote = New TextNote
    mobjNote.LoadNote NoteItem
    frmTextNote.ParentObject = Me
    frmTextNote.Display
    
    EditNote = mobjNote.HasItemBeenSaved
    If Not mobjNote.HasItemBeenSaved Then Exit Function
    
    mNoteList.Remove NoteItem.Note_Collection_Key
    AddNewItemToNotesList
    
End Function
Public Function DeleteNote(NoteItem As uListItem) As Boolean
    
    Set mobjNote = New TextNote
    DeleteNote = mobjNote.DeleteNote(NoteItem)
    If DeleteNote Then mNoteList.Remove NoteItem.Note_Collection_Key
        
End Function

Public Sub Initialize(ByVal robjComm As Object, robjLogonEnviron As Object, robjConfig As Object, robjPatient As Object, Optional rParentNCID As String)
    
    On Error GoTo ErrHandler
    
    '---create pointers to objects---
    Set cwobjCWComm = robjComm
    Set cwobjLogonEnviron = robjLogonEnviron
    Set cwobjConfig = robjConfig
    Set cwobjPatient = robjPatient
    gParentNCID = rParentNCID
    
    '---create OLE objects---
    Set gConceptGems = New GEMS_ConceptCtrl
    
    Exit Sub

ErrHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "NoteTool.Initialize", "Note", vbExclamation)
    'Resume Next

End Sub

Public Function RetrieveNoteList(Optional pEncounterID As Long, Optional LoadRTF As Boolean = True) As Boolean

    Dim cList As New Collection
    Dim objDataString As DataString
    Dim objClinicalText As ClinicalText
    Dim uItem As uListItem
    Dim sSQL As String
    Dim DataID As Long
    Dim oConn As New CHCSII_CONN.Conn
    
    Const COMPRESSED_DATA_ID As String = "COMPRESSED_DATA_ID="
    On Error GoTo ErrHandler
    RetrieveNoteList = False
    
    sSQL = "SELECT DATA_ID, DATE_CREATED, CREATED_BY, " & vbLf
    sSQL = sSQL & "DATE_MODIFIED, MODIFIED_BY, ENTRY_TOOL_NCID " & vbLf
    sSQL = sSQL & "FROM GENERIC_APP_NOTES  WHERE UNIT_NUMBER = " & cwobjPatient.UnitNumber
    sSQL = sSQL & " AND FACILITY_NCID = " & oConn.SQLQuote(cwobjLogonEnviron.FacilityNCID)
    If Not pEncounterID = 0 Then sSQL = sSQL & " AND ENCOUNTER_NUMBER = " & oConn.SQLQuote(pEncounterID)
    If Not Trim(gParentNCID) = "" Then sSQL = sSQL & " AND ENTRY_TOOL_NCID = " & oConn.SQLQuote(gParentNCID)
    sSQL = sSQL & " ORDER BY DATE_CREATED"
    
    gSQL.Execute sSQL
    If Not (gSQL.BOF And gSQL.EOF) Then
       Do Until gSQL.EOF
            uItem.Date_Created = gSQL("DATE_CREATED")
            uItem.Entered_By_NCID = Val(gSQL("CREATED_BY"))
            If uItem.Entered_By_NCID = 0 Then
               uItem.Entered_By = ""
            Else
               gConceptGems.UniqueId = uItem.Entered_By_NCID
               uItem.Entered_By = gConceptGems.PrefRep("2000").Representation
              
            End If
            
            If Not IsEmpty(gSQL("DATE_MODIFIED")) Then
               uItem.Date_Modified = gSQL("DATE_MODIFIED")
            Else
               uItem.Date_Modified = Empty
            End If
            
            uItem.Modified_By_NCID = Val(gSQL("MODIFIED_BY"))
            If uItem.Modified_By_NCID = 0 Then
               uItem.Modified_By = ""
            Else
               gConceptGems.UniqueId = uItem.Modified_By_NCID
               uItem.Modified_By = gConceptGems.PrefRep("2000").Representation
            End If
            
            uItem.Entry_Tool_NCID = gSQL("ENTRY_TOOL_NCID")
            uItem.EventID = gSQL("DATA_ID")
            uItem.Note_Text = ""
            If Not uItem.EventID = "" And LoadRTF = True Then
                  uItem.Note_Text = GetGEMSDocument(CLng(uItem.EventID))
            End If
            uItem.Note_Collection_Key = CStr(cwobjPatient.UnitNumber) & "|" & CStr(uItem.Date_Created) & "|" & CStr(uItem.Entered_By_NCID)
            cList.Add uItem, uItem.Note_Collection_Key
            gSQL.MoveNext
       Loop
    End If
    
    RetrieveNoteList = True
    GoTo CleanUpAndExit
    
ErrHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "NoteTool.Initialize", "Note", vbExclamation)
    'Resume Next

CleanUpAndExit:
    Set objDataString = Nothing
    Set objClinicalText = Nothing
    Set mNoteList = cList
    Set cList = Nothing
    
End Function

Private Sub Class_Initialize()

    Const sSUBR_NAME$ = "Class_Initialize"
    
    #If DEBUGON Then
        Set DL = GetObject(Empty, "dlog.cdebuglog")
        DL.debugStart
        Dim DR As DebugRec
        DL.DebugLog msEXE_NAME$, msCLASS_NAME$, "Class_Initialize", "Begin", DR
    #End If

    '---Declare local variables
    
    On Error GoTo ErrHandler
    Set gobjDas = New DasCtrl
    Set mNoteList = New Collection
    Set gobjClinician = New Clinician
    
    Set gobjShared = New CWShared
    Set gSQL = GetSQLOps
    
    '------------------------------------
    '  This statement is needed to debug
    '    this OLE Server.
    '------------------------------------
    #If ciDEBUG = 1 Then
        If cwbDebugLogon = False Then
            Exit Sub
        End If
    #End If
    #If DEBUGON Then
        DL.DebugLog msEXE_NAME$, msCLASS_NAME$, "Class_Initialize", "End", DR
    #End If

    Exit Sub

ErrHandler:

End Sub 'Class_Initialize


Private Sub Class_Terminate()

    Const sSUBR_NAME$ = "Class_Terminate"

    '============================================================================
    '
    '----------------------------------------------------------------------------
    'Modifications
    '============================================================================
    
    #If DEBUGON Then
        Dim DR As DebugRec
        DL.DebugLog msEXE_NAME$, msCLASS_NAME$, "Class_Terminate", "Begin", DR
    #End If

    '----------------------------------------
    '  If debugging, make sure to "log off".
    '----------------------------------------
    Set mNoteList = Nothing
    'Set gobjDas = Nothing
    'Set gobjClinician = Nothing
    
    #If ciDEBUG = 1 Then
        cwbDebugLogOff
    #End If
    #If DEBUGON Then
        DL.DebugLog msEXE_NAME$, msCLASS_NAME$, "Class_Terminate", "End", DR
    #End If

End Sub 'Class_Terminate

Public Property Get Note() As TextNote
    If mobjNote Is Nothing Then Set mobjNote = New TextNote
    Set Note = mobjNote
End Property

Public Property Let Note(ByVal RHV As TextNote)
    mobjNote = RHV
End Property

Public Property Get NoteList() As Collection
    Set NoteList = mNoteList
End Property
