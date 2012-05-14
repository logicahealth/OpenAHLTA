VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Encounters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"Class1"
Attribute VB_Ext_KEY = "Member0" ,"Class1"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable to hold collection
Private mCol As Collection
Private mobjCWShared As CWShared
Dim mobjEncounterOps As EncounterOps

Public Property Get Item(vntIndexKey As Variant) As Encounter
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property
Public Property Get FullItem(vntIndexKey As Variant) As Encounter
  
    Dim objEncounter As Encounter
    
    'Replace Lite Encounter w/ complete encounter
    Set objEncounter = mCol(vntIndexKey)

'Commented out to improve performance
'    If objEncounter.FullyLoaded = False Then
'        Set objEncounter = mobjEncounterOps.GetEncounter(objEncounter.PatientID, objEncounter.FacilityNCID, objEncounter.EncounterID)
'        mCol.Remove objEncounter.StorageKey
'        mCol.Add objEncounter, objEncounter.StorageKey
'    End If
  
   Set FullItem = mCol(vntIndexKey)
  
End Property
Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mobjCWShared = New CWShared
    Set mobjEncounterOps = GetEncounterOps()
End Sub

Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
    Set mobjCWShared = Nothing
    Set mobjEncounterOps = Nothing
End Sub

Public Sub Initialize(ByVal UnitNumber As String, ByVal Lite As Boolean)

    Dim c As Collection
    Dim objEncounter As CHCSIIEncounterOps.Encounter
    Dim lMousePointer As Long
    Dim bFlag As Boolean
    
    On Error GoTo ErrXIT
    
    Set mCol = New Collection
    
    lMousePointer = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    
    cwobjCWComm.SetStatusText "Retrieving Previous Encounters ..."
    
    bFlag = GetSetting(App.EXEName, "Test", "Get All RTF's", False)
    SaveSetting App.EXEName, "Test", "Get All RTF's", bFlag
    
    ' Begin:SCR #40046;   Developer: MAG 08/29/2003 04:04 PM
    If Not Lite Then
        Set c = mobjEncounterOps.GetPreviousEncounters(UnitNumber)
    Else
        Set c = mobjEncounterOps.GetPreviousEncountersLite(UnitNumber, False)
    End If
    ' End: SCR #40046;
    
    For Each objEncounter In c
        
        Select Case objEncounter.Status
        Case complete, NeedsCoSignature, Updated, PatientLeftWithoutBeingSeen, _
             CancelByFacility, CancelByPatient, NoShow, Updating
            ' we want these
            mCol.Add objEncounter, objEncounter.StorageKey
        Case Else
            'we don't want these
        End Select
        
    Next

Cleanup:
    Screen.MousePointer = lMousePointer
    Call cwobjCWComm.SetStatusText("")
    Set objEncounter = Nothing
    Set c = Nothing
    
    Exit Sub
    
ErrXIT:
    Call mobjCWShared.ShowVBError(Err.Number, Err.Description, "PreviousEncounter.Initialize", App.Title, vbCritical)
    Resume Cleanup
    
End Sub

Public Sub GetInpatientNotes(ByVal UnitNumber As String, ByVal Lite As Boolean)

    Dim c As Collection
    Dim objEncounter As CHCSIIEncounterOps.Encounter
    Dim lMousePointer As Long
    
    On Error GoTo ErrXIT
    
    Set mCol = New Collection
    
    lMousePointer = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    
    cwobjCWComm.SetStatusText "Retrieving Inpatient Notes ..."
    
    Set c = mobjEncounterOps.GetInpatientNoteEncounters(UnitNumber, Lite)
    
    For Each objEncounter In c
        
      If objEncounter.EncounterType = InpatientNote Then
        ' SCR #42995;   Developer:  12/29/2003
        Select Case objEncounter.Status
        Case InProgress, complete, Updating, Updated, NeedsCoSignature
            ' we want these
            mCol.Add objEncounter, objEncounter.StorageKey
        Case Else
            'we don't want these
        End Select
      End If
    Next

Cleanup:
    Screen.MousePointer = lMousePointer
    Call cwobjCWComm.SetStatusText("")
    Set objEncounter = Nothing
    Set c = Nothing
    
    Exit Sub
    
ErrXIT:
    Call mobjCWShared.ShowVBError(Err.Number, Err.Description, "PreviousEncounter.GetInpatientNotes", App.Title, vbCritical)
    Resume Cleanup
    
End Sub


