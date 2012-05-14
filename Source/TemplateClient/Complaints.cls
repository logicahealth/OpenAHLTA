VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Complaints"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements TemplateData

Private Const LookupNCID As String = "658" ' actually symptom domain
Private Const StorageNCID As String = "31048" ' actually symptom domain
Private Const EventSubType As String = "119221" ' from Reason for Visit module
Private Const MyLabel As String = "Associated Reasons for Visit"

Private mListBox As VB.ListBox
Private mItems As Dictionary
Private mbDirty As Boolean

Private oEMCategoryDictionary As Scripting.Dictionary
Private oSpecialtyDictionary As Scripting.Dictionary

Public Property Get EMCategoryDictionary() As Scripting.Dictionary
    Set EMCategoryDictionary = oEMCategoryDictionary
End Property

Public Property Set EMCategoryDictionary(ByVal RHV As Scripting.Dictionary)
    Set oEMCategoryDictionary = RHV
End Property

Public Property Get SpecialtyDictionary() As Scripting.Dictionary
    Set SpecialtyDictionary = oSpecialtyDictionary
End Property

Public Property Set SpecialtyDictionary(ByVal RHV As Scripting.Dictionary)
    Set oSpecialtyDictionary = RHV
End Property


Public Sub AddItem(ByRef rsSnoID As String, Optional ByRef rsDesc As String = "")
    On Error Resume Next
    If rsDesc = "" Then
        InitSnoConvert
        rsDesc = gobjSnoConvert.SnoDescription(rsSnoID) _
            & "  " & gobjSnoConvert.SnoICD9(rsSnoID)
    End If
    mItems.Add rsSnoID, rsDesc
    mbDirty = True
    mDisplayItems
End Sub

Private Sub TemplateData_DoMoveDown()
    
    Dim sNCIDFrom As String
    Dim sDescFrom As String
    
    Dim sNCIDTo As String
    Dim sDescTo As String
    
    Dim i As Integer
    
    For i = mListBox.ListCount - 1 To 0 Step -1
        If mListBox.Selected(i) Then
            sNCIDFrom = CStr(mListBox.ItemData(i))
            sDescFrom = mListBox.List(i)
            sNCIDTo = CStr(mListBox.ItemData(i + 1))
            sDescTo = mListBox.List(i + 1)
            'SCR-53144 5/4/2004 RBELL Fix data related error
            mItems.Key(sNCIDFrom) = "-1"
            mItems.Key(sNCIDTo) = sNCIDFrom
            mItems(sNCIDFrom) = sDescFrom
            mItems.Key("-1") = sNCIDTo
            mItems(sNCIDTo) = sDescTo
            mDisplayItems
            mListBox.Selected(i + 1) = True
            mbDirty = True
        End If
    Next i

End Sub

Private Sub TemplateData_DoMoveUp()
    
    Dim sNCIDFrom As String
    Dim sDescFrom As String
    
    Dim sNCIDTo As String
    Dim sDescTo As String
    
    Dim i As Integer
    
    For i = mListBox.ListCount - 1 To 0 Step -1
        If mListBox.Selected(i) Then
            sNCIDFrom = CStr(mListBox.ItemData(i))
            sDescFrom = mListBox.List(i)
            sNCIDTo = CStr(mListBox.ItemData(i - 1))
            sDescTo = mListBox.List(i - 1)
            'SCR-53144 5/4/2004 RBELL Fix data related error
            mItems.Key(sNCIDFrom) = "-1"
            mItems.Key(sNCIDTo) = sNCIDFrom
            mItems(sNCIDFrom) = sDescFrom
            mItems.Key("-1") = sNCIDTo
            mItems(sNCIDTo) = sDescTo
            mDisplayItems
            mListBox.Selected(i - 1) = True
            mbDirty = True
            Exit For
        End If
    Next i

End Sub


Private Function TemplateData_Export() As String
    TemplateData_Export = SerializeArray(mItems.Keys, gsListDelim)
End Function

Private Function TemplateData_GetFromEncounter(objEncounter As CHCSIIEncounterOps.Encounter) As Boolean
    
    Dim objProblems As MMMHISProblems.Problem
    Dim objProblemOps As MMMHISProblems.ProblemOps
    Dim objComplaint As MMMHISProblems.OneComplaint
    Dim objIConvert As actxSearchSnoMed.IConvert
    
    Dim szDesc As String
    
    On Error GoTo ErrHandler
    
    Set objProblems = cwobjCWComm.InitializeOLEServer(NCID_PROBLEMS_APPLICATION)
    Set objIConvert = New actxSearchSnoMed.IConvert
    
    If Not objProblems Is Nothing Then
        For Each objComplaint In objProblems.ProblemOps.Complaints.Filter(objEncounter.EncounterID, NCID_ACTIVE)
            szDesc = objComplaint.Complaint & Space(2) & objIConvert.SnoICD9(objComplaint.SnoID)
            If Len(Trim(szDesc)) > 0 Then mItems.Add objComplaint.SnoID, Trim(szDesc)
        Next
    End If
    
    
    TemplateData_GetFromEncounter = True
    
    GoTo Cleanup
    
ErrHandler:
    If Err = 457 Then
        Resume Next
    Else
        Call gobjShared.ShowVBError(Err.Number, Err.Description, "Complaints.GetFromEncounter", "TemplateClient", vbExclamation)
    End If
    'Resume Next
Cleanup:
        
    mbDirty = True
    
    Set objProblems = Nothing
    Set objProblemOps = Nothing
    Set objComplaint = Nothing
    Set objIConvert = Nothing
    
    On Error Resume Next
    Call mDisplayItems      'Clear and Fill mListBox
    
End Function

Private Function TemplateData_Import(rsString As String) As Boolean
    Dim colNCIDs As Collection
    Dim varNCID As Variant
    
    'SCR-34348 4/8/03 RBELL DISABLED IMPORT OF THIS ITEM FOR PGUI MODES
    If gobjShared Is Nothing Then Set gobjShared = New CWShared
    If gobjShared.IsAppMode(modeCHCSI_GUI) Then Exit Function '<SCR 36493
    
    Set colNCIDs = DeserializeString(rsString, gsListDelim)
    For Each varNCID In colNCIDs
        On Error Resume Next
        AddItem CStr(varNCID)
        On Error GoTo 0
    Next varNCID
    If Not mListBox Is Nothing Then
        Call mDisplayItems
    End If
    mbDirty = False
    TemplateData_Import = True
End Function

Private Function TemplateData_GetQueryCriteria() As String
    Dim i As Integer
    Dim sStr As String
    For i = 0 To mItems.Count - 1
        If sStr <> "" Then
            sStr = sStr & " or"
        End If
        sStr = sStr & " (type_ncid = " & StorageNCID & " and value_ncid = " & mItems.Keys(i) & ")"
    Next i
    TemplateData_GetQueryCriteria = sStr
End Function

Private Sub mDisplayItems()
    Dim i As Integer
    mListBox.Clear
    For i = 0 To mItems.Count - 1
        If Not mItems.Keys(i) = "" Then
            mListBox.AddItem mItems.Items(i)
            mListBox.ItemData(mListBox.NewIndex) = mItems.Keys(i)
        End If
    Next i
End Sub

Private Sub Class_Initialize()
    Set mItems = New Dictionary
End Sub

Private Sub Class_Terminate()
    Set mItems = Nothing
    Set mListBox = Nothing
    Set oEMCategoryDictionary = Nothing
    Set oSpecialtyDictionary = Nothing
End Sub

Private Property Get TemplateData_CanEdit() As Boolean
    TemplateData_CanEdit = False
End Property

Private Function TemplateData_DeSerialize(rvData As Variant) As Boolean
    Dim objRef As ENCTMPLTATLLib.ETXref
    
    TemplateData_DeSerialize = GEMS_TemplateData_DeSerialize(rvData)

End Function
'<< GEMS    '*-- Coleman 060601: Adding Theater support
Private Function GEMS_TemplateData_DeSerialize(rvData As Variant) As Boolean
    Dim objRef As GEMS_EncTmplt.ETXref
    For Each objRef In rvData.References
        If objRef.eType = StorageNCID Then
            On Error Resume Next
            AddItem objRef.Value
            On Error GoTo 0
        End If
    Next objRef
    If Not mListBox Is Nothing Then
        Call mDisplayItems
    End If
    mbDirty = False
    '*-- Coleman 081001:
    '*-- To Do: Looks like the following line is needed, but must be tested...
    'TemplateData_DeSerialize = True
End Function
' GEMS >>

Private Sub TemplateData_DoAdd()
    'SCR-50502 RBELL 2/27/2004 Added check to see if search form was loaded if not then send core messages
    Dim bIsFormLoaded As Boolean
    Dim oForm As Form
    
    bIsFormLoaded = False
    For Each oForm In Forms
        If UCase(oForm.Name) = "FRMSEARCH" Then
          bIsFormLoaded = True
          Exit For
        End If
    Next
    If Not bIsFormLoaded Then cwobjCWComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_TEMPLATE_MANAGEMENT
    frmMedcin.DoAdd Me, "Complaint Search", mlmComplaints_Search 'Diagnosis_Search
    If Not bIsFormLoaded Then cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_TEMPLATE_MANAGEMENT
End Sub

Private Sub TemplateData_DoEdit()

End Sub

Private Sub TemplateData_DoRemove()
    Dim sNCID As String
    Dim i As Integer
    For i = mListBox.ListCount - 1 To 0 Step -1
        If mListBox.Selected(i) Then
            sNCID = CStr(mListBox.ItemData(i))
            mItems.Remove sNCID
            mListBox.RemoveItem i
            mbDirty = True
        End If
    Next i
End Sub

Private Property Let TemplateData_isDirty(RHS As Boolean)
    mbDirty = RHS
End Property

Private Property Get TemplateData_isDirty() As Boolean
    TemplateData_isDirty = mbDirty
End Property

Private Property Get TemplateData_Items() As Scripting.IDictionary
    Set TemplateData_Items = mItems
End Property

Private Property Get TemplateData_Label() As String
    TemplateData_Label = MyLabel
End Property

Private Property Set TemplateData_ListBox(RHS As Object)
    Set mListBox = RHS
    Call mDisplayItems
End Property

Private Property Get TemplateData_ListBox() As Object

End Property

Private Function TemplateData_Serialize(rvData As Variant) As Boolean
    
    TemplateData_Serialize = GEMS_TemplateData_Serialize(rvData)

End Function
'<< GEMS    '*-- Coleman 060601: Adding Theater support
Private Function GEMS_TemplateData_Serialize(rvData As Variant) As Boolean
    Dim i As Integer
    Dim objRef As GEMS_EncTmplt.ETXref
    
    For i = 0 To mItems.Count - 1
        Set objRef = rvData.References.Add
        objRef.eType = StorageNCID      '*-- Coleman 060601: field name change
        objRef.Value = mItems.Keys(i)
    Next i
    Set objRef = Nothing
    GEMS_TemplateData_Serialize = True
    mbDirty = False
End Function
' GEMS >>

