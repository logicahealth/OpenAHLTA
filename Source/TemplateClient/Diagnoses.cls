VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Diagnoses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements TemplateData

Private Const LookupNCID As String = "658"
Private Const MyLabel As String = "Diagnoses"

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
    Dim objDiagnosis As MMMHISProblems.OneDiagnosis
    Dim objIConvert As actxSearchSnoMed.IConvert
    
    Dim szDesc As String
    
    On Error GoTo ErrHandler
    
    Set objProblems = cwobjCWComm.InitializeOLEServer(NCID_PROBLEMS_APPLICATION)
    Set objIConvert = New actxSearchSnoMed.IConvert
    
    If Not objProblems Is Nothing Then
        For Each objDiagnosis In objProblems.ProblemOps.Diagnosis.Filter(objEncounter.EncounterID, NCID_ACTIVE)
            
            'SCR-43649 RBELL 10/22/03 Added extra data to the key in PGUI mode to determine wether the data is from the ICd9 or SnoID
            If gobjShared.IsAppMode(modeCHCSI_GUI) Then
                If objDiagnosis.FromProblemList Then
                    szDesc = objDiagnosis.Diagnosis & Space(2) & objDiagnosis.SnoID 'SCR-43649 RBELL 11/4/3 Removed call to objIConvert.SnoICD9
                    mItems.Add objDiagnosis.SnoID & ",I", Trim(szDesc)
                Else
                    szDesc = objDiagnosis.Diagnosis & Space(2) & objIConvert.SnoICD9(objDiagnosis.SnoID)
                    mItems.Add objDiagnosis.SnoID & ",M", Trim(szDesc)
                End If
            Else
                szDesc = objDiagnosis.Diagnosis & Space(2) & objIConvert.SnoICD9(objDiagnosis.SnoID)
                mItems.Add objDiagnosis.SnoID, Trim(szDesc)
            End If
            'AddItem objDiagnosis.SnoID
            
        Next
    End If
    
    TemplateData_GetFromEncounter = True
    GoTo Cleanup
ErrHandler:
    If Err = 457 Then Resume Next 'Ignore duplicates diagnosis from encounter scr-36579
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Diagnoses.GetFromEncounter", "TemplateClient", vbExclamation)
Cleanup:
    
    mbDirty = True
    
    Set objProblems = Nothing
    Set objProblemOps = Nothing
    Set objDiagnosis = Nothing
    Set objIConvert = Nothing
    
    On Error Resume Next
    Call mDisplayItems      'Clear and Fill mListBox

End Function

Private Function TemplateData_Import(rsString As String) As Boolean
    TemplateData_Import = TemplateData_DeSerialize(rsString)
End Function

Private Function TemplateData_GetQueryCriteria() As String
End Function

Private Sub mDisplayItems()
    Dim i As Integer
    mListBox.Clear
    For i = 0 To mItems.Count - 1
        mListBox.AddItem mItems.Items(i)
        mListBox.ItemData(mListBox.NewIndex) = mItems.Keys(i)
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
    Dim colItems As Collection
    Dim sValue As Variant
    Set colItems = DeserializeString(CStr(rvData), gsListDelim)
    For Each sValue In colItems
        On Error Resume Next
        AddItem CStr(sValue)
        On Error GoTo 0
    Next sValue
    If Not mListBox Is Nothing Then
        Call mDisplayItems
    End If
    mbDirty = False
    TemplateData_DeSerialize = True
End Function

Private Sub TemplateData_DoAdd()
    'SCR-50502 RBELL 2/27/2004 Added message  to handle modal form display
    cwobjCWComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_TEMPLATE_MANAGEMENT
    frmMedcin.DoAdd Me, "Diagnosis Search", mlmDiagnosis_Search
    cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_TEMPLATE_MANAGEMENT
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
    rvData.Add SerializeArray(mItems.Keys, gsListDelim)
    mbDirty = False
End Function
