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

Implements TemplateData

Private Const MyLabel As String = "Items to Autocite into Note"

Private mListBox As VB.ListBox
Private mItems As Dictionary ' key is serial, item is encautocite
Private mbDirty As Boolean

Private mMasterDict As Dictionary
Private oEMCategoryDictionary As Scripting.Dictionary
Private oSpecialtyDictionary As Scripting.Dictionary

Private Const NCID_EXPIRED_MEDICATIONS               As String = "1999989"

Public Property Get EMCategoryDictionary() As Scripting.Dictionary
    On Error GoTo ErrHandler
    Set EMCategoryDictionary = oEMCategoryDictionary
    Exit Property
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.AutoCites.Get_EMCategory", Err.Description
End Property

Public Property Set EMCategoryDictionary(ByVal RHV As Scripting.Dictionary)
    On Error GoTo ErrHandler
    Set oEMCategoryDictionary = RHV
    Exit Property
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.AutoCites.Set_EMCategory", Err.Description
End Property

Public Property Get SpecialtyDictionary() As Scripting.Dictionary
    On Error GoTo ErrHandler
    Set SpecialtyDictionary = oSpecialtyDictionary
    Exit Property
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.AutoCites.Get_Specialty", Err.Description
End Property

Public Property Set SpecialtyDictionary(ByVal RHV As Scripting.Dictionary)
    On Error GoTo ErrHandler
    Set oSpecialtyDictionary = RHV
    Exit Property
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.AutoCites.Set_Specialty", Err.Description
End Property


Public Sub AddItem(ByRef rsID As String, ByRef rsDesc As String)
    
    On Error GoTo ErrHandler
    mItems.Add rsID, rsDesc
    mbDirty = True
    mDisplayItems
    Exit Sub
    
ErrHandler:
    If Err = 457 Then
       mItems.Remove rsID
       Resume
    End If
    Err.Raise Err.Number, "CHCSII_TemplateClient.AutoCites.AddItem", Err.Description
End Sub

Private Sub TemplateData_DoMoveDown()
    
    Dim sNCIDFrom As String
    Dim sDescFrom As String
    Dim sNCIDTo As String
    Dim sDescTo As String
    Dim i As Integer
    
    On Error GoTo ErrHandler
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

    Exit Sub
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.AutoCites.DoMoveDown", Err.Description

End Sub

Private Sub TemplateData_DoMoveUp()
    
    Dim sNCIDFrom As String
    Dim sDescFrom As String
    Dim sNCIDTo As String
    Dim sDescTo As String
    Dim i As Integer
    
    On Error GoTo ErrHandler
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

    Exit Sub
    
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.AutoCites.DoMoveUp", Err.Description

End Sub


Private Function TemplateData_Export() As String
    On Error GoTo ErrHandler
    TemplateData_Export = SerializeArray(mItems.Keys, gsListDelim)
    Exit Function
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.AutoCites.Export", Err.Description
End Function

Private Function TemplateData_GetFromEncounter(objEncounter As CHCSIIEncounterOps.Encounter) As Boolean
    
    On Error GoTo ErrHandler
    
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.AutoCites.GetFromEncounter", Err.Description

End Function

Private Function TemplateData_Import(rsString As String) As Boolean
    On Error GoTo ErrHandler
    TemplateData_Import = TemplateData_DeSerialize(rsString)
    Exit Function
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.AutoCites.Import", Err.Description
End Function

Private Function TemplateData_GetQueryCriteria() As String
    
    On Error GoTo ErrHandler
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.AutoCites.GetQueryCriteria", Err.Description
    
End Function

Private Sub mDisplayItems()
    
    Dim i As Integer
    
    On Error GoTo ErrHandler
    mListBox.Clear
    For i = 0 To mItems.Count - 1
        mListBox.AddItem mItems.Items(i)
        mListBox.ItemData(mListBox.NewIndex) = mItems.Keys(i)
    Next i
    Exit Sub
    
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.AutoCites.mDisplayItems", Err.Description

End Sub

Private Sub Class_Initialize()
    
    On Error Resume Next
    Set mItems = New Dictionary
    Set mMasterDict = New Dictionary
    
    If gobjShared.IsAppMode(modeCHCSI_GUI) Then   '<SCR 36493
        mMasterDict.Add CStr(NCID_MEDICATIONS_APPLICATION), "Active Medications"
        mMasterDict.Add CStr(NCID_PROBLEMS_APPLICATION), "Active Problems"
        mMasterDict.Add CStr(NCID_ALLERGIES_APPLICATION), "Allergies"
        mMasterDict.Add CStr(NCID_EXPIRED_MEDICATIONS), "Expired Dispensed Medications"
        mMasterDict.Add CStr(NCID_LAB_APPLICATION), "Labs"
        mMasterDict.Add CStr(NCID_RADIOLOGY_RESULTS_APPLICATION), "Rads"
    ElseIf gobjShared.IsAppMode(modeTheater) Then
        'SCR-43860 10/14/03 RBELL Added theater specific list of items for autociting rather than defaults
        mMasterDict.Add CStr(NCID_MEDICATIONS_APPLICATION), "Active Medications"
        mMasterDict.Add CStr(NCID_PROBLEMS_APPLICATION), "Active Problems"
        mMasterDict.Add CStr(NCID_ALLERGIES_APPLICATION), "Allergies"
        mMasterDict.Add CStr(NCID_VITAL_SIGNS_APPLICATION), "Vitals"
    Else
        mMasterDict.Add CStr(NCID_MEDICATIONS_APPLICATION), "Active Medications"
        mMasterDict.Add CStr(NCID_PROBLEMS_APPLICATION), "Active Problems"
        mMasterDict.Add CStr(NCID_ALLERGIES_APPLICATION), "Allergies"
        If cwobjCWComm.CmdLineSwitch("CPG") <> "" Then mMasterDict.Add CStr(NCID_DEVICEREADING_APPLICATION), "Device Readings"
        mMasterDict.Add CStr(NCID_EXPIRED_MEDICATIONS), "Expired Dispensed Medications"
        mMasterDict.Add CStr(NCID_FAMILY_HISTORY_APPLICATION), "Family History"
        If cwobjCWComm.CmdLineSwitch("CPG") <> "" Then mMasterDict.Add CStr(NCID_PATIENT_PLANOFCARE), "Goal Summary"
        mMasterDict.Add CStr(NCID_LAB_APPLICATION), "Labs"
        If cwobjCWComm.CmdLineSwitch("CPG") <> "" Then mMasterDict.Add CStr(NCID_EDUCATION), "P/E Handouts"
        mMasterDict.Add CStr(NCID_PATIENT_QUESTIONNAIRES), "Questionnaires"
        mMasterDict.Add CStr(NCID_RADIOLOGY_RESULTS_APPLICATION), "Rads"
        mMasterDict.Add CStr(NCID_VITAL_SIGNS_APPLICATION), "Vitals"
    End If
    
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Set mItems = Nothing
    Set mListBox = Nothing
    Set mMasterDict = Nothing
    Set oEMCategoryDictionary = Nothing
    Set oSpecialtyDictionary = Nothing
End Sub

Private Property Get TemplateData_CanEdit() As Boolean
    On Error GoTo ErrHandler
    TemplateData_CanEdit = False
    Exit Property
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.AutoCites.CanEdit", Err.Description
End Property

Private Function TemplateData_DeSerialize(rvData As Variant) As Boolean
    
    Dim colItems As Collection
    Dim sValue As Variant
    
    On Error GoTo ErrHandler
    Set colItems = DeserializeString(CStr(rvData), gsListDelim)
    For Each sValue In colItems
        mItems.Add CStr(sValue), mMasterDict(sValue)
    Next sValue
    If Not mListBox Is Nothing Then
        Call mDisplayItems
    End If
    mbDirty = False
    TemplateData_DeSerialize = True
    Exit Function
    
ErrHandler:
    If Err = 457 Then
       mItems.Remove sValue
       Resume
    End If
    Err.Raise Err.Number, "CHCSII_TemplateClient.AutoCites.DeSerialize", Err.Description

End Function

Private Sub TemplateData_DoAdd()
    On Error GoTo ErrHandler
    cwobjCWComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_TEMPLATE_MANAGEMENT
    frmDictionary.DoAdd Me, mMasterDict, "Auto Cite Selections"
    cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_TEMPLATE_MANAGEMENT
    Exit Sub
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.AutoCites.DoAdd", Err.Description
End Sub

Private Sub TemplateData_DoEdit()
    
    On Error GoTo ErrHandler
    
    Exit Sub
    
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.AutoCites.DoEdit", Err.Description

End Sub

Private Sub TemplateData_DoRemove()
    
    Dim sNCID As String
    Dim i As Integer
    
    On Error GoTo ErrHandler
    For i = mListBox.ListCount - 1 To 0 Step -1
        If mListBox.Selected(i) Then
            sNCID = CStr(mListBox.ItemData(i))
            mItems.Remove sNCID
            mListBox.RemoveItem i
            mbDirty = True
        End If
    Next i

    Exit Sub

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.AutoCites.DoRemove", Err.Description

End Sub

Private Property Let TemplateData_isDirty(RHS As Boolean)
    On Error GoTo ErrHandler
    mbDirty = RHS
    Exit Property
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.AutoCites.IsDirty", Err.Description
End Property

Private Property Get TemplateData_isDirty() As Boolean
    On Error GoTo ErrHandler
    TemplateData_isDirty = mbDirty
    Exit Property
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.AutoCites.IsDirty", Err.Description
End Property

Private Property Get TemplateData_Items() As Scripting.IDictionary
    On Error GoTo ErrHandler
    Set TemplateData_Items = mItems
    Exit Property
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.AutoCites.Items", Err.Description
End Property

Private Property Get TemplateData_Label() As String
    On Error GoTo ErrHandler
    TemplateData_Label = MyLabel
    Exit Property
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.AutoCites.Label", Err.Description
End Property

Private Property Set TemplateData_ListBox(RHS As Object)
    On Error GoTo ErrHandler
    Set mListBox = RHS
    Call mDisplayItems
    Exit Property
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.AutoCites.ListBox", Err.Description
End Property

Private Property Get TemplateData_ListBox() As Object
    On Error GoTo ErrHandler
    Exit Property
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.AutoCites.ListBox", Err.Description
End Property

Private Function TemplateData_Serialize(rvData As Variant) As Boolean
    On Error GoTo ErrHandler
    rvData.Add SerializeArray(mItems.Keys, gsListDelim)
    mbDirty = False
    Exit Function
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.AutoCites.Serialize", Err.Description
End Function
