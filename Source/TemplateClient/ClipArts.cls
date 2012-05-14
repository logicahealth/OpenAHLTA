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
Attribute VB_Name = "ClipArts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements TemplateData

Private mItems As Scripting.Dictionary
Private mListBox As VB.ListBox
Private mbDirty As Boolean

Private oEMCategoryDictionary As Scripting.Dictionary
Private oSpecialtyDictionary As Scripting.Dictionary


Public Property Get EMCategoryDictionary() As Scripting.Dictionary
    On Error GoTo ErrHandler
    Set EMCategoryDictionary = oEMCategoryDictionary
    Exit Property
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.ClipArts.Get_EMCategory", Err.Description
End Property

Public Property Set EMCategoryDictionary(ByVal RHV As Scripting.Dictionary)
    On Error GoTo ErrHandler
    Set oEMCategoryDictionary = RHV
    Exit Property
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.ClipArts.Get_EMCategory", Err.Description
End Property

Public Property Get SpecialtyDictionary() As Scripting.Dictionary
    On Error GoTo ErrHandler
    Set SpecialtyDictionary = oSpecialtyDictionary
    Exit Property
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.ClipArts.Get_Specialty", Err.Description
End Property

Public Property Set SpecialtyDictionary(ByVal RHV As Scripting.Dictionary)
    On Error GoTo ErrHandler
    Set oSpecialtyDictionary = RHV
    Exit Property
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.ClipArts.Set_Specialty", Err.Description
End Property

Private Sub Class_Terminate()
    On Error Resume Next
    Set mItems = Nothing
    Set mListBox = Nothing
    Set oEMCategoryDictionary = Nothing
    Set oSpecialtyDictionary = Nothing
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
            'mDisplayItems
            mListBox.Selected(i + 1) = True
            mbDirty = True
        End If
    Next i

    Exit Sub
    
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.ClipArts.DoMoveDown", Err.Description

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
            'mDisplayItems
            mListBox.Selected(i - 1) = True
            mbDirty = True
            Exit For
        End If
    Next i

    Exit Sub
    
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.ClipArts.DoMoveUp", Err.Description

End Sub


Private Function TemplateData_Export() As String
    On Error GoTo ErrHandler
    Exit Function
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.ClipArts.Expotr", Err.Description

End Function

Private Function TemplateData_Import(rsString As String) As Boolean
    On Error GoTo ErrHandler
    TemplateData_Import = True
    Exit Function
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.ClipArts.Import", Err.Description
End Function

Private Function TemplateData_GetFromEncounter(objEncounter As CHCSIIEncounterOps.Encounter) As Boolean
    On Error GoTo ErrHandler
    Exit Function
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.ClipArts.GetFromEncounter", Err.Description
End Function

Private Function TemplateData_GetQueryCriteria() As String
    On Error GoTo ErrHandler
    Exit Function
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.ClipArts.GetQueryCriteria", Err.Description
End Function

Private Property Get TemplateData_CanEdit() As Boolean
    On Error GoTo ErrHandler
    TemplateData_CanEdit = False
    Exit Property
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.ClipArts.CanEdit", Err.Description
End Property

Private Function TemplateData_DeSerialize(rvData As Variant) As Boolean
    On Error GoTo ErrHandler
    Exit Function
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.ClipArts.DeSerialize", Err.Description
    
End Function

Private Sub TemplateData_DoAdd()
    On Error GoTo ErrHandler
    MsgBxARMd "Need Source for finding names of available clip art."
    Exit Sub
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.ClipArts.DoAdd", Err.Description
End Sub

Private Sub TemplateData_DoEdit()
    On Error GoTo ErrHandler
    Exit Sub
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.ClipArts.DoEdit", Err.Description

End Sub

Private Sub TemplateData_DoRemove()
    On Error GoTo ErrHandler
    Exit Sub
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.ClipArts.DoRemove", Err.Description

End Sub

Private Property Let TemplateData_isDirty(RHS As Boolean)
    On Error GoTo ErrHandler
    mbDirty = RHS
    Exit Property
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.ClipArts.IsDirty", Err.Description
End Property

Private Property Get TemplateData_isDirty() As Boolean
    On Error GoTo ErrHandler
    TemplateData_isDirty = mbDirty
    Exit Property
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.ClipArts.IsDirty", Err.Description
End Property

Private Property Get TemplateData_Items() As Scripting.IDictionary
    On Error GoTo ErrHandler
    Exit Property
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.ClipArts.Items", Err.Description

End Property

Private Property Get TemplateData_Label() As String
    On Error GoTo ErrHandler
    TemplateData_Label = "Associated Clipart"
    Exit Property
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.ClipArts.Label", Err.Description
End Property

Private Property Set TemplateData_ListBox(RHS As Object)
    On Error GoTo ErrHandler
    Set mListBox = RHS
    Exit Property
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.ClipArts.ListBox", Err.Description
End Property

Private Property Get TemplateData_ListBox() As Object
    On Error GoTo ErrHandler
    Exit Property
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.ClipArts.ListBox", Err.Description
End Property

Private Function TemplateData_Serialize(rvData As Variant) As Boolean
    On Error GoTo ErrHandler
    rvData.Add ""
    mbDirty = False
    Exit Function
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.ClipArts.Serialize", Err.Description
End Function

