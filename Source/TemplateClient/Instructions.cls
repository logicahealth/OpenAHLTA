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
Attribute VB_Name = "Instructions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements TemplateData

'Private Const MyLabel As String = "Other Therapies"
Private Const MyLabel As String = "Patient Instructions"

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
    
    Dim objSQL As ICHCSII_SQL.ISqlOps
    Dim moConn As CHCSII_CONN.Conn
    
    Set moConn = New CHCSII_CONN.Conn
    Set objSQL = moConn.CHCSII_SQLOPS(Auto)
    
    objSQL.Execute "select SnoID from pi_data" _
               & " where enc_fac_NCID = '" & objEncounter.FacilityNCID & "'" _
               & " and enc_num = '" & objEncounter.EncounterID & "'"
    On Error GoTo ErrHandler
    Do While Not objSQL.EOF
        AddItem objSQL("SnoID")
        objSQL.MoveNext
    Loop
    TemplateData_GetFromEncounter = True
    GoTo Cleanup
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Instructions.GetFromEncounter", "TemplateClient", vbExclamation)
Cleanup:
    Set objSQL = Nothing
    Set moConn = Nothing
    
End Function

Private Function TemplateData_Import(rsString As String) As Boolean
    TemplateData_Import = TemplateData_DeSerialize(rsString) '<SCR 37645
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
    frmMedcin.DoAdd Me, "Patient Instructions", mlmPatient_Instruction_Search
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

