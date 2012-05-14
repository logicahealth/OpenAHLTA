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
Attribute VB_Name = "Specialties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements TemplateData

Private Const StorageNCID As String = "291"
Private Const MyLabel As String = "Associated Specialties"

Private mListBox As VB.ListBox
Private mItems As Dictionary
Private mDirty As Boolean
Private mobjConcept As CONCEPTATLLib.ConceptCtrl
Private mobjShared As MMMHISShared.CWShared
Private mobjPatient As MMMHISPatient.Patient
Private mobjComm As Object
Private mobjUser As CHCSII_ClientSecurity.ICHCSII_User
Private mobjLogon As mmmhislogon.Logon
Private mlProgId As Long

Private mMasterDict As Dictionary

Public Sub AddItem(ByRef rsID As String, ByRef rsDesc As String)
    On Error Resume Next
    mItems.Add rsID, rsDesc
    mDirty = True
    mDisplayItems
End Sub

Private Function TemplateData_Export() As String
    TemplateData_Export = SerializeArray(mItems.Keys, gsListDelim)
End Function

Private Function TemplateData_Import(rsString As String) As Boolean
    Dim colItems As Collection
    Dim sValue As Variant
    Set colItems = DeserializeString(rsString, gsListDelim)
    For Each sValue In colItems
        On Error Resume Next
        mItems.Add sValue, mMasterDict(sValue)
        On Error GoTo 0
    Next sValue
    If Not mListBox Is Nothing Then
        Call mDisplayItems
    End If
    mDirty = False
    TemplateData_Import = True
End Function

Private Function TemplateData_GetFromEncounter(objEncounter As MMMHIS.Encounter) As Boolean
End Function

Private Function TemplateData_GetQueryCriteria() As String
    Dim i As Integer
    Dim sStr As String
    For i = 0 To mItems.Count - 1
        If sStr <> "" Then
            sStr = sStr & " or"
        End If
        sStr = sStr & "( type_ncid = " & StorageNCID & " and value_ncid = " & mItems.Keys(i) & ")"
    Next i
    TemplateData_GetQueryCriteria = sStr
End Function

Private Sub TemplateData_Initialize(robjLogon As mmmhislogon.Logon, robjUser As CHCSII_ClientSecurity.ICHCSII_User, robjShared As MMMHISShared.CWShared, robjConcept As CONCEPTATLLib.IConceptCtrl, robjPatient As MMMHISPatient.Patient, robjComm As Object, vlProgId As Variant)
    Set mobjLogon = robjLogon
    Set mobjUser = robjUser
    Set mobjShared = robjShared
    Set mobjConcept = robjConcept
    Set mobjPatient = robjPatient
    Set mobjComm = robjComm
    mlProgId = vlProgId
End Sub

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
    Set mMasterDict = New Dictionary
    mMasterDict.Add "1", "Aerospace/Aviation Medicine"
    mMasterDict.Add "2", "Allergy/Immunology"
    mMasterDict.Add "3", "Cardiology"
    mMasterDict.Add "4", "Cardiothoracic Surgery"
    mMasterDict.Add "5", "Dermatology"
    mMasterDict.Add "6", "Endocrinology"
    mMasterDict.Add "7", "Family Medicine"
    mMasterDict.Add "8", "Gastroenterology"
    mMasterDict.Add "9", "General Surgery"
    mMasterDict.Add "10", "Gynecology"
    mMasterDict.Add "11", "Hematology"
    mMasterDict.Add "12", "Hyperbaric/Undersea Medicine"
    mMasterDict.Add "13", "Infectious Disease"
    mMasterDict.Add "14", "Nephrology"
    mMasterDict.Add "15", "Neurology"
    mMasterDict.Add "16", "Neurosurgery"
    mMasterDict.Add "17", "Obstetrics/Gynecology"
    mMasterDict.Add "18", "Occupational and Environmental Medicine"
    mMasterDict.Add "19", "Oncology"
    mMasterDict.Add "20", "Ophthalmology"
    mMasterDict.Add "21", "Orthopedics"
    mMasterDict.Add "22", "Otorhinolaryngology"
    mMasterDict.Add "23", "Pediatrics"
    mMasterDict.Add "24", "Preventive Medicine"
    mMasterDict.Add "25", "Psychiatric Disorders"
    mMasterDict.Add "26", "Pulmonary Medicine"
    mMasterDict.Add "27", "Rheumatology"
    mMasterDict.Add "28", "Tropical Medicine"
    mMasterDict.Add "29", "Urology"
End Sub

Private Sub Class_Terminate()
    Set mItems = Nothing
    Set mListBox = Nothing
    Set mobjLogon = Nothing
    Set mobjUser = Nothing
    Set mobjConcept = Nothing
    Set mobjShared = Nothing
    Set mobjComm = Nothing
    Set mobjPatient = Nothing
    Set mMasterDict = Nothing
End Sub

Private Property Get TemplateData_CanEdit() As Boolean
    TemplateData_CanEdit = False
End Property

Private Function TemplateData_DeSerialize(rvData As Variant) As Boolean
    Dim objRef As ENCTMPLTATLLib.ETXref
    For Each objRef In rvData.References
        If objRef.Type = StorageNCID Then
            mItems.Add objRef.Value, mMasterDict(objRef.Value)
        End If
    Next objRef
    If Not mListBox Is Nothing Then
        Call mDisplayItems
    End If
    mDirty = False
End Function

Private Sub TemplateData_DoAdd()
    mobjComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_TEMPLATE_MANAGEMENT
    frmDictionary.DoAdd Me, mMasterDict, "Specialties"
    mobjComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_TEMPLATE_MANAGEMENT
End Sub

Private Sub TemplateData_DoEdit()

End Sub

Private Sub TemplateData_DoRemove()
    Dim i As Integer
    Dim sNCID As String
    sNCID = CStr(mListBox.ItemData(mListBox.ListIndex))
    mItems.Remove sNCID
    mListBox.RemoveItem mListBox.ListIndex
    mDirty = True
End Sub

Private Property Let TemplateData_isDirty(RHS As Boolean)
    mDirty = RHS
End Property

Private Property Get TemplateData_isDirty() As Boolean
    TemplateData_isDirty = mDirty
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
    Dim i As Integer
    Dim objRef As ENCTMPLTATLLib.ETXref
    For i = 0 To mItems.Count - 1
        Set objRef = rvData.References.Add
        objRef.Type = StorageNCID
        objRef.Value = mItems.Keys(i)
    Next i
    Set objRef = Nothing
    TemplateData_Serialize = True
    mDirty = False
End Function

