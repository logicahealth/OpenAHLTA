VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EMCodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements TemplateData

Private Const MyLabel As String = "EM Category"

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
    TemplateData_Import = TemplateData_DeSerialize(rsString)
End Function

Private Function TemplateData_GetFromEncounter(objEncounter As MMMHIS.Encounter) As Boolean
End Function

Private Function TemplateData_GetQueryCriteria() As String
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
    mMasterDict.Add "1", "TELEPHONE CONSULTS"
    mMasterDict.Add "2", "OFFICE / OUTPATIENT VISIT, EST"
    mMasterDict.Add "3", "OFFICE / OUTPATIENT VISIT, NEW"
    mMasterDict.Add "4", "OFFICE CONSULTATION"
    mMasterDict.Add "5", "CARE PLAN OVERSIGHT SERVICES"
    mMasterDict.Add "6", "CASE MANAGEMENT SERVICES (Team Conferences)"
    mMasterDict.Add "7", "EMERGENCY CARE"
    mMasterDict.Add "8", "OBSERVATION CARE"
    mMasterDict.Add "9", "OTHER E/M SERVICES"
    mMasterDict.Add "10", "PREVENTIVE MEDICINE SERVICES, NEW"
    mMasterDict.Add "11", "PREVENTIVE MEDICINE SERVICES, EST"
    mMasterDict.Add "12", "PREVENTIVE MEDICINE SERVICES, OTHER"
    mMasterDict.Add "13", "SPECIAL E/M SERVICES"
    mMasterDict.Add "14", "STANDBY SERVICES"
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
    Dim colItems As Collection
    Dim sValue As Variant
    Set colItems = DeserializeString(CStr(rvData), gsListDelim)
    For Each sValue In colItems
        On Error Resume Next
        mItems.Add sValue, mMasterDict(sValue)
        On Error GoTo 0
    Next sValue
    If Not mListBox Is Nothing Then
        Call mDisplayItems
    End If
    mDirty = False
    TemplateData_DeSerialize = True
End Function

Private Sub TemplateData_DoAdd()
    mobjComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_TEMPLATE_MANAGEMENT
    frmDictionary.DoAdd Me, mMasterDict, "EM Category Selections"
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
    rvData.Add SerializeArray(mItems.Keys, gsListDelim)
    mDirty = False
End Function

