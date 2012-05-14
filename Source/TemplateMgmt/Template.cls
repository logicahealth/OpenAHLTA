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
Attribute VB_Name = "Template"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum TemplateScope
    NoScope = 0
    PersonalScope = 1
    ClinicScope = 2
    MTFScope = 3
    EnterpriseScope = 4
End Enum

Private Const msSerialVersion As String = "V1"

Private mobjEncTmplt As ENCTMPLTATLLib.EncounterTemplate
Private mnVersion As Integer
Private mOwnerNCID As String
Private mOwnerName As String
Private mDirty As Boolean

Public IsShared As Boolean
Public IsActive As Boolean
Public Scope As TemplateScope
Public TemplateName As String
Public TemplateID As String
Public ComplaintList As TemplateData
Public AppointmentList As TemplateData
Public ProblemList As TemplateData
Public AutociteList As TemplateData
Public NotesList As TemplateData
Public ClipartList As TemplateData
Public DiagnosisList As TemplateData
Public OrderList As TemplateData
Public EMCodeList As TemplateData
Public ProcedureList As TemplateData
Public SpecialtyList As TemplateData

Public Property Get isDirty() As Boolean
    If mDirty Then
        isDirty = True
    ElseIf ComplaintList.isDirty Then
        isDirty = True
    ElseIf AppointmentList.isDirty Then
        isDirty = True
    ElseIf ProblemList.isDirty Then
        isDirty = True
    ElseIf AutociteList.isDirty Then
        isDirty = True
    ElseIf NotesList.isDirty Then
        isDirty = True
    ElseIf ClipartList.isDirty Then
        isDirty = True
    ElseIf DiagnosisList.isDirty Then
        isDirty = True
    ElseIf OrderList.isDirty Then
        isDirty = True
    ElseIf EMCodeList.isDirty Then
        isDirty = True
    ElseIf ProcedureList.isDirty Then
        isDirty = True
    ElseIf SpecialtyList.isDirty Then
        isDirty = True
    End If
End Property

Public Property Get OwnerNcid() As String
    OwnerNcid = mOwnerNCID
End Property

Public Property Let OwnerNcid(NCID As String)
    Dim objConcept As CONCEPTATLLib.ConceptCtrl
    mOwnerNCID = NCID
    If NCID = "" Then
        mOwnerName = ""
    Else
        Set objConcept = New CONCEPTATLLib.ConceptCtrl
        objConcept.UniqueId = NCID
        mOwnerName = objConcept.PrefRep("2000").Representation
    End If
    mDirty = True
End Property

Public Property Get OwnerName() As String
    OwnerName = mOwnerName
End Property

Public Function Delete() As Boolean
    On Error GoTo ErrHandler
    If mobjEncTmplt Is Nothing Then
        MsgBxARMd "Error: attempt to delete Encounter Template not properly initialized."
    Else
        mobjEncTmplt.Delete
    End If
    Delete = True
    Exit Function
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Template.Delete()", "TemplateClient", vbExclamation)
End Function

Public Function Deserialize(robjEncTemplate As ENCTMPLTATLLib.EncounterTemplate) As Boolean
    Dim colList As Collection
    Dim i As Integer
    Set mobjEncTmplt = robjEncTemplate
    mobjEncTmplt.Read
    TemplateID = robjEncTemplate.TemplateID
    OwnerNcid = robjEncTemplate.Owner
    IsActive = robjEncTemplate.IsActive
    Select Case robjEncTemplate.Type
    Case tt_facility
        If robjEncTemplate.IsShared Then
            Scope = ClinicScope
        Else
            Scope = MTFScope
        End If
        IsShared = False
    Case tt_enterprise
        IsShared = False
        Scope = EnterpriseScope
    Case Else
        IsShared = robjEncTemplate.IsShared
        Scope = PersonalScope
    End Select
    TemplateName = robjEncTemplate.Name
    ' deserialize the cross-reference type lists
    Call AppointmentList.Deserialize(robjEncTemplate)
    Call ComplaintList.Deserialize(robjEncTemplate)
    Call ProblemList.Deserialize(robjEncTemplate)
    Call SpecialtyList.Deserialize(robjEncTemplate)
    
    'deserialize the encounter data lists
    For i = 1 To mobjEncTmplt.Items.Count
        If mobjEncTmplt.Items(i).Type = dt_blob Then
            Set colList = DeserializeString(mobjEncTmplt.Items(i).dataBlob, gsMainDelim)
            Exit For
        End If
    Next i
    If Not colList Is Nothing Then
        If colList.Count > 0 Then
            If colList(1) = msSerialVersion Then
                Call ClipartList.Deserialize(colList(2))
                Call NotesList.Deserialize(colList(3))
                Call DiagnosisList.Deserialize(colList(4))
                Call OrderList.Deserialize(colList(5))
                Call EMCodeList.Deserialize(colList(6))
                Call ProcedureList.Deserialize(colList(7))
                If colList.Count >= 8 Then
                    Call AutociteList.Deserialize(colList(8))
                End If
            Else
                MsgBxARMd "Unexpected template serialization version '" & colList(1) & "'."
            End If
        End If
        Set colList = Nothing
    End If
    mDirty = False
    Deserialize = True
End Function

Public Function Save() As Boolean
    Dim colList As Collection
    Dim objItem As ENCTMPLTATLLib.ETItem
    On Error GoTo ErrHandler
    If mobjEncTmplt Is Nothing Then
        Set mobjEncTmplt = New ENCTMPLTATLLib.EncounterTemplate
        mobjEncTmplt.ClearAll
    Else
        mobjEncTmplt.Items.ClearAll
        mobjEncTmplt.References.ClearAll
    End If
    With mobjEncTmplt
        .Name = TemplateName
        .Owner = mOwnerNCID
        .IsActive = IsActive
        Select Case Scope
        Case ClinicScope
            .Type = tt_facility
            .IsShared = True
        Case MTFScope
            .Type = tt_facility
            .IsShared = False
        Case EnterpriseScope
            .Type = tt_enterprise
            .IsShared = False
        Case Else
            .Type = tt_user
            .IsShared = IsShared
        End Select
    End With
    'serialize the cross-reference type lists
    Call AppointmentList.Serialize(mobjEncTmplt)
    Call ComplaintList.Serialize(mobjEncTmplt)
    Call ProblemList.Serialize(mobjEncTmplt)
    Call SpecialtyList.Serialize(mobjEncTmplt)
    'serialize the encounter data types
    Set colList = New Collection
    colList.Add msSerialVersion
    Call ClipartList.Serialize(colList)
    Call NotesList.Serialize(colList)
    Call DiagnosisList.Serialize(colList)
    Call OrderList.Serialize(colList)
    Call EMCodeList.Serialize(colList)
    Call ProcedureList.Serialize(colList)
    Call AutociteList.Serialize(colList)
    mobjEncTmplt.Items.Add
    mobjEncTmplt.Items(mobjEncTmplt.Items.Count).Type = dt_blob
    mobjEncTmplt.Items(mobjEncTmplt.Items.Count).dataBlob = SerializeCollection(colList, gsMainDelim)
    mobjEncTmplt.Save
    Save = True
    mDirty = False
    Exit Function
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Template.Save()", "TemplateClient", vbExclamation)
End Function

Private Sub Class_Initialize()
    Set ComplaintList = New Complaints
    Set AppointmentList = New AppointmentTypes
    Set ProblemList = New Problems
    Set AutociteList = New AutoCites
    Set NotesList = New NoteTemplates
    Set ClipartList = New ClipArts
    Set DiagnosisList = New Diagnoses
    Set OrderList = New Orders
    Set EMCodeList = New EMCodes
    Set ProcedureList = New Procedures
    Set SpecialtyList = New Specialties
    IsActive = True
End Sub

Private Sub Class_Terminate()
    Set ComplaintList = Nothing
    Set AppointmentList = Nothing
    Set ProblemList = Nothing
    Set AutociteList = Nothing
    Set NotesList = Nothing
    Set ClipartList = Nothing
    Set DiagnosisList = Nothing
    Set OrderList = Nothing
    Set EMCodeList = Nothing
    Set ProcedureList = Nothing
    Set SpecialtyList = Nothing
    If Not mobjEncTmplt Is Nothing Then
        mobjEncTmplt.ClearAll
        Set mobjEncTmplt = Nothing
    End If
End Sub

