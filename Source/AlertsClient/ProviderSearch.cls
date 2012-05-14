VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProviderSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Private moClinician As MMMHISClinician.Clinician
Private mbSelected As Boolean
Private moItem As ProviderItem
Private moLogon As Object

Public Function FindItem() As ProviderItem

If Not DebugMode Then On Error GoTo ErrHandler
  
    Set moClinician = New Clinician
    If moClinician.SearchClinician(moLogon.FacilityNCID, "", moLogon, True, False) = True Then
        Set moItem = New ProviderItem
        moItem.Name = moClinician.Name
        moItem.NCID = moClinician.NCID
        
        Dim oConcept As New ConceptCtrl
        oConcept.UniqueId = moItem.NCID
        moItem.UnitNum = oConcept.PrefRep(1309).Representation
        Set oConcept = Nothing

    Else
        Set moItem = Nothing
    End If
    Set FindItem = moItem
    Exit Function

ErrHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "DiagnosisOps.GetDiagnosis", "CHCSII_DiagnosisSearch", vbExclamation)
    Exit Function
End Function


Friend Property Let SetClinic(ByVal oNewValue As Object)
    Rem Set moItem = oNewValue
End Property

Public Function GetByNCID(NCID As String) As CHCSII_Search.ProviderItem

If Not DebugMode Then On Error GoTo ErrHandler
    
    
    Dim oConcept As New ConceptCtrl
    Dim oItem As New CHCSII_Search.ProviderItem
    
    oItem.NCID = NCID
    oConcept.UniqueId = oItem.NCID
    oItem.Name = oConcept.PrefRep(NCID_3M_DEFAULT_CONTEXT).Representation
    If oItem.Name = "" Then
        '-- Temporary for testing
        oItem.Name = "Smithers, Jackson"
        Rem Set oItem = Nothing
        Rem Set oConcept = Nothing
        Rem Exit Function
    End If
    Set GetByNCID = oItem
GoTo CleanUp

ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "DiagnosisOps.GetByNCID", "CHCSII_DiagnosisSearch", vbExclamation)
CleanUp:
    Set oConcept = Nothing
    Set oItem = Nothing
    Exit Function
End Function

Public Function Search(SearchText As String) As Collection

If Not DebugMode Then On Error GoTo ErrHandler

    Exit Function
ErrHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "DiagnosisOps.Search", "CHCSII_DiagnosisSearch", vbExclamation)
    Exit Function
End Function

Private Sub Class_Initialize()
    Set gobjShared = New CWShared
End Sub


Private Sub Class_Terminate()
    Set gobjShared = Nothing
End Sub





Public Property Let LogonObject(ByVal oNewValue As Object)
    Set moLogon = oNewValue
End Property
