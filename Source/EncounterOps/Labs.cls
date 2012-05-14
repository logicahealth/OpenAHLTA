VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Labs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mobjEncounter As Encounter
Private mobjParent As SectionParent
Private msRTF As String
Private mbDirty As Boolean
Private mnDataID As Long

Public Property Get RTF() As String
    RTF = mobjEncounter.SectionParent(eAutoCite).mobjSection.LabRTF
End Property

Public Sub Load(ByRef objParent As SectionParent, ByRef colDataStrings As Collection)
    Dim objData As DATASTRATLLib.DataString
    Dim objEvent As CLINTYPEATLLib.ClinicalEvent
    msRTF = ""
    mbDirty = False
    mnDataID = 0
    For Each objData In colDataStrings
        If objData.DataTypeNCID = NCID_ENCEVENT Then
            Set objEvent = objData.PatientData
            If objEvent.Header.ObsBatId.Value.Coded.ValueConcept.UniqueId = SectionEnumToNCID(eLabs) Then
                mnDataID = objData.DataID
                If objEvent.ClinObs.count > 0 Then
                    msRTF = ExtractRTF(objEvent.ClinObs.Item(1))
                End If
                Exit For
            End If
        End If
    Next objData
End Sub

Public Sub OpenEncounter(ByRef objParent As SectionParent, ByRef objEncounter As Encounter)
    Set mobjParent = objParent
    Set mobjEncounter = objEncounter
End Sub

Public Sub CloseEncounter()
    Set mobjParent = Nothing
    Set mobjEncounter = Nothing
End Sub

Public Function Save() As Boolean
' Assume section already locked for update before getting here
    Dim objData As DataString
    On Error GoTo ErrHandler
    If Not mbDirty Then
        Save = True
        Exit Function
    End If
    Set objData = mobjEncounter.InitializeDataString(mobjParent.meSection, mnDataID, Now, Now)
    Call AddRTF(objData.PatientData, msRTF)
    If objData.Save Then
        mnDataID = objData.DataID
        Save = True
        mbDirty = False
    End If
    GoTo Cleanup
ErrHandler:
    Call mobjEncounter.CWShared.ShowVBError(Err.Number, Err.Description, "Labs.Save", "EncounterOps", vbCritical)
Cleanup:
    Set objData = Nothing
    If Save Then
        mobjEncounter.Comm.SetStatusText "Encounter " & mobjEncounter.EncounterID & " " & SectionEnumToName(mobjParent.meSection) & " section was saved."
    Else
        mobjEncounter.Comm.SetStatusText ""
    End If
End Function
