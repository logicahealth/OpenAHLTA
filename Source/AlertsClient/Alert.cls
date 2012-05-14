VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Alert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"38135DFA0191"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private s_message_guid As String
Private s_message_type_guid As String
Private s_message As Variant
Private s_DisplayMessage As Variant
Private l_message_ien As Long
Private d_time_sent As Date
Private i_priority As Integer
Private b_resolved As Boolean
Private s_provider_ncid As String
Private s_patient_ncid As String
Private s_surrogate_ncid As String
Private s_note As String
Private s_document_ncid As String
Private s_facility_ncid As String
Private b_chcs_clear As Boolean

Public Property Get MessageGuid() As String
    Let MessageGuid = s_message_guid
End Property

Public Property Let MessageGuid(ByVal Value As String)
    Let s_message_guid = Value
End Property

Public Property Get MessageTypeGuid() As String
    Let MessageTypeGuid = s_message_type_guid
End Property

Public Property Let MessageTypeGuid(ByVal Value As String)
    Let s_message_type_guid = Value
End Property

Public Property Get Message() As Variant
    Let Message = s_message
End Property

Public Property Let Message(ByVal Value As Variant)
    Let s_message = Value
End Property

Public Property Get DisplayMessage() As Variant
    Let DisplayMessage = s_DisplayMessage
End Property

Public Property Let DisplayMessage(ByVal Value As Variant)
    Let s_DisplayMessage = Value
End Property

Public Property Get MessageIen() As Long
    Let MessageIen = l_message_ien
End Property

Public Property Let MessageIen(ByVal Value As Long)
    Let l_message_ien = Value
End Property

Public Property Get TimeSent() As Date
    Let TimeSent = d_time_sent
End Property

Public Property Let TimeSent(ByVal Value As Date)
    Let d_time_sent = Value
End Property

Public Property Get Priority() As Integer
    Let Priority = i_priority
End Property

Public Property Let Priority(ByVal Value As Integer)
    Let i_priority = Value
End Property

Public Property Get Resolved() As Boolean
    Let Resolved = b_resolved
End Property

Public Property Let Resolved(ByVal Value As Boolean)
    Let b_resolved = Value
End Property

Public Property Get ProviderNcid() As String
    Let ProviderNcid = s_provider_ncid
End Property

Public Property Let ProviderNcid(ByVal Value As String)
    Let s_provider_ncid = Value
End Property

Public Property Get PatientNcid() As String
    Let PatientNcid = s_patient_ncid
End Property

Public Property Let PatientNcid(ByVal Value As String)
    Let s_patient_ncid = Value
End Property

Public Property Get SurrogateNcid() As String
    Let SurrogateNcid = s_surrogate_ncid
End Property

Public Property Let SurrogateNcid(ByVal Value As String)
    Let s_surrogate_ncid = Value
End Property

Public Property Get Note() As String
    Let Note = s_note
End Property

Public Property Let Note(ByVal Value As String)
    Let s_note = Value
End Property

Public Property Get DocumentNcid() As String
    Let DocumentNcid = s_document_ncid
End Property

Public Property Let DocumentNcid(ByVal Value As String)
    Let s_document_ncid = Value
End Property

Public Property Get FacilityNcid() As String
    Let FacilityNcid = s_facility_ncid
End Property

Public Property Let FacilityNcid(ByVal Value As String)
    Let s_facility_ncid = Value
End Property

Public Property Get Chcs_Clear() As Boolean
    Let Chcs_Clear = b_chcs_clear
End Property

Public Property Let Chcs_Clear(ByVal Value As Boolean)
    Let b_chcs_clear = Value
End Property

Private Sub Class_Initialize()
'
End Sub

Private Sub Class_InitProperties()
'
End Sub

Private Sub Class_ReadProperties(PropBag As PropertyBag)
'
End Sub

Private Sub Class_Terminate()
'
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
'
End Sub
