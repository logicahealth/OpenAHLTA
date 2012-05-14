VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BroadcastMessageHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_colListeners As New Collection

Public Sub AddListenerForMessage(DestinationServer As Object, ByVal SourceServerNCID As String, ByVal MessageType As Integer)
    
Dim objListenerEntry As BroadcastMessageEntry

'Make sure it is not a duplicate
For Each objListenerEntry In m_colListeners
    If (objListenerEntry.MessageType = MessageType) Then
        If (objListenerEntry.SourceServerNCID = SourceServerNCID) Then
            Exit Sub             'Duplicate
        End If
    End If
Next
    
Set objListenerEntry = New BroadcastMessageEntry

With objListenerEntry
    Set .DestinationServer = DestinationServer
    .SourceServerNCID = SourceServerNCID
    .MessageType = MessageType
End With
   
m_colListeners.Add objListenerEntry
    
End Sub

Public Sub RemoveListenerForMessage(ByVal SourceServerNCID As String, ByVal MessageType As Integer)

Dim objListenerEntry As BroadcastMessageEntry
Dim iIndex As Integer

For Each objListenerEntry In m_colListeners
    iIndex = iIndex + 1
    If (objListenerEntry.MessageType = MessageType And objListenerEntry.SourceServerNCID = SourceServerNCID) Then
        m_colListeners.Remove iIndex
    End If
Next

End Sub

Public Sub Message(ByVal MessageType As Integer, ByVal MessageText As String, ByVal SourceServerNCID As String)

Dim objListenerEntry As BroadcastMessageEntry

For Each objListenerEntry In m_colListeners
    If (objListenerEntry.MessageType = MessageType And objListenerEntry.SourceServerNCID = SourceServerNCID) Then
        objListenerEntry.DestinationServer.Message MessageType, MessageText, SourceServerNCID
    End If
Next

End Sub

