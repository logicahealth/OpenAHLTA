VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MyEvntBridge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'<< SCR #29069;   Developer: Brian Mowbray 11/11/2002 06:00 PM
'This wrapper class was added to prevent duplicate instatiation of the event bridge
Option Explicit

Public WithEvents EvBridge As EVHBRIDGEMODLib.EvhBridge
Attribute EvBridge.VB_VarHelpID = -1

Private Sub Class_Initialize()

    On Error Resume Next
    If EvBridge Is Nothing Then
       Set EvBridge = New EVHBRIDGEMODLib.EvhBridge
    End If
    
End Sub

Private Sub EvBridge_HandleXMLEvent(ByVal EventID As Long, ByVal EventType As Long, ByVal EventAction As Long, ByVal RequesterSubscription As Long, ByVal XMLString As String)
   Call gEvntSvr.RcvMessage("TMIP_EVENT", XMLString, ED_DMBE)
End Sub

Public Sub SendMyXMLRequest(lEventType As Long, sMsg As String, lEventID As Long)
    EvBridge.SendXMLRequest lEventType, sMsg, lEventID
End Sub

Public Sub SendMyXMLResponse(lEventType As Long, lSubscription As Long, sMsg As String, lEventID As Long)
    EvBridge.SendXMLResponse lEventType, lSubscription, sMsg, lEventID
End Sub

Public Function MySubscribe(lEventType As Long, sModuleName As String) As Boolean
    On Error Resume Next
    EvBridge.Subscribe lEventType, sModuleName
    If Err.number <> 0 Then
        Err.Clear
        Exit Function
    End If
    MySubscribe = True
End Function
