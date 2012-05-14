VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPSR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private oSingleton As New srvSVGComm.SingletonFactory
Public oSVGComm As srvSVGComm.clsSVGComm

Public Property Let upperRight(sUR As String)
   oSVGComm.mObjPSR.upperRight = sUR
End Property
Public Property Get upperRight() As String
   upperRight = oSVGComm.mObjPSR.upperRight
End Property
Public Property Let upperCenter(sUC As String)
    oSVGComm.mObjPSR.upperCenter = sUC
End Property
Public Property Get upperCenter() As String
    upperCenter = oSVGComm.mObjPSR.upperCenter
End Property
Public Property Let upperLeft(sUL As String)
    oSVGComm.mObjPSR.upperLeft = sUL
End Property
Public Property Get upperLeft() As String
    upperLeft = oSVGComm.mObjPSR.upperLeft
End Property
Public Property Let lowerRight(sLR As String)
    oSVGComm.mObjPSR.lowerRight = sLR
End Property
Public Property Get lowerRight() As String
    lowerRight = oSVGComm.mObjPSR.lowerRight
End Property
Public Property Let lowerCenter(sLC As String)
    oSVGComm.mObjPSR.lowerCenter = sLC
End Property
Public Property Get lowerCenter() As String
    lowerCenter = oSVGComm.mObjPSR.lowerCenter
End Property
Public Property Let lowerLeft(sLL As String)
    oSVGComm.mObjPSR.lowerLeft = sLL
End Property
Public Property Get lowerLeft() As String
    lowerLeft = oSVGComm.mObjPSR.lowerLeft
End Property
Public Property Let sNotes(sNotes As String)
    oSVGComm.mObjPSR.sNotes = sNotes
End Property
Public Property Get sNotes() As String
    sNotes = oSVGComm.mObjPSR.sNotes
End Property
Public Property Let bDataChanged(bDChanged As Boolean)
    oSVGComm.mObjPSR.bDataChanged = bDChanged
End Property
Public Property Get bDataChanged() As Boolean
    bDataChanged = oSVGComm.mObjPSR.bDataChanged
End Property


Private Sub Class_Initialize()
    'Set to svgcomm server when loaded
    Set oSVGComm = oSingleton.Singleton

End Sub

Private Sub Class_Terminate()

    Set oSVGComm = Nothing
    Set oSingleton = Nothing

End Sub


