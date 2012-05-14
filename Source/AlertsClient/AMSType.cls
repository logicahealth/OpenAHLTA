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
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AMSType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"38135DF9023A"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private sMessageType As String
Private sInteraction As String
Private lPriority As Long
Private sLabel As String
Private bMultiplicity As Boolean

Public Property Get MessageType() As String
    MessageType = sMessageType
End Property

Public Property Let MessageType(ByVal Value As String)
    sMessageType = Value
End Property

Public Property Get Interaction() As String
    Interaction = sInteraction
End Property

Public Property Let Interaction(ByVal Value As String)
    sInteraction = Value
End Property

Public Property Get Priority() As Long
    Priority = lPriority
End Property

Public Property Let Priority(ByVal Value As Long)
    lPriority = Value
End Property

Public Property Get Label() As String
    Label = sLabel
End Property

Public Property Let Label(ByVal Value As String)
    sLabel = Value
End Property

Public Property Get Multiplicity() As Boolean
    Multiplicity = bMultiplicity
End Property

Public Property Let Multiplicity(ByVal Value As Boolean)
    bMultiplicity = Value
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
