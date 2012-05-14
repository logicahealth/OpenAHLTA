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
Attribute VB_Name = "MetaData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"Disposition"
Option Explicit
Implements IPersist

'Persistence of Controls
Public Disposition As String
Public InjuryIllness As String
Public BillingAdmin As String
Public FollowUp As String
Public Discussed As String
Public EMSelector As String
Public EMCalculator As String
Public SelectedTab As Long
Public AppointmentClass As String
Public EMAdditional_1 As String
Public EMAdditional_2 As String
Public EMAdditional_3 As String
Public DentalStatus As String
'Public PowerBall As Integer
Public TimeFactor As String 'SCR-43819
'SF 65426 2-18-05
Public Discharge As String

Private mobjEandMs As WAMData.EandMs
Private mobjAdminCodes As WAMData.AdminCodes

Private Sub Class_Initialize()
    Dim objWAM As WAM
    Dim objEMs As EandMs
    
    Set objWAM = New WAM
    Set mobjEandMs = objWAM.ClassFactory(ctEandMs)
    Set mobjAdminCodes = objWAM.ClassFactory(ctAdminCodes)

End Sub

Private Function IPersist_ChildFactory() As Object

End Function

Private Function IPersist_Persist() As String
    IPersist_Persist = Persist(Me)
End Function

Private Function IPersist_Restore(ByVal Persist As String) As Boolean
    IPersist_Restore = Restore(Me, Persist)
End Function

Public Property Get EandMs() As WAMData.EandMs
    Set EandMs = mobjEandMs
End Property

Public Property Set EandMs(ByVal vNewValue As WAMData.EandMs)
    Set mobjEandMs = vNewValue
End Property

Public Property Get AdminCodes() As WAMData.AdminCodes
    Set AdminCodes = mobjAdminCodes
End Property

Public Property Set AdminCodes(ByVal vNewValue As WAMData.AdminCodes)
    Set mobjAdminCodes = vNewValue
End Property
'Called by ClearDispositionData SCR-35155
Public Sub Clear()
    Disposition = vbNullString
    InjuryIllness = vbNullString
    BillingAdmin = vbNullString
    FollowUp = vbNullString
    Discussed = vbNullString
    EMSelector = vbNullString
    EMCalculator = vbNullString
    SelectedTab = 0
    AppointmentClass = vbNullString
    EMAdditional_1 = vbNullString
    EMAdditional_2 = vbNullString
    EMAdditional_3 = vbNullString
    DentalStatus = vbNullString
    TimeFactor = vbNullString
    'PowerBall = 0
    mobjAdminCodes.Clear
    mobjEandMs.Clear
End Sub

