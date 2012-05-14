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
Attribute VB_Name = "CTooth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mintToothNumber As Integer
Private mblnChecked As Boolean
Private mblnMissing As Boolean
Private mstrPriName As String
Private mstrSecName As String
Private mstrName As String
Private mintToothPos As Integer
Private mlngPriId As Long
Private mblnEnabled As Boolean
Private mblnImplant As Boolean
Private mblnPontic As Boolean
Private mblnStateChanged As Boolean
Private mstrNewName As String
Private mblnTemporary As Boolean

Public Property Get Temporary() As Boolean
    Temporary = mblnTemporary
End Property

Public Property Let Temporary(ByVal blnTemp As Boolean)
    mblnTemporary = blnTemp
End Property

Public Property Get OralStateChanged() As Boolean
    OralStateChanged = mblnStateChanged
End Property

Public Sub ChangeOralState(ByVal blnChanged As Boolean)
    mblnStateChanged = blnChanged
End Sub

Public Property Get Pontic() As Boolean
    Pontic = mblnPontic
End Property

Public Property Let Pontic(ByVal blnValue As Boolean)
    mblnPontic = blnValue
End Property
Public Property Get Implant() As Boolean
    Implant = mblnImplant
End Property
Public Property Let Implant(ByVal blnValue As Boolean)
    mblnImplant = blnValue
End Property

Public Property Get ToothNumber() As Integer
    ToothNumber = mintToothNumber
End Property
Public Property Get ToothPosition() As Integer
    ToothPosition = mintToothPos
End Property
Public Property Get Checked() As Boolean
    Checked = mblnChecked
End Property
Public Property Get Missing() As Boolean
    Missing = mblnMissing
End Property
Public Property Get ToothName() As String
    ToothName = mstrName
End Property
Friend Property Let ToothPosition(ByVal intPos As Integer)
    mintToothPos = intPos
End Property
Friend Property Let ToothName(ByVal strName As String)
    mstrName = strName
End Property
Friend Property Let ToothNumber(ByVal intNum As Integer)
    mintToothNumber = intNum
End Property

Friend Property Let Checked(ByVal blnChecked As Boolean)
    mblnChecked = blnChecked
End Property

Friend Property Let Missing(ByVal blnMissing As Boolean)
    mblnMissing = blnMissing
End Property

Public Property Let Enabled(ByVal blnValue As Boolean)
    mblnEnabled = blnValue
End Property

Public Property Get Enabled() As Boolean
    Enabled = mblnEnabled
End Property

Private Sub Class_Initialize()
    mblnStateChanged = False
End Sub
