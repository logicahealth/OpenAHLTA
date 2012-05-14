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
Attribute VB_Name = "cGUISupport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private oListener As cGUIListener
Attribute oListener.VB_VarHelpID = -1

Private Sub Class_Initialize()
    Set oListener = New cGUIListener
End Sub

Public Sub SetCmplxTxsRef(oCmplxTxs As cCmplxTransactions)
    oListener.SetCmplxTxsRef oCmplxTxs
End Sub

Private Sub Class_Terminate()
    Set oListener = Nothing
End Sub
