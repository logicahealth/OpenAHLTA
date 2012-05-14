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
Attribute VB_Name = "SnoMed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Event BuildList(ByVal SnoDescId As Long, ByVal Description As String, ByVal HasChildren As Boolean)

Private mConn As CHCSII_CONN.Conn
Private mDAS As ICHCSII_DataAccess.ICHCSII_DAS



Private Sub Class_Initialize()

    On Error GoTo ErrHnd
    
    Set mConn = New CHCSII_CONN.Conn
    Set mDAS = mConn.CHCSII_DAS(Auto)
    
    Exit Sub
ErrHnd:
    MsgBox Err.Description, vbInformation, "SnoMed.Initialize"
    
End Sub


Private Sub Class_Terminate()

    Set mConn = Nothing
    Set mDAS = Nothing
    
End Sub


