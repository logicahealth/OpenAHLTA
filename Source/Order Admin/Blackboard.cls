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
Attribute VB_Name = "Blackboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mComm As MMMHISComm.Comm
Private mConfig As MMMHISConfiguration.Config
Private mLogon As MMMHISLogon.Logon
Private mUser As Object
Private moShared As CWShared
Private mDAS As ICHCSII_DAS
Private mMedOrders As colMedOrders
Private mLabOrders As colLabOrders
Private mRadOrders As colRadOrders
Private moDatabase As CHCSII_conn.Conn

Private Sub Class_Terminate()
 
    Set mComm = Nothing
    Set mConfig = Nothing
    Set mLogon = Nothing
    Set mUser = Nothing
    Set moShared = Nothing
    Set mDAS = Nothing
    Set mMedOrders = Nothing
    Set mLabOrders = Nothing
    Set mRadOrders = Nothing
    
End Sub

Public Property Get Comm() As MMMHISComm.Comm
  Set Comm = mComm
End Property

Public Property Set Comm(ByRef NewValue As Object)
  Set mComm = NewValue
End Property

Public Property Get Config() As MMMHISConfiguration.Config
  Set Config = mConfig
End Property

Public Property Set Config(ByRef NewValue As Object)
  Set mConfig = NewValue
End Property

Public Property Get Logon() As MMMHISLogon.Logon
  Set Logon = mLogon
End Property

Public Property Set Logon(ByRef NewValue As Object)
  Set mLogon = NewValue
End Property

Public Property Get User() As Object
  Set User = mUser
End Property

Public Property Set User(ByRef NewValue As Object)
  Set mUser = NewValue
End Property

Public Property Get SharedObj() As CWShared
    If moShared Is Nothing Then Set moShared = New CWShared
    Set SharedObj = moShared
End Property

Public Sub ClearBlackBoard()
    Set moShared = Nothing
End Sub

Public Property Get I_DAS() As ICHCSII_DAS
    Dim Connection As CHCSII_conn.Conn
    On Error GoTo ErrHandler
  
    If mDAS Is Nothing Then
        Set mDAS = DataConn.CHCSII_DAS
    End If
    Set I_DAS = mDAS
    
    Exit Property
  
ErrHandler:
    moShared.ShowVBError Err.Number, Err.Description, "BlackBoard.", "OrderAdmin", vbExclamation

End Property

Public Property Get DataConn() As CHCSII_conn.Conn
    If moDatabase Is Nothing Then
        Set moDatabase = New CHCSII_conn.Conn
    End If
    Set DataConn = moDatabase
End Property

Public Property Get MedOrders() As colMedOrders
    If mMedOrders Is Nothing Then Set mMedOrders = New colMedOrders
    Set MedOrders = mMedOrders
End Property


Public Property Get labOrders() As colLabOrders
    If mLabOrders Is Nothing Then Set mLabOrders = New colLabOrders
    Set labOrders = mLabOrders
End Property


Public Property Get radOrders() As colRadOrders
    If mRadOrders Is Nothing Then Set mRadOrders = New colRadOrders
    Set radOrders = mRadOrders
End Property


