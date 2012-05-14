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
Attribute VB_Name = "clsDataElemOld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarIsDirty As Boolean
Private mvarDx As CHCSII_Dental_DXTX.clsDx
Private mvarTx As CHCSII_Dental_DXTX.clsTx

Public Property Set Tx(vData As CHCSII_Dental_DXTX.clsTx)

    Set mvarTx = vData
    
End Property
Public Property Get Tx() As CHCSII_Dental_DXTX.clsTx
    
    Set Tx = mvarTx
    
End Property
Public Property Set Dx(vData As CHCSII_Dental_DXTX.clsDx)
    
    Set mvarDx = vData
    
    If Not mvarDx Is Nothing And Not mvarTx Is Nothing Then
        MsgBox "the dx and tx are set."
    End If
    
End Property
Public Property Get Dx() As CHCSII_Dental_DXTX.clsDx
    
    Set Dx = mvarDx
    
End Property


Public Property Let IsDirty(vData As Boolean)
    mvarIsDirty = vData
End Property

Public Property Get IsDirty() As Boolean
    IsDirty = mvarIsDirty
End Property

