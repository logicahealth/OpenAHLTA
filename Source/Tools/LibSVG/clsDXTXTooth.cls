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
Attribute VB_Name = "clsDXTXTooth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'This is used to store the SVG string for the Dxs and Txs per tooth per surface

'Public ToothID As Integer
'Public SurfaceID As Integer
'Public alias As Integer
'Public DxTxSVG As String

Private mvarAliasID As Integer 'local copy
Private mvarToothID As Integer 'local copy
Private mvarSurfaceID As Integer 'local copy
Private mvarDxTxSVG As String 'local copy
Public Property Let DxTxSVG(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DxTxSVG = 5
    mvarDxTxSVG = vData
End Property


Public Property Get DxTxSVG() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DxTxSVG
    DxTxSVG = mvarDxTxSVG
End Property



Public Property Let SurfaceID(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SurfaceID = 5
    mvarSurfaceID = vData
End Property


Public Property Get SurfaceID() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SurfaceID
    SurfaceID = mvarSurfaceID
End Property



Public Property Let ToothID(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ToothID = 5
    mvarToothID = vData
End Property


Public Property Get ToothID() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ToothID
    ToothID = mvarToothID
End Property



Public Property Let AliasID(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AliasID = 5
    mvarAliasID = vData
End Property


Public Property Get AliasID() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AliasID
    AliasID = mvarAliasID
End Property



