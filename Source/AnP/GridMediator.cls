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
Attribute VB_Name = "GridMediator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Created to manage Grid focus and give the user a visual indication of which
'grid has focus.  SCR 28769  Phil Crowder 11/11/02

Option Explicit

Public Enum enumGridFocus
  gfNone
  gfDiagnoses
  gfOrders
End Enum

Private WithEvents mDiagGrid As vsFlexGrid
Attribute mDiagGrid.VB_VarHelpID = -1
Private WithEvents mOrdersGrid As vsFlexGrid
Attribute mOrdersGrid.VB_VarHelpID = -1

Private mFocus As enumGridFocus

Private Sub mDiagGrid_GotFocus()
  mOrdersGrid.BackColorSel = &H80000010
  mDiagGrid.BackColorSel = &H8000000D
  mFocus = gfDiagnoses
End Sub

Private Sub mOrdersGrid_GotFocus()
  mDiagGrid.BackColorSel = &H80000010
  mOrdersGrid.BackColorSel = &H8000000D
  mFocus = gfOrders
End Sub

Public Property Get GridWithFocus() As enumGridFocus
  GridWithFocus = mFocus
End Property

Public Sub Init(ByRef DiagGrid As vsFlexGrid, ByRef OrdersGrid As vsFlexGrid)
  Set mDiagGrid = DiagGrid
  Set mOrdersGrid = OrdersGrid
End Sub
