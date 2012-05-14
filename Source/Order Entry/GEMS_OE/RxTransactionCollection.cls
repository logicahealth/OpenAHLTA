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
Attribute VB_Name = "RxTransactionCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Public Type to be kept in the collection
Public Type RxTransaction
     RXTablet As RxType
     RXWarnings As Collection
     Action As hlACTION
     NewOrderid As String
     OldOrderid As String
End Type


'local variable to hold collection
Private mcolTransactions As Collection

Public Property Get Count() As Long
     Count = mcolTransactions.Count
End Property

Public Property Get Item(Index As Integer) As RxTransaction
     On Error GoTo ErrorHandler:
     'Make sure item requested is in the collection
     If Index > 0 And Index <= mcolTransactions.Count Then
          Item = mcolTransactions.Item(Index)
     End If     'if item not in collection then return Nothing
     Exit Property
     
ErrorHandler:
     'If error, return Nothing
End Property

Friend Function Add(tRxtransaction As RxTransaction)
     mcolTransactions.Add (tRxtransaction)
End Function

'Enable For...Each support
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
     Set NewEnum = mcolTransactions.[_NewEnum]
End Property

Private Sub Class_Initialize()
     Set mcolTransactions = New Collection
     Debug.Print "RxTransactionCollection Initialized"
End Sub

Public Function SearchOrderId(szOrderId As String) As Collection
     Dim tRxtransaction As Variant 'RxTransaction
     Set SearchOrderId = New Collection
     
     For Each tRxtransaction In mcolTransactions
          With tRxtransaction
               If .NewOrderid = szOrderId Or .OldOrderid = szOrderId Then
                    SearchOrderId.Add tRxtransaction
               End If
          End With
     Next
     'If szorderid not found, return empty collection
End Function


Private Sub Class_Terminate()
     Set mcolTransactions = Nothing
     Debug.Print "RxTransactionCollection Terminated"
End Sub





