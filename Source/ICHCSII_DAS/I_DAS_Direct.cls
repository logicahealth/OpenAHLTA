VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "I_DAS_Direct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function ExecuteSQL(ByVal SQL As String, _
                           Optional ByVal UserID As String, _
                           Optional ByVal RecordDateTime As Date) As Variant: End Function
Public Function BeginTrans() As Boolean:             End Function
Public Function CommitTrans() As Boolean:            End Function
Public Function RollbackTrans() As Boolean:          End Function
Public Property Get ActiveConnection() As ADODB.Connection: End Property


