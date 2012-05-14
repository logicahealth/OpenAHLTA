VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iBackend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public SkipFacilityWorkAround As Boolean

Public Function CreateRecordsetFromFile(ByVal sFileName As String, ByVal sDelimiter As String) As ADODB.Recordset
End Function

Public Sub ExecuteSQL(ByVal SQL As String)
End Sub

Public Function OpenRecordset(ByVal SQL As String) As ADODB.Recordset
End Function

