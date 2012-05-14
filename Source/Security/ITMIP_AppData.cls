VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ITMIP_AppData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Function GetAppData(ByVal AppName As String, _
                           ByRef UID As String, _
                           ByRef PWD As String, _
                           ByRef SID As String) As Boolean
'- returns the application userid and password and SID for tnsnames entry point in the provided strings.
'  returns True on success; false otherwise.
'  AppName for us is "CHCSII"
End Function
