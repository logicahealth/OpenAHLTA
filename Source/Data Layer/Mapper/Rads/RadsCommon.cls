VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RadsCommon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function GetMonth(ByRef sMonth As String) As String
On Error GoTo ErrorHandler

    Select Case sMonth
        Case 12
            GetMonth = "Dec"
        Case 11
            GetMonth = "Nov"
        Case 10
            GetMonth = "Oct"
        Case 9
        GetMonth = "Sep"
        Case 8
        GetMonth = "Aug"
        Case 7
        GetMonth = "Jul"
        Case 6
        GetMonth = "Jun"
        Case 5
        GetMonth = "May"
        Case 4
        GetMonth = "Apr"
        Case 3
        GetMonth = "Mar"
        Case 2
        GetMonth = "Feb"
        Case 1
        GetMonth = "Jan"
        
        
    End Select
    
    Exit Function
ErrorHandler:

  Err.Raise Err.Number, "DM_Radiology.MapperCDR.GetMonth", Err.Description, Err.HelpFile, Err.HelpContext

End Function

