VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Spool_Manager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mobj_Files As Collection
Private mobjFSO As Object 'New FileSystemObject
Private mobjTS As Object 'TextStream

Public Sub Spool_On(Filename As String)
On Error GoTo ErrHandler


    Dim TempString As String
    Dim TheFile As String
    Dim TheFolder As String
    Dim i As Long
    
'    TempString = StrReverse(Filename)
'    TheFolder = Mid(TempString, InStr(1, TempString, "\") + 1)
'    TheFolder = StrReverse(TheFolder)
'    TheFile = Mid(TempString, 1, InStr(1, TempString, "\") - 1)
'    TheFile = StrReverse(TheFile)
'
'    Set mobjFSO = New FileSystemObject
'
'    'Set mobjTS = mobjFSO.CreateTextFile("c:\ABCDEF~1\trash.txt", True)

    
    Set mobjTS = mobjFSO.CreateTextFile(Filename, True)

    
    'i = 1 / 0
    '"C:\DOCUME~1\BWEBB~1.CHC\LOCALS~1\Temp\run_patch.sql"

GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "Spool_Manager", "Spool_On", Err.Number, Err.Description
CleanUp:
End Sub

Public Sub Spool_Off()
On Error GoTo ErrHandler

    mobjTS.Close
    
GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "Spool_Manager", "Spool_Off", Err.Number, Err.Description
CleanUp:
End Sub

Public Function Put_Line(Line As String) As Boolean
On Error GoTo ErrHandler

    mobjTS.WriteLine (Line)
    Put_Line = True
    
GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "Spool_Manager", "Put_Line", Err.Number, Err.Description
CleanUp:
End Function

Private Sub Class_Initialize()
    Set mobjFSO = CreateObject("Scripting.FileSystemObject")
End Sub

Private Sub Class_Terminate()
    Set mobjFSO = Nothing
    Set mobjTS = Nothing
End Sub
