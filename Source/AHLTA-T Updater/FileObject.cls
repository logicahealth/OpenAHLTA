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
Attribute VB_Name = "FileObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private m_fso As FileSystemObject

Public Sub CreateFolder(FolderName As String)
    
Dim aFolderArray() As String
Dim sPath As String
Dim iCounter As Integer

    aFolderArray = Split(FolderName, "\")
        
    sPath = aFolderArray(0)
    For iCounter = 1 To UBound(aFolderArray)
        sPath = sPath & "\" & aFolderArray(iCounter)
    
        If Not m_fso.FolderExists(sPath) Then
            m_fso.CreateFolder sPath
        End If
    
    Next

End Sub
 
Public Sub CreateTextFile(FileName As String)

    If Not m_fso.FileExists(FileName) Then
        m_fso.CreateTextFile FileName
    End If
    
End Sub
 
Public Sub DeleteFolder(FolderName As String)
        
    If m_fso.FolderExists(FolderName) Then
        m_fso.DeleteFolder FolderName
    End If
    
End Sub

Public Sub DeleteFile(FileName As String)
    
    On Error Resume Next
    
    If m_fso.FileExists(FileName) Then
        m_fso.DeleteFile FileName
    End If
    
End Sub

Public Sub DeleteAllFiles(FolderName As String)
    
    Dim oFile As File
        
    On Error Resume Next
    
    For Each oFile In m_fso.GetFolder(FolderName).Files
        m_fso.DeleteFile oFile.Path, True
    Next oFile

End Sub

 
Public Sub DeleteAllFolders(FolderName As String)
    
    Dim oFolder As Folder
        
    On Error Resume Next
    
    For Each oFolder In m_fso.GetFolder(FolderName).SubFolders
        m_fso.DeleteFolder oFolder.Path, True
    Next

End Sub

Public Sub CopyFolder(SourceFolder As String, DestinationFolder As String)
 
    If m_fso.FolderExists(DestinationFolder) Then m_fso.DeleteFolder DestinationFolder, True
    m_fso.CopyFolder SourceFolder, DestinationFolder

End Sub

Public Sub CopyFile(SourceFile As String, DestinationFile As String)
    
    If m_fso.FileExists(DestinationFile) Then m_fso.DeleteFile DestinationFile, True
    m_fso.CopyFile SourceFile, DestinationFile

End Sub

Public Sub MoveFolder(SourceFolder As String, DestinationFolder As String)
  
    If m_fso.FolderExists(DestinationFolder) Then m_fso.DeleteFolder DestinationFolder, True
    m_fso.CopyFile SourceFolder, DestinationFolder

End Sub

Public Sub MoveFile(SourceFile As String, DestinationFile As String)
    
    If m_fso.FileExists(DestinationFile) Then m_fso.DeleteFile DestinationFile, True
    m_fso.MoveFile SourceFile, DestinationFile

End Sub

 
Public Function DriveExists(sDrive As String) As Boolean

    DriveExists = m_fso.DriveExists(sDrive)

End Function

Public Function GetFolder(FolderName As String) As Folder

    Set GetFolder = m_fso.GetFolder(FolderName)
    
End Function
 
Public Function GetParentFolderName(FolderName As String) As Folder

    Set GetParentFolderName = m_fso.GetSpecialFolder(TemporaryFolder)

End Function
 
 
 Public Function GetFile(FileName As String) As File
    Set GetFile = m_fso.GetFile(FileName)
 End Function

 
 Public Function FolderExists(FolderName As String) As Boolean
    FolderExists = m_fso.FolderExists(FolderName)
 End Function

 Public Function FileExists(FileName As String) As Boolean
    FileExists = m_fso.FileExists(FileName)
 End Function
 

Private Sub Class_Initialize()
    Set m_fso = New FileSystemObject
End Sub

Private Sub Class_Terminate()
    Set m_fso = Nothing
End Sub
