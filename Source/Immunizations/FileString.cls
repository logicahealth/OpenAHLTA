VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFileString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'FileString Class - encapsulates common file/path string operations
'   Note: the project must reference "Microsoft Scripting Runtime"
'       which is file scrrun.dll in the Windows System directory.
'©1998 SunOpTech, Ltd.

'class member variables
Private m_sPathAndFileName As String
Private m_fso As New FileSystemObject       'for local functions
Public Function CopyFile(sSource As String, sTarget As String, Optional bOverwrite As Boolean = True) As Boolean
  On Error Resume Next
  m_fso.CopyFile sSource, sTarget, bOverwrite
End Function


Public Property Get PathAndFileName() As String
'returns current PathAndFileName

    PathAndFileName = m_sPathAndFileName
    
End Property

Public Property Let PathAndFileName(ByVal NewPathAndFileName As String)
'sets a PathAndFileName and all the object's properties

    m_sPathAndFileName = Trim$(NewPathAndFileName)
    
End Property

Public Property Get FileExists() As Boolean
'returns True if PathAndFileName exists - ReadOnly

    FileExists = m_fso.FileExists(m_sPathAndFileName)
    
End Property

Public Property Get FolderExists() As Boolean
'returns True if the path specified by PathAndFileName exists - ReadOnly

    FolderExists = m_fso.FolderExists(StripFileName())
    
End Property

Public Property Get ParsePath() As String()
'splits the PathAndFileName into an array of strings - ReadOnly

    ParsePath = Split(m_sPathAndFileName, "\")
    
End Property

Public Function MakePath() As Boolean
'checks the path in PathAndFileName is valid, creates if not

    MakePath = MakeFolder(StripFileName())
    
End Function

Private Function MakeFolder(ByVal DirectorySpec As String) As Boolean
'creates a directory structure, assumes only a path is passed

    Dim sDirSpec As String
    
    On Error GoTo MakeFolder_Error
    
    sDirSpec = Trim$(DirectorySpec)
    'if we're out of chars, then we didn't succeed. Works for local drives & UNCs,
    '   but will fail on a UNC on which no priviledge is granted (as it should).
    If Len(sDirSpec) Then
        MakeFolder = True
        'if this dir exists, exit. Otherwise, check next dir & create this one.
        If m_fso.FolderExists(sDirSpec) Then Exit Function
        If Not MakeFolder(StripFileName(sDirSpec)) Then
            MakeFolder = False
            Exit Function
        End If
        m_fso.CreateFolder (sDirSpec)
    End If
    
MakeFolder_Exit:
    Exit Function
    
MakeFolder_Error:
    MakeFolder = False
    Resume MakeFolder_Exit
    
End Function

Public Property Get FileType() As String
'returns the File Type

    Dim filTemp As File
    
    If FileExists Then
        Set filTemp = m_fso.GetFile(m_sPathAndFileName)
        FileType = filTemp.Type
    End If

End Property

Public Function StripFileName(Optional ByVal sFile As String) As String
'returns path
    Dim i As Integer
    On Error Resume Next
    If Len(sFile) = 0 Then sFile = m_sPathAndFileName
    i = InStrRev(sFile, "\")
    StripFileName = Left$(sFile, i - 1)
End Function

Public Function StripPathName(Optional ByVal sFullPath As String = "") As String
'returns filename
    Dim i As Integer
    On Error Resume Next
    If Len(sFullPath) = 0 Then sFullPath = m_sPathAndFileName
    i = InStrRev(sFullPath, "\")
    StripPathName = Trim$(Mid$(sFullPath, i + 1))
End Function

Public Property Get FileName() As String
'returns only the FileName portion of the FileAndPathName - ReadOnly
'Note: this function assumes that all filenames
'   will have at least 1 dot in them. Everything
'   before the last dot is considered part of the
'   filename.

    Dim sFile As String
    Dim sArray() As String
    Dim i As Integer
    
    sFile = StripPathName()
    If InStr(sFile, ".") = 0 Then Exit Property     'if no dot
    sArray = Split(sFile, ".")
    i = UBound(sArray)
    If i > 0 Then           'there was a filename
        If i = 1 Then       'only one dot
            FileName = sArray(0)
        Else                'more than one dot
            'remove extension from end
            i = InStrRev(sFile, ".")
            FileName = Mid$(sFile, 1, i - 1)
        End If
    End If
    
End Property

Public Property Get FileExt() As String
'returns only the File Extension portion of the FileAndPathName - ReadOnly

    Dim i As Integer
    i = InStrRev(m_sPathAndFileName, ".")
    FileExt = Mid$(m_sPathAndFileName, i + 1)
    
End Property

Public Function CompressPath() As String
'returns a compressed path and filename

    Dim sArray() As String
    Dim sNewArray(3) As String
    Dim i As Integer
    
    sArray = Split(m_sPathAndFileName, "\")
    i = UBound(sArray)
    If i > 2 Then
        sNewArray(0) = sArray(0)
        sNewArray(1) = sArray(1)
        sNewArray(2) = "..."
        sNewArray(3) = sArray(i)
        CompressPath = Join(sNewArray, "\")
    Else
        CompressPath = m_sPathAndFileName
    End If
    
End Function

Public Function DriveFreeSpace() As Long
  Dim oDrive As Object
  On Error Resume Next
  Set oDrive = m_fso.GetDrive(m_fso.GetDriveName(m_sPathAndFileName))
  DriveFreeSpace = oDrive.FreeSpace / 1024
End Function

Private Sub Class_Initialize()
  On Error Resume Next
  m_sPathAndFileName = "C:\"
End Sub



Public Function FileSize() As Long
  Dim oFile
  On Error Resume Next
  Set oFile = m_fso.GetFile(m_sPathAndFileName)
  FileSize = oFile.Size
End Function

Private Sub Class_Terminate()
  On Error Resume Next
  Set m_fso = Nothing
End Sub

Public Function DeleteFile(Optional ByVal sPattern As String = "", Optional bForce As Boolean = False)
  On Error Resume Next
  If Len(sPattern) = 0 Then sPattern = m_sPathAndFileName
  m_fso.DeleteFile sPattern, bForce
End Function
