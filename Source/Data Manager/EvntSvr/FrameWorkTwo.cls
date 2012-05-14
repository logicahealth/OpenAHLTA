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
Attribute VB_Name = "FrameWorkTwo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub CheckEFFolders()

'This is called from SubMain

  Dim sRootDirectory As String
  Dim sEICRegPath As String
  
      sRootDirectory = GetRootDirectory
      Call CreateDataManagerRegistry

      Call SaveRegistryPath("BMIST_195b_IN", sRootDirectory & "\TMIPB2\EF\Components\C2T\BMIST\In")
      Call SaveRegistryPath("BMIST_196b_IN", sRootDirectory & "\TMIPB2\EF\Components\C2T\BMIST\In")
      Call SaveRegistryPath("BMIST_IN", sRootDirectory & "\TMIPB2\EF\Components\C2T\BMIST\In")
      Call SaveRegistryPath("PIC_IN", sRootDirectory & "\TMIPB2\EF\Components\C2T\PIC\In")
      Call SaveRegistryPath("DISK_IN", sRootDirectory & "\TMIPB2\EF\Components\C2T\DISK\In")
      Call SaveRegistryPath("C2T_IN", sRootDirectory & "\TMIPB2\EF\Components\C2T\PVF\In")

End Sub



'This is for the Framework2 timer
Private Function GetRootDirectory() As String

  Dim sRegPath As String
  
'SF SCR 71407 6/6/05
On Error GoTo ErrHnd

  If gobjShared Is Nothing Then
    Set gobjShared = New CWShared
  End If

  sRegPath = gobjShared.RTDataPath
  sRegPath = Trim$(sRegPath)
  
  If Len(sRegPath) >= 2 Then
    GetRootDirectory = Left(sRegPath, 2)
  Else
    GetRootDirectory = "C:"
  End If
    
Exit Function
ErrHnd:
    'SF SCR 71407 6/6/05
    If Err.Number = -2147417848 Then 'The object invoked has disconnected from its clients.
        'For now, end.
        End
        'Exit Function
    End If
    
End Function

'This will Create the DataManager Registry Key if it does not exist.
Private Function CreateDataManagerRegistry() As Boolean

On Error GoTo ErrHandler

    Dim sTemp() As String

    Dim objReg      As RegObj.Registry
    Dim ParentKey   As RegObj.RegKey
    Dim strKeyName  As String
    Dim ChildKey    As RegObj.RegKey
    Dim lLong As Long
    
    strKeyName = gobjShared.RegistryKey
    Set objReg = New RegObj.Registry
    Set ParentKey = objReg.RegKeyFromString(strKeyName)

    CreateDataManagerRegistry = False

    For Each ChildKey In ParentKey.SubKeys
      If UCase$(Trim$(ChildKey.Name)) = "DATAMANAGER" Then
        CreateDataManagerRegistry = True
        Exit Function
      End If
    Next ChildKey
    
    Call ParentKey.SubKeys.Add("DataManager")
    CreateDataManagerRegistry = True

Exit Function
ErrHandler:

  'MsgBox "Could not create the Registry Key: DATAMANAGER" & vbCrLf & Err.Description, vbOKOnly, "Data Manager"
  CreateDataManagerRegistry = False

End Function


'We are assuming that "\HKEY_LOCAL_MACHINE\SOFTWARE" already exists=
'This Procedure will Create the Registry Setting:
'     "\HKEY_LOCAL_MACHINE\SOFTWARE\TMIP\EIC\EICPath = sEICPath

Private Sub SaveRegistryEICPath(ByVal sEICPath As String)

On Error GoTo ErrHandler

Dim objReg As RegObj.Registry
Dim ParentKey As RegObj.RegKey
Dim ChildKey As RegObj.RegKey
Dim objValue As RegObj.RegValue

Dim bKeyFound As Boolean

Dim sRegPath As String
Dim sChildKey As String

Dim sErrMsg As String

sErrMsg = vbNullString

  Set objReg = New RegObj.Registry
  
  '*****************************************************
  '******Look for TMIP**********************************
  '*****************************************************
  sChildKey = "TMIP"
  bKeyFound = False
  sRegPath = "\HKEY_LOCAL_MACHINE\SOFTWARE\"
  Set ParentKey = objReg.RegKeyFromString(sRegPath)
  For Each ChildKey In ParentKey.SubKeys
    If UCase$(Trim$(ChildKey.Name)) = sChildKey Then
      bKeyFound = True
      Exit For
    End If
  Next ChildKey
  
  If bKeyFound = False Then
    sErrMsg = "Could not add Registry Key 'TMIP'."
    Call ParentKey.SubKeys.Add("TMIP")
    
  End If
  
  '*****************************************************
  '******Look for EIC***********************************
  '*****************************************************
  sChildKey = "EIC"
  bKeyFound = False
  sRegPath = "\HKEY_LOCAL_MACHINE\SOFTWARE\TMIP"
  Set ParentKey = objReg.RegKeyFromString(sRegPath)
  For Each ChildKey In ParentKey.SubKeys
    If UCase$(Trim$(ChildKey.Name)) = sChildKey Then
      bKeyFound = True
      Exit For
    End If
  Next ChildKey
  
  If bKeyFound = False Then
    sErrMsg = "Could not add Registry Key 'EIC'."
    Call ParentKey.SubKeys.Add("EIC")
  End If
  
  '*****************************************************
  '******Look for EICPath*******************************
  '*****************************************************
  sChildKey = "EICPath"
  bKeyFound = False
  sRegPath = "\HKEY_LOCAL_MACHINE\SOFTWARE\TMIP\EIC"
  Set ParentKey = objReg.RegKeyFromString(sRegPath)
  For Each objValue In ParentKey.Values
    If UCase$(Trim$(objValue.Name)) = "EICPATH" Then
      bKeyFound = True
      Exit For
    End If
  Next objValue
  
  If bKeyFound = False Then
    sErrMsg = "Could not add Registry Key 'EICPath'."
    Call ParentKey.Values.Add("EICPath", sEICPath)
  Else
    If UCase$(Trim$(objValue.Value)) <> UCase$(Trim$(sEICPath)) Then
      sErrMsg = "Could not set Registry Key 'EICPath' to " & sEICPath & "."
      objValue.Value = sEICPath
    End If
  End If
  
  
  GoTo Cleanup

ErrHandler:

  'MsgBox sErrMsg & vbCrLf & Err.Description, vbInformation, "Data Manager"

Cleanup:

  Set objReg = Nothing
  Set ParentKey = Nothing
  Set ChildKey = Nothing
  Set objValue = Nothing

End Sub

Private Function GetRegValue(ByVal sRegValueName As String) As String

On Error GoTo ErrHandler

  Dim objReg As RegObj.Registry
  Dim objKey As RegObj.RegKey
  Dim objValue As RegObj.RegValue
  Dim sRegKey As String
  
  GetRegValue = vbNullString
  
  sRegKey = gobjShared.RegistryKey & "\DataManager"
  
  Set objReg = New RegObj.Registry
  Set objKey = objReg.RegKeyFromString(sRegKey)
  
  For Each objValue In objKey.Values
    If UCase$(Trim$(objValue.Name)) = UCase$(Trim$(sRegValueName)) Then
      GetRegValue = objValue.Value
      Exit For
    End If
  Next objValue
  
  GoTo Cleanup
  
ErrHandler:

    'Call gobjShared.ShowVBError(Err.Number, Err.Description, "DMBE_Ops.GetRegValue", "Data Manager", vbExclamation)
  
Cleanup:
  
  Set objReg = Nothing
  Set objKey = Nothing
  Set objValue = Nothing

End Function


'If Registry Entry and Folder does not exist, then Build it.
Private Sub SaveRegistryPath(ByVal RegKeyName As String, ByVal sDefaultFolder As String)
    
    Dim sTemp() As String
    Dim objReg As RegObj.Registry
    Dim objKey   As RegObj.RegKey
    Dim objValue As RegObj.RegValue
    Dim sExistingFolder As String
    
    Dim sMsg As String

    Dim nRetValue As Integer
    
    'call ParentKey.SubKeys(2).Values.Add("Hello", "0")
    
    On Error GoTo ErrHandler
    
    'sFolder = vbNullString
    sMsg = vbNullString
    sExistingFolder = vbNullString
    
    If RegistryValueExists(RegKeyName, sExistingFolder, sDefaultFolder) = False Then
        
      'This will Build the Directories
      nRetValue = CheckDirs(RegKeyName, sExistingFolder, sDefaultFolder)
      
      Set objReg = New RegObj.Registry
      Set objKey = objReg.RegKeyFromString(gobjShared.RegistryKey & "\DataManager")
      
      sMsg = "Cannot create Registry Value " & sDefaultFolder & " for registry key " & RegKeyName & "."
      Call objKey.Values.Add(UCase$(RegKeyName), sDefaultFolder)
    
    Else
    
      If sExistingFolder = vbNullString Then
      
        'Modify the Existing Registry Key with DefaultFolder Path
        Set objReg = New RegObj.Registry
        Set objKey = objReg.RegKeyFromString(gobjShared.RegistryKey & "\DataManager")
        
        If Not objKey Is Nothing Then
          For Each objValue In objKey.Values
          
            If UCase$(Trim$(objValue.Name)) = UCase$(Trim$(RegKeyName)) Then
              
              sMsg = "Cannot create Registry Value " & sDefaultFolder & " for registry key " & RegKeyName & "."
              objValue.Value = sDefaultFolder
              sExistingFolder = sDefaultFolder
            
            End If
        
          Next objValue
        End If
        
      
      End If
      
    
    End If
      
    nRetValue = CheckDirs(RegKeyName, sExistingFolder, sDefaultFolder)
    

GoTo Cleanup

ErrHandler:

  'MsgBox sMsg & vbCrLf & Err.Description, vbInformation, "Data Manager"

Cleanup:

    Set objReg = Nothing
    Set objKey = Nothing
    Set objValue = Nothing
    
End Sub

'If Registry Value is there, don't change it.
Private Function RegistryValueExists(ByVal RegKeyName As String, ByRef sExistingFolder As String, ByVal sDefaultFolder As String) As Boolean

On Error GoTo ErrHandler
    
    Dim reg As Registry
    Dim sTemp() As String
    
    Dim objReg      As RegObj.Registry
    Dim objKey      As RegObj.RegKey
    Dim strKeyName  As String
    Dim objValue    As RegObj.RegValue
    
    Set objReg = New RegObj.Registry
    
    
    strKeyName = gobjShared.RegistryKey & "\DataManager"
    
    Set objKey = objReg.RegKeyFromString(strKeyName)
  
    RegistryValueExists = False
    sExistingFolder = vbNullString
    
    For Each objValue In objKey.Values
      If UCase$(objValue.Name) = UCase$(Trim(RegKeyName)) Then
        
        RegistryValueExists = True
        
        If objValue.Value <> vbNullString Then
          
          sExistingFolder = objValue.Value
          'Check Validity of Folder
          If BuildFolder(sExistingFolder) <> 0 Then
            sExistingFolder = vbNullString
          End If
          Exit Function
        
        End If
      
        Exit For
      End If
          
    Next objValue
    Set objValue = Nothing
    
GoTo Cleanup

ErrHandler:

    RegistryValueExists = False
    sExistingFolder = vbNullString
    
Cleanup:
    
    Set objReg = Nothing
    Set objKey = Nothing
    Set objValue = Nothing

End Function

Private Function CheckDirs(ByVal sRegKey As String, ByVal sExistingFolder As String, ByVal sDefaultFolder As String) As Long
   
   Dim fso As FileSystemObject
   Dim drv As Variant
   Dim dc As Variant
   Dim DriveStr As String
   Dim nRetValue As Integer
   Dim sPath As String
   
   On Error GoTo CheckOrBuildDirsError
   
   CheckDirs = CheckOrBuildDirs(sExistingFolder, sDefaultFolder)
   
   
      
  Exit Function
   
CheckOrBuildDirsError:
   
   'gobjShared.ShowVBError Err.Number, Err.Description, "CheckDirs", "DMFE", vbExclamation
   CheckDirs = -1
   Exit Function

End Function

Private Sub CreateDirs(FilePath As String)
   On Error Resume Next
   
   MkDir FilePath
End Sub

'Check if the Exisitng Folder is legitimate/
'  If yes, then Build it.
'  If no, then Build the Default Folder.

Private Function CheckOrBuildDirs(ByVal sExistingFolder As String, ByVal sDefaultFolder As String) As Integer
   
   On Error GoTo ErrHandler
   
   CheckOrBuildDirs = 0
      
      If sExistingFolder <> "" Then
        CheckOrBuildDirs = BuildFolder(sExistingFolder)
        If CheckOrBuildDirs <> 0 Then
          CheckOrBuildDirs = BuildFolder(sDefaultFolder)
        End If
      Else
        CheckOrBuildDirs = BuildFolder(sDefaultFolder)
      End If
      
        
Exit Function
   
ErrHandler:

  'gobjShared.ShowVBError Err.Number, Err.Description, "CheckOrBuildDirs", "DMFE", vbExclamation
  CheckOrBuildDirs = -1
  Exit Function

End Function

'This will Check or Build the Folder Path
Private Function BuildFolder(ByVal sFolder As String) As Long

On Error GoTo ErrHandler
   
   Dim Temp() As String
   Dim LB As Long
   Dim UB As Long
   Dim I As Long
   Dim sPath As String
   Dim fso         As FileSystemObject

   
      If sFolder <> vbNullString Then

         Temp = Split(sFolder, "\", -1, vbTextCompare)
         LB = LBound(Temp)
         UB = UBound(Temp)
         Set fso = New FileSystemObject
         If InStr(1, Temp(I), ":", vbTextCompare) <> 0 Then
            If fso.DriveExists(Temp(I) & "\") Then
               sPath = Temp(I)
            Else
               BuildFolder = -100 ' drive does not exist - bad return
               Set fso = Nothing
               Exit Function
            End If
         Else
            BuildFolder = -200 ' drive is not at start of path - bad return
            Set fso = Nothing
            Exit Function
         End If
         
         For I = LB + 1 To UB
            sPath = sPath & "\" & Temp(I)
            If fso.FolderExists(sPath) = False Then
               BuildFolder = -1
               Call fso.CreateFolder(sPath)
               BuildFolder = 0  'Good Return
            Else
                BuildFolder = 0
            End If
         Next I

      Else
         BuildFolder = -200 ' no path specified - bad return
         Set fso = Nothing
         Exit Function
      End If
      
      Set fso = Nothing

Exit Function

ErrHandler:

  'MsgBox "Folder " & sFolder & " could not be created." & vbCrLf & Err.Description, vbOKOnly, "Data Manager"
  
  BuildFolder = -1
  Set fso = Nothing
  Exit Function

End Function








