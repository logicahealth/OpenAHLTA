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
Attribute VB_Name = "Encryption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Const EncryptPassword = "dskfjd3e9322309323212!@!@$@#!@@!SDE@#23s"
Public Function DecryptText(StrToDecrypt As String) As String
On Error GoTo ErrHandler
    
    Dim TheSecret As String

    TheSecret = StrToDecrypt
    DecryptText = EncryptDecrypt(TheSecret, EncryptPassword)

GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "Encryptor", "DecryptText", Err.Number, Err.Description
CleanUp:
End Function

Private Function EncryptDecrypt(secret As String, Password As String) As String
On Error GoTo ErrHandler
    
    Dim L As Long
    Dim x As Long
    Dim Char As String
    
    L = Len(Password)
    
    For x = 1 To Len(secret)
        Char = Asc(Mid$(Password, (x Mod L) - L * ((x Mod L) = 0), 1))
        Mid$(secret$, x, 1) = Chr$(Asc(Mid$(secret, x, 1)) Xor Char)
    Next
    
    EncryptDecrypt = secret
    
GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "Encryptor", "EncryptDecrypt", Err.Number, Err.Description
CleanUp:
End Function

Public Function EncryptText(StrToEncrypt As String) As String
On Error GoTo ErrHandler

    Dim TheString As String
    
    TheString = StrToEncrypt
    EncryptText = EncryptDecrypt(TheString, EncryptPassword)

GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "Encryptor", "EncryptText", Err.Number, Err.Description
CleanUp:
End Function

 

