VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ICHCSII_PKI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum ValidationCode
  SIGNED_GOOD = 0
  SIGNED_NOT = 1
  SIGNED_BAD = 2
  SIGNED_NO_KEY = 3
End Enum

Public Type SignatureProperties
    status As ValidationCode
    UserId As String
    KeyID As String
    SignatureDTS As Date
End Type

Public Function GenerateKey(ByVal user_Ncid As String, ByVal sPassword As String, oUser As ICHCSII_User) As Boolean
End Function

Public Function Sign(ByVal TextToSign As String, ByVal sPassword As String, oUser As ICHCSII_User) As String
End Function

Public Function Validate(ByVal SignedText As String, ByVal Signature As String, oUser As ICHCSII_User, SigProp As SignatureProperties) As ValidationCode
End Function


