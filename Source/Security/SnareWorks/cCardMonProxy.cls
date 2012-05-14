VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCardMonProxy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private bCac_Avail  As Boolean       '-- false if no card reader is found
Private WithEvents CAC_Monitor As CAC_Card_Monitor.CardMonitor
Attribute CAC_Monitor.VB_VarHelpID = -1

Public Event CardInserted(scardID As String)
Public Event CardRemoved(scardID As String)

Private Sub CAC_Monitor_CardInserted(scardID As String)
    RaiseEvent CardInserted(scardID)
End Sub

Private Sub CAC_Monitor_CardRemoved(scardID As String)
    RaiseEvent CardRemoved(scardID)
End Sub

Private Function Initialize() As Boolean
    On Error GoTo errhandler
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    Dim sysdir As String
    sysdir = Space(255)
    
    Dim l As Long
    l = GetSystemDirectory(sysdir, 255)
    If l > 0 Then
        sysdir = Left$(sysdir, l)
    End If
    
    If Not (fso.FileExists(sysdir & "\acbsiprov.dll") And fso.FileExists(sysdir & "\dmdcCAC.dll ")) Then
        Initialize = False
        bCac_Avail = False
        Exit Function
    End If
    
    Set CAC_Monitor = New CAC_Card_Monitor.CardMonitor
    
    If CAC_Monitor.Initialize Then
        bCac_Avail = True
    Else
        bCac_Avail = False
        Set CAC_Monitor = Nothing
    End If
    Initialize = True
    Exit Function
    
errhandler:
    
    Initialize = False
    bCac_Avail = False
End Function
Public Property Get CardPresent() As Boolean
    On Error GoTo errhandler
    If Not bCac_Avail Then Exit Property

    CardPresent = CAC_Monitor.CardPresent
    Exit Property
    
errhandler:
    CardPresent = False
End Property
 
Public Function SignText(ByVal PIN As String, _
                         ByVal TextToSign As String, _
                         ByRef Signature As String, _
                Optional ByRef ErrCode As Long, _
                Optional ByRef ErrMsg As String) As Boolean
                
    On Error GoTo errhandler
    If Not bCac_Avail Then Exit Function
    Dim oCrypto As CCrypto
    Set oCrypto = New CCrypto
    Dim str As String
                   '-- the CAC wants only a 128 byte string
    str = oCrypto.GetHash(TextToSign, True)     '-- this gives us 32
    str = str & str                             '-- now 64
    str = str & str                             '-- now 128
   
    SignText = CAC_Monitor.SignText(PIN, str, Signature, ErrCode, ErrMsg)
    Exit Function
    
errhandler:
    SignText = False
    ErrMsg = "Unknown problem occurred in the signtext function"
End Function
 

Public Property Get ReaderPresent() As Boolean
    On Error GoTo errhandler
    If Not bCac_Avail Then Exit Property

    ReaderPresent = CAC_Monitor.ReaderPresent
    Exit Property
    
errhandler:
    ReaderPresent = False
End Property

Public Function CHCSII_Login(ByVal PIN As String, _
                             ByRef str As String, _
                             ByRef Sig As String, _
                             ByRef Cert As String, _
                             ByRef ErrMsg As String) As Boolean
                             
    On Error GoTo errhandler
    If Not bCac_Avail Then Exit Function

    CHCSII_Login = CAC_Monitor.CHCSII_Login(PIN, str, Sig, Cert, ErrMsg)
    Exit Function
    
errhandler:
    CHCSII_Login = False
    ErrMsg = "Unknown problem occurred in the CAC logon function"
End Function


Public Function ValidatePIN(ByVal PIN As String, _
                   Optional ByRef CardID As String, _
                   Optional ByRef ErrCode As Long, _
                   Optional ByRef ErrMsg As String) As Boolean
                   
    On Error GoTo errhandler
    If Not bCac_Avail Then Exit Function

    ValidatePIN = CAC_Monitor.ValidatePIN(PIN, CardID, ErrCode, ErrMsg)
    Exit Function
    
errhandler:
    ValidatePIN = False
    ErrMsg = "Unknown problem occurred in the CAC logon function"
                       
End Function

Private Sub Class_Initialize()
    Call Initialize
End Sub

Friend Sub Terminate()
    If Not CAC_Monitor Is Nothing Then
        CAC_Monitor.Terminate
    End If
End Sub

