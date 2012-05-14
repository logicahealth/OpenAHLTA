VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CardMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Private Const MAX_CERTIFICATE_LENGTH = 2013
Public Event CardInserted(scardID As String)
Public Event CardRemoved(scardID As String)
Implements SSubTimer6.ISubclass

Private WithEvents m_Timer   As CTimer                     '-- internal, non-VBControl timer
Attribute m_Timer.VB_VarHelpID = -1
Private WithEvents m_ShutdownTimer As CTimer             '-- used to allow references to be dropped before shutting down
Attribute m_ShutdownTimer.VB_VarHelpID = -1
Private WithEvents m_oReader As cReader                  '-- the one and only reader
Attribute m_oReader.VB_VarHelpID = -1

Private m_hCardContext      As Long                          '-- keep track of the context for the life of this object
Private m_udtReaderCardStates() As SCARD_READERSTATEA   '-- might not be used... keep track of reader states from last check
Private m_FirstName         As String
Private m_LastName          As String
Private m_MiddleName        As String
Private m_DateOfBirth       As Date
Private m_Identifier        As String
Private m_IdentifierType    As Integer
Private m_ExpirationDate    As Date
Private m_Certificate       As String
Private m_PIN               As String
Private m_AID               As String
Private m_Algorithm         As Long
Private m_ReaderPresent     As Boolean

Const AID_CONTEXT           As String = "A0000000790300"
Const AID_ID                As String = "A0000000790100"
Const AID_SIG               As String = "A0000000790101"
Const AID_ENC               As String = "A0000000790102"

Public Enum HashType
    ht_NONE = -1
    ht_MD5 = 0
    ht_SHA1 = 1
End Enum
Private m_KillMe            As Boolean

                

Public Function Initialize() As Boolean
    Dim lngResult As Long
    On Error GoTo errhandler
    
    lngResult = SCardEstablishContext(SCARD_SCOPE_SYSTEM, lngNull, lngNull, m_hCardContext)
    If (lngResult <> SCARD_S_SUCCESS) Then
        Err.Raise lngResult, "Smart Card API", "Unable to establish Smart Card context."
    End If
    
    Set m_Timer = New CTimer
    m_Timer.Interval = 1000  '-- one second interval by default
    
    Set m_ShutdownTimer = New CTimer
    m_ShutdownTimer.Interval = 5000
    m_KillMe = False
    InstanceCounter = InstanceCounter + 1
        
    m_AID = AID_SIG
    m_Algorithm = BSI_RSA_NO_PAD
    
    Set m_oReader = New cReader
    
    If Not GetFullReaderList(m_hCardContext, m_oReader) Then
        m_Timer.Interval = 0
        m_ReaderPresent = False
        GoTo errhandler
    Else
        m_Timer.Interval = 1000  '-- one second interval by default
        m_ReaderPresent = True
    End If
    
    '-- to this point we've only used winscard.dll

'    Dim mscard As SCARDSSPLib.CSCard
'    Set mscard = New CSCard
'    mscard.AttachByReader m_oReader.ReaderName
'    'mscard.Context = m_hCardContext
'    Debug.Print mscard.Status
'    Dim mscardCmd As SCARDSSPLib.CSCardCmd
'    Set mscardCmd = New CSCardCmd
'    mscardcmd.
     
    m_ReaderPresent = True
    LoadErrorStrings
    LoadBSIErrorStrings
    ResetPersonalInfo
    Initialize = True
    Exit Function
    
errhandler:
    Initialize = False
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "Smart Card API", Err.Description
    End If
End Function

Public Function GetUserPIN(ByVal lLeft As Long, _
                           ByVal lTop As Long, _
                           Optional ByRef ErrCode As Long) As String
'-- takes a position (left, top) for form
'   returns the pin and/or and error code: success, user canceled, login failed, no card

    On Error GoTo errhandler
    Dim frm As frmCHCSIILogin
    Set frm = New frmCHCSIILogin
    With frm
        .Move lLeft, lTop
        .Show vbModal
        If .HasCancelled Then
            GetUserPIN = ""
            ErrCode = PIN_ERR_USERCANCELED
        Else
            GetUserPIN = .PIN
            ErrCode = PIN_ERR_SUCCESS
        End If
    End With
    Set frm = Nothing
    Exit Function

errhandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

Public Function cardPresent() As Boolean
    On Error GoTo errhandler

     cardPresent = m_oReader.cardPresent
    Exit Function
errhandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function ReaderPresent() As Boolean
    ReaderPresent = m_ReaderPresent
End Function

Private Sub Class_Terminate()
    On Error Resume Next
    Set m_Timer = Nothing
    SCardReleaseContext m_hCardContext
End Sub

Public Property Let Interval(ByVal iValue As Long)
    m_Timer.Interval = iValue
End Property

Public Property Get Interval() As Long
    Interval = m_Timer.Interval
End Property

Private Property Let ISubclass_MsgResponse(ByVal RHS As SSubTimer6.EMsgResponse)

End Property

Private Property Get ISubclass_MsgResponse() As SSubTimer6.EMsgResponse

End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

End Function

Private Sub m_oReader_CardInserted(CUID As String)
    RaiseEvent CardInserted(CUID)
End Sub

Private Sub m_oReader_CardRemoved(CUID As String)
    RaiseEvent CardRemoved(CUID)
End Sub

Private Sub m_ShutdownTimer_ThatTime()
    If m_KillMe Then
        InstanceCounter = InstanceCounter - 1
        If InstanceCounter = 0 Then
            End
        End If
    End If
End Sub


Private Sub m_Timer_ThatTime()
    '-- timer interval passed... check cards, raise events as needed

    '-- event priority:
    '--    1) CardRemoved
    '--    2) CardInserted
    Dim i As Long
    Dim lngResult  As Long


    Dim udtReaderCardStates() As SCARD_READERSTATEA

    ReDim udtReaderCardStates(0)

    udtReaderCardStates(0).dwCurrentState = 4  '-- force a 'change'
    udtReaderCardStates(0).dwEventState = 0
    udtReaderCardStates(0).szReader = m_oReader.ReaderName & vbNullChar


    lngResult = SCardGetStatusChangeA(m_hCardContext, 0&, udtReaderCardStates(0), 1)
    
    udtReaderCardStates(0).dwCurrentState = udtReaderCardStates(0).dwEventState
    
    Call m_oReader.UpdateStatus(udtReaderCardStates(0))
    m_oReader.ReaderCurrentState = udtReaderCardStates(0).dwEventState
    


    Exit Sub

End Sub


Private Function LoadCardData() As Long
    On Error GoTo Error_GetCardData
    
    '******************
    '-- Basic Demographics
    LoadCardData = GetDemographics
    If (LoadCardData <> 0) Then
        ResetPersonalInfo
        Exit Function
    End If
    '******************
    
    '******************
    '-- Certifcate
    LoadCardData = GetCertificate
    If (LoadCardData <> 0) Then
        ResetPersonalInfo
        Exit Function
    End If
    '******************
    
    LoadCardData = 0

Error_GetCardData:
    
    If Err.Number <> 0 Then
        LoadCardData = Err.Number
        Err.Clear
    End If
End Function

Private Sub ResetPersonalInfo()
    On Error Resume Next

    m_FirstName = ""
    m_LastName = ""
    m_MiddleName = ""
    m_DateOfBirth = 0
    m_Identifier = ""
    m_IdentifierType = 0
    m_ExpirationDate = 0
    m_Certificate = ""
End Sub

Public Property Get FirstName() As String
    FirstName = m_FirstName
End Property

Public Property Get MiddleName() As String
    MiddleName = m_MiddleName
End Property
    
Public Property Get LastName() As String
    LastName = m_LastName
End Property

Public Property Get DateOfBirth() As Date
    DateOfBirth = m_DateOfBirth
End Property

Public Property Get Identifier() As String
    Identifier = m_Identifier
End Property
    
Public Property Get IdentifierType() As Integer
    IdentifierType = m_IdentifierType
End Property

Public Property Get ExpirationDate() As Date
    ExpirationDate = m_ExpirationDate
End Property

Public Property Get Certificate() As String
    Certificate = m_Certificate
End Property

Private Function CCACCardDate(ByVal sDate As String) As Date
    If Len(sDate) <> 8 Then
        CCACCardDate = 0
    Else
        CCACCardDate = CDate(Mid(sDate, 7, 2) & "/" & Mid(sDate, 5, 2) & "/" & Mid(sDate, 1, 4))
    End If
End Function

'--- retain for future use
'---------------------------------------
'Private Function ByteArrayToHexString(aryBytes() As Byte) As String
'    Dim sReturn As String
'    Dim bCurrent As Byte
'    Dim i As Long
'
'    sReturn = ""
'    For i = 0 To UBound(aryBytes)
'        bCurrent = aryBytes(i) \ 16  'integer division
'        If bCurrent < 10 Then
'            sReturn = sReturn & Trim(Str(bCurrent))
'        Else
'            sReturn = sReturn & Chr(Asc("A") + bCurrent - 10)
'        End If
'
'        bCurrent = aryBytes(i) - (bCurrent * 16)
'        If bCurrent < 10 Then
'            sReturn = sReturn & Trim(Str(bCurrent))
'        Else
'            sReturn = sReturn & Chr(Asc("A") + bCurrent - 10)
'        End If
'    Next
'
'    ByteArrayToHexString = sReturn
'End Function

Private Function HexStringToByteArray(ByVal sStringToConvert As String) As Byte()
    Dim bReturn()   As Byte
    Dim sChar       As String       '-- charactor buffer
    Dim i           As Long         '-- General Pointer
    
    'TODO: Add error checking
    
    '-- Assumptions:
    '--    sStringToConvert will have a length divisible by 2
    '--    all characters will be in the hex range of 0 - F
    
    sStringToConvert = UCase(sStringToConvert)
    ReDim bReturn((Len(sStringToConvert) / 2) - 1)
    For i = 0 To UBound(bReturn)
        bReturn(i) = 0
        
        sChar = Mid(sStringToConvert, (i * 2) + 1, 1)
        If (sChar >= "A" And sChar <= "F") Then
            bReturn(i) = bReturn(i) + (16 * (10 + Asc(sChar) - 65))   '--  ASC("A") = 65
        ElseIf (sChar >= "0" And sChar <= "9") Then
            bReturn(i) = bReturn(i) + (16 * Val(sChar))
        End If
        
        sChar = Mid(sStringToConvert, (i * 2) + 2, 1)
        If (sChar >= "A" And sChar <= "F") Then
            bReturn(i) = bReturn(i) + (10 + Asc(sChar) - 65)
        ElseIf (sChar >= "0" And sChar <= "9") Then
            bReturn(i) = bReturn(i) + (Val(sChar))
        End If
    Next
    HexStringToByteArray = bReturn
End Function

Public Function SignText(ByVal PIN As String, _
                         ByVal TextToSign As String, _
                         ByRef Signature As String, _
                Optional ByRef ErrCode As Long, _
                Optional ByRef ErrMsg As String) As Boolean
                         
On Error GoTo errhandler
    Dim Base64      As New CBase64
    Dim bEncoded()  As Byte
    Dim SigArray()  As Byte
    Dim i           As Long
    
    m_PIN = PIN
    
    If Len(TextToSign) <> 128 Then
        SignText = False
        Exit Function
    End If
    
    SignText = Sign_String(TextToSign, SigArray(), ErrCode, ErrMsg)

    Base64.EncodeB64 SigArray, bEncoded
    Signature = ""
    For i = 0 To UBound(bEncoded)
        Signature = Signature & Chr(bEncoded(i))
    Next

    Exit Function
errhandler:
   
    ErrMsg = BSIErrStr(Err.Number)
    SignText = False
                         
End Function
                        
                         
Public Function ValidatePIN(ByVal PIN As String, _
                   Optional ByRef CardID As String, _
                   Optional ByRef ErrCode As Long, _
                   Optional ByRef ErrMsg As String) As Boolean


On Error GoTo errhandler
    Dim oCAC As DMDCCACLib.CommonAccessCard
    Dim lVal As Long
    
    Set oCAC = New DMDCCACLib.CommonAccessCard
    With oCAC
        .connectToCard m_oReader.ReaderName, 0
        .getCardStatus m_oReader.ReaderName, lVal
        .logonToCard StrConv(PIN, vbFromUnicode), Len(PIN)
    End With
    
    Call GetCardID(m_oReader.ReaderName, CardID)
    
    ValidatePIN = True
    ErrCode = 0
    ErrMsg = ""
    oCAC.disconnectFromCard
    Exit Function
errhandler:
   
    ErrMsg = BSIErrStr(Err.Number)
    ValidatePIN = False
    oCAC.disconnectFromCard
    
End Function

Private Function GetDemographics() As Long
    '-- load the basic demographic information from the card into this object's
    '--    properties
    '--   m_PIN must have been set
    
    Dim oCAC As DMDCCACLib.CommonAccessCard
    Dim sValue As String
    
    On Error GoTo errhandler
    
    If m_oReader Is Nothing Then
        GetDemographics = SCARD_E_NO_READERS_AVAILABLE
        Exit Function
    End If

    
    Set oCAC = New DMDCCACLib.CommonAccessCard
    With oCAC
        .connectToCard m_oReader.ReaderName, 0
        .getCardStatus m_oReader.ReaderName, 0
        .logonToCard StrConv(m_PIN, vbFromUnicode), Len(m_PIN)
    
        '-- firstname
        .getFirstName sValue
        m_FirstName = RemoveNull(StrConv(sValue, vbUnicode))
        
        '-- middlename
        .getMiddleName sValue
        m_MiddleName = RemoveNull(StrConv(sValue, vbUnicode))
        
        '-- lastname
        .getLastName sValue
        m_LastName = RemoveNull(StrConv(sValue, vbUnicode))
        
        '-- DateOfBirth
        .getDateOfBirth sValue
        m_DateOfBirth = CCACCardDate(RemoveNull(StrConv(sValue, vbUnicode)))
        
        '-- Identifier (SSN?)
        oCAC.getPersonIdentifier sValue
        m_Identifier = RemoveNull(StrConv(sValue, vbUnicode))
        
        '-- IdentifierType (?)
        .getPersonIdentifierTypeCode sValue
        m_IdentifierType = Val(RemoveNull(StrConv(sValue, vbUnicode)))
        
        '-- ExpirationDate (?)
        .getDateCacDataExpires sValue
        m_ExpirationDate = CCACCardDate(RemoveNull(StrConv(sValue, vbUnicode)))
        
        .disconnectFromCard
    End With
    
    If m_ExpirationDate < Date Then   '-- PER SCR 50531
        Set oCAC = Nothing
        GetDemographics = 30040 '-- card expired
        Err.Raise 30040, "Card Data", "Card has expired (Expiration Date: " & m_ExpirationDate & ")"
    Else
        GetDemographics = 0
    End If

    Set oCAC = Nothing
    Exit Function
    
errhandler:
    GetDemographics = Err.Number
    Set oCAC = Nothing

End Function

Public Function GetCertificate() As Long


    Dim lVal        As Long
    Dim lReturn     As Long
    Dim hCard       As Long
    Dim lLenAID     As Long
    Dim Base64      As CBase64
    Dim bEncoded()  As Byte
    Dim i           As Long
    Dim bCertificate() As Byte
    
    '-- blank the cert value in case of error
    '-- m_Certificate = ""
    On Error GoTo GetCertificate_Problem
    
    Set Base64 = New CBase64
    '-- connect
    If m_oReader Is Nothing Then
        GetCertificate = SCARD_E_NO_READERS_AVAILABLE
        Exit Function
    End If
    
    lReturn = AC_XSI_UtilConnect(hCard, m_oReader.ReaderName, 0&)
    GetCertificate = lReturn
    If lReturn <> BSI_OK Then
        Exit Function
    End If

    '-- acquire a context
    Dim bAID() As Byte
    bAID = HexStringToByteArray(m_AID)
    lLenAID = Len(m_AID) / 2

    Dim BSIAuthenticator(0) As BSI_Authenticator
    With BSIAuthenticator(0)
        .unACRType = BSI_ACR_PIN
        .usAuthValue = StrConv(m_PIN, vbFromUnicode)
        .unAuthLen = Len(m_PIN)
        .usKeyValue = 0&
    End With

    lReturn = gscBsiUtilAcquireContext(hCard, bAID(0), lLenAID, BSIAuthenticator(0), 0&)
    GetCertificate = lReturn
    If lReturn <> BSI_OK Then
        GoTo GetCertificate_Problem
    End If
    
    '-- read and store the certificate
    
    ReDim bCertificate(MAX_CERTIFICATE_LENGTH)
    lVal = MAX_CERTIFICATE_LENGTH
    
    lReturn = gscBsiPkiReadCertificate(hCard, bAID(0), lLenAID, bCertificate(0), lVal)
    GetCertificate = lReturn
    If lReturn <> BSI_OK Then
        GoTo GetCertificate_Problem
    End If
        
    ReDim Preserve bCertificate(lVal - 1)
    Base64.EncodeB64 bCertificate, bEncoded
    m_Certificate = ""
    For i = 0 To UBound(bEncoded)
        m_Certificate = m_Certificate & Chr(bEncoded(i))
    Next
    
    '--code to dump to a binary file... just used for testing
    'Open "c:\test.bin" For Binary As #1
    'For lVal = 0 To UBound(bCertificate)
    '    Put #1, lVal + 1, bCertificate(lVal)
    'Next
    'Close #1


    
'    Open "c:\test.pem" For Output As #1
'    m_Certificate = "-----BEGIN CERTIFICATE-----" & vbCrLf & m_Certificate
'    m_Certificate = m_Certificate & vbCrLf & "-----END CERTIFICATE-----"
'
'    Print #1, m_Certificate
'    Close #1
    '-- release the context
    '--   (this seems to be locking the machine... i believe that the disconnect
    '--       will release all contexts, automatically)
    'lReturn = gscBsiUtilReleaseContext(hCard, bAID(0), lLenAID)

    
    lReturn = AC_XSI_UtilDisconnect(hCard)
    GetCertificate = lReturn
    Exit Function
    
GetCertificate_Problem:
    '-- something bombed... we're just going to disconnect, and get out
    '   return code should already show the error
    
    On Error Resume Next
    lReturn = AC_XSI_UtilDisconnect(hCard)
    '-- not reseting the return value, so it should still contain the error code
    
End Function


Public Property Let PIN(ByVal sNewValue As String)
    m_PIN = sNewValue
End Property

Public Function CHCSII_Login(ByVal PIN As String, _
                             ByRef Str As String, _
                             ByRef Sig As String, _
                             ByRef Cert As String, _
                             ByRef ErrMsg As String) As Boolean
    Dim lRes        As Long
    Dim ErrCode     As Long
    Dim bEncoded()  As Byte
    Dim Base64      As New CBase64
    Dim SigArray()  As Byte
    Dim i           As Long
    
    m_PIN = PIN
    If m_PIN = "" Then Exit Function
    
    lRes = LoadCardData   '-- logs on to card with m_pin and gets cert and demographics
    If lRes <> 0 Then
        ErrMsg = InterpretSCardError(lRes)
        Exit Function
    End If
   
    Cert = m_Certificate
    
    CHCSII_Login = Sign_String(Str, SigArray(), ErrCode, ErrMsg)
    
       '- set the return value
    Base64.EncodeB64 SigArray, bEncoded
    Sig = ""
    For i = 0 To UBound(bEncoded)
        Sig = Sig & Chr(bEncoded(i))
    Next
    
End Function

Private Function Sign_String(ByRef Str As String, _
                             ByRef Sig() As Byte, _
                             ByRef ErrCode As Long, _
                             ByRef ErrMsg As String) As Boolean
    '-- uses the CAC card to sign a string, and
    '   returns a blank string on failure
    
    On Error GoTo errhandler:
    If Len(Str) <> 128 Then Exit Function
    
    Dim lVal        As Long
    Dim lReturn     As Long
    Dim hCard       As Long
    Dim lLenAID     As Long
    Dim bArray()    As Byte
    Dim bStr()      As Byte
    Dim i           As Long
    
    If m_oReader Is Nothing Then
        ErrCode = SCARD_E_NO_READERS_AVAILABLE
        ErrMsg = ErrStrings(SCARD_E_NO_READERS_AVAILABLE)
        Sign_String = False
        Exit Function
    End If
        
    lReturn = AC_XSI_UtilConnect(hCard, m_oReader.ReaderName, 0&)
    If lReturn <> BSI_OK Then
        ErrCode = lReturn
        ErrMsg = BSIErrStr(lReturn)
        Exit Function
    End If

    '-- acquire a context
    Dim bAID() As Byte
    bAID = HexStringToByteArray(m_AID)
    lLenAID = Len(m_AID) / 2

    Dim BSIAuthenticator(0) As BSI_Authenticator
    With BSIAuthenticator(0)
        .unACRType = BSI_ACR_PIN
        .usAuthValue = StrConv(m_PIN, vbFromUnicode)
        .unAuthLen = Len(m_PIN)
        .usKeyValue = 0&
    End With

    lReturn = gscBsiUtilAcquireContext(hCard, bAID(0), lLenAID, BSIAuthenticator(0), 0&)
    If lReturn <> BSI_OK Then
        ErrCode = lReturn
        ErrMsg = BSIErrStr(lReturn)
        GoTo errhandler
    End If
    
    '-- build the string
    ReDim bStr(127)
    For i = 0 To UBound(bStr)
        bStr(i) = Asc(Mid$(Str, i + 1, 1))
    Next i
    
    ReDim bArray(2043)
    lVal = UBound(bArray) + 1
    ReDim Sig(2043)
    lVal = UBound(Sig) + 1
    '-- sign the string
    lReturn = gscBsiPkiCompute(hCard, bAID(0), lLenAID, m_Algorithm, bStr(0), UBound(bStr) + 1, Sig(0), lVal)
    If lReturn <> BSI_OK Then
        ErrCode = lReturn
        ErrMsg = BSIErrStr(lReturn)
        GoTo errhandler
    End If
    
    ReDim Preserve bArray(lVal - 1)
    ReDim Preserve Sig(lVal - 1)
    '--code to dump to a binary file... just used for testing
'    Open "c:\sig.bin" For Binary As #1
'    For lVal = 0 To UBound(bArray)
'        Put #1, lVal + 1, bArray(lVal)
'    Next
'    Close #1
    
    

    
    'release the context
    '--   (this seems to be locking the machine... i believe that the disconnect
    '--       will release all contexts, automatically)
    'lReturn = gscBsiUtilReleaseContext(hCard, bAID(0), lLenAID)
    Sign_String = True
    lReturn = AC_XSI_UtilDisconnect(hCard)
    Exit Function
errhandler:
    On Error Resume Next
    '-- disconnect
    lReturn = AC_XSI_UtilDisconnect(hCard)
    
    
End Function


Public Sub Terminate()
    m_KillMe = True
End Sub
