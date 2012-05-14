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
Attribute VB_Name = "GatewayLGS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'LGS Gateway
'data layer support tool responsible for executing data layer calls to the LGS soap server
'
'connect to server
'format params into soap package
'execute call against web service
'decompress result
'process output params
'return results to concrete domain mapper object
'
'maintains security with token
'needs ValidateUser method that will accept username/password and store resulting token
'
'responsible for timeout / retry message
'
'
'
'Configuration Options per session:
' AllowRetry
' Timeout / RetryTimeout
' URL
' Namespace
' Compress
' some type of config to store what errors to rethrow and which to retry
'
'
'Public Methods
' Init(URL, NameSpace, Compress)
' Execute(Method, Parameters, AllowRetry, Timeout, Compress)
' ValidateUser(Username, Password)

Private Enum eErrorAction
    eErrorAction_Rethrow
    eErrorAction_Retry
    eErrorAction_Reconnect
End Enum

Private Const DEBUGMODE = True

Private Const USE_CURRENT_SOAP_TIMEOUT = -1

Private Const DEFAULT_SOAP_TIMEOUT = 300000              '5 mins
Private Const DEFAULT_SOAP_TIMEOUT_RETRIES = 600000      '10 mins
Private Const DEFAULT_NAMESPACE = "http://tempuri.org/"
Private Const DEFAULT_USECOMPRESSION = True
Private Const DEFAULT_ALLOWRETRY = True
Private Const DEFAULT_ERRORMESSAGE_CONNECTIVITY = "Connection to Server Lost.  Click 'RETRY' to resubmit transaction or click 'CANCEL' to exit application."
Private Const DEFAULT_GZIP = True
Private Const DEFAULT_RAISESHUTDOWN = True

Private Const DEFAULT_LOGONMETHODNAME = "ValidateUser"
Private Const DEFAULT_LOGONUSERNAMEPARAM = "AccessCode"
Private Const DEFAULT_LOGONPASSWORDPARAM = "VerifyCode"
Private Const DEFAULT_LOGONTOKENPARAM = "Token"

Private Const MESSAGEBOX_ERROR = False          'display message box on error?

Private m_oSoapConnector As SoapConnector30
Private m_sURL As String
Private m_sNamespace As String
Private m_bRequestCompression As Boolean
Private m_bAllowRetry As Boolean
Private m_bGzip As Boolean
Private m_lSoapTimeout As Long, m_lSoapTimeoutRetries As Long
Private m_bRaiseShutDownOnCancel As Boolean

Private m_sWebServiceLogonMethod As String
Private m_sWebServiceLogonUsernameParam As String
Private m_sWebServiceLogonPasswordParam As String
Private m_sWebServiceLogonTokenParam As String

Private m_sToken As String
Private m_sUserName As String, m_sPassword As String

Public Event RequestShutdown()

'*** initialization routines
Private Sub Class_Initialize()
    Namespace = DEFAULT_NAMESPACE
    RequestCompression = DEFAULT_USECOMPRESSION
    AllowRetry = DEFAULT_ALLOWRETRY
    SoapTimeout = DEFAULT_SOAP_TIMEOUT
    SoapTimeoutRetries = DEFAULT_SOAP_TIMEOUT_RETRIES
    Gzip = DEFAULT_GZIP
    
    LogonMethodName = DEFAULT_LOGONMETHODNAME
    LogonUsernameParam = DEFAULT_LOGONUSERNAMEPARAM
    LogonPasswordParam = DEFAULT_LOGONPASSWORDPARAM
    LogonTokenParam = DEFAULT_LOGONTOKENPARAM
End Sub

Public Sub Init(Optional WebServiceURL As String = vbNullString, Optional WebServiceNamespace As String = vbNullString, _
                Optional WebServiceTimeout As Long = USE_CURRENT_SOAP_TIMEOUT, Optional WebServiceTimeoutRetries As Long = USE_CURRENT_SOAP_TIMEOUT)
    If WebServiceURL <> "" Then URL = WebServiceURL
    If WebServiceNamespace <> "" Then Namespace = WebServiceNamespace
    If SoapTimeout <> USE_CURRENT_SOAP_TIMEOUT Then SoapTimeout = WebServiceTimeout
    If SoapTimeoutRetries <> USE_CURRENT_SOAP_TIMEOUT Then SoapTimeoutRetries = WebServiceTimeoutRetries
    
    If URL = "" Then Err.Raise LGSWEBSERVER_ERRORCODE_INVALID_WEBSERVICE_URL, "Invalid web service url"
    
    Connect
End Sub

Private Sub Connect()
    Const Method = "Connect"
    
    On Error GoTo Error_Handler
    
    Set m_oSoapConnector = New HttpConnector30
    With SoapConnector
        .Property("EndPointURL") = URL
        .Property("Timeout") = SoapTimeout
        .Connect
    End With
    
    Exit Sub
    
Error_Handler:
    Dim sError As String
    sError = Err.Description
    ErrorHandler sError, , Method, True, False
End Sub
'*** initialization routines

'*** termination routines
Private Sub Class_Terminate()
    Disconnect
End Sub

Private Sub Disconnect()
    If Not m_oSoapConnector Is Nothing Then m_oSoapConnector.Reset
    Set SoapConnector = Nothing
End Sub
'*** termination routines


'*** local properties
Private Property Get SoapConnector() As SoapConnector30
    Set SoapConnector = m_oSoapConnector
End Property
    
Private Property Get URL() As String
    URL = m_sURL
End Property
Private Property Let URL(sURL As String)
    m_sURL = sURL
End Property
    
Private Property Get Namespace() As String
    Namespace = m_sNamespace
End Property
Private Property Let Namespace(sNamespace As String)
    m_sNamespace = sNamespace
End Property
    
Private Property Get SoapTimeout() As Long
    SoapTimeout = m_lSoapTimeout
End Property
Private Property Let SoapTimeout(lSoapTimeout As Long)
    m_lSoapTimeout = lSoapTimeout
End Property
Private Property Get SoapTimeoutRetries() As Long
    SoapTimeout = m_lSoapTimeoutRetries
End Property
Private Property Let SoapTimeoutRetries(lSoapTimeoutRetries As Long)
    m_lSoapTimeoutRetries = lSoapTimeoutRetries
End Property
    
Private Property Get RequestCompression() As Boolean
    RequestCompression = m_bRequestCompression
End Property
Private Property Let RequestCompression(bRequestCompression As Boolean)
    m_bRequestCompression = bRequestCompression
End Property
    
Private Property Get AllowRetry() As Boolean
    AllowRetry = m_bAllowRetry
End Property
Private Property Let AllowRetry(bAllowRetry As Boolean)
    m_bAllowRetry = bAllowRetry
End Property

Private Property Get RaiseShutdownMethodOnCancel() As Boolean
    RaiseShutdownMethodOnCancel = m_bRaiseShutDownOnCancel
End Property
Private Property Let RaiseShutdownMethodOnCancel(bRaiseShutdownMethodOnCancel As Boolean)
    m_bRaiseShutDownOnCancel = bRaiseShutdownMethodOnCancel
End Property

Private Property Get Gzip() As Boolean
    Gzip = m_bGzip
End Property
Private Property Let Gzip(bGzip As Boolean)
    m_bGzip = bGzip
End Property

Private Property Get LogonMethodName() As String
    LogonMethodName = m_sWebServiceLogonMethod
End Property
Private Property Let LogonMethodName(WebServiceLogonMethod As String)
    m_sWebServiceLogonMethod = WebServiceLogonMethod
End Property
Private Property Get LogonUsernameParam() As String
    LogonUsernameParam = m_sWebServiceLogonUsernameParam
End Property
Private Property Let LogonUsernameParam(WebServiceLogonUsernameParam As String)
    m_sWebServiceLogonUsernameParam = WebServiceLogonUsernameParam
End Property
Private Property Get LogonPasswordParam() As String
    LogonPasswordParam = m_sWebServiceLogonPasswordParam
End Property
Private Property Let LogonPasswordParam(WebServiceLogonPasswordParam As String)
    m_sWebServiceLogonPasswordParam = WebServiceLogonPasswordParam
End Property
Private Property Get LogonTokenParam() As String
    LogonTokenParam = m_sWebServiceLogonTokenParam
End Property
Private Property Let LogonTokenParam(WebServiceLogonTokenParam As String)
    m_sWebServiceLogonTokenParam = WebServiceLogonTokenParam
End Property
'*** local properties

'*** public function
Public Property Get LogonUsername() As String
    LogonUsername = m_sUserName
End Property
Public Property Let LogonUsername(Username As String)
    m_sUserName = Username
End Property
Public Property Get LogonPassword() As String
    LogonPassword = m_sPassword
End Property
Public Property Let LogonPassword(Password As String)
    m_sPassword = Password
End Property


Public Function Logon(Optional Username As String = vbNullString, Optional Password As String = vbNullString, Optional RaiseErrorOnFailure As Boolean = True) As String
    Dim Parameters As Object
    
    On Error GoTo Error_Handler
    
    If Username <> "" Then LogonUsername = Username
    If Password <> "" Then LogonPassword = Password
    
    Parameters.Add LogonUsernameParam, LogonUsername
    Parameters.Add LogonUsernameParam, LogonPassword
    Logon = Execute(LogonMethodName, Parameters, False)
    Exit Function
    
Error_Handler:
    If RaiseErrorOnFailure Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Logon = "ERROR:" & Err.Description
    End If
End Function

Public Function Execute(ByVal Method As String, Optional Parameters As Object = Nothing, _
                        Optional AllowRetryForThisCall As Boolean = True, Optional SoapTimeoutForThisCall As Long = USE_CURRENT_SOAP_TIMEOUT, _
                        Optional CompressionForThisCall As Boolean = True, Optional GzipThisCall As Boolean = True) As String
    Dim sResultWrap As String
    Dim sResultUnWrap As String
    
    On Error GoTo Error_Handler
    
    Parameters.Add LogonTokenParam, m_sToken
    
    If SoapTimeoutForThisCall = USE_CURRENT_SOAP_TIMEOUT Then
        SoapTimeoutForThisCall = SoapTimeout
    End If
    
    CompressionForThisCall = (CompressionForThisCall And RequestCompression)
    GzipThisCall = (GzipThisCall And Gzip)
    
    Timings Method, "Begin execute (" & Method & ")"
    'sReturn = SoapClient.Execute(Method, oParams, bAllowRetryForThisCall, lSoapTimeout, Option_RequestCompression)
    sResultWrap = Execute_Internal(Method, Parameters, SoapTimeoutForThisCall, CompressionForThisCall, GzipThisCall)
    Timings Method, "End execute"
    
    Timings Method, "Begin unwrap"
    sResultUnWrap = Unwrap(sResultWrap)
    Timings Method, "End unwrap"
    
    Execute = sResultUnWrap
    
    m_sToken = Parameters.Item(LogonTokenParam)
    
    Exit Function
        
Error_Handler:
    
    Dim sOriginalError As String
    Dim sError As String
    Dim lError As Long
    Dim Action As eErrorAction
    
    sOriginalError = Err.Description
    sError = sOriginalError
    lError = Err.Number
    Action = eErrorAction_Rethrow
    
    Debug.Print "Soap client error: " & sError
    
    '*** this needs to be replaced with something more dynamic
    If (UCase(sError) Like UCase("*Unable to validate token*")) Or (UCase(sError) Like "*INVALID TOKEN*") Or (lError = LGSWEBSERVER_ERRORCODE_INVALID_SECURITY_CONNECTION) Then
        sError = "Invalid security connection."
        lError = LGSWEBSERVER_ERRORCODE_INVALID_SECURITY_CONNECTION
        Action = eErrorAction_Reconnect
    ElseIf UCase(sError) Like UCase("*Error connecting to the Connection Manager*") Then
        sError = "Unable to establish connection to CHCS host [CANNOT CONNECT TO CONNECTION MANAGER]."
        lError = LGSWEBSERVER_ERRORCODE_CHCS_NOT_AVAILABLE
        Action = eErrorAction_Retry
    ElseIf (UCase(sError) Like "*REMOTE SERVER UNREACHABLE*") Or (lError = LGSWEBSERVER_ERRORCODE_CHCS_NOT_AVAILABLE) Then
        sError = "Unable to establish connection to CHCS host [REMOTE SERVER UNREACHABLE]."
        lError = LGSWEBSERVER_ERRORCODE_CHCS_NOT_AVAILABLE
        Action = eErrorAction_Retry
    ElseIf UCase(sError) Like "*LOADING THE REQUEST INTO SOAPREADER FAILED*" Then
        sError = "Unable to establish connection to LGS [LOADING THE REQUEST INTO SOAPREADER FAILED]"
        lError = LGSWEBSERVER_ERRORCODE_LGS_NOTAVAILABLE
        Action = eErrorAction_Retry
    ElseIf UCase(sError) Like "*CONNECTION TIME OUT*" Then
        sError = "Connection time out between the client workstation and the LGS."
        lError = LGSWEBSERVER_ERRORCODE_TIMEOUT_CW_TO_LGS
        Action = eErrorAction_Retry
    ElseIf UCase(sError) Like UCase("*Unspecified HTTP error*") Then
        sError = "Unable to establish connection to LGS [UNSPECIFIED HTTP ERROR]"
        lError = LGSWEBSERVER_ERRORCODE_LGS_NOTAVAILABLE
        Action = eErrorAction_Retry
    ElseIf UCase(sError) Like UCase("*Loading of the WSDL file failed*") Then
        sError = "Unable to load WSDL file."
        lError = LGSWEBSERVER_ERRORCODE_LGS_NOTAVAILABLE
        Action = eErrorAction_Retry
    ElseIf UCase(sError) Like UCase("*There is an error in XML document*") Then
        sError = "Configuration error on method " & Method
        lError = LGSWEBSERVER_ERRORCODE_CONFIGURATION
        Action = eErrorAction_Retry
    ElseIf UCase(sError) Like UCase("*Host Not found*") Then
        sError = "Unable to establish connection to LGS [HOST NOT FOUND]"
        lError = LGSWEBSERVER_ERRORCODE_LGS_NOTAVAILABLE
        Action = eErrorAction_Retry
    ElseIf UCase(sError) Like UCase("*Method 'EndMessage' of object 'ISoapConnector' failed*") Then
        sError = "Unable to establish connection to LGS [ENDMESSAGE FAILED]"
        lError = LGSWEBSERVER_ERRORCODE_LGS_NOTAVAILABLE
        Action = eErrorAction_Retry
    ElseIf UCase(sError) Like UCase("*Unable to execute mobject call*") And UCase(sError) Like UCase("*Object reference not set to an instance of an object*") Then
        'when this error occurs, unable to perform mobject calls via eclient - port unavailable
        sError = "Unable to connect to execute mobject call [Object reference not set to an instance of an object]."
        lError = LGSWEBSERVER_ERRORCODE_CHCS_NOT_AVAILABLE
        Action = eErrorAction_Retry
    ElseIf UCase(sError) Like UCase("*Unable to execute mobject call*") And UCase(sError) Like UCase("*Cannot create activex component*") Then
        'when this error occurs, unable to perform mobject calls via eclient - port unavailable
        sError = "Unable to connect to execute mobject call [Cannot create activeX component]."
        lError = LGSWEBSERVER_ERRORCODE_CHCS_NOT_AVAILABLE
        Action = eErrorAction_Retry
    ElseIf UCase(sError) Like UCase("*Invalid procedure call or argument*") Then
        'this is the error that occurs if you shut down IIS / ASPNET while running the request
        sError = "Internal LGS error"
        lError = LGSWEBSERVER_ERRORCODE_INTERNALLGSERROR
        Action = eErrorAction_Retry
    ElseIf UCase(sError) Like UCase("*System error*") Then
        'handle this like a timeout
        sError = "Internal LGS error (system error)"
        lError = LGSWEBSERVER_ERRORCODE_INTERNALLGSERROR
        Action = eErrorAction_Retry
    ElseIf UCase(sError) Like UCase("*Server did not recognize the value of HTTP Header SOAPAction*") Then
        sError = "Action not supported (" & Method & ")"
        lError = LGSWEBSERVER_ERRORCODE_METHOD_NOT_SUPPORTED
        Action = eErrorAction_Rethrow
    ElseIf UCase(sError) Like UCase("*Object reference not set*") And UCase(Method) = "RETRIEVERECORDSET" Then
        sError = "Error accessing database (" & sError & ")"
        lError = LGSWEBSERVER_ERRORCODE_GENERICLCD
        Action = eErrorAction_Retry
    ElseIf UCase(sError) Like UCase("*unable to execute mobject call*") And UCase(Error) Like "*server threw an exception*" Then
        'scr 35409 - 15MAY2003 - j.ray - added retry logic to "server threw an exception" error
        sError = "Unable to establish connection to CHCS host [SERVER THREW AN EXCEPTION]."
        lError = LGSWEBSERVER_ERRORCODE_GENERICMOBJECTERROR
        Action = eErrorAction_Retry
    ElseIf UCase(sError) Like UCase("*Catastrophic failure*") Then
        If UCase(sError) Like ("*connection failed*") Then
          sError = "Unable to establish connection to LCD."
        Else
          sError = "Unable to establish connection to external system."
        End If
        lError = LGSWEBSERVER_ERRORCODE_CATASTROPHICFAILURE
        Action = eErrorAction_Retry
    ElseIf UCase(sError) Like UCase("*Cannot uncompress corrupt data*") Then
        sError = "Error during decompression.  method:[" & Method & "] err:[" & sError & "]"
        lError = LGSWEBSERVER_ERRORCODE_COMPRESSION
        Action = eErrorAction_Retry
    End If
    
    If Not (AllowRetry And AllowRetryForThisCall) And Action = eErrorAction_Retry Then Action = eErrorAction_Rethrow
    
    Select Case Action
        Case eErrorAction_Retry, eErrorAction_Reconnect
            If Action = eErrorAction_Reconnect Then
                Dim ReconnectResult As String
                ReconnectResult = Logon(, , False)
                If UCase(ReconnectResult) Like "ERROR*" Then
                    sError = sError & " (RC=F)"
                Else
                    sError = sError & " (RC=T)"
                End If
            End If
            
            
            If DisplayErrorMessage(CStr(lError), sError & vbCrLf & vbCrLf & sOriginalError) = vbRetry Then
                SoapTimeoutForThisCall = SoapTimeoutRetries
                Resume
            Else
                If RaiseShutdownMethodOnCancel Then
                    RaiseEvent RequestShutdown
                End If
            End If
        Case Else
            'no special processing for case else
    End Select
    
    'if here, then need to rethrow error
    Err.Raise lError, Method, sError
        
End Function

Private Function Execute_Internal(ByVal Method As String, Parameters As Object, SoapTimeoutForThisCall As Long, AllowCompressionForThisCall As Boolean, GzipThisCall As Boolean) As String
    Dim Serializer As SoapSerializer30
    Dim Reader As SoapReader30
    Dim Paramater As Object
    Dim oStringAttachment As MSSOAPLib30.StringAttachment30
    
    Const SOAP_PARAM_DIRECTION__INPUT = "?"
    Const SOAP_PARAM_DATATYPE__STRINGATTACHMENT = "?"
    Const SOAP_PARAM_DIRECTION__OUTPUT = "?"
    
    On Error GoTo Error_Handler
    
    If SoapConnector Is Nothing Then
        Connect
    End If
    
    With SoapConnector
        'this request that web server compress the http result - not the same as our wrap / unwrap methods
        If GzipThisCall Then
            .Property("RequestHTTPHeader") = "Accept-Encoding: gzip, deflate"
        End If
    
        .Property("SoapAction") = Namespace & Method
        .Property("Timeout") = SoapTimeoutForThisCall
    
        .BeginMessage
    
        Set Serializer = New SoapSerializer30
        Serializer.Init .InputStream
    
        Serializer.StartEnvelope
    
        Serializer.StartHeader
            Serializer.StartHeaderElement "ClientRequestCompression", Namespace
                Serializer.StartElement "sClientRequestCompression", Namespace
                Serializer.WriteString LCase(AllowCompressionForThisCall)
                Serializer.EndElement
            Serializer.EndHeaderElement
        Serializer.EndHeader
    
        Serializer.StartBody
        Serializer.StartElement Method, Namespace, , "m"
        
        If Not Parameters Is Nothing Then
            For Each Paramater In Parameters
                If (Paramater.ParamDirection And SOAP_PARAM_DIRECTION__INPUT) Then
                    If Paramater.ParamDataType = SOAP_PARAM_DATATYPE__STRINGATTACHMENT Then
                        Set oStringAttachment = New StringAttachment30
                        oStringAttachment.String = Paramater.ParamValue_Convert
                        oStringAttachment.Property("DimeTNF") = "media-type"
                        oStringAttachment.Property("DimeType") = "text/xml, charset=utf-8"
                        Serializer.AddAttachmentAndReference oStringAttachment
                    Else
                        Serializer.StartElement CStr(Paramater.ParamName), Namespace
                        Serializer.WriteString Paramater.ParamValue_Convert             'oParam.ParamValue
                        Serializer.EndElement
                    End If
                End If
            Next Paramater
        End If
    
        Serializer.EndElement
        Serializer.EndBody
    
        Debug.Print Timer & vbTab & "Begin soap call: " & Method
        Serializer.EndEnvelope
        Debug.Print Timer & vbTab & "End soap call: " & Method
    
        .EndMessage
        
        Set Reader = New SoapReader30
        Reader.Load .OutputStream
    
        If Not Reader.Fault Is Nothing Then
            Err.Raise LGSWEBSERVER_ERRORCODE_GENERAL, Method, Reader.FaultString.Text
        Else
            If Not Reader.RpcResult Is Nothing Then
                
                Execute_Internal = Reader.RpcResult.Text
                
                'get by ref values
                If Not Paramater Is Nothing Then
                    For Each Paramater In Parameters
                        If (Paramater.ParamDirection And SOAP_PARAM_DIRECTION__OUTPUT) Then
                            If Not Reader.RpcParameter(Paramater.ParamName, Namespace) Is Nothing Then
                                Paramater.ParamValue = Reader.RpcParameter(Paramater.ParamName, Namespace).nodeTypedValue
                            End If
                        End If
                    Next Paramater
                End If
            Else
                'no result
            End If
        End If
        
    End With
    Exit Function
    
Error_Handler:
    Err.Raise vbObjectError, "Execute_Internal", "Error occurred executing soap method [" & Err.Description & "]"
End Function
    
Private Sub Timings(sMethod As String, sEvent As String)
    If DEBUGMODE Then Debug.Print Timer & "   " & vbTab & sMethod & " - " & sEvent
End Sub

Private Function DisplayErrorMessage(Optional lErrorNumber As String, Optional sMessageDetail As String) As VbMsgBoxResult
    Dim sMessage As String
            
    sMessage = DEFAULT_ERRORMESSAGE_CONNECTIVITY
    If Len(lErrorNumber) > 0 Then
        sMessage = sMessage & " [" & lErrorNumber & "]"
    End If
    
    DisplayErrorMessage = PromptReconnect(sMessage, sMessageDetail)
End Function

Private Function ErrorHandler(ByVal sError As String, Optional ByVal lErrorNumber As Long = LGSWEBSERVER_ERRORCODE_GENERAL, Optional ByVal sMethod As String = "", Optional bRethrowError As Boolean = True, Optional bShowMessageBox As Boolean = MESSAGEBOX_ERROR, Optional bDisplayMethod As Boolean = False) As Boolean
    If bDisplayMethod Then
        If sMethod <> "" Then
            sError = "Error occurred in " & sMethod & " [" & sError & "]"
        Else
            sError = "Error occurred [" & sError & "]"
        End If
    End If
        
    If bShowMessageBox Then MsgBox sError, vbExclamation + vbOKOnly
    If bRethrowError Then Err.Raise lErrorNumber, sMethod, sError
End Function

Public Function PromptReconnect(Message As String, Optional MessageDetail As String = vbNullString) As VbMsgBoxResult
Dim Reconnect As ReconnectDialog

  Set Reconnect = New ReconnectDialog
    
  Screen.MousePointer = MousePointerConstants.vbDefault
    
  With Reconnect
    .Message = Message
    .MessageDetail = MessageDetail
    .Show vbModal
    
    PromptReconnect = .Action
  End With
  
  Unload Reconnect
  Set Reconnect = Nothing
  
End Function

