VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"WorkStatus"
Attribute VB_Ext_KEY = "Member0" ,"WorkStatus"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


Private mobjdas     As ICHCSII_DAS
Private oDatabase   As CHCSII_CONN.Conn
Private msName      As String
Private msComment   As String
Private msResult    As String
Private msPrvComment As String
Private msStatus    As Inpatient_Utilities.enumItemStatus
Private msAssessment As Inpatient_Utilities.enumItemCritical
Private msOrderId    As String
Private msOrderType  As Integer
Private msSign       As Integer
Private msNurse_Ack  As Integer
Private msDataId     As String
Private msAppendEnc   As Integer
Private msOrderBy     As String
Private msResultedBy     As String

Public Property Let orderId(ByVal vData As String)
    msOrderId = vData
End Property

Public Property Get orderId() As String
    orderId = msOrderId
End Property
Public Property Let dataId(ByVal vData As String)
    msDataId = vData
End Property

Public Property Get dataId() As String
    dataId = msDataId
End Property

Public Property Let OrderType(ByVal vData As Integer)
    msOrderType = vData
End Property

Public Property Get OrderType() As Integer
    OrderType = msOrderType
End Property
Public Property Let oName(ByVal vData As String)
    msName = vData
End Property

Public Property Get oName() As String
    oName = msName
End Property

Public Property Let oComment(ByVal vData As String)
    msComment = vData
End Property

Public Property Get oComment() As String
    oComment = msComment
End Property

Public Property Let Resutls(ByVal vData As String)
    msResult = vData
End Property

Public Property Get Resutls() As String
    Resutls = msResult
End Property

Public Property Let pComment(ByVal vData As String)
    msPrvComment = vData
End Property

Public Property Get pComment() As String
    pComment = msPrvComment
End Property

Public Property Let oStatus(ByVal vData As enumItemStatus)
    msStatus = vData
End Property

Public Property Get oStatus() As enumItemStatus
    oStatus = msStatus
End Property

Public Property Let oAssessment(ByVal vData As enumItemCritical)
    msAssessment = vData
End Property

Public Property Get oAssessment() As enumItemCritical
    oAssessment = msAssessment
End Property

Public Property Let Sign(ByVal vData As Integer)
    msSign = vData
End Property

Public Property Get Sign() As Integer
    Sign = msSign
End Property

Public Property Let Nurse_Ack(ByVal vData As Integer)
    msNurse_Ack = vData
End Property

Public Property Get Nurse_Ack() As Integer
    Nurse_Ack = msNurse_Ack
End Property

Public Property Let appendEnc(ByVal vData As Integer)
    msAppendEnc = vData
End Property

Public Property Get appendEnc() As Integer
    appendEnc = msAppendEnc
End Property

Public Property Let OrderBy(ByVal vData As String)
    msOrderBy = vData
End Property

Public Property Get OrderBy() As String
    OrderBy = msOrderBy
End Property

Public Property Let ResultedBy(ByVal vData As String)
    msResultedBy = vData
End Property

Public Property Get ResultedBy() As String
    ResultedBy = msResultedBy
End Property


Public Sub saveResultsforAction(ByVal ResultsChanged As Boolean)
                              
    Dim strSQL  As String
    Dim strSQL2 As String
    Dim inptSQL As String
    Dim AckName As String
    
    On Error GoTo ErrHnd
        
    If oStatus > 1 Then
        strSQL2 = ", COMPLETED_DATE = " & oDatabase.SQLDate(Now)
    End If
    
    If IsInpatient Then
    'SCR 104523 - LCI - 07/05/2007
    AckName = GetLogonFromNCID(mobjShared.CurrentUserNCID)
        inptSQL = ", SIGN_FLAG = " & Sign & ", NURSE_ACK_FALG = " & Nurse_Ack & ", NURSE_ACK_DATE = " & oDatabase.FormatDateForDBInsert(Now) & ", NURSE_ACK_NAME = " & oDatabase.SQLQuote(AckName)
    End If
    
    If OrderType = hlRAD_ORDERS_CODE Then
        strSQL = "update oe_rad set RESULTS = " & oDatabase.SQLQuote(Resutls, 4000)
        strSQL = strSQL & ", STATUS = " & Me.oStatus
        strSQL = strSQL & ", NORMAL_STATUS = " & Me.oAssessment
        ' Quan 85937
        strSQL = strSQL & ", RESULTS_COMMENTS = " & oDatabase.SQLQuote(pComment, 256)
        strSQL = strSQL & ", APPEND_ENC = " & appendEnc
        ' SCR 78761, JQ, 6-15-2006 -- only save resultedby if results have changed
        If ResultsChanged Then
            strSQL = strSQL & ", ResultedBy  = " & ResultedBy
        End If
        strSQL = strSQL & strSQL2 & inptSQL
        strSQL = strSQL & " where OE_RAD_DATA_ID = " & CLng(dataId)
    Else
        strSQL = "update oe_lab set RESULTS = " & oDatabase.SQLQuote(Resutls, 4000)
        strSQL = strSQL & ", STATUS = " & oStatus
        strSQL = strSQL & ", NORMAL_STATUS = " & oAssessment
        strSQL = strSQL & ", RESULTS_COMMENTS = " & oDatabase.SQLQuote(pComment, 256)
        strSQL = strSQL & ", APPEND_ENC = " & appendEnc
        ' SCR 78761, JQ, 6-15-2006 -- only save resultedby if results have changed
        If ResultsChanged Then
            strSQL = strSQL & ", ResultedBy  = " & ResultedBy
        End If
        strSQL = strSQL & strSQL2 & inptSQL
        strSQL = strSQL & " where OE_LAB_DATA_ID = " & CLng(dataId)
    End If
   
    mobjdas.ExecuteSQL strSQL
       
    If oStatus = isCompleted_Addr Then
        Call CheckCompletedOrders(CLng(dataId), OrderType)
    End If
    
    ' Quan 82065
    If oStatus = isCompleted_Unaddr Then
        Call CheckCompletedOrders(CLng(dataId), OrderType, True)
    End If
    
    Exit Sub
    
ErrHnd:
     mobjShared.ShowVBError Err.Number, Err.Description, "clsResults.saveResultsforAction", "clsResults", vbExclamation
     Exit Sub
     Resume
End Sub

Public Function retrieveData() As Boolean

Dim strSQL      As String
Dim i           As Integer
Dim Status      As enumItemStatus
Dim rs     As Recordset

On Error GoTo ErrHnd
    
    If OrderType = hlRAD_ORDERS_CODE Then
        strSQL = "select ORDER_ID, ORDER_NAME, STATUS, NORMAL_STATUS, RESULTS, RESULTS_COMMENTS, Append_Enc, ORDERED_BY,"
        If IsInpatient Then
            strSQL = strSQL & " SIGN_FLAG, NURSE_ACK_FALG, "
        End If
        strSQL = strSQL & " CLINICAL_IMPRESSION_TEXT, RAD_COMMENT as prv_comment  "
        strSQL = strSQL & " from OE_RAD "
        strSQL = strSQL & " where OE_RAD_DATA_ID = " & CLng(dataId)
        
        Set rs = mobjdas.OpenRecordset(strSQL)

        If Not rs.EOF Then
            oComment = rs!RESULTS_COMMENTS & vbNullString
            pComment = rs!CLINICAL_IMPRESSION_TEXT & vbCrLf & rs!prv_comment & vbNullString
            oName = rs!ORDER_NAME & vbNullString
            Resutls = rs!RESULTS & vbNullString
            orderId = rs!ORDER_ID
            appendEnc = Val(rs!Append_Enc & vbNullString)
            If IsInpatient Then
                Me.Sign = Val(rs!SIGN_FLAG & vbNullString)
                Me.Nurse_Ack = Val(rs!NURSE_ACK_FALG & vbNullString)
            End If
        End If
    Else
        strSQL = "select ORDER_ID, ORDER_IEN, ORDER_NAME, STATUS, NORMAL_STATUS,Append_Enc,ORDERED_BY,"
        strSQL = strSQL & " RESULTS, RESULTS_COMMENTS,"
        If IsInpatient Then
            strSQL = strSQL & " SIGN_FLAG, NURSE_ACK_FALG,  "
        End If
        strSQL = strSQL & " LAB_COMMENT as prv_comment "
        strSQL = strSQL & " from OE_LAB "
        strSQL = strSQL & " where OE_LAB_DATA_ID = " & CLng(dataId)
        
        Set rs = mobjdas.OpenRecordset(strSQL)
        
        If Not rs.EOF Then
            oComment = rs!RESULTS_COMMENTS & vbNullString
            pComment = rs!prv_comment & vbNullString
            oName = rs!ORDER_NAME & vbNullString
            Resutls = rs!RESULTS & vbNullString
            orderId = rs!ORDER_ID
            appendEnc = Val(rs!Append_Enc & vbNullString)
            If IsInpatient Then
                Me.Sign = Val(rs!SIGN_FLAG & vbNullString)
                Me.Nurse_Ack = Val(rs!NURSE_ACK_FALG & vbNullString)
            End If
            
        End If
                        
        If Trim$(Resutls) = vbNullString Then
            Resutls = getLabValueRange(rs!ORDER_IEN)
        End If
        
    End If

    OrderBy = rs!ORDERED_BY
    
    i = Val(rs!Status & vbNullString)
    oStatus = i

    If rs!NORMAL_STATUS = vbNullString Then
        i = 0
    Else
        i = Val(rs!NORMAL_STATUS & vbNullString)
    End If
    oAssessment = i
    
    orgResult = Resutls
    orgComment = oComment
    orgStatus = oStatus
    orgCritical = oAssessment
   
    rs.Close
    retrieveData = True
    Exit Function
    
ErrHnd:
     mobjShared.ShowVBError Err.Number, Err.Description, "clsResults.retrieveData", "clsResults", vbExclamation
    retrieveData = False
End Function

Private Sub Class_Initialize()

    Set oDatabase = New CHCSII_CONN.Conn
    Set mobjdas = oDatabase.CHCSII_DAS

End Sub

Private Function getLabValueRange(ByVal ien As Long) As String
Dim rs     As Recordset
Dim strSQL As String
On Error GoTo ErrHnd
    
    getLabValueRange = vbNullString
    
    strSQL = "select VALUE_RANGE from LAB_TEST_OE where LAB_TEST_IEN = " & ien
    Set rs = mobjdas.OpenRecordset(strSQL)
    If Not rs.EOF Then
        getLabValueRange = rs!VALUE_RANGE & vbNullString
    End If
    
    rs.Close
    Exit Function
    
ErrHnd:
    mobjShared.ShowVBError Err.Number, Err.Description, "clsResults.getLabValueRange", "clsResults", vbExclamation
End Function

Private Sub Class_Terminate()

    Set mobjdas = Nothing
    Set oDatabase = Nothing

End Sub


Private Sub CheckCompletedOrders(ByVal dataId As Long, ByVal OrderType As Integer, Optional ByVal bUnaddr As Boolean)
Dim sSql     As String
Dim sSQL2    As String
Dim oRs      As ADODB.Recordset
Dim oRs2    As ADODB.Recordset
Dim oRs3    As ADODB.Recordset
Dim orderId  As String
Dim tableName As String
Dim oCount  As Integer
Dim mCount  As Integer
    
On Error GoTo ErrHandler

    If OrderType = hlLAB_ORDERS_CODE Then
        tableName = "OE_LAB"
        sSql = "Select ORDER_ID From OE_LAB WHERE OE_LAB_DATA_ID = " & dataId & " And IS_ORDER_OCCURRENCE = 1 "
    Else
        tableName = "OE_RAD"
        sSql = "Select ORDER_ID From OE_RAD WHERE OE_RAD_DATA_ID = " & dataId & " And IS_ORDER_OCCURRENCE = 1 "
    End If
    
    Set oRs = mobjdas.OpenRecordset(sSql)
    
    If Not oRs.EOF Then
        orderId = oRs.Fields("ORDER_ID")
        'SCR-135135 - 7/13/07 - FL -- Rollback SCR-82065
            'Display "Completed/Unaddressed" in A&P only when all child (including future generated pending child) is
            'in "Completed/Unaddressed"
            'SA updating requirement tracker #60176
            Call SetFrequenceOrderToComplete(tableName, orderId)
    End If
    
    Exit Sub
    
ErrHandler:
    mobjShared.ShowVBError Err.Number, Err.Description, "clsResults.CheckCompletedOrders", "clsResults", vbExclamation
End Sub
Private Sub SetFrequenceOrderToComplete(ByVal tableName As String, ByVal orderId As String)
Dim oParam As ADODB.Parameter
Dim oCmd As ADODB.Command
Dim oDAS As IADO_DAS
Dim rs  As Recordset

On Error GoTo ErrHandler
    
    Set oDAS = oDatabase.ADO_DAS
    Set oCmd = New ADODB.Command
               
    With oCmd
    
        .CommandType = adCmdStoredProc
        .CommandText = "Check_CompletedStatus_Pkg.prc_OrderToComplete"
        
        Set oParam = .CreateParameter("p_vOrder_Id", adVarChar, adParamInput, 128, orderId)
        .Parameters.Append oParam
        Set oParam = .CreateParameter("p_vTable_Name ", adVarChar, adParamInput, 128, tableName)
        .Parameters.Append oParam
        
    End With
    Set rs = oDAS.OpenRecordsetSP(oCmd)

    Exit Sub
   
ErrHandler:
    mobjShared.ShowVBError Err.Number, Err.Description, "clsResults.SetFrequenceOrderToComplete", "clsResults", vbExclamation
End Sub

Private Function GetLogonFromNCID(ByVal sNCID As String) As String
    Dim rs     As Recordset
    Dim strSQL As String
    Dim UserName() As String
    
    On Error GoTo ErrHnd
    
    strSQL = "select name from provider where ncid = " & sNCID
    Set rs = mobjdas.OpenRecordset(strSQL)
    If Not rs.EOF Then
        GetLogonFromNCID = rs!Name & vbNullString
    End If
    rs.Close
    'Make sure there is a space after comma between last and first name
    UserName = Split(GetLogonFromNCID, ",")
    GetLogonFromNCID = UserName(0) & ", " & UserName(1)
    Exit Function
    
ErrHnd:
    mobjShared.ShowVBError Err.Number, Err.Description, "clsResults - GetLogonFromNCID", "Order Results", vbExclamation
End Function



