VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MapperTheater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private objShared As CWShared
Implements DL_Interfaces.iMapperFillMed

Private mConn As CHCSII_CONN.Conn
Private SQLOps As ICHCSII_DAS
Private mMapperFactory As DL_Interfaces.iDataMapperFactory
Private mCoreMapper    As DL_Interfaces.iMapperCore

Private Sub iMapperFillMed_Init(Support As DL_Support.iParameters, Cache As DL_Support.CacheMapper)
On Error GoTo ErrHandler

    Set mConn = New CHCSII_CONN.Conn
    Set SQLOps = mConn.CHCSII_DAS(Auto)
    Set objShared = New CWShared

    Set mMapperFactory = Support.Item("DataMapperFactory").Value
    
    Exit Sub

ErrHandler:
  Err.Raise Err.Number, CACHE_NAME & ".MapperTheater.iMapperFillMed_Init", Err.Description, Err.HelpFile, Err.HelpContext

End Sub

Private Function iMapperFillMed_Retrieve(Params As DL_Support.iParameters) As DTOs.MedFills
    On Error GoTo ErrHandler
    Dim MyOrders  As DTOs.MedFills
    Dim param     As DM_FillMedHistory.Retrieve_Parameters
    
    Set MyOrders = New DTOs.MedFill
    Set param = Params

    Set MyOrders = GetFillMedHistory(param, MyOrders)

    Set iMapperFillMed_Retrieve = MyOrders

    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, CACHE_NAME & ".MapperTheater.iMapperFillMed_Retrieve", Err.Description, Err.HelpFile, Err.HelpContext
End Function

Private Function iMapperFillMed_Save(Meds As DTOs.MedFills) As Boolean

    On Error GoTo ErrHandler
    Dim nOrderCount As Integer
    Dim MyOrder As DTOs.MedFill
    
    For nOrderCount = 1 To Meds.Count
        Set MyOrder = Meds.Item(nOrderCount)
        iMapperFillMed_Save = SaveMedFillHistory(MyOrder)
    Next nOrderCount
    
    Exit Function

ErrHandler:
  Err.Raise Err.Number, CACHE_NAME & ".MapperTheater.iMapperFillMed_Save", Err.Description, Err.HelpFile, Err.HelpContext

End Function


Private Function GetFillMedHistory(ByVal oParams As DM_FillMedHistory.Retrieve_Parameters, _
                                  ByRef objFillMeds As DTOs.MedFills) As DTOs.MedFills
    Dim strSQL     As String
    Dim sWhere      As String
    Dim oRs         As ADODB.Recordset
    Dim oFillMed   As DTOs.MedFill
    
    On Error GoTo ErrHandler
   
    strSQL = "Select ORDER_IEN, ACTION_DATE, ACTION_BY, QUANTITY, SIG_CODE, ACTION_FLAG, COMMENTS, REFILLS, ORDER_NAME, UNIT_NUMBER, FACILITY_NCID "
    strSQL = strSQL & " From MED_FILL_HISTORY WHERE ORDER_ID =  " & oParams.OrderId
    strSQL = strSQL & " ORDER BY ACTION_DATE  "
    
    Set oRs = SQLOps.OpenRecordset(strSQL)
    If oRs.RecordCount > 0 Then
        oRs.MoveFirst
        Do Until oRs.EOF
            Set oFillMed = New DTOs.MedFill
            With oFillMed
                .Action = ConvertToActionString(oRs.Fields("ACTION_FLAG"))
                .FillDate = oRs.Fields("ACTION_DATE")
                .Name = oRs.Fields("ACTION_BY")
                .Quantity = oRs.Fields("QUANTITY")
                .Sig = oRs.Fields("SIG_CODE")
                .Comment = oRs.Fields("Comments")
                .OrderIEN = oRs.Fields("OrderIEN")
                .Refills = oRs.Fields("Refills")
                .OrderName = oRs.Fields("Order_Name")
                .UnitNumber = oRs.Fields("UNIT_NUMBER")
                .FacilityNCID = oRs.Fields("FACILITY_NCID")
            End With
            objFillMeds.Add oFillMed
            oRs.MoveNext
            Set oFillMed = Nothing
        Loop
    End If
    
    Set GetFillMedHistory = objFillMeds
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, CACHE_NAME & ".MapperTheater.GetFillMedHistory", Err.Description, Err.HelpFile, Err.HelpContext
End Function


Private Function GetFillMedHistoryByOrderIEN(ByVal oParams As DM_FillMedHistory.Retrieve_Parameters, _
                                  ByRef objFillMeds As DTOs.MedFills) As DTOs.MedFills
    Dim strSQL     As String
    Dim sWhere      As String
    Dim oRs         As ADODB.Recordset
    Dim oFillMed   As DTOs.MedFill
    
    On Error GoTo ErrHandler
   
    strSQL = "Select ORDER_IEN, ACTION_DATE, ACTION_BY, QUANTITY, SIG_CODE, ACTION_FLAG, COMMENTS, UNIT_NUMBER, FACILITY_NCID "
    strSQL = strSQL & " From MED_FILL_HISTORY WHERE ORDER_ID =  " & oParams.OrderIEN
    strSQL = strSQL & " ORDER BY ACTION_DATE  "
    
    Set oRs = SQLOps.OpenRecordset(strSQL)
    If oRs.RecordCount > 0 Then
        oRs.MoveFirst
        Do Until oRs.EOF
            Set oFillMed = New DTOs.MedFill
            With oFillMed
                .Action = ConvertToActionString(oRs.Fields("ACTION_FLAG"))
                .FillDate = oRs.Fields("ACTION_DATE")
                .Name = oRs.Fields("ACTION_BY")
                .Quantity = oRs.Fields("QUANTITY")
                .Sig = oRs.Fields("SIG_CODE")
                .Comment = oRs.Fields("Comments")
                .OrderIEN = oRs.Fields("OrderIEN")
                .OrderName = oRs.Fields("ORDER_NAME")
                .Refills = oRs.Fields("REFILLS")
                .UnitNumber = oRs.Fields("UNIT_NUMBER")
                .FacilityNCID = oRs.Fields("FACILITY_NCID")
            End With
            objFillMeds.Add oFillMed
            oRs.MoveNext
            Set oFillMed = Nothing
        Loop
    End If
    
    Set GetFillMedHistoryByOrderIEN = objFillMeds
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, CACHE_NAME & ".MapperTheater.GetFillMedHistoryByOrderIEN", Err.Description, Err.HelpFile, Err.HelpContext
End Function

Private Function SaveMedFillHistory(ByRef oOrder As DTOs.MedFill) As Boolean

    Dim strSQL  As String
    
    On Error GoTo ErrHandler
    
    '--- UPDATE OR_RX TABLE
    '--- SCR 79823  Sherry Wang  10/11/2005  Add UnitNumber and FacilityNCID
    With oOrder
        strSQL = "INSERT INTO MED_FILL_HISTORY ( " & vbNewLine
        strSQL = strSQL & " ORDER_ID, ACTION_ORDERID, ORDER_IEN, ACTION_DATE, ACTION_BY, QUANTITY, SIG_CODE, COMMENTS, ORDER_NAME, REFILLS, UNIT_NUMBER, FACILITY_NCID, ACTION_FLAG )" & vbNewLine
        strSQL = strSQL & " VALUES ( " & vbNewLine
        strSQL = strSQL & mConn.SQLQuote(.OrderId) & ","
        strSQL = strSQL & mConn.SQLQuote(.ActionOrderId) & ","
        strSQL = strSQL & mConn.SQLQuote(.OrderIEN) & ","
        strSQL = strSQL & mConn.FormatDateForDBInsert(.FillDate) & ","
        strSQL = strSQL & mConn.SQLQuote(.Name) & ","
        strSQL = strSQL & .Quantity & ","
        strSQL = strSQL & mConn.SQLQuote(.Sig) & ","
        strSQL = strSQL & mConn.SQLQuote(.Comment) & ","
        strSQL = strSQL & mConn.SQLQuote(.OrderName) & ","
        strSQL = strSQL & .Refills & ","
        strSQL = strSQL & .UnitNumber & ","
        strSQL = strSQL & .FacilityNCID & ","
        strSQL = strSQL & ConvertToActionFlag(.Action) & ")" & vbNewLine
    End With

    SQLOps.ExecuteSQL strSQL

    SaveMedFillHistory = True
    
    Exit Function
    
ErrHandler:
    SaveMedFillHistory = False
    Err.Raise Err.Number, CACHE_NAME & ".MapperTheater.SaveMedFillHistory", Err.Description, Err.HelpFile, Err.HelpContext
End Function


