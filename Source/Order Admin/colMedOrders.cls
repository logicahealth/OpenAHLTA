VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colMedOrders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mCol As Collection

Public Function loadData() As Boolean

    Dim sSql As String
    Dim rs As New ADODB.Recordset
    Dim objOrder As medOrder
    Dim PrevMedID As Long
    
    On Error GoTo ErrHandler
        
    If Len(goBB.SharedObj.CmdLineSwitches("INPATIENT2")) > 0 Then
      
        sSql = "Select A.MEDID, A.DESCDISPLAY, A.SYNONYMS, A.DEFAULT_SIG, A.DEFAULT_REFILL, A.DEFAULT_ROUTE"
        sSql = sSql & ", A.DOSAGE_FORM, A.DEFAULT_FREQUENCY, A.DEFAULT_UNIT, A.CURRENT_STOCK, A.DRUG_CHECK_DISABLED"
        sSql = sSql & ", A.NOTE_TO_PROVIDER, A.NDC, A.NSN, A.DEFAULT_DOSE, A.DOSAGE_STRENGTH, A.CONTENT_UNIT "
        sSql = sSql & " From RX_LIST_FDB A "
        
        sSql = sSql & " Order By DESCDISPLAY"
        
        Set rs = goBB.I_DAS.ExecuteSQL(sSql)

        Do Until rs.EOF
          Set objOrder = New medOrder
          
          If IsNull(rs.Fields("DRUG_CHECK_DISABLED")) Then
             objOrder.Active = True
          Else
             If Val(rs.Fields("DRUG_CHECK_DISABLED")) = 1 Then
                 objOrder.Active = False
             Else
                 objOrder.Active = True
             End If
          End If
          
          With objOrder
            .Ien = rs("MEDID")
            .DrugName = CStr("" & rs("DESCDISPLAY"))
            .Synonyms = CStr("" & rs.Fields("SYNONYMS"))
            .DefSig = CStr("" & rs.Fields("DEFAULT_SIG"))
            .DefRefill = IsNullValue(rs.Fields("DEFAULT_REFILL"), 0)
            .DefRoute = CStr("" & rs.Fields("DEFAULT_ROUTE"))
            .DefForm = CStr("" & rs.Fields("DOSAGE_FORM"))
            .DefFreq = CStr("" & rs.Fields("DEFAULT_FREQUENCY"))
            .DefQty = IsNullValue(rs.Fields("DEFAULT_UNIT"), 0)
            .CurrentStock = IsNullValue(rs.Fields("CURRENT_STOCK"), 0)
                       
            .NDC = CStr("" & rs.Fields("NDC"))
            .NSN = CStr("" & rs.Fields("NSN"))
            .DefDose = CStr("" & rs.Fields("DEFAULT_DOSE"))
            .DosageStrength = CStr("" & rs.Fields("DOSAGE_STRENGTH"))
            .ContentUnit = CStr("" & rs.Fields("CONTENT_UNIT"))
            
'            If Len(.NSN) > 0 Then Stop
            
            .NoteToProvider = CStr("" & rs.Fields("NOTE_TO_PROVIDER"))
          End With
          mCol.Add objOrder, objOrder.Ien & objOrder.NDC
          rs.MoveNext
      Loop
    Else
      sSql = "select * "
      sSql = sSql & " from RX_LIST "
      sSql = sSql & " where (NDC_INDICATOR is NULL or NDC_INDICATOR = 0)"
      sSql = sSql & " and drug_name is not null"
      sSql = sSql & " and drug_ien <> 0"
      sSql = sSql & " and FACILITY_NCID = " & goBB.Logon.FacilityNCID
      sSql = sSql & " Order by drug_name "
      
      Set rs = goBB.I_DAS.ExecuteSQL(sSql)
      
      Do Until rs.EOF
          Set objOrder = New medOrder
           If IsNull(rs.Fields("DRUG_CHECK_DISABLED")) Then
              objOrder.Active = True
           Else
              If Val(rs.Fields("DRUG_CHECK_DISABLED")) = 1 Then
                  objOrder.Active = False
              Else
                  objOrder.Active = True
              End If
           End If
          objOrder.Ien = IsNullValue(rs.Fields("DRUG_IEN"), 0)
          objOrder.DrugName = CStr("" & rs.Fields("DRUG_NAME"))
          objOrder.Synonyms = CStr("" & rs.Fields("SYNONYMS"))
          objOrder.DefSig = CStr("" & rs.Fields("DEFAULT_SIG"))
          objOrder.DefRefill = IsNullValue(rs.Fields("DEFAULT_REFILL"), 0)
          objOrder.DefRoute = CStr("" & rs.Fields("DEFAULT_ROUTE"))
          objOrder.DefForm = CStr("" & rs.Fields("DOSAGE_FORM"))
          objOrder.DefFreq = CStr("" & rs.Fields("CONTENT_UNIT"))
          objOrder.DefQty = IsNullValue(rs.Fields("DEFAULT_UNIT"), 0)
          objOrder.CurrentStock = IsNullValue(rs.Fields("CURRENT_STOCK"), 0)
          
          mCol.Add objOrder
          Set objOrder = Nothing
          rs.MoveNext
      Loop
    End If
    Set rs = Nothing
    loadData = True
     
Exit Function

ErrHandler:
    Call goBB.SharedObj.ShowVBError(Err.Number, Err.Description, "colMedOrder.loadData()", "ReasonForVisit", vbExclamation)
    loadData = False
    Exit Function
    Resume
End Function

Public Function AddMedOrder(ByVal oMedOrder As medOrder) As Boolean
    
    Dim SQLStmt As String
    Dim DataID As String
    Dim ActiveFlag As Integer
    
    On Error GoTo ErrHandler

    AddMedOrder = False
    DataID = goBB.I_DAS.GetNextID()
    
    With oMedOrder
        'INPATIENT2 12/20/04 Frank
        'Use FDB Drug Interaction checking mechanism
        If Len(goBB.SharedObj.CmdLineSwitches("INPATIENT2")) > 0 Then

            If .Active = True Then
                ActiveFlag = 0
            Else
                ActiveFlag = 1
            End If

            SQLStmt = "INSERT INTO RX_LIST_FDB (MEDID, DRUG_LIST, DESCDISPLAY, " _
                    & "DEFAULT_UNIT, NDC_INDICATOR, " _
                    & "DRUG_CHECK_DISABLED, " _
                    & "DOSAGE_STRENGTH, DEFAULT_FREQUENCY, DOSAGE_FORM, DEFAULT_DOSE, CONTENT_UNIT, Synonyms, " _
                    & "DEFAULT_SIG, DEFAULT_REFILL, DEFAULT_ROUTE, CURRENT_STOCK, " _
                    & "NSN, NDC, NOTE_TO_PROVIDER) " _
                    & "VALUES (" & .Ien & "," & _
                                "1," & _
                                goBB.DataConn.SQLQuote(.DrugName) & "," & _
                                goBB.DataConn.SQLQuote(.DefQty) & "," & _
                                "0," & _
                                ActiveFlag & "," & _
                                goBB.DataConn.SQLQuote(.DosageStrength) & "," & _
                                goBB.DataConn.SQLQuote(.DefFreq) & "," & _
                                goBB.DataConn.SQLQuote(.DefForm) & "," & _
                                goBB.DataConn.SQLQuote(.DefDose) & "," & _
                                goBB.DataConn.SQLQuote(.ContentUnit) & "," & _
                                goBB.DataConn.SQLQuote(.Synonyms) & "," & _
                                goBB.DataConn.SQLQuote(.DefSig) & "," & _
                                .DefRefill & "," & _
                                goBB.DataConn.SQLQuote(.DefRoute) & "," & _
                                .CurrentStock & "," & _
                                goBB.DataConn.SQLQuote(.NSN) & "," & _
                                goBB.DataConn.SQLQuote(.NDC) & "," & _
                                goBB.DataConn.SQLQuote(.NoteToProvider) & ")"
        Else
            SQLStmt = "INSERT INTO RX_LIST (DRUG_IEN, DRUG_NAME, SYNONYMS, " & _
                "DEFAULT_SIG, DEFAULT_REFILL, DEFAULT_ROUTE, DOSAGE_FORM, " & _
                "CONTENT_UNIT, DEFAULT_UNIT, CURRENT_STOCK, FACILITY_NCID ) " & _
                "VALUES (" & DataID & "," & _
                goBB.DataConn.SQLQuote(UCase$(.DrugName), DrugName_LEN) & "," & _
                goBB.DataConn.SQLQuote(UCase$(.Synonyms), Synonyms_LEN) & "," & _
                goBB.DataConn.SQLQuote(.DefSig, DefSig_LEN) & "," & _
                Val(.DefRefill) & ", " & _
                goBB.DataConn.SQLQuote(.DefRoute, DefRoute_LEN) & "," & _
                goBB.DataConn.SQLQuote(.DefDose, DefDose_LEN) & "," & _
                goBB.DataConn.SQLQuote(.DefFreq, DefFreq_LEN) & "," & _
                Val(.DefQty) & "," & _
                .CurrentStock & "," & _
                goBB.Logon.FacilityNCID & ")"
        End If
    End With
    
    goBB.I_DAS.ExecuteSQL SQLStmt
    AddMedOrder = True
    Exit Function

ErrHandler:
'    Call goBB.SharedObj.ShowVBError(Err.Number, Err.Description, "colMedOrders.AddMedOrder()", "OrderAdmin", vbExclamation)
    Exit Function
    Resume
End Function
Public Function saveData() As Boolean

    Dim sSql As String
    Dim rs As New ADODB.Recordset
    Dim objOrder As medOrder
    
    On Error GoTo ErrHandler
        
    sSql = "select DRUG_IEN, DRUG_NAME, DEFAULT_UNIT, DOSAGE_STRENGTH,"
    sSql = sSql & " CONTENT_UNIT, DOSAGE_FORM"
    sSql = sSql & " from RX_LIST "
    sSql = sSql & " where (DRUG_CHECK_DISABLED is NULL or DRUG_CHECK_DISABLED = 0)"
    sSql = sSql & " and FACILITY_NCID = " & goBB.Logon.FacilityNCID
    If goBB.SharedObj.UseMySQLDB Then
        sSql = sSql & " LIMIT 10"
    Else
        sSql = sSql & " and rownum <= 10"
    End If
    
    Set rs = goBB.I_DAS.ExecuteSQL(sSql)
    
    Do Until rs.EOF
        Set objOrder = New medOrder
        objOrder.Ien = IsNullValue(rs.Fields("DRUG_IEN"), 0)
        objOrder.DrugName = CStr("" & rs.Fields("DRUG_NAME"))
        objOrder.Synonyms = CStr("" & rs.Fields("DRUG_NAME"))
        objOrder.DefSig = ""
        objOrder.DefRefill = CStr("" & rs.Fields("DEFAULT_UNIT"))
        objOrder.DefRoute = CStr("" & rs.Fields("DOSAGE_FORM"))
        objOrder.DefFreq = CStr("" & rs.Fields("CONTENT_UNIT"))
        objOrder.DefQty = IsNullValue(rs.Fields("DOSAGE_STRENGTH"), 0)
        objOrder.CurrentStock = 0
        objOrder.Active = True
        mCol.Add objOrder
        Set objOrder = Nothing
        rs.MoveNext
    Loop

    Set rs = Nothing
    saveData = True
     
Exit Function

ErrHandler:
    Call goBB.SharedObj.ShowVBError(Err.Number, Err.Description, "colMedOrder.saveData()", "ReasonForVisit", vbExclamation)
    saveData = False
    
End Function
Public Function DeleteOrder(myIen As String, NDC As String) As Boolean

    Dim SQLStmt As String
    
    On Error GoTo ErrHandler
    
    DeleteOrder = False
    
    If Len(goBB.SharedObj.CmdLineSwitches("INPATIENT2")) > 0 Then
        'SCR 12/20/04 Frank
        'Use FDB Drug Interaction checking mechanism
        
        SQLStmt = "delete from RX_LIST_FDB"
        SQLStmt = SQLStmt & " where MEDID = " & myIen
        If NDC = "" Then
            SQLStmt = SQLStmt & " and NDC is NULL"
        Else
            SQLStmt = SQLStmt & " and upper(NDC) = " & UCase$(goBB.DataConn.SQLQuote(NDC))
        End If
    Else
        SQLStmt = "update RX_LIST "
        SQLStmt = SQLStmt & " set NDC_INDICATOR = 1 "
        SQLStmt = SQLStmt & " where DRUG_IEN = " & myIen
        SQLStmt = SQLStmt & " and FACILITY_NCID = " & goBB.Logon.FacilityNCID
    End If
    
    goBB.I_DAS.ExecuteSQL SQLStmt
    DeleteOrder = True
    Exit Function

ErrHandler:
'    Call goBB.SharedObj.ShowVBError(Err.Number, Err.Description, "colMedOrders.DeleteOrder()", "OrderAdmin", vbExclamation)

End Function

Public Function DeleteAllMedOrder() As Boolean
    
    'SCR 12/20/04 Frank
    'Use FDB Drug Interaction checking mechanism
    
    '*** NEW SUB for INPATIENT2
    
    Dim SQLStmt As String
    
    On Error GoTo ErrHandler
    
    DeleteAllMedOrder = False
    If Len(goBB.SharedObj.CmdLineSwitches("INPATIENT2")) > 0 Then
        
        SQLStmt = "delete from RX_LIST_FDB "
'        SQLStmt = "truncate table " & sTableName   '*** Truncate does not work from here
                                                    '*** But works in Toad!!

        goBB.I_DAS.ExecuteSQL SQLStmt
    End If
    DeleteAllMedOrder = True
    Exit Function

ErrHandler:
    On Error Resume Next
'    Call goBB.SharedObj.ShowVBError(Err.Number, Err.Description, "colMedOrders.DeleteAllMedOrder()", "OrderAdmin", vbExclamation)

End Function

Public Sub Add(orderItem As Object, Optional sKey As String)
    'create a new object
    Dim objNewMember As medOrder
    Set objNewMember = orderItem

    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If
    
    Set objNewMember = Nothing

End Sub

Public Property Get Item(vntIndexKey As Variant) As medOrder
    Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
        Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
        mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
        Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub


Public Function IsMedOrderExist(MedID As Long) As Boolean

    Dim SQLStmt As String
    Dim oRS As ADODB.Recordset

    On Error GoTo ErrHandler
    
    IsMedOrderExist = True = False
    
    SQLStmt = "select *  from RX_LIST_FDB" _
            & " where MEDID = " & MedID
    
    Set oRS = goBB.I_DAS.OpenRecordset(SQLStmt, adOpenStatic)
        
    IsMedOrderExist = Not (oRS.BOF And oRS.EOF)
    
    Set oRS = Nothing
    Exit Function

ErrHandler:
    Call goBB.SharedObj.ShowVBError(Err.Number, Err.Description, "colMedOrders.IsMedOrderExist()", "OrderAdmin", vbExclamation)

End Function


Public Function UpdateMedOrder(ByVal oMedOrder As medOrder) As Boolean

    Dim SQLStmt As String
    Dim inactiveFlag As Integer
    
    On Error GoTo ErrHandler

    UpdateMedOrder = False
    If oMedOrder.Active = False Then
        inactiveFlag = 1
    Else
        inactiveFlag = 0
    End If
    
    With oMedOrder
        If Len(goBB.SharedObj.CmdLineSwitches("INPATIENT2")) > 0 Then
            SQLStmt = "UPDATE RX_LIST_FDB SET " & _
                    " SYNONYMS = " & goBB.DataConn.SQLQuote(UCase$(.Synonyms), Synonyms_LEN) & _
                    ", DEFAULT_SIG = " & goBB.DataConn.SQLQuote(.DefSig, DefSig_LEN) & _
                    ", DEFAULT_REFILL = " & Val(.DefRefill) & _
                    ", DEFAULT_ROUTE = " & goBB.DataConn.SQLQuote(.DefRoute, DefRoute_LEN) & _
                    ", DOSAGE_FORM = " & goBB.DataConn.SQLQuote(.DefForm) & _
                    ", DEFAULT_DOSE = " & goBB.DataConn.SQLQuote(.DefDose) & _
                    ", DEFAULT_FREQUENCY = " & goBB.DataConn.SQLQuote(.DefFreq, DefFreq_LEN) & _
                    ", DEFAULT_UNIT = " & Val(.DefQty) & _
                    ", DOSAGE_STRENGTH = " & goBB.DataConn.SQLQuote(.DosageStrength) & _
                    ", CONTENT_UNIT = " & goBB.DataConn.SQLQuote(.ContentUnit) & _
                    ", CURRENT_STOCK =  " & .CurrentStock & _
                    ", NOTE_TO_PROVIDER =  " & goBB.DataConn.SQLQuote(.NoteToProvider) & _
                    ", NDC = " & goBB.DataConn.SQLQuote(.NDC) & _
                    ", NSN = " & goBB.DataConn.SQLQuote(.NSN) & _
                    ", DRUG_CHECK_DISABLED = " & inactiveFlag & _
                    " WHERE MEDID = " & .Ien
        Else
            SQLStmt = "UPDATE  RX_LIST SET " & _
                    " SYNONYMS = " & goBB.DataConn.SQLQuote(UCase$(.Synonyms), Synonyms_LEN) & _
                    ", DEFAULT_SIG = " & goBB.DataConn.SQLQuote(.DefSig, DefSig_LEN) & _
                    ", DEFAULT_REFILL = " & Val(.DefRefill) & _
                    ", DEFAULT_ROUTE = " & goBB.DataConn.SQLQuote(.DefRoute, DefRoute_LEN) & _
                    ", DOSAGE_FORM = '" & Val(.DefDose) & _
                    "', CONTENT_UNIT = " & goBB.DataConn.SQLQuote(.DefFreq, DefFreq_LEN) & _
                    ", DEFAULT_UNIT = " & Val(.DefQty) & _
                    ", CURRENT_STOCK =  " & .CurrentStock & _
                    ", DRUG_CHECK_DISABLED = " & inactiveFlag & _
                    " WHERE DRUG_IEN = " & .Ien
            SQLStmt = SQLStmt & " and FACILITY_NCID = " & goBB.Logon.FacilityNCID
        End If
    End With

    goBB.I_DAS.ExecuteSQL SQLStmt
    UpdateMedOrder = True
    Exit Function

ErrHandler:
    On Error Resume Next
'    Call goBB.SharedObj.ShowVBError(Err.Number, Err.Description, "colMedOrders.UpdateMedOrder()", "OrderAdmin", vbExclamation)
End Function
