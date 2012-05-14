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
Attribute VB_Name = "colRadOrders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mCol As Collection

Public Function loadData() As Boolean

    Dim sSql As String
    Dim rs As New ADODB.Recordset
    Dim objOrder As radOrder
    
    On Error GoTo ErrHandler
        
     sSql = "select INACTIVE_FLAG, RAD_TEST_IEN, RAD_TEST_NAME, SYNONYMS, RAD_TEST_TYPE "
     sSql = sSql & " from RAD_TEST "
     sSql = sSql & " where " '(INACTIVE_FLAG is NULL or INACTIVE_FLAG = 0)"
     sSql = sSql & "  rad_test.rad_test_name is not null"
     sSql = sSql & " AND rad_test.Rad_test_ien <> 0"
     sSql = sSql & " and FACILITY_NCID = " & goBB.Logon.FacilityNCID
'      sSql = sSql & " and rad_test_name like 'JANE%' "
    
    '** SCR 68553 - 4/6/05 - FRANK
    '** New column DELETE_FLAG = 1 when Lab is DELETE from Order Admin
    sSql = sSql & " and DELETE_FLAG <> 1"
    
     sSql = sSql & " Order by rad_test_name "
    
     Set rs = goBB.I_DAS.ExecuteSQL(sSql)
     
     Do Until rs.EOF
         Set objOrder = New radOrder
         If IsNull(rs.Fields("INACTIVE_FLAG")) Then
            objOrder.Active = True
         Else
            If Val(rs.Fields("INACTIVE_FLAG")) = 1 Then
                objOrder.Active = False
            Else
                objOrder.Active = True
            End If
         End If
         objOrder.Ien = IsNullValue(rs.Fields("RAD_TEST_IEN"), 0)
         objOrder.TextName = CStr("" & "" & rs.Fields("RAD_TEST_NAME"))
         objOrder.Synonyms = CStr("" & "" & rs.Fields("SYNONYMS"))
         objOrder.TestType = CStr("" & "" & rs.Fields("RAD_TEST_TYPE"))
         mCol.Add objOrder
         Set objOrder = Nothing
         rs.MoveNext
     Loop
     
     
     Set rs = Nothing
     loadData = True
     
Exit Function

ErrHandler:
    Call goBB.SharedObj.ShowVBError(Err.Number, Err.Description, "colRadOrders.loadData()", "ReasonForVisit", vbExclamation)
    loadData = False
    
End Function

Public Sub AddRadOrder(ByVal oRadOrder As radOrder)

Dim SQLStmt As String
Dim DataID As String

    On Error GoTo ErrHandler
    DataID = goBB.I_DAS.GetNextID()
    
    With oRadOrder
        SQLStmt = "INSERT INTO RAD_TEST (RAD_TEST_IEN, RAD_TEST_NAME," & _
                " SYNONYMS, FACILITY_NCID, RAD_TEST_TYPE )" & _
                "VALUES (" & DataID & "," & _
                goBB.DataConn.SQLQuote(UCase$(.TextName), RadName_LEN) & "," & _
                goBB.DataConn.SQLQuote(UCase$(.Synonyms), RadSynonyms_LEN) & "," & _
                goBB.Logon.FacilityNCID & ", '" & .TestType & "')"
    End With
    
    goBB.I_DAS.ExecuteSQL SQLStmt
    
Exit Sub

ErrHandler:
    Call goBB.SharedObj.ShowVBError(Err.Number, Err.Description, "colRadOrders.AddRadOrder()", "OrderAdmin", vbExclamation)
End Sub

Public Sub DeleteOrder(myIen As String)

    Dim SQLStmt As String
    
    On Error GoTo ErrHandler
        
'    SQLStmt = "Delete RAD_TEST "
    SQLStmt = "update RAD_TEST "
    
    '** SCR 68553 - 4/6/05 - FRANK
    '** New column DELETE_FLAG = 1 when Lab is DELETE from Order Admin
'    SQLStmt = SQLStmt & " set INACTIVE_FLAG = 1 "
    SQLStmt = SQLStmt & " set DELETE_FLAG = 1 "
    
    SQLStmt = SQLStmt & " where RAD_TEST_IEN = " & myIen
    SQLStmt = SQLStmt & " and FACILITY_NCID = " & goBB.Logon.FacilityNCID
    

    goBB.I_DAS.ExecuteSQL SQLStmt
          
Exit Sub

ErrHandler:
    Call goBB.SharedObj.ShowVBError(Err.Number, Err.Description, "colRadOrders.DeleteOrder()", "OrderAdmin", vbExclamation)

End Sub

Public Sub Add(orderItem As Object, Optional sKey As String)
    'create a new object
    Dim objNewMember As radOrder
    Set objNewMember = orderItem

    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If
    
    Set objNewMember = Nothing

End Sub

Public Property Get Item(vntIndexKey As Variant) As radOrder
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

Public Sub UpdateRadOrder(ByVal oRadOrder As radOrder)

Dim SQLStmt As String
Dim inactiveFlag As Integer

    On Error GoTo ErrHandler
    
    If oRadOrder.Active = False Then
        inactiveFlag = 1
    Else
        inactiveFlag = 0
    End If
    
    With oRadOrder
        SQLStmt = "UPDATE RAD_TEST SET " & _
                " INACTIVE_FLAG = " & inactiveFlag & _
                ", SYNONYMS = " & goBB.DataConn.SQLQuote(UCase$(.Synonyms), RadSynonyms_LEN) & _
                ", RAD_TEST_TYPE = '" & .TestType & "'" & _
                " WHERE RAD_TEST_IEN = '" & .Ien & "'"
        SQLStmt = SQLStmt & " and FACILITY_NCID = " & goBB.Logon.FacilityNCID
    End With
    
    goBB.I_DAS.ExecuteSQL SQLStmt
    
Exit Sub

ErrHandler:
    Call goBB.SharedObj.ShowVBError(Err.Number, Err.Description, "colRadOrders.UpdateRadOrder()", "OrderAdmin", vbExclamation)
End Sub

