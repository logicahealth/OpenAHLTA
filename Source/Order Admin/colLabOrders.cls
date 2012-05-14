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
Attribute VB_Name = "colLabOrders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"WellnessReminder"
Attribute VB_Ext_KEY = "Member0" ,"WellnessReminder"
Option Explicit
Private mCol As Collection

Public Function loadData() As Boolean

    Dim sSql As String
    Dim rs As New ADODB.Recordset
    Dim objOrder As labOrder
    
    On Error GoTo ErrHandler
        
    '--- SCR 64695   Sherry Wang   1/21/05
    '--- Deleted the following code:  (INACTIVE_FLAG is NULL or INACTIVE_FLAG = 0) from where clause
     sSql = "select INACTIVE_FLAG, LAB_TEST_IEN, LAB_TEST_NAME, SYNONYMS, VALUE_RANGE "
     sSql = sSql & " from LAB_TEST_OE "
     sSql = sSql & " where "
     sSql = sSql & "  lab_test_name is not null"
     sSql = sSql & " AND lab_test_ien <> 0"
     sSql = sSql & " and FACILITY_NCID = " & goBB.Logon.FacilityNCID
     
    '** SCR 68553 - 4/6/05 - FRANK
    '** New column DELETE_FLAG = 1 when Lab is DELETE from Order Admin
    sSql = sSql & " and DELETE_FLAG <> 1"

     sSql = sSql & " Order by lab_test_name "
    
     Set rs = goBB.I_DAS.ExecuteSQL(sSql)
     
     Do Until rs.EOF
         Set objOrder = New labOrder
         If IsNull(rs.Fields("INACTIVE_FLAG")) Then
            objOrder.Active = True
         Else
            If Val(rs.Fields("INACTIVE_FLAG")) = 1 Then
                objOrder.Active = False
            Else
                objOrder.Active = True
            End If
         End If
         objOrder.Ien = IsNullValue(rs.Fields("LAB_TEST_IEN"), 0)
         objOrder.TextName = CStr("" & "" & rs.Fields("LAB_TEST_NAME"))
         objOrder.Synonyms = CStr("" & "" & rs.Fields("SYNONYMS"))
         objOrder.ValueRange = CStr("" & "" & rs.Fields("VALUE_RANGE"))

         mCol.Add objOrder
         Set objOrder = Nothing
         rs.MoveNext
     Loop
     
     Set rs = Nothing
     loadData = True
     
Exit Function

ErrHandler:
    Call goBB.SharedObj.ShowVBError(Err.Number, Err.Description, "colLabOrders.loadData()", "ReasonForVisit", vbExclamation)
    loadData = False
    
End Function
Public Sub AddLabOrder(ByVal oLabOrder As labOrder)

Dim SQLStmt As String
Dim DataID As String

    On Error GoTo ErrHandler
    DataID = goBB.I_DAS.GetNextID()
    
    With oLabOrder
        SQLStmt = "INSERT INTO LAB_TEST_OE (LAB_TEST_IEN, LAB_TEST_NAME," & _
                " SYNONYMS, VALUE_RANGE, FACILITY_NCID ) " & _
                "VALUES (" & DataID & "," & _
                goBB.DataConn.SQLQuote(UCase$(.TextName), LabName_LEN) & "," & _
                goBB.DataConn.SQLQuote(UCase$(.Synonyms), LabSynonyms_LEN) & "," & _
                goBB.DataConn.SQLQuote(.ValueRange, ValueRange_LEN) & "," & _
                goBB.Logon.FacilityNCID & ")"
    End With
    goBB.I_DAS.ExecuteSQL SQLStmt
    
Exit Sub

ErrHandler:
    Call goBB.SharedObj.ShowVBError(Err.Number, Err.Description, "colLabOrders.AddLabOrder()", "OrderAdmin", vbExclamation)
End Sub

Public Sub DeleteOrder(myIen As String)

    Dim SQLStmt As String
    
    On Error GoTo ErrHandler
        
    SQLStmt = "update LAB_TEST_OE "
    
    '** SCR 68553 - 4/6/05 - FRANK
    '** New column DELETE_FLAG = 1 when Lab is DELETE from Order Admin
'    SQLStmt = SQLStmt & " set INACTIVE_FLAG = 1 "
    SQLStmt = SQLStmt & " set DELETE_FLAG = 1 "
    
    SQLStmt = SQLStmt & " where LAB_TEST_IEN = " & myIen
    SQLStmt = SQLStmt & " and FACILITY_NCID = " & goBB.Logon.FacilityNCID
    
    goBB.I_DAS.ExecuteSQL SQLStmt
          
Exit Sub

ErrHandler:
    Call goBB.SharedObj.ShowVBError(Err.Number, Err.Description, "colLabOrders.DeleteOrder()", "OrderAdmin", vbExclamation)

End Sub

Public Sub Add(orderItem As Object, Optional sKey As String)
    'create a new object
    Dim objNewMember As labOrder
    Set objNewMember = orderItem


    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If
    
    Set objNewMember = Nothing

End Sub

Public Property Get Item(vntIndexKey As Variant) As labOrder
Attribute Item.VB_UserMemId = 0
    Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
        Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
        mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
        Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

Public Sub UpdateLabOrder(ByVal oLabOrder As labOrder)
Dim SQLStmt As String
Dim inactiveFlag As Integer

    On Error GoTo ErrHandler
    
    If oLabOrder.Active = False Then
        inactiveFlag = 1
    Else
        inactiveFlag = 0
    End If
    
    With oLabOrder
        'SCR 47780 - Jane Shen 01/07/2004
        SQLStmt = "UPDATE  LAB_TEST_OE SET " & _
                " INACTIVE_FLAG = " & inactiveFlag & _
                ", SYNONYMS = " & goBB.DataConn.SQLQuote(UCase$(.Synonyms), LabSynonyms_LEN) & _
                ",VALUE_RANGE= " & goBB.DataConn.SQLQuote(UCase$(.ValueRange), ValueRange_LEN) & _
                " WHERE LAB_TEST_IEN = '" & oLabOrder.Ien & "'"
        SQLStmt = SQLStmt & " and FACILITY_NCID = " & goBB.Logon.FacilityNCID
        
   End With
    
    goBB.I_DAS.ExecuteSQL SQLStmt
    
    Exit Sub

ErrHandler:
    Call goBB.SharedObj.ShowVBError(Err.Number, Err.Description, "colLabOrders.UpdateLabOrder()", "OrderAdmin", vbExclamation)
End Sub

