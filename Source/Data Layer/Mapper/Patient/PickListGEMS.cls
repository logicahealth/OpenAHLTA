VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PickListGEMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements iPickList


Private Function iPickList_ComboxSelectStrings(ByVal sDomainNCID As String) As DTOs.PatientSelectStrings
    Dim oRS As adodb.Recordset
    Dim oConn As CHCSII_CONN.Conn
    Dim sSQL As String
    Dim oSelectString As DTOs.PatientSelectString
    Dim oSelectStrings As DTOs.PatientSelectStrings

    On Error GoTo ErrHandler
    
    Set oSelectStrings = New DTOs.PatientSelectStrings
    Set oConn = New CHCSII_CONN.Conn
    
    sSQL = "select n.ncid, n.description"
    sSQL = sSQL & " from ncids n"
    sSQL = sSQL & " where n.domainncid = " & sDomainNCID
    
    Set oRS = oConn.CHCSII_DAS(Auto).OpenRecordset(sSQL)
    
    If oRS.RecordCount > 0 Then
        oRS.MoveFirst
        Do While Not oRS.EOF
            Set oSelectString = Nothing
            Set oSelectString = New DTOs.PatientSelectString
            If oRS("ncid") <> 152964 Then
                oSelectString.ID = oRS("ncid")
                oSelectString.Value = Trim$(oRS("description"))
                oSelectStrings.Add oSelectString, oSelectString.ID
            End If
            oRS.MoveNext
        Loop
    End If
    
    Set iPickList_ComboxSelectStrings = oSelectStrings
    
    GoTo CleanUp
    
ErrHandler:
    Err.Raise Err.Number, "PickListGEMS.iPickList_ComboxSelectStrings", Err.Description
CleanUp:
    Set oRS = Nothing
    Set oConn = Nothing
    Set oSelectString = Nothing
    Set oSelectStrings = Nothing

End Function

Private Function iPickList_FMPSelectStrings(ByVal sNumDomainNCID As String, ByVal sTextDomainNCID As String) As DTOs.PatientSelectStrings
    Dim oRS As adodb.Recordset
    Dim oConn As CHCSII_CONN.Conn
    Dim sSQL As String
    Dim oSelectString As DTOs.PatientSelectString
    Dim oSelectStrings As DTOs.PatientSelectStrings

    On Error GoTo ErrHandler
    
    Set oSelectStrings = New DTOs.PatientSelectStrings
    Set oConn = New CHCSII_CONN.Conn
    
    sSQL = "select n.ncid FMP_Ncid, n.description FMP_Num, t.Description FMP_Text"
    sSQL = sSQL & " from ncids n, ncids t"
    sSQL = sSQL & " where n.domainncid = " & sNumDomainNCID
    sSQL = sSQL & " and t.domainncid = " & sTextDomainNCID
    sSQL = sSQL & " and n.ncid = t.ncid"
    
    Set oRS = oConn.CHCSII_DAS(Auto).OpenRecordset(sSQL)
    
    If oRS.RecordCount > 0 Then
        oRS.MoveFirst
        Do While Not oRS.EOF
            Set oSelectString = Nothing
            Set oSelectString = New DTOs.PatientSelectString
            
            oSelectString.ID = oRS("FMP_Ncid")
            oSelectString.Value = Trim$(oRS("FMP_Num")) & " - " & Trim$(oRS("FMP_Text"))
            oSelectStrings.Add oSelectString, oSelectString.ID
            
            oRS.MoveNext
        Loop
    End If
    
    Set iPickList_FMPSelectStrings = oSelectStrings
    
    GoTo CleanUp
    
ErrHandler:
    Err.Raise Err.Number, "PickListGEMS.iPickList_ComboxSelectStrings", Err.Description
CleanUp:
    Set oRS = Nothing
    Set oConn = Nothing
    Set oSelectString = Nothing
    Set oSelectStrings = Nothing

End Function

Private Function iPickList_PATCATSelectStrings(ByVal sNumDomainNCID As String) As DTOs.PatientSelectStrings
    Dim oRS As adodb.Recordset
    Dim oConn As CHCSII_CONN.Conn
    Dim sSQL As String
    Dim oSelectString As DTOs.PatientSelectString
    Dim oSelectStrings As DTOs.PatientSelectStrings

    On Error GoTo ErrHandler
    
    Set oSelectStrings = New DTOs.PatientSelectStrings
    Set oConn = New CHCSII_CONN.Conn
    
    sSQL = "select n.ncid, n.description"
    sSQL = sSQL & " from ncids n"
    sSQL = sSQL & " where n.domainncid = " & sNumDomainNCID
    sSQL = sSQL & " Order by n.description"
    
    Set oRS = oConn.CHCSII_DAS(Auto).OpenRecordset(sSQL)
    
    If oRS.RecordCount > 0 Then
        oRS.MoveFirst
        Do While Not oRS.EOF
            Set oSelectString = Nothing
            Set oSelectString = New DTOs.PatientSelectString
            
            oSelectString.ID = oRS("ncid")
            oSelectString.Value = Trim$(oRS("description"))
            oSelectStrings.Add oSelectString, oSelectString.ID
            
            oRS.MoveNext
        Loop
    End If
    
    Set iPickList_PATCATSelectStrings = oSelectStrings
    
    GoTo CleanUp
    
ErrHandler:
    Err.Raise Err.Number, "PickListGEMS.iPickList_PATCATSelectStrings", Err.Description
CleanUp:
    Set oRS = Nothing
    Set oConn = Nothing
    Set oSelectString = Nothing
    Set oSelectStrings = Nothing

End Function

Private Function iPickList_ReturnRepFromNCID(ByVal sDomainNCID As String, ByVal sNCID As String) As String
    Dim oRS As adodb.Recordset
    Dim oConn As CHCSII_CONN.Conn
    Dim sSQL As String

    On Error GoTo ErrHandler
    
    If LenB(sNCID) = 0 Then
        iPickList_ReturnRepFromNCID = vbNullString
        GoTo CleanUp
    End If
    
    
    Set oConn = New CHCSII_CONN.Conn
    
    sSQL = "select n.ncid, n.description"
    sSQL = sSQL & " from ncids n"
    sSQL = sSQL & " where n.domainncid = " & sDomainNCID
    sSQL = sSQL & " and ncid = " & sNCID
    
    Set oRS = oConn.CHCSII_DAS(Auto).OpenRecordset(sSQL)
    
    If oRS.RecordCount > 0 Then
        iPickList_ReturnRepFromNCID = Trim$(oRS("description"))
    Else
        iPickList_ReturnRepFromNCID = vbNullString
    End If
     
    GoTo CleanUp
    
ErrHandler:
    Err.Raise Err.Number, "PickListGEMS.iPickList_ReturnRepFromNCID", Err.Description
CleanUp:
    Set oRS = Nothing
    Set oConn = Nothing
End Function




