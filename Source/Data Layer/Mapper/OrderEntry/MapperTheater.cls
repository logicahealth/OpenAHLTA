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
Implements DL_Interfaces.iMapperOrderEntry

Private Const MODULE_NAME As String = "DM_OrderEntry.MapperTheater"

Public Sub iMapperOrderEntry_Init(ByRef Support As iParameters, ByRef Cache As CacheMapper)
    On Error GoTo ErrHandler
    
    Set mCacheMapper = Cache
    Set mConn = Support.Item("DBConn").Value
    
    Exit Sub

ErrHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function iMapperOrderEntry_RetrieveOtherOrderTypes(ByRef params As iParameters) As OEOtherOrderTypes
    On Error GoTo ErrHandler
    
    Dim dtoOtherOrderTypes As DTOs.OEOtherOrderTypes
    Dim dtoOtherOrderType As DTOs.OEOtherOrderType
    Dim param As OEOtherOrderTypeSP
    Dim sWhereClause As String
    Dim rs As ADODB.Recordset
    Dim sSQL As String
    Dim fldID As ADODB.Field
    Dim fldCategory As ADODB.Field
    Dim fldDescription As ADODB.Field
    Dim fldComment As ADODB.Field
    Dim fldStatus As ADODB.Field
    Dim fldFacilityNCID As ADODB.Field
                
    'typecast the params object
    Set param = params
    
    'build the SQL where clause based on params
    sWhereClause = ""
    
    If param.FacilityNCID <> "" Then
           sWhereClause = " FACILITY_NCID = '" & param.FacilityNCID & "' "
    End If
    
    If params.Exists("Category") Then
        If (sWhereClause <> "") Then
            sWhereClause = sWhereClause & "AND"
        End If
        sWhereClause = sWhereClause & " CATEGORY = '" & param.Category & "' "
    End If
    
    If params.Exists("Status") Then
        If (sWhereClause <> "") Then
            sWhereClause = sWhereClause & "AND"
        End If
        sWhereClause = sWhereClause & " STATUS = " & param.Status & " "
    End If
    
    If params.Exists("ID") Then
        If (sWhereClause <> "") Then
            sWhereClause = sWhereClause & "AND"
        End If
        sWhereClause = sWhereClause & " ID = '" & param.ID & "' "
    End If
    
    'specifying the columns so I can refer to them by index later (faster)
    sSQL = "select ID, CATEGORY, DESCRIPTION, COMMENTS, STATUS, FACILITY_NCID " _
                & " from OE_OTHER_ORDER_TYPE "
    
    If (sWhereClause <> "") Then
        sSQL = sSQL & "WHERE " & sWhereClause
    End If
    
    'apply the default sort here
    '   (sorting by all three will also make populating the "filter" dropdowns more efficient)
    sSQL = sSQL & " order by CATEGORY, DESCRIPTION, COMMENTS"

    'run the query and populate the DTOs
    Set dtoOtherOrderTypes = New DTOs.OEOtherOrderTypes
    Set rs = mConn.CHCSII_DAS.OpenRecordset(sSQL)
    
    
    If (rs.RecordCount > 0) Then
        
        'uncertain of the syntax for this... will debug as needed
        Set fldID = rs.Fields("ID")
        Set fldCategory = rs.Fields("Category")
        Set fldDescription = rs.Fields("Description")
        Set fldComment = rs.Fields("Comments")
        Set fldStatus = rs.Fields("Status")
        Set fldFacilityNCID = rs.Fields("Facility_NCID")
        
        'iterate through the recordset
        rs.MoveFirst
        While Not rs.EOF
            Set dtoOtherOrderType = New DTOs.OEOtherOrderType
            With dtoOtherOrderType
            
                'again, not entirely certain of the syntax; but,
                '   pushing this to get the interface built.
                .ID = fldID.Value
                .Category = fldCategory.Value
                .Description = fldDescription.Value
                If (IsNull(fldComment.Value)) Then
                    .Comment = ""
                Else
                    .Comment = fldComment.Value
                End If
                .Status = fldStatus.Value
                .FacilityNCID = fldFacilityNCID.Value
            End With
            dtoOtherOrderTypes.Add dtoOtherOrderType
            rs.MoveNext
        Wend
    End If
    
    Set iMapperOrderEntry_RetrieveOtherOrderTypes = dtoOtherOrderTypes
    
    Exit Function
    
ErrHandler:
    Set iMapperOrderEntry_RetrieveOtherOrderTypes = Nothing
    Err.Raise Err.Number, MODULE_NAME & ".RetrieveOtherOrderTypeCategories", Err.Description
End Function

Public Function iMapperOrderEntry_RetrieveOtherOrderTypeCategories(ByVal FacilityNCID As String) As Collection
    'return a collection of strings containing the Categories found in the OtherOrderTypes table for
    '   a given facility
    On Error GoTo ErrHandler
    
    Dim sSQL  As String
    Dim colResult As Collection


    Dim rs As ADODB.Recordset

    
    Set colResult = New Collection
    sSQL = "SELECT DISTINCT ORDER_TYPE FROM OE_OTHER WHERE FACILITY_NCID = '" & FacilityNCID & "'"
    Set rs = mConn.CHCSII_DAS.OpenRecordset(sSQL)
    
    If (rs.RecordCount > 0) Then
        rs.MoveFirst
        While Not (rs.EOF)
            colResult.Add (rs(0).Value)
            rs.MoveNext
        Wend
    End If
    
    Set iMapperOrderEntry_RetrieveOtherOrderTypeCategories = colResult
    
    Exit Function
    
ErrHandler:
    Set iMapperOrderEntry_RetrieveOtherOrderTypeCategories = Nothing
    Err.Raise Err.Number, MODULE_NAME & ".RetrieveOtherOrderTypeCategories", Err.Description
End Function

Public Function iMapperOrderEntry_SaveOtherOrderTypes(ByRef OrderTypes As OEOtherOrderTypes) As Boolean
    On Error GoTo ErrHandler
    Dim i As Long
    Dim dtoOrderType As DTOs.OEOtherOrderType
    Dim rs As ADODB.Recordset
    Dim sSQL As String
    
    
    'iterate through the individual DTOs
    For i = 1 To OrderTypes.Count
        Set dtoOrderType = OrderTypes.Item(i)
        
        With dtoOrderType
            
            If .ID = "" Then .ID = "0"
            
            'temporary work-around
            If .FacilityNCID = "" Then .FacilityNCID = "0"
            
            'do we have a data id?  if so, we're "editing" a record
            If .ID <> 0 Then
                
                'have data id... get the record
                sSQL = "select * from OE_OTHER_ORDER_TYPE where ID = " & Val(.ID)
                Set rs = mConn.CHCSII_DAS.OpenRecordset(sSQL)
                
                
                If rs.RecordCount > 0 Then
                    'if the info is the same, we can just update the status
                    If rs.Fields("CATEGORY").Value = .Category _
                            And rs.Fields("DESCRIPTION").Value & vbNullString = .Description _
                            And rs.Fields("COMMENTS").Value & vbNullString = .Comment _
                            And rs.Fields("FACILITY_NCID").Value & vbNullString = .FacilityNCID Then
                        
                        sSQL = "update OE_OTHER_ORDER_TYPE set STATUS = " & .Status & " where ID = " & .ID
                        mConn.CHCSII_DAS.ExecuteSQL sSQL
                        
                        GoTo NextRecord
                        
                    'otherwise, if the record changed, mark the old one as deleted
                    Else
                        sSQL = "update OE_OTHER_ORDER_TYPE set STATUS = 2 where ID = " & .ID
                        mConn.CHCSII_DAS.ExecuteSQL sSQL
                        'GoTo NextRecord
                    End If
                End If
            End If
            
            'is there a matching record? (other than ID)
            sSQL = "select * from OE_OTHER_ORDER_TYPE where DESCRIPTION = " & SQLQuoteThunk(.Description) _
                        & " and COMMENTS = " & SQLQuoteThunk(.Comment) & " and FACILITY_NCID = " & SQLQuoteThunk(.FacilityNCID) _
                        & " and CATEGORY = " & SQLQuoteThunk(.Category)
            Set rs = mConn.CHCSII_DAS.OpenRecordset(sSQL)
            
            'having a problem where blank strings are getting saved (by Oracle?) as null
            '   falling back to check for that...
            If rs.RecordCount = 0 And .Comment = "" Then
                rs.Close
                sSQL = "select * from OE_OTHER_ORDER_TYPE where DESCRIPTION = " & SQLQuoteThunk(.Description) _
                            & " and COMMENTS is NULL and FACILITY_NCID = " & SQLQuoteThunk(.FacilityNCID) _
                            & " and CATEGORY = " & SQLQuoteThunk(.Category)
                Set rs = mConn.CHCSII_DAS.OpenRecordset(sSQL)
            End If
            
            If rs.RecordCount > 0 Then
                'found one... just update the status (and change the DTO's data ID accordingly)
                .ID = rs.Fields("ID").Value
                sSQL = "update OE_OTHER_ORDER_TYPE set STATUS = " & .Status & " where ID = " & .ID
                mConn.CHCSII_DAS.ExecuteSQL sSQL
                rs.Close
                GoTo NextRecord
            End If
            
            'we're still in this loop... we need to get a new ID and add the record
            sSQL = "select OE_OTHER_ORDER_TYPE_SEQ.NextVal from DUAL"
            Set rs = mConn.CHCSII_DAS.OpenRecordset(sSQL)
            .ID = rs.Fields(0).Value
            rs.Close
            
            sSQL = "insert into OE_OTHER_ORDER_TYPE fields (ID, CATEGORY, DESCRIPTION, COMMENTS, STATUS, FACILITY_NCID) " _
                    & " values (" & .ID & ", " & SQLQuoteThunk(.Category) & "," & SQLQuoteThunk(.Description) & "," & SQLQuoteThunk(.Comment) & "," & .Status & "," & SQLQuoteThunk(.FacilityNCID) & ")"
            mConn.CHCSII_DAS.ExecuteSQL sSQL
        End With
        
NextRecord:
    
    Next
    
    iMapperOrderEntry_SaveOtherOrderTypes = True
    Exit Function
    
ErrHandler:
    iMapperOrderEntry_SaveOtherOrderTypes = False
    Err.Raise Err.Number, MODULE_NAME & ".SaveOtherOrderTypes", Err.Description
End Function


Private Function SQLQuoteThunk(sSQL As String) As String
    If sSQL = "" Then
        SQLQuoteThunk = "''"
    Else
        SQLQuoteThunk = mConn.SQLQuote(sSQL)
    End If
End Function
