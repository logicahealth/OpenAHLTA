VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XMLToRS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Count As Long
Private RS_Hold As Dictionary
Private CurrentTableName As String
Private SegmentName As String

Private Sub Class_Initialize()
   
   On Error GoTo Class_InitializeError
   
   ' initialize properties
   Set RS_Hold = New Dictionary
   Count = 0
   Exit Sub
Class_InitializeError:
   Dim ErrMsg As String
   ErrMsg = Err.Number & vbCrLf
   ErrMsg = ErrMsg & Err.Description & vbCrLf
   Call MsgBxARMd(ErrMsg, vbCritical, "EvntProc Class_Initialize")
End Sub

Private Sub Class_Terminate()
   Set RS_Hold = Nothing
   Count = 0
End Sub

Public Function FindOrCreateRS(TableName As String, ColumnName As String, Value As String, RetVal As EP_Errors) As ADODB.Recordset
   Dim RS As ADODB.Recordset
   Dim lResult As Long
   Dim TabName As String
   Dim vEntry As Variant
   Dim FieldName(1) As String
   Dim FieldValue(1) As String
   
   On Error GoTo FindOrCreateRSError
   
   Set FindOrCreateRS = Nothing
   TabName = UCase(TableName)
   If TabName <> "" Then
      If RS_Hold.Exists(TabName) Then
         Set RS = RS_Hold(TabName)
         lResult = PositionRS(RS)
         If lResult = ERROR_SUCCESS Then
            FieldName(0) = ColumnName
            FieldValue(0) = Value
            lResult = PopulateRS(RS, FieldValue, FieldName)
            Set FindOrCreateRS = RS_Hold(TabName)
         End If
         RetVal = lResult
         Exit Function
      Else
         lResult = GetFieldProps(TabName, RS, True)
         If lResult = ERROR_SUCCESS Then
            FieldName(0) = ColumnName
            FieldValue(0) = Value
            lResult = PopulateRS(RS, FieldValue, FieldName)
            RS_Hold.Add TabName, RS
            Set FindOrCreateRS = RS
            RetVal = ERROR_SUCCESS
            Exit Function
         Else
            Set FindOrCreateRS = Nothing
            Call MsgBxARMd("ERROR_UNABLE_TO_CREATE_RECORDSET", vbCritical, "FindOrCreateRS")
            RetVal = ERROR_UNABLE_TO_CREATE_RECORDSET
            Exit Function
         End If
      End If
   Else
      ' invalid TableName
      RetVal = ERROR_INVALID_TABLENAME ' bad return
      Set FindOrCreateRS = Nothing
   End If
   Exit Function
FindOrCreateRSError:
   Dim ErrMsg As String
   ErrMsg = Err.Number & vbCrLf
   ErrMsg = ErrMsg & Err.Description & vbCrLf
   Call MsgBxARMd(ErrMsg, vbCritical, "EvntProc Class_Initialize")
End Function

Public Function UpdateRS_Hold(TableName As String, RS As ADODB.Recordset) As Long
   Dim TabName As String
   
   On Error GoTo UpdateRS_HoldError
   
   TabName = UCase(TableName)
   If TabName <> "" Then
      If Not (RS Is Nothing) Then
         If RS_Hold.Exists(TabName) Then
            RS_Hold.Remove TabName
            RS_Hold.Add TabName, RS
            UpdateRS_Hold = ERROR_SUCCESS ' good return
            Exit Function
         Else
            RS_Hold.Add TabName, RS
            UpdateRS_Hold = ERROR_SUCCESS ' good return
            Exit Function
         End If
      Else
         UpdateRS_Hold = ERROR_INVALID_RECORDSET_OPERATION ' bad return
         Exit Function
      End If
   Else
      UpdateRS_Hold = ERROR_INVALID_TABLENAME ' bad return
      Exit Function
   End If
   Exit Function
UpdateRS_HoldError:
   Dim ErrMsg As String
   ErrMsg = Err.Number & vbCrLf
   ErrMsg = ErrMsg & Err.Description & vbCrLf
   Call MsgBxARMd(ErrMsg, vbCritical, "EvntProc UpdateRS_Hold")
End Function

Public Function PositionRS(RS As ADODB.Recordset) As Long
   On Error GoTo PositionRSError
   
   If Count = 0 Then
      If Not (RS Is Nothing) Then
         RS.MoveFirst
      End If
      PositionRS = ERROR_SUCCESS ' good return
      Exit Function
   Else
      If Not (RS Is Nothing) Then
         If Count = 1 Then
            RS.MoveFirst
         Else
         End If
      Else
      End If
   End If
   Exit Function
PositionRSError:
   Dim ErrMsg As String
   ErrMsg = Err.Number & vbCrLf
   ErrMsg = ErrMsg & Err.Description & vbCrLf
   Call MsgBxARMd(ErrMsg, vbCritical, "EvntProc PositionRS")
End Function

Public Function SaveXMLValueToRS(RS As ADODB.Recordset, ColumnName As String, Value As String) As Long
   Dim vTemp As Variant
   Dim RetVal As Long
   Dim DataValues() As String
   Dim DataFields() As String
   
   On Error GoTo SaveXMLValueToRSError
   If Not (RS Is Nothing) Then
      If ColumnName <> "" Then
         DataValues(0) = Value
         DataFields(0) = ColumnName
         RetVal = PopulateRS(RS, DataValues, DataFields)
         If RetVal <> ERROR_SUCCESS Then
         End If
      Else
         SaveXMLValueToRS = ERROR_INVALID_COLUMNNAME ' bad return
         Exit Function
      End If
   Else
      SaveXMLValueToRS = ERROR_RECORDSET_IS_NOTHING  ' bad return
      Exit Function
   End If
   Exit Function
SaveXMLValueToRSError:
   Dim ErrMsg As String
   ErrMsg = Err.Number & vbCrLf
   ErrMsg = ErrMsg & Err.Description & vbCrLf
   Call MsgBxARMd(ErrMsg, vbCritical, "EvntProc SaveXMLValueToRS")
End Function

Public Property Let StoreTableName(TableName As String)
   CurrentTableName = TableName
End Property

Public Property Get RetrieveTableName() As String
   RetrieveTableName = CurrentTableName
End Property

Public Property Let StoreSegmentName(sSegmentName As String)
   SegmentName = sSegmentName
End Property

Public Property Get RetrieveSegmentName() As String
   RetrieveSegmentName = SegmentName
End Property

