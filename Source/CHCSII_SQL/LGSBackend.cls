VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LGSBackend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements iBackend

Private mCHCSConnection As Object   'Client-side interface to LGS (shared object in CWShared)

Private mRecordSet As ADODB.Recordset

Private Property Get iBackend_BOF() As Boolean
  iBackend_BOF = mRecordSet.BOF
End Property

Private Property Get iBackend_ColumnCount() As Integer
  iBackend_ColumnCount = mRecordSet.Fields.Count
End Property

Private Property Get iBackend_ColumnIndex(Name As String) As Integer
Dim Flds As Fields
Dim i As Integer
Dim sText As String

  iBackend_ColumnIndex = -1
  
  Set Flds = mRecordSet.Fields
  
  For i = 0 To Flds.Count - 1
    If UCase$(Flds(i).Name) = UCase$(Name) Then
      iBackend_ColumnIndex = i
      
      Exit For
    End If
  Next
  
  If iBackend_ColumnIndex = -1 Then
      sText = "There is no column named '" & Name & "'.  Please select from:"
    For i = 0 To Flds.Count - 1
        sText = sText & Chr(10) & Flds(i).Name
    Next
  
    Err.Raise vbObjectError + 1, "CHCSII_SQL:SqlOps.ColumnIndex", sText
  
  End If
  
End Property

Private Property Get iBackend_ColumnName(ByVal index As Integer) As String
    
    If Not (mRecordSet.BOF Or mRecordSet.EOF) Then
        If index > mRecordSet.Fields.Count Then
            Err.Raise vbObjectError + 1, "CHCSII_SQL:SqlOps.ColumnName", "Index (" & index & ") out of range.  There are " & mRecordSet.Fields.Count & " colummns."
        End If
        
        iBackend_ColumnName = mRecordSet.Fields(index - 1).Name
    End If

End Property

Private Function iBackend_CreateNewCDRClinicNCID(ByVal strClinicName As String, ByVal strFacilityNCID As String, ByVal strEnterpriseNCID As String) As String
  'Not Supported
End Function

Private Property Get iBackend_EOF() As Boolean
  iBackend_EOF = mRecordSet.EOF
End Property

Private Sub iBackend_Execute(SqlStatement As String)

  InitCHCSConnection
  
  If Left$(UCase$(SqlStatement), 6) = "SELECT" Then
    Set mRecordSet = mCHCSConnection.OpenRecordset(SqlStatement)
    
    If mRecordSet Is Nothing Then
      Set mRecordSet = New ADODB.Recordset
    End If
  Else
    mCHCSConnection.Execute SqlStatement
  End If
  
End Sub

Private Sub iBackend_ExecuteMulti(SqlStatements As Collection)
Dim i As Long

  For i = 1 To SqlStatements.Count
    iBackend_Execute SqlStatements.Item(i)
  Next
  
End Sub

Private Property Get iBackend_InfoType(Column As Variant) As String
  iBackend_InfoType = mRecordSet.Fields(Column).Type
End Property

Private Sub iBackend_MoveFirst()
  mRecordSet.MoveFirst
End Sub

Private Sub iBackend_MoveLast()
  mRecordSet.MoveLast
End Sub

Private Sub iBackend_MoveNext()
  mRecordSet.MoveNext
End Sub

Private Sub iBackend_MovePrevious()
  mRecordSet.MovePrevious
End Sub

Private Property Get iBackend_RecordCount() As Long
  iBackend_RecordCount = mRecordSet.RecordCount
End Property

Private Property Get iBackend_Source() As String
  iBackend_Source = mRecordSet.Source
End Property

Private Function iBackend_ValiadeFacilityAgainstCDR(FacilityNCID As String, ByVal EnterpriseNCID As String, NCIDIsRepresentation As Boolean, ByVal CreateIfNotFound As Boolean) As Boolean
  'Not Supported
End Function

Private Function iBackend_ValidateClinicAgainstCDR(ClinicNCID As String, ByVal FacilityNCID As String, FacilityIsValid As Boolean, NCIDIsRepresentation As Boolean, ByVal EnterpriseNCID As String, ByVal CreateIfNotFound As Boolean) As Boolean
  'Not Supported
End Function

Private Property Get iBackend_Value(Column As Variant) As Variant
  iBackend_Value = mRecordSet.Fields(Column).Value
End Property

Private Sub InitCHCSConnection()
Dim CoreShared As CWShared

  If mCHCSConnection Is Nothing Then
    Set CoreShared = New CWShared
    
    Set mCHCSConnection = CoreShared.CHCSConnection
  End If
  
  If mCHCSConnection Is Nothing Then
    Err.Raise 62000, "CHCSII_DataAccess:iBackend_OpenRecordset", "Unable to Connect to CHCS."
  End If
  
End Sub

