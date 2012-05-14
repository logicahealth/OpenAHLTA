VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SqlOps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Implements ICHCSII_SQL.ISqlOps

Private mBackend As iBackend

Private Sub Class_Initialize()
Dim CoreShared As CWShared

  Set CoreShared = New CWShared
  
  If CoreShared.AppMode <> modeCHCSI_GUI Then
    Set mBackend = New CDRBackend
  Else
    Set mBackend = New LGSBackend
  End If

End Sub

Private Sub Class_Terminate()
  Set mBackend = Nothing
End Sub

Private Property Get ISqlOps_BOF() As Boolean
  ISqlOps_BOF = BOF
End Property

Private Property Get ISqlOps_ColumnCount() As Integer
  ISqlOps_ColumnCount = ColumnCount
End Property

Private Property Get ISqlOps_ColumnIndex(Name As String) As Integer
  ISqlOps_ColumnIndex = ColumnIndex(Name)
End Property

Private Property Get ISqlOps_ColumnName(index As Integer) As String
  ISqlOps_ColumnName = ColumnName(index)
End Property

Private Property Get ISqlOps_EOF() As Boolean
  ISqlOps_EOF = EOF
End Property

Private Function ISqlOps_Execute(SqlStatement As String) As Variant
  Call Execute(SqlStatement)
End Function

Private Function ISqlOps_ExecuteCmd(oAdoCmd As ADODB.Command) As Variant
  ISqlOps_ExecuteCmd = False
End Function

Private Sub ISqlOps_ExecuteMulti(SqlStatements As Collection)
  Call ExecuteMulti(SqlStatements)
End Sub

Private Function ISqlOps_ExecuteSP(SParms As CHCSII_Command.CHCSII_Cmd) As Variant
  ISqlOps_ExecuteSP = False
End Function

Private Function ISqlOps_GetNextID(Optional ByVal sSeqName As ICHCSII_SQL.eSeq = 0&) As Variant
  ISqlOps_GetNextID = 0&
End Function

Private Function ISqlOps_GetNextNcid() As Variant
  ISqlOps_GetNextNcid = 0
End Function

Private Property Get ISqlOps_InfoType(Column As Variant) As String
  ISqlOps_InfoType = InfoType(Column)
End Property

Private Sub ISqlOps_MoveFirst()
  MoveFirst
End Sub

Private Sub ISqlOps_MoveLast()
  MoveLast
End Sub

Private Sub ISqlOps_MoveNext()
  MoveNext
End Sub

Private Sub ISqlOps_MovePrevious()
  MovePrevious
End Sub

Private Property Get ISqlOps_RecordCount() As Long
  ISqlOps_RecordCount = RecordCount
End Property

Private Property Get ISqlOps_Source() As String
  ISqlOps_Source = Source
End Property

Private Function ISqlOps_ValidateFacility(ByVal FacilityNCID As Long, _
                                          ByVal CreateIfNotFound As Boolean, _
                                          Optional ByVal ClinicNCID As Long = 0&) As Boolean
  ISqlOps_ValidateFacility = False
End Function

Private Property Get ISqlOps_Value(Column As Variant) As Variant
  ISqlOps_Value = Value(Column)
End Property

Private Function ISqlOps_ValiadeFacilityAgainstCDR(FacilityNCID As String, ByVal EnterpriseNCID As String, NCIDIsRepresentation As Boolean, ByVal CreateIfNotFound As Boolean) As Boolean
  ISqlOps_ValiadeFacilityAgainstCDR = mBackend.ValiadeFacilityAgainstCDR(FacilityNCID, EnterpriseNCID, NCIDIsRepresentation, CreateIfNotFound)
End Function

Private Function ISqlOps_ValidateClinicAgainstCDR(ClinicNCID As String, ByVal FacilityNCID As String, FacilityIsValid As Boolean, NCIDIsRepresentation As Boolean, ByVal EnterpriseNCID As String, ByVal CreateIfNotFound As Boolean) As Boolean
  ISqlOps_ValidateClinicAgainstCDR = mBackend.ValidateClinicAgainstCDR(ClinicNCID, FacilityNCID, FacilityIsValid, NCIDIsRepresentation, EnterpriseNCID, CreateIfNotFound)
End Function

Public Property Get BOF() As Boolean
  BOF = mBackend.BOF
End Property

Public Property Get ColumnCount() As Integer
  ColumnCount = mBackend.ColumnCount
End Property

Public Property Get ColumnIndex(Name As String) As Integer
  ColumnIndex = mBackend.ColumnIndex(Name)
End Property

Public Property Get ColumnName(index As Integer) As String
  ColumnName = mBackend.ColumnName(index)
End Property

Public Property Get EOF() As Boolean
  EOF = mBackend.EOF
End Property

Public Property Get InfoType(Column As Variant) As String
  InfoType = mBackend.InfoType(Column)
End Property

Public Property Get RecordCount() As Long
  RecordCount = mBackend.RecordCount
End Property

Public Property Get Source() As String
  Source = mBackend.Source
End Property

Public Property Get Value(Column As Variant) As Variant
Attribute Value.VB_UserMemId = 0
  Value = mBackend.Value(Column)
End Property

Public Function CreateNewCDRClinicNCID(ByVal strClinicName As String, ByVal strFacilityNCID As String, ByVal strEnterpriseNCID As String) As String
  CreateNewCDRClinicNCID = mBackend.CreateNewCDRClinicNCID(strClinicName, strFacilityNCID, strEnterpriseNCID)
End Function

Public Sub Execute(SqlStatement As String)
  mBackend.Execute SqlStatement
End Sub

Public Sub ExecuteMulti(SqlStatements As Collection)
  mBackend.ExecuteMulti SqlStatements
End Sub

Public Sub MoveFirst()
  mBackend.MoveFirst
End Sub

Public Sub MoveLast()
  mBackend.MoveLast
End Sub

Public Sub MoveNext()
  mBackend.MoveNext
End Sub

Public Sub MovePrevious()
  mBackend.MovePrevious
End Sub

