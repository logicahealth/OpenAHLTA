VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ISqlOps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Enum eSeq
    esMAINSEQ = 0
    esAPPT_ID
    esAPPT_TRANS_ID
    esENC_TEMPLATE_ID
    esENCOUNTER_ID
    esIEN
    esPCM_CODE
    esPI_DATA_ID
    esUNIT_NUMBER
    esUSER_NUMBER
    esJDOE_NUMBER
    esTEMP_SSN
End Enum

Public Function Execute(SqlStatement As String) As Variant:     End Function
Public Sub ExecuteMulti(SqlStatements As Collection):           End Sub
Public Function ExecuteCmd(oAdoCmd As ADODB.Command):           End Function
Public Function GetNextID(Optional ByVal eSeqName As eSeq = esMAINSEQ): End Function
Public Function GetNextNcid():                                  End Function
Public Sub MovePrevious():                                      End Sub
Public Sub MoveNext():                                          End Sub
Public Sub MoveLast():                                          End Sub
Public Sub MoveFirst():                                         End Sub
Public Property Get Source() As String:                         End Property
Public Property Get RecordCount() As Long:                      End Property
Public Property Get EOF() As Boolean:                           End Property
Public Property Get BOF() As Boolean:                           End Property
Public Property Get Value(Column As Variant) As Variant:        End Property
Attribute Value.VB_UserMemId = 0
Attribute Value.VB_MemberFlags = "200"
Public Property Get InfoType(Column As Variant) As String:      End Property
Public Property Get ColumnIndex(Name As String) As Integer:     End Property
Public Property Get ColumnName(index As Integer) As String:     End Property
Public Property Get ColumnCount() As Integer:                   End Property


'<SCR 22426>

'PIC - add a function in CHCSII_SQL to add a new clinic into CDR

'Public Function NCIDIsCDRFacility(ByVal FacilityNCID As String) As Boolean
'
'End Function
'
'Public Function NCIDIsCDRFacilityRepresentation(ByRef InOutFacilityNCID As String) As Boolean
'
'End Function
'
'Public Function CreateNewCDRFacilityNCID(ByVal FacilityName As String, ByVal LogonObject As Variant) As String
'
'End Function
'
'Public Function NCIDIsCDRClinic(ByVal FacilityNCID As String) As Boolean
'
'End Function
'
'Public Function NCIDIsCDRCLinicRepresentation(ByRef InOutFacilityNCID As String) As Boolean
'
'End Function
'
'Public Function CreateNewCDRClinicNCID(ByVal FacilityName As String, ByVal LogonObject As Variant) As String
'
'End Function

Public Function ValiadeFacilityAgainstCDR(ByRef FacilityNcid As String, ByVal EnterpriseNCID As String, ByRef NCIDIsRepresentation As Boolean, ByVal CreateIfNotFound As Boolean) As Boolean


End Function

Public Function ValidateClinicAgainstCDR(ByRef ClinicNCID As String, ByVal FacilityNcid As String, ByRef FacilityIsValid As Boolean, ByRef NCIDIsRepresentation As Boolean, ByVal EnterpriseNCID As String, ByVal CreateIfNotFound As Boolean) As Boolean


End Function

'</SCR 22426>
Public Function ExecuteSP(Parms As CHCSII_Command.CHCSII_Cmd) As Variant:     End Function



