VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InpatientAdmission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--- begin DTO Builder; DTO=InpatientAdmission
'--- 3/18/2004 11:37:03 AM
'---
'--- NOTE: Any time you modify this class, increment the version number
'
'

Option Explicit

Implements iObjectPersist

Private Const OBJECT_VERSION = "1.0"
Private Const UNIQUE_IDENTIFIER = "InpatientAdmission"


Private mDateTime As Date   '
Private mDispositionDate As Date    '
Private mType As String '
Private mSynonym As String  '
Private mDispositionType As String  '
Private mAbsent As String   '
Private mMedicalHoldDispositionDate As Date '
Private mMeprsCode As String    '
Private mPatientIEN As String   '
Private mRegisterNumber As String   '



Public Property Get DateTime() As Date
    DateTime = mDateTime
End Property
Public Property Let DateTime(ByVal value As Date)
    mDateTime = value
End Property
Public Property Get DispositionDate() As Date
    DispositionDate = mDispositionDate
End Property
Public Property Let DispositionDate(ByVal value As Date)
    mDispositionDate = value
End Property
Public Property Get TypeAdmission() As String
    TypeAdmission = mType
End Property
Public Property Let TypeAdmission(ByVal value As String)
    mType = value
End Property
Public Property Get Synonym() As String
    Synonym = mSynonym
End Property
Public Property Let Synonym(ByVal value As String)
    mSynonym = value
End Property
Public Property Get DispositionType() As String
    DispositionType = mDispositionType
End Property
Public Property Let DispositionType(ByVal value As String)
    mDispositionType = value
End Property
Public Property Get Absent() As String
    Absent = mAbsent
End Property
Public Property Let Absent(ByVal value As String)
    mAbsent = value
End Property
Public Property Get MedicalHoldDispositionDate() As Date
    MedicalHoldDispositionDate = mMedicalHoldDispositionDate
End Property
Public Property Let MedicalHoldDispositionDate(ByVal value As Date)
    mMedicalHoldDispositionDate = value
End Property
Public Property Get MeprsCode() As String
    MeprsCode = mMeprsCode
End Property
Public Property Let MeprsCode(ByVal value As String)
    mMeprsCode = value
End Property
Public Property Get PatientIEN() As String
    PatientIEN = mPatientIEN
End Property
Public Property Let PatientIEN(ByVal value As String)
    mPatientIEN = value
End Property
Public Property Get RegisterNumber() As String
    RegisterNumber = mRegisterNumber
End Property
Public Property Let RegisterNumber(ByVal value As String)
    mRegisterNumber = value
End Property




Private Function iObjectPersist_Copy() As Object
    Dim NewCopy As InpatientAdmission
    Set NewCopy = New InpatientAdmission
    
    With NewCopy
        .DateTime = mDateTime
        .DispositionDate = mDispositionDate
        .TypeAdmission = mType
        .Synonym = mSynonym
        .DispositionType = mDispositionType
        .Absent = mAbsent
        .MedicalHoldDispositionDate = mMedicalHoldDispositionDate
        .MeprsCode = mMeprsCode
        .PatientIEN = mPatientIEN
        .RegisterNumber = mRegisterNumber
    End With

    Set iObjectPersist_Copy = NewCopy

End Function




Private Sub iObjectPersist_Persist(Writer As iWriter)
    With Writer
        .InitHeader UNIQUE_IDENTIFIER, OBJECT_VERSION
        .WriteProperty "DateTime", mDateTime
        .WriteProperty "DispositionDate", mDispositionDate
        .WriteProperty "Type", mType
        .WriteProperty "Synonym", mSynonym
        .WriteProperty "DispositionType", mDispositionType
        .WriteProperty "Absent", mAbsent
        .WriteProperty "MedicalHoldDispositionDate", mMedicalHoldDispositionDate
        .WriteProperty "MeprsCode", mMeprsCode
        .WriteProperty "PatientIEN", mPatientIEN
        .WriteProperty "RegisterNumber", mRegisterNumber
    End With
End Sub




Private Function iObjectPersist_Restore(Reader As iReader) As Boolean
Dim Failed As Boolean

    On Error GoTo ErrHandler
  
    With Reader
        mDateTime = .ReadProperty("DateTime")
        mDispositionDate = .ReadProperty("DispositionDate")
        mType = .ReadProperty("Type")
        mSynonym = .ReadProperty("Synonym")
        mDispositionType = .ReadProperty("DispositionType")
        mAbsent = .ReadProperty("Absent")
        mMedicalHoldDispositionDate = .ReadProperty("MedicalHoldDispositionDate")
        mMeprsCode = .ReadProperty("MeprsCode")
        mPatientIEN = .ReadProperty("PatientIEN")
        mRegisterNumber = .ReadProperty("RegisterNumber")
    End With

    iObjectPersist_Restore = Not Failed
  
    Exit Function
  
ErrHandler:
    Failed = True
    Resume Next

End Function


Private Property Get iObjectPersist_UniqueIdentifier() As String
    iObjectPersist_UniqueIdentifier = UNIQUE_IDENTIFIER
End Property

Private Property Get iObjectPersist_Version() As String
    iObjectPersist_Version = OBJECT_VERSION
End Property
'--- end DTO Builder

