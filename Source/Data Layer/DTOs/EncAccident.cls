VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EncAccident"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--- begin DTO Builder; DTO=EncAccident
'--- 3/16/04 7:57:01 PM
'---
'--- NOTE: Any time you modify this class, increment the version number
'
'

Option Explicit

Implements iObjectPersist
Implements iState

Private Const OBJECT_VERSION = "1.0"
Private Const UNIQUE_IDENTIFIER = "EncAccident"

Private mState As enumObjectState
Private mAccidentDate As Date   '
Private mCausedByAccident As Boolean    '
Private mCauseCodeNCID As String    '
Private mLocationIEN As String  '
Private mLocationCode As String '
Private mApptID As String   '
Private mClinicNCID As String   '



Public Property Get AccidentDate() As Date
    AccidentDate = mAccidentDate
End Property
Public Property Let AccidentDate(value As Date)
    mAccidentDate = value
End Property
Public Property Get CausedByAccident() As Boolean
    CausedByAccident = mCausedByAccident
End Property
Public Property Let CausedByAccident(value As Boolean)
    mCausedByAccident = value
End Property
Public Property Get CauseCodeNCID() As String
    CauseCodeNCID = mCauseCodeNCID
End Property
Public Property Let CauseCodeNCID(value As String)
    mCauseCodeNCID = value
End Property
Public Property Get LocationIEN() As String
    LocationIEN = mLocationIEN
End Property
Public Property Let LocationIEN(value As String)
    mLocationIEN = value
End Property
Public Property Get LocationCode() As String
    LocationCode = mLocationCode
End Property
Public Property Let LocationCode(value As String)
    mLocationCode = value
End Property
Public Property Get ApptID() As String
    ApptID = mApptID
End Property
Public Property Let ApptID(value As String)
    mApptID = value
End Property
Public Property Get ClinicNCID() As String
    ClinicNCID = mClinicNCID
End Property
Public Property Let ClinicNCID(value As String)
    mClinicNCID = value
End Property




Private Function iObjectPersist_Copy() As Object
    Dim NewCopy As EncAccident
    Dim State As iState
    
    Set NewCopy = New EncAccident
    
    With NewCopy
        .AccidentDate = mAccidentDate
        .CausedByAccident = mCausedByAccident
        .CauseCodeNCID = mCauseCodeNCID
        .LocationIEN = mLocationIEN
        .LocationCode = mLocationCode
        .ApptID = mApptID
        .ClinicNCID = mClinicNCID
        
        'iState interface
        Set State = NewCopy
        State.State = mState
    End With

    Set iObjectPersist_Copy = NewCopy

End Function




Private Sub iObjectPersist_Persist(Writer As iWriter)
    With Writer
        .InitHeader UNIQUE_IDENTIFIER, OBJECT_VERSION
        .WriteProperty "AccidentDate", mAccidentDate
        .WriteProperty "CausedByAccident", mCausedByAccident
        .WriteProperty "CauseCodeNCID", mCauseCodeNCID
        .WriteProperty "LocationIEN", mLocationIEN
        .WriteProperty "LocationCode", mLocationCode
        .WriteProperty "ApptID", mApptID
        .WriteProperty "ClinicNCID", mClinicNCID
        
        'iState interface
        .WriteProperty "iState.State", mState
    End With
End Sub




Private Function iObjectPersist_Restore(Reader As iReader) As Boolean
Dim Failed As Boolean

    On Error GoTo ErrHandler
  
    With Reader
        mAccidentDate = .ReadProperty("AccidentDate")
        mCausedByAccident = .ReadProperty("CausedByAccident")
        mCauseCodeNCID = .ReadProperty("CauseCodeNCID")
        mLocationIEN = .ReadProperty("LocationIEN")
        mLocationCode = .ReadProperty("LocationCode")
        mApptID = .ReadProperty("ApptID")
        mClinicNCID = .ReadProperty("ClinicNCID")
        
        'iState interface
        mState = .ReadProperty("iState.State")
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

Private Property Let iState_State(ByVal NewValue As enumObjectState)
    mState = NewValue
End Property

Private Property Get iState_State() As enumObjectState
    iState_State = mState
End Property

Private Property Let iState_ID(ByVal RHS As String)

End Property

Private Property Get iState_ID() As String
    iState_ID = UNIQUE_IDENTIFIER
End Property
