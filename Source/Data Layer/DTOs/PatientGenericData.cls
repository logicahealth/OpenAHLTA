VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PatientGenericData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True


Option Explicit

Implements iObjectPersist
Implements iState

Private Const OBJECT_VERSION = "1.0"
Private Const UNIQUE_IDENTIFIER = "PatientGenericData"


Private mErrorText As String
Private mOrdinal As Long
Private mParent As Long
Private mTypeNcid As Long
Private mValueNcid As Long
Private mValueRsid As Long
Private mValueText As String

Private mState As enumObjectState

Public Property Get ErrorText() As String
    ErrorText = mErrorText
End Property
Public Property Let ErrorText(ByVal value As String)
    mErrorText = value
End Property
Public Property Get Ordinal() As Long
    Ordinal = mOrdinal
End Property
Public Property Let Ordinal(ByVal value As Long)
    mOrdinal = value
End Property
Public Property Get Parent() As Long
    Parent = mParent
End Property
Public Property Let Parent(ByVal value As Long)
    mParent = value
End Property
Public Property Get TypeNcid() As Long
    TypeNcid = mTypeNcid
End Property
Public Property Let TypeNcid(ByVal value As Long)
    mTypeNcid = value
End Property
Public Property Get ValueNcid() As Long
    ValueNcid = mValueNcid
End Property
Public Property Let ValueNcid(ByVal value As Long)
    mValueNcid = value
End Property
Public Property Get ValueRsid() As Long
    ValueRsid = mValueRsid
End Property
Public Property Let ValueRsid(ByVal value As Long)
    mValueRsid = value
End Property
Public Property Get ValueText() As String
    ValueText = mValueText
End Property
Public Property Let ValueText(ByVal value As String)
    mValueText = value
End Property




Private Function iObjectPersist_Copy() As Object
    Dim NewCopy As PatientGenericData
    Dim State As iState
    
    On Error GoTo ErrHandler
    
    Set NewCopy = New PatientGenericData
    
    With NewCopy
        .ErrorText = mErrorText
        .Ordinal = mOrdinal
        .Parent = mParent
        .TypeNcid = mTypeNcid
        .ValueNcid = mValueNcid
        .ValueRsid = mValueRsid
        .ValueText = mValueText
        
        'iState interface
        Set State = NewCopy
        State.State = mState
    End With

    Set iObjectPersist_Copy = NewCopy

    Set NewCopy = Nothing
    Exit Function
ErrHandler:
    Err.Raise Err.Number, "DTOsPatientGenericData.iObjectPersist_Persist", Err.Description, Err.HelpFile, Err.HelpContext

End Function




Private Sub iObjectPersist_Persist(ByRef Writer As iWriter)

    On Error GoTo ErrHandler
    
    With Writer
        .InitHeader UNIQUE_IDENTIFIER, OBJECT_VERSION
        .WriteProperty "ErrorText", mErrorText
        .WriteProperty "Ordinal", mOrdinal
        .WriteProperty "Parent", mParent
        .WriteProperty "TypeNcid", mTypeNcid
        .WriteProperty "ValueNcid", mValueNcid
        .WriteProperty "ValueRsid", mValueRsid
        .WriteProperty "ValueText", mValueText
        
        'iState interface
        .WriteProperty "iState.State", mState
    End With
    
    Exit Sub
    
ErrHandler:
    Err.Raise Err.Number, "DTOsPatientGenericData.iObjectPersist_Persist", Err.Description, Err.HelpFile, Err.HelpContext
End Sub




Private Function iObjectPersist_Restore(ByRef Reader As iReader) As Boolean
Dim Failed As Boolean

    On Error GoTo ErrHandler
  
    With Reader
        mErrorText = .ReadProperty("ErrorText")
        mOrdinal = .ReadProperty("Ordinal")
        mParent = .ReadProperty("Parent")
        mTypeNcid = .ReadProperty("TypeNcid")
        mValueNcid = .ReadProperty("ValueNcid")
        mValueRsid = .ReadProperty("ValueRsid")
        mValueText = .ReadProperty("ValueText")
        
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

