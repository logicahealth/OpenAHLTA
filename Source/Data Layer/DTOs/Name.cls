VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Name"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Implements iObjectPersist
Implements iState

Private Const OBJECT_VERSION = "1.0"
Private Const UNIQUE_IDENTIFIER = "Name"


Private mFirst_Name As String
Private mLast_Name As String
Private mMiddle_Name As String
Private mFirst_Name_Soundex As String
Private mLast_Name_Soundex As String
Private mFirst_Name2 As String
Private mLast_Name2 As String
Private mMiddle_Name2 As String
Private mSuffix As String
Private mDegree As String
Private mTitle As String
Private mName_Type_NCID As String
Private mState As enumObjectState

Public Property Get First_Name() As String
    First_Name = mFirst_Name
End Property
Public Property Let First_Name(ByVal value As String)
    mFirst_Name = value
End Property
Public Property Get Last_Name() As String
    Last_Name = mLast_Name
End Property
Public Property Let Last_Name(ByVal value As String)
    mLast_Name = value
End Property
Public Property Get Middle_Name() As String
    Middle_Name = mMiddle_Name
End Property
Public Property Let Middle_Name(ByVal value As String)
    mMiddle_Name = value
End Property
Public Property Get First_Name_Soundex() As String
    First_Name_Soundex = mFirst_Name_Soundex
End Property
Public Property Let First_Name_Soundex(ByVal value As String)
    mFirst_Name_Soundex = value
End Property
Public Property Get Last_Name_Soundex() As String
    Last_Name_Soundex = mLast_Name_Soundex
End Property
Public Property Let Last_Name_Soundex(ByVal value As String)
    mLast_Name_Soundex = value
End Property
Public Property Get First_Name2() As String
    First_Name2 = mFirst_Name2
End Property
Public Property Let First_Name2(ByVal value As String)
    mFirst_Name2 = value
End Property
Public Property Get Last_Name2() As String
    Last_Name2 = mLast_Name2
End Property
Public Property Let Last_Name2(ByVal value As String)
    mLast_Name2 = value
End Property
Public Property Get Middle_Name2() As String
    Middle_Name2 = mMiddle_Name2
End Property
Public Property Let Middle_Name2(ByVal value As String)
    mMiddle_Name2 = value
End Property
Public Property Get Suffix() As String
    Suffix = mSuffix
End Property
Public Property Let Suffix(ByVal value As String)
    mSuffix = value
End Property
Public Property Get Degree() As String
    Degree = mDegree
End Property
Public Property Let Degree(ByVal value As String)
    mDegree = value
End Property
Public Property Get Title() As String
    Title = mTitle
End Property
Public Property Let Title(ByVal value As String)
    mTitle = value
End Property
Public Property Get Name_Type_NCID() As String
    Name_Type_NCID = mName_Type_NCID
End Property
Public Property Let Name_Type_NCID(ByVal value As String)
    mName_Type_NCID = value
End Property




Private Function iObjectPersist_Copy() As Object
    Dim NewCopy As Name
    Dim State As iState
    
    On Error GoTo ErrHandler
    
    Set NewCopy = New Name
    
    With NewCopy
        .First_Name = mFirst_Name
        .Last_Name = mLast_Name
        .Middle_Name = mMiddle_Name
        .First_Name_Soundex = mFirst_Name_Soundex
        .Last_Name_Soundex = mLast_Name_Soundex
        .First_Name2 = mFirst_Name2
        .Last_Name2 = mLast_Name2
        .Middle_Name2 = mMiddle_Name2
        .Suffix = mSuffix
        .Degree = mDegree
        .Title = mTitle
        .Name_Type_NCID = mName_Type_NCID
        
        'iState interface
        Set State = NewCopy
        State.State = mState
    End With

    Set iObjectPersist_Copy = NewCopy

    Set NewCopy = Nothing
    Exit Function
ErrHandler:
    Err.Raise Err.Number, "DTOsName.iObjectPersist_Persist", Err.Description, Err.HelpFile, Err.HelpContext

End Function




Private Sub iObjectPersist_Persist(ByRef Writer As iWriter)

    On Error GoTo ErrHandler
    
    With Writer
        .InitHeader UNIQUE_IDENTIFIER, OBJECT_VERSION
        .WriteProperty "First_Name", mFirst_Name
        .WriteProperty "Last_Name", mLast_Name
        .WriteProperty "Middle_Name", mMiddle_Name
        .WriteProperty "First_Name_Soundex", mFirst_Name_Soundex
        .WriteProperty "Last_Name_Soundex", mLast_Name_Soundex
        .WriteProperty "First_Name2", mFirst_Name2
        .WriteProperty "Last_Name2", mLast_Name2
        .WriteProperty "Middle_Name2", mMiddle_Name2
        .WriteProperty "Suffix", mSuffix
        .WriteProperty "Degree", mDegree
        .WriteProperty "Title", mTitle
        .WriteProperty "Name_Type_NCID", mName_Type_NCID
        
        'iState interface
        .WriteProperty "iState.State", mState
    End With
    
    Exit Sub
    
ErrHandler:
    Err.Raise Err.Number, "DTOsName.iObjectPersist_Persist", Err.Description, Err.HelpFile, Err.HelpContext
End Sub




Private Function iObjectPersist_Restore(ByRef Reader As iReader) As Boolean
Dim Failed As Boolean

    On Error GoTo ErrHandler
  
    With Reader
        mFirst_Name = .ReadProperty("First_Name")
        mLast_Name = .ReadProperty("Last_Name")
        mMiddle_Name = .ReadProperty("Middle_Name")
        mFirst_Name_Soundex = .ReadProperty("First_Name_Soundex")
        mLast_Name_Soundex = .ReadProperty("Last_Name_Soundex")
        mFirst_Name2 = .ReadProperty("First_Name2")
        mLast_Name2 = .ReadProperty("Last_Name2")
        mMiddle_Name2 = .ReadProperty("Middle_Name2")
        mSuffix = .ReadProperty("Suffix")
        mDegree = .ReadProperty("Degree")
        mTitle = .ReadProperty("Title")
        mName_Type_NCID = .ReadProperty("Name_Type_NCID")
        
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
