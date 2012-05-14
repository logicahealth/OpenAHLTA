VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Gems_Names"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "collection of names"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"GEMS_Name"
Attribute VB_Ext_KEY = "Member0" ,"GEMS_Name"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable to hold collection
Private mCol As Collection
Private mCurr   As GEMS_Name
Private mvarIdx As Long
Private mvarEOF As Boolean

Private Sub Class_Initialize()
    Set mCol = New Collection
    Set mCurr = New GEMS_Name
    mvarIdx = 0
End Sub


Private Sub Class_Terminate()
    Set mCol = Nothing
    Set mCurr = Nothing
End Sub

Public Function MoveNext() As Boolean
  mvarIdx = mvarIdx + 1
  If mvarIdx <= mCol.Count Then
    Set mCurr = mCol.mvarIdx
    mvarEOF = False
    MoveNext = True
  Else
    Set mCurr = New GEMS_Name
    mvarEOF = True
    mvarIdx = 0
    MoveNext = False
  End If
End Function

Public Function MoveFirst() As Boolean
   mvarIdx = 1
  If mvarIdx <= mCol.Count Then
    Set mCurr = mCol.Item(mvarIdx)
    mvarEOF = False
    MoveFirst = True
  Else
    Set mCurr = New GEMS_Name
    mvarEOF = True
    mvarIdx = 0
    MoveFirst = False
  End If
 
End Function
Public Function AddName(newName As GEMS_Name)
On Error GoTo errHandler    '- just in case there's already one there
  If Not newName Is Nothing Then
    If Not newName.Name_type_cid = "" Then
      mCol.Add newName, CStr(newName.Name_type_cid)
    Else
      mCol.Add newName
    End If
    AddName = True
    Set mCurr = newName
    mvarEOF = False
    mvarIdx = mCol.Count
  End If
  
errHandler:


End Function

Public Function Add(Name_type_cid As String, Last_name_soundex As String, _
                First_name_soundex As String, Degree As String, _
                Title As String, Middle_name2 As String, Suffix As String, _
                Middle_name As String, First_name As String, Last_name As String, _
                UpdateFlag As String, Optional sKey As String) As GEMS_Name
                
    'create a new object
    Dim oGems_Name As GEMS_Name
    Set oGems_Name = New GEMS_Name


    'set the properties passed into the method
    With oGems_Name
        .Name_type_cid = Name_type_cid
        .Last_name_soundex = Last_name_soundex
        .First_name_soundex = First_name_soundex
        .Degree = Degree
        .Title = Title
        .Suffix = Suffix
        .Middle_name = Middle_name
        .Middle_name2 = Middle_name2
        .First_name = First_name
        .Last_name = Last_name
        .UpdateFlag = UpdateFlag
    End With
    If Len(sKey) = 0 Then
        mCol.Add oGems_Name
    Else
        mCol.Add oGems_Name, sKey
    End If

    'return the object created
    Set Add = oGems_Name
    Set oGems_Name = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As GEMS_Name
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


