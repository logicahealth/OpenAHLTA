VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colSelectedTeeth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const WHITE = "WhiteOnly"
Private Const YELLOW = "YellowOnly"
Private Const YELLOW2 = "FillYellow"
Private Const BLUE = "BlueOnly"
Private Const GREY = "GreyOnly"

Private Const WHITE_HIDE = "NoFill"
Private Const YELLOW_HIDE = "YellowYellow"
Private Const BLUE_HIDE = "BlueBlue"
Private Const GREY_HIDE = "GreyGrey"

Private m_Items As Collection
Public m_gTeethDoc As Object
Public m_gTeethWin As Object
' Add an item at its correct position.
Public Function Add(ByVal new_item As clsgcTooth) As clsgcTooth

Dim i As Integer
Dim oRect As Object

    'prepare to change rect to yellow by the mere fact that we are adding to collection
    Set oRect = m_gTeethDoc.getElementById(CStr(new_item.ToothID))
    
    'Tracker 24234
    If oRect Is Nothing Then
        Exit Function
    End If
    
    Call oRect.setAttribute("class", YELLOW)
    
    ' See where the item belongs.
    For i = 1 To m_Items.Count
        If m_Items(i).ToothID >= new_item.ToothID Then Exit For
    Next i

    ' Insert the item.
    If i > m_Items.Count Then
        ' Add at the end.
        m_Items.Add new_item
    Else
        ' Add at the right position.
        m_Items.Add new_item, , i
    End If
    
    Set Add = new_item
    
End Function
Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = m_Items.Count
End Property
Public Sub removeAll()
Dim i As Integer

    For i = 1 To m_Items.Count
       Remove (1)
    Next i
    
End Sub
Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)
Dim oRect As Object

    'prepare to change rect to yellow by the mere fact that we are adding to collection
    Set oRect = m_gTeethDoc.getElementById(CStr(m_Items.Item(vntIndexKey).ToothID))
    Call oRect.setAttribute("class", WHITE)
 

    m_Items.Remove vntIndexKey
    
End Sub
Public Property Get Item(vntIndexKey As Variant) As clsgcTooth
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = m_Items(vntIndexKey)
End Property
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = m_Items.[_NewEnum]
End Property

' Create the new collection.
Private Sub Class_Initialize()
    Set m_Items = New Collection
End Sub

Private Sub Class_Terminate()
    
    Set m_Items = Nothing
    Set m_gTeethDoc = Nothing
    Set m_gTeethWin = Nothing

End Sub
