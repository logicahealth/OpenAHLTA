VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMultiListItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Collection1"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mColItems As Collection
Private mobjItem As CMultiListItem
Private mlngCnt As Long
Private mlngIndex As Long
Private mstrValues As String
Private mstrInitValues As String
Private mintMaxDescLen As Integer
Private mlngListType As ListType

Private Const const_MODULE_NAME = "CMultiListItems"

Public Enum ListType
    StandardList
    PrimarySecondaryList
End Enum


Public Function Value_String_Display() As String

    Dim objItem As CMultiListItem
    
    On Error GoTo ErrHandler
    
    For Each objItem In mColItems
        If objItem.Checked = True Then
            Value_String_Display = Value_String_Display & ", " & objItem.Description
        End If
    Next
        
    'Remove the first ", "
    If Value_String_Display <> "" Then
        Value_String_Display = Mid(Value_String_Display, Len(", ") + 1)
    End If
    
    Set objItem = Nothing
    
Exit Function
    
ErrHandler:
    mstrValues = Empty

End Function

Private Sub Class_Initialize()
    Set mColItems = New Collection
    Set mobjItem = New CMultiListItem
End Sub

Private Sub Class_Terminate()
    Set mColItems = Nothing
    Set mobjItem = Nothing
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mColItems.[_NewEnum]
End Property

Public Property Get ListType() As ListType
    ListType = mlngListType
End Property
Public Property Let ListType(ByVal lngListType As ListType)
    mlngListType = lngListType
End Property
Public Property Get Count() As Long
    Count = mlngCnt
End Property

Public Property Get Value_String() As String
    Value_String = mstrValues
End Property

Public Property Get Max_Desc_Length() As Integer
    Max_Desc_Length = mintMaxDescLen
End Property
Public Function Item(ByVal Index As Variant) As CMultiListItem
Attribute Item.VB_UserMemId = 0
   Set Item = mColItems.Item(Index)
End Function

'*******************************************************************
'Function CheckDescLength
'Author: KDunne
'Date: 10/18/2001
'Purpose: Checks to see if the Description Length is greater
    'then our Max Length. If it is, truncate the Description.
    'Then, we keep a running tally of our longest Description.
    'When we build our checkboxes, we'll make the width of all of them
    'the value of the longest description, up to the value of the max length.
'*******************************************************************
Private Function CheckDescLength(ByVal strDesc As String) As String
    Dim lngLen As String
    Dim strReturn As String
    
    lngLen = Len(strDesc)
    
    strReturn = strDesc
    
    If lngLen > const_MAX_DESC_LENGTH Then
        strReturn = Left(strDesc, const_MAX_DESC_LENGTH)
        mintMaxDescLen = const_MAX_DESC_LENGTH
    ElseIf lngLen > mintMaxDescLen Then
        mintMaxDescLen = lngLen
    End If
    
    CheckDescLength = strReturn
End Function

'*******************************************************************
'Sub: Add
'Author: KDunne
'Date: 10/08/2001
'Purpose: Adds a CMultiListItem to the CMultiListItems collection class.
'*******************************************************************
Public Sub Add(ByVal strKey As String, ByVal strDescription As String, ByVal blnChecked As Boolean, _
        Optional ByVal blnPrimary As Boolean = False, Optional ByVal blnSecondary As Boolean = False)
    
    Dim lngDescLen As Long
    
    On Error GoTo ErrHandler
        
    'First, we make sure the key and the description field don't contain the
    'Item or value delimeter.
    
    If InStr(strKey, const_VALUE_DELIMETER) Or InStr(strKey, const_ITEM_DELIMETER) _
        Or InStr(strDescription, const_VALUE_DELIMETER) Or InStr(strDescription, const_ITEM_DELIMETER) Then
        Err.Raise const_ERR_INVALID_STRING, const_MODULE_NAME, "Item Key or Item Description contains an invalid character"
    End If
    
    'Next, we Check to see if the Description Length is greater
    'then our Max Length. If it is, truncate the Description.
    strDescription = CheckDescLength(strDescription)
    
    Set mobjItem = New CMultiListItem
    
    With mobjItem
        .Key = strKey
        .Description = strDescription
        .Checked = blnChecked
        If mlngListType = PrimarySecondaryList Then
            .Primary = blnPrimary
            .Secondary = blnSecondary
        End If
        
    End With 'objItem
    
    mColItems.Add mobjItem, strKey
    
    mlngCnt = mColItems.Count
    
    Exit Sub
ErrHandler:
    If Err.Number = 457 Then
        Err.Raise Err.Number, "CMultiListItems.Add", "Key Name for Item Not Unique"
    Else
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Sub

Friend Function Update(ByVal blnChecked As Boolean, Optional ByVal lngIndex As Long = -1, _
        Optional ByVal strKey As String = "", Optional ByVal blnPrimary As Boolean, _
        Optional ByVal blnSecondary As Boolean) As Boolean
    
    Dim blnUpdate As Boolean
    
    On Error GoTo ErrHandler
    
    'Make sure that either an index or a key was supplied.
    If strKey <> "" Then
        Set mobjItem = mColItems(strKey)
    ElseIf lngIndex >= 0 Then
        Set mobjItem = mColItems.Item(lngIndex)
    Else
        'Neither an index nor a key were supplied. Return false
        blnUpdate = False
        GoTo ExitFunc
    End If
    
    If mlngListType = PrimarySecondaryList Then
        mobjItem.Primary = blnPrimary
        mobjItem.Secondary = blnSecondary
    End If
    
    mobjItem.Checked = blnChecked
    
    blnUpdate = True

ExitFunc:
    Update = blnUpdate
    Exit Function

ErrHandler:
    blnUpdate = False
    GoTo ExitFunc
End Function

'*******************************************************************
'Sub: BuildValueString
'Author: KDunne
'Date: 10/16/2001
'Purpose: Loops through Collection and builds a string containing
'         all of the values.
'*******************************************************************
Friend Function BuildValueString() As String
    Dim objItem As CMultiListItem
    Dim strValues As String
    Dim blnReturn As Boolean
    
    On Error GoTo ErrHandler
    
    strValues = Empty
    
    For Each objItem In mColItems
        strValues = strValues & objItem.Key & const_VALUE_DELIMETER & _
            objItem.Description
        If mlngListType = PrimarySecondaryList Then
            strValues = strValues & const_VALUE_DELIMETER & objItem.Primary & _
                const_VALUE_DELIMETER & objItem.Secondary
        Else
            strValues = strValues & const_VALUE_DELIMETER & objItem.Checked
        End If
        
        strValues = strValues & const_ITEM_DELIMETER
    Next
    
    blnReturn = True
    mstrValues = strValues
    
FuncExit:
    BuildValueString = strValues
    Set objItem = Nothing
    
    Exit Function
    
ErrHandler:
    mstrValues = Empty
    blnReturn = False
    GoTo FuncExit
    
End Function
'*******************************************************************
'Function: Load_String
'Author: KDunne
'Date: 10/08/2001
'Returns: boolean. True if successful, False if error occurs
'Purpose: Populates internal collection from | deliminated string
'*******************************************************************
Public Function LoadFromValueString(ByVal strValues As String) As Boolean
 
    Dim objItem As CMultiListItem
    Dim arrItems() As String
    Dim arrValues() As String
    Dim lngItemCnt As Long
    Dim lngValueCnt As Long
    Dim X As Integer
    Dim blnReturn As Boolean
    
    
    On Error GoTo ErrHandler
    
    mstrValues = strValues
    
    arrItems = Split(strValues, const_ITEM_DELIMETER)
    
    lngItemCnt = UBound(arrItems)
    
    Set mColItems = New Collection
    
    For X = 0 To lngItemCnt
        Set objItem = New CMultiListItem
        
        If arrItems(X) <> "" Then
        
            arrValues = Split(arrItems(X), const_VALUE_DELIMETER)
            lngValueCnt = UBound(arrValues)
            With objItem
                If lngValueCnt = 0 Then
                    'The first is the key and Desc
                    .Key = arrValues(0)
                    .Description = CheckDescLength(arrValues(0))
                    .Checked = False
                Else
                    .Key = arrValues(0)
                    .Description = CheckDescLength(arrValues(1))
                    '.Checked = CBool(arrValues(2))
                End If
                If mlngListType = PrimarySecondaryList Then
                    If lngValueCnt > 2 Then
                        .Primary = CBool(arrValues(2))
                        .Secondary = CBool(arrValues(3))
                    End If
                ElseIf lngValueCnt > 1 Then
                    .Checked = CBool(arrValues(2))
                End If
                
                mColItems.Add objItem, objItem.Key
                
            End With
        End If
    
    Next
    
    mlngCnt = mColItems.Count
    
    blnReturn = True
    
    
FuncExit:
    LoadFromValueString = blnReturn
    
    Exit Function
ErrHandler:
    blnReturn = False
    
    Err.Raise Err.Number, Err.Source, Err.Description
    
    GoTo FuncExit
    
End Function
