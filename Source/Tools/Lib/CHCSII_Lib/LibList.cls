'***************************************************************************************
'
'  Copyright (c) 2007-2012 Northrop Grumman Corporation
'
'  Licensed by Tricare Management Activity under license from the Copyright owner.
'
'  This text file must be included in all Derivative Works of the licensed Source Code.
'
'***************************************************************************************

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mclsLibList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''************************************************************************
'''Module:        mclsLibList/LibList.cls
'''
'''Description:   List object. allows to create and manipulate list of mclsLibItem
'''               objects
'''
'''Modules Required:
'''
'''Note:
'''
'''Author         : Alexander Alberstein
'''Modified       :
'''               :
'''************************************************************************
Option Explicit
'Public P_bMoveToEnd(byval sKey as string) as boolean
''''''''''''''' Constanst Section '''''''''''''''''''''''''''''''''''''''''

Const msMODULE_NAME              As String = "mclsLibList."

''''''''''''''' Varibles Section '''''''''''''''''''''''''''''''''''''''''

Private mlNextKey                As Long

''''''''''''''' Objects Section '''''''''''''''''''''''''''''''''''''''''

Private mcolList                 As Collection
Private mobjItemCurrent          As mclsLibItem
Private mobjItemNew              As mclsLibItem

''''''''''''''' Events Section '''''''''''''''''''''''''''''''''''''''''

Public Event pevBeforeSort(ByRef bCancel As Boolean)
Public Event pevAfterSort()

Public Function P_bSort _
      (Optional ByVal bAscending As Boolean = True, Optional ByVal bByName As Boolean = True) As Boolean
'''********************************************************************************************
'''Called By:   Any
'''
'''Arguments:
''' bAscending
'''   Indicates sort order: Ascending/Descending. Defaults to bAscending
''' bByName
'''   Indicates whether to sort by items name or values
'''
'''Description:   Sort current list based on specified criteria
'''
'''Note:          Loops are replicated for performance reason
'''Warning:
'''Updates Globals:
'''
'''Returns:       If successful
'''                  True
'''               else
'''                  false
'''
'''Programmer:    Alexander Alberstein
'''********************************************************************************************
   Const sPROC_NAME     As String = "P_bSort"
   
   Dim bResult          As Boolean
   Dim objCheck         As mclsLibItem
   Dim objNext          As mclsLibItem
   Dim colSorted        As Collection
   
On Error GoTo ErrHnd
   
20    Set colSorted = New Collection
   
      If bByName And bAscending Then         '''Sort by item name in ascending order
100      Do While mcolList.Count > 0
110         Set objNext = mcolList(1)
130         For Each objCheck In mcolList
140            If UCase$(objCheck.psName) < UCase$(objNext.psName) Then
150               Set objNext = objCheck
               End If
            Next
190         GoSub MoveItem
         Loop
      ElseIf bByName Then                    '''Sort by item name in descending order
200      Do While mcolList.Count > 0
210         Set objNext = mcolList(1)
230         For Each objCheck In mcolList
240            If UCase$(objCheck.psName) > UCase$(objNext.psName) Then
250               Set objNext = objCheck
               End If
            Next
290         GoSub MoveItem
         Loop
      ElseIf bAscending Then                 '''Sort by item value in ascending order
300      Do While mcolList.Count > 0
310         Set objNext = mcolList(1)
330         For Each objCheck In mcolList
340            If objCheck.pvValue < objNext.pvValue Then
350               Set objNext = objCheck
               End If
            Next
390         GoSub MoveItem
         Loop
      Else                                   '''Sort by item value in descending order
400      Do While mcolList.Count > 0
410         Set objNext = mcolList(1)
430         For Each objCheck In mcolList
440            If objCheck.pvValue > objNext.pvValue Then
450               Set objNext = objCheck
               End If
            Next
490         GoSub MoveItem
         Loop
      End If
         
900   Set mcolList = colSorted
   
999   P_bSort = True
   
Exit Function
MoveItem:
500   Call colSorted.Add(objNext, objNext.psKey)
620   Call mcolList.Remove(objNext.psKey)

Return
ErrHnd:
   P_bSort = False
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & Erl)

End Function


Public Sub P_SelectCurrent(ByVal sKey As String)
   Set mobjItemCurrent = mcolList(sKey)
End Sub


Public Function P_sFindByName _
      (ByVal sName As String, Optional ByVal bCaseSensitive As Boolean, Optional ByVal bLike) As String
'''************************************************************************
'''Called By:
'''
'''Arguments:
''' sName
'''   Name search pattern
''' bCaseSensitive
'''   Indicates whether search case should be ignored
''' bLike
'''   Indicates whether to use Like operator or search for exact match
'''
'''Description:   Attempt to find Item matching specified name pattern
'''
'''Note:
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       If successful
'''                  Key value of the matching item
'''               else
'''                  Empty string
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "P_sFindByName"
   
   Dim sResult          As String
   Dim objItem          As mclsLibItem
   
On Error GoTo ErrHnd
   
2  For Each objItem In mcolList
4     If bLike Then
6        If bCaseSensitive Then
8           If objItem.psName Like sName Then
10             sResult = objItem.psKey
            End If
         Else
20          If UCase$(objItem.psName) Like UCase$(sName) Then
22             sResult = objItem.psKey
            End If
         End If
      Else
30       If bCaseSensitive Then
32          If objItem.psName = sName Then
34             sResult = objItem.psKey
            End If
         Else
40          If gobjLibStr.LibStr_bTxtEql(objItem.psName, sName) Then
42             sResult = objItem.psKey
            End If
         End If
      End If
60    If Len(sResult) > 0 Then
62       Exit For
      End If
   Next objItem
   
999   P_sFindByName = sResult

Exit Function
ErrHnd:
   P_sFindByName = ""
   Call gobjLibApp.LibApp_Error _
        (Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & Erl, True)

End Function

Public Function P_sFindByValue(ByVal vValue As Variant) As String
'''************************************************************************
'''Called By:
'''
'''Arguments:
''' vValue
'''   Value of the Item requested to be found
'''
'''Description:   Attempt to find Item with matching value
'''
'''Note:
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       If successful
'''                  Key value of the matching item
'''               else
'''                  Empty string
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "P_sFindByValue"
   
   Dim sResult          As String
   Dim objItem          As mclsLibItem
   
On Error GoTo ErrHnd
   
2  For Each objItem In mcolList
4     If objItem.pvValue = vValue Then
10       sResult = objItem.psKey
         Exit For
      End If
   Next objItem
   
999   P_sFindByValue = sResult

Exit Function
ErrHnd:
   P_sFindByValue = ""
   Call gobjLibApp.LibApp_Error _
        (Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & Erl, True)

End Function

Public Property Get pbExistKey(ByVal sKey As String) As Boolean
On Error Resume Next
   If mcolList.Item(sKey).pbEditable Then
   End If
   pbExistKey = Err.Number = 0
End Property

Public Property Get pbExistValue(ByVal vVal As Variant) As Boolean
   pbExistValue = Len(P_sFindByValue(vVal)) > 0
   
End Property

Public Property Get plCount() As Long
   Dim lResult As Long
   
   If Not mcolList Is Nothing Then
      lResult = mcolList.Count
   End If
   plCount = lResult
   
End Property

Public Property Get pobjItemCurrent() As mclsLibItem
   Set pobjItemCurrent = mobjItemCurrent
End Property


Public Property Get pobjItem(ByVal vKeyIndex As Variant) As mclsLibItem
   Set pobjItem = mcolList(vKeyIndex)
End Property

Public Property Get pobjItemNew() As mclsLibItem
   Set pobjItemNew = mobjItemNew
End Property

Private Sub Class_Initialize()
   Set mcolList = New Collection
End Sub

Private Sub Class_Terminate()
   Set mobjItemCurrent = Nothing
   Set mobjItemNew = Nothing
   Set mcolList = Nothing
   
End Sub
Public Function P_sAddItem _
(Optional ByVal sKey As String, Optional ByVal sName As String, Optional ByVal vValue As Variant) As String
'''************************************************************************
'''Called By:
'''
'''Arguments:
''' sKey
'''   Optional specified Key value. If empty will be generated.
''    Note: generated key values always starts with plus sign
''' sName
'''   Optional Text representation of the Item
''' vValue
'''   Optional Associated data value
'''
'''Description:   Add new List Item. If Key value is not specified generate new
'''
'''Note:          On successful completion of this procedure pobjItemNew property will
'''               contain reference to the newly added Item
'''
'''Warning:       Specified Key value will be locked on completion of this procedure
'''
'''Updates Globals:
'''
'''Returns:       If successful
'''                  Key value of the added item
'''               else
'''                  Empty string
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "P_sAddItem"

   Dim objItem          As mclsLibItem
   
On Error GoTo ErrHnd
   
10    Set objItem = New mclsLibItem
20    objItem.psName = sName
30    objItem.pvValue = vValue
      '''If key was requested to be generated
      If sKey = "" Then
40       mlNextKey = mlNextKey + 1  '''Generate new key value
50       sKey = mlNextKey
      End If
60    objItem.psKey = sKey
70    Call objItem.P_LockKey     '''Make sure nobody tempers with Item identifier
80    Call mcolList.Add(objItem, objItem.psKey)
90    Set mobjItemNew = objItem
100   Set objItem = Nothing
   
999   P_sAddItem = sKey

Exit Function
ErrHnd:
   P_sAddItem = ""
   Call gobjLibApp.LibApp_Error _
        (Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & Erl, True)

End Function

Public Sub P_Clear()
'''************************************************************************
'''Called By:
'''
'''Arguments:
'''
'''Description:   Clears the list
'''
'''Note:          pctlData must be Listbox or ComboBox
'''
'''Warning:
'''Updates Globals:
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "P_Clear"

On Error GoTo ErrHnd
   
   Set mobjItemNew = Nothing
   Set mobjItemCurrent = Nothing
   Set mcolList = New Collection
   
Exit Sub
ErrHnd:
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "", True)

End Sub

Public Sub P_RemoveItem(Optional ByVal sKey As String, Optional ByVal vValue As Variant)
'''************************************************************************
'''Called By:     Any
'''
'''Arguments:
''' sKey
'''   Key value of the Item requested to be removed. Ignored if empty
''' vValue
'''   Item value of the Item requested to be removed
'''
'''Description:   Removes requested List Item(s)
'''
'''Note:       If Key or value is not specified - Current Item will be removed
'''
'''Warning:    if vValue is specified all matching items will be removed
'''
'''Updates Globals:
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "P_RemoveItem"
   
   Dim objItem          As mclsLibItem
   Dim bCurrent         As Boolean
   
On Error GoTo ErrHnd

      '''If Remove by Key is requested
20    If Len(sKey) > 0 Then
40       GoSub RemoveItem
      Else
         bCurrent = True
      End If
      '''If requested to remove all items matching specified Item value
      If Not IsMissing(vValue) Then
         bCurrent = False
100       For Each objItem In mcolList
120          If objItem.pvValue = vValue Then
140             sKey = objItem.psKey
160             GoSub RemoveItem
            End If
         Next objItem
      End If
      If bCurrent Then
200      If Not mobjItemCurrent Is Nothing Then
220         sKey = mobjItemCurrent.psKey
240         GoSub RemoveItem
         End If
      End If
   
Exit Sub
RemoveItem:
500   Call mcolList.Remove(sKey)

520   If Not mobjItemCurrent Is Nothing Then
540      If mobjItemCurrent.psKey = sKey Then
560         Set mobjItemCurrent = Nothing
         End If
      End If
600   If Not mobjItemNew Is Nothing Then
620      If mobjItemNew.psKey = sKey Then
640         Set mobjItemNew = Nothing
         End If
      End If

Return
ErrHnd:
   Call gobjLibApp.LibApp_Error _
      (Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & Erl, True)

End Sub
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
   Set NewEnum = mcolList.[_NewEnum]
End Function

