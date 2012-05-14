VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OSHCustom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Used to interface with the OSH_INI_SETTINGS table for configuration information.
Option Explicit

Private mConn As CHCSII_CONN.Conn
Private mCache As CacheMapper

Private mUseCache As Boolean                'True - cache all OshBinItem returned by GetValue method

Private Sub AddCacheItem(ByRef Item As OshBinItem)
Dim Entry As iCacheEntry
Dim Params As iParameters

  On Error Resume Next
  
  Set Entry = mCache.Retrieve("OSH Items")
  
  If Not Entry.Valid Then
    Set Params = New ParametersGeneric
  Else
    Set Params = Entry.Value
  End If
  
  If Not Params.Exists(Item.ApplicationNCID & "_" & Item.OwnerNCID & "_" & Item.Section & "_" & Item.Key) Then
    Params.Add Item.ApplicationNCID & "_" & Item.OwnerNCID & "_" & Item.Section & "_" & Item.Key, Item
  
    mCache.Save "OSH Items", Params
  End If
  
End Sub

Private Sub DeleteOSHSetting(ByRef Item As OshBinItem)
Dim SQLStmt As String

Const SQL_DELETE As String = "DELETE FROM OSH_INI_Settings WHERE " & _
                             "Application_NCID = %AppNCID AND " & _
                             "apOwnerNCID = %OwnerNCID AND " & _
                             "lower(apSection) = '%Sec' AND " & _
                             "lower(apKey) = '%Key'"

  On Error GoTo ErrHandler
  
  SQLStmt = Replace(SQL_DELETE, "%AppNCID", Item.ApplicationNCID)
  SQLStmt = Replace(SQLStmt, "%OwnerNCID", Item.OwnerNCID)
  SQLStmt = Replace(SQLStmt, "%Sec", LCase$(Item.Section)) 'SCR-41051
  SQLStmt = Replace(SQLStmt, "%Key", LCase$(Item.Key))
  
  mConn.CHCSII_DAS.ExecuteSQL SQLStmt

  Exit Sub
  
ErrHandler:
  Err.Raise Err.Number, "OSHCustom.DeleteOSHSetting", Err.Description
  
End Sub

Private Function GetCachedNCIDs() As iParameters
Dim Entry As iCacheEntry

  Set Entry = mCache.Retrieve("OSH")
  Set GetCachedNCIDs = Entry.Value
  
End Function

Public Function GetValue(ByRef BinItem As DTOs.OshBinItem) As DTOs.OshBinItem
Dim Found As Boolean
Dim i As Long
Dim Params As iParameters
Dim Item As OshBinItem
Dim Persist As iObjectPersist

  On Error GoTo ErrHandler
  
  If Len(BinItem.ApplicationNCID) < 1 Or Len(BinItem.Section) < 1 Or Len(BinItem.Key) < 1 Then Exit Function
      
  Set Persist = BinItem
  Set Item = Persist.Copy
  
  Item.BinNCID = vbNullString
  
  Set Params = GetCachedNCIDs
  
  For i = 1 To Params.Count
    If Len(Params.Item(i).Value) > 0 Then
      Item.OwnerNCID = Params.Item(i).Value
      
      If mUseCache Then Found = LoadCachedItem(Item)
      If Not Found Then Found = LoadValueFromDB(Item, Params.Item(i).Value = 0)
      
      If Found Then
        Item.BinNCID = Item.OwnerNCID                     'Indicator that Item was found
                
        If mUseCache Then AddCacheItem Item
        Exit For
      End If
    End If
  Next

  Set GetValue = Item

  Exit Function
  
ErrHandler:
  Err.Raise Err.Number, "OSHCustom.GetValue", Err.Description

End Function

Public Sub Init(ByRef Conn As CHCSII_CONN.Conn, ByRef Cache As DL_Support.CacheMapper, Optional ByVal CacheConfigItems As Boolean = False)
  Set mConn = Conn
  Set mCache = Cache
  
  mUseCache = CacheConfigItems
End Sub

Private Sub InsertOSHItem(ByRef Item As OshBinItem)
Dim SQLStmt As String
      
Const SQL_SET As String = "INSERT INTO OSH_INI_Settings (" & _
                                  "Application_NCID, " & _
                                  "apOwnerNCID, " & _
                                  "apSection, " & _
                                  "apKey, " & _
                                  "apValue) VALUES (" & _
                                  "'%AppNCID', '%OwnerNCID', '%Sec', '%Key', %Val)"
      
  On Error GoTo ErrHandler
      
  SQLStmt = Replace(SQL_SET, "%AppNCID", Item.ApplicationNCID)
  SQLStmt = Replace(SQLStmt, "%OwnerNCID", Item.OwnerNCID)
  SQLStmt = Replace(SQLStmt, "%Sec", LCase$(Item.Section))
  SQLStmt = Replace(SQLStmt, "%Key", LCase$(Item.Key))
  SQLStmt = Replace(SQLStmt, "%Val", mConn.SQLQuote(Item.Value))
  
  mConn.CHCSII_DAS.ExecuteSQL SQLStmt

  Exit Sub
  
ErrHandler:
  Err.Raise Err.Number, "OSHCustom.InsertOSHItem", Err.Description

End Sub

Private Function LoadCachedItem(ByRef Item As OshBinItem) As Boolean
Dim Entry As iCacheEntry
Dim Params As iParameters
Dim CachedItem As OshBinItem

  On Error Resume Next
  Set Entry = mCache.Retrieve("OSH Items")
  
  If Not Entry.Valid Then Exit Function
  
  Set Params = Entry.Value
  
  Set CachedItem = Params.Item(Item.ApplicationNCID & "_" & Item.OwnerNCID & "_" & Item.Section & "_" & Item.Key).Value

  If Not (CachedItem Is Nothing) Then
    Item.Value = CachedItem.Value
    
    LoadCachedItem = True
  End If
End Function

'Load config info from database.
Private Function LoadValueFromDB(ByRef Item As OshBinItem, _
                                 Optional ByVal UseNullGet As Boolean = False) As Boolean
Dim SQLStmt As String
Dim rs As Recordset
  
Const SQL_GET As String = "SELECT apValue FROM OSH_INI_Settings WHERE " & _
                                  "Application_NCID = %AppNCID AND " & _
                                  "lower(apSection) = '%ApSec' AND " & _
                                  "lower(apKey) = '%Key' AND " & _
                                  "apOwnerNCID = '%OwnerNCID'"

Const SQL_GET_NULL As String = "SELECT apValue FROM OSH_INI_Settings WHERE " & _
                                       "Application_NCID = %AppNCID AND " & _
                                       "lower(apSection) = '%ApSec' AND " & _
                                       "lower(apKey) = '%Key' AND " & _
                                       "apOwnerNCID is null"

  On Error GoTo ErrHandler
  
  If Not UseNullGet Then
    SQLStmt = Replace(SQL_GET, "%AppNCID", Item.ApplicationNCID)
    SQLStmt = Replace(SQLStmt, "%OwnerNCID", Item.OwnerNCID)
  Else
    'OwnerNCID is null
    SQLStmt = Replace(SQL_GET_NULL, "%AppNCID", Item.ApplicationNCID)
  End If
  
  SQLStmt = Replace(SQLStmt, "%ApSec", LCase$(Item.Section))
  SQLStmt = Replace(SQLStmt, "%Key", LCase$(Item.Key))

  Set rs = mConn.CHCSII_DAS.OpenRecordset(SQLStmt)
  
  If Not rs.EOF And Not rs.BOF Then
    Item.Value = vbNullString & rs!apValue
    
    LoadValueFromDB = True
  Else
    Item.Value = vbNullString
  End If
    
  rs.Close

  Exit Function
  
ErrHandler:
  Err.Raise Err.Number, "OSHCustom.LoadValueFromDB", Err.Description
  
End Function

Public Sub RetrieveConfiguration(ByRef Parameters As DL_Support.iParameters)
Dim Params As iParameters
Dim Persist As iObjectPersist

  On Error GoTo ErrHandler
  
  Set Persist = Parameters
  Set Params = Persist.Copy
  
  Params.Add "Null", 0        'Owner not set (ie. not used)
  
  mCache.Save "OSH", Params
  
  Exit Sub
  
ErrHandler:
  If Err.Number >= 60050 Then
    Err.Raise Err.Number, Err.Source, Err.Description
  Else
    Err.Raise Err.Number, "OSHCustom.RetrieveConfiguration", Err.Description
  End If

End Sub

Public Function SaveChanges(ByRef BinItems As DTOs.OshBinItems) As Boolean
Dim i As Long
Dim Item As OshBinItem

  On Error GoTo ErrHandler
  
  For i = 1 To BinItems.Count
    Set Item = BinItems.Item(i)
    
    DeleteOSHSetting Item
    
    If Not Item.Delete And Len(Item.Value) > 0 Then
      InsertOSHItem Item
    End If
  Next
  
  If mUseCache Then UpdateCachedItems BinItems
  
  SaveChanges = True
  
  Exit Function
  
ErrHandler:
  If InStr(1, Err.Source, "OSH_Item") = 0 Then
    Err.Source = "OSHCustom.SaveChanges"
  End If
  
  Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Sub UpdateCachedItems(ByRef Items As OshBinItems)
Dim Entry As iCacheEntry
Dim Params As iParameters
Dim i As Long
Dim Item As OshBinItem

  On Error Resume Next
  
  Set Entry = mCache.Retrieve("OSH Items")
  
  If Not Entry.Valid Then Exit Sub
  
  Set Params = Entry.Value
  
  For i = 1 To Items.Count
    If Items.Item(i).Delete Then
      Params.Remove Items.Item(i).ApplicationNCID & "_" & Items.Item(i).OwnerNCID & "_" & Items.Item(i).Section & "_" & Items.Item(i).Key
    Else
      Set Item = Params.Item(Items.Item(i).ApplicationNCID & "_" & Items.Item(i).OwnerNCID & "_" & Items.Item(i).Section & "_" & Items.Item(i).Key).Value
      
      If Not (Item Is Nothing) Then
        Item.Value = Items.Item(i).Value
      End If
    End If
  Next
    
End Sub

