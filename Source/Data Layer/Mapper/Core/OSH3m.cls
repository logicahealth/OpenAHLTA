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
Attribute VB_Name = "OSH3m"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Used to get and save configuration information using 3M's OSH objects.  The configuration is cached in an
'ini file on the workstation by the 3M object.
Option Explicit

Private Const NCID_BIN_ITEM_FILE_TYPE As String = "72"

Private mCache As CacheMapper

'Remove temp ini files created by 3M OSH object.
Public Sub CleanUpTempFiles()
Dim Entry As iCacheEntry
Dim Files As iParameters
Dim i As Long

  On Error Resume Next
  
  Set Entry = mCache.Retrieve("OSHTempFiles")
  
  If Entry.Valid Then
    Set Files = Entry.Value
    
    For i = 1 To Files.Count
      Kill Files.Item(i).Value
    Next
  End If
  
End Sub

Public Function DeleteBin(ByVal OwnerNCID As String) As Boolean
Dim OSHBin As Bin
Dim OSH As Storehouse

  On Error GoTo ErrHandler
  
  Set OSH = LoadBins(OwnerNCID)
    
  If Not (OSH Is Nothing) Then
    Set OSHBin = LoadBin(OSH, OwnerNCID)
    
    If Not (OSHBin Is Nothing) Then
      OSHBin.DeleteMe = True
      
      If OSHBin.Save(True) = 0 Then
        Err.Raise 60051, "OSH3m.DeleteBin", "Unable to Delete the Bin for: " & OwnerNCID & "."
      End If
    End If
  End If
    
  DeleteBin = True

  Exit Function
  
ErrHandler:
  If Err.Number >= 60050 Then
    Err.Raise Err.Number, Err.Source, Err.Description
  Else
    Err.Raise Err.Number, "OSH3m.DeleteBin", Err.Description
  End If

End Function

Public Function GetValue(ByRef BinItem As DTOs.OshBinItem, ByVal SessionKey As String) As DTOs.OshBinItem
Dim i As Long
Dim OSHBin As Bin
Dim OshBinItem As BinItem
Dim File As INIFile
Dim Found As Boolean
Dim Params As iParameters
Dim Persist As iObjectPersist
Dim Item As OshBinItem

  On Error GoTo ErrHandler
        
  Set Params = GetCachedBins
  
  'Order of search is define in the loading of Bin objects in RetrieveConfiguration method.
  For i = 1 To Params.Count
    Set OSHBin = Params.Item(i).Value
    BinItem.BinNCID = OSHBin.Owner
    
    Set OshBinItem = LoadBinItem(OSHBin, BinItem.ApplicationNCID)
    
    Set File = Nothing
    Set File = New INIFile
    File.Init OSHBin.Owner & "~" & BinItem.ApplicationNCID & "~" & SessionKey
      
    If Not (OshBinItem Is Nothing) Then
      If OshBinItem.FileLocation = vbNullString Then
        OshBinItem.FileLocation = File.FullFileName
              
        If Not OshBinItem.Load Then
          BinItem.BinNCID = vbNullString
          Err.Raise 60052, "OSH3m.GetValue", "Unable to load the configuration file: " & File.FullFileName & " in Config.GetValue."
        End If
      End If
      
      Set Persist = BinItem
      Set Item = Persist.Copy         'Use copy to return
      
      Found = File.GetINIValue(Item)  'Get value from INI temp file
      
      If Found Then Exit For
    End If
    
  Next
  
  If Item Is Nothing Then
    Set Persist = BinItem
    Set Item = Persist.Copy
    Item.Value = vbNullString         'return null object
  End If
  
  Set GetValue = Item
  
  If Not (File Is Nothing) Then
    CacheFileName File.FullFileName   'cache temp ini file name to delete later
  End If
  
  If Not Found Then Item.BinNCID = vbNullString      'return of null indicates config item not found
  
  Exit Function
  
ErrHandler:
  If Err.Number >= 60050 Or InStr(1, Err.Source, "INIFiles") > 0 Then
    Err.Raise Err.Number, Err.Source, Err.Description
  Else
    Err.Raise Err.Number, "OSH3m.GetValue", Err.Description
  End If

End Function

Public Sub Init(ByRef Cache As DL_Support.CacheMapper)
  Set mCache = Cache
End Sub

Public Sub RetrieveConfiguration(ByRef Parameters As DL_Support.iParameters)
Dim Item As Bin
Dim Params As iParameters
Dim i As Long

  On Error GoTo ErrHandler
  
  Set Params = New ParametersGeneric
  
  'Load Bins Search order used by GetValue method
  For i = 1 To Parameters.Count
    Set Item = InitBin(Parameters.Item(i).Value)
    
    If Not (Item Is Nothing) Then Params.Add Parameters.Item(i).Value, Item
    
    Set Item = Nothing
  Next
  
  mCache.Save "OSH", Params         'Save NCIDs of Bins to search for config values
  
  Exit Sub
  
ErrHandler:
  If Err.Number >= 60050 Then
    Err.Raise Err.Number, Err.Source, Err.Description
  Else
    Err.Raise Err.Number, "OSH3m.RetrieveConfiguration", Err.Description
  End If
  
End Sub

Public Function SaveChanges(ByRef BinItems As DTOs.OshBinItems, ByVal SessionKey As String) As Boolean
Dim OSHBin As Bin
Dim OshBinItem As BinItem
Dim Success As Boolean

  On Error GoTo ErrHandler
  
  If BinItems.Count = 0 Then
    SaveChanges = True
    Exit Function
  End If
  
  Set OSHBin = GetBin(BinItems.Item(1).OwnerNCID)
  
  If OSHBin Is Nothing Then Exit Function
    
  Set OshBinItem = LoadBinItem(OSHBin, BinItems.Item(1).ApplicationNCID)
            
  SetValues BinItems, OSHBin, OshBinItem, SessionKey
          
  If Not (OshBinItem Is Nothing) Then
    
    If OshBinItem.FileChanged Then
      'IMPORTANT: Make sure the OSH OLE "thinks" the file needs to be saved.  This
      '           is accomplished by changing a "property" of the OLE.  Any time you
      '           change an OLE property, it "marks" the file as "dirty", and will
      '           save out the changes on the next "Save".
      With OshBinItem
        .FileChanged = False
        .FileChanged = True
      
        Success = .Save(True)
      End With
      
      SaveChanges = Success
    End If
  Else
    Err.Raise 60054, "OSH3m.SaveChanges", "BinItem does not exist for Application NCID: " & BinItems.Item(1).ApplicationNCID & " in the Bin owned by: " & BinItems.Item(1).OwnerNCID & " in SaveChanges.  The changes will not be saved."
  End If
  
  Exit Function
  
ErrHandler:
  If Err.Number >= 60050 Then
    Err.Raise Err.Number, Err.Source, Err.Description
  Else
    Err.Raise Err.Number, "OSH3m.SaveChanges", Err.Description
  End If
  
End Function

Private Sub CacheFileName(ByVal FileName As String)
Dim Entry As iCacheEntry
Dim Files As iParameters

  Set Entry = mCache.Retrieve("OSHTempFiles")
  
  If Not Entry.Valid Then
    Set Files = New ParametersGeneric
  Else
    Set Files = Entry.Value
  End If
  
  If Not Files.Exists(FileName) Then
    Files.Add FileName, FileName
    mCache.Save "OSHTempFiles", Files
  End If
  
End Sub

Private Function GetBin(ByVal OwnerNCID As String) As Bin
Dim Entry As iCacheEntry
Dim Params As iParameters

  On Error Resume Next
  
  Set Entry = mCache.Retrieve("OSH")
  
  If Entry.Valid Then
    Set Params = Entry.Value
    
    Set GetBin = Params.Item(OwnerNCID).Value
  End If
  
End Function

Private Function GetCachedBins() As iParameters
Dim Entry As iCacheEntry

  Set Entry = mCache.Retrieve("OSH")
  Set GetCachedBins = Entry.Value
  
End Function

'Load a Bin object for a given OwnerNCID.
Private Function InitBin(ByVal BinOwnerNCID As String) As Bin
Dim OSH As Storehouse
Dim OSHBin As Bin
Dim OshBinItem As BinItem

  On Error GoTo ErrHandler
  
  If Len(BinOwnerNCID) = 0 Then Exit Function
    
  Set OSH = LoadBins(BinOwnerNCID)
  Set OSHBin = LoadBin(OSH, BinOwnerNCID)
            
  If OSHBin Is Nothing Then
    Set OSHBin = New Bin
                
    With OSHBin
      .Name = BinOwnerNCID
      .Owner = BinOwnerNCID
      .Class = "cw_configuration"
    End With
                
    'save the new "Bin" to the CDR, this is needed to enable adding/modifying "BinItems" within this "Bin"---
    If OSHBin.Save(True) = 0 Then
      Err.Raise 60052, "MapperCDR.InitBin", "Unable to Create initial configuration Bin for BinNCID: " & BinOwnerNCID & " in Config.mbLoadBinInitBinItems."
    End If
  Else
    'Since you only want to "load" files as needed, the "FileLocation"
    'property is used to tell if a "BinItem's" file has been loaded or
    'not.  If it is NULL, it has NOT been loaded yet.
    For Each OshBinItem In OSHBin.BinItems
      With OshBinItem
        .FileChanged = False              'file has not be modified yet
        
        'Do not use vbNullString for next line because it doesn't work with the 3M object on NT.  Why?
        .FileLocation = ""                'file has not be "loaded" or copied to the PC yet
      End With
    Next
  End If

  Set InitBin = OSHBin
  
  Exit Function
  
ErrHandler:
  If Err.Number >= 60050 Then
    Err.Raise Err.Number, Err.Source, Err.Description
  Else
    Err.Raise Err.Number, "MapperCDR.InitBin", Err.Description
  End If
  
End Function

Private Function LoadBin(ByRef OSH As Storehouse, ByVal BinOwnerNCID As String) As Bin
  On Error Resume Next
  
  Set LoadBin = OSH.Bins(BinOwnerNCID)
End Function

Private Function LoadBinItem(ByRef OSHBin As Bin, ByVal ApplicationNCID As String) As BinItem
  On Error Resume Next
  
  Set LoadBinItem = OSHBin.BinItems(ApplicationNCID)
End Function

Private Function LoadBins(ByVal BinOwnerNCID As String) As Storehouse
Dim OSH As Storehouse

  On Error GoTo ErrHandler
  
  Set OSH = New Storehouse
    
  With OSH
    .SearchName = BinOwnerNCID
    .SearchOwner = BinOwnerNCID
    .SearchClass = "cw_configuration"
  End With

  If OSH.Load(True) = 0 Then
    Err.Raise 60050, "OSH3m.LoadBins", "Unable to Load the Bin for: " & BinOwnerNCID & ".  This Bin will not be deleted."
  End If
    
  Set LoadBins = OSH
  
  Exit Function
  
ErrHandler:
  If Err.Number >= 60050 Then
    Err.Raise Err.Number, Err.Source, Err.Description
  Else
    Err.Raise Err.Number, "OSH3m.LoadBins", Err.Description
  End If
  Exit Function
  Resume
End Function

Private Sub SetValues(ByRef BinItems As DTOs.OshBinItems, ByRef OSHBin As Bin, ByRef OshBinItem As BinItem, _
                      ByVal SessionKey As String)
Dim File As INIFile
Dim Success As Boolean
Dim BinItem As DTOs.OshBinItem
Dim i As Long

  On Error GoTo ErrHandler

  If BinItems.Count = 0 Then Exit Sub
    
  Set File = New INIFile
  File.Init OSHBin.Owner & "~" & BinItems.Item(1).ApplicationNCID & "~" & SessionKey
      
  If Not (OshBinItem Is Nothing) Then
    'OSH BinItem exists
    If OshBinItem.FileLocation = vbNullString Then
      OshBinItem.FileLocation = File.FullFileName
            
      If Not OshBinItem.Load Then
        Err.Raise 60052, "OSH3m.SetValue", "Unable to load the configuration file: " & File.FullFileName & " in Config.SetValue."
      End If
    End If
  Else
    'New OSH BinItem needed
    Set OshBinItem = New BinItem
    
    With OshBinItem
      .Type = NCID_BIN_ITEM_FILE_TYPE
      .Name = BinItems.Item(1).ApplicationNCID
      .BinName = OSHBin.Name
      .BinClass = OSHBin.Class
      .FileChanged = False
      .FileLocation = File.FullFileName
    End With
        
    OSHBin.BinItems.Add OshBinItem, -1, OshBinItem.Name   'Add new OSH Bin Item
  End If
      
  'Set all of the value in the INI file
  For i = 1 To BinItems.Count
    Set BinItem = BinItems.Item(i)
    
    With BinItem
      If .Delete Then
        Success = File.SetINIValue(.Section, .Key, vbNullString)   'Save change to local file
      Else
        Success = File.SetINIValue(.Section, .Key, .Value)          'Save change to local file
      End If
    End With
    
    If Success Then OshBinItem.FileChanged = True
  Next
  
  If Not OshBinItem.FileChanged Then
    Err.Raise 60057, "OSH3m.SetValues", "Unable to write changes for the configuration file " & File.FullFileName & " in SetValues."
  End If
  
  Exit Sub
  
ErrHandler:
  If Err.Number >= 60050 Or InStr(1, Err.Source, "INIFiles") > 0 Then
    Err.Raise Err.Number, Err.Source, Err.Description
  Else
    Err.Raise Err.Number, "OSH3m.SetValues", Err.Description
  End If
  
End Sub




