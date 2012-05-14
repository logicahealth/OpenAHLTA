VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Saves and retrieve CacheEntry objects in memory and the file system.
'CacheEntry items stay in cache until they are retieved and the expiration date is greater than the current date/time.

Option Explicit

Private mConfigFile As String           'Contains default expiration times in minutes by Cache Entry key
Private mConfig As iParameters          'Collection of default expiration times

Private mDataCachePath As String        'File system location to save Cache Entry objects

Private mCacheEntries As Collection     'Cache Entry objects (in memory cache)

Private Sub Class_Initialize()
  Set mCacheEntries = New Collection
End Sub

'Creates a key from the passed in values to be used for a CacheEntry objects' key.
Private Function GetCacheKey(ByVal MapperName As String, ByVal Method As String, Optional ByRef Parameters As iParameters = Nothing) As String
  GetCacheKey = MapperName & Method & GetValueKey(Parameters)
End Function

Private Function GetValueKey(ByRef Parameters As iParameters) As String
Dim i As Long
Dim Key As String

  If Parameters Is Nothing Then Exit Function
  
  For i = 1 To Parameters.Count
    Key = Key & Parameters.Item(i).Value
  Next
  
  GetValueKey = Key
  
End Function

Private Function ConfigFileExist() As Boolean
  On Error Resume Next
  ConfigFileExist = Len(Dir$(mConfigFile)) > 0
End Function

'Init Data Cach file system location and load config information.
Public Sub Init(ByVal ConfigFile As String, ByVal DataCachePath As String)
  
  If Len(DataCachePath) > 0 Then
    mDataCachePath = DataCachePath
    If Right$(mDataCachePath, 1) <> "\" Then mDataCachePath = mDataCachePath & "\"
  End If
  
  If Len(ConfigFile) > 0 Then
    mConfigFile = ConfigFile
    
    LoadConfig
  End If
  
End Sub

Private Sub LoadConfig()
Dim Reader As ReaderXML
Dim Persist As iObjectPersist

  On Error GoTo ErrHandler
  
  If Not ConfigFileExist Then Exit Sub
  
  Set Reader = New ReaderXML
  
  Reader.Load mConfigFile
  
  Set mConfig = New ParametersGeneric
  Set Persist = mConfig
  
  Persist.Restore Reader
  
  Exit Sub
  
ErrHandler:
  'Unable to load config info so don't use
  
End Sub

'Removes a cache entry from the in memory cache if found.
Public Sub Remove(ByVal MapperName As String, ByVal Method As String, Optional ByRef Parameters As iParameters = Nothing)
  On Error Resume Next
  
  mCacheEntries.Remove GetCacheKey(MapperName, Method, Parameters)
  
End Sub

Private Function ReplaceSpecialChar(ByVal Key As String) As String
  
  Key = Replace(Key, "\", "~")
  Key = Replace(Key, "/", "!")
  Key = Replace(Key, ":", "@")
  Key = Replace(Key, "*", "#")
  Key = Replace(Key, "?", "$")
  Key = Replace(Key, """", "%")
  Key = Replace(Key, "<", "^")
  Key = Replace(Key, ">", "&")
  Key = Replace(Key, "|", "-")
  
  ReplaceSpecialChar = Key
  
End Function

'Retrieves a cache entry from the in memory cache if found.
Public Function Retrieve(ByVal MapperName As String, ByVal Method As String, Optional ByRef Parameters As iParameters = Nothing) As iCacheEntry
Dim Entry As iCacheEntry
Dim Key As String

  On Error Resume Next
  
  Key = GetCacheKey(MapperName, Method, Parameters)
  
  Set Entry = mCacheEntries.Item(Key)
  
  If Not (Entry Is Nothing) Then
    If Entry.ExpirationDate <= Now Then
      Entry.Valid = False                   'Cache Entry object has expired
      
      mCacheEntries.Remove Key              'Expired so remove from cache
    End If
  End If
  
  If Entry Is Nothing Then
    Set Entry = New CacheNullEntry          'Cache Entry not found
    Entry.Key = Key
  End If
        
  Set Retrieve = Entry
    
End Function

'Retrieves a cache entry from the file system cache if found.
Public Function RetrieveLocal(ByVal MapperName As String, ByVal Method As String, Optional ByRef Parameters As iParameters = Nothing) As iCacheEntry
Dim Key As String
Dim Entry As iCacheEntry
Dim Reader As iReader
Dim XMLReader As ReaderXML
Dim Value As String

  On Error GoTo ErrHandler
  
  Key = GetCacheKey(MapperName, Method, Parameters)
  Key = ReplaceSpecialChar(Key)
  
  Set Reader = New ReaderXML
  
  Set XMLReader = Reader
  XMLReader.Load mDataCachePath & Key & ".xml"
  
  Set Entry = New CacheEntry
  
  Entry.Key = Reader.ReadProperty("Key")
  
  If Len(Entry.Key) > 0 Then
    With Entry
      .ExpirationDate = CDate(Reader.ReadProperty("ExpirationDate"))
      
      'The 3rd party component is inserting a line feed at regular intervals so we need to strip it out (effects large strings).
      Value = Reader.ReadProperty("Value")
      .Value = Replace(Value, vbLf, vbNullString)
      
      .Valid = Entry.ExpirationDate >= Now          'Cache Entry object has expired
    End With
  Else
    Set Entry = New CacheNullEntry                  'Cache Entry not found
    Entry.Key = Key
  End If
    
  Set RetrieveLocal = Entry
  
  Exit Function
  
ErrHandler:
  Set Entry = New CacheNullEntry
  Entry.Key = Key

End Function

'Save a Cache Entry object to the in memory cache.
'Replaces an existing Cache Entry object if one with the same Key already exists.
Public Function Save(ByVal MapperName As String, ByVal Method As String, ByRef Value As iObjectPersist, Optional ByRef Parameters As iParameters = Nothing, Optional ByVal ExpirationDate As Date = #12/12/9999#) As Boolean
Dim Entry As iCacheEntry
Dim Key As String

  On Error Resume Next
    
  If Value Is Nothing Then Exit Function
  
  Key = GetCacheKey(MapperName, Method, Parameters)
  
  Set Entry = mCacheEntries.Item(Key)

  
  On Error GoTo ErrHandler
  
  If Not (Entry Is Nothing) Then
    Set Entry = Nothing
    mCacheEntries.Remove Key
  End If
  
  Set Entry = New CacheEntry
  
  SetExpirationDate Entry, ExpirationDate, Key
  
  With Entry
    .Key = Key
    .Valid = True
    Set .Value = Value
  End With
  
  mCacheEntries.Add Entry, Key
  
  Save = True
  
  Exit Function
  
ErrHandler:
  'Cache not modified
  
End Function

'Save a Cache Entry object to the file system cache.
'Replaces an existing Cache Entry object if one with the same Key already exists (ie. over writes existing file).
Public Function SaveLocal(ByVal MapperName As String, ByVal Method As String, ByVal SerializeData As String, Optional ByRef Parameters As iParameters = Nothing, Optional ByVal ExpirationDate As Date = #12/12/9999#) As Boolean
Dim Key As String
Dim Writer As iWriter
Dim XMLWriter As WriterXML

  On Error GoTo ErrHandler
  
  If Len(SerializeData) = 0 Then Exit Function
  
  Key = GetCacheKey(MapperName, Method, Parameters)
  Key = ReplaceSpecialChar(Key)
  
  Set Writer = New WriterXML
  
  With Writer
    .InitHeader "CacheEntry", "1.0"
    
    .WriteProperty "Key", Key
    .WriteProperty "ExpirationDate", ExpirationDate
    
    .WriteProperty "Value", SerializeData
  End With
  
  Set XMLWriter = Writer
  XMLWriter.Save mDataCachePath & Key & ".xml"
  
  SaveLocal = True
  
  Exit Function
  
ErrHandler:
  
End Function

'Sets the expiration date for a Cache Entry.
'Checks the Config object for a default value if the Config object was loaded.
Private Sub SetExpirationDate(ByRef Entry As iCacheEntry, ByVal ExpirationDate As Date, ByVal Key As String)
Dim ConfigFileValue As String
Dim IntervalMin As Long

  If ExpirationDate = #12/12/9999# Then
    If Not (mConfig Is Nothing) Then ConfigFileValue = mConfig.Item(Key)
    
    If Len(ConfigFileValue) > 0 Then
      IntervalMin = Val(ConfigFileValue)
    End If
    
    If IntervalMin > 0 Then
      ExpirationDate = DateAdd("n", IntervalMin, Now)
    End If
  End If
  
  Entry.ExpirationDate = ExpirationDate
  
End Sub
