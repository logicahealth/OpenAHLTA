VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GEMS_PickListCtrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mConn As CHCSII_CONN.Conn
Private moDAS As ICHCSII_DAS
Private rs As ADODB.Recordset
Private Const DEFAULT_CONTEXT As String = "2000"

Private sContext As String
Private sDomain As String
Private bPrefOnlyFlag As Boolean
Private bNeeded As Boolean
Private sRelationship As String     '- relationship ncid
    
Private cPickListItems As New Collection
Private iPickListIndex As Integer

Public Property Get Relationship() As String:            Relationship = sRelationship:  End Property
Public Property Let Relationship(ByVal RHV As String):   sRelationship = RHV:           End Property

Public Property Get ClassHandle() As Long
    'ClassHandle =
End Property
Public Property Get Context() As String
    Context = sContext
    If Trim(sContext) = "" Then
        sContext = DEFAULT_CONTEXT
    End If
End Property
Public Property Let Context(ByVal sContextValue As String)
    sContext = sContextValue
End Property
Private Sub Class_Terminate()
    Set mConn = Nothing
    Set moDAS = Nothing
    Set rs = Nothing
End Sub
Public Property Get Domain() As String
    Domain = sDomain
End Property
Public Property Let Domain(ByVal sDomainValue As String)
    sDomain = sDomainValue
End Property
Public Property Get ListItem(itemidx As Long) As Object
    If itemidx <= 0 Then
       Set ListItem = Nothing
    End If
    Set ListItem = cPickListItems(itemidx)
End Property
Public Property Get ListItemCnt() As Integer
    ListItemCnt = cPickListItems.Count
End Property
Public Sub Load()
    Dim oItem As clsPickListItem
    Dim lCount As Integer
    
    If sDomain = "" Then
       Exit Sub
    End If
  
    For lCount = 1 To cPickListItems.Count
         cPickListItems.Remove 1
    Next
    
    If moDAS Is Nothing Then
        Set mConn = New CHCSII_CONN.Conn
        Set moDAS = mConn.CHCSII_DAS(Auto)
    End If
    
    Set rs = New ADODB.Recordset
    Set rs = moDAS.ExecuteSQL("Select Description, NCID from NCIDs where DomainNCID = '" & Domain & "'")
    If rs.RecordCount = 0 Then
       Exit Sub
    End If
    
    rs.MoveFirst
    Do Until rs.EOF
       Set oItem = New clsPickListItem
       oItem.Representation = rs.Fields("Description")
       oItem.NCID = CStr(rs.Fields("NCID"))
       cPickListItems.Add oItem, oItem.NCID
       rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
    
End Sub
Public Property Get PrefOnlyFlag() As Boolean
    PrefOnlyFlag = bPrefOnlyFlag
End Property
Public Property Let PrefOnlyFlag(ByVal bPrefOnlyFlagValue As Boolean)
    bPrefOnlyFlag = bPrefOnlyFlagValue
End Property
Public Property Get Needed() As Boolean
    Needed = bNeeded
End Property
Public Property Let Needed(ByVal bNeededValue As Boolean)
    bNeeded = bNeededValue
End Property

