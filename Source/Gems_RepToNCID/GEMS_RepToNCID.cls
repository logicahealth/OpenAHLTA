VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GEMS_RepToNCID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"CHCSII_Demographic"
Attribute VB_Ext_KEY = "Member0" ,"CHCSII_Demographic"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mCol As Collection
Private mvarRepresentation As String
Private mvarNeeded As Long
Private mvarContext As String
Private mvarDomain As String
Private mConn As CHCSII_CONN.Conn
Private moDAS As ICHCSII_DAS
Private rs As ADODB.Recordset
Public Property Get Representation() As String:          Representation = mvarRepresentation: End Property
Public Property Let Representation(ByVal RHV As String): mvarRepresentation = RHV:            End Property

Public Property Get Needed() As Long:                    Needed = mvarNeeded:                 End Property
Public Property Let Needed(ByVal RHV As Long):           mvarNeeded = RHV:                    End Property

Public Property Get Context() As String:                 Context = mvarContext:               End Property
Public Property Let Context(ByVal RHV As String):        mvarContext = RHV:                   End Property

Public Property Get Domain() As String:                  Domain = mvarDomain:                 End Property
Public Property Let Domain(ByVal RHV As String):         mvarDomain = RHV:                    End Property


Public Function SaveCache()
End Function

Public Function Load()

    Dim sql As String
    
    If moDAS Is Nothing Then
        Set mConn = New CHCSII_CONN.Conn
        Set moDAS = mConn.CHCSII_DAS(Auto)
     End If
       
    Set rs = New ADODB.Recordset
    Set rs = moDAS.ExecuteSQL("Select NCID from NCIDs where upper(description)  = '" & UCase(mvarRepresentation) & "' and " _
                & "domainncid = " & mvarDomain)
       
    If rs.EOF Then
     Set rs = Nothing
     Exit Function
    End If
    
    Do While Not rs.EOF
      mCol.Add "" & rs("NCID")
      rs.MoveNext
    Loop
    
    rs.Close

End Function


Public Property Get Item(vntIndexKey As Variant) As String
Attribute Item.VB_UserMemId = 0
   Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
  Set mCol = New Collection

End Sub


Private Sub Class_Terminate()
  Set mCol = Nothing
  Set moDAS = Nothing
  Set mConn = Nothing
  Set rs = Nothing
    
End Sub

