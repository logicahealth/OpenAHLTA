VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GEMS_ConceptCtrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mDAS As ICHCSII_DAS
Private rs As ADODB.Recordset
Private mvarUniqueID As String
Private mUseOracle As Boolean
Private mUseMySQL As Boolean
Public Property Get UniqueID() As String
  UniqueID = mvarUniqueID
End Property
Public Property Let UniqueID(RHV As String)
  mvarUniqueID = RHV
End Property
Public Function PrefRep(RHV As String) As GEMS_PrefSForm
  Dim Sql As String
  Dim oPref As GEMS_PrefSForm
  Set oPref = New GEMS_PrefSForm

  If mvarUniqueID = "" Then
    Exit Function
  End If
   
  On Error GoTo ErrHandler
  
  oPref.Representation = ""
  Set rs = New ADODB.Recordset
   
  Sql = "SELECT Description, DomainNCID From ncids where NCID = '" & mvarUniqueID & "'"
  Set rs = mDAS.OpenRecordset(Sql)
    
  If Not rs.EOF Then
    ' SCR 135084, jq, 7/12/06; clinic lookup was not working
    
    ' start out with the description from the NCID table
    If Not IsNumeric(rs(0)) Then
      oPref.Representation = rs(0)
    End If
    
    ' this was the wrong domainNCID:
    ' ElseIf rs(1) = 99990501 Then                'Domain NCID for Clinic IENs in PGUI database
    
    ' We don't want ElseIf, because we *always* want to look up a clinic if possible
    If rs(1) = 1003 Then  ' this is a clinic entry
      If rs.State = adStateOpen Then rs.Close
      Sql = "select Name from Clinic Where NCID = '" & mvarUniqueID & "'"
      
      Set rs = mDAS.OpenRecordset(Sql)  ' we are re-using rs!!!
    
      If Not rs.EOF Then
          ' found it in the clinic table, so replace the name with our value
          oPref.Representation = rs(0).Value & vbNullString
      End If ' clinic lookup succeeded
    
    End If  ' domainNCID is 1003
    
  End If ' we found an entry in the NCID table
  
  If Len(oPref.Representation) = 0 Then
    If rs.State = adStateOpen Then rs.Close
    
    If mUseMySQL Then
        Sql = "select concat_ws(', ', last_name, first_name) as FullName From mmi_name where unit_number = " & mvarUniqueID
    ElseIf mUseOracle Then
        Sql = "select last_name || ', ' || first_name as FullName From mmi_name where unit_number = " & mvarUniqueID
    Else
        Sql = "select last_name + ', ' + first_name as FullName From mmi_name where unit_number = " & mvarUniqueID
    End If
    
    Set rs = mDAS.OpenRecordset(Sql)
    
    If rs.State = adStateOpen Then
        If Not rs.EOF Then
          If Not IsNull(rs(0).Value) Then  'tfb  changed to handle null values  07-13-2001
              oPref.Representation = rs(0).Value
          Else
              oPref.Representation = vbNullString
              Exit Function
          End If
        End If
    End If
  End If
  
  If Len(oPref.Representation) = 0 Then
    If rs.State = adStateOpen Then rs.Close
    Sql = "select Name from provider where ncid = '" & mvarUniqueID & "'"
    
    Set rs = mDAS.OpenRecordset(Sql)
    
    If Not rs.EOF Then
       If Not IsNull(rs(0).Value) Then  'tfb  changed to handle null values  07-13-2001
           oPref.Representation = rs(0).Value
       Else
           oPref.Representation = vbNullString
       End If
    End If
  End If
  
  rs.Close
  Set rs = Nothing
  Set PrefRep = oPref
  
  Set oPref = Nothing
  
  Exit Function
  
ErrHandler:
  If Err.Number <> 1000 Then
    Err.Raise Err.Number, Err.Source, Err.Description
  Else
    Resume Next
  End If
End Function

Private Sub Class_Initialize()
Dim Connection As CHCSII_CONN.Conn

   Set Connection = New CHCSII_CONN.Conn
   
   Set mDAS = Connection.CHCSII_DAS(Auto)
   
   mUseOracle = Connection.UseLocalDB
   mUseMySQL = Connection.UseMySQL
   
End Sub

Private Sub Class_Terminate()
  Set mDAS = Nothing
  Set rs = Nothing
End Sub
