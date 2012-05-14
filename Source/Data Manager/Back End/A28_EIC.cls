VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "A28_EIC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private m_dicDom As Dictionary
Private m_sPath As String

Public Function GetPatientDict() As Dictionary
  Set GetPatientDict = m_dicDom
End Function

Public Function Check4New() As Boolean
Dim oSQL As ICHCSII_SQl.ISqlOpsEx
Dim sSQL As String

On Error Resume Next

  Set oSQL = oDatabase.CHCSII_SQLOPS(Auto)  'GEMS_SQL.SqlOps
  
  sSQL = "Select count(*) from TMIP_A28"
  
  oSQL.Execute sSQL
  
  If oSQL(0) > 0 Then
    If Err.Number = 0 Then
      Check4New = True
    End If
  End If
  
  Set oSQL = Nothing
  
End Function

'FOR TEXT LOAD
'Public Function LoadA28(sPath As String) As Boolean
'Dim nReturn As Long
'Dim nCnt As Long
'
'    On Error GoTo ErrHandler
'
'    Set m_dicDom = New Dictionary
'
'    m_sPath = sPath
'
'    nReturn = QueryA28(sPath)
'
'    For nCnt = 1 To m_dicDom.Count
'        LoadPatient m_dicDom.Item(nCnt)
'    Next nCnt
'
'Exit Function
'
'ErrHandler:
'    gobjShared.ShowVBError Err.Number, Err.Description, "MC4_EIC:ProcessMC4", "DMBE", vbExclamation
'End Function

Public Function LoadA28() As Boolean
  Dim nReturn As Long
  Dim nCnt As Long

    On Error GoTo ErrHandler

    Set m_dicDom = New Dictionary

    nReturn = QueryA28DB

    For nCnt = 1 To m_dicDom.Count
        LoadPatient m_dicDom.Item(nCnt)
    Next nCnt

Exit Function

ErrHandler:
    gobjShared.ShowVBError Err.Number, Err.Description, "MC4_EIC:ProcessMC4", "DMBE", vbExclamation
End Function

Private Function QueryA28DB() As DMErrorCodes
  Dim oSQL As ICHCSII_SQl.ISqlOpsEx
  Dim oSQLAudit As ICHCSII_SQl.ISqlOpsEx
  Dim oSQLError As ICHCSII_SQl.ISqlOpsEx
  Dim oDas As ICHCSII_DAS
  Dim oFSO As Scripting.FileSystemObject
  Dim oDom As MSXML2.DOMDocument30
  Dim oCOl As Collection
  Dim oMC4XML As MC4_Ops
  Dim sTempXML As String
  Dim nReturn As Long
  Dim nCnt As Long
  Dim sErrorMsg As String
  Dim sSQL As String
  
  On Error GoTo ErrHandler
  
    QueryA28DB = ERROR_UNHANDLED_SYSTEM_ERROR
   
    Set oSQL = oDatabase.CHCSII_SQLOPS(Auto) 'New GEMS_SQL.SqlOps
    Set oSQLError = oDatabase.CHCSII_SQLOPS(Auto) 'New GEMS_SQL.SqlOps
    Set oSQLAudit = oDatabase.CHCSII_SQLOPS(Auto) 'New GEMS_SQL.SqlOps
    Set oDas = oDatabase.CHCSII_DAS(Auto) 'New GEMS_DAS
    
     
    sSQL = "SELECT * FROM TMIP_A28"
          
    oSQL.Execute (sSQL)
    
    If oSQL.RecordCount > 0 Then
      oSQLError.Execute ("SELECT * FROM TMIP_INVALID_A28 Where 0 = 1")
      oSQLAudit.Execute ("SELECT * FROM TMIP_A28_AUDIT Where 0 = 1")
      oSQL.MoveFirst
      Do While Not oSQL.EOF
        nCnt = nCnt + 1
          
        sTempXML = oSQL("XML_MESSAGE_BODY")
        
        If oDom Is Nothing Then
          Set oDom = New MSXML2.DOMDocument30
        End If
          
        oDom.async = False
        oDom.validateOnParse = False
        oDom.resolveExternals = False
          
        If oDom.loadXML(sTempXML) Then
          If InStr(1, sTempXML, "<?xml", vbTextCompare) = 0 Then
            sTempXML = "<?xml version='1.0'?>" & oDom.xml
          Else
            sTempXML = oDom.xml
          End If
            
          If Not m_dicDom.Exists(nCnt) Then
            m_dicDom.Add nCnt, sTempXML
          End If
          
          oSQLAudit.AddNew
          oSQLAudit("CREATEDON") = oSQL("CREATEDON")
          oSQLAudit("FILENAME") = oSQL("FILENAME")
          oSQLAudit("XML_MESSAGE_BODY") = oSQL("XML_MESSAGE_BODY")
          oSQLAudit("UPDATEDON") = Now
          oSQLAudit("UPDATEDBY") = gobjShared.CurrentUserNCID
          oSQLAudit.Update
          
        Else
          oSQLError.AddNew
          oSQLError("CREATEDON") = oSQL("CREATEDON")
          oSQLError("FILENAME") = oSQL("FILENAME")
          oSQLError("XML_MESSAGE_BODY") = oSQL("XML_MESSAGE_BODY")
          oSQLError.Update

          sErrorMsg = sErrorMsg & "Error reading file- """ & oSQL("FILENAME") & """ XML" & vbCrLf
        End If
'        Set oCOl = New Collection
'        sSQL = "Delete From TMIP_A28 Where FILENAME = '" & oSQL("FILENAME") & "'"
'        oCOl.Add sSQL
'        oSQLDelete.ExecuteMulti oCOl
        oDas.ExecuteSQL "Delete From TMIP_A28 Where FILENAME = '" & oSQL("FILENAME") & "'"
        oSQL.MoveNext
     Loop
   End If
            
   If Len(sErrorMsg) > 0 Then
     MsgBox sErrorMsg, vbInformation, "Error reading some Demographic Files"
   End If
   
   QueryA28DB = ERROR_SUCCESS
   
DOExit:
    Set oSQL = Nothing
    Set oSQLError = Nothing
    Set oSQLAudit = Nothing
    Set oDas = Nothing
    Set oDom = Nothing
    
Exit Function

ErrHandler:
    
   If gobjShared Is Nothing Then
      Set gobjShared = New CWShared
   End If
   gobjShared.ShowVBError Err.Number, Err.Description, "QueryA28DB Function", "A28_EIC", vbExclamation
   Err.Clear
   GoTo DOExit
   Resume
End Function

Private Function LoadPatient(sXML As String) As Long
Dim colRSObjects As Collection
Dim objEICTransfer As DMBE.HL7XMLParser
Dim objPatient As MMMHISPatient.Patient
Dim nReturn As Long
Dim enmPatientSearchResult As eSearchResults

  Set objEICTransfer = New DMBE.HL7XMLParser
  
  nReturn = objEICTransfer.GetRecordsetsFromEIC(colRSObjects, CLng(0), eDM_Demographics, sXML, TMIP_Message)

  Set objPatient = New MMMHISPatient.Patient
  Call objPatient.Validation_Criteria_Eic(colRSObjects)

  enmPatientSearchResult = objPatient.ValidatePatient(True)
End Function

Private Sub Class_Initialize()
  If gobjShared Is Nothing Then
    Set gobjShared = New MMMHISShared.CWShared
  End If
End Sub

