VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExamXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const XML_DOCUMENT_ELEMENT_NAME As String = "Exam"

Private mvarDentalEncounterID As Long

Private sPrevExamXMLString As String

Private mvarXMLDoc As MSXML2.DOMDocument30

Public Function Load(DentalEncounterID As Long) As Boolean

'Get the xml string from the db
Dim rs As Recordset
Dim sSQL As String
Dim arXML() As String

    Load = True
        
    If mvarDentalEncounterID = DentalEncounterID Then
        'Have already loaded it for this encounter so exit
        Exit Function
    End If
    
    sSQL = "Select Exam_XML from Templ_Exam_XML where Dental_Encounter_ID = " & DentalEncounterID & " order by Exam_ID"
    
    Set rs = goconn.OpenRecordset(sSQL)
        
    ReDim Preserve arXML(rs.RecordCount) As String
    
    Do Until rs.EOF
        arXML(rs.AbsolutePosition) = rs("Exam_XML")
        
        rs.MoveNext
    Loop

    sPrevExamXMLString = Join(arXML)
    
    'Load the xmldoc
    Set mvarXMLDoc = New MSXML2.DOMDocument30
    mvarXMLDoc.async = False
    
    If sPrevExamXMLString = "" Then
        'Create an empty document
        Set mvarXMLDoc.documentElement = mvarXMLDoc.createElement(XML_DOCUMENT_ELEMENT_NAME)
    Else
        'Load the prev xml
        If mvarXMLDoc.loadXML(sPrevExamXMLString) = False Then
            '*** MsgBox Fix - Code_Validator auto change made 3/21/2003
            MsgBxARMd "clsExamXML, Could not parse xml document: " & sPrevExamXMLString & " - " & mvarXMLDoc.parseError.reason
            Load = False
        End If
    End If
    
    mvarDentalEncounterID = DentalEncounterID
    
End Function

Public Function Save() As Boolean

'Save the xml string to the db
Dim sSQL As String
Dim lSeq As Long
Dim sTempXML As String

On Error GoTo ErrorHnd

    Save = True
    
    If sPrevExamXMLString = mvarXMLDoc.xml Then
        Exit Function
    End If
    
    'Delete the existing records
    sSQL = "Delete from Templ_Exam_XML Where Dental_Encounter_ID = " & mvarDentalEncounterID
    goconn.ExecuteSQL sSQL
    
    'Save the xml string to the db
    sTempXML = mvarXMLDoc.xml
    Do Until sTempXML = ""
        lSeq = goconn.GetNextID(esTEMPL_EXAM_XML_SEQ)
        sSQL = "Insert into Templ_Exam_XML (Dental_Encounter_ID, Exam_ID, Exam_XML) Values (" _
            & mvarDentalEncounterID & ", " & lSeq & ", '" & RQ(Left(sTempXML, 2000)) & "')"
        goconn.ExecuteSQL sSQL
        
        sTempXML = Mid(sTempXML, 2000)
    Loop
    
    sPrevExamXMLString = mvarXMLDoc.xml

Exit Function

ErrorHnd:
    If Err.Number = -2147221503 Then
        'Oracle message - no data found on the delete. Not a problem.
        Resume Next
    End If
    

End Function

Public Function TerminateObjects()

    Set mvarXMLDoc = Nothing
    sPrevExamXMLString = ""
    
End Function


Public Property Get XMLDoc() As MSXML2.DOMDocument30
    
    Set XMLDoc = mvarXMLDoc
    
End Property

