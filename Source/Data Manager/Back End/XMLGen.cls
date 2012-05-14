VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XMLGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private sOutputFile As String
Private sOutPutXML As String
Private rst_DOM As MSXML2.DOMDocument30
Private rstXSL  As ADODB.Recordset
Private dicXML_Msg As Dictionary

Private oDom As MSXML2.DOMDocument30
Attribute oDom.VB_VarHelpID = -1
Private oRootNode As MSXML2.IXMLDOMElement
Private oNode As MSXML2.IXMLDOMElement

'Not Used yet
Private oNodeLevel1 As MSXML2.IXMLDOMElement
Private oNodeLevel2 As MSXML2.IXMLDOMElement
Private oNodeLevel3 As MSXML2.IXMLDOMElement


Public Enum eMsgName
  eADT_A04 = 1
  eADT_ZA1 = 2
  eVXU_V04 = 3
End Enum

Public Enum eSegAction
  eCopyRST_Load = 0
  eNoLoad = 99
End Enum
Public Property Get ReturnRawXML() As String
  
  'Change the current dom to the temp dom for transformation
  rst_DOM.loadXML oRootNode.xml
  
  ReturnRawXML = rst_DOM.xml
  
End Property

Public Sub AppendNode(ByVal sNodeName As String)

  On Error GoTo Sub_Err
  
  


Sub_Exit:
  On Error Resume Next
  Exit Sub
  
Sub_Err:
  gobjShared.ShowVBError Err.Number, Err.Description, "AppendNode Procedure", "DMBE.XMLGen", vbExclamation
  Resume Sub_Exit
  Resume 'Debug

End Sub

Public Sub CreateRootNodeElement(sElementText As String)
  Set oRootNode = oDom.createElement(sElementText)
End Sub


' Used to add any final touches necessary to the XML
Public Sub FinalizeXML()

  sOutPutXML = "<?xml version=""1.0""?><!DOCTYPE " & oRootNode.nodeName & " SYSTEM ""hl7_v231.dtd"">"
  sOutPutXML = sOutPutXML & oRootNode.xml

End Sub

Private Function GetXSLFile(sMsgName As String, eAction As eSegAction) As String
  
  Dim sXSLName As String
  On Error GoTo Func_Err
  
  

'************************************************
'Bill Fote (21-March-2005)

'If this Function runs without error, the
'   RowsetSchema will NOT appear in the JMEWS File.

'For build 301.5h, I added the next 3 LOC and the
'   function ran without error.  However, the RowsetSchema
'   no longer displayed in the JMEWS file.

'   copy_rst.xsl

'???????????????????????????????????????
'''?  If gobjShared Is Nothing Then    ?
'''?    Set gobjShared = New CWShared  ?
'''?  End If                           ?
'???????????????????????????????????????
'************************************************

  '*****************************************************************
  '#  68485  Open  TH-Datamanager - need to remove schema section
  'Bill Fote - 4-Apr-2005
  If gobjShared Is Nothing Then
    Set gobjShared = New CWShared
  End If
  '*****************************************************************

  If dicXSL Is Nothing Then
    GoTo Func_Exit
  End If
  
  If dicXSL.Exists(sMsgName) Then
    
    sXSLName = dicXSL.Item(sMsgName)
    If Len(sXSLName) <> 0 Then
      GetXSLFile = gobjShared.RTDataPath & "\" & sXSLName
    End If
      
    'Add code to validate version
    
    
  End If


Func_Exit:
  On Error Resume Next
  Exit Function
  
Func_Err:
  gobjShared.ShowVBError Err.Number, Err.Description, "GetXSLFile Procedure", "DMBE.XMLGen", vbExclamation
  Resume Func_Exit
  Resume 'Debug

End Function

Private Sub LoadXSLFiles()
'  Dim rst As ADODB.Recordset
'  Dim sValue As String
'  Dim sSegName As String
'  On Error GoTo Sub_Err
'
'  Set rst = New ADODB.Recordset
  
'  CreateRst eXSL_LookupRst, rst
'
'  rst.AddNew
'  rst.Fields("Seg_Name") = "CopyRST"
'  rst.Fields("XSL_Name") = "CopyRST.XSL"
'  rst.Fields("XSL_Version") = "20020909"
'  rst.Update
'
'  rst.AddNew
'  rst.Fields("Seg_Name") = "PID"
'  rst.Fields("XSL_Name") = "PID.XSL"
'  rst.Fields("XSL_Version") = "20020909"
'  rst.Update
'
'  rst.AddNew
'  rst.Fields("Seg_Name") = "MSH"
'  rst.Fields("XSL_Name") = "MSH.XSL"
'  rst.Fields("XSL_Version") = "20020909"
'  rst.Update
'
'  rst.AddNew
'  rst.Fields("Seg_Name") = "ZD1"
'  rst.Fields("XSL_Name") = "ZD1.XSL"
'  rst.Fields("XSL_Version") = "20020909"
'  rst.Update
'
'  rst.Save "C:\XSL_Lookup.xml", adPersistXML
  
'  Set dicXSL = New Dictionary
'
'  rst.Open gobjShared.RTDataPath & "\" & "XSL_Lookup.xml"
'
'  If Not rst.BOF And Not rst.EOF Then
'    Do Until rst.EOF
'      sSegName = rst.Fields("Seg_Name").Value
'      If Not dicXSL.Exists(sSegName) Then
'        sValue = rst.Fields("XSL_Name") & FieldSepChar & rst.Fields("XSL_Version")
'        dicXSL.Add sSegName, sValue
'      End If
'      rst.MoveNext
'    Loop
'  End If
'
'
'
'Sub_Exit:
'  On Error Resume Next
'  rst.Close
'  Set rst = Nothing
'  Exit Sub
'
'Sub_Err:
'  gobjShared.ShowVBError Err.Number, Err.Description, "LoadXSLFiles Procedure", "DMBE.XMLGen", vbExclamation
'  Resume Sub_Exit
'  Resume 'Debug

End Sub
Public Sub BuildRawXML(ByVal colRSObjects As Collection, sMessageType As String, Optional bIncludeZEG As Boolean = True, Optional lEncNumber As Long = 0)
   
    Dim rst As ADODB.Recordset
    Dim Data As I_RSObjects
    Dim rstCol As Collection
    Dim sTemp As String
    Dim bInpDiagFound As Boolean
    On Error GoTo Func_Err
    
   '**********************************************************
   'SCR 71255
   'Replace the Existing Inpatient Diagnosis Table
   'With the New one.  The new one includes the ICD9 Code.
    bInpDiagFound = False
    For Each Data In colRSObjects
      If UCase$(Data.TableName) = "<~>INPATIENT_DIAGNOSIS" Then
        bInpDiagFound = True
      End If
    Next
    If bInpDiagFound = True Then
      For Each Data In colRSObjects
        If UCase$(Data.TableName) = "INPATIENT_DIAGNOSIS" Then
          Data.TableName = "IGNORETHISTABLE"
        End If
      Next

      For Each Data In colRSObjects
        If UCase$(Data.TableName) = "<~>INPATIENT_DIAGNOSIS" Then
          Data.TableName = "INPATIENT_DIAGNOSIS"
        End If
      Next
    End If
    '*******************************************************
  
    If oDas Is Nothing Then
        Set oDas = oDatabase.CHCSII_DAS(Auto) 'New GEMS_DAS
    End If
  
    CreateRootNodeElement sMessageType
    
    For Each Data In colRSObjects
      Set rst = Data.DBRECORDSET
      'Ignoring the DIAGNOSES table because I'm having to build one to include the ICD9 and text, will create an SCR so Problems will supply this data
      On Error Resume Next
      rst.Filter = ""
      '<< Begin:SCR #30824;   Developer: Brian Mowbray 12/16/2002 11:45 PM
      If rst.RecordCount > 0 Then
      
        rst.MoveFirst
        If lEncNumber > 0 Then
            
            FilterRS "ENCOUNTERNUMBER", lEncNumber, rst
            If Err.Number <> 0 Then
              rst.Filter = ""
              
              Err.Clear
            End If
        Else
          rst.Filter = ""
        End If
      End If
      On Error GoTo Func_Err
      If Not (UCase$(Data.TableName) = "DIAGNOSES" Or UCase$(Data.TableName) = "IGNORETHISTABLE") Then
        PersistRst rst, UCase$(Data.TableName), eCopyRST_Load, , , , bIncludeZEG
      End If
      'This is needed to remove the filter from Enounters, otherwise it will not generate all the JMEWS files
      If Not rst Is Nothing Then
        rst.Filter = ""
      End If
    Next Data
      '>> End: SCR #30824;
      
    'Change the current dom to the temp dom for transformation
    rst_DOM.loadXML oRootNode.xml
    
    Exit Sub
    
Func_Exit:
  On Error Resume Next
  Exit Sub

Func_Err:
   gobjShared.ShowVBError Err.Number, Err.Description, "BuildRawXML Function ", "XMLGen", vbExclamation
   GoTo Func_Exit
   Resume 'Debug
End Sub

Public Function WriteXML(ByVal sPath As String, ByVal sFileName As String) As Boolean
   
  Dim oFS As FileSystemObject
  Dim strFilePath As String
  Dim sXMLData As String
  Dim myDMBE_Ops As DMBE_Ops
  Dim lError As DMErrorCodes
   
  On Error Resume Next
    CheckOrBuildDirs sPath
  On Error GoTo Func_Err
  
  If Right(sPath, 1) <> "\" Then
    sPath = sPath & "\"
  End If
  
  strFilePath = sPath & sFileName
   
  strFilePath = EnsureUniqueFile(strFilePath)
  
  If Not (rst_DOM Is Nothing) Then
    
    Set oFS = New FileSystemObject
    
    
    If oFS.FileExists(strFilePath) Then
       oFS.DeleteFile strFilePath, True  ' just in case of failure last time in clearing file
    End If
    
    Set oFS = Nothing
    
    '*****************************************************************
    'SCR 62595
    'We Must Encrypt the data before writing to the file
    
    sXMLData = rst_DOM.xml
    Set myDMBE_Ops = New DMBE_Ops
    lError = myDMBE_Ops.WriteFileContents(strFilePath, sXMLData, False, True)
    If lError = ERROR_SUCCESS Then
      WriteXML = True
    End If
    
    'rst_DOM.Save strFilePath
    '*****************************************************************
  
  End If
  
  WriteXML = True
  
Func_Exit:
  On Error Resume Next
  Set oFS = Nothing
  Exit Function

Func_Err:
  If gobjShared Is Nothing Then
    Set gobjShared = New CWShared
  End If
  gobjShared.ShowVBError Err.Number, Err.Description, "WriteXML Function", "XMLGen", vbExclamation
  Resume Func_Exit
  Resume 'Debug
End Function

Private Sub LoadMsgDict()
  Dim rst As ADODB.Recordset
  Dim sValue As String
  Dim sMsgName As String
  On Error GoTo Sub_Err
  
  Set rst = New ADODB.Recordset
  
  If gobjShared Is Nothing Then
    Set gobjShared = New CWShared
  End If

'  CreateRst eXML_MsgDict, rst
'
'  rst.AddNew
'  rst.Fields("Msg_Name") = "VXU_V04"
'  rst.Fields("Msg_Struc") = "MSH" & FieldSepChar & "PID" & FieldSepChar & "ZD1" & FieldSepChar & "RXA"
'  rst.Fields("Seg_Name") = "20020909"
'  rst.Update
'
'  rst.Save "C:\XML_MsgDict.xml", adPersistXML
  
  Set dicXML_Msg = New Dictionary
  
  rst.Open gobjShared.RTDataPath & "\" & "XML_MsgDict.xml"
  
  If Not rst.BOF And Not rst.EOF Then
    Do Until rst.EOF
      sMsgName = rst.Fields("Msg_Name").Value
      If Not dicXML_Msg.Exists(sMsgName) Then
        sValue = rst.Fields("Msg_Struc")
        dicXML_Msg.Add sMsgName, sValue
      End If
      rst.MoveNext
    Loop
  End If



Sub_Exit:
  On Error Resume Next
  rst.Close
  Set rst = Nothing
  Exit Sub
  
Sub_Err:
  gobjShared.ShowVBError Err.Number, Err.Description, "LoadMsgDict Procedure", "DMBE.XMLGen", vbExclamation
  Resume Sub_Exit
  Resume 'Debug

End Sub



Private Sub LoadNode(ByVal sXMLString As String, sSegName As String)
  Dim oTmpDOM As MSXML2.DOMDocument30
  
  On Error GoTo Sub_Err
  
  Set oTmpDOM = New MSXML2.DOMDocument30
  
  'This loads the XML string created by the transformation of the XML->XSL to XML
  If oTmpDOM.loadXML(sXMLString) Then
'    If eSegName = eNoLoad Then
'      oRootNode.appendChild oTmpDOM.childNodes(0)
'    Else
      Set oNode = oDom.createElement(sSegName)
      oNode.appendChild oTmpDOM.childNodes(0)
      oRootNode.appendChild oNode
    End If
'  End If
  
Sub_Exit:
  On Error Resume Next
  Set oTmpDOM = Nothing
  Exit Sub
  
Sub_Err:
  gobjShared.ShowVBError Err.Number, Err.Description, "LoadNode Procedure", "DMBE.XMLGen", vbExclamation
  Resume Sub_Exit
  Resume 'Debug
End Sub

Private Sub LoadZ_SEG(ByRef rstIn As ADODB.Recordset, adFormat As ADODB.PersistFormatEnum, sSegName As String)
  Dim oStream As ADODB.Stream
  
  On Error GoTo Sub_Err
  
  Set oStream = New ADODB.Stream
  
  'This loads the rst into a stream
  rstIn.Save oStream, adFormat
      
  Set oNode = oRootNode.selectSingleNode("Z_SEG")
  
  If oNode Is Nothing Then
    Set oNode = oDom.createElement("Z_SEG")
  End If
  
  If adFormat = adPersistXML Then
    Set oNodeLevel1 = oDom.createElement(sSegName)
        oNodeLevel1.Text = oStream.ReadText(-1)
        oNode.appendChild oNodeLevel1
        oRootNode.appendChild oNode
  Else
    'Handle the ADT format
  
  End If
  
Sub_Exit:
  On Error Resume Next
  Set oStream = Nothing
  Exit Sub
  
Sub_Err:
  gobjShared.ShowVBError Err.Number, Err.Description, "LoadZ_SEG Procedure", "DMBE.XMLGen", vbExclamation
  Resume Sub_Exit
  Resume 'Debug
End Sub


Public Property Get OutputFile() As String
    OutputFile = sOutputFile
End Property

Public Property Get OutputXMLString() As String
    OutputXMLString = sOutPutXML
End Property


'This procedure gets called for two purposes
' 1. It can be used to persist an rst to a file for testing purposes
' 2. It is used to build a dom with variaous pieces to the when competed it can be processed against the xsl file
'
Public Sub PersistRst(ByRef rstIn As ADODB.Recordset, sSegName As String, eAction As eSegAction, Optional sMsgName As String = "CopyRST", Optional sPath As String = vbNullString, Optional adFormat As ADODB.PersistFormatEnum = adPersistXML, Optional bIncludeZSEG As Boolean = True)
  On Error Resume Next
  
  If Len(sPath) > 0 Then
  
    'Delete if already exists
    Kill sPath
    
    'Persists the ADODB.Recordset to the path specified
    rstIn.Save sPath, adFormat
  
  Else
  
    'Persists the ADODB.Recordset directly into the DOM object so it can be transformed
    rstIn.Save rst_DOM, adPersistXML
    
    'Set this property variable so it can be retrieved with the OutPutXMLString property if needed
    sOutPutXML = rst_DOM.xml
    
    'Transform the data into the temporary dom until it is processed against the complete xsl
    DoTransform eAction, sMsgName
    
    'Add this output as a node to the current tree
    If eAction = eCopyRST_Load Then
      LoadNode sOutPutXML, sSegName
      If bIncludeZSEG Then
        LoadZ_SEG rstIn, adFormat, sSegName
      End If
    End If

    
    
  End If
End Sub


Public Function CompleteMsg(sMsgName As String) As Boolean
  
  On Error GoTo Func_Exit
  
  'Change the current dom to the temp dom for transformation
  rst_DOM.loadXML oRootNode.xml

  'Now transform it against
  DoTransform eNoLoad, sMsgName
  
  CompleteMsg = True


Func_Exit:
  On Error Resume Next
  Exit Function

Func_Err:
  gobjShared.ShowVBError Err.Number, Err.Description, "CompleteMsg Procedure", "DMBE.XMLGen", vbExclamation
  Resume Func_Exit
  Resume 'Debug


End Function

Private Function ReturnSegName(eSegName As eMsgName) As String
  
  Select Case eSegName
    Case 0
      ReturnSegName = "CopyRST"
    Case 1
      ReturnSegName = "ADT_A04"
    Case 2
      ReturnSegName = "ADT_ZA1"
    Case 3
      ReturnSegName = "VXU_V04"
  End Select

End Function


Public Property Get rstDOM() As MSXML2.DOMDocument30
    Set rstDOM = rst_DOM
End Property




Public Sub DoTransform(eAction As eSegAction, sMsgName As String)
    Dim xsld As MSXML2.FreeThreadedDOMDocument30
    Dim xsltemp As MSXML2.XSLTemplate30
    Dim xslproc As MSXML2.IXSLProcessor
    Dim xsldoc As String
    
    On Error GoTo Sub_Err
    
    Set xsld = New FreeThreadedDOMDocument30
    Set xsltemp = New MSXML2.XSLTemplate30
    
    
    xsld.async = False
    
    'Load the XSL File
    xsldoc = GetXSLFile(sMsgName, eAction)
    
    If Len(xsldoc) = 0 Then
      'Set the property to vbnullstring to prevent previous data from being loaded
      sOutPutXML = vbNullString
      GoTo Sub_Exit
    End If
    
    
    xsld.Load xsldoc
    
    'Add the XSL file to the template
    Set xsltemp.stylesheet = xsld
    
    'Setup the XSL Processor
    Set xslproc = xsltemp.createProcessor
      
    'Add the XML to transform
    xslproc.input = rst_DOM
    
    xslproc.Transform
    
    'Set the output to a variable or it will be lost
    sOutPutXML = xslproc.Output
          
Sub_Exit:
  On Error Resume Next
  Set xsld = Nothing
  Set xsltemp = Nothing
  Exit Sub

Sub_Err:
  gobjShared.ShowVBError Err.Number, Err.Description, "DoTransform Procedure", "DMBE.XMLGen", vbExclamation
  Resume Sub_Exit
  Resume 'Debug

End Sub

Private Sub Class_Initialize()
  If gobjShared Is Nothing Then
    Set gobjShared = New MMMHISShared.CWShared
  End If
  Set oDom = New MSXML2.DOMDocument30
  Set rst_DOM = New MSXML2.DOMDocument30
  
  'LoadXSLFiles
  LoadMsgDict
  
End Sub

Private Sub Class_Terminate()
    Set oDom = Nothing
    Set rst_DOM = Nothing
End Sub




