VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MC4_Ops"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'<< SCR #29291;   Developer: Brian Mowbray 11/14/2002 12:23 PM
'Even thoough this is called MC4 it impliments the ability to read in an xml file from an external source such as BMIST

Option Explicit

' Private Class Variables
Private m_strPatientName As String
Private m_strPatientSSN As String
Private m_strPatientUnitNumber As String
Private m_enmMC4Status As DMStatus
Private m_strPatients() As String
Private m_strEIC As String

Private sOutputFile As String
Private sOutPutXML As String
Private rst_DOM As MSXML2.DOMDocument30
Private rstXSL  As ADODB.Recordset
Private dicDom As Dictionary


Private oDom As MSXML2.DOMDocument30
Private oRootNode As MSXML2.IXMLDOMElement
Private oNode As MSXML2.IXMLDOMElement

'Not Used yet
Private oNodeLevel1 As MSXML2.IXMLDOMElement
Private oNodeLevel2 As MSXML2.IXMLDOMElement
Private oNodeLevel3 As MSXML2.IXMLDOMElement

'SF SCR 77513
Private mFilesProcessed As Dictionary

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


Public Function DeleteFilesProcessed() As Boolean
    
    'SF SCR 77513 Delete the files that have been processed. Called from DataManager.frmRegularView.QueryPatientEIC
    
Dim I As Integer

On Error GoTo ErrHnd

    For I = 0 To mFilesProcessed.Count - 1
        Kill mFilesProcessed.Items(I)
    Next
        
    mFilesProcessed.RemoveAll
    
Exit Function

ErrHnd:
    If Err.Number = 53 Then 'File not found
        Resume Next
    End If
    
    If gobjShared Is Nothing Then
      Set gobjShared = New CWShared
   End If
   gobjShared.ShowVBError Err.Number, Err.Description, "DeleteFilesProcessed", "MC4_Ops", vbInformation
   
End Function

' Used to add any final touches necessary to the XML
Public Sub FinalizeXML()


End Sub

Private Function GetXSLFile(sMsgName As String, eAction As eSegAction) As String


Func_Exit:
  On Error Resume Next
  Exit Function
  
Func_Err:
  gobjShared.ShowVBError Err.Number, Err.Description, "GetXSLFile Procedure", "DMBE.XMLGen", vbExclamation
  Resume Func_Exit
  Resume 'Debug

End Function

Public Sub LoadNode(ByVal sXMLString As String, sSegName As String)
  Dim oTmpDOM As MSXML2.DOMDocument30
  
  On Error GoTo Sub_Err
  
  Set oTmpDOM = New MSXML2.DOMDocument30
  
  'This loads the XML string created by the transformation of the XML->XSL to XML
  If oTmpDOM.loadXML(sXMLString) Then
    Set oNode = oDom.createElement(sSegName)
    oNode.appendChild oTmpDOM.childNodes(0)
    oRootNode.appendChild oNode
  End If
  
Sub_Exit:
  Set oTmpDOM = Nothing
  Exit Sub
  
Sub_Err:
  gobjShared.ShowVBError Err.Number, Err.Description, "LoadNode Procedure", "DMBE.XMLGen", vbExclamation
  Exit Sub
  Resume 'Debug
End Sub

Public Property Get OutputFile() As String
    OutputFile = sOutputFile
End Property

Public Property Get OutputXMLString() As String
    OutputXMLString = sOutPutXML
End Property

Public Function CompleteMsg(sMsgName As String) As Boolean
  
  On Error GoTo Func_Exit
  
  'Change the current dom to the temp dom for transformation
  rst_DOM.loadXML oRootNode.xml

  CompleteMsg = True


Func_Exit:
  Exit Function

Func_Err:
  gobjShared.ShowVBError Err.Number, Err.Description, "CompleteMsg Procedure", "DMBE.XMLGen", vbExclamation
  Exit Function
  Resume 'Debug


End Function

Public Property Get rstDOM() As MSXML2.DOMDocument30
    Set rstDOM = rst_DOM
End Property

Public Sub DoTransform(sXSLFile As String)
    Dim xsld As MSXML2.FreeThreadedDOMDocument30
    Dim xsltemp As MSXML2.XSLTemplate30
    Dim xslproc As MSXML2.IXSLProcessor
    Dim xsldoc As String
    
    On Error GoTo Sub_Err
    
    Set xsld = New FreeThreadedDOMDocument30
    Set xsltemp = New MSXML2.XSLTemplate30
    
    
    xsld.async = False
    
    'Load the XSL File
    xsldoc = sXSLFile
    
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
    
    'Debug.Print sOutPutXML
      
Sub_Exit:
  Set xsld = Nothing
  Set xsltemp = Nothing
  Exit Sub

Sub_Err:
  gobjShared.ShowVBError Err.Number, Err.Description, "DoTransform Procedure", "DMBE.XMLGen", vbExclamation
  Exit Sub
  Resume 'Debug

End Sub

Public Sub BuildBMISTXML(ByRef rstIn As ADODB.Recordset)
   
    Dim rst As ADODB.Recordset
    Dim sTemp As String
    
    On Error GoTo Func_Err
    
    If oDas Is Nothing Then
        Set oDas = oDatabase.CHCSII_DAS(Auto) 'New GEMS_DAS
    End If
    
    PersistRst rstIn
      
    'Transform the data into the temporary dom until it is processed against the complete xsl
    DoTransform gobjShared.RTDataPath & "\CHCSIIToBMIST.xsl"
    
    rst_DOM.loadXML sOutPutXML
    
    
Func_Exit:
  On Error Resume Next
  Exit Sub

Func_Err:
   gobjShared.ShowVBError Err.Number, Err.Description, "BuildBMISTXML Function ", "MC4_Ops", vbExclamation
   Resume Func_Exit
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
  Set myDMBE_Ops = Nothing
  
  Exit Function

Func_Err:
  If gobjShared Is Nothing Then
    Set gobjShared = New CWShared
  End If
  gobjShared.ShowVBError Err.Number, Err.Description, "WriteXML Function", "XMLGen", vbExclamation
  Resume Func_Exit
  Resume 'Debug
End Function








Private Function XMLToOracleDate(ByVal sValue As String) As Variant

    Dim dDate As Date
    
    sValue = Replace(sValue, "/", "")
    sValue = Replace(sValue, "-", "")
    sValue = Replace(sValue, " ", "")
    sValue = Replace(sValue, ":", "")
    
    If IsNumeric(sValue) Then
    
        sValue = Left$(sValue & "00000000000000", 14)
        
        sValue = Mid$(sValue, 5, 2) & "/" & Mid$(sValue, 7, 2) & "/" & Mid$(sValue, 1, 4) & " " & Mid$(sValue, 9, 2) & ":" & Mid$(sValue, 11, 2) & ":" & Mid$(sValue, 13, 2)
        dDate = CDate(sValue)
        XMLToOracleDate = dDate
    Else
        XMLToOracleDate = Null
    End If
    
End Function

Private Function ConvertTextToRTF(ByRef Text As String) As String
    Dim sTmp As String
    If Left$(Text, 5) = "{\rtf" Then
        ConvertTextToRTF = Text
    Else
        sTmp = Replace(Text, Chr(13) & Chr(10), "\par ")
        sTmp = Replace(sTmp, Chr(13), "\par ")
        sTmp = Replace(sTmp, Chr(10), "\par ")
        sTmp = Replace(sTmp, Chr(9), "\tab ")
        ConvertTextToRTF = "{\rtf1\ansi\ansicpg1252\uc0\deff0" _
            & "{\fonttbl{\f0\fswiss\fcharset0\fprq2 Arial;}}" _
            & "{\colortbl;\red0\green0\blue0;}" _
            & "\deftab720 \plain \pard \f0\fs16 " _
            & sTmp & " }"
    End If
    
End Function

Public Function QueryMC4(sPathToMC4 As String, RetVal As Long, Optional ShowMessages As Boolean = True) As Boolean
   Dim lngReturnValue As Long
   Dim objFSO As FileSystemObject
   Dim lngSearchHandle As Long
   Dim strFileName As String
   Dim strFilePattern As String
   Dim sDetailMsg As String
   Dim varAvail As Variant
   Dim strAbsolutePathName As String
   Dim sTempXML As String
   Dim sDetail As String
   Dim lCnt As Long
   Dim myDMBE_Ops As DMBE_Ops
   Dim sBMISTData As String
   Dim bCompressed As Boolean
   Dim bEncrypted As Boolean


On Error GoTo ErrHandler
    
    Set mFilesProcessed = New Dictionary
    
   ' assume MC4 Files are not available
   m_enmMC4Status = eDM_NotReady
   
   ' check if operation concerns a PIC device
      If sPathToMC4 <> "" Then
         Set objFSO = New FileSystemObject
         
         If objFSO.FolderExists(sPathToMC4) Then
            ' drive exists
             
            strAbsolutePathName = objFSO.GetDriveName(objFSO.GetAbsolutePathName(sPathToMC4))
            
            Set varAvail = objFSO.GetDrive(strAbsolutePathName)
            If varAvail.IsReady Then
                m_enmMC4Status = eDM_Ready
            Else
               m_enmMC4Status = eDM_NotReady
               lngReturnValue = ERROR_PATH_DOES_NOT_EXIST
               Set objFSO = Nothing
               Set varAvail = Nothing
               Exit Function
            End If
    
            Set objFSO = Nothing
            Set varAvail = Nothing
            
            
            '>> Begin:SCR #37536;   Developer: Brian Mowbray 07/28/2003 03:46 PM
            strFilePattern = sPathToMC4 & "\CHCSII_*.xml"
            '<< End: SCR #37536;
            strFilePattern = Replace(strFilePattern, "\\", "\")
            lngReturnValue = GetFirstFile(strFilePattern, lngSearchHandle, strFileName)
            'strFileName is the file name found above
            sDetail = vbNullString
            lCnt = 0
            If lngReturnValue = gclFOUND Then

              Do Until lngReturnValue = gclNOTFOUND
                lCnt = lCnt + 1
                
                If rst_DOM Is Nothing Then
                  Set rst_DOM = New MSXML2.DOMDocument30
                End If
                
                rst_DOM.async = False
                rst_DOM.validateOnParse = False
                rst_DOM.resolveExternals = False
                
                'SCR 76478 By default BMIST files are not encrypted. If this fails, then assume it is encrytped.
                rst_DOM.Load sPathToMC4 & "\" & strFileName
                
                'SCR 62595 - The BMIST File needs to be
                '  Decrypted before it's loaded into the MSXML2.DOMDocument30 object.
                If rst_DOM.parseError <> 0 Then
                  Set myDMBE_Ops = New DMBE_Ops
                  
                  sBMISTData = vbNullString
                  lngReturnValue = 0
                  
                  bCompressed = IsFileCompressed(strFileName)
                  
                  'SCR 62595 BMIST files can be encrypted, but their filename will not reflect it.
                  bEncrypted = True
                  'bEncrypted = IsFileEncrypted(strFileName)
                  
                  sBMISTData = myDMBE_Ops.ReadFileContents(sPathToMC4 & "\" & strFileName, bCompressed, bEncrypted, lngReturnValue)
                  'lngReturnValue = myDMBE_Ops.WriteFileContents(sPathToMC4 & "\" & strFileName, sBMISTData, False, False, False)
                  Set myDMBE_Ops = Nothing
                  'rst_DOM.Load sPathToMC4 & "\" & strFileName
                  If Len(sBMISTData) > 0 Then
                    If rst_DOM.loadXML(sBMISTData) Then
                    End If
                  End If
                End If
                
                Set gobjShared = New CWShared
              
                If rst_DOM.parseError <> 0 Then
                  m_enmMC4Status = eDM_NotReady
                  '>> Begin:SCR #37536;   Developer: Brian Mowbray 07/28/2003 03:46 PM
                  sDetailMsg = sDetailMsg & "Error Loading XML File: " & strFileName & vbCrLf & "Error Description: " & rst_DOM.parseError.reason & vbCrLf & "Reason: The file being processed may not be a BMIST generated file, please check the source path."
                  '<< End: SCR #37536;
                  lngReturnValue = ERROR_UNABLE_TO_PARSE_BMIST_XML_DATA
                Else
                  
                  'SF SCR 77513 Save the filename that has been processed.
                  mFilesProcessed.Add strFileName, sPathToMC4 & "\" & strFileName
                  
                  DoTransform gobjShared.RTDataPath & "\BMISTToA28.xsl"
                  sTempXML = "<?xml version='1.0'?>" & sOutPutXML & "<?xml version='1.0'?>"
                  
                  DoTransform gobjShared.RTDataPath & "\BMISTToA04.xsl"
                  sTempXML = sTempXML & sOutPutXML
                  
                  If Not dicDom.Exists(lCnt) Then
                    dicDom.Add lCnt, sTempXML
                    m_strEIC = m_strEIC & sTempXML
                  End If
                End If
                lngReturnValue = GetNextFile(lngSearchHandle, strFileName)
              Loop
              FindClose lngSearchHandle
              
              If Len(sDetailMsg) > 0 Then
                '>> Begin:SCR #37536;   Developer: Brian Mowbray 07/28/2003 03:47 PM
                MsgBox sDetailMsg, vbInformation, "Error reading some BMIST Files"
              '<< End: SCR #37536;
              End If
              m_enmMC4Status = eDM_ReadyAndReadable
            Else
               'SCR 78771, Don't display a message that there are not any BMIST files to process if polling.
'              ' No BMIST files are available to process
              RetVal = ERROR_NO_BMIST_FILES_TO_PROCESS
              'SCR 78771
              If ShowMessages Then
                  'SCR 128280
                  MsgBox "No AHLTA-MOBILE files found in specified path!" & vbCrLf & sPathToMC4, vbInformation, "NO FILES FOUND"
              End If
              
              Exit Function
            End If
          End If
      Else
        ' No BMIST files are available to process
        lngReturnValue = ERROR_NO_BMIST_FILES_TO_PROCESS
        Exit Function
      End If
            
   QueryMC4 = True
   
Exit Function

ErrHandler:
    
   If gobjShared Is Nothing Then
      Set gobjShared = New CWShared
   End If
   gobjShared.ShowVBError Err.Number, Err.Description, "QueryMC4", "MC4_Ops", vbInformation
   Err.Clear
    
    Exit Function
    Resume
    
End Function

Private Function IsFileCompressed(ByVal sFileName As String) As Boolean

On Error GoTo ErrHandler

Dim sTempString As String
  
  Dim sArray() As String
  'FileName = "CHCSII_CE_BMIST1140005805.XML"
  
  IsFileCompressed = False
  
  sFileName = UCase$(Trim$(sFileName))
  
  If sFileName = vbNullString Then
    Exit Function
  End If
  
  If InStr(sFileName, "_") > 0 Then
    sArray = Split(sFileName, "_")
    If UBound(sArray) > 1 Then
      sTempString = Trim$(sArray(1))
      If Len(sTempString) = 2 Then
        If sTempString = "CX" Or sTempString = "CE" Then
          IsFileCompressed = True
        End If
      End If
    End If
  End If

Exit Function

ErrHandler:

   If gobjShared Is Nothing Then
      Set gobjShared = New CWShared
   End If
   gobjShared.ShowVBError Err.Number, Err.Description, "IsFileCompressed", "MC4_Ops", vbExclamation
  
    IsFileCompressed = False

End Function


Private Function IsFileEncrypted(ByVal sFileName As String) As Boolean

On Error GoTo ErrHandler

Dim sTempString As String
  
  Dim sArray() As String
  'FileName = "CHCSII_CE_BMIST1140005805.XML"
  
  IsFileEncrypted = False
  
  sFileName = UCase$(Trim$(sFileName))
  
  If sFileName = vbNullString Then
    Exit Function
  End If
  
  If InStr(sFileName, "_") > 0 Then
    sArray = Split(sFileName, "_")
    If UBound(sArray) > 1 Then
      sTempString = Trim$(sArray(1))
      If Len(sTempString) = 2 Then
        If sTempString = "XE" Or sTempString = "CE" Then
          IsFileEncrypted = True
        End If
      End If
    End If
  End If

Exit Function

ErrHandler:

   If gobjShared Is Nothing Then
      Set gobjShared = New CWShared
   End If
   gobjShared.ShowVBError Err.Number, Err.Description, "IsFileEncrypted", "MC4_Ops", vbExclamation
  
    IsFileEncrypted = False

End Function
Public Property Get MC4Status() As DMStatus
  MC4Status = m_enmMC4Status
End Property
'This procedure gets called for two purposes
' 1. It can be used to persist an rst to a file for testing purposes
' 2. It is used to build a dom with variaous pieces to the when competed it can be processed against the xsl file
'
Public Sub PersistRst(ByRef rstIn As ADODB.Recordset, Optional sSegName As String = vbNullString)
  
    On Error GoTo ErrHnd
  
  'Persists the ADODB.Recordset directly into the DOM object so it can be transformed
  rstIn.Save rst_DOM, adPersistXML
    
  'Set this property variable so it can be retrieved with the OutPutXMLString property if needed
  sOutPutXML = rst_DOM.xml
  
  'Add this output as a node to the current tree
  If Len(sSegName) > 0 Then
    LoadNode sOutPutXML, sSegName
  End If
    
Exit Sub

ErrHnd:
    MsgBox "Application Error: " & Err.Description & " Occurred in DMBE.MC4_Ops.PersistRst."

End Sub


Public Property Let MC4Status(StatusIn As DMStatus)
  m_enmMC4Status = StatusIn
End Property

Public Property Get PatientName() As String
  PatientName = m_strPatientName
End Property
Public Property Let PatientName(sNameIn As String)
  m_strPatientName = sNameIn
End Property

Public Property Get PatientSSN() As String
  PatientSSN = m_strPatientSSN
End Property
Public Property Let PatientSSN(sSSNIn As String)
  m_strPatientSSN = sSSNIn
End Property

Public Property Get MC4Patients() As Variant
  MC4Patients = m_strPatients
End Property
Public Property Let MC4Patients(ByVal enmDMFPatientInfo As Variant)
  m_strPatients = enmDMFPatientInfo
End Property

Public Property Get PatientUnitNumber() As String
  PatientUnitNumber = m_strPatientUnitNumber
End Property
Public Property Let PatientUnitNumber(sUnitIn As String)
  m_strPatientUnitNumber = sUnitIn
End Property

Public Property Get EIC_XML() As String
  EIC_XML = m_strEIC
End Property

Private Sub Class_Initialize()

On Error GoTo ErrHnd

  If gobjShared Is Nothing Then
    Set gobjShared = New MMMHISShared.CWShared
  End If
  Set dicDom = New Dictionary
  Set oDom = New MSXML2.DOMDocument30
  Set rst_DOM = New MSXML2.DOMDocument30
  
Exit Sub
ErrHnd:
    MsgBox "Application Error: " & Err.Description & " Occurred in DMBE.MC$_Ops.Initialize"
    
End Sub


Public Function GetPatientDict() As Dictionary
  Set GetPatientDict = dicDom

End Function
