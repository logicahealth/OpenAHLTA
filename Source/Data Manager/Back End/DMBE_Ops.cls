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
Attribute VB_Name = "DMBE_Ops"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IDMBE.I_DMBE

'Private DELIM_CN As ADODB.Connection
'Private Const DB_CNSTR As String = "Provider=MSDAORA.1;Password=gems;User ID=gems;Data Source=cdr;Persist Security Info=True"

Private m_JmewsRSObjects As Collection

Private unitNumber As String
Private m_EICEncounterNumber As String

' TMIP Message Variables
Public EventID As Long
Public EventType As Long
Public sModuleName As String

'<< Begin:SCR #29291;   Developer: Brian Mowbray 11/14/2002 11:29 AM
'Private oMC4Dom As MSXML2.DOMDocument30
'>> End: SCR #29291;

Public WithEvents EvntSvr As CDMEvntSvr
Attribute EvntSvr.VB_VarHelpID = -1
Private mcctEvents As EvntSvrConnector
'Private SafeStore As Collection
Private DMBE_Ack As Long

Private Const gNCID_DATA_MANAGER_BACKEND As String = "1999983"

' IDMBE Properties
Private m_enmDataSet As DMDataSets
Private m_enmDestination As DMLocations
Private m_enmSource As DMLocations
Private m_enmDMFormat As DMFormats
Private m_colPatientList As Collection
Private m_strPicPath As String
Private m_strAltPicPath As String
Private m_blnUpdateIsLocked As Boolean
Private m_enmCompression As DMCompression
Private m_enmCrypto As DMCrypto
Private m_objQueryPIC As Object

'<< SCR #30953;   Developer: Brian Mowbray 12/18/2002 02:02 PM
Private m_lEncSource As enEncSource

'--- SCR 85597   Sherry Wang   1/25/2006
Private mc_strPassword As String

'Private DMBEPatient As MMMHISPatient.Patient

'Private m_blnIsSAMSTransaction As Boolean
'Private m_strSAMSTempHL7 As String
'Private m_bSAMSIsDestination As Boolean
'Private m_bSAMSIsSource As Boolean

'Private m_blnIsEICTransaction As Boolean
Private m_strEICTempHL7 As String
Private m_bEICIsDestination As Boolean
Private m_bEICIsSource As Boolean

Public Temp As String
Private Const DELIMITOR_CONSTANT As String = "|"
Private Const DAT_LINE_BREAK_CHARACTERS As String = "<@$%^%$@>" & vbCr '<SCR 24045, 24055 CC>
Private Const DELIMITOR_CHARACTERS As String = "|*FB*|"
'<SCR 24045 CC>

Private m_dicDatasets As Dictionary
Private m_sPicFile As String

Private Const mc_strEndTableName = "ETN"



Private Enum TmipStatus
    Sent = 1
    Complete = 2
    LDDBError = 3
    TMIPError = 4
End Enum

Public Enum InpatientMsgType
  enBedStatusSetup = 1
  enBedStatusUpdate = 2
  enAdmit = 3
  enDischarge = 4
End Enum

'Private m_DMBE_IDMappers As DMBE_IDMappers
Private m_DMBE_IDMappers As IDMBE.I_IDMappers '<SCR 24303 CC>
Private m_enmXMLMessageSource As XMLMessageSource
Private m_enmDMErrorCode As DMErrorCodes
Private m_colRSObjects As Collection

Private m_lngPICPatientUnitNumber As Long
Private m_sFileName As String

Private Sub AddToJMEWSCollection(Col As Collection)
        
'SF 62831
'This copies the collection of RSObject_Ops to a new persisted collection
On Error GoTo ErrHnd
Dim oRsOps As RSObject_Ops

    If m_JmewsRSObjects Is Nothing Then
        Set m_JmewsRSObjects = New Collection
    End If
    
    For Each oRsOps In Col           'Make sure there are no duplicate tables
        m_JmewsRSObjects.Add oRsOps, UCase(oRsOps.TableName)
    Next
    
Exit Sub

ErrHnd:
    If Err.Number = 457 Then
        'Duplicate key found, so continue
        Resume Next
    Else
        MsgBox "Application error: " & Err.Description & " Occurred in: DMBE.DMBE_OPS.AddToJMEWSCollection"
    End If
    Exit Sub
    Resume
    
End Sub



'SCR 84451 and 84452 Regen
Public Sub DeleteRegenMessageLog()
    
    DeleteRegenLog
    
End Sub

Public Property Let RegenPatientName(Name As String)

    gRegenPatientName = Name
    
End Property

'<< SCR #43003;   Developer: Brian Mowbray 01/05/2004 04:42 PM
Public Function SendTMIPInpatientMsg(lMsgType As InpatientMsgType, ByVal sXML As String) As Long
   
    Dim rst As ADODB.Recordset
    Dim sTemp As String
    Dim oXML As MC4_Ops
    Dim sMessageType As String
    Dim sMessageGUID As String
    Dim enmResult As DMErrorCodes
    Dim lErrresult As Long
    Dim sXSLFile As String
    Dim sHL7Type As String
    
    Dim oDom As MSXML.DOMDocument
    Dim oItemNode As MSXML.IXMLDOMNode
    Dim oRoot As MSXML.IXMLDOMNode
    Dim oNode As MSXML.IXMLDOMNode
    
    '102123
    GetTMIPFlag
    If gbTMIP = False Then
        SendTMIPInpatientMsg = ERROR_SUCCESS
        Exit Function
    End If
    
    Set oDom = New MSXML.DOMDocument

'    oDom.Load "c:\temp\Sample_BedStatusSetup.xml"
'
'    sXML = oDom.xml

    On Error GoTo Func_Err
    
    SendTMIPInpatientMsg = ERROR_SUCCESS
    
    If Len(sXML) <= 0 Then
      SendTMIPInpatientMsg = ERROR_EMPTY_RECORDSET
      Exit Function
    End If
    
    Set oXML = New MC4_Ops
   
    If Not (gobjPatient Is Nothing) Then
        If Not BuildPatientRst(False, False) Then
          Err.Raise 999, "BuildTMIPMsg Function", "Error creating the Patient recordset, patient segments cannot be generated!"
          GoTo Func_Exit
        End If
    Else
      Err.Raise 999, "BuildTMIPMsg Function", "Error creating the Patient recordset, patient segments cannot be generated!"
      GoTo Func_Exit
    End If

    sMessageGUID = GetNewGUID
    sMessageType = "INPATIENT"
       
    oXML.CreateRootNodeElement sMessageType
    
    sXSLFile = gobjShared.RTDataPath & "\"
    
    'SCR 43003
    Select Case lMsgType
    Case enBedStatusSetup 'Bed Status Setup
      sMessageType = "BedStatusSetup"
      BuildMSHRst "Z20", sMessageGUID
      sXSLFile = sXSLFile & "ADT_Z20.xsl"
      sHL7Type = "TMIP_Z20" '120555
    Case enBedStatusUpdate 'Bed Status Update
      sMessageType = "BedStatusUpdate"
      BuildMSHRst "A20", sMessageGUID
      sXSLFile = sXSLFile & "ADT_A20.xsl"
      sHL7Type = "TMIP_A20" '120555
    Case enAdmit 'Admit
      sMessageType = "Admit"
      BuildMSHRst "A01", sMessageGUID
      sXSLFile = sXSLFile & "ADT_A01.xsl"
    Case enDischarge 'Discharge
      sMessageType = "Discharge"
      BuildMSHRst "A03", sMessageGUID
      sXSLFile = sXSLFile & "ADT_A03.xsl"
    Case Else
      sMessageType = "Unknown"
    End Select
    
    oXML.PersistRst g_rstMSH, "MSH"
    oXML.PersistRst g_rstPatient, "PID"
    
    oXML.LoadNode sXML, "RECORDSETS"
      
    oXML.CompleteMsg "INPATIENT"
      
    oXML.DoTransform sXSLFile
    
    Set oDom = New MSXML.DOMDocument

    oDom.loadXML oXML.OutputXMLString

    Set oRoot = oDom.selectSingleNode("//INPATIENT_MSGS")

    For Each oItemNode In oRoot.childNodes

      Set oNode = oItemNode.selectSingleNode("MSH/MSH.10")

      enmResult = PersistTMIPMessageToDatabase(oNode.Text, sMessageType, "", g_rstPatient.Clone, oItemNode.xml, sMessageType, False, sHL7Type)

    Next oItemNode
    
    If enmResult = ERROR_SUCCESS Then
        SendTMIPInpatientMsg = ERROR_SUCCESS ' good return
    Else
        SendTMIPInpatientMsg = enmResult
    End If
  
Func_Exit:
  Set oXML = Nothing
  Set oDom = Nothing
  Exit Function

Func_Err:
   gobjShared.ShowVBError Err.Number, Err.Description, "SendTMIPInpatientMsg Function", "DMBE", vbExclamation
   SendTMIPInpatientMsg = ERROR_UNHANDLED_SYSTEM_ERROR
   GoTo Func_Exit
   
   Resume 'Debug
End Function

Private Function I_DMBE_AppendTransferRS(ByVal TableName As String, RS As ADODB.Recordset) As Boolean
    
    Dim rsObject As IDMBE.I_RSObjects
    
    On Error GoTo ErrHandler
    
    Set rsObject = New DMBE.RSObject_Ops
    rsObject.TableName = TableName
    Set rsObject.DBRECORDSET = RS
    
    m_colRSObjects.Add rsObject, TableName

    I_DMBE_AppendTransferRS = True
Exit Function

ErrHandler:
    
    I_DMBE_AppendTransferRS = False
    Err.Clear
    
End Function

Private Function I_DMBE_ClearTransferRS() As Boolean
Dim oItem As I_RSObjects
Dim oRS As ADODB.Recordset

On Error GoTo ErrHandler
    For Each oItem In m_colRSObjects
      Set oRS = oItem.DBRECORDSET
      oRS.Close
      Set oRS = Nothing
      Set oItem = Nothing
    Next oItem
  
    Set m_colRSObjects = New Collection
    I_DMBE_ClearTransferRS = True
Exit Function

ErrHandler:
  'Continue on error
  Resume Next
End Function

Private Sub Class_Initialize()

   On Error GoTo ErrHnd
   
   Set oDas = oDatabase.CHCSII_DAS(Auto) 'New GEMS_DAS
   
   Set mobjConvertSno = New actxSearchSnoMed.IConvert
   Set gobjShared = New CWShared
   
   '--- SCR 85597   Sherry Wang   1/25/2006
   mc_strPassword = GetEncryptionKey
   
   Set m_colRSObjects = New Collection
   
   If CheckEFFolderStructure = True Then
   End If
   
   LoadDataSet
   
   GetTMIPSendNowFlag
    
   m_enmDMErrorCode = ERROR_SUCCESS
   
Exit Sub
ErrHnd:
    
    MsgBox "Application Error: " & Err.Description & " Occurred in DMBE.DMBE_Ops.Initialize"
    Exit Sub
    
    Resume
End Sub

Private Sub Class_Terminate()
   Set mobjConvertSno = Nothing
   Set gobjShared = Nothing
   Set oDas = Nothing
End Sub

Private Sub EvntSvr_SendEvent(Status As Long, DataType As String, cData As Collection, XMlMsg As String, AppNCID As String)
   If AppNCID = gNCID_DATA_MANAGER_BACKEND Then
      DMBE_Ack = Status
   End If
End Sub

' BEGINNING OF DMBE INTERFACE DEFINITION OBJECT

Private Property Get I_DMBE_AltPicPath() As String
   I_DMBE_AltPicPath = m_strAltPicPath
End Property

Private Property Let I_DMBE_AltPicPath(rhs As String)
   If Len(rhs) > 0 Then
     If Right(rhs, 1) = "\" Then
       m_strAltPicPath = rhs
     Else
       m_strAltPicPath = rhs & "\"
     End If
   End If
End Property

Private Function I_DMBE_ClearParameters() As Boolean
   I_DMBE_ClearParameters = True
'   Exit Function
'   PICPATH = ""
'   Set m_colPatientList = Nothing
'   DATASET = eDM_All
'   m_enmDestination = eDM_CDR
'   m_enmSource = eDM_CDR
'   m_enmDMFormat = eDM_ADT
End Function

Private Property Let I_DMBE_Compression(rhs As IDMBE.DMCompression)
   
   m_enmCompression = rhs
   
End Property

Private Property Get I_DMBE_Compression() As IDMBE.DMCompression
   
   I_DMBE_Compression = m_enmCompression
   
End Property

Private Function I_DMBE_CopyPIC() As Long
   
   Dim unitNumber As String
   Dim sFormat As String
   Dim FilePattern As String
   Dim FilePIC As String
   Dim FileName As String
   Dim FileName2 As String
   Dim RetVal As Long
   Dim RetVal2 As Long
   Dim hSearch As Long
   Dim fso As FileSystemObject
   Dim SrcFilePath As String
   Dim DestFilePath As String
   Dim idx As Integer
   Dim sMsg As String
   Dim bCopyError As Boolean
   
   On Error GoTo ErrTrap
   
   ' assume that the copy will succeed;  i.e., prove failure
   I_DMBE_CopyPIC = ERROR_SUCCESS  ' good return
   
   Set fso = New FileSystemObject
   
   If m_strPicPath <> "" Then
      If m_strAltPicPath <> "" Then
        unitNumber = gobjPatient.unitNumber
        'Get Data File because the PIC file will always already be in the destination path
        sFormat = FormatEnumToTxt(m_enmDMFormat)
        
        ' Begin:SCR #44865;   Developer:  11/07/2003
        'Need to check this machines UnitNumber on PIC file
        ' or if PIC created on another machine (different Unit Number)
        ' SCR #47954;   Developer: MAG 01/13/2004
        For idx = 0 To 1
          FileName = ""
          Select Case idx
          Case 0
            FilePattern = m_strPicPath & unitNumber & "*" & "." & sFormat
            FilePIC = m_strPicPath & unitNumber & "*" & ".PIC"
          Case 1
            FilePattern = m_strPicPath & fso.GetBaseName(m_sPicFile) & "." & sFormat
            FilePIC = m_strPicPath & fso.GetBaseName(m_sPicFile) & ".PIC"
          End Select
          hSearch = 0
          RetVal = GetFirstFile(FilePattern, hSearch, FileName)
          FindClose hSearch
          hSearch = 0
          RetVal2 = GetFirstFile(FilePIC, hSearch, FileName2)
          FindClose hSearch
          If (RetVal = gclFOUND) And (RetVal2 = gclFOUND) Then
             SrcFilePath = m_strPicPath & FileName
             DestFilePath = m_strAltPicPath & FileName
             
             'This will Write the file to the Destination
             'XML File
             'SCR 128754, 129250
             sMsg = "Cannot copy file from: " & SrcFilePath & " to: " & DestFilePath _
                & vbCrLf & " Please verify that the windows user account has write permissions to the destination directory."
             bCopyError = True
             fso.CopyFile SrcFilePath, DestFilePath, True
             bCopyError = False
             
             'SCR 84451 and 84452 Regen
             LogRegenMessage DestFilePath
             
             SrcFilePath = m_strPicPath & FileName2
             DestFilePath = m_strAltPicPath & FileName2
             
             'This will Write the file to the Destination
             'PIC File
             'SCR 128754, 129250
             sMsg = "Cannot copy file from: " & SrcFilePath & " to: " & DestFilePath _
                & vbCrLf & " Please verify that the windows user account has write permissions to the destination directory."
             bCopyError = True
             fso.CopyFile SrcFilePath, DestFilePath, True
             I_DMBE_CopyPIC = ERROR_SUCCESS  ' good return
             bCopyError = False
             
             'SCR 84451 and 84452 Regen
             LogRegenMessage DestFilePath
             
             Exit For
          Else
             I_DMBE_CopyPIC = ERROR_DATA_NOT_FOUND  ' bad return
          End If
        Next idx
        ' End: SCR #44865;
      Else
        I_DMBE_CopyPIC = ERROR_PATH_DOES_NOT_EXIST  ' bad return
      End If
   Else
      I_DMBE_CopyPIC = ERROR_PATH_DOES_NOT_EXIST  ' bad return
   End If

Cleanup:
   Set fso = Nothing
Exit Function

ErrTrap:
   'SCR 128754, 129250
    If bCopyError = True Then
         MsgBox sMsg, vbInformation
         Err.Clear
         I_DMBE_CopyPIC = ERROR_UNHANDLED_SYSTEM_ERROR
         Resume Cleanup
    End If
   
   If gobjShared Is Nothing Then
      Set gobjShared = New CWShared
   End If
   gobjShared.ShowVBError Err.Number, Err.Description, "I_DMBE_CopyPIC Function", "DMBE", vbExclamation
   Err.Clear
   I_DMBE_CopyPIC = ERROR_UNHANDLED_SYSTEM_ERROR
   Resume Cleanup
End Function

Private Function I_DMBE_CreateDetachedRS(TableName As String, objRS As ADODB.Recordset, Optional ByVal OpenRS As Boolean = True) As Long
    
    I_DMBE_CreateDetachedRS = GetFieldProps(TableName, objRS, OpenRS)
    
End Function

Private Property Let I_DMBE_Crypto(rhs As IDMBE.DMCrypto)
   m_enmCrypto = rhs
End Property

Private Property Get I_DMBE_Crypto() As IDMBE.DMCrypto
   I_DMBE_Crypto = m_enmCrypto
End Property

Private Property Get I_DMBE_CurrentDMErrorCode() As Long
    I_DMBE_CurrentDMErrorCode = m_enmDMErrorCode
End Property

Public Property Get I_DMBE_DataSet() As IDMBE.DMDataSets
   I_DMBE_DataSet = m_enmDataSet
End Property

Private Property Let I_DMBE_DataSet(rhs As IDMBE.DMDataSets)

    '8 = Encounter
    
    m_enmDataSet = rhs
    
    
End Property

Public Property Get I_DMBE_Destination() As IDMBE.DMLocations
   I_DMBE_Destination = m_enmDestination
End Property

Private Property Let I_DMBE_Destination(rhs As IDMBE.DMLocations)
    m_enmDestination = rhs
End Property

Private Property Let I_DMBE_EICHL7Message(rhs As String)
    m_strEICTempHL7 = rhs
End Property

Private Property Get I_DMBE_EICHL7Message() As String
     I_DMBE_EICHL7Message = m_strEICTempHL7
End Property

Private Property Let I_DMBE_EICIsDestination(ByVal rhs As Boolean)
    m_bEICIsDestination = rhs
End Property

Private Property Get I_DMBE_EICIsDestination() As Boolean
    I_DMBE_EICIsDestination = m_bEICIsDestination
End Property

Private Property Let I_DMBE_EICIsSource(ByVal rhs As Boolean)
    m_bEICIsSource = rhs
End Property

Private Property Get I_DMBE_EICIsSource() As Boolean
    I_DMBE_EICIsSource = m_bEICIsSource
End Property

'<< SCR #30953;   Developer: Brian Mowbray 12/18/2002 03:03 PM
Private Property Let I_DMBE_EncounterSource(ByVal rhs As IDMBE.enEncSource)
  m_lEncSource = rhs
End Property

'<< SCR #30953;   Developer: Brian Mowbray 12/18/2002 03:03 PM
Private Property Get I_DMBE_EncounterSource() As Long
  I_DMBE_EncounterSource = m_lEncSource
End Property

'<SCR 22576 CC>
Private Function I_DMBE_Execute(cObj As Collection, Optional ByRef TMIPDataRetrievalResult As TMIPDataRetrievalResult) As Long
'</SCR 22576 CC>
   
   Dim RetVal As Long
   Dim objHL7Parser As HL7XMLParser
   Dim oItem As I_RSObjects
   Dim rsProvider As ADODB.Recordset
   Dim oSQLOps As ICHCSII_SQl.ISqlOpsEx
   Dim sNCID As String
   Dim sName() As String
   Dim sLName As String
   Dim sFName As String
   
   m_enmDMErrorCode = ERROR_SUCCESS
   
   
   If CheckEFFolderStructure = True Then
   End If
   
   If Not m_colRSObjects Is Nothing Then
        For Each oItem In m_colRSObjects
            cObj.Add oItem
        Next
   End If
   
   Set m_colRSObjects = New Collection
   
   ' check for valid patient object
   If gobjPatient Is Nothing Then
      I_DMBE_Execute = ERROR_PATIENT_OBJECT_IS_NOTHING
      m_enmDMErrorCode = ERROR_PATIENT_OBJECT_IS_NOTHING
      Exit Function
   Else
      ' set the unit_number from the patient object
      unitNumber = gobjPatient.unitNumber
      ' save variables for transaction audit
      gPatientName = gobjPatient.Name
      gPatientSSN = gobjPatient.SSN
      gUnitNumber = gobjPatient.unitNumber
      geDataSet = I_DMBE_DataSet
      geFormat = I_DMBE_Format
      geSource = I_DMBE_Source
      geDestination = I_DMBE_Destination
   End If
   
   If m_enmSource = eDM_EIC And m_enmDestination = eDM_LDDB Then
        Set objHL7Parser = New HL7XMLParser
        Select Case geDataSet
            Case eDM_Demographics
                'Should never come here
                'RetVal = objHL7Parser.GetRecordsetsFromEIC(cObj, UnitNumber, eDM_Demographics)
            Case eDM_Encounters
                RetVal = objHL7Parser.GetRecordsetsFromEIC(cObj, unitNumber, eDM_Encounters, m_strEICTempHL7, m_enmXMLMessageSource)
                Set rsProvider = GetRS(cObj, "PROVIDER")
                If Not rsProvider Is Nothing Then
                  Set oSQLOps = oDatabase.CHCSII_SQLOPS_EX(Auto)
                  Do While Not rsProvider.EOF
                    If rsProvider("SSN") & vbNullString <> vbNullString Then
                      sNCID = rsProvider("NCID") & vbNullString
                      sName = Split(rsProvider("NAME") & vbNullString, ",")
                      If UBound(sName) = 0 Then
                        sLName = sName(0)
                        sFName = ""
                      Else
                        sLName = sName(0)
                        sFName = sName(1)
                      End If
                      Call oSQLOps.ValidateProvider2(sNCID, rsProvider("SSN"), sFName, "", sLName, True)
                      Call m_DMBE_IDMappers.Add(NCID_PROVIDER_SERVICES, "NCID", rsProvider("NCID") & vbNullString, sNCID)
                    End If
                    rsProvider.MoveNext
                  Loop
                End If
        End Select
        I_DMBE_Execute = RetVal
        m_enmDMErrorCode = RetVal
        Set objHL7Parser = Nothing
        Exit Function
   End If
   
   ' check if operation concerns a PIC device
   If m_enmSource = eDM_PIC Or m_enmSource = eDM_AGGPIC Or m_enmDestination = eDM_PIC Or m_enmDestination = eDM_AGGPIC Then
      
      unitNumber = m_lngPICPatientUnitNumber
      RetVal = I_DMBE_ValidatePIC()
      If RetVal <> ERROR_SUCCESS Then
         I_DMBE_Execute = RetVal
         m_enmDMErrorCode = RetVal
         Exit Function
      End If
   End If
   
   'Validate PIC First then read the recordset
   If m_enmSource = eDM_PIC Or m_enmSource = eDM_AGGPIC Then
         If cObj.Count = 0 Then
            'Read PIC will put the recordset into the collection
             RetVal = ReadPIC(cObj, m_enmDataSet, m_enmDMFormat, m_strPicPath)
                          
             'Could not Validate Facility and Clinic, keep going
             If RetVal = -934 Then
                RetVal = 0
             End If

         End If
         If Not cObj Is Nothing Then
             For Each oItem In cObj
              'Save the RS Collection for future use
              m_colRSObjects.Add oItem
            Next
         End If
         I_DMBE_Execute = RetVal
         m_enmDMErrorCode = RetVal
         Exit Function
   End If

   If m_enmDestination = eDM_PIC Or m_enmDestination = eDM_AGGPIC Then
      If m_enmDataSet = eDM_All Then
         I_DMBE_Execute = ERROR_DATASET_NOT_SUPPORTED ' bad return - all health context does not have any meaning for writing to a PIC
         m_enmDMErrorCode = ERROR_DATASET_NOT_SUPPORTED
         Exit Function
      Else
         If Not cObj Is Nothing Then
            For Each oItem In cObj
                m_colRSObjects.Add oItem
                Debug.Print oItem.TableName
            Next
         End If
         
'         m_dicDatasets.Add m_enmDataSet, ReturnDataSetName(m_enmDataSet)
'         RetVal = WritePIC(cObj, m_enmDMFormat, m_strPicPath)
         I_DMBE_Execute = RetVal
         m_enmDMErrorCode = RetVal
         Exit Function
      End If
   End If
    
   If m_enmDestination = 8 Then ' eDM_Inpatient SF 62831
        If m_enmDMFormat = eDM_XML Then
            If Not IsMissing(TMIPDataRetrievalResult) Then
               RetVal = JMEWS_EXECUTE(cObj, TMIPDataRetrievalResult)
            Else
               RetVal = JMEWS_EXECUTE(cObj)
            End If
            Exit Function
        End If
   End If
   
   If m_enmDestination = eDM_TMIP Or m_enmDestination = eDM_EIC Then
   
      If m_enmDMFormat = eDM_XML Then
'         If m_enmDataSet = eDM_Readiness Then
'            Call ReadinessTemporaryFix(cObj)
'         End If
'         If m_enmDataSet = eDM_Vitals Then
'            Call VitalsTemporaryFix(cObj)
'         End If
'         If m_enmDataSet = eDM_Encounters Then
'            Call EncountersTemporaryFix(cObj)
'         End If
'         If m_enmDataSet = eDM_Medications Then
'            Call MedicationsTemporaryFix(cObj)
'         End If
'         If m_enmDataSet = eDM_Demographics Then
'            Call DemographicsTemporaryFix(cObj)
'         End If
         
         If Not IsMissing(TMIPDataRetrievalResult) Then
            RetVal = TMIP_EXECUTE(cObj, TMIPDataRetrievalResult)
         Else
            RetVal = TMIP_EXECUTE(cObj)
         End If
         m_enmDMErrorCode = RetVal
         I_DMBE_Execute = RetVal
      Else
         I_DMBE_Execute = ERROR_FORMAT_NOT_SUPPORTED ' format not supported
         m_enmDMErrorCode = ERROR_FORMAT_NOT_SUPPORTED
      End If
   Else
      I_DMBE_Execute = ERROR_DESTINATION_NOT_SUPPORTED ' destination not supported
      m_enmDMErrorCode = ERROR_DESTINATION_NOT_SUPPORTED
   End If
   
End Function

Public Property Get I_DMBE_Format() As IDMBE.DMFormats
   I_DMBE_Format = m_enmDMFormat
End Property

Private Property Let I_DMBE_Format(rhs As IDMBE.DMFormats)
    m_enmDMFormat = rhs
End Property

Private Function I_DMBE_GetDMBEErrorMessage(errorCode As Long) As String
   I_DMBE_GetDMBEErrorMessage = GetErrMessage(errorCode)
End Function

'<SCR 24303 CC>
'Private Property Set I_DMBE_IDMappers(ByRef RHS As Object)
'    Set m_DMBE_IDMappers = RHS
'End Property
'
'Private Property Get I_DMBE_IDMappers() As Object
'    Set I_DMBE_IDMappers = m_DMBE_IDMappers
'End Property


Private Property Set I_DMBE_IDMappers(ByRef rhs As IDMBE.I_IDMappers)
    Set m_DMBE_IDMappers = rhs
End Property

Private Property Get I_DMBE_IDMappers() As IDMBE.I_IDMappers
    Set I_DMBE_IDMappers = m_DMBE_IDMappers
End Property

Private Property Let I_DMBE_PICPatientUnitNumber(ByVal rhs As Long)
    m_lngPICPatientUnitNumber = rhs
End Property

Private Property Get I_DMBE_PICPatientUnitNumber() As Long
    I_DMBE_PICPatientUnitNumber = m_lngPICPatientUnitNumber
End Property

Private Property Let I_DMBE_FileName(ByVal rhs As String)
    m_sFileName = rhs
End Property

Private Property Get I_DMBE_FileName() As String
    I_DMBE_FileName = m_sFileName
End Property

'</SCR 24303 CC>

Private Function I_DMBE_InitDMBE(eSource As IDMBE.DMLocations, eDest As IDMBE.DMLocations, eFormat As IDMBE.DMFormats, eDataSet As IDMBE.DMDataSets, eUnitNumber As Collection, PicPath As String, Optional AltPicPath As String, Optional eCompress As IDMBE.DMCompression, Optional eCrypto As IDMBE.DMCrypto) As Long
   Dim fs As FileSystemObject
On Error GoTo ErrHandler
   
   m_enmSource = eSource
   m_enmDestination = eDest
   If m_enmSource = m_enmDestination Then
      I_DMBE_InitDMBE = ERROR_SOURCE_CANNOT_BE_DESTINATION ' bad return - source cannot equal the destination
      Call I_DMBE_ClearParameters
   End If
   m_enmDMFormat = eFormat
   m_enmDataSet = eDataSet
   
   If IsMissing(eCompress) Then
      m_enmCompression = eDM_NoCompression
   Else
      m_enmCompression = eCompress
   End If
   
   If IsMissing(eCrypto) Then
      m_enmCrypto = eCrypto
   Else
      m_enmCrypto = eDM_NoCrypto
   End If
   
   Set m_colPatientList = eUnitNumber
   PicPath = m_strPicPath
   
   Set fs = New FileSystemObject
   If fs.FolderExists(PicPath) Then
      I_DMBE_InitDMBE = ERROR_SUCCESS ' good return
   Else
      I_DMBE_InitDMBE = ERROR_SOURCE_CANNOT_BE_DESTINATION ' bad return
      Call I_DMBE_ClearParameters
   End If
   If IsMissing(AltPicPath) Then
      AltPicPath = Environ("TEMP")
   Else
      AltPicPath = m_strAltPicPath
   End If
   If fs.FolderExists(AltPicPath) Then
      I_DMBE_InitDMBE = ERROR_SUCCESS ' good return
   Else
      I_DMBE_InitDMBE = ERROR_SOURCE_CANNOT_BE_DESTINATION ' bad return
      Call I_DMBE_ClearParameters
   End If
   
   Set fs = Nothing
Exit Function

ErrHandler:
  Set fs = Nothing
  Err.Raise Err.Number, "I_DMBE.InitDMBE", Err.Description
End Function

Private Function I_DMBE_InitializePIC() As Long
   
   Dim fso As FileSystemObject
   Dim FilePath As String
   Dim RetVal As Boolean
   Dim RetOption As Long
   Dim ErrMsg As String
   Dim sPICFileNAME As String
   Dim bCompress As Boolean
   Dim bEncrypt As Boolean
   Dim sFacilityNCID As String
   
   On Error GoTo InitPICError
   
   Set fso = New FileSystemObject
   If gobjPatient Is Nothing Then
      I_DMBE_InitializePIC = ERROR_PATIENT_OBJECT_IS_NOTHING
      Exit Function
   Else
      sFacilityNCID = gobjLogon.FacilityNCID
      unitNumber = gobjPatient.unitNumber
   End If
   
   sPICFileNAME = unitNumber & "_"
   
   m_enmCompression = eDM_NoCompression
   m_enmCrypto = eDM_Encrypted
   
   Select Case m_enmCompression
   Case eDM_Compressed
     sPICFileNAME = sPICFileNAME & "C" 'Compressed
     bCompress = True
   Case eDM_NoCompression
     sPICFileNAME = sPICFileNAME & "X" 'Not Compressed
   Case Else
     sPICFileNAME = sPICFileNAME & "U" 'Unknown Compression
   End Select
   
   bEncrypt = False
   Select Case m_enmCrypto
   Case eDM_Encrypted
     sPICFileNAME = sPICFileNAME & "E"  'Encryption
     bEncrypt = True
   Case eDM_NoCrypto
     sPICFileNAME = sPICFileNAME & "X"  'No Encryption
   Case Else
     sPICFileNAME = sPICFileNAME & "U"  'Unknown Encryption
   End Select
   
   sPICFileNAME = sPICFileNAME & "_" & sFacilityNCID
   sPICFileNAME = sPICFileNAME & ".PIC"
   
   If m_strAltPicPath = "" Then
      I_DMBE_InitializePIC = ERROR_PIC_DEVICE_PATH_HAS_NOT_BEEN_SET
      Exit Function
   End If
   
   FilePath = m_strAltPicPath & sPICFileNAME
   RetVal = fso.FileExists(FilePath)
   If RetVal Then
      ErrMsg = "This PIC device has been initialized." & vbCrLf
      ErrMsg = ErrMsg & "Do you want to re-initialize the PIC?" & vbCrLf
      RetOption = MsgBox(ErrMsg, vbYesNo, "Initialize PIC Device")
      If RetOption = vbYes Then
         fso.DeleteFile (FilePath)
      Else
         I_DMBE_InitializePIC = ERROR_SUCCESS ' good return
         Exit Function
      End If
   End If
   
'   If m_enmDestination = eDM_PIC Then
'      FilePattern = m_strPicPath & "*.*" ' delete all files from PIC device
'      RetCode = GetFirstFile(FilePattern, hSearch, FileName)
'      While RetCode = gclFOUND
'         Kill m_strPicPath & "\" & FileName
'         FileName = ""
'         RetCode = GetNextFile(hSearch, FileName)
'      Wend
'   End If
   'Used to be .FullName
   I_DMBE_InitializePIC = WritePICInfo(FilePath, bCompress, bEncrypt)
   
   m_sPicFile = sPICFileNAME 'Set the Module Variable PIC File Name
   
   Set fso = Nothing
   
   Exit Function
InitPICError:
   Resume Next ' kill is probably in error...may have a directory on PIC...occurred during testing
End Function

Private Property Get I_DMBE_IsEICTransaction() As Boolean

End Property

Private Property Let I_DMBE_IsEICTransaction(ByVal rhs As Boolean)

End Property

Private Property Set I_DMBE_LogonObj(rhs As Object)
    Set gobjLogon = rhs
End Property

Private Property Get I_DMBE_Password() As String
   I_DMBE_Password = mc_strPassword
End Property

Private Property Get I_DMBE_PatientList() As Collection
   Set I_DMBE_PatientList = m_colPatientList
End Property

Private Property Set I_DMBE_PatientList(rhs As Collection)
    Set m_colPatientList = rhs
End Property

Private Property Set I_DMBE_PatientObj(rhs As Object)
   Set gobjPatient = rhs
End Property

Private Property Get I_DMBE_PatientObj() As Object
   Set I_DMBE_PatientObj = gobjPatient
End Property

Private Property Let I_DMBE_PicPath(rhs As String)
   If Len(rhs) > 0 Then
     If Right(rhs, 1) = "\" Then
        m_strPicPath = rhs
     Else
        m_strPicPath = rhs & "\"
     End If
   End If
End Property

Public Property Get I_DMBE_PicPath() As String
   I_DMBE_PicPath = m_strPicPath
End Property



Private Function I_DMBE_QueryPIC(strPathToPIC As String, lngReturnStatus As Long) As Object
   
   Dim lngReturnValue As Long
   Dim objFSO As FileSystemObject
   Dim lngSearchHandle As Long
   Dim strFileName As String
   Dim strFilePattern As String
   Dim strFilePath As String
   Dim asPatientInfo() As String
   Dim strData() As String
   Dim strFileData As String
   Dim varAvail As Variant
   Dim strAbsolutePathName As String
   Dim sDatasets As String
   Dim nPatientCnt As Long
   Dim sMsg As String
   
   Dim oDom As MSXML2.DOMDocument30
   Dim oNode As MSXML2.IXMLDOMElement
   Dim oItem As MSXML2.IXMLDOMElement
   Dim bCompress As Boolean
   Dim bEncrypt As Boolean
   Dim lMaxImportCount As Long 'SF SCR 87423
   
   Dim objQueryPIC As IDMBE.I_QueryPIC
   
On Error GoTo ErrHandler

   'SF 87423 Read the max value from the registry.
   lMaxImportCount = GetMaxImportCnt

   Set objQueryPIC = New DMBE.QueryPIC_Ops 'This is the return object
   
   ' assume PIC is not available
   'objQueryPIC.I_QueryPIC_PicStatus = eDM_NotReady
   objQueryPIC.PicStatus = eDM_NotReady
   
   ' check if operation concerns a PIC device
   If m_enmSource = eDM_PIC Or m_enmSource = eDM_AGGPIC Then
      If strPathToPIC <> "" Then
         Set objFSO = New FileSystemObject
         
         'If objFso.DriveExists(strPathToPIC) Then
         If objFSO.FolderExists(strPathToPIC) Then
            ' drive exists
             
            strAbsolutePathName = objFSO.GetDriveName(objFSO.GetAbsolutePathName(strPathToPIC))
            
            Set varAvail = objFSO.GetDrive(strAbsolutePathName)
            If varAvail.IsReady Then
                'objQueryPIC.I_QueryPIC_PicStatus = eDM_Ready
                objQueryPIC.PicStatus = eDM_Ready
            Else
               'objQueryPIC.I_QueryPIC_PicStatus = eDM_NotReady
               objQueryPIC.PicStatus = eDM_NotReady
               lngReturnStatus = ERROR_PIC_DEVICE_DOES_NOT_EXIST
               Set I_DMBE_QueryPIC = Nothing
               Set varAvail = Nothing
               GoTo Cleanup
            End If
    
            Set varAvail = Nothing
            
            strFilePattern = strPathToPIC & "\*.pic"
            strFilePattern = Replace(strFilePattern, "\\", "\")
            lngReturnValue = GetFirstFile(strFilePattern, lngSearchHandle, strFileName)
            
            If lngReturnValue = gclNOTFOUND Then
' Begin:SCR #38406;   Developer: Mag 07/03/2003 02:51 PM
'               If m_enmSource = eDM_PIC Or m_enmSource = eDM_AGGPIC Then
                 lngReturnStatus = ERROR_PIC_DEVICE_DOES_NOT_EXIST
'               Else
'                 ' PIC device is not initialized
'                 lngReturnStatus = ERROR_PIC_DEVICE_HAS_NOT_BEEN_INITIALIZED
'               End If
' End: SCR #38406;
               Set I_DMBE_QueryPIC = Nothing
               GoTo Cleanup
            End If
            
            ReDim asPatientInfo(5, 0)
            nPatientCnt = 0
            Do Until lngReturnValue = gclNOTFOUND
               strData = Split(strFileName, "_", -1, vbTextCompare)
               If UBound(strData) > 0 Then
                 ' get Patient Unit Number
                 objQueryPIC.PatientUnitNumber = strData(0)
                 
                 If UBound(strData) > 1 Then
                   If Len(strData(2)) > 4 Then
                     objQueryPIC.FileName = Split(strFileName, ".")(0)
                   End If
                 End If
                 ' Begin:SCR #36997;   Developer: Mag 06/11/2003 02:27 PM
                 'Initialize Compression & Encryption flags each time per file
                 bCompress = False
                 bEncrypt = False
                 ' End: SCR #36997;
                 
                 If Left(strData(1), 1) = "C" Then
                   bCompress = True
                 End If
                
                 If Mid(strData(1), 2, 1) = "E" Then
                   bEncrypt = True
                 End If
                
                 objQueryPIC.PicStatus = eDM_ReadyAndReadable
               Else
                 objQueryPIC.PicStatus = eDM_NotReadable
                 lngReturnStatus = ERROR_DAT_FILE_PROCESSING
                 Set I_DMBE_QueryPIC = Nothing
                 GoTo Cleanup
               End If
            
               strFilePath = strPathToPIC & "\" & strFileName
               strFilePath = Replace(strFilePath, "\\", "\")
               
               strFileData = ReadFileContents(strFilePath, bCompress, bEncrypt)
            
               Set oDom = New MSXML2.DOMDocument30
               If Not oDom.loadXML(strFileData) Then
                 MsgBox "Invalid PIC device format. " & vbCrLf & vbCrLf & "File : """ & strFilePath & """", vbExclamation, "DMBE"
                 lngReturnStatus = DMErrorCodes.ERROR_DATA_NOT_FOUND
                 GoTo Cleanup
               End If
            
               Set oNode = oDom.selectSingleNode("//PICINFO/PATIENT/NAME")
               If Not oNode Is Nothing Then
                  objQueryPIC.PatientName = oNode.Text
               End If
            
            
               Set oNode = oDom.selectSingleNode("//PICINFO/PATIENT/SSN")
               If Not oNode Is Nothing Then
                 objQueryPIC.PatientSSN = oNode.Text
               End If
            
'               Set oNode = oDom.selectSingleNode("//PICINFO/PATIENT/BIRTHDATE")
'               If Not oNode Is Nothing Then
'                 If IsDate(oNode.Text) Then
'                    objQueryPIC.PatientDOB = CDate(oNode.Text)
'                 Else
'                    objQueryPIC.PatientDOB = 0
'                End If
'               End If
               
               objQueryPIC.PicFormat = GetPicFormat(strPathToPIC, objQueryPIC.PatientUnitNumber)

               ' Identify all of the datasets on the PIC device
               Dim lngDataSets() As Long
               Dim intArrayCount As Integer
            
               Set oNode = oDom.selectSingleNode("//PICINFO/DATASETS")
               
               sDatasets = ""
               ReDim lngDataSets(0)
               If Not oNode Is Nothing Then
                 intArrayCount = 0
                 For Each oItem In oNode.childNodes
                   ReDim Preserve lngDataSets(intArrayCount)
                   lngDataSets(intArrayCount) = oItem.Text
                   sDatasets = sDatasets & oItem.Text & ","
                   intArrayCount = intArrayCount + 1
                 Next oItem
            
                 If intArrayCount > 0 Then
                    objQueryPIC.DataSetsInPIC = lngDataSets
                 Else
                  If m_enmSource = eDM_AGGPIC Then
                    sMsg = "No datasets exist for patient: " & objQueryPIC.PatientName _
                           & vbCrLf & vbCrLf & "   Unit number: " & objQueryPIC.PatientUnitNumber
                  Else
                    sMsg = "No dataset exists in the PIC device."
                  End If
                  MsgBox sMsg, vbExclamation, "DMBE Error"
                  lngReturnStatus = DMErrorCodes.ERROR_DATA_NOT_FOUND
                  GoTo Cleanup
                 End If
               Else
                objQueryPIC.DataSetsInPIC = lngDataSets
               End If
               
               'Setup Array of Patient Info for Multiple Patient Processing (DISK)
               '0 - Unit #
               '1 - SSN
               '2 - Name
               '3 - Format
               '4 - datasets (comma delimited)
               '5 - Birth Date
               ReDim Preserve asPatientInfo(5, nPatientCnt)
               asPatientInfo(0, nPatientCnt) = objQueryPIC.PatientUnitNumber
               asPatientInfo(1, nPatientCnt) = objQueryPIC.PatientSSN
               asPatientInfo(2, nPatientCnt) = objQueryPIC.PatientName
               asPatientInfo(3, nPatientCnt) = objQueryPIC.PicFormat
               asPatientInfo(4, nPatientCnt) = sDatasets
               asPatientInfo(5, nPatientCnt) = objQueryPIC.FileName
               
'               If objQueryPIC.PatientDOB <> 0 Then
'                 asPatientInfo(5, nPatientCnt) = objQueryPIC.PatientDOB
'               Else
'                 asPatientInfo(5, nPatientCnt) = ""
'               End If
               
               lngReturnValue = GetNextFile(lngSearchHandle, strFileName)
               nPatientCnt = nPatientCnt + 1
               
               If lngReturnValue = gclFOUND And m_enmSource = eDM_PIC Then
                 MsgBox "Multiple patients found on PIC Device, only one patient per PIC device allowed.", vbExclamation, "DMBE Error"
                 lngReturnStatus = DMErrorCodes.ERROR_PIC_DEVICE_TYPE_OR_AVAILABILITY
                 GoTo Cleanup
               End If
               
               'SF 2/8/06 87423 Limit the number of patients to load at one time to 500
               If nPatientCnt = lMaxImportCount Then
                  MsgBox lMaxImportCount & " files is the maximum recommended number that should be processed at one time." & vbCrLf & "Please close and restart the Theater Application before processing the next " & lMaxImportCount & " files." & vbCrLf & "You may continue without restarting the application, but your computer may run low on memory.", vbInformation
                  Exit Do
               End If
            Loop
            FindClose lngSearchHandle
            
            objQueryPIC.AGGPICPatients = asPatientInfo
            
            lngReturnStatus = ERROR_SUCCESS   ' good return
            Set I_DMBE_QueryPIC = objQueryPIC
            GoTo Cleanup
         
         Else
            objQueryPIC.PicStatus = eDM_NotReady
            lngReturnStatus = ERROR_PIC_DEVICE_DOES_NOT_EXIST
            Set I_DMBE_QueryPIC = Nothing
            GoTo Cleanup
         End If
      
      Else
         lngReturnStatus = ERROR_PIC_DEVICE_PATH_HAS_NOT_BEEN_SET  ' bad return
         Set I_DMBE_QueryPIC = Nothing
         GoTo Cleanup
      End If
   
'   '<SCR 22437>
'   'Fix for reading DISK
'   ElseIf m_enmSource = eDM_AGGPIC Then
'
'        Set m_objQueryPIC = QueryAGGPIC(strPathToPIC, lngReturnStatus)
'        Set I_DMBE_QueryPIC = m_objQueryPIC
'        lngReturnStatus = DMErrorCodes.ERROR_SUCCESS
   Else
      lngReturnStatus = ERROR_DESTINATION_NOT_SUPPORTED  ' bad return
      Set I_DMBE_QueryPIC = Nothing
      GoTo Cleanup
   End If
   
Cleanup:
  Set objFSO = Nothing
  Set m_objQueryPIC = Nothing

Exit Function

ErrHandler:
    
   If gobjShared Is Nothing Then
      Set gobjShared = New CWShared
   End If
   gobjShared.ShowVBError Err.Number, Err.Description, "QueryPIC", "DMBE", vbExclamation
   Err.Clear
   Set I_DMBE_QueryPIC = Nothing
   Resume Cleanup
End Function

'Private Function QueryAGGPIC(ByVal strPathToPIC As String, ByVal lngReturnStatus As DMErrorCodes) As IDMBE.I_QueryPIC
'
'    Dim objQueryPIC As IDMBE.I_QueryPIC
'    Dim objFSO As FileSystemObject
'    Dim lngReturnValue As Long
'    Dim lngSearchHandle As Long
'    Dim strFileName As String
'    Dim strFilePattern As String
'    Dim strFilePath As String
'    Dim strData() As String
'    Dim strTemp As String
'    'Dim FileType As String
'    Dim strTableName As String
'    Dim strDataSet As String
'    'Dim enmDataSet As IDMBE.DMDataSets
'    'Dim enmFormat As IDMBE.DMFormats
'    Dim varAvail As Variant
'    Dim strPatientInfo() As String
'    Dim intCount As Integer
'    Dim strTempData() As String
'    Dim lngDataSets() As Long
'    Dim lngTemp As Long
'    Dim intArrayCount As Integer
'
'    ' check if operation concerns a PIC device
'    If Not m_enmSource = eDM_AGGPIC Then
'      lngReturnStatus = ERROR_DESTINATION_NOT_SUPPORTED  ' bad return
'      Set objQueryPIC = Nothing
'      Exit Function
'    End If
'
'    Set objQueryPIC = New DMBE.QueryPIC_Ops
'
'    If strPathToPIC = "" Then
'        lngReturnStatus = ERROR_PIC_DEVICE_PATH_HAS_NOT_BEEN_SET  ' bad return
'        Set objQueryPIC = Nothing
'        Exit Function
'    End If
'
'    Set objFSO = New FileSystemObject
'    If Not objFSO.FolderExists(strPathToPIC) Then
'        Set objFSO = Nothing
'        objQueryPIC.PicStatus = eDM_NotReady
'        lngReturnStatus = ERROR_PIC_DEVICE_DOES_NOT_EXIST
'        'Set I_DMBE_QueryPIC = Nothing
'        Exit Function
'    End If
'
'    'Folder exists
'    'Get the absolute path name first
'    strTemp = objFSO.GetDriveName(objFSO.GetAbsolutePathName(strPathToPIC))
'    'Then check if the drive is "ready"
'    Set varAvail = objFSO.GetDrive(strTemp)
'    If varAvail.IsReady Then
'        'objQueryPIC.PicStatus = eDM_Ready
'        objQueryPIC.PicStatus = eDM_Ready
'    Else
'        objQueryPIC.PicStatus = eDM_NotReady
'        lngReturnStatus = ERROR_PIC_DEVICE_DOES_NOT_EXIST
'        Set objFSO = Nothing
'        Set objQueryPIC = Nothing
'        Set varAvail = Nothing
'        Exit Function
'    End If
'
'    'This is the only place the file system object is needed
'    Set objFSO = Nothing
'    Set varAvail = Nothing
'
'    strTemp = strPathToPIC & "\*.pic"
'    lngReturnValue = GetFirstFile(strTemp, lngSearchHandle, strFileName)
'    'strFileName is the file name found above (ByRef)
'
'    If lngReturnValue = gclNOTFOUND Then
'        ' PIC device is not initialized
'        lngReturnStatus = ERROR_PIC_DEVICE_HAS_NOT_BEEN_INITIALIZED
'        Set objQueryPIC = Nothing
'        Exit Function
'    Else
'
'        objQueryPIC.PicStatus = eDM_ReadyAndReadable
'        intCount = 0
'
'        Do Until lngReturnValue = gclNOTFOUND
'            'First get all the .EIC data
'            ReDim Preserve strTempData(intCount)
'            strTemp = strPathToPIC & "\" & strFileName
'            ' Begin:SCR #14025;   Developer: MAG 07/17/2002 11:45 AM
'            strData = Split(strFileName, ".", -1, vbTextCompare)
'            'Add Patient Unit Number to string
'            strTempData(intCount) = ReadFileContents(strTemp) & "|" & strData(0)
'            ' End: SCR #14025;
'            intCount = intCount + 1
'            lngReturnValue = GetNextFile(lngSearchHandle, strFileName)
'        Loop  'lngReturnStatus = gclNOTFOUND
'
'        Dim intI As Integer
'
'        'Then split them into patient data array
'        ReDim Preserve strPatientInfo(intCount - 1, 2)
'        For intI = 0 To intCount - 1
'            strData = Split(strTempData(intI), "|", -1, vbTextCompare)
'            strPatientInfo(intI, 0) = strData(3)    'Unit Number
'            strPatientInfo(intI, 1) = strData(2)    'SSN
'            strPatientInfo(intI, 2) = strData(0)    'Name
'            objQueryPIC.PatientUnitNumber = ""
'        Next intI
'
'        objQueryPIC.AGGPICPatients = strPatientInfo
'
'
'    End If 'lngReturnValue = gclNOTFOUND
'
'    ' Identify all of the datasets on the DISK device
'
'    intArrayCount = 0
'    lngSearchHandle = 0
'    strFilePattern = strPathToPIC & "\*.*"
'    strFileName = ""
'
'    lngReturnValue = GetFirstFile(strFilePattern, lngSearchHandle, strFileName)
'
'    If lngReturnValue = gclFOUND Then
'        ' Initialize values
'        strTableName = ""
'        'FileType = ""
'        'enmDataSet = -1   ' illegal values
'        'enmFormat = -1  ' illegal values
'        strData = Split(strFileName, "_", -1, vbTextCompare)
'        If UBound(strData) > 4 Then
'          strTableName = strData(3)
'          'strDataSet = GetClinicalDataAreaByName(strTableName, enmDataSet)
'          lngTemp = TranslateDataSetCode(strTableName)
'          If lngTemp <> 9999 Then
'                ReDim lngDataSets(0)
'                lngDataSets(0) = lngTemp
'                intArrayCount = 1
'                objQueryPIC.PICDataSetExists(strDataSet) = True
'          End If
'        End If
'
'
'        Dim blnValueIsFound As Boolean
'        Dim blnExitThisLoop As Boolean
'
'        blnExitThisLoop = False
'
'        Do Until blnExitThisLoop = True
'          strFileName = ""
'          lngReturnValue = GetNextFile(lngSearchHandle, strFileName)
'          If lngReturnValue = gclFOUND Then
'             ' Initialize values
'             strTableName = ""
'             'FileType = ""
'             'enmDataSet = -1   ' illegal values
'             'enmFormat = -1  ' illegal values
'             strData = Split(strFileName, "_", -1, vbTextCompare)
'             If UBound(strData) > 4 Then
'                'strTableName = strData(4) & "_" & strData(4)
'                strTableName = strData(3)
'                lngTemp = TranslateDataSetCode(strTableName)
'
'
'                If lngTemp <> 9999 Then
'                    blnValueIsFound = False
'                    For intI = 0 To intArrayCount - 1
'                        If lngTemp = lngDataSets(intI) Then
'                            blnValueIsFound = True
'                        End If
'                    Next intI
'                    If blnValueIsFound = False Then
'                        ReDim Preserve lngDataSets(intArrayCount)
'                        lngDataSets(intArrayCount) = lngTemp
'                        intArrayCount = intArrayCount + 1
'                    End If
'                    objQueryPIC.PICDataSetExists(strDataSet) = True
'                End If 'lngTemp <> 9999
'
'              End If 'UBound(strData) > 4
'          Else
'            blnExitThisLoop = True
'          End If
'
'       Loop 'While lngReturnValue = gclFOUND
'
'    End If 'lngReturnValue = gclFOUND from GetFirstFile
'
'    If intArrayCount > 0 Then
'        objQueryPIC.DataSetsInPIC = lngDataSets
'    Else
'
'    End If
'
'    lngReturnStatus = ERROR_SUCCESS   ' good return
'    Set QueryAGGPIC = objQueryPIC
'
'Exit Function
'    If gobjShared Is Nothing Then
'      Set gobjShared = New CWShared
'   End If
'   gobjShared.ShowVBError Err.Number, Err.Description, "QueryAGGPIC Function", "DMBE", vbExclamation
'   Err.Clear
'   Set QueryAGGPIC = Nothing
'
'End Function


'Private Property Let I_DMBE_SAMSHL7Message(RHS As String)
'    m_strSAMSTempHL7 = RHS
'End Property
'
'Public Property Get I_DMBE_SAMSHL7Message() As String
'    I_DMBE_SAMSHL7Message = m_strSAMSTempHL7
'End Property

'Private Property Let I_DMBE_SAMSIsDestination(ByVal RHS As Boolean)
'    m_bSAMSIsDestination = RHS
'End Property
'
'Private Property Get I_DMBE_SAMSIsDestination() As Boolean
'    I_DMBE_SAMSIsDestination = m_bSAMSIsDestination
'End Property
'
'Private Property Let I_DMBE_SAMSIsSource(ByVal RHS As Boolean)
'    m_bSAMSIsSource = RHS
'End Property
'
'Private Property Get I_DMBE_SAMSIsSource() As Boolean
'    I_DMBE_SAMSIsSource = m_bSAMSIsSource
'End Property

Public Property Get I_DMBE_Source() As IDMBE.DMLocations
   I_DMBE_Source = m_enmSource
End Property

Private Property Let I_DMBE_Source(rhs As IDMBE.DMLocations)
    m_enmSource = rhs
End Property


Private Property Let I_DMBE_UpdateLock(rhs As Boolean)
   m_blnUpdateIsLocked = rhs
End Property

Private Property Get I_DMBE_UpdateLock() As Boolean
   I_DMBE_UpdateLock = m_blnUpdateIsLocked
End Property

Private Function I_DMBE_ValidatePIC() As Long
   Dim hSearch As Long
   Dim FileName As String
   Dim FilePattern As String
   Dim RetCode As Long
   Dim sReadVal As String
   Dim sValArr() As String
   Dim fso As FileSystemObject
   Dim UserOption As Long
   Dim ErrMsg As String
   Dim DriveType As String
   Dim DriveReady As Boolean
   
   Dim oDom As MSXML2.DOMDocument30
   Dim oNode As MSXML2.IXMLDOMElement
   Dim bCompress As Boolean
   Dim bEncrypt As Boolean
   
   On Error GoTo ErrHandler
   
   Set fso = New FileSystemObject
   
   If m_strPicPath = "" And m_strAltPicPath = "" Then
      I_DMBE_ValidatePIC = ERROR_PIC_DEVICE_PATH_HAS_NOT_BEEN_SET
      GoTo Cleanup
   Else
      ErrMsg = ""
      RetCode = CheckPICDriveType(m_strPicPath, DriveType, DriveReady)
      
      If m_enmSource = eDM_PIC Or m_enmDestination = eDM_PIC Then
         '<SCR 22701 CC>
         If RetCode = DRIVE_FIXED Or DRIVE_REMOVABLE Then
            FilePattern = m_strPicPath & "*" & ".PIC"
         Else
            ErrMsg = "PIC must be a local REMOVABLE OR FIXED DRIVE." & vbCrLf
            ErrMsg = ErrMsg & "The path (" & m_strPicPath & ") provided is a " & DriveType & "drive type." & vbCrLf & vbCrLf
         End If
         '</SCR 22701 CC>
      End If
      
      If m_enmSource = eDM_AGGPIC Or m_enmDestination = eDM_AGGPIC Then
         If RetCode = DRIVE_FIXED Or RetCode = DRIVE_REMOVABLE Then
            If m_enmDestination = eDM_AGGPIC Then
                ' Begin:SCR #36985;   Developer: Mag 06/11/2003 02:33 PM
                FilePattern = m_strAltPicPath & gobjPatient.unitNumber & "_??_" & gobjLogon.FacilityNCID & ".PIC"
                ' End: SCR #36985;
            Else
                FilePattern = m_strPicPath & m_lngPICPatientUnitNumber & "*" & ".PIC"
            End If
            
         '-- SCR 190312 Webb 6/10/2009 CD-ROM error
         ElseIf m_enmDestination = eDM_AGGPIC Then
            ErrMsg = "Aggregated PIC (Floppy, ZIP Disk or Hard Disk) must be a local REMOVABLE or FIXED DRIVE." & vbCrLf
            ErrMsg = ErrMsg & "The path (" & m_strPicPath & ") provided is a " & DriveType & " drive type." & vbCrLf & vbCrLf
         End If
      End If
      
      If DriveReady = False Then
         ErrMsg = ErrMsg & "The selected PIC device is not ready."
      End If
      If ErrMsg <> "" Then
         Call MsgBox(ErrMsg, vbCritical, "ValidatePIC")
         I_DMBE_ValidatePIC = ERROR_PIC_DEVICE_TYPE_OR_AVAILABILITY
         GoTo Cleanup
      End If
      
      hSearch = 0
      FileName = ""
      
      RetCode = GetFirstFile(FilePattern, hSearch, FileName) 'FileName is ByRef
      FindClose hSearch
      
      If RetCode = gclNOTFOUND Then
        If m_enmDestination = eDM_AGGPIC Then 'DISK
            I_DMBE_ValidatePIC = I_DMBE_InitializePIC
            'I_DMBE_ValidatePIC = ERROR_SUCCESS ' good return
            GoTo Cleanup
        ' Begin:SCR #38406;   Developer: Mag 07/02/2003 05:20 PM
        ElseIf m_enmDestination = eDM_PIC Then    'PIC Device
        ' End: SCR #38406;
          ErrMsg = "PIC device has not been initialized." & vbCrLf
          ErrMsg = ErrMsg & "Do you wish to Initialize for current patient?" & vbCrLf
          ErrMsg = ErrMsg & "Patient Name: " & gobjPatient.Name & vbCrLf
          UserOption = MsgBox(ErrMsg, vbYesNo, "ValidatePIC Function")
          If UserOption = vbNo Then
            I_DMBE_ValidatePIC = ERROR_PIC_DEVICE_HAS_NOT_BEEN_INITIALIZED
            GoTo Cleanup
          Else
            Call I_DMBE_InitializePIC
            I_DMBE_ValidatePIC = ERROR_SUCCESS ' good return
            GoTo Cleanup
          End If
        End If
      Else

'<< SCR #24298; 24307; 25142;   Developer: Brian Mowbray 07/11/2002 05:46 PM
'Patient Object is ensuring we have the correct patient so this is unecessary. The unit number could be different when transmitting between
'LDDB's as each LDDB will generate it's own Unit number.
      
'<< Begin:SCR #14052;   Developer: MAG 01/08/2003 03:30 PM
      
            If m_enmDestination = eDM_PIC Or m_enmDestination = eDM_AGGPIC Then

                'Check for Compressed
                sValArr = Split(FileName, "_")
                If Left(sValArr(1), 1) = "C" Then
                   bCompress = True
                End If
                
                'Check for Encrypt
                If Mid(sValArr(1), 2, 1) = "E" Then
                   bEncrypt = True
                End If
                
                RetCode = 0
                FilePattern = m_strAltPicPath & FileName
                sReadVal = ReadFileContents(FilePattern, bCompress, bEncrypt, RetCode)
                
                If RetCode <> 0 Then
                    I_DMBE_ValidatePIC = RetCode
                    GoTo Cleanup
                End If
                
                Set oDom = New MSXML2.DOMDocument30
                If Not oDom.loadXML(sReadVal) Then
                    I_DMBE_ValidatePIC = ERROR_PIC_DEVICE_HAS_NOT_BEEN_INITIALIZED
                    GoTo Cleanup
                End If
                Set oNode = oDom.selectSingleNode("//PICINFO/PATIENT/SSN")
                
                ' SCR #50493;   Developer: MAG 02/26/2004
                If gobjPatient.SSN <> oNode.Text And m_lEncSource <> enEncounterSigned Then
                   ' Begin:SCR #36985;   Developer: Mag 06/11/2003 02:36 PM
                   If m_enmDestination = eDM_AGGPIC Then
                     ErrMsg = "A different patient with this patient's unit number and facility already exists" & vbCrLf
                     ErrMsg = ErrMsg & "Would you like to overwrite file"
                     
                     If MsgBox(ErrMsg, vbYesNo, "Disk Validation") = vbYes Then
                       fso.DeleteFile m_strPicPath & fso.GetBaseName(FilePattern) & ".*", True
                       I_DMBE_ValidatePIC = I_DMBE_InitializePIC
                       GoTo Cleanup
                     Else
                       I_DMBE_ValidatePIC = ERROR_PIC_DEVICE_IS_NOT_FOR_CURRENT_PATIENT
                       GoTo Cleanup
                     End If
                   Else
                     ErrMsg = "You are attempting to write data belonging to:"
                     ErrMsg = ErrMsg & "Patient X to a PIC belonging to Patient Y" & vbCrLf
                     MsgBox ErrMsg, vbOKOnly, "Failed Validation"
                   
                     I_DMBE_ValidatePIC = ERROR_PIC_DEVICE_IS_NOT_FOR_CURRENT_PATIENT
                     GoTo Cleanup
                   End If
                   ' End: SCR #36985;
                End If
                
                'PIC exists and is valid
                 If m_enmDMFormat <> GetPicFormat(m_strAltPicPath, gobjPatient.unitNumber, m_enmDMFormat) Then
                   ErrMsg = "Can not overwrite PIC device (" & FileName & ") in a different format"
                   MsgBox ErrMsg, vbOKOnly, "Format Invalid"
                   I_DMBE_ValidatePIC = ERROR_INVALID_FORMAT
                   GoTo Cleanup
                 End If
                 
                 'Correct file so set Module Pic File Name
                 m_sPicFile = FileName
            End If
'            Else
'               sValArr = Split(sVal, "|", -1, vbTextCompare)
'               sReadValArr = Split(sReadVal, "|", -1, vbTextCompare)
'               If sValArr(2) = sReadValArr(2) Then ' check SSN
'                  I_DMBE_ValidatePIC = ERROR_SUCCESS ' good return
'                  Exit Function
'                  ErrMsg = "SSN matches current patient on PIC device but Name and/or DOB DO NOT Match." & vbCrLf
'                  ErrMsg = ErrMsg & Chr(9) & Chr(9) & "[PIC]" & Chr(9) & Chr(9) & "versus" & Chr(9) & "<Current Patient>" & vbCrLf
'                  ErrMsg = ErrMsg & Chr(9) & "SSN" & Chr(9) & "[" & sReadValArr(2) & "]" & Chr(9) & Chr(9) & "<" & sValArr(2) & ">" & vbCrLf
'                  ErrMsg = ErrMsg & Chr(9) & "Name" & Chr(9) & "[" & sReadValArr(0) & "]" & Chr(9) & Chr(9) & "<" & sValArr(0) & ">" & vbCrLf
'                  ErrMsg = ErrMsg & Chr(9) & "DOB" & Chr(9) & "[" & sReadValArr(1) & "]" & Chr(9) & Chr(9) & "<" & sValArr(1) & ">" & vbCrLf
'                  ErrMsg = ErrMsg & vbCrLf & "Continue (YES) or Initialize PIC device (NO) or Abort (Cancel) current activity?" & vbCrLf
'                  UserOption = MsgBox(ErrMsg, vbYesNoCancel, "Validate PIC Function")
'                  If UserOption = vbYes Then
'                     I_DMBE_ValidatePIC = ERROR_SUCCESS ' good return
'                     Exit Function
'                  End If
'                  If UserOption = vbNo Then
'                     Call I_DMBE_InitializePIC
'                     I_DMBE_ValidatePIC = ERROR_SUCCESS ' good return
'                     Exit Function
'                  End If
'                  If UserOption = vbCancel Then
'                     I_DMBE_ValidatePIC = ERROR_PIC_DEVICE_IS_NOT_FOR_CURRENT_PATIENT
'                     Exit Function
'                  End If
'                  If sValArr(0) = sReadValArr(0) Then ' check name
'                     If sValArr(1) = sReadValArr(1) Then ' check DOB
'                        I_DMBE_ValidatePIC = ERROR_SUCCESS ' good return
'                        Exit Function
'                     Else
'                        ' issue warning
'                        I_DMBE_ValidatePIC = ERROR_PIC_DEVICE_PATIENT_DOB_IS_DIFFERENT
'                        Exit Function
'                     End If
'                  Else
'                     ' issue warning
'                     UserOption = DisplayPICMsg(FileName)
'                     I_DMBE_ValidatePIC = ERROR_PIC_DEVICE_PATIENT_SSN_IS_DIFFERENT
'                     Exit Function
'                  End If

      End If
  End If
  I_DMBE_ValidatePIC = ERROR_SUCCESS ' good return
  
Cleanup:
  Set fso = Nothing
  Set oDom = Nothing
  Set oNode = Nothing
  Exit Function
  
ErrHandler:
  Set fso = Nothing
  Set oDom = Nothing
  Set oNode = Nothing
  
  Err.Raise Err.Number
End Function
' END OF DMBE INTERFACE DEFINITION OBJECT
'
'
Private Function DisplayPICMsg(FileName As String) As Long
   Dim sReadVal As String
   Dim PatientData() As String
   Dim ErrMsg As String
      
   sReadVal = ReadFileContents(m_strPicPath & FileName)
   PatientData = Split(sReadVal, "|", -1, vbTextCompare)
   ErrMsg = "This PIC device has been initialized for another patient:" & vbCrLf & vbCrLf
   ErrMsg = ErrMsg & Chr(9) & "Unit Number = " & Mid(FileName, 1, InStr(1, FileName, ".PIC", vbTextCompare) - 1) & vbCrLf
   ErrMsg = ErrMsg & Chr(9) & "Patient Name = " & PatientData(0) & vbCrLf
   ErrMsg = ErrMsg & Chr(9) & "Patient SSN = " & PatientData(2) & vbCrLf
   ErrMsg = ErrMsg & Chr(9) & "Patient DOB = " & PatientData(1) & vbCrLf & vbCrLf
   ErrMsg = ErrMsg & "Check PIC SSN with Patient SSN" & vbCrLf
   DisplayPICMsg = MsgBox(ErrMsg, vbInformation & vbCancel, "DisplayPICMsg Function")
End Function

Private Function WarningPICMsg(FileName As String) As Long
   Dim sReadVal As String
   Dim PatientData() As String
   Dim ErrMsg As String
   
   sReadVal = ReadFileContents(m_strPicPath & FileName)
   PatientData = Split(sReadVal, "|", -1, vbTextCompare)
   ErrMsg = "This PIC device matches on Unit Number and SSN but does NOT match on name and/or DOB:" & vbCrLf
   ErrMsg = ErrMsg & Chr(9) & "Unit Number = " & Mid(FileName, 1, InStr(1, FileName, ".PIC", vbTextCompare) - 1) & vbCrLf
   ErrMsg = ErrMsg & Chr(9) & "Patient Name = " & PatientData(0) & vbCrLf
   ErrMsg = ErrMsg & Chr(9) & "Patient SSN = " & PatientData(2) & vbCrLf
   ErrMsg = ErrMsg & Chr(9) & "Patient DOB = " & PatientData(1) & vbCrLf
   ErrMsg = ErrMsg & "Do you wish to continue with this card? " & vbCrLf
   ErrMsg = ErrMsg & "(Warning: Answer NO and the current process will be terminated)" & vbCrLf
   WarningPICMsg = MsgBox(ErrMsg, vbYesNo + vbExclamation, "WarningPICMsg Function")
End Function

Private Function GetAllergiesRecordsets(ByRef colRSObjects As Collection) As Boolean

    Dim strSQL As String
    Dim RS As ADODB.Recordset
    Dim objData As IDMBE.I_RSObjects
    Dim objDAS As ICHCSII_DataAccess.ICHCSII_DAS
    
    Set objDAS = oDatabase.CHCSII_DAS(Auto) 'New GEMS_DAS

    strSQL = "SELECT * FROM ALLERGIES WHERE UNIT_NUMBER = " & gobjPatient.unitNumber
    strSQL = strSQL & " AND (UPDATEFLAG <> 'D' OR UPDATEFLAG IS NULL)"
    
    Set RS = objDAS.OpenRecordset(strSQL)
    
    If Not RS Is Nothing Then
        If Not RS.EOF Then
            Set objData = New DMBE.RSObject_Ops
            With objData
                Set .DBRECORDSET = RS
                .TableName = "ALLERGIES"
            End With
            colRSObjects.Add objData, "ALLERGIES"
        Else
            Exit Function
        End If
    Else
        Exit Function
    End If
    
    strSQL = "SELECT * FROM ALLERGYREACTION WHERE UNIT_NUMBER = " & gobjPatient.unitNumber
    strSQL = strSQL & " AND (UPDATEFLAG <> 'D' OR UPDATEFLAG IS NULL)"
    
    Set RS = objDAS.OpenRecordset(strSQL)
    
    If Not RS Is Nothing Then
        If Not RS.EOF Then
            Set objData = New DMBE.RSObject_Ops
            With objData
                Set .DBRECORDSET = RS
                .TableName = "ALLERGYREACTION"
            End With
            colRSObjects.Add objData, "ALLERGYREACTION"
        End If
    End If
    
    strSQL = "SELECT * FROM ALLERGYCOMMENT WHERE UNIT_NUMBER = " & gobjPatient.unitNumber
    strSQL = strSQL & " AND (UPDATEFLAG <> 'D' OR UPDATEFLAG IS NULL)"

    Set RS = objDAS.OpenRecordset(strSQL)
    
    If Not RS Is Nothing Then
        If Not RS.EOF Then
            Set objData = New DMBE.RSObject_Ops
            With objData
                Set .DBRECORDSET = RS
                .TableName = "ALLERGYCOMMENT"
            End With
            colRSObjects.Add objData, "ALLERGYCOMMENT"
        End If
    End If
    
    GetAllergiesRecordsets = True
    
End Function


'***************************************************************************
'TO BE WORKED: ACCCEPT ACK FROM TMIP AND COORDINATE IT WITH ENTRY IN DB TABLE

Private Function TMIP_EXECUTE(colRSObjects As Collection, Optional ByVal enmTMIPRetrieval As TMIPDataRetrievalResult) As Long
   Dim RetVal As Long
   
   '<SCR 22579 CC>'
   RetVal = DMErrorCodes.ERROR_SUCCESS
   
   If Not IsMissing(enmTMIPRetrieval) Then
        Select Case enmTMIPRetrieval
            Case No_Data
                RetVal = DMErrorCodes.ERROR_NO_RECORDS_FOR_DATASET_FOUND
            Case All_Data_Have_Been_Sent
                RetVal = DMErrorCodes.ERROR_NO_DATA_TO_TRANSFER
            Case Application_Error
                RetVal = DMErrorCodes.ERROR_CLINICAL_MODULE_ERROR
        End Select
   End If
   
   If RetVal <> DMErrorCodes.ERROR_SUCCESS Then
        TMIP_EXECUTE = RetVal
        Exit Function
    End If
   '</SCR 22579 CC>'
   
   '102123
   GetTMIPFlag
    
   Select Case m_enmDataSet
   
      Case eDM_Allergies
         ' build the allergy message for TMIP
         RetVal = BuildNewTMIPAllergyMsg(colRSObjects)
         TMIP_EXECUTE = RetVal
         
         Exit Function
      Case eDM_Demographics
         ' build the Demographics message for TMIP
         RetVal = BuildTMIPDemographicsMsg(colRSObjects)
         TMIP_EXECUTE = RetVal
         Exit Function
      Case eDM_Problems
         ' build the Problems message for TMIP
         RetVal = BuildTMIPProblemsMsg(colRSObjects)
         TMIP_EXECUTE = RetVal
         Exit Function
      Case eDM_Medications
         ' build the Medications message for TMIP
         RetVal = BuildTMIPMedicationsMsg(colRSObjects)
         TMIP_EXECUTE = RetVal
         Exit Function
      Case eDM_Readiness
         ' build the Readiness message for TMIP
         RetVal = BuildTMIPReadinessMsg(colRSObjects)
         TMIP_EXECUTE = RetVal
         Exit Function
      Case eDM_Immunizations
         ' build the Immunizations message for TMIP
         RetVal = BuildTMIPImmunizationsMsg(colRSObjects)
         TMIP_EXECUTE = RetVal
         Exit Function
      Case eDM_Vitals
         ' build the Vitals message for TMIP
         RetVal = BuildTMIPVitalsMsg(colRSObjects)
         TMIP_EXECUTE = RetVal
         Exit Function
         
      'SF 62830
      Case 12 ' eDM_JMEWS
          GenerateJMEWSInpatientStatusFile colRSObjects
          TMIP_EXECUTE = DMErrorCodes.ERROR_SUCCESS
         
      Case eDM_Encounters
      
         ' build the Encounters message for TMIP
         RetVal = BuildTMIPEncountersMsg(colRSObjects)
         TMIP_EXECUTE = RetVal
         
         '>> Begin:SCR #44244;   Developer: Brian Mowbray
         If m_bEICIsDestination = False Then
          GenerateJMEWSFile colRSObjects
         End If
         '<< End : SCR #44244;
         
         
         If TMIP_EXECUTE <> DMErrorCodes.ERROR_SUCCESS Then
            Exit Function
         End If
         
         
'         << Begin:SCR #25812;   Developer: Brian Mowbray 07/18/2002 11:22 PM
'         This is no longer neccesary since the OBX is no included in the Encounter Message
'          build the Symptoms message for TMIP if possible
'         If m_bEICIsDestination = False Then
'            RetVal = BuildTMIPSymptomsMsg(colRSObjects)
'         End If
'         >> End: SCR #25812;
         
         If RetVal = ERROR_RECORDSET_IS_NOTHING Then
            RetVal = ERROR_SUCCESS
         End If
         
         If TMIP_EXECUTE <> RetVal Then
            'TMIP_EXECUTE = ERROR_SUCCESS_WITH_SELECTED_RECORDSETS
         Else
            Exit Function
         End If
      Case eDM_ClinicNotes
         ' build the Clinic Note message for TMIP
         RetVal = BuildTMIPClinicalNotesMsg(colRSObjects)
         TMIP_EXECUTE = RetVal
         Exit Function
      Case eDM_All
         TMIP_EXECUTE = ERROR_DATASET_NOT_SUPPORTED
         Exit Function
                 
      Case Else
         TMIP_EXECUTE = ERROR_DATASET_NOT_SUPPORTED
         Exit Function
   End Select
   
End Function


'***************************************************************************
'TO BE WORKED: ACCCEPT ACK FROM TMIP AND COORDINATE IT WITH ENTRY IN DB TABLE

Private Function JMEWS_EXECUTE(colRSObjects As Collection, Optional ByVal enmJMEWSRetrieval As TMIPDataRetrievalResult) As Long
   Dim RetVal As Long
   
   '<SCR 22579 CC>'
   RetVal = DMErrorCodes.ERROR_SUCCESS
   
   If Not IsMissing(enmJMEWSRetrieval) Then
        Select Case enmJMEWSRetrieval
            Case No_Data
                RetVal = DMErrorCodes.ERROR_NO_RECORDS_FOR_DATASET_FOUND
            Case All_Data_Have_Been_Sent
                RetVal = DMErrorCodes.ERROR_NO_DATA_TO_TRANSFER
            Case Application_Error
                RetVal = DMErrorCodes.ERROR_CLINICAL_MODULE_ERROR
        End Select
   End If
   
   If RetVal <> DMErrorCodes.ERROR_SUCCESS Then
        JMEWS_EXECUTE = RetVal
        Exit Function
    End If
   '</SCR 22579 CC>'
      
   Select Case m_enmDataSet
      Case eDM_Allergies
         ' build the allergy message for TMIP
'         RetVal = BuildNewTMIPAllergyMsg(colRSObjects)
'         JMEWS_Execute = RetVal
         AddToJMEWSCollection colRSObjects
         RetVal = ERROR_SUCCESS
         Exit Function
      Case eDM_Demographics
         ' build the Demographics message for TMIP
'         RetVal = BuildTMIPDemographicsMsg(colRSObjects)
'         JMEWS_Execute = RetVal
         AddToJMEWSCollection colRSObjects
         RetVal = ERROR_SUCCESS
         Exit Function
      Case eDM_Problems
         ' build the Problems message for TMIP
'         RetVal = BuildTMIPProblemsMsg(colRSObjects)
'         JMEWS_Execute = RetVal
         AddToJMEWSCollection colRSObjects
         RetVal = ERROR_SUCCESS
         Exit Function
      Case eDM_Medications
         ' build the Medications message for TMIP
'         RetVal = BuildTMIPMedicationsMsg(colRSObjects)
'         JMEWS_Execute = RetVal
         AddToJMEWSCollection colRSObjects
         RetVal = ERROR_SUCCESS
         Exit Function
      Case eDM_Readiness
         ' build the Readiness message for TMIP
'         RetVal = BuildTMIPReadinessMsg(colRSObjects)
'         JMEWS_Execute = RetVal
         RetVal = ERROR_SUCCESS
         Exit Function
      Case eDM_Immunizations
         ' build the Immunizations message for TMIP
'         RetVal = BuildTMIPImmunizationsMsg(colRSObjects)
'         JMEWS_Execute = RetVal
         RetVal = ERROR_SUCCESS
         Exit Function
      Case eDM_Vitals
         ' build the Vitals message for TMIP
'         RetVal = BuildTMIPVitalsMsg(colRSObjects)
'         JMEWS_Execute = RetVal
         RetVal = ERROR_SUCCESS
         Exit Function
         
      'SF 62830
      Case 12 ' eDM_JMEWS
          'This is the last dataset to be called so send it out.
          AddToJMEWSCollection colRSObjects
          
          GenerateJMEWSInpatientStatusFile m_JmewsRSObjects
          JMEWS_EXECUTE = DMErrorCodes.ERROR_SUCCESS
         
      Case eDM_Encounters
'         ' build the Encounters message for TMIP
'         RetVal = BuildTMIPEncountersMsg(colRSObjects)
'         JMEWS_Execute = RetVal
'
'         '>> Begin:SCR #44244;   Developer: Brian Mowbray
'         If m_bEICIsDestination = False Then
'          GenerateJMEWSFile colRSObjects
'         End If
'         '<< End : SCR #44244;
'
'
'         If JMEWS_Execute <> DMErrorCodes.ERROR_SUCCESS Then
'            Exit Function
'         End If
'
'
''         << Begin:SCR #25812;   Developer: Brian Mowbray 07/18/2002 11:22 PM
''         This is no longer neccesary since the OBX is no included in the Encounter Message
''          build the Symptoms message for TMIP if possible
''         If m_bEICIsDestination = False Then
''            RetVal = BuildTMIPSymptomsMsg(colRSObjects)
''         End If
''         >> End: SCR #25812;
'
'         If RetVal = ERROR_RECORDSET_IS_NOTHING Then
'            RetVal = ERROR_SUCCESS
'         End If
'
'         If JMEWS_Execute <> RetVal Then
'            'JMEWS_Execute = ERROR_SUCCESS_WITH_SELECTED_RECORDSETS
'         Else
'            Exit Function
'         End If
      Case eDM_ClinicNotes
         ' build the Clinic Note message for TMIP
'         RetVal = BuildTMIPClinicalNotesMsg(colRSObjects)
'         JMEWS_Execute = RetVal
         Exit Function
      Case eDM_All
'         JMEWS_Execute = ERROR_DATASET_NOT_SUPPORTED
         Exit Function
                 
      Case Else
         JMEWS_EXECUTE = ERROR_DATASET_NOT_SUPPORTED
         Exit Function
   End Select
   
End Function




Private Function GenerateJMEWSFile(ByVal colRSObjects As Collection) As Boolean
   
  Dim oXML As XMLGen
  Dim sMessageType As String
  Dim sFileName As String
  Dim sTemp As String
  Dim rsObject As IDMBE.I_RSObjects
  Dim sMessageGUID As String
  Dim rstEnc As ADODB.Recordset
  Dim vbookmrk As Variant
  
  On Error GoTo Func_Err
    
   
  'Look in the registry for the values pertaining to JMEWS
  GetJMEWSFlags
    
  'If the JMEWS flag is True, perform action to create file
  If gbJMEWS Then
    
     
    If oDas Is Nothing Then
        Set oDas = oDatabase.CHCSII_DAS(Auto) 'New GEMS_DAS
    End If
         
    '<< Begin:SCR #30824;   Developer: Brian Mowbray 12/16/2002 11:46 PM
    Set rstEnc = New ADODB.Recordset
    
    Set rstEnc = GetRS(colRSObjects, "ENCOUNTERS")
    
    On Error Resume Next
    If Not rstEnc Is Nothing Then
      If rstEnc.BOF And rstEnc.EOF Then
        GoTo Func_Exit
      Else
        rstEnc.MoveFirst
      End If
    Else
      GoTo Func_Exit
    End If
    
    On Error GoTo Func_Err
    
    Do Until rstEnc.EOF
    
      Set oXML = New XMLGen
       
      If Not (gobjPatient Is Nothing) Then
          If Not BuildPatientRst(True, False) Then
            Err.Raise 999, "BuildTMIPMsg Function", "Cannot create the Immunization Message because patient segments cannot be generated!"
            GoTo Func_Exit
          Else
            If Not g_rstPatient Is Nothing Then
              'Add the rst to the collection
              Set rsObject = New DMBE.RSObject_Ops
              rsObject.TableName = "PID"
              Set rsObject.DBRECORDSET = g_rstPatient
              colRSObjects.Add rsObject, "PID"
              Set rsObject = Nothing
            End If
          End If
      Else
        Err.Raise 999, "BuildTMIPMsg Function", "Cannot create the Immunization Message because patient segments cannot be generated!"
        GoTo Func_Exit
      End If
     
       'Get the a unique message guid
       sMessageGUID = GetNewGUID
       
       'Get the clinical module name
       sMessageType = ReturnDataSetName(m_enmDataSet)
          
       BuildMSHRst sMessageType, sMessageGUID
           
       If Not g_rstMSH Is Nothing Then
         'Add the rst to the collection
         Set rsObject = New DMBE.RSObject_Ops
         rsObject.TableName = "MSH"
         Set rsObject.DBRECORDSET = g_rstMSH
         colRSObjects.Add rsObject, "MSH"
         Set rsObject = Nothing
       End If
           
       'Add the rst to the collection
       Set rsObject = New DMBE.RSObject_Ops
       rsObject.TableName = "Providers"
       Set rsObject.DBRECORDSET = GetProviderInfo()
       colRSObjects.Add rsObject, "Providers"
       Set rsObject = Nothing
       
       'Add the rst to the collection
       Set rsObject = New DMBE.RSObject_Ops
       rsObject.TableName = "DIAGNOSIS"
       Set rsObject.DBRECORDSET = GetDiagnosis(colRSObjects, rstEnc![encounterNumber])
       colRSObjects.Add rsObject, "DIAGNOSIS"
       Set rsObject = Nothing
       
       '*************************************************************************************
       'SCR 71255
       'This will be the New Inpatient Disgnosis Recordset that will replace the Existing one
       Set rsObject = New DMBE.RSObject_Ops
       rsObject.TableName = "<~>INPATIENT_DIAGNOSIS"
       Set rsObject.DBRECORDSET = GetInpatientDiagnosis(colRSObjects, rstEnc![encounterNumber])
       colRSObjects.Add rsObject, "<~>INPATIENT_DIAGNOSIS"
       Set rsObject = Nothing
       
       vbookmrk = rstEnc.Bookmark
       
      'Build the raw xml file based on the recordsets in the collection
      oXML.BuildRawXML colRSObjects, sMessageType, False, rstEnc![encounterNumber]
      
      If Not IsEmpty(vbookmrk) Then
        rstEnc.Bookmark = vbookmrk
      End If
      
      'Tracker #55390, #55393 : DBM 06/18/2004
      'Create a unique files name using the data/time and facilityncid
      
      sTemp = gobjLogon.FacilityNCID & "_" & Format(Now, "yyyymmddhhmmss") & Format(Int(Rnd(Now) * 100000000), "00000000")
      sFileName = "CHCSIIT_XE_" & sTemp & ".xml"
      
      '*****************************************************
      'SCR 61581
      'Needs to write to JMEWS_OUT and JMEWSII_OUT
    
'SF SCR 71126 Removed the logic that saves the JMEWS 1 file.
'      If gsJMEWS_Path <> vbNullString Then
'        oXML.WriteXML gsJMEWS_Path, sFileName
'      End If
'
      If gsJMEWSII_Path <> vbNullString Then
        oXML.WriteXML gsJMEWSII_Path, sFileName
      End If
      '*****************************************************
      
      
      Set oXML = Nothing

      rstEnc.MoveNext
    Loop
    '>> End: SCR #30824;
  
  End If
  
Func_Exit:
  On Error Resume Next
  Set oXML = Nothing
  Exit Function

Func_Err:
   If Err.Number = 457 Then
      colRSObjects.Remove rsObject.TableName
      Resume
    Else
      Resume Next
   End If
   
   gobjShared.ShowVBError Err.Number, Err.Description, "GenerateJMEWSFile Function ", "DMBE", vbExclamation
   Resume Func_Exit
   Resume 'Debug
End Function

Private Function GenerateJMEWSInpatientStatusFile(ByVal colRSObjects As Collection) As Boolean
   
  Dim oXML As XMLGen
  Dim sMessageType As String
  Dim sFileName As String
  Dim sTemp As String
  Dim rsObject As IDMBE.I_RSObjects
  Dim sMessageGUID As String
  Dim rstEnc As ADODB.Recordset
  Dim vbookmrk As Variant
  
  On Error GoTo Func_Err
   
  'Look in the registry for the values pertaining to JMEWS
  GetJMEWSFlags
    
  'If the JMEWS flag is True, perform action to create file
  If gbJMEWS Then
    
     
    If oDas Is Nothing Then
        Set oDas = oDatabase.CHCSII_DAS(Auto) 'New GEMS_DAS
    End If
         
'    '<< Begin:SCR #30824;   Developer: Brian Mowbray 12/16/2002 11:46 PM
'    Set rstEnc = New ADODB.Recordset
'
'    Set rstEnc = GetRS(colRSObjects, "ENCOUNTERS")
'
'    On Error Resume Next
'    If Not rstEnc Is Nothing Then
'      If rstEnc.BOF And rstEnc.EOF Then
'        GoTo Func_Exit
'      Else
'        rstEnc.MoveFirst
'      End If
'    Else
'      GoTo Func_Exit
'    End If
'
'    On Error GoTo Func_Err
'
'    Do Until rstEnc.EOF
    
      Set oXML = New XMLGen
       
       'SF 62831, Don't build the patient header yet, no requirement for it for JMEWS Inpatient Status
'      If Not (gobjPatient Is Nothing) Then
'          If Not BuildPatientRst(True) Then
'            Err.Raise 999, "BuildTMIPMsg Function", "Cannot create the Immunization Message because patient segments cannot be generated!"
'            GoTo Func_Exit
'          Else
'            If Not g_rstPatient Is Nothing Then
'              'Add the rst to the collection
'              Set rsObject = New DMBE.RSObject_Ops
'              rsObject.TableName = "PID"
'              Set rsObject.DBRECORDSET = g_rstPatient
'              colRSObjects.Add rsObject, "PID"
'              Set rsObject = Nothing
'            End If
'          End If
'      Else
'        Err.Raise 999, "BuildTMIPMsg Function", "Cannot create the Immunization Message because patient segments cannot be generated!"
'        GoTo Func_Exit
'      End If
     
       'Get the a unique message guid
       sMessageGUID = GetNewGUID
       
       'Get the clinical module name
       sMessageType = ReturnDataSetName(m_enmDataSet)
          
       BuildMSHRst sMessageType, sMessageGUID
           
       If Not g_rstMSH Is Nothing Then
         'Add the rst to the collection
         Set rsObject = New DMBE.RSObject_Ops
         rsObject.TableName = "MSH"
         Set rsObject.DBRECORDSET = g_rstMSH
         colRSObjects.Add rsObject, "MSH"
         Set rsObject = Nothing
       End If
           
'       'Add the rst to the collection
'       Set rsObject = New DMBE.RSObject_Ops
'       rsObject.TableName = "Providers"
'       Set rsObject.DBRECORDSET = GetProviderInfo()
'       colRSObjects.Add rsObject, "Providers"
'       Set rsObject = Nothing
'
       'Add the rst to the collection
       Set rsObject = New DMBE.RSObject_Ops
       rsObject.TableName = "DIAGNOSIS"
       Set rsObject.DBRECORDSET = GetDiagnosis(colRSObjects, 0) ', rstEnc![EncounterNumber])
       'Remove the diagnosis table from the collection
       DeleteRS colRSObjects, "DIAGNOSIS"
       'Add the new RS
       colRSObjects.Add rsObject, "DIAGNOSIS"
       Set rsObject = Nothing
'
'       vbookmrk = rstEnc.Bookmark
       
      'Build the raw xml file based on the recordsets in the collection
      oXML.BuildRawXML colRSObjects, sMessageType, False ', rstEnc![EncounterNumber]
      
'      If Not IsEmpty(vbookmrk) Then
'        rstEnc.Bookmark = vbookmrk
'      End If
      
      'Tracker #55390, #55393 : DBM 06/18/2004
      'Create a unique files name using the data/time and facilityncid
      sTemp = gobjLogon.FacilityNCID & "_" & Format(Now, "yyyymmddhhmmss") & Format(Int(Rnd(Now) * 100000000), "00000000")
      sFileName = "CHCSIIT_XE_" & sTemp & ".xml"
      
      '*****************************************************
      'SCR 61581
      'Needs to write to JMEWS_OUT and JMEWSII_OUT
      
'SF SCR 71126
'      If gsJMEWS_Path <> vbNullString Then
'        oXML.WriteXML gsJMEWS_Path, sFileName
'      End If
      
      If gsJMEWSII_Path <> vbNullString Then
        oXML.WriteXML gsJMEWSII_Path, sFileName
      End If
      '*****************************************************
      
      
      Set oXML = Nothing

'      rstEnc.MoveNext
'    Loop
    '>> End: SCR #30824;
  
  End If
  
Func_Exit:
  On Error Resume Next
  Set oXML = Nothing
  Exit Function

Func_Err:
   If Err.Number = 457 Then
      colRSObjects.Remove rsObject.TableName
      Resume
    Else
      Resume Next
   End If
   
   gobjShared.ShowVBError Err.Number, Err.Description, "GenerateJMEWSFile Function ", "DMBE", vbExclamation
   Resume Func_Exit
   Resume 'Debug
End Function


Public Function GenerateJMEWSBedStatusFile(lMsgType As InpatientMsgType, ByVal sXML As String) As Long
   
    'Dim rst As ADODB.Recordset
    Dim sTemp As String
    Dim oXML As MC4_Ops
    Dim sMessageType As String
    Dim sMessageGUID As String
    Dim enmResult As DMErrorCodes
    Dim lErrresult As Long
    'Dim sXSLFile As String
    Dim sFileName As String
    Dim sFacilityID As String
    Dim sCompleteXML As String
    
    Dim oDom As MSXML.DOMDocument
    Dim oItemNode As MSXML.IXMLDOMNode
    Dim oRoot As MSXML.IXMLDOMNode
    Dim oNode As MSXML.IXMLDOMNode
    
    Set oDom = New MSXML.DOMDocument

'    oDom.Load "c:\temp\Sample_BedStatusSetup.xml"
'
'    sXML = oDom.xml

    On Error GoTo Func_Err
    
    If gobjShared Is Nothing Then
      Set gobjShared = New CWShared
    End If
    
    'Look in the registry for the values pertaining to JMEWS
    GetJMEWSFlags
    
    'If the JMEWS flag is True, perform action to create file
    If gbJMEWS = False Then
        Exit Function
    End If
    
    GenerateJMEWSBedStatusFile = ERROR_SUCCESS
    
    If Len(sXML) <= 0 Then
      GenerateJMEWSBedStatusFile = ERROR_EMPTY_RECORDSET
      Exit Function
    End If
    
    Set oXML = New MC4_Ops
   
    If Not (gobjPatient Is Nothing) Then
        If Not BuildPatientRst(False, False) Then
          Err.Raise 999, "GenerateJMEWSBedStatusFile Function", "Error creating the Patient recordset, patient segments cannot be generated!"
          GoTo Func_Exit
        End If
    Else
      Err.Raise 999, "GenerateJMEWSBedStatusFile Function", "Error creating the Patient recordset, patient segments cannot be generated!"
      GoTo Func_Exit
    End If

    sMessageGUID = GetNewGUID
    sMessageType = "INPATIENT"
       
    oXML.CreateRootNodeElement sMessageType
    
    'sXSLFile = gobjShared.RTDataPath & "\"
    
    'SCR 43003
    Select Case lMsgType
    Case enBedStatusSetup 'Bed Status Setup
      sMessageType = "BedStatusSetup"
      BuildMSHRst "Z20", sMessageGUID
      'sXSLFile = sXSLFile & "ADT_Z20.xsl"
    Case enBedStatusUpdate 'Bed Status Update
      sMessageType = "BedStatusUpdate"
      BuildMSHRst "A20", sMessageGUID
      'sXSLFile = sXSLFile & "ADT_A20.xsl"
    Case enAdmit 'Admit
      sMessageType = "Admit"
      BuildMSHRst "A01", sMessageGUID
      'sXSLFile = sXSLFile & "ADT_A01.xsl"
    Case enDischarge 'Discharge
      sMessageType = "Discharge"
      BuildMSHRst "A03", sMessageGUID
      'sXSLFile = sXSLFile & "ADT_A03.xsl"
    Case Else
      sMessageType = "Unknown"
    End Select
    
    oXML.PersistRst g_rstMSH, "MSH"
    oXML.PersistRst g_rstPatient, "PID"
    
    oXML.LoadNode sXML, "RECORDSETS"
      
    oXML.CompleteMsg "INPATIENT"
      
    sCompleteXML = oXML.rstDOM.xml
    
    If Not (gobjLogon Is Nothing) Then
      sFacilityID = gobjLogon.FacilityNCID & "_"
    Else
      sFacilityID = vbNullString
    End If
    
    sTemp = sFacilityID & Format(Now, "yyyymmddhhmmss") & Format(Int(Rnd(Now) * 100000000), "00000000")
    sFileName = sMessageType & "_XE_" & sTemp & ".xml"
    
'SF SCR 71126
'    'Needs to write to JMEWS_OUT and JMEWSII_OUT
'    If gsJMEWS_Path <> vbNullString Then
'        'oXML.WriteXML gsJMEWS_Path, sFileName
'        GenerateJMEWSBedStatusFile = WriteFileContents(gsJMEWS_Path & "\" & sFileName, sCompleteXML, False, True)
'    End If
    If gsJMEWSII_Path <> vbNullString Then
        'oXML.WriteXML gsJMEWSII_Path, sFileName
        GenerateJMEWSBedStatusFile = WriteFileContents(gsJMEWSII_Path & "\" & sFileName, sCompleteXML, False, True)
    End If
      
'    'File Path
'    sTMIPII_Path = GetPath("TMIPII_OUT")
'    sTMIPII_Path = sTMIPII_Path & "\" & sFileName
'    'WriteFileContents
'    'Don't Compress
'    'Encrypt
'    PersistTMIPMessageToFile = WriteFileContents(sTMIPII_Path, strMessage, False, True)
      
'    oXML.DoTransform sXSLFile
'
'    Set oDom = New MSXML.DOMDocument
'
'    oDom.loadXML oXML.OutputXMLString
'
'    Set oRoot = oDom.selectSingleNode("//INPATIENT_MSGS")
'
'    For Each oItemNode In oRoot.childNodes
'
'      Set oNode = oItemNode.selectSingleNode("MSH/MSH.10")
'
'      enmResult = PersistTMIPMessageToDatabase(oNode.Text, sMessageType, "", g_rstPatient.Clone, oItemNode.xml, sMessageType, False)
'
'    Next oItemNode
    
    If enmResult = ERROR_SUCCESS Then
        GenerateJMEWSBedStatusFile = ERROR_SUCCESS ' good return
    Else
        GenerateJMEWSBedStatusFile = enmResult
    End If
  
Func_Exit:
  Set oXML = Nothing
  Set oDom = Nothing
  Exit Function

Func_Err:
   gobjShared.ShowVBError Err.Number, Err.Description, "SendTMIPInpatientMsg Function", "DMBE", vbExclamation
   GenerateJMEWSBedStatusFile = ERROR_UNHANDLED_SYSTEM_ERROR
   GoTo Func_Exit
   
   Resume 'Debug
   
End Function


Private Function GetPath(ByVal sRegKeyName As String) As String

  Dim objReg As RegObj.Registry
  Dim objKey As RegObj.RegKey
  Dim strKeyName As String
  Dim objValue As RegObj.RegValue
  
  On Error Resume Next
  
  GetPath = vbNullString
  
  If gobjShared Is Nothing Then
    Set gobjShared = New CWShared
  End If

  strKeyName = gobjShared.RegistryKey & "\DataManager"
  
  Set objReg = New RegObj.Registry
  Set objKey = objReg.RegKeyFromString(strKeyName)
  
  If Not objKey Is Nothing Then
    For Each objValue In objKey.Values
      If UCase$(Trim$((objValue.Name))) = UCase$(Trim$(sRegKeyName)) Then
        GetPath = UCase$(Trim$((objValue.Value)))
        Exit Function
      End If
    Next objValue
  End If

End Function

Private Function GetProviderInfo() As ADODB.Recordset
   
   Dim sSQL As String
   On Error Resume Next
   
   'Get All providers since we could have multiple Encounters to process
   sSQL = "SELECT * FROM PROVIDER "
   
   Set GetProviderInfo = GetTableDataDAS(sSQL)

End Function

Private Function GetDiagnosis(colRS As Collection, lEncNum As Long) As ADODB.Recordset
  Dim rst As ADODB.Recordset
  Dim rstDiag As ADODB.Recordset
  Dim tempString As String
  Dim diagNCID As String
  Dim extenderName As String
  ' Build Diagnoses Data Segment and Associated Data Segments
  
  On Error Resume Next
  
  Set rst = GetRS(colRS, "DIAGNOSES")
  
  Set rstDiag = New ADODB.Recordset
  
  Call GetFieldProps("DIAGNOSES", rstDiag)
  
  If Not (rst Is Nothing) Then
         
    If lEncNum <> 0 Then
        FilterRS "ENCOUNTERNUMBER", lEncNum, rst
    End If
    
    If rst Is Nothing Then
        GoTo ExitFunction
    End If
    
    rst.MoveFirst
    '>> End: SCR #30824;
         
    If Not rst.BOF And Not rst.EOF Then
      
      Do Until rst.EOF
          rstDiag.AddNew
          rstDiag.Fields("DATAID").Value = rst![DataId]
          rstDiag.Fields("FACILITYNCID") = rst![FacilityNCID]
          diagNCID = rst![DIAGNOSISNCID]
          rstDiag.Fields("DIAGNOSISNCID") = diagNCID
          rstDiag.Fields("ENCOUNTERNUMBER") = rst![encounterNumber]
          rstDiag.Fields("CLINICIANNCID") = rst![ClinicianNCID]
          rstDiag.Fields("UNIT_NUMBER") = rst![Unit_Number]
          rstDiag.Fields("SnoID") = rst![SnoID]
          rstDiag.Fields("POINTOFCAREFACILITYNCID") = rst![PointOfCareFacilityNCID]
          rstDiag.Fields("STATUSNCID") = rst![STATUSNCID]
          rstDiag.Fields("NOTEDDATE") = rst![NotedDate]
          rstDiag.Fields("ONSETDATE") = rst![ONSETDATE]
          rstDiag.Fields("CHRONICITYNCID") = rst![CHRONICITYNCID]
          rstDiag.Fields("SOURCENCID") = rst![SOURCENCID]
          rstDiag.Fields("EXTENDEDICD9") = rst![EXTENDEDICD9]
          rstDiag.Fields("ICD9NCID") = rst![ICD9NCID]
          
          'Get the diagnosis name
          If mobjConvertSno Is Nothing Then
              Set mobjConvertSno = New actxSearchSnoMed.IConvert
          End If
          
          tempString = mobjConvertSno.SnoDescription(diagNCID)
          
          If Len(rst![EXTENDEDICD9]) > 3 Then
              extenderName = GetDiagnosisExtenderName(rst![EXTENDEDICD9])
              If Len(extenderName) > 0 Then
                 tempString = tempString & " (" & extenderName & ")"
              End If
          End If
          
          If Len(rst![DIAGNOSESCOMMENT]) > 0 Then
            tempString = tempString & ": "
          End If
          
          rstDiag.Fields("DIAGNOSESCOMMENT") = tempString & rst![DIAGNOSESCOMMENT]
          
          rstDiag.Fields("CREATEUSERNCID") = rst![CREATEUSERNCID]
          rstDiag.Fields("CREATETIME") = rst![CreateTime]
          rstDiag.Fields("CREATEDBY") = rst![CREATEDBY]
          rstDiag.Fields("CREATEDON") = rst![CREATEDON]
          rstDiag.Fields("PRIORITY") = rst![Priority]
          rstDiag.Fields("SENSITIVITY") = rst![Sensitivity]
          rstDiag.Fields("NEWFOLLOWUP") = rst![NEWFOLLOWUP]
          rstDiag.Fields("BILLINGCODENCID") = rst![BILLINGCODENCID]
          rstDiag.Fields("TMIP_STATUS") = rst![TMIP_Status]
          rstDiag.Update
         
         rst.MoveNext
      Loop
    End If
  End If
  
ExitFunction:
  
  rst.Filter = ""
  
  Set GetDiagnosis = rstDiag




End Function

'SCR 71255
'Add ICD9_CODE to Inpatient Disgnosis (JMEWSII File)
'We are creating a New "Inpatient_Diagnosis" recordset based on the Existing One.
'We are adding another Field to the Table...ICD9_CODE.
'Later, will will delete the Original RS from the Collection and just use this one.
Private Function GetInpatientDiagnosis(colRS As Collection, lEncNum As Long) As ADODB.Recordset
  Dim rst As ADODB.Recordset
  Dim rstInpDiag As ADODB.Recordset
  Dim sTemp As String
  Dim sDiagNCID As String
  Dim sICD9Code As String
  Dim sSnoID As String
  Dim sMedcinPrefix As String
  ' Build Diagnoses Data Segment and Associated Data Segments

  On Error Resume Next

  Set rst = GetRS(colRS, "INPATIENT_DIAGNOSIS")

  Set rstInpDiag = New ADODB.Recordset

  Call GetFieldPropsInpDiag("INPATIENT_DIAGNOSIS", rstInpDiag)

  If Not (rst Is Nothing) Then
  
    If rst Is Nothing Then
        GoTo ExitFunction
    End If

    rst.MoveFirst
    '>> End: SCR #30824;

    If Not rst.BOF And Not rst.EOF Then

      Do Until rst.EOF

          rstInpDiag.AddNew
          rstInpDiag.Fields("INPATIENT_ID").Value = rst![INPATIENT_ID]
          rstInpDiag.Fields("SnoID").Value = rst![SnoID]
          rstInpDiag.Fields("MEDCIN_PREFIX").Value = rst![MEDCIN_PREFIX]
          rstInpDiag.Fields("COMPLAINT_NCID").Value = rst![COMPLAINT_NCID]
          rstInpDiag.Fields("COMPLAINT_TYPE").Value = rst![COMPLAINT_TYPE]
          rstInpDiag.Fields("DESCRIPTION").Value = rst![Description]
          rstInpDiag.Fields("DIAGNOSIS_TYPE").Value = rst![DIAGNOSIS_TYPE]

          sSnoID = rst.Fields("SnoID").Value & vbNullString
          sMedcinPrefix = rst.Fields("MEDCIN_PREFIX").Value & vbNullString

          If mobjConvertSno Is Nothing Then
            Set mobjConvertSno = New actxSearchSnoMed.IConvert
          End If
    
          sICD9Code = vbNullString
          If sSnoID <> vbNullString Then
            sICD9Code = mobjConvertSno.SnoICD9(sSnoID)
          End If

          If sICD9Code = vbNullString Then
            sICD9Code = "Unknown"
          End If

          rstInpDiag.Fields("ICD9_CODE").Value = sICD9Code

          rstInpDiag.Update

         rst.MoveNext
      Loop
    End If
  End If

ExitFunction:

  rst.Filter = ""

  Set GetInpatientDiagnosis = rstInpDiag

End Function

Private Function ReturnDataSetName(lDataSet As Long) As String
  Dim sTemp As String
  
  Select Case lDataSet
  Case eDM_Demographics
    sTemp = "DEMOGRAPHICS"
  Case eDM_Allergies
    sTemp = "ALLERGIES"
  Case eDM_Problems
    sTemp = "PROBLEMS"
  Case eDM_Medications
    sTemp = "MEDICATIONS"
  Case eDM_Readiness
    sTemp = "READINESS"
  Case eDM_Immunizations
    sTemp = "IMMUNIZATIONS"
  Case eDM_Vitals
    sTemp = "VITALS"
  Case eDM_Encounters
    sTemp = "ENCOUNTERS"
  Case eDM_ClinicNotes
    sTemp = "CLINICALNOTES"
  Case eDM_Couplers
    sTemp = "COUPLERS"
  Case eDM_All
    sTemp = "ALL"
    
  'SF 62830
  Case 12 'eDM_INPATIENT
   'Replace spaces with "_" and remove "(" and ")"  They cause a problem when creating the JMEWS Inpatient node
   sTemp = "INPATIENTSTATUS-" & gobjLogon.FacilityNCID
  End Select

  ReturnDataSetName = sTemp

End Function

'<SCR 23902 CC>
Private Function BuildNewTMIPAllergyMsg(ByVal cData As Collection) As Long
   
    '102123
    'SCR 130299 - LCI - 05/16/2007: Function is used for EIC export too.
    If gbTMIP = False And m_bEICIsDestination = False Then
        BuildNewTMIPAllergyMsg = ERROR_SUCCESS
        Exit Function
    End If
    
    Dim strAllergyMessage As String
    Dim rsALLERGIES As ADODB.Recordset
    Dim rsALLERGYREACTION As ADODB.Recordset
    Dim rsALLERGYCOMMENT As ADODB.Recordset
    Dim lngAL1MsgNumber As Long
    Dim strMessageType As String
    
    Dim oHL7 As New HL7XMLBuilder
    Dim oDom As MSXML.DOMDocument
    Dim oRootNode As MSXML.IXMLDOMElement
    Dim oNodeLevel1 As MSXML.IXMLDOMElement
    Dim oNodeLevel2 As MSXML.IXMLDOMElement
    Dim blnHasReaction As Boolean
    Dim blnHasComment As Boolean
    Dim lResult As Long
    
    On Error GoTo ErrTrap
    
    Set rsALLERGIES = GetRS(cData, "ALLERGIES")
     
    lResult = CheckRst(rsALLERGIES, "")
    If lResult <> 0 Then
      BuildNewTMIPAllergyMsg = lResult
      Exit Function
    End If
    
    Set rsALLERGYREACTION = GetRS(cData, "ALLERGYREACTION")
    Set rsALLERGYCOMMENT = GetRS(cData, "ALLERGYCOMMENT")
    
    '<< Begin:SCR #26531;   Developer: Brian Mowbray 08/13/2002 17:22 PM
    'If rsALLERGYREACTION Is Nothing Then
    If Not rsALLERGYREACTION Is Nothing Then
        If rsALLERGYREACTION.RecordCount > 0 Then
            blnHasReaction = True
            rsALLERGYREACTION.MoveFirst
        End If
    End If
    
    'If rsALLERGYCOMMENT Is Nothing Then
    If Not rsALLERGYCOMMENT Is Nothing Then
        If rsALLERGYCOMMENT.RecordCount > 0 Then
            blnHasComment = True
            rsALLERGYCOMMENT.MoveFirst
        End If
    End If
    '<< Begin:SCR #26531


    ' build message from parts
    
    strMessageType = "ADT_ZA1"
    
    Set oDom = New MSXML.DOMDocument
    Set oRootNode = oDom.createElement("ADT_ZA1")
    
    If oDas Is Nothing Then
       Set oDas = oDatabase.CHCSII_DAS(Auto) 'New GEMS_DAS
    End If
    
    Dim strMessageGUID As String
    
    strMessageGUID = GetNewGUID
    
    Set oNodeLevel1 = oHL7.ReturnNewMSHNode("ZA1", "ADT", strMessageGUID)  'CStr(g_lngTMIPMsgNumber))
    oRootNode.appendChild oNodeLevel1
    Set oNodeLevel1 = oHL7.ReturnNewEVNNode("ZA1")
    oRootNode.appendChild oNodeLevel1
    
    If Not (gobjPatient Is Nothing) Then
       PIDCount = PIDCount + 1
       Set oNodeLevel1 = oHL7.ReturnNewPIDNode(PIDCount)
       oRootNode.appendChild oNodeLevel1
'       << Begin:SCR #26531;   Developer: Brian Mowbray 08/12/2002 02:22 PM
'       Set oNodeLevel1 = oHL7.ReturnTMIPDefinedZD1Node()
'       oRootNode.appendChild oNodeLevel1
'       >> End: SCR #26531;
    End If
    
    '<< Begin:SCR #26531;   Developer: Brian Mowbray 08/12/2002 02:23 PM
    ' build Patient Visit data segment for Allergies
'    lngPV1Count = lngPV1Count + 1
'    'Temp = Temp & BuildNewALGPV1Segment(lngPV1Count)
'    Set oNodeLevel1 = oHL7.ReturnNewALGPV1Node(lngPV1Count)
'    oRootNode.appendChild oNodeLevel1
    '>> End: SCR #26531;
    
    '<< Begin:SCR #26531;   Developer: Brian Mowbray 08/12/2002 02:22 PM
    '    Set oNodeLevel1 = oDom.createElement("ADT_A28.LST.4")
    '>> End: SCR #26531;
    
    Set oNodeLevel1 = oDom.createElement("ADT_ZA1.LST.1")
    
    Do Until rsALLERGIES.EOF
      
      'SCR 71415 - We donot need Deleted Allergies.
      If rsALLERGIES.Fields("UPDATEFLAG") <> "D" Then 'These have been deleted
      
        lngAL1MsgNumber = lngAL1MsgNumber + 1
        
        If blnHasComment Then
            rsALLERGYCOMMENT.Filter = "ALLERGENNCID = " & rsALLERGIES.Fields("ALLERGENNCID").Value
        End If
        
        If blnHasReaction Then
            rsALLERGYREACTION.Filter = "ALLERGENNCID = " & rsALLERGIES.Fields("ALLERGENNCID").Value
        End If
        
        Set oNodeLevel2 = oHL7.ReturnNewAL1Node(lngAL1MsgNumber, rsALLERGIES, rsALLERGYREACTION)
        oNodeLevel1.appendChild oNodeLevel2
        
        If blnHasReaction Then
            If Not (rsALLERGYREACTION.BOF And rsALLERGYREACTION.EOF) Then
                rsALLERGYREACTION.MoveFirst
            End If
        End If
        
        Set oNodeLevel2 = oHL7.ReturnNewZALNode(lngAL1MsgNumber, rsALLERGIES, rsALLERGYCOMMENT, rsALLERGYREACTION)
        oNodeLevel1.appendChild oNodeLevel2
        
        
      End If
    
      rsALLERGIES.MoveNext
      rsALLERGYCOMMENT.Filter = ""
      rsALLERGYREACTION.Filter = ""
    
    Loop
    
    oRootNode.appendChild oNodeLevel1
    
    strAllergyMessage = "<?xml version=""1.0""?>"
    strAllergyMessage = strAllergyMessage & oRootNode.xml
    
    '***********************************************
    'SCR 71415
    'Allergy Data could be written to the EIC File
    '***********************************************
    
    Dim enmResult As DMErrorCodes
    
    
    If m_bEICIsDestination = True Then
        
      m_strEICTempHL7 = m_strEICTempHL7 & strAllergyMessage
        
    Else
      
      enmResult = PersistTMIPMessageToDatabase(strMessageGUID, "ALLERGIES", "DATAID", rsALLERGIES.Clone, strAllergyMessage, strMessageType, , "TMIP_ZA1")
    
      If enmResult = ERROR_SUCCESS Then
           BuildNewTMIPAllergyMsg = ERROR_SUCCESS ' good return
      Else
           BuildNewTMIPAllergyMsg = enmResult
      End If
      
    End If

Exit Function
   
ErrTrap:
   Err.Clear
   BuildNewTMIPAllergyMsg = ERROR_UNHANDLED_SYSTEM_ERROR
End Function



Private Function BuildTMIPClinicalNotesMsg(ByVal cData As Collection) As Long
   
    '102123
    If gbTMIP = False Then
        BuildTMIPClinicalNotesMsg = ERROR_SUCCESS
        Exit Function
    End If
   
   Dim strMsg As String
   Dim strMsgType As String
   Dim ClinNoteRS As ADODB.Recordset
   Dim OBRCount As Long
   Dim OBXCount As Long
   Dim EofFlag As Boolean
   
   Dim oHL7 As New HL7XMLBuilder
   Dim oDom As MSXML.DOMDocument
   Dim oRootNode As MSXML.IXMLDOMElement
   Dim oNodeLevel1 As MSXML.IXMLDOMElement
   Dim oNodeLevel2 As MSXML.IXMLDOMElement
   Dim oNodeLevel3 As MSXML.IXMLDOMElement
   Dim oNodeLevel4 As MSXML.IXMLDOMElement
   Dim oNodeLevel5 As MSXML.IXMLDOMElement
   Dim oNodeLevel6 As MSXML.IXMLDOMElement
   Dim oNodeLevel7 As MSXML.IXMLDOMElement
   Dim strMessageGUID As String
   Dim lResult As Long
   
   On Error GoTo ErrTrap
   
   BuildTMIPClinicalNotesMsg = ERROR_SUCCESS
   
   Set ClinNoteRS = GetRS(cData, "CLINNOTE")
   
   lResult = CheckRst(ClinNoteRS, "")
   
   If lResult <> 0 Then
      BuildTMIPClinicalNotesMsg = lResult
      Exit Function
   End If
   
   If oDas Is Nothing Then
      Set oDas = oDatabase.CHCSII_DAS(Auto) 'GEMS_DAS
   End If
   
   Set oDom = New MSXML.DOMDocument
   Set oRootNode = oDom.createElement("ORU_R01")
   
   ' build message from parts
   strMsgType = "ORU_R01"
   
    strMessageGUID = GetNewGUID
   
    Set oNodeLevel1 = oHL7.ReturnNewMSHNode("R01", "ORU", strMessageGUID)
    oRootNode.appendChild oNodeLevel1
    
    Set oNodeLevel1 = oDom.createElement("ORU_R01.LST.8")
    Set oNodeLevel2 = oDom.createElement("ORU_R01.GRP.5")
    Set oNodeLevel3 = oDom.createElement("ORU_R01.GRP.4")
    
    If Not (gobjPatient Is Nothing) Then
        PIDCount = PIDCount + 1
        Set oNodeLevel4 = oHL7.ReturnNewPIDNode(PIDCount)
        oNodeLevel3.appendChild oNodeLevel4
    End If
   
    oNodeLevel2.appendChild oNodeLevel3
    Set oNodeLevel3 = oDom.createElement("ORU_R01.LST.7")
    Set oNodeLevel4 = oDom.createElement("ORU_R01.GRP.2")

    If Not (ClinNoteRS Is Nothing) Then
        EofFlag = False
        OBRCount = OBRCount + 1

        Set oNodeLevel5 = oHL7.ReturnNewCLNOBRNode(OBRCount, ClinNoteRS)
        oNodeLevel4.appendChild oNodeLevel5
        
        Set oNodeLevel5 = oDom.createElement("ORU_R01.LST.5")
        
        Do
            Set oNodeLevel6 = oDom.createElement("ORU_R01.GRP.1")
            OBXCount = OBXCount + 1
            Set oNodeLevel7 = oHL7.ReturnNewCLNOBXNode(OBRCount, ClinNoteRS)
            oNodeLevel6.appendChild oNodeLevel7
            oNodeLevel5.appendChild oNodeLevel6
            
            ClinNoteRS.MoveNext
            EofFlag = ClinNoteRS.EOF
        Loop While EofFlag = False
        
        oNodeLevel4.appendChild oNodeLevel5
        oNodeLevel3.appendChild oNodeLevel4
    End If

    oNodeLevel2.appendChild oNodeLevel3
    oNodeLevel1.appendChild oNodeLevel2
    oRootNode.appendChild oNodeLevel1
    
    strMsg = "<?xml version=""1.0""?>"
    strMsg = strMsg & oRootNode.xml
    
   Dim enmResult As DMErrorCodes
   
   enmResult = PersistTMIPMessageToDatabase(strMessageGUID, "CLINNOTE", "DATAID", ClinNoteRS, strMsg, strMsgType, , "TMIP_R01")
   
   If enmResult = ERROR_SUCCESS Then
        BuildTMIPClinicalNotesMsg = ERROR_SUCCESS ' good return
   Else
        BuildTMIPClinicalNotesMsg = enmResult
   End If

Exit Function
ErrTrap:
   Err.Clear
   BuildTMIPClinicalNotesMsg = DMErrorCodes.ERROR_UNHANDLED_SYSTEM_ERROR
End Function
Private Function BuildTMIPEncountersMsg(ByVal colRSObjects As Collection) As Long
   
'This function generates the TMIP HL7 A04 message

   Dim rsENCOUNTERS As ADODB.Recordset
   
    '102123
    'SCR 130299 - LCI - 05/16/2007: Function is used for EIC export too.
    If gbTMIP = False And m_bEICIsDestination = False Then
        'SCR 135927 Update TMIP_STATUS even if not writeing the TMIP file
        If oSQL Is Nothing Then
          Set oSQL = oDatabase.CHCSII_SQLOPS_EX(Auto) 'New GEMS_SQL.SqlOps
        End If
        
        ' Get main encounters recordset
        Set rsENCOUNTERS = GetRS(colRSObjects, "ENCOUNTERS")
        Call UpdateTMIP_Status(rsENCOUNTERS, "ENCOUNTERS", "ENCOUNTERNUMBER & FACILITYNCID")
        
        BuildTMIPEncountersMsg = ERROR_SUCCESS
        Exit Function
    End If
   
   Dim rsENC_DISPOSITION As ADODB.Recordset
   Dim rsDIAGNOSES As ADODB.Recordset
   Dim rsENC_DIAGNOSIS As ADODB.Recordset
   Dim rsDIAGNOSESRELATIONSHIPS As ADODB.Recordset
   Dim rsPROCEDURES As ADODB.Recordset
   Dim rsENC_SECTIONS As ADODB.Recordset
   Dim rsENCTEXTNOTES As ADODB.Recordset
   Dim rsPI_DATA As ADODB.Recordset
   Dim rsENC_RTFS As ADODB.Recordset
   Dim rsENC_SIGBLOCK As ADODB.Recordset
   Dim rsLISTTOOLRECORD As ADODB.Recordset
   Dim rsLISTTOOLRECORDS As ADODB.Recordset
   Dim rsCOMPLAINTS As ADODB.Recordset
   Dim rsPROBLEMS As ADODB.Recordset
   Dim rsPROBLEMSRELATIONSHIPS As ADODB.Recordset
   Dim rsCLINNOTE As ADODB.Recordset
   Dim rsSONOTES As ADODB.Recordset
   Dim rsWORK_STATUS As ADODB.Recordset
   Dim rsVITALS As ADODB.Recordset
   Dim rsVITALSIGNSVALUE As ADODB.Recordset
   Dim rsVITALSIGNSMODIFIER As ADODB.Recordset
   Dim rsREADINESS As ADODB.Recordset
   Dim rsSYMPTOMS As ADODB.Recordset
   Dim rsMedications As ADODB.Recordset
   Dim rsRXLIST As ADODB.Recordset
   Dim rsOE_RX As ADODB.Recordset
   
   Dim lngPV1Count As Long, lngZPVCount As Long, lngDG1Count As Long
   Dim lngOBXCount As Long, lngZEDCount As Long, lngPR1Count As Long
   Dim lngZESCount As Long, lngZEXCount As Long, lngZEPCount As Long
   Dim lngZERCount As Long, lngZEBCount As Long, lngZLTCount As Long
   Dim lngZLRCount As Long, lngZCPCount As Long, lngZPBCount As Long
   Dim lngZBRCount As Long, lngZCNCount As Long, lngZSOCount As Long
   Dim lngZWSCount As Long, lngZRECount As Long
   
   Dim lngPrime As Long, lngClinic As Long, lngType As Long
   
   Dim varTemp As Variant, varAdmitDate As Variant, varDischargeDate As Variant, varStatus As Variant
   
   Dim strHL7XML As String, strMessageType As String, strMessageGUID As String
   
   Dim oHL7 As New HL7XMLBuilder
   Dim oDom As MSXML.DOMDocument
   Dim oRootNode As MSXML.IXMLDOMElement
   Dim oNodeLevel1 As MSXML.IXMLDOMElement
   Dim oNodeLevel2 As MSXML.IXMLDOMElement
   Dim oNodeLevel3 As MSXML.IXMLDOMElement
   Dim oAL1Node As MSXML.IXMLDOMElement
   Dim lResult As Long
   Dim sUnitNumber As String
   Dim sSQL As String
   
   On Error GoTo ErrTrap
   
   If oSQL Is Nothing Then
     Set oSQL = oDatabase.CHCSII_SQLOPS_EX(Auto) 'New GEMS_SQL.SqlOps
   End If
   
   ' Get main encounters recordset
   Set rsENCOUNTERS = GetRS(colRSObjects, "ENCOUNTERS")
   
   lResult = CheckRst(rsENCOUNTERS, "")
   
   If lResult <> 0 Then
      BuildTMIPEncountersMsg = lResult
      Exit Function
   End If
   
   
   If oDas Is Nothing Then
      Set oDas = oDatabase.CHCSII_DAS(Auto) 'New GEMS_DAS
   End If
   
   
   '<< Begin:SCR #26641;   Developer: Brian Mowbray 08/30/2002 03:57 PM
   'Moved all code outside of loop inside so there would only be 1 encounter per message sent
   'If Not (rsENCOUNTERS Is Nothing) Then
   Do Until rsENCOUNTERS.EOF
   
      
     ' build message from parts - this will be sent regardless of whether an encounters recordset is present
     strMessageType = "ADT_A01"
     
     Set oDom = New MSXML.DOMDocument
     
     Set oRootNode = oDom.createElement(strMessageType)
     
     strMessageGUID = GetNewGUID
     
     Set oNodeLevel1 = oHL7.ReturnNewMSHNode("A04", "ADT", strMessageGUID)  ', CStr(g_lngTMIPMsgNumber))
     oRootNode.appendChild oNodeLevel1
     Set oNodeLevel1 = oHL7.ReturnNewEVNNode("A04")
     oRootNode.appendChild oNodeLevel1
     
     If Not (gobjPatient Is Nothing) Then
        PIDCount = PIDCount + 1
        Set oNodeLevel1 = oHL7.ReturnNewPIDNode(PIDCount)
        oRootNode.appendChild oNodeLevel1
        Set oNodeLevel1 = oHL7.ReturnTMIPDefinedZD1Node()
        oRootNode.appendChild oNodeLevel1
     End If
     
  ' This Allergy section is not included in the EIC transfer to date. Eventually it will be included - Brian Mowbray 7/18/2002
  '    '<SCR 22801 CC> - Add AL1 to EIC
  '    If m_enmDestination = eDM_EIC And m_enmSource = eDM_LDDB Then
  '
  '       Set oAL1Node = oDom.createElement("ADT_A01.LST.4")
  '
  '       Call GetAllergiesRecordsets(colRSObjects)
  '
  '       Dim rsALLERGIES As ADODB.Recordset
  '       Dim rsALLERGYREACTION As ADODB.Recordset
  '       Dim rsALLERGYCOMMENT As ADODB.Recordset
  '       Dim blnHasReaction As Boolean
  '       Dim blnHasComment As Boolean
  '
  '       Set rsALLERGIES = GetRS(colRSObjects, "ALLERGIES")
  '
  '       If Not rsALLERGIES Is Nothing Then
  '           Set rsALLERGYREACTION = GetRS(colRSObjects, "ALLERGYREACTION")
  '           Set rsALLERGYCOMMENT = GetRS(colRSObjects, "ALLERGYCOMMENT")
  '
  '           If Not rsALLERGYREACTION Is Nothing Then
  '               If rsALLERGYREACTION.RecordCount > 0 Then
  '                   blnHasReaction = True
  '               End If
  '           End If
  '           If Not rsALLERGYCOMMENT Is Nothing Then
  '               If rsALLERGYCOMMENT.RecordCount > 0 Then
  '                   blnHasComment = True
  '               End If
  '           End If
  '
  '           Do Until rsALLERGIES.EOF
  '
  '               If blnHasReaction = True Then
  '                   rsALLERGYREACTION.Filter = ""
  '                   rsALLERGYREACTION.MoveFirst
  '                   rsALLERGYREACTION.Filter = "ALLERGENNCID = " & rsALLERGIES.Fields("ALLERGENNCID").Value
  '               End If
  '               If blnHasComment = True Then
  '                   rsALLERGYCOMMENT.Filter = ""
  '                   rsALLERGYCOMMENT.MoveFirst
  '                   rsALLERGYCOMMENT.Filter = "ALLERGENNCID = " & rsALLERGIES.Fields("ALLERGENNCID").Value
  '               End If
  '
  '               lngAL1Count = lngAL1Count + 1
  '
  '               Set oNodeLevel2 = oHL7.ReturnNewAL1Node(lngAL1Count, rsALLERGIES, rsALLERGYREACTION)
  '               oAL1Node.appendChild oNodeLevel2
  '
  '               rsALLERGIES.MoveNext
  '           Loop
  '
  '       End If 'rsAllergies is nothing
  '
  '       'oRootNode.appendChild oNodeLevel1
  '    End If 'm_enmDestination = eDM_EIC
  '    '</SCR 22801 CC>
      
      
      
      varTemp = rsENCOUNTERS.Fields("PRIMARYPROVIDERNCID").Value
      If IsNull(varTemp) Then
         lngPrime = 0
      Else
         lngPrime = CLng(varTemp)
      End If
      varTemp = rsENCOUNTERS.Fields("CLINICNCID").Value
      If IsNull(varTemp) Then
         lngClinic = 0
      Else
         lngClinic = CLng(varTemp)
      End If
      varTemp = rsENCOUNTERS.Fields("Type").Value
      If IsNull(varTemp) Then
         lngType = 0
      Else
         lngType = CLng(varTemp)
      End If
      ' Build Patient Visit Segment
      lngPV1Count = 1
      varAdmitDate = rsENCOUNTERS.Fields("STARTDTS").Value
      If IsNull(varAdmitDate) Then
         varAdmitDate = ""
      End If
      
      Set rsENC_DISPOSITION = GetRS(colRSObjects, "ENC_DISPOSITION")
      
      Dim lngEncounterNumber As Long
      
      varTemp = rsENCOUNTERS.Fields("EncounterNumber").Value
      lngEncounterNumber = rsENCOUNTERS.Fields("EncounterNumber").Value
      
      m_EICEncounterNumber = CStr(lngEncounterNumber)
      
      
      If Not IsNull(varTemp) Then
         
         Call FilterRS("ENCOUNTERNUMBER", lngEncounterNumber, rsENC_DISPOSITION)
         If Not (rsENC_DISPOSITION Is Nothing) Then
            If rsENC_DISPOSITION.RecordCount <> 0 Then
               varStatus = rsENC_DISPOSITION.Fields("DISPOSITIONNCID").Value
               If IsNull(varStatus) Then
                  varStatus = ""
               End If
            Else
               varStatus = ""
            End If
         Else
             varStatus = ""
         End If
            
         '<< Begin:SCR #27050;   Developer: Brian Mowbray 09/06/2002 01:52 PM
         ' Get the Disposition Date to populate the TMIP Disharge Date/time
         'Using this was desided over using when the "WhenSigned" field in the encounter table since there could be a lapse of time between
         'when the encounter is signed and when the disposition is completed.
         
         Set rsENC_SECTIONS = GetRS(colRSObjects, "ENC_SECTIONS")
         Call FilterRS("ENCOUNTERNUMBER", lngEncounterNumber, rsENC_SECTIONS)
         If Not (rsENC_SECTIONS Is Nothing) Then
            If rsENC_SECTIONS.EOF = False Then
               rsENC_SECTIONS.MoveFirst
               While rsENC_SECTIONS.EOF = False
                  'Only look for the Disposition section
                  If rsENC_SECTIONS.Fields("ENC_SECTIONSINDEX") = NCID_DISPOSITION_APPLICATION Then
                    'Get the disposition date/time
                    varDischargeDate = rsENC_SECTIONS.Fields("DTS").Value
                    If IsNull(varDischargeDate) Then
                       varDischargeDate = ""
                    End If
                  End If
                  rsENC_SECTIONS.MoveNext
               Wend
            End If
         End If
         '>> End: SCR #27050;
         
         
         '<SCR 20841 CC>
         Set oNodeLevel1 = oHL7.ReturnNewPV1Node(lngPV1Count, lngPrime, lngClinic, lngType, varAdmitDate, varDischargeDate, varStatus)
         oRootNode.appendChild oNodeLevel1
         lngZPVCount = lngZPVCount + 1
         ' Build Additional Patient Visit Data Segment
         Set oNodeLevel1 = oHL7.ReturnNewZPVNode(lngZPVCount, rsENCOUNTERS, rsENC_DISPOSITION)
         oRootNode.appendChild oNodeLevel1
         'End If
         '</SCR 20841 CC>
         
        If Not oAL1Node Is Nothing Then  '<SCR 22801 CC> AL1 node should be appended only once
            oRootNode.appendChild oAL1Node
            Set oAL1Node = Nothing
        End If
         
        '<< Begin:SCR #25812;   Developer: Brian Mowbray 07/18/2002 11:20 PM
        ' Uncommented this code to allow the symptoms to be included in the A04 message
        ' Build Symptoms Segment (OBX)
        '<< Begin:SCR #26545;   Developer: Darren Magnuson 09/05/2002 2:20 PM
        
        Set oNodeLevel1 = oDom.createElement("ADT_A01.LST.3") '- To be used for OBX

        Set rsSYMPTOMS = GetRS(colRSObjects, "SYMPTOMS")
        Call FilterRS("ENCOUNTERNUMBER", varTemp, rsSYMPTOMS)
        If Not (rsSYMPTOMS Is Nothing) Then
           If rsSYMPTOMS.EOF = False Then
              rsSYMPTOMS.MoveFirst
              While rsSYMPTOMS.EOF = False
                 lngOBXCount = lngOBXCount + 1
                 Set oNodeLevel2 = oHL7.ReturnNewSYMOBXNode(lngOBXCount, rsSYMPTOMS)
                 oNodeLevel1.appendChild oNodeLevel2
                 rsSYMPTOMS.MoveNext
              Wend
           End If
        End If
        
        oRootNode.appendChild oNodeLevel1

        '>> End: SCR #25812, 26545;
         
         
         ' Build Diagnoses Data Segment and Associated Data Segments
         Set rsDIAGNOSES = GetRS(colRSObjects, "DIAGNOSES")
         Call FilterRS("ENCOUNTERNUMBER", lngEncounterNumber, rsDIAGNOSES)
         
         '<SCR 24239 CC> 'NGIT SPR 477 - SCR 24239 CC
         If Not (rsDIAGNOSES Is Nothing) Then
         
            Set rsREADINESS = GetRS(colRSObjects, "READINESS")
            If Not rsREADINESS Is Nothing Then
               If rsREADINESS.RecordCount > 0 Then
                  rsREADINESS.Filter = "ENCOUNTER_NUMBER = " & lngEncounterNumber
               End If
            Else
                'Dim strSQL As String
                'strSQL = "SELECT * FROM READINESS WHERE Patient_Unit_Number = 128548 AND Temp_Profile_Limitation = 'This is a test 2'"
                'Set rsREADINESS = oDas.OpenRecordset(strSQL)
                'MsgBox rsREADINESS.RecordCount
            End If
            
            If rsDIAGNOSES.EOF = False Then
                
               'SCR 24614 CC
               Set oNodeLevel3 = oDom.createElement("ADT_A01.LST.5") '- To be used for DG1
               
               rsDIAGNOSES.MoveFirst
               
               While rsDIAGNOSES.EOF = False
                  lngDG1Count = lngDG1Count + 1
                  Set oNodeLevel1 = oHL7.ReturnNewDG1Node(lngDG1Count, rsDIAGNOSES)
                  oNodeLevel3.appendChild oNodeLevel1 'SCR 24614 CC
                  
                  
                  'Set oNodeLevel1 = oHL7.ReturnTMIPDefinedZD1Node()
                  'oRootNode.appendChild oNodeLevel1
                  Set oNodeLevel1 = oHL7.ReturnNewZD2Node(lngDG1Count, rsDIAGNOSES)
                  oRootNode.appendChild oNodeLevel1
                  
                  Set rsDIAGNOSESRELATIONSHIPS = GetRS(colRSObjects, "DIAGNOSESRELATIONSHIPS")
                  Call FilterRS("DIAGNOSISDATAID", rsDIAGNOSES.Fields("DATAID").Value, rsDIAGNOSESRELATIONSHIPS)
                  
                  ' add check for nothing, BOF, and EOF
                  If Not rsDIAGNOSESRELATIONSHIPS Is Nothing Then
                     If rsDIAGNOSESRELATIONSHIPS.BOF = False And rsDIAGNOSESRELATIONSHIPS.EOF = False Then
                         Set oNodeLevel1 = oHL7.ReturnNewZDRNode(lngDG1Count, rsDIAGNOSESRELATIONSHIPS)
                         oRootNode.appendChild oNodeLevel1
                     End If
                  End If
                  
                  rsDIAGNOSES.MoveNext
                Wend
            
                oRootNode.appendChild oNodeLevel3 'Append DG1 under ADT_A01.LST.5
               
               '<< Begin:SCR #27593;   Developer: Brian Mowbray 09/12/2002 12:45 PM
                If Not (rsENC_DISPOSITION Is Nothing) Then
                   If rsENC_DISPOSITION.EOF = False Then
                      Set oNodeLevel1 = oHL7.ReturnNewZG1Node(rsENC_DISPOSITION, rsREADINESS)
                      oRootNode.appendChild oNodeLevel1
                      rsENC_DISPOSITION.MoveNext
                   End If
                End If
               '<< End:SCR #27593
                
                If Not rsREADINESS Is Nothing Then 'SCR 22801 CC
                    rsREADINESS.Filter = ""
                End If
            
            End If
         End If
         

         ' Build Encounter Diagnosis Segment
         'Set rsENC_DIAGNOSIS = GetRS(colRSObjects, "ENC_DIAGNOSES"
         Set rsENC_DIAGNOSIS = GetRS(colRSObjects, "ENC_DIAGNOSIS")
         Call FilterRS("ENCOUNTERNUMBER", lngEncounterNumber, rsENC_DIAGNOSIS)
         If Not (rsENC_DIAGNOSIS Is Nothing) Then
            If rsENC_DIAGNOSIS.EOF = False Then
               rsENC_DIAGNOSIS.MoveFirst
               While rsENC_DIAGNOSIS.EOF = False
                  lngZEDCount = lngZEDCount + 1
                  Set oNodeLevel1 = oHL7.ReturnNewZEDNode(lngZEDCount, rsENC_DIAGNOSIS)
                  oRootNode.appendChild oNodeLevel1
                  rsENC_DIAGNOSIS.MoveNext
               Wend
            End If
         End If
         
         ' Build Procedures Data Segment
         Set rsPROCEDURES = GetRS(colRSObjects, "PROCEDURES")
         '<< Begin:SCR #26688;   Developer: Brian Mowbray 08/16/2002 12:21 PM
         Call FilterRS("ENCOUNTERNUMBER", lngEncounterNumber, rsPROCEDURES)
         '>> End: SCR #26688;
         If Not (rsPROCEDURES Is Nothing) Then
            If rsPROCEDURES.EOF = False Then
               rsPROCEDURES.MoveFirst
               
               '<< Begin:SCR #26641;   Developer: Brian Mowbray 08/15/2002 06:36 PM
               Set oNodeLevel1 = oDom.createElement("ADT_A01.LST.7") 'SCR 24614 CC
                    Set oNodeLevel2 = oDom.createElement("ADT_A01.GRP.1")
               
                    While rsPROCEDURES.EOF = False
                      lngPR1Count = lngPR1Count + 1
                      
                      Set oNodeLevel3 = oHL7.ReturnNewPR1Node(lngPR1Count, rsPROCEDURES)
                      oNodeLevel2.appendChild oNodeLevel3
                      rsPROCEDURES.MoveNext
                    Wend
                    
                    oNodeLevel1.appendChild oNodeLevel2
                oRootNode.appendChild oNodeLevel1 ''SCR 24614 CC
            
                rsPROCEDURES.MoveFirst
                lngPR1Count = 0
            
                While rsPROCEDURES.EOF = False
                  lngPR1Count = lngPR1Count + 1
                  
                  Set oNodeLevel1 = oHL7.ReturnNewZP1Node(lngPR1Count, rsPROCEDURES)
                  oRootNode.appendChild oNodeLevel1
                  rsPROCEDURES.MoveNext
                Wend
               '>> End: SCR #26641;
            
            End If
         End If
         
         ' Build Encounter Sections Segment
         Set rsENC_SECTIONS = GetRS(colRSObjects, "ENC_SECTIONS")
         Call FilterRS("ENCOUNTERNUMBER", lngEncounterNumber, rsENC_SECTIONS)
         If Not (rsENC_SECTIONS Is Nothing) Then
            If rsENC_SECTIONS.EOF = False Then
               rsENC_SECTIONS.MoveFirst
               While rsENC_SECTIONS.EOF = False
                  lngZESCount = lngZESCount + 1
                  Set oNodeLevel1 = oHL7.ReturnNewZESNode(lngZESCount, rsENC_SECTIONS)
                  oRootNode.appendChild oNodeLevel1
                  rsENC_SECTIONS.MoveNext
               Wend
            End If
         End If
         
         ' Build Encounter Extended Notes Segment
         Set rsENCTEXTNOTES = GetRS(colRSObjects, "ENCTEXTNOTES")
         Call FilterRS("ENCOUNTERNUMBER", lngEncounterNumber, rsENCTEXTNOTES)
         If Not (rsENCTEXTNOTES Is Nothing) Then
            If rsENCTEXTNOTES.EOF = False Then
               rsENCTEXTNOTES.MoveFirst
               While rsENCTEXTNOTES.EOF = False
                  lngZEXCount = lngZEXCount + 1
                  Set oNodeLevel1 = oHL7.ReturnNewZEXNode(lngZEXCount, rsENCTEXTNOTES)
                  oRootNode.appendChild oNodeLevel1
                  rsENCTEXTNOTES.MoveNext
               Wend
            End If
         End If
         
         ' Build Encounter Patient Information Segment
         Set rsPI_DATA = GetRS(colRSObjects, "PI_DATA")
         Call FilterRS("ENC_NUM", lngEncounterNumber, rsPI_DATA)
         If Not (rsPI_DATA Is Nothing) Then
            If rsPI_DATA.EOF = False Then
               rsPI_DATA.MoveFirst
               While rsPI_DATA.EOF = False
                  lngZEPCount = lngZEPCount + 1
                  Set oNodeLevel1 = oHL7.ReturnNewZEPNode(lngZEPCount, rsPI_DATA)
                  oRootNode.appendChild oNodeLevel1
                  rsPI_DATA.MoveNext
               Wend
            End If
         End If
         
         ' Build Encounter OE_RX Segment
         Set rsOE_RX = GetRS(colRSObjects, "OE_RX")
         Call FilterRS("ENC_NUM", lngEncounterNumber, rsOE_RX)
         If Not (rsOE_RX Is Nothing) Then
            If rsOE_RX.EOF = False Then
               rsOE_RX.MoveFirst
               While rsOE_RX.EOF = False
                  lngZEPCount = lngZEPCount + 1
                  Set oNodeLevel1 = oHL7.ReturnNewZEPNode_OE_RX(lngZEPCount, rsOE_RX)
                  oRootNode.appendChild oNodeLevel1
                  rsOE_RX.MoveNext
               Wend
            End If
         End If

         
         ' Build Encounter RTF Segment and Encounter Signature Block Segment
         Set rsENC_RTFS = GetRS(colRSObjects, "ENC_RTFS")
         Call FilterRS("ENCOUNTERNUMBER", lngEncounterNumber, rsENC_RTFS)
         If Not (rsENC_RTFS Is Nothing) Then
            If rsENC_RTFS.EOF = False Then
               rsENC_RTFS.MoveFirst
               While rsENC_RTFS.EOF = False
                  lngZERCount = lngZERCount + 1
                  Set oNodeLevel1 = oHL7.ReturnNewZERNode(lngZERCount, rsENC_RTFS)
                  oRootNode.appendChild oNodeLevel1
                  Set rsENC_SIGBLOCK = GetRS(colRSObjects, "ENC_SIGBLOCK")
                  Call FilterRS("ENCOUNTERNUMBER", varTemp, rsENC_SIGBLOCK)
                  If Not (rsENC_SIGBLOCK Is Nothing) Then
                     If rsENC_SIGBLOCK.EOF = False Then
                        rsENC_SIGBLOCK.MoveFirst
                        While rsENC_SIGBLOCK.EOF = False
                           lngZEBCount = lngZEBCount + 1
                           Set oNodeLevel1 = oHL7.ReturnNewZEBNode(lngZEBCount, rsENC_SIGBLOCK)
                           oRootNode.appendChild oNodeLevel1
                           rsENC_SIGBLOCK.MoveNext
                        Wend
                     End If
                  End If
                  rsENC_RTFS.MoveNext
               Wend
            End If
         End If
         
         ' Build Encounter List Tool Record Segment and Encounter List Tool Records Segment
         Set rsLISTTOOLRECORD = GetRS(colRSObjects, "LISTTOOLRECORD")
         Call FilterRS("ENCOUNTERNUMBER", lngEncounterNumber, rsLISTTOOLRECORD)
         If Not (rsLISTTOOLRECORD Is Nothing) Then
            If rsLISTTOOLRECORD.EOF = False Then
               rsLISTTOOLRECORD.MoveFirst
               While rsLISTTOOLRECORD.EOF = False
                  lngZLTCount = lngZLTCount + 1
                  Set oNodeLevel1 = oHL7.ReturnNewZLTNode(lngZLTCount, rsLISTTOOLRECORD)
                  oRootNode.appendChild oNodeLevel1
                  Set rsLISTTOOLRECORDS = GetRS(colRSObjects, "LISTTOOLRECORDS")
                  Call FilterRS("DATAID", rsLISTTOOLRECORD.Fields("DATAID").Value, rsLISTTOOLRECORDS)
                  If Not (rsLISTTOOLRECORDS Is Nothing) Then
                     If rsLISTTOOLRECORDS.EOF = False Then
                        rsLISTTOOLRECORDS.MoveFirst
                        While rsLISTTOOLRECORDS.EOF = False
                           lngZLRCount = lngZLRCount + 1
                           Set oNodeLevel1 = oHL7.ReturnNewZLRNode(lngZLRCount, rsLISTTOOLRECORDS)
                           oRootNode.appendChild oNodeLevel1
                           rsLISTTOOLRECORDS.MoveNext
                        Wend
                     End If
                  End If
                  rsLISTTOOLRECORD.MoveNext
               Wend
            End If
         End If
         
         ' Build Complaint Segment
         Set rsCOMPLAINTS = GetRS(colRSObjects, "COMPLAINTS")
         Call FilterRS("ENCOUNTERNUMBER", lngEncounterNumber, rsCOMPLAINTS)
         If Not (rsCOMPLAINTS Is Nothing) Then
            If rsCOMPLAINTS.EOF = False Then
               rsCOMPLAINTS.MoveFirst
               While rsCOMPLAINTS.EOF = False
                  lngZCPCount = lngZCPCount + 1
                  Set oNodeLevel1 = oHL7.ReturnNewZCPNode(lngZCPCount, rsCOMPLAINTS)
                  oRootNode.appendChild oNodeLevel1
                  rsCOMPLAINTS.MoveNext
               Wend
            End If
         End If
         
        ' Begin:SCR #28326;   Developer: MAG 10/16/2002 12:33 PM
        'Build Medication Segment
        Set rsMedications = GetRS(colRSObjects, "MEDICATIONS")
        Call FilterRS("ENC_NUM", lngEncounterNumber, rsMedications)
        If Not (rsMedications Is Nothing) Then
            If rsMedications.EOF = False Then
                rsMedications.MoveFirst
                Set oNodeLevel1 = oDom.createElement("RDE_O01.LST.14")
                Do Until rsMedications.EOF
                  Set oNodeLevel2 = oDom.createElement("RDE_O01.GRP.5")
                      Set oNodeLevel3 = oHL7.ReturnNewORCNode(rsMedications)
                      oNodeLevel2.appendChild oNodeLevel3
                      Set oNodeLevel3 = oHL7.ReturnNewRXENode(rsMedications)
                      oNodeLevel2.appendChild oNodeLevel3
                  
                      lngZRECount = lngZRECount + 1
                      Set rsRXLIST = GetRS(colRSObjects, "RX_LIST")
                      Call FilterRS("DRUG_IEN", rsMedications("MEDICATIONNCID"), rsRXLIST)
                      Set oNodeLevel3 = oHL7.ReturnNewZRENode(lngZRECount, rsMedications, rsRXLIST)
                      oNodeLevel2.appendChild oNodeLevel3
                  oNodeLevel1.appendChild oNodeLevel2
                  rsMedications.MoveNext
                Loop
            
                oRootNode.appendChild oNodeLevel1
            End If
        End If
        ' End: SCR #28326;
          
         ' Build Problem Segment and Problem Relationship Segment
         Set rsPROBLEMS = GetRS(colRSObjects, "PROBLEMS")
         'Call FilterRS("ENCOUNTERNUMBER", lngEncounterNumber, rsPROBLEMS)
         If Not (rsPROBLEMS Is Nothing) Then
            If rsPROBLEMS.EOF = False Then
               rsPROBLEMS.MoveFirst
               While rsPROBLEMS.EOF = False
                  lngZPBCount = lngZPBCount + 1
                  Set oNodeLevel1 = oHL7.ReturnNewZPBNode(lngZPBCount, rsPROBLEMS)
                  oRootNode.appendChild oNodeLevel1
                  varTemp = rsPROBLEMS.Fields("DATAID").Value
                  
                  Set rsPROBLEMSRELATIONSHIPS = GetRS(colRSObjects, "PROBLEMSRELATIONSHIPS")
                  Call FilterRS("PROBLEMDATAID", varTemp, rsPROBLEMSRELATIONSHIPS)
                  If Not rsPROBLEMSRELATIONSHIPS Is Nothing Then
                     If rsPROBLEMSRELATIONSHIPS.EOF = False Then
                        rsPROBLEMSRELATIONSHIPS.MoveFirst
                        While rsPROBLEMSRELATIONSHIPS.EOF = False
                            lngZBRCount = lngZBRCount + 1
                            Set oNodeLevel1 = oHL7.ReturnNewZBRNode(lngZBRCount, rsPROBLEMSRELATIONSHIPS)
                            oRootNode.appendChild oNodeLevel1
                            rsPROBLEMSRELATIONSHIPS.MoveNext
                        Wend
                    End If
                 End If
                 sUnitNumber = rsPROBLEMS("UNIT_NUMBER") & vbNullString
                 sSQL = "Update Problems set TMIP_STATUS = 1"
                 sSQL = sSQL & " Where UNIT_NUMBER = " & sUnitNumber
                 sSQL = sSQL & " And DATAID = " & varTemp
                 
                 oSQL.Execute sSQL
                 rsPROBLEMS.MoveNext
               Wend
            End If
         End If
         
         ' Build Clinician Notes Segment
         Set rsCLINNOTE = GetRS(colRSObjects, "CLINNOTE")
         Call FilterRS("ENCOUNTER", lngEncounterNumber, rsCLINNOTE)
         If Not (rsCLINNOTE Is Nothing) Then
            If rsCLINNOTE.EOF = False Then
               rsCLINNOTE.MoveFirst
               While rsCLINNOTE.EOF = False
                  lngZCNCount = lngZCNCount + 1
                  Set oNodeLevel1 = oHL7.ReturnNewZCNNode(lngZCNCount, gobjPatient, rsCLINNOTE)
                  oRootNode.appendChild oNodeLevel1
                  rsCLINNOTE.MoveNext
               Wend
            End If
         End If
         
         ' Build SO Notes Segment
         Set rsSONOTES = GetRS(colRSObjects, "SONOTES")
         Call FilterRS("ENCOUNTERNUMBER", lngEncounterNumber, rsSONOTES)
         If Not (rsSONOTES Is Nothing) Then
            If rsSONOTES.EOF = False Then
               rsSONOTES.MoveFirst
               While rsSONOTES.EOF = False
                  lngZSOCount = lngZSOCount + 1
                  Set oNodeLevel1 = oHL7.ReturnNewZSONode(lngZSOCount, rsSONOTES)
                  oRootNode.appendChild oNodeLevel1
                  rsSONOTES.MoveNext
               Wend
            End If
         End If
         
         '<SCR 22838 CC>
         
         'Build vital segment
         Dim lngCount As Long
                 
         Set rsVITALS = GetRS(colRSObjects, "VITALS")
         Set rsVITALSIGNSVALUE = GetRS(colRSObjects, "VITALSIGNSVALUE")
         Set rsVITALSIGNSMODIFIER = GetRS(colRSObjects, "VITALSIGNSMODIFIER")
         lngCount = 1
         
         If Not rsVITALS Is Nothing Then
             If rsVITALS.RecordCount > 0 Then
                 
                 rsVITALS.MoveFirst
                 rsVITALS.Filter = "ENCOUNTERNUMBER = " & lngEncounterNumber
                 
                 Do Until rsVITALS.EOF
             
                       If Not rsVITALSIGNSVALUE Is Nothing Then
                            rsVITALSIGNSVALUE.Filter = ""
                            If rsVITALSIGNSVALUE.RecordCount > 0 Then
                                rsVITALSIGNSVALUE.MoveFirst
                                rsVITALSIGNSVALUE.Filter = "DATAID = " & rsVITALS.Fields("DATAID").Value
                            End If
                       End If
                       
                       If Not rsVITALSIGNSMODIFIER Is Nothing Then
                            rsVITALSIGNSMODIFIER.Filter = ""
                            If rsVITALSIGNSMODIFIER.RecordCount > 0 Then
                                rsVITALSIGNSMODIFIER.MoveFirst
                                rsVITALSIGNSMODIFIER.Filter = "DATAID = " & rsVITALS.Fields("DATAID").Value
                            End If
                       End If
                       
                       lngCount = lngCount + 1
                       
                       Set oNodeLevel1 = oHL7.ReturnNewZVTNode(lngCount, rsVITALS, rsVITALSIGNSMODIFIER, rsVITALSIGNSVALUE)
                       oRootNode.appendChild oNodeLevel1
                       rsVITALS.MoveNext
                     
                 Loop
             End If
         End If
         
         
         'Build READINESS segment
         'Set rsREADINESS = GetRS(colRSObjects, "READINESS")
'            If Not rsREADINESS Is Nothing Then
'                lngCount = 1
'                Do Until rsREADINESS.EOF
'                    Set oNodeLevel1 = oHL7.ReturnNewZRDNode(lngCount, rsREADINESS)
'                    oRootNode.appendChild oNodeLevel1
'                    lngCount = lngCount + 1
'                Loop
'            End If
         
         ' Build Work varStatus Segment
         Set rsWORK_STATUS = GetRS(colRSObjects, "WORK_STATUS")
         Call FilterRS("ENCOUNTER_NUMBER", lngEncounterNumber, rsWORK_STATUS)
         If Not (rsWORK_STATUS Is Nothing) Then
            If rsWORK_STATUS.EOF = False Then
               rsWORK_STATUS.MoveFirst
               While rsWORK_STATUS.EOF = False
                  lngZWSCount = lngZWSCount + 1
                  Set oNodeLevel1 = oHL7.ReturnNewZWSNode(lngZWSCount, rsWORK_STATUS)
                  oRootNode.appendChild oNodeLevel1
                  rsWORK_STATUS.MoveNext
               Wend
            End If
         End If
      End If
      
     strHL7XML = "<?xml version=""1.0""?>"
     strHL7XML = strHL7XML & oRootNode.xml
     
     'Call CreateOutputFile("Encounters", strHL7XML)
        
     ' Update the TMIP_Status
     Call UpdateTMIP_Status(rsENCOUNTERS, "ENCOUNTERS", "ENCOUNTERNUMBER & FACILITYNCID")
     
     If m_bEICIsDestination = True Then
       
       m_strEICTempHL7 = m_strEICTempHL7 & strHL7XML
     
     Else 'This is a TMIP Transaction
       
       Dim enmResult As DMErrorCodes

       enmResult = PersistTMIPMessageToDatabase(strMessageGUID, "ENCOUNTERS", "ENCOUNTERNUMBER & FACILITYNCID", rsENCOUNTERS.Clone, strHL7XML, strMessageType, , "TMIP_A04")
       
       If enmResult = ERROR_SUCCESS Then
           BuildTMIPEncountersMsg = ERROR_SUCCESS ' good return
       Else
           BuildTMIPEncountersMsg = enmResult
       End If
       
     End If
     
     
     rsENCOUNTERS.MoveNext
   
   Loop 'Until rsENCOUNTERS.EOF
   '>> End: SCR #26641;

  '<< Begin:SCR #30953;   Developer: Brian Mowbray 12/18/2002 02:01 PM
  'Reset the EncounterSource flag so that it will default to waiting for the EventServer unless set by the Signed Encounter Event to not wait
  m_lEncSource = 0
  '>> End: SCR #30953;

Exit Function

ErrTrap:
  Err.Clear
  BuildTMIPEncountersMsg = ERROR_UNHANDLED_SYSTEM_ERROR
End Function


'***************************************************************
'New Function
'Trackers: 43003, 48315, 48316, 48318, 48319, 48320, 48321
'TMIP Framework 2
'***************************************************************                                                                                                                                                                                                                         '123519
Private Function PersistTMIPMessageToFile(ByVal strMSH10GUID As String, ByVal strTableName As String, ByVal strPKColumnName As String, ByVal rsClinRecord As ADODB.Recordset, ByVal strMessage As String, ByVal strMessageType As String, Optional bUpdateTMIP_Status As Boolean = True, Optional HL7Type As String = "") As DMErrorCodes

On Error GoTo ErrHandler

Dim sFileName As String
Dim sTemp As String
Dim sTMIPII_Path As String
Dim sFacilityID As String
Dim sPatientUnitNumber As String

  '****************************************************
  'SCR 62364 - Add Patient Unit Number to the File Name
  '****************************************************
  
  'FileName
  'SCR 43003
  If Not (gobjLogon Is Nothing) Then
    sFacilityID = gobjLogon.FacilityNCID & "_"
  Else
    sFacilityID = vbNullString
  End If
  
  sPatientUnitNumber = vbNullString
  If Not (gobjPatient Is Nothing) Then
    If gobjPatient.unitNumber <> vbNull Then
      sPatientUnitNumber = gobjPatient.unitNumber & "_"
    End If
  End If
  
  'sTemp = sFacilityID & Format(Now, "yyyymmddhhmmss") & Format(Int(Rnd(Now) * 100000000), "00000000")
  sTemp = sFacilityID & sPatientUnitNumber & Format(Now, "yyyymmddhhmmss") & Format(Int(Rnd(Now) * 100000000), "00000000")
                '123519
  If HL7Type <> "" Then
    HL7Type = HL7Type & "_"
  End If
  sFileName = Replace(HL7Type, "TMIP_", "") & strTableName & "_XE_" & sTemp & ".xml"
  
  'File Path
  sTMIPII_Path = GetPath("TMIP_OUT")
  sTMIPII_Path = sTMIPII_Path & "\" & sFileName
  'WriteFileContents
  'Don't Compress
  'Encrypt
  
  '**************************************************
  PersistTMIPMessageToFile = WriteFileContents(sTMIPII_Path, strMessage, False, True)
  '**************************************************
  
  '****************************
  'SCR 63969
  If PersistTMIPMessageToFile = ERROR_SUCCESS And bUpdateTMIP_Status Then
    If UpdateTMIP_Status(rsClinRecord, strTableName, strPKColumnName) = False Then
      MsgBox "An error occured when attempting to update the TMIP Status Flag. Please retry sending this message " & _
          "by using Data Manager to send the message again manually", vbCritical, "TMIP Status Update Error"
      
      PersistTMIPMessageToFile = ERROR_UPDATING_TMIP_STATUS_FLAG
    End If
  End If
  '****************************

Exit Function

ErrHandler:
  gobjShared.ShowVBError Err.Number, Err.Description, "PersistTMIPMessageToFile Function", "DMBE", vbExclamation
  PersistTMIPMessageToFile = ERROR_UNHANDLED_SYSTEM_ERROR

End Function

'Framework 1
'XML String being written to a database.
'Framework 2
'XML Written to a file somewhere.
Private Function PersistTMIPMessageToDatabase(ByVal strMSH10GUID As String, _
    ByVal strTableName As String, ByVal strPKColumnName As String, _
    ByVal rsClinRecord As ADODB.Recordset, ByVal strMessage As String, _
    ByVal strMessageType As String, Optional bUpdateTMIP_Status As Boolean = True, _
    Optional ByVal HL7Type As String = "") As DMErrorCodes

  Dim sSQL As String
  Dim sCompName As String
  Dim oCmd As ADODB.Command
  Dim oParam As ADODB.Parameter
  Dim nTMIPFlag As Integer
  On Error GoTo ErrHandler
   
  nTMIPFlag = TMIPFrameworkFlag

  'SCR 120555
  If HL7Type <> "" Then
    If Message_Ops.GetHL7MessageFlag(HL7Type) = False Then
        PersistTMIPMessageToDatabase = ERROR_SUCCESS
        Exit Function
    End If
  End If
  
  If nTMIPFlag <> 1 Then
  
    'We need to write something to a file somewhere.                                                                                                                   '123519
    PersistTMIPMessageToDatabase = PersistTMIPMessageToFile(strMSH10GUID, strTableName, strPKColumnName, rsClinRecord, strMessage, strMessageType, bUpdateTMIP_Status, HL7Type)
  
  Else
    If oSQL Is Nothing Then
      Set oSQL = oDatabase.CHCSII_SQLOPS_EX(Auto) 'New GEMS_SQL.SqlOps
    End If
    
    'The current Micorosft OLEDB Oracle provider used does not support LOB's, CLOB's or BLOB's. It only supports the LONG and LONG RAW Oracle
    'Data type. We cannot insert a value directly into a CLOB datatype as needed so we insert it first into the LONG then call the next SQL
    'Statement to perform the conversion using 'TO_LOB' to get it into the CLOB data type.
    
    sCompName = Trim$(Environ("Computername") & vbNullString)
      
    Call oSQL.Execute("Select * From TMIP_COORDINATOR Where TMIP_MESSAGE_GUID = '0'")
      
    oSQL.AddNew
      
    oSQL("TMIP_MESSAGE_GUID") = strMSH10GUID
    
    oSQL("TMIP_MESSAGE_TYPE") = IIf(Len(strMessageType) > 0, strMessageType, "ADT_UNK")
  
    Call oSQL.AppendAsChunk("XML_MESSAGE_BODY", strMessage)
              
    If oSQL.Update = True Then
      
    ' Begin:SCR #36878;   Developer: Brian Mowbray 06/02/2003 07:21 PM
    
    'This SQL statement will convert the LONG from the TMIP_Coordinator table and convert it to a CLOB prior to inserting into the
    'TMIP_INTERFACE_STORAGE table. This was needed instead of the stored procedure for the following reasons:
    'Oracle 8.1.7 will not intrinsically convert the LONG to a CLOB as we were able to do in 8.1.6. Instead of using packages or other
    'conversion mechanisms the 'TO_LOB' is a built in Oracle SQL function that cannot be called in a PL/SQL block and can only be called as
    'an INSERT statement with the nested SELECT.
      
     sSQL = "INSERT INTO TMIP_INTERFACE_STORAGE (cMessage, strApplication, strComputerName) " & _
        "SELECT TO_LOB(a.XML_MESSAGE_BODY), 'CHCSII-T', '" & sCompName & "' " & _
        "FROM TMIP_COORDINATOR a"
        
     Call oSQL.Execute(sSQL)
      
      
      'Now delete any rows from the TMIP_Coordinator table
      Call oSQL.Execute("DELETE FROM TMIP_COORDINATOR")
    
              
  '    Set oCmd = New ADODB.Command
  '    Set oParam = New ADODB.Parameter
  '
  '    With oCmd
  '        .CommandText = "SENDTO_TMIP"
  '        .CommandType = adCmdStoredProc
  '        Set oParam = .CreateParameter("vCompName", adVarChar, adParamInput, Len(sCompName) + 10, sCompName)
  '        .Parameters.Append oParam
  '        Set oParam = .CreateParameter("vAppName", adVarChar, adParamInput, Len("CHCSII-T") + 10, "CHCSII-T")
  '        .Parameters.Append oParam
  '    End With
  '
  '    Call oSQL.ExecuteCmd(oCmd)
    
    ' End:SCR #36878
    
      If bUpdateTMIP_Status Then
        If UpdateTMIP_Status(rsClinRecord, strTableName, strPKColumnName) = False Then
          MsgBox "An error occured when attempting to update the TMIP Status Flag. Please retry sending this message " & _
            "by using Data Manager to send the message again manually", vbCritical, "TMIP Status Update Error"
        
          PersistTMIPMessageToDatabase = ERROR_UPDATING_TMIP_STATUS_FLAG
          GoTo Func_Exit
        End If
      End If
    
    Else
      MsgBox "An error occured when attempting to update the TMIP_Coordinator Table. Please retry sending this message " & _
        "by using Data Manager to send the message again manually", vbCritical, "TMIP_Coorinator Table Error"
    
      PersistTMIPMessageToDatabase = ERROR_UPDATING_TMIP_COORDINATOR
      GoTo Func_Exit
    End If
          
    PersistTMIPMessageToDatabase = ERROR_SUCCESS
  
  End If

Func_Exit:
  Set oCmd = Nothing
  Set oParam = Nothing
  Exit Function


ErrHandler:
  MsgBox "An unknown error occured when attempting to add this clinical record to the database. Please retry sending this message " & _
     "by using Data Manager to send the message again manually", vbCritical, "TMIP Message Update Error"
  
  PersistTMIPMessageToDatabase = ERROR_UNHANDLED_SYSTEM_ERROR
  
  GoTo Func_Exit
End Function


Private Function UpdateTMIP_Status(ByVal rst As ADODB.Recordset, ByVal sTableName As String, ByVal strPKColumnName As String) As Boolean

    Dim objDAS As ICHCSII_DAS
    Dim strSQL As String
    
    On Error GoTo ErrHandler
    
    If rst Is Nothing Then
      UpdateTMIP_Status = True
      GoTo Func_Exit
    End If
    
    If rst.EOF Then
      rst.MoveFirst
'    Else
'      UpdateTMIP_Status = True
'      GoTo Func_Exit
    End If
        
    Set objDAS = oDatabase.CHCSII_DAS(Auto) 'New GEMS_DAS
    
    
    Do Until rst.EOF
      strSQL = "UPDATE " & sTableName & " SET TMIP_STATUS = 1 "
      Select Case sTableName
          Case Is = "IMM_PEOPLE_VACCINE"
              'Exceptional case for immunization
              strSQL = strSQL & "WHERE SPON_SSAN_FMP = '" & rst.Fields("SPON_SSAN_FMP").Value & "' AND "
              strSQL = strSQL & "VACCINEID = " & rst.Fields("VACCINEID").Value & " AND "
              strSQL = strSQL & "VACC_DATE = '" & Format(rst.Fields("VACC_DATE").Value, "dd-mmm-yyyy") & "'"
          Case Is = "ENCOUNTERS"
              'exceptional case for encounters
              strSQL = strSQL & "WHERE ENCOUNTERNUMBER = " & rst.Fields("ENCOUNTERNUMBER").Value & " AND "
              strSQL = strSQL & "FACILITYNCID = " & rst.Fields("FACILITYNCID").Value
          Case Is = "READINESS"
              strSQL = strSQL & "WHERE LAST_READINESS_DT = " & oDatabase.SQLDate(rst.Fields("LAST_READINESS_DT").Value) & " AND "
              strSQL = strSQL & "PATIENT_UNIT_NUMBER = " & rst.Fields("PATIENT_UNIT_NUMBER").Value
          Case Else
              strSQL = strSQL & "WHERE " & strPKColumnName & " = " & rst.Fields(strPKColumnName).Value
      End Select
      
      Call objDAS.ExecuteSQL(strSQL) 'Update the timp status of each row
    
      On Error GoTo ErrHandler
      rst.MoveNext
    
    Loop

    UpdateTMIP_Status = True

Func_Exit:
  On Error Resume Next
  Set objDAS = Nothing
  Exit Function

ErrHandler:
  Resume Func_Exit
End Function



Private Function UpdateTMIPStatusToClinicalTable(ByVal lngMsgNumber As Long, ByVal strTableName As String, ByVal strPKColumnName As String, ByVal rsClinRecord As ADODB.Recordset, ByVal strMessage As String)

'******************************************************************************************
    'This function is to update the TMIP_STAUS fields of all the clinical tables involved
    'in TMIP data transfer from LDDB.
    
    'In Parameters:
        'lngMsgNumber       The number of the newly created XML message
        'strTableName       name of the table to be updated
        'strPKColumnName    name of the primary key field
        'rsClinRecord       recordset of the data
        'strMessage         the XML message itself
        
'******************************************************************************************

    Dim strSQL As String
    Dim intI As Integer
    Dim strDate As String
    Dim strTime As String
        
    On Error GoTo ErrHandler
        
    If oDas Is Nothing Then
        Set oDas = oDatabase.CHCSII_DAS(Auto) 'New GEMS_DAS
    End If
   
    If oSQL Is Nothing Then
        Set oSQL = oDatabase.CHCSII_SQLOPS(Auto) 'New GEMS_SQL.SqlOps
    End If
    Call oSQL.Execute("TMIP_COORDINATOR")
    
    rsClinRecord.MoveFirst
    'oDas.BeginTrans
    
    Do Until rsClinRecord.EOF
           
        'Update the clinical table for the status to be "sent"
        
        strSQL = "UPDATE " & strTableName & " SET TMIP_STATUS = " & TmipStatus.Sent & " "
        
        Select Case strTableName
            Case Is = "IMM_PEOPLE_VACCINE"
                'Exceptional case for immunization
                strSQL = strSQL & "WHERE SPON_SSAN_FMP = " & rsClinRecord.Fields("SPON_SSAN_FMP").Value & " AND "
                strSQL = strSQL & "VACCINEID = " & rsClinRecord.Fields("VACCINEID").Value & " AND "
                strSQL = strSQL & "VACC_DATE = '" & Format(rsClinRecord.Fields("VACC_DATE").Value, "dd-mmm-yyyy") & "'"
            Case Is = "ENCOUNTERS"
                'exceptional case for encounters
                strSQL = strSQL & "WHERE ENCOUNTERNUMBER = " & rsClinRecord.Fields("ENCOUNTERNUMBER").Value & " AND "
                strSQL = strSQL & "FACILITYNCID = " & rsClinRecord.Fields("FACILITYNCID").Value
            Case Is = "READINESS"
                strSQL = strSQL & "WHERE LAST_READINESS_DT = " & oDatabase.SQLDate(rsClinRecord.Fields("LAST_READINESS_DT").Value) & " AND "
                strSQL = strSQL & "PATIENT_UNIT_NUMBER = " & rsClinRecord.Fields("PATIENT_UNIT_NUMBER").Value
            Case Else
                strSQL = strSQL & "WHERE " & strPKColumnName & " = " & rsClinRecord.Fields(strPKColumnName).Value
        End Select
        
        intI = oDas.ExecuteSQL(strSQL)
    
    
        'Insert a row in to TMIP_COORDINATOR
        
        'strSQL = "INSERT INTO TMIP_COORDINATOR (DATAID, GEMS_TO_TMIP_MSG_NUMBER, TABLE_NAME, PK_COLUMN_NAME, PK_VALUE, GEMS_TO_TMIP_MSG) VALUES (" ', GEMS_TO_TMIP_MSG
        'strSQL = strSQL & oDas.GetNextID & ", "                             'DATAID
        'strSQL = strSQL & lngMsgNumber & ", "                               'GEMS_TO_TMIP_MSG_NUMBER
        'strSQL = strSQL & "'" & strTableName & "', "                        'TABLE_NAME
        'strSQL = strSQL & "'" & strPKColumnName & "', "                     'PK_COLUMN_NAME
        
        'PK_VALUE
        'Select Case strTableName
        '    Case Is = "IMM_PEOPLE_VACCINE"
                'Exceptional case for immunization
        '        strSQL = strSQL & "'" & rsClinRecord.Fields("SPON_SSAN_FMP").Value & ","
        '        strSQL = strSQL & rsClinRecord.Fields("VACCINEID").Value & ","
        '        strSQL = strSQL & rsClinRecord.Fields("VACC_DATE").Value & "'"
        '    Case Is = "ENCOUNTERS"
                'exceptional case for encounters
        '        strSQL = strSQL & "'" & rsClinRecord.Fields("ENCOUNTERNUMBER").Value & ","
        '        strSQL = strSQL & rsClinRecord.Fields("FACILITYNCID").Value & "'"
        '    Case Is = "READINESS"
                'exceptional case for readiness
        '        strSQL = strSQL & "'" & rsClinRecord.Fields("LAST_READINESS_DT").Value & ","
        '        strSQL = strSQL & rsClinRecord.Fields("PATIENT_UNIT_NUMBER").Value & "'"
        '    Case Else
                'For all other tables
         '       strSQL = strSQL & "'" & rsClinRecord.Fields(strPKColumnName).Value & "'"
        'End Select
        'strSQL = strSQL & ", '" & FixParans(strMessage) & "')"                           'GEMS_TO_TMIP_MSG
        'strSQL = strSQL & ")"

        oSQL.AddNew
            
            oSQL("DATAID") = oDas.GetNextID
            oSQL("GEMS_TO_TMIP_MSG_NUMBER") = lngMsgNumber
            oSQL("TABLE_NAME") = strTableName
            oSQL("PK_COLUMN_NAME") = strPKColumnName
            
            Select Case strTableName
                Case Is = "IMM_PEOPLE_VACCINE"
                    'Exceptional case for immunization
                    strSQL = rsClinRecord.Fields("SPON_SSAN_FMP").Value & ","
                    strSQL = strSQL & rsClinRecord.Fields("VACCINEID").Value & ","
                    strSQL = strSQL & rsClinRecord.Fields("VACC_DATE").Value
                Case Is = "ENCOUNTERS"
                    'exceptional case for encounters
                    strSQL = rsClinRecord.Fields("ENCOUNTERNUMBER").Value & ","
                    strSQL = strSQL & rsClinRecord.Fields("FACILITYNCID").Value
                Case Is = "READINESS"
                    'exceptional case for readiness
                    strSQL = rsClinRecord.Fields("LAST_READINESS_DT").Value & ","
                    strSQL = strSQL & rsClinRecord.Fields("PATIENT_UNIT_NUMBER").Value
                Case Else
                    'For all other tables
                    strSQL = rsClinRecord.Fields(strPKColumnName).Value
            End Select
            

            oSQL("PK_VALUE") = strSQL
            ' Begin:SCR #28321;   Developer: MAG 10/08/2002 03:59 PM
            'Replaced FixParans(strMessage) with strMessage only
            Call oSQL.AppendAsChunk("GEMS_TO_TMIP_MSG", strMessage)
            ' End: SCR #28321;
        
        oSQL.Update
    
        'intI = oDas.ExecuteSQL(strSQL)
        
        rsClinRecord.MoveNext
    Loop

    'oDas.CommitTrans

Exit Function

ErrHandler:
    'oDas.RollbackTrans
    Err.Raise Err.Number, Err.Source, Err.Description
End Function


Private Function BuildTMIPDemographicsMsg(ByVal cData As Collection) As Long
   
    '102123
    'SCR 130299 - LCI - 05/16/2007: Function is used for EIC export too.
    If gbTMIP = False And m_bEICIsDestination = False Then
        BuildTMIPDemographicsMsg = ERROR_SUCCESS
        Exit Function
    End If
   
   Dim Msg As String
   Dim Mmi_IDRS As ADODB.Recordset
   Dim Mmi_GenRS As ADODB.Recordset
   Dim Mmi_NameRS As ADODB.Recordset
   Dim Mmi_InstRS As ADODB.Recordset
   Dim Mmi_SiteRS As ADODB.Recordset
   Dim rsMMI_ADDRESS As ADODB.Recordset
   Dim rsMMI_GENERIC_ID As ADODB.Recordset
   Dim EofFlag As Boolean
   Dim ZPICount As Long
   Dim ZPGCount As Long
   Dim ZPTCount As Long
   
   Dim oHL7 As New HL7XMLBuilder
   Dim oDom As MSXML.DOMDocument
   Dim oRootNode As MSXML.IXMLDOMElement
   Dim oNodeLevel1 As MSXML.IXMLDOMElement
   Dim oRootNode31 As MSXML.IXMLDOMElement
   Dim oNodeLevel131 As MSXML.IXMLDOMElement
   Dim strMsg31 As String
   Dim enmResult As DMErrorCodes
   
   On Error GoTo ErrTrap
   
   ' prove failure as the message is processed
   BuildTMIPDemographicsMsg = ERROR_SUCCESS ' good return
   
   Set Mmi_IDRS = GetRS(cData, "MMI_ID")
   If Not (Mmi_IDRS Is Nothing) Then
      If Mmi_IDRS.EOF = False Then
         Mmi_IDRS.MoveFirst
      End If
   End If
   Set Mmi_NameRS = GetRS(cData, "MMI_NAME")
   If Not (Mmi_NameRS Is Nothing) Then
      If Mmi_NameRS.EOF = False Then
         Mmi_NameRS.MoveFirst
      End If
   End If
   Set Mmi_SiteRS = GetRS(cData, "MMI_SITE")
   If Not (Mmi_SiteRS Is Nothing) Then
      If Mmi_SiteRS.EOF = False Then
         Mmi_SiteRS.MoveFirst
      End If
   End If
   Set rsMMI_GENERIC_ID = GetRS(cData, "MMI_GENERIC_ID")
   If Not (rsMMI_GENERIC_ID Is Nothing) Then
      If rsMMI_GENERIC_ID.EOF = False Then
         rsMMI_GENERIC_ID.MoveFirst
      End If
   End If
   Set rsMMI_ADDRESS = GetRS(cData, "MMI_ADDRESS")
   If Not (rsMMI_ADDRESS Is Nothing) Then
      If rsMMI_ADDRESS.EOF = False Then
         rsMMI_ADDRESS.MoveFirst
      End If
   End If
   
   Set oDom = New MSXML.DOMDocument
   
   If oDas Is Nothing Then
      Set oDas = oDatabase.CHCSII_DAS(Auto) 'New GEMS_DAS
   End If
   
   PIDCount = 0
   
   Dim strMessageGUID As String
   '<< Begin:SCR #28652;   Developer: Brian Mowbray 10/21/2002 12:59 PM
   Dim strMessageGUID31 As String
   '>> End: SCR #28652;

   
   Do
      ' build message from parts
      
'      MsgType = "ADT_A28"
'      strMsgType31 = "ADT_A31" '<SCR 23922 CC>
      
      Set oRootNode = oDom.createElement("ADT_A28")
      Set oRootNode31 = oDom.createElement("ADT_A01") '<SCR 23922 CC>
      
      strMessageGUID = GetNewGUID
      '<< Begin:SCR #28652;   Developer: Brian Mowbray 10/21/2002 01:00 PM
      strMessageGUID31 = GetNewGUID
      '>> End: SCR #28652;
      
      Set oNodeLevel1 = oHL7.ReturnNewMSHNode("A28", "ADT", strMessageGUID) ' CStr(g_lngTMIPMsgNumber))
      '<< Begin:SCR #28652;   Developer: Brian Mowbray 10/21/2002 01:01 PM
      Set oNodeLevel131 = oHL7.ReturnNewMSHNode("A31", "ADT", strMessageGUID31)
      '>> End: SCR #28652;
      oRootNode.appendChild oNodeLevel1
      oRootNode31.appendChild oNodeLevel131 '<SCR 23922 CC>
      
      Set oNodeLevel1 = oHL7.ReturnNewEVNNode("A28")
      Set oNodeLevel131 = oNodeLevel1.cloneNode(True) '<SCR 23922 CC>
      oRootNode.appendChild oNodeLevel1
      oRootNode31.appendChild oNodeLevel131 '<SCR 23922 CC>
      
      If Not (gobjPatient Is Nothing) Then
         PIDCount = PIDCount + 1
         
         Set oNodeLevel1 = oHL7.ReturnNewPIDNode(PIDCount)
         Set oNodeLevel131 = oNodeLevel1.cloneNode(True)
         oRootNode.appendChild oNodeLevel1
         oRootNode31.appendChild oNodeLevel131 '<SCR 23922 CC>
         
         Set oNodeLevel1 = oHL7.ReturnTMIPDefinedZD1Node()
         Set oNodeLevel131 = oNodeLevel1.cloneNode(True)
         oRootNode.appendChild oNodeLevel1
         oRootNode31.appendChild oNodeLevel131 '<SCR 23922 CC>
      End If
      
      Set oNodeLevel1 = oHL7.ReturnFakePV1Segment
      Set oNodeLevel131 = oNodeLevel1.cloneNode(True)
      oRootNode.appendChild oNodeLevel1
      oRootNode31.appendChild oNodeLevel131 '<SCR 23922 CC>
      
      If Not (Mmi_IDRS Is Nothing) Then
         ZPICount = ZPICount + 1
         Set oNodeLevel1 = oHL7.ReturnNewZPINode(Mmi_IDRS, Mmi_NameRS, Mmi_SiteRS, rsMMI_GENERIC_ID, rsMMI_ADDRESS)
         Set oNodeLevel131 = oNodeLevel1.cloneNode(True)
         oRootNode.appendChild oNodeLevel1
         oRootNode31.appendChild oNodeLevel131 '<SCR 23922 CC>
         
         Set Mmi_GenRS = GetRS(cData, "MMI_GENERICDATA")
         If Not (Mmi_GenRS Is Nothing) Then
            If Mmi_GenRS.EOF = False Then
               Mmi_GenRS.MoveFirst
               While Mmi_GenRS.EOF = False
                  ZPGCount = ZPGCount + 1
                  Set oNodeLevel1 = oHL7.ReturnNewZPGNode(ZPGCount, Mmi_GenRS)
                  Set oNodeLevel131 = oNodeLevel1.cloneNode(True)
                  oRootNode.appendChild oNodeLevel1
                  oRootNode31.appendChild oNodeLevel131 '<SCR 23922 CC>
                  'Temp = Temp & BuildNewZPGSegment(ZPGCount, Mmi_GenRS)
                  Mmi_GenRS.MoveNext
               Wend
            End If
         End If
         
         Set Mmi_InstRS = GetRS(cData, "MMI_INST_NAME")
         If Not (Mmi_InstRS Is Nothing) Then
            If Mmi_InstRS.EOF = False Then
               Mmi_InstRS.MoveFirst
               While Mmi_InstRS.EOF = False
                  ZPTCount = ZPTCount + 1
                  Set oNodeLevel1 = oHL7.ReturnNewZPTNode(ZPTCount, Mmi_InstRS)
                  Set oNodeLevel131 = oNodeLevel1.cloneNode(True)
                  oRootNode.appendChild oNodeLevel1
                  oRootNode31.appendChild oNodeLevel131 '<SCR 23922 CC>
                  Mmi_InstRS.MoveNext
               Wend
            End If
         End If
      End If
      
      Msg = "<?xml version=""1.0""?>"
      Msg = Msg & oRootNode.xml
      
      If m_bEICIsDestination = True Then
        
        m_strEICTempHL7 = m_strEICTempHL7 & Msg
        
      Else 'This is a TMIP Transaction
            enmResult = PersistTMIPMessageToDatabase(strMessageGUID, "DEMOGRAPHICS", "", Mmi_GenRS, Msg, "", False, "TMIP_A28")
            
            If enmResult = ERROR_SUCCESS Then
                BuildTMIPDemographicsMsg = ERROR_SUCCESS ' good return
            Else
                BuildTMIPDemographicsMsg = enmResult
                Exit Function
            End If
        
            strMsg31 = "<?xml version=""1.0""?>"
            strMsg31 = strMsg31 & oRootNode31.xml
            
            enmResult = PersistTMIPMessageToDatabase(strMessageGUID31, "DEMOGRAPHICS", "", Mmi_GenRS, strMsg31, "", False, "TMIP_A31")
            
            If enmResult = ERROR_SUCCESS Then
                BuildTMIPDemographicsMsg = ERROR_SUCCESS ' good return
            Else
                BuildTMIPDemographicsMsg = enmResult
                Exit Function
            End If
      End If
      
      If Mmi_IDRS Is Nothing Then
         EofFlag = True
      Else
         Mmi_IDRS.MoveNext
         EofFlag = Mmi_IDRS.EOF
      End If
      
   Loop While EofFlag = False
   
   If Mmi_IDRS Is Nothing Then
      BuildTMIPDemographicsMsg = DMErrorCodes.ERROR_EMPTY_RECORDSET '<SCR22548>
      Exit Function
   End If
   If Mmi_IDRS.BOF = True And Mmi_IDRS.EOF = True Then
      BuildTMIPDemographicsMsg = DMErrorCodes.ERROR_EMPTY_RECORDSET '<SCR22548>
      Exit Function
   End If
   
   Exit Function

ErrTrap:
   Err.Clear
   BuildTMIPDemographicsMsg = ERROR_UNHANDLED_SYSTEM_ERROR
   Exit Function
   Resume
End Function

Private Property Let I_DMBE_WhoseXMLMessage(ByVal rhs As IDMBE.XMLMessageSource)

    m_enmXMLMessageSource = rhs
    
End Property

Private Property Get I_DMBE_WhoseXMLMessage() As IDMBE.XMLMessageSource
    
    I_DMBE_WhoseXMLMessage = m_enmXMLMessageSource

End Property

Private Function I_DMBE_WriteBMISTFile(ByVal enmDMFormat As IDMBE.DMFormats, ByVal strFilePath As String) As Long
    
    Dim oXML As MC4_Ops
    Dim sFileName As String
    Dim sTemp As String
    Dim lRetVal As DMErrorCodes
    Dim oFS As FileSystemObject
    Dim nCnt As Integer
    
    On Error GoTo Func_Err
    
    Set oXML = New MC4_Ops
    Set oFS = New FileSystemObject
     
    If oDas Is Nothing Then
        Set oDas = oDatabase.CHCSII_DAS(Auto) 'New GEMS_DAS
    End If
    
    lRetVal = ERROR_SUCCESS
         
    If Not (gobjPatient Is Nothing) Then
        If Not BuildPatientRst(True, True) Then
          lRetVal = ERROR_WRITING_BMIST_FILE
          GoTo Func_Exit
        End If
    Else
      lRetVal = ERROR_WRITING_BMIST_FILE
      GoTo Func_Exit
    End If
     
    'Build the raw xml file based on the recordsets in the collection
    oXML.BuildBMISTXML g_rstPatient
    
     ' Begin:SCR #44201;   Developer: Mag 12/01/2003 04:20 PM
     'Test to see if the temp path has a value
     If Len(strFilePath) = 0 Then
        MsgBox "Please select a destination path!", vbCritical, "NO PATH SELECTED"
        I_DMBE_WriteBMISTFile = ERROR_WRITING_BMIST_FILE
        Exit Function
     End If
     
     If Right(strFilePath, 1) <> "\" Then
       strFilePath = strFilePath & "\"
     End If
     
     nCnt = 0
     Do
      'Create a unique files name using the data/time and facilityncid
      sTemp = gobjLogon.FacilityNCID & "_" & Format(Now, "yyyymmddhhmmss")
      
      sFileName = "CHCSIITBMIST_XE_" & sTemp & nCnt & ".xml"
    Loop Until (Not oFS.FileExists(strFilePath & sFileName))
    ' End: SCR #44201;
    
    'Persist xml to path
    If Not oXML.WriteXML(strFilePath, sFileName) Then
       lRetVal = ERROR_WRITING_BMIST_FILE
    End If

    I_DMBE_WriteBMISTFile = lRetVal
  
Func_Exit:
  On Error Resume Next
  Set oXML = Nothing
  Set oFS = Nothing
  Exit Function

Func_Err:
   I_DMBE_WriteBMISTFile = ERROR_WRITING_BMIST_FILE
   Resume Func_Exit
   Resume 'Debug
End Function

Private Function I_DMBE_WriteToEIC(ByVal strSponsorSSN As String) As Long
' SCR #44079;    Developer: Mag 10/20/2003 01:08 PM
' Changed SSN to SponsorSSN in Header information of EIC.
    Dim oEIC As EIC_SERVERLib.EIC
    Dim lngError As Long
    Dim strError As String
    Dim strTemp As String
    
    On Error GoTo ErrHandler
    
    Set oEIC = New EIC_SERVERLib.EIC
    
    'This next LOC will write the file "EIC.dat" to the folder
    '  that's setup in the regristy "TMIP\EIC"
    oEIC.EICWrite strSponsorSSN, m_strEICTempHL7, lngError, strError
    
    If lngError <> 0 Then
        
        strTemp = TranslateEICErrorCode(lngError)
        If oDas Is Nothing Then Set oDas = oDatabase.CHCSII_DAS(Auto) 'GEMS_DAS
        If oSQL Is Nothing Then Set oSQL = oDatabase.CHCSII_SQLOPS(Auto) 'New GEMS_SQL.SqlOps
        
        oSQL.Execute "TMIP_INVALID_MESSAGE"
        
        oSQL.AddNew
            oSQL("DATAID") = oDas.GetNextID
            oSQL("DESCRIPTION") = strTemp & " " & strError
            oSQL("TIMESTAMP") = Now
            ' Begin:SCR #28321;   Developer: MAG 10/08/2002 03:59 PM
            'Replaced FixParans(strMessage) with strMessage only
            Call oSQL.AppendAsChunk("MESSAGE", m_strEICTempHL7)
            ' End: SCR #28321;
        oSQL.Update
        
        'If gobjShared Is Nothing Then
        '    Set gobjShared = New CWShared
        'End If
        'gobjShared.ShowVBError lngError, TranslateEICErrorCode(lngError), "WriteToEIC Function", "DMBE", vbExclamation
        
        I_DMBE_WriteToEIC = TranslateEICErrorCode(lngError)
    Else
       ' Begin:SCR #40034;   Developer: Mag 08/22/2003 02:23 PM
       'Call WriteEICTestingHL7(m_strEICTempHL7)
       ' End: SCR #40034;
        
        ' Begin:SCR #27495;   Developer: MAG 02/19/2004
      If Not (m_EICEncounterNumber = "" Or m_EICEncounterNumber = "0") Then
        If oSQL Is Nothing Then Set oSQL = oDatabase.CHCSII_SQLOPS(Auto) 'New GEMS_SQL.SqlOps
        
          oSQL.Execute "Select UpdateLock From Encounters where EncounterNumber = " & m_EICEncounterNumber
          
     'We need to set the UpadteLock to "0" or False.
    '-1 is preventing the Encounter from being modified.
    'Because when we send out an encounter, we need to be able to amend the encounter.
         If oSQL.RecordCount > 0 Then
            oSQL("UPDATELOCK") = False
            oSQL.Update
          End If
          ' End: SCR #27495;
        End If
        
        
        I_DMBE_WriteToEIC = DMErrorCodes.ERROR_SUCCESS
      End If
    
Exit Function

ErrHandler:

    If oDas Is Nothing Then Set oDas = oDatabase.CHCSII_DAS(Auto) 'GEMS_DAS
    If oSQL Is Nothing Then Set oSQL = oDatabase.CHCSII_SQLOPS(Auto) 'New GEMS_SQL.SqlOps
        
    Call oSQL.Execute("TMIP_INVALID_MESSAGE")
    
    oSQL.AddNew
        oSQL("DATAID") = oDas.GetNextID
        oSQL("DESCRIPTION") = Err.Number & ": " & Err.Description & " " & Err.Source
        oSQL("TIMESTAMP") = Now
        
        ' Begin:SCR #28321;   Developer: MAG 10/08/2002 03:59 PM
        'Replaced FixParans(strMessage) with strMessage only
        Call oSQL.AppendAsChunk("MESSAGE", m_strEICTempHL7)
        ' End: SCR #28321;
    oSQL.Update
'
'   If gobjShared Is Nothing Then
'      Set gobjShared = New CWShared
'   End If
'   gobjShared.ShowVBError Err.Number, Err.Description, "WriteToEIC Function", "DMBE", vbExclamation
   
   I_DMBE_WriteToEIC = DMErrorCodes.ERROR_WRITE_TO_EIC_ERROR

End Function

Private Sub WriteEICTestingHL7(ByVal strHL7 As String)

    Dim strPath As String
    Dim strFileName As String
    
    On Error GoTo ErrHandler
    
    strPath = Environ("TEMP") & "\TMIP\MSGS\"
    strFileName = Format(Now, "mmm d yyyy hh:mm:ss AMPM")
    strFileName = Replace(strFileName, ":", "_", 1, -1, vbTextCompare)
    strFileName = Replace(strFileName, " ", "_", 1, -1, vbTextCompare)
    strPath = strPath & "EIC_" & strFileName & ".xml"
    strPath = EnsureUniqueFile(strPath)
    Call WriteFileContents(strPath, strHL7)

Exit Sub

ErrHandler:
    If gobjShared Is Nothing Then
      Set gobjShared = New CWShared
   End If
   gobjShared.ShowVBError Err.Number, Err.Description, "WriteToEIC Function", "DMBE", vbExclamation
   Err.Clear
End Sub

Private Function BuildTMIPImmunizationsMsg(ByVal colRSObjects As Collection) As Long
   
    '102123
    If gbTMIP = False Then
        BuildTMIPImmunizationsMsg = ERROR_SUCCESS
        Exit Function
    End If
   
    Dim strMessage As String
    Dim rsIMMUNIZATION As ADODB.Recordset
    Dim strMessageType As String
    
    Dim oHL7 As New HL7XMLBuilder
    Dim oDom As MSXML.DOMDocument
    Dim oRootNode As MSXML.IXMLDOMElement
    Dim oNodeLevel1 As MSXML.IXMLDOMElement
    Dim oNodeLevel2 As MSXML.IXMLDOMElement
    Dim oNodeLevel3 As MSXML.IXMLDOMElement
    Dim strMessageGUID As String
    Dim lResult As Long
    
    On Error GoTo ErrTrap
    
    BuildTMIPImmunizationsMsg = ERROR_SUCCESS
   
    '<< Begin:SCR #38306;   Developer: Brian Mowbray 06/30/2003 03:37 PM
    'Set rsIMMUNIZATION = GetRS(colRSObjects, "IMMUNIZATIONS")
    Set rsIMMUNIZATION = GetRS(colRSObjects, "IMMUNIZATION")

    
    lResult = CheckRst(rsIMMUNIZATION, "")
    
    If lResult <> 0 Then
      BuildTMIPImmunizationsMsg = lResult
      Exit Function
    End If
    
    Set oDom = New MSXML.DOMDocument
    
    strMessageType = "VXU_V04"
    Set oRootNode = oDom.createElement("VXU_V04")
    
    If oDas Is Nothing Then
        Set oDas = oDatabase.CHCSII_DAS(Auto) 'New GEMS_DAS
    End If
    
    strMessageGUID = GetNewGUID
    
    Set oNodeLevel1 = oHL7.ReturnNewMSHNode("V04", "VXU", strMessageGUID)  ',  CStr(g_lngTMIPMsgNumber))
    oRootNode.appendChild oNodeLevel1
    
    If Not (gobjPatient Is Nothing) Then
        PIDCount = PIDCount + 1
        
        Set oNodeLevel1 = oHL7.ReturnNewPIDNode(PIDCount)
        oRootNode.appendChild oNodeLevel1
        Set oNodeLevel1 = oHL7.ReturnTMIPDefinedZD1Node()
        oRootNode.appendChild oNodeLevel1
    End If
    
    '<SCR 24313 CC>
    Set oNodeLevel1 = oDom.createElement("VXU_V04.LST.5")
    
    Do Until rsIMMUNIZATION.EOF
         ImmCount = ImmCount + 1
         
         '<< Begin:SCR #27454;   Developer: Brian Mowbray 09/05/2002 05:45 PM
         Set oNodeLevel2 = oDom.createElement("VXU_V04.GRP.4")
         Set oNodeLevel3 = oHL7.ReturnNewRXANode(ImmCount, rsIMMUNIZATION)
         oNodeLevel2.appendChild oNodeLevel3
         
         Set oNodeLevel3 = oHL7.ReturnNewZRANode(rsIMMUNIZATION)
         oNodeLevel2.appendChild oNodeLevel3
         
         oNodeLevel1.appendChild oNodeLevel2
         '>> End: SCR #27454;
      
      rsIMMUNIZATION.MoveNext
    Loop
    
    oRootNode.appendChild oNodeLevel1
    
    strMessage = "<?xml version=""1.0""?>"
    strMessage = strMessage & oRootNode.xml
                    
    Dim enmResult As DMErrorCodes
    
    enmResult = PersistTMIPMessageToDatabase(strMessageGUID, "IMM_PEOPLE_VACCINE", "SPON_SSAN_FMP", rsIMMUNIZATION.Clone, strMessage, strMessageType, , "TMIP_V04")
      
    If enmResult = ERROR_SUCCESS Then
       BuildTMIPImmunizationsMsg = ERROR_SUCCESS ' good return
    Else
       BuildTMIPImmunizationsMsg = enmResult
    End If
Exit Function
ErrTrap:
   Err.Clear
   BuildTMIPImmunizationsMsg = ERROR_UNHANDLED_SYSTEM_ERROR
End Function

Private Function BuildTMIPMedicationsMsg(ByVal cData As Collection) As Long
   
    '102123
    If gbTMIP = False Then
        BuildTMIPMedicationsMsg = ERROR_SUCCESS
        Exit Function
    End If
   
  Dim MedicationsRS As ADODB.Recordset
  Dim RXListRS As ADODB.Recordset
  'Dim PIDCount As Long
  Dim ZRECount As Long
  Dim Msg As String
  Dim MsgType As String
  
  Dim oHL7 As New HL7XMLBuilder
  Dim oDom As MSXML.DOMDocument
  Dim oRootNode As MSXML.IXMLDOMElement
  Dim oNodeLevel1 As MSXML.IXMLDOMElement
  Dim oNodeLevel2 As MSXML.IXMLDOMElement
  Dim oNodeLevel3 As MSXML.IXMLDOMElement
  
  Dim strMessageGUID As String
  Dim lResult As Long
  
  On Error GoTo ErrTrap
  
  BuildTMIPMedicationsMsg = ERROR_SUCCESS
  
  Set MedicationsRS = GetRS(cData, "MEDICATIONS")
  
  lResult = CheckRst(MedicationsRS, "")
  
  If lResult <> 0 Then
    BuildTMIPMedicationsMsg = lResult ' bad return
    Exit Function
  End If
  
  Set oDom = New MSXML.DOMDocument
      
  Set oRootNode = oDom.createElement("RDE_O01")
      
  If oDas Is Nothing Then
    Set oDas = oDatabase.CHCSII_DAS(Auto) 'GEMS_DAS
  End If
      
  strMessageGUID = GetNewGUID
      
  ' build message from parts
  MsgType = "RDE_O01"
      
  Set oNodeLevel1 = oHL7.ReturnNewMSHNode("O01", "RDE", strMessageGUID)   ', CLng(g_lngTMIPMsgNumber))
  oRootNode.appendChild oNodeLevel1
  
  If Not (gobjPatient Is Nothing) Then
     PIDCount = PIDCount + 1
     Set oNodeLevel1 = oHL7.ReturnNewPIDNode(PIDCount)
     oRootNode.appendChild oNodeLevel1
     Set oNodeLevel1 = oHL7.ReturnTMIPDefinedZD1Node()
     oRootNode.appendChild oNodeLevel1
  End If
      
  '<< Begin:SCR #27568;   Developer: Brian Mowbray 09/11/2002 15:16
  Set oNodeLevel1 = oDom.createElement("RDE_O01.LST.14")
      
  Do Until MedicationsRS.EOF
    Set RXListRS = GetRS(cData, "RX_LIST")
    Call FilterRS("DRUG_IEN", MedicationsRS("MEDICATIONNCID"), RXListRS)
    
    Set oNodeLevel2 = oDom.createElement("RDE_O01.GRP.5")
    Set oNodeLevel3 = oHL7.ReturnNewORCNode(MedicationsRS)
    oNodeLevel2.appendChild oNodeLevel3
    Set oNodeLevel3 = oHL7.ReturnNewRXENode(MedicationsRS)
    oNodeLevel2.appendChild oNodeLevel3

    ZRECount = ZRECount + 1
    Set oNodeLevel3 = oHL7.ReturnNewZRENode(ZRECount, MedicationsRS, RXListRS)
    oNodeLevel2.appendChild oNodeLevel3
    oNodeLevel1.appendChild oNodeLevel2

    MedicationsRS.MoveNext
  
  Loop
      
  oRootNode.appendChild oNodeLevel1
  '<< End:SCR #27568
    
    
   Msg = "<?xml version=""1.0""?>"
   Msg = Msg & oRootNode.xml
   
   Dim enmResult As DMErrorCodes
   
   enmResult = PersistTMIPMessageToDatabase(strMessageGUID, "MEDICATIONS", "DATAID", MedicationsRS.Clone, Msg, MsgType, "O01")
     
   If enmResult = ERROR_SUCCESS Then
      BuildTMIPMedicationsMsg = ERROR_SUCCESS ' good return
   Else
      BuildTMIPMedicationsMsg = enmResult
   End If
  
Exit Function

ErrTrap:
   Err.Clear
   BuildTMIPMedicationsMsg = ERROR_UNHANDLED_SYSTEM_ERROR
End Function
      
Private Function BuildTMIPProblemsMsg(ByVal cData As Collection) As Long
   
    '102123
    If gbTMIP = False Then
        BuildTMIPProblemsMsg = ERROR_SUCCESS
        Exit Function
    End If
   
  Dim ProblemsRS As ADODB.Recordset
  Dim ProblemsRelationshipsRS As ADODB.Recordset
  'Dim PIDCount As Long
  Dim PV1Count As Long
  Dim ZPBCount As Long
  Dim Msg As String
  Dim MsgType As String
  Dim strMessageGUID As String
  Dim sDataIds As String
  Dim oHL7 As New HL7XMLBuilder
  Dim oDom As MSXML.DOMDocument
  Dim oRootNode As MSXML.IXMLDOMElement
  Dim oNodeLevel1 As MSXML.IXMLDOMElement
  Dim lResult As Long
  Dim sSQL As String
  Dim sUnitNumber As String
  Dim sDataID As String
  On Error GoTo ErrTrap
   
  BuildTMIPProblemsMsg = ERROR_SUCCESS
   
  Set ProblemsRS = GetRS(cData, "PROBLEMS")
  
  lResult = CheckRst(ProblemsRS, "")
  
  If lResult <> 0 Then
    BuildTMIPProblemsMsg = lResult
    Exit Function
  End If
   
  sUnitNumber = ProblemsRS("UNIT_NUMBER") & vbNullString
  
  Do While Not ProblemsRS.EOF
    sDataID = sDataID & ProblemsRS("DATAID") & ", "
    ProblemsRS.MoveNext
  Loop
  
  If Len(sDataID) > 0 Then
    sDataID = Left(sDataID, Len(sDataID) - 2)
  
    If oDas Is Nothing Then
      Set oDas = oDatabase.CHCSII_DAS(Auto)
    End If
      
    Set ProblemsRS = Nothing
    
    sSQL = "Select * From PROBLEMS"
    sSQL = sSQL & " Where Unit_Number = " & sUnitNumber
    sSQL = sSQL & " And DataID IN (" & sDataID & ")"
    sSQL = sSQL & " And TMIP_STATUS <> 1"
    
    Set ProblemsRS = oDas.OpenRecordset(sSQL)
    
    lResult = CheckRst(ProblemsRS, "")
    
    If lResult <> 0 Then
      BuildTMIPProblemsMsg = lResult
      Exit Function
    End If
  End If
  
  Set oDom = New MSXML.DOMDocument
  Set oRootNode = oDom.createElement("ADT_A01")
  
  MsgType = "ADT_A01"
  
  strMessageGUID = GetNewGUID
  
  Set oNodeLevel1 = oHL7.ReturnNewMSHNode("A04", "ADT", strMessageGUID) ', CLng(g_lngTMIPMsgNumber))
  oRootNode.appendChild oNodeLevel1
  Set oNodeLevel1 = oHL7.ReturnNewEVNNode("A04")
  oRootNode.appendChild oNodeLevel1
  
  If Not (gobjPatient Is Nothing) Then
    PIDCount = PIDCount + 1
    Set oNodeLevel1 = oHL7.ReturnNewPIDNode(PIDCount)
    oRootNode.appendChild oNodeLevel1
    Set oNodeLevel1 = oHL7.ReturnTMIPDefinedZD1Node()
    oRootNode.appendChild oNodeLevel1
  End If
    
  'No Problem record in the 84 box has an EncounterNumber
  PV1Count = PV1Count + 1
  
  Set oNodeLevel1 = oHL7.ReturnNewALGPV1Node(PV1Count)
  oRootNode.appendChild oNodeLevel1
 
  Do Until ProblemsRS.EOF
   
      ZPBCount = ZPBCount + 1
      Set oNodeLevel1 = oHL7.ReturnNewZPBNode(ZPBCount, ProblemsRS)
      oRootNode.appendChild oNodeLevel1
      
      ProblemsRS.MoveNext
    
  Loop
  
  Msg = "<?xml version=""1.0""?>"
  Msg = Msg & oRootNode.xml

  Dim enmResult As DMErrorCodes
    
  enmResult = PersistTMIPMessageToDatabase(strMessageGUID, "PROBLEMS", "DATAID", ProblemsRS.Clone, Msg, MsgType, , "TMIP_A04")
      
  If enmResult = ERROR_SUCCESS Then
     BuildTMIPProblemsMsg = DMErrorCodes.ERROR_SUCCESS '<SCR 22549 CC>
  Else
     BuildTMIPProblemsMsg = enmResult
  End If

Exit Function

ErrTrap:
   Err.Clear
   BuildTMIPProblemsMsg = ERROR_UNHANDLED_SYSTEM_ERROR
End Function

Private Function BuildTMIPReadinessMsg(ByVal cData As Collection) As Long
   
    '102123
    If gbTMIP = False Then
        BuildTMIPReadinessMsg = ERROR_SUCCESS
        Exit Function
    End If
   
  Dim Msg As String
  Dim ReadinessRS As ADODB.Recordset
  Dim RDCount As Long
  Dim MsgType As String
  
  Dim oHL7 As New HL7XMLBuilder
  Dim oDom As MSXML.DOMDocument
  Dim oRootNode As MSXML.IXMLDOMElement
  Dim oNodeLevel1 As MSXML.IXMLDOMElement
  Dim strMSH10GUID As String
  Dim lResult As Long
  
  On Error GoTo ErrTrap
  
  BuildTMIPReadinessMsg = ERROR_SUCCESS ' good return
   
  Set ReadinessRS = GetRS(cData, "READINESS")
     
  lResult = CheckRst(ReadinessRS, "")
  
  If lResult <> 0 Then
    BuildTMIPReadinessMsg = lResult
    Exit Function
  End If

  ' build message from parts
  If oDas Is Nothing Then
    Set oDas = oDatabase.CHCSII_DAS(Auto) 'GEMS_DAS
  End If
  
  Set oDom = New MSXML.DOMDocument
  Set oRootNode = oDom.createElement("ADT_A01")
      
  strMSH10GUID = GetNewGUID
      
  MsgType = "ADT_A01"
  
  Set oNodeLevel1 = oHL7.ReturnNewMSHNode("A04", "ADT", strMSH10GUID)
  oRootNode.appendChild oNodeLevel1
  Set oNodeLevel1 = oHL7.ReturnNewEVNNode("A04")
  oRootNode.appendChild oNodeLevel1
      
  If Not (gobjPatient Is Nothing) Then
     PIDCount = PIDCount + 1
     'Temp = Temp & BuildNewPIDSegment(PIDCount, gobjPatient)
     'Temp = Temp & BuildTMIPDefinedZD1Segment(gobjPatient)
     Set oNodeLevel1 = oHL7.ReturnNewPIDNode(PIDCount)
     oRootNode.appendChild oNodeLevel1
     Set oNodeLevel1 = oHL7.ReturnTMIPDefinedZD1Node()
     oRootNode.appendChild oNodeLevel1
  End If
      
  ' Begin:SCR #43913; Developer: Mag 10/15/2003 12:11 PM
  Set oNodeLevel1 = oHL7.ReturnFakePV1Segment
  oRootNode.appendChild oNodeLevel1
  ' End: SCR #43913;
  
  Do Until ReadinessRS.EOF
     
'>> Begin:SCR #40466;   Developer: Brian Mowbray 08/25/2003 10:17 AM
'     PV1Count = PV1Count + 1
'     vTemp = ReadinessRS.Fields("PRIMARYPROVIDERNCID").Value
'     If IsNull(vTemp) Then
'        lPrime = 0
'     Else
'        lPrime = CLng(vTemp)
'     End If
'     vTemp = ReadinessRS.Fields("CLINICNCID").Value
'     If IsNull(vTemp) Then
'        lClinic = 0
'     Else
'        lClinic = CLng(vTemp)
'     End If
'     vTemp = ReadinessRS.Fields("Type").Value
'     If IsNull(vTemp) Then
'        lType = 0
'     Else
'        lType = CLng(vTemp)
'     End If
'     vTemp = ReadinessRS.Fields("STARTDTS").Value
'     If IsNull(vTemp) Then
'        AdmitDate = ""
'     Else
'        AdmitDate = vTemp
'     End If
'     vTemp = ReadinessRS.Fields("ENDDTS").Value
'     If IsNull(vTemp) Then
'        DischargeDate = ""
'     Else
'        DischargeDate = vTemp
'     End If
'     vTemp = ReadinessRS.Fields("DISPOSITIONNCID").Value
'     If IsNull(vTemp) Then
'        Status = ""
'     Else
'        Status = vTemp
'     End If
'     Set oNodeLevel1 = oHL7.ReturnNewPV1Node(PV1Count, lPrime, lClinic, lType, AdmitDate, DischargeDate, Status)
'     oRootNode.appendChild oNodeLevel1
     '<< End: SCR #40466;
     RDCount = RDCount + 1
     Set oNodeLevel1 = oHL7.ReturnNewZRDNode(RDCount, ReadinessRS)
     oRootNode.appendChild oNodeLevel1
  
    ReadinessRS.MoveNext
    
  Loop
      
  Msg = "<?xml version=""1.0""?>"
  Msg = Msg & oRootNode.xml
   
  Dim enmResult As DMErrorCodes
   
  enmResult = PersistTMIPMessageToDatabase(strMSH10GUID, "READINESS", "LAST_READINESS_DT & PATIENT_UNIT_NUMBER", ReadinessRS.Clone, Msg, MsgType, , "TMIP_A04")
     
  If enmResult = ERROR_SUCCESS Then
    BuildTMIPReadinessMsg = ERROR_SUCCESS ' good return
  Else
    BuildTMIPReadinessMsg = enmResult
  End If
   
   Exit Function

ErrTrap:
   Err.Clear
   BuildTMIPReadinessMsg = ERROR_UNHANDLED_SYSTEM_ERROR
End Function
      
Private Function BuildTMIPVitalsMsg(ByVal cData As Collection) As Long
  
    '102123
    If gbTMIP = False Then
        BuildTMIPVitalsMsg = ERROR_SUCCESS
        Exit Function
    End If
  
  Dim Msg As String
  Dim VitalsRS As ADODB.Recordset
  Dim VitalSignsModRS As ADODB.Recordset
  Dim VitalSignsValRS As ADODB.Recordset
  Dim PV1Count As Long
  Dim VTCount As Long
  Dim MsgType As String
  Dim strMessageGUID As String
  Dim oHL7 As New HL7XMLBuilder
  Dim oDom As MSXML.DOMDocument
  Dim oRootNode As MSXML.IXMLDOMElement
  Dim oNodeLevel1 As MSXML.IXMLDOMElement
  Dim lResult As Long
  
  On Error GoTo ErrTrap
   
  BuildTMIPVitalsMsg = ERROR_SUCCESS
    
  Set VitalsRS = GetRS(cData, "VITALS")
  Set VitalSignsValRS = GetRS(cData, "VITALSIGNSVALUE")
  Set VitalSignsModRS = GetRS(cData, "VITALSIGNSMODIFIER")
    
  lResult = CheckRst(VitalsRS, "")
  If lResult <> 0 Then
    BuildTMIPVitalsMsg = lResult
    Exit Function
  End If
  
  lResult = CheckRst(VitalSignsValRS, "")
  If lResult <> 0 Then
    BuildTMIPVitalsMsg = lResult
    Exit Function
  End If
    
  ' build message from parts
  Set oDom = New MSXML.DOMDocument
      
  MsgType = "ADT_A01"
  Set oRootNode = oDom.createElement("ADT_A01")
      
  strMessageGUID = GetNewGUID
      
  Set oNodeLevel1 = oHL7.ReturnNewMSHNode("A04", "ADT", strMessageGUID) ', CStr(g_lngTMIPMsgNumber))
  oRootNode.appendChild oNodeLevel1
  Set oNodeLevel1 = oHL7.ReturnNewEVNNode("A04")
  oRootNode.appendChild oNodeLevel1
      
  If Not (gobjPatient Is Nothing) Then
     PIDCount = PIDCount + 1
     Set oNodeLevel1 = oHL7.ReturnNewPIDNode(PIDCount)
     oRootNode.appendChild oNodeLevel1
     Set oNodeLevel1 = oHL7.ReturnTMIPDefinedZD1Node()
     oRootNode.appendChild oNodeLevel1
  End If
  
  ' build minimum patient visit data segment
  PV1Count = PV1Count + 1
  Set oNodeLevel1 = oHL7.ReturnNewALGPV1Node(PV1Count)
  oRootNode.appendChild oNodeLevel1
      
  Do Until VitalsRS.EOF
    
    VitalSignsValRS.Filter = ""
    VitalSignsModRS.Filter = ""
    
    If VitalSignsValRS.RecordCount > 0 Then
        VitalSignsValRS.MoveFirst
        VitalSignsValRS.Filter = "DATAID = " & VitalsRS.Fields("DATAID").Value
    End If
    
    If VitalSignsModRS.RecordCount > 0 Then
        VitalSignsModRS.MoveFirst
        VitalSignsModRS.Filter = "DATAID = " & VitalsRS.Fields("DATAID").Value
    End If
    
    VTCount = VTCount + 1
    
    Set oNodeLevel1 = oHL7.ReturnNewZVTNode(VTCount, VitalsRS, VitalSignsModRS, VitalSignsValRS)
    oRootNode.appendChild oNodeLevel1
    
    VitalsRS.MoveNext
  
  Loop
      
  Msg = "<?xml version=""1.0""?>"
  Msg = Msg & oRootNode.xml
      
  Dim enmResult As DMErrorCodes
   
  enmResult = PersistTMIPMessageToDatabase(strMessageGUID, "VITALS", "DATAID", VitalsRS.Clone, Msg, MsgType, , "TMIP_A04")
  
  If enmResult = ERROR_SUCCESS Then
    BuildTMIPVitalsMsg = ERROR_SUCCESS ' good return
  Else
    BuildTMIPVitalsMsg = enmResult
  End If
   
Exit Function

ErrTrap:
   Err.Clear
   BuildTMIPVitalsMsg = ERROR_UNHANDLED_SYSTEM_ERROR
End Function
      
Private Function GetRS(cObjs As Collection, TableName As String) As ADODB.Recordset
   
   Dim Data As I_RSObjects
   
   On Error GoTo GetRSError
   
   For Each Data In cObjs
      If UCase(Data.TableName) = UCase(TableName) Then
         Set GetRS = Data.DBRECORDSET
         Exit Function
      End If
   Next Data

Exit Function

GetRSError:
   Set GetRS = Nothing
End Function

Private Function DeleteRS(cObjs As Collection, TableName As String)

'SF 62831 Want to avoid duplicate recordsets

Dim Data As I_RSObjects
Dim iCounter As Long

   On Error GoTo GetRSError
   
   For iCounter = 1 To cObjs.Count
      Set Data = cObjs(iCounter)
      If UCase(Data.TableName) = UCase(TableName) Then
         cObjs.Remove (iCounter)
         Exit Function
      End If
   Next

Exit Function

GetRSError:
    MsgBox "Application error: " & Err.Description & " Occurred in: DMBE.DMBE_OPS.DeleteRS"
   
End Function

Private Function GetAssociatedRS(Value As Variant, FieldName As String, RS As ADODB.Recordset) As Long
   On Error GoTo GetAssociatedRSError
   If RS.BOF = True Then
      If RS.EOF = True Then
         GetAssociatedRS = ERROR_EMPTY_RECORDSET ' bad return
         Exit Function
      End If
   End If
   If FieldName = "" Then
      GetAssociatedRS = ERROR_UNABLE_TO_ACCESS_TABLE ' bad return
      Exit Function
   End If
   RS.MoveFirst
      
   Do
      If RS.Fields(FieldName).Value = Value Then
          GetAssociatedRS = ERROR_SUCCESS ' good return
          Exit Function
      End If
      RS.MoveNext
   Loop While RS.EOF = False
GetAssociatedRSError:
   GetAssociatedRS = ERROR_DATA_NOT_FOUND ' bad return
   Err.Clear
End Function


Private Function GetTableDataADO(ByVal sSQL As String, ByVal DSNConnect As String, RS As ADODB.Recordset) As Variant
'   Dim CN As ADODB.Connection
   On Error GoTo GetTableDataADOError
   
   ' make sure that we prove success
   GetTableDataADO = ERROR_UNABLE_TO_ACCESS_TABLE
   If oDas Is Nothing Then
      Set oDas = oDatabase.CHCSII_DAS(Auto) 'New GEMS_DAS
   End If
'   If sSQL <> "" Then
'      Set CN = New ADODB.Connection
'      CN.ConnectionString = DSNConnect
'      CN.Open
'   End If
   Set RS = oDas.OpenRecordset(sSQL, adOpenDynamic, adLockOptimistic)
'   RS.Open sSQL, CN, adOpenStatic, adLockOptimistic
   If RS.BOF = True And RS.EOF = True Then
      Set RS = Nothing
   Else
      GetTableDataADO = ERROR_SUCCESS
   End If
'   Set CN = Nothing
   Exit Function
GetTableDataADOError:
   If gobjShared Is Nothing Then
      Set gobjShared = New CWShared
   End If
   gobjShared.ShowVBError Err.Number, Err.Description, "GetTableDataADO Function", "DMBE", vbExclamation
   GetTableDataADO = -123454321
   Set RS = Nothing
End Function

Private Function ReadPIC(cRSObjs As Collection, DSet As DMDataSets, DFormat As DMFormats, strFilePath As String) As Long
   
   Dim rsData As ADODB.Recordset
   Dim objRSObject As I_RSObjects
   Dim lngSearchHandle As Long
   Dim strFileName As String
   Dim strFilePattern As String
   Dim lngReturnValue As Long
   Dim strPath As String
   Dim strTempPath As String
   Dim strPersistedFileType As String
   Dim lngSearchReturn As Long
   
   Dim oDom As MSXML2.DOMDocument30
   Dim oNode As MSXML2.IXMLDOMElement
   Dim oItem As MSXML2.IXMLDOMElement
   Dim oTempNode As MSXML2.IXMLDOMElement
   Dim bCompress As Boolean
   Dim bEncrypt As Boolean
   Dim asArray() As String
   Dim sData As String
   Dim nCnt As Long
   Dim nCol As Long
   Dim asFields() As String
   Dim asItems() As String
   Dim sSQL As String
   Dim sDoc As String
   
   On Error GoTo ReadPICError

   ' check for persisted recordset format
   Select Case DFormat
      Case eDM_ADT
        strPersistedFileType = ".ADT"
      Case eDM_XML
        strPersistedFileType = ".XML"
      Case eDM_DAT
        strPersistedFileType = ".DAT"
      Case Else
        ReadPIC = ERROR_INVALID_FORMAT ' bad return - data format not initialized
        Exit Function
   End Select
   ' get possible persisted recordset table names
   'strDataSet = GetDataSetCode()

    ' make empty collection
    While cRSObjs.Count <> 0
       cRSObjs.Remove 1
    Wend
    
    ' get upper bound for tabel name array
    ' Will check first for the name of the patient and dataset
    ' Note: here it is matched by unitnumber
    
    'Need to do better with unitnumber
    
    'strFilePattern = strFilePath & "\" & "*" & gobjPatient.UnitNumber & "*" & strDataSet & "*" & strPersistedFileType
    strFilePattern = strFilePath & "\" & m_sFileName & strPersistedFileType
    
    strFilePattern = Replace(strFilePattern, "\\", "\", 1, -1, vbTextCompare)
    strFileName = ""
    
    ' check if right PIC device is still there
    lngReturnValue = I_DMBE_ValidatePIC()
    If lngReturnValue <> ERROR_SUCCESS Then
       ReadPIC = lngReturnValue
       Exit Function
    End If
    
    'Go to the first file
    lngSearchReturn = GetFirstFile(strFilePattern, lngSearchHandle, strFileName)
    FindClose lngSearchHandle
    'Could not find a match by Patient Unit Number (ie. Unit Numbers Change from system to system)
    If lngSearchReturn = gclNOTFOUND Then
'        strFileName = GetPatientFile
    End If
    
    If lngSearchReturn = gclFOUND Then
       strPath = strFilePath & "\" & strFileName
       strPath = Replace(strPath, "\\", "\", 1, -1, vbTextCompare)
       
       'check for compression and encryption
       asArray = Split(strFileName, "_")
       
       If UBound(asArray) > 0 Then
         If Left(asArray(1), 1) = "C" Then
           bCompress = True
         End If
        
         If Mid(asArray(1), 2, 1) = "E" Then
           bEncrypt = True
         End If
       End If
       
       
       If bEncrypt Or bCompress Then
         sData = ReadFileContents(strPath, bCompress, bEncrypt, lngReturnValue)
       
         If lngReturnValue <> 0 Then
           ReadPIC = lngReturnValue
           Exit Function
         End If
       End If
       
        If DFormat = eDM_XML Or DFormat = eDM_ADT Then
          Set oDom = New MSXML2.DOMDocument30
          If bEncrypt Or bCompress Then
            If Not oDom.loadXML(sData) Then
              ReadPIC = ERROR_DAT_FILE_PROCESSING
              Exit Function
            End If
          Else 'PVF File
            If Not oDom.Load(strPath) Then
              sData = ReadFileContents(strPath, False, False, lngReturnValue)
              If lngReturnValue <> 0 Then
                ReadPIC = ERROR_DAT_FILE_PROCESSING
                Exit Function
              ElseIf Not oDom.loadXML(sData) Then
                ReadPIC = ERROR_DAT_FILE_PROCESSING
                Exit Function
              End If
            End If
          End If
          
          Set oNode = oDom.selectSingleNode("//RECORDSETS")
          For Each oItem In oNode.childNodes
            strTempPath = Environ("TEMP") & "\" & oItem.nodeName & strPersistedFileType
            strTempPath = Replace(strTempPath, "\\", "\", 1, -1, vbTextCompare)
            If DFormat = eDM_ADT Then
                sData = DecodeString(oItem.childNodes(0).Text, , True)
                Call PutBinaryData(strTempPath, sData)
            Else
                sData = oItem.childNodes(0).xml
                Call WriteFileContents(strTempPath, sData)
            End If
            Set rsData = New ADODB.Recordset
          
            If UCase(strPersistedFileType) = ".XML" Then
               rsData.Open strTempPath, , adOpenDynamic, adLockOptimistic, adPersistXML
            Else
               rsData.Open strTempPath, , adOpenDynamic, adLockOptimistic, adPersistADTG
            End If
            
            Select Case UCase$(oItem.nodeName)
            Case "ENC_RTFS", "ENC_SECTIONS", "LISTTOOLRECORD"
              If UCase$(oItem.nodeName) = "LISTTOOLRECORD" Then
                sDoc = "RTF"
              Else
                sDoc = "DOC"
              End If
              
              nCnt = 1
              If rsData.RecordCount > 0 Then
                rsData.MoveFirst
                Do While Not rsData.EOF
                  Set oTempNode = oItem.childNodes(nCnt)
                  ' SCR #39837;   Developer: Mag 08/12/2003 04:10 PM
                  If oTempNode Is Nothing Then
                    sData = ""
                  Else
                    sData = oTempNode.Text
                  End If
                  
                  If (rsData("OriginalSize") > 1) Then
                        
'                        'SF 1/10/05 added compression logic for testing - incluclusive results (may need to revisit)
'                        Dim objZlib As EncZLib
'                        Dim aBytes() As Byte
'
'                        Set objZlib = New EncZLib
'                        aBytes = StrConv(Replace(sData, Chr(0), ""), vbFromUnicode)
'                        objZlib.CompressData aBytes
'
'                         If rsData("OriginalSize") = 3 Then
                            sData = DecodeString(sData, rsData("OriginalSize"), True)
'                        End If
                        Call rsData(sDoc).AppendChunk(sData)
                       
                  Else
                    Call rsData(sDoc).AppendChunk(FilterXmlToRtf(sData))
                  End If
                  rsData.Update
                  rsData.MoveNext
                  nCnt = nCnt + 1
                Loop
              End If
            ' Begin:SCR #38841;   Developer: Mag 07/15/2003 12:43 PM
            'MMI_ID Fix SSN
            Case "MMI_ID"
              If rsData.RecordCount > 0 Then
                rsData.MoveFirst
                Do While Not rsData.EOF
                  If rsData("SSN") & "" <> "" Then
                    rsData("SSN") = Trim$(Replace$(rsData("SSN"), "-", ""))
                    If FieldExist(rsData, "PATIENT_FLAG") Then
                      rsData("PATIENT_FLAG") = "Y"
                    End If
                    rsData.Update
                  End If
                  rsData.MoveNext
                Loop
              End If
            ' End: SCR #38841;
            End Select
            Kill strTempPath
          
          ' create object for collection
          Set objRSObject = New RSObject_Ops
          Set objRSObject.DBRECORDSET = rsData
          objRSObject.TableName = oItem.nodeName
          cRSObjs.Add objRSObject, oItem.nodeName
          Set objRSObject = Nothing
          Set rsData = Nothing
        Next oItem
      ElseIf DFormat = eDM_DAT Then
          Set oDom = New MSXML2.DOMDocument30
          If Not oDom.loadXML(sData) Then
            ReadPIC = ERROR_DAT_FILE_PROCESSING
            Exit Function
          End If
      
          Set oNode = oDom.selectSingleNode("//RECORDSETS")
          For Each oItem In oNode.childNodes
            sData = FilterXmlToRtf(oItem.childNodes(0).xml)
            
            If oDas Is Nothing Then
              Set oDas = oDatabase.CHCSII_DAS(Auto) 'New GEMS_DAS
            End If
            
            'Get a Empty RS for Table
            sSQL = "Select * From " & oItem.nodeName & " Where 1 = 0"
            Set rsData = Nothing
            Set rsData = oDas.OpenRecordset(sSQL, adOpenDynamic, adLockOptimistic)
            Set rsData.ActiveConnection = Nothing
          
            asArray = Split(sData, DELIMITOR_CHARACTERS & vbCrLf)
            asFields = Split(asArray(0), DELIMITOR_CHARACTERS)
            
            If UBound(asArray) > 0 Then
                On Error Resume Next
                For nCnt = 1 To UBound(asArray)
                    If Not Trim(asArray(nCnt)) = "" Then
                      rsData.AddNew
                      asItems = Split(asArray(nCnt), DELIMITOR_CHARACTERS)
                      For nCol = 0 To UBound(asItems)
                        If rsData(asFields(nCol)).Type = adLongVarBinary Then
                          Select Case UCase$(oItem.nodeName)
                          Case "ENC_RTFS", "ENC_SECTIONS", "LISTTOOLRECORD"
                            If rsData("OriginalSize").Value > 1 Then
                              Call rsData(asFields(nCol)).AppendChunk(DecodeString(asItems(nCol)))
                            Else
                              rsData(asFields(nCol)).Value = FilterXmlToRtf(asItems(nCol))
                            End If
                          Case Else
                              rsData(asFields(nCol)).Value = FilterXmlToRtf(asItems(nCol))
                          End Select
                        Else
                          rsData(asFields(nCol)).Value = asItems(nCol)
                        End If
                      Next nCol
                      rsData.Update
                    End If
                Next nCnt
                
            End If
            
            On Error GoTo ReadPICError
          
            ' create object for collection
            Set objRSObject = New RSObject_Ops
            Set objRSObject.DBRECORDSET = rsData
            objRSObject.TableName = oItem.nodeName
            cRSObjs.Add objRSObject, oItem.nodeName
            Set objRSObject = Nothing
            Set rsData = Nothing
        Next oItem
      End If

    Else
       ' flag an error here because record not found on PIC device but continue processing
       Call MsgBox("Record not found on PIC device but continue processing", vbCritical, "ReadPIC Function")
    End If
    
   If cRSObjs.Count <> 0 Then
   
      'SCR 24615 CC
      lngReturnValue = ValidatePICFacilityAndClinic(cRSObjs)
        If lngReturnValue = DMErrorCodes.ERROR_SUCCESS Then
          ReadPIC = ERROR_SUCCESS ' good return
        Else
          ReadPIC = lngReturnValue
        End If
          
    '      If m_enmDataSet = eDM_Encounters Then
    '         Set m_colPatientList = cRSObjs
    '      End If
   Else
      ReadPIC = ERROR_NO_RECORDS_FOR_DATASET_FOUND ' bad return
   End If

Exit Function

ReadPICError:
   If Err.Number = 53 Then ' file does not exist
      Err.Clear
      Resume Next
   End If
   If gobjShared Is Nothing Then
      Set gobjShared = New CWShared
   End If
   
   'Clean up
  gobjShared.ShowVBError Err.Number, Err.Description, "ReadPIC Function", "DMBE", vbExclamation
   Set rsData = Nothing
   Set objRSObject = Nothing

Exit Function
  Resume
End Function
'<SCR 24615 CC>
Private Function ValidatePICFacilityAndClinic(ByRef colRSObjects As Collection) As DMErrorCodes

    Dim objSQL As ICHCSII_SQl.ISqlOpsEx
    Dim objData As IDMBE.I_RSObjects
    
    On Error GoTo ErrHandler
    
    
    Set objSQL = oDatabase.CHCSII_SQLOPS(Auto) 'New GEMS_SQL.SqlOps
    
    For Each objData In colRSObjects
        
        Select Case UCase(objData.TableName)
            
            Case Is = "ALLERGIES"
                'FACILITY_NCID, POINTOFCAREFACILITYNCID
                If ValidatePICFacility(objData.DBRECORDSET, "FACILITY_NCID") = False Then
                    ValidatePICFacilityAndClinic = ERROR_PIC_INVALID_FACILITY
                    Exit Function
                End If
                If ValidatePICFacility(objData.DBRECORDSET, "POINTOFCAREFACILITYNCID") = False Then
                    ValidatePICFacilityAndClinic = ERROR_PIC_INVALID_FACILITY
                    Exit Function
                End If
                        
            Case Is = "CLINNOTE"
                'ENCFACILITY and POCFACILITY
                If ValidatePICFacility(objData.DBRECORDSET, "ENCFACILITY") = False Then
                    ValidatePICFacilityAndClinic = ERROR_PIC_INVALID_FACILITY
                    Exit Function
                End If
                If ValidatePICFacility(objData.DBRECORDSET, "POCFACILITY") = False Then
                    ValidatePICFacilityAndClinic = ERROR_PIC_INVALID_FACILITY
                    Exit Function
                End If
                
            Case Is = "COMPLAINTS"
                'POINTOFCAREFACILITYNCID
                If ValidatePICFacility(objData.DBRECORDSET, "POINTOFCAREFACILITYNCID") = False Then
                    ValidatePICFacilityAndClinic = ERROR_PIC_INVALID_FACILITY
                    Exit Function
                End If
                
            Case Is = "DIAGNOSES"
                'FACILITYNCID,POINTOFCAREFACILITYNCID
                If ValidatePICFacility(objData.DBRECORDSET, "FACILITYNCID") = False Then
                    ValidatePICFacilityAndClinic = ERROR_PIC_INVALID_FACILITY
                    Exit Function
                End If
                If ValidatePICFacility(objData.DBRECORDSET, "POINTOFCAREFACILITYNCID") = False Then
                    ValidatePICFacilityAndClinic = ERROR_PIC_INVALID_FACILITY
                    Exit Function
                End If
                
            Case ",ENC_DIAGNOSIS"
                ',FACILITYNCID,
                If ValidatePICFacility(objData.DBRECORDSET, "FACILITYNCID") = False Then
                    ValidatePICFacilityAndClinic = ERROR_PIC_INVALID_FACILITY
                    Exit Function
                End If
                
            Case ",ENC_DISPOSITION"
                'FACILITYNCID
                If ValidatePICFacility(objData.DBRECORDSET, "FACILITYNCID") = False Then
                    ValidatePICFacilityAndClinic = ERROR_PIC_INVALID_FACILITY
                    Exit Function
                End If
                
            Case "ENC_RTFS"
                'FACILITYNCID
                If ValidatePICFacility(objData.DBRECORDSET, "FACILITYNCID") = False Then
                    ValidatePICFacilityAndClinic = ERROR_PIC_INVALID_FACILITY
                    Exit Function
                End If
                
            Case "ENC_SECTIONS"
                'FACILITYNCID
                If ValidatePICFacility(objData.DBRECORDSET, "FACILITYNCID") = False Then
                    ValidatePICFacilityAndClinic = ERROR_PIC_INVALID_FACILITY
                    Exit Function
                End If
                
            Case Is = "ENCOUNTERS"
                'FACILITYNCID, CLINICNCID
                If ValidatePICFacility(objData.DBRECORDSET, "FACILITYNCID") = False Then
                    ValidatePICFacilityAndClinic = ERROR_PIC_INVALID_FACILITY
                    Exit Function
                End If
                If ValidatePICClinic(objData.DBRECORDSET, "FACILITYNCID", "CLINICNCID") = False Then
                    ValidatePICFacilityAndClinic = ERROR_PIC_INVALID_CLINIC
                    Exit Function
                End If
                
            Case "ENCTEXTNOTES"
                'FACILITYNCID
                If ValidatePICFacility(objData.DBRECORDSET, "FACILITYNCID") = False Then
                    ValidatePICFacilityAndClinic = ERROR_PIC_INVALID_FACILITY
                    Exit Function
                End If
                
            Case Is = "LISTTOOLRECORD"
                'FACILITYNCID, CLINICNCID
                If ValidatePICFacility(objData.DBRECORDSET, "FACILITYNCID") = False Then
                    ValidatePICFacilityAndClinic = ERROR_PIC_INVALID_FACILITY
                    Exit Function
                End If
                If ValidatePICClinic(objData.DBRECORDSET, "FACILITYNCID", "CLINICNCID") = False Then
                    ValidatePICFacilityAndClinic = ERROR_PIC_INVALID_CLINIC
                    Exit Function
                End If
                
            Case Is = "MEDICATIONS"
                'DISPENSINGLOCATIONNCID, CLINICNCID
                If ValidatePICFacility(objData.DBRECORDSET, "DISPENSINGLOCATIONNCID") = False Then
                    ValidatePICFacilityAndClinic = ERROR_PIC_INVALID_FACILITY
                    Exit Function
                End If
                If ValidatePICClinic(objData.DBRECORDSET, "DISPENSINGLOCATIONNCID", "CLINICNCID") = False Then
                    ValidatePICFacilityAndClinic = ERROR_PIC_INVALID_CLINIC
                    Exit Function
                End If
                
            Case Is = "PROBLEMS"
                'FACILITYNCID, POINTOFCAREFACILITYNCID
                If ValidatePICFacility(objData.DBRECORDSET, "FACILITYNCID") = False Then
                    ValidatePICFacilityAndClinic = ERROR_PIC_INVALID_FACILITY
                    Exit Function
                End If
                If ValidatePICFacility(objData.DBRECORDSET, "POINTOFCAREFACILITYNCID") = False Then
                    ValidatePICFacilityAndClinic = ERROR_PIC_INVALID_FACILITY
                    Exit Function
                End If
                
            Case Is = "PROCEDURES"
                'FACILITYNCID, POINTOFCAREFACILITYNCID
                If ValidatePICFacility(objData.DBRECORDSET, "FACILITYNCID") = False Then
                    ValidatePICFacilityAndClinic = ERROR_PIC_INVALID_FACILITY
                    Exit Function
                End If
                If ValidatePICFacility(objData.DBRECORDSET, "POINTOFCAREFACILITYNCID") = False Then
                    ValidatePICFacilityAndClinic = ERROR_PIC_INVALID_FACILITY
                    Exit Function
                End If
                
            Case Is = "SONOTES"
                'FACILITYNCID,
                If ValidatePICFacility(objData.DBRECORDSET, "FACILITYNCID") = False Then
                    ValidatePICFacilityAndClinic = ERROR_PIC_INVALID_FACILITY
                    Exit Function
                End If
                
            Case Is = "VITALS"
                'FACILITYNCID
                If ValidatePICFacility(objData.DBRECORDSET, "FACILITYNCID") = False Then
                    ValidatePICFacilityAndClinic = ERROR_PIC_INVALID_FACILITY
                    Exit Function
                End If
            
            Case Is = "MMI_SITE"
                'FACILITY_CID
                If ValidatePICFacility(objData.DBRECORDSET, "FACILITY_CID") = False Then
                    ValidatePICFacilityAndClinic = ERROR_PIC_INVALID_FACILITY
                    Exit Function
                End If
            
        End Select
    Next objData

    ValidatePICFacilityAndClinic = ERROR_SUCCESS

Exit Function

ErrHandler:
    ValidatePICFacilityAndClinic = ERROR_PIC_INVALID_FACILITY
    MsgBox "Error #: " & Err.Number & Chr(10) & Err.Description & ". Occured in ValidatePICFacilityAndClinic.", vbExclamation, "DMBE"
    Err.Clear
End Function

Private Function ValidatePICFacility(ByRef rsData As ADODB.Recordset, ByVal strFieldName As String) As Boolean

    Dim objSQL As ICHCSII_SQl.ISqlOpsEx
    Dim lngNCID As Long
    
    ValidatePICFacility = False
    Set objSQL = oDatabase.CHCSII_SQLOPS(Auto) 'New GEMS_SQL.SqlOps
    
    If rsData Is Nothing Then
        ValidatePICFacility = True
        Exit Function
    End If
    
    With rsData
        If Not .RecordCount = 0 Then
            .MoveFirst
            Do Until .EOF
                If Not IsNull(.Fields(strFieldName).Value) Then
                    lngNCID = .Fields(strFieldName).Value
                    If lngNCID <> 0 Then
                        If objSQL.ValidateFacility2(lngNCID, "", True) = False Then
                            ValidatePICFacility = False
                            Exit Function
                        End If
                    End If
                End If
                .MoveNext
            Loop
        End If
    End With
    
    ValidatePICFacility = True
    
End Function

Private Function ValidatePICClinic(ByRef rsData As ADODB.Recordset, ByVal strFacilityFieldName As String, ByVal strClinicFieldName As String) As Boolean

    Dim objSQL As ICHCSII_SQl.ISqlOpsEx
    Dim lngFacilityNCID As Long
    Dim lngClinicNCID As Long
    
    ValidatePICClinic = False
    Set objSQL = oDatabase.CHCSII_SQLOPS(Auto) 'New GEMS_SQL.SqlOps
    
    If rsData Is Nothing Then
        ValidatePICClinic = True
        Exit Function
    End If
    
    With rsData
        If Not .RecordCount = 0 Then
            .MoveFirst
            Do Until .EOF
                If Not IsNull(.Fields(strFacilityFieldName).Value) Then
                    lngFacilityNCID = .Fields(strFacilityFieldName).Value
                End If
                If Not IsNull(.Fields(strClinicFieldName).Value) Then
                    lngClinicNCID = .Fields(strClinicFieldName).Value
                End If
                If lngFacilityNCID = 0 Or lngClinicNCID = 0 Then
                    ValidatePICClinic = True
                    Exit Function
                Else
                    If objSQL.ValidateFacility2(lngFacilityNCID, "", True, lngClinicNCID, "") = False Then
                        ValidatePICClinic = False
                        Exit Function
                    End If
                End If
                .MoveNext
            Loop
        End If
    End With
    
    ValidatePICClinic = True
    
End Function
'</SCR 24615 CC>

Private Function TranslateProviderSSNToNCID(ByRef colRSObjects As Collection) As DMErrorCodes
    
    'This is used on the way "In" from PIC
    
    Dim strSSN As String
    Dim lngNCID As Long
    Dim objData As IDMBE.I_RSObjects
    
    For Each objData In colRSObjects
        Select Case UCase(objData.TableName)
            Case Is = "ALLERGIES"
                'CLINICIANNCID
                If Not IsNull(objData.DBRECORDSET.Fields("CLINICIANNCID").Value) Then
                    strSSN = objData.DBRECORDSET.Fields("CLINICIANNCID").Value
                    lngNCID = GetProviderNCIDFromSSN(strSSN)
                    If lngNCID <> 0 Then
                        objData.DBRECORDSET.Fields("CLINICIANNCID").Value = CLng(lngNCID)
                        objData.DBRECORDSET.Update
                    End If
                End If
            Case Is = "COMPLAINTS"
                'CLINICIANNCID
                If Not IsNull(objData.DBRECORDSET.Fields("CLINICIANNCID").Value) Then
                    strSSN = objData.DBRECORDSET.Fields("CLINICIANNCID").Value
                    lngNCID = GetProviderSSNFromNCID(strSSN)
                    If lngNCID <> 0 Then
                        objData.DBRECORDSET.Fields("CLINICIANNCID").Value = CLng(lngNCID)
                        objData.DBRECORDSET.Update
                    End If
                End If
            Case Is = "DIAGNOSES"
                'CLINICIANNCID
                If Not IsNull(objData.DBRECORDSET.Fields("CLINICIANNCID").Value) Then
                    strSSN = objData.DBRECORDSET.Fields("CLINICIANNCID").Value
                    lngNCID = GetProviderSSNFromNCID(strSSN)
                    If lngNCID <> 0 Then
                        objData.DBRECORDSET.Fields("CLINICIANNCID").Value = CLng(lngNCID)
                        objData.DBRECORDSET.Update
                    End If
                End If
            Case Is = "ENCOUNTERS"
                'WHOCOSIGNEDNCID,WHOSIGNEDNCID,COSIGNERNCID,
                If Not IsNull(objData.DBRECORDSET.Fields("WHOCOSIGNEDNCID").Value) Then
                    strSSN = objData.DBRECORDSET.Fields("WHOCOSIGNEDNCID").Value
                    lngNCID = GetProviderSSNFromNCID(strSSN)
                    If lngNCID <> 0 Then
                        objData.DBRECORDSET.Fields("WHOCOSIGNEDNCID").Value = CLng(lngNCID)
                        objData.DBRECORDSET.Update
                    End If
                End If
                If Not IsNull(objData.DBRECORDSET.Fields("WHOSIGNEDNCID").Value) Then
                    strSSN = objData.DBRECORDSET.Fields("WHOSIGNEDNCID").Value
                    lngNCID = GetProviderSSNFromNCID(strSSN)
                    If lngNCID <> 0 Then
                        objData.DBRECORDSET.Fields("WHOSIGNEDNCID").Value = CLng(lngNCID)
                        objData.DBRECORDSET.Update
                    End If
                End If
                If Not IsNull(objData.DBRECORDSET.Fields("COSIGNERNCID").Value) Then
                    strSSN = objData.DBRECORDSET.Fields("COSIGNERNCID").Value
                    lngNCID = GetProviderSSNFromNCID(strSSN)
                    If lngNCID <> 0 Then
                        objData.DBRECORDSET.Fields("COSIGNERNCID").Value = CLng(lngNCID)
                        objData.DBRECORDSET.Update
                    End If
                End If
            Case Is = "IMM_PEOPLE_VACCINE"
                'PROVIDERID
                If Not IsNull(objData.DBRECORDSET.Fields("PROVIDERID").Value) Then
                    strSSN = objData.DBRECORDSET.Fields("PROVIDERID").Value
                    lngNCID = GetProviderSSNFromNCID(strSSN)
                    If lngNCID <> 0 Then
                        objData.DBRECORDSET.Fields("PROVIDERID").Value = CLng(lngNCID)
                        objData.DBRECORDSET.Update
                    End If
                End If
            Case Is = "LISTTOOLRECORD"
                'CLINICIANNCID
                If Not IsNull(objData.DBRECORDSET.Fields("CLINICIANNCID").Value) Then
                    strSSN = objData.DBRECORDSET.Fields("CLINICIANNCID").Value
                    lngNCID = GetProviderSSNFromNCID(strSSN)
                    If lngNCID <> 0 Then
                        objData.DBRECORDSET.Fields("CLINICIANNCID").Value = CLng(lngNCID)
                        objData.DBRECORDSET.Update
                    End If
                End If
            Case Is = "PROBLEMS"
                'CLINICIANNCID
                If Not IsNull(objData.DBRECORDSET.Fields("CLINICIANNCID").Value) Then
                    strSSN = objData.DBRECORDSET.Fields("CLINICIANNCID").Value
                    lngNCID = GetProviderSSNFromNCID(strSSN)
                    If lngNCID <> 0 Then
                        objData.DBRECORDSET.Fields("CLINICIANNCID").Value = CLng(lngNCID)
                        objData.DBRECORDSET.Update
                    End If
                End If
            Case Is = "PROCEDURES"
                'CLINICIANNCID
                If Not IsNull(objData.DBRECORDSET.Fields("CLINICIANNCID").Value) Then
                    strSSN = objData.DBRECORDSET.Fields("CLINICIANNCID").Value
                    lngNCID = GetProviderSSNFromNCID(strSSN)
                    If lngNCID <> 0 Then
                        objData.DBRECORDSET.Fields("CLINICIANNCID").Value = CLng(lngNCID)
                        objData.DBRECORDSET.Update
                    End If
                End If
            Case Is = "VITALS"
                'CLINICIANNCID
                If Not IsNull(objData.DBRECORDSET.Fields("CLINICIANNCID").Value) Then
                    strSSN = objData.DBRECORDSET.Fields("CLINICIANNCID").Value
                    lngNCID = GetProviderSSNFromNCID(strSSN)
                    If lngNCID <> 0 Then
                        objData.DBRECORDSET.Fields("CLINICIANNCID").Value = CLng(lngNCID)
                        objData.DBRECORDSET.Update
                    End If
                End If
            Case Is = "WORK_STATUS"
                'PROVIDER_NCID
                If Not IsNull(objData.DBRECORDSET.Fields("PROVIDER_NCID").Value) Then
                    strSSN = objData.DBRECORDSET.Fields("PROVIDER_NCID").Value
                    lngNCID = GetProviderSSNFromNCID(strSSN)
                    If lngNCID <> 0 Then
                        objData.DBRECORDSET.Fields("PROVIDER_NCID").Value = CLng(lngNCID)
                        objData.DBRECORDSET.Update
                    End If
                End If
        End Select
    Next objData
    
    
    TranslateProviderSSNToNCID = ERROR_SUCCESS
        
End Function


Private Function GetProviderNCIDFromSSN(ByVal strSSN As String) As Long

    'This is used on the way "In" From PIC

    Dim rsProvider As ADODB.Recordset
    Dim strSQL As String
    Dim objDAS As ICHCSII_DAS
    Dim lngNCID As Long
    Dim objSQL As ICHCSII_SQl.ISqlOpsEx
    
    strSQL = "SELECT NCID FROM PROVIDER WHERE SSN = '" & strSSN & "'"
    Set objDAS = oDatabase.CHCSII_DAS(Auto) 'New GEMS_DAS
    Set rsProvider = objDAS.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)
    If Not rsProvider Is Nothing Then
        If rsProvider.EOF = False Then
            If Not IsNull(rsProvider.Fields(0).Value) Then
                lngNCID = rsProvider.Fields(0).Value
                
            End If
        End If
        
    Else
        'Need to add a new provider
        'Need to assign new NCID
        
        Set objSQL = oDatabase.CHCSII_SQLOPS(Auto) 'New GEMS_SQL.SqlOps
        
        'SF SCR 62192, 65523 2-21-05 This needs to share the same sequence as the patient because the of the assumptions in the concept control.
        lngNCID = objDAS.GetNextID(esUNIT_NUMBER)
        'lngNCID = objDAS.GetNextNcid
 
        If objSQL.ValidateProvider2(CStr(lngNCID), strSSN, "Unknown", "", "Theatre Provicer", True) = False Then
            'Validation Failed
            lngNCID = 0
        End If
        
    End If
    
    Set rsProvider = Nothing
    Set objDAS = Nothing

    GetProviderNCIDFromSSN = lngNCID

End Function

Private Function TranslateProviderNCIDToSSN(ByRef colRSObjects As Collection) As Long
    
    'This is used on the way "out" to PIC
    
    Dim lngNCID As Long
    Dim strSSN As String
    Dim objData As IDMBE.I_RSObjects
    
    For Each objData In colRSObjects
        Select Case UCase(objData.TableName)
            Case Is = "ALLERGIES"
                'CLINICIANNCID
                If Not IsNull(objData.DBRECORDSET.Fields("CLINICIANNCID").Value) Then
                    lngNCID = objData.DBRECORDSET.Fields("CLINICIANNCID").Value
                    strSSN = GetProviderSSNFromNCID(lngNCID)
                    If strSSN <> "" Then
                        objData.DBRECORDSET.Fields("CLINICIANNCID").Value = CLng(strSSN)
                        objData.DBRECORDSET.Update
                    End If
                End If
            Case Is = "CLINNOTE"
                'CLINICIAN
                
            Case Is = "COMPLAINTS"
                'CLINICIANNCID
                If Not IsNull(objData.DBRECORDSET.Fields("CLINICIANNCID").Value) Then
                    lngNCID = objData.DBRECORDSET.Fields("CLINICIANNCID").Value
                    strSSN = GetProviderSSNFromNCID(lngNCID)
                    If strSSN <> "" Then
                        objData.DBRECORDSET.Fields("CLINICIANNCID").Value = CLng(strSSN)
                        objData.DBRECORDSET.Update
                    End If
                End If
            Case Is = "DIAGNOSES"
                'CLINICIANNCID
                If Not IsNull(objData.DBRECORDSET.Fields("CLINICIANNCID").Value) Then
                    lngNCID = objData.DBRECORDSET.Fields("CLINICIANNCID").Value
                    strSSN = GetProviderSSNFromNCID(lngNCID)
                    If strSSN <> "" Then
                        objData.DBRECORDSET.Fields("CLINICIANNCID").Value = CLng(strSSN)
                        objData.DBRECORDSET.Update
                    End If
                End If
            Case Is = "ENCOUNTERS"
                'WHOCOSIGNEDNCID,WHOSIGNEDNCID,COSIGNERNCID,
                If Not IsNull(objData.DBRECORDSET.Fields("WHOCOSIGNEDNCID").Value) Then
                    lngNCID = objData.DBRECORDSET.Fields("WHOCOSIGNEDNCID").Value
                    strSSN = GetProviderSSNFromNCID(lngNCID)
                    If strSSN <> "" Then
                        objData.DBRECORDSET.Fields("WHOCOSIGNEDNCID").Value = CLng(strSSN)
                        objData.DBRECORDSET.Update
                    End If
                End If
                If Not IsNull(objData.DBRECORDSET.Fields("WHOSIGNEDNCID").Value) Then
                    lngNCID = objData.DBRECORDSET.Fields("WHOSIGNEDNCID").Value
                    strSSN = GetProviderSSNFromNCID(lngNCID)
                    If strSSN <> "" Then
                        objData.DBRECORDSET.Fields("WHOSIGNEDNCID").Value = CLng(strSSN)
                        objData.DBRECORDSET.Update
                    End If
                End If
                If Not IsNull(objData.DBRECORDSET.Fields("COSIGNERNCID").Value) Then
                    lngNCID = objData.DBRECORDSET.Fields("COSIGNERNCID").Value
                    strSSN = GetProviderSSNFromNCID(lngNCID)
                    If strSSN <> "" Then
                        objData.DBRECORDSET.Fields("COSIGNERNCID").Value = CLng(strSSN)
                        objData.DBRECORDSET.Update
                    End If
                End If
            Case Is = "IMM_PEOPLE_VACCINE"
                'PROVIDERID
                If Not IsNull(objData.DBRECORDSET.Fields("PROVIDERID").Value) Then
                    lngNCID = objData.DBRECORDSET.Fields("PROVIDERID").Value
                    strSSN = GetProviderSSNFromNCID(lngNCID)
                    If strSSN <> "" Then
                        objData.DBRECORDSET.Fields("PROVIDERID").Value = CLng(strSSN)
                        objData.DBRECORDSET.Update
                    End If
                End If
            Case Is = "LISTTOOLRECORD"
                'CLINICIANNCID
                If Not IsNull(objData.DBRECORDSET.Fields("CLINICIANNCID").Value) Then
                    lngNCID = objData.DBRECORDSET.Fields("CLINICIANNCID").Value
                    strSSN = GetProviderSSNFromNCID(lngNCID)
                    If strSSN <> "" Then
                        objData.DBRECORDSET.Fields("CLINICIANNCID").Value = CLng(strSSN)
                        objData.DBRECORDSET.Update
                    End If
                End If
            Case Is = "MEDICATIONS"
                'ORDERINGPROVIDERNCID,
            
            Case Is = "PROBLEMS"
                'CLINICIANNCID
                If Not IsNull(objData.DBRECORDSET.Fields("CLINICIANNCID").Value) Then
                    lngNCID = objData.DBRECORDSET.Fields("CLINICIANNCID").Value
                    strSSN = GetProviderSSNFromNCID(lngNCID)
                    If strSSN <> "" Then
                        objData.DBRECORDSET.Fields("CLINICIANNCID").Value = CLng(strSSN)
                        objData.DBRECORDSET.Update
                    End If
                End If
            Case Is = "PROCEDURES"
                'CLINICIANNCID
                If Not IsNull(objData.DBRECORDSET.Fields("CLINICIANNCID").Value) Then
                    lngNCID = objData.DBRECORDSET.Fields("CLINICIANNCID").Value
                    strSSN = GetProviderSSNFromNCID(lngNCID)
                    If strSSN <> "" Then
                        objData.DBRECORDSET.Fields("CLINICIANNCID").Value = CLng(strSSN)
                        objData.DBRECORDSET.Update
                    End If
                End If
            Case Is = "VITALS"
                'CLINICIANNCID
                If Not IsNull(objData.DBRECORDSET.Fields("CLINICIANNCID").Value) Then
                    lngNCID = objData.DBRECORDSET.Fields("CLINICIANNCID").Value
                    strSSN = GetProviderSSNFromNCID(lngNCID)
                    If strSSN <> "" Then
                        objData.DBRECORDSET.Fields("CLINICIANNCID").Value = CLng(strSSN)
                        objData.DBRECORDSET.Update
                    End If
                End If
            Case Is = "WORK_STATUS"
                'PROVIDER_NCID
                If Not IsNull(objData.DBRECORDSET.Fields("PROVIDER_NCID").Value) Then
                    lngNCID = objData.DBRECORDSET.Fields("PROVIDER_NCID").Value
                    strSSN = GetProviderSSNFromNCID(lngNCID)
                    If strSSN <> "" Then
                        objData.DBRECORDSET.Fields("PROVIDER_NCID").Value = CLng(strSSN)
                        objData.DBRECORDSET.Update
                    End If
                End If
        End Select
    Next objData
    
    TranslateProviderNCIDToSSN = ERROR_SUCCESS
    
End Function

Private Function GetProviderSSNFromNCID(ByVal lngProviderNCID As Long) As String

    'This is used on the way "out" to PIC
    
    Dim rsProvider As ADODB.Recordset
    Dim strSQL As String
    Dim objDAS As ICHCSII_DAS
    Dim strSSN As String
    
    strSQL = "SELECT SSN FROM PROVIDER WHERE NCID =  & lngProviderNCID "
    Set objDAS = oDatabase.CHCSII_DAS(Auto) 'New GEMS_DAS
    Set rsProvider = objDAS.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)

    strSSN = ""
    
    If Not rsProvider Is Nothing Then
        If rsProvider.EOF = False Then
            If Not IsNull(rsProvider.Fields(0).Value) Then
                strSSN = rsProvider.Fields(0).Value
                strSSN = Replace(strSSN, "-", "")
                strSSN = Replace(strSSN, " ", "")
            End If
        End If
    End If
    
    Set rsProvider = Nothing
    Set objDAS = Nothing
    
    GetProviderSSNFromNCID = strSSN
    
End Function



'Private Function ProcessDATFile(cRSObjs As Collection, strDataSet As String, FilePath As String) As Long
'   Const PersistedFileType As String = ".DAT"
'   Dim FilePattern As String
'   Dim hSearch As Long
'   Dim FileName As String
'   Dim Path As String
'   Dim TempPath As String
'   Dim fs As FileSystemObject
'   Dim PICFileData As String
'   Dim EncryptCompressFlag As String
'   Dim DataLine() As String
'   Dim TableName As String
'   Dim Version As String
'   Dim FieldNames As String
'   Dim DataFields() As String
'   Dim DataValues() As String
'   Dim UpperTN As Long
'   Dim LastLine As Long
'   Dim ColonPos As Long
'   Dim DelimitorPos As Long
'   Dim oRS As ADODB.Recordset
'   Dim I As Long
'   Dim J As Long
'   Dim RetVal As Long
'   Dim Data As I_RSObjects
'   Dim ByPassRSFlag As Long
'   Dim ColumnStr As String
'   Dim lUIdx As Long
'   Dim lPos As Long
'   Dim sTemp() As String
'   Dim strOldTableName As String
'
'   On Error GoTo ProcessDATFileErrTrap
'
'   ' find the DAT file on the PIC device - should only be one per clinical area
'   I = 0
''   UpperTN = UBound(TN) - 1
''   FilePattern = FilePath & "\" & "*" & TN(I) & "*" & PersistedFileType  ' should use PersistedFileType here?
''   FilePattern = Replace(FilePattern, "\\", "\", 1, -1, vbTextCompare)
'   FilePattern = FilePath & "\" & "*" & gobjPatient.UnitNumber & "*" & strDataSet & "*" & PersistedFileType
'   FilePattern = Replace(FilePattern, "\\", "\", 1, -1, vbTextCompare)
'   FileName = ""
''   FileName = ""
'   ' check if right PIC device is still there
'   RetVal = I_DMBE_ValidatePIC()
'   If RetVal <> ERROR_SUCCESS Then
'      ProcessDATFile = RetVal
'      Exit Function
'   End If
'   RetVal = GetFirstFile(FilePattern, hSearch, FileName)
'   If RetVal = gclFOUND Then
'      Path = FilePath & "\" & FileName
'      Path = Replace(Path, "\\", "\", 1, -1, vbTextCompare)
'      TempPath = Environ("TEMP") & "\" & strDataSet & PersistedFileType
'      TempPath = Replace(TempPath, "\\", "\", 1, -1, vbTextCompare)
'      Set fs = New FileSystemObject
'      If fs.FileExists(TempPath) Then
'         fs.DeleteFile (TempPath)
'      End If
'      fs.CopyFile Path, TempPath, True
'      Set fs = Nothing
'      ' check for encryption and/or compression of DAT file
'      EncryptCompressFlag = Mid(FileName, 1, 1)
'      If EncryptCompressFlag = "E" Then
'         RetVal = DecryptDataFile(TempPath)   ' decrypt the DAT file
'         If RetVal <> 0 Then
'            ProcessDATFile = RetVal
'            Exit Function
'         End If
'      End If
'      EncryptCompressFlag = Mid(FileName, 2, 1)
'      If EncryptCompressFlag = "C" Then
'         RetVal = DecompressDataFile(TempPath)   ' decompress the DAT file
'         If RetVal <> 0 Then
'            ProcessDATFile = RetVal
'            Exit Function
'         End If
'      End If
'      ' check if right PIC device is still there
'      RetVal = I_DMBE_ValidatePIC()
'      If RetVal <> ERROR_SUCCESS Then
'         ProcessDATFile = RetVal
'         Exit Function
'      End If
'
'      ' read DAT file
'      PICFileData = ReadFileContents(TempPath)
'      If PICFileData <> "" Then
'         ' get the table lines of the DAT file
'         DataLine = Split(PICFileData, vbCr, -1, vbTextCompare)
'         LastLine = UBound(DataLine) - 1
'         ' get first recordset
'
'         ByPassRSFlag = 0
'
'         For I = 0 To LastLine
'
'            Set oRS = New ADODB.Recordset
'
'            ' get the values from a table line
'            DataValues = Split(DataLine(I), DELIMITOR_CONSTANT, -1, vbTextCompare)
'            ColonPos = InStr(1, DataValues(0), ":", vbTextCompare)
'            ' get the table name
'            strOldTableName = TableName
'            TableName = Mid(DataValues(0), 1, ColonPos - 1)
'            DelimitorPos = InStr(1, DataLine(I), DELIMITOR_CONSTANT, vbTextCompare)
'            ' get the field mapping table version
'            Version = Mid(DataLine(I), ColonPos + 1, DelimitorPos - ColonPos - 1)
'
'
'            If strOldTableName <> TableName Then
'               FieldNames = GetFieldNames(TableName, Version)
'               If FieldNames = "" Then
'                  ProcessDATFile = ERROR_FIELD_MAPPING_TABLE
'                  Exit Function
'               End If
'               DataFields = Split(FieldNames, DELIMITOR_CONSTANT, -1, vbTextCompare)
'               lUIdx = UBound(DataFields) - 1
'               For J = 0 To lUIdx
'                  sTemp = Split(DataFields(J), ",", -1, vbTextCompare)
'                  DataFields(J) = sTemp(0)
'               Next J
'               ' build recordset
'               RetVal = CreateDisconRS(FieldNames, oRS)
''               RetVal = GetFieldProps(TableName, oRS)
'               If RetVal <> ERROR_SUCCESS Then
'                  ProcessDATFile = RetVal  ' bad return
'                  Exit Function
'               End If
'            End If
'            ' put data into the recordset
'            RetVal = PopulateRS(oRS, DataValues(), DataFields())
'            If RetVal <> ERROR_SUCCESS Then
'               ProcessDATFile = ERROR_INVALID_RECORDSET_OPERATION
'               Exit Function
'            End If
'
'            If I < LastLine Then
'               If InStr(1, UCase(DataLine(I + 1)), UCase(TableName), vbTextCompare) = 0 Then
'                  ' create object for collection
'                  Set Data = New RSObject_Ops
'                  Set Data.DBRECORDSET = oRS
'                  ' special case
'                  If TableName = "IMMUNIZATIONS" Then
'                     Data.TableName = "IMMUNIZATION"
'                  Else
'                     Data.TableName = TableName
'                  End If
'                  cRSObjs.Add Data
'                  Set Data = Nothing
'                  Set oRS = Nothing
'                  Set oRS = New ADODB.Recordset
'                  ByPassRSFlag = 0
'               Else
'                  ByPassRSFlag = 1
'               End If
'            Else
'               If I = LastLine Then
'                  ' create object for collection
'                  Set Data = New RSObject_Ops
'                  Set Data.DBRECORDSET = oRS
'                  Data.TableName = TableName
'                  cRSObjs.Add Data
'                  Set Data = Nothing
'                  Set oRS = Nothing
'               End If
'            End If
'         Next I
'      Else
'         ProcessDATFile = ERROR_NO_DATA_IN_DAT_FILE  ' bad return
'         Exit Function
'      End If
'   Else
'      ProcessDATFile = ERROR_DAT_FILE_NOT_FOUND  ' bad return
'      Exit Function
'   End If
'   ' we have a collection of RSOjbects
'   ProcessDATFile = ERROR_SUCCESS ' good return
'   Exit Function
'ProcessDATFileErrTrap:
'   Dim ErrMsg As String
'   If Err.Number = 53 Then ' file does not exist
'      Err.Clear
'      Resume Next
'   End If
'   If gobjShared Is Nothing Then
'      Set gobjShared = New CWShared
'   End If
'   gobjShared.ShowVBError Err.Number, Err.Description, "ProcessDATFile Function", "DMBE", vbExclamation
'   ProcessDATFile = ERROR_DAT_FILE_PROCESSING  ' bad return
'   Resume
'End Function

'<SCR 14016 CC>

'Private Function ProcessDATFile(ByRef cRSObjs As Collection, ByVal strDataSet As String, ByVal FilePath As String) As Long
'
'   Const PersistedFileType As String = ".DAT"
'
'   Dim strFilePattern As String, strPath As String, strFileName As String
'   Dim lngSearchHandle As Long
'   Dim strTempPath As String
'   Dim objFSO As FileSystemObject
'   Dim strDataInLine() As String
'   Dim strTableName As String
'   Dim strVersion As String
'   Dim strFieldNames As String
'   Dim strDataFields() As String
'   Dim strDataValues() As String
'   Dim intColonPos As Integer
'   Dim lngDelimitorPos As Long
'   Dim rsData As ADODB.Recordset
'   Dim intLineNo As Long
'   Dim intI As Long
'   Dim lngRetVal As Long
'   Dim objRSObject As I_RSObjects
'   Dim strField() As String
'   Dim strOldTableName As String
'   Dim strTemp As String
'   Dim blnFound As Boolean
'   Dim intCount As Integer
'
'   On Error GoTo ErrHandler
'
'   ' find the DAT file on the PIC device - should only be one per clinical area
'   intLineNo = 0
'
'   'Build the search pattern to retrieve the file
'   strFilePattern = FilePath & "\" & "*" & gobjPatient.UnitNumber & "*" & strDataSet & "*" & PersistedFileType
'   strFilePattern = Replace(strFilePattern, "\\", "\", 1, -1, vbTextCompare)
'   strFileName = ""
'
'   ' check if right PIC device is still there
'   lngRetVal = I_DMBE_ValidatePIC()
'   If lngRetVal <> ERROR_SUCCESS Then
'      ProcessDATFile = lngRetVal
'      Exit Function
'   End If
'
'   'Then go get the first file
'   lngRetVal = GetFirstFile(strFilePattern, lngSearchHandle, strFileName)
'
'   If lngRetVal = gclFOUND Then
'
'      strPath = FilePath & "\" & strFileName
'      strPath = Replace(strPath, "\\", "\", 1, -1, vbTextCompare)
'      'Temp Path to copy the file to
'      strTempPath = Environ("TEMP") & "\" & strDataSet & PersistedFileType
'      strTempPath = Replace(strTempPath, "\\", "\", 1, -1, vbTextCompare)
'      Set objFSO = New FileSystemObject
'      If objFSO.FileExists(strTempPath) Then
'         objFSO.DeleteFile (strTempPath)
'      End If
'      objFSO.CopyFile strPath, strTempPath, True
'      Set objFSO = Nothing
'
'      ' check for encryption and/or compression of DAT file, at the first character of the file name
'      strTemp = Mid(strFileName, 1, 1)
'      If strTemp = "E" Then
'         lngRetVal = DecryptDataFile(strTempPath)   ' decrypt the DAT file
'         If lngRetVal <> 0 Then
'            ProcessDATFile = lngRetVal
'            Exit Function
'         End If
'      End If
'      'Then check for compression
'      strTemp = Mid(strFileName, 2, 1)
'      If strTemp = "C" Then
'         lngRetVal = DecompressDataFile(strTempPath)   ' decompress the DAT file
'         If lngRetVal <> 0 Then
'            ProcessDATFile = lngRetVal
'            Exit Function
'         End If
'      End If
'
'      ' check if right PIC device is still there
'      lngRetVal = I_DMBE_ValidatePIC()
'      If lngRetVal <> ERROR_SUCCESS Then
'         ProcessDATFile = lngRetVal
'         Exit Function
'      End If
'
'      ' read DAT file
'      strTemp = ReadFileContents(strTempPath)
'      If strTemp <> "" Then
'         ' get the table lines of the DAT file
'         strDataInLine = Split(strTemp, DAT_LINE_BREAK_CHARACTERS, -1, vbTextCompare) '<SCR 24045, 24055 CC>
'         ' get first recordset
'
'         For intLineNo = 0 To UBound(strDataInLine) - 1
'
'            ' get the values from a table line
'            strDataValues = Split(strDataInLine(intLineNo), DELIMITOR_CONSTANT, -1, vbTextCompare)
'            intColonPos = InStr(1, strDataValues(0), ":", vbTextCompare)
'
'            ' get the table name
'            strOldTableName = strTableName
'            strTableName = Mid(strDataValues(0), 1, intColonPos - 1)
'            ' get the field mapping table strVersion
'            lngDelimitorPos = InStr(1, strDataInLine(intLineNo), DELIMITOR_CONSTANT, vbTextCompare)
'            strVersion = Mid(strDataInLine(intLineNo), intColonPos + 1, lngDelimitorPos - intColonPos - 1)
'
'            If strOldTableName <> strTableName Then
'               Set rsData = New ADODB.Recordset
'               'If new table in Data then need to create new recordset
'               strFieldNames = GetFieldNames(strTableName, strVersion)
'               If strFieldNames = "" Then
'                  ProcessDATFile = ERROR_FIELD_MAPPING_TABLE
'                  Exit Function
'               End If
'               strDataFields = Split(strFieldNames, DELIMITOR_CONSTANT, -1, vbTextCompare)
'
'               For intI = 0 To UBound(strDataFields) - 1
'                  strField = Split(strDataFields(intI), ",", -1, vbTextCompare)
'                  strDataFields(intI) = strField(0)
'               Next intI
'
'               ' build recordset
'               lngRetVal = CreateDisconRS(strFieldNames, rsData)
'
'               If lngRetVal <> ERROR_SUCCESS Then
'                  ProcessDATFile = lngRetVal  ' bad return
'                  Exit Function
'               End If
'
'            End If
'
'            ' put Data into the recordset
'            lngRetVal = PopulateRS(rsData, strDataValues(), strDataFields())
'            If lngRetVal <> ERROR_SUCCESS Then
'               ProcessDATFile = ERROR_INVALID_RECORDSET_OPERATION
'               Exit Function
'            End If
'
'            blnFound = False
'            strTemp = strTableName
'            If strTemp = "IMMUNIZATIONS" Then
'                strTemp = "IMMUNIZATION"
'            End If
'
'            For intI = 1 To cRSObjs.Count
'                If cRSObjs(intI).TableName = strTemp Then
'                    blnFound = True
'                    Exit For
'                End If
'            Next
'
'            If blnFound = True Then 'The recordset is already in the collection
'                rsData.MoveLast
'                cRSObjs(intI).DBRECORDSET.AddNew
'                For intCount = 0 To rsData.Fields.Count - 1
'                    cRSObjs(intI).DBRECORDSET.Fields(intCount).Value = rsData.Fields(intCount).Value
'                Next intCount
'                cRSObjs(intI).DBRECORDSET.Update
'
'            Else
'                Set objRSObject = New RSObject_Ops
'                Set objRSObject.DBRECORDSET = rsData
'                objRSObject.TableName = strTemp
'                cRSObjs.Add objRSObject
'            End If
'
'         Next intLineNo
'      Else
'         ProcessDATFile = ERROR_NO_DATA_IN_DAT_FILE  ' bad return
'         Exit Function
'      End If
'   Else
'      ProcessDATFile = ERROR_DAT_FILE_NOT_FOUND  ' bad return
'      Exit Function
'   End If
'
'   ' we have a collection of RSOjbects
'   ProcessDATFile = ERROR_SUCCESS ' good return
'
'Exit Function
'
'ErrHandler:
'   Dim ErrMsg As String
'   If Err.Number = 53 Then ' file does not exist
'      Err.Clear
'      Resume Next
'   End If
'   If gobjShared Is Nothing Then
'      Set gobjShared = New CWShared
'   End If
'   gobjShared.ShowVBError Err.Number, Err.Description, "ProcessDATFile Function", "DMBE", vbExclamation
'   ProcessDATFile = ERROR_DAT_FILE_PROCESSING  ' bad return
'End Function
''<SCR 14016 CC>

'<SCR 24045 CC>
'Private Function ProcessDATFile(ByRef cRSObjs As Collection, ByVal strDataSet As String, ByVal FilePath As String) As Long
'Private Function ProcessDATFile(ByRef cRSObjs As Collection, ByVal strDataSet As String, ByVal strFilePath As String) As Long
'
'   Const PersistedFileType As String = ".DAT"
'
'   Dim strFilePattern As String, strPath As String, strFileName As String
'   Dim lngSearchHandle As Long
'   Dim strTempPath As String
'   Dim objFSO As FileSystemObject
'   Dim strDataInLine() As String
'   Dim strTableName As String
'   Dim strVersion As String
'   Dim strFieldNames As String
'   Dim strDataFields() As String
'   Dim strDataValues() As String
'   Dim intColonPos As Integer
'   Dim lngDelimitorPos As Long
'   Dim rsData As ADODB.Recordset
'   Dim intLineNo As Long
'   Dim intI As Long
'   Dim lngRetVal As Long
'   Dim objRSObject As I_RSObjects
'   Dim strField() As String
'   Dim strOldTableName As String
'   Dim strTemp As String
'   Dim blnFound As Boolean
'   Dim intCount As Integer
'   Dim blnIsClinicalNote As Boolean
'
'   On Error GoTo ErrHandler
'
'   ' find the DAT file on the PIC device - should only be one per clinical area
'   intLineNo = 0
'
'   'Build the search pattern to retrieve the file
'   'strFilePattern = strFilePath & "\" & "*" & gobjPatient.UnitNumber & "*" & strDataSet & "*" & PersistedFileType
'   strFilePattern = strFilePath & "\" & "*" & m_lngPICPatientUnitNumber & "*" & strDataSet & "*" & PersistedFileType
'   strFilePattern = Replace(strFilePattern, "\\", "\", 1, -1, vbTextCompare)
'   strFileName = ""
'
'   ' check if right PIC device is still there
'   lngRetVal = I_DMBE_ValidatePIC()
'   If lngRetVal <> ERROR_SUCCESS Then
'      ProcessDATFile = lngRetVal
'      Exit Function
'   End If
'
'   'Then go get the first file
'   lngRetVal = GetFirstFile(strFilePattern, lngSearchHandle, strFileName)
'
'   If lngRetVal = gclFOUND Then
'
'      strPath = strFilePath & "\" & strFileName
'      strPath = Replace(strPath, "\\", "\", 1, -1, vbTextCompare)
'      'Temp Path to copy the file to
'      strTempPath = Environ("TEMP") & "\" & strDataSet & PersistedFileType
'      strTempPath = Replace(strTempPath, "\\", "\", 1, -1, vbTextCompare)
'      Set objFSO = New FileSystemObject
'      If objFSO.FileExists(strTempPath) Then
'         objFSO.DeleteFile (strTempPath)
'      End If
'      objFSO.CopyFile strPath, strTempPath, True
'      Set objFSO = Nothing
'
'      ' check for encryption and/or compression of DAT file, at the first character of the file name
'      strTemp = Mid(strFileName, 1, 1)
'      If strTemp = "E" Then
'         lngRetVal = DecryptDataFile(strTempPath)   ' decrypt the DAT file
'         If lngRetVal <> 0 Then
'            ProcessDATFile = lngRetVal
'            Exit Function
'         End If
'      End If
'      'Then check for compression
'      strTemp = Mid(strFileName, 2, 1)
'      If strTemp = "C" Then
'         lngRetVal = DecompressDataFile(strTempPath)   ' decompress the DAT file
'         If lngRetVal <> 0 Then
'            ProcessDATFile = lngRetVal
'            Exit Function
'         End If
'      End If
'
'      ' check if right PIC device is still there
'      lngRetVal = I_DMBE_ValidatePIC()
'      If lngRetVal <> ERROR_SUCCESS Then
'         ProcessDATFile = lngRetVal
'         Exit Function
'      End If
'
'      ' read DAT file
'      strTemp = ReadFileContents(strTempPath)
'      blnIsClinicalNote = False
'
'      If strTemp <> "" Then
'
'           ' get the table lines of the DAT file
'           '<SCR 14016 CC>
'         strDataInLine = Split(strTemp, DAT_LINE_BREAK_CHARACTERS, -1, vbTextCompare)
'
'         For intLineNo = 0 To UBound(strDataInLine) - 1
'
'            If blnIsClinicalNote = True Then
'                'If intLineNo = UBound(strDataInLine) - 1 Then
'                '    Exit For
'                'End If
'            End If
'
'            ' get the values from a table line
'            'DELIMITOR_CHARACTERS
'            'strDataValues = Split(strDataInLine(intLineNo), DELIMITOR_CONSTANT, -1, vbTextCompare)
'            strDataValues = Split(strDataInLine(intLineNo), DELIMITOR_CHARACTERS, -1, vbTextCompare)
'
'            intColonPos = InStr(1, strDataValues(0), ":", vbTextCompare)
'
'            ' get the table name
'            strOldTableName = strTableName
'
'            strTableName = Mid(strDataValues(0), 1, intColonPos - 1)
'            ' get the field mapping table strVersion
'            'lngDelimitorPos = InStr(1, strDataInLine(intLineNo), DELIMITOR_CONSTANT, vbTextCompare)
'            lngDelimitorPos = InStr(1, strDataInLine(intLineNo), DELIMITOR_CHARACTERS, vbTextCompare)
'            strVersion = Mid(strDataInLine(intLineNo), intColonPos + 1, lngDelimitorPos - intColonPos - 1)
'
'            If strOldTableName <> strTableName Then
'               Set rsData = New ADODB.Recordset
'               'If new table in Data then need to create new recordset
'               strFieldNames = GetFieldNames(strTableName, strVersion)
'               If strFieldNames = "" Then
'                  ProcessDATFile = ERROR_FIELD_MAPPING_TABLE
'                  Exit Function
'               End If
'               strDataFields = Split(strFieldNames, DELIMITOR_CONSTANT, -1, vbTextCompare)
'
'               For intI = 0 To UBound(strDataFields) - 1
'                  strField = Split(strDataFields(intI), ",", -1, vbTextCompare)
'                  strDataFields(intI) = strField(0)
'               Next intI
'
'               ' build recordset
'               lngRetVal = CreateDisconRS(strFieldNames, rsData)
'
'               If lngRetVal <> ERROR_SUCCESS Then
'                  ProcessDATFile = lngRetVal  ' bad return
'                  Exit Function
'               End If
'
'            End If
'
'            ' put Data into the recordset
'            lngRetVal = PopulateRS(rsData, strDataValues(), strDataFields())
'            If lngRetVal <> ERROR_SUCCESS Then
'               ProcessDATFile = ERROR_INVALID_RECORDSET_OPERATION
'               Exit Function
'            End If
'
'            blnFound = False
'            strTemp = strTableName
'            If strTemp = "IMMUNIZATIONS" Then
'                strTemp = "IMMUNIZATION"
'            End If
'
'            For intI = 1 To cRSObjs.Count
'                If cRSObjs(intI).TableName = strTemp Then
'                    blnFound = True
'                    Exit For
'                End If
'            Next
'
'            If blnFound = True Then 'The recordset is already in the collection
'                rsData.MoveFirst
'                Set objRSObject.DBRECORDSET = rsData
'
'            Else
'                Set objRSObject = New RSObject_Ops
'                Set objRSObject.DBRECORDSET = rsData
'                objRSObject.TableName = strTemp
'                cRSObjs.Add objRSObject, strTemp
'            End If
'
'         Next intLineNo
'
'      Else
'         ProcessDATFile = ERROR_NO_DATA_IN_DAT_FILE  ' bad return
'         Exit Function
'      End If
'   Else
'      ProcessDATFile = ERROR_DAT_FILE_NOT_FOUND  ' bad return
'      Exit Function
'   End If
'
'   ' we have a collection of RSOjbects
'   ProcessDATFile = ERROR_SUCCESS ' good return
'
'Exit Function
'
'ErrHandler:
'   Dim ErrMsg As String
'   If Err.Number = 53 Then ' file does not exist
'      Err.Clear
'      Resume Next
'   End If
'   If gobjShared Is Nothing Then
'      Set gobjShared = New CWShared
'   End If
'   gobjShared.ShowVBError Err.Number, Err.Description, "ProcessDATFile Function", "DMBE", vbExclamation
'   ProcessDATFile = ERROR_DAT_FILE_PROCESSING  ' bad return
'
'End Function


'This will Create a *.PIC File
Private Function AppendPICInfo(dicDatasets As Scripting.Dictionary) As Long
Dim oDom As MSXML2.DOMDocument30
Dim oNode As MSXML2.IXMLDOMElement
Dim oNode1 As MSXML2.IXMLDOMElement
Dim sData As String
Dim sValArr() As String
Dim idx As Long
Dim bCompress As Boolean
Dim bEncrypt As Boolean
Dim nError As Long
On Error GoTo ErrHandler
    
    'Check for Compressed
    sValArr = Split(m_sPicFile, "_")
    If UBound(sValArr) > 0 Then
        If Left(sValArr(1), 1) = "C" Then
           bCompress = True
        End If
        
        'Check for Encrypt
        If Mid(sValArr(1), 2, 1) = "E" Then
           bEncrypt = True
        End If
    End If
    
    
    'm_strAltPicPath is the PIC File in the user's Temp Folder
    sData = ReadFileContents(m_strAltPicPath & m_sPicFile, bCompress, bEncrypt, nError)
    
    If nError <> 0 Then
        AppendPICInfo = nError
        GoTo Function_Exit
    End If
    
    Set oDom = New MSXML2.DOMDocument30
    If Not oDom.loadXML(sData) Then
        AppendPICInfo = ERROR_INVALID_FORMAT
        GoTo Function_Exit
    End If
    
    Set oNode = oDom.selectSingleNode("//PICINFO/DATASETS")
    If oNode Is Nothing Then
        Set oNode = oDom.selectSingleNode("PICINFO")
        If Not oNode Is Nothing Then
            Set oNode1 = oDom.createElement("DATASETS")
            oNode.appendChild oNode1
        Else
            AppendPICInfo = ERROR_INVALID_FORMAT
            GoTo Function_Exit
        End If
    End If
    
    For idx = 0 To dicDatasets.Count - 1
        Set oNode = oDom.selectSingleNode("//PICINFO/DATASETS/" & UCase$(dicDatasets.Keys(idx)))
        If Not oNode Is Nothing Then
            oNode.Text = dicDatasets.Items(idx)
        Else
            Set oNode = oDom.selectSingleNode("//PICINFO/DATASETS")
            If Not oNode Is Nothing Then
                Set oNode1 = oDom.createElement(UCase$(dicDatasets.Keys(idx)))
                oNode1.Text = dicDatasets.Items(idx)
                oNode.appendChild oNode1
            Else
                AppendPICInfo = ERROR_INVALID_FORMAT
                GoTo Function_Exit
            End If
            
        End If
        'oDOM.appendChild oNode
    Next idx
    
    sData = oDom.xml
    
    AppendPICInfo = WriteFileContents(m_strAltPicPath & m_sPicFile, sData, bCompress, bEncrypt)
Function_Exit:
    Set oDom = Nothing
    Set oNode = Nothing
Exit Function

ErrHandler:
   If gobjShared Is Nothing Then
      Set gobjShared = New CWShared
   End If
   gobjShared.ShowVBError Err.Number, Err.Description, "AppendPICInfo Function", "DMBE", vbExclamation
   Err.Clear
   GoTo Function_Exit
End Function

Private Function WritePICInfo(sPathFile As String, bCompress As Boolean, bEncrypt As Boolean) As Long
   Dim oDom As MSXML2.DOMDocument30
   Dim oRoot As MSXML2.IXMLDOMElement
   Dim oNode0 As MSXML2.IXMLDOMElement
   Dim oNode1 As MSXML2.IXMLDOMElement
   Dim sData As String
   
On Error GoTo ErrHandler

    bCompress = False
    bEncrypt = True
    
    Set oDom = New MSXML2.DOMDocument30
    
    Set oRoot = oDom.createElement("PICINFO")
     
    'Patient Demographics
    Set oNode0 = oDom.createElement("PATIENT")
    Set oNode1 = oDom.createNode(NODE_ELEMENT, "NAME", "")
    oNode1.Text = Trim(gobjPatient.Name)
    oNode0.appendChild oNode1
        
    Set oNode1 = oDom.createNode(NODE_ELEMENT, "BIRTHDATE", "")
    oNode1.Text = Trim(Format(gobjPatient.DateOfBirth, "dd mmm yyyy"))
    oNode0.appendChild oNode1
    
    Set oNode1 = oDom.createNode(NODE_ELEMENT, "SSN", "")
    oNode1.Text = Trim(gobjPatient.SSN)
    oNode0.appendChild oNode1

    oRoot.appendChild oNode0
    
    Set oNode0 = oDom.createNode(NODE_ELEMENT, "COMPRESS", "")
    oNode0.Text = m_enmCompression
    oRoot.appendChild oNode0
    
    Set oNode0 = oDom.createNode(NODE_ELEMENT, "FORMAT", "")
    oNode0.Text = m_enmDMFormat
    oRoot.appendChild oNode0
    
    Set oNode0 = oDom.createNode(NODE_ELEMENT, "ENCRYPT", "")
    oNode0.Text = m_enmCrypto
    oRoot.appendChild oNode0
    
    sData = "<?xml version=""1.0""?>" & oRoot.xml
    WritePICInfo = WriteFileContents(sPathFile, sData, False, True)

    Exit Function
ErrHandler:
   If gobjShared Is Nothing Then
      Set gobjShared = New CWShared
   End If
   gobjShared.ShowVBError Err.Number, Err.Description, "WritePICInfo Function", "DMBE", vbExclamation
   Err.Clear
End Function

'This will create both a *.PIC and *.XML
Private Function I_DMBE_WriteToPIC(ByVal strSSN As String, oDataset As Object) As Long
Dim oDom As MSXML2.DOMDocument30
Dim oFS As Scripting.FileSystemObject
Dim sData As String
Dim sFormat As String
Dim bCompress As Boolean
Dim bEncrypt As Boolean
Dim aArray() As String
Dim dicDatasets As Scripting.Dictionary
Dim nReturn As Long
Dim sPVFCopyLocation As String
Dim sTempPath As String
Dim bNeedCopy As Boolean

    On Error GoTo ErrHandler
    
    Set dicDatasets = oDataset
    If dicDatasets.Count > 0 Then
        'This will Create ther *.PIC File
        Call AppendPICInfo(dicDatasets)
    End If
    
    Select Case m_enmDMFormat
    Case eDM_ADT
        sFormat = ".ADT"
    Case eDM_XML
        sFormat = ".XML"
    Case eDM_DAT
        sFormat = ".DAT"
    Case Else
          I_DMBE_WriteToPIC = ERROR_INVALID_FORMAT ' bad return - data format not initialized
          Exit Function
    End Select

    If m_strAltPicPath <> m_strPicPath Then
      'Copy PIC file from temporary location to TMIP location
      sTempPath = Environ$("TEMP") & "\"
      bNeedCopy = True
    Else
      sTempPath = m_strPicPath
    End If
    
    aArray = Split(m_sPicFile, "_")
    If UBound(aArray) = 0 Then
        I_DMBE_WriteToPIC = ERROR_INVALID_FORMAT
        Exit Function
    End If
    
    If Left(aArray(1), 1) = "C" Then
        bCompress = True
    End If
    
    If Mid(aArray(1), 2, 1) = "E" Then
        bEncrypt = True
    End If
    
    'Create XML File in Temp folder
    Set oFS = New FileSystemObject
    Set oDom = New MSXML2.DOMDocument30
    nReturn = CreatePICRSCol2XML(m_colRSObjects, m_enmDMFormat, sData, oDom)
        
    If nReturn = ERROR_SUCCESS Then
      If bCompress Or bEncrypt Then
        Do
          I_DMBE_WriteToPIC = ERROR_SUCCESS
          I_DMBE_WriteToPIC = WritePICData(sTempPath & oFS.GetBaseName(m_sPicFile) & sFormat, sData, bCompress, bEncrypt)
        Loop While I_DMBE_WriteToPIC = ERROR_RETRY
      Else
        oDom.Save sTempPath & oFS.GetBaseName(m_sPicFile) & sFormat
      End If
    End If
    
    If bNeedCopy Then
    ' Begin:SCR #38308;   Developer: Mag 06/30/2003 05:45 PM
'    If m_lEncSource = enEncounterSigned Then
      'This will copy the Temporary PVF file to the location where the framework expects the file
      sPVFCopyLocation = m_strPicPath
      m_strPicPath = m_strAltPicPath       'The PVF went to .Destination(m_strAltPicPath) Path set it to be .Source(m_strPicPath) to copy
      m_strAltPicPath = sPVFCopyLocation
      'This calls I_DMBE_CopyPIC,
      '  which will copy *.PIC and *.XML from user's Temp folder
      '  To PVF_OUT
      I_DMBE_WriteToPIC = I_DMBE_CopyPIC
      If unitNumber <> "" Then
        ' Begin:SCR #44865;   Developer:  11/07/2003
        On Error Resume Next
        Call oFS.DeleteFile(sTempPath & unitNumber & "*.*", True)
        If Err.Number <> 0 Then
          If m_sPicFile <> "" Then
            Err.Clear
            Call oFS.DeleteFile(sTempPath & oFS.GetBaseName(m_sPicFile) & ".*", True)
          End If
        End If
        On Error GoTo ErrHandler
        ' End: SCR #44865;
      End If
      'Reset path back to original values
      sPVFCopyLocation = m_strPicPath
      m_strPicPath = m_strAltPicPath
      m_strAltPicPath = sPVFCopyLocation
      
    End If
'    End If
    ' End: SCR #38308;
    
Cleanup:
  Set oDom = Nothing
  Set oFS = Nothing
Exit Function

ErrHandler:
   If gobjShared Is Nothing Then
      Set gobjShared = New CWShared
   End If
   gobjShared.ShowVBError Err.Number, Err.Description, "WriteToPic Function", "DMBE", vbExclamation
   Err.Clear
   Resume Cleanup
End Function

Private Function CreatePICRSCol2XML(oColRS As Collection, enmFormat As DMFormats, sReturnData As String, oDom As MSXML2.DOMDocument30) As Long
    Dim oRoot As MSXML2.IXMLDOMElement
    Dim oNode As MSXML2.IXMLDOMElement
    Dim oDMRS As IDMBE.I_RSObjects
    
    On Error GoTo ErrHandler
    
    If oColRS.Count = 0 Then
      CreatePICRSCol2XML = ERROR_NO_RECORDS_FOR_DATASET_FOUND
      Exit Function
    End If
  
    Set oRoot = oDom.createElement("RECORDSETS")
    For Each oDMRS In oColRS
        If oDMRS.DBRECORDSET.RecordCount <> 0 Or UCase$(oDMRS.TableName) = "ENCOUNTERS" Then
            Set oNode = oRoot.selectSingleNode(UCase$(oDMRS.TableName))
            Debug.Print oDMRS.TableName
            If oNode Is Nothing Then
              Set oNode = StreamRS2Node(oDMRS.DBRECORDSET, UCase$(oDMRS.TableName), enmFormat)
              oRoot.appendChild oNode
            End If
        End If
    Next oDMRS
    
    sReturnData = "<?xml version=""1.0""?>" & oRoot.xml
    
    oDom.appendChild oRoot
    Exit Function
ErrHandler:
   If gobjShared Is Nothing Then
      Set gobjShared = New CWShared
   End If
   gobjShared.ShowVBError Err.Number, Err.Description, "CreatePICRSCol2XML Function", "DMBE", vbExclamation

End Function

Private Function StreamRS2Node(ByRef rstIn As ADODB.Recordset, sSegName As String, enmFormat As DMFormats) As Object
  Dim oDom As MSXML2.DOMDocument30
  Dim oTmpDOM As MSXML2.DOMDocument30
  Dim oNode As MSXML2.IXMLDOMElement
  Dim oStream As ADODB.Stream
  Dim sTemp As String
  Dim sDoc As String
  Dim nCol As Long
  On Error GoTo Sub_Err

  Set oDom = New MSXML2.DOMDocument30
  Set oTmpDOM = New MSXML2.DOMDocument30
  Set oStream = New ADODB.Stream
  
  'This loads the rst into a stream
  If enmFormat = eDM_ADT Then
    rstIn.Save oStream, adPersistADTG
    sDoc = oStream.Read(-1)
    sTemp = EncodeString(sDoc, True)
    Set StreamRS2Node = oDom.createElement(sSegName)
    StreamRS2Node.Text = sTemp
      
    Select Case UCase$(sSegName)
    Case "ENC_RTFS", "ENC_SECTIONS", "LISTTOOLRECORD"
        If UCase$(sSegName) = "LISTTOOLRECORD" Then
           sDoc = "RTF"
        Else
           sDoc = "DOC"
        End If
        If rstIn.RecordCount > 0 Then
          rstIn.MoveFirst
          Do While Not rstIn.EOF
            Set oNode = oTmpDOM.createElement(sDoc)
            sTemp = rstIn(sDoc).Value & vbNullString
            If rstIn("OriginalSize").Value > 1 Then
              oNode.Text = EncodeString(sTemp, True)
            Else
              oNode.Text = FilterRtfToXml(sTemp)
            End If
            StreamRS2Node.appendChild oNode
            rstIn.MoveNext
          Loop
        End If
    End Select
  
  ElseIf enmFormat = eDM_XML Then
    rstIn.Save oStream, adPersistXML
    If oTmpDOM.loadXML(oStream.ReadText(-1)) Then
      Set StreamRS2Node = oDom.createElement(sSegName)
      StreamRS2Node.appendChild oTmpDOM.childNodes(0)
      
      Select Case UCase$(sSegName)
      Case "ENC_RTFS", "ENC_SECTIONS", "LISTTOOLRECORD"
        If UCase$(sSegName) = "LISTTOOLRECORD" Then
           sDoc = "RTF"
        Else
           sDoc = "DOC"
        End If
        If rstIn.RecordCount > 0 Then
          rstIn.MoveFirst
          Do While Not rstIn.EOF
            Set oNode = oTmpDOM.createElement(sDoc)
            sTemp = rstIn(sDoc).Value & vbNullString
            If rstIn("OriginalSize").Value > 1 Then
              oNode.Text = EncodeString(sTemp, True)
            Else
              oNode.Text = FilterRtfToXml(sTemp)
            End If
            StreamRS2Node.appendChild oNode
            rstIn.MoveNext
          Loop
        End If
      End Select
    
    End If
  ElseIf enmFormat = eDM_DAT Then
    Set StreamRS2Node = oDom.createElement(sSegName)
    
    sTemp = ""
    If Not (rstIn.BOF And rstIn.EOF) Then
        For nCol = 0 To rstIn.Fields.Count - 1
            sTemp = sTemp & rstIn.Fields(nCol).Name & DELIMITOR_CHARACTERS
        Next nCol
        sTemp = sTemp & vbCrLf
        
        rstIn.MoveFirst
        Do While Not rstIn.EOF
            For nCol = 0 To rstIn.Fields.Count - 1
              If IsNull(rstIn.Fields(nCol).Value) Then
                sTemp = sTemp & "NULL" & DELIMITOR_CHARACTERS
              ElseIf rstIn.Fields(nCol).Type = adLongVarBinary Then
                Select Case UCase$(sSegName)
                Case "ENC_RTFS", "ENC_SECTIONS", "LISTTOOLRECORD"
                  If rstIn("OriginalSize").Value > 1 Then
                    sTemp = sTemp & EncodeString(rstIn.Fields(nCol).Value) & DELIMITOR_CHARACTERS
                  Else
                    sTemp = sTemp & FilterRtfToXml(rstIn.Fields(nCol).Value) & DELIMITOR_CHARACTERS
                  End If
                Case Else
                    sTemp = sTemp & FilterRtfToXml(rstIn.Fields(nCol).Value & vbNullString) & DELIMITOR_CHARACTERS
                End Select
              Else
                sTemp = sTemp & rstIn.Fields(nCol).Value & DELIMITOR_CHARACTERS
              End If
            Next nCol
            sTemp = sTemp & vbCrLf
            rstIn.MoveNext
        Loop
    End If
    StreamRS2Node.Text = FilterRtfToXml(sTemp)
  Else
    Set StreamRS2Node = Nothing
  End If
  
  
Sub_Exit:
  On Error Resume Next
  Set oDom = Nothing
  Set oTmpDOM = Nothing
  Set oStream = Nothing
  Exit Function
  
Sub_Err:
  gobjShared.ShowVBError Err.Number, Err.Description, "StreamRS2Node", "DMBE", vbExclamation
  Resume Sub_Exit
  Resume 'Debug
End Function

Private Function WritePICData(sPathFile As String, sData As String, ByVal bCompress As Boolean, ByVal bEncrypt As Boolean) As Long
    Dim oFS As Scripting.FileSystemObject
    Dim oFile As Scripting.File
    Dim dblBytesFree As Double
    Dim lngRetVal As Long
    Dim lngFileSize As Long
    Dim sMsg As String
    Dim bForceBinary As Boolean
    Dim abTemp() As Byte
    
   On Error GoTo WritePICErrTrap
   
    bCompress = False
    bEncrypt = True
   
     dblBytesFree = DiskSpaceCheck(sPathFile)
                    
     Set oFS = New Scripting.FileSystemObject
     If oFS.FileExists(sPathFile) Then
       Set oFile = oFS.GetFile(sPathFile)
       If Not oFile Is Nothing Then
         dblBytesFree = dblBytesFree + oFile.Size
       End If
       Set oFile = Nothing
     End If
     Set oFS = Nothing
     
     If dblBytesFree < 0 Then
        'Need new error
        WritePICData = dblBytesFree
        GoTo Cleanup
     End If
                  
     'Compress data if selected for disk space validation
     If bCompress Then
         ' compress the data
         sData = CompressString(sData)
         If sData = "" Then
            WritePICData = ERROR_WITH_COMPRESSION_FUNCTION ' bad return
            GoTo Cleanup
         End If
         'Since data is compressed do not compress again on WriteFileContents
         bCompress = False
         bForceBinary = True
     End If

'     If bEncrypt Then
'         ' encrypt the file
'         Set cipher = New CCrypto
'         sData = cipher.Encrypt(sData, mc_strPassword, False, frezBlockEncryption)
'         If sData = "" Then
'            WritePICData = ERROR_WITH_ENCRYPTION_FUNCTION ' bad return
'            Exit Function
'         End If
'         'Since data is Encrypted do not Encrypt again on WriteFileContents
'         bEncrypt = False
'         bForceBinary = True
'     End If
     
     ' Begin:SCR #14066;   Developer: Mag 03/05/2003 03:21 PM
     abTemp = sData
     lngFileSize = UBound(abTemp) + 1
     ' End: SCR #14066;
                  
     If dblBytesFree >= lngFileSize Then
        ' check if right PIC device is still there
        lngRetVal = I_DMBE_ValidatePIC()
        If lngRetVal <> ERROR_SUCCESS Then
           WritePICData = lngRetVal
           GoTo Cleanup
        End If
         ' write the file contents
        Call WriteFileContents(sPathFile, sData, bCompress, bEncrypt, bForceBinary)
     Else
        sMsg = "No records will be written. The file size exceeds the capacity of the destination device."
        sMsg = sMsg & vbCrLf & vbCrLf & "Press cancel to return to the Data Manager Transaction screen"
        If MsgBox(sMsg, vbRetryCancel) = vbRetry Then
          WritePICData = ERROR_RETRY
        Else
          WritePICData = ERROR_INSUFFICIENT_SPACE_ON_DEVICE ' bad return - insufficient space on output device
        End If
        GoTo Cleanup
     End If

     WritePICData = ERROR_SUCCESS
    
Cleanup:
  Set oFS = Nothing
  Set oFile = Nothing
Exit Function


WritePICErrTrap:
   If Err.Number = 53 Then ' file does not exist
      Err.Clear
      Resume Next
   End If
   If gobjShared Is Nothing Then
      Set gobjShared = New CWShared
   End If
   gobjShared.ShowVBError Err.Number, Err.Description, "WritePICData Function", "DMBE", vbExclamation
   WritePICData = ERROR_INVALID_COLLECTION  ' bad return
   Resume Cleanup
   'Resume
End Function

'Private Function WritePIC(colRSObjects As Collection, DFormat As DMFormats, FilePath As String) As Long
'   Dim strPersistedFileType As String
'   Dim rsData As ADODB.Recordset
'   Dim objRSObject As IDMBE.I_RSObjects
'   Dim objDeletedRSObject As IDMBE.I_RSObjects
'   Dim lngRetVal As Long
'   Dim strPath As String
'   Dim strTempPath As String
'   Dim objFSO As FileSystemObject
'   Dim lngFileSize As Long
'   Dim dblBytesFree As Double
'   Dim strFieldNames As String
'   Dim strFieldName() As String
'   Dim lngUpperLimit As Long
'   Dim strCat As String
'   Dim strVersion As String
'   Dim strClinicalTableName As String
'   Dim lngJ As Long
'   Dim lngRecordCount As Long
'   Dim strTableName As String
'   Dim lngUIndex As Long
'   Dim strTemp() As String
'   Dim lngRSCount As Long
'
'   On Error GoTo WritePICErrTrap
'
'   If colRSObjects.Count = 0 Then
'      WritePIC = ERROR_NO_RECORDS_FOR_DATASET_FOUND
'      Exit Function
'   End If
'
'   'Here - translate provider NCID to SSN
'
'
'
'   ' check for persisted recordset format
'   Select Case DFormat
'      Case eDM_ADT
'            strPersistedFileType = ".ADT"
'      Case eDM_XML
'            strPersistedFileType = ".XML"
'      Case eDM_DAT
'            strPersistedFileType = ".DAT"
'      Case Else
'            WritePIC = ERROR_INVALID_FORMAT ' bad return - data format not initialized
'            Exit Function
'   End Select
'
'   ' get recordsets from collection for deleting from the PIC
'   For Each objRSObject In colRSObjects
'      If Not (objRSObject Is Nothing) Then
'         lngRetVal = DeletePIC(m_strPicPath, objRSObject.TableName, ".*")
'         If lngRetVal <> Error_Success Then
'            WritePIC = lngRetVal
'            Exit Function
'         End If
'      Else
'         WritePIC = ERROR_COLLECTION_RSOBJECT_IS_NOTHING ' bad return
'         Exit Function
'      End If
'   Next objRSObject
'
'
'   ' get recordsets from collection for adding to the PIC
'   strCat = ""
'   lngRSCount = 0
'
'   For Each objRSObject In colRSObjects
'      Set rsData = objRSObject.DBRECORDSET
'      If rsData.RecordCount = 0 Then
'         Set rsData = Nothing
'      Else
'         lngRSCount = lngRSCount + rsData.RecordCount
'      End If
'
'      If Not (rsData Is Nothing) Then
'
'         If rsData.Fields.Count <> 0 Then
'
'            strPath = FormatPath(m_strPicPath, objRSObject.TableName, strPersistedFileType)
'            ' get rid of records of this type on the PIC
'            strPath = EnsureUniqueFile(strPath)
'            Select Case UCase(strPersistedFileType)
'               Case ".XML", ".ADT"
'                     strTempPath = Environ("TEMP") & "\" & objRSObject.TableName & strPersistedFileType
'                     strTempPath = Replace(strTempPath, "\\", "\", 1, -1, vbTextCompare)
'                     Set objFSO = New FileSystemObject
'
'                     If objFSO.FileExists(strTempPath) Then
'                        objFSO.DeleteFile strTempPath, True ' just in case of failure last time in clearing file
'                     End If
'                     Set objFSO = Nothing
'
'                     If UCase(strPersistedFileType) = ".XML" Then
'                        rsData.Save strTempPath, adPersistXML
'                     End If
'                     If UCase(strPersistedFileType) = ".ADT" Then
'                        rsData.Save strTempPath, adPersistADTG
'                     End If
'                     dblBytesFree = DiskSpaceCheck(strPath)
'
'                     If dblBytesFree < 0 Then
'                        'Need new error
'                        WritePIC = dblBytesFree
'                        Kill strTempPath
'                        ' delete all of this data type from output device - it is all or nothing here
'                        For Each objDeletedRSObject In colRSObjects
'                           lngRetVal = DeletePIC(m_strPicPath, objDeletedRSObject.TableName, strPersistedFileType)
'                           If lngRetVal <> Error_Success Then
'                              WritePIC = lngRetVal
'                              Exit Function
'                           End If
'                        Next objDeletedRSObject
'                        Exit Function
'                     End If
'
'                     If m_enmCompression = eDM_Compressed Then
'                        lngRetVal = CompressDataFile(strTempPath)
'                        If lngRetVal <> 0 Then
'                           WritePIC = lngRetVal
'                           Exit Function
'                        End If
'                     End If
'                     If m_enmCrypto = eDM_Encrypted Then
'                        lngRetVal = EncryptDataFile(strTempPath)
'                        If lngRetVal <> 0 Then
'                           WritePIC = lngRetVal
'                           Exit Function
'                        End If
'                     End If
'                     lngFileSize = FileLen(strTempPath)
'
'                     If dblBytesFree >= lngFileSize Then
'                        ' check if right PIC device is still there
'                        lngRetVal = I_DMBE_ValidatePIC()
'                        If lngRetVal <> Error_Success Then
'                           WritePIC = lngRetVal
'                           Exit Function
'                        End If
'                        'Here copy from strTempPath to strPath
'                        Set objFSO = New FileSystemObject
'                        objFSO.CopyFile strTempPath, strPath, True
'                        objFSO.DeleteFile strTempPath, True
'                        Set objFSO = Nothing
'                     Else
'                        ' check if right PIC device is still there
'                        lngRetVal = I_DMBE_ValidatePIC()
'                        If lngRetVal <> Error_Success Then
'                           WritePIC = lngRetVal
'                           Exit Function
'                        End If
'                        ' delete temp file
'                        Kill strTempPath
'                        ' delete all of this data type from output device - it is all or nothing here
'                        For Each objDeletedRSObject In colRSObjects
'                           lngRetVal = DeletePIC(m_strPicPath, objDeletedRSObject.TableName, strPersistedFileType)
'                           If lngRetVal <> Error_Success Then
'                              WritePIC = lngRetVal
'                              Exit Function
'                           End If
'                        Next objDeletedRSObject
'                        WritePIC = ERROR_INSUFFICIENT_SPACE_ON_DEVICE ' bad return - insufficient space on output device
'                        Exit Function
'                     End If
'               Case ".DAT"
'                     strClinicalTableName = GetClinicalDataAreaByEnum(m_enmDataSet)
'                     strTempPath = Environ("TEMP") & "\" & strClinicalTableName & strPersistedFileType
'                     strTempPath = Replace(strTempPath, "\\", "\", 1, -1, vbTextCompare)
'                     Set objFSO = New FileSystemObject
'                     If objFSO.FileExists(strTempPath) Then
'                        objFSO.DeleteFile strTempPath, True ' just in case of failure last time in clearing file
'                     End If
'                     Set objFSO = Nothing
'                     strVersion = ""
'                     ' take care of special case
'                     strTableName = objRSObject.TableName
'                     If UCase(strTableName) = "IMMUNIZATION" Then
'                        strTableName = "IMMUNIZATIONS"
'                     End If
'                     strFieldNames = GetFieldNames(strTableName, strVersion)
'
'                     If strFieldNames <> "" Then
'                        strFieldName = Split(strFieldNames, DELIMITOR_CONSTANT, -1, vbTextCompare)
'                        lngUIndex = UBound(strFieldName) - 1
'                        For lngJ = 0 To lngUIndex
'                           strTemp = Split(strFieldName(lngJ), ",", -1, vbTextCompare)
'                           strFieldName(lngJ) = strTemp(0)
'                        Next lngJ
'                        lngUpperLimit = UBound(strFieldName) - 1
'                        If rsData.EOF = False Then
'                           rsData.MoveFirst
'                           Do
'                              strCat = strCat & strTableName & ":" & strVersion & DELIMITOR_CHARACTERS   'DELIMITOR_CONSTANT
'                              ' go through the field names for the current database table version
'                              For lngJ = 0 To lngUpperLimit
'                                 strCat = strCat & rsData.Fields(strFieldName(lngJ)).Value & DELIMITOR_CHARACTERS    'DELIMITOR_CONSTANT
'                              Next lngJ
'                              '<SCR 14016 CC>'
'                              strCat = strCat & DAT_LINE_BREAK_CHARACTERS
'                              ' save file to temp location while we continue to aggregate recordsets
'                              Call WriteFileContents(strTempPath, strCat)
'                              rsData.MoveNext
'                           Loop While rsData.EOF = False
'                        End If
'                     Else
'                        WritePIC = ERROR_TABLE_COLUMN_NAMES_NOT_LOADED
'                        Exit Function
'                     End If
'            End Select
'            ' destroy recordset object
'            Set rsData = Nothing
'         Else 'rsData.Fields.Count <> 0
'            lngRecordCount = lngRecordCount + 1
'         End If
'      End If 'rsData Is Nothing
'   Next objRSObject
'
'   '<SCR 22715 CC>
'   If lngRSCount = 0 Then
'    WritePIC = DMErrorCodes.ERROR_NO_RECORDS_FOR_DATASET_FOUND
'    Exit Function
'   End If
'   '</SCR 22715 CC>
'
'   ' special processing of DAT file
'   If UCase(strPersistedFileType) = ".DAT" Then
'      If m_enmCompression = eDM_Compressed Then
'         lngRetVal = CompressDataFile(strTempPath)
'         If lngRetVal <> 0 Then
'            WritePIC = lngRetVal
'            Exit Function
'         End If
'      End If
'      If m_enmCrypto = eDM_Encrypted Then
'         lngRetVal = EncryptDataFile(strTempPath)
'         If lngRetVal <> 0 Then
'            WritePIC = lngRetVal
'            Exit Function
'         End If
'      End If
'
'      strClinicalTableName = GetClinicalDataAreaByEnum(m_enmDataSet)
'      If UCase(strTableName) = "IMMUNIZATION" Then
'            strTableName = "IMMUNIZATIONS"
'      End If
'
'      strPath = FormatPath(m_strPicPath, strClinicalTableName, strPersistedFileType)
'      lngFileSize = FileSystem.FileLen(strTempPath) '<SCR 24045 CC>   'FileLen(strTempPath)
'      dblBytesFree = DiskSpaceCheck(strPath)
'      If dblBytesFree < 0 Then
'         WritePIC = dblBytesFree
'         Kill strTempPath
'         ' delete all of this data type from output device - it is all or nothing here
'         For Each objDeletedRSObject In colRSObjects
'            lngRetVal = DeletePIC(m_strPicPath, objDeletedRSObject.TableName, strPersistedFileType)
'            If lngRetVal <> Error_Success Then
'               WritePIC = lngRetVal
'               Exit Function
'            End If
'         Next objDeletedRSObject
'         Exit Function
'      End If
'      If dblBytesFree >= lngFileSize Then
'         ' check if right PIC device is still there
'         lngRetVal = I_DMBE_ValidatePIC()
'         If lngRetVal <> Error_Success Then
'            WritePIC = lngRetVal
'            Exit Function
'         End If
'         Set objFSO = New FileSystemObject
'         objFSO.CopyFile strTempPath, strPath, True
'         objFSO.DeleteFile strTempPath, True
'         Set objFSO = Nothing
'      Else
'         Set objFSO = New FileSystemObject
'         If objFSO.FileExists(strTempPath) Then
'            objFSO.DeleteFile strTempPath, True ' just in case of failure last time in clearing file
'         End If
'         Set objFSO = Nothing
'         ' delete all of this data type from output device - it is all or nothing here
'         For Each objDeletedRSObject In colRSObjects
'            lngRetVal = DeletePIC(m_strPicPath, objDeletedRSObject.TableName, strPersistedFileType)
'            If lngRetVal <> Error_Success Then
'               WritePIC = lngRetVal
'               Exit Function
'            End If
'         Next objDeletedRSObject
'         WritePIC = ERROR_INSUFFICIENT_SPACE_ON_DEVICE ' bad return - insufficient space on output device
'         Exit Function
'      End If
'   End If
'
'   If lngRecordCount = 0 Then
'      If lngRSCount <> 0 Then
'         WritePIC = Error_Success ' good return
'      Else
'         WritePIC = ERROR_NO_RECORDS_FOR_DATASET_FOUND  ' bad return
'      End If
'   Else
'      If colRSObjects.Count = lngRecordCount Then
'         WritePIC = ERROR_INVALID_COLLECTION ' bad return
'      Else
'         WritePIC = ERROR_SUCCESS_WITH_SELECTED_RECORDSETS ' good/bad return
'      End If
'   End If
'   Exit Function
'WritePICErrTrap:
'   Dim ErrMsg As String
'   If Err.Number = 53 Then ' file does not exist
'      Err.Clear
'      Resume Next
'   End If
'   If gobjShared Is Nothing Then
'      Set gobjShared = New CWShared
'   End If
'   gobjShared.ShowVBError Err.Number, Err.Description, "WritePIC Function", "DMBE", vbExclamation
'   WritePIC = ERROR_INVALID_COLLECTION  ' bad return
'   'Resume
'End Function

'Private Function WritePICNew(ByVal colRSObjs As Collection, ByVal enmDMFormat As DMFormats, ByVal strFilePath As String) As Long
'
'   Dim strPersistedFileType As String
'   Dim rsData As ADODB.Recordset
'   Dim oRsObj As IDMBE.I_RSObjects
'   Dim oRsDeleteObj As IDMBE.I_RSObjects
'   Dim strFileName As String
'   Dim lngReturnValue As Long
'   Dim strPath As String
'   Dim strTempPath As String
'   Dim oFS As FileSystemObject
'   Dim lngFileSize As Long
'   Dim dblBytesFree As Double
'   Dim strFieldNames As String
'   Dim strFN() As String
'   Dim strUpperLimit As Long
'   Dim strCat As String
'   Dim strVersion As String
'   Dim strClinicalTableName As String
'   Dim lngJ As Long
'   Dim lngRecordCount As Long
'   Dim enmDataSet As IDMBE.DMDataSets
'   Dim strTableName As String
'   Dim lngUIdx As Long
'   Dim strTemp() As String
'   Dim lngRSCount As Long
'
'   'On Error GoTo WritePICErrTrap
'
'   If colRSObjs.Count = 0 Then
'      WritePICNew = ERROR_NO_RECORDS_FOR_DATASET_FOUND
'      Exit Function
'   End If
'
'   ' check for persisted recordset format
'   Select Case enmDMFormat
'      Case eDM_ADT
'            strPersistedFileType = ".ADT"
'      Case eDM_XML
'            strPersistedFileType = ".XML"
'      Case eDM_DAT
'            strPersistedFileType = ".DAT"
'      Case Else
'            WritePICNew = ERROR_INVALID_FORMAT ' bad return - data format not initialized
'            Exit Function
'   End Select
'
'
'
'   If lngReturnValue = ERROR_SUCCESS Then
'      ' get recordsets from collection for deleting from the PIC
'      For Each oRsObj In colRSObjs
'         If Not (oRsObj Is Nothing) Then
'            lngReturnValue = DeletePIC(strFilePath, oRsObj.TableName, ".*")
'            If lngReturnValue <> ERROR_SUCCESS Then
'               WritePICNew = lngReturnValue
'               Exit Function
'            End If
'         Else
'            WritePICNew = ERROR_COLLECTION_RSOBJECT_IS_NOTHING ' bad return
'            Exit Function
'         End If
'      Next oRsObj
'   Else
'      WritePICNew = ERROR_PIC_DEVICE_IS_NOT_FOR_CURRENT_PATIENT
'   End If
'   ' get recordsets from collection for adding to the PIC
'   strCat = ""
'   lngRSCount = 0
'
'
'
'
'
'
'   For Each oRsObj In colRSObjs
'      Set rsData = oRsObj.DBRECORDSET
'      If rsData.RecordCount = 0 Then
'         Set rsData = Nothing
'      Else
'         lngRSCount = lngRSCount + rsData.RecordCount
'      End If
'      If Not (rsData Is Nothing) Then
'         If rsData.Fields.Count <> 0 Then
'            strPath = FormatPath(strFilePath, oRsObj.TableName, strPersistedFileType)
'            ' get rid of records of this type on the PIC
'            strPath = EnsureUniqueFile(strPath)
'            Select Case UCase(strPersistedFileType)
'               Case ".XML", ".ADT"
'                     strTempPath = Environ("TEMP") & "\" & oRsObj.TableName & strPersistedFileType
'                     strTempPath = Replace(strTempPath, "\\", "\", 1, -1, vbTextCompare)
'                     Set oFS = New FileSystemObject
'                     If oFS.FileExists(strTempPath) Then
'                        oFS.DeleteFile strTempPath, True ' just in case of failure last time in clearing file
'                     End If
'                     Set oFS = Nothing
'                     If UCase(strPersistedFileType) = ".XML" Then
'                        rsData.Save strTempPath, adPersistXML
'                     End If
'                     If UCase(strPersistedFileType) = ".ADT" Then
'                        rsData.Save strTempPath, adPersistADTG
'                     End If
'                     dblBytesFree = DiskSpaceCheck(strPath)
'                     If dblBytesFree < 0 Then
'                        WritePICNew = dblBytesFree
'                        Kill strTempPath
'                        ' delete all of this data type from output device - it is all or nothing here
'                        For Each oRsDeleteObj In colRSObjs
'                           lngReturnValue = DeletePIC(strFilePath, oRsDeleteObj.TableName, strPersistedFileType)
'                           If lngReturnValue <> ERROR_SUCCESS Then
'                              WritePICNew = lngReturnValue
'                              Exit Function
'                           End If
'                        Next oRsDeleteObj
'                        Exit Function
'                     End If
'                     If m_enmCompression = eDM_Compressed Then
'                        lngReturnValue = CompressDataFile(strTempPath)
'                        If lngReturnValue <> 0 Then
'                           WritePICNew = lngReturnValue
'                           Exit Function
'                        End If
'                     End If
'                     If m_enmCrypto = eDM_Encrypted Then
'                        lngReturnValue = EncryptDataFile(strTempPath)
'                        If lngReturnValue <> 0 Then
'                           WritePICNew = lngReturnValue
'                           Exit Function
'                        End If
'                     End If
'                     lngFileSize = FileLen(strTempPath)
'                     If dblBytesFree >= lngFileSize Then
'                        ' check if right PIC device is still there
'                        lngReturnValue = I_DMBE_ValidatePIC()
'                        If lngReturnValue <> ERROR_SUCCESS Then
'                           WritePICNew = lngReturnValue
'                           Exit Function
'                        End If
'                        Set oFS = New FileSystemObject
'                        oFS.CopyFile strTempPath, strPath, True
'                        oFS.DeleteFile strTempPath, True
'                        Set oFS = Nothing
'                     Else
'                        ' check if right PIC device is still there
'                        lngReturnValue = I_DMBE_ValidatePIC()
'                        If lngReturnValue <> ERROR_SUCCESS Then
'                           WritePICNew = lngReturnValue
'                           Exit Function
'                        End If
'                        ' delete temp file
'                        Kill strTempPath
'                        ' delete all of this data type from output device - it is all or nothing here
'                        For Each oRsDeleteObj In colRSObjs
'                           lngReturnValue = DeletePIC(strFilePath, oRsDeleteObj.TableName, strPersistedFileType)
'                           If lngReturnValue <> ERROR_SUCCESS Then
'                              WritePICNew = lngReturnValue
'                              Exit Function
'                           End If
'                        Next oRsDeleteObj
'                        WritePICNew = ERROR_INSUFFICIENT_SPACE_ON_DEVICE ' bad return - insufficient space on output device
'                        Exit Function
'                     End If
'               Case ".DAT"
'                     strClinicalTableName = GetClinicalDataAreaByEnum(m_enmDataSet)
'                     strTempPath = Environ("TEMP") & "\" & strClinicalTableName & strPersistedFileType
'                     strTempPath = Replace(strTempPath, "\\", "\", 1, -1, vbTextCompare)
'                     Set oFS = New FileSystemObject
'                     If oFS.FileExists(strTempPath) Then
'                        oFS.DeleteFile strTempPath, True ' just in case of failure last time in clearing file
'                     End If
'                     Set oFS = Nothing
'                     strVersion = ""
'                     ' take care of special case
'                     strTableName = oRsObj.TableName
'                     If UCase(strTableName) = "IMMUNIZATION" Then
'                        strTableName = "IMMUNIZATIONS"
'                     End If
'                     strFieldNames = GetFieldNames(strTableName, strVersion)
'                     If strFieldNames <> "" Then
'                        strFN = Split(strFieldNames, DELIMITOR_CONSTANT, -1, vbTextCompare)
'                        lngUIdx = UBound(strFN) - 1
'                        For lngJ = 0 To lngUIdx
'                           strTemp = Split(strFN(lngJ), ",", -1, vbTextCompare)
'                           strFN(lngJ) = strTemp(0)
'                        Next lngJ
'                        strUpperLimit = UBound(strFN) - 1
'                        If rsData.EOF = False Then
'                           rsData.MoveFirst
'                           Do
'                              strCat = strCat & strTableName & ":" & strVersion & DELIMITOR_CONSTANT
'                              ' go through the field names for the current database table version
'                              For lngJ = 0 To strUpperLimit
'                                 strCat = strCat & rsData.Fields(strFN(lngJ)).Value & DELIMITOR_CONSTANT
'                              Next lngJ
'                              strCat = strCat & vbCr
'                              ' save file to temp location while we continue to aggregate recordsets
'                              Call WriteFileContents(strTempPath, strCat)
'                              rsData.MoveNext
'                           Loop While rsData.EOF = False
'                        End If
'                     Else
'                        WritePICNew = ERROR_TABLE_COLUMN_NAMES_NOT_LOADED
'                        Exit Function
'                     End If
'            End Select
'            ' destroy recordset object
'            Set rsData = Nothing
'         Else
'            lngRecordCount = lngRecordCount + 1
'         End If
'      End If
'   Next oRsObj
'   ' special processing of DAT file
'   If UCase(strPersistedFileType) = ".DAT" Then
'      If m_enmCompression = eDM_Compressed Then
'         lngReturnValue = CompressDataFile(strTempPath)
'         If lngReturnValue <> 0 Then
'            WritePICNew = lngReturnValue
'            Exit Function
'         End If
'      End If
'      If m_enmCrypto = eDM_Encrypted Then
'         lngReturnValue = EncryptDataFile(strTempPath)
'         If lngReturnValue <> 0 Then
'            WritePICNew = lngReturnValue
'            Exit Function
'         End If
'      End If
'      strPath = FormatPath(strFilePath, strClinicalTableName, strPersistedFileType)
'      lngFileSize = FileLen(strTempPath)
'      dblBytesFree = DiskSpaceCheck(strPath)
'      If dblBytesFree < 0 Then
'         WritePICNew = dblBytesFree
'         Kill strTempPath
'         ' delete all of this data type from output device - it is all or nothing here
'         For Each oRsDeleteObj In colRSObjs
'            lngReturnValue = DeletePIC(strFilePath, oRsDeleteObj.TableName, strPersistedFileType)
'            If lngReturnValue <> ERROR_SUCCESS Then
'               WritePICNew = lngReturnValue
'               Exit Function
'            End If
'         Next oRsDeleteObj
'         Exit Function
'      End If
'      If dblBytesFree >= lngFileSize Then
'         ' check if right PIC device is still there
'         lngReturnValue = I_DMBE_ValidatePIC()
'         If lngReturnValue <> ERROR_SUCCESS Then
'            WritePICNew = lngReturnValue
'            Exit Function
'         End If
'         Set oFS = New FileSystemObject
'         oFS.CopyFile strTempPath, strPath, True
'         oFS.DeleteFile strTempPath, True
'         Set oFS = Nothing
'      Else
'         Set oFS = New FileSystemObject
'         If oFS.FileExists(strTempPath) Then
'            oFS.DeleteFile strTempPath, True ' just in case of failure last time in clearing file
'         End If
'         Set oFS = Nothing
'         ' delete all of this data type from output device - it is all or nothing here
'         For Each oRsDeleteObj In colRSObjs
'            lngReturnValue = DeletePIC(strFilePath, oRsDeleteObj.TableName, strPersistedFileType)
'            If lngReturnValue <> ERROR_SUCCESS Then
'               WritePICNew = lngReturnValue
'               Exit Function
'            End If
'         Next oRsDeleteObj
'         WritePICNew = ERROR_INSUFFICIENT_SPACE_ON_DEVICE ' bad return - insufficient space on output device
'         Exit Function
'      End If
'   End If
'   If lngRecordCount = 0 Then
'      If lngRSCount <> 0 Then
'         WritePICNew = ERROR_SUCCESS ' good return
'      Else
'         WritePICNew = ERROR_NO_RECORDS_FOR_DATASET_FOUND  ' bad return
'      End If
'   Else
'      If colRSObjs.Count = lngRecordCount Then
'         WritePICNew = ERROR_INVALID_COLLECTION ' bad return
'      Else
'         WritePICNew = ERROR_SUCCESS_WITH_SELECTED_RECORDSETS ' good/bad return
'      End If
'   End If
'   Exit Function
'WritePICErrTrap:
'   Dim ErrMsg As String
'   If Err.Number = 53 Then ' file does not exist
'      Err.Clear
'      Resume Next
'   End If
'   If gobjShared Is Nothing Then
'      Set gobjShared = New CWShared
'   End If
'   gobjShared.ShowVBError Err.Number, Err.Description, "writepicnew Function", "DMBE", vbExclamation
'   WritePICNew = ERROR_INVALID_COLLECTION  ' bad return
'   'Resume
'End Function
'
'Private Function EnsureUniqueFile(PathName As String) As String
'   Dim fs As FileSystemObject
'   Dim Temp() As String
'   Dim SavePath As String
'   Dim Count As Long
'
'   On Error GoTo EnsureUniqueFileErrTrap
'
'   Set fs = CreateObject("Scripting.FileSystemObject")
'   SavePath = PathName
'   While fs.FileExists(SavePath)
'      Temp = Split(PathName, ".", -1, vbTextCompare)
'      SavePath = Temp(0) & "_" & Count & "." & Temp(1)
'      Count = Count + 1
'   Wend
'   Set fs = Nothing
'   EnsureUniqueFile = SavePath
'   Exit Function
'EnsureUniqueFileErrTrap:
'   Temp = Split(PathName, ".", -1, vbTextCompare)
'   SavePath = Temp(0) & "_" & "999" & "." & Temp(1)
'   EnsureUniqueFile = SavePath
'End Function

Private Function FormatPath(ByVal SavePath As String, ByVal TableName As String, ByVal PersistFileType As String) As String
' The naming convention of the individual files stored on the PIC device will be as follows:
'
'   AB_UnitNumber_TableName_DataSet_MM_DD_YYY_HH_MM_SS_DH.EXT
'
' Where
'   A is the Crypto flag = N(o encryption) or E(ncrypted)
'   B is the Compress flag = X(no compression) or C(ompressed)
'   UnitNumber is the Patient ID
'   Tablename is the name of the database table for the persisted recordset
'   DataSet indicates the TableName assoication with the Clinic Data module
'   MM is the three character month identifier
'   DD is the day of the month
'   YYYY is the four digit year
'   HH is the hour of the day when the file was created
'   MM is the minute of the day when the file was created
'   SS is the second of the day when the file was created
'   DH = AM Or PM
'   EXT = XML, ADT or DAT
'
'   The DataSet Mnemonics are as follows:
'        ALG = Allergies
'        VIT = Vitals
'        MED = Medications
'        DEM = Demographics
'        IMM = Immunizations
'        ENC = Encounters
'        RDN = Readiness
'        PRB = Problems
'        CLN = Clinical Notes

' Examples of this format is NX_UnitNumber_ALLERGIES_ALG_Apr_10_2001_11_20_33_AM.XML or EC_ALLERGIES_Apr_25_2001_4_17_52_PM.XML
'
' If the value of A is E, then the file must be decrypted first.
' If the value of B is C, then the length of the decompressed string can be found by reading the contents of the file and performing the following:
'   mStr = "171:CompressedContentsOfFile"
'   DataLen = Val(mStr) ' here DataLen will be 171 the length of the decompressed string
' The character after the colon (:) is the start of the compressed string.
'
   Dim sData As String
   Dim sCEData As String
   
If m_enmCrypto = 0 Then
   m_enmCrypto = eDM_NoCrypto
End If
If m_enmCompression = 0 Then
   m_enmCompression = eDM_NoCompression
End If
   sCEData = Chr(m_enmCrypto) & Chr(m_enmCompression) & "_" & unitNumber & "_"
   If m_enmSource = eDM_CDR Then
      sCEData = sCEData & "CDR_"
   ElseIf m_enmSource = eDM_PIC Then
      sCEData = sCEData & "PIC_"
   ElseIf m_enmSource = eDM_TMIP Then
      sCEData = sCEData & "TMIP_"
   ElseIf m_enmSource = eDM_LDDB Then
      sCEData = sCEData & "LDDB_"
   End If
   sData = Format(Now, "mmm d yyyy hh:mm:ss AMPM")
   sData = Replace(sData, ":", "_", 1, -1, vbTextCompare)
   sData = Replace(sData, " ", "_", 1, -1, vbTextCompare)
   
   '<SCR 21661 CC>
   If m_enmDataSet = eDM_Immunizations Then
        FormatPath = "\" & sCEData & UCase(TableName) & "_" & mc_strEndTableName & "_" & sData & PersistFileType
   Else
        FormatPath = "\" & sCEData & GetDataSetCode & "_" & UCase(TableName) & "_" & mc_strEndTableName & "_" & sData & PersistFileType
   End If
   
   If SavePath <> "" Then
        FormatPath = SavePath & FormatPath
   Else
        FormatPath = App.Path & FormatPath
   End If
   
'   If SavePath <> "" Then
'      FormatPath = SavePath & "\" & sCEData & GetDataSetCode & "_" & UCase(TableName) & "_" & mc_strEndTableName & "_" & sData & PersistFileType
'   Else
'      FormatPath = App.Path & "\" & sCEData & GetDataSetCode & "_" & UCase(TableName) & "_" & mc_strEndTableName & "_" & sData & PersistFileType
'   End If
   '</SCR 21661 CC>
   
   FormatPath = Replace(FormatPath, "\\", "\", 1, -1, vbTextCompare)

End Function

Private Function GetDataSetCode() As String
'   The DataSet Mnemonics are as follows:
'        ALG = Allergies
'        VIT = Vitals
'        MED = Medications
'        DEM = Demographics
'        IMM = Immunizations
'        ENC = Encounters
'        RDN = Readiness
'        PRB = Problems
'        CLN = Clinical Notes
   Select Case m_enmDataSet
      Case eDM_Allergies
         GetDataSetCode = "ALG"
      Case eDM_ClinicNotes
         GetDataSetCode = "CLN"
      Case eDM_Demographics
         GetDataSetCode = "DEM"
      Case eDM_Encounters
         GetDataSetCode = "ENC"
      Case eDM_Immunizations
         GetDataSetCode = "IMM"
      Case eDM_Medications
         GetDataSetCode = "MED"
      Case eDM_Problems
         GetDataSetCode = "PRB"
      Case eDM_Readiness
         GetDataSetCode = "RDN"
      Case eDM_Vitals
         GetDataSetCode = "VIT"
   End Select

End Function

Private Function TranslateDataSetCode(ByVal strDMCode) As Long

    Dim enmDataSet As IDMBE.DMDataSets

   Select Case strDMCode
      Case "ALG"
         enmDataSet = eDM_Allergies
      Case "CLN"
         enmDataSet = eDM_ClinicNotes
      Case "DEM"
         enmDataSet = eDM_Demographics
      Case "ENC"
         enmDataSet = eDM_Encounters
      Case "IMM", "IMMUNIZATIONS"
         enmDataSet = eDM_Immunizations
      Case "MED"
         enmDataSet = eDM_Medications
      Case "PRB"
         enmDataSet = eDM_Problems
      Case "RDN"
         enmDataSet = eDM_Readiness
      Case "VIT"
         enmDataSet = eDM_Vitals
      Case Else
         enmDataSet = 9999
   End Select
 
    TranslateDataSetCode = enmDataSet

End Function

Private Function GetFirstFile(FilePathPlusPattern As String, hSearch As Long, FileName As String) As Long
   
   Dim FileData As WIN32_FIND_DATA
   Dim tempLen As Long
   On Error GoTo GetFirstFileError
   
   ' Initialize FileName
   FileData.cFileName = " "
   ' Start searching for FilePattern files in a directory.
   If FilePathPlusPattern = "" Then
      GetFirstFile = gclNOTFOUND ' bad return
      Exit Function
   End If
   
   hSearch = FindFirstFile(FilePathPlusPattern, FileData)
   If hSearch = INVALID_HANDLE_VALUE Or hSearch = 0 Then
'      Call MsgBox("No " & FilePathPlusPattern & " files found.", vbCritical, "GetNextFile Function Error")
      GetFirstFile = gclNOTFOUND
      Exit Function
   End If
   FileName = FileData.cFileName
   ' Begin:SCR #36997;   Developer: Mag 07/17/2003 12:22 PM
   tempLen = InStr(1, FileName, Chr(0), vbTextCompare)
   FileName = Mid(FileName, 1, tempLen - 1)
   ' End: SCR #36997;
   GetFirstFile = gclFOUND
   Exit Function
GetFirstFileError:
   If gobjShared Is Nothing Then
      Set gobjShared = New CWShared
   End If
   gobjShared.ShowVBError Err.Number, Err.Description, "GetFirstFile Function", "DMBE", vbExclamation
   GetFirstFile = gclNOTFOUND ' bad return
End Function

Private Function GetNextFile(ByRef hSearch As Long, ByRef FileName As String) As Long
   Dim FileData As WIN32_FIND_DATA
   Dim tempLen As Long
   Dim ReturnValue As Long
   
   ' Initialize FileName
   FileData.cFileName = " "
   ' Start searching for FilePattern files in a directory.
   If hSearch = 0 Then
      GetNextFile = gclNOTFOUND
      Exit Function
   End If
   ReturnValue = FindNextFile(hSearch, FileData)
   If ReturnValue = 0 Then
'      Call MsgBox("No files found.", vbCritical, "GetNextFile Function Error")
      GetNextFile = gclNOTFOUND
      Call FindClose(hSearch)
      Exit Function
   End If
   FileName = FileData.cFileName
   ' Begin:SCR #36997;   Developer: Mag 07/17/2003 12:22 PM
   tempLen = InStr(1, FileName, Chr(0), vbTextCompare)
   FileName = Mid(FileName, 1, tempLen - 1)
   ' End: SCR #36997;
   GetNextFile = gclFOUND
End Function

Private Sub GetAllergyTableNames(TN() As String)
   ReDim TN(3) As String
   TN(0) = "ALLERGIES"
   TN(1) = "ALLERGYREACTION"
   TN(2) = "ALLERGYCOMMENT"
End Sub

Private Sub GetDemographicsTableNames(TN() As String)
   ReDim TN(7) As String

   TN(0) = "MMI_ID"
   TN(1) = "MMI_NAME"
   TN(2) = "MMI_ADDRESS"
   TN(3) = "MMI_PHONE"
   TN(4) = "MMI_SITE"
   TN(5) = "MMI_GENERIC_ID"
   TN(6) = "MMI_GENERICDATA"
End Sub

Private Sub GetEncountersTableNames(TN() As String)
   ReDim TN(17) As String
   TN(0) = "ENC_SECTIONS"
   TN(1) = "ENC_RTFS"
   TN(2) = "SONOTES"
   TN(3) = "VITALS"
   TN(4) = "PROBLEMSRELATIONSHIPS"
   TN(5) = "ENCOUNTERS"
   TN(6) = "ENC_DIAGNOSIS"
   TN(7) = "ENCOUNTER_LOCK"
   TN(8) = "ENC_DISPOSITION"
   TN(9) = "DIAGNOSESRELATIONSHIPS"
   TN(10) = "CLINNOTE"
   TN(11) = "WORK_STATUS"
   TN(12) = "PROCEDURES"
   TN(13) = "ENCTEXTNOTES"
   TN(14) = "PI_DATA"
   TN(15) = "COMPLAINTS"
   TN(16) = "MEDICATIONS"

   'TN(4) = "LISTTOOLRECORDS"
End Sub

Private Sub GetImmunizationsTableNames(TN() As String)
'   ReDim TN(2) As String
'   TN(0) = "ALLERGIES"
'   TN(1) = "MMI_GENERIC_ID"
   ReDim TN(1) As String
   TN(0) = "IMMUNIZATIONS"
End Sub

Private Sub GetMedicationsTableNames(TN() As String)
      ReDim TN(2) As String
      TN(0) = "MEDICATIONS"
      TN(1) = "RX_LIST"
      End Sub

Private Sub GetProblemsTableNames(TN() As String)
   ReDim TN(6) As String
   TN(0) = "PROBLEMSRELATIONSHIPS"
   TN(1) = "DIAGNOSESRELATIONSHIPS"
   TN(2) = "COMPLAINTS"
   TN(3) = "PROCEDURES"
   TN(4) = "PROBLEMS"
   TN(5) = "DIAGNOSES"
End Sub

Private Sub GetReadinessTableNames(TN() As String)
   ReDim TN(1) As String
   TN(0) = "READINESS"
End Sub

Private Sub GetVitalsTableNames(TN() As String)
   ReDim TN(3) As String
   TN(0) = "VITALS"
   TN(1) = "VITALSIGNSVALUE"
   TN(2) = "VITALSIGNSMODIFIER"
End Sub

'Private Function DeletePIC(ByRef FilePath As String, ByRef TableName As String, ByRef PersistedFileType As String) As Long
'
'   Dim Path As String
'   Dim FilePattern As String
'   Dim FileName As String
'   Dim hSearch As Long
'   Dim RetVal As Long
'
'   ' search for first file of this type
'   If m_enmDestination = eDM_AGGPIC Then
'      FilePattern = FilePath & "\" & "*" & gobjPatient.UnitNumber & "*" & TableName & "*" & PersistedFileType
'   End If
'   If m_enmDestination = eDM_PIC Then
'      FilePattern = FilePath & "\" & "*" & TableName & "*" & PersistedFileType
'   End If
'
'   FilePattern = Replace(FilePattern, "\\", "\", 1, -1, vbTextCompare)
'   FileName = ""
'
'   ' check if right PIC device is still there
'   RetVal = I_DMBE_ValidatePIC()
'   If RetVal <> ERROR_SUCCESS Then
'      DeletePIC = RetVal
'      Exit Function
'   End If
'
'   RetVal = GetFirstFile(FilePattern, hSearch, FileName)
'   If RetVal = gclFOUND Then
'      Path = FilePath & "\" & FileName
'      Path = Replace(Path, "\\", "\", 1, -1, vbTextCompare)
'      If InStr(1, UCase(Path), "XML", vbTextCompare) <> 0 Then
'         Kill (Path)
'      ElseIf InStr(1, UCase(Path), "ADT", vbTextCompare) <> 0 Then
'         Kill (Path)
'      ElseIf InStr(1, UCase(Path), "DAT", vbTextCompare) <> 0 Then
'         Kill (Path)
'      End If
'      Do
'         FileName = ""
'         ' check if right PIC device is still there
'         RetVal = I_DMBE_ValidatePIC()
'         If RetVal <> ERROR_SUCCESS Then
'            DeletePIC = RetVal
'            Exit Function
'         End If
'         ' search for next file(s) of this type
'         RetVal = GetNextFile(hSearch, FileName)
'         If RetVal = gclFOUND Then
'            Path = FilePath & "\" & FileName
'            Path = Replace(Path, "\\", "\", 1, -1, vbTextCompare)
'            If InStr(1, UCase(Path), "XML", vbTextCompare) <> 0 Then
'               Kill (Path)
'            ElseIf InStr(1, UCase(Path), "ADT", vbTextCompare) <> 0 Then
'               Kill (Path)
'            ElseIf InStr(1, UCase(Path), "DAT", vbTextCompare) <> 0 Then
'               Kill (Path)
'            End If
'         End If
'      Loop While RetVal = gclFOUND
'      FindClose hSearch
'   End If
'
'   DeletePIC = ERROR_SUCCESS ' good return
'
'
'End Function
    
Private Function DiskSpaceCheck(Path As String) As Double
   Dim RetVal As Long
   Dim lpSectorsPerCluster As Long
   Dim lpBytesPerSector As Long
   Dim lpNumberOfFreeClusters As Long
   Dim lpTotalNumberOfClusters As Long
   Dim TotalSpaceFree As Double
   Dim Drive As String
   Dim fs As FileSystemObject
   
   Set fs = New FileSystemObject
   Drive = fs.GetDriveName(Path) & "\"
   RetVal = GetDiskFreeSpace(Drive, _
                              lpSectorsPerCluster, _
                              lpBytesPerSector, _
                              lpNumberOfFreeClusters, _
                              lpTotalNumberOfClusters _
                              )
   If RetVal = 0 Then
      DiskSpaceCheck = ERROR_UNABLE_TO_CHECK_DISK_SPACE ' bad return
      Exit Function
   End If
   TotalSpaceFree = lpSectorsPerCluster * lpBytesPerSector
   TotalSpaceFree = TotalSpaceFree * lpNumberOfFreeClusters
   DiskSpaceCheck = TotalSpaceFree
   Set fs = Nothing
End Function

'Private Function GetFieldNames(TableName As String, VersionNumber As String) As String
'   Dim RS As New ADODB.Recordset
'   Dim sSQL As String
'   Dim FindFlag As Long
'   Dim TabName As String
'
'   On Error GoTo ErrTrap
'
''   If DELIM_CN Is Nothing Then
''      Set DELIM_CN = New ADODB.Connection
''      DELIM_CN.ConnectionString = DB_CNSTR
''      DELIM_CN.Open
''   End If
'   TabName = UCase(TableName)
'   sSQL = "SELECT * FROM DELIMITEDFILEMAPPING WHERE TABLENAME = '" & TabName & "' ORDER BY VERSION DESC"
'   If oDas Is Nothing Then
'      Set oDas = New GEMS_DAS
'   End If
'   Set RS = oDas.OpenRecordset(sSQL, adOpenDynamic, adLockOptimistic)
''   RS.Open sSQL, DELIM_CN, adOpenStatic, adLockOptimistic, -1
'   If RS.BOF = True And RS.EOF = True Then
'      VersionNumber = ""
'      GetFieldNames = ""
'      Exit Function
'   End If
'   If VersionNumber = "" Then
'      VersionNumber = RS.Fields("VERSION").Value  ' write function call - use most recent
'      GetFieldNames = RS.Fields("DELIMITEDFIELDNAMES").Value
'   Else
'      If RS.EOF = False Then
'         RS.MoveFirst
'         FindFlag = 0
'         Do
'            If VersionNumber = RS.Fields("VERSION").Value Then
'               FindFlag = 1
'               GetFieldNames = RS.Fields("DELIMITEDFIELDNAMES").Value
'               Set RS = Nothing
'               Exit Function
'            End If
'            RS.MoveNext
'         Loop While RS.EOF = False
'      Else
'         Call MsgBox("No records found for table (" & TableName & ").", vbCritical, "GetFieldNames Function")
'         GetFieldNames = ""
'         Set RS = Nothing
'         Exit Function
'      End If
'   End If
'   Exit Function
'ErrTrap:
'   If gobjShared Is Nothing Then
'      Set gobjShared = New CWShared
'   End If
'   gobjShared.ShowVBError Err.Number, Err.Description, "GetFieldNames Function", "DMBE", vbExclamation
'   GetFieldNames = ""  ' bad return
''   Resume
'End Function

'Private Sub AllergiesTemporaryFix(cObj As Collection)
'   Dim vItem As Variant
'   Dim RS As New ADODB.Recordset
'   Dim sSQL As String
'   Dim Data As I_RSObjects
'   Dim colRSData As New Collection
'   Dim RetVal As Long
'
'   ' get unit_number form collection
'   For Each vItem In cObj
'      Set Data = New RSObject_Ops
'      sSQL = "SELECT * FROM ALLERGIES WHERE UNIT_NUMBER = '" & vItem & "'"
'      Set RS = GetTableDataDAS(sSQL)
'      If RS.EOF = False Then
'         Data.TableName = "ALLERGIES"
'         Set Data.DBRECORDSET = RS
'         colRSData.Add Data
'         Set Data = Nothing
'      Else
'         Call MsgBox("No data retrieved from LDDB for ALLERGIES.", vbCritical, "Execute Sub")
'      End If
'      Set RS = Nothing
'      Set RS = New ADODB.Recordset
'      Set Data = New RSObject_Ops
'      sSQL = "SELECT * FROM ALLERGYREACTION WHERE UNIT_NUMBER = '" & vItem & "'"
'      Set RS = GetTableDataDAS(sSQL)
'      If RS.EOF = False Then
'         Data.TableName = "ALLERGYREACTION"
'         Set Data.DBRECORDSET = RS
'         colRSData.Add Data
'         Set Data = Nothing
'      Else
'         Call MsgBox("No data retrieved from LDDB for ALLERGYREACTION.", vbCritical, "Execute Sub")
'      End If
'      Set RS = Nothing
'      Set RS = New ADODB.Recordset
'      Set Data = New RSObject_Ops
'      sSQL = "SELECT * FROM ALLERGYCOMMENT WHERE UNIT_NUMBER = '" & vItem & "'"
'      Set RS = GetTableDataDAS(sSQL)
'      If RS.EOF = False Then
'         Data.TableName = "ALLERGYCOMMENT"
'         Set Data.DBRECORDSET = RS
'         colRSData.Add Data
'         Set Data = Nothing
'      Else
'         Call MsgBox("No data retrieved from LDDB for ALLERGYCOMMENT.", vbCritical, "Execute Sub")
'      End If
'      Set RS = Nothing
'
''      Set colRSData = Nothing
'   Next vItem
'   Set cObj = Nothing
'   Set cObj = colRSData
'   ' set the path for PIC device
''   PICPATH = "A:\"
'End Sub

'Private Sub ReadinessTemporaryFix(cObj As Collection)
'   Dim vItem As Variant
'   Dim RS As New ADODB.Recordset
'   Dim sSQL As String
'   Dim Data As I_RSObjects
'   Dim colRSData As New Collection
'   Dim RetVal As Long
'
'   Set Data = New RSObject_Ops
'   sSQL = "select r.*, e.PRIMARYPROVIDERNCID, e.CLINICNCID, e.TYPE from readiness r, encounters e where r.ENCOUNTER_NUMBER = e.ENCOUNTERNUMBER(+) and r.PATIENT_UNIT_NUMBER = <?1>"
'   sSQL = Replace(sSQL, "<?1>", gobjPatient.UnitNumber, 1, -1, vbTextCompare)
'   Set RS = GetTableDataDAS(sSQL)
'   If RS.EOF = False Then
'      Data.TableName = "READINESS"
'      Set Data.DBRECORDSET = RS
'      colRSData.Add Data
'      Set Data = Nothing
'   Else
'      Call MsgBox("No data retrieved from LDDB for Readiness.", vbCritical, "ReadinessTemporaryFix")
'   End If
'   Set RS = Nothing
'
''      Set colRSData = Nothing
'   Set cObj = Nothing
'   Set cObj = colRSData
'   ' set the path for PIC device
''   PICPATH = "A:\"
'End Sub

'Private Sub VitalsTemporaryFix(cObj As Collection)
'   Dim vItem As Variant
'   Dim RS As New ADODB.Recordset
'   Dim sSQL As String
'   Dim Data As I_RSObjects
'   Dim colRSData As New Collection
'   Dim RetVal As Long
'
'   Set Data = New RSObject_Ops
'   sSQL = "select v.*, vv.*, vm.*, e.PRIMARYPROVIDERNCID, e.CLINICNCID, e.Type from vitals v, vitalsignsvalue vv, vitalsignsmodifier vm, encounters e where v.unit_number = <?1> and v.dataid = vv.dataid(+) and v.dataid = vm.dataid(+) and v.ENCOUNTERNUMBER = e.ENCOUNTERNUMBER(+)"
'   sSQL = Replace(sSQL, "<?1>", gobjPatient.UnitNumber, 1, -1, vbTextCompare)
'   Set RS = GetTableDataDAS(sSQL)
'   If RS.EOF = False Then
'      Data.TableName = "VITALS"
'      Set Data.DBRECORDSET = RS
'      colRSData.Add Data
'      Set Data = Nothing
'   Else
'      Call MsgBox("No data retrieved from LDDB for Vitals.", vbCritical, "VitalsTemporaryFix")
'   End If
'   Set RS = Nothing
'
''      Set colRSData = Nothing
'   Set cObj = Nothing
'   Set cObj = colRSData
'   ' set the path for PIC device
''   PICPATH = "A:\"
'End Sub

'Private Sub MedicationsTemporaryFix(cObj As Collection)
'   Dim vItem As Variant
'   Dim RS As New ADODB.Recordset
'   Dim sSQL As String
'   Dim Data As I_RSObjects
'   Dim colRSData As New Collection
'   Dim RetVal As Long
'
'   Set Data = New RSObject_Ops
'   sSQL = "select m.* from Medications m where m.unit_number = <?1>"
'   sSQL = Replace(sSQL, "<?1>", gobjPatient.UnitNumber, 1, -1, vbTextCompare)
'   Set RS = GetTableDataDAS(sSQL)
'   If RS.EOF = False Then
'      Data.TableName = "MEDICATIONS"
'      Set Data.DBRECORDSET = RS
'      colRSData.Add Data
'      Set Data = Nothing
'   Else
'      Call MsgBox("No data retrieved from LDDB for Medications.", vbCritical, "MedicationsTemporaryFix")
'   End If
'   Set RS = Nothing
'
''      Set colRSData = Nothing
'   Set cObj = Nothing
'   Set cObj = colRSData
'   ' set the path for PIC device
''   PICPATH = "A:\"
'End Sub

'Private Sub EncountersTemporaryFix(cObj As Collection)
'   Dim vItem As Variant
'   Dim EncountersRS As New ADODB.Recordset
'   Dim DiagnosesRS As New ADODB.Recordset
'   Dim Enc_DiagnosisRS As New ADODB.Recordset
'   Dim ProcedureRS As New ADODB.Recordset
'   Dim Enc_SectionsRS As New ADODB.Recordset
'   Dim sSQL As String
'   Dim Data As I_RSObjects
'   Dim colRSData As New Collection
'   Dim RetVal As Long
'   Dim UN As Variant
'   Dim EN As Variant
'
'   sSQL = "select e.*, ed.* from encounters e, enc_disposition ed where e.unit_number = 128563 and e.encounternumber = ed.encounternumber(+)"
'   Set EncountersRS = GetTableDataDAS(sSQL)
'   If EncountersRS.EOF = False Then
'      Set Data = New RSObject_Ops
'      Data.TableName = "ENCOUNTERS"
'      Set Data.DBRECORDSET = EncountersRS
'      colRSData.Add Data
'      Set Data = Nothing
'   Else
'      Call MsgBox("No data retrieved from LDDB for Encounters.", vbCritical, "Execute Sub")
'   End If
'
'   If Not (EncountersRS Is Nothing) Then
'      EncountersRS.MoveFirst
'      While EncountersRS.EOF = False
'         UN = EncountersRS.Fields("Unit_Number").Value
'         If IsNull(UN) Then
'            UN = ""
'         End If
'         EN = EncountersRS.Fields("EncounterNumber").Value
'         If IsNull(EN) Then
'            EN = ""
'         End If
'         sSQL = "select d.*, dr.* from diagnoses d, diagnosesrelationships dr where d.unit_number = <?1> and d.encounternumber = <?2> and d.dataid = dr.diagnosisdataid(+)"
'         sSQL = Replace(sSQL, "<?1>", UN, 1, -1, vbTextCompare)
'         If EN = "" Then  ' remove this after test just to get data for test
'            sSQL = "select d.*, dr.* from diagnoses d, diagnosesrelationships dr where d.unit_number = 200001 and d.dataid = dr.diagnosisdataid(+)"
'         Else
'            sSQL = Replace(sSQL, "<?2>", EN, 1, -1, vbTextCompare)
'         End If
'         Set DiagnosesRS = GetTableDataDAS(sSQL)
'         If DiagnosesRS.EOF = False Then
'            Set Data = New RSObject_Ops
'            Data.TableName = "DIAGNOSES_" & EN
'            Set Data.DBRECORDSET = DiagnosesRS
'            colRSData.Add Data
'            Set Data = Nothing
'         Else
'            Call MsgBox("No data retrieved from LDDB for Encounters.", vbCritical, "Execute Sub")
'         End If
'
'         sSQL = "select * from enc_diagnosis where encounternumber = 1262"  ' change 1262 to <?2>
'         EN = 1262
'         sSQL = Replace(sSQL, "<?2>", EN, 1, -1, vbTextCompare)
'         If EN <> "" Then  ' this was a null value
'            Set Enc_DiagnosisRS = GetTableDataDAS(sSQL)
'            If Enc_DiagnosisRS.EOF = False Then
'               Set Data = New RSObject_Ops
'               Data.TableName = "ENC_DIAGNOSES_" & EN
'               Set Data.DBRECORDSET = Enc_DiagnosisRS
'               colRSData.Add Data
'               Set Data = Nothing
'            Else
'               Call MsgBox("No data retrieved from LDDB for Encounters.", vbCritical, "Execute Sub")
'            End If
'         End If
'
'         sSQL = "select * from procedures where unit_number = <?1> and encounternumber = <?2>"
'         sSQL = "select * from procedures where unit_number = 200001" ' remove later
'         sSQL = Replace(sSQL, "<?1>", UN, 1, -1, vbTextCompare)
'         sSQL = Replace(sSQL, "<?2>", EN, 1, -1, vbTextCompare)
'         If EN <> "" Then  ' this was a null value
'            Set ProcedureRS = GetTableDataDAS(sSQL)
'            If ProcedureRS.EOF = False Then
'               Set Data = New RSObject_Ops
'               Data.TableName = "PROCEDURES_" & EN
'               Set Data.DBRECORDSET = ProcedureRS
'               colRSData.Add Data
'               Set Data = Nothing
'            Else
'               Call MsgBox("No data retrieved from LDDB for Encounters.", vbCritical, "Execute Sub")
'            End If
'         End If
'
'         sSQL = "select * from enc_sections where encounternumber = 1262"  ' change 1262 to <?2>
'         sSQL = Replace(sSQL, "<?2>", EN, 1, -1, vbTextCompare)
'         If EN <> "" Then  ' this was a null value
'            Set Enc_SectionsRS = GetTableDataDAS(sSQL)
'            If Enc_SectionsRS.EOF = False Then
'               Set Data = New RSObject_Ops
'               Data.TableName = "ENC_SECTIONS_" & EN
'               Set Data.DBRECORDSET = Enc_SectionsRS
'               colRSData.Add Data
'               Set Data = Nothing
'            Else
'               Call MsgBox("No data retrieved from LDDB for Encounters.", vbCritical, "Execute Sub")
'            End If
'         End If
'
'         EncountersRS.MoveNext
'      Wend
'   End If
'
'   EncountersRS.MoveFirst
'   Set EncountersRS = Nothing
'   Set DiagnosesRS = Nothing
'   Set Enc_DiagnosisRS = Nothing
'   Set ProcedureRS = Nothing
'   Set Enc_SectionsRS = Nothing
''      Set colRSData = Nothing
'   Set cObj = Nothing
'   Set cObj = colRSData
'   ' set the path for PIC device
''   PICPATH = "A:\"
'End Sub


'Private Sub DemographicsTemporaryFix(cObj As Collection)
'   Dim vItem As Variant
'   Dim RS As New ADODB.Recordset
'   Dim sSQL As String
'   Dim Data As I_RSObjects
'   Dim colRSData As New Collection
'   Dim RetVal As Long
'
'   Set Data = New RSObject_Ops
'   sSQL = "select id.*, name.NAME_INDEX, name.LAST_NAME_2, name.MIDDLE_NAME_2, name.FIRST_NAME_SOUNDEX, name.LAST_NAME_SOUNDEX, site.RADIOLOGY_NUMBER from mmi_id id, mmi_name name, mmi_site site where id.unit_number = <?1> and id.unit_number = name.unit_number(+) and id.unit_number = site.unit_number(+)"
'   sSQL = Replace(sSQL, "<?1>", gobjPatient.UnitNumber, 1, -1, vbTextCompare)
'   Set RS = GetTableDataDAS(sSQL)
'   If RS.EOF = False Then
'      Data.TableName = "MMI_ID"
'      Set Data.DBRECORDSET = RS
'      colRSData.Add Data
'      Set Data = Nothing
'   Else
'      Call MsgBox("No data retrieved from LDDB for Demographics MMI_ID.", vbCritical, "DemographicsTemporaryFix")
'   End If
'   Set RS = Nothing
'
'   Set Data = New RSObject_Ops
'   sSQL = "select * from mmi_genericdata where unit_number = <?1>"
'   sSQL = Replace(sSQL, "<?1>", gobjPatient.UnitNumber, 1, -1, vbTextCompare)
'   Set RS = GetTableDataDAS(sSQL)
'   If RS.EOF = False Then
'      Data.TableName = "MMI_GENERICDATA"
'      Set Data.DBRECORDSET = RS
'      colRSData.Add Data
'      Set Data = Nothing
'   Else
'      Call MsgBox("No data retrieved from LDDB for Demographics MMI_GENERICDATA.", vbCritical, "DemographicsTemporaryFix")
'   End If
'   Set RS = Nothing
'
'   Set Data = New RSObject_Ops
'   sSQL = "select * from mmi_inst_name where unit_number = <?1>"
'   sSQL = Replace(sSQL, "<?1>", gobjPatient.UnitNumber, 1, -1, vbTextCompare)
'   Set RS = GetTableDataDAS(sSQL)
'   If RS.EOF = False Then
'      Data.TableName = "MMI_INST_NAME"
'      Set Data.DBRECORDSET = RS
'      colRSData.Add Data
'      Set Data = Nothing
'   Else
'      Call MsgBox("No data retrieved from LDDB for Demographics MMI_INST_NAME.", vbCritical, "DemographicsTemporaryFix")
'   End If
'   Set RS = Nothing
'
''      Set colRSData = Nothing
'   Set cObj = Nothing
'   Set cObj = colRSData
'   ' set the path for PIC device
''   PICPATH = "A:\"
'End Sub


' Decision table for when to read/write binary and when to read/write ascii
'
'              |  Encrypt     |   Decrypt            |        Compress         |  UnCompress
'    --------------------------------------------------------------------------------------------
'              | read ascii   |                      | depends on encrypt flag |
'     WritePIC |              |                      |                         |
'              | write binary |                      | write binary            |
'              |              |                      |                         |
'    --------------------------------------------------------------------------------------------
'              |              | read binary          |                         | read binary
'     ReadPIC  |              |                      |                         |
'              |              | write binary         |                         | write ascii
'              |              | depends on compress  |                         |
'    --------------------------------------------------------------------------------------------
'
'  Assumptions for WritePIC: encrypt first and compress second
'  Assumptions for ReadPIC: decompress first and decrypt second
'
' EncryptDataFile function is only called from WritePIC
' DecryptDataFile function is only called from ReadPIC
' CompressDataFile function is only called from WritePIC
' DecompressDataFile function is only called from ReadPIC
'

Private Function EncryptDataFile(FilePath As String) As Long
   Dim Data As String
   Dim EncryptedData As String
   Dim cipher As CCrypto
   ' get file contents
   Data = ReadBinaryData(FilePath)
   ' encrypt the file
   Set cipher = New CCrypto
   EncryptedData = cipher.Encrypt(Data, mc_strPassword)
   If EncryptedData = "" Then
      EncryptDataFile = ERROR_WITH_ENCRYPTION_FUNCTION ' bad return
      Exit Function
   End If
   ' write the file contents
   Call PutBinaryData(FilePath, EncryptedData)
   Set cipher = Nothing
   EncryptDataFile = ERROR_SUCCESS ' good return
End Function

Private Function DecryptDataFile(FilePath As String) As Long
   Dim Data As String
   Dim DecryptedData As String
   Dim cipher As CCrypto
   
   ' get file contents
   Data = ReadBinaryData(FilePath)
   ' decrypt the file
   Set cipher = New CCrypto
   DecryptedData = cipher.Decrypt(Data, mc_strPassword)
   If DecryptedData = "" Then
      DecryptDataFile = ERROR_WITH_DECRYPTION_FUNCTION ' bad return
      Exit Function
   End If
   ' write the file contents
   Call PutBinaryData(FilePath, DecryptedData)
   Set cipher = Nothing
   DecryptDataFile = ERROR_SUCCESS ' good return
End Function

Private Function CompressDataFile(FilePath As String) As Long
   Dim Data As String
   Dim CompressData As String
   ' get the file contents
   Data = ReadFileContents(FilePath)
   ' compress the data
   CompressData = CompressString(Data)
   If CompressData = "" Then
      CompressDataFile = ERROR_WITH_COMPRESSION_FUNCTION ' bad return
      Exit Function
   End If
   ' write the file contents
   Call PutBinaryData(FilePath, CompressData)
   CompressDataFile = ERROR_SUCCESS ' good return
End Function

Private Function DecompressDataFile(FilePath As String) As Long
   Dim Data As String
   Dim DecompressData As String
   ' get the file contents
   Data = ReadBinaryData(FilePath)
   ' decompress the data
   DecompressData = UncompressString(Data)
   If DecompressData = "" Then
      DecompressDataFile = ERROR_WITH_DECOMPRESSION_FUNCTION ' bad return
      Exit Function
   End If
   ' write the file contents
   Call WriteFileContents(FilePath, DecompressData)
   DecompressDataFile = ERROR_SUCCESS ' good return
End Function

Public Function ReadFileContents(FullFilePath As String, _
                                  Optional bCompress As Boolean = False, _
                                  Optional bEncrypt As Boolean = False, _
                                  Optional nError As Long = 0, _
                                  Optional bForceBinary As Boolean = False) As String
   Dim fso As FileSystemObject
   Dim TextObj As TextStream
   Dim Data As String
   Dim cipher As CCrypto
   
   On Error GoTo ReadFileContentsErrTrap
   
   Set fso = New FileSystemObject
   If bEncrypt Or bCompress Or bForceBinary Then
     Data = ReadBinaryData(FullFilePath)
   Else
     ' get text stream object
     ' Begin:SCR #34311;   Developer: Mag 04/04/2003 02:21 PM
     'DBM 4/4/2003 - Change this to write ASCII file instead of Unicode by changing TristateTrue to TristateFalse
     Set TextObj = fso.OpenTextFile(FullFilePath, ForReading, False, TristateFalse)
     ' End: SCR #34311;
     
     If TextObj.AtEndOfStream Then
       ReadFileContents = ""
       GoTo Cleanup
     End If
     
     ' read the contents of file
     Data = TextObj.ReadAll
     TextObj.Close
     
     If Data = "" Then
       ReadFileContents = Data
       GoTo Cleanup
     End If
   
     ' Begin:SCR #34311;   Developer: Mag 04/04/2003 02:21 PM
     If Asc(Left(Data, 1)) = 255 Or Asc(Mid(Data, 2, 1)) = 0 Then
       'Reading in ASCII File was invalid so try Unicode read
       'As of 4/4/2003 all Files should be written in ASCII format
       Set TextObj = fso.OpenTextFile(FullFilePath, ForReading, False, TristateTrue)
       
       If TextObj.AtEndOfStream Then
         ReadFileContents = ""
         GoTo Cleanup
       End If
       
       Data = TextObj.ReadAll
       TextObj.Close
     End If
     ' End: SCR #34311;
     
     ' destroy the objects
     Set TextObj = Nothing
     Set fso = Nothing
   End If

'Decrypt first then uncompress, based on how WriteFileContents order.
   'Check for Encrypt
   If bEncrypt Then
      Set cipher = New CCrypto
      Data = cipher.Decrypt(Data, mc_strPassword)
      If Data = "" Then
        nError = ERROR_WITH_DECRYPTION_FUNCTION
        ReadFileContents = ""
        GoTo Cleanup
      End If
      Set cipher = Nothing
   End If
   
   'Check for Compressed
   If bCompress Then
      Data = UncompressString(Data)
      If Data = "" Then
        nError = ERROR_WITH_DECOMPRESSION_FUNCTION
        ReadFileContents = ""
        GoTo Cleanup
      End If
   End If
  
   ReadFileContents = Data
   
Cleanup:
   On Error Resume Next
   TextObj.Close
   Set TextObj = Nothing
   Set fso = Nothing
   Exit Function

ReadFileContentsErrTrap:
   If gobjShared Is Nothing Then
      Set gobjShared = New CWShared
   End If
   gobjShared.ShowVBError Err.Number, Err.Description, "ReadFileContents Function", "DMBE", vbExclamation
   Err.Clear
   Resume Cleanup
   'Resume Next
End Function

Private Function GetPicFormat(ByVal FullFilePath As String _
                            , Optional ByVal sUnitNumber As String = "" _
                            , Optional ByVal vDefault As DMFormats = -1) As DMFormats
Dim hSearch As Long
Dim sFileName As String
On Error GoTo ErrHandler
    If Right(FullFilePath, 1) <> "\" Then
        FullFilePath = FullFilePath & "\"
    End If
    
    sFileName = sUnitNumber & "*"
    If GetFirstFile(FullFilePath & sFileName & ".XML", hSearch, sFileName) = gclFOUND Then
        GetPicFormat = eDM_XML
    ElseIf GetFirstFile(FullFilePath & sFileName & ".ADT", hSearch, sFileName) = gclFOUND Then
        GetPicFormat = eDM_ADT
    ElseIf GetFirstFile(FullFilePath & sFileName & ".DAT", hSearch, sFileName) = gclFOUND Then
        GetPicFormat = eDM_DAT
    Else
        GetPicFormat = vDefault
    End If
    
    FindClose hSearch
Exit Function

ErrHandler:
   If gobjShared Is Nothing Then
      Set gobjShared = New CWShared
   End If
   gobjShared.ShowVBError Err.Number, Err.Description, "GetPICFormat Function", "DMBE", vbExclamation
   Err.Clear
   'Resume Next
End Function

'This will create a File in the user's Temp folder
Public Function WriteFileContents(FullFilePath As String, ByVal Data As String, _
                        Optional bCompress As Boolean = False, _
                        Optional bEncrypt As Boolean = False, _
                        Optional bForceBinary As Boolean = False) As Long
   
   Dim fso As FileSystemObject
   Dim TextObj As TextStream
   Dim cipher As CCrypto
   
   On Error GoTo WriteFileContentsErrTrap
   
   If Data = "" Then
      WriteFileContents = ERROR_NO_DATA_TO_WRITE ' bad return
      Exit Function
   End If
   
   If bCompress Then
      ' compress the data
      Data = CompressString(Data)
      If Data = "" Then
         WriteFileContents = ERROR_WITH_COMPRESSION_FUNCTION ' bad return
         Exit Function
      End If
   End If
     
   If bEncrypt Then
      ' encrypt the file
      Set cipher = New CCrypto
      Data = cipher.Encrypt(Data, mc_strPassword)
      If Data = "" Then
         WriteFileContents = ERROR_WITH_ENCRYPTION_FUNCTION ' bad return
         Exit Function
      End If
   End If

   If bCompress Or bEncrypt Or bForceBinary Then
        WriteFileContents = PutBinaryData(FullFilePath, Data)
   Else
     Set fso = New FileSystemObject
     

     'Set TextObj = fso.OpenTextFile(FullFilePath, ForWriting, True, TristateTrue)
     'BM 4/3/2003 - Change this to write ASCII file instead of Unicode by changing TristateTrue to TristateFalse
     Set TextObj = fso.OpenTextFile(FullFilePath, ForWriting, True, TristateFalse)
     ' read the contents of file
     TextObj.Write (Data)
     TextObj.Close
     
     WriteFileContents = ERROR_SUCCESS ' good return
     
     'This is needed to Flush the Windows Cache to the storage device.
     Call FlushFile(FullFilePath)
   End If
   
   LogRegenMessage FullFilePath
    
Function_Exit:
   Set fso = Nothing
   Set TextObj = Nothing

Exit Function

WriteFileContentsErrTrap:
   If gobjShared Is Nothing Then
      Set gobjShared = New CWShared
   End If
   WriteFileContents = ERROR_PIC_DEVICE_TYPE_OR_AVAILABILITY
   
   'Path Not Found
   If Err.Number = 76 Then
     WriteFileContents = ERROR_PIC_DEVICE_PATH_HAS_NOT_BEEN_SET
   Else
     gobjShared.ShowVBError Err.Number, Err.Description, "WriteFileContents Function", "DMBE", vbExclamation
   End If
   Err.Clear
   'Resume Next
   Resume Function_Exit
End Function


Private Function ReadBinaryData(FullFilePath As String) As String
   Dim FileNumber As Long
   Dim Data As String
   Dim DataLen As Long
   
   FileNumber = FreeFile   ' Get unused file
   Open FullFilePath For Binary Access Read As #FileNumber   ' Create file name.
   DataLen = FileLen(FullFilePath)
   Data = InputB(DataLen, #FileNumber)
'   Input #FileNumber, Data
   Close #FileNumber   ' Close file.
   ReadBinaryData = Data
End Function

Private Function PutBinaryData(FullFilePath As String, Data As String) As Long
   Dim FileNumber As Long
   Dim Temp() As Byte
   Dim fs As FileSystemObject
   On Error GoTo ErrHandler
   
   If Data = "" Then
      PutBinaryData = ERROR_NO_DATA_TO_WRITE ' bad return
      Exit Function
   End If
   
   Set fs = New FileSystemObject
   
   ' Begin:SCR #14025;   Developer: MAG 07/17/2002 11:45 AM
   'Need to delete file before Saving Binary Data
   If fs.FileExists(FullFilePath) Then
     fs.DeleteFile FullFilePath, True
   End If
   ' End: SCR #14025;
      
   Set fs = Nothing
   
   Temp = Data
   FileNumber = FreeFile   ' Get unused file
   Open FullFilePath For Binary Access Write As #FileNumber   ' Create file name.
   Put #FileNumber, , Temp
   Close #FileNumber   ' Close file.
   
   Call FlushFile(FullFilePath)
   PutBinaryData = ERROR_SUCCESS ' good return
   
   

Exit Function

ErrHandler:
  Set fs = Nothing
  MsgBox "Application Error: " & Err.Description & " Occurred in DMBE.DMBE_Ops.PutBinaryData."
End Function

Private Function FlushFile(sFileName As String) As Boolean
Dim hHandle As Long
Dim tStruct As OFSTRUCT
On Error GoTo ErrHandler

   hHandle = OpenFile(sFileName, tStruct, OF_EXIST)
   If hHandle <> -1 Then
     Call FlushFileBuffers(hHandle)
     Call CloseHandle(hHandle)
   End If

Exit Function

ErrHandler:
    'Do not raise error, if flush does not occur.
    'Err.Clear
End Function

Private Function GetNewGUID() As String

    On Error GoTo ErrHnd
    
    Dim objGUIDFactory As GUIDFactory
    
    Set objGUIDFactory = New GUIDFactory
    
    GetNewGUID = objGUIDFactory.CreateGUID
    
Exit Function

ErrHnd:
    MsgBox "Application Error: " & Err.Description & " Occurred in DMBE.DMBE_Ops.GetNewGUID"
    
End Function

Private Function FieldExist(oRS As ADODB.Recordset, sFieldName As String) As Boolean

On Error Resume Next

  If Len(oRS(sFieldName).Name) > 0 Then
    FieldExist = True
  End If
End Function

Public Function TMIPFrameworkFlag() As Integer
    
    Dim objReg As RegObj.Registry
    Dim objKey As RegObj.RegKey
    Dim strKeyName As String
    Dim objValue    As RegObj.RegValue
    Dim sValue As String
    
    On Error GoTo ErrHandler
    
    Set objReg = New RegObj.Registry
    
    'SF 1-8-05
    If gobjShared Is Nothing Then
        Set gobjShared = New CWShared
    End If
    
    '<< SCR #29068;   Developer: Brian Mowbray 02/04/2004 09:19 PM
    strKeyName = gobjShared.RegistryKey & "\EventServer"
    
    Set objReg = New RegObj.Registry
    Set objKey = objReg.RegKeyFromString(strKeyName)
    
    'Set the default to 0 if the value is not in the registry
    TMIPFrameworkFlag = 0
    
    If Not objKey Is Nothing Then
        For Each objValue In objKey.Values
          'Default Value
          If UCase$(objValue.Name) = "TMIP_FRAMEWORK" Then
            sValue = UCase$(Trim$(objValue.Value))
            Select Case sValue
              Case "1"
                TMIPFrameworkFlag = 1
              Case "2" '
                TMIPFrameworkFlag = 2
              Case Else
                TMIPFrameworkFlag = 0
            End Select
            Exit For
          End If
          Set objValue = Nothing
        Next objValue
    End If
    Set objReg = Nothing
    Set objKey = Nothing
    
Exit Function

ErrHandler:
    MsgBox "Application Error: " & Err.Description & " Occurred in DMBE.DMBE_Ops.TMIPFrameworkFlag. You may not have the resgistry setting: " & strKeyName
    TMIPFrameworkFlag = 0
    Exit Function
    
End Function

Private Function CheckEFFolderStructure() As Boolean

On Error GoTo ErrHandler

Dim sRootDirectory As String
Dim sEICRegPath As String

    CheckEFFolderStructure = False

      sRootDirectory = GetRootDirectory
      
      Call CreateDataManagerRegistry
      Call SaveRegistryPath("BMIST_195b_OUT", sRootDirectory & "\TMIPB2\EF\Components\C2T\BMIST\Out")
      Call SaveRegistryPath("BMIST_196b_OUT", sRootDirectory & "\TMIPB2\EF\Components\C2T\BMIST\Out")
      Call SaveRegistryPath("BMIST_195b_IN", sRootDirectory & "\TMIPB2\EF\Components\C2T\BMIST\In")
      Call SaveRegistryPath("BMIST_196b_IN", sRootDirectory & "\TMIPB2\EF\Components\C2T\BMIST\In")
      Call SaveRegistryPath("BMIST_OUT", sRootDirectory & "\TMIPB2\EF\Components\C2T\BMIST\Out")
      Call SaveRegistryPath("BMIST_IN", sRootDirectory & "\TMIPB2\EF\Components\C2T\BMIST\In")
      Call SaveRegistryPath("PIC_IN", sRootDirectory & "\TMIPB2\EF\Components\C2T\PIC\In")
      Call SaveRegistryPath("PIC_OUT", sRootDirectory & "\TMIPB2\EF\Components\C2T\PIC\Out")
      Call SaveRegistryPath("DISK_IN", sRootDirectory & "\TMIPB2\EF\Components\C2T\DISK\In")
      Call SaveRegistryPath("DISK_OUT", sRootDirectory & "\TMIPB2\EF\Components\C2T\DISK\Out")
      Call SaveRegistryPath("JMEWS_OUT", sRootDirectory & "\TMIPB2\EF\Components\C2T\JMEWS\Out")
      Call SaveRegistryPath("JMEWSII_OUT", sRootDirectory & "\TMIPB2\EF\Components\C2T\JMEWSII\Out")
      Call SaveRegistryPath("C2T_IN", sRootDirectory & "\TMIPB2\EF\Components\C2T\PVF\In")
      Call SaveRegistryPath("C2T_OUT", sRootDirectory & "\TMIPB2\EF\Components\C2T\PVF\Out")
      Call SaveRegistryPath("TMIP_OUT", sRootDirectory & "\TMIPB2\EF\Components\C2T\TMIP\Out")
      Call SaveRegistryPath("C2TRep_OUT", sRootDirectory & "\TMIPB2\EF\Components\C2T\C2TRep\Out")
      Call SaveRegistryPath("CDR_OUT", sRootDirectory & "\TMIPB2\EF\Components\C2T\CDR\Out")
      Call SaveRegistryPath("EIC_IN_OUT", sRootDirectory & "\TMIPB2\EF\Components\C2T\EIC\In_Out")
      
      '****************************************************
      sEICRegPath = GetRegValue("EIC_IN_OUT")
      If Trim$(sEICRegPath) = vbNullString Then
        sEICRegPath = "TMIP\EIC\EICPath"
      End If
      'This will set the value of "TMIP\EIC\EICPath"
      Call SaveRegistryEICPath(sEICRegPath)
      '****************************************************

    CheckEFFolderStructure = True
    Exit Function

ErrHandler:
    CheckEFFolderStructure = False
End Function


'We are assuming that "\HKEY_LOCAL_MACHINE\SOFTWARE" already exists=
'This Procedure will Create the Registry Setting:
'     "\HKEY_LOCAL_MACHINE\SOFTWARE\TMIP\EIC\EICPath = sEICPath

Private Sub SaveRegistryEICPath(ByVal sEICPath As String)

On Error GoTo ErrHandler

Dim objReg As RegObj.Registry
Dim ParentKey As RegObj.RegKey
Dim ChildKey As RegObj.RegKey
Dim objValue As RegObj.RegValue

Dim bKeyFound As Boolean

Dim sRegPath As String
Dim sChildKey As String

Dim sErrMsg As String

sErrMsg = vbNullString

  Set objReg = New RegObj.Registry
  
  '*****************************************************
  '******Look for TMIP**********************************
  '*****************************************************
  sChildKey = "TMIP"
  bKeyFound = False
  sRegPath = "\HKEY_LOCAL_MACHINE\SOFTWARE\"
  Set ParentKey = objReg.RegKeyFromString(sRegPath)
  For Each ChildKey In ParentKey.SubKeys
    If UCase$(Trim$(ChildKey.Name)) = sChildKey Then
      bKeyFound = True
      Exit For
    End If
  Next ChildKey
  
  If bKeyFound = False Then
    sErrMsg = "Could not add Registry Key 'TMIP'."
    Call ParentKey.SubKeys.Add("TMIP")
    
  End If
  
  '*****************************************************
  '******Look for EIC***********************************
  '*****************************************************
  sChildKey = "EIC"
  bKeyFound = False
  sRegPath = "\HKEY_LOCAL_MACHINE\SOFTWARE\TMIP"
  Set ParentKey = objReg.RegKeyFromString(sRegPath)
  For Each ChildKey In ParentKey.SubKeys
    If UCase$(Trim$(ChildKey.Name)) = sChildKey Then
      bKeyFound = True
      Exit For
    End If
  Next ChildKey
  
  If bKeyFound = False Then
    sErrMsg = "Could not add Registry Key 'EIC'."
    Call ParentKey.SubKeys.Add("EIC")
  End If
  
  '*****************************************************
  '******Look for EICPath*******************************
  '*****************************************************
  sChildKey = "EICPath"
  bKeyFound = False
  sRegPath = "\HKEY_LOCAL_MACHINE\SOFTWARE\TMIP\EIC"
  Set ParentKey = objReg.RegKeyFromString(sRegPath)
  For Each objValue In ParentKey.Values
    If UCase$(Trim$(objValue.Name)) = "EICPATH" Then
      bKeyFound = True
      Exit For
    End If
  Next objValue
  
  If bKeyFound = False Then
    sErrMsg = "Could not add Registry Key 'EICPath'."
    Call ParentKey.Values.Add("EICPath", sEICPath)
  Else
    
    If UCase$(Trim$(objValue.Value)) <> UCase$(Trim$(sEICPath)) Then
      sErrMsg = "Could not set Registry Key 'EICPath' to " & sEICPath & "."
      objValue.Value = sEICPath
    End If
    
  End If
  
  GoTo Cleanup

ErrHandler:

  MsgBox sErrMsg & vbCrLf & Err.Description, vbInformation, "Data Manager"

Cleanup:

  Set objReg = Nothing
  Set ParentKey = Nothing
  Set ChildKey = Nothing
  Set objValue = Nothing

End Sub

Private Function GetRegValue(ByVal sRegValueName As String) As String

On Error GoTo ErrHandler

  Dim objReg As RegObj.Registry
  Dim objKey As RegObj.RegKey
  Dim objValue As RegObj.RegValue
  Dim sRegKey As String
  
  GetRegValue = vbNullString
  
  sRegKey = gobjShared.RegistryKey & "\DataManager"
  
  Set objReg = New RegObj.Registry
  Set objKey = objReg.RegKeyFromString(sRegKey)
  
  For Each objValue In objKey.Values
    If UCase$(Trim$(objValue.Name)) = UCase$(Trim$(sRegValueName)) Then
      GetRegValue = objValue.Value
      Exit For
    End If
  Next objValue
  
  GoTo Cleanup
  
ErrHandler:
    
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "DMBE_Ops.GetRegValue", "Data Manager", vbExclamation)
  
Cleanup:
  
  Set objReg = Nothing
  Set objKey = Nothing
  Set objValue = Nothing

End Function

Private Function GetRootDirectory() As String

  Dim sRegPath As String
  
  If gobjShared Is Nothing Then
    Set gobjShared = New CWShared
  End If

  sRegPath = gobjShared.RTDataPath
  sRegPath = Trim$(sRegPath)
  
  If Len(sRegPath) >= 2 Then
    GetRootDirectory = Left(sRegPath, 2)
  Else
    GetRootDirectory = "C:"
  End If

End Function

'This will Create the DataManager Registry Key if it does not exist.
Private Function CreateDataManagerRegistry() As Boolean

On Error GoTo ErrHandler

    Dim sTemp() As String

    Dim objReg      As RegObj.Registry
    Dim ParentKey   As RegObj.RegKey
    Dim strKeyName  As String
    Dim ChildKey    As RegObj.RegKey
    Dim lLong As Long
    
    strKeyName = gobjShared.RegistryKey
    Set objReg = New RegObj.Registry
    Set ParentKey = objReg.RegKeyFromString(strKeyName)

    CreateDataManagerRegistry = False

    For Each ChildKey In ParentKey.SubKeys
      If UCase$(Trim$(ChildKey.Name)) = "DATAMANAGER" Then
        CreateDataManagerRegistry = True
        Exit Function
      End If
    Next ChildKey
    
    Call ParentKey.SubKeys.Add("DataManager")
    CreateDataManagerRegistry = True

Exit Function
ErrHandler:

  MsgBox "Could not create the Registry Key: DATAMANAGER" & vbCrLf & Err.Description, vbOKOnly, "Data Manager"
  CreateDataManagerRegistry = False

End Function

'If Registry Entry and Folder does not exist, then Build it.
Private Sub SaveRegistryPath(ByVal RegKeyName As String, ByVal sDefaultFolder As String)
    
    Dim sTemp() As String
    Dim objReg As RegObj.Registry
    Dim objKey   As RegObj.RegKey
    Dim objValue As RegObj.RegValue
    Dim sExistingFolder As String
    
    Dim sMsg As String

    Dim nRetValue As Integer
    
    'call ParentKey.SubKeys(2).Values.Add("Hello", "0")
    
    On Error GoTo ErrHandler
    
    'sFolder = vbNullString
    sMsg = vbNullString
    sExistingFolder = vbNullString
    
    If RegistryValueExists(RegKeyName, sExistingFolder, sDefaultFolder) = False Then
        
      'This will Build the Directories
      nRetValue = CheckDirs(RegKeyName, sExistingFolder, sDefaultFolder)
      
      Set objReg = New RegObj.Registry
      Set objKey = objReg.RegKeyFromString(gobjShared.RegistryKey & "\DataManager")
      
      sMsg = "Cannot create Registry Value " & sDefaultFolder & " for registry key " & RegKeyName & "."
      Call objKey.Values.Add(UCase$(RegKeyName), sDefaultFolder)
    
    Else
    
      If sExistingFolder = vbNullString Then
      
        'Modify the Existing Registry Key with DefaultFolder Path
        Set objReg = New RegObj.Registry
        Set objKey = objReg.RegKeyFromString(gobjShared.RegistryKey & "\DataManager")
        
        If Not objKey Is Nothing Then
          For Each objValue In objKey.Values
          
            If UCase$(Trim$(objValue.Name)) = UCase$(Trim$(RegKeyName)) Then
              
              sMsg = "Cannot create Registry Value " & sDefaultFolder & " for registry key " & RegKeyName & "."
              objValue.Value = sDefaultFolder
              sExistingFolder = sDefaultFolder
            
            End If
        
          Next objValue
        End If
        
      
      End If
      
    
    End If
      
    nRetValue = CheckDirs(RegKeyName, sExistingFolder, sDefaultFolder)
    

GoTo Cleanup

ErrHandler:

  MsgBox sMsg & vbCrLf & Err.Description, vbInformation, "Data Manager"

Cleanup:

    Set objReg = Nothing
    Set objKey = Nothing
    Set objValue = Nothing
    
End Sub

'If Registry Value is there, don't change it.
Private Function RegistryValueExists(ByVal RegKeyName As String, ByRef sExistingFolder As String, ByVal sDefaultFolder As String) As Boolean

On Error GoTo ErrHandler
    
    Dim reg As Registry
    Dim sTemp() As String
    
    Dim objReg      As RegObj.Registry
    Dim objKey      As RegObj.RegKey
    Dim strKeyName  As String
    Dim objValue    As RegObj.RegValue
    
    Set objReg = New RegObj.Registry
    
    
    strKeyName = gobjShared.RegistryKey & "\DataManager"
    
    Set objKey = objReg.RegKeyFromString(strKeyName)
  
    RegistryValueExists = False
    sExistingFolder = vbNullString
    
    For Each objValue In objKey.Values
      If UCase$(objValue.Name) = UCase$(Trim(RegKeyName)) Then
        
        RegistryValueExists = True
        
        If objValue.Value <> vbNullString Then
          
          sExistingFolder = objValue.Value
          'Check Validity of Folder
          If BuildFolder(sExistingFolder) <> 0 Then
            sExistingFolder = vbNullString
          End If
          Exit Function
        
        End If
      
        Exit For
      End If
          
    Next objValue
    Set objValue = Nothing
    
GoTo Cleanup

ErrHandler:

    RegistryValueExists = False
    sExistingFolder = vbNullString
    
Cleanup:
    
    Set objReg = Nothing
    Set objKey = Nothing
    Set objValue = Nothing

End Function

Private Function CheckDirs(ByVal sRegKey As String, ByVal sExistingFolder As String, ByVal sDefaultFolder As String) As Long
   
   Dim fso As FileSystemObject
   Dim drv As Variant
   Dim dc As Variant
   Dim DriveStr As String
   Dim nRetValue As Integer
   Dim sPath As String
   
   On Error GoTo CheckOrBuildDirsError
   
   CheckDirs = CheckOrBuildDirs(sExistingFolder, sDefaultFolder)
   
   
      
  Exit Function
   
CheckOrBuildDirsError:
   
   gobjShared.ShowVBError Err.Number, Err.Description, "CheckDirs", "DMFE", vbExclamation
   CheckDirs = -1
   Exit Function

End Function

Private Sub CreateDirs(FilePath As String)
   On Error Resume Next
   
   MkDir FilePath
End Sub

'Check if the Exisitng Folder is legitimate/
'  If yes, then Build it.
'  If no, then Build the Default Folder.

Private Function CheckOrBuildDirs(ByVal sExistingFolder As String, ByVal sDefaultFolder As String) As Integer
   
   On Error GoTo ErrHandler
   
   CheckOrBuildDirs = 0
      
      If sExistingFolder <> "" Then
        CheckOrBuildDirs = BuildFolder(sExistingFolder)
        If CheckOrBuildDirs <> 0 Then
          CheckOrBuildDirs = BuildFolder(sDefaultFolder)
        End If
      Else
        CheckOrBuildDirs = BuildFolder(sDefaultFolder)
      End If
      
        
Exit Function
   
ErrHandler:

  gobjShared.ShowVBError Err.Number, Err.Description, "CheckOrBuildDirs", "DMFE", vbExclamation
  CheckOrBuildDirs = -1
  Exit Function

End Function

'This will Check or Build the Folder Path
Private Function BuildFolder(ByVal sFolder As String) As Long

On Error GoTo ErrHandler
   
   Dim Temp() As String
   Dim LB As Long
   Dim UB As Long
   Dim I As Long
   Dim sPath As String
   Dim fso         As FileSystemObject

   
      If sFolder <> vbNullString Then

         Temp = Split(sFolder, "\", -1, vbTextCompare)
         LB = LBound(Temp)
         UB = UBound(Temp)
         Set fso = New FileSystemObject
         If InStr(1, Temp(I), ":", vbTextCompare) <> 0 Then
            If fso.DriveExists(Temp(I) & "\") Then
               sPath = Temp(I)
            Else
               BuildFolder = -100 ' drive does not exist - bad return
               Set fso = Nothing
               Exit Function
            End If
         Else
            BuildFolder = -200 ' drive is not at start of path - bad return
            Set fso = Nothing
            Exit Function
         End If
         
         For I = LB + 1 To UB
            sPath = sPath & "\" & Temp(I)
            If fso.FolderExists(sPath) = False Then
               BuildFolder = -1
               Call fso.CreateFolder(sPath)
               BuildFolder = 0  'Good Return
            Else
                BuildFolder = 0
            End If
         Next I

      Else
         BuildFolder = -200 ' no path specified - bad return
         Set fso = Nothing
         Exit Function
      End If
      
      Set fso = Nothing

Exit Function

ErrHandler:

  MsgBox "Folder " & sFolder & " could not be created." & vbCrLf & Err.Description, vbOKOnly, "Data Manager"
  
  BuildFolder = -1
  Set fso = Nothing
  Exit Function

End Function


'--- SCR 85597   Sherry Wang   1/25/2006
Private Function GetEncryptionKey() As String
    On Error GoTo ErrHandler
    
    Dim sKey As String
    Dim oCrypto As C2T_Security.CCrypto
    Const sDefaultKey = "07AC44E13C503AD56AB59CCABFBF943AC1913212A79F190331DD41729F26513B0F198FE2BB548D66D79FF226B67B1BDB"

    sKey = GetRegValue("PWDx")

    If LenB(sKey) = 0 Then
        Call SaveRegistryPath("PWDx", sDefaultKey)
        sKey = sDefaultKey
    End If
    
    Set oCrypto = New C2T_Security.CCrypto
    GetEncryptionKey = oCrypto.Decrypt(sKey, oCrypto.GetHash(AppName & gsExp, True), True, frezBlockEncryption)

    Exit Function
ErrHandler:
    MsgBox "Application Error: " & Err.Description & " Occurred in DMBE.DMBE_Ops.GetEncryptionKey"
End Function


Private Function GetDiagnosisExtenderName(ByVal extenderCode As String) As String
    Dim dodCodesForm As DodCentricCodes

    On Error Resume Next
  
    If Len(extenderCode) < 3 Then Exit Function
  
    Set dodCodesForm = New DodCentricCodes
    dodCodesForm.Visible = False
  
    GetDiagnosisExtenderName = dodCodesForm.GetDiagnosisExtenderName(extenderCode)
  
    Unload dodCodesForm
    Set dodCodesForm = Nothing
    
End Function


