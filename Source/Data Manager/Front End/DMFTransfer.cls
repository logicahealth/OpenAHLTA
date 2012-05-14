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
Attribute VB_Name = "DMFTransferManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' SCR #44079;   Developer: Mag 10/20/2003 01:11 PM
' Added sPatientFMP & sPatientSponsorSSN
Public Type DMFTransferResult
    sPatientUnitNumber    As String
    sPatientSSN           As String
    sPatientFMP           As String
    sPatientSponsorSSN    As String
    sPatientName          As String
    Source                As Long           'DMLocations
    Destination           As Long           'DMLocations
    Dataset               As Long           'DMDataSets
    Result                As Long           'DMErrorCodes
    sErrorDescription     As String
End Type

Private Const GETRESULTMEMORY As Integer = 1000
Private m_objDMFData            As DMFTransactionData
Private m_tpTransferResult()    As DMFTransferResult
Private m_objDMBE               As IDMBE.I_DMBE
Private m_DM                    As Object
Private m_nRecordIndicator      As Long

Public Function ExecuteTransfer(ByVal objDMFTD As DMFTransactionData, _
    Optional bResendTMIPMessages As Boolean = False, _
    Optional RegeneratePVFFiles As Boolean = True, _
    Optional FullGrid As Boolean = False) As Variant
            
    'The return variant is m_tpTransferResult
    
    Dim lDataSets() As Long
    Dim myPatient(0) As DMFPatientInfo  'There will only be one patient in the transfer
    Dim TransferResultState() As DMFTransferResult
    Dim mEncSignedState As Long
    Dim ImportTransferResultState() As DMFTransferResult
    
    On Error GoTo ExecuteTransferError
    
    Set m_objDMFData = objDMFTD
    
    
    
    'Initialize all the module level variables
    Set m_objDMBE = New DMBE.DMBE_Ops
    
    '******************
    'Set the patient property and logon property
    Set m_objDMBE.PatientObj = cwobjPatient
    Set m_objDMBE.LogonObj = cwobjLogonEnviron
    
    'cwobjPatient
    'cwobjLogonEnviron
    
    '**************************
    
    m_nRecordIndicator = 0
    
    m_objDMBE.ClearParameters
    
    '*-- Set DMBE PatientObj to point to cwobjPatient just one time.
    'Set m_objDMBE.PatientObj = LoadPatient(m_objDMFData.PatientInfo(0).sUnitNumber)
    'Set m_objDMBE.PatientObj = cwobjPatient
    
    If m_objDMFData.Destination = eDM_EIC Then
        'm_objDMFData.Destination = eDM_TMIP
        m_objDMFData.IsEICDestination = True
    End If
    
    m_objDMBE.Destination = m_objDMFData.Destination
    m_objDMBE.Source = m_objDMFData.Source
    m_objDMBE.Format = m_objDMFData.Format
    m_objDMBE.Crypto = m_objDMFData.Encryption
    m_objDMBE.Compression = m_objDMFData.Compression
    m_objDMBE.PicPath = m_objDMFData.SourcePath
    m_objDMBE.AltPicPath = m_objDMFData.DestinationPath
    m_objDMBE.EICIsDestination = m_objDMFData.IsEICDestination
    mEncSignedState = m_objDMFData.EncounterSource
    m_objDMBE.EncounterSource = mEncSignedState
    
    'Determining is the transfer type
    
    'PIC To PIC
    If (m_objDMBE.Source = eDM_PIC Or m_objDMBE.Source = eDM_AGGPIC) _
        And (m_objDMBE.Destination = eDM_PIC Or m_objDMBE.Destination = eDM_AGGPIC) Then
        
        Call PICtoPICTransfer
        
    ElseIf m_objDMBE.Source = eDM_LDDB And m_objDMBE.Destination = 8 Then 'eDM_JMEWS SF 62831
        Call ValidateCriteriaCDRAndLDDBAsSource(bResendTMIPMessages)
        
    'CDR Or LDDB is the source
    ElseIf m_objDMBE.Source = eDM_CDR Or m_objDMBE.Source = eDM_LDDB Then
        
        'SCR 84451 and 84452
        If bResendTMIPMessages Then
            Dim oDMBEOps_Regen As DMBE.DMBE_Ops
            Set oDMBEOps_Regen = m_objDMBE
            oDMBEOps_Regen.DeleteRegenMessageLog
        End If
        
        Call ValidateCriteriaCDRAndLDDBAsSource(bResendTMIPMessages)
        
        If GetPVFFlag And m_objDMBE.Destination = eDM_TMIP Then
        
          TransferResultState = m_tpTransferResult
    
          'PIC PVF File Transfer
          '=====================
          If RegeneratePVFFiles Then
              Call SendPVF(mEncSignedState, bResendTMIPMessages)
          End If
         
          m_tpTransferResult = TransferResultState
        
        End If 'PVF Flag

    ' PIC or AGGPIC is the Source of data: DESTINATION: CDR Or LDDB
    ElseIf (m_objDMBE.Source = eDM_PIC Or m_objDMBE.Source = eDM_AGGPIC) _
       And (m_objDMBE.Destination = eDM_CDR Or m_objDMBE.Destination = eDM_LDDB) Then
        
        Call PICtoCDROrLDDBTransfer
        
    '<< SCR #29291;   Developer: Brian Mowbray 11/14/2002 12:30 PM
    ElseIf m_objDMBE.Source = eDM_EIC And m_objDMBE.Destination = eDM_LDDB Then
        
        Call EICtoLDDBTransfer
    '<< Begin:SCR #29043;   Developer: Brian Mowbray 11/18/2002 12:05 PM
    ElseIf m_objDMBE.Source = eDM_BMIST And m_objDMBE.Destination = eDM_LDDB Then
    '>> End: SCR #29043;
        
      
      Call EICtoLDDBTransfer
      
      'Save the result temporarily
      ImportTransferResultState = m_tpTransferResult
      
      '<< Begin:SCR #29597;   Developer: Brian Mowbray 11/21/2002 03:44 PM
'      If m_tpTransferResult(1).Result = ERROR_SUCCESS Then
          

          '<< Begin:SCR #29827;   Developer: Brian Mowbray 11/26/2002 05:59 PM
          If bResendTMIPMessages = True Then
                m_objDMFData.PatientInfo = GetAllPatientsFromDB
          Else
                m_objDMFData.PatientInfo = GetSelectedPatientFromGrid(FullGrid)
          End If
          
          '>> End: SCR #29827;
      
          m_nRecordIndicator = 0
    
          m_objDMBE.ClearParameters
          
          ReDim lDataSets(1)
          With m_objDMFData
            lDataSets(0) = NCID_DEMOGRAPHICS_APPLICATION
            lDataSets(1) = NCID_ENCOUNTER_APPLICATION
          End With
    
          m_objDMFData.Destination = eDM_TMIP
          m_objDMFData.Source = eDM_LDDB
          m_objDMFData.SourcePath = ""
          m_objDMFData.DestinationPath = ""
          m_objDMFData.IsEICDestination = False
    
          m_objDMBE.Destination = eDM_TMIP
          m_objDMBE.Source = eDM_LDDB
          m_objDMBE.Format = eDM_XML
          m_objDMBE.Crypto = eDM_NoCrypto
          m_objDMBE.Compression = eDM_NoCompression
          m_objDMBE.PicPath = ""
          m_objDMBE.AltPicPath = ""
          m_objDMBE.EICIsDestination = False
    
          If Not gobjComm Is Nothing Then
              Call gobjComm.SetStatusText("Sending Encounter to TMIP ...")
          End If
          
          Call ValidateCriteriaCDRAndLDDBAsSource(bResendTMIPMessages)

        If GetPVFFlag And m_objDMBE.Destination = eDM_TMIP Then
        
          TransferResultState = m_tpTransferResult
    
          'PIC PVF File Transfer
          '=====================
          Call SendPVF(mEncSignedState, bResendTMIPMessages)
         
          m_tpTransferResult = TransferResultState
        
        End If 'PVF Flag
        
        'Set the transfer result from the import, not from the message generation.
        m_tpTransferResult = ImportTransferResultState
    End If
      '>> End: SCR #29597;
    
    Call SetCWStatusText("")
    
    'Set return equal to the module level variable
    ExecuteTransfer = m_tpTransferResult
    
    '<SCR 22675 CC>
    'Forgot to set the destination back
    If m_objDMBE.EICIsDestination = True Then
        m_objDMBE.Destination = eDM_EIC
    End If
    '</SCR 22675 CC>
    
    Set m_objDMFData = Nothing
    Set m_objDMBE = Nothing
    m_nRecordIndicator = 0
    
Exit Function

ExecuteTransferError:
    Set m_objDMFData = Nothing
    Set m_objDMBE = Nothing
    m_nRecordIndicator = 0
    Screen.MousePointer = vbDefault
    gobjShared.ShowVBError Err.Number, Err.Description, "ExecuteTransfer", "DMF TransferManager", vbCritical
    Err.Clear
    Call SetCWStatusText("")
    Exit Function
    Resume
End Function

'<SCR 20841 CC>
'SCR 71415 - Import Allergy data into LDDB from EIC File.
Private Function EICtoLDDBTransfer()

    Dim objEICTransfer As DMBE.HL7XMLParser
    Dim colRSObjects As Collection
    Dim nTemp As Long
    Dim transItem As IRSTransfer.I_RSTransfer
    Dim lngResult As DMErrorCodes
    Dim lngUnitNumber As Long
    Dim strTemp As String
    Dim blnDemographicsFound As Boolean
    Dim blnEncountersFound As Boolean
    Dim blnAllergiesFound As Boolean        'SCR 71415
    Dim oNode As MSXML.IXMLDOMNode
    Dim nPatInfo As Long
    Dim sPtCnt As String
    Dim objAllergy As Object
    
    Dim enmWhoseMessage As XMLMessageSource
    
    On Error GoTo ErrHandler

    ReDim m_tpTransferResult(0)
    
    Set objEICTransfer = New DMBE.HL7XMLParser
    Set m_objDMBE.IDMappers = New DMBE.DMBE_IDMappers
    
    For nPatInfo = 0 To UBound(m_objDMFData.PatientInfo) 'Loop based on number of patient selected on the RegView Screen
      
      '<< Begin:SCR #30824;   Developer: Brian Mowbray 12/16/2002 01:32 PM
      ' The patients should be parsed in the same order as processed when the xml is parsed initially
      '>> End: SCR #30824;
      
      lngUnitNumber = m_objDMFData.PatientInfo(nPatInfo).sUnitNumber
      '<< Begin:SCR #30824;   Developer: Brian Mowbray 12/16/2002 11:51 PM
      sPtCnt = m_objDMFData.PatientInfo(nPatInfo).sId
      '>> End: SCR #30824;
      
      m_tpTransferResult(0).sPatientUnitNumber = lngUnitNumber
      m_tpTransferResult(0).sPatientName = m_objDMFData.PatientInfo(nPatInfo).sName
      m_tpTransferResult(0).sPatientSSN = m_objDMFData.PatientInfo(nPatInfo).sSSN
      m_tpTransferResult(0).Source = m_objDMFData.Source
      m_tpTransferResult(0).Destination = m_objDMFData.Destination
      m_tpTransferResult(0).Dataset = NCID_DEMOGRAPHICS_APPLICATION
      
      
      If m_objDMFData.Source = eDM_BMIST Then
      
        Dim strXMLDec As String
        Dim strDTDDec As String
        Dim oDom As MSXML.DOMDocument
        Dim strXMLData() As String
        Dim sKey As String
      
        strXMLDec = "<?xml version" '""1.0""?>"
        strDTDDec = "<!DOCTYPE" ' ADT_A31 SYSTEM ""hl7_v231.dtd"">"
        '<< Begin:SCR #30824;   Developer: Brian Mowbray 12/16/2002 01:32 PM
        sKey = m_objDMFData.PatientInfo(nPatInfo).sSSN & CStr(lngUnitNumber) & sPtCnt
        '>> End: SCR #30824;
      
        'This next line is causing an error
        strXMLData = Split(g_dicBMISTFile(sKey), strXMLDec)
        
        Set g_EICMessage.DemographicsDOM = New MSXML.DOMDocument
        Set g_EICMessage.EncountersDOM = New MSXML.DOMDocument
        
        '*****************************************************
        'SCR 71415
        Set g_EICMessage.AllergiesDOM = New MSXML.DOMDocument
        '*****************************************************
        
        
    
        For nTemp = 0 To UBound(strXMLData)
        
            If strXMLData(nTemp) <> "" Then
              
              strXMLData(nTemp) = Replace(strXMLData(nTemp), "encoding = 'UTF-8'", "")
              strXMLData(nTemp) = Replace(strXMLData(nTemp), "encoding='UTF-8'", "")
              strXMLData(nTemp) = Replace(strXMLData(nTemp), "encoding= 'UTF-8'", "")
              strXMLData(nTemp) = Replace(strXMLData(nTemp), "encoding ='UTF-8'", "")
              
              'If blnXMLDeclarationFound = True Then
                  strXMLData(nTemp) = strXMLDec & strXMLData(nTemp)
              'End If
              
              Set oDom = New MSXML.DOMDocument
              'Disable DTD Validation
              oDom.async = False
              oDom.validateOnParse = False
              oDom.resolveExternals = False
                      
              If oDom.loadXML(strXMLData(nTemp)) = True Then
                  
                  Set oNode = oDom.selectSingleNode("//MSH.9/CM_MSG_TYPE.2")
                  If Not oNode Is Nothing Then
                      Select Case UCase(Trim(oNode.text))
                          Case Is = "A28"
                              Set g_EICMessage.DemographicsDOM = oDom.cloneNode(True)
                              'g_EICMessage.DemographicsMessage = strXMLData(nTemp)
                          Case Is = "A04"
                              Set g_EICMessage.EncountersDOM = oDom.cloneNode(True)
                      End Select
                  End If 'oNode is Nothing
              End If
            End If
          Next nTemp
      End If
      
      
      'EIC Starts Here
      
      'Just a sanity check
      For nTemp = 0 To UBound(m_objDMFData.Dataset)
          If m_objDMFData.Dataset(nTemp) = NCID_DEMOGRAPHICS_APPLICATION Then
              blnDemographicsFound = True
              Exit For
          End If
      Next nTemp
      
      If blnDemographicsFound = False Then
          m_tpTransferResult(0).Result = DMErrorCodes.ERROR_EIC_INCORRECT_FORMAT
          Exit Function
      End If
      
      'Check where the message was sent from
      Set oNode = g_EICMessage.DemographicsDOM.selectSingleNode("//MSH.3/HD.1")
      If Not oNode Is Nothing Then
          strTemp = UCase(oNode.text)
          Select Case strTemp
              Case "SAMS"
                  lngUnitNumber = CLng(Date) + Hour(Time) + Minute(Time) + Second(Time)
                  enmWhoseMessage = SAMS_Message
              '<< Begin:SCR #29291;   Developer: Brian Mowbray 11/14/2002 12:31 PM
              Case "AHLTA-MOBILE", "BMIST"
                  If lngUnitNumber = 0 Then
                    lngUnitNumber = CLng(Date) + Hour(Time) + Minute(Time) + Second(Time)
                  End If
                  enmWhoseMessage = BMIST_Message
              '>> End: SCR #29291;
              Case "DMU"
                  enmWhoseMessage = CHCSII_Message
              Case Else
                  enmWhoseMessage = SAMS_Message
                  'lngUnitNumber = CLng(Date) + Hour(Time) + Minute(Time) + Second(Time)
          End Select
      Else
          m_tpTransferResult(0).Result = DMErrorCodes.ERROR_EIC_INCORRECT_FORMAT
          Exit Function
      End If 'oNode is nothing
  
      'Validate patient first
      
      lngResult = objEICTransfer.GetRecordsetsFromEIC(colRSObjects, CLng(0), setDatasetEnum(NCID_DEMOGRAPHICS_APPLICATION), g_EICMessage.DemographicsDOM.xml, enmWhoseMessage)
      
      Dim enmPatientSearchResult As eSearchResults
  
      Call cwobjPatient.Validation_Criteria_Eic(colRSObjects)
      
      'colRSObjects is holding ONLY Dem data
  
      '-- 188795 B. Webb different search criteria for data coming from AHLTA Mobile (No Patient SSN)
      If m_objDMBE.Source = eDM_BMIST And m_objDMBE.Destination = eDM_LDDB And cwobjPatient.Validation_Criteria.PatientFMP <> "20" Then
        enmPatientSearchResult = cwobjPatient.ValidatePatient(True, True)
      Else
        enmPatientSearchResult = cwobjPatient.ValidatePatient(True)
      End If
      
      Select Case enmPatientSearchResult
          Case esr_MatchedSinglePatient, esr_CreatedNewPatient, esr_MatchedNameSSN
              'The only time want to select the patient
              'This will prevent the patient from being loaded again
              If Trim$(m_objDMFData.PatientInfo(nPatInfo).sUnitNumber) <> Trim(cwobjPatient.UnitNumber) Then
                m_objDMFData.PatientInfo(nPatInfo).sUnitNumber = cwobjPatient.UnitNumber
                UpdateSearchGridUnit m_objDMFData.PatientInfo(nPatInfo).sUnitNumber, cwobjPatient.UnitNumber
              End If
              
              Call ChangeCorePatient(cwiSELECT_PATIENT, cwobjPatient.UnitNumber & ";" & cwobjPatient.MedicalRecordNumber)
          
          Case Else
              'Clear the patient
              Call ChangeCorePatient(cwiCLEAR_PATIENT, lngUnitNumber & ";" & "")
              lngResult = DMErrorCodes.ERROR_CANNOT_VALIDATE_PATIENT
              '-- raise some error condition here?
              m_tpTransferResult(0).Result = lngResult
          Exit Function
      End Select
      
      'Update the Unitnumber
      
      '********************************************************
      'Encounter transfer from now on
      '********************************************************
      
      Set m_objDMBE.PatientObj = cwobjPatient
      
      m_nRecordIndicator = 1
      
      For nTemp = 0 To UBound(m_objDMFData.Dataset)
          If m_objDMFData.Dataset(nTemp) = NCID_ENCOUNTER_APPLICATION Then
              blnEncountersFound = True
              Exit For
          End If
      Next nTemp
      
      If blnEncountersFound = False Then
          Exit Function
      End If
      
      ReDim Preserve m_tpTransferResult(1)
          
      m_tpTransferResult(1).sPatientUnitNumber = cwobjPatient.UnitNumber
      m_tpTransferResult(1).sPatientName = m_objDMFData.PatientInfo(nPatInfo).sName
      m_tpTransferResult(1).sPatientSSN = m_objDMFData.PatientInfo(nPatInfo).sSSN
      m_tpTransferResult(1).Source = m_objDMFData.Source
      m_tpTransferResult(1).Destination = m_objDMFData.Destination
      m_tpTransferResult(1).Dataset = NCID_ENCOUNTER_APPLICATION
      m_objDMBE.Dataset = eDM_Encounters
      
      
      If enmWhoseMessage = CHCSII_Message Then
          m_objDMBE.WhoseXMLMessage = CHCSII_Message
          m_objDMBE.EICHL7Message = g_EICMessage.EncountersDOM.xml
          Set transItem = gobjComm.InitializeOLEServer(NCID_ENCOUNTER_APPLICATION)
          If Not (transItem Is Nothing) Then
              lngResult = transItem.Transfer(m_objDMBE)
              Set transItem = Nothing
          Else
              lngResult = ERROR_CANNOT_INSTANTIATE_CLINICAL_MODULES
          End If
      Else
      
          'Need to construct an Encounter on the fly
          '<< Begin:SCR #29291;   Developer: Brian Mowbray 11/14/2002 12:33 PM
          lngResult = objEICTransfer.CreateEncounterFromSAMSEIC(g_EICMessage.EncountersDOM, enmWhoseMessage)
          '>> End: SCR #29291;
  
          Set oNode = g_EICMessage.EncountersDOM.selectSingleNode("//AL1")
          If Not oNode Is Nothing Then
              Dim oZALNode As MSXML.IXMLDOMNode
              Dim rsALLERGIES As ADODB.Recordset
              Dim rsALLERGYREACTION As ADODB.Recordset
              Dim rsALLERGYCOMMENT As ADODB.Recordset
              Dim objRSObject As IDMBE.I_RSObjects
              Set oZALNode = g_EICMessage.EncountersDOM.selectSingleNode("//ZAL")
              If Not oZALNode Is Nothing Then
                              
                  lngResult = objEICTransfer.FillRecordsetFromZALNode(oZALNode, lngUnitNumber, rsALLERGIES, rsALLERGYREACTION, rsALLERGYCOMMENT)
                  
                  If lngResult = ERROR_SUCCESS Then
                      'Put them into a collection and call transfer from PIC to LDDB
                      Set colRSObjects = New Collection
      
                      Set objRSObject = New DMBE.RSObject_Ops
                      objRSObject.TableName = "ALLERGIES"
                      Set objRSObject.DBRecordSet = rsALLERGIES
                      colRSObjects.Add objRSObject
                  
                      Set objRSObject = New DMBE.RSObject_Ops
                      objRSObject.TableName = "ALLERGYREACTION"
                      Set objRSObject.DBRecordSet = rsALLERGYREACTION
                      colRSObjects.Add objRSObject
                  
                      Set objRSObject = New DMBE.RSObject_Ops
                      objRSObject.TableName = "ALLERGYCOMMENT"
                      Set objRSObject.DBRecordSet = rsALLERGYCOMMENT
                      colRSObjects.Add objRSObject
                  
                      
                      Set objAllergy = gobjComm.InitializeOLEServer(NCID_ALLERGIES_APPLICATION)
                      If Not (objAllergy Is Nothing) Then
                          Call objAllergy.AlleryRecordsetTransferToLDDB(colRSObjects)
                      End If
                          
                  End If
              
              Else 'Need to update from AL1
                  
                  Dim strSQL As String
                  Dim oDAS As ICHCSII_DAS
                  Dim oSQLOpsEx As ICHCSII_SQL.ISqlOpsEx
                      
                  lngResult = objEICTransfer.FillRecordsetFromAL1Node(oNode, lngUnitNumber, rsALLERGIES, rsALLERGYREACTION, rsALLERGYCOMMENT)
                  If lngResult = ERROR_SUCCESS Then
                              
                      If Not rsALLERGIES Is Nothing Then
                          If rsALLERGIES.RecordCount > 0 Then
                              
                              Set oDAS = oDatabase.CHCSII_DAS(Auto) 'New GEMS_DAS
                              Set oSQLOpsEx = oDatabase.CHCSII_SQLOPS(Auto) 'New GEMS_SQL.SqlOps
                              
                              oSQLOpsEx.Execute "ALLERGIES"
                              
                              Do Until rsALLERGIES.EOF
                                  oSQLOpsEx.AddNew
                                      oSQLOpsEx("DATAID") = oDAS.GetNextID
                                      oSQLOpsEx("FACILITY_NCID") = cwobjLogonEnviron.FacilityNCID
                                      oSQLOpsEx("UNIT_NUMBER") = lngUnitNumber
                                      oSQLOpsEx("CLINICIANNCID") = cwobjLogonEnviron.UserName
                                      oSQLOpsEx("ALLERGENNCID") = rsALLERGIES.Fields("ALLERGENNCID").Value
                                      oSQLOpsEx("ALLEREN") = rsALLERGIES.Fields("ALLEREN").Value
                                      oSQLOpsEx("NOTEDDATE") = rsALLERGIES.Fields("NOTEDDATE").Value
                                      oSQLOpsEx("CREATEDBY") = cwobjLogonEnviron.UserName
                                      oSQLOpsEx("CREATEDON") = Date
                                  oSQLOpsEx.Update
                                  rsALLERGIES.MoveNext
                              Loop
                              
                              If Not rsALLERGYREACTION Is Nothing Then
                                  If rsALLERGYREACTION.RecordCount > 0 Then
                                      oSQLOpsEx.Execute "ALLERGYREACTION"
                                      Do Until rsALLERGYREACTION.EOF
                                          oSQLOpsEx.AddNew
                                              oSQLOpsEx("UNIT_NUMBER") = lngUnitNumber
                                              oSQLOpsEx("REACTIONNCID") = rsALLERGYREACTION.Fields("REACTIONNCID").Value
                                              oSQLOpsEx("REACTION") = rsALLERGYREACTION.Fields("REACTION").Value
                                              oSQLOpsEx("ALLERENNCID") = rsALLERGYREACTION.Fields("ALLERENNCID").Value
                                              oSQLOpsEx("CREATEDBY") = cwobjLogonEnviron.UserName
                                              oSQLOpsEx("CREATEDON") = Date
                                          oSQLOpsEx.Update
                                          rsALLERGYREACTION.MoveNext
                                      Loop
                                  End If
                              End If 'rsALLERGYREACTION Is Nothing
                          End If 'rsALLERGIES.RecordCount > 0
                      End If 'rsALLERGIES Is Nothing
                  End If 'lngResult = ERROR_SUCCESS
              End If 'oZALNode Is Nothing
          End If 'oNode Is Nothing
      End If 'CHCSII_MESSAGE
      
      If lngResult = ERROR_SUCCESS Then
          If m_objDMBE.CurrentDMErrorCode <> DMErrorCodes.ERROR_SUCCESS Then
              lngResult = m_objDMBE.CurrentDMErrorCode
          End If
      End If
      
      m_tpTransferResult(1).Result = lngResult
      
      '*************************************************************
      'Begin 71415
      'SCR 71415 - Need to Import EIC Allergy Data
      'BFote - 14-Jun-2005
      'Let's Give Allergies a Shot
      blnAllergiesFound = False
      For nTemp = 0 To UBound(m_objDMFData.Dataset)
        If m_objDMFData.Dataset(nTemp) = NCID_ALLERGIES_APPLICATION Then
          blnAllergiesFound = True
          Exit For
        End If
      Next nTemp
      
      '********************************************************
      'SCR 72865
      'If there is no Allergy data in the EIC File, then Leave.
      If g_EICMessage.AllergiesDOM.xml = vbNullString Then
        blnAllergiesFound = False
      End If
      '********************************************************
      
      'SF SCR TBD 'The BMIST file does not have allergies, so do not exit, just skip allergies.
      If blnAllergiesFound = True Then
      
          m_nRecordIndicator = 2
          
          ReDim Preserve m_tpTransferResult(2)
    
          m_tpTransferResult(2).sPatientUnitNumber = cwobjPatient.UnitNumber
          m_tpTransferResult(2).sPatientName = m_objDMFData.PatientInfo(nPatInfo).sName
          m_tpTransferResult(2).sPatientSSN = m_objDMFData.PatientInfo(nPatInfo).sSSN
          m_tpTransferResult(2).Source = m_objDMFData.Source
          m_tpTransferResult(2).Destination = m_objDMFData.Destination
          m_tpTransferResult(2).Dataset = NCID_ALLERGIES_APPLICATION
          m_objDMBE.Dataset = eDM_Allergies
          
          lngUnitNumber = cwobjPatient.UnitNumber
          
          lngResult = objEICTransfer.GetRecordsetsFromEIC(colRSObjects, lngUnitNumber, setDatasetEnum(NCID_ALLERGIES_APPLICATION), g_EICMessage.AllergiesDOM.xml, enmWhoseMessage)
          
          If enmWhoseMessage = CHCSII_Message Then
              m_objDMBE.WhoseXMLMessage = CHCSII_Message
              m_objDMBE.EICHL7Message = g_EICMessage.AllergiesDOM.xml
              
              
              Set objAllergy = gobjComm.InitializeOLEServer(NCID_ALLERGIES_APPLICATION)
              If Not (objAllergy Is Nothing) Then
                lngResult = objAllergy.AlleryRecordsetTransferToLDDB(colRSObjects)
              End If
            
          End If
          
          If lngResult = ERROR_SUCCESS Then
              If m_objDMBE.CurrentDMErrorCode <> DMErrorCodes.ERROR_SUCCESS Then
                  lngResult = m_objDMBE.CurrentDMErrorCode
              End If
          End If
          
          m_tpTransferResult(2).Result = lngResult
          
          'END 71415
          '*************************************************************
    End If
    ' 85915 Quan 5/11/07
    Call ChangeCorePatient(cwiREFRESH_ALLERGIES, cwobjPatient.UnitNumber & ";" & cwobjPatient.MedicalRecordNumber)
        
   Next nPatInfo 'Next Patient
        
        
    Set m_objDMBE.IDMappers = Nothing
    
Exit Function

ErrHandler:
    Screen.MousePointer = vbDefault
    gobjShared.ShowVBError Err.Number, Err.Description, "CDRAndLBBBSourceExecute", "DMF TransferManager", vbCritical
    Err.Clear
    Call SetCWStatusText("")

Exit Function
Resume

End Function
'</SCR 20841 CC>

Private Function ChangeCorePatient(ByVal lngNCIDEventType As Long, Optional sMessage As String = "")
    '--- SCR 134937  Sherry Wang   7/9/2007
    Dim strPatientUnitNumber() As String
    strPatientUnitNumber = Split(sMessage, ";")
    If IsPatientAlreadyChanged(strPatientUnitNumber(0)) = False Then
        Call cwobjCWComm.Message(lngNCIDEventType, sMessage, NCID_CW_APPLICATION, NCID_DATA_MANAGER)
        
        If gcolPatientsChanged Is Nothing Then Set gcolPatientsChanged = New Collection
        gcolPatientsChanged.Add strPatientUnitNumber(0)
    End If
End Function

Private Function ValidateCriteriaCDRAndLDDBAsSource(bResendTMIPMessages As Boolean) As Boolean

    'This function will call CDRAndLDDBAsSourceExecute for each dataset
    'It is a level to control each patient -- loading of patient is done
    'in the ExecuteMethod

    Dim nPatInfo As Long
    Dim transItem As I_RSTransfer
    Dim lngResult As DMErrorCodes
    Dim objPatient As MMMHISPatient.Patient
    Dim bLocalDBState As Boolean
    Dim sLastUnitNumber As String
    Dim oDMBEOps_Regen As DMBE.DMBE_Ops

    On Error GoTo ErrHandler
        
    ' Begin:SCR #27451;   Developer: MAG 09/25/2002 12:22 PM
    Set objPatient = New MMMHISPatient.Patient  'cwobjPatient
    objPatient.dataMgr = True
    ' End: SCR #27451;
    
    m_nRecordIndicator = 0
    
    ReDim m_tpTransferResult(0)
    
    'Set up the TransferResult of the first patient first
    m_tpTransferResult(m_nRecordIndicator).sPatientUnitNumber = m_objDMFData.PatientInfo(nPatInfo).sUnitNumber
    m_tpTransferResult(m_nRecordIndicator).sPatientName = m_objDMFData.PatientInfo(nPatInfo).sName
    m_tpTransferResult(m_nRecordIndicator).sPatientSSN = m_objDMFData.PatientInfo(nPatInfo).sSSN
    ' SCR #44079;   Developer: Mag 10/20/2003 01:12 PM
    m_tpTransferResult(m_nRecordIndicator).sPatientSponsorSSN = m_objDMFData.PatientInfo(nPatInfo).sSponsorSSN
    m_tpTransferResult(m_nRecordIndicator).Source = m_objDMFData.Source
    
    If bResendTMIPMessages = True Then
        m_nRecordIndicator = 0
            
        If m_nRecordIndicator + GETRESULTMEMORY < UBound(m_objDMFData.PatientInfo) Then
          ReDim Preserve m_tpTransferResult(m_nRecordIndicator + GETRESULTMEMORY)
        Else
          ReDim Preserve m_tpTransferResult(UBound(m_objDMFData.PatientInfo))
        End If
        
        'SCR 84451 and 84452 Regen
        Set oDMBEOps_Regen = m_objDMBE
    End If
    
    'FROM down here is about patient
    Set m_objDMBE.IDMappers = New DMBE.DMBE_IDMappers
    For nPatInfo = 0 To UBound(m_objDMFData.PatientInfo) 'Loop based on number of patient selected on the RegView Screen
        If bResendTMIPMessages = True Then
            LogTransaction nPatInfo & ") " & m_objDMFData.PatientInfo(nPatInfo).sSSN, "Start Patient", GetMemoryUsage
            
            'SCR 84451 and 84452 Set the patient's name for logging regen messages
            oDMBEOps_Regen.RegenPatientName = m_objDMFData.PatientInfo(nPatInfo).sName
            
            Screen.MousePointer = vbHourglass
        End If
        ' Begin:SCR #34534;   Developer: Brian Mowbray 04/17/2002 12:22 PM
        If sLastUnitNumber = m_objDMFData.PatientInfo(nPatInfo).sUnitNumber Then
          GoTo NextPerson
        End If
        
        sLastUnitNumber = m_objDMFData.PatientInfo(nPatInfo).sUnitNumber
        ' End:SCR #34534
        m_tpTransferResult(m_nRecordIndicator).sPatientUnitNumber = m_objDMFData.PatientInfo(nPatInfo).sUnitNumber
        m_tpTransferResult(m_nRecordIndicator).sPatientName = m_objDMFData.PatientInfo(nPatInfo).sName
        m_tpTransferResult(m_nRecordIndicator).sPatientSSN = m_objDMFData.PatientInfo(nPatInfo).sSSN
        ' SCR #44079;   Developer: Mag 10/20/2003 01:12 PM
        m_tpTransferResult(m_nRecordIndicator).sPatientSponsorSSN = m_objDMFData.PatientInfo(nPatInfo).sSponsorSSN
        m_tpTransferResult(m_nRecordIndicator).Source = m_objDMFData.Source
        
        If m_objDMFData.IsEICDestination = False Then
            m_tpTransferResult(m_nRecordIndicator).Destination = m_objDMFData.Destination
        Else: m_tpTransferResult(m_nRecordIndicator).Destination = eDM_EIC
        End If
        
        If CLng(m_objDMFData.PatientInfo(nPatInfo).sUnitNumber) <> objPatient.UnitNumber Then
            'If the patient is not the "current" patient
            bChangeDBFlagToLocal objPatient, (m_objDMFData.Source = eDM_LDDB)
            Call objPatient.LoadPatient(CLng(m_objDMFData.PatientInfo(nPatInfo).sUnitNumber), "") 'There is only one patient
            Call ChangeCorePatient(cwiSELECT_PATIENT, objPatient.UnitNumber & ";" & objPatient.MedicalRecordNumber)  '- Instruct core to display the patient
            Set m_objDMBE.PatientObj = objPatient
        End If
        
        'Source: CDR   Destination: LDDB
        If m_objDMFData.Source = eDM_CDR And m_objDMFData.Destination = eDM_LDDB Then
        
        'Source: LDDB   Destination: CDR
        ElseIf m_objDMFData.Source = eDM_LDDB And m_objDMFData.Destination = eDM_CDR Then
                            
        'SOURCE: LDDB    DESTINATION: TMIP Or EIC Or PIC Or DISK
        Else
            If bResendTMIPMessages = True Then
              Screen.MousePointer = vbHourglass
            End If
            lngResult = CDRAndLDDBAsSourceExecute(m_objDMFData.PatientInfo(nPatInfo).sUnitNumber, bResendTMIPMessages)
            
        End If
    
NextPerson:
    If bResendTMIPMessages = True Then
        
        m_tpTransferResult(0).sPatientName = "Process Complete"
        m_tpTransferResult(0).sPatientSponsorSSN = ""
        m_tpTransferResult(0).sPatientSSN = ""
        m_tpTransferResult(0).Result = ERROR_SUCCESS
        
    End If
    
    Call SetCWStatusText("")
    Set m_objDMBE.IDMappers = Nothing

    If bResendTMIPMessages = True Then
        LogTransaction m_objDMFData.PatientInfo(nPatInfo).sSSN, "End Patient", GetMemoryUsage
    End If
    
    Next nPatInfo 'Next Patient
ExitRoutine:
    Screen.MousePointer = vbDefault
        
    If bResendTMIPMessages = False Then
        If m_nRecordIndicator > 0 Then
           ReDim Preserve m_tpTransferResult(m_nRecordIndicator - 1)
        End If
    End If
    
    If bResendTMIPMessages = True Then
        oDMBEOps_Regen.RegenPatientName = ""
        Set oDMBEOps_Regen = Nothing
    End If

    objPatient.dataMgr = False
Exit Function

ErrHandler:
    Screen.MousePointer = vbDefault
    MsgBox "Error " & Err.Number & " occured in ValidateCriteriaCDRandLDDB. " & Chr(10) & Err.Description, vbCritical, "DMFE"
    Err.Clear
    Call SetCWStatusText("")

End Function


'<SCR 21771 CC>
Private Function MovePatientBetweenCDRAndLDDB(ByVal enmFrom As DMLocations, ByVal enmTo As DMLocations) As Boolean

    '<SCR 21771 CC>
    Dim blnResult As Boolean
    Dim objPatient As MMMHISPatient.Patient
    
    Set objPatient = cwobjPatient

    ' SCR #47633;   Developer:  01/02/2004
    If objPatient.SSN = "" Then
        MsgBox "The patient: " & objPatient.Name & " is missing an SSN, patient cannot be transferred." & _
               vbCrLf & vbCrLf & "Edit the patient with a valid SSN and retry sending", vbOKOnly, "MISSING PATIENT'S SSN"
        blnResult = False
    ElseIf enmFrom = eDM_LDDB And enmTo = eDM_CDR Then
        blnResult = objPatient.Copy_LDDBtoCDR
    ElseIf enmFrom = eDM_CDR And enmTo = eDM_LDDB Then
        blnResult = objPatient.Copy_CDRtoLDDB
    End If

    MovePatientBetweenCDRAndLDDB = blnResult
    
End Function

Private Function CDRAndLDDBAsSourceExecute(ByVal strUnitNumber As String, bResendTMIPMessages As Boolean) As DMErrorCodes

    Dim nDataSets As Long
    Dim transItem As IRSTransfer.I_RSTransfer
    Dim nCnt As Long
    Dim lngResult As DMErrorCodes
    Dim enmDataset As DMDataSets
    Dim dicDatasets As Scripting.Dictionary
    Dim sProblem As String
    Dim sAction As String
    Dim sModule As String
    
    On Error GoTo ErrHandler
    
    nCnt = m_nRecordIndicator
    
    If m_tpTransferResult(m_nRecordIndicator).Destination = eDM_PIC _
    Or m_tpTransferResult(m_nRecordIndicator).Destination = eDM_AGGPIC Then
        m_objDMBE.PICPatientUnitNumber = cwobjPatient.UnitNumber
        Set dicDatasets = New Scripting.Dictionary
    End If
    
    For nDataSets = 0 To UBound(m_objDMFData.Dataset) 'Inside loop for DataSet (In Validation, the outer loop is based on patient)
        
        'Build the result
        If (m_objDMBE.Destination <> eDM_EIC And bResendTMIPMessages = False) Or (bResendTMIPMessages = True And m_objDMBE.Destination <> eDM_EIC And m_objDMBE.Destination <> eDM_TMIP And m_objDMBE.Destination <> eDM_AGGPIC) Then
            If Not nDataSets = 0 Then
                'If it is the first transaction of this patient
                m_tpTransferResult(m_nRecordIndicator).sPatientUnitNumber = m_tpTransferResult(nCnt).sPatientUnitNumber
                m_tpTransferResult(m_nRecordIndicator).sPatientName = m_tpTransferResult(nCnt).sPatientName
                m_tpTransferResult(m_nRecordIndicator).sPatientSSN = m_tpTransferResult(nCnt).sPatientSSN
                m_tpTransferResult(m_nRecordIndicator).Source = m_tpTransferResult(nCnt).Source
                m_tpTransferResult(m_nRecordIndicator).Destination = m_tpTransferResult(nCnt).Destination
            End If
                
            m_tpTransferResult(m_nRecordIndicator).Dataset = m_objDMFData.Dataset(nDataSets)
            
        End If
        
        enmDataset = setDatasetEnum(m_objDMFData.Dataset(nDataSets))
        
        
        If enmDataset = eDM_Demographics And (m_objDMFData.Destination = eDM_CDR Or m_objDMFData.Destination = eDM_LDDB) Then
            'If it is a Demographics transfer between CDR and LDDB then do not need to call Demographics
            'The data has already been transfered in ValidateCriteriaCDRAndLDDBAsSource
            lngResult = ERROR_SUCCESS
        Else
                    
            If Len(Trim(m_tpTransferResult(nCnt).sPatientSSN)) <= 0 And m_objDMFData.Destination <> eDM_TMIP _
            And m_objDMFData.Destination <> eDM_EIC And m_objDMFData.Destination <> 8 Then
                                                                                    'eDM_JMEWS
                                                                                    'SF 62831
                m_tpTransferResult(m_nRecordIndicator).Result = ERROR_MISSING_SSN
                CDRAndLDDBAsSourceExecute = lngResult
                
                If m_objDMFData.EncounterSource = 1 Then
                   sProblem = "the Patient View File (PVF) file cannot be written." & vbCrLf & _
                              "The CHCSII-T system on the receiving end will not be capable of processing " & _
                              "these files without an SSN."
                              
                   sAction = "Edit the patient with a valid SSN, use Tools->Data Manager to resend the TMIP message " & _
                              "manually and a PVF will be sent."
                Else
                
                  sAction = "Edit the patient with a valid SSN and retry sending "
                  
                  Select Case m_objDMFData.Destination
                  Case eDM_AGGPIC
                      sProblem = "the Patient View File (PVF) or DISK"
                      sAction = sAction & "."
                  Case eDM_PIC
                    sProblem = "the PIC"
                    sAction = sAction & "to the PIC."
                  Case eDM_BMIST
                    sProblem = "the AHLTA-MOBILE"
                    sAction = sAction & "to AHLTA-MOBILE."
                  End Select
                
                  sProblem = sProblem & " file cannot be written." & vbCrLf & _
                              "The system on the receiving end will not be capable of processing " & _
                              "these files without an SSN."
                End If
                
                MsgBox "The patient: " & cwobjPatient.Name & " is missing an SSN, " & sProblem & _
                       vbCrLf & vbCrLf & sAction, vbOKOnly, "MISSING PATIENT'S SSN"

                Exit Function
            
            End If
            
            'Initialize the clinical module
            'Server NCID = 77610 - NCID_DEMOGRAPHICS_APPLICATION
            'Server NCID = 77611 - NCID_ENCOUNTER_APPLICATION
            
            
            
            Set transItem = gobjComm.InitializeOLEServer(CStr(m_objDMFData.Dataset(nDataSets)))
            
            If Not (transItem Is Nothing) Then
                m_objDMBE.Dataset = enmDataset
                If m_objDMBE.Destination = eDM_BMIST And m_objDMBE.Dataset = eDM_Demographics Then
                    
                   lngResult = m_objDMBE.WriteBMISTFile(eDM_XML, m_objDMFData.DestinationPath)
'
                Else
                  ' SCR #50386;   Developer: MAG 02/25/2004
                  ' This code was added to try to find out what module this error occurs in.
                  Select Case CStr(m_objDMFData.Dataset(nDataSets))
                    Case NCID_DEMOGRAPHICS_APPLICATION
                      sModule = "Demographics"
                    Case NCID_ENCOUNTER_APPLICATION
                      sModule = "Encounters"
                    Case NCID_ALLERGIES_APPLICATION
                      sModule = "Allergies"
                    Case NCID_CITA_APPLICATION
                      sModule = "Immunizations"
                    Case NCID_MEDICATIONS_APPLICATION
                      sModule = "Medications"
                    Case NCID_READINESS_APPLICATION
                      sModule = "Readiness"
                    Case NCID_PROBLEMS_APPLICATION
                      sModule = "Problems"
                    Case NCID_CLINICAL_NOTES_APPLICATION
                      sModule = "Clinical Notes"
                    Case NCID_VITAL_SIGNS_APPLICATION
                      sModule = "Vitals"
                    Case NCID_HEAR_APPLICATION
                      sModule = "Couplers"
                    'SF 62831
                    Case NCID_INPATIENT_APPLICATION
                      sModule = "Inpatient"
                    Case Else
                      sModule = "Unknown"
                  End Select

                    'Everything else will be executed here
                    
                    'Need to fake out the JMEWS destination
                    'if
                    
                    'Don't Do Couplers
                    If sModule <> "Couplers" Then
                      'If sModule = "Encounter" then the JMEWS files is written to EF when Encounter is signed
                      lngResult = transItem.Transfer(m_objDMBE)
                    Else
                      lngResult = ERROR_SUCCESS
                    End If
                    sModule = "DMFE"
                End If
                
                '<SCR 23125 CC> - if it is a PIC then need additional check
                If m_objDMBE.Destination = eDM_PIC Or m_objDMBE.Destination = eDM_AGGPIC Then
                    If TerminalPICError(lngResult) = True Then
                    
                      If bResendTMIPMessages = False Then
                        m_tpTransferResult(m_nRecordIndicator).Result = lngResult
                        m_nRecordIndicator = m_nRecordIndicator + 1
                        ReDim Preserve m_tpTransferResult(m_nRecordIndicator)
                      End If

                      CDRAndLDDBAsSourceExecute = lngResult
                      'Terminate all subsequent datasets
                      Exit Function
                    
                    ElseIf lngResult = ERROR_SUCCESS Then
                      dicDatasets.Add DatasetEnumToTxt(enmDataset), enmDataset
                    End If
                End If  ''</SCR 23125 CC>
                
                'SF SCR 76426 9-7-05
                'If the encounter module does not return any records and the destination is EIC, then
                'popup a message for the user.
                If m_objDMBE.Destination = eDM_EIC And CStr(m_objDMFData.Dataset(nDataSets)) = NCID_ENCOUNTER_APPLICATION And lngResult = ERROR_NO_RECORDS_FOR_DATASET_FOUND Then
                    MsgBox "There is no encounter data available for transfer.", vbInformation
                End If
                
                Set transItem = Nothing
            Else
                lngResult = ERROR_CANNOT_INSTANTIATE_CLINICAL_MODULES
            End If 'transItem Is Nothing
                  
        End If
        
        If bResendTMIPMessages = True Then
            If lngResult <> ERROR_SUCCESS And lngResult <> ERROR_NO_RECORDS_FOR_DATASET_FOUND Then
              If m_tpTransferResult(m_nRecordIndicator).Result = ERROR_SUCCESS Then
                m_tpTransferResult(m_nRecordIndicator).Result = lngResult
              End If
            End If
        ElseIf bResendTMIPMessages = False Then
            m_tpTransferResult(m_nRecordIndicator).Result = lngResult
            m_nRecordIndicator = m_nRecordIndicator + 1
            ReDim Preserve m_tpTransferResult(m_nRecordIndicator)
        End If
        
    Next nDataSets
    
    sModule = "Modules Complete"

    If m_objDMBE.EICIsDestination = True Then
        'Write to EIC only after all the strings have been accumulated (implemented in DMBE)
        ' SCR #44079;   Developer: Mag 10/20/2003 01:13 PM
        lngResult = m_objDMBE.WriteToEIC(m_tpTransferResult(nCnt).sPatientSponsorSSN)
        m_objDMBE.EICHL7Message = ""
        
        m_tpTransferResult(nCnt).Result = lngResult
        If lngResult = ERROR_SUCCESS Then
          If UBound(m_tpTransferResult) > 1 Then
            If m_tpTransferResult(1).Result = No_Data Then
              m_tpTransferResult(nCnt).Result = ERROR_EIC_SUCCESS_NO_ENCOUNTER
            End If
          End If
        End If
'        m_nRecordIndicator = m_nRecordIndicator + 1
'        ReDim Preserve m_tpTransferResult(m_nRecordIndicator)
    ElseIf m_objDMBE.Destination = eDM_PIC Or m_objDMBE.Destination = eDM_AGGPIC Then
        sModule = "Writing PIC"
        lngResult = m_objDMBE.WriteToPIC(m_tpTransferResult(nCnt).sPatientSSN, dicDatasets)
        m_objDMBE.ClearTransferRS
        
        If lngResult = ERROR_INSUFFICIENT_SPACE_ON_DEVICE Then
          'Remove All Dataset results related to this patient, except for Demographics
          m_nRecordIndicator = m_nRecordIndicator - UBound(m_objDMFData.Dataset) - 1
          If m_nRecordIndicator < 0 Then
            m_nRecordIndicator = 0
          End If
          ReDim Preserve m_tpTransferResult(m_nRecordIndicator)
        End If
        
        m_tpTransferResult(nCnt).Result = lngResult
        m_nRecordIndicator = m_nRecordIndicator + 1
        ReDim Preserve m_tpTransferResult(m_nRecordIndicator)
    End If
    sModule = "DONE"
Exit Function

ErrHandler:
    
    m_tpTransferResult(m_nRecordIndicator).Result = DMErrorCodes.ERROR_UNHANDLED_SYSTEM_ERROR
    m_nRecordIndicator = m_nRecordIndicator + 1
    ReDim Preserve m_tpTransferResult(m_nRecordIndicator)
    
    If m_objDMFData.EncounterSource = 1 And m_objDMFData.Destination = eDM_AGGPIC Then
      MsgBox "An error occurred when attempting to send PVF File to TMIP. Please try sending " & _
        "by using Data Manager to send (Source = LDDB, Destination = PIC, Path = " & m_objDMFData.SourcePath & ")", vbCritical, "PVF Send Error: " & Err.Number & ". Module: " & sModule
    ElseIf m_objDMFData.EncounterSource = 1 Then
      MsgBox "An error occurred when attempting to send data to TMIP. Please retry sending this message " & _
        "by using Data Manager to send the message again manually", vbCritical, "TMIP Send Error: " & Err.Number & ". Module: " & sModule
    Else
      MsgBox "Error " & Err.Number & " occurred in CDRAndLDDBAsSourceExecute. Module:" & sModule & Chr(10) & Err.Description, vbCritical, "DMFE"
    End If
    
    Err.Clear
    Screen.MousePointer = vbDefault
    Call SetCWStatusText("")
End Function


'<SCR 23125 CC>
Private Function TerminalPICError(ByVal enmResult As DMErrorCodes) As Boolean

    Dim blnReturn As Boolean

    Select Case enmResult
        Case DMErrorCodes.ERROR_PIC_DEVICE_DOES_NOT_EXIST
            blnReturn = True
        Case DMErrorCodes.ERROR_PIC_DEVICE_HAS_NOT_BEEN_INITIALIZED
            blnReturn = True
        Case DMErrorCodes.ERROR_PIC_DEVICE_IS_NOT_FOR_CURRENT_PATIENT
            blnReturn = True
        Case DMErrorCodes.ERROR_PIC_DEVICE_PATH_HAS_NOT_BEEN_SET
            blnReturn = True
        Case DMErrorCodes.ERROR_PIC_DEVICE_PATIENT_DOB_IS_DIFFERENT
            blnReturn = True
        Case DMErrorCodes.ERROR_PIC_DEVICE_PATIENT_SSN_IS_DIFFERENT
            blnReturn = True
        Case DMErrorCodes.ERROR_PIC_DEVICE_TYPE_OR_AVAILABILITY
            blnReturn = True
        Case DMErrorCodes.ERROR_PIC_MULTIPLE_PATIENT
            blnReturn = True
        Case DMErrorCodes.ERROR_INVALID_FORMAT
            blnReturn = True
        Case Else
            blnReturn = False
    End Select

    TerminalPICError = blnReturn
    
End Function
'</SCR 23125 CC>

Private Sub PICtoPICTransfer()

    Dim nCnt As Long
    Dim nDataSets As Long
    Dim lngResult As DMErrorCodes
    
    On Error GoTo ErrHandler
    
    m_nRecordIndicator = 0
    ReDim m_tpTransferResult(0)
    
    For nCnt = 0 To UBound(m_objDMFData.PatientInfo)
        'First count to see how many patient transferring
        
        'There is a possibility that the patient will not be in LDDB or CDR
        
        Set m_objDMBE.PatientObj = LoadPatient(m_objDMFData.PatientInfo(nCnt).sUnitNumber)
        Call TellCoreToLoadPatient(m_objDMFData.PatientInfo(nCnt).sUnitNumber)
        
        m_tpTransferResult(m_nRecordIndicator).sPatientUnitNumber = m_objDMFData.PatientInfo(nCnt).sUnitNumber
        m_tpTransferResult(m_nRecordIndicator).sPatientName = m_objDMFData.PatientInfo(nCnt).sName
        m_tpTransferResult(m_nRecordIndicator).sPatientSSN = m_objDMFData.PatientInfo(nCnt).sSSN
        m_tpTransferResult(m_nRecordIndicator).Source = m_objDMFData.Source
        m_tpTransferResult(m_nRecordIndicator).Destination = m_objDMFData.Destination
        
        DoEvents
        
        Call SetCWStatusText("Transferring " & m_objDMFData.PatientInfo(nCnt).sName & " ...")

        lngResult = m_objDMBE.CopyPIC
        If lngResult <> ERROR_SUCCESS Then
            m_tpTransferResult(m_nRecordIndicator).Result = lngResult
            Exit Sub
        End If
        
        For nDataSets = 0 To UBound(m_objDMFData.Dataset)

            Call SetCWStatusText(m_objDMFData.PatientInfo(nCnt).sName & " ... " & ResultTitle(m_objDMFData.Dataset(nDataSets)))
            'm_objDMBE.Dataset = m_objDMFData.Dataset(nDataSets)
            m_objDMBE.Dataset = setDatasetEnum(m_objDMFData.Dataset(nDataSets))
                
            If Not nDataSets = 0 Then
                'Need to populate the transfer result
                m_tpTransferResult(m_nRecordIndicator).sPatientUnitNumber = m_tpTransferResult(nCnt).sPatientUnitNumber
                m_tpTransferResult(m_nRecordIndicator).sPatientName = m_tpTransferResult(nCnt).sPatientName
                m_tpTransferResult(m_nRecordIndicator).sPatientSSN = m_tpTransferResult(nCnt).sPatientSSN
                m_tpTransferResult(m_nRecordIndicator).Source = m_tpTransferResult(nCnt).Source
                m_tpTransferResult(m_nRecordIndicator).Destination = m_tpTransferResult(nCnt).Destination
            End If
            
            m_tpTransferResult(m_nRecordIndicator).Result = lngResult
            m_tpTransferResult(m_nRecordIndicator).Dataset = m_objDMFData.Dataset(nDataSets)
                
            m_nRecordIndicator = m_nRecordIndicator + 1
        
            ReDim Preserve m_tpTransferResult(m_nRecordIndicator)
    
        Next nDataSets
    
    Next nCnt

    ReDim Preserve m_tpTransferResult(m_nRecordIndicator - 1)

Exit Sub

ErrHandler:
    Screen.MousePointer = vbDefault
    gobjShared.ShowVBError Err.Number, Err.Description, "PICtoPICTransfer", "DMF TransferManager", vbCritical
    Err.Clear
    Call SetCWStatusText("")
End Sub

Private Sub PICtoCDROrLDDBTransfer()

    Dim nCnt As Long
    Dim nDataSets As Long
    Dim lngResult As DMErrorCodes
    Dim bSaveUseLocalFlag As Boolean
    Dim newPatient As MMMHISPatient.Patient
    Dim transItem As I_RSTransfer
    Dim lngUnitNumber As Long
    Dim blnEncountersSelected As Boolean
    
    On Error GoTo ErrHandler
    
    Call bChangeDBFlagToLocal(m_objDMBE.PatientObj, True)
    
    m_nRecordIndicator = 0
    ReDim m_tpTransferResult(0)
    
    'PICPatientUnitNumber
    
    For nCnt = 0 To UBound(m_objDMFData.PatientInfo) '- Loop based on number of patient selected
        Set m_objDMBE.IDMappers = New DMBE.DMBE_IDMappers
        
        'First count to see how many patient transferring
    
        'Set m_objDMBE.IDMappers = New IDMBE.I_IDMappers 'DMBE.DMBE_IDMappers
        
        m_objDMBE.PICPatientUnitNumber = m_objDMFData.PatientInfo(nCnt).sUnitNumber 'SCR 21659 CC
        m_objDMBE.FileName = m_objDMFData.PatientInfo(nCnt).sId
        
        m_tpTransferResult(m_nRecordIndicator).sPatientUnitNumber = m_objDMFData.PatientInfo(nCnt).sUnitNumber
        m_tpTransferResult(m_nRecordIndicator).sPatientName = m_objDMFData.PatientInfo(nCnt).sName
        m_tpTransferResult(m_nRecordIndicator).sPatientSSN = m_objDMFData.PatientInfo(nCnt).sSSN
        m_tpTransferResult(m_nRecordIndicator).Source = m_objDMFData.Source
        m_tpTransferResult(m_nRecordIndicator).Destination = m_objDMFData.Destination
    
        If PICPatientIsInLDDB(nCnt) = False Then
            m_tpTransferResult(m_nRecordIndicator).Result = DMErrorCodes.ERROR_CANNOT_VALIDATE_PATIENT
            Exit Sub
        Else
            If cwobjPatient.UnitNumber = 0 Then
                m_tpTransferResult(m_nRecordIndicator).Result = DMErrorCodes.ERROR_CANNOT_VALIDATE_PATIENT
                Exit Sub
            End If
        End If
        If Len(m_objDMFData.PatientInfo(nCnt).sSSN) > 0 Then
          blnEncountersSelected = False
          
          'Need to Add Providers to IDMapper
          Call SetProviderIDMapper(m_objDMBE)
          
          For nDataSets = 0 To UBound(m_objDMFData.Dataset)
          
              If m_objDMFData.Dataset(nDataSets) = NCID_ENCOUNTER_APPLICATION Then
                blnEncountersSelected = True
              End If
                  
              '<SCR 25154 CC>
              If blnEncountersSelected = True And (m_objDMFData.Dataset(nDataSets) = eDM_Vitals Or m_objDMFData.Dataset(nDataSets) = eDM_Problems) Then
                  'Do nothing should already have been called from EncounterOps
                  
              Else
                    
                  If Not nDataSets = 0 Then
                      m_tpTransferResult(m_nRecordIndicator).sPatientUnitNumber = m_objDMFData.PatientInfo(nCnt).sUnitNumber
                      m_tpTransferResult(m_nRecordIndicator).sPatientName = m_objDMFData.PatientInfo(nCnt).sName
                      m_tpTransferResult(m_nRecordIndicator).sPatientSSN = m_objDMFData.PatientInfo(nCnt).sSSN
                      m_tpTransferResult(m_nRecordIndicator).Source = m_objDMFData.Source
                      m_tpTransferResult(m_nRecordIndicator).Destination = m_objDMFData.Destination
                  End If
                  m_tpTransferResult(m_nRecordIndicator).Dataset = m_objDMFData.Dataset(nDataSets)
                  
                  
                  Set transItem = gobjComm.InitializeOLEServer(CStr(m_objDMFData.Dataset(nDataSets)))
                  If Not (transItem Is Nothing) Then
                      'Here Need To Find Someway to put the rs back
                      'Force the EIC number in there
                      'Temporary replacement for unitnumber
                      
                      'm_objDMBE.EICHL7Message = m_objDMFData.PatientInfo(nCnt).sUnitNumber
                      m_objDMBE.PICPatientUnitNumber = m_objDMFData.PatientInfo(nCnt).sUnitNumber
                      
                      m_objDMBE.Dataset = setDatasetEnum(m_objDMFData.Dataset(nDataSets))
                      lngResult = transItem.Transfer(m_objDMBE)
                      Set transItem = Nothing
                  Else
                      lngResult = DMErrorCodes.ERROR_CANNOT_INSTANTIATE_CLINICAL_MODULES
                  End If
                          
                  m_tpTransferResult(m_nRecordIndicator).Result = lngResult
                  m_nRecordIndicator = m_nRecordIndicator + 1
                  
                  ReDim Preserve m_tpTransferResult(m_nRecordIndicator)
              
              End If
              'Set m_objDMBE.IDMappers = Nothing
              
          Next nDataSets
                  
          Dim enmPatientSearchResult As eSearchResults
    
          enmPatientSearchResult = cwobjPatient.ValidatePatient(True)
          Select Case enmPatientSearchResult
            Case esr_MatchedSinglePatient, esr_CreatedNewPatient, esr_MatchedNameSSN
                'The only time want to select the patient
                'This will prevent the patient from being loaded again
                If Trim$(m_objDMFData.PatientInfo(nCnt).sUnitNumber) <> Trim(cwobjPatient.UnitNumber) Then
                  m_objDMFData.PatientInfo(nCnt).sUnitNumber = cwobjPatient.UnitNumber
                  UpdateSearchGridUnit m_objDMFData.PatientInfo(nCnt).sUnitNumber, cwobjPatient.UnitNumber
                End If
                
                Call ChangeCorePatient(cwiSELECT_PATIENT, cwobjPatient.UnitNumber & ";" & cwobjPatient.MedicalRecordNumber)
                ' Begin:SCR #34301;   Developer: Mag 04/15/2003 04:22 PM
                Call cwobjCWComm.Message(cwiREFRESH_DATA, "", NCID_CW_APPLICATION, NCID_DEMOGRAPHICS_APPLICATION)
                ' End: SCR #34301;
            Case Else
                'Clear the patient
                Call ChangeCorePatient(cwiCLEAR_PATIENT, lngUnitNumber & ";" & "")
                lngResult = DMErrorCodes.ERROR_CANNOT_VALIDATE_PATIENT
                '-- raise some error condition here?
                m_tpTransferResult(0).Result = lngResult
          End Select
                    
          m_objDMBE.ClearTransferRS
      Else
          MsgBox "The patient: " & m_objDMFData.PatientInfo(nCnt).sName & " is missing an SSN. CHCSII-T Cannot import data for " & _
                 "Patients that do not have SSN's", vbOKOnly, "MISSING PATIENT'S SSN"
          
      End If
    Next nCnt
    
    Set m_objDMBE.IDMappers = Nothing

    Call bChangeDBFlagToLocal(m_objDMBE.PatientObj, bSaveUseLocalFlag)

Exit Sub

ErrHandler:

    m_tpTransferResult(m_nRecordIndicator).Result = DMErrorCodes.ERROR_UNHANDLED_SYSTEM_ERROR
    'm_nRecordIndicator = m_nRecordIndicator + 1
    'ReDim Preserve m_tpTransferResult(m_nRecordIndicator)
    
    Screen.MousePointer = vbDefault
    
    MsgBox "Error " & Err.Number & " occured in PICtoCDROrLDDBTransfer. " & Chr(10) & Err.Description, vbCritical, "DMFE"
    'gobjShared.ShowVBError Err.Number, Err.Description, "PICtoCDROrLDDBTransfer", "DMF TransferManager", vbCritical
    Err.Clear
    Call SetCWStatusText("")
Exit Sub
Resume Next
End Sub

Private Function PICPatientIsInLDDB(ByVal nCnt As Long) As Boolean

    Dim strPatientName() As String
    Dim rsMMI_ID As ADODB.Recordset
    Dim rsMMI_NAME As ADODB.Recordset
    Dim colRSObjects As Collection
    Dim objRSObject As IDMBE.I_RSObjects
    
    On Error GoTo ErrHandler
        
    Set rsMMI_ID = New ADODB.Recordset
    Set rsMMI_NAME = New ADODB.Recordset
    
    Call GetFieldProps("MMI_ID", rsMMI_ID)
    Call GetFieldProps("MMI_NAME", rsMMI_NAME)
    
    rsMMI_ID.AddNew
        rsMMI_ID.Fields("UNIT_NUMBER").Value = m_objDMFData.PatientInfo(nCnt).sUnitNumber
        rsMMI_ID.Fields("SSN").Value = m_objDMFData.PatientInfo(nCnt).sSSN
    rsMMI_ID.Update
        
    rsMMI_NAME.AddNew
        rsMMI_NAME.Fields("UNIT_NUMBER").Value = m_objDMFData.PatientInfo(nCnt).sUnitNumber
        strPatientName = Split(m_objDMFData.PatientInfo(nCnt).sName, ",")
        If UBound(strPatientName) = 1 Then
            rsMMI_NAME.Fields("LAST_NAME").Value = Trim(strPatientName(0))
            strPatientName = Split(Trim(strPatientName(1)), " ")
            If UBound(strPatientName) = 1 Then
              rsMMI_NAME.Fields("FIRST_NAME").Value = Trim(strPatientName(0))
              rsMMI_NAME.Fields("MIDDLE_NAME").Value = Trim(strPatientName(1))
            Else
              rsMMI_NAME.Fields("FIRST_NAME").Value = Trim(strPatientName(0))
            End If
        Else
            'Error
             PICPatientIsInLDDB = False
             Exit Function
        End If
    rsMMI_NAME.Update
    
    Set colRSObjects = New Collection
    
    Set objRSObject = New DMBE.RSObject_Ops
    objRSObject.TableName = "MMI_ID"
    Set objRSObject.DBRecordSet = rsMMI_ID
    colRSObjects.Add objRSObject, "MMI_ID"
    
    Set objRSObject = New DMBE.RSObject_Ops
    objRSObject.TableName = "MMI_NAME"
    Set objRSObject.DBRecordSet = rsMMI_NAME
    colRSObjects.Add objRSObject, "MMI_NAME"
    
    Call cwobjPatient.Validation_Criteria_Eic(colRSObjects)
    Dim retcode As eSearchResults
    
    retcode = cwobjPatient.ValidatePatient(True)
    Select Case retcode
        Case esr_MatchedSinglePatient, eSearchResults.esr_CreatedNewPatient, eSearchResults.esr_MatchedNameSSN
            ChangeCorePatient cwiSELECT_PATIENT, cwobjPatient.UnitNumber & ";" & cwobjPatient.MedicalRecordNumber
            PICPatientIsInLDDB = True
    Case Else
        ChangeCorePatient cwiCLEAR_PATIENT, cwobjPatient.UnitNumber & ";" & cwobjPatient.MedicalRecordNumber
    
        PICPatientIsInLDDB = False
    End Select
        
        
Exit Function

ErrHandler:
    Screen.MousePointer = vbDefault
    gobjShared.ShowVBError Err.Number, Err.Description, "PICPatientIsInLDDB", "DMF TransferManager", vbCritical
    Call SetCWStatusText("")
    PICPatientIsInLDDB = False
End Function

Private Function LDDBFacilityIsValid() As Boolean
    
    '-- see if the facility ncid is 'real' and available in the cdr
    
    Dim oGEMSSQL As ICHCSII_SQL.ISqlOpsEx
    Dim oCDRSQL As ICHCSII_SQL.ISqlOps
    Dim strOldNCID As String
    Dim strNewNCID As String
    Dim strMachineNCID As String
    Dim strSQL As String
    Dim oGEMSDAS As ICHCSII_DAS
    'Dim oGEMSSQL1 As ICHCSII_SQL.ISqlOpsEx
    
    Dim rsMachine As New ADODB.Recordset
    Dim rsFacility As New ADODB.Recordset
    Dim blnFound As Boolean
    Dim oCmd As ADODB.command
    Dim oParam As ADODB.Parameter
    Dim bNCIDExists As Boolean
    Dim oDbCDR As New CHCSII_CONN.Conn
    Dim oDbLDDB As New CHCSII_CONN.Conn
    
    Set oGEMSDAS = oDbLDDB.CHCSII_DAS(Auto) 'New GEMS_DAS
    Set oCDRSQL = oDbCDR.CHCSII_SQLOPS(CDR)  'CHCSII_SQL.SqlOps
    
    '<< Begin:SCR #28284;   Developer: Brian Mowbray 10/09/2002 05:18 PM
    On Error Resume Next
    '>> End: SCR #28284;
    
    Set rsFacility = oGEMSDAS.ExecuteSQL("SELECT FACILITY_NCID FROM FACILITY_NCID_CORRELATION")
    'oGEMSSQL.Execute "SELECT FACILITY_NCID FROM FACILITY_NCID_CORRELATION"
    blnFound = False
    
    Do Until rsFacility.EOF
        
        Set oGEMSSQL = oDbLDDB.CHCSII_SQLOPS(Auto) 'New GEMS_SQL.SqlOps
        
        strOldNCID = rsFacility.Fields("FACILITY_NCID").Value
        strNewNCID = strOldNCID
        
        Dim bRepresentation As Boolean
        
        
        
        'If oCDRSQL.NCIDIsCDRFacility(strOldNCID) = False Then
        If oCDRSQL.ValiadeFacilityAgainstCDR(strNewNCID, cwobjLogonEnviron.EnterpriseNCID, bRepresentation, True) = False Then
            'If it returns false without an error then it means that new ncid is returned,
            'either because it is a representation or because it was new created
            
            'Will be called only once
            If blnFound = False Then
                'Get the machine Facility NCID
                strSQL = "SELECT APVALUE FROM OSH_INI_SETTINGS "
                strSQL = strSQL & "WHERE APSECTION = 'LOCAL_ENV' AND APKEY = 'FACILITY_NCID'"
    
                Set rsMachine = oGEMSDAS.OpenRecordset(strSQL, adOpenForwardOnly, adLockOptimistic)
                If Not rsMachine.EOF Then
                    strMachineNCID = rsMachine.Fields(0).Value
                Else
                    'Need to raise error
                    LDDBFacilityIsValid = False
                    Exit Function
                End If
                
                Set rsMachine = Nothing
                blnFound = True
                
            End If
            
            'strSQL = "UPDATE FACILITY_NCID_CORRELATION "
            'strSQL = strSQL & "SET FACILITY_NCID = " & strNewNCID & " "
            'strSQL = strSQL & "WHERE FACILITY_NCID = " & strOldNCID
            
            
            ' Begin:SCR ;   Developer: MAG 09/25/2002 06:02 PM
            '<< Begin:SCR #28284;   Developer: Brian Mowbray 10/09/2002 05:17 PM
            'Changed the table name
            strSQL = "SELECT * FROM FACILITY_NCID_CORRELATION "
            strSQL = strSQL & "Where FACILITY_NCID = " & strNewNCID
            '>> End: SCR #28284;
            
            Call oGEMSSQL.Execute(strSQL)
            
            bNCIDExists = Not oGEMSSQL.EOF
            ' End: SCR #20463;
            
            On Error GoTo TransError
            
            'Begin trans because need to update more than one tables
            oGEMSSQL.BeginTrans
            
            ' Begin:SCR #20463;   Developer: MAG 09/25/2002 06:02 PM
            'RUN UPDFACILITYNCID Stored Procedure : First change the correlation table
            Set oCmd = New ADODB.command
            Set oParam = New ADODB.Parameter
            
            With oCmd
                .CommandText = "UPDFACILITYNCID"
                .CommandType = adCmdStoredProc
                Set oParam = .CreateParameter("oldid", adNumeric, adParamInput)
                oParam.Value = CLng(strOldNCID)
                .Parameters.Append oParam
                Set oParam = .CreateParameter("newid", adNumeric, adParamInput)
                oParam.Value = CLng(strNewNCID)
                .Parameters.Append oParam
            End With
            
            Call oGEMSSQL.ExecuteCmd(oCmd)
            
            If bNCIDExists Then
              strSQL = "Delete From Facility_NCID_CORRELATION "
              strSQL = strSQL & "Where Facility_NCID = " & strOldNCID
            Else
              strSQL = "Update FACILITY_NCID_CORRELATION "
              strSQL = strSQL & " Set FACILITY_NCID = " & strNewNCID
              strSQL = strSQL & " Where FACILITY_NCID = " & strOldNCID
            End If
            Call oGEMSSQL.Execute(strSQL)
            ' End: SCR #20463;
                
            strSQL = "SELECT * FROM OSH_INI_SETTINGS "
            strSQL = strSQL & "WHERE APSECTION = 'LOCAL_ENV' AND APKEY = 'FACILITY_NCID' "
            
            Call oGEMSSQL.Execute(strSQL)
            
            If CLng(strMachineNCID) = CLng(strOldNCID) Then
                strMachineNCID = strNewNCID
            End If
            
            If Not oGEMSSQL.EOF Then
                Call oGEMSSQL.AppendAsChunk("APVALUE", strMachineNCID)
                oGEMSSQL.Update
            End If
                
            'strSQL = "UPDATE OSH_INI_SETTINGS "
            'strSQL = strSQL & "SET APVALUE = " & strNewNCID & " "
            'strSQL = strSQL & "WHERE APSECTION = 'LOCAL_ENV' AND APKEY = 'FACILITY_NCID' "
            'strSQL = strSQL & "AND APVALUE = " & strOldNCID
            'then update the stored INI setting
            'Call oGEMSDAS.ExecuteSQL(strSQL)

            oGEMSSQL.CommitTrans

            On Error GoTo ErrHandler
            ' End: SCR #20463;
            
        End If
    
        rsFacility.MoveNext
    Loop
    
    Set oGEMSDAS = Nothing
    Set oCDRSQL = Nothing
    Set oGEMSSQL = Nothing
    
    LDDBFacilityIsValid = True

Exit Function

TransError:
    oGEMSSQL.RollbackTrans
    'oGEMSDAS.RollbackTrans
ErrHandler:
    LDDBFacilityIsValid = False
    gobjShared.ShowVBError Err.Number, Err.Description, _
                        "LDDBFacilityIsValid", "DMFTransferManager", vbCritical
    Err.Clear
    
End Function

Private Sub Class_Initialize()
    '---create shared object used to show VB errors, etc.---
    
    If gobjShared Is Nothing Then
        Set gobjShared = New CWShared
    End If
    
    '---set the help file for this application---
    App.HelpFile = gobjShared.CWHelpFile

End Sub
'
Private Function TellCoreToLoadPatient(sUnitNumber As String) As Boolean
    
    Dim sPatientInfo As String
    
    On Error GoTo UnexpectedError
    
    sPatientInfo = sUnitNumber & ";"    '*-- Null string MedRecNumber
    gobjComm.Message cwiSELECT_PATIENT, sPatientInfo, _
                        NCID_CW_APPLICATION, NCID_DATA_MANAGER
    TellCoreToLoadPatient = True
    Exit Function
    
UnexpectedError:
    TellCoreToLoadPatient = False
    gobjShared.ShowVBError Err.Number, Err.Description, _
                        "bGetNewPatient", "DM:gfrmRegularView", vbCritical
    Exit Function
    Resume
End Function


'*-- Coleman 101801: Use this function to change the database flag...it will
'*-- make sure the Patient DB object is set correctly.
Private Function bChangeDBFlagToLocal(ByRef oPatient As MMMHISPatient.Patient, _
                                    ByVal bUseLocal As Boolean) As Boolean
    On Error GoTo UnexpectedError
    
    Call oPatient.InitClassForDBSwitch
    bChangeDBFlagToLocal = True
    
    Exit Function
    
UnexpectedError:
    bChangeDBFlagToLocal = False
    Exit Function
    Resume
End Function

Private Sub SetCWStatusText(ByVal strScreenText As String)

    If Not gobjComm Is Nothing Then
        Call gobjComm.SetStatusText(strScreenText)
    End If

End Sub

Private Sub RaiseMsgBox(ByVal strMessage As String, ByVal intStyle As VbMsgBoxStyle, strTitle As String)

    'Until I can figure something out
    MsgBox strMessage, intStyle, strTitle
    
End Sub

Private Sub CheckTransferResult(lngResult As Long, ErrText As String, item As String)
  
    If lngResult = 0 Then
        MsgBox "The " + item + " transaction was Successful! ", vbOKOnly + vbInformation, "Data Manager"
    Else
        Select Case lngResult
            Case -100
                RaiseMsgBox "The " + item + "  transaction was failed!  Destination not supported. ", vbOKOnly + vbInformation, "Data Manager"
            Case -200
                RaiseMsgBox "The " + item + "  transaction was failed! Source not supported. ", vbOKOnly + vbInformation, "Data Manager"
            Case -300
                RaiseMsgBox "The " + item + "  transaction was failed! Format not supported. ", vbOKOnly + vbInformation, "Data Manager"
            Case -400
                RaiseMsgBox "The " + item + "  transaction was failed! Dataset not supported. ", vbOKOnly + vbInformation, "Data Manager"
            Case -700
                RaiseMsgBox "The " + item + " transacton was failed. The data selected has already been transferred. ", vbInformation, "Data manager"
            Case Else
                RaiseMsgBox "Error in processing " + item + "  request; result = " & lngResult & vbCrLf & ErrText, vbCritical, "Data Manager"
        End Select
  End If
  
End Sub


Private Function ResultTitle(ByVal lngNCID As Long) As String

    Select Case lngNCID
        Case NCID_CLINICAL_NOTES_APPLICATION
            ResultTitle = "Clinical Notes "
        Case NCID_ALLERGIES_APPLICATION
            ResultTitle = "Allergies "
        Case NCID_ENCOUNTER_APPLICATION
            ResultTitle = "Encounters "
        Case NCID_DEMOGRAPHICS_APPLICATION
            ResultTitle = "Demographincs "
        Case NCID_CITA_APPLICATION
            ResultTitle = "Immunization "
        Case NCID_MEDICATIONS_APPLICATION
            ResultTitle = "Medications "
        Case NCID_READINESS_APPLICATION
            ResultTitle = "Readiness "
        Case NCID_PROBLEMS_APPLICATION
            ResultTitle = "Probem List "
        Case NCID_HEAR_APPLICATION
            ResultTitle = "Couplers "
        Case NCID_VITAL_SIGNS_APPLICATION
            ResultTitle = "Vitals "
        Case Else
            ResultTitle = "Unknown "
    End Select

End Function

Public Sub SetCommunicationObject(objComm As MMMHISComm.Comm)

    Set gobjComm = objComm
    
End Sub

Public Function LoadPatient(ByVal strUnitNumber As String) As MMMHISPatient.Patient

    Dim objPatient As MMMHISPatient.Patient
    
    Set objPatient = New MMMHISPatient.Patient
    Call objPatient.LoadPatient(CLng(strUnitNumber), "")
    Set LoadPatient = objPatient

End Function

Private Function setDatasetEnum(ByVal lAppNCID As Long) As DMDataSets
    Select Case lAppNCID
        Case NCID_DEMOGRAPHICS_APPLICATION
            setDatasetEnum = eDM_Demographics
        Case NCID_ALLERGIES_APPLICATION
            setDatasetEnum = eDM_Allergies
        Case NCID_PROBLEMS_APPLICATION
            setDatasetEnum = eDM_Problems
        Case NCID_MEDICATIONS_APPLICATION
            setDatasetEnum = eDM_Medications
        Case NCID_READINESS_APPLICATION
            setDatasetEnum = eDM_Readiness
        Case NCID_IMMUNIZATION_SYNCH_APPLICATION
            setDatasetEnum = eDM_Immunizations
        '<SCR 22566 CC>
        Case NCID_CITA_APPLICATION
            setDatasetEnum = eDM_Immunizations
        '</SCR 22566 CC>
        Case NCID_VITAL_SIGNS_APPLICATION
            setDatasetEnum = eDM_Vitals
        Case NCID_ENCOUNTER_APPLICATION
            setDatasetEnum = eDM_Encounters
        Case NCID_CLINICAL_NOTES_APPLICATION
            setDatasetEnum = eDM_ClinicNotes
        Case NCID_HEAR_APPLICATION
            setDatasetEnum = eDM_Couplers
        'SF 62830
        Case NCID_INPATIENT_APPLICATION
            setDatasetEnum = 12 'eDM_Inpatient
    End Select
End Function

Private Function SetProviderIDMapper(oDMBE As IDMBE.I_DMBE) As Boolean
Dim oRSt As IDMBE.I_RSObjects
Dim colRst As Collection
Dim oDAS As ICHCSII_DAS
Dim oSQL As ICHCSII_SQL.ISqlOpsEx
Dim oRS As ADODB.Recordset
Dim bReturn As Boolean
Dim asName() As String
Dim sFName As String
Dim sLName As String
Dim sOldProvNCID As String
Dim sProvNCID As String
Dim sProvSSN As String

  Set colRst = New Collection    '-- Of oRSt

  If Not oDMBE.Execute(colRst) = 0 Then Exit Function
  
  For Each oRSt In colRst
    If oRSt.TableName = "PROVIDER" Then
      Set oDAS = oDatabase.CHCSII_DAS(Auto)
      Set oSQL = oDatabase.CHCSII_SQLOPS(Auto)
      Set oRS = oRSt.DBRecordSet
      Do While Not oRS.EOF
        asName = Split(oRS("NAME") & vbNullString, ",")
        If UBound(asName) = 0 Then
          sLName = asName(0)
        Else
          sLName = asName(0)
          sFName = asName(1)
        End If
        
        sProvNCID = oRS("NCID") & vbNullString
        sOldProvNCID = sProvNCID
        sProvSSN = oRS("SSN") & vbNullString
        bReturn = oSQL.ValidateProvider2(sProvNCID, sProvSSN, sFName, "", sLName, True)
        If Not oDMBE.IDMappers.Exists(NCID_PROVIDER_SERVICES, "NCID", sOldProvNCID) Then
          oDMBE.IDMappers.Add NCID_PROVIDER_SERVICES, "NCID", sOldProvNCID, sProvNCID
        End If
        oRS.MoveNext
      Loop
      Exit For
    End If
  Next oRSt
  
End Function

'<< Begin:SCR #34987;   Developer: Brian Mowbray 04/18/2003 03:36 PM
Private Function GetPVFFlag() As Boolean
    
    Dim objReg As RegObj.Registry
    Dim objKey As RegObj.regKey
    Dim strKeyName As String
    Dim objValue    As RegObj.regValue
    
    On Error Resume Next
    Set objReg = New RegObj.Registry
    
    '<< SCR #29068;   Developer: Brian Mowbray 02/04/2004 09:23 PM
    strKeyName = gobjShared.RegistryKey & "\DataManager"
    
    Set objReg = New RegObj.Registry
    Set objKey = objReg.RegKeyFromString(strKeyName)
    
    'Set the default to False if the value is not in the registry. PVF will not be sent if the TMIP flag is not set to False also
    GetPVFFlag = False
    
    If Not objKey Is Nothing Then
        For Each objValue In objKey.Values
          If UCase$(objValue.Name) = "PVF" Then
              GetPVFFlag = (Trim$(objValue.Value) = "Y")
          End If
          Set objValue = Nothing
        Next objValue
    End If
    Set objReg = Nothing
    Set objKey = Nothing

End Function

Private Function SendPVF(bEncSigned As Long, ByVal bResendTMIPMessages As Boolean) As Boolean
Dim oDMPatient(0) As DMFPatientInfo
Dim lDataSets() As Long

  On Error GoTo ErrHandler:

  '************************************************
  'SCR 68262
  'B.Fote 5-Apr-2005
  If bResendTMIPMessages = False Then
    Set m_objDMFData = New DMFTransactionData
  Else
    'SCR 68262
    'Don't Wipe Out the Patient List.
    'We use the same Patient List for PVF Files.
  End If
  '************************************************
  
  m_nRecordIndicator = 0

  m_objDMBE.ClearParameters

  With m_objDMFData
      If bEncSigned And bResendTMIPMessages = False Then
        oDMPatient(0).sUnitNumber = cwobjPatient.UnitNumber
        oDMPatient(0).sSSN = cwobjPatient.SSN
        oDMPatient(0).sName = cwobjPatient.Name
        .PatientInfo = oDMPatient
      '************************************************
      'SCR 68262
      'B.Fote 5-Apr-2005
      ElseIf bResendTMIPMessages = True Then
        gfrmRegularView.cmdTMIP.Tag = "PVF"
      '************************************************
      Else
        .PatientInfo = GetSelectedPatientFromGrid
      End If

      ReDim lDataSets(8)
      lDataSets(0) = NCID_DEMOGRAPHICS_APPLICATION
      lDataSets(1) = NCID_ENCOUNTER_APPLICATION
      lDataSets(2) = NCID_ALLERGIES_APPLICATION
      lDataSets(3) = NCID_CITA_APPLICATION
      lDataSets(4) = NCID_MEDICATIONS_APPLICATION
      lDataSets(5) = NCID_READINESS_APPLICATION
      lDataSets(6) = NCID_PROBLEMS_APPLICATION
      lDataSets(7) = NCID_CLINICAL_NOTES_APPLICATION
      lDataSets(8) = NCID_VITAL_SIGNS_APPLICATION

      .Dataset = lDataSets

      .Destination = eDM_AGGPIC
      .Source = eDM_LDDB

      '**************************************
      .Compression = eDM_NoCompression
      '**************************************


      '**************************************
      'SCR 62595 Always Encrypt
      '.Encryption = eDM_NoCrypto
      .Encryption = eDM_Encrypted
      '**************************************
      .Format = eDM_XML
      .IsEICDestination = False

      If bEncSigned Then
        .EncounterSource = 1
      Else
        .EncounterSource = 0
      End If
      ' Begin:SCR #38308;   Developer: Mag 06/30/2003 05:46 PM
      ' SCR #47954;   Developer: MAG 01/13/2004
      .DestinationPath = Environ$("TEMP") & "\" 'GetPVFPath(KEY_WRITE)
      .SourcePath = GetPVFPath(KEY_WRITE)
      m_objDMBE.Destination = .Destination
      m_objDMBE.Source = .Source
      m_objDMBE.Format = .Format
      m_objDMBE.Crypto = .Encryption
      m_objDMBE.Compression = .Compression
      m_objDMBE.PicPath = .SourcePath
      m_objDMBE.AltPicPath = .DestinationPath
      m_objDMBE.EICIsDestination = .IsEICDestination

      '<< Begin:SCR #30953;   Developer: Brian Mowbray 12/18/2002 03:36 PM
      .EncounterSource = .EncounterSource
      '>> End: SCR #30953;

  End With

  If Not gobjComm Is Nothing Then
      Call gobjComm.SetStatusText("Sending Encounter to PVF ...")
  End If

  Call ValidateCriteriaCDRAndLDDBAsSource(bResendTMIPMessages)

Exit Function

ErrHandler:
    gobjShared.ShowVBError Err.Number, Err.Description, "SendPVF", "DMFTransferManager", vbCritical
End Function

