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
Attribute VB_Name = "HL7XMLParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Note: this whole class was created for SCR 20841 and SCR 22439. CC

Option Explicit
Private m_objGEMSDAS As GEMS_DAS

Private menMessageSource As XMLMessageSource
Private Const rtfTag As String = "{\rtf"


Public Function GetRecordsetsFromEIC(ByRef colRSObject As Collection, ByVal lngUnitNumber As Long, ByVal enmDataSet As DMDataSets, ByVal strEICXmlData As String, ByVal enmWhoseMessage As XMLMessageSource) As Long

    Dim oNode As MSXML.IXMLDOMElement
    Dim oDom As MSXML.DOMDocument
    Dim lngResult As Long
    
    On Error GoTo ErrHandler
    
    Set oDom = New MSXML.DOMDocument
    
    'Disable DTD Validation
    oDom.async = False
    oDom.validateOnParse = False
    oDom.resolveExternals = False

    If oDom.loadXML(strEICXmlData) = True Then
         
        Set oNode = oDom.selectSingleNode("//MSH.9/CM_MSG_TYPE.2")
        If Not oNode Is Nothing Then
            Select Case UCase(Trim(oNode.Text))
                Case Is = "A28"
                    If enmDataSet = eDM_Demographics Then
                        lngResult = ReturnDemographicsRecordsets(oDom, lngUnitNumber, colRSObject, enmWhoseMessage)
                    End If
                Case Is = "A31"
                    If enmDataSet = eDM_Demographics Then
                        lngResult = ReturnDemographicsRecordsets(oDom, lngUnitNumber, colRSObject, enmWhoseMessage)
                    End If
                Case Is = "A04"
                    If enmDataSet = eDM_Encounters Then
                        If enmWhoseMessage = SAMS_Message Then
                        '<< Begin:SCR #29291;   Developer: Brian Mowbray 11/14/2002 11:38 AM
                            lngResult = CreateEncounterFromSAMSEIC(oDom, enmWhoseMessage)
                        ElseIf enmWhoseMessage = BMIST_Message Then
                            lngResult = CreateEncounterFromSAMSEIC(oDom, enmWhoseMessage)
                        '>> End: SCR #29291;
                        Else
                            lngResult = ReturnEncountersRecordsets(oDom, lngUnitNumber, colRSObject)
                        End If
                    End If
                Case Is = "ZA1"
                    'SCR 71415 - Import Allergy data from EIC.
                    lngResult = ReturnAllergyRecordsets(oDom, lngUnitNumber, colRSObject, enmWhoseMessage)
            End Select
        End If 'oNode is Nothing
    
    End If 'oDom.loadXML(strEICXmlData) = True
    
    GetRecordsetsFromEIC = lngResult
    
Exit Function

ErrHandler:
    If gobjShared Is Nothing Then
        Set gobjShared = New CWShared
    End If
    gobjShared.ShowVBError Err.Number, Err.Description, "UpdateFromEIC", "DMBE", vbExclamation
    Err.Clear
    GetRecordsetsFromEIC = DMErrorCodes.ERROR_EIC_READ_EIC_DATA
End Function

Private Function ReturnDemographicsRecordsets(ByVal oDom As MSXML.DOMDocument, ByVal lngUnitNumber As Long, ByRef colRSObject As Collection, ByVal enmWhoseMessage As XMLMessageSource) As Long

    Dim objRSObject As IDMBE.I_RSObjects
    Dim rsMMI_ID As Adodb.Recordset
    Dim rsMMI_NAME As Adodb.Recordset
    Dim rsMMI_SITE As Adodb.Recordset
    Dim rsMMI_GENERICDATA As Adodb.Recordset
    Dim rsMMI_GENERIC_ID As Adodb.Recordset '<SCR 22707 CC>
    Dim rsMMI_INST_NAME As Adodb.Recordset
    Dim rsMMI_ADDRESS As Adodb.Recordset
    Dim rsMMI_PHONE As Adodb.Recordset
    Dim rsALLERGIES As Adodb.Recordset
    Dim rsALLERGYREACTION As Adodb.Recordset
    Dim rsALLERGYCOMMENT As Adodb.Recordset
    
    Dim intI As Integer
    Dim oNode As MSXML.IXMLDOMElement
    Dim oItemNode As MSXML.IXMLDOMElement
    
    
    On Error Resume Next
    
    If Not colRSObject Is Nothing Then
        For intI = 1 To colRSObject.Count
            colRSObject.Remove intI
        Next intI
    End If
       
    On Error GoTo ErrHandler
    
    'Intitialize Demographics recordsets
    
    Set rsMMI_ID = New Adodb.Recordset
    Set rsMMI_NAME = New Adodb.Recordset
    Set rsMMI_SITE = New Adodb.Recordset
    Set rsMMI_GENERICDATA = New Adodb.Recordset
    Set rsMMI_GENERIC_ID = New Adodb.Recordset '<SCR 22707 CC>
    Set rsMMI_INST_NAME = New Adodb.Recordset
    Set rsMMI_ADDRESS = New Adodb.Recordset
    Set rsMMI_PHONE = New Adodb.Recordset
    Set rsALLERGIES = New Adodb.Recordset
    Set rsALLERGYREACTION = New Adodb.Recordset
    Set rsALLERGYCOMMENT = New Adodb.Recordset
    
    Call GetFieldProps("MMI_ID", rsMMI_ID)
    Call GetFieldProps("MMI_NAME", rsMMI_NAME)
    Call GetFieldProps("MMI_SITE", rsMMI_SITE)
    Call GetFieldProps("MMI_GENERICDATA", rsMMI_GENERICDATA)
    Call GetFieldProps("MMI_GENERIC_ID", rsMMI_GENERIC_ID) '<SCR 22707 CC>
    Call GetFieldProps("MMI_INST_NAME", rsMMI_INST_NAME)
    Call GetFieldProps("MMI_ADDRESS", rsMMI_ADDRESS)
    Call GetFieldProps("MMI_PHONE", rsMMI_PHONE)
    Call GetFieldProps("ALLERGIES", rsALLERGIES)
    Call GetFieldProps("ALLERGYREACTION", rsALLERGYREACTION)
    Call GetFieldProps("ALLERGYCOMMENT", rsALLERGYCOMMENT)
    
    '<SCR 22764 CC>
    Set oNode = oDom.selectSingleNode("//PID")
    If Not oNode Is Nothing Then
        If FillRecordsetFromPIDNode(oNode, lngUnitNumber, rsMMI_ID, rsMMI_NAME, rsMMI_GENERIC_ID, rsMMI_ADDRESS, rsMMI_PHONE, enmWhoseMessage) <> ERROR_SUCCESS Then
            ReturnDemographicsRecordsets = DMErrorCodes.ERROR_EIC_HEADER_DATA_NOT_FOUND
        End If
    Else
        ReturnDemographicsRecordsets = DMErrorCodes.ERROR_EIC_HEADER_DATA_NOT_FOUND
        Exit Function
    End If
    
    Set oNode = oDom.selectSingleNode("//ADT_A28")
    
    If Not oNode Is Nothing Then
        For Each oItemNode In oNode.childNodes
            Select Case UCase(oItemNode.nodeName)
                Case "ZD1"
                    Call FillRecordsetFromZD1Node(oItemNode, lngUnitNumber, rsMMI_ID, rsMMI_GENERICDATA, enmWhoseMessage)
                Case "ZPI"
                    Call FillRecordsetFromZPINode(oItemNode, lngUnitNumber, rsMMI_ID, rsMMI_NAME, rsMMI_SITE, rsMMI_GENERIC_ID, rsMMI_ADDRESS)
                Case "ZPG"
                    Call FillRecordsetFromZPGNode(oItemNode, lngUnitNumber, rsMMI_GENERICDATA)
                Case "ZPT"
                    Call FillRecordsetFromZPTNode(oItemNode, lngUnitNumber, rsMMI_INST_NAME)
                Case "AL1"
                    Call FillRecordsetFromAL1Node(oItemNode, lngUnitNumber, rsALLERGIES, rsALLERGYREACTION, rsALLERGYCOMMENT)
            End Select
        Next
    End If
    
    ' SCR #44193;   Developer: MAG 01/29/2004
    '>> Begin:SCR #49201;   Developer: Brian Mowbray 02/05/2004 01:06 PM - Added BMIST_Message logic
    ' SCR #53557, 53599; Developer: DBM 05/11/2004 - Check to see if Patient_Flag is set otherwise set it to "Y"
'    If enmWhoseMessage = SAMS_Message Or enmWhoseMessage = BMIST_Message Or bA28Demo Then
      If rsMMI_ID.RecordCount > 0 Then
        rsMMI_ID.MoveFirst
        Do While Not rsMMI_ID.EOF
          rsMMI_ID.Update "PATIENT_FLAG", "Y"
          rsMMI_ID.MoveNext
        Loop
      End If
'    End If
    
    '</SCR 22764 CC>
    
    'Append all the recordsets when all the above transactions are successful
    
    Set colRSObject = New Collection
    
     Set objRSObject = New RSObject_Ops
    objRSObject.TableName = "MMI_ID"
    Set objRSObject.DBRECORDSET = MoveRecordsetFirst(rsMMI_ID)
    colRSObject.Add objRSObject
    
    Set objRSObject = New RSObject_Ops
    objRSObject.TableName = "MMI_NAME"
    Set objRSObject.DBRECORDSET = MoveRecordsetFirst(rsMMI_NAME)
    colRSObject.Add objRSObject
    
    Set objRSObject = New RSObject_Ops
    objRSObject.TableName = "MMI_SITE"
    Set objRSObject.DBRECORDSET = MoveRecordsetFirst(rsMMI_SITE)
    colRSObject.Add objRSObject
    
    Set objRSObject = New RSObject_Ops
    objRSObject.TableName = "MMI_GENERICDATA"
    Set objRSObject.DBRECORDSET = MoveRecordsetFirst(rsMMI_GENERICDATA)
    colRSObject.Add objRSObject
    
    '<SCR 22707 CC>
    Set objRSObject = New RSObject_Ops
    objRSObject.TableName = "MMI_GENERIC_ID"
    Set objRSObject.DBRECORDSET = MoveRecordsetFirst(rsMMI_GENERIC_ID)
    colRSObject.Add objRSObject
    
    Set objRSObject = New RSObject_Ops
    objRSObject.TableName = "MMI_INST_NAME"
    Set objRSObject.DBRECORDSET = MoveRecordsetFirst(rsMMI_INST_NAME)
    colRSObject.Add objRSObject
    
    Set objRSObject = New RSObject_Ops
    objRSObject.TableName = "MMI_ADDRESS"
    Set objRSObject.DBRECORDSET = MoveRecordsetFirst(rsMMI_ADDRESS)
    colRSObject.Add objRSObject
    
    Set objRSObject = New RSObject_Ops
    objRSObject.TableName = "MMI_PHONE"
    Set objRSObject.DBRECORDSET = MoveRecordsetFirst(rsMMI_PHONE)
    colRSObject.Add objRSObject
    
    ReturnDemographicsRecordsets = DMErrorCodes.ERROR_SUCCESS
    
Exit Function

ErrHandler:
    ReturnDemographicsRecordsets = DMErrorCodes.ERROR_EIC_XML_TO_RECORDSET
    If gobjShared Is Nothing Then
        Set gobjShared = New CWShared
    End If
    gobjShared.ShowVBError Err.Number, Err.Description, "ReturnDemographicsRecordsets Function", "DMBE", vbExclamation
End Function

Private Function MoveRecordsetFirst(ByRef rsToMove As Adodb.Recordset) As Adodb.Recordset

    If Not rsToMove Is Nothing Then
        If rsToMove.RecordCount > 0 Then
            rsToMove.MoveFirst
        End If
    End If
    Set MoveRecordsetFirst = rsToMove
    
End Function

'<SCR 2276 CC>

Public Function FillRecordsetFromAL1Node(ByVal oAL1Node As MSXML.IXMLDOMElement, ByVal lngUnitNumber As Long, _
                                        ByRef rsALLERGIES As Adodb.Recordset, ByRef rsALLERGYREACTION As Adodb.Recordset, ByRef rsALLERGYCOMMENT As Adodb.Recordset) As Long
    
    Dim oItemNode As MSXML.IXMLDOMNode
    Dim oChildNode As MSXML.IXMLDOMNode
    Dim oChildNode1 As MSXML.IXMLDOMNode
    Dim strSQL As String
    Dim rsAllergy As New Adodb.Recordset
    Dim rsReaction As New Adodb.Recordset
    Dim strTemp As String
    Dim lngAllergenNCID As Long
    
    On Error GoTo ErrHandler
    
    If oAL1Node Is Nothing Then
        Exit Function
    End If
    
    If m_objGEMSDAS Is Nothing Then Set m_objGEMSDAS = oDatabase.CHCSII_DAS(Auto) 'New GEMS_DAS
    
    'lngFakeDataID = oSQL.GetNextID
    
    If rsALLERGIES Is Nothing Then
        Set rsALLERGIES = New Adodb.Recordset
        Call GetFieldProps("ALLERGIES", rsALLERGIES)
    End If
    
    If rsALLERGYREACTION Is Nothing Then
        Set rsALLERGYREACTION = New Adodb.Recordset
         Call GetFieldProps("ALLERGYREACTION", rsALLERGYREACTION)
    End If
    
    If rsALLERGYCOMMENT Is Nothing Then
        Set rsALLERGYCOMMENT = New Adodb.Recordset
         Call GetFieldProps("ALLERGYCOMMENT", rsALLERGYCOMMENT)
    End If
    
    rsALLERGIES.AddNew
    rsALLERGIES.Fields("UNIT_NUMBER").Value = lngUnitNumber
    
    'Need to fill-in Unit Number for each RS
    'DataID is the PK of DIAGNOSES
    
    For Each oItemNode In oAL1Node.childNodes
        Select Case UCase(oItemNode.nodeName)
            Case "AL1.2"
                
            Case "AL1.3"
            
                Dim blnNCIDFound As Boolean
                
                blnNCIDFound = False
                
                Set oChildNode = oItemNode.selectSingleNode("CE.6")
                If Not oChildNode Is Nothing Then
                    If oChildNode.Text = "99NCI" Then
                        Set oChildNode = oItemNode.selectSingleNode("CE.4")
                        If Not oChildNode Is Nothing Then
                            lngAllergenNCID = XMLToOracleNumber(oChildNode.Text)
                            rsALLERGIES.Fields("ALLERGENNCID").Value = lngAllergenNCID
                            blnNCIDFound = True
                        End If
                        Set oChildNode = oItemNode.selectSingleNode("CE.5")
                        If Not oChildNode Is Nothing Then
                            rsALLERGIES.Fields("ALLERGEN").Value = NullToNothing(oChildNode.Text)
                        End If
                        
                    End If
                End If
                
                If blnNCIDFound = False Then
            
                    Set oChildNode = oItemNode.selectSingleNode("CE.2")
                    If Not oChildNode Is Nothing Then
    
                        strTemp = NullToNothing(oChildNode.Text)
                        If Not strTemp = "" Then
                            strSQL = "SELECT IEN FROM CHCS_ALLERGIES WHERE BN_GNN = '" & oChildNode.Text & "'"
                            Set rsAllergy = m_objGEMSDAS.ExecuteSQL(strSQL)
                            If Not rsAllergy.EOF Then
                                lngAllergenNCID = rsAllergy.Fields(0).Value
                                rsALLERGIES.Fields("ALLEGENNCID").Value = rsAllergy.Fields(0).Value
                                rsALLERGIES.Fields("ALLEGEN").Value = strTemp
                            Else
                                lngAllergenNCID = 998000
                                rsALLERGIES.Fields("ALLEGENNCID").Value = 998000 'OTHER
                                rsALLERGIES.Fields("ALLEGEN").Value = strTemp
                            End If
                        End If
                    End If
                End If
                        
            Case "AL1.5.LST"
                For Each oChildNode In oItemNode.childNodes
                    
                    blnNCIDFound = False
                    
                    If oChildNode.nodeName = "AL1.5" Then
                        
                        Set oChildNode1 = oChildNode.selectSingleNode("CE.6")
                        If Not oChildNode1 Is Nothing Then
                            If oChildNode1.Text = "99NCI" Then
                                Set oChildNode1 = oChildNode.selectSingleNode("CE.4")
                                If Not oChildNode1 Is Nothing Then
                                    lngAllergenNCID = XMLToOracleNumber(oChildNode1.Text)
                                    rsALLERGYREACTION.AddNew
                                        rsALLERGYREACTION.Fields("ALLERGENNCID").Value = lngAllergenNCID
                                        rsALLERGYREACTION.Fields("UNIT_NUMBER").Value = lngUnitNumber
                                        rsALLERGYREACTION.Fields("REACTIONNCID").Value = lngAllergenNCID
                                    blnNCIDFound = True
                                    Set oChildNode1 = oChildNode.selectSingleNode("CE.5")
                                    If Not oChildNode1 Is Nothing Then
                                        rsALLERGYREACTION.Fields("REACTION").Value = oChildNode1.Text
                                    End If
                                    rsALLERGYREACTION.Update
                                End If
                            End If
                        End If
                    
                        If blnNCIDFound = False Then
                            Set oChildNode1 = oChildNode.selectSingleNode("CE.2")
                            If Not oChildNode1 Is Nothing Then
                                strTemp = NullToNothing(oChildNode1.Text)
                                If Not strTemp = "" Then
                                    strSQL = "SELECT NCID FROM NCIDS WHERE domainncid = 68230 and upper(description) = '" & oChildNode1.Text & "'"
                                    rsALLERGYREACTION.AddNew
                                    rsALLERGYREACTION.Fields("ALLERGENNCID").Value = lngAllergenNCID
                                    rsALLERGYREACTION.Fields("UNIT_NUMBER").Value = lngUnitNumber
                                    Set rsReaction = m_objGEMSDAS.ExecuteSQL(strSQL)
                                    If Not rsReaction.EOF Then
                                        rsALLERGYREACTION.Fields("REACTIONNCID").Value = rsReaction.Fields(0).Value
                                        rsALLERGYREACTION.Fields("REACTION").Value = strTemp
                                    Else
                                        rsALLERGYREACTION.Fields("REACTIONNCID").Value = 76281 'Unknown
                                        rsALLERGYREACTION.Fields("REACTION").Value = strTemp
                                    End If
                                    rsALLERGYREACTION.Update
                                End If
                            End If
                        End If
                                
                    End If
                Next
            
            Case "AL1.16"
                Set oChildNode = oItemNode.selectSingleNode("//XCN.1")
                If Not oChildNode Is Nothing Then
                    rsALLERGIES.Fields("NOTEDATE").Value = XMLToOracleDate(oChildNode.Text)
                End If
                
        End Select
    
    Next oItemNode
    
    rsALLERGIES.Update
    
    FillRecordsetFromAL1Node = DMErrorCodes.ERROR_SUCCESS

Exit Function

ErrHandler:
   FillRecordsetFromAL1Node = DMErrorCodes.ERROR_EIC_XML_TO_RECORDSET
   If gobjShared Is Nothing Then
      Set gobjShared = New CWShared
   End If
   gobjShared.ShowVBError Err.Number, Err.Description, "FillRecordsetFromAL1Node Function", "DMBE", vbExclamation
   Resume Next 'There is no reason this procedure should not continue to create xml if it errors
   
End Function

'*****************************************
'SCR 71415 - Import Allergy data From EIC.
'*****************************************
Private Function ReturnAllergyRecordsets(ByVal oDom As MSXML.DOMDocument, ByVal lngUnitNumber As Long, ByRef colRSObject As Collection, ByVal enmWhoseMessage As XMLMessageSource) As DMErrorCodes

    Dim objRSObject As IDMBE.I_RSObjects
    Dim rsALLERGIES As Adodb.Recordset
    Dim rsALLERGYREACTION As Adodb.Recordset
    Dim rsALLERGYCOMMENT As Adodb.Recordset
    
    
    Dim intI As Integer
    Dim oNode As MSXML.IXMLDOMElement
    Dim oItemNode As MSXML.IXMLDOMElement
    Dim oItemSubNode As MSXML.IXMLDOMElement
    
    
    On Error Resume Next
    
    If Not colRSObject Is Nothing Then
        For intI = 1 To colRSObject.Count
            colRSObject.Remove intI
        Next intI
    End If
       
    On Error GoTo ErrHandler
    
    Set rsALLERGIES = New Adodb.Recordset
    Set rsALLERGYREACTION = New Adodb.Recordset
    Set rsALLERGYCOMMENT = New Adodb.Recordset
    
    Call GetFieldProps("ALLERGIES", rsALLERGIES)
    Call GetFieldProps("ALLERGYREACTION", rsALLERGYREACTION)
    Call GetFieldProps("ALLERGYCOMMENT", rsALLERGYCOMMENT)
    
    
    Set oNode = oDom.selectSingleNode("//ADT_ZA1")
    
    If Not oNode Is Nothing Then
        For Each oItemNode In oNode.childNodes
            Select Case UCase(oItemNode.nodeName)
                Case "ADT_ZA1.LST.1"
                  For Each oItemSubNode In oItemNode.childNodes
                    'If UCase(oItemSubNode.nodeName) = "AL1" Then
                    '  Call FillRecordsetFromAL1Node(oItemSubNode, lngUnitNumber, rsALLERGIES, rsALLERGYREACTION, rsALLERGYCOMMENT)
                    'End If
                    If UCase(oItemSubNode.nodeName) = "ZAL" Then
                      Call FillRecordsetFromZALNode(oItemSubNode, lngUnitNumber, rsALLERGIES, rsALLERGYREACTION, rsALLERGYCOMMENT)
                    End If
                  Next
                                    
                  Set colRSObject = New Collection
                  
              
                  Set objRSObject = New RSObject_Ops
                  objRSObject.TableName = "ALLERGIES"
                  Set objRSObject.DBRECORDSET = rsALLERGIES
                  colRSObject.Add objRSObject
              
                  Set objRSObject = New RSObject_Ops
                  objRSObject.TableName = "ALLERGYREACTION"
                  Set objRSObject.DBRECORDSET = rsALLERGYREACTION
                  colRSObject.Add objRSObject
              
                  Set objRSObject = New RSObject_Ops
                  objRSObject.TableName = "ALLERGYCOMMENT"
                  Set objRSObject.DBRECORDSET = rsALLERGYCOMMENT
                  colRSObject.Add objRSObject

            End Select
        Next
    End If
    
    ReturnAllergyRecordsets = DMErrorCodes.ERROR_SUCCESS
    
Exit Function

ErrHandler:
    ReturnAllergyRecordsets = DMErrorCodes.ERROR_EIC_XML_TO_RECORDSET
    If gobjShared Is Nothing Then
        Set gobjShared = New CWShared
    End If
    gobjShared.ShowVBError Err.Number, Err.Description, "ReturnAllergyRecordsets Function", "DMBE", vbExclamation

End Function


'</SCR 2276 CC>
Private Function ReturnEncountersRecordsets(ByVal oDom As MSXML.DOMDocument, ByVal lngUnitNumber As Long, ByRef colRSObject As Collection) As DMErrorCodes

    Dim objRSObject As IDMBE.I_RSObjects
    Dim oNodeList As MSXML.IXMLDOMNodeList
    Dim oNode As MSXML.IXMLDOMElement
    Dim oItemNode As MSXML.IXMLDOMNode
    Dim oADTNode As MSXML.IXMLDOMNode
    Dim oSQLOps As ICHCSII_SQl.ISqlOpsEx
    
    Dim rsENCOUNTERS As Adodb.Recordset
    Dim rsENC_DISPOSITION As Adodb.Recordset
    Dim rsDIAGNOSES As Adodb.Recordset
    Dim rsDIAGNOSESRELATIONSHIPS As Adodb.Recordset
    Dim rsENC_DIAGNOSIS As Adodb.Recordset
    Dim rsPROCEDURES As Adodb.Recordset
    Dim rsENC_SECTIONS As Adodb.Recordset
    Dim rsENCTEXTNOTES As Adodb.Recordset
    Dim rsPI_DATA As Adodb.Recordset
    Dim rsENC_RTFS As Adodb.Recordset
    Dim rsENC_SIGBLOCK As Adodb.Recordset
    Dim rsLISTTOOLRECORD As Adodb.Recordset
    Dim rsLISTTOOLRECORDS As Adodb.Recordset
    Dim rsCOMPLAINTS As Adodb.Recordset
    Dim rsPROBLEMS As Adodb.Recordset
    Dim rsPROBLEMSRELATIONSHIPS As Adodb.Recordset
    Dim rsCLINNOTE As Adodb.Recordset
    Dim rsSONOTES As Adodb.Recordset
    Dim rsWORK_STATUS As Adodb.Recordset
    Dim rsREADINESS As Adodb.Recordset
    Dim rsVITALS As Adodb.Recordset
    Dim rsVITALSIGNSMODIFIER As Adodb.Recordset
    Dim rsVITALSIGNSVALUE As Adodb.Recordset
    Dim rsMedications As Adodb.Recordset
    Dim rsRXLIST As Adodb.Recordset
    Dim rsProvider As Adodb.Recordset
    
    Dim enmResult As DMErrorCodes
    
    On Error GoTo ErrHandler
    
    'Intitialize Demographics recordsets
    
    Set rsENCOUNTERS = New Adodb.Recordset
    Set rsENC_DISPOSITION = New Adodb.Recordset
    Set rsDIAGNOSES = New Adodb.Recordset
    Set rsDIAGNOSESRELATIONSHIPS = New Adodb.Recordset
    Set rsENC_DIAGNOSIS = New Adodb.Recordset
    Set rsPROCEDURES = New Adodb.Recordset
    Set rsENC_SECTIONS = New Adodb.Recordset
    Set rsENCTEXTNOTES = New Adodb.Recordset
    Set rsPI_DATA = New Adodb.Recordset
    Set rsENC_RTFS = New Adodb.Recordset
    Set rsENC_SIGBLOCK = New Adodb.Recordset
    Set rsLISTTOOLRECORD = New Adodb.Recordset
    Set rsLISTTOOLRECORDS = New Adodb.Recordset
    Set rsCOMPLAINTS = New Adodb.Recordset
    Set rsPROBLEMS = New Adodb.Recordset
    Set rsPROBLEMSRELATIONSHIPS = New Adodb.Recordset
    Set rsCLINNOTE = New Adodb.Recordset
    Set rsSONOTES = New Adodb.Recordset
    Set rsWORK_STATUS = New Adodb.Recordset
    Set rsREADINESS = New Adodb.Recordset
    Set rsVITALS = New Adodb.Recordset
    Set rsVITALSIGNSMODIFIER = New Adodb.Recordset
    Set rsVITALSIGNSVALUE = New Adodb.Recordset
    Set rsMedications = New Adodb.Recordset
    Set rsRXLIST = New Adodb.Recordset
    Set rsProvider = New Adodb.Recordset
    
    Call GetFieldProps("ENCOUNTERS", rsENCOUNTERS)
    Call GetFieldProps("ENC_DISPOSITION", rsENC_DISPOSITION)
    Call GetFieldProps("DIAGNOSES", rsDIAGNOSES)
    Call GetFieldProps("DIAGNOSESRELATIONSHIPS", rsDIAGNOSESRELATIONSHIPS)
    Call GetFieldProps("ENC_DIAGNOSIS", rsENC_DIAGNOSIS)
    Call GetFieldProps("PROCEDURES", rsPROCEDURES)
    Call GetFieldProps("ENC_SECTIONS", rsENC_SECTIONS)
    Call GetFieldProps("ENCTEXTNOTES", rsENCTEXTNOTES)
    Call GetFieldProps("PI_DATA", rsPI_DATA)
    Call GetFieldProps("ENC_RTFS", rsENC_RTFS)
    Call GetFieldProps("ENC_SIGBLOCK", rsENC_SIGBLOCK)
    Call GetFieldProps("LISTTOOLRECORD", rsLISTTOOLRECORD)
    Call GetFieldProps("LISTTOOLRECORDS", rsLISTTOOLRECORDS)
    Call GetFieldProps("COMPLAINTS", rsCOMPLAINTS)
    Call GetFieldProps("PROBLEMS", rsPROBLEMS)
    Call GetFieldProps("PROBLEMSRELATIONSHIPS", rsPROBLEMSRELATIONSHIPS)
    Call GetFieldProps("CLINNOTE", rsCLINNOTE)
    Call GetFieldProps("SONOTES", rsSONOTES)
    Call GetFieldProps("WORK_STATUS", rsWORK_STATUS)
    Call GetFieldProps("READINESS", rsREADINESS)
    Call GetFieldProps("VITALS", rsVITALS)
    Call GetFieldProps("VITALSIGNSMODIFIER", rsVITALSIGNSMODIFIER)
    Call GetFieldProps("VITALSIGNSVALUE", rsVITALSIGNSVALUE)
    Call GetFieldProps("MEDICATIONS", rsMedications)
    Call GetFieldProps("RX_LIST", rsRXLIST)
    Call GetFieldProps("PROVIDER", rsProvider)
    
    Set oADTNode = oDom.selectSingleNode("//ADT_A01")
    
    Set oNodeList = oDom.selectNodes("//ZPV")
    If oNodeList.Length = 0 Then
        ReturnEncountersRecordsets = DMErrorCodes.ERROR_EIC_HEADER_DATA_NOT_FOUND
        Exit Function
    End If
    'Set Node to "ZPV" Node
    Set oNode = oNodeList.Item(0)
    If oNode Is Nothing Then
        ReturnEncountersRecordsets = DMErrorCodes.ERROR_EIC_HEADER_DATA_NOT_FOUND
        Exit Function
    End If
    
    ReturnEncountersRecordsets = FillRecordsetFromZPVNode(oNode, lngUnitNumber, rsENCOUNTERS, rsENC_DISPOSITION, rsProvider)
    ' Begin:SCR #44179;   Developer: Mag 10/22/2003 06:33 PM
    If ReturnEncountersRecordsets <> DMErrorCodes.ERROR_SUCCESS Then
      Exit Function
    End If
    ' End: SCR #44179;
    
    If Not oADTNode Is Nothing Then
        For Each oItemNode In oADTNode.childNodes
        
            enmResult = DMErrorCodes.ERROR_SUCCESS
        
            Select Case UCase(oItemNode.nodeName)
            Case "ZG1"
                'Call FillRecordsetFromZG1Node(oItemNode, lngUnitNumber, rsENC_DISPOSITION)
                'Does not have dataID
            Case "ADT_A01.LST.5"
                'Call FillRecordsetFromDG1Node(oItemNode, lngUnitNumber, rsDIAGNOSES)
                'Does not have dataID
            Case "ZD2"
                enmResult = FillRecordsetFromZD2Node(oItemNode, lngUnitNumber, rsDIAGNOSES)
                If enmResult <> ERROR_SUCCESS Then
                    GoTo ErrHandler
                End If
            Case "ZDR"
                enmResult = FillRecordsetFromZDRNode(oItemNode, lngUnitNumber, rsDIAGNOSESRELATIONSHIPS)
                If enmResult <> ERROR_SUCCESS Then
                    GoTo ErrHandler
                End If
            Case "ZED"
                enmResult = FillRecordsetFromZEDNode(oItemNode, lngUnitNumber, rsENC_DIAGNOSIS)
                If enmResult <> ERROR_SUCCESS Then
                    GoTo ErrHandler
                End If
            
'            'NGIT SPR 772: Darren Magnuson
'            Case "ADT_A01.LST.7" 'ADT_A01.GRP.1
'                'Call FillRecordsetFromPR1Node(oItemNode, lngUnitNumber, rsPROCEDURES)
'                'Does not have dataID
'
'                '<SCR 25123 CC>
'
'                Dim oNodeLevel1 As MSXML.IXMLDOMElement
'
'                Set oNodeLevel1 = oItemNode.selectSingleNode("//ADT_A01.GRP.1/ZP1")
'                enmResult = FillRecordsetFromZP1Node(oNodeLevel1, lngUnitNumber, rsPROCEDURES)
'                If enmResult <> ERROR_SUCCESS Then
'                    GoTo ErrHandler
'                End If
'            'End: NGIT SPR 772
            
            Case "ZP1"
                enmResult = FillRecordsetFromZP1Node(oItemNode, lngUnitNumber, rsPROCEDURES)
                If enmResult <> ERROR_SUCCESS Then
                    GoTo ErrHandler
                End If
            Case "ZES"
                enmResult = FillRecordsetFromZESNode(oItemNode, lngUnitNumber, rsENC_SECTIONS)
                If enmResult <> ERROR_SUCCESS Then
                    GoTo ErrHandler
                End If
            Case "ZEX"
                enmResult = FillRecordsetFromZEXNode(oItemNode, lngUnitNumber, rsENCTEXTNOTES)
                If enmResult <> ERROR_SUCCESS Then
                    GoTo ErrHandler
                End If
            Case "ZEP"
                enmResult = FillRecordsetFromZEPNode(oItemNode, lngUnitNumber, rsPI_DATA)
                If enmResult <> ERROR_SUCCESS Then
                    GoTo ErrHandler
                End If
            Case "ZER"
                enmResult = FillRecordsetFromZERNode(oItemNode, lngUnitNumber, rsENC_RTFS)
                If enmResult <> ERROR_SUCCESS Then
                    GoTo ErrHandler
                End If
            Case "ZEB"
                enmResult = FillRecordsetFromZEBNode(oItemNode, lngUnitNumber, rsENC_SIGBLOCK)
                If enmResult <> ERROR_SUCCESS Then
                    GoTo ErrHandler
                End If
            Case "ZLT"
                enmResult = FillRecordsetFromZLTNode(oItemNode, lngUnitNumber, rsLISTTOOLRECORD)
                If enmResult <> ERROR_SUCCESS Then
                    GoTo ErrHandler
                End If
            Case "ZLR"
                enmResult = FillRecordsetFromZLRNode(oItemNode, lngUnitNumber, rsLISTTOOLRECORDS)
                If enmResult <> ERROR_SUCCESS Then
                    GoTo ErrHandler
                End If
            Case "ZCP"
                enmResult = FillRecordsetFromZCPNode(oItemNode, lngUnitNumber, rsCOMPLAINTS)
                If enmResult <> ERROR_SUCCESS Then
                    GoTo ErrHandler
                End If
            Case "ZPB"
                enmResult = FillRecordsetFromZPBNode(oItemNode, lngUnitNumber, rsPROBLEMS)
                If enmResult <> ERROR_SUCCESS Then
                    GoTo ErrHandler
                End If
            Case "ZBR"
                enmResult = FillRecordsetFromZBRNode(oItemNode, lngUnitNumber, rsPROBLEMSRELATIONSHIPS)
                If enmResult <> ERROR_SUCCESS Then
                    GoTo ErrHandler
                End If
            Case "ZCN"
                enmResult = FillRecordsetFromZCNNode(oItemNode, lngUnitNumber, rsCLINNOTE)
                If enmResult <> ERROR_SUCCESS Then
                    GoTo ErrHandler
                End If
            Case "ZSO"
                enmResult = FillRecordsetFromZSONode(oItemNode, lngUnitNumber, rsSONOTES)
                If enmResult <> ERROR_SUCCESS Then
                    GoTo ErrHandler
                End If
            Case "ZWS"
                enmResult = FillRecordsetFromZWSNode(oItemNode, lngUnitNumber, rsWORK_STATUS)
                If enmResult <> ERROR_SUCCESS Then
                    GoTo ErrHandler
                End If
            Case "ZRD"
                enmResult = FillRecordsetFromZRDNode(oItemNode, lngUnitNumber, rsREADINESS)
                If enmResult <> ERROR_SUCCESS Then
                    GoTo ErrHandler
                End If
            Case "ZVT"
                enmResult = FillRecordsetFromZVTNode(oItemNode, lngUnitNumber, rsVITALS, rsVITALSIGNSMODIFIER, rsVITALSIGNSVALUE)
                If enmResult <> ERROR_SUCCESS Then
                    GoTo ErrHandler
                End If
            ' Begin:SCR #28326;   Developer: MAG 10/16/2002 12:36 PM
            Case "RDE_O01.LST.14"   'Medications Group
                enmResult = FillRecordsetFromRDE_O01Node(oItemNode, lngUnitNumber, rsMedications, rsRXLIST)
                If enmResult <> ERROR_SUCCESS Then
                    GoTo ErrHandler
                End If
             ' End: SCR #28326;
             End Select
    
        
        Next oItemNode
    End If
    
    Set colRSObject = New Collection
    
    Set objRSObject = New RSObject_Ops
    objRSObject.TableName = "ENCOUNTERS"
    Set objRSObject.DBRECORDSET = MoveRecordsetFirst(rsENCOUNTERS)
    colRSObject.Add objRSObject, objRSObject.TableName
    
    Set objRSObject = New RSObject_Ops
    objRSObject.TableName = "ENC_DISPOSITION"
    Set objRSObject.DBRECORDSET = MoveRecordsetFirst(rsENC_DISPOSITION)
    colRSObject.Add objRSObject, objRSObject.TableName
    
    Set objRSObject = New RSObject_Ops
    objRSObject.TableName = "DIAGNOSES"
    Set objRSObject.DBRECORDSET = MoveRecordsetFirst(rsDIAGNOSES)
    colRSObject.Add objRSObject, objRSObject.TableName
    
    Set objRSObject = New RSObject_Ops
    objRSObject.TableName = "DIAGNOSESRELATIONSHIPS"
    Set objRSObject.DBRECORDSET = MoveRecordsetFirst(rsDIAGNOSESRELATIONSHIPS)
    colRSObject.Add objRSObject, objRSObject.TableName
    
    Set objRSObject = New RSObject_Ops
    objRSObject.TableName = "ENC_DIAGNOSIS"
    Set objRSObject.DBRECORDSET = MoveRecordsetFirst(rsENC_DIAGNOSIS)
    colRSObject.Add objRSObject, objRSObject.TableName
    
    Set objRSObject = New RSObject_Ops
    objRSObject.TableName = "PROCEDURES"
    Set objRSObject.DBRECORDSET = MoveRecordsetFirst(rsPROCEDURES)
    colRSObject.Add objRSObject, objRSObject.TableName
    
    Set objRSObject = New RSObject_Ops
    objRSObject.TableName = "ENC_SECTIONS"
    Set objRSObject.DBRECORDSET = MoveRecordsetFirst(rsENC_SECTIONS)
    colRSObject.Add objRSObject, objRSObject.TableName
    
    Set objRSObject = New RSObject_Ops
    objRSObject.TableName = "ENCTEXTNOTES"
    Set objRSObject.DBRECORDSET = MoveRecordsetFirst(rsENCTEXTNOTES)
    colRSObject.Add objRSObject, objRSObject.TableName
    
    Set objRSObject = New RSObject_Ops
    objRSObject.TableName = "PI_DATA"
    Set objRSObject.DBRECORDSET = MoveRecordsetFirst(rsPI_DATA)
    colRSObject.Add objRSObject, objRSObject.TableName
    
    Set objRSObject = New RSObject_Ops
    objRSObject.TableName = "ENC_RTFS"
    Set objRSObject.DBRECORDSET = MoveRecordsetFirst(rsENC_RTFS)
    colRSObject.Add objRSObject, objRSObject.TableName
    
    Set objRSObject = New RSObject_Ops
    objRSObject.TableName = "ENC_SIGBLOCK"
    Set objRSObject.DBRECORDSET = MoveRecordsetFirst(rsENC_SIGBLOCK)
    colRSObject.Add objRSObject, objRSObject.TableName
    
    Set objRSObject = New RSObject_Ops
    objRSObject.TableName = "LISTTOOLRECORD"
    Set objRSObject.DBRECORDSET = MoveRecordsetFirst(rsLISTTOOLRECORD)
    colRSObject.Add objRSObject, objRSObject.TableName
    
    Set objRSObject = New RSObject_Ops
    objRSObject.TableName = "LISTTOOLRECORDS"
    Set objRSObject.DBRECORDSET = MoveRecordsetFirst(rsLISTTOOLRECORDS)
    colRSObject.Add objRSObject, objRSObject.TableName
    
    Set objRSObject = New RSObject_Ops
    objRSObject.TableName = "COMPLAINTS"
    Set objRSObject.DBRECORDSET = MoveRecordsetFirst(rsCOMPLAINTS)
    colRSObject.Add objRSObject, objRSObject.TableName
    
    ' Begin:SCR #28326;   Developer: MAG 10/16/2002 12:36 PM
    Set objRSObject = New RSObject_Ops
    objRSObject.TableName = "MEDICATIONS"
    Set objRSObject.DBRECORDSET = MoveRecordsetFirst(rsMedications)
    colRSObject.Add objRSObject, objRSObject.TableName
    ' End: SCR #28326;
    
    ' Begin:SCR #47403;   Developer:  01/02/2004
    Set objRSObject = New RSObject_Ops
    objRSObject.TableName = "RX_LIST"
    Set objRSObject.DBRECORDSET = MoveRecordsetFirst(rsRXLIST)
    colRSObject.Add objRSObject, objRSObject.TableName
    ' End: SCR #47403;
    
    Set objRSObject = New RSObject_Ops
    objRSObject.TableName = "PROBLEMS"
    Set objRSObject.DBRECORDSET = MoveRecordsetFirst(rsPROBLEMS)
    colRSObject.Add objRSObject, objRSObject.TableName
    
    Set objRSObject = New RSObject_Ops
    objRSObject.TableName = "PROBLEMSRELATIONSHIPS"
    Set objRSObject.DBRECORDSET = MoveRecordsetFirst(rsPROBLEMSRELATIONSHIPS)
    colRSObject.Add objRSObject, objRSObject.TableName
    
    Set objRSObject = New RSObject_Ops
    objRSObject.TableName = "CLINNOTE"
    Set objRSObject.DBRECORDSET = MoveRecordsetFirst(rsCLINNOTE)
    colRSObject.Add objRSObject, objRSObject.TableName
    
    Set objRSObject = New RSObject_Ops
    objRSObject.TableName = "SONOTES"
    Set objRSObject.DBRECORDSET = MoveRecordsetFirst(rsSONOTES)
    colRSObject.Add objRSObject, objRSObject.TableName
    
    Set objRSObject = New RSObject_Ops
    objRSObject.TableName = "WORK_STATUS"
    Set objRSObject.DBRECORDSET = MoveRecordsetFirst(rsWORK_STATUS)
    colRSObject.Add objRSObject, objRSObject.TableName
    
    Set objRSObject = New RSObject_Ops
    objRSObject.TableName = "READINESS"
    Set objRSObject.DBRECORDSET = MoveRecordsetFirst(rsREADINESS)
    colRSObject.Add objRSObject, objRSObject.TableName
    
    Set objRSObject = New RSObject_Ops
    objRSObject.TableName = "VITALS"
    Set objRSObject.DBRECORDSET = MoveRecordsetFirst(rsVITALS)
    colRSObject.Add objRSObject, objRSObject.TableName
    
    Set objRSObject = New RSObject_Ops
    objRSObject.TableName = "VITALSIGNSMODIFIER"
    Set objRSObject.DBRECORDSET = MoveRecordsetFirst(rsVITALSIGNSMODIFIER)
    colRSObject.Add objRSObject, objRSObject.TableName
    
    Set objRSObject = New RSObject_Ops
    objRSObject.TableName = "VITALSIGNSVALUE"
    Set objRSObject.DBRECORDSET = MoveRecordsetFirst(rsVITALSIGNSVALUE)
    colRSObject.Add objRSObject, objRSObject.TableName

    If Not rsProvider Is Nothing Then
      MoveRecordsetFirst rsProvider
      Set oSQLOps = oDatabase.CHCSII_SQLOPS_EX(Auto)
      Call oSQLOps.ValidateProvider2(rsProvider("NCID"), rsProvider("SSN"), "", "", rsProvider("NAME") & vbNullString, True)
      
      Set objRSObject = New RSObject_Ops
      objRSObject.TableName = "PROVIDER"
      Set objRSObject.DBRECORDSET = MoveRecordsetFirst(rsProvider)
      colRSObject.Add objRSObject, objRSObject.TableName
    End If
    
    ReturnEncountersRecordsets = ERROR_SUCCESS

Exit Function

ErrHandler:
    If Err.Number <> 0 Then
        If gobjShared Is Nothing Then
            Set gobjShared = New CWShared
        End If
        gobjShared.ShowVBError Err.Number, Err.Description, "ReturnEncountersRecordsets Function", "DMBE", vbExclamation
        Err.Clear
        ReturnEncountersRecordsets = DMErrorCodes.ERROR_EIC_XML_TO_RECORDSET
    Else
        ReturnEncountersRecordsets = enmResult
    End If
End Function

Private Function FillRecordsetFromZWSNode(ByVal oZWSNode As MSXML.IXMLDOMElement, ByVal lngUnitNumber As Long, _
                                        ByRef rsWORK_STATUS As Adodb.Recordset) As DMErrorCodes

    Dim oItemNode As MSXML.IXMLDOMNode
    Dim oChildNode As MSXML.IXMLDOMNode
    
    On Error GoTo ErrHandler
    
    If oZWSNode Is Nothing Then
        Exit Function
    End If
    
    rsWORK_STATUS.AddNew

    For Each oItemNode In oZWSNode.childNodes
        Select Case UCase(oItemNode.nodeName)
            Case "ZWS.1"
                
            Case "ZWS.2"
                rsWORK_STATUS.Fields("Encounter_Number").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZWS.3"
                'rsWORK_STATUS.Fields("Facility_ID").Value = XMLToOracleNumber(oItemNode.Text)
                rsWORK_STATUS.Fields("Facility_ID").Value = gobjLogon.FacilityNCID
            Case "ZWS.4"
                rsWORK_STATUS.Fields("Patient_NCID").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZWS.5"
                Set oChildNode = oItemNode.selectSingleNode("//XCN.1")
                If Not oChildNode Is Nothing Then
                    'Returning SSN
                    rsWORK_STATUS.Fields("Provider_NCID").Value = XMLToOracleNumber(oChildNode.Text)
                End If
            Case "ZWS.6"
                Set oChildNode = oItemNode.selectSingleNode("CE.5")
                If Not oChildNode Is Nothing Then
                    rsWORK_STATUS.Fields("Work_Status_NCID").Value = XMLToOracleNumber(oChildNode.Text)
                End If
            Case "ZWS.7"
                rsWORK_STATUS.Fields("Work_Status_Selected").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZWS.8"
                rsWORK_STATUS.Fields("Work_Status_Qualified").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZWS.9"
                rsWORK_STATUS.Fields("LAST_UPDATED").Value = XMLToOracleDate(oItemNode.Text)
        End Select
    Next oItemNode
    
    rsWORK_STATUS.Update
    
    FillRecordsetFromZWSNode = DMErrorCodes.ERROR_SUCCESS

Exit Function

ErrHandler:
    FillRecordsetFromZWSNode = DMErrorCodes.ERROR_EIC_XML_TO_RECORDSET
    If gobjShared Is Nothing Then
        Set gobjShared = New CWShared
    End If
    gobjShared.ShowVBError Err.Number, Err.Description, "FillRecordsetFromZWSNode Function", "DMBE", vbExclamation
    Resume Next 'There is no reason this procedure should not continue to create xml if it errors
End Function



Private Function FillRecordsetFromZSONode(ByVal oZSONode As MSXML.IXMLDOMElement, ByVal lngUnitNumber As Long, _
                                        ByRef rsSONOTES As Adodb.Recordset) As DMErrorCodes

    Dim oItemNode As MSXML.IXMLDOMNode
    Dim oChildNode As MSXML.IXMLDOMNode
    
    On Error GoTo ErrHandler
    
    If oZSONode Is Nothing Then
        Exit Function
    End If
    
    rsSONOTES.AddNew

    For Each oItemNode In oZSONode.childNodes
        Select Case UCase(oItemNode.nodeName)
            Case "ZSO.1"
                
            Case "ZSO.2"
                rsSONOTES.Fields("Unit_Number").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZSO.3"
                rsSONOTES.Fields("Encounter").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZSO.4"
                rsSONOTES.Fields("DataID").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZSO.5"
                rsSONOTES.Fields("Type").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZSO.6"
                rsSONOTES.Fields("EncounterDataID").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZSO.7"
                rsSONOTES.Fields("DTS").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZSO.8"
                For Each oChildNode In oItemNode.childNodes
                    Select Case oChildNode.nodeName
                        Case "CE.1"
                            rsSONOTES.Fields("OwnerNCID").Value = XMLToOracleNumber(oChildNode.Text)
                        Case "CE.2"
                            rsSONOTES.Fields("OwnerName").Value = NullToNothing(oChildNode.Text)
                    End Select
                Next
            Case "ZSO.9"
                rsSONOTES.Fields("Title").Value = NullToNothing(oItemNode.Text)
            Case "ZSO.10"
                rsSONOTES.Fields("RTF").Value = NullToNothing(oItemNode.Text)
            Case "ZSO.11"
                rsSONOTES.Fields("UpdateFlag").Value = NullToNothing(oItemNode.Text)
            Case "ZSO.12"
                rsSONOTES.Fields("ProfileNCID").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZSO.13"
                rsSONOTES.Fields("UpdatedBy").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZSO.14"
                rsSONOTES.Fields("UpdatedOn").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZSO.15"
                rsSONOTES.Fields("CreatedBy").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZSO.16"
                rsSONOTES.Fields("CreatedOn").Value = XMLToOracleDate(oItemNode.Text)
        End Select
    Next oItemNode
    
    rsSONOTES.Update
    
    FillRecordsetFromZSONode = DMErrorCodes.ERROR_SUCCESS

Exit Function

ErrHandler:
    FillRecordsetFromZSONode = DMErrorCodes.ERROR_EIC_XML_TO_RECORDSET
    If gobjShared Is Nothing Then
        Set gobjShared = New CWShared
    End If
    gobjShared.ShowVBError Err.Number, Err.Description, "FillRecordsetFromZSONode Function", "DMBE", vbExclamation
    Resume Next 'There is no reason this procedure should not continue to create xml if it errors
End Function


Private Function FillRecordsetFromZCNNode(ByVal oZCNNode As MSXML.IXMLDOMElement, ByVal lngUnitNumber As Long, _
                                        ByRef rsCLINNOTE As Adodb.Recordset) As DMErrorCodes

    Dim oItemNode As MSXML.IXMLDOMNode
    Dim oChildNode As MSXML.IXMLDOMNode
    
    On Error GoTo ErrHandler
    
    If oZCNNode Is Nothing Then
        Exit Function
    End If
    
    rsCLINNOTE.AddNew

    For Each oItemNode In oZCNNode.childNodes
        Select Case UCase(oItemNode.nodeName)
            Case "ZCN.1"
                
            Case "ZCN.2"
                rsCLINNOTE.Fields("Unit_Number").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZCN.3"
                rsCLINNOTE.Fields("Encounter").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZCN.4"
                rsCLINNOTE.Fields("DataID").Value = XMLToOracleNumber(XMLToOracleNumber(oItemNode.Text))
            Case "ZCN.5"
                rsCLINNOTE.Fields("ApplicationNCID").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZCN.6"
                For Each oChildNode In oItemNode.childNodes
                    Select Case oChildNode.nodeName
                        Case "CE.1"
                            rsCLINNOTE.Fields("NoteTypeNCID").Value = XMLToOracleNumber(oChildNode.Text)
                        Case "CE.2"
                            rsCLINNOTE.Fields("NoteTypeName").Value = NullToNothing(oChildNode.Text)
                    End Select
                Next
            Case "ZCN.7"
                rsCLINNOTE.Fields("ApplicationNCID").Value = XMLToOracleNumber(oItemNode.Text)
             Case "ZCN.8"
                'rsCLINNOTE.Fields("POCFacility").Value = XMLToOracleNumber(oItemNode.Text)
                rsCLINNOTE.Fields("POCFacility").Value = gobjLogon.FacilityNCID
            Case "ZCN.9"
                'rsCLINNOTE.Fields("EncFacility").Value = XMLToOracleNumber(oItemNode.Text)
                rsCLINNOTE.Fields("EncFacility").Value = gobjLogon.FacilityNCID
            Case "ZCN.10"
                Set oChildNode = oItemNode.selectSingleNode("//XCN.1")
                If Not oChildNode Is Nothing Then
                    'Returning SSN
                    rsCLINNOTE.Fields("ClinicianNCID").Value = XMLToOracleNumber(oChildNode.Text)
                End If
            Case "ZCN.11"
                rsCLINNOTE.Fields("HTMLText").Value = NullToNothing(oItemNode.Text)
            Case "ZCN.12"
                rsCLINNOTE.Fields("Caller").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZCN.13"
                rsCLINNOTE.Fields("Status").Value = NullToNothing(oItemNode.Text)
            Case "ZCN.14"
                rsCLINNOTE.Fields("CreatedBy").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZCN.15"
                rsCLINNOTE.Fields("CreateDate").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZCN.16"
                rsCLINNOTE.Fields("ModifiedBy").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZCN.17"
                rsCLINNOTE.Fields("ModifiedDate").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZCN.18"
                rsCLINNOTE.Fields("Priority").Value = NullToNothing(oItemNode.Text)
            Case "ZCN.19"
                rsCLINNOTE.Fields("ImageAvailable").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZCN.20"
                rsCLINNOTE.Fields("ImageNumber").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZCN.21"
                rsCLINNOTE.Fields("CDR_DataID").Value = XMLToOracleNumber(oItemNode.Text)
        End Select
    Next oItemNode
    
    rsCLINNOTE.Update
    
    FillRecordsetFromZCNNode = DMErrorCodes.ERROR_SUCCESS

Exit Function

ErrHandler:
    FillRecordsetFromZCNNode = ERROR_EIC_XML_TO_RECORDSET
    If gobjShared Is Nothing Then
        Set gobjShared = New CWShared
    End If
    gobjShared.ShowVBError Err.Number, Err.Description, "FillRecordsetFromZCNNode Function", "DMBE", vbExclamation
    Resume Next 'There is no reason this procedure should not continue to create xml if it errors
End Function

'*****************************************
'SCR 71415 - Import Allergy Data From EIC.
'*****************************************
Public Function FillRecordsetFromZALNode(ByVal oZALNode As MSXML.IXMLDOMElement, ByVal lngUnitNumber As Long, _
                                        ByRef rsALLERGIES As Adodb.Recordset, ByRef rsALLERGYREACTION As Adodb.Recordset, ByRef rsALLERGYCOMMENT As Adodb.Recordset) As Long

    Dim oItemNode As MSXML.IXMLDOMNode
    Dim oChildNode As MSXML.IXMLDOMNode
    
    Dim sProviderSSN As String
    Dim sProviderLName As String
    Dim sProviderFName As String
    Dim sProviderMName As String
    Dim sProviderNCID As String
    
    Dim oSQLOps As ICHCSII_SQl.ISqlOpsEx
    
    Dim sFacilityNCID As String
    
    On Error GoTo ErrHandler
    
    If oZALNode Is Nothing Then
        Exit Function
    End If
    
    If rsALLERGIES Is Nothing Then
        Set rsALLERGIES = New Adodb.Recordset
        Call GetFieldProps("ALLERGIES", rsALLERGIES)
    End If
    
    If rsALLERGYREACTION Is Nothing Then
        Set rsALLERGYREACTION = New Adodb.Recordset
         Call GetFieldProps("ALLERGYREACTION", rsALLERGYREACTION)
    End If
    
    If rsALLERGYCOMMENT Is Nothing Then
        Set rsALLERGYCOMMENT = New Adodb.Recordset
         Call GetFieldProps("ALLERGYCOMMENT", rsALLERGYCOMMENT)
    End If
    
    rsALLERGIES.AddNew
    rsALLERGYREACTION.AddNew
    rsALLERGYCOMMENT.AddNew
    
    rsALLERGIES.Fields("Unit_Number").Value = lngUnitNumber
    rsALLERGYCOMMENT.Fields("Unit_Number").Value = lngUnitNumber
    rsALLERGYREACTION.Fields("Unit_Number").Value = lngUnitNumber
    
    For Each oItemNode In oZALNode.childNodes
        Select Case UCase(oItemNode.nodeName)
            Case "ZAL.1"
                
            Case "ZAL.2"
            
                rsALLERGIES.Fields("DataID").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZAL.3"
                rsALLERGIES.Fields("Unit_Number").Value = lngUnitNumber 'XMLToOracleNumber(oItemNode.Text)
            Case "ZAL.4"
            
              sFacilityNCID = vbNullString
              
              '--- SCR 89510   Sherry Wang   3/8/2006
              If LenB(oItemNode.Text) > 0 Then
                sFacilityNCID = XMLToOracleNumber(XMLToOracleNumber(oItemNode.Text))
              End If
              
              If sFacilityNCID <> vbNullString Then
                rsALLERGIES.Fields("Facility_NCID").Value = sFacilityNCID
              Else
                rsALLERGIES.Fields("Facility_NCID").Value = gobjLogon.FacilityNCID
              End If
            
            Case "ZAL.5"
            
              sProviderSSN = vbNullString
              sProviderLName = vbNullString
              sProviderFName = vbNullString
              sProviderMName = vbNullString
              sProviderNCID = vbNullString

                 For Each oChildNode In oItemNode.childNodes
                    Select Case UCase(oChildNode.nodeName)
                    Case "XCN.1"
                      sProviderSSN = NullToNothing(oChildNode.Text)
                    Case "XCN.2"
                      sProviderLName = NullToNothing(oChildNode.Text)
                    Case "XCN.3"
                      sProviderFName = NullToNothing(oChildNode.Text)
                    Case "XCN.4"
                      sProviderMName = NullToNothing(oChildNode.Text)
                    End Select
                Next
                
                If Trim$(sProviderSSN) <> "" Then
                  Set oSQLOps = oDatabase.CHCSII_SQLOPS_EX(Auto)
                  Call oSQLOps.ValidateProvider2(sProviderNCID, sProviderSSN, sProviderFName, sProviderMName, sProviderLName, True)
                Else
                  sProviderNCID = gobjLogon.UserNCID
                End If
                
                rsALLERGIES.Fields("ClinicianNCID").Value = sProviderNCID
        
'''                Set oChildNode = oItemNode.selectSingleNode("CE.1")
''''                rsALLERGIES.Fields("ClinicianNCID").Value = gobjLogon.UserNCID
'''                If Not oChildNode Is Nothing Then
'''                  If Len(oItemNode.Text) = 0 Then
'''                    rsALLERGIES.Fields("ClinicianNCID").Value = gobjLogon.UserNCID
'''                  Else
'''                    rsALLERGIES.Fields("ClinicianNCID").Value = XMLToOracleNumber(oChildNode.Text)
'''                  End If
'''                End If
            Case "ZAL.6"
                rsALLERGIES.Fields("CreateTime").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZAL.7"
                '<SCR 22974 CC>
                'rsALLERGIES.Fields("InfoSourceNCID").Value = XMLToOracleNumber(oItemNode.Text)
                Set oChildNode = oItemNode.selectSingleNode("CE.5")
                If Not oItemNode Is Nothing Then
                    rsALLERGIES.Fields("InfoSourceNCID").Value = XMLToOracleNumber(oChildNode.Text)
                End If
            Case "ZAL.8"
                rsALLERGIES.Fields("Modified").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZAL.9"
                rsALLERGIES.Fields("ModifyTime").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZAL.10"
                'SCR 164074
'                If Len(oItemNode.Text) = 0 Then
                  rsALLERGIES.Fields("ModifyUserNCID").Value = gobjLogon.UserNCID
'                Else
'                  rsALLERGIES.Fields("ModifyUserNCID").Value = XMLToOracleNumber(oItemNode.Text)
'                End If
            Case "ZAL.11"
                '<SCR 22974 CC>
                'rsALLERGIES.Fields("PointOfCareFacilityNCID").Value = XMLToOracleNumber(oItemNode.Text)
                'Set oChildNode = oItemNode.selectSingleNode("CE.5")
                'If Not oItemNode Is Nothing Then
                '    rsALLERGIES.Fields("PointOfCareFacilityNCID").Value = XMLToOracleNumber(oChildNode.Text)
                'End If
                'rsALLERGIES.Fields("PointOfCareFacilityNCID").Value = XMLToOracleNumber(oChildNode.Text)
                rsALLERGIES.Fields("PointOfCareFacilityNCID").Value = gobjLogon.FacilityNCID
            Case "ZAL.12"
                 'SCR 163958
'                If Len(oItemNode.Text) = 0 Then
                  rsALLERGIES.Fields("CreateUserNCID").Value = gobjLogon.UserNCID
'                Else
'                  rsALLERGIES.Fields("CreateUserNCID").Value = XMLToOracleNumber(oItemNode.Text)
'                End If
            Case "ZAL.13"
                rsALLERGIES.Fields("UpdateFlag").Value = NullToNothing(oItemNode.Text)
            Case "ZAL.14"
                rsALLERGIES.Fields("CDR_DataID").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZAL.15"
                rsALLERGIES.Fields("CreatedBy").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZAL.16"
                rsALLERGIES.Fields("CreatedOn").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZAL.17"
                rsALLERGIES.Fields("UpdatedBy").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZAL.18"
                rsALLERGIES.Fields("UpdatedOn").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZAL.19"
                rsALLERGYCOMMENT.Fields("Unit_Number").Value = lngUnitNumber 'XMLToOracleNumber(oItemNode.Text)
            Case "ZAL.20"
                rsALLERGYCOMMENT.Fields("AllergenNCID").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZAL.21"
                rsALLERGYCOMMENT.Fields("ALLERGYCOMMENT").Value = NullToNothing(oItemNode.Text)
            Case "ZAL.22"
                rsALLERGYCOMMENT.Fields("UpdateFlag").Value = NullToNothing(oItemNode.Text)
            Case "ZAL.23"
                rsALLERGYCOMMENT.Fields("CreatedBy").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZAL.24"
                rsALLERGYCOMMENT.Fields("CreatedOn").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZAL.25"
                rsALLERGYCOMMENT.Fields("UpdatedBy").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZAL.26"
                rsALLERGYCOMMENT.Fields("UpdatedOn").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZAL.27"
                rsALLERGYREACTION.Fields("Unit_Number").Value = lngUnitNumber 'XMLToOracleNumber(oItemNode.Text)
            Case "ZAL.28"
                rsALLERGYREACTION.Fields("AllergenNCID").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZAL.29"
                rsALLERGYREACTION.Fields("UpdateFlag").Value = NullToNothing(oItemNode.Text)
            Case "ZAL.30"
                rsALLERGYREACTION.Fields("CreatedBy").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZAL.31"
                rsALLERGYREACTION.Fields("CreatedOn").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZAL.32"
                rsALLERGYREACTION.Fields("UpdatedBy").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZAL.33"
                rsALLERGYREACTION.Fields("UpdatedOn").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZAL.34"
                Set oChildNode = oItemNode.selectSingleNode("CE.4")
                If Not oChildNode Is Nothing Then
                    rsALLERGYREACTION.Fields("REACTIONNCID").Value = XMLToOracleNumber(oChildNode.Text)
                End If
                Set oChildNode = oItemNode.selectSingleNode("CE.5")
                If Not oChildNode Is Nothing Then
                    rsALLERGYREACTION.Fields("REACTION").Value = NullToNothing(oChildNode.Text)
                End If
            Case "ZAL.35"
                 Set oChildNode = oItemNode.selectSingleNode("CE.2")
                If Not oChildNode Is Nothing Then
                    rsALLERGIES.Fields("ALLERGYTYPENCID").Value = XMLToOracleNumber(oChildNode.Text)
                End If
            Case "ZAL.36"
                Set oChildNode = oItemNode.selectSingleNode("CE.4")
                If Not oChildNode Is Nothing Then
                    rsALLERGIES.Fields("ALLERGENNCID").Value = XMLToOracleNumber(oChildNode.Text)
                    rsALLERGIES.Fields("FDB_SEQ_ID").Value = XMLToOracleNumber(oChildNode.Text)
                End If
                Set oChildNode = oItemNode.selectSingleNode("CE.5")
                If Not oChildNode Is Nothing Then
                    rsALLERGIES.Fields("ALLERGEN").Value = NullToNothing(oChildNode.Text)
                End If
            Case "ZAL.37"
                rsALLERGIES.Fields("NOTEDDATE").Value = XMLToOracleDate(oItemNode.Text)
        End Select
    Next oItemNode
    
    rsALLERGIES.Update
    rsALLERGYREACTION.Update
    rsALLERGYCOMMENT.Update
    
    FillRecordsetFromZALNode = DMErrorCodes.ERROR_SUCCESS

Exit Function

ErrHandler:

    FillRecordsetFromZALNode = ERROR_EIC_XML_TO_RECORDSET
    If gobjShared Is Nothing Then
        Set gobjShared = New CWShared
    End If
    gobjShared.ShowVBError Err.Number, Err.Description, "FillRecordsetFromZALNode Function", "DMBE", vbExclamation
    Resume Next 'There is no reason this procedure should not continue to create xml if it errors
End Function

Private Function FillRecordsetFromZVTNode(ByVal oZVTNode As MSXML.IXMLDOMElement, ByVal lngUnitNumber As Long, _
                                        ByRef rsVITALS As Adodb.Recordset, ByRef rsVITALSIGNSMODIFIER As Adodb.Recordset, ByRef rsVITALSIGNSVALUE As Adodb.Recordset) As DMErrorCodes

    Dim oItemNode As MSXML.IXMLDOMNode
    Dim oChildNode As MSXML.IXMLDOMNode
    Dim oChildNodeLevel1 As MSXML.IXMLDOMNode
    Dim oSQLOps As ICHCSII_SQl.ISqlOpsEx
    Dim sProviderSSN As String
    Dim sProviderLName As String
    Dim sProviderFName As String
    Dim sProviderMName As String
    Dim sProviderNCID As String
    
    On Error GoTo ErrHandler
    
    If oZVTNode Is Nothing Then
        Exit Function
    End If
    
    rsVITALS.AddNew
    ' SCR #47595;   Developer:  01/08/2004
    'rsVITALSIGNSMODIFIER.AddNew
    
    For Each oItemNode In oZVTNode.childNodes
        Select Case UCase(oItemNode.nodeName)
            Case "ZVT.1"
                
            Case "ZVT.2"
                'rsVITALS.Fields("FacilityNCID").Value = XMLToOracleNumber(oItemNode.Text)
                 rsVITALS.Fields("DataID").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZVT.3"
                'rsVITALS.Fields("FacilityNCID").Value = XMLToOracleNumber(oItemNode.Text)
                ' Begin:SCR #33071;   Developer: Mag 05/13/2003 02:24 PM
                rsVITALS.Fields("FacilityNCID").Value = XMLToOracleNumber(gobjLogon.FacilityNCID)
                ' End: SCR #33071;
            Case "ZVT.4"
                rsVITALS.Fields("Unit_Number").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZVT.5"
                rsVITALS.Fields("EncounterNumber").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZVT.6"
                'Tracker #51305
                For Each oChildNode In oItemNode.childNodes
                    Select Case UCase(oChildNode.nodeName)
                    Case "XCN.1"
                      'SSN instead of unit number
                      sProviderSSN = NullToNothing(oChildNode.Text)
                    Case "XCN.2"
                      sProviderLName = NullToNothing(oChildNode.Text)
                    Case "XCN.3"
                      sProviderFName = NullToNothing(oChildNode.Text)
                    Case "XCN.4"
                      sProviderMName = NullToNothing(oChildNode.Text)
                    End Select
                Next
                
                If Trim$(sProviderSSN) <> "" Then
                  Set oSQLOps = oDatabase.CHCSII_SQLOPS_EX(Auto)
                  Call oSQLOps.ValidateProvider2(sProviderNCID, sProviderSSN, sProviderFName, sProviderMName, sProviderLName, True)
                Else
                  sProviderNCID = gobjLogon.UserNCID
                End If
                rsVITALS.Fields("ClinicianNCID").Value = sProviderNCID
                
            Case "ZVT.7"
                rsVITALS.Fields("PanelType").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZVT.8"
                rsVITALS.Fields("EntryDate").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZVT.9"
                rsVITALS.Fields("VitalsComment").Value = NullToNothing(oItemNode.Text)
            Case "ZVT.10"
                rsVITALS.Fields("UpdateFlag").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZVT.11"
                rsVITALS.Fields("CDR_DataID").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZVT.12"
                rsVITALS.Fields("CREATEDBY").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZVT.13"
                rsVITALS.Fields("CREATEDON").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZVT.14"
                rsVITALS.Fields("UPDATEDBY").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZVT.15"
                rsVITALS.Fields("UPDATEDON").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZVT.16.LST"
                For Each oChildNode In oItemNode.childNodes 'SCR 23969
                    If oChildNode.nodeName = "ZVT.16" Then
                        rsVITALSIGNSVALUE.AddNew
                        For Each oChildNodeLevel1 In oChildNode.childNodes 'SCR 23969
                            Select Case UCase(oChildNodeLevel1.nodeName)
                                Case "ZVT.16.1"
                                    rsVITALSIGNSVALUE.Fields("DataID").Value = XMLToOracleNumber(oChildNodeLevel1.Text)
                                Case "ZVT.16.2"
                                    rsVITALSIGNSVALUE.Fields("VitalsType").Value = XMLToOracleNumber(oChildNodeLevel1.Text)
                                Case "ZVT.16.3"
                                    rsVITALSIGNSVALUE.Fields("VitalsValue").Value = NullToNothing(oChildNodeLevel1.Text)
                                Case "ZVT.16.4"
                                    rsVITALSIGNSVALUE.Fields("Units").Value = XMLToOracleNumber(oChildNodeLevel1.Text)
                                Case "ZVT.16.5"
                                    rsVITALSIGNSVALUE.Fields("CREATEDBY").Value = XMLToOracleNumber(oChildNodeLevel1.Text)
                                Case "ZVT.16.6"
                                    rsVITALSIGNSVALUE.Fields("CREATEDON").Value = XMLToOracleDate(oChildNodeLevel1.Text)
                                Case "ZVT.16.7"
                                    rsVITALSIGNSVALUE.Fields("UPDATEDBY").Value = XMLToOracleNumber(oChildNodeLevel1.Text)
                                Case "ZVT.16.8"
                                    rsVITALSIGNSVALUE.Fields("UPDATEDON").Value = XMLToOracleDate(oChildNodeLevel1.Text)
                            End Select
                        Next
                        rsVITALSIGNSVALUE.Update
                    End If
                Next
            
            Case "ZVT.17.LST"
                For Each oChildNode In oItemNode.childNodes 'SCR 23969
                    If oChildNode.nodeName = "ZVT.17" Then
                        rsVITALSIGNSMODIFIER.AddNew
                        For Each oChildNodeLevel1 In oChildNode.childNodes
                            Select Case UCase(oChildNodeLevel1.nodeName)
                                Case "ZVT.17.1"
                                    rsVITALSIGNSMODIFIER.Fields("DataID").Value = XMLToOracleNumber(oChildNodeLevel1.Text)
                                Case "ZVT.17.2"
                                    rsVITALSIGNSMODIFIER.Fields("VitalsType").Value = XMLToOracleNumber(oChildNodeLevel1.Text)
                                Case "ZVT.17.3"
                                    rsVITALSIGNSMODIFIER.Fields("OBSERVATIONMODIFIER").Value = XMLToOracleNumber(oChildNodeLevel1.Text)
                                Case "ZVT.17.4"
                                    rsVITALSIGNSMODIFIER.Fields("Type").Value = XMLToOracleNumber(oChildNodeLevel1.Text)
                                Case "ZVT.17.5"
                                    rsVITALSIGNSMODIFIER.Fields("Domain").Value = XMLToOracleNumber(oChildNodeLevel1.Text)
                                Case "ZVT.17.6"
                                    rsVITALSIGNSMODIFIER.Fields("ModifierValue").Value = NullToNothing(oChildNodeLevel1.Text)
                                Case "ZVT.17.7"
                                    rsVITALSIGNSMODIFIER.Fields("BodySide").Value = XMLToOracleNumber(oChildNodeLevel1.Text)
                                Case "ZVT.17.8"
                                    rsVITALSIGNSMODIFIER.Fields("AggObs").Value = NullToNothing(oChildNodeLevel1.Text)
                                Case "ZVT.17.9"
                                    rsVITALSIGNSMODIFIER.Fields("CREATEDBY").Value = XMLToOracleNumber(oChildNodeLevel1.Text)
                                Case "ZVT.17.10"
                                    rsVITALSIGNSMODIFIER.Fields("CREATEDON").Value = XMLToOracleDate(oChildNodeLevel1.Text)
                                Case "ZVT.17.11"
                                    rsVITALSIGNSMODIFIER.Fields("UPDATEDBY").Value = XMLToOracleNumber(oChildNodeLevel1.Text)
                                Case "ZVT.17.12"
                                    rsVITALSIGNSMODIFIER.Fields("UPDATEDON").Value = XMLToOracleDate(oChildNodeLevel1.Text)
                            End Select
                        Next oChildNodeLevel1
                        rsVITALSIGNSMODIFIER.Update
                    End If
                Next oChildNode
        End Select
    Next oItemNode
    
    rsVITALS.Update
    'rsVITALSIGNSMODIFIER.Update
    'rsVITALSIGNSVALUE.Update
    
    FillRecordsetFromZVTNode = DMErrorCodes.ERROR_SUCCESS

Exit Function

ErrHandler:
    FillRecordsetFromZVTNode = DMErrorCodes.ERROR_EIC_XML_TO_RECORDSET
    If gobjShared Is Nothing Then
        Set gobjShared = New CWShared
    End If
    gobjShared.ShowVBError Err.Number, Err.Description, "FillRecordsetFromZVTNode Function", "DMBE", vbExclamation
    Resume Next 'There is no reason this procedure should not continue to create xml if it errors
    Exit Function
    Resume
End Function



'<SCR 22800 CC>
Private Function FillRecordsetFromZRDNode(ByVal oZRDNode As MSXML.IXMLDOMElement, ByVal lngUnitNumber As Long, _
                                        ByRef rsREADINESS As Adodb.Recordset) As DMErrorCodes

    Dim oItemNode As MSXML.IXMLDOMNode
    Dim oChildNode As MSXML.IXMLDOMNode
    
    On Error GoTo ErrHandler
    
    If oZRDNode Is Nothing Then
        Exit Function
    End If
    
    rsREADINESS.AddNew
    
    For Each oItemNode In oZRDNode.childNodes
        Select Case UCase(oItemNode.nodeName)
            Case "ZRD.1"
                
            Case "ZRD.2"
                rsREADINESS.Fields("Last_Readiness_DT").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZRD.3"
                rsREADINESS.Fields("HIV_Test_DT").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZRD.4"
                rsREADINESS.Fields("HIV_Test_Rslt").Value = NullToNothing(XMLToOracleNumber(oItemNode.Text))
            Case "ZRD.5"
                rsREADINESS.Fields("G6PD_Test_DT").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZRD.6"
                rsREADINESS.Fields("G6PD_Test_Rslt").Value = NullToNothing(oItemNode.Text)
            Case "ZRD.7"
                rsREADINESS.Fields("Sickle_Cell_Test_DT").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZRD.8"
                rsREADINESS.Fields("Sickle_Cell_Test_Rslt").Value = NullToNothing(oItemNode.Text)
            Case "ZRD.9"
                rsREADINESS.Fields("Blood_Type_Test_DT").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZRD.10"
                rsREADINESS.Fields("Blood_Type_Test_Rslt").Value = NullToNothing(oItemNode.Text)
            Case "ZRD.11"
                rsREADINESS.Fields("DNA_Test_DT").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZRD.12"
                rsREADINESS.Fields("Hear_Exam_DT").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZRD.13"
                rsREADINESS.Fields("Hear_Aid_Required").Value = NullToNothing(oItemNode.Text)
            Case "ZRD.14"
                rsREADINESS.Fields("Hear_Aid_Issued").Value = NullToNothing(oItemNode.Text)
            Case "ZRD.15"
                rsREADINESS.Fields("Hear_Aid_Issued_DT").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZRD.16"
                rsREADINESS.Fields("Last_Dental_Exam_DT").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZRD.17"
                rsREADINESS.Fields("Dental_Readiness_CD").Value = NullToNothing(oItemNode.Text)
            Case "ZRD.18"
                rsREADINESS.Fields("Dental_Readiness_DT").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZRD.19"
                rsREADINESS.Fields("Medic_Assessment_DT").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZRD.20"
                rsREADINESS.Fields("Temp_Profile").Value = NullToNothing(oItemNode.Text)
            Case "ZRD.21"
                rsREADINESS.Fields("Temp_Profile_Start_DT").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZRD.22"
                rsREADINESS.Fields("Temp_Profile_End_DT").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZRD.23"
                For Each oChildNode In oItemNode.childNodes
                    Select Case UCase(oChildNode.nodeName)
                        Case "CE.1"
                            rsREADINESS.Fields("Temp_Profile_Dignos_CD").Value = NullToNothing(oChildNode.Text)
                        Case "CE.2"
                            rsREADINESS.Fields("Temp_Profile_Dignos_Dsc").Value = NullToNothing(oChildNode.Text)
                    End Select
                Next
            Case "ZRD.24"
                rsREADINESS.Fields("Temp_Profile_Limitation").Value = NullToNothing(oItemNode.Text)
            Case "ZRD.25"
                rsREADINESS.Fields("Last_Visual_Acuity_DT").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZRD.26"
                rsREADINESS.Fields("Uncorrected_OD").Value = NullToNothing(oItemNode.Text)
            Case "ZRD.27"
                rsREADINESS.Fields("Corrected_OD").Value = NullToNothing(oItemNode.Text)
            Case "ZRD.28"
                rsREADINESS.Fields("Uncorrected_OS").Value = NullToNothing(oItemNode.Text)
            Case "ZRD.29"
                rsREADINESS.Fields("Corrected_OS").Value = NullToNothing(oItemNode.Text)
            Case "ZRD.30"
                rsREADINESS.Fields("Uncorrected_OU").Value = NullToNothing(oItemNode.Text)
            Case "ZRD.31"
                rsREADINESS.Fields("Corrected_OU").Value = NullToNothing(oItemNode.Text)
            Case "ZRD.32"
                rsREADINESS.Fields("Protect_Insert_Required").Value = NullToNothing(oItemNode.Text)
            Case "ZRD.33"
                rsREADINESS.Fields("Protect_Insert_Issued").Value = NullToNothing(oItemNode.Text)
            Case "ZRD.34"
                rsREADINESS.Fields("Protect_Insert_Exam_DT").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZRD.35"
                rsREADINESS.Fields("DNA_On_File").Value = NullToNothing(oItemNode.Text)
            Case "ZRD.36"
                rsREADINESS.Fields("PATIENT_UNIT_NUMBER").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZRD.37"
                rsREADINESS.Fields("ENCOUNTER_NUMBER").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZRD.38"
                rsREADINESS.Fields("CREATEDBY").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZRD.39"
                rsREADINESS.Fields("CREATEDON").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZRD.40"
                rsREADINESS.Fields("UPDATEDBY").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZRD.41"
                rsREADINESS.Fields("UPDATEDON").Value = XMLToOracleDate(oItemNode.Text)
        End Select
    Next oItemNode
    
    rsREADINESS.Update
    
    FillRecordsetFromZRDNode = DMErrorCodes.ERROR_SUCCESS

Exit Function

ErrHandler:
    FillRecordsetFromZRDNode = DMErrorCodes.ERROR_EIC_XML_TO_RECORDSET
    If gobjShared Is Nothing Then
        Set gobjShared = New CWShared
    End If
    gobjShared.ShowVBError Err.Number, Err.Description, "FillRecordsetFromZRDNode Function", "DMBE", vbExclamation
    Resume Next 'There is no reason this procedure should not continue to create xml if it errors
End Function
'</SCR 22800 CC>

Private Function FillRecordsetFromZBRNode(ByVal oZBRNode As MSXML.IXMLDOMElement, ByVal lngUnitNumber As Long, _
                                        ByRef rsPROBLEMSRELATIONSHIPS As Adodb.Recordset) As DMErrorCodes

    Dim oItemNode As MSXML.IXMLDOMNode
    
    On Error GoTo ErrHandler
    
    If oZBRNode Is Nothing Then
        Exit Function
    End If
    
    rsPROBLEMSRELATIONSHIPS.AddNew

    For Each oItemNode In oZBRNode.childNodes
        Select Case UCase(oItemNode.nodeName)
            Case "ZBR.1"
                
            Case "ZBR.2"
                rsPROBLEMSRELATIONSHIPS.Fields("ProblemDataID").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZBR.3"
                rsPROBLEMSRELATIONSHIPS.Fields("RelatedDataID").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZBR.4"
                rsPROBLEMSRELATIONSHIPS.Fields("RelationshipNCID").Value = XMLToOracleNumber(XMLToOracleNumber(oItemNode.Text))
            Case "ZBR.5"
                rsPROBLEMSRELATIONSHIPS.Fields("CreatedBy").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZBR.6"
                rsPROBLEMSRELATIONSHIPS.Fields("CreatedOn").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZBR.7"
                rsPROBLEMSRELATIONSHIPS.Fields("UpdatedBy").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZBR.8"
                rsPROBLEMSRELATIONSHIPS.Fields("UpdatedOn").Value = XMLToOracleDate(oItemNode.Text)
        End Select
    Next oItemNode
    
    rsPROBLEMSRELATIONSHIPS.Update
    
    FillRecordsetFromZBRNode = DMErrorCodes.ERROR_SUCCESS

Exit Function

ErrHandler:
    FillRecordsetFromZBRNode = ERROR_EIC_XML_TO_RECORDSET
    If gobjShared Is Nothing Then
        Set gobjShared = New CWShared
    End If
    gobjShared.ShowVBError Err.Number, Err.Description, "FillRecordsetFromZBRNode Function", "DMBE", vbExclamation
    Resume Next 'There is no reason this procedure should not continue to create xml if it errors
End Function

Private Function FillRecordsetFromZPBNode(ByVal oZPBNode As MSXML.IXMLDOMElement, ByVal lngUnitNumber As Long, _
                                        ByRef rsPROBLEMS As Adodb.Recordset) As DMErrorCodes

    Dim oItemNode As MSXML.IXMLDOMNode
    Dim oChildNode As MSXML.IXMLDOMNode
    
    On Error GoTo ErrHandler
    
    If oZPBNode Is Nothing Then
        Exit Function
    End If
    
    rsPROBLEMS.AddNew

    For Each oItemNode In oZPBNode.childNodes
        Select Case UCase(oItemNode.nodeName)
            Case "ZPB.1"
                
            Case "ZPB.2"
                rsPROBLEMS.Fields("DataID").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZPB.3"
                'rsPROBLEMS.Fields("FacilityNCID").Value = XMLToOracleNumber(oItemNode.Text)
                rsPROBLEMS.Fields("FacilityNCID").Value = gobjLogon.FacilityNCID
            Case "ZPB.4"
                rsPROBLEMS.Fields("Unit_Number").Value = XMLToOracleNumber(XMLToOracleNumber(oItemNode.Text))
            Case "ZPB.5"
                rsPROBLEMS.Fields("EncounterNumber").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZPB.6"
                Set oChildNode = oItemNode.selectSingleNode("CE.5")
                If Not oChildNode Is Nothing Then
                    rsPROBLEMS.Fields("ProblemNCID").Value = XMLToOracleNumber(oChildNode.Text)
                End If
            Case "ZPB.7"
                rsPROBLEMS.Fields("ProblemsComment").Value = NullToNothing(oItemNode.Text)
            Case "ZPB.8"
                '<SCR 22973 CC>
                'rsPROBLEMS.Fields("ChronicityNCID").Value = XMLToOracleNumber(oItemNode.Text)
                Set oChildNode = oItemNode.selectSingleNode("CE.5")
                If Not oChildNode Is Nothing Then
                    rsPROBLEMS.Fields("ChronicityNCID").Value = XMLToOracleNumber(oChildNode.Text)
                End If
            Case "ZPB.9"
                'rsPROBLEMS.Fields("StatusNCID").Value = XMLToOracleNumber(oItemNode.Text)
                Set oChildNode = oItemNode.selectSingleNode("CE.5")
                If Not oChildNode Is Nothing Then
                    rsPROBLEMS.Fields("StatusNCID").Value = XMLToOracleNumber(oChildNode.Text)
                End If
            Case "ZPB.10"
                rsPROBLEMS.Fields("OnSetDate").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZPB.11"
                'rsPROBLEMS.Fields("SourceNCID").Value = XMLToOracleNumber(oItemNode.Text)
                Set oChildNode = oItemNode.selectSingleNode("CE.5")
                If Not oChildNode Is Nothing Then
                    rsPROBLEMS.Fields("SourceNCID").Value = XMLToOracleNumber(oChildNode.Text)
                End If
            Case "ZPB.12"
                Set oChildNode = oItemNode.selectSingleNode("//XCN.1")
                If Not oChildNode Is Nothing Then
                    'Returning SSN
                    rsPROBLEMS.Fields("ClinicianNCID").Value = XMLToOracleNumber(oChildNode.Text)
                End If
            Case "ZPB.13"
                rsPROBLEMS.Fields("SnoID").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZPB.14"
                'rsPROBLEMS.Fields("PointOfCareFacilityNCID").Value = XMLToOracleNumber(oItemNode.Text)
                'Set oChildNode = oItemNode.selectSingleNode("CE.5")
                'If Not oChildNode Is Nothing Then
                '    rsPROBLEMS.Fields("PointOfCareFacilityNCID").Value = XMLToOracleNumber(oChildNode.Text)
                'End If
                rsPROBLEMS.Fields("PointOfCareFacilityNCID").Value = gobjLogon.FacilityNCID
                '</SCR 22973 CC>
            Case "ZPB.15"
                rsPROBLEMS.Fields("NotedDate").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZPB.16"
                rsPROBLEMS.Fields("CreateUserNCID").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZPB.17"
                rsPROBLEMS.Fields("CreateTime").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZPB.18"
                rsPROBLEMS.Fields("ModifyUserNCID").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZPB.19"
                rsPROBLEMS.Fields("ModifyTime").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZPB.20"
                rsPROBLEMS.Fields("UpdateFlag").Value = NullToNothing(oItemNode.Text)
            Case "ZPB.21"
                rsPROBLEMS.Fields("CDR_DataID").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZPB.22"
                rsPROBLEMS.Fields("CreatedBy").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZPB.23"
                rsPROBLEMS.Fields("CreatedOn").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZPB.24"
                rsPROBLEMS.Fields("UpdatedBy").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZPB.25"
                rsPROBLEMS.Fields("UpdatedOn").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZPB.26"
                rsPROBLEMS.Fields("Prefix").Value = NullToNothing(oItemNode.Text)
        End Select
    Next oItemNode
    
    rsPROBLEMS.Update
    
    FillRecordsetFromZPBNode = DMErrorCodes.ERROR_SUCCESS

Exit Function

ErrHandler:
    FillRecordsetFromZPBNode = DMErrorCodes.ERROR_EIC_XML_TO_RECORDSET
    If gobjShared Is Nothing Then
        Set gobjShared = New CWShared
    End If
    gobjShared.ShowVBError Err.Number, Err.Description, "FillRecordsetFromZPBNode Function", "DMBE", vbExclamation
    Resume Next 'There is no reason this procedure should not continue to create xml if it errors
End Function

Private Function FillRecordsetFromZCPNode(ByVal oZCPNode As MSXML.IXMLDOMElement, ByVal lngUnitNumber As Long, _
                                        ByRef rsCOMPLAINTS As Adodb.Recordset) As DMErrorCodes

    Dim oItemNode As MSXML.IXMLDOMNode
    Dim oChildNode As MSXML.IXMLDOMNode
    
    On Error GoTo ErrHandler
    
    If oZCPNode Is Nothing Then
        Exit Function
    End If
    
    rsCOMPLAINTS.AddNew

    For Each oItemNode In oZCPNode.childNodes
        Select Case UCase(oItemNode.nodeName)
            Case "ZCP.1"
                
            Case "ZCP.2"
                rsCOMPLAINTS.Fields("DataID").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZCP.3"
                'rsCOMPLAINTS.Fields("FacilityNCID").Value = XMLToOracleNumber(oItemNode.Text)
                rsCOMPLAINTS.Fields("FacilityNCID").Value = gobjLogon.FacilityNCID
            Case "ZCP.4"
                rsCOMPLAINTS.Fields("Unit_Number").Value = XMLToOracleNumber(XMLToOracleNumber(oItemNode.Text))
            Case "ZCP.5"
                rsCOMPLAINTS.Fields("EncounterNumber").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZCP.6"
                rsCOMPLAINTS.Fields("SnoID").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZCP.7"
                Set oChildNode = oItemNode.selectSingleNode("XCN.1")
                If Not oChildNode Is Nothing Then
                    'Returning SSN
                    rsCOMPLAINTS.Fields("ClinicianNCID").Value = XMLToOracleNumber(oChildNode.Text)
                End If
            Case "ZCP.8"
                'rsCOMPLAINTS.Fields("PointOfCareFacilityNCID").Value = XMLToOracleNumber(oItemNode.Text)
                Set oChildNode = oItemNode.selectSingleNode("CE.5")
                'If Not oChildNode Is Nothing Then
                '    rsCOMPLAINTS.Fields("PointOfCareFacilityNCID").Value = XMLToOracleNumber(oChildNode.Text)
                'End If
                rsCOMPLAINTS.Fields("PointOfCareFacilityNCID").Value = gobjLogon.FacilityNCID
            Case "ZCP.9"
                'rsCOMPLAINTS.Fields("StatusNCID").Value = XMLToOracleNumber(oItemNode.Text)
                Set oChildNode = oItemNode.selectSingleNode("CE.5")
                If Not oChildNode Is Nothing Then
                    rsCOMPLAINTS.Fields("StatusNCID").Value = XMLToOracleNumber(oChildNode.Text)
                End If
            Case "ZCP.10"
                rsCOMPLAINTS.Fields("NotedDate").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZCP.11"
                rsCOMPLAINTS.Fields("OnsetDate").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZCP.12"
                'rsCOMPLAINTS.Fields("ComplaintNCID").Value = XMLToOracleNumber(oItemNode.Text)
                Set oChildNode = oItemNode.selectSingleNode("CE.5")
                If Not oChildNode Is Nothing Then
                    rsCOMPLAINTS.Fields("ComplaintNCID").Value = XMLToOracleNumber(oChildNode.Text)
                End If
            Case "ZCP.13"
                rsCOMPLAINTS.Fields("NewFollowUp").Value = NullToNothing(oItemNode.Text)
            Case "ZCP.14"
                rsCOMPLAINTS.Fields("CreateUserNCID").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZCP.15"
                rsCOMPLAINTS.Fields("CreateTime").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZCP.16"
                rsCOMPLAINTS.Fields("ModifyUserNCID").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZCP.17"
                rsCOMPLAINTS.Fields("ModifyTime").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZCP.18"
                rsCOMPLAINTS.Fields("UpdateFlag").Value = NullToNothing(oItemNode.Text)
            Case "ZCP.19"
                rsCOMPLAINTS.Fields("ComplaintsComment").Value = NullToNothing(oItemNode.Text)
            Case "ZCP.20"
                rsCOMPLAINTS.Fields("CDR_DataID").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZCP.21"
                rsCOMPLAINTS.Fields("CreatedBy").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZCP.22"
                rsCOMPLAINTS.Fields("CreatedOn").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZCP.23"
                rsCOMPLAINTS.Fields("UpdatedBy").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZCP.24"
                rsCOMPLAINTS.Fields("UpdatedOn").Value = XMLToOracleDate(oItemNode.Text)
        End Select
    Next oItemNode
    
    rsCOMPLAINTS.Update
    
    FillRecordsetFromZCPNode = DMErrorCodes.ERROR_SUCCESS

Exit Function

ErrHandler:
    FillRecordsetFromZCPNode = ERROR_EIC_XML_TO_RECORDSET
    If gobjShared Is Nothing Then
        Set gobjShared = New CWShared
    End If
    gobjShared.ShowVBError Err.Number, Err.Description, "FillRecordsetFromZCPNode Function", "DMBE", vbExclamation
    Resume Next 'There is no reason this procedure should not continue to create xml if it errors
End Function

Private Function FillRecordsetFromZLRNode(ByVal oZLRNode As MSXML.IXMLDOMElement, ByVal lngUnitNumber As Long, _
                                        ByRef rsLISTTOOLRECORDS As Adodb.Recordset) As DMErrorCodes

    Dim oItemNode As MSXML.IXMLDOMNode
    
    On Error GoTo ErrHandler
    
    If oZLRNode Is Nothing Then
        Exit Function
    End If
    
    rsLISTTOOLRECORDS.AddNew

    For Each oItemNode In oZLRNode.childNodes
        Select Case UCase(oItemNode.nodeName)
            Case "ZLR.1"
                
            Case "ZLR.2"
                rsLISTTOOLRECORDS.Fields("DataID").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZLR.3"
                rsLISTTOOLRECORDS.Fields("SnoID").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZLR.4"
                rsLISTTOOLRECORDS.Fields("Prefix").Value = NullToNothing(oItemNode.Text) 'XMLToOracleNumber
                'SCR 114875 2/12/07 Quan
                If rsLISTTOOLRECORDS.Fields("Prefix").Value = "" Then
                    rsLISTTOOLRECORDS.Fields("Prefix").Value = "0"
                End If
            Case "ZLR.5"
                rsLISTTOOLRECORDS.Fields("Modifier").Value = NullToNothing(oItemNode.Text)
            Case "ZLR.6"
                rsLISTTOOLRECORDS.Fields("Result").Value = NullToNothing(oItemNode.Text)
            Case "ZLR.7"
                rsLISTTOOLRECORDS.Fields("Status").Value = NullToNothing(oItemNode.Text)
            Case "ZLR.8"
                rsLISTTOOLRECORDS.Fields("Onset").Value = NullToNothing(oItemNode.Text)
            Case "ZLR.9"
                rsLISTTOOLRECORDS.Fields("Duration").Value = NullToNothing(oItemNode.Text)
            Case "ZLR.10"
                rsLISTTOOLRECORDS.Fields("Value").Value = NullToNothing(oItemNode.Text)
            Case "ZLR.11"
                rsLISTTOOLRECORDS.Fields("Unit").Value = NullToNothing(oItemNode.Text)
            Case "ZLR.12"
                rsLISTTOOLRECORDS.Fields("RangeScale").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZLR.13"
                rsLISTTOOLRECORDS.Fields("ReferID").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZLR.14"
                rsLISTTOOLRECORDS.Fields("RangeNormalLow").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZLR.15"
                rsLISTTOOLRECORDS.Fields("RangeNormalHigh").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZLR.16"
                rsLISTTOOLRECORDS.Fields("ChartFlag").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZLR.17"
                rsLISTTOOLRECORDS.Fields("Qlink").Value = NullToNothing(oItemNode.Text)
            Case "ZLR.18"
                rsLISTTOOLRECORDS.Fields("QlinkSeq").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZLR.19"
                rsLISTTOOLRECORDS.Fields("Note").Value = NullToNothing(oItemNode.Text)
            Case "ZLR.20"
                rsLISTTOOLRECORDS.Fields("UpdateFlag").Value = NullToNothing(oItemNode.Text)
            Case "ZLR.21"
                rsLISTTOOLRECORDS.Fields("CreatedBy").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZLR.22"
                rsLISTTOOLRECORDS.Fields("CreatedOn").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZLR.23"
                rsLISTTOOLRECORDS.Fields("UpdatedBy").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZLR.24"
                rsLISTTOOLRECORDS.Fields("UpdatedOn").Value = XMLToOracleDate(oItemNode.Text)
        
        End Select
    Next oItemNode
    
    rsLISTTOOLRECORDS.Update
    
    FillRecordsetFromZLRNode = DMErrorCodes.ERROR_SUCCESS

Exit Function

ErrHandler:
    FillRecordsetFromZLRNode = DMErrorCodes.ERROR_EIC_XML_TO_RECORDSET
'    If gobjShared Is Nothing Then
'        Set gobjShared = New CWShared
'    End If
'    gobjShared.ShowVBError Err.Number, Err.Description, "FillRecordsetFromZLRNode Function", "DMBE", vbExclamation
    Err.Clear
End Function

Private Function FillRecordsetFromZLTNode(ByVal oZLTNode As MSXML.IXMLDOMElement, ByVal lngUnitNumber As Long, _
                                        ByRef rsLISTTOOLRECORD As Adodb.Recordset) As DMErrorCodes

    Dim oItemNode As MSXML.IXMLDOMNode
    Dim oChildNode As MSXML.IXMLDOMNode
    
    On Error GoTo ErrHandler
    
    If oZLTNode Is Nothing Then
        Exit Function
    End If
    
    rsLISTTOOLRECORD.AddNew

    For Each oItemNode In oZLTNode.childNodes
        Select Case UCase(oItemNode.nodeName)
            Case "ZLT.1"
                
            Case "ZLT.2"
                rsLISTTOOLRECORD.Fields("DataID").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZLT.3"
                'rsLISTTOOLRECORD.Fields("FacilityNCID").Value = XMLToOracleNumber(oItemNode.Text)
                rsLISTTOOLRECORD.Fields("FacilityNCID").Value = gobjLogon.FacilityNCID
            Case "ZLT.4"
                rsLISTTOOLRECORD.Fields("EncounterNumber").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZLT.5"
                rsLISTTOOLRECORD.Fields("SnoID").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZLT.6"
                Set oChildNode = oItemNode.selectSingleNode("//XCN.1")
                If Not oChildNode Is Nothing Then
                    'Returning SSN
                    rsLISTTOOLRECORD.Fields("ClinicianNCID").Value = XMLToOracleNumber(oChildNode.Text)
                End If
            Case "ZLT.7"
                rsLISTTOOLRECORD.Fields("ClinicNCID").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZLT.8"
                rsLISTTOOLRECORD.Fields("StartDTS").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZLT.9"
                rsLISTTOOLRECORD.Fields("EndDTS").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZLT.10"
                rsLISTTOOLRECORD.Fields("SourceNCID").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZLT.11"
                '<< Begin:SCR #25866;   Developer: Brian Mowbray 07/18/2002 11:42 PM
                If menMessageSource = BMIST_Message Then
                  rsLISTTOOLRECORD.Fields("RTF").AppendChunk NullToNothing(oItemNode.Text)
                Else
                  rsLISTTOOLRECORD.Fields("RTF").AppendChunk DecodeString((NullToNothing(oItemNode.Text)), , True)
                End If
            '>> End: SCR #25866;
            Case "ZLT.12"
                rsLISTTOOLRECORD.Fields("UpdateFlag").Value = NullToNothing(oItemNode.Text)
            Case "ZLT.13"
                rsLISTTOOLRECORD.Fields("CreatedBy").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZLT.14"
                rsLISTTOOLRECORD.Fields("CreatedOn").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZLT.15"
                rsLISTTOOLRECORD.Fields("UpdatedBy").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZLT.16"
                rsLISTTOOLRECORD.Fields("UpdatedOn").Value = XMLToOracleDate(oItemNode.Text)
            ' Begin:SCR #22830;   Developer: MAG 07/19/2002 03:47 PM
            'Case "ZLT.25"
            ' End: SCR #22830;
            ' Begin:SCR #22958;   Developer: Brian Mowbray 11/06/2002 12:47 PM
            Case "ZLT.17"
                rsLISTTOOLRECORD.Fields("OriginalSize").Value = XMLToOracleNumber(oItemNode.Text)
            ' End: SCR #22958;
        
        End Select
    Next oItemNode
    
    rsLISTTOOLRECORD.Update
    
    FillRecordsetFromZLTNode = DMErrorCodes.ERROR_SUCCESS

Exit Function

ErrHandler:
    FillRecordsetFromZLTNode = DMErrorCodes.ERROR_EIC_XML_TO_RECORDSET
'    If gobjShared Is Nothing Then
'        Set gobjShared = New CWShared
'    End If
'    gobjShared.ShowVBError Err.Number, Err.Description, "FillRecordsetFromZLTNode Function", "DMBE", vbExclamation
    Err.Clear
End Function



Private Function FillRecordsetFromZEBNode(ByVal oZEBNode As MSXML.IXMLDOMElement, ByVal lngUnitNumber As Long, _
                                        ByRef rsENC_SIGBLOCK As Adodb.Recordset) As DMErrorCodes

    Dim oItemNode As MSXML.IXMLDOMNode
    
    On Error GoTo ErrHandler
    
    If oZEBNode Is Nothing Then
        Exit Function
    End If
    
    rsENC_SIGBLOCK.AddNew

    For Each oItemNode In oZEBNode.childNodes
        Select Case UCase(oItemNode.nodeName)
            Case "ZEB.1"
                
            Case "ZEB.2"
                rsENC_SIGBLOCK.Fields("DataID").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZEB.3"
                rsENC_SIGBLOCK.Fields("OwnerDataID").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZEB.4"
                rsENC_SIGBLOCK.Fields("UserNCID").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZEB.5"
                rsENC_SIGBLOCK.Fields("Line1").Value = NullToNothing(oItemNode.Text)
            Case "ZEB.6"
                rsENC_SIGBLOCK.Fields("Line2").Value = NullToNothing(oItemNode.Text)
            Case "ZEB.7"
                rsENC_SIGBLOCK.Fields("Line3").Value = NullToNothing(oItemNode.Text)
            Case "ZEB.8"
                rsENC_SIGBLOCK.Fields("DTS").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZEB.9"
                rsENC_SIGBLOCK.Fields("PKI").Value = NullToNothing(oItemNode.Text)
            Case "ZEB.10"
                rsENC_SIGBLOCK.Fields("CreatedBy").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZEB.11"
                rsENC_SIGBLOCK.Fields("CreatedOn").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZEB.12"
                rsENC_SIGBLOCK.Fields("UpdatedBy").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZEB.13"
                rsENC_SIGBLOCK.Fields("UpdatedOn").Value = XMLToOracleDate(oItemNode.Text)
        End Select
    Next oItemNode
    
    rsENC_SIGBLOCK.Update
    
    FillRecordsetFromZEBNode = DMErrorCodes.ERROR_SUCCESS

Exit Function

ErrHandler:
    FillRecordsetFromZEBNode = ERROR_EIC_XML_TO_RECORDSET
    If gobjShared Is Nothing Then
        Set gobjShared = New CWShared
    End If
    gobjShared.ShowVBError Err.Number, Err.Description, "FillRecordsetFromZEBNode Function", "DMBE", vbExclamation
    Err.Clear
End Function

Private Function FillRecordsetFromZERNode(ByVal oZERNode As MSXML.IXMLDOMElement, ByVal lngUnitNumber As Long, _
                                        ByRef rsENC_RTFS As Adodb.Recordset) As DMErrorCodes

    Dim oItemNode As MSXML.IXMLDOMNode
    Dim oChildNode As MSXML.IXMLDOMNode
    Dim sDoc As String
    Dim nOrigSize As Long
    
    On Error GoTo ErrHandler
    
    If oZERNode Is Nothing Then
        Exit Function
    End If
    
    rsENC_RTFS.AddNew

    For Each oItemNode In oZERNode.childNodes
        Select Case UCase(oItemNode.nodeName)
            Case "ZER.1"
                
            Case "ZER.2"
                'rsENC_RTFS.Fields("DataID").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZER.3"
                rsENC_RTFS.Fields("EncounterNumber").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZER.4"
                'rsENC_RTFS.Fields("FacilityNCID").Value = XMLToOracleNumber(oItemNode.Text)
                rsENC_RTFS.Fields("FacilityNCID").Value = gobjLogon.FacilityNCID
            Case "ZER.5"
                rsENC_RTFS.Fields("RTFType").Value = NullToNothing(oItemNode.Text)
            Case "ZER.6"
                rsENC_RTFS.Fields("Enc_RTFSIndex").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZER.7"
                rsENC_RTFS.Fields("NoteID").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZER.8"
                rsENC_RTFS.Fields("Source").Value = NullToNothing(oItemNode.Text)
            Case "ZER.9"
                For Each oChildNode In oItemNode.childNodes
                    Select Case UCase(oChildNode.nodeName)
                        Case "CE.1"
                            rsENC_RTFS.Fields("UserNCID").Value = XMLToOracleNumber(oChildNode.Text)
                        Case "CE.2"
                            Dim strTemp As String
                            
                            strTemp = NullToNothing(oItemNode.Text)
                            If Len(strTemp) > 20 Then
                                strTemp = Left(strTemp, 20)
                            End If
                            rsENC_RTFS.Fields("UserName").Value = strTemp
                            'Fix here
                    End Select
                Next
            Case "ZER.10"
                rsENC_RTFS.Fields("DateAdded").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZER.11"
                rsENC_RTFS.Fields("DTS").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZER.12"
                '<< Begin:SCR #25866;   Developer: Brian Mowbray 07/18/2002 11:43 PM
                sDoc = NullToNothing(oItemNode.Text)
                'rsENC_RTFS.Fields("Doc").AppendChunk DecodeString(NullToNothing(oItemNode.Text), , True)
                '>> End: SCR #25866;
            Case "ZER.13"
                rsENC_RTFS.Fields("Title").Value = NullToNothing(oItemNode.Text)
            Case "ZER.14"
                rsENC_RTFS.Fields("Status").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZER.15"
                rsENC_RTFS.Fields("SensitivityLevel").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZER.16"
                rsENC_RTFS.Fields("CoSignerNCID").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZER.17"
                rsENC_RTFS.Fields("OriginalSize").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZER.18"
                rsENC_RTFS.Fields("InitialSignature").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZER.19"
                rsENC_RTFS.Fields("FinalSignature").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZER.20"
                rsENC_RTFS.Fields("Complete").Value = NullToNothing(oItemNode.Text)
            Case "ZER.21"
                rsENC_RTFS.Fields("Category").Value = NullToNothing(oItemNode.Text)
            Case "ZER.22"
                rsENC_RTFS.Fields("UpdateFlag").Value = NullToNothing(oItemNode.Text)
            Case "ZER.23"
                rsENC_RTFS.Fields("CreatedBy").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZER.24"
                rsENC_RTFS.Fields("CreatedOn").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZER.25"
                rsENC_RTFS.Fields("UpdatedBy").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZER.26"
                rsENC_RTFS.Fields("UpdatedOn").Value = XMLToOracleDate(oItemNode.Text)
        End Select
    Next oItemNode
    
    nOrigSize = Val(rsENC_RTFS.Fields("OriginalSize") & vbNullString)
    
    If nOrigSize = 0 Then
      rsENC_RTFS.Fields("DOC").AppendChunk sDoc
    Else
      rsENC_RTFS.Fields("Doc").AppendChunk DecodeString(sDoc, , True)
    End If
    
    rsENC_RTFS.Update
    
    FillRecordsetFromZERNode = DMErrorCodes.ERROR_SUCCESS

Exit Function

ErrHandler:
    FillRecordsetFromZERNode = ERROR_EIC_XML_TO_RECORDSET
    If gobjShared Is Nothing Then
        Set gobjShared = New CWShared
    End If
    gobjShared.ShowVBError Err.Number, Err.Description, "FillRecordsetFromZERNode Function", "DMBE", vbExclamation
    Err.Clear
End Function

Private Function FillRecordsetFromZEPNode(ByVal oZEPNode As MSXML.IXMLDOMElement, ByVal lngUnitNumber As Long, _
                                        ByRef rsPI_DATA As Adodb.Recordset) As DMErrorCodes

    Dim oItemNode As MSXML.IXMLDOMNode
    
    On Error GoTo ErrHandler
    
    If oZEPNode Is Nothing Then
        Exit Function
    End If
    
    rsPI_DATA.AddNew

    For Each oItemNode In oZEPNode.childNodes
        Select Case UCase(oItemNode.nodeName)
            Case "ZEP.1"
                
            Case "ZEP.2"
                rsPI_DATA.Fields("PI_Data_ID").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZEP.3"
                rsPI_DATA.Fields("SnoID").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZEP.4"
                rsPI_DATA.Fields("Unit_Number").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZEP.5"
                rsPI_DATA.Fields("PI_NCID").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZEP.6"
                rsPI_DATA.Fields("Enc_Num").Value = NullToNothing(oItemNode.Text)
            Case "ZEP.7"
                rsPI_DATA.Fields("Enc_Fac_NCID").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZEP.8"
                rsPI_DATA.Fields("PI_Time").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZEP.9"
                rsPI_DATA.Fields("CreatedBy").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZEP.10"
                rsPI_DATA.Fields("CreatedOn").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZEP.11"
                rsPI_DATA.Fields("UpdatedBy").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZEP.12"
                rsPI_DATA.Fields("UpdatedOn").Value = XMLToOracleDate(oItemNode.Text)
        End Select
    Next oItemNode
    
    rsPI_DATA.Update
    
    FillRecordsetFromZEPNode = DMErrorCodes.ERROR_SUCCESS

Exit Function

ErrHandler:
    FillRecordsetFromZEPNode = ERROR_EIC_XML_TO_RECORDSET
    If gobjShared Is Nothing Then
        Set gobjShared = New CWShared
    End If
    gobjShared.ShowVBError Err.Number, Err.Description, "FillRecordsetFromZEPNode Function", "DMBE", vbExclamation
    Err.Clear
End Function

Private Function FillRecordsetFromZEXNode(ByVal oZEXNode As MSXML.IXMLDOMElement, ByVal lngUnitNumber As Long, _
                                        ByRef rsENCTEXTNOTES As Adodb.Recordset) As DMErrorCodes

    Dim oItemNode As MSXML.IXMLDOMNode
    Dim oChildNode As MSXML.IXMLDOMNode
    
    On Error GoTo ErrHandler
    
    If oZEXNode Is Nothing Then
        Exit Function
    End If
    
    rsENCTEXTNOTES.AddNew

    For Each oItemNode In oZEXNode.childNodes
        Select Case UCase(oItemNode.nodeName)
            Case "ZEX.1"
                
            Case "ZEX.2"
                rsENCTEXTNOTES.Fields("DataID").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZEX.3"
                rsENCTEXTNOTES.Fields("EncounterNumber").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZEX.4"
                'rsENCTEXTNOTES.Fields("FacilityNCID").Value = XMLToOracleNumber(oItemNode.Text)
                 rsENCTEXTNOTES.Fields("FacilityNCID").Value = gobjLogon.FacilityNCID
            Case "ZEX.5"
                Set oChildNode = oItemNode.selectSingleNode("//XCN.1")
                If Not oChildNode Is Nothing Then
                    'Returning SSN
                    rsENCTEXTNOTES.Fields("ModifyClinicanNCID").Value = XMLToOracleNumber(oChildNode.Text)
                End If
            Case "ZEX.6"
                Set oChildNode = oItemNode.selectSingleNode("//XCN.1")
                If Not oChildNode Is Nothing Then
                    'Returning SSN
                    rsENCTEXTNOTES.Fields("CreateClinicanNCID").Value = XMLToOracleNumber(oChildNode.Text)
                End If
            Case "ZEX.7"
                rsENCTEXTNOTES.Fields("StartTime").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZEX.8"
                rsENCTEXTNOTES.Fields("RTF").Value = NullToNothing(oItemNode.Text)
            Case "ZEX.9"
                rsENCTEXTNOTES.Fields("Complete").Value = NullToNothing(oItemNode.Text)
            Case "ZEX.10"
                rsENCTEXTNOTES.Fields("Category").Value = NullToNothing(oItemNode.Text)
            Case "ZEX.11"
                rsENCTEXTNOTES.Fields("Title").Value = NullToNothing(oItemNode.Text)
            Case "ZEX.12"
                rsENCTEXTNOTES.Fields("EndTime").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZEX.13"
                rsENCTEXTNOTES.Fields("CreatedBy").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZEX.14"
                rsENCTEXTNOTES.Fields("CreatedOn").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZEX.15"
                rsENCTEXTNOTES.Fields("UpdatedBy").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZEX.16"
                rsENCTEXTNOTES.Fields("UpdatedOn").Value = XMLToOracleDate(oItemNode.Text)
             Case "ZEX.17"
                rsENCTEXTNOTES.Fields("DTS").Value = XMLToOracleDate(oItemNode.Text)
        End Select
    Next oItemNode
    
    rsENCTEXTNOTES.Update
    
    FillRecordsetFromZEXNode = DMErrorCodes.ERROR_SUCCESS

Exit Function

ErrHandler:
    FillRecordsetFromZEXNode = DMErrorCodes.ERROR_EIC_XML_TO_RECORDSET
    If gobjShared Is Nothing Then
        Set gobjShared = New CWShared
    End If
    gobjShared.ShowVBError Err.Number, Err.Description, "FillRecordsetFromZEXNode Function", "DMBE", vbExclamation
    Err.Clear
End Function

Private Function FillRecordsetFromZESNode(ByVal oZESNode As MSXML.IXMLDOMElement, ByVal lngUnitNumber As Long, _
                                        ByRef rsENC_SECTIONS As Adodb.Recordset) As DMErrorCodes

    Dim oItemNode As MSXML.IXMLDOMNode
    Dim oItemNode1 As MSXML.IXMLDOMNode
    Dim strTemp As String
    Dim sDoc As String
    Dim nOrigSize As Long
    
    On Error GoTo ErrHandler
    
    If oZESNode Is Nothing Then
        Exit Function
    End If
    
    rsENC_SECTIONS.AddNew

    For Each oItemNode In oZESNode.childNodes
        Select Case UCase(oItemNode.nodeName)
            Case "ZES.1"
                
            Case "ZES.2"
                rsENC_SECTIONS.Fields("DataID").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZES.3"
                rsENC_SECTIONS.Fields("EncounterNumber").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZES.4"
                'rsENC_SECTIONS.Fields("FacilityNCID").Value = XMLToOracleNumber(oItemNode.Text)
                rsENC_SECTIONS.Fields("FacilityNCID").Value = gobjLogon.FacilityNCID
            Case "ZES.5"
                rsENC_SECTIONS.Fields("Enc_SectionsIndex").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZES.6"
                rsENC_SECTIONS.Fields("DTS").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZES.7"
                '<< Begin:SCR #25866;   Developer: Brian Mowbray 07/18/2002 11:44 PM
                sDoc = NullToNothing(oItemNode.Text)
                'rsENC_SECTIONS.Fields("DOC").AppendChunk DecodeString(NullToNothing(oItemNode.Text), , True)
                '>> End: SCR #25866;
            Case "ZES.8"
                ' Begin:SCR #40526;   Developer: Mag 08/26/2003 12:06 PM
                For Each oItemNode1 In oItemNode.childNodes
                    Select Case UCase(oItemNode1.nodeName)
                        Case "CE.1"
                            rsENC_SECTIONS.Fields("OwnerNCID").Value = XMLToOracleNumber(oItemNode.Text)
                        Case "CE.2"
                            rsENC_SECTIONS.Fields("OwnerName").Value = NullToNothing(oItemNode.Text)
                    End Select
                Next
              ' End: SCR #40526;
            Case "ZES.9"
                rsENC_SECTIONS.Fields("SensitivityLevel").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZES.10"
                rsENC_SECTIONS.Fields("Status").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZES.11"
                rsENC_SECTIONS.Fields("OriginalSize").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZES.12"
                rsENC_SECTIONS.Fields("LabRTF").Value = NullToNothing(oItemNode.Text)
            Case "ZES.13"
                rsENC_SECTIONS.Fields("RadRTF").Value = NullToNothing(oItemNode.Text)
            Case "ZES.14"
                rsENC_SECTIONS.Fields("CreatedBy").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZES.15"
                rsENC_SECTIONS.Fields("CreatedOn").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZES.16"
                rsENC_SECTIONS.Fields("UpdatedBy").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZES.17"
                rsENC_SECTIONS.Fields("UpdatedOn").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZES.18"
                rsENC_SECTIONS.Fields("Complete").Value = NullToNothing(oItemNode.Text)
            Case "ZES.19"
                rsENC_SECTIONS.Fields("Category").Value = NullToNothing(oItemNode.Text)
            Case "ZES.20"
                rsENC_SECTIONS.Fields("Title").Value = NullToNothing(oItemNode.Text)
            Case "ZES.21"
                rsENC_SECTIONS.Fields("WSRFV").Value = NullToNothing(oItemNode.Text)
            Case "ZES.22"
                Set oItemNode1 = oItemNode.selectSingleNode("CE.5")
                If Not oItemNode1 Is Nothing Then
                    strTemp = NullToNothing(oItemNode1.Text)
                    If Not strTemp = "" Then
                        rsENC_SECTIONS.Fields("DataTypeNCID").Value = XMLToOracleNumber(strTemp)
                    End If
                End If
        End Select
    Next oItemNode
    
    nOrigSize = Val(rsENC_SECTIONS.Fields("OriginalSize") & vbNullString)
    
    If nOrigSize = 0 Then
      rsENC_SECTIONS.Fields("DOC").AppendChunk sDoc
    Else
      rsENC_SECTIONS.Fields("Doc").AppendChunk DecodeString(sDoc, , True)
    End If
    
    rsENC_SECTIONS.Update
    
    FillRecordsetFromZESNode = DMErrorCodes.ERROR_SUCCESS

Exit Function

ErrHandler:
    FillRecordsetFromZESNode = ERROR_EIC_XML_TO_RECORDSET
    If gobjShared Is Nothing Then
        Set gobjShared = New CWShared
    End If
    gobjShared.ShowVBError Err.Number, Err.Description, "FillRecordsetFromZESNode Function", "DMBE", vbExclamation
    Err.Clear
End Function

Private Function FillRecordsetFromZP1Node(ByVal oZP1Node As MSXML.IXMLDOMElement, ByVal lngUnitNumber As Long, _
                                        ByRef rsPROCEDURES As Adodb.Recordset) As DMErrorCodes

    Dim oItemNode As MSXML.IXMLDOMNode
    Dim oChildNode As MSXML.IXMLDOMNode
    On Error GoTo ErrHandler
    
    If oZP1Node Is Nothing Then
        Exit Function
    End If
    
    rsPROCEDURES.AddNew

    For Each oItemNode In oZP1Node.childNodes
        Select Case UCase(oItemNode.nodeName)
            Case "ZP1.1"
                
            Case "ZP1.2"
                rsPROCEDURES.Fields("DataID").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZP1.3"
                rsPROCEDURES.Fields("CPTID").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZP1.4"
                Set oChildNode = oItemNode.selectSingleNode("CE.5")
                'If Not oChildNode Is Nothing Then
                '    rsPROCEDURES.Fields("PointOfCareFacilityNCID").Value = XMLToOracleNumber(oChildNode.Text)
                'End If
                rsPROCEDURES.Fields("PointOfCareFacilityNCID").Value = gobjLogon.FacilityNCID
            Case "ZP1.5"
                'rsPROCEDURES.Fields("StatusNCID").Value = XMLToOracleNumber(oItemNode.Text)
                Set oChildNode = oItemNode.selectSingleNode("CE.5")
                If Not oChildNode Is Nothing Then
                    rsPROCEDURES.Fields("StatusNCID").Value = XMLToOracleNumber(oChildNode.Text)
                End If
            Case "ZP1.6"
                rsPROCEDURES.Fields("NotedDate").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZP1.7"
                rsPROCEDURES.Fields("ProblemListIndicator").Value = oItemNode.Text
            Case "ZP1.8"
                rsPROCEDURES.Fields("ProceduresComment").Value = oItemNode.Text
            Case "ZP1.9"
                rsPROCEDURES.Fields("CreateUserNCID").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZP1.10"
                rsPROCEDURES.Fields("CreateTime").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZP1.11"
                rsPROCEDURES.Fields("ModifyUserNCID").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZP1.12"
                rsPROCEDURES.Fields("ModifyTime").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZP1.13"
                rsPROCEDURES.Fields("UpdateFlag").Value = oItemNode.Text
            Case "ZP1.14"
                rsPROCEDURES.Fields("CDR_DataID").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZP1.15"
                rsPROCEDURES.Fields("CreatedBy").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZP1.16"
                rsPROCEDURES.Fields("CreatedOn").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZP1.17"
                rsPROCEDURES.Fields("UpdatedBy").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZP1.18"
                rsPROCEDURES.Fields("UpdatedOn").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZP1.19"
                rsPROCEDURES.Fields("Unit_Number").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZP1.20"
                Set oChildNode = oItemNode.selectSingleNode("CE.5")
                If Not oChildNode Is Nothing Then
                    rsPROCEDURES.Fields("FACILITYNCID").Value = XMLToOracleNumber(oChildNode.Text)
                End If
        End Select
    Next oItemNode
    
    rsPROCEDURES.Update
    
    FillRecordsetFromZP1Node = DMErrorCodes.ERROR_SUCCESS

Exit Function

ErrHandler:
    FillRecordsetFromZP1Node = DMErrorCodes.ERROR_EIC_XML_TO_RECORDSET
    If gobjShared Is Nothing Then
        Set gobjShared = New CWShared
    End If
    gobjShared.ShowVBError Err.Number, Err.Description, "FillRecordsetFromZP1Node Function", "DMBE", vbExclamation
    Err.Clear
End Function

Private Function FillRecordsetFromPR1Node(ByVal oPR1Node As MSXML.IXMLDOMElement, ByVal lngUnitNumber As Long, _
                                        ByRef rsPROCEDURES As Adodb.Recordset) As DMErrorCodes
    
    Dim oItemNode As MSXML.IXMLDOMNode
    Dim oChildNode As MSXML.IXMLDOMNode
    
    On Error GoTo ErrHandler
    
    If oPR1Node Is Nothing Then
        Exit Function
    End If
    
    rsPROCEDURES.AddNew

    'DataID is the PK
    'The table has EncounterNumber and FacilityNCID as fk and ClinicianNCID and UnitNumber
    
    For Each oItemNode In oPR1Node.childNodes
        Select Case UCase(oItemNode.nodeName)
            Case "PR1.3"
                For Each oChildNode In oItemNode.childNodes
                    Select Case UCase(oChildNode.nodeName)
                        Case "CE.1"
                            rsPROCEDURES.Fields("CPT4ChangableCode").Value = XMLToOracleNumber(oChildNode.Text)
                        Case "CE.5"
                            rsPROCEDURES.Fields("CPTID").Value = XMLToOracleNumber(oChildNode.Text)
                    End Select
                Next
            Case "PR1.5"
                rsPROCEDURES.Fields("OnSetDate").Value = XMLToOracleDate(oItemNode.Text)
            Case "PR1.12"
                Set oChildNode = oItemNode.selectSingleNode("//XCN.1")
                If Not oChildNode Is Nothing Then
                    rsPROCEDURES.Fields("ClinicianNCID").Value = XMLToOracleNumber(oChildNode.Text)
                End If
        End Select
    Next oItemNode
    
    rsPROCEDURES.Update
    
    FillRecordsetFromPR1Node = DMErrorCodes.ERROR_SUCCESS

Exit Function

ErrHandler:
   FillRecordsetFromPR1Node = ERROR_EIC_XML_TO_RECORDSET
   If gobjShared Is Nothing Then
      Set gobjShared = New CWShared
   End If
   gobjShared.ShowVBError Err.Number, Err.Description, "FillRecordsetFromPR1Node Function", "DMBE", vbExclamation
   Err.Clear
End Function

Private Function FillRecordsetFromZEDNode(ByVal oZEDNode As MSXML.IXMLDOMElement, ByVal lngUnitNumber As Long, _
                                        ByRef rsENC_DIAGNOSIS As Adodb.Recordset) As DMErrorCodes

    Dim oItemNode As MSXML.IXMLDOMNode
    
    On Error GoTo ErrHandler
    
    If oZEDNode Is Nothing Then
        Exit Function
    End If
    
    rsENC_DIAGNOSIS.AddNew

    For Each oItemNode In oZEDNode.childNodes
        Select Case UCase(oItemNode.nodeName)
            Case "ZED.1"
                
            Case "ZED.2"
                rsENC_DIAGNOSIS.Fields("DiagnosisNCID").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZED.3"
                'rsENC_DIAGNOSIS.Fields("FacilityNCID").Value = XMLToOracleNumber(oItemNode.Text)
                rsENC_DIAGNOSIS.Fields("FacilityNCID").Value = gobjLogon.FacilityNCID
            Case "ZED.4"
                rsENC_DIAGNOSIS.Fields("EncounterNumber").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZED.5"
                rsENC_DIAGNOSIS.Fields("CreatedBy").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZED.6"
                rsENC_DIAGNOSIS.Fields("CreatedOn").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZED.7"
                rsENC_DIAGNOSIS.Fields("UpdatedBy").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZED.8"
                rsENC_DIAGNOSIS.Fields("UpdatedOn").Value = XMLToOracleDate(oItemNode.Text)
        End Select
    Next oItemNode
    
    rsENC_DIAGNOSIS.Update
    
    FillRecordsetFromZEDNode = DMErrorCodes.ERROR_SUCCESS

Exit Function

ErrHandler:
    FillRecordsetFromZEDNode = ERROR_EIC_XML_TO_RECORDSET
    If gobjShared Is Nothing Then
        Set gobjShared = New CWShared
    End If
    gobjShared.ShowVBError Err.Number, Err.Description, "FillRecordsetFromZEDNode Function", "DMBE", vbExclamation
    Err.Clear
End Function

Private Function GetDiagnosisDescription(ByVal varNCID As Variant) As String
   
   Dim varTemp As Variant
   Dim lngTemp As Long
   Dim strOutPut As String
   Dim objConvertSno As actxSearchSnoMed.IConvert
   strOutPut = ""
   
   If varNCID <> "" Then
      lngTemp = varNCID
      If lngTemp <> 0 Then
         If objConvertSno Is Nothing Then
            Set objConvertSno = New actxSearchSnoMed.IConvert
         End If
         varTemp = objConvertSno.SnoDescription(lngTemp)
         If IsNull(varTemp) Or IsEmpty(varTemp) Then
            strOutPut = ""
         Else
            strOutPut = Trim(varTemp)
         End If
      End If
    End If
   
    GetDiagnosisDescription = strOutPut
   
End Function
Private Function FillRecordsetFromZDRNode(ByVal oZDRNode As MSXML.IXMLDOMElement, ByVal lngUnitNumber As Long, _
                                        ByRef rsDIAGNOSESRELATIONSHIP As Adodb.Recordset) As DMErrorCodes

    Dim oItemNode As MSXML.IXMLDOMNode
    
    On Error GoTo ErrHandler
    
    If oZDRNode Is Nothing Then
        Exit Function
    End If
    
    rsDIAGNOSESRELATIONSHIP.AddNew

    'DIAGNOSES PK is DIAGNOSESDATAID, RELATEDDATAID, AND RELATIONSHIP NCID
                        
    For Each oItemNode In oZDRNode.childNodes
        Select Case UCase(oItemNode.nodeName)
            Case "ZDR.1"
                
            Case "ZDR.2"
                rsDIAGNOSESRELATIONSHIP.Fields("DiagnosisDataID").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZDR.3"
                rsDIAGNOSESRELATIONSHIP.Fields("RelatedDataID").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZDR.4"
                rsDIAGNOSESRELATIONSHIP.Fields("RelationshipNCID").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZDR.5"
                rsDIAGNOSESRELATIONSHIP.Fields("CreatedBy").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZDR.6"
                rsDIAGNOSESRELATIONSHIP.Fields("CreatedOn").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZDR.7"
                rsDIAGNOSESRELATIONSHIP.Fields("UpdatedBy").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZDR.8"
                rsDIAGNOSESRELATIONSHIP.Fields("UpdatedOn").Value = XMLToOracleDate(oItemNode.Text)
            
        End Select
    Next oItemNode
    
    rsDIAGNOSESRELATIONSHIP.Update
    
    FillRecordsetFromZDRNode = DMErrorCodes.ERROR_SUCCESS

Exit Function

ErrHandler:
    FillRecordsetFromZDRNode = ERROR_EIC_XML_TO_RECORDSET
    If gobjShared Is Nothing Then
        Set gobjShared = New CWShared
    End If
    gobjShared.ShowVBError Err.Number, Err.Description, "FillRecordsetFromZDRNode Function", "DMBE", vbExclamation
    Err.Clear
End Function

Private Function FillRecordsetFromZD2Node(ByVal oZD2Node As MSXML.IXMLDOMElement, ByVal lngUnitNumber As Long, _
                                        ByRef rsDIAGNOSES As Adodb.Recordset) As DMErrorCodes

    Dim oItemNode As MSXML.IXMLDOMNode
    Dim oChildNode As MSXML.IXMLDOMNode
    
    On Error GoTo ErrHandler
    
    If oZD2Node Is Nothing Then
        Exit Function
    End If
    
    rsDIAGNOSES.AddNew

    'DIAGNOSES PK is DATAID
    
    'rsDIAGNOSES.fields("UNIT_NUMBER").value = lngUnitNumber
    
    For Each oItemNode In oZD2Node.childNodes
        Select Case UCase(oItemNode.nodeName)
            Case "ZD2.1"
                
            Case "ZD2.2"
                rsDIAGNOSES.Fields("DataID").Value = XMLToOracleNumber(oItemNode.Text)
            
            Case "ZD2.3"
                'rsDIAGNOSES.Fields("FacilityNCID").Value = XMLToOracleNumber(oItemNode.Text)
                rsDIAGNOSES.Fields("FacilityNCID").Value = gobjLogon.FacilityNCID
            Case "ZD2.4"
                rsDIAGNOSES.Fields("Unit_Number").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZD2.5"
                rsDIAGNOSES.Fields("EncounterNumber").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZD2.6"
                rsDIAGNOSES.Fields("SnoID").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZD2.7"
                Set oChildNode = oItemNode.selectSingleNode("CE.5")
                'If Not oChildNode Is Nothing Then
                '    rsDIAGNOSES.Fields("PointOfCareFacilityNCID").Value = XMLToOracleNumber(oChildNode.Text)
                'End If
                rsDIAGNOSES.Fields("PointOfCareFacilityNCID").Value = gobjLogon.FacilityNCID
            Case "ZD2.8"
                'rsDIAGNOSES.Fields("StatusNCID").Value = XMLToOracleNumber(oItemNode.Text)
                Set oChildNode = oItemNode.selectSingleNode("CE.5")
                If Not oChildNode Is Nothing Then
                    rsDIAGNOSES.Fields("StatusNCID").Value = XMLToOracleNumber(oChildNode.Text)
                End If
            Case "ZD2.9"
                rsDIAGNOSES.Fields("NotedDate").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZD2.10"
                'Medcin Description
                'rsDIAGNOSES.fields("ChronicityNCID").value = oItemNode.Text
                Set oChildNode = oItemNode.selectSingleNode("CE.5")
                If Not oChildNode Is Nothing Then
                    rsDIAGNOSES.Fields("ChronicityNCID").Value = XMLToOracleNumber(oChildNode.Text)
                End If
            Case "ZD2.11"
                'NCIDDescription
                'rsDIAGNOSES.fields("SourceNCID").value = oItemNode.Text
                Set oChildNode = oItemNode.selectSingleNode("CE.5")
                If Not oChildNode Is Nothing Then
                    rsDIAGNOSES.Fields("SourceNCID").Value = XMLToOracleNumber(oChildNode.Text)
                End If
            Case "ZD2.12"
                rsDIAGNOSES.Fields("NewFollowUp").Value = oItemNode.Text
            Case "ZD2.13"
                'NCID Description
                'rsDIAGNOSES.fields("BillingCodeNCID").value = oItemNode.Text
                Set oChildNode = oItemNode.selectSingleNode("CE.5")
                If Not oChildNode Is Nothing Then
                    rsDIAGNOSES.Fields("BillingCodeNCID").Value = XMLToOracleNumber(oChildNode.Text)
                End If
            Case "ZD2.14"
                rsDIAGNOSES.Fields("Sensitivity").Value = oItemNode.Text
            Case "ZD2.15"
                rsDIAGNOSES.Fields("ExtendedICD9").Value = oItemNode.Text
            Case "ZD2.16"
                rsDIAGNOSES.Fields("DiagnosesComment").Value = oItemNode.Text
            Case "ZD2.17"
                rsDIAGNOSES.Fields("CreateUserNCID").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZD2.18"
                rsDIAGNOSES.Fields("CreateTime").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZD2.19"
                rsDIAGNOSES.Fields("ModifyUserNCID").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZD2.20"
                rsDIAGNOSES.Fields("ModifyTime").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZD2.21"
                rsDIAGNOSES.Fields("UpdateFlag").Value = oItemNode.Text
            Case "ZD2.22"
                rsDIAGNOSES.Fields("CDR_DataID").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZD2.23"
                rsDIAGNOSES.Fields("CreatedBy").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZD2.24"
                rsDIAGNOSES.Fields("CreatedOn").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZD2.25"
                rsDIAGNOSES.Fields("UpdatedBy").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZD2.26"
                rsDIAGNOSES.Fields("UpdatedOn").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZD2.27"
                rsDIAGNOSES.Fields("ICD9NCID").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZD2.28"
                rsDIAGNOSES.Fields("DiagnosisNCID").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZD2.29"
                rsDIAGNOSES.Fields("OnSetDate").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZD2.30"
                rsDIAGNOSES.Fields("Priority").Value = NullToNothing(oItemNode.Text)
            Case "ZD2.31"
                rsDIAGNOSES.Fields("ClinicianNCID").Value = XMLToOracleNumber(oItemNode.Text)
        End Select
    Next oItemNode
    
    rsDIAGNOSES.Update
    
    FillRecordsetFromZD2Node = DMErrorCodes.ERROR_SUCCESS

Exit Function

ErrHandler:
    FillRecordsetFromZD2Node = ERROR_EIC_XML_TO_RECORDSET
    Err.Clear
End Function

Private Function FillRecordsetFromDG1Node(ByVal oDG1Node As MSXML.IXMLDOMElement, ByVal lngUnitNumber As Long, _
                                        ByRef rsDIAGNOSES As Adodb.Recordset) As DMErrorCodes
    
    Dim oItemNode As MSXML.IXMLDOMNode
    Dim oChildNode As MSXML.IXMLDOMNode
    
    On Error GoTo ErrHandler
    
    If oDG1Node Is Nothing Then
        Exit Function
    End If
    
    rsDIAGNOSES.AddNew

    'Need to fill-in Unit Number for each RS
    'DataID is the PK of DIAGNOSES
    
    For Each oItemNode In oDG1Node.childNodes
        Select Case UCase(oItemNode.nodeName)
            Case "DG1.3"
                For Each oChildNode In oItemNode.childNodes
                    Select Case UCase(oChildNode.nodeName)
                        Case "CE.5"
                            rsDIAGNOSES.Fields("ICD9NCID").Value = XMLToOracleNumber(oChildNode.Text)
                        Case "CE.4"
                            rsDIAGNOSES.Fields("DiagnosisNCID").Value = NullToNothing(oChildNode.Text)
                            'mobjConvertSno.SnoDescription(strTemp)
                    End Select
                Next
            Case "DG1.5"
                rsDIAGNOSES.Fields("OnSetDate").Value = XMLToOracleDate(oItemNode.Text)
            Case "DG1.15"
                rsDIAGNOSES.Fields("Priority").Value = NullToNothing(oItemNode.Text)
            Case "DG1.16"
                Set oChildNode = oItemNode.selectSingleNode("//XCN.1")
                If Not oChildNode Is Nothing Then
                    rsDIAGNOSES.Fields("ClinicianNCID").Value = XMLToOracleNumber(oChildNode.Text)
                End If
        End Select
    Next oItemNode
    
    rsDIAGNOSES.Update
    
    FillRecordsetFromDG1Node = DMErrorCodes.ERROR_SUCCESS

Exit Function

ErrHandler:
    FillRecordsetFromDG1Node = DMErrorCodes.ERROR_EIC_XML_TO_RECORDSET
'    If gobjShared Is Nothing Then
'        Set gobjShared = New CWShared
'    End If
'    gobjShared.ShowVBError Err.Number, Err.Description, "FillRecordsetFromDG1Node Function", "DMBE", vbExclamation
    Err.Clear
End Function

Private Function FillRecordsetFromZG1Node(ByVal oDG1Node As MSXML.IXMLDOMElement, ByVal lngUnitNumber As Long, _
                                        ByRef rsENC_DISPOSITION As Adodb.Recordset) As Long
    
    Dim oItemNode As MSXML.IXMLDOMNode
    
    On Error Resume Next
    
    If oDG1Node Is Nothing Then
        Exit Function
    End If
    
    rsENC_DISPOSITION.AddNew
    
    'Need to fill-in Unit Number for each RS
    'DataID is the PK of DIAGNOSES
    
    For Each oItemNode In oDG1Node.childNodes
        Select Case UCase(oItemNode.nodeName)
            Case "ZG1.1"
                rsENC_DISPOSITION.Fields("INJ_ILL_CAUSE").Value = NullToNothing(oItemNode.Text)
            Case "ZG1.3"
                rsENC_DISPOSITION.Fields("INJ_ILL_CATEGORY").Value = NullToNothing(oItemNode.Text)
            
            Case "ZG1.5"
                rsENC_DISPOSITION.Fields("DISPOSITIONNCID").Value = NullToNothing(oItemNode.Text)
            Case "ZG1.6"
                rsENC_DISPOSITION.Fields("DISPOSITIONTEXT").Value = NullToNothing(oItemNode.Text)
            Case "ZG1.7"
                rsENC_DISPOSITION.Fields("NCID").Value = NullToNothing(oItemNode.Text)
            Case "ZG1.8"
                rsENC_DISPOSITION.Fields("WSDISP").Value = NullToNothing(oItemNode.Text)
                
        End Select
    Next oItemNode
    
    rsENC_DISPOSITION.Update
    
    If Err.Number <> 0 Then
      FillRecordsetFromZG1Node = DMErrorCodes.ERROR_ADDING_BMIST_ENCOUNTER_DATA
      Err.Clear
      Exit Function
    End If
    
    On Error GoTo ErrHandler
    
 
    
    FillRecordsetFromZG1Node = DMErrorCodes.ERROR_SUCCESS

Exit Function

ErrHandler:
    FillRecordsetFromZG1Node = DMErrorCodes.ERROR_EIC_XML_TO_RECORDSET
    If gobjShared Is Nothing Then
        Set gobjShared = New CWShared
    End If
    gobjShared.ShowVBError Err.Number, Err.Description, "FillRecordsetFromZG1Node Function", "DMBE", vbExclamation
    Err.Clear
End Function

Private Function FillRecordsetFromZPVNode(ByVal oZPVNode As MSXML.IXMLDOMElement, ByVal lngUnitNumber As Long, ByRef rsENCOUNTERS As Adodb.Recordset, _
                                        ByRef rsENC_DISPOSITION As Adodb.Recordset, ByRef rsProvider As Adodb.Recordset) As Long
    
    Dim oItemNode As MSXML.IXMLDOMNode
    Dim oChildNode As MSXML.IXMLDOMNode
    Dim nError As Long
    Dim sName As String
    
    On Error GoTo ErrHandler
    
    If oZPVNode Is Nothing Then
        Exit Function
    End If
    
    rsENCOUNTERS.AddNew
    rsENC_DISPOSITION.AddNew
    
    'Need to fill-in Unit Number for each RS
    'ENCOUNTERNUMBER AND FACILITYNCID are the PK of ENCOUNTERS
    'ENCOUNTERNUMBER AND FACILITYNCID are the PK of ENC_DISPOSITION
    '?????????? for MMI_SITE
    
    For Each oItemNode In oZPVNode.childNodes
        Select Case UCase(oItemNode.nodeName)
            Case "ZPV.1"
                
            Case "ZPV.2"
                For Each oChildNode In oItemNode.childNodes
                    Select Case UCase(oChildNode.nodeName)
                        Case "CE.1"
                            rsENCOUNTERS.Fields("TEMPLATEID").Value = XMLToOracleNumber(oChildNode.Text)
                        Case "CE.2"
                            '<< Begin:SCR #36996;   Developer: Brian Mowbray 06/05/2003 08:33 PM
                            rsENCOUNTERS.Fields("TEMPLATEDATA").AppendChunk NullToNothing(oChildNode.Text)
                            '>> End: SCR #36996;
                    End Select
                Next
            Case "ZPV.3"
                ' Begin:SCR #44179;   Developer: Mag 10/22/2003 06:33 PM
                '
                '**************************************************************
                'SCR 69379 - Completed and Updated., SCR 164074 added "Or oItemNode.Text ="
                'If oItemNode.Text = "205657" Then   'Completed Encounters Only
                If oItemNode.Text = "205657" Or oItemNode.Text = "14510309" Then
                '**************************************************************
                  rsENCOUNTERS.Fields("STATUS").Value = XMLToOracleNumber(oItemNode.Text)
                Else
                  nError = DMErrorCodes.ERROR_EIC_NOT_COMPLETE_ENCOUNTER
                  Exit For
                End If
                ' End: SCR #44179;
            Case "ZPV.4"
                'UNIT NUMBER
                rsENCOUNTERS.Fields("UNIT_NUMBER").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZPV.5"
                rsENCOUNTERS.Fields("CATEGORY").Value = NullToNothing(oItemNode.Text)
            Case "ZPV.6"
                '<SCR 22798 CC>
                'rsencounters.fields("PRIMARYDIAGNOSTICNCID").value = nulltonothing(oitemnode.text)
                'IT IS BEING RETURNED AS THE REPRESENTATION (TEXT)
                Set oChildNode = oItemNode.selectSingleNode("CE.5")
                If Not oChildNode Is Nothing Then
                    rsENCOUNTERS.Fields("PRIMARYDIAGNOSISNCID").Value = XMLToOracleNumber(oChildNode.Text)
                End If
            Case "ZPV.7"
                rsENCOUNTERS.Fields("APPTID").Value = NullToNothing(oItemNode.Text)
            Case "ZPV.8"
                rsENCOUNTERS.Fields("SADRStatus").Value = NullToNothing(oItemNode.Text)
            Case "ZPV.9"
                rsENCOUNTERS.Fields("TimeZone").Value = NullToNothing(oItemNode.Text)
            Case "ZPV.10"
                ''rsencounters.fields("APRole1NCID").value = nulltonothing(oitemnode.text)
                'REPRESENTATION
                Set oChildNode = oItemNode.selectSingleNode("CE.5")
                If Not oChildNode Is Nothing Then
                    rsENCOUNTERS.Fields("APRole1NCID").Value = XMLToOracleNumber(oChildNode.Text)
                End If
            Case "ZPV.11"
                'rsencounters.fields("APRole2NCID").value = nulltonothing(oitemnode.text)
                'REPRESENTATION
                Set oChildNode = oItemNode.selectSingleNode("CE.5")
                If Not oChildNode Is Nothing Then
                    rsENCOUNTERS.Fields("APRole2NCID").Value = XMLToOracleNumber(oChildNode.Text)
                End If
            Case "ZPV.12"
                rsENCOUNTERS.Fields("WasUpdated").Value = NullToNothing(oItemNode.Text)
            Case "ZPV.13"
                rsENCOUNTERS.Fields("WhenCosigned").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZPV.14"
                rsENCOUNTERS.Fields("CosignerLine2").Value = NullToNothing(oItemNode.Text)
            Case "ZPV.15"
                rsENCOUNTERS.Fields("CosignerLine3").Value = NullToNothing(oItemNode.Text)
            Case "ZPV.16"
                rsENCOUNTERS.Fields("WhenSigned").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZPV.17"
                rsENCOUNTERS.Fields("SignerLine2").Value = NullToNothing(oItemNode.Text)
            Case "ZPV.18"
                rsENCOUNTERS.Fields("SignerLine3").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZPV.19"
                rsENCOUNTERS.Fields("NoCount").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZPV.20"
                Set oChildNode = oItemNode.selectSingleNode("XCN.5")
                If Not oChildNode Is Nothing Then
                  rsENCOUNTERS.Fields("ADDITIONALPROVIDER1NCID").Value = XMLToOracleNumber(oChildNode.Text)
                  If NullToNothing(oChildNode.Text) <> vbNullString Then
                    rsProvider.Filter = "NCID = " & oChildNode.Text
                    If rsProvider.RecordCount = 0 Then
                      rsProvider.Filter = vbNullString
                      rsProvider.AddNew
                      For Each oChildNode In oItemNode.childNodes
                        Select Case UCase(oChildNode.nodeName)
                        Case "XCN.1" 'SSN
                            rsProvider.Fields("SSN").Value = NullToNothing(oChildNode.Text)
                        Case "XCN.2" 'LAST NAME
                            sName = NullToNothing(oChildNode.Text)
                        Case "XCN.3" 'FIRST NAME
                            If sName = "" Then
                              sName = NullToNothing(oChildNode.Text)
                            Else
                              sName = sName & ", " & NullToNothing(oChildNode.Text)
                            End If
                        Case "XCN.5" 'NCID
                          rsProvider.Fields("NCID") = NullToNothing(oChildNode.Text)
                        End Select
                      Next
                      rsProvider.Fields("NAME") = sName
                      rsProvider.Update
                    End If
                    rsProvider.Filter = vbNullString
                  End If
                Else 'Keep for Old EIC Files
                  Set oChildNode = oItemNode.selectSingleNode("XCN.1")
                  If Not oChildNode Is Nothing Then
                      rsENCOUNTERS.Fields("ADDITIONALPROVIDER1NCID").Value = XMLToOracleNumber(oChildNode.Text)
                  End If
                End If
            Case "ZPV.21"
                rsENCOUNTERS.Fields("ApptType").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZPV.22"
                Set oChildNode = oItemNode.selectSingleNode("XCN.5")
                If Not oChildNode Is Nothing Then
                  rsENCOUNTERS.Fields("ADDITIONALPROVIDER2NCID").Value = XMLToOracleNumber(oChildNode.Text)
                  If NullToNothing(oChildNode.Text) <> vbNullString Then
                    rsProvider.Filter = "NCID = " & oChildNode.Text
                    If rsProvider.RecordCount = 0 Then
                      rsProvider.Filter = vbNullString
                      rsProvider.AddNew
                      For Each oChildNode In oItemNode.childNodes
                        Select Case UCase(oChildNode.nodeName)
                        Case "XCN.1" 'SSN
                            rsProvider.Fields("SSN").Value = NullToNothing(oChildNode.Text)
                        Case "XCN.2" 'LAST NAME
                            sName = NullToNothing(oChildNode.Text)
                        Case "XCN.3" 'FIRST NAME
                            If sName = "" Then
                              sName = NullToNothing(oChildNode.Text)
                            Else
                              sName = sName & ", " & NullToNothing(oChildNode.Text)
                            End If
                        Case "XCN.5" 'NCID
                          rsProvider.Fields("NCID") = NullToNothing(oChildNode.Text)
                        End Select
                      Next
                      rsProvider.Fields("NAME") = sName
                      rsProvider.Update
                    End If
                    rsProvider.Filter = vbNullString
                  End If
                Else 'Keep for Old EIC Files
                  Set oChildNode = oItemNode.selectSingleNode("XCN.1")
                  If Not oChildNode Is Nothing Then
                      rsENCOUNTERS.Fields("ADDITIONALPROVIDER2NCID").Value = XMLToOracleNumber(oChildNode.Text)
                  End If
                End If
                
            Case "ZPV.23"
                rsENCOUNTERS.Fields("AllergyVerifiedByDate").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZPV.24"
                'rsencounters.fields("AllergyVerifiedByNCID").value = XMLToOracleNumber(oitemnode.text)
                'This is representation
                Set oChildNode = oItemNode.selectSingleNode("CE.5")
                If Not oChildNode Is Nothing Then
                    rsENCOUNTERS.Fields("AllergyVerifiedByNCID").Value = XMLToOracleNumber(oChildNode.Text)
                End If
            Case "ZPV.25"
                Set oChildNode = oItemNode.selectSingleNode("CE.5")
                If Not oChildNode Is Nothing Then
                  rsENCOUNTERS.Fields("CoSignerNCID").Value = XMLToOracleNumber(oChildNode.Text)
                  If NullToNothing(oChildNode.Text) <> vbNullString Then
                    rsProvider.Filter = "NCID = " & oChildNode.Text
                    If rsProvider.RecordCount = 0 Then
                      rsProvider.Filter = vbNullString
                      rsProvider.AddNew
                      For Each oChildNode In oItemNode.childNodes
                        Select Case UCase(oChildNode.nodeName)
                        Case "CE.1" 'SSN
                            rsProvider.Fields("SSN").Value = NullToNothing(oChildNode.Text)
                        Case "CE.2" 'FULL NAME
                            rsProvider.Fields("NAME") = NullToNothing(oChildNode.Text)
                        Case "CE.5" 'NCID
                          rsProvider.Fields("NCID") = NullToNothing(oChildNode.Text)
                        End Select
                      Next
                      rsProvider.Update
                    End If
                    rsProvider.Filter = vbNullString
                  End If
                Else 'Keep for Old EIC Files
                  'rsencounters.fields("CoSignerNCID").value = XMLToOracleNumber(oitemnode.text)
                  'This is representation
                  Set oChildNode = oItemNode.selectSingleNode("CE.5")
                  If Not oChildNode Is Nothing Then
                      rsENCOUNTERS.Fields("CoSignerNCID").Value = XMLToOracleNumber(oChildNode.Text)
                  End If
                End If
            Case "ZPV.26"
                Set oChildNode = oItemNode.selectSingleNode("CE.5")
                If Not oChildNode Is Nothing Then
                  rsENCOUNTERS.Fields("WhoSignedNCID").Value = XMLToOracleNumber(oChildNode.Text)
                  If NullToNothing(oChildNode.Text) <> vbNullString Then
                    rsProvider.Filter = "NCID = " & oChildNode.Text
                    If rsProvider.RecordCount = 0 Then
                      rsProvider.Filter = vbNullString
                      rsProvider.AddNew
                      For Each oChildNode In oItemNode.childNodes
                        Select Case UCase(oChildNode.nodeName)
                        Case "CE.1" 'SSN
                            rsProvider.Fields("SSN").Value = NullToNothing(oChildNode.Text)
                        Case "CE.2" 'FULL NAME
                            rsProvider.Fields("NAME") = NullToNothing(oChildNode.Text)
                        Case "CE.5" 'NCID
                          rsProvider.Fields("NCID") = NullToNothing(oChildNode.Text)
                        End Select
                      Next
                      rsProvider.Update
                    End If
                    rsProvider.Filter = vbNullString
                  End If
                Else 'Keep for Old EIC Files
                  'rsencounters.fields("WhoSignedNCID").value = XMLToOracleNumber(oitemnode.text)
                  'This is representation
                  Set oChildNode = oItemNode.selectSingleNode("CE.5")
                  If Not oChildNode Is Nothing Then
                      rsENCOUNTERS.Fields("WhoSignedNCID").Value = XMLToOracleNumber(oChildNode.Text)
                  End If
                End If
            Case "ZPV.27"
                'rsencounters.fields("CoSignerNCID").value = XMLToOracleNumber(oitemnode.text)
                'This is representation
                Set oChildNode = oItemNode.selectSingleNode("CE.5")
                If Not oChildNode Is Nothing Then
                    rsENCOUNTERS.Fields("CoSignerNCID").Value = XMLToOracleNumber(oChildNode.Text)
                End If
            Case "ZPV.28"
                rsENCOUNTERS.Fields("CreatedBy").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZPV.29"
                rsENCOUNTERS.Fields("CreatedOn").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZPV.30"
                rsENCOUNTERS.Fields("UpdatedBy").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZPV.31"
                rsENCOUNTERS.Fields("UpdatedOn").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZPV.32"
                rsENC_DISPOSITION.Fields("EncounterNumber").Value = XMLToOracleNumber(oItemNode.Text) '   = rsENCOUNTERS.Fields("ENCOUNTERNUMBER").Value
                'This is returned as representation
            Case "ZPV.33"
                'Set oChildNode = oItemNode.selectSingleNode("//CE.1")
                'If Not oChildNode Is Nothing Then
                '    rsENC_DISPOSITION.Fields("FacilityNCID").Value = XMLToOracleNumber(oChildNode.Text)
                    'rsENCOUNTERS.Fields("FacilityNCID").Value = XMLToOracleNumber(oItemNode.Text)
                'End If
                rsENC_DISPOSITION.Fields("FacilityNCID").Value = gobjLogon.FacilityNCID
            Case "ZPV.34"
                rsENC_DISPOSITION.Fields("NCID").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZPV.35"
                rsENC_DISPOSITION.Fields("EandMCodeReview").Value = NullToNothing(oItemNode.Text)
            Case "ZPV.36"
                rsENC_DISPOSITION.Fields("EandMReviewedDate").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZPV.37"
                rsENC_DISPOSITION.Fields("EandMReviewedBy").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZPV.38"
                rsENC_DISPOSITION.Fields("EandMReviewed").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZPV.39"
                rsENC_DISPOSITION.Fields("ReviewedByCoder").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZPV.40"
                rsENC_DISPOSITION.Fields("DispAdminOption").Value = NullToNothing(oItemNode.Text)
            Case "ZPV.41"
                For Each oChildNode In oItemNode.childNodes
                    Select Case UCase(oChildNode.nodeName)
                        Case "CE.1"
                            rsENC_DISPOSITION.Fields("DispositionNCID").Value = XMLToOracleNumber(oChildNode.Text)
                        Case "CE.2"
                            rsENC_DISPOSITION.Fields("DispositionText").Value = NullToNothing(oChildNode.Text)
                    End Select
                Next
            Case "ZPV.42"
                rsENC_DISPOSITION.Fields("ItemsDiscussed").Value = NullToNothing(oItemNode.Text)
            Case "ZPV.43"
                rsENC_DISPOSITION.Fields("FollowUpTimeframe").Value = NullToNothing(oItemNode.Text)
            Case "ZPV.44"
                rsENC_DISPOSITION.Fields("FollowUUpComments").Value = NullToNothing(oItemNode.Text)
            Case "ZPV.45"
                rsENC_DISPOSITION.Fields("WSDisp").Value = NullToNothing(oItemNode.Text)
            Case "ZPV.46"
                rsENC_DISPOSITION.Fields("EandMCalc").Value = NullToNothing(oItemNode.Text)
            '<SCR 22810 CC>
            Case "ZPV.47"
                rsENC_DISPOSITION.Fields("INJ_ILL_CAUSE").Value = NullToNothing(oItemNode.Text)
            Case "ZPV.48"
                rsENC_DISPOSITION.Fields("INJ_ILL_CATEGORY").Value = NullToNothing(oItemNode.Text)
            Case "ZPV.49"
                rsENC_DISPOSITION.Fields("DISCUSSEDCOMMENT").Value = NullToNothing(oItemNode.Text)
            Case "ZPV.50"
                rsENCOUNTERS.Fields("EncounterNumber").Value = NullToNothing(oItemNode.Text)
            Case "ZPV.51"
                'rsENCOUNTERS.Fields("FACILITYNCID").Value = NullToNothing(oItemNode.Text)
                rsENCOUNTERS.Fields("FACILITYNCID").Value = gobjLogon.FacilityNCID
            Case "ZPV.52"
                Set oChildNode = oItemNode.selectSingleNode("XCN.5")
                If Not oChildNode Is Nothing Then
                  rsENCOUNTERS.Fields("PRIMARYPROVIDERNCID").Value = NullToNothing(oChildNode.Text)
                  If NullToNothing(oChildNode.Text) <> vbNullString Then
                    rsProvider.Filter = "NCID = " & oChildNode.Text
                    If rsProvider.RecordCount = 0 Then
                      rsProvider.Filter = vbNullString
                      rsProvider.AddNew
                      For Each oChildNode In oItemNode.childNodes
                        Select Case UCase(oChildNode.nodeName)
                        Case "XCN.1" 'SSN
                            rsProvider.Fields("SSN").Value = NullToNothing(oChildNode.Text)
                        Case "XCN.2" 'LAST NAME
                            sName = NullToNothing(oChildNode.Text)
                        Case "XCN.3" 'FIRST NAME
                            If sName = "" Then
                              sName = NullToNothing(oChildNode.Text)
                            Else
                              sName = sName & ", " & NullToNothing(oChildNode.Text)
                            End If
                        Case "XCN.5" 'NCID
                          rsProvider.Fields("NCID") = NullToNothing(oChildNode.Text)
                        End Select
                      Next
                      rsProvider.Fields("NAME") = sName
                      rsProvider.Update
                    End If
                    rsProvider.Filter = vbNullString
                  End If
                Else 'Keep for Old EIC Files
                  rsENCOUNTERS.Fields("PRIMARYPROVIDERNCID").Value = NullToNothing(oItemNode.Text)
                End If
            Case "ZPV.53"
                rsENCOUNTERS.Fields("CLINICNCID").Value = NullToNothing(oItemNode.Text)
            Case "ZPV.54"
                rsENCOUNTERS.Fields("TYPE").Value = NullToNothing(oItemNode.Text)
            Case "ZPV.55"
                rsENCOUNTERS.Fields("STARTDTS").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZPV.56"
                rsENCOUNTERS.Fields("ENDDTS").Value = XMLToOracleDate(oItemNode.Text)
            ' Begin:SCR #26544;   Developer: MAG 08/19/2002 10:01 AM
            Case "ZPV.57"
                rsENCOUNTERS.Fields("DXSnoID").Value = XMLToOracleNumber(oItemNode.Text)
            ' End: SCR #26544;
            Case "ZPV.58"  'Needed to Move ZPV.52 to ZPV.58 to keep backward compatibility with 704.20
                Set oChildNode = oItemNode.selectSingleNode("XCN.5")
                If Not oChildNode Is Nothing Then
                  rsENCOUNTERS.Fields("PRIMARYPROVIDERNCID").Value = NullToNothing(oChildNode.Text)
                  If NullToNothing(oChildNode.Text) <> vbNullString Then
                    rsProvider.Filter = "NCID = " & oChildNode.Text
                    If rsProvider.RecordCount = 0 Then
                      rsProvider.Filter = vbNullString
                      rsProvider.AddNew
                      For Each oChildNode In oItemNode.childNodes
                        Select Case UCase(oChildNode.nodeName)
                        Case "XCN.1" 'SSN
                            rsProvider.Fields("SSN").Value = NullToNothing(oChildNode.Text)
                        Case "XCN.2" 'LAST NAME
                            sName = NullToNothing(oChildNode.Text)
                        Case "XCN.3" 'FIRST NAME
                            If sName = "" Then
                              sName = NullToNothing(oChildNode.Text)
                            Else
                              sName = sName & ", " & NullToNothing(oChildNode.Text)
                            End If
                        Case "XCN.5" 'NCID
                          rsProvider.Fields("NCID") = NullToNothing(oChildNode.Text)
                        End Select
                      Next
                      rsProvider.Fields("NAME") = sName
                      rsProvider.Update
                    End If
                    rsProvider.Filter = vbNullString
                  End If
                End If
        End Select
    Next oItemNode
    
    If nError = 0 Then
      rsENCOUNTERS.Update '- still have to figure out how to get FacilityNCID
      rsENC_DISPOSITION.Update
    
      FillRecordsetFromZPVNode = DMErrorCodes.ERROR_SUCCESS
    Else
      FillRecordsetFromZPVNode = nError
    End If
Exit Function

ErrHandler:
    FillRecordsetFromZPVNode = DMErrorCodes.ERROR_EIC_XML_TO_RECORDSET
    Err.Clear
    Resume Next
End Function


'<SCR 22707 CC - added a parameter>
Private Function FillRecordsetFromZPINode(ByVal oZPINode As MSXML.IXMLDOMElement, ByVal lngUnitNumber As Long, ByRef rsMMI_ID As Adodb.Recordset, _
                                        ByRef rsMMI_NAME As Adodb.Recordset, ByRef rsMMI_SITE As Adodb.Recordset, ByRef rsMMI_GENERIC_ID As Adodb.Recordset, _
                                        ByRef rsMMI_ADDRESS As Adodb.Recordset) As DMErrorCodes

    Dim oItemNode As MSXML.IXMLDOMNode
    Dim oChildNodeLevel1 As MSXML.IXMLDOMNode
    Dim oChildNodeLevel2 As MSXML.IXMLDOMNode
    
    On Error GoTo ErrHandler
    
    If oZPINode Is Nothing Then
        Exit Function
    End If
    
    rsMMI_SITE.AddNew

    'Need to fill-in Unit Number for each RS
    'UnitNumber is the PK of MMI_ID
    'UNIT_NUMNER and NAME_INDEX are the PK of MMI_NAME
    '?????????? for MMI_SITE
    
    rsMMI_SITE("UNIT_NUMBER") = lngUnitNumber
    
    For Each oItemNode In oZPINode.childNodes
        Select Case UCase(oItemNode.nodeName)
            Case "ZPI.1"
                rsMMI_ID("REFERENCE") = oItemNode.Text
            Case "ZPI.2"
                rsMMI_ID("CORRECT_SSN") = oItemNode.Text
            Case "ZPI.3"
                rsMMI_ID("SSN_MISSING_REASON_CID") = XMLToOracleNumber(oItemNode.Text)
            Case "ZPI.4"
                rsMMI_ID("GOVT_HEALTH_NUMBER") = oItemNode.Text
            Case "ZPI.5"
                rsMMI_ID("PASSPORT_NUMBER") = oItemNode.Text
            Case "ZPI.6"
                rsMMI_ID("DOB_EST_FLAG") = oItemNode.Text
            Case "ZPI.7"
                rsMMI_ID("DEATH_DATE") = XMLToOracleDate(oItemNode.Text)
            Case "ZPI.8"
                rsMMI_ID("MAIDEN_NAME") = oItemNode.Text
            Case "ZPI.9"
                rsMMI_ID("PATIENT_FLAG") = oItemNode.Text
            Case "ZPI.10"
                rsMMI_ID("GUARANTOR_FLAG") = oItemNode.Text
            Case "ZPI.11"
                rsMMI_ID("EXPIRED_FLAG") = oItemNode.Text
            Case "ZPI.12"
                rsMMI_ID("ORGAN_DONOR_FLAG") = oItemNode.Text
            Case "ZPI.13"
                rsMMI_ID("ADOPTION_FLAG") = oItemNode.Text
            Case "ZPI.14"
                rsMMI_ID("UNIDENTIFIED_FLAG") = oItemNode.Text
            Case "ZPI.15"
                rsMMI_ID("DEACTIVATE_FLAG") = oItemNode.Text
            Case "ZPI.16"
                rsMMI_ID("PREVIOUS_LINK_ERROR_FLAG") = oItemNode.Text
            Case "ZPI.17"
                rsMMI_ID("INSTITUTIONAL_FLAG") = oItemNode.Text
            Case "ZPI.18"
                rsMMI_ID("BIRTH_DATE_RESOLUTION") = XMLToOracleNumber(oItemNode.Text)
            Case "ZPI.19"
                rsMMI_ID("DEATH_DATE_RESOLUTION") = XMLToOracleNumber(oItemNode.Text)
            Case "ZPI.20"
                rsMMI_ID("DEPLOYLOCATION") = oItemNode.Text
            Case "ZPI.21"
                rsMMI_ID("DEPLOYUNIT") = oItemNode.Text
            Case "ZPI.22"
                rsMMI_ID("DNBI") = oItemNode.Text
            Case "ZPI.23"
                rsMMI_ID("CREATEDINGEMS") = XMLToOracleNumber(oItemNode.Text)
            Case "ZPI.24"
                rsMMI_NAME("Name_Index") = oItemNode.Text
            Case "ZPI.25"
                rsMMI_NAME("Last_Name_2") = oItemNode.Text
            Case "ZPI.26"
                rsMMI_NAME("Middle_Name_2") = oItemNode.Text
            Case "ZPI.27"
                rsMMI_NAME("First_Name_Soundex") = oItemNode.Text
            Case "ZPI.28"
                rsMMI_NAME("Last_Name_Soundex") = oItemNode.Text
            Case "ZPI.29"
                rsMMI_SITE("Radiology_Number") = oItemNode.Text
            Case "ZPI.30.LST"
                '<SCR 22707 CC>
                For Each oChildNodeLevel1 In oItemNode.childNodes
                    rsMMI_GENERIC_ID.AddNew
                    For Each oChildNodeLevel2 In oChildNodeLevel1.childNodes
                        Select Case oChildNodeLevel2.nodeName
                            Case "CE.1"
                                rsMMI_GENERIC_ID.Fields("ID_TYPE_NCID").Value = XMLToOracleNumber(oChildNodeLevel2.Text)
                            Case "CE.2"
                                rsMMI_GENERIC_ID.Fields("ID_VALUE").Value = NullToNothing(oChildNodeLevel2.Text)
                            Case "CE.3"
                                rsMMI_GENERIC_ID.Fields("UNIT_NUMBER").Value = lngUnitNumber 'XMLToOracleNumber(oChildNodeLevel2.Text)
                        End Select
                    Next
                    rsMMI_GENERIC_ID.Update
                Next
                '</SCR 22707 CC>
            Case "ZPI.31"
                If oItemNode.childNodes.Length > 0 Then
                    rsMMI_ADDRESS.AddNew
                    For Each oChildNodeLevel1 In oItemNode.childNodes
                        Select Case oChildNodeLevel1.nodeName
                            Case "XAD.1"
                                rsMMI_ADDRESS.Fields("STREET_1").Value = NullToNothing(oChildNodeLevel1.Text)
                            Case "XAD.2"
                                rsMMI_ADDRESS.Fields("STREET_2").Value = NullToNothing(oChildNodeLevel1.Text)
                            Case "XAD.3"
                                rsMMI_ADDRESS.Fields("CITY").Value = NullToNothing(oChildNodeLevel1.Text)
                            Case "XAD.4"
                                rsMMI_ADDRESS.Fields("STATE").Value = NullToNothing(oChildNodeLevel1.Text)
                            Case "XAD.5"
                                rsMMI_ADDRESS.Fields("POST_CODE").Value = NullToNothing(oChildNodeLevel1.Text)
                            Case "XAD.6"
                                rsMMI_ADDRESS.Fields("COUNTRY").Value = NullToNothing(oChildNodeLevel1.Text)
                            Case "XAD.9"
                                rsMMI_ADDRESS.Fields("COUNTY").Value = NullToNothing(oChildNodeLevel1.Text)
                        End Select
                    Next
                    rsMMI_ADDRESS.Fields("UNIT_NUMBER").Value = lngUnitNumber
                    rsMMI_ADDRESS.Fields("ADDRESS_TYPE_CID").Value = NCID_WORK_ADDRESS
                    rsMMI_ADDRESS.Update
                End If
            'Begin:SCR #44724;    Developer:  Mag 11/03/2003 08:39 PM
            Case "ZPI.32"
              If Len(oItemNode.Text) > 0 Then
                rsMMI_ID("Foreign_ID_Flag") = XMLToOracleNumber(oItemNode.Text)
              Else
                rsMMI_ID("Foreign_ID_Flag") = 0
              End If
            ' End: SCR #44724
        End Select
    Next oItemNode
    
    rsMMI_ID.Update
    rsMMI_NAME.Update
    rsMMI_SITE.Update
    
    FillRecordsetFromZPINode = DMErrorCodes.ERROR_SUCCESS

Exit Function

ErrHandler:
    FillRecordsetFromZPINode = DMErrorCodes.ERROR_EIC_XML_TO_RECORDSET
    If gobjShared Is Nothing Then
        Set gobjShared = New CWShared
    End If
    gobjShared.ShowVBError Err.Number, Err.Description, "FillRecordsetFromZPINode Function", "DMBE", vbExclamation
    Err.Clear
End Function
'<SCR 22858, 22764 CC>
Private Function FillRecordsetFromPIDNode(ByVal oPIDNode As MSXML.IXMLDOMElement, ByVal lngUnitNumber As Long, ByRef rsMMI_ID As Adodb.Recordset, _
                                        ByRef rsMMI_NAME As Adodb.Recordset, ByRef rsMMI_GENERIC_ID As Adodb.Recordset, _
                                        ByRef rsMMI_ADDRESS As Adodb.Recordset, ByRef rsMMI_PHONE As Adodb.Recordset, _
                                        ByVal enmWhoseMessage As XMLMessageSource) As DMErrorCodes
    
    'Dim lngFakeUnitNumber As Long
    Dim oItemNode As MSXML.IXMLDOMNode
    Dim oChildNodeLevel1 As MSXML.IXMLDOMNode
    Dim oChildNodeLevel2 As MSXML.IXMLDOMNode
    Dim oChildNodeLevel3 As MSXML.IXMLDOMNode
    Dim rsSAMS As Adodb.Recordset
    Dim strSQL As String
    Dim strSSN As String
    Dim strFMP As String
    Dim strSponsorSSN As String
    Dim blnNCIDIsFound As Boolean
    
    On Error GoTo ErrHandler
    
    'lngFakeUnitNumber = CLng(Time) + Hour(Time) + Minute(Time) + Second(Time)

    If oPIDNode Is Nothing Then
        Exit Function
    End If
    
    rsMMI_ID.AddNew

    


    For Each oItemNode In oPIDNode.childNodes
        Select Case UCase(oItemNode.nodeName)
                
            Case "PID.2"
                Set oChildNodeLevel1 = oItemNode.selectSingleNode("//CX.1")
                'If Not oChildNodeLevel1 Is Nothing Then
                '    lngUnitNumber = XMLToOracleNumber(oChildNodeLevel1.Text)
                'Else
                '    lngUnitNumber = lngFakeUnitNumber
                'End If
                rsMMI_ID.Fields("UNIT_NUMBER").Value = lngUnitNumber
            Case "PID.3.LST"
                For Each oChildNodeLevel1 In oItemNode.childNodes
                    If UCase(oChildNodeLevel1.nodeName) = "PID.3" Then
                        Set oChildNodeLevel2 = oChildNodeLevel1.selectSingleNode("CX.5")
                        If Not oChildNodeLevel2 Is Nothing Then
                            Select Case UCase(oChildNodeLevel2.Text)
                                Case "SSN"
                                    ' Begin:SCR #28377;   Developer: MAG 10/10/2002 03:43 PM
                                    Set oChildNodeLevel3 = oChildNodeLevel1.selectSingleNode("CX.1")
'                                    rsMMI_ID.Fields("SSN").Value = Replace(oChildNodeLevel3.Text, "-", "")

                                    
                                    '<< Begin:SCR #27601;   Developer: Brian Mowbray 09/12/2002 04:04 PM
                                    'According to the customer, the SSN will ALWAYS be the SPONSOR SSN, not the patient's SSN when coming from SAMS
                                    ' We need to change this so we do not rely on any TMIP segments to pull data back in
                                    strSponsorSSN = Replace(oChildNodeLevel3.Text, "-", "")
                                    '>> End: SCR #27601;
                                    ' End: SCR #28377;
                                Case "FMP"
                                    Set oChildNodeLevel3 = oChildNodeLevel1.selectSingleNode("CX.1")
                                    rsMMI_GENERIC_ID.AddNew
                                        rsMMI_GENERIC_ID.Fields("ID_TYPE_NCID").Value = NCID_PATIENT_FMP
                                        strFMP = NullToNothing(oChildNodeLevel3.Text)
                                        If strFMP = "" Then
                                            strFMP = "20"
                                        End If
                                        rsMMI_GENERIC_ID.Fields("ID_VALUE").Value = strFMP
                                        rsMMI_GENERIC_ID.Fields("UNIT_NUMBER").Value = lngUnitNumber
                                        'strFMP = oChildNodeLevel3.Text
                                    rsMMI_GENERIC_ID.Update
                                '<< Begin:SCR #27601;   Developer: Brian Mowbray 09/12/2002 03:55 PM
                                'Removed per request of the customer
                                ' Needed to Add SPONSORSSN Node
                                'Case "SSSN"
                                '    Set oChildNodeLevel3 = oChildNodeLevel1.selectSingleNode("CX.1")
                                '    strSponsorSSN = Replace(oChildNodeLevel3.Text, "-", "")
                                '>> End: SCR #27601;
                            End Select
                        End If
                    End If
                Next
                
                ' Begin:SCR #26018;   Developer: MAG 07/30/2002 03:57 PM
                If strSSN = "" Then
                    If strFMP = "" Then
                        strFMP = "20"
                    End If
                    
                    If CInt(strFMP) = 20 And strSponsorSSN <> "" Then
                        rsMMI_ID.Fields("SSN").Value = strSponsorSSN
                        strSSN = strSponsorSSN
                    Else
                        Set oChildNodeLevel1 = oItemNode.selectSingleNode("//PID.19")
                        If Not oChildNodeLevel1 Is Nothing Then
                            strSSN = Replace(oChildNodeLevel1.Text, "-", "")
                            rsMMI_ID.Fields("SSN").Value = strSSN
                        End If
                    End If
                End If
                
                ' End: SCR #26018;
                
                rsMMI_GENERIC_ID.AddNew
                    rsMMI_GENERIC_ID.Fields("ID_TYPE_NCID").Value = NCID_FMP_SPONSOR_SSN
                    ' Begin:SCR #44087;   Developer: Mag 10/21/2003 12:25 PM
                    rsMMI_GENERIC_ID.Fields("ID_VALUE").Value = strFMP & "/" & Replace(strSponsorSSN, "-", "")
                    ' End: SCR #44087;
                    rsMMI_GENERIC_ID.Fields("UNIT_NUMBER").Value = lngUnitNumber
                rsMMI_GENERIC_ID.Update
                rsMMI_GENERIC_ID.AddNew
                    rsMMI_GENERIC_ID.Fields("ID_TYPE_NCID").Value = NCID_SPONSOR_SSN
                    rsMMI_GENERIC_ID.Fields("ID_VALUE").Value = strSponsorSSN
                    rsMMI_GENERIC_ID.Fields("UNIT_NUMBER").Value = lngUnitNumber
                rsMMI_GENERIC_ID.Update

            Case "PID.5.LST"
                For Each oChildNodeLevel1 In oItemNode.childNodes
                    If UCase(oChildNodeLevel1.nodeName) = "PID.5" Then
                        rsMMI_NAME.AddNew
                        rsMMI_NAME.Fields("UNIT_NUMBER").Value = lngUnitNumber
                        For Each oChildNodeLevel2 In oChildNodeLevel1.childNodes
                            '<< Begin:SCR #25949;   Developer: Brian Mowbray 08/13/2002 01:07 PM
                            Select Case UCase(oChildNodeLevel2.nodeName)
                                Case "XPN.1"
                                    rsMMI_NAME.Fields("LAST_NAME").Value = UCase(Left$(oChildNodeLevel2.Text, 26))
                                Case "XPN.2"
                                    rsMMI_NAME.Fields("FIRST_NAME").Value = UCase(Left$(oChildNodeLevel2.Text, 16))
                                Case "XPN.3"
                                    rsMMI_NAME.Fields("MIDDLE_NAME").Value = UCase(Left$(oChildNodeLevel2.Text, 16))
                                Case "XPN.4"
                                    rsMMI_NAME.Fields("SUFFIX").Value = UCase(Left$(oChildNodeLevel2.Text, 4))
                                Case "XPN.5"
                                    rsMMI_NAME.Fields("TITLE").Value = UCase(Left$(oChildNodeLevel2.Text, 16))
                                Case "XPN.6"
                                    rsMMI_NAME.Fields("DEGREE").Value = UCase(Left$(oChildNodeLevel2.Text, 10))
                            End Select
                        '>> End: SCR #25949;
                        Next
                        rsMMI_NAME.Update
                    End If
                Next
            Case "PID.6"
                Set oChildNodeLevel1 = oItemNode.selectSingleNode("//XPN.1")
                If Not oChildNodeLevel1 Is Nothing Then
                    rsMMI_ID.Fields("MAIDEN_NAME").Value = UCase(oChildNodeLevel1.Text)
                End If
            Case "PID.7"
                rsMMI_ID.Fields("BIRTH_DATE").Value = XMLToOracleDate(oItemNode.Text)
            Case "PID.8"
                'SCR 22851 CC
                Select Case UCase(oItemNode.Text)
                    Case Is = "MALE", "M"
                        rsMMI_ID.Fields("SEX").Value = "M"
                    Case Is = "FEMALE", "F"
                        rsMMI_ID.Fields("SEX").Value = "F"
                    Case Is = "BOTH", "B"
                        rsMMI_ID.Fields("SEX").Value = "B"
                    Case Is = "OTHER", "O"
                        rsMMI_ID.Fields("SEX").Value = "O"
                    Case Else
                        rsMMI_ID.Fields("SEX").Value = "U"
                End Select
            Case "PID.9"
                'Also a name
            Case "PID.10.LST"
                'Race
                Set oChildNodeLevel1 = oItemNode.selectSingleNode("PID.10")
                If Not oChildNodeLevel1 Is Nothing Then
                    
                    blnNCIDIsFound = False
                    Set oChildNodeLevel2 = oChildNodeLevel1.selectSingleNode("CE.6")
                    If Not oChildNodeLevel2 Is Nothing Then
                        If Trim(oChildNodeLevel2.Text) = "99NCI" Then
                            blnNCIDIsFound = True
                        End If
                    End If
                    
                    If blnNCIDIsFound = True Then
                        Set oChildNodeLevel2 = oChildNodeLevel1.selectSingleNode("CE.4")
                        If Not oChildNodeLevel2 Is Nothing Then
                            rsMMI_ID.Fields("RACE_CID").Value = XMLToOracleNumber(oChildNodeLevel2.Text)
                        End If
                    Else
                        Set oChildNodeLevel2 = oChildNodeLevel1.selectSingleNode("CE.2")
    
                        
                        If Not oChildNodeLevel2 Is Nothing Then
                            strSQL = "SELECT CHCSII_NCID FROM TMIP_SAMS_NCID_MAPPING "
                            strSQL = strSQL & "WHERE SAMS_DESCRIPTION = '" & oChildNodeLevel2.Text & "'"
                            
                            If m_objGEMSDAS Is Nothing Then Set m_objGEMSDAS = oDatabase.CHCSII_DAS(Auto) 'New GEMS_DAS
                                
                            Set rsSAMS = New Adodb.Recordset
                            Set rsSAMS = m_objGEMSDAS.ExecuteSQL(strSQL)
                            If Not rsSAMS.EOF Then
                                rsMMI_ID.Fields("RACE_CID").Value = rsSAMS.Fields(0).Value
                            Else
                                'Unknown
                                rsMMI_ID.Fields("RACE_CID").Value = 146
                            End If
                        End If
                    End If
                End If
            Case "PID.11.LST"
                For Each oChildNodeLevel1 In oItemNode.childNodes
                    If UCase(oChildNodeLevel1.nodeName) = "PID.11" Then
                        rsMMI_ADDRESS.AddNew
                        rsMMI_ADDRESS.Fields("UNIT_NUMBER").Value = lngUnitNumber
                        rsMMI_ADDRESS.Fields("ADDRESS_TYPE_CID").Value = NCID_HOME_ADDRESS
                        For Each oChildNodeLevel2 In oChildNodeLevel1.childNodes
                            '<< Begin:SCR #25949;   Developer: Brian Mowbray 08/13/2002 01:07 PM
                            Select Case UCase(oChildNodeLevel2.nodeName)
                                Case "XAD.1"
                                    rsMMI_ADDRESS.Fields("STREET_1").Value = Left$(oChildNodeLevel2.Text, 30)
                                Case "XAD.2"
                                    rsMMI_ADDRESS.Fields("STREET_2").Value = Left$(oChildNodeLevel2.Text, 30)
                                Case "XAD.3"
                                    rsMMI_ADDRESS.Fields("CITY").Value = Left$(oChildNodeLevel2.Text, 34)
                                Case "XAD.4"
                                    rsMMI_ADDRESS.Fields("STATE").Value = Left$(oChildNodeLevel2.Text, 10)
                                Case "XAD.5"
                                    rsMMI_ADDRESS.Fields("POST_CODE").Value = Left$(oChildNodeLevel2.Text, 12)
                                Case "XAD.6"
                                    rsMMI_ADDRESS.Fields("COUNTRY").Value = Left$(oChildNodeLevel2.Text, 12)
                                Case "XAD.9"
                                    rsMMI_ADDRESS.Fields("COUNTY").Value = Left$(oChildNodeLevel2.Text, 34)
                            End Select
                        '>> End: SCR #25949;
                        Next
                        rsMMI_ADDRESS.Update
                    End If
                Next
            Case "PID.13.LST"
                'Home phone
                Dim strTemp As String
                For Each oChildNodeLevel1 In oItemNode.childNodes
                    If UCase(oChildNodeLevel1.nodeName) = "PID.13" Then
                        rsMMI_PHONE.AddNew
                        rsMMI_PHONE.Fields("UNIT_NUMBER").Value = lngUnitNumber
                        rsMMI_PHONE.Fields("PHONE_TYPE_CID").Value = NCID_HOME_PHONE
                        
                        For Each oChildNodeLevel2 In oChildNodeLevel1.childNodes '<SCR 25148 CC>
                            '<< Begin:SCR #25949;   Developer: Brian Mowbray 08/13/2002 01:07 PM
                            Select Case oChildNodeLevel2.nodeName
                                Case Is = "XTN.5"
                                    rsMMI_PHONE.Fields("COUNTRY_CODE").Value = Left$(oChildNodeLevel2.Text, 16)
                                Case Is = "XTN.6"
                                    rsMMI_PHONE.Fields("AREA_CODE").Value = Left$(oChildNodeLevel2.Text, 16)
                                Case Is = "XTN.7"
                                    rsMMI_PHONE.Fields("LOCAL_NUMBER").Value = Left$(oChildNodeLevel2.Text, 16)
                                Case Is = "XTN.8"
                                    rsMMI_PHONE.Fields("INTERNAL_EXTENSION").Value = Left$(oChildNodeLevel2.Text, 16)
                            End Select
                        '>> End: SCR #25949;
                        Next
                        
'                        Set oChildNodeLevel2 = oChildNodeLevel1.selectSingleNode("XTN.1")
'                        If Not oChildNodeLevel2 Is Nothing Then
'                            strTemp = oChildNodeLevel2.Text
'                            If Not strTemp = "" Then
'                                strTemp = Replace(strTemp, "(", "")
'                                strTemp = Replace(strTemp, ")", "")
'                                strTemp = Replace(strTemp, "[", "")
'                                strTemp = Replace(strTemp, "]", "")
'                                strTemp = Replace(strTemp, "-", "")
'                                strTemp = Replace(strTemp, " ", "")
'                                Select Case Len(strTemp)
'                                    Case 10
'                                        rsMMI_PHONE.Fields("AREA_CODE").Value = Left(strTemp, 3)
'                                        rsMMI_PHONE.Fields("LOCAL_NUMBER").Value = Right(strTemp, 7)
'                                    Case 7
'                                        rsMMI_PHONE.Fields("LOCAL_NUMBER").Value = Right(strTemp, 7)
'                                    Case Else
'                                        intI = InStr(1, "x", strTemp)
'                                        If intI > 0 Then
'                                            rsMMI_PHONE.Fields("INTERNAL_EXTENSION").Value = Right(strTemp, Len(strTemp) - intI)
'                                            strTemp = Left(strTemp, Len(strTemp) - intI)
'                                        Else
'                                            Select Case Len(strTemp)
'                                                Case 10
'                                                    rsMMI_PHONE.Fields("AREA_CODE").Value = Left(strTemp, 3)
'                                                    rsMMI_PHONE.Fields("LOCAL_NUMBER").Value = Right(strTemp, 7)
'                                                Case 7
'                                                    rsMMI_PHONE.Fields("LOCAL_NUMBER").Value = Right(strTemp, 7)
'                                                Case Is > 10
'                                                    rsMMI_PHONE.Fields("LOCAL_NUMBER").Value = Right(strTemp, 7)
'                                                    strTemp = Left(strTemp, Len(strTemp) - 7)
'                                                    rsMMI_PHONE.Fields("AREA_CODE").Value = Right(strTemp, 3)
'                                                    rsMMI_PHONE.Fields("COUNTRY_CODE").Value = Left(strTemp, Len(strTemp) - 3)
'                                            End Select
'                                        End If
'                                End Select
'                            End If
'                        End If
                        
                        rsMMI_PHONE.Update
                    
                    End If
                Next
            Case "PID.14.LST"
                'Work Phone
                For Each oChildNodeLevel1 In oItemNode.childNodes
                    If UCase(oChildNodeLevel1.nodeName) = "PID.14" Then
                        rsMMI_PHONE.AddNew
                        rsMMI_PHONE.Fields("UNIT_NUMBER").Value = lngUnitNumber
                        rsMMI_PHONE.Fields("PHONE_TYPE_CID").Value = NCID_WORK_PHONE
                        
                        For Each oChildNodeLevel2 In oChildNodeLevel1.childNodes '<SCR 25148 CC>
                            '<< Begin:SCR #25949;   Developer: Brian Mowbray 08/13/2002 01:07 PM
                            Select Case oChildNodeLevel2.nodeName
                                Case Is = "XTN.5"
                                    rsMMI_PHONE.Fields("COUNTRY_CODE").Value = Left$(oChildNodeLevel2.Text, 16)
                                Case Is = "XTN.6"
                                    rsMMI_PHONE.Fields("AREA_CODE").Value = Left$(oChildNodeLevel2.Text, 16)
                                Case Is = "XTN.7"
                                    rsMMI_PHONE.Fields("LOCAL_NUMBER").Value = Left$(oChildNodeLevel2.Text, 16)
                                Case Is = "XTN.8"
                                    rsMMI_PHONE.Fields("INTERNAL_EXTENSION").Value = Left$(oChildNodeLevel2.Text, 16)
                            End Select
                        '>> End: SCR #25949;
                        Next
                        
'                        Set oChildNodeLevel2 = oChildNodeLevel1.selectSingleNode("XTN.1")
'                        If Not oChildNodeLevel2 Is Nothing Then
'                            strTemp = oChildNodeLevel2.Text
'                            If Not strTemp = "" Then
'                                strTemp = Replace(strTemp, "(", "")
'                                strTemp = Replace(strTemp, ")", "")
'                                strTemp = Replace(strTemp, "[", "")
'                                strTemp = Replace(strTemp, "]", "")
'                                strTemp = Replace(strTemp, "-", "")
'                                Select Case Len(strTemp)
'                                    Case 10
'                                        rsMMI_PHONE.Fields("AREA_CODE").Value = Left(strTemp, 3)
'                                        rsMMI_PHONE.Fields("LOCAL_NUMBER").Value = Right(strTemp, 7)
'                                    Case 7
'                                        rsMMI_PHONE.Fields("LOCAL_NUMBER").Value = Right(strTemp, 7)
'                                    Case Else
'                                        intI = InStr(1, "x", strTemp)
'                                        If intI > 0 Then
'                                            rsMMI_PHONE.Fields("INTERNAL_EXTENSION").Value = Right(strTemp, Len(strTemp) - intI)
'                                            strTemp = Left(strTemp, Len(strTemp) - intI)
'                                        Else
'                                            Select Case Len(strTemp)
'                                                Case 10
'                                                    rsMMI_PHONE.Fields("AREA_CODE").Value = Left(strTemp, 3)
'                                                    rsMMI_PHONE.Fields("LOCAL_NUMBER").Value = Right(strTemp, 7)
'                                                Case 7
'                                                    rsMMI_PHONE.Fields("LOCAL_NUMBER").Value = Right(strTemp, 7)
'                                                Case Is > 10
'                                                    rsMMI_PHONE.Fields("LOCAL_NUMBER").Value = Right(strTemp, 7)
'                                                    strTemp = Left(strTemp, Len(strTemp) - 7)
'                                                    rsMMI_PHONE.Fields("AREA_CODE").Value = Right(strTemp, 3)
'                                                    rsMMI_PHONE.Fields("COUNTRY_CODE").Value = Left(strTemp, Len(strTemp) - 3)
'                                            End Select
'                                        End If
'                                End Select
'                            End If
                        
                        'End If
                        rsMMI_PHONE.Update
                    End If
                Next
            Case "PID.16"

                'Marital Status
                Set oChildNodeLevel1 = oItemNode.selectSingleNode("CE.6")
                If Not oChildNodeLevel1 Is Nothing Then
                    If oChildNodeLevel1.Text = "99NCI" Then
                        Set oChildNodeLevel1 = oItemNode.selectSingleNode("CE.4")
                        If Not oChildNodeLevel1 Is Nothing Then
                            strTemp = NullToNothing(oChildNodeLevel1.Text)
                            If Not strTemp = "" Then
                                rsMMI_ID.Fields("MARITAL_STATUS_CID").Value = strTemp
                            End If
                        End If
                    Else
                        Set oChildNodeLevel1 = oItemNode.selectSingleNode("CE.2")
                        If Not oChildNodeLevel1 Is Nothing Then
                            
                            Dim oRS As Adodb.Recordset
                            
                            strTemp = NullToNothing(oChildNodeLevel1.Text)
                            
                            strSQL = "SELECT NCID FROM NCIDS WHERE DOMAINNCID = " & NCID_MARTIAL_STATUS
                            strSQL = strSQL & " AND UPPER(DESCRIPTION) LIKE '" & UCase(Left(strTemp, 3)) & "%'"
                            
                            If oDas Is Nothing Then Set oDas = oDatabase.CHCSII_DAS(Auto) 'New GEMS_DAS
                            Set oRS = oDas.OpenRecordset(strSQL)
                            If Not oRS.EOF Then
                                strTemp = oRS("NCID")
                                rsMMI_ID.Fields("MARITAL_STATUS_CID").Value = strTemp
                            End If
                            oRS.Close
                        End If
                    End If
                End If
                
            Case "PID.17"
                'Religion
                
                blnNCIDIsFound = False
                Set oChildNodeLevel1 = oItemNode.selectSingleNode("CE.6")
                If Not oChildNodeLevel1 Is Nothing Then
                    If Trim(oChildNodeLevel1.Text) = "99NCI" Then
                        blnNCIDIsFound = True
                    End If
                End If
                
                If blnNCIDIsFound = True Then
                    Set oChildNodeLevel1 = oItemNode.selectSingleNode("CE.4")
                    If Not oChildNodeLevel1 Is Nothing Then
                        rsMMI_ID.Fields("RELIGION_CID").Value = XMLToOracleNumber(oChildNodeLevel1.Text)
                    End If
                Else
                    Set oChildNodeLevel1 = oItemNode.selectSingleNode("CE.2")
                    If Not oChildNodeLevel1 Is Nothing Then
                        strSQL = "SELECT CHCSII_NCID FROM TMIP_SAMS_NCID_MAPPING "
                        strSQL = strSQL & "WHERE SAMS_DESCRIPTION = '" & oChildNodeLevel1.Text & "'"
                        
                        If m_objGEMSDAS Is Nothing Then Set m_objGEMSDAS = oDatabase.CHCSII_DAS(Auto) 'New GEMS_DAS
                            
                        Set rsSAMS = New Adodb.Recordset
                        Set rsSAMS = m_objGEMSDAS.ExecuteSQL(strSQL)
                        If Not rsSAMS.EOF Then
                            rsMMI_ID.Fields("RELIGION_CID").Value = rsSAMS.Fields(0).Value
                        Else
                            rsMMI_ID.Fields("RELIGION_CID").Value = 180 'Unknown
                        End If
                    End If
                End If
            Case "PID.19"
                ' Begin:SCR #28377;   Developer: MAG 10/10/2002 03:44 PM
                rsMMI_ID.Fields("SSN").Value = Replace(oItemNode.Text, "-", "")
                If strSponsorSSN = "" Then
                    strSponsorSSN = rsMMI_ID.Fields("SSN")
                End If
                ' End: SCR #28377;
            
                ' Begin:SCR #44724;  Developer: Mag 11/03/2003 08:40 PM
                If Len(rsMMI_ID.Fields("SSN") & vbNullString) <> 9 Then
                  rsMMI_ID("Foreign_ID_Flag") = 1
                End If
                ' End: SCR #44724;
            Case "PID.22"
                'Ethnicity
                'No Use to US
        End Select
    Next
    rsMMI_ID.Update
    
    FillRecordsetFromPIDNode = ERROR_SUCCESS
    
Exit Function

ErrHandler:
    FillRecordsetFromPIDNode = ERROR_EIC_XML_TO_RECORDSET
'    If gobjShared Is Nothing Then
'        Set gobjShared = New CWShared
'    End If
'    gobjShared.ShowVBError Err.Number, Err.Description, "FillRecordsetFromPIDNode Function", "DMBE", vbExclamation
    Err.Clear
    'Resume 'BM
End Function
'<SCR 22858, 22764 CC>


Private Function FillRecordsetFromZD1Node(ByVal oZD1Node As MSXML.IXMLDOMElement, ByVal lngUnitNumber As Long, _
                                        ByRef rsMMI_ID As Adodb.Recordset, ByRef rsMMI_GENERICDATA As Adodb.Recordset, _
                                        ByVal enmWhoseMessage As XMLMessageSource) As DMErrorCodes
    
    'Dim oZD1Node As MSXML.IXMLDOMElement
    Dim oItemNode As MSXML.IXMLDOMNode
    Dim oChildNode As MSXML.IXMLDOMNode
    Dim strSQL As String
    Dim strTemp As String
    Dim rsTemp As Adodb.Recordset
    Dim strService As String
    Dim oPat As MMMHISPatient.Patient
    Dim rstP As Adodb.Recordset
    Const NCID_SERVICE As Long = 1013
    
    On Error GoTo ErrHandler
        
    Set oPat = New MMMHISPatient.Patient
    Set rstP = New Adodb.Recordset
    
    ' Begin:SCR #27452;   Developer: MAG 09/06/2002 10:31 AM
    If oZD1Node Is Nothing Then
        Exit Function
    End If
    ' End: SCR #27452;
    
    If oDas Is Nothing Then
        Set oDas = oDatabase.CHCSII_DAS(Auto) 'New GEMS_DAS
    End If
    
    strService = "U"
    
    For Each oItemNode In oZD1Node.childNodes
        Select Case UCase(oItemNode.nodeName)
            Case "ZD1.1"
               'Branch of Service
              Set oChildNode = oItemNode.selectSingleNode("XON.1")
              
              If Not oChildNode Is Nothing Then
                strTemp = oChildNode.Text
                
                ' Begin:SCR #31184;   Developer: MAG 01/07/2003 11:20 AM
                ' Note: SAMS is giving us one character representation
                strSQL = "SELECT * FROM TMIP_SAMS_SERVICE_RANK_MAPPING WHERE "
                If Len(strTemp) = 1 Then
                    strSQL = strSQL & "CHCSII_REPRESENTATION = " & oDatabase.SQLQuote(UCase(oChildNode.Text)) & " AND "
                Else
                    strSQL = strSQL & "TMIP_SAMS_DISPLAY = " & oDatabase.SQLQuote(UCase(oChildNode.Text)) & " AND "
                End If
                ' End: SCR #31184;
                strSQL = strSQL & "CHCSII_DOMAIN_NCID = " & NCID_SERVICE
                                    
                '<SCR 23295 CC>
                If oDas Is Nothing Then Set oDas = oDatabase.CHCSII_DAS(Auto) 'New GEMS_DAS
                Set rsTemp = oDas.ExecuteSQL(strSQL)
                '</SCR 23295 CC>
                If Not rsTemp Is Nothing Then
                  
                  If Not rsTemp.EOF Then
                    
                    strService = rsTemp.Fields("CHCSII_REPRESENTATION").Value
                        
'>> SCR #37693;   Developer: Brian Mowbray 07/31/2003 07:58 PM
'This is now being handled by the BMIST_ZD1.xsl file
'                    '<< Begin:SCR #30825;   Developer: Brian Mowbray 12/16/2002 12:54 PM
'                    If enmWhoseMessage = BMIST_Message Then
'                      Set rstP = oPat.PatientOps.GetPatCatSelectStrings(NCID_PATIENT_CATEGORY)
'
'                      Select Case UCase$(Left$(strService, 1))
'
'                      Case "A"
'                       lPatNCID = 99990013
'                      Case "B"
'                       lPatNCID = 99990049
'                      Case "C"
'                       lPatNCID = 99990062
'                      Case "F"
'                       lPatNCID = 119069
'                      Case "K"
'                       lPatNCID = 99990127
'                      Case "M"
'                       lPatNCID = 99990262
'                      Case "N"
'                       lPatNCID = 99990287
'                      Case "P"
'                       lPatNCID = 99990314
'                      Case "R"
'                       lPatNCID = 99990328
'                      End Select
'
'                      FilterRS "PATCAT_NCID", lPatNCID, rstP
'
'                      sPatCat = rstP.Fields("PATCAT_TEXT").Value & vbNullString
'
'                      If Len(Trim$(sPatCat)) > 0 Then
'
'                        rsMMI_GENERICDATA.AddNew
'                        rsMMI_GENERICDATA.Fields("UNIT_NUMBER").Value = lngUnitNumber
'                        rsMMI_GENERICDATA.Fields("TYPE_NCID").Value = NCID_PATIENT_CATEGORY
'                        rsMMI_GENERICDATA.Fields("VALUE_TEXT").Value = rstP.Fields("PATCAT_TEXT").Value
'                        rsMMI_GENERICDATA.Fields("VALUE_NCID").Value = lPatNCID
'                        rsMMI_GENERICDATA.Update
'                        rstP.Filter = ""
'
'                      End If 'sPatCat test
'                    End If  'enmWhoseMessage = BMIST_Message
                  '>> End: SCR #30825;
                  End If 'rsTemp.EOF
                End If 'rsTemp Is Nothing
              End If  'oChildNode Is Nothing
                
                      
            Case "ZD1.2"
                strTemp = ""
                Set oChildNode = oItemNode.selectSingleNode("XON.1")
                If Not oChildNode Is Nothing Then
                    '<< Begin:SCR #30703;   Developer: Brian Mowbray 12/11/2002 10:16 AM
                    rsMMI_ID.Fields("DEPLOYUNIT").Value = NullToNothing(oChildNode.Text)
                    '>> End: SCR #30703;
                
                ' Begin:SCR #31184;   Developer: MAG 01/07/2003 11:21 AM
                ' Note : Adding UIC code to MMI_GenericData
                    strTemp = NullToNothing(oChildNode.Text)
                End If
               
                Set oChildNode = oItemNode.selectSingleNode("XON.3")
                If Not oChildNode Is Nothing Then
                    If NullToNothing(oChildNode.Text) <> "" Then
                        strTemp = strTemp & "     (" & oChildNode.Text & ")"
                    End If
                End If
                If strTemp <> "" Then
                    rsMMI_GENERICDATA.AddNew
                    rsMMI_GENERICDATA.Fields("UNIT_NUMBER").Value = lngUnitNumber
                    rsMMI_GENERICDATA.Fields("TYPE_NCID").Value = NCID_UNIT_IDENTIFICATION_CODE
                    rsMMI_GENERICDATA.Fields("VALUE_TEXT").Value = strTemp
                    rsMMI_GENERICDATA.Update
                End If
            ' End: SCR #31184;
            Case "ZD1.3"
            'Rank
            'NCID_RANK = 1015
            
            strTemp = ""
            If enmWhoseMessage = CHCSII_Message Then
'               << Begin:SCR #27270;   Developer: Brian Mowbray 09/09/2002 08:31 PM
'               Rank is already being added by the ZPG segment, this causes a duplicate record
'
'                    Set oChildNode = oItemNode.selectSingleNode("CE.4")
'                    If Not oChildNode Is Nothing Then
'                        strTemp = NullToNothing(oChildNode.Text)
'                    End If
'                    If strTemp <> "" Then
'                        rsMMI_GENERICDATA.AddNew
'                            rsMMI_GENERICDATA.Fields("UNIT_NUMBER").Value = lngUnitNumber
'                            rsMMI_GENERICDATA.Fields("TYPE_NCID").Value = NCID_RANK
'                            rsMMI_GENERICDATA.Fields("VALUE_NCID").Value = strTemp
'                        rsMMI_GENERICDATA.Update
'                    End If
'                 >> End: SCR #27270;
            
            Else 'This is from an outside source
                Set oChildNode = oItemNode.selectSingleNode("CE.2")
                If Not oChildNode Is Nothing Then
                    strSQL = "SELECT * FROM TMIP_SAMS_SERVICE_RANK_MAPPING "
                    strSQL = strSQL & "WHERE CHCSII_DOMAIN_NCID = " & NCID_RANK & " "
                    strSQL = strSQL & "AND TMIP_SAMS_DISPLAY = " & oDatabase.SQLQuote(UCase(oChildNode.Text)) & ""
                    
                    Set rsTemp = oDas.ExecuteSQL(strSQL)
                    If Not rsTemp.EOF Then
                        
                        Dim rsRank As Adodb.Recordset
                        Dim intI As Integer
                        
                        '<< Begin:SCR #27270,31184;   Developer: Brian Mowbray 09/09/2002 08:45 PM
                        'Added the Value_text and Full_rank so this info will display rank in demographics when SAMS is the source
                        For intI = 0 To rsTemp.RecordCount - 1
                            strSQL = "SELECT NCID, FULL_RANK, PAYGRADE FROM RANK_ABBREV "
                            strSQL = strSQL & "WHERE NCID = " & rsTemp.Fields("CHCSII_NCID").Value & " "
'                            strSQL = strSQL & "AND BRANCH = '" & strService & "'"
                        
                            Set rsRank = oDas.ExecuteSQL(strSQL)
                            If Not rsRank.EOF Then
                                rsMMI_GENERICDATA.AddNew
                                    rsMMI_GENERICDATA.Fields("UNIT_NUMBER").Value = lngUnitNumber
                                    rsMMI_GENERICDATA.Fields("TYPE_NCID").Value = NCID_RANK
                                    
                                    strTemp = rsRank.Fields("FULL_RANK").Value & " " & rsRank.Fields("PAYGRADE").Value & "-" & strService
                                    rsMMI_GENERICDATA.Fields("VALUE_TEXT").Value = strTemp & vbNullString
                                    rsMMI_GENERICDATA.Fields("VALUE_NCID").Value = rsRank.Fields("NCID").Value
                                rsMMI_GENERICDATA.Update
                                Exit For
                            End If
                        Next intI
                    '>> End: SCR #27270;
                    End If 'rsTemp.EOF
                End If 'oChildNode Is Nothing
            End If 'enmWhoseMessage = CHCSII_Message
                                
            Case "ZD1.4"
                '<< Begin:SCR #30703;   Developer: Brian Mowbray 12/11/2002 10:16 AM
                'Pay Grade
                  Set oChildNode = oItemNode.selectSingleNode("CE.2")
                  If Not oChildNode Is Nothing Then
  
                    strSQL = "SELECT NCID, FULL_RANK, PAYGRADE FROM RANK_ABBREV "
                    strSQL = strSQL & "WHERE PAYGRADE = '" & UCase(oChildNode.Text) & "' "
                    strSQL = strSQL & "AND BRANCH = '" & strService & "'"
                
                    Set rsRank = oDas.ExecuteSQL(strSQL)
                    ' Begin:SCR #31184;   Developer: MAG 01/07/2003 11:29 AM
                    ' Note: SAMS PayGrade comes in as E04 instead of E4
                    If rsRank.EOF Then
                        strTemp = oChildNode.Text
                        If Len(strTemp) >= 2 Then
                            If Mid(strTemp, 2, 1) = "0" Then
                                strTemp = Left(strTemp, 1) & Right(strTemp, Len(strTemp) - 2)
                                strSQL = "SELECT NCID, FULL_RANK, PAYGRADE FROM RANK_ABBREV "
                                strSQL = strSQL & "WHERE PAYGRADE = '" & UCase(strTemp) & "' "
                                strSQL = strSQL & "AND BRANCH = '" & strService & "'"
                            End If
                        End If
                        Set rsRank = oDas.ExecuteSQL(strSQL)
                    End If
                    ' End: SCR #31184;
                    
                    If Not rsRank.EOF Then
                        rsMMI_GENERICDATA.AddNew
                            rsMMI_GENERICDATA.Fields("UNIT_NUMBER").Value = lngUnitNumber
                            rsMMI_GENERICDATA.Fields("TYPE_NCID").Value = NCID_RANK
                            
                            strTemp = rsRank.Fields("FULL_RANK").Value '& " " & rsRank.Fields("PAYGRADE").Value & "-" & strService
                            rsMMI_GENERICDATA.Fields("VALUE_TEXT").Value = strTemp & vbNullString
                            rsMMI_GENERICDATA.Fields("VALUE_NCID").Value = rsRank.Fields("NCID").Value
                        rsMMI_GENERICDATA.Update
                        
                    End If
                  End If
            '>> End: SCR #30703;
            Case "ZD1.6"
                'Flight Status
                rsMMI_GENERICDATA.AddNew
                rsMMI_GENERICDATA.Fields("UNIT_NUMBER").Value = lngUnitNumber
                rsMMI_GENERICDATA.Fields("TYPE_NCID").Value = NCID_FLYING_STATUS
                Set oChildNode = oItemNode.selectSingleNode("CE.2")
                If Not oChildNode Is Nothing Then
                    If Not oChildNode.Text = "" Then
                        If UCase(Left(oChildNode.Text, 1)) = "Y" Then
                            rsMMI_GENERICDATA.Fields("VALUE_TEXT").Value = "Y"
                        Else: rsMMI_GENERICDATA.Fields("VALUE_TEXT").Value = "N" '<SCR 25148 CC>
                        End If
                    Else: rsMMI_GENERICDATA.Fields("VALUE_TEXT").Value = "N"
                    End If
                Else: rsMMI_GENERICDATA.Fields("VALUE_TEXT").Value = "N"
                End If
                rsMMI_GENERICDATA.Update
            Case "ZD1.7"
                'Patient Category
                'NCID_PATIENT_CATEGORY
                
                ' Begin:SCR #28091, #30932;   Developer: MAG 12/17/2002 03:45 PM
                Set oChildNode = oItemNode.selectSingleNode("CE.2")
                If Not oChildNode Is Nothing Then
                    If Not oChildNode.Text = "" Then
                        rsMMI_GENERICDATA.AddNew
                        rsMMI_GENERICDATA.Fields("UNIT_NUMBER").Value = lngUnitNumber
                        rsMMI_GENERICDATA.Fields("TYPE_NCID").Value = NCID_PATIENT_CATEGORY
                        rsMMI_GENERICDATA.Fields("VALUE_TEXT").Value = oChildNode.Text
                        rsMMI_GENERICDATA.Update
                    End If
                End If
                ' End: SCR #28091, #30932;
            
            '<< Begin:SCR #30900, #30932(MAG);   Developer: Brian Mowbray 12/16/2002 11:54 PM
            'This was added for BMIST import to handle the UIC
            Case "ZD1.8"
                'UIC
                Set oChildNode = oItemNode.selectSingleNode("CE.2")
                If Not oChildNode Is Nothing Then
                    If Not oChildNode.Text = "" Then
                        rsMMI_GENERICDATA.AddNew
                        rsMMI_GENERICDATA.Fields("UNIT_NUMBER").Value = lngUnitNumber
                        rsMMI_GENERICDATA.Fields("TYPE_NCID").Value = NCID_UNIT_IDENTIFICATION_CODE
                        rsMMI_GENERICDATA.Fields("VALUE_TEXT").Value = oChildNode.Text
                        rsMMI_GENERICDATA.Update
                    End If
                End If
        '>> End: SCR #30900, #30932;
        End Select
    Next oItemNode
    rsMMI_ID.Update
    
    FillRecordsetFromZD1Node = DMErrorCodes.ERROR_SUCCESS
    On Error Resume Next
    Set rstP = Nothing

Exit Function

ErrHandler:
    FillRecordsetFromZD1Node = ERROR_EIC_XML_TO_RECORDSET
'    If gobjShared Is Nothing Then
'        Set gobjShared = New CWShared
'    End If
'    gobjShared.ShowVBError Err.Number, Err.Description, "FillRecordsetFromZD1Node Function", "DMBE", vbExclamation
    Err.Clear
End Function

Private Function GetBranchOfService(ByVal strBranch As String) As String

    Dim strCode As String

    If InStr(1, "CIV", UCase(strBranch)) > 0 Then
        'Civilian
        strCode = "K"
    ElseIf InStr(1, "ARM", UCase(strBranch)) > 0 Then
        strCode = "A"
    ElseIf InStr(1, "NOAA", UCase(strBranch)) > 0 Then
        strCode = "B"
    ElseIf InStr(1, "COAST", UCase(strBranch)) > 0 Then
        strCode = "C"
    ElseIf InStr(1, "AIR", UCase(strBranch)) > 0 Then
        strCode = "F"
    ElseIf InStr(1, "CONTRACT", UCase(strBranch)) > 0 Then
        strCode = "K"
    ElseIf InStr(1, "MARINE", UCase(strBranch)) > 0 Then
        strCode = "M"
    ElseIf InStr(1, "NAV", UCase(strBranch)) > 0 Then
        strCode = "N"
    ElseIf InStr(1, "USPHS", UCase(strBranch)) > 0 Then
        strCode = "P"
    ElseIf InStr(1, "NATO", UCase(strBranch)) > 0 Then
        strCode = "R"
    Else
        
    End If
    
End Function

Private Function FillRecordsetFromZPGNode(ByVal oZPGNode As MSXML.IXMLDOMElement, ByVal lngUnitNumber As Long, _
                                        ByRef rsMMI_GENERICDATA As Adodb.Recordset) As DMErrorCodes
    
    'Dim oZPGNode As MSXML.IXMLDOMElement
    Dim oItemNode As MSXML.IXMLDOMNode
    
    Dim oNodeChild2 As MSXML.IXMLDOMNode
    
    On Error GoTo ErrHandler
    
    If oZPGNode Is Nothing Then
        Exit Function
    End If
    
    'Set Node to "ZPG" Node
    'For Each oZPGNode In oZPGNodeList
        rsMMI_GENERICDATA.AddNew
        rsMMI_GENERICDATA("UNIT_NUMBER") = lngUnitNumber
        
        For Each oItemNode In oZPGNode.childNodes
            Select Case UCase(oItemNode.nodeName)
                Case "ZPG.1"
                   'rsMMI_GENERICDATA("") = oItemNode.Text
                Case "ZPG.2"
                    rsMMI_GENERICDATA("TYPE_NCID") = XMLToOracleNumber(oItemNode.Text)
                Case "ZPG.3"
                    Set oNodeChild2 = oItemNode.selectSingleNode("CE.1")
                    rsMMI_GENERICDATA("VALUE_NCID") = XMLToOracleNumber(oNodeChild2.Text)
                    Set oNodeChild2 = oItemNode.selectSingleNode("CE.2")
                    rsMMI_GENERICDATA("VALUE_TEXT") = oNodeChild2.Text
                Case "ZPG.4"
                    rsMMI_GENERICDATA("ORDINAL") = XMLToOracleNumber(oItemNode.Text)
                Case "ZPG.5"
                    rsMMI_GENERICDATA("PARENT") = XMLToOracleNumber(oItemNode.Text)
                Case "ZPG.6"
                    rsMMI_GENERICDATA("VALUE_RSID") = XMLToOracleNumber(oItemNode.Text)
                Case "ZPG.7"
                    rsMMI_GENERICDATA("ERROR_TEXT") = oItemNode.Text
            End Select
        Next oItemNode
        rsMMI_GENERICDATA.Update
    'Next oZPGNode
    
    FillRecordsetFromZPGNode = DMErrorCodes.ERROR_SUCCESS

Exit Function

ErrHandler:
    FillRecordsetFromZPGNode = DMErrorCodes.ERROR_EIC_XML_TO_RECORDSET
    If gobjShared Is Nothing Then
        Set gobjShared = New CWShared
    End If
    gobjShared.ShowVBError Err.Number, Err.Description, "FillRecordsetFromZPGNode Function", "DMBE", vbExclamation
    Err.Clear
End Function

Private Function FillRecordsetFromZPTNode(ByVal oZPTNode As MSXML.IXMLDOMNode, ByVal lngUnitNumber As Long, _
                                        ByRef rsMMI_INST_NAME As Adodb.Recordset) As DMErrorCodes

    'Dim oZPTNode As MSXML.IXMLDOMNode
    Dim oItemNode As MSXML.IXMLDOMNode
    Dim oChildNode As MSXML.IXMLDOMNode
    
    On Error GoTo ErrHandler
    
    If oZPTNode Is Nothing Then
        Exit Function
    End If
    
    'Set Node to "ZPT" Node
    'For Each oZPTNode In oZPTNodeList
        rsMMI_INST_NAME.AddNew
        rsMMI_INST_NAME("UNIT_NUMBER") = lngUnitNumber
        
        For Each oItemNode In oZPTNode.childNodes
            Select Case UCase(oItemNode.nodeName)
                Case "ZPT.1"
    '                    RsMMI_Inst_Name("") = oItemNode.Text
                Case "ZPT.2"
                    rsMMI_INST_NAME("UNIT_NUMBER") = XMLToOracleNumber(oItemNode.Text)
                Case "ZPT.3"
                    rsMMI_INST_NAME("NAME_INDEX") = XMLToOracleNumber(oItemNode.Text)
                Case "ZPT.4"
                    rsMMI_INST_NAME("NAME") = oItemNode.Text
                Case "ZPT.5"
                    rsMMI_INST_NAME("CREATE_AUDIT_NUM") = XMLToOracleNumber(oItemNode.Text)
                Case "ZPT.6"
                    rsMMI_INST_NAME("ROW_UPDATE_AUDIT_NUM") = XMLToOracleNumber(oItemNode.Text)
                Case "ZPT.7"
                    'Need to change Description to ID
                    'rsMMI_INST_NAME("NAME_TYPE_CID") = XMLToOracleNumber(oItemNode.Text)
                    Set oChildNode = oItemNode.selectSingleNode("CE.5")
                    If Not oChildNode Is Nothing Then
                        rsMMI_INST_NAME.Fields("NAME_TYPE_CID").Value = XMLToOracleNumber(oChildNode.Text)
                    End If
                Case "ZPT.8"
                    If UCase(oItemNode.Text) <> "CREATED BY" Then
                        rsMMI_INST_NAME("CREATEDBY") = XMLToOracleNumber(oItemNode.Text)
                    End If
                Case "ZPT.9"
                    If UCase(oItemNode.Text) <> "CREATED ON" Then
                        rsMMI_INST_NAME("CREATEDON") = XMLToOracleDate(oItemNode.Text)
                    End If
                Case "ZPT.10"
                    If UCase(oItemNode.Text) <> "UPDATED BY" Then
                        rsMMI_INST_NAME("UPDATEDBY") = XMLToOracleNumber(oItemNode.Text)
                    End If
                Case "ZPT.11"
                    If UCase(oItemNode.Text) <> "UPDATED ON" Then
                        rsMMI_INST_NAME("UPDATEDON") = XMLToOracleDate(oItemNode.Text)
                    End If
            End Select
        Next oItemNode
        rsMMI_INST_NAME.Update
    'Next oZPTNode
    
    FillRecordsetFromZPTNode = DMErrorCodes.ERROR_SUCCESS
                          
Exit Function

ErrHandler:
    FillRecordsetFromZPTNode = DMErrorCodes.ERROR_EIC_XML_TO_RECORDSET
    If gobjShared Is Nothing Then
        Set gobjShared = New CWShared
    End If
    gobjShared.ShowVBError Err.Number, Err.Description, "FillRecordsetFromZPTNode Function", "DMBE", vbExclamation
    Err.Clear
End Function

Private Function XMLToOracleNumber(ByVal sValue As String) As Variant
    ' Begin:SCR #33071;   Developer: Mag 05/13/2003 02:24 PM
    On Error Resume Next
    
    XMLToOracleNumber = Null
    ' End: SCR #33071;
    
    If IsNumeric(sValue) Then
        XMLToOracleNumber = sValue
'    Else
'       XMLToOracleNumber = Null
    End If
    
End Function

Private Function XMLToOracleDate(ByVal sValue As String) As Variant

    Dim dDate As Date

    ' Begin:SCR #33071;   Developer: Mag 05/13/2003 02:24 PM
    On Error Resume Next
    
    XMLToOracleDate = Null
    ' End: SCR #33071;
    
    
    If IsDate(sValue) Then
      XMLToOracleDate = CDate(sValue)
      Exit Function
    End If
        
    sValue = Replace(sValue, "/", "")
    sValue = Replace(sValue, "-", "")
    sValue = Replace(sValue, " ", "")
    sValue = Replace(sValue, ":", "")
    
    If IsNumeric(sValue) Then
    
        sValue = Left$(sValue & "00000000000000", 14)
        
        sValue = Mid$(sValue, 5, 2) & "/" & Mid$(sValue, 7, 2) & "/" & Mid$(sValue, 1, 4) & " " & Mid$(sValue, 9, 2) & ":" & Mid$(sValue, 11, 2) & ":" & Mid$(sValue, 13, 2)
        dDate = CDate(sValue)
        XMLToOracleDate = dDate
'    Else
'        XMLToOracleDate = Null
    End If
    
'    If IsDate(sValue) Then
'        'Need to add Date format
'        XMLToOracleDate = Format(sValue, DATE_FORMAT_ORA)
'        If Not IsDate(XMLToOracleDate) Then
'            XMLToOracleDate = DateSerial(Year(sValue), Month(sValue), Day(sValue))
'            'XMLToOracleDate = Null
'        End If
'    Else
'        sValue = Replace(sValue, "/", "")
'        sValue = Replace(sValue, "-", "")
'        sValue = Left$(sValue & "00000000000000", 14)
'        XMLToOracleDate = DateSerial(CInt(Left(sValue, 4)), Mid(sValue, 5, 2), Mid(sValue, 7, 2))
'
'    End If
    
End Function


Private Function NullToNothing(ByVal varPassIn As Variant) As String

    If IsNull(varPassIn) Then
        NullToNothing = ""
    Else
        NullToNothing = varPassIn
    End If
    
End Function

'<< SCR #29291;   Developer: Brian Mowbray 11/14/2002 11:40 AM
Public Function CreateEncounterFromSAMSEIC(oDom As MSXML.DOMDocument, Optional enmWhoseMessage As XMLMessageSource = SAMS_Message) As Long

'___ we expect to receive a DOMDocument containing only the A04 message
    On Error GoTo ErrHandler
'___ assume the current patient is the patient for this encounter.
    Dim sDataID     As String
    Dim oNode       As MSXML.IXMLDOMNode
    Dim oNodes      As MSXML.IXMLDOMNodeList
    Dim sSigLine1 As String
    Dim sSigLine2 As String
    Dim sSigLine3 As String
    Dim sOBXPairs As String
    Dim dEncDate  As Date
    Dim sDT As String
    Dim objSQL As ICHCSII_SQl.ISqlOpsEx
    '<< Begin:SCR #29291;   Developer: Brian Mowbray 11/14/2002 11:40 AM
    Dim sProviderName As String
    Dim sProviderTitle As String
    Dim sWhoAmITitle As String
    
    Dim sProviderSSN As String
    Dim sProviderLName As String
    Dim sProviderFName As String
    Dim sProviderNCID As String
    Dim lChar As Long
    Dim encounterSignedDate As Date
    Dim cosignerSSN As String
    Dim cosignerLastName As String
    Dim cosignerFirstName As String
    Dim cosignerNCID As String
    Dim externalEncounterKey As String
    Dim externalPreviousEncounterKey As String
    Dim externalDeviceId As String
    Dim reAssessedEncounter As Boolean
   
    '>> End: SCR #29291;
    CreateEncounterFromSAMSEIC = DMErrorCodes.ERROR_EIC_XML_TO_RECORDSET
    
    menMessageSource = enmWhoseMessage
    
    Set objSQL = oDatabase.CHCSII_SQLOPS(Auto)
        
'___ get the enc date
    Set oNode = oDom.selectSingleNode("//PV1.44")
    If Not oNode Is Nothing Then
        dEncDate = ReformatDate(oNode.Text)
    End If
    
    If enmWhoseMessage = SAMS_Message Then
    
      sWhoAmITitle = "SAMS Encounter Note"
    
  '___ collect the obx data elements as sets of:
  '    Obs description: value
      Set oNodes = oDom.selectNodes("//PV1.7.LST/PV1.7")
      If Not oNodes Is Nothing Then
          sSigLine1 = oNodes(0).childNodes(1).Text & ", " & oNodes(0).childNodes(2).Text
          sSigLine2 = oNodes(0).childNodes(5).Text & ", " & oNodes(0).childNodes(6).Text
          sSigLine3 = oNodes(0).childNodes(7).Text & vbCrLf & "=========================" & vbCrLf
          sOBXPairs = sSigLine1 & vbCrLf & sSigLine2 & vbCrLf & sSigLine3 & vbCrLf
      End If
      
      sProviderName = oNodes(0).childNodes(1).Text & vbNullString
      sProviderTitle = oNodes(0).childNodes(6).Text & vbNullString
      
      Set oNodes = oDom.selectNodes("//ADT_A01.LST.3/OBX")
      
      If Not oNodes Is Nothing Then
          For Each oNode In oNodes
              If Not oNode.childNodes(1).Text = "" Then
                  sOBXPairs = sOBXPairs & oNode.childNodes(1).Text & ": " & oNode.childNodes(3).Text & vbCrLf
              End If
          Next oNode
      End If
      
      '<< Begin:SCR #29043;   Developer: Brian Mowbray 11/13/2002 07:58 PM
      Set oNodes = oDom.selectNodes("//ADT_A01.LST.5/DG1/DG1.3")
      
      sOBXPairs = sOBXPairs & vbCrLf & "Diagnosis" & vbCrLf & "==============================" & vbCrLf
      
      If Not oNodes Is Nothing Then
          For Each oNode In oNodes
              If Not oNode.childNodes(1).Text = "" Then
                  sOBXPairs = sOBXPairs & oNode.childNodes(1).Text & vbTab & "ICD9:" & oNode.childNodes(0).Text & vbCrLf
              End If
          Next oNode
      End If
      '>> End: SCR #29043;
      
      sDataID = oDas.GetNextID(esMAINSEQ)

      objSQL.Execute "Select * from clinnote where dataid = 0"
      'objSQL.Execute "clinnote"
      objSQL.AddNew
      objSQL("dataid") = sDataID
      objSQL("APPLICATIONNCID") = 77616
      objSQL("NOTETYPENCID") = 123
      objSQL("NOTETYPENAME") = "SAMS Enc Data"
      objSQL("NOTEDATE") = CDate(dEncDate)
      objSQL("POCFACILITY") = gobjLogon.FacilityNCID
      objSQL("ENCFACILITY") = gobjLogon.FacilityNCID
      objSQL("CLINICIAN") = gobjLogon.UserName
      objSQL("ENCOUNTER") = 0
      objSQL("STATUS") = "Final"
      objSQL("Unit_Number") = gobjPatient.unitNumber
      objSQL.AppendAsChunk "HTMLTEXT", oDatabase.SQLQuote(sOBXPairs)
      objSQL.Update
  
  ElseIf enmWhoseMessage = BMIST_Message Then
  
        '-- Get encounter signed date
        Set oNode = oDom.selectSingleNode("//ZPV.16")
        If Not oNode Is Nothing Then
            encounterSignedDate = ReformatDate(oNode.Text)
        End If
        
        '-- Get device id
        Set oNode = oDom.selectSingleNode("//ZPV.2")
        If Not oNode Is Nothing Then
            externalDeviceId = oNode.Text
        End If
        
        '-- Get encounter key
        Set oNode = oDom.selectSingleNode("//ZPV.3")
        If Not oNode Is Nothing Then
            externalEncounterKey = oNode.Text
        End If
        
        '-- Get previous encounter key
        Set oNode = oDom.selectSingleNode("//ZPV.4")
        If Not oNode Is Nothing Then
            externalPreviousEncounterKey = oNode.Text
        End If
        
        If Len(externalPreviousEncounterKey) > 0 Then
            reAssessedEncounter = True
        Else
            reAssessedEncounter = False
        End If
          
        If reAssessedEncounter Then
            sWhoAmITitle = "AHLTA-Mobile Appended/Amended/Reassessed Encounter Note"
        Else
            sWhoAmITitle = "AHLTA-Mobile Encounter Note"
        End If
        
        '>> Begin:SCR #47334;   Developer: Brian Mowbray 02/04/2004 09:39 AM
        Set oNodes = oDom.selectNodes("//ZPV.52")
        If oNodes.Length > 0 Then
          If Not oNodes Is Nothing Then
            sProviderSSN = oNodes(0).childNodes(0).Text
            sProviderLName = oNodes(0).childNodes(1).Text
            sProviderFName = oNodes(0).childNodes(2).Text
            cosignerSSN = oNodes(0).childNodes(3).Text
            cosignerLastName = oNodes(0).childNodes(4).Text
            cosignerFirstName = oNodes(0).childNodes(5).Text
          End If
        End If
              
        If Len(sProviderSSN) > 0 Then
          'SF SCR 62192, 65523 2-21-05 This needs to share the same sequence as the patient because the of the assumptions in the concept control.
          sProviderNCID = objSQL.GetNextID(exUNIT_NUMBER)
          
          Call objSQL.ValidateProvider2(sProviderNCID, sProviderSSN, sProviderFName, "", sProviderLName, True)
        Else
          sProviderNCID = gobjLogon.UserNCID
        End If
        '<< End: SCR #47334;
        
        If Len(cosignerSSN) > 0 Then
           cosignerNCID = objSQL.GetNextID(exUNIT_NUMBER)
           Call objSQL.ValidateProvider2(cosignerNCID, cosignerSSN, cosignerFirstName, "", cosignerLastName, True)
        End If
        
        
        Set oNodes = oDom.selectNodes("//PV1.7.LST/PV1.7")
        If oNodes.Length > 0 Then
            If Not oNodes Is Nothing Then
                sSigLine1 = oNodes(0).childNodes(1).Text
                sSigLine2 = oNodes(0).childNodes(6).Text
                sSigLine3 = ConvertTextToBoldRTF(oNodes(0).childNodes(7).Text) & vbCrLf & "===================================" & vbCrLf
                sOBXPairs = sSigLine1 & vbCrLf & sSigLine2 & vbCrLf & sSigLine3 & vbCrLf
                
                sProviderName = oNodes(0).childNodes(1).Text & vbNullString
                sProviderTitle = oNodes(0).childNodes(6).Text & vbNullString
            End If
          End If
        
        Set oNodes = oDom.selectNodes("//ADT_A01.LST.3/OBX")
        
        If Not oNodes Is Nothing Then
            For Each oNode In oNodes
                If Not oNode.childNodes(2).Text = "" Then
                    If oNode.childNodes(2).Text = "|" Then
                       lChar = InStr(oNode.childNodes(4).Text, ":")
                       If lChar > 0 Then
                          sOBXPairs = sOBXPairs & ConvertTextToBoldRTF(Left(oNode.childNodes(4).Text, lChar)) & " " & ConvertTextToRTF(Mid(oNode.childNodes(4).Text, lChar + 1)) & vbCrLf
                       Else
                          sOBXPairs = sOBXPairs & ConvertTextToRTF(oNode.childNodes(4).Text) & vbCrLf
                       End If
                    Else
                       sOBXPairs = sOBXPairs & ConvertTextToBoldRTF(oNode.childNodes(2).Text & ":") & " " & ConvertTextToRTF(oNode.childNodes(4).Text) & vbCrLf
                    End If
                End If
            Next oNode
        End If
        
        If Len(Trim$(sOBXPairs)) > 0 Then
            CreateEncounterFromSAMSEIC = CreateExternalEncounter(dEncDate, sProviderLName & ", " & sProviderFName, sProviderTitle, sOBXPairs, sWhoAmITitle, oDom, sProviderNCID, _
                                        cosignerNCID, encounterSignedDate, externalDeviceId, externalEncounterKey, externalPreviousEncounterKey, reAssessedEncounter, cosignerLastName & ", " & cosignerFirstName)
          Exit Function
        Else
          CreateEncounterFromSAMSEIC = DMErrorCodes.ERROR_NO_BMIST_ENCOUNTER_DATA_FOUND
          Exit Function
        End If
    End If
    
    CreateEncounterFromSAMSEIC = DMErrorCodes.ERROR_SUCCESS

Exit Function
    
ErrHandler:
    CreateEncounterFromSAMSEIC = DMErrorCodes.ERROR_EIC_XML_TO_RECORDSET

Exit Function
Resume
End Function

'<< SCR #29291;   Developer: Brian Mowbray 11/14/2002 12:03 PM
Private Function CreateExternalEncounter(ByVal encounterDate As Date, _
                                         ByVal providerName As String, _
                                         ByVal title As String, _
                                         ByVal note As String, _
                                         ByVal whoseMessageTitle As String, _
                                         ByRef domDoc As MSXML.DOMDocument, _
                                         ByVal providerNCID As String, _
                                         ByVal cosignerNCID As String, _
                                         ByVal encounterSignedDate As Date, _
                                         ByVal externalDeviceId As String, _
                                         ByVal externalEncounterKey As String, _
                                         ByVal externalPreviousEncounterKey As String, _
                                         ByVal reassessed As Boolean, _
                                         ByVal cosignerName As String) As Long
                                         
    Dim apptID             As String
    Dim encounterID        As String
    Dim sql                As String
    Dim objSQL             As ICHCSII_SQl.ISqlOpsEx
    Dim node               As MSXML.IXMLDOMNode
    Dim itemNode           As MSXML.IXMLDOMNode
    Dim errorCode          As DMErrorCodes
    Dim isBMISTEnc         As Boolean 'SCR 128313
    Dim workStatuses       As workStatuses
    Dim signatureRtf       As String
    Const rtfNewLine       As String = " \par "
    
    Dim multisql As Collection
    
    On Error GoTo ErrHnd
        
    Set objSQL = oDatabase.CHCSII_SQLOPS(Auto)
    
    apptID = objSQL.GetNextID(esAPPT_ID)
    encounterID = objSQL.GetNextID(esENCOUNTER_ID)
    errorCode = ERROR_SUCCESS
    
    'SCR 128313
    isBMISTEnc = InStr(whoseMessageTitle, "AHLTA-Mobile") 'BMIST
    
    Set multisql = New Collection
    
    If EncounterAlreadyAdded(encounterSignedDate, gobjPatient.unitNumber) Then
        CreateExternalEncounter = ERROR_ENCOUNTER_ALREADY_EXISTS
        GoTo Cleanup
    End If
    
    If Val(cosignerNCID) > 0 Then
        signatureRtf = rtfNewLine & ConvertTextToBoldRTF("Co-Signed By") & " " & ConvertTextToRTF(UCase(cosignerName) & " @ " & Format(encounterSignedDate, DefaultDateTimeFormat)) & rtfNewLine & rtfNewLine
    Else
        signatureRtf = rtfNewLine & ConvertTextToBoldRTF("Signed By") & " " & ConvertTextToRTF(UCase(providerName) & " @ " & Format(encounterSignedDate, DefaultDateTimeFormat)) & rtfNewLine & rtfNewLine
    End If
    note = note & signatureRtf

    'For appended or amended encounter
    If reassessed And Len(externalEncounterKey) > 0 And Len(externalDeviceId) > 0 Then
        sql = "select EncounterNumber from encounters " & _
              "where EXTERNAL_DEVICE_ID = '" & externalDeviceId & "'" & _
              " AND EXTERNAL_ENCOUNTER_KEY = '" & externalPreviousEncounterKey & "'" & _
              " AND UNIT_NUMBER = " & gobjPatient.unitNumber
        objSQL.Execute sql
        
        If objSQL.RecordCount > 0 Then
            encounterID = objSQL("EncounterNumber")
            If Not EncounterRtfAlreadyAdded(encounterSignedDate, encounterID) Then
                errorCode = AddEncounterRtf(encounterID, encounterSignedDate, note, providerNCID, providerName, whoseMessageTitle, isBMISTEnc, True)
                
                ' Update the external encounter key with the latest encounter key in AHLTA-MOBILE
                ' reset tmip_status so the amendment will be forwarded to tmip
                sql = "update encounters set EXTERNAL_ENCOUNTER_KEY = '" & externalEncounterKey & "' " & _
                      ", tmip_status = 0 " & _
                      "where EncounterNumber= " & encounterID
                    
                objSQL.Execute sql
            End If
            
            GoTo Cleanup  ' alternate goto: Rejoin - to revisit content
            
        End If
    End If
        
    '>> Begin:SCR #47334;   Developer: Brian Mowbray 02/04/2004 09:39 AM
    '<< Begin:SCR 34694;   Developer: Brian Mowbray 04/16/2003 07:58 PM
    'Removed the Patient ssn and sponsor ssn because the BMIST unknown patient exceeds the table column width in appoinments
     sql = "insert into appointment ( CLINICIAN_NCID, FACILITY_NCID, APPT_ID, CLINIC_NCID," & vbCrLf _
    & "  PATIENT_UNIT_NUMBER, APPT_DATE_TIME, APPT_TYPE, CHECKIN_TIME," & vbCrLf _
    & " APPT_COMMENT, APPT_CLASSIFICATION, ENCOUNTER_STATUS, DATE_CREATED, PATIENT_NAME," & vbCrLf _
    & "  APPT_STATUS) values (" & vbCrLf _
    & providerNCID & "," & gobjLogon.FacilityNCID & "," & apptID & "," & gobjLogon.NursingDivisionNCID & vbCrLf _
    & "," & gobjPatient.unitNumber & "," & oDatabase.SQLDate(encounterDate) & ",'WI',sysdate" & vbCrLf _
    & ",'Rapid Patient Generation',0,4,sysdate," & oDatabase.SQLQuote(gobjPatient.Name) & vbCrLf _
    & ",'KEPT'" & ")"
    '>> End:SCR 34694
    multisql.Add sql
    
    
    'SCR 134253 added in WHENSIGNED
    '<< Begin:SCR 36319  Developer: Brian Mowbray 05/19/2003 04:14 PM - Changed Type value from 4 (Dental) to 1 (Outpatient)
    sql = "insert into encounters (EncounterNumber,Type,Category,ApptID,StartDTS,FacilityNCID," & vbCrLf _
    & "ClinicNCID,Unit_NUmber,Status,PRIMARYPROVIDERNCID,UPDATELOCK, BMIST_ENC, EXTERNAL_DEVICE_ID, EXTERNAL_ENCOUNTER_KEY, WHOSIGNEDNCID, WHENSIGNED) values (" & vbCrLf _
    & encounterID & ",1,0," & apptID & "," & oDatabase.SQLDate(encounterDate) & "," & gobjLogon.FacilityNCID & vbCrLf _
    & "," & gobjLogon.NursingDivisionNCID & "," & gobjPatient.unitNumber & ", 205657," & providerNCID & ",0, " & _
    Abs(isBMISTEnc) & " , '" & externalDeviceId & "', '" & externalEncounterKey & "', " & _
    providerNCID & ", " & oDatabase.SQLDate(encounterSignedDate) & ")"
    
    If Val(cosignerNCID) > 0 Then
        sql = Replace(sql, "WHOSIGNEDNCID, WHENSIGNED)", "WHENSIGNED, WHENCOSIGNED, COSIGNERNCID, WHOCOSIGNEDNCID)")
        sql = Replace(sql, providerNCID & ", " & oDatabase.SQLDate(encounterSignedDate) & ")", oDatabase.SQLDate(encounterSignedDate) & ", " & oDatabase.SQLDate(encounterSignedDate) & ", " & cosignerNCID & ", " & cosignerNCID & ")")
    End If
    
    'We need to set the UpadteLock to "0".
    '-1 is preventing the Encounter from being modified.
    'Because when we send out an encounter, we need to be able to amend the encounter.
    
    multisql.Add sql
    
    sql = "Update appointment set ENCOUNTER_NUMBER = " & encounterID _
        & " where appt_id = " & apptID
    
    multisql.Add sql
                  
    objSQL.ExecuteMulti multisql
    
    If Err.Number <> 0 Then
      errorCode = ERROR_ADDING_BMIST_ENCOUNTER_DATA
      Err.Clear
    End If
    
    DoEvents
    
    If errorCode = ERROR_SUCCESS Then
        errorCode = AddEncounterRtf(encounterID, encounterSignedDate, note, providerNCID, providerName, whoseMessageTitle, isBMISTEnc, False)
    End If
        
Rejoin:

    If errorCode = ERROR_SUCCESS Or errorCode = ERROR_ADDING_BMIST_ENCOUNTER_RTF_DATA Then 'This is okay since the rtf is not vital to getting the other data
    
      Dim rsDIAGNOSIS As Adodb.Recordset
      Dim rsLISTTOOLRECORD As Adodb.Recordset
      Dim rsLISTTOOLRECORDS As Adodb.Recordset
      Dim rsENC_DISP As Adodb.Recordset
      Dim rsVITALS As Adodb.Recordset
      Dim rsVITALSIGNSVALUE As Adodb.Recordset
      Dim lngResult As Long
      Dim lDataID As Long
      Dim lLTRDataID As Long
      Dim lPriDiagNCID As Long
      Dim rsALLERGIES As Adodb.Recordset
      Dim rsALLERGYREACTION As Adodb.Recordset
      Dim rsALLERGYCOMMENT As Adodb.Recordset
      Dim rsWORKSTATUS As Adodb.Recordset
      Dim rsMedications As Adodb.Recordset
      
      Set rsDIAGNOSIS = New Adodb.Recordset
      Set rsENC_DISP = New Adodb.Recordset
      Set rsLISTTOOLRECORDS = New Adodb.Recordset
      Set rsLISTTOOLRECORD = New Adodb.Recordset
      Set rsVITALS = New Adodb.Recordset
      Set rsVITALSIGNSVALUE = New Adodb.Recordset
      Set rsWORKSTATUS = New Adodb.Recordset
      Set rsMedications = New Adodb.Recordset
      
      Call GetFieldProps("DIAGNOSES", rsDIAGNOSIS)
      Call GetFieldProps("LISTTOOLRECORD", rsLISTTOOLRECORD)
      Call GetFieldProps("LISTTOOLRECORDS", rsLISTTOOLRECORDS)
      Call GetFieldProps("ENC_DISPOSITION", rsENC_DISP)
      Call GetFieldProps("VITALS", rsVITALS)
      Call GetFieldProps("VITALSIGNSVALUE", rsVITALSIGNSVALUE)
      Call GetFieldProps("WORK_STATUS", rsWORKSTATUS)
      Call GetFieldProps("MEDICATIONS", rsMedications)
      
      lPriDiagNCID = 0
      
      Set node = domDoc.selectSingleNode("//ADT_A04")
      If Not node Is Nothing Then
        For Each itemNode In node.childNodes
          Select Case itemNode.nodeName
          Case "ZD2"
            
              'Update Diagnosis
            
              lngResult = FillRecordsetFromZD2Node(itemNode, gobjPatient.unitNumber, rsDIAGNOSIS)
                
              lDataID = objSQL.GetNextID(exMAINSEQ)
              If lngResult = ERROR_SUCCESS Then
                 
                If Not IsNull(rsDIAGNOSIS![DIAGNOSISNCID]) Then
                  'SCR 128313
                  objSQL.Execute "Select * from DIAGNOSES where DataID = 0"
                  objSQL.AddNew
                  objSQL("DATAID") = lDataID
                  objSQL("FACILITYNCID") = rsDIAGNOSIS![FacilityNCID]
                  objSQL("DIAGNOSISNCID") = rsDIAGNOSIS![DIAGNOSISNCID]
                  objSQL("ENCOUNTERNUMBER") = encounterID
                  objSQL("CLINICIANNCID") = providerNCID
                  objSQL("UNIT_NUMBER") = gobjPatient.unitNumber
                  objSQL("SnoID") = rsDIAGNOSIS![SnoID]
                  objSQL("POINTOFCAREFACILITYNCID") = rsDIAGNOSIS![PointOfCareFacilityNCID]
                  objSQL("STATUSNCID") = rsDIAGNOSIS![STATUSNCID]
                  objSQL("NOTEDDATE") = oDatabase.SQLDate(rsDIAGNOSIS![NotedDate])
                  objSQL("ONSETDATE") = oDatabase.SQLDate(rsDIAGNOSIS![ONSETDATE])
                  objSQL("CHRONICITYNCID") = rsDIAGNOSIS![CHRONICITYNCID]
                  objSQL("SOURCENCID") = rsDIAGNOSIS![SOURCENCID]
                  objSQL("ICD9NCID") = rsDIAGNOSIS![ICD9NCID]
                  objSQL("DIAGNOSESCOMMENT") = rsDIAGNOSIS![DIAGNOSESCOMMENT]
                  objSQL("CREATEUSERNCID") = gobjLogon.UserNCID
                  objSQL("CREATETIME") = rsDIAGNOSIS![CreateTime]
                  objSQL("CREATEDBY") = rsDIAGNOSIS![CREATEDBY]
                  objSQL("CREATEDON") = rsDIAGNOSIS![CREATEDON]
                  objSQL("PRIORITY") = rsDIAGNOSIS![Priority]
                  objSQL("SENSITIVITY") = rsDIAGNOSIS![Sensitivity]
                  objSQL("NEWFOLLOWUP") = rsDIAGNOSIS![NEWFOLLOWUP]
                  objSQL("BILLINGCODENCID") = rsDIAGNOSIS![BILLINGCODENCID]
                  objSQL("TMIP_STATUS") = rsDIAGNOSIS![TMIP_Status]
                  objSQL.Update
                  
                  If Err.Number <> 0 Then
                    errorCode = ERROR_PROCESSING_BMIST_ENCOUNTER_DATA
                    Err.Clear
                  End If
                  
                  'Only take the first diagnosis ncid as the primary
                  If lPriDiagNCID = 0 Then
                    lPriDiagNCID = rsDIAGNOSIS![DIAGNOSISNCID]
                    objSQL.Execute "Select * from encounters WHERE ENCOUNTERNUMBER = " & encounterID
                    objSQL("PRIMARYDIAGNOSISNCID") = lPriDiagNCID
                    objSQL("DXSnoID") = lPriDiagNCID
                    objSQL.Update
                    
                    '>> Begin:SCR #31683; 31249;   Developer: Brian Mowbray 07/28/2003 03:38 PM
                    'Update appointment with Reason for visit
                     If lPriDiagNCID > 0 Then
                      
                      sql = "Update appointment set APPT_REASON = '" & GetDiagnosisDescription(lPriDiagNCID) & "'" _
                             & " where appt_id = " & apptID
                     
                     Else
                      'SCR 128313
                      sql = "Update appointment set APPT_REASON = 'AHLTA-Mobile Encounter, Unable to Translate AHLTA-Mobile Diagnosis into OpenAHLTA'" _
                             & " where appt_id = " & apptID
                     
                     End If
                            
                     objSQL.Execute sql
                    '<< End: SCR #31683; 31249;
                    
                  End If
                  
                  If Err.Number <> 0 Then
                    errorCode = ERROR_PROCESSING_BMIST_ENCOUNTER_DATA
                    Err.Clear
                  End If
                  
                  'Update Problems
                  '--- SCR 89535  Sherry Wang  3/8/2009
                  sql = "select DATAID from PROBLEMS " & vbNewLine
                  sql = sql & " Where UNIT_NUMBER = " & gobjPatient.unitNumber & vbNewLine
                  sql = sql & " and SnoID = " & rsDIAGNOSIS![SnoID] & vbNewLine
                
                  objSQL.Execute sql
                  
                  If objSQL.RecordCount > 0 Then
                    sql = "Delete from PROBLEMS where DATAID = " & objSQL("DATAID")
                    objSQL.Execute sql
                  End If
                  
                  lDataID = objSQL.GetNextID(exMAINSEQ)
                  'SCR 128313
                  objSQL.Execute "Select * from PROBLEMS where DataID = 0"
                  objSQL.AddNew
                  objSQL("DATAID") = lDataID
                  objSQL("FACILITYNCID") = rsDIAGNOSIS![FacilityNCID]
                  objSQL("PROBLEMNCID") = rsDIAGNOSIS![DIAGNOSISNCID]
                  objSQL("ENCOUNTERNUMBER") = encounterID
                  objSQL("CLINICIANNCID") = providerNCID
                  objSQL("UNIT_NUMBER") = gobjPatient.unitNumber
                  objSQL("SnoID") = rsDIAGNOSIS![SnoID]
                  objSQL("POINTOFCAREFACILITYNCID") = rsDIAGNOSIS![PointOfCareFacilityNCID]
                  objSQL("STATUSNCID") = rsDIAGNOSIS![STATUSNCID]
                  objSQL("NOTEDDATE") = CDate(rsDIAGNOSIS![NotedDate])
                  objSQL("ONSETDATE") = CDate(rsDIAGNOSIS![ONSETDATE])
                  objSQL("CHRONICITYNCID") = rsDIAGNOSIS![CHRONICITYNCID]
                  objSQL("SOURCENCID") = rsDIAGNOSIS![SOURCENCID]
                  objSQL("PROBLEMSCOMMENT") = ""
                  objSQL("CREATEUSERNCID") = gobjLogon.UserNCID
                  objSQL("CREATETIME") = rsDIAGNOSIS![CreateTime]
                  objSQL("CREATEDBY") = rsDIAGNOSIS![CREATEDBY]
                  objSQL("CREATEDON") = rsDIAGNOSIS![CREATEDON]
                  objSQL.Update
                  
                  '--- SCR 89765  Sherry Wang  3/9/2009
                  sql = "INSERT INTO PROBLEMSRELATIONSHIPS ( PROBLEMDATAID, RELATIONSHIPNCID, RELATEDDATAID) "
                  sql = sql & "VALUES (" & lDataID & ", " & NCID_RELATED_ENCOUNTERS & ", " & encounterID & ")"
                  objSQL.Execute sql
         
                  If Err.Number <> 0 Then
                    errorCode = ERROR_PROCESSING_BMIST_ENCOUNTER_DATA
                    Err.Clear
                  End If
                
                End If
              End If
          
        Case "ZG1"
        
            lngResult = FillRecordsetFromZG1Node(itemNode, gobjPatient.unitNumber, rsENC_DISP)
                
              If lngResult = ERROR_SUCCESS Then
                'SCR 128313
                objSQL.Execute "Select * from ENC_DISPOSITION where encounternumber=0"
                objSQL.AddNew
                objSQL("ENCOUNTERNUMBER") = encounterID
                objSQL("NCID") = rsENC_DISP![NCID]
                objSQL("FACILITYNCID") = gobjLogon.FacilityNCID
                objSQL("DISPOSITIONNCID") = rsENC_DISP![DispositionNCID]
                objSQL("DISPOSITIONTEXT") = rsENC_DISP![DispositionText]
                objSQL("WSDISP") = rsENC_DISP![WSDISP]
                objSQL("INJ_ILL_CATEGORY") = rsENC_DISP![INJ_ILL_CATEGORY]
                objSQL("INJ_ILL_CAUSE") = rsENC_DISP![INJ_ILL_CAUSE]
                objSQL.Update
              End If
        
              If Err.Number <> 0 Then
                errorCode = ERROR_PROCESSING_BMIST_ENCOUNTER_DATA
                Err.Clear
              End If
        
        
        Case "ZLT"
        
            lngResult = FillRecordsetFromZLTNode(itemNode, gobjPatient.unitNumber, rsLISTTOOLRECORD)
                
            If lngResult = ERROR_SUCCESS Then
            
              lLTRDataID = objSQL.GetNextID(exMAINSEQ)
              
              'SCR 128313
              objSQL.Execute "Select * from LISTTOOLRECORD where dataid = 0"
              objSQL.AddNew
              objSQL("DATAID") = lLTRDataID
              objSQL("FACILITYNCID") = gobjLogon.FacilityNCID
              objSQL("SnoID") = rsLISTTOOLRECORD![SnoID]
              objSQL("CLINICIANNCID") = providerNCID
              objSQL("ENCOUNTERNUMBER") = encounterID
              objSQL("CLINICNCID") = gobjLogon.NursingDivisionNCID
              objSQL("STARTDTS") = rsLISTTOOLRECORD![StartDTS]
              objSQL("ENDDTS") = rsLISTTOOLRECORD![EndDTS]
              objSQL("SOURCENCID") = rsLISTTOOLRECORD![SOURCENCID]
              objSQL.AppendAsChunk "RTF", rsLISTTOOLRECORD![RTF]
              objSQL("UPDATEFLAG") = rsLISTTOOLRECORD![UpdateFlag]
              objSQL("CREATEDBY") = rsLISTTOOLRECORD![CREATEDBY]
              objSQL("CREATEDON") = rsLISTTOOLRECORD![CREATEDON]
              objSQL("UPDATEDBY") = rsLISTTOOLRECORD![UPDATEDBY]
              objSQL("UPDATEDON") = rsLISTTOOLRECORD![UPDATEDON]
              objSQL("ORIGINALSIZE") = rsLISTTOOLRECORD![OriginalSize]
              objSQL.Update
            
              If Err.Number <> 0 Then
                errorCode = ERROR_PROCESSING_BMIST_ENCOUNTER_DATA
                Err.Clear
              End If
          
            End If
        
        
        Case "ZLR"
        
            lngResult = FillRecordsetFromZLRNode(itemNode, gobjPatient.unitNumber, rsLISTTOOLRECORDS)
                
            If lngResult = ERROR_SUCCESS Then
              
              objSQL.Execute "Select * from LISTTOOLRECORDS where dataid = 0"
              
              If rsLISTTOOLRECORDS.RecordCount > 0 Then
                  objSQL.AddNew
                  objSQL("DATAID") = lLTRDataID
                  objSQL("SnoID") = rsLISTTOOLRECORDS![SnoID]
                  objSQL("PREFIX") = rsLISTTOOLRECORDS![prefix]
                  objSQL("MODIFIER") = rsLISTTOOLRECORDS![Modifier]
                  objSQL("RESULT") = rsLISTTOOLRECORDS![result]
                  objSQL("STATUS") = rsLISTTOOLRECORDS![Status]
                  objSQL("ONSET") = rsLISTTOOLRECORDS![Onset]
                  objSQL("DURATION") = rsLISTTOOLRECORDS![Duration]
                  objSQL("VALUE") = rsLISTTOOLRECORDS![Value]
                  objSQL("UNIT") = rsLISTTOOLRECORDS![Unit]
                  objSQL("RANGESCALE") = rsLISTTOOLRECORDS![RangeScale]
                  objSQL("REFERID") = rsLISTTOOLRECORDS![Referid]
                  objSQL("RANGENORMALLOW") = rsLISTTOOLRECORDS![RangeNormalLow]
                  objSQL("RANGENORMALHIGH") = rsLISTTOOLRECORDS![RangeNormalHigh]
                  objSQL("CHARTFLAG") = rsLISTTOOLRECORDS![CHARTFLAG]
                  objSQL("QLINK") = rsLISTTOOLRECORDS![QLINK]
                  objSQL("QLINKSEQ") = rsLISTTOOLRECORDS![QLINKSEQ]
                  objSQL("NOTE") = rsLISTTOOLRECORDS![note]
                  objSQL("UPDATEFLAG") = rsLISTTOOLRECORDS![UpdateFlag]
                  objSQL("CREATEDBY") = rsLISTTOOLRECORDS![CREATEDBY]
                  objSQL("CREATEDON") = rsLISTTOOLRECORDS![CREATEDON]
                  objSQL("UPDATEDBY") = rsLISTTOOLRECORDS![UPDATEDBY]
                  objSQL("UPDATEDON") = rsLISTTOOLRECORDS![UPDATEDON]
                  objSQL.Update
                
                  If Err.Number <> 0 Then
                    errorCode = ERROR_PROCESSING_BMIST_ENCOUNTER_DATA
                    Err.Clear
                  End If
               End If
            
            End If
        Case "ZAL"
        
              'test to see if these allergies have already been added
              
              lngResult = FillRecordsetFromZALNode(itemNode, gobjPatient.unitNumber, rsALLERGIES, rsALLERGYREACTION, rsALLERGYCOMMENT)
              
              '-- SCR 189697 Webb, Check for NCID value match and Allergen text match to prevent duplicates
              sql = "select ALLERGENNCID, UNIT_NUMBER " & vbNewLine
              sql = sql & "from ALLERGIES " & vbNewLine
              sql = sql & "where (ALLERGENNCID = " & rsALLERGIES![AllergenNCID] & " or" & vbNewLine
              sql = sql & "UPPER(ALLERGEN) = UPPER('" & Replace(rsALLERGIES![Allergen], "'", "''") & "'))" & vbNewLine
              sql = sql & " AND UNIT_NUMBER = " & gobjPatient.unitNumber
              
              objSQL.Execute sql
              
              If objSQL.RecordCount = 0 Then
                lDataID = objSQL.GetNextID(exMAINSEQ)
                
                If lngResult = ERROR_SUCCESS Then
                  
                  'SCR 128313
                  objSQL.Execute "Select * from ALLERGIES where DataID = 0"
                  objSQL.AddNew
                  objSQL("DATAID") = lDataID
                  objSQL("FACILITY_NCID") = rsALLERGIES![FACILITY_NCID]
                  objSQL("UNIT_NUMBER") = rsALLERGIES![Unit_Number]
                  objSQL("ALLERGEN") = rsALLERGIES![Allergen]
                  objSQL("CLINICIANNCID") = rsALLERGIES![ClinicianNCID]
                  objSQL("ALLERGENNCID") = rsALLERGIES![AllergenNCID]
                  objSQL("ALLERGYTYPENCID") = rsALLERGIES![AllergyTypeNCID]
                  objSQL("CREATETIME") = rsALLERGIES![CreateTime]
                  objSQL("INFOSOURCENCID") = rsALLERGIES![InfoSourceNCID]
                  objSQL("NOTEDDATE") = rsALLERGIES![NotedDate]
                  objSQL("POINTOFCAREFACILITYNCID") = rsALLERGIES![PointOfCareFacilityNCID]
                  objSQL("UPDATEFLAG") = rsALLERGIES![UpdateFlag]
                  objSQL("CREATEDBY") = rsALLERGIES![CREATEDBY]
                  objSQL("CREATEDON") = rsALLERGIES![CREATEDON]
                  objSQL("UPDATEDBY") = rsALLERGIES![UPDATEDBY]
                  objSQL("UPDATEDON") = rsALLERGIES![UPDATEDON]
                  objSQL("MODIFYUSERNCID") = rsALLERGIES![MODIFYUSERNCID]
                  objSQL("CREATEUSERNCID") = rsALLERGIES![CREATEUSERNCID]
                  If isBMISTEnc Then
                    objSQL("FDB_SEQ_ID") = rsALLERGIES![FDB_SEQ_ID]
                  End If
                  objSQL.Update
                  
                  'SCR 128313
                  objSQL.Execute "Select * from ALLERGYCOMMENT where Unit_Number=0"
                  objSQL.AddNew
                  objSQL("ALLERGENNCID") = rsALLERGYCOMMENT![AllergenNCID]
                  objSQL("UNIT_NUMBER") = rsALLERGYCOMMENT![Unit_Number]
                  objSQL("ALLERGYCOMMENT") = rsALLERGYCOMMENT![ALLERGYCOMMENT]
                  objSQL("UPDATEFLAG") = rsALLERGYCOMMENT![UpdateFlag]
                  objSQL("CREATEDBY") = rsALLERGYCOMMENT![CREATEDBY]
                  objSQL("CREATEDON") = rsALLERGYCOMMENT![CREATEDON]
                  objSQL("UPDATEDBY") = rsALLERGYCOMMENT![UPDATEDBY]
                  objSQL("UPDATEDON") = rsALLERGYCOMMENT![UPDATEDON]
                  objSQL.Update
                  
                  'SCR 128313
                  objSQL.Execute "Select * from ALLERGYREACTION where Unit_Number=0"
                  objSQL.AddNew
                  objSQL("ALLERGENNCID") = rsALLERGYREACTION![AllergenNCID]
                  objSQL("UNIT_NUMBER") = rsALLERGYREACTION![Unit_Number]
                  objSQL("REACTIONNCID") = rsALLERGYREACTION![ReactionNCID]
                  objSQL("REACTION") = rsALLERGYREACTION![Reaction]
                  objSQL("UPDATEFLAG") = rsALLERGYREACTION![UpdateFlag]
                  objSQL("CREATEDBY") = rsALLERGYREACTION![CREATEDBY]
                  objSQL("CREATEDON") = rsALLERGYREACTION![CREATEDON]
                  objSQL("UPDATEDBY") = rsALLERGYREACTION![UPDATEDBY]
                  objSQL("UPDATEDON") = rsALLERGYREACTION![UPDATEDON]
                  objSQL.Update
          
                  If Err.Number <> 0 Then
                    errorCode = ERROR_PROCESSING_BMIST_ENCOUNTER_DATA
                    Err.Clear
                  End If
                End If
              End If
              
        Case "ZVT" 'Vitals
            If isBMISTEnc Then
                lngResult = FillRecordsetFromZVTNode(itemNode, gobjPatient.unitNumber, rsVITALS, Nothing, rsVITALSIGNSVALUE)
                  
                sql = "Select Unit_Number from VITALS Where EntryDate = " & oDatabase.SQLDate(rsVITALS("EntryDate")) & " And Unit_Number = " & gobjPatient.unitNumber
                  
                objSQL.Execute sql
                
                If objSQL.RecordCount = 0 Then
                    lDataID = objSQL.GetNextID(exMAINSEQ)
                    
                    If lngResult = ERROR_SUCCESS Then
                        objSQL.Execute "Select * from VITALS where DataID = 0"
                        objSQL.AddNew
                        objSQL("CLINICIANNCID") = providerNCID
                        objSQL("UNIT_NUMBER") = gobjPatient.unitNumber
                        objSQL("FACILITY_NCID") = gobjLogon.FacilityNCID
                        objSQL("ENCOUNTERNUMBER") = encounterID
                        objSQL("DATAID") = lDataID
                        objSQL("PANELTYPE") = rsVITALS![PANELTYPE]
                        objSQL("ENTRYDATE") = rsVITALS![EntryDate]
                        objSQL("CREATEDON") = Now
                        objSQL("UPDATEFLAG") = "N"
                        objSQL.Update
                        
                        objSQL.Execute "Select * from VITALSIGNSVALUE where DataID=0"
                        rsVITALSIGNSVALUE.MoveFirst
                        Do Until rsVITALSIGNSVALUE.EOF
                            objSQL.AddNew
                            objSQL("DATAID") = lDataID
                            objSQL("VITALSTYPE") = rsVITALSIGNSVALUE![VITALSTYPE]
                            objSQL("VITALSVALUE") = rsVITALSIGNSVALUE![VITALSVALUE]
                            objSQL("UNITS") = rsVITALSIGNSVALUE![UNITS]
                            objSQL("CREATEDBY") = providerNCID
                            objSQL("CREATEDON") = Now
                            objSQL.Update
                            rsVITALSIGNSVALUE.MoveNext
                        Loop
                        
                        If Err.Number <> 0 Then
                          errorCode = ERROR_PROCESSING_BMIST_ENCOUNTER_DATA
                          Err.Clear
                        End If
                    End If
                    
                End If
            End If
            
        Case "ZWS" 'Work Status
            If isBMISTEnc Then
                lngResult = FillRecordsetFromZWSNode(itemNode, gobjPatient.unitNumber, rsWORKSTATUS)
                
                'Only load this once
                If workStatuses Is Nothing Then
                    Set workStatuses = New workStatuses
                    workStatuses.Retrieve gobjPatient.unitNumber, encounterID, gobjLogon.FacilityNCID, providerNCID
                End If
                
                workStatuses(CStr(rsWORKSTATUS![Work_Status_NCID])).Selected = rsWORKSTATUS![Work_Status_Selected]
                workStatuses(CStr(rsWORKSTATUS![Work_Status_NCID])).Qualified = rsWORKSTATUS![Work_Status_Qualified]
                workStatuses.Save 'This saves to the workstatus table and the mmigenericdata table
                
            End If
        
        Case "RXE" 'Outside Medications
            If isBMISTEnc Then
                lngResult = FillRecordsetFromBMISTRXENode(itemNode, gobjPatient.unitNumber, rsMedications, encounterDate)
                sql = "Select Unit_Number from MEDICATIONS Where " _
                      & " Unit_Number = " & gobjPatient.unitNumber _
                      & " And SIG = '" & rsMedications("SIG") _
                      & "' And STARTDATE = " & oDatabase.SQLDate(rsMedications("STARTDATE")) _
                      & " And FDB_MEDID = " & rsMedications("FDB_MEDID")
                  
                objSQL.Execute sql
                
                If objSQL.RecordCount = 0 Then
                    lDataID = objSQL.GetNextID(exMAINSEQ)
                    
                    If lngResult = ERROR_SUCCESS Then
                        objSQL.Execute "Select * from MEDICATIONS where DATAID = 0"
                        objSQL.AddNew
                        objSQL("DATAID") = lDataID
                        objSQL("UNIT_NUMBER") = gobjPatient.unitNumber
                        objSQL("MEDICATIONNCID") = rsMedications![FDB_MEDID]
                        objSQL("SIG") = rsMedications![SIG]
                        objSQL("OTC") = 0
                        objSQL("ORDERNUMBER") = providerNCID & "-" & gobjPatient.unitNumber & "-" & Time & "-0"
                        objSQL("ORDERDATE") = rsMedications![STARTDATE]
                        objSQL("STARTDATE") = rsMedications![STARTDATE]
                        objSQL("ORDERINGPROVIDERNCID") = providerNCID
                        objSQL("ORDERINGPROVIDER") = providerName
                        objSQL("CLINICNCID") = gobjLogon.NursingDivisionNCID
                        objSQL("PATIENTTYPE") = "O"
                        objSQL("STATUSNCID") = 1024
                        objSQL("DISPENSINGLOCATIONNCID") = gobjLogon.FacilityNCID
                        objSQL("MEDICATIONSCOMMENT") = rsMedications![MEDICATIONSCOMMENT]
                        objSQL("CDRDATAID") = 0
                        objSQL("STORINGAPP") = NCID_MEDICATIONS_APPLICATION
                        objSQL("NON_CODED_MEDICATION_NAME") = rsMedications![NON_CODED_MEDICATION_NAME]
                        objSQL("ENC_NUM") = encounterID
                        objSQL("FDB_MEDID") = rsMedications![FDB_MEDID]
                        objSQL.Update
                        
                        If Err.Number <> 0 Then
                          errorCode = ERROR_PROCESSING_BMIST_ENCOUNTER_DATA
                          Err.Clear
                        End If
                    End If
                End If
            End If
        End Select
      Next
    End If
  End If
  
  CreateExternalEncounter = errorCode
    
  'Refresh the patient icon bar
  gobjPatient.LoadPatient gobjPatient.unitNumber, ""

  GoTo Cleanup

Cleanup:
    On Error Resume Next
    Screen.MousePointer = vbDefault
    Set rsALLERGIES = Nothing
    Set rsALLERGYCOMMENT = Nothing
    Set rsALLERGYREACTION = Nothing
    Set rsDIAGNOSIS = Nothing
    Set rsENC_DISP = Nothing
    Set rsLISTTOOLRECORD = Nothing
    Set rsLISTTOOLRECORDS = Nothing
    
    Exit Function
    
ErrHnd:
    errorCode = ERROR_PROCESSING_BMIST_ENCOUNTER_DATA
    Resume Next
    Resume

End Function

Private Function EncounterAlreadyAdded(ByVal encounterSignedDate As Date, ByVal unitNumber As String) As Boolean
    Dim sqlOps  As ICHCSII_SQl.ISqlOpsEx
    Dim sql     As String
    
    Set sqlOps = oDatabase.CHCSII_SQLOPS(Auto)
    
    sql = "select EncounterNumber from encounters where WHENSIGNED = " & oDatabase.SQLDate(encounterSignedDate) & " AND UNIT_NUMBER = " & unitNumber
    
    sqlOps.Execute sql
    If sqlOps.RecordCount > 0 Then
        EncounterAlreadyAdded = True
    Else
        EncounterAlreadyAdded = False
    End If
End Function

Private Function EncounterRtfAlreadyAdded(ByVal encounterSignedDate As Date, ByVal encounterNumber As String) As Boolean
    Dim sqlOps  As ICHCSII_SQl.ISqlOpsEx
    Dim sql     As String
    
    Set sqlOps = oDatabase.CHCSII_SQLOPS(Auto)
    
    sql = "select ENCOUNTERNUMBER from ENC_RTFS Where DTS = " & oDatabase.SQLDate(encounterSignedDate) & " And ENCOUNTERNUMBER = " & encounterNumber
    
    sqlOps.Execute sql
    If sqlOps.RecordCount > 0 Then
        EncounterRtfAlreadyAdded = True
    Else
        EncounterRtfAlreadyAdded = False
    End If
End Function


Private Function AddEncounterRtf(ByVal encounterNumber As String, _
                                 ByVal rtfDate As Date, _
                                 ByVal rtfNote As String, _
                                 ByVal providerNCID As String, _
                                 ByVal providerName As String, _
                                 ByVal title As String, _
                                 ByVal isBMISTEnc As Boolean, _
                                 ByVal reassessed As Boolean) As DMErrorCodes
    
    Dim sql                 As String
    Dim sqlOps              As ICHCSII_SQl.ISqlOpsEx
    Dim rtfIndex            As Long
    Dim noteType            As String
    
    AddEncounterRtf = DMErrorCodes.ERROR_SUCCESS

    Set sqlOps = oDatabase.CHCSII_SQLOPS(Auto)
      
      rtfIndex = sqlOps.GetNextID(exMAINSEQ)
      
      If isBMISTEnc Then
          noteType = "NOTE"
      Else
          noteType = "RTF"
      End If
      
      sql = "insert into enc_rtfs (FacilityNCID,ENCOUNTERNUMBER,DTS,RTFTYPE,ENC_RTFSINDEX," _
       & "NOTEID,UserNCID,USERNAME,DATEADDED,Title,COMPLETE)"
      sql = sql & " Values (" & gobjLogon.FacilityNCID & "," & encounterNumber & "," & oDatabase.SQLDate(rtfDate) _
      & " - 0.002,'" & noteType & "'," & rtfIndex & ",1," & providerNCID & "," & SQLQuote(Left(providerName, 20)) & ",sysdate,'" & title & "','Y')"
      
      If reassessed Then
        sql = Replace(sql, " - 0.002", vbNullString)
        sql = Replace(sql, ",1,", ",2,")
      End If
      
      sqlOps.Execute sql
      
      sqlOps.Execute "Select * from enc_rtfs WHERE ENCOUNTERNUMBER = " & encounterNumber & " and ENC_RTFSINDEX = " & rtfIndex
      sqlOps.AppendAsChunk "DOC", ConvertTextToRTF(rtfNote)
      sqlOps("OriginalSize") = 0
      sqlOps.Update
      
      If Err.Number <> 0 Then
        AddEncounterRtf = ERROR_ADDING_BMIST_ENCOUNTER_RTF_DATA
        Err.Clear
      End If

End Function


'<< SCR #29291;   Developer: Brian Mowbray 11/14/2002 12:04 PM
Private Function ConvertTextToBoldRTF(ByRef Text As String) As String
    Dim tempText As String
    If Left$(Text, 5) = rtfTag Then
        ConvertTextToBoldRTF = Text
    Else
        If InStr(1, Text, rtfTag) > 0 Then
            tempText = Text
        Else
            tempText = HandleSpecialCharactersInRTF(Text)
        End If
        tempText = HandleNewLineAndTabInRTF(tempText)
        
        ConvertTextToBoldRTF = "{\rtf1\ansi\ansicpg1252\deff0\deflang1033{\fonttbl{\f0\fswiss\fcharset0 Arial;}}" _
        & "{\colortbl ;\red0\green0\blue0;}" _
        & "\viewkind4\uc1\pard\cf1\ul\b\f0\fs16 " & tempText & "\b0\ulnone}"
            
    End If
    
End Function

Private Function HandleSpecialCharactersInRTF(ByVal Text As String) As String
    Dim tempText As String
    
    tempText = Replace(Text, "\", "\\")
    tempText = Replace(tempText, "{", "\{")
    tempText = Replace(tempText, "}", "\}")
    
    HandleSpecialCharactersInRTF = tempText
End Function

Private Function HandleNewLineAndTabInRTF(ByVal Text As String) As String
    Dim tempText As String
    
    tempText = Replace(Text, Chr(13) & Chr(10), "\par ")
    tempText = Replace(tempText, Chr(13), "\par ")
    tempText = Replace(tempText, Chr(10), "\par ")
    tempText = Replace(tempText, Chr(9), "\tab ")
    
    HandleNewLineAndTabInRTF = tempText
End Function

Private Function ConvertTextToRTF(ByRef Text As String) As String
    Dim tempText As String
    If Left$(Text, 5) = rtfTag Then
        ConvertTextToRTF = Text
    Else
        If InStr(1, Text, rtfTag) > 0 Then
            tempText = Text
        Else
            tempText = HandleSpecialCharactersInRTF(Text)
        End If
        tempText = HandleNewLineAndTabInRTF(tempText)

        ConvertTextToRTF = "{\rtf1\ansi\ansicpg1252\uc0\deff0" _
            & "{\fonttbl{\f0\fswiss\fcharset0\fprq2 Arial;}}" _
            & "{\colortbl;\red0\green0\blue0;}" _
            & "\deftab720 \plain \pard \f0\fs16 " _
            & tempText & " }"
    End If
    
End Function


' Begin:SCR #28326;   Developer: MAG 10/16/2002 12:40 PM
Private Function FillRecordsetFromZRENode(ByVal oZRENode As MSXML.IXMLDOMElement, ByVal lngUnitNumber As Long, _
                                        ByRef rsMedications As Adodb.Recordset, ByRef rsRXLIST As Adodb.Recordset) As DMErrorCodes

    Dim oItemNode As MSXML.IXMLDOMNode
    
    On Error GoTo ErrHandler
    
    If oZRENode Is Nothing Then
        Exit Function
    End If
    
    For Each oItemNode In oZRENode.childNodes
        Select Case UCase(oItemNode.nodeName)
            Case "ZRE.2"
                 rsMedications.Fields("UNIT_NUMBER").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZRE.3"
                rsMedications.Fields("MEDICATIONSCOMMENT").Value = NullToNothing(oItemNode.Text)
            Case "ZRE.4"
                rsMedications.Fields("DISPENSINGLOCATIONNCID").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZRE.5"
                rsMedications.Fields("DAYSOFSUPPLY").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZRE.6"
                rsMedications.Fields("DATAID").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZRE.7"
                rsMedications.Fields("QUANTITY").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZRE.8"
                rsMedications.Fields("OTC").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZRE.9"
                rsMedications.Fields("ORDEREXPIRATIONDATE").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZRE.10"
                rsMedications.Fields("ORDERINGPROVIDER").Value = NullToNothing(oItemNode.Text)
            Case "ZRE.11"
                rsMedications.Fields("CLINICNCID").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZRE.12"
                rsMedications.Fields("PATIENTTYPE").Value = NullToNothing(oItemNode.Text)
            Case "ZRE.13"
                rsMedications.Fields("UPDATEFLAG").Value = NullToNothing(oItemNode.Text)
            Case "ZRE.14"
                rsMedications.Fields("CDR_DATAID").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZRE.15"
                rsMedications.Fields("STORINGAPP").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZRE.16"
                rsMedications.Fields("CREATEDBY").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZRE.17"
                rsMedications.Fields("CREATEDON").Value = XMLToOracleDate(oItemNode.Text)
            Case "ZRE.18"
                rsMedications.Fields("UPDATEDBY").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZRE.19"
                rsMedications.Fields("UPDATEDON").Value = XMLToOracleDate(oItemNode.Text)
            ' Begin:SCR #47403;   Developer:  12/24/2003
            Case "ZRE.20"
                rsMedications.Fields("ORDERINGPROVIDERNCID").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZRE.21"
                rsMedications.Fields("STATUSNCID").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZRE.22"
                rsMedications.Fields("NATIONALDRUGCODE").Value = NullToNothing(oItemNode.Text)
            Case "ZRE.23"
                rsMedications.Fields("NON_CODED_MEDICATION_NAME").Value = NullToNothing(oItemNode.Text)
            Case "ZRE.24"
                rsMedications.Fields("TMIP_STATUS").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZRE.25"
                rsMedications.Fields("ENC_NUM").Value = NullToNothing(oItemNode.Text)
            Case "ZRE.26"
                rsRXLIST.Fields("FACILITY_NCID").Value = gobjLogon.FacilityNCID
                'rsRXLIST.Fields("FACILITY_NCID").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZRE.27"
                rsRXLIST.Fields("DRUG_NAME").Value = NullToNothing(oItemNode.Text)
            Case "ZRE.28"
                rsRXLIST.Fields("DRUG_IEN").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZRE.29"
                rsRXLIST.Fields("DEFAULT_UNIT").Value = NullToNothing(oItemNode.Text)
            Case "ZRE.30"
                rsRXLIST.Fields("NDC_INDICATOR").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZRE.31"
                rsRXLIST.Fields("DRUG_CHECK_DISABLED").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZRE.32"
                rsRXLIST.Fields("DOSAGE_STRENGTH").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZRE.33"
                rsRXLIST.Fields("CONTENT_UNIT").Value = NullToNothing(oItemNode.Text)
            Case "ZRE.34"
                rsRXLIST.Fields("DOSAGE_FORM").Value = NullToNothing(oItemNode.Text)
            Case "ZRE.35"
                rsRXLIST.Fields("SYNONYMS").Value = NullToNothing(oItemNode.Text)
            Case "ZRE.36"
                rsRXLIST.Fields("DEFAULT_SIG").Value = NullToNothing(oItemNode.Text)
            Case "ZRE.37"
                rsRXLIST.Fields("DEFAULT_REFILL").Value = XMLToOracleNumber(oItemNode.Text)
            Case "ZRE.38"
                rsRXLIST.Fields("DEFAULT_ROUTE").Value = NullToNothing(oItemNode.Text)
            Case "ZRE.39"
                rsRXLIST.Fields("CURRENT_STOCK").Value = XMLToOracleNumber(oItemNode.Text)
            ' End: SCR #47403;
            End Select
    Next oItemNode
    
    FillRecordsetFromZRENode = DMErrorCodes.ERROR_SUCCESS

Exit Function

ErrHandler:
    FillRecordsetFromZRENode = DMErrorCodes.ERROR_EIC_XML_TO_RECORDSET
    If gobjShared Is Nothing Then
        Set gobjShared = New CWShared
    End If
    gobjShared.ShowVBError Err.Number, Err.Description, "FillRecordsetFromZRENode Function", "DMBE", vbExclamation
    Err.Clear
End Function

Private Function FillRecordsetFromRXENode(ByVal oRXENode As MSXML.IXMLDOMElement, ByVal lngUnitNumber As Long, _
                                        ByRef rsMedications As Adodb.Recordset) As DMErrorCodes

    Dim oItemNode As MSXML.IXMLDOMNode
    Dim oChildNode As MSXML.IXMLDOMNode
    Dim oChildNodeLevel1 As MSXML.IXMLDOMNode
    Dim asDates() As String
    On Error GoTo ErrHandler
    
    If oRXENode Is Nothing Then
        Exit Function
    End If
    
    For Each oItemNode In oRXENode.childNodes
        Select Case UCase(oItemNode.nodeName)
            Case "RXE.1"    'Format ^^^StartDate^EndDate
                asDates = Split(oItemNode.Text, "^")
                '**********************************************************************
                'SCR 64927
                If UBound(asDates) >= 3 Then
                  rsMedications.Fields("STARTDATE").Value = XMLToOracleDate(asDates(3))
                End If
                If UBound(asDates) >= 4 Then
                  rsMedications.Fields("ENDDATE").Value = XMLToOracleDate(asDates(4))
                End If
                '**********************************************************************
            Case "RXE.2"
                Set oChildNode = oItemNode.selectSingleNode("CE.4")
                If Not oChildNode Is Nothing Then
                    rsMedications.Fields("MEDICATIONNCID").Value = XMLToOracleNumber(oChildNode.Text)
                End If
            Case "RXE.7.LST"
                Set oChildNode = oItemNode.selectSingleNode("RXE.7")
                If Not (oChildNode Is Nothing) Then
                    Set oChildNodeLevel1 = oChildNode.selectSingleNode("CE.2")
                    If Not (oChildNode Is Nothing) Then
                        rsMedications.Fields("SIG").Value = NullToNothing(oChildNodeLevel1.Text)
                    End If
                End If
            Case "RXE.12"
                rsMedications.Fields("Refills").Value = XMLToOracleNumber(oItemNode.Text)
            Case "RXE.16"
                rsMedications.Fields("RefillsRemaining").Value = XMLToOracleNumber(oItemNode.Text)
            Case "RXE.18"
                rsMedications.Fields("FillDate").Value = XMLToOracleDate(oItemNode.Text)
        End Select
    Next oItemNode
    
    FillRecordsetFromRXENode = DMErrorCodes.ERROR_SUCCESS

Exit Function

ErrHandler:
    FillRecordsetFromRXENode = DMErrorCodes.ERROR_EIC_XML_TO_RECORDSET
    If gobjShared Is Nothing Then
        Set gobjShared = New CWShared
    End If
    gobjShared.ShowVBError Err.Number, Err.Description, "FillRecordsetFromRXENode Function", "DMBE", vbExclamation
    Err.Clear
End Function

Private Function FillRecordsetFromBMISTRXENode(ByVal oRXENode As MSXML.IXMLDOMElement, ByVal lngUnitNumber As Long, _
                                        ByRef rsMedications As Adodb.Recordset, EncDate As Date) As DMErrorCodes

    Dim oItemNode As MSXML.IXMLDOMNode
    Dim oChildNode As MSXML.IXMLDOMNode
    Dim oChildNodeLevel1 As MSXML.IXMLDOMNode
    Dim asDates() As String
    On Error GoTo ErrHandler
    
    If oRXENode Is Nothing Then
        Exit Function
    End If
    
    rsMedications.AddNew
    
    For Each oItemNode In oRXENode.childNodes
        Select Case UCase(oItemNode.nodeName)
            Case "RXE.1"
                If oItemNode.Text = "" Then
                    rsMedications.Fields("STARTDATE").Value = EncDate
                Else
                    rsMedications.Fields("STARTDATE").Value = XMLToOracleDate(oItemNode.Text)
                End If
            Case "RXE.2"
                Set oChildNode = oItemNode.selectSingleNode("CE.4")
                If Not oChildNode Is Nothing Then
                    rsMedications.Fields("FDB_MEDID").Value = XMLToOracleNumber(oChildNode.Text)
                End If
            Case "RXE.7.LST"
                Set oChildNode = oItemNode.selectSingleNode("RXE.7")
                If Not (oChildNode Is Nothing) Then
                    Set oChildNodeLevel1 = oChildNode.selectSingleNode("CE.2")
                    If Not (oChildNode Is Nothing) Then
                        rsMedications.Fields("SIG").Value = NullToNothing(oChildNodeLevel1.Text)
                    End If
                End If
            Case "RXE.8"
                rsMedications.Fields("MEDICATIONSCOMMENT").Value = NullToNothing(oItemNode.Text)
            Case "RXE.9"
                rsMedications.Fields("NON_CODED_MEDICATION_NAME").Value = NullToNothing(oItemNode.Text)
            Case "RXE.12"
                rsMedications.Fields("Refills").Value = XMLToOracleNumber(oItemNode.Text)
            Case "RXE.16"
                rsMedications.Fields("RefillsRemaining").Value = XMLToOracleNumber(oItemNode.Text)
            Case "RXE.18"
                rsMedications.Fields("FillDate").Value = XMLToOracleDate(oItemNode.Text)
        End Select
    Next oItemNode
    
    rsMedications.Update
    
    FillRecordsetFromBMISTRXENode = DMErrorCodes.ERROR_SUCCESS

Exit Function

ErrHandler:
    FillRecordsetFromBMISTRXENode = DMErrorCodes.ERROR_EIC_XML_TO_RECORDSET
    If gobjShared Is Nothing Then
        Set gobjShared = New CWShared
    End If
    gobjShared.ShowVBError Err.Number, Err.Description, "FillRecordsetFromBMISTRXENode Function", "DMBE", vbExclamation
    Err.Clear
    Exit Function
    Resume
End Function


Private Function FillRecordsetFromORCNode(ByVal oORCNode As MSXML.IXMLDOMElement, ByVal lngUnitNumber As Long, _
                                        ByRef rsMedications As Adodb.Recordset) As DMErrorCodes

    Dim oItemNode As MSXML.IXMLDOMNode
    Dim oChildNode As MSXML.IXMLDOMNode
    On Error GoTo ErrHandler
    
    If oORCNode Is Nothing Then
        Exit Function
    End If
    
    For Each oItemNode In oORCNode.childNodes
        Select Case UCase(oItemNode.nodeName)
            Case "ORC.2"
                Set oChildNode = oItemNode.selectSingleNode("EI.1")
                If Not (oChildNode Is Nothing) Then
                    rsMedications.Fields("ORDERNUMBER").Value = NullToNothing(oChildNode.Text)
                End If
            Case "ORC.9"
                rsMedications.Fields("ORDERDATE").Value = XMLToOracleDate(oItemNode.Text)
            Case "ORC.15"
                rsMedications.Fields("STARTDATE").Value = XMLToOracleDate(oItemNode.Text)
        End Select
    Next oItemNode
    
    FillRecordsetFromORCNode = DMErrorCodes.ERROR_SUCCESS

Exit Function

ErrHandler:
    FillRecordsetFromORCNode = DMErrorCodes.ERROR_EIC_XML_TO_RECORDSET
    If gobjShared Is Nothing Then
        Set gobjShared = New CWShared
    End If
    gobjShared.ShowVBError Err.Number, Err.Description, "FillRecordsetFromORCNode Function", "DMBE", vbExclamation
    Err.Clear
End Function

Private Function FillRecordsetFromRDE_O01Node(ByVal oRDENode As MSXML.IXMLDOMElement, ByVal lngUnitNumber As Long, _
                                        ByRef rsMedications As Adodb.Recordset, ByRef rsRXLIST As Adodb.Recordset) As DMErrorCodes

    Dim enmResult As DMErrorCodes
    Dim oGroupNode As MSXML.IXMLDOMElement
    Dim oSegmentNode As MSXML.IXMLDOMElement
    On Error GoTo ErrHandler
    
    If oRDENode Is Nothing Then
        Exit Function
    End If
    
    For Each oGroupNode In oRDENode.childNodes
        rsMedications.AddNew
        rsRXLIST.AddNew
        For Each oSegmentNode In oGroupNode.childNodes
            Select Case UCase(oSegmentNode.nodeName)
            Case "ORC"
                enmResult = FillRecordsetFromORCNode(oSegmentNode, lngUnitNumber, rsMedications)
                If enmResult <> ERROR_SUCCESS Then
                    GoTo ErrHandler
                End If
            Case "ZRE"
                enmResult = FillRecordsetFromZRENode(oSegmentNode, lngUnitNumber, rsMedications, rsRXLIST)
                If enmResult <> ERROR_SUCCESS Then
                    GoTo ErrHandler
                End If
                
            Case "RXE"
                enmResult = FillRecordsetFromRXENode(oSegmentNode, lngUnitNumber, rsMedications)
                If enmResult <> ERROR_SUCCESS Then
                    GoTo ErrHandler
                End If
            Case "ORC"
            End Select
        Next oSegmentNode
        rsRXLIST.Update
        rsMedications.Update
    Next oGroupNode
    
    FillRecordsetFromRDE_O01Node = DMErrorCodes.ERROR_SUCCESS

Exit Function

ErrHandler:
    FillRecordsetFromRDE_O01Node = DMErrorCodes.ERROR_EIC_XML_TO_RECORDSET
    If gobjShared Is Nothing Then
        Set gobjShared = New CWShared
    End If
    gobjShared.ShowVBError Err.Number, Err.Description, "FillRecordsetFromRDE_O01Node Function", "DMBE", vbExclamation
    Err.Clear
End Function
' End: SCR #28326;

Private Function ReformatDate(ByVal textDate As String) As Date
    If Len(textDate) = 14 Then
        textDate = Mid$(textDate, 5, 2) & "/" & Mid$(textDate, 7, 2) & "/" & Mid$(textDate, 1, 4) & " " & Mid$(textDate, 9, 2) & ":" & Mid$(textDate, 11, 2) & ":" & Mid$(textDate, 13, 2)
        ReformatDate = CDate(textDate)
    End If
End Function

