VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HL7XMLBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' SCR #39837;   Developer: Mag 08/12/2003 04:10 PM
' Made changes to OriginalSize > 1 instead of 0 . (ie. Fields("OriginalSize").Value) > 1)
Option Explicit
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" _
    (ByVal lpBuffer As String, nSize As Long) As Boolean

Private Function IsValidDate(ByVal strDate As String) As Boolean

    IsValidDate = False
    If IsDate(strDate) = True Then
        If CLng(CDate(strDate)) <> 0 Then 'To take care of "12:00:00 AM"
            IsValidDate = True
        End If
    End If
            
End Function

Private Function NullToZero(ByVal vVal As Variant) As String
  On Error Resume Next
  
  If IsNull(vVal) Then
    NullToZero = "0"
  ElseIf Len(Trim$(vVal)) = 0 Then
    NullToZero = "0"
  Else
    NullToZero = vVal
  End If
  
End Function

Public Function ReturnNewSYMOBXNode(ByVal lngCount As Long, ByVal rsSYMPTOMS As ADODB.Recordset) As MSXML.IXMLDOMElement

' Observation Segment
' sample function output
'   <OBX>
'      <OBX.1>1</OBX.1>
'      <OBX.2>ST</OBX.2>
'      <OBX.3>
'         <CE.1></CE.1>
'         <CE.2>CODED_SYMPTOM</CE.2>
'         <CE.3></CE.3>
'         <CE.4></CE.4>
'         <CE.5></CE.5>
'         <CE.6></CE.6>
'      </OBX.3>
'      <OBX.4>1</OBX.4>
'      <OBX.5.LST>
'         <OBX.5>headache</OBX.5>
'      </OBX.5.LST>
'      <OBX.11>P</OBX.11>
'   </OBX>

    Dim oDom As MSXML.DOMDocument
    Dim oOBXNode As MSXML.IXMLDOMElement
    Dim oOBXChildLevel1 As MSXML.IXMLDOMElement
    Dim oOBXChildLevel2 As MSXML.IXMLDOMElement
    
    On Error GoTo ErrTrap

    Set oDom = New MSXML.DOMDocument
    Set oOBXNode = oDom.createElement("OBX")
    
    ' Set ID
    Set oOBXChildLevel1 = oDom.createElement("OBX.1")
    oOBXChildLevel1.Text = CStr(lngCount)
    oOBXNode.appendChild oOBXChildLevel1

    ' Value Type - ST for Symptoms
    Set oOBXChildLevel1 = oDom.createElement("OBX.2")
    oOBXChildLevel1.Text = "ST"
    oOBXNode.appendChild oOBXChildLevel1

    ' Observation Identifier
    Set oOBXChildLevel1 = oDom.createElement("OBX.3")
        ' Identifier
        Set oOBXChildLevel2 = oDom.createElement("CE.1")
        oOBXChildLevel2.Text = NullToNothing(rsSYMPTOMS.Fields("SnoID").Value)
        oOBXChildLevel1.appendChild oOBXChildLevel2
        ' Text
        Set oOBXChildLevel2 = oDom.createElement("CE.2")
        oOBXChildLevel2.Text = "CODED_SYMPTOM"
        oOBXChildLevel1.appendChild oOBXChildLevel2
        ' Name of Coding System
        Set oOBXChildLevel2 = oDom.createElement("CE.3")
        oOBXChildLevel2.Text = "99MED"
        oOBXChildLevel1.appendChild oOBXChildLevel2
        ' Alternate Identifier
        Set oOBXChildLevel2 = oDom.createElement("CE.4")
        oOBXChildLevel2.Text = ""
        oOBXChildLevel1.appendChild oOBXChildLevel2
        ' Alternate Text - Table WhatRS:
        Set oOBXChildLevel2 = oDom.createElement("CE.5")
        oOBXChildLevel2.Text = ""
        oOBXChildLevel1.appendChild oOBXChildLevel2
        ' Alternate Name of Coding System
        Set oOBXChildLevel2 = oDom.createElement("CE.6")
        oOBXChildLevel2.Text = ""
        oOBXChildLevel1.appendChild oOBXChildLevel2
    oOBXNode.appendChild oOBXChildLevel1

    ' Observation SubID
    Set oOBXChildLevel1 = oDom.createElement("OBX.4")
    oOBXChildLevel1.Text = CStr(lngCount)
    oOBXNode.appendChild oOBXChildLevel1

    ' Observation Value
    Set oOBXChildLevel1 = oDom.createElement("OBX.5.LST")
        Set oOBXChildLevel2 = oDom.createElement("OBX.5")
       '<< Begin:SCR #26545;   Developer: Darren Magnuson 09/05/2002 2:20 PM
        oOBXChildLevel2.Text = NullToNothing(rsSYMPTOMS.Fields("MEDCINDESC").Value)
        
'        'strTemp = NullToNothing(rsCOMPLAINTS.Fields("COMPLAINTNCID").Value) '<SCR 23005 CC>
'        strTemp = NullToNothing(rsSYMPTOMS.Fields("SnoID").Value) '<SCR 23005 CC>
'        If strTemp <> "" Then
'            'If mobjConvertSno Is Nothing Then '<SCR 23005 CC>
'            '    Set mobjConvertSno = New actxSearchSnoMed.IConvert '<SCR 23005 CC>
'            'End If '<SCR 23005 CC>
'            strTemp = mobjConvertSno.SnoDescription(strTemp)
'            oOBXChildLevel2.Text = strTemp
'        End If
        '<< End:SCR#26545
        oOBXChildLevel1.appendChild oOBXChildLevel2
    oOBXNode.appendChild oOBXChildLevel1
    
    ' Observation Result Status
    Set oOBXChildLevel1 = oDom.createElement("OBX.11")
     '<< Begin:SCR #26545;   Developer: Darren Magnuson 09/05/2002 2:20 PM
'    If IsNull(rsCOMPLAINTS.Fields("STATUSNCID").Value) Then
'        strTemp = "X"
'    Else
'        Select Case DecodeSYMStatusNCID(rsCOMPLAINTS.Fields("STATUSNCID").Value)
'            Case "All"
'                strTemp = "X"
'            Case "Active"
'                strTemp = "P"
'            Case "Deleted(Error)"
'                strTemp = "D"
'            Case "Inactive"
'                strTemp = "F"
'            Case Else
'                strTemp = "X"
'        End Select
'    End If
    oOBXChildLevel1.Text = "F" 'strTemp
    '<< End:SCR#26545
    oOBXNode.appendChild oOBXChildLevel1

    Set ReturnNewSYMOBXNode = oOBXNode
    
Exit Function

ErrTrap:
   If gobjShared Is Nothing Then
      Set gobjShared = New CWShared
   End If
   gobjShared.ShowVBError Err.Number, Err.Description, "ReturnNewSYMOBXNode Function", "DMBE", vbExclamation
   'Err.Clear
   Resume Next 'There is no reason this procedure should not continue to create xml if it errors
   
End Function

Public Function ReturnNewSYMOBRNode(ByVal lngCount As Long, ByVal RS As ADODB.Recordset) As MSXML.IXMLDOMElement
' Observation Request
' sample function output
'   <OBR>
'      <OBR.1>1</OBR.1>
'      <OBR.3>
'         <EI.1>7568</EI.1>
'         <EI.2>CHCSII-T</EI.2>
'         <EI.3></EI.3>
'         <EI.4></EI.4>
'      </OBR.3>
'      <OBR.4>
'         <CE.1></CE.1>
'         <CE.2>SYMPTOMS</CE.2>
'         <CE.3></CE.3>
'         <CE.4></CE.4>
'         <CE.5></CE.5>
'         <CE.6></CE.6>
'      </OBR.4>
'   </OBR>

    Dim oDom As MSXML.DOMDocument
    Dim oOBRNode As MSXML.IXMLDOMElement
    Dim oOBRChildLevel1 As MSXML.IXMLDOMElement
    Dim oOBRChildLevel2 As MSXML.IXMLDOMElement
    Dim strDocFirstName As String
    Dim strDocLastName As String
    Dim strDocMiddleName As String
    Dim strDocID As String
    Dim strDocSpecialty As String
    Dim strTemp As String
    Dim blnRsIsGood As Boolean
    
    On Error GoTo ErrTrap

    Set oDom = New MSXML.DOMDocument
    Set oOBRNode = oDom.createElement("OBR")
    
    ' Count
    Set oOBRChildLevel1 = oDom.createElement("OBR.1")
    oOBRChildLevel1.Text = CStr(lngCount)
    oOBRNode.appendChild oOBRChildLevel1

    ' Placer Order Number - Table Clinnote:DATAID
    Set oOBRChildLevel1 = oDom.createElement("OBR.3")
        Set oOBRChildLevel2 = oDom.createElement("EI.1")
        oOBRChildLevel2.Text = NullToNothing(RS.Fields("DATAID").Value)
        oOBRChildLevel1.appendChild oOBRChildLevel2
    
        Set oOBRChildLevel2 = oDom.createElement("EI.2")
        oOBRChildLevel2.Text = "CHCSII-T"
        oOBRChildLevel1.appendChild oOBRChildLevel2
    
        Set oOBRChildLevel2 = oDom.createElement("EI.3")
        oOBRChildLevel2.Text = ""
        oOBRChildLevel1.appendChild oOBRChildLevel2
    
        Set oOBRChildLevel2 = oDom.createElement("EI.4")
        oOBRChildLevel2.Text = ""
        oOBRChildLevel1.appendChild oOBRChildLevel2
    oOBRNode.appendChild oOBRChildLevel1
    
    ' Universal Service ID
    Set oOBRChildLevel1 = oDom.createElement("OBR.4")
        ' Identifier
        Set oOBRChildLevel2 = oDom.createElement("CE.1")
        oOBRChildLevel2.Text = ""
        oOBRChildLevel1.appendChild oOBRChildLevel2
        ' Text
        Set oOBRChildLevel2 = oDom.createElement("CE.2")
        oOBRChildLevel2.Text = "SYMTOMS"
        oOBRChildLevel1.appendChild oOBRChildLevel2
        ' Name of Coding System -
        Set oOBRChildLevel2 = oDom.createElement("CE.3")
        oOBRChildLevel2.Text = ""
        oOBRChildLevel1.appendChild oOBRChildLevel2
        ' Alternate Identifier
        Set oOBRChildLevel2 = oDom.createElement("CE.4")
        oOBRChildLevel2.Text = ""
        oOBRChildLevel1.appendChild oOBRChildLevel2
        ' Alternate Text
        Set oOBRChildLevel2 = oDom.createElement("CE.5")
        oOBRChildLevel2.Text = ""
        oOBRChildLevel1.appendChild oOBRChildLevel2
        ' Alternate Name of Coding System
        Set oOBRChildLevel2 = oDom.createElement("CE.6")
        oOBRChildLevel2.Text = ""
        oOBRChildLevel1.appendChild oOBRChildLevel2
    oOBRNode.appendChild oOBRChildLevel1
    
    Set ReturnNewSYMOBRNode = oOBRNode
    
Exit Function

ErrTrap:
   If gobjShared Is Nothing Then
      Set gobjShared = New CWShared
   End If
   gobjShared.ShowVBError Err.Number, Err.Description, "ReturnNewSYMOBRNode Function", "DMBE", vbExclamation
   'Err.Clear
   Resume Next 'There is no reason this procedure should not continue to create xml if it errors

End Function

Public Function ReturnNewZPVNode(ByVal lngCount As Long, ByVal EncountersRS As ADODB.Recordset, Optional ByVal Enc_DispositionRS As ADODB.Recordset) As MSXML.IXMLDOMElement
' Additional Patient Visit Data Segment
' sample function output
'   <ZPV>
'      <ZPV.1>Count</ZPV.1>
'      <ZPV.2>
'           <CE.1>Template Identifier</CE.1>
'           <CE.2>Template Text</CE.2>
'     </ZPV.2>
'      <ZPV.3>Encounter Status</ZPV.3>
'      <ZPV.4>Unit Number</ZPV.4>
'      <ZPV.5>Category</ZPV.5>
'      <ZPV.6>Primary Diagnosis</ZPV.6>
'      <ZPV.7>Appointment ID</ZPV.7>
'      <ZPV.8>SADR Status</ZPV.8>
'      <ZPV.9>TimeZone</ZPV.9>
'      <ZPV.10>AP Role1</ZPV.10>
'      <ZPV.11>AP Role2</ZPV.11>
'      <ZPV.12>Was Updated</ZPV.12>
'      <ZPV.13>Date/Time Cosigned</ZPV.13>
'      <ZPV.14>Cosigner Line2</ZPV.14>
'      <ZPV.15>Cosigner Line3</ZPV.15>
'      <ZPV.16>Date/Time Signed</ZPV.16>
'      <ZPV.17>Signer Line2</ZPV.17>
'      <ZPV.18>Signer Line3</ZPV.18>
'      <ZPV.19>No Count</ZPV.19>
'      <ZPV.20>
'         <XCN.1>Additional Provider 1 ID Number (SSN)</XCN.1>
'         <XCN.2>Additional Provider 1 Family Last Name</XCN.2>
'         <XCN.3>Additional Provider 1 Family First Name</XCN.3>
'         <XCN.4>Additional Provider 1 Family Middle Name</XCN.4>
'      </ZPV.20>
'      <ZPV.21>Appointment Type</ZPV.21>
'      <ZPV.22>
'         <XCN.1>Additional Provider 2 ID Number (SSN)</XCN.1>
'         <XCN.2>Additional Provider 2 Family Last Name</XCN.2>
'         <XCN.3>Additional Provider 2 Family First Name</XCN.3>
'         <XCN.4>Additional Provider 2 Family Middle Name</XCN.4>
'      </ZPV.22>
'      <ZPV.23>Allergy Verified By Date/Time</ZPV.23>
'      <ZPV.24>Allergy Verified By</ZPV.24>
'      <ZPV.25>Who Cosigned</ZPV.25>
'      <ZPV.26>Who Signed</ZPV.26>
'      <ZPV.27>Cosigner</ZPV.27>
'      <ZPV.28>Created By</ZPV.28>
'      <ZPV.29>Created On</ZPV.29>
'      <ZPV.30>Updated By</ZPV.30>
'      <ZPV.31>Updated On</ZPV.31>
'      <ZPV.32>Encounter Number</ZPV.32>
'      <ZPV.33>
'           <CE.1>Facility NCID</CE.1>
'           <CE.2>Facility Name</CE.2>
'      </ZPV.33>
'      <ZPV.34>Disposition NCID</ZPV.34>
'      <ZPV.35>EandM Code Review</ZPV.35>
'      <ZPV.36>EandM Reviewed Date/Time</ZPV.36>
'      <ZPV.37>EandM Reviewed By</ZPV.37>
'      <ZPV.38>EandM Reviewed</ZPV.38>
'      <ZPV.39>Reviewed By Coder</ZPV.39>
'      <ZPV.40>Disposition Admin Option</ZPV.40>
'      <ZPV.41>
'           <CE.1>Disposition Identifier</CE.1>
'           <CE.2>Disposition Text</CE.2>
'      </ZPV.41>
'      <ZPV.42>Item Discussed</ZPV.42>
'      <ZPV.43>Follow Up Timeframe</ZPV.43>
'      <ZPV.44>Follow Up Comments</ZPV.44>
'      <ZPV.45>WS Disposition</ZPV.45>
'      <ZPV.46>EandM Calculation</ZPV.46>
'   </ZPV>
    
    Dim oDom As MSXML.DOMDocument
    Dim oZPVNode As MSXML.IXMLDOMElement
    Dim oZPVChildLevel1 As MSXML.IXMLDOMElement
    Dim oZPVChildLevel2 As MSXML.IXMLDOMElement
    Dim strDocFirstName As String
    Dim strDocLastName As String
    Dim strDocMiddleName As String
    Dim strDocID As String
    Dim strDocSpecialty As String
    Dim strTemp As String
    Dim blnRsIsGood As Boolean
    
    On Error GoTo ErrTrap

    Set oDom = New MSXML.DOMDocument
    Set oZPVNode = oDom.createElement("ZPV")
    
    ' Count
    Set oZPVChildLevel1 = oDom.createElement("ZPV.1")
    oZPVChildLevel1.Text = CStr(lngCount)
    oZPVNode.appendChild oZPVChildLevel1

    ' Identifier - Table Encounters:TemplateID
    Set oZPVChildLevel1 = oDom.createElement("ZPV.2")
        ' Identifier - Table Encounters:TemplateID
        Set oZPVChildLevel2 = oDom.createElement("CE.1")
        '<< Begin:SCR #36895;   Developer: Brian Mowbray 06/04/2003 11:12 AM
        'oZPVChildLevel1.Text = NullToNothing(EncountersRS.Fields("TemplateID").Value)
        oZPVChildLevel2.Text = NullToNothing(EncountersRS.Fields("TemplateID").Value)
        oZPVChildLevel1.appendChild oZPVChildLevel2
        ' Text - Table Encounters:TemplateData
        Set oZPVChildLevel2 = oDom.createElement("CE.2")
'        oZPVChildLevel1.Text = NullToNothing(EncountersRS.Fields("TemplateData").Value)
        oZPVChildLevel2.Text = NullToNothing(EncountersRS.Fields("TemplateData").Value)
        '>> End: SCR #36895;
        oZPVChildLevel1.appendChild oZPVChildLevel2
    oZPVNode.appendChild oZPVChildLevel1

    ' Encounter Status - Table Encounters:Status
    Set oZPVChildLevel1 = oDom.createElement("ZPV.3")
    oZPVChildLevel1.Text = NullToNothing(EncountersRS.Fields("Status").Value)
    oZPVNode.appendChild oZPVChildLevel1
    
    ' Unit Number - Table Encounters:Unit_Number
    Set oZPVChildLevel1 = oDom.createElement("ZPV.4")
    oZPVChildLevel1.Text = NullToNothing(EncountersRS.Fields("Unit_Number").Value)
    oZPVNode.appendChild oZPVChildLevel1
    
    ' Category - Table Encounters:Category
    Set oZPVChildLevel1 = oDom.createElement("ZPV.5")
    oZPVChildLevel1.Text = NullToNothing(EncountersRS.Fields("Category").Value)
    oZPVNode.appendChild oZPVChildLevel1
    
    ' Primary Diagnosis - Table Encounters:PrimaryDiagnosisNCID
    '<SCR 22798 CC>
    Set oZPVChildLevel1 = oDom.createElement("ZPV.6")
    strTemp = NullToNothing(EncountersRS.Fields("PrimaryDiagnosisNCID").Value)
        Set oZPVChildLevel2 = oDom.createElement("CE.1")
        oZPVChildLevel1.appendChild oZPVChildLevel2
        Set oZPVChildLevel2 = oDom.createElement("CE.2")
        If strTemp <> "" Then
            oZPVChildLevel2.Text = GetNCIDDescription(CLng(strTemp))
        End If
        oZPVChildLevel1.appendChild oZPVChildLevel2
        Set oZPVChildLevel2 = oDom.createElement("CE.3")
        oZPVChildLevel1.appendChild oZPVChildLevel2
        Set oZPVChildLevel2 = oDom.createElement("CE.4")
        oZPVChildLevel1.appendChild oZPVChildLevel2
        Set oZPVChildLevel2 = oDom.createElement("CE.5")
        oZPVChildLevel2.Text = strTemp
        oZPVChildLevel1.appendChild oZPVChildLevel2
        Set oZPVChildLevel2 = oDom.createElement("CE.6")
        oZPVChildLevel2.Text = "99NCI"
        oZPVChildLevel1.appendChild oZPVChildLevel2
    oZPVNode.appendChild oZPVChildLevel1
    
    ' Appointment ID - Table Encounters:ApptID
    Set oZPVChildLevel1 = oDom.createElement("ZPV.7")
    oZPVChildLevel1.Text = NullToNothing(EncountersRS.Fields("ApptID").Value)
    oZPVNode.appendChild oZPVChildLevel1
    
    ' SADR Status - Table Encounters:SADRStatus
    Set oZPVChildLevel1 = oDom.createElement("ZPV.8")
    oZPVChildLevel1.Text = NullToNothing(EncountersRS.Fields("SADRStatus").Value)
    oZPVNode.appendChild oZPVChildLevel1
    
    ' TimeZone - Table Encounters:TimeZone
    Set oZPVChildLevel1 = oDom.createElement("ZPV.9")
    oZPVChildLevel1.Text = NullToNothing(EncountersRS.Fields("TimeZone").Value)
    oZPVNode.appendChild oZPVChildLevel1
    
    ' AP Role1 - Table Encounters:APRole1NCID
    Set oZPVChildLevel1 = oDom.createElement("ZPV.10")
    strTemp = NullToNothing(EncountersRS.Fields("APRole1NCID").Value)
    Set oZPVChildLevel2 = oDom.createElement("CE.1")
        oZPVChildLevel1.appendChild oZPVChildLevel2
        Set oZPVChildLevel2 = oDom.createElement("CE.2")
        If strTemp <> "" Then
            oZPVChildLevel2.Text = GetNCIDDescription(CLng(strTemp))
        End If
        oZPVChildLevel1.appendChild oZPVChildLevel2
        Set oZPVChildLevel2 = oDom.createElement("CE.3")
        oZPVChildLevel1.appendChild oZPVChildLevel2
        Set oZPVChildLevel2 = oDom.createElement("CE.4")
        oZPVChildLevel1.appendChild oZPVChildLevel2
        Set oZPVChildLevel2 = oDom.createElement("CE.5")
        oZPVChildLevel2.Text = strTemp
        oZPVChildLevel1.appendChild oZPVChildLevel2
        Set oZPVChildLevel2 = oDom.createElement("CE.6")
        oZPVChildLevel2.Text = "99NCI"
        oZPVChildLevel1.appendChild oZPVChildLevel2
    oZPVNode.appendChild oZPVChildLevel1
    
    ' AP Role2 - Table Encounters:APRole2NCID
    Set oZPVChildLevel1 = oDom.createElement("ZPV.11")
    strTemp = NullToNothing(EncountersRS.Fields("APRole2NCID").Value)
    Set oZPVChildLevel2 = oDom.createElement("CE.1")
        oZPVChildLevel1.appendChild oZPVChildLevel2
        Set oZPVChildLevel2 = oDom.createElement("CE.2")
        If strTemp <> "" Then
            oZPVChildLevel2.Text = GetNCIDDescription(CLng(strTemp))
        End If
        oZPVChildLevel1.appendChild oZPVChildLevel2
        Set oZPVChildLevel2 = oDom.createElement("CE.3")
        oZPVChildLevel1.appendChild oZPVChildLevel2
        Set oZPVChildLevel2 = oDom.createElement("CE.4")
        oZPVChildLevel1.appendChild oZPVChildLevel2
        Set oZPVChildLevel2 = oDom.createElement("CE.5")
        oZPVChildLevel2.Text = strTemp
        oZPVChildLevel1.appendChild oZPVChildLevel2
        Set oZPVChildLevel2 = oDom.createElement("CE.6")
        oZPVChildLevel2.Text = "99NCI"
        oZPVChildLevel1.appendChild oZPVChildLevel2
    oZPVNode.appendChild oZPVChildLevel1
    
    ' Was Updated - Table Encounters:WasUpdated
    Set oZPVChildLevel1 = oDom.createElement("ZPV.12")
    oZPVChildLevel1.Text = NullToNothing(EncountersRS.Fields("WasUpdated").Value)
    oZPVNode.appendChild oZPVChildLevel1
    
    ' Date/Time Cosigned - Table Encounters:WhenCosigned
    Set oZPVChildLevel1 = oDom.createElement("ZPV.13")
    strTemp = NullToNothing(EncountersRS.Fields("WhenCosigned").Value)
    If IsValidDate(strTemp) Then
        oZPVChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZPVNode.appendChild oZPVChildLevel1
    
    ' Cosigner Line2 - Table Encounters:CosignerLine2
    Set oZPVChildLevel1 = oDom.createElement("ZPV.14")
    oZPVChildLevel1.Text = NullToNothing(EncountersRS.Fields("CosignerLine2").Value)
    oZPVNode.appendChild oZPVChildLevel1
    
    ' Cosigner Line3 - Table Encounters:CosignerLine3
    Set oZPVChildLevel1 = oDom.createElement("ZPV.15")
    oZPVChildLevel1.Text = NullToNothing(EncountersRS.Fields("CosignerLine3").Value)
    oZPVNode.appendChild oZPVChildLevel1
    
    ' Date/Time Signed - Table Encounters:WhenSigned
    Set oZPVChildLevel1 = oDom.createElement("ZPV.16")
    strTemp = NullToNothing(EncountersRS.Fields("WhenSigned").Value)
    If IsValidDate(strTemp) Then
        oZPVChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZPVNode.appendChild oZPVChildLevel1
    
    ' Signer Line2 - Table Encounters:SignerLine2
    Set oZPVChildLevel1 = oDom.createElement("ZPV.17")
    oZPVChildLevel1.Text = NullToNothing(EncountersRS.Fields("SignerLine2").Value)
    oZPVNode.appendChild oZPVChildLevel1
    
    ' Signer Line3 - Table Encounters:SignerLine3
    Set oZPVChildLevel1 = oDom.createElement("ZPV.18")
    oZPVChildLevel1.Text = NullToNothing(EncountersRS.Fields("SignerLine3").Value)
    oZPVNode.appendChild oZPVChildLevel1
    
    ' No Count - Table Encounters:NoCount
    Set oZPVChildLevel1 = oDom.createElement("ZPV.19")
    oZPVChildLevel1.Text = NullToNothing(EncountersRS.Fields("NoCount").Value)
    oZPVNode.appendChild oZPVChildLevel1
    
    ' Additional Provider 1 Information
    Set oZPVChildLevel1 = oDom.createElement("ZPV.20")
        strTemp = NullToNothing(EncountersRS.Fields("ADDITIONALPROVIDER1NCID").Value)
        If strTemp <> "" Then
            Call GetNewDoctorInfo(CLng(strTemp), strDocLastName, strDocFirstName, strDocMiddleName, strDocID, strDocSpecialty)
            blnRsIsGood = True
        End If
        ' Additional Provider 1 ID Number (SSN) - Table Encounters:ADDITIONALPROVIDER1NCID
        Set oZPVChildLevel2 = oDom.createElement("XCN.1")
        If blnRsIsGood = True Then
            oZPVChildLevel2.Text = strDocID
        End If
        oZPVChildLevel1.appendChild oZPVChildLevel2
        ' Additional Provider 1 Family Last Name
        Set oZPVChildLevel2 = oDom.createElement("XCN.2")
        If blnRsIsGood = True Then
            oZPVChildLevel2.Text = strDocLastName
        End If
        oZPVChildLevel1.appendChild oZPVChildLevel2
        ' Additional Provider 1 Family First Name
        Set oZPVChildLevel2 = oDom.createElement("XCN.3")
        If blnRsIsGood = True Then
            oZPVChildLevel2.Text = strDocFirstName
        End If
        oZPVChildLevel1.appendChild oZPVChildLevel2
        ' Additional Provider 1 Family Middle Name
        Set oZPVChildLevel2 = oDom.createElement("XCN.4")
        If blnRsIsGood = True Then
            oZPVChildLevel2.Text = strDocMiddleName
        End If
        oZPVChildLevel1.appendChild oZPVChildLevel2
        ' Additional Provider 1 NCID
        Set oZPVChildLevel2 = oDom.createElement("XCN.5")
        If blnRsIsGood = True Then
            oZPVChildLevel2.Text = strTemp
        End If
        oZPVChildLevel1.appendChild oZPVChildLevel2
    
    oZPVNode.appendChild oZPVChildLevel1
    
    ' Appointment Type - Table Encounters:ApptType
    Set oZPVChildLevel1 = oDom.createElement("ZPV.21")
    oZPVChildLevel1.Text = NullToNothing(EncountersRS.Fields("ApptType").Value)
    oZPVNode.appendChild oZPVChildLevel1
    
    ' Additional Provider 2 Information - Table Encounters:ADDITIONALPROVIDER2NCID
    Set oZPVChildLevel1 = oDom.createElement("ZPV.22")
        strTemp = NullToNothing(EncountersRS.Fields("ADDITIONALPROVIDER2NCID").Value)
        blnRsIsGood = False
        If strTemp <> "" Then
            Call GetNewDoctorInfo(CLng(strTemp), strDocLastName, strDocFirstName, strDocMiddleName, strDocID, strDocSpecialty)
            blnRsIsGood = True
        End If
        ' Additional Provider 2 ID Number (SSN) - Table Encounters:ADDITIONALPROVIDER2NCID
        Set oZPVChildLevel2 = oDom.createElement("XCN.1")
        If blnRsIsGood = True Then
            oZPVChildLevel2.Text = strDocID
        End If
        oZPVChildLevel1.appendChild oZPVChildLevel2
        ' Additional Provider 2 Family Last Name
        Set oZPVChildLevel2 = oDom.createElement("XCN.2")
        If blnRsIsGood = True Then
            oZPVChildLevel2.Text = strDocLastName
        End If
        oZPVChildLevel1.appendChild oZPVChildLevel2
        ' Additional Provider 2 Family First Name
        Set oZPVChildLevel2 = oDom.createElement("XCN.3")
        If blnRsIsGood = True Then
            oZPVChildLevel2.Text = strDocFirstName
        End If
        oZPVChildLevel1.appendChild oZPVChildLevel2
        ' Additional Provider 2 Family Middle Name
        Set oZPVChildLevel2 = oDom.createElement("XCN.4")
        If blnRsIsGood = True Then
            oZPVChildLevel2.Text = strDocMiddleName
        End If
        oZPVChildLevel1.appendChild oZPVChildLevel2
        ' Additional Provider 2 NCID
        Set oZPVChildLevel2 = oDom.createElement("XCN.5")
        If blnRsIsGood = True Then
            oZPVChildLevel2.Text = strTemp
        End If
        oZPVChildLevel1.appendChild oZPVChildLevel2
    oZPVNode.appendChild oZPVChildLevel1
    
    ' Allergy Verified By Date/Time - Table Encounters:AllergyVerifiedByDate
    Set oZPVChildLevel1 = oDom.createElement("ZPV.23")
    strTemp = NullToNothing(EncountersRS.Fields("AllergyVerifiedByDate").Value)
    If IsValidDate(strTemp) Then
        oZPVChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZPVNode.appendChild oZPVChildLevel1
    
    ' Allergy Verified By - Table Encounters:AllergyVerifiedByNCID
    Set oZPVChildLevel1 = oDom.createElement("ZPV.24")
    strTemp = NullToNothing(EncountersRS.Fields("AllergyVerifiedByNCID").Value)
    Set oZPVChildLevel2 = oDom.createElement("CE.1")
        oZPVChildLevel1.appendChild oZPVChildLevel2
        Set oZPVChildLevel2 = oDom.createElement("CE.2")
        If strTemp <> "" Then
            oZPVChildLevel2.Text = GetNCIDDescription(CLng(strTemp))
        End If
        oZPVChildLevel1.appendChild oZPVChildLevel2
        Set oZPVChildLevel2 = oDom.createElement("CE.3")
        oZPVChildLevel1.appendChild oZPVChildLevel2
        Set oZPVChildLevel2 = oDom.createElement("CE.4")
        oZPVChildLevel1.appendChild oZPVChildLevel2
        Set oZPVChildLevel2 = oDom.createElement("CE.5")
        oZPVChildLevel2.Text = strTemp
        oZPVChildLevel1.appendChild oZPVChildLevel2
        Set oZPVChildLevel2 = oDom.createElement("CE.6")
        oZPVChildLevel2.Text = "99NCI"
        oZPVChildLevel1.appendChild oZPVChildLevel2
    oZPVNode.appendChild oZPVChildLevel1
    
    ' Who CoSigned - Table Encounters:CoSignerNCID
    Set oZPVChildLevel1 = oDom.createElement("ZPV.25")
        strTemp = NullToNothing(EncountersRS.Fields("CoSignerNCID").Value)
        If strTemp <> "" Then
            Call GetNewDoctorInfo(CLng(strTemp), strDocLastName, strDocFirstName, strDocMiddleName, strDocID, strDocSpecialty)
            blnRsIsGood = True
        End If
        
        'CoSignerNCID : SSN
        Set oZPVChildLevel2 = oDom.createElement("CE.1")
        If blnRsIsGood = True Then
            oZPVChildLevel2.Text = strDocID
        End If
        oZPVChildLevel1.appendChild oZPVChildLevel2
        
        'CoSignerNCID : Full Name
        oZPVChildLevel1.appendChild oZPVChildLevel2
        Set oZPVChildLevel2 = oDom.createElement("CE.2")
        If strTemp <> "" Then
            oZPVChildLevel2.Text = GetNCIDDescription(CLng(strTemp))
        End If
        oZPVChildLevel1.appendChild oZPVChildLevel2
        
        Set oZPVChildLevel2 = oDom.createElement("CE.3")
        oZPVChildLevel1.appendChild oZPVChildLevel2
        Set oZPVChildLevel2 = oDom.createElement("CE.4")
        oZPVChildLevel1.appendChild oZPVChildLevel2
        
        'CoSignerNCID : NCID
        Set oZPVChildLevel2 = oDom.createElement("CE.5")
        oZPVChildLevel2.Text = strTemp
        oZPVChildLevel1.appendChild oZPVChildLevel2
        
        Set oZPVChildLevel2 = oDom.createElement("CE.6")
        oZPVChildLevel2.Text = "99NCI"
        oZPVChildLevel1.appendChild oZPVChildLevel2
    oZPVNode.appendChild oZPVChildLevel1
    
    ' Who Signed - Table Encounters:WhoSignedNCID
    Set oZPVChildLevel1 = oDom.createElement("ZPV.26")
        strTemp = NullToNothing(EncountersRS.Fields("WhoSignedNCID").Value)
        If strTemp <> "" Then
            Call GetNewDoctorInfo(CLng(strTemp), strDocLastName, strDocFirstName, strDocMiddleName, strDocID, strDocSpecialty)
            blnRsIsGood = True
        End If
        
        'WhoSignedNCID : SSN
        Set oZPVChildLevel2 = oDom.createElement("CE.1")
        If blnRsIsGood = True Then
            oZPVChildLevel2.Text = strDocID
        End If
        oZPVChildLevel1.appendChild oZPVChildLevel2
        
        'WhoSignerNCID : Full Name
        Set oZPVChildLevel2 = oDom.createElement("CE.2")
        If strTemp <> "" Then
            oZPVChildLevel2.Text = GetNCIDDescription(CLng(strTemp))
        End If
        oZPVChildLevel1.appendChild oZPVChildLevel2
        
        Set oZPVChildLevel2 = oDom.createElement("CE.3")
        oZPVChildLevel1.appendChild oZPVChildLevel2
        Set oZPVChildLevel2 = oDom.createElement("CE.4")
        oZPVChildLevel1.appendChild oZPVChildLevel2
        
        'WhoSignedNCID : NCID
        Set oZPVChildLevel2 = oDom.createElement("CE.5")
        oZPVChildLevel2.Text = strTemp
        oZPVChildLevel1.appendChild oZPVChildLevel2
        
        Set oZPVChildLevel2 = oDom.createElement("CE.6")
        oZPVChildLevel2.Text = "99NCI"
        oZPVChildLevel1.appendChild oZPVChildLevel2
    oZPVNode.appendChild oZPVChildLevel1
    
    ' COSigner - Table Encounters:CoSignerNCID
    Set oZPVChildLevel1 = oDom.createElement("ZPV.27")
    strTemp = NullToNothing(EncountersRS.Fields("CoSignerNCID").Value)
    Set oZPVChildLevel2 = oDom.createElement("CE.1")
        oZPVChildLevel1.appendChild oZPVChildLevel2
        Set oZPVChildLevel2 = oDom.createElement("CE.2")
        If strTemp <> "" Then
            oZPVChildLevel2.Text = GetNCIDDescription(CLng(strTemp))
        End If
        oZPVChildLevel1.appendChild oZPVChildLevel2
        Set oZPVChildLevel2 = oDom.createElement("CE.3")
        oZPVChildLevel1.appendChild oZPVChildLevel2
        Set oZPVChildLevel2 = oDom.createElement("CE.4")
        oZPVChildLevel1.appendChild oZPVChildLevel2
        Set oZPVChildLevel2 = oDom.createElement("CE.5")
        oZPVChildLevel2.Text = strTemp
        oZPVChildLevel1.appendChild oZPVChildLevel2
        Set oZPVChildLevel2 = oDom.createElement("CE.6")
        oZPVChildLevel2.Text = "99NCI"
        oZPVChildLevel1.appendChild oZPVChildLevel2
    oZPVNode.appendChild oZPVChildLevel1
    
    ' Created By - Table Encounters:CreatedBy
    Set oZPVChildLevel1 = oDom.createElement("ZPV.28")
    oZPVChildLevel1.Text = NullToNothing(EncountersRS.Fields("CreatedBy").Value)
    oZPVNode.appendChild oZPVChildLevel1
    
    ' Created On - Table Encounters:CreatedOn
    Set oZPVChildLevel1 = oDom.createElement("ZPV.29")
    strTemp = NullToNothing(EncountersRS.Fields("CreatedOn").Value)
    If IsValidDate(strTemp) Then
        oZPVChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZPVNode.appendChild oZPVChildLevel1
    
    ' Updated By - Table Encounters:UpdatedBy
    Set oZPVChildLevel1 = oDom.createElement("ZPV.30")
    oZPVChildLevel1.Text = NullToNothing(EncountersRS.Fields("UpdatedBy").Value)
    oZPVNode.appendChild oZPVChildLevel1
    
    ' Updated On - Table Encounters:UpdatedOn
    Set oZPVChildLevel1 = oDom.createElement("ZPV.31")
    strTemp = NullToNothing(EncountersRS.Fields("UpdatedOn").Value)
    If IsValidDate(strTemp) Then
        oZPVChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZPVNode.appendChild oZPVChildLevel1
    
    ' Encounter Number - Table Enc_Disposition:EncounterNumber
    Set oZPVChildLevel1 = oDom.createElement("ZPV.32")
    '<SCR 2294 CC>
    'strTemp = NullToNothing(EncountersRS.Fields("EncounterNumber").Value)
    'If strTemp <> "" Then
    '    oZPVChildLevel1.Text = GetNCIDDescription(CLng(strTemp))
    'End If
    oZPVChildLevel1.Text = NullToNothing(EncountersRS.Fields("EncounterNumber").Value)
    '</SCR 2294 CC>
    oZPVNode.appendChild oZPVChildLevel1
    
    ' Facility
    Set oZPVChildLevel1 = oDom.createElement("ZPV.33")
        ' Identifier - Table Enc_Disposition:FacilityNCID
        Set oZPVChildLevel2 = oDom.createElement("CE.1")
        oZPVChildLevel2.Text = NullToNothing(EncountersRS.Fields("FacilityNCID").Value)
        oZPVChildLevel1.appendChild oZPVChildLevel2
        ' Text - Table Enc_Disposition:FacilityNCID
        Set oZPVChildLevel2 = oDom.createElement("CE.2")
        strTemp = NullToNothing(EncountersRS.Fields("FacilityNCID").Value)
        If strTemp <> "" Then
            oZPVChildLevel2.Text = GetNCIDDescription(CLng(strTemp))
        End If
        oZPVChildLevel1.appendChild oZPVChildLevel2
    oZPVNode.appendChild oZPVChildLevel1
    
    'Table Enc_Disposition
    blnRsIsGood = False
    If Not (Enc_DispositionRS Is Nothing) Then
        If Enc_DispositionRS.BOF = False And Enc_DispositionRS.EOF = False Then
            blnRsIsGood = True
        End If
    End If
    
    ' NCID - Table Enc_Disposition:NCID
    Set oZPVChildLevel1 = oDom.createElement("ZPV.34")
    If blnRsIsGood = True Then
        oZPVChildLevel1.Text = NullToNothing(Enc_DispositionRS.Fields("NCID").Value)
    End If
    oZPVNode.appendChild oZPVChildLevel1
    
    ' EandM Code Review - Table Enc_Disposition:EandMCodeReview
    Set oZPVChildLevel1 = oDom.createElement("ZPV.35")
    If blnRsIsGood = True Then
        oZPVChildLevel1.Text = NullToNothing(Enc_DispositionRS.Fields("EandMCodeReview").Value)
    End If
    oZPVNode.appendChild oZPVChildLevel1
    
    ' EandM Code Reviewed Date/Time - Table Enc_Disposition:EandMReviewedDate
    Set oZPVChildLevel1 = oDom.createElement("ZPV.36")
    If blnRsIsGood Then
        strTemp = NullToNothing(Enc_DispositionRS.Fields("EandMReviewedDate").Value)
        If IsValidDate(strTemp) Then
            oZPVChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
        End If
    End If
    oZPVNode.appendChild oZPVChildLevel1
    
    ' EandM Code Reviewed By - Table Enc_Disposition:EandMReviewedBy
    Set oZPVChildLevel1 = oDom.createElement("ZPV.37")
    If blnRsIsGood = True Then
        oZPVChildLevel1.Text = NullToNothing(Enc_DispositionRS.Fields("EandMReviewedBy").Value)
    End If
    oZPVNode.appendChild oZPVChildLevel1
    
    ' EandM Code Reviewed - Table Enc_Disposition:EandMReviewed
    Set oZPVChildLevel1 = oDom.createElement("ZPV.38")
    If blnRsIsGood = True Then
        oZPVChildLevel1.Text = NullToNothing(Enc_DispositionRS.Fields("EandMReviewed").Value)
    End If
    oZPVNode.appendChild oZPVChildLevel1
    
    ' Reviewed By Coder - Table Enc_Disposition:ReviewedByCoder
    Set oZPVChildLevel1 = oDom.createElement("ZPV.39")
    If blnRsIsGood = True Then
        oZPVChildLevel1.Text = NullToNothing(Enc_DispositionRS.Fields("ReviewedByCoder").Value)
    End If
    oZPVNode.appendChild oZPVChildLevel1
    
    ' Disposition Admin Option - Table Enc_Disposition:DispAdminOption
    Set oZPVChildLevel1 = oDom.createElement("ZPV.40")
    If blnRsIsGood = True Then
        oZPVChildLevel1.Text = NullToNothing(Enc_DispositionRS.Fields("DispAdminOption").Value)
    End If
    oZPVNode.appendChild oZPVChildLevel1
    
    ' Disposition
    Set oZPVChildLevel1 = oDom.createElement("ZPV.41")
        ' Identifier - Table Enc_Disposition:DispositionNCID
        Set oZPVChildLevel2 = oDom.createElement("CE.1")
        If blnRsIsGood = True Then
            oZPVChildLevel2.Text = NullToNothing(Enc_DispositionRS.Fields("DispositionNCID").Value)
        End If
        oZPVChildLevel1.appendChild oZPVChildLevel2
        ' Text - Table Enc_Disposition:DispositionText
        Set oZPVChildLevel2 = oDom.createElement("CE.2")
        If blnRsIsGood = True Then
            oZPVChildLevel2.Text = NullToNothing(Enc_DispositionRS.Fields("DispositionText").Value)
        End If
        oZPVChildLevel1.appendChild oZPVChildLevel2
    oZPVNode.appendChild oZPVChildLevel1
        
    ' Items Discussed - Table Enc_Disposition:ItemsDiscussed
    Set oZPVChildLevel1 = oDom.createElement("ZPV.42")
    If blnRsIsGood = True Then
        oZPVChildLevel1.Text = NullToNothing(Enc_DispositionRS.Fields("ItemsDiscussed").Value)
    End If
    oZPVNode.appendChild oZPVChildLevel1
    
    ' Follow Up Timeframe - Table Enc_Disposition:FollowUpTimeframe
    Set oZPVChildLevel1 = oDom.createElement("ZPV.43")
    If blnRsIsGood = True Then
        oZPVChildLevel1.Text = NullToNothing(Enc_DispositionRS.Fields("FollowUpTimeframe").Value)
    End If
    oZPVNode.appendChild oZPVChildLevel1
    
    ' Follow Up Comments - Table Enc_Disposition:FollowUUpComments
    Set oZPVChildLevel1 = oDom.createElement("ZPV.44")
    If blnRsIsGood = True Then
        oZPVChildLevel1.Text = NullToNothing(Enc_DispositionRS.Fields("FollowUUpComments").Value)
    End If
    oZPVNode.appendChild oZPVChildLevel1
    
    ' WS Disposition - Table Enc_Disposition:WSDisp
    Set oZPVChildLevel1 = oDom.createElement("ZPV.45")
    If blnRsIsGood = True Then
        oZPVChildLevel1.Text = NullToNothing(Enc_DispositionRS.Fields("WSDisp").Value)
    End If
    oZPVNode.appendChild oZPVChildLevel1
    
    ' EandM Calculation - Table Enc_Disposition:EandMCalc
    Set oZPVChildLevel1 = oDom.createElement("ZPV.46")
    If blnRsIsGood = True Then
        oZPVChildLevel1.Text = NullToNothing(Enc_DispositionRS.Fields("EandMCalc").Value)
    End If
    oZPVNode.appendChild oZPVChildLevel1
    
     '<SCR 22810 CC>
    
    ' Type Case Code
    Set oZPVChildLevel1 = oDom.createElement("ZPV.47")
    If blnRsIsGood = True Then
        oZPVChildLevel1.Text = NullToNothing(Enc_DispositionRS.Fields("INJ_ILL_CAUSE").Value)
    End If
    oZPVNode.appendChild oZPVChildLevel1
    
    ' Disease Non-Battle Injury (DNBI) Category
    Set oZPVChildLevel1 = oDom.createElement("ZPV.48")
    If blnRsIsGood = True Then
        oZPVChildLevel1.Text = NullToNothing(Enc_DispositionRS.Fields("INJ_ILL_CATEGORY").Value)
    End If
    oZPVNode.appendChild oZPVChildLevel1
    
    Set oZPVChildLevel1 = oDom.createElement("ZPV.49")
    If blnRsIsGood = True Then
        oZPVChildLevel1.Text = NullToNothing(Enc_DispositionRS.Fields("DISCUSSEDCOMMENT").Value)
    End If
    oZPVNode.appendChild oZPVChildLevel1
    
    '</SCR 22810 CC>

    Set oZPVChildLevel1 = oDom.createElement("ZPV.50")
    oZPVChildLevel1.Text = NullToNothing(EncountersRS.Fields("EncounterNumber").Value)
    oZPVNode.appendChild oZPVChildLevel1
    
    Set oZPVChildLevel1 = oDom.createElement("ZPV.51")
    oZPVChildLevel1.Text = NullToNothing(EncountersRS.Fields("FACILITYNCID").Value)
    oZPVNode.appendChild oZPVChildLevel1
    
    Set oZPVChildLevel1 = oDom.createElement("ZPV.52")
    oZPVChildLevel1.Text = NullToNothing(EncountersRS.Fields("PRIMARYPROVIDERNCID").Value)
    oZPVNode.appendChild oZPVChildLevel1
    
    
    Set oZPVChildLevel1 = oDom.createElement("ZPV.53")
    oZPVChildLevel1.Text = NullToNothing(EncountersRS.Fields("CLINICNCID").Value)
    oZPVNode.appendChild oZPVChildLevel1
    
    Set oZPVChildLevel1 = oDom.createElement("ZPV.54")
    oZPVChildLevel1.Text = NullToNothing(EncountersRS.Fields("TYPE").Value)
    oZPVNode.appendChild oZPVChildLevel1
    
    Set oZPVChildLevel1 = oDom.createElement("ZPV.55")
    strTemp = NullToNothing(EncountersRS.Fields("STARTDTS").Value)
    If IsValidDate(strTemp) Then
        oZPVChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZPVNode.appendChild oZPVChildLevel1
    
    Set oZPVChildLevel1 = oDom.createElement("ZPV.56")
    strTemp = NullToNothing(EncountersRS.Fields("ENDDTS").Value)
    If IsValidDate(strTemp) Then
        oZPVChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZPVNode.appendChild oZPVChildLevel1
    
    ' Begin:SCR #26544;   Developer: MAG 08/19/2002 10:01 AM
    ' Category - Table Encounters:DxSnoID
    Set oZPVChildLevel1 = oDom.createElement("ZPV.57")
    oZPVChildLevel1.Text = NullToNothing(EncountersRS.Fields("DXSnoID").Value)
    oZPVNode.appendChild oZPVChildLevel1
    ' End: SCR #26544;
    
    Set oZPVChildLevel1 = oDom.createElement("ZPV.58")
    strTemp = NullToNothing(EncountersRS.Fields("PRIMARYPROVIDERNCID").Value)
    If strTemp <> "" Then
        Call GetNewDoctorInfo(CLng(strTemp), strDocLastName, strDocFirstName, strDocMiddleName, strDocID, strDocSpecialty)
        blnRsIsGood = True
    End If
    ' Primary Provider ID Number (SSN) - Table Encounters:PRIMARYPROVIDERNCID
    Set oZPVChildLevel2 = oDom.createElement("XCN.1")
    If blnRsIsGood = True Then
        oZPVChildLevel2.Text = strDocID
    End If
    oZPVChildLevel1.appendChild oZPVChildLevel2
    ' Primary Provider Last Name
    Set oZPVChildLevel2 = oDom.createElement("XCN.2")
    If blnRsIsGood = True Then
        oZPVChildLevel2.Text = strDocLastName
    End If
    oZPVChildLevel1.appendChild oZPVChildLevel2
    ' Primary Provider First Name
    Set oZPVChildLevel2 = oDom.createElement("XCN.3")
    If blnRsIsGood = True Then
        oZPVChildLevel2.Text = strDocFirstName
    End If
    oZPVChildLevel1.appendChild oZPVChildLevel2
    ' Primary Provider Middle Name
    Set oZPVChildLevel2 = oDom.createElement("XCN.4")
    If blnRsIsGood = True Then
        oZPVChildLevel2.Text = strDocMiddleName
    End If
    oZPVChildLevel1.appendChild oZPVChildLevel2
    ' Primary Provider NCID
    Set oZPVChildLevel2 = oDom.createElement("XCN.5")
    If blnRsIsGood = True Then
        oZPVChildLevel2.Text = strTemp
    End If
    oZPVChildLevel1.appendChild oZPVChildLevel2
    oZPVNode.appendChild oZPVChildLevel1
    
    Set ReturnNewZPVNode = oZPVNode
    
Exit Function

ErrTrap:
   If gobjShared Is Nothing Then
      Set gobjShared = New CWShared
   End If
   gobjShared.ShowVBError Err.Number, Err.Description, "ReturnNewZPVNode Function", "DMBE", vbExclamation
   Resume Next 'There is no reason this procedure should not continue to create xml if it errors
End Function

Public Function ReturnNewZWSNode(ByVal lngCount As Long, ByVal Work_StatusRS As ADODB.Recordset) As MSXML.IXMLDOMElement
' Work Status Segment
' sample function output
'   <ZWS>
'      <ZWS.1>Count</ZWS.1>
'      <ZWS.2>Encounter Number</ZWS.2>
'      <ZWS.3>Facility NCID</ZWS.3>
'      <ZWS.4>Patient NCID</ZWS.4>
'      <ZWS.5>
'         <XCN.1>Provider SSN</XCN.1>
'         <XCN.2>Provider Last Name</XCN.2>
'         <XCN.3>Provider First Name</XCN.3>
'         <XCN.4>Provider Middle Name</XCN.4>
'      </ZWS.5>
'      <ZWS.6>Work Status</ZWS.6>
'      <ZWS.7>Work Status Selected</ZWS.7>
'      <ZWS.8>Work Status Qualified</ZWS.8>
'      <ZWS.9>Last Updated</ZWS.9>
'   </ZWS>
   
    Dim oDom As MSXML.DOMDocument
    Dim oZWSNode As MSXML.IXMLDOMElement
    Dim oZWSChildLevel1 As MSXML.IXMLDOMElement
    Dim oZWSChildLevel2 As MSXML.IXMLDOMElement
    Dim strDocFirstName As String
    Dim strDocLastName As String
    Dim strDocMiddleName As String
    Dim strDocID As String
    Dim strDocSpecialty As String
    Dim strTemp As String
    Dim blnRsIsGood As Boolean
    
    On Error GoTo ErrTrap

    Set oDom = New MSXML.DOMDocument
    Set oZWSNode = oDom.createElement("ZWS")
    
    ' Set ID - Count
    Set oZWSChildLevel1 = oDom.createElement("ZWS.1")
    oZWSChildLevel1.Text = CStr(lngCount)
    oZWSNode.appendChild oZWSChildLevel1

    ' Encounter Number - Table Work_Status:Encounter_Number
    Set oZWSChildLevel1 = oDom.createElement("ZWS.2")
    oZWSChildLevel1.Text = NullToNothing(Work_StatusRS.Fields("Encounter_Number").Value)
    oZWSNode.appendChild oZWSChildLevel1

    ' Facility - Table Work_Status:Facility_ID
    Set oZWSChildLevel1 = oDom.createElement("ZWS.3")
    oZWSChildLevel1.Text = NullToNothing(Work_StatusRS.Fields("Facility_ID").Value)
    oZWSNode.appendChild oZWSChildLevel1
    
    ' Patient NCID - Table Work_Status:Patient_NCID
    Set oZWSChildLevel1 = oDom.createElement("ZWS.4")
    oZWSChildLevel1.Text = NullToNothing(Work_StatusRS.Fields("Patient_NCID").Value)
    oZWSNode.appendChild oZWSChildLevel1
    
    ' Provider
    Set oZWSChildLevel1 = oDom.createElement("ZWS.5")
        strTemp = NullToNothing(Work_StatusRS.Fields("Provider_NCID").Value)
        If strTemp <> "" Then
            Call GetNewDoctorInfo(CLng(strTemp), strDocLastName, strDocFirstName, strDocMiddleName, strDocID, strDocSpecialty)
            blnRsIsGood = True
        End If
        ' Provider SSN - Table Work_Status:Provider_NCID
        Set oZWSChildLevel2 = oDom.createElement("XCN.1")
        If blnRsIsGood = True Then
            oZWSChildLevel2.Text = strDocID
        End If
        oZWSChildLevel1.appendChild oZWSChildLevel2
        ' Provider Last Name
        Set oZWSChildLevel2 = oDom.createElement("XCN.2")
        If blnRsIsGood = True Then
            oZWSChildLevel2.Text = strDocLastName
        End If
        oZWSChildLevel1.appendChild oZWSChildLevel2
        ' Provider First Name
        Set oZWSChildLevel2 = oDom.createElement("XCN.3")
        If blnRsIsGood = True Then
            oZWSChildLevel2.Text = strDocFirstName
        End If
        oZWSChildLevel1.appendChild oZWSChildLevel2
        ' Provider Middle Name
        Set oZWSChildLevel2 = oDom.createElement("XCN.4")
        If blnRsIsGood = True Then
            oZWSChildLevel2.Text = strDocMiddleName
        End If
        oZWSChildLevel1.appendChild oZWSChildLevel2
    oZWSNode.appendChild oZWSChildLevel1
    
    ' Work Status - Table Work_Status:Work_Status_NCID
    Set oZWSChildLevel1 = oDom.createElement("ZWS.6")
    'oZWSChildLevel1.Text = NullToNothing(Work_StatusRS.Fields("Work_Status_NCID").Value)
        strTemp = NullToNothing(Work_StatusRS.Fields("Work_Status_NCID").Value)
        
        Set oZWSChildLevel2 = oDom.createElement("CE.1")
        oZWSChildLevel1.appendChild oZWSChildLevel2
        
        Set oZWSChildLevel2 = oDom.createElement("CE.2")
        If strTemp <> "" Then
            oZWSChildLevel2.Text = GetNCIDDescription(CLng(strTemp))
        End If
        oZWSChildLevel1.appendChild oZWSChildLevel2
        
        Set oZWSChildLevel2 = oDom.createElement("CE.3")
        oZWSChildLevel1.appendChild oZWSChildLevel2
        Set oZWSChildLevel2 = oDom.createElement("CE.4")
        oZWSChildLevel1.appendChild oZWSChildLevel2
        Set oZWSChildLevel2 = oDom.createElement("CE.5")
        oZWSChildLevel2.Text = strTemp 'The NCID
        oZWSChildLevel1.appendChild oZWSChildLevel2
        Set oZWSChildLevel2 = oDom.createElement("CE.6")
        oZWSChildLevel2.Text = "99NCI"
        oZWSChildLevel1.appendChild oZWSChildLevel2
    oZWSNode.appendChild oZWSChildLevel1
    
    ' Work Status Selected - Table Work_Status:Work_Status_Selected
    Set oZWSChildLevel1 = oDom.createElement("ZWS.7")
    oZWSChildLevel1.Text = NullToNothing(Work_StatusRS.Fields("Work_Status_Selected").Value)
    oZWSNode.appendChild oZWSChildLevel1
    
    ' Work Status Qualified - Table Work_Status:Work_Status_Qualified
    Set oZWSChildLevel1 = oDom.createElement("ZWS.8")
    oZWSChildLevel1.Text = NullToNothing(Work_StatusRS.Fields("Work_Status_Qualified").Value)
    oZWSNode.appendChild oZWSChildLevel1
    
    ' Last Updated - Table Work_Status:Last_Updated
    Set oZWSChildLevel1 = oDom.createElement("ZWS.9")
    strTemp = NullToNothing(Work_StatusRS.Fields("Work_Status_Qualified").Value)
    If IsValidDate(strTemp) Then
        oZWSChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZWSNode.appendChild oZWSChildLevel1
    
    Set ReturnNewZWSNode = oZWSNode
    
Exit Function

ErrTrap:
   If gobjShared Is Nothing Then
      Set gobjShared = New CWShared
   End If
   gobjShared.ShowVBError Err.Number, Err.Description, "ReturnNewZWSNode Function", "DMBE", vbExclamation
   Resume Next 'There is no reason this procedure should not continue to create xml if it errors
End Function

Public Function ReturnNewZSONode(ByVal lngCount As Long, ByVal SoNotesRS As ADODB.Recordset) As MSXML.IXMLDOMElement
' SO Notes Segment
' sample function output
'   <ZSO>
'      <ZSO.1>Count</ZSO.1>
'      <ZSO.2>DataID</ZSO.2>
'      <ZSO.3>Encounter Number</ZSO.3>
'      <ZSO.4>Facility NCID</ZSO.4>
'      <ZSO.5>Type</ZSO.5>
'      <ZSO.6>Encounter Data ID</ZSO.6>
'      <ZSO.7>Date/Time Stamp</ZSO.7>
'      <ZSO.8>
'         <CE.1>Owner Identifier</CE.1>
'         <CE.2>Owner Name</CE.2>
'      </ZSO.8>
'      <ZSO.9>Title</ZSO.9>
'      <ZSO.10>RTF</ZSO.10>
'      <ZSO.11>Update Flag</ZSO.11>
'      <ZSO.12>Profile NCID</ZSO.12>
'      <ZSO.13>Updated By</ZSO.13>
'      <ZSO.14>Updated On</ZSO.14>
'      <ZSO.15>Created By</ZSO.15>
'      <ZSO.16>Created On</ZSO.16>
'   </ZSO>

    Dim oDom As MSXML.DOMDocument
    Dim oZSONode As MSXML.IXMLDOMElement
    Dim oZSOChildLevel1 As MSXML.IXMLDOMElement
    Dim oZSOChildLevel2 As MSXML.IXMLDOMElement
    Dim strTemp As String
    
    On Error GoTo ErrTrap

    Set oDom = New MSXML.DOMDocument
    Set oZSONode = oDom.createElement("ZSO")
    
    ' Set ID - Count
    Set oZSOChildLevel1 = oDom.createElement("ZSO.1")
    oZSOChildLevel1.Text = CStr(lngCount)
    oZSONode.appendChild oZSOChildLevel1

    ' DataID - Table SONotes:DataID
    Set oZSOChildLevel1 = oDom.createElement("ZSO.2")
    oZSOChildLevel1.Text = NullToNothing(SoNotesRS.Fields("DataID").Value)
    oZSONode.appendChild oZSOChildLevel1

    ' Encounter Number - Table SONotes:EncounterNumber
    Set oZSOChildLevel1 = oDom.createElement("ZSO.3")
    oZSOChildLevel1.Text = NullToNothing(SoNotesRS.Fields("EncounterNumber").Value)
    oZSONode.appendChild oZSOChildLevel1

    ' Facility - Table SONotes:FacilityNCID
    Set oZSOChildLevel1 = oDom.createElement("ZSO.4")
    oZSOChildLevel1.Text = NullToNothing(SoNotesRS.Fields("FacilityNCID").Value)
    oZSONode.appendChild oZSOChildLevel1

    ' Type - Table SONotes:Type
    Set oZSOChildLevel1 = oDom.createElement("ZSO.5")
    oZSOChildLevel1.Text = NullToNothing(SoNotesRS.Fields("Type").Value)
    oZSONode.appendChild oZSOChildLevel1

    ' Encounter Data ID - Table SONotes:EncounterDataID
    Set oZSOChildLevel1 = oDom.createElement("ZSO.6")
    oZSOChildLevel1.Text = NullToNothing(SoNotesRS.Fields("EncounterDataID").Value)
    oZSONode.appendChild oZSOChildLevel1

    ' Date/Time Stamp - Table SONotes:DTS
    Set oZSOChildLevel1 = oDom.createElement("ZSO.7")
    strTemp = NullToNothing(SoNotesRS.Fields("DTS").Value)
    If IsValidDate(strTemp) Then
        oZSOChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZSONode.appendChild oZSOChildLevel1

    ' Owner
    Set oZSOChildLevel1 = oDom.createElement("ZSO.8")
        ' Identifier - Table SONotes:OwnerNCID
        Set oZSOChildLevel2 = oDom.createElement("CE.1")
        oZSOChildLevel2.Text = NullToNothing(SoNotesRS.Fields("OwnerNCID").Value)
        oZSOChildLevel1.appendChild oZSOChildLevel2
        ' Name - Table SONotes:OwnerName
        Set oZSOChildLevel2 = oDom.createElement("CE.2")
        oZSOChildLevel2.Text = NullToNothing(SoNotesRS.Fields("OwnerName").Value)
        oZSOChildLevel1.appendChild oZSOChildLevel2
    oZSONode.appendChild oZSOChildLevel1

    ' Title - Table SONotes:Title
    Set oZSOChildLevel1 = oDom.createElement("ZSO.9")
    oZSOChildLevel1.Text = NullToNothing(SoNotesRS.Fields("Title").Value)
    oZSONode.appendChild oZSOChildLevel1

    ' RTF - Table SONotes:RTF
    Set oZSOChildLevel1 = oDom.createElement("ZSO.10")
    strTemp = NullToNothing(SoNotesRS.Fields("RTF").Value)
    If strTemp <> "" Then
        oZSOChildLevel1.Text = FilterRtfToXml(strTemp)
    End If
    oZSONode.appendChild oZSOChildLevel1

    ' Update Flag - Table SONotes:UpdateFlag
    Set oZSOChildLevel1 = oDom.createElement("ZSO.11")
    oZSOChildLevel1.Text = NullToNothing(SoNotesRS.Fields("UpdateFlag").Value)
    oZSONode.appendChild oZSOChildLevel1
    
    ' Profile - Table SONotes:ProfileNCID
    Set oZSOChildLevel1 = oDom.createElement("ZSO.12")
    oZSOChildLevel1.Text = NullToNothing(SoNotesRS.Fields("ProfileNCID").Value)
    oZSONode.appendChild oZSOChildLevel1
    
    ' Updated By - Table SONotes:UpdatedBy
    Set oZSOChildLevel1 = oDom.createElement("ZSO.13")
    oZSOChildLevel1.Text = NullToNothing(SoNotesRS.Fields("UpdatedBy").Value)
    oZSONode.appendChild oZSOChildLevel1
    
    ' Updated On - Table SONotes:UpdatedOn
    Set oZSOChildLevel1 = oDom.createElement("ZSO.14")
    strTemp = NullToNothing(SoNotesRS.Fields("UpdatedOn").Value)
    If IsValidDate(strTemp) Then
        oZSOChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZSONode.appendChild oZSOChildLevel1
    
    ' Created By - Table SONotes:CreatedBy
    Set oZSOChildLevel1 = oDom.createElement("ZSO.15")
    oZSOChildLevel1.Text = NullToNothing(SoNotesRS.Fields("CreatedBy").Value)
    oZSONode.appendChild oZSOChildLevel1
    
    ' Created On - Table SONotes:CreatedOn
    Set oZSOChildLevel1 = oDom.createElement("ZSO.16")
    strTemp = NullToNothing(SoNotesRS.Fields("CreatedOn").Value)
    If IsValidDate(strTemp) Then
        oZSOChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZSONode.appendChild oZSOChildLevel1
    
    Set ReturnNewZSONode = oZSONode
    
Exit Function

ErrTrap:
   If gobjShared Is Nothing Then
      Set gobjShared = New CWShared
   End If
   gobjShared.ShowVBError Err.Number, Err.Description, "ReturnNewZSONode Function", "DMBE", vbExclamation
   Resume Next 'There is no reason this procedure should not continue to create xml if it errors
End Function

Public Function ReturnNewZCNNode(ByVal lngCount As Long, ByVal DMBEPatient As MMMHISPatient.Patient, ByVal ClinNoteRS As ADODB.Recordset) As MSXML.IXMLDOMElement
' Clinician Notes Segment
' sample function output
'   <ZCN>
'      <ZCN.1>Count</ZCN.1>
'      <ZCN.2>Unit Number</ZCN.2>
'      <ZCN.3>Encounter Number</ZCN.3>
'      <ZCN.4>DataID</ZCN.4>
'      <ZCN.5>Application</ZCN.5>
'      <ZCN.6>
'         <CE.1>Note Type Identifier</CE.1>
'         <CE.2>Note Type Name</CE.2>
'      </ZCN.6>
'      <ZCN.7>Noted Date/Time</ZCN.7>
'      <ZCN.8>Point of Care Facility NCID</ZCN.8>
'      <ZCN.9>Encounter Facility NCID</ZCN.9>
'      <ZCN.10>
'         <XCN.1>Clinician SSN</XCN.1>
'         <XCN.2>Clinician Last Name</XCN.2>
'         <XCN.3>Clinician First Name</XCN.3>
'         <XCN.4>Clinician Middle Name</XCN.4>
'      </ZCN.10>
'      <ZCN.11>HTML Text</ZCN.11>
'      <ZCN.12>Caller</ZCN.12>
'      <ZCN.13>Status</ZCN.13>
'      <ZCN.14>Created By</ZCN.14>
'      <ZCN.15>Created Date/Time</ZCN.15>
'      <ZCN.16>Modify By</ZCN.16>
'      <ZCN.17>Modify Date/Time</ZCN.17>
'      <ZCN.18>Priority</ZCN.18>
'      <ZCN.19>Image Available</ZCN.19>
'      <ZCN.20>Image Number</ZCN.20>
'      <ZCN.21>CDR Data ID</ZCN.21>
'   </ZCN>

    Dim oDom As MSXML.DOMDocument
    Dim oZCNNode As MSXML.IXMLDOMElement
    Dim oZCNChildLevel1 As MSXML.IXMLDOMElement
    Dim oZCNChildLevel2 As MSXML.IXMLDOMElement
    Dim strDocFirstName As String
    Dim strDocLastName As String
    Dim strDocMiddleName As String
    Dim strDocID As String
    Dim strDocSpecialty As String
    Dim strTemp As String
    Dim blnRsIsGood As Boolean
    
    On Error GoTo ErrTrap

    Set oDom = New MSXML.DOMDocument
    Set oZCNNode = oDom.createElement("ZCN")
    
    ' Set ID - Count
    Set oZCNChildLevel1 = oDom.createElement("ZCN.1")
    oZCNChildLevel1.Text = CStr(lngCount)
    oZCNNode.appendChild oZCNChildLevel1

    ' Unit Number - Table ClinNote:Unit_Number
    Set oZCNChildLevel1 = oDom.createElement("ZCN.2")
    oZCNChildLevel1.Text = NullToNothing(DMBEPatient.unitNumber)
    oZCNNode.appendChild oZCNChildLevel1

    ' Encounter Number - Table ClinNote:Encounter
    Set oZCNChildLevel1 = oDom.createElement("ZCN.3")
    oZCNChildLevel1.Text = NullToNothing(ClinNoteRS.Fields("Encounter").Value)
    oZCNNode.appendChild oZCNChildLevel1
    
    ' DataID - Table ClinNote:DataID
    Set oZCNChildLevel1 = oDom.createElement("ZCN.4")
    oZCNChildLevel1.Text = NullToNothing(ClinNoteRS.Fields("DataID").Value)
    oZCNNode.appendChild oZCNChildLevel1
    
    ' Application - Table ClinNote:ApplicationNCID
    Set oZCNChildLevel1 = oDom.createElement("ZCN.5")
    oZCNChildLevel1.Text = NullToNothing(ClinNoteRS.Fields("ApplicationNCID").Value)
    oZCNNode.appendChild oZCNChildLevel1
    
    ' Note Type
    Set oZCNChildLevel1 = oDom.createElement("ZCN.6")
        ' Identifier - Table ClinNote:NoteTypeNCID
        Set oZCNChildLevel2 = oDom.createElement("CE.1")
        oZCNChildLevel2.Text = NullToNothing(ClinNoteRS.Fields("NoteTypeNCID").Value)
        oZCNChildLevel1.appendChild oZCNChildLevel2
        ' Name - Table ClinNote:NoteTypeName
        Set oZCNChildLevel2 = oDom.createElement("CE.2")
        oZCNChildLevel2.Text = NullToNothing(ClinNoteRS.Fields("NoteTypeName").Value)
        oZCNChildLevel1.appendChild oZCNChildLevel2
    oZCNNode.appendChild oZCNChildLevel1
    
    ' Noted Date/Time - Table ClinNote:NoteDate
    Set oZCNChildLevel1 = oDom.createElement("ZCN.7")
    strTemp = NullToNothing(ClinNoteRS.Fields("ApplicationNCID").Value)
    If IsValidDate(strTemp) Then
        oZCNChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZCNNode.appendChild oZCNChildLevel1
    
    ' Point of Care Facility - Table ClinNote:POCFacility
    Set oZCNChildLevel1 = oDom.createElement("ZCN.8")
    oZCNChildLevel1.Text = NullToNothing(ClinNoteRS.Fields("POCFacility").Value)
    oZCNNode.appendChild oZCNChildLevel1
    
    ' Encounter Facility - Table ClinNote:EncFacility
    Set oZCNChildLevel1 = oDom.createElement("ZCN.9")
    oZCNChildLevel1.Text = NullToNothing(ClinNoteRS.Fields("EncFacility").Value)
    oZCNNode.appendChild oZCNChildLevel1
    
    ' Clinician
    Set oZCNChildLevel1 = oDom.createElement("ZCN.10")
        strTemp = NullToNothing(ClinNoteRS.Fields("ClinicianNCID").Value)
        If strTemp <> "" Then
            Call GetNewDoctorInfo(CLng(strTemp), strDocLastName, strDocFirstName, strDocMiddleName, strDocID, strDocSpecialty)
            blnRsIsGood = True
        End If
        ' Clinician SSN - Table ClinNote:Clinician
        Set oZCNChildLevel2 = oDom.createElement("XCN.1")
        If blnRsIsGood = True Then
            oZCNChildLevel2.Text = strDocID
        End If
        oZCNChildLevel1.appendChild oZCNChildLevel2
        ' Clinician Last Name
        Set oZCNChildLevel2 = oDom.createElement("XCN.2")
        If blnRsIsGood = True Then
            oZCNChildLevel2.Text = strDocLastName
        End If
        oZCNChildLevel1.appendChild oZCNChildLevel2
        ' Clinician First Name
        Set oZCNChildLevel2 = oDom.createElement("XCN.3")
        If blnRsIsGood = True Then
            oZCNChildLevel2.Text = strDocFirstName
        End If
        oZCNChildLevel1.appendChild oZCNChildLevel2
        ' Clinician Middle Name
        Set oZCNChildLevel2 = oDom.createElement("XCN.4")
        If blnRsIsGood = True Then
            oZCNChildLevel2.Text = strDocMiddleName
        End If
        oZCNChildLevel1.appendChild oZCNChildLevel2
    oZCNNode.appendChild oZCNChildLevel1

    ' HTML Text - Table ClinNote:HTMLText
    Set oZCNChildLevel1 = oDom.createElement("ZCN.11")
    strTemp = NullToNothing(ClinNoteRS.Fields("HTMLText").Value)
    If strTemp <> "" Then
        oZCNChildLevel1.Text = FilterRtfToXml(strTemp)
    End If
    oZCNNode.appendChild oZCNChildLevel1
    
    ' Caller - Table ClinNote:Caller
    Set oZCNChildLevel1 = oDom.createElement("ZCN.12")
    oZCNChildLevel1.Text = NullToNothing(ClinNoteRS.Fields("Caller").Value)
    oZCNNode.appendChild oZCNChildLevel1
    
    ' Status - Table ClinNote:Status
    Set oZCNChildLevel1 = oDom.createElement("ZCN.13")
    oZCNChildLevel1.Text = NullToNothing(ClinNoteRS.Fields("Status").Value)
    oZCNNode.appendChild oZCNChildLevel1
    
    ' Created By - Table ClinNote:CreatedBy
    Set oZCNChildLevel1 = oDom.createElement("ZCN.14")
    oZCNChildLevel1.Text = NullToNothing(ClinNoteRS.Fields("CreatedBy").Value)
    oZCNNode.appendChild oZCNChildLevel1
    
    ' Create Date/Time - Table ClinNote:CreateDate
    Set oZCNChildLevel1 = oDom.createElement("ZCN.15")
    strTemp = NullToNothing(ClinNoteRS.Fields("CreateDate").Value)
    If IsValidDate(strTemp) Then
        oZCNChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZCNNode.appendChild oZCNChildLevel1
    
    ' Modified By - Table ClinNote:ModifiedBy
    Set oZCNChildLevel1 = oDom.createElement("ZCN.16")
    oZCNChildLevel1.Text = NullToNothing(ClinNoteRS.Fields("ModifiedBy").Value)
    oZCNNode.appendChild oZCNChildLevel1
    
    ' Modified Date/Time - Table ClinNote:ModifiedDate
    Set oZCNChildLevel1 = oDom.createElement("ZCN.17")
    strTemp = NullToNothing(ClinNoteRS.Fields("ModifiedDate").Value)
    If IsValidDate(strTemp) Then
        oZCNChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZCNNode.appendChild oZCNChildLevel1
    
    ' Priority - Table ClinNote:Priority
    Set oZCNChildLevel1 = oDom.createElement("ZCN.18")
    oZCNChildLevel1.Text = NullToNothing(ClinNoteRS.Fields("Priority").Value)
    oZCNNode.appendChild oZCNChildLevel1
    
    ' Image Available - Table ClinNote:ImageAvailable
    Set oZCNChildLevel1 = oDom.createElement("ZCN.19")
    oZCNChildLevel1.Text = NullToNothing(ClinNoteRS.Fields("ImageAvailable").Value)
    oZCNNode.appendChild oZCNChildLevel1
    
    ' Image Number - Table ClinNote:ImageNumber
    Set oZCNChildLevel1 = oDom.createElement("ZCN.20")
    oZCNChildLevel1.Text = NullToNothing(ClinNoteRS.Fields("ImageNumber").Value)
    oZCNNode.appendChild oZCNChildLevel1
    
    ' CDR Data ID - Table ClinNote:CDR_DataID
    Set oZCNChildLevel1 = oDom.createElement("ZCN.21")
    oZCNChildLevel1.Text = NullToNothing(ClinNoteRS.Fields("CDR_DataID").Value)
    oZCNNode.appendChild oZCNChildLevel1
    
    Set ReturnNewZCNNode = oZCNNode

Exit Function

ErrTrap:
   If gobjShared Is Nothing Then
      Set gobjShared = New CWShared
   End If
   gobjShared.ShowVBError Err.Number, Err.Description, "ReturnNewZCNNode Function", "DMBE", vbExclamation
   'Err.Clear
   Resume Next 'There is no reason this procedure should not continue to create xml if it errors

End Function

Public Function ReturnNewZBRNode(ByVal lngCount As Long, ByVal ProblemsRelationshipsRS As ADODB.Recordset) As MSXML.IXMLDOMElement
' Problem Relationship Segment
' sample function output
'   <ZBR>
'      <ZBR.1>Count</ZBR.1>
'      <ZBR.2>Problems Data ID</ZBR.2>
'      <ZBR.3>Related Data ID</ZBR.3>
'      <ZBR.4>Relationship NCID</ZBR.4>
'      <ZBR.5>Created By</ZBR.5>
'      <ZBR.6>Created On</ZBR.6>
'      <ZBR.7>Updated By</ZBR.7>
'      <ZBR.8>Updated On</ZBR.8>
'   </ZBR>

    Dim oDom As MSXML.DOMDocument
    Dim oZBRNode As MSXML.IXMLDOMElement
    Dim oZBRChildLevel1 As MSXML.IXMLDOMElement
    Dim strTemp As String
    
    On Error GoTo ErrTrap

    Set oDom = New MSXML.DOMDocument
    Set oZBRNode = oDom.createElement("ZBR")
    
    ' Set ID - Count
    Set oZBRChildLevel1 = oDom.createElement("ZBR.1")
    oZBRChildLevel1.Text = CStr(lngCount)
    oZBRNode.appendChild oZBRChildLevel1

    ' Problems Data ID - Table ProblemsRelationships:ProblemDataID
    Set oZBRChildLevel1 = oDom.createElement("ZBR.2")
    oZBRChildLevel1.Text = NullToNothing(ProblemsRelationshipsRS.Fields("ProblemDataID").Value)
    oZBRNode.appendChild oZBRChildLevel1

    ' Related Data ID - Table ProblemsRelationships:RelatedDataID
    Set oZBRChildLevel1 = oDom.createElement("ZBR.3")
    oZBRChildLevel1.Text = NullToNothing(ProblemsRelationshipsRS.Fields("RelatedDataID").Value)
    oZBRNode.appendChild oZBRChildLevel1

    ' Relationship NCID - Table ProblemsRelationships:RelationshipNCID
    Set oZBRChildLevel1 = oDom.createElement("ZBR.4")
    oZBRChildLevel1.Text = NullToNothing(ProblemsRelationshipsRS.Fields("RelationshipNCID").Value)
    oZBRNode.appendChild oZBRChildLevel1

    ' Created By - Table ProblemsRelationships:CreatedBy
    Set oZBRChildLevel1 = oDom.createElement("ZBR.5")
    oZBRChildLevel1.Text = NullToNothing(ProblemsRelationshipsRS.Fields("CreatedBy").Value)
    oZBRNode.appendChild oZBRChildLevel1
    
    ' Created On - Table ProblemsRelationships:CreatedOn
    Set oZBRChildLevel1 = oDom.createElement("ZBR.6")
    strTemp = NullToNothing(ProblemsRelationshipsRS.Fields("CreatedOn").Value)
    If IsValidDate(strTemp) Then
        oZBRChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZBRNode.appendChild oZBRChildLevel1
    
    ' Updated By - Table ProblemsRelationships:UpdatedBy
    Set oZBRChildLevel1 = oDom.createElement("ZBR.7")
    oZBRChildLevel1.Text = NullToNothing(ProblemsRelationshipsRS.Fields("UpdatedBy").Value)
    oZBRNode.appendChild oZBRChildLevel1
    
    ' Updated On - Table ProblemsRelationships:UpdatedOn
    Set oZBRChildLevel1 = oDom.createElement("ZBR.8")
    strTemp = NullToNothing(ProblemsRelationshipsRS.Fields("UpdatedOn").Value)
    If IsValidDate(strTemp) Then
        oZBRChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZBRNode.appendChild oZBRChildLevel1
    
    Set ReturnNewZBRNode = oZBRNode

Exit Function

ErrTrap:
   If gobjShared Is Nothing Then
      Set gobjShared = New CWShared
   End If
   gobjShared.ShowVBError Err.Number, Err.Description, "ReturnNewZBRNode Function", "DMBE", vbExclamation
   'Err.Clear
   Resume Next 'There is no reason this procedure should not continue to create xml if it errors

End Function

Public Function ReturnNewZCPNode(ByVal lngCount As Long, ByVal ComplaintsRS As ADODB.Recordset) As MSXML.IXMLDOMElement
' Complaint Segment
' sample function output
'   <ZCP>
'      <ZCP.1>Count</ZCP.1>
'      <ZCP.2>DataID</ZCP.2>
'      <ZCP.3>Facility NCID</ZCP.3>
'      <ZCP.4>Unit Number</ZCP.4>
'      <ZCP.5>Encounter Number</ZCP.5>
'      <ZCP.6>Medcin ID</ZCP.6>
'      <ZCP.7>
'         <XCN.1>Clinician SSN</XCN.1>
'         <XCN.2>Clinician Last Name</XCN.2>
'         <XCN.3>Clinician First Name</XCN.3>
'         <XCN.4>Clinician Middle Name</XCN.4>
'      </ZCP.7>
'      <ZCP.8>Point of Care Facility</ZCP.8>
'      <ZCP.9>Status</ZCP.9>
'      <ZCP.10>Noted Date</ZCP.10>
'      <ZCP.11>Onset Date</ZCP.11>
'      <ZCP.12>Complaint NCID</ZCP.12>
'      <ZCP.13>New Follow Up</ZCP.13>
'      <ZCP.14>Create User</ZCP.14>
'      <ZCP.15>Create Date/Time</ZCP.15>
'      <ZCP.16>Modify User</ZCP.16>
'      <ZCP.17>Modify Date/Time</ZCP.17>
'      <ZCP.18>Update Flag</ZCP.18>
'      <ZCP.19>Complaint Comments</ZCP.19>
'      <ZCP.20>CDR Data ID</ZCP.20>
'      <ZCP.21>Created By</ZCP.21>
'      <ZCP.22>Created On</ZCP.22>
'      <ZCP.23>Updated By</ZCP.23>
'      <ZCP.24>Updated On</ZCP.24>
'   </ZCP>

    Dim oDom As MSXML.DOMDocument
    Dim oZCPNode As MSXML.IXMLDOMElement
    Dim oZCPChildLevel1 As MSXML.IXMLDOMElement
    Dim oZCPChildLevel2 As MSXML.IXMLDOMElement
    Dim strDocFirstName As String
    Dim strDocLastName As String
    Dim strDocMiddleName As String
    Dim strDocID As String
    Dim strDocSpecialty As String
    Dim strTemp As String
    Dim blnRsIsGood As Boolean
    
    On Error GoTo ErrTrap

    Set oDom = New MSXML.DOMDocument
    Set oZCPNode = oDom.createElement("ZCP")
    
    ' Set ID - Count
    Set oZCPChildLevel1 = oDom.createElement("ZCP.1")
    oZCPChildLevel1.Text = CStr(lngCount)
    oZCPNode.appendChild oZCPChildLevel1

    ' DataID - Table Complaints:DataID
    Set oZCPChildLevel1 = oDom.createElement("ZCP.2")
    oZCPChildLevel1.Text = NullToNothing(ComplaintsRS.Fields("DataID").Value)
    oZCPNode.appendChild oZCPChildLevel1

    ' Facility - Table Complaints:FacilityNCID
    Set oZCPChildLevel1 = oDom.createElement("ZCP.3")
    oZCPChildLevel1.Text = NullToNothing(ComplaintsRS.Fields("FacilityNCID").Value)
    oZCPNode.appendChild oZCPChildLevel1

    ' Unit Number - Table Complaints:Unit_Number
    Set oZCPChildLevel1 = oDom.createElement("ZCP.4")
    oZCPChildLevel1.Text = NullToNothing(ComplaintsRS.Fields("Unit_Number").Value)
    oZCPNode.appendChild oZCPChildLevel1
    
    ' Encounter Number - Table Complaints:EncounterNumber
    Set oZCPChildLevel1 = oDom.createElement("ZCP.5")
    oZCPChildLevel1.Text = NullToNothing(ComplaintsRS.Fields("EncounterNumber").Value)
    oZCPNode.appendChild oZCPChildLevel1
    
    ' Medcin ID - Table Complaints:SnoID
    Set oZCPChildLevel1 = oDom.createElement("ZCP.6")
    oZCPChildLevel1.Text = NullToNothing(ComplaintsRS.Fields("SnoID").Value)
    oZCPNode.appendChild oZCPChildLevel1
    
    ' Clinician
    Set oZCPChildLevel1 = oDom.createElement("ZCP.7")
        strTemp = NullToNothing(ComplaintsRS.Fields("ClinicianNCID").Value)
        If strTemp <> "" Then
            Call GetNewDoctorInfo(CLng(strTemp), strDocLastName, strDocFirstName, strDocMiddleName, strDocID, strDocSpecialty)
            blnRsIsGood = True
        End If
         ' Clinician SSN - Table Complaints:ClinicianNCID
        Set oZCPChildLevel2 = oDom.createElement("XCN.1")
        If blnRsIsGood = True Then
            oZCPChildLevel2.Text = strDocID
        End If
        oZCPChildLevel1.appendChild oZCPChildLevel2
        ' Clinician Last Name
        Set oZCPChildLevel2 = oDom.createElement("XCN.2")
        If blnRsIsGood = True Then
            oZCPChildLevel2.Text = strDocLastName
        End If
        oZCPChildLevel1.appendChild oZCPChildLevel2
        ' Clinician First Name
        Set oZCPChildLevel2 = oDom.createElement("XCN.3")
        If blnRsIsGood = True Then
            oZCPChildLevel2.Text = strDocFirstName
        End If
        oZCPChildLevel1.appendChild oZCPChildLevel2
        ' Clinician Middle Name
        Set oZCPChildLevel2 = oDom.createElement("XCN.4")
        If blnRsIsGood = True Then
            oZCPChildLevel2.Text = strDocMiddleName
        End If
        oZCPChildLevel1.appendChild oZCPChildLevel2
    oZCPNode.appendChild oZCPChildLevel1

    ' Point of Care Facility - Table Complaints:PointOfCareFacilityNCID
    Set oZCPChildLevel1 = oDom.createElement("ZCP.8")
    'oZCPChildLevel1.Text = NullToNothing(ComplaintsRS.Fields("PointOfCareFacilityNCID").Value)
        strTemp = NullToNothing(ComplaintsRS.Fields("PointOfCareFacilityNCID").Value)
        
        Set oZCPChildLevel2 = oDom.createElement("CE.1")
        oZCPChildLevel1.appendChild oZCPChildLevel2
        
        Set oZCPChildLevel2 = oDom.createElement("CE.2")
        If strTemp <> "" Then
            oZCPChildLevel2.Text = GetNCIDDescription(CLng(strTemp))
        End If
        oZCPChildLevel1.appendChild oZCPChildLevel2
        
        Set oZCPChildLevel2 = oDom.createElement("CE.3")
        oZCPChildLevel1.appendChild oZCPChildLevel2
        Set oZCPChildLevel2 = oDom.createElement("CE.4")
        oZCPChildLevel1.appendChild oZCPChildLevel2
        Set oZCPChildLevel2 = oDom.createElement("CE.5")
        oZCPChildLevel2.Text = strTemp 'The NCID
        oZCPChildLevel1.appendChild oZCPChildLevel2
        Set oZCPChildLevel2 = oDom.createElement("CE.6")
        oZCPChildLevel2.Text = "99NCI"
        oZCPChildLevel1.appendChild oZCPChildLevel2
    oZCPNode.appendChild oZCPChildLevel1
    
    ' Status - Table Complaints:StatusNCID
    Set oZCPChildLevel1 = oDom.createElement("ZCP.9")
    'oZCPChildLevel1.Text = NullToNothing(ComplaintsRS.Fields("StatusNCID").Value)
        strTemp = NullToNothing(ComplaintsRS.Fields("StatusNCID").Value)
        
        Set oZCPChildLevel2 = oDom.createElement("CE.1")
        oZCPChildLevel1.appendChild oZCPChildLevel2
        
        Set oZCPChildLevel2 = oDom.createElement("CE.2")
        If strTemp <> "" Then
            oZCPChildLevel2.Text = GetNCIDDescription(CLng(strTemp))
        End If
        oZCPChildLevel1.appendChild oZCPChildLevel2
        
        Set oZCPChildLevel2 = oDom.createElement("CE.3")
        oZCPChildLevel1.appendChild oZCPChildLevel2
        Set oZCPChildLevel2 = oDom.createElement("CE.4")
        oZCPChildLevel1.appendChild oZCPChildLevel2
        Set oZCPChildLevel2 = oDom.createElement("CE.5")
        oZCPChildLevel2.Text = strTemp 'The NCID
        oZCPChildLevel1.appendChild oZCPChildLevel2
        Set oZCPChildLevel2 = oDom.createElement("CE.6")
        oZCPChildLevel2.Text = "99NCI"
        oZCPChildLevel1.appendChild oZCPChildLevel2
    oZCPNode.appendChild oZCPChildLevel1
    
    ' Noted Date - Table Complaints:NotedDate
    Set oZCPChildLevel1 = oDom.createElement("ZCP.10")
    strTemp = NullToNothing(ComplaintsRS.Fields("NotedDate").Value)
    If IsValidDate(strTemp) Then
        oZCPChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZCPNode.appendChild oZCPChildLevel1
    
    ' Onset Date - Table Complaints:OnsetDate
    Set oZCPChildLevel1 = oDom.createElement("ZCP.11")
    strTemp = NullToNothing(ComplaintsRS.Fields("OnsetDate").Value)
    If IsValidDate(strTemp) Then
        oZCPChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZCPNode.appendChild oZCPChildLevel1
    
    ' Complaint - Table Complaints:ComplaintNCID
    Set oZCPChildLevel1 = oDom.createElement("ZCP.12")
    'oZCPChildLevel1.Text = NullToNothing(ComplaintsRS.Fields("ComplaintNCID").Value)
        strTemp = NullToNothing(ComplaintsRS.Fields("ComplaintNCID").Value)
        
        Set oZCPChildLevel2 = oDom.createElement("CE.1")
        oZCPChildLevel1.appendChild oZCPChildLevel2
        
        Set oZCPChildLevel2 = oDom.createElement("CE.2")
        If strTemp <> "" Then
            oZCPChildLevel2.Text = GetNCIDDescription(CLng(strTemp))
        End If
        oZCPChildLevel1.appendChild oZCPChildLevel2
        
        Set oZCPChildLevel2 = oDom.createElement("CE.3")
        oZCPChildLevel1.appendChild oZCPChildLevel2
        Set oZCPChildLevel2 = oDom.createElement("CE.4")
        oZCPChildLevel1.appendChild oZCPChildLevel2
        Set oZCPChildLevel2 = oDom.createElement("CE.5")
        oZCPChildLevel2.Text = strTemp 'The NCID
        oZCPChildLevel1.appendChild oZCPChildLevel2
        Set oZCPChildLevel2 = oDom.createElement("CE.6")
        oZCPChildLevel2.Text = "99NCI"
        oZCPChildLevel1.appendChild oZCPChildLevel2
    oZCPNode.appendChild oZCPChildLevel1
    
    ' New Follow Up - Table Complaints:NewFollowUp
    Set oZCPChildLevel1 = oDom.createElement("ZCP.13")
    oZCPChildLevel1.Text = NullToNothing(ComplaintsRS.Fields("NewFollowUp").Value)
    oZCPNode.appendChild oZCPChildLevel1
    
    ' Create User - Table Complaints:CreateUserNCID
    Set oZCPChildLevel1 = oDom.createElement("ZCP.14")
    oZCPChildLevel1.Text = NullToNothing(ComplaintsRS.Fields("CreateUserNCID").Value)
    oZCPNode.appendChild oZCPChildLevel1
    
    ' Create Date/Time - Table Complaints:CreateTime
    Set oZCPChildLevel1 = oDom.createElement("ZCP.15")
    strTemp = NullToNothing(ComplaintsRS.Fields("CreateTime").Value)
    If IsValidDate(strTemp) Then
        oZCPChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZCPNode.appendChild oZCPChildLevel1
    
    ' Modify User - Table Complaints:ModifyUserNCID
    Set oZCPChildLevel1 = oDom.createElement("ZCP.16")
    oZCPChildLevel1.Text = NullToNothing(ComplaintsRS.Fields("ModifyUserNCID").Value)
    oZCPNode.appendChild oZCPChildLevel1
    
    ' Modify Date/Time - Table Complaints:ModifyTime
    Set oZCPChildLevel1 = oDom.createElement("ZCP.17")
    strTemp = NullToNothing(ComplaintsRS.Fields("ModifyTime").Value)
    If IsValidDate(strTemp) Then
        oZCPChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZCPNode.appendChild oZCPChildLevel1
    
    ' Update Flag - Table Complaints:UpdateFlag
    Set oZCPChildLevel1 = oDom.createElement("ZCP.18")
    oZCPChildLevel1.Text = NullToNothing(ComplaintsRS.Fields("UpdateFlag").Value)
    oZCPNode.appendChild oZCPChildLevel1
    
    ' Comments - Table Complaints:ComplaintsComment
    Set oZCPChildLevel1 = oDom.createElement("ZCP.19")
    oZCPChildLevel1.Text = NullToNothing(ComplaintsRS.Fields("ComplaintsComment").Value)
    oZCPNode.appendChild oZCPChildLevel1
    
    ' CDR Data ID - Table Complaints:CDR_DataID
    Set oZCPChildLevel1 = oDom.createElement("ZCP.20")
    oZCPChildLevel1.Text = NullToNothing(ComplaintsRS.Fields("CDR_DataID").Value)
    oZCPNode.appendChild oZCPChildLevel1
    
    ' Created By - Table Complaints:CreatedBy
    Set oZCPChildLevel1 = oDom.createElement("ZCP.21")
    oZCPChildLevel1.Text = NullToNothing(ComplaintsRS.Fields("CreatedBy").Value)
    oZCPNode.appendChild oZCPChildLevel1
    
    ' Created On - Table Complaints:CreatedOn
    Set oZCPChildLevel1 = oDom.createElement("ZCP.22")
    strTemp = NullToNothing(ComplaintsRS.Fields("CreatedOn").Value)
    If IsValidDate(strTemp) Then
        oZCPChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZCPNode.appendChild oZCPChildLevel1
    
    ' Updated By - Table Complaints:UpdatedBy
    Set oZCPChildLevel1 = oDom.createElement("ZCP.23")
    oZCPChildLevel1.Text = NullToNothing(ComplaintsRS.Fields("UpdatedBy").Value)
    oZCPNode.appendChild oZCPChildLevel1
    
    ' Updated On - Table Complaints:UpdatedOn
    Set oZCPChildLevel1 = oDom.createElement("ZCP.24")
    strTemp = NullToNothing(ComplaintsRS.Fields("UpdatedOn").Value)
    If IsValidDate(strTemp) Then
        oZCPChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZCPNode.appendChild oZCPChildLevel1
    
    Set ReturnNewZCPNode = oZCPNode

Exit Function

ErrTrap:
   If gobjShared Is Nothing Then
      Set gobjShared = New CWShared
   End If
   gobjShared.ShowVBError Err.Number, Err.Description, "ReturnNewZCPNode Function", "DMBE", vbExclamation
   'Err.Clear
   Resume Next 'There is no reason this procedure should not continue to create xml if it errors

End Function

Public Function ReturnNewZLRNode(ByVal lngCount As Long, ByVal ListToolRecordsRS As ADODB.Recordset) As MSXML.IXMLDOMElement
' Encounter List Tool Records Segment
' sample function output
'   <ZLR>
'      <ZLR.1>Count</ZLR.1>
'      <ZLR.2>DataID</ZLR.2>
'      <ZLR.3>Medcin ID</ZLR.3>
'      <ZLR.4>Prefix</ZLR.4>
'      <ZLR.5>Modifier</ZLR.5>
'      <ZLR.6>Result</ZLR.6>
'      <ZLR.7>Status</ZLR.7>
'      <ZLR.8>Onset</ZLR.8>
'      <ZLR.9>Duration</ZLR.9>
'      <ZLR.10>Value</ZLR.10>
'      <ZLR.11>Unit</ZLR.11>
'      <ZLR.12>RangeScale</ZLR.12>
'      <ZLR.13>Refer ID</ZLR.13>
'      <ZLR.14>Range Normal Low</ZLR.14>
'      <ZLR.15>Range Normal High</ZLR.15>
'      <ZLR.16>Chart Flag</ZLR.16>
'      <ZLR.17>Qlink</ZLR.17>
'      <ZLR.18>Qlink Seq</ZLR.18>
'      <ZLR.19>Note</ZLR.19>
'      <ZLR.20>Update Flag</ZLR.20>
'      <ZLR.21>Created By</ZLR.21>
'      <ZLR.22>Created On</ZLR.22>
'      <ZLR.23>Updated By</ZLR.23>
'      <ZLR.24>Updated On</ZLR.24>
'   </ZLR>

    Dim oDom As MSXML.DOMDocument
    Dim oZLRNode As MSXML.IXMLDOMElement
    Dim oZLRChildLevel1 As MSXML.IXMLDOMElement
    Dim strTemp As String
    
    On Error GoTo ErrTrap

    Set oDom = New MSXML.DOMDocument
    Set oZLRNode = oDom.createElement("ZLR")
    
    ' Count
    Set oZLRChildLevel1 = oDom.createElement("ZLR.1")
    oZLRChildLevel1.Text = CStr(lngCount)
    oZLRNode.appendChild oZLRChildLevel1

    ' DataID - Table ListToolRecords:DataID
    Set oZLRChildLevel1 = oDom.createElement("ZLR.2")
    oZLRChildLevel1.Text = NullToNothing(ListToolRecordsRS.Fields("DataID").Value)
    oZLRNode.appendChild oZLRChildLevel1

    ' Medcin ID - Table ListToolRecords:SnoID
    Set oZLRChildLevel1 = oDom.createElement("ZLR.3")
    oZLRChildLevel1.Text = NullToNothing(ListToolRecordsRS.Fields("SnoID").Value)
    oZLRNode.appendChild oZLRChildLevel1

    ' Prefix - Table ListToolRecords:Prefix
    Set oZLRChildLevel1 = oDom.createElement("ZLR.4")
    oZLRChildLevel1.Text = NullToNothing(ListToolRecordsRS.Fields("Prefix").Value)
    oZLRNode.appendChild oZLRChildLevel1

    ' Modifier - Table ListToolRecords:Modifier
    Set oZLRChildLevel1 = oDom.createElement("ZLR.5")
    oZLRChildLevel1.Text = NullToNothing(ListToolRecordsRS.Fields("Modifier").Value)
    oZLRNode.appendChild oZLRChildLevel1
    
    ' Result - Table ListToolRecords:Result
    Set oZLRChildLevel1 = oDom.createElement("ZLR.6")
    oZLRChildLevel1.Text = NullToNothing(ListToolRecordsRS.Fields("Result").Value)
    oZLRNode.appendChild oZLRChildLevel1
    
    ' Status - Table ListToolRecords:Status
    Set oZLRChildLevel1 = oDom.createElement("ZLR.7")
    oZLRChildLevel1.Text = NullToNothing(ListToolRecordsRS.Fields("Status").Value)
    oZLRNode.appendChild oZLRChildLevel1
    
    ' Onset - Table ListToolRecords:Onset
    Set oZLRChildLevel1 = oDom.createElement("ZLR.8")
    oZLRChildLevel1.Text = NullToNothing(ListToolRecordsRS.Fields("Onset").Value)
    oZLRNode.appendChild oZLRChildLevel1
    
    ' Duration - Table ListToolRecords:Duration
    Set oZLRChildLevel1 = oDom.createElement("ZLR.9")
    oZLRChildLevel1.Text = NullToNothing(ListToolRecordsRS.Fields("Duration").Value)
    oZLRNode.appendChild oZLRChildLevel1
    
    ' Value - Table ListToolRecords:Value
    Set oZLRChildLevel1 = oDom.createElement("ZLR.10")
    oZLRChildLevel1.Text = NullToNothing(ListToolRecordsRS.Fields("Value").Value)
    oZLRNode.appendChild oZLRChildLevel1
    
    ' Unit - Table ListToolRecords:Unit
    Set oZLRChildLevel1 = oDom.createElement("ZLR.11")
    oZLRChildLevel1.Text = NullToNothing(ListToolRecordsRS.Fields("Unit").Value)
    oZLRNode.appendChild oZLRChildLevel1
    
    ' RangeScale - Table ListToolRecords:RangeScale
    Set oZLRChildLevel1 = oDom.createElement("ZLR.12")
    oZLRChildLevel1.Text = NullToNothing(ListToolRecordsRS.Fields("RangeScale").Value)
    oZLRNode.appendChild oZLRChildLevel1
    
    ' Refer ID - Table ListToolRecords:ReferID
    Set oZLRChildLevel1 = oDom.createElement("ZLR.13")
    oZLRChildLevel1.Text = NullToNothing(ListToolRecordsRS.Fields("ReferID").Value)
    oZLRNode.appendChild oZLRChildLevel1
    
    ' Range Normal Low - Table ListToolRecords:RangeNormalLow
    Set oZLRChildLevel1 = oDom.createElement("ZLR.14")
    oZLRChildLevel1.Text = NullToNothing(ListToolRecordsRS.Fields("RangeNormalLow").Value)
    oZLRNode.appendChild oZLRChildLevel1
    
    ' Range Normal High - Table ListToolRecords:RangeNormalHigh
    Set oZLRChildLevel1 = oDom.createElement("ZLR.15")
    oZLRChildLevel1.Text = NullToNothing(ListToolRecordsRS.Fields("RangeNormalHigh").Value)
    oZLRNode.appendChild oZLRChildLevel1
    
    ' Chart Flag - Table ListToolRecords:ChartFlag
    Set oZLRChildLevel1 = oDom.createElement("ZLR.16")
    oZLRChildLevel1.Text = NullToNothing(ListToolRecordsRS.Fields("ChartFlag").Value)
    oZLRNode.appendChild oZLRChildLevel1
    
    ' Qlink - Table ListToolRecords:Qlink
    Set oZLRChildLevel1 = oDom.createElement("ZLR.17")
    oZLRChildLevel1.Text = NullToNothing(ListToolRecordsRS.Fields("Qlink").Value)
    oZLRNode.appendChild oZLRChildLevel1
    
    ' Qlink Seq - Table ListToolRecords:QlinkSeq
    Set oZLRChildLevel1 = oDom.createElement("ZLR.18")
    oZLRChildLevel1.Text = NullToNothing(ListToolRecordsRS.Fields("QlinkSeq").Value)
    oZLRNode.appendChild oZLRChildLevel1
    
    ' Note - Table ListToolRecords:Note
    Set oZLRChildLevel1 = oDom.createElement("ZLR.19")
    oZLRChildLevel1.Text = NullToNothing(ListToolRecordsRS.Fields("Note").Value)
    oZLRNode.appendChild oZLRChildLevel1
    
    ' Update Flag - Table ListToolRecords:UpdateFlag
    Set oZLRChildLevel1 = oDom.createElement("ZLR.20")
    oZLRChildLevel1.Text = NullToNothing(ListToolRecordsRS.Fields("UpdateFlag").Value)
    oZLRNode.appendChild oZLRChildLevel1
    
    ' Created By - Table ListToolRecords:CreatedBy
    Set oZLRChildLevel1 = oDom.createElement("ZLR.21")
    oZLRChildLevel1.Text = NullToNothing(ListToolRecordsRS.Fields("CreatedBy").Value)
    oZLRNode.appendChild oZLRChildLevel1
    
    ' Created On - Table ListToolRecords:CreatedOn
    Set oZLRChildLevel1 = oDom.createElement("ZLR.22")
    strTemp = NullToNothing(ListToolRecordsRS.Fields("CreatedOn").Value)
    If IsValidDate(strTemp) Then
        oZLRChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZLRNode.appendChild oZLRChildLevel1
    
    ' Updated By - Table ListToolRecords:UpdatedBy
    Set oZLRChildLevel1 = oDom.createElement("ZLR.23")
    oZLRChildLevel1.Text = NullToNothing(ListToolRecordsRS.Fields("UpdatedBy").Value)
    oZLRNode.appendChild oZLRChildLevel1
    
    ' Updated On - Table ListToolRecords:UpdatedOn
    Set oZLRChildLevel1 = oDom.createElement("ZLR.24")
    strTemp = NullToNothing(ListToolRecordsRS.Fields("UpdatedOn").Value)
    If IsValidDate(strTemp) Then
        oZLRChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZLRNode.appendChild oZLRChildLevel1
    
    Set ReturnNewZLRNode = oZLRNode

Exit Function

ErrTrap:
   If gobjShared Is Nothing Then
      Set gobjShared = New CWShared
   End If
   gobjShared.ShowVBError Err.Number, Err.Description, "ReturnNewZLRNode Function", "DMBE", vbExclamation
   Resume Next 'There is no reason this procedure should not continue to create xml if it errors
End Function

Public Function ReturnNewZLTNode(ByVal lngCount As Long, ByVal ListToolRecordRS As ADODB.Recordset) As MSXML.IXMLDOMElement
' Encounter List Tool Record Segment
' sample function output
'   <ZLT>
'      <ZLT.1>Count</ZLT.1>
'      <ZLT.2>DataID</ZLT.2>
'      <ZLT.3>Facility NCID</ZLT.3>
'      <ZLT.4>Encounter Number</ZLT.4>
'      <ZLT.5>Medcin ID</ZLT.5>
'      <ZLT.6>
'         <XCN.1>Clinician SSN</XCN.1>
'         <XCN.2>Clinician Last Name</XCN.2>
'         <XCN.3>Clinician First Name</XCN.3>
'         <XCN.4>Clinician Middle Name</XCN.4>
'      </ZLT.6>
'      <ZLT.7>Clinic</ZLT.7>
'      <ZLT.8>Start Date/Time</ZLT.8>
'      <ZLT.9>End Date/Time</ZLT.9>
'      <ZLT.10>Source</ZLT.10>
'      <ZLT.11>RTF</ZLT.11>
'      <ZLT.12>Update Flag</ZLT.12>
'      <ZLT.13>Created By</ZLT.13>
'      <ZLT.14>Created On</ZLT.14>
'      <ZLT.15>Updated By</ZLT.15>
'      <ZLT.16>Updated On</ZLT.16>
' Begin:SCR #22830;   Developer: MAG 07/19/2002 03:55 PM
'      <ZLT.17>Original Size</ZLT.17>
' End: SCR #22830;
'   </ZLT>

    Dim oDom As MSXML.DOMDocument
    Dim oZLTNode As MSXML.IXMLDOMElement
    Dim oZLTChildLevel1 As MSXML.IXMLDOMElement
    Dim oZLTChildLevel2 As MSXML.IXMLDOMElement
    Dim strDocFirstName As String
    Dim strDocLastName As String
    Dim strDocMiddleName As String
    Dim strDocID As String
    Dim strDocSpecialty As String
    Dim strTemp As String
    Dim blnRsIsGood As Boolean
    
    On Error GoTo ErrTrap

    Set oDom = New MSXML.DOMDocument
    Set oZLTNode = oDom.createElement("ZLT")
    
    ' Count
    Set oZLTChildLevel1 = oDom.createElement("ZLT.1")
    oZLTChildLevel1.Text = CStr(lngCount)
    oZLTNode.appendChild oZLTChildLevel1

    ' DataID - Table ListToolRecord:DataID
    Set oZLTChildLevel1 = oDom.createElement("ZLT.2")
    oZLTChildLevel1.Text = NullToNothing(ListToolRecordRS.Fields("DataID").Value)
    oZLTNode.appendChild oZLTChildLevel1

    ' Facility - Table ListToolRecord:FacilityNCID
    Set oZLTChildLevel1 = oDom.createElement("ZLT.3")
    oZLTChildLevel1.Text = NullToNothing(ListToolRecordRS.Fields("FacilityNCID").Value)
    oZLTNode.appendChild oZLTChildLevel1

    ' Encounter Number - Table ListToolRecord:EncounterNumber
    Set oZLTChildLevel1 = oDom.createElement("ZLT.4")
    oZLTChildLevel1.Text = NullToNothing(ListToolRecordRS.Fields("EncounterNumber").Value)
    oZLTNode.appendChild oZLTChildLevel1
    
    ' Medcin ID - Table ListToolRecord:SnoID
    Set oZLTChildLevel1 = oDom.createElement("ZLT.5")
    oZLTChildLevel1.Text = NullToNothing(ListToolRecordRS.Fields("SnoID").Value)
    oZLTNode.appendChild oZLTChildLevel1
    
    ' Clinician
    Set oZLTChildLevel1 = oDom.createElement("ZLT.6")
        strTemp = NullToNothing(ListToolRecordRS.Fields("ClinicianNCID").Value)
        If strTemp <> "" Then
            Call GetNewDoctorInfo(CLng(strTemp), strDocLastName, strDocFirstName, strDocMiddleName, strDocID, strDocSpecialty)
            blnRsIsGood = True
        End If
        ' Clinician SSN - Table ListToolRecord:ClinicianNCID
        Set oZLTChildLevel2 = oDom.createElement("XCN.1")
        If blnRsIsGood = True Then
            oZLTChildLevel2.Text = strDocID
        End If
        oZLTChildLevel1.appendChild oZLTChildLevel2
        ' Clinician Last Name
        Set oZLTChildLevel2 = oDom.createElement("XCN.2")
        If blnRsIsGood = True Then
            oZLTChildLevel2.Text = strDocLastName
        End If
        oZLTChildLevel1.appendChild oZLTChildLevel2
        ' Clinician First Name
        Set oZLTChildLevel2 = oDom.createElement("XCN.3")
        If blnRsIsGood = True Then
            oZLTChildLevel2.Text = strDocFirstName
        End If
        oZLTChildLevel1.appendChild oZLTChildLevel2
        ' Clinician Middle Name
        Set oZLTChildLevel2 = oDom.createElement("XCN.4")
        If blnRsIsGood = True Then
            oZLTChildLevel2.Text = strDocMiddleName
        End If
        oZLTChildLevel1.appendChild oZLTChildLevel2
    oZLTNode.appendChild oZLTChildLevel1
    
    ' Clinic - Table ListToolRecord:ClinicNCID
    Set oZLTChildLevel1 = oDom.createElement("ZLT.7")
    oZLTChildLevel1.Text = NullToNothing(ListToolRecordRS.Fields("ClinicNCID").Value)
    oZLTNode.appendChild oZLTChildLevel1
    
    ' Start Date/Time - Table ListToolRecord:StartDTS
    Set oZLTChildLevel1 = oDom.createElement("ZLT.8")
    strTemp = NullToNothing(ListToolRecordRS.Fields("StartDTS").Value)
    If IsValidDate(strTemp) Then
        oZLTChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZLTNode.appendChild oZLTChildLevel1
    
    ' End Date/Time - Table ListToolRecord:EndDTS
    Set oZLTChildLevel1 = oDom.createElement("ZLT.9")
    strTemp = NullToNothing(ListToolRecordRS.Fields("EndDTS").Value)
    If IsValidDate(strTemp) Then
        oZLTChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZLTNode.appendChild oZLTChildLevel1
    
    ' Source - Table ListToolRecord:SourceNCID
    Set oZLTChildLevel1 = oDom.createElement("ZLT.10")
    oZLTChildLevel1.Text = NullToNothing(ListToolRecordRS.Fields("SourceNCID").Value)
    oZLTNode.appendChild oZLTChildLevel1
    
    ' RTF - Table ListToolRecord:RTF
    Set oZLTChildLevel1 = oDom.createElement("ZLT.11")
    strTemp = NullToNothing(ListToolRecordRS.Fields("RTF").Value)
    If strTemp <> "" Then
      '<< Begin:SCR #25718;   Developer: Brian Mowbray 07/18/2002 01:48 PM
      If ListToolRecordRS.Fields("OriginalSize").Value > 1 Then
        oZLTChildLevel1.Text = EncodeString(strTemp, True)
      Else
        oZLTChildLevel1.Text = FilterRtfToXml(strTemp)
      End If
    '>> End: SCR #25718;
    End If
    oZLTNode.appendChild oZLTChildLevel1
    
    ' Update Flag - Table ListToolRecord:UpdateFlag
    Set oZLTChildLevel1 = oDom.createElement("ZLT.12")
    oZLTChildLevel1.Text = NullToNothing(ListToolRecordRS.Fields("UpdateFlag").Value)
    oZLTNode.appendChild oZLTChildLevel1
    
    ' Created By - Table ListToolRecord:CreatedBy
    Set oZLTChildLevel1 = oDom.createElement("ZLT.13")
    oZLTChildLevel1.Text = NullToNothing(ListToolRecordRS.Fields("CreatedBy").Value)
    oZLTNode.appendChild oZLTChildLevel1
    
    ' Created On - Table ListToolRecord:CreatedOn
    Set oZLTChildLevel1 = oDom.createElement("ZLT.14")
    strTemp = NullToNothing(ListToolRecordRS.Fields("CreatedOn").Value)
    If IsValidDate(strTemp) Then
        oZLTChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZLTNode.appendChild oZLTChildLevel1
    
    ' Updated By - Table ListToolRecord:UpdatedBy
    Set oZLTChildLevel1 = oDom.createElement("ZLT.15")
    oZLTChildLevel1.Text = NullToNothing(ListToolRecordRS.Fields("UpdatedBy").Value)
    oZLTNode.appendChild oZLTChildLevel1
    
    ' Updated On - Table ListToolRecord:UpdatedOn
    Set oZLTChildLevel1 = oDom.createElement("ZLT.16")
    strTemp = NullToNothing(ListToolRecordRS.Fields("UpdatedOn").Value)
    If IsValidDate(strTemp) Then
        oZLTChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZLTNode.appendChild oZLTChildLevel1
    
    ' Begin:SCR #22830;   Developer: MAG 07/19/2002 03:53 PM
    ' Original Size - Table ListToolRecord:UpdatedOn
    Set oZLTChildLevel1 = oDom.createElement("ZLT.17")
    oZLTChildLevel1.Text = NullToNothing(ListToolRecordRS.Fields("OriginalSize").Value)
    oZLTNode.appendChild oZLTChildLevel1
    ' End: SCR #22830;
    
    Set ReturnNewZLTNode = oZLTNode

Exit Function

ErrTrap:
   If gobjShared Is Nothing Then
      Set gobjShared = New CWShared
   End If
   gobjShared.ShowVBError Err.Number, Err.Description, "ReturnNewZLTNode Function", "DMBE", vbExclamation
   Resume Next 'There is no reason this procedure should not continue to create xml if it errors
End Function

Public Function ReturnNewZEBNode(ByVal lngCount As Long, ByVal Enc_SigBlockRS As ADODB.Recordset) As MSXML.IXMLDOMElement
' Encounter Signature Block Segment
' sample function output
'   <ZEB>
'      <ZEB.1>Count</ZEB.1>
'      <ZEB.2>DataID</ZEB.2>
'      <ZEB.3>Owner Data ID</ZEB.3>
'      <ZEB.4>User NCID</ZEB.4>
'      <ZEB.5>Line 1</ZEB.5>
'      <ZEB.6>Line 2</ZEB.6>
'      <ZEB.7>Line 3</ZEB.7>
'      <ZEB.8>Date/Time Stamp</ZEB.8>
'      <ZEB.9>PKI</ZEB.9>
'      <ZEB.10>Created By</ZEB.10>
'      <ZEB.11>Created On</ZEB.11>
'      <ZEB.12>Updated By</ZEB.12>
'      <ZEB.13>Updated On</ZEB.13>
'   </ZEB>

    Dim oDom As MSXML.DOMDocument
    Dim oZEBNode As MSXML.IXMLDOMElement
    Dim oZEBChildLevel1 As MSXML.IXMLDOMElement
    Dim strTemp As String
    
    On Error GoTo ErrTrap

    Set oDom = New MSXML.DOMDocument
    Set oZEBNode = oDom.createElement("ZEB")
    
    ' Set ID - Count
    Set oZEBChildLevel1 = oDom.createElement("ZEB.1")
    oZEBChildLevel1.Text = CStr(lngCount)
    oZEBNode.appendChild oZEBChildLevel1

    ' DataID - Table Enc_SIGBlock:DataID
    Set oZEBChildLevel1 = oDom.createElement("ZEB.2")
    oZEBChildLevel1.Text = NullToNothing(Enc_SigBlockRS.Fields("DataID").Value)
    oZEBNode.appendChild oZEBChildLevel1

    ' Owner Data ID - Table Enc_SIGBlock:OwnerDataID
    Set oZEBChildLevel1 = oDom.createElement("ZEB.3")
    oZEBChildLevel1.Text = NullToNothing(Enc_SigBlockRS.Fields("OwnerDataID").Value)
    oZEBNode.appendChild oZEBChildLevel1
    
    ' User - Table Enc_SIGBlock:UserNCID
    Set oZEBChildLevel1 = oDom.createElement("ZEB.4")
    strTemp = NullToNothing(Enc_SigBlockRS.Fields("UserNCID").Value)
    If strTemp <> "" Then
        oZEBChildLevel1.Text = GetNCIDDescription(CLng(strTemp))
    End If
    oZEBNode.appendChild oZEBChildLevel1
    
    ' Line 1 - Table Enc_SIGBlock:Line1
    Set oZEBChildLevel1 = oDom.createElement("ZEB.5")
    oZEBChildLevel1.Text = NullToNothing(Enc_SigBlockRS.Fields("Line1").Value)
    oZEBNode.appendChild oZEBChildLevel1
    
    ' Line 2 - Table Enc_SIGBlock:Line2
    Set oZEBChildLevel1 = oDom.createElement("ZEB.6")
    oZEBChildLevel1.Text = NullToNothing(Enc_SigBlockRS.Fields("Line2").Value)
    oZEBNode.appendChild oZEBChildLevel1
    
    ' Line 3 - Table Enc_SIGBlock:Line3
    Set oZEBChildLevel1 = oDom.createElement("ZEB.7")
    oZEBChildLevel1.Text = NullToNothing(Enc_SigBlockRS.Fields("Line3").Value)
    oZEBNode.appendChild oZEBChildLevel1
    
    ' Date/Time Stamp - Table Enc_SIGBlock:DTS
    Set oZEBChildLevel1 = oDom.createElement("ZEB.8")
    strTemp = NullToNothing(Enc_SigBlockRS.Fields("DTS").Value)
    If IsValidDate(strTemp) Then
        oZEBChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZEBNode.appendChild oZEBChildLevel1
    
    ' PKI - Table Enc_SIGBlock:PKI
    Set oZEBChildLevel1 = oDom.createElement("ZEB.9")
    oZEBChildLevel1.Text = NullToNothing(Enc_SigBlockRS.Fields("PKI").Value)
    oZEBNode.appendChild oZEBChildLevel1
    
    ' Created By - Table Enc_SigBlock:CreatedBy
    Set oZEBChildLevel1 = oDom.createElement("ZEB.10")
    oZEBChildLevel1.Text = NullToNothing(Enc_SigBlockRS.Fields("CreatedBy").Value)
    oZEBNode.appendChild oZEBChildLevel1
    
    ' Created On - Table Enc_SigBlock:CreatedOn
    Set oZEBChildLevel1 = oDom.createElement("ZEB.11")
    strTemp = NullToNothing(Enc_SigBlockRS.Fields("CreatedOn").Value)
    If IsValidDate(strTemp) Then
        oZEBChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZEBNode.appendChild oZEBChildLevel1
    
    ' Updated By - Table Enc_SigBlock:UpdatedBy
    Set oZEBChildLevel1 = oDom.createElement("ZEB.12")
    oZEBChildLevel1.Text = NullToNothing(Enc_SigBlockRS.Fields("UpdatedBy").Value)
    oZEBNode.appendChild oZEBChildLevel1
    
    ' Updated On - Table Enc_SigBlock:UpdatedOn
    Set oZEBChildLevel1 = oDom.createElement("ZEB.13")
    strTemp = NullToNothing(Enc_SigBlockRS.Fields("UpdatedOn").Value)
    If IsValidDate(strTemp) Then
        oZEBChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZEBNode.appendChild oZEBChildLevel1
    
    Set ReturnNewZEBNode = oZEBNode

Exit Function

ErrTrap:
   If gobjShared Is Nothing Then
      Set gobjShared = New CWShared
   End If
   gobjShared.ShowVBError Err.Number, Err.Description, "ReturnNewZEBNode Function", "DMBE", vbExclamation
   Resume Next 'There is no reason this procedure should not continue to create xml if it errors
End Function

Public Function ReturnNewZERNode(ByVal lngCount As Long, ByVal Enc_RtfsRS As ADODB.Recordset) As MSXML.IXMLDOMElement
   ' Encounter RTF Segment
' sample function output
'   <ZER>
'      <ZER.1>Count</ZER.1>
'      <ZER.2>DataID</ZER.2>
'      <ZER.3>Encounter Number</ZER.3>
'      <ZER.4>Facility NCID</ZER.4>
'      <ZER.5>RTF Type</ZER.5>
'      <ZER.6>Encounter RTF Index</ZER.6>
'      <ZER.7>Note ID</ZER.7>
'      <ZER.8>Source</ZER.8>
'      <ZER.9>
'         <CE.1>User Identifier</CE.1>
'         <CE.2>User Name</CE.2>
'      </ZER.9>
'      <ZER.10>Date/Time Added</ZER.10>
'      <ZER.11>Date/Time Stamp</ZER.11>
'      <ZER.12>Document</ZER.12>
'      <ZER.13>Title</ZER.13>
'      <ZER.14>Status</ZER.14>
'      <ZER.15>Sensitivity Level</ZER.15>
'      <ZER.16>Co-Signer</ZER.16>
'      <ZER.17>Original Size</ZER.17>
'      <ZER.18>Initial Signature</ZER.18>
'      <ZER.19>Final Signature</ZER.19>
'      <ZER.20>Complete</ZER.20>
'      <ZER.21>Category</ZER.21>
'      <ZER.22>Update Flag</ZER.22>
'      <ZER.23>Created By</ZER.23>
'      <ZER.24>Created On</ZER.24>
'      <ZER.25>Updated By</ZER.25>
'      <ZER.26>Updated On</ZER.26>
'   </ZER>
   
    Dim oDom As MSXML.DOMDocument
    Dim oZERNode As MSXML.IXMLDOMElement
    Dim oZERChildLevel1 As MSXML.IXMLDOMElement
    Dim oZERChildLevel2 As MSXML.IXMLDOMElement
    Dim strTemp As String
    
    On Error GoTo ErrTrap

    Set oDom = New MSXML.DOMDocument
    Set oZERNode = oDom.createElement("ZER")
    
    ' Set ID - Count
    Set oZERChildLevel1 = oDom.createElement("ZER.1")
    oZERChildLevel1.Text = CStr(lngCount)
    oZERNode.appendChild oZERChildLevel1

    ' DataID - Table Enc_RTFs:DataID
    Set oZERChildLevel1 = oDom.createElement("ZER.2")
    ' dataid is not in table
    'oZERChildLevel1.Text = NullToNothing(Enc_RtfsRS.Fields("DataID").Value)
    oZERNode.appendChild oZERChildLevel1

    ' Encounter Number - Table Enc_RTFs:EncounterNumber
    Set oZERChildLevel1 = oDom.createElement("ZER.3")
    oZERChildLevel1.Text = NullToNothing(Enc_RtfsRS.Fields("EncounterNumber").Value)
    oZERNode.appendChild oZERChildLevel1

    ' Facility - Table Enc_RTFs:FacilityNCID
    Set oZERChildLevel1 = oDom.createElement("ZER.4")
    oZERChildLevel1.Text = NullToNothing(Enc_RtfsRS.Fields("FacilityNCID").Value)
    oZERNode.appendChild oZERChildLevel1

    ' RTF Type - Table Enc_RTFs:RTFType
    Set oZERChildLevel1 = oDom.createElement("ZER.5")
    oZERChildLevel1.Text = NullToNothing(Enc_RtfsRS.Fields("RTFType").Value)
    oZERNode.appendChild oZERChildLevel1

    ' Encounter RTF Index - Table Enc_RTFs:Enc_RTFIndex
    Set oZERChildLevel1 = oDom.createElement("ZER.6")
    oZERChildLevel1.Text = NullToNothing(Enc_RtfsRS.Fields("Enc_RTFSIndex").Value)
    oZERNode.appendChild oZERChildLevel1
    
    ' Note ID - Table Enc_RTFs:NoteID
    Set oZERChildLevel1 = oDom.createElement("ZER.7")
    oZERChildLevel1.Text = NullToNothing(Enc_RtfsRS.Fields("NoteID").Value)
    oZERNode.appendChild oZERChildLevel1
    
    ' Source - Table Enc_RTFs:Source
    Set oZERChildLevel1 = oDom.createElement("ZER.8")
    oZERChildLevel1.Text = NullToNothing(Enc_RtfsRS.Fields("Source").Value)
    oZERNode.appendChild oZERChildLevel1
    
    ' User
    Set oZERChildLevel1 = oDom.createElement("ZER.9")
        ' Identifier - Table Enc_RTFs:UserNCID
        Set oZERChildLevel2 = oDom.createElement("CE.1")
        oZERChildLevel2.Text = NullToNothing(Enc_RtfsRS.Fields("UserNCID").Value)
        oZERChildLevel1.appendChild oZERChildLevel2
        ' Name - Table Enc_RTFs:UserName
        Set oZERChildLevel2 = oDom.createElement("CE.2")
        oZERChildLevel2.Text = NullToNothing(Enc_RtfsRS.Fields("UserName").Value)
        oZERChildLevel1.appendChild oZERChildLevel2
    oZERNode.appendChild oZERChildLevel1

    ' Date/Time Added - Table Enc_RTFs:DateAdded
    Set oZERChildLevel1 = oDom.createElement("ZER.10")
    strTemp = NullToNothing(Enc_RtfsRS.Fields("DateAdded").Value)
    If IsValidDate(strTemp) Then
        oZERChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZERNode.appendChild oZERChildLevel1
    
    ' Date/Time Stamp - Table Enc_RTFs:DTS
    Set oZERChildLevel1 = oDom.createElement("ZER.11")
    strTemp = NullToNothing(Enc_RtfsRS.Fields("DTS").Value)
    If IsValidDate(strTemp) Then
        oZERChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZERNode.appendChild oZERChildLevel1
    
    ' Document - Table Enc_RTFs:Doc
    Set oZERChildLevel1 = oDom.createElement("ZER.12")
    strTemp = NullToNothing(Enc_RtfsRS.Fields("Doc").Value)
    If strTemp <> "" Then
      '<< Begin:SCR #25718;   Developer: Brian Mowbray 07/18/2002 01:48 PM
      ' Begin:SCR #43939;    Developer: Mag 10/15/2003 03:44 PM
      If Enc_RtfsRS.Fields("OriginalSize").Value > 1 Then
      ' End: SCR #43939
        oZERChildLevel1.Text = EncodeString(strTemp, True)
      Else
        oZERChildLevel1.Text = FilterRtfToXml(strTemp)
      End If
    '>> End: SCR #25718;
    End If
    oZERNode.appendChild oZERChildLevel1
    
    ' Title - Table Enc_RTFs:Title
    Set oZERChildLevel1 = oDom.createElement("ZER.13")
    oZERChildLevel1.Text = NullToNothing(Enc_RtfsRS.Fields("Title").Value)
    oZERNode.appendChild oZERChildLevel1
    
    ' Status - Table Enc_RTFs:Status
    Set oZERChildLevel1 = oDom.createElement("ZER.14")
    oZERChildLevel1.Text = NullToNothing(Enc_RtfsRS.Fields("Status").Value)
    oZERNode.appendChild oZERChildLevel1
    
    ' Sensitivity Level - Table Enc_RTFs:SensitivityLevel
    Set oZERChildLevel1 = oDom.createElement("ZER.15")
    oZERChildLevel1.Text = NullToNothing(Enc_RtfsRS.Fields("SensitivityLevel").Value)
    oZERNode.appendChild oZERChildLevel1
    
    ' Co-Signer - Table Enc_RTFs:CoSignerNCID
    Set oZERChildLevel1 = oDom.createElement("ZER.16")
    strTemp = NullToNothing(Enc_RtfsRS.Fields("CoSignerNCID").Value)
    If strTemp <> "" Then
        oZERChildLevel1.Text = GetNCIDDescription(CLng(strTemp))
    End If
    oZERNode.appendChild oZERChildLevel1
    
    ' Original Size - Table Enc_RTFs:OriginalSize
    Set oZERChildLevel1 = oDom.createElement("ZER.17")
    oZERChildLevel1.Text = NullToNothing(Enc_RtfsRS.Fields("OriginalSize").Value)
    oZERNode.appendChild oZERChildLevel1
    
    ' Initial Signature - Table Enc_RTFs:InitialSignature
    Set oZERChildLevel1 = oDom.createElement("ZER.18")
    oZERChildLevel1.Text = NullToNothing(Enc_RtfsRS.Fields("InitialSignature").Value)
    oZERNode.appendChild oZERChildLevel1
    
    ' Final Signature - Table Enc_RTFs:FinalSignature
    Set oZERChildLevel1 = oDom.createElement("ZER.19")
    oZERChildLevel1.Text = NullToNothing(Enc_RtfsRS.Fields("FinalSignature").Value)
    oZERNode.appendChild oZERChildLevel1
    
    ' Complete - Table Enc_RTFs:Complete
    Set oZERChildLevel1 = oDom.createElement("ZER.20")
    oZERChildLevel1.Text = NullToNothing(Enc_RtfsRS.Fields("Complete").Value)
    oZERNode.appendChild oZERChildLevel1
    
    ' Category - Table Enc_RTFs:Category
    Set oZERChildLevel1 = oDom.createElement("ZER.21")
    oZERChildLevel1.Text = NullToNothing(Enc_RtfsRS.Fields("Category").Value)
    oZERNode.appendChild oZERChildLevel1
    
    ' Update Flag - Table Enc_RTFs:UpdateFlag
    Set oZERChildLevel1 = oDom.createElement("ZER.22")
    oZERChildLevel1.Text = NullToNothing(Enc_RtfsRS.Fields("UpdateFlag").Value)
    oZERNode.appendChild oZERChildLevel1
    
    ' Created By - Table Enc_RTFs:CreatedBy
    Set oZERChildLevel1 = oDom.createElement("ZER.23")
    oZERChildLevel1.Text = NullToNothing(Enc_RtfsRS.Fields("CreatedBy").Value)
    oZERNode.appendChild oZERChildLevel1
    
    ' Created On - Table Enc_RTFs:CreatedOn
    Set oZERChildLevel1 = oDom.createElement("ZER.24")
    strTemp = NullToNothing(Enc_RtfsRS.Fields("CreatedOn").Value)
    If IsValidDate(strTemp) Then
        oZERChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZERNode.appendChild oZERChildLevel1
    
    ' Updated By - Table Enc_RTFs:UpdatedBy
    Set oZERChildLevel1 = oDom.createElement("ZER.25")
    oZERChildLevel1.Text = NullToNothing(Enc_RtfsRS.Fields("UpdatedBy").Value)
    oZERNode.appendChild oZERChildLevel1
    
    ' Updated On - Table Enc_RTFs:UpdatedOn
    Set oZERChildLevel1 = oDom.createElement("ZER.26")
    strTemp = NullToNothing(Enc_RtfsRS.Fields("UpdatedOn").Value)
    If IsValidDate(strTemp) Then
        oZERChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZERNode.appendChild oZERChildLevel1
    
    Set ReturnNewZERNode = oZERNode

Exit Function

ErrTrap:
   If gobjShared Is Nothing Then
      Set gobjShared = New CWShared
   End If
   gobjShared.ShowVBError Err.Number, Err.Description, "ReturnNewZERNode Function", "DMBE", vbExclamation
   Resume Next 'There is no reason this procedure should not continue to create xml if it errors
End Function


Public Function ReturnNewZEPNode_OE_RX(ByVal lngCount As Long, ByVal OE_RX_rs As ADODB.Recordset) As MSXML.IXMLDOMElement
    
    Dim oDom As MSXML.DOMDocument
    Dim oZEPNode As MSXML.IXMLDOMElement
    Dim oZEPChildLevel1 As MSXML.IXMLDOMElement
    Dim strTemp As String
    
    On Error GoTo ErrTrap

    Set oDom = New MSXML.DOMDocument
    Set oZEPNode = oDom.createElement("ZEP")
    
    Set oZEPChildLevel1 = oDom.createElement("ZEP.1")
    oZEPChildLevel1.Text = CStr(lngCount)
    oZEPNode.appendChild oZEPChildLevel1

    Set oZEPChildLevel1 = oDom.createElement("ZEP.2")
    oZEPChildLevel1.Text = NullToNothing(OE_RX_rs.Fields("OE_RX_DATA_ID").Value)
    oZEPNode.appendChild oZEPChildLevel1

    Set oZEPChildLevel1 = oDom.createElement("ZEP.3")
    oZEPChildLevel1.Text = NullToNothing(OE_RX_rs.Fields("ENC_NUM").Value)
    oZEPNode.appendChild oZEPChildLevel1
    
    Set oZEPChildLevel1 = oDom.createElement("ZEP.4")
    oZEPChildLevel1.Text = NullToNothing(OE_RX_rs.Fields("ENC_FAC_NCID").Value)
    oZEPNode.appendChild oZEPChildLevel1
    
    Set oZEPChildLevel1 = oDom.createElement("ZEP.5")
    strTemp = NullToNothing(OE_RX_rs.Fields("ORDER_DATE").Value)
    If IsValidDate(strTemp) Then
        oZEPChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZEPNode.appendChild oZEPChildLevel1
    
    Set oZEPChildLevel1 = oDom.createElement("ZEP.6")
    oZEPChildLevel1.Text = NullToNothing(OE_RX_rs.Fields("RX_COMMENT").Value)
    oZEPNode.appendChild oZEPChildLevel1
    
    Set oZEPChildLevel1 = oDom.createElement("ZEP.7")
    oZEPChildLevel1.Text = NullToNothing(OE_RX_rs.Fields("SIG_CODE").Value)
    oZEPNode.appendChild oZEPChildLevel1
    
    Set oZEPChildLevel1 = oDom.createElement("ZEP.8")
    oZEPChildLevel1.Text = NullToNothing(OE_RX_rs.Fields("QUANTITY").Value)
    oZEPNode.appendChild oZEPChildLevel1
        
    Set oZEPChildLevel1 = oDom.createElement("ZEP.9")
    oZEPChildLevel1.Text = NullToNothing(OE_RX_rs.Fields("REFILLS").Value)
    oZEPNode.appendChild oZEPChildLevel1
    
    Set oZEPChildLevel1 = oDom.createElement("ZEP.10")
    oZEPChildLevel1.Text = NullToNothing(OE_RX_rs.Fields("PHARMACY_NAME_SENT").Value)
    oZEPNode.appendChild oZEPChildLevel1
    
    Set oZEPChildLevel1 = oDom.createElement("ZEP.11")
    oZEPChildLevel1.Text = NullToNothing(OE_RX_rs.Fields("PHARMACY_IEN_SENT").Value)
    oZEPNode.appendChild oZEPChildLevel1
    
    Set oZEPChildLevel1 = oDom.createElement("ZEP.12")
    oZEPChildLevel1.Text = NullToNothing(OE_RX_rs.Fields("CHILD_CAP").Value)
    oZEPNode.appendChild oZEPChildLevel1
    
    Set oZEPChildLevel1 = oDom.createElement("ZEP.13")
    oZEPChildLevel1.Text = NullToNothing(OE_RX_rs.Fields("ORDER_IEN").Value)
    oZEPNode.appendChild oZEPChildLevel1
    
    Set oZEPChildLevel1 = oDom.createElement("ZEP.14")
    oZEPChildLevel1.Text = NullToNothing(OE_RX_rs.Fields("ORDER_NAME").Value)
    oZEPNode.appendChild oZEPChildLevel1
    
    Set oZEPChildLevel1 = oDom.createElement("ZEP.15")
    strTemp = NullToNothing(OE_RX_rs.Fields("TIME_ORDERED").Value)
    If IsValidDate(strTemp) Then
        oZEPChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZEPNode.appendChild oZEPChildLevel1

    Set oZEPChildLevel1 = oDom.createElement("ZEP.16")
    oZEPChildLevel1.Text = NullToNothing(OE_RX_rs.Fields("UNIT_NUMBER").Value)
    oZEPNode.appendChild oZEPChildLevel1
    
    Set oZEPChildLevel1 = oDom.createElement("ZEP.17")
    oZEPChildLevel1.Text = NullToNothing(OE_RX_rs.Fields("ORDER_ID").Value)
    oZEPNode.appendChild oZEPChildLevel1
    
    Set oZEPChildLevel1 = oDom.createElement("ZEP.18")
    oZEPChildLevel1.Text = NullToNothing(OE_RX_rs.Fields("SENSITIVE").Value)
    oZEPNode.appendChild oZEPChildLevel1
    
    Set oZEPChildLevel1 = oDom.createElement("ZEP.19")
    oZEPChildLevel1.Text = NullToNothing(OE_RX_rs.Fields("ORDERED_BY").Value)
    oZEPNode.appendChild oZEPChildLevel1
    
    Set oZEPChildLevel1 = oDom.createElement("ZEP.20")
    oZEPChildLevel1.Text = NullToNothing(OE_RX_rs.Fields("PROVIDER").Value)
    oZEPNode.appendChild oZEPChildLevel1
    
    Set oZEPChildLevel1 = oDom.createElement("ZEP.21")
    oZEPChildLevel1.Text = NullToNothing(OE_RX_rs.Fields("UPDATEDBY").Value)
    oZEPNode.appendChild oZEPChildLevel1
    
    Set oZEPChildLevel1 = oDom.createElement("ZEP.22")
    strTemp = NullToNothing(OE_RX_rs.Fields("UPDATEDON").Value)
    If IsValidDate(strTemp) Then
        oZEPChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZEPNode.appendChild oZEPChildLevel1

    Set oZEPChildLevel1 = oDom.createElement("ZEP.23")
    oZEPChildLevel1.Text = NullToNothing(OE_RX_rs.Fields("CREATEDBY").Value)
    oZEPNode.appendChild oZEPChildLevel1
    
    Set oZEPChildLevel1 = oDom.createElement("ZEP.24")
    strTemp = NullToNothing(OE_RX_rs.Fields("CREATEDON").Value)
    If IsValidDate(strTemp) Then
        oZEPChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZEPNode.appendChild oZEPChildLevel1
    
    Set oZEPChildLevel1 = oDom.createElement("ZEP.25")
    oZEPChildLevel1.Text = NullToNothing(OE_RX_rs.Fields("STATUS_NCID").Value)
    oZEPNode.appendChild oZEPChildLevel1
    
    Set oZEPChildLevel1 = oDom.createElement("ZEP.26")
    strTemp = NullToNothing(OE_RX_rs.Fields("FILL_DATE").Value)
    If IsValidDate(strTemp) Then
        oZEPChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZEPNode.appendChild oZEPChildLevel1
    
    Set oZEPChildLevel1 = oDom.createElement("ZEP.27")
    oZEPChildLevel1.Text = NullToNothing(OE_RX_rs.Fields("SIGN_FLAG").Value)
    oZEPNode.appendChild oZEPChildLevel1
    
    Set oZEPChildLevel1 = oDom.createElement("ZEP.28")
    oZEPChildLevel1.Text = NullToNothing(OE_RX_rs.Fields("NURSE_ACK_FALG").Value)
    oZEPNode.appendChild oZEPChildLevel1
    
    Set oZEPChildLevel1 = oDom.createElement("ZEP.29")
    oZEPChildLevel1.Text = NullToNothing(OE_RX_rs.Fields("PATIENT_TYPE").Value)
    oZEPNode.appendChild oZEPChildLevel1
    
    Set oZEPChildLevel1 = oDom.createElement("ZEP.30")
    oZEPChildLevel1.Text = NullToNothing(OE_RX_rs.Fields("PROCESS_PRIORITY").Value)
    oZEPNode.appendChild oZEPChildLevel1
    
    Set oZEPChildLevel1 = oDom.createElement("ZEP.31")
    oZEPChildLevel1.Text = NullToNothing(OE_RX_rs.Fields("NURSE_ACK_NAME").Value)
    oZEPNode.appendChild oZEPChildLevel1
    
    Set oZEPChildLevel1 = oDom.createElement("ZEP.32")
    strTemp = NullToNothing(OE_RX_rs.Fields("NURSE_ACK_DATE").Value)
    If IsValidDate(strTemp) Then
        oZEPChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZEPNode.appendChild oZEPChildLevel1
    
    Set oZEPChildLevel1 = oDom.createElement("ZEP.33")
    oZEPChildLevel1.Text = NullToNothing(OE_RX_rs.Fields("PREVIOUS_ORDER_STATUS").Value)
    oZEPNode.appendChild oZEPChildLevel1
    
    Set ReturnNewZEPNode_OE_RX = oZEPNode

Exit Function

ErrTrap:

   If gobjShared Is Nothing Then
      Set gobjShared = New CWShared
   End If
   gobjShared.ShowVBError Err.Number, Err.Description, "ReturnNewZEPNode_OE_RX Function", "DMBE", vbExclamation
   Resume Next 'There is no reason this procedure should not continue to create xml if it errors

End Function


Public Function ReturnNewZEPNode(ByVal lngCount As Long, ByVal PI_DataRS As ADODB.Recordset) As MSXML.IXMLDOMElement
' Encounter Patient Information Segment
' sample funcion output
'   <ZEP>
'      <ZEP.1>Count</ZEP.1>
'      <ZEP.2>Patient Info Data ID</ZEP.2>
'      <ZEP.3>Patient Info Medcin ID</ZEP.3>
'      <ZEP.4>Unit Number</ZEP.4>
'      <ZEP.5>Patient Information NCID</ZEP.5>
'      <ZEP.6>Encounter Number</ZEP.6>
'      <ZEP.7>Encounter Facility NCID</ZEP.7>
'      <ZEP.8>Patient Info Date/Time</ZEP.8>
'      <ZEP.9>Created By</ZEP.9>
'      <ZEP.10>Created On</ZEP.10>
'      <ZEP.11>Updated By</ZEP.11>
'      <ZEP.12>Updated On</ZEP.12>
'   </ZEP>

    Dim oDom As MSXML.DOMDocument
    Dim oZEPNode As MSXML.IXMLDOMElement
    Dim oZEPChildLevel1 As MSXML.IXMLDOMElement
    Dim strTemp As String
    
    On Error GoTo ErrTrap

    Set oDom = New MSXML.DOMDocument
    Set oZEPNode = oDom.createElement("ZEP")
    
    ' Set ID - Count
    Set oZEPChildLevel1 = oDom.createElement("ZEP.1")
    oZEPChildLevel1.Text = CStr(lngCount)
    oZEPNode.appendChild oZEPChildLevel1

    ' Patient Info Data ID - Table PI_Data:PI_Data_ID
    Set oZEPChildLevel1 = oDom.createElement("ZEP.2")
    oZEPChildLevel1.Text = NullToNothing(PI_DataRS.Fields("PI_Data_ID").Value)
    oZEPNode.appendChild oZEPChildLevel1

    ' Patient Info Medcin ID - Table PI_Data:PI_SnoID
    Set oZEPChildLevel1 = oDom.createElement("ZEP.3")
    oZEPChildLevel1.Text = NullToNothing(PI_DataRS.Fields("SnoID").Value)
    oZEPNode.appendChild oZEPChildLevel1
    
    ' Unit Number - Table PI_Data:Unit_Number
    Set oZEPChildLevel1 = oDom.createElement("ZEP.4")
    oZEPChildLevel1.Text = NullToNothing(PI_DataRS.Fields("Unit_Number").Value)
    oZEPNode.appendChild oZEPChildLevel1
    
    ' Patient Information - Table PI_Data:PI_NCID
    Set oZEPChildLevel1 = oDom.createElement("ZEP.5")
    oZEPChildLevel1.Text = NullToNothing(PI_DataRS.Fields("PI_NCID").Value)
    oZEPNode.appendChild oZEPChildLevel1
    
    ' Encounter Number - Table PI_Data:Enc_Num
    Set oZEPChildLevel1 = oDom.createElement("ZEP.6")
    oZEPChildLevel1.Text = NullToNothing(PI_DataRS.Fields("Enc_Num").Value)
    oZEPNode.appendChild oZEPChildLevel1
    
    ' Encounter Facility - Table PI_Data:Enc_Fac_NCID
    Set oZEPChildLevel1 = oDom.createElement("ZEP.7")
    oZEPChildLevel1.Text = NullToNothing(PI_DataRS.Fields("Enc_Fac_NCID").Value)
    oZEPNode.appendChild oZEPChildLevel1
    
    ' Patient Info Date/Time - Table PI_Data:PI_Time
    Set oZEPChildLevel1 = oDom.createElement("ZEP.8")
    strTemp = NullToNothing(PI_DataRS.Fields("PI_Time").Value)
    If IsValidDate(strTemp) Then
        oZEPChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZEPNode.appendChild oZEPChildLevel1
    
    ' Created By - Table PI_Data:CreatedBy
    Set oZEPChildLevel1 = oDom.createElement("ZEP.9")
    oZEPChildLevel1.Text = NullToNothing(PI_DataRS.Fields("CreatedBy").Value)
    oZEPNode.appendChild oZEPChildLevel1
    
    ' Created On - Table PI_Data:CreatedOn
    Set oZEPChildLevel1 = oDom.createElement("ZEP.10")
    strTemp = NullToNothing(PI_DataRS.Fields("CreatedOn").Value)
    If IsValidDate(strTemp) Then
        oZEPChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZEPNode.appendChild oZEPChildLevel1
    
    ' Updated By - Table PI_Data:UpdatedBy
    Set oZEPChildLevel1 = oDom.createElement("ZEP.11")
    oZEPChildLevel1.Text = NullToNothing(PI_DataRS.Fields("UpdatedBy").Value)
    oZEPNode.appendChild oZEPChildLevel1
    
    ' Updated On - Table PI_Data:UpdatedOn
    Set oZEPChildLevel1 = oDom.createElement("ZEP.12")
    strTemp = NullToNothing(PI_DataRS.Fields("UpdatedOn").Value)
    If IsValidDate(strTemp) Then
        oZEPChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZEPNode.appendChild oZEPChildLevel1
    
    Set ReturnNewZEPNode = oZEPNode

Exit Function

ErrTrap:
   If gobjShared Is Nothing Then
      Set gobjShared = New CWShared
   End If
   gobjShared.ShowVBError Err.Number, Err.Description, "ReturnNewZEPNode Function", "DMBE", vbExclamation
   Resume Next 'There is no reason this procedure should not continue to create xml if it errors

End Function

Public Function ReturnNewZEXNode(ByVal lngCount As Long, ByVal EncTextNotesRS As ADODB.Recordset) As MSXML.IXMLDOMElement
    ' Encounter Extended Notes Segment
' sample function output
'   <ZEX>
'      <ZEX.1>Count</ZEX.1>
'      <ZEX.2>DataID</ZEX.2>
'      <ZEX.3>Encounter Number</ZEX.3>
'      <ZEX.4>Facility NCID</ZEX.4>
'      <ZEX.5>
'         <XCN.1>Modifying Clinician SSN</XCN.1>
'         <XCN.2>Modifying Clinician Last Name</XCN.2>
'         <XCN.3>Modifying Clinician First Name</XCN.3>
'         <XCN.4>Modifying Clinician Middle Name</XCN.4>
'      </ZEX.5>
'      <ZEX.6>
'         <XCN.1>Create Clinician SSN</XCN.1>
'         <XCN.2>Create Clinician Last Name</XCN.2>
'         <XCN.3>Create Clinician First Name</XCN.3>
'         <XCN.4>Create Clinician Middle Name</XCN.4>
'      </ZEX.6>
'      <ZEX.7>Start Date/Time</ZEX.7>
'      <ZEX.8>RTF</ZEX.8>
'      <ZEX.9>Complete</ZEX.9>
'      <ZEX.10>Category</ZEX.10>
'      <ZEX.11>Title</ZEX.11>
'      <ZEX.12>End Date/Time</ZEX.12>
'      <ZEX.13>Created By</ZEX.13>
'      <ZEX.14>Created On</ZEX.14>
'      <ZEX.15>Updated By</ZEX.15>
'      <ZEX.16>Updated On</ZEX.16>
'   </ZEX>
    
    
    Dim oDom As MSXML.DOMDocument
    Dim oZEXNode As MSXML.IXMLDOMElement
    Dim oZEXChildLevel1 As MSXML.IXMLDOMElement
    Dim oZEXChildLevel2 As MSXML.IXMLDOMElement
    Dim strDocFirstName As String
    Dim strDocLastName As String
    Dim strDocMiddleName As String
    Dim strDocID As String
    Dim strDocSpecialty As String
    Dim strTemp As String
    Dim blnRsIsGood As Boolean
    
    On Error GoTo ErrTrap

    Set oDom = New MSXML.DOMDocument
    Set oZEXNode = oDom.createElement("ZEX")
    
    ' Set ID - Count
    Set oZEXChildLevel1 = oDom.createElement("ZEX.1")
    oZEXChildLevel1.Text = CStr(lngCount)
    oZEXNode.appendChild oZEXChildLevel1

    ' DataID - Table EncTextNotes:DataID
    Set oZEXChildLevel1 = oDom.createElement("ZEX.2")
    oZEXChildLevel1.Text = NullToNothing(EncTextNotesRS.Fields("DataID").Value)
    oZEXNode.appendChild oZEXChildLevel1
    
    ' Encounter Number - Table EncTextNotes:EncounterNumber
    Set oZEXChildLevel1 = oDom.createElement("ZEX.3")
    oZEXChildLevel1.Text = NullToNothing(EncTextNotesRS.Fields("EncounterNumber").Value)
    oZEXNode.appendChild oZEXChildLevel1
    
    ' Facility - Table EncTextNotes:FacilityNCID
    Set oZEXChildLevel1 = oDom.createElement("ZEX.4")
    oZEXChildLevel1.Text = NullToNothing(EncTextNotesRS.Fields("FacilityNCID").Value)
    oZEXNode.appendChild oZEXChildLevel1
    
    ' Modifying Clinician - Table EncTextNotes:ModifyClinicanNCID
    Set oZEXChildLevel1 = oDom.createElement("ZEX.5")
        strTemp = NullToNothing(EncTextNotesRS.Fields("ModifyClinicanNCID").Value)
        If strTemp <> "" Then
            Call GetNewDoctorInfo(CLng(strTemp), strDocLastName, strDocFirstName, strDocMiddleName, strDocID, strDocSpecialty)
            blnRsIsGood = True
        End If
        ' Modifying Clinician SSN
        Set oZEXChildLevel2 = oDom.createElement("XCN.1")
        If blnRsIsGood = True Then
            oZEXChildLevel2.Text = strDocID
        End If
        oZEXChildLevel1.appendChild oZEXChildLevel2
        ' Modifying Clinician Last Name
        Set oZEXChildLevel2 = oDom.createElement("XCN.2")
        If blnRsIsGood = True Then
            oZEXChildLevel2.Text = strDocLastName
        End If
        oZEXChildLevel1.appendChild oZEXChildLevel2
        ' Modifying Clinician First Name
        Set oZEXChildLevel2 = oDom.createElement("XCN.3")
        If blnRsIsGood = True Then
            oZEXChildLevel2.Text = strDocFirstName
        End If
        oZEXChildLevel1.appendChild oZEXChildLevel2
        ' Modifying Clinician Middle Name
        Set oZEXChildLevel2 = oDom.createElement("XCN.4")
        If blnRsIsGood = True Then
            oZEXChildLevel2.Text = strDocMiddleName
        End If
        oZEXChildLevel1.appendChild oZEXChildLevel2
    oZEXNode.appendChild oZEXChildLevel1
    
    ' Create Clinician - Table EncTextNotes:CreateClinicanNCID
    Set oZEXChildLevel1 = oDom.createElement("ZEX.6")
        strTemp = NullToNothing(EncTextNotesRS.Fields("CreateClinicanNCID").Value)
        blnRsIsGood = False
        If strTemp <> "" Then
            Call GetNewDoctorInfo(CLng(strTemp), strDocLastName, strDocFirstName, strDocMiddleName, strDocID, strDocSpecialty)
            blnRsIsGood = True
        End If
        ' Create Clinician SSN
        Set oZEXChildLevel2 = oDom.createElement("XCN.1")
        If blnRsIsGood = True Then
            oZEXChildLevel2.Text = strDocID
        End If
        oZEXChildLevel1.appendChild oZEXChildLevel2
        ' Create Clinician Last Name
        Set oZEXChildLevel2 = oDom.createElement("XCN.2")
        If blnRsIsGood = True Then
            oZEXChildLevel2.Text = strDocLastName
        End If
        oZEXChildLevel1.appendChild oZEXChildLevel2
        ' Create Clinician First Name
        Set oZEXChildLevel2 = oDom.createElement("XCN.3")
        If blnRsIsGood = True Then
            oZEXChildLevel2.Text = strDocFirstName
        End If
        oZEXChildLevel1.appendChild oZEXChildLevel2
        ' Create Clinician Middle Name
        Set oZEXChildLevel2 = oDom.createElement("XCN.4")
        If blnRsIsGood = True Then
            oZEXChildLevel2.Text = strDocMiddleName
        End If
        oZEXChildLevel1.appendChild oZEXChildLevel2
    oZEXNode.appendChild oZEXChildLevel1
    
    ' Start Date/Time - Table EncTextNotes:StartTime
    Set oZEXChildLevel1 = oDom.createElement("ZEX.7")
    strTemp = NullToNothing(EncTextNotesRS.Fields("StartTime").Value)
    If IsValidDate(strTemp) Then
        oZEXChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZEXNode.appendChild oZEXChildLevel1
    
    ' RTF - Table EncTextNotes:RTF
    Set oZEXChildLevel1 = oDom.createElement("ZEX.8")
    strTemp = NullToNothing(EncTextNotesRS.Fields("RTF").Value)
    If strTemp <> "" Then
          oZEXChildLevel1.Text = FilterRtfToXml(strTemp)
    End If
    oZEXNode.appendChild oZEXChildLevel1
    
    ' Complete - Table EncTextNotes:Complete
    Set oZEXChildLevel1 = oDom.createElement("ZEX.9")
    oZEXChildLevel1.Text = NullToNothing(EncTextNotesRS.Fields("Complete").Value)
    oZEXNode.appendChild oZEXChildLevel1
    
    ' Category - Table EncTextNotes:Category
    Set oZEXChildLevel1 = oDom.createElement("ZEX.10")
    oZEXChildLevel1.Text = NullToNothing(EncTextNotesRS.Fields("Category").Value)
    oZEXNode.appendChild oZEXChildLevel1
    
    ' Title - Table EncTextNotes:Title
    Set oZEXChildLevel1 = oDom.createElement("ZEX.11")
    oZEXChildLevel1.Text = NullToNothing(EncTextNotesRS.Fields("Title").Value)
    oZEXNode.appendChild oZEXChildLevel1
    
    ' End Date/Time - Table EncTextNotes:EndTime
    Set oZEXChildLevel1 = oDom.createElement("ZEX.12")
    strTemp = NullToNothing(EncTextNotesRS.Fields("EndTime").Value)
    If IsValidDate(strTemp) Then
        oZEXChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZEXNode.appendChild oZEXChildLevel1
    
    ' Created By - Table EncTextNotes:CreatedBy
    Set oZEXChildLevel1 = oDom.createElement("ZEX.13")
    oZEXChildLevel1.Text = NullToNothing(EncTextNotesRS.Fields("CreatedBy").Value)
    oZEXNode.appendChild oZEXChildLevel1
    
    ' Created On - Table EncTextNotes:CreatedOn
    Set oZEXChildLevel1 = oDom.createElement("ZEX.14")
    strTemp = NullToNothing(EncTextNotesRS.Fields("CreatedOn").Value)
    If IsValidDate(strTemp) Then
        oZEXChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZEXNode.appendChild oZEXChildLevel1
    
    ' Updated By - Table EncTextNotes:UpdatedBy
    Set oZEXChildLevel1 = oDom.createElement("ZEX.15")
    oZEXChildLevel1.Text = NullToNothing(EncTextNotesRS.Fields("UpdatedBy").Value)
    oZEXNode.appendChild oZEXChildLevel1
    
    ' Updated On - Table EncTextNotes:UpdatedOn
    Set oZEXChildLevel1 = oDom.createElement("ZEX.16")
    strTemp = NullToNothing(EncTextNotesRS.Fields("UpdatedOn").Value)
    If IsValidDate(strTemp) Then
        oZEXChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZEXNode.appendChild oZEXChildLevel1
    
    Set ReturnNewZEXNode = oZEXNode
    
    Set oZEXChildLevel1 = oDom.createElement("ZEX.17")
    strTemp = NullToNothing(EncTextNotesRS.Fields("DTS").Value)
    If IsValidDate(strTemp) Then
        oZEXChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZEXNode.appendChild oZEXChildLevel1
    

Exit Function

ErrTrap:
   If gobjShared Is Nothing Then
      Set gobjShared = New CWShared
   End If
   gobjShared.ShowVBError Err.Number, Err.Description, "ReturnNewZEXNode Function", "DMBE", vbExclamation
   Resume Next 'There is no reason this procedure should not continue to create xml if it errors
End Function

Public Function ReturnNewZESNode(ByVal lngCount As Long, ByVal Enc_SectionsRS As ADODB.Recordset) As MSXML.IXMLDOMElement
   ' Encounter Sections Segment
' sample function output
'   <ZES>
'      <ZES.1>Count</ZES.1>
'      <ZES.2>DataID</ZES.2>
'      <ZES.3>Encounter Number</ZES.3>
'      <ZES.4>Facility NCID</ZES.4>
'      <ZES.5>Encounter Sections Index</ZES.5>
'      <ZES.6>Date/Time Stamp</ZES.6>
'      <ZES.7>Document</ZES.7>
'      <ZES.8>
'         <CE.1>Owner Identifier</CE.1>
'         <CE.2>Owner Name</CE.2>
'      </ZES.8>
'      <ZES.9>Sensitivity Level</ZES.9>
'      <ZES.10>Status</ZES.10>
'      <ZES.11>Original Size</ZES.11>
'      <ZES.12>LabRTF</ZES.12>
'      <ZES.13>RadRTF</ZES.13>
'      <ZES.14>Created By</ZES.14>
'      <ZES.15>Created On</ZES.15>
'      <ZES.16>Updated By</ZES.16>
'      <ZES.17>Updated On</ZES.17>
'      <ZES.18>Complete</ZES.18>
'      <ZES.19>Category</ZES.19>
'      <ZES.20>Title</ZES.20>
'      <ZES.21>WSRFV</ZES.21>
'      <ZES.22>Data Type NCID</ZES.22>
'   </ZES>
   
    Dim oDom As MSXML.DOMDocument
    Dim oZESNode As MSXML.IXMLDOMElement
    Dim oZESChildLevel1 As MSXML.IXMLDOMElement
    Dim oZESChildLevel2 As MSXML.IXMLDOMElement
    Dim strTemp As String
    
    On Error GoTo ErrTrap

    Set oDom = New MSXML.DOMDocument
    Set oZESNode = oDom.createElement("ZES")
        
    ' Set ID - Count
    Set oZESChildLevel1 = oDom.createElement("ZES.1")
    oZESChildLevel1.Text = CStr(lngCount)
    oZESNode.appendChild oZESChildLevel1

    ' DataID - Table Enc_Sections:DataID
    Set oZESChildLevel1 = oDom.createElement("ZES.2")
    oZESChildLevel1.Text = NullToNothing(Enc_SectionsRS.Fields("DataID").Value)
    oZESNode.appendChild oZESChildLevel1
    
    ' Encounter Number - Table Enc_Sections:EncounterNumber
    Set oZESChildLevel1 = oDom.createElement("ZES.3")
    oZESChildLevel1.Text = NullToNothing(Enc_SectionsRS.Fields("EncounterNumber").Value)
    oZESNode.appendChild oZESChildLevel1
    
    ' Facility - Table Enc_Sections:FacilityNCID
    Set oZESChildLevel1 = oDom.createElement("ZES.4")
    oZESChildLevel1.Text = NullToNothing(Enc_SectionsRS.Fields("FacilityNCID").Value)
    oZESNode.appendChild oZESChildLevel1
    
    ' Encounter Sections Index - Table Enc_Sections:Enc_SectionsIndex
    Set oZESChildLevel1 = oDom.createElement("ZES.5")
    oZESChildLevel1.Text = NullToNothing(Enc_SectionsRS.Fields("Enc_SectionsIndex").Value)
    oZESNode.appendChild oZESChildLevel1
    
    ' Date/Time Stamp - Table Enc_Sections:DTS
    Set oZESChildLevel1 = oDom.createElement("ZES.6")
    strTemp = NullToNothing(Enc_SectionsRS.Fields("DTS").Value)
    If IsValidDate(strTemp) Then
        oZESChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZESNode.appendChild oZESChildLevel1
    
    ' Document - Table Enc_Sections:DOC
    Set oZESChildLevel1 = oDom.createElement("ZES.7")
    strTemp = NullToNothing(Enc_SectionsRS.Fields("DOC").Value)
    If strTemp <> "" Then
      '<< Begin:SCR #25718;   Developer: Brian Mowbray 07/18/2002 01:48 PM
      If Enc_SectionsRS.Fields("OriginalSize").Value > 1 Then
        oZESChildLevel1.Text = EncodeString(strTemp, True)
      Else
        oZESChildLevel1.Text = FilterRtfToXml(strTemp)
      End If
    '>> End: SCR #25718;
    End If
    oZESNode.appendChild oZESChildLevel1
    
    ' Owner
    Set oZESChildLevel1 = oDom.createElement("ZES.8")
        ' Identifier - Table Enc_Sections:OwnerNCID
        Set oZESChildLevel2 = oDom.createElement("CE.1")
        oZESChildLevel2.Text = NullToNothing(Enc_SectionsRS.Fields("OwnerNCID").Value)
        oZESChildLevel1.appendChild oZESChildLevel2
        ' Name - Table Enc_Sections:OwnerName
        Set oZESChildLevel2 = oDom.createElement("CE.2")
        oZESChildLevel2.Text = NullToNothing(Enc_SectionsRS.Fields("OwnerName").Value)
        oZESChildLevel1.appendChild oZESChildLevel2
    oZESNode.appendChild oZESChildLevel1
    
    ' Sensitivity Level - Table Enc_Sections:SensitivityLevel
    Set oZESChildLevel1 = oDom.createElement("ZES.9")
    oZESChildLevel1.Text = NullToNothing(Enc_SectionsRS.Fields("SensitivityLevel").Value)
    oZESNode.appendChild oZESChildLevel1
    
    ' Status - Table Enc_Sections:Status
    Set oZESChildLevel1 = oDom.createElement("ZES.10")
    oZESChildLevel1.Text = NullToNothing(Enc_SectionsRS.Fields("Status").Value)
    oZESNode.appendChild oZESChildLevel1
    
    ' Original Size - Table Enc_Sections:OriginalSize
    Set oZESChildLevel1 = oDom.createElement("ZES.11")
    oZESChildLevel1.Text = NullToNothing(Enc_SectionsRS.Fields("OriginalSize").Value)
    oZESNode.appendChild oZESChildLevel1
    
    ' LabRTF - Table Enc_Sections:LabRTF
    Set oZESChildLevel1 = oDom.createElement("ZES.12")
    strTemp = NullToNothing(Enc_SectionsRS.Fields("LabRTF").Value)
    If strTemp <> "" Then
        oZESChildLevel1.Text = FilterRtfToXml(strTemp)
    End If
    oZESNode.appendChild oZESChildLevel1
    
    ' RadRTF - Table Enc_Sections:RadRTF
    Set oZESChildLevel1 = oDom.createElement("ZES.13")
    strTemp = NullToNothing(Enc_SectionsRS.Fields("RadRTF").Value)
    If strTemp <> "" Then
        oZESChildLevel1.Text = FilterRtfToXml(strTemp)
    End If
    oZESNode.appendChild oZESChildLevel1
    
    ' Created By - Table Enc_Sections:CreatedBy
    Set oZESChildLevel1 = oDom.createElement("ZES.14")
    oZESChildLevel1.Text = NullToNothing(Enc_SectionsRS.Fields("CreatedBy").Value)
    oZESNode.appendChild oZESChildLevel1
    
    ' Created On - Table Enc_Sections:CreatedOn
    Set oZESChildLevel1 = oDom.createElement("ZES.15")
    strTemp = NullToNothing(Enc_SectionsRS.Fields("CreatedOn").Value)
    If IsValidDate(strTemp) Then
        oZESChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZESNode.appendChild oZESChildLevel1
    
    ' Updated By - Table Enc_Sections:UpdatedBy
    Set oZESChildLevel1 = oDom.createElement("ZES.16")
    oZESChildLevel1.Text = NullToNothing(Enc_SectionsRS.Fields("UpdatedBy").Value)
    oZESNode.appendChild oZESChildLevel1
    
    ' Updated On - Table Enc_Sections:UpdatedOn
    Set oZESChildLevel1 = oDom.createElement("ZES.17")
    strTemp = NullToNothing(Enc_SectionsRS.Fields("UpdatedOn").Value)
    If IsValidDate(strTemp) Then
        oZESChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZESNode.appendChild oZESChildLevel1
    
    ' Complete - Table Enc_Sections:Complete
    Set oZESChildLevel1 = oDom.createElement("ZES.18")
    oZESChildLevel1.Text = NullToNothing(Enc_SectionsRS.Fields("Complete").Value)
    oZESNode.appendChild oZESChildLevel1
    
    ' Category - Table Enc_Sections:Category
    Set oZESChildLevel1 = oDom.createElement("ZES.19")
    oZESChildLevel1.Text = NullToNothing(Enc_SectionsRS.Fields("Category").Value)
    oZESNode.appendChild oZESChildLevel1
    
    ' Title - Table Enc_Sections:Title
    Set oZESChildLevel1 = oDom.createElement("ZES.20")
    oZESChildLevel1.Text = NullToNothing(Enc_SectionsRS.Fields("Title").Value)
    oZESNode.appendChild oZESChildLevel1
    
    ' WSRFV - Table Enc_Sections:WSRFV
    Set oZESChildLevel1 = oDom.createElement("ZES.21")
    oZESChildLevel1.Text = NullToNothing(Enc_SectionsRS.Fields("WSRFV").Value)
    oZESNode.appendChild oZESChildLevel1
    
    ' Date Type - Table Enc_Sections:DataTypeNCID
    Set oZESChildLevel1 = oDom.createElement("ZES.22")
    strTemp = NullToNothing(Enc_SectionsRS.Fields("DataTypeNCID").Value)
        
        Set oZESChildLevel2 = oDom.createElement("CE.1")
        oZESChildLevel1.appendChild oZESChildLevel2
        
        Set oZESChildLevel2 = oDom.createElement("CE.2")
        If strTemp <> "" Then
            oZESChildLevel2.Text = GetNCIDDescription(CLng(strTemp))
        End If
        oZESChildLevel1.appendChild oZESChildLevel2
        
        Set oZESChildLevel2 = oDom.createElement("CE.3")
        oZESChildLevel1.appendChild oZESChildLevel2
        Set oZESChildLevel2 = oDom.createElement("CE.4")
        oZESChildLevel1.appendChild oZESChildLevel2
        Set oZESChildLevel2 = oDom.createElement("CE.5")
        oZESChildLevel2.Text = strTemp 'The NCID
        oZESChildLevel1.appendChild oZESChildLevel2
        Set oZESChildLevel2 = oDom.createElement("CE.6")
        oZESChildLevel2.Text = "99NCI"
        oZESChildLevel1.appendChild oZESChildLevel2
    oZESNode.appendChild oZESChildLevel1
    
    'Set oZESChildLevel1 = oDom.createElement("ZES.23")
    'oZESChildLevel1.Text = NullToNothing(Enc_SectionsRS.Fields("UNIT_NUMBER").Value)
    'oZESNode.appendChild oZESChildLevel1
    
    
    Set ReturnNewZESNode = oZESNode

Exit Function

ErrTrap:
   If gobjShared Is Nothing Then
      Set gobjShared = New CWShared
   End If
   gobjShared.ShowVBError Err.Number, Err.Description, "ReturnNewZESNode Function", "DMBE", vbExclamation
   Resume Next 'There is no reason this procedure should not continue to create xml if it errors
End Function

Public Function ReturnNewZP1Node(ByVal lngCount As Long, ByVal ProceduresRS As ADODB.Recordset) As MSXML.IXMLDOMElement

' Additional Procedure Data Segment
' sample function output
'   <ZP1>
'      <ZP1.1>Count</ZP1.1>
'      <ZP1.2>DataID</ZP1.2>
'      <ZP1.3>CPTID</ZP1.3>
'      <ZP1.4>Point of Care Facility NCID</ZP1.4>
'      <ZP1.5>Status</ZP1.5>
'      <ZP1.6>Noted Date</ZP1.6>
'      <ZP1.7>Problem List Indicator</ZP1.7>
'      <ZP1.8>Procedures Comment</ZP1.8>
'      <ZP1.9>Create User</ZP1.9>
'      <ZP1.10>Create Date/Time</ZP1.10>
'      <ZP1.11>Modify User</ZP1.11>
'      <ZP1.12>Modify Date/Time</ZP1.12>
'      <ZP1.13>Update Flag</ZP1.13>
'      <ZP1.14>CDR Data ID</ZP1.14>
'      <ZP1.15>Created By</ZP1.15>
'      <ZP1.16>Created On</ZP1.16>
'      <ZP1.17>Updated By</ZP1.17>
'      <ZP1.18>Updated On</ZP1.18>
'   </ZP1>

    Dim oDom As MSXML.DOMDocument
    Dim oZP1Node As MSXML.IXMLDOMElement
    Dim oZP1ChildLevel1 As MSXML.IXMLDOMElement
    Dim oZP1ChildLevel2 As MSXML.IXMLDOMElement
    Dim strTemp As String
    
    On Error GoTo ErrTrap

    Set oDom = New MSXML.DOMDocument
    Set oZP1Node = oDom.createElement("ZP1")
        
    ' Set ID - Count
    Set oZP1ChildLevel1 = oDom.createElement("ZP1.1")
    oZP1ChildLevel1.Text = CStr(lngCount)
    oZP1Node.appendChild oZP1ChildLevel1

    ' DataID - Table Procedures:DataID - Table Procedures:DataID
    Set oZP1ChildLevel1 = oDom.createElement("ZP1.2")
    oZP1ChildLevel1.Text = NullToNothing(ProceduresRS.Fields("DataID").Value)
    oZP1Node.appendChild oZP1ChildLevel1

    ' CPTID - Table Procedures:CPTID
    Set oZP1ChildLevel1 = oDom.createElement("ZP1.3")
    oZP1ChildLevel1.Text = NullToNothing(ProceduresRS.Fields("CPTID").Value)
    oZP1Node.appendChild oZP1ChildLevel1

    ' Point of Care Facility - Table Procedures:PointOfCareFacilityNCID
    Set oZP1ChildLevel1 = oDom.createElement("ZP1.4")
    strTemp = NullToNothing(ProceduresRS.Fields("PointOfCareFacilityNCID").Value)
        
        Set oZP1ChildLevel2 = oDom.createElement("CE.1")
        oZP1ChildLevel1.appendChild oZP1ChildLevel2
        
        Set oZP1ChildLevel2 = oDom.createElement("CE.2")
        If strTemp <> "" Then
            oZP1ChildLevel2.Text = GetNCIDDescription(CLng(strTemp))
        End If
        oZP1ChildLevel1.appendChild oZP1ChildLevel2
        
        Set oZP1ChildLevel2 = oDom.createElement("CE.3")
        oZP1ChildLevel1.appendChild oZP1ChildLevel2
        Set oZP1ChildLevel2 = oDom.createElement("CE.4")
        oZP1ChildLevel1.appendChild oZP1ChildLevel2
        Set oZP1ChildLevel2 = oDom.createElement("CE.5")
        oZP1ChildLevel2.Text = strTemp 'The NCID
        oZP1ChildLevel1.appendChild oZP1ChildLevel2
        Set oZP1ChildLevel2 = oDom.createElement("CE.6")
        oZP1ChildLevel2.Text = "99NCI"
        oZP1ChildLevel1.appendChild oZP1ChildLevel2
    oZP1Node.appendChild oZP1ChildLevel1


    ' Status - Table Procedures:StatusNCID
    Set oZP1ChildLevel1 = oDom.createElement("ZP1.5")
    'strTemp = NullToNothing(ProceduresRS.Fields("StatusNCID").Value)
    'If strTemp <> "" Then
    '    oZP1ChildLevel1.Text = GetNCIDDescription(CLng(strTemp))
    'End If
    strTemp = NullToNothing(ProceduresRS.Fields("StatusNCID").Value)
        
        Set oZP1ChildLevel2 = oDom.createElement("CE.1")
        oZP1ChildLevel1.appendChild oZP1ChildLevel2
        
        Set oZP1ChildLevel2 = oDom.createElement("CE.2")
        If strTemp <> "" Then
            oZP1ChildLevel2.Text = GetNCIDDescription(CLng(strTemp))
        End If
        oZP1ChildLevel1.appendChild oZP1ChildLevel2
        
        Set oZP1ChildLevel2 = oDom.createElement("CE.3")
        oZP1ChildLevel1.appendChild oZP1ChildLevel2
        Set oZP1ChildLevel2 = oDom.createElement("CE.4")
        oZP1ChildLevel1.appendChild oZP1ChildLevel2
        Set oZP1ChildLevel2 = oDom.createElement("CE.5")
        oZP1ChildLevel2.Text = strTemp 'The NCID
        oZP1ChildLevel1.appendChild oZP1ChildLevel2
        Set oZP1ChildLevel2 = oDom.createElement("CE.6")
        oZP1ChildLevel2.Text = "99NCI"
        oZP1ChildLevel1.appendChild oZP1ChildLevel2
    oZP1Node.appendChild oZP1ChildLevel1
    'oZP1Node.appendChild oZP1ChildLevel1
    
    ' Noted Date - Table Procedures:NotedDate
    Set oZP1ChildLevel1 = oDom.createElement("ZP1.6")
    strTemp = NullToNothing(ProceduresRS.Fields("NotedDate").Value)
    If IsValidDate(strTemp) Then
        oZP1ChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZP1Node.appendChild oZP1ChildLevel1

    ' Problem List Indicator - Table Procedures:ProblemListIndicator
    Set oZP1ChildLevel1 = oDom.createElement("ZP1.7")
    oZP1ChildLevel1.Text = NullToNothing(ProceduresRS.Fields("ProblemListIndicator").Value)
    oZP1Node.appendChild oZP1ChildLevel1

    ' Procedures Comment - Table Procedures:ProceduresComment
    Set oZP1ChildLevel1 = oDom.createElement("ZP1.8")
    oZP1ChildLevel1.Text = NullToNothing(ProceduresRS.Fields("ProceduresComment").Value)
    oZP1Node.appendChild oZP1ChildLevel1

    ' Create User - Table Procedures:CreateUserNCID
    Set oZP1ChildLevel1 = oDom.createElement("ZP1.9")
    oZP1ChildLevel1.Text = NullToNothing(ProceduresRS.Fields("CreateUserNCID").Value)
    oZP1Node.appendChild oZP1ChildLevel1

    ' Create Date/Time - Table Procedures:CreateTime
    Set oZP1ChildLevel1 = oDom.createElement("ZP1.10")
    strTemp = NullToNothing(ProceduresRS.Fields("CreateTime").Value)
    If IsValidDate(strTemp) Then
        oZP1ChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZP1Node.appendChild oZP1ChildLevel1

    ' Modify User - Table Procedures:ModifyUserNCID
    Set oZP1ChildLevel1 = oDom.createElement("ZP1.11")
    oZP1ChildLevel1.Text = NullToNothing(ProceduresRS.Fields("ModifyUserNCID").Value)
    oZP1Node.appendChild oZP1ChildLevel1
    
    ' Modify Date/Time - Table Procedures:ModifyTime
    Set oZP1ChildLevel1 = oDom.createElement("ZP1.12")
    strTemp = NullToNothing(ProceduresRS.Fields("ModifyTime").Value)
    If IsValidDate(strTemp) Then
        oZP1ChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZP1Node.appendChild oZP1ChildLevel1
    
    ' Update Flag - Table Procedures:UpdateFlag
    Set oZP1ChildLevel1 = oDom.createElement("ZP1.13")
    oZP1ChildLevel1.Text = NullToNothing(ProceduresRS.Fields("UpdateFlag").Value)
    oZP1Node.appendChild oZP1ChildLevel1

    ' CDR Data ID - Table Procedures:CDR_DataID
    Set oZP1ChildLevel1 = oDom.createElement("ZP1.14")
    oZP1ChildLevel1.Text = NullToNothing(ProceduresRS.Fields("CDR_DataID").Value)
    oZP1Node.appendChild oZP1ChildLevel1

    ' Created By - Table Procedures:CreatedBy
    Set oZP1ChildLevel1 = oDom.createElement("ZP1.15")
    oZP1ChildLevel1.Text = NullToNothing(ProceduresRS.Fields("CreatedBy").Value)
    oZP1Node.appendChild oZP1ChildLevel1

    ' Created On - Table Procedures:CreatedOn
    Set oZP1ChildLevel1 = oDom.createElement("ZP1.16")
    strTemp = NullToNothing(ProceduresRS.Fields("CreatedOn").Value)
    If IsValidDate(strTemp) Then
        oZP1ChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZP1Node.appendChild oZP1ChildLevel1
    
    ' Updated By - Table Procedures:UpdatedBy
    Set oZP1ChildLevel1 = oDom.createElement("ZP1.17")
    oZP1ChildLevel1.Text = NullToNothing(ProceduresRS.Fields("UpdatedBy").Value)
    oZP1Node.appendChild oZP1ChildLevel1

    ' Updated On - Table Procedures:UpdatedOn
    Set oZP1ChildLevel1 = oDom.createElement("ZP1.18")
    strTemp = NullToNothing(ProceduresRS.Fields("UpdatedOn").Value)
    If IsValidDate(strTemp) Then
        oZP1ChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZP1Node.appendChild oZP1ChildLevel1
    
    ' Unit Number - Table Procedures:Unit_Number
    Set oZP1ChildLevel1 = oDom.createElement("ZP1.19")
    oZP1ChildLevel1.Text = NullToNothing(ProceduresRS.Fields("UNIT_NUMBER").Value)
    oZP1Node.appendChild oZP1ChildLevel1

    ' Facility - Table Procedures:FacilityNCID
    Set oZP1ChildLevel1 = oDom.createElement("Zp1.20")
    strTemp = NullToNothing(ProceduresRS.Fields("FACILITYNCID").Value)
    Set oZP1ChildLevel2 = oDom.createElement("CE.1")
        oZP1ChildLevel1.appendChild oZP1ChildLevel2
        Set oZP1ChildLevel2 = oDom.createElement("CE.2")
        If strTemp <> "" Then
            oZP1ChildLevel2.Text = GetNCIDDescription(CLng(strTemp))
        End If
        oZP1ChildLevel1.appendChild oZP1ChildLevel2
        Set oZP1ChildLevel2 = oDom.createElement("CE.3")
        oZP1ChildLevel1.appendChild oZP1ChildLevel2
        Set oZP1ChildLevel2 = oDom.createElement("CE.4")
        oZP1ChildLevel1.appendChild oZP1ChildLevel2
        Set oZP1ChildLevel2 = oDom.createElement("CE.5")
        oZP1ChildLevel2.Text = strTemp
        oZP1ChildLevel1.appendChild oZP1ChildLevel2
        Set oZP1ChildLevel2 = oDom.createElement("CE.6")
        oZP1ChildLevel2.Text = "99NCI"
        oZP1ChildLevel1.appendChild oZP1ChildLevel2
    oZP1Node.appendChild oZP1ChildLevel1
    
    Set ReturnNewZP1Node = oZP1Node

Exit Function

ErrTrap:
   If gobjShared Is Nothing Then
      Set gobjShared = New CWShared
   End If
   gobjShared.ShowVBError Err.Number, Err.Description, "ReturnNewZP1Node Function", "DMBE", vbExclamation
   Resume Next 'There is no reason this procedure should not continue to create xml if it errors
   Exit Function
   Resume
End Function

Public Function ReturnNewPR1Node(ByVal lngCount As Long, ByVal ProceduresRS As ADODB.Recordset) As MSXML.IXMLDOMElement
   ' Procedure Segment
' sample function output
'   <PR1>
'      <PR1.1>1</PR1.1>
'      <PR1.3>
'         <CE.1></CE.1>
'         <CE.2>Anesthesia Pelvis For Body Cast Application Or Revision</CE.2>
'         <CE.3>C4</CE.3>
'         <CE.4></CE.4>
'         <CE.5>Anesthesia Pelvis For Body Cast Application Or Revision</CE.5>
'         <CE.6>99NCI</CE.6>
'      </PR1.3>
'      <PR1.5>20010812142954</PR1.5>
'      <PR1.12>
'         <XCN.1>123459876</XCN.1>
'         <XCN.2>SYSTEM</XCN.2>
'         <XCN.3>GENERATED</XCN.3>
'         <XCN.4></XCN.4>
'      </PR1.12>
'   </PR1>

    Dim oDom As MSXML.DOMDocument
    Dim oPR1Node As MSXML.IXMLDOMElement
    Dim oPR1ChildLevel1 As MSXML.IXMLDOMElement
    Dim oPR1ChildLevel2 As MSXML.IXMLDOMElement
    Dim oPR1ChildLevel3 As MSXML.IXMLDOMElement
    Dim strDocFirstName As String
    Dim strDocLastName As String
    Dim strDocMiddleName As String
    Dim strDocID As String
    Dim strDocSpecialty As String
    Dim strTemp As String
    Dim blnRsIsGood As Boolean
    
    On Error GoTo ErrTrap

    Set oDom = New MSXML.DOMDocument
    Set oPR1Node = oDom.createElement("PR1")
        
    ' Set ID - Count
    Set oPR1ChildLevel1 = oDom.createElement("PR1.1")
    oPR1ChildLevel1.Text = CStr(lngCount)
    oPR1Node.appendChild oPR1ChildLevel1
    
    ' Procedure Code
    Set oPR1ChildLevel1 = oDom.createElement("PR1.3")
        
        ' Procedure Code Identifier (CPT4) - Table Procedures:CPT4ChangableCode
        Set oPR1ChildLevel2 = oDom.createElement("CE.1")
        
        strTemp = NullToNothing(ProceduresRS.Fields("CPTID").Value)
        If strTemp <> "" Then
            If mobjConvertSno Is Nothing Then
                Set mobjConvertSno = New actxSearchSnoMed.IConvert
            End If
            oPR1ChildLevel2.Text = mobjConvertSno.SnoCPT4(CLng(strTemp))
        End If
        oPR1ChildLevel1.appendChild oPR1ChildLevel2
        
        'If strTemp <> "" Then
        '    If mobjConvertSno Is Nothing Then
        '        Set mobjConvertSno = New actxSearchSnoMed.IConvert
        '    End If
        '    oPR1ChildLevel2.Text = mobjConvertSno.SnoCPT4(CLng(strTemp))
        'End If
        'oPR1ChildLevel1.appendChild oPR1ChildLevel2
        
        ' CPT-4 CODE Description
        Set oPR1ChildLevel2 = oDom.createElement("CE.2")
        
        'strTemp = NullToNothing(ProceduresRS.Fields("CPT4ChangableCode").Value)
        '<SCR 22597 CC>
        strTemp = NullToNothing(ProceduresRS.Fields("CPTID").Value)
        If strTemp <> "" Then
            If mobjConvertSno Is Nothing Then
                Set mobjConvertSno = New actxSearchSnoMed.IConvert
            End If
            oPR1ChildLevel2.Text = mobjConvertSno.CPTDescription(CLng(strTemp)) '  .SnoCPT4(CLng(strTemp))
        End If
        oPR1ChildLevel1.appendChild oPR1ChildLevel2
        
        Set oPR1ChildLevel2 = oDom.createElement("CE.3")
        oPR1ChildLevel2.Text = "C4"
        oPR1ChildLevel1.appendChild oPR1ChildLevel2
        
        ' Procedure Code Identifier (CPT4) - Table Procedures:CPT4ChangableCode
        Set oPR1ChildLevel2 = oDom.createElement("CE.4")
        oPR1ChildLevel2.Text = NullToNothing(ProceduresRS.Fields("CPTID").Value)
        oPR1ChildLevel1.appendChild oPR1ChildLevel2
        
        ' Procedure Code Text - Table Procedures:ProcedureNCID
        Set oPR1ChildLevel2 = oDom.createElement("CE.5")
        strTemp = NullToNothing(ProceduresRS.Fields("CPTID").Value)
        If strTemp <> "" Then
            If mobjConvertSno Is Nothing Then
                Set mobjConvertSno = New actxSearchSnoMed.IConvert
            End If
            oPR1ChildLevel2.Text = mobjConvertSno.CPTDescription(CLng(strTemp)) '  .SnoCPT4(CLng(strTemp))
        End If
        oPR1ChildLevel1.appendChild oPR1ChildLevel2
        
        '</SCR 22597 CC>
        
        Set oPR1ChildLevel2 = oDom.createElement("CE.6")
        oPR1ChildLevel2.Text = "99NCI"
        oPR1ChildLevel1.appendChild oPR1ChildLevel2
        
    oPR1Node.appendChild oPR1ChildLevel1
    
    ' Procedure Date/Time - Table Procedures:OnSetDate
    Set oPR1ChildLevel1 = oDom.createElement("PR1.5")
    strTemp = NullToNothing(ProceduresRS.Fields("OnSetDate").Value)
    If IsValidDate(strTemp) Then
        oPR1ChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oPR1Node.appendChild oPR1ChildLevel1
    
    '<< Begin:SCR #26641;   Developer: Brian Mowbray 08/15/2002 06:37 PM
    ' Procedure Practitioner - Table Procedures:ClinicianNCID
    Set oPR1ChildLevel1 = oDom.createElement("PR1.12.LST")
          
        Set oPR1ChildLevel2 = oDom.createElement("PR1.12")
            
            blnRsIsGood = False
            strTemp = NullToNothing(ProceduresRS.Fields("ClinicianNCID").Value)
            If strTemp <> "" Then
                Call GetNewDoctorInfo(CLng(strTemp), strDocLastName, strDocFirstName, strDocMiddleName, strDocID, strDocSpecialty)
                blnRsIsGood = True
            End If
            
            ' Procedure Practitioner Identifier Number (SSN)
            Set oPR1ChildLevel3 = oDom.createElement("XCN.1")
            If blnRsIsGood Then
                oPR1ChildLevel3.Text = strDocID
            End If
            oPR1ChildLevel2.appendChild oPR1ChildLevel3
            ' Procedure Practitioner Last Name
            Set oPR1ChildLevel3 = oDom.createElement("XCN.2")
            If blnRsIsGood Then
                oPR1ChildLevel3.Text = strDocLastName
            End If
            oPR1ChildLevel2.appendChild oPR1ChildLevel3
            ' Procedure Practitioner First Name
            Set oPR1ChildLevel3 = oDom.createElement("XCN.3")
            If blnRsIsGood Then
                oPR1ChildLevel3.Text = strDocFirstName
            End If
            oPR1ChildLevel2.appendChild oPR1ChildLevel3
            ' Procedure Practitioner Middle Name
            Set oPR1ChildLevel3 = oDom.createElement("XCN.4")
            If blnRsIsGood Then
                oPR1ChildLevel3.Text = strDocMiddleName
            End If
            oPR1ChildLevel2.appendChild oPR1ChildLevel3
        
        oPR1ChildLevel1.appendChild oPR1ChildLevel2
    
    oPR1Node.appendChild oPR1ChildLevel1
    '>> End: SCR #26641;
    
    Set ReturnNewPR1Node = oPR1Node
    
Exit Function

ErrTrap:
   If gobjShared Is Nothing Then
      Set gobjShared = New CWShared
   End If
   gobjShared.ShowVBError Err.Number, Err.Description, "ReturnNewPR1Node Function", "DMBE", vbExclamation
   'Err.Clear
   Resume Next 'There is no reason this procedure should not continue to create xml if it errors
   Exit Function
   Resume
End Function
   
Public Function ReturnNewZEDNode(ByVal lngCount As Long, ByVal Enc_DiagnosisRS As ADODB.Recordset) As MSXML.IXMLDOMElement
' Encounter Diagnosis Segment
' sample function output
'   <ZED>
'      <ZED.1>Count</ZED.1>
'      <ZED.2>Diagnosis</ZED.2>
'      <ZED.3>Facility NCID</ZED.3>
'      <ZED.4>Encounter Number</ZED.4>
'      <ZED.5>Created By</ZED.5>
'      <ZED.6>Created On</ZED.6>
'      <ZED.7>Updated By</ZED.7>
'      <ZED.8>Updated On</ZED.8>
'   </ZED>
    
    Dim oDom As MSXML.DOMDocument
    Dim oZEDNode As MSXML.IXMLDOMElement
    Dim oZEDChildLevel1 As MSXML.IXMLDOMElement
    Dim oZEDChildLevel2 As MSXML.IXMLDOMElement
    Dim strTemp As String
    
    On Error GoTo ErrTrap
    
    Set oDom = New MSXML.DOMDocument
    Set oZEDNode = oDom.createElement("ZED")
        
    ' Set ID - Count
    Set oZEDChildLevel1 = oDom.createElement("ZED.1")
    oZEDChildLevel1.Text = CStr(lngCount)
    oZEDNode.appendChild oZEDChildLevel1
    
    ' Diagnosis - Table Enc_Diagnosis:DiagnosisNCID
    Set oZEDChildLevel1 = oDom.createElement("ZED.2")
    strTemp = NullToNothing(Enc_DiagnosisRS.Fields("DiagnosisNCID").Value)

'    If strTemp <> "" Then
'        oZEDChildLevel1.Text = GetNCIDDescription(CLng(strTemp))
'    End If
    
        Set oZEDChildLevel2 = oDom.createElement("CE.1")
        oZEDChildLevel1.appendChild oZEDChildLevel2
        
        Set oZEDChildLevel2 = oDom.createElement("CE.2")
        If strTemp <> "" Then
            ' Begin:SCR #22812;   Developer: MAG 07/18/2002 06:00 PM
            oZEDChildLevel2.Text = GetDiagnosisDescription(CLng(strTemp))
            ' End: SCR #22812;
        End If
        oZEDChildLevel1.appendChild oZEDChildLevel2
        
        Set oZEDChildLevel2 = oDom.createElement("CE.3")
        oZEDChildLevel1.appendChild oZEDChildLevel2
        Set oZEDChildLevel2 = oDom.createElement("CE.4")
        oZEDChildLevel1.appendChild oZEDChildLevel2
        Set oZEDChildLevel2 = oDom.createElement("CE.5")
        oZEDChildLevel2.Text = strTemp 'The NCID
        oZEDChildLevel1.appendChild oZEDChildLevel2
        Set oZEDChildLevel2 = oDom.createElement("CE.6")
        oZEDChildLevel2.Text = "99NCI"
        oZEDChildLevel1.appendChild oZEDChildLevel2
    
    oZEDNode.appendChild oZEDChildLevel1
    
    
    ' Facility - Table Enc_Diagnosis:FacilityNCID
    Set oZEDChildLevel1 = oDom.createElement("ZED.3")
    oZEDChildLevel1.Text = NullToNothing(Enc_DiagnosisRS.Fields("FacilityNCID").Value)
    oZEDNode.appendChild oZEDChildLevel1
    
    ' Encounter Number - Table Enc_Diagnosis:EncounterNumber
    Set oZEDChildLevel1 = oDom.createElement("ZED.4")
    oZEDChildLevel1.Text = NullToNothing(Enc_DiagnosisRS.Fields("EncounterNumber").Value)
    oZEDNode.appendChild oZEDChildLevel1
    
    ' Created By - Table Enc_Diagnosis:CreatedBy
    Set oZEDChildLevel1 = oDom.createElement("ZED.5")
    oZEDChildLevel1.Text = NullToNothing(Enc_DiagnosisRS.Fields("CreatedBy").Value)
    oZEDNode.appendChild oZEDChildLevel1
    
    ' Created On - Table Enc_Diagnosis:CreatedOn
    Set oZEDChildLevel1 = oDom.createElement("ZED.6")
    strTemp = NullToNothing(Enc_DiagnosisRS.Fields("CreatedOn").Value)
    If IsValidDate(strTemp) Then
        oZEDChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZEDNode.appendChild oZEDChildLevel1
    
    ' Updated By - Table Enc_Diagnosis:UpdatedBy
    Set oZEDChildLevel1 = oDom.createElement("ZED.7")
    oZEDChildLevel1.Text = NullToNothing(Enc_DiagnosisRS.Fields("UpdatedBy").Value)
    oZEDNode.appendChild oZEDChildLevel1
    
    ' Updated On - Table Enc_Diagnosis:UpdatedOn
    Set oZEDChildLevel1 = oDom.createElement("ZED.8")
    strTemp = NullToNothing(Enc_DiagnosisRS.Fields("UpdatedOn").Value)
    If IsValidDate(strTemp) Then
        oZEDChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZEDNode.appendChild oZEDChildLevel1
    
    Set ReturnNewZEDNode = oZEDNode

Exit Function

ErrTrap:
   If gobjShared Is Nothing Then
      Set gobjShared = New CWShared
   End If
   gobjShared.ShowVBError Err.Number, Err.Description, "ReturnNewZEDNode Function", "DMBE", vbExclamation
   Resume Next 'There is no reason this procedure should not continue to create xml if it errors
End Function

Public Function ReturnNewZDRNode(ByVal lngCount As Long, ByVal DiagnosesRelationshipsRS As ADODB.Recordset) As MSXML.IXMLDOMElement
' Diagnosis Relationship Segment
' sample function output
'   <ZDR>
'      <ZDR.1>Count</ZDR.1>
'      <ZDR.2>Diagnosis Data ID</ZDR.2>
'      <ZDR.3>Related Data ID</ZDR.3>
'      <ZDR.4>Relationship NCID</ZDR.4>
'      <ZDR.5>Created By</ZDR.5>
'      <ZDR.6>Created On</ZDR.6>
'      <ZDR.7>Updated By</ZDR.7>
'      <ZDR.8>Updated On</ZDR.8>
'   </ZDR>
    
    Dim oDom As MSXML.DOMDocument
    Dim oZDRNode As MSXML.IXMLDOMElement
    Dim oZDRChildLevel1 As MSXML.IXMLDOMElement
    Dim strTemp As String
    
    On Error GoTo ErrTrap
    
    Set oDom = New MSXML.DOMDocument
    Set oZDRNode = oDom.createElement("ZDR")
        
    ' Set ID - Count
    Set oZDRChildLevel1 = oDom.createElement("ZDR.1")
    oZDRChildLevel1.Text = CStr(lngCount)
    oZDRNode.appendChild oZDRChildLevel1
    
    ' Diagnosis Data ID - Table DiagnosesRelationships:DiagnosisDataID
    Set oZDRChildLevel1 = oDom.createElement("ZDR.2")
    oZDRChildLevel1.Text = NullToNothing(DiagnosesRelationshipsRS.Fields("DiagnosisDataID").Value)
    oZDRNode.appendChild oZDRChildLevel1
    
    ' Related Data ID - Table DiagnosesRelationships:RelatedDataID
    Set oZDRChildLevel1 = oDom.createElement("ZDR.3")
    oZDRChildLevel1.Text = NullToNothing(DiagnosesRelationshipsRS.Fields("RelatedDataID").Value)
    oZDRNode.appendChild oZDRChildLevel1
    
    ' Relationship NCID - Table DiagnosesRelationships:RelationshipNCID
    Set oZDRChildLevel1 = oDom.createElement("ZDR.4")
    strTemp = NullToNothing(DiagnosesRelationshipsRS.Fields("RelationshipNCID").Value)
    If strTemp <> "" Then
        oZDRChildLevel1.Text = GetNCIDDescription(CLng(strTemp))
    End If
    oZDRNode.appendChild oZDRChildLevel1
    
    ' Created By - Table DiagnosesRelationships:CreatedBy
    Set oZDRChildLevel1 = oDom.createElement("ZDR.5")
    oZDRChildLevel1.Text = NullToNothing(DiagnosesRelationshipsRS.Fields("CreatedBy").Value)
    oZDRNode.appendChild oZDRChildLevel1
    
    ' Created On - Table DiagnosesRelationships:CreatedOn
    Set oZDRChildLevel1 = oDom.createElement("ZDR.6")
    strTemp = NullToNothing(DiagnosesRelationshipsRS.Fields("CreatedOn").Value)
    If IsValidDate(strTemp) Then
        oZDRChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZDRNode.appendChild oZDRChildLevel1
    
    ' Updated By - Table DiagnosesRelationships:UpdatedBy
    Set oZDRChildLevel1 = oDom.createElement("ZDR.7")
    oZDRChildLevel1.Text = NullToNothing(DiagnosesRelationshipsRS.Fields("UpdatedBy").Value)
    oZDRNode.appendChild oZDRChildLevel1
    
    ' Updated On - Table DiagnosesRelationships:UpdatedOn
    Set oZDRChildLevel1 = oDom.createElement("ZDR.8")
    strTemp = NullToNothing(DiagnosesRelationshipsRS.Fields("UpdatedOn").Value)
    If IsValidDate(strTemp) Then
        oZDRChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZDRNode.appendChild oZDRChildLevel1
    
    Set ReturnNewZDRNode = oZDRNode

Exit Function

ErrTrap:
   If gobjShared Is Nothing Then
      Set gobjShared = New CWShared
   End If
   gobjShared.ShowVBError Err.Number, Err.Description, "ReturnNewZDRNode Function", "DMBE", vbExclamation
   'Err.Clear
   Resume Next 'There is no reason this procedure should not continue to create xml if it errors

End Function

Public Function ReturnNewZD2Node(ByVal lngCount As Long, ByVal DiagnosesRS As ADODB.Recordset) As MSXML.IXMLDOMElement
    ' Additional Diagnosis Data Segment
' sample function output
'   <ZD2>
'      <ZD2.1>2</ZD2.1>
'      <ZD2.2>6016</ZD2.2>
'      <ZD2.3>1046961</ZD2.3>
'      <ZD2.4>128540</ZD2.4>
'      <ZD2.5>1606</ZD2.5>
'      <ZD2.6>47681</ZD2.6>
'      <ZD2.7>1046961</ZD2.7>
'      <ZD2.8>1024</ZD2.8>
'      <ZD2.9>20010808180345</ZD2.9>
'      <ZD2.10>11312</ZD2.10>
'      <ZD2.11>Patient</ZD2.11>
'      <ZD2.12>New</ZD2.12>
'      <ZD2.13></ZD2.13>
'      <ZD2.14>False</ZD2.14>
'      <ZD2.15></ZD2.15>
'      <ZD2.16> </ZD2.16>
'      <ZD2.17>0</ZD2.17>
'      <ZD2.18></ZD2.18>
'      <ZD2.19>0</ZD2.19>
'      <ZD2.20></ZD2.20>
'      <ZD2.21></ZD2.21>
'      <ZD2.22>6016</ZD2.22>
'      <ZD2.23></ZD2.23>
'      <ZD2.24></ZD2.24>
'      <ZD2.25></ZD2.25>
'      <ZD2.26></ZD2.26>
'   </ZD2>
    
    Dim oDom As MSXML.DOMDocument
    Dim oZD2Node As MSXML.IXMLDOMElement
    Dim oZD2ChildLevel1 As MSXML.IXMLDOMElement
    Dim oZD2ChildLevel2 As MSXML.IXMLDOMElement
    Dim strTemp As String
    
    On Error GoTo ErrTrap
    
    Set oDom = New MSXML.DOMDocument
    Set oZD2Node = oDom.createElement("ZD2")
    
    ' Set ID - Count
    Set oZD2ChildLevel1 = oDom.createElement("ZD2.1")
    oZD2ChildLevel1.Text = CStr(lngCount)
    oZD2Node.appendChild oZD2ChildLevel1
    
    ' Data ID - Table Diagnoses:DataID - table Diagnoses:DataID
    Set oZD2ChildLevel1 = oDom.createElement("ZD2.2")
    oZD2ChildLevel1.Text = NullToNothing(DiagnosesRS.Fields("DataID").Value)
    oZD2Node.appendChild oZD2ChildLevel1
    
    ' Facility NCID - Table Diagnoses:Facility_NCID
    Set oZD2ChildLevel1 = oDom.createElement("ZD2.3")
    oZD2ChildLevel1.Text = NullToNothing(DiagnosesRS.Fields("FacilityNCID").Value)
    oZD2Node.appendChild oZD2ChildLevel1
    
    ' Unit Number - Table Diagnoses:Unit_Number
    Set oZD2ChildLevel1 = oDom.createElement("ZD2.4")
    oZD2ChildLevel1.Text = NullToNothing(DiagnosesRS.Fields("Unit_Number").Value)
    oZD2Node.appendChild oZD2ChildLevel1
    
    ' Encounter Number - Table Diagnoses:EncounterNumber
    Set oZD2ChildLevel1 = oDom.createElement("ZD2.5")
    oZD2ChildLevel1.Text = NullToNothing(DiagnosesRS.Fields("EncounterNumber").Value)
    oZD2Node.appendChild oZD2ChildLevel1
    
    ' Medcin ID - Table Diagnoses:SnoID
    Set oZD2ChildLevel1 = oDom.createElement("ZD2.6")
    oZD2ChildLevel1.Text = NullToNothing(DiagnosesRS.Fields("SnoID").Value)
    oZD2Node.appendChild oZD2ChildLevel1
    
    ' Point of Care Facility - Table Diagnoses:PointOfCareFacilityNCID
    Set oZD2ChildLevel1 = oDom.createElement("ZD2.7")
    strTemp = NullToNothing(DiagnosesRS.Fields("PointOfCareFacilityNCID").Value)
        
        Set oZD2ChildLevel2 = oDom.createElement("CE.1")
        oZD2ChildLevel1.appendChild oZD2ChildLevel2
        
        Set oZD2ChildLevel2 = oDom.createElement("CE.2")
        If strTemp <> "" Then
            oZD2ChildLevel2.Text = GetNCIDDescription(CLng(strTemp))
        End If
        oZD2ChildLevel1.appendChild oZD2ChildLevel2
        
        Set oZD2ChildLevel2 = oDom.createElement("CE.3")
        oZD2ChildLevel1.appendChild oZD2ChildLevel2
        Set oZD2ChildLevel2 = oDom.createElement("CE.4")
        oZD2ChildLevel1.appendChild oZD2ChildLevel2
        Set oZD2ChildLevel2 = oDom.createElement("CE.5")
        oZD2ChildLevel2.Text = strTemp 'The NCID
        oZD2ChildLevel1.appendChild oZD2ChildLevel2
        Set oZD2ChildLevel2 = oDom.createElement("CE.6")
        oZD2ChildLevel2.Text = "99NCI"
        oZD2ChildLevel1.appendChild oZD2ChildLevel2
    oZD2Node.appendChild oZD2ChildLevel1
    
    ' Status - Table Diagnoses:StatusNCID
    Set oZD2ChildLevel1 = oDom.createElement("ZD2.8")
    strTemp = NullToNothing(DiagnosesRS.Fields("StatusNCID").Value)
        
        Set oZD2ChildLevel2 = oDom.createElement("CE.1")
        oZD2ChildLevel1.appendChild oZD2ChildLevel2
        
        Set oZD2ChildLevel2 = oDom.createElement("CE.2")
        If strTemp <> "" Then
            oZD2ChildLevel2.Text = GetNCIDDescription(CLng(strTemp))
        End If
        oZD2ChildLevel1.appendChild oZD2ChildLevel2
        
        Set oZD2ChildLevel2 = oDom.createElement("CE.3")
        oZD2ChildLevel1.appendChild oZD2ChildLevel2
        Set oZD2ChildLevel2 = oDom.createElement("CE.4")
        oZD2ChildLevel1.appendChild oZD2ChildLevel2
        Set oZD2ChildLevel2 = oDom.createElement("CE.5")
        oZD2ChildLevel2.Text = strTemp 'The NCID
        oZD2ChildLevel1.appendChild oZD2ChildLevel2
        Set oZD2ChildLevel2 = oDom.createElement("CE.6")
        oZD2ChildLevel2.Text = "99NCI"
        oZD2ChildLevel1.appendChild oZD2ChildLevel2
    oZD2Node.appendChild oZD2ChildLevel1
    
    ' Noted Date - Table Diagnoses:NotedDate
    Set oZD2ChildLevel1 = oDom.createElement("ZD2.9")
    strTemp = NullToNothing(DiagnosesRS.Fields("NotedDate").Value)
    If IsValidDate(strTemp) Then
        oZD2ChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZD2Node.appendChild oZD2ChildLevel1
    
    ' Chronicity - Table Diagnoses:ChronicityNCID
    Set oZD2ChildLevel1 = oDom.createElement("ZD2.10")
    strTemp = NullToNothing(DiagnosesRS.Fields("ChronicityNCID").Value)
        Set oZD2ChildLevel2 = oDom.createElement("CE.1")
        oZD2ChildLevel1.appendChild oZD2ChildLevel2
        
        Set oZD2ChildLevel2 = oDom.createElement("CE.2")
        If strTemp <> "" Then
            '<SCR 23299 CC>
            'If mobjConvertSno Is Nothing Then
            '    Set mobjConvertSno = New actxSearchSnoMed.IConvert
            'End If
            'oZD2ChildLevel2.Text = mobjConvertSno.SnoDescription(CLng(strTemp))
            strTemp = GetNCIDDescription(strTemp)
            oZD2ChildLevel2.Text = strTemp
            '</SCR 23299 CC>
        End If
        oZD2ChildLevel1.appendChild oZD2ChildLevel2
        
        Set oZD2ChildLevel2 = oDom.createElement("CE.3")
        oZD2ChildLevel1.appendChild oZD2ChildLevel2
        Set oZD2ChildLevel2 = oDom.createElement("CE.4")
        oZD2ChildLevel1.appendChild oZD2ChildLevel2
        Set oZD2ChildLevel2 = oDom.createElement("CE.5")
        oZD2ChildLevel2.Text = strTemp
        oZD2ChildLevel1.appendChild oZD2ChildLevel2
        Set oZD2ChildLevel2 = oDom.createElement("CE.6")
        oZD2ChildLevel2.Text = "99NCI"
        oZD2ChildLevel1.appendChild oZD2ChildLevel2
    oZD2Node.appendChild oZD2ChildLevel1
    
    ' Source - Table Diagnoses:SourceNCID
    Set oZD2ChildLevel1 = oDom.createElement("ZD2.11")
    strTemp = NullToNothing(DiagnosesRS.Fields("SourceNCID").Value)
        Set oZD2ChildLevel2 = oDom.createElement("CE.1")
        oZD2ChildLevel1.appendChild oZD2ChildLevel2
        Set oZD2ChildLevel2 = oDom.createElement("CE.2")
        If strTemp <> "" Then
            oZD2ChildLevel2.Text = GetNCIDDescription(CLng(strTemp))
        End If
        oZD2ChildLevel1.appendChild oZD2ChildLevel2
        Set oZD2ChildLevel2 = oDom.createElement("CE.3")
        oZD2ChildLevel1.appendChild oZD2ChildLevel2
        Set oZD2ChildLevel2 = oDom.createElement("CE.4")
        oZD2ChildLevel1.appendChild oZD2ChildLevel2
        Set oZD2ChildLevel2 = oDom.createElement("CE.5")
        oZD2ChildLevel2.Text = strTemp
        oZD2ChildLevel1.appendChild oZD2ChildLevel2
        Set oZD2ChildLevel2 = oDom.createElement("CE.6")
        oZD2ChildLevel2.Text = "99NCI"
        oZD2ChildLevel1.appendChild oZD2ChildLevel2
    oZD2Node.appendChild oZD2ChildLevel1
    
    ' New Follow Up - Table Diagnoses:NewFollowUp
    Set oZD2ChildLevel1 = oDom.createElement("ZD2.12")
    oZD2ChildLevel1.Text = NullToNothing(DiagnosesRS.Fields("NewFollowUp").Value)
    oZD2Node.appendChild oZD2ChildLevel1
    
    ' Billing Code - Table Diagnoses:BillingCodeNCID
    Set oZD2ChildLevel1 = oDom.createElement("ZD2.13")
    strTemp = NullToNothing(DiagnosesRS.Fields("BillingCodeNCID").Value)
        Set oZD2ChildLevel2 = oDom.createElement("CE.1")
        oZD2ChildLevel1.appendChild oZD2ChildLevel2
        Set oZD2ChildLevel2 = oDom.createElement("CE.2")
        If strTemp <> "" Then
            oZD2ChildLevel2.Text = GetNCIDDescription(CLng(strTemp))
        End If
        oZD2ChildLevel1.appendChild oZD2ChildLevel2
        Set oZD2ChildLevel2 = oDom.createElement("CE.3")
        oZD2ChildLevel1.appendChild oZD2ChildLevel2
        Set oZD2ChildLevel2 = oDom.createElement("CE.4")
        oZD2ChildLevel1.appendChild oZD2ChildLevel2
        Set oZD2ChildLevel2 = oDom.createElement("CE.5")
        oZD2ChildLevel2.Text = strTemp
        oZD2ChildLevel1.appendChild oZD2ChildLevel2
        Set oZD2ChildLevel2 = oDom.createElement("CE.6")
        oZD2ChildLevel2.Text = "99NCI"
        oZD2ChildLevel1.appendChild oZD2ChildLevel2
    oZD2Node.appendChild oZD2ChildLevel1
    
    ' Sensitivity - Table Diagnoses:Sensitivity
    Set oZD2ChildLevel1 = oDom.createElement("ZD2.14")
    oZD2ChildLevel1.Text = NullToNothing(DiagnosesRS.Fields("Sensitivity").Value)
    oZD2Node.appendChild oZD2ChildLevel1
    
    ' Extended ICD9 - Table Diagnoses:ExtendedICD9
    Set oZD2ChildLevel1 = oDom.createElement("ZD2.15")
    oZD2ChildLevel1.Text = NullToNothing(DiagnosesRS.Fields("ExtendedICD9").Value)
    oZD2Node.appendChild oZD2ChildLevel1
    
    ' Diagnosis Comment - Table Diagnoses:DiagnosisComment
    Set oZD2ChildLevel1 = oDom.createElement("ZD2.16")
    oZD2ChildLevel1.Text = NullToNothing(DiagnosesRS.Fields("DiagnosesComment").Value)
    oZD2Node.appendChild oZD2ChildLevel1
    
    ' Create User - Table Diagnoses:CreateUserNCID
    Set oZD2ChildLevel1 = oDom.createElement("ZD2.17")
    oZD2ChildLevel1.Text = NullToNothing(DiagnosesRS.Fields("CreateUserNCID").Value)
    oZD2Node.appendChild oZD2ChildLevel1
    
    ' Create Date/Time - Table Diagnoses:CreateTime
    Set oZD2ChildLevel1 = oDom.createElement("ZD2.18")
    strTemp = NullToNothing(DiagnosesRS.Fields("CreateTime").Value)
    If IsValidDate(strTemp) Then
        oZD2ChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZD2Node.appendChild oZD2ChildLevel1
    
    ' Modify User - Table Diagnoses:ModifyUserNCID
    Set oZD2ChildLevel1 = oDom.createElement("ZD2.19")
    oZD2ChildLevel1.Text = NullToNothing(DiagnosesRS.Fields("ModifyUserNCID").Value)
    oZD2Node.appendChild oZD2ChildLevel1
    
    ' Modify Date/Time - Table Diagnoses:ModifyTime
    Set oZD2ChildLevel1 = oDom.createElement("ZD2.20")
    strTemp = NullToNothing(DiagnosesRS.Fields("ModifyTime").Value)
    If IsValidDate(strTemp) Then
        oZD2ChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZD2Node.appendChild oZD2ChildLevel1
    
    ' Update Flag - Table Diagnoses:UpdateFlag
    Set oZD2ChildLevel1 = oDom.createElement("ZD2.21")
    oZD2ChildLevel1.Text = NullToNothing(DiagnosesRS.Fields("UpdateFlag").Value)
    oZD2Node.appendChild oZD2ChildLevel1
    
    ' CDR Data ID - Table Diagnoses:CDR_DataID
    Set oZD2ChildLevel1 = oDom.createElement("ZD2.22")
    oZD2ChildLevel1.Text = NullToNothing(DiagnosesRS.Fields("CDR_DataID").Value)
    oZD2Node.appendChild oZD2ChildLevel1
    
    ' Created By - Table Diagnoses:CreatedBy
    Set oZD2ChildLevel1 = oDom.createElement("ZD2.23")
    oZD2ChildLevel1.Text = NullToNothing(DiagnosesRS.Fields("CreatedBy").Value)
    oZD2Node.appendChild oZD2ChildLevel1
    
    ' Created On - Table Diagnoses:CreatedOn
    Set oZD2ChildLevel1 = oDom.createElement("ZD2.24")
    strTemp = NullToNothing(DiagnosesRS.Fields("CreatedOn").Value)
    If IsValidDate(strTemp) Then
        oZD2ChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZD2Node.appendChild oZD2ChildLevel1
    
    ' Updated By - Table Diagnoses:UpdatedBy
    Set oZD2ChildLevel1 = oDom.createElement("ZD2.25")
    oZD2ChildLevel1.Text = NullToNothing(DiagnosesRS.Fields("UpdatedBy").Value)
    oZD2Node.appendChild oZD2ChildLevel1
    
    ' Updated On - Table Diagnoses:UpdatedOn
    Set oZD2ChildLevel1 = oDom.createElement("ZD2.26")
    strTemp = NullToNothing(DiagnosesRS.Fields("UpdatedOn").Value)
    If IsValidDate(strTemp) Then
        oZD2ChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZD2Node.appendChild oZD2ChildLevel1
    
    Set oZD2ChildLevel1 = oDom.createElement("ZD2.27")
    oZD2ChildLevel1.Text = NullToNothing(DiagnosesRS.Fields("ICD9NCID").Value)
    oZD2Node.appendChild oZD2ChildLevel1
    
    Set oZD2ChildLevel1 = oDom.createElement("ZD2.28")
    oZD2ChildLevel1.Text = NullToNothing(DiagnosesRS.Fields("DiagnosisNCID").Value)
    oZD2Node.appendChild oZD2ChildLevel1
    
    Set oZD2ChildLevel1 = oDom.createElement("ZD2.29")
    strTemp = NullToNothing(DiagnosesRS.Fields("OnSetDate").Value)
    If IsValidDate(strTemp) Then
        oZD2ChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
        oZD2Node.appendChild oZD2ChildLevel1
    End If
    
    Set oZD2ChildLevel1 = oDom.createElement("ZD2.30")
    oZD2ChildLevel1.Text = NullToNothing(DiagnosesRS.Fields("Priority").Value)
    oZD2Node.appendChild oZD2ChildLevel1
    
    Set oZD2ChildLevel1 = oDom.createElement("ZD2.31")
    oZD2ChildLevel1.Text = NullToNothing(DiagnosesRS.Fields("ClinicianNCID").Value)
    oZD2Node.appendChild oZD2ChildLevel1
    
    
    
    
    
    
    Set ReturnNewZD2Node = oZD2Node

Exit Function

ErrTrap:
   If gobjShared Is Nothing Then
      Set gobjShared = New CWShared
   End If
   gobjShared.ShowVBError Err.Number, Err.Description, "ReturnNewZD2Node Function", "DMBE", vbExclamation
   'Err.Clear
   Resume Next 'There is no reason this procedure should not continue to create xml if it errors

End Function

'NGIT SPR 477 - SCR 24239 CC
Public Function ReturnNewZG1Node(ByVal rsENC_DISPOSITION As ADODB.Recordset, Optional ByVal rsREADINESS As ADODB.Recordset = Nothing) As MSXML.IXMLDOMElement
    ' Procedure Segment - Enc_DispositionRS
    Dim oDom As MSXML.DOMDocument
    Dim oZG1Node As MSXML.IXMLDOMElement
    Dim oZG1ChildLevel1 As MSXML.IXMLDOMElement
    Dim strValue As String
    Dim lngValue As Long
    Dim dteStart As Date
    Dim dteEnd As Date
    
    On Error GoTo ErrTrap
    
    Set oDom = New MSXML.DOMDocument
    Set oZG1Node = oDom.createElement("ZG1")
    
    ' Type Case Code
    Set oZG1ChildLevel1 = oDom.createElement("ZG1.1")
    '<< SCR #25942;   Developer: Brian Mowbray 07/22/2002 07:39 PM
    strValue = NullToNothing(rsENC_DISPOSITION.Fields("INJ_ILL_CAUSE").Value)
    If Len(strValue) > 0 Then
      strValue = Format$(strValue, "00")
    End If
    oZG1ChildLevel1.Text = strValue
    oZG1Node.appendChild oZG1ChildLevel1
    
    ' Disease Non-Battle Injury (DNBI) Category
    Set oZG1ChildLevel1 = oDom.createElement("ZG1.2")
    
    '<< Begin:SCR #25943;   Developer: Brian Mowbray 07/22/2002 07:56 PM
    strValue = NullToNothing(rsENC_DISPOSITION.Fields("INJ_ILL_CATEGORY").Value)
    If Len(strValue) = 1 Then
      strValue = Format$(strValue, "00")
    End If
    oZG1ChildLevel1.Text = strValue
    '>> End: SCR #25943;
    oZG1Node.appendChild oZG1ChildLevel1
    
    'NGIT SPR 477 - SCR 24239 CC
    
    'Light Duty Days
    Set oZG1ChildLevel1 = oDom.createElement("ZG1.3")
    If (NullToNothing(rsENC_DISPOSITION.Fields("DISPOSITIONNCID").Value) = NCID_ReleaseWDutyLimitations) _
    Or (NullToNothing(rsENC_DISPOSITION.Fields("DISPOSITIONNCID").Value) = NCID_ReleaseWDutyLimitations_2) Then
        If Not rsREADINESS Is Nothing Then
            If rsREADINESS.RecordCount > 0 Then
                If Not rsREADINESS.EOF Then
                    If IsNull(rsREADINESS.Fields("TEMP_PROFILE_START_DT").Value) = False Then
                      dteStart = rsREADINESS.Fields("TEMP_PROFILE_START_DT").Value
                    End If
                    If IsNull(rsREADINESS.Fields("TEMP_PROFILE_END_DT").Value) = False Then
                      dteEnd = rsREADINESS.Fields("TEMP_PROFILE_END_DT").Value
                    End If
                    If IsDate(dteStart) And IsDate(dteEnd) Then
                        oZG1ChildLevel1.Text = DateDiff("d", dteStart, dteEnd) ' SCR 24250 CC
                    End If
                End If
            End If
        End If
    End If
    oZG1Node.appendChild oZG1ChildLevel1
    
    'Lost Duty Days
    Set oZG1ChildLevel1 = oDom.createElement("ZG1.4")
    If (NullToNothing(rsENC_DISPOSITION.Fields("DISPOSITIONNCID").Value) = NCID_SickAtHome) _
        Or (NullToNothing(rsENC_DISPOSITION.Fields("DISPOSITIONNCID").Value) = NCID_SickAtHome_2) Then
        strValue = NullToNothing(rsENC_DISPOSITION.Fields("DISPOSITIONTEXT").Value)
        If strValue <> "" Then
            If IsNumeric(strValue) Then
                lngValue = CInt(strValue)
                lngValue = lngValue / 24   '- convert it to the number of days
                oZG1ChildLevel1.Text = lngValue
            End If
        End If
    End If
    oZG1Node.appendChild oZG1ChildLevel1
    
    Set ReturnNewZG1Node = oZG1Node
    
Exit Function

ErrTrap:
   If gobjShared Is Nothing Then
      Set gobjShared = New CWShared
   End If
   gobjShared.ShowVBError Err.Number, Err.Description, "ReturnNewZG1Node Function", "DMBE", vbExclamation
   'Err.Clear
   Resume Next 'There is no reason this procedure should not continue on to create xml
           
End Function


Public Function ReturnNewDG1Node(ByVal lngCount As Long, ByVal DiagnosesRS As ADODB.Recordset) As MSXML.IXMLDOMElement
   ' Additional Patient Visit Data Segment
' sample function output
'   <DG1>
'      <DG1.1>1</DG1.1>
'      <DG1.3>
'         <CE.1>3859</CE.1>
'         <CE.2>exposed to chronic vibration from tool operation</CE.2>
'         <CE.3>I9</CE.3>
'         <CE.4>3859</CE.4>
'         <CE.5>exposed to chronic vibration from tool operation</CE.5>
'         <CE.6>99NCI</CE.6>
'      </DG1.3>
'      <DG1.5>20010808172404</DG1.5>
'      <DG1.15>1</DG1.15>
'      <DG1.16>
'         <XCN.1>123459876</XCN.1>
'         <XCN.2>SYSTEM</XCN.2>
'         <XCN.3>GENERATED</XCN.3>
'         <XCN.4></XCN.4>
'      </DG1.16>
'   </DG1>

    Dim oDom As MSXML.DOMDocument
    Dim oDG1Node As MSXML.IXMLDOMElement
    Dim oDG1ChildLevel1 As MSXML.IXMLDOMElement
    Dim oDG1ChildLevel2 As MSXML.IXMLDOMElement
    Dim oDG1ChildLevel3 As MSXML.IXMLDOMElement
    Dim strDocFirstName As String
    Dim strDocLastName As String
    Dim strDocMiddleName As String
    Dim strDocID As String
    Dim strDocSpecialty As String
    Dim strTemp As String
    Dim blnRsIsGood As Boolean
    
    On Error GoTo ErrTrap
    
    Set oDom = New MSXML.DOMDocument
    Set oDG1Node = oDom.createElement("DG1")
    
    ' Count
    Set oDG1ChildLevel1 = oDom.createElement("DG1.1")
    oDG1ChildLevel1.Text = CStr(lngCount)
    oDG1Node.appendChild oDG1ChildLevel1
    
    '<SCR 22597 CC>
    Set oDG1ChildLevel1 = oDom.createElement("DG1.3")
        ' Diagnosis Code Identifier (ICD-9) - Table Diagnoses:ICD9NCID
        Set oDG1ChildLevel2 = oDom.createElement("CE.1")
        strTemp = NullToNothing(DiagnosesRS.Fields("ICD9NCID").Value)
        If mobjConvertSno Is Nothing Then
            Set mobjConvertSno = New actxSearchSnoMed.IConvert
        End If
        If strTemp <> "" Then
            strTemp = mobjConvertSno.SnoICD9(CLng(strTemp))
            'strTemp = mobjConvertSno.SnoCPT4(CLng(strTemp))
        End If
        
        'Data migration Issue
        'This is being added based on errors in TMIP parsing of HL7 Messages
        ' If <DG1.3><CE.1> messages are blank then message is placed in TMIP Error log.
        'The following two codes are being seen during data migration
        If strTemp = "" Then
          If NullToNothing(DiagnosesRS.Fields("ICD9NCID").Value) = 8337 Then
            strTemp = "307.9"
          ElseIf NullToNothing(DiagnosesRS.Fields("ICD9NCID").Value) = 11057 Then
            strTemp = "919.0"
          Else
            strTemp = "0"
          End If
          
        End If
        
        oDG1ChildLevel2.Text = strTemp
        'oDG1ChildLevel2.Text = NullToNothing(DiagnosesRS.Fields("ICD9NCID").Value)
        oDG1ChildLevel1.appendChild oDG1ChildLevel2
        
        'ICD-9 CODE Description - based on DiagnosisNCID
        Set oDG1ChildLevel2 = oDom.createElement("CE.2")
        strTemp = NullToNothing(DiagnosesRS.Fields("DiagnosisNCID").Value)
        'strTemp = NullToNothing(DiagnosesRS.Fields("ICD9NCID").Value)
        '</SCR 22597 CC>
        If strTemp <> "" Then
            oDG1ChildLevel2.Text = mobjConvertSno.SnoDescription(strTemp) '<SCR 22597 CC>
        End If
        oDG1ChildLevel1.appendChild oDG1ChildLevel2
        
        'Name of Coding System
        Set oDG1ChildLevel2 = oDom.createElement("CE.3")
        oDG1ChildLevel2.Text = "I9"
        oDG1ChildLevel1.appendChild oDG1ChildLevel2
        
        'Alternate Identifier  Table Diagnoses:DiagnosisNCID
        Set oDG1ChildLevel2 = oDom.createElement("CE.4")
        'oDG1ChildLevel2.Text = NullToNothing(DiagnosesRS.Fields("ICD9NCID").Value)
        strTemp = NullToNothing(DiagnosesRS.Fields("DiagnosisNCID").Value)
        oDG1ChildLevel1.appendChild oDG1ChildLevel2
        
        ' Alternate Text - Table Diagnoses:ICD9NCID
        Set oDG1ChildLevel2 = oDom.createElement("CE.5")
        strTemp = NullToNothing(DiagnosesRS.Fields("DiagnosisNCID").Value)
        'strTemp = NullToNothing(DiagnosesRS.Fields("ICD9NCID").Value)
        '</SCR 22597 CC>
        If strTemp <> "" Then
            oDG1ChildLevel2.Text = mobjConvertSno.SnoDescription(strTemp) '<SCR 22597 CC>
        End If
        oDG1ChildLevel1.appendChild oDG1ChildLevel2
        
        Set oDG1ChildLevel2 = oDom.createElement("CE.6")
        oDG1ChildLevel2.Text = "99NCI"
        oDG1ChildLevel1.appendChild oDG1ChildLevel2
    oDG1Node.appendChild oDG1ChildLevel1
    '</SCR 22597 CC>
    
    ' Diagnosis Date/Time - Table Diagnoses:OnSetDate
    Set oDG1ChildLevel1 = oDom.createElement("DG1.5")
    strTemp = NullToNothing(DiagnosesRS.Fields("OnSetDate").Value)
    If IsValidDate(strTemp) Then
        oDG1ChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oDG1Node.appendChild oDG1ChildLevel1
    
    'Diagnosis Type
    Set oDG1ChildLevel1 = oDom.createElement("DG1.6")
    oDG1ChildLevel1.Text = "F"
    oDG1Node.appendChild oDG1ChildLevel1
    
    ' Diagnosis Priority - Table Diagnoses:Priority (values 1-6)
    Set oDG1ChildLevel1 = oDom.createElement("DG1.15")
    oDG1ChildLevel1.Text = NullToNothing(DiagnosesRS.Fields("Priority").Value)
    oDG1Node.appendChild oDG1ChildLevel1
    
    '<< Begin:SCR #26641;   Developer: Brian Mowbray 08/15/2002 06:37 PM
    ' Diagnosing Clinician - Table Diagnoses:ClinicianNCID
    
    Set oDG1ChildLevel1 = oDom.createElement("DG1.16.LST")
        Set oDG1ChildLevel2 = oDom.createElement("DG1.16")
          strTemp = NullToNothing(DiagnosesRS.Fields("ClinicianNCID").Value)
          If strTemp <> "" Then
              Call GetNewDoctorInfo(CLng(strTemp), strDocLastName, strDocFirstName, strDocMiddleName, strDocID, strDocSpecialty)
              blnRsIsGood = True
          End If
          
          ' Diagnosing Clinician ID Number (SSN)
          Set oDG1ChildLevel3 = oDom.createElement("XCN.1")
          If blnRsIsGood = True Then
              oDG1ChildLevel3.Text = strDocID
          End If
          oDG1ChildLevel2.appendChild oDG1ChildLevel3
          ' Diagnosing Clinician Family Last Name
          Set oDG1ChildLevel3 = oDom.createElement("XCN.2")
          If blnRsIsGood = True Then
              oDG1ChildLevel3.Text = strDocLastName
          End If
          oDG1ChildLevel2.appendChild oDG1ChildLevel3
          ' Diagnosing Clinician Family First Name
          Set oDG1ChildLevel3 = oDom.createElement("XCN.3")
          If blnRsIsGood = True Then
              oDG1ChildLevel3.Text = strDocFirstName
          End If
          oDG1ChildLevel2.appendChild oDG1ChildLevel3
          ' Diagnosing Clinician Family Middle Name
          Set oDG1ChildLevel3 = oDom.createElement("XCN.4")
          If blnRsIsGood = True Then
              oDG1ChildLevel3.Text = strDocMiddleName
          End If
        oDG1ChildLevel2.appendChild oDG1ChildLevel3
    oDG1ChildLevel1.appendChild oDG1ChildLevel2
  
  oDG1Node.appendChild oDG1ChildLevel1
    '>> End: SCR #26641;
    
  Set ReturnNewDG1Node = oDG1Node
    
Exit Function

ErrTrap:
   If gobjShared Is Nothing Then
      Set gobjShared = New CWShared
   End If
   gobjShared.ShowVBError Err.Number, Err.Description, "ReturnNewDG1Node Function", "DMBE", vbExclamation
   'Err.Clear
   Resume Next 'There is no reason this procedure should not continue to create xml if it errors

End Function


Public Function ReturnNewZRANode(ByVal ImmRS As ADODB.Recordset) As MSXML.IXMLDOMElement

   ' Additional Pharmacy Admistration Data Segment
   
' sample function output
'   <ZRA>
'      <ZRA.1>Sponsor SSN</ZRA.1>
'      <ZRA.2>Vaccination Series</ZRA.2>
'      <ZRA.3>Vaccination Next Due</ZRA.3>
'      <ZRA.4>
'         <CE.1>Identifier</CE.1>
'         <CE.2>Text</CE.2>
'         <CE.4>Alternate Identifier</CE.4>
'         <CE.5>Alternate Text</CE.5>
'      </ZRA.4>
'      <ZRA.5>Last DEERS Ack</ZRA.5>
'      <ZRA.6>DEERS Action</ZRA.6>
'      <ZRA.7>DEERS Instance</ZRA.7>
'      <ZRA.8>Route</ZRA.8>
'      <ZRA.9>Site</ZRA.9>
'      <ZRA.10>Exemption Expire Date</ZRA.10>
'      <ZRA.11>Result</ZRA.11>
'      <ZRA.12>RXN Size</ZRA.12>
'      <ZRA.13>Central Date</ZRA.13>
'      <ZRA.14>INCDR</ZRA.14>
'      <ZRA.15>Manufacturers Note</ZRA.15>
'      <ZRA.16>CPT4 Code</ZRA.16>
'      <ZRA.17>ICD9 Code</ZRA.17>
'      <ZRA.18>Number in Series</ZRA.18>
'      <ZRA.19>Schedule</ZRA.19>
'      <ZRA.20>Booster</ZRA.20>
'      <ZRA.21>Report Header</ZRA.21>
'      <ZRA.22>MPF DIN</ZRA.22>
'      <ZRA.23>Generic Name</ZRA.23>
'      <ZRA.24>Display</ZRA.24>
'      <ZRA.25>Barcode</ZRA.25>
'      <ZRA.26>Childhood</ZRA.26>
'      <ZRA.27>
'         <CE.1>Identifier</CE.1>
'         <CE.2>Text</CE.2>
'      </ZRA.27>
'      <ZRA.28>
'         <CE.1>Identifier</CE.1>
'         <CE.2>Text</CE.2>
'      </ZRA.28>
'      <ZRA.29>Last Edited By</ZRA.29>
'      <ZRA.30>Last Edited Date/Time</ZRA.30>
'   </ZRA>
   
    Dim oDom As MSXML.DOMDocument
    Dim oZRANode As MSXML.IXMLDOMElement
    Dim oZRAChildLevel1 As MSXML.IXMLDOMElement
    Dim oZRAChildLevel2 As MSXML.IXMLDOMElement
    Dim strTemp As String
    Dim blnTemp As Boolean
    
    On Error GoTo ErrTrap

    Set oDom = New MSXML.DOMDocument
    Set oZRANode = oDom.createElement("ZRA")
    
    ' Sponsor SSN - Table Imm_People_Vaccine:Spon_SSAN_FMP
    Set oZRAChildLevel1 = oDom.createElement("ZRA.1")
    oZRAChildLevel1.Text = NullToNothing(ImmRS.Fields("SPON_SSAN_FMP").Value)
    oZRANode.appendChild oZRAChildLevel1

    ' Vaccination Series - Table Imm_People_Vaccine:Vacc_Series
    Set oZRAChildLevel1 = oDom.createElement("ZRA.2")
    oZRAChildLevel1.Text = NullToNothing(ImmRS.Fields("VACC_SERIES").Value)
    oZRANode.appendChild oZRAChildLevel1

    ' Vaccination Next Due - Table Imm_People_Vaccine:Vacc_NextDue
    Set oZRAChildLevel1 = oDom.createElement("ZRA.3")
    strTemp = NullToNothing(ImmRS.Fields("VACC_NEXTDUE").Value)
    If IsValidDate(strTemp) Then
        oZRAChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZRANode.appendChild oZRAChildLevel1

    ' Vaccination Exempt
    Set oZRAChildLevel1 = oDom.createElement("ZRA.4")
        ' Identifier - Table Imm_People_Vaccine:Vacc_Exempt
        Set oZRAChildLevel2 = oDom.createElement("CE.1")
        oZRAChildLevel2.Text = NullToNothing(ImmRS.Fields("VACC_EXEMPT").Value)
        oZRAChildLevel1.appendChild oZRAChildLevel2
        ' Text - Table TblVacc_Exempt:Vacc_Exempt_Text
        Set oZRAChildLevel2 = oDom.createElement("CE.2")
        oZRAChildLevel2.Text = NullToNothing(ImmRS.Fields("VACC_EXEMPT_TEXT").Value)
        oZRAChildLevel1.appendChild oZRAChildLevel2
        ' Alternate Identifier - Table TblVacc_Exempt:Vacc_Exempt_Code
        Set oZRAChildLevel2 = oDom.createElement("CE.4")
        oZRAChildLevel2.Text = NullToNothing(ImmRS.Fields("VACC_EXEMPT_CODE").Value)
        oZRAChildLevel1.appendChild oZRAChildLevel2
        ' Alternate Text - Table TblVacc_Exempt:Vacc_Exempt_Desc
        Set oZRAChildLevel2 = oDom.createElement("CE.5")
        oZRAChildLevel2.Text = NullToNothing(ImmRS.Fields("VACC_EXEMPT_DESC").Value)
        oZRAChildLevel1.appendChild oZRAChildLevel2
        
    oZRANode.appendChild oZRAChildLevel1
        
    ' Last DEERS Ack - Table Imm_People_Vaccine:Last_DEERS_Ack
    Set oZRAChildLevel1 = oDom.createElement("ZRA.5")
    strTemp = NullToNothing(ImmRS.Fields("LASTDEERS_ACK").Value)
    If IsValidDate(strTemp) Then
        oZRAChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZRANode.appendChild oZRAChildLevel1
    
    ' DEERS Action - Table Imm_People_Vaccine:DEERS_Action
    Set oZRAChildLevel1 = oDom.createElement("ZRA.6")
    oZRAChildLevel1.Text = NullToNothing(ImmRS.Fields("DEERS_ACTION").Value)
    oZRANode.appendChild oZRAChildLevel1
    
    ' DEERS Instance - Table Imm_People_Vaccine:DEERS_Instance
    Set oZRAChildLevel1 = oDom.createElement("ZRA.7")
    oZRAChildLevel1.Text = NullToNothing(ImmRS.Fields("DEERS_INSTANCE").Value)
    oZRANode.appendChild oZRAChildLevel1
    
    ' Route - Table Imm_People_Vaccine:Route
    Set oZRAChildLevel1 = oDom.createElement("ZRA.8")
    oZRAChildLevel1.Text = NullToNothing(ImmRS.Fields("ROUTE").Value)
    oZRANode.appendChild oZRAChildLevel1
    
    ' Site - Table Imm_People_Vaccine:Site
    Set oZRAChildLevel1 = oDom.createElement("ZRA.9")
    oZRAChildLevel1.Text = NullToNothing(ImmRS.Fields("SITE").Value)
    oZRANode.appendChild oZRAChildLevel1
    
    ' Exemption Expire Date - Table Imm_People_Vaccine:ExemptionExpirationDate
    Set oZRAChildLevel1 = oDom.createElement("ZRA.10")
    strTemp = NullToNothing(ImmRS.Fields("EXEMPTIONEXPIRATIONDATE").Value)
    If IsValidDate(strTemp) Then
        oZRAChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZRANode.appendChild oZRAChildLevel1
    
    ' Result - Table Imm_People_Vaccine:Result
    Set oZRAChildLevel1 = oDom.createElement("ZRA.11")
    oZRAChildLevel1.Text = NullToNothing(ImmRS.Fields("RESULT").Value)
    oZRANode.appendChild oZRAChildLevel1
    
    ' RXN Size - Table Imm_People_Vaccine:Rxn_Size
    Set oZRAChildLevel1 = oDom.createElement("ZRA.12")
    oZRAChildLevel1.Text = NullToNothing(ImmRS.Fields("RXN_SIZE").Value)
    oZRANode.appendChild oZRAChildLevel1
    
    ' Central Date - Table Imm_People_Vaccine:CentralDate
    Set oZRAChildLevel1 = oDom.createElement("ZRA.13")
    strTemp = NullToNothing(ImmRS.Fields("CENTRALDATE").Value)
    If IsValidDate(strTemp) Then
        oZRAChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZRANode.appendChild oZRAChildLevel1
    
    ' INCDR - Table Imm_People_Vaccine:InCDR
    Set oZRAChildLevel1 = oDom.createElement("ZRA.14")
    If Not IsNull(ImmRS.Fields("INCDR").Value) Then
        blnTemp = ImmRS.Fields("INCDR").Value
        If blnTemp = True Then
            oZRAChildLevel1.Text = "True"
        Else
            oZRAChildLevel1.Text = "False"
        End If
    End If
    oZRANode.appendChild oZRAChildLevel1
    
    ' Manufacturers Note - table TblVaccine_MFG:MFG_Note
    Set oZRAChildLevel1 = oDom.createElement("ZRA.15")
    oZRAChildLevel1.Text = NullToNothing(ImmRS.Fields("MFG_NOTE").Value)
    oZRANode.appendChild oZRAChildLevel1
    
    ' CPT4 Code - Table TblVaccines:CPT
    Set oZRAChildLevel1 = oDom.createElement("ZRA.16")
    oZRAChildLevel1.Text = NullToNothing(ImmRS.Fields("CPT").Value)
    oZRANode.appendChild oZRAChildLevel1
    
    ' ICD9 Code - Table TblVaccines:ICD
    Set oZRAChildLevel1 = oDom.createElement("ZRA.17")
    oZRAChildLevel1.Text = NullToNothing(ImmRS.Fields("ICD").Value)
    oZRANode.appendChild oZRAChildLevel1
    
     ' Number in Series - Table TblVaccines:NbrInSeries
    Set oZRAChildLevel1 = oDom.createElement("ZRA.18")
    oZRAChildLevel1.Text = NullToNothing(ImmRS.Fields("NBRINSERIES").Value)
    oZRANode.appendChild oZRAChildLevel1
    
    ' Schedule - Table TblVaccines:Schedule
    Set oZRAChildLevel1 = oDom.createElement("ZRA.19")
    oZRAChildLevel1.Text = NullToNothing(ImmRS.Fields("SCHEDULE").Value)
    oZRANode.appendChild oZRAChildLevel1
    
    ' Booster - Table TblVaccines:Booster
    Set oZRAChildLevel1 = oDom.createElement("ZRA.20")
    oZRAChildLevel1.Text = NullToNothing(ImmRS.Fields("BOOSTER").Value)
    oZRANode.appendChild oZRAChildLevel1
    
    ' Report Header - Table TblVaccines:ReportHeader
    Set oZRAChildLevel1 = oDom.createElement("ZRA.21")
    oZRAChildLevel1.Text = NullToNothing(ImmRS.Fields("REPORTHEADER").Value)
    oZRANode.appendChild oZRAChildLevel1
    
    ' MPF DIN - Table TblVaccines:MPF_DIN
    Set oZRAChildLevel1 = oDom.createElement("ZRA.22")
    oZRAChildLevel1.Text = NullToNothing(ImmRS.Fields("MPF_DIM").Value)
    oZRANode.appendChild oZRAChildLevel1
    
    ' Generic Name - Table TblVaccines:GenericName
    Set oZRAChildLevel1 = oDom.createElement("ZRA.23")
    oZRAChildLevel1.Text = NullToNothing(ImmRS.Fields("GENERICNAME").Value)
    oZRANode.appendChild oZRAChildLevel1
    
    ' Display - Table TblVaccines:Display
    Set oZRAChildLevel1 = oDom.createElement("ZRA.24")
    If Not IsNull(ImmRS.Fields("DISPLAY").Value) Then
        blnTemp = ImmRS.Fields("DISPLAY").Value
        If blnTemp = True Then
            oZRAChildLevel1.Text = "True"
        Else
            oZRAChildLevel1.Text = "False"
        End If
    End If
    oZRANode.appendChild oZRAChildLevel1
    
    ' Barcode - Table TblVaccines:BarCode
    Set oZRAChildLevel1 = oDom.createElement("ZRA.25")
    If Not IsNull(ImmRS.Fields("BARCODE").Value) Then
        blnTemp = ImmRS.Fields("BARCODE").Value
        If blnTemp = True Then
            oZRAChildLevel1.Text = "True"
        Else
            oZRAChildLevel1.Text = "False"
        End If
    End If
    oZRANode.appendChild oZRAChildLevel1
    
    ' Childhood - Table TblVaccines:Childhood
    Set oZRAChildLevel1 = oDom.createElement("ZRA.26")
    If Not IsNull(ImmRS.Fields("CHILDHOOD").Value) Then
        blnTemp = ImmRS.Fields("CHILDHOOD").Value
        If blnTemp = True Then
            oZRAChildLevel1.Text = "True"
        Else
            oZRAChildLevel1.Text = "False"
        End If
    End If
    oZRANode.appendChild oZRAChildLevel1
    
    ' Patient Service
    Set oZRAChildLevel1 = oDom.createElement("ZRA.27")
        ' Identifier - Table TblService:PT_Service
        Set oZRAChildLevel2 = oDom.createElement("CE.1")
        oZRAChildLevel2.Text = NullToNothing(ImmRS.Fields("PT_SERVICE").Value)
        oZRAChildLevel1.appendChild oZRAChildLevel2
        ' Text - Table TblService:Service
        Set oZRAChildLevel2 = oDom.createElement("CE.2")
        oZRAChildLevel2.Text = NullToNothing(ImmRS.Fields("SERVICE").Value)
        oZRAChildLevel1.appendChild oZRAChildLevel2
    oZRANode.appendChild oZRAChildLevel1
       
    ' Patient Status
    Set oZRAChildLevel1 = oDom.createElement("ZRA.28")
        ' Identifier - Table TblPatientStatus:Status
        Set oZRAChildLevel2 = oDom.createElement("CE.1")
        oZRAChildLevel2.Text = NullToNothing(ImmRS.Fields("STATUS").Value)
        oZRAChildLevel1.appendChild oZRAChildLevel2
        ' Text - Table TblPatientStatus:StatusText
        Set oZRAChildLevel2 = oDom.createElement("CE.2")
        oZRAChildLevel2.Text = NullToNothing(ImmRS.Fields("STATUSTEXT").Value)
        oZRAChildLevel1.appendChild oZRAChildLevel2
    oZRANode.appendChild oZRAChildLevel1
    
    ' Last Edited By - Table Imm_People_Vaccine:LASTEDITEDBY
    Set oZRAChildLevel1 = oDom.createElement("ZRA.29")
    oZRAChildLevel1.Text = NullToNothing(ImmRS.Fields("LASTEDITEDBY").Value)
    oZRANode.appendChild oZRAChildLevel1
    
    ' Last Edited Date/Time - Table Imm_People_Vaccine:LASTEDITED
    Set oZRAChildLevel1 = oDom.createElement("ZRA.30")
    strTemp = NullToNothing(ImmRS.Fields("LASTEDITED").Value)
    If IsValidDate(strTemp) Then
        oZRAChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZRANode.appendChild oZRAChildLevel1
              
    Set ReturnNewZRANode = oZRANode

Exit Function

ErrTrap:
   If gobjShared Is Nothing Then
      Set gobjShared = New CWShared
   End If
   gobjShared.ShowVBError Err.Number, Err.Description, "ReturnNewZRANode Function", "DMBE", vbExclamation
   Resume Next 'There is no reason this procedure should not continue to create xml if it errors
       
End Function

Public Function ReturnNewRXANode(ByVal lngCount As Long, ByVal ImmRS As ADODB.Recordset) As MSXML.IXMLDOMElement
   ' Pharmacy Admistration Segment

' sample function output
'   <RXA>
'      <RXA.1>1</RXA.1>
'      <RXA.2>1</RXA.2>
'      <RXA.3>20010911</RXA.3>
'      <RXA.4></RXA.4>
'      <RXA.5>
'         <CE.1>90581</CE.1>
'         <CE.2>Anthrax</CE.2>
'         <CE.3>C4</CE.3>
'         <CE.4>14515233</CE.4>
'         <CE.5>Anthrax</CE.5>
'         <CE.6>99NCI</CE.6>
'      </RXA.5>
'      <RXA.6></RXA.6>
'      <RXA.10.LST>
'         <RXA.10>
'            <XCN.1>TRANSCRIBED</XCN.1>
'            <XCN.2>TRANSCRIBED</XCN.2>
'            <XCN.3>RECORD</XCN.3>
'            <XCN.4></XCN.4>
'         </RXA.10>
'      </RXA.10.LST>
'      <RXA.11>Unknown</RXA.11>
'      <RXA.15.LST>
'         <RXA.15>TRANSCRIBED</RXA.15>
'      </RXA.15.LST>
'      <RXA.17.LST>
'         <RXA.17>
'            <CE.1></CE.1>
'            <CE.2>Bioport</CE.2>
'            <CE.3></CE.3>
'            <CE.4></CE.4>
'            <CE.5>Bioport</CE.5>
'            <CE.6>99NCI</CE.6>
'         </RXA.17>
'      </RXA.17.LST>
'      <RXA.22>20010920</RXA.22>
'   </RXA>

    Dim oDom As MSXML.DOMDocument
    Dim oRXANode As MSXML.IXMLDOMElement
    Dim oRXAChildLevel1 As MSXML.IXMLDOMElement
    Dim oRXAChildLevel2 As MSXML.IXMLDOMElement
    Dim oRXAChildLevel3 As MSXML.IXMLDOMElement
    Dim strTemp As String
    Dim strDosage() As String
    Dim blnRsIsGood As Boolean
    
    On Error GoTo ErrTrap

    '**********************************************
    'SCR 68617 - Bill Fote - 07-Apr-2005
    If gobjShared Is Nothing Then
      Set gobjShared = New CWShared
    End If
    '**********************************************

    Set oDom = New MSXML.DOMDocument
    Set oRXANode = oDom.createElement("RXA")
    
    ' Count
    Set oRXAChildLevel1 = oDom.createElement("RXA.1")
    oRXAChildLevel1.Text = CStr(lngCount)
    oRXANode.appendChild oRXAChildLevel1

    ' Medical Administration Count
    Set oRXAChildLevel1 = oDom.createElement("RXA.2")
    oRXAChildLevel1.Text = CStr(lngCount)
    oRXANode.appendChild oRXAChildLevel1

    ' Date/Time Start of Administration - Table IMM_People_Vaccine:Vacc_Date
    Set oRXAChildLevel1 = oDom.createElement("RXA.3")
    strTemp = NullToNothing(ImmRS.Fields("VACC_DATE").Value)
    If IsValidDate(strTemp) Then
        oRXAChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oRXANode.appendChild oRXAChildLevel1
    
    ' Date/Time End of Administration - Table IMM_People_Vaccine:???
    Set oRXAChildLevel1 = oDom.createElement("RXA.4")
    'strTemp = NullToNothing(ImmRS.Fields("VACC_DATE").Value)
    'If isvaliddate(strTemp) Then
    '    oRXAChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    'End If
    oRXAChildLevel1.Text = ""
    oRXANode.appendChild oRXAChildLevel1
    
    ' Administered Code
    Set oRXAChildLevel1 = oDom.createElement("RXA.5")
        ' Identifier - Table IMM_People_Vaccine:CPT
        Set oRXAChildLevel2 = oDom.createElement("CE.1")
        oRXAChildLevel2.Text = NullToNothing(ImmRS.Fields("CPT").Value)
        oRXAChildLevel1.appendChild oRXAChildLevel2
        ' Text - Table TblVaccines:Vaccine
        Set oRXAChildLevel2 = oDom.createElement("CE.2")
        oRXAChildLevel2.Text = NullToNothing(ImmRS.Fields("VACCINE").Value)
        oRXAChildLevel1.appendChild oRXAChildLevel2
        ' Name of Coding System -
        Set oRXAChildLevel2 = oDom.createElement("CE.3")
        oRXAChildLevel2.Text = "C4"
        oRXAChildLevel1.appendChild oRXAChildLevel2
        ' Alternate Identifier - Table TblVaccines:Vacc_NCID
        Set oRXAChildLevel2 = oDom.createElement("CE.4")
        oRXAChildLevel2.Text = NullToNothing(ImmRS.Fields("VACC_NCID").Value)
        oRXAChildLevel1.appendChild oRXAChildLevel2
        ' Alternate Text - Table TblVaccines:GenericName
        Set oRXAChildLevel2 = oDom.createElement("CE.5")
        oRXAChildLevel2.Text = NullToNothing(ImmRS.Fields("GENERICNAME").Value)
        oRXAChildLevel1.appendChild oRXAChildLevel2
        ' Alternate Name of Coding System
        Set oRXAChildLevel2 = oDom.createElement("CE.6")
        oRXAChildLevel2.Text = "99NCI"
        oRXAChildLevel1.appendChild oRXAChildLevel2
    oRXANode.appendChild oRXAChildLevel1
      
    ' Administered Amount - Table IMM_People_Vaccine:Vacc_Dosage
    Set oRXAChildLevel1 = oDom.createElement("RXA.6")
    blnRsIsGood = False
    strTemp = NullToNothing(ImmRS.Fields("VACC_DOSAGE").Value)
    '<< Begin:SCR #25477;#27345   Developer: Brian Mowbray 09/09/2002 11:10 AM
    '>> SCR #38306;   Developer: Brian Mowbray 07/02/2003 05:24 PM
    If Trim$(strTemp) <> "" And UCase$(strTemp) <> "NONE" Then
        strDosage = Split(strTemp, " ", -1, vbTextCompare)
        blnRsIsGood = True
        oRXAChildLevel1.Text = NullToZero(strDosage(0))
    Else
      oRXAChildLevel1.Text = NullToZero(strTemp)
    End If
    '>> End: SCR #25477;#27345
    oRXANode.appendChild oRXAChildLevel1

    '<SCR 24313 CC>
    'Unit of the dosage
    Set oRXAChildLevel1 = oDom.createElement("RXA.7")
        ' Identifier
        Set oRXAChildLevel2 = oDom.createElement("CE.1")
        oRXAChildLevel1.appendChild oRXAChildLevel2
        ' Text
        Set oRXAChildLevel2 = oDom.createElement("CE.2")
        If blnRsIsGood = True Then
            oRXAChildLevel2.Text = strDosage(1)
        End If
        oRXAChildLevel1.appendChild oRXAChildLevel2
        ' Name of Coding System -
        Set oRXAChildLevel2 = oDom.createElement("CE.3")
        oRXAChildLevel1.appendChild oRXAChildLevel2
        ' Alternate Identifier -
        Set oRXAChildLevel2 = oDom.createElement("CE.4")
        oRXAChildLevel1.appendChild oRXAChildLevel2
        ' Alternate Text - Table TblVaccines:GenericName
        Set oRXAChildLevel2 = oDom.createElement("CE.5")
        If blnRsIsGood = True Then
            oRXAChildLevel2.Text = strDosage(1)
        End If
        oRXAChildLevel1.appendChild oRXAChildLevel2
        ' Alternate Name of Coding System
        Set oRXAChildLevel2 = oDom.createElement("CE.6")
        oRXAChildLevel1.appendChild oRXAChildLevel2
    oRXANode.appendChild oRXAChildLevel1

    ' Administering Provider
    Set oRXAChildLevel1 = oDom.createElement("RXA.10.LST")
        Set oRXAChildLevel2 = oDom.createElement("RXA.10")
            ' ID Number - Table IMM_People_Vaccine:ProviderID
            Set oRXAChildLevel3 = oDom.createElement("XCN.1")
            oRXAChildLevel3.Text = NullToNothing(ImmRS.Fields("PROVIDERID").Value)
            oRXAChildLevel2.appendChild oRXAChildLevel3
            ' Provider Last Name - Use ProviderID NCID to find
            Set oRXAChildLevel3 = oDom.createElement("XCN.2")
            oRXAChildLevel3.Text = NullToNothing(ImmRS.Fields("IP_LASTNAME").Value)
            oRXAChildLevel2.appendChild oRXAChildLevel3
            ' Provider First Name - Use ProviderID NCID to find
            Set oRXAChildLevel3 = oDom.createElement("XCN.3")
            oRXAChildLevel3.Text = NullToNothing(ImmRS.Fields("IP_FIRSTNAME").Value)
            oRXAChildLevel2.appendChild oRXAChildLevel3
            ' Provider Middle Name - Use ProviderID NCID to find
            Set oRXAChildLevel3 = oDom.createElement("XCN.4")
            oRXAChildLevel3.Text = NullToNothing(ImmRS.Fields("IP_MI").Value)
            oRXAChildLevel2.appendChild oRXAChildLevel3
        oRXAChildLevel1.appendChild oRXAChildLevel2
    oRXANode.appendChild oRXAChildLevel1
    
    ' Administered at Location - Table IMM_People_Vaccine:SITE
    Set oRXAChildLevel1 = oDom.createElement("RXA.11")
    strTemp = NullToNothing(gobjShared.CurrentFacilityNCID)
    '<SCR 24311 CC>
    If strTemp <> "" Then
        oRXAChildLevel1.Text = NullToNothing(GetNCIDDescription(strTemp))
        'oRXAChildLevel1.Text = NullToNothing(ImmRS.Fields("SITE").Value)
    Else
    
        oRXAChildLevel1.Text = "CHCSII-T ADT"
    End If
    '</SCR 24311 CC>
    oRXANode.appendChild oRXAChildLevel1
    
    ' Substance Lot Number - Table IMM_People_Vaccine:Lot_NBR
    Set oRXAChildLevel1 = oDom.createElement("RXA.15.LST")
        Set oRXAChildLevel2 = oDom.createElement("RXA.15")
        oRXAChildLevel2.Text = NullToNothing(ImmRS.Fields("LOT_NBR").Value)
        oRXAChildLevel1.appendChild oRXAChildLevel2
    oRXANode.appendChild oRXAChildLevel1
    
    ' Substance Manufacture Name
    Set oRXAChildLevel1 = oDom.createElement("RXA.17.LST")
        Set oRXAChildLevel2 = oDom.createElement("RXA.17")
            ' Identifier - Table IMM_People_Vaccine:MFG_Code)
            Set oRXAChildLevel3 = oDom.createElement("CE.1")
            oRXAChildLevel3.Text = NullToNothing(ImmRS.Fields("MFG_CODE").Value)
            oRXAChildLevel2.appendChild oRXAChildLevel3
            ' Text - Table TblVaccine_MFG:Vaccine_MFG
            Set oRXAChildLevel3 = oDom.createElement("CE.2")
            oRXAChildLevel3.Text = NullToNothing(ImmRS.Fields("VACCINE_MFG").Value)
            oRXAChildLevel2.appendChild oRXAChildLevel3
            ' Name of Coding System -
            Set oRXAChildLevel3 = oDom.createElement("CE.3")
            oRXAChildLevel3.Text = ""
            oRXAChildLevel2.appendChild oRXAChildLevel3
            ' Alternate Identifier - Table TblVaccine_MFG:MFG_NCID
            Set oRXAChildLevel3 = oDom.createElement("CE.4")
            oRXAChildLevel3.Text = NullToNothing(ImmRS.Fields("MFG_NCID").Value)
            oRXAChildLevel2.appendChild oRXAChildLevel3
            ' Alternate Text - Table TblVaccine_MFG:Vaccine_MFG
            Set oRXAChildLevel3 = oDom.createElement("CE.5")
            oRXAChildLevel3.Text = NullToNothing(ImmRS.Fields("Vaccine_MFG").Value)
            oRXAChildLevel2.appendChild oRXAChildLevel3
            ' Alternate Name of Coding System -
            Set oRXAChildLevel3 = oDom.createElement("CE.6")
            oRXAChildLevel3.Text = "99NCI"
            oRXAChildLevel2.appendChild oRXAChildLevel3
        oRXAChildLevel1.appendChild oRXAChildLevel2
    oRXANode.appendChild oRXAChildLevel1
    
    ' System Entry Date/Time - CHSCII System Date/Time
    Set oRXAChildLevel1 = oDom.createElement("RXA.22")
    oRXAChildLevel1.Text = Format(Now, "YYYYMMDDHHNNSS")
    oRXANode.appendChild oRXAChildLevel1
    
    Set ReturnNewRXANode = oRXANode
     
Exit Function

ErrTrap:
   If gobjShared Is Nothing Then
      Set gobjShared = New CWShared
   End If
   gobjShared.ShowVBError Err.Number, Err.Description, "ReturnNewRXANode Function", "DMBE", vbExclamation
   'Err.Clear
   Resume Next 'There is no reason this procedure should not continue to create xml if it errors
     
End Function

Public Function ReturnNewZPBNode(ByVal lngCount As Long, ByVal ProblemsRS As ADODB.Recordset) As MSXML.IXMLDOMElement
   ' Problem Segment

' sample funcion output
'   <ZPB>
'      <ZPB.1>Count</ZPB.1>
'      <ZPB.2>Data ID</ZPB.2>
'      <ZPB.3>Facility NCID</ZPB.3>
'      <ZPB.4>Unit Number</ZPB.4>
'      <ZPB.5>Encounter Number</ZPB.5>
'      <ZPB.6>Problem NCID</ZPB.6>
'      <ZPB.7>Problem Comments</ZPB.7>
'      <ZPB.8>Problem Chronicity NCID</ZPB.8>
'      <ZPB.9>Problem Status NCID</ZPB.9>
'      <ZPB.10>Onset Date/Time</ZPB.10>
'      <ZPB.11>Source NCID</ZPB.11>
'      <ZPB.12>
'            <XCN.1>Clinician SSN</XCN.1>
'            <XCN.2>Clinician Last Name</XCN.2>
'            <XCN.3>Clinician First Name</XCN.3>
'            <XCN.4>Clinician Middle Name</XCN.4>
'      </ZPB.12>
'      <ZPB.13>Medcin ID</ZPB.13>
'      <ZPB.14>Point of Care Facility NCID</ZPB.14>
'      <ZPB.15>Noted Date</ZPB.15>
'      <ZPB.16>Create User</ZPB.16>
'      <ZPB.17>Create Date/Time</ZPB.17>
'      <ZPB.18>Modify User</ZPB.18>
'      <ZPB.19>Modify Date/Time</ZPB.19>
'      <ZPB.20>Update Flag</ZPB.20>
'      <ZPB.21>CDR Data ID</ZPB.21>
'      <ZPB.22>Created By</ZPB.22>
'      <ZPB.23>Created On</ZPB.23>
'      <ZPB.24>Updated By</ZPB.24>
'      <ZPB.25>Updated On</ZPB.25>
'   </ZPB>

    Dim oDom As MSXML.DOMDocument
    Dim oZPBNode As MSXML.IXMLDOMElement
    Dim oZPBChildLevel1 As MSXML.IXMLDOMElement
    Dim oZPBChildLevel2 As MSXML.IXMLDOMElement
    Dim strDocFirstName As String
    Dim strDocLastName As String
    Dim strDocMiddleName As String
    Dim strDocID As String
    Dim strDocSpecialty As String
    Dim strTemp As String
    Dim blnRsIsGood As Boolean
    
    On Error GoTo ErrTrap

    Set oDom = New MSXML.DOMDocument
    Set oZPBNode = oDom.createElement("ZPB")
    
    ' Count
    Set oZPBChildLevel1 = oDom.createElement("ZPB.1")
    oZPBChildLevel1.Text = CStr(lngCount)
    oZPBNode.appendChild oZPBChildLevel1

    ' Data ID - Table Problems:DataID
    Set oZPBChildLevel1 = oDom.createElement("ZPB.2")
    oZPBChildLevel1.Text = NullToNothing(ProblemsRS.Fields("DataID").Value)
    oZPBNode.appendChild oZPBChildLevel1

    ' Facility NCID - Table Problems:FacilityNCID
    Set oZPBChildLevel1 = oDom.createElement("ZPB.3")
    oZPBChildLevel1.Text = NullToNothing(ProblemsRS.Fields("FacilityNCID").Value)
    oZPBNode.appendChild oZPBChildLevel1

    ' Unit Number - Table Problems:Unit_Number
    Set oZPBChildLevel1 = oDom.createElement("ZPB.4")
    oZPBChildLevel1.Text = NullToNothing(ProblemsRS.Fields("Unit_Number").Value)
    oZPBNode.appendChild oZPBChildLevel1

    ' Encounter Number - Table Problems:EncounterNumber
    Set oZPBChildLevel1 = oDom.createElement("ZPB.5")
    oZPBChildLevel1.Text = NullToNothing(ProblemsRS.Fields("EncounterNumber").Value)
    oZPBNode.appendChild oZPBChildLevel1

    ' Problem NCID - Table Problems:ProblemNCID
    Set oZPBChildLevel1 = oDom.createElement("ZPB.6")
    'oZPBChildLevel1.Text = NullToNothing(ProblemsRS.Fields("ProblemNCID").Value)
        strTemp = NullToNothing(ProblemsRS.Fields("ProblemNCID").Value)
        
        Set oZPBChildLevel2 = oDom.createElement("CE.1")
        oZPBChildLevel1.appendChild oZPBChildLevel2
        
        Set oZPBChildLevel2 = oDom.createElement("CE.2")
        If strTemp <> "" Then
            oZPBChildLevel2.Text = GetNCIDDescription(CLng(strTemp))
        End If
        oZPBChildLevel1.appendChild oZPBChildLevel2
        
        Set oZPBChildLevel2 = oDom.createElement("CE.3")
        oZPBChildLevel1.appendChild oZPBChildLevel2
        Set oZPBChildLevel2 = oDom.createElement("CE.4")
        oZPBChildLevel1.appendChild oZPBChildLevel2
        Set oZPBChildLevel2 = oDom.createElement("CE.5")
        oZPBChildLevel2.Text = strTemp 'The NCID
        oZPBChildLevel1.appendChild oZPBChildLevel2
        Set oZPBChildLevel2 = oDom.createElement("CE.6")
        oZPBChildLevel2.Text = "99NCI"
        oZPBChildLevel1.appendChild oZPBChildLevel2
    oZPBNode.appendChild oZPBChildLevel1

    ' Problem Comments - Table Problems:ProblemsComment
    Set oZPBChildLevel1 = oDom.createElement("ZPB.7")
    oZPBChildLevel1.Text = NullToNothing(ProblemsRS.Fields("ProblemsComment").Value)
    oZPBNode.appendChild oZPBChildLevel1

    '<SCR 22973 CC>
    ' Problem Chronicity NCID - Table Problems:ChronicityNCID
    Set oZPBChildLevel1 = oDom.createElement("ZPB.8")
    'oZPBChildLevel1.Text = NullToNothing(ProblemsRS.Fields("ChronicityNCID").Value)
        strTemp = NullToNothing(ProblemsRS.Fields("ChronicityNCID").Value)
        
        Set oZPBChildLevel2 = oDom.createElement("CE.1")
        oZPBChildLevel1.appendChild oZPBChildLevel2
        
        Set oZPBChildLevel2 = oDom.createElement("CE.2")
        If strTemp <> "" Then
            oZPBChildLevel2.Text = GetNCIDDescription(CLng(strTemp))
        End If
        oZPBChildLevel1.appendChild oZPBChildLevel2
        
        Set oZPBChildLevel2 = oDom.createElement("CE.3")
        oZPBChildLevel1.appendChild oZPBChildLevel2
        Set oZPBChildLevel2 = oDom.createElement("CE.4")
        oZPBChildLevel1.appendChild oZPBChildLevel2
        Set oZPBChildLevel2 = oDom.createElement("CE.5")
        oZPBChildLevel2.Text = strTemp 'The NCID
        oZPBChildLevel1.appendChild oZPBChildLevel2
        Set oZPBChildLevel2 = oDom.createElement("CE.6")
        oZPBChildLevel2.Text = "99NCI"
        oZPBChildLevel1.appendChild oZPBChildLevel2
    oZPBNode.appendChild oZPBChildLevel1

    ' Problem Status NCID - Table Problems:StatusNCID
    Set oZPBChildLevel1 = oDom.createElement("ZPB.9")
    'oZPBChildLevel1.Text = NullToNothing(ProblemsRS.Fields("StatusNCID").Value)
        strTemp = NullToNothing(ProblemsRS.Fields("StatusNCID").Value)
        
        Set oZPBChildLevel2 = oDom.createElement("CE.1")
        oZPBChildLevel1.appendChild oZPBChildLevel2
        
        Set oZPBChildLevel2 = oDom.createElement("CE.2")
        If strTemp <> "" Then
            oZPBChildLevel2.Text = GetNCIDDescription(CLng(strTemp))
        End If
        oZPBChildLevel1.appendChild oZPBChildLevel2
        
        Set oZPBChildLevel2 = oDom.createElement("CE.3")
        oZPBChildLevel1.appendChild oZPBChildLevel2
        Set oZPBChildLevel2 = oDom.createElement("CE.4")
        oZPBChildLevel1.appendChild oZPBChildLevel2
        Set oZPBChildLevel2 = oDom.createElement("CE.5")
        oZPBChildLevel2.Text = strTemp 'The NCID
        oZPBChildLevel1.appendChild oZPBChildLevel2
        Set oZPBChildLevel2 = oDom.createElement("CE.6")
        oZPBChildLevel2.Text = "99NCI"
        oZPBChildLevel1.appendChild oZPBChildLevel2
    oZPBNode.appendChild oZPBChildLevel1

    ' Onset Date/Time - Table Problems:OnSetDate
    Set oZPBChildLevel1 = oDom.createElement("ZPB.10")
    strTemp = NullToNothing(ProblemsRS.Fields("OnSetDate").Value)
    If IsValidDate(strTemp) Then
        oZPBChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZPBNode.appendChild oZPBChildLevel1
    
    ' Source NCID - Table Problems:SourceNCID
    Set oZPBChildLevel1 = oDom.createElement("ZPB.11")
    'oZPBChildLevel1.Text = NullToNothing(ProblemsRS.Fields("SourceNCID").Value)
        strTemp = NullToNothing(ProblemsRS.Fields("SourceNCID").Value)
        
        Set oZPBChildLevel2 = oDom.createElement("CE.1")
        oZPBChildLevel1.appendChild oZPBChildLevel2
        
        Set oZPBChildLevel2 = oDom.createElement("CE.2")
        If strTemp <> "" Then
            oZPBChildLevel2.Text = GetNCIDDescription(CLng(strTemp))
        End If
        oZPBChildLevel1.appendChild oZPBChildLevel2
        
        Set oZPBChildLevel2 = oDom.createElement("CE.3")
        oZPBChildLevel1.appendChild oZPBChildLevel2
        Set oZPBChildLevel2 = oDom.createElement("CE.4")
        oZPBChildLevel1.appendChild oZPBChildLevel2
        Set oZPBChildLevel2 = oDom.createElement("CE.5")
        oZPBChildLevel2.Text = strTemp 'The NCID
        oZPBChildLevel1.appendChild oZPBChildLevel2
        Set oZPBChildLevel2 = oDom.createElement("CE.6")
        oZPBChildLevel2.Text = "99NCI"
        oZPBChildLevel1.appendChild oZPBChildLevel2
    oZPBNode.appendChild oZPBChildLevel1
    
    ' Clinician
    Set oZPBChildLevel1 = oDom.createElement("ZPB.12")
        
        blnRsIsGood = False
        strTemp = NullToNothing(ProblemsRS.Fields("ClinicianNCID").Value)
        If Not strTemp = "" Then
            Call GetNewDoctorInfo(CLng(strTemp), strDocLastName, strDocFirstName, strDocMiddleName, strDocID, strDocSpecialty)
            blnRsIsGood = True
        End If
            
        ' Clinician ID
        Set oZPBChildLevel2 = oDom.createElement("XCN.1")
        If blnRsIsGood = True Then
            oZPBChildLevel2.Text = strDocID
        End If
        oZPBChildLevel1.appendChild oZPBChildLevel2
        ' Clinician Last Name
        Set oZPBChildLevel2 = oDom.createElement("XCN.2")
        If blnRsIsGood = True Then
            oZPBChildLevel2.Text = strDocLastName
        End If
        oZPBChildLevel1.appendChild oZPBChildLevel2
        ' Clinician first Name
        Set oZPBChildLevel2 = oDom.createElement("XCN.3")
        If blnRsIsGood = True Then
            oZPBChildLevel2.Text = strDocFirstName
        End If
        oZPBChildLevel1.appendChild oZPBChildLevel2
        ' Clinician Middle Name
        Set oZPBChildLevel2 = oDom.createElement("XCN.4")
        If blnRsIsGood = True Then
            oZPBChildLevel2.Text = strDocMiddleName
        End If
        oZPBChildLevel1.appendChild oZPBChildLevel2
    
    oZPBNode.appendChild oZPBChildLevel1
    
    ' SnoID - Table Problems:SnoID
    Set oZPBChildLevel1 = oDom.createElement("ZPB.13")
    oZPBChildLevel1.Text = NullToNothing(ProblemsRS.Fields("SnoID").Value)
    oZPBNode.appendChild oZPBChildLevel1
    
    ' Point of Care Facility - Table Problems:PointOfCareFacilityNCID
    Set oZPBChildLevel1 = oDom.createElement("ZPB.14")
    'oZPBChildLevel1.Text = NullToNothing(ProblemsRS.Fields("PointOfCareFacilityNCID").Value)
        strTemp = NullToNothing(ProblemsRS.Fields("PointOfCareFacilityNCID").Value)
        
        Set oZPBChildLevel2 = oDom.createElement("CE.1")
        oZPBChildLevel1.appendChild oZPBChildLevel2
        
        Set oZPBChildLevel2 = oDom.createElement("CE.2")
        If strTemp <> "" Then
            oZPBChildLevel2.Text = GetNCIDDescription(CLng(strTemp))
        End If
        oZPBChildLevel1.appendChild oZPBChildLevel2
        
        Set oZPBChildLevel2 = oDom.createElement("CE.3")
        oZPBChildLevel1.appendChild oZPBChildLevel2
        Set oZPBChildLevel2 = oDom.createElement("CE.4")
        oZPBChildLevel1.appendChild oZPBChildLevel2
        Set oZPBChildLevel2 = oDom.createElement("CE.5")
        oZPBChildLevel2.Text = strTemp 'The NCID
        oZPBChildLevel1.appendChild oZPBChildLevel2
        Set oZPBChildLevel2 = oDom.createElement("CE.6")
        oZPBChildLevel2.Text = "99NCI"
        oZPBChildLevel1.appendChild oZPBChildLevel2
    oZPBNode.appendChild oZPBChildLevel1
    '</SCR 22973 CC>
    
    ' Noted Date - Table Problems:NotedDate
    Set oZPBChildLevel1 = oDom.createElement("ZPB.15")
    strTemp = NullToNothing(ProblemsRS.Fields("NotedDate").Value)
    If IsValidDate(strTemp) Then
        oZPBChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZPBNode.appendChild oZPBChildLevel1
    
    ' Create User - Table Problems:CreateUserNCID
    Set oZPBChildLevel1 = oDom.createElement("ZPB.16")
    oZPBChildLevel1.Text = NullToNothing(ProblemsRS.Fields("CreateUserNCID").Value)
    oZPBNode.appendChild oZPBChildLevel1
    
    ' Create Date/Time - Table Problems:CreateTime
    Set oZPBChildLevel1 = oDom.createElement("ZPB.17")
    strTemp = NullToNothing(ProblemsRS.Fields("CreateTime").Value)
    If IsValidDate(strTemp) Then
        oZPBChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZPBNode.appendChild oZPBChildLevel1
    
    ' Modify User - Table Problems:ModifyUserNCID
    Set oZPBChildLevel1 = oDom.createElement("ZPB.18")
    oZPBChildLevel1.Text = NullToNothing(ProblemsRS.Fields("ModifyUserNCID").Value)
    oZPBNode.appendChild oZPBChildLevel1
    
    ' Modify Date/Time - Table Problems:ModifyTime
    Set oZPBChildLevel1 = oDom.createElement("ZPB.19")
    strTemp = NullToNothing(ProblemsRS.Fields("ModifyTime").Value)
    If IsValidDate(strTemp) Then
        oZPBChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZPBNode.appendChild oZPBChildLevel1
    
    ' Update Flag - Table Problems:UpdateFlag
    Set oZPBChildLevel1 = oDom.createElement("ZPB.20")
    oZPBChildLevel1.Text = NullToNothing(ProblemsRS.Fields("UpdateFlag").Value)
    oZPBNode.appendChild oZPBChildLevel1
    
    ' CDR Data ID - Table Problems:CDR_DataID
    Set oZPBChildLevel1 = oDom.createElement("ZPB.21")
    oZPBChildLevel1.Text = NullToNothing(ProblemsRS.Fields("CDR_DataID").Value)
    oZPBNode.appendChild oZPBChildLevel1
    
    ' Created By - Table Problems:CreatedBy
    Set oZPBChildLevel1 = oDom.createElement("ZPB.22")
    oZPBChildLevel1.Text = NullToNothing(ProblemsRS.Fields("CreatedBy").Value)
    oZPBNode.appendChild oZPBChildLevel1
    
    ' Created On - Table Problems:CreatedOn
    Set oZPBChildLevel1 = oDom.createElement("ZPB.23")
    strTemp = NullToNothing(ProblemsRS.Fields("CreatedOn").Value)
    If IsValidDate(strTemp) Then
        oZPBChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZPBNode.appendChild oZPBChildLevel1
    
    ' Updated By - Table Problems:UpdatedBy
    Set oZPBChildLevel1 = oDom.createElement("ZPB.24")
    oZPBChildLevel1.Text = NullToNothing(ProblemsRS.Fields("UpdatedBy").Value)
    oZPBNode.appendChild oZPBChildLevel1
    
    ' Created On - Table Problems:CreatedOn
    Set oZPBChildLevel1 = oDom.createElement("ZPB.25")
    strTemp = NullToNothing(ProblemsRS.Fields("UpdatedOn").Value)
    If IsValidDate(strTemp) Then
        oZPBChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZPBNode.appendChild oZPBChildLevel1
    
    Set oZPBChildLevel1 = oDom.createElement("ZPB.26")
    oZPBChildLevel1.Text = NullToNothing(ProblemsRS.Fields("PREFIX").Value)
    oZPBNode.appendChild oZPBChildLevel1
    
    Set ReturnNewZPBNode = oZPBNode

Exit Function

ErrTrap:
   If gobjShared Is Nothing Then
      Set gobjShared = New CWShared
   End If
   gobjShared.ShowVBError Err.Number, Err.Description, "ReturnNewZPBNode Function", "DMBE", vbExclamation
   Resume Next 'There is no reason this procedure should not continue to create xml if it errors
    
End Function

'Public Function ReturnNewZVTNode(ByVal lngCount As Long, ByVal VitalsRS As ADODB.Recordset, ByVal VitalSignsModRS As ADODB.Recordset, ByVal VitalSignsValRS As ADODB.Recordset) As MSXML.IXMLDOMElement
'' Vitals Segment
'' This function uses a joined recordset from Vitals, VitalSignsModifier, VitalSignsValue, and Encounters table
'' sample function output
''   <ZVT>
''      <ZVT.1>Count</ZVT.1>
''      <ZVT.2>Vitals Data ID</ZVT.2>
''      <ZVT.3>Facility NCID</ZVT.3>
''      <ZVT.4>Unit Number</ZVT.4>
''      <ZVT.5>Encounter Number</ZVT.5>
''      <ZVT.6>
''            <XCN.1>Clinician SSN</XCN.1>
''            <XCN.2>Clinician Last Name</XCN.2>
''            <XCN.3>Clinician First Name</XCN.3>
''            <XCN.4>Clinician Middle Name</XCN.4>
''      </ZVT.6>
''      <ZVT.7>Panel Type</ZVT.7>
''      <ZVT.8>Entry Date</ZVT.8>
''      <ZVT.9>Comment</ZVT.9>
''      <ZVT.10>Update Flag</ZVT.10>
''      <ZVT.11>CDR Data ID</ZVT.11>
''      <ZVT.12>Vital Signs Modifier Data ID</ZVT.12>
''      <ZVT.13>Vital Signs Modifier Vitals Type</ZVT.13>
''      <ZVT.14>Observation Modifier</ZVT.14>
''      <ZVT.15>Vital Signs Modifier Type</ZVT.15>
''      <ZVT.16>Vital Signs Modifier Domain</ZVT.16>
''      <ZVT.17>Vital Signs Modifier Value</ZVT.17>
''      <ZVT.18>Vital Signs Modifier Body Side</ZVT.18>
''      <ZVT.19>Vital Signs Modifier Agg Obs</ZVT.19>
''      <ZVT.20>Vital Signs Value Data ID</ZVT.20>
''      <ZVT.21>Type</ZVT.21>
''      <ZVT.22>Value</ZVT.22>
''      <ZVT.23>Units</ZVT.23>
''   </ZVT>
'
'    Dim oDom As MSXML.DOMDocument
'    Dim oZVTNode As MSXML.IXMLDOMElement
'    Dim oZVTChildLevel1 As MSXML.IXMLDOMElement
'    Dim oZVTChildLevel2 As MSXML.IXMLDOMElement
'    Dim oZVTChildLevel3 As MSXML.IXMLDOMElement
'    Dim strDocFirstName As String
'    Dim strDocLastName As String
'    Dim strDocMiddleName As String
'    Dim strDocID As String
'    Dim strDocSpecialty As String
'    Dim strTemp As String
'    Dim blnRSIsGood As Boolean
'    Dim blnModRsIsEOF As Boolean
'
'    On Error GoTo ErrTrap
'
'    Set oDom = New MSXML.DOMDocument
'    Set oZVTNode = oDom.createElement("ZVT")
'
'    ' Count
'    Set oZVTChildLevel1 = oDom.createElement("ZVT.1")
'    oZVTChildLevel1.Text = CStr(lngCount)
'    oZVTNode.appendChild oZVTChildLevel1
'
'    ' Data ID - Table Vitals:DataID
'    Set oZVTChildLevel1 = oDom.createElement("ZVT.2")
'    oZVTChildLevel1.Text = NullToNothing(VitalsRS.Fields("DataID").Value)
'    oZVTNode.appendChild oZVTChildLevel1
'
'    ' Facility NCID - Table Vitals:FacilityNCID
'    Set oZVTChildLevel1 = oDom.createElement("ZVT.3")
'    oZVTChildLevel1.Text = NullToNothing(VitalsRS.Fields("FacilityNCID").Value)
'    oZVTNode.appendChild oZVTChildLevel1
'
'    ' Unit Number - Table Vitals:Unit_Number
'    Set oZVTChildLevel1 = oDom.createElement("ZVT.4")
'    oZVTChildLevel1.Text = NullToNothing(VitalsRS.Fields("Unit_Number").Value)
'    oZVTNode.appendChild oZVTChildLevel1
'
'    ' Encounter Number - Table Vitals:EncounterNumber
'    Set oZVTChildLevel1 = oDom.createElement("ZVT.5")
'    oZVTChildLevel1.Text = NullToNothing(VitalsRS.Fields("EncounterNumber").Value)
'    oZVTNode.appendChild oZVTChildLevel1
'
'    ' Clinician
'    Set oZVTChildLevel1 = oDom.createElement("ZVT.6")
'
'        blnRSIsGood = False
'        strTemp = NullToNothing(VitalsRS.Fields("ClinicianNCID").Value)
'        If Not strTemp = "" Then
'            Call GetNewDoctorInfo(CLng(strTemp), strDocLastName, strDocFirstName, strDocMiddleName, strDocID, strDocSpecialty)
'            blnRSIsGood = True
'        End If
'
'        ' Clinician ID
'        Set oZVTChildLevel2 = oDom.createElement("XCN.1")
'        If blnRSIsGood = True Then
'            oZVTChildLevel2.Text = strDocID
'        End If
'        oZVTChildLevel1.appendChild oZVTChildLevel2
'        ' Clinician Last Name
'        Set oZVTChildLevel2 = oDom.createElement("XCN.2")
'        If blnRSIsGood = True Then
'            oZVTChildLevel2.Text = strDocLastName
'        End If
'        oZVTChildLevel1.appendChild oZVTChildLevel2
'        ' Clinician first Name
'        Set oZVTChildLevel2 = oDom.createElement("XCN.3")
'        If blnRSIsGood = True Then
'            oZVTChildLevel2.Text = strDocFirstName
'        End If
'        oZVTChildLevel1.appendChild oZVTChildLevel2
'        ' Clinician Middle Name
'        Set oZVTChildLevel2 = oDom.createElement("XCN.4")
'        If blnRSIsGood = True Then
'            oZVTChildLevel2.Text = strDocMiddleName
'        End If
'        oZVTChildLevel1.appendChild oZVTChildLevel2
'
'    oZVTNode.appendChild oZVTChildLevel1
'
'    ' Panel Type - Table Vitals:PanelType
'    Set oZVTChildLevel1 = oDom.createElement("ZVT.7")
'    oZVTChildLevel1.Text = NullToNothing(VitalsRS.Fields("PanelType").Value)
'    oZVTNode.appendChild oZVTChildLevel1
'
'    ' Entry Date - Table Vitals:EntryDate
'    Set oZVTChildLevel1 = oDom.createElement("ZVT.8")
'    strTemp = NullToNothing(VitalsRS.Fields("EntryDate").Value)
'    If IsValidDate(strTemp) Then
'        oZVTChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
'    End If
'    oZVTNode.appendChild oZVTChildLevel1
'
'    ' Comment - Table Vitals:VitalsComment
'    Set oZVTChildLevel1 = oDom.createElement("ZVT.9")
'    oZVTChildLevel1.Text = NullToNothing(VitalsRS.Fields("VitalsComment").Value)
'    oZVTNode.appendChild oZVTChildLevel1
'
'    ' Update Flag - Table Vitals:UpdateFlag
'    Set oZVTChildLevel1 = oDom.createElement("ZVT.10")
'    oZVTChildLevel1.Text = NullToNothing(VitalsRS.Fields("UpdateFlag").Value)
'    oZVTNode.appendChild oZVTChildLevel1
'
'    ' CDR Data ID - Table Vitals:CDR_DataID
'    Set oZVTChildLevel1 = oDom.createElement("ZVT.11")
'    oZVTChildLevel1.Text = NullToNothing(VitalsRS.Fields("CDR_DataID").Value)
'    oZVTNode.appendChild oZVTChildLevel1
'
'    blnModRsIsEOF = True
'
'    If Not VitalSignsModRS Is Nothing Then
'        If VitalSignsModRS.RecordCount > 0 Then
'            blnModRsIsEOF = False
'        End If
'    End If
'
'    ' Data ID - Table VitalSignsModifier:DataID
'    Set oZVTChildLevel1 = oDom.createElement("ZVT.12")
'    If blnModRsIsEOF = False Then
'        oZVTChildLevel1.Text = NullToNothing(VitalSignsModRS.Fields("DataID").Value)
'    End If
'    oZVTNode.appendChild oZVTChildLevel1
'
'    ' Vitals Type - Table VitalSignsModifier:VitalsType
'    Set oZVTChildLevel1 = oDom.createElement("ZVT.13")
'    If blnModRsIsEOF = False Then
'        oZVTChildLevel1.Text = NullToNothing(VitalSignsModRS.Fields("VitalsType").Value)
'    End If
'    oZVTNode.appendChild oZVTChildLevel1
'
'    blnRSIsGood = False
'    If Not (VitalSignsModRS Is Nothing) Then
'        If Not VitalSignsModRS.EOF Then
'            blnRSIsGood = True
'        End If
'    End If
'
'    ' Observation Modifier - Table VitalSignsModifier:ObservationModifier
'    Set oZVTChildLevel1 = oDom.createElement("ZVT.14")
'    If blnRSIsGood = True Then
'        oZVTChildLevel1.Text = NullToNothing(VitalSignsModRS.Fields("OBSERVATIONMODIFIER").Value)
'    End If
'    oZVTNode.appendChild oZVTChildLevel1
'
'    ' Type - Table VitalSignsModifier:Type
'    Set oZVTChildLevel1 = oDom.createElement("ZVT.15")
'    If blnRSIsGood = True Then
'        oZVTChildLevel1.Text = NullToNothing(VitalSignsModRS.Fields("Type").Value)
'    End If
'    oZVTNode.appendChild oZVTChildLevel1
'
'    ' Domain - Table VitalSignsModifier:Domain
'    Set oZVTChildLevel1 = oDom.createElement("ZVT.16")
'    If blnRSIsGood = True Then
'        oZVTChildLevel1.Text = NullToNothing(VitalSignsModRS.Fields("Domain").Value)
'    End If
'    oZVTNode.appendChild oZVTChildLevel1
'
'    ' Modifier Value - Table VitalSignsModifier:ModifierValue
'    Set oZVTChildLevel1 = oDom.createElement("ZVT.17")
'    If blnRSIsGood = True Then
'        oZVTChildLevel1.Text = NullToNothing(VitalSignsModRS.Fields("ModifierValue").Value)
'    End If
'    oZVTNode.appendChild oZVTChildLevel1
'
'    ' Body Side - Table VitalSignsModifier:BodySide
'    Set oZVTChildLevel1 = oDom.createElement("ZVT.18")
'    If blnRSIsGood = True Then
'        oZVTChildLevel1.Text = NullToNothing(VitalSignsModRS.Fields("BodySide").Value)
'    End If
'    oZVTNode.appendChild oZVTChildLevel1
'
'    ' Agg Obs - Table VitalSignsModifier:AggObs
'    Set oZVTChildLevel1 = oDom.createElement("ZVT.19")
'    If blnRSIsGood = True Then
'        oZVTChildLevel1.Text = NullToNothing(VitalSignsModRS.Fields("AggObs").Value)
'    End If
'    oZVTNode.appendChild oZVTChildLevel1
'
'    blnRSIsGood = False
'    If Not VitalSignsValRS Is Nothing Then
'        If VitalSignsValRS.RecordCount > 0 Then
'            If Not VitalSignsValRS.EOF Then
'                blnRSIsGood = True
'            End If
'        End If
'    End If
'
'    ' Data ID - Table VitalSignsValue:DataID
'    Set oZVTChildLevel1 = oDom.createElement("ZVT.20")
'    If blnRSIsGood = True Then
'        oZVTChildLevel1.Text = NullToNothing(VitalSignsValRS.Fields("DataID").Value)
'    End If
'    oZVTNode.appendChild oZVTChildLevel1
'
'    ' Type - Table VitalSignsValue:VitalsType
'    Set oZVTChildLevel1 = oDom.createElement("ZVT.21")
'    If blnRSIsGood = True Then
'        oZVTChildLevel1.Text = NullToNothing(VitalSignsValRS.Fields("VitalsType").Value)
'    End If
'    oZVTNode.appendChild oZVTChildLevel1
'
'    ' Value - Table VitalSignsValue:VitalsValue
'    Set oZVTChildLevel1 = oDom.createElement("ZVT.22")
'    If blnRSIsGood = True Then
'        oZVTChildLevel1.Text = NullToNothing(VitalSignsValRS.Fields("VitalsValue").Value)
'    End If
'    oZVTNode.appendChild oZVTChildLevel1
'
'    ' Units - Table VitalSignsValue:Units
'    Set oZVTChildLevel1 = oDom.createElement("ZVT.23")
'    If blnRSIsGood = True Then
'        oZVTChildLevel1.Text = NullToNothing(VitalSignsValRS.Fields("Units").Value)
'    End If
'    oZVTNode.appendChild oZVTChildLevel1
'
'    Set ReturnNewZVTNode = oZVTNode
'
'Exit Function
'
'ErrTrap:
'   If gobjShared Is Nothing Then
'      Set gobjShared = New CWShared
'   End If
'   gobjShared.ShowVBError Err.Number, Err.Description, "ReturnNewZVTNode Function", "DMBE", vbExclamation
'   'Err.Clear
'
'End Function

Public Function ReturnNewZVTNode(ByVal lngCount As Long, ByVal VitalsRS As ADODB.Recordset, ByVal VitalSignsModRS As ADODB.Recordset, ByVal VitalSignsValRS As ADODB.Recordset) As MSXML.IXMLDOMElement
' Vitals Segment
' This function uses a joined recordset from Vitals, VitalSignsModifier, VitalSignsValue, and Encounters table
' sample function output
'   <ZVT>
'      <ZVT.1>Count</ZVT.1>
'      <ZVT.2>Vitals Data ID</ZVT.2>
'      <ZVT.3>Facility NCID</ZVT.3>
'      <ZVT.4>Unit Number</ZVT.4>
'      <ZVT.5>Encounter Number</ZVT.5>
'      <ZVT.6>
'            <XCN.1>Clinician SSN</XCN.1>
'            <XCN.2>Clinician Last Name</XCN.2>
'            <XCN.3>Clinician First Name</XCN.3>
'            <XCN.4>Clinician Middle Name</XCN.4>
'      </ZVT.6>
'      <ZVT.7>Panel Type</ZVT.7>
'      <ZVT.8>Entry Date</ZVT.8>
'      <ZVT.9>Comment</ZVT.9>
'      <ZVT.10>Update Flag</ZVT.10>
'      <ZVT.11>CDR Data ID</ZVT.11>
'      <ZVT.12>Vital Signs Modifier Data ID</ZVT.12>
'      <ZVT.13>Vital Signs Modifier Vitals Type</ZVT.13>
'      <ZVT.14>Observation Modifier</ZVT.14>
'      <ZVT.15>Vital Signs Modifier Type</ZVT.15>
'      <ZVT.16>Vital Signs Modifier Domain</ZVT.16>
'      <ZVT.17>Vital Signs Modifier Value</ZVT.17>
'      <ZVT.18>Vital Signs Modifier Body Side</ZVT.18>
'      <ZVT.19>Vital Signs Modifier Agg Obs</ZVT.19>
'      <ZVT.20>Vital Signs Value Data ID</ZVT.20>
'      <ZVT.21>Type</ZVT.21>
'      <ZVT.22>Value</ZVT.22>
'      <ZVT.23>Units</ZVT.23>
'   </ZVT>
   
    Dim oDom As MSXML.DOMDocument
    Dim oZVTNode As MSXML.IXMLDOMElement
    Dim oZVTChildLevel1 As MSXML.IXMLDOMElement
    Dim oZVTChildLevel2 As MSXML.IXMLDOMElement
    Dim oZVTChildLevel3 As MSXML.IXMLDOMElement
    Dim strDocFirstName As String
    Dim strDocLastName As String
    Dim strDocMiddleName As String
    Dim strDocID As String
    Dim strDocSpecialty As String
    Dim strTemp As String
    Dim blnRsIsGood As Boolean
    
    On Error GoTo ErrTrap

    Set oDom = New MSXML.DOMDocument
    Set oZVTNode = oDom.createElement("ZVT")
    
    ' Count
    Set oZVTChildLevel1 = oDom.createElement("ZVT.1")
    oZVTChildLevel1.Text = CStr(lngCount)
    oZVTNode.appendChild oZVTChildLevel1

    ' Data ID - Table Vitals:DataID
    Set oZVTChildLevel1 = oDom.createElement("ZVT.2")
    oZVTChildLevel1.Text = NullToNothing(VitalsRS.Fields("DataID").Value)
    oZVTNode.appendChild oZVTChildLevel1

    ' Facility NCID - Table Vitals:FacilityNCID
    Set oZVTChildLevel1 = oDom.createElement("ZVT.3")
    oZVTChildLevel1.Text = NullToNothing(VitalsRS.Fields("FacilityNCID").Value)
    oZVTNode.appendChild oZVTChildLevel1

    ' Unit Number - Table Vitals:Unit_Number
    Set oZVTChildLevel1 = oDom.createElement("ZVT.4")
    oZVTChildLevel1.Text = NullToNothing(VitalsRS.Fields("Unit_Number").Value)
    oZVTNode.appendChild oZVTChildLevel1
    
    ' Encounter Number - Table Vitals:EncounterNumber
    Set oZVTChildLevel1 = oDom.createElement("ZVT.5")
    oZVTChildLevel1.Text = NullToNothing(VitalsRS.Fields("EncounterNumber").Value)
    oZVTNode.appendChild oZVTChildLevel1
    
    ' Clinician
    Set oZVTChildLevel1 = oDom.createElement("ZVT.6")
        
        blnRsIsGood = False
        strTemp = NullToNothing(VitalsRS.Fields("ClinicianNCID").Value)
        If Not strTemp = "" Then
            Call GetNewDoctorInfo(CLng(strTemp), strDocLastName, strDocFirstName, strDocMiddleName, strDocID, strDocSpecialty)
            blnRsIsGood = True
        End If
            
        ' Clinician ID
        Set oZVTChildLevel2 = oDom.createElement("XCN.1")
        If blnRsIsGood = True Then
            oZVTChildLevel2.Text = strDocID
        End If
        oZVTChildLevel1.appendChild oZVTChildLevel2
        ' Clinician Last Name
        Set oZVTChildLevel2 = oDom.createElement("XCN.2")
        If blnRsIsGood = True Then
            oZVTChildLevel2.Text = strDocLastName
        End If
        oZVTChildLevel1.appendChild oZVTChildLevel2
        ' Clinician first Name
        Set oZVTChildLevel2 = oDom.createElement("XCN.3")
        If blnRsIsGood = True Then
            oZVTChildLevel2.Text = strDocFirstName
        End If
        oZVTChildLevel1.appendChild oZVTChildLevel2
        ' Clinician Middle Name
        Set oZVTChildLevel2 = oDom.createElement("XCN.4")
        If blnRsIsGood = True Then
            oZVTChildLevel2.Text = strDocMiddleName
        End If
        oZVTChildLevel1.appendChild oZVTChildLevel2
        
    oZVTNode.appendChild oZVTChildLevel1
    
    ' Panel Type - Table Vitals:PanelType
    Set oZVTChildLevel1 = oDom.createElement("ZVT.7")
    oZVTChildLevel1.Text = NullToNothing(VitalsRS.Fields("PanelType").Value)
    oZVTNode.appendChild oZVTChildLevel1
    
    ' Entry Date - Table Vitals:EntryDate
    Set oZVTChildLevel1 = oDom.createElement("ZVT.8")
    strTemp = NullToNothing(VitalsRS.Fields("EntryDate").Value)
    If IsValidDate(strTemp) Then
        oZVTChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZVTNode.appendChild oZVTChildLevel1
    
    ' Comment - Table Vitals:VitalsComment
    Set oZVTChildLevel1 = oDom.createElement("ZVT.9")
    oZVTChildLevel1.Text = NullToNothing(VitalsRS.Fields("VitalsComment").Value)
    oZVTNode.appendChild oZVTChildLevel1
    
    ' Update Flag - Table Vitals:UpdateFlag
    Set oZVTChildLevel1 = oDom.createElement("ZVT.10")
    oZVTChildLevel1.Text = NullToNothing(VitalsRS.Fields("UpdateFlag").Value)
    oZVTNode.appendChild oZVTChildLevel1
    
    ' CDR Data ID - Table Vitals:CDR_DataID
    Set oZVTChildLevel1 = oDom.createElement("ZVT.11")
    oZVTChildLevel1.Text = NullToNothing(VitalsRS.Fields("CDR_DataID").Value)
    oZVTNode.appendChild oZVTChildLevel1
    
    Set oZVTChildLevel1 = oDom.createElement("ZVT.12")
    oZVTChildLevel1.Text = NullToNothing(VitalsRS.Fields("CREATEDBY").Value)
    oZVTNode.appendChild oZVTChildLevel1
    
    Set oZVTChildLevel1 = oDom.createElement("ZVT.13")
    strTemp = NullToNothing(VitalsRS.Fields("CREATEDON").Value)
    If IsValidDate(strTemp) Then
        oZVTChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZVTNode.appendChild oZVTChildLevel1
    
    Set oZVTChildLevel1 = oDom.createElement("ZVT.14")
    oZVTChildLevel1.Text = NullToNothing(VitalsRS.Fields("UPDATEDBY").Value)
    oZVTNode.appendChild oZVTChildLevel1
    
    Set oZVTChildLevel1 = oDom.createElement("ZVT.15")
    strTemp = NullToNothing(VitalsRS.Fields("CREATEDON").Value)
    If IsValidDate(strTemp) Then
        oZVTChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZVTNode.appendChild oZVTChildLevel1
    
    'New Subsection based on how many value there is
    'Vital Signs Value
    Set oZVTChildLevel1 = oDom.createElement("ZVT.16.LST")
        
        blnRsIsGood = False
        If Not (VitalSignsValRS Is Nothing) Then
            If Not VitalSignsValRS.EOF Then
                blnRsIsGood = True
            End If
        End If
    
        If blnRsIsGood = True Then
            
            Do Until VitalSignsValRS.EOF
            
                Set oZVTChildLevel2 = oDom.createElement("ZVT.16")
                    
                    ' Data ID - Table VitalSignsValue:DataID
                    Set oZVTChildLevel3 = oDom.createElement("ZVT.16.1")
                    oZVTChildLevel3.Text = NullToNothing(VitalSignsValRS.Fields("DataID").Value)
                    oZVTChildLevel2.appendChild oZVTChildLevel3
                    
                    ' Type - Table VitalSignsValue:VitalsType
                    Set oZVTChildLevel3 = oDom.createElement("ZVT.16.2")
                    oZVTChildLevel3.Text = NullToNothing(VitalSignsValRS.Fields("VitalsType").Value)
                    oZVTChildLevel2.appendChild oZVTChildLevel3
                    
                    ' Value - Table VitalSignsValue:VitalsValue
                    Set oZVTChildLevel3 = oDom.createElement("ZVT.16.3")
                    oZVTChildLevel3.Text = NullToNothing(VitalSignsValRS.Fields("VitalsValue").Value)
                    oZVTChildLevel2.appendChild oZVTChildLevel3
                    
                    ' Units - Table VitalSignsValue:Units
                    Set oZVTChildLevel3 = oDom.createElement("ZVT.16.4")
                    oZVTChildLevel3.Text = NullToNothing(VitalSignsValRS.Fields("Units").Value)
                    oZVTChildLevel2.appendChild oZVTChildLevel3
                
                    Set oZVTChildLevel3 = oDom.createElement("ZVT.16.5")
                    oZVTChildLevel3.Text = NullToNothing(VitalSignsValRS.Fields("CREATEDBY").Value)
                    oZVTChildLevel2.appendChild oZVTChildLevel3
                    
                    Set oZVTChildLevel3 = oDom.createElement("ZVT.16.6")
                    strTemp = NullToNothing(VitalSignsValRS.Fields("CREATEDON").Value)
                    If IsValidDate(strTemp) Then
                        oZVTChildLevel3.Text = Format(strTemp, "YYYYMMDDHHNNSS")
                    End If
                    oZVTChildLevel2.appendChild oZVTChildLevel3
                    
                    Set oZVTChildLevel3 = oDom.createElement("ZVT.16.7")
                    oZVTChildLevel3.Text = NullToNothing(VitalSignsValRS.Fields("UPDATEDBY").Value)
                    oZVTChildLevel2.appendChild oZVTChildLevel3
                    
                    Set oZVTChildLevel3 = oDom.createElement("ZVT.16.8")
                    strTemp = NullToNothing(VitalSignsValRS.Fields("UPDATEDON").Value)
                    If IsValidDate(strTemp) Then
                        oZVTChildLevel3.Text = Format(strTemp, "YYYYMMDDHHNNSS")
                    End If
                    oZVTChildLevel2.appendChild oZVTChildLevel3
                
                oZVTChildLevel1.appendChild oZVTChildLevel2
                VitalSignsValRS.MoveNext
            Loop
        End If
    oZVTNode.appendChild oZVTChildLevel1
    
    'Vital Signs Modifier
    
    Set oZVTChildLevel1 = oDom.createElement("ZVT.17.LST")
        
        blnRsIsGood = False
        If Not VitalSignsModRS Is Nothing Then
            If VitalSignsModRS.RecordCount > 0 Then
                blnRsIsGood = True
            End If
        End If
        
        If blnRsIsGood = True Then
            
            Do Until VitalSignsModRS.EOF
                Set oZVTChildLevel2 = oDom.createElement("ZVT.17")
                    
                    ' Data ID - Table VitalSignsValue:DataID
                    Set oZVTChildLevel3 = oDom.createElement("ZVT.17.1")
                    oZVTChildLevel3.Text = NullToNothing(VitalSignsModRS.Fields("DataID").Value)
                    oZVTChildLevel2.appendChild oZVTChildLevel3
                    
                    ' Vitals Type - Table VitalSignsModifier:VitalsType
                    Set oZVTChildLevel3 = oDom.createElement("ZVT.17.2")
                    oZVTChildLevel3.Text = NullToNothing(VitalSignsModRS.Fields("VitalsType").Value)
                    oZVTChildLevel2.appendChild oZVTChildLevel3
                    
                    ' Observation Modifier - Table VitalSignsModifier:ObservationModifier
                    Set oZVTChildLevel3 = oDom.createElement("ZVT.17.3")
                    oZVTChildLevel3.Text = NullToNothing(VitalSignsModRS.Fields("OBSERVATIONMODIFIER").Value)
                    oZVTChildLevel2.appendChild oZVTChildLevel3
                    
                    ' Type - Table VitalSignsModifier:Type
                    Set oZVTChildLevel3 = oDom.createElement("ZVT.17.4")
                    oZVTChildLevel3.Text = NullToNothing(VitalSignsModRS.Fields("Type").Value)
                    oZVTChildLevel2.appendChild oZVTChildLevel3
                    
                    ' Domain - Table VitalSignsModifier:Domain
                    Set oZVTChildLevel3 = oDom.createElement("ZVT.17.5")
                    oZVTChildLevel3.Text = NullToNothing(VitalSignsModRS.Fields("Domain").Value)
                    oZVTChildLevel2.appendChild oZVTChildLevel3
                    
                    ' Modifier Value - Table VitalSignsModifier:ModifierValue
                    Set oZVTChildLevel3 = oDom.createElement("ZVT.17.6")
                    oZVTChildLevel3.Text = NullToNothing(VitalSignsModRS.Fields("ModifierValue").Value)
                    oZVTChildLevel2.appendChild oZVTChildLevel3
                    
                    ' Body Side - Table VitalSignsModifier:BodySide
                    Set oZVTChildLevel3 = oDom.createElement("ZVT.17.7")
                    oZVTChildLevel3.Text = NullToNothing(VitalSignsModRS.Fields("BodySide").Value)
                    oZVTChildLevel2.appendChild oZVTChildLevel3
                    
                    ' Agg Obs - Table VitalSignsModifier:AggObs
                    Set oZVTChildLevel3 = oDom.createElement("ZVT.17.8")
                    oZVTChildLevel3.Text = NullToNothing(VitalSignsModRS.Fields("AggObs").Value)
                    oZVTChildLevel2.appendChild oZVTChildLevel3
                        
                    Set oZVTChildLevel3 = oDom.createElement("ZVT.17.9")
                    oZVTChildLevel3.Text = NullToNothing(VitalSignsModRS.Fields("CREATEDBY").Value)
                    oZVTChildLevel2.appendChild oZVTChildLevel3
                    
                    Set oZVTChildLevel3 = oDom.createElement("ZVT.17.10")
                    strTemp = NullToNothing(VitalSignsModRS.Fields("CREATEDON").Value)
                    If IsValidDate(strTemp) Then
                        oZVTChildLevel3.Text = Format(strTemp, "YYYYMMDDHHNNSS")
                    End If
                    oZVTChildLevel2.appendChild oZVTChildLevel3
                    
                    Set oZVTChildLevel3 = oDom.createElement("ZVT.17.11")
                    oZVTChildLevel3.Text = NullToNothing(VitalSignsModRS.Fields("UPDATEDBY").Value)
                    oZVTChildLevel2.appendChild oZVTChildLevel3
                    
                    Set oZVTChildLevel3 = oDom.createElement("ZVT.17.12")
                    strTemp = NullToNothing(VitalSignsModRS.Fields("UPDATEDON").Value)
                    If IsValidDate(strTemp) Then
                        oZVTChildLevel3.Text = Format(strTemp, "YYYYMMDDHHNNSS")
                    End If
                    oZVTChildLevel2.appendChild oZVTChildLevel3
                        
                oZVTChildLevel1.appendChild oZVTChildLevel2
                VitalSignsModRS.MoveNext
            Loop
        End If
    oZVTNode.appendChild oZVTChildLevel1
        
    Set ReturnNewZVTNode = oZVTNode
    
Exit Function

ErrTrap:
   If gobjShared Is Nothing Then
      Set gobjShared = New CWShared
   End If
   gobjShared.ShowVBError Err.Number, Err.Description, "ReturnNewZVTNode Function", "DMBE", vbExclamation
   Resume Next 'There is no reason this procedure should not continue to create xml if it errors
    
End Function






Public Function ReturnNewZRDNode(ByVal lngCount As Long, ByVal ReadinessRS As ADODB.Recordset) As MSXML.IXMLDOMElement
' Readiness Segment
' sample function output
'   <ZRD>
'      <ZRD.1>Count</ZRD.1>
'      <ZRD.2>Last Readiness Date/Time</ZRD.2>
'      <ZRD.3>HIV Test Date/Time</ZRD.3>
'      <ZRD.4>HIV Test Result</ZRD.4>
'      <ZRD.5>G6PD Test Date/Time</ZRD.5>
'      <ZRD.6>G6PD Test Result</ZRD.6>
'      <ZRD.7>Sickle Cell Test Date/Time</ZRD.7>
'      <ZRD.8>Sickle Cell Test Result</ZRD.8>
'      <ZRD.9>Blood Type Test Date/Time</ZRD.9>
'      <ZRD.10>Blood Type Test Result</ZRD.10>
'      <ZRD.11>DNA Test Date/Time</ZRD.11>
'      <ZRD.12>Hearing Exam Date/Time</ZRD.12>
'      <ZRD.13>Hearing Aid Required</ZRD.13>
'      <ZRD.14>Hearing Aid Issued</ZRD.14>
'      <ZRD.15>Hearing Aid Issued Date/Time</ZRD.15>
'      <ZRD.16>Last Dental Exam Date/Time</ZRD.16>
'      <ZRD.17>Dental Readiness Code</ZRD.17>
'      <ZRD.18>Dental Readiness Date/Time</ZRD.18>
'      <ZRD.19>Medic Assessment Date/Time</ZRD.19>
'      <ZRD.20>Temporary Profile</ZRD.20>
'      <ZRD.21>Temporary Profile Start Date/Time</ZRD.21>
'      <ZRD.22>Temporary Profile End Date/Time</ZRD.22>
'      <ZRD.23>
'         <CE.1>Temporary Profile Diagnosis Identifier</CE.1>
'         <CE.2>Temporary Profile Diagnosis Text</CE.2>
'      </ZRD.23>
'      <ZRD.24>Temporary Profile Limitation</ZRD.24>
'      <ZRD.25>Last Visual Acuity Date/Time</ZRD.25>
'      <ZRD.26>Uncorrected OD</ZRD.26>
'      <ZRD.27>Corrected OD</ZRD.27>
'      <ZRD.28>Uncorrected OS</ZRD.28>
'      <ZRD.29>Corrected OS</ZRD.29>
'      <ZRD.30>Uncorrected OU</ZRD.30>
'      <ZRD.31>Corrected OU</ZRD.31>
'      <ZRD.32>Protect Insert Required</ZRD.32>
'      <ZRD.33>Protect Insert Issued</ZRD.33>
'      <ZRD.34>Protect Insert Exam Date/Time</ZRD.34>
'      <ZRD.35>DNA Onfile</ZRD.35>
'   </ZRD>

    Dim oDom As MSXML.DOMDocument
    Dim oZRDNode As MSXML.IXMLDOMElement
    Dim oZRDChildLevel1 As MSXML.IXMLDOMElement
    Dim oZRDChildLevel2 As MSXML.IXMLDOMElement
    Dim strTemp As String
    
    On Error GoTo ErrTrap

    Set oDom = New MSXML.DOMDocument
    Set oZRDNode = oDom.createElement("ZRD")
    
    ' Set ID - Count
    Set oZRDChildLevel1 = oDom.createElement("ZRD.1")
    oZRDChildLevel1.Text = CStr(lngCount)
    oZRDNode.appendChild oZRDChildLevel1

    ' Last Readiness Date/Time - Table Readiness:Last_Readiness_DT
    Set oZRDChildLevel1 = oDom.createElement("ZRD.2")
    strTemp = NullToNothing(ReadinessRS.Fields("Last_Readiness_DT").Value)
    If IsValidDate(strTemp) Then
        oZRDChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZRDNode.appendChild oZRDChildLevel1

    ' HIV Test Date/Time - Table Readiness:HIV_Test_DT
    Set oZRDChildLevel1 = oDom.createElement("ZRD.3")
    strTemp = NullToNothing(ReadinessRS.Fields("HIV_Test_DT").Value)
    If IsValidDate(strTemp) Then
        oZRDChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZRDNode.appendChild oZRDChildLevel1

    ' HIV Test Result - Table Readiness:HIV_Test_Rslt
    Set oZRDChildLevel1 = oDom.createElement("ZRD.4")
    oZRDChildLevel1.Text = NullToNothing(ReadinessRS.Fields("HIV_Test_Rslt").Value)
    oZRDNode.appendChild oZRDChildLevel1

    ' G6PD Test Date/Time - Table Readiness:G6PD_Test_DT
    Set oZRDChildLevel1 = oDom.createElement("ZRD.5")
    strTemp = NullToNothing(ReadinessRS.Fields("G6PD_Test_DT").Value)
    If IsValidDate(strTemp) Then
        oZRDChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZRDNode.appendChild oZRDChildLevel1

    ' G6PD Test Result - Table Readiness:G6PD_Test_Rslt
    Set oZRDChildLevel1 = oDom.createElement("ZRD.6")
    oZRDChildLevel1.Text = NullToNothing(ReadinessRS.Fields("G6PD_Test_Rslt").Value)
    oZRDNode.appendChild oZRDChildLevel1

    ' Sickle Cell Test Date/Time - Table Readiness:Sickle_Cell_Test_DT
    Set oZRDChildLevel1 = oDom.createElement("ZRD.7")
    strTemp = NullToNothing(ReadinessRS.Fields("Sickle_Cell_Test_DT").Value)
    If IsValidDate(strTemp) Then
        oZRDChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZRDNode.appendChild oZRDChildLevel1

    ' Sickle Cell Test Result - Table Readiness:Sickle_Cell_Test_Rslt
    Set oZRDChildLevel1 = oDom.createElement("ZRD.8")
    oZRDChildLevel1.Text = NullToNothing(ReadinessRS.Fields("Sickle_Cell_Test_Rslt").Value)
    oZRDNode.appendChild oZRDChildLevel1

    ' Blood Type Test Date/Time - Table Readiness:Blood_Type_Test_DT
    Set oZRDChildLevel1 = oDom.createElement("ZRD.9")
    strTemp = NullToNothing(ReadinessRS.Fields("Blood_Type_Test_DT").Value)
    If IsValidDate(strTemp) Then
        oZRDChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZRDNode.appendChild oZRDChildLevel1

    ' Blood Type Test Result - Table Readiness:Blood_Type_Test_Rslt
    Set oZRDChildLevel1 = oDom.createElement("ZRD.10")
    oZRDChildLevel1.Text = NullToNothing(ReadinessRS.Fields("Blood_Type_Test_Rslt").Value)
    oZRDNode.appendChild oZRDChildLevel1

    ' DNA Test Date/Time - Table Readiness:DNA_Test_DT
    Set oZRDChildLevel1 = oDom.createElement("ZRD.11")
    strTemp = NullToNothing(ReadinessRS.Fields("DNA_Test_DT").Value)
    If IsValidDate(strTemp) Then
        oZRDChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZRDNode.appendChild oZRDChildLevel1

    ' Hearing Exam Date/Time - Table Readiness:Hear_Exam_DT
    Set oZRDChildLevel1 = oDom.createElement("ZRD.12")
    strTemp = NullToNothing(ReadinessRS.Fields("Hear_Exam_DT").Value)
    If IsValidDate(strTemp) Then
        oZRDChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZRDNode.appendChild oZRDChildLevel1

    ' Hearing Aid Required - Table Readiness:Hear_Aid_Required
    Set oZRDChildLevel1 = oDom.createElement("ZRD.13")
    oZRDChildLevel1.Text = NullToNothing(ReadinessRS.Fields("Hear_Aid_Required").Value)
    oZRDNode.appendChild oZRDChildLevel1

    ' Hearing Aid Issued - Table Readiness:Hear_Aid_Issued
    Set oZRDChildLevel1 = oDom.createElement("ZRD.14")
    oZRDChildLevel1.Text = NullToNothing(ReadinessRS.Fields("Hear_Aid_Issued").Value)
    oZRDNode.appendChild oZRDChildLevel1

    ' Hearing Aid Issued Date/Time - Table Readiness:Hear_Aid_Issued_DT
    Set oZRDChildLevel1 = oDom.createElement("ZRD.15")
    strTemp = NullToNothing(ReadinessRS.Fields("Hear_Aid_Issued_DT").Value)
    If IsValidDate(strTemp) Then
        oZRDChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZRDNode.appendChild oZRDChildLevel1

    ' Last Dental Exam Date/Time - Table Readiness:Last_Dental_Exam_DT
    Set oZRDChildLevel1 = oDom.createElement("ZRD.16")
    strTemp = NullToNothing(ReadinessRS.Fields("Last_Dental_Exam_DT").Value)
    If IsValidDate(strTemp) Then
        oZRDChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZRDNode.appendChild oZRDChildLevel1
    
    ' Dental Readiness Code - Table Readiness:Dental_Readiness_CD
    Set oZRDChildLevel1 = oDom.createElement("ZRD.17")
    oZRDChildLevel1.Text = NullToNothing(ReadinessRS.Fields("Dental_Readiness_CD").Value)
    oZRDNode.appendChild oZRDChildLevel1

    ' Dental Readiness Date/Time - Table Readiness:Dental_Readiness_DT
    Set oZRDChildLevel1 = oDom.createElement("ZRD.18")
    strTemp = NullToNothing(ReadinessRS.Fields("Dental_Readiness_DT").Value)
    If IsValidDate(strTemp) Then
        oZRDChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZRDNode.appendChild oZRDChildLevel1
    
    ' Medic Assessment Date/Time - Table Readiness:Medic_Assessment_DT
    Set oZRDChildLevel1 = oDom.createElement("ZRD.19")
    strTemp = NullToNothing(ReadinessRS.Fields("Medic_Assessment_DT").Value)
    If IsValidDate(strTemp) Then
        oZRDChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZRDNode.appendChild oZRDChildLevel1
    
    ' Temporary Profile - Table Readiness:Temp_Profile
    Set oZRDChildLevel1 = oDom.createElement("ZRD.20")
    oZRDChildLevel1.Text = NullToNothing(ReadinessRS.Fields("Temp_Profile").Value)
    oZRDNode.appendChild oZRDChildLevel1

    ' Temporary Profile Start Date/Time - Table Readiness:Temp_Profile_Start_DT
    Set oZRDChildLevel1 = oDom.createElement("ZRD.21")
    strTemp = NullToNothing(ReadinessRS.Fields("Temp_Profile_Start_DT").Value)
    If IsValidDate(strTemp) Then
        oZRDChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZRDNode.appendChild oZRDChildLevel1
    
    ' Temporary Profile End Date/Time - Table Readiness:Temp_Profile_End_DT
    Set oZRDChildLevel1 = oDom.createElement("ZRD.22")
    strTemp = NullToNothing(ReadinessRS.Fields("Temp_Profile_End_DT").Value)
    If IsValidDate(strTemp) Then
        oZRDChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZRDNode.appendChild oZRDChildLevel1
    
    ' Temporary Profile Diagnosis
    Set oZRDChildLevel1 = oDom.createElement("ZRD.23")
        ' Identifier - Table Readiness:Temp_Profile_Dignos_CD
        Set oZRDChildLevel2 = oDom.createElement("CE.1")
        oZRDChildLevel2.Text = NullToNothing(ReadinessRS.Fields("Temp_Profile_Dignos_CD").Value)
        oZRDChildLevel1.appendChild oZRDChildLevel2
        ' Text - Table Readiness:Temp_Profile_Dignos_Dsc
        Set oZRDChildLevel2 = oDom.createElement("CE.2")
        oZRDChildLevel2.Text = NullToNothing(ReadinessRS.Fields("Temp_Profile_Dignos_Dsc").Value)
        oZRDChildLevel1.appendChild oZRDChildLevel2
    oZRDNode.appendChild oZRDChildLevel1

    ' Temporary Profile Limitation - Table Readiness:Temp_Profile_Limitation
    Set oZRDChildLevel1 = oDom.createElement("ZRD.24")
    oZRDChildLevel1.Text = NullToNothing(ReadinessRS.Fields("Temp_Profile_Limitation").Value)
    oZRDNode.appendChild oZRDChildLevel1

    ' Last Visual Acuity Date/Time - Table Readiness:Last_Visual_Acuity_DT
    Set oZRDChildLevel1 = oDom.createElement("ZRD.25")
    strTemp = NullToNothing(ReadinessRS.Fields("Last_Visual_Acuity_DT").Value)
    If IsValidDate(strTemp) Then
        oZRDChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZRDNode.appendChild oZRDChildLevel1
    
    ' Uncorrected OD - Table Readiness:Uncorrected_OD
    Set oZRDChildLevel1 = oDom.createElement("ZRD.26")
    oZRDChildLevel1.Text = NullToNothing(ReadinessRS.Fields("Uncorrected_OD").Value)
    oZRDNode.appendChild oZRDChildLevel1

    ' Corrected OD - Table Readiness:Corrected_OD
    Set oZRDChildLevel1 = oDom.createElement("ZRD.27")
    oZRDChildLevel1.Text = NullToNothing(ReadinessRS.Fields("Corrected_OD").Value)
    oZRDNode.appendChild oZRDChildLevel1

    ' Uncorrected OS - Table Readiness:Uncorrected_OS
    Set oZRDChildLevel1 = oDom.createElement("ZRD.28")
    oZRDChildLevel1.Text = NullToNothing(ReadinessRS.Fields("Uncorrected_OS").Value)
    oZRDNode.appendChild oZRDChildLevel1

    ' Corrected OS - Table Readiness:Corrected_OS
    Set oZRDChildLevel1 = oDom.createElement("ZRD.29")
    oZRDChildLevel1.Text = NullToNothing(ReadinessRS.Fields("Corrected_OS").Value)
    oZRDNode.appendChild oZRDChildLevel1

    ' Uncorrected OS - Table Readiness:Uncorrected_OU
    Set oZRDChildLevel1 = oDom.createElement("ZRD.30")
    oZRDChildLevel1.Text = NullToNothing(ReadinessRS.Fields("Uncorrected_OU").Value)
    oZRDNode.appendChild oZRDChildLevel1

    ' Corrected OS - Table Readiness:Corrected_OS
    Set oZRDChildLevel1 = oDom.createElement("ZRD.31")
    oZRDChildLevel1.Text = NullToNothing(ReadinessRS.Fields("Corrected_OU").Value)
    oZRDNode.appendChild oZRDChildLevel1

    ' Protect Insert Required - Table Readiness:Protect_Insert_Required
    Set oZRDChildLevel1 = oDom.createElement("ZRD.32")
    oZRDChildLevel1.Text = NullToNothing(ReadinessRS.Fields("Protect_Insert_Required").Value)
    oZRDNode.appendChild oZRDChildLevel1

    ' Protect Insert Issued - Table Readiness:Protect_Insert_Issued
    Set oZRDChildLevel1 = oDom.createElement("ZRD.33")
    oZRDChildLevel1.Text = NullToNothing(ReadinessRS.Fields("Protect_Insert_Issued").Value)
    oZRDNode.appendChild oZRDChildLevel1

    ' Protect Insert Exam Date/Time - Table Readiness:Protect_Insert_Exam_DT
    Set oZRDChildLevel1 = oDom.createElement("ZRD.34")
    strTemp = NullToNothing(ReadinessRS.Fields("Protect_Insert_Exam_DT").Value)
    If IsValidDate(strTemp) Then
        oZRDChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZRDNode.appendChild oZRDChildLevel1
    
    ' DNA Onfile - Table Readiness:DNA_On_File
    Set oZRDChildLevel1 = oDom.createElement("ZRD.35")
    oZRDChildLevel1.Text = NullToNothing(ReadinessRS.Fields("DNA_On_File").Value)
    oZRDNode.appendChild oZRDChildLevel1

    Set oZRDChildLevel1 = oDom.createElement("ZRD.36")
    oZRDChildLevel1.Text = NullToNothing(ReadinessRS.Fields("PATIENT_UNIT_NUMBER").Value)
    oZRDNode.appendChild oZRDChildLevel1

    Set oZRDChildLevel1 = oDom.createElement("ZRD.37")
    oZRDChildLevel1.Text = NullToNothing(ReadinessRS.Fields("ENCOUNTER_NUMBER").Value)
    oZRDNode.appendChild oZRDChildLevel1
    
    Set oZRDChildLevel1 = oDom.createElement("ZRD.38")
    oZRDChildLevel1.Text = NullToNothing(ReadinessRS.Fields("CREATEDBY").Value)
    oZRDNode.appendChild oZRDChildLevel1
    
    Set oZRDChildLevel1 = oDom.createElement("ZRD.39")
    strTemp = NullToNothing(ReadinessRS.Fields("CREATEDON").Value)
    If IsValidDate(strTemp) Then
        oZRDChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZRDNode.appendChild oZRDChildLevel1
    
    Set oZRDChildLevel1 = oDom.createElement("ZRD.40")
    oZRDChildLevel1.Text = NullToNothing(ReadinessRS.Fields("UPDATEDBY").Value)
    oZRDNode.appendChild oZRDChildLevel1
    
    Set oZRDChildLevel1 = oDom.createElement("ZRD.41")
    strTemp = NullToNothing(ReadinessRS.Fields("UPDATEDON").Value)
    If IsValidDate(strTemp) Then
        oZRDChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZRDNode.appendChild oZRDChildLevel1
    
    Set ReturnNewZRDNode = oZRDNode

Exit Function

ErrTrap:
   If gobjShared Is Nothing Then
      Set gobjShared = New CWShared
   End If
   gobjShared.ShowVBError Err.Number, Err.Description, "ReturnNewPV1Node Function", "DMBE", vbExclamation
   Resume Next 'There is no reason this procedure should not continue to create xml if it errors
        
End Function

Public Function ReturnNewPV1Node(ByVal lngCount As Long, ByVal lngProviderNCID As Long, ByVal lngFacilityNCID As Long, ByVal lngEncType As Long, ByVal varAdmitDate As Variant, ByVal varDischargeDate As Variant, ByVal varStatus As Variant) As MSXML.IXMLDOMElement

' Patient Visit Segment
' sample function output
'   <PV1>
'      <PV1.1>1</PV1.1>
'      <PV1.2>O</PV1.2>
'      <PV1.3>
'         <PL.1>Midwifery</PL.1>
'         <PL.4>
'            <HD.1>ADOLESCENT MEDICINE CLINIC</HD.1>
'         </PL.4>
'      </PV1.3>
'      <PV1.7.LST>
'         <PV1.7>
'            <XCN.1>9987656554</XCN.1>
'            <XCN.2>SHEN</XCN.2>
'            <XCN.3>JANE</XCN.3>
'            <XCN.4></XCN.4>
'         </PV1.7>
'      </PV1.7.LST>
'      <PV1.18>1</PV1.18>
'      <PV1.36>00</PV1.36>
'      <PV1.39>
'            <CE.1></CE.1>
'            <CE.2>ADOLESCENT MEDICINE CLINIC</CE.2>
'            <CE.3></CE.3>
'            <CE.4>1047017</CE.4>
'            <CE.5>ADOLESCENT MEDICINE CLINIC</CE.5>
'            <CE.6>99NCI</CE.6>
'      </PV1.39>
'      <PV1.44>20010808180345</PV1.44>
'      <PV1.45></PV1.45>
'   </PV1>

    Dim oDom As MSXML.DOMDocument
    Dim oPV1Node As MSXML.IXMLDOMElement
    Dim oPV1ChildLevel1 As MSXML.IXMLDOMElement
    Dim oPV1ChildLevel2 As MSXML.IXMLDOMElement
    Dim oPV1ChildLevel3 As MSXML.IXMLDOMElement
    Dim strDocFirstName As String
    Dim strDocLastName As String
    Dim strDocMiddleName As String
    Dim strDocID As String
    Dim strDocSpecialty As String
    
    On Error GoTo ErrTrap

    Set oDom = New MSXML.DOMDocument
    Set oPV1Node = oDom.createElement("PV1")
    
    ' Segment Count
    Set oPV1ChildLevel1 = oDom.createElement("PV1.1")
    oPV1ChildLevel1.Text = CStr(lngCount)
    oPV1Node.appendChild oPV1ChildLevel1

    ' Patient Class - default to "O" - Outpatient
    Set oPV1ChildLevel1 = oDom.createElement("PV1.2")
    oPV1ChildLevel1.Text = "O"
    oPV1Node.appendChild oPV1ChildLevel1

    ' Assigned patient Location
    Set oPV1ChildLevel1 = oDom.createElement("PV1.3")
        
        Call GetNewDoctorInfo(lngProviderNCID, strDocLastName, strDocFirstName, strDocMiddleName, strDocID, strDocSpecialty)
        ' Medical Specialty
        Set oPV1ChildLevel2 = oDom.createElement("PL.1")
        oPV1ChildLevel2.Text = strDocSpecialty
        oPV1ChildLevel1.appendChild oPV1ChildLevel2
        ' Medical Treatment Facility
        Set oPV1ChildLevel2 = oDom.createElement("PL.4")
            Set oPV1ChildLevel3 = oDom.createElement("HD.1")
            'oPV1ChildLevel3.Text = GetNCIDDescription(lngFacilityNCID, True) 'SCR 25083 CC
            oPV1ChildLevel2.appendChild oPV1ChildLevel3
        oPV1ChildLevel1.appendChild oPV1ChildLevel2
    oPV1Node.appendChild oPV1ChildLevel1
    '<SCR 22597 CC>
    
    ' Attending Doctor
    Set oPV1ChildLevel1 = oDom.createElement("PV1.7.LST")
        Set oPV1ChildLevel2 = oDom.createElement("PV1.7")
            ' Attending Doctor ID (SSN)
            Set oPV1ChildLevel3 = oDom.createElement("XCN.1")
            oPV1ChildLevel3.Text = Trim(strDocID)
            oPV1ChildLevel2.appendChild oPV1ChildLevel3
            ' Attending Doctor Last Name
            Set oPV1ChildLevel3 = oDom.createElement("XCN.2")
            oPV1ChildLevel3.Text = Trim(strDocLastName)
            oPV1ChildLevel2.appendChild oPV1ChildLevel3
            ' Attending Doctor First Name
            Set oPV1ChildLevel3 = oDom.createElement("XCN.3")
            oPV1ChildLevel3.Text = Trim(strDocFirstName)
            oPV1ChildLevel2.appendChild oPV1ChildLevel3
            ' Attending Doctor Middle Name
            Set oPV1ChildLevel3 = oDom.createElement("XCN.4")
            oPV1ChildLevel3.Text = Trim(strDocMiddleName)
            oPV1ChildLevel2.appendChild oPV1ChildLevel3
        oPV1ChildLevel1.appendChild oPV1ChildLevel2
    oPV1Node.appendChild oPV1ChildLevel1
    
    '</SCR 22597 CC>
    
    ' Patient Type - Table Encounters:Type
    Set oPV1ChildLevel1 = oDom.createElement("PV1.18")
    oPV1ChildLevel1.Text = CStr(lngEncType)
    oPV1Node.appendChild oPV1ChildLevel1
    
    ' Discharge Disposition - Table Enc_Disposition:DispositionNCID
    Set oPV1ChildLevel1 = oDom.createElement("PV1.36")
    
'        Set oPV1ChildLevel2 = oDom.createElement("CE.1")
'        oPV1ChildLevel2.Text = ""
'        oPV1ChildLevel1.appendChild oPV1ChildLevel2
'        ' Disposition Code
'        Set oPV1ChildLevel2 = oDom.createElement("CE.2")
'        oPV1ChildLevel2.Text = DischargeDispositionText(varStatus)
'        oPV1ChildLevel1.appendChild oPV1ChildLevel2
'
'        Set oPV1ChildLevel2 = oDom.createElement("CE.3")
'        oPV1ChildLevel2.Text = ""
'        oPV1ChildLevel1.appendChild oPV1ChildLevel2
'
'        Set oPV1ChildLevel2 = oDom.createElement("CE.4")
'        oPV1ChildLevel2.Text = lngFacilityNCID
'        oPV1ChildLevel1.appendChild oPV1ChildLevel2
'        ' Disposition NCID
'        Set oPV1ChildLevel2 = oDom.createElement("CE.5")
'        oPV1ChildLevel2.Text = varStatus 'This is the NCID
'        oPV1ChildLevel1.appendChild oPV1ChildLevel2
'
'        Set oPV1ChildLevel2 = oDom.createElement("CE.6")
'        oPV1ChildLevel2.Text = "99NCI"
'        oPV1ChildLevel1.appendChild oPV1ChildLevel2
    
    
    oPV1ChildLevel1.Text = DischargeDispositionText(varStatus)
    oPV1Node.appendChild oPV1ChildLevel1
    
    ' Servicing Facility
    Set oPV1ChildLevel1 = oDom.createElement("PV1.39")
        ' Serving Facility ICD-9
        Set oPV1ChildLevel2 = oDom.createElement("CE.1")
        oPV1ChildLevel2.Text = ""
        oPV1ChildLevel1.appendChild oPV1ChildLevel2
        ' Clinic Name
        Set oPV1ChildLevel2 = oDom.createElement("CE.2")
        oPV1ChildLevel2.Text = GetNCIDDescription(lngFacilityNCID)
        oPV1ChildLevel1.appendChild oPV1ChildLevel2
             
        Set oPV1ChildLevel2 = oDom.createElement("CE.3")
        oPV1ChildLevel2.Text = ""
        oPV1ChildLevel1.appendChild oPV1ChildLevel2
        ' Serving Facility ICD-9
        Set oPV1ChildLevel2 = oDom.createElement("CE.4")
        oPV1ChildLevel2.Text = lngFacilityNCID
        oPV1ChildLevel1.appendChild oPV1ChildLevel2
        ' Clinic Name
        Set oPV1ChildLevel2 = oDom.createElement("CE.5")
        oPV1ChildLevel2.Text = GetNCIDDescription(lngFacilityNCID)
        oPV1ChildLevel1.appendChild oPV1ChildLevel2
        
        Set oPV1ChildLevel2 = oDom.createElement("CE.6")
        oPV1ChildLevel2.Text = "99NCI"
        oPV1ChildLevel1.appendChild oPV1ChildLevel2
    
    oPV1Node.appendChild oPV1ChildLevel1
    
    ' Admit Date/Time
    Set oPV1ChildLevel1 = oDom.createElement("PV1.44")
    If IsValidDate(varAdmitDate) Then
        oPV1ChildLevel1.Text = Format(varAdmitDate, "YYYYMMDDHHNNSS")
    End If
    oPV1Node.appendChild oPV1ChildLevel1
    
    ' Discharge Date/Time
    Set oPV1ChildLevel1 = oDom.createElement("PV1.45")
    If IsValidDate(varDischargeDate) Then
        oPV1ChildLevel1.Text = Format(varDischargeDate, "YYYYMMDDHHNNSS")
    End If
    oPV1Node.appendChild oPV1ChildLevel1
    
    Set ReturnNewPV1Node = oPV1Node

Exit Function

ErrTrap:
   If gobjShared Is Nothing Then
      Set gobjShared = New CWShared
   End If
   gobjShared.ShowVBError Err.Number, Err.Description, "ReturnNewPV1Node Function", "DMBE", vbExclamation
   'Err.Clear
   Resume Next 'There is no reason this procedure should not continue to create xml if it errors
        
End Function

Public Function ReturnNewZRENode(ByVal lngCount As Long, ByVal MedicationsRS As ADODB.Recordset, ByVal RXListRS As ADODB.Recordset) As MSXML.IXMLDOMElement

' Additional Pharmacy Encoded Order Data Segment
' sample function output
'   <ZRE>
'      <ZRE.1>Count</ZRE.1>
'      <ZRE.2>Unit Number</ZRE.2>
'      <ZRE.3>Medications Comment</ZRE.3>
'      <ZRE.4>Dispense Location</ZRE.4>
'      <ZRE.5>Days of Supply</ZRE.5>
'      <ZRE.6>Data ID</ZRE.6>
'      <ZRE.7>Quantity</ZRE.7>
'      <ZRE.8>OTC</ZRE.8>
'      <ZRE.9>Order Expiration Data/Time</ZRE.9>
'      <ZRE.10>Ordering Provider</ZRE.10>
'      <ZRE.11>Clinic</ZRE.11>
'      <ZRE.12>Patient Type</ZRE.12>
'      <ZRE.13>Update Flag</ZRE.13>
'      <ZRE.14>CDR Data ID</ZRE.14>
'      <ZRE.15>Storing App</ZRE.15>
'      <ZRE.16>Created By</ZRE.16>
'      <ZRE.17>Created On</ZRE.17>
'      <ZRE.18>Updated By</ZRE.18>
'      <ZRE.19>Updated On</ZRE.19>
'      <ZRE.20>Ordering Provider NCID</ZRE.20>
'      <ZRE.21>Status NCID</ZRE.21>
'      <ZRE.22>National Drug Code</ZRE.22>
'      <ZRE.23>Non Coded Medication Name</ZRE.23>
'      <ZRE.24>TMIP Status</ZRE.24>
'      <ZRE.25>Encounter Number</ZRE.25>

'      <ZRE.26>Facility NCID</ZRE.26>
'      <ZRE.27>Drug Name</ZRE.27>
'      <ZRE.28>Drug IEN</ZRE.28>
'      <ZRE.29>Default Unit</ZRE.29>
'      <ZRE.30>NDC Indicator</ZRE.30>
'      <ZRE.31>Drug Check Disabled</ZRE.31>
'      <ZRE.32>Dosage Strength</ZRE.32>
'      <ZRE.33>Content Unit</ZRE.33>
'      <ZRE.34>Dosage Form</ZRE.34>
'      <ZRE.35>Synonyms</ZRE.35>
'      <ZRE.36>Default SIG</ZRE.36>
'      <ZRE.37>Default Refill</ZRE.37>
'      <ZRE.38>Default Route</ZRE.38>
'      <ZRE.39>Current Stock</ZRE.39>
'   </ZRE>

    Dim oDom As MSXML.DOMDocument
    Dim oZRENode As MSXML.IXMLDOMElement
    Dim oZREChildLevel1 As MSXML.IXMLDOMElement
    Dim strTemp As String
    
    On Error GoTo ErrTrap

    Set oDom = New MSXML.DOMDocument
    Set oZRENode = oDom.createElement("ZRE")
    
    ' Segment Count
    Set oZREChildLevel1 = oDom.createElement("ZRE.1")
    oZREChildLevel1.Text = CStr(lngCount)
    oZRENode.appendChild oZREChildLevel1
    
    ' Unit Number - Table Medications:Unit_Number
    Set oZREChildLevel1 = oDom.createElement("ZRE.2")
    oZREChildLevel1.Text = NullToNothing(MedicationsRS.Fields("Unit_Number").Value)
    oZRENode.appendChild oZREChildLevel1
    
    ' Medications Comment - Table Medications:MedicationsComment
    Set oZREChildLevel1 = oDom.createElement("ZRE.3")
    oZREChildLevel1.Text = NullToNothing(MedicationsRS.Fields("MedicationsComment").Value)
    oZRENode.appendChild oZREChildLevel1
    
    ' Dispense Location NCID - Table Medications:DispensingLocationNCID
    Set oZREChildLevel1 = oDom.createElement("ZRE.4")
    oZREChildLevel1.Text = NullToNothing(MedicationsRS.Fields("DispensingLocationNCID").Value)
    oZRENode.appendChild oZREChildLevel1
    
    ' Days of Supply - Table Medications:DaysOfSupply
    Set oZREChildLevel1 = oDom.createElement("ZRE.5")
    oZREChildLevel1.Text = NullToNothing(MedicationsRS.Fields("DaysOfSupply").Value)
    oZRENode.appendChild oZREChildLevel1
    
    ' Data ID - Table Medications:DataID
    Set oZREChildLevel1 = oDom.createElement("ZRE.6")
    oZREChildLevel1.Text = NullToNothing(MedicationsRS.Fields("DataID").Value)
    oZRENode.appendChild oZREChildLevel1
    
    ' Quantity - Table Medications:Quantity
    Set oZREChildLevel1 = oDom.createElement("ZRE.7")
    oZREChildLevel1.Text = NullToNothing(MedicationsRS.Fields("Quantity").Value)
    oZRENode.appendChild oZREChildLevel1
    
    ' OTC - Table Medications:OTC
    Set oZREChildLevel1 = oDom.createElement("ZRE.8")
    oZREChildLevel1.Text = NullToNothing(MedicationsRS.Fields("OTC").Value)
    oZRENode.appendChild oZREChildLevel1
    
    ' Order Expiration Date/Time - Table Medications:OrderExpirationDate
    Set oZREChildLevel1 = oDom.createElement("ZRE.9")
    strTemp = NullToNothing(MedicationsRS.Fields("OrderExpirationDate").Value)
    If IsValidDate(strTemp) Then
        oZREChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZRENode.appendChild oZREChildLevel1
    
    ' Ordering Provider - Table Medications:OrderingProvider
    Set oZREChildLevel1 = oDom.createElement("ZRE.10")
    oZREChildLevel1.Text = NullToNothing(MedicationsRS.Fields("OrderingProvider").Value)
    oZRENode.appendChild oZREChildLevel1
    
    ' Clinic - Table Medications:ClinicNCID
    Set oZREChildLevel1 = oDom.createElement("ZRE.11")
    oZREChildLevel1.Text = NullToNothing(MedicationsRS.Fields("ClinicNCID").Value)
    oZRENode.appendChild oZREChildLevel1
    
    ' Patient Type - Table Medications:PatientType
    Set oZREChildLevel1 = oDom.createElement("ZRE.12")
    oZREChildLevel1.Text = NullToNothing(MedicationsRS.Fields("PatientType").Value)
    oZRENode.appendChild oZREChildLevel1
    
    ' Update Flag - Table Medications:UpdateFlag
    Set oZREChildLevel1 = oDom.createElement("ZRE.13")
    oZREChildLevel1.Text = NullToNothing(MedicationsRS.Fields("UpdateFlag").Value)
    oZRENode.appendChild oZREChildLevel1
    
    ' CDR Data ID - Table Medications:CDR_DataID
    Set oZREChildLevel1 = oDom.createElement("ZRE.14")
    oZREChildLevel1.Text = NullToNothing(MedicationsRS.Fields("CDR_DataID").Value)
    oZRENode.appendChild oZREChildLevel1
    
    ' Storing App - Table Medications:StoringApp
    Set oZREChildLevel1 = oDom.createElement("ZRE.15")
    oZREChildLevel1.Text = NullToNothing(MedicationsRS.Fields("StoringApp").Value)
    oZRENode.appendChild oZREChildLevel1
    
    ' Created By - Table Medications:CreatedBy
    Set oZREChildLevel1 = oDom.createElement("ZRE.16")
    oZREChildLevel1.Text = NullToNothing(MedicationsRS.Fields("CreatedBy").Value)
    oZRENode.appendChild oZREChildLevel1
    
    ' Created On - Table Medications:CreatedOn
    Set oZREChildLevel1 = oDom.createElement("ZRE.17")
    strTemp = NullToNothing(MedicationsRS.Fields("CreatedOn").Value)
    If IsValidDate(strTemp) Then
        oZREChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZRENode.appendChild oZREChildLevel1
    
    ' Updated By - Table Medications:UpdatedBy
    Set oZREChildLevel1 = oDom.createElement("ZRE.18")
    oZREChildLevel1.Text = NullToNothing(MedicationsRS.Fields("UpdatedBy").Value)
    oZRENode.appendChild oZREChildLevel1
    
    ' Updated On - Table Medications:UpdatedOn
    Set oZREChildLevel1 = oDom.createElement("ZRE.19")
    strTemp = NullToNothing(MedicationsRS.Fields("UpdatedOn").Value)
    If IsValidDate(strTemp) Then
        oZREChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZRENode.appendChild oZREChildLevel1
    
    ' Begin:SCR #47403;   Developer:  12/24/2003
    'Ordering Provider NCID
    Set oZREChildLevel1 = oDom.createElement("ZRE.20")
    oZREChildLevel1.Text = NullToNothing(MedicationsRS.Fields("ORDERINGPROVIDERNCID").Value)
    oZRENode.appendChild oZREChildLevel1
    
    'Status NCID
    Set oZREChildLevel1 = oDom.createElement("ZRE.21")
    oZREChildLevel1.Text = NullToNothing(MedicationsRS.Fields("STATUSNCID").Value)
    oZRENode.appendChild oZREChildLevel1
    
    'National Drug Code
    Set oZREChildLevel1 = oDom.createElement("ZRE.22")
    oZREChildLevel1.Text = NullToNothing(MedicationsRS.Fields("NATIONALDRUGCODE").Value)
    oZRENode.appendChild oZREChildLevel1
    
    'Non Coded Medication Name
    Set oZREChildLevel1 = oDom.createElement("ZRE.23")
    oZREChildLevel1.Text = NullToNothing(MedicationsRS.Fields("NON_CODED_MEDICATION_NAME").Value)
    oZRENode.appendChild oZREChildLevel1
    
    'TMIP Status
    Set oZREChildLevel1 = oDom.createElement("ZRE.24")
    oZREChildLevel1.Text = NullToNothing(MedicationsRS.Fields("TMIP_STATUS").Value)
    oZRENode.appendChild oZREChildLevel1
    
    'Encounter Number
    Set oZREChildLevel1 = oDom.createElement("ZRE.25")
    oZREChildLevel1.Text = NullToNothing(MedicationsRS.Fields("ENC_NUM").Value)
    oZRENode.appendChild oZREChildLevel1
    
  'MAG: Tracker #52316 04/14/2004
  If Not RXListRS Is Nothing Then
    If RXListRS.RecordCount <> 0 Then
      'Facility NCID
      Set oZREChildLevel1 = oDom.createElement("ZRE.26")
      oZREChildLevel1.Text = NullToNothing(RXListRS.Fields("FACILITY_NCID").Value)
      oZRENode.appendChild oZREChildLevel1
  
      'Drug Name
      Set oZREChildLevel1 = oDom.createElement("ZRE.27")
      oZREChildLevel1.Text = NullToNothing(RXListRS.Fields("DRUG_NAME").Value)
      oZRENode.appendChild oZREChildLevel1
      
      'Drug IEN
      Set oZREChildLevel1 = oDom.createElement("ZRE.28")
      oZREChildLevel1.Text = NullToNothing(RXListRS.Fields("DRUG_IEN").Value)
      oZRENode.appendChild oZREChildLevel1
      
      'Default Unit
      Set oZREChildLevel1 = oDom.createElement("ZRE.29")
      oZREChildLevel1.Text = NullToNothing(RXListRS.Fields("DEFAULT_UNIT").Value)
      oZRENode.appendChild oZREChildLevel1
      
      'NDC Indicator
      Set oZREChildLevel1 = oDom.createElement("ZRE.30")
      oZREChildLevel1.Text = NullToNothing(RXListRS.Fields("NDC_INDICATOR").Value)
      oZRENode.appendChild oZREChildLevel1
      
      'Drug Check Disabled
      Set oZREChildLevel1 = oDom.createElement("ZRE.31")
      oZREChildLevel1.Text = NullToNothing(RXListRS.Fields("DRUG_CHECK_DISABLED").Value)
      oZRENode.appendChild oZREChildLevel1
      
      'Dosage Strength
      Set oZREChildLevel1 = oDom.createElement("ZRE.32")
      oZREChildLevel1.Text = NullToNothing(RXListRS.Fields("DOSAGE_STRENGTH").Value)
      oZRENode.appendChild oZREChildLevel1
      
      'Content Unit
      Set oZREChildLevel1 = oDom.createElement("ZRE.33")
      oZREChildLevel1.Text = NullToNothing(RXListRS.Fields("CONTENT_UNIT").Value)
      oZRENode.appendChild oZREChildLevel1
      
      'Dosage Form
      Set oZREChildLevel1 = oDom.createElement("ZRE.34")
      oZREChildLevel1.Text = NullToNothing(RXListRS.Fields("DOSAGE_FORM").Value)
      oZRENode.appendChild oZREChildLevel1
      
      'Synonyms
      Set oZREChildLevel1 = oDom.createElement("ZRE.35")
      oZREChildLevel1.Text = NullToNothing(RXListRS.Fields("SYNONYMS").Value)
      oZRENode.appendChild oZREChildLevel1
      
      'Default SIG
      Set oZREChildLevel1 = oDom.createElement("ZRE.36")
      oZREChildLevel1.Text = NullToNothing(RXListRS.Fields("DEFAULT_SIG").Value)
      oZRENode.appendChild oZREChildLevel1
      
      'Default Refill
      Set oZREChildLevel1 = oDom.createElement("ZRE.37")
      oZREChildLevel1.Text = NullToNothing(RXListRS.Fields("DEFAULT_REFILL").Value)
      oZRENode.appendChild oZREChildLevel1
      
      'Default Route
      Set oZREChildLevel1 = oDom.createElement("ZRE.38")
      oZREChildLevel1.Text = NullToNothing(RXListRS.Fields("DEFAULT_ROUTE").Value)
      oZRENode.appendChild oZREChildLevel1
      
      'Current Stock
      Set oZREChildLevel1 = oDom.createElement("ZRE.39")
      oZREChildLevel1.Text = NullToNothing(RXListRS.Fields("CURRENT_STOCK").Value)
      oZRENode.appendChild oZREChildLevel1
      ' End: SCR #47403;
    End If 'RXListRS.RecordCount
  End If
  Set ReturnNewZRENode = oZRENode
    
Exit Function

ErrTrap:
   If gobjShared Is Nothing Then
      Set gobjShared = New CWShared
   End If
   gobjShared.ShowVBError Err.Number, Err.Description, "ReturnNewZRENode Function", "DMBE", vbExclamation
   Resume Next 'There is no reason this procedure should not continue to create xml if it errors
    
End Function

Public Function ReturnNewRXENode(ByVal MedicationsRS As ADODB.Recordset) As MSXML.IXMLDOMElement
' Pharmacy Encoded Order Segment
' sample function output
'   <RXE>
'      <RXE.1>^^^20010719^20010726</RXE.1>
'      <RXE.2>
'         <CE.1></CE.1>
'         <CE.2></CE.2>
'         <CE.3></CE.3>
'         <CE.4></CE.4>
'         <CE.5></CE.5>
'         <CE.6>99NCI</CE.6>
'      </RXE.2>
'      <RXE.3></RXE.3>
'      <RXE.4></RXE.4>
'      <RXE.5>
'         <CE.1></CE.1>
'         <CE.2></CE.2>
'         <CE.3></CE.3>
'         <CE.4></CE.4>
'         <CE.5></CE.5>
'         <CE.6>99NCI</CE.6>
'      </RXE.5>
'      <RXE.7.LST>
'         <RXE.7>
'            <CE.1></CE.1>
'            <CE.2>Sysadmin</CE.2>
'            <CE.3>I9</CE.3>
'            <CE.4></CE.4>
'            <CE.5>Sysadmin</CE.5>
'            <CE.6></CE.6>
'         </RXE.7>
'      </RXE.7.LST>
'      <RXE.10></RXE.10>
'      <RXE.12>2</RXE.12>
'      <RXE.16>1</RXE.16>
'      <RXE.18>20010726</RXE.18>
'   </RXE>

    Dim oDom As MSXML.DOMDocument
    Dim oRXENode As MSXML.IXMLDOMElement
    Dim oRXEChildLevel1 As MSXML.IXMLDOMElement
    Dim oRXEChildLevel2 As MSXML.IXMLDOMElement
    Dim oRXEChildLevel3 As MSXML.IXMLDOMElement
    Dim strTemp As String
    Dim strQuantity As String
    
    On Error GoTo ErrTrap

    Set oDom = New MSXML.DOMDocument
    Set oRXENode = oDom.createElement("RXE")
    
    ' Quantity/Timing
    Set oRXEChildLevel1 = oDom.createElement("RXE.1")
    strTemp = NullToNothing(MedicationsRS.Fields("StartDate").Value)
    If Not IsValidDate(strTemp) Then
        strQuantity = "^^^"
    Else
        strQuantity = "^^^" & Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    strTemp = NullToNothing(MedicationsRS.Fields("EndDate").Value)
    If Not IsValidDate(strTemp) Then
        strQuantity = strQuantity & "^"
    Else
        strQuantity = strQuantity & Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oRXEChildLevel1.Text = strQuantity
    oRXENode.appendChild oRXEChildLevel1

    ' Give Code
    Set oRXEChildLevel1 = oDom.createElement("RXE.2")
        ' Identifier NCID - Table Medications:MedicationsNCID
        Set oRXEChildLevel2 = oDom.createElement("CE.1")
        oRXEChildLevel2.Text = ""
        oRXEChildLevel1.appendChild oRXEChildLevel2
        ' Text - Table Medications:
        Set oRXEChildLevel2 = oDom.createElement("CE.2")
        strTemp = NullToNothing(MedicationsRS.Fields("MedicationNCID").Value)
        If strTemp <> "" Then
            oRXEChildLevel2.Text = GetMedicationNameText(CDbl(strTemp)) '<SCR 23902 CC> - Change from long to double
        End If
        oRXEChildLevel1.appendChild oRXEChildLevel2
        ' Name of Coding System (=99NCI) - Table Medications:
        Set oRXEChildLevel2 = oDom.createElement("CE.3")
        oRXEChildLevel2.Text = ""
        oRXEChildLevel1.appendChild oRXEChildLevel2
        ' Alternate ID - Table Medications:NationalDrugCode
        Set oRXEChildLevel2 = oDom.createElement("CE.4")
        oRXEChildLevel2.Text = strTemp 'MedicationNCID
        oRXEChildLevel1.appendChild oRXEChildLevel2
        ' Text - Table Medications:
        Set oRXEChildLevel2 = oDom.createElement("CE.5")
        strTemp = NullToNothing(MedicationsRS.Fields("MedicationNCID").Value)
        If strTemp <> "" Then
            oRXEChildLevel2.Text = GetMedicationNameText(CDbl(strTemp)) '<SCR 23902 CC> - Change from long to double
        End If
        oRXEChildLevel1.appendChild oRXEChildLevel2
        ' Alternate Coding System  - Table Medications:
        Set oRXEChildLevel2 = oDom.createElement("CE.6")
        oRXEChildLevel2.Text = "99NCI"
        oRXEChildLevel1.appendChild oRXEChildLevel2
    oRXENode.appendChild oRXEChildLevel1
    
    ' Give Amount - Minimum - Table Medications:Low_Dose_Value
    Set oRXEChildLevel1 = oDom.createElement("RXE.3")
    
    ' Begin:SCR #27481;   Developer: MAG 09/06/2002 05:05 PM
    '<< Begin:SCR #26440; 27061;   Developer: Brian Mowbray 08/30/2002 04:48 PM
    oRXEChildLevel1.Text = "0"
    '>> End: SCR #26440; 27061;
    
    oRXENode.appendChild oRXEChildLevel1  '' value not currently in DB Table

    ' Give Amount - Maximum - Table Medications:High_Dose_Value
    Set oRXEChildLevel1 = oDom.createElement("RXE.4")
    '<< Begin:SCR #26440; 27061;   Developer: Brian Mowbray 08/30/2002 04:48 PM
    oRXEChildLevel1.Text = "0"
    '>> End: SCR #26440; 27061;
    oRXENode.appendChild oRXEChildLevel1  '' value not currently in DB Table
    
    ' Begin:SCR #28647;   Developer: MAG 10/28/2002 04:49 PM
    ' Give Units (NCID) - Table Medications:Dose_Units_NCID
    Set oRXEChildLevel1 = oDom.createElement("RXE.5")
        ' Identifier NCID
        Set oRXEChildLevel2 = oDom.createElement("CE.1")
        oRXEChildLevel2.Text = ""
        oRXEChildLevel1.appendChild oRXEChildLevel2
        ' Text
        Set oRXEChildLevel2 = oDom.createElement("CE.2")
        oRXEChildLevel2.Text = "0"
        oRXEChildLevel1.appendChild oRXEChildLevel2
        ' Name of Coding System (=99NCI) - Table Medications:
        Set oRXEChildLevel2 = oDom.createElement("CE.3")
        oRXEChildLevel2.Text = ""
        oRXEChildLevel1.appendChild oRXEChildLevel2
        ' Alternate ID
        Set oRXEChildLevel2 = oDom.createElement("CE.4")
        oRXEChildLevel2.Text = ""
        oRXEChildLevel1.appendChild oRXEChildLevel2
        ' Text - Table Medications:
        Set oRXEChildLevel2 = oDom.createElement("CE.5")
            oRXEChildLevel2.Text = ""
        oRXEChildLevel1.appendChild oRXEChildLevel2
        ' Alternate Coding System  - Table Medications:
        Set oRXEChildLevel2 = oDom.createElement("CE.6")
        oRXEChildLevel2.Text = "99NCI"
        oRXEChildLevel1.appendChild oRXEChildLevel2
    oRXENode.appendChild oRXEChildLevel1
    ' End: SCR #28647;

    ' Providers Administration Instructions - Table Medications:Sig
    Set oRXEChildLevel1 = oDom.createElement("RXE.7.LST")
        Set oRXEChildLevel2 = oDom.createElement("RXE.7")
            ' Identifier
            Set oRXEChildLevel3 = oDom.createElement("CE.1")
            oRXEChildLevel3.Text = ""
            oRXEChildLevel2.appendChild oRXEChildLevel3
            ' Text
            Set oRXEChildLevel3 = oDom.createElement("CE.2")
            oRXEChildLevel3.Text = NullToNothing(MedicationsRS.Fields("Sig").Value)
            oRXEChildLevel2.appendChild oRXEChildLevel3
            ' Name of Coding System
            Set oRXEChildLevel3 = oDom.createElement("CE.3")
            oRXEChildLevel3.Text = ""
            oRXEChildLevel2.appendChild oRXEChildLevel3
            ' Alternate Identifier
            Set oRXEChildLevel3 = oDom.createElement("CE.4")
            oRXEChildLevel3.Text = ""
            oRXEChildLevel2.appendChild oRXEChildLevel3
            ' Alternate Text
            Set oRXEChildLevel3 = oDom.createElement("CE.5")
            oRXEChildLevel3.Text = NullToNothing(MedicationsRS.Fields("Sig").Value)
            oRXEChildLevel2.appendChild oRXEChildLevel3
            ' Alternate Name of Codeing System
            Set oRXEChildLevel3 = oDom.createElement("CE.6")
            oRXEChildLevel3.Text = ""
            oRXEChildLevel2.appendChild oRXEChildLevel3
        oRXEChildLevel1.appendChild oRXEChildLevel2
    oRXENode.appendChild oRXEChildLevel1
    
    ' Dispense Amount - Table Medications:Dispense_Quantity
    Set oRXEChildLevel1 = oDom.createElement("RXE.10")
    ' Begin:SCR #27481;   Developer: MAG 09/06/2002 05:05 PM
    '<< Begin:SCR #26440; 27061;   Developer: Brian Mowbray 08/30/2002 04:50 PM
    oRXEChildLevel1.Text = NullToZero(MedicationsRS.Fields("Quantity").Value)
    '>> End: SCR #26440; 27061;
    ' End: SCR #27481;
    oRXENode.appendChild oRXEChildLevel1    ' value not currently in DB Table
    
    ' Number of Refills - Table Medications:Refills
    Set oRXEChildLevel1 = oDom.createElement("RXE.12")
    oRXEChildLevel1.Text = NullToNothing(MedicationsRS.Fields("Refills").Value)
    oRXENode.appendChild oRXEChildLevel1
    
    ' Number of Refills Remaining - Table Medications:RefillsRemaining
    Set oRXEChildLevel1 = oDom.createElement("RXE.16")
    oRXEChildLevel1.Text = NullToNothing(MedicationsRS.Fields("RefillsRemaining").Value)
    oRXENode.appendChild oRXEChildLevel1
    
    ' Date/Time of Most Recent Refill or Dose Dispensed - Table Medications:FillDate
    Set oRXEChildLevel1 = oDom.createElement("RXE.18")
    strTemp = NullToNothing(MedicationsRS.Fields("FillDate").Value)
    If IsValidDate(strTemp) Then
        oRXEChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oRXENode.appendChild oRXEChildLevel1
    
    Set ReturnNewRXENode = oRXENode

Exit Function

ErrTrap:
   If gobjShared Is Nothing Then
      Set gobjShared = New CWShared
   End If
   gobjShared.ShowVBError Err.Number, Err.Description, "ReturnNewRXENode Function", "DMBE", vbExclamation
   'Err.Clear
   Resume Next 'There is no reason this procedure should not continue to create xml if it errors
            
End Function

Public Function ReturnNewORCNode(ByVal MedicationsRS As ADODB.Recordset) As MSXML.IXMLDOMElement
   ' Common Order Segment
   
' sample function output
'   <ORC>
'      <ORC.1>NW</ORC.1>
'      <ORC.2>13319926-200487-2:57:16 PM-0</ORC.2>
'      <ORC.5>fainting followed by a headache</ORC.5>
'      <ORC.9>20010726000000</ORC.9>
'      <ORC.12>
'         <XCN.1></XCN.1>
'         <XCN.2></XCN.2>
'         <XCN.3></XCN.3>
'         <XCN.4></XCN.4>
'      </ORC.12>
'      <ORC.15>20010724</ORC.15>
'   </ORC>

    Dim oDom As MSXML.DOMDocument
    Dim oORCNode As MSXML.IXMLDOMElement
    Dim oORCChildLevel1 As MSXML.IXMLDOMElement
    Dim oORCChildLevel2 As MSXML.IXMLDOMElement
    Dim oORCChildLevel3 As MSXML.IXMLDOMElement
    Dim strDocFirstName As String
    Dim strDocLastName As String
    Dim strDocMiddleName As String
    Dim strDocID As String
    Dim strDocSpecialty As String
    Dim strTemp As String
    
    On Error GoTo ErrTrap

    Set oDom = New MSXML.DOMDocument
    Set oORCNode = oDom.createElement("ORC")
    
    ' Order Control - default to "NW" (new order)
    Set oORCChildLevel1 = oDom.createElement("ORC.1")
    oORCChildLevel1.Text = "NW"
    oORCNode.appendChild oORCChildLevel1

    Set oORCChildLevel1 = oDom.createElement("ORC.2")
        ' Placer Order Number - Table Medications:OrderNumber
        Set oORCChildLevel2 = oDom.createElement("EI.1")
        oORCChildLevel2.Text = NullToNothing(MedicationsRS.Fields("OrderNumber").Value)
        oORCChildLevel1.appendChild oORCChildLevel2
    oORCNode.appendChild oORCChildLevel1
    
    Set oORCChildLevel1 = oDom.createElement("ORC.5")
    strTemp = NullToNothing(MedicationsRS.Fields("StatusNCID").Value)
    If Not strTemp = "" Then
        oORCChildLevel1.Text = GetMedicationsStatusText(CLng(strTemp))
    End If
    oORCNode.appendChild oORCChildLevel1

    ' Order Data/Time - Table Medications:OrderDate
    Set oORCChildLevel1 = oDom.createElement("ORC.9")
    strTemp = NullToNothing(MedicationsRS.Fields("OrderDate").Value)
    If IsValidDate(strTemp) Then
        oORCChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oORCNode.appendChild oORCChildLevel1
    
    ' Ordering Provider
    Set oORCChildLevel1 = oDom.createElement("ORC.12.LST")
        Set oORCChildLevel2 = oDom.createElement("ORC.12")
            
            strTemp = NullToNothing(MedicationsRS.Fields("OrderingProviderNCID").Value)
            If strTemp <> "" Then
                Call GetNewDoctorInfo(CLng(strTemp), strDocLastName, strDocFirstName, strDocMiddleName, strDocID, strDocSpecialty)
            End If
            
            ' Ordering Provider SSN - Table Medications:OrderingProviderNCID
            Set oORCChildLevel3 = oDom.createElement("XCN.1")
            oORCChildLevel3.Text = strDocID
            oORCChildLevel2.appendChild oORCChildLevel3
            ' Ordering Provider Last Name
            Set oORCChildLevel3 = oDom.createElement("XCN.2")
            oORCChildLevel3.Text = strDocLastName
            oORCChildLevel2.appendChild oORCChildLevel3
            ' Ordering Provider First Name
            Set oORCChildLevel3 = oDom.createElement("XCN.3")
            oORCChildLevel3.Text = strDocFirstName
            oORCChildLevel2.appendChild oORCChildLevel3
            ' Ordering Provider Middle Name
            Set oORCChildLevel3 = oDom.createElement("XCN.4")
            oORCChildLevel3.Text = strDocMiddleName
            oORCChildLevel2.appendChild oORCChildLevel3
            
        oORCChildLevel1.appendChild oORCChildLevel2
    oORCNode.appendChild oORCChildLevel1
    
    ' Order Effective Data/Time - Table Medications:StartDate
    Set oORCChildLevel1 = oDom.createElement("ORC.15")
    strTemp = NullToNothing(MedicationsRS.Fields("StartDate").Value)
    If IsValidDate(strTemp) Then
        oORCChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oORCNode.appendChild oORCChildLevel1
    
    Set ReturnNewORCNode = oORCNode

Exit Function

ErrTrap:
   If gobjShared Is Nothing Then
      Set gobjShared = New CWShared
   End If
   gobjShared.ShowVBError Err.Number, Err.Description, "ReturnNewORCNode Function", "DMBE", vbExclamation
   'Err.Clear
   Resume Next 'There is no reason this procedure should not continue to create xml if it errors

End Function


Public Function ReturnNewZALNode(ByVal lngCount As Long, ByVal AllergyRS As ADODB.Recordset, ByVal CommentRS As ADODB.Recordset, ByVal ReactionRS As ADODB.Recordset) As MSXML.IXMLDOMElement
' Additional Allergy Data Segment

' sample function output
'   <ZAL>
'      <ZAL.1>Count</ZAL.1>
'      <ZAL.2>Data ID</ZAL.2>
'      <ZAL.3>Unit Number</ZAL.3>
'      <ZAL.4>Facility NCID</ZAL.4>
'      <ZAL.5>
'            <XCN.1>Attending Doctor ID (SSN)</XCN.1>
'            <XCN.2>Attending Doctor Last Name</XCN.2>
'            <XCN.3>Attending Doctor First Name</XCN.3>
'            <XCN.4>Attending Doctor Middle Name</XCN.4>
'      </ZAL.5>
'      <ZAL.6>Create Date/Time</ZAL.6>
'      <ZAL.7>Info Source NCID</ZAL.7>
'      <ZAL.8>Modified</ZAL.8>
'      <ZAL.9>Modified Date/Time</ZAL.9>
'      <ZAL.10>Modify User</ZAL.10>
'      <ZAL.11>Point of Care Facility NCID</ZAL.11>
'      <ZAL.12>Create User</ZAL.12>
'      <ZAL.13>Update Flag</ZAL.13>
'      <ZAL.14>CDR Data ID</ZAL.14>
'      <ZAL.15>Created By</ZAL.15>
'      <ZAL.16>Created On</ZAL.16>
'      <ZAL.17>Updated By</ZAL.17>
'      <ZAL.18>Updated On</ZAL.18>
'      <ZAL.19>Allergy Comments Unit Number</ZAL.19>
'      <ZAL.20>Allergy Comments Allergen NCID</ZAL.20>
'      <ZAL.21>Doctor Comments</ZAL.21>
'      <ZAL.22>Allergy Comments Update Flag</ZAL.22>
'      <ZAL.23>Allergy Comments Created By</ZAL.23>
'      <ZAL.24>Allergy Comments Created On</ZAL.24>
'      <ZAL.25>Allergy Comments Updated By</ZAL.25>
'      <ZAL.26>Allergy Comments Updated On</ZAL.26>
'      <ZAL.27>Allergy Reaction Unit Number</ZAL.27>
'      <ZAL.28>Allergy Reaction Allergen NCID</ZAL.28>
'      <ZAL.29>Allergy Reaction Update Flag</ZAL.29>
'      <ZAL.30>Allergy Reaction Created By</ZAL.30>
'      <ZAL.31>Allergy Reaction Created On</ZAL.31>
'      <ZAL.32>Allergy Reaction Updated By</ZAL.32>
'      <ZAL.33>Allergy Reaction Updated On</ZAL.33>
'   </ZAL>

    Dim oDom As MSXML.DOMDocument
    Dim oZALNode As MSXML.IXMLDOMElement
    Dim oZALChildLevel1 As MSXML.IXMLDOMElement
    Dim oZALChildLevel2 As MSXML.IXMLDOMElement
    Dim strDocFirstName As String
    Dim strDocLastName As String
    Dim strDocMiddleName As String
    Dim strDocID As String
    Dim strDocSpecialty As String
    Dim strTemp As String
    Dim blnRsIsGood As Boolean
    
    On Error GoTo ErrTrap

    Set oDom = New MSXML.DOMDocument
    Set oZALNode = oDom.createElement("ZAL")
    
    ' Additional Allergy Count
    Set oZALChildLevel1 = oDom.createElement("ZAL.1")
    oZALChildLevel1.Text = CStr(lngCount)
    oZALNode.appendChild oZALChildLevel1

    ' Data ID - Table Allergies:DataID
    Set oZALChildLevel1 = oDom.createElement("ZAL.2")
    oZALChildLevel1.Text = AllergyRS.Fields("DataID").Value
    oZALNode.appendChild oZALChildLevel1

    ' Unit Number - Table Allergies:Unit_Number
    Set oZALChildLevel1 = oDom.createElement("ZAL.3")
    oZALChildLevel1.Text = AllergyRS.Fields("Unit_Number").Value
    oZALNode.appendChild oZALChildLevel1

    ' Facility NCID - Table Allergies:Facility_NCID
    Set oZALChildLevel1 = oDom.createElement("ZAL.4")
    oZALChildLevel1.Text = AllergyRS.Fields("Facility_NCID").Value
    oZALNode.appendChild oZALChildLevel1

    ' Allergy Clinician - Table Allergies:ClinicianNCID
    Set oZALChildLevel1 = oDom.createElement("ZAL.5")
        
        Call GetNewDoctorInfo(AllergyRS.Fields("ClinicianNCID").Value, strDocLastName, strDocFirstName, strDocMiddleName, strDocID, strDocSpecialty)
            
        ' Attending Doctor ID (SSN)
        Set oZALChildLevel2 = oDom.createElement("XCN.1")
        oZALChildLevel2.Text = strDocID
        oZALChildLevel1.appendChild oZALChildLevel2
        ' Attending Doctor Last Name
        Set oZALChildLevel2 = oDom.createElement("XCN.2")
        oZALChildLevel2.Text = strDocLastName
        oZALChildLevel1.appendChild oZALChildLevel2
        ' Attending Doctor first Name
        Set oZALChildLevel2 = oDom.createElement("XCN.3")
        oZALChildLevel2.Text = strDocFirstName
        oZALChildLevel1.appendChild oZALChildLevel2
        ' Attending Doctor Middle Name
        Set oZALChildLevel2 = oDom.createElement("XCN.4")
        oZALChildLevel2.Text = strDocMiddleName
        oZALChildLevel1.appendChild oZALChildLevel2
        
    oZALNode.appendChild oZALChildLevel1
    
    ' Create Date/Time - Table Allergies:CreateTime
    Set oZALChildLevel1 = oDom.createElement("ZAL.6")
    oZALChildLevel1.Text = NullToNothing(AllergyRS.Fields("CreateTime").Value)
    oZALNode.appendChild oZALChildLevel1
    
    '<SCR 22974 CC>
    ' Info Source NCID - Table Allergies:InfoSourceNCID
    Set oZALChildLevel1 = oDom.createElement("ZAL.7")
    strTemp = NullToNothing(AllergyRS.Fields("InfoSourceNCID").Value)
        Set oZALChildLevel2 = oDom.createElement("CE.1")
        oZALChildLevel1.appendChild oZALChildLevel2
        
        Set oZALChildLevel2 = oDom.createElement("CE.2")
        If strTemp <> "" Then
            oZALChildLevel2.Text = GetNCIDDescription(CLng(strTemp))
        End If
        oZALChildLevel1.appendChild oZALChildLevel2
        
        Set oZALChildLevel2 = oDom.createElement("CE.3")
        oZALChildLevel1.appendChild oZALChildLevel2
        Set oZALChildLevel2 = oDom.createElement("CE.4")
        oZALChildLevel1.appendChild oZALChildLevel2
        Set oZALChildLevel2 = oDom.createElement("CE.5")
        oZALChildLevel2.Text = strTemp 'The NCID
        oZALChildLevel1.appendChild oZALChildLevel2
        Set oZALChildLevel2 = oDom.createElement("CE.6")
        oZALChildLevel2.Text = "99NCI"
        oZALChildLevel1.appendChild oZALChildLevel2
    oZALNode.appendChild oZALChildLevel1
    '</SCR 22974 CC>

    ' Modified - Table Allergies:Modified
    Set oZALChildLevel1 = oDom.createElement("ZAL.8")
    oZALChildLevel1.Text = NullToNothing(AllergyRS.Fields("Modified").Value)
    oZALNode.appendChild oZALChildLevel1

    ' Modify Date/Time - Table Allergies:ModifyTime
    Set oZALChildLevel1 = oDom.createElement("ZAL.9")
    strTemp = NullToNothing(AllergyRS.Fields("ModifyTime").Value)
    If IsValidDate(strTemp) Then
        oZALChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZALNode.appendChild oZALChildLevel1

    ' Modify User - Table Allergies:ModifyUserNCID
    Set oZALChildLevel1 = oDom.createElement("ZAL.10")
    oZALChildLevel1.Text = NullToNothing(AllergyRS.Fields("ModifyUserNCID").Value)
    oZALNode.appendChild oZALChildLevel1
    
    ' Point of Care Facility NCID - Table Allergies:PointOfCareFacilityNCID
    Set oZALChildLevel1 = oDom.createElement("ZAL.11")
    'oZALChildLevel1.Text = NullToNothing(AllergyRS.Fields("PointOfCareFacilityNCID").Value)
    strTemp = NullToNothing(AllergyRS.Fields("PointOfCareFacilityNCID").Value)
        Set oZALChildLevel2 = oDom.createElement("CE.1")
        oZALChildLevel1.appendChild oZALChildLevel2
        
        Set oZALChildLevel2 = oDom.createElement("CE.2")
        If strTemp <> "" Then
            oZALChildLevel2.Text = GetNCIDDescription(CLng(strTemp))
        End If
        oZALChildLevel1.appendChild oZALChildLevel2
        
        Set oZALChildLevel2 = oDom.createElement("CE.3")
        oZALChildLevel1.appendChild oZALChildLevel2
        Set oZALChildLevel2 = oDom.createElement("CE.4")
        oZALChildLevel1.appendChild oZALChildLevel2
        Set oZALChildLevel2 = oDom.createElement("CE.5")
        oZALChildLevel2.Text = strTemp 'The NCID
        oZALChildLevel1.appendChild oZALChildLevel2
        Set oZALChildLevel2 = oDom.createElement("CE.6")
        oZALChildLevel2.Text = "99NCI"
        oZALChildLevel1.appendChild oZALChildLevel2
    oZALNode.appendChild oZALChildLevel1
    '</SCR 22974 CC>
    
    ' Create User - Table Allergies:CreateUserNCID
    Set oZALChildLevel1 = oDom.createElement("ZAL.12")
    oZALChildLevel1.Text = NullToNothing(AllergyRS.Fields("CreateUserNCID").Value)
    oZALNode.appendChild oZALChildLevel1
    
    ' Update Flag - Table Allergies:UpdateFlag
    Set oZALChildLevel1 = oDom.createElement("ZAL.13")
    oZALChildLevel1.Text = NullToNothing(AllergyRS.Fields("UpdateFlag").Value)
    oZALNode.appendChild oZALChildLevel1
    
    ' CDR Data ID
    Set oZALChildLevel1 = oDom.createElement("ZAL.14")
    oZALChildLevel1.Text = NullToNothing(AllergyRS.Fields("CDR_DataID").Value)
    oZALNode.appendChild oZALChildLevel1

    ' Created By - Table Allergies:CreatedBy
    Set oZALChildLevel1 = oDom.createElement("ZAL.15")
    oZALChildLevel1.Text = NullToNothing(AllergyRS.Fields("CreatedBy").Value)
    oZALNode.appendChild oZALChildLevel1

    ' Created On - Table Allergies:CreatedOn
    Set oZALChildLevel1 = oDom.createElement("ZAL.16")
    strTemp = NullToNothing(AllergyRS.Fields("CreatedOn").Value)
    If IsValidDate(strTemp) Then
        oZALChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZALNode.appendChild oZALChildLevel1
    
    ' Updated By - Table Allergies:UpdatedBy
    Set oZALChildLevel1 = oDom.createElement("ZAL.17")
    oZALChildLevel1.Text = NullToNothing(AllergyRS.Fields("UpdatedBy").Value)
    oZALNode.appendChild oZALChildLevel1

    ' Updated On - Table Allergies:UpdatedOn
    Set oZALChildLevel1 = oDom.createElement("ZAL.18")
    strTemp = NullToNothing(AllergyRS.Fields("UpdatedOn").Value)
    If IsValidDate(strTemp) Then
        oZALChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZALNode.appendChild oZALChildLevel1
    
    blnRsIsGood = False
    If Not (CommentRS Is Nothing) Then
        If Not CommentRS.EOF Then
            blnRsIsGood = True
        End If
    End If
     
    ' AllergyComment Unit Number - Table AllergyComment:Unit_Number
    Set oZALChildLevel1 = oDom.createElement("ZAL.19")
    If blnRsIsGood = True Then
        oZALChildLevel1.Text = NullToNothing(CommentRS.Fields("Unit_Number").Value)
    End If
    oZALNode.appendChild oZALChildLevel1

    ' AllergyComment AllergenNCID - Table AllergyComment:AllergenNCID
    Set oZALChildLevel1 = oDom.createElement("ZAL.20")
    If blnRsIsGood = True Then
        oZALChildLevel1.Text = NullToNothing(CommentRS.Fields("AllergenNCID").Value)
    End If
    oZALNode.appendChild oZALChildLevel1

    ' Allergy Comment - Table AllergyComment:ALLERGYCOMMENT
    Set oZALChildLevel1 = oDom.createElement("ZAL.21")
    If blnRsIsGood = True Then
        oZALChildLevel1.Text = NullToNothing(CommentRS.Fields("ALLERGYCOMMENT").Value)
    End If
    oZALNode.appendChild oZALChildLevel1

    ' AllergyComment Update Flag - Table AllergyComment:UpdateFlag
    Set oZALChildLevel1 = oDom.createElement("ZAL.22")
    If blnRsIsGood = True Then
        oZALChildLevel1.Text = NullToNothing(CommentRS.Fields("UpdateFlag").Value)
    End If
    oZALNode.appendChild oZALChildLevel1

    ' AllergyComment Created By - Table AllergyComment:CreatedBy
    Set oZALChildLevel1 = oDom.createElement("ZAL.23")
    If blnRsIsGood = True Then
        oZALChildLevel1.Text = NullToNothing(CommentRS.Fields("CreatedBy").Value)
    End If
    oZALNode.appendChild oZALChildLevel1

    ' AllergyComment Created On - Table AllergyComment:CreatedOn
    Set oZALChildLevel1 = oDom.createElement("ZAL.24")
    If blnRsIsGood = True Then
        strTemp = NullToNothing(CommentRS.Fields("CreatedOn").Value)
        If IsValidDate(strTemp) Then
            oZALChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
        End If
    End If
    oZALNode.appendChild oZALChildLevel1

    ' AllergyComment Updated By - Table AllergyComment:UpdatedBy
    Set oZALChildLevel1 = oDom.createElement("ZAL.25")
    If blnRsIsGood = True Then
        oZALChildLevel1.Text = NullToNothing(CommentRS.Fields("UpdatedBy").Value)
    End If
    oZALNode.appendChild oZALChildLevel1

    ' AllergyComment Updated On - Table AllergyComment:UpdatedOn
    Set oZALChildLevel1 = oDom.createElement("ZAL.26")
    If blnRsIsGood = True Then
        strTemp = NullToNothing(CommentRS.Fields("UpdatedOn").Value)
        If IsValidDate(strTemp) Then
            oZALChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
        End If
    End If
    oZALNode.appendChild oZALChildLevel1

    blnRsIsGood = False
    If Not (ReactionRS Is Nothing) Then
        If Not ReactionRS.EOF Then
            blnRsIsGood = True
        End If
    End If

    ' Allergy Reaction Unit Number - Table AllergyReaction:Unit_Number
    Set oZALChildLevel1 = oDom.createElement("ZAL.27")
    If blnRsIsGood = True Then
        oZALChildLevel1.Text = NullToNothing(ReactionRS.Fields("Unit_Number").Value)
    End If
    oZALNode.appendChild oZALChildLevel1

    ' Allergy Reaction Allergen NCID - Table AllergyReaction:AllergenNCID
    Set oZALChildLevel1 = oDom.createElement("ZAL.28")
    If blnRsIsGood = True Then
        oZALChildLevel1.Text = NullToNothing(ReactionRS.Fields("AllergenNCID").Value)
    End If
    oZALNode.appendChild oZALChildLevel1

    ' Allergy Reaction Update Flag - Table AllergyReaction:UpdateFlag
    Set oZALChildLevel1 = oDom.createElement("ZAL.29")
    If blnRsIsGood = True Then
        oZALChildLevel1.Text = NullToNothing(ReactionRS.Fields("UpdateFlag").Value)
    End If
    oZALNode.appendChild oZALChildLevel1

    ' Allergy Reaction Created By - Table AllergyReaction:CreatedBy
    Set oZALChildLevel1 = oDom.createElement("ZAL.30")
    If blnRsIsGood = True Then
        oZALChildLevel1.Text = NullToNothing(ReactionRS.Fields("CreatedBy").Value)
    End If
    oZALNode.appendChild oZALChildLevel1

    ' Allergy Reaction Created On - Table AllergyReaction:CreatedOn
    Set oZALChildLevel1 = oDom.createElement("ZAL.31")
    If blnRsIsGood = True Then
        strTemp = NullToNothing(ReactionRS.Fields("CreatedOn").Value)
        If IsValidDate(strTemp) Then
            oZALChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
        End If
    End If
    oZALNode.appendChild oZALChildLevel1

    ' Allergy Reaction Updated By - Table AllergyReaction:UpdatedBy
    Set oZALChildLevel1 = oDom.createElement("ZAL.32")
    If blnRsIsGood = True Then
        oZALChildLevel1.Text = NullToNothing(ReactionRS.Fields("UpdatedBy").Value)
    End If
    oZALNode.appendChild oZALChildLevel1

    ' Allergy Reaction Updated On - Table AllergyReaction:UpdatedOn
    Set oZALChildLevel1 = oDom.createElement("ZAL.33")
    If blnRsIsGood = True Then
        strTemp = NullToNothing(ReactionRS.Fields("UpdatedOn").Value)
        If IsValidDate(strTemp) Then
            oZALChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
        End If
    End If
    oZALNode.appendChild oZALChildLevel1

    'New requirement
    Set oZALChildLevel1 = oDom.createElement("ZAL.34")
    If blnRsIsGood = True Then
        
        strTemp = NullToNothing(ReactionRS.Fields("REACTIONNCID").Value)
        
        Set oZALChildLevel2 = oDom.createElement("CE.1")
        oZALChildLevel2.Text = "477.9"
        oZALChildLevel1.appendChild oZALChildLevel2
        
        Set oZALChildLevel2 = oDom.createElement("CE.2")
        If strTemp <> "" Then
            oZALChildLevel2.Text = GetNCIDDescription(CLng(strTemp))
        End If
        oZALChildLevel1.appendChild oZALChildLevel2
        
        Set oZALChildLevel2 = oDom.createElement("CE.3")
        oZALChildLevel2.Text = "I9"
        oZALChildLevel1.appendChild oZALChildLevel2
                
        Set oZALChildLevel2 = oDom.createElement("CE.4")
        oZALChildLevel2.Text = strTemp 'The NCID
        oZALChildLevel1.appendChild oZALChildLevel2
        
        Set oZALChildLevel2 = oDom.createElement("CE.5")
        oZALChildLevel2.Text = NullToNothing(ReactionRS.Fields("REACTION").Value)
        oZALChildLevel1.appendChild oZALChildLevel2
        
        Set oZALChildLevel2 = oDom.createElement("CE.6")
        oZALChildLevel2.Text = "99NCI"
        oZALChildLevel1.appendChild oZALChildLevel2
    End If
    oZALNode.appendChild oZALChildLevel1

    Set oZALChildLevel1 = oDom.createElement("ZAL.35")
    strTemp = NullToNothing(AllergyRS.Fields("ALLERGYTYPENCID").Value)
        
        Set oZALChildLevel2 = oDom.createElement("CE.1")
        If strTemp <> "" Then
            oZALChildLevel2.Text = GetNCIDDescription(CLng(strTemp))
        End If
        oZALChildLevel1.appendChild oZALChildLevel2
        
        Set oZALChildLevel2 = oDom.createElement("CE.2")
        oZALChildLevel2.Text = strTemp
        oZALChildLevel1.appendChild oZALChildLevel2
        
        Set oZALChildLevel2 = oDom.createElement("CE.3")
        oZALChildLevel2.Text = "99NCI"
        oZALChildLevel1.appendChild oZALChildLevel2
                
        Set oZALChildLevel2 = oDom.createElement("CE.4")
        oZALChildLevel1.appendChild oZALChildLevel2
        Set oZALChildLevel2 = oDom.createElement("CE.5")
        oZALChildLevel1.appendChild oZALChildLevel2
        Set oZALChildLevel2 = oDom.createElement("CE.6")
        oZALChildLevel1.appendChild oZALChildLevel2
    oZALNode.appendChild oZALChildLevel1

    Set oZALChildLevel1 = oDom.createElement("ZAL.36")
    strTemp = NullToNothing(AllergyRS.Fields("ALLERGENNCID").Value)
        
        Set oZALChildLevel2 = oDom.createElement("CE.1")
        oZALChildLevel2.Text = "477.9"
        oZALChildLevel1.appendChild oZALChildLevel2
        
        Set oZALChildLevel2 = oDom.createElement("CE.2")
        oZALChildLevel2.Text = "Allergy (Other)"
        oZALChildLevel1.appendChild oZALChildLevel2
        
        Set oZALChildLevel2 = oDom.createElement("CE.3")
        oZALChildLevel2.Text = "I9"
        oZALChildLevel1.appendChild oZALChildLevel2
                
        Set oZALChildLevel2 = oDom.createElement("CE.4")
        oZALChildLevel2.Text = NullToNothing(AllergyRS.Fields("ALLERGENNCID").Value)
        oZALChildLevel1.appendChild oZALChildLevel2
        
        Set oZALChildLevel2 = oDom.createElement("CE.5")
        oZALChildLevel2.Text = NullToNothing(AllergyRS.Fields("ALLERGEN").Value)
        oZALChildLevel1.appendChild oZALChildLevel2
        
        Set oZALChildLevel2 = oDom.createElement("CE.6")
        oZALChildLevel2.Text = "99NCI"
        oZALChildLevel1.appendChild oZALChildLevel2
    
    oZALNode.appendChild oZALChildLevel1

    Set oZALChildLevel1 = oDom.createElement("ZAL.37")
    strTemp = NullToNothing(AllergyRS.Fields("NOTEDDATE").Value)
    If IsValidDate(strTemp) Then
        oZALChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
    End If
    oZALNode.appendChild oZALChildLevel1

    Set ReturnNewZALNode = oZALNode

Exit Function

ErrTrap:
   If gobjShared Is Nothing Then
      Set gobjShared = New CWShared
   End If
   gobjShared.ShowVBError Err.Number, Err.Description, "ReturnNewZALNode Function", "DMBE", vbExclamation
   'Err.Clear
   Resume Next 'There is no reason this procedure should not continue to create xml if it errors

End Function

'<SCR #19356;   Developer: CC 05/02/2002 10:26 AM>
'SCR Description: Old logic was wrong

'Public Function ReturnNewAL1Node(ByVal lngCount As Long, ByVal AllergyRS As ADODB.Recordset, ByVal ReactionRS As ADODB.Recordset) As MSXML.IXMLDOMElement
'
'' sample function output
''   <ADT_A28.LST.4>
''      <AL1>
''         <AL1.1>1</AL1.1>
''         <AL1.2>MA</AL1.2>
''         <AL1.3>
''            <CE.1>477.9</CE.1>
''            <CE.2>Allergy (Other)</CE.2>
''            <CE.3>I9</CE.3>
''            <CE.4>1000</CE.4>
''            <CE.5>Penicillins</CE.5>
''            <CE.6>99NCI</CE.6>
''         </AL1.3>
''      <AL1.5.LST>
''         <AL1.5>
''            <CE.1></CE.1>
''            <CE.2>Rash</CE.2>
''            <CE.3></CE.3>
''            <CE.4>82559</CE.4>
''            <CE.5>Rash</CE.5>
''            <CE.6>99NCI</CE.6>
''         </AL1.5>
''      </AL1.5.LST>
''         <AL1.6>20010326</AL1.6>
''      </AL1>
''   </ADT_A28.LST.4>
'
'    Dim oDom As MSXML.DOMDocument
'    Dim oAL1Node As MSXML.IXMLDOMElement
'    Dim oAL1ChildLevel1 As MSXML.IXMLDOMElement
'    Dim oAL1ChildLevel2 As MSXML.IXMLDOMElement
'    Dim oAL1ChildLevel3 As MSXML.IXMLDOMElement
'    Dim strAllergyType As String
'    Dim strName As String
'
'    On Error GoTo ErrTrap
'
'    Set oDom = New MSXML.DOMDocument
'    Set oAL1Node = oDom.createElement("AL1")
'
'    ' Allergy Count
'    Set oAL1ChildLevel1 = oDom.createElement("AL1.1")
'    oAL1ChildLevel1.Text = CStr(lngCount)
'    oAL1Node.appendChild oAL1ChildLevel1
'
'    ' Allergy Type
'    If oConCept Is Nothing Then
'        Set oConCept = New GEMS_ConceptCtrl
'    End If
'    oConCept.uniqueID = AllergyRS.Fields("AllergyTypeNCID").Value
'    strName = oConCept.PrefRep("2000").Representation
'    strAllergyType = "MA"
'    If InStr(1, UCase(strName), "DRUG", vbTextCompare) <> 0 Then
'        strAllergyType = "DA"
'    End If
'    If InStr(1, UCase(strName), "FOOD", vbTextCompare) <> 0 Then
'        strAllergyType = "FA"
'    End If
'    If InStr(1, UCase(strName), "MISCELLANEOUS ALLERGY", vbTextCompare) <> 0 Then
'       strAllergyType = "MA"
'    End If
'    If InStr(1, UCase(strName), "MISCELLANEOUS CONTRAINDICATION", vbTextCompare) <> 0 Then
'       strAllergyType = "MC"
'    End If
'    Set oAL1ChildLevel1 = oDom.createElement("AL1.2")
'    oAL1ChildLevel1.Text = strAllergyType
'    oAL1Node.appendChild oAL1ChildLevel1
'
'    ' allergy
'    Set oAL1ChildLevel1 = oDom.createElement("AL1.3")
'        ' Allergy ICD-9 Code - default to 477.9
'        Set oAL1ChildLevel2 = oDom.createElement("CE.1")
'        oAL1ChildLevel2.Text = "477.9"
'        oAL1ChildLevel1.appendChild oAL1ChildLevel2
'        ' Allergen ICD-9 Description
'        Set oAL1ChildLevel2 = oDom.createElement("CE.2")
'        oAL1ChildLevel2.Text = "Allergy (Other)"
'        oAL1ChildLevel1.appendChild oAL1ChildLevel2
'        ' Identifies previous is ICD-9 - default to "I9"
'        Set oAL1ChildLevel2 = oDom.createElement("CE.3")
'        oAL1ChildLevel2.Text = "I9"
'        oAL1ChildLevel1.appendChild oAL1ChildLevel2
'        ' Allergen NCID
'        Set oAL1ChildLevel2 = oDom.createElement("CE.4")
'        oAL1ChildLevel2.Text = NullToNothing(AllergyRS.Fields("ALLERGENNCID").Value)
'        oAL1ChildLevel1.appendChild oAL1ChildLevel2
'        ' Allergy Description
'        Set oAL1ChildLevel2 = oDom.createElement("CE.5")
'        oAL1ChildLevel2.Text = NullToNothing(AllergyRS.Fields("Allergen").Value)
'        oAL1ChildLevel1.appendChild oAL1ChildLevel2
'        ' Identifies previous as NCID - default to "99NCI"
'        Set oAL1ChildLevel2 = oDom.createElement("CE.6")
'        oAL1ChildLevel2.Text = "99NCI"
'        oAL1ChildLevel1.appendChild oAL1ChildLevel2
'    oAL1Node.appendChild oAL1ChildLevel1
'
'    ' Allergy Reaction
'    Set oAL1ChildLevel1 = oDom.createElement("AL1.5.LST")
'    Do
'        Set oAL1ChildLevel2 = oDom.createElement("AL1.5")
'        If Not (ReactionRS Is Nothing) Then
'            If ReactionRS.EOF = False Then
'                ' Reaction NCID
'                Set oAL1ChildLevel3 = oDom.createElement("CE.1")
'                oAL1ChildLevel3.Text = "477.9"
'                oAL1ChildLevel2.appendChild oAL1ChildLevel3
'                ' Reaction Description
'                Set oAL1ChildLevel3 = oDom.createElement("CE.2")
'                oAL1ChildLevel3.Text = NullToNothing(GetNCIDDescription(ReactionRS.Fields("REACTIONNCID").Value))
'                oAL1ChildLevel2.appendChild oAL1ChildLevel3
'                ' Reaction NCID
'                Set oAL1ChildLevel3 = oDom.createElement("CE.3")
'                oAL1ChildLevel3.Text = "I9"
'                oAL1ChildLevel2.appendChild oAL1ChildLevel3
'                ' Reaction NCID
'                Set oAL1ChildLevel3 = oDom.createElement("CE.4")
'                oAL1ChildLevel3.Text = NullToNothing(ReactionRS.Fields("REACTIONNCID").Value)
'                oAL1ChildLevel2.appendChild oAL1ChildLevel3
'                ' Reaction Description
'                Set oAL1ChildLevel3 = oDom.createElement("CE.5")
'                oAL1ChildLevel3.Text = NullToNothing(ReactionRS.Fields("REACTION").Value) '  (GetNCIDDescription(ReactionRS.Fields("REACTIONNCID").Value))
'                oAL1ChildLevel2.appendChild oAL1ChildLevel3
'
'                Set oAL1ChildLevel3 = oDom.createElement("CE.6")
'                oAL1ChildLevel3.Text = "99NCI"
'                oAL1ChildLevel2.appendChild oAL1ChildLevel3
'
'                oAL1ChildLevel1.appendChild oAL1ChildLevel2
'                ReactionRS.MoveNext
'
'             Else 'ReactionRS.EOF = False
'                ' Reaction NCID
'                Set oAL1ChildLevel3 = oDom.createElement("CE.1")
'                oAL1ChildLevel3.Text = ""
'                oAL1ChildLevel2.appendChild oAL1ChildLevel3
'                ' Reaction Description
'                Set oAL1ChildLevel3 = oDom.createElement("CE.2")
'                oAL1ChildLevel3.Text = ""
'                oAL1ChildLevel2.appendChild oAL1ChildLevel3
'                ' Reaction Coding System
'                Set oAL1ChildLevel3 = oDom.createElement("CE.3")
'                oAL1ChildLevel3.Text = ""
'                oAL1ChildLevel2.appendChild oAL1ChildLevel3
'                ' Reaction Alternate Identifier
'                Set oAL1ChildLevel3 = oDom.createElement("CE.4")
'                oAL1ChildLevel3.Text = ""
'                oAL1ChildLevel2.appendChild oAL1ChildLevel3
'                ' Reaction Alternate Description
'                Set oAL1ChildLevel3 = oDom.createElement("CE.5")
'                oAL1ChildLevel3.Text = ""
'                oAL1ChildLevel2.appendChild oAL1ChildLevel3
'                ' Reaction Alternate Coding System
'                Set oAL1ChildLevel3 = oDom.createElement("CE.6")
'                oAL1ChildLevel3.Text = ""
'                oAL1ChildLevel2.appendChild oAL1ChildLevel3
'
'                oAL1ChildLevel1.appendChild oAL1ChildLevel2
'
'            End If 'ReactionRS.EOF = False
'
'        Else 'Not (ReactionRS Is Nothing)
'            ' Reaction NCID
'            Set oAL1ChildLevel3 = oDom.createElement("CE.1")
'            oAL1ChildLevel3.Text = ""
'            oAL1ChildLevel2.appendChild oAL1ChildLevel3
'            ' Reaction Description
'            Set oAL1ChildLevel3 = oDom.createElement("CE.2")
'            oAL1ChildLevel3.Text = ""
'            oAL1ChildLevel2.appendChild oAL1ChildLevel3
'            ' Reaction Coding System
'            Set oAL1ChildLevel3 = oDom.createElement("CE.3")
'            oAL1ChildLevel3.Text = ""
'            oAL1ChildLevel2.appendChild oAL1ChildLevel3
'            ' Reaction Alternate Identifier
'            Set oAL1ChildLevel3 = oDom.createElement("CE.4")
'            oAL1ChildLevel3.Text = ""
'            oAL1ChildLevel2.appendChild oAL1ChildLevel3
'            ' Reaction Alternate Description
'            Set oAL1ChildLevel3 = oDom.createElement("CE.5")
'            oAL1ChildLevel3.Text = ""
'            oAL1ChildLevel2.appendChild oAL1ChildLevel3
'            ' Reaction Alternate Coding System
'            Set oAL1ChildLevel3 = oDom.createElement("CE.6")
'            oAL1ChildLevel3.Text = ""
'            oAL1ChildLevel2.appendChild oAL1ChildLevel3
'
'            oAL1ChildLevel1.appendChild oAL1ChildLevel2
'
'        End If ''Not (ReactionRS Is Nothing)
'    Loop While ReactionRS.EOF = False
'
'    oAL1Node.appendChild oAL1ChildLevel1
'
'    Set oAL1ChildLevel1 = oDom.createElement("AL1.6")
'    If Not IsNull(AllergyRS.Fields("NOTEDDATE").Value) Then
'        oAL1ChildLevel1.Text = Format(AllergyRS.Fields("NOTEDDATE").Value, "YYYYMMDDHHNNSS")
'    End If
'    oAL1Node.appendChild oAL1ChildLevel1
'
'    Set ReturnNewAL1Node = oAL1Node
'
'Exit Function
'
'ErrTrap:
'   If gobjShared Is Nothing Then
'      Set gobjShared = New CWShared
'   End If
'   gobjShared.ShowVBError Err.Number, Err.Description, "ReturnNewAL1Node Function", "DMBE", vbExclamation
'   'Err.Clear
'
'End Function


Public Function ReturnNewAL1Node(ByVal lngCount As Long, ByVal rsALLERGIES As ADODB.Recordset, ByVal rsALLERGYREACTION As ADODB.Recordset) As MSXML.IXMLDOMElement

    Dim oDom As MSXML.DOMDocument
    Dim oAL1Node As MSXML.IXMLDOMElement
    Dim oAL1ChildLevel1 As MSXML.IXMLDOMElement
    Dim oAL1ChildLevel2 As MSXML.IXMLDOMElement
    Dim oAL1ChildLevel3 As MSXML.IXMLDOMElement
    Dim strAllergyType As String
    Dim strName As String
    
    On Error GoTo ErrTrap

    Set oDom = New MSXML.DOMDocument
    Set oAL1Node = oDom.createElement("AL1")
    
    ' Allergy Count
    Set oAL1ChildLevel1 = oDom.createElement("AL1.1")
    oAL1ChildLevel1.Text = CStr(lngCount)
    oAL1Node.appendChild oAL1ChildLevel1
    
    ' Allergy Type
    If oConCept Is Nothing Then
        Set oConCept = New GEMS_ConceptCtrl
    End If
    oConCept.UniqueID = rsALLERGIES.Fields("AllergyTypeNCID").Value
    strName = oConCept.PrefRep("2000").Representation
    strAllergyType = "MA"
    If InStr(1, UCase(strName), "DRUG", vbTextCompare) <> 0 Then
        strAllergyType = "DA"
    End If
    If InStr(1, UCase(strName), "FOOD", vbTextCompare) <> 0 Then
        strAllergyType = "FA"
    End If
    If InStr(1, UCase(strName), "MISCELLANEOUS ALLERGY", vbTextCompare) <> 0 Then
       strAllergyType = "MA"
    End If
    If InStr(1, UCase(strName), "MISCELLANEOUS CONTRAINDICATION", vbTextCompare) <> 0 Then
       strAllergyType = "MC"
    End If
    Set oAL1ChildLevel1 = oDom.createElement("AL1.2")
    oAL1ChildLevel1.Text = strAllergyType
    oAL1Node.appendChild oAL1ChildLevel1
    
    ' allergy
    Set oAL1ChildLevel1 = oDom.createElement("AL1.3")
        ' Allergy ICD-9 Code - default to 477.9
        Set oAL1ChildLevel2 = oDom.createElement("CE.1")
        oAL1ChildLevel2.Text = "477.9"
        oAL1ChildLevel1.appendChild oAL1ChildLevel2
        ' Allergen ICD-9 Description
        Set oAL1ChildLevel2 = oDom.createElement("CE.2")
        oAL1ChildLevel2.Text = "Allergy (Other)"
        oAL1ChildLevel1.appendChild oAL1ChildLevel2
        ' Identifies previous is ICD-9 - default to "I9"
        Set oAL1ChildLevel2 = oDom.createElement("CE.3")
        oAL1ChildLevel2.Text = "I9"
        oAL1ChildLevel1.appendChild oAL1ChildLevel2
        ' Allergen NCID
        Set oAL1ChildLevel2 = oDom.createElement("CE.4")
        oAL1ChildLevel2.Text = NullToNothing(rsALLERGIES.Fields("ALLERGENNCID").Value)
        oAL1ChildLevel1.appendChild oAL1ChildLevel2
        ' Allergy Description
        Set oAL1ChildLevel2 = oDom.createElement("CE.5")
        oAL1ChildLevel2.Text = NullToNothing(rsALLERGIES.Fields("Allergen").Value)
        oAL1ChildLevel1.appendChild oAL1ChildLevel2
        ' Identifies previous as NCID - default to "99NCI"
        Set oAL1ChildLevel2 = oDom.createElement("CE.6")
        oAL1ChildLevel2.Text = "99NCI"
        oAL1ChildLevel1.appendChild oAL1ChildLevel2
    oAL1Node.appendChild oAL1ChildLevel1
    
    ' Allergy Reaction
    Set oAL1ChildLevel1 = oDom.createElement("AL1.5.LST")
    
    Dim blnRsIsGood As Boolean

    If Not rsALLERGYREACTION Is Nothing Then
        If rsALLERGYREACTION.RecordCount > 0 Then
            blnRsIsGood = True
            rsALLERGYREACTION.MoveFirst
        End If
    End If
    
    If blnRsIsGood = False Then
        Set oAL1ChildLevel2 = oDom.createElement("AL1.5")
        ' Reaction NCID
            Set oAL1ChildLevel3 = oDom.createElement("CE.1")
            oAL1ChildLevel3.Text = ""
            oAL1ChildLevel2.appendChild oAL1ChildLevel3
            ' Reaction Description
            Set oAL1ChildLevel3 = oDom.createElement("CE.2")
            oAL1ChildLevel3.Text = ""
            oAL1ChildLevel2.appendChild oAL1ChildLevel3
            ' Reaction Coding System
            Set oAL1ChildLevel3 = oDom.createElement("CE.3")
            oAL1ChildLevel3.Text = ""
            oAL1ChildLevel2.appendChild oAL1ChildLevel3
            ' Reaction Alternate Identifier
            Set oAL1ChildLevel3 = oDom.createElement("CE.4")
            oAL1ChildLevel3.Text = ""
            oAL1ChildLevel2.appendChild oAL1ChildLevel3
            ' Reaction Alternate Description
            Set oAL1ChildLevel3 = oDom.createElement("CE.5")
            oAL1ChildLevel3.Text = ""
            oAL1ChildLevel2.appendChild oAL1ChildLevel3
            ' Reaction Alternate Coding System
            Set oAL1ChildLevel3 = oDom.createElement("CE.6")
            oAL1ChildLevel3.Text = ""
            oAL1ChildLevel2.appendChild oAL1ChildLevel3
            
        oAL1ChildLevel1.appendChild oAL1ChildLevel2
    Else
        Do Until rsALLERGYREACTION.EOF
            Set oAL1ChildLevel2 = oDom.createElement("AL1.5")
            ' Reaction NCID
                Set oAL1ChildLevel3 = oDom.createElement("CE.1")
                oAL1ChildLevel3.Text = "477.9"
                oAL1ChildLevel2.appendChild oAL1ChildLevel3
                ' Reaction Description
                Set oAL1ChildLevel3 = oDom.createElement("CE.2")
                oAL1ChildLevel3.Text = NullToNothing(GetNCIDDescription(rsALLERGYREACTION.Fields("REACTIONNCID").Value))
                oAL1ChildLevel2.appendChild oAL1ChildLevel3
                ' Reaction NCID
                Set oAL1ChildLevel3 = oDom.createElement("CE.3")
                oAL1ChildLevel3.Text = "I9"
                oAL1ChildLevel2.appendChild oAL1ChildLevel3
                ' Reaction NCID
                Set oAL1ChildLevel3 = oDom.createElement("CE.4")
                oAL1ChildLevel3.Text = NullToNothing(rsALLERGYREACTION.Fields("REACTIONNCID").Value)
                oAL1ChildLevel2.appendChild oAL1ChildLevel3
                ' Reaction Description
                Set oAL1ChildLevel3 = oDom.createElement("CE.5")
                oAL1ChildLevel3.Text = NullToNothing(rsALLERGYREACTION.Fields("REACTION").Value) '  (GetNCIDDescription(rsALLERGYREACTION.Fields("REACTIONNCID").Value))
                oAL1ChildLevel2.appendChild oAL1ChildLevel3
                
                Set oAL1ChildLevel3 = oDom.createElement("CE.6")
                oAL1ChildLevel3.Text = "99NCI"
                oAL1ChildLevel2.appendChild oAL1ChildLevel3
            oAL1ChildLevel1.appendChild oAL1ChildLevel2
            rsALLERGYREACTION.MoveNext
        Loop
    End If
    
    oAL1Node.appendChild oAL1ChildLevel1
    
    Set oAL1ChildLevel1 = oDom.createElement("AL1.6")
    If Not IsNull(rsALLERGIES.Fields("NOTEDDATE").Value) Then
        oAL1ChildLevel1.Text = Format(rsALLERGIES.Fields("NOTEDDATE").Value, "YYYYMMDDHHNNSS")
    End If
    oAL1Node.appendChild oAL1ChildLevel1
    
    Set ReturnNewAL1Node = oAL1Node
    
Exit Function

ErrTrap:
   If gobjShared Is Nothing Then
      Set gobjShared = New CWShared
   End If
   gobjShared.ShowVBError Err.Number, Err.Description, "ReturnNewAL1Node Function", "DMBE", vbExclamation
   'Err.Clear
   Resume Next 'There is no reason this procedure should not continue on to create xml

End Function
'<SCR #19356;   Developer: CC 05/02/2002 10:26 AM>


Public Function ReturnNewALGPV1Node(ByVal lngCount As Long) As MSXML.IXMLDOMElement
' Patient Visit Segment

' sample function output
'   <PV1>
'      <PV1.1>1</PV1.1>
'      <PV1.2>O</PV1.2>
'      <PV1.3>
'         <PL.1>Midwifery</PL.1>
'         <PL.4>
'            <HD.1>ADOLESCENT MEDICINE CLINIC</HD.1>
'         </PL.4>
'      </PV1.3>
'      <PV1.7.LST>
'         <PV1.7>
'            <XCN.1>9987656554</XCN.1>
'            <XCN.2>SHEN</XCN.2>
'            <XCN.3>JANE</XCN.3>
'            <XCN.4></XCN.4>
'         </PV1.7>
'      </PV1.7.LST>
'      <PV1.18>1</PV1.18>
'      <PV1.36>00</PV1.36>
'      <PV1.39>
'            <CE.1></CE.1>
'            <CE.2>ADOLESCENT MEDICINE CLINIC</CE.2>
'            <CE.3></CE.3>
'            <CE.4>1047017</CE.4>
'            <CE.5>ADOLESCENT MEDICINE CLINIC</CE.5>
'            <CE.6>99NCI</CE.6>
'      </PV1.39>
'      <PV1.44>8/8/01 3:48:50 PM</PV1.44>
'      <PV1.45></PV1.45>
'   </PV1>

    Dim oDom As MSXML.DOMDocument
    Dim oPV1Node As MSXML.IXMLDOMElement
    Dim oPV1ChildLevel1 As MSXML.IXMLDOMElement
    
    On Error GoTo ErrTrap

    Set oDom = New MSXML.DOMDocument
    Set oPV1Node = oDom.createElement("PV1")
    
    ' Segment Count
    Set oPV1ChildLevel1 = oDom.createElement("PV1.1")
    oPV1ChildLevel1.Text = CStr(lngCount)
    oPV1Node.appendChild oPV1ChildLevel1
    
    ' Patient Class - default to "O" - Outpatient
    Set oPV1ChildLevel1 = oDom.createElement("PV1.2")
    oPV1ChildLevel1.Text = "O"
    oPV1Node.appendChild oPV1ChildLevel1
    
    Set ReturnNewALGPV1Node = oPV1Node

Exit Function

ErrTrap:
   If gobjShared Is Nothing Then
      Set gobjShared = New CWShared
   End If
   gobjShared.ShowVBError Err.Number, Err.Description, "ReturnNewALGPV1Node Function", "DMBE", vbExclamation
   'Err.Clear
   Resume Next 'There is no reason this procedure should not continue on to create xml

End Function

Public Function ReturnTMIPDefinedZD1Node() As MSXML.IXMLDOMElement

' Additional Diagnosis Data Segment
' sample function output
'   <ZD1>
'      <ZD1.1>
'         <XON.1>AIR FORCE</XON.1>
'         <XON.2></XON.2>
'      </ZD1.1>
'      <ZD1.2>
'         <XON.1>The Unit Wizzard...1</XON.1>
'         <XON.2></XON.2>
'      </ZD1.2>
'      <ZD1.3>
'         <CE.1></CE.1>
'         <CE.2>TECHNICAL SERGEANT</CE.2>
'         <CE.3></CE.3>
'         <CE.4></CE.4>
'         <CE.5>TECHNICAL SERGEANT</CE.5>
'         <CE.6></CE.6>
'      </ZD1.3>
'      <ZD1.4>
'         <CE.1></CE.1>
'         <CE.2></CE.2>
'         <CE.3></CE.3>
'         <CE.4></CE.4>
'         <CE.5></CE.5>
'         <CE.6></CE.6>
'      </ZD1.4>
'      <ZD1.5>
'         <CE.1></CE.1>
'         <CE.2></CE.2>
'         <CE.3></CE.3>
'         <CE.4></CE.4>
'         <CE.5></CE.5>
'         <CE.6></CE.6>
'      </ZD1.5>
'      <ZD1.6>
'         <CE.1></CE.1>
'         <CE.2>False</CE.2>
'         <CE.3></CE.3>
'         <CE.4></CE.4>
'         <CE.5>False</CE.5>
'         <CE.6></CE.6>
'      </ZD1.6>
'      <ZD1.7>
'         <CE.1></CE.1>
'         <CE.2>F41 USAF FAM MBR AD</CE.2>
'         <CE.3></CE.3>
'         <CE.4></CE.4>
'         <CE.5>F41 USAF FAM MBR AD</CE.5>
'         <CE.6></CE.6>
'      </ZD1.7>
'   </ZD1>

    Dim oDom As MSXML.DOMDocument
    Dim oZD1Node As MSXML.IXMLDOMElement
    Dim oZD1ChildLevel1 As MSXML.IXMLDOMElement
    Dim oZD1ChildLevel2 As MSXML.IXMLDOMElement
    Dim strTemp As String
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrTrap
    
    
    If oDas Is Nothing Then
        Set oDas = oDatabase.CHCSII_DAS(Auto) 'New GEMS_DAS
    End If
        
    Set oDom = New MSXML.DOMDocument
    Set oZD1Node = oDom.createElement("ZD1")

'<SCR 23517 CC>
    ' Branch of Service - SCR 25147
    Set oZD1ChildLevel1 = oDom.createElement("ZD1.1")
        Set oZD1ChildLevel2 = oDom.createElement("XON.1")
        '<SCR 22766>
        strTemp = gobjPatient.BranchOfService
        If Not IsNull(strTemp) Then
            oZD1ChildLevel2.Text = strTemp
        End If
        oZD1ChildLevel1.appendChild oZD1ChildLevel2
'        Set oZD1ChildLevel2 = oDom.createElement("XON.3")
'        If Not IsNull(strTemp) Then
'            oZD1ChildLevel2.Text = GetCodeBranchOfService(Trim(strTemp)) SCR 25147
'        End If
'Greg, in case you are looking at this right now I commented out the below line because it is a duplicate of the above. Brian
'        oZD1ChildLevel1.appendChild oZD1ChildLevel2
    oZD1Node.appendChild oZD1ChildLevel1

    ' Home Unit <SCR 25146 CC>
    Set oZD1ChildLevel1 = oDom.createElement("ZD1.2")
        
        strTemp = NullToNothing(gobjPatient.CurrentPatient.UIC)
        
        Dim strUIC() As String '<SCR 25146 CC>
        
        If strTemp <> "" Then
            
            '<< Begin:SCR #25146;   Developer: Brian Mowbray 07/22/2002 12:47 PM
            'This was only setup to handle the the UIC if there is the UIC was in parenthesis somewhere in the text, eg (N1232)
            'Change made to handle any UIC
            
            strUIC = Split(strTemp, "(")
            If UBound(strUIC) = 1 Then
                Set oZD1ChildLevel2 = oDom.createElement("XON.1")
                oZD1ChildLevel2.Text = Trim(strUIC(0))
                oZD1ChildLevel1.appendChild oZD1ChildLevel2
                Set oZD1ChildLevel2 = oDom.createElement("XON.3")
                strUIC = Split(strUIC(1), ")")
                oZD1ChildLevel2.Text = Trim(strUIC(0))
                oZD1ChildLevel1.appendChild oZD1ChildLevel2
            Else
                'Added the "Else" to handle all other UIC's as part of SCR #25146
                Set oZD1ChildLevel2 = oDom.createElement("XON.1")
                oZD1ChildLevel2.Text = Trim(strTemp)
                oZD1ChildLevel1.appendChild oZD1ChildLevel2
            End If
            '>> End: SCR #25146;
        End If
        
        '<SCR 25146 CC>
        'Set oZD1ChildLevel2 = oDom.createElement("XON.1")
        'oZD1ChildLevel2.Text = gobjPatient.CurrentPatient.DeployUnit
        'oZD1ChildLevel1.appendChild oZD1ChildLevel2
        'Set oZD1ChildLevel2 = oDom.createElement("XON.3")
        'oZD1ChildLevel2.Text = ""   'Deploy Unit is a text field in MMI_ID
        'oZD1ChildLevel2.Text = gobjPatient.CurrentPatient.UIC '<SCR 25146 CC>
        'oZD1ChildLevel1.appendChild oZD1ChildLevel2
    oZD1Node.appendChild oZD1ChildLevel1

'</SCR 23517 CC>

    'Rank
    'Change Here
    Set oZD1ChildLevel1 = oDom.createElement("ZD1.3")
        Set oZD1ChildLevel2 = oDom.createElement("CE.1")
        oZD1ChildLevel2.Text = ""
        oZD1ChildLevel1.appendChild oZD1ChildLevel2
        Set oZD1ChildLevel2 = oDom.createElement("CE.2")
        oZD1ChildLevel2.Text = gobjPatient.Rank
        oZD1ChildLevel1.appendChild oZD1ChildLevel2
        Set oZD1ChildLevel2 = oDom.createElement("CE.3")
        oZD1ChildLevel2.Text = ""
        oZD1ChildLevel1.appendChild oZD1ChildLevel2
        Set oZD1ChildLevel2 = oDom.createElement("CE.4")
        'oZD1ChildLevel2.Text = "" 'Rank NCID
        
        strSQL = "SELECT VALUE_NCID FROM MMI_GENERICDATA WHERE "
        strSQL = strSQL & "UNIT_NUMBER = " & gobjPatient.unitNumber & " "
        strSQL = strSQL & "AND TYPE_NCID = " & NCID_RANK
        
        Set rsTemp = oDas.ExecuteSQL(strSQL)
        If Not rsTemp.EOF Then
            If Not IsNull(rsTemp.Fields(0).Value) Then
                oZD1ChildLevel2.Text = rsTemp.Fields(0).Value
            End If
        End If
        
        oZD1ChildLevel1.appendChild oZD1ChildLevel2
        Set oZD1ChildLevel2 = oDom.createElement("CE.5")
        oZD1ChildLevel2.Text = gobjPatient.Rank
        oZD1ChildLevel1.appendChild oZD1ChildLevel2
        Set oZD1ChildLevel2 = oDom.createElement("CE.6")
        oZD1ChildLevel2.Text = "99NCI"
        oZD1ChildLevel1.appendChild oZD1ChildLevel2
    oZD1Node.appendChild oZD1ChildLevel1
    
    ' Pay Grade - we don't have paygrade description
    Set oZD1ChildLevel1 = oDom.createElement("ZD1.4")
        Set oZD1ChildLevel2 = oDom.createElement("CE.1")
        oZD1ChildLevel2.Text = ""
        oZD1ChildLevel1.appendChild oZD1ChildLevel2
        Set oZD1ChildLevel2 = oDom.createElement("CE.2")
        
        '<< Begin:SCR #27266;   Developer: Brian Mowbray 09/06/2002 04:40 PM
        oZD1ChildLevel2.Text = ReturnSAMSPayGrade(gobjPatient.PayGrade & vbNullString)
        '>> End: SCR #27266;
        oZD1ChildLevel1.appendChild oZD1ChildLevel2
        
        Set oZD1ChildLevel2 = oDom.createElement("CE.3")
        oZD1ChildLevel2.Text = ""
        oZD1ChildLevel1.appendChild oZD1ChildLevel2
        
        Set oZD1ChildLevel2 = oDom.createElement("CE.4")
        oZD1ChildLevel2.Text = ""
        oZD1ChildLevel1.appendChild oZD1ChildLevel2
        
        Set oZD1ChildLevel2 = oDom.createElement("CE.5")
        '<< Begin:SCR #27266;   Developer: Brian Mowbray 09/06/2002 04:40 PM
        oZD1ChildLevel2.Text = ReturnSAMSPayGrade(gobjPatient.PayGrade & vbNullString)
        '>> End: SCR #27266;
        oZD1ChildLevel1.appendChild oZD1ChildLevel2
        
        Set oZD1ChildLevel2 = oDom.createElement("CE.6")
        oZD1ChildLevel2.Text = ""
        oZD1ChildLevel1.appendChild oZD1ChildLevel2
    oZD1Node.appendChild oZD1ChildLevel1
    
    ' Ocupational Specialty
    Set oZD1ChildLevel1 = oDom.createElement("ZD1.5")
    
        Set oZD1ChildLevel2 = oDom.createElement("CE.1")
        oZD1ChildLevel2.Text = ""
        oZD1ChildLevel1.appendChild oZD1ChildLevel2
        Set oZD1ChildLevel2 = oDom.createElement("CE.2")
        oZD1ChildLevel2.Text = ""
        oZD1ChildLevel1.appendChild oZD1ChildLevel2
        Set oZD1ChildLevel2 = oDom.createElement("CE.3")
        oZD1ChildLevel2.Text = ""
        oZD1ChildLevel1.appendChild oZD1ChildLevel2
        Set oZD1ChildLevel2 = oDom.createElement("CE.4")
        oZD1ChildLevel2.Text = ""
        oZD1ChildLevel1.appendChild oZD1ChildLevel2
        Set oZD1ChildLevel2 = oDom.createElement("CE.5")
        oZD1ChildLevel2.Text = ""
        oZD1ChildLevel1.appendChild oZD1ChildLevel2
        Set oZD1ChildLevel2 = oDom.createElement("CE.6")
        oZD1ChildLevel2.Text = ""
        oZD1ChildLevel1.appendChild oZD1ChildLevel2
    oZD1Node.appendChild oZD1ChildLevel1
    
    ' Flight Status - True or False
    Set oZD1ChildLevel1 = oDom.createElement("ZD1.6")
        Set oZD1ChildLevel2 = oDom.createElement("CE.1")
        oZD1ChildLevel2.Text = ""
        oZD1ChildLevel1.appendChild oZD1ChildLevel2
        
        Set oZD1ChildLevel2 = oDom.createElement("CE.2")
        strTemp = NullToNothing(gobjPatient.CurrentPatient.SWSFlying)
        If strTemp <> "" Then
            If strTemp = "Y" Then
                oZD1ChildLevel2.Text = "True"
            Else
                oZD1ChildLevel2.Text = "False"
            End If
        End If
        'For Each oStatus In gobjPatient.CurrentPatient.SWSFlying
        '    If oStatus.type_ncid = NCID_FLYING_STATUS Then
        '        If oStatus.value_ncid = NCID_YES Then
        '            oZD1ChildLevel2.Text = "True"
        '        Else
        '            oZD1ChildLevel2.Text = "False"
        '        End If
        '        Exit For
        '    End If
        'Next
        oZD1ChildLevel1.appendChild oZD1ChildLevel2
        
        Set oZD1ChildLevel2 = oDom.createElement("CE.3")
        oZD1ChildLevel2.Text = ""
        oZD1ChildLevel1.appendChild oZD1ChildLevel2
        Set oZD1ChildLevel2 = oDom.createElement("CE.4")
        oZD1ChildLevel2.Text = ""
        oZD1ChildLevel1.appendChild oZD1ChildLevel2
        
        Set oZD1ChildLevel2 = oDom.createElement("CE.5")
        If IsNull(gobjPatient.CurrentPatient.SWSFlying) Then
            oZD1ChildLevel2.Text = "False"
        Else
            If gobjPatient.CurrentPatient.SWSFlying = "Y" Then
                oZD1ChildLevel2.Text = "True"
            Else
                oZD1ChildLevel2.Text = "False"
            End If
        End If
'        For Each oStatus In gobjPatient.SpecialWorkStatus
'            If oStatus.type_ncid = NCID_FLYING_STATUS Then
'                If oStatus.value_ncid = NCID_YES Then
'                    oZD1ChildLevel2.Text = "True"
'                Else
'                    oZD1ChildLevel2.Text = "False"
'                End If
'                Exit For
'            End If
'        Next
        oZD1ChildLevel1.appendChild oZD1ChildLevel2
        
        Set oZD1ChildLevel2 = oDom.createElement("CE.6")
        oZD1ChildLevel2.Text = ""
        oZD1ChildLevel1.appendChild oZD1ChildLevel2
    oZD1Node.appendChild oZD1ChildLevel1
    
    ' Patient Category
    Set oZD1ChildLevel1 = oDom.createElement("ZD1.7")
        Set oZD1ChildLevel2 = oDom.createElement("CE.1")
        oZD1ChildLevel2.Text = ""
        oZD1ChildLevel1.appendChild oZD1ChildLevel2
        Set oZD1ChildLevel2 = oDom.createElement("CE.2")
        oZD1ChildLevel2.Text = gobjPatient.PatientCategory
        oZD1ChildLevel1.appendChild oZD1ChildLevel2
        Set oZD1ChildLevel2 = oDom.createElement("CE.3")
        oZD1ChildLevel2.Text = ""
        oZD1ChildLevel1.appendChild oZD1ChildLevel2
        Set oZD1ChildLevel2 = oDom.createElement("CE.4")
        'oZD1ChildLevel2.Text = ""
        
        strSQL = "SELECT VALUE_NCID FROM MMI_GENERICDATA WHERE "
        strSQL = strSQL & "UNIT_NUMBER = " & gobjPatient.unitNumber & " "
        strSQL = strSQL & "AND TYPE_NCID = " & NCID_PATIENT_CATEGORY
        
        Set rsTemp = oDas.ExecuteSQL(strSQL)
        If Not rsTemp.EOF Then
            If Not IsNull(rsTemp.Fields(0).Value) Then
                oZD1ChildLevel2.Text = rsTemp.Fields(0).Value
            End If
        End If
    
        oZD1ChildLevel1.appendChild oZD1ChildLevel2
        Set oZD1ChildLevel2 = oDom.createElement("CE.5")
        oZD1ChildLevel2.Text = gobjPatient.PatientCategory
        oZD1ChildLevel1.appendChild oZD1ChildLevel2
        Set oZD1ChildLevel2 = oDom.createElement("CE.6")
        oZD1ChildLevel2.Text = "99NCI"
        oZD1ChildLevel1.appendChild oZD1ChildLevel2
    oZD1Node.appendChild oZD1ChildLevel1
     
    Set ReturnTMIPDefinedZD1Node = oZD1Node
    
Exit Function

ErrTrap:
   If gobjShared Is Nothing Then
      Set gobjShared = New CWShared
   End If
   gobjShared.ShowVBError Err.Number, Err.Description, "ReturnTMIPDefinedZD1Node Function", "DMBE", vbExclamation
   Resume Next 'There is no reason this procedure should not continue to create xml if it errors
End Function
    
'<< SCR #27266;   Developer: Brian Mowbray 09/06/2002 04:41 PM
Private Function ReturnSAMSPayGrade(sCHCSIIPaygrade As String) As String
  Dim sTemp As String
  On Error Resume Next
  
  If geDestination = eDM_EIC Then
    
      Select Case sCHCSIIPaygrade
        Case "CD"
          sTemp = "CAD"
        Case "E1"
          sTemp = "E01"
        Case "E2"
          sTemp = "E02"
        Case "E3"
          sTemp = "E03"
        Case "E4"
          sTemp = "E04"
        Case "E5"
          sTemp = "E05"
        Case "E6"
          sTemp = "E06"
        Case "E7"
          sTemp = "E07"
        Case "E8"
          sTemp = "E08"
        Case "E9"
          sTemp = "E09"
        Case "O1"
          sTemp = "O01"
        Case "O2"
          sTemp = "O02"
        Case "O3"
          sTemp = "O03"
        Case "O4"
          sTemp = "O04"
        Case "O5"
          sTemp = "O05"
        Case "O6"
          sTemp = "O06"
        Case "O7"
          sTemp = "O07"
        Case "O8"
          sTemp = "O08"
        Case "O9"
          sTemp = "O09"
        Case "W1"
          sTemp = "W01"
        Case "W2"
          sTemp = "W02"
        Case "W3"
          sTemp = "W03"
        Case "W4"
          sTemp = "W04"
        Case "W5"
          sTemp = "W05"
        Case Else
          sTemp = sCHCSIIPaygrade
      End Select
  Else
    sTemp = sCHCSIIPaygrade
  End If
  
  ReturnSAMSPayGrade = sTemp

End Function


Private Function GetCodeBranchOfService(ByVal strBranch As String) As String
    
    Dim strReturn As String
    
    Select Case UCase(strBranch)
        Case "ARMY"
            strReturn = "A"
        Case "NOAA"
            strReturn = "B"
        Case "COAST GUARD"
            strReturn = "C"
        Case "AIR FORCE"
            strReturn = "F"
        Case "DOD CONTRACTOR OR CIVILIAN"
            strReturn = "K"
        Case "MARINES", "MARINE"
            strReturn = "M" '2/19/02 - Walker - Modified to plural
        Case "NAVY"
            strReturn = "N"
        Case "USPHS"
            strReturn = "P"
        Case "NATO RECIP AGREE"
            strReturn = "R"
        Case "FOREIGN MILITARY"
            strReturn = "X"
        Case "UNKNOWN"
            strReturn = "U"
        Case Else
            strReturn = "U"
    End Select

    GetCodeBranchOfService = strReturn

End Function

Public Function ReturnNewZPTNode(ByVal lngCount As Long, ByVal Mmi_Inst_NameRS As ADODB.Recordset) As MSXML.IXMLDOMElement
' Patient Institution Segment

' sample function output
'   <ZPT>
'      <ZPT.1>Count</ZPT.1>
'      <ZPT.2>Unit Number</ZPT.2>
'      <ZPT.3>Name Index</ZPT.3>
'      <ZPT.4>Name</ZPT.4>
'      <ZPT.5>Create Audit Number</ZPT.5>
'      <ZPT.6>Row Update Audit Number</ZPT.6>
'      <ZPT.7>Name Type</ZPT.7>
'      <ZPT.8>Created By</ZPT.8>
'      <ZPT.9>Created On</ZPT.9>
'      <ZPT.10>Updated By</ZPT.10>
'      <ZPT.11>Updated On</ZPT.11>
'   </ZPT>

    Dim oDom As MSXML.DOMDocument
    Dim oZPTNode As MSXML.IXMLDOMElement
    Dim oZPTChildLevel1 As MSXML.IXMLDOMElement
    Dim oZPTChildLevel2 As MSXML.IXMLDOMElement
    Dim strTemp As String
    
    On Error GoTo ErrTrap
    
    Set oDom = New MSXML.DOMDocument
    Set oZPTNode = oDom.createElement("ZPT")
    
    ' Set ID - Count
    Set oZPTChildLevel1 = oDom.createElement("ZPT.1")
    oZPTChildLevel1.Text = CStr(lngCount)
    oZPTNode.appendChild oZPTChildLevel1
    
    ' Unit Number - Table MMI_Inst_Name:Unit_Number
    Set oZPTChildLevel1 = oDom.createElement("ZPT.2")
    oZPTChildLevel1.Text = NullToNothing(Mmi_Inst_NameRS.Fields("Unit_Number").Value)
    oZPTNode.appendChild oZPTChildLevel1
    
    ' Name Index - Table MMI_Inst_Name:Name_Index
    Set oZPTChildLevel1 = oDom.createElement("ZPT.3")
    oZPTChildLevel1.Text = NullToNothing(Mmi_Inst_NameRS.Fields("Name_Index").Value)
    oZPTNode.appendChild oZPTChildLevel1
    
    ' Name - Table MMI_Inst_Name:Name
    Set oZPTChildLevel1 = oDom.createElement("ZPT.4")
    oZPTChildLevel1.Text = NullToNothing(Mmi_Inst_NameRS.Fields("Name").Value)
    oZPTNode.appendChild oZPTChildLevel1
    
    ' Create Audit Number - Table MMI_Inst_Name:Create_Audit_Num
    Set oZPTChildLevel1 = oDom.createElement("ZPT.5")
    oZPTChildLevel1.Text = NullToNothing(Mmi_Inst_NameRS.Fields("Create_Audit_Num").Value)
    oZPTNode.appendChild oZPTChildLevel1
    
    ' Row Update Audit Number - Table MMI_Inst_Name:Row_Update_Audit_Num
    Set oZPTChildLevel1 = oDom.createElement("ZPT.6")
    oZPTChildLevel1.Text = NullToNothing(Mmi_Inst_NameRS.Fields("Row_Update_Audit_Num").Value)
    oZPTNode.appendChild oZPTChildLevel1
    
    ' Name Type - Table MMI_Inst_Name:Name_Type_CID
    Set oZPTChildLevel1 = oDom.createElement("ZPT.7")
    'If Not IsNull(Mmi_Inst_NameRS.Fields("Name_Type_CID").Value) Then
    '    oZPTChildLevel1.Text = GetNCIDDescription(Mmi_Inst_NameRS.Fields("Name_Type_CID").Value)
    'End If
    strTemp = Mmi_Inst_NameRS.Fields("Name_Type_CID").Value
        Set oZPTChildLevel2 = oDom.createElement("CE.1")
        oZPTChildLevel1.appendChild oZPTChildLevel2
        Set oZPTChildLevel2 = oDom.createElement("CE.2")
        If strTemp <> "" Then
            oZPTChildLevel2.Text = GetNCIDDescription(CLng(strTemp))
        End If
        oZPTChildLevel1.appendChild oZPTChildLevel2
        Set oZPTChildLevel2 = oDom.createElement("CE.3")
        oZPTChildLevel1.appendChild oZPTChildLevel2
        Set oZPTChildLevel2 = oDom.createElement("CE.4")
        oZPTChildLevel1.appendChild oZPTChildLevel2
        Set oZPTChildLevel2 = oDom.createElement("CE.5")
        oZPTChildLevel2.Text = strTemp
        oZPTChildLevel1.appendChild oZPTChildLevel2
        Set oZPTChildLevel2 = oDom.createElement("CE.6")
        oZPTChildLevel2.Text = "99NCI"
        oZPTChildLevel1.appendChild oZPTChildLevel2
    oZPTNode.appendChild oZPTChildLevel1
    
    ' Created By - Table MMI_Inst_Name:CreatedBy
    Set oZPTChildLevel1 = oDom.createElement("ZPT.8")
    oZPTChildLevel1.Text = "Created By"
    oZPTNode.appendChild oZPTChildLevel1
    
    ' Created On - Table MMI_Inst_Name:CreatedOn
    Set oZPTChildLevel1 = oDom.createElement("ZPT.9")
    oZPTChildLevel1.Text = "Created On"
    oZPTNode.appendChild oZPTChildLevel1
    
    ' Updated By - Table MMI_Inst_Name:UpdatedBy
    Set oZPTChildLevel1 = oDom.createElement("ZPT.10")
    oZPTChildLevel1.Text = "Updated By"
    oZPTNode.appendChild oZPTChildLevel1
    
    ' Updated On - Table MMI_Inst_Name:UpdatedOn
    Set oZPTChildLevel1 = oDom.createElement("ZPT.11")
    oZPTChildLevel1.Text = "Updated On"
    oZPTNode.appendChild oZPTChildLevel1
        
    Set ReturnNewZPTNode = oZPTNode

Exit Function

ErrTrap:
   If gobjShared Is Nothing Then
      Set gobjShared = New CWShared
   End If
   gobjShared.ShowVBError Err.Number, Err.Description, "ReturnNewZPTNode Function", "DMBE", vbExclamation
   Resume Next 'There is no reason this procedure should not continue to create xml if it errors
End Function


'<SCR 22707 CC - added a parameter>
Public Function ReturnNewZPINode(ByVal Mmi_IDRS As ADODB.Recordset, ByVal Mmi_NameRS As ADODB.Recordset, ByVal Mmi_SiteRS As ADODB.Recordset, ByVal rsMMI_GENERIC_ID As ADODB.Recordset, ByVal rsMMI_ADDRESS As ADODB.Recordset) As MSXML.IXMLDOMElement
    ' Additional Patient Data Segment

' sample funcion output
'   <ZPI>
'      <ZPI.1>Reference</ZPI.1>
'      <ZPI.2>Correct SSN</ZPI.2>
'      <ZPI.3>SSN Missing Reason</ZPI.3>
'      <ZPI.4>Govt Health Number</ZPI.4>
'      <ZPI.5>Passport Number</ZPI.5>
'      <ZPI.6>DOB Est Flag</ZPI.6>
'      <ZPI.7>Death Date/Time</ZPI.7>
'      <ZPI.8>Maiden Name</ZPI.8>
'      <ZPI.9>Patient Flag</ZPI.9>
'      <ZPI.10>Guarantor Flag</ZPI.10>
'      <ZPI.11>Expired Flag</ZPI.11>
'      <ZPI.12>Organ Donor Flag</ZPI.12>
'      <ZPI.13>Adoption Flag</ZPI.13>
'      <ZPI.14>Unidentified Flag</ZPI.14>
'      <ZPI.15>Deactivate Flag</ZPI.15>
'      <ZPI.16>Previous Link Error Flag</ZPI.16>
'      <ZPI.17>Institutional Flag</ZPI.17>
'      <ZPI.18>Birth Date Resolution</ZPI.18>
'      <ZPI.19>Death Date Resolution</ZPI.19>
'      <ZPI.20>Deploy Location</ZPI.20>
'      <ZPI.21>Deploy Unit</ZPI.21>
'      <ZPI.22>DNBI</ZPI.22>
'      <ZPI.23>Created in GEMS</ZPI.23>
'      <ZPI.24>Name Index</ZPI.24>
'      <ZPI.25>Last Name2</ZPI.25>
'      <ZPI.26>Middle Name2</ZPI.26>
'      <ZPI.27>First Name Soundex</ZPI.27>
'      <ZPI.28>Last Name Soundex</ZPI.28>
'      <ZPI.29>Radiology Number</ZPI.29>
'      <ZPI.30>MMI_GENERIC_ID
'      <ZPI.31>MMI_ADDRESS
'      <ZPI.32>Foreign_Id_Flag
'   </ZPI>

    Dim oDom As MSXML.DOMDocument
    Dim oZPINode As MSXML.IXMLDOMElement
    Dim oZPIChildLevel1 As MSXML.IXMLDOMElement
    Dim oZPIChildLevel2 As MSXML.IXMLDOMElement
    Dim oZPIChildLevel3 As MSXML.IXMLDOMElement
    Dim strTemp As String
    
    On Error GoTo ErrTrap
    
    Set oDom = New MSXML.DOMDocument
    Set oZPINode = oDom.createElement("ZPI")
    
    ' Reference - Table MMI_ID:Reference
    Set oZPIChildLevel1 = oDom.createElement("ZPI.1")
    oZPIChildLevel1.Text = NullToNothing(Mmi_IDRS.Fields("Reference").Value)
    oZPINode.appendChild oZPIChildLevel1
    
    ' Correct SSN - Table MMI_ID:Correct_SSN
    Set oZPIChildLevel1 = oDom.createElement("ZPI.2")
    oZPIChildLevel1.Text = NullToNothing(Mmi_IDRS.Fields("Correct_SSN").Value)
    oZPINode.appendChild oZPIChildLevel1
    
    ' SSN Missing Reason - Table MMI_ID:SSN_Missing_Reason_CID
    Set oZPIChildLevel1 = oDom.createElement("ZPI.3")
    If Not NullToNothing(Mmi_IDRS.Fields("SSN_Missing_Reason_CID").Value) = "" Then
        oZPIChildLevel1.Text = GetNCIDDescription(CLng(strTemp))
    End If
    oZPINode.appendChild oZPIChildLevel1
    
    ' Govt Health Number - Table MMI_ID:Govt_Health_Number
    Set oZPIChildLevel1 = oDom.createElement("ZPI.4")
    oZPIChildLevel1.Text = NullToNothing(Mmi_IDRS.Fields("Govt_Health_Number").Value)
    oZPINode.appendChild oZPIChildLevel1
    
    ' Passport Number - Table MMI_ID:Passport_Number
    Set oZPIChildLevel1 = oDom.createElement("ZPI.5")
    oZPIChildLevel1.Text = NullToNothing(Mmi_IDRS.Fields("Passport_Number").Value)
    oZPINode.appendChild oZPIChildLevel1
    
    ' DOB Est Flag - Table MMI_ID:DOB_Est_Flag
    Set oZPIChildLevel1 = oDom.createElement("ZPI.6")
    oZPIChildLevel1.Text = NullToNothing(Mmi_IDRS.Fields("DOB_Est_Flag").Value)
    oZPINode.appendChild oZPIChildLevel1
    
    ' Death Date/Time - Table MMI_ID:Death_Date
    Set oZPIChildLevel1 = oDom.createElement("ZPI.7")
    If Not IsNull(Mmi_IDRS.Fields("Death_Date").Value) Then
        oZPIChildLevel1.Text = Format(Mmi_IDRS.Fields("Death_Date").Value, "YYYYMMDDHHNNSS")
    End If
    oZPINode.appendChild oZPIChildLevel1
    
    ' Maiden Name - Table MMI_ID:Maiden_Name
    Set oZPIChildLevel1 = oDom.createElement("ZPI.8")
    oZPIChildLevel1.Text = NullToNothing(Mmi_IDRS.Fields("Maiden_Name").Value)
    oZPINode.appendChild oZPIChildLevel1
    
    ' Patient Flag - Table MMI_ID:Patient_Flag
    Set oZPIChildLevel1 = oDom.createElement("ZPI.9")
    oZPIChildLevel1.Text = NullToNothing(Mmi_IDRS.Fields("Patient_Flag").Value)
    oZPINode.appendChild oZPIChildLevel1
    
    ' Guarantor Flag - Table MMI_ID:Guarantor_Flag
    Set oZPIChildLevel1 = oDom.createElement("ZPI.10")
    oZPIChildLevel1.Text = NullToNothing(Mmi_IDRS.Fields("Guarantor_Flag").Value)
    oZPINode.appendChild oZPIChildLevel1
    
    ' Expired Flag - Table MMI_ID:Expired_Flag
    Set oZPIChildLevel1 = oDom.createElement("ZPI.11")
    oZPIChildLevel1.Text = NullToNothing(Mmi_IDRS.Fields("Expired_Flag").Value)
    oZPINode.appendChild oZPIChildLevel1
    
    ' Organ Donor Flag - Table MMI_ID:Organ_Donor_Flag
    Set oZPIChildLevel1 = oDom.createElement("ZPI.12")
    oZPIChildLevel1.Text = NullToNothing(Mmi_IDRS.Fields("Organ_Donor_Flag").Value)
    oZPINode.appendChild oZPIChildLevel1
    
    ' Adoption Flag - Table MMI_ID:Adoption_Flag
    Set oZPIChildLevel1 = oDom.createElement("ZPI.13")
    oZPIChildLevel1.Text = NullToNothing(Mmi_IDRS.Fields("Adoption_Flag").Value)
    oZPINode.appendChild oZPIChildLevel1
    
    ' Unidentified Flag - Table MMI_ID:Unidentified_Flag
    Set oZPIChildLevel1 = oDom.createElement("ZPI.14")
    oZPIChildLevel1.Text = NullToNothing(Mmi_IDRS.Fields("Unidentified_Flag").Value)
    oZPINode.appendChild oZPIChildLevel1
    
    ' Deactivate Flag - Table MMI_ID:Deactivate_Flag
    Set oZPIChildLevel1 = oDom.createElement("ZPI.15")
    oZPIChildLevel1.Text = NullToNothing(Mmi_IDRS.Fields("Deactivate_Flag").Value)
    oZPINode.appendChild oZPIChildLevel1
    
    ' Previous Link Error Flag - Table MMI_ID:Previous_Link_Error_Flag
    Set oZPIChildLevel1 = oDom.createElement("ZPI.16")
    oZPIChildLevel1.Text = NullToNothing(Mmi_IDRS.Fields("Previous_Link_Error_Flag").Value)
    oZPINode.appendChild oZPIChildLevel1
    
    ' Institutional Flag - Table MMI_ID:Institutional_Flag
    Set oZPIChildLevel1 = oDom.createElement("ZPI.17")
    oZPIChildLevel1.Text = NullToNothing(Mmi_IDRS.Fields("Institutional_Flag").Value)
    oZPINode.appendChild oZPIChildLevel1
    
    ' Birth Date Resolution - Table MMI_ID:Birth_Date_Resolution
    Set oZPIChildLevel1 = oDom.createElement("ZPI.18")
    oZPIChildLevel1.Text = NullToNothing(Mmi_IDRS.Fields("Birth_Date_Resolution").Value)
    oZPINode.appendChild oZPIChildLevel1
    
    ' Death Date Resolution - Table MMI_ID:Death_Date_Resolution
    Set oZPIChildLevel1 = oDom.createElement("ZPI.19")
    oZPIChildLevel1.Text = NullToNothing(Mmi_IDRS.Fields("Death_Date_Resolution").Value)
    oZPINode.appendChild oZPIChildLevel1
    
    ' Deploy Location - Table MMI_ID:DeployLocation
    Set oZPIChildLevel1 = oDom.createElement("ZPI.20")
    oZPIChildLevel1.Text = NullToNothing(Mmi_IDRS.Fields("DeployLocation").Value)
    oZPINode.appendChild oZPIChildLevel1
    
    ' Deploy Unit - Table MMI_ID:DeployUnit
    Set oZPIChildLevel1 = oDom.createElement("ZPI.21")
    oZPIChildLevel1.Text = NullToNothing(Mmi_IDRS.Fields("DeployUnit").Value)
    oZPINode.appendChild oZPIChildLevel1
    
    ' DNBI - Table MMI_ID:DNBI
    Set oZPIChildLevel1 = oDom.createElement("ZPI.22")
    oZPIChildLevel1.Text = NullToNothing(Mmi_IDRS.Fields("DNBI").Value)
    oZPINode.appendChild oZPIChildLevel1
    
    ' Created in GEMS - Table MMI_ID:CreatedInGEMS
    Set oZPIChildLevel1 = oDom.createElement("ZPI.23")
    oZPIChildLevel1.Text = NullToNothing(Mmi_IDRS.Fields("CreatedInGEMS").Value)
    oZPINode.appendChild oZPIChildLevel1
    
    ' Name Index - Table MMI_Name:Name_Index
    Set oZPIChildLevel1 = oDom.createElement("ZPI.24")
    oZPIChildLevel1.Text = NullToNothing(Mmi_NameRS.Fields("Name_Index").Value)
    oZPINode.appendChild oZPIChildLevel1
    
    ' Last Name2 - Table MMI_Name:Last_Name_2
    Set oZPIChildLevel1 = oDom.createElement("ZPI.25")
    oZPIChildLevel1.Text = NullToNothing(Mmi_NameRS.Fields("Last_Name_2").Value)
    oZPINode.appendChild oZPIChildLevel1
    
    ' Middle Name2 - Table MMI_Name:Middle_Name_2
    Set oZPIChildLevel1 = oDom.createElement("ZPI.26")
    oZPIChildLevel1.Text = NullToNothing(Mmi_NameRS.Fields("Middle_Name_2").Value)
    oZPINode.appendChild oZPIChildLevel1
    
    ' First Name Soundex - Table MMI_Name:First_Name_Soundex
    Set oZPIChildLevel1 = oDom.createElement("ZPI.27")
    oZPIChildLevel1.Text = NullToNothing(Mmi_NameRS.Fields("First_Name_Soundex").Value)
    oZPINode.appendChild oZPIChildLevel1
    
    ' Last Name Soundex - Table MMI_Name:Last_Name_Soundex
    Set oZPIChildLevel1 = oDom.createElement("ZPI.28")
    oZPIChildLevel1.Text = NullToNothing(Mmi_NameRS.Fields("Last_Name_Soundex").Value)
    oZPINode.appendChild oZPIChildLevel1
    
    ' Radiology Number - Table MMI_Site:Radiology_Number
    Set oZPIChildLevel1 = oDom.createElement("ZPI.29")
    oZPIChildLevel1.Text = NullToNothing(Mmi_SiteRS.Fields("Radiology_Number").Value)
    oZPINode.appendChild oZPIChildLevel1
    
    '<SCR 22707 CC>
    Set oZPIChildLevel1 = oDom.createElement("ZPI.30.LST")
    If Not rsMMI_GENERIC_ID Is Nothing Then
        If rsMMI_GENERIC_ID.RecordCount > 0 Then
            rsMMI_GENERIC_ID.MoveFirst
            Do Until rsMMI_GENERIC_ID.EOF
                Set oZPIChildLevel2 = oDom.createElement("ZPI.30")
                    Set oZPIChildLevel3 = oDom.createElement("CE.1")
                    oZPIChildLevel3.Text = NullToNothing(rsMMI_GENERIC_ID.Fields("ID_TYPE_NCID").Value)
                    oZPIChildLevel2.appendChild oZPIChildLevel3
                    Set oZPIChildLevel3 = oDom.createElement("CE.2")
                    oZPIChildLevel3.Text = NullToNothing(rsMMI_GENERIC_ID.Fields("ID_VALUE").Value)
                    oZPIChildLevel2.appendChild oZPIChildLevel3
                    Set oZPIChildLevel3 = oDom.createElement("CE.3")
                    oZPIChildLevel3.Text = NullToNothing(rsMMI_GENERIC_ID.Fields("UNIT_NUMBER").Value)
                    oZPIChildLevel2.appendChild oZPIChildLevel3
                oZPIChildLevel1.appendChild oZPIChildLevel2
                rsMMI_GENERIC_ID.MoveNext
            Loop
        End If
    End If
    oZPINode.appendChild oZPIChildLevel1
    '</SCR 22707 CC>
    
    Set oZPIChildLevel1 = oDom.createElement("ZPI.31")
    
    If Not rsMMI_ADDRESS Is Nothing Then
        If Not rsMMI_ADDRESS.RecordCount = 0 Then
            rsMMI_ADDRESS.MoveFirst
            Do Until rsMMI_ADDRESS.EOF
                If rsMMI_ADDRESS.Fields("ADDRESS_TYPE_CID").Value = NCID_WORK_ADDRESS Then
                    ' Begin:SCR #26004;   Developer: MAG 07/24/2002 03:16 PM
                    Set oZPIChildLevel2 = oDom.createElement("XAD.1")
                    oZPIChildLevel2.Text = NullToNothing(rsMMI_ADDRESS.Fields("STREET_1").Value)
                    oZPIChildLevel1.appendChild oZPIChildLevel2
                    Set oZPIChildLevel2 = oDom.createElement("XAD.2")
                    oZPIChildLevel2.Text = NullToNothing(rsMMI_ADDRESS.Fields("STREET_2").Value)
                    oZPIChildLevel1.appendChild oZPIChildLevel2
                    Set oZPIChildLevel2 = oDom.createElement("XAD.3")
                    oZPIChildLevel2.Text = NullToNothing(rsMMI_ADDRESS.Fields("CITY").Value)
                    oZPIChildLevel1.appendChild oZPIChildLevel2
                    Set oZPIChildLevel2 = oDom.createElement("XAD.4")
                    oZPIChildLevel2.Text = NullToNothing(rsMMI_ADDRESS.Fields("STATE").Value)
                    oZPIChildLevel1.appendChild oZPIChildLevel2
                    Set oZPIChildLevel2 = oDom.createElement("XAD.5")
                    oZPIChildLevel2.Text = NullToNothing(rsMMI_ADDRESS.Fields("POST_CODE").Value)
                    oZPIChildLevel1.appendChild oZPIChildLevel2
                    Set oZPIChildLevel2 = oDom.createElement("XAD.6")
                    oZPIChildLevel2.Text = NullToNothing(rsMMI_ADDRESS.Fields("COUNTRY").Value)
                    oZPIChildLevel1.appendChild oZPIChildLevel2
                    Set oZPIChildLevel2 = oDom.createElement("XAD.9")
                    oZPIChildLevel2.Text = NullToNothing(rsMMI_ADDRESS.Fields("COUNTY").Value)
                    oZPIChildLevel1.appendChild oZPIChildLevel2
                    ' End: SCR #26004;
                    
                    Exit Do
                End If
                rsMMI_ADDRESS.MoveNext
            Loop
        End If
    End If
    
    oZPINode.appendChild oZPIChildLevel1
    
    ' Begin:SCR #44724;   Developer: Mag 11/03/2003 08:40 PM
    Set oZPIChildLevel1 = oDom.createElement("ZPI.32")
    oZPIChildLevel1.Text = NullToNothing(Mmi_IDRS.Fields("Foreign_ID_Flag").Value)
    oZPINode.appendChild oZPIChildLevel1
    ' End: SCR #44724;
    
    Set ReturnNewZPINode = oZPINode

Exit Function

ErrTrap:
'   If gobjShared Is Nothing Then
'      Set gobjShared = New CWShared
'   End If
'   gobjShared.ShowVBError Err.Number, Err.Description, "ReturnNewZPINode Function", "DMBE", vbExclamation
   Err.Clear
        
End Function

Public Function ReturnNewZPGNode(ByVal intCount As Long, ByVal Mmi_GenericDataRS As ADODB.Recordset) As MSXML.IXMLDOMElement
' Generic Patient Data Segment

' sample function output
'   <ZPG>
'      <ZPG.1>Count</ZPG.1>
'      <ZPG.2>Data Type</ZPG.2>
'      <ZPG.3>
'         <CE.1>Data Type Identifier</CE.1>
'         <CE.2>Data Type Text</CE.2>
'      </ZPG.3>
'      <ZPG.4>Ordinal</ZPG.4>
'      <ZPG.5>Parent</ZPG.5>
'      <ZPG.6>Value_RSID</ZPG.6>
'      <ZPG.7>Error Text</ZPG.7>
'   </ZPG>

    Dim oDom As MSXML.DOMDocument
    Dim oZPGNode As MSXML.IXMLDOMElement
    Dim oZPGChildLevel1 As MSXML.IXMLDOMElement
    Dim oZPGChildLevel2 As MSXML.IXMLDOMElement
    
    On Error GoTo ErrTrap
    
    Set oDom = New MSXML.DOMDocument
    Set oZPGNode = oDom.createElement("ZPG")
    
    ' Set ID - Count
    Set oZPGChildLevel1 = oDom.createElement("ZPG.1")
    oZPGChildLevel1.Text = CStr(intCount)
    oZPGNode.appendChild oZPGChildLevel1
    
    ' Data Type - Table MMI_GenericData:Type_NCID
    Set oZPGChildLevel1 = oDom.createElement("ZPG.2")
    If Not IsNull(Mmi_GenericDataRS.Fields("Type_NCID").Value) Then
         'oZPGChildLevel1.Text = GetNCIDDescription(Mmi_GenericDataRS.Fields("Type_NCID").Value)
        oZPGChildLevel1.Text = NullToNothing(Mmi_GenericDataRS.Fields("Type_NCID").Value)
    End If
    oZPGNode.appendChild oZPGChildLevel1
     
    ' Data Value
    Set oZPGChildLevel1 = oDom.createElement("ZPG.3")
        ' Identifier - Table MMI_GenericData:Value_NCID
        Set oZPGChildLevel2 = oDom.createElement("CE.1")
        oZPGChildLevel2.Text = NullToNothing(Mmi_GenericDataRS.Fields("Value_NCID").Value)
        oZPGChildLevel1.appendChild oZPGChildLevel2
        ' Text - Table MMI_GenericData:Value_Text
        Set oZPGChildLevel2 = oDom.createElement("CE.2")
        oZPGChildLevel2.Text = NullToNothing(Mmi_GenericDataRS.Fields("Value_Text").Value)
        oZPGChildLevel1.appendChild oZPGChildLevel2
    oZPGNode.appendChild oZPGChildLevel1
    
    ' Ordinal - Table MMI_GenericData:Ordinal
    Set oZPGChildLevel1 = oDom.createElement("ZPG.4")
    oZPGChildLevel1.Text = NullToNothing(Mmi_GenericDataRS.Fields("Ordinal").Value)
    oZPGNode.appendChild oZPGChildLevel1
    
    ' Parent - Table MMI_GenericData:Parent
    Set oZPGChildLevel1 = oDom.createElement("ZPG.5")
    oZPGChildLevel1.Text = NullToNothing(Mmi_GenericDataRS.Fields("Parent").Value)
    oZPGNode.appendChild oZPGChildLevel1
    
    ' Value_RSID - Table MMI_GenericData:Value_RSID
    Set oZPGChildLevel1 = oDom.createElement("ZPG.6")
    oZPGChildLevel1.Text = NullToNothing(Mmi_GenericDataRS.Fields("Value_RSID").Value)
    oZPGNode.appendChild oZPGChildLevel1
    
    ' Error Text - Table MMI_GenericData:Error_Text
    Set oZPGChildLevel1 = oDom.createElement("ZPG.7")
    oZPGChildLevel1.Text = NullToNothing(Mmi_GenericDataRS.Fields("Error_Text").Value)
    oZPGNode.appendChild oZPGChildLevel1
    
    Set ReturnNewZPGNode = oZPGNode

Exit Function

ErrTrap:
   If gobjShared Is Nothing Then
      Set gobjShared = New CWShared
   End If
   gobjShared.ShowVBError Err.Number, Err.Description, "ReturnNewZPGNode Function", "DMBE", vbExclamation
   Resume Next 'There is no reason this procedure should not continue to create xml if it errors
    
End Function


Public Function ReturnNewEVNNode(ByVal strEventType As String) As MSXML.IXMLDOMElement
   
' sample function output
'   <EVN>
'      <EVN.1>Event Type</EVN.1>
'      <EVN.2>Event Date/Time</EVN.2>
'      <EVN.6>Date/Time Event Occurred</EVN.6>
'   </EVN>
   
    Dim oDom As MSXML.DOMDocument
    Dim oEVNNode As MSXML.IXMLDOMElement
    Dim oEVNChildLevel1 As MSXML.IXMLDOMElement
    
    On Error GoTo ErrTrap
    
    Set oDom = New MSXML.DOMDocument
    Set oEVNNode = oDom.createElement("EVN")
    
    ' Recorded Date/Time - Should be based on GMT
    Set oEVNChildLevel1 = oDom.createElement("EVN.2")
    oEVNChildLevel1.Text = Format(Now, "YYYYMMDDHHNNSS")
    oEVNNode.appendChild oEVNChildLevel1
       
    ' Date/Time Event Occurred
    Set oEVNChildLevel1 = oDom.createElement("EVN.6")
    oEVNChildLevel1.Text = Format(Now, "YYYYMMDDHHNNSS")
    'oEVNChildLevel1.Text = "EventOccurred"
    oEVNNode.appendChild oEVNChildLevel1
   
    Set ReturnNewEVNNode = oEVNNode
    
Exit Function

ErrTrap:
   If gobjShared Is Nothing Then
      Set gobjShared = New CWShared
   End If
   gobjShared.ShowVBError Err.Number, Err.Description, "ReturnNewEVNNode Function", "DMBE", vbExclamation
   'Err.Clear
   Resume Next 'There is no reason this procedure should not continue to create xml if it errors
   
End Function

Public Function ReturnFakePV1Segment() As MSXML.IXMLDOMElement

    Dim oDom As MSXML.DOMDocument
    Dim oPV1Element As MSXML.IXMLDOMElement
    Dim oPV1ChildNode As MSXML.IXMLDOMElement
        
    Set oDom = New MSXML.DOMDocument
    Set oPV1Element = oDom.createElement("PV1")
'        Set oPV1ChildNode = oDOM.createNode(NODE_ELEMENT, "PV1.1", "")
'        oPV1ChildNode.Text = "1"
'    oPV1Element.appendChild oPV1ChildNode
    
        Set oPV1ChildNode = oDom.createNode(NODE_ELEMENT, "PV1.2", "")
        oPV1ChildNode.Text = "O"
    
    oPV1Element.appendChild oPV1ChildNode
    
    Set ReturnFakePV1Segment = oPV1Element
    
'        Set oPV1ChildNode = oDOM.createElement("PV1.3")
'            Set oPV1ChildChildNode = oDOM.createElement("PL.1")
'            oPV1ChildChildNode.Text = ""
'        oPV1ChildNode.appendChild oPV1ChildChildNode
'            Set oPV1ChildChildNode = oDOM.createElement("PL.2")
'            oPV1ChildChildNode.Text = ""
'        oPV1ChildNode.appendChild oPV1ChildChildNode
'            Set oPV1ChildChildNode = oDOM.createElement("PL.3")
'            oPV1ChildChildNode.Text = ""
'        oPV1ChildNode.appendChild oPV1ChildChildNode
'            Set oPV1ChildChildNode = oDOM.createElement("PL.4")
'            oPV1ChildChildNode.Text = ""
'        oPV1ChildNode.appendChild oPV1ChildChildNode
'            Set oPV1ChildChildNode = oDOM.createElement("PL.5")
'            oPV1ChildChildNode.Text = ""
'        oPV1ChildNode.appendChild oPV1ChildChildNode
'            Set oPV1ChildChildNode = oDOM.createElement("PL.6")
'            oPV1ChildChildNode.Text = ""
'        oPV1ChildNode.appendChild oPV1ChildChildNode
'            Set oPV1ChildChildNode = oDOM.createElement("PL.7")
'            oPV1ChildChildNode.Text = ""
'        oPV1ChildNode.appendChild oPV1ChildChildNode
'            Set oPV1ChildChildNode = oDOM.createElement("PL.8")
'            oPV1ChildChildNode.Text = ""
'        oPV1ChildNode.appendChild oPV1ChildChildNode
'            Set oPV1ChildChildNode = oDOM.createElement("PL.9")
'            oPV1ChildChildNode.Text = ""
'        oPV1ChildNode.appendChild oPV1ChildChildNode
'
'    oPV1Element.appendChild oPV1ChildNode
'
'        Set oPV1ChildNode = oDOM.createElement("PV1.7")
'            Set oPV1ChildChildNode = oDOM.createElement("XCN.1")
'            oPV1ChildChildNode.Text = ""
'        oPV1ChildNode.appendChild oPV1ChildChildNode
'            Set oPV1ChildChildNode = oDOM.createElement("XCN.2")
'            oPV1ChildChildNode.Text = ""
'        oPV1ChildNode.appendChild oPV1ChildChildNode
'            Set oPV1ChildChildNode = oDOM.createElement("XCN.3")
'            oPV1ChildChildNode.Text = ""
'        oPV1ChildNode.appendChild oPV1ChildChildNode
'            Set oPV1ChildChildNode = oDOM.createElement("XCN.4")
'            oPV1ChildChildNode.Text = ""
'        oPV1ChildNode.appendChild oPV1ChildChildNode
'
'    oPV1Element.appendChild oPV1ChildNode
'
'        Set oPV1ChildNode = oDOM.createNode(NODE_ELEMENT, "PV1.18", "")
'        oPV1ChildNode.Text = ""
'    oPV1Element.appendChild oPV1ChildNode
'
'        Set oPV1ChildNode = oDOM.createNode(NODE_ELEMENT, "PV1.36", "")
'        oPV1ChildNode.Text = ""
'    oPV1Element.appendChild oPV1ChildNode
'
'        Set oPV1ChildNode = oDOM.createNode(NODE_ELEMENT, "PV1.49", "")
'        oPV1ChildNode.Text = ""
'    oPV1Element.appendChild oPV1ChildNode
'
'        Set oPV1ChildNode = oDOM.createNode(NODE_ELEMENT, "PV1.44", "")
'        oPV1ChildNode.Text = ""
'    oPV1Element.appendChild oPV1ChildNode
'
'        Set oPV1ChildNode = oDOM.createNode(NODE_ELEMENT, "PV1.45", "")
'        oPV1ChildNode.Text = ""
'    oPV1Element.appendChild oPV1ChildNode

End Function

Public Function ReturnNewMSHNode(ByVal strEventName As String, strEventType As String, strMsgNumber As String) As MSXML.IXMLDOMElement

' sample function output
'   <MSH>
'      <MSH.1>|</MSH.1>
'      <MSH.2>^~\&amp;</MSH.2>
'      <MSH.3>
'         <HD.1>DMU</HD.1>
'      </MSH.3>
'      <MSH.4>
'         <HD.1>NTGHOST77</HD.1>
'         <HD.2>131072</HD.2>
'      </MSH.4>
'      <MSH.5>
'         <HD.1>RHL</HD.1>
'      </MSH.5>
'      <MSH.6>
'         <HD.1>NTGHOST77</HD.1>
'         <HD.2>65536</HD.2>
'      </MSH.6>
'      <MSH.7>20010911183456</MSH.7>
'      <MSH.9>
'         <CM_MSG_TYPE.1>ADT</CM_MSG_TYPE.1>
'         <CM_MSG_TYPE.2>A28</CM_MSG_TYPE.2>
'      </MSH.9>
'      <MSH.10>MSG000001</MSH.10>
'      <MSH.11>
'         <PT.1>P</PT.1>
'      </MSH.11>
'      <MSH.12>
'         <VID.1>2.3.1</VID.1>
'      </MSH.12>
'      <MSH.15>NE</MSH.15>
'      <MSH.16>AL</MSH.16>
'   </MSH>

    Dim oDom As MSXML.DOMDocument
    Dim oMSHNode As MSXML.IXMLDOMElement
    Dim oMSHChildLevel1 As MSXML.IXMLDOMElement
    Dim oMSHChildLevel2 As MSXML.IXMLDOMElement
    Dim intMaxComputerNameLength As Integer
    Dim lngReturnLength As Long
    Dim strComputerName As String
    Dim blnIsSuccessful As Boolean
    
    On Error GoTo ErrTrap
    
    Set oDom = New MSXML.DOMDocument
    Set oMSHNode = oDom.createElement("MSH")
    
    Set oMSHChildLevel1 = oDom.createNode(NODE_ELEMENT, "MSH.1", "")
    ' Field separator - default to "|"
    oMSHChildLevel1.Text = "|"
    oMSHNode.appendChild oMSHChildLevel1
    
    Set oMSHChildLevel1 = oDom.createElement("MSH.2")
    ' Encoding Characters - default to "^~\&amp;"
    oMSHChildLevel1.Text = "^~\&"
    oMSHNode.appendChild oMSHChildLevel1
    
    Set oMSHChildLevel1 = oDom.createNode(NODE_ELEMENT, "MSH.3", "")
        Set oMSHChildLevel2 = oDom.createNode(NODE_ELEMENT, "HD.1", "")
        ' Sending Application - default to "DMU"
        oMSHChildLevel2.Text = "DMU"
        oMSHChildLevel1.appendChild oMSHChildLevel2
    oMSHNode.appendChild oMSHChildLevel1
            
    Set oMSHChildLevel1 = oDom.createNode(NODE_ELEMENT, "MSH.4", "")
    ' Sending Facility
        
        Set oMSHChildLevel2 = oDom.createNode(NODE_ELEMENT, "HD.1", "")
        'Workstation Computer Name
        intMaxComputerNameLength = 256
        strComputerName = String$(intMaxComputerNameLength + 1, " ")
        lngReturnLength = intMaxComputerNameLength + 1
        blnIsSuccessful = GetComputerName(strComputerName, lngReturnLength)
        If blnIsSuccessful Then
            strComputerName = Trim(strComputerName)
            If lngReturnLength > 0 Then
                strComputerName = Mid(strComputerName, 1, lngReturnLength)
            Else
                strComputerName = "UNKNOWN"
            End If
        Else
            strComputerName = "UNKNOWN"
        End If
        
        oMSHChildLevel2.Text = strComputerName
        oMSHChildLevel1.appendChild oMSHChildLevel2
        
        Set oMSHChildLevel2 = oDom.createNode(NODE_ELEMENT, "HD.2", "")
        ' Subscription Event Type - default to 131072
        oMSHChildLevel2.Text = "131072"
        oMSHChildLevel1.appendChild oMSHChildLevel2
    oMSHNode.appendChild oMSHChildLevel1
    
    Set oMSHChildLevel1 = oDom.createNode(NODE_ELEMENT, "MSH.5", "")
    ' Receiving Facility
        Set oMSHChildLevel2 = oDom.createNode(NODE_ELEMENT, "HD.1", "")
        oMSHChildLevel2.Text = "RHL"
        ' Workstation Computer Name - default to RHL
        oMSHChildLevel1.appendChild oMSHChildLevel2
    oMSHNode.appendChild oMSHChildLevel1
    
    Set oMSHChildLevel1 = oDom.createNode(NODE_ELEMENT, "MSH.6", "")
        Set oMSHChildLevel2 = oDom.createNode(NODE_ELEMENT, "HD.1", "")
        oMSHChildLevel2.Text = strComputerName
        oMSHChildLevel1.appendChild oMSHChildLevel2
        Set oMSHChildLevel2 = oDom.createNode(NODE_ELEMENT, "HD.2", "")
        ' Publish Event Type - default to 65536
        oMSHChildLevel2.Text = "65536"
        oMSHChildLevel1.appendChild oMSHChildLevel2
    oMSHNode.appendChild oMSHChildLevel1
        
    Set oMSHChildLevel1 = oDom.createNode(NODE_ELEMENT, "MSH.7", "")
    oMSHChildLevel1.Text = Format(Now, "YYYYMMDDHHNNSS")
    oMSHNode.appendChild oMSHChildLevel1
    
    'Set oMSHChildLevel1 = oDom.createNode(NODE_ELEMENT, "MSH.8", "")
    'No MSH.8 node in old code
    
    Set oMSHChildLevel1 = oDom.createNode(NODE_ELEMENT, "MSH.9", "")
        ' Message Type
        Set oMSHChildLevel2 = oDom.createNode(NODE_ELEMENT, "CM_MSG_TYPE.1", "")
        oMSHChildLevel2.Text = strEventType
        oMSHChildLevel1.appendChild oMSHChildLevel2
        ' Trigger Event
        Set oMSHChildLevel2 = oDom.createNode(NODE_ELEMENT, "CM_MSG_TYPE.2", "")
        oMSHChildLevel2.Text = strEventName
        oMSHChildLevel1.appendChild oMSHChildLevel2
    oMSHNode.appendChild oMSHChildLevel1
    
    Set oMSHChildLevel1 = oDom.createNode(NODE_ELEMENT, "MSH.10", "")
     ' Message Control ID - MSGxxxxxx
    oMSHChildLevel1.Text = "MSG" & Format(CStr(strMsgNumber), "0#####")
    oMSHNode.appendChild oMSHChildLevel1
    
    Set oMSHChildLevel1 = oDom.createNode(NODE_ELEMENT, "MSH.11", "")
    ' Processing ID = Production (P)
        Set oMSHChildLevel2 = oDom.createNode(NODE_ELEMENT, "PT.1", "")
        oMSHChildLevel2.Text = "P"
        oMSHChildLevel1.appendChild oMSHChildLevel2
    oMSHNode.appendChild oMSHChildLevel1
    
    Set oMSHChildLevel1 = oDom.createNode(NODE_ELEMENT, "MSH.12", "")
    ' Version ID - default to 2.3.1
        Set oMSHChildLevel2 = oDom.createNode(NODE_ELEMENT, "VID.1", "")
        oMSHChildLevel2.Text = "2.3.1"
        oMSHChildLevel1.appendChild oMSHChildLevel2
    oMSHNode.appendChild oMSHChildLevel1
    
    'Set oMSHChildLevel1 = oDom.createNode(NODE_ELEMENT, "MSH.13", "")
    'Set oMSHChildLevel1 = oDom.createNode(NODE_ELEMENT, "MSH.14", "")
    
    Set oMSHChildLevel1 = oDom.createNode(NODE_ELEMENT, "MSH.15", "")
    ' Accept Acknowlegment Type - default to "NE" to never receive an "accept acknowledgement"
    oMSHChildLevel1.Text = "NE"
    oMSHNode.appendChild oMSHChildLevel1
    
    Set oMSHChildLevel1 = oDom.createNode(NODE_ELEMENT, "MSH.16", "")
    ' Application Acknowledgement Type - default to "AL" to always send an "application acknowledgement"
    oMSHChildLevel1.Text = "AL"
    oMSHNode.appendChild oMSHChildLevel1
        
    Set ReturnNewMSHNode = oMSHNode

Exit Function

ErrTrap:
   If gobjShared Is Nothing Then
      Set gobjShared = New CWShared
   End If
   gobjShared.ShowVBError Err.Number, Err.Description, "ReturnNewMSHNode Function", "DMBE", vbExclamation
   'Err.Clear
   Resume Next 'There is no reason this procedure should not continue to create xml if it errors

End Function

Public Function ReturnNewPIDNode(ByVal lngCount As Long) As MSXML.IXMLDOMElement

' sample function output
'   <PID>
'      <PID.1>1</PID.1>
'      <PID.2>
'         <CX.1>128548</CX.1>
'      </PID.2>
'      <PID.3.LST>
'         <PID.3>
'            <CX.1>279875785</CX.1>
'            <CX.5>SSN</CX.5>
'         </PID.3>
'         <PID.3>
'            <CX.1>20</CX.1>
'            <CX.5>FMP</CX.5>
'         </PID.3>
'         <PID.3>
'            <CX.1>1254004</CX.1>
'            <CX.5>PRN</CX.5>
'         </PID.3>
'         <PID.3>
'            <CX.1>279875785</CX.1>
'            <CX.5>SPONSORSSN</CX.5>
'         </PID.3>
'      </PID.3.LST>
'      <PID.5.LST>
'         <PID.5>
'            <XPN.1>LOPEZ</XPN.1>
'            <XPN.2>ERIKA</XPN.2>
'            <XPN.3></XPN.3>
'            <XPN.4></XPN.4>
'            <XPN.5></XPN.5>
'            <XPN.6></XPN.6>
'         </PID.5>
'      </PID.5.LST>
'      <PID.6>
'         <XPN.1></XPN.1>
'      </PID.6>
'      <PID.7>19540321</PID.7>
'      <PID.8>F</PID.8>
'      <PID.9>
'         <XPN.1></XPN.1>
'         <XPN.2></XPN.2>
'         <XPN.3></XPN.3>
'         <XPN.4></XPN.4>
'         <XPN.5></XPN.5>
'         <XPN.6></XPN.6>
'      </PID.9>
'      <PID.10.LST>
'         <PID.10>
'            <CE.1></CE.1>
'            <CE.2>White</CE.2>
'            <CE.3></CE.3>
'            <CE.4>139</CE.4>
'            <CE.5>White</CE.5>
'            <CE.6>99NCI</CE.6>
'         </PID.10>
'      </PID.10.LST>
'      <PID.11.LST>
'         <PID.11>
'            <XAD.1>991627 HOAPONO PLACE</XAD.1>
'            <XAD.2>Michigan Road</XAD.2>
'            <XAD.3>SCHOFIELD BARRACKS</XAD.3>
'            <XAD.4>HI</XAD.4>
'            <XAD.5>96857</XAD.5>
'            <XAD.6></XAD.6>
'            <XAD.9></XAD.9>
'         </PID.11>
'      </PID.11.LST>
'      <PID.13.LST>
'         <PID.13>
'            <XTN.1>629-5984</XTN.1>
'         </PID.13>
'      </PID.13.LST>
'      <PID.14.LST>
'         <PID.14>
'            <XTN.1>655-9094</XTN.1>
'         </PID.14>
'      </PID.14.LST>
'      <PID.16>
'         <CE.1></CE.1>
'         <CE.2>Single, Never Married</CE.2>
'         <CE.3></CE.3>
'         <CE.4>14503679</CE.4>
'         <CE.5>Single, Never Married</CE.5>
'         <CE.6>99NCI</CE.6>
'      </PID.16>
'      <PID.17>
'         <CE.1></CE.1>
'         <CE.2>Roman Catholic</CE.2>
'         <CE.3></CE.3>
'         <CE.4>841</CE.4>
'         <CE.5>Roman Catholic</CE.5>
'         <CE.6>99NCI</CE.6>
'      </PID.17>
'   </PID>

    Dim oDom As MSXML.DOMDocument
    Dim oPIDNode As MSXML.IXMLDOMElement
    Dim oPIDChildLevel1 As MSXML.IXMLDOMElement
    Dim oPIDChildLevel2 As MSXML.IXMLDOMElement
    Dim oPIDChildLevel3 As MSXML.IXMLDOMElement
    Dim RS As ADODB.Recordset
    Dim strSQL As String
    Dim strTemp As String
    Dim bNoSponSSN As Boolean
    On Error GoTo ErrTrap
    
    
    Set oDom = New MSXML.DOMDocument
    Set oPIDNode = oDom.createElement("PID")
    
    Set oPIDChildLevel1 = oDom.createNode(NODE_ELEMENT, "PID.1", "")
    ' Patient Count
    oPIDChildLevel1.Text = CStr(lngCount)
    oPIDNode.appendChild oPIDChildLevel1
    
    Set oPIDChildLevel1 = oDom.createNode(NODE_ELEMENT, "PID.2", "")
        Set oPIDChildLevel2 = oDom.createNode(NODE_ELEMENT, "CX.1", "")
        ' Patient Unit Number
        oPIDChildLevel2.Text = NullToNothing(gobjPatient.unitNumber)
        oPIDChildLevel1.appendChild oPIDChildLevel2
    oPIDNode.appendChild oPIDChildLevel1
    
    Set oPIDChildLevel1 = oDom.createNode(NODE_ELEMENT, "PID.3.LST", "")
        bNoSponSSN = False
        ' Patient SSN
        '<< Begin:SCR #27554;   Developer: Brian Mowbray 09/10/2002 04:40 PM
        ' Begin:SCR #28377;   Developer: MAG 11/13/2002 07:04 PM
        strTemp = NullToNothing(gobjPatient.SponsorSSN)
        ' End: SCR #28377;
        
        If Len(Trim$(strTemp)) = 0 Then
          strTemp = gobjPatient.SSN
          bNoSponSSN = True
        End If
        
        If Len(Trim$(strTemp)) <> 0 Then
            Set oPIDChildLevel2 = oDom.createNode(NODE_ELEMENT, "PID.3", "")
                Set oPIDChildLevel3 = oDom.createNode(NODE_ELEMENT, "CX.1", "")
                oPIDChildLevel3.Text = strTemp
                oPIDChildLevel2.appendChild oPIDChildLevel3
                Set oPIDChildLevel3 = oDom.createNode(NODE_ELEMENT, "CX.5", "")
                oPIDChildLevel3.Text = "SSN"
                oPIDChildLevel2.appendChild oPIDChildLevel3
            oPIDChildLevel1.appendChild oPIDChildLevel2
        End If
        ' Family Member Prefix
        If bNoSponSSN Then
          strTemp = "20"
        Else
          strTemp = NullToNothing(gobjPatient.FMP)
        End If
        If Len(Trim$(strTemp)) <> 0 Then
            Set oPIDChildLevel2 = oDom.createNode(NODE_ELEMENT, "PID.3", "")
                Set oPIDChildLevel3 = oDom.createNode(NODE_ELEMENT, "CX.1", "")
                oPIDChildLevel3.Text = strTemp
                oPIDChildLevel2.appendChild oPIDChildLevel3
                Set oPIDChildLevel3 = oDom.createNode(NODE_ELEMENT, "CX.5", "")
                oPIDChildLevel3.Text = "FMP"
                oPIDChildLevel2.appendChild oPIDChildLevel3
            oPIDChildLevel1.appendChild oPIDChildLevel2
        End If
        ' Patient Record Number
        strTemp = NullToNothing(gobjPatient.MedicalRecordNumber)
        If Len(Trim$(strTemp)) <> 0 Then
            Set oPIDChildLevel2 = oDom.createNode(NODE_ELEMENT, "PID.3", "")
                Set oPIDChildLevel3 = oDom.createNode(NODE_ELEMENT, "CX.1", "")
                oPIDChildLevel3.Text = strTemp
                oPIDChildLevel2.appendChild oPIDChildLevel3
                Set oPIDChildLevel3 = oDom.createNode(NODE_ELEMENT, "CX.5", "")
                oPIDChildLevel3.Text = "PRN"
                oPIDChildLevel2.appendChild oPIDChildLevel3
            oPIDChildLevel1.appendChild oPIDChildLevel2
        End If
        
'<< Begin:SCR #27601;   Developer: Brian Mowbray 09/12/2002 03:38 PM
'Removed per customer request
'        ' Begin:SCR #26570;   Developer: MAG 08/13/2002 03:13 PM
'        ' Added Sponsor SSN
'        strTemp = NullToNothing(gobjPatient.SponsorSSN)
'        If Len(Trim$(strTemp)) <> 0 Then
'            Set oPIDChildLevel2 = oDom.createNode(NODE_ELEMENT, "PID.3", "")
'                Set oPIDChildLevel3 = oDom.createNode(NODE_ELEMENT, "CX.1", "")
'                oPIDChildLevel3.Text = strTemp
'                oPIDChildLevel2.appendChild oPIDChildLevel3
'                Set oPIDChildLevel3 = oDom.createNode(NODE_ELEMENT, "CX.5", "")
'                oPIDChildLevel3.Text = "SSSN"
'                oPIDChildLevel2.appendChild oPIDChildLevel3
'            oPIDChildLevel1.appendChild oPIDChildLevel2
'        End If
'        ' End: SCR #26570;
'        '>> End: SCR #27554;
'>> End: SCR #27601;
    
    'Complete this level
    oPIDNode.appendChild oPIDChildLevel1
    
    Set oPIDChildLevel1 = oDom.createNode(NODE_ELEMENT, "PID.5.LST", "")
        Set oPIDChildLevel2 = oDom.createNode(NODE_ELEMENT, "PID.5", "")
            'Patient Last Name
            Set oPIDChildLevel3 = oDom.createNode(NODE_ELEMENT, "XPN.1", "")
            oPIDChildLevel3.Text = NullToNothing(gobjPatient.CurrentPatient.PatientName.Last_name)
            oPIDChildLevel2.appendChild oPIDChildLevel3
            'Patient First Name
            Set oPIDChildLevel3 = oDom.createNode(NODE_ELEMENT, "XPN.2", "")
            oPIDChildLevel3.Text = NullToNothing(gobjPatient.CurrentPatient.PatientName.First_name)
            oPIDChildLevel2.appendChild oPIDChildLevel3
            'Patient Middle Name
            Set oPIDChildLevel3 = oDom.createNode(NODE_ELEMENT, "XPN.3", "")
            oPIDChildLevel3.Text = NullToNothing(gobjPatient.CurrentPatient.PatientName.Middle_name)
            oPIDChildLevel2.appendChild oPIDChildLevel3
            'Patient Suffix
            Set oPIDChildLevel3 = oDom.createNode(NODE_ELEMENT, "XPN.4", "")
            oPIDChildLevel3.Text = NullToNothing(gobjPatient.CurrentPatient.PatientName.Suffix)
            oPIDChildLevel2.appendChild oPIDChildLevel3
            'Patient Prefix
            Set oPIDChildLevel3 = oDom.createNode(NODE_ELEMENT, "XPN.5", "")
            oPIDChildLevel3.Text = NullToNothing(gobjPatient.CurrentPatient.PatientName.title)
            oPIDChildLevel2.appendChild oPIDChildLevel3
            'Patient Degree
            Set oPIDChildLevel3 = oDom.createNode(NODE_ELEMENT, "XPN.6", "")
            oPIDChildLevel3.Text = NullToNothing(gobjPatient.CurrentPatient.PatientName.Degree)
            oPIDChildLevel2.appendChild oPIDChildLevel3
        oPIDChildLevel1.appendChild oPIDChildLevel2
    oPIDNode.appendChild oPIDChildLevel1
    
    'Mother's Maiden Name
    Set oPIDChildLevel1 = oDom.createNode(NODE_ELEMENT, "PID.6", "")
        Set oPIDChildLevel2 = oDom.createNode(NODE_ELEMENT, "XPN.1", "")
        oPIDChildLevel2.Text = NullToNothing(gobjPatient.MothersMaidenName)
        oPIDChildLevel1.appendChild oPIDChildLevel2
    oPIDNode.appendChild oPIDChildLevel1
    
    'Patient Birth Date
    Set oPIDChildLevel1 = oDom.createNode(NODE_ELEMENT, "PID.7", "")
    strTemp = NullToNothing(gobjPatient.DateOfBirth)
    If Val(strTemp) > 0 Then
        strTemp = CDate(strTemp)
        If IsValidDate(strTemp) = True Then
            oPIDChildLevel1.Text = Format(strTemp, "YYYYMMDDHHNNSS")
        End If
    End If
    oPIDNode.appendChild oPIDChildLevel1
    
    'Patient Sex
    '<SCR 22851 CC>
    Set oPIDChildLevel1 = oDom.createNode(NODE_ELEMENT, "PID.8", "")
    Select Case gobjPatient.SEX
        Case "F"
            oPIDChildLevel1.Text = "F"
        Case "M"
            oPIDChildLevel1.Text = "M"
        Case "B"
            oPIDChildLevel1.Text = "B"
        Case "O"
            oPIDChildLevel1.Text = "O"
        Case Else
            oPIDChildLevel1.Text = "U"
    End Select
    oPIDNode.appendChild oPIDChildLevel1
    '</SCR 22851 CC>
    
    'Patient Alias
    Set oPIDChildLevel1 = oDom.createNode(NODE_ELEMENT, "PID.9", "")
        
        Set oPIDChildLevel2 = oDom.createNode(NODE_ELEMENT, "XPN.1", "")
        oPIDChildLevel2.Text = NullToNothing(gobjPatient.CurrentPatient.PatientName.Last_name2)
        oPIDChildLevel1.appendChild oPIDChildLevel2
        
        Set oPIDChildLevel2 = oDom.createNode(NODE_ELEMENT, "XPN.2", "")
        oPIDChildLevel2.Text = NullToNothing(gobjPatient.CurrentPatient.PatientName.First_name2)
        oPIDChildLevel1.appendChild oPIDChildLevel2
        
        Set oPIDChildLevel2 = oDom.createNode(NODE_ELEMENT, "XPN.3", "")
        oPIDChildLevel2.Text = NullToNothing(gobjPatient.CurrentPatient.PatientName.Middle_name2)
        oPIDChildLevel1.appendChild oPIDChildLevel2
        
        Set oPIDChildLevel2 = oDom.createNode(NODE_ELEMENT, "XPN.4", "")
        oPIDChildLevel2.Text = NullToNothing(gobjPatient.CurrentPatient.PatientName.Suffix)
        oPIDChildLevel1.appendChild oPIDChildLevel2

        Set oPIDChildLevel2 = oDom.createNode(NODE_ELEMENT, "XPN.5", "")
        oPIDChildLevel2.Text = NullToNothing(gobjPatient.CurrentPatient.PatientName.title)
        oPIDChildLevel1.appendChild oPIDChildLevel2
    
        Set oPIDChildLevel2 = oDom.createNode(NODE_ELEMENT, "XPN.6", "")
        oPIDChildLevel2.Text = NullToNothing(gobjPatient.CurrentPatient.PatientName.Degree)
        oPIDChildLevel1.appendChild oPIDChildLevel2

    oPIDNode.appendChild oPIDChildLevel1
    
    'Patient Race
    
    '<SCR 20384 CC>
    
    Dim lngNCID As Long
    
    If oDas Is Nothing Then
        Set oDas = oDatabase.CHCSII_DAS(Auto) 'New GEMS_DAS
    End If
    
    strTemp = NullToNothing(gobjPatient.Race)
       
    If strTemp = "" Then
        lngNCID = 0
    Else: lngNCID = CLng(strTemp) 'gobjPatient.Race
    End If
    
    '<< Begin:SCR #26532;   Developer: Brian Mowbray 08/12/2002 02:26 PM
    If geDestination = eDM_EIC Then
        
      strSQL = "SELECT SAMS_DESCRIPTION FROM TMIP_SAMS_NCID_MAPPING "
      strSQL = strSQL & "WHERE CHCSII_NCID = " & lngNCID
      'This applies to both TMIP and EIC
  
      Set RS = New ADODB.Recordset
      Set RS = oDas.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)
                  
      If Not RS Is Nothing Then
          If Not RS.EOF Then
              strTemp = NullToNothing(RS.Fields(0).Value) 'rs has only one field
          End If
      Else: strTemp = ""
      End If
    Else
      If lngNCID <> 0 Then
        strTemp = GetNCIDDescription(lngNCID)
      End If
    End If
    '>> End: SCR #26532;
    
    Set oPIDChildLevel1 = oDom.createNode(NODE_ELEMENT, "PID.10.LST", "")
        Set oPIDChildLevel2 = oDom.createNode(NODE_ELEMENT, "PID.10", "")
            
            Set oPIDChildLevel3 = oDom.createNode(NODE_ELEMENT, "CE.1", "")
            oPIDChildLevel3.Text = ""
            oPIDChildLevel2.appendChild oPIDChildLevel3
                        
            Set oPIDChildLevel3 = oDom.createNode(NODE_ELEMENT, "CE.2", "")
            oPIDChildLevel3.Text = strTemp
            oPIDChildLevel2.appendChild oPIDChildLevel3
            
            Set oPIDChildLevel3 = oDom.createNode(NODE_ELEMENT, "CE.3", "")
            oPIDChildLevel3.Text = ""
            oPIDChildLevel2.appendChild oPIDChildLevel3
            
            Set oPIDChildLevel3 = oDom.createNode(NODE_ELEMENT, "CE.4", "")
            oPIDChildLevel3.Text = gobjPatient.Race
            oPIDChildLevel2.appendChild oPIDChildLevel3
            
            Set oPIDChildLevel3 = oDom.createNode(NODE_ELEMENT, "CE.5", "")
            oPIDChildLevel3.Text = strTemp
            oPIDChildLevel2.appendChild oPIDChildLevel3
            
            Set oPIDChildLevel3 = oDom.createNode(NODE_ELEMENT, "CE.6", "")
            oPIDChildLevel3.Text = "99NCI"
            oPIDChildLevel2.appendChild oPIDChildLevel3
        
        oPIDChildLevel1.appendChild oPIDChildLevel2
    oPIDNode.appendChild oPIDChildLevel1
    
    'Patient Address
    Dim oAddress As GEMS_MMIObj.GEMS_Address
    Dim blnAddressFound As Boolean
    
    Set oPIDChildLevel1 = oDom.createNode(NODE_ELEMENT, "PID.11.LST", "")
    
    For Each oAddress In gobjPatient.CurrentPatient.Addresses
        If oAddress.AddrType = NCID_HOME_ADDRESS Then
            blnAddressFound = True
            Exit For
        End If
    Next
        
        'PID.11 is a required element
        Set oPIDChildLevel2 = oDom.createNode(NODE_ELEMENT, "PID.11", "")
        
        If blnAddressFound = True Then
          
            'Patient Street Address
            Set oPIDChildLevel3 = oDom.createNode(NODE_ELEMENT, "XAD.1", "")
            oPIDChildLevel3.Text = NullToNothing(oAddress.Street1)
            oPIDChildLevel2.appendChild oPIDChildLevel3
            'Patient Street Address
            Set oPIDChildLevel3 = oDom.createNode(NODE_ELEMENT, "XAD.2", "")
            oPIDChildLevel3.Text = NullToNothing(oAddress.Street2)
            oPIDChildLevel2.appendChild oPIDChildLevel3
            'City
            Set oPIDChildLevel3 = oDom.createNode(NODE_ELEMENT, "XAD.3", "")
            oPIDChildLevel3.Text = NullToNothing(oAddress.City)
            oPIDChildLevel2.appendChild oPIDChildLevel3
            'State
            Set oPIDChildLevel3 = oDom.createNode(NODE_ELEMENT, "XAD.4", "")
            oPIDChildLevel3.Text = NullToNothing(oAddress.State)
            oPIDChildLevel2.appendChild oPIDChildLevel3
            'Zip Code
            Set oPIDChildLevel3 = oDom.createNode(NODE_ELEMENT, "XAD.5", "")
            oPIDChildLevel3.Text = NullToNothing(oAddress.Zip)
            oPIDChildLevel2.appendChild oPIDChildLevel3
            'Country
            Set oPIDChildLevel3 = oDom.createNode(NODE_ELEMENT, "XAD.6", "")
            oPIDChildLevel3.Text = NullToNothing(oAddress.Country)
            oPIDChildLevel2.appendChild oPIDChildLevel3
            'County
            Set oPIDChildLevel3 = oDom.createNode(NODE_ELEMENT, "XAD.9", "")
            oPIDChildLevel3.Text = NullToNothing(oAddress.County)
            oPIDChildLevel2.appendChild oPIDChildLevel3
            
        End If
        oPIDChildLevel1.appendChild oPIDChildLevel2
    oPIDNode.appendChild oPIDChildLevel1
    
    Dim oPhone As GEMS_MMIObj.GEMS_Phone
    Dim blnPhoneFound As Boolean
    
    blnPhoneFound = False
    
    For Each oPhone In gobjPatient.CurrentPatient.phones
        If oPhone.PhoneType = NCID_HOME_PHONE Then
            blnPhoneFound = True
            'Patient Phone Numbers
            Set oPIDChildLevel1 = oDom.createNode(NODE_ELEMENT, "PID.13.LST", "")
                Set oPIDChildLevel2 = oDom.createNode(NODE_ELEMENT, "PID.13", "")
                    'Patient Home Phone Number
                    '1 [(999)] 999-9999 [X99999][C Any Text]
                    Set oPIDChildLevel3 = oDom.createNode(NODE_ELEMENT, "XTN.1", "")
                    
                    '<< Begin:SCR #25923;   Developer: Brian Mowbray 07/22/2002 06:44 PM
                    oPIDChildLevel3.Text = FormatAreaCode(oPhone.AreaCode) & " " & FormatPhone(oPhone.Phone) & FormatExtension(oPhone.Extension)
                    '>> End: SCR #25923;
                    oPIDChildLevel2.appendChild oPIDChildLevel3
                    '<SCR 25148 CC>
                    '2 Telecommunication Use Code
                    '3 Telecommunication Equipment Type (ID)
                    '4 Email Address
                    '5 Country Code
                    Set oPIDChildLevel3 = oDom.createNode(NODE_ELEMENT, "XTN.5", "")
                    oPIDChildLevel3.Text = oPhone.CountryCode
                    oPIDChildLevel2.appendChild oPIDChildLevel3
                    '6 Area/City Code
                    Set oPIDChildLevel3 = oDom.createNode(NODE_ELEMENT, "XTN.6", "")
                    oPIDChildLevel3.Text = oPhone.AreaCode
                    oPIDChildLevel2.appendChild oPIDChildLevel3
                    '7 Phone Number
                    Set oPIDChildLevel3 = oDom.createNode(NODE_ELEMENT, "XTN.7", "")
                    oPIDChildLevel3.Text = oPhone.Phone
                    oPIDChildLevel2.appendChild oPIDChildLevel3
                    '8 Extension
                    Set oPIDChildLevel3 = oDom.createNode(NODE_ELEMENT, "XTN.8", "")
                    oPIDChildLevel3.Text = oPhone.Extension
                    oPIDChildLevel2.appendChild oPIDChildLevel3
                    '9 Any Text
                oPIDChildLevel1.appendChild oPIDChildLevel2
            oPIDNode.appendChild oPIDChildLevel1
            Exit For
        End If
    Next
    
    ' Phone Number - Business
    For Each oPhone In gobjPatient.CurrentPatient.phones
        If oPhone.PhoneType = NCID_WORK_PHONE Then
            blnPhoneFound = True
            Set oPIDChildLevel1 = oDom.createNode(NODE_ELEMENT, "PID.14.LST", "")
                Set oPIDChildLevel2 = oDom.createNode(NODE_ELEMENT, "PID.14", "")
                    'Patient Home Phone Number
                    '1 [(999)] 999-9999 [X99999][C Any Text]
                    Set oPIDChildLevel3 = oDom.createNode(NODE_ELEMENT, "XTN.1", "")
                    '<< Begin:SCR #25923;   Developer: Brian Mowbray 07/22/2002 06:45 PM
                    oPIDChildLevel3.Text = FormatAreaCode(oPhone.AreaCode) & " " & FormatPhone(oPhone.Phone) & FormatExtension(oPhone.Extension)
                    '>> End: SCR #25923;
                    oPIDChildLevel2.appendChild oPIDChildLevel3
                    '<SCR 25148 CC>
                    '2 Telecommunication Use Code
                    '3 Telecommunication Equipment Type (ID)
                    '4 Email Address
                    '5 Country Code
                    Set oPIDChildLevel3 = oDom.createNode(NODE_ELEMENT, "XTN.5", "")
                    oPIDChildLevel3.Text = oPhone.CountryCode
                    oPIDChildLevel2.appendChild oPIDChildLevel3
                    '6 Area/City Code
                    Set oPIDChildLevel3 = oDom.createNode(NODE_ELEMENT, "XTN.6", "")
                    oPIDChildLevel3.Text = oPhone.AreaCode
                    oPIDChildLevel2.appendChild oPIDChildLevel3
                    '7 Phone Number
                    Set oPIDChildLevel3 = oDom.createNode(NODE_ELEMENT, "XTN.7", "")
                    oPIDChildLevel3.Text = oPhone.Phone
                    oPIDChildLevel2.appendChild oPIDChildLevel3
                    '8 Extension
                    Set oPIDChildLevel3 = oDom.createNode(NODE_ELEMENT, "XTN.8", "")
                    oPIDChildLevel3.Text = oPhone.Extension
                    oPIDChildLevel2.appendChild oPIDChildLevel3
                    '9 Any Text
                oPIDChildLevel1.appendChild oPIDChildLevel2
            oPIDNode.appendChild oPIDChildLevel1
            Exit For
        End If
    Next
    
    ' Patient Martial Status
    Set oPIDChildLevel1 = oDom.createNode(NODE_ELEMENT, "PID.16", "")
        Set oPIDChildLevel2 = oDom.createNode(NODE_ELEMENT, "CE.1", "")
        oPIDChildLevel2.Text = ""
        oPIDChildLevel1.appendChild oPIDChildLevel2
        Set oPIDChildLevel2 = oDom.createNode(NODE_ELEMENT, "CE.2", "")
        oPIDChildLevel2.Text = GetNCIDDescription(gobjPatient.MaritalStatus)
        oPIDChildLevel1.appendChild oPIDChildLevel2
        Set oPIDChildLevel2 = oDom.createNode(NODE_ELEMENT, "CE.3", "")
        oPIDChildLevel2.Text = ""
        oPIDChildLevel1.appendChild oPIDChildLevel2
        Set oPIDChildLevel2 = oDom.createNode(NODE_ELEMENT, "CE.4", "")
        oPIDChildLevel2.Text = gobjPatient.MaritalStatus
        oPIDChildLevel1.appendChild oPIDChildLevel2
        Set oPIDChildLevel2 = oDom.createNode(NODE_ELEMENT, "CE.5", "")
        oPIDChildLevel2.Text = GetNCIDDescription(gobjPatient.MaritalStatus)
        oPIDChildLevel1.appendChild oPIDChildLevel2
        Set oPIDChildLevel2 = oDom.createNode(NODE_ELEMENT, "CE.6", "")
        oPIDChildLevel2.Text = "99NCI"
        oPIDChildLevel1.appendChild oPIDChildLevel2
    oPIDNode.appendChild oPIDChildLevel1
        
    ' Patient Religion
    
    '<SCR 20384 CC>
    
    strTemp = NullToNothing(gobjPatient.Religion)
    If strTemp = "" Then
        lngNCID = 0
    Else: lngNCID = CLng(strTemp) 'gobjPatient.Religion
    End If
        
        
    '<< Begin:SCR #26532;   Developer: Brian Mowbray 08/12/2002 02:26 PM
    If geDestination = eDM_EIC Then
                    
      strSQL = "SELECT SAMS_DESCRIPTION FROM TMIP_SAMS_NCID_MAPPING WHERE "
      strSQL = strSQL & " CHCSII_NCID = " & lngNCID
      'This applies to both TMIP and SAMS
      
      Set RS = New ADODB.Recordset
      Set RS = oDas.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)
      If Not RS Is Nothing Then
          If Not RS.EOF Then
              strTemp = NullToNothing(RS.Fields(0).Value) 'rs has only 1 field
          Else
              strTemp = ""
          End If
      Else: strTemp = ""
      End If
    Else
      If lngNCID <> 0 Then
        strTemp = GetNCIDDescription(lngNCID)
      End If
    End If
    '>> End: SCR #26532;
    
    Set oPIDChildLevel1 = oDom.createNode(NODE_ELEMENT, "PID.17", "")
        Set oPIDChildLevel2 = oDom.createNode(NODE_ELEMENT, "CE.1", "")
        oPIDChildLevel2.Text = ""
        oPIDChildLevel1.appendChild oPIDChildLevel2
        Set oPIDChildLevel2 = oDom.createNode(NODE_ELEMENT, "CE.2", "")
        oPIDChildLevel2.Text = strTemp
        oPIDChildLevel1.appendChild oPIDChildLevel2
        Set oPIDChildLevel2 = oDom.createNode(NODE_ELEMENT, "CE.3", "")
        oPIDChildLevel2.Text = ""
        oPIDChildLevel1.appendChild oPIDChildLevel2
        Set oPIDChildLevel2 = oDom.createNode(NODE_ELEMENT, "CE.4", "")
        oPIDChildLevel2.Text = gobjPatient.Religion
        oPIDChildLevel1.appendChild oPIDChildLevel2
        Set oPIDChildLevel2 = oDom.createNode(NODE_ELEMENT, "CE.5", "")
        oPIDChildLevel2.Text = strTemp
        oPIDChildLevel1.appendChild oPIDChildLevel2
        Set oPIDChildLevel2 = oDom.createNode(NODE_ELEMENT, "CE.6", "")
        oPIDChildLevel2.Text = "99NCI"
        oPIDChildLevel1.appendChild oPIDChildLevel2
    oPIDNode.appendChild oPIDChildLevel1
    
    Set oPIDChildLevel1 = oDom.createNode(NODE_ELEMENT, "PID.19", "")
        oPIDChildLevel1.Text = gobjPatient.SSN
    oPIDNode.appendChild oPIDChildLevel1
    
    strTemp = NullToNothing(gobjPatient.Race)
    If strTemp = "" Then
        lngNCID = 0
    Else: lngNCID = CLng(strTemp) 'gobjPatient.Race
    End If
        
    '<< Begin:SCR #26532;   Developer: Brian Mowbray 08/12/2002 02:27 PM
    If geDestination = eDM_EIC Then
    
      strSQL = "SELECT SAMS_ETHNICITY_NCID, SAMS_ETHNICITY_DESCRIPTION FROM "
      strSQL = strSQL & "TMIP_SAMS_NCID_MAPPING WHERE "
      strSQL = strSQL & "CHCSII_NCID = " & lngNCID
      'This applies to both TMIP and SAMS
  
      Set RS = New ADODB.Recordset
      Set RS = oDas.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)
      If Not RS Is Nothing Then
          If Not RS.EOF Then
              strSQL = NullToNothing(RS.Fields(0).Value)  'strSQL is the ncid
              strTemp = NullToNothing(RS.Fields(1).Value) 'strtemp is the description
          Else
              '<SCR 22468 CC>
              strSQL = ""
              strTemp = ""
          End If
      Else: strTemp = ""
          strSQL = ""
          '</SCR 22468 CC>
      End If
    
    Else
      If lngNCID <> 0 Then
        strTemp = GetNCIDDescription(lngNCID)
        strSQL = CStr(lngNCID)
      End If
    End If
    '>> End: SCR #26532;
    
    '<NGIT SPR 423-1 CC>
    Set oPIDChildLevel1 = oDom.createNode(NODE_ELEMENT, "PID.22.LST", "")
    
        Set oPIDChildLevel2 = oDom.createNode(NODE_ELEMENT, "PID.22", "")
            Set oPIDChildLevel3 = oDom.createNode(NODE_ELEMENT, "CE.1", "")
            oPIDChildLevel3.Text = ""           'strTemp is the description
            oPIDChildLevel2.appendChild oPIDChildLevel3
            Set oPIDChildLevel3 = oDom.createNode(NODE_ELEMENT, "CE.2", "")
            oPIDChildLevel3.Text = strTemp 'Already description
            oPIDChildLevel2.appendChild oPIDChildLevel3
            Set oPIDChildLevel3 = oDom.createNode(NODE_ELEMENT, "CE.3", "")
            oPIDChildLevel3.Text = ""
            oPIDChildLevel2.appendChild oPIDChildLevel3
            Set oPIDChildLevel3 = oDom.createNode(NODE_ELEMENT, "CE.4", "")
            oPIDChildLevel3.Text = strSQL       'strSQL is the NCID
            oPIDChildLevel2.appendChild oPIDChildLevel3
            Set oPIDChildLevel3 = oDom.createNode(NODE_ELEMENT, "CE.5", "")
            oPIDChildLevel3.Text = strTemp      'strTemp is the description
            oPIDChildLevel2.appendChild oPIDChildLevel3
            Set oPIDChildLevel3 = oDom.createNode(NODE_ELEMENT, "CE.6", "")
            oPIDChildLevel3.Text = "99NCI"
            oPIDChildLevel2.appendChild oPIDChildLevel3
                
        oPIDChildLevel1.appendChild oPIDChildLevel2
    oPIDNode.appendChild oPIDChildLevel1
    '</NGIT SPR 423-1 CC>
    
    Set ReturnNewPIDNode = oPIDNode

Exit Function

ErrTrap:
   If gobjShared Is Nothing Then
      Set gobjShared = New CWShared
   End If
   gobjShared.ShowVBError Err.Number, Err.Description, "ReturnNewPIDNode Function", "DMBE", vbExclamation
   ''Err.Clear
   Resume Next 'There is no reason this procedure should not continue to create xml if it errors

End Function

Public Function ReturnNewCLNOBRNode(ByVal lngCount As Long, ByVal rsOBR As ADODB.Recordset) As MSXML.IXMLDOMElement
         
' Observation Request
'
' sample function output
'   <OBR>
'      <OBR.1>1</OBR.1>
'      <OBR.3>
'         <EI.1></EI.1>
'      </OBR.3>
'      <OBR.4>
'         <CE.1></CE.1>
'         <CE.2>CLINICAL_NOTES</CE.2>
'         <CE.3></CE.3>
'         <CE.4></CE.4>
'         <CE.5></CE.5>
'         <CE.6>99NCI</CE.6>
'      </OBR.4>
'   </OBR>

    Dim oDom As MSXML.DOMDocument
    Dim oOBRNode As MSXML.IXMLDOMElement
    Dim oOBRChildLevel1 As MSXML.IXMLDOMElement
    Dim oOBRChildLevel2 As MSXML.IXMLDOMElement
    
    On Error GoTo ErrTrap
    
    Set oDom = New MSXML.DOMDocument
    Set oOBRNode = oDom.createElement("OBR")
    
    'Set Order Number
    Set oOBRChildLevel1 = oDom.createElement("OBR.1")
    oOBRChildLevel1.Text = lngCount
    oOBRNode.appendChild oOBRChildLevel1
    
    Set oOBRChildLevel1 = oDom.createElement("OBR.3")
        ' Placer Order Number - Table Clinnote:DATAID
        Set oOBRChildLevel2 = oDom.createElement("EI.1")
        oOBRChildLevel2.Text = NullToNothing(rsOBR.Fields("DATAID").Value)
        oOBRChildLevel1.appendChild oOBRChildLevel2
        
        Set oOBRChildLevel2 = oDom.createElement("EI.2")
        oOBRChildLevel2.Text = "CHCSII-T"
        oOBRChildLevel1.appendChild oOBRChildLevel2
        
        Set oOBRChildLevel2 = oDom.createElement("EI.3")
        oOBRChildLevel2.Text = ""
        oOBRChildLevel1.appendChild oOBRChildLevel2
        
        Set oOBRChildLevel2 = oDom.createElement("EI.4")
        oOBRChildLevel2.Text = ""
        oOBRChildLevel1.appendChild oOBRChildLevel2
    oOBRNode.appendChild oOBRChildLevel1
    
    ' Universal Service ID
    Set oOBRChildLevel1 = oDom.createElement("OBR.4")
        ' Identifier
        Set oOBRChildLevel2 = oDom.createElement("CE.1")
        oOBRChildLevel2.Text = "" 'NullToNothing(rsOBR.Fields("DATAID").Value)
        oOBRChildLevel1.appendChild oOBRChildLevel2
        'Text
        Set oOBRChildLevel2 = oDom.createElement("CE.2")
        oOBRChildLevel2.Text = "CLINICAL_NOTES"
        oOBRChildLevel1.appendChild oOBRChildLevel2
        ' Name of Coding System -
        Set oOBRChildLevel2 = oDom.createElement("CE.3")
        oOBRChildLevel2.Text = ""
        oOBRChildLevel1.appendChild oOBRChildLevel2
        ' Alternate Identifier
        Set oOBRChildLevel2 = oDom.createElement("CE.4")
        oOBRChildLevel2.Text = ""
        oOBRChildLevel1.appendChild oOBRChildLevel2
        ' Alternate Text
        Set oOBRChildLevel2 = oDom.createElement("CE.5")
        oOBRChildLevel2.Text = ""
        oOBRChildLevel1.appendChild oOBRChildLevel2
        ' Alternate Name of Coding System
        Set oOBRChildLevel2 = oDom.createElement("CE.6")
        oOBRChildLevel2.Text = ""
        oOBRChildLevel1.appendChild oOBRChildLevel2
    oOBRNode.appendChild oOBRChildLevel1
    
    Set ReturnNewCLNOBRNode = oOBRNode

Exit Function

ErrTrap:
   If gobjShared Is Nothing Then
      Set gobjShared = New CWShared
   End If
   gobjShared.ShowVBError Err.Number, Err.Description, "ReturnNewCLNOBRNode Function", "DMBE", vbExclamation
   'Err.Clear
   Resume Next 'There is no reason this procedure should not continue on to create xml

End Function

Public Function ReturnNewCLNOBXNode(ByVal lngCount As Long, rsClinicalNotes As ADODB.Recordset) As MSXML.IXMLDOMElement

' sample function output
'   <OBX>
'      <OBX.1>1</OBX.1>
'      <OBX.2>ST</OBX.2>
'      <OBX.3>
'         <CE.1></CE.1>
'         <CE.2>CLINICAL_NOTES_TEXT</CE.2>
'         <CE.3></CE.3>
'         <CE.4></CE.4>
'         <CE.5></CE.5>
'         <CE.6></CE.6>
'      </OBX.3>
'      <OBX.5.LST>
'         <OBX.5>{\rtf1\ansi\ansicpg1252\uc0\deff0{\fonttbl
            '{\f0\fswiss\fcharset0\fprq2 Arial;}
            '{\f1\froman\fcharset2\fprq2 Symbol;}}
            '{\colortbl;\red0\green0\blue0;\red255\green255\blue255;}
            '{\info{\comment TX_RTF32 8.0.301.500}}
            '\deftab1134\pard\plain\f0\fs16 The patient has been examined is fit to return to normal duty activities.\par }</OBX.5>
'      </OBX.5.LST>
'      <OBX.11>F</OBX.11>
'   </OBX>

    Dim oDom As MSXML.DOMDocument
    Dim oOBXNode As MSXML.IXMLDOMElement
    Dim oOBXChildLevel1 As MSXML.IXMLDOMElement
    Dim oOBXChildLevel2 As MSXML.IXMLDOMElement
    Dim strTemp As String
    
    On Error GoTo ErrTrap
    
    Set oDom = New MSXML.DOMDocument
    Set oOBXNode = oDom.createElement("OBX")
    
    'Set Order Number
    Set oOBXChildLevel1 = oDom.createElement("OBX.1")
    oOBXChildLevel1.Text = lngCount
    oOBXNode.appendChild oOBXChildLevel1

    ' Value Type - ST for Clinical Notes
    Set oOBXChildLevel1 = oDom.createElement("OBX.2")
    oOBXChildLevel1.Text = "ST"
    oOBXNode.appendChild oOBXChildLevel1

    ' Observation Identifier
    Set oOBXChildLevel1 = oDom.createElement("OBX.3")
        ' Observation Identifier
        Set oOBXChildLevel2 = oDom.createElement("CE.1")
        oOBXChildLevel2.Text = ""
        oOBXChildLevel1.appendChild oOBXChildLevel2
        ' Text
        Set oOBXChildLevel2 = oDom.createElement("CE.2")
        oOBXChildLevel2.Text = "CLINICAL_NOTES_TEXT"
        oOBXChildLevel1.appendChild oOBXChildLevel2
        ' Name of Coding System
        Set oOBXChildLevel2 = oDom.createElement("CE.3")
        oOBXChildLevel2.Text = ""
        oOBXChildLevel1.appendChild oOBXChildLevel2
        ' Alternate Identifier
        Set oOBXChildLevel2 = oDom.createElement("CE.4")
        oOBXChildLevel2.Text = ""
        oOBXChildLevel1.appendChild oOBXChildLevel2
        ' Alternate Text - Table WhatRS:
        Set oOBXChildLevel2 = oDom.createElement("CE.5")
        oOBXChildLevel2.Text = ""
        oOBXChildLevel1.appendChild oOBXChildLevel2
        ' Alternate Name of Coding System
        Set oOBXChildLevel2 = oDom.createElement("CE.6")
        oOBXChildLevel2.Text = ""
        oOBXChildLevel1.appendChild oOBXChildLevel2
    oOBXNode.appendChild oOBXChildLevel1
        
    ' Observation SubID
    Set oOBXChildLevel1 = oDom.createElement("OBX.4")
    oOBXChildLevel1.Text = CStr(lngCount)
    oOBXNode.appendChild oOBXChildLevel1
    
    Set oOBXChildLevel1 = oDom.createElement("OBX.5.LST")
        ' Observation Identifier
        Set oOBXChildLevel2 = oDom.createElement("OBX.5")
        If IsNull(rsClinicalNotes.Fields("HTMLTEXT").Value) Then
            strTemp = ""
        Else
            strTemp = FilterRtfToXml(rsClinicalNotes.Fields("HTMLTEXT").Value)
        End If
        oOBXChildLevel2.Text = strTemp
        oOBXChildLevel1.appendChild oOBXChildLevel2
    oOBXNode.appendChild oOBXChildLevel1
        
    Set oOBXChildLevel1 = oDom.createElement("OBX.11")
    If IsNull(rsClinicalNotes.Fields("STATUS").Value) Then
        strTemp = "X"
    Else
        If rsClinicalNotes.Fields("STATUS").Value = "Resolved" Then
            strTemp = "F"
        Else
            strTemp = "P"
        End If
    End If
    oOBXChildLevel1.Text = strTemp
    oOBXNode.appendChild oOBXChildLevel1
        
    Set ReturnNewCLNOBXNode = oOBXNode
    
Exit Function

ErrTrap:
   If gobjShared Is Nothing Then
      Set gobjShared = New CWShared
   End If
   gobjShared.ShowVBError Err.Number, Err.Description, "ReturnNewCLNOBXNode Function", "DMBE", vbExclamation
   'Err.Clear
   Resume Next 'There is no reason this procedure should not continue to create xml if it errors
   
End Function

Private Function GetNCIDDescription(ByVal varNCID As Variant, Optional ByVal blnIsFacilityNCID As Boolean = False) As String
   
   'SCR 25083 - Added optional parameter
   
   Dim varTemp As Variant
   Dim lngTemp As Long
   Dim strOutPut As String
   
   strOutPut = ""
   
   If varNCID <> "" Then
      lngTemp = varNCID
      If lngTemp <> 0 Then
         If oConCept Is Nothing Then
            Set oConCept = New GEMS_ConceptCtrl
         End If
         oConCept.UniqueID = lngTemp
         varTemp = oConCept.PrefRep("2000").Representation
         If IsNull(varTemp) Or IsEmpty(varTemp) Then
            strOutPut = ""
         Else
            strOutPut = Trim(varTemp)
         End If
      End If
    End If
   
    If blnIsFacilityNCID = True And strOutPut = "" Then
        strOutPut = "CHCSII-T ADT"
    End If
   
    GetNCIDDescription = strOutPut
   
End Function


Private Function NullToNothing(ByVal varPassIn As Variant) As String

    If IsNull(varPassIn) Then
        NullToNothing = ""
    Else
        NullToNothing = varPassIn
    End If
    
End Function

Private Sub GetNewDoctorInfo(ClinicianNCID As Long, _
                           DocLastName As String, _
                           DocFirstName As String, _
                           DocMiddleName As String, _
                           DocID As String, _
                           DocSpecialty As String)
   
   Dim RS As ADODB.Recordset
   Dim vTemp As Variant
   Dim sSQL As String
   Dim sData() As String
   
   sSQL = "SELECT * FROM PROVIDER WHERE NCID = '" & ClinicianNCID & "'"
   Set RS = GetTableDataDAS(sSQL)
   If RS Is Nothing Then
      DocLastName = ""
      DocFirstName = ""
      DocMiddleName = ""
      DocID = ""
      Exit Sub
   End If
   If RS.BOF = False Then
      If RS.EOF = False Then
         RS.MoveFirst
         vTemp = RS.Fields("SSN").Value
         If IsNull(vTemp) Then
            DocID = ""
         Else
            DocID = Trim(vTemp)
         End If
         ' Begin:SCR #34092;   Developer: Mag 03/27/2003 09:59 AM
         vTemp = RS.Fields("SPECIALTY_CODE").Value
         ' End: SCR #34092;
         If IsNull(vTemp) Then
            DocSpecialty = ""
         Else
            DocSpecialty = Trim(vTemp)
            If IsNumeric(DocSpecialty) Then
            DocSpecialty = GetNCIDDescription(CLng(DocSpecialty))
            End If
         End If
         vTemp = RS.Fields("NAME").Value
         If IsNull(vTemp) Then
            DocLastName = ""
            DocFirstName = ""
            DocMiddleName = ""
         Else
            DocLastName = Trim(vTemp)
            DocLastName = Replace(DocLastName, ",", " ", 1, -1, vbTextCompare)
            DocLastName = Replace(DocLastName, "  ", " ", 1, -1, vbTextCompare)
            sData = Split(DocLastName, " ", -1, vbTextCompare)
            If UBound(sData) > 0 Then
               DocLastName = sData(0)
            Else
               DocLastName = ""
            End If
            If UBound(sData) >= 1 Then
               DocFirstName = sData(1)
            Else
               DocFirstName = ""
            End If
            If UBound(sData) = 2 Then
               DocMiddleName = sData(2)
            Else
               DocMiddleName = ""
            End If
         End If
      End If
   End If
   Set RS = Nothing

End Sub

Private Function GetMedicationsStatusText(ByVal lngData As Long) As String
   
   ' this should be coming from a DB table and NOT hard coded in this function or in Medications
   Const NCID_ORDERSTATUS_SCHEDULED = "1555"
   Const NCID_ORDERSTATUS_ACTIVE = "1024"
   'Public Const NCID_INACTIVE = "1025"
   Const NCID_ORDERSTATUS_DISCONTINUED = "1524"
   Const NCID_ONHOLD = "163002"
   'Public Const NCID_CANCELLED = 11981
   Const NCID_ORDERSTATUS_EXPIRED = "206497"
   Const NCID_Expired = "14516641"
   Const NCID_PARTIAL = "14516644"
   Dim DataStr As String
   
   Select Case lngData
      Case NCID_ORDERSTATUS_SCHEDULED
         DataStr = "SCHEDULED"
      Case NCID_ORDERSTATUS_ACTIVE
         DataStr = "ACTIVE"
      Case NCID_ORDERSTATUS_DISCONTINUED
         DataStr = "DISCONTINUED"
      Case NCID_ONHOLD
         DataStr = "ONHOLD"
      Case NCID_ORDERSTATUS_EXPIRED
         DataStr = "EXPIRED"
      Case NCID_Expired
         DataStr = "NCID EXPIRED"
      Case NCID_PARTIAL
         DataStr = "PARTIAL"
      Case Else
         DataStr = ""
   End Select
   
   GetMedicationsStatusText = DataStr

End Function

Private Function GetMedicationNameText(ByVal Data As Double) As String

'<SCR 23902 CC> - Change from long to double

   Dim sSQL As String
   Dim RS As ADODB.Recordset
   Dim vTemp As Variant
   
   sSQL = "select * from rx_list where drug_ien = ?0"
   sSQL = Replace(sSQL, "?0", CStr(Data), 1, -1, vbTextCompare)
   Set RS = GetTableDataDAS(sSQL)
   If RS.RecordCount <> 0 Then
      vTemp = RS.Fields("DRUG_NAME").Value
      If IsNull(vTemp) = True Then
         vTemp = ""
      Else
         GetMedicationNameText = vTemp
      End If
   Else
      GetMedicationNameText = ""
      Exit Function
   End If
End Function

Public Function DischargeDispositionText(ByVal vNCID As Variant) As String
   DischargeDispositionText = DischargeDisposition(vNCID)
End Function
