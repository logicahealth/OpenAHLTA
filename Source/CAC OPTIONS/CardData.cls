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
Attribute VB_Name = "CardData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private msSSN As String                         'Card Holder SSN
Private msSponsorSSN As String                  'Sponsor SSN
Private msVersionCode As String                 'Card Version
Private msPersonDesignatorIdentifier As String
Private msPersonDesignatorTypeCode As String
Private msPersonelCategory As String
Private msSPDI As String
Private msFirstName As String
Private msSurname As String
Private msDOB As String
Private msPayGrade As String
Private msPayPlanCode As String
Private msPayPlanGradeCode As String
Private msPersonnelEntitlementConditionType As String
Private msBranchCode As String
Private msRank As String
Private msServiceCode As String
Private msCardIssueDate As String
Private msCardExpirationDate As String
Private msCardInstanceIdentifier As String
Private msCardSecurityCode As String
Private msStatus As String
Private msDDS_DDSC As String
Private msDF As String
Private msGender As String
Private msMedBenExpirationDate As String

Public Property Let MedBenExpirationDate(vData As String)
    msMedBenExpirationDate = vData
End Property

Public Property Get MedBenExpirationDate() As String
    MedBenExpirationDate = msMedBenExpirationDate
End Property

Public Property Let Gender(vData As String)
    msGender = vData
End Property

Public Property Get Gender() As String
    Gender = msGender
End Property

Public Property Let SSN(vData As String)
    msSSN = vData
End Property

Public Property Get SSN() As String
    SSN = msSSN
End Property

Public Property Let SponsorSSN(vData As String)
    msSponsorSSN = vData
End Property

Public Property Get SponsorSSN() As String
    SponsorSSN = msSponsorSSN
End Property

Public Property Let VersionCode(vData As String)
    msVersionCode = vData
End Property

Public Property Get VersionCode() As String
    VersionCode = msVersionCode
End Property

Public Property Let PersonDesignatorIdentifier(vData As String)
    msPersonDesignatorIdentifier = vData
End Property

Public Property Get PersonDesignatorIdentifier() As String
    PersonDesignatorIdentifier = msPersonDesignatorIdentifier
End Property

Public Property Let PersonDesignatorTypeCode(vData As String)
    msPersonDesignatorTypeCode = vData
End Property

Public Property Get PersonDesignatorTypeCode() As String
    PersonDesignatorTypeCode = msPersonDesignatorTypeCode
End Property

Public Property Let PersonelCategory(vData As String)
    msPersonelCategory = vData
End Property

Public Property Get PersonelCategory() As String
    PersonelCategory = msPersonelCategory
End Property

'Sponsor's PDI
Public Property Let SPDI(vData As String)
    msSPDI = vData
End Property

'Sponsor's PDI
Public Property Get SPDI() As String
    SPDI = msSPDI
End Property

Public Property Let FirstName(vData As String)
    msFirstName = vData
End Property

Public Property Get FirstName() As String
    FirstName = msFirstName
End Property

Public Property Let Surname(vData As String)
    msSurname = vData
End Property

Public Property Get Surname() As String
    Surname = msSurname
End Property

Public Property Let DateOfBirth(vData As String)
    msDOB = vData
End Property

Public Property Get DateOfBirth() As String
    DateOfBirth = msDOB
End Property

Public Property Let PayGrade(vData As String)
    msPayGrade = vData
End Property

Public Property Get PayGrade() As String
    PayGrade = msPayGrade
End Property

Public Property Let PayPlanCode(vData As String)
    msPayPlanCode = vData
End Property

Public Property Get PayPlanCode() As String
    PayPlanCode = msPayPlanCode
End Property

Public Property Let PayPlanGradeCode(vData As String)
    msPayPlanGradeCode = vData
End Property

Public Property Get PayPlanGradeCode() As String
    PayPlanGradeCode = msPayPlanGradeCode
End Property

Public Property Let PersonnelEntitlementConditionType(vData As String)
    msPersonnelEntitlementConditionType = vData
End Property

Public Property Get PersonnelEntitlementConditionType() As String
    PersonnelEntitlementConditionType = msPersonnelEntitlementConditionType
End Property

Public Property Let BranchCode(vData As String)
    msBranchCode = vData
End Property

Public Property Get BranchCode() As String
    BranchCode = msBranchCode
End Property

Public Property Let Rank(vData As String)
    msRank = vData
End Property

Public Property Get Rank() As String
    Rank = msRank
End Property

'Service/Component Code
Public Property Let ServiceCode(vData As String)
    msServiceCode = vData
End Property

'Service/Component Code
Public Property Get ServiceCode() As String
    ServiceCode = msServiceCode
End Property

Public Property Let CardIssueDate(vData As String)
    msCardIssueDate = vData
End Property

Public Property Get CardIssueDate() As String
    CardIssueDate = msCardIssueDate
End Property

Public Property Let CardExpirationDate(vData As String)
    msCardExpirationDate = vData
End Property

Public Property Get CardExpirationDate() As String
    CardExpirationDate = msCardExpirationDate
End Property

Public Property Let CardInstanceIdentifier(vData As String)
    msCardInstanceIdentifier = vData
End Property

Public Property Get CardInstanceIdentifier() As String
    CardInstanceIdentifier = msCardInstanceIdentifier
End Property

Public Property Let CardSecurityCode(vData As String)
    msCardSecurityCode = vData
End Property

Public Property Get CardSecurityCode() As String
    CardSecurityCode = msCardSecurityCode
End Property

Public Property Let Status(vData As String)
    msStatus = vData
End Property

Public Property Get Status() As String
    Status = msStatus
End Property

'DEERS Dependent Suffix/DMDC Dependent Suffix Code
Public Property Let DDS_DDSC(vData As String)
    msDDS_DDSC = vData
End Property

'DEERS Dependent Suffix/DMDC Dependent Suffix Code
Public Property Get DDS_DDSC() As String
    DDS_DDSC = msDDS_DDSC
End Property

'DDSC Flag
Public Property Let DF(vData As String)
    msDF = vData
End Property

'DDSC Flag
Public Property Get DF() As String
    DF = msDF
End Property

Public Function Serialize() As String
    Serialize = XMLSerialize_Deserialize(Me, "", False)
End Function

Public Function Deserialize(XMLString As String)
    XMLSerialize_Deserialize Me, XMLString, False
End Function

Private Function XMLSerialize_Deserialize(ByVal Object As Object, Optional XMLString As String = "", Optional ByVal _
    AsAttributes = True, Optional ByVal OmitProperties As String = "XMLProperties") As String
  Dim tTLI      As TLIApplication
  Dim tMem      As MemberInfo
  Dim tDom      As DOMDocument
  Dim tNode     As IXMLDOMNode
  Dim tInvoke   As InvokeKinds
  Dim tOmit     As String
  Dim tName     As String     'used as lower case....
  Dim tString   As String
On Error GoTo ErrorHandler

  Set tTLI = New TLIApplication
  Set tDom = New DOMDocument
  
  If Len(XMLString) Then
  
  'If string given, then we are letting new property values from xmlstring
      tInvoke = VbLet
      tDom.loadXML (XMLString)
  
  Else
  
  'Else we are getting existing property values
      tInvoke = VbGet
      tDom.appendChild tDom.createNode(NODE_ELEMENT, _
          TypeName(Object), "")
  
  End If
  
  tOmit = "," & LCase(OmitProperties) & ","
  
  'handle each get or let member from object
  
  For Each tMem In _
      tli.InterfaceInfoFromObject(Object).Members
      
      tName = LCase(tMem.Name)
      'tName = tMem.Name
      
      '... get or let and not omitted property
      '    for example object etc...
      
      If tMem.InvokeKind = tInvoke And InStr(tOmit, "," & tName & ",") = 0 And tMem.Parameters.Count = 0 Then
           
          On Error Resume Next  'could be object or
              ' something else that can't handle
          If tInvoke = VbGet Then
             '... put data to XML-node
             If AsAttributes Then
               Set tNode = tDom.createAttribute(tName)
               tNode.Text = CallByName(Object, tMem.Name, _
                   VbGet)
               tDom.documentElement.Attributes.setNamedItem _
                   tNode
             Else
               Set tNode = tDom.createElement(tName)
               tNode.Text = CallByName(Object, tMem.Name, _
                   VbGet)
               tDom.documentElement.appendChild tNode
             End If
          Else
             '... get data from XML-node
             If AsAttributes Then
               CallByName Object, tMem.Name, VbLet, _
                 tDom.documentElement.Attributes.getNamedItem(tName).Text
             Else
             
                If tMem.Name = "SSN" And _
                        (tDom.documentElement.selectSingleNode("ssn") Is Nothing) Then
                       Object.SSN = tDom.documentElement.selectSingleNode("PatientSSN").Text
               Else
                   CallByName Object, tMem.Name, VbLet, _
                     tDom.documentElement.selectSingleNode(tName).Text
               End If
             End If
          End If
          On Error GoTo 0
      End If
  Next
  
  XMLSerialize_Deserialize = tDom.xml

    Exit Function
ErrorHandler:
    Err.Clear
End Function

