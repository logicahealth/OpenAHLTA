VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLabDentSpec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================================================================
'  Class Name: clsLabDentSpec
'
'  Description: This abstract class is used by Dental Lab component
''              to store lookup values
'
'================================================================================================
Option Explicit

'////////////////////  C L A S S   P R O P E R T I E S    S E C T I O N  ////////////////////

'LAB_TASK_CATEGORY_DEF table                    'Table Fields
Private msLabTaskCategoryID As String           'Lab_Task_Category_ID (PK/FK)
Private msTaskCategoryDefinition As String      'Definition
Private msAlias As String                       'Alias
Private msEffectiveDate As String               'EffectiveDate
Private msEffectiveEndDate As String            'EffectiveEndDate

'LAB_PART_SPECIFICATIONS table                  'Table Fields
Private msLabPartSpecificationsID As String     'Lab_Part_Specifications_ID
Private msLabTxID As String                     'LAB_TX_ID
Private msLabMaterialElementsID As String       'Lab_Material_Elements_ID
Private msLabMaterialID As String               'Lab_Material_ID
Private msLineItemToothLocation As String       'Line_Item_Tooth_Location

'Note
Private msLabNote As String

'RPD GUI variables

Private msCaseNo As String
Private msToothLocation As String

Private msFrameworkMaterial As String
Private msVeneerMaterial As String
Private msGuide As String
Private msShade As String
Private msOcclusalSurfaceMaterial As String
Private msOcclusalScheme As String
Private msMargin As String
Private msMProximalContact As String
Private msDProximalContact As String
Private msGuidePlane As String
Private msRest1 As String
Private msRest2 As String
Private msUndercutLocation As String
Private msUndercutAmount As String
Private msAttachment As String
Private msAttachmentLocation As String
Private msMetalFinish As String
Private msVeneerFinish As String

Private msMConnector As String
Private msDConnector As String
Private msPonticType As String
Private msDistanceGingiva As String
Private msTissueContaceArea As String
Private msTissueContactMaterial As String

Private msCheckAt As String
Private msReturnAt As String

'link to the object LabMaterials which has private collection of LabMaterial
Private mobjLabMaterials As clsLabMaterials

Private Sub Class_Initialize()

'LAB_TASK_CATEGORY_DEF table
 msLabTaskCategoryID = ""
 msTaskCategoryDefinition = ""
 msAlias = ""
 msEffectiveDate = ""
 msEffectiveEndDate = ""

'LAB_PART_SPECIFICATIONS table
 msLabPartSpecificationsID = ""
 msLabTxID = ""
 msLabTaskCategoryID = ""
 msLabMaterialElementsID = ""
 msLabMaterialID = ""
 msLineItemToothLocation = ""
 
 Set mobjLabMaterials = New clsLabMaterials

'GUI
msCaseNo = ""
msToothLocation = ""

msLabNote = ""

msFrameworkMaterial = ""
msVeneerMaterial = ""
msGuide = ""
msShade = ""
msOcclusalSurfaceMaterial = ""
msOcclusalScheme = ""
msMargin = ""
msMProximalContact = ""
msDProximalContact = ""
msGuidePlane = ""
msRest1 = ""
msRest2 = ""
msUndercutLocation = ""
msUndercutAmount = ""
msAttachment = ""
msAttachmentLocation = ""
msMetalFinish = ""
msVeneerFinish = ""

msMConnector = ""
msDConnector = ""
msPonticType = ""
msDistanceGingiva = ""
msTissueContaceArea = ""
msTissueContactMaterial = ""

msCheckAt = ""
msReturnAt = ""

End Sub

Public Property Get LabMaterials() As clsLabMaterials
    Set LabMaterials = mobjLabMaterials
End Property

Public Property Set LabMaterials(robjValue As clsLabMaterials)
    Set mobjLabMaterials = robjValue
End Property

Public Property Let LabTaskCategoryID(ByVal vlValue As Long)
    msLabTaskCategoryID = vlValue
End Property

Public Property Get LabTaskCategoryID() As Long
    LabTaskCategoryID = msLabTaskCategoryID
End Property

Public Property Let TaskCategoryDefinition(ByVal vsValue As String)
    msTaskCategoryDefinition = vsValue
End Property

Public Property Get TaskCategoryDefinition() As String
    TaskCategoryDefinition = msTaskCategoryDefinition
End Property

Public Property Let Alias(ByVal vsValue As String)
    msAlias = vsValue
End Property

Public Property Get Alias() As String
    Alias = msAlias
End Property

Public Property Let LabPartSpecificationsID(ByVal vsValue As String)
    msLabPartSpecificationsID = vsValue
End Property

Public Property Get LabPartSpecificationsID() As String
    LabPartSpecificationsID = msLabPartSpecificationsID
End Property

Public Property Let LabTxID(ByVal vsValue As String)
    msLabTxID = vsValue
End Property

Public Property Get LabTxID() As String
    LabTxID = msLabTxID
End Property

Public Property Let LabMaterialID(ByVal vsValue As String)
    msLabMaterialID = vsValue
End Property

Public Property Get LabMaterialID() As String
    LabMaterialID = msLabMaterialID
End Property

Public Property Let LabMaterialElementsID(ByVal vsValue As String)
    msLabMaterialElementsID = vsValue
End Property

Public Property Get LabMaterialElementsID() As String
    LabMaterialElementsID = msLabMaterialElementsID
End Property

Public Property Let LineItemToothLocation(ByVal vsValue As String)
    msLineItemToothLocation = vsValue
End Property

Public Property Get LineItemToothLocation() As String
    LineItemToothLocation = msLineItemToothLocation
End Property

Public Property Let EffectiveDate(ByVal vsValue As String)
    msEffectiveDate = vsValue
End Property

Public Property Get EffectiveDate() As String
    EffectiveDate = msEffectiveDate
End Property

Public Property Let EffectiveEndDate(ByVal vsValue As String)
    msEffectiveEndDate = vsValue
End Property

Public Property Get EffectiveEndDate() As String
    EffectiveEndDate = msEffectiveEndDate
End Property

Public Property Let CaseNo(ByVal vsValue As String)
    msCaseNo = vsValue
End Property

Public Property Get CaseNo() As String
    CaseNo = msCaseNo
End Property

Public Property Let ToothLocation(ByVal vsValue As String)
    msToothLocation = vsValue
End Property

Public Property Get ToothLocation() As String
    ToothLocation = msToothLocation
End Property

Public Property Let LabNote(ByVal vsValue As String)
    msLabNote = vsValue
End Property

Public Property Get LabNote() As String
    LabNote = msLabNote
End Property

Public Property Let FrameworkMaterial(ByVal vsValue As String)
    msFrameworkMaterial = vsValue
End Property

Public Property Get FrameworkMaterial() As String
    FrameworkMaterial = msFrameworkMaterial
End Property

Public Property Let VeneerMaterial(ByVal vsValue As String)
    msVeneerMaterial = vsValue
End Property

Public Property Get VeneerMaterial() As String
    VeneerMaterial = msVeneerMaterial
End Property

Public Property Let Guide(ByVal vsValue As String)
    msGuide = vsValue
End Property

Public Property Get Guide() As String
    Guide = msGuide
End Property

Public Property Let Shade(ByVal vsValue As String)
    msShade = vsValue
End Property

Public Property Get Shade() As String
    Shade = msShade
End Property

Public Property Let OcclusalSurfaceMaterial(ByVal vsValue As String)
    msOcclusalSurfaceMaterial = vsValue
End Property

Public Property Get OcclusalSurfaceMaterial() As String
    OcclusalSurfaceMaterial = msOcclusalSurfaceMaterial
End Property

Public Property Let OcclusalScheme(ByVal vsValue As String)
    msOcclusalScheme = vsValue
End Property

Public Property Get OcclusalScheme() As String
    OcclusalScheme = msOcclusalScheme
End Property

Public Property Let Margin(ByVal vsValue As String)
    msMargin = vsValue
End Property

Public Property Get Margin() As String
    Margin = msMargin
End Property

Public Property Let MProximalContact(ByVal vsValue As String)
    msMProximalContact = vsValue
End Property

Public Property Get MProximalContact() As String
    MProximalContact = msMProximalContact
End Property

Public Property Let DProximalContact(ByVal vsValue As String)
    msDProximalContact = vsValue
End Property

Public Property Get DProximalContact() As String
    DProximalContact = msDProximalContact
End Property

Public Property Let GuidePlane(ByVal vsValue As String)
    msGuidePlane = vsValue
End Property

Public Property Get GuidePlane() As String
    GuidePlane = msGuidePlane
End Property

Public Property Let Rest1(ByVal vsValue As String)
    msRest1 = vsValue
End Property

Public Property Get Rest1() As String
    Rest1 = msRest1
End Property

Public Property Let Rest2(ByVal vsValue As String)
    msRest2 = vsValue
End Property

Public Property Get Rest2() As String
    Rest2 = msRest1
End Property

Public Property Let UndercutLocation(ByVal vsValue As String)
    msUndercutLocation = vsValue
End Property

Public Property Get UndercutLocation() As String
    UndercutLocation = msUndercutLocation
End Property

Public Property Let UndercutAmount(ByVal vsValue As String)
    msUndercutAmount = vsValue
End Property

Public Property Get UndercutAmount() As String
    UndercutAmount = msUndercutAmount
End Property

Public Property Let Attachment(ByVal vsValue As String)
    msAttachment = vsValue
End Property

Public Property Get Attachment() As String
    Attachment = msAttachment
End Property

Public Property Let AttachmentLocation(ByVal vsValue As String)
    msAttachmentLocation = vsValue
End Property

Public Property Get AttachmentLocation() As String
    AttachmentLocation = msAttachmentLocation
End Property

Public Property Let MetalFinish(ByVal vsValue As String)
    msMetalFinish = vsValue
End Property

Public Property Get MetalFinish() As String
    MetalFinish = msMetalFinish
End Property

Public Property Let VeneerFinish(ByVal vsValue As String)
    msVeneerFinish = vsValue
End Property

Public Property Get VeneerFinish() As String
    VeneerFinish = msVeneerFinish
End Property

Public Property Let MConnector(ByVal vsValue As String)
    msMConnector = vsValue
End Property

Public Property Get MConnector() As String
    MConnector = msMConnector
End Property

Public Property Let DConnector(ByVal vsValue As String)
    msDConnector = vsValue
End Property

Public Property Get DConnector() As String
    DConnector = msDConnector
End Property

Public Property Let PonticType(ByVal vsValue As String)
    msPonticType = vsValue
End Property

Public Property Get PonticType() As String
    PonticType = msPonticType
End Property

Public Property Let DistanceGingiva(ByVal vsValue As String)
    msDistanceGingiva = vsValue
End Property

Public Property Get DistanceGingiva() As String
    DistanceGingiva = msDistanceGingiva
End Property

Public Property Let TissueContaceArea(ByVal vsValue As String)
    msTissueContaceArea = vsValue
End Property

Public Property Get TissueContaceArea() As String
    TissueContaceArea = msTissueContaceArea
End Property

Public Property Let TissueContactMaterial(ByVal vsValue As String)
    msTissueContactMaterial = vsValue
End Property

Public Property Get TissueContactMaterial() As String
    TissueContactMaterial = msTissueContactMaterial
End Property

Public Property Let CheckAt(ByVal vsValue As String)
    msCheckAt = vsValue
End Property

Public Property Get CheckAt() As String
    CheckAt = msCheckAt
End Property

Public Property Let ReturnAt(ByVal vsValue As String)
    msReturnAt = vsValue
End Property

Public Property Get ReturnAt() As String
    ReturnAt = msReturnAt
End Property

Public Sub FillFromRecordset(roRs As Recordset)
''===========================================================================================
''  Description: This method populates fields properties with data from recordset fields
''
''  Called By: this method is called by Add method in clsLabPartSpecs class code
''
''  Argument(s):
''     roRs - Contains recordset
''
''  Returns:
''
''  Notes:
''
''  Programmer:    Alek Samedov
''===========================================================================================

'--- error handling ---
On Error GoTo ErrHandler

    '--- fill in the properties from the LAB_PART_SPECIFICATIONS table fields ---

Select Case giLabCaseStatus
   Case gcCreateNewLabCase
      'LAB_TASK_CATEGORY_DEF table
       msLabTaskCategoryID = CStr(roRs("Lab_Task_Category_ID")) & ""
       msTaskCategoryDefinition = roRs("Definition") & ""
       msAlias = roRs("Alias") & ""
       msEffectiveDate = CDate(roRs("EffectiveDate")) & ""
       msEffectiveEndDate = CDate(roRs("EffectiveEndDate")) & ""
      
      
      'LAB_PART_SPECIFICATIONS table
       'msLabPartSpecificationsID = CStr(roRs("Lab_Part_Specifications_ID")) & ""
       msLabTxID = CStr(roRs("LAB_TX_ID")) & ""
       msLabTaskCategoryID = CStr(roRs("Lab_Task_Category_ID")) & ""
       msLabMaterialElementsID = CStr(roRs("Lab_Material_Elements_ID")) & ""
       msLabMaterialID = CStr(roRs("Lab_Material_ID")) & ""
       msLineItemToothLocation = CStr(roRs("Line_Item_Tooth_Location")) & ""
   
   Case Else
      'update status
      'LAB_TASK_CATEGORY_DEF table
       msLabTaskCategoryID = CStr(roRs("Lab_Task_Category_ID")) & ""
       msTaskCategoryDefinition = roRs("Definition") & ""
       msAlias = roRs("Alias") & ""
       msEffectiveDate = CDate(roRs("EffectiveDate")) & ""
       msEffectiveEndDate = CDate(roRs("EffectiveEndDate")) & ""
      
      
      'LAB_PART_SPECIFICATIONS table
       msLabPartSpecificationsID = CStr(roRs("Lab_Part_Specifications_ID")) & ""
       msLabTxID = CStr(roRs("LAB_TX_ID")) & ""
       msLabTaskCategoryID = CStr(roRs("Lab_Task_Category_ID")) & ""
       msLabMaterialElementsID = CStr(roRs("Lab_Material_Elements_ID")) & ""
       msLabMaterialID = CStr(roRs("Lab_Material_ID")) & ""
       msLineItemToothLocation = CStr(roRs("Line_Item_Tooth_Location")) & ""
      
End Select

Exit Sub

ErrHandler:
    Resume Next
   Call gobjShared.ShowVBError(Err.Number, Err.Description, _
                "clsLabDentSpec.FillFromRecordset", "Labs", vbCritical)
End Sub

Public Function GetLabDentSpecChildMembers() As Boolean
''===========================================================================================
''  Description: This method gets the recordset of the Part Spec related Materials
''               and creates the collection of the Material and Material Elements items
''
''  Called By: this method is called by
''                                     - GetLabCaseChildMembers in clsLabCase code
''
''  Argument(s):
''
''  Returns:
''         TRUE  - if all Lab Materials lookup data successfully retrieved from database
''         FALSE - if failed
''
''  Notes:
''
''  Programmer:    Alek Samedov
''===========================================================================================

Dim oLabMaterials As clsLabMaterials

Dim oRs As Recordset
Dim bResult As Boolean
Dim i As Integer

On Error GoTo ErrHandler

GetLabDentSpecChildMembers = True

'--- reference to the object
Set oLabMaterials = gobjLabCase.LabDentSpec.LabMaterials

If gLookupLabMaterialsRS Is Nothing Then
   bResult = oLabMaterials.GetLookupMaterials(oRs)
Else
   '--- gLookupLabMaterialsRS recordset is not Nothing
   '--- (already exist)
   bResult = True
   Set oRs = gLookupLabMaterialsRS
End If

If bResult Then
   If gLookupLabMaterialElemsRS Is Nothing Then
      bResult = GetDentMaterialElems
   Else
      '--- gLookupLabMaterialElemsRS recordset is not Nothing
      bResult = True
   End If
End If

If bResult Then
  
   If gobjLookupLabMaterials Is Nothing Then
      Set gobjLookupLabMaterials = gobjLabCase.LabDentSpec.LabMaterials
   End If
  
  i = 1
  oRs.MoveFirst
  Do While Not oRs.EOF
    oLabMaterials.Add oRs
         
    'gLookupLabMaterialElemsRS.MoveFirst
    gLookupLabMaterialElemsRS.Find ("Lab_Material_ID = " & "'" & oLabMaterials.Item(i).LabMaterialID & "'")
    
    Do While Not gLookupLabMaterialElemsRS.EOF
      If CInt(gLookupLabMaterialElemsRS("Lab_Material_ID")) = CInt(oLabMaterials.Item(i).LabMaterialID) Then
         oLabMaterials.Item(i).LabMaterialElems.Add gLookupLabMaterialElemsRS
      End If
      If CInt(gLookupLabMaterialElemsRS("Lab_Material_ID")) > CInt(oLabMaterials.Item(i).LabMaterialID) Then
         'Control break
         Exit Do
      End If
      gLookupLabMaterialElemsRS.MoveNext
    Loop

    oRs.MoveNext
    i = i + 1

  Loop
         
Else
    GetLabDentSpecChildMembers = False
    Exit Function
End If
       
'release reference
If Not oLabMaterials Is Nothing Then Set oLabMaterials = Nothing
Set oRs = Nothing

Exit Function

ErrHandler:
GetLabDentSpecChildMembers = False
'release reference
If Not oLabMaterials Is Nothing Then Set oLabMaterials = Nothing

   Call gobjShared.ShowVBError(Err.Number, Err.Description, _
                "clsLabDentSpec.GetLabDentSpecChildMembers", "Labs", vbCritical)

End Function

Private Sub Class_Terminate()
   Set mobjLabMaterials = Nothing
End Sub

'Public Function GetLabDentSpecChildMembersBackup() As Boolean
'''===========================================================================================
'''  Description: This method gets the recordset of the Part Spec related Materials
'''               and creates the collection of the Material items
'''
'''  Called By: this method is called by
'''                                     - GetLabCaseChildMembers in clsLabCase code
'''
'''  Argument(s):
'''
'''  Returns:
'''         TRUE  - if all Lab Materials lookup data successfully retrieved from database
'''         FALSE - if failed
'''
'''  Notes:
'''
'''  Programmer:    Alek Samedov
'''===========================================================================================
'
'Dim oLabMaterials As clsLabMaterials
'Dim oRs As Recordset
'Dim bResult As Boolean
'
'On Error GoTo ErrHandler
'
'GetLabDentSpecChildMembers = True
'
''reference to the object
'Set oLabMaterials = gobjLabCase.LabDentSpec.LabMaterials
'
'If gLookupLabMaterialsRS Is Nothing Then
'   bResult = oLabMaterials.GetLookupMaterials(oRs)
'Else
'   'gLookupLabMaterialsRS recordset is not Nothing
'   '(already exist)
'   bResult = True
'   Set oRs = gLookupLabMaterialsRS
'End If
'
'If bResult Then
'   If gLookupLabMaterialElemsRS Is Nothing Then
'      bResult = GetDentMaterialElems
'   Else
'      'gLookupLabMaterialElemsRS recordset is not Nothing
'      bResult = True
'   End If
'End If
'
'If bResult Then
'
'  Dim i
'  i = 1
'  oRs.MoveFirst
'  Do While Not oRs.EOF
'    oLabMaterials.Add oRs
'
'    'retrieve MaterialElements data
'    If oLabMaterials.Item(i).GetLookupMaterialChildMembers(i) = False Then
'       GetLabDentSpecChildMembers = False
'       Exit Function
'    End If
'
'    oRs.MoveNext
'    i = i + 1
'  Loop
'
'Else
'    GetLabDentSpecChildMembers = False
'    Exit Function
'End If
'
''release reference
'If Not oLabMaterials Is Nothing Then Set oLabMaterials = Nothing
'Set oRs = Nothing
'Exit Function
'
'ErrHandler:
'GetLabDentSpecChildMembers = False
''release reference
'If Not oLabMaterials Is Nothing Then Set oLabMaterials = Nothing
'
'   Call gobjShared.ShowVBError(Err.Number, Err.Description, _
'                "clsLabDentSpec.GetLabDentSpecChildMembers", "Labs", vbCritical)
'
'End Function
'

'Public Sub FillToRecordset(roRs As Recordset)
'''===========================================================================================
'''  Description: This method populates fields of the recordset with data from properties
'''
'''  Called By: this method is called by ... in ... code
'''
'''  Argument(s):
'''     roRs - Contains recordset
'''
'''  Returns:
'''
'''  Notes:
'''
'''  Programmer:    Alek Samedov
'''===========================================================================================
'
'On Error GoTo ErrHandler
'
''assigned values come from Tasks flex grid columns !!!
'
'    '--- fill in the ... table fields from the properties
'''    roRs("LAB_CASE_NUMBER") = msLabCaseNo
'''    roRs("Date_CASE_COMPLETED") = msDateCaseCompleted
'''    roRs("Status") = msStatus
'''    roRs("Comments") = msComments
'''    roRs("ProviderID") = msProviderId
'''    roRs("CreateDateTime") = msCreateDateTime
'''    roRs("DENTAL_ENCOUNTER_ID") = mlEncounterNumber
'''    roRs("LINKED") = mlCaseLinkedTo
'
'Exit Sub
'
'ErrHandler:
'    Resume Next
'   'MsgBox Err.Description, vbExclamation, "Error " & Err.Number & " in " & gsProcName
'   Call gobjShared.ShowVBError(Err.Number, Err.Description, _
'                "clsLabDentSpec.FillToRecordset", "Labs", vbCritical)
'End Sub


