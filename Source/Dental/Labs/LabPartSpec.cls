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
Attribute VB_Name = "clsLabPartSpec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================================================================
'  Class Name: clsLabPartSpec
'
'================================================================================================
Option Explicit

'////////////////////  C L A S S   P R O P E R T I E S    S E C T I O N  ////////////////////

'LAB_TASK_CATEGORY_DEF table                    'Table Fields
Private msTaskCategoryDefinition As String      'Definition
Private msAlias As String                       'Alias
Private msEffectiveDate As String               'EffectiveDate
Private msEffectiveEndDate As String            'EffectiveEndDate

'PK/FK
Private msLabTaskCategoryID As String           'Lab_Task_Category_ID

'LAB_PART_SPECIFICATIONS table                  'Table Fields
Private msLabPartSpecificationsID As String     'Lab_Part_Specifications_ID
Private msLabTxID As String                     'LAB_TX_ID
Private msLabMaterialElementsID As String       'Lab_Material_Elements_ID
Private msLabMaterialID As String               'Lab_Material_ID
Private msLineItemToothLocation As String       'Line_Item_Tooth_Location

Private miLineItemNumber As Integer             'Line_Item_Number
Private msLabMaterialCategory As String         'Lab_Material_Category

'Note
Private msLabNote As String

'RPD GUI variables
Private mbPreserveFacebrow              As Boolean
Private mbRemountCastsAndMountMaxillary As Boolean
Private mbLabRemount                    As Boolean
Private mbLabOcclusalAdjustment         As Boolean
Private mbCustomCharacterization        As Boolean
Private mbStippling                     As Boolean

'link to the object LabMaterials which has private collection of LabMaterial
Private mobjLabMaterials As clsLabMaterials

'link to the LabNote object
Private mobjLabNote As clsLabNote

Private Sub Class_Initialize()

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
Set mobjLabNote = New clsLabNote

End Sub

Public Property Get LabMaterials() As clsLabMaterials
    Set LabMaterials = mobjLabMaterials
End Property

Public Property Set LabMaterials(robjValue As clsLabMaterials)
    Set mobjLabMaterials = robjValue
End Property

Public Property Get LabNote() As clsLabNote
    Set LabNote = mobjLabNote
End Property

Public Property Set LabNote(robjValue As clsLabNote)
    Set mobjLabNote = robjValue
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

Public Property Let LabMaterialCategory(ByVal vsValue As String)
    msLabMaterialCategory = vsValue
End Property

Public Property Get LabMaterialCategory() As String
    LabMaterialCategory = msLabMaterialCategory
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

Public Property Let LineItemNumber(ByVal viValue As Integer)
    miLineItemNumber = viValue
End Property

Public Property Get LineItemNumber() As Integer
    LineItemNumber = miLineItemNumber
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

Select Case giLabCaseStatus
   Case gcCreateNewLabCase
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
       miLineItemNumber = roRs("Lab_Item_Number") & ""
       msLabMaterialCategory = roRs("Lab_Material_Category") & ""
   
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
       miLineItemNumber = roRs("Line_Item_Number") & ""
       msLabMaterialCategory = roRs("Lab_Material_Category") & ""
      
End Select

Exit Sub

ErrHandler:
    Resume Next
   Call gobjShared.ShowVBError(Err.Number, Err.Description, _
                "clsLabPartSpecs.FillFromRecordset", "Labs", vbCritical)
End Sub

Public Function GetLabPartSpecChildMembers(ByRef roCurLabPartSpec As clsLabPartSpec) As Boolean
''===========================================================================================
''  Description: This method gets the recordset of the Part Spec related Materials
''               and creates the collection of the Material items
''
''  Called By: this method is called by GetLabTXChildMembers in clsLabTX code
''
''  Argument(s):
''         vValue - index of the TX item in the TXs collection
''
''  Returns:
''         TRUE  - if all treatment related tasks data successfully retrieved from database
''         FALSE - if failed
''
''  Notes:
''
''  Programmer:    Alek Samedov
''===========================================================================================

Dim oRs As Recordset
Dim i As Integer

On Error GoTo ErrHandler

GetLabPartSpecChildMembers = True

If roCurLabPartSpec.LabMaterials.GetLabMaterials(roCurLabPartSpec, oRs) Then
   oRs.MoveFirst
   i = 1
   Do While Not oRs.EOF
      roCurLabPartSpec.LabMaterials.Add oRs
            'to get lab materials elements from LAB_PART_SPECIFICATIONS table
            If roCurLabPartSpec.LabMaterials.Item(i).GetLabMaterialChildMembers(roCurLabPartSpec, i) = False Then
               GetLabPartSpecChildMembers = False
               Exit Function
            End If
      oRs.MoveNext
      i = i + 1
   Loop
Else
   'Err.Raise vbObjectError, "", "Collection is not defined"
   GetLabPartSpecChildMembers = False
   Exit Function
End If
    
Exit Function

ErrHandler:
GetLabPartSpecChildMembers = False

   Call gobjShared.ShowVBError(Err.Number, Err.Description, _
                "clsLabPartSpec.GetLabTXChildMembers", "Labs", vbCritical)

End Function

Private Sub Class_Terminate()
   Set mobjLabMaterials = Nothing
   Set mobjLabNote = Nothing
End Sub

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
'                "clsLabTask.FillToRecordset", "Labs", vbCritical)
'End Sub


