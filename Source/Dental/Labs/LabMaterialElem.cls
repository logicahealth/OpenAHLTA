VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLabMaterialElem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================================================================
'  Class Name: clsLabMaterialElem
'
'================================================================================================
Option Explicit

'////////////////////  C L A S S   P R O P E R T I E S    S E C T I O N  ////////////////////

'LAB_MATERIAL_ELEMENTS table                  'Table Fields
Private msLabMaterialElementsID As String     'Lab_Material_Elements_ID
Private msLabMaterialID As String             'Lab_Material_ID
Private msAbbrev As String                    'Abbrev
Private msDescription As String               'Description
Private msActive As String                    'Active

'LAB_PART_SPECIFICATIONS table
'Private miLabPartSpecificationsID As Integer
Private msLabPartSpecificationsID As String

'Private miLabTxID As Integer
Private msLabTxID As String

'Private miLabTaskCategoryID As Integer
Private msLabTaskCategoryID As String

Private msLineItemToothLocation As String
Private miLineItemNumber As Integer           'Line_Item_Number
'to distinguish RPD Framework from Pontics
Private msLabMaterialCategory As String       'Lab_Material_Category: Framework or Pontics grid

'Update/Insert status
Private msUpdateStatus As String

Public Property Let LabMaterialElementsID(ByVal vsValue As String)
    msLabMaterialElementsID = vsValue
End Property

Public Property Get LabMaterialElementsID() As String
    LabMaterialElementsID = msLabMaterialElementsID
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

Public Property Let LabTaskCategoryID(ByVal vsValue As String)
    msLabTaskCategoryID = vsValue
End Property

Public Property Get LabTaskCategoryID() As String
    LabTaskCategoryID = msLabTaskCategoryID
End Property

Public Property Let Abbrev(ByVal vsValue As String)
    msAbbrev = vsValue
End Property

Public Property Get Abbrev() As String
    Abbrev = msAbbrev
End Property

Public Property Let Description(ByVal vsValue As String)
    msDescription = vsValue
End Property

Public Property Get Description() As String
    Description = msDescription
End Property

Public Property Let Active(ByVal vsValue As String)
    msActive = vsValue
End Property

Public Property Get Active() As String
    Active = msActive
End Property

Public Property Let UpdateStatus(ByVal vsValue As String)
    msUpdateStatus = vsValue
End Property

Public Property Get UpdateStatus() As String
    UpdateStatus = msUpdateStatus
End Property
Public Property Let LabMaterialID(ByVal vsValue As String)
    msLabMaterialID = vsValue
End Property

Public Property Get LabMaterialCategory() As String
    LabMaterialCategory = msLabMaterialCategory
End Property

Public Property Let LabMaterialCategory(ByVal vsValue As String)
    msLabMaterialCategory = vsValue
End Property

Public Property Get LabMaterialID() As String
    LabMaterialID = msLabMaterialID
End Property

Public Property Let LineItemNumber(ByVal viValue As Integer)
    miLineItemNumber = viValue
End Property

Public Property Get LineItemNumber() As Integer
    LineItemNumber = miLineItemNumber
End Property

Public Property Let LineItemToothLocation(ByVal vsValue As String)
    msLineItemToothLocation = vsValue
End Property

Public Property Get LineItemToothLocation() As String
    LineItemToothLocation = msLineItemToothLocation
End Property

Public Sub FillFromRecordset(roRs As Recordset)
''===========================================================================================
''  Description: This method populates fields properties with data from recordset fields
''
''  Called By: this method is called by Add method in clsLabMaterials class code
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

On Error GoTo ErrHandler

    '--- fill in the properties from the LAB_MATERIALS table fields ---
      msLabMaterialElementsID = CStr(roRs("Lab_Material_Elements_ID")) & ""
      msLabMaterialID = roRs("Lab_Material_ID") & ""
      msAbbrev = roRs("Abbrev") & ""
      msDescription = roRs("Description") & ""
      msActive = roRs("Active") & ""
      
      msLabPartSpecificationsID = CStr(roRs("Lab_Part_Specifications_ID")) & ""
      msLabTxID = CStr(roRs("Lab_TX_ID")) & ""
      msLabTaskCategoryID = CStr(roRs("Lab_Task_Category_ID")) & ""
      msLineItemToothLocation = CStr(roRs("Line_Item_Tooth_Location")) & ""
      miLineItemNumber = roRs("Line_Item_Number")
      msLabMaterialCategory = roRs("Lab_Material_Category") & ""
      
Exit Sub

ErrHandler:
    Resume Next
   Call gobjShared.ShowVBError(Err.Number, Err.Description, _
                "clsLabMaterial.FillFromRecordset", "Labs", vbCritical)
End Sub

