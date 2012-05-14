VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLabMaterial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================================================================
'  Class Name: clsLabMaterial
'
'  Description: This class is used by Dental Labs component to
'               describe, store and retrieve data for/from LabMaterial object
'
'================================================================================================
Option Explicit

'////////////////////  C L A S S   P R O P E R T I E S    S E C T I O N  ////////////////////

'LAB_MATERIALS table                          'Table Fields
Private msLabMaterialID            As String  'Lab_Material_ID PK & FK
Private msAbbrev                   As String  'Abbrev
Private msDescription              As String  'Description
Private msActive                   As String  'Active

'LAB_PART_SPECIFICATIONS table                'Table Fields
Private msLabPartSpecificationsID As String   'Lab_Part_Specifications_ID
Private msLabTxID                 As String   'LAB_TX_ID
Private msLabTaskCategoryID       As String   'Lab_Task_Category_ID
Private msLabMaterialElemsID      As String   'Lab_Material_Elements_ID
Private msLineItemToothLocation   As String   'Line_Item_Tooth_Location


'link to the object LabMaterialElems which has private collection of LabMaterialElem
Private mobjLabMaterialElems As clsLabMaterialElems

Public Property Set LabMaterialElems(rcolValue As clsLabMaterialElems)
    Set mobjLabMaterialElems = rcolValue
End Property

Public Property Get LabMaterialElems() As clsLabMaterialElems
    Set LabMaterialElems = mobjLabMaterialElems
End Property

Public Property Let LabMaterialID(ByVal vsValue As String)
    msLabMaterialID = vsValue
End Property

Public Property Get LabMaterialID() As String
    LabMaterialID = msLabMaterialID
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
    LabTaskCategoryID = vsValue
End Property

Public Property Get LabTaskCategoryID() As String
    LabTaskCategoryID = msLabTaskCategoryID
End Property

Public Property Let LabMaterialElemsID(ByVal vsValue As String)
    LabMaterialElemsID = vsValue
End Property

Public Property Get LabMaterialElemsID() As String
    LabMaterialElemsID = msLabMaterialElemsID
End Property

Public Property Let LineItemToothLocation(ByVal vsValue As String)
    LineItemToothLocation = vsValue
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
Select Case giLabCaseStatus
   Case gcCreateNewLabCase
      msLabMaterialID = roRs("Lab_Material_ID") & ""
      msAbbrev = roRs("Abbrev") & ""
      msDescription = roRs("Description") & ""
      msActive = roRs("Active") & ""
      
   Case Else
      msLabMaterialID = roRs("Lab_Material_ID") & ""
      msAbbrev = roRs("Abbrev") & ""
      msDescription = roRs("Description") & ""
      msActive = roRs("Active") & ""
      
End Select

Exit Sub

ErrHandler:
    Resume Next
   Call gobjShared.ShowVBError(Err.Number, Err.Description, _
                "clsLabMaterial.FillFromRecordset", "Labs", vbCritical)
End Sub

Private Sub Class_Initialize()
    Set mobjLabMaterialElems = New clsLabMaterialElems
End Sub

Private Sub Class_Terminate()
   Set mobjLabMaterialElems = Nothing
End Sub

Public Function GetLabMaterialChildMembers(ByRef robjLabPartSpec As clsLabPartSpec, _
                                           ByVal vMaterialIndex As Variant _
                                           ) As Boolean
''===========================================================================================
''  Description: This method gets the recordset of the Lab Material related Elements
''               and creates the collection of the Elements items
''
''  Called By:   This method is called by GetLabPartSpecChildMembers
''               in cls clsLabPartSpec code
''
''  Argument(s):
''         robjLabPartSpec         - reference to the LabPartSpec object
''         vMaterialIndex  - index of the Lab Material item in the LabMaterials collection
''
''  Returns:
''         TRUE  - if all LabMaterial related Elements data successfully retrieved from database
''         FALSE - if failed
''
''  Notes: Material Child Members are lab material elements
''
''  Programmer:    Alek Samedov
''===========================================================================================

Dim oLabMaterials As clsLabMaterials
Dim sLabMaterialID As String
Dim oRs As Recordset

On Error GoTo ErrHandler

GetLabMaterialChildMembers = True

'set reference to existing object
Set oLabMaterials = robjLabPartSpec.LabMaterials

sLabMaterialID = oLabMaterials.Item(vMaterialIndex).LabMaterialID

If oLabMaterials.Item(vMaterialIndex).LabMaterialElems.GetLabMaterialElems(sLabMaterialID, oRs) Then

   oRs.MoveFirst
   Do While Not oRs.EOF
      oLabMaterials.Item(vMaterialIndex).LabMaterialElems.Add oRs
      oRs.MoveNext
   Loop
End If

'release reference
If Not oLabMaterials Is Nothing Then Set oLabMaterials = Nothing

Exit Function

ErrHandler:
GetLabMaterialChildMembers = False
'release reference
If Not oLabMaterials Is Nothing Then Set oLabMaterials = Nothing

   Call gobjShared.ShowVBError(Err.Number, Err.Description, _
                "clsLabMaterial.GetLabMaterialChildMembers", "Labs", vbCritical)

End Function
'
'
'Public Function GetLabMaterialChildMembers(ByRef robjLabPartSpec As clsLabPartSpec, _
'                                           ByVal vMaterialIndex As Variant _
'                                           ) As Boolean
'''===========================================================================================
'''  Description: This method gets the recordset of the Lab Material related Elements
'''               and creates the collection of the Elements items
'''
'''  Called By:   This method is called by GetLabPartSpecChildMembers
'''               in cls clsLabPartSpec code
'''
'''  Argument(s):
'''         robjLabPartSpec         - reference to the LabPartSpec object
'''         vMaterialIndex  - index of the Lab Material item in the LabMaterials collection
'''
'''  Returns:
'''         TRUE  - if all LabMaterial related Elements data successfully retrieved from database
'''         FALSE - if failed
'''
'''  Notes: Material Child Members are lab material elements
'''
'''  Programmer:    Alek Samedov
'''===========================================================================================
'
'Dim oLabMaterials As clsLabMaterials
'Dim sLabMaterialID As String
'Dim oRs As Recordset
'
'On Error GoTo ErrHandler
'
'GetLabMaterialChildMembers = True
'
''set reference to existing object
'Set oLabMaterials = robjLabPartSpec.LabMaterials
'
'sLabMaterialID = oLabMaterials.Item(vMaterialIndex).LabMaterialID
'
''If oLabMaterials.Item(vMaterialIndex).LabMaterialElems.GetLabMaterialElems(sLabMaterialID, oRs) Then
''If gPartSpecMaterialElemsRS Is Nothing Then
'   If oLabMaterials.Item(vMaterialIndex).LabMaterialElems.GetLabMaterialElems(gPartSpecMaterialElemsRS) Then
''End If
'
'   gPartSpecMaterialElemsRS.MoveFirst  'find!!!
'   Do While Not gPartSpecMaterialElemsRS.EOF
'      If CInt(gPartSpecMaterialElemsRS("Lab_Material_ID")) = CInt(oLabMaterials.Item(vMaterialIndex).LabMaterialID) _
'         And CInt(gPartSpecMaterialElemsRS("Lab_Task_Category_ID")) = CInt(gobjCurLabPartSpec.LabTaskCategoryID) _
'         And CInt(gPartSpecMaterialElemsRS("Lab_TX_ID")) = CInt(gobjCurLabTX.LabTxID) Then
'         oLabMaterials.Item(vMaterialIndex).LabMaterialElems.Add gPartSpecMaterialElemsRS
'      End If
'      If CInt(gPartSpecMaterialElemsRS("Lab_Material_ID")) > CInt(oLabMaterials.Item(vMaterialIndex).LabMaterialID) Then
'         'Ctrl Break
'         Exit Do
'      End If
'      gPartSpecMaterialElemsRS.MoveNext
'   Loop
'
' End If
'
''release reference
'If Not oLabMaterials Is Nothing Then Set oLabMaterials = Nothing
'
'Exit Function
'
'ErrHandler:
'GetLabMaterialChildMembers = False
''release reference
'If Not oLabMaterials Is Nothing Then Set oLabMaterials = Nothing
'
'   Call gobjShared.ShowVBError(Err.Number, Err.Description, _
'                "clsLabMaterial.GetLabMaterialChildMembers", "Labs", vbCritical)
'
'End Function

'Public Function GetLookupMaterialChildMembers(ByVal vMaterialIndex As Variant) As Boolean
'''===========================================================================================
'''  Description: This method gets the recordset of the Lab Material related Elements
'''               and creates the collection of the Elements items
'''
'''  Called By:   This method is called by GetLabDentSpecChildMembers
'''               in cls clsLabDentSpec code
'''
'''  Argument(s):
'''         vMaterialIndex - LabMaterialID
'''
'''  Returns:
'''         TRUE  - if all LabMaterial related Elements data successfully retrieved from database
'''         FALSE - if failed
'''
'''  Notes:
'''
'''  Programmer:    Alek Samedov
'''===========================================================================================
'
'Dim oRs As Recordset
'
'On Error GoTo ErrHandler
'
'GetLookupMaterialChildMembers = True
'
'If gobjLookupLabMaterials Is Nothing Then
'   Set gobjLookupLabMaterials = gobjLabCase.LabDentSpec.LabMaterials
'End If
'
''Set gobjLookupLabMaterials = gobjLabCase.LabDentSpec.LabMaterials
'
'With gobjLookupLabMaterials.Item(vMaterialIndex).LabMaterialElems
'   If .GetLookupMaterialElems(vMaterialIndex, oRs) Then
'      oRs.MoveFirst
'      Do While Not oRs.EOF
'         .Add oRs
'         oRs.MoveNext
'        Loop
'
'    Else
'        GetLookupMaterialChildMembers = False
'        Exit Function
'    End If
' End With
'
'Exit Function
'
'ErrHandler:
''If Not gobjLookupLabMaterials Is Nothing Then Set gobjLookupLabMaterials = Nothing
'
'GetLookupMaterialChildMembers = False
'
'   Call gobjShared.ShowVBError(Err.Number, Err.Description, _
'                "clsLabTX.GetLookupMaterialChildMembers", "Labs", vbCritical)
'
'End Function

'Public Function GetLookupMaterialChildMembersNew(ByVal vMaterialIndex As Variant) As Boolean
'''===========================================================================================
'''  Description: This method gets the recordset of the Lab Material related Elements
'''               and creates the collection of the Elements items
'''
'''  Called By:   This method is called by GetLabDentSpecChildMembers
'''               in cls clsLabDentSpec code
'''
'''  Argument(s):
'''         vTXIndex       - index of the Lab TX in the collection of the treatment items
'''         vMaterialIndex - LabMaterialID
'''
'''  Returns:
'''         TRUE  - if all LabMaterial related Elements data successfully retrieved from database
'''         FALSE - if failed
'''
'''  Notes:
'''
'''  Programmer:    Alek Samedov
'''===========================================================================================
'
'On Error GoTo ErrHandler
'
'GetLookupMaterialChildMembersNew = True
'
'If Not gobjLookupLabMaterials Is Nothing Then
'   Set gobjLookupLabMaterials = Nothing
'End If
'
'Set gobjLookupLabMaterials = gobjLabCase.LabDentSpec.LabMaterials
'
'With gobjLookupLabMaterials.Item(vMaterialIndex).LabMaterialElems
'   gLookupLabMaterialElemsRS.MoveFirst
'
'   Do While Not gLookupLabMaterialElemsRS.EOF
'      If CInt(gLookupLabMaterialElemsRS("Lab_Material_ID")) = CInt(vMaterialIndex) Then
'         .Add gLookupLabMaterialElemsRS
'      End If
'      gLookupLabMaterialElemsRS.MoveNext
'   Loop
'
' End With
'
'Exit Function
'
'ErrHandler:
'If Not gobjLookupLabMaterials Is Nothing Then Set gobjLookupLabMaterials = Nothing
'
'GetLookupMaterialChildMembersNew = False
'
'   Call gobjShared.ShowVBError(Err.Number, Err.Description, _
'                "clsLabMaterial.GetLookupMaterialChildMembersNew", "Labs", vbCritical)
'
'End Function

'Public Function GetLookupMaterialChildMembersBackup(ByVal vTXIndex, _
'                                              ByVal vMaterialIndex As Variant) As Boolean
'''===========================================================================================
'''  Description: This method gets the recordset of the Lab Material related Elements
'''               and creates the collection of the Elements items
'''
'''  Called By:   This method is called by GetLabDentSpecChildMembers
'''               in cls clsLabDentSpec code
'''
'''  Argument(s):
'''         vTXIndex       - index of the Lab TX in the collection of the treatment items
'''         vMaterialIndex - LabMaterialID
'''
'''  Returns:
'''         TRUE  - if all LabMaterial related Elements data successfully retrieved from database
'''         FALSE - if failed
'''
'''  Notes:
'''
'''  Programmer:    Alek Samedov
'''===========================================================================================
'
'Dim oRs As Recordset
'
'On Error GoTo ErrHandler
'
'GetLookupMaterialChildMembersBackup = True
'
'If Not gobjLookupLabMaterials Is Nothing Then
'   Set gobjLookupLabMaterials = Nothing
'End If
'
''<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
''Set gobjLookupLabMaterials = gobjLabCase.LabDentSpec.LabMaterials
''<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'
''<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'Set gobjLookupLabMaterials = gobjLabCase.LabTXs.Item(vTXIndex).LabDentSpec.LabMaterials
''<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'
'With gobjLookupLabMaterials.Item(vMaterialIndex).LabMaterialElems
'   If .GetLookupMaterialElems(vMaterialIndex, oRs) Then
'      oRs.MoveFirst
'      Do While Not oRs.EOF
'         .Add oRs
'         oRs.MoveNext
'        Loop
'
'    Else
'        'Err.Raise vbObjectError, "", "Collection is not defined"
'        GetLookupMaterialChildMembersBackup = False
'        Exit Function
'    End If
' End With
'
''release reference
''Set gobjLookupLabMaterials = Nothing
'
'Exit Function
'
'ErrHandler:
'If Not gobjLookupLabMaterials Is Nothing Then Set gobjLookupLabMaterials = Nothing
'
'GetLookupMaterialChildMembersBackup = False
'
'   Call gobjShared.ShowVBError(Err.Number, Err.Description, _
'                "clsLabTX.GetLookupMaterialChildMembersBackup", "Labs", vbCritical)
'
'End Function
'

