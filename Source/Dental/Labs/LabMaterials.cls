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
Attribute VB_Name = "clsLabMaterials"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================================================================
'  Class Name: clsLabMaterials
'
'================================================================================================
Option Explicit

'////////////////////  C L A S S   P R O P E R T I E S    S E C T I O N  ////////////////////

Private mcolLabMaterials As Collection

'Methods of the LabMaterials class.

Public Function Add(ByRef oRst As Recordset) As clsLabMaterial
Dim oLabMaterial As New clsLabMaterial
Dim sLabMaterialID As String
        
On Error GoTo ErrHandler

oLabMaterial.FillFromRecordset oRst

'Add the LabMaterial object reference to the
'collection, using the ID property as the key.
        
sLabMaterialID = oLabMaterial.LabMaterialID
If sLabMaterialID <> "" Then
   mcolLabMaterials.Add oLabMaterial, sLabMaterialID
Else
   mcolLabMaterials.Add oLabMaterial
End If
            
' Return a reference to the new LabMaterial
Set Add = oLabMaterial
    
Exit Function

ErrHandler:
Resume Next
   Call gobjShared.ShowVBError(Err.Number, Err.Description, _
                "clsLabMaterials.Add", "Labs", vbCritical)
    
End Function

Public Function Count() As Long
    Count = mcolLabMaterials.Count
End Function

Public Sub Delete(ByVal vIndex As Variant)
    mcolLabMaterials.Remove vIndex
End Sub

Public Function Item(ByVal vIndex As Variant) As clsLabMaterial
Attribute Item.VB_UserMemId = 0
    Set Item = mcolLabMaterials.Item(vIndex)
End Function

'NewEnum must return the IUnknown interface of a
'collection's enumerator
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mcolLabMaterials.[_NewEnum]
End Function

Private Sub Class_Initialize()
On Error GoTo ErrHandler
'---set/initialize all data members

Set mcolLabMaterials = New Collection

Exit Sub
ErrHandler:
   Call gobjShared.ShowVBError(Err.Number, Err.Description, _
                "clsLabMaterials.Initialize", "Labs", vbCritical)
End Sub

Public Function GetLabMaterials(ByRef roCurLabPartSpec As clsLabPartSpec, _
                                 ByRef robjRS As Recordset) As Boolean
''===========================================================================================
''  Description: This method gets the recordset with LabMaterials data from database
''
''  Called By:   This function is called by GetLabPartSpecChildMembers method
''               in clsLabPartSpec code
''
''  Argument(s):
''     roCurLabPartSpec - reference to the current LabPartSpec object
''     robjRS           - Recordset
''
''  Returns:
''         TRUE  - if success
''         FALSE - if failed
''
''  Notes:
''
''  Programmer:    Alek Samedov
''===========================================================================================
Dim sSql As String
Dim sControlName As String
Dim sCategoryType As String

'--- error handling ---

On Error GoTo ErrHandler

GetLabMaterials = True

Select Case roCurLabPartSpec.TaskCategoryDefinition
   Case gcCrown
      sCategoryType = gcCrown
   Case gcFPD
      sCategoryType = gcFPD
   Case gcRPD
      sCategoryType = gcRPD
   Case Else
   '
End Select

sSql = GetSQLLabMaterials(sCategoryType)

If LabDataProcessing(sSql, robjRS) = gcRowsSelected Then
   GetLabMaterials = True
Else
   GetLabMaterials = False
End If

Exit Function

ErrHandler:
GetLabMaterials = False

   Call gobjShared.ShowVBError(Err.Number, Err.Description, _
                "modLabsData.GetLabMaterials", "Labs", vbCritical)

End Function

Private Sub Class_Terminate()
   'Set mcolLabMaterials = Nothing
End Sub

Public Function GetLookupMaterials(ByRef robjRS As Recordset) As Boolean
''===========================================================================================
''  Description: This method gets the recordset with LabMaterials data from database
''
''  Called By:   This function is called by GetLabDentSpecChildMembers method
''               in clsLabDentSpec code
''
''  Argument(s):
''               robjRS - Recordset
''
''  Returns:
''         TRUE  - if success
''         FALSE - if failed
''
''  Notes:
''
''  Programmer:    Alek Samedov
''===========================================================================================
Dim sSql As String

'--- error handling ---

On Error GoTo ErrHandler

GetLookupMaterials = True

'lab Material records are getting from Lab_Materials lookup table
sSql = "SELECT Lab_Material_ID, Abbrev, Description, Active " _
      & "FROM LAB_MATERIALS ORDER BY Lab_Material_ID"
        
If LabDataProcessing(sSql, robjRS) = gcRowsSelected Then
   GetLookupMaterials = True
   Set gLookupLabMaterialsRS = robjRS
Else
   GetLookupMaterials = False
End If

Exit Function

ErrHandler:
GetLookupMaterials = False

   'MsgBox Err.Description, vbExclamation, "Error " & Err.Number & " in " & gsProcName
   Call gobjShared.ShowVBError(Err.Number, Err.Description, _
                "modLabsData.GetLookupMaterials", "Labs", vbCritical)

End Function

Public Function GetSQLLabMaterials(ByVal vValue As Variant) As String
''===========================================================================================
''  Description: This method created SQL to read Lab Material IDs from database
''               in accordance with Material ID values presented in GUI controls
''               (Material IDs are in tags for comboboxes, and textboxes,
''               and in column data for grids)
''
''  Called By:   This function gets called in GetLabMaterials method in clsLabMaterials code
''
''  Argument(s):
''     vValue - Contains abbrev of control name, i.e. bear information about
''              lab part specifications
''
''
''  Returns:
''          GetSQLLabMaterials - contains SQL string
''
''  Notes:
''
''  Programmer:    Alek Samedov
''===========================================================================================
Dim sSql As String
Dim iControl As Control
Dim i As Integer

'--- error handling ---
On Error GoTo ErrHandler

sSql = "SELECT Lab_Material_ID, Abbrev, Description, Active " _
        & "FROM LAB_MATERIALS WHERE Lab_Material_ID = -1"

With frmLabs
   Select Case vValue
      Case gcCrown
         For Each iControl In .cboCrownMaterial
            sSql = sSql & " OR " & "Lab_Material_ID = '" & Val(iControl.Tag) & "' "
         Next
         
      Case gcFPD
         For Each iControl In .cboFPDMaterial
            sSql = sSql & " OR " & "Lab_Material_ID = '" & Val(iControl.Tag) & "' "
         Next
      Case gcRPD
         For Each iControl In .cboRPDMaterial
            sSql = sSql & " OR " & "Lab_Material_ID = '" & Val(iControl.Tag) & "' "
         Next
         
         For Each iControl In .chkRPDMaterial
            sSql = sSql & " OR " & "Lab_Material_ID = '" & Val(iControl.Tag) & "' "
         Next
               
         For i = 1 To .VSFlexGridRPDFramework.Cols - 1
            sSql = sSql & " OR " & "Lab_Material_ID = '" & Val(.VSFlexGridRPDFramework.ColData(i)) & "' "
         Next
                  
         For i = 1 To .VSFlexGridRPDPontics.Cols - 1
   
   '''''                  'COMMENTED OUT: TO BE DONE IN RELEASE 3
   '''''                  'special case for Shade column
   '''''                  'because of relationship between Guide and Shade columns
   '''''                  'the following logic should be implemented
   '''''                  If vsControl.ColData(i) = gcRPDShade Then
   '''''                     'gcRPDShade = 46
   '''''                     sSql = sSql & " OR " & "Lab_Material_ID = '" & vsControl.ColData(i) & "' "
   '''''                     'gcRPDShade = 47
   '''''                     sSql = sSql & " OR " & "Lab_Material_ID = '" & CStr(CInt(vsControl.ColData(i)) + 1) & "' "
   '''''                     'gcRPDShade = 48
   '''''                     sSql = sSql & " OR " & "Lab_Material_ID = '" & CStr(CInt(vsControl.ColData(i)) + 2) & "' "
   '''''                  Else
   '''''                     sSql = sSql & " OR " & "Lab_Material_ID = '" & vsControl.ColData(i) & "' "
   '''''                  End If
                     
            sSql = sSql & " OR " & "Lab_Material_ID = '" & Val(.VSFlexGridRPDPontics.ColData(i)) & "' "
                     
         Next
         
      Case Else
      '
   End Select
End With

GetSQLLabMaterials = sSql

Exit Function

ErrHandler:
GetSQLLabMaterials = ""

   Call gobjShared.ShowVBError(Err.Number, Err.Description, _
                "modLabsData.GetSQLLabMaterials", "Labs", vbCritical)

End Function

