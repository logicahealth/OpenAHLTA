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
Attribute VB_Name = "clsLabMaterialElems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================================================================
'  Class Name: clsLabMaterialElems
'
'================================================================================================
Option Explicit

'////////////////////  C L A S S   P R O P E R T I E S    S E C T I O N  ////////////////////

Private mcolLabMaterialElems As Collection

'Methods of the LabMaterialElems class.

Public Function Add(ByRef robjValue As Variant) As clsLabMaterialElem
Dim oLabMaterialElem As New clsLabMaterialElem
Dim sLabMaterialElemsID As String
Dim sLabMaterialElemsItemNumber As String
Dim sLabMaterialCategory As String
Dim oRs As Recordset
        
Dim sKey As String
        
On Error GoTo ErrHandler

If TypeOf robjValue Is Recordset Then
   Set oRs = robjValue
   oLabMaterialElem.FillFromRecordset oRs
Else
   Set oLabMaterialElem = robjValue
End If

'Add LabMaterialElem object reference to the
'collection, using ID property as the key.
        
sLabMaterialElemsID = oLabMaterialElem.LabMaterialElementsID
If oLabMaterialElem.LineItemNumber > 0 Then
   sLabMaterialElemsItemNumber = CStr(oLabMaterialElem.LineItemNumber)
End If

sLabMaterialCategory = oLabMaterialElem.LabMaterialCategory

If sLabMaterialElemsItemNumber <> "" Then
   sKey = sLabMaterialElemsID & sLabMaterialElemsItemNumber & sLabMaterialCategory
Else
   sKey = sLabMaterialElemsID
End If

If sLabMaterialElemsID <> "" Then
   mcolLabMaterialElems.Add oLabMaterialElem, sKey
Else
   mcolLabMaterialElems.Add oLabMaterialElem
End If
            
' Return a reference to the new LabMaterialElem
Set Add = oLabMaterialElem
    
Exit Function

ErrHandler:
Resume Next
   Call gobjShared.ShowVBError(Err.Number, Err.Description, _
                "clslabmaterialelems.Add", "Labs", vbCritical)
    
End Function

Public Function Count() As Long
    Count = mcolLabMaterialElems.Count
End Function

Public Sub Delete(ByVal vIndex As Variant)
    mcolLabMaterialElems.Remove vIndex
End Sub

Public Function Item(ByVal vIndex As Variant) As clsLabMaterialElem
    Set Item = mcolLabMaterialElems.Item(vIndex)
End Function

'NewEnum must return the IUnknown interface of a
'collection's enumerator
Public Function NewEnum() As IUnknown
    Set NewEnum = mcolLabMaterialElems.[_NewEnum]
End Function

Private Sub Class_Initialize()
On Error GoTo ErrHandler
'---set/initialize all data members

Set mcolLabMaterialElems = New Collection

Exit Sub
ErrHandler:
   Call gobjShared.ShowVBError(Err.Number, Err.Description, _
                "clslabmaterialelems.Initialize", "Labs", vbCritical)
End Sub

Public Function GetLabMaterialElems(ByVal vLabMaterialID As String, _
                                    ByRef robjRS As Recordset _
                                    ) As Boolean
''===========================================================================================
''  Description: This method gets the recordset with labmaterialelems data from database
''
''  Called By:   This function is called by GetLabMaterialChildMembers method in clsLabMaterial code
''
''  Argument(s):
''     vLabMaterialID - LabMaterialId value
''     vValue         - Contains index of the LabMaterial item in the LabMaterial collection
''     robjRS         - Recordset
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

GetLabMaterialElems = True

'''<DEV Timer
'''Dim dTimeStart As Double
'''dTimeStart = Timer

sSql = "SELECT a.Lab_Material_Elements_ID, a.Description, " _
        & "b.Lab_Part_Specifications_ID, b.Lab_TX_ID, b.Lab_Task_Category_ID, b.Lab_Material_ID, b.Line_Item_Tooth_Location, b.Line_Item_Number, b.Lab_Material_Category " _
        & "FROM LAB_MATERIAL_ELEMENTS a, LAB_PART_SPECIFICATIONS b " _
        & "WHERE a.Lab_Material_Elements_ID = b.Lab_Material_Elements_ID " _
        & "AND   b.LAB_TX_ID =  '" & gobjCurLabTX.LabTxID & "' " _
        & "AND   b.Lab_Task_Category_ID = '" & gobjCurLabPartSpec.LabTaskCategoryID & "' " _
        & "AND   b.Lab_Material_ID = '" & vLabMaterialID & "' "

Dim sResult As String
sResult = LabDataProcessing(sSql, robjRS)

If sResult = gcRowsSelected Then
   'no row selected - Ok
   GetLabMaterialElems = True
Else
   GetLabMaterialElems = False
End If

'''<DEV timer
'''Debug.Print Format$(Timer - dTimeStart, "0.00000")

Exit Function

ErrHandler:
GetLabMaterialElems = False

   Call gobjShared.ShowVBError(Err.Number, Err.Description, _
                "clsLabMaterialElems.GetLabMaterialElems", "Labs", vbCritical)

End Function


Private Sub Class_Terminate()
   Set mcolLabMaterialElems = Nothing
End Sub

'Public Function GetLookupMaterialElems(ByVal vValue As Variant, _
'                                       ByRef robjRS As Recordset) As Boolean
'''===========================================================================================
'''  Description:  This method gets the recordset with labmaterialelems data from database
'''
'''  Called By:    This function is called by GetLookupMaterialChildMembers method
'''                in clsLabMaterial code
'''
'''  Argument(s):
'''     vValue - Contains index of the LabMaterialElem item in the LabMaterialElems collection
'''     robjRS - Recordset
'''
'''  Returns:
'''         TRUE  - if success
'''         FALSE - if failed
'''
'''  Notes:
'''
'''  Programmer:    Alek Samedov
'''===========================================================================================
'Dim sSql As String
'Dim sResult As String
'
''--- error handling ---
'
'On Error GoTo ErrHandler
'
'GetLookupMaterialElems = True
'
''Lab Material records are getting from Lab_Materials_Elements lookup table
''joining with Lab_Materials table
'
'sSql = "SELECT a.Lab_Material_Elements_ID, a.Lab_Material_ID, a.Abbrev, a.Description, a.Active " _
'    & "FROM LAB_MATERIAL_ELEMENTS a, LAB_MATERIALS b " _
'    & "WHERE a.Lab_Material_ID = b.Lab_Material_ID " _
'    & "AND b.Lab_Material_ID = '" & gobjLookupLabMaterials.Item(vValue).LabMaterialID & "' " _
'    & "ORDER BY a.Description "
'
'sResult = LabDataProcessing(sSql, robjRS)
'
'If sResult = gcRowsSelected Then       'Or sResult = gcNoRowsSelected
'   'no row selected - Ok
'   GetLookupMaterialElems = True
'Else
'   GetLookupMaterialElems = False
'End If
'
'Exit Function
'
'ErrHandler:
'GetLookupMaterialElems = False
'
'   Call gobjShared.ShowVBError(Err.Number, Err.Description, _
'                "clsLabMaterialElems.GetLookupMaterialElems", "Labs", vbCritical)
'
'End Function
'

'Public Function GetLabMaterialElemsBackup(ByVal vLabMaterialID As String, _
'                                    ByRef robjRS As Recordset _
'                                    ) As Boolean
'''===========================================================================================
'''  Description: This method gets the recordset with labmaterialelems data from database
'''
'''  Called By:   This function is called by GetLabMaterialChildMembers method in clsLabMaterial code
'''
'''  Argument(s):
'''     vLabMaterialID - LabMaterialId value
'''     vValue         - Contains index of the LabMaterial item in the LabMaterial collection
'''     robjRS         - Recordset
'''
'''  Returns:
'''         TRUE  - if success
'''         FALSE - if failed
'''
'''  Notes:
'''
'''  Programmer:    Alek Samedov
'''===========================================================================================
'Dim sSql As String
'
''--- error handling ---
'
'On Error GoTo ErrHandler
'
'GetLabMaterialElems = True
'
''''<DEV Timer
''''Dim dTimeStart As Double
''''dTimeStart = Timer
'
'sSql = "SELECT a.Lab_Material_Elements_ID, a.Description, " _
'        & "b.Lab_Part_Specifications_ID, b.Lab_TX_ID, b.Lab_Task_Category_ID, b.Lab_Material_ID, b.Line_Item_Tooth_Location, b.Line_Item_Number, b.Lab_Material_Category " _
'        & "FROM LAB_MATERIAL_ELEMENTS a, LAB_PART_SPECIFICATIONS b " _
'        & "WHERE a.Lab_Material_Elements_ID = b.Lab_Material_Elements_ID " _
'        & "AND   b.LAB_TX_ID =  '" & gobjCurLabTX.LabTxID & "' " _
'        & "AND   b.Lab_Task_Category_ID = '" & gobjCurLabPartSpec.LabTaskCategoryID & "' " _
'        & "AND   b.Lab_Material_ID = '" & vLabMaterialID & "' "
'
'Dim sResult As String
'sResult = LabDataProcessing(sSql, robjRS)
'
'If sResult = gcRowsSelected Then       'Or sResult = gcNoRowsSelected
'   'no row selected - Ok
'   GetLabMaterialElems = True
'Else
'   GetLabMaterialElems = False
'End If
'
''''<DEV timer
''''Debug.Print Format$(Timer - dTimeStart, "0.00000")
'
'Exit Function
'
'ErrHandler:
'GetLabMaterialElems = False
'
'   Call gobjShared.ShowVBError(Err.Number, Err.Description, _
'                "clsLabMaterialElems.GetLabMaterialElems", "Labs", vbCritical)
'
'End Function
'

'Public Function GetLabMaterialElems(ByRef robjRS As Recordset _
'                                    ) As Boolean
'''===========================================================================================
'''  Description: This method gets the recordset with labmaterialelems data from database
'''
'''  Called By:   This function is called by GetLabMaterialChildMembers method in clsLabMaterial code
'''
'''  Argument(s):
'''     robjRS         - Recordset
'''
'''  Returns:
'''         TRUE  - if success
'''         FALSE - if failed
'''
'''  Notes:
'''
'''  Programmer:    Alek Samedov
'''===========================================================================================
'Dim sSql As String
'
''--- error handling ---
'
'On Error GoTo ErrHandler
'
'GetLabMaterialElems = True
'
''''<DEV Timer
''''Dim dTimeStart As Double
''''dTimeStart = Timer
'
'sSql = "SELECT a.Lab_Material_Elements_ID, a.Description, " _
'        & "b.Lab_Part_Specifications_ID, b.Lab_TX_ID, b.Lab_Task_Category_ID, b.Lab_Material_ID, b.Line_Item_Tooth_Location, b.Line_Item_Number, b.Lab_Material_Category " _
'        & "FROM LAB_MATERIAL_ELEMENTS a, LAB_PART_SPECIFICATIONS b " _
'        & "WHERE a.Lab_Material_Elements_ID = b.Lab_Material_Elements_ID " _
'        & "AND   b.LAB_TX_ID =  '" & gobjCurLabTX.LabTxID & "' " _
'        & "AND   b.Lab_Task_Category_ID = '" & gobjCurLabPartSpec.LabTaskCategoryID & "' " _
'        & "ORDER BY b.Lab_Material_ID "
'
'        '& "AND   b.Lab_Material_ID = '" & vLabMaterialID & "' "
'
'Dim sResult As String
'sResult = LabDataProcessing(sSql, robjRS)
'
'If sResult = gcRowsSelected Then       'Or sResult = gcNoRowsSelected
'   'no row selected - Ok
'   GetLabMaterialElems = True
'Else
'   GetLabMaterialElems = False
'End If
'
''''<DEV timer
''''Debug.Print Format$(Timer - dTimeStart, "0.00000")
'
'Exit Function
'
'ErrHandler:
'GetLabMaterialElems = False
'
'   Call gobjShared.ShowVBError(Err.Number, Err.Description, _
'                "clsLabMaterialElems.GetLabMaterialElems", "Labs", vbCritical)
'
'End Function
'
'
