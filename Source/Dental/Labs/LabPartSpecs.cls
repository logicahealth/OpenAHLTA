VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLabPartSpecs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================================================================
'  Class Name: clsLabPartSpecs
'
'================================================================================================
Option Explicit

'////////////////////  C L A S S   P R O P E R T I E S    S E C T I O N  ////////////////////

Private mcolLabPartSpecs As Collection

'Methods of the LabPartSpecs class.

'This method is called by GetLabTXChildMembers in clsLabTX
Public Function Add(ByRef oRst As Recordset) As clsLabPartSpec
Dim oLabPartSpec As New clsLabPartSpec
Dim sLabPartSpecificationsID As String
Dim sLabTXID As String
Dim sLabTaskCategoryID As String
Dim sLabMaterialElementsID As String
Dim sLabMaterialID As String
Dim sLineItemToothLocation As String
Dim sLabTaskCategoryDefinition As String

Dim sKey As String
        
On Error GoTo ErrHandler

oLabPartSpec.FillFromRecordset oRst
  
'Add the LabPartSpec object reference to the
'collection, using the ID property as the key.
        
sLabTaskCategoryDefinition = oLabPartSpec.TaskCategoryDefinition
sLabTaskCategoryID = oLabPartSpec.LabTaskCategoryID
'sLabPartSpecificationsID = oLabPartSpec.LabPartSpecificationsID

'sKey = Trim$(sLabTaskCategoryID)
sKey = sLabTaskCategoryDefinition
       
If sKey <> "" Then
   mcolLabPartSpecs.Add oLabPartSpec, sKey
Else
   mcolLabPartSpecs.Add oLabPartSpec
End If
           
' Return a reference to the new LabPartSpec
Set Add = oLabPartSpec
    
Exit Function

ErrHandler:
Resume Next

   Call gobjShared.ShowVBError(Err.Number, Err.Description, _
                "clsLabPartSpecs.Add", "Labs", vbCritical)
End Function

Public Function Count() As Long
    Count = mcolLabPartSpecs.Count
End Function

Public Sub Delete(ByVal vIndex As Variant)
    mcolLabPartSpecs.Remove vIndex
End Sub

Public Function Item(ByVal vIndex As Variant) As clsLabPartSpec
    Set Item = mcolLabPartSpecs.Item(vIndex)
End Function

'NewEnum must return the IUnknown interface of a
'collection's enumerator
Public Function NewEnum() As IUnknown
    Set NewEnum = mcolLabPartSpecs.[_NewEnum]
End Function

Private Sub Class_Initialize()
On Error GoTo ErrHandler
'---set/initialize all data members

Set mcolLabPartSpecs = New Collection

Exit Sub
ErrHandler:
   Call gobjShared.ShowVBError(Err.Number, Err.Description, _
                "clsLabLabPartSpecs.Initialize", "Labs", vbCritical)
End Sub

Public Function GetLabPartSpecs(ByRef robjRS As Recordset) As Boolean
''===========================================================================================
''  Description: This method gets the recordset with LabPartSpecs data from database
''
''  Called By:   This function is called by GetLabTXChildMembers method in clsLabTX code
''
''  Argument(s):
''     vValue - Contains index of the LabTX item in the LabTXs collection
''     robjRS - Recordset
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

'<<temp

'--- error handling ---

On Error GoTo ErrHandler

GetLabPartSpecs = True

'lab task category records are getting from Lab_Task_Category_Def static table

sSql = "SELECT * FROM LAB_TASK_CATEGORY_DEF " _
    & "WHERE Definition = '" & gcCrown & "' " _
    & "OR    Definition = '" & gcFPD & "' " _
    & "OR    Definition = '" & gcRPD & "' "

If LabDataProcessing(sSql, robjRS) = gcRowsSelected Then
   GetLabPartSpecs = True
Else
   GetLabPartSpecs = False
End If

Exit Function

ErrHandler:
GetLabPartSpecs = False

   Call gobjShared.ShowVBError(Err.Number, Err.Description, _
                "clsLabPartSpecs.GetLabPartSpecs", "Labs", vbCritical)

End Function

Private Sub Class_Terminate()
   Set mcolLabPartSpecs = Nothing
End Sub
