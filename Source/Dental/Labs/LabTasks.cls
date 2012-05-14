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
Attribute VB_Name = "clsLabTasks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================================================================
'  Class Name: clsLabTasks
'
'================================================================================================
Option Explicit

'////////////////////  C L A S S   P R O P E R T I E S    S E C T I O N  ////////////////////

Private mcolLabTasks As Collection

'Methods of the LabTasks class.
'This method is called by GetLabTXChildMembers in clsLabTX

Public Function Add(ByRef robjValue As Variant) As clsLabTask
Dim oLabTask As New clsLabTask
Dim sLabTaskID As String
Dim sLabTxTaskSpecsID As String
Dim oRs As Recordset
Dim sKey As String
        
On Error GoTo ErrHandler

If TypeOf robjValue Is Recordset Then
   Set oRs = robjValue
   oLabTask.FillFromRecordset oRs
Else
   Set oLabTask = robjValue
End If
 
'Add the LabTask object reference to the
'collection, using the ID property as the key.
        
sLabTxTaskSpecsID = oLabTask.LabTxTaskSpecsID

sKey = Trim$(sLabTxTaskSpecsID)
               
If CStr(oLabTask.LabTxTaskSpecsID) = "" Then
   'there is not associated LabTxTaskSpecsID value, i.e. this information was not retrieved
   'from LAB_TX_SPECIFICATIONS table, therefore this data
   'was manually added to the task grid
   'or this task associated with newly added treatment in BigTooth after LabCase was created
   'or it was New Lab Case issue and
   'should be inserted to the DB
   'Initially all items have gcInsert value
   oLabTask.RecordStatus = gcINSERT
   If Len(oLabTask.DateOrdered) = 0 Then
      oLabTask.DateOrdered = gsDefaultTXDateOrdered
   End If
Else
   'particular item has been retrieved from DB
   'therefore its record status is gcStatusQuo
   oLabTask.RecordStatus = gcStatusQuo
   
End If
                
If sKey <> "" Then
   mcolLabTasks.Add oLabTask, sKey
Else
   mcolLabTasks.Add oLabTask
End If
        
' Return a reference to the new LabCase
Set Add = oLabTask
    
Exit Function

ErrHandler:
   Call gobjShared.ShowVBError(Err.Number, Err.Description, _
                "clsLabTasks.Add", "Labs", vbCritical)
    
Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

Public Function Count() As Long
    Count = mcolLabTasks.Count
End Function

Public Sub Delete(ByVal vIndex As Variant)
    mcolLabTasks.Remove vIndex
End Sub

Public Function Item(ByVal vIndex As Variant) As clsLabTask
    Set Item = mcolLabTasks.Item(vIndex)
End Function

'NewEnum must return the IUnknown interface of a
'collection's enumerator
Public Function NewEnum() As IUnknown
    Set NewEnum = mcolLabTasks.[_NewEnum]
End Function

Private Sub Class_Initialize()
On Error GoTo ErrHandler
'---set/initialize all data members

Set mcolLabTasks = New Collection

Exit Sub
ErrHandler:
   'MsgBox Err.Description, vbExclamation, "Error " & Err.Number & " in " & gsProcName
   Call gobjShared.ShowVBError(Err.Number, Err.Description, _
                "clsLabTasks.Initialize", "Labs", vbCritical)
End Sub

Public Function GetLabTasks(ByRef robjCurLabTX As clsLabTX, _
                            ByRef robjRS As Recordset, _
                            ByVal iChartType As Integer) As Boolean
''===========================================================================================
''  Description: This method gets the recordset with tasks data from database
''
''  Called By:   This function is called by GetLabTXChildMembers method in clsLabTX code
''
''  Argument(s):
''     robjCurLabTX - reference to the current LabTX object
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
Dim sToothLocation
Dim sResult As String

'--- error handling ---
On Error GoTo ErrHandler

GetLabTasks = True

Select Case iChartType
   Case gcGraphChart

      If giLabCaseStatus = gcCreateNewLabCase Then
      
         sSql = "SELECT b.Lab_Task_Category_ID, b.Tx_ADA_Code_ID, b.CLV_Code_ID, a.Definition " _
              & "FROM LAB_TASK_DEFINITION a, LAB_ADA_CLV b, LAB_TASK_CATEGORY_DEF c " _
              & "WHERE b.Tx_ADA_Code_ID = '" & robjCurLabTX.TxADACodeID & "' " _
              & "AND   a.CLV_Code_ID = b.CLV_Code_ID " _
              & "AND   b.Lab_Task_Category_ID = c.Lab_Task_Category_ID " _
              & "ORDER BY b.CLV_Code_ID "
                      
      Else
          'lab case does exist in dental db and
          'lab task records are getting from Lab_Task_Definition table
          'by using Lab_Tx_Tasks dynamic intersect table, and Lab_Tx_Task_Specs table
                      
         sToothLocation = robjCurLabTX.ToothLocation
            
         If CStr(robjCurLabTX.LabTxID) <> "" Then
            'get lab task data from dynamic tables
            sSql = "SELECT b.Lab_Task_Category_ID, b.Tx_ADA_Code_ID, b.CLV_Code_ID, b.Location, f.Definition, T1.Tooth_Name, " _
                   & "b.Lab_TX_Task_Specs_ID, b.Quantity, b.Priority, b.TL_Date_Ordered, b.TL_Date_Completed, b.TL_Date_REQUIRED, b.TL_Date_STARTED, " _
                   & "d.LAB_TX_ID, d.TX_Code_ID, d.Lab_Case_ID " _
                   & "FROM LAB_TASK_DEFINITION f, LAB_TX_TASK_SPECS b, LAB_TX_TASKS c, LAB_TX_SPECIFICATIONS d, " _
                   & "DENTAL_ENCOUNTER_TXS_SS T2, DENTAL_ENCOUNTER_TEETH T1, DENTAL_ENC_SURFACE_SET T3 " _
                   & "WHERE b.Lab_TX_Task_Specs_ID = c.Lab_TX_Task_Spec_ID AND c.LAB_TX_ID = d.LAB_TX_ID " _
                   & "AND   d.Lab_Case_ID = '" & gobjLabCase.LabCaseID & "' AND b.Tx_ADA_Code_ID = '" & robjCurLabTX.TxADACodeID & "' " _
                   & "AND T1.Tooth_Name = '" & sToothLocation & "' AND T3.DENTAL_ENCOUNTER_TEETH_ID = T1.DENTAL_ENCOUNTER_TEETH_ID and T2.SURFACE_SET_ID = T3.SURFACE_SET_ID and d.tx_code_id = T2.tx_code_id AND T3.Surface_Set_ID = '" & robjCurLabTX.SurfaceSetID & "' " _
                   & "AND d.tx_code_id = '" & robjCurLabTX.TxCodeID & "' AND f.clv_code_id = b.clv_code_id ORDER BY b.CLV_Code_ID "
         Else
            'get lab task data from lookup tables
            sSql = "SELECT b.Lab_Task_Category_ID, b.Tx_ADA_Code_ID, b.CLV_Code_ID, a.Definition " _
                 & "FROM LAB_TASK_DEFINITION a, LAB_ADA_CLV b, LAB_TASK_CATEGORY_DEF c " _
                 & "WHERE b.Tx_ADA_Code_ID = '" & robjCurLabTX.TxADACodeID & "' " _
                 & "AND   a.CLV_Code_ID = b.CLV_Code_ID " _
                 & "AND   b.Lab_Task_Category_ID = c.Lab_Task_Category_ID " _
                 & "ORDER BY b.CLV_Code_ID "
                            
         End If
      
      End If
   
   Case gcNonGraphChart
      If giLabCaseStatus = gcCreateNewLabCase Then
      
         sSql = "SELECT b.Lab_Task_Category_ID, b.Tx_ADA_Code_ID, b.CLV_Code_ID, a.Definition " _
              & "FROM LAB_TASK_DEFINITION a, LAB_ADA_CLV b, LAB_TASK_CATEGORY_DEF c " _
              & "WHERE b.Tx_ADA_Code_ID = '" & robjCurLabTX.TxADACodeID & "' " _
              & "AND   a.CLV_Code_ID = b.CLV_Code_ID " _
              & "AND   b.Lab_Task_Category_ID = c.Lab_Task_Category_ID " _
              & "ORDER BY b.CLV_Code_ID "
                      
      Else
          'lab case does exist in dental db and
          'lab task records are getting from Lab_Task_Definition table
          'by using Lab_Tx_Tasks dynamic intersect table, and Lab_Tx_Task_Specs table
                      
         sToothLocation = robjCurLabTX.ToothLocation
            
         If CStr(robjCurLabTX.LabTxID) <> "" Then
            'get lab task data from dynamic tables
            sSql = "SELECT b.Lab_Task_Category_ID, b.Tx_ADA_Code_ID, b.CLV_Code_ID, b.Location, f.Definition, T1.Location_Name, " _
                   & "b.Lab_TX_Task_Specs_ID, b.Quantity, b.Priority, b.TL_Date_Ordered, b.TL_Date_Completed, b.TL_Date_REQUIRED, b.TL_Date_STARTED, " _
                   & "d.LAB_TX_ID, d.TX_Code_ID, d.Lab_Case_ID " _
                   & "FROM LAB_TASK_DEFINITION f, LAB_TX_TASK_SPECS b, LAB_TX_TASKS c, LAB_TX_SPECIFICATIONS d, " _
                   & "NOTE_TXS_LOCATION T2, NOTE_LOCATION T1 " _
                   & "WHERE b.Lab_TX_Task_Specs_ID = c.Lab_TX_Task_Spec_ID AND c.LAB_TX_ID = d.LAB_TX_ID " _
                   & "AND d.Lab_Case_ID = '" & gobjLabCase.LabCaseID & "' AND b.Tx_ADA_Code_ID = '" & robjCurLabTX.TxADACodeID & "' " _
                   & "AND T1.Location_Name = '" & sToothLocation & "' AND d.tx_code_id = T2.tx_code_id " _
                   & "AND d.tx_code_id = '" & robjCurLabTX.TxCodeID & "' AND f.clv_code_id = b.clv_code_id ORDER BY b.CLV_Code_ID "
         Else
            'get lab task data from lookup tables
            sSql = "SELECT b.Lab_Task_Category_ID, b.Tx_ADA_Code_ID, b.CLV_Code_ID, a.Definition " _
                 & "FROM LAB_TASK_DEFINITION a, LAB_ADA_CLV b, LAB_TASK_CATEGORY_DEF c " _
                 & "WHERE b.Tx_ADA_Code_ID = '" & robjCurLabTX.TxADACodeID & "' " _
                 & "AND   a.CLV_Code_ID = b.CLV_Code_ID " _
                 & "AND   b.Lab_Task_Category_ID = c.Lab_Task_Category_ID " _
                 & "ORDER BY b.CLV_Code_ID "
                            
         End If
      
      End If
   
End Select

sResult = LabDataProcessing(sSql, robjRS)

If sResult = gcRowsSelected Or sResult = gcNoRowsSelected Then
   GetLabTasks = True
Else
   GetLabTasks = False
End If

Exit Function

ErrHandler:
GetLabTasks = False

   Call gobjShared.ShowVBError(Err.Number, Err.Description, _
                "clsLabTasks.GetLabTasks", "Labs", vbCritical)

End Function

Private Sub Class_Terminate()
   Set mcolLabTasks = Nothing
End Sub
