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
Attribute VB_Name = "clsLabTask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================================================================
'  Class Name: clsLabTask
'
'================================================================================================
Option Explicit

'////////////////////  C L A S S   P R O P E R T I E S    S E C T I O N  ////////////////////

'LAB_TX_TASK_SPECS table AND LAB_ADA_CLV table       'Table Fields
Private msLabTaskCategoryID          As String       'Lab_Task_Category_ID
Private msTxADACodeID                As String       'Tx_ADA_Code_ID
'Private msTaskDefinition             As String      'Definition
Private msCLVCode                    As String       'CLV_Code_ID
Private msEffectiveDate              As String       'EffectiveDate
Private msEffectiveEndDate           As String       'EffectiveEndDate
Private msAlias                      As String       'Alias

'LAB_TASK_DEFINITION table
Private msTaskDefinition             As String       'Definition

'LAB_TASK_DEFINITION (PK) and
'LAB_TX_TASK_SPECS (FK)
Private msLabTaskID                  As String       'Lab_Task_ID: PK and FK

'LAB_TX_TASK_SPECS table
Private msLabTxTaskSpecsID           As String       'Lab_Tx_Task_Specs_ID
Private mlQty                        As Long         'Quantity
Private mlPriority                   As Long         'Priority
Private msDateOrdered                As String       'TL_Date_Ordered
Private msDateCompleted              As String       'TL_Date_Completed
Private msDateRequired               As String       'TL_Date_Required
Private msDateStarted                As String       'TL_Date_Started
Private msComments                   As String       'Comments
Private msStatus                     As String       'Status
Private msCreateDateTime             As String       'CreateDateTime

'GUI fields
Private msQI                         As String       '?
Private msLabLocation                As String       '?
Private msToothLocation              As String       '?

Private msRecordStatus               As String       'INSERT/UPDATE

Private Sub Class_Initialize()

    msLabTaskID = ""
    msLabTaskCategoryID = ""
    msTxADACodeID = ""
    msTaskDefinition = ""
    msCLVCode = ""
    msEffectiveDate = ""
    msEffectiveEndDate = ""

    msLabTxTaskSpecsID = ""
    msLabTaskID = ""
    mlQty = 1
    mlPriority = 0
    msDateOrdered = ""
    msDateCompleted = ""
    msDateRequired = ""
    msDateStarted = ""
    msComments = ""
    msStatus = ""
    msCreateDateTime = ""
    
    msQI = ""
    msLabLocation = ""
    msToothLocation = ""
    
    msRecordStatus = gcStatusQuo
    
End Sub

Public Property Let Qty(ByVal vlValue As Long)
    mlQty = vlValue
End Property

Public Property Get Qty() As Long
    Qty = mlQty
End Property

Public Property Let LabTaskID(ByVal vsValue As String)
    msLabTaskID = vsValue
End Property

Public Property Get LabTaskID() As String
    LabTaskID = msLabTaskID
End Property

Public Property Let LabTaskCategoryID(ByVal vsValue As String)
    msLabTaskCategoryID = vsValue
End Property

Public Property Get LabTaskCategoryID() As String
    LabTaskCategoryID = msLabTaskCategoryID
End Property

Public Property Let TxADACodeID(ByVal vsValue As String)
    msTxADACodeID = vsValue
End Property

Public Property Get TxADACodeID() As String
    TxADACodeID = msTxADACodeID
End Property

Public Property Let LabTxTaskSpecsID(ByVal vsValue As String)
    msLabTxTaskSpecsID = vsValue
End Property

Public Property Get LabTxTaskSpecsID() As String
    LabTxTaskSpecsID = msLabTxTaskSpecsID
End Property

Public Property Let CLVCode(ByVal vsValue As String)
    msCLVCode = vsValue
End Property

Public Property Get CLVCode() As String
    CLVCode = msCLVCode
   'add leading zero if CLV code is one digit
   CLVCode = IIf(Len(CLVCode) = 1, "0" & CLVCode, CLVCode)
End Property

Public Property Let TaskDefinition(ByVal vsValue As String)
    msTaskDefinition = vsValue
End Property

Public Property Get TaskDefinition() As String
    TaskDefinition = msTaskDefinition
End Property

Public Property Let Alias(ByVal vsValue As String)
    msAlias = vsValue
End Property

Public Property Get Alias() As String
    Alias = msAlias
End Property

Public Property Let ToothLocation(ByVal vsValue As String)
    msToothLocation = vsValue
End Property

Public Property Get ToothLocation() As String
    ToothLocation = msToothLocation
End Property

Public Property Let LabLocation(ByVal vsValue As String)
   vsValue = Replace(vsValue, "''", "'")
   vsValue = Replace(vsValue, "'", "''")
   If Len(vsValue) >= gcMAX_STRING_LENGTH Then
      'DB column length validation
      vsValue = Mid$(vsValue, 1, gcMAX_STRING_LENGTH)
   End If
   msLabLocation = vsValue
End Property

Public Property Get LabLocation() As String
    LabLocation = msLabLocation
End Property

Public Property Let DateOrdered(ByVal vsValue As String)
    msDateOrdered = vsValue
End Property

Public Property Get DateOrdered() As String
    DateOrdered = msDateOrdered
End Property

Public Property Let DateRequired(ByVal vsValue As String)
    msDateRequired = vsValue
End Property

Public Property Get DateRequired() As String
    DateRequired = msDateRequired
End Property

Public Property Let DateStarted(ByVal vsValue As String)
    msDateStarted = vsValue
End Property

Public Property Get DateStarted() As String
    DateStarted = msDateStarted
End Property

Public Property Let DateCompleted(ByVal vsValue As String)
    msDateCompleted = vsValue
End Property

Public Property Get DateCompleted() As String
    DateCompleted = msDateCompleted
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

Public Property Let CreateDateTime(ByVal vsValue As String)
    msCreateDateTime = vsValue
End Property

Public Property Get CreateDateTime() As String
    CreateDateTime = msCreateDateTime
End Property

Public Property Let QI(ByVal vsValue As String)
    msQI = vsValue
End Property

Public Property Get QI() As String
    QI = msQI
End Property

Public Property Let Comments(ByVal vsValue As String)
    msComments = vsValue
End Property

Public Property Get Comments() As String
    Comments = msComments
End Property

Public Property Let Status(ByVal vsValue As String)
    msStatus = vsValue
End Property

Public Property Get Status() As String
    Status = msStatus
End Property

Public Property Let Priority(ByVal vlValue As Long)
    mlPriority = vlValue
End Property

Public Property Get Priority() As Long
    Priority = mlPriority
End Property

Public Property Let RecordStatus(ByVal vsValue As String)
    msRecordStatus = vsValue
End Property

Public Property Get RecordStatus() As String
    RecordStatus = msRecordStatus
End Property

Public Sub FillFromRecordset(roRs As Recordset)
''===========================================================================================
''  Description: This method populates fields properties with data from recordset fields
''
''  Called By: this method is called by Add method in clsLabTask class code
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
      
      msLabTaskCategoryID = roRs("Lab_Task_Category_ID") & ""
      msCLVCode = roRs("CLV_Code_ID") & ""
      msTxADACodeID = roRs("Tx_ADA_Code_ID") & ""
      msTaskDefinition = roRs("Definition") & ""
      
      'msRecordStatus = gcUPDATE
      
   Case Else
      'update status
      
      msLabTaskCategoryID = CStr(roRs("Lab_Task_Category_ID")) & ""
      msCLVCode = CStr(roRs("CLV_Code_ID")) & ""
      msTxADACodeID = CStr(roRs("Tx_ADA_Code_ID")) & ""
      msTaskDefinition = roRs("Definition") & ""
      msLabLocation = roRs("Location") & ""
      msLabTxTaskSpecsID = CStr(roRs("Lab_TX_Task_Specs_ID")) & ""
      mlQty = roRs("Quantity")
      mlPriority = roRs("Priority")
      msDateCompleted = CDate(roRs("TL_Date_Completed")) & ""
      
      msDateOrdered = CDate(roRs("TL_Date_Ordered")) & ""
      msDateOrdered = Format$(roRs("TL_Date_Ordered"), "mm/dd/yyyy")
      
      msDateRequired = CDate(roRs("TL_Date_REQUIRED")) & ""
      msDateRequired = Format$(roRs("TL_DATE_REQUIRED"), "mm/dd/yyyy")
      
      msDateStarted = CDate(roRs("TL_Date_STARTED")) & ""
      
      'msRecordStatus = gcUPDATE
      
      'the following should be added to the SQL statement if needed
      'msComments = roRs("Comments") & ""
      'msStatus = roRs("Status") & ""
      'msCreateDateTime = roRs("CreateDateTime") & ""
      
End Select

Exit Sub

ErrHandler:
    Resume Next
   'MsgBox Err.Description, vbExclamation, "Error " & Err.Number & " in " & gsProcName
   Call gobjShared.ShowVBError(Err.Number, Err.Description, _
                "clsLabTask.FillFromRecordset", "Labs", vbCritical)
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
'
