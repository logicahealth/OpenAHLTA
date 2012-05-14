VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLabTX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================================================================
'  Class Name: clsLabTX
'
'================================================================================================
Option Explicit

'////////////////////  C L A S S   P R O P E R T I E S    S E C T I O N  ////////////////////

'LAB_TX_SPECIFICATIONS table
Private msLabTxID           As String
Private msTxCodeID          As String
Private msLabCaseID         As String
Private msCreateDateTime    As String
Private msTXDateOrdered     As String
Private msTXDateRequired    As String
Private msTXDateStarted     As String
Private msTXDateCompleted   As String
Private mlQty               As Long

'TX_ADA_CODE_DEFINITION table
Private msTxADACodeID       As String
Private msADACategoryCode  As String
Private msADACode As String
Private msADADefinition     As String
Private msADAAlias As String
Private msADAHelpNarative   As String
Private msChartRelated      As String
Private msLabFlag           As String

Private miChartType         As Integer

'GUI variables
Private msToothLocation     As String
Private msQI                As String

Private msRecordStatus      As String     'INSERT/UDPATE

Private msSurfaceSetID      As String

'link to the object LabTasks (collection class)
'which has private collection of LabTask
Private mobjLabTasks As clsLabTasks

'link to the object LabPartSpecs (collection class)
'which has private collection of LabPartSpecs
Private mobjLabPartSpecs As clsLabPartSpecs

Private Sub Class_Initialize()
On Error GoTo ErrHandler
    
msLabTxID = ""
msTxCodeID = ""
msLabCaseID = ""
msCreateDateTime = ""

msADACategoryCode = ""
msADADefinition = ""
msADAAlias = ""
msADAHelpNarative = ""
msChartRelated = ""
msLabFlag = ""

msToothLocation = ""
mlQty = 1

msRecordStatus = gcStatusQuo

msSurfaceSetID = ""

msADACode = ""
msTxADACodeID = ""
msTXDateOrdered = ""
msTXDateRequired = ""
msTXDateStarted = ""
msTXDateCompleted = ""
msQI = ""
    
miChartType = 0
    
Set mobjLabTasks = New clsLabTasks
Set mobjLabPartSpecs = New clsLabPartSpecs
   
Exit Sub
ErrHandler:
   Call gobjShared.ShowVBError(Err.Number, Err.Description, _
                "clsLabTX.Initialize", "Labs", vbCritical)
End Sub

Private Sub Class_Terminate()
    Set mobjLabTasks = Nothing
    Set mobjLabPartSpecs = Nothing
End Sub

Public Property Set LabTasks(rcolValue As clsLabTasks)
    Set mobjLabTasks = rcolValue
End Property

Public Property Get LabTasks() As clsLabTasks
    Set LabTasks = mobjLabTasks
End Property

Public Property Set LabPartSpecs(rcolValue As clsLabPartSpecs)
    Set mobjLabPartSpecs = rcolValue
End Property

Public Property Get LabPartSpecs() As clsLabPartSpecs
    Set LabPartSpecs = mobjLabPartSpecs
End Property

Public Property Let Qty(ByVal vlValue As Long)
    mlQty = vlValue
End Property

Public Property Get Qty() As Long
    Qty = mlQty
End Property

Public Property Let ADACode(ByVal vsValue As String)
    msADACode = vsValue
End Property

Public Property Get ADACode() As String
    ADACode = msADACode
End Property

Public Property Let ADACategoryCode(ByVal vsValue As String)
    msADACategoryCode = vsValue
End Property

Public Property Get ADACategoryCode() As String
    ADACategoryCode = msADACategoryCode
End Property

Public Property Let TxADACodeID(ByVal vsValue As String)
    msTxADACodeID = vsValue
End Property

Public Property Get TxADACodeID() As String
    TxADACodeID = msTxADACodeID
End Property

Public Property Let ADAAlias(ByVal vsValue As String)
    msADAAlias = vsValue
End Property

Public Property Get ADAAlias() As String
    ADAAlias = msADAAlias
End Property

Public Property Let LabTxID(ByVal vsValue As String)
    msLabTxID = vsValue
End Property

Public Property Get LabTxID() As String
    LabTxID = msLabTxID
End Property

Public Property Let TxCodeID(ByVal vsValue As String)
    msTxCodeID = vsValue
End Property

Public Property Get TxCodeID() As String
    TxCodeID = msTxCodeID
End Property

Public Property Let LabCaseID(ByVal vsValue As String)
    msLabCaseID = vsValue
End Property

Public Property Get LabCaseID() As String
    LabCaseID = msLabCaseID
End Property

Public Property Let ADADefinition(ByVal vsValue As String)
    msADADefinition = vsValue
End Property

Public Property Get ADADefinition() As String
    ADADefinition = msADADefinition
End Property

Public Property Let ToothLocation(ByVal vsValue As String)
    msToothLocation = vsValue
End Property

Public Property Get ToothLocation() As String
    ToothLocation = msToothLocation
End Property

Public Property Let TXDateOrdered(ByVal vsValue As String)
    msTXDateOrdered = vsValue
End Property

Public Property Get TXDateOrdered() As String
    TXDateOrdered = msTXDateOrdered
End Property

Public Property Let TXDateRequired(ByVal vsValue As String)
    msTXDateRequired = vsValue
End Property

Public Property Get TXDateRequired() As String
    TXDateRequired = msTXDateRequired
End Property

Public Property Let TXDateStarted(ByVal vsValue As String)
    msTXDateStarted = vsValue
End Property

Public Property Get TXDateStarted() As String
    TXDateStarted = msTXDateStarted
End Property

Public Property Let TXDateCompleted(ByVal vsValue As String)
    msTXDateCompleted = vsValue
End Property

Public Property Get TXDateCompleted() As String
    TXDateCompleted = msTXDateCompleted
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

Public Property Let ADAHelpNarative(ByVal vsValue As String)
    msADAHelpNarative = vsValue
End Property

Public Property Get ADAHelpNarative() As String
    ADAHelpNarative = msADAHelpNarative
End Property

Public Property Let ChartRelated(ByVal vsValue As String)
    msChartRelated = vsValue
End Property

Public Property Get ChartRelated() As String
    ChartRelated = msChartRelated
End Property

Public Property Let LabFlag(ByVal vsValue As String)
    msLabFlag = vsValue
End Property

Public Property Get LabFlag() As String
    LabFlag = msLabFlag
End Property

Public Property Let SurfaceSetID(ByVal vsValue As String)
    msSurfaceSetID = vsValue
End Property

Public Property Get SurfaceSetID() As String
    SurfaceSetID = msSurfaceSetID
End Property

Public Property Let RecordStatus(ByVal vsValue As String)
    msRecordStatus = vsValue
End Property

Public Property Get RecordStatus() As String
    RecordStatus = msRecordStatus
End Property

Public Property Let ChartType(ByVal viValue As Integer)
    miChartType = viValue
End Property

Public Property Get ChartType() As Integer
    ChartType = miChartType
End Property

Public Sub FillFromRecordset(roRs As Recordset, ByVal iChartType As Integer)
''===========================================================================================
''  Description: This method populates fields properties with data from recordset fields
''
''  Called By: this method is called by Add method in clsLabTXs class code
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

Select Case giLabCaseStatus
   Case gcCreateNewLabCase
      msADACode = roRs("ADA_Code") & ""
      msTxADACodeID = roRs("Tx_ADA_Code_ID") & ""
      msADADefinition = roRs("ADA_Definition") & ""
      msTxCodeID = CStr(roRs("TX_Code_ID")) & ""
      msLabFlag = CStr(roRs("LabFlag")) & ""
            
   Case Else
      'update status
      'LAB_TX_SPECIFICATIONS table
      msLabTxID = CStr(roRs("LAB_TX_ID")) & ""
      msTxCodeID = CStr(roRs("TX_Code_ID")) & ""
      msTXDateOrdered = CDate(roRs("TX_DATE_ORDERED")) & ""
      msTXDateRequired = Format$(roRs("TX_DATE_REQUIRED"), "mm/dd/yyyy")
      msTXDateStarted = CDate(roRs("TX_DATE_STARTED")) & ""
      msTXDateCompleted = CDate(roRs("TX_DATE_COMPLETED")) & ""
      mlQty = roRs("Quantity")
            
      'TX_ADA_CODE_DEFINITION table
      msTxADACodeID = CStr(roRs("Tx_ADA_Code_ID")) & ""
      msADACategoryCode = CStr(roRs("ADA_Category_Code")) & ""
      msADACode = roRs("ADA_Code") & ""
      msADADefinition = roRs("ADA_Definition") & ""
      msLabFlag = CStr(roRs("LabFlag")) & ""
                  
End Select

Select Case iChartType
   Case gcGraphChart
      msToothLocation = CStr(roRs("Tooth_Name")) & ""
      msSurfaceSetID = CStr(roRs("Surface_Set_ID")) & ""
   Case gcNonGraphChart
      msToothLocation = CStr(roRs("Location_Name")) & ""
      msSurfaceSetID = "" 'CStr(roRs("Sub_Location_ID")) & ""
End Select

Exit Sub

ErrHandler:
Resume Next

   Call gobjShared.ShowVBError(Err.Number, Err.Description, _
                "clsLabTX.FillFromRecordset", "Labs", vbCritical)
End Sub

Public Function GetLabTXChildMembers(ByRef robjCurLabTX As clsLabTX, _
                                       ByVal iChartType As Integer) As Boolean
''===========================================================================================
''  Description: This method gets the recordset of the treatment related tasks
''               and creates the collection of the tasks items
''
''  Called By: this method is called by GetLabCaseChildMembers in clsLabCase code
''
''  Argument(s):
''         robjCurLabTX - reference to the current LabTX object
''
''  Returns:
''         TRUE  - if all treatment related tasks data successfully retrieved from database
''         FALSE - if failed
''
''  Notes:
''
''  Programmer:    Alek Samedov
''===========================================================================================

'vValue is the index of TX item in the TXs collection

Dim oLabTasks As clsLabTasks

Dim bResult       As Boolean
Dim bPartSpecErr  As Boolean
Dim oRs           As Recordset

bResult = True
On Error GoTo ErrHandler

'set reference to existing object
Set oLabTasks = robjCurLabTX.LabTasks

If oLabTasks.GetLabTasks(robjCurLabTX, oRs, iChartType) Then
  If oRs.RecordCount <> 0 Then
  oRs.MoveFirst
  Do While Not oRs.EOF
      oLabTasks.Add oRs
      oRs.MoveNext
   Loop
 End If
 
   Set oRs = Nothing
   bResult = True
Else
   bResult = False
End If

GetLabTXChildMembers = bResult

'release reference
Set oLabTasks = Nothing

Exit Function

ErrHandler:
GetLabTXChildMembers = False
If Not oLabTasks Is Nothing Then Set oLabTasks = Nothing

   Call gobjShared.ShowVBError(Err.Number, Err.Description, _
                "clsLabTX.GetLabTXChildMembers", "Labs", vbCritical)
End Function

Public Function GetLabTXPartSpecifications(ByVal vValue As Variant) As Boolean
''===========================================================================================
''  Description: This method gets the recordset of the treatment related tasks
''               and creates the collection of the tasks items
''
''  Called By: this method is called by vsDentalTab_Switch event in frmLabs code
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

'vValue is the index of TX item in the TXs collection
Dim bResult       As Boolean
Dim bPartSpecErr  As Boolean
Dim oRs           As Recordset

Dim oLabPartSpecs As clsLabPartSpecs

On Error GoTo ErrHandler

'set reference to the current LabTx object
Set gobjCurLabTX = gobjLabCase.LabTXs.Item(vValue)

'set reference to current LabPartSpecs object
Set oLabPartSpecs = gobjCurLabTX.LabPartSpecs

'to get lab part specifications (Task Category)
'such as Crown, FPD, and RPD from LAB_TASK_CATEGORY_DEF

'to verify whether lab part specs already populated with data for particular treatment
If oLabPartSpecs.Count = 0 Then
   If oLabPartSpecs.GetLabPartSpecs(oRs) Then
      oRs.MoveFirst
      Dim i As Integer
      i = 1
      Do While Not oRs.EOF And Not bPartSpecErr
         oLabPartSpecs.Add oRs
         
         'gobjCurLabPartSpec is Crown, FPD, or RPD object
         
         'based on data in LAB_TASK_CATEGORY_DEF
         'i = 1 -> LabPartSpec (Task Category) is FPD
         'i = 2 ->                             is RPD
         'i = 3 ->                             is Crown
         'for each task category lab materials collection is being created
         
         Set gobjCurLabPartSpec = gobjCurLabTX.LabPartSpecs.Item(i)
                  
         If gobjCurLabPartSpec.GetLabPartSpecChildMembers(gobjCurLabPartSpec) Then
            oRs.MoveNext
            i = i + 1
         Else
            bPartSpecErr = True
         End If
         
         Set gobjCurLabPartSpec = Nothing
         
      Loop
      
      Set oRs = Nothing
      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Result
      bResult = Not bPartSpecErr
   End If
Else
   bResult = True
End If
  
'release reference
Set oLabPartSpecs = Nothing
Set gobjCurLabTX = Nothing
  
GetLabTXPartSpecifications = bResult

Exit Function

ErrHandler:
GetLabTXPartSpecifications = False
'release references
If Not oLabPartSpecs Is Nothing Then Set oLabPartSpecs = Nothing
If Not gobjCurLabTX Is Nothing Then Set gobjCurLabTX = Nothing
If Not gobjCurLabPartSpec Is Nothing Then Set gobjCurLabPartSpec = Nothing

   Call gobjShared.ShowVBError(Err.Number, Err.Description, _
                "clsLabTX.GetLabTXPartSpecifications", "Labs", vbCritical)

End Function

Public Sub UpdateFromRecordset(roRs As Recordset)
''===========================================================================================
''  Description: This method LabTX object members with data from recordset fields
''
''  Called By: this method is called by GetLabCaseChildMembers method in clsLabCase code
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

   '--- fill in the properties from the LAB_TX_SPECIFICATIONS table fields ---

     msLabTxID = CStr(roRs("LAB_TX_ID")) & ""
     msTXDateOrdered = CDate(roRs("TX_DATE_ORDERED")) & ""
     msTXDateRequired = Format$(roRs("TX_DATE_REQUIRED"), "mm/dd/yyyy")
     msTXDateStarted = CDate(roRs("TX_DATE_STARTED")) & ""
     msTXDateCompleted = CDate(roRs("TX_DATE_COMPLETED")) & ""
     mlQty = roRs("Quantity")

Exit Sub

ErrHandler:
Resume Next

   Call gobjShared.ShowVBError(Err.Number, Err.Description, _
                "clsLabTX.UpdateFromRecordset", "Labs", vbCritical)
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
'                "clsLabTX.FillToRecordset", "Labs", vbCritical)
'End Sub


