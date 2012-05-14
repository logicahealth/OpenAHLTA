VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLabNote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================================================================
'  Class Name: clsLabNote
'
'  Description: This class is used by ... to
'               Its main purpose is to enable ...applications to
'
'  Properties:
'    ... - Contains the ... .
'
'  Methods:
'     - Allows an application to ... .
'================================================================================================
Option Explicit

'////////////////////  C L A S S   P R O P E R T I E S    S E C T I O N  ////////////////////

'LAB_NOTE table                        'Table Fields
Private msLabNoteID As String          'Lab_Note_ID
Private msTaskCategoryID As String     'Lab_Task_Category_ID
Private msProviderID As String         'Provider_ID
Private msTitle As String              'Title
Private msBody As String               'Body
Private msCreateDateTime As String     'CreateDateTime

'LAB_NOTES table                       'Field
Private msLabTxID As String            'LAB_TX_ID

Public Property Let LabNoteID(ByVal vsValue As String)
    msLabNoteID = vsValue
End Property

Public Property Get LabNoteID() As String
    LabNoteID = msLabNoteID
End Property

Public Property Let TaskCategoryID(ByVal vsValue As String)
    msTaskCategoryID = vsValue
End Property

Public Property Get TaskCategoryID() As String
    TaskCategoryID = msTaskCategoryID
End Property

Public Property Let ProviderID(ByVal vsValue As String)
    msProviderID = vsValue
End Property

Public Property Get ProviderID() As String
    ProviderID = msProviderID
End Property

Public Property Let Title(ByVal vsValue As String)
    msTitle = vsValue
End Property

Public Property Get Title() As String
    Title = msTitle
End Property

Public Property Let Body(ByVal vsValue As String)
    msBody = vsValue
End Property

Public Property Get Body() As String
    Body = msBody
End Property

Public Property Let CreateDateTime(ByVal vsValue As String)
    msCreateDateTime = vsValue
End Property

Public Property Get CreateDateTime() As String
    CreateDateTime = msCreateDateTime
End Property

Public Property Let LabTxID(ByVal vsValue As String)
    msLabTxID = vsValue
End Property

Public Property Get LabTxID() As String
    LabTxID = msLabTxID
End Property

Public Sub FillFromRecordset(roRs As Recordset)
''===========================================================================================
''  Description: This method populates fields properties with data from recordset fields
''
''  Called By: this method is called by ... method in clsLabNote class code
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

    '--- fill in the properties from the ... table fields ---

Select Case giLabCaseStatus

   Case gcCreateNewLabCase
      msLabNoteID = roRs("Lab_Note_ID") & ""
      msTaskCategoryID = roRs("Lab_Task_Category_ID") & ""
      msProviderID = roRs("Provider") & ""
      msTitle = roRs("Title") & ""
      msBody = roRs("Body") & ""
      msCreateDateTime = CDate(roRs("CreateDateTime")) & ""
   Case Else
      'update status
      msLabNoteID = roRs("Lab_Note_ID") & ""
      msTaskCategoryID = roRs("Lab_Task_Category_ID") & ""
      msProviderID = roRs("Provider") & ""
      msTitle = roRs("Title") & ""
      msBody = roRs("Body") & ""
      msCreateDateTime = CDate(roRs("CreateDateTime")) & ""
      
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
                "clsLabNote.FillFromRecordset", "Labs", vbCritical)
End Sub

