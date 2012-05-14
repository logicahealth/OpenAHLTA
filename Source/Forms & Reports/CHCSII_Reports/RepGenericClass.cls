VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RepGenericClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private moCollRecSet As Collection
Private moCollFormulas As Collection

'hwcc 12/15/2003 - Use goDBSqlOps, goDBDAS, goDBSqlOpsEx global objects for database access
'  Do not delete those objects within this class

Public Function BuildReport(sReportName As String, _
                            oReportItem As ReportItem, _
                            oReport As CRAXDRT.Report, _
                            oCollFormulas As Collection) As Boolean
 
    On Error GoTo ErrHandler
     
    BuildReport = False
    
    '*************************************************************************************************************************************************
    '*** THIS CLASS IS TO BE USED AS A TEMPLATE FOR NEW CLASSES OR TO PROVIDE A BLANK POPUP REPORT ***
    '*** WHEN THE APPLICATION IS RUN FOR ONE OF THE NEW POPULATION HEALTH REPORTS.                                ***
    '*************************************************************************************************************************************************
    
    Dim oCrystal As Crystal
    Dim oRec As ADODB.Recordset
    Dim sSponsorUnitNumber As String
    Dim sUnitNumber As String
    DoEvents
    
    '-- Get Data
    '----------------------------------------------------------------------------
    Set moCollRecSet = Nothing
    Set moCollRecSet = New Collection

    '-- Get Candidates
    '----------------------------------------------------------
    If Not GetData(oReportItem, oRec) Then GoTo CleanUp
    'moCollRecSet.Add oRec, "answers"

'-- Create Report Object
'-----------------------------------------------------------------------------
    If Not CreateReport(sReportName & ".rpt", oReport, oCollFormulas, moCollRecSet) Then GoTo CleanUp

    BuildReport = True

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepGenericClass.BuildReport", "CHCSII_Reports", vbExclamation
CleanUp:
    Set oCrystal = Nothing
    Set oRec = Nothing
    
End Function

Private Function GetData(oReportItem As ReportItem, _
                           oRec As ADODB.Recordset) As Boolean

    Dim sStmt As String
    
    On Error GoTo ErrHandler
    GetData = False

    GetData = True
    GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepGenericClass.GetData", "CHCSII_Reports", vbExclamation
CleanUp:
    Exit Function
End Function

