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
  DataSourceBehavior  = 1  'vbDataSource
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDataSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'''************************************************************************
'''Module:        clsDataSource/clsDataSource.cls
'''
'''Description:   Data Source object.
'''                  This class used to supply data where from either DB or other object
'''
'''Modules Required:
'''
'''Note:
'''
'''Author         : Wilson Sun
'''Modified       :
'''               :
'''************************************************************************

Option Explicit
Private mrsNotes As ADODB.Recordset
Private mrsViews As ADODB.Recordset
Attribute mrsViews.VB_VarHelpID = -1
Private moCHCSIIDas As CHCSII_DAS
Private Const msMODULE_NAME As String = "clsDataSource."

Private Sub Class_GetDataMember(DataMember As String, Data As Object)
    ' TODO:  Return the appropriate recordset based on DataMember. For example:
    Select Case DataMember
        Case "Z_NOTES"
           Set Data = mrsNotes
        Case "Z_VIEWS"
           Set Data = mrsViews
        Case ""
           ' Provide a default record source when no Data Member is specified.
           Set Data = Nothing
    End Select

End Sub
Private Sub Class_Initialize()
    Dim sSql As String
    Set moCHCSIIDas = New CHCSII_DAS

    Set gobjShared = New MMMHISShared.CWShared
    
    ' Add strings to the DataMembers collection.
    With DataMembers
        .Add "Z_NOTES"
        .Add "Z_VIEWS"
    End With
   
End Sub
Private Sub Class_Terminate()
    Set mrsNotes = Nothing
    Set mrsViews = Nothing
    Set moCHCSIIDas = Nothing
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Name:  PreProcess
'Description:   Thie method performs two functionalities
'               1. Process and Staging data
'               2. Set Record set to be returned
'               This method is called when the form load and Patient changed
'Paramenters:   pnPatientUnitId:
'               pnEncounterFacilityNcid:
'               pbDataFromDB:
'                   - Ture, Data from Database table.
'                   - False, Data from Collection
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function PreProcess( _
        ByVal pnPatientUnitId As Long, _
        ByVal pnEncounterFacilityNCID As Long, _
        ByVal pbDataFromDB As Boolean) As Boolean
    Const sPROC_NAME As String = "PreProcess"
    Dim sSql As String
    Dim sErrorMsg As String
    
    On Error GoTo ErrorHandler
    
    ' Set Recordset objects.
    Set mrsNotes = Nothing
    Set mrsViews = Nothing
    Set mrsNotes = New ADODB.Recordset
    Set mrsViews = New ADODB.Recordset
    
    
    If pbDataFromDB Then 'Data from DB
       '************Test 1 Dental_Notes Object ******************************************
       '''''''''''''''''''''''''''''''''''''''''
       'From Dental_Notes object
       '''''''''''''''''''''''''''''''''''''''''
        Dim oDentalNotes As New CHCSII_Dental_Notes.clsDTNotesOps
        
        If Not oDentalNotes.Mtd_bGetNotes(pnPatientUnitId, mrsNotes, False, pnEncounterFacilityNCID) Then GoTo ErrorHandler
        
'        '****************Test 2 Temp Table **************************************************
'        'select new data
'        sSql = "select * from Z_NOTES order by DISPLAY_DATE desc"
'        Set mrsNotes = moCHCSIIDas.OpenRecordset(sSql)
'
        sSql = "select * from DENTAL.Z_VIEWS"
        Set mrsViews = moCHCSIIDas.OpenRecordset(sSql)
    Else    'Data from other object
        '''''''''''''''''''''''''''''''''''''''''''''''''
        'Notes
        '''''''''''''''''''''''''''''''''''''''''''''''''
        mrsNotes.CursorLocation = adUseClient
        mrsNotes.Fields.Append "DISPLAY_DATE", adDate
        mrsNotes.Fields.Append "NOTES", adVarChar, 255
        mrsNotes.Fields.Append "DRC", adInteger
        mrsNotes.Open
        
        'TO DO. Adding Loop to get multiple rows' values
        mrsNotes.AddNew
'        mrsNotes.Fields.Item(NotesEnum.eDate) = #1/10/1999#
'        mrsNotes.Fields.Item(NotesEnum.eNotes) = "BB"
'        mrsNotes.Fields.Item(NotesEnum.eDRC) = 2
        mrsNotes!DISPLAY_DATE = Now
        mrsNotes!Notes = "Test Notes 2"
        mrsNotes!DRC = 2
        'mrsNotes.Update
        
        mrsNotes.AddNew
        mrsNotes!DISPLAY_DATE = Now - 1
        mrsNotes!Notes = "Test Notes 1"
        mrsNotes!DRC = 1
        'mrsNotes.Update
        
        '''''''''''''''''''''''''''''''''''''''''''''''''
        'Views
        '''''''''''''''''''''''''''''''''''''''''''''''''
        mrsViews.CursorLocation = adUseClient
        mrsViews.Fields.Append "DIAGNOSIS", adVarChar, 255
        mrsViews.Fields.Append "DX_CODE", adInteger
        mrsViews.Fields.Append "FINDING", adVarChar, 20
        mrsViews.Fields.Append "DX_LOCATION", adVarChar, 50
        mrsViews.Fields.Append "DX_SURFACE", adVarChar, 50
        mrsViews.Fields.Append "DX_DATE", adDate
        mrsViews.Fields.Append "DX_PROVIDER", adVarChar, 40
        mrsViews.Fields.Append "DX_STATUS", adVarChar, 20
        mrsViews.Fields.Append "DR", adInteger
        mrsViews.Fields.Append "PROPOSED_TREATMENT", adVarChar, 200
        mrsViews.Fields.Append "PROPOSED_TX_CODE", adInteger
        mrsViews.Fields.Append "PROPOSED_TX_LOCATION", adVarChar, 50
        mrsViews.Fields.Append "PROPOSED_TX_SURFACE", adVarChar, 50
        mrsViews.Fields.Append "PROPOSED_TREATMENT_PROVIDER", adVarChar, 40
        mrsViews.Fields.Append "COMPLETED_TREATMENT", adVarChar, 200
        mrsViews.Fields.Append "COMPLETED_TX_CODE", adInteger
        mrsViews.Fields.Append "COMPLETED_TX_LOCATION", adVarChar, 50
        mrsViews.Fields.Append "COMPLETED_TX_SURFACE", adVarChar, 50
        mrsViews.Fields.Append "COMPLETED_TX_DATE", adDate
        mrsViews.Fields.Append "COMPLETED_TX_DOCTOR", adVarChar, 40
        mrsViews.Open
        
        'TO DO. Adding Loop to get multiple rows' values
        mrsViews.AddNew
        mrsViews!DIAGNOSIS = "Diagnosis dummy test 1"
        mrsViews!DX_CODE = 1234567890
        mrsViews!FINDING = "finding test"
        mrsViews!DX_LOCATION = "dx location test"
        mrsViews!DX_SURFACE = "dx surface test"
        mrsViews!DX_DATE = Now
        mrsViews!DX_PROVIDER = "dx provider name"
        mrsViews!DX_STATUS = "dx status test"
        mrsViews!DR = 3
        mrsViews!PROPOSED_TREATMENT = "proposed treatment test"
        mrsViews!PROPOSED_TX_CODE = 4
        mrsViews!PROPOSED_TX_LOCATION = "proposed tx location test"
        mrsViews!PROPOSED_TX_SURFACE = "proposed tx surface test"
        mrsViews!PROPOSED_TREATMENT_PROVIDER = "proposed treatment provider name"
        mrsViews!COMPLETED_TREATMENT = "completed treatment test"
        mrsViews!COMPLETED_TX_CODE = 5
        mrsViews!COMPLETED_TX_LOCATION = "completed tx location test"
        mrsViews!COMPLETED_TX_SURFACE = "completed tx surface test"
        mrsViews!COMPLETED_TX_DATE = Now
        mrsViews!COMPLETED_TX_DOCTOR = "completed tx doctor test"
    End If
    PreProcess = True
    Exit Function
ErrorHandler:
    Set mrsViews = Nothing
    Set mrsNotes = Nothing
    Set moCHCSIIDas = Nothing
    PreProcess = False
    Call gobjShared.ShowVBError(Err.Number, Err.Description, msMODULE_NAME & sPROC_NAME, App.Title, vbCritical)
End Function
