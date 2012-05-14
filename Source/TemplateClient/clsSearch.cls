VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Type SearchItem
    szDisplayName As String
    iCollectionIndex As Integer
End Type


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' SearchSpecimen - Seaches the Lab Specimen table
'
' Globals       none
' Call by ref   pListBox        List Box to return SQL quiery results
'               szSearchText    Seach Text
'               szSpecimenData  Optional Search text for the Lab Data column
' Call by val
'
' Returns       the data saved in the list box
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function SearchSpecimen(plistbox As CSComboBox) As srcSearchResult

    Dim oDAL As ICHCSII_DAS
    Dim moConn As CHCSII_CONN.Conn
    Dim szOrderBy As String
    Dim szSelect As String
    Dim szQst As String
    Dim szSQL As String
    Dim szSpecimenDataSQL As String
    Dim gszDBName As String

    On Error GoTo ErrorHandler
    
    szSelect = "SELECT  lab_specimen.specimen_name, lab_specimen.specimen_data, lab_specimen.specimen_ien"
    szQst = ""
    szQst = szQst & " From " & gszDBName & "LAB_SPECIMEN"
    szQst = szQst & " Where lab_specimen.specimen_name is not null"
    szQst = szQst & " AND lab_specimen.specimen_ien <> 0"
    szOrderBy = " ORDER BY lab_specimen.specimen_name;"
    szSQL = szSelect & szQst & szOrderBy
    
    Dim oRecordSet As Recordset
    Dim iFieldIndex As Integer
    Dim sListBoxLine As String
    Dim sDelimiter As String
    
    sDelimiter = "^"  'vbTab
    plistbox.Clear                                      'clear listbox
    Set moConn = New CHCSII_CONN.Conn
    Set oDAL = moConn.CHCSII_DAS(Auto)
    Set oRecordSet = oDAL.OpenRecordset(szSQL)
    
    While Not oRecordSet.EOF                        'add each name and ien
        sListBoxLine = ""
        
        iFieldIndex = 0
        While iFieldIndex < 3          'Build the text list for the list box
            sListBoxLine = sListBoxLine
                    '           sListBoxLine = sListBoxLine & Trim(oRecordSet.GetString(iFieldIndex)) & sDelimiter
            'NEED a blank colum for TOPO
            If iFieldIndex = 1 Then
                    sListBoxLine = sListBoxLine & "[ USE TABLE ]" & sDelimiter
            End If
            
          szSpecimenDataSQL = Trim$(oRecordSet.Fields(iFieldIndex))
          sListBoxLine = sListBoxLine & szSpecimenDataSQL & sDelimiter
                    ''szSpecimenDataSQL = Trim(oRecordSet.Fields(iFieldIndex))
                    ''sListBoxLine = sListBoxLine & szSpecimenDataSQL & sDelimiter
            iFieldIndex = iFieldIndex + 1
        
        Wend
                         ''If iFieldIndex = 3 Then
               sListBoxLine = sListBoxLine & "-1" & sDelimiter & sDelimiter
                         ''  End If

    '??    Trim (oRecordSet.GetString(iFieldIndex)) & sDelimiter
        
        plistbox.AddItem sListBoxLine               'add text list to list box
        oRecordSet.MoveNext                         'get the next record
    Wend
    
          'Screen.ActiveForm.MousePointer = vbDefault          'display default cursor
          'Removed 6-15-01 AAB #12754 and #12790

    If plistbox.ListCount > 0 Then                      'check for empty ListBox
        SearchSpecimen = srcErrNone
    End If
    
    GoTo CleanUpAndExit

ErrorHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "OE_SearchEngine: SearchSpecimen", "TemplateClient", vbExclamation)
    
          'Screen.ActiveForm.MousePointer = vbDefault          'display default cursor
          'Removed 6-15-01 AAB #12754 and #12790
CleanUpAndExit:
    Set moConn = Nothing
    Set oDAL = Nothing

End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GetLabSpecimenData - Returns the Lab Specimen table
'
' Globals       none
' Call by ref   SampleIEN       Lab Speciment IEN to search for
' Call by val
'
' Returns       Text - name of the specified lab specimen
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GetLabSpecimenData(SampleIEN As String) As String

    Dim oDAL As ICHCSII_DAS
    Dim moConn As CHCSII_CONN.Conn
    Dim oRecordSet As Recordset
    Dim szOrderBy As String
    Dim szSelect As String
    Dim szQst As String
    Dim szSQL As String
    Dim gszDBName As String
  
    On Error GoTo ErrorHandler:
    Set moConn = New CHCSII_CONN.Conn
    Set oDAL = moConn.CHCSII_DAS(Auto)
  
    szSelect = "Select lab_specimen.specimen_name, lab_specimen.specimen_data"
    szQst = ""
    szQst = szQst & " From " & gszDBName & "LAB_SPECIMEN"
    szQst = szQst & " Where lab_specimen.specimen_ien = " & SampleIEN
    szQst = szQst & " AND lab_specimen.specimen_name is not null"
    szOrderBy = " ORDER BY lab_specimen.specimen_name;"
    szSQL = szSelect & szQst & szOrderBy

    GetLabSpecimenData = ""
    
    Set oRecordSet = oDAL.OpenRecordset(szSQL)
    
    If Not oRecordSet.EOF Then
        GetLabSpecimenData = Trim$(oRecordSet.Fields(1))
    End If

    GoTo CleanUpAndExit
  
ErrorHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "OE_SearchEngine: GetLabSpecimenData", "TemplateClient", vbExclamation)
CleanUpAndExit:
    Set moConn = Nothing
    Set oDAL = Nothing

End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' SearchTOPO - Seaches the Lab TOPO table
'
' Globals       none
' Call by ref   pListBox        CSComboBox to return SQL quiery results
'               szSearchText    Seach Text
' Call by val
'
' Returns       the data saved in the CSComboBox
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function SearchTOPO(plistbox As CSComboBox, szSearchText As String) As srcSearchResult
    
    Dim szOrderBy As String
    Dim szSelect As String
    Dim szQst As String
    Dim szSQL As String
    Dim gszDBName As String

    On Error GoTo ErrorHandler
    
    ''      'Build query string:
''        szQst = "Select TYPE_NAME, TYPE_IEN"
''        szQst = szQst & " From " & gszDBName & "LAB_TOPO "
''
''        If pbPeriodDetected = False Then
''          szQst = szQst & "where TYPE_NAME like '%" & pszKeyWord & "%'"
''        Else
''          szQst = szQst & "where TYPE_NAME like '" & pszKeyWord & "'"
''        End If
''          ''szQst = szQst & "Order by TYPE_NAME"
''
''        'MsgBxARMd szQst
    
    
    szSelect = "SELECT Distinct lab_topo.type_name, lab_topo.type_ien"
    szQst = ""
    szQst = szQst & " From " & gszDBName & "LAB_TOPO"
    szQst = szQst & " Where " & "lab_topo.type_name like '%" & szSearchText & "%'"       'was disabled 9-24-99; AAB
    szQst = szQst & " AND lab_topo.type_name is not null"
    szQst = szQst & " AND lab_topo.type_ien <> 0"
    'Tracker 24622
    szQst = szQst & " UNION "
    szQst = szQst & " SELECT Distinct lab_topo.type_name, lab_topo.type_ien"
    szQst = szQst & " From " & gszDBName & "LAB_TOPO, " & gszDBName & "topo_synonym "
    szQst = szQst & " Where lab_topo.type_ien = topo_synonym.topo_ien"
    szQst = szQst & " AND topo_synonym.synonym_name like '%" & szSearchText & "%'"
    szQst = szQst & " AND lab_topo.type_name is not null"
    szQst = szQst & " AND lab_topo.type_ien <> 0"

    szOrderBy = " ORDER BY type_name;"
    szSQL = szSelect & szQst & szOrderBy
    
    SearchTOPO = TransferToListBox(szSQL, plistbox, 2)
    
    Exit Function

ErrorHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "OE_SearchEngine: SearchTOPO", "TemplateClient", vbExclamation)
    
End Function

