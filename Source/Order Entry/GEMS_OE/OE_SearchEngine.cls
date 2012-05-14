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
Attribute VB_Name = "OE_SearchEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' Begin:SCR #318;   Developer: Jane Shen 09/26/2002 10:22 AM
' End: SCR #318;
Option Explicit

Public Enum srcSynonyms
    srcSimple
    srcSynonym
    srcboth
End Enum

Public Enum srcIdentifier
    srcGYN
    srcAutopsy
End Enum

Public Enum srcSearchResult
    srcErrNone
    srcErrUnknown
End Enum

Public Type LabRadItem
    szDisplayName As String
    szIEN As String
    szName As String
    szSynonym As String
    szPanelFlag As String
    szSubscript As String
End Type

Private Type SearchItem
    szDisplayName As String
    iCollectionIndex As Integer
End Type

Private oDAL As ICHCSII_DAS


Function szCreateConnectString(szPassword As String) As String

    szCreateConnectString = GetAppProfile("ADO", "ConnectStringOE", "")
        
End Function
Friend Function Search(plistbox As CSComboBox, _
                        OrderType As hlORDERTYPECODE, _
                        szSearchText As String, _
                        Optional srcSearchWhat As srcSynonyms = srcboth, _
                        Optional srcSection As Variant = "") As srcSearchResult
    
Dim szOrderBy As String
Dim szSelect As String
Dim szQst As String
Dim szSQL As String
Dim szSectionSQL As String
Dim iNumberofFields As Integer
Dim szFacility As String

    On Error GoTo ErrorHandler
    iNumberofFields = 2
    
    Select Case OrderType
    
        Case hlRX_ORDERS_CODE:  'search both name and synonyms
                    
            ' Begin:SCR #43030;   Developer: Jane Shen 10/24/2003 05:19 PM
            ' Note: Function Changed:    changes made in OE_SearchEngin.cls:search().    Impact of Change:    Search on both name and synonym
            ' End: SCR #43030;
            szSelect = "SELECT Distinct rx_list.drug_name, rx_list.drug_ien"
            szQst = ""
            szQst = szQst & " From  RX_LIST"
            szQst = szQst & " Where rx_list.drug_name like " & pobjConn.SQLQuote("%" & szSearchText & "%", 100)
            szQst = szQst & " and (rx_list.DRUG_CHECK_DISABLED is NULL or rx_list.DRUG_CHECK_DISABLED = 0)"
            szQst = szQst & " and (rx_list.NDC_INDICATOR is NULL or rx_list.NDC_INDICATOR = 0)"
            szQst = szQst & " and rx_list.drug_name is not null"
            szQst = szQst & " and rx_list.drug_ien <> 0"
            If gobjShared.IsAppMode(modeTheater) Then
                szQst = szQst & " UNION"
                szQst = szQst & " SELECT DISTINCT RX_LIST.drug_name, RX_LIST.drug_ien "
                szQst = szQst & " FROM  RX_LIST "
                szQst = szQst & " WHERE  rx_list.synonyms like " & pobjConn.SQLQuote("%" & szSearchText & "%", 100)
                szQst = szQst & " and (rx_list.DRUG_CHECK_DISABLED is NULL or rx_list.DRUG_CHECK_DISABLED = 0)"
                szQst = szQst & " and (rx_list.NDC_INDICATOR is NULL or rx_list.NDC_INDICATOR = 0)"
                szQst = szQst & " and rx_list.drug_name is not null"
                szQst = szQst & " and rx_list.drug_ien <> 0"
            End If
        '===============================================================================
        Case hlRAD_ORDERS_CODE:
            
            If LenB(srcSection) > 0 Then
                szSectionSQL = " AND rad_test.rad_test_type like '" & srcSection & "%'"
            Else
                szSectionSQL = ""
            End If
            
            szSelect = "SELECT Distinct rad_test.rad_test_name, rad_test.rad_test_ien, rad_test.rad_test_type"
            szQst = ""
            szQst = szQst & " From rad_test"
            szQst = szQst & " Where rad_test.rad_test_name like " & pobjConn.SQLQuote("%" & szSearchText & "%", 100)
            szQst = szQst & " and (INACTIVE_FLAG is NULL or INACTIVE_FLAG = 0)"
            szQst = szQst & " and DELETE_FLAG <> 1 "  'SCR 69610
            szQst = szQst & " AND rad_test.rad_test_name is not null"
            szQst = szQst & " AND rad_test.Rad_test_ien <> 0"
            szQst = szQst & szSectionSQL
            If gobjShared.IsAppMode(modeTheater) Then
                szQst = szQst & " UNION"
                szQst = szQst & " SELECT DISTINCT rad_test_name, rad_test_ien, rad_test_type"
                szQst = szQst & " FROM  rad_test "
                szQst = szQst & " WHERE  rad_test.synonyms like " & pobjConn.SQLQuote("%" & szSearchText & "%", 100)
                szQst = szQst & " and (INACTIVE_FLAG is NULL or INACTIVE_FLAG = 0)"
                szQst = szQst & " and DELETE_FLAG <> 1 "  'SCR 69610
                szQst = szQst & " AND rad_test.rad_test_name is not null"
                szQst = szQst & " AND rad_test.Rad_test_ien <> 0"
                szQst = szQst & szSectionSQL
            End If
        '===============================================================================
        Case hlLAB_ORDERS_CODE:
            iNumberofFields = 4
            
            If LenB(srcSection) > 0 Then
                If gobjShared.IsAppMode(modeTheater) Then '<SCR 37864
                    szSectionSQL = " AND LAB_TEST_OE.SUBSCRIPT = '" & srcSection & "'"
                Else
                    szSectionSQL = " AND LAB_TEST_OE.subscript = '" & srcSection & "'"
                End If
            Else
                szSectionSQL = ""
            End If
                        
            szSelect = "SELECT Distinct LAB_TEST_OE.lab_test_name, LAB_TEST_OE.lab_test_ien,"
            szSelect = szSelect & " LAB_TEST_OE.panel_flag, LAB_TEST_OE.subscript "
            szQst = ""
            szQst = szQst & " From  LAB_TEST_OE"
            szQst = szQst & " Where LAB_TEST_OE.lab_test_name like " & pobjConn.SQLQuote("%" & szSearchText & "%", 100)
            szQst = szQst & " and (INACTIVE_FLAG is NULL or INACTIVE_FLAG = 0)"
            szQst = szQst & " and DELETE_FLAG <> 1 " 'SCR 69610
            szQst = szQst & " AND LAB_TEST_OE.lab_test_name is not null"
            szQst = szQst & " AND LAB_TEST_OE.lab_test_ien <> 0"
            szQst = szQst & szSectionSQL
            
            If gobjShared.IsAppMode(modeTheater) Then
                szQst = szQst & " UNION"
                szQst = szQst & " SELECT Distinct LAB_TEST_OE.lab_test_name, LAB_TEST_OE.lab_test_ien,"
                szQst = szQst & " LAB_TEST_OE.panel_flag, LAB_TEST_OE.subscript "
                szQst = szQst & " FROM  LAB_TEST_OE "
                szQst = szQst & " WHERE  LAB_TEST_OE.synonyms like " & pobjConn.SQLQuote("%" & szSearchText & "%", 100)
                szQst = szQst & " and (INACTIVE_FLAG is NULL or INACTIVE_FLAG = 0)"
                szQst = szQst & " and DELETE_FLAG <> 1 " 'SCR 69610
                szQst = szQst & " AND LAB_TEST_OE.lab_test_name is not null"
                szQst = szQst & " AND LAB_TEST_OE.lab_test_ien <> 0"
                szQst = szQst & szSectionSQL
            End If
            
    End Select
    If gobjShared.IsAppMode(modeTheater) Then
        szFacility = " and FACILITY_NCID = " & cwobjLogonEnviron.FacilityNCID
    End If
    szOrderBy = " Order by  1"
    szSQL = szSelect & szQst & szFacility & szOrderBy
   
    If gobjShared.IsAppMode(modeITT) Then   '<SCR 37864
        If OrderType = hlRX_ORDERS_CODE Then
            Search = TransferToListBox(szSQL, plistbox, iNumberofFields)
        Else
            Search = TransferToRadLabListBox(szSQL, plistbox, OrderType)
        End If

    Else
        Search = TransferToListBox(szSQL, plistbox, iNumberofFields)
    End If
    
    Exit Function

ErrorHandler:
    DisplayLogErr "OE_SearchEngine: Search"
    
    Exit Function

    Err.Clear
    Resume
    
End Function
Friend Function LoadAllRadTestTypes(plistbox As CSComboBox) As srcSearchResult

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' LoadAllRadTestTypes - Seaches the Rad Test Type Table
'
' Globals       none
' Call by ref   pListBox        CSComboBox to return SQL quiery results
'
' Returns       the data saved in the CSComoBox
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim szSelect As String

    On Error GoTo ErrorHandler

    szSelect = "SELECT Distinct rad_test_type"
   
    szSelect = szSelect & " From RAD_TEST"
    szSelect = szSelect & " Where rad_test_type is not null"
    
    LoadAllRadTestTypes = TransferToListBox(szSelect, plistbox, 1)
    
    Exit Function

ErrorHandler:
    DisplayLogErr "OE_SearchEngine: LoadAllRadTestTypes"
    
End Function
Friend Function LoadAllRadLocations(plistbox As CSComboBox) As srcSearchResult

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' LoadAllRadLocations - Seaches the Clinic Locations for Imaging Type
'
' Globals       none
' Call by ref   pListBox        CSComboBox to return SQL quiery results
'
' Returns       the data saved in the CSComoBox
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim szOrderBy As String
    Dim szSelect As String
    Dim szQst As String
    Dim szSQL As String
    Dim gszDBName As String

    On Error GoTo ErrorHandler

    szSelect = "SELECT hospital_loc_name, hospital_loc_ien"
    szQst = ""
     szQst = szQst & " From " & gszDBName & " MEPRS"
    szQst = szQst & " Where hospital_loc_type = 'IMAGING'"
    szOrderBy = " ORDER BY 1"
    szSQL = szSelect & szQst & szOrderBy
    
    LoadAllRadLocations = TransferToListBox(szSQL, plistbox, 2)
    
    Exit Function

ErrorHandler:
    DisplayLogErr "OE_SearchEngine: LoadAllRadLocations"
    
End Function


Friend Function LoadAllLabTestTypes(plistbox As CSComboBox) As srcSearchResult
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' LoadAllLabTestTypes - Seaches the Lab Test Type table
'
' Globals       none
' Call by ref   pListBox        CSComboBoxBox to return SQL quiery results
' Call by val
'
' Returns       the data saved in the CSComboBox
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    On Error GoTo ErrorHandler
    
    If plistbox.ListCount < 12 Then
        plistbox.AddItem "AUTOPSY" & plistbox.ColDelim & "AU"
        plistbox.AddItem "BACTERIOLOGY" & plistbox.ColDelim & "BT"
        plistbox.AddItem "BLOOD BANK" & plistbox.ColDelim & "BB"
        plistbox.AddItem "BONE MARROW" & plistbox.ColDelim & "BM"
        plistbox.AddItem "CHEMISTRY" & plistbox.ColDelim & "CH"
        plistbox.AddItem "CYTOLOGY GYN" & plistbox.ColDelim & "CG"
        plistbox.AddItem "CYTOLOGY NON-GYN" & plistbox.ColDelim & "CN"
        plistbox.AddItem "MYCOBACTERIUM" & plistbox.ColDelim & "TB"
        plistbox.AddItem "MYCOLOGY" & plistbox.ColDelim & "MY"
        plistbox.AddItem "PARASITOLOGY" & plistbox.ColDelim & "OP"
        plistbox.AddItem "SURGICAL PATHOLOGY" & plistbox.ColDelim & "SP"
        plistbox.AddItem "VIROLOGY" & plistbox.ColDelim & "VR"
    End If
    
    Exit Function

ErrorHandler:
    DisplayLogErr "OE_SearchEngine: LoadAllLabTestTypes"
    
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
    
    
    szSelect = "SELECT Distinct lab_topo.type_name, lab_topo.type_ien"
    szQst = ""
    szQst = szQst & " From " & gszDBName & "LAB_TOPO"
    szQst = szQst & " Where " & "lab_topo.type_name like '%" & szSearchText & "%'"      'was disabled 9-24-99; AAB
    szQst = szQst & " AND lab_topo.type_name is not null"
    szQst = szQst & " AND lab_topo.type_ien <> 0"
    szOrderBy = " ORDER BY lab_topo.type_name;"
    szSQL = szSelect & szQst & szOrderBy
    
    SearchTOPO = TransferToListBox(szSQL, plistbox, 2)
    
    Exit Function

ErrorHandler:
    DisplayLogErr "OE_SearchEngine: SearchTOPO"
    
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
Friend Function SearchTOPOWithIdentifier(plistbox As CSComboBox, srcSearchWhat As srcIdentifier) As srcSearchResult
    
    Dim szOrderBy As String
    Dim szSelect As String
    Dim szQst As String
    Dim szSQL As String
    Dim gszDBName As String
    Dim szSearchIdentifer As String
    
    On Error GoTo ErrorHandler
    If srcSearchWhat = srcAutopsy Then
        szSearchIdentifer = "AY"
    Else
        szSearchIdentifer = "GS"
    End If
    
    szSelect = "SELECT Distinct lab_topo.type_name, lab_topo.type_ien"
    szQst = ""
    szQst = szQst & " From " & gszDBName & "LAB_TOPO"
    szQst = szQst & " Where " & "lab_topo.identifier = '" & szSearchIdentifer & "'"      'was disabled 9-24-99; AAB
    szQst = szQst & " AND lab_topo.type_name is not null"
    szQst = szQst & " AND lab_topo.type_ien <> 0"
    szOrderBy = " ORDER BY lab_topo.type_name;"
    szSQL = szSelect & szQst & szOrderBy
    
    SearchTOPOWithIdentifier = TransferToListBox(szSQL, plistbox, 2)
    
    Exit Function

ErrorHandler:
    DisplayLogErr "OE_SearchEngine: SearchTOPOWithIdentifier"
    
End Function
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

    sDelimiter = "^"
    plistbox.Clear                                      'clear listbox
    Set oRecordSet = oDAL.OpenRecordset(szSQL)
    
    While Not oRecordSet.EOF                        'add each name and ien
        sListBoxLine = ""
        
        iFieldIndex = 0
        While iFieldIndex < 3          'Build the text list for the list box
            sListBoxLine = sListBoxLine
            If iFieldIndex = 1 Then
                    sListBoxLine = sListBoxLine & "[ USE TABLE ]" & sDelimiter
            End If
            
            szSpecimenDataSQL = Trim$(oRecordSet.Fields(iFieldIndex) & " ")
            sListBoxLine = sListBoxLine & szSpecimenDataSQL & sDelimiter
            iFieldIndex = iFieldIndex + 1
        
        Wend
        sListBoxLine = sListBoxLine & "-1" & sDelimiter & sDelimiter
        plistbox.AddItem sListBoxLine               'add text list to list box
        oRecordSet.MoveNext                         'get the next record
    Wend
    
    If plistbox.ListCount > 0 Then                      'check for empty ListBox
        SearchSpecimen = srcErrNone
    End If
    
    Exit Function

ErrorHandler:
    DisplayLogErr "OE_SearchEngine: SearchSpecimen"
    
    Exit Function

    Err.Clear
    Resume
    
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

    Dim oRecordSet As Recordset
    Dim szOrderBy As String
    Dim szSelect As String
    Dim szQst As String
    Dim szSQL As String
    Dim gszDBName As String
  
    On Error GoTo ErrorHandler:
  
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

    Exit Function
  
ErrorHandler:
    DisplayLogErr "OE_SearchEngine: GetLabSpecimenData"

End Function
Public Function GetSpecimenDataCollection() As Collection

    Dim oRecordSet As Recordset
    Dim szOrderBy As String
    Dim szSelect As String
    Dim szQst As String
    Dim szSQL As String
    Dim objSpecimen As TypeSpecimen
    Dim cols As Collection
    
On Error GoTo ErrorHandler:
  
    szSelect = "Select specimen_name, specimen_data,specimen_ien"
    szQst = ""
    szQst = szQst & " From LAB_SPECIMEN"
    szQst = szQst & " Where specimen_data is not null "
    szQst = szQst & " AND specimen_name is not null"
    szOrderBy = " ORDER BY lab_specimen.specimen_name;"
    szSQL = szSelect & szQst & szOrderBy
    
    Set oRecordSet = oDAL.OpenRecordset(szSQL)
    Set cols = New Collection
    
    While Not oRecordSet.EOF
        Set objSpecimen = New TypeSpecimen
        objSpecimen.Name = Trim$(oRecordSet.Fields(0))
        objSpecimen.CapColorOrContainer = Trim$(oRecordSet.Fields(1))
        objSpecimen.IEN = Trim$(oRecordSet.Fields(2))
        cols.Add objSpecimen
        oRecordSet.MoveNext
    Wend
    
    Set GetSpecimenDataCollection = cols
    Exit Function
  
ErrorHandler:
    DisplayLogErr "OE_SearchEngine: GetSpecimenDataCollection"

End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' SearchLabPanel - Seaches the Lab Panel table
'
' Globals       none
' Call by ref   ListBox         List Box to return SQL quiery results
'               SampleIEN       Lab Speciment IEN to search for
' Call by val
'
' Returns       the data saved in the list box
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function SearchLabPanel(plistbox As ListBox, _
                            SampleIEN As String) As srcSearchResult
                            
    Dim szQst As String
    Dim gszDBName As String

    On Error GoTo ErrorHandler:
  

    szQst = "Select LAB_TEST_NAME"
    szQst = szQst & " FROM " & gszDBName & "LAB_TEST_OE " & gszDBName
    szQst = szQst & " Where LAB_TEST_OE.LAB_TEST_IEN = " & SampleIEN
    szQst = szQst & " order by LAB_TEST_NAME"

    SearchLabPanel = TransferToListBox(szQst, plistbox, 1)

    Exit Function
  
ErrorHandler:
    DisplayLogErr "OE_SearchEngine: SearchLabPanel"

End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' LoadAllMeprs - Seaches the Meprs table
'
' Globals       none
' Call by ref   pListBox        CSComboBox to return SQL quiery results
' Call by val
'
' Returns       the data saved in the list box
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function LoadAllMeprs(plistbox As CSComboBox) As srcSearchResult

    Dim szOrderBy As String
    Dim szSelect As String
    Dim szQst As String
    Dim szSQL As String
    Dim gszDBName As String

    On Error GoTo ErrorHandler
    
    szSelect = "SELECT Distinct meprs.hospital_loc_name, meprs.meprs_code, meprs.meprs_ien, meprs.hospital_loc_ien"
    szQst = ""
    szQst = szQst & " From " & gszDBName & "MEPRS"
    szQst = szQst & " Where meprs.hospital_loc_name is not null"
    szQst = szQst & " AND meprs.meprs_code is not null"
    szQst = szQst & " AND meprs.meprs_ien <> 0"
    szOrderBy = " ORDER BY meprs.hospital_loc_name"
    szSQL = szSelect & szQst & szOrderBy
    
    LoadAllMeprs = TransferToListBox(szSQL, plistbox, 4)
    Exit Function

ErrorHandler:
    DisplayLogErr "OE_SearchEngine: LoadAllMeprs"
    
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' LoadAllMeprs - Seaches the Meprs table
'
' Globals       none
' Call by ref   pColl        Collection to hold results
' Call by val
'
' Returns       the data saved in the list box
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function LoadAllMeprsIntoCollection(pColl As ClinicCollection) As srcSearchResult

    Dim szOrderBy As String
    Dim szSelect As String
    Dim szQst As String
    Dim szSQL As String
    Dim gszDBName As String
    Dim oRecordSet As Recordset
    Dim szClinicIen As String
    Dim szClinicName As String
    Dim szMeprsIEN As String
    Dim szMeprsCode As String
    
    On Error GoTo ErrorHandler
    
    szSelect = "SELECT Distinct meprs.hospital_loc_name, meprs.meprs_code, meprs.meprs_ien, meprs.hospital_loc_ien"
    szQst = ""
    szQst = szQst & " From " & gszDBName & "MEPRS"
    szQst = szQst & " Where meprs.hospital_loc_name is not null"
    szQst = szQst & " AND meprs.meprs_code is not null"
    szQst = szQst & " AND meprs.meprs_ien <> 0"
    
    szOrderBy = " ORDER BY meprs.hospital_loc_name"
    szSQL = szSelect & szQst & szOrderBy
                                        
    Set oRecordSet = oDAL.OpenRecordset(szSQL)
    
    While Not oRecordSet.EOF                        'add each name and ien
        
        szClinicName = Trim$(oRecordSet.Fields(0))
        szMeprsCode = Trim$(oRecordSet.Fields(1))
        szMeprsIEN = Trim$(oRecordSet.Fields(2))
        szClinicIen = Trim$(oRecordSet.Fields(3))
        
        pColl.Add szClinicIen, szMeprsCode, szMeprsIEN, szClinicName      'add text list to list box
        oRecordSet.MoveNext                         'get the next record
    Wend
    LoadAllMeprsIntoCollection = srcErrNone
    Exit Function

ErrorHandler:
    DisplayLogErr "OE_SearchEngine: LoadAllMeprsIntoCollection"
    LoadAllMeprsIntoCollection = srcErrUnknown
    
End Function
Public Sub LoadAllLabOrdersIntoCollections(pSumm As Collection, _
                                                pLab As Collection, _
                                                hlErr As hlERRORS)

Dim Coll2               As Collection
Dim strSQL              As String
Dim SummaryRecord       As SummaryType
Dim Lab_Type            As LabType
Dim oRecordSet          As Recordset
    
    On Error GoTo ErrHandler

    Set Coll2 = New Collection
    strSQL = "select * from OE_LAB where IS_ORDER_OCCURRENCE = 0 and unit_number = " & gobjPatient.UnitNumber
    
    Set oRecordSet = oDAL.OpenRecordset(strSQL)

    Do While Not oRecordSet.EOF
        
        Set SummaryRecord = New SummaryType
    
        With SummaryRecord
            .data_id = Val(oRecordSet("OE_LAB_DATA_ID"))
            .OrderTypeCode = hlLAB_ORDERS_CODE
            .OrderTypeText = oRecordSet("ORDER_NAME") & vbNullString
            .OrderId = oRecordSet("ORDER_ID") & vbNullString
            .FullText = oRecordSet("ORDER_NAME") & vbNullString
            
            .Status = OrderSatusForDisplay(oRecordSet("STATUS") & vbNullString)
            If gobjShared.IsAppMode(modeITT) Then   '<SCR 37864
                .SigNeed = Val(oRecordSet("SIGN") & vbNullString)
            End If
        End With
        ' End: SCR #24287;
        
        pSumm.Add SummaryRecord
        
        Set Lab_Type = New LabType
        
        With Lab_Type
            .OrderId = oRecordSet("ORDER_ID") & vbNullString
            .OrderIen = oRecordSet("ORDER_IEN") & vbNullString
            .OrderName = oRecordSet("ORDER_NAME") & vbNullString
            .OrderingProviderLname = oRecordSet("ORDERED_BY") & vbNullString
            .Comment = oRecordSet("LAB_COMMENT") & vbNullString
            .CollectMethod = oRecordSet("COLLECT_METHOD") & vbNullString
            .ProcessPriority = oRecordSet("PROCESS_PRIORITY") & vbNullString
            .LabCollectPriority = oRecordSet("LAB_COLLECT_PRIORITY") & vbNullString
            .InPatientType = oRecordSet("PATIENT_TYPE") & vbNullString
            .ProviderComment = oRecordSet("RESULTS_COMMENTS") & vbNullString
            .ResultAssessment = Val(oRecordSet("NORMAL_STATUS") & vbNullString)
            
            If Not IsNull(oRecordSet("RESULTS")) Then
                If Trim$(oRecordSet("RESULTS")) <> "" Then
                    .WarningReason = oRecordSet("RESULTS")
                Else
                    .WarningReason = ""
                End If
            Else
                .WarningReason = ""
            End If
            
           '.APSpecimenColl.Count = # of specimen records
            With .Specimen
                .Name = oRecordSet("SPECIMEN_NAME") & vbNullString
                .IEN = oRecordSet("SPECIMEN_IEN") & vbNullString
                .Topography = oRecordSet("SPECIMEN_TOPOGRAPHY") & vbNullString
                .TopographyIen = oRecordSet("SPECIMEN_TOPOGRAPHY_IEN") & vbNullString
                .AnswerCode = oRecordSet("SPECIMEN_ANSWER_CODE") & vbNullString
                .AnswerText = oRecordSet("SPECIMEN_ANSWER_TEXT") & vbNullString
                .CapColorOrContainer = oRecordSet("SPECIMEN_CAPCOLOR_OR_CONTAINER") & vbNullString
            End With
            With .Timing
                .AdminTime = oRecordSet("TIMING_ADMIN_TIME") & vbNullString
                .Duration = oRecordSet("TIMING_DURATION") & vbNullString
                .Frequency = oRecordSet("TIMING_FREQUENCY") & vbNullString
                .StartDate = oRecordSet("TIMING_START_DATE") & vbNullString
                .StartTime = oRecordSet("TIME_ORDERED") & vbNullString
                .OnexOrContinued = oRecordSet("TIMING_ONE_X_OR_CONTINUED") & vbNullString
            End With
            
            .WarningHeader = oRecordSet("STATUS") & vbNullString  'Use this field to hold status, need to change later.
            
        End With

        Coll2.Add Lab_Type
        
        oRecordSet.MoveNext
        
    Loop
    
    Set pLab = Coll2
    
    Set oRecordSet = Nothing
    
    hlErr = hlERROR_NONE
    
    Exit Sub
  
ErrHandler:
    hlErr = hlERROR_GET_ORDERS_FAILED
    Set oRecordSet = Nothing
    gobjShared.ShowVBError Err.Number, Err.Description, "OE_SearchEngine.LoadAllLabOrdersIntoCollections", "OE", vbExclamation
    Exit Sub
    
    Err.Clear
    Resume
    
End Sub

Public Sub LoadAllRxOrdersIntoCollections(pSumm As Collection, _
                                                pRx As Collection, _
                                                hlErr As hlERRORS)

Dim Coll2          As Collection
Dim sSql           As String
Dim sSQLsub      As String
Dim SummaryRecord  As SummaryType
Dim rx_type        As RxType
Dim oRS            As Recordset
Dim oRSub          As Recordset

Dim bSkip As Boolean

    On Error GoTo ErrHandler

    Set Coll2 = New Collection
    
    '** This JOINT SQL cause problem for Refill Med order
    '** 2 record w/ same ORDER_ID exist in OE_RX with satus PENDING and ACTIVE
    '** but only one record exist in MEDICATIONS table with status ACTIVE
    '********* Result show both as ACTIVE
    
    'SF SCR 76468, 76736 Add order_ien to the query from the oe_rx table. 77760
'    sSql = "SELECT MEDICATIONS.*, OE_RX.ORDER_IEN AS ORDER_IEN, OE_RX.PATIENT_TYPE AS PATIENT_TYPE" _
'            & " From MEDICATIONS, OE_RX WHERE MEDICATIONS.ORDERNUMBER = OE_RX.ORDER_ID (+) and "
    
    sSql = "select * from medications where  "

    If gobjShared.IsAppMode(modeITT) Then   '<SCR 37864
        sSql = sSql & "  MEDICATIONS.StatusNCID not in (" & NCID_ORDERSTATUS_DISCONTINUED & ") and "
    End If
    
    '** 5/12/05 - FRANK
    '** Display only ACTIVE or PENDING Med
    If LenB(gobjShared.CmdLineSwitches("INPATIENT2")) > 0 Then
        sSql = sSql & "  MEDICATIONS.StatusNCID in (" & NCID_PARTIAL & "," & NCID_ORDERSTATUS_ACTIVE & "," & NCID_ORDERSTATUS_COMPLETE & ") and "
    End If
    
    sSql = sSql & "  MEDICATIONS.unit_number = " & gobjPatient.UnitNumber
    
    Set oRS = oDAL.OpenRecordset(sSql)

    Do While Not oRS.EOF
        
        '** SCR 71640 - 6/2/05 - Frank
        '** Display only Pending or Active Inaptient Med Orders
        '** Or Pending, Active and Completed w/in last 30 days for outpatient med orders
        bSkip = False
        If LenB(gobjShared.CmdLineSwitches("INPATIENT2")) > 0 Then
            If oRS.Fields("PATIENTTYPE") & vbNullString = "I" Then
                '** For Inpatient Med Orders, display only Pending and Active orders
                bSkip = oRS.Fields("STATUSNCID") & vbNullString = NCID_ORDERSTATUS_COMPLETE
            Else
                '** For Outpatient Med Orders, display Pending, Active and Completed w/in last 30 days
                If oRS.Fields("STATUSNCID") & vbNullString = NCID_ORDERSTATUS_COMPLETE Then
                    bSkip = (DateDiff("d", CDate(oRS.Fields("FILLDATE") & vbNullString), Date$) > 30)
                End If
            End If
           
        End If
        
        If Not bSkip Then
            Set SummaryRecord = New SummaryType
            
            With SummaryRecord
                .data_id = oRS("dataid")
                .OrderTypeCode = hlRX_ORDERS_CODE
                .Status = oRS("StatusNCID") & vbNullString
                .OrderId = oRS("OrderNumber") & vbNullString
                ' Quan 82845
                .ExpirationDate = oRS("ENDDATE") & vbNullString 'oRS("ORDEREXPIRATIONDATE") & vbNullString
                .LastFillDate = oRS("FillDate") & vbNullString
                .RemainingRF = oRS("REFILLSREMAINING") & vbNullString
                .OrderTypeText = oRS("NationalDrugCode") & vbNullString
                
                .Refill = oRS("REFILLS") & vbNullString ' SCR 107029, JQ, 6/6/07
                
                .FDBMedID = oRS("FDB_MEDID") & vbNullString
                
                If Not IsNull(oRS("OTC")) Then
                    .Modifiable = hlIS_NOT_MODIFIABLE
                Else
                    .Modifiable = hlIS_MODIFIABLE
                End If
                
                If .Status = NCID_ORDERSTATUS_ACTIVE Then  'If status is "Active", check if there is pending one
                    Dim Count As Integer
                    Count = getMultiOrderCount(.OrderId)
                    If Count > 1 Then
                        .Modifiable = hlIS_NOT_MODIFIABLE
                    End If
                End If
                
                .ParentDataID = oRS("PARENT_DATA_ID").Value & vbNullString
                
                'To support medications added through the medications module.
                If oRS("NON_CODED_MEDICATION_NAME") & vbNullString = "" Then
                    sSQLsub = "Select Distinct DRUG_NAME from RX_LIST Where drug_ien = " & oRS("MedicationNCID")
                    Set oRSub = oDAL.OpenRecordset(sSQLsub)
                    If Not oRSub.EOF Then
                        .FullText = oRSub("DRUG_NAME") & vbNullString
                    End If
                    Set oRSub = Nothing
                Else
                    .FullText = oRS("NON_CODED_MEDICATION_NAME") & vbNullString
                End If
                                
                'SF SCR 76468, 76736 Add order_ien to the query from the oe_rx table.
''                .OrderIen = oRS("ORDER_IEN") & vbNullString
                '77760
                .ParentDataID = oRS("PARENT_DATA_ID").Value & vbNullString
                
                If oRS("PATIENTTYPE") & "" = "I" Then
                    .PatientType = "Inpatient"
                Else
                    .PatientType = "Outpatient"
                End If
                
                If gobjShared.IsAppMode(modeITT) Then   '<SCR 37864
                    Dim ChildCap As Integer
                    sSQLsub = "Select SIGN,CHILD_CAP from oe_rx where OE_RX_DATA_ID = " & .data_id
                    Set oRSub = oDAL.OpenRecordset(sSQLsub)
                    
                    If Not oRSub.EOF Then
                         .SigNeed = Val(oRSub("SIGN") & vbNullString)
                         ChildCap = Val(oRSub("CHILD_CAP") & vbNullString)
                    End If
                    Set oRSub = Nothing
                End If
            End With
        
        
            pSumm.Add SummaryRecord
            
            Call Deserialize_RX(oRS, rx_type, True)
            rx_type.OrderName = SummaryRecord.FullText
            
            Coll2.Add rx_type
        End If
            
        oRS.MoveNext
        

    Loop
    
    Set pRx = Coll2
    Set oRS = Nothing
    
    hlErr = hlERROR_NONE
    
    Exit Sub
  
ErrHandler:
    hlErr = hlERROR_GET_ORDERS_FAILED
    Set oRS = Nothing
    gobjShared.ShowVBError Err.Number, Err.Description, "OE_SearchEngine.LoadAllRXOrdersIntoCollections", "OE", vbExclamation
    Exit Sub
    
    Err.Clear
    Resume
    
End Sub

Public Sub LoadAllRadOrdersIntoCollections(pSumm As Collection, _
                                                pRad As Collection, _
                                                hlErr As hlERRORS)

Dim Coll2               As Collection
Dim strSQL              As String
Dim SummaryRecord       As SummaryType
Dim rad_type            As RadType
Dim oRecordSet          As Recordset

    On Error GoTo ErrHandler

    Set Coll2 = New Collection
    
    strSQL = "select * from oe_rad where IS_ORDER_OCCURRENCE = 0 and unit_number = " & gobjPatient.UnitNumber
    
    Set oRecordSet = oDAL.OpenRecordset(strSQL)

    Do While Not oRecordSet.EOF
    
        Set SummaryRecord = New SummaryType
        
        ' Begin:SCR #24287;   Developer: Jane Shen 06/03/2002 09:19 AM
        'SCR Description: Changes made in Lab and Rad to match with Order Management results form.
        With SummaryRecord
            .data_id = oRecordSet("oe_rad_data_id")
            .OrderTypeCode = hlRAD_ORDERS_CODE
            .FullText = oRecordSet("ORDER_NAME")
            .OrderId = oRecordSet("ORDER_ID")
            If Not IsNull(oRecordSet("STATUS")) Then
                .Status = OrderSatusForDisplay(oRecordSet("STATUS"))
            End If
            If gobjShared.IsAppMode(modeITT) Then   '<SCR 37864
                If Not IsNull(oRecordSet("SIGN")) Then .SigNeed = oRecordSet("SIGN")
            End If
        End With
        ' End: SCR #24287;
        
        pSumm.Add SummaryRecord
        
        Set rad_type = New RadType
        With rad_type
        
            .OrderName = oRecordSet("ORDER_NAME") & vbNullString
            .OrderIen = oRecordSet("ORDER_IEN") & vbNullString
            .Comment = oRecordSet("RAD_COMMENT") & vbNullString
            .OrderId = oRecordSet("ORDER_ID") & vbNullString
            .OrderingProviderLname = oRecordSet("ORDERED_BY") & vbNullString
            .TransportMode = oRecordSet("TRANSPORT_MODE") & vbNullString
            .ClinicalImpressionText = oRecordSet("CLINICAL_IMPRESSION_TEXT") & vbNullString
            .ActionReason = oRecordSet("RESULTS") & vbNullString
            .PortableAllowed = oRecordSet("STATUS") & vbNullString
            .WarningReason = oRecordSet("OE_RAD_DATA_ID") & vbNullString
            .ProcessPriority = oRecordSet("PROCESS_PRIORITY") & vbNullString
            .WetReadingWanted = oRecordSet("WET_READING_WANTED") & vbNullString
            .InPatientType = oRecordSet("PATIENT_TYPE") & vbNullString
            .ProviderComment = oRecordSet("RESULTS_COMMENTS") & vbNullString
            .ResultAssessment = Val(oRecordSet("NORMAL_STATUS") & vbNullString)
            
            With .Timing
                .StartDate = oRecordSet("timing_start_date") & vbNullString
                .AdminTime = oRecordSet("TIMING_ADMIN_TIME") & vbNullString
                .Frequency = oRecordSet("TIMING_FREQUENCY") & vbNullString
                .Duration = oRecordSet("TIMING_DURATION") & vbNullString
                .OnexOrContinued = oRecordSet("TIMING_ONE_X_OR_CONTINUED") & vbNullString
            End With

        End With
        
        Coll2.Add rad_type
        
        oRecordSet.MoveNext
        
    Loop
    
    Set pRad = Coll2
    
    Set oRecordSet = Nothing
    
    hlErr = hlERROR_NONE
    
    Exit Sub
  
ErrHandler:
    hlErr = hlERROR_GET_ORDERS_FAILED
    Set oRecordSet = Nothing
    gobjShared.ShowVBError Err.Number, Err.Description, "OE_SearchEngine.LoadAllRadOrdersIntoCollections", "OE", vbExclamation
    Exit Sub
    
    Err.Clear
    Resume
    
End Sub


Friend Function TransferToListBox(szSQL As String, _
                                    plistbox As Control, _
                                    iFieldMaxCount As Integer) As srcSearchResult

    Dim oRecordSet As Recordset
    Dim iFieldIndex As Integer
    Dim sListBoxLine As String
    Dim sDelimiter As String
    Dim sControlName As String
    Dim sWidth As Single

    On Error GoTo ErrorHandler
    
    If iFieldMaxCount < 1 Then iFieldMaxCount = 1       'check iFieldMaxCount
    
    'The following code supports more than one type of ComboBox control
    
    sControlName = TypeName(plistbox)
    Select Case UCase$(sControlName)
        Case "COMBOBOX"                                 'check for VB ComboBox
            sDelimiter = vbTab
        Case "CSCOMBOBOX"                               'check for Cresent ComboBox
            sDelimiter = plistbox.ColDelim              'Delimiter character
        Case "LISTBOX"                                  'added 9-20-99; AAB
            sDelimiter = vbTab
        Case Else
            sDelimiter = vbTab

    End Select
    
    TransferToListBox = srcErrUnknown
    
    plistbox.Clear                                      'clear listbox
    Set oRecordSet = oDAL.OpenRecordset(szSQL)
    
    sWidth = 0
    While Not oRecordSet.EOF                        'add each name and ien
        sListBoxLine = ""
        iFieldIndex = 0
        'Begin:SCR #23953;   Developer: Jane Shen 05/17/2002 06:14 PM
        If Not Trim$(oRecordSet.Fields(iFieldIndex)) = "" Then
          While iFieldIndex < iFieldMaxCount          'Build the text list for the list box
            If iFieldIndex = 0 Then
                If plistbox.Parent.TextWidth(Trim$(oRecordSet.Fields(iFieldIndex))) > sWidth Then
                    sWidth = plistbox.Parent.TextWidth(oRecordSet.Fields(iFieldIndex) & "       ")
                End If
            End If
     
            If Not IsNull(oRecordSet.Fields(iFieldIndex)) Then
                sListBoxLine = sListBoxLine & Trim$(oRecordSet.Fields(iFieldIndex)) & sDelimiter
            Else
                sListBoxLine = sListBoxLine & sDelimiter
            End If
            
            iFieldIndex = iFieldIndex + 1
          Wend
        
          plistbox.AddItem sListBoxLine               'add text list to list box
        End If
        'End: SCR #23953;
        oRecordSet.MoveNext                         'get the next record
    Wend

    If UCase$(sControlName) <> "LISTBOX" Then
        If sWidth > plistbox.Width Then
            plistbox.ListBoxWidth = sWidth / Screen.TwipsPerPixelX
        Else
            plistbox.ListBoxWidth = plistbox.Width / Screen.TwipsPerPixelX
        End If
    End If

    If plistbox.ListCount > 0 Then                      'check for empty ListBox
        TransferToListBox = srcErrNone
    End If
    
    Exit Function

ErrorHandler:
     DisplayLogErr "OE_SearchEngine: TransferToListBox"
    Exit Function
     
    Err.Clear
    Resume

End Function

Friend Function TransferToRadLabListBox(szSQL As String, plistbox As Control, _
                                OrderType As hlORDERTYPECODE) As srcSearchResult

    Dim oRecordSet As Recordset
    Dim sListBoxLine As String
    Dim sDelimiter As String
    Dim sControlName As String
    Dim sWidth As Single
    
    Dim pCol As New Collection
    Dim pItem As LabRadItem
    
    Dim arItems() As SearchItem
    Dim lngArray As Long
    Dim lngStep As Long
    Dim i As Integer
    
    On Error GoTo ErrorHandler

    
    sControlName = TypeName(plistbox)
    Select Case UCase$(sControlName)
        Case "COMBOBOX"                                 'check for VB ComboBox
            sDelimiter = vbTab
        Case "CSCOMBOBOX"                               'check for Cresent ComboBox
            sDelimiter = plistbox.ColDelim              'Delimiter character
        Case "LISTBOX"                                  'added 9-20-99; AAB
            sDelimiter = vbTab
        Case Else
            sDelimiter = vbTab
    End Select
    
    TransferToRadLabListBox = srcErrUnknown
    
    plistbox.Clear                                      'clear listbox
    Set oRecordSet = oDAL.OpenRecordset(szSQL)
    
    
    While Not oRecordSet.EOF                        'add each name and ien
        
        pItem.szDisplayName = ""
        pItem.szIEN = Trim$(oRecordSet.Fields(1))
        pItem.szName = Trim$(oRecordSet.Fields(0))
        If OrderType = hlLAB_ORDERS_CODE Then
            pItem.szPanelFlag = Trim$(oRecordSet.Fields("panel_flag") & " ")
            pItem.szSubscript = Trim$(oRecordSet.Fields("subscript") & " ")
        End If
        If pCol.Count > 0 Then
            If pItem.szIEN = pCol.Item(pCol.Count).szIEN Then
                If LenB(pItem.szSynonym) > 0 And pItem.szSynonym < pItem.szName And pCol.Item(pCol.Count).szDisplayName <> pItem.szSynonym & " " & pItem.szName Then
                    pItem.szDisplayName = pItem.szSynonym & " " & pItem.szName
                    pCol.Remove pCol.Count
                    pCol.Add pItem
                    arItems(lngArray).szDisplayName = pItem.szDisplayName
                End If
            Else
                pItem.szDisplayName = pItem.szName
                pCol.Add pItem
                lngArray = pCol.Count - 1
                ReDim Preserve arItems(0 To lngArray)
                arItems(lngArray).szDisplayName = pItem.szDisplayName
                arItems(lngArray).iCollectionIndex = pCol.Count
            End If
        Else
            pItem.szDisplayName = pItem.szName
            pCol.Add pItem
            lngArray = pCol.Count - 1
            ReDim Preserve arItems(0 To lngArray)
            arItems(lngArray).szDisplayName = pItem.szDisplayName
            arItems(lngArray).iCollectionIndex = pCol.Count
        End If
        
        oRecordSet.MoveNext
    Wend
    
    sWidth = 0
    For lngStep = 0 To lngArray
        
        i = arItems(lngStep).iCollectionIndex
        If OrderType = hlRAD_ORDERS_CODE Then
            sListBoxLine = pCol(i).szDisplayName & sDelimiter & pCol(i).szIEN & sDelimiter & pCol(i).szName
        Else
            sListBoxLine = pCol(i).szDisplayName & sDelimiter & pCol(i).szIEN & sDelimiter & pCol(i).szPanelFlag & sDelimiter & pCol(i).szSubscript & sDelimiter & pCol(i).szName
        End If
        
        plistbox.AddItem sListBoxLine               'add text list to list box
        
        If plistbox.Parent.TextWidth(pCol(i).szDisplayName) > sWidth Then
            sWidth = plistbox.Parent.TextWidth(pCol(i).szDisplayName & "       ")
        End If
    Next lngStep
    
    If sWidth > plistbox.Width Then
        plistbox.ListBoxWidth = sWidth / Screen.TwipsPerPixelX
    Else
        plistbox.ListBoxWidth = plistbox.Width / Screen.TwipsPerPixelX
    End If

    If plistbox.ListCount > 0 Then                      'check for empty ListBox
        TransferToRadLabListBox = srcErrNone
    End If
    
    Exit Function

ErrorHandler:
     DisplayLogErr "OE_SearchEngine: TransferToRadLabListBox"

End Function
Friend Function GetLabCollection(objSpecimen As TypeSpecimen) As Integer
    
    Dim oRecordSet As Recordset
    Dim szQst As String
    Dim gszDBName As String
  
    On Error GoTo ErrorHandler:
  
    szQst = "select TEST_IEN, TEST_NAME, COLLECTION_SAMPLE,CONTAINER, SPECIMEN"
    szQst = szQst & " From " & gszDBName & "lab_collection"
    szQst = szQst & " Where TEST_NAME LIKE '" & objSpecimen.Name & "%'"
    
    Set oRecordSet = oDAL.OpenRecordset(szQst)
    
    If Not oRecordSet.EOF Then
        objSpecimen.IEN = Trim$(oRecordSet.Fields(0))
        objSpecimen.Name = Trim$(oRecordSet.Fields(2))
        objSpecimen.CapColorOrContainer = Trim$(oRecordSet.Fields(3))
        ' Begin:SCR #1022;   Developer: Jane Shen 11/08/2002 03:39 PM
        If Not IsNull(Trim$(oRecordSet.Fields(4))) Then
            objSpecimen.Topography = Trim$(oRecordSet.Fields(4))
        End If
        ' End: SCR #1022;
        GetLabCollection = 1
    Else
        GetLabCollection = 0
    End If

    Exit Function
  
ErrorHandler:
    DisplayLogErr "OE_SearchEngine: GetLabSpecimenData"

End Function
Private Sub Class_Initialize()
   Set oDAL = pobjConn.CHCSII_DAS '<SCR 37864
End Sub

Private Sub Class_Terminate()
 Set oDAL = Nothing
End Sub


