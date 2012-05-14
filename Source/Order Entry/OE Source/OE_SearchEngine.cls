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
    srcKeywordTooShort
End Enum

Public Type RxItem
    szIEN As String
    szName As String
    szDefUnit As String
    szSynonym As String
    ip_op_indicator As Integer
    drug_check_disabled As Integer
    non_formulary As Integer
    Message As String
    szDisplayName As String
    Stocked As Integer
    formulary_group_ien As Integer
End Type

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

Private Const mcINDENT = "     "
Private oDAL As ICHCSII_DAS
Private szWildcard As String




'Public Function bLogoffDatabase() As Boolean
'    Dim bDisconnect As Boolean
'
'    bLogoffDatabase = False
'
'    On Error GoTo Errorhandler
'
'    If Not oDAL Is Nothing Then
'        bDisconnect = oDAL.CacheConnection("")
'        Set oDAL = Nothing
'    End If
'
'    bLogoffDatabase = True
'    Exit Function
'
'Errorhandler:
'    DisplayLogErr "OE_SearchEngine: bLogoffDatabase"
'
'End Function
'
'Public Function bLogonDatabase() As Boolean
'    Dim iResult As Integer
'    Dim szResult As String
'    Dim szUNIXPassword As String
'    Dim szDBString As String
'    Dim szNewConnectString As String
'    Dim szMsg As String
'    Dim bConnect As Boolean
'
'    bLogonDatabase = False
'
'    On Error GoTo Errorhandler
'
'    Set oDAL = New CDataAccess
'
'
'    szDBString = GetAppProfile("ADO", "PWD", "")
'    szNewConnectString = szCreateConnectString(szDBString)
'
'    If szNewConnectString = "" Then
'        MsgBxARMd "Clinical Integrated Workstation INI file missing DB Connect String", 0, "Startup Configuration Error"
'        bLogonDatabase = False
'        Exit Function
'    End If
'
'    bConnect = oDAL.CacheConnection(szNewConnectString)
'    If bConnect Then
'        bLogonDatabase = True
'    Else
'        MsgBxARMd "The Data Base Logon failed.", 0, "CacheConnection logon failed"
'    End If
'    Exit Function
'
'Errorhandler:
'   DisplayLogErr "OE_SearchEngine: bLogonDatabase"
'
'End Function
Function szCreateConnectString(szPassword As String) As String

'KELLY DEAD 7-27-01     Dim szConnectString As String
'KELLY DEAD 7-27-01    szConnectString = "Driver={Microsoft Access Driver (*.mdb)};" & _
'KELLY DEAD 7-27-01                     "Dbq=hl7order.mdb;" & _
'KELLY DEAD 7-27-01                      "DefaultDir=D:\;" & _
'KELLY DEAD 7-27-01                      "Uid=Admin;Pwd=;"
                     
    szCreateConnectString = GetAppProfile("ADO", "ConnectStringOE", "")
  '  szCreateConnectString = szConnectString
    
        
End Function
Friend Function Search(plistbox As CSComboBox, _
                        OrderType As hlORDERTYPECODE, _
                        szSearchText As String, _
                        Optional srcSearchWhat As srcSynonyms = srcboth, _
                        Optional srcSection As Variant = "", _
                        Optional szFormularyGroupIEN As String) As srcSearchResult
    
Dim szOrderBy As String
Dim szSelect As String
Dim szQst As String
Dim szSQL As String
Dim szSectionSQL As String
Dim gszDBName As String
Dim iNumberofFields As Integer
    
               'srcSearchWhat = srcsimple                      'debug
               'srcSearchWhat = srcSynonym                     'debug
               'srcSearchWhat = srcBoth                        'debug
               'srcSection = "MAGNETIC RESONANCE IMAGING"      'debug
               'srcSection = 1                                 'debug
    
               ''5-15-01; AAB #12754
               'Screen.ActiveForm.MousePointer = vbHourglass    'display hour glass cursor
               'Removed 6-15-01 AAB #12754 and #12790
    If (LenB(szSearchText) < 4) And (LenB(srcSection) = 0) Then
        MsgBxARMd "Keyword must be at least two characters long.", vbCritical, "Search Criteria Error"
        Search = srcKeywordTooShort
        GoTo ErrorHandler
    End If
    
    On Error GoTo ErrorHandler
    iNumberofFields = 2
'    #If coiAccess Then
'        srcSearchWhat = srcSimple
'    #End If
    Select Case OrderType
    
    Case hlRX_ORDERS_CODE:
''      #If coiPickList = True Then
        iNumberofFields = 7
''      #Else
''      iNumberofFields = 2
''      #End If
               Select Case srcSearchWhat 'Simple, Synonym, Both
                

                Case srcSimple:
''                  #If coiPickList = True Then
                    szSelect = "SELECT RX_LIST.drug_name, RX_LIST.drug_ien, RX_LIST.default_unit, RX_FORMULARY.ip_op_indicator, RX_LIST.drug_check_disabled,  RX_FORMULARY.non_formulary, RX_FORMULARY.message"
                    szQst = szQst & " FROM " & gszDBName & " RX_LIST, RX_FORMULARY"
''                  #Else
''                      szSelect = "SELECT RX_LIST.drug_name, RX_LIST.drug_ien"
''                      szQst = szQst & " FROM " & gszDBName & " RX_LIST"
''                  #End If
                    szQst = szQst & " WHERE RX_LIST.drug_name like '" & szWildcard & szSearchText & szWildcard & "'" ' string match
                    szQst = szQst & " AND RX_LIST.drug_name is not null" 'drug name must not be empty
                    szQst = szQst & " AND RX_LIST.drug_ien <> 0" 'drug IEN must be present
''                  #If coiPickList = True Then
                    szQst = szQst & " AND RX_FORMULARY.drug_ien = RX_LIST.drug_ien "
                    szQst = szQst & " AND RX_FORMULARY.formulary_group_ien = " & szFormularyGroupIEN
                    szQst = szQst & " AND RX_LIST.ndc_indicator=2 AND RX_FORMULARY.inactive_flag=0"
''                  #End If
                    szOrderBy = " ORDER BY 1;"
                    szSQL = szSelect & szQst & szOrderBy
                    
                Case srcSynonym:
                    szSelect = "SELECT Distinct RX_LIST.drug_name, RX_LIST.drug_ien"
                    szQst = ""
                    szQst = szQst & " FROM " & gszDBName & "RX_LIST, " & gszDBName & "RX_SYNONYM"
                    szQst = szQst & " WHERE RX_LIST.drug_ien = rx_synonym.drug_ien"
                    szQst = szQst & " AND rx_synonym.synonym_name like '" & szWildcard & szSearchText & szWildcard & "'" ' string match
                    szQst = szQst & " AND rx_synonym.synonym_name is not null" 'synonym name must not be empty
                    szQst = szQst & " AND rx_synonym.drug_ien <> 0" 'drug IEN must be present
                    szOrderBy = " ORDER BY RX_LIST.drug_name;"
                    szSQL = szSelect & szQst & szOrderBy
                                        
                Case srcboth:
            
                    iNumberofFields = 9
                    szSelect = "SELECT DISTINCT RX_LIST.drug_name, RX_LIST.drug_ien, RX_LIST.default_unit, RX_FORMULARY.ip_op_indicator, RX_LIST.drug_check_disabled,  RX_FORMULARY.non_formulary, RX_FORMULARY.message, NULL as MinSyn, 1 as STOCKED, RX_FORMULARY.formulary_group_ien"
                    szQst = " FROM " & gszDBName & " RX_LIST, " & gszDBName & " RX_FORMULARY"
                    szQst = szQst & " WHERE RX_LIST.drug_name like '" & szWildcard & szSearchText & szWildcard & "'"
                    szQst = szQst & " AND RX_LIST.drug_name is not null" 'drug name must be present
                    szQst = szQst & " AND RX_LIST.drug_ien <> 0" 'drug IEN must be present
                    szQst = szQst & " AND RX_FORMULARY.drug_ien = RX_LIST.drug_ien "
                    szQst = szQst & " AND RX_FORMULARY.inactive_flag=0"
                    szQst = szQst & " AND RX_FORMULARY.non_formulary=0"
                    szQst = szQst & " UNION"
                    szQst = szQst & " SELECT DISTINCT RX_LIST.drug_name, RX_LIST.drug_ien, RX_LIST.default_unit, RX_FORMULARY.ip_op_indicator, RX_LIST.drug_check_disabled,  RX_FORMULARY.non_formulary, RX_FORMULARY.message, MIN(rx_synonym.synonym_name),  1 as STOCKED, RX_FORMULARY.formulary_group_ien"
                    szQst = szQst & " FROM " & gszDBName & " RX_LIST, " & gszDBName & " RX_FORMULARY, " & gszDBName & " RX_SYNONYM"
                    szQst = szQst & " WHERE RX_LIST.drug_ien = rx_synonym.drug_ien"
                    szQst = szQst & " AND rx_synonym.synonym_name like '" & szWildcard & szSearchText & szWildcard & "'"
                    szQst = szQst & " AND RX_FORMULARY.drug_ien = RX_LIST.drug_ien "
                    szQst = szQst & " AND RX_FORMULARY.inactive_flag=0"
                    szQst = szQst & " AND RX_FORMULARY.non_formulary= 0 "
                    szQst = szQst & " GROUP BY RX_LIST.drug_name, RX_LIST.drug_ien, RX_LIST.default_unit, RX_FORMULARY.ip_op_indicator, RX_LIST.drug_check_disabled,RX_FORMULARY.non_formulary, RX_FORMULARY.message, RX_FORMULARY.formulary_group_ien"
                    szOrderBy = " Order By 1 ASC, 9 DESC;"
                    szSQL = szSelect & szQst & szOrderBy
                                                            
            End Select 'Simple, Synonym, Both
            
        '===============================================================================
        
        Case hlRAD_ORDERS_CODE:
            
            If LenB(srcSection) > 0 Then
                szSectionSQL = " AND RAD_TEST.rad_test_type = '" & srcSection & "'"
            Else
                szSectionSQL = ""
            End If
            
            Select Case srcSearchWhat 'Simple, Synonym, Both
            
                Case srcSimple:
           '        szSelect = "SELECT Distinct RAD_TEST.rad_test_name, RAD_TEST.rad_test_ien, RAD_TEST.rad_test_type"
                    szSelect = "SELECT Distinct RAD_TEST.rad_test_name, RAD_TEST.rad_test_ien"
                    szQst = ""
''                  #If coiPickList = True Then
                    szQst = szQst & " FROM  " & gszDBName & "RAD_TEST, " & gszDBName & " RAD_TEST_LOC "
''                  #Else
''                     szQst = szQst & " FROM  " & gszDBName & "RAD_TEST "
''                  #End If
                    '#If coiAccess Then
                    szQst = szQst & " WHERE rad_test_name like '" & szWildcard & szSearchText & szWildcard & "'" 'string match
                    '#Else
                    '    szQst = szQst & " WHERE rad_test_name like '%" & szSearchText & "%'" 'string match
                    '#End If
                    szQst = szQst & " AND rad_test_name is not null"
''                  #If coiPickList = True Then
                    szQst = szQst & " AND suppress = 1"
                    szQst = szQst & " AND inactive_flag = 0"
                    szQst = szQst & " AND RAD_TEST.rad_test_ien = RAD_TEST_LOC.rad_test_ien"
                    szQst = szQst & " AND RAD_TEST.type_code > 0"
''                  #End If
                    szQst = szQst & szSectionSQL
                    szOrderBy = " ORDER BY rad_test_name;"
                    szSQL = szSelect & szQst & szOrderBy

                Case srcSynonym:
                    'szSelect = "SELECT Distinct RAD_TEST.rad_test_name, RAD_TEST.rad_test_ien, RAD_TEST.rad_test_type"
                    szSelect = "SELECT Distinct RAD_TEST.rad_test_name, RAD_TEST.rad_test_ien"
                    szQst = ""
''                  #If coiPickList = True Then
                    szQst = szQst & " FROM  " & gszDBName & "RAD_TEST_SYNONYM, " & gszDBName & " RAD_TEST, " & gszDBName & " RAD_TEST_LOC "
''                  #Else
''                  szQst = szQst & " FROM  " & gszDBName & "RAD_TEST_SYNONYM , " & gszDBName & " RAD_TEST"
''                  #End If
                    szQst = szQst & " WHERE RAD_TEST.rad_test_ien = RAD_TEST_SYNONYM.rad_test_ien"
                    '#If coiAccess Then
                        szQst = szQst & " AND RAD_TEST_SYNONYM.synonym_name like '" & szWildcard & szSearchText & szWildcard & "'" 'string match
                    '#Else
                    '    szQst = szQst & " AND RAD_TEST_SYNONYM.synonym_name like '%" & szSearchText & "%'" 'string match
                    '#End If
                    szQst = szQst & " AND RAD_TEST_SYNONYM.synonym_name is not null" 'synonym name must not be empty"
                    szQst = szQst & " AND RAD_TEST_SYNONYM.rad_test_ien <> 0" 'RAD test IEN meust be present
''                  #If coiPickList = True Then
                    szQst = szQst & " AND suppress = 1"
                    szQst = szQst & " AND inactive_flag = 0"
                    szQst = szQst & " AND RAD_TEST.rad_test_ien = RAD_TEST_LOC.rad_test_ien"
                    szQst = szQst & " AND RAD_TEST.type_code > 0"
''                  #End If
                    szQst = szQst & szSectionSQL
                    szOrderBy = " ORDER BY 1;"
                    szSQL = szSelect & szQst & szOrderBy
                    
                Case srcboth:
''                  #If coiPickList = True Then
                    iNumberofFields = 3
''                  #Else
''                      iNumberofFields = 2
''                  #End If
                    szSelect = "SELECT Distinct RAD_TEST.rad_test_name, RAD_TEST.rad_test_ien, NULL as MinSyn"
                    szQst = ""
''                  #If coiPickList = True Then
                    szQst = szQst & " FROM  " & gszDBName & "RAD_TEST, " & gszDBName & " RAD_TEST_LOC "
''                  #Else
''                     szQst = szQst & " FROM  " & gszDBName & "RAD_TEST "
''                  #End If
                    '#If coiAccess Then
                    szQst = szQst & " WHERE RAD_TEST.rad_test_name like '" & szWildcard & szSearchText & szWildcard & "'"
                    '#Else
                    '    szQst = szQst & " WHERE RAD_TEST.rad_test_name like '%" & szSearchText & "%'"
                    '#End If
                    szQst = szQst & " AND RAD_TEST.rad_test_name is not null"
                    szQst = szQst & " AND RAD_TEST.Rad_test_ien <> 0"
''                  #If coiPickList = True Then
                    szQst = szQst & " AND suppress = 1"
                    szQst = szQst & " AND inactive_flag = 0"
                    szQst = szQst & " AND RAD_TEST.rad_test_ien = RAD_TEST_LOC.rad_test_ien"
                    szQst = szQst & " AND RAD_TEST.type_code > 0"
''                  #End If
                    szQst = szQst & szSectionSQL
                    szQst = szQst & " UNION"
''                  #If coiPickList = True Then
                    szQst = szQst & " SELECT Distinct RAD_TEST.rad_test_name, RAD_TEST.rad_test_ien, Min(RAD_TEST_SYNONYM.synonym_name)"
                    szQst = szQst & " FROM  " & gszDBName & "RAD_TEST_SYNONYM, " & gszDBName & " RAD_TEST, " & gszDBName & " RAD_TEST_LOC "
''                  #Else
''                      szQst = szQst & " SELECT Distinct RAD_TEST.rad_test_name, RAD_TEST.rad_test_ien, RAD_TEST_SYNONYM.synonym_name"
''                      szQst = szQst & " FROM  " & gszDBName & "RAD_TEST_SYNONYM , " & gszDBName & " RAD_TEST"
''                  #End If
                    szQst = szQst & " WHERE RAD_TEST.rad_test_ien = RAD_TEST_SYNONYM.rad_test_ien"
                    '#If coiAccess Then
                    szQst = szQst & " AND RAD_TEST_SYNONYM.synonym_name like '" & szWildcard & szSearchText & szWildcard & "'"
                    '#Else
                    '    szQst = szQst & " AND RAD_TEST_SYNONYM.synonym_name like '%" & szSearchText & "%'"
                    '
                    '#End If
                    
''                  #If coiPickList = True Then
                    szQst = szQst & " AND suppress = 1"
                    szQst = szQst & " AND inactive_flag = 0"
                    szQst = szQst & " AND RAD_TEST.rad_test_ien = RAD_TEST_LOC.rad_test_ien"
                    szQst = szQst & " AND RAD_TEST.type_code > 0"
''                  #End If
                    
                    szQst = szQst & szSectionSQL
''                  #If coiPickList = True Then
                    szQst = szQst & " GROUP BY RAD_TEST.rad_test_name, RAD_TEST.rad_test_ien"
''                  #End If
''                  Tracker 18886
                    szQst = szQst & " UNION"
                    szQst = szQst & " SELECT /*+ FULL(CPT_CODES) */ Distinct RAD_TEST.rad_test_name, RAD_TEST.rad_test_ien, NULL as MinSyn"
                    szQst = szQst & " FROM  " & gszDBName & "RAD_TEST, " & gszDBName & "RAD_TEST_LOC, " & gszDBName & "CPT_CODES "
                    szQst = szQst & " WHERE RAD_TEST.cpt_code = CPT_CODES.code"
                    szQst = szQst & " AND CPT_CODES.dod_description like '" & szWildcard & szSearchText & szWildcard & "'"
                    szQst = szQst & " AND RAD_TEST.rad_test_name is not null"
                    szQst = szQst & " AND RAD_TEST.Rad_test_ien <> 0"
                    szQst = szQst & " AND suppress = 1"
                    szQst = szQst & " AND inactive_flag = 0"
                    szQst = szQst & " AND RAD_TEST.rad_test_ien = RAD_TEST_LOC.rad_test_ien"
                    szQst = szQst & " AND RAD_TEST.type_code > 0"
                    szQst = szQst & szSectionSQL

                    szOrderBy = " Order By 1;"
                    szSQL = szSelect & szQst & szOrderBy
                
        End Select 'Simple, Synonym, Both
        
        '===============================================================================

        Case hlLAB_ORDERS_CODE:
        
            ' #If coiFuture = False Then
            '     iNumberofFields = 3
            ' #Else
            iNumberofFields = 4
            '  #End If
            
            If LenB(srcSection) > 0 Then
''              #If coiPickList = False Then
''                  szSectionSQL = " AND LAB_TEST_WTYPE.test_type_name = '" & srcSection & "'"
''              #Else
                szSectionSQL = " AND LAB_TEST_OE.subscript = '" & srcSection & "'"
''              #End If
            Else
                szSectionSQL = ""
            End If
            
            Select Case srcSearchWhat 'Simple, Synonym, Both
            
                Case srcSimple:
''                  #If coiPickList = False Then
''                      szSelect = "SELECT Distinct LAB_TEST_WTYPE.lab_test_name, LAB_TEST_WTYPE.lab_test_ien, LAB_TEST_WTYPE.panel_flag, LAB_TEST_WTYPE.subscript"
''                  #Else
                    szSelect = "SELECT Distinct LAB_TEST_OE.lab_test_name, LAB_TEST_OE.lab_test_ien, LAB_TEST_OE.panel_flag, LAB_TEST_OE.subscript"
''                  #End If
                    szQst = ""
''                  #If coiPickList = False Then
''                      szQst = szQst & " FROM  " & gszDBName & "LAB_TEST_WTYPE"
''                      szQst = szQst & " WHERE LAB_TEST_WTYPE.lab_test_name like '" & szWildcard & szSearchText & szWildcard & "'"
''                      szQst = szQst & " AND LAB_TEST_WTYPE.lab_test_name is not null"
''                      szQst = szQst & " AND LAB_TEST_WTYPE.lab_test_ien <> 0"
''                  #Else
                    szQst = szQst & " FROM  " & gszDBName & "LAB_TEST_OE"
                    szQst = szQst & " WHERE LAB_TEST_OE.lab_test_name like '" & szWildcard & szSearchText & szWildcard & "'"
                    szQst = szQst & " AND LAB_TEST_OE.lab_test_name is not null"
                    szQst = szQst & " AND LAB_TEST_OE.lab_test_ien <> 0"
                    szQst = szQst & " AND LAB_TEST_OE.Inactive_Flag = 0"
                    szQst = szQst & " AND LAB_TEST_OE.QC_Only = 0 "
                    szQst = szQst & " AND LAB_TEST_OE.test_type_code > 3"
''                  #End If
                    
                    szQst = szQst & szSectionSQL
                    szOrderBy = " Order by  1;"
                    szSQL = szSelect & szQst & szOrderBy
                    
                Case srcSynonym:
                  '  #If coiFuture = False Then
                  '      szSelect = "SELECT Distinct LAB_TEST_WTYPE.lab_test_name, LAB_TEST_WTYPE.lab_test_ien, LAB_TEST_WTYPE.panel_flag"
                  '  #Else
''                  #If coiPickList = False Then
''                      szSelect = "SELECT Distinct LAB_TEST_WTYPE.lab_test_name, LAB_TEST_WTYPE.lab_test_ien, LAB_TEST_WTYPE.panel_flag, LAB_TEST_WTYPE.subscript"
''                  #Else
                    szSelect = "SELECT Distinct LAB_TEST_OE.lab_test_name, LAB_TEST_OE.lab_test_ien, LAB_TEST_OE.panel_flag, LAB_TEST_OE.subscript"
''                  #End If
                  '  #End If
                    szQst = ""
                    szQst = szQst & " FROM  " & gszDBName & "LAB_TEST_WTYPE, " & gszDBName & "LAB_TEST_SYNONYM"
''                  #If coiPickList = False Then
''                      szQst = szQst & " WHERE LAB_TEST_WTYPE.lab_test_ien = lab_test_synonym.lab_test_ien"
''                      szQst = szQst & " AND lab_test_synonym.synonym_name like '" & szWildcard & szSearchText & szWildcard & "'" ' string match"
''                      szQst = szQst & " AND lab_test_synonym.synonym_name is not null" 'synonym name must not be empty
''                      szQst = szQst & " AND lab_test_synonym.lab_test_ien <> 0" 'drug IEN must be present
''                  #Else
                    szQst = szQst & " WHERE LAB_TEST_OE.lab_test_ien = lab_test_synonym.lab_test_ien"
                    szQst = szQst & " AND lab_test_synonym.synonym_name like '" & szWildcard & szSearchText & szWildcard & "'" ' string match
                    szQst = szQst & " AND lab_test_synonym.synonym_name is not null" 'synonym name must not be empty
                    szQst = szQst & " AND lab_test_synonym.lab_test_ien <> 0" 'drug IEN must be present
                    szQst = szQst & " AND LAB_TEST_OE.Inactive_Flag = 0"
                    szQst = szQst & " AND LAB_TEST_OE.QC_Only = 0 "
                    szQst = szQst & " AND LAB_TEST_OE.test_type_code > 3"
                    szQst = szQst & szSectionSQL
''                  #End If
                    szOrderBy = " ORDER BY 1;"
                    szSQL = szSelect & szQst & szOrderBy
                    
                Case srcboth:
''                  #If coiPickList = False Then
''                     szSelect = "SELECT Distinct LAB_TEST_WTYPE.lab_test_name, LAB_TEST_WTYPE.lab_test_ien, LAB_TEST_WTYPE.panel_flag, LAB_TEST_WTYPE.subscript"
''                  #Else
                    szSelect = "SELECT Distinct LAB_TEST_OE.lab_test_name, LAB_TEST_OE.lab_test_ien, LAB_TEST_OE.panel_flag, LAB_TEST_OE.subscript, NULL as MinSyn"
                    iNumberofFields = 5
''                  #End If
                    szQst = ""
                        
''                  #If coiPickList = False Then
''                      szQst = szQst & " FROM " & gszDBName & "LAB_TEST_WTYPE"
''                      szQst = szQst & " WHERE LAB_TEST_WTYPE.lab_test_name like '" & szWildcard & szSearchText & szWildcard & "'"
''                      szQst = szQst & " AND LAB_TEST_WTYPE.lab_test_name is not null"
''                      szQst = szQst & " AND LAB_TEST_WTYPE.lab_test_ien <> 0"
''                  #Else
                    szQst = szQst & " FROM " & gszDBName & "LAB_TEST_OE"
                    szQst = szQst & " WHERE LAB_TEST_OE.lab_test_name like '" & szWildcard & szSearchText & szWildcard & "'"
                    szQst = szQst & " AND LAB_TEST_OE.lab_test_name is not null"
                    szQst = szQst & " AND LAB_TEST_OE.lab_test_ien <> 0"
                    szQst = szQst & " AND LAB_TEST_OE.Inactive_Flag = 0"
                    szQst = szQst & " AND LAB_TEST_OE.QC_Only = 0 "
                    szQst = szQst & " AND LAB_TEST_OE.test_type_code > 3"
''                  #End If
                    szQst = szQst & szSectionSQL
                    szQst = szQst & " UNION"
                    
''                  #If coiPickList = False Then
''                      szQst = szQst & " SELECT Distinct LAB_TEST_WTYPE.lab_test_name, LAB_TEST_WTYPE.lab_test_ien, LAB_TEST_WTYPE.panel_flag, LAB_TEST_WTYPE.subscript"
''                      szQst = szQst & " FROM " & gszDBName & "LAB_TEST_WTYPE, " & gszDBName & "LAB_TEST_SYNONYM"
''                      szQst = szQst & " WHERE LAB_TEST_WTYPE.lab_test_ien = lab_test_synonym.lab_test_ien"
''                      szQst = szQst & " AND lab_test_synonym.synonym_name like '" & szWildcard & szSearchText & szWildcard & "'"
''                  #Else
                    szQst = szQst & " SELECT Distinct LAB_TEST_OE.lab_test_name, LAB_TEST_OE.lab_test_ien, LAB_TEST_OE.panel_flag, LAB_TEST_OE.subscript, MIN(LAB_TEST_SYNONYM.synonym_name)"
                    szQst = szQst & " FROM " & gszDBName & "LAB_TEST_OE, " & gszDBName & "LAB_TEST_SYNONYM"
                    szQst = szQst & " WHERE LAB_TEST_OE.lab_test_ien = lab_test_synonym.lab_test_ien"
                    szQst = szQst & " AND lab_test_synonym.synonym_name like '" & szWildcard & szSearchText & szWildcard & "'"
                    szQst = szQst & " AND LAB_TEST_OE.Inactive_Flag = 0"
                    szQst = szQst & " AND LAB_TEST_OE.QC_Only = 0 "
                    szQst = szQst & " AND LAB_TEST_OE.test_type_code > 3"
''                  #End If
                    
                    szQst = szQst & szSectionSQL
''                  #If coiPickList = True Then
                    szQst = szQst & " GROUP BY LAB_TEST_OE.lab_test_name, LAB_TEST_OE.lab_test_ien, LAB_TEST_OE.panel_flag, LAB_TEST_OE.subscript"
''                  #End If
                    szOrderBy = " Order By 1;"
                    szSQL = szSelect & szQst & szOrderBy
                    
            End Select 'Simple, Synonym, Both
            
    End Select
''  #If coiPickList = True Then
        If OrderType = hlRX_ORDERS_CODE Then
            Search = TransferToRxListBox(szSQL, plistbox, Val(szFormularyGroupIEN))
        Else
            Search = TransferToRadLabListBox(szSQL, plistbox, OrderType)
        End If
''  #Else
''      Search = TransferToListBox(szSQL, plistbox, iNumberofFields)
''  #End If
    
               ''5-15-01; AAB #12754
               'Screen.ActiveForm.MousePointer = vbDefault          'display default cursor
               'Removed 6-15-01 AAB #12754 and #12790

    Exit Function

ErrorHandler:
    DisplayLogErr "OE_SearchEngine: Search"
   
               ''5-15-01; AAB #12754
               'Screen.ActiveForm.MousePointer = vbDefault          'display default cursor
               'Removed 6-15-01 AAB #12754 and #12790

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

    Dim szOrderBy As String
    Dim szSelect As String
    Dim szQst As String
    Dim szSQL As String
    Dim gszDBName As String

    On Error GoTo ErrorHandler

    szSelect = "SELECT Distinct rad_test_type.test_type_name"
    szQst = ""
    szQst = szQst & " From " & gszDBName & "RAD_TEST_TYPE"
    szQst = szQst & " Where rad_test_type.test_type_name is not null"
    szQst = szQst & " AND rad_test_type.test_ien <> 0"
    szOrderBy = " ORDER BY rad_test_type.test_type_name;"
    szSQL = szSelect & szQst & szOrderBy
    
    LoadAllRadTestTypes = TransferToListBox(szSQL, plistbox, 1)
    
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
    szOrderBy = " ORDER BY 1;"
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

    Dim szOrderBy As String
    Dim szSelect As String
    Dim szQst As String
    Dim szSQL As String
    Dim gszDBName As String

    On Error GoTo ErrorHandler
    
    szSelect = "SELECT Distinct lab_test_type.test_type_name"
    szQst = ""
    szQst = szQst & " From " & gszDBName & "LAB_TEST_TYPE"
    szQst = szQst & " Where lab_test_type.test_type_name is not null"
    szQst = szQst & " AND lab_test_type.test_type_ien <> 0"
    szOrderBy = " ORDER BY lab_test_type.test_type_name;"
    szSQL = szSelect & szQst & szOrderBy
    
''  #If coiPickList = False Then
''      LoadAllLabTestTypes = TransferToListBox(szSQL, plistbox, 1)
''  #Else
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
        
''  #End If
    
    Exit Function

ErrorHandler:
    DisplayLogErr "OE_SearchEngine: LoadAllLabTestTypes"
    
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' SearchLabTestWType - Seaches the Lab Test WType table
'
' Globals       none
' Call by ref   pListBox        CSComboBox to return SQL quiery results
'               srcSection      Optional Search text for the Lab Data column
' Call by val
'
' Returns       the data saved in the CSComboBox
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function SearchLabTestWType(plistbox As CSComboBox, _
                                    Optional srcSection As String = "") As srcSearchResult

    Dim szOrderBy As String
    Dim szSelect As String
    Dim szQst As String
    Dim szSQL As String
    Dim szSectionSQL As String
    Dim gszDBName As String

    On Error GoTo ErrorHandler
    
    If LenB(srcSection) > 0 Then
        szSectionSQL = " AND lab_test_wtype.test_type_name = '" & srcSection & "'"
    Else
        szSectionSQL = ""
    End If

    szSelect = "SELECT Distinct lab_test_wtype.lab_test_name, lab_test_wtype.lab_test_ien, lab_test_wtype.test_type_name"
    szQst = ""
    szQst = szQst & " From " & gszDBName & "LAB_TEST_WTYPE"
    szQst = szQst & " Where lab_test_wtype.lab_test_name is not null"
    szQst = szQst & " AND lab_test_wtype.lab_test_ien <> 0"
    szQst = szQst & szSectionSQL
    szOrderBy = " ORDER BY lab_test_wtype.lab_test_name;"
    szSQL = szSelect & szQst & szOrderBy
    
    SearchLabTestWType = TransferToListBox(szSQL, plistbox, 2)
    
    Exit Function

ErrorHandler:
    DisplayLogErr "OE_SearchEngine: SearchLabTestWType"
    
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' LoadAllPharmacy - Seaches the Pharmacy table
'
' Globals       none
' Call by ref   pListBox        CSComboBox to return SQL quiery results
'               szDrugIEN       added 10-30-01 for picklist functionality
' Call by val
'
' Returns       the data saved in the CSComboBox, srcErrUnkown if error and
'               srcErrNone if no errors
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function LoadAllPharmacy(plistbox As CSComboBox, Optional szDrugIEN As String = "", Optional szDefPharmacyIEN As String = "") As srcSearchResult

    Dim szOrderBy As String
    Dim szSelect As String
    Dim szQst As String
    Dim szSQL As String
    Dim gszDBName As String
    Dim oRecordSet As Recordset

    Dim pszRxName As String
    Dim pszFormularyGroupIEN As String
    'Dim bNotStocked As Boolean
    Dim pszRxIen As String
    Dim szTempB As String
    Dim szTempA As String
    Dim sWidth As Single
    Dim iIndexDefault As Integer
    
    Const cNOTSTOCKED = "*** DRUG NOT STOCKED ***"

    On Error GoTo ErrorHandler
    
    'if szDrugIEN is blank is only for backward compatibility.  It should do query against a drug ien
    
'#If coiPickList = True Then
 '#Else
   iIndexDefault = -1
    
   If szDrugIEN <> "" Then
        #If coiAccess = True Then
        'this part of query returns pharmacies that the drug is NOT IN stock.  2nd Last item in select is 1, Formulary Message is Null
   
            szQst = "SELECT pharmacy.pharmacy_name, pharmacy.pharmacy_ien, pharmacy.FORMULARY_GROUP_IEN, 1, NULL"
            szQst = szQst & " FROM pharmacy "
            szQst = szQst & " WHERE pharmacy.inactive = 0 AND pharmacy_ien NOT IN "
            szQst = szQst & " (SELECT  pharmacy.pharmacy_ien"
            szQst = szQst & " FROM pharmacy left JOIN RX_FORMULARY ON pharmacy.FORMULARY_GROUP_IEN = RX_FORMULARY.FORMULARY_GROUP_IEN"
            szQst = szQst & " WHERE RX_FORMULARY.DRUG_IEN = " & szDrugIEN & ")"
        
        'this part of query returns pharmacies that the drug is in stock.  2nd Last item in select is 0
            szQst = szQst & " Union "
            szQst = szQst & " SELECT pharmacy.pharmacy_name, pharmacy.pharmacy_ien,  pharmacy.FORMULARY_GROUP_IEN, 0,  RX_FORMULARY.MESSAGE"
            szQst = szQst & " FROM pharmacy left JOIN RX_FORMULARY ON pharmacy.FORMULARY_GROUP_IEN = RX_FORMULARY.FORMULARY_GROUP_IEN"
            szQst = szQst & " WHERE RX_FORMULARY.DRUG_IEN =" & szDrugIEN
            szQst = szQst & " AND pharmacy.inactive = 0 "
            szSQL = szQst & " ORDER By 1"
        #Else
         'this part of query returns pharmacies that the drug is NOT IN stock.  2nd Last item in select is 1, Formulary Message is Null
   
            szQst = "SELECT pharmacy.pharmacy_name,pharmacy.pharmacy_ien,pharmacy.FORMULARY_GROUP_IEN,1, NULL"
            szQst = szQst & " FROM pharmacy, RX_FORMULARY"
            szQst = szQst & " WHERE pharmacy.inactive = 0 AND"
            szQst = szQst & " pharmacy.FORMULARY_GROUP_IEN = RX_FORMULARY.FORMULARY_GROUP_IEN (+) AND"
            szQst = szQst & " RX_FORMULARY.FORMULARY_GROUP_IEN is NULL AND"
            szQst = szQst & " RX_FORMULARY.DRUG_IEN (+) =" & szDrugIEN
            
        'this part of query returns pharmacies that the drug is in stock.  2nd Last item in select is 0
            szQst = szQst & " UNION "
            szQst = szQst & " SELECT pharmacy.pharmacy_name,pharmacy.pharmacy_ien,pharmacy.FORMULARY_GROUP_IEN,0,RX_FORMULARY.Message "
            szQst = szQst & " FROM pharmacy, RX_FORMULARY"
            szQst = szQst & " WHERE pharmacy.FORMULARY_GROUP_IEN (+) = RX_FORMULARY.FORMULARY_GROUP_IEN "
            szQst = szQst & " AND RX_FORMULARY.DRUG_IEN =" & szDrugIEN
            szQst = szQst & " AND pharmacy.inactive = 0 "
            szSQL = szQst & " ORDER By 1;"

        #End If
    
    Else
            'include Formulary; 11-11-01;AAB
            szSelect = "SELECT pharmacy.pharmacy_name, pharmacy.pharmacy_ien, pharmacy.FORMULARY_GROUP_IEN"
            szQst = ""
            szQst = szQst & " From " & gszDBName & "PHARMACY"
            szQst = szQst & " Where pharmacy.pharmacy_name is not null"   'P1270 5-14-96
            szQst = szQst & " AND pharmacy.pharmacy_ien <> 0"       'new for CHCS-II query
            szOrderBy = " ORDER BY pharmacy.pharmacy_name;"
            szSQL = szSelect & szQst & szOrderBy

    End If

    plistbox.Clear                                  'clear listbox
    Set oRecordSet = oDAL.OpenRecordset(szSQL)

    If szDrugIEN = "" Then
        While Not oRecordSet.EOF                        'add each name and ien
                
               pszRxName = Trim$(oRecordSet.Fields("PHARMACY_NAME") & " ")
               pszRxIen = Trim$(oRecordSet.Fields("PHARMACY_IEN") & " ")
               pszFormularyGroupIEN = Trim$(oRecordSet.Fields("FORMULARY_GROUP_IEN") & " ")
    
                   'CBO format is : NAME ^ IEN\NAME\99PSI ^ IEN ^ FORM GROUP IEN ^ NAME ^ DEF FLAG
               szTempB = pszRxIen & "\" & pszRxName & "\99PSI"
               'include Formulary; 11-11-01;AAB
               If pszRxIen = szDefPharmacyIEN Then
                 szTempA = pszRxName & mcINDENT & "(Default) ^" & szTempB & "^" & pszRxIen & "^" & pszFormularyGroupIEN & "^" & pszRxName & "^1"
                 iIndexDefault = plistbox.ListCount
               Else
                 szTempA = pszRxName & "^" & szTempB & "^" & pszRxIen & "^" & pszFormularyGroupIEN & "^" & pszRxName & "^1"
    
               End If
               plistbox.AddItem szTempA
               oRecordSet.MoveNext                         'get the next record
        Wend
        
    Else
            sWidth = 0
            While Not oRecordSet.EOF                        'add each name and ien
                 pszRxName = Trim$(oRecordSet.Fields("PHARMACY_NAME"))
                 pszRxIen = Trim$(oRecordSet.Fields("PHARMACY_IEN"))
                 pszFormularyGroupIEN = Trim$(oRecordSet.Fields("FORMULARY_GROUP_IEN") & " ")
                 
                        'CBO format is : NAME ^ IEN\NAME\99PSI ^ IEN ^ FORM GROUP IEN ^ NAME ^ DEF FLAG
                 szTempB = pszRxIen & "\" & pszRxName & "\99PSI"
                 If pszRxIen = szDefPharmacyIEN Then
                    szTempA = pszRxName & mcINDENT & "(Default) ^" & szTempB & "^" & pszRxIen & "^" & pszFormularyGroupIEN & "^" & pszRxName & "^1"
                    iIndexDefault = plistbox.ListCount
   
                 Else
                    szTempA = pszRxName & "^" & szTempB & "^" & pszRxIen & "^" & pszFormularyGroupIEN & "^" & pszRxName & "^0"
                 
                 End If
                 
                        'Cbo_RxPharmacy.AddItem szTempA
                 plistbox.AddItem szTempA
                 If Trim$(oRecordSet.Fields(3) & " ") = "1" Then
                     szTempA = mcINDENT & cNOTSTOCKED & "^" & szTempB & "^" & pszRxIen & "^" & pszFormularyGroupIEN & "^" & pszRxName
                     plistbox.AddItem szTempA
                 End If
                 If Trim$(oRecordSet.Fields(4) & " ") <> "" Then
                     szTempA = mcINDENT & Trim$(oRecordSet.Fields(4) & " ") & "^" & szTempB & "^" & pszRxIen & "^" & pszFormularyGroupIEN & "^" & pszRxName
                     plistbox.AddItem szTempA
                 End If
                 oRecordSet.MoveNext 'get the next record
                 
                 'plistbox.ListIndex = plistbox.ListCount - 1
                 plistbox.ValueCol = 0
                 szTempB = plistbox.Value
                 
                 If plistbox.Parent.TextWidth(szTempB) > sWidth Then
                    sWidth = plistbox.Parent.TextWidth(szTempB & "     ")
                 End If
            Wend

    End If
          If sWidth > plistbox.Width Then
            plistbox.ListBoxWidth = sWidth / Screen.TwipsPerPixelX
          Else
            plistbox.ListBoxWidth = plistbox.Width / Screen.TwipsPerPixelX
          End If
          'plistbox.ListBoxHeight = (plistbox.Parent.TextHeight("ABC") * plistbox.ListCount) / Screen.TwipsPerPixelY
          
          plistbox.ListIndex = -1
          If plistbox.ListCount > 0 Then                      'check for empty ListBox
              LoadAllPharmacy = srcErrNone
          End If

    '===================================================
    Exit Function

ErrorHandler:
    DisplayLogErr "OE_SearchEngine: LoadAllPharmacy"
    

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
' SearchRequestingLocation - Seaches the Meprs table
'
' Globals       none
' Call by ref   pColl        Collection to hold results
' Call by val
'
' Returns       the data saved in the list box
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function SearchRequestingLocation(plistbox As ListView, szSearchText As String, Optional srcSection As Variant = "") As srcSearchResult
 
    Dim szOrderBy As String
    Dim szSelect As String
    Dim szQst As String
    Dim szSQL As String
  '  Dim gszDBName As String
    Dim oRecordSet As Recordset
    Dim szClinicIEN As String
    Dim szClinicName As String
    Dim szMEPRSIEN As String
    Dim szMEPRSCode As String
    Dim szPharmIEN As String
    Dim szFormGroupIEN As String
    
    On Error GoTo ErrorHandler
    If (LenB(szSearchText) < 4) And (LenB(srcSection) = 0) Then
        MsgBxARMd "Keyword must be at least two characters long.", vbCritical, "Search Criteria Error"
        SearchRequestingLocation = srcKeywordTooShort
        GoTo ErrorHandler
    End If
    
''  #If coiPickList = True Then
    szSelect = "SELECT  Clinic.NAME, Clinic.IEN, Clinic.MEPRS_CODE, Clinic.MEPRS_IEN, Clinic.DEF_PHARMACY_IEN, pharmacy.FORMULARY_GROUP_IEN, clinic.location_type "
    szQst = ""
    #If coiAccess = True Then
        szQst = szQst & " FROM Clinic LEFT JOIN pharmacy ON Clinic.DEF_PHARMACY_IEN = pharmacy.PHARMACY_IEN"
        szQst = szQst & " WHERE Clinic.INACTIVE = '0'"
        szQst = szQst & " AND Clinic.NAME like '" & szWildcard & szSearchText & szWildcard & "'" ' string match
        If LenB(srcSection) > 0 Then
            szQst = szQst & " AND Clinic.location_type = '" & srcSection & "'"  ' string match
        End If
                   
    #Else
        szQst = szQst & " FROM Clinic, pharmacy"
        szQst = szQst & " WHERE Clinic.DEF_PHARMACY_IEN (+) = pharmacy.PHARMACY_IEN"
        szQst = szQst & " AND Clinic.INACTIVE = '0'"
        szQst = szQst & " AND Clinic.NAME like '" & szWildcard & szSearchText & szWildcard & "'" ' string match
        If LenB(srcSection) > 0 Then
            szQst = szQst & " AND Clinic.location_type = '" & srcSection & "'"  ' string match
        End If
 
    #End If
 
    szQst = szQst & " AND MEPRS_CODE is not null"
    szQst = szQst & " AND MEPRS_IEN <> 0"
    szQst = szQst & " AND Clinic.NAME is not Null"
    szOrderBy = " ORDER BY Clinic.NAME"
''  #Else
''          szSelect = "SELECT Distinct meprs.hospital_loc_name, meprs.hospital_loc_ien, meprs.meprs_code, meprs.meprs_ien "
''          szQst = ""
''          szQst = szQst & " From " & gszDBName & "MEPRS"
''          szQst = szQst & " Where meprs.hospital_loc_name is not null"
''          szQst = szQst & " AND meprs.meprs_code is not null"
''          szQst = szQst & " AND meprs.meprs_ien <> 0"
''          szOrderBy = " ORDER BY meprs.hospital_loc_name;"
''      #End If
    
    szSQL = szSelect & szQst & szOrderBy
    Set oRecordSet = oDAL.OpenRecordset(szSQL)
    
    Dim itmX As ListItem
    plistbox.ListItems.Clear
 
 
    While Not oRecordSet.EOF                        'add each name and ien
        szClinicName = Trim$(oRecordSet.Fields(0) & " ")
        szMEPRSCode = Trim$(oRecordSet.Fields(2) & " ")
        szMEPRSIEN = Trim$(oRecordSet.Fields(3) & " ")
        szClinicIEN = Trim$(oRecordSet.Fields(1) & " ")
 
        szPharmIEN = Trim$(oRecordSet.Fields(4) & " ")
        szFormGroupIEN = Trim$(oRecordSet.Fields(5) & " ")
 
        Set itmX = plistbox.ListItems.Add(, , szClinicName)
        itmX.SubItems(1) = Trim$(oRecordSet.Fields(6) & " ")
        'itmX.ListSubItems(colLastFill_RX).Tag = Format$(goOESession.SummaryCollection(i).LastFillDate, "yyyymmddhhss")
        itmX.SubItems(2) = szMEPRSCode
        itmX.Tag = szClinicIEN & "^" & szMEPRSIEN & "^" & szPharmIEN & "^" & szFormGroupIEN & "^"
        
        itmX.Selected = False
      '  plistbox.Add szClinicName & "^" & szClinicIEN & "^" & szMEPRSCode & "^" & szMEPRSIEN & "^" & szClinicName & "^" & szPharmIEN & "^" & szFormGroupIEN           'add text list to list box
        oRecordSet.MoveNext                         'get the next record
    Wend
    SearchRequestingLocation = srcErrNone
    Exit Function
 
ErrorHandler:
    DisplayLogErr "OE_SearchEngine: SearchRequestingLocation"
    SearchRequestingLocation = srcErrUnknown
    
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
'KELLY DEAD 7-27-01     Dim szSearchTextSQL As String
    Dim gszDBName As String

    On Error GoTo ErrorHandler
    
   ''        pszAddItem = pszName & "^" & pZoc2_TopographyName & "^" & pszCap & "^" & pszIEN & "^" & pZoc2_szTopoIen & "^" & szRslt_Comment
''        Cbo_ZocList.AddItem pszAddItem

    ''szSelect = "SELECT Distinct lab_specimen.specimen_name, lab_specimen.specimen_data, lab_specimen.specimen_ien"
      'DISTINCT will NOT send us all that we want;
               szSelect = "SELECT  lab_specimen.specimen_name, lab_specimen.specimen_data, lab_specimen.specimen_ien"
 
               szQst = ""
               szQst = szQst & " From " & gszDBName & "LAB_SPECIMEN"
               szQst = szQst & " Where lab_specimen.specimen_name is not null"
               szQst = szQst & " AND lab_specimen.specimen_ien <> 0"
               szOrderBy = " ORDER BY lab_specimen.specimen_name;"
               szSQL = szSelect & szQst & szOrderBy
    
    ''SearchSpecimen = TransferToListBox(szSQL, plistbox, 3)
    
Dim oRecordSet As Recordset
Dim iFieldIndex As Integer
Dim sListBoxLine As String
Dim sDelimiter As String
'KELLY DEAD 7-27-01 Dim sControlName As String
'KELLY DEAD 7-27-01 Dim iFieldMaxCount As Integer

 ''   On Error GoTo Errorhandler
    
 'KELLY DEAD 7-27-01    iFieldMaxCount = 3
            sDelimiter = "^"  'vbTab

          'Screen.ActiveForm.MousePointer = vbHourglass        'display hour glass cursor
          'Removed 6-15-01 AAB #12754 and #12790

    plistbox.Clear                                      'clear listbox
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
    
    Exit Function

ErrorHandler:
    DisplayLogErr "OE_SearchEngine: SearchSpecimen"
    
          'Screen.ActiveForm.MousePointer = vbDefault          'display default cursor
          'Removed 6-15-01 AAB #12754 and #12790

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
                            
'KELLY DEAD 7-27-01     Dim szOrderBy As String
'KELLY DEAD 7-27-01     Dim szSelect As String
    Dim szQst As String
'KELLY DEAD 7-27-01     Dim szSQL As String
    Dim gszDBName As String

    On Error GoTo ErrorHandler:
  
''    szSelect = "Select distinct lab_test.lab_test_name, lab_panel.lab_test_ien, lab_panel.lab_test_ien"
''    szQst = ""
''    szQst = szQst & " From " & gszDBName & "LAB_TEST, " & gszDBName & "LAB_PANEL"
''    szQst = szQst & " Where lab_test.lab_test_ien = lab_panel.lab_test_ien"
''    szQst = szQst & " AND lab_panel.lab_panel_ien = " & SampleIEN
''    szQst = szQst & " ORDER BY lab_test.lab_test_name;"
''    szSQL = szSelect & szQst & szOrderBy
''
''    SearchLabPanel = TransferToListBox(szSQL, plistbox, 3)

'fixed 9-20-99; AAB

    szQst = "Select LAB_TEST_NAME  "
''  #If coiPickList = False Then
''      szQst = szQst & " FROM " & gszDBName & "LAB_TEST, " & gszDBName & "LAB_PANEL "
''      szQst = szQst & "Where LAB_TEST.LAB_TEST_IEN = LAB_PANEL.LAB_TEST_IEN "
''      szQst = szQst & " AND LAB_PANEL_IEN = " & SampleIEN
''      szQst = szQst & " order by LAB_TEST_NAME"

''  #Else
        szQst = szQst & " FROM " & gszDBName & "LAB_TEST_OE, " & gszDBName & "LAB_PANEL "
        szQst = szQst & "Where LAB_TEST_OE.LAB_TEST_IEN = LAB_PANEL.LAB_TEST_IEN "
        szQst = szQst & " AND LAB_PANEL_IEN = " & SampleIEN
        szQst = szQst & " order by LAB_TEST_NAME"

''  #End If
    
    SearchLabPanel = TransferToListBox(szQst, plistbox, 1)

    Exit Function
  
ErrorHandler:
    DisplayLogErr "OE_SearchEngine: SearchLabPanel"

End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' SearchFormularyGroupIEN - Seaches the clinic and pharmacy table
'
' Globals       none
' Call by ref   szClinicIEN         The clinic we are searching for its formulary group
'               szFormGroupIEN       The answer
' Call by val
'
' Returns       Result
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function SearchFormularyGroupIEN(szClinicIEN As String, szFormGroupIEN As String, szDefPharmacyIEN As String, szPharmacyName As String) As srcSearchResult
                            
    Dim szQst As String
    Dim gszDBName As String
    Dim oRecordSet As Recordset
    
    On Error GoTo ErrorHandler:
    SearchFormularyGroupIEN = srcErrUnknown
    szQst = "Select PHARMACY.Formulary_group_ien, CLINIC.def_pharmacy_ien, PHARMACY.pharmacy_name  "
    
    szQst = szQst & " FROM " & gszDBName & "CLINIC, " & gszDBName & "PHARMACY "
    szQst = szQst & " WHERE CLINIC.def_pharmacy_ien = PHARMACY.pharmacy_ien "
    szQst = szQst & " AND CLINIC.ien = '" & szClinicIEN & "'"

    Set oRecordSet = oDAL.OpenRecordset(szQst)
    If Not oRecordSet.EOF Then
        szFormGroupIEN = Trim$(oRecordSet.Fields(0) & " ")
        szDefPharmacyIEN = Trim$(oRecordSet.Fields(1) & " ")
        szPharmacyName = Trim$(oRecordSet.Fields(2) & " ")
    Else
        szFormGroupIEN = ""
        szDefPharmacyIEN = ""
        szPharmacyName = ""
    End If
    
    SearchFormularyGroupIEN = srcErrNone

    Exit Function
  
ErrorHandler:
    DisplayLogErr "OE_SearchEngine: SearchFormularyGroupIEN"

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
'Friend Function LoadAllMeprs(plistbox As CSComboBox) As srcSearchResult
'
'    Dim szOrderBy As String
'    Dim szSelect As String
'    Dim szQst As String
'    Dim szSQL As String
'    Dim gszDBName As String
'
'    On Error GoTo ErrorHandler
'
'    szSelect = "SELECT Distinct meprs.hospital_loc_name, meprs.meprs_code, meprs.meprs_ien, meprs.hospital_loc_ien"
'    szQst = ""
'    szQst = szQst & " From " & gszDBName & "MEPRS"
'    szQst = szQst & " Where meprs.hospital_loc_name is not null"
'    szQst = szQst & " AND meprs.meprs_code is not null"
'    szQst = szQst & " AND meprs.meprs_ien <> 0"
'    szOrderBy = " ORDER BY meprs.hospital_loc_name;"
'    szSQL = szSelect & szQst & szOrderBy
'
'    LoadAllMeprs = TransferToListBox(szSQL, plistbox, 4)
'    Exit Function
'
'ErrorHandler:
'    DisplayLogErr "OE_SearchEngine: LoadAllMeprs"
'
'End Function
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
  '  Dim gszDBName As String
    Dim oRecordSet As Recordset
    Dim szClinicIEN As String
    Dim szClinicName As String
    Dim szMEPRSIEN As String
    Dim szMEPRSCode As String
    Dim szPharmIEN As String
    Dim szFormGroupIEN As String
    
    On Error GoTo ErrorHandler
''  #If coiPickList = True Then
    szSelect = "SELECT  Clinic.NAME, Clinic.IEN, Clinic.MEPRS_CODE, Clinic.MEPRS_IEN, Clinic.DEF_PHARMACY_IEN, pharmacy.FORMULARY_GROUP_IEN "
    szQst = ""
    #If coiAccess = True Then
        szQst = szQst & " FROM Clinic LEFT JOIN pharmacy ON Clinic.DEF_PHARMACY_IEN = pharmacy.PHARMACY_IEN"
        szQst = szQst & " WHERE Clinic.INACTIVE = '0'"
    #Else
        szQst = szQst & " FROM Clinic, pharmacy"
        szQst = szQst & " WHERE Clinic.DEF_PHARMACY_IEN (+) = pharmacy.PHARMACY_IEN"
        szQst = szQst & " AND Clinic.INACTIVE = '0'"
    #End If

    szQst = szQst & " AND MEPRS_CODE is not null"
    szQst = szQst & " AND MEPRS_IEN <> 0"
    szQst = szQst & " AND Clinic.NAME is not Null"
    szOrderBy = " ORDER BY Clinic.NAME"
''  #Else
''          szSelect = "SELECT Distinct meprs.hospital_loc_name, meprs.hospital_loc_ien, meprs.meprs_code, meprs.meprs_ien "
''          szQst = ""
''          szQst = szQst & " From " & gszDBName & "MEPRS"
''          szQst = szQst & " Where meprs.hospital_loc_name is not null"
''          szQst = szQst & " AND meprs.meprs_code is not null"
''          szQst = szQst & " AND meprs.meprs_ien <> 0"
''          szOrderBy = " ORDER BY meprs.hospital_loc_name;"
''      #End If
    
    szSQL = szSelect & szQst & szOrderBy
    Set oRecordSet = oDAL.OpenRecordset(szSQL)
    
    While Not oRecordSet.EOF                        'add each name and ien
        
        szClinicName = Trim$(oRecordSet.Fields(0) & " ")
        szMEPRSCode = Trim$(oRecordSet.Fields(2) & " ")
        szMEPRSIEN = Trim$(oRecordSet.Fields(3) & " ")
        szClinicIEN = Trim$(oRecordSet.Fields(1) & " ")
''      #If coiPickList = True Then
        szPharmIEN = Trim$(oRecordSet.Fields(4) & " ")
        szFormGroupIEN = Trim$(oRecordSet.Fields(5) & " ")
''      #Else
''          szPharmIEN = ""
''          szFormGroupIEN = ""
''      #End If
            
        pColl.Add szClinicIEN, szMEPRSCode, szMEPRSIEN, szClinicName, szPharmIEN, szFormGroupIEN      'add text list to list box
        oRecordSet.MoveNext                         'get the next record
    Wend
    LoadAllMeprsIntoCollection = srcErrNone
    Exit Function

ErrorHandler:
    DisplayLogErr "OE_SearchEngine: LoadAllMeprsIntoCollection"
    LoadAllMeprsIntoCollection = srcErrUnknown
    
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
Friend Function SearchClinicTableforNCID() As srcSearchResult

    Dim szOrderBy As String
    Dim szSelect As String
    Dim szQst As String
    Dim szSQL As String
  
    Dim oRecordSet As Recordset
    Dim szClinicIEN As String
    Dim szClinicName As String
    Dim szMEPRSIEN As String
    Dim szMEPRSCode As String
'    Dim szPharmIEN As String
'    Dim szFormGroupIEN As String
    Dim szTempA As String
    Dim szNCID As String
    
    On Error GoTo ErrorHandler

    If Not (goAnP Is Nothing) Then
    
    
        szNCID = goAnP.EncounterClinicNCID
        
        szSelect = "SELECT  Clinic.NAME, Clinic.IEN, Clinic.MEPRS_CODE, Clinic.MEPRS_IEN, Clinic.DEF_PHARMACY_IEN "
        szQst = ""
'        #If coiAccess = True Then
'            szQst = szQst & " FROM Clinic "
'            szQst = szQst & " WHERE Clinic.INACTIVE = '0'"
'        #Else
        szQst = szQst & " FROM Clinic"
        szQst = szQst & " WHERE Clinic.INACTIVE = '0'"
'        #End If
    
        szQst = szQst & " AND Clinic.NCID = " & szNCID & ""
        szQst = szQst & " AND MEPRS_CODE is not null"
        szQst = szQst & " AND MEPRS_IEN <> 0"
        szQst = szQst & " AND Clinic.NAME is not Null"
    
        szSQL = szSelect & szQst & szOrderBy
        Set oRecordSet = oDAL.OpenRecordset(szSQL)
        
        If Not oRecordSet.EOF Then
            goOESession.ClinicName = Trim$(oRecordSet.Fields(0) & " ")
            goOESession.ClinicIEN = Trim$(oRecordSet.Fields(1) & " ")
            goOESession.MEPRSCode = Trim$(oRecordSet.Fields(2) & " ")
            goOESession.MEPRSIen = Trim$(oRecordSet.Fields(3) & " ")
        End If
     End If
'        If plistbox.ListCount > 1 Then
'                SetRequestingLocationinComboBox plistbox, False, False
'        Else
'                plistbox.Clear
'                szTempA = goOESession.ClinicName & "^" & _
'                    goOESession.MEPRSCode & "^" & _
'                    goOESession.MEPRSIen & "^" & _
'                    goOESession.ClinicIEN
'
'                    plistbox.AddItem szTempA
'                    plistbox.ListIndex = 0
'        End If
   
'    While Not oRecordSet.EOF                        'add each name and ien
'
'
'        szPharmIEN = Trim$(oRecordSet.Fields(4) & " ")
'        szFormGroupIEN = Trim$(oRecordSet.Fields(5) & " ")
'
'        pColl.Add szClinicIEN, szMEPRSCode, szMEPRSIEN, szClinicName, szPharmIEN, szFormGroupIEN      'add text list to list box
'        oRecordSet.MoveNext                         'get the next record
'    Wend
    SearchClinicTableforNCID = srcErrNone
    Exit Function

ErrorHandler:
    DisplayLogErr "OE_SearchEngine: SearchClinicTableforNCID"
    SearchClinicTableforNCID = srcErrUnknown
    
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' TransferToListBox - Does SQL quiery and transfers record set
'                     to a List Box
'
' Globals       none
' Call by ref   szSQL           SQL quiery statements
'               pListBox        Cresent/Combo Box ,VB ComboBox, or VB ListBox to
'                               return SQL quiery results
'               iFieldMaxCount  The number of fields from SQL search
'                               that will be put into each line of the control.
' Call by val
'
' Returns       the data saved in the list box, srcErrUnkown if error and
'               srcErrNone if no errors
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function TransferToListBox(szSQL As String, _
                                    plistbox As Control, _
                                    iFieldMaxCount As Integer) As srcSearchResult

    Dim oRecordSet As Recordset
    Dim iFieldIndex As Integer
    Dim sListBoxLine As String
    Dim sDelimiter As String
    Dim sControlName As String
    Dim sWidth As Single

   ' On Error GoTo ErrorHandler
    
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
'            ** PR28 ** Default to vbTab and not a MsgBxARMd
'            MsgBxARMd ("Unsupported ComboBox control: " & sControlName)
'            GoTo Errorhandler

            sDelimiter = vbTab

    End Select

    'Screen.ActiveForm.MousePointer = vbHourglass        'display hour glass cursor
    'Removed 6-15-01 AAB #12754 and #12790
    
    TransferToListBox = srcErrUnknown
    
    plistbox.Clear                                      'clear listbox
    Set oRecordSet = oDAL.OpenRecordset(szSQL)
    
    sWidth = 0
    While Not oRecordSet.EOF                        'add each name and ien
        sListBoxLine = ""
        iFieldIndex = 0
        While iFieldIndex < iFieldMaxCount          'Build the text list for the list box
            If iFieldIndex = 0 Then
                If plistbox.Parent.TextWidth(Trim$(oRecordSet.Fields(iFieldIndex))) > sWidth Then
                    sWidth = plistbox.Parent.TextWidth(oRecordSet.Fields(iFieldIndex) & "       ")
                End If
            End If
            'sListBoxLine = sListBoxLine
'           sListBoxLine = sListBoxLine & Trim(oRecordSet.GetString(iFieldIndex)) & sDelimiter
            sListBoxLine = sListBoxLine & Trim$(oRecordSet.Fields(iFieldIndex) & " ") & sDelimiter
            iFieldIndex = iFieldIndex + 1
        Wend
        plistbox.AddItem sListBoxLine               'add text list to list box
        oRecordSet.MoveNext                         'get the next record
    Wend
    If sWidth > plistbox.Width Then
        plistbox.ListBoxWidth = sWidth / Screen.TwipsPerPixelX
    Else
        plistbox.ListBoxWidth = plistbox.Width / Screen.TwipsPerPixelX
    End If
          'Screen.ActiveForm.MousePointer = vbDefault          'display default cursor
          'Removed 6-15-01 AAB #12754 and #12790

    If plistbox.ListCount > 0 Then                      'check for empty ListBox
        TransferToListBox = srcErrNone
    End If
    
    Exit Function

ErrorHandler:
     DisplayLogErr "OE_SearchEngine: TransferToListBox"
     
          'added 5-16-01; AAB
          'Screen.ActiveForm.MousePointer = vbDefault          'display default cursor
          'Removed 6-15-01 AAB #12754 and #12790

End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' TransferToRadLabListBox - Does SQL quiery and transfers record set
'                     to a Lab or Rad List Box
'
' Globals       none
' Call by ref   szSQL           SQL quiery statements
'               pListBox        Cresent/Combo Box ,VB ComboBox, or VB ListBox to
'                               return SQL quiery results
'               iFieldMaxCount  The number of fields from SQL search
'                               that will be put into each line of the control.
'               OrderType
' Call by val
'
' Returns       the data saved in the list box, srcErrUnkown if error and
'               srcErrNone if no errors
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function TransferToRadLabListBox(szSQL As String, _
                                    plistbox As Control, _
                                    OrderType As hlORDERTYPECODE) As srcSearchResult

    Dim oRecordSet As Recordset
    'Dim iFieldIndex As Integer
    Dim sListBoxLine As String
    Dim sDelimiter As String
    Dim sControlName As String
    Dim sWidth As Single
    'Dim szDisplayName As String
    'Dim szName As String
    
    Dim pCol As New Collection
    Dim pItem As LabRadItem
    
    'Bubble sort stuff
    Dim arItems() As SearchItem
    Dim lngArray As Long
    Dim lngStep As Long
    Dim i As Integer
    
    On Error GoTo ErrorHandler

   ' If iFieldMaxCount < 1 Then iFieldMaxCount = 1       'check iFieldMaxCount
    
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
    
    TransferToRadLabListBox = srcErrUnknown
    
    plistbox.Clear                                      'clear listbox
    Set oRecordSet = oDAL.OpenRecordset(szSQL)
    
    
    While Not oRecordSet.EOF                        'add each name and ien
        
        pItem.szDisplayName = ""
        pItem.szIEN = Trim$(oRecordSet.Fields(1))
        pItem.szName = Trim$(oRecordSet.Fields(0))
        pItem.szSynonym = Trim$(oRecordSet.Fields("MinSyn") & " ")
        If OrderType = hlLAB_ORDERS_CODE Then
            pItem.szPanelFlag = Trim$(oRecordSet.Fields("panel_flag") & " ")
            pItem.szSubscript = Trim$(oRecordSet.Fields("subscript") & " ")
        End If
        If pCol.Count > 0 Then
            If pItem.szIEN = pCol.Item(pCol.Count).szIEN Then
                If LenB(pItem.szSynonym) > 0 And pItem.szSynonym < pItem.szName And pCol.Item(pCol.Count).szDisplayName <> pItem.szSynonym & mcINDENT & pItem.szName Then
                    pItem.szDisplayName = pItem.szSynonym & mcINDENT & pItem.szName
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
    
    'Bubble Sort
    BubbleSort pCol.Count, arItems  'alphabetize by Display Name
    
    
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



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' TransferToRxListBox - Does SQL quiery and transfers record set
'                     to a the RxList Box
'
' Globals       none
' Call by ref   szSQL           SQL quiery statements
'               pListBox        Cresent/Combo Box ,VB ComboBox, or VB ListBox to
'                               return SQL quiery results
'               iFieldMaxCount  The number of fields from SQL search
'                               that will be put into each line of the control.
' Call by val
'
' Returns       the data saved in the list box, srcErrUnkown if error and
'               srcErrNone if no errors
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function TransferToRxListBox(szSQL As String, _
                                    plistbox As Control, _
                                    currentFormulary As Integer) As srcSearchResult

    Dim oRecordSet As Recordset
    'Dim iFieldIndex As Integer
    Dim sListBoxLine As String
    Dim sDelimiter As String
    Dim sControlName As String
    Dim sWidth As Single
    'Dim szIEN As String
    'Dim szName As String
    'Dim szDefUnit As String
    'Dim szSynonym As String

    Dim pCol As New Collection
    Dim pItem As RxItem
    Dim szSuffix As String
    Dim newDisplayName As String
    Dim i As Integer, j As Integer, currentIEN As String
    Dim iStocked As Integer
    
    'Bubble sort stuff
    Dim arItems() As SearchItem
    Dim outOfFormularyRepeat() As Boolean
   ' Dim lngArray As Long
    Dim lngStep As Long
    On Error GoTo ErrorHandler
    
'    If iFieldMaxCount < 1 Then iFieldMaxCount = 1       'check iFieldMaxCount
    
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
'            ** PR28 ** Default to vbTab and not a MsgBxARMd
'            MsgBxARMd ("Unsupported ComboBox control: " & sControlName)
'            GoTo Errorhandler

            sDelimiter = vbTab

    End Select

    'Screen.ActiveForm.MousePointer = vbHourglass        'display hour glass cursor
    'Removed 6-15-01 AAB #12754 and #12790
    
    TransferToRxListBox = srcErrUnknown
    
    plistbox.Clear                                      'clear listbox
    Set oRecordSet = oDAL.OpenRecordset(szSQL)
    

    While Not oRecordSet.EOF 'add each name and ien
        pItem.szDisplayName = ""
        pItem.szIEN = Trim$(oRecordSet.Fields(1))
        pItem.szName = Trim$(oRecordSet.Fields(0))
        If IsNull(oRecordSet.Fields("ip_op_indicator").Value) Or _
                  oRecordSet.Fields("ip_op_indicator").Value = "" Then 'Tracker 51328
            pItem.ip_op_indicator = 0
        Else
            pItem.ip_op_indicator = oRecordSet.Fields("ip_op_indicator")
        End If
        If IsNull(oRecordSet.Fields("drug_check_disabled").Value) Or _
                  oRecordSet.Fields("drug_check_disabled").Value = "" Then 'Tracker 51328
            pItem.drug_check_disabled = 0
        Else
            pItem.drug_check_disabled = oRecordSet.Fields("drug_check_disabled")
        End If
        If IsNull(oRecordSet.Fields("non_formulary").Value) Or _
                  oRecordSet.Fields("non_formulary").Value = "" Then 'Tracker 51328
            pItem.non_formulary = 0
        Else
            pItem.non_formulary = oRecordSet.Fields("non_formulary")
        End If
        If IsNull(oRecordSet.Fields("formulary_group_ien").Value) Or _
                  oRecordSet.Fields("formulary_group_ien").Value = "" Then 'Tracker 51328
            pItem.formulary_group_ien = 0
        Else
            pItem.formulary_group_ien = oRecordSet.Fields("formulary_group_ien")
        End If
        pItem.Message = Trim$(oRecordSet.Fields("message") & " ")
        pItem.szSynonym = Trim$(oRecordSet.Fields("MinSyn") & " ")
        pItem.szDefUnit = Trim$(oRecordSet.Fields(2) & " ")
        If IsNull(oRecordSet.Fields("Stocked").Value) Or _
                  oRecordSet.Fields("Stocked").Value = "" Then
            pItem.Stocked = 0
        Else
            pItem.Stocked = oRecordSet.Fields("Stocked")
        End If
        
        If pCol.Count > 0 Then
           If pItem.szIEN = pCol.Item(pCol.Count).szIEN And _
              pItem.formulary_group_ien = currentFormulary Then
                newDisplayName = pItem.szSynonym & mcINDENT & pItem.szName & mcINDENT & pItem.szDefUnit
                If LenB(pItem.szSynonym) > 0 And pItem.szSynonym < pItem.szName And _
                   pCol.Item(pCol.Count).szDisplayName <> newDisplayName Then
                    pItem.szDisplayName = newDisplayName
                    pCol.Remove pCol.Count
                    pCol.Add pItem
                    'arItems(lngArray).szDisplayName = pItem.szDisplayName
                Else 'No synonym, but may be 1st within formulary
                    If pCol.Item(pCol.Count).formulary_group_ien <> currentFormulary Then
                        pItem.szDisplayName = pItem.szName & mcINDENT & pItem.szDefUnit
                        pCol.Add pItem
                    End If
                End If
            Else
                If pItem.szIEN <> pCol.Item(pCol.Count).szIEN Then
                'Exclude useless non formulary Items
                    pItem.szDisplayName = pItem.szName & mcINDENT & pItem.szDefUnit
                    pCol.Add pItem
                'lngArray = pCol.Count - 1
                'ReDim Preserve arItems(0 To lngArray)
                'arItems(lngArray).szDisplayName = pItem.szDisplayName
                'arItems(lngArray).iCollectionIndex = pCol.Count
                End If
           End If
        Else
            pItem.szDisplayName = pItem.szName & mcINDENT & pItem.szDefUnit
            pCol.Add pItem
            'lngArray = pCol.Count - 1
            'ReDim Preserve arItems(0 To lngArray)
            'arItems(lngArray).szDisplayName = pItem.szDisplayName
            'arItems(lngArray).iCollectionIndex = pCol.Count
        End If
        
        oRecordSet.MoveNext                         'get the next record
    Wend
    
    ReDim outOfFormularyRepeat(1 To pCol.Count)
    
    'Remove unnecessary items from other formularies
    For i = 1 To pCol.Count
        currentIEN = pCol.Item(i).szIEN
        For j = (i + 1) To pCol.Count
            If pCol.Item(j).szIEN = currentIEN Then
                If pCol.Item(j).formulary_group_ien <> currentFormulary Then
                    outOfFormularyRepeat(j) = True
                Else
                    outOfFormularyRepeat(i) = True
                End If
            End If
        Next j
    Next i
        
    For i = pCol.Count To 1 Step -1
        If outOfFormularyRepeat(i) Then pCol.Remove i
    Next i
       
    ReDim arItems(0 To (pCol.Count - 1))
    For i = 1 To pCol.Count
        arItems(i - 1).szDisplayName = pCol.Item(i).szDisplayName
        arItems(i - 1).iCollectionIndex = i
    Next i
  
    'Bubble Sort
    BubbleSort pCol.Count, arItems
    lngStep = -1
    
    'For i = 1 To pCol.Count
    sWidth = 0
    For iStocked = 1 To 0 Step -1 'first show the stocked items and then show unstocked
        For lngStep = 0 To pCol.Count - 1
            
            i = arItems(lngStep).iCollectionIndex
            If pCol(i).Stocked = iStocked Then
               szSuffix = sDelimiter & pCol(i).szIEN & sDelimiter & pCol(i).szName
               sListBoxLine = pCol(i).szDisplayName & " " & szSuffix
               plistbox.AddItem sListBoxLine
               
               If pCol(i).ip_op_indicator = 1 Then
                   sListBoxLine = mcINDENT & "*** INPATIENT ONLY ***" & szSuffix
                   plistbox.AddItem sListBoxLine
               End If
               If pCol(i).drug_check_disabled = 1 Then
                   sListBoxLine = mcINDENT & "*** DRUG CHECKS DISABLED ***" & szSuffix
                   plistbox.AddItem sListBoxLine
               End If
               
               If pCol(i).non_formulary = 1 Then
                   sListBoxLine = mcINDENT & "*** NON FORMULARY ***" & szSuffix
                   plistbox.AddItem sListBoxLine
               End If
               If pCol(i).formulary_group_ien <> currentFormulary Then
                   sListBoxLine = mcINDENT & "*** NOT STOCKED AT DEFAULT PHARMACY ***" & szSuffix
                   plistbox.AddItem sListBoxLine
               End If
               
               If LenB(pCol(i).Message) > 0 Then
                   sListBoxLine = mcINDENT & pCol(i).Message & szSuffix
                   plistbox.AddItem sListBoxLine
               End If
               
            '   lngStep = lngStep + 1
            '   arItems(lngStep).szDisplayName = pCol(i).szDisplayName
            '   arItems(lngStep).iCollectionIndex = lngStep
               sListBoxLine = szGetColumn(sListBoxLine, 1, sDelimiter) & mcINDENT & mcINDENT
               If plistbox.Parent.TextWidth(sListBoxLine) > sWidth Then
                   sWidth = plistbox.Parent.TextWidth(sListBoxLine & mcINDENT & mcINDENT)
               End If
            End If
        Next
    Next iStocked
  
    If sWidth > plistbox.Width Then
        plistbox.ListBoxWidth = sWidth / Screen.TwipsPerPixelX
    Else
        plistbox.ListBoxWidth = plistbox.Width / Screen.TwipsPerPixelX
    End If
          'Screen.ActiveForm.MousePointer = vbDefault          'display default cursor
          'Removed 6-15-01 AAB #12754 and #12790

    If plistbox.ListCount > 0 Then                      'check for empty ListBox
        TransferToRxListBox = srcErrNone
    End If
    
    Exit Function

ErrorHandler:
      DisplayLogErr "OE_SearchEngine: TransferToRxListBox"
'     If Err.Number = 457 Then 'Already added item
'        Resume Next
'
'     End If
     Debug.Print Err.Description
          'added 5-16-01; AAB
          'Screen.ActiveForm.MousePointer = vbDefault          'display default cursor
          'Removed 6-15-01 AAB #12754 and #12790

End Function


Private Sub Class_Initialize()
    Dim oDatabase As New CHCSII_CONN.Conn

    Set oDAL = oDatabase.CHCSII_DAS
    #If coi2000 = True Then
        szWildcard = "*"
    #Else
        #If coiAccess = True Then
            szWildcard = "%"
        #Else
            szWildcard = "%"
        #End If
    #End If

End Sub

Private Sub Class_Terminate()
 Set oDAL = Nothing
 
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GetSignatureClass - Returns the CHCS I Signature Class based upon Provider IEN
'
' Globals       none
' Call by ref   ProviderIEN       Provider IEN of logged in user
' Call by val
'
' Returns       Integer - Sig Class 0-5 or 99
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GetSignatureClass(ProviderIEN As String, FACILITY_NCID As Long) As Integer

    Dim oRecordSet As Recordset
    'Dim szOrderBy As String
    Dim szSelect As String
    Dim szQst As String
    Dim szSQL As String
    Dim gszDBName As String
  
    On Error GoTo ErrorHandler:
  
    szSelect = "Select SIGNATURE_CLASS"
    szQst = ""
    szQst = szQst & " From " & gszDBName & "PROVIDER"
    szQst = szQst & " Where IEN = '" & ProviderIEN
    szQst = szQst & "' AND FACILITY_NCID = " & FACILITY_NCID
    
    'szSelect = "Select SIGNATURE_CLASS"
    'szQst = ""
    'szQst = szQst & " From " & gszDBName & "CLINIC_PROVIDER"
    'szQst = szQst & " Where CLINIC_PROVIDER.PROVIDER_IEN = " & ProviderIEN
    'szQst = szQst & " AND CLINIC_PROVIDER.FACILITY_NCID = " & FACILITY_NCID
    
    szSQL = szSelect & szQst
    GetSignatureClass = -1
    
    Set oRecordSet = oDAL.OpenRecordset(szSQL)
    
    If Not oRecordSet.EOF Then
        GetSignatureClass = oRecordSet.Fields("SIGNATURE_CLASS")
    End If

    Exit Function
  
ErrorHandler:
    DisplayLogErr "OE_SearchEngine: GetSignatureClass"

End Function

Private Sub BubbleSort(lngCollectionCount As Long, arItems() As SearchItem)
   'Bubble Sort
    Dim lngStep As Long
    Dim strTemp As String
    Dim iTempCollectionIndex As Integer
    Dim lngPassNum As Long
    
    For lngPassNum = 1 To lngCollectionCount - 1
        For lngStep = 0 To lngCollectionCount - lngPassNum - 1
            If arItems(lngStep).szDisplayName > arItems(lngStep + 1).szDisplayName Then
                strTemp = arItems(lngStep).szDisplayName
                iTempCollectionIndex = arItems(lngStep).iCollectionIndex
                arItems(lngStep).szDisplayName = arItems(lngStep + 1).szDisplayName
                arItems(lngStep).iCollectionIndex = arItems(lngStep + 1).iCollectionIndex
                arItems(lngStep + 1).szDisplayName = strTemp
                arItems(lngStep + 1).iCollectionIndex = iTempCollectionIndex
            End If
        Next lngStep
    Next lngPassNum
End Sub
