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
Attribute VB_Name = "colGroups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsGroup"
Attribute VB_Ext_KEY = "Member0" ,"clsGroup"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable to hold collection
Private mCol As Collection
Private mvarCurrentDentalEncounterID As Long

'Collection to hold all Txs to ensure unique data integrity
Private mvarTxsRef As Collection

'Collection to hold all Dxs to ensure unique data integrity
Private mvarDxsRef As Collection

'Disconnected recordset to hold all the existing Dxs to improve loading performance.
Private rsDxsToLoad As Recordset
'Disconnected recordset to hold all the existing Dxs to improve loading performance.
Private rsTxsToLoad As Recordset
'Disconnected recordset to hold the surfaceset and tooth info for the Dxs
Private rsDxSSTeethToLoad As Recordset
'Disconnected recordset to hold the surfaceset and tooth info for the Dxs
Private rsTxSSTeethToLoad As Recordset
'Disconnected recordset to hold the surfaceset and tooth info for the Dxs
Private rsDxLocSubLocToLoad As Recordset
'Disconnected recordset to hold the surfaceset and tooth info for the Dxs
Private rsTxLocSubLocToLoad As Recordset

Friend Property Let CurrentDentalEncounterID(vData As Long)
    
    mvarCurrentDentalEncounterID = vData
    
    Dim oGroup As clsGroup
    Dim oDx As clsDx
    Dim oTx As clsTx
    
    For Each oGroup In mCol
        oGroup.Dxs.CurrentDentalEncounterID = mvarCurrentDentalEncounterID
        oGroup.Txs.CurrentDentalEncounterID = mvarCurrentDentalEncounterID
    Next
        
End Property

Private Sub LoadSSTeethandLocSubLoc(oDxs As colDxs, oTxs As colTxs)

'This sub loads the surfacesets, teeth, locations, sublocations for Dxs and Txs

    If Not oDxs Is Nothing Then
        'Dxs part
        'Load the SS and Teeth for the dxs in this group
        oDxs.RefreshSSTeethFromRS rsDxSSTeethToLoad
        'Load the Loc and Sub loc for the dxs in this group
        oDxs.RefreshLocSubLocFromRS rsDxLocSubLocToLoad

    End If
    
    If Not oTxs Is Nothing Then
        'Txs part
        'Load the SS and Teeth for the txs in this group
        oTxs.RefreshSSandTeethFromRS rsTxSSTeethToLoad
        'Load the Loc and SubLoc for the txs in this group
        oTxs.RefreshLocSubLocFromRS rsTxLocSubLocToLoad
        
    End If
    
End Sub

Public Function SaveGroups() As Boolean

Dim oGroup As clsGroup

    SaveGroups = True
    
    For Each oGroup In mCol
        If oGroup.SaveGroup = False Then
           'If one dx doesn't save properly, then set SaveGroups = false
            SaveGroups = False
        End If
    Next
    
    SaveGroups = True
    
End Function


Friend Property Get TxsRef() As Collection

    Set TxsRef = mvarTxsRef

End Property

Friend Property Get DxsRef() As Collection

    Set DxsRef = mvarDxsRef

End Property
Friend Property Let TxsRef(vData As Collection)

    Set mvarTxsRef = vData
    
End Property

Friend Property Let DxsRef(vData As Collection)

    Set mvarDxsRef = vData
    
End Property

Public Function AddObj(Group As clsGroup) As Boolean

'This function adds the passed in Group to this collection

    AddObj = False
    
    'The Group must have a groupid so add it to the collection with the groupid
    mCol.Add Group, CStr(Group.GroupID)
    
    AddObj = True

End Function
Private Function AddInternal(GroupID As Long) As clsGroup
    
    'create a new object
    Dim objNewMember As clsGroup
    Set objNewMember = New clsGroup

    'set the properties passed into the method
    objNewMember.GroupID = GroupID
    
    'Set the DxsRef and the TxsRef
    objNewMember.DxsRef = mvarDxsRef
    objNewMember.TxsRef = mvarTxsRef
    
    'Add the object to the collection
    mCol.Add objNewMember, CStr(GroupID)

    'return the object created
    Set AddInternal = objNewMember
    Set objNewMember = Nothing

End Function


Friend Function RefreshGroupsFromPatientID(PatientID As Long, CurrentDentalEncounterID As Long, Optional ToothID As Long = 0) As Boolean

'Get all the groups based on the PatientID

', Source, CurrentDentalEncounterID, and ToothID

Dim rs As Recordset
Dim sDxSQL As String
Dim sTxSQL As String
Dim rsDx As Recordset
Dim rsTx As Recordset
Dim lDxGroupID As Long
Dim lTxGroupID As Long
Dim TempGroup As clsGroup
Dim sSQL As String
Dim ar2DTemp As Variant
Dim ar1DTemp() As String
Dim lCounter As Long
Dim sDxCodeIDs As String
Dim sTxCodeIDs As String
Dim oCmd As CHCSII_Command.CHCSII_Cmd

    If gbUseStoredProc = True Then
        'Use stored procs
        Set oCmd = New CHCSII_Command.CHCSII_Cmd
        
        oCmd.SPName = "Dl_Note_Dxs_Load"

        Call oCmd.AddParam(False, False, "textString", PatientID)
        Call oCmd.AddParam(True, True, "", "")
        
        Set rsDxsToLoad = goconn.OpenRecordsetSP(oCmd)
        
    Else
        'Get the Dx disconnected recordsets for the given patient.
        sSQL = "SELECT DENTAL_ENCOUNTER.PATIENT_UNIT_ID," _
            & " DENTAL_ENCOUNTER_DX.DX_CODE_ID," _
            & " DENTAL_ENCOUNTER_DX.DX_ICD_CODE_ID," _
            & " DENTAL_ENCOUNTER_DX.CREATE_DATE_TIME," _
            & " DENTAL_ENCOUNTER_DX.MODIFIED_DATE_TIME," _
            & " DENTAL_ENCOUNTER_DX.PRIORITY," _
            & " DENTAL_ENCOUNTER_DX.PROVIDER_ID," _
            & " DENTAL_ENCOUNTER_DX.DENTAL_ENCOUNTER_ID," _
            & " DENTAL_ENCOUNTER_DX.STATUS_ID," _
            & " DENTAL_ENCOUNTER_DX.CHARTABLE," _
            & " DENTAL_ENCOUNTER_DX.DRC," _
            & " DENTAL_ENCOUNTER_DX.SIGNED," _
            & " DENTAL_ENCOUNTER_DX.PERIO_STATE," _
            & " DENTAL_ENCOUNTER_DX.DELETE_REASON," _
            & " DENTAL_ENCOUNTER_DX.MODIFIED_IN_DENTAL_ENCOUNTER," _
            & " DENTAL_ENCOUNTER_GROUPS.GROUP_ID," _
            & " DX_ICD_CODE_DEFINITION.ICD_CATEGORY," _
            & " DX_ICD_CODE_DEFINITION.ICD_CODE," _
            & " DX_ICD_CODE_DEFINITION.ICD_DEFINITION," _
            & " DX_ICD_CODE_DEFINITION.DRC AS DEFAULT_DRC," _
            & " DX_ICD_CODE_DEFINITION.CHART_RELATED," _
            & " DX_ICD_CODE_DEFINITION.EFFECTIVE_DATE," _
            & " DX_ICD_CODE_DEFINITION.EFFECTIVE_END_DATE," _
            & " DX_ICD_CODE_DEFINITION.DRCMIN," _
            & " DX_ICD_CODE_DEFINITION.DRCMax"
        
            sSQL = sSQL & " FROM DENTAL_ENCOUNTER_DX, DENTAL_ENCOUNTER," _
                & " DX_ICD_CODE_DEFINITION," _
                & " DENTAL_ENCOUNTER_GROUPS"
            sSQL = sSQL & " Where DENTAL_ENCOUNTER_DX.DENTAL_ENCOUNTER_ID = DENTAL_ENCOUNTER.DENTAL_ENCOUNTER_ID" _
                & " AND DENTAL_ENCOUNTER_DX.DX_ICD_CODE_ID = DX_ICD_CODE_DEFINITION.DX_ICD_CODE_ID" _
                & " AND DENTAL_ENCOUNTER_DX.DX_CODE_ID = DENTAL_ENCOUNTER_GROUPS.DX_CODE_ID" _
                & " AND (DENTAL_ENCOUNTER.PATIENT_UNIT_ID = " & PatientID & ") AND" _
                & " (DENTAL_ENCOUNTER_DX.STATUS_ID <> 21) AND" _
                & " (DENTAL_ENCOUNTER_DX.DX_CODE_ID > 0)" _
                & " ORDER BY DENTAL_ENCOUNTER_DX.CREATE_DATE_TIME"
        
        Set rsDxsToLoad = goconn.OpenRecordset(sSQL)
        rsDxsToLoad.ActiveConnection = Nothing
    End If
    
    
    '****************** BEGINING OF CHANGE
    'Get a comma delimited list of dx_code_ids to query the ss, tooth, loc, and subloc
    'This is the fastest way - string concat is slow
    If rsDxsToLoad.EOF = False Then
        ar2DTemp = rsDxsToLoad.GetRows(, , "Dx_Code_ID")
        ReDim ar1DTemp(UBound(ar2DTemp, 2)) As String
        For lCounter = 0 To UBound(ar2DTemp, 2)
            ar1DTemp(lCounter) = ar2DTemp(0, lCounter)
        Next
        sDxCodeIDs = Join(ar1DTemp, ",")
            
        If gbUseStoredProc = True Then
            'Use stored procs
            'Get the SurfaceSets and Teeth info
            Set oCmd = New CHCSII_Command.CHCSII_Cmd
            
            oCmd.SPName = "Dl_Note_Dxssteeth_Load"
    
            Call oCmd.AddParam(False, False, "textString", sDxCodeIDs)
            Call oCmd.AddParam(True, True, "", "")
            
            Set rsDxSSTeethToLoad = goconn.OpenRecordsetSP(oCmd)
                
            'Get the Location and Sublocation info
            Set oCmd = New CHCSII_Command.CHCSII_Cmd
            
            oCmd.SPName = "Dl_Note_Dxlocsubloc_Load"
    
            Call oCmd.AddParam(False, False, "textString", sDxCodeIDs)
            Call oCmd.AddParam(True, True, "", "")
            
            Set rsDxLocSubLocToLoad = goconn.OpenRecordsetSP(oCmd)
            
        Else
            'Get the recordset to load the surface and tooth for each Dx
            sSQL = "SELECT DENTAL_ENC_SURFACE_SET.SURFACE_SET_ID, DENTAL_ENC_SURFACE_SET.SURFACESET," _
            & " DENTAL_ENC_SURFACE_SET.COLOR_ID, DENTAL_ENC_SURFACE_SET.STATE," _
            & " DENTAL_ENC_SURFACE_SET.HATCH_ID, DENTAL_ENC_SURFACE_SET.SYMBOL_ID," _
            & " DENTAL_ENC_SURFACE_SET.SURFACEABBR, DENTAL_ENC_SURFACE_SET.SUB_MENU_ID," _
            & " DENTAL_ENCOUNTER_TEETH.DENTAL_ENCOUNTER_ID, DENTAL_ENCOUNTER_TEETH.TOOTH_ID," _
            & " DENTAL_ENCOUNTER_TEETH.DRC, DENTAL_ENCOUNTER_TEETH.TOOTH_NAME," _
            & " DENTAL_ENCOUNTER_TEETH.MESIAL_VERSION, DENTAL_ENCOUNTER_TEETH.DISTAL_VERSION," _
            & " DENTAL_ENCOUNTER_TEETH.FACIAL_VERSION, DENTAL_ENCOUNTER_TEETH.LINGUAL_VERSION," _
            & " DENTAL_ENCOUNTER_TEETH.ROTATED_MESIAL, DENTAL_ENCOUNTER_TEETH.ROTATED_DISTAL," _
            & " DENTAL_ENCOUNTER_TEETH.INTRUDED, DENTAL_ENCOUNTER_TEETH.EXTRUDED," _
            & " DENTAL_ENCOUNTER_TEETH.DELETED AS TEETH_DELETED, DENTAL_ENCOUNTER_TEETH.ORIENTATION," _
            & " DENTAL_ENCOUNTER_TEETH.DENTAL_ENCOUNTER_TEETH_ID, DENTAL_ENCOUNTER_DXS_SS.DX_CODE_ID"
            
            If goChcsiiConn.UseSQLServer Then
                'SQL Server join syntax
                sSQL = sSQL & " FROM DENTAL_ENCOUNTER_DXS_SS INNER JOIN DENTAL_ENC_SURFACE_SET ON" _
                & " DENTAL_ENCOUNTER_DXS_SS.SURFACE_SET_ID = DENTAL_ENC_SURFACE_SET.SURFACE_SET_ID LEFT OUTER JOIN" _
                & " DENTAL_ENCOUNTER_TEETH ON DENTAL_ENC_SURFACE_SET.DENTAL_ENCOUNTER_TEETH_ID = DENTAL_ENCOUNTER_TEETH.DENTAL_ENCOUNTER_TEETH_ID" _
                & " WHERE (DENTAL_ENC_SURFACE_SET.DELETED = 0) AND (DENTAL_ENCOUNTER_DXS_SS.DX_CODE_ID IN (" & sDxCodeIDs & "))" _
                & " ORDER BY DENTAL_ENC_SURFACE_SET.SURFACE_SET_ID"
            Else
                'Oracle join syntax
                sSQL = sSQL & " FROM DENTAL_ENCOUNTER_DXS_SS, DENTAL_ENC_SURFACE_SET, DENTAL_ENCOUNTER_TEETH" _
                & " Where DENTAL_ENCOUNTER_DXS_SS.SURFACE_SET_ID = DENTAL_ENC_SURFACE_SET.SURFACE_SET_ID AND" _
                & " DENTAL_ENC_SURFACE_SET.DENTAL_ENCOUNTER_TEETH_ID = DENTAL_ENCOUNTER_TEETH.DENTAL_ENCOUNTER_TEETH_ID" _
                & " (+) AND (DENTAL_ENC_SURFACE_SET.DELETED = 0) AND" _
                & " (DENTAL_ENCOUNTER_DXS_SS.DX_CODE_ID IN (" & sDxCodeIDs & "))" _
                & " ORDER BY DENTAL_ENC_SURFACE_SET.SURFACE_SET_ID"
            
            End If
        
            Set rsDxSSTeethToLoad = goconn.OpenRecordset(sSQL)
            Set rsDxSSTeethToLoad.ActiveConnection = Nothing
    
            'Get the recordset to load the loc and subloc for each Dx
            sSQL = "SELECT NOTE_DXS_LOCATION.DX_LOCATION_ID, NOTE_DXS_LOCATION.DX_CODE_ID," _
                & " NOTE_DXS_LOCATION.DELETED, NOTE_LOCATION.LOCATION_ID," _
                & " NOTE_LOCATION.LOCATION_NAME, NOTE_DX_SUB_LOCATION.DELETED AS SUBLOC_DELETED," _
                & " NOTE_DX_SUB_LOCATION.DX_SUBLOC_ID, NOTE_SUB_LOCATION.SUB_LOCATION_ID," _
                & " NOTE_SUB_LOCATION.SUB_LOCATION_NAME"
                
            If goChcsiiConn.UseSQLServer Then
                'SQL Server join syntax
                sSQL = sSQL & " FROM NOTE_LOCATION INNER JOIN NOTE_DXS_LOCATION ON NOTE_LOCATION.LOCATION_ID = NOTE_DXS_LOCATION.LOCATION_ID LEFT OUTER JOIN" _
                & " NOTE_DX_SUB_LOCATION ON NOTE_DXS_LOCATION.DX_LOCATION_ID = NOTE_DX_SUB_LOCATION.DX_LOCATION_ID LEFT OUTER JOIN" _
                & " NOTE_SUB_LOCATION ON NOTE_DX_SUB_LOCATION.SUB_LOCATION_ID = NOTE_SUB_LOCATION.SUB_LOCATION_ID" _
                & " WHERE (NOTE_DXS_LOCATION.DELETED = 0) AND (NOTE_DXS_LOCATION.DX_CODE_ID IN (" & sDxCodeIDs & "))"
            Else
                'Oracle join syntax
                sSQL = sSQL & " FROM NOTE_LOCATION, NOTE_DXS_LOCATION, NOTE_DX_SUB_LOCATION, NOTE_SUB_LOCATION" _
                & " Where NOTE_LOCATION.LOCATION_ID = NOTE_DXS_LOCATION.LOCATION_ID AND" _
                & " NOTE_DXS_LOCATION.DX_LOCATION_ID = NOTE_DX_SUB_LOCATION.DX_LOCATION_ID (+) AND" _
                & " NOTE_DX_SUB_LOCATION.SUB_LOCATION_ID = NOTE_SUB_LOCATION.SUB_LOCATION_ID (+) AND" _
                & " (NOTE_DXS_LOCATION.DELETED = 0) AND" _
                & " (NOTE_DXS_LOCATION.DX_CODE_ID IN (" & sDxCodeIDs & "))"
            End If
            
            Set rsDxLocSubLocToLoad = goconn.OpenRecordset(sSQL)
            rsDxLocSubLocToLoad.ActiveConnection = Nothing
        End If
    End If
    '********** END OF CHANGE
    

    If gbUseStoredProc = True Then
        'Use stored procs
        Set oCmd = New CHCSII_Command.CHCSII_Cmd
        
        oCmd.SPName = "Dl_Note_Txs_Load"

        Call oCmd.AddParam(False, False, "textString", PatientID)
        Call oCmd.AddParam(False, False, "textString", CInt(gbDentalTriage))
        Call oCmd.AddParam(True, True, "", "")
        
        Set rsTxsToLoad = goconn.OpenRecordsetSP(oCmd)
        
    Else
        
        'Get the Tx disconnected recordsets for the given patient.
        sSQL = "SELECT DENTAL_ENCOUNTER_TX.TX_CODE_ID," _
            & " DENTAL_ENCOUNTER_TX.DENTAL_ENCOUNTER_ID," _
            & " DENTAL_ENCOUNTER_TX.TX_ADA_CODE_ID," _
            & " DENTAL_ENCOUNTER_TX.STATUS_ID," _
            & " DENTAL_ENCOUNTER_TX.CREATE_DATE_TIME," _
            & " DENTAL_ENCOUNTER_TX.MODIFIED_DATE_TIME," _
            & " DENTAL_ENCOUNTER_TX.PROVIDER_ID," _
            & " DENTAL_ENCOUNTER_TX.CHARTABLE," _
            & " DENTAL_ENCOUNTER_TX.COMMENTS," _
            & " DENTAL_ENCOUNTER_TX.MATERIALS," _
            & " DENTAL_ENCOUNTER_TX.PROGNOSIS_ID," _
            & " DENTAL_ENCOUNTER_TX.SIGNED," _
            & " DENTAL_ENCOUNTER_TX.PHASE," _
            & " DENTAL_ENCOUNTER_TX.WORKLOAD_ID," _
            & " DENTAL_ENCOUNTER_TX.DELETE_REASON," _
            & " DENTAL_ENCOUNTER_TX.MODIFIED_IN_DENTAL_ENCOUNTER,"
        sSQL = sSQL & " TX_ADA_CODE_DEFINITION.ADA_CATEGORY_CODE," _
            & " TX_ADA_CODE_DEFINITION.ADA_DEFINITION," _
            & " TX_ADA_CODE_DEFINITION.ADA_ALIAS," _
            & " TX_ADA_CODE_DEFINITION.CHART_RELATED," _
            & " TX_ADA_CODE_DEFINITION.LABFLAG," _
            & " TX_ADA_CODE_DEFINITION.ADA_HELP_NARATIVE," _
            & " TX_ADA_CODE_DEFINITION.ADA_CODE," _
            & " TX_ADA_CODE_DEFINITION.PERIO," _
            & " TX_ADA_CODE_DEFINITION.FPD_FLAG," _
            & " TX_ADA_CODE_DEFINITION.RPD_FLAG," _
            & " TX_ADA_CODE_DEFINITION.DENTURE_FLAG," _
            & " DENTAL_ENCOUNTER_GROUPS.GROUP_ID,"
        sSQL = sSQL & " TX_ADA_CODE_DEFINITION.DWV" _
            & " FROM DENTAL_ENCOUNTER_TX, TX_ADA_CODE_DEFINITION," _
            & " DENTAL_ENCOUNTER," _
            & " DENTAL_ENCOUNTER_GROUPS" _
            & " Where DENTAL_ENCOUNTER_TX.TX_ADA_CODE_ID = TX_ADA_CODE_DEFINITION.TX_ADA_CODE_ID" _
            & " AND"
        sSQL = sSQL & " DENTAL_ENCOUNTER_TX.DENTAL_ENCOUNTER_ID = DENTAL_ENCOUNTER.DENTAL_ENCOUNTER_ID" _
            & " AND" _
            & " DENTAL_ENCOUNTER_TX.TX_CODE_ID = DENTAL_ENCOUNTER_GROUPS.TX_CODE_ID" _
            & " AND (((DENTAL_ENCOUNTER_TX.STATUS_ID <> 104) AND" _
            & " (DENTAL_ENCOUNTER_TX.SIGNED = 0) AND" _
            & " (DENTAL_ENCOUNTER.PATIENT_UNIT_ID = " & PatientID & ")) OR" _
            & " "
                
        If gbDentalTriage = False Then
            sSQL = sSQL & " ((DENTAL_ENCOUNTER_TX.STATUS_ID IN (100, 101, 102, 103, 105)) AND"
        Else 'If the dentaltriage command line switch is on, then query all statuses except deleted
            sSQL = sSQL & " ((DENTAL_ENCOUNTER_TX.STATUS_ID <> 104) AND"
        End If
            
        sSQL = sSQL & " (DENTAL_ENCOUNTER.PATIENT_UNIT_ID = " & PatientID & ") AND" _
            & " (DENTAL_ENCOUNTER_TX.SIGNED = -1)" _
            & " ))" _
            & " ORDER BY DENTAL_ENCOUNTER_TX.CREATE_DATE_TIME"
        Set rsTxsToLoad = goconn.OpenRecordset(sSQL)
        rsTxsToLoad.ActiveConnection = Nothing
    End If
    
    '**************** BEGINNING OF CHANGE
    'Get a comma delimited list of tx_code_ids to query the ss, tooth, loc, and subloc
    'This is the fastest way - string concat is slow
    If rsTxsToLoad.EOF = False Then
        ar2DTemp = rsTxsToLoad.GetRows(, , "Tx_Code_ID")
        ReDim ar1DTemp(UBound(ar2DTemp, 2)) As String
        For lCounter = 0 To UBound(ar2DTemp, 2)
            ar1DTemp(lCounter) = ar2DTemp(0, lCounter)
        Next
        sTxCodeIDs = Join(ar1DTemp, ",")
    
        If gbUseStoredProc = True Then
            'Use stored procs
            'Get the SurfaceSets and Teeth info
            Set oCmd = New CHCSII_Command.CHCSII_Cmd
            
            oCmd.SPName = "Dl_Note_Txssteeth_Load"
    
            Call oCmd.AddParam(False, False, "textString", sTxCodeIDs)
            Call oCmd.AddParam(True, True, "", "")
            
            Set rsTxSSTeethToLoad = goconn.OpenRecordsetSP(oCmd)
                
            'Get the Location and Sublocation info
            Set oCmd = New CHCSII_Command.CHCSII_Cmd
            
            oCmd.SPName = "Dl_Note_Txlocsubloc_Load"
    
            Call oCmd.AddParam(False, False, "textString", sTxCodeIDs)
            Call oCmd.AddParam(True, True, "", "")
            
            Set rsTxLocSubLocToLoad = goconn.OpenRecordsetSP(oCmd)
            
        Else
    
            'Get the recordset to load the surface and tooth for each Tx
            sSQL = "SELECT DENTAL_ENC_SURFACE_SET.SURFACE_SET_ID, DENTAL_ENC_SURFACE_SET.SURFACESET," _
                & " DENTAL_ENC_SURFACE_SET.COLOR_ID, DENTAL_ENC_SURFACE_SET.STATE," _
                & " DENTAL_ENC_SURFACE_SET.HATCH_ID, DENTAL_ENC_SURFACE_SET.SYMBOL_ID," _
                & " DENTAL_ENC_SURFACE_SET.SURFACEABBR, DENTAL_ENC_SURFACE_SET.DELETED," _
                & " DENTAL_ENC_SURFACE_SET.SUB_MENU_ID, DENTAL_ENCOUNTER_TEETH.DENTAL_ENCOUNTER_ID," _
                & " DENTAL_ENCOUNTER_TEETH.DENTAL_ENCOUNTER_TEETH_ID, DENTAL_ENCOUNTER_TEETH.TOOTH_ID," _
                & " DENTAL_ENCOUNTER_TEETH.DRC, DENTAL_ENCOUNTER_TEETH.TOOTH_NAME," _
                & " DENTAL_ENCOUNTER_TEETH.MESIAL_VERSION, DENTAL_ENCOUNTER_TEETH.DISTAL_VERSION," _
                & " DENTAL_ENCOUNTER_TEETH.FACIAL_VERSION, DENTAL_ENCOUNTER_TEETH.LINGUAL_VERSION," _
                & " DENTAL_ENCOUNTER_TEETH.ROTATED_MESIAL, DENTAL_ENCOUNTER_TEETH.ROTATED_DISTAL," _
                & " DENTAL_ENCOUNTER_TEETH.INTRUDED, DENTAL_ENCOUNTER_TEETH.EXTRUDED," _
                & " DENTAL_ENCOUNTER_TEETH.DELETED AS TEETH_DELETED, DENTAL_ENCOUNTER_TEETH.ORIENTATION," _
                & " DENTAL_ENCOUNTER_TXS_SS.TX_CODE_ID"
                
            If goChcsiiConn.UseSQLServer Then
                'SQL Server join format
                sSQL = sSQL & " FROM DENTAL_ENCOUNTER_TXS_SS INNER JOIN DENTAL_ENC_SURFACE_SET ON" _
                & " DENTAL_ENCOUNTER_TXS_SS.SURFACE_SET_ID = DENTAL_ENC_SURFACE_SET.SURFACE_SET_ID LEFT OUTER JOIN DENTAL_ENCOUNTER_TEETH ON" _
                & " DENTAL_ENC_SURFACE_SET.DENTAL_ENCOUNTER_TEETH_ID = DENTAL_ENCOUNTER_TEETH.DENTAL_ENCOUNTER_TEETH_ID" _
                & " WHERE (DENTAL_ENC_SURFACE_SET.DELETED = 0) AND (DENTAL_ENCOUNTER_TXS_SS.TX_CODE_ID IN (" & sTxCodeIDs & "))" _
                & " ORDER BY DENTAL_ENC_SURFACE_SET.SURFACE_SET_ID"
            Else
                'Oracle join format
                sSQL = sSQL & " FROM DENTAL_ENCOUNTER_TXS_SS, DENTAL_ENC_SURFACE_SET," _
                & " DENTAL_ENCOUNTER_TEETH" _
                & " Where DENTAL_ENCOUNTER_TXS_SS.SURFACE_SET_ID = DENTAL_ENC_SURFACE_SET.SURFACE_SET_ID AND" _
                & " DENTAL_ENC_SURFACE_SET.DENTAL_ENCOUNTER_TEETH_ID = DENTAL_ENCOUNTER_TEETH.DENTAL_ENCOUNTER_TEETH_ID" _
                & " (+) AND (DENTAL_ENC_SURFACE_SET.DELETED = 0) AND" _
                & " (DENTAL_ENCOUNTER_TXS_SS.TX_CODE_ID IN (" & sTxCodeIDs & "))" _
                & " ORDER BY DENTAL_ENC_SURFACE_SET.SURFACE_SET_ID"
            End If
            
            Set rsTxSSTeethToLoad = goconn.OpenRecordset(sSQL)
            rsTxSSTeethToLoad.ActiveConnection = Nothing
    
            'Get the recordset to load the loc and subloc for each Tx
            sSQL = "SELECT NOTE_LOCATION.LOCATION_NAME, NOTE_TXS_LOCATION.DELETED," _
                & " NOTE_TXS_LOCATION.LOCATION_ID, NOTE_TXS_LOCATION.TX_LOCATION_ID," _
                & " NOTE_TXS_LOCATION.TX_CODE_ID, NOTE_TX_SUB_LOCATION.SUB_LOCATION_ID," _
                & " NOTE_TX_SUB_LOCATION.DELETED AS SUBLOC_DELETED, NOTE_TX_SUB_LOCATION.TX_SUBLOC_ID," _
                & " NOTE_SUB_LOCATION.SUB_LOCATION_NAME"
            If goChcsiiConn.UseSQLServer Then
                'SQL Server join syntax
                sSQL = sSQL & " FROM NOTE_LOCATION INNER JOIN NOTE_TXS_LOCATION ON NOTE_LOCATION.LOCATION_ID = NOTE_TXS_LOCATION.LOCATION_ID LEFT OUTER JOIN" _
                & " NOTE_SUB_LOCATION RIGHT OUTER JOIN NOTE_TX_SUB_LOCATION ON NOTE_SUB_LOCATION.SUB_LOCATION_ID = NOTE_TX_SUB_LOCATION.SUB_LOCATION_ID ON" _
                & " NOTE_TXS_LOCATION.TX_LOCATION_ID = NOTE_TX_SUB_LOCATION.TX_LOCATION_ID" _
                & " WHERE (NOTE_TXS_LOCATION.DELETED = 0) AND (NOTE_TXS_LOCATION.TX_CODE_ID IN (" & sTxCodeIDs & "))"
            Else
                'Oracle join syntax
                sSQL = sSQL & " FROM NOTE_SUB_LOCATION, NOTE_TX_SUB_LOCATION, NOTE_TXS_LOCATION , NOTE_LOCATION" _
                & " WHERE NOTE_SUB_LOCATION.SUB_LOCATION_ID (+) = NOTE_TX_SUB_LOCATION.SUB_LOCATION_ID AND" _
                & " NOTE_TX_SUB_LOCATION.TX_LOCATION_ID (+) = NOTE_TXS_LOCATION.TX_LOCATION_ID AND" _
                & " NOTE_TXS_LOCATION.LOCATION_ID = NOTE_LOCATION.LOCATION_ID AND (NOTE_TXS_LOCATION.DELETED = 0) AND" _
                & " (NOTE_TXS_LOCATION.TX_CODE_ID IN (" & sTxCodeIDs & "))"
            End If
            
            Set rsTxLocSubLocToLoad = goconn.OpenRecordset(sSQL)
            rsTxLocSubLocToLoad.ActiveConnection = Nothing
        End If
    End If
    '******** END OF CHANGE
    
    RefreshGroupsFromPatientID = False
    
    mvarCurrentDentalEncounterID = CurrentDentalEncounterID
    
    Set mCol = New Collection

    If gbUseStoredProc = True Then
        'Use stored procs
        Set oCmd = New CHCSII_Command.CHCSII_Cmd
        
        oCmd.SPName = "Dl_Note_Groups_Load"

        Call oCmd.AddParam(False, False, "textString", PatientID)
        Call oCmd.AddParam(True, True, "", "")
        
        Set rs = goconn.OpenRecordsetSP(oCmd)
                
        Do Until rs.EOF
            Set TempGroup = AddInternal(rs("Group_ID"))
            'Load the Dxs for this group
            TempGroup.Dxs.RefreshDxsFromRS rsDxsToLoad, rs("Group_ID"), CurrentDentalEncounterID
            'Load the Txs for this group
            TempGroup.Txs.RefreshTxsFromRS rsTxsToLoad, rs("Group_ID"), CurrentDentalEncounterID
            
            LoadSSTeethandLocSubLoc TempGroup.Dxs, TempGroup.Txs
            
            rs.MoveNext
        Loop
        
    Else
        'The following is a cludge because CHCSII_DAS does not support UNIONs.
        'Get a list of groups that are used by Dxs and Txs
            sDxSQL = "SELECT DISTINCT DENTAL_ENCOUNTER_GROUPS.GROUP_ID" _
                & " FROM DENTAL_ENCOUNTER_DX, DENTAL_ENCOUNTER," _
                    & " DENTAL_ENCOUNTER_GROUPS" _
                & " Where DENTAL_ENCOUNTER_DX.DENTAL_ENCOUNTER_ID = DENTAL_ENCOUNTER.DENTAL_ENCOUNTER_ID" _
                     & " AND" _
                    & " DENTAL_ENCOUNTER_DX.DX_CODE_ID = DENTAL_ENCOUNTER_GROUPS.DX_CODE_ID" _
                     & " AND (DENTAL_ENCOUNTER.PATIENT_UNIT_ID = " & PatientID & ")" _
                & " ORDER BY DENTAL_ENCOUNTER_GROUPS.GROUP_ID"
        
            sTxSQL = "SELECT DISTINCT DENTAL_ENCOUNTER_GROUPS.GROUP_ID" _
                & " FROM DENTAL_ENCOUNTER, DENTAL_ENCOUNTER_TX," _
                    & " DENTAL_ENCOUNTER_GROUPS" _
                & " Where DENTAL_ENCOUNTER.DENTAL_ENCOUNTER_ID = DENTAL_ENCOUNTER_TX.DENTAL_ENCOUNTER_ID" _
                     & " AND" _
                    & " DENTAL_ENCOUNTER_TX.TX_CODE_ID = DENTAL_ENCOUNTER_GROUPS.TX_CODE_ID" _
                     & " AND (DENTAL_ENCOUNTER.PATIENT_UNIT_ID = " & PatientID & ")" _
                & " ORDER BY DENTAL_ENCOUNTER_GROUPS.GROUP_ID"
        
            Set rsDx = goconn.OpenRecordset(sDxSQL)
            Set rsTx = goconn.OpenRecordset(sTxSQL)
            
            'Loop through all the groups and load the Dxs and Txs for each
            Do Until rsDx.EOF And rsTx.EOF
                If rsDx.EOF = False Then
                    lDxGroupID = Val(rsDx("Group_ID") & "")
                Else
                    lDxGroupID = -1
                End If
                
                If rsTx.EOF = False Then
                    lTxGroupID = Val(rsTx("Group_ID") & "")
                Else
                    lTxGroupID = -1
                End If
                
                If lDxGroupID = -1 Then
                    'The rsDx is at the end of the recordset so add the Tx group
                    Set TempGroup = AddInternal(lTxGroupID)
                    'Load the Dxs for this group
                    TempGroup.Dxs.RefreshDxsFromRS rsDxsToLoad, lTxGroupID, CurrentDentalEncounterID
                    'Load the Txs for this group
                    TempGroup.Txs.RefreshTxsFromRS rsTxsToLoad, lTxGroupID, CurrentDentalEncounterID
                    LoadSSTeethandLocSubLoc TempGroup.Dxs, TempGroup.Txs
        '            TempGroup.Dxs.RefreshDxsFromPatientID lTxGroupID, PatientID, Source, CurrentDentalEncounterID, ToothID
        '            TempGroup.Txs.RefreshTxsFromPatientID lTxGroupID, PatientID, Source, CurrentDentalEncounterID, ToothID
                    rsTx.MoveNext
                ElseIf lTxGroupID = -1 Then
                    'The rsTx is at the end of the recordset so add the Dx group
                    Set TempGroup = AddInternal(lDxGroupID)
                    TempGroup.Dxs.RefreshDxsFromRS rsDxsToLoad, lDxGroupID, CurrentDentalEncounterID
                    TempGroup.Txs.RefreshTxsFromRS rsTxsToLoad, lDxGroupID, CurrentDentalEncounterID
                    LoadSSTeethandLocSubLoc TempGroup.Dxs, TempGroup.Txs
        '            TempGroup.Dxs.RefreshDxsFromPatientID lDxGroupID, PatientID, Source, CurrentDentalEncounterID, ToothID
        '            TempGroup.Txs.RefreshTxsFromPatientID lDxGroupID, PatientID, Source, CurrentDentalEncounterID, ToothID
                    rsDx.MoveNext
                ElseIf lDxGroupID = lTxGroupID Then
                    'The groups are the same so add the group
                    Set TempGroup = AddInternal(lDxGroupID)
                    TempGroup.Dxs.RefreshDxsFromRS rsDxsToLoad, lDxGroupID, CurrentDentalEncounterID
                    TempGroup.Txs.RefreshTxsFromRS rsTxsToLoad, lDxGroupID, CurrentDentalEncounterID
                    LoadSSTeethandLocSubLoc TempGroup.Dxs, TempGroup.Txs
        '            TempGroup.Dxs.RefreshDxsFromPatientID lDxGroupID, PatientID, Source, CurrentDentalEncounterID, ToothID
        '            TempGroup.Txs.RefreshTxsFromPatientID lDxGroupID, PatientID, Source, CurrentDentalEncounterID, ToothID
                    rsDx.MoveNext
                    rsTx.MoveNext
                ElseIf lDxGroupID < lTxGroupID Then
                    'The Dx group is less than the Tx group, so load the Dx group
                    Set TempGroup = AddInternal(lDxGroupID)
                    TempGroup.Dxs.RefreshDxsFromRS rsDxsToLoad, lDxGroupID, CurrentDentalEncounterID
                    TempGroup.Txs.RefreshTxsFromRS rsTxsToLoad, lDxGroupID, CurrentDentalEncounterID
                    LoadSSTeethandLocSubLoc TempGroup.Dxs, TempGroup.Txs
        '            TempGroup.Dxs.RefreshDxsFromPatientID lDxGroupID, PatientID, Source, CurrentDentalEncounterID, ToothID
        '            TempGroup.Txs.RefreshTxsFromPatientID lDxGroupID, PatientID, Source, CurrentDentalEncounterID, ToothID
                    rsDx.MoveNext
                ElseIf lDxGroupID > lTxGroupID Then
                    'The Dx group is greater than the Tx group, so load the Tx group
                    Set TempGroup = AddInternal(lTxGroupID)
                    TempGroup.Dxs.RefreshDxsFromRS rsDxsToLoad, lTxGroupID, CurrentDentalEncounterID
                    TempGroup.Txs.RefreshTxsFromRS rsTxsToLoad, lTxGroupID, CurrentDentalEncounterID
                    LoadSSTeethandLocSubLoc TempGroup.Dxs, TempGroup.Txs
        '            TempGroup.Dxs.RefreshDxsFromPatientID lTxGroupID, PatientID, Source, CurrentDentalEncounterID, ToothID
        '            TempGroup.Txs.RefreshTxsFromPatientID lTxGroupID, PatientID, Source, CurrentDentalEncounterID, ToothID
                    rsTx.MoveNext
                End If
            Loop
            Set TempGroup = Nothing
    End If
    
    'Remove any groups that don't have Dxs and Txs because the LoadDxs or LoadTxs has
    ' filtered them out.
    For Each TempGroup In mCol
        If TempGroup.Dxs.Count = 0 And TempGroup.Txs.Count = 0 Then
            'Remove the group from the collection
            mCol.Remove TempGroup.GroupID
        End If
    Next
    
    Set rsDxsToLoad = Nothing
    Set rsTxsToLoad = Nothing
    Set rsDxSSTeethToLoad = Nothing
    Set rsDxLocSubLocToLoad = Nothing
    Set rsTxSSTeethToLoad = Nothing
    Set rsTxLocSubLocToLoad = Nothing
    
    RefreshGroupsFromPatientID = True
    
End Function

Public Function Add() As clsGroup
    
    'create a new object
    Dim objNewMember As clsGroup
    Dim sSQL As String
    Dim rs As Recordset
    
    Set objNewMember = New clsGroup

'    'Get a new groupid from the group_id_seq sequence
'    sSQL = "Select GROUP_ID_SEQ.NEXTVAL as ID from Dual"
'
'    Set rs = goconn.OpenRecordset(sSQL)
'
'    If rs.EOF Then
'        Err.Raise vbObjectError + 1, "colGroups.Add", "Cannot get NextVal from the sequence 'Group_ID_Seq'"
'    Else
        
        
        objNewMember.GroupID = goconn.GetNextID(esGROUP_ID_SEQ)
        'Get the defaults
        objNewMember.Dxs.CurrentDentalEncounterID = mvarCurrentDentalEncounterID
        objNewMember.Dxs.GroupID = objNewMember.GroupID
        
        objNewMember.Txs.CurrentDentalEncounterID = mvarCurrentDentalEncounterID
        objNewMember.Txs.GroupID = objNewMember.GroupID
        
        mCol.Add objNewMember, CStr(objNewMember.GroupID)
        
        'return the object created
        Set Add = objNewMember
'    End If
'    rs.Close
    
    'Set the DxsRef and the TxsRef
    objNewMember.DxsRef = mvarDxsRef
    objNewMember.TxsRef = mvarTxsRef

    Set rs = Nothing
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As clsGroup
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    
    'destroys collection when this class is terminated
    Set mCol = Nothing
    Set mvarTxsRef = Nothing
    Set mvarDxsRef = Nothing
    Set rsDxsToLoad = Nothing
    Set rsTxsToLoad = Nothing
    Set rsDxSSTeethToLoad = Nothing
    Set rsTxSSTeethToLoad = Nothing
    Set rsDxLocSubLocToLoad = Nothing
    Set rsTxLocSubLocToLoad = Nothing
    
End Sub

