VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EncTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const Module = "EncTransfer"
Private mobjTables As XferTables
Private oDMBE As IDMBE.I_DMBE
Private DataSQL As ICHCSII_DAS
Private ImportExport As ImportExport
Private mobjShared As CWShared
Public Comm As Comm

'--- array containing data for audit history report ---
'Public gatReadinessHist()       As gtReadiness
'Public AllRecordsInMemory()     As gtReadiness  'As gtAllergy
Public AllRecordsInMemory     As Dictionary
Public NumAllRecordsInMemory    As Long
Public Sub Init()
    Set mobjShared = New CWShared
    
    Set mobjTables = New XferTables
    Set AllRecordsInMemory = New Dictionary
    
    LoadEncounterTables mobjTables
    'mobjTables.Load

End Sub


Public Function Transfer(DMBEObj As IDMBE.I_DMBE) As Variant
Dim objTransfer As I_RSTransfer
Dim lRet As Long



   '     This routine will take records from:
   '
   '           1.  The outside world anything other than CDR & LDDB
   '                and save them to LDDB
   '                It will then move to CDR from LDDB if the original
   '                request was from the outside world to CDR
   
   '
   '           2.  The inside world either CDR or LDDB
   '               and save them to the outside world
   '               CDR and LDDB (if it exists) will be Identical after
   '               running this routine
   

 On Error GoTo haveerror

 Dim Patient As Variant
 Dim Data_LDDB As Boolean
 Dim count As Long
 Dim MergeWithLDDB As Boolean
 Dim GoodReturnCode As Long         ' returns 0 if good to be compatible with DMBE
 Dim OperatorResponse As VbMsgBoxResult
 
 '...Kluge fix for not logged on to both databases
 Dim NotBothCDRandLDDB As Boolean   ' Remove this variable and all references once can talk to both CDR and LDDB
 
 ' SCR #22581;   Developer: Jane Shen 04/01/2002 05:05 PM
 Dim TransferResult As TMIPDataRetrievalResult
  
 NotBothCDRandLDDB = True
     
 Set oDMBE = DMBEObj
 Set DataSQL = gobjCHCSIIConn.CHCSII_DAS(Auto)
 
 '...Assume a bad deal
 GoodReturnCode = -15  ' no records for dataset
 
 '...Module defined variable
 Set ImportExport = New ImportExport
 
 With oDMBE

    If .Source = eDM_LDDB And Not (mobjShared.LocalDBConnected Or mobjShared.MySQLDBConnected) Then
       MsgBxARMd "Local Database (LDDB) not Connected", vbInformation + vbOKOnly, "Terminating Request"
       Exit Function
    End If

    If .Source = eDM_CDR And Not mobjShared.CDRConnected Then
       MsgBxARMd "CHCSII Database (CDR) not Connected", vbInformation + vbOKOnly, "Terminating Request"
       Exit Function
    End If
      
    If .Destination = eDM_LDDB And Not (mobjShared.LocalDBConnected Or mobjShared.MySQLDBConnected) Then
       MsgBxARMd "Local Database (LDDB) not Connected", vbInformation + vbOKOnly, "Terminating Request"
       Exit Function
    ElseIf .Destination = eDM_CDR And Not mobjShared.CDRConnected Then
       MsgBxARMd "CHCSII Database (CDR) not Connected", vbInformation + vbOKOnly, "Terminating Request"
       Exit Function
    End If

    '...Where will data come from
    Select Case .Source
    
      Case eDM_LDDB, eDM_CDR    '...Get Data from CDR or LDDB
          
           Data_LDDB = (.Source = eDM_LDDB)
           
           '...Patient List should have only 1 entry in it _
               however, if they decide to do multiple patients _
               at a time this code can handle it
            Patient = .PatientObj.UnitNumber
'           For Each Patient In .PatientList
           
            '...Reset Merged Records in Memory if necessary
            'Erase AllRecordsInMemory
            AllRecordsInMemory.RemoveAll
            'NumAllRecordsInMemory = 0
            NumAllRecordsInMemory = AllRecordsInMemory.count
           
            '...Indicate not merging To LDDB
            MergeWithLDDB = False
            
            '...Always bring in the most recent data if CDR connected
            If mobjShared.CDRConnected And (Not Data_LDDB And NotBothCDRandLDDB _
               Or Data_LDDB And Not NotBothCDRandLDDB) Then
                '...Get rquired records in CDR (user must already be logged on)
                If BuildRecordSetsCDR(CStr(Patient)) Then
                  '...If LDDB Exists then always Update LDDB
                  If mobjShared.LocalDBConnected Then
                   OperatorResponse = MsgBxARMd("Obtained Records from CDR..." & vbCrLf & "Do you wish to merge records with LDDB Records?", vbInformation + vbYesNoCancel, "Update LDDB")
                   Select Case OperatorResponse
                     Case vbYes
                       MergeWithLDDB = True
                     Case vbNo
                     Case Else
                       '...Clean up stuff here before exiting
                       GoTo getout
                    End Select
                  End If
                End If
            End If

            If Data_LDDB Or MergeWithLDDB Then
              '...Get required records in LDDB (User must already be logged on)
              ' SCR #22581;   Developer: Jane Shen 04/01/2002 05:06 PM
              TransferResult = BuildRecordSetsGEMS(CStr(Patient), DMBEObj)
              If TransferResult = All_Data_Have_Been_Sent Or TransferResult = No_Data Then
                Transfer = -15  'No Data to Send
                Exit Function
              End If
            End If
            
            If MergeWithLDDB Then
              '...Save Data to LDDB
              MergeRecords CStr(Patient), False
              
              '...Put Records in Memory
              MergeRecords CStr(Patient), True
              
              '...Update CDR
              MSaveDataToCDR
              
              '...remove all recordsets built
              ImportExport.ClearRSCollection
              
              '...Get rquired records in LDDB
              TransferResult = BuildRecordSetsGEMS(CStr(Patient), DMBEObj)
              
            End If
           
'           Next Patient
     
          '...Where does this data go
          
          If ImportExport.RecordSetsTransfered <> 0 Then
          
            Select Case oDMBE.Destination
                
                Case eDM_CDR
                
                If Not MergeWithLDDB Then
                  MsgBxARMd "CHCSII Database (CDR) not Connected", vbInformation + vbOKOnly, "Terminating Request"
                  Exit Function
                  End If
                   
               Case eDM_LDDB
                  ' Process already completed
                  Exit Function
                   
               Case Else            ' Outside world
'
'                    If oDMBE.Source = eDM_LDDB And oDMBE.Destination = eDM_TMIP Then
'                        Dim i As Integer
'                        For i = 1 To colTMIPEncNum.count
'                            TransferSymptoms oDMBE, colTMIPEncNum(i)
'                        Next
'                    End If
                    Set objTransfer = Me.Comm.InitializeOLEServer(NCID_PROBLEMS_APPLICATION)
                    lRet = objTransfer.Transfer(DMBEObj)
                                
                    ' Begin:SCR #24762;   Developer: Jane Shen 06/14/2002 04:08 PM
                    ' Note: Function Changed:    I_Transfer:Transfer()    Impact of Change:
                    'Begin:SCR #23307;   Developer: Jane Shen 04/24/2002 03:43 PM
                    ' End: SCR #24762;
                    Set objTransfer = Me.Comm.InitializeOLEServer(NCID_MEDICATIONS_APPLICATION)
                    lRet = objTransfer.Transfer(DMBEObj)
                    'End: SCR #23307;
                    
                    'OA 4/12/2010 No Readiness
                    'Begin:SCR #24250;   Developer: Jane Shen 06/12/2002 03:43 PM
                    'Set objTransfer = Me.Comm.InitializeOLEServer(NCID_READINESS_APPLICATION)
                    'lRet = objTransfer.Transfer(DMBEObj)
                    'End: SCR #24250;
                
                GoodReturnCode = ImportExport.ExecuteDMBE(DMBEObj, TransferResult)
            End Select
            
          End If
       
      Case Else      '...Get data from Outside World

          '...Get Records from outside world and save in either CDR or LDDB
          If (ImportExport.ExecuteDMBE(DMBEObj)) = 0 Then
              '...0 is a valid return code from Execute

               Data_LDDB = (.Destination = eDM_LDDB)
              
               '...Patient List should have only 1 entry in it _
                   however, if they decide to do multiple patients _
                   THIS CODE WILL NOT WORK _
                     Changes Required _
                        1.  Copy of TransferRecordSets must be Created _
                        2.  IRSObject needs to have a Variable _
                            Indicating the patient associated with _
                            that recordset and Table so that the _
                            routines SaveREcordSetsXXX can load the _
                            correct Recordset associated with a given _
                            Table.
                        
                Patient = .PatientObj.UnitNumber
'               For Each Patient In .PatientList
               
                    '...Update LDDB
                    If mobjShared.LocalDBConnected And Data_LDDB Then
                    
                        '...Save Data to LDDB
                        MergeRecords CStr(Patient), False
                        
                        '...Clear record sets
                        '...In the future need to make a copy _
                            of TransferRecordsets
                        ImportExport.ClearRSCollection
                    
                    End If
                    
                    '...If CDR connected update CDR
                    If mobjShared.CDRConnected And Not NotBothCDRandLDDB Then
                        '...Get Data from CDR
                        BuildRecordSetsCDR CStr(Patient)
                        If mobjShared.LocalDBConnected Then
                          '...Get rquired records in LDDB
                          BuildRecordSetsGEMS CStr(Patient), DMBEObj
                        End If
                        
                        '...Merge Records in Memory
                        MergeRecords CStr(Patient), True
                        
                        '...Update CDR
                        MSaveDataToCDR
                    End If
              
                '...Need to restore TRansferRecordsets Here
                '   If we do multiple patients
               
'               Next Patient
               
               '...Indicate everything is ok
               GoodReturnCode = 0
               
          End If
          
    End Select
getout:
 End With
 
ExitRoutine:

Transfer = GoodReturnCode
 
 Set DataSQL = Nothing
 Set oDMBE = Nothing
 Set ImportExport = Nothing
 
 Exit Function
 
haveerror:

 GoodReturnCode = Err.Number
 
 MsgBxARMd "Error Number:  " & CStr(Err.Number) & vbCrLf & _
        "Description:   " & Err.Description, vbExclamation, "Encounters"
        
 Debug.Print Err.Number, Err.Description
 Resume ExitRoutine
 

End Function


Public Function BuildRecordSetsCDR(UnitNumber As String)
' '
' '     Build recordsets from CDR
' '
'  On Local Error GoTo haveerror
'  Dim atProcedures() As gtProcedure
'  Dim hold As gtProcedure
'  Dim rsProc As ADODB.Recordset
'  Dim lngFieldData As Long
'  Dim Xfield As ADODB.Field
'  Dim ArrayNum As Long
'  Dim xTable As Long
'  Dim outersort As Integer
'  Dim innersort As Integer
'  Dim TestRecordSet As ADODB.Recordset
'  Dim value As Variant
'  Dim Reaction As Integer
'  Dim TableName As String
'  Dim ReturnCode  As Boolean
'
'  '...Set Patient workingon
'  glCurPat = UnitNumber
'
'  '...Retrieve data from CDR
'  'Call mRetrieveCDR(atProcedures)
'
'  '...Are there any records
'  If LBound(atProcedures) = 0 Then Exit Function
'
'  '...Assure data is in chronological order
'  For outersort = LBound(atProcedures) To UBound(atProcedures) - 1
'    For innersort = outersort + 1 To UBound(atProcedures)
'     If CDate(atProcedures(outersort).CreateTime) > CDate(atProcedures(innersort).CreateTime) Then
'       hold = atProcedures(outersort)
'       atProcedures(outersort) = atProcedures(innersort)
'       atProcedures(innersort) = hold
'     End If
'    Next innersort
'  Next outersort
'
'    '...Process all the tables for Procedure
'    For xTable = 1 To 1
'
'      With oDMBE
'       Select Case xTable
'          Case 1          ' Procedures
'            TableName = "Procedures"
'            Set rsProc = New ADODB.Recordset
'            Set TestRecordSet = rsProc
'       End Select
'
'       If .CreateDetachedRS(TableName, TestRecordSet) = 0 Then
'
'            '...Build Record Data Object used by Data Manager Back End
'            ImportExport.SaveCDRrs TestRecordSet, UnitNumber, TableName
'
'       Else
'
'          MsgBxARMd "Error Creating Recordsets for " & TableName, vbExclamation + vbOKOnly, "Processing Terminated"
'
'       End If
'      End With
'
'      '...Convert Data to Universal Recordset
'      '...Process all the Records for this table
'      For ArrayNum = LBound(atProcedures) To UBound(atProcedures)
'
'         '...Create New Record
'         TestRecordSet.AddNew
'
'         '...Map all fields to Record
'         For lngFieldData = 0 To TestRecordSet.Fields.count - 1
'
'            hold = atProcedures(ArrayNum)
'            Set Xfield = TestRecordSet.Fields(lngFieldData)
'
'            With ImportExport
'               Select Case xTable
'
'                  '... Procedures Table
'                  Case 1          ' Procedure
'                           Select Case LCase(Xfield.Name)
'                               Case "facilityncid"
'                                   .ConvertDataToField Xfield, hold.FacilityNCID
'                               Case "dataid"
'                                   .ConvertDataToField Xfield, hold.DataId
'                                Case "procedurencid"
'                                   .ConvertDataToField Xfield, hold.ProcedureNCID
'                               Case "encounternumber"
'                                   .ConvertDataToField Xfield, hold.EncounterNumber
'                               Case "clinicianncid"
'                                   .ConvertDataToField Xfield, hold.ClinicianNCID
'                               Case "unit_number"
'                                   .ConvertDataToField Xfield, hold.UnitNumber
'                               Case "SnoID"
'                                   .ConvertDataToField Xfield, hold.SnoID
'                               Case "pointofcarefacilityncid"
'                                   .ConvertDataToField Xfield, hold.PointOfCareFacilityNCID
'                               Case "statusncid"
'                                   .ConvertDataToField Xfield, hold.StatusNCID
'                               Case "noteddate"
'                                   .ConvertDataToField Xfield, hold.NotedDate
'                               Case "onsetdate"
'                                   .ConvertDataToField Xfield, hold.OnSetDate
'                               Case "problemlistindicator"
'                                   .ConvertDataToField Xfield, hold.ProblemListIndicator
'                               Case "cpt4changeablecode"
'                                   .ConvertDataToField Xfield, hold.Cpt4ChangeableCode
'                               Case "procedurescomment"
'                                   .ConvertDataToField Xfield, hold.ProceduresComment
'                               Case "createuserncid"
'                                   .ConvertDataToField Xfield, hold.CreateUserNCID
'                               Case "createtime"
'                                   .ConvertDataToField Xfield, hold.CreateTime
'                               Case "modifyuserncid"
'                                   .ConvertDataToField Xfield, hold.ModifyUserNCID
'                               Case "modifytime"
'                                   .ConvertDataToField Xfield, hold.ModifyTime
'                               Case "updateflag"
'                                   .ConvertDataToField Xfield, hold.UpdateFlag
'                               Case "cdr_dataid"
'                                   .ConvertDataToField Xfield, hold.CdrDataID
'                               'Case "createdby"
'                               '    .ConvertDataToField Xfield, hold.CreatedBy
'                               'Case "createdon"
'                               '    .ConvertDataToField Xfield, hold.CreatedOn
'                               'Case "updatedby"
'                               '    .ConvertDataToField Xfield, hold.UpdatedBy
'                               'Case "updatedon"
'                               '    .ConvertDataToField Xfield, hold.UpdatedOn
'                           End Select
'               End Select
'            End With
'         Next lngFieldData
'         '...Save Data in Record Set
'         TestRecordSet.Update
'      Next ArrayNum
'
'    Next xTable
'    BuildRecordSetsCDR = True
'ExitRoutine:
'  Exit Function
'haveerror:
'  Debug.Print Err.Number, Err.Description
'  Resume ExitRoutine
End Function

Public Function BuildRecordSetsGEMS(UnitNumber As String, DMBEObj As IDMBE.I_DMBE) As TMIPDataRetrievalResult
' SCR #22581;   Developer: Jane Shen 04/01/2002 05:08 PM
' Change BuildRecordSetsGEMS from sub to function with TMIPDataRetrievalResult as return type

  '    This routine builds recordsets for all the required tables
  
   On Local Error GoTo haveerror
   Dim sql As String
   Dim WherePart As String
   Dim SelectPart As String
   Dim FromPart As String
   Dim FileNumber As Integer
   Dim AdditionalSort As String
   Dim objTable As XferTable
    'Begin:SCR #22819, 22837  Developer: Jane Shen 04/17/2002 03:37 PM
    Dim sSQL As String
    Dim rsEnc As Adodb.Recordset
    Dim sDataIDs As String
    Dim idx As Integer
    Dim encNum As Long
    Dim rsVital As Adodb.Recordset
    'End: SCR #22819, 22837
   Dim rs As Adodb.Recordset
   Dim whoSignedNCID As Long
   Dim rsEncSigBlock As Adodb.Recordset
   Dim OutPutRS As Adodb.Recordset
   
   'SF Get the facility ncid so we will have it when we transfer the ncids table
   Dim sFacilityNCID As String
   
   For Each objTable In mobjTables
   
        '...SQL statement used for all tables
        SelectPart = objTable.SelectPart
        
        If objTable.TableName <> "NCIDS" Then
            WherePart = " WHERE " & objTable.WherePart
            If Len(Trim(objTable.WherePart)) > 0 Then
                WherePart = WherePart & " AND "
            End If
        Else
            'Make an exception for the NCIDS table because we need to add the facility ncid
            WherePart = " WHERE " & objTable.WherePart & sFacilityNCID
        End If
        
        'Begin:SCR #23307;   Developer: Jane Shen 04/24/2002 03:43 PM
        If DMBEObj.Destination = eDM_TMIP Then 'only send completed encounters
        
            If objTable.TableName = "Encounters" Then
                WherePart = WherePart & " (TMIP_STATUS = 0 OR TMIP_STATUS is Null) AND "
                '******************************************************
                'SCR 64993 - Look for Completed and Updated Encounters
                WherePart = WherePart & " (STATUS = 205657 OR STATUS = 14510309) AND"
                '******************************************************
            ElseIf UCase(objTable.TableName) = "DIAGNOSES" Then
                WherePart = WherePart & " STATUSNCID = 1024 AND"
            ElseIf UCase(objTable.TableName) <> "NCIDS" Then
              '*********************************************************
                'SCR 67140 - TH-IP2p1: DM - JMeWS file sent for an updated encounter appears incomplete
               WherePart = WherePart & " (enc.TMIP_STATUS = 0 OR enc.TMIP_STATUS is Null) AND "
                 WherePart = WherePart & " (enc.STATUS = 205657 OR enc.STATUS = 14510309) AND"
              '*********************************************************
            End If
            
            If UCase(objTable.TableName) <> "NCIDS" Then
                WherePart = WherePart & objTable.WherePatientPart & UnitNumber
            End If
            
        ElseIf DMBEObj.Destination = eDM_EIC Then 'only send last completed encounter
        
            If objTable.TableName = "Encounters" Then
                'Get last completed or updated encounter SCR-25818
                sSQL = "select  max(encounternumber) from encounters where STATUS in (205657, 14510309) AND "
                sSQL = sSQL & " TYPE <> 5 AND " 'SCR 76426 Don't export inpatient encounters.
                sSQL = sSQL & objTable.WherePatientPart & UnitNumber
                
                Set rsEnc = DataSQL.OpenRecordset(sSQL)
                'Begin:SCR #  Developer: Jane Shen 05/01/2002 05:53 PM
                If IsNull(rsEnc.Fields.Item(0).value) Then
                    encNum = 0
                Else
                    encNum = rsEnc.Fields.Item(0).value
                End If
                'End: SCR #
                WherePart = WherePart & " encounternumber =  " & encNum
                Set rsEnc = Nothing
            ElseIf UCase(objTable.TableName) <> "NCIDS" Then
                WherePart = WherePart & " enc.encounternumber =  " & encNum
            End If
          
        Else
            If objTable.TableName = "Encounters" Then  'SCR#36893, Complete=205657, Updated=14510309
                WherePart = WherePart & " STATUS in (205657, 14510309) AND "
            ElseIf UCase(objTable.TableName) <> "NCIDS" Then
                 WherePart = WherePart & " enc.STATUS in (205657, 14510309) AND "
            End If
            
            If UCase(objTable.TableName) <> "NCIDS" Then
                WherePart = WherePart & objTable.WherePatientPart & UnitNumber
            End If
        End If
        'End: SCR #23307;
        
    'SF SCR 62831
    '    'Exclude Inpatient Note encounters from being transferred SCR#47895
    '    'Added "AND" SCR#48202
    '    If objTable.TableName = "Encounters" Then
    '        WherePart = WherePart & " AND TYPE <> 5 "
    '    Else
    '       WherePart = WherePart & " AND enc.TYPE <> 5 "
    '    End If
        
         AdditionalSort = objTable.AdditionalSort
         
         FromPart = objTable.TableName
         
         '...Get all data from this file for this patient
         sql = SelectPart & objTable.FromPart & WherePart
         If Len(Trim(AdditionalSort)) > 0 Then
            sql = sql & " Order By " & AdditionalSort & " "
         End If
         
         Set rs = DataSQL.OpenRecordset(sql)
        
        'Handel case of vitals
        ' SCR #47595;   Developer:  01/08/2004
        'Begin:SCR #22819, 22837   Developer: Jane Shen 04/17/2002 03:38 PM
        If objTable.TableName = "Vitals" And Not rs.EOF Then
            idx = 0
            Do Until rs.EOF
                If idx = 0 Then
                  sDataIDs = rs.Fields("DATAID").value
                Else
                  sDataIDs = sDataIDs & ", " & rs.Fields("DATAID").value
                End If
                
                idx = idx + 1
                rs.MoveNext
            Loop 'Until rs.EOF
          
            sSQL = "select * from VITALSignsValue where DATAID in ( " & sDataIDs & ")"
            Set rsVital = DataSQL.OpenRecordset(sSQL)
            ImportExport.SaveGEMSrs rsVital, UnitNumber, "VITALSignsValue"
            Set rsVital = Nothing
            
            sSQL = "select * from VITALSignsModifier where DATAID in ( " & sDataIDs & ")"
            Set rsVital = DataSQL.OpenRecordset(sSQL)
            ImportExport.SaveGEMSrs rsVital, UnitNumber, "VITALSignsModifier"
            Set rsVital = Nothing
        End If
        
        'Pass Enc Sig Block SCR-22826
        If Not rs.EOF And objTable.TableName = "Encounters" And DMBEObj.Destination = eDM_EIC Then
            'Get who signed
            ' Begin:SCR #30967;   Developer: Mag 12/18/2002 06:24 PM
            If IsNull(rs.Fields("WHOSIGNEDNCID")) Then
                whoSignedNCID = 0
            Else
                whoSignedNCID = rs.Fields("WHOSIGNEDNCID")
            End If
            ' End: SCR #30967;
            
            'Get provider's sig block
            sSQL = "SELECT * FROM ENC_SIGBLOCK WHERE USERNCID = " & whoSignedNCID
            
            Set rsEncSigBlock = DataSQL.OpenRecordset(sSQL)
            
            ImportExport.SaveGEMSrs rsEncSigBlock, UnitNumber, "Enc_SigBlock"
            
            Set rsEncSigBlock = Nothing
        End If
         
        If Not rs.EOF And objTable.TableName = "Encounters" Then
            'SF Get the Facility NCID of this facility to retrieve the facility info from the ncids table.
            sFacilityNCID = rs("FacilityNCID")
        End If
        
        If rs.EOF And objTable.TableName = "Encounters" Then
           BuildRecordSetsGEMS = No_Data
           'Begin:SCR #23307;   Developer: Jane Shen 04/24/2002 03:43 PM
           If DMBEObj.Destination = eDM_TMIP Then
               sql = "select * from Encounters where"
               sql = sql & " TMIP_STATUS <> 0 AND STATUS = 205657 AND "
               sql = sql & " unit_number = " & UnitNumber
               Set rs = DataSQL.OpenRecordset(sql)
               If Not rs.EOF Then
                   BuildRecordSetsGEMS = All_Data_Have_Been_Sent
               End If
            End If
           'End: SCR #23307;
           
           If Not rs.EOF Then
             ImportExport.SaveGEMSrs rs, UnitNumber, FromPart
           End If
           
           Set rs = Nothing
           Exit Function
           
        Else
            '<Fixed in build cube 4/26/2002>
           'If rs.EOF Then
           '    BuildRecordSetsGEMS = No_Data
           'Else
           '    BuildRecordSetsGEMS = Success
           'End If
           '</Fixed in build cube 4/26/2002>
           '...Build Record Data Object used by Data Manager Back End
           ImportExport.SaveGEMSrs rs, UnitNumber, FromPart
        End If
        
        'Transfer Symptoms SCR-26545 also do for EIC SCR-28327
        If DMBEObj.Source = eDM_LDDB And (DMBEObj.Destination = eDM_TMIP Or DMBEObj.Destination = eDM_EIC) And objTable.TableName = "Encounters" Then
            rs.MoveFirst
               
            Do Until rs.EOF
    '            Set colTMIPEncNum = New Collection
    '            colTMIPEncNum.Add "" & rs.Fields("ENCOUNTERNUMBER")
                TransferSymptoms oDMBE, rs.Fields("ENCOUNTERNUMBER"), OutPutRS
                rs.MoveNext
            Loop
            
            ImportExport.SaveGEMSrs OutPutRS, UnitNumber, "Symptoms"
            
            Set OutPutRS = Nothing
        End If
        
        'End: SCR #22581;
        
        Set rs = Nothing
     
   Next objTable

   Exit Function
haveerror:
    MsgBxARMd "Error Number:  " & CStr(Err.Number) & vbCrLf & _
            "Description:   " & Err.Description, vbExclamation, "Encounters"
    BuildRecordSetsGEMS = Application_Error
    Debug.Print Err.Number, Err.Description
    Exit Function
    Resume
End Function

Public Sub MergeRecords(UnitNumber As String, MemoryMerge As Boolean)

  ' This routine Determines the Tables to be moved.
  ' It then extracts the necessary tables from the
  ' collection supplied by TransferRecordSets.  These tables are
  ' extracted in a specific order.  The order is
  ' determined by which tables generate keys that
  ' are used by other tables in this move.  The tables
  ' generating these keys are extracted first.
  
  ' The Key fields are defined for the table and mapping
  ' is performed on the input recordset to optimize
  ' indexing.
  
  ' Business decisions are then implemented to determine
  ' if a record is added, deleted or updated to the
  ' destination table.

  On Local Error GoTo haveerror
  
  Dim CREATEDONFIELDFROM As Adodb.Field          ' Field Mapped to RS recordset for field CreatedON
  Dim CREATEDONFIELDTO As Adodb.Field            ' Field Mapped to TORS recordset for field CreatedON
  Dim UPDATEDONFIELDFROM As Adodb.Field          ' Field Mapped to RS recordset for field UpdatedON
  Dim UPDATEDONFIELDTO As Adodb.Field            ' Field Mapped to TORS recordset for field UPDATEDONON
  'Dim objCreateTimeFrom As ADODB.Field    'Field mapped to FromRs recordset for field Last_Readiness_Dt
  'Dim objCreateTimeTo As ADODB.Field      'Field mapped to ToRs recordset for field Last_Readiness_Dt
  Dim count As Integer
  Dim GetNextToRecord As Boolean                 ' Indicates move to next record in To Table
  Dim GetNextFromRecord As Boolean               ' Indicates move to next record in From Table
  Dim GoodIO As Boolean                          ' Indicate insertion or update was successful
  Dim Recordsets As Integer                      ' Loops through all record Sets in Collection
  Dim dataDAS As ICHCSII_DAS
  Dim FieldNames() As Variant
  Dim SequenceSource() As Variant
  Dim SequenceType() As Variant
  Dim DictionaryKey() As Variant
  Dim DictionarySynonym() As Variant
  Dim ThisIsAKeyTable As Boolean                 ' Indicates that Processing should be terminated if _
                                                   value is True and Table in TransferRecordSets is not found
  Dim objTable As XferTable
  Dim strTableName As String
  
   Set dataDAS = gobjCHCSIIConn.CHCSII_DAS(Auto)
   With ImportExport
    '...Loop through all recordsets supplied by DMBE
    'For Recordsets = 1 To .RecordSetsTransfered
    For Each objTable In mobjTables
    
      '...Indicate this table not necessary
      ThisIsAKeyTable = False
      
      '...Initialize all arrays
      .EraseArrays
      
      '...Reset Name of Key used to insert a new record
      .ADDKEY = ""
      
      '...Reset name of Field to indicate when record was created
      .CreatedOn = ""
       
      '...Determine next Table to work on
      '...Establish WHERE Fields that are used when doing an update to that table
      '...Establish Key Field names that need new sequential numbers when records _
          are inserted into that Table
'      Select Case Recordsets
'
'          Case 1
'             '...Table being moved
'             '.FromPart = "procedures"
'             strTableName = "Encounters"
'          Case 2
'            strTableName = "Encounter_Lock"
'      End Select
'
'      Set objTable = mobjTables.Item(strTableName)
      
      
      
      
      
      '.FromPart = objTable.TableName
      .FromPart = objTable.FromPart
             
      '...Table must exist and there has to be entries in Table
      ThisIsAKeyTable = True
    
      '...All keys Created (Generated) when a new record is added
      'FieldNames = Array("DATAID")
      'FieldNames = Array("dataid")
      FieldNames = Array("")
    
      '...Source to create the above Keys 1 to 1 relationship
      'SequenceSource = Array(SequentialIDSource.DataId)
      'SequenceSource = Array("SequentialIDSource.DataId")
      SequenceSource = Array("")
    
      '...Type of Sequential ID (used by Source)
      '   NOT REQUIRED IF SOURCE IS NCIDS but a place holder is required
      '   for multiple keys
      'SequenceType = Array(ICHCSII_SQL.eSeq.esMAINSEQ)
      SequenceType = Array("")
    
      '...Synonym name for DATAID used in other tables where DATAID is a foreign key
      '   The synonym is required where the key representing DATAID has a different name
      '   Any name can be used...but the name must be consistent when mapping the
      '   foreing keys in other tables
      'DictionaryKey = Array("AllergiesDATAID")
      DictionaryKey = Array("")
    
      '...Synonym name for the Fieldnames key in this table
      '   "" means this table has no foreign keys THAT WE NEED TO WORRY ABOUT in the FUTURE
      '   If a foreign key has been previuosly Created or Used then the DictionaryKey used to define that
      '   mapping to that Field must be entered here.  This entry must map to FieldNames defined above
      'DictionarySynonym = Array("")
      DictionarySynonym = Array("")
    
      '...Set Sequential KEY information for ImportExport
    
      'emc readiness doesn't have DATAID fields, so skip this step ?
      '.SetNewSequential FieldNames, SequenceSource, SequenceType, DictionaryKey, DictionarySynonym
    
      '...Define Field Names used in where clause to Extract records from Source Table
      '   Besides the unique patient ID Field name include all keys that are generated by an insert
      '   to this table
      'FieldNames = Array("DATAID", "UNIT_NUMBER")
      FieldNames = Array("")
      Dim lngIndex As Long
      If objTable.PatientUniqueFields.count > 0 Then
        ReDim Preserve FieldNames(objTable.PatientUniqueFields.count)
        
        'For lngIndex = 1 To objTable.PatientUniqueFields.count
        '  FieldNames(lngIndex) = objTable.PatientUniqueFields.Key(lngIndex)
        'Next lngIndex
        FieldNames = objTable.PatientUniqueFields.Keys
      End If

      '...Set Where Criteria
      .SetWhereCondition FieldNames
    
      '...Include all keys required to sort records in proper order
      '   Most of the time this will be a sequence number
      '   "" means records will be sorted by date created
      '.ADDKEY = ", CREATETIME"
      .ADDKEY = ""
    
      '...Define name of field used to indicate when record was created
      .CreatedOn = "CREATEDON"
      .UpdatedOn = "UPDATEDON"
      'End Select
      
      '...THIS CODE MUST APPEAR AFTER THE TABLES AND KEYS ARE DEFINED ABOVE
      
      '...Find Proper Table in Collection
      '...and Load imported Records
      'Set .FromRs = .GetATable(.FromPart)
      Set .FromRs = .GetATable(objTable.TableName)
      If Not .FromRs Is Nothing Then
        
        If MemoryMerge Then
          '...Get all pertinent records in CDR Table that we are moving to
          Set .ToRs = .GetCDRrs(UnitNumber, .FromPart)
          
        Else
          '...Get all pertinent records in LDDB Table that we are moving to
          Dim strSelectSQL As String
          Dim strWhereSQL As String
          'Dim strUniqueFields() As Variant
          'Dim strField As Variant
          
          strSelectSQL = objTable.SelectPart
          strWhereSQL = objTable.WherePart
          'strUniqueFields = objTable.PatientUniqueFields.Keys
          lngIndex = 0
          
'          For Each strField In strUniqueFields
'            If Len(Trim(strWhereSQL)) > 0 Then
'                strWhereSQL = strWhereSQL & " AND "
'            End If
'            Select Case LCase(strField)
'                Case "unitnumber"
'                    strWhereSQL = strWhereSQL & " unitnumber = " & UnitNumber
'                Case "unit_number"
'                    strWhereSQL = strWhereSQL & " unit_number = " & UnitNumber
'            End Select
'            lngIndex = lngIndex + 1
'          Next strField
'         strWhereSQL = " WHERE " & strWhereSQL
            If Len(Trim(strWhereSQL)) > 0 Then
                strWhereSQL = strWhereSQL & " AND "
            End If
            strWhereSQL = strWhereSQL & objTable.WherePatientPart & UnitNumber
            strWhereSQL = " WHERE " & strWhereSQL


          
          'Set .ToRs = dataDAS.ICHCSII_DAS_ExecuteSQL(strSelectSQL & .FromPart & strWhereSQL & " Order By " & .CreatedOn & .ADDKEY & " ASC")
          Set .ToRs = dataDAS.ExecuteSQL(strSelectSQL & .FromPart & strWhereSQL)
          
        End If
        
        '...Map all fields to Proper Record Sets
        .MapFields
        
        '...Map CreatedON Fields
        'Set CREATEDONFIELDFROM = .FindFieldInRS(.FromRs, .CreatedOn)
        'Set CREATEDONFIELDTO = .FindFieldInRS(.ToRs, .CreatedOn)
        
        'Map Last_Readiness_Dt fields
        'Set objCreateTimeFrom = .FindFieldInRS(.FromRs, "CreateTime")
        'Set objCreateTimeTo = .FindFieldInRS(.ToRs, "CreateTime")
        
        
        '...Map UpdatedON Fields
        'Set UPDATEDONFIELDFROM = .FindFieldInRS(.FromRs, .UpdatedOn)
        'Set UPDATEDONFIELDTO = .FindFieldInRS(.ToRs, .UpdatedOn)
        
          If Not .FromRs.EOF Then
           .FromRs.MoveFirst
          Else
           '... Need to see if table exists here
           '    Build the table if it does not
          
          End If
            
          'emc set data structure indicating which records have been processed
          'Dim objProcessedFrom As Dictionary
          'Dim objProcessedTo As Dictionary
          Dim dtTo As Date
          Dim dtFrom As Date
          
'          .FromRs.MoveFirst
'          Do While Not .FromRs.EOF
'            dtFrom = .FromRs.Fields("last_readiness_dt")
'            objProcessedFrom.Add CStr(dtFrom), vbFalse
'            .FromRs.MoveNext
'          Loop
'          .FromRs.MoveFirst
          
          .ToRs.MoveFirst
'          Do While Not .ToRs.EOF
'            dtTo = .ToRs.Fields("last_readiness_dt")
'            objProcessedTo.Add CStr(dtTo), vbFalse
'            .ToRs.MoveNext
'          Loop
'          .ToRs.MoveFirst
          
        Dim rsMerge As Recordset
        Dim objField As XferField
        Set rsMerge = New Recordset
        rsMerge.CursorLocation = adUseClient
        rsMerge.CursorType = adOpenStatic
        rsMerge.LockType = adLockPessimistic
        rsMerge.Source = .ToRs.Source
        rsMerge.ActiveConnection = .ToRs.ActiveConnection
        rsMerge.Open
          
            '...Loop through all imported records
            Do While Not .FromRs.EOF
                .ToRs.MoveFirst
                
                Dim strFind As String
                Dim strDelete As String
                Dim strKeyFields() As Variant
                Dim strKeyField As Variant
                
                strKeyFields = objTable.KeyFields.Keys
                lngIndex = 0
                strFind = ""
                'strDelete = ""
                For Each strKeyField In strKeyFields
                    
                    If lngIndex > 0 Then
                        strFind = strFind & " and "
                        strDelete = strDelete & " AND "
                    End If
                    
                    Select Case LCase(strKeyField)
                        Case "unitnumber"
                            strFind = strFind & " unitnumber = " & UnitNumber
                            'strDelete = strDelete & " unitnumber = " & UnitNumber
                        Case "unit_number"
                            strFind = strFind & " unit_number = " & UnitNumber
                            'strDelete = strDelete & " unit_number = " & UnitNumber
                        'Case "encounternumber"
                        '    strFind = strFind & " encounternumber = " & .FromRs("encounternumber")
                           ' strDelete = strDelete & " encounternumber = " & .FromRs("encounternumber")
                        Case Else
                            Select Case .FromRs.Fields(strKeyField).Type
                                Case adDate, adDBTimeStamp, adDBTime
                                    strFind = strFind & " " & strKeyField & " = #" & .FromRs(strKeyField) & "# "
                                Case adVarChar
                                    strFind = strFind & " " & strKeyField & " = '" & .FromRs(strKeyField) & "' "
                                Case Else
                                    strFind = strFind & " " & strKeyField & " = " & .FromRs(strKeyField)
                            End Select
                    End Select
                    
                    lngIndex = lngIndex + 1
                    DoEvents
                Next strKeyField
                
                .ToRs.Filter = strFind
                
                If Not .ToRs.EOF Then  'matching record found, update
                  rsMerge.MoveFirst
                  rsMerge.Filter = strFind
                Else  'record not found, insert it
                    rsMerge.AddNew
                End If
                  
                For Each objField In objTable.Fields
                  rsMerge.Fields(objField.FieldName) = .FromRs.Fields(objField.FieldName)
                  DoEvents
                Next objField
                rsMerge.Update
                
                
                .FromRs.MoveNext
            Loop
            
            'Now insert all the "from" records
            .FromRs.MoveFirst
            .ActionToTake = "INSERT"
            
            Do While Not .FromRs.EOF
                If MemoryMerge Then
                   '...Perform Transaction
                   MergeToMemory UnitNumber, .FromPart, .FromRs, .ToRs
                Else
                   '...Perform Transaction
                   'GoodIO = .MoveRecordToGems(.FromRs)
                End If
                .FromRs.MoveNext
                DoEvents
            Loop
            
      Else                  ' Table not found in collection
        
        '...Test to see if this table was required to continue with other tables
        If LCase(Trim(objTable.TableName)) = "encounters" Then
            Exit For
        End If
        'If ThisIsAKeyTable Then Exit For

        
      End If
          
    'Next Recordsets
    Next objTable
    
   End With                 ' end of ImportExport
   Set dataDAS = Nothing
   Exit Sub
haveerror:
    MsgBxARMd "Error Number:  " & CStr(Err.Number) & vbCrLf & _
            "Description:   " & Err.Description, vbExclamation, "Encounters"
            
     
     Debug.Print Err.Number, Err.Description
     
End Sub

Public Sub MergeToMemory(UnitNumber As String, xTableName As String, FromRs As Adodb.Recordset, ToRs As Adodb.Recordset)
   
   '    Updated always takes data from ToRs and Updates it with Data from  FromRs _
        Works like an Insert except there is no New Keys Created Unless the CDR _
        key is Different from the LDDB Key  (Records from LDDB)
   
   '    Inserted data always comes from FromRs (Records from LDDB) _
        Sequential Numbers are set to 0
   
   '    Copy Data Always Comes from FromRs (Records from CDR) _
        Sequential Numbers from CDR are used
   
   Dim Xfield As Adodb.Field
   'Dim hold As gtProcedure
   Dim hold As Dictionary
   Dim ArrayNumber As Long
   Dim objTable As XferTable
   Dim objField As XferField
   
   Set objTable = mobjTables.Item(xTableName)
'   Select Case xTableName
'      Case "Procedures"
'
'        '...Increase Number of Records in Array
'        NumAllRecordsInMemory = NumAllRecordsInMemory + 1
'        If NumAllRecordsInMemory = 1 Then
'         ReDim AllRecordsInMemory(1 To 1)
'        Else
'         ReDim Preserve AllRecordsInMemory(NumAllRecordsInMemory)
'        End If
'        ArrayNumber = NumAllRecordsInMemory
'
'        Select Case ImportExport.ActionToTake
'
'           Case "Copy"
'           Case "Insert"
'           Case "Update"
'        End Select
'   End Select
   
        '...Load Proper Array
        'hold = AllRecordsInMemory(ArrayNumber)
        Set hold = New Dictionary
        AllRecordsInMemory.Add objTable.TableName, hold
         
        '...Map all fields to Record
        'For Each Xfield In FromRs.Fields
        For Each objField In objTable.Fields
    
             'With ImportExport
             hold.Add objField.FieldName, "" & objField.value
             
'                Select Case LCase(Xfield.Name)
'                    Case "facilityncid"
'                        hold.FacilityNCID = "" & Xfield.value
'                    Case "dataid"
'                        hold.DataId = "" & Xfield.value
'                    Case "procedurencid"
'                        hold.ProcedureNCID = "" & Xfield.value
'                    Case "encounternumber"
'                        hold.EncounterNumber = "" & Xfield.value
'                    Case "clinicianncid"
'                        hold.ClinicianNCID = "" & Xfield.value
'                    Case "unit_number"
'                        hold.UnitNumber = "" & Xfield.value
'                    Case "SnoID"
'                        hold.SnoID = "" & Xfield.value
'                    Case "pointofcarefacilityncid"
'                        hold.PointOfCareFacilityNCID = "" & Xfield.value
'                    Case "statusncid"
'                        hold.StatusNCID = "" & Xfield.value
'                    Case "noteddate"
'                        hold.NotedDate = "" & Xfield.value
'                    Case "onsetdate"
'                        hold.OnSetDate = "" & Xfield.value
'                    Case "problemlistindicator"
'                        hold.ProblemListIndicator = "" & Xfield.value
'                    Case "cpt4changeablecode"
'                        hold.Cpt4ChangeableCode = "" & Xfield.value
'                    Case "procedurescomment"
'                        hold.ProceduresComment = "" & Xfield.value
'                    Case "createuserncid"
'                        hold.CreateUserNCID = "" & Xfield.value
'                    Case "creattime"
'                        hold.CreateTime = "" & Xfield.value
'                    Case "modifyuserncid"
'                        hold.ModifyUserNCID = "" & Xfield.value
'                    Case "modifytime"
'                        hold.ModifyTime = "" & Xfield.value
'                    Case "updateflag"
'                        hold.UpdateFlag = "" & Xfield.value
'                    Case "cdr_dataid"
'                        hold.CdrDataID = "" & Xfield.value
'                    'Case "createdby"
'                    '   hold.CreatedBy = "" & Xfield.Value
'                    'Case "createdon"
'                    '   hold.CreatedOn = "" & Xfield.Value
'                    'Case "updatedby"
'                    '   hold.UpdatedBy = "" & Xfield.Value
'                    'Case "updatedon"
'                    '   hold.UpdatedOn = "" & Xfield.Value
'               End Select
            'End With
        
        'Next Xfield
        Next objField
End Sub


Public Sub MSaveDataToCDR()
 
  ' Saves Data in Memory to CDR
  
  Dim count As Long
  
   For count = 1 To NumAllRecordsInMemory
        'mSaveCDR AllRecordsInMemory(count)
   Next count

End Sub





Private Function LoadEncounterTables(objTables As XferTables) As Boolean

Dim blnRet As Boolean
On Error GoTo ErrHandler

blnRet = True

'Load Encounters Table
blnRet = LoadEncounters
'If blnRet Then
'    blnRet = LoadEncounters_Audit
'End If
'If blnRet Then
'    blnRet = LoadEncounter_Lock
'End If
If blnRet Then
    blnRet = LoadEncounter_Sections
End If
If blnRet Then
    blnRet = LoadEncounter_Rtfs
End If
If blnRet Then
    blnRet = Load_SoNotes
End If
If blnRet Then
    blnRet = Load_ListToolRecord
End If
If blnRet Then
    blnRet = Load_ListToolRecords
End If
If blnRet Then
    blnRet = LoadEncounter_Diagnosis
End If
If blnRet Then
    blnRet = LoadEncounter_Disposition
End If
If blnRet Then
    blnRet = LoadEncounter_TextNotes
End If
If blnRet Then
    blnRet = Load_PiData
End If
If blnRet Then
    blnRet = Load_Vitals
End If
If blnRet Then
    blnRet = Load_ClinNote
End If
If blnRet Then
    blnRet = Load_WorkStatus
End If
'**************************************************************
'SCR 67372
'Datamanager needs to export the OE_RX table for PVF and JMEWS.
'**************************************************************
If blnRet Then
    blnRet = Load_OE_RX
End If

'If blnRet Then
'    blnRet = Load_EncSigBlock
'End If

'*************
'SF Added inpatient tables
If blnRet Then
    blnRet = Load_Inpatient_Admission
End If
If blnRet Then
    blnRet = Load_InptWard
End If
If blnRet Then
    blnRet = Load_InptWard_Location
End If
If blnRet Then
    blnRet = Load_Inpatient_Transfer
End If
If blnRet Then
    blnRet = Load_Inpatient_Diagnosis
End If

'SF 76357 8-25-05
If blnRet Then
    blnRet = Load_Inpatient_Discharge
End If
'*************

'Add the NCIDS table
If blnRet Then
    blnRet = Load_NCIDS_Table ' Load_NICDS_Table
End If

LoadEncounterTables = blnRet

Exit Function
ErrHandler:
Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Function



Private Function LoadEncounters_Audit() As Boolean
Dim objTable As XferTable
Dim objField As XferField

Set objTable = New XferTable

With objTable
    .TableName = "Encounters_Audit"
    .SelectPart = "Select * "
    .FromPart = " FROM Encounters_Audit ea "
    .WherePart = " "
    .WherePatientPart = " Unit_Number = "
    
    .AdditionalSort = " CreatedOn ASC, UpDatedOn ASC "
    .PatientUniqueFields.Add "UnitNumber", "UnitNumber"
    .KeyFields.Add "EncounterNumber", "EncounterNumber"
    
    .AddField "TemplateId"
    .AddField "UnitNumber"
    .AddField "PrimaryProviderNCID"
    .AddField "EandmNCID"
    .AddField "FacilityNCID"
    .AddField "EncounterNumber"
    .AddField "ClinicNCID"
    .AddField "Status"
    .AddField "Unit_Number"
    .AddField "Category"
    .AddField "Class"
    .AddField "Type"
    .AddField "EndDts"
    .AddField "PrimaryDiagnosisNCID"
    .AddField "ApptId"
    .AddField "StartDts"
    .AddField "SadrStatus"
    .AddField "DispositionText"
    .AddField "DispositionNCID"
    .AddField "DispAdminOption"
    .AddField "ReviewedByCoder"
    .AddField "EandmReviewed"
    .AddField "EandmReviewedBy"
    .AddField "EandmReviewedDate"
    .AddField "EandmCalc"
    .AddField "EandmCodeReview"
    .AddField "TemplateData"
    .AddField "TimeZone"
    .AddField "WsDisp"
    .AddField "WsRfv"
    .AddField "ApRole1NCID"
    .AddField "ApRole2NCID"
    .AddField "WasupDated"
    .AddField "WhenCosigned"
    .AddField "CosignerLine2"
    .AddField "CosignerLine3"
    .AddField "WhenSigned"
    .AddField "SignerLine2"
    .AddField "SignerLine3"
    .AddField "NoCount"
    .AddField "AdditionalProvider1NCID"
    .AddField "ApptType"
    .AddField "AdditionalProvider2NCID"
    .AddField "AllergyVerifiedByDate"
    .AddField "AllergyVerifiedByNCID"
    .AddField "WhoCosignedNCID"
    .AddField "FollupComments"
    .AddField "WhoSignedNCID"
    .AddField "FollowupTimeFrame"
    .AddField "CosignerNCID"
    .AddField "ItemsDiscussed"
    .AddField "CreatedBy"
    .AddField "CreatedOn"
    .AddField "UpdatedBy"
    .AddField "UpdatedOn"
End With

mobjTables.Add objTable

LoadEncounters_Audit = True

End Function

Private Function LoadEncounters() As Boolean
Dim objTable As XferTable
Dim objField As XferField

Set objTable = New XferTable

With objTable
    .TableName = "Encounters"
    .SelectPart = "Select * "
    .FromPart = " FROM Encounters "
    '.WherePart = " WHERE UnitNumber = "
    .WherePart = " "
    .WherePatientPart = " Unit_Number = "
    .AdditionalSort = " CreatedOn ASC, UpDatedOn ASC "
    .PatientUniqueFields.Add "Unit_Number", "Unit_Number"
    .KeyFields.Add "EncounterNumber", "EncounterNumber"
        
    .AddField "TemplateId"
    .AddField "PrimaryProviderNCID"
    .AddField "FacilityNCID"
    .AddField "EncounterNumber"
    .AddField "ClinicNCID"
    .AddField "Status"
    .AddField "Unit_Number"
    .AddField "Category"
    .AddField "Class"
    .AddField "Type"
    .AddField "EndDts"
    .AddField "PrimaryDiagnosisNCID"
    .AddField "ApptId"
    .AddField "StartDts"
    .AddField "SadrStatus"
    .AddField "TemplateData"
    .AddField "TimeZone"
    .AddField "ApRole1NCID"
    .AddField "ApRole2NCID"
    .AddField "WasupDated"
    .AddField "WhenCosigned"
    .AddField "CosignerLine2"
    .AddField "CosignerLine3"
    .AddField "WhenSigned"
    .AddField "SignerLine2"
    .AddField "SignerLine3"
    .AddField "NoCount"
    .AddField "AdditionalProvider1NCID"
    .AddField "ApptType"
    .AddField "AdditionalProvider2NCID"
    .AddField "AllergyVerifiedByDate"
    .AddField "AllergyVerifiedByNCID"
    .AddField "WhoCosignedNCID"
    .AddField "WhoSignedNCID"
    .AddField "CosignerNCID"
    .AddField "CreatedBy"
    .AddField "CreatedOn"
    .AddField "UpdatedBy"
    .AddField "UpdatedOn"
    .AddField "Sensitive"
    
    'SF Added inpatient fields.
    .AddField "Inpatient_Service"
    .AddField "INPT_NOTE_TYPE_NCID"
    .AddField "INPT_HOSP_DAY"
    
End With

mobjTables.Add objTable

LoadEncounters = True

End Function


Private Function LoadEncounter_Lock() As Boolean
Dim objTable As XferTable
Dim objField As XferField

Set objTable = New XferTable

With objTable
    .TableName = "Encounter_Lock"
    '.SelectPart = "Select enc.unit_number, enc.encounternumber, el.* "
    .SelectPart = "Select el.* "
    .FromPart = " FROM Encounters enc, Encounter_Lock el "
    .WherePart = " enc.encounternumber = el.encounter_number "
    .WherePatientPart = " enc.Unit_Number = "
    .AdditionalSort = ""
    
    .PatientUniqueFields.Add "Unit_Number", "Unit_Number"
    .KeyFields.Add "Encounter_Number", "Encounter_Number"
    .KeyFields.Add "Facility_NCID", "Facility_NCID"
    .KeyFields.Add "Section_Number", "Section_Number"
    .KeyFields.Add "Event_Id", "Event_Id"
        
    .AddField "FACILITY_NCID"
    .AddField "ENCOUNTER_NUMBER"
    .AddField "SECTION_NUMBER"
    .AddField "EVENT_ID"
    .AddField "USER_NCID"
    .AddField "WORKSTATION"
    .AddField "DATE_TIME"
    .AddField "REASON"
    .AddField "STATUS"
    .AddField "ITEM_NAME"
End With

mobjTables.Add objTable

LoadEncounter_Lock = True

End Function

Private Function LoadEncounter_Sections() As Boolean
Dim objTable As XferTable
Dim objField As XferField

Set objTable = New XferTable

With objTable
    .TableName = "Enc_Sections"
    .SelectPart = "Select es.* "
    .FromPart = " FROM Encounters enc, Enc_Sections es "
    .WherePart = " enc.encounternumber = es.encounternumber "
    .WherePatientPart = " enc.Unit_Number = "
    .AdditionalSort = ""
    
    .PatientUniqueFields.Add "Unit_Number", "Unit_Number"
    .KeyFields.Add "DataID", "DataID"
        
    .AddField "DATAID"
    .AddField "FACILITYNCID"
    .AddField "ENC_SECTIONSINDEX"
    .AddField "ENCOUNTERNUMBER"
    .AddField "DTS"
    .AddField "DOC"
    .AddField "OWNERNCID"
    .AddField "OWNERNAME"
    .AddField "SENSITIVITYLEVEL"
    .AddField "STATUS"
    .AddField "ORIGINALSIZE"
    .AddField "LABRTF"
    .AddField "RADRTF"
    .AddField "CREATEDBY"
    .AddField "CREATEDON"
    .AddField "UPDATEDBY"
    .AddField "UPDATEDON"
    .AddField "COMPLETE"
    .AddField "CATEGORY"
    .AddField "TITLE"
    .AddField "WSRFV"
    .AddField "DATATYPENCID"
    .AddField "UNIT_NUMBER"
    
End With

mobjTables.Add objTable

LoadEncounter_Sections = True

End Function

Private Function LoadEncounter_Rtfs() As Boolean
Dim objTable As XferTable
Dim objField As XferField

Set objTable = New XferTable

With objTable
    .TableName = "Enc_Rtfs"
    .SelectPart = "Select er.* "
    .FromPart = " FROM Encounters enc, Enc_Rtfs er "
    .WherePart = " enc.encounternumber = er.encounternumber "
    .WherePatientPart = " enc.Unit_Number = "
    .AdditionalSort = ""
    
    .PatientUniqueFields.Add "Unit_Number", "Unit_Number"
    .KeyFields.Add "FacilityNCID", "FacilityNCID"
    .KeyFields.Add "EncounterNumber", "EncounterNumber"
    .KeyFields.Add "Dts", "Dts"
    .KeyFields.Add "Enc_RtfsIndex", "Enc_RtfsIndex"

    .AddField "FACILITYNCID"
    .AddField "ENCOUNTERNUMBER"
    .AddField "DTS"
    .AddField "RTFTYPE"
    .AddField "ENC_RTFSINDEX"
    .AddField "NOTEID"
    .AddField "SOURCE"
    .AddField "USERNCID"
    .AddField "USERNAME"
    .AddField "DATEADDED"
    .AddField "DOC"
    .AddField "TITLE"
    .AddField "STATUS"
    .AddField "SENSITIVITYLEVEL"
    .AddField "COSIGNERNCID"
    .AddField "ORIGINALSIZE"
    .AddField "INITIALSIGNATURE"
    .AddField "FINALSIGNATURE"
    .AddField "COMPLETE"
    .AddField "CATEGORY"
    .AddField "UPDATEFLAG"
    .AddField "CREATEDBY"
    .AddField "CREATEDON"
    .AddField "UPDATEDBY"
    .AddField "UPDATEDON"
End With

mobjTables.Add objTable

LoadEncounter_Rtfs = True

End Function

Private Function Load_SoNotes() As Boolean
Dim objTable As XferTable
Dim objField As XferField

Set objTable = New XferTable

With objTable
    .TableName = "SoNotes"
    .SelectPart = "Select sn.* "
    .FromPart = " FROM Encounters enc, SoNotes sn "
    .WherePart = " enc.encounternumber = sn.encounternumber "
    .WherePatientPart = " enc.Unit_Number = "
    .AdditionalSort = ""
    
    .PatientUniqueFields.Add "Unit_Number", "Unit_Number"
    .KeyFields.Add "FacilityNCID", "FacilityNCID"
    .KeyFields.Add "EncounterNumber", "EncounterNumber"

    .AddField "FACILITYNCID"
    .AddField "DATAID"
    .AddField "TYPE"
    .AddField "ENCOUNTERNUMBER"
    .AddField "ENCOUNTERDATAID"
    .AddField "DTS"
    .AddField "OWNERNCID"
    .AddField "OWNERNAME"
    .AddField "TITLE"
    .AddField "RTF"
    .AddField "UPDATEFLAG"
    .AddField "UPDATEDBY"
    .AddField "PROFILENCID"
    .AddField "UPDATEDON"
    .AddField "CREATEDBY"
    .AddField "CREATEDON"
End With

mobjTables.Add objTable

Load_SoNotes = True

End Function

Private Function LoadEncounter_SoNotes() As Boolean
Dim objTable As XferTable
Dim objField As XferField

Set objTable = New XferTable

With objTable
    .TableName = "SoNotes"
    .SelectPart = "Select sn.* "
    .FromPart = " FROM Encounters enc, SoNotes sn "
    .WherePart = " enc.encounternumber = sn.encounternumber "
    .WherePatientPart = " enc.Unit_Number = "
    .AdditionalSort = ""
    
    .PatientUniqueFields.Add "Unit_Number", "Unit_Number"
    .KeyFields.Add "FacilityNCID", "FacilityNCID"
    .KeyFields.Add "EncounterNumber", "EncounterNumber"

    .AddField "FACILITYNCID"
    .AddField "DATAID"
    .AddField "TYPE"
    .AddField "ENCOUNTERNUMBER"
    .AddField "ENCOUNTERDATAID"
    .AddField "DTS"
    .AddField "OWNERNCID"
    .AddField "OWNERNAME"
    .AddField "TITLE"
    .AddField "RTF"
    .AddField "UPDATEFLAG"
    .AddField "UPDATEDBY"
    .AddField "PROFILENCID"
    .AddField "UPDATEDON"
    .AddField "CREATEDBY"
    .AddField "CREATEDON"
End With

mobjTables.Add objTable

LoadEncounter_SoNotes = True

End Function

Private Function Load_Inpatient_Admission() As Boolean

Dim objTable As XferTable
Dim objField As XferField

Set objTable = New XferTable

With objTable
    .TableName = "Inpatient_Admission"
    .SelectPart = "Select distinct ia.* "
    .FromPart = " FROM Encounters enc, Inpatient_Admission ia "
    .WherePart = " enc.apptid = ia.appt_id and enc.facilityncid = ia.facility_ncid and enc.unit_number = ia.unit_number "
    .WherePatientPart = " enc.Unit_Number = "
    .AdditionalSort = ""
    
    .PatientUniqueFields.Add "Unit_Number", "Unit_Number"
    .KeyFields.Add "Inpatient_ID", "Inpatient_ID"

    .AddField "Inpatient_ID"
    .AddField "Unit_Number"
    .AddField "Ward_NCID"
    .AddField "Assigned_Bed_Number"
    .AddField "Admission_Type_Code"
    .AddField "Admission_Date"
    .AddField "Discharge_Date"
    .AddField "Coments_Text"
    .AddField "Admitting_Provider_NCID"
    .AddField "Attending_Provider_NCID"
    .AddField "Primary_Service_NCID"
    .AddField "Appt_ID"
    .AddField "Time_Bias"
End With

mobjTables.Add objTable

Load_Inpatient_Admission = True

End Function


Private Function Load_NCIDS_Table() As Boolean

Dim objTable As XferTable
Dim objField As XferField

Set objTable = New XferTable

With objTable
    .TableName = "NCIDS"
    .SelectPart = "Select * "
    .FromPart = "FROM NCIDS "
    .WherePart = "NCIDS.DOMAINNCID = 119 AND NCIDS.NCID = "
    .WherePatientPart = ""
    .AdditionalSort = ""
    
    .KeyFields.Add "NCIDS", "NCID"
    .KeyFields.Add "DOMAINNCID", "DOMAINNCID"

    .AddField "NCID"
    .AddField "DOMAINNCID"
    .AddField "DESCRIPTION"
End With

mobjTables.Add objTable

Load_NCIDS_Table = True

End Function



Private Function Load_InptWard() As Boolean

Dim objTable As XferTable
Dim objField As XferField

Set objTable = New XferTable

With objTable
    .TableName = "Ward"
    .SelectPart = "Select distinct w.* "
    .FromPart = " FROM Encounters enc, Inpatient_Admission ia, Ward w "
    .WherePart = " enc.apptid = ia.appt_id and enc.facilityncid = ia.facility_ncid and enc.unit_number = ia.unit_number and ia.ward_ncid = w.ward_ncid and enc.facilityncid = w.facility_ncid "
    .WherePatientPart = " enc.Unit_Number = "
    .AdditionalSort = ""
    
    .PatientUniqueFields.Add "Unit_Number", "Unit_Number"
    .KeyFields.Add "Facility_NCID", "Facility_NCID"
    .KeyFields.Add "Ward_NCID", "Ward_NCID"

    .AddField "Ward_Location_NCID"
    .AddField "Ward_Name"
    .AddField "Ward_Bed_Number"
    .AddField "Ward_Status_Code"
End With

mobjTables.Add objTable

Load_InptWard = True

End Function


Private Function Load_InptWard_Location() As Boolean

Dim objTable As XferTable
Dim objField As XferField

Set objTable = New XferTable

With objTable
    .TableName = "Ward_Location"
    .SelectPart = "Select distinct wl.* "
    .FromPart = " FROM Encounters enc, Inpatient_Admission ia, Ward w, Ward_Location wl "
    .WherePart = " enc.apptid = ia.appt_id and enc.facilityncid = ia.facility_ncid and enc.unit_number = ia.unit_number and ia.ward_ncid = w.ward_ncid and w.ward_location_ncid = wl.ward_location_ncid and enc.facilityncid = wl.facility_ncid "
    .WherePatientPart = " enc.Unit_Number = "
    .AdditionalSort = ""
    
    .PatientUniqueFields.Add "Unit_Number", "Unit_Number"
    .KeyFields.Add "Facility_NCID", "Facility_NCID"
    .KeyFields.Add "Ward_Location_NCID", "Ward_Location_NCID"

    .AddField "Facility_NCID"
    .AddField "Ward_Location_NCID"
    .AddField "Ward_Location_Name"
End With

mobjTables.Add objTable

Load_InptWard_Location = True

End Function




Private Function Load_Inpatient_Transfer() As Boolean

Dim objTable As XferTable
Dim objField As XferField

Set objTable = New XferTable

With objTable
    .TableName = "Inpatient_Transfer"
    .SelectPart = "Select distinct it.* "
    .FromPart = " FROM Encounters enc, Inpatient_Admission ia, Inpatient_Transfer it "
    .WherePart = " enc.apptid = ia.appt_id and enc.facilityncid = ia.facility_ncid and enc.unit_number = ia.unit_number and it.inpatient_id = ia.inpatient_id "
    .WherePatientPart = " enc.Unit_Number = "
    .AdditionalSort = ""
    
    .PatientUniqueFields.Add "Unit_Number", "Unit_Number"
    .KeyFields.Add "Inpatient_ID", "Inpatient_ID"

    .AddField "Inpatient_ID"
    .AddField "Transfer_Date"
    .AddField "Ward_NCID"
    .AddField "Assigned_Bed_Number"
    .AddField "Attending_Provider_NCID"
    .AddField "Primary_Service_NCID"
    .AddField "Time_Bias"
End With

mobjTables.Add objTable

Load_Inpatient_Transfer = True

End Function




Private Function Load_Inpatient_Diagnosis() As Boolean

Dim objTable As XferTable
Dim objField As XferField

Set objTable = New XferTable

With objTable
    .TableName = "Inpatient_Diagnosis"
    .SelectPart = "Select distinct id.* "
    .FromPart = " FROM Encounters enc, Inpatient_Admission ia, Inpatient_Diagnosis id "
    .WherePart = " enc.apptid = ia.appt_id and enc.facilityncid = ia.facility_ncid and enc.unit_number = ia.unit_number and ia.inpatient_id = id.inpatient_id "
    .WherePatientPart = " enc.Unit_Number = "
    .AdditionalSort = ""

    .PatientUniqueFields.Add "Unit_Number", "Unit_Number"
    .KeyFields.Add "Inpatient_ID", "Inpatient_ID"
    .KeyFields.Add "SnoID", "SnoID"
    .KeyFields.Add "Diagnosis_Type", "Diagnosis_Type"

    .AddField "Inpatient_ID"
    .AddField "SnoID"
    .AddField "Diagnosis_Type"
    .AddField "Medcin_Prefix"
    .AddField "Complaint_NCID"
    .AddField "Complaint_Type"
    .AddField "Description"
End With

mobjTables.Add objTable

Load_Inpatient_Diagnosis = True

End Function


Private Function Load_Inpatient_Discharge() As Boolean

Dim objTable As XferTable
Dim objField As XferField

Set objTable = New XferTable

With objTable
    .TableName = "Inpatient_Discharge"
    .SelectPart = "Select id.* "  '--- SCR 82046  Sherry Wang   12/14/2005
    .FromPart = " FROM Encounters enc, Inpatient_Admission ia, Inpatient_Discharge id "
    .WherePart = " enc.apptid = ia.appt_id and enc.facilityncid = ia.facility_ncid and enc.unit_number = ia.unit_number and ia.inpatient_id = id.inpatient_id "
    .WherePatientPart = " enc.Unit_Number = "
    .AdditionalSort = ""

    .PatientUniqueFields.Add "Unit_Number", "Unit_Number"
    .KeyFields.Add "Inpatient_ID", "Inpatient_ID"

    .AddField "Inpatient_ID"
    .AddField "DISCHARGEDATE"
    .AddField "TIMEBIAS"
    .AddField "DISCHARGETEXT"
    .AddField "DISPOSITIONNCID"
    .AddField "DISPOSITIONTEXT"
    .AddField "ITEMSDISCUSSED"
    .AddField "DISCUSSEDTEXT"
    .AddField "POSTDISCHARGETEXT"
End With

mobjTables.Add objTable

Load_Inpatient_Discharge = True

End Function








Private Function Load_ListToolRecord() As Boolean
Dim objTable As XferTable
Dim objField As XferField

Set objTable = New XferTable

With objTable
    .TableName = "ListToolRecord"
    .SelectPart = "Select ltr.* "
    .FromPart = " FROM Encounters enc, ListToolRecord ltr "
    .WherePart = " enc.encounternumber = ltr.encounternumber "
    .WherePatientPart = " enc.Unit_Number = "
    .AdditionalSort = ""
    
    .PatientUniqueFields.Add "Unit_Number", "Unit_Number"
    .KeyFields.Add "DataID", "DataID"

    .AddField "DATAID"
    .AddField "FACILITYNCID"
    .AddField "SnoID"
    .AddField "CLINICIANNCID"
    .AddField "ENCOUNTERNUMBER"
    .AddField "CLINICNCID"
    .AddField "STARTDTS"
    .AddField "ENDDTS"
    .AddField "SOURCENCID"
    .AddField "RTF"
    .AddField "UPDATEFLAG"
    .AddField "CREATEDBY"
    .AddField "CREATEDON"
    .AddField "UPDATEDBY"
    .AddField "UPDATEDON"

End With

mobjTables.Add objTable

Load_ListToolRecord = True

End Function

Private Function Load_ListToolRecords() As Boolean
Dim objTable As XferTable
Dim objField As XferField

Set objTable = New XferTable

With objTable
    .TableName = "ListToolRecords"
    .SelectPart = "Select ltrs.* "
    .FromPart = " FROM Encounters enc, ListToolRecord ltr, ListToolRecords ltrs "
    .WherePart = " enc.encounternumber = ltr.encounternumber and ltr.dataid = ltrs.dataid "
    .WherePatientPart = " enc.Unit_Number = "
    .AdditionalSort = ""
    
    .PatientUniqueFields.Add "Unit_Number", "Unit_Number"
    .KeyFields.Add "DataID", "DataID"
    .KeyFields.Add "SnoID", "SnoID"
    .KeyFields.Add "Prefix", "Prefix"

    .AddField "DATAID"
    .AddField "SnoID"
    .AddField "PREFIX"
    .AddField "MODIFIER"
    .AddField "RESULT"
    .AddField "STATUS"
    .AddField "ONSET"
    .AddField "DURATION"
    .AddField "VALUE"
    .AddField "UNIT"
    .AddField "RANGESCALE"
    .AddField "REFERID"
    .AddField "RANGENORMALLOW"
    .AddField "RANGENORMALHIGH"
    .AddField "CHARTFLAG"
    .AddField "QLINK"
    .AddField "QLINKSEQ"
    .AddField "NOTE"
    .AddField "UPDATEFLAG"
    .AddField "CREATEDBY"
    .AddField "CREATEDON"
    .AddField "UPDATEDBY"
    .AddField "UPDATEDON"

End With

mobjTables.Add objTable

Load_ListToolRecords = True

End Function


'Changed Table name from ENC_DIAGNOSIS to DIAGNOSES SCR-22812
Private Function LoadEncounter_Diagnosis() As Boolean
Dim objTable As XferTable
Dim objField As XferField

Set objTable = New XferTable

With objTable
    .TableName = "Diagnoses"
    .SelectPart = "Select ed.* "
    .FromPart = " FROM Encounters enc, Diagnoses ed "
    .WherePart = " enc.encounternumber = ed.encounternumber "
    .WherePatientPart = " enc.Unit_Number = "
    .AdditionalSort = ""
    
    .PatientUniqueFields.Add "Unit_Number", "Unit_Number"
    .KeyFields.Add "DiagnosisNCID", "DiagnosisNCID"

    .AddField "FACILITYNCID"
    .AddField "DIAGNOSISNCID"
    .AddField "ENCOUNTERNUMBER"
    .AddField "CREATEDBY"
    .AddField "CREATEDON"
    .AddField "UPDATEDBY"
    .AddField "UPDATEDON"

End With

mobjTables.Add objTable

LoadEncounter_Diagnosis = True

End Function

'**************************************************************
'SCR 67372
'Datamanager needs to export the OE_RX table for PVF and JMEWS.
'**************************************************************
Private Function Load_OE_RX() As Boolean

Dim objTable As XferTable
Dim objField As XferField

Set objTable = New XferTable

With objTable
    .TableName = "OE_RX"
    .SelectPart = "Select RX.* "
    .FromPart = " FROM Encounters enc, OE_RX rx "
    .WherePart = " enc.encounternumber = rx.ENC_NUM "
    .WherePatientPart = " enc.Unit_Number = "
    .AdditionalSort = ""
    
    .PatientUniqueFields.Add "Unit_Number", "Unit_Number"
    .KeyFields.Add "ENC_NUM", "ENC_NUM"
    .KeyFields.Add "ENC_FAC_NCID", "ENC_FAC_NCID"
    
    .AddField "OE_RX_DATA_ID"
    .AddField "ENC_NUM"
    .AddField "ENC_FAC_NCID"
    .AddField "ORDER_DATE"
    .AddField "RX_COMMENT"
    .AddField "SIG_CODE"
    .AddField "QUANTITY"
    .AddField "REFILLS"
    .AddField "PHARMACY_NAME_SENT"
    .AddField "PHARMACY_IEN_SENT"
    .AddField "CHILD_CAP"
    .AddField "ORDER_IEN"
    .AddField "ORDER_NAME"
    .AddField "TIME_ORDERED"
    .AddField "UNIT_NUMBER"
    .AddField "ORDER_ID"
    .AddField "SENSITIVE"
    .AddField "ORDERED_BY"
    .AddField "Provider"
    .AddField "UPDATEDBY"
    .AddField "UpdatedOn"
    .AddField "CreatedBy"
    .AddField "CreatedOn"
    .AddField "STATUS_NCID"
    .AddField "FILL_DATE"
    .AddField "SIGN_FLAG"
    .AddField "NURSE_ACK_FALG"
    .AddField "PATIENT_TYPE"
    .AddField "PROCESS_PRIORITY"
    .AddField "NURSE_ACK_NAME"
    .AddField "NURSE_ACK_DATE"
    .AddField "PREVIOUS_ORDER_STATUS"

End With

mobjTables.Add objTable

Load_OE_RX = True

End Function

Private Function LoadEncounter_Disposition() As Boolean
Dim objTable As XferTable
Dim objField As XferField

Set objTable = New XferTable

With objTable
    .TableName = "Enc_Disposition"
    .SelectPart = "Select ed.* "
    .FromPart = " FROM Encounters enc, Enc_Disposition ed "
    .WherePart = " enc.encounternumber = ed.encounternumber "
    .WherePatientPart = " enc.Unit_Number = "
    .AdditionalSort = ""
    
    .PatientUniqueFields.Add "Unit_Number", "Unit_Number"
    .KeyFields.Add "EncounterNumber", "EncounterNumber"
    .KeyFields.Add "FacilityNCID", "FacilityNCID"

    .AddField "ENCOUNTERNUMBER"
    .AddField "FACILITYNCID"
    .AddField "NCID"
    .AddField "EANDMCODEREVIEW"
    .AddField "EANDMREVIEWEDDATE"
    .AddField "EANDMREVIEWEDBY"
    .AddField "EANDMREVIEWED"
    .AddField "REVIEWEDBYCODER"
    .AddField "DISPADMINOPTION"
    .AddField "DISPOSITIONNCID"
    .AddField "DISPOSITIONTEXT"
    .AddField "ITEMSDISCUSSED"
    .AddField "FOLLOWUPTIMEFRAME"
    .AddField "FOLLOWUUPCOMMENTS"
    .AddField "WSDISP"
    .AddField "EANDMCALC"
    .AddField "INJ_ILL_CATEGORY"
    .AddField "INJ_ILL_CAUSE"
End With

mobjTables.Add objTable

LoadEncounter_Disposition = True

End Function

Private Function LoadEncounter_TextNotes() As Boolean
Dim objTable As XferTable
Dim objField As XferField

Set objTable = New XferTable

With objTable
    .TableName = "EncTextNotes"
    .SelectPart = "Select etn.* "
    .FromPart = " FROM Encounters enc, EncTextNotes etn "
    .WherePart = " enc.encounternumber = etn.encounternumber "
    .WherePatientPart = " enc.Unit_Number = "
    .AdditionalSort = ""
    
    .PatientUniqueFields.Add "Unit_Number", "Unit_Number"
    .KeyFields.Add "DataID", "DataID"

    .AddField "DATAID"
    .AddField "FACILITYNCID"
    .AddField "MODIFYCLINICIANNCID"
    .AddField "ENCOUNTERNUMBER"
    .AddField "CREATECLINICIANNCID"
    .AddField "STARTTIME"
    .AddField "RTF"
    .AddField "COMPLETE"
    .AddField "CATEGORY"
    .AddField "TITLE"
    .AddField "ENDTIME"
    .AddField "CREATEDBY"
    .AddField "CREATEDON"
    .AddField "UPDATEDBY"
    .AddField "UPDATEDON"
End With

mobjTables.Add objTable

LoadEncounter_TextNotes = True

End Function


Private Function Load_PiData() As Boolean
Dim objTable As XferTable
Dim objField As XferField

Set objTable = New XferTable

With objTable
    .TableName = "Pi_Data"
    .SelectPart = "Select pd.* "
    .FromPart = " FROM Encounters enc, Pi_Data pd "
    .WherePart = " enc.encounternumber = pd.enc_num "
    .WherePatientPart = " enc.Unit_Number = "
    .AdditionalSort = ""
    
    .PatientUniqueFields.Add "Unit_Number", "Unit_Number"
    .KeyFields.Add "Pi_Data_Id", "Pi_Data_Id"

    .AddField "PI_DATA_ID"
    .AddField "SNOID"
    .AddField "UNIT_NUMBER"
    .AddField "PI_NCID"
    .AddField "ENC_NUM"
    .AddField "ENC_FAC_NCID"
    .AddField "PI_TIME"

End With

mobjTables.Add objTable

Load_PiData = True

End Function

Private Function Load_Vitals() As Boolean
Dim objTable As XferTable
Dim objField As XferField

Set objTable = New XferTable

With objTable
    .TableName = "Vitals"
    .SelectPart = "Select v.* "
    .FromPart = " FROM Encounters enc, Vitals v "
    .WherePart = " enc.encounternumber = v.encounternumber "
    .WherePatientPart = " enc.Unit_Number = "
    .AdditionalSort = ""
    
    .PatientUniqueFields.Add "Unit_Number", "Unit_Number"
    .KeyFields.Add "DataID", "DataID"

    .AddField "CLINICIANNCID"
    .AddField "UNIT_NUMBER"
    .AddField "FACILITYNCID"
    .AddField "ENCOUNTERNUMBER"
    .AddField "DATAID"
    .AddField "PANELTYPE"
    .AddField "ENTRYDATE"
    .AddField "VITALSCOMMENT"
    .AddField "UPDATEFLAG"
    .AddField "CDR_DATAID"
    .AddField "CREATEDBY"
    .AddField "CREATEDON"
    .AddField "UPDATEDBY"
    .AddField "UPDATEDON"

End With

mobjTables.Add objTable

Load_Vitals = True

End Function

Private Function Load_ClinNote() As Boolean
Dim objTable As XferTable
Dim objField As XferField

Set objTable = New XferTable

With objTable
    .TableName = "ClinNote"
    .SelectPart = "Select c.* "
    .FromPart = " FROM Encounters enc, ClinNote c "
    .WherePart = " enc.encounternumber = c.encounter "
    .WherePatientPart = " enc.Unit_Number = "
    .AdditionalSort = ""
    
    .PatientUniqueFields.Add "Unit_Number", "Unit_Number"
    .KeyFields.Add "DataId", "DataId"

    .AddField "DATAID"
    .AddField "APPLICATIONNCID"
    .AddField "NOTETYPENCID"
    .AddField "NOTETYPENAME"
    .AddField "NOTEDATE"
    .AddField "POCFACILITY"
    .AddField "ENCFACILITY"
    .AddField "CLINICIAN"
    .AddField "ENCOUNTER"
    .AddField "HTMLTEXT"
    .AddField "CALLER"
    .AddField "STATUS"
    .AddField "PRIORITY"
    .AddField "CREATEDBY"
    .AddField "CREATEDON"
    .AddField "UPDATEDBY"
    .AddField "UPDATEDON"
    .AddField "IMAGEAVAILABLE"
    .AddField "IMAGENUMBER"
    .AddField "CDR_DATAID"
    .AddField "UNIT_NUMBER"
    .AddField "ORIGINALSIZE"
    .AddField "MODIFIEDDATEBIAS"
    .AddField "NOTEDATEBIAS"
    .AddField "UPDATEFLAG"
    .AddField "CREATEDATEBIAS"

End With

mobjTables.Add objTable

Load_ClinNote = True

End Function

Private Function Load_WorkStatus() As Boolean
Dim objTable As XferTable
Dim objField As XferField

Set objTable = New XferTable

With objTable
    .TableName = "Work_Status"
    .SelectPart = "Select w.* "
    .FromPart = " FROM Encounters enc, Work_Status w "
    .WherePart = " enc.encounternumber = w.encounter_number"
    .WherePatientPart = " enc.Unit_Number = "
    .AdditionalSort = ""
    
    .PatientUniqueFields.Add "Unit_Number", "Unit_Number"
    .KeyFields.Add "Encounter_Number", "Encounter_Number"
    .KeyFields.Add "Facility_Id", "Facility_Id"
    .KeyFields.Add "Work_Status_NCID", "Work_Status_NCID"

    .AddField "PATIENT_NCID"
    .AddField "ENCOUNTER_NUMBER"
    .AddField "WORK_STATUS_NCID"
    .AddField "WORK_STATUS_SELECTED"
    .AddField "WORK_STATUS_QUALIFIED"
    .AddField "FACILITY_ID"
    .AddField "PROVIDER_NCID"
    .AddField "LAST_UPDATED"

End With

mobjTables.Add objTable

Load_WorkStatus = True

End Function


Private Function Load_EncSigBlock() As Boolean
Dim objTable As XferTable
Dim objField As XferField

Set objTable = New XferTable

With objTable
    .TableName = "Enc_SigBlock"
    .SelectPart = "Select e.* "
    .FromPart = " FROM Enc_SigBlock e "
    .WherePart = ""
    .WherePatientPart = ""
    .AdditionalSort = ""
    
    '.PatientUniqueFields.Add "Unit_Number", "Unit_Number"
    '.KeyFields.Add "Encounter_Number", "Encounter_Number"
    '.KeyFields.Add "Facility_Id", "Facility_Id"
    .KeyFields.Add "DATAID", "DATAID"

    .AddField "DATAID"
    .AddField "OWNERDATAID"
    .AddField "USERNCID"
    .AddField "LINE1"
    .AddField "LINE2"
    .AddField "LINE3"
    .AddField "DTS"
    .AddField "PKI"
    .AddField "CREATEDBY"
    .AddField "CREATEDON"
    .AddField "UPDATEDBY"
    .AddField "UPDATEDON"

End With

mobjTables.Add objTable

Load_EncSigBlock = True

End Function

'Transfer Symptoms Function SCR-26545
Private Function TransferSymptoms(DMBEObj As IDMBE.I_DMBE, ByVal EncID As Long, ByRef OutPutRS As Adodb.Recordset) As Variant

    Dim oChart      As CHCSIIEncounterOps.ListToolRecord
    Dim LTRRS       As Adodb.Recordset
    Dim ltrsRS      As Adodb.Recordset
    'Dim OutPutRS    As ADODB.Recordset
    Dim oDas        As ICHCSII_DAS
    Dim sTxt        As String
    Dim objRS       As IDMBE.I_RSObjects
    Dim colRS       As Collection
    
    Set oDas = gobjCHCSIIConn.CHCSII_DAS(Auto)
    Set oChart = New CHCSIIEncounterOps.ListToolRecord
    Set colRS = New Collection
'__ sanity check:  Only for transfer from LDDB to TMIP or EIC
    If Not (DMBEObj.Source = eDM_LDDB _
     And (DMBEObj.Destination = eDM_TMIP Or DMBEObj.Destination = eDM_EIC)) Then Exit Function
     
    If OutPutRS Is Nothing Then
        Set OutPutRS = New Adodb.Recordset
        With OutPutRS
            .CursorLocation = adUseClient
            .LockType = adLockBatchOptimistic
            .ActiveConnection = Nothing
            .Fields.Append "EncounterNumber", adInteger, 20
            .Fields.Append "SnoID", adInteger, 20
            .Fields.Append "MedcinDesc", adVarChar, 1024
            .Open
        End With
    End If
    
    Set LTRRS = oDas.ExecuteSQL("select dataid from listtoolrecord where encounternumber = " & EncID)

    Do While Not LTRRS.EOF
    Dim oConvert As New IConvert
        Set ltrsRS = oDas.ExecuteSQL("select * from listtoolrecords where result <> 'N' and dataid = " & LTRRS(0))
        Do While Not ltrsRS.EOF
          
            'SCR 72733 Only should include symptoms, not other data. Basing it on " A" is not good.
            'Added code to look up the medcin term type.
            If oConvert.SnoTermType(ltrsRS("SnoID")) = ettSymptom Then
                '***********************************************************************
                'SCR 67271 - This If Satement is producing a False result.  No Data
                'If Left$(oConvert.MedcinNodeKey(ltrsRS("SnoID")), 2) = " A" Then  'Note that is "Space-A"
                OutPutRS.AddNew
                OutPutRS("EncounterNumber") = EncID
                OutPutRS("SnoID") = ltrsRS("SnoID")
                OutPutRS("MedcinDesc") = oConvert.SnoDescription(ltrsRS("SnoID"))
                OutPutRS.Update
                'End If
                '***********************************************************************
            End If
          
            ltrsRS.MoveNext
        Loop
        LTRRS.MoveNext
    Loop
'__ Setup the RecordSetTransfer objects
'    Set objRS = New DMBE.RSObject_Ops
'    objRS.TableName = "SYMPTOMS"
'    Set objRS.DBRecordSet = OutPutRS
'    colRS.Add objRS
'    Call DMBEObj.Execute(colRS)
    
End Function





