VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IMM_Tables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'...Error Handling Variables
Private lLastError As Long
Private sLastErrorDesc As String

'...Data contains all Group Names for a given MTF_ID indexed by MTF_ID
Private UserGroupsMTF_ID As Dictionary

'...Data contains all Group Names for a given Unit / MTF_ID indexed by MTF_ID
Private UserGroupsUnitMTF_ID As Dictionary

'...contains a record for a Group Indexed by
'...    MTF_ID / GROUP    OR     DEFAULT / GROUP
'...Last field in Record contains all Generic Immunizations
'...associated with this group
Private PatientGroups As Dictionary

Private obd_AllGenericVacNames As Dictionary

'...Holds Vaccine to VacID conversion Indexed by Vaccine
Private obd_VaccineToVacID As Dictionary

'...Holds VacID to GenericName conversion Indexed by VacID
Private obd_VacIDToGenName As Dictionary

'...Holds Generic Name to all vacIDs Indexed by Generic Name
Public obd_GenNametoVacIDs As Dictionary

'...Dictionary of Vaccine Data Indexed by VacID
Private obd_Vaccines As Dictionary

'...Dictionary of Manufacturers Indexed by MFG Code
Private obd_MfgCodeToMfgName As Dictionary

'<< SCR #25155;   Developer: Brian Mowbray 06/27/2002 02:01 PM
'...Dictionary of Vaccine Information Statements Indexed by VaccineID
Private obd_VIS_Versions As Dictionary

'...Dictionary of Manufacturers Indexed by MFG Name
Private obd_MfgNameToMfgCode As Dictionary

'...Dictionary of Vaccines Solved By Index is Vaccine being solved
Private obd_VaccinesSolvedBy As Dictionary

'...Dictionary of Generic Names that are
'...Related to other generic names
Private obd_RelatedGenericNames As Dictionary

'...All Groups Required for each Patient in a Unit
Private UnitGroupReqForPatient As Dictionary
Private LastUnitUsed As String

'....String of VacIDs that are Combination Drugs
'...Format  |V1|V2|...|Vx|
Private CombinationVaccines As String

Public Sub LoadVaccines()

' This Routine will populate the following dictionaries

'     obd_Vaccines
'     obd_VacIDToGenName
'     obd_GenNametoVacIDs


  On Local Error GoTo haveerror

 Dim sql As String
 Dim gn As String
 Dim GenName_RS As ADODB.Recordset
 Dim Schedule_RS As ADODB.Recordset
 Dim Count As Integer
 Dim sepchar As String
 Dim sepchar2 As String
 Dim GenRec As String
 Dim xfield As String
 Dim RelatedNames As String
 Dim Grp1VaccineData As String
 Dim Grp2VaccineData As String
 Dim Schedule As String
 Dim Conditions As String
 Dim swork As String
 Dim VaccineID As Long
 Dim condition As Integer
 Dim Vaccine As String
 Dim MinAge As String
 Dim MaxAge As String
 Dim TakesPlaceOf As String
 Dim Solved() As String
 Dim solver As String
 Dim VacID As Variant
 '<< SCR #25155;   Developer: Brian Mowbray 06/27/2002 02:01 PM
 Dim sVIS_Version As String
 
 
 
   FrmStatus.LabStatus = FrmStatus.LabStatus & vbCrLf & "Loading Immunizations:"
   FrmStatus.Refresh
 
 ' Use the following Data shaping once the DAS connection allows it

'sql = _
'     "SHAPE " & _
'      "{ " & _
'        "SELECT GenericName, VaccineID,Vaccine, NBRINSERIES " & _
'        "FROM TBLVaccines " & _
'      "} " & _
'    "APPEND " & _
'    "( " & _
'       "{ " & _
'       "SELECT age_from, Age_To, Min_Age, Max_Age, Min_Interval, Min_Count, Recm_Age, Recm_Interval, Skip_age " & _
'       "FROM TBLSchedule " & _
'       "Where TBLVaccines.VaccineID =  tblschedule.vaccineID Order by DOSENUMBER" & _
'       "} " & _
'     "RELATE TBLVaccines.VaccineID TO TBLSchedule.VaccineID " & _
'    ") " & _
'    "AS appSchedule"

' The following code takes the place of data shaping

 '<< SCR #25155;   Developer: Brian Mowbray 06/27/2002 02:01 PM
 '...Read Vaccine Table
 sql = "Select GenericName, Inactive, Min_Age, Max_Age, VaccineID, HL7_ID, Vaccine, NBRINSERIES, EffectsVacID, " & _
       "CPT, ICD, Booster, MPF_DIN, ChildHood, HistKey, GenNameID, ImmRelNCID, VIS_Version" & " From tblvaccines order by VaccineID"
 Set GenName_RS = gobjSQL.OpenRecordset(sql, adOpenForwardOnly, adLockReadOnly)
 If Not GenName_RS.EOF Then
 
    '...Read Schedule Table
    sql = "Select VaccineID, DoseNumber, age_from, Age_To, Min_Age, Min_Interval, Min_Count, Recm_Age, Recm_Interval, Skip_age " & _
          "From TBLSchedule order by VaccineID, DoseNumber"
    Set Schedule_RS = gobjSQL.OpenRecordset(sql, adOpenForwardOnly, adLockReadOnly)
    
    With GenName_RS
      Do While Not .EOF
        sepchar = ""
        GenRec = ""
'        Debug.Assert CStr(.Fields("VaccineID")) <> "900"
        For Count = 0 To .Fields.Count - 1
           Select Case Count
             Case VacDictGrp1Fields.GenericName
              xfield = .Fields("GenericName") & vbNullString
              gn = xfield
             Case VacDictGrp1Fields.HL7_ID
              xfield = .Fields("HL7_ID") & vbNullString
             Case VacDictGrp1Fields.NbrInSeries
              xfield = CStr(.Fields("NBRInSeries"))
             Case VacDictGrp1Fields.Vaccine
              xfield = .Fields("Vaccine") & vbNullString
              Vaccine = xfield
             Case VacDictGrp1Fields.VaccineID
              VaccineID = .Fields("VaccineID")
              xfield = CStr(VaccineID)
             Case VacDictGrp1Fields.MinAge
              xfield = .Fields("Min_Age") & vbNullString
              MinAge = xfield
             Case VacDictGrp1Fields.MaxAge
              xfield = .Fields("Max_Age") & vbNullString
              MaxAge = xfield
             Case VacDictGrp1Fields.EffectsVacID
              xfield = .Fields("EffectsVacID") & vbNullString
              TakesPlaceOf = xfield
             Case VacDictGrp1Fields.CPT
              xfield = .Fields("CPT") & vbNullString
             Case VacDictGrp1Fields.ICD
              xfield = .Fields("ICD") & vbNullString
             Case VacDictGrp1Fields.Booster
              xfield = .Fields("Booster")
             Case VacDictGrp1Fields.MPF_DIN
              xfield = .Fields("MPF_Din") & vbNullString
             Case VacDictGrp1Fields.ChildHood
              xfield = .Fields("Childhood")
             Case VacDictGrp1Fields.InActive
              xfield = .Fields("InActive")
             Case VacDictGrp1Fields.HistKey
              xfield = .Fields("HistKey") & vbNullString
             Case VacDictGrp1Fields.GenNameID
              xfield = .Fields("GenNameID")
             Case VacDictGrp1Fields.ImmRel_Ncid
              xfield = .Fields("ImmRelNcid")
             '<< Begin:SCR #25155;   Developer: Brian Mowbray 06/27/2002 02:01 PM
             Case VacDictGrp1Fields.VIS_Version
              xfield = .Fields("VIS_Version") & vbNullString
              sVIS_Version = xfield
             '>> End: SCR #25155;
           End Select
           '...Build General Data Record for Vaccine
           GenRec = GenRec & sepchar & xfield
           sepchar = FieldSepChar
        Next Count
        
        ' SCR 18130 T.Noren  10/29/01 remove 'Force to upper case
'        If Not obd_VaccineToVacID.Exists(UCase$(Vaccine)) Then
        If Not obd_VaccineToVacID.Exists(Vaccine) Then
            '...Create inverted file for Vaccine
            ' SCR 18130 T.Noren  10/29/01 remove 'Force to upper case
'            obd_VaccineToVacID.Add UCase$(Vaccine), CStr(VaccineID)
            obd_VaccineToVacID.Add Vaccine, CStr(VaccineID)
            
            
            '...Create inverted file for VaccineID
            obd_VacIDToGenName.Add CStr(VaccineID), gn & FieldSepChar & Vaccine & FieldSepChar & MinAge & FieldSepChar & MaxAge & FieldSepChar & TakesPlaceOf
            
    '        Debug.Assert CStr(VaccineID) <> "51"
            '...Create Inverted file for Vaccines solved by
            If TakesPlaceOf <> "" Then
            
              '...Save Combination Vaccine
              CombinationVaccines = CombinationVaccines & WithinFieldSepChar & CStr(VaccineID)
              
              '...Load Vaccines solved by VaccineID
              Solved = Split(Mid$(TakesPlaceOf, 2, Len(TakesPlaceOf) - 2), WithinFieldSepChar)
              
              For Each VacID In Solved
                 If Not obd_VaccinesSolvedBy.Exists(CStr(VacID)) Then
                   obd_VaccinesSolvedBy.Add CStr(VacID), ""
                   solver = ""
                 Else
                   solver = obd_VaccinesSolvedBy.Item(VacID)
                 End If
                 solver = solver & WithinFieldSepChar & CStr(VaccineID)
                 obd_VaccinesSolvedBy.Item(VacID) = solver
              Next VacID
            
            End If
            
            
            '...Create inverted file for Generic Names referencing a vacID
            If obd_GenNametoVacIDs.Exists(gn) Then
             obd_GenNametoVacIDs.Item(gn) = obd_GenNametoVacIDs.Item(gn) & InvertedTblSepChar & CStr(VaccineID)
            Else
             obd_GenNametoVacIDs.Add gn, CStr(VaccineID)
            End If
            
            '<< Begin:SCR #25155;   Developer: Brian Mowbray 06/27/2002 02:02 PM
            '...Create inverted file for VIS Versions
            If Len(sVIS_Version) > 0 Then
                If Not obd_VIS_Versions.Exists(Get_A_Generic_Name(VaccineID)) Then
                  obd_VIS_Versions.Add Get_A_Generic_Name(VaccineID), sVIS_Version
                End If
            End If
            '>> End: SCR #25155;
            
            
            '...Get Schedule information for this VaccineID
            '...VaccineID are in ascending order to match
            '...the order of Vaccine Record Set
            With Schedule_RS
             Schedule = ""
             sepchar2 = ""
             condition = -1
    '         Debug.Assert CStr(.Fields("VaccineID")) <> "900"
foundID:
             If Not .EOF Then
                Select Case CLng(.Fields("VaccineID"))
                
                     Case Is = VaccineID    ' Found proper VaccineID
                       condition = 0
                       
                     Case Is > VaccineID    ' This condition should never happen
                       condition = 1
                       
                     Case Else
                       condition = 2
                       MsgBxARMd Disclaimer & _
                              "Vaccine ID:  " & CStr(.Fields("VaccineID")) & " Is not defined in the Table TBLVaccines." & vbCrLf & _
                              "  OR  " & vbCrLf & _
                              "Table TblSchedule is corrupt" & vbCrLf & "VaccineID is out of sequence", vbExclamation + vbOKOnly, "Invalid Database Condition"
                       
                End Select
                
                If condition = 0 Then
                   Do
                      swork = ""
                      sepchar = ""
    '                  Debug.Assert CStr(.Fields("VaccineID")) <> "900"
                      For Count = 0 To .Fields.Count - 1
                         Select Case Count
                            Case VacDictGrp2Fields.VaccineID
                             xfield = .Fields("VaccineID")
                            Case VacDictGrp2Fields.DoseNumber
                             xfield = .Fields("DoseNumber")
                            Case VacDictGrp2Fields.AgeFrom
                             xfield = .Fields("Age_From") & vbNullString
                            Case VacDictGrp2Fields.AgeTo
                             xfield = .Fields("Age_To") & vbNullString
                            Case VacDictGrp2Fields.MinAge
                             xfield = .Fields("Min_Age") & vbNullString
                            Case VacDictGrp2Fields.MinCount
                             xfield = .Fields("Min_Count") & vbNullString
                            Case VacDictGrp2Fields.MinInterval
                             xfield = .Fields("Min_Interval") & vbNullString
                            Case VacDictGrp2Fields.RecmAge
                             xfield = .Fields("Recm_Age") & vbNullString
                            Case VacDictGrp2Fields.RecmInterval
                             xfield = .Fields("Recm_Interval") & vbNullString
                            Case VacDictGrp2Fields.SkipAge
                             xfield = .Fields("Skip_age") & vbNullString
                         End Select
                         swork = swork & sepchar & xfield
                         sepchar = FieldSepChar
                      Next Count
                      Schedule = Schedule & sepchar2 & swork
                      sepchar2 = SeriesSepChar
                      .MoveNext
                      If .EOF Then Exit Do
                   Loop While VaccineID = .Fields("VaccineID")
                ElseIf condition = 2 Then
                  '...loop untill we find vaccineID or pass it up
                   Do
                      .MoveNext
                      If .EOF Then Exit Do
                   Loop While VaccineID > .Fields("VaccineID")
                   If Not .EOF Then
                     If VaccineID = .Fields("VaccineID") Then GoTo foundID
                   End If
                End If
             End If
            End With
            
            Conditions = ""
            
            ' Process Conditions here
            
            obd_Vaccines.Add CStr(VaccineID), GenRec & GroupSepChar & Schedule & GroupSepChar & Conditions
            If condition = 2 Then Exit Do
            .MoveNext
        Else
          MsgBxARMd Disclaimer & _
                 "Vaccine:  " & Vaccine & vbCrLf & vbCrLf & _
                 "Is defined multiple times in Table: TBLVACCINES.", vbExclamation + vbOKOnly, "Invalid Situation"
        End If
      Loop
    End With
 End If

 Set GenName_RS = Nothing
 Set Schedule_RS = Nothing
   FrmStatus.LabStatus = FrmStatus.LabStatus & vbCrLf & "Successful"
   FrmStatus.Refresh

 Exit Sub
 
haveerror:
'  Debug.Print Err.Number, Err.Description
'  Resume
  lLastError = Err.Number
  sLastErrorDesc = "Imm_Tables: LoadVaccines: " & Err.Description
  Err.Raise lLastError, "Imm_Tables: LoadVaccines: ", sLastErrorDesc
 
End Sub

' SCR #19803;   Developer: Brian Mowbray 05/10/2002 01:00 PM
Public Sub LoadMfgs()

' This Routine will populate the following dictionaries

'     obd_MfgCodeToMfgName
'     obd_MfgNameToMfgCode
' Need to add remaining elements of mfg to new dictionary

  On Local Error GoTo haveerror

 Dim sSQL As String
 Dim iCnt As Integer
 Dim rst As ADODB.Recordset
 Dim sfield As String
 Dim sMfg_Code As String
 Dim sMfg_Name As String
 Dim sMfgRec As String
 
 
   FrmStatus.LabStatus = FrmStatus.LabStatus & vbCrLf & "Loading Manufacturers:"
   FrmStatus.Refresh
 
 '...Read Manufacturer Table
 sSQL = "SELECT Mfg_Code, Vaccine_Mfg FROM tblVaccine_Mfg ORDER BY Vaccine_Mfg"
 Set rst = gobjSQL.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
 If Not rst.EOF Then
 
    With rst
      Do While Not .EOF
        For iCnt = 0 To .Fields.Count - 1
           Select Case iCnt
             Case MfgDictFields.Mfg_Code
              sfield = .Fields("Mfg_Code") & vbNullString
              sMfg_Code = sfield
             Case MfgDictFields.Vaccine_Mfg
              sfield = .Fields("Vaccine_Mfg") & vbNullString
              sMfg_Name = sfield
           End Select
        Next iCnt
        
        If Not obd_MfgCodeToMfgName.Exists(sMfg_Code) Then
            '...Create inverted file for Mfg Code
            obd_MfgCodeToMfgName.Add sMfg_Code, sMfg_Name
        End If
        If Not obd_MfgNameToMfgCode.Exists(sMfg_Name) Then
          '...Create inverted file for Mfg Name
          obd_MfgNameToMfgCode.Add sMfg_Name, sMfg_Code
        End If
        rst.MoveNext
      Loop
    End With
 End If

 Set rst = Nothing
 FrmStatus.LabStatus = FrmStatus.LabStatus & vbCrLf & "Successful"
 FrmStatus.Refresh

 Exit Sub
 
haveerror:
'  Debug.Print Err.Number, Err.Description
'  Resume
  lLastError = Err.Number
  sLastErrorDesc = "Imm_Tables: LoadMfgs: " & Err.Description
  Err.Raise lLastError, "Imm_Tables: LoadMfgs: ", sLastErrorDesc
End Sub


Private Sub LoadGenericNames()

  On Local Error GoTo haveerror

 Dim sql As String
 Dim gn As String
 Dim xSelect As String
 Dim xFrom  As String
 Dim xWhere As String
 Dim GenName_RS As ADODB.Recordset
 Dim Count As Integer
 Dim sepchar As String
 Dim GenRec As String
 Dim xfield As String
 Dim RelatedNames As String
 
 
   FrmStatus.LabStatus = FrmStatus.LabStatus & vbCrLf & "Loading Generic Names:"
   FrmStatus.Refresh
   
 xSelect = "Select N.GenericName, N.MinAge, N.MaxAge, N.GENNAMEID, " & _
           "N.Type, N.Cycle, N.Range, Defunct, PHYSREQONLY, Combination, N.Related_Gen_Names"
 
 ' Read in all Generic Names
 xFrom = Name_Of_Generic_Vac_Names_Table & "N"
 '...Concantenate SQL statement
 sql = xSelect & " FROM " & xFrom     '   & " WHERE " & xWhere & IIf(xWhere = "", "", "AND ") & "N.Defunct = 0 AND N.PhysReqOnly = 0"         ' TJN_NEED ... completed
 '...Retrieve Records
 Set GenName_RS = gobjSQL.OpenRecordset(sql, adOpenForwardOnly, adLockReadOnly)
 
 With GenName_RS
 
  If Not .EOF Then
  
    '...Load Generic Name Dictionary for fast Generic Search
    Do While Not .EOF
      sepchar = ""
      GenRec = ""
      For Count = 0 To .Fields.Count - 1
        Select Case Count
          Case GenVacDictFields.GenericName
             gn = Trim(.Fields("GenericName"))
             xfield = gn
          Case GenVacDictFields.GenericID
             xfield = .Fields("GENNAMEID")
          Case GenVacDictFields.MinAge
             xfield = .Fields("MinAge") & vbNullString
          Case GenVacDictFields.MaxAge
             xfield = .Fields("MaxAge") & vbNullString
          Case GenVacDictFields.xtype
             xfield = .Fields("Type") & vbNullString
          Case GenVacDictFields.Cycle
             xfield = .Fields("Cycle") & vbNullString
          Case GenVacDictFields.Range
             xfield = .Fields("Range") & vbNullString
          Case GenVacDictFields.Discontinued
             xfield = CStr(.Fields("Defunct").Value & vbNullString)
          Case GenVacDictFields.RecmdDTROnly
             xfield = CStr(.Fields("PHYSREQONLY").Value & vbNullString)
          Case GenVacDictFields.Combination
             xfield = CStr(.Fields("Combination").Value & vbNullString)
          Case GenVacDictFields.RelatedNames
             RelatedNames = .Fields("Related_Gen_Names") & vbNullString
        End Select
        If Count <> GenVacDictFields.RelatedNames Then
         GenRec = GenRec & sepchar & xfield
         sepchar = FieldSepChar
        End If
      Next Count
      If gn <> "" Then
    '     Debug.Assert InStr(GenRec, "DTaP") = 0
          GenRec = GenRec & sepchar & "N"                ' Put in Ignore Age
          obd_AllGenericVacNames.Add gn, GenRec
          obd_RelatedGenericNames.Add gn, RelatedNames
      Else
        MsgBxARMd Disclaimer & _
               "Invalid Generic Name.  Name is Blank" & vbCrLf & _
               "Generic ID: " & CStr(.Fields("GENNAMEID")), vbInformation + vbOKOnly, "Invalid Generic Name"
      End If
      .MoveNext
    Loop
    '...reset RS to beginning for the future
    .MoveFirst
   
  End If
 End With
 
 Set GenName_RS = Nothing
exitRoutine:
   FrmStatus.LabStatus = FrmStatus.LabStatus & vbCrLf & "Successful"
   FrmStatus.Refresh
 Exit Sub
 
haveerror:
'  Debug.Print Err.Number, Err.Description
'  'Resume
  lLastError = Err.Number
  sLastErrorDesc = "ImmDATA: LoadGenericNames: " & Err.Description
  Err.Raise lLastError, "Imm_Tables: LoadGenericNames: ", sLastErrorDesc
 
End Sub




Public Sub InitializeUserVaccineGroups()

   '  This Routine populates the 2 Dictionaries
   
   '   1. UserGroupsMTF_ID    Inverted Dictionary record consists of Group names
   '                          Key:  MTF_ID
   '   2. PatientGroups       Consists of Group Records defined by enum  GroupRec
   '                          Key:  MTF_ID | GroupName
 
 On Local Error GoTo haveerror
 
 Dim rs As ADODB.Recordset
 Dim mtfid As ADODB.field
 Dim GroupName As ADODB.field
 Dim GenericName As ADODB.field
 Dim LastMTFID As String
 Dim RecMtfID As String
 Dim recfields() As String
 Dim IndDict As Dictionary
 Dim objCmd As CHCSII_Command.CHCSII_Cmd
 Dim objDASsp As ICHCSII_DataAccess.ICHCSII_DAS
 
   FrmStatus.LabStatus = FrmStatus.LabStatus & vbCrLf & "Building Patient Definitions:"
   FrmStatus.Refresh
   
'Begin:SCR #20068 - Brian Mowbray 01/03/2002
'SCR Description: Performance - Open Immunizations fails the ORD6 second req in the field
' Set rs = gobjSQL.OpenRecordset("Select Group_Name, MTF_ID, Max_Age, Min_Age, Consider_Age " & _
'                                "FROM User_Vaccine_Groups Order By MTF_ID,Group_Name", adOpenForwardOnly, adLockReadOnly)


 '>> SCR #35880; 35890;   Developer: Brian Mowbray 03/03/2004 02:29 PM
 '#SP
  If gobjShared.AppMode = modeCDR Then
  
    If objDASsp Is Nothing Then
      Set objDASsp = gobjConn.CHCSII_DAS
    End If
    If objCmd Is Nothing Then
      Set objCmd = New CHCSII_Command.CHCSII_Cmd
    End If
    
    objCmd.SPName = "IMMUNIZATION_PERFORMANCE_PKG.prc_Get_MTF_Vaccine_Groups"
    Call objCmd.AddParam(False, False, "textString", SayMTF_ID())
    Call objCmd.AddParam(True, True, "", "")  '--- for Output Parameter
        
      
    Set rs = objDASsp.OpenRecordsetSP(objCmd)

  
  Else
   
    Set rs = gobjSQL.OpenRecordset("Select Group_Name, MTF_ID, Max_Age, Min_Age, Consider_Age " & _
                                "FROM User_Vaccine_Groups WHERE MTF_ID = '" & SayMTF_ID() & "' Order By MTF_ID,Group_Name", adOpenForwardOnly, adLockReadOnly)
  
  End If
  '#SP


 
 With rs
  If Not .EOF Then
  
   Set mtfid = .Fields("MTF_ID")
   Set GroupName = .Fields("Group_Name")
   
   Do While Not .EOF
     
     If LastMTFID <> mtfid.Value Then
      If LastMTFID <> "" Then
       UserGroupsMTF_ID.Add LastMTFID, RecMtfID
      End If
      LastMTFID = mtfid.Value
      RecMtfID = ""
     End If
     
     '...Build Inverted Group Demis
     RecMtfID = RecMtfID & IIf(RecMtfID = "", "", InvertedTblSepChar) & GroupName.Value
     
     '...Create a Record for this Demis | Group
     PatientGroups.Add .Fields("MTF_ID").Value & vbNullString & WithinFieldSepChar & .Fields("Group_Name").Value & vbNullString, _
                               .Fields("Min_Age").Value & vbNullString & FieldSepChar & _
                               .Fields("Max_Age").Value & vbNullString & FieldSepChar & _
                               .Fields("Consider_Age").Value & FieldSepChar & _
                               "" & FieldSepChar & "" & FieldSepChar & "" & FieldSepChar & CStr(DefIDControl.UserGroup)
                               'Put in Place holder for Status code, Service code, Generic Names and Type of Group respectively
     .MoveNext
   Loop
   
   '...Put last record in
   UserGroupsMTF_ID.Add LastMTFID, RecMtfID
   
  End If
  rs.Close
 End With
 
 Set rs = Nothing
 LastMTFID = ""
 
 '  Read Generic Names for groups
 
 '...Process User_Vaccine_Groups_Vaccine FIRST

'Begin:SCR #20068 - Brian Mowbray 01/03/2002
'SCR Description: Performance - Open Immunizations fails the ORD6 second req in the field
' Set rs = gobjSQL.OpenRecordset("Select GenericName,Group_Name,MTF_ID from User_Vaccine_Groups_Vaccine Order by MTF_ID, Group_Name", adOpenForwardOnly, adLockReadOnly)
 
 '>> SCR #35880; 35890;   Developer: Brian Mowbray 03/03/2004 02:29 PM
 '#SP
  If gobjShared.AppMode = modeCDR Then
  
    If objDASsp Is Nothing Then
      Set objDASsp = gobjConn.CHCSII_DAS
    End If
    
    Set objCmd = New CHCSII_Command.CHCSII_Cmd
    
    objCmd.SPName = "IMMUNIZATION_PERFORMANCE_PKG.prc_Get_MTF_Vac_Grps_Vaccines"
    Call objCmd.AddParam(False, False, "textString", SayMTF_ID())
    Call objCmd.AddParam(True, True, "", "")  '--- for Output Parameter
        
      
    Set rs = objDASsp.OpenRecordsetSP(objCmd)

  
  Else
   
      Set rs = gobjSQL.OpenRecordset("Select GenericName,Group_Name,MTF_ID from User_Vaccine_Groups_Vaccine WHERE MTF_ID = '" & SayMTF_ID() & "' Order by MTF_ID, Group_Name", adOpenForwardOnly, adLockReadOnly)
  
  End If
  '#SP
 
 With rs
   
   If Not .EOF Then
   
    Set mtfid = .Fields("MTF_ID")
    Set GroupName = .Fields("Group_Name")
    Set GenericName = .Fields("GenericName")
    
     Do While Not .EOF
     
        If LastMTFID <> mtfid.Value & WithinFieldSepChar & GroupName.Value Then
         If LastMTFID <> "" Then
          '...Add Generic Names to this record
          recfields = Split(PatientGroups.Item(LastMTFID), FieldSepChar)
          '...IF there isn't garbage in the file save the data
          If UBound(recfields) <> -1 Then
            recfields(grouprec.GenericNames) = RecMtfID
            PatientGroups.Item(LastMTFID) = Join(recfields, FieldSepChar)
          Else
         MsgBxARMd Disclaimer & _
                "The following data exists in File: " & "User_Vaccine_Groups_Vaccine" & vbCrLf & vbCrLf & _
                "Group_Name:  " & IIf(Len(Mid$(LastMTFID, InStr(LastMTFID, WithinFieldSepChar) + 1)) = 0, "VOID", Mid$(LastMTFID, InStr(LastMTFID, WithinFieldSepChar) + 1)) & vbCrLf & _
                "MTF_ID Key:  " & IIf(Len(Mid$(LastMTFID, 1, InStr(LastMTFID, WithinFieldSepChar) - 1)) = 0, "VOID", Mid$(LastMTFID, 1, InStr(LastMTFID, WithinFieldSepChar) - 1)) & vbCrLf & vbCrLf & _
                "but does not exist in File: " & "User_Vaccine_Groups", _
                vbInformation + vbOKOnly, "Corrupt File"
          End If
         End If
         LastMTFID = Trim$(mtfid.Value) & WithinFieldSepChar & Trim$(GroupName.Value)
         RecMtfID = ""
        End If
        
         '...Build Inverted Group Demis
         RecMtfID = RecMtfID & IIf(LenB(RecMtfID) = 0, "", InvertedTblSepChar) & GenericName.Value
         .MoveNext
         
     Loop
     
     '...Add Generic Names to this record
     recfields = Split(PatientGroups.Item(LastMTFID), FieldSepChar)
     
     If UBound(recfields) <> -1 Then
         recfields(grouprec.GenericNames) = RecMtfID
         PatientGroups.Item(LastMTFID) = Join(recfields, FieldSepChar)
     Else
         MsgBxARMd Disclaimer & _
                "The following data exists in File: " & "User_Vaccine_Groups_Vaccine" & vbCrLf & vbCrLf & _
                "MTF_ID Key:  " & IIf(Len(Mid$(LastMTFID, InStr(LastMTFID, WithinFieldSepChar) + 1)) = 0, "VOID", Mid$(LastMTFID, InStr(LastMTFID, WithinFieldSepChar) + 1)) & vbCrLf & _
                "Group_Name:  " & IIf(Len(Mid$(LastMTFID, 1, InStr(LastMTFID, WithinFieldSepChar) - 1)) = 0, "VOID", Mid$(LastMTFID, 1, InStr(LastMTFID, WithinFieldSepChar) - 1)) & vbCrLf & vbCrLf & _
                "but does not exist in File: " & "User_Vaccine_Groups", _
                vbInformation + vbOKOnly, "Corrupt File"
     End If
     
   End If
   
   .Close
   
 End With
 
 Set rs = Nothing
 Set mtfid = Nothing
 Set GroupName = Nothing
 Set GenericName = Nothing
 
exitRoutine:

   FrmStatus.LabStatus = FrmStatus.LabStatus & vbCrLf & "Successful"
   FrmStatus.Refresh

 Exit Sub
 
haveerror:
 ErrMsgBox "Imm_Tables_InitializeUserVaccineGroups", Error$, Err
 Resume exitRoutine
 Resume 'Debug
End Sub
Public Sub InitIalizeDefaultVaccineGroups()

   '  This Routine populates the Dictionary
   '  PatientGroups with records from
   '  TBLVaccine_Groups (The Default Group Table)
 
 On Local Error GoTo haveerror
 
 Dim rs As ADODB.Recordset
 Dim sql As String
 Dim LastMTFID As String
 Dim Prefix As String
 Dim recfields() As String
 Dim GroupName As ADODB.field
 Dim GenericName As ADODB.field
 Dim RecMtfID As String
 Dim objCmd As CHCSII_Command.CHCSII_Cmd
 Dim objDASsp As ICHCSII_DataAccess.ICHCSII_DAS
 
      
 
   FrmStatus.LabStatus = FrmStatus.LabStatus & vbCrLf & "Building Default Groups:"
   FrmStatus.Refresh
   
 Prefix = "DEFAULT" & WithinFieldSepChar
 
  '>> SCR #35880; 35890;   Developer: Brian Mowbray 03/03/2004 02:29 PM
  '#SP
  If gobjShared.AppMode = modeCDR Then
  
    Set objDASsp = gobjConn.CHCSII_DAS
    Set objCmd = New CHCSII_Command.CHCSII_Cmd
        
    objCmd.SPName = "IMMUNIZATION_PERFORMANCE_PKG.prc_Get_Vaccine_Groups"
    Call objCmd.AddParam(True, True, "", "")  '--- for Output Parameter
        
      
    Set rs = objDASsp.OpenRecordsetSP(objCmd)

  
  Else
   
   Set rs = gobjSQL.OpenRecordset("Select Group_Name, Min_Age, Max_Age, Status_Code, Service_Code, Consider_Age " & _
                                  "FROM TBLVaccine_Groups", adOpenForwardOnly, adLockReadOnly)
  End If
  '#SP
 
   With rs
   Do While Not .EOF
   
     PatientGroups.Add Prefix & .Fields("Group_Name").Value & vbNullString & vbNullString, _
                               .Fields("Min_Age").Value & vbNullString & vbNullString & FieldSepChar & _
                               .Fields("Max_Age").Value & vbNullString & vbNullString & FieldSepChar & _
                               .Fields("Consider_Age").Value & FieldSepChar & _
                               UCase$(.Fields("Status_code").Value & vbNullString) & FieldSepChar & _
                               UCase$(.Fields("Service_Code").Value & vbNullString) & FieldSepChar & _
                               " " & FieldSepChar & _
                               CStr(DefIDControl.UnitDefault)     'Put in Place holder for Generic Names and indicate Unit Default
     .MoveNext
   Loop
   rs.Close
 End With
 
 Set rs = Nothing
 LastMTFID = ""
 
   FrmStatus.LabStatus = FrmStatus.LabStatus & vbCrLf & "Successful"
   FrmStatus.Refresh
   
 '  Read Generic Names for groups
 
   FrmStatus.LabStatus = FrmStatus.LabStatus & vbCrLf & "Building Generic Names / Inversion:"
   FrmStatus.Refresh
   
 '...Process TBLVaccine_Groups_Vaccine FIRST
   '>> SCR #35880; 35890;   Developer: Brian Mowbray 03/03/2004 02:29 PM
   '#SP
  If gobjShared.AppMode = modeCDR Then
  
    If objDASsp Is Nothing Then
      Set objDASsp = gobjConn.CHCSII_DAS
    End If
    
    Set objCmd = New CHCSII_Command.CHCSII_Cmd
    
    objCmd.SPName = "IMMUNIZATION_PERFORMANCE_PKG.prc_Get_Vac_Groups_Vaccines"
    Call objCmd.AddParam(True, True, "", "")  '--- for Output Parameter
        
      
    Set rs = objDASsp.OpenRecordsetSP(objCmd)

  
  Else
   
    Set rs = gobjSQL.OpenRecordset("Select GenericName,Group_Name from TBLVaccine_Groups_Vaccine Order by Group_Name", adOpenForwardOnly, adLockReadOnly)
  
  End If
  '#SP
 
 
 
 With rs
   
   If Not .EOF Then
   
    Set GroupName = .Fields("Group_Name")
    Set GenericName = .Fields("GenericName")
    
     Do While Not .EOF
     
     If LastMTFID <> GroupName.Value Then
      If LastMTFID <> "" Then
       '...Add Generic Names to this record
       recfields = Split(PatientGroups.Item(Prefix & LastMTFID), FieldSepChar)
       '...Is there isn't garbage in the file save the data
       If UBound(recfields) <> -1 Then
         recfields(grouprec.GenericNames) = RecMtfID
         PatientGroups.Item(Prefix & LastMTFID) = Join(recfields, FieldSepChar)
       Else
         MsgBxARMd Disclaimer & _
                "Invalid Data in File: " & "TBLVaccine_Groups_Vaccine" & vbCrLf & vbCrLf & _
                "Group_Name:  " & IIf(Len(LastMTFID) = 0, "VOID", LastMTFID), _
                vbInformation + vbOKOnly, "Corrupt File"
       End If
      End If
      LastMTFID = GroupName.Value
      RecMtfID = ""
     End If
     
      '...Build Inverted Group Demis
      RecMtfID = RecMtfID & IIf(RecMtfID = "", "", InvertedTblSepChar) & GenericName.Value
      .MoveNext
     Loop
     '...Add Generic Names to this record
     recfields = Split(PatientGroups.Item(Prefix & LastMTFID), FieldSepChar)
     
     If UBound(recfields) <> -1 Then
         recfields(grouprec.GenericNames) = RecMtfID
         PatientGroups.Item(Prefix & LastMTFID) = Join(recfields, FieldSepChar)
     Else
         MsgBxARMd Disclaimer & _
                "Invalid Data in File: " & "TBLVaccine_Groups_Vaccine" & vbCrLf & vbCrLf & _
                "Group_Name:  " & IIf(Len(LastMTFID) = 0, "VOID", LastMTFID), _
                vbInformation + vbOKOnly, "Corrupt File"
     End If
     
   End If
   
   .Close
   
 End With
 
 Set rs = Nothing
 
   FrmStatus.LabStatus = FrmStatus.LabStatus & vbCrLf & "Successful"
   FrmStatus.Refresh
   
exitRoutine:

 Exit Sub
 
haveerror:
 ErrMsgBox "Imm_Tables_InitializeDefaultVaccineGroups", Error$, Err
 Resume exitRoutine
 Resume '
End Sub

Private Sub Class_Initialize()
  '--- COMMON TO MULTIPLE SCR-ARMs: MSolano 9/13/01 SCR-ARM-16110, 16130
  '--- Init routine should be called from first top level creatable coclass
  '--- It can be called multiple times, but will actually run only once.
  '--- Instrument all creatable classes in case dll's multiple entry
  '--- points are actually used
  Call ARM_InitAllTransactionsForImmunizations
 '...Dictionaries used to Indicate Universe of
 '...Generic names to choose from
 Set obd_Vaccines = New Dictionary
 Set obd_VacIDToGenName = New Dictionary
 Set obd_GenNametoVacIDs = New Dictionary
 Set obd_AllGenericVacNames = New Dictionary
 Set obd_RelatedGenericNames = New Dictionary
 Set UserGroupsMTF_ID = New Dictionary
 Set PatientGroups = New Dictionary
 Set obd_VaccineToVacID = New Dictionary
 Set obd_VaccinesSolvedBy = New Dictionary
 ' Begin:SCR #19803;   Developer: Brian Mowbray 05/10/2002 01:00 PM
 Set obd_MfgCodeToMfgName = New Dictionary
 Set obd_MfgNameToMfgCode = New Dictionary
 ' End: SCR #19803;
 Set obd_VIS_Versions = New Dictionary
 Set UserGroupsUnitMTF_ID = New Dictionary

 UserGroupsUnitMTF_ID.CompareMode = TextCompare
 obd_VaccinesSolvedBy.CompareMode = TextCompare
 obd_VaccineToVacID.CompareMode = TextCompare
 UserGroupsMTF_ID.CompareMode = TextCompare
 PatientGroups.CompareMode = TextCompare
 obd_AllGenericVacNames.CompareMode = TextCompare
 obd_RelatedGenericNames.CompareMode = TextCompare
 obd_Vaccines.CompareMode = TextCompare
 obd_VacIDToGenName.CompareMode = TextCompare
 obd_GenNametoVacIDs.CompareMode = TextCompare
 ' Begin:SCR #19803;   Developer: Brian Mowbray 05/10/2002 01:00 PM
 obd_MfgCodeToMfgName.CompareMode = TextCompare
 obd_MfgNameToMfgCode.CompareMode = TextCompare
 ' End: SCR #19803;
 
 obd_VIS_Versions.CompareMode = TextCompare
  
  With FrmStatus
   .Left = (Screen.Width - .Width) / 2
   .Top = (Screen.Height - .Height) / 2
   .LabStatus = "Initialization Progress:" & vbCrLf & vbCrLf
   .Show
  End With
  
  InitIalizeDefaultVaccineGroups
  InitializeUserVaccineGroups
  LoadGenericNames
  LoadVaccines
  ' Beginc:SCR #19803;   Developer: Brian Mowbray 05/10/2002 01:00 PM
  LoadMfgs
  ' End: SCR #19803;
  StatusAck = False
  Unload FrmStatus
  
  
  '...Debug Purposes ONly
  '  Test_Display_Tables
  '  Dump_To_Disk
End Sub



Public Property Get Vaccine_Record(ByVal VacKey As String) As String()
 '...Get a specific Vaccine Record
 If obd_Vaccines.Exists(VacKey) Then
   Vaccine_Record = Split(obd_Vaccines.Item(VacKey), GroupSepChar)
  Else
    Vaccine_Record = SetUboundToMinusOne
  End If
End Property
Public Property Get Vaccine_ExpertCriteria(ByVal VacKey As String) As String
 '...Get a specific Vaccine Criteria to use
 Dim Groups() As String
 If obd_Vaccines.Exists(VacKey) Then
   Groups = Split(obd_Vaccines.Item(VacKey), GroupSepChar)
   Vaccine_ExpertCriteria = Groups(2)
 End If
End Property
Public Property Get Vaccine_Series_String(ByVal VacKey As String) As String
 '...Get a specific Vaccine Series Schedule
 Dim Groups() As String
 If obd_Vaccines.Exists(VacKey) Then
   Groups = Split(obd_Vaccines.Item(VacKey), GroupSepChar)
   Vaccine_Series_String = Groups(1)
 End If
End Property
Public Property Get Vaccine_Series(ByVal VacKey As String) As String()
 '...Get a specific Vaccine Series Schedule
 Dim Groups() As String
 If obd_Vaccines.Exists(VacKey) Then
   Groups = Split(obd_Vaccines.Item(VacKey), GroupSepChar)
   Vaccine_Series = Split(Groups(1), SeriesSepChar)
 Else
    Vaccine_Series = SetUboundToMinusOne
 End If
End Property

Public Property Get Vaccine_BasicData(ByVal VacKey As String) As String()
 '...Get a specific Vaccine's basic data
 Dim Groups() As String
 If obd_Vaccines.Exists(VacKey) Then
   Groups = Split(obd_Vaccines.Item(VacKey), GroupSepChar)
   Vaccine_BasicData = Split(Groups(0), FieldSepChar)
 Else
    Vaccine_BasicData = SetUboundToMinusOne
 End If
End Property


Public Property Get VacIDsFromGenName(ByVal gn As String) As String()
 '...Get all Vac IDs for a Generic Name
 If obd_GenNametoVacIDs.Exists(gn) Then
   VacIDsFromGenName = Split(obd_GenNametoVacIDs.Item(gn), InvertedTblSepChar)
 Else
   VacIDsFromGenName = SetUboundToMinusOne
 End If
End Property


Public Property Get VacIdKeys() As Variant()
 'Return all VacIDs
   VacIdKeys = obd_VacIDToGenName.Keys
End Property


Public Property Get VacIDToGenNameRecord(ByVal VacKey As String) As String()
 '...Load Generic Name for given a vacID
 If obd_VacIDToGenName.Exists(VacKey) Then
   VacIDToGenNameRecord = Split(obd_VacIDToGenName.Item(VacKey), FieldSepChar)
 Else
   VacIDToGenNameRecord = SetUboundToMinusOne
 End If
End Property

Private Sub Class_Terminate()

'************************************************
  
  ' The Following Variables are global variables
  ' that have no way of guaranteeing that they
  ' have been set to nothing.  So the solution is
  ' to use this class to get rid of them.  The
  ' Unload in FrmAdmin_Main could have also been used
  
  ' This problem will be resolved when functionality
  ' is taken out of FrmAdmin_Main and put into sepearate
  ' forms

 Set ModifiedVaccineSeries = Nothing
 Set OrigVaccineSeriesDel = Nothing
 
'************************************************
 Set obd_VaccinesSolvedBy = Nothing
 Set obd_Vaccines = Nothing
 Set obd_VacIDToGenName = Nothing
 Set obd_GenNametoVacIDs = Nothing
 Set obd_AllGenericVacNames = Nothing
 Set obd_RelatedGenericNames = Nothing
 Set UnitGroupReqForPatient = Nothing
 Set UserGroupsMTF_ID = Nothing
 Set PatientGroups = Nothing
 Set obd_VaccineToVacID = Nothing
 Set UserGroupsUnitMTF_ID = Nothing
 Set obd_MfgCodeToMfgName = Nothing
 Set obd_MfgNameToMfgCode = Nothing
 Set obd_VIS_Versions = Nothing
End Sub



Public Sub LoadRequiredUnitGroups(ByVal xUnit As String, ByVal Service As String, ByVal Status As String, ByVal vDOB As Variant)
 
 '...Get Default Information Based on Service, Status and maybe AGE
 
 On Local Error GoTo haveerror
 
 Dim DefaultKeys() As Variant
 Dim DefKey As Variant
 
 Dim GrpRecord() As String
 Dim Statuses() As String
 Dim InpStatuses() As String
 Dim ValidGroup As Boolean
 Dim MinAge As Variant
 Dim MaxAge As Variant
 Dim Inpstat As Variant
 Dim FldStat As Variant
 Dim foundit As Boolean
 Dim GroupsForUnitMTF As String
 
'------------------------------------------------------------------

    'FUTURE POTENTIAL OPTIMIZATION
    
' if LastUnitUsed = xUnit and LastService = Service then

'------------------------------------------------------------------

 If LenB(Status) <> 0 Then
  InpStatuses() = Split(UCase$(Status), WithinFieldSepChar)
 Else
  InpStatuses = SetUboundToMinusOne
 End If
 
 DefaultKeys = PatientGroups.Keys

 '...Create New Groups required for this unit
 Set UnitGroupReqForPatient = Nothing
 Set UnitGroupReqForPatient = New Dictionary
 Set UserGroupsUnitMTF_ID = Nothing
 Set UserGroupsUnitMTF_ID = New Dictionary
 
 For Each DefKey In DefaultKeys

   If Mid$(DefKey, 1, 8) = "DEFAULT" & WithinFieldSepChar Then

    ValidGroup = False
    GrpRecord = Split(PatientGroups.Item(DefKey), FieldSepChar)
    
    '...Load Statuses if any
    If LenB(GrpRecord(grouprec.Status)) <> 0 Then
      Statuses = Split(UCase$(GrpRecord(grouprec.Status)), WithinFieldSepChar)
    Else
      Statuses = SetUboundToMinusOne
    End If
    
    '...Check Status First
    If UBound(InpStatuses) <> -1 And UBound(Statuses) <> -1 Then
    
         foundit = False
         '...See If Status Match
         For Each FldStat In Statuses
           For Each Inpstat In InpStatuses

             If Inpstat = FldStat Then
               foundit = True
               Exit For
             End If

           Next Inpstat

           If foundit Then Exit For

         Next FldStat
         ValidGroup = foundit
    
    End If
    
    '...Is this the right service and we have a valid group
    If ValidGroup _
       And ( _
            (GrpRecord(grouprec.ServiceCode) = Service And LenB(Service) <> 0) _
              Or _
            (GrpRecord(grouprec.ServiceCode) = "") _
           ) Then
     ' Yes

     If IsDate(vDOB) Then
       '...Do we need to consider age
       If GrpRecord(grouprec.ConsiderAge) <> "0" Then
         '...Calculate age range
         MinAge = CalcAgeOrTime(CDate(vDOB), GrpRecord(grouprec.MinAge))
         MaxAge = CalcAgeOrTime(CDate(vDOB), GrpRecord(grouprec.MaxAge))
         '...Are we within limits
         If Now < MinAge Or Now > MaxAge Then ValidGroup = False
       End If
       
     ElseIf GrpRecord(grouprec.ConsiderAge) = "Y" Then
       '...Need to consider age but no age given
       ValidGroup = False
       
     End If

    Else

     '...Not the right situation
     ValidGroup = False

    End If

    If ValidGroup Then
     GrpRecord(grouprec.idcontrol) = CStr(Val(GrpRecord(grouprec.idcontrol)) Or DefIDControl.RequiredGroup)
     UnitGroupReqForPatient.Add DefKey, Join(GrpRecord, FieldSepChar)
    End If

   End If

 Next DefKey
 
 '...Was a UIC passed to us
 If xUnit <> "" Then
   ' Yes
   '...Get Additional Unit information based on
   '...Data defined in the IMM_UIC_Unit Table
     On Local Error GoTo haveerror
    
     ' This Routine Builds a Dictionary of Group Records required by
     ' the current Unit working with
    
     Dim rs As ADODB.Recordset
     Dim Group As ADODB.field
     Dim Default As String
     
'Tracker #53478  Developer: MAG 05/07/2004
'     Static LastUnit As String
'     Static LastDemis As String
'
'     '...May gain some economy here
'     If LastUnit = xUnit And LastDemis = gsMTF_ID Then Exit Sub
'
'     '...Create New Key for Unit
'     '...Used only in above statement
'     LastUnit = xUnit
'     LastDemis = gsMTF_ID
     
     '...Create Key for Default Groups
     Default = "DEFAULT" & WithinFieldSepChar
     
     If gsMTF_ID <> vbNullString Then
       '<< SCR #42779;   Developer: Brian Mowbray 10/21/2003 04:18 PM
       Set rs = gobjSQL.OpenRecordset("Select Group_Name from IMM_Unit_Group where Unit = '" & ValidateDBChar(xUnit) & "' AND MTF_ID = '" & gsMTF_ID & "'", adOpenForwardOnly, adLockReadOnly)
       With rs
        If Not .EOF Then
         Set Group = .Fields("Group_Name")
         Do While Not .EOF
         
           GroupsForUnitMTF = GroupsForUnitMTF & IIf(LenB(GroupsForUnitMTF) = 0, "", InvertedTblSepChar) & Group.Value
           
           If Not UnitGroupReqForPatient.Exists(Default & Group.Value) Then
           
                '...Is this a Default Group
                If PatientGroups.Exists(Default & Group.Value) Then
                  ' this Group is a Default Group for the unit
                  UnitGroupReqForPatient.Add Group.Value, PatientGroups.Item(Default & Group.Value)
                  
                '...Is this a User Group assigned to this Demis
                ElseIf PatientGroups.Exists(gsMTF_ID & WithinFieldSepChar & Group.Value) Then
                  ' have a user defined Group for the Unit
                  UnitGroupReqForPatient.Add Group.Value, PatientGroups.Item(gsMTF_ID & WithinFieldSepChar & Group.Value)
                  
                '...this group is Invalid <Corrupt Database>
                Else
                  MsgBxARMd "This Unit has no Groups assigned", vbInformation + vbOKOnly, "For Your Information"
                End If
                
           End If
           .MoveNext
         Loop
         .Close
         UserGroupsUnitMTF_ID.Add xUnit & WithinFieldSepChar & gsMTF_ID, GroupsForUnitMTF
         
        End If
       End With
       Set rs = Nothing
     End If
 End If
 
 Exit Sub
 
haveerror:
'  Debug.Print Err.Number, Err.Description
'  Resume     ' TJNXXX
  lLastError = Err.Number
  sLastErrorDesc = "Imm_Tables: LoadRequiredUnitGroups: " & Err.Description
  Err.Raise lLastError, "Imm_Tables: LoadRequiredUnitGroups: ", sLastErrorDesc
 
 
End Sub

Public Sub GetUnitGroups(xDict As Dictionary)

 '  Fill a Patient's Group dictionary with Required groups from Unit
 '  Since Units groups override User groups this will always populate
 '  With items from UnitGroupReqForPatient
 
 Dim GroupKey As Variant
 
 For Each GroupKey In UnitGroupReqForPatient
   If xDict.Exists(GroupKey) Then
    xDict.Item(GroupKey) = UnitGroupReqForPatient.Item(GroupKey)
   Else
    xDict.Add GroupKey, UnitGroupReqForPatient.Item(GroupKey)
   End If
 Next GroupKey

End Sub

'<< SCR #25155;   Developer: Brian Mowbray 06/27/2002 02:02 PM
Public Function GetVIS_Version() As String()
 Dim sTemp() As String
 Dim a, i As Integer
   
 ReDim sTemp(5)
 a = obd_VIS_Versions.Keys
 
 For i = 0 To obd_VIS_Versions.Count - 1
    If i > UBound(sTemp) Then ReDim Preserve sTemp(UBound(sTemp) + 1)
    sTemp(i) = a(i) & " (" & obd_VIS_Versions.Item(a(i)) & ")"
 Next i

 GetVIS_Version = sTemp

End Function

Public Function UserGroupRecord(ByVal xkey As String) As String
  If PatientGroups.Exists(xkey) Then UserGroupRecord = PatientGroups.Item(xkey)
End Function

Public Sub Test_Display_Tables()

  '  Used to debug all the tables built by Imm_Tables
 On Local Error GoTo haveerror
 Dim xrec() As String
 Dim xShaping As String
 Dim Group As Variant
 
 Debug.Print "-------------------"
 Debug.Print "Related Vaccines   "
 Debug.Print "-------------------"
 Debug.Print
 Debug.Print "ID            Vaccine       Related Vaccines"
 Debug.Print
 For Each Group In obd_Vaccines
 
   xrec = Split(obd_Vaccines.Item(Group), GroupSepChar)
   xrec = Split(xrec(0), FieldSepChar)
   Debug.Print xrec(VacDictGrp1Fields.VaccineID), xrec(VacDictGrp1Fields.Vaccine), ;
   If xrec(VacDictGrp1Fields.EffectsVacID) <> vbNullString Then
    Debug.Print xrec(VacDictGrp1Fields.EffectsVacID)
   Else
    Debug.Print
   End If
   
 Next Group
 
 Debug.Print "-------------------"
 Debug.Print "Table PatientGroups"
 Debug.Print "-------------------"
 Debug.Print
 Debug.Print "MinAge    MaxAge    CA        Status    SvcCode   Generic_Names"
 Debug.Print
 For Each Group In PatientGroups
 
   xrec = Split(PatientGroups.Item(Group), FieldSepChar)
   Debug.Print Group
   If UBound(xrec) <> -1 Then
    Debug.Print xrec(grouprec.MinAge) & Space$(5 + Len(xrec(grouprec.MinAge))) & _
                xrec(grouprec.MaxAge) & Space$(5 + Len(xrec(grouprec.MaxAge))) & _
                xrec(grouprec.ConsiderAge) & Space$(5 + Len(xrec(grouprec.ConsiderAge))) & _
                xrec(grouprec.Status) & Space$(5 + Len(xrec(grouprec.Status))) & _
                xrec(grouprec.ServiceCode) & Space$(5 + Len(xrec(grouprec.ServiceCode))) & _
                xrec(grouprec.GenericNames)
   Else
    Debug.Print "Invalid Group Name"
   End If
 Next Group
 
 Debug.Print
 Debug.Print "----------------------"
 Debug.Print "Table UserGroupsMTF_ID"
 Debug.Print "----------------------"
 Debug.Print
 For Each Group In UserGroupsMTF_ID
    Debug.Print Group
    Debug.Print UserGroupsMTF_ID.Item(Group)
 Next Group
 
 Debug.Print
 Debug.Print "----------------------------"
 Debug.Print "Table obd_AllGenericVacNames"
 Debug.Print "----------------------------"
 Debug.Print "GenericName       GenericID    MinAge        MaxAge        xType   Cycle    range    DC    RO    C   IA  Related"
 Debug.Print
 For Each Group In obd_AllGenericVacNames
   If obd_RelatedGenericNames.Exists(Group) Then
     xShaping = obd_RelatedGenericNames.Item(Group)
   Else
    xShaping = ""
   End If
   xrec = Split(Mid$(obd_AllGenericVacNames.Item(Group), 2, Len(obd_AllGenericVacNames.Item(Group)) - 2), FieldSepChar)
   If UBound(xrec) <> -1 Then
    Debug.Print IIf(Len(xrec(GenVacDictFields.GenericName)) > 17, Left$(xrec(GenVacDictFields.GenericName), 17), xrec(GenVacDictFields.GenericName)) & Space$(18 - Len(IIf(Len(xrec(GenVacDictFields.GenericName)) > 17, Left$(xrec(GenVacDictFields.GenericName), 17), xrec(GenVacDictFields.GenericName)))) & _
                IIf(Len(xrec(GenVacDictFields.GenericID)) > 12, Left$(xrec(GenVacDictFields.GenericID), 12), xrec(GenVacDictFields.GenericID)) & Space$(13 - Len(IIf(Len(xrec(GenVacDictFields.GenericID)) > 12, Left$(xrec(GenVacDictFields.GenericID), 12), xrec(GenVacDictFields.GenericID)))) & _
                IIf(Len(xrec(GenVacDictFields.MinAge)) > 13, Left$(xrec(GenVacDictFields.MinAge), 13), xrec(GenVacDictFields.MinAge)) & Space$(14 - Len(IIf(Len(xrec(GenVacDictFields.MinAge)) > 13, Left$(xrec(GenVacDictFields.MinAge), 13), xrec(GenVacDictFields.MinAge)))) & _
                IIf(Len(xrec(GenVacDictFields.MaxAge)) > 13, Left$(xrec(GenVacDictFields.MaxAge), 13), xrec(GenVacDictFields.MaxAge)) & Space$(14 - Len(IIf(Len(xrec(GenVacDictFields.MaxAge)) > 13, Left$(xrec(GenVacDictFields.MaxAge), 13), xrec(GenVacDictFields.MaxAge)))) & _
                IIf(Len(xrec(GenVacDictFields.xtype)) > 7, Left$(xrec(GenVacDictFields.xtype), 7), xrec(GenVacDictFields.xtype)) & Space$(8 - Len(IIf(Len(xrec(GenVacDictFields.xtype)) > 7, Left$(xrec(GenVacDictFields.xtype), 7), xrec(GenVacDictFields.xtype)))) & _
                IIf(Len(xrec(GenVacDictFields.Cycle)) > 8, Left$(xrec(GenVacDictFields.Cycle), 8), xrec(GenVacDictFields.Cycle)) & Space$(9 - Len(IIf(Len(xrec(GenVacDictFields.Cycle)) > 8, Left$(xrec(GenVacDictFields.Cycle), 8), xrec(GenVacDictFields.Cycle)))) & _
                IIf(Len(xrec(GenVacDictFields.Range)) > 8, Left$(xrec(GenVacDictFields.Range), 8), xrec(GenVacDictFields.Range)) & Space$(9 - Len(IIf(Len(xrec(GenVacDictFields.Range)) > 8, Left$(xrec(GenVacDictFields.Range), 8), xrec(GenVacDictFields.Range)))) & _
                IIf(Len(xrec(GenVacDictFields.Discontinued)) > 5, Left$(xrec(GenVacDictFields.Discontinued), 5), xrec(GenVacDictFields.Discontinued)) & Space$(6 - Len(IIf(Len(xrec(GenVacDictFields.Discontinued)) > 5, Left$(xrec(GenVacDictFields.Discontinued), 5), xrec(GenVacDictFields.Discontinued)))) & _
                IIf(Len(xrec(GenVacDictFields.RecmdDTROnly)) > 5, Left$(xrec(GenVacDictFields.RecmdDTROnly), 5), xrec(GenVacDictFields.RecmdDTROnly)) & Space$(6 - Len(IIf(Len(xrec(GenVacDictFields.RecmdDTROnly)) > 5, Left$(xrec(GenVacDictFields.RecmdDTROnly), 5), xrec(GenVacDictFields.RecmdDTROnly)))) & _
                IIf(Len(xrec(GenVacDictFields.Combination)) > 3, Left$(xrec(GenVacDictFields.Combination), 3), xrec(GenVacDictFields.Combination)) & Space$(4 - Len(IIf(Len(xrec(GenVacDictFields.Combination)) > 3, Left$(xrec(GenVacDictFields.Combination), 3), xrec(GenVacDictFields.Combination)))) & _
                xrec(GenVacDictFields.IgnoreAge) & Space$(3) & xShaping
   Else
    Debug.Print "Invalid Group Name"
   End If
 Next Group
 
 Debug.Print
 Debug.Print "------------------------"
 Debug.Print "Table obd_VacIDToGenName"
 Debug.Print "------------------------"
 Debug.Print
 Debug.Print "VacID  GenericName            Vaccine        Min_Age     Max_Age     VacIDsEffected"
 Debug.Print
 For Each Group In obd_VacIDToGenName
   xrec = Split(obd_VacIDToGenName.Item(Group), FieldSepChar)
   If UBound(xrec) <> -1 Then
   Debug.Print IIf(Len(Group) > 6, Left$(Group, 6), Group) & Space$(7 - Len(IIf(Len(Group) > 6, Left$(Group, 6), Group))) & _
                IIf(Len(xrec(VIDToGenName.GenericName)) > 22, Left$(xrec(VIDToGenName.GenericName), 22), xrec(VIDToGenName.GenericName)) & Space$(23 - Len(IIf(Len(xrec(VIDToGenName.GenericName)) > 22, Left$(xrec(VIDToGenName.GenericName), 22), xrec(VIDToGenName.GenericName)))) & _
                IIf(Len(xrec(VIDToGenName.Vaccine)) > 14, Left$(xrec(VIDToGenName.Vaccine), 14), xrec(VIDToGenName.Vaccine)) & Space$(15 - Len(IIf(Len(xrec(VIDToGenName.Vaccine)) > 14, Left$(xrec(VIDToGenName.Vaccine), 14), xrec(VIDToGenName.Vaccine)))) & _
                IIf(Len(xrec(VIDToGenName.MinAge)) > 11, Left$(xrec(VIDToGenName.MinAge), 11), xrec(VIDToGenName.MinAge)) & Space$(12 - Len(IIf(Len(xrec(VIDToGenName.MinAge)) > 11, Left$(xrec(VIDToGenName.MinAge), 11), xrec(VIDToGenName.MinAge)))) & _
                IIf(Len(xrec(VIDToGenName.MaxAge)) > 11, Left$(xrec(VIDToGenName.MaxAge), 11), xrec(VIDToGenName.MaxAge)) & Space$(12 - Len(IIf(Len(xrec(VIDToGenName.MaxAge)) > 11, Left$(xrec(VIDToGenName.MaxAge), 11), xrec(VIDToGenName.MaxAge)))) & _
                xrec(VIDToGenName.VacIdsEffected)
   Else
    Debug.Print "Invalid VacID"
   End If
 Next Group
 
 Debug.Print
 Debug.Print "-------------------------"
 Debug.Print "Table obd_GenNametoVacIDs"
 Debug.Print "-------------------------"
 Debug.Print
 Debug.Print "GenericName                    VacIDs"
 Debug.Print
 For Each Group In obd_GenNametoVacIDs
    Debug.Print IIf(Len(Group) > 30, Left$(Group, 30), Group) & Space$(31 - Len(IIf(Len(Group) > 30, Left$(Group, 30), Group))) & _
                obd_GenNametoVacIDs.Item(Group)
    Debug.Print
 Next Group
 
exitRoutine:

 Exit Sub
 
haveerror:
 ErrMsgBox "Imm_Tables_Test_Display_Tables", Error$, Err
 Resume exitRoutine
 Resume

End Sub

Public Property Get Get_All_Generic_Keys() As Variant()
  Get_All_Generic_Keys = obd_AllGenericVacNames.Keys
End Property


Public Property Get Get_A_Generic_Name_Record(ByVal GenericName As String) As String()
  '...Load a Generic Record
  If obd_AllGenericVacNames.Exists(GenericName) Then
    Get_A_Generic_Name_Record = Split(obd_AllGenericVacNames(GenericName), FieldSepChar)
  Else
    Get_A_Generic_Name_Record = SetUboundToMinusOne
  End If
End Property


Public Sub Add_VacID_To_Generic_Name(ByVal VacID As String, ByVal GenericName As String)

  '  Add VacID to Generic Name
  
   If GenericName = "" Or VacID = "" Then Exit Sub
   
   If obd_GenNametoVacIDs.Exists(GenericName) Then
   
    Dim vacidrec As String
    vacidrec = obd_GenNametoVacIDs.Item(PriorGenericName)
    obd_GenNametoVacIDs.Item(PriorGenericName) = vacidrec & IIf(vacidrec = "", "", InvertedTblSepChar) & VacID
    
   Else
    obd_GenNametoVacIDs.Add GenericName, VacID
    
   End If
End Sub
Public Sub ReMove_VacID_From_GenericName(ByVal VacID As String, ByVal GenericName As String)
 
  ' This removes a Vaccine ID from the inverted generic Table

  Dim VacRec() As String
  Dim sepchar As String
  Dim vacidrec As String
  Dim Count As Integer
  
   If GenericName = "" Or VacID = "" Then Exit Sub

  
   '...Load record
   VacRec = Split(obd_GenNametoVacIDs.Item(GenericName), InvertedTblSepChar)
   sepchar = ""
   
   For Count = 0 To UBound(VacRec)
    '...If this isn't vacid then put it back in again
    If VacRec(Count) <> VacID Then
     vacidrec = vacidrec & sepchar & CStr(VacRec(Count))
     sepchar = InvertedTblSepChar
    End If
   Next Count
   
   '...Restore record
   obd_GenNametoVacIDs.Item(GenericName) = vacidrec

End Sub

Public Sub Insert_VacID_To_GenName(ByVal VacID As String, ByVal GenericName As String)

 '...Insert a VacID into the Generic Name Inverted file
 
 If obd_GenNametoVacIDs.Exists(GenericName) Then
 
  Dim vacidrec As String
  
  vacidrec = obd_GenNametoVacIDs.Item(GenericName)
  obd_GenNametoVacIDs.Item(GenericName) = vacidrec & IIf(vacidrec = "", "", InvertedTblSepChar) & VacID
  
 Else
  obd_GenNametoVacIDs.Add GenericName, VacID
  
 End If

End Sub


Public Property Get Get_A_Generic_Name(ByVal VacID As String) As String

  '...Pull Generic name out of Record
  
  Dim Record() As String
  
  Record = Split(obd_VacIDToGenName.Item(VacID), FieldSepChar)
  If UBound(Record) <> -1 Then
    Get_A_Generic_Name = Record(VIDToGenName.GenericName)
  End If
  
End Property


Public Sub Remove_GenericName(NameToRemove As String)
    If obd_AllGenericVacNames.Exists(NameToRemove) Then obd_AllGenericVacNames.Remove NameToRemove
End Sub

Public Sub Add_GenericName(GenRec As GenName_Rec)
      '...Add a New Generic Name record to the dictionary
  With GenRec
   If Not obd_AllGenericVacNames.Exists(.GenericName) Then
      obd_AllGenericVacNames.Add .GenericName, .Record
   End If
  End With
End Sub

Public Function Delete_GenericName(ByVal NameToDelete As String) As Boolean

 On Local Error GoTo haveerror

 Dim Record() As String
 Dim sql As String
 
   If obd_AllGenericVacNames.Exists(NameToDelete) Then
   
      Record = Get_A_Generic_Name_Record(NameToDelete)
      sql = "DELETE FROM" & Name_Of_Generic_Vac_Names_Table & "WHERE GenNameID = " & Record(GenVacDictFields.GenericID)
      gobjSQL.ExecuteSQL sql
      obd_AllGenericVacNames.Remove NameToDelete
      Delete_GenericName = True
      
   End If
exitRoutine:
  Exit Function
haveerror:
 ErrMsgBox "Imm_Tables_Delete_GenericName", Error$, Err
 Resume exitRoutine
End Function

Public Sub Remove_Vaccine(ByVal VaccineToRemove As String)
    If obd_Vaccines.Exists(VaccineToRemove) Then obd_Vaccines.Remove VaccineToRemove
End Sub

Public Sub Add_Vaccine(VacRec As Vaccine_Rec)
      '...Add a New Generic Name record to the dictionary
  With VacRec
   If Not obd_Vaccines.Exists(.VaccineID) Then
      obd_Vaccines.Add .VaccineID, .Record & GroupSepChar & .SeriesArray & GroupSepChar & .ExpertArray
   End If
  End With

End Sub

Public Function Delete_Vaccine(ByVal VacID As String) As Boolean

  On Local Error GoTo haveerror
  Dim sql As String
  
  If obd_Vaccines.Exists(VacID) Then
    sql = "DELETE FROM" & Name_Of_Vaccine_Table & "WHERE VaccineID = " & VacID
    
    gobjSQL.ExecuteSQL sql
    
    '...remove from dictionary
    obd_Vaccines.Remove VacID
    
    Delete_Vaccine = True
    
  End If
exitRoutine:
   Exit Function
haveerror:
 ErrMsgBox "Imm_Tables_Delete_Vaccine", Error$, Err
 Resume exitRoutine
End Function

Public Function Delete_Related_Generic_Names(ByVal GenName As String) As String
 If obd_RelatedGenericNames.Exists(GenName) Then obd_RelatedGenericNames.Remove GenName
End Function
Public Function Save_Related_Generic_Names(ByVal GenName As String, RelNames() As String) As String
 Dim Record As String
 Dim RelatedGenName As Variant
 If obd_RelatedGenericNames.Exists(GenName) Then obd_RelatedGenericNames.Remove GenName
 For Each RelatedGenName In RelNames
 
  Record = Record & WithinFieldSepChar & RelatedGenName
 
 Next RelatedGenName
 
 If Record <> "" Then Record = Record & WithinFieldSepChar
 
 obd_RelatedGenericNames.Add GenName, Record
 
End Function
Public Function Get_Related_Generic_Names(ByVal GenName As String) As String
 If obd_RelatedGenericNames.Exists(GenName) Then Get_Related_Generic_Names = obd_RelatedGenericNames.Item(GenName)
End Function


Public Function VaccineToVaccineID(ByVal Vaccine As String) As String
  '...Retrieve VaccineId given Vaccine Description (Name)
  ' SCR 18130 T.Noren  10/29/01 remove 'Force to upper case
'  If obd_VaccineToVacID.Exists(UCase$(Vaccine)) Then VaccineToVaccineID = obd_VaccineToVacID.Item(UCase$(Vaccine))
  If obd_VaccineToVacID.Exists(Vaccine) Then VaccineToVaccineID = obd_VaccineToVacID.Item(Vaccine)
End Function
' SCR #19514;   Developer: Brian Mowbray 05/09/2002 11:10 AM
Public Function VaccineExists(ByVal Vaccine As String) As Boolean
  
  '...Indicate if vaccine name already exists given Vaccine Name. Using the VaccineToID dictionary that is built on startup and will be updated if
  ' a new vaccine is introduced
  
  VaccineExists = obd_VaccineToVacID.Exists(Vaccine)
  
End Function

Public Function IsVaccineInActive(ByVal VacKey As String) As Boolean

  '...Determine if vaccine is a Test I.E  IPPD
  
  Dim Vaccine_Record() As String
  Dim Group1() As String
  
 If obd_Vaccines.Exists(VacKey) Then
   Vaccine_Record = Split(obd_Vaccines.Item(VacKey), GroupSepChar)
   Group1 = Split(Vaccine_Record(0), FieldSepChar)
   IsVaccineInActive = (Group1(VacDictGrp1Fields.InActive) = "1")
 End If
  
End Function

Public Function IsVaccineATest(ByVal VacID As String) As Boolean

  '...Determine if vaccine is a Test I.E  IPPD
  
  Dim GenRec() As String
  Dim gn As String
  
  If obd_VacIDToGenName.Exists(VacID) Then
    GenRec = Split(obd_AllGenericVacNames.Item(Get_A_Generic_Name(VacID)), FieldSepChar)
    IsVaccineATest = IIf(CInt("0" & GenRec(GenVacDictFields.xtype)) = GenericTYpeDrug.test, True, False)
  End If
  
End Function


Public Function IsVaccineASingleShot(ByVal VacID As String) As Boolean

  '...Determine if vaccine is a single Shot  I.E.  Influenza
  
  Dim GenRec() As String
  Dim gn As String
  
  If obd_VacIDToGenName.Exists(VacID) Then
    GenRec = Split(obd_AllGenericVacNames.Item(Get_A_Generic_Name(VacID)), FieldSepChar)
    IsVaccineASingleShot = IIf(CInt("0" & GenRec(GenVacDictFields.xtype)) = GenericTYpeDrug.SingleShot, True, False)
  End If
  
End Function
Public Function GetVaccineType(ByVal VacID As String) As Integer

  '...Determine if vaccine is a series  I.E.  Hep A
  
  Dim GenRec() As String
  Dim gn As String
  
  If obd_VacIDToGenName.Exists(VacID) Then
    GenRec = Split(obd_AllGenericVacNames.Item(Get_A_Generic_Name(VacID)), FieldSepChar)
    GetVaccineType = CInt("0" & GenRec(GenVacDictFields.xtype))
  Else
    '...Invalid Vaccine
    GetVaccineType = -1
  End If
  
End Function
Public Function IsVaccineASeries(ByVal VacID As String) As Boolean

  '...Determine if vaccine is a series
  
  Dim GenRec() As String
  Dim gn As String
  
  If obd_VacIDToGenName.Exists(VacID) Then
    GenRec = Split(obd_AllGenericVacNames.Item(Get_A_Generic_Name(VacID)), FieldSepChar)
    IsVaccineASeries = IIf(CInt("0" & GenRec(GenVacDictFields.xtype)) = GenericTYpeDrug.SeriesOfShots, True, False)
  End If
  
End Function

Public Function VacIDToVaccine(ByVal VacID As String) As String
  ' Get Vaccine Name given VacID
  Dim VacRec() As String
  
  VacRec = Vaccine_BasicData(VacID)
  If UBound(VacRec) <> -1 Then VacIDToVaccine = VacRec(VacDictGrp1Fields.Vaccine)
  
End Function

'<< SCR #25155;   Developer: Brian Mowbray 06/27/2002 02:02 PM
Public Function VacIDToVISVersion(ByVal VacID As String) As String
  ' Get Vaccine Name given VacID
  Dim VacRec() As String
  
  VacRec = Vaccine_BasicData(VacID)
  If UBound(VacRec) <> -1 Then VacIDToVISVersion = VacRec(VacDictGrp1Fields.VIS_Version)
  
End Function

' SCR #19803;   Developer: Brian Mowbray 05/10/2002 12:58 PM
Public Function MfgCodeToMfgName(ByVal sMfg_Code As String) As String
' Get Manufacturer Name given Mfg code
' SCR #52654;   Developer: Brian Mowbray 04/22/2004 12:58 PM
  If obd_MfgCodeToMfgName.Exists(sMfg_Code) Then
    MfgCodeToMfgName = obd_MfgCodeToMfgName.Item(sMfg_Code)
  Else
    MfgCodeToMfgName = sMfg_Code
  End If
End Function
' SCR #19803;   Developer: Brian Mowbray 05/10/2002 12:58 PM
Public Function MfgNameToMfgCode(ByVal sMfg_Name As String) As String
  ' Get Manufacturer Code given Mfg Name
  If obd_MfgNameToMfgCode.Exists(sMfg_Name) Then MfgNameToMfgCode = obd_MfgNameToMfgCode.Item(sMfg_Name)
End Function


' SCR #19514;   Developer: Brian Mowbray 05/09/2002 11:10 AM
Public Function VaccineIDExists(ByVal VacID As String) As Boolean
  
  ' See is vaccine id already exists given VacID
  Dim VacRec() As String
  
  VacRec = Vaccine_BasicData(VacID)
  If UBound(VacRec) <> -1 Then VaccineIDExists = True
  
End Function
Public Function ThisVaccineIsSolvedBy(ByVal VacID As String) As String()

 ' This routine retrieves the vaccines that Solve VacID
 
 Dim VacRec() As String
 Dim Hold As String
 
  VacRec = Vaccine_BasicData(VacID)
  If obd_VaccinesSolvedBy.Exists(VacID) Then
    Hold = obd_VaccinesSolvedBy.Item(VacID)
    If Hold <> vbNullString Then
     ThisVaccineIsSolvedBy = Split(Mid$(Hold, 2, Len(Hold) - 2), WithinFieldSepChar)
    Else
     ThisVaccineIsSolvedBy = SetUboundToMinusOne
    End If
  Else
   ThisVaccineIsSolvedBy = SetUboundToMinusOne
  End If
End Function
Public Function ThisVaccineSolves(ByVal VacID As String) As String()

 ' This routine retrieves the vaccines that are effected by VacID
 
 Dim VacRec() As String
 Dim Hold As String
 
  VacRec = Vaccine_BasicData(VacID)
  If UBound(VacRec) <> -1 Then
    Hold = VacRec(VacDictGrp1Fields.EffectsVacID)
    If Hold <> vbNullString Then
     ThisVaccineSolves = Split(Mid$(Hold, 2, Len(Hold) - 2), WithinFieldSepChar)
    Else
     VacRec = SetUboundToMinusOne
     ThisVaccineSolves = VacRec
    End If
  Else
   ThisVaccineSolves = VacRec
  End If
End Function


Public Function GetCombinationVaccines() As String()

   '...Return all Vaccines that effect other vaccines
   
   Dim Vaccines() As String
   
   If CombinationVaccines = "" Then
     Vaccines = SetUboundToMinusOne
   Else
     Vaccines = Split(right(CombinationVaccines, Len(CombinationVaccines) - 1), WithinFieldSepChar)
   End If
   GetCombinationVaccines = Vaccines
   
End Function

Public Sub Dump_To_Disk()
  On Local Error GoTo haveerror
  Dim Keys() As Variant
  Dim key As Variant
  Dim filen As Integer
  Dim output As String
  Dim test As String
  Dim fo As FileSystemObject
  Dim f As File
  Dim FileSize As Long
  Dim fn As String
  
  Keys = obd_Vaccines.Keys
  Set fo = New FileSystemObject
  For Each key In Keys
xover:
    filen = FreeFile
    output = obd_Vaccines.Item(CStr(key))
    Debug.Print Len(output)
    Open "c:\Data\Source\Immunizations\Data\" & CStr(key) For Random Access Write As #filen Len = LenB(output) + 2
    Put #filen, 1, output
    Close #filen
    Set f = fo.GetFile("c:\Data\Source\Immunizations\Data\" & CStr(key))
    FileSize = (f.Size - 4) * 2 + 2
    FileSize = IIf(FileSize <= 0, 2, FileSize)
    Debug.Print LenB(test)
    Open "c:\Data\Source\Immunizations\Data\" & CStr(key) For Random Access Read As #filen Len = FileSize
    Get #filen, 1, test
    Close #filen
    If test <> obd_Vaccines.Item(CStr(key)) Or test <> output Then Beep
  Next key
  Keys = obd_AllGenericVacNames.Keys
  For Each key In Keys
over2:
    filen = FreeFile
    output = obd_AllGenericVacNames.Item(CStr(key))
    fn = Replace(key, "/", "{")
    fn = Replace(fn, "\", "|")
    Open "c:\Data\Source\Immunizations\Data\" & fn For Random Access Write As #filen Len = LenB(output) + 2
    Put #filen, 1, output
    Close #filen
    Set f = fo.GetFile("c:\Data\Source\Immunizations\Data\" & fn)
    FileSize = (f.Size - 4) * 2 + 2
    FileSize = IIf(FileSize <= 0, 2, FileSize)
    test = Space$(FileSize - 2)
    Open "c:\Data\Source\Immunizations\Data\" & fn For Random Access Read As #filen Len = FileSize
    Get #filen, 1, test
    Close #filen
    If test <> obd_AllGenericVacNames.Item(CStr(key)) Or test <> output Then Beep
    If test <> "" Then Debug.Print fn
  Next key
  Set f = Nothing
  Set fo = Nothing
exitRoutine:
  Exit Sub
haveerror:
  Debug.Print Err.Number, Err.Description
  Close #filen
  Resume xover
  Resume
  Resume over2
End Sub

Public Function GetGroupsForMTF_ID(ByVal MTF_ID As String) As String()
 If UserGroupsMTF_ID.Exists(MTF_ID) Then
   GetGroupsForMTF_ID = Split(UserGroupsMTF_ID.Item(MTF_ID), InvertedTblSepChar)
 Else
  GetGroupsForMTF_ID = SetUboundToMinusOne
 End If
End Function

Public Function GetGroupsForAServiceCode(ByVal ServiceCode As String) As Dictionary

  ' Retrieve all Group Records in the form of a dictionary for a given Service Code

 Dim Keys() As Variant
 Dim key As Variant
 Dim lookfor As String
 Dim Results As Dictionary
 Dim grouprecord() As String
 Dim Group As String
 
 
 ServiceCode = UCase$(ServiceCode)
 lookfor = "DEFAULT" & WithinFieldSepChar
 
 Set Results = New Dictionary
 Keys = PatientGroups.Keys
 
 For Each key In Keys
 
  If Left$(key, 8) = lookfor Then
  
    grouprecord = Split(PatientGroups.Item(key), FieldSepChar)
    If LenB(grouprecord(grouprec.ServiceCode)) = 0 Or UCase$(grouprecord(grouprec.ServiceCode)) = ServiceCode Then
      Group = Mid$(key, 9)
      If Not Results.Exists(Group) Then Results.Add Group, PatientGroups.Item(key)
    End If
    
  End If
 
 Next key

 Set GetGroupsForAServiceCode = Results
 Set Results = Nothing

End Function

Public Function GetGroupsForUnitMTF_ID(ByVal xUnit As String, ByVal MTF_ID As String) As String()
 Dim key As String
 key = xUnit & WithinFieldSepChar & MTF_ID
 If UserGroupsUnitMTF_ID.Exists(key) Then
  GetGroupsForUnitMTF_ID = Split(UserGroupsUnitMTF_ID.Item(key), InvertedTblSepChar)
 Else
  GetGroupsForUnitMTF_ID = SetUboundToMinusOne
 End If
End Function

Public Function Add_Generic_Name_For_A_User_Group(ByVal xGroup As String, ByVal MTF_ID As String, ByVal GenericName As String) As Boolean
 
 ' Delete a Generic Vaccine Names Defined to this xGroup and MTF_ID
 
 Dim key As String
 Dim rec() As String
 Dim newgenrec As String
 Dim foundit As Boolean
 Dim Count As Integer
 Dim MoveRec As Integer
 Dim holdrec() As String
 Dim MinAge As String
 Dim MaxAge As String
 Dim ConsiderAge As String
 
 key = MTF_ID & WithinFieldSepChar & xGroup
 ' SCR 18130 T.Noren  10/29/01 remove 'Force to upper case
' GenericName = UCase$(GenericName)
 If PatientGroups.Exists(key) Then
   holdrec = Split(PatientGroups.Item(key), FieldSepChar)
   If LenB(holdrec(grouprec.GenericNames)) <> 0 Then
    rec = Split(holdrec(grouprec.GenericNames), InvertedTblSepChar)
    For Count = 0 To UBound(rec)
 ' SCR 18130 T.Noren  10/29/01 remove 'Force to upper case
'      If UCase$(rec(count)) = GenericName Then
      If rec(Count) = GenericName Then
       foundit = True
       Exit For
      End If
    Next Count
    If Not foundit Then
      ReDim Preserve rec(UBound(rec) + 1)
    End If
   Else
     ReDim rec(0)
   End If
   If Not foundit Then
      rec(UBound(rec)) = GenericName
      newgenrec = Join(rec, InvertedTblSepChar)
      holdrec(grouprec.GenericNames) = newgenrec
      PatientGroups.Item(key) = Join(holdrec, FieldSepChar)
   End If
 Else
    '...Create a Record for this Demis | Group
    PatientGroups.Add key, _
                      MinAge & FieldSepChar & _
                      MaxAge & FieldSepChar & _
                      ConsiderAge & FieldSepChar & _
                               "" & FieldSepChar & "" & FieldSepChar & GenericName & FieldSepChar & CStr(DefIDControl.UserGroup)
                               'Put in Place holder for Status code, Service code, Generic Names and Type of Group respectively
 End If
 
 '...Indicate a True condition if Generic name NOT found
 Add_Generic_Name_For_A_User_Group = Not foundit
 
End Function

Public Function Delete_Generic_Name_For_A_User_Group(ByVal xGroup As String, ByVal MTF_ID As String, ByVal GenericName As String) As Boolean
 
 ' Delete a Generic Vaccine Names Defined to this xGroup and MTF_ID
 
 Dim key As String
 Dim rec() As String
 Dim newgenrec As String
 Dim foundit As Boolean
 Dim Count As Integer
 Dim MoveRec As Integer
 Dim holdrec() As String
 
 key = MTF_ID & WithinFieldSepChar & xGroup
 ' SCR 18130 T.Noren  10/29/01 remove 'Force to upper case
' GenericName = UCase$(GenericName)
 If PatientGroups.Exists(key) Then
   holdrec = Split(PatientGroups.Item(key), FieldSepChar)
   rec = Split(holdrec(grouprec.GenericNames), InvertedTblSepChar)
   For Count = 0 To UBound(rec)
 ' SCR 18130 T.Noren  10/29/01 remove 'Force to upper case
'     If UCase$(rec(count)) = GenericName Then
     If rec(Count) = GenericName Then
      For MoveRec = Count + 1 To UBound(rec)
        rec(MoveRec - 1) = rec(MoveRec)
      Next MoveRec
      If UBound(rec) <> 0 Then
       ReDim Preserve rec(UBound(rec) - 1)
       newgenrec = Join(rec, InvertedTblSepChar)
      Else
       newgenrec = ""
      End If
      foundit = True
     End If
     holdrec(grouprec.GenericNames) = newgenrec
     PatientGroups.Item(key) = Join(holdrec, FieldSepChar)
     If foundit Then Exit For
   Next Count
 End If
 
 Delete_Generic_Name_For_A_User_Group = foundit
 
End Function

Public Function GetGenericNamesForAGroup(ByVal xGroup As String, ByVal MTF_ID As String) As String()
 
 ' Get Generic Vaccine Names Defined to this xGroup and MTF_ID
 
 Dim key As String
 Dim rec() As String
 key = MTF_ID & WithinFieldSepChar & xGroup
 If PatientGroups.Exists(key) Then
   rec = Split(PatientGroups.Item(key), FieldSepChar)
   GetGenericNamesForAGroup = Split(rec(grouprec.GenericNames), InvertedTblSepChar)
 Else
   GetGenericNamesForAGroup = SetUboundToMinusOne
 End If
End Function

Public Sub Delete_A_User_Defined_Group(ByVal MTF_ID As String, ByVal GroupName As String)

  ' Delete a User Defined Group

 On Local Error GoTo haveerror
 
 Dim RecMtfID() As String
 Dim MinAge As String
 Dim MaxAge As String
 Dim ConsiderAge As String
 Dim Count As Integer
 Dim MoveNames As Integer
 Dim foundit As Boolean
 Dim NewRecMTFID As String
    
    '...Modify Inverted Group Demis
    If UserGroupsMTF_ID.Exists(MTF_ID) Then
    
     RecMtfID = Split(UserGroupsMTF_ID.Item(MTF_ID), InvertedTblSepChar)
     
     For Count = 0 To UBound(RecMtfID)
     
       If RecMtfID(Count) = GroupName Then
         foundit = True
         '...Shift other group names down
         For MoveNames = Count + 1 To UBound(RecMtfID)
           RecMtfID(MoveNames - 1) = RecMtfID(MoveNames)
         Next MoveNames
         
         '...Create new Record
         If UBound(RecMtfID) <> 0 Then
          ReDim Preserve RecMtfID(UBound(RecMtfID) - 1)
          NewRecMTFID = Join(RecMtfID, InvertedTblSepChar)
         Else
          NewRecMTFID = ""
         End If
         
       End If
       
       If foundit Then Exit For
       
     Next Count
     
     If LenB(NewRecMTFID) <> 0 Then
      '...Save New Record
      UserGroupsMTF_ID.Item(MTF_ID) = NewRecMTFID
      
     Else
      '...Get rid of MTF_ID Reference
      UserGroupsMTF_ID.Remove MTF_ID
      
     End If
     
    End If
    
    '...Remove the Record for this Demis | Group
    PatientGroups.Remove MTF_ID & WithinFieldSepChar & GroupName
    
exitRoutine:
  Exit Sub
haveerror:
  lLastError = Err.Number
  sLastErrorDesc = "ImmDATA: Add_A_User_Defined_Group: " & Err.Description
  Err.Raise lLastError, "Imm_Tables: Add_A_User_Defined_Group: ", sLastErrorDesc
  

End Sub
Public Sub Add_A_User_Defined_Group(ByVal MTF_ID As String, ByVal GroupName As String)

  ' Add a New User Defined Group

 On Local Error GoTo haveerror
 
 Dim RecMtfID As String
 Dim MinAge As String
 Dim MaxAge As String
 Dim ConsiderAge As String
    
    '...Build Inverted Group Demis
    If Not UserGroupsMTF_ID.Exists(MTF_ID) Then
     UserGroupsMTF_ID.Add MTF_ID, GroupName
    Else
     RecMtfID = UserGroupsMTF_ID.Item(MTF_ID)
     RecMtfID = RecMtfID & IIf(RecMtfID = "", "", InvertedTblSepChar) & GroupName
     UserGroupsMTF_ID.Item(MTF_ID) = RecMtfID
    End If
    '...Create a Record for this Demis | Group
    PatientGroups.Add MTF_ID & WithinFieldSepChar & GroupName, _
                               MinAge & FieldSepChar & _
                               MaxAge & FieldSepChar & _
                               ConsiderAge & FieldSepChar & _
                               "" & FieldSepChar & "" & FieldSepChar & "" & FieldSepChar & CStr(DefIDControl.UserGroup)
                               'Put in Place holder for Status code, Service code, Generic Names and Type of Group respectively
exitRoutine:
  Exit Sub
haveerror:
  lLastError = Err.Number
  sLastErrorDesc = "ImmDATA: Add_A_User_Defined_Group: " & Err.Description
  Err.Raise lLastError, "Imm_Tables: Add_A_User_Defined_Group: ", sLastErrorDesc
  

End Sub



Public Function Get_Number_of_Vaccines_In_Series(ByVal VacKey As String) As Integer
 '...Get Number of Vaccine schedules in this Series
 Dim Groups() As String
 Dim Series() As String
 If obd_Vaccines.Exists(VacKey) Then
   Groups = Split(obd_Vaccines.Item(VacKey), GroupSepChar)
   Series = Split(Groups(1), SeriesSepChar)
   Get_Number_of_Vaccines_In_Series = UBound(Series) + 1
 Else
    Get_Number_of_Vaccines_In_Series = -1
 End If

End Function

'Public Property Get NewEnum() As IUnknown
'
'    No Can Do For Dictionary
'
'    '...Enable for each on UnitGroupReqForPatient
'    Set NewEnum = UnitGroupReqForPatient.[_NewEnum]
'End Property


Public Function GetThisPatientsGroup(ByVal Group As String) As String()
 
  '...Load a Patient's Group record
  If UnitGroupReqForPatient.Exists(Group) Then
   GetThisPatientsGroup = Split(UnitGroupReqForPatient.Item(Group), FieldSepChar)
  Else
   GetThisPatientsGroup = SetUboundToMinusOne
  End If

End Function

Public Function GetThisPatientsGroupKeys() As Variant()
 
  '...Load all Patient's Group Keys
  If UnitGroupReqForPatient.Count <> 0 Then
   GetThisPatientsGroupKeys = UnitGroupReqForPatient.Keys
  End If


End Function

