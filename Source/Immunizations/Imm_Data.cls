VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Imm_Data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Event CheckForContinue(bValue As Boolean)

Public Event StatusMsg(sMsg As String)
Public Event StatusMsg2(sMsg As String)
Public Event ErrorMsg(sProc As String, sError As String, lError As Long, sDetails As String, bLogError As Boolean)
Public Event ChangeMousePointer(lPointerValue As Long)
Public Event InitMeter(sMsg As String, lMax As Long, lMin As Long, bShowCancel As Boolean, sCaption As String)
Public Event UpdateMeter(lVal As Long)
Public Event RemoveMeter()

Private bContinue As Boolean

Private lLastError As Long
Private sLastErrorDesc As String


Public Enum UpdateType
  Append = 1
  Update = 2
  Delete = 3
  Archive = 4
End Enum




Public Function GenericNameToID(ByVal sGenericName As String) As Long
  Dim lVaccineID As Long
  Dim rst As ADODB.Recordset
  Dim sSQL As String
  On Error Resume Next
  Select Case sGenericName
  Case "DTaP"
    lVaccineID = 20
  Case "Hep A"
    lVaccineID = 52
  Case "Hep B"
    lVaccineID = 45
  Case "Hib"
    lVaccineID = 17
  Case "Influenza"
    lVaccineID = 16
  Case "IPPD"
    lVaccineID = 96
  Case "MMR"
    lVaccineID = 3
  Case "Polio"
    lVaccineID = 10
  Case "Rabies"
    lVaccineID = 18
  Case "Td"
    lVaccineID = 9
  Case "Typhoid"
    lVaccineID = 91
  Case Else
    IsPGM_MDB_Open
    sSQL = "SELECT VaccineID FROM tblVaccines WHERE GenericName='" & sGenericName & "'"
    Set rst = gobjSQL.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
    If rst.BOF And rst.EOF Then
    Else
      lVaccineID = rst![VaccineID]
    End If
  End Select
 
  GenericNameToID = lVaccineID
End Function

Sub DeleteGroup(sspon_SSAN_FMP As String, sGroup As String)
  Dim sSQL As String
  Dim rst As ADODB.Recordset
  On Error GoTo DeleteGroup_Err
  '>> Begin:SCR #38820;   Developer: Brian Mowbray 07/21/2003 10:43 AM
  If Len(sspon_SSAN_FMP) > 0 Then
    If Len(sGroup) > 0 Then
      sSQL = "SELECT SPON_SSAN_FMP, Group_Name FROM Imm_People_Group WHERE (Spon_SSAN_FMP ='" & sspon_SSAN_FMP & "') AND (Group_Name ='" & ValidateDBChar(sGroup) & "')"
      Set rst = gobjSQL.OpenRecordset(sSQL, adOpenForwardOnly, adLockOptimistic)
      If rst.BOF And rst.EOF Then
      Else
        sSQL = "DELETE FROM Imm_People_Group WHERE SPON_SSAN_FMP = '" & sspon_SSAN_FMP & "' AND (Group_Name ='" & ValidateDBChar(sGroup) & "')"
        gobjSQL.ExecuteSQL sSQL
        WriteAuditTrail sSQL
      End If
    End If
  End If
  '<< End: SCR #38820;
DeleteGroup_Exit:
  On Error Resume Next
  rst.Close
  Set rst = Nothing
  Exit Sub

DeleteGroup_Err:
  lLastError = Err.Number
  sLastErrorDesc = "CITA_DATA: DeleteGroup: " & Err.Description
  Err.Raise lLastError, "CITA_DATA: DeleteGroup: ", sLastErrorDesc

End Sub


'>> SCR #33675;   Developer: Brian Mowbray 07/22/2003 10:45 AM
Public Function GetJabs(lVaccineID As Long, sspon_SSAN_FMP As String) As Long
  On Error GoTo Func_Err
  
  If lVaccineID = 75 Or lVaccineID = 105 Then
    'Tracker #52998; Developer: MAG 04/28/2004
    If PatientIMM Is Nothing Then
      Exit Function
    End If
    
    '<< SCR #52421;  Developer: Brian Mowbray 04/14/2004 01:18 PM
    'Changed this to get called once and filter the rst each time
    If PatientIMM.ImmRecordExist(lVaccineID, Null) Then
      GetJabs = 15
    Else
      GetJabs = 3
    End If
  Else
    GetJabs = 0
  End If

Func_Exit:
  On Error Resume Next
  Exit Function

Func_Err:
  lLastError = Err.Number
  sLastErrorDesc = "IMM_DATA: GetJabs: " & Err.Description
  Err.Raise lLastError, "IMM_DATA: GetJabs: ", sLastErrorDesc
End Function

Public Property Get LastError() As Long
  On Error Resume Next
  LastError = lLastError
End Property
Public Property Get LastErrorDesc() As String
  On Error Resume Next
  LastErrorDesc = sLastErrorDesc
End Property
Private Sub SayStatus(sMsg As String)
  On Error Resume Next
  RaiseEvent StatusMsg(sMsg)
End Sub
Public Sub UpdateExemptions(sspon_SSAN_FMP As String, iExempt As Integer, vDate As Variant, Optional sUnit As String = vbNullString, Optional vDOB As Variant = Null, Optional bOD_Anthrax As Boolean = False)
  Dim rst As ADODB.Recordset
  Dim rstP As ADODB.Recordset
  Dim sSQL As String
  Dim iVacc_Series As Integer
  Dim lVaccineID As Long
  Dim bUpdateSeries As Boolean
  On Error GoTo Err_Handler
  Screen.MousePointer = vbHourglass
  SayStatus "Updating exemptions..."
   
  If iExempt = 0 Then
    sSQL = "SELECT * FROM Imm_People_Vaccine where SPON_SSAN_FMP='" & sspon_SSAN_FMP & "' AND Vacc_Exempt > 0"
    Set rst = gobjSQL.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
    If rst.BOF And rst.EOF Then
    Else
      Do Until rst.EOF
        Select Case rst![Vacc_Exempt]
        Case 2, 3, 4, 5, 7, 8, 10, 11, 12
          DeleteImm_People_Vaccine sspon_SSAN_FMP, rst![VaccineID], rst![Vacc_Date]
        End Select
        rst.MoveNext
      Loop
    End If
    rst.Close
    '<< SCR #42805; 42809;   Developer: Brian Mowbray 10/02/2003 03:20 PM
    If Not gbITT Then
      PatientIMM.GetPatientsVaccineSchedule True
    End If
  Else
    '...TJNLOOKAT
    Set rst = GetLatestGenericVaccines(sspon_SSAN_FMP, True, sUnit, vDOB)
    If rst.BOF And rst.EOF Then
    Else
      Do Until rst.EOF
        If IsNull(rst![VaccineID]) Then
          lVaccineID = GenericNameToID(rst![GenericName])
        Else
          lVaccineID = rst![VaccineID]
        End If
        If rst![Vacc_Exempt] > 0 Then
          bUpdateSeries = False
          Select Case rst![Vacc_Exempt]
          Case 2, 3, 4, 8, 10, 12
            DeleteImm_People_Vaccine sspon_SSAN_FMP, lVaccineID, rst![Vacc_Date]
          Case 7, 11
            If bOD_Anthrax Then
              Exit Sub
            Else
              DeleteImm_People_Vaccine sspon_SSAN_FMP, lVaccineID, rst![Vacc_Date]
            End If
          End Select
        Else
          bUpdateSeries = True
        End If
        If bUpdateSeries Then
          iVacc_Series = utNullToZero(rst![Vacc_Series]) + 1
        Else
          iVacc_Series = utNullToZero(rst![Vacc_Series])
        End If
        '<< SCR #42805; 42809;   Developer: Brian Mowbray 10/02/2003 03:21 PM
        If Not gbITT Then PatientIMM.AddExemption CStr(lVaccineID), iVacc_Series, iExempt, vDate
        AddImm_People_Vaccine sspon_SSAN_FMP, lVaccineID, Date, iVacc_Series, , , , , , iExempt, vDate
        rst.MoveNext
      Loop
    End If
  End If
  '...kluge to set everything right
  '<< SCR #42805; 42809;   Developer: Brian Mowbray 10/02/2003 03:21 PM
  If Not gbITT Then
    PatientIMM.GetPatientsVaccineSchedule True
  End If
  Imm_People_Recalc sspon_SSAN_FMP, vDOB, sUnit
Sub_exit:
  On Error Resume Next
  rst.Close
  rstP.Close
  Set rstP = Nothing
  Set rst = Nothing
  Screen.MousePointer = vbDefault
  SayStatus ""
  Exit Sub

Err_Handler:
  lLastError = Err.Number
  sLastErrorDesc = "Imm_DATA: UpdateExemptions: " & Err.Description
  Err.Raise lLastError, "Imm_DATA: UpdateExemptions: ", sLastErrorDesc
End Sub
Public Sub Imm_People_Recalc(sspon_SSAN_FMP As String, _
                             Optional vDOB As Variant = vbNull, _
                             Optional vUnit As Variant = vbNull, Optional CalcNextSeries As Boolean = True)
  
  
  Dim rstRecalc As ADODB.Recordset
  Dim sVaccine As String
  Dim vNextDue As Variant
  Dim vDate As Variant
  Dim sSQL As String
  Dim iSeries As Integer
  Dim RecSeries As Integer
  Dim VacRec() As String
  Dim LastGenName As String
  
  On Error GoTo Imm_People_Recalc_Err
  
  '>> Begin:SCR #38820;   Developer: Brian Mowbray 07/21/2003 10:44 AM
  If Len(sspon_SSAN_FMP) = 0 Then
    Exit Sub
  End If
  '<< End: SCR #38820;
  
  
  
  sSQL = "SELECT Imm_People_Vaccine.SPON_SSAN_FMP, tblVaccines.GenericName, " & _
      "Imm_People_Vaccine.VaccineID, Imm_People_Vaccine.Vacc_Date, " & _
      "Imm_People_Vaccine.Vacc_Series, Imm_People_Vaccine.Vacc_NextDue, " & _
      "Imm_People_Vaccine.Vacc_Exempt, Imm_People_Vaccine.ExemptionExpirationDate, " & _
      "Imm_People_Vaccine.Result, Imm_People_Vaccine.Deers_Action, Imm_People_Vaccine.LastEdited " & _
      "FROM Imm_People_Vaccine, tblVaccines " & _
      "WHERE Imm_People_Vaccine.VaccineID = tblVaccines.VaccineID " & _
      "AND (((Imm_People_Vaccine.SPON_SSAN_FMP) ='" & sspon_SSAN_FMP & "')) " & _
      "ORDER BY Imm_People_Vaccine.SPON_SSAN_FMP, tblVaccines.GenericName, " & _
      "Imm_People_Vaccine.Vacc_Date DESC"
    
  Set rstRecalc = gobjSQL.OpenRecordset(sSQL, adOpenForwardOnly, adLockOptimistic)
  
  Do Until rstRecalc.EOF
  
    sVaccine = rstRecalc![GenericName]
    
    ' T.Noren 10/8/01  avoid Mass enoculation problem
    '   I.E  PatientImm not instantiated
    If CalcNextSeries Then
        '...Get last Series given for this Generic Name
        '<< SCR #42805; 42809;   Developer: Brian Mowbray 10/02/2003 03:21 PM
        If Not gbITT Then
         If LastGenName <> sVaccine Then
          iSeries = PatientIMM.GetLastGenSeries(sVaccine)
          LastGenName = sVaccine
         End If
        End If
    End If
    'Check for null series
    If utNullToZero(rstRecalc![Vacc_Series]) = 0 Then
      RecSeries = 1
    Else
      RecSeries = rstRecalc![Vacc_Series]
    End If
   
    If rstRecalc![Vacc_Date] < #1/1/1930# Then
      sSQL = "UPDATE Imm_People_Vaccine SET Imm_People_Vaccine.DEERS_Action = '" & "X" & "'" & " WHERE (SPON_SSAN_FMP='" & sspon_SSAN_FMP & "') AND (VaccineID=" & CLng(rstRecalc![VaccineID]) & ") AND (Vacc_Date= " & SaySQLDateString(rstRecalc![Vacc_Date]) & ")"
      gobjSQL.ExecuteSQL sSQL
      
    End If
    '<< SCR #42805; 42809;   Developer: Brian Mowbray 10/02/2003 03:21 PM
    If Not gbITT Then
     If Not IsEmpty(iSeries) And iSeries = CStr(rstRecalc!Vacc_Series) Then
       vNextDue = PatientIMM.VaxNextDue(CStr(rstRecalc!VaccineID), RecSeries, rstRecalc![Vacc_Date], NextDateTypeCalc.StatusAsOf)
       vNextDue = IIf(Not IsDate(vNextDue), "", Format(vNextDue, "dd mmm yyyy"))
     Else
      vNextDue = ""
     End If
    Else
     vNextDue = Format(VaxNextDue(sspon_SSAN_FMP, rstRecalc!VaccineID, RecSeries, rstRecalc![Vacc_Date], rstRecalc![Vacc_Exempt], vDOB, rstRecalc![result], vUnit & vbNullString), "dd mmm yyyy")
    End If
    vDate = Format(rstRecalc![Vacc_Date], "mm/dd/yyyy")
    If utNullToZero(rstRecalc![Vacc_NextDue]) <> utNullToZero(vNextDue) Then
      sSQL = "UPDATE Imm_People_Vaccine SET Vacc_NextDue =" & SaySQLDateString(vNextDue) & " , LastEdited=" & SaySQLDateString(Now()) & " WHERE (SPON_SSAN_FMP='" & sspon_SSAN_FMP & "') AND (VaccineID=" & CLng(rstRecalc![VaccineID]) & ") AND (Vacc_Date= " & SaySQLDateString(vDate) & ")"
      gobjSQL.ExecuteSQL sSQL
    End If
    rstRecalc.MoveNext
    If rstRecalc.EOF Then Exit Do
    
    '...Vnextdue should probably be set to vbnullstring in theater too
    '<< SCR #42805; 42809;   Developer: Brian Mowbray 10/02/2003 03:21 PM
    If Not gbITT Then vNextDue = vbNullString
    
    Do While sVaccine = rstRecalc![GenericName]
      If utNullToZero(rstRecalc![Vacc_NextDue]) = 0 Then
      Else
        sSQL = "UPDATE Imm_People_Vaccine SET Vacc_NextDue =" & SaySQLDateString(vNextDue) & " , LastEdited=" & SaySQLDateString(Now()) & " WHERE (SPON_SSAN_FMP='" & sspon_SSAN_FMP & "') AND (VaccineID=" & CLng(rstRecalc![VaccineID]) & ") AND (Vacc_Date= " & SaySQLDateString(rstRecalc![Vacc_Date]) & ")"
        gobjSQL.ExecuteSQL sSQL
      End If
      rstRecalc.MoveNext
      If rstRecalc.EOF Then Exit Do
    Loop
  Loop
  
Imm_People_Recalc_Exit:
  On Error Resume Next
  rstRecalc.Close
  Set rstRecalc = Nothing
  Exit Sub

Imm_People_Recalc_Err:
  lLastError = Err.Number
  sLastErrorDesc = "CITA_DATA: Imm_People_Recalc: " & Err.Description
  Err.Raise lLastError, "CITA_DATA: Imm_People_Recalc: ", sLastErrorDesc
  Resume Imm_People_Recalc_Exit
  Resume
End Sub
Public Function VaxNextDue(sspon_SSAN_FMP As String, ByVal lVaccineID As Long, _
                           vSeries As Variant, vDate, _
                           Optional vExempt As Variant = Null, _
                           Optional vDOB As Variant = Null, _
                           Optional vTST_Result As Variant = Null, _
                           Optional vUnit As Variant = Null) As Variant
  
    
  Dim dblDays As Double
  Dim iAge As Integer
  On Error GoTo VaxNextDue_Err
  If utNullToZero(vExempt) > 0 Then
    VaxNextDue = Null
    Exit Function
  End If
  dblDays = 0
  Select Case lVaccineID
  Case 3 'MMR for Children
    If vDOB = Null Then
    Else
      If iAge < 7 Then
        Select Case vSeries
        Case 1
          dblDays = 1710
        End Select
      Else
      End If
    End If
  Case 8 ' Hep B - Child
    Select Case vSeries
    Case 1
      dblDays = 30
    Case 2
      dblDays = 150
    End Select
  Case 9 ' Td
    dblDays = 3650
  Case 10 'IPV
    If vDOB = Null Then
    Else
      If iAge < 7 Then
        Select Case vSeries
        Case 1
          dblDays = 60
        Case 2
          dblDays = 540 '? Should is be due at 6 mo or 18 mo
        Case 3
          dblDays = 1620 'Should have number 5 between 4-6 years of age
        End Select
      End If
    End If
  Case 16 'Influenza
    If Year(vDate) = Year(Date) Then
      If Month(vDate) < 9 Then
        VaxNextDue = DateSerial(Year(Date), 10, 1)
      Else
        VaxNextDue = DateSerial(Year(Date) + 1, 10, 1)
      End If
    Else
      If (Year(vDate) = (Year(Date) - 1)) Then
        If Month(vDate) < 9 Then
          VaxNextDue = DateSerial(Year(Date) - 1, 10, 1)
        Else
          VaxNextDue = DateSerial(Year(Date), 10, 1)
        End If
      Else
        VaxNextDue = DateSerial(Year(Date) - 1, 10, 1)
      End If
    End If
    Exit Function
  Case 18, 40 'Rabies - Intramuscular or Intradermal
    Select Case vSeries
    Case 1
      dblDays = 7
    Case 2
      dblDays = 14
    End Select
  Case 20 ' DTaP
    Select Case vSeries
    Case 1
      dblDays = 60
    Case 2
      dblDays = 60
    Case 3
      dblDays = 270
    Case 4
      dblDays = 1620 'Should have number 5 between 4-6 years of age
    End Select
  Case 21 ' Varicella
    If vDOB = Null Then
    Else
      If GetAge(vDOB, Date) > 12 Then
        If vSeries = 1 Then dblDays = 30
      End If
    End If
  Case 23 ' Plague
    Select Case vSeries
    Case 1
      dblDays = 30
    Case 2
      dblDays = 60
    Case Is > 2
      dblDays = 365
    End Select
  Case 24 ' Anthrax
    Select Case vSeries
    Case 0 ' Should not happen, but...
      dblDays = 14
    Case 1
      dblDays = 14
    Case 2
      dblDays = 14
    Case 3
      dblDays = 150
    Case 4
      dblDays = 180
    Case 5
      dblDays = 180
    Case Else
      dblDays = 365
    End Select
  Case 25 ' Typhoid - Oral
    dblDays = 1825
  Case 31 ' Hep A (Pediatric)
    If vSeries = 1 Then dblDays = 180
  Case 32 ' Meningococcal
      dblDays = 1825
  Case 33 ' Pneumococcal
    dblDays = 2190
  Case 35 ' Tetanus toxoid
    dblDays = 3650
  Case 37 ' Yellow Fever
    dblDays = 3650
  Case 39 ' Japanese Encephalitis
    Select Case vSeries
    Case 1
      dblDays = 7
    Case 2
      dblDays = 23
    Case Is > 2
      dblDays = 1095
    End Select
  Case 41, 91, 101 ' Typhoid - Typhim Vi
    dblDays = 730
  Case 43 ' Hep B - Adult
    Select Case vSeries
    Case 1
      dblDays = 30
    Case 2
      dblDays = 150
    End Select
  Case 17, 46, 47, 48 ' Hib - HbOC or Hib - PRP-T
    Select Case vSeries
    Case 1
      dblDays = 60
    Case 2
      dblDays = 60
    Case 3
      dblDays = 180
    End Select
  Case 48 ' Hib - PRP-OMP
    Select Case vSeries
    Case 1
      dblDays = 60
    Case 2
      dblDays = 240
    End Select
  Case 51 'Hib-Hep B
    Select Case vSeries
    Case 1
      dblDays = 60
    Case 2
      dblDays = 280
    End Select
  Case 52, 85, 83 ' Hep A (Adult)
    If vSeries = 1 Then
      dblDays = 180
    End If
  Case 53 ' Typhoid - Parenteral-Military Unique
    Select Case vSeries
    Case 1
      dblDays = 30
    Case Is > 1
      dblDays = 1095
    End Select
  Case 95, 96, 97, 98, 1001 'TB Skin Test
    VaxNextDue = IPPDNextDue(sspon_SSAN_FMP, vDate, vTST_Result, Null, vExempt, vUnit)
    Exit Function
  Case 100  'Pneumococcal Conjugate
    Select Case vSeries
    Case 1
      dblDays = 60
    Case 2
      dblDays = 60
    Case 3
      dblDays = 180
    End Select
  Case 75, 105
    VaxNextDue = SPNextDue(sspon_SSAN_FMP, vDate, vTST_Result, Null, vExempt, vUnit)
    Exit Function
  End Select
  If dblDays = 0 Then
    VaxNextDue = Null
  Else
    VaxNextDue = DateAdd("d", dblDays, vDate)
  End If
VaxNextDue_Exit:
  On Error Resume Next
  Exit Function

VaxNextDue_Err:
  lLastError = Err.Number
  sLastErrorDesc = "CITA_DATA: VaxNextDue: " & Err.Description
  Err.Raise lLastError, "CITA_DATA: VaxNextDue: ", sLastErrorDesc
End Function
Private Function IPPDNextDue(sspon_SSAN_FMP As String, vDate As Variant, vResult As Variant, vDays As Variant, vExempt As Variant, Optional vUnit As Variant = Null) As Variant
  Dim sSQL As String
  Dim rstR As ADODB.Recordset
  Dim rstHE As ADODB.Recordset
  Dim sUnit As String
  On Error GoTo IPPDNextDue_Err
  IPPDNextDue = Null
  If utNullToZero(vExempt) > 0 Then
    IPPDNextDue = Null
    Exit Function
  End If
   
  sSQL = "SELECT Imm_People_Group.Spon_SSAN_FMP "
  sSQL = sSQL & "FROM Imm_People_Group, tblVaccine_Groups_Vaccine WHERE Imm_People_Group.Group_Name = tblVaccine_Groups_Vaccine.Group_Name "
  sSQL = sSQL & "AND ((Imm_People_Group.Spon_SSAN_FMP='" & sspon_SSAN_FMP & "') AND (tblVaccine_Groups_Vaccine.GenericName='IPPD')) "
  sSQL = sSQL & "UNION SELECT Imm_People_Group.Spon_SSAN_FMP "
  sSQL = sSQL & "FROM Imm_People_Group , User_Vaccine_Groups_Vaccine WHERE Imm_People_Group.Group_Name = User_Vaccine_Groups_Vaccine.Group_Name "
  sSQL = sSQL & "AND ((Imm_People_Group.Spon_SSAN_FMP='" & sspon_SSAN_FMP & "') AND (User_Vaccine_Groups_Vaccine.GenericName='IPPD'))"
  sUnit = ValidateDBChar(vUnit & vbNullString)
  If Len(sUnit) = 0 Then
  Else
    sSQL = sSQL & " UNION SELECT '999999999' AS SPON_SSAN_FMP "
    sSQL = sSQL & "FROM Imm_Unit_Group , tblVaccine_Groups_Vaccine WHERE Imm_Unit_Group.Group_Name = tblVaccine_Groups_Vaccine.Group_Name "
    sSQL = sSQL & "AND ((Imm_Unit_Group.UNIT='" & ValidateDBChar(sUnit) & "') AND (tblVaccine_Groups_Vaccine.GenericName='IPPD'))"
  End If
  Set rstR = gobjSQL.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  If rstR.BOF And rstR.EOF Then
  Else
    If Len(vResult & "") = 0 Then
      IPPDNextDue = vDate + 2
    Else
      Select Case vResult
      Case "P", "Positive", "E", "Exempt", "X"
      Case "None"
        IPPDNextDue = vDate + 2
      Case Else
        If utNullToZero(vDate) = 0 Then
        Else
          ' This sets the requirement to every two years if they are not part of the Hospital Employee Group
          sSQL = "SELECT Imm_People_Group.Spon_SSAN_FMP, Imm_People_Group.Group_Name From Imm_People_Group " & _
            "WHERE (((Imm_People_Group.Spon_SSAN_FMP)='" & sspon_SSAN_FMP & "') AND ((Imm_People_Group.Group_Name)='Hospital Employee' Or (Imm_People_Group.Group_Name)='AMC Mobility'))"
          
          Set rstHE = gobjSQL.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
          If rstHE.BOF And rstHE.EOF Then
            If IsNull(vDays) Then vDays = 730
          Else
            If IsNull(vDays) Then vDays = 365
          End If
          IPPDNextDue = DateAdd("d", vDays, vDate)
        End If
      End Select
    End If
  End If
IPPDNextDue_Exit:
  On Error Resume Next
  rstR.Close
  Set rstR = Nothing
  rstHE.Close
  Set rstHE = Nothing
  Exit Function
IPPDNextDue_Err:
  lLastError = Err.Number
  sLastErrorDesc = "CITA_DATA: IPPDNextDue: " & Err.Description
  Err.Raise lLastError, "CITA_DATA: IPPDNextDue: ", sLastErrorDesc
End Function
Private Function SPNextDue(sspon_SSAN_FMP As String, vDate As Variant, vResult As Variant, vDays As Variant, vExempt As Variant, Optional vUnit As Variant = Null) As Variant
  Dim sSQL As String
  Dim rstR As ADODB.Recordset
  Dim rstHE As ADODB.Recordset
  Dim sUnit As String
  On Error GoTo IPPDNextDue_Err
  SPNextDue = Null
  If utNullToZero(vExempt) > 0 Then
    SPNextDue = Null
    Exit Function
  End If
   
  
    If Len(vResult & "") = 0 Then
      SPNextDue = vDate + 7
    Else
    End If

IPPDNextDue_Exit:
  On Error Resume Next
  rstR.Close
  Set rstR = Nothing
  rstHE.Close
  Set rstHE = Nothing
  Exit Function
IPPDNextDue_Err:
  lLastError = Err.Number
  sLastErrorDesc = "CITA_DATA: IPPDNextDue: " & Err.Description
  Err.Raise lLastError, "CITA_DATA: IPPDNextDue: ", sLastErrorDesc
End Function

Public Sub AddImm_People_Vaccine(sspon_SSAN_FMP As String, lVaccineID As Long, _
                                 vDate As Variant, Optional ByVal iSeries As Integer = 1, _
                                 Optional vLot As Variant = Null, Optional vMfg As Variant = Null, _
                                 Optional vVacc_Dosage As Variant = Null, _
                                 Optional vSite As Variant = Null, _
                                 Optional vProvider As Variant = Null, _
                                 Optional vExempt As Variant = Null, _
                                 Optional vExpire As Variant = Null, _
                                 Optional vResult As Variant = Null, _
                                 Optional vRxn_Size As Variant = vbNullString, _
                                 Optional vRoute As Variant = vbNullString, _
                                 Optional vVISVersion As Variant = vbNullString)
  Dim rst As ADODB.Recordset
  Dim sSQL As String
  Dim ImmRelNcid As String
  Dim VacRec() As String
  On Error GoTo AIPV_Err
  
  If utNullToZero(vDate) = 0 Then
    ErrMsgBox "AddImm_People_Vaccine", "Missing the date of immunization.", 0, sspon_SSAN_FMP & "|" & CStr(lVaccineID)
    Exit Sub
  End If
  If vDate > (Date + 1) Then
    ErrMsgBox "AddImm_People_Vaccine", "Invalid date of immunization: " & Format$(vDate, "d mmm yyyy") & ".", 0, sspon_SSAN_FMP & "|" & CStr(lVaccineID) & "|" & Format$(vDate, "yyyymmdd")
    Exit Sub
  End If
  '<< SCR #42805; 42809;   Developer: Brian Mowbray 10/02/2003 03:21 PM
  If Not gbITT Then
    VacRec = IMM_Tables.Vaccine_BasicData(CStr(lVaccineID))
    ImmRelNcid = VacRec(VacDictGrp1Fields.ImmRel_Ncid)
  End If
  sSQL = "SELECT * FROM Imm_People_Vaccine WHERE ((SPON_SSAN_FMP='" & sspon_SSAN_FMP & "') AND (VaccineID=" & lVaccineID & ") AND (Vacc_Date=" & SaySQLDateString(vDate) & "))"
  Set rst = gobjSQL.OpenRecordset(sSQL, adOpenForwardOnly, adLockOptimistic)
  If rst.BOF And rst.EOF Then
    With rst
      '<< SCR #42805; 42809;   Developer: Brian Mowbray 10/02/2003 03:21 PM
      If Not gbITT Then
        '<< Begin:SCR #25155;   Developer: Brian Mowbray 06/27/2002 01:57 PM
        sSQL = "INSERT INTO Imm_People_Vaccine (SPON_SSAN_FMP, IMMRELNCID, VaccineID, Vacc_Series, Vacc_Date, Lot_Nbr, Mfg_Code, Vacc_Dosage, ProviderID, Vacc_Exempt, ExemptionExpirationDate, Result, RXN_Size, Site, Route, LastEdited, LastEditedBy, MTF_Code, VIS_Version) " & _
          "VALUES ('" & sspon_SSAN_FMP & "'" & " ," & ImmRelNcid & " ," & lVaccineID & " ," & iSeries & " ," & SaySQLDateString(vDate) & " ,'" & Left(vLot & vbNullString, 16) & "'" & " ,'" & Left(vMfg & vbNullString, 4) & "'" & " ,'" & Left(vVacc_Dosage & vbNullString, 10) & "'" & " ,'" & vProvider & "'" & " ," & utNullToZero(vExempt) & " ," & SaySQLDateString(vExpire) & " ,'" & Left(vResult & vbNullString, 1) & "'" & " ,'" & Left(vRxn_Size, 3) & "'" & " ,'" & Left(vSite & vbNullString, 50) & "'" & " ,'" & ValidateDBChar(Left(vRoute & vbNullString, 50)) & "'" & " ," & SaySQLDateString(Now) & " ,'" & CurrentUser() & "', '" & SayMTF_ID() & "', '" & ValidateDBChar(Left(vVISVersion, 255)) & "')"
        '>> End: SCR #25155;
      Else
        '<< Begin:SCR #25360;   Developer: Brian Mowbray 07/03/2002 02:38 PM
        sSQL = "INSERT INTO Imm_People_Vaccine (SPON_SSAN_FMP,  VaccineID, Vacc_Series, Vacc_Date, Lot_Nbr, Mfg_Code, Vacc_Dosage, ProviderID, Vacc_Exempt, ExemptionExpirationDate, Result, RXN_Size, Site, Route, LastEdited, LastEditedBy, MTF_Code, VIS_Version) " & _
          "VALUES ('" & sspon_SSAN_FMP & "'" & " ," & lVaccineID & " ," & iSeries & " ," & SaySQLDateString(vDate) & " ,'" & Left(vLot & vbNullString, 16) & "'" & " ,'" & Left(vMfg & vbNullString, 4) & "'" & " ,'" & Left(vVacc_Dosage & vbNullString, 10) & "'" & " ,'" & vProvider & "'" & " ," & utNullToZero(vExempt) & " ," & SaySQLDateString(vExpire) & " ,'" & Left(vResult & vbNullString, 1) & "'" & " ,'" & Left(vRxn_Size, 3) & "'" & " ,'" & Left(vSite & vbNullString, 50) & "'" & " ,'" & ValidateDBChar(Left(vRoute & vbNullString, 50)) & "'" & " ," & SaySQLDateString(Now) & " ,'" & CurrentUser() & "', '" & SayMTF_ID() & "', '" & ValidateDBChar(Left(vVISVersion, 255)) & "')"
      '>> End: SCR #25360;
      End If
      gobjSQL.ExecuteSQL sSQL
      WriteAuditTrail sSQL
    End With
  End If

AIPV_Exit:
  On Error Resume Next
  rst.Close
  Set rst = Nothing
  Exit Sub

AIPV_Err:
  lLastError = Err.Number
  sLastErrorDesc = "CITA_DATA: AddImm_People_Vaccine: " & Err.Description
  Err.Raise lLastError, "CITA_DATA: AddImm_People_Vaccine: ", sLastErrorDesc

End Sub


Public Function Recalc_All_NextDue(sGenericName As String) As Boolean
  Dim rst As ADODB.Recordset
  Dim lCnt As Long
  Dim sKey As String
  Dim vNextDue As Variant
  Dim lTotal As Long
  Dim sSQL As String
  On Error GoTo Recalc_All_NextDue_Err
  Screen.MousePointer = vbHourglass
  SayStatus "Opening recordset..."
   
  If sGenericName = "All" Then
    sSQL = "SELECT Imm_People_Vaccine.SPON_SSAN_FMP, Imm_People.PT_DOB, Imm_People.PT_UNIT, tblVaccines.GenericName, Imm_People_Vaccine.Vacc_Date, Imm_People_Vaccine.VaccineID, Imm_People_Vaccine.Vacc_Series, Imm_People_Vaccine.Vacc_NextDue, Imm_People_Vaccine.Vacc_Exempt, Imm_People_Vaccine.Result, Imm_People_Vaccine.LastEdited "
    sSQL = sSQL & "FROM Imm_People INNER JOIN (Imm_People_Vaccine INNER JOIN tblVaccines ON Imm_People_Vaccine.VaccineID = tblVaccines.VaccineID) ON Imm_People.SPON_SSAN_FMP = Imm_People_Vaccine.SPON_SSAN_FMP "
    sSQL = sSQL & "ORDER BY Imm_People_Vaccine.SPON_SSAN_FMP, tblVaccines.GenericName, Imm_People_Vaccine.Vacc_Date DESC"
  Else
    sSQL = "SELECT Imm_People_Vaccine.SPON_SSAN_FMP, Imm_People.PT_DOB, Imm_People.PT_UNIT, tblVaccines.GenericName, Imm_People_Vaccine.Vacc_Date, Imm_People_Vaccine.VaccineID, Imm_People_Vaccine.Vacc_Series, Imm_People_Vaccine.Vacc_NextDue, Imm_People_Vaccine.Vacc_Exempt, Imm_People_Vaccine.Result, Imm_People_Vaccine.LastEdited "
    sSQL = sSQL & "FROM Imm_People INNER JOIN (Imm_People_Vaccine INNER JOIN tblVaccines ON Imm_People_Vaccine.VaccineID = tblVaccines.VaccineID) ON Imm_People.SPON_SSAN_FMP = Imm_People_Vaccine.SPON_SSAN_FMP "
    sSQL = sSQL & "WHERE (((tblVaccines.GenericName)='" & sGenericName & "')) "
    sSQL = sSQL & "ORDER BY Imm_People_Vaccine.SPON_SSAN_FMP, tblVaccines.GenericName, Imm_People_Vaccine.Vacc_Date DESC"
  End If
  Set rst = gobjSQL.OpenRecordset(sSQL, adOpenForwardOnly, adLockOptimistic)
  If rst.BOF And rst.EOF Then
    GoTo Recalc_No_Error
  End If
  bContinue = True
  Do Until rst.EOF
    DoEvents
    If bContinue Then
    Else
      GoTo Recalc_All_NextDue_Exit
    End If
    lCnt = lCnt + 1
    If lCnt Mod 50 = 0 Then
      SayStatus Format$(lCnt, "###,###,##0")
    End If
    If rst![SPON_SSAN_FMP] & rst![GenericName] <> sKey Then
      sKey = rst![SPON_SSAN_FMP] & rst![GenericName]
      '<< SCR #42805; 42809;   Developer: Brian Mowbray 10/02/2003 03:21 PM
      If Not gbITT Then
          vNextDue = PatientIMM.VaxNextDue(CStr(rst!VaccineID), rst![Vacc_Series], rst![Vacc_Date], NextDateTypeCalc.StatusAsOf)
      Else
          vNextDue = VaxNextDue(rst![SPON_SSAN_FMP], rst![VaccineID], rst![Vacc_Series], rst![Vacc_Date], rst![Vacc_Exempt], rst![PT_DOB], rst![result], rst![Pt_Unit])
      End If
      If utNullToZero(rst![Vacc_NextDue]) <> utNullToZero(vNextDue) Then
        DeleteImm_People_Vaccine rst![SPON_SSAN_FMP], rst![VaccineID], rst![Vacc_Date], False
        rst![Vacc_NextDue] = vNextDue
        rst![LastEdited] = Now
        rst.Update
      End If
    Else
      If utNullToZero(rst![Vacc_NextDue]) = 0 Then
      Else
        rst![Vacc_NextDue] = Null
        rst.Update
      End If
    End If
    rst.MoveNext
  Loop
Recalc_No_Error:
  Recalc_All_NextDue = True
Recalc_All_NextDue_Exit:
  On Error Resume Next
  rst.Close
  Set rst = Nothing
  Screen.MousePointer = vbDefault
  SayStatus ""
  Exit Function

Recalc_All_NextDue_Err:
  lLastError = Err.Number
  sLastErrorDesc = "CITA_DATA: Recalc_All_NextDue: " & Err.Description
  Err.Raise lLastError, "CITA_DATA: Recalc_All_NextDue: ", sLastErrorDesc

End Function
Public Function CalcUnit_Vaccine_Stats(sType As String, Optional bAllowCancel As Boolean = True) As Boolean
  Dim rstUR As ADODB.Recordset
  Dim rstP As ADODB.Recordset
  Dim rstStat As ADODB.Recordset
  Dim rstV As ADODB.Recordset
  Dim rstR As ADODB.Recordset
  Dim lCnt As Long
  Dim iAdd As Integer
  Dim bAddNew As Boolean
  Dim lTotal As Long
  Dim sVaccine As String
  Dim vVacc_Exempt As Variant
  Dim vVacc_Date As Variant
  Dim vVacc_NextDue As Variant
  Dim vVacc_Series As Variant
  Dim vIPPD_REsult As Variant
  Dim sVaxNeeded As String
  Dim iPerStat As Integer
  Dim sHepSeries As String
  Dim sText As String
  Dim sSQL As String
  Dim sPT_Unit As String
  Dim sspon_SSAN_FMP  As String
  Dim sSSAN As String
  Dim sMsg As String
  Dim sFilter As String
  On Error GoTo CalcUnit_Vaccine_Stats_Err
  Screen.MousePointer = vbHourglass
  SayStatus "Clearing the totals..."
  CreateRst Unit_Roster, rstUR
  CreateRst Unit_Vaccine_Count, rstStat
  gobjSQL.ExecuteSQL "DELETE FROM Unit_Roster"
  If sType = "ALL" Then
    gobjSQL.ExecuteSQL "DELETE  FROM tblUnit_Selected"
    gobjSQL.ExecuteSQL "DELETE  FROM Unit_Vaccine_Count"
    sSQL = "INSERT INTO tblUnit_Selected ( UNIT ) SELECT  Imm_Unit.UNIT FROM Imm_Unit"
    gobjSQL.ExecuteSQL sSQL
  Else
    gobjSQL.ExecuteSQL "DELETE FROM Unit_Vaccine_Count FROM Unit_Vaccine_Count, tblUnit_Selected WHERE Unit_Vaccine_Count.UNIT = tblUnit_Selected.UNIT"
  End If
  
  gobjSQL.ExecuteSQL "UPDATE Imm_Unit SET Imm_Unit.NbrAssigned = 0, Imm_Unit.NbrCurrent = 0 From imm_unit, tblUnit_Selected Where imm_unit.Unit = tblUnit_Selected.Unit"
  
  sSQL = "SELECT Imm_Personnel.SSAN + '20' AS SPON_SSAN_FMP, Imm_Personnel.SSAN, Imm_Personnel.UNIT, "
  sSQL = sSQL & "Imm_Personnel.OFFICE_SYMBOL, Imm_Personnel.NAME, Imm_Personnel.RANK, Imm_Personnel.NeededVax, "
  sSQL = sSQL & "Imm_Personnel.VaxCurrent "
  sSQL = sSQL & "FROM tblUnit_Selected, Imm_Personnel WHERE tblUnit_Selected.UNIT = Imm_Personnel.UNIT"

  Set rstP = gobjSQL.OpenRecordset(sSQL, adOpenStatic, adLockOptimistic)
  bContinue = True
  If bAllowCancel Then
    sMsg = "Evaluating immunization status.  Press cancel to halt this procedure."
  Else
    sMsg = "Evaluating immunization status.  This procedure should be run at least twice a month.  Please be patient until this procedure is finished."
  End If
  InitMeter sMsg, , , bAllowCancel, "Calculate Unit Status"
  Do Until rstP.EOF
    
    DoEvents
    If bContinue Then
    Else
      GoTo CalcUnit_Vaccine_Stats_exit
    End If
    lCnt = lCnt + 1
    sspon_SSAN_FMP = rstP![SPON_SSAN_FMP]
    sSSAN = rstP![SSAN]
    sPT_Unit = rstP![Unit] & ""
    If Len(sPT_Unit) = 0 Then
      sPT_Unit = "N/A"
    End If
    SayStatus "Processing person #" & Format$(lCnt, "###,##0")
    '...TJNLOOKAT
    Set rstV = oData.GetLatestGenericVaccines(sspon_SSAN_FMP, True, sPT_Unit)
    If rstUR.BOF And rstUR.EOF Then
    Else
      rstUR.MoveFirst
    End If
    rstUR.Find "SSAN='" & sSSAN & "'"
    If rstUR.EOF Then
      rstUR.AddNew
      rstUR![SSAN] = sSSAN
    End If
    sVaxNeeded = ""
    iPerStat = -1
    Do Until rstV.EOF
      sVaccine = rstV![GenericName]
      sFilter = "UNIT='" & sPT_Unit & "' and Vaccine='" & sVaccine & "'"
      rstStat.Filter = sFilter
      With rstStat
        If .BOF And .EOF Then
          .AddNew
          ![Unit] = sPT_Unit
          ![Vaccine] = sVaccine
        End If
        ![NbrReq] = utNullToZero(![NbrReq]) + 1
        If rstV![VAcc_Current] Then
          ![NbrCurrent] = utNullToZero(![NbrCurrent]) + 1
        Else
          ![NbrDue] = utNullToZero(![NbrDue]) + 1
        End If
        ![LastEval] = Now
        .Update
      End With
      rstStat.Filter = ""
      iPerStat = Max(iPerStat, rstV![VAcc_Current])
      If Not rstV![VAcc_Current] Then
        Select Case sVaccine
        Case "Anthrax"
          If Len(sVaxNeeded) = 0 Then
            sVaxNeeded = "Anthrax (" & Format$(utNullToZero(rstV![Vacc_Series]) + 1, "##") & ")"
          Else
            sVaxNeeded = sVaxNeeded & ", " & "Anthrax (" & Format$(utNullToZero(rstV![Vacc_Series]) + 1, "##") & ")"
          End If
        Case "Hep A"
          If rstV![Vacc_Series] = 1 Then
            sHepSeries = "2"
          Else
            sHepSeries = "1"
          End If
          If Len(sVaxNeeded) = 0 Then
            sVaxNeeded = "Hep A (" & sHepSeries & ")"
          Else
            sVaxNeeded = sVaxNeeded & ", " & "Hep A (" & sHepSeries & ")"
          End If
        Case Else
          If Len(sVaxNeeded) = 0 Then
            sVaxNeeded = sVaccine
          Else
            sVaxNeeded = sVaxNeeded & ", " & sVaccine
          End If
        End Select
      End If
      sText = ""
      If IsNull(rstV![Vacc_Exempt]) Then
      Else
        If rstV![Vacc_Exempt] > 0 Then
          sText = SayVacc_Exempt(rstV![Vacc_Exempt])
        End If
      End If
      If Len(sText) > 0 Then
      Else
        If utNullToZero(rstV![Vacc_Date]) = 0 Then
        Else
          sText = Format$(rstV![Vacc_Date], "yymmdd")
        End If
        If rstV![VAcc_Current] Then
        Else
          sText = sText & " *"
        End If
      End If
      Select Case sVaccine
      Case "Anthrax"
        rstUR![Anthrax] = rstV![Vacc_Series] & " " & sText
      Case "Hep A"
        rstUR![Hep_A] = rstV![Vacc_Series] & " " & sText
      Case "Hep B"
        rstUR![Hep_B] = rstV![Vacc_Series] & " " & sText
      Case "Influenza"
        rstUR![Influenza] = " " & sText
      Case "IPPD"
        rstUR![IPPD] = " " & sText
        rstUR![IPPD_RESULT] = rstV![result] & vbNullString
      Case "Meningococcal"
        rstUR![Meningococcal] = " " & sText
      Case "MMR"
        rstUR![MMR] = " " & sText
      Case "Polio"
        rstUR![Polio] = " " & sText
      Case "Td"
        rstUR![Td] = " " & sText
      Case "Typhoid"
        rstUR![Typhoid] = " " & sText
      Case "Varicella"
        rstUR![Varicella] = " " & sText
      Case "Yellow Fever"
        rstUR![Yellow_Fever] = " " & sText
      Case Else
      End Select
        
      rstV.MoveNext
    Loop
NextPerson:
    rstUR![LastEvaluated] = Now
    rstUR.Update
    sSQL = "UPDATE Imm_Personnel SET VaxCurrent = " & iPerStat & ", NeededVax = '" & sVaxNeeded & "' WHERE SSAN = '" & rstP![SSAN] & "'"
    gobjSQL.ExecuteSQL sSQL
    rstP.MoveNext
  Loop
  SayStatus "Copying data to database..."
  sSQL = "DELETE * FROM Unit_Vaccine_Count"
  gobjSQL.ExecuteSQL sSQL
  With rstStat
    If .RecordCount > 0 Then
      .MoveFirst
      Do Until .EOF
        sSQL = "INSERT INTO Unit_Vaccine_Count (Unit, Vaccine, NbrReq, NbrCurrent, NbrDue, NbrOverdue, LastEval) VALUES ('" & ![Unit] & "','" & ![Vaccine] & "'," & utNullToZero(![NbrReq]) & "," & utNullToZero(![NbrCurrent]) & "," & utNullToZero(![NbrDue]) & "," & utNullToZero(![NbrOverdue]) & "," & SaySQLDateString(![LastEval]) & ")"
        gobjSQL.ExecuteSQL sSQL
        .MoveNext
      Loop
    End If
  End With
  
  sSQL = "DELETE * FROM Unit_Roster"
  gobjSQL.ExecuteSQL sSQL
  
  With rstUR
    .MoveFirst
    Do Until .EOF
      sSQL = "UPDATE Unit_Roster " & _
        "SET SSAN = '" & ![SSAN] & "'," & _
        "Anthrax = '" & ![Anthrax] & "'," & _
        "Hep_A = '" & ![Hep_A] & "'," & _
        "Hep_B = '" & ![Hep_B] & "'," & _
        "Influenza = '" & ![Influenza] & "'," & _
        "IPPD = '" & ![IPPD] & "'," & _
        "Meningococcal = '" & ![Meningococcal] & "'," & _
        "Td = '" & ![Td] & "'," & _
        "Typhoid = '" & ![Typhoid] & "'," & _
        "Yellow_Fever = '" & ![Yellow_Fever] & "'," & _
        "MMR = '" & ![MMR] & "'," & _
        "Polio = '" & ![Polio] & "'," & _
        "Varicella = '" & ![Varicella] & "'," & _
        "IPPD_RESULT= '" & ![IPPD_RESULT] & "'," & _
        "LastEvaluated = '" & ![LastEvaluated] & "'"
      gobjSQL.ExecuteSQL sSQL
      .MoveNext
    Loop
  End With
  On Error Resume Next
  rstP.Close
  Set rstP = Nothing
  rstV.Close
  Set rstV = Nothing
  rstUR.Close
  Set rstUR = Nothing
  rstR.Close
  Set rstR = Nothing
  On Error GoTo CalcUnit_Vaccine_Stats_Err
  SumUnitTotals
  
  If sType = "ALL" Then
    'CalcAFHS_Metric
  End If
  
  CalcUnit_Vaccine_Stats = True

CalcUnit_Vaccine_Stats_exit:
  On Error Resume Next
  RemoveMeter
  Screen.MousePointer = vbDefault
  SayStatus ""
  rstP.Close
  Set rstP = Nothing
  rstStat.Close
  Set rstStat = Nothing
  rstV.Close
  Set rstV = Nothing
  rstUR.Close
  Set rstUR = Nothing
  rstR.Close
  Set rstR = Nothing
  Beep
  Exit Function
  
CalcUnit_Vaccine_Stats_Err:
  ErrMsgBox "CalcUnit_Vaccine_Stats", Error$, Err
  Resume CalcUnit_Vaccine_Stats_exit
End Function


Public Sub SumUnitTotals()
  Dim rstT As ADODB.Recordset
  Dim rstU As ADODB.Recordset
  Dim sSQL As String
  On Error GoTo SumUnitTotals_Err
  Screen.MousePointer = vbHourglass
  SayStatus "Gathering unit totals..."
  Set rstU = New ADODB.Recordset
  sSQL = "SELECT Imm_Personnel.UNIT, Count(*) AS NbrAssigned, Sum(Abs(Imm_Personnel.VaxCurrent)) AS NbrCurrent From Imm_Personnel GROUP BY Imm_Personnel.UNIT"
  Set rstT = gobjSQL.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  With rstT
    Do Until .EOF
      Set rstU = gobjSQL.OpenRecordset("SELECT Unit, NbrAssigned, NbrCurrent FROM IMM_UNIT WHERE UNIT='" & rstT![Unit] & "'", adOpenForwardOnly, adLockOptimistic)
      If rstU.EOF And rstU.BOF Then
      Else
        sSQL = "UPDATE Imm_Unit SET NbrAggigned = " & rstT![NbrAssigned] & ", NbrCurrent = " & rstT![NbrCurrent] & " WHERE UNIT='" & rstT![Unit] & "'"
        gobjSQL.ExecuteSQL sSQL
      End If
      rstU.Close
      .MoveNext
    Loop
    .Close
  End With
  
SumUnitTotals_Exit:
  On Error Resume Next
  rstU.Close
  Set rstU = Nothing
  rstT.Close
  Set rstT = Nothing
  Screen.MousePointer = vbDefault
  SayStatus ""
  Exit Sub
  
SumUnitTotals_Err:
  ErrMsgBox "SumUnitTotals", Error$, Err
  Resume SumUnitTotals_Exit
End Sub



Public Function StuffPersonalVaxStatus(dStartDate As Date, dEndDate As Date) As Boolean
  Dim rstPLbl As ADODB.Recordset
  Dim rstP As ADODB.Recordset
  Dim rstPers As ADODB.Recordset
  Dim rstV As ADODB.Recordset
  Dim lCnt As Long
  Dim lCntPers As Long
  Dim sVax As String
  Dim sText As String
  Dim vLabel As Variant
  Dim sfield As String
  Dim sSQL As String
  Dim sSSAN As String
  Dim bExempt As Boolean
  Dim sNotCurrent As String
  On Error GoTo SPVS_Err
  Screen.MousePointer = vbHourglass
   
  gobjSQL.ExecuteSQL "DELETE  FROM tblPers_Vacc_Status_Labels"
  gobjSQL.ExecuteSQL "DELETE  FROM tblPers_Vacc_Status"
  If SayReportByUIC = 1 Then
    sSQL = "SELECT Imm_Personnel.SSAN, Imm_Personnel.UIC AS Unit FROM Imm_Personnel INNER JOIN tblUnit_Selected ON Imm_Personnel.UIC = tblUnit_Selected.UNIT"
  Else
    sSQL = "SELECT Imm_Personnel.SSAN, Imm_Personnel.Unit FROM Imm_Personnel INNER JOIN tblUnit_Selected ON Imm_Personnel.UNIT = tblUnit_Selected.UNIT"
  End If
  Set rstPers = gobjSQL.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  If rstPers.BOF And rstPers.EOF Then
    GoTo SPVS_Exit
  End If
  '!Fix
  'InitMeter "Checking status of personnel for Units selected...", rstPers.RecordCount, 0, True, "Unit Immunization Roster"
  Do Until rstPers.EOF
    sSSAN = rstPers![SSAN] & "20"
    lCntPers = lCntPers + 1
    'UpdateMeter lCntPers
    '...TJNLOOKAT
    Set rstV = GetLatestGenericVaccines(sSSAN, True, rstPers![Unit], , dEndDate)
    Do Until rstV.EOF
      sSQL = "SELECT tblPers_Vacc_Status.SPON_SSAN_FMP, tblPers_Vacc_Status.PT_SSAN, tblPers_Vacc_Status.Vax01, tblPers_Vacc_Status.Vax02, tblPers_Vacc_Status.Vax03, tblPers_Vacc_Status.Vax04, tblPers_Vacc_Status.Vax05, tblPers_Vacc_Status.Vax06, tblPers_Vacc_Status.Vax07, tblPers_Vacc_Status.Vax08, tblPers_Vacc_Status.Vax09, tblPers_Vacc_Status.Vax10, tblPers_Vacc_Status.Vax11 From tblPers_Vacc_Status WHERE (((tblPers_Vacc_Status.SPON_SSAN_FMP)= '" & sSSAN & "'))"
      Set rstP = gobjSQL.OpenRecordset(sSQL, adOpenForwardOnly, adLockOptimistic)
      If rstP.BOF And rstP.EOF Then
        rstP.AddNew
        rstP![SPON_SSAN_FMP] = sSSAN
        rstP![PT_SSAN] = rstPers![SSAN]
      End If
      sText = ""
      If IsNull(rstV![Vacc_Exempt]) Then
        bExempt = False
      Else
        If rstV![Vacc_Exempt] > 0 Then
          bExempt = True
          If rstV![ExemptionExpirationDate] >= dStartDate And rstV![ExemptionExpirationDate] <= dEndDate Then
            sNotCurrent = " *"
            sText = "Exemption Expired" & sNotCurrent
          Else
            sNotCurrent = ""
            sText = SayVacc_Exempt(rstV![Vacc_Exempt])
          End If
        End If
      End If
      If Len(sText) > 0 Then
      Else
        If rstV![Vacc_NextDue] >= dStartDate And rstV![Vacc_NextDue] <= dEndDate Then
          sNotCurrent = " *"
          sText = Format$(rstV![Vacc_Date], "yymmdd") & sNotCurrent
        Else
          If utNullToZero(rstV![Vacc_Date]) = 0 Then
            sNotCurrent = " *"
            sText = sNotCurrent
          Else
            sNotCurrent = ""
            sText = Format$(rstV![Vacc_Date], "yymmdd")
          End If
        End If
      End If
      Select Case rstV![GenericName]
      Case "IPPD"
        If Len(rstV![result]) = 0 Then
        Else
          sText = Format$(rstV![Vacc_Date], "yymmdd") & " " & rstV![result] & sNotCurrent
        End If
      Case "Influenza"
        sText = " " & sText
      Case "Meningococcal"
        sText = " " & sText
      Case "MMR"
        sText = " " & sText
      Case "Polio"
        sText = " " & sText
      Case "Td"
        sText = " " & sText
      Case "Typhoid"
        sText = " " & sText
      Case "Varicella"
        sText = " " & sText
      Case "Yellow Fever"
        sText = " " & sText
      Case Else
        If bExempt Then
          If Len(rstV![result]) = 0 Then
            sText = " " & sText
          Else
            sText = sText & rstV![result]
          End If
        Else
          sText = rstV![Vacc_Series] & " " & sText
        End If
      End Select
      rstP.Fields(lCnt + 2) = sText & ""
      rstP.Update
      rstP.Close
      sSQL = "SELECT tblPers_Vacc_Status_Labels.SPON_SSAN_FMP, tblPers_Vacc_Status_Labels.PT_SSAN, tblPers_Vacc_Status_Labels.lblVax01, tblPers_Vacc_Status_Labels.lblVax02, tblPers_Vacc_Status_Labels.lblVax03, tblPers_Vacc_Status_Labels.lblVax04, tblPers_Vacc_Status_Labels.lblVax05, tblPers_Vacc_Status_Labels.lblVax06, tblPers_Vacc_Status_Labels.lblVax07, tblPers_Vacc_Status_Labels.lblVax08, tblPers_Vacc_Status_Labels.lblVax09, tblPers_Vacc_Status_Labels.lblVax10, tblPers_Vacc_Status_Labels.lblVax11 From tblPers_Vacc_Status_Labels WHERE (((tblPers_Vacc_Status_Labels.SPON_SSAN_FMP)= '" & sSSAN & "'))"
      Set rstPLbl = gobjSQL.OpenRecordset(sSQL, adOpenForwardOnly, adLockOptimistic)
      If rstPLbl.BOF And rstPLbl.EOF Then
        rstPLbl.AddNew
        rstPLbl![SPON_SSAN_FMP] = sSSAN
        rstPLbl![PT_SSAN] = rstPers![SSAN]
      End If
      sVax = "lblVax" & Format$(lCnt, "00")
      rstPLbl.Fields(lCnt + 2) = rstV![GenericName] & ""
      rstPLbl.Update
      rstPLbl.Close
      rstV.MoveNext
      lCnt = lCnt + 1
      If lCnt > 10 Then
        Exit Do
      End If
    Loop
    lCnt = 0
    rstPers.MoveNext
    rstV.Close
  Loop

SPVS_Exit:
  On Error Resume Next
  rstPers.Close
  rstPLbl.Close
  rstP.Close
  rstV.Close
  Set rstPers = Nothing
  Set rstPLbl = Nothing
  Set rstP = Nothing
  Set rstV = Nothing
  Screen.MousePointer = vbDefault
  'RemoveMeter
  SayStatus ""
  Exit Function
SPVS_Err:
  lLastError = Err.Number
  sLastErrorDesc = "CITA_DATA: StuffPersonalVaxStatus: " & Err.Description
  Err.Raise lLastError, "CITA_DATA: StuffPersonalVaxStatus: ", sLastErrorDesc
End Function
Public Sub DeleteImm_People_Vaccine(sspon_SSAN_FMP As String, lVaccineID As Long, vVacc_Date As Variant, Optional bDelete As Boolean = True, Optional MergeRS As ADODB.Recordset = Nothing)
  Dim rstVacc As ADODB.Recordset
  Dim rstVaccDel As ADODB.Recordset
  Dim vItem As Variant
  Dim sSQL As String
  Dim sVaccDate As String
  On Error GoTo Err_Handler
  If utNullToZero(vVacc_Date) = 0 Then
    Exit Sub
  End If
   
  sVaccDate = SaySQLDateString(vVacc_Date)
  ' T.Noren  Take advantage of this code when deleting records
  '           in CDR that have been deleted in Theater (auto)
  If MergeRS Is Nothing Then
    sSQL = "SELECT Imm_People_Vaccine.* FROM Imm_People_Vaccine WHERE (SPON_SSAN_FMP='" & sspon_SSAN_FMP & "') AND (VaccineID=" & lVaccineID & ") AND (Vacc_Date=" & sVaccDate & ")"
    Set rstVacc = gobjSQL.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  Else
   Set rstVacc = MergeRS
  End If
    
    If rstVacc.BOF And rstVacc.EOF Then
    Else
      
    'Update the Imm_People_Vaccine_Delete table so the records can be deleted if neccessary
      
      
      ' T.Noren XXX Added Gbgems to force entry in Theater (auto)
      '<< SCR #42805; 42809;   Developer: Brian Mowbray 10/02/2003 03:22 PM
      If utNullToZero(rstVacc![LastDEERS_Ack]) = 0 And (Not gbITT Or Not gbGEMS) And MergeRS Is Nothing Then
      Else
        sSQL = "SELECT Imm_People_Vaccine_Delete.* FROM Imm_People_Vaccine_Delete WHERE (SPON_SSAN_FMP='" & sspon_SSAN_FMP & "') AND (VaccineID=" & lVaccineID & ") AND (Vacc_Date=" & sVaccDate & ")"
        Set rstVaccDel = gobjSQL.OpenRecordset(sSQL, adOpenForwardOnly, adLockOptimistic)
        If rstVaccDel.BOF And rstVaccDel.EOF Then
          sSQL = "INSERT INTO Imm_People_Vaccine_Delete (SPON_SSAN_FMP, VaccineID, Vacc_Series, Vacc_Date, MTF_Code, Deers_Instance, LastDEERS_Ack, LastEdited, DEERS_Action, LastEditedBy, InCDR) " & _
            "VALUES ('" & rstVacc![SPON_SSAN_FMP] & "' ," & CLng(rstVacc![VaccineID]) & " ,'" & rstVacc![Vacc_Series] & "' ," & SaySQLDateString(rstVacc![Vacc_Date]) & " ,'" & rstVacc![MTF_Code] & "' ,'" & rstVacc![Deers_Instance] & "' ," & SaySQLDateString(rstVacc![LastDEERS_Ack]) & " , " & SaySQLDateString(Now()) & " , '" & "D" & "' ,'" & CurrentUser() & "', " & utNullToZero(rstVacc![InCDR]) & ")"
          gobjSQL.ExecuteSQL sSQL
          ' T.Noren  Don't write audit if merging
          If MergeRS Is Nothing Then WriteAuditTrail sSQL

        End If
        rstVaccDel.Close
        Set rstVaccDel = Nothing
      End If
      
'This code was commented out because the requirement was removed to keep Wellness in sync with the immunization history. 12/17/2001 B. Mowbray

'      Update the Imm_People_Vaccine_Del_CDR if the value is 1 that means the CDR has been posted with this vaccine
'      If rstVacc![InCDR] = 1 Then
'        sSQL = "SELECT Imm_People_Vaccine_Del_CDR.* FROM Imm_People_Vaccine_Del_CDR WHERE (SPON_SSAN_FMP='" & sspon_SSAN_FMP & "') AND (VaccineID=" & lVaccineID & ") AND (Vacc_Date=" & sVaccDate & ")"
'        Set rstVaccDel = gobjSQL.OpenRecordset(sSQL, adOpenForwardOnly, adLockOptimistic)
'        If rstVaccDel.BOF And rstVaccDel.EOF Then
'          sSQL = "INSERT INTO Imm_People_Vaccine_Del_CDR (SPON_SSAN_FMP, VaccineID, Vacc_Series, Vacc_Date, MTF_Code, Deers_Instance, LastDEERS_Ack, LastEdited, DEERS_Action, LastEditedBy, InCDR) " & _
'            "VALUES ('" & rstVacc![SPON_SSAN_FMP] & "' ," & CLng(rstVacc![VaccineID]) & " ,'" & rstVacc![Vacc_Series] & "' ," & SaySQLDateString(rstVacc![Vacc_Date]) & " ,'" & rstVacc![MTF_Code] & "' ,'" & rstVacc![Deers_Instance] & "' ," & SaySQLDateString(rstVacc![LastDEERS_Ack]) & " , " & SaySQLDateString(Now()) & " , '" & "D" & "' ,'" & CurrentUser() & "', " & utNullToZero(rstVacc![InCDR]) & ")"
'          gobjSQL.ExecuteSQL sSQL
'        End If
'        rstVaccDel.Close
'        Set rstVaccDel = Nothing
'      End If
      
      'This flag is used for DEERS updates because there are situations where a delete will need to occur on DEERS but we don't want to
      'delete the record locally because the record was edited. DEERS never gets updated, records are deleted then added back
      'correctly if an edit was made. This ensures that the record is the same as what is in the global CDR
      
      If bDelete Then
        sSQL = "DELETE FROM Imm_People_Vaccine WHERE SPON_SSAN_FMP='" & sspon_SSAN_FMP & "' AND VaccineID=" & lVaccineID & " AND Vacc_Date=" & sVaccDate
        
        gobjSQL.ExecuteSQL sSQL
        
        WriteAuditTrail sSQL
      End If
    End If
    
Sub_exit:
  On Error Resume Next
  If MergeRS Is Nothing Then rstVacc.Close
  Set rstVacc = Nothing
  If Not rstVaccDel Is Nothing Then rstVaccDel.Close
  Set rstVaccDel = Nothing
  Exit Sub
  
Err_Handler:
  lLastError = Err.Number
  sLastErrorDesc = "Imm_DATA: DeleteImm_People_Vaccine: " & Err.Description
  Err.Raise lLastError, "Imm_DATA: DeleteImm_People_Vaccine:", sLastErrorDesc
  
End Sub


Private Function GetProgSetting(Section As String, key As String, Optional Default As String = vbNullString, Optional sIniFile As String = "CITA.INI") As String
  Dim iResult As Variant
  Dim sResult As String
  On Error Resume Next
  sResult = Space$(200)
  iResult = utGetPrivateProfileString(Section, key, Default, sResult, 200, sIniFile)
  GetProgSetting = Left$(sResult, iResult)
End Function

Public Sub Archive_Imm_People(sspon_SSAN_FMP As String)
  Dim sSQL As String
  On Error GoTo Err_Handler
   
  sSQL = "INSERT INTO Imm_People_Archive SELECT Imm_People.* From Imm_People WHERE (Imm_People.SPON_SSAN_FMP='" & sspon_SSAN_FMP & "')"
  gobjSQL.ExecuteSQL sSQL  ', , adExecuteNoRecords
  sSQL = "INSERT INTO Imm_People_Vaccine_Archive SELECT Imm_People_Vaccine.* From Imm_People_Vaccine WHERE (((Imm_People_Vaccine.SPON_SSAN_FMP)='" & sspon_SSAN_FMP & "'))"
  gobjSQL.ExecuteSQL sSQL  ', , adExecuteNoRecords
  gobjSQL.ExecuteSQL "DELETE FROM Imm_People_Vaccine WHERE SPON_SSAN_FMP='" & sspon_SSAN_FMP & "'"  ', , adExecuteNoRecords
  sSQL = "INSERT INTO Imm_People_GROUP_Archive SELECT Imm_People_Group.* From Imm_People_Group WHERE (((Imm_People_Group.SPON_SSAN_FMP)='" & sspon_SSAN_FMP & "'))"
  gobjSQL.ExecuteSQL sSQL  ' , , adExecuteNoRecords
  gobjSQL.ExecuteSQL "DELETE  FROM Imm_People_Group WHERE SPON_SSAN_FMP='" & sspon_SSAN_FMP & "'"  ', , adExecuteNoRecords
  gobjSQL.ExecuteSQL "DELETE  FROM Imm_People_Group WHERE SPON_SSAN_FMP='" & sspon_SSAN_FMP & "'"  ', , adExecuteNoRecords
  DeleteImm_People sspon_SSAN_FMP
Sub_exit:
  On Error Resume Next
  Exit Sub
  
Err_Handler:
  lLastError = Err.Number
  sLastErrorDesc = "CITA_DATA: Archive_Imm_People: " & Err.Description
  Err.Raise lLastError, "CITA_DATA: Archive_Imm_People: ", sLastErrorDesc
End Sub
Public Sub Archive_Imm_Personnel(sSSAN As String)
  Dim sSQL As String
  On Error GoTo Err_Handler
  sSQL = "INSERT INTO Imm_Personnel_Archive SELECT Imm_Personnel.* From Imm_Personnel WHERE (Imm_Personnel.SSAN='" & sSSAN & "')"
  gobjSQL.ExecuteSQL sSQL
  DeleteImm_Personnel sSSAN
Sub_exit:
  On Error Resume Next
  Exit Sub
  
Err_Handler:
  lLastError = Err.Number
  sLastErrorDesc = "CITA_DATA: Archive_Imm_Personnel: " & Err.Description
  Err.Raise lLastError, "CITA_DATA: Archive_Imm_Personnel: ", sLastErrorDesc
End Sub

Public Function DeleteImm_People(sspon_SSAN_FMP As String) As Boolean
  On Error GoTo DeleteImm_People_Err
   
  If Len(sspon_SSAN_FMP) > 0 Then
    gobjSQL.ExecuteSQL "DELETE * FROM Imm_People WHERE SPON_SSAN_FMP='" & sspon_SSAN_FMP & "'"
    WriteAuditTrail "DELETE * FROM Imm_People WHERE SPON_SSAN_FMP='" & sspon_SSAN_FMP & "'"
    DeleteImm_People = True
  End If
DeleteImm_People_exit:
  On Error Resume Next
  Exit Function

DeleteImm_People_Err:
  lLastError = Err.Number
  sLastErrorDesc = "CITA_DATA: DeleteImm_People: " & Err.Description
  Err.Raise lLastError, "CITA_DATA: DeleteImm_People: ", sLastErrorDesc
End Function
Public Sub DeleteImm_People_Archive(sspon_SSAN_FMP As String)
  On Error GoTo DeleteImm_People_Arc_Err
  If Len(sspon_SSAN_FMP) > 0 Then
    gobjSQL.ExecuteSQL "DELETE FROM Imm_People_Archive WHERE SPON_SSAN_FMP='" & sspon_SSAN_FMP & "'"  ', , adExecuteNoRecords
    gobjSQL.ExecuteSQL "DELETE FROM Imm_People_Group_Archive WHERE SPON_SSAN_FMP='" & sspon_SSAN_FMP & "'"  ', , adExecuteNoRecords
    gobjSQL.ExecuteSQL "DELETE FROM Imm_People_Vaccine_Archive WHERE SPON_SSAN_FMP='" & sspon_SSAN_FMP & "'"  ', , adExecuteNoRecords
  End If
DeleteImm_People_Arc_exit:
  On Error Resume Next
  Exit Sub

DeleteImm_People_Arc_Err:
  lLastError = Err.Number
  sLastErrorDesc = "CITA_DATA: DeleteImm_People_Archive: " & Err.Description
  Err.Raise lLastError, "CITA_DATA: DeleteImm_People_Archive: ", sLastErrorDesc
End Sub
Public Function DeleteImm_Personnel(sSSAN As String) As Boolean
  On Error GoTo DeleteImm_Personnel_Err
  If Len(sSSAN) > 0 Then
    gobjSQL.ExecuteSQL "DELETE FROM Imm_Personnel WHERE SSAN='" & sSSAN & "'"
    DeleteImm_Personnel = True
  End If
DeleteImm_Personnel_exit:
  On Error Resume Next
  Exit Function
DeleteImm_Personnel_Err:
  lLastError = Err.Number
  sLastErrorDesc = "CITA_DATA: DeleteImm_People: " & Err.Description
  Err.Raise lLastError, "CITA_DATA: DeleteImm_People: ", sLastErrorDesc
End Function
Public Function DeleteImm_Personnel_Archive(sSSAN As String) As Boolean
  On Error GoTo Err_Handler
  If Len(sSSAN) > 0 Then
    gobjSQL.ExecuteSQL "DELETE FROM Imm_Personnel_Archive WHERE SSAN='" & sSSAN & "'"  ', , adExecuteNoRecords
    DeleteImm_Personnel_Archive = True
  End If
Func_Exit:
  On Error Resume Next
  Exit Function
  
Err_Handler:
  lLastError = Err.Number
  sLastErrorDesc = "CITA_DATA: DeleteImm_Personnel_Archive: " & Err.Description
  Err.Raise lLastError, "CITA_DATA: DeleteImm_Personnel_Archive: ", sLastErrorDesc
End Function
Public Function GetImm_People_Archive(rst As ADODB.Recordset) As Boolean
'  Dim rstP As ADODB.Recordset
'  On Error GoTo GetImm_People_Archive_Err
''  If Len(sSPON_SSAN_FMP) = 11 Then
''    mvarSPON_SSAN_FMP = sSPON_SSAN_FMP
''    mvarFMP = Right(sSPON_SSAN_FMP, 2)
''    If mvarFMP = "20" Then
''      mvarSSAN = Left(sSPON_SSAN_FMP, 9)
''    End If
''  End If
'  If Len(rst![SPON_SSAN_FMP]) = 0 Then
'    Exit Function
'  End If
'  Set rstP = New ADODB.Recordset
'  If oConn.GetRecordSet(rstP, "SELECT * FROM Imm_People_Archive WHERE SPON_SSAN_FMP='" & rst![SPON_SSAN_FMP] & "'", adOpenForwardOnly, adLockReadOnly, adCmdText) Then
'  Else
'    Exit Function
'  End If
'  With rstP
'    If .BOF And .EOF Then
'    Else
'      If ![DataSource] = "D" Then
'        rst![DataSource] = "D"
'        rst![PT_FIRST_NAME] = ![PT_FIRST_NAME] & ""
'        rst![PT_MIDDLE_NAME] = ![PT_MIDDLE_NAME] & ""
'        rst![PT_LAST_NAME] = ![PT_LAST_NAME] & ""
'        rst![PT_SUFFIX] = ![PT_SUFFIX] & ""
'        rst![PT_DOB] = ![PT_DOB]
'        rst![PT_SSAN] = ![PT_SSAN] & ""
'      Else
'        If Len(rst![PT_SSAN]) = 0 Then
'          rst![PT_SSAN] = ![PT_SSAN] & ""
'        Else
'          If rst![PT_SSAN] <> ![PT_SSAN] Then mvarImm_People_Dirty = True
'        End If
'        If Len(mvarFirstName) = 0 Then
'          mvarFirstName = ![PT_FIRST_NAME] & ""
'        Else
'          If mvarFirstName <> ![PT_FIRST_NAME] Then mvarImm_People_Dirty = True
'        End If
'        If Len(mvarMiddleInitial) = 0 Then
'          mvarMiddleInitial = ![PT_MIDDLE_NAME] & ""
'        End If
'        If Len(mvarLastName) = 0 Then
'          mvarLastName = ![PT_LAST_NAME] & ""
'        Else
'          If mvarLastName <> ![PT_LAST_NAME] Then mvarImm_People_Dirty = True
'        End If
'        If Len(mvarSuffix) = 0 Then
'          mvarSuffix = ![PT_SUFFIX] & ""
'        End If
'        If IsNull(mvarDOB) Then
'          mvarDOB = ![PT_DOB]
'        Else
'          If mvarDOB <> ![PT_DOB] Then mvarImm_People_Dirty = True
'        End If
'      End If
'      If Len(mvarUnit) = 0 Then
'        mvarUnit = ![PT_Unit] & ""
'      Else
'        If mvarUnit <> ![PT_Unit] & "" Then
'          mvarImm_People_Dirty = True
'        End If
'      End If
'      If Len(mvarOffice_Symbol) = 0 Then
'        mvarOffice_Symbol = ![PT_Office_Symbol] & ""
'      Else
'        If mvarOffice_Symbol <> ![PT_Office_Symbol] Then mvarImm_People_Dirty = True
'      End If
'      If Len(mvarService) = 0 Then
'        mvarService = ![PT_Service] & ""
'      Else
'        If mvarService <> ![PT_Service] Then mvarImm_People_Dirty = True
'      End If
'      If Len(mvarStatus) = 0 Then
'        mvarStatus = ![PT_Status] & ""
'      Else
'        If mvarStatus <> ![PT_Status] Then mvarImm_People_Dirty = True
'      End If
'      If Len(mvarGender) = 0 Then
'        mvarGender = ![PT_Gender] & ""
'      Else
'        If mvarGender <> ![PT_Gender] Then mvarImm_People_Dirty = True
'      End If
'      If Len(mvarRank) = 0 Then
'        mvarRank = ![PT_Rank] & ""
'      Else
'        If mvarRank <> ![PT_Rank] Then mvarImm_People_Dirty = True
'      End If
'      mvarDEERS_DupID = ![DEERS_DupID] & ""
'      mvarDEERS_Date = ![DEERS_Date]
'      If Len(mvarDataSource) = 0 Then
'        mvarDataSource = ![DataSource] & ""
'      End If
'      mvarLastEdited = ![LastEdited]
'      mvarExemptionExpirationDate = ![ExemptionExpirationDate]
'      mvarExemptionCode = ![ExemptionCode] 'Removed deIMM.rscmmImm_People
'      mvarDEERS_DupID = ![DEERS_DupID]
'      mvarDEERS_Date = ![DEERS_Date]
'      mvarImm_People_Found = True
'      GetImm_People_Archive = True
'    End If
'  End With
'GetImm_People_Archive_Exit:
'  On Error Resume Next
'  rst.Close
'  Set rst = Nothing
'  Exit Function
'
'GetImm_People_Archive_Err:
  lLastError = Err.Number
  sLastErrorDesc = "CITA_DATA: GetImm_People_Archive: " & Err.Description
  Err.Raise lLastError, "CITA_DATA: GetImm_People_Archive: ", sLastErrorDesc
End Function
Public Function Recall_Imm_People_Archive(sspon_SSAN_FMP As String) As Boolean
'  Dim sSQL As String
'  Dim rst As ADODB.Recordset
'  Dim rstA As ADODB.Recordset
'  On Error GoTo Err_Handler
'  If Len(sSPON_SSAN_FMP) = 11 Then
'  Else
'    Exit Function
'  End If
'  Set rst = New ADODB.Recordset
'
'  mvarSPON_SSAN_FMP = sSPON_SSAN_FMP
'  mvarFMP = Right$(sSPON_SSAN_FMP, 2)
'  If mvarFMP = "20" Then
'    mvarSSAN = Left$(sSPON_SSAN_FMP, 9)
'  End If
'  If GetImm_People_Archive(sSPON_SSAN_FMP) Then
'  Else
'    GoTo Function_Exit
'  End If
'  UpdateImm_People
'  Set rstA = New ADODB.Recordset
'  If oConn.GetRecordSet(rstA, "SELECT Imm_People_Group_Archive.Spon_SSAN_FMP, Imm_People_Group_Archive.Group_Name From Imm_People_Group_Archive WHERE Imm_People_Group_Archive.Spon_SSAN_FMP='" & mvarSPON_SSAN_FMP & "'", adOpenForwardOnly, adLockOptimistic, adCmdText) Then
'  Else
'    GoTo Function_Exit
'  End If
'  With rstA
'    If .BOF And .EOF Then
'    Else
'      Do Until rstA.EOF
'        If oConn.GetRecordSet(rst, "SELECT Imm_People_Group.SPON_SSAN_FMP, Imm_People_Group.Group_Name FROM Imm_People_Group WHERE Imm_People_Group.SPON_SSAN_FMP='" & mvarSPON_SSAN_FMP & "'", adOpenForwardOnly, adLockOptimistic, adCmdText) Then
'        Else
'          GoTo Function_Exit
'        End If
'        rst.AddNew
'        rst![SPON_SSAN_FMP] = mvarSPON_SSAN_FMP
'        rst![Group_Name] = rstA![Group_Name]
'        On Error Resume Next
'        rst.Update
'        ' If an error is blown because of a duplicate value, press on!
'        If err <> 0 Then
'          rst.CancelUpdate
'        End If
'        rst.Close
'        rstA.MoveNext
'      Loop
'      rst.Close
'      Set rst = Nothing
'    End If
'  End With
'  rstA.Close
'  Set rstA = Nothing
'  oConn.Execute "DELETE  FROM Imm_People_Group_Archive WHERE SPON_SSAN_FMP='" & mvarSPON_SSAN_FMP & "'", , adExecuteNoRecords
'  sSQL = "INSERT INTO Imm_People_Vaccine "
'  sSQL = sSQL & "SELECT Imm_People_Vaccine_Archive.* "
'  sSQL = sSQL & "From Imm_People_Vaccine_Archive "
'  sSQL = sSQL & "WHERE (((Imm_People_Vaccine_Archive.SPON_SSAN_FMP)='" & mvarSPON_SSAN_FMP & "'))"
'  oConn.Execute sSQL, , adExecuteNoRecords
'  oConn.Execute "DELETE  FROM Imm_People_Vaccine_Archive WHERE SPON_SSAN_FMP='" & mvarSPON_SSAN_FMP & "'", , adExecuteNoRecords
'  oConn.Execute "DELETE  FROM Imm_People_Archive WHERE SPON_SSAN_FMP='" & mvarSPON_SSAN_FMP & "'", , adExecuteNoRecords
'  Recall_Imm_People_Archive = True
'Function_Exit:
'  On Error Resume Next
'  Exit Function
'
'Err_Handler:
  lLastError = Err.Number
  sLastErrorDesc = "CITA_DATA: Recall_Imm_People_Archive: " & Err.Description
  Err.Raise lLastError, "CITA_DATA: Recall_Imm_People_Archive: ", sLastErrorDesc
End Function
Public Sub UpdateImm_People(rstP As ADODB.Recordset, Action As UpdateType)
  Dim sSQL As String
  On Error GoTo UIP_Err
  If rstP.BOF And rstP.EOF Then
  Else
    Select Case Action
    Case 1 ' Append
      'SCR #22333;   Developer: Brian Mowbray 03/20/2002 10:49 AM
      '>> SCR #51212;   Developer: Brian Mowbray 03/17/2004 07:42 AM
      'Added Left$() to handle invlaid lengths on PT_LAST_NAME, PT_FIRST_NAME and PT_MIDDLE_NAME columns
      sSQL = "INSERT INTO Imm_People (SPON_SSAN_FMP, PT_SSAN, PT_LAST_NAME, PT_FIRST_NAME, PT_MIDDLE_NAME, PT_SUFFIX, PT_DOB, PT_GENDER, PT_STATUS, " & _
        "PT_RANK, PT_Service, PT_Unit, Pt_OFFICE_SYMBOL, DataSource, LastEdited, ExemptionExpirationDate, ExemptionCode ) " & _
        "VALUES ('" & rstP![SPON_SSAN_FMP] & "' ,'" & rstP![PT_SSAN] & "' ,'" & ValidateDBChar(Left$(rstP![PT_LAST_NAME] & vbNullString, 40)) & _
          "' ,'" & ValidateDBChar(Left$(rstP![PT_First_Name] & vbNullString, 30)) & "' ,'" & ValidateDBChar(Left$(rstP![PT_MIDDLE_NAME] & vbNullString, 1)) & _
          "' ,'" & ValidateDBChar(rstP![PT_SUFFIX] & vbNullString) & "' ," & SaySQLDateString(rstP![PT_DOB]) & " ,'" & rstP![PT_GENDER] & "' ,'" & rstP![PT_Status] & "' ,'" & ValidateDBChar(rstP![PT_RANK] & vbNullString) & _
          "' ,'" & rstP![PT_Service] & "' ,'" & ValidateDBChar(rstP![Pt_Unit] & vbNullString) & "' ,'" & ValidateDBChar(rstP![PT_OFFICE_SYMBOL] & vbNullString) & _
          "' ,'" & rstP![DataSource] & "' ," & SaySQLDateString(Now()) & "  ," & SaySQLDateString(rstP![ExemptionExpirationDate]) & _
          " ," & utNullToZero(rstP![ExemptionCode]) & " )"
      gobjSQL.ExecuteSQL sSQL
      WriteAuditTrail sSQL
    Case 2 ' Update
      'SCR #22333;   Developer: Brian Mowbray 03/20/2002 10:49 AM
      '>> SCR #51212;   Developer: Brian Mowbray 03/17/2004 07:42 AM
      'Added Left$() to handle invlaid lengths on PT_LAST_NAME, PT_FIRST_NAME and PT_MIDDLE_NAME columns
      sSQL = "UPDATE Imm_People SET SPON_SSAN_FMP = '" & rstP![SPON_SSAN_FMP] & "'" & " , PT_SSAN = '" & rstP![PT_SSAN] & "'" & " , PT_LAST_NAME = '" & ValidateDBChar(Left$(rstP![PT_LAST_NAME] & vbNullString, 40)) & "'" & " , PT_FIRST_NAME = '" & ValidateDBChar(Left$(rstP![PT_First_Name] & vbNullString, 30)) & "'" & " , PT_MIDDLE_NAME = '" & ValidateDBChar(Left$(rstP![PT_MIDDLE_NAME] & vbNullString, 1)) & "'" & _
        " , PT_SUFFIX = '" & ValidateDBChar(rstP![PT_SUFFIX] & vbNullString) & "'" & " , PT_DOB = " & SaySQLDateString(rstP![PT_DOB]) & " , PT_GENDER = '" & rstP![PT_GENDER] & "'" & " , PT_STATUS = '" & rstP![PT_Status] & "'" & " , PT_RANK = '" & ValidateDBChar(rstP![PT_RANK] & vbNullString) & "'" & " , PT_Service = '" & rstP![PT_Service] & "'" & _
        " , PT_Unit= '" & ValidateDBChar(rstP![Pt_Unit] & vbNullString) & "'" & " , Pt_OFFICE_SYMBOL = '" & ValidateDBChar(rstP![PT_OFFICE_SYMBOL] & vbNullString) & "'" & " , DataSource = '" & rstP![DataSource] & "'" & " , LastEdited = " & SaySQLDateString(rstP![LastEdited]) & " , ExemptionExpirationDate = " & SaySQLDateString(rstP![ExemptionExpirationDate]) & " , ExemptionCode = " & utNullToZero(rstP![ExemptionCode]) & ", Comments='" & ValidateDBChar(rstP![Comments] & vbNullString) & "' " & _
        "WHERE Imm_People.SPON_SSAN_FMP = '" & rstP![SPON_SSAN_FMP] & "'"
      gobjSQL.ExecuteSQL sSQL
      WriteAuditTrail sSQL
    Case 3 ' Delete
      sSQL = "DELETE FROM Imm_People WHERE Imm_People.SPON_SSAN_FMP = '" & rstP![SPON_SSAN_FMP] & "'"
      gobjSQL.ExecuteSQL sSQL
      WriteAuditTrail sSQL
    Case 4 ' Archive
      ' Remove any previous archive records since this is just demographics
      On Error Resume Next
      sSQL = "DELETE FROM Imm_People_Archive WHERE Imm_People_Archive.SPON_SSAN_FMP = '" & rstP![SPON_SSAN_FMP] & "'"
      gobjSQL.ExecuteSQL sSQL
      ' Insert records into archive table
      sSQL = "INSERT INTO Imm_People_Archive (SPON_SSAN_FMP, PT_SSAN, PT_LAST_NAME, PT_FIRST_NAME, PT_MIDDLE_NAME, PT_SUFFIX, PT_DOB, PT_GENDER, PT_STATUS, PT_RANK, PT_Service, PT_Unit, Pt_OFFICE_SYMBOL, DataSource, LastEdited, ExemptionExpirationDate, ExemptionCode ) " & _
        "VALUES (rstP![SPON_SSAN_FMP], rstP![PT_SSAN], rstP![PT_LAST_NAME], rstP![PT_FIRST_NAME], rstP![PT_MIDDLE_NAME], rstP![PT_SUFFIX], rstP![PT_DOB], rstP![PT_GENDER], rstP![PT_STATUS], rstP![PT_RANK], rstP![PT_Service], rstP![PT_Unit], rstP![Pt_OFFICE_SYMBOL], rstP![DataSource], " & Format$(Now(), "dd mmmm yyyy:HH:nn:ss") & "  , rstP![ExemptionExpirationDate], rstP![ExemptionCode])"
      gobjSQL.ExecuteSQL sSQL
      ' Delete record from Imm_People Table
      sSQL = "DELETE FROM Imm_People WHERE Imm_People.SPON_SSAN_FMP = '" & rstP![SPON_SSAN_FMP] & "'"
      gobjSQL.ExecuteSQL sSQL
    End Select
    'Resync Groups with changes just made
     CheckImm_People_Group rstP![SPON_SSAN_FMP] & vbNullString, rstP![PT_Status], rstP![PT_Service], rstP![PT_DOB]
  End If
  
UIP_Exit:
  On Error Resume Next
  Exit Sub
  
UIP_Err:
  lLastError = Err.Number
  sLastErrorDesc = "Imm_DATA: UpdateImm_People: " & Err.Description
  Err.Raise lLastError, "Imm_DATA: UpdateImm_People: ", sLastErrorDesc
End Sub
Sub CheckImm_People_Group(sspon_SSAN_FMP As String, vStatus As Variant, vService As Variant, Optional vDOB As Variant = Null)
  Dim sServiceStatus As String
  On Error Resume Next
  '>> SCR #38820;   Developer: Brian Mowbray 07/21/2003 10:45 AM
  If Len(sspon_SSAN_FMP) > 0 Then
 
    If right$(sspon_SSAN_FMP, 2) <> 20 Then 'Prevent Family Member from being assigned a group
      Exit Sub
    End If

   If IsNull(vStatus) Then
     Exit Sub
   End If
   If IsNull(vService) Then
     Exit Sub
   End If
   sServiceStatus = vService & vStatus
   Select Case sServiceStatus
   Case "AAD"
     AddGroup sspon_SSAN_FMP, "Active Duty (Army)"
   Case "ARV", "ANG"
     AddGroup sspon_SSAN_FMP, "Reserve (Army)"
   Case "BAD"
     AddGroup sspon_SSAN_FMP, "Active Duty (Navy)"
   Case "BRV", "BNG"
     AddGroup sspon_SSAN_FMP, "Reserve (Navy)"
   Case "CAD"
     AddGroup sspon_SSAN_FMP, "Active Duty (AF)"
   Case "CRV"
     AddGroup sspon_SSAN_FMP, "Reserve (AF)"
   Case "CNG"
     AddGroup sspon_SSAN_FMP, "Air National Guard"
   Case "DAD"
     AddGroup sspon_SSAN_FMP, "Active Duty (Marine)"
   Case "DRV", "DNG"
     AddGroup sspon_SSAN_FMP, "Reserve (Marine)"
   Case "EAD"
     AddGroup sspon_SSAN_FMP, "Active Duty (CG)"
   Case "ERV", "ENG"
     AddGroup sspon_SSAN_FMP, "Reserve (CG)"
   End Select
  End If
End Sub
Sub AddGroup(sspon_SSAN_FMP As String, sGroup As String)
  Dim sSQL As String
  Dim rst As ADODB.Recordset
  On Error GoTo AddGroup_Err
  '>> SCR #38820;   Developer: Brian Mowbray 07/21/2003 10:45 AM
  If Len(sspon_SSAN_FMP) > 0 Then
    If Len(sGroup) > 0 Then
      sSQL = "SELECT SPON_SSAN_FMP, Group_Name, MTF_ID FROM Imm_People_Group WHERE (Spon_SSAN_FMP ='" & sspon_SSAN_FMP & "' AND Group_Name ='" & ValidateDBChar(sGroup) & "' AND MTF_ID='" & SayMTF_ID() & "')"
      Set rst = gobjSQL.OpenRecordset(sSQL, adOpenForwardOnly, adLockOptimistic)
      If rst.BOF And rst.EOF Then
        sSQL = "INSERT INTO Imm_People_Group (SPON_SSAN_FMP, GROUP_NAME, MTF_ID) VALUES ('" & sspon_SSAN_FMP & "' ,'" & ValidateDBChar(sGroup) & "', '" & SayMTF_ID() & "')"
        gobjSQL.ExecuteSQL sSQL
        WriteAuditTrail sSQL
       ' bookmark
       '<< SCR #42805; 42809;   Developer: Brian Mowbray 10/02/2003 03:22 PM
       If Not gbITT Then
           ' Add this group to Patient's dictionary
           PatientIMM.AddGroupToPatient sGroup, True
       End If
      End If
    End If
  End If
AddGroup_Exit:
  On Error Resume Next
  rst.Close
  Set rst = Nothing
  Exit Sub

AddGroup_Err:
  lLastError = Err.Number
  sLastErrorDesc = "CITA_DATA: AddGroup: " & Err.Description
  Err.Raise lLastError, "CITA_DATA: AddGroup: ", sLastErrorDesc
End Sub
Public Function GetLatestGenericVaccines(sspon_SSAN_FMP As String, Optional RequiredOnly As Boolean = True, Optional sUnit As String = "", Optional vDOB As Variant = Null, Optional vEvalDate = Null) As ADODB.Recordset
  '<< SCR #42805; 42809;   Developer: Brian Mowbray 10/02/2003 03:22 PM
  If Not gbITT Then
   '...08/08/01  T.Noren
    Set GetLatestGenericVaccines = PatientIMM.BuildGenericRecordSet(RequiredOnly)
    Exit Function
  End If
  Dim rstGroup As ADODB.Recordset
  Dim sSQL As String
  Dim bEdit As Boolean
  Dim sGeneric As String
  Dim lDaysOld As Long
  Dim rst As ADODB.Recordset
  Dim rstG As ADODB.Recordset
  On Error GoTo GLGV_Err
   
  If utNullToZero(vEvalDate) = 0 Then
    vEvalDate = Date
  End If
'  Set rst = New ADODB.Recordset
'  If rst.Fields.Count = 0 Then
'    rst.Fields.Append "GenericName", adVarChar, 50, adFldKeyColumn
'    rst.Fields.Append "VaccineID", adInteger, , adFldIsNullable
'    rst.Fields.Append "Vacc_Series", adInteger, , adFldIsNullable
'    rst.Fields.Append "Vacc_Date", adDate, , adFldIsNullable
'    rst.Fields.Append "Vacc_NextDue", adDate, , adFldIsNullable
'    rst.Fields.Append "Vacc_Current", adBoolean, , False
'    rst.Fields.Append "Vacc_Exempt", adInteger, , adFldIsNullable
'    rst.Fields.Append "ExemptionExpirationDate", adDate, , adFldIsNullable
'    rst.Fields.Append "Result", adChar, 1, adFldIsNullable
'  End If
'  rst.Open , , adOpenDynamic, adLockOptimistic

  '<< Begin:SCR #37777;   Developer: Brian Mowbray 06/16/2003 05:00 PM
  CreateRst Generic_rst, rst
  '>> End: SCR #37777;
  
  If right$(sspon_SSAN_FMP, 2) <> "20" Then
    If utNullToZero(vDOB) = 0 Then
    Else
      lDaysOld = DateDiff("d", vDOB, Date)
      If lDaysOld < 60 Then
        rst.AddNew
        rst![GenericName] = "Hep B"
        rst.Update
        'GoTo LoadVaccineData
      ElseIf lDaysOld < 365 Then  '1 Year
        rst.AddNew
        rst![GenericName] = "Hep B"
        rst.Update
        rst.AddNew
        rst![GenericName] = "DTaP"
        rst.Update
        rst.AddNew
        rst![GenericName] = "Hib"
        rst.Update
        rst.AddNew
        rst![GenericName] = "Polio"
        rst.Update
        'GoTo LoadVaccineData
      ElseIf lDaysOld < 4015 Then  '11 years
        rst.AddNew
        rst![GenericName] = "Hep B"
        rst.Update
        rst.AddNew
        rst![GenericName] = "DTaP"
        rst.Update
        rst.AddNew
        rst![GenericName] = "Hib"
        rst.Update
        rst.AddNew
        rst![GenericName] = "Polio"
        rst.Update
        rst.AddNew
        rst![GenericName] = "MMR"
        rst.Update
        rst.AddNew
        rst![GenericName] = "Varicella"
        rst.Update
        'GoTo LoadVaccineData
      ElseIf lDaysOld < 6570 Then  '18 years
        rst.AddNew
        rst![GenericName] = "Hep B"
        rst.Update
        rst.AddNew
        rst![GenericName] = "Td"
        rst.Update
        rst.AddNew
        rst![GenericName] = "Polio"
        rst.Update
        rst.AddNew
        rst![GenericName] = "MMR"
        rst.Update
        rst.AddNew
        rst![GenericName] = "Varicella"
        rst.Update
        'GoTo LoadVaccineData
      End If
    End If
  End If
  sSQL = "SELECT DISTINCT tblVaccine_Groups_Vaccine.GenericName " & _
    "FROM tblVaccine_Groups_Vaccine, Imm_People_Group " & _
    "WHERE (Imm_People_Group.Group_Name = tblVaccine_Groups_Vaccine.Group_Name) " & _
    "AND (Imm_People_Group.SPON_SSAN_FMP='" & sspon_SSAN_FMP & "') " & _
    "UNION " & _
    "SELECT DISTINCT User_Vaccine_Groups_Vaccine.GenericName " & _
    "FROM Imm_People_Group, User_Vaccine_Groups_Vaccine " & _
    "WHERE (Imm_People_Group.Group_Name = User_Vaccine_Groups_Vaccine.Group_Name) " & _
    "AND (Imm_People_Group.SPON_SSAN_FMP='" & sspon_SSAN_FMP & "') "
    If Len(sUnit) > 0 Then
      sSQL = sSQL & "Union " & _
        "SELECT tblVaccine_Groups_Vaccine.GenericName " & _
        "FROM Imm_Unit_Group, tblVaccine_Groups_Vaccine " & _
        "WHERE Imm_Unit_Group.Group_Name = tblVaccine_Groups_Vaccine.Group_Name " & _
        "AND Imm_Unit_Group.UNIT='" & ValidateDBChar(sUnit) & "' " & _
        "UNION " & _
        "SELECT User_Vaccine_Groups_Vaccine.GenericName " & _
        "FROM Imm_Unit_Group, User_Vaccine_Groups_Vaccine " & _
        "WHERE Imm_Unit_Group.Group_Name = User_Vaccine_Groups_Vaccine.Group_Name " & _
        "AND Imm_Unit_Group.UNIT='" & ValidateDBChar(sUnit) & "'"
    End If
  
  
  
  
  Set rstGroup = gobjSQL.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  If rstGroup.BOF And rst.EOF Then
    If RequiredOnly Then
      If rst.EOF And rst.BOF Then
      Else
        rst.MoveFirst
      End If
      Set GetLatestGenericVaccines = rst
      Exit Function
' <-----------------
    End If
  End If
  Do Until rstGroup.EOF
    With rst
      '<< Begin:SCR #767;#811   Developer: Brian Mowbray 09/27/2002 01:12 PM
      If Not .BOF And Not .EOF Then
        .MoveFirst
        .Find "GenericName='" & Trim(rstGroup![GenericName]) & "'"
      End If
      '>> End: SCR #811;
      If .EOF Then
        .AddNew
        ![GenericName] = Trim(rstGroup![GenericName])
        ![VaccineID] = Null
        ![Vacc_Series] = Null
        ![Vacc_Date] = Null
        ![Vacc_NextDue] = Null
        ![VAcc_Current] = False
        ![Vacc_Exempt] = 0
        ![ExemptionExpirationDate] = Null
        ![result] = ""
        .Update
      End If
    '>> End: SCR #767;
    End With
    rstGroup.MoveNext
  Loop
  rstGroup.Close
  
LoadVaccineData:

  If rst.EOF And rst.BOF Then
  Else
    rst.MoveFirst
  End If
    
  sSQL = "SELECT Imm_People_Vaccine.SPON_SSAN_FMP, tblVaccines.GenericName, Imm_People_Vaccine.Vacc_Date, Imm_People_Vaccine.VaccineID, Imm_People_Vaccine.Vacc_Series, Imm_People_Vaccine.Vacc_NextDue, Imm_People_Vaccine.Vacc_Exempt, Imm_People_Vaccine.ExemptionExpirationDate, Imm_People_Vaccine.RXN_Size, Imm_People_Vaccine.Result, Imm_People_Vaccine.Mfg_Code, Imm_People_Vaccine.Lot_Nbr, tblVaccines.Vaccine "
  sSQL = sSQL & "FROM Imm_People_Vaccine, tblVaccines "
  sSQL = sSQL & "WHERE Imm_People_Vaccine.VaccineID = tblVaccines.VaccineID AND (Imm_People_Vaccine.SPON_SSAN_FMP ='" & sspon_SSAN_FMP & "') ORDER BY tblVaccines.GenericName, Imm_People_Vaccine.Vacc_Date DESC"
  
  
  
'  sSQL = "SELECT Imm_People_Vaccine.SPON_SSAN_FMP, tblVaccines.GenericName, Imm_People_Vaccine.Vacc_Date, Imm_People_Vaccine.VaccineID, Imm_People_Vaccine.Vacc_Series, Imm_People_Vaccine.Vacc_NextDue, Imm_People_Vaccine.Vacc_Exempt, Imm_People_Vaccine.ExemptionExpirationDate, Imm_People_Vaccine.RXN_Size, Imm_People_Vaccine.Result, Imm_People_Vaccine.Mfg_Code, Imm_People_Vaccine.Lot_Nbr, tblVaccines.Vaccine "
'  sSQL = sSQL & "FROM Imm_People_Vaccine, tblVaccines "
'  sSQL = sSQL & "WHERE Imm_People_Vaccine.VaccineID = tblVaccines.VaccineID AND (Imm_People_Vaccine.SPON_SSAN_FMP ='" & sspon_SSAN_FMP & "') ORDER BY tblVaccines.GenericName, Imm_People_Vaccine.Vacc_Date DESC"
  Set rstG = gobjSQL.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  
  
  Do Until rstG.EOF
    bEdit = False
    If rst.BOF And rst.EOF Then
    Else
      rst.MoveFirst
    End If
    If sGeneric <> rstG![GenericName] Then
      sGeneric = rstG![GenericName]
'      Commented out so Anthrax will only show due when the person has the CENTCOM AOR or Anthrax assigned

'      If sGeneric = "Anthrax" Then
'        rst.Find "GenericName='Anthrax'"
'        If rst.EOF Then
'          rst.AddNew
'          rst![GenericName] = "Anthrax"
'          bEdit = True
'        Else
'          bEdit = True
'        End If
'      Else
        rst.Find "GenericName='" & sGeneric & "'"
        If rst.EOF Then
          If RequiredOnly Then
            bEdit = False
            If sGeneric = "MMR" Then
              GoTo UpdateMMR
            End If
          Else
            rst.AddNew
            rst![GenericName] = sGeneric
            bEdit = True
          End If
        Else
          bEdit = True
        End If
'      End If
      If bEdit Then
        With rst
          ![VaccineID] = rstG![VaccineID]
          ![Vacc_Series] = rstG![Vacc_Series]
          ![Vacc_Date] = rstG![Vacc_Date]
          ![Vacc_NextDue] = rstG![Vacc_NextDue]
          If utNullToZero(rstG![Vacc_NextDue]) = 0 Then
            ![VAcc_Current] = True
          Else
            ![VAcc_Current] = (rstG![Vacc_NextDue] > vEvalDate)
          End If
          ![Vacc_Exempt] = rstG![Vacc_Exempt]
          ![ExemptionExpirationDate] = rstG![ExemptionExpirationDate]
          ![result] = rstG![result]
          .Update
        End With
        If sGeneric = "Hib" And rstG![VaccineID] = 51 Then
          rst.MoveFirst
          rst.Find "GenericName='Hep B'"
          If rst.EOF Then
            rst.AddNew
            rst![GenericName] = "Hep B"
          End If
          With rst
            ![VaccineID] = rstG![VaccineID]
            ![Vacc_Series] = rstG![Vacc_Series]
            ![Vacc_Date] = rstG![Vacc_Date]
            ![Vacc_NextDue] = rstG![Vacc_NextDue]
            If utNullToZero(rstG![Vacc_NextDue]) = 0 Then
              ![VAcc_Current] = True
            Else
              ![VAcc_Current] = (rstG![Vacc_NextDue] > vEvalDate)
            End If
            ![Vacc_Exempt] = rstG![Vacc_Exempt]
            ![ExemptionExpirationDate] = rstG![ExemptionExpirationDate]
            ![result] = rstG![result]
            .Update
          End With
        End If

UpdateMMR:
        
        If sGeneric = "MMR" And rstG![VaccineID] = 3 Then
          rst.MoveFirst
          rst.Find "GenericName='Mumps'"
          If rst.EOF Then
          Else
            With rst
              ![VaccineID] = rstG![VaccineID]
              ![Vacc_Series] = rstG![Vacc_Series]
              ![Vacc_Date] = rstG![Vacc_Date]
              ![Vacc_NextDue] = rstG![Vacc_NextDue]
              If utNullToZero(rstG![Vacc_NextDue]) = 0 Then
                ![VAcc_Current] = True
              Else
                ![VAcc_Current] = (rstG![Vacc_NextDue] > vEvalDate)
              End If
              ![Vacc_Exempt] = rstG![Vacc_Exempt]
              ![ExemptionExpirationDate] = rstG![ExemptionExpirationDate]
              ![result] = rstG![result]
              .Update
            End With
          End If
        End If
      
        If sGeneric = "MMR" And rstG![VaccineID] = 3 Then
          rst.MoveFirst
          rst.Find "GenericName='Rubella'"
          If rst.EOF Then
          
          Else
            With rst
              ![VaccineID] = rstG![VaccineID]
              ![Vacc_Series] = rstG![Vacc_Series]
              ![Vacc_Date] = rstG![Vacc_Date]
              ![Vacc_NextDue] = rstG![Vacc_NextDue]
              If utNullToZero(rstG![Vacc_NextDue]) = 0 Then
                ![VAcc_Current] = True
              Else
                ![VAcc_Current] = (rstG![Vacc_NextDue] > vEvalDate)
              End If
              ![Vacc_Exempt] = rstG![Vacc_Exempt]
              ![ExemptionExpirationDate] = rstG![ExemptionExpirationDate]
              ![result] = rstG![result]
              .Update
            End With
          End If
        End If
        
        If sGeneric = "MMR" And rstG![VaccineID] = 3 Then
          rst.MoveFirst
          rst.Find "GenericName='Rubella/Mumps'"
          If rst.EOF Then
          
          Else
            With rst
              ![VaccineID] = rstG![VaccineID]
              ![Vacc_Series] = rstG![Vacc_Series]
              ![Vacc_Date] = rstG![Vacc_Date]
              ![Vacc_NextDue] = rstG![Vacc_NextDue]
              If utNullToZero(rstG![Vacc_NextDue]) = 0 Then
                ![VAcc_Current] = True
              Else
                ![VAcc_Current] = (rstG![Vacc_NextDue] > vEvalDate)
              End If
              ![Vacc_Exempt] = rstG![Vacc_Exempt]
              ![ExemptionExpirationDate] = rstG![ExemptionExpirationDate]
              ![result] = rstG![result]
              .Update
            End With
          End If
        End If

        
      
      End If
    End If
    rstG.MoveNext
  Loop
  If rst.EOF And rst.BOF Then
  Else
    rst.MoveFirst
  End If
  Set GetLatestGenericVaccines = rst
GLGV_Exit:
  On Error Resume Next
  rstGroup.Close
  Set rstGroup = Nothing
  rstG.Close
  Set rstG = Nothing
  Exit Function
  
GLGV_Err:
  lLastError = Err.Number
  sLastErrorDesc = "Imm_DATA: GetLatestGenericVaccines: " & Err.Description
  Err.Raise lLastError, "Imm_DATA: GetLatestGenericVaccines: ", sLastErrorDesc
End Function

Public Function GetImm_Personnel_Archive(Optional sSSAN As String) As Boolean
'  Dim rst As ADODB.Recordset
'  On Error GoTo GIP_Err
'  If Len(sSSAN) = 9 Then
'    ClearPerson
'    mvarSSAN = sSSAN
'  End If
'  If Len(mvarSSAN) = 0 Then
'    Exit Function
'  End If
'  Set rst = New ADODB.Recordset
'  If oConn.GetRecordSet(rst, "SELECT * FROM Imm_Personnel_Archive WHERE SSAN='" & mvarSSAN & "'", adOpenForwardOnly, adLockReadOnly, adCmdText) Then
'  Else
'    GoTo GIP_Exit
'  End If
'  With rst
'    If .BOF And .EOF Then
'    Else
'      SplitName ![Name], mvarLastName, mvarFirstName, mvarMiddleInitial, mvarSuffix
'      mvarDOB = ![DOB]
'      mvarUnit = ![Unit] & ""
'      mvarOffice_Symbol = ![OFFICE_SYMBOL] & ""
'      mvarService = ![Service] & ""
'      mvarStatus = ![Status] & ""
'      mvarGender = ![Sex] & ""
'      mvarRank = ![RANK] & ""
'      mvarUIC = ![UIC] & ""
'      mvarDAS = ![DAS] & ""
'      mvarWorkPhone = ![PHONE] & ""
'      mvarIsFlyer = ![Flyer]
'      mvarMobility = ![Mobility]
'      mvarDuty_Status = ![Duty_Status] & ""
'      mvarDataSource = "P"
'      GetImm_Personnel_Archive = True
'    End If
'  End With
'GIP_Exit:
'  On Error Resume Next
'  rst.Close
'  Set rst = Nothing
'  Exit Function
'
'GIP_Err:
  lLastError = Err.Number
  sLastErrorDesc = "CITA_DATA: GetImm_Personnel_Archive: " & Err.Description
  Err.Raise lLastError, "CITA_DATA: GetImm_Personnel_Archive: ", sLastErrorDesc

End Function
Public Function Recall_Imm_Personnel_Archive(sSSAN As String) As Boolean
'  On Error GoTo Err_Handler
'  If Len(sSSAN) = 9 Then
'    ClearPerson
'    mvarSSAN = sSSAN
'  Else
'    Exit Function
'  End If
'  If GetImm_Personnel_Archive() Then
'    UpdateImm_Personnel
'    'DeleteImm_Personnel_Archive
'    Recall_Imm_Personnel_Archive = True
'  End If
'Func_Exit:
'  On Error Resume Next
'  Exit Function
'
Err_Handler:
  lLastError = Err.Number
  sLastErrorDesc = "CITA_DATA: Recall_Imm_personnel_Archive: " & Err.Description
  Err.Raise lLastError, "CITA_DATA: Recall_Imm_personnel_Archive: ", sLastErrorDesc

End Function
Public Function GetImm_Personnel(rst As ADODB.Recordset) As Boolean

'Begin:SCR #20068 - Brian Mowbray 01/03/2002
'SCR Description: Performance Tunning - Open Immunizations fails the ORD6 second req in the field
'  Dim sSQL As String
'  Dim vItem As Variant
'  Dim rstP As ADODB.Recordset
'  On Error GoTo GIP_Err
'
'  If Len(rst![PT_SSAN] & vbNullString) = 0 Then
'    Exit Function
'  End If
'  sSQL = "SELECT * FROM Imm_Personnel WHERE SSAN='" & rst![PT_SSAN] & "'"
'  Set rstP = gobjSQL.OpenRecordset(sSQL)  ', adOpenForwardOnly, adLockReadOnly)
'  If rstP.BOF And rstP.EOF Then
'  Else
'    With rst
'      SplitNameRst rst![Name], rst
'      vItem = rst![DOB]
'      ![PT_DOB] = vItem
'      ![Pt_Unit] = rst![Unit] & vbNullString
'      ![PT_OFFICE_SYMBOL] = rst![OFFICE_SYMBOL] & vbNullString
'      ![PT_Service] = rst![Service] & vbNullString
'      ![PT_Status] = rst![Status] & vbNullString
'      ![PT_GENDER] = rst![Sex] & vbNullString
'      ![PT_RANK] = rst![Rank] & vbNullString
'      ![DataSource] = "P"
'    End With
    
    GetImm_Personnel = True

'  End If
'End: SCR #20068

GIP_Exit:
  On Error Resume Next
  Exit Function
  
GIP_Err:
  lLastError = Err.Number
  sLastErrorDesc = "CITA_DATA: Get_Imm_Personnel: " & Err.Description
  Err.Raise lLastError, "CITA_DATA: Get_Imm_Personnel: ", sLastErrorDesc
End Function

Public Function GetPerson(ByRef rstP As ADODB.Recordset) As Boolean
  Dim sFMP As String
  Dim sSSAN As String
  Dim bImm_Personnel_Found As Boolean
  Dim bImm_People_Found As Boolean
  Dim bImm_People_Dirty As Boolean
  On Error GoTo GetPerson_Err

  '>> SCR #38820;   Developer: Brian Mowbray 07/21/2003 10:45 AM
  If Len(rstP![SPON_SSAN_FMP] & vbNullString) > 0 Then
    sFMP = right(rstP![SPON_SSAN_FMP], 2)
    If sFMP = "20" Then
      If Len(rstP![PT_SSAN] & vbNullString) = 0 Then
        rstP![PT_SSAN] = Left(rstP![SPON_SSAN_FMP], 30)
        rstP.Update
      End If
    End If
  Else
    Exit Function
  End If
    
  If sFMP = "20" Then
    '>> SCR #38820;   Developer: Brian Mowbray 07/21/2003 10:57 AM
    If Len(rstP![PT_SSAN]) > 0 Then
      bImm_Personnel_Found = GetImm_Personnel(rstP)
    End If
  End If
  bImm_People_Found = GetImm_People(rstP)
  If Not bImm_People_Found Then
    If Not bImm_Personnel_Found Then
      'GetCHCS
    End If
  End If
  If Len(rstP![PT_Status] & vbNullString) = 0 Then
    If Val(sFMP) < 20 Then
      rstP![PT_Status] = "DEP"
    ElseIf Val(sFMP) > 29 And Val(sFMP) < 40 Then
      rstP!PT_Status = "DEP"
    Else
     ' rstP![pt_status] = oRules.
      'mvarStatus = sPT_Status = Publicly declared in Business Logic
    End If
  End If
  If Len(rstP![PT_Service] & vbNullString) = 0 Then
    'mvarService = sPT_Service = sPT_Status = Publicly declared in Business Logic
  End If
  If Not bImm_People_Found Then
'      UpdateImm_People rstPerson
  Else
    '! BM 8/21/2001
    '<< SCR #42805; 42809;   Developer: Brian Mowbray 10/02/2003 03:22 PM
    If gbITT Then CheckImm_People_Group rstP![SPON_SSAN_FMP] & vbNullString, rstP![PT_Status] & vbNullString, rstP![PT_Service] & vbNullString, rstP![PT_DOB]
    
  End If
  rstP.Update
  GetPerson = True
GetPerson_Exit:
  Exit Function
  
GetPerson_Err:
  lLastError = Err.Number
  sLastErrorDesc = "Imm_DATA: GetPerson: " & Err.Description
  Err.Raise lLastError, "Imm_DATA: GetPerson: ", sLastErrorDesc
 
End Function
Public Function GetImm_People(ByRef rst As ADODB.Recordset) As Boolean
  Dim sSQL As String
  Dim rstP As ADODB.Recordset
  Dim bDirty As Boolean
  
  On Error GoTo GetImm_People_Err
  If Len(rst![SPON_SSAN_FMP]) = 0 Then
    Exit Function
  End If
  sSQL = "SELECT * FROM Imm_People WHERE SPON_SSAN_FMP='" & rst![SPON_SSAN_FMP] & "'"
  Set rstP = gobjSQL.OpenRecordset(sSQL, adOpenForwardOnly, adLockOptimistic)
  If rstP.BOF And rstP.EOF Then
    UpdateImm_People rst, Append
  Else
    With rstP
      If rst![PT_SSAN] <> ![PT_SSAN] Then
        bDirty = True
      End If
      If Len(rst![PT_First_Name] & vbNullString) = 0 Then
        rst![PT_First_Name] = ![PT_First_Name] & ""
      Else
        If rst![PT_First_Name] <> ![PT_First_Name] Then
          bDirty = True
        End If
      End If
      If Len(rst![PT_MIDDLE_NAME] & vbNullString) = 0 Then
        rst![PT_MIDDLE_NAME] = ![PT_MIDDLE_NAME] & ""
      Else
        If rst![PT_MIDDLE_NAME] <> ![PT_MIDDLE_NAME] Then
          bDirty = True
        End If
      End If
      If Len(rst![PT_LAST_NAME] & vbNullString) = 0 Then
        rst![PT_LAST_NAME] = ![PT_LAST_NAME] & ""
      Else
        If rst![PT_LAST_NAME] <> ![PT_LAST_NAME] Then
          bDirty = True
        End If
      End If
      If Len(rst![PT_SUFFIX] & vbNullString) = 0 Then
        rst![PT_SUFFIX] = ![PT_SUFFIX] & ""
      Else
        If rst![PT_SUFFIX] <> ![PT_SUFFIX] Then
          bDirty = True
        End If
      End If
      If utNullToZero(rst![PT_DOB]) = 0 Then
        rst![PT_DOB] = ![PT_DOB]
      Else
        If rst![PT_DOB] <> ![PT_DOB] Then
          bDirty = True
        End If
      End If
      If Len(rst![Pt_Unit] & vbNullString) = 0 Then
        rst![Pt_Unit] = ![Pt_Unit] & ""
      Else
        If rst![Pt_Unit] <> ![Pt_Unit] & "" Then
          bDirty = True
        End If
      End If
      If Len(rst![PT_OFFICE_SYMBOL] & vbNullString) = 0 Then
        rst![PT_OFFICE_SYMBOL] = ![PT_OFFICE_SYMBOL] & ""
      Else
        If rst![PT_OFFICE_SYMBOL] <> ![PT_OFFICE_SYMBOL] Then
          bDirty = True
        End If
      End If
      If Len(rst![PT_Service] & vbNullString) = 0 Then
        rst![PT_Service] = ![PT_Service] & ""
      Else
        If rst![PT_Service] <> ![PT_Service] Then
          bDirty = True
        End If
      End If
      If Len(rst![PT_Status] & vbNullString) = 0 Then
        rst![PT_Status] = ![PT_Status] & ""
      Else
        If rst![PT_Status] <> ![PT_Status] Then
          bDirty = True
        End If
      End If
      If Len(rst![PT_GENDER] & vbNullString) = 0 Then
        rst![PT_GENDER] = ![PT_GENDER] & ""
      Else
        If rst![PT_GENDER] <> ![PT_GENDER] Then
          bDirty = True
        End If
      End If
      If Len(rst![PT_RANK] & vbNullString) = 0 Then
        rst![PT_RANK] = ![PT_RANK] & ""
      Else
        If rst![PT_RANK] <> ![PT_RANK] Then
          bDirty = True
        End If
      End If
      rst![DEERS_DupID] = ![DEERS_DupID] & ""
      rst![DEERS_Date] = ![DEERS_Date]
      If Len(rst![DataSource] & vbNullString) = 0 Then
        rst![DataSource] = ![DataSource] & ""
      End If
      rst![LastEdited] = ![LastEdited]
      rst![ExemptionExpirationDate] = ![ExemptionExpirationDate]
      rst![ExemptionCode] = CInt(utNullToZero(![ExemptionCode]))
      rst![Comments] = ![Comments] & vbNullString
      rst![Dirty] = bDirty
      rst.Update
      GetImm_People = True
    End With
  End If
  

GetImm_People_Exit:
  On Error Resume Next
  rstP.Close
  Set rstP = Nothing
  Exit Function
  
GetImm_People_Err:
  lLastError = Err.Number
  sLastErrorDesc = "CITA_DATA: GetImm_People: " & Err.Description
  Err.Raise lLastError, "CITA_DATA: GetImm_People: ", sLastErrorDesc
End Function




Public Function IsIn_Imm_People(Optional sspon_SSAN_FMP As String = vbNullString, _
                                Optional sSSAN As String = vbNullString, _
                                Optional AutoAdd As Boolean = True) As Boolean
  Dim rst As ADODB.Recordset
  Dim rstP As ADODB.Recordset
  Dim sSQL As String
  On Error GoTo IIP_Err
   
  If Len(sspon_SSAN_FMP) = 0 Then
    If Len(sSSAN) = 0 Then
      GoTo IIP_Exit
    Else
      sSQL = "SELECT Imm_People.PT_SSAN FROM Imm_People WHERE PT_SSAN='" & sSSAN & "'"
    End If
  Else
    sSQL = "SELECT Imm_People.SPON_SSAN_FMP FROM Imm_People WHERE SPON_SSAN_FMP='" & sspon_SSAN_FMP & "'"
  End If
  Set rst = gobjSQL.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  If rst.BOF And rst.EOF Then
    IsIn_Imm_People = False
    If AutoAdd Then
      Set rstP = New ADODB.Recordset
      CreateRst Imm_People, rstP
      rstP.AddNew
      If Len(sspon_SSAN_FMP) = 0 Then
        sLastErrorDesc = "CITA_DATA: IsIn_Imm_People: No SPON_SSAN_FMP entered! Can't add to Imm_People."
        lLastError = 100 + vbObjectError
        Err.Raise lLastError, "CITA_DATA: IsIn_Imm_People: ", sLastErrorDesc
        GoTo IIP_Exit
      Else
        '>> SCR #38820;   Developer: Brian Mowbray 07/21/2003 10:46 AM
        If Len(sspon_SSAN_FMP) > 0 Then
          rstP![SPON_SSAN_FMP] = sspon_SSAN_FMP & vbNullString
          rstP![PT_SSAN] = sSSAN & vbNullString
          rstP.Update
          If GetPerson(rstP) Then
          Else
            sLastErrorDesc = "CITA_DATA: IsIn_Imm_People: Unable to Add the person to Imm_People"
            lLastError = 100 + vbObjectError
            Err.Raise lLastError, "CITA_DATA: IsIn_Imm_People: ", sLastErrorDesc
            GoTo IIP_Exit
          End If
        Else
          sLastErrorDesc = "CITA_DATA: IsIn_Imm_People: Please enter a valid SPON_SSAN_FMP! If has to be at least 11 characters."
          lLastError = 100 + vbObjectError
          Err.Raise lLastError, "CITA_DATA: IsIn_Imm_People: ", sLastErrorDesc
          GoTo IIP_Exit
        End If
      End If
    End If
  End If
  IsIn_Imm_People = True
  
IIP_Exit:
  On Error Resume Next
  rst.Close
  rstP.Close
  Set rst = Nothing
  Set rstP = Nothing
  Exit Function


IIP_Err:
  lLastError = Err.Number
  sLastErrorDesc = "CITA_DATA: IsIn_Imm_People: " & Err.Description
  Err.Raise lLastError, "CITA_DATA: IsIn_Imm_People: ", sLastErrorDesc
End Function

 

