VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GenName_Rec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private msGenericName As String
Private msGenericID As String
Private msMinAge As String
Private msMaxAge As String
Private msType As String
Private msCycle As String
Private msRange As String
Private msRelated_Gen_Names As String
Private msDefunct As String
Private msPhysReqONLY As String
Private msCombination As String

Public Function ModGenericData(OldGenName As String) As Boolean

  ' This Routine modifies an existing Record in the Generic Name DataBase

 On Local Error GoTo haveerror
 
  Dim records As Long
  Dim xSet As String
  Dim sql As String
  
 With GenericVacForm
 
  xSet = "GenericName = '" & msGenericName & "', " & _
         "MinAge = '" & msMinAge & "', " & _
         "MaxAge = '" & msMaxAge & "', " & _
         "Type = " & msType & ", " & _
         "Cycle = '" & msCycle & "', " & _
         "Range = '" & msRange & "', " & _
         "Related_Gen_Names = '" & IMM_Tables.Get_Related_Generic_Names(OldGenName) & "', " & _
         "Combination = " & msCombination & ", " & _
         "Defunct = " & msDefunct & ", " & _
         "PHYSREQONLY = " & msPhysReqONLY & " "
         
 End With
 
    If Not gobjShared.UseLocalDB Then
     xSet = xSet & ", " & _
     "UpdatedOn = " & ConvertDateToIOType(Now) & ", " & _
     "UpdatedBy = " & gobjLogon.UserNCID & " "
    End If
         
  '...Update Generic Name Table
  sql = "Update" & Name_Of_Generic_Vac_Names_Table & "SET " & xSet & "WHERE GenericName = '" & OldGenName & "'"
  gobjSQL.ExecuteSQL sql
  
  '...Modify all entries in the Vaccine Table
  '   and VacHistKey Table
  If Trim$(msGenericName) <> OldGenName Then
    On Local Error Resume Next
    '...Update Generic Name...Remove this
    '   when geneic name is removed from table
    sql = "Update TBLVACCINES SET " & _
          "GenericName = '" & Trim$(msGenericName) & "' " & _
          "WHERE GenericName = '" & OldGenName & "'"
    gobjSQL.ExecuteSQL sql
    
    If Err.Number Then
      '...No records in TBLVaccines was using this Generic Name
      Err.Clear
      GoTo noVaccine
    End If
    
    On Local Error GoTo haveerror
    '...Update HistKey
    sql = "Update TBLVACCINES SET " & _
          "HistKey = '" & Trim$(msGenericName) & "', " & _
          "UpdatedOn = " & ConvertDateToIOType(Now) & ", " & _
          "UpdatedBy = " & gobjLogon.UserNCID & " " & _
          "WHERE HistKey = '" & OldGenName & "'"
    gobjSQL.ExecuteSQL sql
noVaccine:
    On Local Error GoTo haveerror
    '...Update HistKey in HistVacKey Table
'    sql = "Update HistVacKey SET " & _
'          "HistKey = '" & Trim$(msGenericName) & "' " & _
'          "WHERE HistKey = '" & OldGenName & "'"
    sql = "Update HistVacKey SET " & _
          "HistKey = '" & PrefixInitVacidHF & "0" & SufixInitVacidHF & WithinFieldSepChar & msGenericName & WithinFieldSepChar & "' " & _
          "WHERE HistKey = '" & OldGenName & "'"

    gobjSQL.ExecuteSQL sql
    
  End If
  
  ModGenericData = True
  
exitRoutine:

 Exit Function
 
haveerror:
 ErrMsgBox "GenName_Rec:  ModGenericData: Modify a Generic Name:", Error$, Err
 Resume exitRoutine
 Resume
End Function

Public Function AddNewGenericData() As Boolean
 On Local Error GoTo haveerror
 Dim records As Long
 Dim sql As String
 Dim dataid As Long
 Dim rs As ADODB.Recordset
 
    If gobjShared.UseMySQLDB Then
        gobjSQL.ExecuteSQL "UPDATE seq_imm_ref SET id=LAST_INSERT_ID(id+1)"
        Set rs = gobjSQL.OpenRecordset("SELECT LAST_INSERT_ID()")
        dataid = rs(0).Value
        If dataid <> 0 Then
           msGenericID = CStr(dataid)
           sql = "INSERT INTO " & Name_Of_Generic_Vac_Names_Table & " (GenNameID, GenericName, MinAge, MaxAge, Type, Cycle, Range, Related_Gen_Names, Combination, Defunct, PHYSREQONLY, CreatedBy, CreatedOn) " & _
                 "VALUES (" & msGenericID & _
                 ", '" & msGenericName & _
                 "', '" & msMinAge & _
                 "', '" & msMaxAge & _
                 "', " & msType & _
                 ", '" & msCycle & _
                 "', '" & msRange & _
                 "', '" & msRelated_Gen_Names & _
                 "', " & msCombination & _
                 ", " & msDefunct & _
                 ", " & msPhysReqONLY & _
                 ", " & gobjLogon.UserNCID & ", " & ConvertDateToIOType(Now) & " )"
        End If
    ElseIf gobjShared.UseSQLServer Then
        msGenericID = ""
        sql = "INSERT INTO " & Name_Of_Generic_Vac_Names_Table & " (GenericName, MinAge, MaxAge, Type, Cycle, Range, Related_Gen_Names, Combination, Defunct, PHYSREQONLY, CreatedBy, CreatedOn) " & _
              "VALUES ('" & msGenericName & _
              "', '" & msMinAge & _
              "', '" & msMaxAge & _
              "', " & msType & _
              ", '" & msCycle & _
              "', '" & msRange & _
              "', '" & msRelated_Gen_Names & _
              "', " & msCombination & _
              ", " & msDefunct & _
              ", " & msPhysReqONLY & _
              ", " & gobjLogon.UserNCID & ", " & ConvertDateToIOType(Now) & " )"
   ElseIf gobjShared.UseLocalDB Then
        Set rs = gobjSQL.OpenRecordset("Select IMM_REF.Nextval from DUAL")
        dataid = rs(0).Value
        If dataid <> 0 Then
           msGenericID = CStr(dataid)
           sql = "INSERT INTO " & Name_Of_Generic_Vac_Names_Table & " (GenNameID, GenericName, MinAge, MaxAge, Type, Cycle, Range, Related_Gen_Names, Combination, Defunct, PHYSREQONLY, CreatedBy, CreatedOn) " & _
                 "VALUES (" & msGenericID & _
                 ", '" & msGenericName & _
                 "', '" & msMinAge & _
                 "', '" & msMaxAge & _
                 "', " & msType & _
                 ", '" & msCycle & _
                 "', '" & msRange & _
                 "', '" & msRelated_Gen_Names & _
                 "', " & msCombination & _
                 ", " & msDefunct & _
                 ", " & msPhysReqONLY & _
                 ", " & gobjLogon.UserNCID & ", " & ConvertDateToIOType(Now) & " )"
        End If
    
   End If
           
    If sql <> "" Then
        gobjSQL.ExecuteSQL sql
'        If DBType = IOTYPE.Access Then
'            '...Get ImmRelNCID from Generic Names Table
'            sql = "Select GenNameId from " & Name_Of_Generic_Vac_Names_Table & " Where GenericName = '" & msGenericName & "'"
'            Set rs = gobjSQL.OpenRecordset(sql)
'            If Not rs.EOF Then dataid = rs(0).Value
'            rs.Close
'            Set rs = Nothing
'        End If
        If dataid <> 0 Then
            '...Add HistKey to the HistVacKey file
            '<< Begin:SCR #36395;   Developer: Brian Mowbray 05/21/2003 05:05 PM
            sql = "Insert INTO HistVacKey " & _
             "(HISTKEY, PriorNCID, IMMRELNCID) VALUES ('" & _
             PrefixInitVacidHF & "0" & SufixInitVacidHF & WithinFieldSepChar & msGenericName & WithinFieldSepChar & "', 0, " & CStr(dataid) & " )"
            '>> End: SCR #36395;
            
            'sql = "Insert INTO HistVacKey (HistKey, ImmRelNcid) Values ('" & msGenericName & "', " & CStr(dataid) & ")"
            gobjSQL.ExecuteSQL sql
        End If
      
      AddNewGenericData = True
    Else
        MsgBxARMd "Could not Acquire a Valid Sequence Number", vbExclamation + vbOKOnly, "File: " & Name_Of_Generic_Vac_Names_Table
    End If
  
exitRoutine:

    Exit Function
 
haveerror:
    ErrMsgBox "GenName_Rec  AddNewGenericdata:", Error$, Err
    Resume exitRoutine
    Debug.Print Err.Number, Err.Description
    Resume
End Function

Public Property Get GenericName() As String
 GenericName = msGenericName
End Property

Public Property Let GenericName(ByVal vNewValue As String)
 msGenericName = RemoveInvalidChar(vNewValue)
End Property

Public Property Get GenericID() As String
 GenericID = msGenericID
End Property

Public Property Let GenericID(ByVal vNewValue As String)
  msGenericID = RemoveInvalidChar(vNewValue)
End Property

Public Property Get MinAge() As String
   MinAge = msMinAge
End Property

Public Property Let MinAge(ByVal vNewValue As String)
  msMinAge = RemoveInvalidChar(vNewValue)
End Property

Public Property Get MaxAge() As String
 MaxAge = msMaxAge
End Property

Public Property Let MaxAge(ByVal vNewValue As String)
  msMaxAge = RemoveInvalidChar(vNewValue)
End Property

Public Property Get GenType() As String
 GenType = msType
End Property

Public Property Let GenType(ByVal vNewValue As String)
 msType = RemoveInvalidChar(vNewValue)
End Property

Public Property Get Cycle() As String
 Cycle = msCycle
End Property

Public Property Let Cycle(ByVal vNewValue As String)
 msCycle = RemoveInvalidChar(vNewValue)
End Property

Public Property Get Range() As String
  Range = msRange
End Property

Public Property Let Range(ByVal vNewValue As String)
  msRange = RemoveInvalidChar(vNewValue)
End Property

Public Property Get Related_Gen_Names() As String
 Related_Gen_Names = msRelated_Gen_Names
End Property

Public Property Let Related_Gen_Names(ByVal vNewValue As String)
 msRelated_Gen_Names = RemoveInvalidChar(vNewValue)
End Property

Public Property Get Defunct() As String
 Defunct = msDefunct
End Property

Public Property Let Defunct(ByVal vNewValue As String)
 msDefunct = RemoveInvalidChar(vNewValue)
End Property

Public Property Get PhysReqONLY() As String
 PhysReqONLY = msPhysReqONLY
End Property

Public Property Let PhysReqONLY(ByVal vNewValue As String)
 msPhysReqONLY = RemoveInvalidChar(vNewValue)
End Property

Public Property Get Combination() As String
 Combination = msCombination
End Property

Public Property Let Combination(ByVal vNewValue As String)
 msCombination = RemoveInvalidChar(vNewValue)
End Property

Public Function Record() As String

    ' Build Vaccine Record

    Dim xfield As Integer
    Dim field As String
    Dim sepchar As String
    Dim DelRecord As String
    
    
    For xfield = 0 To GenVacDictFields.NumFlds - 1
    
        Select Case xfield
        
            Case GenVacDictFields.GenericName
               field = msGenericName
            Case GenVacDictFields.GenericID
               field = msGenericID
            Case GenVacDictFields.MinAge
               field = MinAge
            Case GenVacDictFields.MaxAge
               field = msMaxAge
            Case GenVacDictFields.xtype
               field = msType
            Case GenVacDictFields.Cycle
               field = msCycle
            Case GenVacDictFields.Range
               field = msRange
            Case GenVacDictFields.Discontinued
               field = msDefunct
            Case GenVacDictFields.RecmdDTROnly
               field = msPhysReqONLY
            Case GenVacDictFields.Combination
               field = msCombination
            Case GenVacDictFields.IgnoreAge
               field = "N"

        End Select
        
        DelRecord = DelRecord & sepchar & field
        sepchar = FieldSepChar
        
    Next xfield
    
    Record = DelRecord

End Function

Private Sub Class_Initialize()
    '--- COMMON TO MULTIPLE SCR-ARMs: MSolano 9/13/01 SCR-ARM-16110, 16130
    '--- Init routine should be called from first top level creatable coclass
    '--- It can be called multiple times, but will actually run only once.
    '--- Instrument all creatable classes in case dll's multiple entry
    '--- points are actually used
    Call ARM_InitAllTransactionsForImmunizations
End Sub
