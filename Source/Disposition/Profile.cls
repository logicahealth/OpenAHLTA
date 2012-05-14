VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Profile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public StartDate As Date
Public EndDate As Date
Public Comment As String
Public DESCRIPTION As String

Public DataID As String
    
Public SnoID As String
Public Prefix As String

Private m_szPersist As String
Implements IPersist

'Use Cases:
'Load First Time, Patient has no Disposioion Record
'Load 2nd Time, Patient Has Disposition Record

'No Data Will Be Set, Module is not Dirty
'Opens Dialog and Sets a Value, Module is Dirty
'Opens Dialog and Does Not Set Value, Module is NOT dirty
'
Public Function Dirty() As Boolean
    If IPersist_Persist() <> m_szPersist Then
        Dirty = True
    End If
    
End Function
Public Function Load() As Boolean
    Dim szSQL As String
    Dim szValues As String
    Dim rs As ADODB.Recordset
    Dim objDAS As ICHCSII_DAS
    Dim F As ADODB.Field
    Dim szFlds As String
    Dim szVals As String
    
    On Error GoTo errXIT
    
    If Dirty = False Then GoTo XIT
    
'SCR 65426
'    If gDisplayMode = InpatientDischarge Then GoTo XIT
    
    Set objDAS = modControls.GetCHCSII_DAS
    
    szSQL = "SELECT * FROM Readiness "
    szSQL = szSQL & "WHERE ENCOUNTER_NUMBER = '%2' "
    szSQL = szSQL & "AND Last_Readiness_Dt IN "
    szSQL = szSQL & "   (SELECT Max(Last_Readiness_Dt) "
    szSQL = szSQL & "   FROM Readiness "
    szSQL = szSQL & "   WHERE Patient_Unit_Number = '%1' )"

    szSQL = Replace(szSQL, "%1", gobjPatient.UnitNumber)
    szSQL = Replace(szSQL, "%2", gobjEncParent.CurrentEncounter.EncounterID)
    
    'Copy Last Row Forward
    Set rs = objDAS.OpenRecordset(szSQL)
    
    If rs.RecordCount Then
        
        'Check for NULL values before loading profile data SCR#50419
        If Not IsNull(rs.Fields("TEMP_PROFILE_START_DT").Value) Then
            Me.StartDate = rs.Fields("TEMP_PROFILE_START_DT").Value
        End If
        
        If Not IsNull(rs.Fields("TEMP_PROFILE_END_DT").Value) Then
            Me.EndDate = rs.Fields("TEMP_PROFILE_END_DT").Value
        End If
        
        If Not IsNull(rs.Fields("TEMP_PROFILE_DIGNOS_CD").Value) Then
            Me.SnoID = rs.Fields("TEMP_PROFILE_DIGNOS_CD").Value
        End If
        
        If Not IsNull(rs.Fields("TEMP_PROFILE_DIGNOS_DSC").Value) Then
            Me.DESCRIPTION = rs.Fields("TEMP_PROFILE_DIGNOS_DSC").Value
        End If
        
        If Not IsNull(rs.Fields("TEMP_PROFILE_LIMITATION").Value) Then
            Me.Comment = rs.Fields("TEMP_PROFILE_LIMITATION").Value
        End If
        
    End If

XIT:
    m_szPersist = IPersist_Persist()
    Load = True
    
    Exit Function
    
errXIT:
    ReportError "Profile", Err, "Load", vbExclamation


End Function
'Handle varchar and nulls in CDR mode SCR#49310
Public Function Save() As Boolean

    Dim szSQL As String
    Dim szValues As String
    Dim rs As ADODB.Recordset
    Dim objDAS As ICHCSII_DAS
    Dim F As ADODB.Field
    Dim szFlds As String
    Dim szVals As String
    Dim bAddField As Boolean
    Dim sMsg As String
    Dim oConn As New Conn
    
    'R. Foster SCR #48177  Developer: Robert Foster 01/20/2004 06:14 PM
    'R. Foster Note: Function Changed: CodedValues.cls - EMCodeNCID (Let); Profile.cls - Save;  modControls.bas - GetEMTable;      Impact of Change:  Modifiy code to account for ITT mode
    'ITT CONVERGENCE MODIFICATION
    Dim objShared As CWShared
    On Error GoTo errXIT
    'R. Foster SCR #48177  Developer: Robert Foster 01/20/2004 06:14 PM
    'R. Foster Note: Function Changed: CodedValues.cls - EMCodeNCID (Let); Profile.cls - Save;  modControls.bas - GetEMTable;      Impact of Change:  Modifiy code to account for ITT mode
    'ITT CONVERGENCE MODIFICATION
    sMsg = "Create Shared"
    Set objShared = New CWShared
    
    If Dirty = False Then GoTo XIT
    
    sMsg = "Get DAS"
    Set objDAS = modControls.GetCHCSII_DAS
    
    sMsg = "Create SQL"
    szSQL = "SELECT * FROM Readiness "
    szSQL = szSQL & "WHERE Patient_Unit_Number = '%1' "
    szSQL = szSQL & "AND Last_Readiness_Dt IN "
    szSQL = szSQL & "   (SELECT Max(Last_Readiness_Dt) "
    szSQL = szSQL & "   FROM Readiness "
    szSQL = szSQL & "   WHERE Patient_Unit_Number = '%1' )"

    szSQL = Replace(szSQL, "%1", gobjPatient.UnitNumber)

    'Copy Last Row Forward
    sMsg = "Get Recordset"
    Set rs = objDAS.OpenRecordset(szSQL)

    If rs.RecordCount Then
    
        sMsg = "Adding Fields"
        szFlds = vbNullString
        szVals = vbNullString
    
        For Each F In rs.Fields
            bAddField = True
            Debug.Print F.Name & " - " & F.Value
            sMsg = F.Name & " : " & F.Value & " : " & F.Type
            Select Case F.Name
            Case "PATIENT_UNIT_NUMBER"
                szVals = szVals & "'" & gobjPatient.UnitNumber & "', "
            Case "ENCOUNTER_NUMBER"
                szVals = szVals & "'" & gobjEncParent.CurrentEncounter.EncounterID & "', "
            Case "TEMP_PROFILE"
                szVals = szVals & "'Y', "
            Case "TEMP_PROFILE_START_DT"
                szVals = szVals & oConn.SQLDate(Me.StartDate) & ", "
            Case "TEMP_PROFILE_END_DT"
                szVals = szVals & oConn.SQLDate(Me.EndDate) & ", "
            Case "TEMP_PROFILE_DIGNOS_CD"
                szVals = szVals & "'" & Me.SnoID & "', "
            Case "TEMP_PROFILE_DIGNOS_DSC"
                 szVals = szVals & oConn.SQLQuote(Me.DESCRIPTION) & ", "
            Case "TEMP_PROFILE_LIMITATION"
                szVals = szVals & oConn.SQLQuote(Me.Comment) & ", "
            Case "LAST_READINESS_DT"
                szVals = szVals & oConn.SQLDate(Now) & ", "
            Case Else
                If Not IsNull(F.Value) Then
                    If F.Type = adDBTimeStamp Or F.Type = adDate Then 'SCR#48704
                        If F.Value <> cwdNULL_DATE Then
                            szVals = szVals & oConn.SQLDate(F.Value) & ", "
                        Else
                            bAddField = False
                        End If
                    Else
                        If F.Type = adLongVarChar Or F.Type = adChar Or F.Type = adVarChar Or F.Type = adLongVarWChar Then  'SCR#48704
                            If Len(Trim$(F.Value)) Then
                                szVals = szVals & "'" & F.Value & "', "
                            Else
                                bAddField = False
                            End If
                        Else
                            szVals = szVals & F.Value & ", "
                        End If
                    End If
                Else
                    bAddField = False
                End If
            End Select
            
            If bAddField Then
                szFlds = szFlds & F.Name & ", " & vbCrLf
            End If
        Next
        
        szFlds = Left(szFlds, Len(szFlds) - 4)
        szVals = Left(szVals, Len(szVals) - 2)
        
        szSQL = "INSERT INTO READINESS (" & szFlds
        szSQL = szSQL & ") Values (" & szVals & ")"
    Else
        'Make Virgin Row
        szSQL = "INSERT INTO READINESS " & vbCrLf
        szSQL = szSQL & "( LAST_READINESS_DT, PATIENT_UNIT_NUMBER, " & vbCrLf
        szSQL = szSQL & " ENCOUNTER_NUMBER, TEMP_PROFILE, " & vbCrLf
        szSQL = szSQL & " TEMP_PROFILE_START_DT, TEMP_PROFILE_END_DT, " & vbCrLf
        szSQL = szSQL & " TEMP_PROFILE_DIGNOS_CD, TEMP_PROFILE_DIGNOS_DSC, " & vbCrLf
        szSQL = szSQL & " TEMP_PROFILE_LIMITATION ) VALUES " & vbCrLf
        szSQL = szSQL & "( %LAST_READINESS_DT, '%PATIENT_UNIT_NUMBER', '%ENCOUNTER_NUMBER', " & vbCrLf
        szSQL = szSQL & "'%TEMP_PROFILE', %TEMP_PROFILE_START_DT, %TEMP_PROFILE_END_DT," & vbCrLf
        szSQL = szSQL & "'%TEMP_PROFILE_DIGNOS_CD', %TEMP_PROFILE_DIGNOS_DSC, %TEMP_PROFILE_LIMITATION )" & vbCrLf
        
        szSQL = Replace(szSQL, "%LAST_READINESS_DT", oConn.SQLDate(Now))

        szSQL = Replace(szSQL, "%PATIENT_UNIT_NUMBER", gobjPatient.UnitNumber)
        
        szSQL = Replace(szSQL, "%ENCOUNTER_NUMBER", gobjEncParent.CurrentEncounter.EncounterID)
        szSQL = Replace(szSQL, "%TEMP_PROFILE_START_DT", oConn.SQLDate(Me.StartDate))
        szSQL = Replace(szSQL, "%TEMP_PROFILE_END_DT", oConn.SQLDate(Me.EndDate))
        szSQL = Replace(szSQL, "%TEMP_PROFILE_DIGNOS_CD", Me.SnoID)
        szSQL = Replace(szSQL, "%TEMP_PROFILE_DIGNOS_DSC", oConn.SQLQuote(Me.DESCRIPTION))
        szSQL = Replace(szSQL, "%TEMP_PROFILE_LIMITATION", oConn.SQLQuote(Me.Comment)) 'SCR-42667
        szSQL = Replace(szSQL, "%TEMP_PROFILE", "Y")
        
    End If
        
    sMsg = "Execute SQL :" & szSQL
    objDAS.ExecuteSQL szSQL
    
XIT:
    
    m_szPersist = IPersist_Persist()
    Save = True
    
    Exit Function
    
errXIT:
    gobjShared.ShowVBError Err.Number, Err.DESCRIPTION & " - " & sMsg, "Profile.Save", "Disposition", vbExclamation
    
End Function

Public Function RTFNote() As String
    Dim sz As String
    
    sz = vbNullString
    
    If Me.StartDate <> 0 Then 'if date = 0 then Profile button was not clicked SCR-31776
        sz = "Profile: %Desc from %Start to %End; Comment: %Comment"
        
        sz = Replace(sz, "%Desc", Me.DESCRIPTION)
        sz = Replace(sz, "%Start", Format(Me.StartDate, "dd mmm yyyy"))
        sz = Replace(sz, "%End", Format(Me.EndDate, "dd mmm yyyy"))
        sz = Replace(sz, "%Comment", Me.Comment)
    End If
    RTFNote = sz
    
End Function

Private Function IPersist_ChildFactory() As Object

End Function

Private Function IPersist_Persist() As String
    IPersist_Persist = Persist(Me)
End Function

Private Function IPersist_Restore(ByVal Persist As String) As Boolean
    IPersist_Restore = Restore(Me, Persist)
    m_szPersist = IPersist_Persist()
End Function
