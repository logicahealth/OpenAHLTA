VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colTrNSECanal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements clsISection

Private mcolCanalnfo As Collection
Private mobjDentalNotes As CHCSII_Dental_DXTX.clsNotes
Private mstrRCTTeeth As String

Private mlngDentalEncounterId As Long
Private mlngPatUnitId As Long
Private mlngDirtyFlag As DirtyFlagEnum

Private Const const_ModuleName = "colTrNSECanal"

'**************************************************************************************
'Sub: Delete
'Author:KDunne
'Date:5/14/2002
'Purpose: Deletes the selected canal object
'**************************************************************************************
Public Sub Delete(ByVal lngIndex As Long)
    Dim objCanalInfo As clsTrNSECanal
    Dim objSection As clsISection
    
    On Error GoTo ErrHandler
    
    'Make sure our Index is valid
    If lngIndex > mcolCanalnfo.Count Then GoTo FuncExit
    
    Set objCanalInfo = mcolCanalnfo.Item(lngIndex)
    
    objCanalInfo.DirtyFlag = dfdelete
    objCanalInfo.SaveSection
    
    mcolCanalnfo.Remove (lngIndex)
FuncExit:
    Set objCanalInfo = Nothing
    Exit Sub

ErrHandler:
    goCWShared.ShowVBError Err.Number, Err.Description, const_ModuleName & ".Delete()", "", vbCritical
    GoTo FuncExit
End Sub
Public Property Get Canals() As Collection
    Set Canals = mcolCanalnfo
End Property
Public Property Get Canal_Info(ByVal lngIndex As Long) As clsTrNSECanal
    Dim objCanalInfo As clsTrNSECanal
    
    If lngIndex > mcolCanalnfo.Count Then
        Set objCanalInfo = New clsTrNSECanal
        objCanalInfo.DentalEncounterID = mlngDentalEncounterId
        
        mcolCanalnfo.Add objCanalInfo
    End If
    
    Set Canal_Info = mcolCanalnfo(lngIndex)
    
    Set objCanalInfo = Nothing
End Property
'********************************************************
'Function:
'Author: KDunne
'Date:12/6/2001
'Purpose: Returns a canal info object based on the id
'********************************************************
Public Function Load_Canal_Info_By_ID(ByVal lngId As Long) As clsTrNSECanal
    
    Dim objCanalInfo As clsTrNSECanal

    For Each objCanalInfo In mcolCanalnfo
        If objCanalInfo.Canal_Info_Id = lngId Then
            Set Load_Canal_Info_By_ID = objCanalInfo
            Exit For
        End If
    Next
    
    Set objCanalInfo = Nothing
End Function
Private Property Get clsISection_DentalEncounterID() As Long
    clsISection_DentalEncounterID = mlngDentalEncounterId
End Property

Private Function clsISection_GetNote() As String
    Dim strNote As String
    Dim strCanalInfo As String
    Dim objCanalInfo As clsTrNSECanal
    
    strNote = Empty
    strCanalInfo = Empty
    
    If mcolCanalnfo Is Nothing = False Then
        For Each objCanalInfo In mcolCanalnfo
            strCanalInfo = strCanalInfo & objCanalInfo.GetNote & vbCrLf
        Next
        
        If strCanalInfo <> "" Then
            strNote = strNote & "Canal Specific Information" & _
                vbCrLf & strCanalInfo
        End If
    End If
    
FuncExit:
    Set objCanalInfo = Nothing
    clsISection_GetNote = strNote
    
    Exit Function
ErrHandler:
    goCWShared.ShowVBError Err.Number, Err.Description, const_ModuleName & ".GetNote()", "", vbCritical
    strNote = Empty
    GoTo FuncExit

End Function

Private Function clsISection_GetRTFNote() As String
    Dim strNote As String
    Dim strCanalInfo As String
    Dim objCanalInfo As clsTrNSECanal
    
    strNote = Empty
    strCanalInfo = Empty
    
    If mcolCanalnfo Is Nothing = False Then
        For Each objCanalInfo In mcolCanalnfo
            strCanalInfo = strCanalInfo & objCanalInfo.GetNote & vbCrLf
        Next
        
        If strCanalInfo <> "" Then
            strNote = strNote & crtfStartControlName & "Canal Specific Information: " & _
                crtfEndControlName & strCanalInfo
        End If
    End If
    
FuncExit:
    Set objCanalInfo = Nothing
    clsISection_GetRTFNote = strNote
    
    Exit Function
ErrHandler:
    goCWShared.ShowVBError Err.Number, Err.Description, const_ModuleName & ".GetRTFNote()", "", vbCritical
    strNote = Empty
    GoTo FuncExit

End Function

Private Function clsISection_LoadSection(ByVal PatientID As Long, ByVal DentalEncounterID As Long) As Boolean
    Dim blnReturn As Boolean
    Dim objRS As ADODB.Recordset
    Dim objToothInfo As clsTrNSETooth
    Dim objCanalInfo As clsTrNSECanal
    Dim blnHasData As Boolean
    
    On Error GoTo ErrHandler
    
    mlngPatUnitId = PatientID
    mlngDentalEncounterId = DentalEncounterID
    
    Set mcolCanalnfo = New Collection
    
    Set objRS = DB_LoadNSECanalInfo(mlngDentalEncounterId)
    
    With objRS
        If .EOF Then
            mlngDirtyFlag = dfAdd
        Else
            mlngDirtyFlag = dfUpdate
            
            Do While Not .EOF
                Set objCanalInfo = New clsTrNSECanal
                
                objCanalInfo.Canal_Info_Id = Val(.Fields("ET_CANAL_INFO_ID") & "")
                objCanalInfo.Canal = Val(.Fields("TOOTH_CANAL") & "")
                objCanalInfo.DentalEncounterID = mlngDentalEncounterId
                objCanalInfo.Tooth_Number = .Fields("TOOTH_NUMBER") & ""
                objCanalInfo.Working_Length = Val(.Fields("WORKING_LENGTH") & "")
                
                mcolCanalnfo.Add objCanalInfo
                
                .MoveNext
            Loop
            
        End If '.EOF
    
    End With 'objRs

    If mlngDirtyFlag = dfUpdate Then
        'We received data on the first, partial query. Now get the rest of the data.
        Set objRS = DB_LoadNSECanalInfo2(mlngDentalEncounterId)
        With objRS
            
            Do While Not .EOF
                Set objCanalInfo = Load_Canal_Info_By_ID(Val(.Fields("ET_CANAL_INFO_ID") & ""))
                objCanalInfo.Comments = .Fields("COMMENTS") & ""
                objCanalInfo.File_System = .Fields("FILE_SYSTEM") & ""
                objCanalInfo.Instrumentation_Techniques = .Fields("INTRUMENT_TECH") & ""
                objCanalInfo.Master_Apical_File = Val(.Fields("MAF") & "")
                objCanalInfo.Obturation_Materials = .Fields("OBTURATION_MAT") & ""
                objCanalInfo.Obturation_Technique = .Fields("OBTURATION_TECH") & ""
                objCanalInfo.References = .Fields("REFERENCES") & ""
                .MoveNext
            Loop
            
        End With 'objRS
    End If
'    mstrRCTTeeth = GetTeethWithRCTTreatment

FuncExit:
    clsISection_LoadSection = blnReturn
    Set objRS = Nothing
    Exit Function
ErrHandler:
   blnReturn = False
   goCWShared.ShowVBError Err.Number, Err.Description, const_ModuleName & ".clsISection_LoadSection()", "", vbCritical

   GoTo FuncExit
End Function

Private Property Get clsISection_PatientID() As Long
    clsISection_PatientID = mlngPatUnitId
End Property

Private Function clsISection_SaveSection() As Boolean
    Dim blnReturn As Boolean
    Dim objCanalInfo As clsTrNSECanal
    
    On Error GoTo ErrHandler:
    
    If mcolCanalnfo Is Nothing = False Then
        For Each objCanalInfo In mcolCanalnfo
            blnReturn = objCanalInfo.SaveSection
        Next
    End If

FuncExit:
    Set objCanalInfo = Nothing
    
    Exit Function
    
ErrHandler:
    blnReturn = False
    goCWShared.ShowVBError Err.Number, Err.Description, const_ModuleName & ".clsISection_SaveSection()", "", vbCritical
    GoTo FuncExit
End Function

'****************************************************
'Sub: UpdateDirtyFlag
'Author: KDunne
'Date: 11/5/2001
'Purpose: Updates module level dirty flag field
'****************************************************
Private Sub UpdateDirtyFlag()

    If mlngDirtyFlag = dfNoAction Then
        mlngDirtyFlag = dfUpdate
    End If
    
End Sub

Public Property Get DentalNotes() As CHCSII_Dental_DXTX.clsNotes
    Set DentalNotes = mobjDentalNotes
End Property
Public Property Set DentalNotes(ByVal objNotes As CHCSII_Dental_DXTX.clsNotes)
    Set mobjDentalNotes = objNotes
End Property


'****************************************************
'Sub: DB_LoadNSECanalInfo
'Author: KDunne
'Date: 12/04/2001
'Purpose: Retrieves all of the Canal Specific information for a dental encounter.
'****************************************************
Private Function DB_LoadNSECanalInfo(ByVal lngDentalEncounterID As Long) As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo ErrHandler:
    strSql = "Select ET_CANAL_INFO_ID,TOOTH_NUMBER, TOOTH_CANAL, WORKING_LENGTH " & _
        " From TEMPL_TREAT_NSE_CANAL_INFO WHERE DENTAL_ENCOUNTER_ID = " & lngDentalEncounterID

            
    Set DB_LoadNSECanalInfo = goconn.OpenRecordset(strSql)
    
    Exit Function
    
ErrHandler:
    goCWShared.ShowVBError Err.Number, Err.Description, const_ModuleName & ".DB_LoadNSECanalInfo()", "", vbCritical
    Set DB_LoadNSECanalInfo = Nothing
End Function

'***************************************************************************
'Function:DB_LoadNSECanalInfo2
'Author: KDunne
'Date:12/6/2001
'Purpose: Loads remaining fields from TEMPL_TREAT_NSE_CANAL_INFO table.
'Notes: There's a limit on the amount of data that can be brought across tuxedo.
'        The full query exceeds that limit, so we had to break it into two seperate
'        queries.
'******************************************************************************
Private Function DB_LoadNSECanalInfo2(ByVal lngDentalEncounterID As Long) As ADODB.Recordset
    Dim strSql As String
    
    Dim sReferencesVar As String
    
    'ITT change
    If goChcsiiConn.UseSQLServer Then
        sReferencesVar = "[REFERENCES]"
    Else
        sReferencesVar = "REFERENCES"
    End If
    
    On Error GoTo ErrHandler
    strSql = "Select INTRUMENT_TECH, FILE_SYSTEM, " & sReferencesVar & ", OBTURATION_MAT, " & _
        "COMMENTS, MAF, OBTURATION_TECH, ET_CANAL_INFO_ID " & _
        " From TEMPL_TREAT_NSE_CANAL_INFO " & _
        " WHERE DENTAL_ENCOUNTER_ID = " & lngDentalEncounterID

    Set DB_LoadNSECanalInfo2 = goconn.OpenRecordset(strSql)
    Exit Function
    
ErrHandler:
    goCWShared.ShowVBError Err.Number, Err.Description, const_ModuleName & ".DB_LoadNSECanalInfo2()", "", vbCritical
    Set DB_LoadNSECanalInfo2 = Nothing
End Function

