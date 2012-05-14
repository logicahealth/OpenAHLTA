VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colTrNSETooth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements clsISection

Private mcolTeethInfo As Collection
Private mcolCanalnfo As Collection
Private mobjDentalNotes As CHCSII_Dental_DXTX.clsNotes

Private mlngDentalEncounterId As Long
Private mlngPatUnitId As Long
Private mlngDirtyFlag As DirtyFlagEnum

Private Const const_ModuleName = "colTrNSETooth"

Private Sub Class_Terminate()
    Set mcolCanalnfo = Nothing
    Set mcolTeethInfo = Nothing
    Set mobjDentalNotes = Nothing
End Sub
'**************************************************************************************
'Sub: Delete
'Author:KDunne
'Date:5/14/2002
'Purpose: Deletes the selected tooth object
'**************************************************************************************
Public Sub Delete(ByVal lngIndex As Long)
    Dim objToothInfo As clsTrNSETooth

    On Error GoTo ErrHandler
    
    'Make sure our Index is valid
    If lngIndex > mcolTeethInfo.Count Then GoTo FuncExit
    
    Set objToothInfo = mcolTeethInfo.Item(lngIndex)
    
    objToothInfo.DirtyFlag = dfdelete
    objToothInfo.SaveSection
    
    mcolTeethInfo.Remove (lngIndex)
FuncExit:
    Set objToothInfo = Nothing
    Exit Sub

ErrHandler:
    goCWShared.ShowVBError Err.Number, Err.Description, const_ModuleName & ".Delete()", "", vbCritical
    GoTo FuncExit
End Sub

Public Property Get Teeth() As Collection
    Set Teeth = mcolTeethInfo
End Property
Public Property Get Tooth_Info(ByVal lngIndex As Long) As clsTrNSETooth
    Dim objToothInfo As clsTrNSETooth
    
    If lngIndex > mcolTeethInfo.Count Then
        Set objToothInfo = New clsTrNSETooth
        objToothInfo.DentalEncounterID = mlngDentalEncounterId
        
        mcolTeethInfo.Add objToothInfo
    End If
    
    Set Tooth_Info = mcolTeethInfo(lngIndex)
    
    Set objToothInfo = Nothing
End Property
Public Property Get clsISection_DentalEncounterID() As Long
    clsISection_DentalEncounterID = mlngDentalEncounterId
End Property

Public Property Get DentalNotes() As CHCSII_Dental_DXTX.clsNotes
    Set DentalNotes = mobjDentalNotes
End Property
Public Property Set DentalNotes(ByVal objNotes As CHCSII_Dental_DXTX.clsNotes)
    Set mobjDentalNotes = objNotes
End Property


Private Function clsISection_GetNote() As String
    Dim strNote As String
    Dim strToothInfo As String
    Dim strCanalInfo As String
    Dim objToothInfo As clsTrNSETooth
    
    strNote = Empty
    strToothInfo = Empty
    strCanalInfo = Empty
    
    If mcolTeethInfo Is Nothing = False Then
        For Each objToothInfo In mcolTeethInfo
            strToothInfo = strToothInfo & objToothInfo.GetNote & vbCrLf
        Next
        
        If strToothInfo <> "" Then
            strNote = strNote & vbCrLf & "Information Relating to the whole tooth" & _
                vbCrLf & strToothInfo
        End If
    End If
    
FuncExit:
    Set objToothInfo = Nothing
    clsISection_GetNote = strNote
    
    Exit Function
ErrHandler:
    goCWShared.ShowVBError Err.Number, Err.Description, const_ModuleName & ".GetNote()", "", vbCritical
    strNote = Empty
    GoTo FuncExit

End Function

Private Function clsISection_GetRTFNote() As String
    Dim strNote As String
    Dim strToothInfo As String
    Dim strCanalInfo As String
    Dim objToothInfo As clsTrNSETooth
    
    strNote = Empty
    strToothInfo = Empty
    strCanalInfo = Empty
    
    If mcolTeethInfo Is Nothing = False Then
        For Each objToothInfo In mcolTeethInfo
            strToothInfo = strToothInfo & objToothInfo.GetNote & vbCrLf
        Next
        
        If strToothInfo <> "" Then
            strNote = strNote & crtfStartControlName & "Information Relating to the whole tooth" & _
                crtfEndControlName & strToothInfo
        End If
    End If
    
FuncExit:
    Set objToothInfo = Nothing
    clsISection_GetRTFNote = strNote
    
    Exit Function
ErrHandler:
    goCWShared.ShowVBError Err.Number, Err.Description, const_ModuleName & ".clsISection_GetRTFNote()", "", vbCritical
    strNote = Empty
    GoTo FuncExit

End Function

Private Function clsISection_LoadSection(ByVal PatientID As Long, ByVal DentalEncounterID As Long) As Boolean
    Dim blnReturn As Boolean
    Dim objRS As ADODB.Recordset
    Dim objToothInfo As clsTrNSETooth
    Dim objCanalInfo As clsTrNSECanal
    
    On Error GoTo ErrHandler
    
    mlngPatUnitId = PatientID
    mlngDentalEncounterId = DentalEncounterID
    
    If mlngDentalEncounterId <= 0 Then
        clsISection_LoadSection = False
        GoTo FuncExit
    End If
    
    Set objRS = DB_LoadNSEToothSpecificInfo(mlngDentalEncounterId)
    
    Set mcolTeethInfo = New Collection
    Set mcolCanalnfo = New Collection
    
    With objRS
        If .EOF Then
            mlngDirtyFlag = dfAdd
            'reset property values here
        
        Else
            mlngDirtyFlag = dfUpdate
            'Set property values here
            
            Do While Not .EOF
                Set objToothInfo = New clsTrNSETooth
                
                objToothInfo.Tooth_Number = .Fields("TOOTH_NUMBER") & ""
                objToothInfo.Solvent = .Fields("LUBRICANTS") & ""
                objToothInfo.Cement = Val(.Fields("CEMENT") & "")
                objToothInfo.Cotton_Pellet = Val(.Fields("COTTON_PELLET") & "")
                objToothInfo.DentalEncounterID = mlngDentalEncounterId
                objToothInfo.ID = Val(.Fields("ET_TOOTH_INFO_ID") & "")
                objToothInfo.Intracanal_Medicament = Val(.Fields("INTRACANAL_MEDICANT") & "")
                objToothInfo.Magnification = .Fields("MAGINIFICATION") & ""
                objToothInfo.Material_Provision = Val(.Fields("MATERIAL") & "")
                objToothInfo.Occulusal_Adjustment = Val(.Fields("OCCUSAL_ADJ") & "")
                objToothInfo.Cement_Text = .Fields("CEMENT_TEXT") & ""
                objToothInfo.Intracanal_Medicament_Text = .Fields("ICM_TEXT") & ""
                objToothInfo.Material_Provision_Text = .Fields("MATERIAL_TEXT") & ""
                
                objToothInfo.DirtyFlag = dfNoAction
                mcolTeethInfo.Add objToothInfo
                
                .MoveNext
            Loop
            
        End If '.EOF
    End With 'objRS
    
    
FuncExit:
    clsISection_LoadSection = blnReturn
    Set objRS = Nothing
    Exit Function
ErrHandler:
   blnReturn = False
   goCWShared.ShowVBError Err.Number, Err.Description, const_ModuleName & ".DB_LoadNSEToothSpecificInfo()", "", vbCritical

   GoTo FuncExit
End Function

Private Property Get clsISection_PatientID() As Long
    clsISection_PatientID = mlngPatUnitId
End Property

Private Function clsISection_SaveSection() As Boolean
    Dim blnReturn As Boolean
    Dim objToothInfo As clsTrNSETooth
    
    On Error GoTo ErrHandler:
    
    If mcolTeethInfo Is Nothing = False Then
        For Each objToothInfo In mcolTeethInfo
            blnReturn = objToothInfo.SaveSection
        Next
    End If

FuncExit:
    Set objToothInfo = Nothing
    
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

'****************************************************
'Sub: DB_LoadNSEToothSpecificInfo
'Author: KDunne
'Date: 11/30/2001
'Purpose: Retrieves all of the Tooth Specific information for a dental encounter.
'****************************************************
Private Function DB_LoadNSEToothSpecificInfo(ByVal lngDentalEncounterID As Long) As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo ErrHandler:
    
    strSql = "Select ET_TOOTH_INFO_ID, DENTAL_ENCOUNTER_ID, " & _
            " TOOTH_NUMBER, LUBRICANTS, MAGINIFICATION, " & _
            " INTRACANAL_MEDICANT, COTTON_PELLET, MATERIAL, " & _
            "  CEMENT, OCCUSAL_ADJ, ICM_TEXT, MATERIAL_TEXT, " & _
            " CEMENT_TEXT From TEMPL_TREAT_NSE_TOOTH_INFO Where " & _
            " DENTAL_ENCOUNTER_ID = " & lngDentalEncounterID
            
    Set DB_LoadNSEToothSpecificInfo = goconn.OpenRecordset(strSql)
    
    Exit Function

ErrHandler:
    goCWShared.ShowVBError Err.Number, Err.Description, const_ModuleName & ".DB_LoadNSEToothSpecificInfo()", "", vbCritical
    Set DB_LoadNSEToothSpecificInfo = Nothing
End Function

