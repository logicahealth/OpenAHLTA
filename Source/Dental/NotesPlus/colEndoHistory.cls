VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colExamEndoSubj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements clsISection

Private mlngDentalEncounterId As Long
Private mlngPatientUnitId As Long

Private mstrChiefComplaint As String
Private mblnPainHistory As Boolean
Private mblnCurrentPain As Boolean
Private mlngPainIntensity As PainIntensity
Private mobjPainChar As clsExamEndoPain
Private mcolItems As Collection
Private mlngSubjectiveId As Long
Private mlngDirtyFlag As DirtyFlagEnum
Private mobjSOAP As clsExamSOAP

Private Const const_SUBJ_CHIEF_COMPLAINT_CODE = 1

Public Enum PainIntensity
    Not_Selected = 0
    NoPain = 11
    Level_01 = 10
    Level_02 = 9
    Level_03 = 8
    Level_04 = 7
    Level_05 = 6
    Level_06 = 5
    Level_07 = 4
    Level_08 = 3
    Level_09 = 2
    Level_10 = 1
End Enum

Private Sub Class_Initialize()
    Set mcolItems = New Collection
End Sub

Private Sub Class_Terminate()
    Set mobjPainChar = Nothing
    Set mcolItems = Nothing
    Set mobjSOAP = Nothing
End Sub
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mcolItems.[_NewEnum]
End Property
Public Property Get Item(vntIndexKey As Variant) As clsTreatPGraft
Attribute Item.VB_UserMemId = 0
  Set Item = mcolItems(vntIndexKey)
End Property
Public Property Get Count() As Long
    Count = mcolItems.Count
End Property
Public Property Get Chief_Complaint() As String
    Chief_Complaint = mstrChiefComplaint
End Property
'Public Property Let Chief_Complaint(ByVal strValue As String)
'    mstrChiefComplaint = strValue
'End Property
Public Property Get SOAP() As clsExamSOAP
    Set SOAP = mobjSOAP
End Property
Public Property Set SOAP(ByVal objSOAP As clsExamSOAP)
    Set mobjSOAP = objSOAP
End Property
Public Property Get Subjective_ID() As Long
    Subjective_ID = mlngSubjectiveId
End Property
Public Property Get History_Of_Pain() As Boolean
    History_Of_Pain = mblnPainHistory
End Property
Public Property Let History_Of_Pain(ByVal blnHistory As Boolean)
    mblnPainHistory = blnHistory
    UpdateDirtyFlag
End Property
Public Property Get DirtyFlag() As DirtyFlagEnum
    DirtyFlag = mlngDirtyFlag
End Property
Public Property Let DirtyFlag(ByVal lngValue As DirtyFlagEnum)
    mlngDirtyFlag = lngValue
End Property
Public Property Get Presently_In_Pain() As Boolean
    Presently_In_Pain = mblnCurrentPain

End Property
Public Property Let Presently_In_Pain(ByVal blnPresentPain As Boolean)
    mblnCurrentPain = blnPresentPain
    UpdateDirtyFlag
End Property

Public Property Get Intensity() As Long
    Intensity = mlngPainIntensity
End Property
Public Property Let Intensity(ByVal lngPain As Long)
    mlngPainIntensity = lngPain
    UpdateDirtyFlag
End Property
Public Property Get Pain_Characteristics(ByVal intIndex As Integer) As clsExamEndoPain
    Dim objPain As clsExamEndoPain
    
    If intIndex > mcolItems.Count Then
        Set objPain = New clsExamEndoPain
        mcolItems.Add objPain
    End If
    
    Set Pain_Characteristics = mcolItems(intIndex)
        
    Set objPain = Nothing
End Property

Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)
    
    If mcolItems(vntIndexKey).Pain_Characteristic_ID <> 0 Then
        'Set the dirtyflag to deleted
        mcolItems(vntIndexKey).DirtyFlag = dfdelete
        mcolItems(vntIndexKey).Clear
    Else
        'The object has not been saved, so remove it.
        mcolItems.Remove vntIndexKey
    End If
    
End Sub
'*******************************************************************
'Function: Get_Intensity_Short_Desc
'Author: KDunne
'Date: 11/1/2001
'Returns: String.
'Purpose: Returns the short description for a given intensity level.
'*******************************************************************
Public Function Get_Intensity_Short_Desc(ByVal intLevel As PainIntensity) As String
    Dim strDesc As String
    
    Select Case intLevel
        Case NoPain
            strDesc = "No Pain"
        Case Level_01
            strDesc = ""
        Case Level_02
            strDesc = "Mild"
        Case Level_03
            strDesc = ""
        Case Level_04
            strDesc = ""
        Case Level_05
            strDesc = "Moderate"
        Case Level_06
            strDesc = ""
        Case Level_07
            strDesc = ""
        Case Level_08
            strDesc = "Severe"
        Case Level_09
            strDesc = ""
        Case Level_10
            strDesc = "Worst Possible Pain"
    
    End Select
    
    Get_Intensity_Short_Desc = strDesc
End Function

'*******************************************************************
'Function: Get_Intensity_Long_Desc
'Author: KDunne
'Date: 11/1/2001
'Returns: String.
'Purpose: Returns the long description for a given intensity level.
'*******************************************************************
Public Function Get_Intensity_Long_Desc(ByVal intLevel As PainIntensity) As String
    Dim strDesc As String
    
    Select Case intLevel
        Case NoPain
            strDesc = "No Pain"
        Case Level_01
            strDesc = "Mild, aware of undercurrent of mild pain."
        Case Level_02
            strDesc = "Mild, aware of undercurrent of mild pain."
        Case Level_03
            strDesc = "Mild, aware of undercurrent of mild pain."
        Case Level_04
            strDesc = "Moderate, but able to continue some physical activity."
        Case Level_05
            strDesc = "Moderate, but able to continue some physical activity."
        Case Level_06
            strDesc = "Moderate, but able to continue some physical activity."
        Case Level_07
            strDesc = "Severe, can't concentrate and can't do all but simple things."
        Case Level_08
            strDesc = "Severe, can't concentrate and can't do all but simple things."
        Case Level_09
            strDesc = "Severe, can't concentrate and can't do all but simple things."
        Case Level_10
            strDesc = "Worst Possible Pain, must take care of pain."
    
    End Select
    
    Get_Intensity_Long_Desc = strDesc
End Function
'*******************************************************************
'Function: Get_Intensity_Text_Desc
'Author: KDunne
'Date: 8/26/2002
'Returns: String.
'Purpose: Returns the text description for a given intensity level.
'*******************************************************************
Public Function Get_Intensity_Text_Desc(ByVal intLevel As PainIntensity) As String
    Dim strDesc As String
    
    Select Case intLevel
        Case NoPain
            strDesc = "No Pain"
        Case Level_01
            strDesc = "Mild"
        Case Level_02
            strDesc = "Mild"
        Case Level_03
            strDesc = "Mild"
        Case Level_04
            strDesc = "Moderate"
        Case Level_05
            strDesc = "Moderate"
        Case Level_06
            strDesc = "Moderate"
        Case Level_07
            strDesc = "Severe"
        Case Level_08
            strDesc = "Severe"
        Case Level_09
            strDesc = "Severe"
        Case Level_10
            strDesc = "Worst Possible Pain"
    
    End Select
    
    Get_Intensity_Text_Desc = strDesc
End Function
Private Property Get clsISection_DentalEncounterID() As Long
    clsISection_DentalEncounterID = mlngDentalEncounterId
End Property

Private Function clsISection_GetNote() As String
    Dim objPainChar As clsExamEndoPain
    Dim strNote As String
    Dim strHistory As String
    Dim strCurrentPain As String
    Dim strPainChar As String
    
    strNote = Empty
        
    If mblnPainHistory Then
        strNote = strNote & "Patient has a history of pain;"
    End If
    
'    If mblnCurrentPain Then
'        strNote = strNote & vbTab & "Patient is presently in pain." & vbCrLf & _
'            "The intensity of current pain is " & mlngPainIntensity & " - " & _
'            Get_Intensity_Short_Desc(mlngPainIntensity) & "." & vbCrLf
'    End If
    
    If mlngPainIntensity > 0 Then
        strNote = strNote & " Pain Intensity: " & ExamEndo_GetPainIntensity_Level(mlngPainIntensity) & " - " & _
            Get_Intensity_Short_Desc(mlngPainIntensity) & ";"
    End If
    
    If mcolItems Is Nothing = False Then
        For Each objPainChar In mcolItems
            strPainChar = strPainChar & objPainChar.GetNote
        Next
    End If
    
    If strPainChar <> "" Then
    
        strNote = strNote & "Pain Characteristics:" & strPainChar
    End If
    
    If strNote <> "" Then
        strNote = "Subjective: " & strNote
    End If
    clsISection_GetNote = strNote
End Function

Private Function clsISection_GetRTFNote() As String
    Dim objPainChar As clsExamEndoPain
    Dim strNote As String
    Dim strHistory As String
    Dim strCurrentPain As String
    Dim strPainChar As String
    Dim strTempPain As String 'scr 26326
    
    strNote = Empty
        
    If mblnPainHistory Then
        strNote = strNote & "Patient has a history of pain;"
    End If
    
'    If mblnCurrentPain Then
'        strNote = strNote & vbTab & "Patient is presently in pain." & vbCrLf & _
'            "The intensity of current pain is " & mlngPainIntensity & " - " & _
'            Get_Intensity_Short_Desc(mlngPainIntensity) & "." & vbCrLf
'    End If
    
    If mlngPainIntensity > 0 Then
        strNote = strNote & " Pain Intensity: " & ExamEndo_GetPainIntensity_Level(mlngPainIntensity) & "/10 " & _
            Get_Intensity_Text_Desc(mlngPainIntensity) & ";"
    End If
    
    If mcolItems Is Nothing = False Then
        For Each objPainChar In mcolItems
            strPainChar = strPainChar & objPainChar.GetNote
        Next
    End If
    
    'KDunne 8/6/2002
    'SCR 26326 - strPainChar has carriage returns in it, so it passes the empty
    '            string check. Use a temp string, and remove all carriage returns.
    strTempPain = Replace(strPainChar, vbCrLf, "")
    
    If Trim(strTempPain) <> "" Then
    
        strNote = strNote & crtfStartControlName & "Pain Characteristics:" & _
            strPainChar & crtfEndControlName
    End If
    
    If strNote <> "" Then
        strNote = crtfStartControlName & "Subjective: " & strNote & crtfEndControlName
    End If
    clsISection_GetRTFNote = strNote
End Function

Private Function clsISection_LoadSection(ByVal PatientID As Long, ByVal DentalEncounterID As Long) As Boolean
    Dim objRS As ADODB.Recordset
    Dim objItem As clsExamEndoPain
    
    mlngDentalEncounterId = DentalEncounterID
    mlngPatientUnitId = PatientID
    
    ResetValues
    
    'Load Chief Complaint.
    'TEMP - We're currently accessing the info straight from the db.
    'when the SOAP object is re-written, we'll get it from there.
    'KDunne 12/17/2001
    'We're now using the SOAP object for Chief Complaint.
'''    Set objRS = DB_LOAD_CHIEF_COMPLAINT(mlngDentalEncounterId)
'''    If Not objRS.EOF Then
'''        mstrChiefComplaint = objRS.Fields("SUBJECTIVE_COMMENT")
'''    Else
'''        mstrChiefComplaint = ""
'''    End If
'''
'''    Set objRS = Nothing
    Set objRS = DB_LOAD_ENDO_SUBJECTIVE(mlngDentalEncounterId)
    
    With objRS
        If .EOF Then
            mlngDirtyFlag = dfAdd
            mlngSubjectiveId = 0
            mblnPainHistory = False
            mblnCurrentPain = False
            mlngPainIntensity = Not_Selected
        
        Else
            mlngDirtyFlag = dfUpdate
            mlngSubjectiveId = .Fields("EX_ENDO_SUBJ_ID")
            mblnPainHistory = CBool(.Fields("HISTORY_PAIN"))
            mlngPainIntensity = .Fields("PAIN_INTENSITY")
            
            Load_Pain_Characteristics .Fields("EX_ENDO_SUBJ_ID")
        End If
    End With

    clsISection_LoadSection = True
    
End Function

Private Property Get clsISection_PatientID() As Long
    clsISection_PatientID = mlngPatientUnitId
End Property

Private Function clsISection_SaveSection() As Boolean
    Dim strSql As String
    Dim blnReturn As Boolean
    Dim objPainChar As clsExamEndoPain
    Dim objRS As ADODB.Recordset
    
    strSql = Empty
    Select Case mlngDirtyFlag
        Case dfNoAction
            blnReturn = True
        Case dfAdd
            
            mlngSubjectiveId = goconn.GetNextID(esTEMPL_EXAM_ENDO_SUBJECTIVE_SEQ)
'            strSql = "SELECT templ_exam_endo_subjective_seq.NEXTVAL From Dual"
'            Set objRS = goconn.OpenRecordset(strSql)
'            mlngSubjectiveId = objRS.Fields("NEXTVAL")
            
            strSql = "INSERT INTO TEMPL_EXAM_ENDO_SUBJECTIVE " & _
                     "(EX_ENDO_SUBJ_ID, DENTAL_ENCOUNTER_ID, " & _
                     " HISTORY_PAIN, PAIN_INTENSITY) " & _
                     "VALUES (" & mlngSubjectiveId & ", " & mlngDentalEncounterId & ", " & _
                     CInt(mblnPainHistory) & ", " & mlngPainIntensity & " )"

        Case dfUpdate
            strSql = "UPDATE TEMPL_EXAM_ENDO_SUBJECTIVE " & _
                     " SET DENTAL_ENCOUNTER_ID = " & mlngDentalEncounterId & ", " & _
                     " HISTORY_PAIN = " & CInt(mblnPainHistory) & ", " & _
                     " PAIN_INTENSITY = " & mlngPainIntensity & _
                     " WHERE EX_ENDO_SUBJ_ID = " & mlngSubjectiveId
        Case dfdelete
            strSql = "DELETE FROM TEMPL_EXAM_ENDO_SUBJECTIVE WHERE " & _
                     " EX_ENDO_SUBJ_ID = " & mlngSubjectiveId
    End Select
    
    If strSql <> "" Then
        goconn.ExecuteSQL (strSql)
    End If
    
    mlngDirtyFlag = dfNoAction
    
    If mcolItems.Count > 0 Then
        For Each objPainChar In mcolItems
            objPainChar.Subjective_ID = mlngSubjectiveId
            objPainChar.SaveSection
        Next
    End If
    
    Set objRS = Nothing
    Set objPainChar = Nothing
End Function
Private Sub Load_Pain_Characteristics(ByVal lngSubjectiveId As Long)
    Dim objRS As ADODB.Recordset
    
    Set objRS = DB_LOAD_ENDO_SUBJ_CHARACTERISTICS(lngSubjectiveId)
    
    Set mcolItems = New Collection
    
    Do While Not objRS.EOF
        Set mobjPainChar = New clsExamEndoPain
        
        With mobjPainChar
            .Comments = objRS.Fields("COMMENTS") & ""
            .Frequency = objRS.Fields("FREQUENCY") & ""
            .HxTrauma = objRS.Fields("TRAUMA_HISTORY")
            .Intensity = objRS.Fields("INTENSITY")
            .Location = objRS.Fields("LOCATION")
            
            .Onset = objRS.Fields("ONSET") & ""
            .PainQuality = objRS.Fields("PAIN_QUALITY") & ""
            .ProvokedBy = objRS.Fields("PROVOKED_BY") & ""
            .RelievedBy = objRS.Fields("RELIEVED_BY") & ""
            .Spontaneous = objRS.Fields("SPONTANEOUS")
            .Pain_Characteristic_ID = objRS.Fields("EX_ENDO_PAIN_CHARS_ID")
            .Subjective_ID = lngSubjectiveId
            .DirtyFlag = dfNoAction
            
        End With
        
        mcolItems.Add mobjPainChar
        
        objRS.MoveNext
    Loop
End Sub
Private Sub ResetValues()
    Set mobjPainChar = Nothing
    mblnPainHistory = False
    mblnCurrentPain = False
    mlngPainIntensity = NoPain
    mlngSubjectiveId = 0
End Sub
Private Function DB_LOAD_ENDO_SUBJECTIVE(ByVal lngDentalEncounterId As Long) As ADODB.Recordset
    Dim strSql As String
    
    strSql = Empty
        
    strSql = "SELECT EX_ENDO_SUBJ_ID, DENTAL_ENCOUNTER_ID, HISTORY_PAIN, " & _
              " PAIN_INTENSITY " & _
              "FROM TEMPL_EXAM_ENDO_SUBJECTIVE " & _
              "WHERE DENTAL_ENCOUNTER_ID = " & lngDentalEncounterId
              
    Set DB_LOAD_ENDO_SUBJECTIVE = goconn.OpenRecordset(strSql)
    
End Function

Private Function DB_LOAD_ENDO_SUBJ_CHARACTERISTICS(ByVal lngSubjectiveId As Long)
    Dim strSql As String
    
    strSql = "SELECT EX_ENDO_PAIN_CHARS_ID, EX_ENDO_SUBJ_ID , LOCATION, " & _
                " ONSET , SPONTANEOUS , PAIN_QUALITY , PROVOKED_BY , FREQUENCY ," & _
                " RELIEVED_BY , INTENSITY , TRAUMA_HISTORY , Comments " & _
                " From TEMPL_EXAM_ENDO_PAIN_CHARS " & _
                " WHERE EX_ENDO_SUBJ_ID = " & lngSubjectiveId
                
    Set DB_LOAD_ENDO_SUBJ_CHARACTERISTICS = goconn.OpenRecordset(strSql)
    
End Function

Private Function DB_LOAD_CHIEF_COMPLAINT(ByVal lngDentalEncounterId As Long) As ADODB.Recordset
    Dim strSql As String
    
    strSql = "SELECT EVALUATION_SUBJECTIVE.SUBJECTIVE_COMMENT " & _
             " From EVALUATION_SUBJECTIVE, NOTE_EVALUATION " & _
             " Where EVALUATION_SUBJECTIVE.NOTE_EVALUATION_ID = NOTE_EVALUATION.NOTE_EVALUATION_ID " & _
             " AND EVALUATION_SUBJECTIVE.NOTE_CODES_ID = " & const_SUBJ_CHIEF_COMPLAINT_CODE & _
             " AND NOTE_EVALUATION.DENTAL_ENCOUNTER_ID = " & lngDentalEncounterId
    
    Set DB_LOAD_CHIEF_COMPLAINT = goconn.OpenRecordset(strSql)

End Function
Private Function UpdateDirtyFlag()

    If mlngDirtyFlag = dfNoAction Then
        mlngDirtyFlag = dfUpdate
    End If
    
End Function
