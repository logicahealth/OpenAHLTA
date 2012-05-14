VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExEndoAdFnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Private mlngDentalEncounterId As Long
Private mlngPatUnitId As Long
Private mlngDirtyFlag As DirtyFlagEnum
Private mstrToothNum As String
Private mstrCanals As String
Private mstrAddFindings As String
Private mlngAddFindId As Long
Private mlngAddFind As ExamEndo_AddFindEnum

Public Enum ExamEndo_AddFindEnum
    exEnAF_None
    exEnAF_ShortFill
    exEnAF_LongFill
    exEnAF_NonDenseObut
    exEnAF_UntreatedCanal
    exEnAF_SilverPoint
    exEnAF_Post
    exEnAF_Perforation
    exEnAF_VerticalRootFracture
End Enum

Public Function Get_AddFind_Enum_Desc(ByVal lngValue As ExamEndo_AddFindEnum) As String
    Dim strReturn
    
    Select Case lngValue
        Case ExamEndo_AddFindEnum.exEnAF_ShortFill
            strReturn = "Short fill"
        Case ExamEndo_AddFindEnum.exEnAF_LongFill
            strReturn = "Long fill"
        Case ExamEndo_AddFindEnum.exEnAF_NonDenseObut
            strReturn = "Non-dense obturation "
        Case ExamEndo_AddFindEnum.exEnAF_UntreatedCanal
            strReturn = "Untreated canal"
        Case ExamEndo_AddFindEnum.exEnAF_SilverPoint
            strReturn = "Silver point"
        Case ExamEndo_AddFindEnum.exEnAF_Post
            strReturn = "Post"
        Case ExamEndo_AddFindEnum.exEnAF_Perforation
            strReturn = "Perforation"
        Case ExamEndo_AddFindEnum.exEnAF_VerticalRootFracture
            strReturn = "Vertical Root Fracture"
        Case Else
            strReturn = Empty
    End Select
    Get_AddFind_Enum_Desc = strReturn
End Function
Private Sub Class_Initialize()
    mlngDirtyFlag = dfNoAction
    mlngAddFindId = 0
End Sub
Public Property Get AdditionalFindings() As String
    AdditionalFindings = mstrAddFindings
End Property
Public Property Let AdditionalFindings(ByVal strValue As String)
    mstrAddFindings = strValue
    UpdateDirtyFlag
End Property
Public Property Get AdditionalFinding() As ExamEndo_AddFindEnum
    AdditionalFinding = mlngAddFind
End Property
Public Property Let AdditionalFinding(ByVal lngValue As Long)
    mlngAddFind = lngValue
End Property
Public Property Get AdditionalFindingsID() As Long
    AdditionalFindingsID = mlngAddFindId
End Property
Public Property Let AdditionalFindingsID(ByVal lngValue As Long)
    mlngAddFindId = lngValue
End Property

Public Property Get ToothNumber() As String
    ToothNumber = mstrToothNum
End Property
Public Property Let ToothNumber(ByVal strValue As String)
    mstrToothNum = strValue
    UpdateDirtyFlag
End Property
Public Property Get Canals() As String
    Canals = mstrCanals
End Property
Public Property Let Canals(ByVal strValue As String)
    mstrCanals = strValue
    UpdateDirtyFlag
End Property
Public Property Get DirtyFlag() As DirtyFlagEnum
    DirtyFlag = mlngDirtyFlag
End Property
Public Property Let DirtyFlag(ByVal lngValue As DirtyFlagEnum)
    mlngDirtyFlag = lngValue
End Property
Public Property Get DentalEncounterID() As Long
    DentalEncounterID = mlngDentalEncounterId
End Property
Public Property Let DentalEncounterID(ByVal lngValue As Long)
    mlngDentalEncounterId = lngValue
End Property

'****************************************************
'Sub: UpdateDirtyFlag
'Author: KDunne
'Date: 11/5/2001
'Purpose: Updates module level dirty flag field
'****************************************************
Public Sub UpdateDirtyFlag()
    If mlngDirtyFlag = dfNoAction Then
        If mlngAddFindId = 0 Then
            mlngDirtyFlag = dfAdd
        Else
            mlngDirtyFlag = dfUpdate
        End If
    End If

End Sub
'*******************************************************************
'Function: SaveSection
'Author: Kieran Dunne
'Date: 11/08/2001
'Purpose: Saves users selections.
'*******************************************************************
Public Function SaveSection() As Boolean
    Dim strSql As String
    Dim blnReturn As Boolean
    Dim objRS As ADODB.Recordset
        
    Select Case mlngDirtyFlag
        Case dfNoAction
            blnReturn = True
            GoTo FuncExit
        Case dfAdd
            'First make sure that our object contains data
            If ContainsData Then
                'Get the primary key
                mlngAddFindId = goconn.GetNextID(esTEMPL_EXAM_ENDO_ADD_F_SEQ)
'                strSql = "SELECT TEMPL_EXAM_ENDO_ADD_F_seq.NEXTVAL From Dual"
'                Set objRS = goconn.OpenRecordset(strSql)
'                mlngAddFindId = objRS("NEXTVAL")
                
                strSql = "INSERT INTO TEMPL_EXAM_ENDO_ADD_FINDINGS " & _
                        " (EX_ENDO_ADD_FIND_ID, DENTAL_ENCOUNTER_ID, TOOTH_NUM, " & _
                        " CANALS, ADD_FINDING) " & _
                        " VALUES (" & mlngAddFindId & ", " & mlngDentalEncounterId & "," & _
                        " '" & mstrToothNum & "', '" & mstrCanals & "', " & mlngAddFind & ")"
            Else
                blnReturn = False
                GoTo FuncExit
            End If
        Case dfUpdate
            strSql = "Update TEMPL_EXAM_ENDO_ADD_FINDINGS " & _
                     " SET TOOTH_NUM = '" & mstrToothNum & "', CANALS = '" & mstrCanals & _
                     "', ADD_FINDING = " & mlngAddFind & " " & _
                     " WHERE EX_ENDO_ADD_FIND_ID = " & mlngAddFindId
                     
        Case dfdelete
            'KDunne 11/4/2003
            'SCR 44672 - Only delete if the record exists in the db.
            If mlngAddFindId > 0 Then
                strSql = "DELETE FROM TEMPL_EXAM_ENDO_ADD_FINDINGS WHERE " & _
                         " EX_ENDO_ADD_FIND_ID = " & mlngAddFindId
                         
            End If
    End Select
        
    If strSql <> "" Then
        goconn.ExecuteSQL (strSql)
    End If
    mlngDirtyFlag = dfNoAction

FuncExit:
    SaveSection = True
    Set objRS = Nothing
End Function
'*******************************************************************
'Function: GetNote
'Author: Kieran Dunne
'Date: 11/08/2001
'Purpose: Builds a narrative summary of the user's selections.
'*******************************************************************
Public Function GetNote() As String
    Dim strNote As String
    Dim objAddFindCtrl As CAddFindings
    Dim objToothCtrl As CToothControl
    
    Set objAddFindCtrl = New CAddFindings
    Set objToothCtrl = New CToothControl
    
    strNote = Empty
    
    If Trim(mstrToothNum) <> "" Then
        strNote = strNote & " Tooth Number = " & objToothCtrl.Get_Tooth_Name(mstrToothNum) & "; "
    End If
    
    If Trim(mstrCanals) <> "" Then
        strNote = strNote & " Canal = " & objAddFindCtrl.Build_Display_String(mstrCanals) & ";"
    End If
    
    If mlngAddFind <> ExamEndo_AddFindEnum.exEnAF_None Then
        strNote = strNote & " Additional Finding = " & objAddFindCtrl.Get_Additional_Findings_Enum_Desc(mlngAddFind) & ";"
    End If
    
    GetNote = strNote
    Set objAddFindCtrl = Nothing
    Set objToothCtrl = Nothing
End Function

'*****************************************************
'Function: ContainsData
'Author: KDunne
'Date: 11/8/2001
'Purpose: Checks to see if object contains data.
'*****************************************************
Private Function ContainsData() As Boolean
    Dim blnReturn As Boolean
    
    blnReturn = False
    
    If Trim(mstrToothNum) <> "" Then
        blnReturn = True
    End If
    
    If Trim(mstrCanals) <> "" Then
        blnReturn = True
    End If
    
    If Trim(mstrAddFindings) <> "" Then
        blnReturn = True
    End If
    
    ContainsData = blnReturn
    
End Function

''*********************************************************************
''Function: Get_Add_Find_Enum_Desc
''Author:Kdunne
''Date:1/30/2002
''Purpose: Given the enum, function returns the description for the
''           ExamEndoAddFind_Canals enums.
''*********************************************************************
'Public Function Get_Add_Find_Enum_Desc(ByVal lngValue As ExamEndoAddFind_Canals) As String
'    Dim strReturn As String
'
'    Select Case lngValue
'    Dim strReturn As String
'
'    Select Case lngValue
'        Case NotesPlusUtil_AddFindings.npUtil_ShortFill
'            strReturn = "Short fill"
'        Case NotesPlusUtil_AddFindings.npUtil_LongFill
'            strReturn = "Long fill"
'        Case NotesPlusUtil_AddFindings.npUtil_NonDenseObturation
'            strReturn = "Non-dense obturation"
'        Case NotesPlusUtil_AddFindings.npUtil_UntreatedCanal
'            strReturn = "Untreated canal"
'        Case NotesPlusUtil_AddFindings.npUtil_SilverPoint
'            strReturn = "Silver point"
'        Case NotesPlusUtil_AddFindings.npUtil_Post
'            strReturn = "Post"
'        Case NotesPlusUtil_AddFindings.npUtil_Perforation
'            strReturn = "Perforation"
'        Case NotesPlusUtil_AddFindings.npUtil_VerticalRootFracture
'            strReturn = "Vertical Root Fracture"
'        Case Else
'            strReturn = Empty
'
'    End Select
'
'    Get_Add_Find_Enum_Desc = strReturn
'End Function
