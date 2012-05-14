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
Attribute VB_Name = "clsExEndoCmmnts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements clsISection

Private mstrComments As String
Private mlngDentalEncId As Long
Private mlngPatUnitId As Long
Private mlngDirtyFlag As DirtyFlagEnum
Private mlngCommentId As Long

Public Enum ExamEndoCommentsModeEnum
    ProblemHistory = 0
    ObjectiveFindings = 1
End Enum
Public Property Get Comments() As String
    Comments = mstrComments
End Property
Public Property Let Comments(ByVal strValue As String)
    mstrComments = strValue
    UpdateDirtyFlag
End Property
Public Property Get DirtyFlag() As DirtyFlagEnum
    DirtyFlag = mlngDirtyFlag
End Property
Public Function Get_User_Comments(ByVal lngCommentsMode As ExamEndoCommentsModeEnum) As Boolean
    Dim objComments As CComments
    Dim blnReturn As Boolean
    
    On Error GoTo ErrHandler:
    
    Set objComments = New CComments
    
    If lngCommentsMode = ObjectiveFindings Then
        objComments.Caption = "Endodontic Objective Findings & Observations Comments"
    Else
        objComments.Caption = "Endodontic Problem History Comments"
    End If
    
    objComments.Comments = mstrComments
    
    If objComments.Get_User_Input() Then
        blnReturn = True
        mstrComments = objComments.Comments
        
        UpdateDirtyFlag
    Else
        blnReturn = False
    End If
    
    Get_User_Comments = blnReturn
    
FuncExit:
    Set objComments = Nothing
    Exit Function
    
ErrHandler:
    Get_User_Comments = False
    GoTo FuncExit
End Function

Private Sub Class_Initialize()
    mlngDirtyFlag = dfNoAction
End Sub

Private Property Get clsISection_DentalEncounterID() As Long
    clsISection_DentalEncounterID = mlngDentalEncId
End Property

Private Function clsISection_GetNote() As String
    Dim strNote As String
    strNote = Empty
    
    If Trim(mstrComments) <> "" Then
        strNote = "Comments: " & mstrComments & ";"
    End If
    
    clsISection_GetNote = strNote
End Function

Private Function clsISection_GetRTFNote() As String
    Dim strNote As String
    strNote = Empty
    
    If Trim(mstrComments) <> "" Then
        strNote = crtfStartControlName & "Endodontic Comments: " & mstrComments & _
            ";" & crtfEndControlName
    End If
    
    clsISection_GetRTFNote = strNote
End Function

Private Function clsISection_LoadSection(ByVal PatientID As Long, ByVal DentalEncounterID As Long) As Boolean
    Dim objRS As ADODB.Recordset
    
    On Error GoTo ErrHandler:
    
    mlngPatUnitId = PatientID
    mlngDentalEncId = DentalEncounterID
    
    Set objRS = DB_LOAD_COMMENTS(mlngDentalEncId)

    If objRS.EOF Then
        mlngDirtyFlag = dfAdd
        
        mstrComments = Empty
    Else
        mlngDirtyFlag = dfUpdate
        mstrComments = objRS.Fields("COMMENTS")
        mlngCommentId = objRS.Fields("EX_ENDO_COMMENTS_ID")
    End If

    clsISection_LoadSection = True
    
    Exit Function
ErrHandler:
    clsISection_LoadSection = False
End Function

Private Property Get clsISection_PatientID() As Long
    clsISection_PatientID = mlngDentalEncId
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
        
            mlngCommentId = goconn.GetNextID(esTEMPL_EXAM_ENDO_COMMENTS_SEQ)
'            strSql = "SELECT TEMPL_EXAM_ENDO_COMMENTS_seq.NEXTVAL From Dual"
'            Set objRS = goconn.OpenRecordset(strSql)
'            mlngCommentId = objRS.Fields("NEXTVAL")
            
            strSql = "INSERT INTO TEMPL_EXAM_ENDO_COMMENTS " & _
                     "(EX_ENDO_COMMENTS_ID, DENTAL_ENCOUNTER_ID, " & _
                     " COMMENTS) " & _
                     "VALUES (" & mlngCommentId & ", " & mlngDentalEncId & ", '" & _
                     Scrub_Apostrophe(mstrComments) & "')"

        Case dfUpdate
            strSql = "UPDATE TEMPL_EXAM_ENDO_COMMENTS " & _
                     " SET DENTAL_ENCOUNTER_ID = " & mlngDentalEncId & ", " & _
                     " COMMENTS = '" & Scrub_Apostrophe(mstrComments) & "'" & _
                     " WHERE EX_ENDO_COMMENTS_ID = " & mlngCommentId
        Case dfdelete
            strSql = "DELETE FROM TEMPL_EXAM_ENDO_COMMENTS WHERE " & _
                     " EX_ENDO_COMMENTS_ID = " & mlngCommentId
    End Select
    
    If strSql <> "" Then
        goconn.ExecuteSQL (strSql)
    End If
    mlngDirtyFlag = dfNoAction
    
'    If mColItems.Count > 0 Then
'        For Each objPainChar In mColItems
'            objPainChar.Subjective_ID = mlngCommentId
'            objPainChar.SaveSection
'        Next
'    End If
    
    Set objRS = Nothing

End Function

Private Function DB_LOAD_COMMENTS(ByVal lngDentalEncounterId As Long) As ADODB.Recordset
    Dim strSql As String
    
    strSql = "Select EX_ENDO_COMMENTS_ID, COMMENTS from TEMPL_EXAM_ENDO_COMMENTS where DENTAL_ENCOUNTER_ID  = " & _
                lngDentalEncounterId
    
    Set DB_LOAD_COMMENTS = goconn.OpenRecordset(strSql)
    
End Function
Private Function UpdateDirtyFlag()

    If mlngDirtyFlag = dfNoAction Then
        mlngDirtyFlag = dfUpdate
    End If
    
End Function
