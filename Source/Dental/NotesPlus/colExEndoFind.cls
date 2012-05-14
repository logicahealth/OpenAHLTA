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
Attribute VB_Name = "colExEndoFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements clsISection

Private mlngDentalEncounterId As Long
Private mlngFindId As Long
Private mstrSwelling As String
Private mlngPatUnitId As Long
Private mlngDirtyFlag As DirtyFlagEnum
Private mcolItems As Collection

Public Property Get Swelling() As String
    Swelling = mstrSwelling
End Property
Public Property Let Swelling(ByVal strValue As String)
    mstrSwelling = strValue
    UpdateDirtyFlag
End Property
Public Property Get FindingsID() As Long
    FindingsID = mlngFindId
End Property

Private Sub Class_Initialize()
    Set mcolItems = New Collection
    Initialize_Values
End Sub

Private Sub Initialize_Values()
    mstrSwelling = "Right Buccal Vestibule|Left Buccal Vestibule|Right Mandibular Subperiosteal|" & _
        "Left Mandibular Subperiosteal|Mental Space|Submental Space|Right Sublingual Space|" & _
        "Left Sublingual Space|Right Submandibular Space|Left Submandibular Space|" & _
        "Ludwig's Angina (sublingual, submandibular, and submental spaces)|" & _
        "Right Buccal Space|Left Buccal Space|Right Submasseteric Space|Left Submasseteric Space|" & _
        "Right Superficial Temporal Space|Left Superficial Temporal Space|Right Deep Temporal Space|" & _
        "Left Deep Temporal Space|Right Pterygomandibular Space|Left Pterygomandibular Space|" & _
        "Right Lateral Pharyngeal Space|Left Lateral Pharyngeal Space|Right Retropharyngeal Space|" & _
        "Left Retropharyngeal Space|Cervical Spaces|Right Palatal Swelling|Left Palatal Swelling|" & _
        "Right Canine (Infraorbital) Space|Left Canine (Infraorbital) Space|Right Periorbital Space|" & _
        "Left Periorbital Space|Base of Upper Lip|Right Cavernous Sinus Thrombosis|Left Cavernous Sinus Thrombosis|"

End Sub
Private Property Get clsISection_DentalEncounterID() As Long
    clsISection_DentalEncounterID = mlngDentalEncounterId
End Property

Private Function clsISection_GetNote() As String
    Dim strNote As String
    Dim arrValues() As String
    Dim strTemp As String
    Dim objItem As clsExEndoAdFnd
    Dim X As Integer
    
    If mstrSwelling = "" Then
        clsISection_GetNote = Empty
        Exit Function
    End If
    strTemp = Get_MultiSelect_DisplayString(mstrSwelling)
    
    If strTemp <> "" Then
        strNote = " Endodontics Template: " & vbCrLf & vbCrLf & vbTab & _
        "Location of Swelling = " & strTemp & vbCrLf
    End If

    
    strTemp = Empty
    If mcolItems Is Nothing = False Then
        For Each objItem In mcolItems
            strTemp = strTemp & vbTab & vbTab & objItem.GetNote & vbCrLf
        Next
    End If
    
    If strTemp <> "" Then
        strNote = strNote & vbTab & "Additional Findings:" & vbCrLf & strTemp
    End If
    clsISection_GetNote = strNote
End Function

Private Function clsISection_GetRTFNote() As String
    Dim strTemp As String
    Dim strNote As String
    Dim objItem As clsExEndoAdFnd
    Dim X As Integer
        
    If mstrSwelling = "" Then
        clsISection_GetRTFNote = Empty
        Exit Function
    End If
    strTemp = Get_MultiSelect_DisplayString(mstrSwelling)
    
    If strTemp <> "" Then
        strNote = crtfStartControlName & "Location of Swelling: " & _
        strTemp & ";" & crtfEndControlName
    End If

    
    strTemp = Empty
    If mcolItems Is Nothing = False Then
        For Each objItem In mcolItems
            strTemp = strTemp & objItem.GetNote
        Next
    End If
    
    If Trim(strTemp) <> "" Then
        strNote = strNote & crtfStartControlName & "Additional Findings: " & _
             strTemp & crtfEndControlName
    End If
    
    Set objItem = Nothing
    
    clsISection_GetRTFNote = strNote
End Function

Private Function clsISection_LoadSection(ByVal PatientID As Long, ByVal DentalEncounterID As Long) As Boolean
    Dim strSql As String
    Dim blnReturn As Boolean
    Dim objRS As ADODB.Recordset
    
    On Error GoTo ErrHandler
    
    mlngPatUnitId = PatientID
    mlngDentalEncounterId = DentalEncounterID
    
    Set objRS = DB_Load_Swelling(mlngDentalEncounterId)
    
    With objRS
        If .EOF Then
            mlngDirtyFlag = dfAdd
            mlngFindId = 0
            Initialize_Values
        
        Else
            mlngDirtyFlag = dfUpdate
            mlngFindId = .Fields("EX_ENDO_FIND_ID")
            mstrSwelling = .Fields("SWELLING") & ""
            
            '''Load_Pain_Characteristics .Fields("EX_ENDO_SUBJ_ID")
        End If
    End With
    
    Load_Additional_Findings
    
    clsISection_LoadSection = True
    Exit Function
ErrHandler:
    clsISection_LoadSection = False
End Function
Public Property Get Item(vntIndexKey As Variant) As clsExEndoAdFnd
  Set Item = mcolItems(vntIndexKey)
End Property
Public Property Get Count() As Long
    If mcolItems Is Nothing = False Then
        Count = mcolItems.Count
    End If
End Property
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = mcolItems.[_NewEnum]
End Property
Private Property Get clsISection_PatientID() As Long
    clsISection_PatientID = mlngPatUnitId
End Property
Public Property Get Additional_Findings(ByVal intIndex As Integer) As clsExEndoAdFnd
    Dim objItem As clsExEndoAdFnd
    
    If intIndex > mcolItems.Count Then
        Set objItem = New clsExEndoAdFnd
        objItem.DentalEncounterID = mlngDentalEncounterId
        mcolItems.Add objItem
    Else
        Set objItem = mcolItems(intIndex)
    End If
    
    Set Additional_Findings = objItem
        
    Set objItem = Nothing
End Property
Private Function clsISection_SaveSection() As Boolean
    Dim strSql As String
    Dim blnReturn As Boolean
    Dim objItem As clsExEndoAdFnd
    
    strSql = Empty
    Select Case mlngDirtyFlag
        Case dfNoAction
            blnReturn = True
        Case dfAdd
        
            mlngFindId = goconn.GetNextID(esTEMPL_EXAM_ENDO_FINDINGS_SEQ)
'            strSql = "SELECT TEMPL_EXAM_ENDO_FINDINGS_seq.NEXTVAL From Dual"
'            Set objRS = goconn.OpenRecordset(strSql)
'            mlngFindId = objRS.Fields("NEXTVAL")
            
            strSql = "INSERT INTO TEMPL_EXAM_ENDO_FINDINGS " & _
                     "(EX_ENDO_FIND_ID, DENTAL_ENCOUNTER_ID, " & _
                     " SWELLING) " & _
                     "VALUES (" & mlngFindId & ", " & mlngDentalEncounterId & ", '" & _
                    Scrub_Apostrophe(mstrSwelling) & "')"

        Case dfUpdate
            strSql = "UPDATE TEMPL_EXAM_ENDO_FINDINGS " & _
                     " SET DENTAL_ENCOUNTER_ID = " & mlngDentalEncounterId & ", " & _
                     " SWELLING = '" & Scrub_Apostrophe(mstrSwelling) & "' " & _
                     " WHERE EX_ENDO_FIND_ID = " & mlngFindId
        Case dfdelete
            strSql = "DELETE FROM TEMPL_EXAM_ENDO_FINDINGS WHERE " & _
                     " EX_ENDO_FIND_ID = " & mlngFindId
    End Select
    
    If strSql <> "" Then
        goconn.ExecuteSQL (strSql)
    End If
    mlngDirtyFlag = dfNoAction
    
    If mcolItems.Count > 0 Then
        For Each objItem In mcolItems
            objItem.DentalEncounterID = mlngDentalEncounterId
            objItem.SaveSection
        Next
    End If
    
End Function

Private Function DB_Load_Swelling(ByVal lngDentalEncId As Long) As ADODB.Recordset
    Dim strSql As String
    
    strSql = "Select EX_ENDO_FIND_ID, SWELLING from TEMPL_EXAM_ENDO_FINDINGS " & _
            "Where DENTAL_ENCOUNTER_ID = " & lngDentalEncId
            
    Set DB_Load_Swelling = goconn.OpenRecordset(strSql)
    
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
'**************************************************
'Sub: Load_Additional_Findings
'Author: KDunne
'Date: 11/8/2001
'Purpose: Loads collction of Additional Findings items.
'***************************************************
Private Sub Load_Additional_Findings()
    Dim strSql As String
    Dim objRS As ADODB.Recordset
    Dim objAddFind As clsExEndoAdFnd
    
    strSql = "Select EX_ENDO_ADD_FIND_ID, TOOTH_NUM, CANALS, ADD_FINDING " & _
             " FROM TEMPL_EXAM_ENDO_ADD_FINDINGS WHERE DENTAL_ENCOUNTER_ID = " & _
             mlngDentalEncounterId
             
    Set objRS = goconn.OpenRecordset(strSql)
    
    Set mcolItems = New Collection
    
    With objRS
        Do While .EOF = False
            Set objAddFind = New clsExEndoAdFnd
            objAddFind.AdditionalFindingsID = .Fields("EX_ENDO_ADD_FIND_ID")
            objAddFind.ToothNumber = .Fields("TOOTH_NUM")
            objAddFind.Canals = .Fields("CANALS")
            objAddFind.DentalEncounterID = mlngDentalEncounterId
            objAddFind.AdditionalFinding = .Fields("ADD_FINDING")
            objAddFind.DirtyFlag = dfNoAction
            .MoveNext
            
            mcolItems.Add objAddFind
        Loop

    End With
    
    Set objRS = Nothing
End Sub

'**************************************************************************
'Function:Remove
'Author:KDunne
'Date:8/5/2002
'Purpose:Removes object from internal collection.
'***************************************************************************
Public Sub Remove(ByVal Index As Long)
       
    Dim objAddFind As clsExEndoAdFnd
    
    If Index > mcolItems.Count Then Exit Sub
    
    'KDunne 11/4/2003
    'SCR 44672 - When removing an item from the collection, we must also
    '            remove the item from the database.
    Set objAddFind = mcolItems.Item(Index)
    
    objAddFind.DirtyFlag = dfdelete
    objAddFind.SaveSection
    
    mcolItems.Remove Index
    
    
End Sub
