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
Attribute VB_Name = "clsTreatFixPro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Implements clsISection

Private mvarPatientID As Long
Private mvarDentalEncounterID As Long

Private dfDirtyFlag As DirtyFlagEnum
'local variable(s) to hold property value(s)
Private mstrTeeth_Number As String
Private mintTeeth_Stained As Integer
Private mintTeeth_Glazed As Integer
Private mintPost_Op As Integer
Private mintOral_Hygiene As Integer
Private mstrTreatment_Plan As String
Private mstrComments As String
Private mstrSupporting_Procedures As String
Private mstrTreatment_Comments As String
Public Property Let Supporting_Procedures(ByVal vData As String)
    mstrSupporting_Procedures = vData
    UpdateDirtyFlag
End Property

Public Property Get Supporting_Procedures() As String

    If mstrSupporting_Procedures = "" Then
        mstrSupporting_Procedures = gcSupportingProcs
    End If
    
    Supporting_Procedures = mstrSupporting_Procedures

End Property
Public Property Let Treatment_Comments(ByVal vData As String)
    mstrTreatment_Comments = vData
    UpdateDirtyFlag
End Property


Public Property Get Treatment_Comments() As String
    Treatment_Comments = mstrTreatment_Comments
End Property

Public Property Let Comments(ByVal vData As String)
    mstrComments = vData
    UpdateDirtyFlag
End Property


Public Property Get Comments() As String
    Comments = mstrComments
End Property

Public Property Let Treatment_Plan(ByVal vData As String)
    mstrTreatment_Plan = vData
    UpdateDirtyFlag
End Property

Public Property Get Treatment_Plan() As String
    Treatment_Plan = mstrTreatment_Plan
End Property

Public Property Let Oral_Hygiene(ByVal vData As Integer)
    mintOral_Hygiene = vData
    UpdateDirtyFlag
End Property

Public Property Get Oral_Hygiene() As Integer
    Oral_Hygiene = mintOral_Hygiene
End Property

Public Property Let Post_Op(ByVal vData As Integer)
    mintPost_Op = vData
    UpdateDirtyFlag
End Property

Public Property Get Post_Op() As Integer
    Post_Op = mintPost_Op
End Property

Public Property Let Teeth_Glazed(ByVal vData As Integer)
    mintTeeth_Glazed = vData
    UpdateDirtyFlag
End Property

Public Property Get Teeth_Glazed() As Integer
    Teeth_Glazed = mintTeeth_Glazed
End Property

Public Property Let Teeth_Stained(ByVal vData As Integer)
    mintTeeth_Stained = vData
    UpdateDirtyFlag
End Property

Public Property Get Teeth_Stained() As Integer
    Teeth_Stained = mintTeeth_Stained
End Property

Public Property Let Teeth_Number(ByVal vData As String)
    mstrTeeth_Number = vData
    UpdateDirtyFlag
End Property

Public Property Get Teeth_Number() As String
    Teeth_Number = mstrTeeth_Number
End Property


Private Property Get clsISection_DentalEncounterID() As Long
    
    clsISection_DentalEncounterID = mvarDentalEncounterID

End Property


Private Property Get clsISection_PatientID() As Long
    
    clsISection_PatientID = mvarPatientID

End Property

Private Function clsISection_GetRTFNote() As String

Dim sNote As String
Dim sTemp As String
Dim sTxt As String
Dim objToothCtrl As NotesPlusUtilities.CToothControl

Const sPROC_NAME As String = "clsTreatFixPro.clsTreatFixPro.GetRTFNote"
        
    If mstrTeeth_Number = "" And _
    mintTeeth_Stained = 0 And _
    mintTeeth_Glazed = 0 And _
    mintPost_Op = 0 And _
    mintOral_Hygiene = 0 And _
    mstrTreatment_Plan = "" And _
    mstrSupporting_Procedures = "" And _
    mstrComments = "" Then
        'No need to update note..
        sNote = ""
        clsISection_GetRTFNote = sNote
        Exit Function
    End If
    
    
    If mstrTeeth_Number <> "" Then
        Set objToothCtrl = New NotesPlusUtilities.CToothControl
        sNote = sNote + crtfStartControlName & "Teeth Numbers: " + objToothCtrl.BuildDisplayString(mstrTeeth_Number) + crtfEndControlName
    End If
    If mintTeeth_Stained = -1 Or mintTeeth_Glazed = -1 Then
        sNote = sNote + crtfStartControlName & "Teeth Color: "
        If mintTeeth_Stained = -1 Then sNote = sNote + "Teeth Stained; "
        If mintTeeth_Glazed = -1 Then sNote = sNote + "Teeth Glazed; "
        'Remove the last "; "
        sNote = Left(sNote, Len(sNote) - Len("; "))
        sNote = sNote + crtfEndControlName
    End If
    
    If mintPost_Op = -1 Or mintOral_Hygiene = -1 Then
            If mintPost_Op = -1 Then sNote = sNote + crtfStartControlName & "Post Op Instructions" + crtfEndControlName
            If mintOral_Hygiene = -1 Then sNote = sNote + crtfStartControlName & "Oral Hygiene Instructions" + crtfEndControlName
    End If
    
    If mstrTreatment_Plan <> "" Or mstrTreatment_Comments <> "" Then
            sNote = sNote + crtfStartControlName + "Preparation/Provisional Prostheses: "
            If mstrTreatment_Plan <> "" Then
                sNote = sNote + Replace(mstrTreatment_Plan, "|", "; ")
            End If
            If mstrTreatment_Comments <> "" Then
                sNote = sNote & "Treatment Plan Comments: " + mstrTreatment_Comments + "; "
            End If
            'Remove the last "; "
            sNote = Left(sNote, Len(sNote) - Len("; "))
            sNote = sNote + crtfEndControlName
            sTxt = ""
    End If
    
    
    If mstrComments <> "" Then
            sNote = sNote + crtfStartControlName & "Fixed Pro Comments: " + mstrComments + crtfEndControlName
    End If
    
    sTxt = SupportingProcsDisplay
    If sTxt <> "" Then
        sNote = sNote & crtfStartControlName & "Supporting Procedures: " & sTxt + crtfEndControlName
    End If

    
    clsISection_GetRTFNote = sNote
Exit Function

ErrHnd:
    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & "." & sPROC_NAME, App.Title, vbCritical


End Function

'*******************************************************************
'Function: GetNote
'Author: Sean McAvoy
'Date: 01/06/02
'Returns: A string that is displayed in the narrative summary for this encounter
'Purpose: To compare the values from this encounter and the previous encounter and
'return a text note.
'*******************************************************************
Private Function clsISection_GetNote() As String

Dim sNote As String
Dim sTxt As String

Const sPROC_NAME As String = "clsTreatFixPro.clsTreatFixPro.GetNote"
        
    If mstrTeeth_Number = "" And _
    mintTeeth_Stained = 0 And _
    mintTeeth_Glazed = 0 And _
    mintPost_Op = 0 And _
    mintOral_Hygiene = 0 And _
    mstrTreatment_Plan = "" And _
    mstrSupporting_Procedures = "" And _
    mstrComments = "" Then
        'No need to update note..
        sNote = ""
        clsISection_GetNote = sNote
        Exit Function
    End If

    sNote = vbCrLf + "Treatment - Fixed Pros" + vbCrLf
    sNote = sNote + vbTab + "Teeth Numbers: " + Replace(mstrTeeth_Number, "|", ", ") + vbCrLf
    If mintTeeth_Stained = -1 Then sNote = sNote + vbTab + "Teeth Stained" + vbCrLf
    If mintTeeth_Glazed = -1 Then sNote = sNote + vbTab + "Teeth Glazed" + vbCrLf
    If mintPost_Op = -1 Then sNote = sNote + vbTab + "Post Op Instructions" + vbCrLf
    If mintOral_Hygiene = -1 Then sNote = sNote + vbTab + "Oral Hygiene Instructions" + vbCrLf
    If mstrTreatment_Plan <> "" Then sNote = sNote + vbTab + "Treatment Plan: " + mstrTreatment_Plan + vbCrLf
    If mstrComments <> "" Then sNote = sNote + vbTab + "Fixed Pros Comments: " + mstrComments + vbCrLf
        sTxt = SupportingProcsDisplay
    If sTxt <> "" Then
        sNote = sNote & vbTab & "Supporting Procedures: " & sTxt & vbCrLf
    End If

    clsISection_GetNote = sNote
Exit Function

ErrHnd:
    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & "." & sPROC_NAME, App.Title, vbCritical

End Function
'*******************************************************************
'Function: SupportingProcsDisplay
'Author: Sean McAvoy
'Date: 1/2/02
'Returns: String
'Purpose: To get the display string for supporting procedures
'*******************************************************************
Private Function SupportingProcsDisplay() As String

Dim oItems As New CMultiListItems

    oItems.ListType = Standard
    oItems.LoadFromValueString mstrSupporting_Procedures
    SupportingProcsDisplay = oItems.Value_String_Display
Set oItems = Nothing
End Function

'*******************************************************************
'Function: SaveSection
'Author: Sean McAvoy
'Date: 01/06/02
'Returns: True/False
'Purpose: Saves the data to the table
'*******************************************************************
Private Function clsISection_SaveSection() As Boolean

Dim sSQL As String

Const sPROC_NAME As String = "clsTreatFixPro.SaveSection"
Dim mstrComments_TMP As String
Dim mstrTreatment_Comments_TMP As String

On Error GoTo ErrHnd

    If dfDirtyFlag = dfNoAction Then
        clsISection_SaveSection = True
        Exit Function
    End If
    If mvarDentalEncounterID = 0 Then
        clsISection_SaveSection = False
        Exit Function
    End If
    
    mstrComments_TMP = RQ(mstrComments)
    mstrTreatment_Comments_TMP = RQ(mstrTreatment_Comments)
                                                                               
    If dfDirtyFlag = dfAdd Then
        'The record does not exist, so create an insert statement
        sSQL = "INSERT INTO TEMPL_TREAT_FIXPRO (" _
            & " DENTAL_ENCOUNTER_ID, TEETH_NUMBER, TEETH_STAINED, TEETH_GLAZED," _
            & " POST_OP, ORAL_HYGIENE, TREATMENT_PLAN, COMMENTS, SUPPORTING_PROCEDURES, TREATMENT_PLAN_COMMENTS)" _
            & " VALUES" _
            & " (" & CStr(mvarDentalEncounterID) & ", " _
            & " '" & mstrTeeth_Number & "', " _
           & CStr(mintTeeth_Stained) & ", " _
           & CStr(mintTeeth_Glazed) & ", " _
           & CStr(mintPost_Op) & ", " _
           & CStr(mintOral_Hygiene) & ", " _
            & " '" & mstrTreatment_Plan & "', " _
            & " '" & mstrComments_TMP & "', " _
            & " '" & mstrSupporting_Procedures & "', " _
            & " '" & mstrTreatment_Comments_TMP & "') " _

    ElseIf dfDirtyFlag = dfUpdate Then
        'The record exists, so create an update statement
        sSQL = "UPDATE TEMPL_TREAT_FIXPRO SET " + _
        "TEETH_NUMBER = '" + mstrTeeth_Number + "', " + _
        "TEETH_STAINED = " + CStr(mintTeeth_Stained) + ", " + _
        "TEETH_GLAZED = " + CStr(mintTeeth_Glazed) + ", " + _
        "POST_OP = " + CStr(mintPost_Op) + ", " + _
        "ORAL_HYGIENE = " + CStr(mintOral_Hygiene) + ", " + _
        "TREATMENT_PLAN = '" + mstrTreatment_Plan + "', " + _
        "COMMENTS = '" + mstrComments_TMP + "',  " + _
        "SUPPORTING_PROCEDURES = '" + mstrSupporting_Procedures + "', " + _
        "TREATMENT_PLAN_COMMENTS = '" + mstrTreatment_Comments_TMP + "'  " + _
        "Where (DENTAL_ENCOUNTER_ID = " & CStr(mvarDentalEncounterID) & ")"


    ElseIf dfDirtyFlag = dfdelete Then
        clsISection_SaveSection = False
        Exit Function
    End If
    
    goconn.ExecuteSQL sSQL
    
    dfDirtyFlag = dfNoAction
    
    clsISection_SaveSection = True

Exit Function

ErrHnd:
    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & "." & sPROC_NAME, App.Title, vbCritical

End Function


'*******************************************************************
'Function: LoadSection
'Author: Sean McAvoy
'Date: 01/06/02
'Returns: True/False
'Purpose: This loads the data from the db
'*******************************************************************
Private Function clsISection_LoadSection(ByVal PatientID As Long, ByVal DentalEncounterID As Long) As Boolean

Dim sSQL As String, rs As Recordset, rsPrev As Recordset
Const sPROC_NAME = "clsTreatFixPro.LoadSection"

On Error GoTo ErrHnd

    mvarPatientID = PatientID
    mvarDentalEncounterID = DentalEncounterID
    
'    TEETH_NUMBER, TEETH_STAINED, TEETH_GLAZED, POST_OP, ORAL_HYGIENE, TREATMENT_PLAN, COMMENTS
    
    'Load the values
    sSQL = "SELECT  TEETH_NUMBER, TEETH_STAINED, TEETH_GLAZED, POST_OP, ORAL_HYGIENE, " + _
    "TREATMENT_PLAN, COMMENTS, SUPPORTING_PROCEDURES FROM TEMPL_TREAT_FIXPRO Where " + _
    "DENTAL_ENCOUNTER_ID = " + CStr(mvarDentalEncounterID)
    
    Set rs = goconn.OpenRecordset(sSQL)
    
    If rs.EOF Then
        dfDirtyFlag = dfAdd
        'Init the values here to ""
        mstrTeeth_Number = ""
        mintTeeth_Stained = 0
        mintTeeth_Glazed = 0
        mintPost_Op = 0
        mintOral_Hygiene = 0
        mstrTreatment_Plan = ""
        mstrComments = ""
        mstrSupporting_Procedures = ""
        mstrTreatment_Comments = ""
    Else
        dfDirtyFlag = dfNoAction
        'Set the values from the db including "" to allow for nulls
        mstrTeeth_Number = rs("TEETH_NUMBER") & ""
        mintTeeth_Stained = Val(rs("TEETH_STAINED") & "")
        mintTeeth_Glazed = Val(rs("TEETH_GLAZED") & "")
        mintPost_Op = Val(rs("POST_OP") & "")
        mintOral_Hygiene = Val(rs("ORAL_HYGIENE") & "")
        mstrTreatment_Plan = rs("TREATMENT_PLAN") & ""
        mstrComments = rs("COMMENTS") & ""
        
        rs.Close
        
        sSQL = "SELECT SUPPORTING_PROCEDURES, TREATMENT_PLAN_COMMENTS FROM TEMPL_TREAT_FIXPRO Where " + _
        "DENTAL_ENCOUNTER_ID = " + CStr(mvarDentalEncounterID)
        
        Set rs = goconn.OpenRecordset(sSQL)
        
        If rs.EOF Then
            mstrSupporting_Procedures = ""
            mstrTreatment_Comments = ""
        Else
            mstrSupporting_Procedures = rs("SUPPORTING_PROCEDURES") & ""
            mstrTreatment_Comments = rs("TREATMENT_PLAN_COMMENTS") & ""
        End If
        rs.Close
    End If
        Set rs = Nothing
        
        clsISection_LoadSection = True
    
Exit Function

ErrHnd:
    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & "." & sPROC_NAME, App.Title, vbCritical
    
End Function

Private Function UpdateDirtyFlag()

    If dfDirtyFlag = dfNoAction Then
        dfDirtyFlag = dfUpdate
    End If
    
End Function

Private Sub Class_Initialize()

    dfDirtyFlag = dfNoAction
    mstrTeeth_Number = ""
    mintTeeth_Stained = 0
    mintTeeth_Glazed = 0
    mintPost_Op = 0
    mintOral_Hygiene = 0
    mstrTreatment_Plan = ""
    mstrComments = ""
    mstrSupporting_Procedures = ""
    mstrTreatment_Comments = ""

End Sub


