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
Attribute VB_Name = "clsPrognosis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Implements clsISection

'Used to determine if the data should be inserted or updated when modified.
Private bSavedToDB As Boolean

Private mvarPatientID As Long 'local copy
Private mvarDentalEncounterID As Long 'local copy
Private dfDirtyFlag As DirtyFlagEnum

'Private mvarShortTerm As String 'local copy
'Private mvarLongTerm As String 'local copy
'Private mvarFairTeeth As String 'local copy
'Private mvarQuestionableTeeth As String 'local copy
'Private mvarPoorTeeth As String 'local copy
'Private mvarHopelessTeeth As String 'local copy
'Private mvarGoodTeeth As String
'Private mvarExcellentTeeth As String

Public Property Let HopelessTeeth(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.HopelessTeeth = 5
    mvarHopelessTeeth = vData
    UpdateDirtyFlag
End Property


Public Property Get HopelessTeeth() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HopelessTeeth
    HopelessTeeth = mvarHopelessTeeth
End Property



Public Property Let PoorTeeth(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PoorTeeth = 5
    mvarPoorTeeth = vData
    UpdateDirtyFlag
End Property


Public Property Get PoorTeeth() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PoorTeeth
    PoorTeeth = mvarPoorTeeth
End Property



Public Property Let QuestionableTeeth(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.QuestionableTeeth = 5
    mvarQuestionableTeeth = vData
    UpdateDirtyFlag
End Property


Public Property Get QuestionableTeeth() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.QuestionableTeeth
    QuestionableTeeth = mvarQuestionableTeeth
End Property



Public Property Let FairTeeth(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FairTeeth = 5
    mvarFairTeeth = vData
    UpdateDirtyFlag
End Property

Public Property Let ExcellentTeeth(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FairTeeth = 5
    mvarExcellentTeeth = vData
    UpdateDirtyFlag
End Property

Public Property Let GoodTeeth(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FairTeeth = 5
    mvarGoodTeeth = vData
    UpdateDirtyFlag
End Property

Public Property Get FairTeeth() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FairTeeth
    FairTeeth = mvarFairTeeth
End Property


Public Property Get ExcellentTeeth() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FairTeeth
    ExcellentTeeth = mvarExcellentTeeth
End Property

Public Property Get GoodTeeth() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FairTeeth
    GoodTeeth = mvarGoodTeeth
End Property

Public Property Let LongTerm(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LongTerm = 5
    mvarLongTerm = Trim(vData)
    UpdateDirtyFlag
End Property


Public Property Get LongTerm() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LongTerm
    LongTerm = mvarLongTerm
End Property



Public Property Let ShortTerm(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ShortTerm = 5
    mvarShortTerm = Trim(vData)
    UpdateDirtyFlag
End Property


Public Property Get ShortTerm() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ShortTerm
    ShortTerm = mvarShortTerm
End Property













Private Function UpdateDirtyFlag()

    If bSavedToDB = False Then
        dfDirtyFlag = dfAdd
    ElseIf dfDirtyFlag = dfNoAction Then
        dfDirtyFlag = dfUpdate
    End If
    
End Function

Private Sub Class_Initialize()

    dfDirtyFlag = dfNoAction
    
End Sub




Private Property Get clsISection_DentalEncounterID() As Long

    clsISection_DentalEncounterID = mvarDentalEncounterID

End Property


'*******************************************************************
'Function: clsISection_GetNote
'Author: Scott Foster
'Date: 10-22-01
'Returns: Text note
'Purpose: Gets the note for the narrative summary
'*******************************************************************
Private Function clsISection_GetNote() As String

Dim sTemp As String
    
    If mvarShortTerm <> "" Then
        sTemp = sTemp & vbTab & "Short term: " & mvarShortTerm & vbCrLf
    End If
    
    If mvarLongTerm <> "" Then
        sTemp = sTemp & vbTab & "Long term: " & mvarLongTerm & vbCrLf
    End If
    
    If mvarFairTeeth <> "" Then
        sTemp = sTemp & vbTab & "Fair teeth: " & Replace(mvarFairTeeth, "|", ", ") & vbCrLf
    End If
    
    If mvarQuestionableTeeth <> "" Then
        sTemp = sTemp & vbTab & "Questionable teeth: " & Replace(mvarQuestionableTeeth, "|", ", ") & vbCrLf
    End If
    
    If mvarPoorTeeth <> "" Then
        sTemp = sTemp & vbTab & "Poor teeth: " & Replace(mvarPoorTeeth, "|", ", ") & vbCrLf
    End If
    
    If mvarHopelessTeeth <> "" Then
        sTemp = sTemp & vbTab & "Hopeless teeth: " & Replace(mvarHopelessTeeth, "|", ", ") & vbCrLf
    End If

    If sTemp <> "" Then
            sTemp = "Prognosis: " & vbCrLf & sTemp
    End If
    
    clsISection_GetNote = sTemp
    
End Function


Private Function clsISection_GetRTFNote() As String

Dim sTemp As String
Dim objToothCtrl As CToothControl

    Set objToothCtrl = New CToothControl
    
    If mvarShortTerm <> "" Then
        sTemp = sTemp & "Short term: " & mvarShortTerm & "; "
    End If
    
    If mvarLongTerm <> "" Then
        sTemp = sTemp & "Long term: " & mvarLongTerm & "; "
    End If
    
    If mvarFairTeeth <> "" Then
        sTemp = sTemp & "Fair teeth: " & objToothCtrl.BuildDisplayString(mvarFairTeeth) & "; "
    End If
    
    If mvarQuestionableTeeth <> "" Then
        sTemp = sTemp & "Questionable teeth: " & objToothCtrl.BuildDisplayString(mvarQuestionableTeeth) & "; "
    End If
    
    If mvarPoorTeeth <> "" Then
        sTemp = sTemp & "Poor teeth: " & objToothCtrl.BuildDisplayString(mvarPoorTeeth) & "; "
    End If
    
    If mvarHopelessTeeth <> "" Then
        sTemp = sTemp & "Hopeless teeth: " & objToothCtrl.BuildDisplayString(mvarHopelessTeeth) & "; "
    End If

    If sTemp <> "" Then
    
        'Remove the last "; "
        sTemp = Left(sTemp, Len(sTemp) - Len("; "))
        
        sTemp = crtfStartControlName & "Prognosis: " & sTemp & crtfEndControlName
    End If
    
    clsISection_GetRTFNote = sTemp
    
End Function


'*******************************************************************
'Function: clsISection_LoadSection
'Author: Scott Foster
'Date: 10-22-01
'Returns: True/False
'Purpose: This loads the data from the db
'*******************************************************************
Private Function clsISection_LoadSection(ByVal PatientID As Long, ByVal DentalEncounterID As Long) As Boolean

Dim sSQL As String, rs As Recordset, rsPrev As Recordset
Const sPROC_NAME = "LoadSection"

On Error GoTo ErrHnd
    
    'So this can load in health history dental dataviews.
    OpenDB
    'Init the XMLPersist class and load the XML data
    Dim XMLPersist As New clsXMLPersist
    XMLPersist.ExamXML.Load DentalEncounterID

    mvarPatientID = PatientID
    mvarDentalEncounterID = DentalEncounterID

    If gbUseStoredProc = True Then
        Dim oCmd As New CHCSII_Command.CHCSII_Cmd
        
        oCmd.SPName = "Dent_Notes_Prognosis_Load"

        Call oCmd.AddParam(False, False, "textString", mvarDentalEncounterID)
        Call oCmd.AddParam(True, True, "", "")
        
        Set rs = goconn.OpenRecordsetSP(oCmd)
    
    Else
    
        'Load the values
        sSQL = "SELECT DENTAL_ENCOUNTER_ID, SHORT_TERM, LONG_TERM, FAIR_TEETH," _
            & " QUESTIONABLE_TEETH, POOR_TEETH, HOPELESS_TEETH," _
            & " EXCELLENT_TEETH, GOOD_TEETH" _
            & " From TEMPL_PERIO_PROGNOSIS" _
            & " Where (DENTAL_ENCOUNTER_ID = " & mvarDentalEncounterID & ")"
    
        Set rs = goconn.OpenRecordset(sSQL)
    End If
    
    If rs.EOF Then
        bSavedToDB = False
        dfDirtyFlag = dfAdd
        'Init the values here to ""
        mvarShortTerm = ""
        mvarLongTerm = ""
        mvarFairTeeth = ""
        mvarQuestionableTeeth = ""
        mvarPoorTeeth = ""
        mvarHopelessTeeth = ""
        mvarGoodTeeth = ""
        mvarExcellentTeeth = ""
    Else
        bSavedToDB = True
        dfDirtyFlag = dfNoAction
        'Set the values from the db including "" to allow for nulls
        mvarShortTerm = rs("Short_Term") & ""
        mvarLongTerm = rs("Long_Term") & ""
        mvarFairTeeth = rs("Fair_Teeth") & ""
        mvarQuestionableTeeth = rs("Questionable_Teeth") & ""
        mvarPoorTeeth = rs("Poor_Teeth") & ""
        mvarHopelessTeeth = rs("Hopeless_Teeth") & ""
        mvarGoodTeeth = rs.Fields("Good_Teeth").Value & ""
        mvarExcellentTeeth = rs.Fields("Excellent_Teeth").Value & ""
    End If
    
    rs.Close
    Set rs = Nothing
    
    clsISection_LoadSection = True
    
Exit Function

ErrHnd:
    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & "." & sPROC_NAME, App.Title, vbCritical

End Function


Private Property Get clsISection_PatientID() As Long

    clsISection_PatientID = mvarPatientID

End Property


'*******************************************************************
'Function: clsISection_SaveSection
'Author: Scott Foster
'Date: 10-22-01
'Returns: True/False
'Purpose: Saves the data to the table
'*******************************************************************
Private Function clsISection_SaveSection() As Boolean

Dim sSQL As String

Const sPROC_NAME As String = "SaveSection"

On Error GoTo ErrHnd

    If dfDirtyFlag = dfNoAction Then
        clsISection_SaveSection = True
        Exit Function
        
    ElseIf dfDirtyFlag = dfAdd Then
        'The record does not exist, so create an insert statement
        sSQL = "INSERT INTO TEMPL_PERIO_PROGNOSIS" _
            & " (DENTAL_ENCOUNTER_ID, SHORT_TERM, LONG_TERM," _
            & " FAIR_TEETH, QUESTIONABLE_TEETH, POOR_TEETH," _
            & " HOPELESS_TEETH, EXCELLENT_TEETH, GOOD_TEETH)" _
            & " VALUES (" _
            & mvarDentalEncounterID & ", '" & mvarShortTerm & "', '" & mvarLongTerm & "', '" _
            & mvarFairTeeth & "', '" & mvarQuestionableTeeth & "', '" _
            & mvarPoorTeeth & "', '" _
            & mvarHopelessTeeth & "', '" & mvarGoodTeeth & "', '" & mvarExcellentTeeth & "')"
            
    ElseIf dfDirtyFlag = dfUpdate Then
        'The record exists, so create an update statement
        sSQL = "UPDATE TEMPL_PERIO_PROGNOSIS" _
            & " SET SHORT_TERM = '" & mvarShortTerm & "'," _
            & " LONG_TERM = '" & mvarLongTerm & "'," _
            & " FAIR_TEETH = '" & mvarFairTeeth & "'," _
            & " QUESTIONABLE_TEETH = '" & QuestionableTeeth & "'," _
            & " POOR_TEETH = '" & mvarPoorTeeth & "'," _
            & " HOPELESS_TEETH = '" & mvarHopelessTeeth & "'," _
            & " GOOD_TEETH = '" & mvarGoodTeeth & "'," _
            & " EXCELLENT_TEETH = '" & mvarExcellentTeeth & "'" _
            & " Where (DENTAL_ENCOUNTER_ID = " & mvarDentalEncounterID & ")"

    End If
    
    goconn.ExecuteSQL sSQL
    
    bSavedToDB = True
    dfDirtyFlag = dfNoAction
    
    clsISection_SaveSection = True

Exit Function

ErrHnd:
    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & "." & sPROC_NAME, App.Title, vbCritical

End Function


