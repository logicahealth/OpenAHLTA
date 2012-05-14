VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTreatRP_RPD"
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


'*****************************************************************************************
'******** Declares Used by Tab:Treatment/Removable Pros/Removable Partial Denture ********
'*****************************************************************************************
Private mstrTeeth_Prepped As String
Private mstrTeeth_Replaced As String
Private mintPolished_Preps As Integer
Private mintTopical_Fluoride As Integer
Private mintFramework_Tryin As Integer
Private mintWaxrims As Integer
Private mintAltered_Cast_Impre As Integer
Private mintAdjust_RPD As Integer
Private mintDeliver_RPD As Integer
Private mintPost_Insertion As Integer
Private mintPost_Delivery_Adj_RPD As Integer
Private mstrRPD_Comments As String
Private mstrSupportingProcs As String


Public Property Let RPD_Comments(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RPD_Comments = 5
    mstrRPD_Comments = vData
    UpdateDirtyFlag
End Property


Public Property Get RPD_Comments() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RPD_Comments
    RPD_Comments = mstrRPD_Comments
End Property



Public Property Let Post_Delivery_Adj_RPD(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Post_Delivery_Adj_RPD = 5
    mintPost_Delivery_Adj_RPD = vData
    UpdateDirtyFlag
End Property


Public Property Get Post_Delivery_Adj_RPD() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Post_Delivery_Adj_RPD
    Post_Delivery_Adj_RPD = mintPost_Delivery_Adj_RPD
End Property



Public Property Let Post_Insertion(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Post_Insertion = 5
    mintPost_Insertion = vData
    UpdateDirtyFlag
End Property


Public Property Get Post_Insertion() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Post_Insertion
    Post_Insertion = mintPost_Insertion
End Property



Public Property Let Deliver_RPD(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Deliver_RPD = 5
    mintDeliver_RPD = vData
    UpdateDirtyFlag
End Property


Public Property Get Deliver_RPD() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Deliver_RPD
    Deliver_RPD = mintDeliver_RPD
End Property



Public Property Let Adjust_RPD(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Adjust_RPD = 5
    mintAdjust_RPD = vData
    UpdateDirtyFlag
End Property


Public Property Get Adjust_RPD() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Adjust_RPD
    Adjust_RPD = mintAdjust_RPD
End Property



Public Property Let Altered_Cast_Impre(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Altered_Cast_Impre = 5
    mintAltered_Cast_Impre = vData
    UpdateDirtyFlag
End Property


Public Property Get Altered_Cast_Impre() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Altered_Cast_Impre
    Altered_Cast_Impre = mintAltered_Cast_Impre
End Property



Public Property Let Waxrims(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Waxrims = 5
    mintWaxrims = vData
    UpdateDirtyFlag
End Property


Public Property Get Waxrims() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Waxrims
    Waxrims = mintWaxrims
End Property



Public Property Let Framework_Tryin(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Framework_Tryin = 5
    mintFramework_Tryin = vData
    UpdateDirtyFlag
End Property


Public Property Get Framework_Tryin() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Framework_Tryin
    Framework_Tryin = mintFramework_Tryin
End Property



Public Property Let Topical_Fluoride(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Topical_Fluoride = 5
    mintTopical_Fluoride = vData
    UpdateDirtyFlag
End Property


Public Property Get Topical_Fluoride() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Topical_Fluoride
    Topical_Fluoride = mintTopical_Fluoride
End Property



Public Property Let Polished_Preps(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Polished_Preps = 5
    mintPolished_Preps = vData
    UpdateDirtyFlag
End Property


Public Property Get Polished_Preps() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Polished_Preps
    Polished_Preps = mintPolished_Preps
End Property



Public Property Let Teeth_Replaced(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Teeth_Replaced = 5
    mstrTeeth_Replaced = vData
    UpdateDirtyFlag
End Property


Public Property Get Teeth_Replaced() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Teeth_Replaced
    Teeth_Replaced = mstrTeeth_Replaced
End Property



Public Property Let Teeth_Prepped(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Teeth_Prepped = 5
    mstrTeeth_Prepped = vData
    UpdateDirtyFlag
End Property


Public Property Get Teeth_Prepped() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Teeth_Prepped
    Teeth_Prepped = mstrTeeth_Prepped
End Property



Public Property Let SupportingProcs(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SupportingProcs = 5
    mstrSupportingProcs = vData
    UpdateDirtyFlag
End Property


Public Property Get SupportingProcs() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SupportingProcs
    
    If mstrSupportingProcs = "" Then
        mstrSupportingProcs = gcSupportingProcs
    End If
    
    SupportingProcs = mstrSupportingProcs
End Property



'*******************************************************************
'Function: SupportingProcsDisplay
'Author: Sean McAvoy
'Date: 1/2/02
'Returns: String
'Purpose: To get the display string for supporting procedures
'*******************************************************************
Private Function SupportingProcsDisplay() As String

Dim oItems As New CMultiListItems
Dim oItem As CMultiListItem

    oItems.ListType = Standard
    oItems.LoadFromValueString mstrSupportingProcs
    SupportingProcsDisplay = oItems.Value_String_Display
Set oItems = Nothing
Set oItem = Nothing
End Function





'*****************************************************************************************
Private Property Get clsISection_DentalEncounterID() As Long
    
    clsISection_DentalEncounterID = mvarDentalEncounterID

End Property


Private Property Get clsISection_PatientID() As Long
    
    clsISection_PatientID = mvarPatientID

End Property



Private Function clsISection_GetRTFNote() As String
Const sPROC_NAME As String = "clsTreatRP_RPD.GetRTFNote"
Dim sNote As String
Dim sTxt As String
Dim objToothCtrl As NotesPlusUtilities.CToothControl
Dim strTeethPreppedDisp As String
Dim strTeethReplDisp As String


On Error GoTo ErrHnd
    
    Set objToothCtrl = New NotesPlusUtilities.CToothControl
    
    If mstrTeeth_Prepped = "" And _
        mstrTeeth_Replaced = "" And _
        mintPolished_Preps = 0 And _
        mintTopical_Fluoride = 0 And _
        mintFramework_Tryin = 0 And _
        mintWaxrims = 0 And _
        mintAltered_Cast_Impre = 0 And _
        mintAdjust_RPD = 0 And _
        mintDeliver_RPD = 0 And _
        mintPost_Insertion = 0 And _
        mintPost_Delivery_Adj_RPD = 0 And _
        mstrRPD_Comments = "" And mstrSupportingProcs = "" Then
        
        'No need to update note..
        sNote = ""
        clsISection_GetRTFNote = sNote
        Exit Function
    End If
    
    If mstrTeeth_Prepped <> "" Then
        strTeethPreppedDisp = objToothCtrl.BuildDisplayString(mstrTeeth_Prepped)
    End If
    
    If mstrTeeth_Replaced <> "" Then
        strTeethReplDisp = objToothCtrl.BuildDisplayString(mstrTeeth_Replaced)
    End If
    
        If mstrTeeth_Prepped <> "" Or mstrTeeth_Replaced <> "" Or mintPolished_Preps <> 0 Or mintTopical_Fluoride <> 0 Then
            sNote = sNote + crtfStartControlName + "Prepped and Replaced: "
            If mstrTeeth_Prepped <> "" Then sNote = sNote + "Teeth Prepped: " + strTeethPreppedDisp + "; "
            If mstrTeeth_Replaced <> "" Then sNote = sNote + "Teeth Replaced: " + strTeethReplDisp + "; "
            If mintPolished_Preps <> 0 Then sNote = sNote + "Polished Preps; "
            If mintTopical_Fluoride <> 0 Then sNote = sNote + "Topical Fluoride; "
            'Remove the last "; "
            sNote = Left(sNote, Len(sNote) - Len("; "))
            sNote = sNote + crtfEndControlName
        End If
    
        'SCR 43361 RWalker 10/21/03 Added check of mintPost_Delivery_Adj_RPD to display even if this one not selected
        If mintFramework_Tryin <> 0 Or mintWaxrims <> 0 Or mintAltered_Cast_Impre <> 0 Then
            sNote = sNote + crtfStartControlName + "Framework and Rims: "
            If mintFramework_Tryin <> 0 Then sNote = sNote + "Framework Try-in; "
            If mintWaxrims <> 0 Then sNote = sNote + "Wax Rims; "
            If mintAltered_Cast_Impre <> 0 Then sNote = sNote + "Altered Cast Impression; "
            'Remove the last "; "
            sNote = Left(sNote, Len(sNote) - Len("; "))
            sNote = sNote + crtfEndControlName
        End If
    
        If mintAdjust_RPD <> 0 Or mintDeliver_RPD <> 0 Or mintPost_Insertion <> 0 Then
            sNote = sNote + crtfStartControlName + "RPD Delivery: "
            If mintAdjust_RPD <> 0 Then sNote = sNote + "Adjust RPD; "
            If mintDeliver_RPD <> 0 Then sNote = sNote + "Deliver RPD; "
            If mintPost_Insertion <> 0 Then sNote = sNote + "Post-insertion Instructions; "
        'SCR 43361 RWalker 10/21/03    If mintPost_Delivery_Adj_RPD <> 0 Then sNote = sNote + "Adjust RPD; "
            'Remove the last "; "
            sNote = Left(sNote, Len(sNote) - Len("; "))
            sNote = sNote + crtfEndControlName
        End If
    
        'SCR 43361 RWalker 10/21/03
        If mintPost_Delivery_Adj_RPD <> 0 Then
            sNote = sNote + crtfStartControlName + "Post-delivery Adjustment: "
            If mintPost_Delivery_Adj_RPD <> 0 Then sNote = sNote + "Adjust Removable Partial Denture; "
            sNote = Left(sNote, Len(sNote) - Len("; "))
            sNote = sNote + crtfEndControlName
        End If
    
        sTxt = SupportingProcsDisplay
        If sTxt <> "" Then
            sNote = sNote & crtfStartControlName & "Supporting Procedures: " & sTxt + crtfEndControlName
        End If
    
        If mstrRPD_Comments <> "" Then
            If mstrRPD_Comments <> "" Then sNote = sNote + crtfStartControlName & "RPD Comments: " + mstrRPD_Comments + crtfEndControlName
        End If
        
FuncExit:
    clsISection_GetRTFNote = sNote
    Set objToothCtrl = Nothing
    Exit Function
    
Exit Function
ErrHnd:
    sNote = "Error in generating note for Treatment - Removable Partial Denture section."
    clsISection_GetRTFNote = sNote
    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & "." & sPROC_NAME, App.Title, vbCritical
    GoTo FuncExit

End Function

'*******************************************************************
'Function: GetNote
'Author: Sean McAvoy
'Date:  12/21/01
'Returns: A string that is displayed in the narrative summary for this encounter
'Purpose: To compare the values from this encounter and the previous encounter and
'return a text note.
'*******************************************************************
Private Function clsISection_GetNote() As String
Const sPROC_NAME As String = "clsTreatRP_RPD.GetNote"
Dim sNote As String
Dim sTxt As String

On Error GoTo ErrHnd
   
    If mstrTeeth_Prepped = "" And _
        mstrTeeth_Replaced = "" And _
        mintPolished_Preps = 0 And _
        mintTopical_Fluoride = 0 And _
        mintFramework_Tryin = 0 And _
        mintWaxrims = 0 And _
        mintAltered_Cast_Impre = 0 And _
        mintAdjust_RPD = 0 And _
        mintDeliver_RPD = 0 And _
        mintPost_Insertion = 0 And _
        mintPost_Delivery_Adj_RPD = 0 And _
        mstrRPD_Comments = "" Then
        
        'No need to update note..
        sNote = ""
        clsISection_GetNote = sNote
        Exit Function
    End If

    sNote = vbCrLf + "Treatment - Removable Pros - Removable Partial Denture" + vbCrLf
    If mstrTeeth_Prepped <> "" Then sNote = sNote + vbTab + "Teeth Prepped: " + Replace(mstrTeeth_Prepped, "|", ", ") + vbCrLf
    If mstrTeeth_Replaced <> "" Then sNote = sNote + vbTab + "Teeth Replaced: " + Replace(mstrTeeth_Replaced, "|", ", ") + vbCrLf
    If mintPolished_Preps <> 0 Then sNote = sNote + vbTab + "Polished Preps" + vbCrLf
    
    If mintTopical_Fluoride <> 0 Then sNote = sNote + vbTab + "Topical Fluoride" + vbCrLf
    If mintFramework_Tryin <> 0 Then sNote = sNote + vbTab + "Framework Try-in" + vbCrLf
    If mintWaxrims <> 0 Then sNote = sNote + vbTab + "Wax Rims" + vbCrLf
    If mintAltered_Cast_Impre <> 0 Then sNote = sNote + vbTab + "Altered Cast Impression" + vbCrLf
    If mintAdjust_RPD <> 0 Then sNote = sNote + vbTab + "Adjust RPD" + vbCrLf
    If mintDeliver_RPD <> 0 Then sNote = sNote + vbTab + "Deliver RPD" + vbCrLf
    If mintPost_Insertion <> 0 Then sNote = sNote + vbTab + "Post-insertion Instructions" + vbCrLf
    If mintPost_Delivery_Adj_RPD <> 0 Then sNote = sNote + vbTab + "Adjust Removable Partial Denture" + vbCrLf
    sTxt = SupportingProcsDisplay
    If sTxt <> "" Then
        sNote = sNote & vbTab & "Supporting Procedures: " & sTxt & vbCrLf
    End If
    
    If mstrRPD_Comments <> "" Then sNote = sNote + vbTab + "RPD Comments: " + mstrRPD_Comments + vbCrLf
    
    clsISection_GetNote = sNote
    Exit Function

Exit Function
ErrHnd:
    sNote = "Error in generating note for Treatment - Removable Partial Denture section."
    clsISection_GetNote = sNote
    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & "." & sPROC_NAME, App.Title, vbCritical
    
End Function

'*******************************************************************
'Function: SaveSection
'Author: Sean McAvoy
'Date:  12/21/01
'Returns: True/False
'Purpose: Saves the data to the table
'*******************************************************************
Private Function clsISection_SaveSection() As Boolean
Dim mstrTeeth_Prepped_TMP As String
Dim mstrTeeth_Replaced_TMP As String
Dim mstrRPD_Comments_TMP As String

Dim sSQL As String

Const sPROC_NAME As String = "clsTreatRP_RPD.SaveSection"

On Error GoTo ErrHnd

    If dfDirtyFlag = dfNoAction Then
        clsISection_SaveSection = True
        Exit Function
    End If
    If mvarDentalEncounterID = 0 Then
        clsISection_SaveSection = False
        Exit Function
    End If
    
    mstrTeeth_Prepped_TMP = RQ(mstrTeeth_Prepped)
    mstrTeeth_Replaced_TMP = RQ(mstrTeeth_Replaced)
    mstrRPD_Comments_TMP = RQ(mstrRPD_Comments)
    
    'KDunne 11/5/2003
    'SCR 44381 This code is unnecessary and is causing a problem with 'Null' being
    '          displayed under certain circumstances.
    '
'    If InStr(1, mstrSupportingProcs, "True", vbTextCompare) = 0 Then mstrSupportingProcs = ""
'    If mstrSupportingProcs = "" Then
'        mstrSupportingProcs = "Null"
'    End If
    
    If dfDirtyFlag = dfAdd Then
        sSQL = "INSERT INTO TEMPL_TREAT_RP_RPD (" _
            & " DENTAL_ENCOUNTER_ID, TEETH_PREPPED, TEETH_REPLACED, POLISHED_PREPS, TOPICAL_Fluoride, FRAMEWORK_TRYIN," _
            & " WAXRIMS, ALTERED_CAST_IMPRE, ADJUST_RPD, DELIVER_RPD, POST_INSERTION, POST_DELIVERY_ADJ_RPD, RPD_COMMENTS, SUPPORTINGPROCEDURES) " _
            & " VALUES" _
            & " (" & CStr(mvarDentalEncounterID) & ", " _
            & " '" & mstrTeeth_Prepped_TMP & "', " _
            & " '" & mstrTeeth_Replaced_TMP & "', " _
            & CStr(mintPolished_Preps) & ", " _
            & CStr(mintTopical_Fluoride) & ", " _
            & CStr(mintFramework_Tryin) & ", " _
            & CStr(mintWaxrims) & ", " _
            & CStr(mintAltered_Cast_Impre) & ", " _
            & CStr(mintAdjust_RPD) & ", " _
            & CStr(mintDeliver_RPD) & ", " _
            & CStr(mintPost_Insertion) & ", " _
            & CStr(mintPost_Delivery_Adj_RPD) & ", " _
            & " '" & mstrRPD_Comments_TMP & "', " _
            & " '" & mstrSupportingProcs & "') "
    
    ElseIf dfDirtyFlag = dfUpdate Then
        sSQL = "UPDATE TEMPL_TREAT_RP_RPD SET " + _
                "TEETH_PREPPED = '" + mstrTeeth_Prepped_TMP + "', " + _
                "TEETH_REPLACED = '" + mstrTeeth_Replaced_TMP + "', " + _
                "POLISHED_PREPS = " + CStr(mintPolished_Preps) + ", " + _
                "TOPICAL_Fluoride = " + CStr(mintTopical_Fluoride) + ", " + _
                "FRAMEWORK_TRYIN = " + CStr(mintFramework_Tryin) + ", " + _
                "WAXRIMS = " + CStr(mintWaxrims) + ", " + _
                "ALTERED_CAST_IMPRE = " + CStr(mintAltered_Cast_Impre) + ", " + _
                "ADJUST_RPD = " + CStr(mintAdjust_RPD) + ", " + _
                "DELIVER_RPD = " + CStr(mintDeliver_RPD) + ", " + _
                "POST_INSERTION = " + CStr(mintPost_Insertion) + ", " + _
                "POST_DELIVERY_ADJ_RPD = " + CStr(mintPost_Delivery_Adj_RPD) + ", " + _
                "RPD_COMMENTS = '" + mstrRPD_Comments_TMP + "', " + _
                "SUPPORTINGPROCEDURES = '" + mstrSupportingProcs + "'  " + _
                "Where (DENTAL_ENCOUNTER_ID = " & CStr(mvarDentalEncounterID) & ")"
    ElseIf dfDirtyFlag = dfdelete Then
        clsISection_SaveSection = False
        Exit Function
    End If
    
    goconn.ExecuteSQL sSQL
'    If mstrSupportingProcs = "Null" Then mstrSupportingProcs = ""  scr 44381
    dfDirtyFlag = dfNoAction
    
    clsISection_SaveSection = True

Exit Function

ErrHnd:
    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & "." & sPROC_NAME, App.Title, vbCritical

End Function


'*******************************************************************
'Function: LoadSection
'Author: Sean McAvoy
'Date:  12/21/01
'Returns: True/False
'Purpose: This loads the data from the db
'*******************************************************************
Private Function clsISection_LoadSection(ByVal PatientID As Long, ByVal DentalEncounterID As Long) As Boolean

Dim sSQL As String, rs As Recordset
Const sPROC_NAME = "clsTreatRP_RPD.LoadSection"

On Error GoTo ErrHnd

    mvarPatientID = PatientID
    mvarDentalEncounterID = DentalEncounterID
    If mvarDentalEncounterID = 0 Then
        clsISection_LoadSection = False
        Exit Function
    End If
    
    'Load the values
    sSQL = "SELECT * FROM TEMPL_TREAT_RP_RPD Where " + _
    "DENTAL_ENCOUNTER_ID = " + CStr(mvarDentalEncounterID)

    Set rs = goconn.OpenRecordset(sSQL)
    
    If rs.EOF Then
        dfDirtyFlag = dfAdd
        mstrTeeth_Prepped = ""
        mstrTeeth_Replaced = ""
        mintPolished_Preps = 0
        mintTopical_Fluoride = 0
        mintFramework_Tryin = 0
        mintWaxrims = 0
        mintAltered_Cast_Impre = 0
        mintAdjust_RPD = 0
        mintDeliver_RPD = 0
        mintPost_Insertion = 0
        mintPost_Delivery_Adj_RPD = 0
        mstrRPD_Comments = ""
        mstrSupportingProcs = ""
    Else
        dfDirtyFlag = dfNoAction
        'Set the values from the db including "" to allow for nulls
        mstrTeeth_Prepped = rs("TEETH_PREPPED") & ""
        mstrTeeth_Replaced = rs("TEETH_REPLACED") & ""
        mintPolished_Preps = Val(rs("POLISHED_PREPS") & "")
        mintTopical_Fluoride = Val(rs("TOPICAL_Fluoride") & "")
        mintFramework_Tryin = Val(rs("FRAMEWORK_TRYIN") & "")
        mintWaxrims = Val(rs("WAXRIMS") & "")
        mintAltered_Cast_Impre = Val(rs("ALTERED_CAST_IMPRE") & "")
        mintAdjust_RPD = Val(rs("ADJUST_RPD") & "")
        mintDeliver_RPD = Val(rs("DELIVER_RPD") & "")
        mintPost_Insertion = Val(rs("POST_INSERTION") & "")
        mintPost_Delivery_Adj_RPD = Val(rs("POST_DELIVERY_ADJ_RPD") & "")
        mstrRPD_Comments = rs("RPD_COMMENTS") & ""
        
        
        'Because of the issues with Tuxedo (4k limit), two sql calls will
        'need to be made to retrieve this additional field.
        rs.Close
        sSQL = "SELECT SUPPORTINGPROCEDURES FROM TEMPL_TREAT_RP_RPD Where " + _
        "DENTAL_ENCOUNTER_ID = " + CStr(mvarDentalEncounterID)
    
        Set rs = goconn.OpenRecordset(sSQL)
        If rs.EOF Then
            mstrSupportingProcs = ""
        Else
            mstrSupportingProcs = rs("SUPPORTINGPROCEDURES") & ""
        End If

        
    End If
    
    rs.Close
    
    
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
    
 mstrTeeth_Prepped = ""
 mstrTeeth_Replaced = ""
 mintPolished_Preps = 0
 mintTopical_Fluoride = 0
 mintFramework_Tryin = 0
 mintWaxrims = 0
 mintAltered_Cast_Impre = 0
 mintAdjust_RPD = 0
 mintDeliver_RPD = 0
 mintPost_Insertion = 0
 mintPost_Delivery_Adj_RPD = 0
 mstrRPD_Comments = ""
 mstrSupportingProcs = ""
    
    
End Sub







