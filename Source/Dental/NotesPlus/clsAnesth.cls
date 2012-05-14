VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAnesth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Implements clsISection

Private Const XML_ELEMENT_NAME As String = "clsAnesth"
Private mvarXMLPersist As New clsXMLPersist
Private mvarXMLElement As IXMLDOMElement

Private ofrmAnesth As frmAnesthDialog

Private oAnesthLocal As colAnesthLocal
Private oIVs As colAnesthIVs

Private mvarPatientID As Long
Private mvarDentalEncounterID As Long
Private mvarAnestheticUsed As Boolean

'Holds a reference to the datalayer
Private mvarDataLayer As clsNotes

Private dfDirtyFlag As DirtyFlagEnum

Public Property Let AnestheticUsed(vData As Boolean)

    mvarAnestheticUsed = vData
    UpdateDirtyFlag
    
End Property

Public Property Get AnestheticUsed() As Boolean

    AnestheticUsed = mvarAnestheticUsed
    
End Property

Public Property Set DataLayer(vData As clsNotes)

    Set mvarDataLayer = vData
    
End Property

'*******************************************************************
'Function: ShowAnestheticForm
'Author: Scott Foster
'Date: 11/07/01
'Returns: Nothing
'Purpose: Shows the Anesthetic form so the user can enter anesthetic info
'*******************************************************************
Public Function ShowAnestheticForm() As Boolean

    'If the app didn't call load section then don't display the form.
    If mvarPatientID = 0 Then
        Exit Function
    End If
    
    'Set the patient id
    ofrmAnesth.PatientID = mvarPatientID
    
    'Refresh the IV grid with the oIVs object
    ofrmAnesth.RefreshIVGrid oIVs
    
    'Refresh the local anesth grid
    ofrmAnesth.RefreshLocalGrid oAnesthLocal
        
    'Set the anesthetic checkbox
    ofrmAnesth.chkNoAnesthUsed.Value = Abs(CInt(Not mvarAnestheticUsed))
    
    'Set the reference to the datalayer
    Set ofrmAnesth.DataLayer = mvarDataLayer
    
    cwobjCWComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, cwlMyProgId
    ofrmAnesth.Show vbModal
    cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, cwlMyProgId
    
    'Update the object with the value of the checkbox
    Me.AnestheticUsed = Not CBool(ofrmAnesth.chkNoAnesthUsed.Value)
    
    'Return true.
    ShowAnestheticForm = True
    
End Function


Private Sub Class_Terminate()
    
    Set mvarXMLElement = Nothing
    Set mvarXMLPersist = Nothing
    Set oAnesthLocal = Nothing
    Set oIVs = Nothing
    Set ofrmAnesth = Nothing

End Sub


Private Property Get clsISection_DentalEncounterID() As Long
    
    clsISection_DentalEncounterID = mvarDentalEncounterID

End Property
'*******************************************************************
'Function: GetNote
'Author: Scott Foster
'Date: 10/19/01
'Returns: A string that is displayed in the narrative summary for this encounter
'Purpose: To compare the values from this encounter and the previous encounter and
'return a text note.
'*******************************************************************
Private Function clsISection_GetNote() As String

Dim sTemp As String
Dim oLocal As clsAnesthLocal
Dim oIV As clsAnesthIV
    
    If mvarAnestheticUsed = False Then
        sTemp = vbTab & "No anesthetic used." & vbCrLf
    End If
    
    For Each oLocal In oAnesthLocal
        sTemp = sTemp & oLocal.GetNote
    Next
    
    For Each oIV In oIVs
        sTemp = sTemp & oIV.GetNote
    Next
    
    If sTemp <> "" Then
        clsISection_GetNote = "Anesthetics" & vbCrLf & sTemp
    End If

End Function

Private Function clsISection_GetRTFNote() As String

Dim sTemp As String
Dim sLocalTemp As String
Dim sIVTemp As String
Dim oLocal As clsAnesthLocal
Dim oIV As clsAnesthIV
    
    If mvarAnestheticUsed = False Then
        sTemp = "No anesthetic used."
    End If
    
    For Each oLocal In oAnesthLocal
        sLocalTemp = sLocalTemp & oLocal.GetRTFNote
    Next
    If sLocalTemp <> "" Then
        sTemp = crtfStartTabName & "Local Anesthetics: " & crtfEndTabName & sLocalTemp
    End If
    
    'Get the note for IV
    For Each oIV In oIVs
        sIVTemp = sIVTemp & oIV.GetRTFNote
    Next
    If sIVTemp <> "" Then
        sTemp = sTemp & crtfStartTabName & "IV sedations and general anesthetics: " & crtfEndTabName & sIVTemp
    End If
    
    If sTemp <> "" Then
        clsISection_GetRTFNote = crtfStartTemplateName & "Anesthetics" & crtfEndTemplateName & sTemp
    End If
    
End Function

Private Function clsISection_GetWorkloadCode() As String

End Function

'*******************************************************************
'Function: LoadSection
'Author: Scott Foster
'Date: 10-22-01
'Returns: True/False
'Purpose: This loads the data from the db
'*******************************************************************
Private Function clsISection_LoadSection(ByVal PatientID As Long, ByVal DentalEncounterID As Long) As Boolean

Dim sSQL As String, rs As Recordset, rsPrev As Recordset

Const sPROC_NAME = "LoadSection"

On Error GoTo ErrHnd

    If mvarDentalEncounterID = DentalEncounterID Then
        Exit Function
    End If
    
    mvarPatientID = PatientID
    mvarDentalEncounterID = DentalEncounterID
    
    'Get the xml element
    Set mvarXMLElement = mvarXMLPersist.ExamXML.XMLDoc.selectSingleNode("//" & XML_ELEMENT_NAME)
    
    If mvarXMLElement Is Nothing Then
        'Set the defaults
        mvarAnestheticUsed = True
    Else
        'Load the value
        mvarAnestheticUsed = CBool(mvarXMLElement.GetAttribute("ANESTH_USED"))
    End If
    
'    'Load the values
'    sSQL = "Select ANESTH_USED FROM TEMPL_ANESTH_MAIN WHERE DENTAL_ENCOUNTER_ID = " & mvarDentalEncounterID
'
'    Set rs = goconn.OpenRecordset(sSQL)
'
'    If rs.EOF Then
'        dfDirtyFlag = dfAdd
'        'Init the values here to ""
'        mvarAnestheticUsed = True
'    Else
'        dfDirtyFlag = dfNoAction
'        'Set the values from the db including "" to allow for nulls
'        mvarAnestheticUsed = CBool(Val(rs("ANESTH_USED") & ""))
'    End If
'
'    rs.Close
'
'    Set rs = Nothing
    
    'Load the form
    Set ofrmAnesth = New frmAnesthDialog
    
    'Load the AnesthLocal grid
    Set oAnesthLocal = New colAnesthLocal
    oAnesthLocal.LoadSection mvarPatientID, mvarDentalEncounterID
    
    'Load the oIVs
    Set oIVs = New colAnesthIVs
    oIVs.LoadSection mvarPatientID, mvarDentalEncounterID
    
    clsISection_LoadSection = True
    
Exit Function

ErrHnd:
    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & "." & sPROC_NAME, App.Title, vbCritical

End Function

Private Property Get clsISection_PatientID() As Long

    clsISection_PatientID = mvarPatientID

End Property
'*******************************************************************
'Function: SaveSection
'Author: Scott Foster
'Date: 10-22-01
'Returns: True/False
'Purpose: Saves the data to the table
'*******************************************************************
Private Function clsISection_SaveSection() As Boolean

Dim sSQL As String

Const sPROC_NAME As String = "SaveSection"

On Error GoTo ErrHnd
    
    'Save the AnesthLocal
    oAnesthLocal.SaveSection
    
    'Save the oIVs
    oIVs.SaveSection
    
    If dfDirtyFlag = dfNoAction Then
        clsISection_SaveSection = True
        Exit Function
        
'    ElseIf dfDirtyFlag = dfAdd Then
'        'The record does not exist, so create an insert statement
'        sSQL = "INSERT INTO TEMPL_ANESTH_MAIN (DENTAL_ENCOUNTER_ID, ANESTH_USED) VALUES (" & mvarDentalEncounterID & ", " & CInt(mvarAnestheticUsed) & ")"
'    ElseIf dfDirtyFlag = dfUpdate Then
'        'The record exists, so create an update statement
'        sSQL = "UPDATE TEMPL_ANESTH_MAIN SET ANESTH_USED = " & CInt(mvarAnestheticUsed) & " WHERE DENTAL_ENCOUNTER_ID = " & mvarDentalEncounterID
    End If
    
    'Save the xmlvalue
    If mvarXMLElement Is Nothing Then
        Set mvarXMLElement = mvarXMLPersist.ExamXML.XMLDoc.selectSingleNode("//" & XML_ELEMENT_NAME)
        If mvarXMLElement Is Nothing Then
            'Create the element
            Set mvarXMLElement = mvarXMLPersist.ExamXML.XMLDoc.createElement(XML_ELEMENT_NAME)
            mvarXMLPersist.ExamXML.XMLDoc.documentElement.appendChild mvarXMLElement
        End If
    End If
    
    'Save the value
    With mvarXMLElement
        .setAttribute "ANESTH_USED", mvarAnestheticUsed
    End With
    
'    goconn.ExecuteSQL sSQL
    
    dfDirtyFlag = dfNoAction
    
    clsISection_SaveSection = True

Exit Function

ErrHnd:
    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & "." & sPROC_NAME, App.Title, vbCritical
Resume
End Function
Private Function UpdateDirtyFlag()

    If dfDirtyFlag = dfNoAction Then
        dfDirtyFlag = dfUpdate
    End If
    
End Function

