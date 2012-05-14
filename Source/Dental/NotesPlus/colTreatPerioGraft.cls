VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colTreatPGrafts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsTreatPerioGraft"
Attribute VB_Ext_KEY = "Member0" ,"clsTreatPerioGraft"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Implements clsISection

'local variable to hold collection
Private mCol As Collection

Private mvarPatientID As Long
Private mvarDentalEncounterID As Long

'Constants for the grid rows
Private Const cColGraft As Integer = 0
Private Const cColMembrane As Integer = 1
Private Const cColSutures As Integer = 2
Private Const cColRoot As Integer = 3
Private Const cColPerio As Integer = 4
'Init the arMembrane in the class initialize event
Private arMembrane(5) As String

'Init the arRootCond in the class initialize event
Private arRootCond(4) As String

'Init the arDressing in the class initialize event
Private arDressing(3) As String

Public Property Get ColGraft() As Integer
    ColGraft = cColGraft
End Property
Public Property Get ColMembrane() As Integer
    ColMembrane = cColMembrane
End Property
Public Property Get ColSutures() As Integer
    ColSutures = cColSutures
End Property
Public Property Get ColRoot() As Integer
    ColRoot = cColRoot
End Property
Public Property Get ColPerio() As Integer
    ColPerio = cColPerio
End Property


'*******************************************************************
'Function: ShowGraftDialog
'Author: Scott Foster
'Date: 10-26-01
'Returns: Nothing
'Purpose: Displays the graft dialog so the user can enter the graft material
' specific data
'*******************************************************************
Public Sub ShowGraftDialog(PGraft As clsTreatPGraft)

Dim frmPGraft As New frmTreatGraftDialog

    'Set the properties
    frmPGraft.sNewGraftMaterial = PGraft.GraftMaterial
    frmPGraft.txtMaterials = PGraft.GraftMaterialDisplay
    frmPGraft.txtManufacturer = PGraft.MaterialManuf
    frmPGraft.txtLot = PGraft.MaterialLot
    frmPGraft.txtComments = PGraft.MaterialComments
    
    Screen.MousePointer = vbDefault
    
    'Show the form
    cwobjCWComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, cwlMyProgId
    frmPGraft.Show vbModal
    cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, cwlMyProgId
    
    Screen.MousePointer = vbHourglass
    
    If frmPGraft.bCancel = False Then
        'Update the values
        PGraft.GraftMaterial = frmPGraft.sNewGraftMaterial
        PGraft.MaterialManuf = frmPGraft.txtManufacturer
        PGraft.MaterialLot = frmPGraft.txtLot
        PGraft.MaterialComments = frmPGraft.txtComments
    End If
    
    Set frmPGraft = Nothing
    
End Sub


'*******************************************************************
'Function: ShowMembraneDialog
'Author: Scott Foster
'Date: 10-29-01
'Returns: Nothing
'Purpose: Displays the membrane dialog so the user can enter the membrane material
' specific data
'*******************************************************************
Public Sub ShowMembraneDialog(PGraft As clsTreatPGraft)

Dim frmPMembrane As New frmTreatMembraneDialog
Dim iCounter As Integer

    'Set the properties
        
    'Add the items to the combo box
    For iCounter = LBound(arMembrane) To UBound(arMembrane)
        frmPMembrane.cboMembrane.AddItem arMembrane(iCounter)
    Next
    
    'Set the one the user selected
    For iCounter = LBound(arMembrane) To UBound(arMembrane)
        If Trim(frmPMembrane.cboMembrane.List(iCounter)) = Trim(PGraft.Membrane) Then
            frmPMembrane.cboMembrane.ListIndex = iCounter
        End If
    Next
    frmPMembrane.txtManuf = PGraft.MembraneManuf
    frmPMembrane.txtLot = PGraft.MembraneLot
    frmPMembrane.txtComments = PGraft.MembraneComments
    
    Screen.MousePointer = vbDefault
    
    'Show the form
    cwobjCWComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, cwlMyProgId
    frmPMembrane.Show vbModal
    cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, cwlMyProgId
    
    Screen.MousePointer = vbHourglass
    
    If frmPMembrane.bCancel = False Then
        'Update the values
        PGraft.Membrane = frmPMembrane.cboMembrane.Text
        PGraft.MembraneManuf = frmPMembrane.txtManuf
        PGraft.MembraneLot = frmPMembrane.txtLot
        PGraft.MembraneComments = frmPMembrane.txtComments
    End If
    
    Set frmPMembrane = Nothing
    
End Sub

'*******************************************************************
'Function: ShowRootCondDialog
'Author: Scott Foster
'Date: 10-29-01
'Returns: Nothing
'Purpose: Displays the root conditioning dialog so the user can enter the root conditioning
' specific data
'*******************************************************************
Public Sub ShowRootCondDialog(PRootCond As clsTreatPGraft)

Dim frmPRootCond As New frmTreatRootCondDialog
Dim iCounter As Integer

    'Set the properties
        
    'Add the items to the combo box
    For iCounter = LBound(arRootCond) To UBound(arRootCond)
        frmPRootCond.cboRootCond.AddItem arRootCond(iCounter)
    Next
    
    'Set the one the user selected
    For iCounter = LBound(arRootCond) To UBound(arRootCond)
        If Trim(frmPRootCond.cboRootCond.List(iCounter)) = Trim(PRootCond.RootCond) Then
            frmPRootCond.cboRootCond.ListIndex = iCounter
        End If
    Next
    frmPRootCond.txtComments = PRootCond.RootCondComments
    
    Screen.MousePointer = vbDefault
    
    'Show the form
    cwobjCWComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, cwlMyProgId
    frmPRootCond.Show vbModal
    cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, cwlMyProgId
    
    Screen.MousePointer = vbHourglass
    
    If frmPRootCond.bCancel = False Then
        'Update the values
        PRootCond.RootCond = frmPRootCond.cboRootCond.Text
        PRootCond.RootCondComments = frmPRootCond.txtComments
    End If
    
    Set frmPRootCond = Nothing
    
End Sub

'*******************************************************************
'Function: ShowDressingDialog
'Author: Scott Foster
'Date: 10-29-01
'Returns: Nothing
'Purpose: Displays the dressing dialog so the user can enter the dressing
' specific data
'*******************************************************************
Public Sub ShowDressingDialog(PDressing As clsTreatPGraft)

Dim frmPDressing As New frmTreatDressingDialog
Dim iCounter As Integer

    'Set the properties
        
    'Add the items to the combo box
    For iCounter = LBound(arDressing) To UBound(arDressing)
        frmPDressing.cboDressing.AddItem arDressing(iCounter)
    Next
    
    'Set the one the user selected
    For iCounter = LBound(arDressing) To UBound(arDressing)
        If Trim(frmPDressing.cboDressing.List(iCounter)) = Trim(PDressing.PerioDressing) Then
            frmPDressing.cboDressing.ListIndex = iCounter
        End If
    Next
    frmPDressing.txtComments = PDressing.PerioDressingComments
    
    Screen.MousePointer = vbDefault
    
    'Show the form
    cwobjCWComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, cwlMyProgId
    frmPDressing.Show vbModal
    cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, cwlMyProgId
    
    Screen.MousePointer = vbHourglass
    
    If frmPDressing.bCancel = False Then
        'Update the values
        PDressing.PerioDressing = frmPDressing.cboDressing.Text
        PDressing.PerioDressingComments = frmPDressing.txtComments
    End If
    
    Set frmPDressing = Nothing
    
End Sub



'*******************************************************************
'Function: ShowSutureDialog
'Author: Scott Foster
'Date: 10-29-01
'Returns: Nothing
'Purpose: Displays the suture dialog so the user can enter the suture type
' specific data
'*******************************************************************
Public Sub ShowSutureDialog(PSuture As clsTreatPGraft)

Dim frmPSuture As New frmTreatSutureDialog

    'Set the properties
    frmPSuture.sNewSuture = PSuture.SutureType
    frmPSuture.txtSuture = PSuture.SutureTypeDisplay
    frmPSuture.txtNumber = PSuture.SutureNumber
    frmPSuture.txtTechnique = PSuture.SutureTech
    
    Screen.MousePointer = vbDefault
    
    'Show the form
    cwobjCWComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, cwlMyProgId
    frmPSuture.Show vbModal
    cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, cwlMyProgId
    
    Screen.MousePointer = vbHourglass
    
    If frmPSuture.bCancel = False Then
        'Update the values
        PSuture.SutureType = frmPSuture.sNewSuture
        PSuture.SutureNumber = frmPSuture.txtNumber
        PSuture.SutureTech = frmPSuture.txtTechnique
    End If
    
    Set frmPSuture = Nothing

End Sub

Public Property Get clsISection_DentalEncounterID() As Long
    
    clsISection_DentalEncounterID = mvarDentalEncounterID

End Property



Public Property Get clsISection_PatientID() As Long
    
    clsISection_PatientID = mvarPatientID

End Property



Public Function clsISection_GetRTFNote() As String

Dim oGraft As clsTreatPGraft
    
    'Get all the notes from the objects
    For Each oGraft In mCol
        clsISection_GetRTFNote = clsISection_GetRTFNote & oGraft.GetRTFNote
    Next

End Function

'*******************************************************************
'Function: clsISection_GetNote
'Author: Scott Foster
'Date: 10/19/01
'Returns: A string that is displayed in the narrative summary for this encounter
'Purpose: To compare the values from this encounter and the previous encounter and
'return a text note.
'*******************************************************************
Public Function clsISection_GetNote() As String

Dim oGraft As clsTreatPGraft
    
    'Get all the notes from the objects
    For Each oGraft In mCol
        clsISection_GetNote = clsISection_GetNote & oGraft.GetNote
    Next

End Function

'*******************************************************************
'Function: clsISection_SaveSection
'Author: Scott Foster
'Date: 10-22-01
'Returns: True/False
'Purpose: Saves the data to the table
'*******************************************************************
Public Function clsISection_SaveSection() As Boolean

Dim oGraft As clsTreatPGraft

Const sPROC_NAME = "clsISection_SaveSection"

    'Save all the objects
    For Each oGraft In mCol
        If oGraft.SaveSection = False Then
            clsISection_SaveSection = False
        End If
        
        If oGraft.DirtyFlag = dfdelete Then
            'Remove it from the collection because it has already been deleted from the db
            Remove oGraft.GraftID
        End If
    Next

Exit Function

ErrHnd:
    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & "." & sPROC_NAME, App.Title, vbCritical

End Function


'*******************************************************************
'Function: clsISection_LoadSection
'Author: Scott Foster
'Date: 10-22-01
'Returns: True/False
'Purpose: This loads the data from the db
'*******************************************************************
Public Function clsISection_LoadSection(ByVal PatientID As Long, ByVal DentalEncounterID As Long) As Boolean

Dim sSQL As String, rs As Recordset, rsPrev As Recordset
Dim objNewMember As clsTreatPGraft

Const sPROC_NAME = "LoadSection"

On Error GoTo ErrHnd

    mvarPatientID = PatientID
    mvarDentalEncounterID = DentalEncounterID
    
    'Load the values
    sSQL = "SELECT GRAFT_ID, GRAFT_MATERIAL, MATERIAL_MANUF, MATERIAL_LOT, MATERIAL_COMMENTS," _
        & " MEMBRANE, MEMBRANE_MANUF, MEMBRANE_LOT, MEMBRANE_COMMENTS," _
        & " SUTURE_TYPE, SUTURE_NUMBER, SUTURE_TECH," _
        & " ROOT_COND, ROOT_COND_COMMENTS, PERIO_DRESSING, PERIO_DRESSING_COMMENTS" _
        & " From TEMPL_TREAT_PERIO_GRAFT" _
        & " Where (DENTAL_ENCOUNTER_ID = " & mvarDentalEncounterID & ")"
    
    Set rs = goconn.OpenRecordset(sSQL)
    
    Do Until rs.EOF
        'Set the values from the db including "" to allow for nulls
        
        Set objNewMember = Add(mvarDentalEncounterID)
            
        'set the properties passed into the method
        objNewMember.GraftID = rs("Graft_ID")
        objNewMember.GraftMaterial = rs("GRAFT_MATERIAL") & ""
        objNewMember.MaterialManuf = rs("MATERIAL_MANUF") & ""
        objNewMember.MaterialLot = rs("MATERIAL_LOT") & ""
        objNewMember.MaterialComments = rs("MATERIAL_COMMENTS") & ""
        objNewMember.Membrane = rs("Membrane") & ""
        objNewMember.MembraneManuf = rs("MEMBRANE_MANUF") & ""
        objNewMember.MembraneLot = rs("MEMBRANE_LOT") & ""
        objNewMember.MembraneComments = rs("MEMBRANE_COMMENTS") & ""
        objNewMember.SutureType = rs("SUTURE_TYPE") & ""
        objNewMember.SutureNumber = rs("SUTURE_NUMBER") & ""
        objNewMember.SutureTech = rs("SUTURE_TECH") & ""
        objNewMember.RootCond = rs("ROOT_COND") & ""
        objNewMember.RootCondComments = rs("ROOT_COND_COMMENTS") & ""
        objNewMember.PerioDressing = rs("PERIO_DRESSING") & ""
        objNewMember.PerioDressingComments = rs("PERIO_DRESSING_COMMENTS") & ""
        objNewMember.DirtyFlag = dfNoAction
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    
    clsISection_LoadSection = True
    
Exit Function

ErrHnd:
    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & "." & sPROC_NAME, App.Title, vbCritical
    
End Function


Private Sub Class_Initialize()

    Set mCol = New Collection

    'Membrane choices
    arMembrane(0) = ""
    arMembrane(1) = "Non resorbable (PTFE)"
    arMembrane(2) = "Non resorbable (e-PTFE)"
    arMembrane(3) = "Resorbable (Collagen)"
    arMembrane(4) = "Resorbable (PLA/PGA)"
    arMembrane(5) = "Resorbable (Other)"
    
    'RootCond choices
    arRootCond(0) = ""
    arRootCond(1) = "Tetracycline"
    arRootCond(2) = "Citric acid"
    arRootCond(3) = "Edta"
    arRootCond(4) = "Other"

    'Dressing choices
    arDressing(0) = ""
    arDressing(1) = "Perio-Pac"
    arDressing(2) = "Barricaid"
    arDressing(3) = "Generic periodontal dressing"
    
End Sub



Public Function Add(EncounterID As Long) As clsTreatPGraft
    
    'create a new object
    Dim objNewMember As clsTreatPGraft
    Set objNewMember = New clsTreatPGraft

    'set the properties passed into the method
    objNewMember.DentalEncounterID = EncounterID
    objNewMember.DirtyFlag = dfAdd
    mCol.Add objNewMember

    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As clsTreatPGraft
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)
    
    If mCol(vntIndexKey).GraftID <> 0 Then
        'Set the dirtyflag to deleted
        mCol(vntIndexKey).DirtyFlag = dfdelete
    Else
        'The object has not been saved, so remove it.
        mCol.Remove vntIndexKey
    End If
    
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property




Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub













