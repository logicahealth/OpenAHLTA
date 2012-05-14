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
Attribute VB_Name = "clsTrNSETooth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mlngId As Long
Private mstrToothNumber As String
Private mlngSolvent As TreatEndo_RCTSolventsEnum
Private mstrSolvent As String
Private mstrMagnification As String
Private mlngIntraCanalMed As TreatEndo_IntracanalMedicantEnum
Private mstrIntraCanalText As String
Private mlngCottenPellet As ExamEndoYesNoEnum
Private mlngProvMaterial As TreatEndo_MaterialEnum
Private mstrProvMaterialText As String
Private mlngCement As TreatEndo_CementEnum
Private mstrCementText As String
Private mlngOcculusalAdj As TreatEndo_OccAdjEnum

Private mlngDirtyFlag As DirtyFlagEnum
Private mlngDentalEncounterId As Long

Public Enum TreatEndo_RCTSolventsEnum
    Not_Addressed = 0
    NaOCl525 = 1
    NaOCl26 = 2
    Saline = 3
    EDTA17 = 4
    Chloroform = 5
    RCPrep = 6
    Glyde = 7
    Other = 8
End Enum

Public Enum TreatEndo_IntracanalMedicantEnum
    Ca0H2 = 1
    Other = 2
    Not_Addressed = 0
End Enum

Public Enum TreatEndo_MaterialEnum
    Cavit = 1
    IRM = 2
    GI = 3
    Amalgam = 4
    Resin = 5
    ZOE = 6
    Other = 7
    Not_Addressed = 0
End Enum

Public Enum TreatEndo_CementEnum
    Dycal = 1
    TempBond = 2
    Other = 3
    Not_Addressed = 0
End Enum

Public Enum TreatEndo_OccAdjEnum
    PriorWL = 1
    PostTx = 2
    Not_Addressed = 0
End Enum

Private Sub Class_Initialize()
    mstrMagnification = cTreatNSEMagnificationList
    mstrSolvent = cTreatNSESolventsList
End Sub
Public Property Get ID() As Long
    ID = mlngId
End Property
Public Property Let ID(ByVal lngValue As Long)
    mlngId = lngValue
End Property
Public Property Get Tooth_Number() As String
    Tooth_Number = mstrToothNumber
End Property
Public Property Let Tooth_Number(ByVal strValue As String)
    mstrToothNumber = strValue
    UpdateDirtyFlag
End Property

Public Property Get Solvent() As String
        Solvent = mstrSolvent
End Property

Public Property Let Solvent(ByVal strValue As String)
    mstrSolvent = strValue
    UpdateDirtyFlag
End Property

Public Property Get Magnification() As String
    Magnification = mstrMagnification
End Property

Public Property Let Magnification(ByVal strValue As String)
    mstrMagnification = strValue
    UpdateDirtyFlag
End Property

Public Property Get Intracanal_Medicament() As TreatEndo_IntracanalMedicantEnum
    Intracanal_Medicament = mlngIntraCanalMed
    UpdateDirtyFlag
End Property

Public Property Let Intracanal_Medicament(ByVal lngValue As TreatEndo_IntracanalMedicantEnum)
    mlngIntraCanalMed = lngValue
    UpdateDirtyFlag
End Property

Public Property Get Intracanal_Medicament_Text() As String
    Intracanal_Medicament_Text = mstrIntraCanalText
End Property
Public Property Let Intracanal_Medicament_Text(ByVal strValue As String)
    mstrIntraCanalText = strValue
    UpdateDirtyFlag
End Property

Public Property Get Cotton_Pellet() As ExamEndoYesNoEnum
    Cotton_Pellet = mlngCottenPellet
    UpdateDirtyFlag
End Property
Public Property Let Cotton_Pellet(ByVal lngValue As ExamEndoYesNoEnum)
    mlngCottenPellet = lngValue
    UpdateDirtyFlag
End Property

Public Property Get Material_Provision() As TreatEndo_MaterialEnum
    Material_Provision = mlngProvMaterial
End Property
Public Property Let Material_Provision(ByVal lngValue As Long)
    mlngProvMaterial = lngValue
    UpdateDirtyFlag
End Property
Public Property Get Material_Provision_Text() As String
    Material_Provision_Text = mstrProvMaterialText
End Property
Public Property Let Material_Provision_Text(ByVal strValue As String)
    mstrProvMaterialText = strValue
    UpdateDirtyFlag
End Property

Public Property Get Cement() As TreatEndo_CementEnum
    Cement = mlngCement
End Property

Public Property Let Cement(ByVal lngValue As TreatEndo_CementEnum)
    mlngCement = lngValue
    UpdateDirtyFlag
End Property

Public Property Get Cement_Text() As String
    Cement_Text = mstrCementText
End Property

Public Property Let Cement_Text(ByVal strValue As String)
    mstrCementText = strValue
    UpdateDirtyFlag
End Property

Public Property Get Occulusal_Adjustment() As TreatEndo_OccAdjEnum
    Occulusal_Adjustment = mlngOcculusalAdj
End Property

Public Property Let Occulusal_Adjustment(ByVal lngValue As TreatEndo_OccAdjEnum)
    mlngOcculusalAdj = lngValue
    UpdateDirtyFlag
End Property

Public Property Get DentalEncounterID() As Long
    DentalEncounterID = mlngDentalEncounterId
End Property
Public Property Let DentalEncounterID(ByVal lngValue As Long)
    mlngDentalEncounterId = lngValue
End Property

Public Property Get DirtyFlag() As DirtyFlagEnum
    DirtyFlag = mlngDirtyFlag
End Property

Public Property Let DirtyFlag(ByVal lngValue As DirtyFlagEnum)
    mlngDirtyFlag = lngValue
End Property
'*******************************************************************
'Function: GetNote
'Author: Kieran Dunne
'Date: 11/08/2001
'Purpose: Builds a narrative summary of the user's selections.
'*******************************************************************
Public Function GetNote() As String
    Dim strReturn As String
    Dim strTemp As String
    
    If mstrToothNumber = "" Then
        'Our documentation here is based off a tooth number. If we have no tooth
        'number, we have no documentation. Return blank.
        strReturn = Empty
        GoTo FuncExit
    Else
        strReturn = strReturn & " Tooth: " & mstrToothNumber & ";"
    End If
    
    strTemp = Get_MultiSelect_DisplayString(mstrSolvent)
    
    If strTemp <> "" Then
        strReturn = strReturn & " Irrigation/Solvents/Lubricants: " & _
            strTemp & ";"
    End If
    
    strTemp = Get_MultiSelect_DisplayString(mstrMagnification)
    
    If strTemp <> "" Then
        strReturn = strReturn & "Magnification: " & strTemp & ";"
    End If
    
    If mlngIntraCanalMed <> TreatEndo_IntracanalMedicantEnum.Not_Addressed Then
        strReturn = strReturn & "Intracanal medicament: " & _
            TreatNSE_GetIntraCanalMedicantEnumDesc(mlngIntraCanalMed, mstrIntraCanalText) & ";"
    End If
    
    Select Case mlngCottenPellet
        Case ExamEndoYesNoEnum.Yes
            strReturn = strReturn & " Cotton Pellet was used." & ";"
        Case ExamEndoYesNoEnum.No
            strReturn = strReturn & " Cotton Pellet was not used." & ";"
    
    End Select
    
    If mlngProvMaterial <> TreatEndo_MaterialEnum.Not_Addressed Then
        strReturn = strReturn & " Material of the provisional was " & _
            TreatNSE_GetProvTreatMaterialEnumDesc(mlngProvMaterial, mstrProvMaterialText) & ";"
    End If
    
    strTemp = TreatNSE_GetCementEnumDesc(mlngCement, mstrCementText)
    If strTemp <> "" Then
        strReturn = strReturn & " Cement: " & _
            strTemp & ";"
            
    End If
    
    If mlngOcculusalAdj <> TreatEndo_OccAdjEnum.Not_Addressed Then
        strReturn = strReturn & " Occlusal adjustment: " & _
            TreatNSE_GetOcclusalAdjustmentEnumDesc(mlngOcculusalAdj) & ";"
    End If
    
    
FuncExit:
    GetNote = strReturn
 
End Function


'*******************************************************************
'Function: SaveSection
'Author: Kieran Dunne
'Date: 11/08/2001
'Purpose: Saves users selections.
'*******************************************************************
Public Function SaveSection() As Boolean
    Dim strSql As String
    Dim blnReturn As Boolean
        
    Select Case mlngDirtyFlag
        Case dfNoAction
            blnReturn = True
        Case dfAdd
            'First make sure that our object contains data
            If ValidateSave Then
                'Get the primary key
                mlngId = goconn.GetNextID(esTEMPL_TREAT_NSE_TOOTH_INFO_SEQ)
'                strSql = "SELECT TEMPL_TREAT_NSE_TOOTH_INFO_seq.NEXTVAL From Dual"
'                Set objRS = goconn.OpenRecordset(strSql)
'                mlngId = objRS("NEXTVAL")
                
                strSql = "INSERT INTO TEMPL_TREAT_NSE_TOOTH_INFO " & _
                        " (ET_TOOTH_INFO_ID, DENTAL_ENCOUNTER_ID, " & _
                        " TOOTH_NUMBER, LUBRICANTS, MAGINIFICATION, " & _
                        " INTRACANAL_MEDICANT, COTTON_PELLET, MATERIAL, " & _
                        "  CEMENT, OCCUSAL_ADJ, CEMENT_TEXT, MATERIAL_TEXT, ICM_TEXT) " & _
                        " VALUES (" & mlngId & ", " & mlngDentalEncounterId & "," & _
                        " '" & mstrToothNumber & "', '" & mstrSolvent & "', '" & mstrMagnification & _
                        "', '" & mlngIntraCanalMed & "', " & mlngCottenPellet & ", " & _
                        mlngProvMaterial & ", '" & mlngCement & "', " & mlngOcculusalAdj & ", '" & _
                        mstrCementText & "', '" & mstrProvMaterialText & "', '" & _
                        mstrIntraCanalText & "')"
            Else
                strSql = Empty
            End If 'ValidateSave
        Case dfUpdate
            If ValidateSave Then
                strSql = "Update TEMPL_TREAT_NSE_TOOTH_INFO " & _
                         " SET TOOTH_NUMBER = '" & mstrToothNumber & "', LUBRICANTS = '" & _
                         mstrSolvent & "', MAGINIFICATION = '" & mstrMagnification & "', " & _
                         "INTRACANAL_MEDICANT = '" & mlngIntraCanalMed & "', COTTON_PELLET = " & _
                         mlngCottenPellet & ", MATERIAL = " & mlngProvMaterial & ", OCCUSAL_ADJ = " & _
                         mlngOcculusalAdj & ", CEMENT = '" & mlngCement & "', CEMENT_TEXT = '" & mstrCementText & _
                         "', MATERIAL_TEXT = '" & mstrProvMaterialText & "', ICM_TEXT = '" & _
                         mstrIntraCanalText & "' WHERE ET_TOOTH_INFO_ID = " & mlngId
            End If 'ValidateSave
                     
        Case dfdelete
            If mlngId > 0 Then
                strSql = "DELETE FROM TEMPL_TREAT_NSE_TOOTH_INFO WHERE " & _
                         " ET_TOOTH_INFO_ID = " & mlngId
            End If
    End Select
        
    If strSql <> "" Then
        goconn.ExecuteSQL (strSql)
    End If
    mlngDirtyFlag = dfNoAction

    
    SaveSection = True
End Function

'******************************************************************************
'Function: ValidateSave
'Author: KDunne
'Date: 11/30/2001
'Purpose: Makes sure that our current record contains enough information to
'           be saved.
'*******************************************************************************
Private Function ValidateSave() As Boolean
    Dim blnReturn As Boolean
    
    blnReturn = False
    
    If Trim(mstrToothNumber) = "" Then
        'We must have a tooth number to save the record.
        ValidateSave = False
        Exit Function
    End If
    
    'Now, make sure that we have at least one additional field
    'before we can save the record.
    
    If mlngCement <> TreatEndo_CementEnum.Not_Addressed Then
        blnReturn = True
    End If
    
    If mlngCottenPellet <> ExamEndoYesNoEnum.Not_Addressed Then
        blnReturn = True
    End If
    
    If mlngIntraCanalMed <> TreatEndo_IntracanalMedicantEnum.Not_Addressed Then
        blnReturn = True
    End If
    
    If mlngOcculusalAdj <> TreatEndo_OccAdjEnum.Not_Addressed Then
        blnReturn = True
    End If
    
    If mlngProvMaterial <> TreatEndo_MaterialEnum.Not_Addressed Then
        blnReturn = True
    End If
    
    If mlngSolvent <> TreatEndo_RCTSolventsEnum.Not_Addressed Then
        blnReturn = True
    End If
    
    If Trim(Get_MultiSelect_DisplayString(mstrMagnification)) <> "" Then
        blnReturn = True
    End If
    
    
    ValidateSave = True
End Function
'****************************************************
'Sub: UpdateDirtyFlag
'Author: KDunne
'Date: 11/5/2001
'Purpose: Updates module level dirty flag field
'****************************************************
Public Sub UpdateDirtyFlag()
    If mlngDirtyFlag = dfNoAction Then
        If mlngId = 0 Then
            mlngDirtyFlag = dfAdd
        Else
            mlngDirtyFlag = dfUpdate
        End If
    End If

End Sub


