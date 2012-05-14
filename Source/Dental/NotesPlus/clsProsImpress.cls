VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProsImpress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mlngDentalEncounterId
Private mlngId As Long
Private mlngDirtyFlag As DirtyFlagEnum

Private mlngLocation As ProsImpress_Location
Private mlngPurpose As ProsImpress_Purpose
Private mlngNumMade As Long
Private mlngMaterial As ProsImpress_Material

Private mlWorkloadID As Long

Private Const const_ModuleName = "clsProsImpress"

Public Enum ProsImpress_Location
    piLoc_Not_Addressed
    piLoc_Maxillary
    piLoc_Mandibular
End Enum

Public Enum ProsImpress_Purpose
    piPur_Not_Addressed
    piPur_Oppossing
    piPur_Bleaching
    piPur_Diagnostic
    piPur_Final
End Enum

Public Enum ProsImpress_Material
    piMat_Not_Addressed
    piMat_Alginate
    piMat_Plastic
    piMat_Polysulfide
    piMat_Polyether
    piMat_Polyvinylsiloxane
End Enum

Public Property Let WorkloadId(vData As Long)
    
    mlWorkloadID = vData
    
End Property

Public Property Get WorkloadId() As Long

    WorkloadId = mlWorkloadID
    
End Property


Private Sub Class_Initialize()
    mlngDirtyFlag = dfNoAction
    mlngId = 0
End Sub

Public Property Get ImpressionId() As Long
    ImpressionId = mlngId
End Property

Public Property Let ImpressionId(ByVal lngValue As Long)
    mlngId = lngValue
End Property
Public Property Get Dental_EncounterId() As Long
    Dental_EncounterId = mlngDentalEncounterId
End Property

Public Property Let Dental_EncounterId(ByVal lngValue As Long)
    mlngDentalEncounterId = lngValue
End Property

Public Property Get Location() As ProsImpress_Location
    Location = mlngLocation
End Property

Public Property Let Location(ByVal lngValue As ProsImpress_Location)
    mlngLocation = lngValue
    UpdateDirtyFlag
End Property

Public Property Get Purpose() As ProsImpress_Purpose
    Purpose = mlngPurpose
End Property

Public Property Let Purpose(ByVal lngValue As ProsImpress_Purpose)
    mlngPurpose = lngValue
    UpdateDirtyFlag
End Property

Public Property Get Number_Made() As Long
    Number_Made = mlngNumMade
End Property

Public Property Let Number_Made(ByVal lngValue As Long)
    mlngNumMade = lngValue
    UpdateDirtyFlag
End Property

Public Property Get Material() As ProsImpress_Material
    Material = mlngMaterial
End Property

Public Property Let Material(ByVal lngValue As ProsImpress_Material)
    mlngMaterial = lngValue
    UpdateDirtyFlag
End Property

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

'*******************************************************************
'Function: SaveSection
'Author: Kieran Dunne
'Date: 11/08/2001
'Purpose: Saves users selections.
'*******************************************************************
Public Function SaveSection() As Boolean
    Dim strSql As String
    Dim blnReturn As Boolean
        
    On Error GoTo ErrHandler
    
    Select Case mlngDirtyFlag
        Case dfNoAction
            blnReturn = True
            GoTo FuncExit
        Case dfAdd
            If mlngLocation <> piLoc_Not_Addressed Then
                mlngId = goconn.GetNextID(esTEMPL_GENPROS_IMPRESS_SEQ)
'                strSql = "SELECT Templ_GENPROS_IMPRESS_seq.NEXTVAL From Dual"
'                Set objRS = goconn.OpenRecordset(strSql)
'                mlngId = objRS("NEXTVAL")
                
                strSql = "INSERT INTO Templ_GENPROS_IMPRESS " & _
                        " (IMPRESSION_ID, DENTAL_ENCOUNTER_ID, LOCATION, " & _
                        " NUM_MADE, MATERIAL, PURPOSE, Workload_ID) " & _
                        " VALUES (" & mlngId & ", " & mlngDentalEncounterId & "," & _
                        " " & mlngLocation & ", " & mlngNumMade & ", " & _
                        mlngMaterial & ", " & mlngPurpose & ", " & mlWorkloadID & ")"
            Else
                blnReturn = True
                GoTo FuncExit
            End If
        Case dfUpdate
            strSql = "Update Templ_GENPROS_IMPRESS " & _
                     " SET LOCATION = " & mlngLocation & ", NUM_MADE = " & mlngNumMade & _
                     ", MATERIAL = " & mlngMaterial & ", PURPOSE =  " & mlngPurpose & _
                     ", Workload_ID = " & mlWorkloadID & _
                     " WHERE IMPRESSION_ID = " & mlngId
                     
        Case dfdelete
            strSql = "DELETE FROM Templ_GENPROS_IMPRESS WHERE " & _
                     " EX_ENDO_ADD_FIND_ID = " & mlngId
    End Select
        
    goconn.ExecuteSQL (strSql)
    mlngDirtyFlag = dfNoAction
    blnReturn = True
    
FuncExit:
    
    SaveSection = blnReturn
    Exit Function

ErrHandler:
    blnReturn = False
    goCWShared.ShowVBError Err.Number, Err.Description, const_ModuleName & ".GetCariesStateEnumDesc()", "", vbCritical
    GoTo FuncExit
End Function

'*******************************************************************
'Function: GetRTFNote
'Author: Kieran Dunne
'Date: 1/07/2001
'Purpose: Builds a narrative summary of the user's selections.
'*******************************************************************
Public Function GetRTFNote() As String
    Dim strNote As String
    
    If mlngLocation <> piLoc_Not_Addressed Then
        strNote = strNote & "Location: " & _
            GetProsImpressLocationEnumDesc(mlngLocation) & "; "
    End If
    
    If mlngPurpose <> piPur_Not_Addressed Then
        strNote = strNote & "Purpose: " & _
            GetProsImpressPurposeEnumDesc(mlngPurpose) & "; "
    End If
    
    If mlngNumMade > 0 Then
        strNote = strNote & "Number made: " & mlngNumMade & "; "
    End If
    
    If mlngMaterial <> piMat_Not_Addressed Then
        strNote = strNote & "Material: " & _
            GetProsImpressMaterialEnumDesc(mlngMaterial) & "; "
    End If
        
    GetRTFNote = strNote
End Function

'******************************************************************************
'Function: GetProsImpressLocationEnumDesc
'Author: KDunne
'Date: 1/7/2002
'Purpose: Returns the description for a given ProsImpress_Location enum value
'*******************************************************************************
Public Function GetProsImpressLocationEnumDesc(ByVal lngValue As ProsImpress_Location) As String
    Dim strReturn As String
    
    Select Case lngValue
        Case ProsImpress_Location.piLoc_Not_Addressed
            strReturn = ""
        Case ProsImpress_Location.piLoc_Maxillary
            strReturn = "Maxillary"
        Case ProsImpress_Location.piLoc_Mandibular
            strReturn = "Mandibular"
        Case Else
            strReturn = Empty
        
    End Select
    
    GetProsImpressLocationEnumDesc = strReturn
End Function

'******************************************************************************
'Function: GetProsImpressPurposeEnumDesc
'Author: KDunne
'Date: 1/7/2002
'Purpose: Returns the description for a given ProsImpress_Purpose enum value
'*******************************************************************************
Public Function GetProsImpressPurposeEnumDesc(ByVal lngValue As ProsImpress_Purpose) As String
    Dim strReturn As String
    
    Select Case lngValue
        Case ProsImpress_Purpose.piPur_Not_Addressed
            strReturn = Empty
        Case ProsImpress_Purpose.piPur_Oppossing
            strReturn = "Opposing"
        Case ProsImpress_Purpose.piPur_Bleaching
            strReturn = "Bleaching Tray"
        Case ProsImpress_Purpose.piPur_Diagnostic
            strReturn = "Diagnostic"
        Case ProsImpress_Purpose.piPur_Final
            strReturn = "Final"
        Case Else
            strReturn = Empty
        
    End Select
    
    GetProsImpressPurposeEnumDesc = strReturn
End Function

'******************************************************************************
'Function: GetProsImpressMaterialEnumDesc
'Author: KDunne
'Date: 1/7/2002
'Purpose: Returns the description for a given ProsImpress_Material enum value
'*******************************************************************************
Public Function GetProsImpressMaterialEnumDesc(ByVal lngValue As ProsImpress_Material) As String
    Dim strReturn As String
    
    Select Case lngValue
        Case ProsImpress_Material.piMat_Not_Addressed
            strReturn = Empty
        Case ProsImpress_Material.piMat_Alginate
            strReturn = "Alginate"
        Case ProsImpress_Material.piMat_Plastic
            strReturn = "Modeling Plastic"
        Case ProsImpress_Material.piMat_Polysulfide
            strReturn = "Polysulfide"
        Case ProsImpress_Material.piMat_Polyether
            strReturn = "Polyether"
        Case ProsImpress_Material.piMat_Polyvinylsiloxane
            strReturn = "Polyvinylsiloxane"
        Case Else
            strReturn = Empty
        
    End Select
    
    GetProsImpressMaterialEnumDesc = strReturn
End Function

'**********************************************************************
'Sub:
'Author:KDunne
'Date: 1/7/2002
'Purpose: Clears the Dirty Flag
'***********************************************************************
Public Sub Clear_Dirty_Flag()
    mlngDirtyFlag = dfNoAction
End Sub
