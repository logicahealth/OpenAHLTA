VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TemplateLite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const NCID_AppointmentType As String = "1021"
Private Const NCID_Complaint As String = "31048"
Private Const NCID_Problem As String = "261"

Private msName As String
Private msOwnerNCID As String
Private msOwnerName As String
Private msSourceNCID As String

Private oSpecialtyDictionary As Scripting.Dictionary
Private mbIsFavorite As Boolean
Private mbIsShared As Boolean
Private mbIsActive As Boolean
Private meOwnerType As TemplateOwnerTypeEnum
Private msFavoriteFolderID As String
Private msTemplateID As String
Private msSpecialtyID As String
Private msCPG_ID As String
Private msTemplateFacilityNCID As String

Private Function GetTemplateCPGXRef(pTemplateID As String) As String
    
    Dim moConn As CHCSII_CONN.Conn
    Dim oSQL As ICHCSII_SQL.ISqlOps
    Dim sSQL As String
    
    On Error GoTo ErrHandler
    Set moConn = New CHCSII_CONN.Conn
    Set oSQL = moConn.CHCSII_SQLOPS(Auto)
    GetTemplateCPGXRef = "0"
    
    sSQL = "SELECT * FROM CPG_TEMPLATE_XREF WHERE TEMPLATE_ID = " & pTemplateID
    oSQL.Execute sSQL
    
    If Not (oSQL.BOF And oSQL.EOF) Then
       If Not oSQL("CPG_ID") = "" Then
          GetTemplateCPGXRef = oSQL("CPG_ID")
       End If
    End If
    
    GoTo CleanUpAndExit
    
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateLite.GetTemplateCPGXRef", Err.Description

CleanUpAndExit:
    Set oSQL = Nothing
    Set moConn = Nothing
    
End Function

Public Function IsFavoriteTemplate(pFavorite_User As String) As Boolean
    
   On Error GoTo ErrHandler
    If pFavorite_User = cwobjLogonEnviron.UserNCID Then mbIsFavorite = True
    IsFavoriteTemplate = mbIsFavorite
    Exit Function
    
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "TemplateLite.IsFavoriteTemplate", "TemplateClient", vbExclamation)
    
End Function

Public Property Get StandaloneOrderSet() As Boolean
'scr-44211, scr-44212 RBELL 2/11/2004 Returns boolean indicating if the template object represents a standalone order set
    On Error GoTo ErrHandler
    StandaloneOrderSet = IsStandaloneOrderSet(msTemplateID)
    Exit Property

ErrHandler:
    If Err = 5 Then
       StandaloneOrderSet = False
       Exit Property
    End If
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateLite.StandaloneOrderSet", Err.Description
    
End Property

Public Property Get SpecialtyDictionary() As Scripting.Dictionary
    
    On Error GoTo ErrHandler
    Set SpecialtyDictionary = oSpecialtyDictionary
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateLite.SpecialtyDictionary", Err.Description
    
End Property

Public Property Set SpecialtyDictionary(ByVal RHV As Scripting.Dictionary)
    
    On Error GoTo ErrHandler
    Set oSpecialtyDictionary = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateLite.SpecialtyDictionary", Err.Description
    
End Property

Public Property Get TemplateType() As TemplateTypeEnum
    'jrm:Replacing mid w/ Left to check the prefix.  Mid returns whole name of template never the prefix.
    
    On Error GoTo ErrHandler
    Select Case UCase(Left(msName, 1))
    Case EncTemplatePrefix
        TemplateType = EncounterTemplateType
    Case ListTemplatePrefix
        TemplateType = ListTemplateType
    Case FormTemplatePrefix 'emc Forms support
        TemplateType = FormTemplateType
    End Select

    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateLite.TemplateType", Err.Description
    
End Property

Public Property Get TemplateName() As String
    
    On Error GoTo ErrHandler
    TemplateName = Mid(msName, 2)
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateLite.TemplateName", Err.Description
    
End Property

Public Property Get OwnerNCID() As String
    
    On Error GoTo ErrHandler
    OwnerNCID = msOwnerNCID
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateLite.OwnerNCID", Err.Description
    
End Property

Public Property Get OwnerName() As String
    
    On Error GoTo ErrHandler
    
    If gobjShared Is Nothing Then Set gobjShared = New CWShared
        
    OwnerName = GEMS_OwnerName
    
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateLite.OwnerName", Err.Description
    
Cleanup:

End Property
Public Property Get GEMS_OwnerName() As String
    
    Dim objConcept As GEMSConceptCtrl.GEMS_ConceptCtrl
    Dim objClinicOps As CHCSII_ClinicClient.ClinicOpsEx
    Dim objClinic As DTOs.Clinic
    
    On Error GoTo ErrHandler
    If msOwnerNCID <> "" And msOwnerName = "" Then
        If OwnerType = EnterpriseOwnerType Then
           msOwnerName = cwobjLogonEnviron.EnterpriseName
        ElseIf OwnerType = ClinicOwnerType And (gobjShared.IsAppMode(modeCDR) And Not gobjShared.IsAppMode(modeITT)) Then
        'SCR-52426 RBELL 4/15/2004 use concept control in theater mode
            On Error GoTo Cleanup
            Set objClinicOps = New CHCSII_ClinicClient.ClinicOpsEx
            Set objClinic = objClinicOps.GetClinic(msOwnerNCID)
            msOwnerName = objClinic.Name & " " & objClinic.MEPRSCode
        Else
            Set objConcept = New GEMSConceptCtrl.GEMS_ConceptCtrl
            objConcept.UniqueId = msOwnerNCID
            msOwnerName = objConcept.PrefRep("2000").Representation
        End If
    End If
    GEMS_OwnerName = msOwnerName
    
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateLite.GEMS_OwnerName", Err.Description
    
Cleanup:
    Set objClinicOps = Nothing
    Set objClinic = Nothing
    Set objConcept = Nothing
End Property
' GEMS >>

Public Property Get Source() As String
    
    On Error GoTo ErrHandler
    Select Case msSourceNCID
    Case NCID_AppointmentType
        Source = "Appt. Type"
    Case NCID_Complaint
        Source = "Rsn for Visit"
    Case NCID_Problem
        Source = "Problem List"
    End Select

    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateLite.Source", Err.Description
    
End Property

Public Property Get SpecialtyName() As String
    
    On Error GoTo ErrHandler
    SpecialtyName = SpecialtyDictionary(SpecialtyID)
    Exit Property

ErrHandler:
    If Err = 5 Then
       SpecialtyName = vbNullString
       Exit Property
    End If
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateLite.SpecialtyName", Err.Description
    
End Property

Friend Sub AdjustSource(robjSQL As ICHCSII_SQL.ISqlOps)
    
    On Error GoTo ErrHandler
    Select Case robjSQL("TYPE_NCID")
    Case NCID_AppointmentType
        ' do nothing
    Case NCID_Complaint
        msSourceNCID = NCID_Complaint
    Case NCID_Problem
        If msSourceNCID <> NCID_Complaint Then
            msSourceNCID = NCID_Problem
        End If
    End Select
    Exit Sub
    
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "TemplateLite.AdjustSource", "TemplateClient", vbExclamation)

End Sub

Friend Function Deserialize(robjSQL As ICHCSII_SQL.ISqlOps, pMTFs As ADODB.Recordset) As Boolean
    
    On Error GoTo ErrHandler
    TemplateID = robjSQL("ID")
    Select Case UCase(TypeName(robjSQL("FAVORITES_USER")))
        Case "STRING"
            If CStr(robjSQL("FAVORITES_USER")) & vbNullString = cwobjLogonEnviron.UserNCID Then
                IsFavorite = True
            Else
                IsFavorite = False
            End If
        Case "DECIMAL", "INTEGER", "LONG"
            If val(robjSQL("FAVORITES_USER")) = val(cwobjLogonEnviron.UserNCID) Then
                IsFavorite = True
            Else
                IsFavorite = False
            End If
            
        Case Else
            IsFavorite = False
    End Select
        
    msOwnerNCID = robjSQL("OWNER")
    IsActive = (robjSQL("ISACTIVE") = "Y")

    'msTemplateFacilityNCID = GetTemplateFacility(TemplateID)
    'Template Facility now returned by SQL query
    If Not IsNull(robjSQL("TEMPLATE_FACILITY")) Then msTemplateFacilityNCID = robjSQL("TEMPLATE_FACILITY")
    
    Select Case robjSQL("TYPE")
    Case tt_facility
        If gobjShared.IsAppMode(modeCDR) And Not gobjShared.IsAppMode(modeITT) Then 'SCR-52426 RBELL 4/15/2004 Only use pMTFs in CDR mode
           pMTFs.Filter = "NCID = " & msOwnerNCID
           If pMTFs.RecordCount > 0 Then
              OwnerType = MTFOwnerType
              IsShared = True 'If Not Share set to False SCR-17939
           Else
              OwnerType = ClinicOwnerType
              IsShared = True  'Is Shared Set to True SCR-17939
           End If
           pMTFs.Filter = ""
        Else
           If msTemplateFacilityNCID = msOwnerNCID Or (msTemplateFacilityNCID = vbNullString And msOwnerNCID = cwobjLogonEnviron.FacilityNCID) Then 'robjSQL("ISSHARED") = "Y" Then
               OwnerType = MTFOwnerType
               IsShared = True 'If Not Share set to False SCR-17939
           Else
               OwnerType = ClinicOwnerType
               IsShared = True  'Is Shared Set to True SCR-17939
           End If
        End If
        'IsShared = False 'Was defaulting to False SCR-17939
    Case tt_enterprise
        IsShared = True
        OwnerType = EnterpriseOwnerType
    Case Else
        IsShared = (robjSQL("ISSHARED") = "Y")
        OwnerType = PersonalOwnerType
    End Select
    msName = robjSQL("NAME")
    
    msFavoriteFolderID = robjSQL("FOLDER_ID") & vbNullString
    
    CPG_ID = "0"
    
    'SCR-46790 RBELL 12/16/03 If Not gobjShared.AppMode = modeCHCSI_GUI Then msSourceNCID = robjSQL("TYPE_NCID")
    Deserialize = True
    Exit Function
    
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "TemplateLite.Deserialize", "TemplateClient", vbExclamation)

End Function


Public Property Get IsFavorite() As Boolean
    
    On Error GoTo ErrHandler
    IsFavorite = mbIsFavorite
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateLite.IsFavorite", Err.Description
    
End Property

Public Property Let IsFavorite(ByVal RHV As Boolean)
    
    On Error GoTo ErrHandler
    mbIsFavorite = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateLite.IsFavorite", Err.Description
    
End Property

Public Property Get IsShared() As Boolean
    
    On Error GoTo ErrHandler
    IsShared = mbIsShared
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateLite.IsShared", Err.Description
    
End Property

Public Property Let IsShared(ByVal RHV As Boolean)
    
    On Error GoTo ErrHandler
    mbIsShared = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateLite.IsShared", Err.Description
    
End Property

Public Property Get IsActive() As Boolean
    
    On Error GoTo ErrHandler
    IsActive = mbIsActive
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateLite.IsActive", Err.Description
    
End Property

Public Property Let IsActive(ByVal RHV As Boolean)
    
    On Error GoTo ErrHandler
    mbIsActive = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateLite.IsActive", Err.Description
    
End Property

Public Property Get OwnerType() As TemplateOwnerTypeEnum
    
    On Error GoTo ErrHandler
    OwnerType = meOwnerType
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateLite.OwnerType", Err.Description
    
End Property

Public Property Let OwnerType(ByVal RHV As TemplateOwnerTypeEnum)
    
    On Error GoTo ErrHandler
    meOwnerType = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateLite.OwnerType", Err.Description
    
End Property

Public Property Get TemplateID() As String
    
    On Error GoTo ErrHandler
    TemplateID = msTemplateID
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateLite.TemplateID", Err.Description
    
End Property

Public Property Let TemplateID(ByVal RHV As String)
    
    On Error GoTo ErrHandler
    msTemplateID = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateLite.TemplateID", Err.Description
    
End Property

Public Property Get SpecialtyID() As String
    
    On Error GoTo ErrHandler
    SpecialtyID = msSpecialtyID
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateLite.SpecialtyID", Err.Description
    
End Property

Public Property Let SpecialtyID(ByVal RHV As String)
    
    On Error GoTo ErrHandler
    msSpecialtyID = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateLite.SpecialtyID", Err.Description
    
End Property

Public Property Get CPG_ID() As String
    
    On Error GoTo ErrHandler
    CPG_ID = msCPG_ID
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateLite.CPG_ID", Err.Description
    
End Property

Public Property Let CPG_ID(ByVal RHV As String)
    
    On Error GoTo ErrHandler
    msCPG_ID = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateLite.CPG_ID", Err.Description
    
End Property

Private Sub Class_Terminate()
    On Error Resume Next
    Set oSpecialtyDictionary = Nothing
End Sub



Public Property Get TemplateFacilityNCID() As String
    
    On Error GoTo ErrHandler
    TemplateFacilityNCID = msTemplateFacilityNCID
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateLite.TemplateFacilityNCID", Err.Description
    
End Property

Public Property Let TemplateFacilityNCID(ByVal RHV As String)
    
    On Error GoTo ErrHandler
    msTemplateFacilityNCID = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateLite.TemplateFacilityNCID", Err.Description
    
End Property


Public Property Get FavoriteFolderID() As String
    
    On Error GoTo ErrHandler
    FavoriteFolderID = msFavoriteFolderID
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateLite.FavoriteFolderID", Err.Description
    
End Property

Public Property Let FavoriteFolderID(ByVal RHV As String)
    
    On Error GoTo ErrHandler
    msFavoriteFolderID = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.TemplateLite.FavoriteFolderID", Err.Description
    
End Property
