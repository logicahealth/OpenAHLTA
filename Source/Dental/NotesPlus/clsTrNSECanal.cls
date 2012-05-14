VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTrNSECanal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mlngId As Long
Private mlngDentalEncounterId As Long
Private mstrToothNum As String
Private mlngCanalNum As Long
Private mdblWorkingLength As Double
Private mdblMAF As Double
Private mlngReferences As TreatEndo_ReferencesEnum
Private mstrInstrTech As String
Private mstrFileSystem As String
Private mlngObtTech As TreatEndo_ObturationTechEnum
Private mstrObtMater As String
Private mstrComments As String

Private mlngDirtyFlag As DirtyFlagEnum

Public Enum TreatEndo_ReferencesEnum
    Not_Addressed
    Mesiobuccal
    Mesiolingual
    DistalBuccal
    Distolingual
    Palatal
    Buccal
    Lingual
    Incisal
End Enum

Public Enum TreatEndo_ObturationTechEnum
    Not_Addressed
    ColdLateral
    WarmLateral
    WarmVertical
    CoreCarrier
End Enum

Private Sub Class_Initialize()
    mstrObtMater = cTreatNSEObturationMaterialsList
    mstrFileSystem = cTreatNSEFileSystemList
    mstrInstrTech = cTreatNSEIntrumentTechsList
    mdblWorkingLength = -1
    mdblMAF = -1
End Sub

Public Property Get Canal_Info_Id() As Long
    Canal_Info_Id = mlngId
End Property
Public Property Let Canal_Info_Id(ByVal lngValue As Long)
    mlngId = lngValue
End Property
Public Property Let DentalEncounterID(ByVal lngValue As Long)
    mlngDentalEncounterId = lngValue

End Property
Public Property Get DentalEncounterID() As Long
    DentalEncounterID = mlngDentalEncounterId
End Property
Public Property Get DirtyFlag() As DirtyFlagEnum
    DirtyFlag = mlngDirtyFlag
End Property
Public Property Let DirtyFlag(ByVal lngValue As DirtyFlagEnum)
    mlngDirtyFlag = lngValue
End Property

Public Property Get Canal() As ToothControlCanalsEnum
    Canal = mlngCanalNum
End Property
Public Property Let Canal(ByVal lngValue As ToothControlCanalsEnum)
    mlngCanalNum = lngValue
    UpdateDirtyFlag
End Property

Public Property Get Tooth_Number() As String
    Tooth_Number = mstrToothNum
End Property
Public Property Let Tooth_Number(ByVal strValue As String)
    mstrToothNum = strValue
    UpdateDirtyFlag
End Property

Public Property Get Working_Length() As Double
    Working_Length = mdblWorkingLength
End Property

Public Property Let Working_Length(ByVal dblValue As Double)
    mdblWorkingLength = dblValue
    UpdateDirtyFlag
End Property

Public Property Get Master_Apical_File() As Double
    Master_Apical_File = mdblMAF
End Property

Public Property Let Master_Apical_File(ByVal dblValue As Double)
    mdblMAF = CDbl(dblValue)
    UpdateDirtyFlag
End Property

Public Property Get References() As TreatEndo_ReferencesEnum
    References = mlngReferences
End Property
Public Property Let References(ByVal lngValue As TreatEndo_ReferencesEnum)
    mlngReferences = lngValue
    UpdateDirtyFlag
    
End Property

Public Property Get Instrumentation_Techniques() As String
    Instrumentation_Techniques = mstrInstrTech
End Property
Public Property Let Instrumentation_Techniques(ByVal strValue As String)
    mstrInstrTech = strValue
    UpdateDirtyFlag
End Property

Public Property Get File_System() As String
    File_System = mstrFileSystem
End Property
Public Property Let File_System(ByVal strValue As String)
    mstrFileSystem = strValue
    UpdateDirtyFlag
End Property

Public Property Get Obturation_Technique() As TreatEndo_ObturationTechEnum
    Obturation_Technique = mlngObtTech
End Property

Public Property Let Obturation_Technique(ByVal lngValue As TreatEndo_ObturationTechEnum)
    mlngObtTech = lngValue
    UpdateDirtyFlag
End Property

Public Property Get Obturation_Materials() As String
    Obturation_Materials = mstrObtMater
End Property
Public Property Let Obturation_Materials(ByVal strValue As String)
    mstrObtMater = strValue
    UpdateDirtyFlag
End Property

Public Property Get Comments() As String
    Comments = mstrComments
End Property
Public Property Let Comments(ByVal strValue As String)
    mstrComments = strValue
    UpdateDirtyFlag
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
    Dim objToothCtrl As CToothControl
    
    Set objToothCtrl = New CToothControl
    
    strReturn = Empty
    
    If mstrToothNum <> "" Then
        strReturn = strReturn & " Tooth Number is " & mstrToothNum & ";"
    End If
    
    If mlngCanalNum > 0 Then
        strReturn = strReturn & " Canal is " & objToothCtrl.Get_Canal_Description(mlngCanalNum) & ";"
    End If
    
    If mdblWorkingLength >= 0 Then
        strReturn = strReturn & " Working Length is " & FormatNumber(mdblWorkingLength, 1, vbTrue) & _
                    " millimeters" & ";"
    End If
    
    If mdblMAF >= 0 Then
        strReturn = strReturn & " Master Apical File is " & FormatNumber(mdblMAF, 1, vbTrue) & _
                    ";"
    End If
    
    If mlngReferences <> TreatEndo_ReferencesEnum.Not_Addressed Then
        strReturn = strReturn & " Reference = " & TreatNSE_GetTreatEndoReferencesEnumDesc(mlngReferences) & ";"
    End If
    
    strTemp = Get_MultiSelect_DisplayString(mstrInstrTech)
    
    If strTemp <> "" Then
        strReturn = strReturn & " Intrumentation Techniques are " & strTemp & ";"
    End If
    
    strTemp = Get_MultiSelect_DisplayString(mstrFileSystem)

    If strTemp <> "" Then
        strReturn = strReturn & " Files System = " & strTemp & ";"
    End If
    
    If mlngObtTech <> TreatEndo_ObturationTechEnum.Not_Addressed Then
        strReturn = strReturn & " Obturation Technique is " & _
            TreatNSE_GetTreatEndoObturationTechEnumDesc(mlngObtTech) & ";"
    End If
    
    strTemp = Get_MultiSelect_DisplayString(mstrObtMater)
    
    If strTemp <> "" Then
        strReturn = strReturn & " Obturation Materials: " & strTemp & ";"
    End If
    
    If Trim(mstrComments) <> "" Then
        strReturn = strReturn & "Comments: " & mstrComments & ";"
    End If
    
    GetNote = strReturn
End Function

'*******************************************************************
'Function: SaveSection
'Author: Kieran Dunne
'Date: 11/08/2001
'Purpose: Saves users selections.
'*******************************************************************
Public Function SaveSection() As Boolean
    Dim blnReturn As Boolean
    Dim strSql As String
    Dim objRS As ADODB.Recordset
    
    Dim sReferencesVar As String
    
    'ITT change
    If goChcsiiConn.UseSQLServer Then
        sReferencesVar = "[REFERENCES]"
    Else
        sReferencesVar = "REFERENCES"
    End If
    
    Select Case mlngDirtyFlag
        Case dfNoAction
            blnReturn = True
        Case dfAdd
            'First make sure that our object contains data
            If ValidateSave Then
                'Get the primary key
                mlngId = goconn.GetNextID(esTEMPL_TREAT_NSE_CANAL_INFO_SEQ)
'                strSql = "SELECT TEMPL_TREAT_NSE_CANAL_INFO_seq.NEXTVAL From Dual"
'                Set objRS = goconn.OpenRecordset(strSql)
'                mlngId = objRS("NEXTVAL")
                
                strSql = "INSERT INTO TEMPL_TREAT_NSE_CANAL_INFO " & _
                    "(ET_CANAL_INFO_ID, DENTAL_ENCOUNTER_ID, " & _
                    " TOOTH_NUMBER, TOOTH_CANAL, WORKING_LENGTH, MAF, " & _
                    " " & sReferencesVar & ", INTRUMENT_TECH, FILE_SYSTEM, " & _
                    " OBTURATION_TECH, OBTURATION_MAT, COMMENTS) " & _
                    " VALUES (" & mlngId & ", " & mlngDentalEncounterId & ", '" & _
                    mstrToothNum & "', " & mlngCanalNum & ", " & mdblWorkingLength & _
                    ", " & mdblMAF & ", " & mlngReferences & ", '" & mstrInstrTech & _
                    "', '" & mstrFileSystem & "', " & mlngObtTech & ",'" & mstrObtMater & "', '" & _
                    RQ(mstrComments) & "')"
            
            End If
        Case dfUpdate
            If ValidateSave Then
                'SCR 47580 Cannot have spaces before Update, Insert or Select statements
                strSql = "Update TEMPL_TREAT_NSE_CANAL_INFO " & _
                    " SET DENTAL_ENCOUNTER_ID = " & mlngDentalEncounterId & ", " & _
                " TOOTH_NUMBER = '" & mstrToothNum & "', TOOTH_CANAL = " & mlngCanalNum & _
                ", WORKING_LENGTH = " & mdblWorkingLength & ", MAF = " & mdblMAF & _
                ", " & sReferencesVar & " = " & mlngReferences & ", INTRUMENT_TECH = '" & mstrInstrTech & _
                "', FILE_SYSTEM = '" & mstrFileSystem & "', OBTURATION_TECH = " & _
                mlngObtTech & ", OBTURATION_MAT = '" & mstrObtMater & "', COMMENTS = '" & _
                RQ(mstrComments) & "' " & _
                " WHERE ET_CANAL_INFO_ID = " & mlngId
            End If
        Case dfdelete
            If mlngId > 0 Then
                strSql = "DELETE FROM TEMPL_TREAT_NSE_CANAL_INFO WHERE " & _
                         " ET_CANAL_INFO_ID = " & mlngId
            End If
        End Select

    If strSql <> "" Then
        goconn.ExecuteSQL (strSql)
    End If
    mlngDirtyFlag = dfNoAction
    
    SaveSection = blnReturn
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

'**************************************************************************
'Function: ValidateSave
'Author: KDunne
'Date: 12/4/2001
'Purpose: Makes sure that our object contains data before we save it.
'**************************************************************************
Public Function ValidateSave() As Boolean
    Dim blnReturn As Boolean
    
    blnReturn = False
    
    
    'We must have at least a tooth number and a canal.
    If mstrToothNum = "" Or mlngCanalNum = 0 Then
        ValidateSave = False
        Exit Function
    End If
    
    'We've got those two fields, now we need to make sure that we
    'have at least one other piece of data.
    If mdblWorkingLength > 0 Then
        blnReturn = True
    End If
    
    If mdblMAF > 0 Then
        blnReturn = True
    End If
    
    If mlngReferences <> TreatEndo_ReferencesEnum.Not_Addressed Then
        blnReturn = True
    End If
    
    If Get_MultiSelect_DisplayString(mstrInstrTech) <> "" Then
        blnReturn = True
    End If
    
    If Get_MultiSelect_DisplayString(mstrFileSystem) <> "" Then
        blnReturn = True
    End If
    
    If mlngObtTech <> TreatEndo_ObturationTechEnum.Not_Addressed Then
        blnReturn = True
    End If
    
    If Get_MultiSelect_DisplayString(mstrObtMater) <> "" Then
        blnReturn = True
    End If
    
    If Trim(mstrComments) <> "" Then
        blnReturn = True
    End If
    
    
    ValidateSave = blnReturn
    
End Function


