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
Attribute VB_Name = "clsExamEndoPain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Implements clsISection

Private mstrLocation As String
Private mstrOnset As String
Private mblnSpontaneous As Boolean
Private mstrQuality As String
Private mstrProvocation As String
Private mstrRelieved As String
Private mstrFrequency As String
Private mintIntensity As Integer
Private mblnHxTrauma As Boolean
Private mstrComments As String
Private mlngDirtyFlag As DirtyFlagEnum
Private mlngSubjId As Long
Private mlngPainCharID As Long

Private Sub Class_Initialize()
    mlngDirtyFlag = dfNoAction
    mstrQuality = "Sharp|Dull|Stabbing|Throbbing|"
    mstrProvocation = "Heat|Cold|Biting|Air|Lying down|Sweets|Bending over"
    mstrRelieved = "Heat|Cold|Biting|Air|Lying down|Bending over"
End Sub
Public Property Get Pain_Characteristic_ID() As Long
    Pain_Characteristic_ID = mlngPainCharID
End Property
Public Property Let Pain_Characteristic_ID(ByVal lngValue As Long)
    mlngPainCharID = lngValue
    UpdateDirtyFlag
End Property
Public Property Get Subjective_ID() As Long
    Subjective_ID = mlngSubjId
End Property
Public Property Let Subjective_ID(ByVal lngValue As Long)
    mlngSubjId = lngValue
End Property
Public Property Get DirtyFlag() As DirtyFlagEnum
    DirtyFlag = mlngDirtyFlag
End Property
Public Property Let DirtyFlag(ByVal lngValue As DirtyFlagEnum)
    mlngDirtyFlag = lngValue
End Property

Public Property Get Location() As String
    Location = mstrLocation
End Property

Public Property Let Location(ByVal strLocation As String)
    mstrLocation = strLocation
    UpdateDirtyFlag
End Property

Public Property Get Onset() As String
    Onset = mstrOnset
End Property
Public Property Let Onset(ByVal strValue As String)
    mstrOnset = strValue
    UpdateDirtyFlag
End Property

Public Property Get Spontaneous() As Boolean
    Spontaneous = mblnSpontaneous
    UpdateDirtyFlag
End Property
Public Property Let Spontaneous(ByVal blnValue As Boolean)
    mblnSpontaneous = blnValue
    UpdateDirtyFlag
End Property
Public Property Get PainQuality() As String
    PainQuality = mstrQuality
End Property
Public Property Let PainQuality(ByVal strQuality As String)
    mstrQuality = strQuality
    UpdateDirtyFlag
End Property
Public Property Get ProvokedBy() As String
    ProvokedBy = mstrProvocation
End Property
Public Property Let ProvokedBy(ByVal strValue As String)
    mstrProvocation = strValue
    UpdateDirtyFlag
End Property
Public Property Get RelievedBy() As String
    RelievedBy = mstrRelieved
End Property
Public Property Let RelievedBy(ByVal strValue As String)
    mstrRelieved = strValue
    UpdateDirtyFlag
End Property
Public Property Get Frequency() As String
    Frequency = mstrFrequency
End Property
Public Property Let Frequency(ByVal strValue As String)
    mstrFrequency = strValue
    UpdateDirtyFlag
End Property

Public Property Get Intensity() As Integer
    Intensity = mintIntensity
End Property
Public Property Let Intensity(ByVal intvalue As Integer)
    mintIntensity = intvalue
    UpdateDirtyFlag
End Property
Public Property Let HxTrauma(ByVal blnValue As Boolean)
    mblnHxTrauma = blnValue
    UpdateDirtyFlag
End Property
Public Property Get HxTrauma() As Boolean
    HxTrauma = mblnHxTrauma
End Property

Public Property Get Comments() As String
    Comments = mstrComments
End Property
Public Property Let Comments(ByVal strValue As String)
    mstrComments = strValue
    UpdateDirtyFlag
End Property

Public Function GetNote() As String
    Dim objToothCtrl As CToothControl
    Dim arrValues() As String
    Dim strLocation As String
    Dim strOnset As String
    Dim strSpont As String
    Dim strHxTrauma As String
    Dim strReturn As String
    Dim strTemp As String
    
    Dim X As Integer
    
    Set objToothCtrl = New CToothControl
    
    
    strReturn = Empty
    strLocation = Empty
    strOnset = Empty
    strSpont = Empty
    strHxTrauma = Empty
    
    If mstrLocation <> "" And mstrLocation <> "100" Then
        
        arrValues = Split(mstrLocation, "|")
        If arrValues(0) >= 100 Then
            strLocation = objToothCtrl.Get_Region_Name(arrValues(0))
        Else
            For X = 0 To UBound(arrValues)
                strLocation = strLocation & ", " & objToothCtrl.Get_Tooth_Name(arrValues(X))
            Next
            
            strLocation = Replace(strLocation, ", ", "", , 1)
            
            If X > 1 Then
                strLocation = "Teeth: " & strLocation
            Else
                strLocation = "Tooth:" & strLocation
            End If
        End If
        
        If strLocation <> "" Then
            strReturn = vbCrLf & "Location: " & strLocation & ";"
        End If
    End If
        
    If mstrOnset <> "" Then
    
        If IsDate(mstrOnset) And mstrOnset <> "12/31/9999" Then
            strReturn = strReturn & " Onset: " & Format(mstrOnset, cCHCSIIDateFormat) & ";" & vbCrLf
        Else
            strReturn = strReturn & "Onset: " & mstrOnset & ";"
        End If
        
    End If
    
    If mblnSpontaneous Then
        strReturn = strReturn & " Pain is spontaneous;"
    End If
        
    strTemp = Get_MultiSelect_DisplayString(mstrQuality)
    If Trim(strTemp) <> "" Then
        strReturn = strReturn & "Pain Quality: " & strTemp & ";"
    End If
    
    strTemp = Get_MultiSelect_DisplayString(mstrProvocation)
    If Trim(strTemp) <> "" Then
        strReturn = strReturn & " Provoked by: " & strTemp & ";"
    End If
    
    strTemp = Get_MultiSelect_DisplayString(mstrRelieved)
    If Trim(strTemp) <> "" Then
        strReturn = strReturn & " Relieved by: " & strTemp & ";"
    End If
    
    If Trim(mstrFrequency) <> "" Then
        strReturn = strReturn & " Frequency: " & mstrFrequency & ";"
    End If
    
    If mintIntensity > 0 Then
        'KDunne 9/4/2003
        'SCR 39287 - Modified intensity Description.
        strReturn = strReturn & "Intensity: " & ExamEndo_GetPainIntensity_Level(mintIntensity) & "/10 " & _
            Get_Intensity_Text(mintIntensity) & ";"
    End If
    
    If mblnHxTrauma Then
        strReturn = strReturn & " Has a history of trauma;"
    End If
    
    If Trim(mstrComments) <> "" Then
        strReturn = strReturn & " Comments: " & mstrComments & ";"
    End If
    
    strReturn = strReturn & vbCrLf
    
    GetNote = strReturn

    Set objToothCtrl = Nothing
    
End Function

'****************************************************
'Sub: UpdateDirtyFlag
'Author: KDunne
'Date: 11/5/2001
'Purpose: Updates module level dirty flag field
'****************************************************
Public Sub UpdateDirtyFlag()
    If mlngDirtyFlag = dfNoAction Then
        If mlngPainCharID = 0 Then
            mlngDirtyFlag = dfAdd
        Else
            mlngDirtyFlag = dfUpdate
        End If
    End If

End Sub

Public Function SaveSection() As Boolean
    Dim strSql As String
    Dim blnReturn As Boolean
    Dim strOnset As String
    Dim lSeq As Long
    
    'SF 39716
    If IsNumeric(mstrOnset) = True Then
    'If IsNumeric(mstrOnset) = False Then
        strOnset = ""
    ElseIf mstrOnset = "12/31/9999" Then
    'ElseIf mstrOnset = 0 Then
        strOnset = ""
    Else
        strOnset = Format(mstrOnset, cCHCSIIDateFormat) '"DD MMM YYYY")
    End If
        
    
    Select Case mlngDirtyFlag
        Case dfNoAction
            blnReturn = True
            GoTo FuncExit
        Case dfAdd
            'Get the primary key
            lSeq = goconn.GetNextID(esTEMPL_EXAM_ENDO_PAIN_CHARS_SEQ)
'            strSql = "SELECT templ_exam_endo_pain_chars_seq.NEXTVAL From Dual"
'            Set objRS = goconn.OpenRecordset(strSql)
            
            strSql = "INSERT INTO TEMPL_EXAM_ENDO_PAIN_CHARS " & _
                    " (EX_ENDO_PAIN_CHARS_ID, EX_ENDO_SUBJ_ID, LOCATION, " & _
                    " ONSET, SPONTANEOUS, PAIN_QUALITY, PROVOKED_BY, " & _
                    " RELIEVED_BY, FREQUENCY, INTENSITY, TRAUMA_HISTORY, " & _
                    " COMMENTS) " & _
                    " VALUES (" & lSeq & ", " & mlngSubjId & ", '" & mstrLocation & "', '" & strOnset & _
                    "', " & CInt(mblnSpontaneous) & ", '" & _
                    mstrQuality & "', '" & mstrProvocation & "', '" & mstrRelieved & _
                    "', '" & RQ(mstrFrequency) & "', " & mintIntensity & ", " & CInt(mblnHxTrauma) & _
                    ", '" & RQ(mstrComments) & "')"

        Case dfUpdate
            strSql = "Update TEMPL_EXAM_ENDO_PAIN_CHARS " & _
                     " SET LOCATION = '" & mstrLocation & "', ONSET = '" & strOnset & _
                     "', SPONTANEOUS = " & CInt(mblnSpontaneous) & ", PAIN_QUALITY = '" & mstrQuality & _
                     "', PROVOKED_BY = '" & mstrProvocation & "', RELIEVED_BY = '" & mstrRelieved & _
                     "', FREQUENCY = '" & mstrFrequency & "', INTENSITY = " & mintIntensity & _
                     ", TRAUMA_HISTORY = " & CInt(mblnHxTrauma) & ", COMMENTS = '" & _
                     RQ(mstrComments) & "' WHERE EX_ENDO_PAIN_CHARS_ID = " & mlngPainCharID
                     
        Case dfdelete
            strSql = "DELETE FROM TEMPL_EXAM_ENDO_PAIN_CHARS WHERE " & _
                     " EX_ENDO_PAIN_CHARS_ID = " & mlngPainCharID
    End Select
        
    If strSql <> "" Then
        goconn.ExecuteSQL (strSql)
    End If
    mlngDirtyFlag = dfNoAction

FuncExit:
    SaveSection = True
    
    Exit Function
End Function
'*******************************************************************
'Function: Get_Intensity_Text
'Author: KDunne
'Date: 11/1/2001
'Returns: String.
'Purpose: Returns the short description for a given intensity level.
'*******************************************************************
Public Function Get_Intensity_Text(ByVal intLevel As PainIntensity) As String
    Dim strDesc As String
    
    Select Case intLevel
        Case NoPain
            strDesc = "No Pain"
        Case Level_01
            strDesc = "Mild"
        Case Level_02
            strDesc = "Mild"
        Case Level_03
            strDesc = "Mild"
        Case Level_04
            strDesc = "Moderate"
        Case Level_05
            strDesc = "Moderate"
        Case Level_06
            strDesc = "Moderate"
        Case Level_07
            strDesc = "Severe"
        Case Level_08
            strDesc = "Severe"
        Case Level_09
            strDesc = "Severe"
        Case Level_10
            strDesc = "Worst Possible Pain"
    
    End Select
    
    Get_Intensity_Text = strDesc
End Function

Public Sub Clear()
    mstrLocation = Empty
    mstrOnset = Empty
    mblnSpontaneous = False
    mstrFrequency = Empty
    mintIntensity = 0
    mblnHxTrauma = False
    mstrComments = Empty
    mstrQuality = "Sharp|Dull|Stabbing|Throbbing|"
    mstrProvocation = "Heat|Cold|Biting|Air|Lying down|Sweets|Bending over"
    mstrRelieved = "Heat|Cold|Biting|Air|Lying down|Bending over"

End Sub
