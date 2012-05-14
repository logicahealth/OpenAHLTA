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
Attribute VB_Name = "GCOps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim sqlString        As String
'<bjr082001 theater
'Private objDAS           As CHCSII_DAS
Private objDAS           As ICHCSII_DAS

Private objRS             As Recordset
Private objBigTooth As BigTooth

Private msDentalEncounterID   As String
Private msEncounterID   As String
Private msEncounterFacility   As String
Private msPatientUnitID   As String
Private objShared As CWShared
Private oLogon As InternalLogon
Private colTeeth As Collection
Private nDxTxView As Integer
Private bIsGCOPs As Boolean

Public objPSR As clsPSR

Private Function nextValSequence(tableName As String) As String
      Set objRS = objDAS.OpenRecordset("select " & tableName & ".nextval from dual")
      nextValSequence = objRS!nextval
End Function

Public Function getDxView(Optional hColor As Long = vbButtonFace) As Collection
Dim j As Integer
Dim I As Integer
Dim R, G, B As Integer
Dim hatch As Integer

 'Load frmGCOps
 'frmGCOps.Show
 frmGraphicalCharting.ReloadDxTx
 'svg Set getDxView = CopyColTeeth(frmGraphicalCharting.teethSVGX1.colTeeth)
'
'    With frmGCOps
'        For j = 1 To 32
'            .toothSVGX1.Initialize hColor
'            .toothSVGX1.currentViewTooth = j
''            For i = 1 To 32
''               R = Int((34 - 1 + 1) * Rnd + 1)
''               hatch = Int((5 - 1 + 1) * Rnd + 1)
''               .toothSVGX1.fillSurface CInt(i), CInt(R)
''             Next i
'
'             '.toothSVGX1.Overlay
'             '.toothSVGX1.SaveCache (j)
'             getDxView.Add frmGraphicalCharting.teethSVGX1.colTeeth(j).ipicScaled, (CStr(j))  '.toothSVGX1.currentScaledImgTooth(j), (CStr(j))
'
'          Next j
'    End With
 'Unload frmGCOps
 
End Function

Public Function getTxView(Optional hColor As Long = vbButtonFace) As Collection
Dim j As Integer
Dim I As Integer
Dim R, G, B As Integer
Dim hatch As Integer

 'Load frmGCOps
 'frmGraphicalCharting.cmdView(1).Value = True
 'svg Set getTxView = CopyColTeeth(frmGraphicalCharting.teethSVGX1.colTeeth)
 
'    With frmGCOps
'        For j = 1 To 32
'            '.toothSVGX1.Initialize hColor
'            '.toothSVGX1.currentViewTooth = j
''            For i = 1 To 32
''               R = Int((34 - 1 + 1) * Rnd + 1)
''               hatch = Int((5 - 1 + 1) * Rnd + 1)
''               .toothSVGX1.fillSurface CInt(i), CInt(R)
''             Next i
'
'             '.toothSVGX1.Overlay
'             '.toothSVGX1.SaveCache (j)
'             getTxView.Add frmGraphicalCharting.teethSVGX1.colTeeth(j).ipicScaled, (CStr(j))  '.toothSVGX1.currentScaledImgTooth(j), (CStr(j))
'
'          Next j
'    End With
' Unload frmGCOps
 
End Function

Private Function CopyColTeeth(OrigColTeeth As Collection) As Collection
    Dim I As Integer
    Dim objTeeth As clsTeeth
        
    Set colTeeth = New Collection
    For I = 1 To 32
        Set objTeeth = New clsTeeth
        'objTeeth.ipicScaled.Handle = OrigColTeeth(i).ipicScaled.Handle
        'objTeeth.ipicScaled.Height = OrigColTeeth(i).ipicScaled.Height
        'objTeeth.ipicScaled.hPal = OrigColTeeth(i).ipicScaled.hPal
        'objTeeth.ipicScaled.Type = OrigColTeeth(i).ipicScaled.Type
        'objTeeth.ipicScaled.Width = OrigColTeeth(i).ipicScaled.Width
        
        'svg Set objTeeth.ipicScaled = OrigColTeeth(I).ipicScaled
        'svg objTeeth.ToothID = OrigColTeeth(I).ToothID
        'svg objTeeth.ToothName = OrigColTeeth(I).ToothName
        'svg objTeeth.ToothComplimentID = OrigColTeeth(I).ToothComplimentID
        colTeeth.Add objTeeth, CStr(I)
    Next I
    Set CopyColTeeth = colTeeth
End Function

Public Function IsDxView() As Boolean
    IsDxView = frmGraphicalCharting.IsDxView
End Function
Public Function savePSR() As Boolean
    If objPSR Is Nothing Then Exit Function
          Call UpdatePSR(CInt(msDentalEncounterID))
End Function

Private Function UpdatePSR(DEncID As Integer) As Boolean
 sqlString = ""
'  With gobjDDD
'          sqlString = "Update " & .psDENTAL_ENCOUNTER & " Set " _
'                      & .psDENTAL_ENCOUNTER(lDE_PSR_LOWER_CENTER) & " = " & objPSR.lowerCenter & "," _
'                      & .psDENTAL_ENCOUNTER(lDE_PSR_LOWER_LEFT) & " = " & objPSR.lowerLeft & "," _
'                      & .psDENTAL_ENCOUNTER(lDE_PSR_LOWER_RIGHT) & " = " & objPSR.lowerRight & "," _
'                      & .psDENTAL_ENCOUNTER(lDE_PSR_UPPER_CENTER) & " = " & objPSR.upperCenter & "," _
'                      & .psDENTAL_ENCOUNTER(lDE_PSR_UPPER_LEFT) & " = " & objPSR.upperLeft & "," _
'                      & .psDENTAL_ENCOUNTER(lDE_PSR_UPPER_RIGHT) & " = " & objPSR.upperRight & " "
'          sqlString = sqlString & "Where " & .psDENTAL_ENCOUNTER(lDE_DENTAL_ENCOUNTER_ID) & " = " & DEncID
'       updatePSR = objDAS.ICHCSII_DAS_BeginTrans
'                       objDAS.ExecuteSQL (sqlString)
'       updatePSR = objDAS.ICHCSII_DAS_CommitTrans
'  End With
End Function

Public Function saveAll() As Boolean
  
  savePSR
 
End Function

Public Sub Initialize(PatUnitNo As String) 'robjComm As Object, robjLogonEnvironment As Object, robjConfig As Object, robjPatient As Object, robjUser As Object, robjEncounter As Object, rstrDentalEncounterID As String)

   On Error GoTo ErrHandler
         
'   Set gobjUser = robjUser
'
'   Dim objRoles      As ICHCSII_Roles
'   Dim objResources  As ICHCSII_Resources
'
'   '-- -create references to passed in objects---
'   Set cwobjCWComm = robjComm
'   Set cwobjLogonEnviron = robjLogonEnvironment
'   Set cwobjConfig = robjConfig
'   Set cwobjPatient = robjPatient
'   Set gobjEnc = robjEncounter
'   Set objShared = New CWShared
'   '-- -set global flags---
'   '-- -Initialize Variables ---
'
'   msDentalEncounterID = rstrDentalEncounterID
'   msEncounterID = gobjEnc.CurrentEncounter.EncounterID
'   msEncounterFacility = gobjEnc.CurrentEncounter.FacilityNCID
'   'msEncounterDate = gobjEnc.CurrentEncounter.en
   msPatientUnitID = PatUnitNo 'cwobjPatient.UnitNumber
   '<bjr082001 theater
   'Set objDAS = New CHCSII_DAS
   Set objDAS = GetDasObj()

   
   Exit Sub
ErrHandler:
   '-- -show the VB error---
   Call gobjShared.ShowVBError(Err.Number, Err.Description, "Dental_GCOps.Initialize", "BigTooth", vbExclamation)
   
   Exit Sub
End Sub

Public Function InitializeEx(objPatient As Object, objEncounter As Object) As Boolean
   
    'If objEncounter Is Nothing Then
    '    InitializeEx = False
    '    Exit Function
    'End If
    
    Set objBigTooth = frmGraphicalCharting.ParentObject
    If objBigTooth Is Nothing Then
        Set objBigTooth = New BigTooth
        InitializeEx = objBigTooth.createGCOPs(objPatient, objEncounter)
        bIsGCOPs = True
        gbIsGCOPs = True
    Else
        InitializeEx = True
    End If
   
End Function

Public Sub terminate_objects()
   'Set objPSR = Nothing
   'Set objDAS = Nothing
   'Set cwobjCWComm = Nothing
   'Set cwobjLogonEnviron = Nothing
   'Set cwobjConfig = Nothing
   'Set cwobjPatient = Nothing
   'Set gobjEnc = Nothing
   'Set objBigTooth = Nothing
   
End Sub

Public Function TestGCOPs(colTeeth As Collection)

'    Dim objTeeth As actXTeethControl.clsTeeth
'    Dim I As Integer
'    Dim objPicture As IPictureDisp
'
'    Load frmTestOps
'    frmTestOps.Show
'
'    I = 0
'    For Each objTeeth In colTeeth
'      If objTeeth.ToothComplimentID <= 32 Then
'        If Not objTeeth.ipicScaled Is Nothing Then
'            frmTestOps.Picture1(objTeeth.ToothComplimentID - 1).Picture = objTeeth.ipicScaled
'        End If
'      End If
'    Next
End Function

Private Sub Class_Initialize()
  
    bIsGCOPs = False
    
#If DebugMode = 1 Then
'        Dim oLogon As InternalLogon
        Dim bConnected As Boolean
        
        Set oLogon = New InternalLogon
        bConnected = oLogon.bGetConnected("mmm", "mmm")
        
        Set oLogon = Nothing
        
        If (Not bConnected) Then
            MsgBox "Logon Failed!"
        End If
#End If
    
End Sub

Private Sub Class_Terminate()
    If Not objBigTooth Is Nothing And bIsGCOPs Then
        objBigTooth.removeGCOPs
        Set objBigTooth = Nothing
        bIsGCOPs = False
    End If
    Set oLogon = Nothing
  
End Sub
