VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOrientation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'The DxICDCodeID for orientation ICDCode "K07.31", ICDDescription = "Displacement"
Private Const ORIENTATION_ID As Integer = 766



Public Function Show(ToothID As Integer) As Boolean

'This function shows the orientation form and saved the selected data.

Dim oSVGCommSF As srvSVGComm.SingletonFactory
Dim oSVGComm As srvSVGComm.clsSVGComm
Dim oGroup As CHCSII_Dental_DXTX.clsGroup
Dim oDx As CHCSII_Dental_DXTX.clsDx
Dim oTempDx As CHCSII_Dental_DXTX.clsDx
Dim oSS As CHCSII_Dental_DXTX.clsSurfaceSet
Dim oTooth As CHCSII_Dental_DXTX.clsTooth
Dim eStatusID As DxStatusEnum
Dim sToothName As String

    Set oSVGCommSF = New SingletonFactory
    Set oSVGComm = oSVGCommSF.Singleton

    'Get the toothname
    If ToothID > 0 And ToothID < 33 Then
        sToothName = ToothID
    ElseIf ToothID > 32 And ToothID < 53 Then
        sToothName = oSVGComm.ToothObject.GetToothName
    End If
    
    'Set up for form for the tooth
    frmOrientation.ToothID = ToothID
    frmOrientation.ToothName = sToothName
    
    'Find the orientation dx
    For Each oTempDx In oSVGComm.DataObject.AllDxs
        If oTempDx.DxICDCodeID = ORIENTATION_ID Then
            eStatusID = oTempDx.StatusID
            If eStatusID = dxActiveFinding Or eStatusID = dxActiveNew Or eStatusID = dxActiveObserve Or eStatusID = dxActiveOngoing Or eStatusID = dxActiveProvisional Then
                If oTempDx.SurfaceSets(1).Teeth(1).ToothID = ToothID Then
                    'Found the orientation dx
                    Set oSS = oTempDx.SurfaceSets(1)
                    Set oDx = oTempDx
                    Exit For
                End If
            End If
        End If
    Next
    
    If Not oDx Is Nothing Then
        'We have a tooth
        Set oTooth = oDx.SurfaceSets(1).Teeth(1)
        With frmOrientation
            .Mesial = CBool(oTooth.MesialVersion)
            .RotatedMesial = CBool(oTooth.RotatedMesial)
            .Distal = CBool(oTooth.DistalVersion)
            .RotatedDistal = CBool(oTooth.RotatedDistal)
            .Facial = CBool(oTooth.FacialVersion)
            .Lingual = CBool(oTooth.LingualVersion)
            .Intruded = CBool(oTooth.Intruded)
            .Extruded = CBool(oTooth.Extruded)
        End With
    Else
        'We don't have a tooth
        With frmOrientation
            .Mesial = False
            .RotatedMesial = False
            .Distal = False
            .RotatedDistal = False
            .Facial = False
            .Lingual = False
            .Intruded = False
            .Extruded = False
        End With
    End If
    
    cwobjCWComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_DENTAL_BIGTOOTH_APPLICATION
    frmOrientation.Show vbModal
    cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_DENTAL_BIGTOOTH_APPLICATION
    
    If frmOrientation.cmdOK.Tag = "OK" And frmOrientation.OrientationDesc <> "" Then
        'Save the data
        If oTooth Is Nothing Then
            'Add a new dx for this tooth
            Set oGroup = oSVGComm.DataObject.Groups.Add

            Set oDx = oGroup.Dxs.Add(sCharts, ORIENTATION_ID, oSVGComm.DataObject.ProviderID)
            Set oSS = oDx.SurfaceSets.Add(0)
'            oSS.SurfaceAbbr = frmOrientation.OrientationDesc
'            oSS.SurfaceSet = "777"

            Set oTooth = oSS.Teeth.Add(oSVGComm.DataObject.DentalEncounterID)
            oTooth.ToothID = ToothID
        End If
        
        With oTooth
            oSS.SurfaceAbbr = frmOrientation.OrientationDesc
            oSS.SurfaceSet = "777"
            .DistalVersion = frmOrientation.Distal
            .RotatedDistal = frmOrientation.RotatedDistal
            .RotatedMesial = frmOrientation.RotatedMesial
            .MesialVersion = frmOrientation.Mesial
            .Extruded = frmOrientation.Extruded
            .Intruded = frmOrientation.Intruded
            .FacialVersion = frmOrientation.Facial
            .LingualVersion = frmOrientation.Lingual
            .ToothName = sToothName
        End With
    ElseIf frmOrientation.cmdOK.Tag = "OK" And frmOrientation.OrientationDesc = "" Then
        If Not oDx Is Nothing Then
            'Delete the Dx
            'This commented out because it should be deleted from the grid to correctly go through the business logic
'            oDx.StatusID = dxNonActiveDeleted
        End If
    End If
    
    Unload frmOrientation
    Set oSVGCommSF = Nothing
    Set oSVGComm = Nothing
    Set oGroup = Nothing
    Set oDx = Nothing
    Set oSS = Nothing
    Set oTooth = Nothing
    
End Function


