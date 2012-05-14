VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Public gobjShared As CWShared
Private mobjParent As HealthHistory
Private mobjSumOLEModules As clsSumOLEModules

Public Property Set SumOLEModules(RHV As clsSumOLEModules): Set mobjSumOLEModules = RHV: End Property
Public Property Set Parent(RHV As HealthHistory):           Set mobjParent = RHV:        End Property

Public Sub Clear()
    sumlNumOleModules = 0
    ReDim sumatOleModules(0)
    sumbReadAndStore = False
    sumsFormTitle = ""
    sumsSummaryOshKey = ""
    sumsSupportOshKey = ""
    sumsOwnerNCID = ""
End Sub

Public Function Add(ByVal vsObjectName As String, _
                   ByVal vbOnSummary As Boolean, _
                   ByVal vlSumLeftPercent As Long, _
                   ByVal vlSumTopPercent As Long, _
                   ByVal vlSumWidthPercent As Long, _
                   ByVal vlSumHeightPercent As Long, _
                   ByVal vsApplicationNCID As String, _
                   ByVal vsApplicationName As String) As Boolean
    
    On Error GoTo errhandler
            
    sumlNumOleModules = sumlNumOleModules + 1
    If (sumlNumOleModules = 1) Then
        ReDim sumatOleModules(1 To 1)
    Else
        ReDim Preserve sumatOleModules(1 To sumlNumOleModules)
    End If
    sumatOleModules(sumlNumOleModules).sObjectName = vsObjectName
    sumatOleModules(sumlNumOleModules).bOnSummary = vbOnSummary
    sumatOleModules(sumlNumOleModules).lSumLeftPercent = vlSumLeftPercent
    sumatOleModules(sumlNumOleModules).lSumTopPercent = vlSumTopPercent
    sumatOleModules(sumlNumOleModules).lSumWidthPercent = vlSumWidthPercent
    sumatOleModules(sumlNumOleModules).lSumHeightPercent = vlSumHeightPercent
    sumatOleModules(sumlNumOleModules).sApplicationNCID = vsApplicationNCID
    sumatOleModules(sumlNumOleModules).sApplicationName = vsApplicationName
    Add = True
    Exit Function
errhandler:
    Add = False
    Exit Function
End Function


Public Function DisplaySummaryConfig() As Long
        
    On Error GoTo errhandler
    
    If mobjSumOLEModules Is Nothing Then Exit Function
    
    Screen.MousePointer = vbHourglass
    
    sumsSummaryOshKey = "support patient summary view"
    sumsOwnerNCID = mobjParent.LogonEnviron.UserNCID
    sSummaryTitle = "Health History"
    
    
    
    Set frmProperties.SumOLEModules = mobjSumOLEModules
    frmProperties.Display
        
    'Fill in the return values
    DisplaySummaryConfig = sumlReturnValue
    
    If (sumlReturnValue = vbOK) Then
        '---Save to the database if we are in ReadAndStore Mode---
        Call StoreOleInfo(mobjSumOLEModules)
    End If
    
    GoTo Cleanup

errhandler:
    DisplaySummaryConfig = vbCancel
Cleanup:
    Screen.MousePointer = vbDefault
    Exit Function
End Function



Private Sub Class_Initialize()
    sumlNumOleModules = 0
    ReDim sumatOleModules(0)
    sumbReadAndStore = False
    sumsFormTitle = ""
    sumsSummaryOshKey = ""
    sumsSupportOshKey = ""
    sumsOwnerNCID = ""
    
    On Error Resume Next
    
End Sub

Public Sub StoreOleInfo(SumOLEModules As clsSumOLEModules)
    Dim i As Integer
    Dim sText As String
    Dim oModule As clsSumOLEModule
    '- for the special gems-meds
    Dim sModuleName As String
    
    '-------------------------------------------------------------
    '  Store out new summary information about each OLE Server.
    '-------------------------------------------------------------
    For Each oModule In SumOLEModules
        With oModule
            '---format a "summary text" string to store---
            If .bOnSummary Then
                sText = "YES;" & _
                        CStr(.lSumLeftPercent) & ";" & _
                        CStr(.lSumTopPercent) & ";" & _
                        CStr(.lSumWidthPercent) & ";" & _
                        CStr(.lSumHeightPercent) & ";"
            Else
                sText = "NO;0;0;0;0;"
            End If
        End With
        
'<< fix the 'save' problem for Gems_Meds
        If oModule.sObjectName = "Gems_MMMHISMeds.Meds" Then
            sModuleName = "MMMHISMeds.Meds"
        Else
            sModuleName = oModule.sObjectName
        End If
           
        If Not mobjParent.Config.setValue(NCID_CW_APPLICATION, sumsOwnerNCID, sumsSummaryOshKey, sModuleName, sText) Then
            MsgBox "Internal Error: Unable to 'Set' summary configuration value, the changes will not be saved.", vbExclamation, "Summary Config"
            Exit Sub
        End If
    Next oModule
    
    If (Not mobjParent.Config.SaveChanges(NCID_CW_APPLICATION, sumsOwnerNCID)) Then
        MsgBox "Internal Error: An error occurred in saving the Summary changes.", vbExclamation, "Summary Config"
    End If
End Sub


