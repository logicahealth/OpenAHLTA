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
Attribute VB_Name = "ModuleConfigHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Provides support functionality for ModuleConfig objects used by Core Mappers.

Option Explicit

'Inits a Module config object with the values in a settings array.
Public Function ExtractConfigInfo(ByRef Config As ModuleConfig, ByVal ConfigValues As Variant, ByVal ReturnAll As Boolean, _
                              Optional ByRef CommandSwitches As DL_Support.iParameters = Nothing) As Boolean
Dim Valid As Boolean

'Osh data positions
Const OSH_MENU_TEXT As Integer = 0
Const OSH_BUTTON_TEXT As Integer = 1
Const OSH_BUTTON_BITMAP As Integer = 2
Const OSH_PATIENT_REQIRED As Integer = 3
Const OSH_TOOL_TIP As Integer = 4
Const OSH_REG_NAME As Integer = 5
Const OSH_APP_NCID As Integer = 6
Const OSH_PATIENT_SUMMARY As Integer = 7
Const OSH_CMD_LINE As Integer = 8
Const OSH_APP_TYPE As Integer = 9
Const OSH_PRINT As Integer = 10
Const OSH_PRINT_PREVIEW As Integer = 11
Const OSH_PROPERTY_VIEW As Integer = 12
Const OSH_SELECTED As Integer = 13
Const OSH_PRIVILEGE As Integer = 14

  On Error GoTo ErrHandler
    
  If UBound(ConfigValues) < 13 Then
    Err.Raise 50000, , "Unable to retrieve the enterprise information for OLE Server: " & Config.ServerName & _
                        " in ExtractConfigInfo.  The CW configuration database may have become corrupted, " & _
                        "or is incompatible."
  End If

  'See If Command line switch is set that is specified in the OSH
  If Len(ConfigValues(OSH_CMD_LINE)) > 0 Then
    If Not (CommandSwitches Is Nothing) Then
      If CommandSwitches.Exists(ConfigValues(OSH_CMD_LINE)) Then
        Valid = (CommandSwitches.Item(UCase$(ConfigValues(OSH_CMD_LINE))).Value = UCase$(ConfigValues(OSH_CMD_LINE)))
      Else
        Valid = False
      End If
    Else
      Valid = False
    End If
  Else
    Valid = True
  End If

  If (Left$(UCase$(ConfigValues(OSH_SELECTED)), 1) = "Y" And Valid) Or (ReturnAll And Valid) Then
    With Config
      .ApplicationNCID = ConfigValues(OSH_APP_NCID)
      .ApplicationType = UCase$(ConfigValues(OSH_APP_TYPE))
      .ButtonCaption = ConfigValues(OSH_BUTTON_TEXT)
      .CommandLineSwitch = ConfigValues(OSH_CMD_LINE)
      .HasPatientSummary = (Left$(UCase$(ConfigValues(OSH_PATIENT_SUMMARY)), 1) = "Y")
      .HasPrint = (Left$(UCase$(ConfigValues(OSH_PRINT)), 1) = "Y")
      .HasPrintPreview = (Left$(UCase$(ConfigValues(OSH_PRINT_PREVIEW)), 1) = "Y")
      .HasPropertyView = (Left$(UCase$(ConfigValues(OSH_PROPERTY_VIEW)), 1) = "Y")
      .ImageName = ConfigValues(OSH_BUTTON_BITMAP)
      .MenuCaption = ConfigValues(OSH_MENU_TEXT)
      .RegisterName = ConfigValues(OSH_REG_NAME)
      .Selected = (UCase$(Left$(ConfigValues(OSH_SELECTED), 1)) = "Y")
      .ToolTip = ConfigValues(OSH_TOOL_TIP)
                                        
      'Set Patient required value
      Select Case UCase$(ConfigValues(OSH_PATIENT_REQIRED))
        Case "YES", "Y"
          .PatientRequired = PATIENT_REQUIRED_YES
        
        Case "NO", "N"
          .PatientRequired = PATIENT_REQUIRED_NO
        
        Case "OPTIONAL", "O"
          .PatientRequired = PATIENT_REQUIRED_OPTIONAL
        
        Case Else
          .PatientRequired = PATIENT_REQUIRED_NO
      End Select
      
      If UBound(ConfigValues) >= 14 Then
        .AccessPriv = ConfigValues(OSH_PRIVILEGE)
      End If
    End With
  Else
    Valid = False
  End If

  ExtractConfigInfo = Valid
  
  Exit Function

ErrHandler:
  Err.Raise Err.Number, "MappperCHCSII.ExtractConfigInfo", Err.Description

End Function


