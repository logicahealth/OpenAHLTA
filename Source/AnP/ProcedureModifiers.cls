VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProcedureModifiers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'BR (SCR 28157) Default Unit of Service to 1  1/17/03  Phil Crowder
Private Const UNIT_OF_SERVICE_DEFAULT As Long = 1
Private Const MAX_MODIFIERS As Long = 3

Private mUnitsOfService As Long
Private mModifiers(1, MAX_MODIFIERS - 1) As String      '0,n = Code and 1,n = Description

Public Event Modified()

Public Property Get Codes() As Collection
Dim CPTCodes As Collection
Dim i As Long

  Set CPTCodes = New Collection
  
  For i = 0 To MAX_MODIFIERS - 1
    If Len(mModifiers(0, i)) > 0 Then
      CPTCodes.Add mModifiers(0, i)
    End If
  Next
  
  Set Codes = CPTCodes
  
End Property

Public Property Get UnitsOfService() As Long
  If mUnitsOfService = 0 Then
    mUnitsOfService = UNIT_OF_SERVICE_DEFAULT
  End If

  UnitsOfService = mUnitsOfService
End Property

'BR (SCR 28158) Allow Edit Unit of Service  1/17/03 Phil Crowder
'
'BR (SCR 28157) Unit of Service must be between 1 and 999  1/17/0  Phil Crowder
Public Property Let UnitsOfService(ByVal NewValue As Long)
  If NewValue >= 1 And NewValue <= 999 Then
    If mUnitsOfService <> 0 Then RaiseEvent Modified
    
    mUnitsOfService = NewValue
  End If
End Property

'BR (28147) Add Modifers (Duplicates are ignored, ie. not allowed)  1/17/03  Phil Crowder
Public Function Add(ByVal Code As String, ByVal Description As String, Optional ByVal Loading As Boolean = False) As Boolean
Dim i As Long
Dim DupFound As Boolean
Dim BlankLocation As Long

  BlankLocation = -1

  For i = 0 To MAX_MODIFIERS - 1
    'Check for Dup.
    If mModifiers(0, i) = Code Then
      DupFound = True
      Exit For
    End If
    
    'Find Blank location in array to add new value
    If Len(mModifiers(0, i)) = 0 And BlankLocation = -1 Then
      BlankLocation = i
    End If
  Next
  
  If BlankLocation >= 0 And Not DupFound Then
    mModifiers(0, BlankLocation) = Code
    mModifiers(1, BlankLocation) = Description
    
    Add = True
    
    If Not Loading Then RaiseEvent Modified
  Else
    Add = False
  End If
  
End Function

Public Sub DeleteAll()
Dim i As Long

  For i = 0 To MAX_MODIFIERS - 1
    If Len(mModifiers(0, i)) > 0 Then RaiseEvent Modified
    
    mModifiers(0, i) = vbNullString
    mModifiers(1, i) = vbNullString
  Next

End Sub

'Array of Codes, Short Decription, Help
Public Function List(ByVal Code As String) As Variant
Dim CDR As ProxyCDRPlus

  If gShared.IsAppMode(modeCDR) Or gShared.IsAppMode(modeCHCSI_GUI) Then
    Set CDR = gBB.CDRPlusInterface
    List = CDR.ListOfModifiers(Code)
  End If
  
End Function

'JT BEGIN SCRs 38959 & 38960 - Separated RTF display from grid display since it is now different
'from what is in the grid.  Only show codes when option is selected and dynamically
'place those codes based upon whether or not modifiers exist and unit of service value
Public Function DisplayRTF(ByVal sCPTCode As String, ByVal bIncludeCodes As Boolean) As String

Dim i As Long
Dim sUOS As String

   'Process Units of Service
   If mUnitsOfService > 1 Then
      'SCR 38959 - Only include units of service if greater than 1
      sUOS = " x " & mUnitsOfService
   End If
      
   'Process Modifier List
   For i = 0 To MAX_MODIFIERS - 1
      If Len(mModifiers(0, i)) <> 0 Then
         If bIncludeCodes Then
            If Len(DisplayRTF) = 0 Then DisplayRTF = sCPTCode & "-"
            DisplayRTF = DisplayRTF & mModifiers(0, i) & " " & mModifiers(1, i) & ", "
         Else
            DisplayRTF = DisplayRTF & mModifiers(1, i) & ", "
         End If
      End If
   Next
   
   'Build the final RTF
   If Len(DisplayRTF) > 0 Then
      DisplayRTF = Left$(DisplayRTF, Len(DisplayRTF) - 2)
      DisplayRTF = sUOS & " (" & DisplayRTF & ")"
   ElseIf bIncludeCodes Then
      DisplayRTF = sUOS & " (" & sCPTCode & ")"
   Else
      DisplayRTF = sUOS
   End If
  
End Function
'JT End SCRs 38959 & 38960

'BR (29573) Modifer and Comment Format  1/17/03 Phil Crowder
'
'BR (28519) Display Unit of Service  1/17/03 Phil Crowder
'
'BR (28150) Display Modifers in Encounter Doc.  1/17/03  Phil Crowder
Public Function Display() As String
Dim i As Long

   For i = 0 To MAX_MODIFIERS - 1
      If Len(mModifiers(0, i)) <> 0 Then
         Display = Display & mModifiers(1, i) & "-" & mModifiers(0, i) & ", "
      End If
   Next
   
   If Len(Display) > 0 Then
      Display = Left$(Display, Len(Display) - 2)
      Display = " x " & mUnitsOfService & " (" & Display & ")"
  ElseIf mUnitsOfService > 1 Then
    Display = " x " & mUnitsOfService
  End If
  
End Function
