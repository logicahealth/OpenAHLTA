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
Attribute VB_Name = "PriorityCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'local variable to hold collection
Private mcolPriority As Collection

Public Property Get Count() As Long
     Count = mcolPriority.Count
End Property

Public Property Get Item(IndexKey As hlLABPRIORITY) As PriorityType
     Dim szIndexKey As String
     
     On Error GoTo ErrorHandler:
     'Make sure item requested is in the collection
     If IndexKey > 0 And IndexKey <= mcolPriority.Count Then
          'Since collection not exposed, assume client will use key to identify item
          szIndexKey = CStr(IndexKey)
          Set Item = mcolPriority.Item(szIndexKey)
     ElseIf mcolPriority.Count > 0 Then  'If indexkey not a collection item, then default to the first item
          Set Item = mcolPriority.Item(1)
     End If     'if collection has no items, return Nothing
     
     Exit Property
     
ErrorHandler:
     'Return item(1) as default
     If mcolPriority.Count > 0 Then
          Set Item = mcolPriority.Item(1)
     End If    'if collection has no items, return Nothing
     
End Property

'Enable For...Each support
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
     Set NewEnum = mcolPriority.[_NewEnum]
End Property


Private Sub Class_Initialize()
     Dim oPriority As PriorityType

     Set oPriority = New PriorityType
     
     Set mcolPriority = New Collection
     'Load up the collection
          oPriority.TypeCode = hlLAB_PRIORITY_ROUTINE
          oPriority.Text = "Routine"
          oPriority.HL7Code = "R"
          mcolPriority.Add oPriority, CStr(oPriority.TypeCode)
          
          Set oPriority = New PriorityType
          oPriority.TypeCode = hlLAB_PRIORITY_STAT
          oPriority.Text = "STAT"
          oPriority.HL7Code = "S"
          mcolPriority.Add oPriority, CStr(oPriority.TypeCode)
          
          Set oPriority = New PriorityType
          oPriority.TypeCode = hlLAB_PRIORITY_ASAP
          oPriority.Text = "ASAP"
          oPriority.HL7Code = "A"
          mcolPriority.Add oPriority, CStr(oPriority.TypeCode)
          
          Set oPriority = New PriorityType
          oPriority.TypeCode = hlLAB_PRIORITY_NOTIFY
          oPriority.Text = "Notify"
          oPriority.HL7Code = "N" 'changed 10-25-01 to match CHCS I change:CHCS_SS_APOMA_4603
          mcolPriority.Add oPriority, CStr(oPriority.TypeCode)
          
          Set oPriority = New PriorityType
          oPriority.TypeCode = hlLAB_PRIORITY_PREOP
          oPriority.Text = "Preop"
          oPriority.HL7Code = "P"
          mcolPriority.Add oPriority, CStr(oPriority.TypeCode)
          
          
          
'Public Enum hlLABPRIORITY
'    hlLAB_PRIORITY_ROUTINE = 1
'    hlLAB_PRIORITY_STAT = 2
'    hlLAB_PRIORITY_ASAP = 3
'    hlLAB_PRIORITY_NOTIFY = 4
'    hlLAB_PRIORITY_PREOP = 5
'End Enum

'Select Case szPriority
'      Case "S": GetLabPriorityfromText = hlLAB_PRIORITY_STAT
'      Case "A": GetLabPriorityfromText = hlLAB_PRIORITY_ASAP
'      Case "P": GetLabPriorityfromText = hlLAB_PRIORITY_PREOP
'      Case "R": GetLabPriorityfromText = hlLAB_PRIORITY_ROUTINE
'      Case "T": GetLabPriorityfromText = hlLAB_PRIORITY_NOTIFY
'  End Select
          
Debug.Print "PriorityCollection Initialize"
End Sub

Public Function SearchDescription(szSearchString As String) As hlLABPRIORITY
     Dim oPriority As PriorityType
          Set oPriority = New PriorityType
     
     For Each oPriority In mcolPriority
          If InStr(1, oPriority.Text, szSearchString, vbTextCompare) > 0 Then
               SearchDescription = oPriority
               Exit For
          End If
     Next
     'If searchstring not found, then 0 returned
End Function

Public Function SearchHL7Code(szSearchString As String) As hlLABPRIORITY
     Dim oPriority As PriorityType
          Set oPriority = New PriorityType
     
     For Each oPriority In mcolPriority
          If InStr(1, oPriority.HL7Code, szSearchString, vbTextCompare) > 0 Then
                SearchHL7Code = oPriority
               Exit For
          End If
     Next
     'If searchstring not found, then 0 returned
End Function


Private Sub Class_Terminate()
     Set mcolPriority = Nothing
     Debug.Print "PriorityCollection Terminate"
End Sub


