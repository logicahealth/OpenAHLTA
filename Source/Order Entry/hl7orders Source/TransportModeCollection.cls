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
Attribute VB_Name = "TransportModeCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'local variable to hold collection
Private mcolTransportMode As Collection

Public Property Get Count() As Long
     Count = mcolTransportMode.Count
End Property

Public Property Get Item(IndexKey As hlRADXPORT) As TransportModeType
     Dim szIndexKey As String
     
     On Error GoTo ErrorHandler:
     'Make sure item requested is in the collection
     If IndexKey > 0 And IndexKey <= mcolTransportMode.Count Then
          'Since collection not exposed, assume client will use key to identify item
          szIndexKey = CStr(IndexKey)
          Set Item = mcolTransportMode.Item(szIndexKey)
     ElseIf mcolTransportMode.Count > 0 Then  'If indexkey not a collection item, then default to the first item
          Set Item = mcolTransportMode.Item(1)
     End If     'if collection has no items, return Nothing
     
     Exit Property
     
ErrorHandler:
     'Return item(1) as default
     If mcolTransportMode.Count > 0 Then
          Set Item = mcolTransportMode.Item(1)
     End If    'if collection has no items, return Nothing
     
End Property

'Enable For...Each support
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
     Set NewEnum = mcolTransportMode.[_NewEnum]
End Property


Private Sub Class_Initialize()
     Dim oTransportMode As TransportModeType

     Set oTransportMode = New TransportModeType
     
     Set mcolTransportMode = New Collection
     'Load up the collection
          oTransportMode.TypeCode = hlRAD_XPORT_AMBULATE
          oTransportMode.Text = "Ambulatory"
          oTransportMode.HL7Code = "WALK"
          mcolTransportMode.Add oTransportMode, CStr(oTransportMode.TypeCode)
          
          Set oTransportMode = New TransportModeType
          oTransportMode.TypeCode = hlRAD_XPORT_WHEELCHAIR
          oTransportMode.Text = "Wheelchair"
          oTransportMode.HL7Code = "WHLC"
          mcolTransportMode.Add oTransportMode, CStr(oTransportMode.TypeCode)
          
          Set oTransportMode = New TransportModeType
          oTransportMode.TypeCode = hlRAD_XPORT_STRETCHER
          oTransportMode.Text = "Stretcher, Crib or Bassinet"
          oTransportMode.HL7Code = "CART"
          mcolTransportMode.Add oTransportMode, CStr(oTransportMode.TypeCode)
          
          Set oTransportMode = New TransportModeType
          oTransportMode.TypeCode = hlRAD_XPORT_PORTABLE
          oTransportMode.Text = "Portable"
          oTransportMode.HL7Code = "PORT"
          mcolTransportMode.Add oTransportMode, CStr(oTransportMode.TypeCode)
          
          
'Public Enum hlRADXPORT
'    hlRAD_XPORT_AMBULATE = 1
'    hlRAD_XPORT_WHEELCHAIR = 2
'    hlRAD_XPORT_STRETCHER = 3
'    hlRAD_XPORT_PORTABLE = 4
'End Enum

'Function GetTextfromRadXport(lngEnumVar As hlRADXPORT) As String
'
'    Select Case lngEnumVar
'    Case hlRAD_XPORT_AMBULATE: GetTextfromRadXport = "WALK"   'AMBULATE
'    Case hlRAD_XPORT_WHEELCHAIR: GetTextfromRadXport = "WHLC"  'WHEELCHAIR
'    Case hlRAD_XPORT_STRETCHER:   GetTextfromRadXport = "CART"  'STRETCHER
'    Case hlRAD_XPORT_PORTABLE: GetTextfromRadXport = "PORT"  'PORTABLE
'    End Select
'
'End Function

'Function szTransportMode(hlTransportMode As hlRADXPORT) As String
'     Select Case hlTransportMode
'     Case hlRADXPORT.hlRAD_XPORT_AMBULATE: szTransportMode = "Ambulatory"
'     Case hlRADXPORT.hlRAD_XPORT_PORTABLE: szTransportMode = "Portable"
'     Case hlRADXPORT.hlRAD_XPORT_STRETCHER: szTransportMode = "Stretcher, Crib or Bassinet"
'     Case hlRADXPORT.hlRAD_XPORT_WHEELCHAIR: szTransportMode = "Wheelchair"
'     End Select
'End Function
          
Debug.Print "TransportModeCollection Initialize"
End Sub

Public Function SearchDescription(szSearchString As String) As hlRADXPORT
Dim oTransportMode As TransportModeType
     Set oTransportMode = New TransportModeType

For Each oTransportMode In mcolTransportMode
     If InStr(1, oTransportMode.Text, szSearchString, vbTextCompare) > 0 Then
           SearchDescription = oTransportMode
          Exit For
     End If
Next
'If searchstring not found, return 0
End Function

Public Function SearchHL7Code(szSearchString As String) As hlRADXPORT
Dim oTransportMode As TransportModeType
     Set oTransportMode = New TransportModeType

For Each oTransportMode In mcolTransportMode
     If InStr(1, oTransportMode.HL7Code, szSearchString, vbTextCompare) > 0 Then
          SearchHL7Code = oTransportMode
          Exit For
     End If
Next
'If searchstring not found, return 0
End Function


Private Sub Class_Terminate()
     Set mcolTransportMode = Nothing
     Debug.Print "TransportModeCollection Terminate"
End Sub




