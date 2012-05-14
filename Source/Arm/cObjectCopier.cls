VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cObjectCopier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function ObjectToStream( _
                  sInstanceTag As String, _
                  oInObjToCpy As Object _
                  ) As Byte()
  Dim oPB As PropertyBag
  On Error GoTo EH
  
  Set oPB = New PropertyBag
  oPB.WriteProperty sInstanceTag, oInObjToCpy
  ObjectToStream = oPB.Contents
  
  Exit Function
EH:
  Call ARMDebugAssert(Err.Number = 0)
  Exit Function
End Function
Public Function StreamToObject( _
                  sInstanceTag As String, _
                  byArr() As Byte _
                  ) As Object
  Dim oPB As PropertyBag
  On Error GoTo EH
  
  Set oPB = New PropertyBag
  oPB.Contents = byArr
  
  Set StreamToObject = oPB.ReadProperty(sInstanceTag)
  
  Exit Function
EH:
  Set StreamToObject = Nothing
  Call ARMDebugAssert(Err.Number = 0)
  Exit Function
End Function

