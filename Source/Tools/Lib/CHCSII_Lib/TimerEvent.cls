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
Attribute VB_Name = "clsTimerEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''************************************************************************
'''Module:        TimerEvent.cls - clsTimerEvent  - Timer class
'''
'''Description:   Timer event information. Used by mclsLibTimer
'''
'''
'''Suggested Use:
'''
'''Note:
'''
'''Modules Required:
'''
'''Author         : Alexander Alberstein
'''Modified       :
'''               :
'''************************************************************************
Option Explicit


Public psKey            As String
Public psID             As String
Public pvTime           As Variant
Public ptTimeStart      As Date
Public psMsg            As String
Public plCallNo          As Long
Public pbActive         As Boolean
Public pvTag            As Variant
