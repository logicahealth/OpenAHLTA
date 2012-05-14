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
Attribute VB_Name = "iEncLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Sub Init(ByRef Conn As CHCSII_CONN.Conn, ByRef Cache As DL_Support.CacheMapper, _
                ByVal DataAccessMode As CHCSII_CONN.eBackend, Optional ByVal CacheConfigItems As Boolean = False)
        
End Sub

'Locks an encounter section.
Public Function LockSection(ByRef Params As iParameters) As ISqlOps
End Function

'Unlocks an encounter section.
Public Function UnlockSection(ByRef Params As iParameters) As Boolean
End Function

