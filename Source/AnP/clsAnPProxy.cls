VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAnPProxy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Class Name:    clsAnPProxy.cls
'Author:        Phillip Crowder
'Created:       12/10/01
'Description:   This class is used to provide an interface to an AnP object.  This
'               class is used to avoid a circular reference in frmAnP.
'Revisions:

Option Explicit

Public Event Save()

'AnP object call its save method.
Public Sub Save()
  RaiseEvent Save
End Sub
