VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSVGLibSF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get SVGLib() As SVG_Lib
    
    If mvarSVGLib Is Nothing Then
        Set mvarSVGLib = New SVG_Lib
        mvarSVGLib.loadSVGToothDefinition
    End If
    
    Set SVGLib = mvarSVGLib
    
End Property


Private Sub Class_Terminate()

    'Set mvarSVGLib = Nothing

End Sub

