VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EncounterProcedure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private msDataId As String
Private msSnoID As String

Property Get DataID() As String
   DataID = msDataId
End Property

Property Get SnoID() As String
   SnoID = msSnoID
End Property

Public Sub Init(ByVal sDataId As String, ByVal sSnoID As String)
   msDataId = sDataId
   msSnoID = sSnoID
End Sub
