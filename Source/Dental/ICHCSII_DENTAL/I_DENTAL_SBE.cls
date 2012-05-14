VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IDENTAL_SBE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub ClearSBE()

End Sub
Public Property Get Comments() As String

End Property

Public Property Get RTFComments() As String

End Property

Public Property Get SBE() As Integer

End Property
Public Property Get SBERequiredFlag() As Boolean

End Property

Public Property Get PatientUnitId() As Long

End Property

Public Property Get DentalEncounterId() As Long

End Property

Public Sub GetUserInput(DentalEncounterId As String, _
                PatientUnitId As String, Optional blnSBEForceDisp As Boolean = True)

End Sub

Public Sub SaveView()

End Sub

