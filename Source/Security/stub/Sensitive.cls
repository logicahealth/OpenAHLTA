VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Sensitive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ICHCSII_Sensitive

Private m_NCID As String
Private m_colSensitivities As New Collection

Friend Function SetNCID(NCID As String)
    m_NCID = NCID
End Function

Private Sub ICHCSII_Sensitive_AddSensitivity(SensitivityLevel As String)
    m_colSensitivities.Add SensitivityLevel
End Sub

Private Function ICHCSII_Sensitive_GetNCID() As String
    ICHCSII_Sensitive_GetNCID = m_NCID
End Function

Private Function ICHCSII_Sensitive_GetSensitivities() As String()
    Dim sSensitivities() As String
    Dim i As Integer
    
    If (m_colSensitivities.Count > 0) Then
        ReDim sSensitivities(1 To m_colSensitivities.Count)
        For i = 1 To m_colSensitivities.Count
            sSensitivities(i) = m_colSensitivities.Item(i)
        Next
    End If
    
    ICHCSII_Sensitive_GetSensitivities = sSensitivities
End Function

Private Function ICHCSII_Sensitive_IsSensitive() As Boolean
    If (m_colSensitivities.Count > 0) Then
        ICHCSII_Sensitive_IsSensitive = True
    Else
        ICHCSII_Sensitive_IsSensitive = False
    End If
End Function
