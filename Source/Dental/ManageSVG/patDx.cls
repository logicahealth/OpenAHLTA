VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "patDx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private meDxStatus As CHCSII_Dental_DXTX.DxStatusEnum

Public bDrawThisDxViewDx As Boolean
Public bDrawThisTxViewDx As Boolean
Public iToothID As Integer
Public sSurface As String
Public eToothLocation As eLocation
Public iMenuID As Integer
Public iRootMenuID As Integer
Public dModifiedDateTime As Date
Public sSVGDX As String
Public iDRC As Integer
Public Property Let eDxStatus(DxStatus As CHCSII_Dental_DXTX.DxStatusEnum)
    meDxStatus = DxStatus
    
    If DxStatus = dxActiveNew Or DxStatus = dxActiveOngoing _
       Or DxStatus = dxActiveProvisional Then
            bDrawThisDxViewDx = True
            bDrawThisTxViewDx = True
    End If

End Property
Public Property Get eDxStatus() As CHCSII_Dental_DXTX.DxStatusEnum
    eDxStatus = meDxStatus
End Property


