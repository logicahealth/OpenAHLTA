VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "patTx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private meTxStatus As CHCSII_Dental_DXTX.DxStatusEnum

Public bDrawThisDxViewTx As Boolean
Public bDrawThisTxViewTx As Boolean
Public iToothID As Integer
Public sSurface As String
Public eToothLocation As eLocation
Public iMenuID As Integer
Public iRootMenuID As Integer
Public dModifiedDateTime As Date
Public sSVGTX As String
Public Property Let eTxStatusBothViews(bTxStatus As Boolean)
    If bTxStatus Then
        bDrawThisDxViewTx = True
        bDrawThisTxViewTx = True
    End If
End Property
Public Property Let eTxStatus(TxStatus As CHCSII_Dental_DXTX.TxStatusEnum)
    meTxStatus = TxStatus
    
    If TxStatus = txCompleted Or TxStatus = txPreExisting Then
            bDrawThisDxViewTx = True
    ElseIf TxStatus = txPlanned Or TxStatus = txInProgress Then
            bDrawThisTxViewTx = True
    End If
    
End Property
Public Property Get eTxStatus() As CHCSII_Dental_DXTX.TxStatusEnum
    eTxStatus = meTxStatus
End Property

