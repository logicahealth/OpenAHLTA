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
Attribute VB_Name = "cGUIListener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'*--
'*-- cGUIListener - This class contains a reference to the cCmplxTransactions
'*--    singleton that is instantiated by cARMWrapper. This allows the Listener
'*--    to receive Events from CmplxTransactions. The reference is set by a
'*--    call to SetCmplxTxsRef() - which is made in ARM_GUIUnitTestCheck()
'*--    if it is determined that the GUI is enabled. ARM_GUIUnitTestCheck() is
'*--    called by cARMWrapper.Class_Initialize.
'*--
'*--    This class merely passes the events along to the GUI by calling public
'*--    methods in the GUI specific module, modARMUnitTestGUI.
'*--
Private WithEvents moCmplxTxs As cCmplxTransactions
Attribute moCmplxTxs.VB_VarHelpID = -1

Private Sub Class_Terminate()
    Set moCmplxTxs = Nothing
End Sub

Public Sub SetCmplxTxsRef(oCmplxTxs As cCmplxTransactions)
    Set moCmplxTxs = oCmplxTxs
End Sub

'<<< SCR 31901 Coleman 2/20/03 - Changed signatures to post more info on GUI
'*-- CmplxTxInstanceTerminated
Private Sub moCmplxTxs_CmplxTxInstanceTerminated(sReadableName As String, _
                                lTxClsID As Long, _
                                lTxInstID As Long, _
                                eCorrFam As CorrelatedFamilies_Enum, _
                                lTimeStamp As Long)
    GUI_CurrTxRemove CStr(lTxInstID), sReadableName, CStr(lTxClsID), _
                        CStr(eCorrFam), CStr(lTimeStamp)
End Sub

'*-- CurrInstancesAdd
Private Sub moCmplxTxs_CurrInstancesAdd(sReadableName As String, _
                                lTxClsID As Long, _
                                lTxInstID As Long, _
                                eCorrelatedFamily As CorrelatedFamilies_Enum, _
                                lTimeStamp As Long)
    GUI_CurrTxAdd sReadableName, CStr(lTxClsID), CStr(lTxInstID), _
                                CStr(eCorrelatedFamily), CStr(lTimeStamp)
End Sub

'*-- FamilyCreated
Private Sub moCmplxTxs_FamilyCreated(eCorrelatedFamily As CorrelatedFamilies_Enum, _
                                        lTimeOut As Long)
    GUI_CorrFamilyAdd eCorrelatedFamily, lTimeOut
End Sub

'*-- FamilyTerminated
Private Sub moCmplxTxs_FamilyTerminated(eFam As CorrelatedFamilies_Enum, _
                                        lParentTxClsID As Long, _
                                        lParentTxInstID As Long)
    GUI_CorrFamilyTerminate eFam, lParentTxClsID, lParentTxInstID
End Sub

'*-- FamilyTxInstAdd
Private Sub moCmplxTxs_FamilyTxInstAdd(sReadableName As String, _
                                lTxClsID As Long, _
                                lTxInstID As Long, _
                                eFam As CorrelatedFamilies_Enum, _
                                lTimeStamp As Long)
    GUI_CorrFamilyAddTxToList sReadableName, CStr(lTxClsID), CStr(lTxInstID), _
                                CStr(eFam), CStr(lTimeStamp)
End Sub

'*-- FamilyTxInstOrphaned
Private Sub moCmplxTxs_FamilyTxInstOrphaned(sReadableName As String, _
                                lTxClsID As Long, _
                                lTxInstID As Long, _
                                eFam As CorrelatedFamilies_Enum, _
                                lTimeStamp As Long)
    GUI_CorrFamilyOrphanThisTx sReadableName, CStr(lTxClsID), _
                                CStr(lTxInstID), CStr(eFam), CStr(lTimeStamp)
End Sub

'*-- FamilyTxInstRemove
Private Sub moCmplxTxs_FamilyTxInstRemove(sReadableName As String, _
                                lTxClsID As Long, _
                                lTxInstID As Long, _
                                eFam As CorrelatedFamilies_Enum, _
                                lTimeStamp As Long)
    GUI_CorrFamilyDeleteTx sReadableName, CStr(lTxClsID), CStr(lTxInstID), _
                                CStr(eFam), CStr(lTimeStamp)
End Sub
'>>> SCR 31901 Coleman 2/20/03 - Changed signatures to post more info on GUI

