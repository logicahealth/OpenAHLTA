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
Attribute VB_Name = "clsQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public mObjManage As clsManage
Public mbPatientChanged As Boolean
Public Cancel As Boolean

Public Event ToothLoaded(iToothID As Integer, strToothSVG As String)
Public Function loadTeeth(Positions As colPositions, ViewMode As eViewMode)

Dim oPos As clsPosition
    
'    Dim aAgePath() As String
'    Dim varTooth As Variant
    Dim iTooth As Integer
    
    Dim oSvg, oSSvg, oRect As Object
    Dim s As String
    Dim colTooth As Collection
    'Dim mobjdataelement As clsDataElem
    Dim oDx As CHCSII_Dental_DXTX.clsDx
    Dim oTx As CHCSII_Dental_DXTX.clsTx
    
'    Debug.Assert False
    
    For Each oPos In Positions
    
        If Cancel = True Then
            Exit Function
        End If
        
        iTooth = oPos.ToothNum
        
        'If the patient has changed then we want to reInitialize our drawing collection
        If ViewMode = eDxView And mbPatientChanged Then
        
            'Reset the dirty flag so the dxs will draw
            For Each oDx In mObjManage.mObjSVGDraw.mObjDataProc.TeethDxData.Item(iTooth)
                oDx.GCIsDirty = True
            Next
            
            'Reset the dirty flag so the dxs will draw
            For Each oTx In mObjManage.mObjSVGDraw.mObjDataProc.TeethTxData.Item(iTooth)
                oTx.GCIsDirty = True
            Next
            
            mObjManage.DrawSVG (iTooth)
            
        End If
        
        If ViewMode = eDxView Then
            s = mObjManage.GetSVGDxViewByTooth(iTooth, False)
        ElseIf ViewMode = eTxView Then
            s = mObjManage.GetSVGTxViewByTooth(iTooth, False)
        ElseIf ViewMode = eDx603aView Then
            s = mObjManage.getsvgDx603aViewbyTooth(iTooth, False)
        End If
   
        If Cancel = True Then
            Exit Function
        End If
   
        If s <> "" Then
            RaiseEvent ToothLoaded(iTooth, s)
        End If
       
    Next
        
'    aAgePath = Split(sAgePath, ",")
'
'    For Each varTooth In aAgePath
'
'        If Cancel = True Then
'            Exit Function
'        End If
'
'        nTooth = CInt(varTooth)
'
'        'If the patient has changed then we want to reInitialize our drawing collection
'        If bIsDx And mbPatientChanged Then
'
'            'Reset the dirty flag so the dxs will draw
'            For Each oDx In mObjManage.mObjSVGDraw.mObjDataProc.TeethDxData.Item(nTooth)
'                oDx.GCIsDirty = True
'            Next
'
'            'Reset the dirty flag so the dxs will draw
'            For Each oTx In mObjManage.mObjSVGDraw.mObjDataProc.TeethTxData.Item(nTooth)
'                oTx.GCIsDirty = True
'            Next
'
'            mObjManage.DrawSVG (nTooth)
'
'        End If
'
'        If bIsDx Then
'            s = mObjManage.GetSVGDxViewByTooth(nTooth, False)
'        Else
'            s = mObjManage.GetSVGTxViewByTooth(nTooth, False)
'        End If
'
'        If Cancel = True Then
'            Exit Function
'        End If
'
'        If s <> "" Then
'            RaiseEvent ToothLoaded(CInt(varTooth), s)
'        End If
'
'    Next
           
    mbPatientChanged = False 'Reset to false no matter what since we are finished
            
End Function

Private Sub Class_Terminate()
  Set mObjManage = Nothing
End Sub
