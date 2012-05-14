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
Attribute VB_Name = "clsDataProc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Const TOTAL_TEETH = 52
Private mvarTeethDxs As colTeethDxs
Private mvarTeethTxs As colTeethTxs

Public Property Get TeethDxData() As colTeethDxs
    Set TeethDxData = mvarTeethDxs
End Property

Public Property Get TeethTxData() As colTeethTxs
    Set TeethTxData = mvarTeethTxs
End Property

Public Function LoadDataLayer(colGroups As CHCSII_Dental_DXTX.colGroups)

'This function groups all the dxs and txs by their respective teeth

Dim oGroup As CHCSII_Dental_DXTX.clsGroup
Dim oDx As CHCSII_Dental_DXTX.clsDx
Dim oTmpDx As CHCSII_Dental_DXTX.clsDx

Dim oSS As CHCSII_Dental_DXTX.clsSurfaceSet
Dim oTooth As CHCSII_Dental_DXTX.clsTooth
Dim oTx As CHCSII_Dental_DXTX.clsTx
Dim strArrTeeth() As String
Dim i As Integer
    
On Error GoTo Err

    'Dx and Tx replaced dataelem
    'Dim oDataElement As clsDataElem
    
    InitializeCollection
    
    If colGroups Is Nothing Then
        Exit Function
    End If
    
    'Loop through the Dxs and add the Dx to the correct ToothDxs collection under the TeethDxs collection
    For Each oGroup In colGroups
        For Each oDx In oGroup.Dxs
            If oDx.Chartable Then
                If oDx.SurfaceSets(1).Teeth(1).ToothName = "Max" Or oDx.SurfaceSets(1).Teeth(1).ToothName = "Mand" Then
                    'The teeth for max and man are comma delimited and stored in the surfaceset field
                    strArrTeeth = Split(oDx.SurfaceSets(1).SurfaceSet, ",")
                    For i = 0 To UBound(strArrTeeth)
                        'Set oDataElement = New clsDataElem
                        'Set oDataElement.Dx = oDx
                        oDx.GCIsDirty = True
                        'oDataElement.IsDirty = True
                        mvarTeethDxs.Item(CInt(strArrTeeth(i))).Add oDx
                    Next
                Else
                    If Not oDx.ICDCode = "K07.31" Then 'do not add to internal collection for displacement
                            'Set oDataElement = New clsDataElem
                            'Set oDataElement.Dx = oDx
                            oDx.GCIsDirty = True
                            'oDataElement.IsDirty = True
                            mvarTeethDxs.Item(oDx.SurfaceSets(1).Teeth(1).ToothID).Add oDx
                    End If
                End If
            End If
        Next
        
        'Loop through the Txs and add the Tx to the correct ToothDxs collection under the TeethDxs collection
        For Each oTx In oGroup.Txs
            If oTx.Chartable Then
               If oTx.SurfaceSets(1).Teeth(1).ToothName = "Max" Or oTx.SurfaceSets(1).Teeth(1).ToothName = "Mand" Then
                    'The teeth for max and man are comma delimited and stored in the surfaceset field
                    strArrTeeth = Split(oTx.SurfaceSets(1).SurfaceSet, ",")
                    For i = 0 To UBound(strArrTeeth)
                        'Set oDataElement = New clsDataElem
                        'Set oDataElement.Tx = oTx
                        oTx.GCIsDirty = True
                        'oDataElement.IsDirty = True
                        mvarTeethTxs.Item(CInt(strArrTeeth(i))).Add oTx
                    Next
                Else
                    'Set oDataElement = New clsDataElem
                    'Set oDataElement.Tx = oTx
                    oTx.GCIsDirty = True
                    'oDataElement.IsDirty = True
                    mvarTeethTxs.Item(oTx.SurfaceSets(1).Teeth(1).ToothID).Add oTx
                End If
            End If
        Next
    Next

Exit Function

Err:
    MsgBxARMd Err.Description
    
End Function
Private Sub InitializeCollection()

'Add placeholders for each tooth
Dim oToothDxs As colToothDxs
Dim oToothTxs As colToothTxs
Dim i As Integer

    Set mvarTeethDxs = New colTeethDxs
    Set mvarTeethTxs = New colTeethTxs

    For i = 1 To TOTAL_TEETH
        Set oToothDxs = New colToothDxs
        mvarTeethDxs.Add oToothDxs
        
        Set oToothTxs = New colToothTxs
        mvarTeethTxs.Add oToothTxs
    Next

End Sub
