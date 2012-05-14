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
Attribute VB_Name = "clsGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarDxs As colDxs
Private WithEvents mvarTxs As colTxs
Attribute mvarTxs.VB_VarHelpID = -1
Private mvarGroupID As Long

'Collection to hold all Txs to ensure unique data integrity
Private mvarTxsRef As Collection

'Collection to hold all Dxs to ensure unique data integrity
Private mvarDxsRef As Collection

Public Function GetNote() As String

'This function gets the note from the dxs and txs

Dim oDx As clsDx
Dim oTx As clsTx
Dim sTemp As String
Dim sNote As String

    For Each oDx In mvarDxs
        sNote = oDx.GetNote

        If sNote <> "" Then
            sTemp = sTemp & sNote
        End If
    Next
    
    For Each oTx In mvarTxs
        sNote = oTx.GetNote

        If sNote <> "" Then
            sTemp = sTemp & sNote
        End If
    Next

    GetNote = sTemp
    
End Function


Public Property Let GroupID(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.GroupID = 5
    mvarGroupID = vData
End Property


Public Property Get GroupID() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.GroupID
    GroupID = mvarGroupID
End Property



Friend Function SaveGroup() As Boolean

'Loop through mvarGroup and Save the Dxs and Txs

Dim oDx As clsDx
Dim oTx As clsTx
    
    mvarDxs.SaveDxs mvarGroupID
    mvarTxs.SaveTxs mvarGroupID
    
'    For Each oDx In mvarDxs
'        oDx.SaveDx mvarGroupID
'    Next
'
'    For Each oTx In mvarTxs
'        oTx.SaveTx mvarGroupID
'    Next

    SaveGroup = True
    
End Function

Public Property Set Txs(ByVal vData As colTxs)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Txs = Form1
    Set mvarTxs = vData
End Property


Public Property Get Txs() As colTxs
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Txs
    Set Txs = mvarTxs
End Property



Friend Property Let TxsRef(vData As Collection)

    Set mvarTxsRef = vData
    mvarTxs.TxsRef = mvarTxsRef
    
End Property

Friend Property Let DxsRef(vData As Collection)

    Set mvarDxsRef = vData
    mvarDxs.DxsRef = mvarDxsRef
    
End Property

Friend Property Get TxsRef() As Collection

    Set TxsRef = mvarTxsRef

End Property

Friend Property Get DxsRef() As Collection

    Set DxsRef = mvarDxsRef

End Property

Public Property Set Dxs(ByVal vData As colDxs)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Dxs = Form1
    Set mvarDxs = vData
End Property


Public Property Get Dxs() As colDxs
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Dxs
    Set Dxs = mvarDxs
End Property



Private Sub Class_Initialize()

    Set mvarDxs = New colDxs
    Set mvarTxs = New colTxs
    
End Sub


Private Sub Class_Terminate()

    Set mvarDxs = Nothing
    Set mvarTxs = Nothing
    Set mvarTxsRef = Nothing
    Set mvarDxsRef = Nothing

End Sub


Private Sub mvarTxs_TxAdded(ByVal NewProviderID As Long, TookOwnership As Boolean)

'Tracker 17486
'The user added a Tx to the Txs collection so Prompt the user to take ownership of
'the unsigned dxs in the group

Dim oDx As clsDx
Dim bPrompted As Boolean
    
    TookOwnership = True
    
    For Each oDx In mvarDxs
        If oDx.Signed = False And oDx.ProviderID <> NewProviderID Then
            'The dx is unsigned and has a different providerid
            If bPrompted = False Then
                'Prompt the user to take ownership
                'The selected diagnosis/procedure originated from another user in an encounter that has not yet been signed.  Do you confirm this diagnosis/procedure?
                If MsgBxARMd("The added procedure is in the same group with one or more unsigned diagnosis. Do you confirm all unsigned diagnosis in the group?", vbInformation + vbYesNo, "Confirm ownership") = vbYes Then
                    oDx.ProviderID = NewProviderID
                Else
                    'The provider does not want to take ownership
                    TookOwnership = False
                    Exit For
                End If
                bPrompted = True
            Else
                oDx.ProviderID = NewProviderID
            End If
        End If
    Next

End Sub


