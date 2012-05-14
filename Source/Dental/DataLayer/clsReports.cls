VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarTxs As colTxs 'local copy
Private mvarDxs As colDxs 'local copy
Private mvarPatientID As Long 'local copy

'Collection to hold all Txs to ensure "unique" Tx data integrity
Private mvarTxsRef As New Collection

'Collection to hold all Dxs to ensure "unique" Dx data integrity
Private mvarDxsRef As New Collection

Private Property Let PatientID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PatientID = 5
    mvarPatientID = vData
End Property


Public Property Get PatientID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PatientID
    PatientID = mvarPatientID
End Property



Public Function LoadTxs(PatientID As Long) As Boolean

    'Open the DB
    OpenDB

'Load all the Txs from the PatientID

    'Set the Dxs = nothing
    Set mvarDxs = Nothing
    
    Set mvarTxs = New colTxs
    
    mvarDxs.DxsRef = mvarDxsRef
    mvarDxs.TxsRef = mvarTxsRef
    
    If mvarTxs.RefreshTxsFromPatientID(PatientID, sReports, 0) Then
        mvarPatientID = PatientID
        LoadTxs = True
    End If

End Function

Public Function LoadDxs(PatientID As Long) As Boolean

    'Open the DB
    OpenDB

    'Load all the Dxs from the PatientID
    
    'Unload the Txs
    Set mvarTxs = Nothing
    
    Set mvarDxs = New colDxs
    
    mvarDxs.DxsRef = mvarDxsRef
    mvarDxs.TxsRef = mvarTxsRef
    
    If mvarDxs.RefreshDxsFromPatientID(PatientID, SourceEnum.sReports, -1) Then
        mvarPatientID = PatientID
        LoadDxs = True
    End If

End Function

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



Private Sub Class_Initialize()
    'OpenDB
End Sub


