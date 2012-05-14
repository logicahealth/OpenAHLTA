VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCharts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarTeeth As colTeeth 'local copy
Private mvarPatientID As Long
Private mvarTxs As colTxs
Private mvarDxs As colDxs
Private mvarDentalEncounterID As Long

'Collection to hold all Txs to ensure "unique" Tx data integrity
Private mvarTxsRef As New Collection

'Collection to hold all Dxs to ensure "unique" Dx data integrity
Private mvarDxsRef As New Collection

Private mvarGroups As colGroups

Public Property Set Groups(ByVal vData As colGroups)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Groups = Form1
    Set mvarGroups = vData
End Property


Public Property Get Groups() As colGroups
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Groups
    Set Groups = mvarGroups
End Property



Public Function SaveGroups() As Boolean

MsgBox "Warning, this is not used anymore: clsCharts.SaveGroups"
Exit Function

    'If the Groups is not nothing, then save
    If Not mvarGroups Is Nothing Then
        SaveGroups = mvarGroups.SaveGroups
    Else
        SaveGroups = True
    End If

End Function

Public Function LoadGroups(PatientID As Long, CurrentDentalEncounterID As Long, Optional ToothID As Long = 0) As Boolean
   
MsgBox "Warning, this is not used anymore: clsCharts.LoadGroups"
Exit Function
   
    'Open the DB
    OpenDB

    'Load all the Dxs from the PatientID

    'Unload the Txs, Dxs and Refs
    Set mvarDxsRef = Nothing
    Set mvarTxsRef = Nothing

    Set mvarGroups = New colGroups
    
    mvarGroups.DxsRef = mvarDxsRef
    mvarGroups.TxsRef = mvarTxsRef
    
    If mvarGroups.RefreshGroupsFromPatientID(PatientID, SourceEnum.sCharts, CurrentDentalEncounterID, ToothID) Then
        mvarPatientID = PatientID
        mvarDentalEncounterID = CurrentDentalEncounterID
        LoadGroups = True
    End If

    LoadGroups = True
 
End Function


Public Property Get DentalEncounterID() As Long

    DentalEncounterID = mvarDentalEncounterID
    
End Property


Public Function LoadDxs(PatientID As Long, CurrentDentalEncounterID As Long, Optional ToothID As Long = 0) As Boolean

    'Open the DB
    OpenDB

    'Load all the Dxs from the PatientID

    'Unload the Txs
    Set mvarTxs = Nothing
    Set mvarDxsRef = Nothing
    Set mvarTxsRef = Nothing
        
    Set mvarDxs = New colDxs
    
    mvarDxs.DxsRef = mvarDxsRef
    mvarDxs.TxsRef = mvarTxsRef
    
    If mvarDxs.RefreshDxsFromPatientID(PatientID, SourceEnum.sCharts, CurrentDentalEncounterID, ToothID) Then
        mvarPatientID = PatientID
        mvarDentalEncounterID = CurrentDentalEncounterID
        LoadDxs = True
    End If

    LoadDxs = True
    
End Function



Public Property Let PatientID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PatientID = 5
    mvarPatientID = vData
End Property


Public Property Get PatientID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PatientID
    PatientID = mvarPatientID
End Property











