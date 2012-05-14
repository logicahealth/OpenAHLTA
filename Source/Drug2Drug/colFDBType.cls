VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colFDBType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsFDBType"
Attribute VB_Ext_KEY = "Member0" ,"clsFDBType"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable to hold collection
Private mCol As Collection

Private Function Add(AllergyType As AllergyTypeEnum, FDBTable As String, PK As String, DescriptionColumn As String, Optional sKey As String) As clsFDBType
    
'create a new object
Dim objNewMember As clsFDBType
Dim oRs As Recordset
Dim sSql As String

    Set objNewMember = New clsFDBType

    'set the properties passed into the method
    objNewMember.AllergyType = AllergyType
    objNewMember.FDBTable = FDBTable
    objNewMember.PK = PK
    objNewMember.DescriptionColumn = DescriptionColumn
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If
    
    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Private Sub Init()
    
'This sub inits the collection of FDB types.

    Add AllergenGroup, "FDB.FDB_ALLERGENGROUP", "AGID", "DESCRIPTION1"
    Add DispensableDrug, "FDB.FDB_DISPENSABLE", "MEDID", "DESCDISPLAY"
    Add Ingredient, "FDB.FDB_INGREDIENT", "HICSEQNO", "DESCRIPTION1"
    'No rows
    'Add Compound, "FDB.FDB_COMPOUND", "COMPOUNDID", "DESCRIPTION1"
    Add DrugName, "FDB.FDB_DRUGNAME", "MNID", "DESCDISPLAY"
    'No rows
    'Add ManufacturedDrug, "FDB.FDB_MANUFDRUG", "IMDID", "DESCRIPTION1"
    Add PackagedDrug, "FDB.FDB_PACKAGEDDRUG", "PMID", "DESCDISPLAY"
    Add RoutedDrug, "FDB.FDB_ROUTEDDRUG", "RMID", "DESCDISPLAY"

End Sub

Private Property Get Item(vntIndexKey As Variant) As clsFDBType
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Private Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property


Private Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Private Property Get NewEnum() As IUnknown
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Public Sub Search(AllergyIEN As String, oAllergyResults As colAllergyResults)
    
    'This searches the chcs_allergies tables for FDB matches
    

    Dim oFDBTypes As colFDBType
    Dim oFDBType As clsFDBType
    Dim sSql As String
    Dim sGNN As String
    Dim rsFDB As Recordset
    Dim rsAllergies As ADODB.Recordset

    On Error GoTo ErrHnd

    If oAllergyResults Is Nothing Then
        Set oAllergyResults = New colAllergyResults
    End If
    
    Set oFDBTypes = New colFDBType
    
    Init

    'Get the allergy
    sSql = "Select GNN, BN_GNN from CHCS_ALLERGIES where IEN = " & AllergyIEN
    Set rsAllergies = goDAS.ExecuteSQL(sSql)
    
    'Get the description to search
    If IsNull(rsAllergies("GNN")) And Not IsNull(rsAllergies("BN_GNN")) Then
        sGNN = rsAllergies("BN_GNN")
    ElseIf IsNull(rsAllergies("GNN")) And IsNull(rsAllergies("BN_GNN")) Then
        sGNN = ""
    Else
        sGNN = rsAllergies("GNN")
    End If
    
    If sGNN <> "" Then
        For Each oFDBType In mCol
            With oFDBType
               'Do an exact match
                sSql = "Select " & .PK & ", " & .DescriptionColumn & " from " & .FDBTable & " where upper(" & .DescriptionColumn & ") = " & goConn.SQLQuote(UCase(sGNN))
                Call Debug_Start(sSql)
                Set rsFDB = goDAS.ExecuteSQL(sSql)
                Call Debug_Stop
                Do Until rsFDB.EOF
                    oAllergyResults.Add .AllergyType, .DescriptionColumn, rsFDB.Fields(1).Value, rsFDB(.PK)
                    rsFDB.MoveNext
                Loop
                
                'Only do a like match if we have not found an exact match
                If oAllergyResults.Count = 0 Then
                    'Do a like match
                    sSql = "Select " & .PK & ", " & .DescriptionColumn & " from " & .FDBTable & " where upper(" & .DescriptionColumn & ") like " & goConn.SQLQuote("%" & UCase(sGNN) & "%")
                    Call Debug_Start(sSql)
                    Set rsFDB = goDAS.ExecuteSQL(sSql)
                    Call Debug_Stop
                    Do Until rsFDB.EOF
                        oAllergyResults.Add .AllergyType, .DescriptionColumn, rsFDB.Fields(1).Value, rsFDB(.PK)
                        rsFDB.MoveNext
                    Loop
                End If
            End With
        Next
    End If
    
    'Clean up
    If rsFDB.State = adStateOpen Then
        rsFDB.Close
    End If
    If rsAllergies.State = adStateOpen Then
        rsAllergies.Close
    End If
    Set rsAllergies = Nothing
    Set rsFDB = Nothing
    Set oFDBType = Nothing
    
Exit Sub

ErrHnd:
    MsgBox "Application error: " & Err.Description & " Occurred in Drug2Drug.colFDBType.Search"

End Sub

Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub
