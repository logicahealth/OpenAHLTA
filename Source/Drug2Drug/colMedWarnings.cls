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
Attribute VB_Name = "colMedWarnings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsField"
Attribute VB_Ext_KEY = "Member0" ,"clsField"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable to hold collection
Private mCol As Collection

Public Function Add(WarningType As String, _
                    WarningText As String, _
                    DrugClass As String, _
                    DrugList As String, _
                    Optional MonographID As String = "", _
                    Optional MonographType As Integer, _
                    Optional SeverityLevel As String, _
                    Optional sKey As String) As MedWarning

    
    'create a new object
    Dim objNewMember As MedWarning
    
    On Error GoTo Err_Handler
    
    '** Do not add duplicate Warning to the Warning Collections
    If IsWarningExist(WarningText) = True Then
        Exit Function
    End If
    
    Set objNewMember = New MedWarning

    If sKey = "" Then
        sKey = MonographID & "-" & WarningType
    End If

    'set the properties passed into the method
    With objNewMember
        .WarningType = WarningType
        .WarningText = WarningText
        .DrugClass = DrugClass
        .DrugList = DrugList
        .MonographID = MonographID
        .MonographType = MonographType
        .SeverityLevel = SeverityLevel
    End With
    
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        On Error Resume Next
        mCol.Add objNewMember, sKey
        If Err.Number > 0 Then
            '*** Tag: FFFE,E00D may occurs multiple times
            '*** This tag is not important, it is used to group
            '*** the tag inside the value field for SQ
            mCol.Add objNewMember, ""
            Err.Clear
        End If
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing
Exit Function
Err_Handler:
    Err.Raise vbObjectError + 1001, "colMedWarnings_Add", "Error adding MedWarning object to MedWarning collection."
    
End Function



Public Function Find(MonographID As String) As MedWarning

    Dim objMedwarning As MedWarning
    
    On Error GoTo Err_Handler
        
     Set Find = Nothing

    For Each objMedwarning In mCol
        With objMedwarning
            If UCase$(.MonographID) = UCase$(MonographID) Then
                Set Find = New MedWarning
                Set Find = objMedwarning
                Exit For
            End If
        End With
    Next objMedwarning
    
    Exit Function

Err_Handler:
    Set Find = Nothing
    
End Function


Public Function IsWarningExist(WarningText As String) As Boolean

    Dim objMedwarning As MedWarning

    IsWarningExist = False
        
    If mCol Is Nothing Then
        Exit Function
    End If
    
    For Each objMedwarning In mCol
        With objMedwarning
            If UCase$(.WarningText) = UCase$(WarningText) Then
                IsWarningExist = True
                Exit For
            End If
        End With
    Next objMedwarning
    
    Exit Function
    
End Function

Public Property Get Item(vntIndexKey As Variant) As MedWarning
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub



