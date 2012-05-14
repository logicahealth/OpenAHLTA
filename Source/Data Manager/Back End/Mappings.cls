VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Mappings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"Mapping"
Attribute VB_Ext_KEY = "Member0" ,"Mapping"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable to hold collection
Private mCol As Collection

Public Function Add(ApplicationNCID As Long, Name As String, Before As String, After As String) As Mapping
    'create a new object
    Dim objNewMember As Mapping
    
    On Error GoTo errXIT
    
    If Me.Exists(ApplicationNCID, Name, Before) Then
        Err.Raise vbObjectError + 100, App.Title & ".Mappings", "Duplicate 'ApplicationNCID, Name, Before' is not allowed."
    End If
    If Me.Exists(ApplicationNCID, Name, , After) Then
        Err.Raise vbObjectError + 100, App.Title & ".Mappings", "Duplicate 'ApplicationNCID, Name, After' is not allowed."
    End If
    
    Set objNewMember = New Mapping

    'set the properties passed into the method
    objNewMember.ApplicationNCID = ApplicationNCID
    objNewMember.Name = Name
    objNewMember.Before = Before
    objNewMember.After = After
    mCol.Add objNewMember


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing
    
    Exit Function

errXIT:
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

Public Property Get Item(vntIndexKey As Variant) As Mapping
Attribute Item.VB_UserMemId = 0
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

Public Sub Clear()
    Set mCol = New Collection
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
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

Public Function Find(ByVal ApplicationNCID As Long, ByVal Name As String, _
                     Optional ByVal Before As String, Optional ByVal After As String) As String
    
    Dim objMapping As Mapping
    
    On Error GoTo errXIT:
    
    If Len(After) + Len(Before) = 0 Then
        Err.Raise vbObjectError, App.Title & ".Mappings", "Ambiguous request, only set Before OR After."
    End If
    
    If Len(After) * Len(Before) <> 0 Then
        Err.Raise vbObjectError, App.Title & ".Mappings", "Ambiguous request, must set Before OR After."
    End If
    
    For Each objMapping In mCol
    
        If objMapping.ApplicationNCID = ApplicationNCID And objMapping.Name = Name Then
        
            If Not Len(After) Then
                If objMapping.After = After Then
                    Find = objMapping.Before
                    Exit For
                End If
            End If
            
            If Not Len(Before) Then
                If objMapping.Before = Before Then
                    Find = objMapping.After
                    Exit For
                End If
            End If
            
        End If
    Next
    
    Exit Function
    
errXIT:
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

Public Function Exists(ByVal ApplicationNCID As Long, ByVal Name As String, _
                        Optional ByVal Before As String, Optional ByVal After As String) As Boolean

    Dim objMapping As Mapping
    
    On Error GoTo errXIT:
    
    If Len(After) + Len(Before) = 0 Then
        Err.Raise vbObjectError, App.Title & ".Mappings", "Ambiguous request, only set Before OR After."
    End If
    
    If Len(After) * Len(Before) <> 0 Then
        Err.Raise vbObjectError, App.Title & ".Mappings", "Ambiguous request, must set Before OR After."
    End If
    
    Exists = False
    
    For Each objMapping In mCol
    
        If objMapping.ApplicationNCID = ApplicationNCID And objMapping.Name = Name Then
        
            If Len(After) Then
                If objMapping.After = After Then
                    Exists = True
                    Exit For
                End If
            End If
            
            If Len(Before) Then
                If objMapping.Before = Before Then
                    Exists = True
                    Exit For
                End If
            End If
            
        End If
    Next
    
    Exit Function

errXIT:
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

