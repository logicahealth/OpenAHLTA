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
Attribute VB_Name = "DMBE_IDMappers"
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

Implements IDMBE.I_IDMappers '<SCR 24303 CC>

Public Function I_IDMappers_Add(ApplicationNCID As Long, Name As String, Before As String, After As String) As IDMBE.I_IDMapper
    
    'create a new object
    Dim objNewMember As IDMBE.I_IDMapper
    
    On Error GoTo errXIT
    

'<< Begin:SCR #36759;   Developer: Brian Mowbray 06/03/2003 10:24 AM
'    No known reason to raise an error if dups exist in this collection, just don't add it

'    If I_IDMappers_Exists(ApplicationNCID, Name, Before) Then
'        Err.Raise vbObjectError + 100, App.Title & ".Mappings", "Duplicate 'ApplicationNCID, Name, Before' is not allowed."
'    End If
'    If I_IDMappers_Exists(ApplicationNCID, Name, , After) Then
'        Err.Raise vbObjectError + 100, App.Title & ".Mappings", "Duplicate 'ApplicationNCID, Name, After' is not allowed."
'    End If
    
    If I_IDMappers_Exists(ApplicationNCID, Name, Before) Or I_IDMappers_Exists(ApplicationNCID, Name, , After) Then
      Exit Function
    End If
'>> End: SCR #36759;
    
    
    Set objNewMember = New DMBE.DMBE_IDMapper

    'set the properties passed into the method
    objNewMember.ApplicationNCID = ApplicationNCID
    objNewMember.Name = Name
    objNewMember.Before = Before
    objNewMember.After = After
    mCol.Add objNewMember


    'return the object created
    Set I_IDMappers_Add = objNewMember
    Set objNewMember = Nothing
    
    Exit Function

errXIT:
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

Public Property Get I_IDMappers_Item(vntIndexKey As Variant) As IDMBE.I_IDMapper
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set I_IDMappers_Item = mCol(vntIndexKey)
End Property



Public Property Get I_IDMappers_Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    I_IDMappers_Count = mCol.Count
End Property

Public Sub I_IDMappers_Clear()
    Set mCol = New Collection
End Sub


Public Property Get I_IDMappers_NewEnum() As IUnknown
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set I_IDMappers_NewEnum = mCol.[_NewEnum]
End Property


Public Function I_IDMappers_Find(ByVal ApplicationNCID As Long, ByVal Name As String, _
                     Optional ByVal Before As String, Optional ByVal After As String) As String
    
    Dim objMapping As IDMBE.I_IDMapper
    
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
                    I_IDMappers_Find = objMapping.Before
                    Exit For
                End If
            End If
            
            If Not Len(Before) Then
                If objMapping.Before = Before Then
                    I_IDMappers_Find = objMapping.After
                    Exit For
                End If
            End If
            
        End If
    Next
    
    Exit Function
    
errXIT:
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

Public Function I_IDMappers_Exists(ByVal ApplicationNCID As Long, ByVal Name As String, _
                        Optional ByVal Before As String, Optional ByVal After As String) As Boolean

    Dim objMapping As IDMBE.I_IDMapper
    
    On Error GoTo errXIT:
    
    If Len(After) + Len(Before) = 0 Then
        Err.Raise vbObjectError, App.Title & ".Mappings", "Ambiguous request, only set Before OR After."
    End If
    
    If Len(After) * Len(Before) <> 0 Then
        Err.Raise vbObjectError, App.Title & ".Mappings", "Ambiguous request, must set Before OR After."
    End If
    
    I_IDMappers_Exists = False
    
    For Each objMapping In mCol
    
        If objMapping.ApplicationNCID = ApplicationNCID And objMapping.Name = Name Then
        
            If Len(After) Then
                If objMapping.After = After Then
                    I_IDMappers_Exists = True
                    Exit For
                End If
            End If
            
            If Len(Before) Then
                If objMapping.Before = Before Then
                    I_IDMappers_Exists = True
                    Exit For
                End If
            End If
            
        End If
    Next
    
    Exit Function

errXIT:
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub


'***********************************************
'Delete from here
'
'
'
'
'Public Function Add(ApplicationNCID As Long, Name As String, Before As String, After As String) As DMBE_IDMapper
'
'    'create a new object
'    Dim objNewMember As DMBE_IDMapper
'
'    On Error GoTo errXIT
'
'    If Me.Exists(ApplicationNCID, Name, Before) Then
'        Err.Raise vbObjectError + 100, App.Title & ".Mappings", "Duplicate 'ApplicationNCID, Name, Before' is not allowed."
'    End If
'    If Me.Exists(ApplicationNCID, Name, , After) Then
'        Err.Raise vbObjectError + 100, App.Title & ".Mappings", "Duplicate 'ApplicationNCID, Name, After' is not allowed."
'    End If
'
'    Set objNewMember = New DMBE_IDMapper
'
'    'set the properties passed into the method
'    objNewMember.ApplicationNCID = ApplicationNCID
'    objNewMember.Name = Name
'    objNewMember.Before = Before
'    objNewMember.After = After
'    mCol.Add objNewMember
'
'
'    'return the object created
'    Set Add = objNewMember
'    Set objNewMember = Nothing
'
'    Exit Function
'
'errXIT:
'    Err.Raise Err.Number, Err.Source, Err.Description
'
'End Function
'
'Public Property Get Item(vntIndexKey As Variant) As DMBE_IDMapper
'    'used when referencing an element in the collection
'    'vntIndexKey contains either the Index or Key to the collection,
'    'this is why it is declared as a Variant
'    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
'  Set Item = mCol(vntIndexKey)
'End Property
'
'
'
'Public Property Get Count() As Long
'    'used when retrieving the number of elements in the
'    'collection. Syntax: Debug.Print x.Count
'    Count = mCol.Count
'End Property
'
'Public Sub Clear()
'    Set mCol = New Collection
'End Sub
'
'
'Public Property Get NewEnum() As IUnknown
'    'this property allows you to enumerate
'    'this collection with the For...Each syntax
'    Set NewEnum = mCol.[_NewEnum]
'End Property
'
'
'
'
'Public Function Find(ByVal ApplicationNCID As Long, ByVal Name As String, _
'                     Optional ByVal Before As String, Optional ByVal After As String) As String
'
'    Dim objMapping As DMBE_IDMapper
'
'    On Error GoTo errXIT:
'
'    If Len(After) + Len(Before) = 0 Then
'        Err.Raise vbObjectError, App.Title & ".Mappings", "Ambiguous request, only set Before OR After."
'    End If
'
'    If Len(After) * Len(Before) <> 0 Then
'        Err.Raise vbObjectError, App.Title & ".Mappings", "Ambiguous request, must set Before OR After."
'    End If
'
'    For Each objMapping In mCol
'
'        If objMapping.ApplicationNCID = ApplicationNCID And objMapping.Name = Name Then
'
'            If Not Len(After) Then
'                If objMapping.After = After Then
'                    Find = objMapping.Before
'                    Exit For
'                End If
'            End If
'
'            If Not Len(Before) Then
'                If objMapping.Before = Before Then
'                    Find = objMapping.After
'                    Exit For
'                End If
'            End If
'
'        End If
'    Next
'
'    Exit Function
'
'errXIT:
'    Err.Raise Err.Number, Err.Source, Err.Description
'
'End Function
'
'Public Function Exists(ByVal ApplicationNCID As Long, ByVal Name As String, _
'                        Optional ByVal Before As String, Optional ByVal After As String) As Boolean
'
'    Dim objMapping As DMBE_IDMapper
'
'    On Error GoTo errXIT:
'
'    If Len(After) + Len(Before) = 0 Then
'        Err.Raise vbObjectError, App.Title & ".Mappings", "Ambiguous request, only set Before OR After."
'    End If
'
'    If Len(After) * Len(Before) <> 0 Then
'        Err.Raise vbObjectError, App.Title & ".Mappings", "Ambiguous request, must set Before OR After."
'    End If
'
'    Exists = False
'
'    For Each objMapping In mCol
'
'        If objMapping.ApplicationNCID = ApplicationNCID And objMapping.Name = Name Then
'
'            If Len(After) Then
'                If objMapping.After = After Then
'                    Exists = True
'                    Exit For
'                End If
'            End If
'
'            If Len(Before) Then
'                If objMapping.Before = Before Then
'                    Exists = True
'                    Exit For
'                End If
'            End If
'
'        End If
'    Next
'
'    Exit Function
'
'errXIT:
'    Err.Raise Err.Number, Err.Source, Err.Description
'
'End Function
'
