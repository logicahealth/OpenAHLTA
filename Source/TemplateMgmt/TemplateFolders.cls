VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TemplateFolders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"TemplateFolder"
Attribute VB_Ext_KEY = "Member0" ,"TemplateFolder"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'local variable to hold collection
Private mCol As Collection

Public Function AddByValues(Folder_ID As String, Folder_Name As String, Folder_Parent_ID As String, USER_NCID As String, Folder_Type As String, Expanded As Boolean, Loaded As String, Optional sKey As String) As TemplateFolder
    'create a new object
    Dim objNewMember As TemplateFolder
    Set objNewMember = New TemplateFolder

    'set the properties passed into the method
    objNewMember.Folder_ID = Folder_ID
    objNewMember.Folder_Name = Folder_Name
    objNewMember.Folder_Parent_ID = Folder_Parent_ID
    objNewMember.USER_NCID = USER_NCID
    objNewMember.Folder_Type = Folder_Type
    objNewMember.Expanded = Expanded
    objNewMember.Loaded = Loaded
    mCol.Add sKey, objNewMember
    

    'return the object created
    Set AddByValues = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As TemplateFolder
Attribute Item.VB_UserMemId = 0
Attribute Item.VB_MemberFlags = "200"
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
    'SCR-50424 RBELL 2/26/2004 Added an error handler to return a nothing object if not found in collection (for dev purposes)
  On Error GoTo EH
  Set Item = mCol(vntIndexKey)
  Exit Property
EH:
  Set Item = Nothing
  Err.Clear
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

Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub


Public Function GetChildrenOf(ByVal sFolderID As String) As TemplateFolders
    Dim tf As TemplateFolder
    Dim sFolderType As String
    
    Set GetChildrenOf = New TemplateFolders
    
     'SCR-48962 RBELL 2/3/4 Rename folders to match naming conventions
     If sFolderID = Enterprise_Root_Key Then
        sFolderID = Enterprise_Folder_Value
        sFolderType = "EN"
     ElseIf sFolderID = MTF_Root_Key Then
        sFolderID = MTF_Folder_Value
        sFolderType = "MT"
     ElseIf sFolderID = Clinic_Root_Key Then
        sFolderID = Clinic_Folder_Value
        sFolderType = "CL"
     ElseIf sFolderID = "My Favorites" Then
        sFolderID = Favorite_Folder_Value
        sFolderType = "FV"
     ElseIf Not IsNumeric(Left(sFolderID, 1)) Then
        'SCR-47593 RBELL 1/13/04  Added mid statement to remove extra character appended to node key when the folder is a DMIS ID folder
        If Left(sFolderID, 1) = "D" Then sFolderID = mid(sFolderID, 2)
    End If
    
    For Each tf In mCol
        If tf.Folder_Parent_ID = sFolderID Then
            If sFolderID = "0" Then
                If tf.Folder_Type = sFolderType Then GetChildrenOf.Add tf.Folder_ID, tf
            Else
                GetChildrenOf.Add tf.Folder_ID, tf
            End If
        End If
    Next
End Function

Public Function GetParentOf(sFolderID As String) As TemplateFolder
    Dim tf As TemplateFolder
    If mCol.Exists(sFolderID) Then
        Set tf = mCol(sFolderID)
        If mCol.Exists(tf.Folder_Parent_ID) Then
            Set GetParentOf = mCol(tf.Folder_Parent_ID)
        End If
    End If
End Function

Public Sub Add(sKey As String, tf As TemplateFolder)
    On Error Resume Next
    mCol.Add tf, sKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Public Function Exists(sKey As String) As Boolean
On Error GoTo Exists_Err
Dim oTemp As Object

        If IsObject(mCol.Item(sKey)) Then
            Exists = True
            Set oTemp = Nothing
            Exit Function
        End If
'else
Exists_Err:
        Exists = False
End Function

Public Sub CollapseChildren(ByVal sFolderID As String)

End Sub
