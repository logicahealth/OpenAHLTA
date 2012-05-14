VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TemplateHx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mobjTemplateFavorites As Collection
Private mcolTemplateHxs As Collection
Private mobjCurrentTemplateHxItem As TemplateHxItem
Private mbCurrentEncounterTemplateSet As Boolean

Public Event FavoriteListUpdated(TemplateFavs As Collection)
Public Event TemplateSelected(TemplHxItem As TemplateHxItem)
Public Property Get IsCurrentEncounterTemplateValid() As Boolean
    IsCurrentEncounterTemplateValid = mbCurrentEncounterTemplateSet
End Property
Public Property Let IsCurrentEncounterTemplateValid(val As Boolean)
   mbCurrentEncounterTemplateSet = val
End Property

Public Property Get TemplateFavoriteList() As Collection
    Set TemplateFavoriteList = mobjTemplateFavorites
End Property
Public Property Get CurrentTemplateHxItem() As TemplateHxItem
    Set CurrentTemplateHxItem = mobjCurrentTemplateHxItem
End Property
Public Property Get TemplateHxs() As Collection
    Set TemplateHxs = mcolTemplateHxs
End Property
Public Function GetDefaultTemplate() As String
    
    Dim sBuf As String
    
    On Error GoTo ErrHandler
    
    sBuf = ""
    
    If Not cwobjConfig Is Nothing Then
        cwobjConfig.getValue NCID_TEMPLATE_MANAGEMENT, "Default", "DefaultTemplate", sBuf
    Else
        sBuf = ""
    End If

    GetDefaultTemplate = sBuf
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "TemplateHx.GetDefaultTemplate", Err.Description, Err.HelpFile, Err.HelpContext
End Function
Public Function Current_History_Item(TemplHxItem As TemplateHxItem, Optional bRaiseEvent As Boolean = True) As Boolean
    
    Set mobjCurrentTemplateHxItem = TemplHxItem
    
    If bRaiseEvent = True Then
        RaiseEvent TemplateSelected(TemplHxItem)
    End If
    
    Current_History_Item = True
End Function
Public Function Add_History_Item(TemplHxItem As TemplateHxItem) As Boolean
Dim i As Integer
Dim oTemplateHxItem As TemplateHxItem

    If mobjCurrentTemplateHxItem Is Nothing Then
        If TemplHxItem.Key = "0" Then
            TemplHxItem.Key = 1
            If checkForDupsInBaseCollection(TemplHxItem.Key) = False Then
                mcolTemplateHxs.Add TemplHxItem, TemplHxItem.Key
            End If
        End If
        
    ElseIf TemplHxItem.Key = "-1" Then
        If checkForDupsInBaseCollection(TemplHxItem.Key) = False Then
            mcolTemplateHxs.Add TemplHxItem, TemplHxItem.Key, , 1
        End If
    ElseIf TemplHxItem.Key = "-2" And mbCurrentEncounterTemplateSet = True Then
        If checkForDupsInBaseCollection(TemplHxItem.Key) = False Then
            mcolTemplateHxs.Add TemplHxItem, TemplHxItem.Key, , 2
        End If
    ElseIf TemplHxItem.Key = "-2" And mbCurrentEncounterTemplateSet = False Then
        If checkForDupsInBaseCollection(TemplHxItem.Key) = False Then
            mcolTemplateHxs.Add TemplHxItem, TemplHxItem.Key, , 1
        End If
    Else
        For i = mcolTemplateHxs.Count To 2 Step -1

            If Not mcolTemplateHxs(i).Key = "-1" And Not mcolTemplateHxs(i).Key = "-2" Then
        'emc: may be issue here with not removing enough items.  In fact it appears to leave the item
        'one beyond current, although we should probably be destroying it and replacing with new.
                If i > mobjCurrentTemplateHxItem.Key + 1 Then
                
                     'emc: error occurs here when item i does not exist.  exm: i = 7, but key of
                     'last element is "6"
                     'Determined cause to be when re-entering module after a current encounter
                     'is added after the initial entry.  Then in code above, the current
                     'encounter history is inserted in the collection, making the key and count
                     'out of synch.
                    'Remove_History_Item CStr(i)
                    RemoveHistoryItemByIndex i
                End If
            End If
            
        Next i
        
        TemplHxItem.Key = mcolTemplateHxs.Count + 1
        mcolTemplateHxs.Add TemplHxItem, TemplHxItem.Key
    End If
    
    Set mobjCurrentTemplateHxItem = TemplHxItem
    Add_History_Item = True
    
End Function
Private Function checkForDupsInBaseCollection(sKey As String) As Boolean
Dim bfound As Boolean
Dim oTemplateHxItem As TemplateHxItem

bfound = False

    For Each oTemplateHxItem In mcolTemplateHxs
        If oTemplateHxItem.Key = sKey Then
            bfound = True
            Exit For
        End If
    Next
    
    checkForDupsInBaseCollection = bfound
    
End Function

'Public Function Remove_History_Item(sKey As String) As Boolean
Private Function Remove_History_Item(sKey As String) As Boolean
    mcolTemplateHxs.Remove sKey
    Remove_History_Item = True
    
End Function
Private Function RemoveHistoryItemByIndex(intIndex As Integer) As Boolean
   On Error GoTo ErrHandler
   mcolTemplateHxs.Remove intIndex
   RemoveHistoryItemByIndex = True
   Exit Function
ErrHandler:
   RemoveHistoryItemByIndex = False
End Function

Private Function checkDefaultInFavList(TemplateFavs As Collection, sDefaultTemplate As String) As Boolean
Dim oTemplateLite As TemplateLite
Dim bfound As Boolean

For Each oTemplateLite In TemplateFavs

    If oTemplateLite.TemplateID = sDefaultTemplate Then
        bfound = True
        Exit For
    End If

Next

checkDefaultInFavList = bfound

End Function
Public Function Add_Favorites(TemplateFavs As Collection) As Boolean
Dim oTemplateLite As TemplateLite
Dim sDefaultTemplate As String

    If Not TemplateFavs Is Nothing Then
        
        sDefaultTemplate = GetDefaultTemplate

        Set mobjTemplateFavorites = New Collection
        
        If Len(sDefaultTemplate) <> 0 And TemplateFavs.Count > 0 Then
            If checkDefaultInFavList(TemplateFavs, sDefaultTemplate) = True Then
               mobjTemplateFavorites.Add TemplateFavs(sDefaultTemplate), TemplateFavs(sDefaultTemplate).TemplateID
            End If
        End If
        
        'SortandGroup TemplateFavs
        For Each oTemplateLite In TemplateFavs
            If Not oTemplateLite.TemplateID = sDefaultTemplate Then
                SortNameAndGroup oTemplateLite, TemplateFavs, sDefaultTemplate
            End If
        Next
        'Set top of tree for history purposes
        Set oTemplateLite = New TemplateLite
        oTemplateLite.TemplateID = 0
        If mobjTemplateFavorites.Count = 0 Then
            mobjTemplateFavorites.Add oTemplateLite, "0"
        Else
            mobjTemplateFavorites.Add oTemplateLite, "0", 1
        End If
        
        'Set top of tree for history purposes with current encounter
        If mbCurrentEncounterTemplateSet = True Then
                Set oTemplateLite = New TemplateLite
                oTemplateLite.TemplateID = -1
                If mobjTemplateFavorites.Count = 0 Then
                    mobjTemplateFavorites.Add oTemplateLite, "-1"
                Else
                    mobjTemplateFavorites.Add oTemplateLite, "-1", , 1
                End If
        End If
        
'        Set mobjTemplateFavorites = TemplateFavs
        RaiseEvent FavoriteListUpdated(mobjTemplateFavorites)
        Add_Favorites = True
    Else
        Add_Favorites = False
    End If
    
    
End Function
Private Function SortNameAndGroup(robjTemplateLite As TemplateLite, TemplateFavs As Collection, sDefaultTemplate As String) As Boolean
Dim objTemplateLite As TemplateLite
Dim bfound As Boolean

bfound = False

    For Each objTemplateLite In mobjTemplateFavorites
        If objTemplateLite.TemplateType = robjTemplateLite.TemplateType Then
            If StrComp(UCase(objTemplateLite.TemplateName), UCase(robjTemplateLite.TemplateName)) = 1 And Not objTemplateLite.TemplateID = sDefaultTemplate Then
                If mobjTemplateFavorites.Count = 0 Then
                    mobjTemplateFavorites.Add robjTemplateLite, robjTemplateLite.TemplateID
                Else
                    mobjTemplateFavorites.Add robjTemplateLite, robjTemplateLite.TemplateID, objTemplateLite.TemplateID
                End If
                bfound = True
                Exit For
            End If
        End If
    Next
    
    If bfound = False Then
        mobjTemplateFavorites.Add robjTemplateLite, robjTemplateLite.TemplateID
    End If

End Function
Public Function Clear_Favorites() As Boolean

    RaiseEvent FavoriteListUpdated(Nothing)
    Set mobjTemplateFavorites = Nothing
    Clear_Favorites = True
    
End Function

Private Sub Class_Initialize()

'    Set mcolTemplateHxs = New Collection
    
End Sub
Public Sub LoadEncounterTemplateHX()
    getEncounterTemplateHx
End Sub
Private Sub getEncounterTemplateHx()
Dim objEncounterOps As CHCSIIEncounterOps.EncounterOps
Dim objEncParent As CHCSIIEncounterCurrent.EncounterParent

    
    Set objEncParent = gobjComm.InitializeOLEServer(NCID_ENCOUNTER_APPLICATION)
    Set objEncounterOps = objEncParent.EncounterOps
    Set mcolTemplateHxs = objEncounterOps.TemplateHx
    If mcolTemplateHxs.Count > 1 Then
        If VarType(mcolTemplateHxs(mcolTemplateHxs.Count)) = vbString Then
            Set mobjCurrentTemplateHxItem = mcolTemplateHxs(mcolTemplateHxs(mcolTemplateHxs.Count))
            mcolTemplateHxs.Remove (mcolTemplateHxs.Count)
        End If
    Else
      'Make sure it's not set  (can be left to something after closing the encounter)
      Set mobjCurrentTemplateHxItem = Nothing
    End If
End Sub
Private Sub Class_Terminate()
    If Not mobjCurrentTemplateHxItem Is Nothing Then
        mcolTemplateHxs.Add CStr(mobjCurrentTemplateHxItem.Key)
    End If
    Set mcolTemplateHxs = Nothing
    Set mobjCurrentTemplateHxItem = Nothing
End Sub
