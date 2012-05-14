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
Attribute VB_Name = "AppointmentTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements TemplateData

Private Const StorageNCID As String = "1021"
Private Const MyLabel As String = "Associated Appointment Types"

Private mListBox As VB.ListBox
Private mItems As Dictionary ' items are appointment ien plus description
Private mbDirty As Boolean

Private oEMCategoryDictionary As Scripting.Dictionary
Private oSpecialtyDictionary As Scripting.Dictionary


Public Sub AddItem(ByRef rsIEN As String, ByRef rsDesc As String)
    
    On Error GoTo ErrHandler
    mItems.Add rsIEN, rsDesc
    mbDirty = True
    mDisplayItems
    Exit Sub

ErrHandler:
    If Err = 457 Then
       mItems.Remove rsIEN
       Resume
    End If
    Err.Raise Err.Number, "CHCSII_TemplateClient.AppointmentTypes.AddItem", Err.Description
    
End Sub

Private Sub TemplateData_DoMoveDown()
    
    Dim sNCIDFrom As String
    Dim sDescFrom As String
    Dim sNCIDTo As String
    Dim sDescTo As String
    Dim i As Integer
    
    On Error GoTo ErrHandler
    For i = mListBox.ListCount - 1 To 0 Step -1
        If mListBox.Selected(i) Then
            sNCIDFrom = CStr(mListBox.ItemData(i))
            sDescFrom = mListBox.List(i)
            sNCIDTo = CStr(mListBox.ItemData(i + 1))
            sDescTo = mListBox.List(i + 1)
            'SCR-53144 5/4/2004 RBELL Fix data related error
            mItems.Key(sNCIDFrom) = "-1"
            mItems.Key(sNCIDTo) = sNCIDFrom
            mItems(sNCIDFrom) = sDescFrom
            mItems.Key("-1") = sNCIDTo
            mItems(sNCIDTo) = sDescTo
            mDisplayItems
            mListBox.Selected(i + 1) = True
            mbDirty = True
        End If
    Next i
    
    Exit Sub
    
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.AppointmentTypes.DoMoveDown", Err.Description

End Sub

Private Sub TemplateData_DoMoveUp()
    
    Dim sNCIDFrom As String
    Dim sDescFrom As String
    Dim sNCIDTo As String
    Dim sDescTo As String
    Dim i As Integer
    
    On Error GoTo ErrHandler
    For i = mListBox.ListCount - 1 To 0 Step -1
        If mListBox.Selected(i) Then
            sNCIDFrom = CStr(mListBox.ItemData(i))
            sDescFrom = mListBox.List(i)
            sNCIDTo = CStr(mListBox.ItemData(i - 1))
            sDescTo = mListBox.List(i - 1)
            'SCR-53144 5/4/2004 RBELL Fix data related error
            mItems.Key(sNCIDFrom) = "-1"
            mItems.Key(sNCIDTo) = sNCIDFrom
            mItems(sNCIDFrom) = sDescFrom
            mItems.Key("-1") = sNCIDTo
            mItems(sNCIDTo) = sDescTo
            mDisplayItems
            mListBox.Selected(i - 1) = True
            mbDirty = True
            Exit For
        End If
    Next i
    
    Exit Sub
    
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.AppointmentTypes.DoMoveUp", Err.Description

End Sub


Private Function TemplateData_Export() As String
    On Error GoTo ErrHandler
    TemplateData_Export = SerializeArray(mItems.Keys, gsListDelim)
    Exit Function
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.AppointmentTypes.Export", Err.Description
End Function

Private Function TemplateData_GetFromEncounter(objEncounter As CHCSIIEncounterOps.Encounter) As Boolean
    
    Dim oConn As CHCSII_CONN.Conn
    Dim oSQL As ICHCSII_SQL.ISqlOps
    
    On Error GoTo ErrHandler
    Set oConn = New CHCSII_CONN.Conn
    Set oSQL = oConn.CHCSII_SQLOPS(Auto)
    
    'GEMS >>
    On Error GoTo ErrHandler
    oSQL.Execute "Select ien, code from appt_type t, appointment a" _
        & " where a.appt_id = " & objEncounter.AppointmentId _
        & " and a.appt_type = t.code" _
        & " and a.facility_ncid = t.facility_ncid"
    If oSQL.BOF And oSQL.EOF Then GoTo CleanUpAndExit
    Do While Not oSQL.EOF
        mItems.Add CStr(oSQL("IEN")), oSQL("CODE")
        oSQL.MoveNext
    Loop
    GoTo CleanUpAndExit
    
ErrHandler:
    If Err = 457 Then
        mItems.Remove oSQL("IEN")
        Resume
    End If
    Err.Raise Err.Number, "CHCSII_TemplateClient.AppointmentTypes.GetFromEncounter", Err.Description
CleanUpAndExit:
    Set oSQL = Nothing
    Set oConn = Nothing
    
End Function

Private Function TemplateData_Import(rsString As String) As Boolean
    
    Dim oConn As CHCSII_CONN.Conn
    Dim oSQL As ICHCSII_SQL.ISqlOps
    Dim cNCIDs As Collection
    Dim varNCID As Variant
    
    'SCR-34348 4/8/03 RBELL DISABLED IMPORT OF THIS ITEM FOR PGUI MODES
    On Error GoTo ErrHandler
    If gobjShared Is Nothing Then Set gobjShared = New CWShared
    If gobjShared.IsAppMode(modeCHCSI_GUI) Then Exit Function '<SCR 36493
    Set oConn = New CHCSII_CONN.Conn
    Set oSQL = oConn.CHCSII_SQLOPS(Auto)
            
    Set cNCIDs = DeserializeString(rsString, gsListDelim)
    For Each varNCID In cNCIDs
        oSQL.Execute "Select CODE from appt_type where ien = " & varNCID _
            & " and facility_ncid = " & cwobjLogonEnviron.FacilityNCID
        If Not (oSQL.BOF And oSQL.EOF) Then
            mItems.Add CStr(varNCID), oSQL("CODE")
        End If
    Next varNCID
    If Not mListBox Is Nothing Then Call mDisplayItems
    
    mbDirty = False
    TemplateData_Import = True
    GoTo CleanUpAndExit

ErrHandler:
    If Err = 457 Then
        mItems.Remove CStr(varNCID)
        Resume
    End If
    Err.Raise Err.Number, "CHCSII_TemplateClient.AppointmentTypes.Import", Err.Description
    
CleanUpAndExit:
    Set oConn = Nothing
    Set oSQL = Nothing

End Function

Private Function TemplateData_GetQueryCriteria() As String
    
    Dim i As Integer
    Dim sStr As String
    
    On Error GoTo ErrHandler
    For i = 0 To mItems.Count - 1
        If sStr <> "" Then
            sStr = sStr & " or"
        End If
        sStr = sStr & "( type_ncid = " & StorageNCID & " and value_ncid = " & mItems.Keys(i) & ")"
    Next i
    
    TemplateData_GetQueryCriteria = sStr
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.AppointmentTypes.GetQueryCriteria", Err.Description
    
End Function

Private Sub mDisplayItems()
    
    Dim i As Integer
    
    On Error GoTo ErrHandler
    mListBox.Clear
    For i = 0 To mItems.Count - 1
        mListBox.AddItem mItems.Items(i)
        mListBox.ItemData(mListBox.NewIndex) = mItems.Keys(i)
    Next i

    Exit Sub
    
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.AppointmentTypes.mDisplayItems", Err.Description

End Sub

Private Sub Class_Initialize()
    On Error Resume Next
    Set mItems = New Dictionary
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Set mItems = Nothing
    Set mListBox = Nothing
    Set oEMCategoryDictionary = Nothing
    Set oSpecialtyDictionary = Nothing
End Sub

Private Property Get TemplateData_CanEdit() As Boolean
    On Error GoTo ErrHandler
    TemplateData_CanEdit = False
    Exit Property
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.AppointmentTypes.CanEdit", Err.Description
End Property

Private Function TemplateData_DeSerialize(rvData As Variant) As Boolean
    
    On Error GoTo ErrHandler
    TemplateData_DeSerialize = GEMS_TemplateData_DeSerialize(rvData)
    
    GoTo CleanUpAndExit
    
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.AppointmentTypes.DeSerialize", Err.Description
    
CleanUpAndExit:

End Function
'<< GEMS        '*-- Coleman 053001: Add Theater support
Private Function GEMS_TemplateData_DeSerialize(rvData As Variant) As Boolean
    
    Dim oRef As GEMS_EncTmplt.ETXref
    Dim oConn As CHCSII_CONN.Conn
    Dim oSQL As ICHCSII_SQL.ISqlOps
    
    On Error GoTo ErrHandler
    Set oConn = New CHCSII_CONN.Conn
    Set oSQL = oConn.CHCSII_SQLOPS(Auto)
    
    For Each oRef In rvData.References
        If oRef.eType = StorageNCID Then
            oSQL.Execute "Select CODE from appt_type where ien = " & oRef.Value _
                & " and facility_ncid = " & cwobjLogonEnviron.FacilityNCID
            If Not (oSQL.BOF And oSQL.EOF) Then
               If Not mItems.Exists(CStr(oRef.Value)) Then
                    mItems.Add CStr(oRef.Value), oSQL("CODE")
               Else
                    mItems(CStr(oRef.Value)) = oSQL("CODE")
               End If
            End If
        End If
    Next oRef
    
    If Not mListBox Is Nothing Then Call mDisplayItems
    mbDirty = False
    GoTo CleanUpAndExit
    
ErrHandler:
    If Err = 457 Then
       If Not oRef Is Nothing Then
          mItems.Remove oRef.Value
          Resume
       End If
    End If
    Err.Raise Err.Number, "CHCSII_TemplateClient.AppointmentTypes.GEMS_TemplateData_DeSerialize", Err.Description
CleanUpAndExit:
    Set oRef = Nothing
    Set oSQL = Nothing
    Set oConn = Nothing

End Function
'GEMS >>

Private Sub TemplateData_DoAdd()
    
    'SCR-50502 RBELL 2/27/2004 Added check to see if search form was loaded if not then send core messages
    Dim bIsFormLoaded As Boolean
    Dim oForm As Form
    
    On Error GoTo ErrHandler
    bIsFormLoaded = False
    For Each oForm In Forms
        If UCase(oForm.Name) = "FRMSEARCH" Then
          bIsFormLoaded = True
          Exit For
        End If
    Next
    If Not bIsFormLoaded Then cwobjCWComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_TEMPLATE_MANAGEMENT
    frmApptType.DoAdd Me
    If Not bIsFormLoaded Then cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_TEMPLATE_MANAGEMENT
    
    Exit Sub
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.AppointmentTypes.DoAdd", Err.Description
    
End Sub

Private Sub TemplateData_DoEdit()

End Sub

Private Sub TemplateData_DoRemove()
    
    Dim i As Integer
    
    On Error GoTo ErrHandler
    For i = mListBox.ListCount - 1 To 0 Step -1
        If mListBox.Selected(i) Then
            mItems.Remove CStr(mListBox.ItemData(i))
            mListBox.RemoveItem i
            mbDirty = True
        End If
    Next i
    
    Exit Sub
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.AppointmentTypes.DoRemove", Err.Description

End Sub

Private Property Let TemplateData_isDirty(RHS As Boolean)
    On Error GoTo ErrHandler
    mbDirty = RHS
    Exit Property
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.AppointmentTypes.IsDirty", Err.Description
End Property

Private Property Get TemplateData_isDirty() As Boolean
    On Error GoTo ErrHandler
    TemplateData_isDirty = mbDirty
    Exit Property
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.AppointmentTypes.IsDirty", Err.Description
End Property

Private Property Get TemplateData_Items() As Scripting.IDictionary
    On Error GoTo ErrHandler
    Set TemplateData_Items = mItems
    Exit Property
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.AppointmentTypes.Items", Err.Description
End Property

Private Property Get TemplateData_Label() As String
    On Error GoTo ErrHandler
    TemplateData_Label = MyLabel
    Exit Property
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.AppointmentTypes.Label", Err.Description
End Property

Private Property Set TemplateData_ListBox(RHS As Object)
    On Error GoTo ErrHandler
    Set mListBox = RHS
    Call mDisplayItems
    Exit Property
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.AppointmentTypes.Listbox", Err.Description
End Property

Private Property Get TemplateData_ListBox() As Object

End Property

' Changes to this module should probably be replicated in GEMS version.
Private Function TemplateData_Serialize(rvData As Variant) As Boolean
    
    On Error GoTo ErrHandler
    
    TemplateData_Serialize = GEMS_TemplateData_Serialize(rvData)
    
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.AppointmentTypes.Serialize", Err.Description

End Function
'<< GEMS    '*-- Coleman 060601: Adding Theater support
' Changes to this module should probably be replicated in CDR version.
Private Function GEMS_TemplateData_Serialize(rvData As Variant) As Boolean
    
    Dim i As Integer
    Dim objRef As GEMS_EncTmplt.ETXref
    
    On Error GoTo ErrHandler
    For i = 0 To mItems.Count - 1
        Set objRef = rvData.References.Add
        objRef.eType = StorageNCID
        objRef.Value = mItems.Keys(i)
    Next i
    
    Set objRef = Nothing
    GEMS_TemplateData_Serialize = True
    mbDirty = False
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.AppointmentTypes.Serialize", Err.Description

End Function


Public Property Get EMCategoryDictionary() As Scripting.Dictionary
    On Error GoTo ErrHandler
    Set EMCategoryDictionary = oEMCategoryDictionary
    Exit Property
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.AppointmentTypes.Get_EMCategory", Err.Description
End Property

Public Property Set EMCategoryDictionary(ByVal RHV As Scripting.Dictionary)
    On Error GoTo ErrHandler
    Set oEMCategoryDictionary = RHV
    Exit Property
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.AppointmentTypes.Set_EMCategory", Err.Description
End Property

Public Property Get SpecialtyDictionary() As Scripting.Dictionary
    On Error GoTo ErrHandler
    Set SpecialtyDictionary = oSpecialtyDictionary
    Exit Property
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.AppointmentTypes.Get_Specialty", Err.Description
End Property

Public Property Set SpecialtyDictionary(ByVal RHV As Scripting.Dictionary)
    On Error GoTo ErrHandler
    Set oSpecialtyDictionary = RHV
    Exit Property
ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.AppointmentTypes.Set_Specialty", Err.Description
End Property
