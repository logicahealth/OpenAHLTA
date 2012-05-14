VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "COtherCritOps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"COtherCrit"
Attribute VB_Ext_KEY = "Member0" ,"COtherCrit"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Implements COtherCrit
'References CWShared
' set GUI to form holding my check boxes

Public GUI          As Form
Private mCol        As Collection
Private oCurrCrit   As COtherCrit


Public Function AddCrit(oNewMember As COtherCrit) As Boolean
    On Error GoTo errhandler
    mCol.Add oNewMember
    AddCrit = True
    
errhandler:
End Function

Public Property Get Item(vntIndexKey As Variant) As COtherCrit
Attribute Item.VB_UserMemId = 0
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

Private Property Set COtherCrit_chkControl(ByVal RHS As CheckBox):  Set oCurrCrit.chkControl = RHS:             End Property
Private Property Get COtherCrit_chkControl() As CheckBox:           Set COtherCrit_chkControl = oCurrCrit.chkControl: End Property

Private Property Let COtherCrit_FieldName(ByVal RHS As String):     oCurrCrit.FieldName = RHS:                  End Property
Private Property Get COtherCrit_FieldName() As String:              COtherCrit_FieldName = oCurrCrit.FieldName: End Property

Private Property Let COtherCrit_TypeNCID(ByVal RHS As String):      oCurrCrit.TypeNCID = RHS:                   End Property
Private Property Get COtherCrit_TypeNCID() As String:               COtherCrit_TypeNCID = oCurrCrit.TypeNCID:   End Property

Private Property Let COtherCrit_ValueText(ByVal RHS As String):     oCurrCrit.ValueText = RHS:                  End Property
Private Property Get COtherCrit_ValueText() As String:              COtherCrit_ValueText = oCurrCrit.ValueText: End Property

Public Function ReadCriteriaFile()
    Dim sIniPath  As String
    Dim f_in      As Integer
    Dim oShared   As CWShared
    Dim sLineIn() As String
    Dim sNewLine  As String
    Dim oCrit     As COtherCrit
    
    Dim lTop      As Long
    Dim lLeft     As Long
    Dim lIdx      As Long
    
    f_in = FreeFile
    Set oShared = New CWShared
    
    sIniPath = oShared.INIPath
    Open sIniPath & "\dental_triage.ini" For Input As f_in
    lLeft = GUI.chkCrit(0).Left
    Do While True
        Line Input #f_in, sNewLine
        If EOF(f_in) Then
            Exit Do
        End If
        sLineIn = Split(sNewLine, ",")
        If UBound(sLineIn) > 0 Then
            Set oCrit = New COtherCrit
            oCrit.TypeNCID = sLineIn(0)
            oCrit.FieldName = Trim(sLineIn(1))
            If oCrit.FieldName = "" Then
                oCrit.FieldName = "(reserved)"
            End If
            Call GUI.oCritOps.AddCrit(oCrit)
            If lIdx = 0 Then
                GUI.chkCrit(lIdx).Caption = oCrit.FieldName
                Set oCrit.chkControl = GUI.chkCrit(lIdx)
            Else
                Load GUI.chkCrit(lIdx)
                With GUI.chkCrit(lIdx)
                    .Caption = oCrit.FieldName
                    If oCrit.FieldName = "(reserved)" Then
                        .Enabled = False
                    End If
                    Set oCrit.chkControl = GUI.chkCrit(lIdx)
                    .Top = lTop
                    .Left = lLeft
                    .Visible = True
                End With
            End If
            With GUI.chkCrit(lIdx)
                '- adjust for double column
                If .Top >= GUI.Height - 400 - (2 * .Height) Then
                    lLeft = .Left + .Width + 200
                    lTop = GUI.chkCrit(0).Top
                Else
                    lTop = .Top + .Height + 100
                End If
            End With
            lIdx = lIdx + 1
        End If
        
    Loop
End Function


Public Function BuildOtherCrit(ByVal PostMortemPatientID As Long) As String
    Dim oCrit   As COtherCrit
    Dim sAlias1 As String
    Dim sAlias2 As String
    Dim sSQL As String
    Dim bSelected As Boolean
    
    'Select all unit_numbers where the type_ncid and value_text is the same for a given unit_number
    sSQL = "SELECT DISTINCT MMI_GENERICDATA1.UNIT_NUMBER" _
        & " FROM MMI_GENERICDATA, MMI_GENERICDATA MMI_GENERICDATA1" _
        & " Where MMI_GENERICDATA.TYPE_NCID = MMI_GENERICDATA1.TYPE_NCID" _
        & " AND lower(MMI_GENERICDATA.VALUE_TEXT) = lower(MMI_GENERICDATA1.VALUE_TEXT)" _
        & " AND (" _
    
    For Each oCrit In mCol
        If oCrit.chkControl.Value = vbChecked Then
            bSelected = True
            sSQL = sSQL & "(MMI_GENERICDATA.TYPE_NCID = " & oCrit.TypeNCID & ") AND" _
                & " (MMI_GENERICDATA.UNIT_NUMBER = " & PostMortemPatientID & ") OR "
        End If
    Next oCrit
    
    'Remove the last " OR "
    sSQL = Left(sSQL, Len(sSQL) - Len(" OR "))
    
    sSQL = sSQL & ")"
    
    If bSelected = True Then
        BuildOtherCrit = sSQL
    Else
        BuildOtherCrit = ""
    End If
    
End Function
