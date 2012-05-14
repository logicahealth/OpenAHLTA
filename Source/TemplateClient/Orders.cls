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
Attribute VB_Name = "Orders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements TemplateData

Private Const MyLabel As String = "Orders"

Private mListBox As VB.ListBox
Private mItems As Dictionary
Private mbDirty As Boolean

Private oEMCategoryDictionary As Scripting.Dictionary
Private oSpecialtyDictionary As Scripting.Dictionary

Public Function DeserializeCHCSOrders(pOrders As String) As Boolean
    
    Dim objMed As HL7Orders.RxType
    Dim objLab As HL7Orders.LabType
    Dim objRad As HL7Orders.RadType
    
    Dim cData As Collection
    Set cData = DeserializeString(pOrders, gsListDelim)
    
    Select Case UCase(cData(8))
        Case "MED"
            Set objMed = New HL7Orders.RxType
            If Not cData(32) = "" Then objMed.Comment = cData(32)
            If Not cData(28) = "" Then objMed.SigCode = cData(28)
            If objMed.SigCode = "" And Not cData(12) = "" Then objMed.SigCode = cData(12)
            If Not cData(17) = "" Then objMed.Quantity = cData(17)
            If Not cData(18) = "" Then objMed.Refills = cData(18)
            'objMed.PharmacyNameSent = pOrders("")
            'objMed.PharmacyIenSent = pOrders("")
            If cData(31) = "" Then
            ElseIf cData(31) = "YES" Then
                objMed.ChildCap = 1
            ElseIf cData(31) = "NO" Then
                objMed.ChildCap = 0
            End If
            
            If Not cData(26) = "" Then objMed.OrderIen = cData(26)
            If Not cData(27) = "" Then objMed.OrderName = cData(27)
            'Set DeSerializeItem = objMed
            mItems.Add "Med" & objMed.OrderIen & CStr(mItems.Count), objMed
        Case "RX"
            Set objMed = New HL7Orders.RxType
            If Not cData(32) = "" Then objMed.Comment = cData(32)
            If Not cData(28) = "" Then objMed.SigCode = cData(28)
            If objMed.SigCode = "" And Not cData(12) = "" Then objMed.SigCode = cData(12)
            If Not cData(17) = "" Then objMed.Quantity = cData(17)
            If Not cData(18) = "" Then objMed.Refills = cData(18)
            'objMed.PharmacyNameSent = pOrders("")
            'objMed.PharmacyIenSent = pOrders("")
            If cData(31) = "" Then
            ElseIf cData(31) = "YES" Then
                objMed.ChildCap = 1
            ElseIf cData(31) = "NO" Then
                objMed.ChildCap = 0
            End If
            
            If Not cData(26) = "" Then objMed.OrderIen = cData(26)
            If Not cData(27) = "" Then objMed.OrderName = cData(27)
            'Set DeSerializeItem = objMed
            mItems.Add "Med" & objMed.OrderIen & CStr(mItems.Count), objMed
        Case "LAB"
            Set objLab = New HL7Orders.LabType
            If Not cData(20) = "" Then objLab.OrderIen = cData(20)
            If Not cData(21) = "" Then objLab.OrderName = cData(21)
            If Not cData(10) = "" Then objLab.Comment = cData(10)
            
            If cData(23) = "" Then
               objLab.CollectMethod = hlLAB_COLLECT_CLINIC
            ElseIf InStr(1, cData(23), "CLINIC COLLECT", vbTextCompare) > 0 Then
               objLab.CollectMethod = hlLAB_COLLECT_CLINIC
            ElseIf InStr(1, cData(23), "CLINIC SEND", vbTextCompare) > 0 Then
               objLab.CollectMethod = hlLAB_COLLECT_SEND
            End If
            
            If cData(14) = "" Then
               objLab.ProcessPriority = hlLAB_PRIORITY_ROUTINE
            ElseIf InStr(1, cData(14), "ASAP", vbTextCompare) > 0 Then
               objLab.ProcessPriority = hlLAB_PRIORITY_ASAP
            ElseIf InStr(1, cData(14), "NOTIFY", vbTextCompare) > 0 Then
               objLab.ProcessPriority = hlLAB_PRIORITY_NOTIFY
            ElseIf InStr(1, cData(14), "PREOP", vbTextCompare) > 0 Then
               objLab.ProcessPriority = hlLAB_PRIORITY_PREOP
            ElseIf InStr(1, cData(14), "ROUTINE", vbTextCompare) > 0 Then
               objLab.ProcessPriority = hlLAB_PRIORITY_ROUTINE
            ElseIf InStr(1, cData(14), "STAT", vbTextCompare) > 0 Then
               objLab.ProcessPriority = hlLAB_PRIORITY_STAT
            End If
            
            If cData(25) = "" Then
               objLab.LabCollectPriority = 0
            ElseIf InStr(1, cData(25), "ASAP", vbTextCompare) > 0 Then
               objLab.LabCollectPriority = hlLAB_PRIORITY_ASAP
            ElseIf InStr(1, cData(25), "NOTIFY", vbTextCompare) > 0 Then
               objLab.LabCollectPriority = hlLAB_PRIORITY_NOTIFY
            ElseIf InStr(1, cData(25), "PREOP", vbTextCompare) > 0 Then
               objLab.LabCollectPriority = hlLAB_PRIORITY_PREOP
            ElseIf InStr(1, cData(25), "ROUTINE", vbTextCompare) > 0 Then
               objLab.LabCollectPriority = hlLAB_PRIORITY_ROUTINE
            ElseIf InStr(1, cData(25), "STAT", vbTextCompare) > 0 Then
               objLab.LabCollectPriority = hlLAB_PRIORITY_STAT
            End If
            
            If Not cData(22) = "" Then objLab.Specimen.Name = cData(22)
            objLab.Specimen.IEN = cData(40)
            objLab.Specimen.Topography = cData(37)
            objLab.Specimen.TopographyIen = cData(39)
            objLab.Specimen.AnswerCode = "" 'pOrders("")
            objLab.Specimen.AnswerText = "" 'pOrders("")
            objLab.Specimen.CapColorOrContainer = cData(38)
            objLab.Timing.OnexOrContinued = hlONETIME
            objLab.Timing.AdminTime = "" 'Now
            objLab.Timing.Duration = "0"
            objLab.Timing.Frequency = ""
            'objLab.Timing.StartDate = pOrders("")
            'Set DeSerializeItem = objLab
            mItems.Add "Lab" & objLab.OrderIen & CStr(mItems.Count), objLab
        Case "RAD"
            Set objRad = New HL7Orders.RadType
            If Not cData(33) = "" Then objRad.OrderIen = cData(33)
            If Not cData(34) = "" Then objRad.OrderName = cData(34)
            If Not cData(10) = "" Then objRad.Comment = cData(10)
            If Not (cData(36) = "" Or cData(36) = "^") Then objRad.ClinicalImpressionText = cData(36)
            
            If cData(14) = "" Then
               objRad.ProcessPriority = 0
            ElseIf InStr(1, cData(14), "ASAP", vbTextCompare) > 0 Then
               objRad.ProcessPriority = hlLAB_PRIORITY_ASAP
            ElseIf InStr(1, cData(14), "NOTIFY", vbTextCompare) > 0 Then
               objRad.ProcessPriority = hlLAB_PRIORITY_NOTIFY
            ElseIf InStr(1, cData(14), "PREOP", vbTextCompare) > 0 Then
               objRad.ProcessPriority = hlLAB_PRIORITY_PREOP
            ElseIf InStr(1, cData(14), "ROUTINE", vbTextCompare) > 0 Then
               objRad.ProcessPriority = hlLAB_PRIORITY_ROUTINE
            ElseIf InStr(1, cData(14), "STAT", vbTextCompare) > 0 Then
               objRad.ProcessPriority = hlLAB_PRIORITY_STAT
            End If
            
            'objRad.TransportMode = pOrders("")
            If cData(35) = "NO" Then
                objRad.WetReadingWanted = 0
            ElseIf cData(35) = "YES" Then
                objRad.WetReadingWanted = 1
            ElseIf Not cData(35) = "" Then
                objRad.WetReadingWanted = cData(35)
            End If
            
            objRad.Timing.OnexOrContinued = hlONETIME
            objRad.Timing.AdminTime = ""
            objRad.Timing.Duration = "0"
            objRad.Timing.Frequency = ""
            'objRad.Timing.StartDate = pOrders("")
            'Set DeSerializeItem = objRad
            
            mItems.Add "Rad" & objRad.OrderIen & CStr(mItems.Count), objRad
    End Select
    
    Set objMed = Nothing
    Set objLab = Nothing
    Set objRad = Nothing

End Function

Public Property Get EMCategoryDictionary() As Scripting.Dictionary
    Set EMCategoryDictionary = oEMCategoryDictionary
End Property

Public Property Set EMCategoryDictionary(ByVal RHV As Scripting.Dictionary)
    Set oEMCategoryDictionary = RHV
End Property


Public Property Get SpecialtyDictionary() As Scripting.Dictionary
    Set SpecialtyDictionary = oSpecialtyDictionary
End Property

Public Property Set SpecialtyDictionary(ByVal RHV As Scripting.Dictionary)
    Set oSpecialtyDictionary = RHV
End Property


Private Sub TemplateData_DoMoveDown()
    
    Dim sNCIDFrom As String
    Dim sNCIDTo As String
    
    Dim oNCIDTo As Object
    Dim oNCIDFrom As Object
    
    Dim i As Integer
    
    For i = mListBox.ListCount - 1 To 0 Step -1
        If mListBox.Selected(i) Then
            sNCIDFrom = mItems.Keys(i)
            Set oNCIDFrom = mItems(sNCIDFrom)
            sNCIDTo = mItems.Keys(i + 1)
            Set oNCIDTo = mItems(sNCIDTo)
            'SCR-53144 5/4/2004 RBELL Fix data related error
            mItems.Key(sNCIDFrom) = "-1"
            mItems.Key(sNCIDTo) = sNCIDFrom
            Set mItems(sNCIDFrom) = oNCIDFrom
            mItems.Key("-1") = sNCIDTo
            Set mItems(sNCIDTo) = oNCIDTo
            mDisplayItems
            mListBox.Selected(i + 1) = True
            mbDirty = True
        End If
    Next i

End Sub

Private Sub TemplateData_DoMoveUp()
    
    Dim sNCIDFrom As String
    Dim sNCIDTo As String
    
    Dim oNCIDTo As Object
    Dim oNCIDFrom As Object
    
    Dim i As Integer
    
    For i = mListBox.ListCount - 1 To 0 Step -1
        If mListBox.Selected(i) Then
            sNCIDFrom = mItems.Keys(i)
            Set oNCIDFrom = mItems(sNCIDFrom)
            sNCIDTo = mItems.Keys(i - 1)
            Set oNCIDTo = mItems(sNCIDTo)
            'SCR-53144 5/4/2004 RBELL Fix data related error
            mItems.Key(sNCIDFrom) = "-1"
            mItems.Key(sNCIDTo) = sNCIDFrom
            Set mItems(sNCIDFrom) = oNCIDFrom
            mItems.Key("-1") = sNCIDTo
            Set mItems(sNCIDTo) = oNCIDTo
            mDisplayItems
            mListBox.Selected(i - 1) = True
            mbDirty = True
            Exit For
        End If
    Next i

End Sub


Private Function TemplateData_Export() As String
    Dim colItems As Collection
    Dim i As Integer
    Dim bForceHL7Export As Boolean
    
    Set colItems = New Collection
    If gobjShared Is Nothing Then Set gobjShared = New CWShared
    ''''<SCR 37645 Alberstein - Route items to correct serialization procedure
    If mItems.Count > 0 Then
      '''if app is in PGUI ITT mode
      If gobjShared.IsAppMode(modeCHCSI_GUI) And gobjShared.IsAppMode(modeITT) Then
         If TypeOf mItems.Items(i) Is GEMS_OE.RxType Then
         ElseIf TypeOf mItems.Items(i) Is GEMS_OE.LabType Then
         ElseIf TypeOf mItems.Items(i) Is GEMS_OE.RadType Then
         Else
            '''If not one of the GEMS_OE types - must be imported templates that are HL7Orders types - Use SerializeItem
            bForceHL7Export = True
         End If
      End If
    End If
    '>
    For i = 0 To mItems.Count - 1
        If Not bForceHL7Export Then '<SCR 37645 Alberstein
           colItems.Add GEMS_SerializeItem(Left$(mItems.Keys(i), 3), mItems.Items(i))
        End If
    Next i
    TemplateData_Export = SerializeCollection(colItems, gsListDelim)
    Set colItems = Nothing
End Function

Private Function SerializeItem(ByVal vsWhat As String, rvItem As Variant) As String
    
    Dim objMed As HL7Orders.RxType
    Dim objLab As HL7Orders.LabType
    Dim objRad As HL7Orders.RadType
    Dim colReturn As Collection
    
    Set colReturn = New Collection
    colReturn.Add vsWhat
    Select Case vsWhat
    Case "Med"
        Set objMed = rvItem
        colReturn.Add objMed.Comment
        colReturn.Add objMed.SigCode
        colReturn.Add objMed.Quantity
        colReturn.Add objMed.Refills
        colReturn.Add objMed.PharmacyNameSent
        colReturn.Add objMed.PharmacyIenSent
        colReturn.Add objMed.ChildCap
        colReturn.Add objMed.OrderIen
        colReturn.Add objMed.OrderName
    Case "Lab"
        Set objLab = rvItem
        colReturn.Add objLab.OrderIen
        colReturn.Add objLab.OrderName
        colReturn.Add objLab.Comment
        colReturn.Add objLab.CollectMethod
        colReturn.Add objLab.ProcessPriority
        colReturn.Add objLab.LabCollectPriority
        colReturn.Add objLab.Specimen.Name
        colReturn.Add objLab.Specimen.IEN
        colReturn.Add objLab.Specimen.Topography
        colReturn.Add objLab.Specimen.TopographyIen
        colReturn.Add objLab.Specimen.AnswerCode
        colReturn.Add objLab.Specimen.AnswerText
        colReturn.Add objLab.Specimen.CapColorOrContainer
        colReturn.Add objLab.Timing.OnexOrContinued
        colReturn.Add objLab.Timing.AdminTime
        colReturn.Add objLab.Timing.Duration
        colReturn.Add objLab.Timing.Frequency
        colReturn.Add objLab.Timing.StartDate
    Case "Rad"
        Set objRad = rvItem
        colReturn.Add objRad.OrderIen
        colReturn.Add objRad.OrderName
        colReturn.Add objRad.Comment
        colReturn.Add objRad.ClinicalImpressionText
        colReturn.Add objRad.ProcessPriority
        colReturn.Add objRad.TransportMode
        colReturn.Add objRad.WetReadingWanted
        colReturn.Add objRad.Timing.OnexOrContinued
        colReturn.Add objRad.Timing.AdminTime
        colReturn.Add objRad.Timing.Duration
        colReturn.Add objRad.Timing.Frequency
        colReturn.Add objRad.Timing.StartDate
    End Select
    SerializeItem = SerializeCollection(colReturn, gsRecordDelim)
    Set colReturn = Nothing
    Set objMed = Nothing
    Set objLab = Nothing
    Set objRad = Nothing
End Function
Private Function GEMS_SerializeItem(ByVal vsWhat As String, rvItem As Variant) As String
    
    Dim objMed As GEMS_OE.RxType
    Dim objLab As GEMS_OE.LabType
    Dim objRad As GEMS_OE.RadType
    Dim colReturn As Collection
    
    Set colReturn = New Collection
    colReturn.Add vsWhat
    Select Case vsWhat
    Case "Med"
        Set objMed = rvItem
        colReturn.Add objMed.Comment
        colReturn.Add objMed.SigCode
        colReturn.Add objMed.Quantity
        colReturn.Add objMed.Refills
        colReturn.Add objMed.PharmacyNameSent
        colReturn.Add objMed.PharmacyIenSent
        colReturn.Add objMed.ChildCap
        colReturn.Add objMed.OrderIen
        colReturn.Add objMed.OrderName
    Case "Lab"
        Set objLab = rvItem
        colReturn.Add objLab.OrderIen
        colReturn.Add objLab.OrderName
        colReturn.Add objLab.Comment
        colReturn.Add objLab.CollectMethod
        colReturn.Add objLab.ProcessPriority
        colReturn.Add objLab.LabCollectPriority
        colReturn.Add objLab.Specimen.Name
        colReturn.Add objLab.Specimen.IEN
        colReturn.Add objLab.Specimen.Topography
        colReturn.Add objLab.Specimen.TopographyIen
        colReturn.Add objLab.Specimen.AnswerCode
        colReturn.Add objLab.Specimen.AnswerText
        colReturn.Add objLab.Specimen.CapColorOrContainer
        colReturn.Add objLab.Timing.OnexOrContinued
        colReturn.Add objLab.Timing.AdminTime
        colReturn.Add objLab.Timing.Duration
        colReturn.Add objLab.Timing.Frequency
        colReturn.Add objLab.Timing.StartDate
    Case "Rad"
        Set objRad = rvItem
        colReturn.Add objRad.OrderIen
        colReturn.Add objRad.OrderName
        colReturn.Add objRad.Comment
        colReturn.Add objRad.ClinicalImpressionText
        colReturn.Add objRad.ProcessPriority
        colReturn.Add objRad.TransportMode
        colReturn.Add objRad.WetReadingWanted
        colReturn.Add objRad.Timing.OnexOrContinued
        colReturn.Add objRad.Timing.AdminTime
        colReturn.Add objRad.Timing.Duration
        colReturn.Add objRad.Timing.Frequency
        colReturn.Add objRad.Timing.StartDate
    End Select
    GEMS_SerializeItem = SerializeCollection(colReturn, gsRecordDelim)
    Set colReturn = Nothing
    Set objMed = Nothing
    Set objLab = Nothing
    Set objRad = Nothing
End Function

Private Function TemplateData_Import(rsString As String) As Boolean
    TemplateData_Import = TemplateData_DeSerialize(rsString)
End Function

Private Function TemplateData_GetFromEncounter(objEncounter As CHCSIIEncounterOps.Encounter) As Boolean
    
    Dim objMed As Object
    Dim objLab As Object
    Dim objRad As Object
    Dim objSQL As ICHCSII_SQL.ISqlOps
    Dim moConn As CHCSII_CONN.Conn
        
    On Error Resume Next
    If gobjShared Is Nothing Then Set gobjShared = New CWShared
    Set moConn = New CHCSII_CONN.Conn
    Set objSQL = moConn.CHCSII_SQLOPS(Auto)
        
    objSQL.Execute "select * from oe_rx" _
         & " where enc_fac_NCID = '" & objEncounter.FacilityNCID & "'" _
         & " and enc_num = '" & objEncounter.EncounterID & "'"
    Do While Not objSQL.EOF
        Set objMed = New GEMS_OE.RxType
        
        objMed.Comment = objSQL("rx_comment")
        objMed.SigCode = objSQL("sig_code")
        objMed.Quantity = objSQL("Quantity")
        objMed.Refills = objSQL("Refills")
        objMed.PharmacyNameSent = objSQL("pharmacy_name_sent")
        objMed.PharmacyIenSent = objSQL("pharmacy_ien_sent")
        objMed.ChildCap = objSQL("child_cap")
        objMed.OrderIen = objSQL("order_ien")
        objMed.OrderName = objSQL("order_name")
        mItems.Add "Med" & objMed.OrderIen, objMed
        objSQL.MoveNext
    Loop
    objSQL.Execute "select * from oe_lab" _
         & " where enc_fac_NCID = '" & objEncounter.FacilityNCID & "'" _
         & " and enc_num = '" & objEncounter.EncounterID & "'"
    Do While Not objSQL.EOF
        Set objLab = New GEMS_OE.LabType
        
        objLab.OrderIen = objSQL("order_ien")
        objLab.OrderName = objSQL("order_name")
        objLab.Comment = objSQL("lab_comment")
        objLab.CollectMethod = objSQL("collect_method")
        objLab.ProcessPriority = objSQL("process_priority")
        objLab.LabCollectPriority = objSQL("lab_collect_priority")
        objLab.Specimen.Name = objSQL("specimen_name")
        objLab.Specimen.IEN = objSQL("specimen_ien")
        objLab.Specimen.Topography = objSQL("specimen_topography")
        objLab.Specimen.TopographyIen = objSQL("specimen_topography_ien")
        objLab.Specimen.AnswerCode = objSQL("specimen_answer_code")
        objLab.Specimen.AnswerText = objSQL("specimen_answer_text")
        objLab.Specimen.CapColorOrContainer = objSQL("specimen_capcolor_or_container")
        objLab.Timing.OnexOrContinued = objSQL("timing_one_x_or_continued")
        objLab.Timing.AdminTime = objSQL("timing_admin_time")
        objLab.Timing.Duration = objSQL("timing_duration")
        objLab.Timing.Frequency = objSQL("timing_frequency")
        objLab.Timing.StartDate = objSQL("timing_start_date")
        objLab.Timing.UnexpandedTime = objSQL("Unexpanded_Time")
        mItems.Add "Lab" & objLab.OrderIen, objLab
        objSQL.MoveNext
    Loop
    objSQL.Execute "select * from oe_rad" _
         & " where enc_fac_NCID = '" & objEncounter.FacilityNCID & "'" _
         & " and enc_num = '" & objEncounter.EncounterID & "'"
    Do While Not objSQL.EOF
        Set objRad = New GEMS_OE.RadType
        
        objRad.OrderIen = objSQL("order_ien")
        objRad.OrderName = objSQL("order_name")
        objRad.Comment = objSQL("rad_comment")
        objRad.ClinicalImpressionText = objSQL("clinical_impression_text")
        objRad.ProcessPriority = objSQL("process_priority")
        objRad.TransportMode = objSQL("transport_mode")
        objRad.WetReadingWanted = objSQL("wet_reading_wanted")
        objRad.Timing.OnexOrContinued = objSQL("timing_one_x_or_continued")
        objRad.Timing.AdminTime = objSQL("timing_admin_time")
        objRad.Timing.Duration = objSQL("timing_duration")
        objRad.Timing.Frequency = objSQL("timing_frequency")
        objRad.Timing.StartDate = objSQL("timing_start_date")
        mItems.Add "Rad" & objRad.OrderIen, objRad
        objSQL.MoveNext
    Loop
Cleanup:
    Set objLab = Nothing
    Set objRad = Nothing
    Set objSQL = Nothing
    Set objMed = Nothing
    Set moConn = Nothing

End Function

Private Function TemplateData_GetQueryCriteria() As String
End Function

Private Sub mDisplayItems()
    Dim i As Integer
    mListBox.Clear
    For i = 0 To mItems.Count - 1
        If IsObject(mItems.Items(i)) Then
        mListBox.AddItem "(" & Left$(mItems.Keys(i), 3) & ")" & mItems.Items(i).OrderName
        mListBox.ItemData(mListBox.NewIndex) = Mid(mItems.Keys(i), 4)
        End If
    Next i
End Sub

Private Sub Class_Initialize()
    Set mItems = New Dictionary
End Sub

Private Sub Class_Terminate()
    Set mItems = Nothing
    Set mListBox = Nothing
    Set oEMCategoryDictionary = Nothing
    Set oSpecialtyDictionary = Nothing
End Sub

Private Property Get TemplateData_CanEdit() As Boolean
    TemplateData_CanEdit = False
End Property

Private Function TemplateData_DeSerialize(rvData As Variant) As Boolean
    Dim colItems As Collection
    Dim sValue As Variant
    Dim objItem As Object
    
    'SCR-31178 Preventing merge from working properly Set mItems = New Dictionary
    Set colItems = DeserializeString(CStr(rvData), gsListDelim)
    For Each sValue In colItems
        On Error Resume Next
        Set objItem = GEMS_DeSerializeItem(sValue)
        If TypeOf objItem Is GEMS_OE.RxType Then
            mItems.Add "Med" & objItem.OrderIen, objItem
        ElseIf TypeOf objItem Is GEMS_OE.LabType Then
            mItems.Add "Lab" & objItem.OrderIen, objItem
        Else
            mItems.Add "Rad" & objItem.OrderIen, objItem
        End If
        
        On Error GoTo 0
    Next sValue
    If Not mListBox Is Nothing Then
        Call mDisplayItems
    End If
    mbDirty = False
    TemplateData_DeSerialize = True
End Function

Private Function DeSerializeItem(rvItem As Variant) As Object
    Dim objMed As HL7Orders.RxType
    Dim objLab As HL7Orders.LabType
    Dim objRad As HL7Orders.RadType
    Dim colFields As Collection
    Set colFields = DeserializeString(CStr(rvItem), gsRecordDelim)
    Select Case colFields(1)
    Case "Med"
        Set objMed = New HL7Orders.RxType
        objMed.Comment = colFields(2)
        objMed.SigCode = colFields(3)
        objMed.Quantity = colFields(4)
        objMed.Refills = colFields(5)
        objMed.PharmacyNameSent = colFields(6)
        objMed.PharmacyIenSent = colFields(7)
        objMed.ChildCap = colFields(8)
        objMed.OrderIen = colFields(9)
        objMed.OrderName = colFields(10)
        Set DeSerializeItem = objMed
    Case "Lab"
        Set objLab = New HL7Orders.LabType
        objLab.OrderIen = colFields(2)
        objLab.OrderName = colFields(3)
        objLab.Comment = colFields(4)
        objLab.CollectMethod = colFields(5)
        objLab.ProcessPriority = colFields(6)
        objLab.LabCollectPriority = colFields(7)
        objLab.Specimen.Name = colFields(8)
        objLab.Specimen.IEN = colFields(9)
        objLab.Specimen.Topography = colFields(10)
        objLab.Specimen.TopographyIen = colFields(11)
        objLab.Specimen.AnswerCode = colFields(12)
        objLab.Specimen.AnswerText = colFields(13)
        objLab.Specimen.CapColorOrContainer = colFields(14)
        objLab.Timing.OnexOrContinued = colFields(15)
        objLab.Timing.AdminTime = colFields(16)
        objLab.Timing.Duration = colFields(17)
        objLab.Timing.Frequency = colFields(18)
        objLab.Timing.StartDate = colFields(19)
        Set DeSerializeItem = objLab
    Case "Rad"
        Set objRad = New HL7Orders.RadType
        objRad.OrderIen = colFields(2)
        objRad.OrderName = colFields(3)
        objRad.Comment = colFields(4)
        objRad.ClinicalImpressionText = colFields(5)
        objRad.ProcessPriority = colFields(6)
        objRad.TransportMode = colFields(7)
        objRad.WetReadingWanted = colFields(8)
        objRad.Timing.OnexOrContinued = colFields(9)
        objRad.Timing.AdminTime = colFields(10)
        objRad.Timing.Duration = colFields(11)
        objRad.Timing.Frequency = colFields(12)
        objRad.Timing.StartDate = colFields(13)
        Set DeSerializeItem = objRad
    End Select
    Set colFields = Nothing
    Set objMed = Nothing
    Set objLab = Nothing
    Set objRad = Nothing
End Function

Private Function GEMS_DeSerializeItem(rvItem As Variant) As Object
    Dim objMed As GEMS_OE.RxType
    Dim objLab As GEMS_OE.LabType
    Dim objRad As GEMS_OE.RadType
    Dim colFields As Collection
    Set colFields = DeserializeString(CStr(rvItem), gsRecordDelim)
    Select Case colFields(1)
    Case "Med"
        Set objMed = New GEMS_OE.RxType
        objMed.Comment = colFields(2)
        objMed.SigCode = colFields(3)
        objMed.Quantity = colFields(4)
        objMed.Refills = colFields(5)
        objMed.PharmacyNameSent = colFields(6)
        objMed.PharmacyIenSent = colFields(7)
        objMed.ChildCap = colFields(8)
        objMed.OrderIen = colFields(9)
        objMed.OrderName = colFields(10)
        Set GEMS_DeSerializeItem = objMed
    Case "Lab"
        Set objLab = New GEMS_OE.LabType
        objLab.OrderIen = colFields(2)
        objLab.OrderName = colFields(3)
        objLab.Comment = colFields(4)
        objLab.CollectMethod = colFields(5)
        objLab.ProcessPriority = colFields(6)
        objLab.LabCollectPriority = colFields(7)
        objLab.Specimen.Name = colFields(8)
        objLab.Specimen.IEN = colFields(9)
        objLab.Specimen.Topography = colFields(10)
        objLab.Specimen.TopographyIen = colFields(11)
        objLab.Specimen.AnswerCode = colFields(12)
        objLab.Specimen.AnswerText = colFields(13)
        objLab.Specimen.CapColorOrContainer = colFields(14)
        objLab.Timing.OnexOrContinued = colFields(15)
        objLab.Timing.AdminTime = colFields(16)
        objLab.Timing.Duration = colFields(17)
        objLab.Timing.Frequency = colFields(18)
        objLab.Timing.StartDate = colFields(19)
        Set GEMS_DeSerializeItem = objLab
    Case "Rad"
        Set objRad = New GEMS_OE.RadType
        objRad.OrderIen = colFields(2)
        objRad.OrderName = colFields(3)
        objRad.Comment = colFields(4)
        objRad.ClinicalImpressionText = colFields(5)
        objRad.ProcessPriority = colFields(6)
        objRad.TransportMode = colFields(7)
        objRad.WetReadingWanted = colFields(8)
        objRad.Timing.OnexOrContinued = colFields(9)
        objRad.Timing.AdminTime = colFields(10)
        objRad.Timing.Duration = colFields(11)
        objRad.Timing.Frequency = colFields(12)
        objRad.Timing.StartDate = colFields(13)
        Set GEMS_DeSerializeItem = objRad
    End Select
    Set colFields = Nothing
    Set objMed = Nothing
    Set objLab = Nothing
    Set objRad = Nothing
End Function

Private Sub TemplateData_DoAdd()
    
End Sub

Private Sub TemplateData_DoEdit()
End Sub

Private Sub TemplateData_DoRemove()
    Dim sKey As String
    Dim i As Integer
    For i = mListBox.ListCount - 1 To 0 Step -1
        If mListBox.Selected(i) Then
            sKey = Mid(mListBox.List(i), 2, 3) & CStr(mListBox.ItemData(i))
            mItems.Remove sKey
            mListBox.RemoveItem i
            mbDirty = True
        End If
    Next i
End Sub

Private Property Let TemplateData_isDirty(RHS As Boolean)
    mbDirty = RHS
End Property

Private Property Get TemplateData_isDirty() As Boolean
    TemplateData_isDirty = mbDirty
End Property

Private Property Get TemplateData_Items() As Scripting.IDictionary
    Set TemplateData_Items = mItems
End Property

Private Property Get TemplateData_Label() As String
    TemplateData_Label = MyLabel
End Property

Private Property Set TemplateData_ListBox(RHS As Object)
    Set mListBox = RHS
    Call mDisplayItems
End Property

Private Property Get TemplateData_ListBox() As Object
    Set TemplateData_ListBox = mListBox
End Property

Private Function TemplateData_Serialize(rvData As Variant) As Boolean
    rvData.Add TemplateData_Export
    mbDirty = False
End Function

