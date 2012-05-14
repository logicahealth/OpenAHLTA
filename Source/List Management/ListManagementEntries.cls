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
Attribute VB_Name = "Entries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"ListManagementEntry"
Option Explicit

'local variable to hold collection


Public Enum EnumType
    stUserComplaints = &HFFF10110
    stUserDiagnosis
    stUserEandMs
    stUserProcedures
    stClinicComplaints
    stClinicDiagnosis
    stClinicEandMs
    stClinicProcedures
    stFacilityComplaints
    stFacilityDiagnosis
    stFacilityEandMs
    stFacilityProcedures
End Enum

Private mCol As Collection
Private IC As New IConvert

Private mbShare As Boolean        'Ignored
Private mbSortByICD9 As Boolean   'Ignored
Public Enum EntriesSort
    esByCode = -1
    esByDescription = -2
    esByNone = 0
End Enum
Public Sort As EntriesSort '{{BOOKMARK TODO: When these values change reorder the list}}
Public Order As Constants_SortIndicator

Private mvarType As EnumType
Private mbIsDirty As Boolean
Private mlOwnerNCID As Long
Private mszKey As String
Private mszSubKey As String
Private mlSaveOwnerNCID As Long

Private mbAllowRead As Boolean
Private mbAllowWrite As Boolean
Private mbAllowDelete As Boolean


Public Property Get AllowDelete() As Boolean
    AllowDelete = mbAllowDelete
End Property
Public Property Get AllowWrite() As Boolean
    AllowWrite = mbAllowWrite
End Property

Public Function MoveItemDown(vntIndexKey As Variant) As Boolean
'  SCR-34313 ADDED SUB FOR NEW FUNCTIONALITYRBELL 4/3/3
    
    Dim sz As String
    Dim oEntry As Entry
    Dim iFindIndex As Integer
    
    Dim sFromClinincNCID As String
    Dim oToEntry As Entry
    
    On Error GoTo xitErr
    MoveItemDown = False
    
    If mbAllowWrite = False Then
        Err.Raise vbObjectError + 10, App.Title, "Insufficient Access Rights."
    End If

    On Error Resume Next
    Set oEntry = mCol(vntIndexKey)
    If Err Then
       Err.Clear
       Exit Function
    End If
    
    For iFindIndex = 1 To mCol.Count
        If mCol(iFindIndex).Key = oEntry.Key Then
           If mCol.Count >= iFindIndex + 1 Then
              Set oToEntry = New Entry
              oToEntry.Init mCol(iFindIndex + 1).SnoID, mCol(iFindIndex + 1).PREFIX, mCol(iFindIndex + 1).DESCRIPTION, mCol(iFindIndex + 1).NCID, mCol(iFindIndex + 1).Code, mCol(iFindIndex + 1).ClinicNCID, mCol(iFindIndex + 1).Key
           End If
           Exit For
        End If
    Next
    
    If Not oToEntry Is Nothing Then
        mCol.Remove iFindIndex
        mCol.Remove iFindIndex
        If iFindIndex = 1 Then
            'If destination is the first item in the collection
          If mCol.Count = 0 Then
            mCol.Add oToEntry, oToEntry.Key
          Else
            mCol.Add oToEntry, oToEntry.Key, 1
          End If
        Else
            mCol.Add oToEntry, oToEntry.Key, , iFindIndex - 1
        End If
        mCol.Add oEntry, oEntry.Key, , iFindIndex
        mbIsDirty = True
    End If
    
    MoveItemDown = True
    Exit Function
    
xitErr:

    sz = "List Management Move failed for OwnerNCID: %1,  Key: %2, SubKey: %3"
    sz = Replace(sz, "%1", mlOwnerNCID)
    sz = Replace(sz, "%2", mszKey)
    sz = Replace(sz, "%3", mszSubKey)
    gobjShared.ShowVBError Err.Number, Err.DESCRIPTION & vbCrLf & sz, "MMMListManagement.Entries.Remove", App.Title, vbExclamation
    'Resume Next
End Function

Public Function MoveItemUp(vntIndexKey As Variant) As Boolean
'  SCR-34313 ADDED SUB FOR NEW FUNCTIONALITYRBELL 4/3/3
    
    Dim sz As String
    Dim oEntry As Entry
    Dim iFindIndex As Integer
    
    Dim sFromClinincNCID As String
    Dim oToEntry As Entry
    
    On Error GoTo xitErr
    MoveItemUp = False
    
    If mbAllowWrite = False Then
        Err.Raise vbObjectError + 10, App.Title, "Insufficient Access Rights."
    End If

    On Error Resume Next
    Set oEntry = mCol(vntIndexKey)
    If Err Then
       Err.Clear
       Exit Function
    End If
    
    For iFindIndex = 1 To mCol.Count
        If mCol(iFindIndex).Key = oEntry.Key Then
           If mCol.Count > 1 Then
              Set oToEntry = New Entry
              oToEntry.Init mCol(iFindIndex - 1).SnoID, mCol(iFindIndex - 1).PREFIX, mCol(iFindIndex - 1).DESCRIPTION, mCol(iFindIndex - 1).NCID, mCol(iFindIndex - 1).Code, mCol(iFindIndex - 1).ClinicNCID, mCol(iFindIndex - 1).Key
           End If
           Exit For
        End If
    Next
    
    If Not oToEntry Is Nothing Then
        mCol.Remove iFindIndex - 1
        mCol.Remove iFindIndex - 1
        If iFindIndex = 2 Then
          'If destination is the first item in the collection
          If mCol.Count = 0 Then
            mCol.Add oToEntry, oToEntry.Key
          Else
            mCol.Add oToEntry, oToEntry.Key, 1
          End If
        Else
            mCol.Add oToEntry, oToEntry.Key, , iFindIndex - 2
        End If
        mCol.Add oEntry, oEntry.Key, iFindIndex - 1
        mbIsDirty = True
    End If
    
    MoveItemUp = True
    Exit Function
    
xitErr:

    sz = "List Management Move failed for OwnerNCID: %1,  Key: %2, SubKey: %3"
    sz = Replace(sz, "%1", mlOwnerNCID)
    sz = Replace(sz, "%2", mszKey)
    sz = Replace(sz, "%3", mszSubKey)
    gobjShared.ShowVBError Err.Number, Err.DESCRIPTION & vbCrLf & sz, "MMMListManagement.Entries.MoveItemUp", App.Title, vbExclamation
    Resume Next

End Function


Public Function Retrieve(ByVal Key As String, ByVal SubKey As String, _
                         ByVal OwnerNCID As Long, ByVal SaveOwnerNCID As Long, _
                         ByVal AllowRead As Boolean, ByVal AllowWrite As Boolean, ByVal AllowDelete As Boolean) As Boolean

    
    Dim bRet As Boolean
    Dim a As Variant
    Dim b As Variant
    Dim i As Long
    Dim lSnoID As String
    Dim szPrefix As String
    Dim sz As String
    Dim szBuf As String
    Dim lClinicNCID As Long
    Dim szSnoID  As String
    
    On Error GoTo xitErr
    
    Set mCol = New Collection
    
    mlSaveOwnerNCID = SaveOwnerNCID
    mlOwnerNCID = OwnerNCID
    mszSubKey = SubKey
    mszKey = Key
    
    'Security
    mbAllowRead = AllowRead
    mbAllowWrite = AllowWrite
    mbAllowDelete = AllowDelete
    
    If mbAllowRead = False Then Exit Function
    
    bRet = gobjConfig.GetValue(NCID_LIST_MGMT_APPLICATION, mszKey, mszSubKey, szBuf)
    
    a = Split(szBuf, PARSE_DELIMITER)
    
    If UBound(a) >= 0 Then
        Select Case a(0)
        Case Constants_SortIndicator.ssSortIndicatorAscending
            Me.Order = ssSortIndicatorAscending
        Case Constants_SortIndicator.ssSortIndicatorDescending
            Me.Order = ssSortIndicatorDescending
        Case Else
            Me.Order = ssSortIndicatorNone
        End Select
    End If
    
    If UBound(a) >= 1 Then
        Select Case a(1)
        Case "True"
            Me.Sort = esByCode
        Case "False"
            Me.Sort = esByDescription
        Case esByCode, esByDescription, esNone
            Me.Sort = a(1)
        Case Else
            Me.Sort = esByDescription
        End Select
    End If
    
    For i = 2 To UBound(a)
        
        b = Split(a(i), PREFIX_DELIMITER)
        
        If UBound(b) >= 0 Then szSnoID = b(0)
        If UBound(b) >= 1 Then szPrefix = b(1)
        
        Call AddInternal(szSnoID, szPrefix)
        
    Next
    
    mbIsDirty = False
    

    Exit Function
    
xitErr:

    sz = "List Management Retrieve failed for OwnerNCID: %1,  Key: %2, SubKey: %3"
    sz = Replace(sz, "%1", mlOwnerNCID)
    sz = Replace(sz, "%2", mszKey)
    sz = Replace(sz, "%3", mszSubKey)
    gobjShared.ShowVBError Err.Number, Err.DESCRIPTION & vbCrLf & sz, "MMMListManagement.Entries.Retrieve", App.Title, vbExclamation
    
    Exit Function
    
End Function

Private Function FindDuplicate(ByVal SnoID As String, ByVal PREFIX As String) As MMMListManagement.Entry
    Dim LME As MMMListManagement.Entry
    
    Set FindDuplicate = Nothing
    
    For Each LME In mCol
    
        If LME.SnoID = SnoID Then
            If LME.PREFIX = PREFIX Then
                Set FindDuplicate = LME
                Exit For
            End If
        End If
        
    Next
    
End Function
Private Function AddInternal(ByVal SnoID As String, ByVal PREFIX As String) As MMMListManagement.Entry
    Dim bAllowWrite
    
    On Error GoTo errXIT
    
    bAllowWrite = mbAllowWrite
    mbAllowWrite = True
    
    Set AddInternal = Add(SnoID, PREFIX)
    
errXIT:
    mbAllowWrite = bAllowWrite
End Function

'JT BEGIN SCR 42399 - Added support for diagnosis added from problem list
Public Function Add(ByVal SnoID As String, ByVal PREFIX As String) As MMMListManagement.Entry
    
    'create a new object
    Dim objNewMember As MMMListManagement.Entry
    Dim szkey As String
    Dim sz As String
    Dim szCode As String
    Dim szDesc As String
    Dim bNotFromMedcin As Boolean
    
    On Error GoTo xitErr
    
    If mbAllowWrite = False Then
        Err.Raise vbObjectError + 10, App.Title, "Insufficient Access Rights."
    End If

    'Determine if the medcin id is really an ICD9 code
    If InStr(1, mszKey, "DiagnosisList", vbTextCompare) > 0 And InStr(1, SnoID, ".", vbTextCompare) > 0 Then
      If Right(SnoID, 1) = "." Then SnoID = Left(SnoID, Len(SnoID) - 1)
      bNotFromMedcin = True
    End If
    
    'Suppress Dups
    Set objNewMember = FindDuplicate(SnoID, PREFIX)
    If Not objNewMember Is Nothing Then
        Set Add = objNewMember
        Exit Function
    End If
    
    'Make A New One
    Set objNewMember = New MMMListManagement.Entry

    'set the properties passed into the method
    szkey = PREFIX & PREFIX_DELIMITER & SnoID
    
    If InStr(1, mszKey, "ProcedureList", vbTextCompare) > 0 Then
    
       szCode = IC.SnoCPT4(SnoID)
       szDesc = IC.CPTDescription(SnoID)
       
    ElseIf InStr(1, mszKey, "DiagnosisList", vbTextCompare) > 0 Then
       
      If bNotFromMedcin Then
         szCode = SnoID
         szDesc = GetICD9Desc(szCode)
      Else
         'This is a real medcin id - handle as such
         szCode = IC.SnoICD9(SnoID)
         szDesc = IC.SnoDescription(SnoID)
       End If
              
    ElseIf InStr(1, mszKey, "ComplaintList", vbTextCompare) > 0 Then
        szCode = IC.SnoICD9(SnoID)
        szDesc = IC.SnoDescription(SnoID)
    Else
        szCode = IC.SnoCPT4(SnoID)
        If szCode = vbNullString Then
            If IsNumeric(SnoID) Then
                szCode = IC.SnoICD9(SnoID)
            End If
        End If
        szDesc = IC.SnoDescription(SnoID)
    End If
    
    objNewMember.Init SnoID, PREFIX, szDesc, 0, szCode, 0, szkey, bNotFromMedcin
                      
    mCol.Add objNewMember, szkey

    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing

    mbIsDirty = True
    
    Exit Function
    
xitErr:

    If Err.Number = vbObjectError + 10 Then
        Err.Raise Err.Number, Err.Source, Err.DESCRIPTION
    Else
        sz = "List Management Add failed for OwnerNCID: %1,  Key: %2, SubKey: %3"
        sz = Replace(sz, "%1", mlOwnerNCID)
        sz = Replace(sz, "%2", mszKey)
        sz = Replace(sz, "%3", mszSubKey)
        gobjShared.ShowVBError Err.Number, Err.DESCRIPTION & vbCrLf & sz, "MMMListManagement.Entries.Add", App.Title, vbExclamation
    End If
    
    Exit Function
    Resume
End Function
'JT END SCR 42399 - Added support for diagnosis added from problem list

Public Property Get Item(vntIndexKey As Variant) As MMMListManagement.Entry
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


Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)

    Dim sz As String
    On Error GoTo xitErr

    If mbAllowDelete = False Then
        Err.Raise vbObjectError + 10, App.Title, "Insufficient Access Rights."
    End If

    mCol.Remove vntIndexKey
    mbIsDirty = True
    Exit Sub
    
xitErr:

    sz = "List Management Remove failed for OwnerNCID: %1,  Key: %2, SubKey: %3"
    sz = Replace(sz, "%1", mlOwnerNCID)
    sz = Replace(sz, "%2", mszKey)
    sz = Replace(sz, "%3", mszSubKey)
    gobjShared.ShowVBError Err.Number, Err.DESCRIPTION & vbCrLf & sz, "MMMListManagement.Entries.Remove", App.Title, vbExclamation
    
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Call Save
    Set mCol = Nothing
End Sub
'Return True if SetValue Success
Friend Function Save() As Boolean

    Dim bRet As Boolean
    Dim sz As String
    
    On Error GoTo xitErr
    
    If mbIsDirty = False Then Exit Function
        
    bRet = gobjConfig.SetValue(NCID_LIST_MGMT_APPLICATION, _
                    mlSaveOwnerNCID, mszKey, mszSubKey, PersistString())
    
    Debug.Print mlOwnerNCID, mszKey, mszSubKey, PersistString()
    
    If Not bRet Then GoTo xitErr
    
    Save = True
    Exit Function
    
xitErr:

    sz = "List Management Save failed for OwnerNCID: %1,  Key: %2, SubKey: %3"
    sz = Replace(sz, "%1", mlOwnerNCID)
    sz = Replace(sz, "%2", mszKey)
    sz = Replace(sz, "%3", mszSubKey)
    gobjShared.ShowVBError Err.Number, Err.DESCRIPTION & vbCrLf & sz, "MMMListManagement.Entries.Save", App.Title, vbExclamation
    
    Save = False
    Exit Function
    
End Function
Private Function PersistString() As String

    Dim sz As String
    Dim LME As MMMListManagement.Entry
    
    sz = Me.Order & PARSE_DELIMITER & Me.Sort
    
    For Each LME In mCol
        If LME.NotFromMedcin Then
            'The Medcin ID is really an ICD9 - save with a . appended so we can identify later
            sz = sz & PARSE_DELIMITER & LME.SnoID & "."
         Else
            sz = sz & PARSE_DELIMITER & LME.SnoID ' & PREFIX_DELIMITER & LME.PREFIX '{{BOOKMARK TODO After complete trasition, turn this on.}}
        End If
    Next

    PersistString = sz
    
End Function

Public Property Get IsDirty() As Boolean
    IsDirty = mbIsDirty
End Property

'JT BEGIN SCR 42399 - Lookup the diagnosis description from ICD9 code
Private Function GetICD9Desc(ByVal sICD9Code As String) As String

On Error GoTo Error_Handler

Dim sSql As String
Dim objConn    As CHCSII_CONN.Conn
Dim objDAS     As ICHCSII_DataAccess.ICHCSII_DAS
Dim adoRS      As ADODB.Recordset

   sSql = "SELECT description " & _
            " FROM vICD " & _
            " WHERE code = '" & Replace(sICD9Code, "'", "''") & "'"
                        
   'Open the recordset
   Set objConn = New CHCSII_CONN.Conn
   Set objDAS = objConn.CHCSII_DAS
   Set adoRS = objDAS.OpenRecordset(sSql)
   
   If Not adoRS Is Nothing Then
      If Not (adoRS.BOF And adoRS.EOF) Then
         adoRS.MoveFirst
         GetICD9Desc = adoRS("Description")
      End If
   End If
   
Error_Handler:
   'Clean Up Objects
   If Not adoRS Is Nothing Then
      If adoRS.State = adStateOpen Then adoRS.Close
      Set adoRS = Nothing
   End If
   Set objConn = Nothing
   Set objDAS = Nothing
   
   If CBool(Err) Then Call mobjshared.ShowVBError(Err.Number, Err.DESCRIPTION, "ListManagementOps.GetICD9Desc", "ListManagement", vbExclamation)

End Function
'JT END SCR 42399


