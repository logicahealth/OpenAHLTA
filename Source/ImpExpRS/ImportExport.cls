VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ImportExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'  Private Const Invalid_DMBE As Long = vbObjectError + 1
  Public FromPart As String                         ' Name of Table working on...retrieved from 1 object
  Public FromRs As Adodb.Recordset                  ' All Record sets from table being moved from
  Public ToRs As Adodb.Recordset                    ' Subset of records sets from Table Data being moved to
  Public ADDKEY As String                           ' Additional Field name used to sortby
  Public CreatedOn As String                        ' Field Name of Field indicating when record was created
  Public UpdatedOn As String                        ' Field Name of Field indicating when record was updated
  Public ActionToTake As String                     ' If To Table needs to be updated...what type of action
  Public KeyMap As Dictionary
  Public oDatabase As New CHCSII_CONN.Conn
  Public DataSQL As ISqlOps
  Private DataDas As ICHCSII_DAS
  Private TransferRecordSets As Collection
  Private NewSequentialFields() As Adodb.Field       ' Fields mapped to rs recordset for new sequential ID's
  Private NewSequentialFieldNames() As String        ' Name of above fields
  Private SourceOfSequentialIDs() As SequentialIDSource   ' Used to Determine Source of New Unique ID
  Private TypeOfSequentialIDs() As ICHCSII_SQL.eSeq  ' Type of Sequential ID...Used for everything but NCIDS
  Private DictionaryPrefix() As String               ' Synonym of Field Name...used for all tables in this group
  Private DictionarySynonym() As String              ' Synonym to test for...get value of inserted Key for new record
  Private WhereUpdateFields() As Adodb.Field         ' Required fields in where clause used in UPDATE
  Private WhereUpdateFieldNames() As String          ' Name of above fields
'  Private RSObject As IRSObject.I_RSObjects
   Private RSObject As IDMBE.I_RSObjects
'  Private DMBE As DMBE.DMBE_Ops                      ' Backend Object
  
    Private Sub Class_Initialize()
     Set DataSQL = oDatabase.CHCSII_SQLOPS(Auto) 'New GEMS_SQL.SqlOps
     Set DataDas = oDatabase.CHCSII_DAS(Auto)
     Set KeyMap = New Dictionary
     Set TransferRecordSets = New Collection
    End Sub
    
    
Private Sub Class_Terminate()
 Set KeyMap = Nothing
 Erase NewSequentialFields
 Erase NewSequentialFieldNames
 Erase SourceOfSequentialIDs
 Erase WhereUpdateFields
 Erase WhereUpdateFields
 Set FromRs = Nothing
 Set ToRs = Nothing
 Set DataSQL = Nothing
 Set RSObject = Nothing
 Set TransferRecordSets = Nothing
End Sub
Public Function InsertField(ByVal xField As Adodb.Field, sepchar As String, xFields As String, xValues As String)

   '   This routine takes a field and converts it to the proper format
   
   '     Date, numeric and string are handled differently
   
   '     Empty does nothing
   
   '     xfields = FieldName
   
   '     string       'Value'
   '     Date         'Todate(Value)'
   '     Numeric      Value
   '
   '     Sepcharacter is modified in this routine but saved in calling routine
   
   
    Dim xfieldstr As String
    Dim Value As String
    Dim quote As String
    Dim count As Integer
    Dim NEWID As Long
    Dim DictionaryKey As String
    Dim bSeqFields As Boolean
    
    If IsEmpty(xField.Value) Then Exit Function
  
    'Need to check to see if NewSequentialFieldNames is Defined
    On Error Resume Next
    Err.Clear
    count = LBound(NewSequentialFieldNames)
    If Err.Number = 0 Then
        bSeqFields = True
    End If
    
    If bSeqFields Then
  
        '...Get new ID's for the necessary fields
        For count = LBound(NewSequentialFieldNames) To UBound(NewSequentialFieldNames)
            If NewSequentialFieldNames(count) = xField.Name Then
                '...Does this table have foreign keys...
                '   If so is this one of the fields that should use the foreign key
                If DictionarySynonym(count) <> "" Then
                   '...get Old key from imported record
                   DictionaryKey = DictionarySynonym(count) & CStr(NewSequentialFields(count).Value)
                   '...Was this key replaced with a new key
                   If KeyMap.Exists(DictionaryKey) Then NEWID = CLng(KeyMap.Item(DictionaryKey))
                End If
                
                '...If newid = 0 need to create a new key
                If NEWID = 0 Then
                   '...Determine source of key
                   Select Case SourceOfSequentialIDs(count)
                      '...Many options depending on table
                      Case DataId
                        NEWID = DataSQL.GetNextID(TypeOfSequentialIDs(count))
                      '...Ncid no options
                      Case NCID
                        NEWID = DataSQL.GetNextNcid
                   End Select
                   '...This is a unique key...Set up dictionary to map to old key value
                   '...Item = New Key Value to use in the future
                   KeyMap.Add DictionaryPrefix(count) & CStr(NewSequentialFields(count).Value), CStr(NEWID)
                End If
            End If
        Next count
    End If
           
    If Err.Number <> 0 Then Err.Clear
    
  
    Select Case xField.Type
  
       Case adDate, adDBDate, adDBTime, adDBTimeStamp
          
          'SCR #44962- SML- attempting to format a null into a date
          If IsNull(xField.Value) Then
            Value = "NULL"
          Else
            Value = oDatabase.SQLDate(CDate(xField.Value))
          End If
          
       Case adBSTR, adChar, adLongVarChar, adLongVarWChar, adVarChar, adVarWChar, adWChar
          ' Begin:SCR #26430;   Developer: MAG 08/12/2002 02:22 PM
          If InStr(xField.Value, "'") > 0 Then
            Value = Replace(xField.Value, "'", "''")
          Else
            Value = xField.Value
          End If
          ' End: SCR #26430;
          quote = "'"
          
       Case Else         ' Convert Numeric to string
          Value = CStr(xField.Value)
          If Value = "" Then
            Value = "NULL"
          End If
    End Select
    
    If bSeqFields Then
        If Not IsEmpty(NewSequentialFieldNames) Then
            '--- If the field is one of the sequential fields, use the generated NEWID
            For count = LBound(NewSequentialFieldNames) To UBound(NewSequentialFieldNames)
                If NewSequentialFieldNames(count) = xField.Name Then
                    Value = CStr(NEWID)
                    Exit For
                End If
            Next count
        End If
    End If
  
    xFields = xFields & sepchar & " " & xField.Name
    xValues = xValues & sepchar & " " & quote & Value & quote
    sepchar = ","

End Function
Public Function IncludeField(ByVal xField As Adodb.Field, sepchar As String) As String

   '   This routine takes a field and converts it to the proper format
   
   '     Date, numeric and string are handled differently
   
   '     Empty returns a Null
   '     string      FieldName = 'Value'
   '     Date        FieldName = 'Todate()'
   '     Numeric     FieldName = Value
   '
   '     Sepcharacter is modified in this routine but saved in calling routine
   
   
  Dim xfieldstr As String
  Dim Value As String
  Dim quote As String
  
  If IsEmpty(xField.Value) Or IsNull(xField.Value) Then Exit Function
  
  Select Case xField.Type
  
       Case adDate, adDBDate, adDBTime, adDBTimeStamp
          Value = oDatabase.SQLDate(CDate(xField.Value))
          
       Case adBSTR, adChar, adLongVarChar, adLongVarWChar, adVarChar, adVarWChar, adWChar
          ' Begin:SCR #26430;   Developer: MAG 08/12/2002 02:21 PM
          If InStr(xField.Value, "'") > 0 Then
            Value = Replace(xField.Value, "'", "''")
          Else
            Value = xField.Value
          End If
          ' End: SCR #26430;
          quote = "'"
          
       Case Else         ' Convert Numeric to string
          Value = CStr(xField.Value)
  End Select
  
  IncludeField = sepchar & " " & xField.Name & " = " & quote & Value & quote
  sepchar = ","

End Function
Public Function MoveRecordToGems(rs As Adodb.Recordset) As Boolean

  On Local Error GoTo haveError

 '  This routine builds an SQL statement _
      For Insert and Update
 
 '  It also Translates any keys that need to be _
    translated to new values or old values _
    depending on the situation.
 
 '  Old values of existing records are used _
    New Values are created for imported records _
    that will be added to the database and their _
    old values are put in a dictionary for future _
    mapping in other tables (foreing keys)

  Dim SqlPrefix As String
  Dim xFieldTo As Variant
  Dim xfieldFrom As Variant
  Dim WhereFields As String
  Dim SetFields As String
  Dim wsepchar As String
  Dim ssepchar As String
  Dim AddFields As String
  Dim AddValues As String
  Dim sql As String
  Dim DirtyData As Boolean

                                     
    Select Case UCase$(Trim$(ActionToTake))
        Case "INSERT"
           
          SqlPrefix = "INSERT INTO " & FromPart & " ("
           
           On Error Resume Next
           
           '...Build new Field and value part of record to add
           'MAG 01/22/2003: ALLERGYTYPETEXT is a alias for ALLERGIN in ALLERGIES Table
           For Each xfieldFrom In rs.Fields
                If UCase(CStr(xfieldFrom.Name)) <> "CREATEDON" And _
                   UCase(CStr(xfieldFrom.Name)) <> "CREATEDBY" And _
                   UCase(CStr(xfieldFrom.Name)) <> "ALLERGYTYPETEXT" Then
                     ' Begin:SCR #45193;   Developer:  11/18/2003
                     If Not FindFieldInRS(ToRs, xfieldFrom.Name) Is Nothing Then
                        InsertField xfieldFrom, ssepchar, AddFields, AddValues
                     End If
                     ' End: SCR #45193;
                End If
           Next xfieldFrom
           
           sql = SqlPrefix & AddFields & ") Values (" & AddValues & ")"
           
        Case "UPDATE"
          
          SqlPrefix = "UPDATE " & FromPart & " SET "
          
          For Each xFieldTo In ToRs.Fields
            '...Build where clause
            If WhereFields = "" Then
                WhereFields = WhereFields & IncludeField(xFieldTo, "")
            Else
                WhereFields = WhereFields & IncludeField(xFieldTo, " And")
            End If
            
            'SCR #44962- SML- need to use "AND" in updates, not ","
            If wsepchar = "," Then
                wsepchar = " AND "
            End If
        
            For Each xfieldFrom In rs.Fields
               '...find proper field and build set Clause
               If UCase$(xfieldFrom.Name) = UCase$(xFieldTo.Name) Then
                 '...are the values different
                 If xfieldFrom.Value <> xFieldTo.Value Then
                  '...Yes update statement and set dirtydata
                  SetFields = SetFields & IncludeField(xfieldFrom, ssepchar)
                  DirtyData = True
                 End If
                 Exit For
               End If
            Next xfieldFrom
          Next xFieldTo
          
          '...If dirtydata is false no need to update
          If DirtyData Then sql = SqlPrefix & SetFields & " Where " & WhereFields
          
        Case Else          ' Invalid input
        
    End Select
    
    '...Process the statement
    If sql <> "" Then
        If DataSQL.Execute(sql) <> 0 Then
            MoveRecordToGems = True
        Else
            MoveRecordToGems = False
            GoTo haveError
        End If
    End If
ExitRoutine:
    Exit Function
    
haveError:
    '...When done testing raise an error here
    Debug.Print Err.Number, Err.Description
    Resume ExitRoutine
    Resume
End Function



Public Function EraseArrays()

      Erase NewSequentialFields
      Erase NewSequentialFieldNames
      Erase SourceOfSequentialIDs
      Erase WhereUpdateFields
      Erase WhereUpdateFieldNames
      Erase DictionaryPrefix
      Erase DictionarySynonym

End Function

Public Sub SetWhereCondition(FieldNames() As Variant)
 Dim newsize As Integer
 Dim count As Integer
 Dim xbase As Integer
 
   newsize = UBound(FieldNames)
   xbase = LBound(FieldNames)
   
    ReDim WhereUpdateFieldNames(xbase To newsize)
    ReDim WhereUpdateFields(xbase To newsize)
    
    For count = LBound(FieldNames) To newsize
     
       WhereUpdateFieldNames(count) = CStr(FieldNames(count))
      
    Next count

End Sub
Public Sub SetNewSequential(FieldNames() As Variant, SeqSource() As Variant, SeqType() As Variant, _
                            Prefix() As Variant, Synonym() As Variant)
 
  ' This routine redefines critical Sequential ID Fields
  ' These fields are Key fields that need to be modified
  ' with a new key when inserted or updating a record
  ' Name of Field in record and Source of ID for that field is the
  ' input to this routine
 
 Dim newsize As Integer
 Dim count As Integer
 Dim xbase As Integer
 
   newsize = UBound(FieldNames)
   xbase = LBound(FieldNames)
   
    ReDim NewSequentialFieldNames(xbase To newsize)
    ReDim NewSequentialFields(xbase To newsize)
    ReDim SourceOfSequentialIDs(xbase To newsize)
    ReDim TypeOfSequentialIDs(xbase To newsize)
    ReDim DictionaryPrefix(xbase To newsize)                ' Synonym of Field Name...used for all tables in this group
    ReDim DictionarySynonym(xbase To newsize)               ' Synonym to test for...get value of inserted Key for new record
    
    For count = xbase To newsize
     
       NewSequentialFieldNames(count) = CStr(FieldNames(count))
       SourceOfSequentialIDs(count) = CLng(SeqSource(count))
       TypeOfSequentialIDs(count) = CLng(SeqType(count))
       DictionaryPrefix(count) = CStr(Prefix(count))
       DictionarySynonym(count) = CStr(Synonym(count))
      
    Next count

End Sub


Public Sub MapFields()
 On Error GoTo haveError
 Dim WhereAreWe As Integer
 
 ' This routine will map Fields in an array to appropriate
 ' Fields in a recordset

 Dim count As Integer
     WhereAreWe = 1
      '...Loop through all critical fields and map to proper Field in imported record set
      For count = LBound(NewSequentialFields) To UBound(NewSequentialFields)
       Set NewSequentialFields(count) = FindFieldInRS(FromRs, NewSequentialFieldNames(count))
      Next count
nextloop:
      WhereAreWe = 2
      '...Loop through all critical fields and map to proper Field in To record set
      For count = LBound(WhereUpdateFields) To UBound(WhereUpdateFields)
       Set WhereUpdateFields(count) = FindFieldInRS(ToRs, WhereUpdateFieldNames(count))
      Next count
      
ExitRoutine:
   Exit Sub
haveError:
   If WhereAreWe = 1 Then
     Resume nextloop
   Else
     Resume ExitRoutine
   End If
End Sub
Public Function FindFieldInRS(rs As Adodb.Recordset, ByVal FieldName As String) As Field

    ' Find a specific Field in a record set
    ' If Field not found a Null Field will be returned
    ' else a Field object for the field name will be returned

 Dim LookFor As String
 Dim xField As Variant
 
 LookFor = UCase$(FieldName)
 
 For Each xField In rs.Fields
   If UCase$(xField.Name) = LookFor Then
    Set FindFieldInRS = xField
    Exit For
   End If
 Next xField
 
End Function


Public Sub ConvertDataToField(xField As Adodb.Field, ByVal Value As Variant)

         Select Case xField.Type
            Case adNumeric
                              If Value <> "" Then
                                 xField.Value = CLng(Value)
                              End If
            Case adVarChar
                              If Value <> "" Then
                                 xField.Value = CStr(Value)
                              End If
            Case adVarWChar
                              If Value <> "" Then
                                 xField.Value = CStr(Value)
                              End If
            Case adBoolean
                              If Value <> "" Then
                                 xField.Value = CBool(Value)
                              End If
            Case adDouble
                              If Value <> "" Then
                                 xField.Value = CDbl(Value)
                              End If
            Case adInteger
                              If Value <> "" Then
                                 xField.Value = CLng(Value)
                              End If
            Case adSmallInt
                              If Value <> "" Then
                                 xField.Value = CInt(Value)
                              End If
            Case adDBTimeStamp
                              If Value <> "" Then
                                 xField.Value = CDate(Value)
                              End If
            Case adLongVarWChar         '203
                              If Value <> "" Then
                                 xField.Value = CStr(Value)
                              End If
            Case adUnsignedTinyInt     ' char
                              If Value <> "" Then
                                 xField.Value = CStr(Value)
                              End If
            Case Else
                              Call MsgBox("Data type = " & xField.Type & " for " & Value, vbOKOnly, "Missing Data Type")
         End Select

End Sub

Public Function SaveGEMSrs(rs As Adodb.Recordset, ByVal unitnumber As String, ByVal TableName As String) As Boolean
 
    If rs Is Nothing Or unitnumber = "" Or TableName = "" Then Exit Function
    
    SaveGEMSrs = SaveRS(rs, unitnumber, TableName, "...LDDB\")
 
    If UCase(TableName) = "ENCOUNTERS" Then
        Call SaveProviderInfo(rs, unitnumber, TableName)
    ElseIf UCase(TableName) = "INPATIENT_ADMISSION" Then
        Call SaveProviderInfo(rs, unitnumber, TableName)
    End If
 
End Function
Public Function SaveCDRrs(rs As Adodb.Recordset, ByVal unitnumber As String, ByVal TableName As String) As Boolean
 
 If rs Is Nothing Or unitnumber = "" Or TableName = "" Then Exit Function
 SaveCDRrs = SaveRS(rs, unitnumber, TableName, "...CDR\")

End Function
Public Function SaveEXTrs(rs As Adodb.Recordset, ByVal unitnumber As String, ByVal TableName As String) As Boolean
 
 If rs Is Nothing Or unitnumber = "" Or TableName = "" Then Exit Function
 SaveEXTrs = SaveRS(rs, unitnumber, TableName, "...EXT\")

End Function


Private Function SaveProviderInfo(oRS As Adodb.Recordset, ByVal sUnitNumber As String, ByVal sTableName As String) As Boolean
Dim sProvNCIDs As String
Dim asProvColumns() As String
Dim oProvRs As Adodb.Recordset
Dim sColumn As String
Dim sSQL As String
Dim idx As Integer
Dim RSObject As DMBE.RSObject_Ops
Dim sProviderTableName As String
Dim iCounter As Long

    On Error GoTo ErrHnd
    
  If oRS Is Nothing Then Exit Function
    
  sProviderTableName = "PROVIDER"
  
  asProvColumns = Split(GetProviderColumnNames(sTableName), ",")
  Do While Not oRS.EOF
    For idx = 0 To UBound(asProvColumns)
      sColumn = Trim$(asProvColumns(idx))
      If oRS(sColumn) & vbNullString <> "" Then
        If InStr(1, sProvNCIDs, oRS(sColumn) & ",") = 0 Then
          'Provider Not in String so Add him
          sProvNCIDs = sProvNCIDs & oRS(sColumn) & ","
        End If
      End If
    Next idx
    oRS.MoveNext
  Loop
  
  ' SCR #50970;   Developer: MAG 03/11/2004
  If Len(sProvNCIDs) > 1 Then
    
    'SCR 71889
    'If the table name is "INPATIENT_ADMISSION", then the "ENCOUNTERS" providers have already been saved, ADD the inpatient_admission providers to the list
    If UCase(sTableName) = "INPATIENT_ADMISSION" Or UCase(sTableName) = "ENCOUNTERS" Then
        For iCounter = 1 To TransferRecordSets.count
            Set RSObject = TransferRecordSets(iCounter)
            If UCase(RSObject.TableName) = sProviderTableName Then
                RSObject.DBRECORDSET.MoveFirst
                Do Until RSObject.DBRECORDSET.EOF
                    sProvNCIDs = sProvNCIDs & RSObject.DBRECORDSET.Fields("NCID").Value & ","
                    RSObject.DBRECORDSET.MoveNext
                Loop
                
                Set RSObject = Nothing
                TransferRecordSets.Remove iCounter
                Exit For
            End If
        Next
    End If
    
    sProvNCIDs = Left(sProvNCIDs, Len(sProvNCIDs) - 1)
  
    sSQL = "Select * From PROVIDER Where NCID In (" & sProvNCIDs & ")"
    Set oProvRs = DataDas.OpenRecordset(sSQL)
  
    SaveProviderInfo = SaveRS(oProvRs, sUnitNumber, sProviderTableName, "...LDDB\")
  End If
  
Exit Function

ErrHnd:
    MsgBox "Application error: " & Err.Description & " occurred in ImpExpRS.ImportExport.cls.SaveProviderInfo"
    
End Function

Private Function GetProviderColumnNames(sTableName As String) As String

  Select Case UCase(sTableName)
    Case "ENCOUNTERS"
        GetProviderColumnNames = "PRIMARYPROVIDERNCID, ADDITIONALPROVIDER1NCID, ADDITIONALPROVIDER2NCID, WHOCOSIGNEDNCID, WHOSIGNEDNCID, COSIGNERNCID"
    
    Case "INPATIENT_ADMISSION"
        GetProviderColumnNames = "ADMITTING_PROVIDER_NCID, ATTENDING_PROVIDER_NCID"
        
  End Select
  
End Function

Public Function RecordSetsTransfered() As Long
 RecordSetsTransfered = TransferRecordSets.count
End Function

Public Property Get GetCDRrs(ByVal unitnumber As String, ByVal TableName As String) As Adodb.Recordset
 On Error Resume Next
 Set GetCDRrs = TransferRecordSets(unitnumber & "...CDR\" & TableName)
End Property


Public Sub ClearRSCollection()
  Dim count As Long
    
    With TransferRecordSets
      Do While .count <> 0
        .Remove 1
      Loop
    End With
    
End Sub

Friend Function SaveRS(rs As Adodb.Recordset, ByVal unitnumber As String, ByVal TableName As String, ByVal xType As String) As Boolean
   
     On Error GoTo ErrHnd
   
     '...Build Record Data Object used by Data Manager Back End
'     Set RSObject = New RSObject_Ops      '<---Modify This Line
     Set RSObject = New DMBE.RSObject_Ops     '<---Modify This Line
     
     '...Make sure a connection doesn't exist
     Set rs.ActiveConnection = Nothing
     With RSObject
        .TableName = TableName
        Set .DBRECORDSET = rs
     End With
     
     '...Add to collection NOTE the key.  This quarantees uniqueness and _
         gives future capability for multiple patients being downloaded _
         at once.
         
     TransferRecordSets.Add RSObject, unitnumber & xType & TableName
    
     SaveRS = True
    
Exit Function

ErrHnd:
    
    MsgBox "Application error: " & Err.Description & " occurred in ImpExpRS.ImportExport.cls.SaveRS"
    
End Function

Public Function ExecuteDMBE(DMBEObj As IDMBE.I_DMBE, Optional ByRef TMIPRetrievalResult As TMIPDataRetrievalResult) As Long
  
  '<SCR 22576 CC>
  If IsMissing(TMIPRetrievalResult) Then
    ExecuteDMBE = DMBEObj.Execute(TransferRecordSets)   '<---Modify This Line ' and remove parm
  Else
    ExecuteDMBE = DMBEObj.Execute(TransferRecordSets, TMIPRetrievalResult)
  End If
  '</SCR 22576 CC>
  
End Function

Public Function GetATable(ByVal TableName As String) As Adodb.Recordset

  Dim FileNameRecordSet As Variant               ' 1 object in TransferRecordSets collection
  Dim TableFound As Boolean

      For Each FileNameRecordSet In TransferRecordSets
       If LCase$(Trim$(TableName)) = LCase$(Trim$(FileNameRecordSet.TableName)) Then
        TableFound = True
        Exit For
       End If
      Next FileNameRecordSet
      
      '...Load imported Records
      If TableFound Then
        Set GetATable = FileNameRecordSet.DBRECORDSET
      Else
       On Local Error Resume Next
       Set GetATable = TransferRecordSets.Item(TableName).DBRECORDSET
       If Err.Number <> 0 Then
         Err.Clear
         Set GetATable = Nothing
       End If
      End If
        
End Function


'Public Property Set SetDMBE(xDMBE As DMBE.DMBE_Ops)
' Set DMBE = xDMBE
'End Property
