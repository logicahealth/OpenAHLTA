VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colPositions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsPosition"
Attribute VB_Ext_KEY = "Member0" ,"clsPosition"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Const TOTAL_POSITIONS = 32

Private meDirtyFlag As DirtyFlagEnum

Private mCol As Collection

Private mlPatientID As Long

Public Function ConvertNumToPos(ToothNum As Integer) As Integer
    
    ConvertNumToPos = ConvNumToPos(ToothNum)

End Function

Public Function ConvertPosToPrimaryNum(ToothPos As Integer) As Integer

    ConvertPosToPrimaryNum = ConvPosToPrimaryNum(ToothPos)
    
End Function
Private Function Add(Position As Integer, Status As ePositionStatus, Seen As Boolean) As clsPosition
    
    'create a new object
    Dim objNewMember As clsPosition
    Set objNewMember = New clsPosition
    
    With objNewMember
    
        'set the properties passed into the method
        .Position = Position
        .Status = Status
        .Seen = Seen
        'This must be set after seen is set.
        .ToothErupted = False

    End With
    
    mCol.Add objNewMember, CStr(Position)

    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Function GetPositionByToothNum(ToothNum As Integer) As clsPosition

'This gets a position given a tooth number
Dim iPos As Integer
    
    'For some strange reason, this gets passed a 0 for the toothnum
    If ToothNum = 0 Then
        Set GetPositionByToothNum = New clsPosition
        GetPositionByToothNum.Position = 0
        Exit Function
    End If
    
    iPos = ConvNumToPos(ToothNum)
    
    Set GetPositionByToothNum = mCol(CStr(iPos))
    
End Function

Public Property Get Item(vntIndexKey As Variant) As clsPosition
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


Private Sub LoadFromTeethList(TeethList As String)

'This takes a comma delimited list of teeth and sets the default for each
'"3,8,9,14,19,23,24,25,26,30,33,34,35,36,39,40,41,42,43,44,45,50,51,52"

Dim arTeeth() As String
Dim sToothNum As Variant
Dim oPos As clsPosition

    ResetPositions
    
    arTeeth = Split(TeethList, ",")
    
    For Each sToothNum In arTeeth
        Set oPos = mCol(CStr(ConvNumToPos(sToothNum)))
        oPos.Seen = False
        oPos.ToothErupted = True
        
        If sToothNum > 32 Then
            oPos.Status = ePositionStatus.ePrimary
        Else
            oPos.Status = ePositionStatus.eSecondary
        End If
    Next

End Sub

Private Function LoadFromXML(PatientID As Long) As Boolean

'This function reads the xml data from the DENTAL_ENCOUNTER_MOUTH_POS table and loads the position collection

Dim sSQL As String
Dim oCHCSIIConn As CHCSII_CONN.Conn
Dim oConn As ICHCSII_DAS
Dim oDAS As ICHCSII_DAS
Dim oRS As Recordset
Dim oXML As ChilkatXml
Dim oNode As ChilkatXml
Dim oPos As clsPosition

On Error GoTo ErrHnd

    Set oCHCSIIConn = New CHCSII_CONN.Conn
    Set oDAS = oCHCSIIConn.CHCSII_DAS
    
    '** This only works for oracle. Determine if the new table exists. If it doesn't then don't use it.
    If oCHCSIIConn.UseSQLServer = False Then
        sSQL = "select table_name from all_tables where table_name = 'DENTAL_ENCOUNTER_MOUTH_POS'"
        Set oRS = oDAS.OpenRecordset(sSQL)
        If oRS.EOF Then
            LoadFromXML = False
            Exit Function
        End If
    End If
    
    sSQL = "SELECT PATIENT_UNIT_ID, POSITION_INFO FROM DENTAL_ENCOUNTER_MOUTH_POS WHERE PATIENT_UNIT_ID = " & PatientID
    
    Set oRS = oDAS.OpenRecordset(sSQL)
    
    ResetPositions
    
    Set oXML = New ChilkatXml
    
    If Not oRS.EOF Then
        oXML.LoadXml oRS("Position_Info")
        
        Set oNode = oXML.FirstChild
        
        Do Until oNode Is Nothing
            Set oPos = mCol(CStr(ConvNumToPos(oNode.GetAttrValue("TN"))))
            oPos.Seen = CBool(oNode.GetAttrValue("Seen"))
            oPos.ToothErupted = CBool(oNode.GetAttrValue("TE"))
            oPos.Status = Val(oNode.GetAttrValue("Status"))
            oPos.PropertyChanged = False
            
            Set oNode = oNode.NextSibling
        Loop
        
        LoadFromXML = True
    End If
    
Exit Function

ErrHnd:
    MsgBox "Could not load the tooth positions from the DENTAL_ENCOUNTER_MOUTH_POS table. " & Err.Description
    Exit Function
    Resume
End Function

Public Function LoadPositions(PatientID As Long, AgeTeethPath As String) As Boolean

'This function loads the positions based on where the data may be.
'It Looks first in the Dental_Encounter_Mouth_Pos table.
'If the data isn't there, it looks in the Dental_Enc_View_Teeth table.
'If the data isn't there, then it loads the info from the AgeTeethPath - The default tooth state based on the patient's age.

On Error GoTo ErrHnd

    If LoadFromXML(PatientID) = True Then
        meDirtyFlag = dfUpdate
    Else
        'Could not load from the new table, so try loading from the old table
        If LoadFromOldTable(PatientID) = True Then
            meDirtyFlag = dfAdd
        Else
            'Could not load from the old table either, so set the teeth to their defaults depending on the patient's age.
            LoadFromTeethList AgeTeethPath
            meDirtyFlag = dfAdd
        End If
    End If
    
    mlPatientID = PatientID
    
Exit Function

ErrHnd:
    MsgBox "Could not load the tooth positions.", vbCritical
    
End Function


Private Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)

    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub ResetPositions()

    Dim iCounter As Integer
    
    'creates the collection when this class is created
    Set mCol = New Collection
    
    'Init the positions
    For iCounter = 1 To TOTAL_POSITIONS
        'Default all the teeth to the earliest tooth
        Select Case iCounter
            'These teeth can only be secondary
            Case 1, 2, 3, 14, 15, 16, 32, 31, 30, 19, 18, 17
                Add iCounter, ePositionStatus.eSecondary, False
            
            'These teeth can be primary or secondary
            Case 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29
                Add iCounter, ePositionStatus.ePrimary, False
        
        End Select
    Next

End Sub

Public Function Save() As Boolean

'This saves the Position information in the Dental_Encounter_Mouth_Pos table

On Error GoTo ErrHnd

Dim oCHCSIIConn As CHCSII_CONN.Conn
Dim oDAS As ICHCSII_DAS
Dim oRS As Recordset
Dim oNode As ChilkatXml
Dim oNodeChild As ChilkatXml
Dim oPos As clsPosition
Dim sSQL As String
Dim bUpdateDB As Boolean

    Set oCHCSIIConn = New Conn
    Set oDAS = oCHCSIIConn.CHCSII_DAS(Auto, NoTimeZoneConversion)
    
    Set oNode = New ChilkatXml
    
    oNode.Tag = "Positions"
    
    For Each oPos In mCol
        Set oNodeChild = oNode.NewChild("Pos" & oPos.Position, "")
        With oNodeChild
            .AddAttribute "TN", oPos.ToothNum
            .AddAttribute "Seen", oPos.Seen
            .AddAttribute "Status", oPos.Status
            .AddAttribute "TE", oPos.ToothErupted
        End With
        If oPos.PropertyChanged = True Then
            bUpdateDB = True
            oPos.PropertyChanged = False
        End If
    Next
    
    If bUpdateDB Then
        
        '** This only works for oracle. Determine if the new table exists. If it doesn't then don't save.
        If oCHCSIIConn.UseSQLServer = False Then
            sSQL = "select table_name from all_tables where table_name = 'DENTAL_ENCOUNTER_MOUTH_POS'"
            Set oRS = oDAS.OpenRecordset(sSQL)
            If oRS.EOF Then
                MsgBox "The table DENTAL_ENCOUNTER_MOUTH_POS does not exist in the database. The tooth position information will not be saved.", vbCritical
                Save = False
                Exit Function
            End If
        End If
        
        Select Case meDirtyFlag
            Case DirtyFlagEnum.dfNoAction
                Exit Function
    
            Case DirtyFlagEnum.dfAdd
                sSQL = "INSERT INTO DENTAL_ENCOUNTER_MOUTH_POS (PATIENT_UNIT_ID, POSITION_INFO)" _
                & " VALUES (" & mlPatientID & ", '" & oNode.GetXml & "')"
    
            Case DirtyFlagEnum.dfUpdate
                sSQL = "UPDATE DENTAL_ENCOUNTER_MOUTH_POS SET" _
                    & " POSITION_INFO = '" & oNode.GetXml & "'" _
                    & " WHERE PATIENT_UNIT_ID = " & mlPatientID
        End Select
        
        oDAS.ExecuteSQL sSQL
        
        If meDirtyFlag = dfAdd Then
            meDirtyFlag = dfUpdate
        End If
    End If
    
    Save = True

Exit Function

ErrHnd:
    MsgBox "Could not save the teeth information. Error in colPositions.Save: " & Err.Description, vbCritical
    Save = False
    
End Function

Private Function LoadFromOldTable(lPatientID As Long) As Boolean

    Dim oRS As ADODB.Recordset
    Dim sSQL As String
    Dim oCHCSIIConn As CHCSII_CONN.Conn
    Dim oDAS As ICHCSII_DAS
    Dim oPos As clsPosition
    
    On Error GoTo ErrorHandler
    
    Set oCHCSIIConn = New CHCSII_CONN.Conn
    Set oDAS = oCHCSIIConn.CHCSII_DAS
    
    '** This only works for oracle. Determine if the old table exists. If it doesn't then don't use it.
    If oCHCSIIConn.UseSQLServer = False Then
        sSQL = "select table_name from all_tables where table_name = 'DENTAL_ENC_VIEW_TEETH'"
        Set oRS = oDAS.OpenRecordset(sSQL)
        If oRS.EOF Then
            LoadFromOldTable = False
            Exit Function
        End If
    End If
    
    ' get SelectedTeeth info for this patient id from db
    sSQL = "SELECT TOOTH_ID, TOOTH_STATUS, TOOTH_SEEN FROM DENTAL_ENC_VIEW_TEETH"
    sSQL = sSQL & " WHERE PATIENT_UNIT_ID = " & lPatientID
           
    Set oRS = oDAS.OpenRecordset(sSQL)
    If oRS.RecordCount <> mCol.Count And oRS.RecordCount <> 0 Then
        'Somehow we don't have all 32 teeth, so delete it and start over 'NO NO NO
'        sSQL = "DELETE FROM DENTAL_ENC_VIEW_TEETH"
'        sSQL = sSQL & " WHERE PATIENT_UNIT_ID = " & lPatientID
'        objDas.ExecuteSQL (sSQL)
'
'        ReadTeethPathFromDB = ""
'        mbNewTeethPath = True
'        Exit Function
        MsgBox Abs(CLng(mCol.Count) - CLng(oRS.RecordCount)) & " tooth (teeth) position(s) were not found in the database. Please reset the tooth state information (primary/secondary) on the grey tooth (teeth) in full mouth.", vbOKOnly + vbInformation
    ElseIf oRS.RecordCount = 0 Then
    
        LoadFromOldTable = False
        'mbNewTeethPath = True
        Exit Function
    End If
    
    Do While Not oRS.EOF
        Set oPos = mCol(CStr(oRS.Fields("TOOTH_ID"))) 'This is really the position number, NOT the tooth id
        With oPos
            'Set seen before Status when loading from the db becuase the db doesn't reflect the real state of the data
            .Seen = CBool(oRS.Fields("TOOTH_SEEN"))
            .Status = CInt(oRS.Fields("TOOTH_STATUS"))
            'If the tooth status is not 0, then it has erupted
            If CInt(oRS.Fields("TOOTH_STATUS")) <> 0 Then
                .ToothErupted = True
            End If
            .PropertyChanged = False
        End With
        
        oRS.MoveNext
    Loop
    
    Set oRS = Nothing
    Set oDAS = Nothing
    Set oCHCSIIConn = Nothing
    
    LoadFromOldTable = True
    
    'ReadTeethPathFromDB = moPositions.TeethPath
    'mbNewTeethPath = False
    Exit Function
    
ErrorHandler:
    MsgBox Err.Number & ": " & Err.Description & ": clsTeeth.readTeethPathFromDB"
    Exit Function
    Resume
End Function


Public Function TeethPath() As String

    'Get the teeth path - for backwards compatibility
    
    Dim oPos As clsPosition
    
    For Each oPos In mCol
        'If oPos.Status <> ePositionStatus.eNone Then
        If oPos.ToothErupted = True Then
            TeethPath = TeethPath & "," & oPos.ToothNum
        End If
    Next
    
    'Remove the first ","
    TeethPath = Mid(TeethPath, 2)
    
End Function

Private Sub Class_Initialize()

    ResetPositions
    
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

