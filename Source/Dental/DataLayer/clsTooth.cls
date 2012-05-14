VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTooth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
'PK
Private mvarDentalEncounterTeethID As Long

Private mvarDentalEncounterID As Long
Private mvarDRC As Long
Private mvarToothID As Long
Private mvarToothName As String
Private mvarMesialVersion As Long
Private mvarDistalVersion As Long
Private mvarFacialVersion As Long
Private mvarLingualVersion As Long
Private mvarRotatedMesial As Long
Private mvarRotatedDistal As Long
Private mvarIntruded As Long
Private mvarExtruded As Long
Private mvarDentalEncOrientationID As Long
Private mvarDirtyFlag As DirtyFlagEnum
Private mvarDeleted As Boolean
Friend Property Let DirtyFlag(ByVal vData As DirtyFlagEnum)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.DirtyFlag = Form1
    mvarDirtyFlag = vData
End Property


Friend Property Get DirtyFlag() As DirtyFlagEnum
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DirtyFlag
    DirtyFlag = mvarDirtyFlag
End Property



Friend Property Let DentalEncounterTeethID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DentalEncounterTeethID = 5
    mvarDentalEncounterTeethID = vData
    SetDirtyFlag

End Property


Public Property Get DentalEncounterTeethID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DentalEncounterTeethID
    DentalEncounterTeethID = mvarDentalEncounterTeethID
End Property





Public Property Let Deleted(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DentalEncOrientationID = 5
    mvarDeleted = vData
    SetDirtyFlag
    
End Property




Public Property Let Extruded(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Extruded = 5
    mvarExtruded = vData
    SetDirtyFlag

End Property


Public Property Get Extruded() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Extruded
    Extruded = mvarExtruded
End Property



Public Property Let Intruded(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Intruded = 5
    mvarIntruded = vData
    SetDirtyFlag

End Property


Public Property Get Intruded() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Intruded
    Intruded = mvarIntruded
End Property



Public Property Let RotatedDistal(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RotatedDistal = 5
    mvarRotatedDistal = vData
    SetDirtyFlag

End Property


Public Property Get RotatedDistal() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RotatedDistal
    RotatedDistal = mvarRotatedDistal
End Property



Public Property Let RotatedMesial(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RotatedMesial = 5
    mvarRotatedMesial = vData
    SetDirtyFlag

End Property


Public Property Get RotatedMesial() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RotatedMesial
    RotatedMesial = mvarRotatedMesial
End Property



Public Property Let LingualVersion(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LingualVersion = 5
    mvarLingualVersion = vData
    SetDirtyFlag

End Property


Public Property Get LingualVersion() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LingualVersion
    LingualVersion = mvarLingualVersion
End Property



Public Property Let FacialVersion(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FacialVersion = 5
    mvarFacialVersion = vData
    SetDirtyFlag

End Property


Public Property Get FacialVersion() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FacialVersion
    FacialVersion = mvarFacialVersion
End Property



Public Property Let DistalVersion(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DistalVersion = 5
    mvarDistalVersion = vData
    SetDirtyFlag

End Property


Public Property Get DistalVersion() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DistalVersion
    DistalVersion = mvarDistalVersion
End Property



Public Property Let MesialVersion(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MesialVersion = 5
    mvarMesialVersion = vData
    SetDirtyFlag

End Property


Public Property Get MesialVersion() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MesialVersion
    MesialVersion = mvarMesialVersion
End Property





Friend Function SaveTooth(ByRef DentalEncounterTeethID As Long) As Boolean

'This function saves the tooth in the
'Dental_Encounter_Teeth table.

Dim sSQL As String
Dim rs As Recordset

On Error GoTo SaveToothErr

    SaveTooth = False

    Select Case mvarDirtyFlag
        Case DirtyFlagEnum.dfNoAction
            'No changes have been made since the last save, so do nothing
   
        Case DirtyFlagEnum.dfUpdate
            'Update the table
            sSQL = "UPDATE DENTAL_ENCOUNTER_TEETH" _
                & " SET TOOTH_ID = " & mvarToothID _
                & ", DRC = " & mvarDRC _
                & ", TOOTH_NAME = '" & mvarToothName & "'" _
                & ", MESIAL_VERSION = '" & mvarMesialVersion & "'" _
                & ", DISTAL_VERSION = '" & mvarDistalVersion & "'" _
                & ", FACIAL_VERSION = '" & mvarFacialVersion & "'" _
                & ", LINGUAL_VERSION = '" & mvarLingualVersion & "'" _
                & ", ROTATED_MESIAL = '" & mvarRotatedMesial & "'" _
                & ", ROTATED_DISTAL = '" & mvarRotatedDistal & "'" _
                & ", INTRUDED = '" & mvarIntruded & "'" _
                & ", EXTRUDED = '" & mvarExtruded & "'" _
                & ", DELETED = " & CInt(mvarDeleted) _
                & " Where (Dental_Encounter_Teeth_ID = " & mvarDentalEncounterTeethID & ")"
            goconn.ExecuteSQL sSQL
            
        Case DirtyFlagEnum.dfAdd
            'Add the record to the table
            
            If gbUseStoredProc = True Then
                    
                Dim oCmd As CHCSII_Command.CHCSII_Cmd
                Dim oParam As CHCSII_Command.CHCSII_SPparam
                
                Set oCmd = New CHCSII_Command.CHCSII_Cmd
                oCmd.SPName = "DL_NOTE_TOOTH_INSERT"
                
                'Call oCmd.AddParam(False, False, "textString", mvarDentalEncounterTeethID)
                Call oCmd.AddParam(False, False, "textString", mvarToothID)
                Call oCmd.AddParam(False, False, "textString", mvarDRC)
                Call oCmd.AddParam(False, False, "textString", mvarToothName)
                Call oCmd.AddParam(False, False, "textString", mvarMesialVersion)
                Call oCmd.AddParam(False, False, "textString", mvarDistalVersion)
                Call oCmd.AddParam(False, False, "textString", mvarFacialVersion)
                Call oCmd.AddParam(False, False, "textString", mvarLingualVersion)
                Call oCmd.AddParam(False, False, "textString", mvarRotatedMesial)
                Call oCmd.AddParam(False, False, "textString", mvarRotatedDistal)
                Call oCmd.AddParam(False, False, "textString", mvarIntruded)
                Call oCmd.AddParam(False, False, "textString", mvarExtruded)
                Call oCmd.AddParam(False, False, "textString", CInt(mvarDeleted))
                Call oCmd.AddParam(False, False, "textString", mvarDentalEncounterID)
                Call oCmd.AddParam(True, True, "", "")
                
                Set rs = goconn.OpenRecordsetSP(oCmd)
                
                If Not rs.EOF Then
                    mvarDentalEncounterTeethID = rs("ID")
                Else
                    MsgBxARMd "Fatal Error: Could not get the sequence from the DL_NOTE_TOOTH_INSERT SP"
                    Exit Function
                End If

            Else
'                'Get the sequence for the table
'                sSQL = "Select DENTAL_ENCOUNTER_TEETH_SEQ.NEXTVAL as PKID from DUAL"
'                Set rs = goconn.OpenRecordset(sSQL)
'
'                'If I don't get the nextval from oracle, then exit
'                If rs.EOF Then
'                    SaveTooth = False
'                    Exit Function
'                End If
'                mvarDentalEncounterTeethID = rs("PKID")
'                rs.Close
'                Set rs = Nothing
                mvarDentalEncounterTeethID = goconn.GetNextID(esDENTAL_ENCOUNTER_TEETH_SEQ)
                
                'Add the record to the table
                sSQL = "INSERT INTO DENTAL_ENCOUNTER_TEETH" _
                    & " (DENTAL_ENCOUNTER_TEETH_ID," _
                    & " TOOTH_ID," _
                    & " DRC," _
                    & " TOOTH_NAME," _
                    & " MESIAL_VERSION," _
                    & " DISTAL_VERSION," _
                    & " FACIAL_VERSION," _
                    & " LINGUAL_VERSION," _
                    & " ROTATED_MESIAL," _
                    & " ROTATED_DISTAL," _
                    & " INTRUDED," _
                    & " EXTRUDED," _
                    & " DELETED, " _
                    & " DENTAL_ENCOUNTER_ID)"
                sSQL = sSQL & " VALUES" _
                    & " (" & mvarDentalEncounterTeethID & "," _
                    & " " & mvarToothID & "," _
                    & " " & mvarDRC & "," _
                    & " '" & mvarToothName & "'," _
                    & " " & mvarMesialVersion & "," _
                    & " " & mvarDistalVersion & "," _
                    & " " & mvarFacialVersion & "," _
                    & " " & mvarLingualVersion & "," _
                    & " " & mvarRotatedMesial & "," _
                    & " " & mvarRotatedDistal & "," _
                    & " " & mvarIntruded & "," _
                    & " " & mvarExtruded & "," _
                    & " " & CInt(mvarDeleted) & "," _
                    & " " & mvarDentalEncounterID & ")"
                goconn.ExecuteSQL sSQL
            End If
        
    End Select
    
    mvarDirtyFlag = dfNoAction
    SaveTooth = True
    
    'This is wierd because the surfaceset cannot be saved until the tooth is saved.
    'So return the PKID of the Tooth to the surface set so it can use it to save.
    DentalEncounterTeethID = mvarDentalEncounterTeethID
    
Exit Function

SaveToothErr:
    MsgBxARMd Err.Description & " -- " & sSQL, vbCritical, Err.Source
    
End Function

Private Sub SetDirtyFlag()
    
    'This is called when a user sets a property
    'If the current flag is dfAdd, then don't change it until SaveTooth is called
    'otherwise, set it to sfUpdate
    
    'SF 10-11-02
    If mvarDirtyFlag = dfNoAction Then
        mvarDirtyFlag = dfUpdate
    End If
    
'    If mvarDirtyFlag = dfAdd Then
'        Exit Sub
'    Else
'        mvarDirtyFlag = dfUpdate
'    End If
    
End Sub






Public Property Let ToothName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ToothName = 5
    mvarToothName = Mid(vData, 1, 25)
    SetDirtyFlag
End Property


Public Property Get ToothName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ToothName
    ToothName = mvarToothName
End Property



Public Property Let ToothID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ToothID = 5
    mvarToothID = vData
    SetDirtyFlag
End Property


Public Property Get ToothID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ToothID
    ToothID = mvarToothID
End Property



Public Property Let DRC(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DRC = 5
    mvarDRC = vData
    SetDirtyFlag
End Property


Public Property Get DRC() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DRC
    DRC = mvarDRC
End Property



Friend Property Let DentalEncounterID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DentalEncounterID = 5
    mvarDentalEncounterID = vData
    SetDirtyFlag
    
End Property


Public Property Get Deleted() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DentalEncounterID
    Deleted = mvarDeleted
End Property




Public Property Get DentalEncounterID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DentalEncounterID
    DentalEncounterID = mvarDentalEncounterID
End Property

