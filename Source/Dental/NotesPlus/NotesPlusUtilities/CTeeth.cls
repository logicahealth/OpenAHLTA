VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTeeth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mColItems As Collection
Private mobjTooth As CTooth
Private mobjDAS As ICHCSII_DAS
Private mobjRS As ADODB.Recordset
Private mstrTeethPath As String

Private Sub Class_Initialize()

    Set goChcsiiConn = New CHCSII_CONN.Conn
    Set mobjDAS = goChcsiiConn.CHCSII_DAS
    
    Set mColItems = New Collection
    
'    Dim blnTheater As Boolean
'
'    blnTheater = gobjCWShared.UseLocalDB
'
'    If blnTheater Then
'        Set mobjDAS = New GEMS_DAS
'    Else
'        Set mobjDAS = New CHCSII_DataAccess.CHCSII_DAS
'    End If
End Sub

Private Sub Class_Terminate()
    Set mobjTooth = Nothing
    Set mColItems = Nothing
    Set mobjDAS = Nothing
    Set mobjRS = Nothing
End Sub

Public Property Let TeethPath(ByVal strPath As String)
    mstrTeethPath = strPath
End Property
Public Property Get TeethPath() As String
    TeethPath = mstrTeethPath
End Property

'********************************************************************
'Sub: Remove
'Author: KDunne
'Date: 1/9/2004
'Purpose: Removes a singe item from the collection.
'********************************************************************
Public Sub Remove(Optional ByVal objRemove As CTooth)
    Dim objTooth As CTooth
    Dim intIndex As Integer
    
    intIndex = 1
    
    'If the tooth wasn't passed in, then use the current, module level tooth.
    If objRemove Is Nothing Then
        Set objRemove = mobjTooth
    End If
    
    For Each objTooth In mColItems
    
        If objTooth Is objRemove Then
            mColItems.Remove intIndex
            Exit For
        End If
        
        intIndex = intIndex + 1
    Next
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mColItems.[_NewEnum]
End Property
Friend Property Set Collection(ByVal colTeeth As Collection)
    Set mColItems = colTeeth
End Property
Public Property Get Tooth() As CTooth
    Set Tooth = mobjTooth
End Property
Public Function Item(ByVal Index As Variant) As CTooth
Attribute Item.VB_UserMemId = 0
   Set Item = mColItems.Item(Index)
End Function
'*******************************************************************
'Sub: Add_Tooth
'Author: KDunne
'Date: 10/16/2001
'Purpose: Adds a CTooth to the CTeeth collection class.
'*******************************************************************
Public Sub Add_Tooth(ByVal intNumber As Integer, ByVal intPosition As Integer, _
        ByVal blnChecked As Boolean, ByVal strName As String, _
        ByVal blnMissing As Boolean, ByVal blnImplant As Boolean, ByVal blnPontic As Boolean, _
        Optional ByVal blnTemp As Boolean = False)
    Set mobjTooth = New CTooth
    
    With mobjTooth
        .ToothNumber = intNumber
        .ToothPosition = intPosition
        .Checked = blnChecked
        .Missing = blnMissing
        .ToothName = strName
        .Implant = blnImplant
        .Pontic = blnPontic
        .Temporary = blnTemp
    End With
    
    mColItems.Add mobjTooth

End Sub

'*******************************************************************
'Function: Load_Tooth_By_Name
'Author: KDunne
'Date: 10/27/2001
'Returns: boolean
'Purpose: Loops through collection of tooth objects, and searches
'         for the given tooth name based on its display name. If it finds it, then
'         then the public tooth object is set to that tooth, and
'         the name is returned.
'Notes: The internal collection of tooth objects must be set prior
'       to calling this function
'*******************************************************************
Public Function Load_Tooth_By_Name(ByVal strName As String) As Boolean
    Dim blnReturn As Boolean
    Dim objTooth As CTooth
    
    blnReturn = False
    
    For Each objTooth In mColItems
        If objTooth.ToothName = strName Then
            blnReturn = True
            Set mobjTooth = objTooth
            Exit For
        End If
    Next
    
    Set objTooth = Nothing
    
    Load_Tooth_By_Name = blnReturn
End Function
'*******************************************************************
'Function: Load_Tooth_By_Position
'Author: KDunne
'Date: 10/24/2001
'Returns: boolean
'Purpose: Loops through collection of tooth objects, and searches
'         for the given tooth position (1-32). If it finds it, then
'         then the public tooth object is set to that tooth, otherwise
'         the function returns false.
'Notes: The internal collection of tooth objects must be set prior
'       to calling this function
'*******************************************************************
Public Function Load_Tooth_By_Position(ByVal lngPos As Long) As Boolean
    Dim blnReturn As Boolean
    Dim objTooth As CTooth
    
    blnReturn = False
    
    For Each objTooth In mColItems
        If objTooth.ToothPosition = lngPos Then
            blnReturn = True
            Set mobjTooth = objTooth
            Exit For
        End If
    Next
    
    Set objTooth = Nothing
    
    Load_Tooth_By_Position = blnReturn
End Function

'*******************************************************************
'Function: Load_Tooth_By_Position
'Author: KDunne
'Date: 10/27/2001
'Returns: boolean
'Purpose: Loops through collection of tooth objects, and searches
'         for the given tooth name based on its number. If it finds it, then
'         then the public tooth object is set to that tooth, and
'         the name is returned.
'Notes: The internal collection of tooth objects must be set prior
'       to calling this function
'*******************************************************************
Public Function Get_Tooth_Name(ByVal lngNumber As Long) As String

    Dim strReturn As String
    
    If Load_Tooth_By_Number(lngNumber) Then
        If mobjTooth.ToothName = "" Then
            strReturn = mobjTooth.ToothNumber
        Else
            strReturn = mobjTooth.ToothName
        End If
    Else
        strReturn = LookupToothName(lngNumber)
    End If
    
    Get_Tooth_Name = strReturn
End Function
'*******************************************************************
'Function: Load_Tooth_By_Number
'Author: KDunne
'Date: 10/24/2001
'Returns: boolean
'Purpose: Loops through collection of tooth objects, and searches
'         for the given tooth number (1-52). If it finds it, then
'         then the public tooth object is set to that tooth, otherwise
'         the function returns false.
'Notes: The internal collection of tooth objects must be set prior
'       to calling this function
'*******************************************************************
Public Function Load_Tooth_By_Number(ByVal lngId As Long) As Boolean
    Dim blnReturn As Boolean
    
    blnReturn = False
    
    For Each mobjTooth In mColItems
        If mobjTooth.ToothNumber = lngId Then
            blnReturn = True
            Exit For
        End If
    Next
    
    Load_Tooth_By_Number = blnReturn
End Function

'*******************************************************************
'Sub: Update_Tooth
'Author: KDunne
'Date: 10/16/2001
'Purpose: Updates a CTooth in the CTeeth collection class. Can only
'         update the Checked property.
'*******************************************************************
Public Sub Update_Tooth(ByVal intNumber As Integer, _
    ByVal blnChecked As Boolean, Optional ByVal blnMissing As Boolean)
    
    Dim strName As String
    
    If Load_Tooth_By_Position(intNumber) Then
        mobjTooth.Checked = blnChecked
    Else
        Call Add_Tooth(intNumber, intNumber, blnChecked, "", True, False, False)
    End If

    
End Sub
'*******************************************************************
'Sub: BuildValueString
'Author: KDunne
'Date: 10/16/2001
'Purpose: Loops through Collection and builds a string containing
'         all of the checked values.
'*******************************************************************
Public Function BuildValueString() As String
    Dim x As Integer
    Dim strValues As String
    Dim objTooth As CTooth


    strValues = Empty
    For Each objTooth In mColItems
        If objTooth.Checked Then
            strValues = strValues & const_ITEM_DELIMETER & objTooth.ToothNumber
        End If
    Next

    'Remove the first Item Delimeter
    strValues = Replace(strValues, const_ITEM_DELIMETER, "", , 1)
    
    BuildValueString = strValues

End Function

'*******************************************************************
'Sub: BuildNameString
'Author: KDunne
'Date: 11/05/2001
'Purpose: Loops through Collection and builds a string containing
'         names of the checked values.
'*******************************************************************
Friend Function BuildNameString() As String
    Dim x As Integer
    Dim strValues As String
    Dim objTooth As CTooth

    strValues = Empty
    For Each objTooth In mColItems
        If objTooth.Checked Then
            strValues = strValues & const_DISPLAY_ITEM_DELIMETER & objTooth.ToothName
        End If
    Next
    
    strValues = Replace(strValues, const_DISPLAY_ITEM_DELIMETER, "", , 1)
    
    BuildNameString = strValues
    
    Set objTooth = Nothing
End Function

Public Function BuildDisplayString() As String
    
    Dim strValues As String
    Dim objTooth As CTooth

On Error GoTo ErrHnd

    strValues = Empty
    For Each objTooth In mColItems
        If objTooth.Checked Then
            'SF 3-19-04 SCR 50380 changed objTooth.ToothNumber to objTooth.ToothName
            strValues = strValues & const_DISPLAY_ITEM_DELIMETER & objTooth.ToothName
        End If
    Next
    
    strValues = Replace(strValues, const_DISPLAY_ITEM_DELIMETER, "", , 1)
    
    BuildDisplayString = strValues
    
    Set objTooth = Nothing

Exit Function

ErrHnd:
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

'*******************************************************************
'Function: Initialize
'Author: KDunne
'Date: 10/24/2001
'Returns: boolean
'Purpose: accepts a string of values, in the format 1|2|3, and
'         initializes the internal collection.

'*******************************************************************
Public Function Initialize(ByVal strValues) As Boolean
    Dim arrValues() As String
    Dim objTooth As CTooth
    Dim x As Integer
    
    ResetTeeth
    arrValues = Split(strValues, const_ITEM_DELIMETER)
    
    For x = 0 To UBound(arrValues)
        If Load_Tooth_By_Number(arrValues(x)) Then
            mobjTooth.Checked = True
        Else
            Add_Tooth arrValues(x), Get_Tooth_Position(arrValues(x)), True, _
                LookupToothName(arrValues(x)), False, False, False, True
        End If
        
    Next

    CheckValueString strValues, mstrTeethPath
    Initialize = True
End Function
'******************************************************************************
'Sub: ResetTeeth
'Author:KDunne
'Date: 10/26/2001
'Purpose: Loops through the collection of teeth and unchecks each tooth.
'******************************************************************************
Public Sub ResetTeeth()
    Dim objTooth As CTooth
    
    For Each objTooth In mColItems
        objTooth.Checked = False
        
        If objTooth.Temporary Then
            Me.Remove (objTooth)
        End If
    Next
    
    Set objTooth = Nothing
End Sub

'******************************************************************************
'Sub: DisableTeeth
'Author:KDunne
'Date: 11/26/2001
'Purpose: Loops through the collection of teeth and disables each tooth.
'******************************************************************************
Public Sub DisableTeeth()
    Dim objTooth As CTooth
    
    For Each objTooth In mColItems
        objTooth.Enabled = False
    Next
    
    Set objTooth = Nothing
End Sub
'******************************************************************************
'Sub: LookupToothName
'Author:KDunne
'Date: 1/8/2004
'Purpose: If our tooth id does not exist in the collection, we need to convert
'         id to name.
'******************************************************************************
Private Function LookupToothName(ByVal lngToothId As Long) As String
    Dim strReturn As String
    
    Select Case lngToothId
        Case Is > 52
            strReturn = ""
        Case Is > 32
            'Tooth is secondary tooth in the 20 - 29 range
            strReturn = UCase(Chr(lngToothId - 32 + 64))
        Case Is < 0
            'Invalid tooth Number
            strReturn = Empty
        Case Else
            'Tooth is secondary
            strReturn = lngToothId
    End Select
     
    LookupToothName = Trim(strReturn)
End Function

Private Sub CheckValueString(ByVal strValues As String, ByVal strPath As String)
    Dim strReturn As String
    Dim arrValues() As String
    Dim x As Integer
    
    arrValues = Split(strValues, const_ITEM_DELIMETER)
    
    For x = 0 To UBound(arrValues)
        If Not CheckToothPath(arrValues(x), strPath) Then
            If Load_Tooth_By_Number(arrValues(x)) Then
                mobjTooth.ChangeOralState True
            End If
        End If
    Next
        
End Sub
'*******************************************************************
'Function: CheckToothPath
'Author: KDunne
'Date: 10/24/2001
'Purpose: Checks to see if a given tooth is in the supplied tooth path
'*******************************************************************
Private Function CheckToothPath(ByVal lngToothId As Long, ByVal strPath As String) As Boolean
    Dim blnReturn As Boolean
    Dim intIndex As Integer
    Dim intCount As Integer
    Dim arrPath() As String
    
    blnReturn = False
    
    
    If strPath = Empty Then
        blnReturn = False
        GoTo FuncExit
    End If
    
    arrPath = Split(strPath, const_ITEM_DELIMETER)
    
    intCount = UBound(arrPath)
    
    For intIndex = 0 To intCount
        If arrPath(intIndex) = lngToothId Then
            blnReturn = True
            GoTo FuncExit
        End If
    Next 'intIndex
    
FuncExit:
    CheckToothPath = blnReturn
    
End Function
'*****************************************************************
'Function:
'Author:KDunne
'Date:1/22/2002
'Purpose: Given a tooth, (1-52) or (A-T), function returns the
'           tooth's position (1-32)
'*****************************************************************
Public Function Get_Tooth_Position(ByVal strTooth As String) As Integer
    Dim intToothNum As Integer
    Dim intReturn As Integer
    
    If IsNumeric(strTooth) Then
        intToothNum = CInt(strTooth)
        
        Select Case intToothNum
            Case Is > 52
                intReturn = 0
            Case Is > 42
                intReturn = intToothNum - 23
            Case Is > 32
                intReturn = intToothNum - 29
            Case Else
                intReturn = intToothNum
        End Select
            
    Else
        intReturn = Get_Tooth_Position(strTooth)
    End If

    Get_Tooth_Position = intReturn
End Function
