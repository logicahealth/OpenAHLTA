VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAddFindings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mstrValueString As String
Private mstrDisplayString As String
Private mstrToothNumber As String
Private mlngAddFinding As NotesPlusUtil_AddFindings
Private mobjDataLayer As CHCSII_Dental_DXTX.clsNotes

Public Enum NotesPlusUtil_AddFindings
    npUtil_None
    npUtil_ShortFill
    npUtil_LongFill
    npUtil_NonDenseObturation
    npUtil_UntreatedCanal
    npUtil_SilverPoint
    npUtil_Post
    npUtil_Perforation
    npUtil_VerticalRootFracture
End Enum

Public Enum ExamEndoAddFind_Canals
    None
    Mesial
    Mesiobuccal
    Mesiobuccal2
    Mesiobuccal3
    Mesiolingual
    Distal
    DistalBuccal
    Distolingual
    Facial
    Palatal
    Palatal2
    Lingual
    NOS
End Enum

Public Property Get Data_Layer() As CHCSII_Dental_DXTX.clsNotes
    Set Data_Layer = mobjDataLayer
End Property
Public Property Set Data_Layer(objDataLayer As CHCSII_Dental_DXTX.clsNotes)
    Set mobjDataLayer = objDataLayer
End Property
Public Property Get Additional_Finding() As NotesPlusUtil_AddFindings
    Additional_Finding = mlngAddFinding
End Property
Public Property Let Additional_Finding(ByVal lngValue As NotesPlusUtil_AddFindings)
    mlngAddFinding = lngValue
End Property

Public Property Get ValueString() As String
    ValueString = mstrValueString
End Property
Public Property Get DisplayString() As String
    DisplayString = mstrDisplayString
End Property

Public Property Get Tooth_Number() As String
    Tooth_Number = mstrToothNumber
End Property
'*******************************************************************
'Sub: Get_Additional_Findings_Enum_Desc
'Author: KDunne
'Date: 11/29/2001
'Purpose: Accepts an Additional Findings Enum and returns the appropriate
'           string description
'*******************************************************************
Public Function Get_Additional_Findings_Enum_Desc(ByVal lngValue As NotesPlusUtil_AddFindings) As String
    Dim strReturn As String
    
    Select Case lngValue
        Case NotesPlusUtil_AddFindings.npUtil_ShortFill
            strReturn = "Short fill"
        Case NotesPlusUtil_AddFindings.npUtil_LongFill
            strReturn = "Long fill"
        Case NotesPlusUtil_AddFindings.npUtil_NonDenseObturation
            strReturn = "Non-dense obturation"
        Case NotesPlusUtil_AddFindings.npUtil_UntreatedCanal
            strReturn = "Untreated canal"
        Case NotesPlusUtil_AddFindings.npUtil_SilverPoint
            strReturn = "Silver point"
        Case NotesPlusUtil_AddFindings.npUtil_Post
            strReturn = "Post"
        Case NotesPlusUtil_AddFindings.npUtil_Perforation
            strReturn = "Perforation"
        Case NotesPlusUtil_AddFindings.npUtil_VerticalRootFracture
            strReturn = "Vertical Root Fracture"
        Case Else
            strReturn = Empty
    
    End Select
    
    Get_Additional_Findings_Enum_Desc = strReturn
End Function

'*******************************************************************
'Sub: Build_Display_String
'Author: KDunne
'Date: 11/29/2001
'Purpose: Deconstructs value string and builds a string which can be
'         presented to the user.
'*******************************************************************
Public Function Build_Display_String(ByVal strValues As String) As String
    Dim lngAddFind As NotesPlusUtil_AddFindings
    Dim strReturn As String
    Dim arrAddFind() As String
    Dim strAddFind As String
    Dim strCanals As String
    Dim arrCanals() As String
    Dim objToothCtrl As CToothControl
    
    Dim x As Integer
    
    
    If strValues = "" Then
        strReturn = Empty
    Else
        Set objToothCtrl = New CToothControl
        arrCanals = Split(strValues, const_VALUE_DELIMETER)
        For x = 0 To UBound(arrCanals)
            strCanals = strCanals & const_DISPLAY_ITEM_DELIMETER & objToothCtrl.Get_Canal_Description(arrCanals(x))
        Next
        
        strCanals = Replace(strCanals, const_DISPLAY_ITEM_DELIMETER, "", , 1)
        
        If strCanals <> "" Then
            strReturn = strCanals
        End If


    
    
    End If
'    If strValues = "" Then
'        strReturn = Empty
'    Else
'        arrAddFind = Split(strValues, const_ITEM_DELIMETER)
'        lngAddFind = arrAddFind(0)
'
'        strReturn = Get_Additional_Findings_Enum_Desc(lngAddFind)
'
'        If AddFindingsIsCanal(lngAddFind) Then
'
'            Set objToothCtrl = New CToothControl
'            strCanals = Empty
'            If UBound(arrAddFind) = 1 Then
'                If arrAddFind(1) <> "" Then
'                    arrCanals = Split(arrAddFind(1), const_VALUE_DELIMETER)
'                    For x = 0 To UBound(arrCanals)
'                        strCanals = strCanals & const_DISPLAY_ITEM_DELIMETER & objToothCtrl.Get_Canal_Description(arrCanals(x))
'                    Next
'
'                    strCanals = Replace(strCanals, const_DISPLAY_ITEM_DELIMETER, "", , 1)
'
'                    If strCanals <> "" Then
'                        strReturn = strReturn & " - " & strCanals
'                    End If
'                End If
'            End If
'
'        End If
'    End If
    Set objToothCtrl = Nothing
    Build_Display_String = strReturn
End Function
'*******************************************************************
'Sub: BuildValueString
'Author: KDunne
'Date: 11/29/2001
'Purpose: Loops through user's choices and builds a string containing
'         all of the checked values.
'*******************************************************************
Private Function Build_Value_String() As String
    Dim strReturn As String
    Dim x As Integer
    Dim strCanals As String
       
    If frmAddFindings.fraCanals.Enabled Then
        strReturn = strReturn & const_ITEM_DELIMETER
        For x = 1 To frmAddFindings.chkCanal.Count
            If frmAddFindings.chkCanal(x).Value = vbChecked Then
                strCanals = strCanals & const_VALUE_DELIMETER & x
            End If
        
        Next
        
        strCanals = Replace(strCanals, const_VALUE_DELIMETER, "", , 1)
        
        strReturn = strCanals
    End If
    
    Build_Value_String = strReturn
End Function
'*******************************************************************
'Function: Load_Value_String
'Author: KDunne
'Date: 10/08/2001
'Returns: boolean. True if successful, False if error occurs
'Purpose: Pre-populates user's choices from | deliminated string
'NOtes: Datalayer object must be set prior to calling this function.
'*******************************************************************
Public Function Load_Value_String(ByVal lngAddFind As Long, ByVal strValues As String, ByVal strToothNum As String) As Boolean
    Dim blnReturn As Boolean
    Dim arrAddFind() As String
    Dim arrCanals() As String
    Dim lngCanal As ToothControlCanalsEnum
    Dim x As Integer
    
    If mobjDataLayer Is Nothing Then
        blnReturn = False
        GoTo FuncExit
    Else
        Load_Tooth_Info mobjDataLayer.PatientID, strToothNum
    End If
    
    blnReturn = Load_Additional_Findings(lngAddFind)
    frmAddFindings.Show_Canal_Frame True
    If strValues = "" Then
        
    Else
    
        arrCanals = Split(strValues, const_VALUE_DELIMETER)
        
        frmAddFindings.cmbAddFinds.ListIndex = lngAddFind
        

        For x = 0 To UBound(arrCanals)
            frmAddFindings.chkCanal(arrCanals(x)).Value = vbChecked
        Next
        
    End If
    Load frmAddFindings
FuncExit:
    Load_Value_String = blnReturn
End Function
'*******************************************************************
'Function: Get_User_Input
'Author: KDunne
'Date: 11/29/2001
'Returns: boolean. True if successful, False if error occurs or
'         User cancels.
'Purpose: Prompts user for Input.
'*******************************************************************
Public Function Get_User_Input() As Boolean
    Dim blnReturn As Boolean
    
    cwobjCWComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, cwlMyProgId
    frmAddFindings.Show vbModal
    cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, cwlMyProgId
    
    If frmAddFindings.User_Cancelled Then
        mstrValueString = Empty
        mstrDisplayString = Empty
        mstrToothNumber = Empty
        blnReturn = False
    Else
        mstrToothNumber = frmAddFindings.ucNotesToothCtrl.Teeth.BuildValueString
        mstrValueString = Build_Value_String()
        mstrDisplayString = Build_Display_String(mstrValueString)
        mlngAddFinding = frmAddFindings.cmbAddFinds.ListIndex
        
        If mstrToothNumber = "" Or mstrToothNumber = "0" Then
            mstrToothNumber = Empty
            mstrValueString = Empty
            mstrDisplayString = Empty
        End If
        
        blnReturn = True
    End If
    
    Unload frmAddFindings
    
    Get_User_Input = blnReturn
    
End Function
'**********************************************************************
'Function:Load_Additional_Findings
'Author:KDunne
'Date: 11/29/2001
'Purpose: Loads Additional Findings Combo box on frmAddFindings
'**********************************************************************
Private Function Load_Additional_Findings(ByVal intIndex As Integer) As Boolean
    Dim x As Integer
    Dim objComboBox As VB.ComboBox
    
    Set objComboBox = frmAddFindings.cmbAddFinds
    
    For x = 0 To 8
        objComboBox.AddItem Get_Additional_Findings_Enum_Desc(x), x
    Next
    
    objComboBox.ListIndex = intIndex
    
    Load_Additional_Findings = True
    Set objComboBox = Nothing
End Function



'*******************************************************************
'Sub:Load_Tooth_Info
'Author:KDunne
'Date:1/39/2002
'Purpose: Loads and populates the teeth control.
'********************************************************************
Public Sub Load_Tooth_Info(ByVal lngPatUnitId As Long, ByVal strToothNumber As String)
    Dim objToothCtrl As CToothControl
    
    Set objToothCtrl = New CToothControl
    
    Set objToothCtrl.DentalNotes = mobjDataLayer
    
    If objToothCtrl.Load_Tooth_Information(lngPatUnitId, strToothNumber) Then
        Set frmAddFindings.ucNotesToothCtrl.Teeth = objToothCtrl.Teeth
        frmAddFindings.ucNotesToothCtrl.Populate_Tool_Bar
        frmAddFindings.Load_Canal_Info objToothCtrl.Teeth.BuildDisplayString
    End If
    
    Set objToothCtrl = Nothing
End Sub


Private Sub Class_Terminate()
    Set mobjDataLayer = Nothing
End Sub
