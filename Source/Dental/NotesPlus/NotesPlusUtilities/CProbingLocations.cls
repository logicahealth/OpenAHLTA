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
Attribute VB_Name = "CProbingLocations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mstrValueString As String
Private mlngMFDepth As Long
Private mlngFDepth As Long
Private mlngDFDepth As Long
Private mlngDLDepth As Long
Private mlngLDepth As Long
Private mlngMLDepth As Long
Private mstrDisplayString As String

Public Enum ProbingLocationEnum
    Mesiofacial
    Facial
    Distofacial
    Distolingual
    Lingual
    Mesiolingual
End Enum

Public Property Get MesiofacialDepth() As Long
    MesiofacialDepth = mlngMFDepth
End Property
Public Property Let MesiofacialDepth(ByVal lngDepth As Long)
    mlngMFDepth = lngDepth
End Property

Public Property Get FacialDepth() As Long
    FacialDepth = mlngFDepth
End Property
Public Property Let FacialDepth(ByVal lngDepth As Long)
    mlngFDepth = lngDepth
End Property

Public Property Get DistofacialDepth() As Long
    DistofacialDepth = mlngDFDepth
End Property
Public Property Let DistofacialDepth(ByVal lngDepth As Long)
    mlngDFDepth = lngDepth
End Property

Public Property Get DistolingualDepth() As Long
    DistolingualDepth = mlngDLDepth
End Property
Public Property Let DistolingualDepth(ByVal lngDepth As Long)
    mlngDLDepth = lngDepth
End Property

Public Property Get LingualDepth() As Long
    LingualDepth = mlngLDepth
End Property
Public Property Let LingualDepth(ByVal lngDepth As Long)
    mlngLDepth = lngDepth
End Property

Public Property Get MesiolingualDepth() As Long
    MesiolingualDepth = mlngMLDepth
End Property
Public Property Let MesiolingualDepth(ByVal lngDepth As Long)
    mlngMLDepth = lngDepth
End Property

Public Property Get ValueString() As String
    ValueString = mstrValueString
End Property

Public Property Get DisplayString() As String
    DisplayString = mstrDisplayString
End Property

Private Sub Class_Initialize()
    'Intialize all of our depths to 0 MM
     mlngMFDepth = 0
     mlngFDepth = 0
     mlngDFDepth = 0
     mlngDLDepth = 0
     mlngLDepth = 0
     mlngMLDepth = 0
End Sub
'*************************************************************************
'Function: Load_Value_String
'Author: KDunne
'Date: 11/15/2001
'Purpose: Loads a string with existing values
'*************************************************************************
Public Function Load_Value_String(ByVal strString As String) As Boolean
    Dim arrValues() As String
    Dim arrDepths() As String
    Dim blnReturn As Boolean
    Dim x As Byte
    Dim intDepth As Integer
    
    mstrValueString = Trim(strString)
    
    If mstrValueString = "" Then
        'No values to load
        blnReturn = True
        mstrValueString = BuildValueString
    Else
        arrValues = Split(mstrValueString, const_ITEM_DELIMETER)
        For x = 0 To UBound(arrValues)
            If arrValues(x) <> "" Then
                arrDepths = Split(arrValues(x), const_VALUE_DELIMETER)
                intDepth = arrDepths(1)
                Select Case arrDepths(0)
                    Case ProbingLocationEnum.Distofacial
                        mlngDFDepth = intDepth
                    Case ProbingLocationEnum.Distolingual
                        mlngDLDepth = intDepth
                    Case ProbingLocationEnum.Facial
                        mlngFDepth = intDepth
                    Case ProbingLocationEnum.Lingual
                        mlngLDepth = intDepth
                    Case ProbingLocationEnum.Mesiofacial
                        mlngMFDepth = intDepth
                    Case ProbingLocationEnum.Mesiolingual
                        mlngMLDepth = intDepth
                End Select
            End If

        Next
    End If
    
    Load_Value_String = True
End Function
'*************************************************************************
'Function: BuildValueString
'Author: KDunne
'Date: 11/15/2001
'Purpose: Builds value string with user's selections.
'*************************************************************************
Private Function BuildValueString() As String
    BuildValueString = ProbingLocationEnum.Distofacial & const_VALUE_DELIMETER & _
        mlngDFDepth & const_ITEM_DELIMETER & ProbingLocationEnum.Distolingual & _
        const_VALUE_DELIMETER & mlngDLDepth & const_ITEM_DELIMETER & ProbingLocationEnum.Facial & _
        const_VALUE_DELIMETER & mlngFDepth & const_ITEM_DELIMETER & ProbingLocationEnum.Lingual & _
        const_VALUE_DELIMETER & mlngLDepth & const_ITEM_DELIMETER & ProbingLocationEnum.Mesiofacial & _
        const_VALUE_DELIMETER & mlngMFDepth & const_ITEM_DELIMETER & ProbingLocationEnum.Mesiolingual & _
        const_VALUE_DELIMETER & mlngMLDepth & const_ITEM_DELIMETER
        
End Function
'*************************************************************************
'Function: Get_User_Input
'Author: KDunne
'Date: 11/15/2001
'Purpose: Prompts user for input
'*************************************************************************
Public Function Get_User_Input() As Boolean
    Dim x As Integer
    Dim strString As String
    Dim strDispString As String
    
    If frmProbingLocations.Load_Value_String(mstrValueString) Then
    
        Load frmProbingLocations
        
        With frmProbingLocations
            
            cwobjCWComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, cwlMyProgId
            .Show vbModal
            cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, cwlMyProgId
        
            If .UserCancelled = False Then
                strString = Empty
                strDispString = Empty
                For x = 0 To .cmbDepth.UBound
                    strString = strString & x & const_VALUE_DELIMETER & _
                        .cmbDepth(x).ListIndex & const_ITEM_DELIMETER
                Next
                mstrValueString = strString
                mstrDisplayString = Build_Display_String(mstrValueString)
            End If
            
        End With
        
        Unload frmProbingLocations
    End If
    
    Get_User_Input = True
    
End Function
'*************************************************************************
'Function: Build_Display_String
'Author: KDunne
'Date: 11/15/2001
'Purpose: Builds a display string from exisitng value string.
'Notes: Assumes that value string is not empty
'*************************************************************************
Public Function Build_Display_String(ByVal strValues As String) As String
    Dim arrValues() As String
    Dim arrDepths() As String
    Dim strReturn As String
    Dim x As Integer
    
    strReturn = Empty
    
    arrValues = Split(strValues, const_ITEM_DELIMETER)
    For x = 0 To UBound(arrValues)
        If arrValues(x) <> "" Then
            arrDepths = Split(arrValues(x), const_VALUE_DELIMETER)
            If arrDepths(1) > 0 Then
                strReturn = strReturn & GetLocationName(x) & " depth is " & _
                    arrDepths(1) & " MM. "
            End If
        End If

    Next
    
    Build_Display_String = strReturn
End Function
'*************************************************************************
'Function: GetLocationName
'Author: KDunne
'Date: 11/15/2001
'Purpose: Returns description for a Location Name enum.
'*************************************************************************
Private Function GetLocationName(ByVal lngValue As ProbingLocationEnum) As String
    Dim strReturn As String
    
    Select Case lngValue
                Case ProbingLocationEnum.Distofacial
                    strReturn = "Distofacial"
                Case ProbingLocationEnum.Distolingual
                    strReturn = "Distolingual"
                Case ProbingLocationEnum.Facial
                    strReturn = "Facial"
                Case ProbingLocationEnum.Lingual
                    strReturn = "Lingual"
                Case ProbingLocationEnum.Mesiofacial
                    strReturn = "Mesiofacial"
                Case ProbingLocationEnum.Mesiolingual
                    strReturn = "Mesiolingual"
    
    End Select
    
    GetLocationName = strReturn
End Function

