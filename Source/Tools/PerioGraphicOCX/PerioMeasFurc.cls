VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PerioMeasFurc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements PerioMeasGrphElem

'Public Tooth As String
Public iMorL As Integer     'Default to perUnknownValue
Public iDorF As Integer     'Default to perUnknownValue
Public iL As Integer        'Default to perUnknownValue
Public b_IsDirty As Boolean 'Dirty Flag
Public bShow As Boolean     'User set False (Default) = no show, True = show
Public MyTooth As PerioGraphicTooth  'Save a reference
Private mstrSVGString As String

Private Sub PerioMeasGrphElem_Dispose()
    Set MyTooth = Nothing
End Sub

'
Private Property Let PerioMeasGrphElem_SVGString(RHS As String)
    mstrSVGString = RHS
End Property

Private Property Get PerioMeasGrphElem_SVGString() As String
    PerioMeasGrphElem_SVGString = mstrSVGString
End Property

Private Sub Class_Initialize()
 PerioMeasGrphElem_ResetDataValue  'init all data values
End Sub

Private Sub Class_Terminate()
  Set MyTooth = Nothing  'Release reference
End Sub

Private Sub PerioMeasGrphElem_Clear()
  If b_IsDirty = False Then
    Exit Sub: End If
  
  Dim objSVGObj As Object
  Dim SVGElementId As String
  Dim cssClass As String
  
  On Error Resume Next
  
  'Secondary Teeth
  Select Case MyTooth.mi_Position
    Case 1, 2, 3, 14, 15, 16, 5, 12
        SVGElementId = MyTooth.mstr_SecondaryToothName & "-" & GE_Mesi_Furc_1
        ClearElement SVGElementId
        SVGElementId = MyTooth.mstr_SecondaryToothName & "-" & GE_Mesi_Furc_2
        ClearElement SVGElementId
        SVGElementId = MyTooth.mstr_SecondaryToothName & "-" & GE_Mesi_Furc_3
        ClearElement SVGElementId
        SVGElementId = MyTooth.mstr_SecondaryToothName & "-" & GE_Mesi_Furc_4
        ClearElement SVGElementId
        SVGElementId = MyTooth.mstr_SecondaryToothName & "-" & GE_Dist_Furc_1
        ClearElement SVGElementId
        SVGElementId = MyTooth.mstr_SecondaryToothName & "-" & GE_Dist_Furc_2
        ClearElement SVGElementId
        SVGElementId = MyTooth.mstr_SecondaryToothName & "-" & GE_Dist_Furc_3
        ClearElement SVGElementId
        SVGElementId = MyTooth.mstr_SecondaryToothName & "-" & GE_Dist_Furc_4
        ClearElement SVGElementId
        
        If (MyTooth.mi_Position <> 5) And (MyTooth.mi_Position <> 12) Then
            SVGElementId = MyTooth.mstr_SecondaryToothName & "-" & GE_Ling_Furc_1
            ClearElement SVGElementId
            SVGElementId = MyTooth.mstr_SecondaryToothName & "-" & GE_Ling_Furc_2
            ClearElement SVGElementId
            SVGElementId = MyTooth.mstr_SecondaryToothName & "-" & GE_Ling_Furc_3
            ClearElement SVGElementId
            SVGElementId = MyTooth.mstr_SecondaryToothName & "-" & GE_Ling_Furc_4
            ClearElement SVGElementId
        End If
      Case 4, 13, 17, 18, 19, 30, 31, 32
        SVGElementId = MyTooth.mstr_SecondaryToothName & "-" & GE_Ling_Furc_1
        ClearElement SVGElementId
        SVGElementId = MyTooth.mstr_SecondaryToothName & "-" & GE_Ling_Furc_2
        ClearElement SVGElementId
        SVGElementId = MyTooth.mstr_SecondaryToothName & "-" & GE_Ling_Furc_3
        ClearElement SVGElementId
        SVGElementId = MyTooth.mstr_SecondaryToothName & "-" & GE_Ling_Furc_4
        ClearElement SVGElementId
        SVGElementId = MyTooth.mstr_SecondaryToothName & "-" & GE_Fac_Furc_1
        ClearElement SVGElementId
        SVGElementId = MyTooth.mstr_SecondaryToothName & "-" & GE_Fac_Furc_2
        ClearElement SVGElementId
        SVGElementId = MyTooth.mstr_SecondaryToothName & "-" & GE_Fac_Furc_3
        ClearElement SVGElementId
        SVGElementId = MyTooth.mstr_SecondaryToothName & "-" & GE_Fac_Furc_4
        ClearElement SVGElementId
    End Select
  
    'Primary teeth
    If MyTooth.mstr_PrimaryToothName <> "" Then
        Select Case MyTooth.mi_Position
          Case 4, 5, 12, 13, 20, 21, 28, 29
            SVGElementId = MyTooth.mstr_PrimaryToothName & "-" & GE_Ling_Furc_1
            ClearElement SVGElementId
            SVGElementId = MyTooth.mstr_PrimaryToothName & "-" & GE_Ling_Furc_2
            ClearElement SVGElementId
            SVGElementId = MyTooth.mstr_PrimaryToothName & "-" & GE_Ling_Furc_3
            ClearElement SVGElementId
            SVGElementId = MyTooth.mstr_PrimaryToothName & "-" & GE_Ling_Furc_4
            ClearElement SVGElementId
        End Select
        
        Select Case MyTooth.mi_Position
          Case 4, 5, 12, 13
            SVGElementId = MyTooth.mstr_PrimaryToothName & "-" & GE_Mesi_Furc_1
            ClearElement SVGElementId
            SVGElementId = MyTooth.mstr_PrimaryToothName & "-" & GE_Mesi_Furc_2
            ClearElement SVGElementId
            SVGElementId = MyTooth.mstr_PrimaryToothName & "-" & GE_Mesi_Furc_3
            ClearElement SVGElementId
            SVGElementId = MyTooth.mstr_PrimaryToothName & "-" & GE_Mesi_Furc_4
            ClearElement SVGElementId
            SVGElementId = MyTooth.mstr_PrimaryToothName & "-" & GE_Dist_Furc_1
            ClearElement SVGElementId
            SVGElementId = MyTooth.mstr_PrimaryToothName & "-" & GE_Dist_Furc_2
            ClearElement SVGElementId
            SVGElementId = MyTooth.mstr_PrimaryToothName & "-" & GE_Dist_Furc_3
            ClearElement SVGElementId
            SVGElementId = MyTooth.mstr_PrimaryToothName & "-" & GE_Dist_Furc_4
            ClearElement SVGElementId
          Case 20, 21, 28, 29
            SVGElementId = MyTooth.mstr_PrimaryToothName & "-" & GE_Fac_Furc_1
            ClearElement SVGElementId
            SVGElementId = MyTooth.mstr_PrimaryToothName & "-" & GE_Fac_Furc_2
            ClearElement SVGElementId
            SVGElementId = MyTooth.mstr_PrimaryToothName & "-" & GE_Fac_Furc_3
            ClearElement SVGElementId
            SVGElementId = MyTooth.mstr_PrimaryToothName & "-" & GE_Fac_Furc_4
            ClearElement SVGElementId
        End Select
  
    End If
    
  
    On Error GoTo 0
  
End Sub

Private Sub PerioMeasGrphElem_Draw()
  'Clear the element
  'PerioMeasGrphElem_Clear
  If bShow = False Then
    Exit Sub    'Go Bye Bye
  End If
  
  If b_IsDirty = False Then
    Exit Sub: End If
   b_IsDirty = False       'Not dirty anymore
  
  Dim objSVGObj As Object
  Dim SVGElementId As String
  Dim cssClass As String

    On Error Resume Next

  If MyTooth.mi_ToothType = perPrimary Then
    Select Case MyTooth.mi_Position
      Case 4, 5, 12, 13
        Select Case iMorL
          Case 1
            SVGElementId = MyTooth.mstr_ToothName & "-" & GE_Mesi_Furc_1
            cssClass = CSS_Furc12
          Case 2
            SVGElementId = MyTooth.mstr_ToothName & "-" & GE_Mesi_Furc_2
            cssClass = CSS_Furc12
          Case 3
            SVGElementId = MyTooth.mstr_ToothName & "-" & GE_Mesi_Furc_3
            cssClass = CSS_Furc3
          Case 4
            SVGElementId = MyTooth.mstr_ToothName & "-" & GE_Mesi_Furc_4
            cssClass = CSS_None
        End Select
        If (iMorL >= 1) And (iMorL <= 4) Then
            DrawElement SVGElementId, cssClass
        End If
        Select Case iDorF
          Case 1
            SVGElementId = MyTooth.mstr_ToothName & "-" & GE_Dist_Furc_1
            cssClass = CSS_Furc12
          Case 2
            SVGElementId = MyTooth.mstr_ToothName & "-" & GE_Dist_Furc_2
            cssClass = CSS_Furc12
          Case 3
            SVGElementId = MyTooth.mstr_ToothName & "-" & GE_Dist_Furc_3
            cssClass = CSS_Furc3
          Case 4
            cssClass = CSS_None
            SVGElementId = MyTooth.mstr_ToothName & "-" & GE_Dist_Furc_4
        End Select
        If (iDorF >= 1) And (iDorF <= 4) Then
            DrawElement SVGElementId, cssClass
        End If
        Select Case iL
          Case 1
            cssClass = CSS_Furc12
            SVGElementId = MyTooth.mstr_ToothName & "-" & GE_Ling_Furc_1
          Case 2
            cssClass = CSS_Furc12
            SVGElementId = MyTooth.mstr_ToothName & "-" & GE_Ling_Furc_2
          Case 3
            cssClass = CSS_Furc3
            SVGElementId = MyTooth.mstr_ToothName & "-" & GE_Ling_Furc_3
          Case 4
            SVGElementId = MyTooth.mstr_ToothName & "-" & GE_Ling_Furc_4
            cssClass = CSS_None
        End Select
        If (iL >= 1) And (iL <= 4) Then
            DrawElement SVGElementId, cssClass
        End If
      Case 20, 21, 28, 29
        Select Case iMorL
          Case 1
            SVGElementId = MyTooth.mstr_ToothName & "-" & GE_Ling_Furc_1
            cssClass = CSS_Furc12
          Case 2
            SVGElementId = MyTooth.mstr_ToothName & "-" & GE_Ling_Furc_2
            cssClass = CSS_Furc12
          Case 3
            SVGElementId = MyTooth.mstr_ToothName & "-" & GE_Ling_Furc_3
            cssClass = CSS_Furc3
          Case 4
            SVGElementId = MyTooth.mstr_ToothName & "-" & GE_Ling_Furc_4
            cssClass = CSS_None
        End Select
        If (iMorL >= 1) And (iMorL <= 4) Then
            DrawElement SVGElementId, cssClass
        End If
        Select Case iDorF
          Case 1
            SVGElementId = MyTooth.mstr_ToothName & "-" & GE_Fac_Furc_1
            cssClass = CSS_Furc12
          Case 2
            SVGElementId = MyTooth.mstr_ToothName & "-" & GE_Fac_Furc_2
            cssClass = CSS_Furc12
          Case 3
            SVGElementId = MyTooth.mstr_ToothName & "-" & GE_Fac_Furc_3
            cssClass = CSS_Furc3
          Case 4
            SVGElementId = MyTooth.mstr_ToothName & "-" & GE_Fac_Furc_4
            cssClass = CSS_None
        End Select
        If (iDorF >= 1) And (iDorF <= 4) Then
            DrawElement SVGElementId, cssClass
        End If
    End Select
  Else  'Secondary or Supernumerary
    Select Case MyTooth.mi_Position
      Case 1, 2, 3, 14, 15, 16, 5, 12
        Select Case iMorL
          Case 1
            SVGElementId = MyTooth.mstr_ToothName & "-" & GE_Mesi_Furc_1
            cssClass = CSS_Furc12
          Case 2
            SVGElementId = MyTooth.mstr_ToothName & "-" & GE_Mesi_Furc_2
            cssClass = CSS_Furc12
          Case 3
            SVGElementId = MyTooth.mstr_ToothName & "-" & GE_Mesi_Furc_3
            cssClass = CSS_Furc3
          Case 4
            SVGElementId = MyTooth.mstr_ToothName & "-" & GE_Mesi_Furc_4
            cssClass = CSS_None
        End Select
        If (iMorL >= 1) And (iMorL <= 4) Then
            DrawElement SVGElementId, cssClass
        End If
        Select Case iDorF
          Case 1
            SVGElementId = MyTooth.mstr_ToothName & "-" & GE_Dist_Furc_1
            cssClass = CSS_Furc12
          Case 2
            SVGElementId = MyTooth.mstr_ToothName & "-" & GE_Dist_Furc_2
            cssClass = CSS_Furc12
          Case 3
            SVGElementId = MyTooth.mstr_ToothName & "-" & GE_Dist_Furc_3
            cssClass = CSS_Furc3
          Case 4
            SVGElementId = MyTooth.mstr_ToothName & "-" & GE_Dist_Furc_4
            cssClass = CSS_None
        End Select
        If (iDorF >= 1) And (iDorF <= 4) Then
            DrawElement SVGElementId, cssClass
        End If
        If (MyTooth.mi_Position <> 5) And (MyTooth.mi_Position <> 12) Then
            Select Case iL
              Case 1
                SVGElementId = MyTooth.mstr_ToothName & "-" & GE_Ling_Furc_1
                cssClass = CSS_Furc12
              Case 2
                SVGElementId = MyTooth.mstr_ToothName & "-" & GE_Ling_Furc_2
                cssClass = CSS_Furc12
              Case 3
                SVGElementId = MyTooth.mstr_ToothName & "-" & GE_Ling_Furc_3
                cssClass = CSS_Furc3
              Case 4
                SVGElementId = MyTooth.mstr_ToothName & "-" & GE_Ling_Furc_4
                cssClass = CSS_None
            End Select
            If (iL >= 1) And (iL <= 4) Then
                DrawElement SVGElementId, cssClass
            End If
        End If
      Case 4, 13, 17, 18, 19, 30, 31, 32
        Select Case iMorL
          Case 1
            SVGElementId = MyTooth.mstr_ToothName & "-" & GE_Ling_Furc_1
            cssClass = CSS_Furc12
          Case 2
            SVGElementId = MyTooth.mstr_ToothName & "-" & GE_Ling_Furc_2
            cssClass = CSS_Furc12
          Case 3
            SVGElementId = MyTooth.mstr_ToothName & "-" & GE_Ling_Furc_3
            cssClass = CSS_Furc3
          Case 4
            SVGElementId = MyTooth.mstr_ToothName & "-" & GE_Ling_Furc_4
            cssClass = CSS_None
        End Select
        If (iMorL >= 1) And (iMorL <= 4) Then
            DrawElement SVGElementId, cssClass
        End If
        Select Case iDorF
          Case 1
            SVGElementId = MyTooth.mstr_ToothName & "-" & GE_Fac_Furc_1
            cssClass = CSS_Furc12
          Case 2
            SVGElementId = MyTooth.mstr_ToothName & "-" & GE_Fac_Furc_2
            cssClass = CSS_Furc12
          Case 3
            SVGElementId = MyTooth.mstr_ToothName & "-" & GE_Fac_Furc_3
            cssClass = CSS_Furc3
          Case 4
            SVGElementId = MyTooth.mstr_ToothName & "-" & GE_Fac_Furc_4
            cssClass = CSS_None
        End Select
        If (iDorF >= 1) And (iDorF <= 4) Then
            DrawElement SVGElementId, cssClass
        End If
    End Select
  End If
  
  On Error GoTo 0
  
End Sub

Private Sub PerioMeasGrphElem_Init(ToothObj As PerioGraphicTooth)
   Set MyTooth = ToothObj
   b_IsDirty = True      '1st time around
End Sub

Public Sub SetShowFlag(Flag As Boolean)
  If bShow <> Flag Then
    If Not bShow Then
      b_IsDirty = True: End If
    bShow = Flag: End If
End Sub

Private Sub PerioMeasGrphElem_ResetDataValue()
 iMorL = perUnknownValue
 iDorF = perUnknownValue
 iL = perUnknownValue
 bShow = False
 b_IsDirty = True      '1st time around
End Sub


Private Sub DrawElement(strElementId As String, strClassName As String)
Dim objSVGObj As Object
Dim objElement As Object
Dim strSVG As String
Dim objSVGTooth As Object
Dim objRet As Object

    On Error Resume Next

    GetSVGElementByID MyTooth.mo_SVGDOMDoc, strElementId, objSVGObj
    
    If objSVGObj Is Nothing Then 'need to load the element

        strSVG = MyTooth.MeasSVGString(strElementId)

        Set objElement = MyTooth.mo_SVGDOMWin.parseXML(CStr(strSVG), MyTooth.mo_SVGDOMDoc)


        Set objSVGTooth = MyTooth.mo_SVGDOMDoc.getElementById(MyTooth.mstr_ToothName)
        If Not objSVGTooth Is Nothing Then
            Set objSVGObj = objSVGTooth.appendChild(objElement.firstChild)
        End If


    End If
    
    On Error GoTo 0
    
    'Set drawing properties
    Call objSVGObj.setAttribute("class", CStr(strClassName))
    
End Sub

Private Sub ClearElement(strElementId As String)
    Dim objSVGObj As Object

    GetSVGElementByID MyTooth.mo_SVGDOMDoc, strElementId, objSVGObj

    If objSVGObj Is Nothing Then 'nothing to do
    Else
        Call objSVGObj.setAttribute("class", CStr(CSS_Opac))
    End If
    
End Sub
