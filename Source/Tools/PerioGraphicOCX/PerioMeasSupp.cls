VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PerioMeasSupp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements PerioMeasGrphElem

'Public Tooth As String
Public iMB As Integer       'Default to perUnknownValue
Public iB As Integer        'Default to perUnknownValue
Public iDB As Integer       'Default to perUnknownValue
Public iDL As Integer       'Default to perUnknownValue
Public iL As Integer        'Default to perUnknownValue
Public iML As Integer       'Default to perUnknownValue
Public bShow As Boolean     'User set False (Default) = no show, True = show
Public b_IsDirty As Boolean 'Dirty Flag
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
  
    On Error Resume Next
    
    SVGElementId = MyTooth.mstr_SecondaryToothName & "-" & GE_Fac_Cntr_Supp
    ClearElement SVGElementId
    SVGElementId = MyTooth.mstr_SecondaryToothName & "-" & GE_Fac_Mesi_Supp
    ClearElement SVGElementId
    SVGElementId = MyTooth.mstr_SecondaryToothName & "-" & GE_Fac_Dist_Supp
    ClearElement SVGElementId
    SVGElementId = MyTooth.mstr_SecondaryToothName & "-" & GE_Ling_Cntr_Supp
    ClearElement SVGElementId
    SVGElementId = MyTooth.mstr_SecondaryToothName & "-" & GE_Ling_Mesi_Supp
    ClearElement SVGElementId
    SVGElementId = MyTooth.mstr_SecondaryToothName & "-" & GE_Ling_Dist_Supp
    ClearElement SVGElementId
    
    If MyTooth.mstr_PrimaryToothName <> "" Then
        SVGElementId = MyTooth.mstr_PrimaryToothName & "-" & GE_Fac_Cntr_Supp
        ClearElement SVGElementId
        SVGElementId = MyTooth.mstr_PrimaryToothName & "-" & GE_Fac_Mesi_Supp
        ClearElement SVGElementId
        SVGElementId = MyTooth.mstr_PrimaryToothName & "-" & GE_Fac_Dist_Supp
        ClearElement SVGElementId
        SVGElementId = MyTooth.mstr_PrimaryToothName & "-" & GE_Ling_Cntr_Supp
        ClearElement SVGElementId
        SVGElementId = MyTooth.mstr_PrimaryToothName & "-" & GE_Ling_Mesi_Supp
        ClearElement SVGElementId
        SVGElementId = MyTooth.mstr_PrimaryToothName & "-" & GE_Ling_Dist_Supp
        ClearElement SVGElementId
    End If

    On Error GoTo 0

End Sub

Private Sub PerioMeasGrphElem_Draw()
  'Clear the element
  'PerioMeasGrphElem_Clear
  If bShow = False Then
    Exit Sub: End If
    
  If b_IsDirty = False Then
    Exit Sub: End If
   b_IsDirty = False       'Not dirty anymore
   
    Dim objSVGObj As Object
    Dim SVGElementId As String

    'Do some drawing based on data
    If (iMB <> 0) And (iMB <> perUnknownValue) Then
      SVGElementId = MyTooth.mstr_ToothName & "-" & GE_Fac_Mesi_Supp
      DrawElement SVGElementId
    End If
    If (iB <> 0) And (iB <> perUnknownValue) Then
      SVGElementId = MyTooth.mstr_ToothName & "-" & GE_Fac_Cntr_Supp
      DrawElement SVGElementId
    End If
  
    If (iDB <> 0) And (iDB <> perUnknownValue) Then
      SVGElementId = MyTooth.mstr_ToothName & "-" & GE_Fac_Dist_Supp
      DrawElement SVGElementId
    End If
    If (iDL <> 0) And (iDL <> perUnknownValue) Then
      SVGElementId = MyTooth.mstr_ToothName & "-" & GE_Ling_Dist_Supp
      DrawElement SVGElementId
    End If
    If (iL <> 0) And (iL <> perUnknownValue) Then
      SVGElementId = MyTooth.mstr_ToothName & "-" & GE_Ling_Cntr_Supp
      DrawElement SVGElementId
    End If
    If (iML <> 0) And (iML <> perUnknownValue) Then
      SVGElementId = MyTooth.mstr_ToothName & "-" & GE_Ling_Mesi_Supp
      DrawElement SVGElementId
    End If
    
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
 iMB = perUnknownValue
 iB = perUnknownValue
 iDB = perUnknownValue
 iDL = perUnknownValue
 iML = perUnknownValue
 iL = perUnknownValue
 bShow = False
 b_IsDirty = True       'After reset always need redraw
End Sub


Private Sub DrawElement(strElementId As String)
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
    Call objSVGObj.setAttribute("fill-opacity", "1")
    Call objSVGObj.setAttribute("stroke-width", "none")
    
End Sub

Private Sub ClearElement(strElementId As String)
  Dim objSVGObj As Object
  
  GetSVGElementByID MyTooth.mo_SVGDOMDoc, strElementId, objSVGObj
  
  If objSVGObj Is Nothing Then 'nothing to do
  Else
    Call objSVGObj.setAttribute("fill-opacity", "0")
    Call objSVGObj.setAttribute("stroke-width", "0")
  End If
End Sub
