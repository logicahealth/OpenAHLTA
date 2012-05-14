VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PerioMeasMob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements PerioMeasGrphElem

'Public Tooth As String
Public iData As Integer     'Default to 0
Public bShow As Boolean     'User set False (Default) = no show, True = show
Public b_IsDirty As Boolean 'Dirty Flag
Public MyTooth As PerioGraphicTooth  'Save a reference
Private Const RangeMax As Integer = 3  'Max Range for mobility
Private Const RangeMin As Integer = 1  'Min value for mobility
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
    SVGElementId = MyTooth.mstr_SecondaryToothName & "-" & GE_Mobility
    ClearElement SVGElementId
    
  If MyTooth.mstr_PrimaryToothName <> "" Then
    SVGElementId = MyTooth.mstr_PrimaryToothName & "-" & GE_Mobility
    ClearElement SVGElementId
  End If

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
    Dim objText As Object

    'Do some drawing based on data
   ' If (iData <> 0) Then
   If (iData <= RangeMax) And (iData >= RangeMin) Then
        SVGElementId = MyTooth.mstr_ToothName & "-" & GE_Mobility
        DrawElement SVGElementId
        
      'change the text data of the element
      SVGElementId = MyTooth.mstr_ToothName & "-" & GE_Mobility & "-" & GE_TSpan
      SetText SVGElementId, CStr(iData)
   End If
End Sub

Private Sub PerioMeasGrphElem_Init(ToothObj As PerioGraphicTooth)
   Set MyTooth = ToothObj
   b_IsDirty = True       'After reset always need redraw
End Sub

Public Sub SetShowFlag(Flag As Boolean)
  If bShow <> Flag Then
    If Not bShow Then
      b_IsDirty = True: End If
    bShow = Flag: End If
End Sub

Private Sub PerioMeasGrphElem_ResetDataValue()
  iData = 0
  bShow = False
  b_IsDirty = True       'After reset always need redraw
End Sub

Private Sub SetText(strElementId As String, strText As String)
Dim objSVGObj As Object
Dim objElement As Object
Dim strSVG As String
Dim objSVGTooth As Object
Dim objRet As Object
Dim objText As Object

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
    
    'Set the text value
      Set objText = objSVGObj.getFirstChild()
      Call objText.setNodeValue(CStr(strText))

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
  
    On Error Resume Next
  GetSVGElementByID MyTooth.mo_SVGDOMDoc, strElementId, objSVGObj
  
  If objSVGObj Is Nothing Then 'nothing to do
  Else
    Call objSVGObj.setAttribute("fill-opacity", "0")
    Call objSVGObj.setAttribute("stroke-width", "0")
  End If
  
  On Error GoTo 0
End Sub
