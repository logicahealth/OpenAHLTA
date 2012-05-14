VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SVG_Lib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'This stores the SVG string for the Dxs and Txs per tooth per surface
Private mvarSVGDxsTxs As colSVGDxsTxs

'This stores the svg for each tooth
Private mvarSVGTeethDef As colSVGTeethDef
'Public colTeeth as Collection

'This stores the svg string for each surface on each tooth, including the svg for the event handlers
Private mvarSVGSurfaces As colSVGSurfaces
'Public colSurfaces As Collection

'This stores a mapping of what tooth has what surfaces for crown, roots, etc.
Private mvarSurfaceGroups As colSurfaceGroups

Private tstStrDXTX As String

Private Sub Class_Terminate()

    Set mvarSurfaceGroups = Nothing
    Set mvarSVGSurfaces = Nothing
    Set mvarSVGTeethDef = Nothing
    Set mvarSVGDxsTxs = Nothing
    Set mvarSVGLib = Nothing
    
End Sub
Public Property Set SurfaceGroups(ByVal vData As colSurfaceGroups)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.SurfaceGroups = Form1
    Set mvarSurfaceGroups = vData
End Property


Public Property Get SurfaceGroups() As colSurfaceGroups
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SurfaceGroups
    Set SurfaceGroups = mvarSurfaceGroups
End Property




Public Property Set SVGSurfaces(ByVal vData As colSVGSurfaces)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.SVGSurfaces = Form1
    Set mvarSVGSurfaces = vData
End Property


Public Property Get SVGSurfaces() As colSVGSurfaces
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SVGSurfaces
    Set SVGSurfaces = mvarSVGSurfaces
End Property




Public Property Set SVGTeethDef(ByVal vData As colSVGTeethDef)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.SVGTeethDef = Form1
    Set mvarSVGTeethDef = vData
End Property


Public Property Get SVGTeethDef() As colSVGTeethDef
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SVGTeethDef
    Set SVGTeethDef = mvarSVGTeethDef
End Property




Public Property Set SVGDxsTxs(ByVal vData As colSVGDxsTxs)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.SVGDxsTxs = Form1
    Set mvarSVGDxsTxs = vData
End Property


Public Property Get SVGDxsTxs() As colSVGDxsTxs
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SVGDxsTxs
    Set SVGDxsTxs = mvarSVGDxsTxs
End Property




Public Function getStrToothDefinition(ToothID As Integer) As String
       
Dim oSVGToothDef As clsSVGToothDef

    Set oSVGToothDef = mvarSVGTeethDef(CStr(ToothID))
    
    getStrToothDefinition = oSVGToothDef.getToothDefHeader & " " & _
                            getStrSurfaceDefinitionFromCol(ToothID) & " " & _
                            oSVGToothDef.getToothDefBody & " " & _
                            oSVGToothDef.getToothDefFooter
       
End Function
Private Function getStrSurfaceDefinitionFromCol(ToothID As Integer) As String
  
Dim oSVGSurface As clsSVGSurface

  For Each oSVGSurface In mvarSVGSurfaces
     If oSVGSurface.ToothID = ToothID Then
       If ToothID = 1 Then 'Testing DXTX
         getStrSurfaceDefinitionFromCol = getStrSurfaceDefinitionFromCol & " " & _
                                           oSVGSurface.getSurfaceDefHeader & " " & _
                                           oSVGSurface.getSurfaceDefBody & " " & _
                                          tstStrDXTX & " " & oSVGSurface.getSurfaceDefFooter
       'Exit For
       Else
         getStrSurfaceDefinitionFromCol = getStrSurfaceDefinitionFromCol & " " & _
                                           oSVGSurface.getSurfaceDefHeader & " " & _
                                           oSVGSurface.getSurfaceDefBody & " " & _
                                           oSVGSurface.getSurfaceDefFooter
       'Exit For
       End If
      End If
  Next
End Function
Friend Sub loadSVGToothDefinition()

   getSurfaceDefFromResource
   getTeethDefFromResource
   getDxTxToothDefFromResource
   getSurfaceMacros

End Sub
Private Sub getSurfaceMacros()

Dim i As Integer
Dim sizeOfString As String
Dim strCompressed As String
Dim strUniSurface As String
Dim strAnsiSurface As String
Dim rows() As String
Dim cols() As String
Dim SurfaceGroup As clsSurfaceGroup
    
    Set mvarSurfaceGroups = New colSurfaceGroups
    
    strUniSurface = LoadResData("TeethSurfaceInformation", "CUSTOM")
    strAnsiSurface = StrConv(strUniSurface, vbUnicode)
    sizeOfString = Trim(Left(strAnsiSurface, 10))
    strCompressed = Right(strAnsiSurface, (Len(strAnsiSurface) - 10))
    Call DecompressString(strCompressed, CLng(sizeOfString))
    rows = Split(strCompressed, "^")
    For i = 0 To UBound(rows)
        cols = Split(rows(i), "|")
        Set SurfaceGroup = New clsSurfaceGroup
        SurfaceGroup.ToothID = cols(0)
        SurfaceGroup.CrownSurfaces = cols(1)
        SurfaceGroup.OcclusalSurfaces = cols(2)
        SurfaceGroup.RootSurfaces = cols(3)
        SurfaceGroup.Root_XY = cols(4)
        SurfaceGroup.Crown_XY = cols(5)
        mvarSurfaceGroups.Add SurfaceGroup, CStr(SurfaceGroup.ToothID)
    Next i

End Sub
Private Sub getTeethDefFromResource()

Dim strUniTeethDef As String
Dim strAnsiTeethDef As String
Dim strArrTeethDef() As String
Dim strArrToothDef() As String
Dim i, j As Integer

Dim oSVGToothDef As clsSVGToothDef
    
    Set mvarSVGTeethDef = New colSVGTeethDef
    
    strUniTeethDef = LoadResData("TeethSVGDefinitions", "CUSTOM")
    strAnsiTeethDef = StrConv(strUniTeethDef, vbUnicode)
    strArrTeethDef = Split(strAnsiTeethDef, "~")
    For i = 0 To UBound(strArrTeethDef)
        strArrToothDef = Split(strArrTeethDef(i), "|")
        Set oSVGToothDef = New clsSVGToothDef
        oSVGToothDef.ToothID = CInt(strArrToothDef(0))
        oSVGToothDef.Header = strArrToothDef(1)
        oSVGToothDef.Footer = strArrToothDef(2)
        oSVGToothDef.CrownMD = strArrToothDef(3)
        oSVGToothDef.Roots = strArrToothDef(4)
        oSVGToothDef.Rect = strArrToothDef(5)
        oSVGToothDef.ToothName = strArrToothDef(6)
        mvarSVGTeethDef.Add oSVGToothDef, strArrToothDef(0)
    Next i
   
End Sub
Private Sub getSurfaceDefFromResource()

Dim strUniTeethDef As String
Dim strAnsiTeethDef As String
Dim strArrTeethDef() As String
Dim strArrToothDef() As String
Dim i, j As Integer

Dim oSVGSurface As clsSVGSurface
    
    Set mvarSVGSurfaces = New colSVGSurfaces
    
    tstStrDXTX = LoadResString(102)
    strUniTeethDef = LoadResData("SURFACESVGDEFINITIONS", "CUSTOM")
    strAnsiTeethDef = StrConv(strUniTeethDef, vbUnicode)
    strArrTeethDef = Split(strAnsiTeethDef, "~")
    For i = 0 To UBound(strArrTeethDef)
        strArrToothDef = Split(strArrTeethDef(i), "|")
        Set oSVGSurface = New clsSVGSurface
        oSVGSurface.ToothID = CInt(strArrToothDef(0))
        oSVGSurface.SurfaceID = strArrToothDef(1)
        oSVGSurface.Surface = strArrToothDef(2)
        oSVGSurface.ASurface = strArrToothDef(3)
        mvarSVGSurfaces.Add oSVGSurface, strArrToothDef(0) & "|" & strArrToothDef(1)
    Next i
   
End Sub
Private Sub getDxTxToothDefFromResource()

'Load the DXTXTOOTHSVGDEFINITIONS into the collection of
'colSVGDxsTxs, which stores the svg string for each Dx and Tx per tooth per surface

Dim strUniTeethDef As String
Dim strAnsiTeethDef As String
Dim strArrTeethDef() As String
Dim strArrToothDef() As String
Dim oSVGDxTx As clsSVGDxTx

Dim i, j As Integer

On Error Resume Next
    
    'Reset the colSVGDxsTxs
    Set mvarSVGDxsTxs = New colSVGDxsTxs
    
    strUniTeethDef = LoadResData("DXTXTOOTHSVGDEFINITIONS", "CUSTOM")
    strAnsiTeethDef = StrConv(strUniTeethDef, vbUnicode)
    strArrTeethDef = Split(strAnsiTeethDef, vbCrLf)
    
    For i = 0 To UBound(strArrTeethDef)
        strArrToothDef = Split(strArrTeethDef(i), "|")
        Set oSVGDxTx = New clsSVGDxTx
        oSVGDxTx.ToothID = CInt(strArrToothDef(0))
        oSVGDxTx.SurfaceID = CInt(strArrToothDef(1))
        oSVGDxTx.AliasID = CInt(strArrToothDef(2))
        oSVGDxTx.DxTxSVG = strArrToothDef(3)
        
        mvarSVGDxsTxs.Add oSVGDxTx, strArrToothDef(0) & "|" & strArrToothDef(1) & "|" & strArrToothDef(2)
    Next i
   
End Sub


