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
Attribute VB_Name = "clsSurface_Definitions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private surfaceDefAttr As clsSurface_Def_Attr
Private mcolSurfaceDef As Collection

Public Property Get surfDefAttr(ToothID As eToothId, surfaceID As Integer) As clsSurface_Def_Attr
Dim strKey As String
On Error GoTo errHandler

   If Not mcolSurfaceDef Is Nothing Then
'       Select Case ToothID
'          Case eToothId.e1_2_3_14_15_16_33_34_41_42
'            strKey = CStr(surfaceID) & "|" & "1|2|3|14|15|16|33|34|41|42"
'          Case eToothId.e6_7_8_9_10_11_22_23_24_25_26_27_35_36_37_38_39_40_45_46_47_48_49_50
'            strKey = CStr(surfaceID) & "|" & "6|7|8|9|10|11|22|23|24|25|26|27|35|36|37|38|39|40|45|46|47|48|49|50"
'                                             '6|7|8|9|10|11|22|23|24|25|26|27|35|36|37|38|39|40|45|46|47|48|49|50
'          Case eToothId.e4_5_12_13
'            strKey = CStr(surfaceID) & "|" & "4|5|12|13"
'          Case eToothId.e4_13s24
'            strKey = CStr(surfaceID) & "|" & "4|13"
'          Case eToothId.e5_12s24
'            strKey = CStr(surfaceID) & "|" & "5|12"
'          Case eToothId.e5_12s25
'            strKey = CStr(surfaceID) & "|" & "5|12"
'          Case eToothId.e17_19_30_32_43_52
'            strKey = CStr(surfaceID) & "|" & "17|19|30|32|43|52"
'          Case eToothId.e18_31_44_51
'            strKey = CStr(surfaceID) & "|" & "18|31|44|51"
'          Case eToothId.e20_21_28_29
'            strKey = CStr(surfaceID) & "|" & "20|21|28|29"
'       End Select
       strKey = CStr(surfaceID) & "|" & CStr(ToothID)
       Set surfDefAttr = mcolSurfaceDef(strKey)
       'mcolSurfaceDef.Add mcolSurfaceDef(141), strKey
       
   End If
   
Exit Property

errHandler:



  Err.Clear
  If surfaceID = 31 Then
    Set surfDefAttr = New clsSurface_Def_Attr
    surfDefAttr.surfaceName = "Mesial"
    surfDefAttr.surfaceAbbr = "M"
    surfDefAttr.surfaceID = 31
    surfDefAttr.ToothID = "All"
  ElseIf surfaceID = 32 Then
    Set surfDefAttr = New clsSurface_Def_Attr
    surfDefAttr.surfaceName = "Distal"
    surfDefAttr.surfaceAbbr = "D"
    surfDefAttr.surfaceID = 32
    surfDefAttr.ToothID = "All"
  Else
    Set surfDefAttr = New clsSurface_Def_Attr
    surfDefAttr.surfaceName = "Outline"
    surfDefAttr.surfaceAbbr = "Outline"
    surfDefAttr.surfaceID = 999
    surfDefAttr.ToothID = "Outline"
  End If
  
  'There is not a surface 31 or 32 so if it is requested then we do the above.
'  Call gobjShared.ShowVBError(Err.Number, Err.Description, "clsSurface_Definitions.surfDefAttr", "srvSVGComm", vbCritical)

End Property


Private Sub Class_Initialize()
   Set mcolSurfaceDef = New Collection
   loadSurfaceDef
End Sub
Private Sub Class_Terminate()
    
    Set mcolSurfaceDef = Nothing
    Set mcolSurfaceDef = Nothing

End Sub
Private Sub loadSurfaceDef()
  Dim strUniSurface As String
  Dim strAnsiSurface As String
  Dim strArrSurfaces() As String
  Dim strArrSurfAttr() As String
  Dim tmpStr As String
  Dim tmpToothIDEnum As Integer
  
  
  Dim i As Integer
  
  
   strUniSurface = LoadResData("Surface_Definitions", "CUSTOM")
   strAnsiSurface = StrConv(strUniSurface, vbUnicode)
   
   strArrSurfaces = Split(strAnsiSurface, "^")
   
   For i = 0 To UBound(strArrSurfaces)
          Set surfaceDefAttr = New clsSurface_Def_Attr
          strArrSurfAttr = Split(strArrSurfaces(i), ";")
          surfaceDefAttr.surfaceName = strArrSurfAttr(0)
            If surfaceDefAttr.surfaceName = "Lingaul Groove" Then surfaceDefAttr.surfaceName = "Lingual Groove"
          surfaceDefAttr.surfaceID = strArrSurfAttr(1)
          surfaceDefAttr.surfaceAbbr = strArrSurfAttr(2)
          surfaceDefAttr.ToothID = strArrSurfAttr(3)
          
          
          Select Case surfaceDefAttr.ToothID
          Case "1|2|3|14|15|16|33|34|41|42"
            tmpToothIDEnum = eToothId.e1_2_3_14_15_16_33_34_41_42
          Case "6|7|8|9|10|11|22|23|24|25|26|27|35|36|37|38|39|40|45|46|47|48|49|50"
            tmpToothIDEnum = eToothId.e6_7_8_9_10_11_22_23_24_25_26_27_35_36_37_38_39_40_45_46_47_48_49_50
          Case "4|5|12|13"
            tmpToothIDEnum = eToothId.e4_5_12_13
          Case "4|13"
            tmpToothIDEnum = eToothId.e4_13s24
          Case "5|12"
            If surfaceDefAttr.surfaceID = 24 Then
              tmpToothIDEnum = eToothId.e5_12s24
            Else
              tmpToothIDEnum = eToothId.e5_12s25
            End If
          Case "17|19|30|32|43|52"
            tmpToothIDEnum = eToothId.e17_19_30_32_43_52
          Case "18|31|44|51"
            tmpToothIDEnum = eToothId.e18_31_44_51
          Case "20|21|28|29"
            tmpToothIDEnum = eToothId.e20_21_28_29
       End Select
       
          tmpStr = CStr(surfaceDefAttr.surfaceID) & "|" & tmpToothIDEnum
          mcolSurfaceDef.Add surfaceDefAttr, tmpStr
   Next i
   Set surfaceDefAttr = Nothing
End Sub
