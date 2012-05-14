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
Attribute VB_Name = "MTFTop50RankDrugs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'-------------------------------------------------------------------------------------
' The top 50 ranked drugs by MTF in "DRUG_RANKING" table need to be highlighted
' in all pharmacy tabular reports
'-------------------------------------------------------------------------------------
Public Function HighlightTop50MTFDrugs(ByRef objRS As ADODB.Recordset, _
                                       ByVal sFacilityNCID As String) As ADODB.Recordset
On Error GoTo ErrHandler
   
   Dim sStmt As String
   Dim oRSrank As ADODB.Recordset
   
   sStmt = "SELECT  DISTINCT DC.GC3_NAME" & vbNewLine & _
            "FROM DRUG_RANKING DR, DRUG_COST_DATA DC" & vbNewLine & _
            "WHERE DR.DRUG_NDC = DC.NDC AND" & vbNewLine & _
                "DR.FACILITY_NCID = " & sFacilityNCID & " AND" & vbNewLine & _
                "DC.GC3_NAME IS NOT NULL AND" & vbNewLine & _
                "DR.RANK < 51" & vbNewLine

   Set oRSrank = goDBDAS.OpenRecordset(sStmt)
   
   If oRSrank.RecordCount = 0 Then
      Set HighlightTop50MTFDrugs = objRS
      GoTo CleanUp
   End If
   
   Do Until objRS.EOF
      '--- Rest oRSrank
      oRSrank.Filter = ""
      oRSrank.MoveFirst
      
      oRSrank.Filter = "GC3_NAME = " & SQLQuote(objRS("GC3_NAME"))
       
      If Not oRSrank.EOF Then
          objRS("GC3_NAME") = "***" & objRS("GC3_NAME") & "***"
      End If
      
      objRS.MoveNext
   Loop
   
   Set HighlightTop50MTFDrugs = objRS
   
   GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "MTFTop50RankDrugs.HighlightTop50MTFDrugs", "CHCSII_Reports", vbExclamation
CleanUp:
    oRSrank.Close
    Set oRSrank = Nothing
    Exit Function
End Function

Public Function HighlightClinicDrugs(ByRef objRS As ADODB.Recordset, _
                                       ByVal sClinicNCID As String) As ADODB.Recordset
On Error GoTo ErrHandler
   
   Dim sStmt As String
   Dim oRSrank As ADODB.Recordset
   
   sStmt = "SELECT  DISTINCT DC.GC3_NAME" & vbNewLine & _
            "FROM DRUG_RANKING DR, DRUG_COST_DATA DC, CLINIC c" & vbNewLine & _
            "WHERE DR.DRUG_NDC = DC.NDC AND" & vbNewLine & _
                "DR.FACILITY_NCID = c.Facility_NCID AND" & vbNewLine & _
                "DC.GC3_NAME IS NOT NULL AND " & vbNewLine & _
                "DC.GC3_NAME IS NOT NULL AND " & vbNewLine & _
                "DR.RANK < 51 AND " & vbNewLine & _
                "c.NCID = " & sClinicNCID
                
   Set oRSrank = goDBDAS.OpenRecordset(sStmt)
   
   If oRSrank.RecordCount = 0 Then
      Set HighlightClinicDrugs = objRS
      GoTo CleanUp
   End If
   
   Do Until objRS.EOF
      '--- Rest oRSrank
      oRSrank.Filter = ""
      oRSrank.MoveFirst
      
      oRSrank.Filter = "GC3_NAME = " & SQLQuote(objRS("GC3_NAME"))
       
      If Not oRSrank.EOF Then
          objRS("GC3_NAME") = "***" & objRS("GC3_NAME") & "***"
      End If
      
      objRS.MoveNext
   Loop
   
   Set HighlightClinicDrugs = objRS
   
   GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "MTFTop50RankDrugs.HighlightClinicDrugs", "CHCSII_Reports", vbExclamation
CleanUp:
    oRSrank.Close
    Set oRSrank = Nothing

End Function

