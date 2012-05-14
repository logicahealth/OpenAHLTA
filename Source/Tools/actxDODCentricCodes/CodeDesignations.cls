VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CodeDesignations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MODULE_NAME As String = "CodeDomainOps"

Private acCodeRanges(cdtRounds To cdtInvalidRNDs, cstEandM To cstProcedure) As Collection

Private Function AddCodeRange(ByVal eDesignation As CodeDesignationTypes, ByVal eCodeSet As CodeSetTypes, _
         ByVal sRangeStart As String, ByVal sRangeEnd As String)

Dim oCodeRange As CodeRange

   Set oCodeRange = New CodeRange
         
   oCodeRange.CodeDesignation = eDesignation
   oCodeRange.CodeSet = eCodeSet
   oCodeRange.Range_Start_Code = sRangeStart
   oCodeRange.Range_End_Code = sRangeEnd
         
   GetCodeRangeCollection(eDesignation, eCodeSet).Add oCodeRange
         
End Function
         
Friend Function IsCodeInDomain(ByVal eCodeDesignation As CodeDesignationTypes, _
         eCodeSet As CodeSetTypes, ByVal sCode As String) As Boolean
     
On Error GoTo ErrorHandler

Dim oCodeRange As CodeRange

   If Not IsLoaded(eCodeDesignation, eCodeSet) Then
      'Attempt to load code range
      If Not LoadCodeRanges(eCodeDesignation, eCodeSet) Then
         'Unable to load code ranges due to an error that was reported in LoadCodeRanges
         IsCodeInDomain = False
         Exit Function
      End If
   End If
     
   'Search the collection to determine if code is in the set
   For Each oCodeRange In GetCodeRangeCollection(eCodeDesignation, eCodeSet)
      If oCodeRange.CodeInRange(sCode) Then
         IsCodeInDomain = True
         Exit For
      End If
   Next
   
ErrorHandler:
   'Clean Up Objects
   Set oCodeRange = Nothing

   'Display Error Message
   If CBool(Err) Then HandleErrors Err, MODULE_NAME, "CodeDesignations.IsCodeInDomain"

End Function

Private Function GetCodeRangeCollection(ByVal eDesignationType As CodeDesignationTypes, eCodeSetType As CodeSetTypes) As Collection

Dim cCodeRanges As Collection

   If Not IsLoaded(eDesignationType, eCodeSetType) Then
      Set acCodeRanges(eDesignationType, eCodeSetType) = New Collection
   End If
   Set GetCodeRangeCollection = acCodeRanges(eDesignationType, eCodeSetType)
   
End Function

Private Function IsLoaded(ByVal eDesignationType As CodeDesignationTypes, eCodeSetType As CodeSetTypes) As Boolean
   IsLoaded = Not acCodeRanges(eDesignationType, eCodeSetType) Is Nothing
End Function

Private Function LoadCodeRanges(ByVal eCodeDesignation As CodeDesignationTypes, ByVal eCodeSet As CodeSetTypes) As Boolean

On Error GoTo ErrorHandler

Dim sSql As String
Dim adoRS As ADODB.Recordset
   
   'Build SQL Statement
   sSql = "SELECT Code_Range_Start, Code_Range_End " & _
            " FROM Code_Designation " & _
            " WHERE Code_Designation_Type_Id = " & eCodeDesignation & _
            "  AND Code_Set_Type_Id = " & eCodeSet
            
   Set adoRS = SharedDAS.OpenRecordset(sSql)
   
   If Not adoRS Is Nothing Then
            
      Do While Not adoRS.EOF
      
         AddCodeRange eCodeDesignation, eCodeSet, adoRS("Code_Range_Start"), adoRS("Code_Range_End")

         adoRS.MoveNext
         
      Loop
      
   End If
   
   LoadCodeRanges = True
   
ErrorHandler:
   'Clean Up Objects
   If Not adoRS Is Nothing Then
      If adoRS.State = adStateOpen Then adoRS.Close
      Set adoRS = Nothing
   End If
      
   'Handle Any Errors
   If CBool(Err) Then HandleErrors Err, MODULE_NAME, "LoadCodeRanges"
      
End Function

