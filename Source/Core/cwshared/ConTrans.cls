VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Library: Concept Translation Library
'
'       Class: ConTrans
'           Methods:
'                StoreNcid(ByVal sNcid As String, ByVal sContext As String) As Integer
'                Representation(ByVal sNcid As String, ByVal sContext As String) As String
'
'           Properties:
'                   none
'
'   Description:        Provides a wrapper for doing batch translations of Ncid/Context pairs.
'               StoreNcid should be called repeatedly to collect the Ncids to be translated
'               en masse.  Calling 'Representation' will cause all untranslated Ncids to be
'               translated.  Calling 'Representation' on an Ncid that has not previously been
'               stored will not be an error, but will cause an immediate tranlations.
'               If the specfied Ncid is NULL or 0, this will just translate to a "" string.  If
'               Context is 0, the default context will be used.  Context cannot be NULL.
'
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Private Const msConTrans As String = "ConceptTransaction"

Private msOldContext As String          ' storage for original context (which is presevered)
Private moConcept As ConceptCtrl     ' sole concept object used to access the cache
Private moContext As ContextCtrl        ' context object used to set default context


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   StoreNcid(ByVal sNcid As String, ByVal sContext As String) As Integer
'
'   Description:    present an Ncid/context pair to the cache
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub StoreNcid(ByVal sNcid As String, ByVal sContext As String)
Attribute StoreNcid.VB_Description = "Creates a Concept object with PrefRep.Needed set to TRUE, if it doesn't already exist."
    
    On Error GoTo ErrHandler

    ' protect us from evil
    If CheckParams(sNcid, sContext) = False Then
        Exit Sub
    End If
    
    
    If sContext <> "0" Then
        msOldContext = moContext.SetDefault(sContext)
    End If
    
    ' this will put the ncid/context in the cache ready for translation
    moConcept.UniqueId = sNcid
    
    If sContext <> "0" Then
        moContext.SetDefault (msOldContext)
    End If

    Exit Sub
ErrHandler:
    Call cwShowVBError(Err.Number, Err.Description, " ConTrans:StoreNcid: object = " + Err.Source, msConTrans, vbExclamation)
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Representation(ByVal sNcid As String, ByVal sContext As String) As String
'
'   Description:    Return the representation for the specified Ncid/Context pair, and
'               translate all pending Ncids.
'
'   Returns:  the Representation for the specfied Concept, or NULL if sNcid is NULL or 0
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Representation(ByVal sNcid As String, ByVal sContext As String) As String
    Dim i As Integer
    
    On Error GoTo ErrHandler
    
    Representation = ""
    
    If CheckParams(sNcid, sContext) = False Then
        Exit Function
    End If
   
   If sContext <> "0" Then
        msOldContext = moContext.SetDefault(sContext)
    End If
        
    ' get the representation
    moConcept.UniqueId = sNcid
    Representation = moConcept.PrefRep(sContext).Representation
    
    If sContext <> "0" Then
        moContext.SetDefault (msOldContext)
    End If
    
    Exit Function
ErrHandler:
    Call cwShowVBError(Err.Number, Err.Description, " ConTrans:Representation: object = " + Err.Source, msConTrans, vbExclamation)
End Function


Private Function CheckParams(sNcid As String, sContext As String) As Boolean
    
    CheckParams = False
    
    If sContext = "" Then
        MsgBox "Error: Context must be 0 or a valid context ncid."
        Exit Function
    End If
    
    ' don't error on the following, just return false
    If sNcid = "" Or Val(sNcid) <= 0 Then
        Exit Function
    End If
   
    
    ' do error on these - they gotta be confused
    If IsNumeric(sNcid) = False Then
        MsgBox "Error: Ncid '" + sNcid + "' is not numeric.", vbOKOnly + vbCritical, "Concept Cache"
        Exit Function
    End If
    If IsNumeric(sContext) = False Then
        MsgBox "Error: Context'" + sContext + "' is not numeric.", vbOKOnly + vbCritical, "Concept Cache"
        Exit Function
    End If
    
    CheckParams = True
    
End Function

Private Sub Class_Initialize()
    
    On Error GoTo ErrHandler
    
    Set moConcept = New ConceptCtrl
    Set moContext = New ContextCtrl
    
    Exit Sub
ErrHandler:
    Call cwShowVBError(Err.Number, Err.Description, " ConTrans:Representation: object = " + Err.Source, msConTrans, vbExclamation)
End Sub

Private Sub Class_Terminate()

    On Error GoTo ErrHandler
    
    Set moConcept = Nothing
    Set moContext = Nothing
    
    Exit Sub
ErrHandler:
    Call cwShowVBError(Err.Number, Err.Description, " ConTrans:Representation: object = " + Err.Source, msConTrans, vbExclamation)
End Sub
