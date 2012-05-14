VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ApptSearchParams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements iParameters
Implements iObjectPersist

'Private Const CLINNCID As String = "ClinicNCID"
'Private Const CLINIEN As String = "ClinicIEN"
'Private Const FACNCID As String = "FacilityNCID"
'Private Const USRNCID As String = "UserNCID"
'Private Const INCNONMTF As String = "IncludeNonMTF"
'Private Const INCWORK As String = "IncludeWorkload"
'Private Const APPTCOD As String = "ApptCode"
'Private Const INCUSRONLY As String = "IncludeUserClinicsOnly"
'Private Const UseCPTABLE As String = "UseClinicProviderTable"

Private mobjParams As iParameters

Private Sub Class_Initialize()

    On Error GoTo errHandler
    
    Set mobjParams = New ParametersGeneric
'    'Initialize ClinicSearchParams
'    SetValue CLINNCID, ""
'    SetValue CLINIEN, ""
'    SetValue FACNCID, ""
'    SetValue USRNCID, ""
'    SetValue INCNONMTF, True
'    SetValue INCWORK, False
'    SetValue APPTCOD, ""
'    SetValue INCUSRONLY, False
'    SetValue UseCPTABLE, True
    
    Exit Sub
    
errHandler:
    Err.Raise Err.Number, "ClinicSearchParam.Initialize", Err.Description, Err.HelpFile, Err.HelpContext
    
End Sub

Private Function iObjectPersist_Copy() As Object
Dim Persist As iObjectPersist

  Set Persist = mobjParams
  Set iObjectPersist_Copy = Persist.Copy

End Function

Private Sub iObjectPersist_Persist(Writer As ObjectPersist.iWriter)
Dim Persist As iObjectPersist

  Set Persist = mobjParams
  Persist.Persist Writer

End Sub

Private Function iObjectPersist_Restore(Reader As ObjectPersist.iReader) As Boolean
Dim Persist As iObjectPersist

  Set Persist = mobjParams
  iObjectPersist_Restore = Persist.Restore(Reader)
End Function


Private Property Get iObjectPersist_UniqueIdentifier() As String
Dim Persist As iObjectPersist

  Set Persist = mobjParams
  iObjectPersist_UniqueIdentifier = Persist.UniqueIdentifier
End Property

Private Property Get iObjectPersist_Version() As String
Dim Persist As iObjectPersist

  Set Persist = mobjParams
  iObjectPersist_Version = Persist.Version
End Property

Private Function iParameters_Add(ByVal Name As String, ByVal Value As Variant) As DL_Support.Parameter
    mobjParams.Add Name, Value
End Function

Private Function iParameters_Count() As Long
    iParameters_Count = mobjParams.Count
End Function

Private Function iParameters_Exists(ByVal Name As String) As Boolean
    iParameters_Exists = mobjParams.Exists(Name)
End Function

Private Function iParameters_Item(ByVal Key As Variant) As DL_Support.Parameter
    Set iParameters_Item = mobjParams.Item(Key)
End Function

Private Sub iParameters_Remove(ByVal Key As Variant)
    mobjParams.Remove (Key)
End Sub

