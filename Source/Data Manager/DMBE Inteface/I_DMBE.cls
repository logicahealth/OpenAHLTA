VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "I_DMBE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Enum DMLocations
    eDM_PIC = 1
    eDM_TMIP = 2
    eDM_LDDB = 3
    eDM_CDR = 4
    eDM_AGGPIC = 5
    eDM_EIC = 6
    '<< Begin:SCR #29291;   Developer: Brian Mowbray 11/14/2002 12:53 PM
    eDM_BMIST = 7
    '>> End: SCR #29291;
End Enum

Public Enum DMFormats
    eDM_XML = 1
    eDM_ADT = 2
    eDM_DAT = 3
End Enum

Public Enum DMDataSets
    eDM_Demographics = 1
    eDM_Allergies = 2
    eDM_Problems = 3
    eDM_Medications = 4
    eDM_Readiness = 5
    eDM_Immunizations = 6
    eDM_Vitals = 7
    eDM_Encounters = 8
    eDM_ClinicNotes = 9
    eDM_Couplers = 10
    eDM_All = 11
End Enum

Public Enum DMCompression
   eDM_NoCompression = 88  ' "X"
   eDM_Compressed = 67     ' "C"
End Enum

Public Enum DMCrypto
   eDM_NoCrypto = 78    ' "N"
   eDM_Encrypted = 69   ' "E"
End Enum

Public Enum DMStatus
   eDM_Ready = 1
   eDM_NotReady = 2
   eDM_NotReadable = 3
   eDM_ReadyAndReadable = 4
End Enum

Public Enum TMIPDataRetrievalResult
    Success = 1
    Application_Error = 2
    No_Data = 4
    All_Data_Have_Been_Sent = 8
End Enum

Public Enum XMLMessageSource
    CHCSII_Message = 1
    TMIP_Message = 2
    SAMS_Message = 3
    '<< Begin:SCR #29291;   Developer: Brian Mowbray 11/14/2002 12:53 PM
    BMIST_Message = 4
    '>> End: SCR #29291;
End Enum

Public Enum enEncSource
    enEncounterSigned = 1
    enDM = 2
End Enum

Public Property Let DataSet(RHV As DMDataSets): End Property
Public Property Get DataSet() As DMDataSets: End Property

Public Property Set PatientList(RHV As Collection): End Property
Public Property Get PatientList() As Collection: End Property

Public Property Let Destination(RHV As DMLocations): End Property
Public Property Get Destination() As DMLocations: End Property

Public Property Let Source(RHV As DMLocations): End Property
Public Property Get Source() As DMLocations: End Property

Public Property Let PicPath(RHV As String): End Property
Public Property Get PicPath() As String: End Property

Public Property Let AltPicPath(RHV As String): End Property
Public Property Get AltPicPath() As String: End Property

Public Property Let Format(RHV As DMFormats): End Property
Public Property Get Format() As DMFormats: End Property

Public Property Let Compression(RHV As DMCompression): End Property
Public Property Get Compression() As DMCompression: End Property

Public Property Let Crypto(RHV As DMCrypto): End Property
Public Property Get Crypto() As DMCrypto: End Property

Public Property Let UpdateLock(RHV As Boolean): End Property
Public Property Get UpdateLock() As Boolean: End Property

Public Property Get PatientObj() As Object: End Property
Public Property Set PatientObj(RHV As Object): End Property

Public Property Get Password() As String: End Property


Public Function Execute(cObjs As Collection, Optional ByRef TMIPDataRetrievalResult As TMIPDataRetrievalResult) As Long:

End Function


Public Function ClearParameters() As Boolean: End Function

Public Function InitializePIC() As Long: End Function

Public Function ValidatePIC() As Long: End Function

Public Function InitDMBE(eSource As DMLocations, eDest As DMLocations, eFormat As DMFormats, eDataSet As DMDataSets, eUnitNumber As Collection, PicPath As String, Optional AltPicPath As String, Optional eCompress As DMCompression, Optional eCrypto As DMCrypto) As Long: End Function

Public Function CreateDetachedRS(TableName As String, objRS As ADODB.Recordset, Optional ByVal OpenRS As Boolean = True) As Long: End Function

Public Function GetDMBEErrorMessage(ErrorCode As Long) As String: End Function

Public Function QueryPIC(PathToPIC As String, RetVal As Long) As Object: End Function

Public Function CopyPIC() As Long

End Function

'Public Property Get IsSAMSTransaction() As Boolean: End Property
'Public Property Let IsSAMSTransaction(ByVal rhv As Boolean): End Property
'
'Public Property Get SAMSHL7Message() As String: End Property
'Public Property Let SAMSHL7Message(strHL7 As String): End Property
'
'Public Function WriteToSAMS(ByVal strSSN As String) As Long
'
'End Function
'
'Public Property Get SAMSIsDestination() As Boolean: End Property
'Public Property Let SAMSIsDestination(ByVal rhv As Boolean): End Property
'
'Public Property Get SAMSIsSource() As Boolean: End Property
'Public Property Let SAMSIsSource(ByVal rhv As Boolean): End Property


'Get rid of the SAMS Stuff after this build

Public Property Get IsEICTransaction() As Boolean: End Property
Public Property Let IsEICTransaction(ByVal RHV As Boolean): End Property

Public Property Get EICHL7Message() As String: End Property
Public Property Let EICHL7Message(RHV As String): End Property

Public Function WriteBMISTFile(ByVal enmDMFormat As DMFormats, ByVal strFilePath As String) As Long

End Function

Public Function WriteToEIC(ByVal RHV As String) As Long

End Function

Public Function WriteToPIC(ByVal RHV As String, dicDatasets As Object) As Long

End Function

Public Property Get EICIsDestination() As Boolean: End Property
Public Property Let EICIsDestination(ByVal RHV As Boolean): End Property

Public Property Get EICIsSource() As Boolean: End Property
Public Property Let EICIsSource(ByVal RHV As Boolean): End Property

'<SCR 24303 CC>
'Public Property Get IDMappers() As Object: End Property
'Public Property Set IDMappers(ByRef RHV As Object):  End Property

Public Property Get IDMappers() As I_IDMappers: End Property
Public Property Set IDMappers(ByRef RHV As I_IDMappers):  End Property


Public Property Set LogonObj(RHV As Object)

End Property

Public Property Get WhoseXMLMessage() As XMLMessageSource

End Property

Public Property Let WhoseXMLMessage(ByVal RHV As XMLMessageSource)

End Property

Public Property Get CurrentDMErrorCode() As Long

End Property

Public Function AppendTransferRS(ByVal TableName As String, ByRef RS As ADODB.Recordset) As Boolean
    
End Function

Public Function ClearTransferRS() As Boolean

End Function

Public Property Get PICPatientUnitNumber() As Long

End Property

Public Property Let PICPatientUnitNumber(ByVal RHV As Long)

End Property

Public Property Get FileName() As String

End Property

Public Property Let FileName(ByVal RHV As String)

End Property

'<< SCR #30953;   Developer: Brian Mowbray 12/18/2002 03:04 PM
Public Property Get EncounterSource() As Long

End Property

'<< SCR #30953;   Developer: Brian Mowbray 12/18/2002 03:04 PM
Public Property Let EncounterSource(ByVal vValueIn As enEncSource)

End Property
