VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DMFTransactionData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_PatientInfo() As DMFPatientInfo
Private m_intDataset() As Long
Private m_intEncryption As DMCrypto
Private m_intFormat As DMFormats
Private m_intCompression As DMCompression
Private m_intSource As DMLocations
Private m_intDestination As DMLocations
Private m_strSourcePath As String
Private m_strDestinationPath As String
Private m_bIsEICDestination As Boolean
Private m_bIsEICSource As Boolean
Private m_lngRecID As Long
'<< SCR #30953;   Developer: Brian Mowbray 12/18/2002 03:04 PM
Private m_lEncSource As enEncSource


'<< SCR #30824;   Developer: Brian Mowbray 12/16/2002 11:50 PM
'Added lId
' SCR #44079;   Developer: Mag 10/20/2003 01:09 PM
'Added sFMP & sSponsorSSN
Public Type DMFPatientInfo
    sUnitNumber As String
    sSSN As String
    sName As String
    sId As String
    sFMP As String
    sSponsorSSN As String
End Type

Public Property Let Source(ByVal SourceEnum As DMLocations)

    m_intSource = SourceEnum

End Property

Public Property Get Source() As DMLocations

    Source = m_intSource
    
End Property

Public Property Let Destination(ByVal DestinationEnum As DMLocations)

    m_intDestination = DestinationEnum
    
End Property

Public Property Get Destination() As DMLocations

    Destination = m_intDestination
    
End Property

Public Property Let SourcePath(ByVal strPath As String)
    
    m_strSourcePath = strPath

End Property

Public Property Get SourcePath() As String
    
    SourcePath = m_strSourcePath
    
End Property

Public Property Let DestinationPath(ByVal strPath As String)
    
    m_strDestinationPath = strPath

End Property

Public Property Get DestinationPath() As String
    
    DestinationPath = m_strDestinationPath
    
End Property

Public Property Let PatientInfo(ByVal PatientInfoArray As Variant)

    m_PatientInfo = PatientInfoArray
        
End Property

Public Property Get PatientInfo() As Variant
    
    PatientInfo = m_PatientInfo
    
End Property

Public Property Let Dataset(ByVal DatasetArray As Variant)

    m_intDataset = DatasetArray

End Property

Public Property Get Dataset() As Variant

    Dataset = m_intDataset

End Property

Public Property Let Encryption(ByVal EncryptionEnum As DMCrypto)
    
    m_intEncryption = EncryptionEnum
    
End Property

Public Property Get Encryption() As DMCrypto

    Encryption = m_intEncryption
    
End Property

Public Property Let Format(ByVal FormatEnum As DMFormats)

    m_intFormat = FormatEnum
    
End Property

Public Property Get Format() As DMFormats
    
    Format = m_intFormat
    
End Property

Public Property Let Compression(ByVal CompressionEnum As DMCompression)

    m_intCompression = CompressionEnum
    
End Property

Public Property Get Compression() As DMCompression

    Compression = m_intCompression
    
End Property

Public Property Get IsEICDestination() As Boolean
    
    IsEICDestination = m_bIsEICDestination
    
End Property

Public Property Let IsEICDestination(ByVal bIsEIC As Boolean)

    m_bIsEICDestination = bIsEIC

End Property

Public Property Get IsEICSource() As Boolean
    
    IsEICSource = m_bIsEICSource
    
End Property

Public Property Let IsEICSource(ByVal bIsEIC As Boolean)

    m_bIsEICSource = bIsEIC

End Property




'<< SCR #30824;   Developer: Brian Mowbray 12/16/2002 11:50 PM
Public Property Get RecID() As Long
  RecID = m_lngRecID
End Property

'<< SCR #30824;   Developer: Brian Mowbray 12/16/2002 11:50 PM
Public Property Let RecID(ByVal lValueIn As Long)
  m_lngRecID = lValueIn
End Property

'<< SCR #30953;   Developer: Brian Mowbray 12/18/2002 03:04 PM
Public Property Get EncounterSource() As Long
  EncounterSource = m_lEncSource
End Property

'<< SCR #30953;   Developer: Brian Mowbray 12/18/2002 03:04 PM
Public Property Let EncounterSource(ByVal lValueIn As enEncSource)
  m_lEncSource = lValueIn
End Property
