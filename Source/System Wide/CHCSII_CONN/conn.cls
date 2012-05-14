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
Attribute VB_Name = "Conn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarUseLocalDB As Boolean 'Oracle
Private mvarUseSQLServer As Boolean
Private mvarUseMySQL As Boolean
Private mvarAppMode As enumCSCSIIAppMode

'The enum values represent the SQL Server codes for DateTime conversions
Public Enum eDateFormat
    DD_MON_YYYY_HH_MM_SS = 113
    DD_MON_YYYY = 106
    MM_DD_YYYY_HH_MM_SS = 100
    MM_DD_YYYY = 101
End Enum

Public Enum eBackend
    Auto = 0
    CDR = 1
    LGS = 2
    GEMS = 4
    ITT = 8
    ICD = 9
    MYSQL = 10
End Enum

'Allow the user to specify a service name
Public Enum eServiceNameMode
    Default = 0
    Reports = 1
End Enum

Public Enum eTimeZoneConversion
    NoTimeZoneConversion = 0
    YesTimeZoneConversion = 1
    BadTimeZoneConversion = 2
End Enum

Private Const cICDConnectionString As String = "Provider=OraOLEDB.Oracle.1;User Id=ico;Password=ico;Data Source=PGUI-CHCSII"
'Private Const cICDSPConnectionString As String = "Provider=MSDAORA;User Id=ico;Password=ico;Data Source=PGUI-CHCSII"
'Private Const cICDConnectionString As String = "filedsn=c:\dsn\msftodbcora.dsn"
'Private Const cICDConnectionString As String = "filedsn=c:\dsn\oracleodbc.dsn"

Public Sub AuditSQL(TableName As String, SqlStatement As String, Cnn As Connection)

'SCR 89413 Allow external object to audit their sql statements
'Primarily used for .addnew methods.
'The calling function contains the error handling, this is a pass through method.

  Call TMIPAudit("C2T_DAS", "GEMS_DAS.ExecuteSQL", CurrentUserName, TableName, SqlStatement, Err.Description, Cnn)

End Sub


Public Function SQLQuote(ByVal sTxt As String, Optional MaxLength As Long = 0, Optional InsertNulls As Boolean = True, Optional AddSingleQuotes As Boolean = True) As String

On Error GoTo ErrHnd

    SQLQuote = modGEMS_DAS.SQLQuote(sTxt, MaxLength, InsertNulls, AddSingleQuotes)

Exit Function

ErrHnd:
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Function
    Resume
End Function

Public Function SQLDate(ByVal dDate As Date, Optional DateFormat As eDateFormat = DD_MON_YYYY_HH_MM_SS) As String

    SQLDate = modGEMS_DAS.SQLDate(dDate, DateFormat)
    
End Function

'Added to Support Data Layer.  Phil Crowder 11/25/03
Public Sub InitDefaultMode(ByVal Mode As eBackend)
  
On Error GoTo ErrHnd

  Select Case Mode
    Case LGS
      mvarAppMode = modeCHCSI_GUI
    Case GEMS, MYSQL
      mvarAppMode = modeTheater
    Case ITT
      mvarAppMode = modeITT
    Case Else
      mvarAppMode = modeCDR
  End Select
  
Exit Sub

ErrHnd:
    Err.Raise Err.Number, Err.Source, Err.Description
  
End Sub

Public Function CHCSII_SQLOPS_EX(Optional Mode As eBackend = eBackend.Auto, Optional ServiceNameMode As eServiceNameMode = Default) As ISqlOpsEx

    '7/14/03  Phil Crowder changed Auto select to use AppMode to determine the
    'correct SQLOps object because UseLocalDB and UseSQLServer do not define the
    'data access completely (ICD moved to Oracle).
    
On Error GoTo ErrHnd

    Select Case Mode
    
        Case eBackend.Auto
            'Figure it out based on mvarUseLocalDB and mvarUseSQLServer
            If mvarAppMode And modeITT Then
                'Return ITT_SQLOPS
                Set CHCSII_SQLOPS_EX = New ITT_SQLOPS
                
            ElseIf mvarAppMode And modeCDR Then
                'Return CDR_SQLOPS
                Err.Raise 1001, "CHCSII_CONN", "CHCSII_SQLOPS_EX is not supported in CDR mode."
                
            ElseIf mvarAppMode And modeTheater Then
                If mvarUseMySQL Then
                    Set CHCSII_SQLOPS_EX = New MYSQL_SQLOPS
                Else
                    'Return GEMS_SQLOPS
                    Set CHCSII_SQLOPS_EX = New GEMS_SQLOPS
                End If
                
            ElseIf mvarAppMode And modeCHCSI_GUI Then
                'Return LGS_SQLOPS
                Err.Raise 1001, "CHCSII_CONN", "CHCSII_SQLOPS_EX is not supported in LGS mode."
                
            Else
                Err.Raise 1001, "CHCSII_CONN", "Cannot automatically determine which mode the app is in. You must specify which mode you want to use."
            
            End If
        
        Case eBackend.CDR
            Err.Raise 1001, "CHCSII_CONN", "CHCSII_SQLOPS_EX is not supported in CDR mode."
        
        Case eBackend.GEMS
            Set CHCSII_SQLOPS_EX = New GEMS_SQLOPS
            
        Case eBackend.MYSQL
            Set CHCSII_SQLOPS_EX = New MYSQL_DAS
        
            
        Case eBackend.ITT
            'Return ITT_SQLOPS
            Set CHCSII_SQLOPS_EX = New ITT_SQLOPS
            
        Case eBackend.LGS
            'Return LGS_SQLOPS
            Err.Raise 1001, "CHCSII_CONN", "CHCSII_SQLOPS_EX is not supported in LGS mode."
            
        Case eBackend.ICD
            'SCR 48849
            '3-8-04 SF Added ICD warning
            Err.Raise 1001, "CHCSII_CONN", "CHCSII_SQLOPS_EX is not supported in ICD mode."
            
    End Select

Exit Function

ErrHnd:
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Public Function FormatDateForDBInsert(sDate As Variant, Optional DateFormat As eDateFormat = DD_MON_YYYY_HH_MM_SS) As String
                                    'This is a variant so it will accept strings types and date types
    'Convert the date to the correct date format string for the appropriate database (SQL Server or Oracle)
    
On Error GoTo ErrHnd

    If IsDate(sDate) = False And Not sDate = Empty Then
        Err.Raise 1002, "CHCSIIConn.Conn.ConvertDate", "Cannot convert " & sDate & " to a date because it is not a valid date"
    Else
        If IsNull(sDate) Or sDate = Empty Then
            FormatDateForDBInsert = "NULL"
        Else
            FormatDateForDBInsert = SQLDate(CDate(sDate), DateFormat)
        End If
    End If
    
Exit Function

ErrHnd:
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function


Public Function FormatDateForDBSelect(FieldName As String, DateFormat As eDateFormat, Optional TruncateTime As Boolean = False) As String
    'Formats a date to be used in a select statement for MySQL or Oracle. You may specify the date format.
    'And you may specify if you want to truncate the time.
   
    Dim sTempField As String
    
On Error GoTo ErrHnd
    
    If mvarUseMySQL Then
        If TruncateTime Then
            sTempField = "DATE(" & FieldName & ", )"
        Else
            sTempField = FieldName
        End If
        
        'MySQL Select                                                                                     'Use the DB format
        FormatDateForDBSelect = "DATE_FORMAT(" & sTempField & ", '" & InternalDateFormatToString(DateFormat, False) & "')"
    ElseIf mvarUseSQLServer = True Then
        If TruncateTime Then
            sTempField = "CAST(FLOOR(CAST(" & FieldName & " AS FLOAT)) AS DATETIME)"
        Else
            sTempField = FieldName
        End If
        
        FormatDateForDBSelect = "CONVERT(VARCHAR, " & sTempField & "," & DateFormat & ")"
    Else
        If TruncateTime Then
            sTempField = "TRUNC(" & FieldName & ")"
        Else
            sTempField = FieldName
        End If
        
        'Oracle Select                                                                                     'Use the DB format
        FormatDateForDBSelect = "TO_CHAR(" & sTempField & ", '" & InternalDateFormatToString(DateFormat, False) & "')"
    End If

Exit Function

ErrHnd:
    Err.Raise Err.Number, Err.Source, Err.Description

End Function


Public Function FormatDisplayDate(sDate As Variant, Optional DateFormat As eDateFormat = DD_MON_YYYY_HH_MM_SS) As String
                                'This is a variant so it will accept string types and date types
'Convert the date to the correct format for output to the screen

On Error GoTo ErrHnd

    If IsNull(sDate) Then
        FormatDisplayDate = Empty
    ElseIf IsDate(sDate) = False Then
        Err.Raise 1002, "CHCSIIConn.Conn.ConvertToDisplayDate", "Cannot convert " & sDate & " to a date because it is not a valid date"
    Else                                                                            'Always use the VB format
        FormatDisplayDate = Format(sDate, InternalDateFormatToString(DateFormat, True))
    End If

Exit Function

ErrHnd:
    Err.Raise Err.Number, Err.Source, Err.Description

End Function


Public Function DateFormatEnumToString(DateFormat As eDateFormat) As String
    
On Error GoTo ErrHnd

    DateFormatEnumToString = InternalDateFormatToString(DateFormat, mvarUseSQLServer)
    
Exit Function

ErrHnd:
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function


Private Function InternalDateFormatToString(DateFormat As eDateFormat, UseVBFormat As Boolean) As String

    'Converts the eDateFormat enum to a string based on UseVBFormat
    
On Error GoTo ErrHnd

    Select Case DateFormat
        Case eDateFormat.DD_MON_YYYY_HH_MM_SS
            If UseVBFormat Then
                InternalDateFormatToString = DATE_FORMAT_VB_DD_MON_YYYY_HH_MM_SS
            Else
                If mvarUseMySQL Then
                    InternalDateFormatToString = DATE_FORMAT_MYSQL_DD_MON_YYYY_HH24_MI_SS
                Else
                    InternalDateFormatToString = DATE_FORMAT_ORA_DD_MON_YYYY_HH24_MI_SS
                End If
            End If
        
        Case eDateFormat.DD_MON_YYYY
            If UseVBFormat Then
                InternalDateFormatToString = DATE_FORMAT_VB_DD_MON_YYYY
            Else
                If mvarUseMySQL Then
                    InternalDateFormatToString = DATE_FORMAT_MYSQL_DD_MON_YYYY
                Else
                    InternalDateFormatToString = DATE_FORMAT_ORA_DD_MON_YYYY
                End If
            End If
        
        Case eDateFormat.MM_DD_YYYY_HH_MM_SS
            If UseVBFormat Then
                InternalDateFormatToString = DATE_FORMAT_VB_MM_DD_YYYY_HH_MM_SS
            Else
                If mvarUseMySQL Then
                    InternalDateFormatToString = DATE_FORMAT_MYSQL_MM_DD_YYYY_HH24_MM_SS
                Else
                    InternalDateFormatToString = DATE_FORMAT_ORA_MM_DD_YYYY_HH24_MM_SS
                End If
            End If
        
        Case eDateFormat.MM_DD_YYYY
            If UseVBFormat Then
                InternalDateFormatToString = DATE_FORMAT_VB_MM_DD_YYYY
            Else
                If mvarUseMySQL Then
                    InternalDateFormatToString = DATE_FORMAT_MYSQL_MM_DD_YYYY
                Else
                    InternalDateFormatToString = DATE_FORMAT_ORA_MM_DD_YYYY
                End If
            End If
            
    End Select

Exit Function

ErrHnd:
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Public Property Let UseSQLServer(ByVal vData As Boolean)
    
    'Only allow the property to be set once
    Static bOnce As Boolean
    
On Error GoTo ErrHnd

    If bOnce Then
        Err.Raise 1001, "CHCSII_CONN", "UseSQLServer can only be set once per instance."
        Exit Property
    End If
    
    bOnce = True
    
    mvarUseSQLServer = vData
    
Exit Property

ErrHnd:
    Err.Raise Err.Number, Err.Source, Err.Description

End Property


Public Property Get UseSQLServer() As Boolean
    
    UseSQLServer = mvarUseSQLServer

End Property



Public Property Let UseLocalDB(ByVal vData As Boolean)
    
    'Only allow the property to be set once
    Static bOnce As Boolean
    
On Error GoTo ErrHnd

    If bOnce Then
        Err.Raise 1001, "CHCSII_CONN", "UseLocalDB can only be set once per instance."
        Exit Property
    End If
    
    bOnce = True
    
    mvarUseLocalDB = vData
    
Exit Property

ErrHnd:
    Err.Raise Err.Number, Err.Source, Err.Description

End Property


Public Property Let UseMySQL(ByVal vData As Boolean)
    
    'Only allow the property to be set once
    Static bOnce As Boolean
    
On Error GoTo ErrHnd

    If bOnce Then
        Err.Raise 1001, "CHCSII_CONN", "UseMySQLDB can only be set once per instance."
        Exit Property
    End If
    
    bOnce = True
    
    mvarUseMySQL = vData
    
Exit Property

ErrHnd:
    Err.Raise Err.Number, Err.Source, Err.Description

End Property



Public Property Get UseLocalDB() As Boolean
    
    UseLocalDB = mvarUseLocalDB

End Property



Public Property Get UseMySQL() As Boolean
    
    UseMySQL = mvarUseMySQL

End Property




Private Sub Class_Initialize()

'Init the property defaults

On Error GoTo ErrHnd

Dim oCWShared As MMMHISShared.CWShared
        
    Set oCWShared = New MMMHISShared.CWShared
    
    mvarUseLocalDB = oCWShared.UseLocalDB
    mvarUseSQLServer = oCWShared.UseSQLServer
    mvarUseMySQL = oCWShared.UseMySQLDB
    
    If Not oCWShared.IsAppMode(modeUnknown) Then
      mvarAppMode = oCWShared.AppMode
    Else
      mvarAppMode = modeCDR     'Default to CDR mode for Mid-Tier process
    End If
    
    Set oCWShared = Nothing
    
Exit Sub

ErrHnd:
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Public Function CHCSII_DAS(Optional Mode As eBackend = eBackend.Auto, Optional TimeZoneConversion As eTimeZoneConversion = BadTimeZoneConversion, Optional ServiceNameMode As eServiceNameMode = Default) As ICHCSII_DataAccess.ICHCSII_DAS
    
    '7/14/03  Phil Crowder changed Auto select to use AppMode to determine the
    'correct DAS object because UseLocalDB and UseSQLServer do not define the
    'data access completely (ICD moved to Oracle).
    
    'SCR 46007
    'Developer: Scott Foster 12-6-03
    'Param created to expose the service name without modifying ICHCSII_DAS only valid for CDR mode.
    'TODO: implement the ServiceNameMode

On Error GoTo ErrHnd

Dim oCDRDAS As CDR_DAS
Dim oICDDAS As ICD_DAS

    Select Case Mode
    
        Case eBackend.Auto
            'Figure it out based on mvarUseLocalDB and mvarUseSQLServer
            If mvarAppMode And modeITT Then
                'Return ITT DAS
                Set CHCSII_DAS = New ITT_DAS
            
            ElseIf mvarAppMode And modeCDR Then
                Set oCDRDAS = New CDR_DAS
                'Set the time zone conversion property and the service name
                oCDRDAS.Init TimeZoneConversion, ServiceNameMode
                Set CHCSII_DAS = oCDRDAS
            ElseIf mvarAppMode And modeTheater Then
                If mvarUseMySQL Then
                    Set CHCSII_DAS = New MYSQL_DAS
                Else
                    Set CHCSII_DAS = New GEMS_DAS
                End If
            ElseIf mvarAppMode And modeCHCSI_GUI Then
                Set CHCSII_DAS = New LGS_DAS
                
            Else
                Err.Raise 1001, "CHCSII_CONN", "Cannot automatically determine which mode the app is in. You must specify which mode you want to use."
            
            End If
        
        Case eBackend.CDR
                Set oCDRDAS = New CDR_DAS
                
                'Set the time zone conversion property and the service name
                oCDRDAS.Init TimeZoneConversion, ServiceNameMode
                
                'Return CDR_DAS
                Set CHCSII_DAS = oCDRDAS
        
        Case eBackend.GEMS
            Set CHCSII_DAS = New GEMS_DAS
            
        Case eBackend.MYSQL
            Set CHCSII_DAS = New MYSQL_DAS
        
        Case eBackend.ITT
            'Return ITT DAs
            Set CHCSII_DAS = New ITT_DAS
            
        Case eBackend.LGS
            'Return LGS DAS
            Set CHCSII_DAS = New LGS_DAS
        
        Case eBackend.ICD
            'SCR 48849
            '3-8-04 SF Added ICD support
            Set CHCSII_DAS = New ICD_DAS
            Set oICDDAS = CHCSII_DAS
            oICDDAS.OpenDatabase cICDConnectionString
            Set oICDDAS = Nothing
            
    End Select
    
    Set oCDRDAS = Nothing
    
Exit Function

ErrHnd:
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Public Function ADO_DAS(Optional Mode As eBackend = eBackend.Auto, Optional TimeZoneConversion As eTimeZoneConversion = BadTimeZoneConversion, Optional ServiceNameMode As eServiceNameMode = Default) As IADO_DAS

    '7/14/03  Phil Crowder changed Auto select to use AppMode to determine the
    'correct DAS object because UseLocalDB and UseSQLServer do not define the
    'data access completely (ICD moved to Oracle).

    'TODO: implement the ServiceNameMode

On Error GoTo ErrHnd

Dim oCDRDAS As CDR_DAS
Dim oICDDAS As ICD_DAS

    Select Case Mode

        Case eBackend.Auto
            'Figure it out based on mvarUseLocalDB and mvarUseSQLServer
            If mvarAppMode And modeITT Then
                'Return ITT DAS
                Set ADO_DAS = New ITT_DAS

            ElseIf mvarAppMode And modeCDR Then
                Set oCDRDAS = New CDR_DAS
                oCDRDAS.Init TimeZoneConversion, ServiceNameMode
                Set ADO_DAS = oCDRDAS
            ElseIf mvarAppMode And modeTheater Then
                If mvarUseMySQL Then
                    Set ADO_DAS = New MYSQL_DAS
                Else
                    Set ADO_DAS = New GEMS_DAS
                End If
            ElseIf mvarAppMode And modeCHCSI_GUI Then
                Err.Raise 1001, "CHCSII_CONN", "ADO_DAS does not work in LGS mode yet."
                'Set ADO_DAS = New LGS_DAS

            Else
                Err.Raise 1001, "CHCSII_CONN", "Cannot automatically determine which mode the app is in. You must specify which mode you want to use."

            End If

        Case eBackend.CDR
                Set oCDRDAS = New CDR_DAS

                'SF 1-22-04
                'Set the time zone conversion property and the service name
                oCDRDAS.Init TimeZoneConversion, ServiceNameMode

                'Return CDR_DAS
                Set ADO_DAS = oCDRDAS

        Case eBackend.GEMS
            Set ADO_DAS = New GEMS_DAS
            
        Case eBackend.MYSQL
            Set ADO_DAS = New MYSQL_DAS

        Case eBackend.ITT
            'Return ITT DAs
            Set ADO_DAS = New ITT_DAS

        Case eBackend.LGS
            'Return LGS DAS
            Err.Raise 1001, "CHCSII_CONN", "ADO_DAS does not work in LGS mode yet."
            'Set ADO_DAS = New LGS_DAS

        Case eBackend.ICD
            'SCR 48849
            '3-8-04 SF Added ICD support
            'Err.Raise 1001, "CHCSII_CONN", "ADO_DAS does not work in ICD mode yet."
            Set ADO_DAS = New ICD_DAS
            Set oICDDAS = ADO_DAS
            oICDDAS.OpenDatabase cICDConnectionString, ""

    End Select

    Set oICDDAS = Nothing
    Set oCDRDAS = Nothing

Exit Function

ErrHnd:
    Err.Raise Err.Number, Err.Source, Err.Description

End Function


Public Function CHCSII_SQLOPS(Optional Mode As eBackend = eBackend.Auto, Optional TimeZoneConversion As eTimeZoneConversion = BadTimeZoneConversion, Optional ServiceNameMode As eServiceNameMode = Default) As ICHCSII_SQL.ISqlOps

    '7/14/03  Phil Crowder changed Auto select to use AppMode to determine the
    'correct SQLOps object because UseLocalDB and UseSQLServer do not define the
    'data access completely (ICD moved to Oracle).
   
'SF 1-22-04 Added the time zone conversion parameter

On Error GoTo ErrHnd

Dim oCdrSqlOps As CDR_SQLOPS
Dim oICDSqlOps As ICD_SQLOPS

    Select Case Mode
    
        Case eBackend.Auto
            'Figure it out based on mvarUseLocalDB and mvarUseSQLServer
            If mvarAppMode And modeITT Then
                'Return ITT_SQLOPS
                Set CHCSII_SQLOPS = New ITT_SQLOPS
                
            ElseIf mvarAppMode And modeCDR Then
                Set oCdrSqlOps = New CDR_SQLOPS
                
                '1/22/04 Set the time zone coversion property and the service name mode
                oCdrSqlOps.Init TimeZoneConversion, ServiceNameMode
                
                'Return CDR_SQLOPS
                Set CHCSII_SQLOPS = oCdrSqlOps
                
            ElseIf mvarAppMode And modeTheater Then
                If mvarUseMySQL Then
                    Set CHCSII_SQLOPS = New MYSQL_SQLOPS
                Else
                    'Return GEMS_SQLOPS
                    Set CHCSII_SQLOPS = New GEMS_SQLOPS
                End If
            ElseIf mvarAppMode And modeCHCSI_GUI Then
                'Return LGS_SQLOPS
                Set CHCSII_SQLOPS = New LGS_SQLOPS
            
            Else
                Err.Raise 1001, "CHCSII_CONN", "Cannot automatically determine which mode the app is in. You must specify which mode you want to use."
                
            End If
        
        Case eBackend.CDR
            Set oCdrSqlOps = New CDR_SQLOPS
            
            '1/22/04 Set the time zone coversion property and the service name mode
            oCdrSqlOps.Init TimeZoneConversion, ServiceNameMode
            Set CHCSII_SQLOPS = oCdrSqlOps
        
        Case eBackend.GEMS
            Set CHCSII_SQLOPS = New GEMS_SQLOPS
        
        Case eBackend.MYSQL
            Set CHCSII_SQLOPS = New MYSQL_SQLOPS
       
        
        Case eBackend.ITT
            'Return ITT_SQLOPS
            Set CHCSII_SQLOPS = New ITT_SQLOPS
            
        Case eBackend.LGS
            'Return LGS_SQLOPS
            Set CHCSII_SQLOPS = New LGS_SQLOPS
            
        Case eBackend.ICD
            'SCR 48849
            '3-8-04 SF Added ICD support
            Set CHCSII_SQLOPS = New ICD_SQLOPS
            Set oICDSqlOps = CHCSII_SQLOPS
            oICDSqlOps.OpenDatabase cICDConnectionString
            Set oICDSqlOps = Nothing
            
    End Select

    Set oCdrSqlOps = Nothing
    
Exit Function

ErrHnd:
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Public Function CHCSII_SQLOPS_SP(Optional Mode As eBackend = eBackend.Auto, Optional TimeZoneConversion As eTimeZoneConversion = BadTimeZoneConversion, Optional ServiceNameMode As eServiceNameMode = Default) As ICHCSII_SQL.ISqlOps

    '7/14/03  Phil Crowder changed Auto select to use AppMode to determine the
    'correct SQLOPS_SP object because UseLocalDB and UseSQLServer do not define the
    'data access completely (ICD moved to Oracle).
    
On Error GoTo ErrHnd

Dim oSQLOpsSP As CDR_SQLOPS_SP

    Select Case Mode
    
        Case eBackend.Auto
            'Figure it out based on mvarUseLocalDB and mvarUseSQLServer
            If mvarAppMode And modeITT Then
                'Return ITT_SQLOPS
                Err.Raise 1001, "CHCSII_CONN", "CHCSII_SQLOPS_SP is not supported in ITT mode."
                
            ElseIf mvarAppMode And modeCDR Then
                'Return CDR_SQLOPS
                Set oSQLOpsSP = New CDR_SQLOPS_SP
                
                'Set the time zone conversion and the service name mode
                oSQLOpsSP.Init TimeZoneConversion, ServiceNameMode
                
                Set CHCSII_SQLOPS_SP = oSQLOpsSP
                
            ElseIf mvarAppMode And modeTheater Then
                'Return GEMS_SQLOPS
                Err.Raise 1001, "CHCSII_CONN", "CHCSII_SQLOPS_SP is not supported in GEMS mode."
                                
            ElseIf mvarAppMode And modeCHCSI_GUI Then
                'Return LGS_SQLOPS
                Err.Raise 1001, "CHCSII_CONN", "CHCSII_SQLOPS_SP is not supported in LGS mode."
                
            Else
                Err.Raise 1001, "CHCSII_CONN", "Cannot automatically determine which mode the app is in. You must specify which mode you want to use."
            
            End If
        
        Case eBackend.CDR
            'Return CDR_SQLOPS
            oSQLOpsSP = New CDR_SQLOPS_SP
            
            'Set the time zone conversion and the service name mode
            oSQLOpsSP.Init TimeZoneConversion, ServiceNameMode
            
            Set CHCSII_SQLOPS_SP = oSQLOpsSP
        
        Case eBackend.GEMS
            Err.Raise 1001, "CHCSII_CONN", "CHCSII_SQLOPS_SP is not supported in GEMS mode."
            
        Case eBackend.MYSQL
            Err.Raise 1001, "CHCSII_CONN", "CHCSII_SQLOPS_SP is not supported in MySQL mode."
        
        Case eBackend.ITT
            'Return ITT_SQLOPS
            Err.Raise 1001, "CHCSII_CONN", "CHCSII_SQLOPS_SP is not supported in ITT mode."
            
        Case eBackend.LGS
            'Return LGS_SQLOPS
            Err.Raise 1001, "CHCSII_CONN", "CHCSII_SQLOPS_SP is not supported in LGS mode."
            
        Case eBackend.ICD
            'SCR 48849
            '3-8-04 SF Added ICD warning
            Err.Raise 1001, "CHCSII_CONN", "CHCSII_SQLOPS_EX is not supported in ICD mode."
            
    End Select

    Set oSQLOpsSP = Nothing
    
Exit Function

ErrHnd:
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Public Sub GemsHandleDBError(ByVal lErrNumber As Long, ByVal sErrDesc As String, ByVal sModule As String, ByRef mCnn As ADODB.Connection)

'SF 1-26-04 SCR:48421 Created the public sub to expose the HandleDBError sub per Robert B's request.

On Error GoTo ErrHnd

    HandleDBError lErrNumber, sErrDesc, sModule, mCnn
    
Exit Sub

ErrHnd:
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Sub
