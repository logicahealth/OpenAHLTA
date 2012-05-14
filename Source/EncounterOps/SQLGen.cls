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
Attribute VB_Name = "SQLGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private msFields As String
Private msValues As String
Private msUpdate As String
Private meTransType As SQLTransType
Public Enum SQLDataType
    eSQLString = 1
    eSQLNumber = 2
    eSQLDate = 3
End Enum
Public Enum SQLTransType
    eSQLInsert = 1
    eSQLUpdate = 2
    eSQLDelete = 3
End Enum
Public Function GenerateSQL(sTable As String, sWhere As String) As String
    
    Dim sSql As String
    
    Select Case meTransType
    Case eSQLInsert
        sSql = "INSERT INTO " & sTable & " ( " & msFields & " ) VALUES ( " & msValues & " )"
    Case eSQLUpdate
        sSql = "UPDATE " & sTable & " SET " & msUpdate & " WHERE " & sWhere
    End Select
    
    GenerateSQL = sSql
    
End Function

Public Sub AddToSQL(NewField As String, NewValue As Variant, DataType As SQLDataType)

    Dim sFormatedValue As String
    Dim oConn As New Conn
    
    Select Case DataType
    
        Case eSQLString
            sFormatedValue = oConn.SQLQuote(CStr(NewValue))
        Case eSQLDate
            sFormatedValue = oConn.SQLDate(CDate(NewValue))
        Case eSQLNumber
            sFormatedValue = CStr(NewValue)
    End Select
        
    Select Case meTransType
    Case eSQLInsert
        If Len(msFields) = 0 Then
            msFields = NewField
        Else
            msFields = msFields & "," & NewField
        End If
        
        If Len(msValues) = 0 Then
            msValues = sFormatedValue
        Else
            msValues = msValues & "," & sFormatedValue
        End If
            
    Case eSQLUpdate
        If Len(msUpdate) = 0 Then
            msUpdate = NewField & " = " & sFormatedValue
        Else
            msUpdate = msUpdate & " , " & NewField & " = " & sFormatedValue
        End If
        
    End Select
    
    Set oConn = Nothing
    
End Sub

Public Property Let SetTransType(ByRef TransType As SQLTransType)
    meTransType = TransType
End Property

Public Function Clear()
    msUpdate = vbNullString
    msValues = vbNullString
    msFields = vbNullString
End Function

