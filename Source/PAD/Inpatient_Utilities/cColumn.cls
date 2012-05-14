VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mColID         As Long         '-- reference number
Private mColName       As String       '-- the grid column name
Private mColumnNumber  As Long         '-- the internal column number
Private mColumnHeader  As String       '-- the text for the column header
Private mColumnOrdinal As Long        '-- the current ordinal position of the column
Private mColumnWidth   As Long
Private mRSName        As String       '-- the database column name as it appears in the recordset
Private mRSType        As String       '-- the database column type
Private mObjectName    As String       '-- (unused)

Public Property Get ColID() As Long
    ColID = mColID
End Property

Public Property Let ColID(ByVal NewValue As Long)
    mColID = NewValue
End Property


Public Property Get ColName() As String
    ColName = mColName
End Property

Public Property Let ColName(ByVal NewValue As String)
    mColName = NewValue
End Property


Public Property Get ColumnNumber() As Long
    ColumnNumber = mColumnNumber
End Property

Public Property Let ColumnNumber(ByVal NewValue As Long)
    mColumnNumber = NewValue
End Property

Public Property Get ColumnHeader() As String
    ColumnHeader = mColumnHeader
End Property

Public Property Let ColumnHeader(ByVal NewValue As String)
    mColumnHeader = NewValue
End Property

Public Property Get ColumnOrdinal() As Long
    ColumnOrdinal = mColumnOrdinal
End Property

Public Property Let ColumnOrdinal(ByVal NewValue As Long)
    mColumnOrdinal = NewValue
End Property

Public Property Get columnWidth() As Long
    columnWidth = mColumnWidth
End Property

Public Property Let columnWidth(ByVal NewValue As Long)
    mColumnWidth = NewValue
End Property

Public Property Get RSName() As String
    RSName = mRSName
End Property

Public Property Let RSName(ByVal NewValue As String)
    mRSName = NewValue
End Property


Public Property Get RSType() As String
    RSType = mRSType
End Property

Public Property Let RSType(ByVal NewValue As String)
    mRSType = NewValue
End Property



Public Property Get ObjectName() As String
    ObjectName = mObjectName
End Property

Public Property Let ObjectName(ByVal NewValue As String)
    mObjectName = NewValue
End Property



