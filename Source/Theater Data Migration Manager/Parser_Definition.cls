VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "parser_definition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'    start       string
'    start_precede   collection of strings
'    start_follow    collection of strings
'    end     string
'    end_precede collection of strings
'    end_follow  collection of strings
'    exception   collection of strings
'    type        type_definition (enumerated): statement, ignore contents, strip contents

Public start_def As String
Public start_precede As New Collection
Public start_follow As New Collection
Public end_def As New Collection
Public end_precede As New Collection
Public end_follow As New Collection
Public exception As String
Public type_def As type_def_enum
Public embedded_sql As Boolean

