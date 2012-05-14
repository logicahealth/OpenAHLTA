VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum enumStandardTools
  stAdd
  stAddProviders
  stCancel
  stDelete
  stEdit
  stRefresh
  stSave
  stSearch
  stSeperator
  stSign
  stTemplates
End Enum

Public Function Add(ByVal StandardTool As enumStandardTools) As iTool

End Function

Public Function AddCustom(ByVal Caption As String, Optional ByVal PicFileName As String = "", _
                          Optional ByVal Tooltip As String = "", Optional ByVal Descrption As String = "", _
                          Optional ByVal Enabled As Boolean = True, Optional ByVal Visible As Boolean = True) As iTool

End Function

Public Function Count() As Long

End Function

Public Sub InitSeed(ByVal SeedValue As Long)

End Sub

Public Function ItemByCaption(ByVal Caption As String) As iTool

End Function

Public Function ItemByID(ByVal ID As Long) As iTool

End Function

Public Function ItemByIndex(ByVal Index As Long) As iTool

End Function

Public Sub Remove(ByVal CaptionOrID As Variant)

End Sub

Public Sub RemoveAll()

End Sub

Public Sub Reset()

End Sub
