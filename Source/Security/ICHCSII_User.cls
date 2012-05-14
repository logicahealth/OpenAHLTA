VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ICHCSII_User"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"378B728601C1"
Option Base 0
Option Explicit
'<< SCR 35657 -  add priv_cosign and change priv_sign's meaning
Enum CHCSII_PrivilegeType
    Priv_Read = 0
    Priv_Write = 1
    Priv_Delete = 2
    Priv_Sign = 3
    Priv_Need_Cosign = 4
    Priv_Break_The_Glass = 5
    Priv_Cosign = 6
    Priv_Other = 99
End Enum
Enum CHCSII_LoginMode
    Mode_Manual = 0
    Mode_CAC = 1
    Mode_Other = 2
End Enum

'##ModelId=378B75FB000A
Public Function HasPrivilege(NCID As String) As Boolean
End Function

Public Function HasPrivilegeEx(PrivilegeArea As String, PrivilegeType As CHCSII_PrivilegeType, Optional OtherType As Integer) As Boolean
End Function

Public Function GetAttributeNames() As Collection
End Function

Public Function GetAttribute(AttributeName As String) As String
Attribute GetAttribute.VB_Description = "Returns a signature for the supplied text"
End Function

'##ModelId=378B76030251
Public Function ValidatePassword(Password As String) As Boolean
End Function

'##ModelId=378B992D00CF
Public Property Get UserName() As String
End Property

'##ModelId=378B992D01D3
Public Property Get UserId() As String
End Property

'##ModelId=378B992D02D7
Public Property Get Roles() As ICHCSII_Roles
End Property

'##ModelId=378B992D03E6
Public Property Get Resources() As ICHCSII_Resources
End Property

'-- next 4 items for CAC support  --Only Snareworks mode for now (03/2004)
Public Function CACSignText(ByVal PIN As String, ByVal Text As String, ByRef Signature As String, Optional ByRef ErrCode As Long, Optional ByRef ErrMsg As String) As Boolean
End Function

Public Function CACPresent() As Boolean
End Function

Public Function CACValidatePIN(ByVal PIN As String) As Boolean
End Function

Public Property Get LoginMode() As CHCSII_LoginMode
End Property

