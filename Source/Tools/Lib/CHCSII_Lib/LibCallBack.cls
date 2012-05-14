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
Attribute VB_Name = "mclsLibCallBack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''************************************************************************
'''Module:        mclsLibCallBack/LibCallBack.cls - Call back object
'''
'''Description:   Allows indirect communication between two modules
'''
'''Usage:
'''   Server
'''      Public pobjCallBack  as mclsLibCallBack
'''      Sub Something ()
'''         if not pobjCallBack is nothing then
'''            ClientResponse = pobjCallBack.P_vCall(DoSomething )
'''            if ClientResponse then
'''               DoAnotherThing
'''
'''   Client
'''      Private mobjServer               as new Server
'''      Private WithEvents mobjCallBack  as mclsLibCallBack
'''
'''      Sub Something
'''         set mobjCallBack = new mclsLibCallBack
'''         set mobjServer.pobjCallBack = mobjCallBack
'''
'''      Sub mobjCallBack_pevCall (ByVal vRequest As Variant, ByRef vResponse As Variant)
'''         if vRequest then
'''            DoSomethingOrDoNot
'''            vResponse = Something
'''         else
'''            vResponse = AnotherThing
'''
'''Modules Required:
'''
'''Note:
'''
'''Warning:
'''
'''Author         : Alexander Alberstein
'''Modified       :
'''               :
'''************************************************************************
Option Explicit
''''''''''''''' Constanst Section '''''''''''''''''''''''''''''''''''''''''

Const msMODULE_NAME           As String = "mclsLibCallBack."


''''''''''''''' Events Section '''''''''''''''''''''''''''''''''''''''''

Public Event pevCall(ByVal vRequest As Variant, ByRef vResponse As Variant)



Public Function P_vCall(ByVal vRequest As Variant) As Variant
'''************************************************************************
'''Called By:
'''
'''Arguments:
''' vRequest
'''   An abstract value or reference requested to be passed to receiving object
'''
'''Description:   Raise event in the object that has with events reference to this object
'''
'''Note:
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       If successful
'''                  Value returned by receiving object
'''               else
'''                  Empty
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "P_vCall"

   Dim vResponse        As Variant
   
On Error GoTo ErrHnd
   
   RaiseEvent pevCall(vRequest, vResponse)
   If VarType(vResponse) = vbObject Then
      Set P_vCall = vResponse
   Else
      P_vCall = vResponse
   End If
   
Exit Function
ErrHnd:
   P_vCall = Empty
   Call gobjLibApp.LibApp_Error _
        (Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & Erl, True)

End Function


