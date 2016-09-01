VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cParmType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type mPrivate
    lParmTypeID As Long
    strParmType As String
    strParmTypeDesc As String
    bParmTypeUsage As Byte
End Type
Private m As mPrivate

Property Get ParmTypeID() As Long
    ParmTypeID = m.lParmTypeID
End Property
Property Get ParmType() As String
    ParmType = m.strParmType
End Property
Property Get ParmTypeDesc() As String
    ParmTypeDesc = m.strParmTypeDesc
End Property
Property Get ParmTypeUsage() As Byte
    ParmTypeUsage = m.bParmTypeUsage
End Property

Property Let ParmTypeID(ByVal pData As Long)
    m.lParmTypeID = pData
End Property
Property Let ParmType(pData As String)
    m.strParmType = pData
End Property
Property Let ParmTypeDesc(pData As String)
    m.strParmTypeDesc = pData
End Property
Property Let ParmTypeUsage(pData As Byte)
    m.bParmTypeUsage = pData
End Property
