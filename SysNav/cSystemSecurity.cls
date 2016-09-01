VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSystemSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type mPrivate
    lSystemNumber As Long
    lParmID As Long
    strPath As String
    strSymbol As String
    strPeriodicity As String
    strFormat As String
    strSecurityType As String
    strSecurityName As String
    strParmName As String
    lRuleID As Long
    strMarketSymbol As String
End Type
Private m As mPrivate

Property Get SystemNumber() As Long
    SystemNumber = m.lSystemNumber
End Property
Property Get ParmID() As Long
    ParmID = m.lParmID
End Property
Property Get Path() As String
    Path = m.strPath
End Property
Property Get Symbol() As String
    Symbol = m.strSymbol
End Property
Property Get Periodicity() As String
    Periodicity = m.strPeriodicity
End Property
Property Get Format() As String
    Format = m.strFormat
End Property
Property Get SecurityType() As String
    SecurityType = m.strSecurityType
End Property
Property Get SecurityName() As String
    SecurityName = m.strSecurityName
End Property
Property Get ParmName() As String
    ParmName = m.strParmName
End Property
Property Get RuleID() As Long
    RuleID = m.lRuleID
End Property
Property Get MarketSymbol() As String
    MarketSymbol = m.strMarketSymbol
End Property

Property Let SystemNumber(pData As Long)
    m.lSystemNumber = pData
End Property
Property Let ParmID(pData As Long)
    m.lParmID = pData
End Property
Property Let Path(pData As String)
    m.strPath = pData
End Property
Property Let Symbol(pData As String)
    m.strSymbol = pData
End Property
Property Let Periodicity(pData As String)
    m.strPeriodicity = pData
End Property
Property Let Format(pData As String)
    m.strFormat = pData
End Property
Property Let SecurityType(pData As String)
    m.strSecurityType = pData
End Property
Property Let SecurityName(pData As String)
    m.strSecurityName = pData
End Property
Property Let ParmName(pData As String)
    m.strParmName = pData
End Property
Property Let RuleID(pData As Long)
    m.lRuleID = pData
End Property
Property Let MarketSymbol(pData As String)
    m.strMarketSymbol = pData
End Property
