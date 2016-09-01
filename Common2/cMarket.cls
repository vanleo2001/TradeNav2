VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMarket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type mPrivate
    lMarketID As Long
    strMarketSymbol As String
    strSecurityType As String
    lSymbolID As Long
    
    dTickMove As Double
    dTickValue As Double
    dMinMoveInTicks As Double
    strMarketDesc As String
    dMargin As Currency
    dContractSize As Double
    bFieldNumber As Byte
End Type
Private m As mPrivate

Property Get FieldNbr() As Byte
    FieldNbr = m.bFieldNumber
End Property
Property Get MarketSymbol() As String
    MarketSymbol = m.strMarketSymbol
End Property
Property Get TickValue() As Double
    TickValue = m.dTickValue
End Property
Property Get MinMoveInTicks() As Double
    MinMoveInTicks = m.dMinMoveInTicks
End Property
Property Get TickMove() As Double
    TickMove = m.dTickMove
End Property
Property Get Margin() As Currency
    Margin = m.dMargin
End Property
Property Get MarketID() As Long
    MarketID = m.lMarketID
End Property
Property Get MarketDesc() As String
    MarketDesc = m.strMarketDesc
End Property
Property Get ContractSize() As Double
    ContractSize = m.dContractSize
End Property
Property Get SecurityType() As String
    SecurityType = m.strSecurityType
End Property
Property Get SymbolID() As Long
    SymbolID = m.lSymbolID
End Property

Property Let MarketID(pData As Long)
    m.lMarketID = pData
End Property
Property Let MarketSymbol(pData As String)
    m.strMarketSymbol = pData
End Property
Property Let TickValue(pData As Double)
    m.dTickValue = pData
End Property
Property Let MinMoveInTicks(pData As Double)
    m.dMinMoveInTicks = pData
End Property
Property Let TickMove(pData As Double)
    m.dTickMove = pData
End Property
Property Let Margin(pData As Currency)
    m.dMargin = pData
End Property
Property Let MarketDesc(pData As String)
    m.strMarketDesc = pData
End Property
Property Let ContractSize(pData As Double)
    m.dContractSize = pData
End Property
Property Let SecurityType(ByVal strSecurityType As String)
    m.strSecurityType = strSecurityType
End Property
Property Let SymbolID(ByVal lSymbolID As Long)
    m.lSymbolID = lSymbolID
End Property
