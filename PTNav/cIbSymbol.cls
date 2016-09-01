VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cIbSymbol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cIbSymbol.cls
'' Description: Class that handles Interactive Brokers symbol information
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 08/10/2012   DAJ         Added primary exchange
'' 10/27/2015   DAJ         Added TradingClass and Multiplier
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    BrokerSym As cBrokerSymbol          ' Base broker symbol information
    
    strIbBaseSymbol As String           ' Interactive Brokers base symbol
    strIbCurrency As String             ' Currency
    strIbExchanges As String            ' Exchanges
    strIbPrimaryExchange As String      ' Primary Exchange
    strTradingClass As String           ' Trading class
    strMultiplier As String             ' Multiplier
End Type
Private m As mPrivate

Public Property Get BrokerBase() As String
    BrokerBase = m.BrokerSym.BrokerBase
End Property
Public Property Let BrokerBase(ByVal strBrokerBase As String)
    m.BrokerSym.BrokerBase = strBrokerBase
End Property

Public Property Get GenesisBase() As String
    GenesisBase = m.BrokerSym.GenesisBase
End Property
Public Property Let GenesisBase(ByVal strGenesisBase As String)
    m.BrokerSym.GenesisBase = strGenesisBase
End Property

Public Property Get PriceMult() As Double
    PriceMult = m.BrokerSym.PriceMult
End Property
Public Property Let PriceMult(ByVal dPriceMult As Double)
    m.BrokerSym.PriceMult = dPriceMult
End Property

Public Property Get PriceFormat() As String
    PriceFormat = m.BrokerSym.PriceFormat
End Property
Public Property Let PriceFormat(ByVal strPriceFormat As String)
    m.BrokerSym.PriceFormat = strPriceFormat
End Property

Public Property Get OrderTypeMask() As String
    OrderTypeMask = m.BrokerSym.OrderTypeMask
End Property
Public Property Let OrderTypeMask(ByVal strOrderTypeMask As String)
    m.BrokerSym.OrderTypeMask = strOrderTypeMask
End Property
Public Property Get OrderTypeAllowed(ByVal nOrderType As eTT_OrderType) As Boolean
    OrderTypeAllowed = m.BrokerSym.OrderTypeAllowed(nOrderType)
End Property

Public Property Get TifMask() As String
    TifMask = m.BrokerSym.TifMask
End Property
Public Property Let TifMask(ByVal strTifMask As String)
    m.BrokerSym.TifMask = strTifMask
End Property
Public Property Get TifAllowed(ByVal nTif As eTT_TimeInForce) As Boolean
    TifAllowed = m.BrokerSym.TifAllowed(nTif)
End Property

Public Property Get RtOverride() As Boolean
    RtOverride = m.BrokerSym.RtOverride
End Property
Public Property Let RtOverride(ByVal bRtOverride As Boolean)
    m.BrokerSym.RtOverride = bRtOverride
End Property

Public Property Get StrikeMult() As Double
    StrikeMult = m.BrokerSym.StrikeMult
End Property
Public Property Let StrikeMult(ByVal dStrikeMult As Double)
    m.BrokerSym.StrikeMult = dStrikeMult
End Property

Public Property Get BrokerPositionBase() As String
    BrokerPositionBase = m.BrokerSym.BrokerPositionBase
End Property
Public Property Let BrokerPositionBase(ByVal strBrokerPositionBase As String)
    m.BrokerSym.BrokerPositionBase = strBrokerPositionBase
End Property

Public Property Get BrokerExchange() As String
    BrokerExchange = m.BrokerSym.BrokerExchange
End Property
Public Property Let BrokerExchange(ByVal strBrokerExchange As String)
    m.BrokerSym.BrokerExchange = strBrokerExchange
End Property

Public Property Get StrikePriceFormat() As String
    StrikePriceFormat = m.BrokerSym.StrikePriceFormat
End Property
Public Property Let StrikePriceFormat(ByVal strStrikePriceFormat As String)
    m.BrokerSym.StrikePriceFormat = strStrikePriceFormat
End Property

Public Property Get IbBaseSymbol() As String
    IbBaseSymbol = m.strIbBaseSymbol
End Property
Public Property Let IbBaseSymbol(ByVal strIbBaseSymbol As String)
    m.strIbBaseSymbol = strIbBaseSymbol
End Property

Public Property Get IbCurrency() As String
    IbCurrency = m.strIbCurrency
End Property
Public Property Let IbCurrency(ByVal strIbCurrency As String)
    m.strIbCurrency = strIbCurrency
End Property

Public Property Get IbExchanges() As String
    IbExchanges = m.strIbExchanges
End Property
Public Property Let IbExchanges(ByVal strIbExchanges As String)
    m.strIbExchanges = strIbExchanges
End Property

Public Property Get IbPrimaryExchange() As String
    IbPrimaryExchange = m.strIbPrimaryExchange
End Property
Public Property Let IbPrimaryExchange(ByVal strIbPrimaryExchange As String)
    m.strIbPrimaryExchange = strIbPrimaryExchange
End Property

Public Property Get TradingClass() As String
    TradingClass = m.strTradingClass
End Property
Public Property Let TradingClass(ByVal strTradingClass As String)
    m.strTradingClass = strTradingClass
End Property

Public Property Get Multiplier() As String
    Multiplier = m.strMultiplier
End Property
Public Property Let Multiplier(ByVal strMultiplier As String)
    m.strMultiplier = strMultiplier
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FromString
'' Description: Set the class members from the given string
'' Inputs:      Broker Symbol Info
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FromString(ByVal strBrokerSymbolInfo As String)
On Error GoTo ErrSection:

    Dim astrFields As cGdArray          ' Fields broken out into an array
    
    Set astrFields = New cGdArray
    astrFields.SplitFields strBrokerSymbolInfo, vbTab
    
    IbBaseSymbol = astrFields(0)
    BrokerBase = astrFields(1)
    BrokerExchange = astrFields(2)
    ' Description = astrFields(3)
    GenesisBase = astrFields(4)
    PriceMult = Val(astrFields(5))
    OrderTypeMask = astrFields(6)
    TifMask = astrFields(7)
    RtOverride = (astrFields(8) = "1")
    IbCurrency = astrFields(9)
    PriceFormat = astrFields(10)
    StrikeMult = Val(astrFields(11))
    IbExchanges = ""
    IbPrimaryExchange = ""
    TradingClass = astrFields(12)
    Multiplier = astrFields(13)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cIbSymbol.FromString"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ContractFromMY
'' Description: Convert a month code and single digit year to a Genesis contract
'' Inputs:      MY, Date
'' Returns:     Contract
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ContractFromMY(ByVal strMY As String, Optional ByVal lDate As Long = kNullData) As String
On Error GoTo ErrSection:

    ContractFromMY = m.BrokerSym.ContractFromMY(strMY, lDate)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cIbSymbol.ContractFromMY"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ContractToMY
'' Description: Convert a Genesis contract to a month code and single digit year
'' Inputs:      Genesis Contract
'' Returns:     MY
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ContractToMY(ByVal strContract As String) As String
On Error GoTo ErrSection:

    ContractToMY = m.BrokerSym.ContractToMY(strContract)

ErrExit:
    Exit Function

ErrSection:
    RaiseError "cIbSymbol.ContractToMY"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ContractFromMMMYY
'' Description: Convert a MMM YY date to a Genesis contract
'' Inputs:      MMMYY, Date
'' Returns:     Contract
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ContractFromMMMYY(ByVal strMMMYY As String, Optional ByVal lDate As Long = kNullData) As String
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim lYear As Long                   ' Year from the string
    Dim lMonth As Long                  ' Month from the string
    
    lReturn = 0&
    If Len(strMMMYY) = 6 Then
        lYear = CLng(Parse(strMMMYY, " ", 2)) + 2000
        lMonth = MonthNumber(Parse(strMMMYY, " ", 1))
        If lMonth > 0 Then
            lReturn = (lYear * 100) + lMonth
            If lDate = kNullData Then
                lDate = Date
            End If
            
            Do While lReturn < ((Year(lDate) * 100) + Month(lDate))
                lReturn = lReturn + 1000
            Loop
        End If
    End If
    
    ContractFromMMMYY = Str(lReturn)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cIbSymbol.ContractFromMMMYY"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ContractToMMMYY
'' Description: Convert a Genesis contract to a MMM YY format
'' Inputs:      Genesis Contract
'' Returns:     MMM YY
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ContractToMMMYY(ByVal strContract As String) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    strReturn = ""
    If Len(strContract) = 6 Then
        If IsNumeric(strContract) Then
            strReturn = MonthName(CLng(Val(Right(strContract, 2))), True, True) & " " & Mid(strContract, 3, 2)
        End If
    End If
    
    ContractToMMMYY = strReturn

ErrExit:
    Exit Function

ErrSection:
    RaiseError "cIbSymbol.ContractToMMMYY"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MakeCopy
'' Description: Make a copy of the object
'' Inputs:      None
'' Returns:     Copy of the object
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MakeCopy() As cIbSymbol
On Error GoTo ErrSection:

    Dim aCopy As New cIbSymbol          ' Copy to return
    
    aCopy.SetPrivateData m
    Set MakeCopy = aCopy

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cIbSymbol.MakeCopy"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetPrivateData
'' Description: Set the private data for copying purposes
'' Inputs:      Private members to copy from
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub SetPrivateData(mCopyFrom As mPrivate)
On Error GoTo ErrSection:

    m = mCopyFrom
    
    Set m.BrokerSym = mCopyFrom.BrokerSym.MakeCopy
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cIbSymbol.SetPrivateData"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Do any initialization that needs to be done for the class
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:

    Set m.BrokerSym = New cBrokerSymbol
    
    m.strIbBaseSymbol = ""
    m.strIbCurrency = ""
    m.strIbExchanges = ""
    m.strIbPrimaryExchange = ""
    m.strTradingClass = ""
    m.strMultiplier = ""

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cIbSymbol.Class_Initialize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Do any clean up that needs to be done for the class
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

    Set m.BrokerSym = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cIbSymbol.Class_Terminate"
    
End Sub

