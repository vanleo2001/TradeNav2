VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBrokerSymbol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cBrokerSymbol.cls
'' Description: Object that represents the relationship between a Genesis symbol
''              and a broker symbol
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 03/07/2011   DAJ         Added contract conversion routines
'' 08/25/2011   DAJ         Added MYY contract conversion
'' 09/07/2011   DAJ         Fixed off-by-one error in TifAllowed, OrderTypeAllowed
'' 12/02/2011   DAJ         Added MMMYY contract conversion
'' 05/03/2012   DAJ         Force month name abbreviations to be in English
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    strBrokerBase As String             ' Base symbol for the broker
    strGenesisBase As String            ' Bsae symbol for Genesis
    dPriceMult As Double                ' Price multiplier
    strPriceFormat As String            ' Price format
    strOrderTypeMask As String          ' Bit mask representing valid order types
    strTifMask As String                ' Bit mask representing valid time-in-force
    bRtOverride As Boolean              ' Real-time override flag
    dStrikeMult As Double               ' Strike price multiplier
    strBrokerPositionBase As String     ' Base symbol for the broker for a position
    strBrokerExchange As String         ' Broker exchange for the symbol
    strStrikeFormat As String           ' Strike price format
End Type
Private m As mPrivate

Public Property Get BrokerBase() As String
    BrokerBase = m.strBrokerBase
End Property
Public Property Let BrokerBase(ByVal strBrokerBase As String)
    m.strBrokerBase = strBrokerBase
End Property

Public Property Get GenesisBase() As String
    GenesisBase = m.strGenesisBase
End Property
Public Property Let GenesisBase(ByVal strGenesisBase As String)
    m.strGenesisBase = strGenesisBase
End Property

Public Property Get PriceMult() As Double
    PriceMult = m.dPriceMult
End Property
Public Property Let PriceMult(ByVal dPriceMult As Double)
    m.dPriceMult = dPriceMult
End Property

Public Property Get PriceFormat() As String
    If Len(m.strPriceFormat) = 0 Then
        PriceFormat = "#.#"
    Else
        PriceFormat = m.strPriceFormat
    End If
End Property
Public Property Let PriceFormat(ByVal strPriceFormat As String)
    m.strPriceFormat = strPriceFormat
End Property

Public Property Get OrderTypeMask() As String
    OrderTypeMask = m.strOrderTypeMask
End Property
Public Property Let OrderTypeMask(ByVal strOrderTypeMask As String)
    m.strOrderTypeMask = strOrderTypeMask
End Property
Public Property Get OrderTypeAllowed(ByVal nOrderType As eTT_OrderType) As Boolean
    If Len(m.strOrderTypeMask) >= nOrderType + 1 Then
        OrderTypeAllowed = (Mid(m.strOrderTypeMask, nOrderType + 1, 1) = "1")
    Else
        OrderTypeAllowed = False
    End If
End Property

Public Property Get TifMask() As String
    TifMask = m.strTifMask
End Property
Public Property Let TifMask(ByVal strTifMask As String)
    m.strTifMask = strTifMask
End Property
Public Property Get TifAllowed(ByVal nTif As eTT_TimeInForce) As Boolean
    If Len(m.strTifMask) >= nTif + 1 Then
        TifAllowed = (Mid(m.strTifMask, nTif + 1, 1) = "1")
    Else
        TifAllowed = False
    End If
End Property

Public Property Get RtOverride() As Boolean
    RtOverride = m.bRtOverride
End Property
Public Property Let RtOverride(ByVal bRtOverride As Boolean)
    m.bRtOverride = bRtOverride
End Property

Public Property Get StrikeMult() As Double
    StrikeMult = m.dStrikeMult
End Property
Public Property Let StrikeMult(ByVal dStrikeMult As Double)
    m.dStrikeMult = dStrikeMult
End Property

Public Property Get BrokerPositionBase() As String
    BrokerPositionBase = m.strBrokerPositionBase
End Property
Public Property Let BrokerPositionBase(ByVal strBrokerPositionBase As String)
    m.strBrokerPositionBase = strBrokerPositionBase
End Property

Public Property Get BrokerExchange() As String
    BrokerExchange = m.strBrokerExchange
End Property
Public Property Let BrokerExchange(ByVal strBrokerExchange As String)
    m.strBrokerExchange = strBrokerExchange
End Property

Public Property Get StrikePriceFormat() As String
    If Len(m.strStrikeFormat) = 0 Then
        StrikePriceFormat = "#.#"
    Else
        StrikePriceFormat = m.strStrikeFormat
    End If
End Property
Public Property Let StrikePriceFormat(ByVal strStrikePriceFormat As String)
    m.strStrikeFormat = strStrikePriceFormat
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
    
    m.strBrokerBase = astrFields(0)
    m.strGenesisBase = astrFields(1)
    m.dPriceMult = Val(astrFields(2))
    m.strPriceFormat = astrFields(3)
    m.strOrderTypeMask = astrFields(4)
    m.strTifMask = astrFields(5)
    m.bRtOverride = (astrFields(6) = "1")
    m.dStrikeMult = Val(astrFields(7))
    m.strBrokerPositionBase = astrFields(8)
    m.strBrokerExchange = astrFields(9)
    m.strStrikeFormat = astrFields(10)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerSymbol.FromString"
    
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

    Dim lReturn As Long                 ' Return value for the function
    Dim lYear As Long                   ' Year from the string
    Dim lMonth As Long                  ' Month from the string
    
    lReturn = 0&
    If Len(strMY) = 2 Then
        If IsDigit(Right(strMY, 1)) Then
            lYear = CLng(Val(Right(strMY, 1))) + 2000
            lMonth = CodeToMonth(Left(strMY, 1))
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
    End If
    
    ContractFromMY = Str(lReturn)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerSymbol.ContractFromMY"
    
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

    Dim strReturn As String             ' Return value for the function
    
    strReturn = ""
    If Len(strContract) = 6 Then
        If IsNumeric(strContract) Then
            strReturn = MonthToCode(CLng(Val(Right(strContract, 2)))) & Mid(strContract, 4, 1)
        End If
    End If
    
    ContractToMY = strReturn

ErrExit:
    Exit Function

ErrSection:
    RaiseError "cBrokerSymbol.ContractToMY"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ContractFromMYY
'' Description: Convert a month code and double digit year to a Genesis contract
'' Inputs:      MYY, Date
'' Returns:     Contract
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ContractFromMYY(ByVal strMYY As String, Optional ByVal lDate As Long = kNullData) As String
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim lYear As Long                   ' Year from the string
    Dim lMonth As Long                  ' Month from the string
    
    lReturn = 0&
    If Len(strMYY) = 3 Then
        If (IsDigit(Right(strMYY, 1)) = True) And (IsDigit(Mid(strMYY, 2, 1)) = True) Then
            lYear = CLng(Val(Mid(strMYY, 2, 2))) + 2000
            lMonth = CodeToMonth(Left(strMYY, 1))
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
    End If
    
    ContractFromMYY = Str(lReturn)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerSymbol.ContractFromMYY"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ContractToMYY
'' Description: Convert a Genesis contract to a month code and double digit year
'' Inputs:      Genesis Contract
'' Returns:     MYY
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ContractToMYY(ByVal strContract As String) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    strReturn = ""
    If Len(strContract) = 6 Then
        If IsNumeric(strContract) Then
            strReturn = MonthToCode(CLng(Val(Right(strContract, 2)))) & Mid(strContract, 3, 2)
        End If
    End If
    
    ContractToMYY = strReturn

ErrExit:
    Exit Function

ErrSection:
    RaiseError "cBrokerSymbol.ContractToMYY"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ContractFromMMMYY
'' Description: Convert a month abbreviation and double digit year to a Genesis contract
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
    If Len(strMMMYY) = 5 Then
        If (IsDigit(Right(strMMMYY, 1)) = True) And (IsDigit(Mid(strMMMYY, 4, 1)) = True) Then
            lYear = CLng(Val(Mid(strMMMYY, 4, 2))) + 2000
            lMonth = MonthNumber(Left(strMMMYY, 3))
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
    End If
    
    ContractFromMMMYY = Str(lReturn)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerSymbol.ContractFromMMMYY"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ContractToMMMYY
'' Description: Convert a Genesis contract to a month abbreviation and double digit year
'' Inputs:      Genesis Contract
'' Returns:     MMMYY
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ContractToMMMYY(ByVal strContract As String) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    strReturn = ""
    If Len(strContract) = 6 Then
        If IsNumeric(strContract) Then
            strReturn = UCase(MonthName(CLng(Val(Right(strContract, 2))), True, True)) & Mid(strContract, 3, 2)
        End If
    End If
    
    ContractToMMMYY = strReturn

ErrExit:
    Exit Function

ErrSection:
    RaiseError "cBrokerSymbol.ContractToMMMYY"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MakeCopy
'' Description: Make a copy of the object
'' Inputs:      None
'' Returns:     Copy of the object
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MakeCopy() As cBrokerSymbol
On Error GoTo ErrSection:

    Dim aCopy As New cBrokerSymbol      ' Copy to return
    
    aCopy.SetPrivateData m
    Set MakeCopy = aCopy

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerSymbol.MakeCopy"

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
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerSymbol.SetPrivateData"
    
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

    m.strBrokerBase = ""
    m.strGenesisBase = ""
    m.dPriceMult = 1#
    m.strPriceFormat = "#.#"
    m.strOrderTypeMask = "00000000"
    m.strTifMask = "000"
    m.bRtOverride = False
    m.dStrikeMult = 1#
    m.strBrokerPositionBase = ""
    m.strBrokerExchange = ""
    m.strStrikeFormat = "#.#"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerSymbol.Class_Initialize"
    
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

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerSymbol.Class_Terminate"
    
End Sub

