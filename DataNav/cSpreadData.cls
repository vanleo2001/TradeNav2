VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSpreadData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cSpreadData.cls
'' Description: Object to calculate and provide bid/ask data for a spread
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 06/11/2012   DAJ         If any component is Null, Null out the bid/ask of the spread
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    strSymbol As String                 ' Symbol for the spread
    dBid As Double                      ' Bid price for the spread
    dBidTime As Double                  ' Time the bid price changed
    dAsk As Double                      ' Ask price for the spread
    dAskTime As Double                  ' Time the ask price changed
    lNumLegs As Long                    ' Number of legs
    
    astrSymbols As cGdArray             ' Array of component symbols
    adMultipliers As cGdArray           ' Array of component multipliers
    adBids As cGdArray                  ' Array of component bids
    adAsks As cGdArray                  ' Array of component asks
    
    Bars As cGdBars                     ' Bars object
End Type
Private m As mPrivate

Public Property Get Symbol() As String
    Symbol = m.strSymbol
End Property
Public Property Let Symbol(ByVal strSymbol As String)
On Error GoTo ErrSection:

    m.strSymbol = strSymbol
    SplitSymbol
    
ErrExit:
    Exit Property

ErrSection:
    RaiseError "cSpreadData.Symbol.Let"
    
End Property

Public Property Get Bid() As Double
    Bid = m.dBid
End Property
Private Property Let Bid(ByVal dBid As Double)
    m.dBid = dBid
    m.Bars(eBARS_Bid, 0) = dBid
End Property

Public Property Get BidTime() As Double
    BidTime = m.dBidTime
End Property

Public Property Get Ask() As Double
    Ask = m.dAsk
End Property
Private Property Let Ask(ByVal dAsk As Double)
    m.dAsk = dAsk
    m.Bars(eBARS_Ask, 0) = dAsk
End Property

Public Property Get AskTime() As Double
    AskTime = m.dAskTime
End Property

Public Property Get NumLegs() As Long
    NumLegs = m.lNumLegs
End Property

Public Property Get LegSymbols() As cGdArray
    Set LegSymbols = m.astrSymbols
End Property

Public Property Get ExchangeTimeZoneInf() As String
    ExchangeTimeZoneInf = m.Bars.Prop(eBARS_ExchangeTimeZoneInf)
End Property

Public Property Get DefaultEndTime() As Double
    DefaultEndTime = m.Bars.Prop(eBARS_DefaultEndTime)
End Property

Public Property Get Bars() As cGdBars
    Set Bars = m.Bars
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateData
'' Description: Update the data
'' Inputs:      None
'' Returns:     True if new values, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function UpdateData() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim bPriceUpdated As Boolean        ' Was a bid or ask updated for any component?
    Dim bBidNull As Boolean             ' Was any component bid null?
    Dim bAskNull As Boolean             ' Was any component ask null?
    Dim dBid As Double                  ' Last known bid price
    Dim dAsk As Double                  ' Last known ask price
    Dim dBidTime As Double              ' Time of the bid
    Dim dAskTime As Double              ' Time of the ask
    Dim dTemp As Double                 ' Temporary variable
    
    bReturn = False
    bPriceUpdated = False
    bBidNull = False
    bAskNull = False
    
    For lIndex = 0 To m.astrSymbols.Size - 1
        dBid = g.RealTime.LastKnownPrice(m.astrSymbols(lIndex), -1, False, dBidTime)
        If dBid <> m.adBids(lIndex) Then
            m.adBids(lIndex) = dBid
            bPriceUpdated = True
        End If
        If m.adMultipliers(lIndex) > 0 Then
            If dBid = kNullData Then
                bBidNull = True
            End If
        Else
            If dAsk = kNullData Then
                bAskNull = True
            End If
        End If
    
        dAsk = g.RealTime.LastKnownPrice(m.astrSymbols(lIndex), 1, False, dAskTime)
        If dAsk <> m.adAsks(lIndex) Then
            m.adAsks(lIndex) = dAsk
            bPriceUpdated = True
        End If
        If m.adMultipliers(lIndex) > 0 Then
            If dAsk = kNullData Then
                bAskNull = True
            End If
        Else
            If dBid = kNullData Then
                bBidNull = True
            End If
        End If
    Next lIndex
    
    If bPriceUpdated Then
        If bBidNull Then
            dBid = kNullData
            dBidTime = kNullData
        Else
            dBid = 0#
            For lIndex = 0 To m.adBids.Size - 1
                If m.adMultipliers(lIndex) > 0 Then
                    If m.adBids(lIndex) = kNullData Then
                        dBid = kNullData
                        dBidTime = kNullData
                        Exit For
                    Else
                        dBid = dBid + (m.adBids(lIndex) * m.adMultipliers(lIndex))
                    End If
                Else
                    If m.adAsks(lIndex) = kNullData Then
                        dBid = kNullData
                        dBidTime = kNullData
                        Exit For
                    Else
                        dBid = dBid + (m.adAsks(lIndex) * m.adMultipliers(lIndex))
                    End If
                End If
            Next lIndex
                
            dBidTime = dBidTime
        End If
        
        If bAskNull Then
            dAsk = kNullData
            dAskTime = kNullData
        Else
            dAsk = 0#
            For lIndex = 0 To m.adAsks.Size - 1
                If m.adMultipliers(lIndex) > 0 Then
                    If m.adAsks(lIndex) = kNullData Then
                        dAsk = kNullData
                        dAskTime = kNullData
                        Exit For
                    Else
                        dAsk = dAsk + (m.adAsks(lIndex) * m.adMultipliers(lIndex))
                    End If
                Else
                    If m.adBids(lIndex) = kNullData Then
                        dAsk = kNullData
                        dAskTime = kNullData
                        Exit For
                    Else
                        dAsk = dAsk + (m.adBids(lIndex) * m.adMultipliers(lIndex))
                    End If
                End If
            Next lIndex
            
            dAskTime = dAskTime
        End If
        
        If (dBid <> kNullData) And (dAsk <> kNullData) Then
            If (dBid < 0) And (dAsk < 0) Then
                dTemp = dBid
                dBid = Abs(dAsk)
                dAsk = Abs(dTemp)
                
                dTemp = dBidTime
                dBidTime = dAskTime
                dAskTime = dTemp
            End If
        End If
        
        If dBid <> m.dBid Then
            Bid = dBid
            m.dBidTime = dBidTime
            bReturn = True
        End If
        If dAsk <> m.dAsk Then
            Ask = dAsk
            m.dAskTime = dAskTime
            bReturn = True
        End If
    End If
    
    UpdateData = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSpreadData.UpdateData"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsValidTradingTime
'' Description: Is the given time a valid trading time
'' Inputs:      Exchange Time, Suspend Time, Resume Time
'' Returns:     True if valid trading time, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsValidTradingTime(ByVal dExchangeDateTime As Double, Optional ByVal dSessionSuspend = kNullData, Optional ByVal dSessionResume = kNullData) As Boolean
On Error GoTo ErrSection:

    IsValidTradingTime = m.Bars.IsValidTradingTime(dExchangeDateTime, dSessionSuspend, dSessionResume)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSpreadData.IsValidTradingTime"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SessionDateForTradeTime
'' Description: Determine the session date for the given time
'' Inputs:      Exchange Time
'' Returns:     Session Date for the given time
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SessionDateForTradeTime(ByVal dExchangeDateTime As Double) As Long
On Error GoTo ErrSection:

    SessionDateForTradeTime = m.Bars.SessionDateForTradeTime(dExchangeDateTime)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSpreadData.SessionDate"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SplitSymbol
'' Description: Split out the symbol information
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SplitSymbol()
On Error GoTo ErrSection:

    Dim astrComponents As cGdArray      ' Array of components from the symbol
    Dim lIndex As Long                  ' Index into a for loop
    
    Set astrComponents = New cGdArray
    astrComponents.SplitFields m.strSymbol, "|"
    
    m.lNumLegs = astrComponents.Size
    ClearArrays m.lNumLegs
    
    For lIndex = 0 To m.lNumLegs - 1
        m.astrSymbols(lIndex) = Parse(astrComponents(lIndex), ";", 2)
        m.adMultipliers(lIndex) = Val(Parse(astrComponents(lIndex), ";", 1))
        m.adBids(lIndex) = kNullData
        m.adAsks(lIndex) = kNullData
    Next lIndex
    
    SetBarProperties m.Bars, m.astrSymbols(0)
    m.Bars.Prop(eBARS_SymbolID) = 0
    m.Bars.Prop(eBARS_Symbol) = m.strSymbol
    m.Bars.Size = 1
    m.Bars(eBARS_Open, 0) = kNullData
    m.Bars(eBARS_High, 0) = kNullData
    m.Bars(eBARS_Low, 0) = kNullData
    m.Bars(eBARS_Close, 0) = kNullData
    m.Bars(eBARS_Bid, 0) = kNullData
    m.Bars(eBARS_Ask, 0) = kNullData
    
    Bid = kNullData
    m.dBidTime = kNullData
    Ask = kNullData
    m.dAskTime = kNullData

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSpreadData.SplitSymbol"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ClearArrays
'' Description: Clear all of the component arrays
'' Inputs:      New Size
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ClearArrays(Optional ByVal lNewSize As Long = -1&)
On Error GoTo ErrSection:

    m.astrSymbols.Clear
    m.adMultipliers.Clear
    m.adBids.Clear
    m.adAsks.Clear
    
    If lNewSize > 0 Then
        m.astrSymbols.Size = lNewSize
        m.adMultipliers.Size = lNewSize
        m.adBids.Size = lNewSize
        m.adAsks.Size = lNewSize
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSpreadData.ClearArrays"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Initialize the class members when the class is initialized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:

    ' Make sure to do this before calling the Bid or Ask setters because they
    ' set the bid and ask in the bars...
    Set m.Bars = New cGdBars
    
    m.strSymbol = ""
    Bid = 0#
    m.dBidTime = kNullData
    Ask = 0#
    m.dAskTime = kNullData
    m.lNumLegs = 0&
    
    Set m.astrSymbols = New cGdArray
    m.astrSymbols.Create eGDARRAY_Strings
    
    Set m.adMultipliers = New cGdArray
    m.adMultipliers.Create eGDARRAY_Longs

    Set m.adBids = New cGdArray
    m.adBids.Create eGDARRAY_Doubles
    
    Set m.adAsks = New cGdArray
    m.adAsks.Create eGDARRAY_Doubles
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSpreadData.Class_Initialize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Clean up the class members when the class is terminated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

    Set m.astrSymbols = Nothing
    Set m.adMultipliers = Nothing
    Set m.adBids = Nothing
    Set m.adAsks = Nothing
    Set m.Bars = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSpreadData.Class_Terminate"
    
End Sub

