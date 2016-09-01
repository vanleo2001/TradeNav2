VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTradeLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cTradeLine.cls
'' Description: Wrapper class for Trade table in the database
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 01/21/2009   DAJ         Now storing the Max Profit/Max Loss to the database
'' 01/30/2008   DAJ         Checking Max Profit/Max Loss against Closed Profit
'' 01/30/2009   DAJ         Also check Max Profit/Max Loss against Open Profit
'' 06/02/2009   DAJ         Use Bid/Ask instead of Last for Options P&L
'' 12/01/2009   DAJ         Added support for automatic commissions on fills
'' 04/15/2013   DAJ         When converting to string, subtract fees off of closed profit
'' 10/24/2013   DAJ         Pass account number to g.Profit.Profit
'' 03/19/2015   DAJ         When recalculating OpenProfit, set the correct class level variable
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    lAccountPositionID As Long          ' ID of the account position
    lTradeNumber As Long                ' Unique sequence number of the trade
    strDirection As String              ' Long or Short
    lQuantity As Long                   ' Quantity of the trade
    lEntryFillID As Long                ' Fill ID of the entry fill
    strEntryBrokerFillID As String      ' Broker ID of the entry fill
    dEntryPrice As Double               ' Entry price of the trade
    dEntryTime As Double                ' Entry time of the trade
    lExitFillID As Long                 ' Fill ID of the exit fill
    strExitBrokerFillID As String       ' Broker ID of the exit fill
    dExitPrice As Double                ' Exit price of the trade
    dExitTime As Double                 ' Exit time of the trade
    dClosedProfit As Double             ' Closed profit of the trade
    lEntryRuleID As Long                ' Entry rule ID for the trade
    lExitRuleID As Long                 ' Exit rule ID for the trade
    nRealSimFlag As Byte                ' Real vs Sim flag for the trade
    bIsOpen As Boolean                  ' Is the trade currently open?
    bIsSnapshot As Boolean              ' Is the trade a snapshot trade?
    dFees As Double                     ' Total Fees for the trade
    
    lAccountID As Long                  ' Account ID
    nBroker As eTT_AccountType          ' Broker for the account ID
    lSymbolID As Long                   ' Symbol ID
    strSymbol As String                 ' Symbol
    dOpenProfit As Double               ' Open Profit
    dMaxProfit As Double                ' Max Profit for the trade
    dMaxLoss As Double                  ' Max Loss for the trade
    
    Bars As cGdBars                     ' Bars object
End Type
Private m As mPrivate

Public Property Get AccountPositionID() As Long
    AccountPositionID = m.lAccountPositionID
End Property
Public Property Let AccountPositionID(ByVal lAccountPositionID As Long)
    m.lAccountPositionID = lAccountPositionID
End Property

Public Property Get TradeNumber() As Long
    TradeNumber = m.lTradeNumber
End Property
Public Property Let TradeNumber(ByVal lTradeNumber As Long)
    m.lTradeNumber = lTradeNumber
End Property

Public Property Get Direction() As String
    Direction = m.strDirection
End Property
Public Property Let Direction(ByVal strDirection As String)
    m.strDirection = strDirection
End Property
Public Property Get DirectionString() As String
    If m.strDirection = "L" Then
        DirectionString = "Long"
    Else
        DirectionString = "Short"
    End If
End Property

Public Property Get Quantity() As Long
    Quantity = m.lQuantity
End Property
Public Property Let Quantity(ByVal lQuantity As Long)
    m.lQuantity = lQuantity
End Property

Public Property Get EntryFillID() As Long
    EntryFillID = m.lEntryFillID
End Property
Public Property Let EntryFillID(ByVal lEntryFillID As Long)
    m.lEntryFillID = lEntryFillID
End Property

Public Property Get EntryBrokerFillID() As String
    EntryBrokerFillID = m.strEntryBrokerFillID
End Property
Public Property Let EntryBrokerFillID(ByVal strEntryBrokerFillID As String)
    m.strEntryBrokerFillID = strEntryBrokerFillID
End Property

Public Property Get EntryPrice() As Double
    EntryPrice = m.dEntryPrice
End Property
Public Property Let EntryPrice(ByVal dEntryPrice As Double)
    m.dEntryPrice = dEntryPrice
End Property
Public Property Get EntryPriceString() As String
    If m.dEntryPrice = 0# Then
        EntryPriceString = ""
    Else
        EntryPriceString = m.Bars.PriceDisplay(m.dEntryPrice)
    End If
End Property

Public Property Get EntryTime() As Double
    EntryTime = m.dEntryTime
End Property
Public Property Let EntryTime(ByVal dEntryTime As Double)
    m.dEntryTime = dEntryTime
End Property

Public Property Get ExitFillID() As Long
    ExitFillID = m.lExitFillID
End Property
Public Property Let ExitFillID(ByVal lExitFillID As Long)
    m.lExitFillID = lExitFillID
End Property

Public Property Get ExitBrokerFillID() As String
    ExitBrokerFillID = m.strExitBrokerFillID
End Property
Public Property Let ExitBrokerFillID(ByVal strExitBrokerFillID As String)
    m.strExitBrokerFillID = strExitBrokerFillID
End Property

Public Property Get ExitPrice() As Double
    ExitPrice = m.dExitPrice
End Property
Public Property Let ExitPrice(ByVal dExitPrice As Double)
    m.dExitPrice = dExitPrice
End Property
Public Property Get ExitPriceString() As String
    If m.dExitPrice = 0# Then
        ExitPriceString = ""
    Else
        ExitPriceString = m.Bars.PriceDisplay(m.dExitPrice)
    End If
End Property

Public Property Get ExitTime() As Double
    ExitTime = m.dExitTime
End Property
Public Property Let ExitTime(ByVal dExitTime As Double)
    m.dExitTime = dExitTime
End Property

Public Property Get ClosedProfit() As Double
    ClosedProfit = m.dClosedProfit
End Property
Public Property Let ClosedProfit(ByVal dClosedProfit As Double)
    m.dClosedProfit = dClosedProfit
End Property

Public Property Get EntryRuleID() As Long
    EntryRuleID = m.lEntryRuleID
End Property
Public Property Let EntryRuleID(ByVal lEntryRuleID As Long)
    m.lEntryRuleID = lEntryRuleID
End Property

Public Property Get ExitRuleID() As Long
    ExitRuleID = m.lExitRuleID
End Property
Public Property Let ExitRuleID(ByVal lExitRuleID As Long)
    m.lExitRuleID = lExitRuleID
End Property

Public Property Get RealSimFlag() As Byte
    RealSimFlag = m.nRealSimFlag
End Property
Public Property Let RealSimFlag(ByVal nRealSimFlag As Byte)
    m.nRealSimFlag = nRealSimFlag
End Property

Public Property Get IsOpen() As Boolean
    IsOpen = m.bIsOpen
End Property
Public Property Let IsOpen(ByVal bIsOpen As Boolean)
    m.bIsOpen = bIsOpen
End Property

Public Property Get IsSnapshot() As Boolean
    IsSnapshot = m.bIsSnapshot
End Property
Public Property Let IsSnapshot(ByVal bIsSnapshot As Boolean)
    m.bIsSnapshot = bIsSnapshot
End Property

Public Property Get Fees() As Double
    Fees = m.dFees
End Property
Public Property Let Fees(ByVal dFees As Double)
    m.dFees = dFees
End Property

Public Property Get AccountID() As Long
    AccountID = m.lAccountID
End Property
Public Property Let AccountID(ByVal lAccountID As Long)
    m.lAccountID = lAccountID
    m.nBroker = g.Broker.AccountTypeForID(lAccountID)
End Property
Public Sub SetAccountID(ByVal lAccountID As Long, Optional ByVal nBroker As eTT_AccountType = -1)
    m.lAccountID = lAccountID
    If nBroker = -1 Then
        m.nBroker = g.Broker.AccountTypeForID(lAccountID)
    Else
        m.nBroker = nBroker
    End If
End Sub

Public Property Get Broker() As eTT_AccountType
    Broker = m.nBroker
End Property

Public Property Get SymbolID() As Long
    SymbolID = m.lSymbolID
End Property

Public Property Get Symbol() As String
    Symbol = m.strSymbol
End Property

Public Property Get SymbolOrSymbolID() As Variant
    If m.lSymbolID = 0& Then
        SymbolOrSymbolID = m.strSymbol
    Else
        SymbolOrSymbolID = m.lSymbolID
    End If
End Property
Public Property Let SymbolOrSymbolID(ByVal vSymbolOrSymbolID As Variant)
    m.lSymbolID = GetSymbolID(vSymbolOrSymbolID)
    m.strSymbol = GetSymbol(vSymbolOrSymbolID)
    SetBarProperties m.Bars, vSymbolOrSymbolID
End Property

Public Sub SetSymbolOrSymbolID(ByVal lSymbolID As Long, ByVal strSymbol As String, Optional ByVal Bars As cGdBars = Nothing)
    m.lSymbolID = lSymbolID
    m.strSymbol = strSymbol
    If Bars Is Nothing Then
        SetBarProperties m.Bars, SymbolOrSymbolID
    Else
        Set m.Bars = Bars
    End If
End Sub

Public Property Get MaxProfit() As Double
    MaxProfit = m.dMaxProfit
End Property

Public Property Get MaxLoss() As Double
    MaxLoss = m.dMaxLoss
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Load
'' Description: Load the trade line from the database
'' Inputs:      Account Position ID, Trade Number, Symbol, Recordset
'' Returns:     True if successful, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Load(ByVal lAccountPositionID As Long, ByVal lTradeNumber As Long, ByVal vSymbolOrSymbolID As Variant, Optional ByVal rs As Recordset = Nothing, Optional ByVal lAccountID As Long = -1, Optional ByVal nBroker As eTT_AccountType = -1, Optional Bars As cGdBars = Nothing) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value from the function
    Dim rs2 As Recordset                ' Recordset into the database
    
    If rs Is Nothing Then
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccountPositionTrades] " & _
                "WHERE [AccountPositionID]=" & Str(lAccountPositionID) & " AND [TradeNumber]=" & Str(lTradeNumber) & ";", dbOpenDynaset)
    End If
    
    If (rs.BOF And rs.EOF) Then
        bReturn = False
    Else
        m.lAccountPositionID = rs!AccountPositionID
        m.lTradeNumber = rs!TradeNumber
        m.strDirection = rs!Direction
        m.lQuantity = rs!Quantity
        m.lEntryFillID = rs!EntryFillID
        m.strEntryBrokerFillID = rs!EntryBrokerFillID
        m.dEntryPrice = rs!EntryPrice
        m.dEntryTime = rs!EntryTime
        m.lExitFillID = rs!ExitFillID
        m.strExitBrokerFillID = rs!ExitBrokerFillID
        m.dExitPrice = rs!ExitPrice
        m.dExitTime = rs!ExitTime
        m.dClosedProfit = rs!ClosedProfit
        m.lEntryRuleID = rs!EntryRuleID
        m.lExitRuleID = rs!ExitRuleID
        m.nRealSimFlag = rs!RealFlag
        m.bIsOpen = rs!IsOpen
        m.bIsSnapshot = rs!IsSnapshot
        m.dFees = rs!Fees
        
        m.dMaxProfit = rs!MaxProfit
        m.dMaxLoss = rs!MaxLoss
        
        ' If the account ID was not passed in, then we need to get it from the database...
        If lAccountID = -1& Then
            Set rs2 = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccountPositions] " & _
                        "WHERE [AccountPositionID]=" & Str(m.lAccountPositionID) & ";", dbOpenDynaset)
            If Not (rs2.BOF And rs2.EOF) Then
                m.lAccountID = rs2!AccountID
            End If
        Else
            m.lAccountID = lAccountID
        End If
        
        ' If the broker was not passed in, then we need to get it based on the account ID...
        If nBroker = -1 Then
            m.nBroker = g.Broker.AccountTypeForID(m.lAccountID)
        Else
            m.nBroker = nBroker
        End If
        
        m.lSymbolID = GetSymbolID(vSymbolOrSymbolID)
        m.strSymbol = GetSymbol(vSymbolOrSymbolID)
        
        ' If the bars were not passed in, then we need to set the bar properties...
        If Bars Is Nothing Then
            SetBarProperties m.Bars, vSymbolOrSymbolID
        Else
            Set m.Bars = Bars
        End If
        
        bReturn = True
    End If
    
    Load = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeLine.Load"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Save
'' Description: Save the trade line to the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Save()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database

    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccountPositionTrades] " & _
            "WHERE [AccountPositionID]=" & Str(m.lAccountPositionID) & " AND [TradeNumber]=" & Str(m.lTradeNumber) & ";", dbOpenDynaset)
    If (rs.BOF And rs.EOF) Then
        rs.AddNew
    Else
        rs.Edit
    End If
    
    rs!AccountPositionID = m.lAccountPositionID
    rs!TradeNumber = m.lTradeNumber
    rs!Direction = m.strDirection
    rs!Quantity = m.lQuantity
    rs!EntryFillID = m.lEntryFillID
    rs!EntryBrokerFillID = m.strEntryBrokerFillID
    rs!EntryPrice = m.dEntryPrice
    rs!EntryTime = m.dEntryTime
    rs!ExitFillID = m.lExitFillID
    rs!ExitBrokerFillID = m.strExitBrokerFillID
    rs!ExitPrice = m.dExitPrice
    rs!ExitTime = m.dExitTime
    rs!ClosedProfit = m.dClosedProfit
    rs!EntryRuleID = m.lEntryRuleID
    rs!ExitRuleID = m.lExitRuleID
    rs!RealFlag = m.nRealSimFlag
    rs!IsOpen = m.bIsOpen
    rs!IsSnapshot = m.bIsSnapshot
    rs!Fees = m.dFees
    
    rs!MaxProfit = m.dMaxProfit
    rs!MaxLoss = m.dMaxLoss
    
    rs.Update
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeLine.Save"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsTheSame
'' Description: Determines if the passed trade line is equivalent to this one
'' Inputs:      Trade Line
'' Returns:     True if equivalent, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsTheSame(ByVal TradeLine As cTradeLine) As Boolean
On Error GoTo ErrSection:

    IsTheSame = (ToString = TradeLine.ToString)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeLine.IsTheSame"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ToString
'' Description: Convert this tradeline to a string for the trade-by-trade file
'' Inputs:      None
'' Returns:     String
''
'' Trade line:  Direction, Entry Time, Entry Price, Entry Rule, Exit Time, Exit Price,
''              Exit Rule, Profit, Max Profit, Max Loss, Bars in Trade, Entry Rule ID,
''              Exit Rule ID, TDOW, TDOM, TDOY, -1 if Low before High or 1 otherwise,
''              Quantity
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ToString() As String
On Error GoTo ErrSection:

    Dim astrReturn As cGdArray          ' Array of information to join together to return

    Set astrReturn = New cGdArray
    astrReturn.Create eGDARRAY_Strings

    astrReturn(0) = m.strDirection
    astrReturn(1) = Str(m.dEntryTime)
    astrReturn(2) = Str(m.dEntryPrice)
    astrReturn(3) = m.strEntryBrokerFillID
    astrReturn(4) = Str(m.dExitTime)
    astrReturn(5) = Str(m.dExitPrice)
    astrReturn(6) = m.strExitBrokerFillID
    astrReturn(7) = Str(m.dClosedProfit - m.dFees)
    astrReturn(8) = Str(m.dMaxProfit)
    astrReturn(9) = Str(m.dMaxLoss)
    astrReturn(10) = ""
    astrReturn(11) = Str(kEntryTradeRuleID + m.lEntryRuleID)
    If m.bIsOpen Then astrReturn(12) = "(null)" Else astrReturn(12) = Str(kExitTradeRuleID + m.lExitRuleID)
    astrReturn(13) = ""
    astrReturn(14) = ""
    astrReturn(15) = ""
    astrReturn(16) = ""
    astrReturn(17) = Str(m.lQuantity)

    ToString = astrReturn.JoinFields(vbTab)
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeLine.ToString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OpenProfit
'' Description: Determine the open profit given the current price
'' Inputs:      Current Price, Date/Time of Current Price
'' Returns:     Open Profit
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OpenProfit(Optional ByVal dCurrentPrice As Double = -99999#, Optional ByVal dLocalDateTime As Double = -99999#) As Double
On Error GoTo ErrSection:

    Dim dReturn As Double               ' Return value for the function
    
    dReturn = 0#
    If m.bIsOpen Then
        If dCurrentPrice = -99999# Then
            dReturn = m.dOpenProfit
        Else
            If dLocalDateTime = -99999# Then
                m.dExitTime = CurrentTime
            Else
                m.dExitTime = dLocalDateTime
            End If
            
            m.dExitTime = ConvertToBrokerDate(m.dExitTime, m.nBroker, m.strSymbol, True)
            If m.strDirection = "L" Then
                dReturn = g.Profit.Profit(m.strSymbol, dCurrentPrice - m.dEntryPrice, m.lQuantity, , m.Bars, , m.lAccountID)
            Else
                dReturn = g.Profit.Profit(m.strSymbol, m.dEntryPrice - dCurrentPrice, m.lQuantity, , m.Bars, , m.lAccountID)
            End If
            
            m.dExitPrice = dCurrentPrice
            m.dOpenProfit = dReturn
        End If
    End If
    
    OpenProfit = dReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeLine.OpenProfit"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CalcMaxPNL
'' Description: Calculate the maximum profit and maximum loss for the trade
'' Inputs:      Bars
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CalcMaxPNL(Optional Bars As cGdBars = Nothing)
On Error GoTo ErrSection:

    Dim dEndTime As Double              ' Ending time
    Dim lStartBar As Long               ' Starting bar for the start date
    Dim lEndBar As Long                 ' Ending bar for the end date
    Dim lIndex As Long                  ' Index into a for loop
    Dim dHighestHigh As Double          ' Highest high
    Dim dLowestLow As Double            ' Lowest low

    ' Determine the ending time for the call to get bars.  If the trade line is
    ' still open, set to zero so that we load to the end of the data...
    If m.bIsOpen Then
        dEndTime = 0#
    Else
        dEndTime = m.dExitTime
    End If
    
    ' If the Bars were not passed in, or the size is zero, then try to load the
    ' data from the entry time of the trade line to the ending time we just figured...
    If Bars Is Nothing Then
        Set Bars = New cGdBars
        DM_GetBars Bars, SymbolOrSymbolID, "1 minute", m.dEntryTime, dEndTime
    ElseIf Bars.Size = 0 Then
        DM_GetBars Bars, SymbolOrSymbolID, "1 minute", m.dEntryTime, dEndTime
    End If
    
    ' Determine the bar that the entry fell on and the bar that the ending time fell
    ' on.  If the trade line is open, set the ending bar to Bars.Size since we will
    ' be subtracting one off later anyway...
    lStartBar = Bars.FindDateTime(m.dEntryTime)
    If dEndTime = 0# Then
        lEndBar = Bars.Size
    Else
        lEndBar = Bars.FindDateTime(dEndTime)
    End If
    
    ' Initialize the highest high and the lowest low to the entry price.  That way if
    ' the entry price was the highest or lowest price, it will already be set that way...
    dHighestHigh = m.dEntryPrice
    dLowestLow = m.dEntryPrice
    
    ' Walk the bars from one after the entry bar to one before the exit bar and check
    ' to see if the high is greater than the highest high so far (and if so set the
    ' highest high) and if the low is less then the lowest low so far (and if so set
    ' the lowest low)...
    For lIndex = lStartBar + 1 To lEndBar - 1
        If Bars(eBARS_High, lIndex) > dHighestHigh Then
            dHighestHigh = Bars(eBARS_High, lIndex)
        End If
        If Bars(eBARS_Low, lIndex) < dLowestLow Then
            dLowestLow = Bars(eBARS_Low, lIndex)
        End If
    Next lIndex
    
    ' Determine the maximum profit and loss based on the difference between the entry price
    ' and the highest high/lowest low as appropriate...
    If m.strDirection = "L" Then
        m.dMaxProfit = g.Profit.Profit(m.strSymbol, dHighestHigh - m.dEntryPrice, m.lQuantity, , , , m.lAccountID)
        m.dMaxLoss = g.Profit.Profit(m.strSymbol, dLowestLow - m.dEntryPrice, m.lQuantity, , , , m.lAccountID)
    Else
        m.dMaxProfit = g.Profit.Profit(m.strSymbol, m.dEntryPrice - dLowestLow, m.lQuantity, , , , m.lAccountID)
        m.dMaxLoss = g.Profit.Profit(m.strSymbol, m.dEntryPrice - dHighestHigh, m.lQuantity, , , , m.lAccountID)
    End If
    
    ' Now check the max profit and max loss against the closed and open profit of the
    ' trade and adjust appropriately...
    m.dMaxProfit = MaxDouble(m.dMaxProfit, m.dClosedProfit)
    m.dMaxProfit = MaxDouble(m.dMaxProfit, OpenProfit(Bars(eBARS_Close, Bars.Size - 1)))
    m.dMaxLoss = MinDouble(m.dMaxLoss, m.dClosedProfit)
    m.dMaxLoss = MinDouble(m.dMaxLoss, OpenProfit)
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeLine.CalcMaxPNL"
    
End Sub

Public Sub ClearMaxPNL()
    m.dMaxLoss = 0
    m.dMaxProfit = 0
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CurrentPrice
'' Description: Determine the current price to use for open profit
'' Inputs:      Last Price, Bid, Ask
'' Returns:     Current Price
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CurrentPrice(ByVal dLast As Double, ByVal dBid As Double, ByVal dAsk As Double) As Double
On Error GoTo ErrSection:

    Dim dReturn As Double               ' Return value for the function

    dReturn = dLast
    
    If InStr(m.strSymbol, " ") <> 0 Then
        If UCase(m.strDirection) = "L" Then
            If dBid <> kNullData Then
                dReturn = dBid
            End If
        ElseIf UCase(m.strDirection) = "S" Then
            If dAsk <> kNullData Then
                dReturn = dAsk
            End If
        End If
    End If
    
    CurrentPrice = dReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeLine.CurrentPrice"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Setup and intialize class members when the class is intialized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:

    Set m.Bars = New cGdBars

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeLine.Class_Initialize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Cleanup class members when the class is terminated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

    Set m.Bars = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeLine.Class_Terminate"
    
End Sub
