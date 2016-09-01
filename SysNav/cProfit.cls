VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cProfit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cProfit.cls
'' Description: Class for calculating profit and loss numbers
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 04/29/2010   DAJ         Try to use prior day on BaseForex if previous session not found
'' 10/24/2013   DAJ         Take account number in Profit; Contract size override
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    astrSymbolMap As cGdArray           ' Symbol map for appropriate forex exchange rate
    astrCache As cGdArray               ' Data cache for current session data
    
    ContractSizes As cGdTree            ' Contract size overrides
    
    bDumpDebug As Boolean               ' Do we want to dump the debug information?
End Type
Private m As mPrivate

Public Property Get ContractSize(ByVal strGenesisSymbol As String, ByVal strAccountNumber As String) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the property
    Dim strKey As String                ' Key into the collection
    
    lReturn = kNullData
    strKey = strGenesisSymbol & "|" & strAccountNumber
    
    If m.ContractSizes.Exists(strKey) Then
        lReturn = m.ContractSizes(strKey)
    End If
    
    ContractSize = lReturn

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cProfit.ContractSize.Get"
    
End Property
Public Property Let ContractSize(ByVal strGenesisSymbol As String, ByVal strAccountNumber As String, ByVal lContractSize As Long)
On Error GoTo ErrSection:

    Dim strKey As String                ' Key into the collection
    
    strKey = strGenesisSymbol & "|" & strAccountNumber
    If m.ContractSizes.Exists(strKey) Then
        m.ContractSizes(strKey) = lContractSize
    Else
        m.ContractSizes.Add lContractSize, strKey
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cProfit.ContractSize.Let"
    
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadSymbolMap
'' Description: Load the symbol map from the file
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub LoadSymbolMap()
On Error GoTo ErrSection:

    Set m.astrSymbolMap = New cGdArray
    m.astrSymbolMap.Create eGDARRAY_Strings
    m.astrSymbolMap.FromFile AddSlash(App.Path) & "Provided\RateMap.TXT"
    m.astrSymbolMap.Sort
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cProfit.LoadSymbolMap"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BaseForex
'' Description: Attempt to retreive the base forex for a given symbol
'' Inputs:      Symbol, Lot Size, Divide?
'' Returns:     Base Forex Symbol (Blank if not applicable)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function BaseForex(ByVal strSymbol As String, Optional lLotSize As Long, Optional bDivide As Boolean) As String
On Error GoTo ErrSection:

    Dim lPos As Long                    ' Position in sorted symbol map
    Dim strReturn As String             ' Return value for the function
    Dim astrSymbolInfo As cGdArray      ' Symbol information array
    
    strReturn = ""
    If m.astrSymbolMap.BinarySearch(strSymbol & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) Then
        Set astrSymbolInfo = New cGdArray
        astrSymbolInfo.SplitFields m.astrSymbolMap(lPos), vbTab
        
        strReturn = Trim(astrSymbolInfo(1))
        lLotSize = CLng(Val(astrSymbolInfo(2)))
        If astrSymbolInfo(3) = "0" Then
            bDivide = False
        Else
            bDivide = True
        End If
    Else
        lLotSize = 1&
        bDivide = True
    End If
    
    BaseForex = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cProfit.BaseForex"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Profit
'' Description: Determine the profit (or loss) for the given information
'' Inputs:      Symbol, Price Difference, Quantity, Exit Date/Time (in exchange time),
''              Bars for Symbol, Bars for Base Forex, Account number
'' Returns:     None
''
'' Note:        If passing in base forex bars, pass in full history of daily
''              bars for the correct base forex symbol
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Profit(ByVal strSymbol As String, ByVal dPriceDifference As Double, Optional ByVal lQuantity As Long = 1&, Optional ByVal dExchangeExitTime As Double = -99999#, Optional ByVal SymbolBars As cGdBars = Nothing, Optional ByVal BaseForexBars As cGdBars = Nothing, Optional ByVal vAccountNumberOrID As Variant = "") As Double
On Error GoTo ErrSection:

    Dim lSessionDate As Long            ' Session date for exit date/time
    Dim lPreviousSession As Long        ' Previous session date for exit date/time
    Dim strBaseForex As String          ' Base forex symbol
    Dim lLotSize As Long                ' Lot size for the symbol
    Dim bDivide As Boolean              ' Do we divide or multiply?
    Dim dReturn As Double               ' Return value for the function
    Dim lBarIndex As Long               ' Index into the bars
    Dim dPrev As Double                 ' Previous value
    Dim bCurrentSession As Boolean      ' Is this being requested for the current session?
    Dim lPos As Long                    ' Position in the cache array
    Dim dExchangeRate As Double         ' Exchange rate to apply
    Dim strKey As String                ' Key into the contract size collection
    Dim strAccountNumber As String      ' Account number

    m.bDumpDebug = FileExist(AddSlash(App.Path) & "DumpProfit.FLG")

    DumpDebug "Profit(" & strSymbol & ", " & Str(dPriceDifference) & ", " & Str(lQuantity) & ", " & Str(dExchangeExitTime) & ")"

    ' Set the bar properties for the symbol bars if not passed in...
    If SymbolBars Is Nothing Then
        Set SymbolBars = New cGdBars
        SetBarProperties SymbolBars, strSymbol
    End If
    
    ' Detrmine the correct base forex symbol...
    strBaseForex = BaseForex(strSymbol, lLotSize, bDivide)
    DumpDebug vbTab & "BaseForex(" & strSymbol & ", " & Str(lLotSize) & ", " & Str(bDivide) & ") = " & Chr(34) & strBaseForex & Chr(34)
    
    If Len(vAccountNumberOrID) > 0 Then
        strAccountNumber = g.Broker.GetAccountNumber(vAccountNumberOrID)
        strKey = strSymbol & "|" & strAccountNumber
        If m.ContractSizes.Exists(strKey) Then
            lLotSize = m.ContractSizes(strKey)
        End If
    End If
    DumpDebug vbTab & "Lot Size overridden to " & Str(lLotSize)
    
    ' Calculate profit normally...
    dReturn = dPriceDifference * SymbolBars.TickValuePerMove * CDbl(lQuantity) * CDbl(lLotSize)
    DumpDebug vbTab & Str(dPriceDifference) & " * " & Str(SymbolBars.TickValuePerMove) & " * " & Str(lQuantity) & " * " & Str(lLotSize) & " = " & Str(dReturn)
    
    ' If there is a base forex symbol, then caluclate the profit based on the previous session
    ' close of the base forex...
    If Len(strBaseForex) > 0 Then
        ' Set the exchange exit time to the current exchange time if not passed in...
        If dExchangeExitTime = -99999# Then
            dExchangeExitTime = ConvertTimeZone(CurrentTime, "", SymbolBars.Prop(eBARS_ExchangeTimeZoneInf))
            lSessionDate = SymbolBars.SessionDateForTradeTime(dExchangeExitTime)
            bCurrentSession = True
        Else
            lSessionDate = SymbolBars.SessionDateForTradeTime(dExchangeExitTime)
            bCurrentSession = (lSessionDate = SymbolBars.SessionDateForTradeTime(ConvertTimeZone(CurrentTime, "", SymbolBars.Prop(eBARS_ExchangeTimeZoneInf))))
        End If
    
        ' Calculate the session date for the exit time and the previous session date...
        lPreviousSession = lSessionDate - 1&
        Do While Not IsWeekday(lPreviousSession)
            lPreviousSession = lPreviousSession - 1&
        Loop
        DumpDebug vbTab & "Exit Time = " & DateFormat(dExchangeExitTime, MM_DD_YYYY, HH_MM_SS, AMPM_UPPER) & " ; Session Date = " & DateFormat(lSessionDate, MM_DD_YYYY) & " ; Previous Session = " & DateFormat(lPreviousSession, MM_DD_YYYY)
    
        ' If this value has already been asked for and has been cached, get it out of the cache
        ' instead of loading the bars...
        If m.astrCache.BinarySearch(strBaseForex & vbTab & Str(lPreviousSession) & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) Then
            dExchangeRate = Val(Parse(m.astrCache(lPos), vbTab, 3))
            DumpDebug vbTab & "Exchange Rate = " & Str(dExchangeRate) & " (Pulled from Cache)"
        Else
            ' Initialize the base forex bars if not passed in...
            If BaseForexBars Is Nothing Then
                Set BaseForexBars = New cGdBars
            End If
            
            ' Load the base forex bars if not passed in (or not passed in correctly)...
            If (BaseForexBars.Size = 0) Or (BaseForexBars.Prop(eBARS_Symbol) <> strBaseForex) Then
                DM_GetBars BaseForexBars, strBaseForex
                g.RealTime.SpliceBars BaseForexBars
            ElseIf (BaseForexBars(eBARS_DateTime, BaseForexBars.Size - 1) < lSessionDate) Or (BaseForexBars(eBARS_DateTime, 0) > lSessionDate) Then
                DM_GetBars BaseForexBars, strBaseForex
                g.RealTime.SpliceBars BaseForexBars
            End If
            
            ' Find the previous session close (if not found or is Null, default to 1)...
            lBarIndex = BaseForexBars.FindDateTime(lPreviousSession, False)
            
            ' DAJ 04/29/2010: If the "previous session" ends up after the currently loaded bars,
            ' use the last known price in the base forex bars...
            If lBarIndex >= BaseForexBars.Size Then
                lBarIndex = BaseForexBars.Size - 1
            End If
            
            If lBarIndex > -1 Then
                dExchangeRate = BaseForexBars(eBARS_Close, lBarIndex)
                DumpDebug vbTab & "Exchange Rate = " & Str(dExchangeRate) & " (Pulled from Data)"
                
                ' If the price out of the bars was Null, default to 1 and dump to log regardless
                ' of the dump to log flag...
                If dExchangeRate = kNullData Then
                    DebugLog "Changing Exchange Rate to 1 because of Null -- Symbol = " & strSymbol & "; Base Forex = " & strBaseForex & "; Previous Session = " & DateFormat(lPreviousSession, MM_DD_YYYY)
                    dExchangeRate = 1#
                End If
            Else
                dExchangeRate = 1#
                DumpDebug vbTab & "Exchange Rate = " & Str(dExchangeRate) & " (Bar doesn't exist)"
            End If
            
            ' If this data is for the current session, cache the data since it is very likely
            ' to be called again (especially if it is being used for open equity while streaming)...
            If bCurrentSession Then
                m.astrCache.Add strBaseForex & vbTab & Str(lPreviousSession) & vbTab & Str(dExchangeRate), lPos
                DumpDebug "Added to Cache: " & m.astrCache(lPos)
            End If
        End If
            
        ' Apply the exchange rate to the profit calculated above as appropriate...
        dPrev = dReturn
        If bDivide Then
            ' If the exchange rate happens to be zero (which it obviously shouldn't be), then
            ' don't apply an exchange rate...
            If dExchangeRate <> 0 Then
                dReturn = RoundNum(dReturn / dExchangeRate, 2)
                DumpDebug vbTab & Str(dPrev) & " / " & Str(dExchangeRate) & " = " & Str(dReturn)
            End If
        Else
            dReturn = RoundNum(dReturn * dExchangeRate, 2)
            DumpDebug vbTab & Str(dPrev) & " * " & Str(dExchangeRate) & " = " & Str(dReturn)
        End If
    End If
    
    Profit = dReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cProfit.Profit"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DumpDebug
'' Description: Dump the given message to a log
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DumpDebug(ByVal strMessage As String)
On Error GoTo ErrSection:

    If m.bDumpDebug = True Then
        DebugLog strMessage
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cProfit.DumpDebug"
    
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

    LoadSymbolMap

    Set m.astrCache = New cGdArray
    m.astrCache.Create eGDARRAY_Strings
    Set m.ContractSizes = New cGdTree
    
    m.bDumpDebug = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cProfit.Class_Initialize"
    
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
    
    Set m.astrSymbolMap = Nothing
    Set m.astrCache = Nothing
    Set m.ContractSizes = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cProfit.Class_Terminate"
    
End Sub
