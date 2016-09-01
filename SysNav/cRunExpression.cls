VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRunExpression"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cRunExpression.cls
'' Description: Object that will run a coded text expression
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 05/20/2010   DAJ         Null out HLC on last bar so engine doesn't use them
'' 06/03/2010   DAJ         Changes for new Trade Sense Order Groups
'' 08/12/2010   DAJ         Check for streaming availability
'' 08/19/2010   DAJ         Added SessionDate function
'' 10/05/2010   DAJ         On a new bar, recalculate the bars since functions
'' 10/08/2010   DAJ         Get rid of "next bar" when pass off to RunExpressions
'' 10/15/2010   DAJ         Commented out DebugLog calls for efficiency
'' 12/17/2010   DAJ         Added UpdateBarsSince argument to FillCallback
'' 01/07/2011   DAJ         Changed the activation calls
'' 01/11/2011   DAJ         Recalc BarsSince on new data if it is invalid
'' 01/26/2011   DAJ         Fix for HHLL bars when don't have an AcctPos object
'' 04/29/2011   DAJ         Added CurrentPrices call
'' 05/18/2011   DAJ         Added custom start/stop time for Market1
'' 06/08/2011   DAJ         Added last bar good flag on Init calls
'' 09/08/2011   DAJ         Moved auto breakout calculation down to cRunExpression
'' 10/07/2011   DAJ         Moved the logging to a routine based on flag file
'' 12/09/2011   DAJ         Allow Next Bar Open if evaluating on each new bar
'' 01/24/2012   DAJ         New method for handling Next Bar Open
'' 09/26/2012   DAJ         Added New Session flag to UpdateBars
'' 10/03/2012   DAJ         Lot size for forex symbols in TradeSense order groups
'' 10/03/2012   DAJ         Added LastGoodBar mode for cRunExpression and cLateCalculating
'' 11/27/2012   DAJ         Fix for secondary markets with only bar period ( e.g. ",30 Min" )
'' 06/26/2014   DAJ         Roll position on a TradeSense order group
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    astrCodedText As cGdArray           ' Coded text expression
    strRunName As String                ' Unique name of the run for the engine
    strSymbol As String                 ' Symbol for the primary market
    lSymbolID As Long                   ' Symbol ID for the primary market
    strPeriod As String                 ' Period for the primary market
    bAutoBreakoutPeriod As Boolean      ' Automated breakout period?
    lNumBars As Long                    ' Number of bars required for the expression
    dEarliestDateTime As Double         ' Earliest date for bars since functions
    bNeedHHLL As Boolean                ' Do we need the HhLl bars?
    dStartTime As Double                ' Custom start time for the data
    dStopTime As Double                 ' Custom stop time for the data
    lSessionDate As Long                ' Current session date for the bars
    bLastBarGood As Boolean             ' Is the last bar considered "good"?
    
    AcctPos As cAccountPosition         ' Account position object
    Bars As cGdBars                     ' Primary bars for the expression
    SecondaryMarkets As cGdTree         ' Collection of secondary market bars
    HhLlBars As cGdBars                 ' Bars needed for Highest High/Lowest Low calculations
    lateCalc As cLateCalculating        ' Object to handle late calculating items
    Results As cGdTree                  ' Collection of results arrays
End Type
Private m As mPrivate

Public Property Get Expression(ByVal lIndex As Long) As String
    Expression = m.astrCodedText(lIndex)
End Property
Public Property Let Expression(ByVal lIndex As Long, ByVal strExpression As String)
    m.astrCodedText(lIndex) = strExpression
End Property

Private Property Get SymbolOrSymbolID() As Variant
    If m.lSymbolID = 0& Then
        SymbolOrSymbolID = m.strSymbol
    Else
        SymbolOrSymbolID = m.lSymbolID
    End If
End Property
Private Property Let SymbolOrSymbolID(ByVal vSymbolOrSymbolID As Variant)
    m.strSymbol = GetSymbol(vSymbolOrSymbolID)
    m.lSymbolID = GetSymbolID(vSymbolOrSymbolID)
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitSingle
'' Description: Initialize the object with the given values
'' Inputs:      Coded Text, Run Name, Symbol, Period, Num Bars Required,
''              Account position information, Custom Start Time, Custom Stop Time,
''              Last Bar Good?, Lot Size
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub InitSingle(ByVal strCodedText As String, ByVal strRunName As String, ByVal vSymbolOrSymbolID As Variant, ByVal strPeriod As String, ByVal lNumBars As Long, Optional AcctPos As cAccountPosition = Nothing, Optional ByVal dStartTime As Double = kNullData, Optional ByVal dStopTime As Double = kNullData, Optional ByVal bLastBarGood As Boolean = False, Optional ByVal lLotSize As Long = 1&)
On Error GoTo ErrSection:

    Dim astrCodedText As New cGdArray   ' Array of coded text expressions

    astrCodedText.Create eGDARRAY_Strings, 1
    astrCodedText(0) = strCodedText
    
    InitMultiple astrCodedText, strRunName, vSymbolOrSymbolID, strPeriod, lNumBars, AcctPos, Nothing, dStartTime, dStopTime, bLastBarGood, lLotSize

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRunExpression.InitSingle"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitMultipleWithAcctPos
'' Description: Initialize the object with the given values
'' Inputs:      Coded Text, Run Name, Symbol, Period, Num Bars Required,
''              Account position information, Custom Start Time, Custom Stop Time,
''              Last Bar Good, Lot Size
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub InitMultipleWithAcctPos(ByVal astrCodedText As cGdArray, ByVal strRunName As String, ByVal vSymbolOrSymbolID As Variant, ByVal strPeriod As String, ByVal lNumBars As Long, ByVal AcctPos As cAccountPosition, Optional ByVal dStartTime As Double = kNullData, Optional ByVal dStopTime As Double = kNullData, Optional ByVal bLastBarGood As Boolean = False, Optional ByVal lLotSize As Long = 1&)
On Error GoTo ErrSection:

    InitMultiple astrCodedText, strRunName, vSymbolOrSymbolID, strPeriod, lNumBars, AcctPos, Nothing, dStartTime, dStopTime, bLastBarGood, lLotSize
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRunExpression.InitMultipleWithAcctPos"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitMultipleWithFills
'' Description: Initialize the object with the given values
'' Inputs:      Coded Text, Run Name, Symbol, Period, Num Bars Required,
''              Fills, Custom Start Time, Custom Stop Time, Last Bar Good?, Lot Size
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub InitMultipleWithFills(ByVal astrCodedText As cGdArray, ByVal strRunName As String, ByVal vSymbolOrSymbolID As Variant, ByVal strPeriod As String, ByVal lNumBars As Long, ByVal Fills As cGdTree, Optional ByVal dStartTime As Double = kNullData, Optional ByVal dStopTime As Double = kNullData, Optional ByVal bLastBarGood As Boolean = False, Optional ByVal lLotSize As Long = 1&)
On Error GoTo ErrSection:

    InitMultiple astrCodedText, strRunName, vSymbolOrSymbolID, strPeriod, lNumBars, Nothing, Fills, dStartTime, dStopTime, bLastBarGood, lLotSize
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRunExpression.InitMultipleWithFills"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateBars
'' Description: Update the bars
'' Inputs:      Run Expression if New Bar?, Force Run Expression?, New Session?
'' Returns:     True if New Bar, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function UpdateBars(ByVal bRunExpressionIfNewBar As Boolean, Optional ByVal bForceRunExpression As Boolean = False, Optional bNewSession As Boolean) As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bNewBar As Boolean              ' Was a new bar started?
    Dim bUpdateBars As Boolean          ' Were the bars updated?
    Dim strPeriod As String             ' Period for the bars
    
    For lIndex = 4 To m.SecondaryMarkets.Count
        If g.RealTime.UpdateBars(m.SecondaryMarkets(lIndex), bNewBar) = True Then
            If bNewBar = True Then
                LoadBars m.SecondaryMarkets(lIndex).SymbolOrSymbolID, m.SecondaryMarkets(lIndex).Prop(eBARS_PeriodicityStr), m.SecondaryMarkets(lIndex)
            End If
        End If
    Next lIndex

    UpdateHHLL

    bNewSession = False
    bUpdateBars = g.RealTime.UpdateBars(m.Bars, bNewBar)
    If bUpdateBars = True Then
        If bNewBar = True Then
            If m.Bars.SessionDate(m.Bars.Size - 1) > m.lSessionDate Then
                bNewSession = True
                m.lSessionDate = m.Bars.SessionDate(m.Bars.Size - 1)
                If m.bAutoBreakoutPeriod Then
                    strPeriod = g.FractZen.GetFractZenPeriod(m.strSymbol)
                    If strPeriod <> m.strPeriod Then
                        m.strPeriod = strPeriod
                        LoadAllMarkets
                    Else
                        g.RealTime.SpliceBars m.Bars
                    End If
                Else
                    g.RealTime.SpliceBars m.Bars
                End If
            Else
                g.RealTime.SpliceBars m.Bars
            End If
            DumpDebug "cRunExpression.UpdateBars: New Bar for " & m.strSymbol & ", " & m.strPeriod & " --> " & DateFormat(m.Bars(eBARS_DateTime, m.Bars.Size - 1), MM_DD_YYYY, HH_MM_SS, AMPM_UPPER)
        End If
        
        If (bNewBar = True) Or (m.lateCalc.BarsSinceValid = False) Then
            m.lateCalc.UpdateBarsSince m.Bars
        End If

        If bForceRunExpression Then
            DumpDebug vbTab & "cRunExpression.UpdateBars: Running Expressions (Force Run)"
            RunExpression
        ElseIf bNewBar And bRunExpressionIfNewBar Then
            DumpDebug vbTab & "cRunExpression.UpdateBars: Running Expressions (New Bar)"
            RunExpression
        End If
    End If
    
    UpdateBars = bNewBar
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRunExpression.UpdateBars"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdatePosition
'' Description: Update the position information for the late calculating functions
'' Inputs:      Account Position
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdatePosition(ByVal AcctPos As cAccountPosition)
On Error GoTo ErrSection:

    Dim dEntryDate As Double            ' Previous entry date

    dEntryDate = m.lateCalc.EntryDateTime
    m.lateCalc.FromAcctPos AcctPos, m.Bars
    m.dEarliestDateTime = EarliestDateTime
    
    ' If there is an entry now and there wasn't before, reload the HHLL bars...
    If (dEntryDate = kNullData) And (m.lateCalc.EntryDateTime <> kNullData) Then
        LoadHHLL
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRunExpression.UpdatePosition"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillCallback
'' Description: Update the position information for the late calculating functions
'' Inputs:      Fill
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FillCallback(ByVal Fill As cPtFill)
On Error GoTo ErrSection:

    Dim dEntryDate As Double            ' Previous entry date

    dEntryDate = m.lateCalc.EntryDateTime
    m.lateCalc.FillCallback Fill, m.Bars
    m.dEarliestDateTime = EarliestDateTime
    
    ' If there is an entry now and there wasn't before, reload the HHLL bars...
    If (dEntryDate = kNullData) And (m.lateCalc.EntryDateTime <> kNullData) Then
        LoadHHLL
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRunExpression.FillCallback"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RunExpression
'' Description: Run the expression
'' Inputs:      None
'' Returns:     Zero on success, Non-Zero otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function RunExpression() As Long
On Error GoTo ErrSection:

    Dim astrCodedText As New cGdArray   ' Array of coded text expressions
    Dim astrParms As New cGdArray       ' Array of parameters to hand to engine
    Dim astrBarNames As New cGdArray    ' Array of bar names to hand to engine
    Dim alBarHandles As New cGdArray    ' Array of bar handles to hand to engine
    Dim ahResults As New cGdArray       ' Array of handles to the results arrays
    Dim Bars As New cGdBars             ' Bars object
    Dim SecondaryMarkets As New cGdTree ' Secondary markets collection
    Dim Daily As New cGdBars            ' Daily version of the Bars object
    Dim Weekly As New cGdBars           ' Weekly version of the Bars object
    Dim lBars As Long                   ' Index into a for loop
    Dim lReturn As Long                 ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim bRcCode As Boolean              ' Return code from engine calls
    Dim strNbo As String                ' Next bar open
    
    lReturn = kNullData
    If (m.astrCodedText.Size > 0) And (m.Bars.Size > 0) Then
        ' Set the first parameter to the given unique run name...
        astrParms.Create eGDARRAY_Strings
        astrParms(0) = m.strRunName
        
        ' Set up the coded text expressions...
        astrCodedText.Create eGDARRAY_Strings, m.astrCodedText.Size
        For lIndex = 0 To m.astrCodedText.Size - 1
            astrCodedText(lIndex) = m.lateCalc.FixExpression(m.astrCodedText(lIndex))
        
            If m.bLastBarGood = False Then
                strNbo = m.Bars.PriceDisplay(m.Bars(eBARS_Open, m.Bars.Size - 1), False)
                strNbo = "~13" & Format(Len(strNbo), "000") & strNbo
            Else
                strNbo = "~13007-999999"
            End If
            astrCodedText(lIndex) = Replace(astrCodedText(lIndex), "~01011OpenNextBar ~16001( ~07007Market1 ~17001)", strNbo)
        Next lIndex
        
        ' Fill in the array of bar names...
        astrBarNames.Create eGDARRAY_Strings
        astrBarNames.Add "Market1"
        astrBarNames.Add "Daily"
        astrBarNames.Add "Weekly"
        MarketsInExpressions astrCodedText, m.Bars(eBARS_DateTime, 0), True, astrBarNames, Nothing, m.strPeriod
                
        ' Attempt to have the engine setup the expressions...
        bRcCode = SetupExpressions(astrParms, astrBarNames, astrCodedText)
        If Not bRcCode Then
            lReturn = -999998
        Else
            ' Make a copy of the main Bars and Null out the HLC on the last bar...
            Set Bars = m.Bars.MakeCopy
            If m.bLastBarGood = False Then
                Bars.Size = Bars.Size - 1
            End If
            DumpDebug "Current Prices for m.Bars: '" & CurrentPrices & "'"
            DumpDebug "Current Prices for Bars: '" & "o:" & Bars.PriceDisplay(Bars(eBARS_Open, Bars.Size - 1)) & "; h:" & Bars.PriceDisplay(Bars(eBARS_High, Bars.Size - 1)) & "; l: " & Bars.PriceDisplay(Bars(eBARS_Low, Bars.Size - 1)) & "; c: " & Bars.PriceDisplay(Bars(eBARS_Close, Bars.Size - 1)) & "; b: " & Bars.PriceDisplay(Bars(eBARS_Bid, Bars.Size - 1)) & "; a: " & Bars.PriceDisplay(Bars(eBARS_Ask, Bars.Size - 1))
        
            ' Build the Daily and Weekly bars...
            Daily.BuildBars "Daily", Bars.BarsHandle
            Weekly.BuildBars "Weekly", Bars.BarsHandle
                
            ' Set the array of bar handles...
            alBarHandles.Create eGDARRAY_Longs
            alBarHandles.Num(0) = Bars.BarsHandle
            alBarHandles.Num(1) = Daily.BarsHandle
            alBarHandles.Num(2) = Weekly.BarsHandle
            For lBars = 4 To m.SecondaryMarkets.Count
                alBarHandles.Num(lBars - 1) = m.SecondaryMarkets(lBars).BarsHandle
            Next lBars
            
            ahResults.Create eGDARRAY_Longs, m.Results.Count
            For lIndex = 1 To m.Results.Count
                ahResults(lIndex - 1) = m.Results(lIndex).ArrayHandle
            Next lIndex
            
            ' Make sure the parameters array has only one element in it...
            astrParms.Size = 1
        
            ' Run the expression...
            lReturn = RunExpressions(astrParms.ArrayHandle, astrBarNames.ArrayHandle, alBarHandles.ArrayHandle, ahResults.ArrayHandle, ByVal 0&, ByVal 0&)
            
            For lIndex = 1 To m.Results.Count
                DumpDebug "cRunExpressions.RunExpression: " & astrCodedText(lIndex - 1)
                DumpDebug "cRunExpressions.RunExpression: " & vbTab & "Array(" & Str(lIndex) & ") --> Size = " & Str(m.Results(lIndex).Size) & "; Last Value = " & Str(m.Results(lIndex)(m.Results(lIndex).Size - 1))
                DumpDebug "cRunExpressions.RunExpression: " & vbTab & "Bars Size = " & Str(Bars.Size) & "; Daily Bars Size = " & Str(Daily.Size) & "; Weekly Bars Size = " & Str(Weekly.Size)
            Next lIndex
            
            ' Have the engine clean things up...
            SetupExpressions astrParms
        End If
    End If
    
    RunExpression = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRunExpression.RunExpression"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetResults
'' Description: Get the results array for the given index
'' Inputs:      Index
'' Returns:     Results array (Nothing if invalid index)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetResults(ByVal lIndex As Long) As cGdArray
On Error GoTo ErrSection:

    Dim adReturn As cGdArray            ' Array to return from the function
       
    Set adReturn = Nothing
    If (lIndex >= 0) And (lIndex < m.Results.Count) Then
        Set adReturn = m.Results(lIndex + 1)
    End If
    
    Set GetResults = adReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRunExpression.GetResults"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LastValue
'' Description: Get the last value out of the results
'' Inputs:      Index, Last Non Null Value?
'' Returns:     Last Value
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function LastValue(ByVal lIndex As Long, Optional ByVal bLastNonNull As Boolean = True) As Double
On Error GoTo ErrSection:

    Dim dReturn As Double               ' Return value for the function
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim adResults As cGdArray           ' Results array
    
    dReturn = -999999
    Set adResults = GetResults(lIndex)
    If Not adResults Is Nothing Then
        If bLastNonNull = False Then
            If adResults.Size > 0 Then
                dReturn = adResults(adResults.Size - 1)
            End If
        Else
            For lIndex2 = adResults.Size - 1 To 0 Step -1
                If adResults(lIndex2) <> adResults.NullValue Then
                    dReturn = adResults(lIndex2)
                    Exit For
                End If
            Next lIndex2
        End If
    End If
    
    LastValue = dReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRunExpression.LastValue"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EnglishExpression
'' Description: Get the English version of a coded text expression
'' Inputs:      Index
'' Returns:     English text
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function EnglishExpression(ByVal lIndex As Long, Optional ByVal bReplaceLateCalc As Boolean = True) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function

    strReturn = ""
    If (lIndex >= 0) And (lIndex < m.astrCodedText.Size) Then
        If bReplaceLateCalc Then
            strReturn = EnglishFromCoded(m.lateCalc.FixExpression(m.astrCodedText(lIndex)))
        Else
            strReturn = EnglishFromCoded(m.astrCodedText(lIndex))
        End If
        
        strReturn = Replace(strReturn, "Market1", Chr(34) & m.strSymbol & ", " & m.strPeriod & Chr(34))
    End If
    
    EnglishExpression = Trim(strReturn)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRunExpression.EnglishExpression"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AllRtDataAvailable
'' Description: Determine if streaming data is available for all symbols
'' Inputs:      None
'' Returns:     True if streaming data available, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AllRtDataAvailable() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    
    bReturn = g.RealTime.RtDataAvailableForBars(m.Bars)
    If bReturn = True Then
        For lIndex = 1 To m.SecondaryMarkets.Count
            If g.RealTime.RtDataAvailableForBars(m.SecondaryMarkets(lIndex)) = False Then
                bReturn = False
                Exit For
            End If
        Next lIndex
    End If
    
    AllRtDataAvailable = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRunExpression.AllRtDataAvailable"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SessionDate
'' Description: Get the current session date for the bars
'' Inputs:      None
'' Returns:     Session Date (kNullData if invalid)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SessionDate() As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    
    If m.Bars Is Nothing Then
        lReturn = kNullData
    ElseIf m.Bars.Size = 0 Then
        lReturn = kNullData
    Else
        lReturn = m.Bars.SessionDate(m.Bars.Size - 1)
    End If
    
    SessionDate = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRunExpression.SessionDate"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CurrentPrices
'' Description: Return a string of the current prices
'' Inputs:      None
'' Returns:     String version of current prices
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CurrentPrices() As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    strReturn = ""
    If Not m.Bars Is Nothing Then
        If m.bAutoBreakoutPeriod Then
            strReturn = "Period: '" & m.strPeriod & "'; "
        End If
        strReturn = strReturn & "o:" & m.Bars.PriceDisplay(m.Bars(eBARS_Open, m.Bars.Size - 1)) & "; h:" & m.Bars.PriceDisplay(m.Bars(eBARS_High, m.Bars.Size - 1)) & "; l: " & m.Bars.PriceDisplay(m.Bars(eBARS_Low, m.Bars.Size - 1)) & "; c: " & m.Bars.PriceDisplay(m.Bars(eBARS_Close, m.Bars.Size - 1)) & "; b: " & m.Bars.PriceDisplay(m.Bars(eBARS_Bid, m.Bars.Size - 1)) & "; a: " & m.Bars.PriceDisplay(m.Bars(eBARS_Ask, m.Bars.Size - 1))
    End If
    
    CurrentPrices = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRunExpression.CurrentPrices"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ReloadData
'' Description: Reload the data
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ReloadData()
On Error GoTo ErrSection:

    Dim strPeriod As String             ' Auto breakout period

    If m.bAutoBreakoutPeriod Then
        strPeriod = g.FractZen.GetFractZenPeriod(m.strSymbol)
        If strPeriod <> m.strPeriod Then
            m.strPeriod = strPeriod
        End If
    End If

    LoadAllMarkets

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRunExpression.ReloadData"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MakeCopy
'' Description: Make a copy of the exit strategy
'' Inputs:      None
'' Returns:     New Copy
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MakeCopy() As cRunExpression
On Error GoTo ErrSection:

    Dim aCopy As New cRunExpression
    
    aCopy.SetPrivateData m
    Set MakeCopy = aCopy

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRunExpression.MakeCopy"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetPrivateData
'' Description: Set the private properties for a copy of the exit strategy
'' Inputs:      Copy From
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub SetPrivateData(mCopyFrom As mPrivate)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim adResults As cGdArray           ' Results array

    m = mCopyFrom
    
    Set m.astrCodedText = mCopyFrom.astrCodedText.MakeCopy
    Set m.Bars = mCopyFrom.Bars.MakeCopy
    Set m.SecondaryMarkets = mCopyFrom.SecondaryMarkets.MakeCopy
    Set m.Results = mCopyFrom.Results.MakeCopy

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRunExpression.SetPrivateData"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitMultiple
'' Description: Initialize the object with the given values
'' Inputs:      Coded Text, Run Name, Symbol, Period, Num Bars Required,
''              Account position information, Fills, Custom Start Time,
''              Custom Stop Time, Last Bar Good?, Lot Size
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitMultiple(ByVal astrCodedText As cGdArray, ByVal strRunName As String, ByVal vSymbolOrSymbolID As Variant, ByVal strPeriod As String, ByVal lNumBars As Long, ByVal AcctPos As cAccountPosition, ByVal Fills As cGdTree, ByVal dStartTime As Double, ByVal dStopTime As Double, ByVal bLastBarGood As Boolean, ByVal lLotSize As Long)
On Error GoTo ErrSection:

    Set m.astrCodedText = astrCodedText.MakeCopy
    
    m.strRunName = strRunName
    SymbolOrSymbolID = vSymbolOrSymbolID
    If UCase(strPeriod) = "AUTO BREAKOUT" Or UCase(strPeriod) = "FRACTZEN" Then
        m.strPeriod = g.FractZen.GetFractZenPeriod(m.strSymbol)
        m.bAutoBreakoutPeriod = True
    Else
        m.strPeriod = strPeriod
        m.bAutoBreakoutPeriod = False
    End If
    m.lNumBars = lNumBars
    m.bNeedHHLL = NeedsHHLL
    m.dStartTime = dStartTime
    m.dStopTime = dStopTime
    m.bLastBarGood = bLastBarGood
    
    m.lateCalc.LotSize = lLotSize
    m.lateCalc.LastBarGood = bLastBarGood
    
    ' Need to set up the late calculating object from the account position before
    ' loading the bars so that we know if we need to load additional data on the main
    ' set of bars for "Bars Since..." functions...
    If Fills Is Nothing Then
        Set m.AcctPos = AcctPos
        m.lateCalc.FromAcctPos AcctPos
    Else
        Set m.AcctPos = Nothing
        m.lateCalc.FromFills Fills
    End If
    m.dEarliestDateTime = EarliestDateTime
        
    LoadAllMarkets
    
    ' Now that the markets are loaded, need to have the late calculating object figure
    ' out the "Bars Since..." functions...
    m.lateCalc.UpdateBarsInfo m.Bars
    m.lateCalc.UpdateBarsSince m.Bars
    m.lateCalc.UpdateHHLL m.HhLlBars
    
    CreateResultsArray m.astrCodedText.Size, m.Bars.Size
    RunExpression

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRunExpression.InitMultiple"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadBars
'' Description: Load the bars for the expression
'' Inputs:      Symbol, Period, Bars, Earliest Date/Time
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadBars(ByVal vSymbolOrSymbolID As Variant, ByVal strPeriod As String, Bars As cGdBars, Optional ByVal dEarliestDateTime As Double = kNullData)
On Error GoTo ErrSection:

    Dim lNumBars As Long                ' Number of bars
    
    If m.lNumBars > 0 Then
        lNumBars = m.lNumBars * -1&
    Else
        lNumBars = m.lNumBars
    End If

    If DM_GetBars(Bars, vSymbolOrSymbolID, strPeriod, lNumBars) Then
        If (dEarliestDateTime <> kNullData) And (dEarliestDateTime < Bars(eBARS_DateTime, 0)) Then
            DM_GetBars Bars, vSymbolOrSymbolID, strPeriod, dEarliestDateTime
        End If
        
        g.RealTime.AddTickBuffer Bars
        g.RealTime.SpliceBars Bars
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRunExpression.LoadBars"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadAllMarkets
'' Description: Load up all of the markets
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadAllMarkets()
On Error GoTo ErrSection:

    Dim astrBarNames As New cGdArray    ' Bar names
    Dim Daily As New cGdBars            ' Temporary bars object
    Dim Weekly As New cGdBars           ' Temporary bars object
    Dim lIndex As Long                  ' Index into a for loop
    Dim strSymbol As String             ' Symbol for secondary market
    Dim strPeriod As String             ' Period for secondary market
    
    If m.astrCodedText.Size > 0 Then
        SetBarProperties m.Bars, SymbolOrSymbolID
        If m.dStartTime <> kNullData Then
            m.Bars.Prop(eBARS_StartTime) = m.dStartTime
        End If
        If m.dStopTime <> kNullData Then
            m.Bars.Prop(eBARS_EndTime) = m.dStopTime
        End If
        
        LoadBars SymbolOrSymbolID, m.strPeriod, m.Bars, Int(m.dEarliestDateTime)
        If m.Bars.Size > 0 Then
            m.lSessionDate = m.Bars.SessionDate(m.Bars.Size - 1)
        End If
        DumpDebug "cRunExpression.LoadAllMarkets: " & Str(m.Bars.Size) & " Bars Loaded for " & m.strSymbol & ", " & m.strPeriod & " (EarliestDate = " & DateFormat(Int(m.dEarliestDateTime), MM_DD_YYYY) & "; NumBars = " & Str(m.lNumBars) & ")"

        Daily.BuildBars "Daily", m.Bars.BarsHandle
        Weekly.BuildBars "Weekly", Daily.BarsHandle
        
        LoadHHLL
        
        Set m.SecondaryMarkets = New cGdTree
        m.SecondaryMarkets.Add m.Bars
        m.SecondaryMarkets.Add Daily
        m.SecondaryMarkets.Add Weekly
        
        astrBarNames.Add "Market1"
        astrBarNames.Add "Daily"
        astrBarNames.Add "Weekly"
        
        ' DAJ 11/27/2012: Pass in the default symbols so that secondary markets like ",30 Min" can work...
        MarketsInExpressions m.astrCodedText, m.Bars(eBARS_DateTime, 0), True, astrBarNames, m.SecondaryMarkets, m.strPeriod, m.strSymbol
        For lIndex = 4 To m.SecondaryMarkets.Count
            strSymbol = m.SecondaryMarkets(lIndex).Prop(eBARS_Symbol)
            strPeriod = m.SecondaryMarkets(lIndex).Prop(eBARS_PeriodicityStr)
            
            LoadBars strSymbol, strPeriod, m.SecondaryMarkets(lIndex)
            DumpDebug "cRunExpression.LoadAllMarkets: " & Str(m.SecondaryMarkets(lIndex).Size) & " Bars Loaded for " & strSymbol & ", " & strPeriod
        Next lIndex
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRunExpression.LoadAllMarkets"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CreateResultsArray
'' Description: Create the results array
'' Inputs:      Number of Expressions, Size of Results
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CreateResultsArray(ByVal lNumExpressions As Long, ByVal lResultsSize As Long)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim adResults As cGdArray           ' Results array
    
    m.Results.Clear
    For lIndex = 0 To lNumExpressions - 1
        Set adResults = New cGdArray
        adResults.Create eGDARRAY_Doubles, lResultsSize
        
        m.Results.Add adResults
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRunExpression.CreateResultsArray"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EarliestDateTime
'' Description: Earliest date/time needed to calculate bars since functions
'' Inputs:      None
'' Returns:     Earliest Date/Time (-999999 if none)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function EarliestDateTime() As Double
On Error GoTo ErrSection:

    Dim dReturn As Double               ' Return value for the function
    Dim dResult As Double               ' Earliest date/time for the expression
    Dim lIndex As Long                  ' Index into a for loop
    
    dReturn = kNullData
    For lIndex = 0 To m.astrCodedText.Size - 1
        dResult = m.lateCalc.EarliestDateTime(m.astrCodedText(lIndex))
        If (dReturn = kNullData) Or (dResult < dReturn) Then
            dReturn = dResult
        End If
    Next lIndex
    
    EarliestDateTime = dReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRunExpression.EarliestDateTime"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NeedsHHLL
'' Description: Do these expressions need Highest High/Lowest Low?
'' Inputs:      None
'' Returns:     True if needs HHLL, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function NeedsHHLL() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    
    bReturn = False
    If m.astrCodedText.Size > 0 Then
        For lIndex = 0 To m.astrCodedText.Size - 1
            If m.lateCalc.HasHHLL(m.astrCodedText(lIndex)) = True Then
                bReturn = True
                Exit For
            End If
        Next lIndex
    End If
    
    NeedsHHLL = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRunExpression.NeedsHHLL"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadHHLL
'' Description: Load up the HHLL bars
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadHHLL()
On Error GoTo ErrSection:

    Dim dEntryDate As Double            ' Entry date/time
    Dim lSessionDate As Long            ' Session date for the entry date/time

    If m.bNeedHHLL Then
        SetBarProperties m.HhLlBars, m.strSymbol
        dEntryDate = m.lateCalc.EntryDateTime
        If dEntryDate <> kNullData Then
            lSessionDate = m.HhLlBars.SessionDateForTradeTime(dEntryDate)
            
            If lSessionDate > LastDailyDownload Then
                DumpDebug "Load Minute HHLL Bars: Loading HHLL Bars from " & DateFormat(lSessionDate, MM_DD_YYYY)
                LoadBars SymbolOrSymbolID, "1 Minute", m.HhLlBars, lSessionDate
                
                If HhllDataAvailable Then
                    If m.HhLlBars.Size = 0 Then
                        DumpDebug vbTab & "Load Minute HHLL Bars: No HHLL Bars Loaded"
                    Else
                        DumpDebug vbTab & "Load Minute HHLL Bars: Bars Loaded from " & DateFormat(m.HhLlBars(eBARS_DateTime, 0), MM_DD_YYYY, HH_MM_SS, AMPM_UPPER) & " to " & DateFormat(m.HhLlBars(eBARS_DateTime, m.HhLlBars.Size - 1), MM_DD_YYYY, HH_MM_SS, AMPM_UPPER)
                        DumpDebug vbTab & "Load Minute HHLL Bars: HH=" & Str(m.HhLlBars.HighestHigh) & "; LL=" & Str(m.HhLlBars.LowestLow)
                        m.HhLlBars.DeleteFirstBars m.HhLlBars.FindDateTime(dEntryDate) + 1
                        DumpDebug vbTab & "Load Minute HHLL Bars: Bars Chopped from " & DateFormat(dEntryDate, MM_DD_YYYY, HH_MM_SS, AMPM_UPPER)
                        DumpDebug vbTab & "Load Minute HHLL Bars: HH=" & Str(m.HhLlBars.HighestHigh) & "; LL=" & Str(m.HhLlBars.LowestLow)
                    End If
                Else
                    DumpDebug vbTab & "Load Minute HHLL Bars: Data not available"
                End If
            Else
                DumpDebug "Load Daily HHLL Bars: Loading HHLL Bars from " & DateFormat(lSessionDate, MM_DD_YYYY)
                LoadBars SymbolOrSymbolID, "Daily", m.HhLlBars, lSessionDate
                
                If HhllDataAvailable Then
                    If m.HhLlBars.Size = 0 Then
                        DumpDebug vbTab & "Load Daily HHLL Bars: No HHLL Bars Loaded"
                    Else
                        DumpDebug vbTab & "Load Daily HHLL Bars: Bars Loaded from " & DateFormat(m.HhLlBars(eBARS_DateTime, 0), MM_DD_YYYY) & " to " & DateFormat(m.HhLlBars(eBARS_DateTime, m.HhLlBars.Size - 1), MM_DD_YYYY)
                        DumpDebug vbTab & "Load Daily HHLL Bars: HH=" & Str(m.HhLlBars.HighestHigh) & "; LL=" & Str(m.HhLlBars.LowestLow)
                        ChopDailyBars m.HhLlBars, dEntryDate
                        DumpDebug vbTab & "Load Daily HHLL Bars: Bars Chopped from " & DateFormat(dEntryDate, MM_DD_YYYY, HH_MM_SS, AMPM_UPPER)
                        DumpDebug vbTab & "Load Daily HHLL Bars: HH=" & Str(m.HhLlBars.HighestHigh) & "; LL=" & Str(m.HhLlBars.LowestLow)
                    End If
                Else
                    DumpDebug vbTab & "Load Daily HHLL Bars: Data not available"
                End If
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRunExpression.LoadHHLL"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateHHLL
'' Description: Update the HHLL bars
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UpdateHHLL()
On Error GoTo ErrSection:

    Dim dEntryDate As Double            ' Entry date/time
    Dim lSessionDate As Long            ' Session date for the entry date/time
    Dim bNewBar As Boolean              ' Did a new bar occur?

    If m.bNeedHHLL Then
        dEntryDate = m.lateCalc.EntryDateTime
        If dEntryDate <> kNullData Then
            lSessionDate = m.HhLlBars.SessionDateForTradeTime(dEntryDate)
            
            If HhllDataAvailable Then
                If g.RealTime.UpdateBars(m.HhLlBars, bNewBar) Then
                    If bNewBar Then
                        DumpDebug "Update HHLL Bars: New Bar = " & Str(bNewBar)
                        g.RealTime.SpliceBars m.HhLlBars
                        DumpDebug vbTab & "Update HHLL Bars: Bars Spliced from " & DateFormat(m.HhLlBars(eBARS_DateTime, 0), MM_DD_YYYY, HH_MM_SS, AMPM_UPPER) & " to " & DateFormat(m.HhLlBars(eBARS_DateTime, m.HhLlBars.Size - 1), MM_DD_YYYY, HH_MM_SS, AMPM_UPPER)
                        DumpDebug vbTab & "Update HHLL Bars: HH=" & Str(m.HhLlBars.HighestHigh) & "; LL=" & Str(m.HhLlBars.LowestLow)
                        
                        If lSessionDate > LastDailyDownload Then
                            m.HhLlBars.DeleteFirstBars m.HhLlBars.FindDateTime(dEntryDate) + 1
                            DumpDebug vbTab & "Update HHLL Bars: Bars Chopped from " & DateFormat(dEntryDate, MM_DD_YYYY, HH_MM_SS, AMPM_UPPER)
                            DumpDebug vbTab & "Update HHLL Bars: HH=" & Str(m.HhLlBars.HighestHigh) & "; LL=" & Str(m.HhLlBars.LowestLow)
                        End If
                    End If
                
                    m.lateCalc.UpdateHHLL m.HhLlBars
                End If
            Else
                DumpDebug "Update HHLL Bars: Data not available"
            End If
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRunExpression.UpdateHHLL"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HhllDataAvailable
'' Description: Determine if the HHLL data is available from streaming or not
'' Inputs:      None
'' Returns:     True if Available, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function HhllDataAvailable() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = True
    If g.RealTime.SalmonIsRunning Then
        bReturn = (g.RealTime.SymbolInfo(m.HhLlBars.SymbolOrSymbolID).GetDataRequestStatus(m.HhLlBars.Prop(eBARS_Periodicity)) = eSalmonAvailable)
    End If
    
    HhllDataAvailable = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRunExpression.HhllDataAvailable"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DumpDebug
'' Description: Dump to the debug log if there is a flag file
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DumpDebug(ByVal strMessage As String)
On Error GoTo ErrSection:

    If FileExist(AddSlash(App.Path) & "LogRunExpr.FLG") Then
        DebugLog strMessage
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRunExpression.DumpDebug"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Initialize the class members when the object is initialized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:

    Set m.astrCodedText = New cGdArray
    Set m.AcctPos = Nothing
    Set m.Bars = New cGdBars
    Set m.HhLlBars = New cGdBars
    Set m.SecondaryMarkets = New cGdTree
    Set m.lateCalc = New cLateCalculating
    Set m.Results = New cGdTree
    
    m.dEarliestDateTime = kNullData
    m.dStartTime = kNullData
    m.dStopTime = kNullData
    m.bAutoBreakoutPeriod = False
    m.lSessionDate = 0

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRunExpression.Class_Initialize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Clean up the class members when the object is terminated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop

    Set m.astrCodedText = Nothing
    Set m.AcctPos = Nothing
    Set m.Bars = Nothing
    Set m.HhLlBars = Nothing
    Set m.SecondaryMarkets = Nothing
    Set m.lateCalc = Nothing
    Set m.Results = Nothing
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRunExpression.Class_Terminate"
    
End Sub
