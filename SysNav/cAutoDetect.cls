VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAutoDetect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cAutoDetect.cls
'' Description: Class to auto detect required bars for an expression
''
'' Author:      Genesis Financial Data Services
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 03/11/2010   DAJ         Mods to MarketsInExpressions (#5580)
'' 06/14/2010   DAJ         Added Daily to default markets in RunExpression (#5777)
'' 06/08/2011   DAJ         Split minute and non-minute time frames, changed times
'' 11/27/2012   DAJ         Changed default futures symbol from SP-067 to ES-067
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    strExpression As String             ' Expression to run
    lSymbolID As Long                   ' Symbol ID to run expression on
    strSymbol As String                 ' Symbol to run expression on
    strPeriod As String                 ' Period to run the expression on
End Type
Private m As mPrivate

Private Property Get SymbolOrSymbolID() As Variant
    If m.lSymbolID <> 0 Then
        SymbolOrSymbolID = m.lSymbolID
    Else
        SymbolOrSymbolID = m.strSymbol
    End If
End Property

Private Property Let SymbolOrSymbolID(ByVal vSymbolOrSymbolID As Variant)
    m.lSymbolID = GetSymbolID(vSymbolOrSymbolID)
    m.strSymbol = GetSymbol(vSymbolOrSymbolID)
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AutoDetect
'' Description: Attempt to auto detect the number bars required for expression
'' Inputs:      Expression, Symbol, Period
'' Returns:     Num Bars Required (-1 if cannot detect)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AutoDetect(ByVal strExpression As String, Optional ByVal vSymbolOrSymbolID As Variant = "", Optional ByVal strPeriod As String = "Daily") As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim nPeriod As eBarsPeriodType      ' Bar period
    
    lReturn = -1&
    If Len(strExpression) > 0 Then
        m.strExpression = strExpression
        SymbolOrSymbolID = vSymbolOrSymbolID
        m.strPeriod = strPeriod
        
        nPeriod = GetPeriodType(GetPeriodicity(strPeriod))
        If nPeriod >= ePRD_Days Then
            lReturn = AutoDetectNonIntraday(strExpression, strPeriod)
            ' TLB 1/23/2014: need to add more daily bars if has "of Monthly" or "of Weekly"
            If lReturn >= 0 And nPeriod = ePRD_Days Then
                If InStr(UCase(strExpression), "MONTHLY") > 0 Then
                    lReturn = lReturn + 25
                ElseIf InStr(UCase(strExpression), "WEEKLY") > 0 Then
                    lReturn = lReturn + 6
                End If
            End If
        ElseIf InStr(UCase(strPeriod), "MINUTE") <> 0 Then
            lReturn = AutoDetectMinute(strExpression, strPeriod)
        Else
            lReturn = AutoDetectNonMinute(strExpression, strPeriod)
        End If
    End If
    
    AutoDetect = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAutoDetect.AutoDetect"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AutoDetectNonIntraday
'' Description: Attempt to auto detect the number bars required for expression
'' Inputs:      Expression, Period
'' Returns:     Num Bars Required (-1 if cannot detect)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AutoDetectNonIntraday(ByVal strExpression As String, ByVal strPeriod As String) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    
    ' Test symbol against one year of data, ten years of data, and full history if needed...
    If m.lSymbolID <> 0 Then
        lReturn = AutoDetectSymbol(m.lSymbolID, 365, 3650)
    Else
        ' 11936: IBM, 41142: ES-067, 50: $DJIA
        lReturn = AutoDetectSymbol(11936, 365, 3650)
        If lReturn = 0& Then
            lReturn = AutoDetectSymbol(41142, 365, 3650)
        End If
        If lReturn = 0& Then
            lReturn = AutoDetectSymbol(50, 365, 3650)
        End If
    End If
    
    AutoDetectNonIntraday = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAutoDetect.AutoDetectNonIntraday"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AutoDetectMinute
'' Description: Attempt to auto detect the number bars required for expression
'' Inputs:      Expression, Period
'' Returns:     Num Bars Required (-1 if cannot detect)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AutoDetectMinute(ByVal strExpression As String, ByVal strPeriod As String) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    
    If m.lSymbolID <> 0 Then
        lReturn = AutoDetectSymbol(m.lSymbolID, 5, 30, 60)
    Else
        ' 11936: IBM, 41142: ES-067, 50: $DJIA
        lReturn = AutoDetectSymbol(11936, 5, 30, 60)
        If lReturn = 0& Then
            lReturn = AutoDetectSymbol(41142, 5, 30, 60)
        End If
        If lReturn = 0& Then
            lReturn = AutoDetectSymbol(50, 5, 30, 60)
        End If
    End If
    
    AutoDetectMinute = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAutoDetect.AutoDetectMinute"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AutoDetectNonMinute
'' Description: Attempt to auto detect the number bars required for expression
'' Inputs:      Expression, Period
'' Returns:     Num Bars Required (-1 if cannot detect)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AutoDetectNonMinute(ByVal strExpression As String, ByVal strPeriod As String) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    
    If m.lSymbolID <> 0 Then
        lReturn = AutoDetectSymbol(m.lSymbolID, 1, 5, 15)
    Else
        ' 11936: IBM, 41142: ES-067, 50: $DJIA
        lReturn = AutoDetectSymbol(11936, 1, 5, 15)
        If lReturn = 0& Then
            lReturn = AutoDetectSymbol(41142, 1, 5, 15)
        End If
        If lReturn = 0& Then
            lReturn = AutoDetectSymbol(50, 1, 5, 15)
        End If
    End If
    
    AutoDetectNonMinute = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAutoDetect.AutoDetectNonMinute"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AutoDetectSymbol
'' Description: Attempt to auto detect the number bars required for expression
'' Inputs:      Symbol ID, Num Days 1, Num Days 2, Num Days 3
'' Returns:     Num Bars Required (-1 if cannot detect)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AutoDetectSymbol(ByVal lSymbolID As Long, ByVal lNumDays1 As Long, ByVal lNumDays2 As Long, Optional ByVal lNumDays3 As Long = -1&) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function

    lReturn = RunExpression(lSymbolID, lNumDays1)
    If lReturn > 0 Then
        If RunExpression(lSymbolID, lNumDays1, True) <> lReturn Then
            lReturn = -1
        End If
    End If
    
    If lReturn = 0& Then
        lReturn = RunExpression(lSymbolID, lNumDays2)
        If lReturn > 0 Then
            If RunExpression(lSymbolID, lNumDays2, True) <> lReturn Then
                lReturn = -1
            End If
        End If
    End If
    
    If lReturn = 0& Then
        lReturn = RunExpression(lSymbolID, lNumDays3)
        If lReturn > 0 Then
            If RunExpression(lSymbolID, lNumDays3, True) <> lReturn Then
                lReturn = -1
            End If
        End If
    End If
    
    AutoDetectSymbol = lReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAutoDetect.AutoDetectSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RunExpression
'' Description: Run the expression to attempt to determine number bars needed
'' Inputs:      Symbol ID, Num Days to Load, Delay one Bar
'' Returns:     Num Bars Required
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function RunExpression(ByVal lSymbolID As Long, Optional ByVal nNumDaysToLoad As Long = -1&, Optional ByVal bDelayOneBar As Boolean = False) As Long
On Error GoTo ErrSection:

    Dim i&, ii&, rc&, d#, hArray&, nAutoDetect&
    Dim nCount&, nStartDate&
    Dim strCodedText$
    Dim Bars As New cGdBars
    Dim Daily As New cGdBars
    Dim Weekly As New cGdBars
    Dim Monthly As New cGdBars
    Dim GC As New cGdBars
    
    Dim astrParms As New cGdArray, astrBarNames As New cGdArray
    Dim aScanExpr As New cGdArray, aArrayOfResults As New cGdArray
    Dim aArrayOfBars As New cGdArray
    Dim aScanArrays As New cGdArray
    Dim aMinBarsReq As New cGdArray
    
    Dim iDayOfWeek As Integer
    
    Dim SecondaryMarkets As New cGdTree ' Bars collection of secondary markets
    Dim lBars As Long                   ' Index into a for loop
    Dim strSymbol As String             ' Symbol for the given symbol ID
    Dim bInvalidSecondaryPeriod As Boolean

    ' Get coded text and handle of values array from each Criteria
    aScanExpr.Create eGDARRAY_Strings
    aScanArrays.Create eGDARRAY_Longs
    aArrayOfResults.Create eGDARRAY_Longs
    aMinBarsReq.Create eGDARRAY_Longs
    strCodedText = Trim(m.strExpression)
    strSymbol = GetSymbol(lSymbolID)

    If Len(strCodedText) > 0 Then
        If nNumDaysToLoad = -1& Then
            nStartDate = 0
        Else
            nStartDate = LastDailyDownload - nNumDaysToLoad
            iDayOfWeek = Weekday(nStartDate)
            nStartDate = nStartDate - (iDayOfWeek - vbMonday)
        End If

        aArrayOfBars.Create eGDARRAY_Longs
        Bars.Size = 0
            
        If lSymbolID <> 0 Then
            ' load a year's worth of data
            If Not DM_GetBars(Bars, lSymbolID, m.strPeriod, nStartDate, 0, , , , False) Then
                Bars.Size = 0
            ElseIf bDelayOneBar Then
                ' need to start with next bar
                nStartDate = Int(Bars(eBARS_DateTime, 0)) + 1
                'If optWeekly.Value = True Then
                If (UCase(m.strPeriod) = "WEEKLY") Then
                    nStartDate = nStartDate + 6
                End If
                If Not DM_GetBars(Bars, lSymbolID, m.strPeriod, nStartDate, 0, , , , False) Then
                    Bars.Size = 0
                End If
            End If
            nStartDate = Bars(eBARS_DateTime, 0)
        End If
            
        If Bars.Size > 0 Then
            aScanExpr.Add strCodedText
            
            astrBarNames(0) = "Market1"
            astrBarNames(1) = "Daily"
            astrBarNames(2) = "Weekly"
            astrBarNames(3) = "Monthly"
            
            MarketsInExpressions aScanExpr, nStartDate, False, astrBarNames, SecondaryMarkets, m.strPeriod, strSymbol, bInvalidSecondaryPeriod
            
            If bInvalidSecondaryPeriod Then
                nAutoDetect = -1&
            Else
                Daily.BuildBars "Daily", Bars.BarsHandle
                Weekly.BuildBars "Weekly", Bars.BarsHandle
                Monthly.BuildBars "Monthly", Bars.BarsHandle
                
                ' create a temporary result array to be used
                ' by the expression evaluator
                hArray = gdCreateArray(eGDARRAY_Doubles, Bars.Size)
                aArrayOfResults.Add hArray
                
                ' Init the expression evaluator with list of scan expressions
                'astrBarNames(0) = "Market1"
                'astrBarNames(1) = "Weekly"
                'astrBarNames(2) = "GC"
                astrParms(0) = "AutoDetectRunExpr"
                If Not SetupExpressions(astrParms, astrBarNames, aScanExpr) Then
                    'InfBox "i=[] ; h=Auto Detect ; An error exists in expression."
                    Exit Function
                End If
        
                ' run engine to evaluate expressions for this symbol
                aArrayOfBars.Num(0) = Bars.BarsHandle '(in case changed)
                aArrayOfBars.Num(1) = Daily.BarsHandle
                aArrayOfBars.Num(2) = Weekly.BarsHandle
                aArrayOfBars.Num(3) = Monthly.BarsHandle
                For lBars = 4 To astrBarNames.Size - 1
                    aArrayOfBars.Num(lBars) = SecondaryMarkets(lBars + 1).BarsHandle
                Next lBars
                astrParms.Size = 1
                rc = RunExpressions(astrParms.ArrayHandle, _
                    astrBarNames.ArrayHandle, aArrayOfBars.ArrayHandle, _
                    aArrayOfResults.ArrayHandle, aMinBarsReq.ArrayHandle, ByVal 0&)
                If rc = 0 Then
                    ' see if last value is not null
                    If aMinBarsReq.Size > 0 Then
                        ' new method (engine calculates the number)
                        If aMinBarsReq(0) < Bars.Size Then
                            nAutoDetect = aMinBarsReq(0) + 1
                            If (UCase(m.strPeriod) <> "MONTHLY") And (InStr(UCase(m.strExpression), "~07007MONTHLY") <> 0) Then
                                If nAutoDetect = 0 Then
                                    nAutoDetect = 23
                                Else
                                    ' figure number of daily bars for full months
                                    d = Bars(eBARS_DateTime, nAutoDetect - 1) - Bars(eBARS_DateTime, 0)
                                    nAutoDetect = Int((d + 30) / 31) * 23
                                End If
                            ElseIf (UCase(m.strPeriod) <> "WEEKLY") And (InStr(UCase(m.strExpression), "~07006WEEKLY") <> 0) Then
                                If nAutoDetect = 0 Then
                                    nAutoDetect = 5
                                Else
                                    ' figure number of daily bars for full weeks
                                    d = Bars(eBARS_DateTime, nAutoDetect - 1) - Bars(eBARS_DateTime, 0)
                                    nAutoDetect = Int((d + 6) / 7) * 5
                                End If
                            End If
                        End If
                    Else
                        hArray = aArrayOfResults.Num(0)
                        d = gdGetNum(hArray, gdGetSize(hArray) - 1)
                        If d <> gdNullValue(hArray) Then
                            ' if so, find first non-null item
                            For i = 0 To gdGetSize(hArray) - 1
                                d = gdGetNum(hArray, i)
                                If d <> gdNullValue(hArray) Then
                                    nAutoDetect = i + 1
                                    If InStr(UCase(m.strExpression), "~07006WEEKLY") <> 0 Then
                                        gdCopy Bars.ArrayHandle(eBARS_Close), hArray
                                        Bars.BuildBars "Weekly"
                                        For ii = 0 To Bars.Size - 1
                                            If Bars(eBARS_Close, ii) <> gdNullValue(Bars.ArrayHandle(eBARS_Close)) Then
                                                nAutoDetect = Int((ii + 4) / 5)
                                                Exit For
                                            End If
                                        Next ii
                                    End If
                                    Exit For
                                End If
                            Next
                        End If
                    End If
                End If
                ' clear the expression evaluator when done with it
                SetupExpressions astrParms '(clear expressions)
            End If
        End If
    End If
    
    ' destroy all the temporary result arrays
    For i = 0 To aArrayOfResults.Size - 1
        gdDestroyArray aArrayOfResults(i)
    Next
    aArrayOfResults.Size = 0
    
    RunExpression = nAutoDetect

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAutoDetect.RunExpression"

End Function
