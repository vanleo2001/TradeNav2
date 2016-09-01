VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLateCalculating"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cLateCalculating.cls
'' Description: Handle macro style replacement of late calculating functions
''              in a coded text expression
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 06/03/2010   DAJ         Changes for new Trade Sense Order Groups
'' 06/14/2010   DAJ         Extra checks in FromAcctPos to stop "With Block..." error
'' 08/11/2010   DAJ         Fixed off by one error in CalcBarsSinceDate
'' 10/05/2010   DAJ         Added the UpdateBarsSince function
'' 10/08/2010   DAJ         Don't calculate BarsSince functions if date is null
'' 10/15/2010   DAJ         Don't replace token in expression unless it is there
'' 01/07/2011   DAJ         Moved position/fill related stuff to different class
'' 01/11/2011   DAJ         Don't allow negative BarsSinceDate
'' 01/13/2011   DAJ         Make sure to convert fill times to NY from broker time
'' 01/18/2011   DAJ         Use gdFixDateTime instead of RoundToSigDigits
'' 04/01/2011   DAJ         Convert date/time to bars time in CalcBarsSinceDate
'' 10/03/2012   DAJ         Lot size for forex symbols in TradeSense order groups
'' 10/03/2012   DAJ         Added LastGoodBar mode for cRunExpression and cLateCalculating
'' 10/23/2012   DAJ         Have BarsSinceEntry/BarsSinceLastEntry return -1 if date past last bar
'' 10/24/2012   DAJ         Have BarsSinceEntry/BarsSinceLastEntry return -1 if date on last bar and calc each new bar
'' 10/24/2013   DAJ         Pass account number to g.Profit.Profit
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Enum eGDToken
    eGDToken_Type = 0
    eGDToken_Name
    eGDToken_CodedName
    eGDToken_Value
    eGDToken_NumFields
End Enum

Private Type mPrivate
    Tokens As cGdTree                   ' Collection of tokens
    Fills As cLateCalculatingFills      ' Fill collection
    lLotSize As Long                    ' Lot size for adjusting positions
    bLastBarGood As Boolean             ' Is the last bar considered "good"?
    lAccountID As Long                  ' Account ID for the fills
End Type
Private m As mPrivate

Private Function TokenField(ByVal nTokenField As eGDToken) As Long
    TokenField = nTokenField
End Function

Public Property Get AverageEntryPrice() As Double
    AverageEntryPrice = Val(GetValue("Average Entry Price"))
End Property
Public Property Let AverageEntryPrice(ByVal dAverageEntryPrice As Double)
    SetValue "Average Entry Price", Str(RoundToSigDigits(dAverageEntryPrice))
End Property

Public Property Get BarsSinceEntry() As Long
    BarsSinceEntry = CLng(Val(GetValue("Bars Since Entry")))
End Property
Public Property Let BarsSinceEntry(ByVal lBarsSinceEntry As Long)
    SetValue "Bars Since Entry", Str(lBarsSinceEntry)
End Property

Public Property Get BarsSinceExit() As Long
    BarsSinceExit = CLng(Val(GetValue("Bars Since Exit")))
End Property
Public Property Let BarsSinceExit(ByVal lBarsSinceExit As Long)
    SetValue "Bars Since Exit", Str(lBarsSinceExit)
End Property

Public Property Get BarsSinceLastEntry() As Long
    BarsSinceLastEntry = CLng(Val(GetValue("Bars Since Last Entry")))
End Property
Public Property Let BarsSinceLastEntry(ByVal lBarsSinceLastEntry As Long)
    SetValue "Bars Since Last Entry", Str(lBarsSinceLastEntry)
End Property

Public Property Get EntryDate() As Double
    EntryDate = Val(GetValue("Entry Date"))
End Property
Public Property Let EntryDate(ByVal dEntryDate As Double)
    SetValue "Entry Date", Str(gdFixDateTime(dEntryDate))
End Property

Public Property Get EntryDateTime() As Double
    EntryDateTime = Val(GetValue("Entry DateTime"))
End Property
Public Property Let EntryDateTime(ByVal dEntryDateTime As Double)
    SetValue "Entry DateTime", Str(gdFixDateTime(dEntryDateTime))
End Property

Public Property Get EntryPrice() As Double
    EntryPrice = Val(GetValue("Entry Price"))
End Property
Public Property Let EntryPrice(ByVal dEntryPrice As Double)
    SetValue "Entry Price", Str(RoundToSigDigits(dEntryPrice))
End Property

Public Property Get EntrySignal() As String
    EntrySignal = GetValue("Entry Signal")
End Property
Public Property Let EntrySignal(ByVal strEntrySignal As String)
    SetValue "Entry Signal", strEntrySignal
End Property

Public Property Get EntryTime() As Double
    EntryTime = Val(GetValue("Entry Time"))
End Property
Public Property Let EntryTime(ByVal dEntryTime As Double)
    SetValue "Entry Time", Str(gdFixDateTime(dEntryTime))
End Property

Public Property Get ExitPosition() As Long
    Position = CLng(Val(GetValue("Exit Position")))
End Property
Public Property Let ExitPosition(ByVal lExitPosition As Long)
    SetValue "Exit Position", Str(lExitPosition)
End Property

Public Property Get ExitPrice() As Double
    ExitPrice = Val(GetValue("Exit Price"))
End Property
Public Property Let ExitPrice(ByVal dExitPrice As Double)
    SetValue "Exit Price", Str(RoundToSigDigits(dExitPrice))
End Property

Public Property Get ExitProfit() As Double
    ExitProfit = Val(GetValue("Exit Profit"))
End Property
Public Property Let ExitProfit(ByVal dExitProfit As Double)
    SetValue "Exit Profit", Str(RoundToSigDigits(dExitProfit))
End Property

Public Property Get ExitDateTime() As Double
    ExitDateTime = Val(GetValue("Exit DateTime"))
End Property
Public Property Let ExitDateTime(ByVal dExitDateTime As Double)
    SetValue "Exit DateTime", Str(gdFixDateTime(dExitDateTime))
End Property

Public Property Get ExitSignal() As String
    ExitSignal = GetValue("Exit Signal")
End Property
Public Property Let ExitSignal(ByVal strExitSignal As String)
    SetValue "Exit Signal", strExitSignal
End Property

Public Property Get HighestHighInTrade() As Double
    HighestHighInTrade = Val(GetValue("Highest High in Trade"))
End Property
Public Property Let HighestHighInTrade(ByVal dHighestHighInTrade As Double)
    SetValue "Highest High in Trade", Str(RoundToSigDigits(dHighestHighInTrade))
End Property

Public Property Get LastEntryDateTime() As Double
    LastEntryDateTime = Val(GetValue("Last Entry DateTime"))
End Property
Public Property Let LastEntryDateTime(ByVal dLastEntryDateTime As Double)
    SetValue "Last Entry DateTime", Str(gdFixDateTime(dLastEntryDateTime))
End Property

Public Property Get LastEntryPrice() As Double
    LastEntryPrice = Val(GetValue("Last Entry Price"))
End Property
Public Property Let LastEntryPrice(ByVal dLastEntryPrice As Double)
    SetValue "Last Entry Price", Str(RoundToSigDigits(dLastEntryPrice))
End Property

Public Property Get LastEntryProfit() As Double
    LastEntryProfit = Val(GetValue("Last Entry Profit"))
End Property
Public Property Let LastEntryProfit(ByVal dLastEntryProfit As Double)
    SetValue "Last Entry Profit", Str(RoundToSigDigits(dLastEntryProfit))
End Property

Public Property Get LastEntrySignal() As String
    LastEntrySignal = GetValue("Last Entry Signal")
End Property
Public Property Let LastEntrySignal(ByVal strLastEntrySignal As String)
    SetValue "Last Entry Signal", strLastEntrySignal
End Property

Public Property Get LowestLowInTrade() As Double
    LowestLowInTrade = Val(GetValue("Lowest Low in Trade"))
End Property
Public Property Let LowestLowInTrade(ByVal dLowestLowInTrade As Double)
    SetValue "Lowest Low in Trade", Str(RoundToSigDigits(dLowestLowInTrade))
End Property

Public Property Get MaxPositionLoss() As Double
    MaxPositionLoss = Val(GetValue("Max Position Loss"))
End Property
Public Property Let MaxPositionLoss(ByVal dMaxPositionLoss As Double)
    SetValue "Max Position Loss", Str(RoundToSigDigits(dMaxPositionLoss))
End Property

Public Property Get MaxPositionProfit() As Double
    MaxPositionProfit = Val(GetValue("Max Position Profit"))
End Property
Public Property Let MaxPositionProfit(ByVal dMaxPositionProfit As Double)
    SetValue "Max Position Profit", Str(RoundToSigDigits(dMaxPositionProfit))
End Property

Public Property Get Position() As Long
    Position = CLng(Val(GetValue("Position")))
End Property
Public Property Let Position(ByVal lPosition As Long)
    SetValue "Position", Str(lPosition)
End Property

Public Property Get PositionProfit() As Double
    PositionProfit = Val(GetValue("Position Profit"))
End Property
Public Property Let PositionProfit(ByVal dPositionProfit As Double)
    SetValue "Position Profit", Str(RoundToSigDigits(dPositionProfit))
End Property

Public Property Get TradedToday() As Boolean
    TradedToday = (UCase(GetValue("Traded Today")) = "TRUE")
End Property
Public Property Let TradedToday(ByVal bTradedToday As Boolean)
    If bTradedToday = True Then
        SetValue "Traded Today", "True"
    Else
        SetValue "Traded Today", "False"
    End If
End Property

Public Property Get LotSize() As Long
    LotSize = m.lLotSize
End Property
Public Property Let LotSize(ByVal lLotSize As Long)
    m.lLotSize = lLotSize
End Property

Public Property Get LastBarGood() As Boolean
    LastBarGood = m.bLastBarGood
End Property
Public Property Let LastBarGood(ByVal bLastBarGood As Boolean)
    m.bLastBarGood = bLastBarGood
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FixExpression
'' Description: Replace the late calculating functions with values in the expression
'' Inputs:      Expression
'' Returns:     Fixed expression
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FixExpression(ByVal strExpression As String) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value in the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim strToReplace As String          ' String to replace
    
    strReturn = strExpression
    For lIndex = 1 To m.Tokens.Count
        strToReplace = CodedTextForToken(m.Tokens(lIndex))
        If InStr(strReturn, strToReplace) <> 0 Then
            strReturn = Replace(strReturn, strToReplace, ReplacementValue(m.Tokens(lIndex)))
        End If
    Next lIndex
    
    FixExpression = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cLateCalculating.FixExpression"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FromAcctPos
'' Description: Fill in the properties from an account position object
'' Inputs:      Account Position, Bars
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FromAcctPos(ByVal AcctPos As cAccountPosition, Optional ByVal Bars As cGdBars = Nothing)
On Error GoTo ErrSection:

    m.Fills.FromAcctPos AcctPos
    UpdateFromFills Bars

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cLateCalculating.FromAcctPos"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FromFills
'' Description: Fill in the properties from a collection of fills
'' Inputs:      Fills, Bars
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FromFills(ByVal Fills As cGdTree, Optional ByVal Bars As cGdBars = Nothing)
On Error GoTo ErrSection:

    m.Fills.FromFills Fills
    UpdateFromFills Bars

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cLateCalculating.FromFills"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateBarsInfo
'' Description: Update the information that needs the Bars
'' Inputs:      Bars
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdateBarsInfo(ByVal Bars As cGdBars)
On Error GoTo ErrSection:

    Dim dLastPrice As Double            ' Last price out of the bars

    ' Determine the last price of the data out of the bars if we can...
    dLastPrice = kNullData
    If Not Bars Is Nothing Then
        If Bars.Size > 0 Then
            If m.bLastBarGood Then
                dLastPrice = Bars(eBARS_Close, Bars.Size - 1)
            ElseIf Bars.Size > 1 Then
                dLastPrice = Bars(eBARS_Close, Bars.Size - 2)
            End If
        End If
    End If

    If Position = 0 Then
        PositionProfit = kNullData
    Else
        PositionProfit = m.Fills.OpenProfit(dLastPrice)
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cLateCalculating.UpdateBarsInfo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillCallback
'' Description: If we get a fill callback, recalc things
'' Inputs:      Fill, Bars
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FillCallback(ByVal Fill As cPtFill, Optional ByVal Bars As cGdBars = Nothing)
On Error GoTo ErrSection:

    m.Fills.FillCallback Fill
    UpdateFromFills Bars

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cLateCalculating.FillCallback"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EarliestDateTime
'' Description: Earliest date/time needed to calculate bars since functions
'' Inputs:      Expression
'' Returns:     Earliest Date/Time (-999999 if none)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function EarliestDateTime(ByVal strExpression As String) As Double
On Error GoTo ErrSection:

    Dim dReturn As Double               ' Return value for the function
    
    dReturn = kNullData
    
    If Len(strExpression) > 0 Then
        If InStr(strExpression, CodedTextForToken(m.Tokens("Bars Since Exit"))) > 0 Then
            dReturn = ExitDateTime
        End If
        If InStr(strExpression, CodedTextForToken(m.Tokens("Bars Since Entry"))) > 0 Then
            If (dReturn = kNullData) Or (EntryDateTime < dReturn) Then
                dReturn = EntryDateTime
            End If
        End If
        If InStr(strExpression, CodedTextForToken(m.Tokens("Bars Since Last Entry"))) > 0 Then
            If (dReturn = kNullData) Or (LastEntryDateTime < dReturn) Then
                dReturn = LastEntryDateTime
            End If
        End If
    End If
    
    EarliestDateTime = dReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cLateCalculating.EarliestDateTime"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HasHHLL
'' Description: Do these expressions have Highest High/Lowest Low/Max Profit/Min Profit?
'' Inputs:      Expression
'' Returns:     True if has HHLL, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function HasHHLL(ByVal strExpression As String) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    If Len(strExpression) > 0 Then
        If InStr(strExpression, CodedTextForToken(m.Tokens("Highest High in Trade"))) > 0 Then
            bReturn = True
        ElseIf InStr(strExpression, CodedTextForToken(m.Tokens("Lowest Low in Trade"))) > 0 Then
            bReturn = True
        ElseIf InStr(strExpression, CodedTextForToken(m.Tokens("Max Position Profit"))) > 0 Then
            bReturn = True
        ElseIf InStr(strExpression, CodedTextForToken(m.Tokens("Max Position Loss"))) > 0 Then
            bReturn = True
        End If
    End If
    
    HasHHLL = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cLateCalculating.HasHHLL"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateHHLL
'' Description: Update Highest High/Lowest Low/Max Profit/Min Profit
'' Inputs:      Bars
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdateHHLL(ByVal Bars As cGdBars)
On Error GoTo ErrSection:

    Dim dMinValue As Double             ' Min value in the array
    Dim BarsCopy As cGdBars             ' Copy of the bars passed in
    
    Set BarsCopy = Bars.MakeCopy
    If m.bLastBarGood = False Then
        BarsCopy.Size = BarsCopy.Size - 1
    End If

    If (EntryPrice <> kNullData) And (Position <> kNullData) And (Position <> 0) Then
        dMinValue = BarsCopy.LowestLow
        
        If dMinValue = kNullData Then
            dMinValue = EntryPrice
        End If
        
        HighestHighInTrade = MaxDouble(EntryPrice, BarsCopy.HighestHigh)
        LowestLowInTrade = MinDouble(EntryPrice, dMinValue)
        
        If Position > 0 Then
            MaxPositionProfit = g.Profit.Profit(BarsCopy.Prop(eBARS_Symbol), HighestHighInTrade - EntryPrice, , , , , m.lAccountID)
            MaxPositionLoss = g.Profit.Profit(BarsCopy.Prop(eBARS_Symbol), EntryPrice - LowestLowInTrade, , , , , m.lAccountID)
        Else
            MaxPositionProfit = g.Profit.Profit(BarsCopy.Prop(eBARS_Symbol), EntryPrice - LowestLowInTrade, , , , , m.lAccountID)
            MaxPositionLoss = g.Profit.Profit(BarsCopy.Prop(eBARS_Symbol), HighestHighInTrade - EntryPrice, , , , , m.lAccountID)
        End If
    Else
        HighestHighInTrade = kNullData
        LowestLowInTrade = kNullData
        MaxPositionLoss = kNullData
        MaxPositionProfit = kNullData
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cLateCalculating.UpdateHHLL"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateBarsSince
'' Description: Update the BarsSince functions
'' Inputs:      Bars
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdateBarsSince(ByVal Bars As cGdBars)
On Error GoTo ErrSection:

    BarsSinceEntry = CalcBarsSinceDate(EntryDateTime, Bars, True)
    BarsSinceLastEntry = CalcBarsSinceDate(LastEntryDateTime, Bars, True)
    BarsSinceExit = CalcBarsSinceDate(ExitDateTime, Bars, False)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cLateCalculating.UpdateBarsSince"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BarsSinceValid
'' Description: Are the BarsSince functions in a valid state?
'' Inputs:      None
'' Returns:     True if valid, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function BarsSinceValid() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = True
    If (Not m.Fills.FirstEntry Is Nothing) And (BarsSinceEntry = kNullData) Then
        bReturn = False
    ElseIf (Not m.Fills.LastEntry Is Nothing) And (BarsSinceLastEntry = kNullData) Then
        bReturn = False
    ElseIf (Not m.Fills.LastExit Is Nothing) And (BarsSinceExit = kNullData) Then
        bReturn = False
    End If
    
    BarsSinceValid = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cLateCalculating.BarsSinceValid"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MakeCopy
'' Description: Make a copy of the exit strategy
'' Inputs:      None
'' Returns:     New Copy
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MakeCopy() As cLateCalculating
On Error GoTo ErrSection:

    Dim aCopy As New cLateCalculating
    
    aCopy.SetPrivateData m
    Set MakeCopy = aCopy

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cLateCalculating.MakeCopy"
    
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

    m = mCopyFrom
    
    Set m.Tokens = mCopyFrom.Tokens.MakeCopy
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cLateCalculating.SetPrivateData"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Item
'' Description: Get an item from the collection
'' Inputs:      Index or Key
'' Returns:     Token (Nothing if not found)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function Item(ByVal vIndexOrKey As Variant) As cGdArray
On Error GoTo ErrSection:

    Dim tokReturn As cGdArray           ' Token to return
    
    Set tokReturn = Nothing
    If m.Tokens.Exists(vIndexOrKey) = True Then
        Set tokReturn = m.Tokens(vIndexOrKey)
    End If
    
    Set Item = tokReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cLateCalculating.Item"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetValue
'' Description: Get a value out of the collection
'' Inputs:      Name
'' Returns:     Value
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetValue(ByVal strName As String) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function

    strReturn = ""
    If m.Tokens.Exists(strName) = True Then
        strReturn = Item(strName)(TokenField(eGDToken_Value))
    End If
    
    GetValue = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cLateCalculating.GetValue"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetValue
'' Description: Set a value in the collection
'' Inputs:      Name, Value
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetValue(ByVal strName As String, ByVal strValue As String)
On Error GoTo ErrSection:

    If m.Tokens.Exists(strName) = True Then
        Item(strName)(TokenField(eGDToken_Value)) = strValue
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cLateCalculating.SetValue"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CodedTextForToken
'' Description: Create the coded text for the given function name
'' Inputs:      Token
'' Returns:     Coded Text
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CodedTextForToken(ByVal astrToken As cGdArray) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim strCodedName As String          ' Coded text name for the function
    
    strReturn = ""
    If astrToken.Size > 0 Then
        strCodedName = astrToken(TokenField(eGDToken_CodedName))
        strReturn = "~" & astrToken(TokenField(eGDToken_Type)) & Format(Len(strCodedName), "000") & strCodedName
        If astrToken(TokenField(eGDToken_Type)) = "02" Then
            strReturn = strReturn & " ~16001( ~07007Market1 ~17001)"
        Else
            strReturn = strReturn & " ~16001( ~17001)"
        End If
    End If
    
    CodedTextForToken = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cLateCalculating.CodedTextForToken"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ReplacementValue
'' Description: Coded text for the replacement value for the function
'' Inputs:      Token
'' Returns:     Replacement
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ReplacementValue(ByVal astrToken As cGdArray) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim strValue As String              ' Value from the token
    
    strReturn = ""
    If astrToken.Size > 0 Then
        strValue = astrToken(TokenField(eGDToken_Value))
        Select Case astrToken(TokenField(eGDToken_Type))
            Case "02"
                If UCase(strValue) = "TRUE" Then
                    strReturn = "~03004True ~16001( ~17001)"
                ElseIf UCase(strValue) = "FALSE" Then
                    strReturn = "~03005False ~16001( ~17001)"
                End If
            Case "04"
                strReturn = "~13" & Format(Len(strValue), "000") & strValue
            Case "30"
                strReturn = "~20" & Format(Len(strValue), "000") & strValue
        End Select
    End If
    
    ReplacementValue = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cLateCalculating.ReplacementValue"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NewToken
'' Description: Create a new token with the given information
'' Inputs:      Type, Name, Coded Name, Default value
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub NewToken(ByVal strFunctionType As String, ByVal strFunctionName As String, ByVal strCodedName As String, ByVal strDefaultValue As String)
On Error GoTo ErrSection:

    Dim astrToken As New cGdArray       ' Array of token information
    
    astrToken.Create eGDARRAY_Strings, TokenField(eGDToken_NumFields)
    astrToken(TokenField(eGDToken_Type)) = strFunctionType
    astrToken(TokenField(eGDToken_Name)) = strFunctionName
    astrToken(TokenField(eGDToken_CodedName)) = strCodedName
    astrToken(TokenField(eGDToken_Value)) = strDefaultValue

    m.Tokens.Add astrToken, strFunctionName

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cLateCalculating.NewToken"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadTokens
'' Description: Load up the tokens
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadTokens()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim fnFunction As cFunction         ' Function
    
    m.Tokens.Clear
    For lIndex = 1 To g.Functions.Count
        Set fnFunction = g.Functions.Item(lIndex)
        
        If fnFunction.LateCalculating = True Then
            If IsGenesisLibrary(fnFunction.LibraryID) = True Then
                Select Case fnFunction.ReturnTypeID
                    Case 1:
                        NewToken "04", fnFunction.FunctionName, fnFunction.CodedName, Str(kNullData)
                    Case 2:
                        NewToken "30", fnFunction.FunctionName, fnFunction.CodedName, ""
                    Case 3:
                        NewToken "02", fnFunction.FunctionName, fnFunction.CodedName, "False"
                End Select
            End If
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cLateCalculating.LoadTokens"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CalcBarsSinceDate
'' Description: Calculate the bars since a given date
'' Inputs:      Date/Time, Bars, BarsSinceEntry?
'' Returns:     Bars Since Date (Null if not found)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CalcBarsSinceDate(ByVal dDateTime As Double, ByVal Bars As cGdBars, ByVal bAllowNegativeOne As Boolean) As Long
On Error GoTo ErrSection:

    Dim lBarNumber As Long              ' Bar number for the date/time
    Dim dReturn As Double               ' Return value for the function
    Dim dSearch As Double               ' Date time to search for
    Dim dBarsDateTime As Double         ' Date time converted to Bars date/time
    
    dReturn = kNullData
    
    If dDateTime <> kNullData Then
        If Not Bars Is Nothing Then
            dBarsDateTime = ConvertTimeZone(dDateTime, "NY", Bars.Prop(eBARS_ExchangeTimeZoneInf))
            
            If Bars.IsIntraday Then
                dSearch = dBarsDateTime
            Else
                dSearch = Bars.SessionDateForTradeTime(dBarsDateTime)
            End If
        
            lBarNumber = Bars.FindDateTime(dSearch)
            If (lBarNumber < Bars.Size) Then
                If m.bLastBarGood Then
                    dReturn = (Bars.Size - 1) - lBarNumber
                ElseIf (lBarNumber < Bars.Size - 1) Then
                    dReturn = (Bars.Size - 2) - lBarNumber
                ElseIf bAllowNegativeOne Then
                    dReturn = -1#
                End If
            ElseIf bAllowNegativeOne Then
                dReturn = -1#
            End If
        End If
    End If
    
    CalcBarsSinceDate = dReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cLateCalculating.CalcBarsSinceDate"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CalcFirstEntry
'' Description: Calculate statistics for the first entry fill
'' Inputs:      First Entry Fill, Bars
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CalcFirstEntry(ByVal FirstEntry As cPtFill, ByVal Bars As cGdBars)
On Error GoTo ErrSection:

    Dim dFillDate As Double             ' Converted date of the fill

    If FirstEntry Is Nothing Then
        EntryPrice = kNullData
        EntryDate = kNullData
        EntrySignal = ""
        EntryTime = kNullData
        EntryDateTime = kNullData
        
        m.lAccountID = 0
    Else
        dFillDate = FirstEntry.FillDateConverted("NY")
        
        EntryPrice = FirstEntry.Price
        EntryDate = CLng(Val(Format(dFillDate, "YYYYMMDD")))
        EntrySignal = FirstEntry.BrokerID
        EntryTime = CLng(Val(Format(dFillDate, "HHMM")))
        EntryDateTime = dFillDate
        
        m.lAccountID = FirstEntry.AccountID
    End If

    BarsSinceEntry = CalcBarsSinceDate(EntryDateTime, Bars, True)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cLateCalculating.CalcFirstEntry"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CalcLastEntry
'' Description: Calculate statistics for the last entry fill
'' Inputs:      Last Entry Fill, Bars
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CalcLastEntry(ByVal LastEntry As cPtFill, ByVal Bars As cGdBars)
On Error GoTo ErrSection:

    If LastEntry Is Nothing Then
        LastEntryPrice = kNullData
        LastEntrySignal = ""
        LastEntryDateTime = kNullData
    Else
        LastEntryPrice = LastEntry.Price
        LastEntrySignal = LastEntry.BrokerID
        LastEntryDateTime = LastEntry.FillDateConverted("NY")
    End If

    BarsSinceLastEntry = CalcBarsSinceDate(LastEntryDateTime, Bars, True)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cLateCalculating.CalcLastEntry"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CalcLastExit
'' Description: Calculate statistics for the last exit fill
'' Inputs:      Last Exit Fill, Bars, Position
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CalcLastExit(ByVal LastExit As cPtFill, ByVal Bars As cGdBars, ByVal lPosition As Long)
On Error GoTo ErrSection:

    If LastExit Is Nothing Then
        ExitSignal = ""
        ExitProfit = kNullData
        ExitDateTime = kNullData
        ExitPrice = kNullData
        ExitPosition = kNullData
    Else
        ExitSignal = LastExit.BrokerID
        ExitProfit = LastExit.ClosedProfitCategory
        ExitDateTime = LastExit.FillDateConverted("NY")
        ExitPrice = LastExit.Price
        If LastExit.Buy Then
            ExitPosition = Int((lPosition - LastExit.Quantity) / m.lLotSize)
        Else
            ExitPosition = Int((lPosition + LastExit.Quantity) / m.lLotSize)
        End If
    End If

    BarsSinceExit = CalcBarsSinceDate(ExitDateTime, Bars, False)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cLateCalculating.CalcLastExit"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateFromFills
'' Description: Update statistics from the fills collection
'' Inputs:      Bars
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UpdateFromFills(Optional ByVal Bars As cGdBars = Nothing)
On Error GoTo ErrSection:

    Position = Int(m.Fills.Position / m.lLotSize)
    If Position = 0 Then
        AverageEntryPrice = kNullData
    Else
        AverageEntryPrice = m.Fills.AverageEntry
    End If
    
    CalcFirstEntry m.Fills.FirstEntry, Bars
    CalcLastEntry m.Fills.LastEntry, Bars
    If m.Fills.LastExit Is Nothing Then
        CalcLastExit Nothing, Bars, 0&
    Else
        CalcLastExit m.Fills.LastExit, Bars, m.Fills.LastExit.CurrentPositionCategory
    End If
    
    UpdateBarsInfo Bars

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cLateCalculating.UpdateFromFills"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Setup the class members when the object is intialized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:

    Set m.Tokens = New cGdTree
    Set m.Fills = New cLateCalculatingFills
    m.lLotSize = 1&
    m.bLastBarGood = False
    m.lAccountID = 0
    
    LoadTokens
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cLateCalculating.Class_Initialize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Setup the class members when the object is terminated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

    Set m.Tokens = Nothing
    Set m.Fills = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cLateCalculating.Class_Terminate"
    
End Sub
