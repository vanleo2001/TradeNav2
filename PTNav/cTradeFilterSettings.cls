VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTradeFilterSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cTradeFilterSettings.cls
'' Description: Settings for the trade report filter
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 05/20/2015   DAJ         Allow multiple accounts for the trade report filter
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    bUseDateRange As Boolean            ' Filter the trade report by a date range?
    dFromDate As Double                 ' Only include trades after this date
    dToDate As Double                   ' Only include trades before this date
    bUseAccount As Boolean              ' Filter the trade report by account?
    alAccountIds As cGdArray            ' Only include trades from these accounts
    bUseSymbol As Boolean               ' Filter the trade report by symbol?
    strSymbol As String                 ' Only include trades from this symbol
    nDirection As eGDFilterDirection    ' Only include trades in this direction
    bUseEntryRule As Boolean            ' Filter the trade report by entry rule?
    lEntryRuleID As Long                ' Only include trades with this entry rule
    bUseExitRule As Boolean             ' Filter the trade report by exit rule?
    lExitRuleID As Long                 ' Only include trades with this exit rule
    nTradeType As eGDFilterTradeType    ' Only include trades with this trade type
    bCalcPnl As Boolean                 ' Recalculate profit and loss for the trades?
    bUseAutoTrade As Boolean            ' Filter the trade report by auto trade?
    lAutoTradeID As Long                ' Only include trades for this automated trading item
End Type
Private m As mPrivate

Public Property Get UseDateRange() As Boolean
    UseDateRange = m.bUseDateRange
End Property
Public Property Let UseDateRange(ByVal bUseDateRange As Boolean)
    m.bUseDateRange = bUseDateRange
End Property

Public Property Get FromDate() As Double
    FromDate = m.dFromDate
End Property
Public Property Let FromDate(ByVal dFromDate As Double)
    m.dFromDate = dFromDate
End Property

Public Property Get ToDate() As Double
    ToDate = m.dToDate
End Property
Public Property Let ToDate(ByVal dToDate As Double)
    m.dToDate = dToDate
End Property

Public Property Get UseAccount() As Boolean
    UseAccount = m.bUseAccount
End Property
Public Property Let UseAccount(ByVal bUseAccount As Boolean)
    m.bUseAccount = bUseAccount
End Property

Public Property Get AccountIds() As cGdArray
    Set AccountIds = m.alAccountIds
End Property
Public Property Let AccountIds(ByVal alAccountIds As cGdArray)
    Set m.alAccountIds = alAccountIds
End Property

Public Property Get UseSymbol() As Boolean
    UseSymbol = m.bUseSymbol
End Property
Public Property Let UseSymbol(ByVal bUseSymbol As Boolean)
    m.bUseSymbol = bUseSymbol
End Property

Public Property Get Symbol() As String
    Symbol = m.strSymbol
End Property
Public Property Let Symbol(ByVal strSymbol As String)
    m.strSymbol = strSymbol
End Property

Public Property Get Direction() As eGDFilterDirection
    Direction = m.nDirection
End Property
Public Property Let Direction(ByVal nDirection As eGDFilterDirection)
    m.nDirection = nDirection
End Property

Public Property Get UseEntryRule() As Boolean
    UseEntryRule = m.bUseEntryRule
End Property
Public Property Let UseEntryRule(ByVal bUseEntryRule As Boolean)
    m.bUseEntryRule = bUseEntryRule
End Property

Public Property Get EntryRuleID() As Long
    EntryRuleID = m.lEntryRuleID
End Property
Public Property Let EntryRuleID(ByVal lEntryRuleID As Long)
    m.lEntryRuleID = lEntryRuleID
End Property

Public Property Get UseExitRule() As Boolean
    UseExitRule = m.bUseExitRule
End Property
Public Property Let UseExitRule(ByVal bUseExitRule As Boolean)
    m.bUseExitRule = bUseExitRule
End Property

Public Property Get ExitRuleID() As Long
    ExitRuleID = m.lExitRuleID
End Property
Public Property Let ExitRuleID(ByVal lExitRuleID As Long)
    m.lExitRuleID = lExitRuleID
End Property

Public Property Get TradeType() As eGDFilterTradeType
    TradeType = m.nTradeType
End Property
Public Property Let TradeType(ByVal nTradeType As eGDFilterTradeType)
    m.nTradeType = nTradeType
End Property

Public Property Get CalcPnl() As Boolean
    CalcPnl = m.bCalcPnl
End Property
Public Property Let CalcPnl(ByVal bCalcPnl As Boolean)
    m.bCalcPnl = bCalcPnl
End Property

Public Property Get UseAutoTrade() As Boolean
    UseAutoTrade = m.bUseAutoTrade
End Property
Public Property Let UseAutoTrade(ByVal bUseAutoTrade As Boolean)
    m.bUseAutoTrade = bUseAutoTrade
End Property

Public Property Get AutoTradeID() As Long
    AutoTradeID = m.lAutoTradeID
End Property
Public Property Let AutoTradeID(ByVal lAutoTradeID As Long)
    m.lAutoTradeID = lAutoTradeID
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadFromIni
'' Description: Load the settings from the INI file
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub LoadFromIni()
On Error GoTo ErrSection:

    Dim strSettings As String           ' Settings from the INI file

    strSettings = GetIniFileProperty("Last", "", "TradeReportFilter", g.strIniFile)
    FromString strSettings, ";", "|"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeFilterSettings.LoadFromIni"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SaveToIni
'' Description: Save the settings to the INI file
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SaveToIni()
On Error GoTo ErrSection:

    SetIniFileProperty "Last", ToString(";", "|"), "TradeReportFilter", g.strIniFile

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeFilterSettings.SaveToIni"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ToString
'' Description: Build a string out of the settings
'' Inputs:      Delimiter
'' Returns:     Delimited string, Delimiter for Accounts list
''
'' Fields:      chkDateRange, From Date, To Date, chkAccount, Account ID, chkSymbol,
''              Symbol, Direction, chkEntryRule, Entry Rule ID, chkExitRule
''              Exit Rule ID, RealSimFlag, chkCalcPnl, chkAutoTradeItem, Auto Trade Item

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ToString(Optional ByVal strDelimiter As String = vbTab, Optional ByVal strAccountsDelimiter As String = ",") As String
On Error GoTo ErrSection:

    Dim astrSettings As cGdArray        ' Array of settings
    
    Set astrSettings = New cGdArray
    astrSettings.Create eGDARRAY_Strings, 16
    
    astrSettings(0) = BoolToString(m.bUseDateRange)
    astrSettings(1) = Str(m.dFromDate)
    astrSettings(2) = Str(m.dToDate)
    astrSettings(3) = BoolToString(m.bUseAccount)
    astrSettings(4) = m.alAccountIds.JoinFields(strAccountsDelimiter)
    astrSettings(5) = BoolToString(m.bUseSymbol)
    astrSettings(6) = m.strSymbol
    astrSettings(7) = DirectionToString(m.nDirection)
    astrSettings(8) = BoolToString(m.bUseEntryRule)
    astrSettings(9) = Str(m.lEntryRuleID)
    astrSettings(10) = BoolToString(m.bUseExitRule)
    astrSettings(11) = Str(m.lExitRuleID)
    astrSettings(12) = TradeTypeToString(m.nTradeType)
    astrSettings(13) = BoolToString(m.bCalcPnl)
    astrSettings(14) = BoolToString(m.bUseAutoTrade)
    astrSettings(15) = Str(m.lAutoTradeID)
    
    ToString = astrSettings.JoinFields(strDelimiter)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeFilterSettings.ToString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FromString
'' Description: Set the settings based on the given string
'' Inputs:      Settings string, Delimiter, Delimiter for Accounts list
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FromString(ByVal strSettings As String, Optional ByVal strDelimiter As String = vbTab, Optional ByVal strAccountsDelimiter As String = ",")
On Error GoTo ErrSection:

    Dim astrSettings As cGdArray        ' Array of settings
    
    If Len(strSettings) > 0 Then
        Set astrSettings = New cGdArray
        astrSettings.SplitFields strSettings, strDelimiter
        
        m.bUseDateRange = BoolFromString(astrSettings(0))
        m.dFromDate = Val(astrSettings(1))
        m.dToDate = Val(astrSettings(2))
        m.bUseAccount = BoolFromString(astrSettings(3))
        m.alAccountIds.SplitFields astrSettings(4), strAccountsDelimiter
        m.bUseSymbol = BoolFromString(astrSettings(5))
        m.strSymbol = astrSettings(6)
        m.nDirection = DirectionFromString(astrSettings(7))
        m.bUseEntryRule = BoolFromString(astrSettings(8))
        m.lEntryRuleID = CLng(Val(astrSettings(9)))
        m.bUseExitRule = BoolFromString(astrSettings(10))
        m.lExitRuleID = CLng(Val(astrSettings(11)))
        m.nTradeType = TradeTypeFromString(astrSettings(12))
        m.bCalcPnl = BoolFromString(astrSettings(13))
        
        If astrSettings.Size > 14 Then
            m.bUseAutoTrade = BoolFromString(astrSettings(14))
            m.lAutoTradeID = CLng(Val(astrSettings(15)))
        Else
            m.bUseAutoTrade = False
            m.lAutoTradeID = 0&
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeFilterSettings.FromString"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HasFilter
'' Description: Are there any filters turned on?
'' Inputs:      Include Account in check?
'' Returns:     True if filter turned on, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function HasFilter(Optional ByVal bIncludeAccount As Boolean = True) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    If (m.bUseAccount = True) And (bIncludeAccount = True) Then
        bReturn = True
    ElseIf m.bUseAutoTrade = True Then
        bReturn = True
    ElseIf m.bUseDateRange = True Then
        bReturn = True
    ElseIf m.bUseEntryRule = True Then
        bReturn = True
    ElseIf m.bUseExitRule = True Then
        bReturn = True
    ElseIf m.bUseSymbol = True Then
        bReturn = True
    ElseIf m.nDirection <> eGDFilterDirection_All Then
        bReturn = True
    ElseIf m.nTradeType <> eGDFilterTradeType_All Then
        bReturn = True
    End If
    
    HasFilter = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeFilterSettings.HasFilter"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Description
'' Description: Build descriptive text for the options used
'' Inputs:      Include Account?, Delimiter, Trade Rules
'' Returns:     Description
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Description(Optional ByVal bIncludeAccount As Boolean = True, Optional ByVal strDelimiter As String = ";", Optional ByVal TradeRules As cTradeRules = Nothing) As String
On Error GoTo ErrSection:

    Dim astrDescription As cGdArray     ' Array of options
    Dim strReturn As String             ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim astrAccounts As cGdArray        ' List of account names
    
    Set astrDescription = New cGdArray
    astrDescription.Create eGDARRAY_Strings
    
    If TradeRules Is Nothing Then
        Set TradeRules = New cTradeRules
        TradeRules.Load
    End If
    
    If (m.bUseAccount = True) And (bIncludeAccount = True) Then
        Set astrAccounts = New cGdArray
        astrAccounts.Create eGDARRAY_Strings
        
        For lIndex = 0 To m.alAccountIds.Size - 1
            astrAccounts.Add g.Broker.AccountNameForID(m.alAccountIds(lIndex))
        Next lIndex
        
        astrDescription.Add "Account='" & astrAccounts.JoinFields(",") & "'"
    End If
    If m.bUseAutoTrade = True Then
        If m.lAutoTradeID = 0& Then
            astrDescription.Add "Category='Manual Trades'"
        Else
            astrDescription.Add "Category='" & g.TradingItems.NameForID(m.lAutoTradeID) & "'"
        End If
    End If
    If m.bUseDateRange = True Then
        astrDescription.Add "Date Range=" & DateFormat(m.dFromDate, MM_DD_YYYY) & " to " & DateFormat(m.dToDate, MM_DD_YYYY)
    End If
    If m.nDirection = eGDFilterDirection_Longs Then
        astrDescription.Add "Direction='Longs Only'"
    ElseIf m.nDirection = eGDFilterDirection_Shorts Then
        astrDescription.Add "Direction='Shorts Only'"
    End If
    If m.bUseEntryRule = True Then
        astrDescription.Add "Entry Rule='" & TradeRules.EntryRuleForID(m.lEntryRuleID) & "'"
    End If
    If m.bUseExitRule = True Then
        astrDescription.Add "Exit Rule='" & TradeRules.ExitRuleForID(m.lExitRuleID) & "'"
    End If
    If m.bUseSymbol = True Then
        astrDescription.Add "Symbol='" & m.strSymbol & "'"
    End If
    If m.nTradeType = eGDFilterTradeType_Real Then
        astrDescription.Add "Type='Real Trades Only'"
    ElseIf m.nTradeType = eGDFilterTradeType_Sim Then
        astrDescription.Add "Type='Sim Trades Only'"
    End If
    If astrDescription.Size = 0 Then
        strReturn = "None"
    Else
        strReturn = astrDescription.JoinFields(strDelimiter)
    End If
    
    Description = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeFilterSettings.Description"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IncludeAccount
'' Description: Determine if the given account ID should be included
'' Inputs:      Account ID
'' Returns:     True if should be included, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IncludeAccount(ByVal lAccountID As Long) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    
    bReturn = False
    For lIndex = 0 To m.alAccountIds.Size - 1
        If m.alAccountIds(lIndex) = lAccountID Then
            bReturn = True
            Exit For
        End If
    Next lIndex
    
    IncludeAccount = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeFilterSettings.IncludeAccount"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BoolToString
'' Description: Convert a boolean to a string
'' Inputs:      Boolean Value
'' Returns:     String equivalent
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function BoolToString(ByVal bValue As Boolean) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function

    If bValue Then
        strReturn = Str(vbChecked)
    Else
        strReturn = Str(vbUnchecked)
    End If
    
    BoolToString = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeFilterSettings.BoolToString"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BoolFromString
'' Description: Convert a string to a boolean
'' Inputs:      String Value
'' Returns:     Boolean equivalent
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function BoolFromString(ByVal strValue As String) As Boolean
On Error GoTo ErrSection:

    BoolFromString = (strValue = Str(vbChecked))

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeFilterSettings.BoolFromString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DirectionToString
'' Description: Convert a direction to a string
'' Inputs:      Direction Value
'' Returns:     String equivalent
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DirectionToString(ByVal nDirection As eGDFilterDirection) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    Select Case nDirection
        Case eGDFilterDirection_All
            strReturn = "B"
        Case eGDFilterDirection_Longs
            strReturn = "L"
        Case eGDFilterDirection_Shorts
            strReturn = "S"
    End Select
    
    DirectionToString = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeFilterSettings.DirectionToString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DirectionFromString
'' Description: Convert a string to a direction
'' Inputs:      String Value
'' Returns:     Direction equivalent
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DirectionFromString(ByVal strValue As String) As eGDFilterDirection
On Error GoTo ErrSection:

    Dim nReturn As eGDFilterDirection   ' Return value for the function

    Select Case strValue
        Case "L"
            nReturn = eGDFilterDirection_Longs
        Case "S"
            nReturn = eGDFilterDirection_Shorts
        Case Else
            nReturn = eGDFilterDirection_All
    End Select
    
    DirectionFromString = nReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeFilterSettings.DirectionFromString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TradeTypeToString
'' Description: Convert a Trade Type to a string
'' Inputs:      Trade Type Value
'' Returns:     String equivalent
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function TradeTypeToString(ByVal nTradeType As eGDFilterTradeType) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    Select Case nTradeType
        Case eGDFilterTradeType_All
            strReturn = "A"
        Case eGDFilterTradeType_Real
            strReturn = "R"
        Case eGDFilterTradeType_Sim
            strReturn = "S"
    End Select
    
    TradeTypeToString = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeFilterSettings.TradeTypeToString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TradeTypeFromString
'' Description: Convert a string to a Trade Type
'' Inputs:      String Value
'' Returns:     Trade Type equivalent
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function TradeTypeFromString(ByVal strValue As String) As eGDFilterTradeType
On Error GoTo ErrSection:

    Dim nReturn As eGDFilterTradeType   ' Return value for the function

    Select Case strValue
        Case "R"
            nReturn = eGDFilterTradeType_Real
        Case "S"
            nReturn = eGDFilterTradeType_Sim
        Case Else
            nReturn = eGDFilterTradeType_All
    End Select
    
    TradeTypeFromString = nReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeFilterSettings.TradeTypeFromString"
    
End Function

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

    UseDateRange = False
    FromDate = 0#
    ToDate = 0#
    UseAccount = False
    Set m.alAccountIds = New cGdArray
    m.alAccountIds.Create eGDARRAY_Longs
    UseSymbol = False
    Symbol = ""
    Direction = eGDFilterDirection_All
    UseEntryRule = False
    EntryRuleID = 0&
    UseExitRule = False
    ExitRuleID = 0&
    TradeType = eGDFilterTradeType_All
    CalcPnl = False
    UseAutoTrade = False
    AutoTradeID = 0&

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeFilterSettings.Class_Initialize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Terminate the class members when the object is terminated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

    Set m.alAccountIds = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeFilterSettings.Class_Terminate"
    
End Sub

