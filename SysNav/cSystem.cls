VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cSystem.cls
'' Description: Wrapper class for a strategy
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 04/17/2009   DAJ         If not pyramiding, pass hard coded values for rules
'' 05/06/2009   DAJ         Error in auto trade orders if no data in the bars
'' 03/15/2010   DAJ         Error out in RunEngine if there is no data in the bars
'' 03/22/2010   DAJ         Check for Bars being Nothing before checking if they are empty
'' 05/13/2010   DAJ         Change tick value for Forex symbols for profit calculations
'' 08/10/2010   DAJ         Check if last two bars have same date/time before running engine
'' 06/16/2011   DAJ         Added code for the Highlight Bar Reporter
'' 11/03/2011   DAJ         Don't double apply tick value adjustment when running system
'' 09/06/2012   DAJ         Handle error conditions better in RunEngine, AutoTradeOrders
'' 01/22/2013   DAJ         Return Rule Based report in AutoTradeOrders
'' 04/03/2013   DAJ         Automated Strategy Baskets
'' 04/09/2013   DAJ         Apply lot size to tick value for back-testing Forex
'' 05/01/2013   DAJ         Shadow Trading
'' 07/25/2013   DAJ         Have LoadForMultRun take the ToDate
'' 02/27/2014   DAJ         For the period between LDD and midnight, make sure to splice bars
'' 03/19/2014   DAJ         Fix for running a system with 'Of ",<period>"' on a symbol group
'' 06/26/2014   DAJ         Added the HasOnCloseRules function; Only do backtest on OnClose
''                          automated trading check
'' 08/08/2014   DAJ         Fix for starting date for a backtest if pyramiding
'' 06/12/2015   DAJ         Added flag to add symbol to stream or not when loading Market1 bars
'' 07/08/2015   DAJ         Removed the #If ISIDE block in AutoTradeOrders
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Compare Text
Option Explicit
'This class controls the gathering System information required for testing.
'A system can have collection of Rules, Inputs, and Markets.  This class
'handles the maintenance of this information.

Public Enum eGDNextBarMode
    eGDNextBarMode_Normal = 0
    eGDNextBarMode_UseHwnd
    eGDNextBarMode_RunMult
End Enum

Public Enum gdRunOption
    groOptimize
    groAlternate
    groBoth
    groSingleIteration
    groNextBar
    groCurrentBarOrders
End Enum

Public Enum eGDRunMode
    eGDRunMode_All = 0
    eGDRunMode_Short
    eGDRunMode_Long
End Enum

Public Enum eGDAutoTradeOrdersMode
    eGDAutoTradeOrdersMode_Normal = 0
    eGDAutoTradeOrdersMode_NewBar
    eGDAutoTradeOrdersMode_OnClose
End Enum

Private Type mPrivate
    Rules As cRules
    Markets As cGdTree

    lSystemNumber As Long
    strSystemName As String
    dFromDate As Date
    dToDate As Date
    bToEndOfData As Boolean
    strBarTimeFrame As String
    strNotes As String
    strDeveloper As String
    strTradesPath As String
    Expenses As Currency
    lLibraryID As Long
    strPassword As String
    nSecurityLevel As Byte
    bCannotDelete As Boolean
    bReverify As Boolean
    lMMid As Long
    dLastModified As Date
    lBarsLoadedBeforeTrading As Long
    lBarsTradedBeforeOrders As Long
    bPyramid As Boolean
    lTradeDepth As Long
    bCanPyramidSameEntry As Boolean
    bLinkInputs As Boolean
    bUseSharesPerTrade As Boolean
    lNumShares As Long
    dDollarsPerTrade As Currency
    dStockExpenses As Currency
    dForexExpenses As Currency
    bAllowReverse As Boolean
    bForceLimitThrough As Boolean
    nErrNbr As Byte
    bIsGuru As Boolean
    
    strOverrides As String

'    bForceChartNavData As Boolean
'    bAllChartNavData As Boolean
    
    hWnd As Long
    Bars As cGdBars
    PrevBars As cGdBars
    RptBridges As cGdTree
    astrSingleIterationParms As cGdArray
    
    nRunMode As eGDRunMode
    lTestRuleID As Long
    bHighlightBarReport As Boolean
    strHighlightBarInfo As String
    
    tblSystem As cGdTable
    tblSystemRules As cGdTable
    tblSystemParms As cGdTable
    tblRptRules As cGdTable
    tblRptMarkets As cGdTable
End Type
Private m As mPrivate

Private Enum etblSystem
    etblSystem_ID = 0
    etblSystem_Name
    etblSystem_Handle
    etblSystem_Expenses
    etblSystem_StockExpenses
    etblSystem_TradeDepth
    etblSystem_IfPyramid
    etblSystem_CanPyramidSameEntry
    etblSystem_SystemScope
    etblSystem_DoShuffle
    etblSystem_AllowReverse
    etblSystem_StockQty
    etblSystem_NumFields
End Enum

Private Enum etblSystemRule
    etblSystemRule_ID = 0
    etblSystemRule_SysID
    etblSystemRule_Name
    etblSystemRule_CondCoded
    etblSystemRule_PriceCoded
    etblSystemRule_LimitCoded
    etblSystemRule_OrderPlacement
    etblSystemRule_BuySell
    etblSystemRule_Selected
    etblSystemRule_Alternate
    etblSystemRule_Use
    etblSystemRule_ExitOnEntryBar
    etblSystemRule_ExitOnEachTrade
    etblSystemRule_NumberContracts
    etblSystemRule_AsPercentPos
    etblSystemRule_LinkedRules
    etblSystemRule_Sequence
    etblSystemRule_NumFields
End Enum

Private Enum etblSystemParm
    etblSystemParm_ID = 0
    etblSystemParm_Name
    etblSystemParm_SysID
    etblSystemParm_RuleID
    etblSystemParm_TypeID
    etblSystemParm_Value
    etblSystemParm_OptFromValue
    etblSystemParm_OptToValue
    etblSystemParm_OptStepValue
    etblSystemParm_IfOptimize
    etblSystemParm_NumFields
End Enum

Private Enum etblRptRules
    etblRptRules_RuleID = 0             ' Long
    etblRptRules_OrderTypeID            ' Short
    etblRptRules_RuleName               ' String
    etblRptRules_NumFields
End Enum

Private Enum etblRptMarkets
    etblRptMarkets_Symbol = 0           ' String
    etblRptMarkets_LocalStart           ' Double
    etblRptMarkets_LocalEnd             ' Double
End Enum

Private Function SystemField(ByVal SysFld As etblSystem) As Long
    SystemField = SysFld
End Function
Private Function SystemRuleField(ByVal SysRuleFld As etblSystemRule) As Long
    SystemRuleField = SysRuleFld
End Function
Private Function SystemParmField(ByVal SystemParmFld As etblSystemParm) As Long
    SystemParmField = SystemParmFld
End Function
Private Function RptRulesField(ByVal RptRulesFld As etblRptRules) As Long
    RptRulesField = RptRulesFld
End Function
Private Function RptMarketsField(ByVal RptMarketsFld As etblRptMarkets) As Long
    RptMarketsField = RptMarketsFld
End Function

Property Get ErrNbr() As Byte
    ErrNbr = m.nErrNbr
End Property
Property Get Rules() As cRules
    Set Rules = m.Rules
End Property
Public Property Get Markets() As cGdTree
    Set Markets = m.Markets
End Property
Property Let Rules(pData As cRules)
    Set m.Rules = pData
End Property
Public Property Let Markets(pData As cGdTree)
    Set m.Markets = pData
End Property
Property Get SystemNumber() As Long
    SystemNumber = m.lSystemNumber
End Property
Property Get SystemName() As String
    SystemName = m.strSystemName
End Property
Property Get FromDate() As Date
    FromDate = m.dFromDate
End Property
Property Get ToDate() As Date
    ToDate = m.dToDate
End Property
Property Get BarTimeFrame() As String
    BarTimeFrame = m.strBarTimeFrame
End Property
Property Get ToEndOfData() As Boolean
    ToEndOfData = m.bToEndOfData
End Property
Property Get TradesPath() As String
    TradesPath = m.strTradesPath
End Property
Property Get Expenses() As Currency
    Expenses = m.Expenses
End Property
Property Get LibraryID() As Long
    LibraryID = m.lLibraryID
End Property
Property Get Developer() As String
    Developer = m.strDeveloper
End Property
Property Get Notes() As String
    Notes = m.strNotes
End Property
Property Get SecurityLevel() As Integer
    SecurityLevel = m.nSecurityLevel
End Property
Property Get CannotDelete() As Boolean
    CannotDelete = m.bCannotDelete
End Property
Property Get Password() As String
    Password = m.strPassword
End Property
Property Get LastModified() As Date
    LastModified = m.dLastModified
End Property
Property Get Reverify() As Boolean
    Reverify = m.bReverify
End Property
Property Get MMid() As Long
    MMid = m.lMMid
End Property
Property Get BarsLoadedBeforeTrading() As Long
    BarsLoadedBeforeTrading = m.lBarsLoadedBeforeTrading
End Property
Property Get BarsTradedBeforeOrders() As Long
    BarsTradedBeforeOrders = m.lBarsTradedBeforeOrders
End Property
Property Get Pyramid() As Boolean
    Pyramid = m.bPyramid
End Property
Property Get TradeDepth() As Long
    TradeDepth = m.lTradeDepth
End Property
Property Get CanPyramidSameEntry() As Boolean
    CanPyramidSameEntry = m.bCanPyramidSameEntry
End Property
Property Get LinkInputs() As Boolean
    LinkInputs = m.bLinkInputs
End Property
Property Get RunMode() As eGDRunMode
    RunMode = m.nRunMode
End Property
Property Get AllowReverse() As Boolean
    AllowReverse = m.bAllowReverse
End Property
Property Get ForceLimitThrough() As Boolean
    ForceLimitThrough = m.bForceLimitThrough
End Property
Property Get Bars() As cGdBars
    Set Bars = m.Bars
End Property
Property Get TestRuleID() As Long
    TestRuleID = m.lTestRuleID
End Property
Property Get RptRulesHandle() As Long
    RptRulesHandle = m.tblRptRules.TableHandle
End Property
Property Get hWnd() As Long
    hWnd = m.hWnd
End Property

Property Let SystemNumber(pData As Long)
    m.lSystemNumber = pData
End Property
Property Let SystemName(pData As String)
    m.strSystemName = pData
End Property
Property Let FromDate(pData As Date)
    m.dFromDate = pData
End Property
Property Let ToDate(pData As Date)
    m.dToDate = pData
End Property
Property Let BarTimeFrame(pData As String)
    m.strBarTimeFrame = pData
End Property
Property Let ToEndOfData(pData As Boolean)
    m.bToEndOfData = pData
End Property
Property Let Notes(pData As String)
    m.strNotes = pData
End Property
Property Let Expenses(pData As Currency)
    m.Expenses = pData
End Property
Property Let TradesPath(pData As String)
    m.strTradesPath = pData
End Property
Property Let LibraryID(pData As Long)
    m.lLibraryID = pData
End Property
Property Let Developer(pData As String)
    m.strDeveloper = pData
End Property
Property Let SecurityLevel(pData As Integer)
    m.nSecurityLevel = pData
End Property
Property Let CannotDelete(pData As Boolean)
    m.bCannotDelete = pData
End Property
Property Let Password(pData As String)
    m.strPassword = pData
End Property
Property Let Reverify(pData As Boolean)
    m.bReverify = pData
End Property
Property Let MMid(pData As Long)
    m.lMMid = pData
End Property
Property Let BarsLoadedBeforeTrading(pData As Long)
    m.lBarsLoadedBeforeTrading = pData
End Property
Property Let BarsTradedBeforeOrders(pData As Long)
    m.lBarsTradedBeforeOrders = pData
End Property
Property Let Pyramid(pData As Boolean)
    m.bPyramid = pData
End Property
Property Let TradeDepth(pData As Long)
    m.lTradeDepth = pData
End Property
Property Let CanPyramidSameEntry(pData As Boolean)
    m.bCanPyramidSameEntry = pData
End Property
Property Let LinkInputs(pData As Boolean)
    m.bLinkInputs = pData
End Property
Property Let SingleIterationParms(pData As cGdArray)
    Set m.astrSingleIterationParms = pData
End Property
Property Get UseSharesPerTrade() As Boolean
    UseSharesPerTrade = m.bUseSharesPerTrade
End Property
Property Let UseSharesPerTrade(ByVal bUseSharesPerTrade As Boolean)
    m.bUseSharesPerTrade = bUseSharesPerTrade
End Property
Property Get NumShares() As Long
    If m.lNumShares = 0 Then m.lNumShares = 100
    NumShares = m.lNumShares
End Property
Property Let NumShares(ByVal lNumShares As Long)
    m.lNumShares = lNumShares
    If m.lNumShares = 0 Then m.lNumShares = 100
End Property
Property Get DollarsPerTrade() As Currency
    DollarsPerTrade = m.dDollarsPerTrade
End Property
Property Let DollarsPerTrade(ByVal dDollarsPerTrade As Currency)
    m.dDollarsPerTrade = dDollarsPerTrade
End Property
Property Get StockExpenses() As Currency
    StockExpenses = m.dStockExpenses
End Property
Property Let StockExpenses(ByVal dStockExpenses As Currency)
    m.dStockExpenses = dStockExpenses
End Property
Property Get ForexExpenses() As Currency
    ForexExpenses = m.dForexExpenses
End Property
Property Let ForexExpenses(ByVal dForexExpenses As Currency)
    m.dForexExpenses = dForexExpenses
End Property
Property Let RunMode(ByVal nMode As eGDRunMode)
    m.nRunMode = nMode
End Property
Property Let AllowReverse(ByVal bAllowReverse As Boolean)
    m.bAllowReverse = bAllowReverse
End Property
Property Let ForceLimitThrough(ByVal bForceLimitThrough As Boolean)
    m.bForceLimitThrough = bForceLimitThrough
End Property
Property Let Bars(NewBars As cGdBars)
    If NewBars Is Nothing Then
        Set m.Bars = Nothing
    Else
        Set m.Bars = NewBars.MakeCopy
        ChangeMarket1 m.Bars
    End If
End Property
Property Let hWnd(ByVal lHwnd As Long)
    m.hWnd = lHwnd
End Property

Public Property Get HighlightBarReport() As Boolean
    HighlightBarReport = m.bHighlightBarReport
End Property
Public Property Let HighlightBarReport(ByVal bHighlightBarReport As Boolean)
    m.bHighlightBarReport = bHighlightBarReport
End Property

Public Property Get HighlightBarInfo() As String
    HighlightBarInfo = m.strHighlightBarInfo
End Property
Public Property Let HighlightBarInfo(ByVal strHighlightBarInfo As String)
    m.strHighlightBarInfo = strHighlightBarInfo
End Property

Public Property Get Overrides() As String
    Overrides = m.strOverrides
End Property
Public Property Let Overrides(ByVal strOverrides As String)
    m.strOverrides = strOverrides
End Property

Public Property Get IsGuru() As Boolean
    IsGuru = m.bIsGuru
End Property
Public Property Let IsGuru(ByVal bIsGuru As Boolean)
    m.bIsGuru = bIsGuru
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Load
'' Description: Load the system object from the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Load(ByVal lSystemNumber As Long)
On Error GoTo ErrSection:
    
    Dim rs As Recordset                 ' Recordset from the database
    Dim lRule As Long
    Dim lParm As Long
    Dim Bars As New cGdBars
    Dim bFix As Boolean
    Dim Parms As cGdTree
    Dim astrToFix As cGdArray
    Dim lIndex As Long
    Dim astrOverrides As New cGdArray   ' Array of user overrides (if mult run)
    Dim lPos As Long                    ' Position found in overrides
    Dim bFound As Boolean               ' Was the pararmeter found in overrides?
    Dim strRuleName As String           ' Name of the current rule
    Dim astrLinked As New cGdArray      ' Array of linked rules
    Dim bChanged As Boolean             ' Did the linked rules list change?
    
    m.lSystemNumber = lSystemNumber
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblSystems] " & _
        "WHERE [SystemNumber]=" & lSystemNumber & ";", dbOpenDynaset)
    ValidateCheckSums rs, "tblSystems"
    
    ' Load the system from the database if found and valid
    If Not (rs.BOF And rs.EOF) Then
        rs.MoveFirst
        
        If rs!CheckSum <> 0.5 Then
            m.strSystemName = rs!SystemName
            m.lSystemNumber = rs!SystemNumber
            m.strBarTimeFrame = rs!BarTimeFrame
            m.Expenses = rs!Expenses
            m.lLibraryID = rs!LibraryID
            m.strDeveloper = rs!Developer
            m.lBarsLoadedBeforeTrading = rs!BarsLoadedBeforeTrading
            m.lBarsTradedBeforeOrders = rs!BarsTradedBeforeOrders
            m.strTradesPath = rs!TradesPath
            m.dFromDate = rs!FromDate
            m.dToDate = rs!ToDate
            m.bToEndOfData = rs!ToEndOfData
            m.dLastModified = rs!LastModified
            m.strNotes = rs!Notes
            m.bReverify = rs!Reverify
            m.lMMid = rs!MMid
            m.nSecurityLevel = NullChk(rs!SecurityLevel, 0)
            m.bCannotDelete = rs!CannotDelete
            m.strPassword = DecryptField(rs!Password)
            m.bPyramid = rs!Pyramid
            m.lTradeDepth = rs!TradeDepth
            m.bCanPyramidSameEntry = rs!CanPyramidSameEntry
            m.bLinkInputs = rs!LinkInputs
            m.bUseSharesPerTrade = rs!UseSharesPerTrade
            m.lNumShares = rs!NumShares
            If m.lNumShares = 0 Then m.lNumShares = 100
            m.dDollarsPerTrade = rs!DollarsPerTrade
            m.dStockExpenses = rs!StockExpenses
            m.dForexExpenses = rs!ForexExpenses
            m.bAllowReverse = rs!AllowReverse
            m.bForceLimitThrough = rs!ForceLimitThrough
            m.bIsGuru = rs!IsGuru
        Else
            Err.Raise vbObjectError + 1000, , "Strategy is Invalid"
        End If
    End If
    
    astrOverrides.SplitFields m.strOverrides, ","
    astrOverrides.Sort
    
    ' Load the Rules if necessary
    'If m.bLoadRules Then LoadRulesForSystem
    Set m.Rules = New cRules
    m.Rules.LoadSystemRules m.lSystemNumber
    
    ' Load up a collection of unique markets...
    Set m.Markets = New cGdTree
    Set Parms = New cGdTree
    Set astrToFix = New cGdArray
    astrToFix.Create eGDARRAY_Strings
    bFix = False

    For lRule = 1 To m.Rules.Count
        With m.Rules.Item(lRule)
            ' Make sure that all of the local rules have the same library and
            ' security information as the system does (DAJ: 3/24/2003)...
            If .LibraryID <> m.lLibraryID Then
                .LibraryID = m.lLibraryID
                .SecurityLevel = m.nSecurityLevel
                .CannotDelete = m.bCannotDelete
                .Password = m.strPassword
                m.bReverify = True
            End If
            
            ' Need to fix the System Rules version of the pyramid information
            ' if the number of contracts is zero (DAJ: 4/15/2004)...
            If .SysNumContracts = 0 Then
                .SysAsPercentOfPosition = .AsPercentOfPosition
                .SysExitBasedOnEachTrade = .ExitBasedOnEachTrade
                .SysNumContracts = .NumberContracts
            End If
            
            ' For some reason at some point a bunch of rules got the ExitBasedOnEachTrade
            ' turned on - make sure to turn them off here (DAJ: 4/22/2003)...
            If .ExitBasedOnEachTrade = True Then
                .ExitBasedOnEachTrade = False
                m.bReverify = True
            End If
            
            ' If a rule gets verified with a carriage return in the condition, the
            ' order placement can end up being blanked out, so we need to make the
            ' user go in and reverify it. (DAJ: 12/15/2003)...
            If Len(.OrderPlacement) = 0 Then
                .Reverify = True
                m.bReverify = True
            End If
            
            ' If there are linked rules associated with this rule, then remove any
            ' rules that no longer exist in this system.  If anything changes, then
            ' mark the system as dirty -- if there are no longer any linked rules,
            ' then turn the rule off for now (DAJ: 06/27/2005)...
            If Len(.LinkedRules) > 0 Then
                bChanged = False
                astrLinked.SplitFields .LinkedRules, ","
                For lIndex = astrLinked.Size - 1 To 0 Step -1
                    If m.Rules.Found(astrLinked(lIndex)) = False Then
                        astrLinked.Remove lIndex
                        bChanged = True
                    End If
                Next lIndex
                If bChanged Then
                    .LinkedRules = astrLinked.JoinFields(",")
                    If Len(.LinkedRules) = 0 Then
                        .Selected = False
                    End If
                    m.bReverify = True
                End If
            End If
        
            strRuleName = .Name
            
            For lParm = 1 To .Inputs.Count
                With .Inputs.Item(lParm)
                    If .SystemID = 0& Then
                        bFix = True
                        astrToFix.Add CStr(.RuleID) & vbTab & CStr(.ParmID)
                    End If
                    
                    If .ParmTypeID = 5 Then
                        If .SecurityName = "" And .Format = "CN" Then
                            If .SymbolID = 0 Then
                                If SetBarProperties(Bars, .Symbol) Then
                                    .SecurityName = Bars.Prop(eBARS_Desc)
                                End If
                            Else
                                If SetBarProperties(Bars, .SymbolID) Then
                                    .SecurityName = Bars.Prop(eBARS_Desc)
                                End If
                            End If
                        End If
                        If Not m.Markets.Exists(.ParmName) Then
                            m.Markets.Add m.Rules.Item(lRule).Inputs.Item(lParm), .ParmName
                        End If
                    Else
                        If m.bLinkInputs Then
                            bFound = astrOverrides.BinarySearch(.ParmName & "=", lPos, eGdSort_MatchUsingSearchStringLength)
                        Else
                            bFound = astrOverrides.BinarySearch(strRuleName & "." & .ParmName & "=", lPos, eGdSort_MatchUsingSearchStringLength)
                        End If
                        If bFound Then
                            .Value = Parse(astrOverrides(lPos), "=", 2)
                        End If
                    
                        If Not Parms.Exists(.ParmName) Then
                            Parms.Add m.Rules.Item(lRule).Inputs.Item(lParm), .ParmName
                        End If
                    End If
                End With
            Next lParm
        End With
    Next lRule
    
    ' For some reason, sometimes there are rule parms with no corresponding system
    ' rule records.  Try to fix these situations here (DAJ 03/24/2003)...
    If bFix Then
        For lIndex = 0 To astrToFix.Size - 1
            lRule = CLng(Parse(astrToFix(lIndex), vbTab, 1))
            lParm = CLng(Parse(astrToFix(lIndex), vbTab, 2))
            
            With m.Rules.Item(CStr(lRule)).Inputs.Item(CStr(lParm))
                .SystemID = m.lSystemNumber
                If .ParmTypeID = 5 Then
                    If m.Markets.Exists(.ParmName) Then
                        .Path = m.Markets(.ParmName).Path
                        .SymbolID = m.Markets(.ParmName).SymbolID
                        If .SymbolID = 0 Then
                            .Symbol = m.Markets(.ParmName).Symbol
                        Else
                            .Symbol = GetSymbol(.SymbolID)
                        End If
                        .MarketSymbol = m.Markets(.ParmName).MarketSymbol
                        .Periodicity = m.Markets(.ParmName).Periodicity
                        .Format = m.Markets(.ParmName).Format
                        .SecurityType = m.Markets(.ParmName).SecurityType
                        .SecurityName = m.Markets(.ParmName).SecurityName
                        .GroupID = m.Markets(.ParmName).GroupID
                    End If
                ElseIf m.bLinkInputs Then
                    If Parms.Exists(.ParmName) Then
                        .Value = Parms(.ParmName).Value
                        .IfOptimize = Parms(.ParmName).IfOptimize
                        .OptFromValue = Parms(.ParmName).OptFromValue
                        .OptToValue = Parms(.ParmName).OptToValue
                        .OptStepValue = Parms(.ParmName).OptStepValue
                    End If
                End If
            End With
        Next lIndex
        
        m.bReverify = True
    End If

ErrExit:
    Set Bars = Nothing
    Set rs = Nothing
    Exit Sub

ErrSection:
    Set Bars = Nothing
    Set rs = Nothing
    RaiseError "cSystem.Load", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Save
'' Description: Save the system to the database
'' Inputs:      Whether to use the Hwnd
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Save(Optional ByVal bUseHwnd As Boolean = False)
On Error GoTo ErrSection:
    
    Dim rs As Recordset                 ' Recordset from the database
    Dim bNewRec As Boolean              ' Is this a new record?
    Dim bTransaction As Boolean         ' Are we in a transaction?
'    Dim strSystems As String            ' List of systems to force hard drive data
    Dim lRule As Long
    Dim lInput As Long
    Dim lNewSystemID As Long
    Dim alOldRules As cGdArray          ' Array of rules that no longer exist in the system
    
    Set alOldRules = New cGdArray
    alOldRules.Create eGDARRAY_Longs
    
    ' Validate the system and components
    Validate
    
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblSystems] " & _
                "WHERE [SystemNumber]=" & m.lSystemNumber & ";", dbOpenDynaset)

    'Start a transaction
    'g.WrkJet.BeginTrans
    bTransaction = True
    
    If rs.EOF Then
        rs.AddNew
        m.lSystemNumber = rs!SystemNumber
        lNewSystemID = rs!SystemNumber
    Else
        rs.Edit
        lNewSystemID = 0&
    End If
    
    ' update LastModified (unless skip file exists -- e.g. when updating master mdb)
    If m.dLastModified = 0 Or Not FileExist(App.Path & "\lastmod.skp") Then
        m.dLastModified = Now()
    Else
        StatusMsg "LastModified not changed"
    End If
    
    rs!SystemName = m.strSystemName
    rs!LibraryID = m.lLibraryID
    rs!BarTimeFrame = m.strBarTimeFrame
    rs!Expenses = m.Expenses
    rs!Developer = m.strDeveloper
    rs!BarsLoadedBeforeTrading = m.lBarsLoadedBeforeTrading
    rs!BarsTradedBeforeOrders = m.lBarsTradedBeforeOrders
    rs!TradesPath = m.strTradesPath
    rs!FromDate = m.dFromDate
    rs!ToDate = m.dToDate
    rs!ToEndOfData = m.bToEndOfData
    rs!LastModified = m.dLastModified
    rs!Notes = m.strNotes
    rs!SecurityLevel = m.nSecurityLevel
    rs!CannotDelete = m.bCannotDelete
    EncryptField rs!Password, m.strPassword
    rs!Reverify = False
    rs!MMid = m.lMMid
    rs!Pyramid = m.bPyramid
    rs!TradeDepth = m.lTradeDepth
    rs!CanPyramidSameEntry = m.bCanPyramidSameEntry
    rs!LinkInputs = m.bLinkInputs
    rs!UseSharesPerTrade = m.bUseSharesPerTrade
    If m.lNumShares = 0 Then m.lNumShares = 100
    rs!NumShares = m.lNumShares
    rs!DollarsPerTrade = m.dDollarsPerTrade
    rs!StockExpenses = m.dStockExpenses
    rs!ForexExpenses = m.dForexExpenses
    rs!AllowReverse = m.bAllowReverse
    rs!ForceLimitThrough = m.bForceLimitThrough
    rs!IsGuru = m.bIsGuru
    
    rs!CheckSum = BuildCheckSum(rs, "tblSystems")
    rs.Update
    
    ' Save the existing rules...
    m.Rules.SaveSystemRules lNewSystemID
    For lRule = 1 To m.Rules.Count
        RefreshRule m.Rules.Item(lRule)
    Next lRule
    
    ' Check for rules that no longer exist in the system...
    For lRule = 1 To g.tblRule.NumRecords
        If g.tblRule(RuleField(etblRule_SystemNumber), lRule) = m.lSystemNumber Then
            If m.Rules.Found(Str(g.tblRule(RuleField(etblRule_RuleID), lRule))) = False Then
                alOldRules.Add g.tblRule(RuleField(etblRule_RuleID), lRule)
            End If
        End If
    Next lRule
    
    ' Delete any rules that no longer exist from the global rules table...
    For lRule = 0 To alOldRules.Size - 1
        DeleteRule alOldRules(lRule)
    Next lRule
    
    ' Let the toolbox know it needs to reload the rules grid...
    If FormIsLoaded("frmToolbox") Then frmToolbox.fgRules.Rows = frmToolbox.fgRules.FixedRows
    
    ' Save the system components
    'SaveRules bNewRec
    'SaveInputs
    'SaveSecurities
    
    ' Force the updates to flush to disk (so engine can read it)
    'g.WrkJet.CommitTrans dbForceOSFlush
    bTransaction = False

    ' Delete output files since now invalid with newly saved system.
    KillFile TradesFile(bUseHwnd)
    If bUseHwnd Then
        KillFile NextBarFile(eGDNextBarMode_UseHwnd)
    Else
        KillFile NextBarFile(eGDNextBarMode_Normal)
    End If
        
    'strSystems = GetIniFileProperty("HardDriveData", "", "Data", g.strIniFile)
    
    ' If the user has changed a security to be from the hard drive, add it to
    ' the list not to force a change to ChartNav data
    'If InStr(strSystems, "," & m.lSystemNumber & ",") = 0 Then
    '    If g.bHasChartNavData And m.bAllChartNavData = False Then
    '        SetIniFileProperty "HardDriveData", strSystems & "," & m.lSystemNumber & ",", "Data", g.strIniFile
    '    End If
    '
    '' Otherwise if there is all ChartNav data, remove it from the list
    'Else
    '    If g.bHasChartNavData And m.bAllChartNavData = True Then
    '        SetIniFileProperty "HardDriveData", Replace(strSystems, "," & m.lSystemNumber & ",", ""), "Data", g.strIniFile
    '    End If
    'End If
    
ErrExit:
    Set rs = Nothing
    Exit Sub

ErrSection:
    'If error, rollback entire transaction
    If bTransaction Then
        'g.WrkJet.Rollback
        bTransaction = False
    End If
    Set rs = Nothing
    RaiseError "cSystem.Save", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Validate
'' Description: Validate the system information
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Validate()
On Error GoTo ErrSection:
    
    Dim lIndex As Integer               ' Index into a for loop
    Dim rs As Recordset                 ' Recordset into the database
    Dim bReverify As Boolean            ' Do we need to reverify a rule?
    Dim strRules As String              ' List of rules to be reverified
    
    ' Make sure that the name of the system is between 1 and 50 characters
    If Len(m.strSystemName) = 0 Or Len(m.strSystemName) > 50 Then
        m.nErrNbr = 1
        Err.Raise vbObjectError + 1000, , "Strategy name must be 1 to 50 characters"
    End If
        
    ' Single quote's not allowed in system names
    If InStr(1, m.strSystemName, "'") > 0 Then
        Err.Raise vbObjectError + 1000, , "Strategy name cannot contain single quotes (')"
    End If

    ' Verify that the dates are valid
    If m.dFromDate > m.dToDate And m.bToEndOfData = False Then
        m.nErrNbr = 2
        Err.Raise vbObjectError + 1000, , "From date must be before the To date."
    End If

    If m.lBarsLoadedBeforeTrading <= 0 Then m.lBarsLoadedBeforeTrading = 350
    If m.lBarsTradedBeforeOrders <= 0 Then m.lBarsTradedBeforeOrders = 350
    If m.lTradeDepth > 1000 Then m.lTradeDepth = 1000
    If Not m.bPyramid Then m.lTradeDepth = 0
    
    ' If the Reverify flag is on then check the Rules of the system.  If
    ' any of the rules' Reverify flag is on then issue an error
    Set rs = g.dbNav.OpenRecordset("SELECT [Reverify] FROM [tblSystems] " & _
        "WHERE [SystemNumber]=" & m.lSystemNumber & ";", dbOpenSnapshot)
    If Not rs.EOF Then
        bReverify = rs!Reverify
    End If
    
    If bReverify Then
        For lIndex = 1 To m.Rules.Count
            If m.Rules.Item(lIndex).Reverify Then
                strRules = strRules & m.Rules.Item(lIndex).Name & "|"
            End If
        Next lIndex
        If Right(strRules, 1) = "|" Then strRules = Left(strRules, Len(strRules) - 1)
        If Len(strRules) > 0 Then Err.Raise vbObjectError + 1000, , "The following rules need to be resaved:||" & strRules
    End If
    
ErrExit:
    Set rs = Nothing
    Exit Sub

ErrSection:
    Set rs = Nothing
    RaiseError "cSystem.Validate", eGDRaiseError_Raise

End Sub

Private Sub ValidateRules()
On Error GoTo ErrSection:
    
    If m.Rules.Count < 1 Then
        Err.Raise vbObjectError + 1000, , _
            "At least one rule must exist before you can run a test."
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cSystem.ValidateRules", eGDRaiseError_Raise

End Sub

Public Function Optimized() As Boolean
On Error GoTo ErrSection:

    Dim X As Integer, r As Integer
    Dim lRule As Long
    Dim lInput As Long
    
    ' if any rules are alternating, then we're optimizing
    Optimized = False
    For lRule = 1 To m.Rules.Count
        If m.Rules.Item(lRule).Alternate Then
            Optimized = True
            Exit For
        Else
            For lInput = 1 To m.Rules.Item(lRule).Inputs.Count
                If m.Rules.Item(lRule).Selected And m.Rules.Item(lRule).Inputs.Item(lInput).IfOptimize Then
                    Optimized = True
                    Exit For
                End If
            Next lInput
        End If
        
        If Optimized Then Exit For
    Next lRule
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSystem.Optimized", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Delete
'' Description: Delete the system and any corresponding local rules
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Delete()
On Error GoTo ErrSection:
    
    Dim rs As Recordset                 ' Recordset from the database
    
    ' Delete any local rules first (both from the database and the global table)...
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblRules] WHERE [SystemNumber]=" & m.lSystemNumber & ";", dbOpenDynaset)
    Do While Not rs.EOF
        DeleteRule rs!RuleID
        rs.Delete
        rs.MoveNext
    Loop
    
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblSystems] " & _
                    "WHERE [SystemNumber]= " & m.lSystemNumber & ";", dbOpenDynaset)
    If Not rs.EOF Then
        KillFile AddSlash(App.Path) & "TRADES\S" & rs!SystemNumber & ".TXT"
        rs.Delete
    End If
    
ErrExit:
    Set rs = Nothing
    Exit Sub

ErrSection:
    Set rs = Nothing
    RaiseError "cSystem.Delete", eGDRaiseError_Raise

End Sub

Public Sub Test(Optional ByVal bUseHwnd As Boolean = False, _
                Optional ByVal lTestRuleID As Long = 0, _
                Optional astrTrades As cGdArray = Nothing, _
                Optional bOptimize As Boolean = True)
On Error GoTo ErrSection:
    
    Dim RunOptimizer    As Boolean
    Dim iMarketTraded   As Long
    Dim lIndex          As Long
    Dim strTradesFile   As String
    
    ' Validate market and rules data...
    ValidateRules
    
    ' If Trades folder not found then create one...
    If DirExist(App.Path & "\Trades") = False Then MakeDir App.Path & "\Trades", True
    
    ' Get the name of the Trades and Next Bar text files...
    strTradesFile = TradesFile(bUseHwnd)
    
    ' For now, always force a re-run (May be readressed later) 11/16/00 DAJ
    KillFile strTradesFile
    
    ' Check if at least one optimization check box is checked
    ' Or if at least one alternating rule is checked
    ''RunOptimizer = Optimized
    
    ' TLB 7/27/00: find the Market1 security
    'iMarketTraded = GetMarketTraded
    
    If bUseHwnd Then bOptimize = False
    
    ' Run the system...
    ''If RunOptimizer And bUseHwnd Then
    ''    ' Do not allow opimization from a chart...
    ''    InfBox "A strategy set to optimize cannot be run from a chart", "!"
    ''Else
        RunEngine False, groBoth, bOptimize, strTradesFile, , (Not bUseHwnd), lTestRuleID, astrTrades
    ''End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cSystem.Test", eGDRaiseError_Raise

End Sub

Public Sub TestFractZen(ByVal bShowReports As Boolean, Optional astrTrades As cGdArray = Nothing)
On Error GoTo ErrSection:
    
    Dim i&, nDate&, nEndDate&, dtTrade#, dPrice#, dCurEntry#, nSaveFromDate&, d#
    Dim strSymbol$, strTradesFile$, strPos$, strCurPos$, strTrade$
    Dim aTrades As cGdArray, aFzTrades As New cGdArray, aFlds As New cGdArray
    
'Specify a date range and symbol -- then for each date:
'- save current position
'- calc autobrkout for data through previous session
'- set bars to that brkout
'- run engine on bars for that day
'- for 1st bar of day, see if position changed (based on previous entry)
'- then process entries and exits for that day (unless entry into same position)
    
    ' Validate market and rules data...
    ValidateRules
    
    ' If Trades folder not found then create one...
    If DirExist(App.Path & "\Trades") = False Then MakeDir App.Path & "\Trades", True
    
    ' Get the name of the Trades and Next Bar text files...
    strTradesFile = TradesFile(False)
    KillFile strTradesFile
    
    'If bUseHwnd Then bOptimize = False
    
    strSymbol = m.Bars.Prop(eBARS_Symbol)
    
    If m.bToEndOfData Then
        nEndDate = LastDailyDownload
    Else
        nEndDate = m.dToDate
    End If
    
    strCurPos = "F"
    nSaveFromDate = m.dFromDate
    For nDate = m.dFromDate To nEndDate
        If IsWeekday(nDate) Then
            StatusMsg "FractZen backtest for: " & strSymbol & " " & DateFormat(nDate, MM_DD_YYYY)
        
            ' Load the data bars through today, but based on FractZen# through previous session
            i = g.FractZen.GetFractZenRange(strSymbol, nDate)
            m.Bars.Size = 0
            DM_GetBars m.Bars, strSymbol, Str(i) & "b", nDate - 14, nDate
            If m.Bars.SessionDate(m.Bars.Size - 1) < nDate Then
                i = i ' must have been a holiday?
            Else
                ' Run the system
                Set aTrades = New cGdArray
                m.dFromDate = m.Bars.SessionDate(0)
                RunEngine False, groSingleIteration, False, "", False, False, 0, aTrades
                
                If aFzTrades.Size = 0 Then
                    aFzTrades.Add aTrades(0)
                End If
                Set aTrades = SplitTrades(aTrades)
If IsIDE Then
    aTrades.ToFile App.Path & "\Trades\Temp\" & strSymbol & "-" & Str(JulToLong(nDate, True)) & ".txt"
End If
                'L   41670.664109    1132    FractZen Long Entry 41673.426088    1114.6  FractZen Long Fixed Stop    -1752.00    348.00  -1752.00    4   8517    8519    5   0   0   1   1   0
    
                ' get the initial position from the last trade prior to this trading session
                strPos = ""
                For i = aTrades.Size - 1 To 0 Step -1
                    strTrade = aTrades(i)
                    dtTrade = Val(Parse(strTrade, vbTab, 2))
                    If m.Bars.SessionDateForTime(dtTrade, False) < nDate Then
                        aTrades.Remove 0, i + 1 ' delete all the trades prior to this session
                        strPos = Parse(strTrade, vbTab, 1)
                        If Left(strPos, 1) = "X" Then
                            If Len(Parse(strTrade, vbTab, 4)) > 0 Then
                                strPos = "F"
                            Else
                                strPos = Mid(strPos, 2, 1)
                            End If
                        End If
                        Exit For
                    End If
                Next
                
                ' if initial position is same as current position
                If strPos = strCurPos Then
                    ' then just remove the last "open exit" in the cumulative trades
                    If aFzTrades.Size > 1 And strCurPos <> "F" Then
                        aFzTrades.Size = aFzTrades.Size - 1
                    End If
                ElseIf Len(strPos) > 0 Then
                    
                    If strPos <> "F" Then
                        strPos = "F"
                        ' rerun starting with today (so will pick up the next signal)
                        Set aTrades = New cGdArray
                        m.dFromDate = nDate
                        RunEngine False, groSingleIteration, False, "", False, False, 0, aTrades
                        Set aTrades = SplitTrades(aTrades)
                        If aTrades.Size > 0 Then
                            d = Val(Parse(aTrades(0), vbTab, 2))
                            If m.Bars.SessionDateForTime(d, False) < nDate Then
                                d = d
                            End If
                        End If
                    End If
                    
                    ' use open price, and time = 1 minute after start of session
                    dtTrade = nDate + (m.Bars.Prop(eBARS_StartTime) + 1) / 1440#
                    If m.Bars.Prop(eBARS_StartTime) > m.Bars.Prop(eBARS_EndTime) Then
                        dtTrade = dtTrade - 1 ' (is previous evening if an overnight session)
                    End If
                    For i = m.Bars.Size - 1 To 0 Step -1
                        If m.Bars.SessionDate(i - 1) < nDate Then
                            dPrice = m.Bars(eBARS_Open, i)
                            Exit For
                        End If
                    Next
                    
                    ' exit current position (reset info for last "open exit")
                    If strCurPos = "S" Or strCurPos = "L" Then
                        aFlds.SplitFields aFzTrades(aFzTrades.Size - 1), vbTab
                        aFlds(0) = "X" & strCurPos
                        aFlds(1) = Str(dtTrade)
                        aFlds(2) = Str(dPrice)
                        aFlds(3) = "New Position"
                        If Left(strTrade, 1) = "X" Then
                            aFlds(9) = Parse(strTrade, vbTab, 10) ' RuleID from an Exit
                        Else
                            aFlds(9) = Parse(strTrade, Chr(27), 2) ' RuleID from an Entry
                        End If
                        If aFlds(9) = "" Then
                            i = i
                        End If
                        aFzTrades(aFzTrades.Size - 1) = aFlds.JoinFields(vbTab)
                        If IsIDE And aFzTrades.Size > 2 Then
                            If Len(Parse(aFzTrades(aFzTrades.Size - 1), vbTab, 1)) + Len(Parse(aFzTrades(aFzTrades.Size - 2), vbTab, 1)) <> 3 Then
                                InfBox "Out of sequence for trades", "e", , "Error"
                            End If
                        End If
                    End If
                    ' and enter new position
                    If strPos = "S" Or strPos = "L" Then
                        aFlds.SplitFields strTrade, vbTab
                        aFlds(0) = strPos
                        aFlds(1) = Str(dtTrade)
                        aFlds(2) = Str(dPrice)
                        aFzTrades.Add aFlds.JoinFields(vbTab)
                        If IsIDE And aFzTrades.Size > 2 Then
                            If Len(Parse(aFzTrades(aFzTrades.Size - 1), vbTab, 1)) + Len(Parse(aFzTrades(aFzTrades.Size - 2), vbTab, 1)) <> 3 Then
                                InfBox "Out of sequence for trades", "e", , "Error"
                            End If
                        End If
                    End If
                    strCurPos = strPos
                End If
                
                ' then append all trades from today's session
                For i = 0 To aTrades.Size - 1
                    strTrade = aTrades(i)
                    If Left(strTrade, 1) = "X" And strCurPos = "F" And i = 0 Then
                        strPos = strPos
                    Else
                        aFzTrades.Add strTrade
                        If IsIDE And aFzTrades.Size > 2 Then
                            If Len(Parse(aFzTrades(aFzTrades.Size - 1), vbTab, 1)) + Len(Parse(aFzTrades(aFzTrades.Size - 2), vbTab, 1)) <> 3 Then
                                InfBox "Out of sequence for trades", "e", , "Error"
                            End If
                        End If
                        ' and track the current position
                        strPos = Left(strTrade, 1)
                        If strPos = "S" Or strPos = "L" Then
                            strCurPos = strPos
                        ElseIf Len(Parse(strTrade, vbTab, 4)) > 0 Then
                            strCurPos = "F" ' is flat only if a Rule is given for the exit
                        End If
                    End If
                Next
            End If
            DoEvents
        End If
    Next
    m.dFromDate = nSaveFromDate
    StatusMsg
    
    Set aFzTrades = MergeTrades(aFzTrades)
    If Not astrTrades Is Nothing Then
        Set astrTrades = aFzTrades
    End If
    KillFile strTradesFile
    If aFzTrades.Size > 0 Then
        aFzTrades.ToFile strTradesFile
        If FileExist(strTradesFile) And bShowReports Then
            ShowReports -1 ', strReportName
        End If
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cSystem.TestFractZen"
End Sub

' to split trades into separate entries and exits
Private Function SplitTrades(ByVal aTrades As cGdArray) As cGdArray
On Error GoTo ErrSection:

    Dim i&, iPos&, iTabs&, s$, strEntryRuleID$
    Dim aSplit As New cGdArray
    
    ' A Trade:
    'L   41670.664109    1132    FractZen Long Entry 41673.426088    1114.6  FractZen Long Fixed Stop    -1752.00    348.00  -1752.00    4   8517    8519    5   0   0   1   1   0
    ' Gets split into Entry and Exit:
    'L   41670.664109    1132    FractZen Long Entry    <Esc>EntryRuleID
    'XL  41673.426088    1114.6  FractZen Long Fixed Stop    -1752.00    348.00  -1752.00    4   8517    8519    5   0   0   1   1   0
    
    'aSplit.Add aTrades(0)
    For i = 1 To aTrades.Size - 1
        s = aTrades(i)
        iTabs = 0
        For iPos = 1 To Len(s)
            If Mid(s, iPos, 1) = vbTab Then
                iTabs = iTabs + 1
                If iTabs = 4 Then
                    strEntryRuleID = Parse(s, vbTab, 12)
                    aSplit.Add Left(s, iPos - 1) & vbTab & Chr(27) & strEntryRuleID
                    aSplit.Add "X" & Left(s, 1) & Mid(s, iPos)
                    Exit For
                End If
            End If
        Next
    Next
    
    Set SplitTrades = aSplit

ErrExit:
    Exit Function

ErrSection:
    RaiseError "cSystem.SplitTrades"
End Function

' to merge trades from separate entries and exits
Private Function MergeTrades(ByVal aTrades As cGdArray) As cGdArray
On Error GoTo ErrSection:

    Dim i&, iPos&, iTabs&, dProfit#, dDPP#, dExp#, dtLastTrade#
    Dim s$, strEntry$, strExit$, strEntryRuleID$
    Dim aUnsplit As New cGdArray, aFlds As New cGdArray
    
    ' a Trade:
    'L   41670.664109    1132    FractZen Long Entry 41673.426088    1114.6  FractZen Long Fixed Stop    -1752.00    348.00  -1752.00    4   8517    8519    5   0   0   1   1   0
    ' from Entry and Exit:
    'L   41670.664109    1132    FractZen Long Entry    <Esc>EntryRuleID
    'XL  41673.426088    1114.6  FractZen Long Fixed Stop    -1752.00    348.00  -1752.00    4   8517    8519    5   0   0   1   1   0
    
    If m.Bars.Prop(eBARS_TickMove) > 0 Then
        dDPP = Round(m.Bars.Prop(eBARS_TickValue) / m.Bars.Prop(eBARS_TickMove), 2)
    End If
    Select Case m.Bars.SecurityType
    Case "F"
        dExp = m.Expenses
    Case "S"
        dExp = m.dStockExpenses
    Case Else
        dExp = m.dForexExpenses
    End Select
    
    For i = 1 To aTrades.Size - 1 Step 2
        strEntry = aTrades(i)
        strExit = aTrades(i + 1)
        If Mid(strEntry, 2, 1) = vbTab And Mid(strExit, 3, 1) = vbTab Then
            strEntryRuleID = Parse(strEntry, Chr(27), 2)
            strEntry = Parse(strEntry, Chr(27), 1)
            If Right(strEntry, 1) = vbTab Then
                strEntry = Left(strEntry, Len(strEntry) - 1)
            End If
            s = strEntry & Mid(strExit, 3)
            aFlds.SplitFields s, vbTab
            dProfit = Val(aFlds(5)) - Val(aFlds(2))
            If Left(strEntry, 1) = "S" Then
                dProfit = -dProfit
            End If
            If dDPP > 0 Then
                dProfit = dProfit * dDPP - dExp
            End If
            aFlds(7) = Format(dProfit, "#0.00")
            If aFlds(11) <> strEntryRuleID Then
                aFlds(11) = strEntryRuleID
            End If
            dtLastTrade = Val(aFlds(4))
            s = aFlds.JoinFields(vbTab)
            aUnsplit.Add s
        Else
            ' error?
            If IsIDE Then
                InfBox "Out of sequence for trades", "e", , "Error"
            End If
            Exit For
        End If
    Next
    
    ' add/fix header
    aFlds.SplitFields aTrades(0), vbTab
    aFlds(4) = Str(dtLastTrade)
    s = aFlds.JoinFields(vbTab)
    aUnsplit.Add s, 0
    
    Set MergeTrades = aUnsplit

ErrExit:
    Exit Function

ErrSection:
    RaiseError "cSystem.MergeTrades"
End Function


' returns true if there are orders to place
Public Function NextBarReport(Optional ByVal NextBarMode As eGDNextBarMode = eGDNextBarMode_Normal, Optional ByVal dNextBarDate As Double = -99999#, Optional ByVal bAssumeNoPosition As Boolean = False, Optional ByVal bIgnoreNextBarData As Boolean = False) As Boolean
On Error GoTo ErrSection:
    
    Dim strNextBarFile As String
    Dim iIndex%, strSymbol$, strSecType$, rc&, strMarket$
    Dim bNeedToRun As Boolean, fh%, strLine$, dChkDate&
    Dim lIndex As Long, lCount As Long, strFile$
    Dim aDates As New cGdArray
    Dim aStrings As New cGdArray
    Dim lRule As Long
    Dim lParm As Long
    Dim lStart As Long
    
    Dim lLastDateOfData&
    Dim Bars As New cGdBars
    Dim strFormat As String
    Dim strBarPeriod As String
    
    Dim bVerifyFrom As Boolean
    'Dim bIgnoreNextBarData As Boolean
    'Dim bAssumeNoPosition As Boolean
    Dim dFromDate As Double
    Dim nMaxBars&
    
    Dim lScreenPointer As Long
    Dim dNewYorkTime As Double          ' Current time in New York
    
    lScreenPointer = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    
    ValidateRules
    
    'If TRADES folder not found then create one...
    If DirExist(App.Path & "\Trades") = False Then MakeDir App.Path & "\Trades", True

    strNextBarFile = NextBarFile(NextBarMode)
    
    'Get symbol for default Market1
    strSymbol = ""
    For lRule = 1 To m.Rules.Count
        With m.Rules.Item(lRule)
            For lParm = 1 To .Inputs.Count
                With .Inputs.Item(lParm)
                    If .ParmTypeID = 5 And UCase(.ParmName) = "MARKET1" Then
                        strMarket = .MarketSymbol
                        If .SymbolID = 0 Then
                            strSymbol = .Symbol
                        Else
                            strSymbol = GetSymbol(.SymbolID)
                        End If
                        strSecType = .SecurityType
                        strFormat = .Format
                        strBarPeriod = .Periodicity
                    End If
                End With
                If strSymbol <> "" Then Exit For
            Next lParm
        End With
        If strSymbol <> "" Then Exit For
    Next lRule
    
    If strSymbol = "" Then
        Err.Raise vbObjectError + 1000, , "You must supply information for Market1"
        Exit Function
    End If
   
    If dNextBarDate <> -99999# Then
        nMaxBars = m.lBarsLoadedBeforeTrading + m.lBarsTradedBeforeOrders
        If m.Bars.Size > nMaxBars Then m.Bars.DeleteFirstBars m.Bars.Size - nMaxBars
        aDates.Size = 1
        aDates(0) = dNextBarDate
                        
        If bAssumeNoPosition Then
            dFromDate = dNextBarDate
        Else
            dFromDate = 0
        End If
    Else
        Select Case NextBarMode
            Case eGDNextBarMode_UseHwnd
                ' cut down bars to just how much is needed
                nMaxBars = m.lBarsLoadedBeforeTrading + m.lBarsTradedBeforeOrders
                If m.Bars.Size > nMaxBars Then m.Bars.DeleteFirstBars m.Bars.Size - nMaxBars
                dNextBarDate = m.Bars(eBARS_DateTime, m.Bars.Size - 1)
                aDates.Size = 1
                aDates(0) = dNextBarDate
                dFromDate = 0
                bIgnoreNextBarData = False
                bAssumeNoPosition = False
                
            Case eGDNextBarMode_RunMult
                ' cut down bars to just how much is needed
                nMaxBars = m.lBarsLoadedBeforeTrading + m.lBarsTradedBeforeOrders
                If m.Bars.Size > nMaxBars Then m.Bars.DeleteFirstBars m.Bars.Size - nMaxBars
                dNewYorkTime = ConvertTimeZone(Now)
                
                If m.bToEndOfData = 0 Then
                    dNextBarDate = m.dToDate + 1
                ElseIf Hour(dNewYorkTime) < 14 Then
                    'if before 11am, assume still doing orders for today
                    '(use 11am since grains open at 10:30 ET and close at 11:10 PT)
                    dNextBarDate = Date
                Else 'otherwise, assume for tomorrow
                    dNextBarDate = Date + 1
                End If
                
                Do While Not IsWeekday(dNextBarDate)
                    dNextBarDate = dNextBarDate + 1
                Loop
                
                ' Verify with user
                Screen.MousePointer = vbDefault
                If frmNextBarOpt.ShowMe(dNextBarDate, dNextBarDate <> Int(dNextBarDate), bVerifyFrom, bAssumeNoPosition, bIgnoreNextBarData) Then
                    If Not bVerifyFrom Then
                        ' Just the one date
                        aDates.Size = 1
                        aDates(0) = dNextBarDate
                        If bAssumeNoPosition Then
                            dFromDate = dNextBarDate
                        Else
                            dFromDate = 0
                        End If
                    Else
                        ' VERIFY Process: run all dates from dump array file
                        fh = FreeFile
                        Open App.Path + "\VerifyNBR" For Input As #fh
                        Do While Not EOF(fh)
                            Line Input #fh, strLine
                            strLine = Parse(strLine, ",", 1)
                            If Len(strLine) > 0 Then
                                If InStr(strLine, "-") > 0 Then
                                    dChkDate = Val(StripStr(strLine, "-"))
                                    dChkDate = JulFromLong(dChkDate)
                                Else
                                    dChkDate = Val(strLine)
                                End If
                                If dChkDate >= dNextBarDate Then
                                    aDates.Add dChkDate
                                End If
                            End If
                        Loop
                        Close #fh
                        KillFile strNextBarFile
                    
                        MakeDir App.Path & "\Verify", True
                        KillFile App.Path & "\Verify\19*.*"
                        KillFile App.Path & "\Verify\20*.*"
                    End If
                Else
                    Exit Function
                End If
                Screen.MousePointer = vbHourglass
            
            Case eGDNextBarMode_Normal
                ' estimate date of next bar
                If Left(strFormat, 2) <> "CN" Then
                    If m.bToEndOfData = 0 Then
                        dNextBarDate = m.dToDate + 1
                    ElseIf Hour(Now) < 11 Then
                        'if before 11am, assume still doing orders for today
                        '(use 11am since grains open at 10:30 ET and close at 11:10 PT)
                        dNextBarDate = Date
                    Else 'otherwise, assume for tomorrow
                        dNextBarDate = Date + 1
                    End If
                    
                    Do While Not IsWeekday(dNextBarDate)
                        dNextBarDate = dNextBarDate + 1
                    Loop
                Else 'If g.bHasChartNavData Then
                    lLastDateOfData = LastDailyDownload + 1
                    Do While Not IsWeekday(lLastDateOfData)
                        lLastDateOfData = lLastDateOfData + 1
                    Loop
                    
                    If DM_GetBars(Bars, strSymbol, strBarPeriod, lLastDateOfData) Then
                        If Bars(eBARS_DateTime, Bars.Size - 1) > lLastDateOfData Then
                            dNextBarDate = Bars(eBARS_DateTime, Bars.Size - 1)
                        ElseIf Not IsIntraday(GetPeriodicity(strBarPeriod)) Then
                            dNextBarDate = lLastDateOfData
                        ElseIf Bars.Prop(eBARS_StartTime) <= Bars.Prop(eBARS_EndTime) Then
                            dNextBarDate = lLastDateOfData + Bars.Prop(eBARS_StartTime) / 1440#
                        Else
                            dNextBarDate = lLastDateOfData + Bars.Prop(eBARS_StartTime) / 1440# - 1
                        End If
                    Else
                        dNextBarDate = lLastDateOfData
                    End If
                End If
        
                ' Verify with user
                Screen.MousePointer = vbDefault
                If frmNextBarOpt.ShowMe(dNextBarDate, dNextBarDate <> Int(dNextBarDate), bVerifyFrom, bAssumeNoPosition, bIgnoreNextBarData) Then
                    If Not bVerifyFrom Then
                        ' Just the one date
                        aDates.Size = 1
                        aDates(0) = dNextBarDate
                        If bAssumeNoPosition Then
                            dFromDate = dNextBarDate
                        Else
                            dFromDate = 0
                        End If
                    Else
                        ' VERIFY Process: run all dates from dump array file
                        fh = FreeFile
                        Open App.Path + "\VerifyNBR" For Input As #fh
                        Do While Not EOF(fh)
                            Line Input #fh, strLine
                            strLine = Parse(strLine, ",", 1)
                            If Len(strLine) > 0 Then
                                If InStr(strLine, "-") > 0 Then
                                    dChkDate = Val(StripStr(strLine, "-"))
                                    dChkDate = JulFromLong(dChkDate)
                                Else
                                    dChkDate = Val(strLine)
                                End If
                                If dChkDate >= dNextBarDate Then
                                    aDates.Add dChkDate
                                End If
                            End If
                        Loop
                        Close #fh
                        KillFile strNextBarFile
                    
                        MakeDir App.Path & "\Verify", True
                        KillFile App.Path & "\Verify\19*.*"
                        KillFile App.Path & "\Verify\20*.*"
                    End If
                Else
                    Exit Function
                End If
                Screen.MousePointer = vbHourglass
        
        End Select
    End If
    
    DoEvents 'to let form hide
    
    ' For now, always force a re-run (May be readressed later) 11/16/00 DAJ
    KillFile strNextBarFile
    
    ' see if correct file already exists
    bNeedToRun = True
    If FileLength(strNextBarFile) > 0 Then
        fh = FreeFile
        Open strNextBarFile For Input As #fh
        Line Input #fh, strLine
        Close #fh
        strLine = Trim(Parse(strLine, vbTab, 1))
        dChkDate = Val(strLine)
        If dChkDate = dNextBarDate Then
            bNeedToRun = False
        End If
    End If
   
    'RUN THE ENGINE HERE...
    '======================
    If bNeedToRun Then
        KillFile strNextBarFile
        For lIndex = 0 To aDates.Size - 1
            RunEngine True, groNextBar, False, strNextBarFile, bIgnoreNextBarData, False, , , aDates(0), dFromDate
            If aDates.Size > 1 Then
                dChkDate = JulToLong(aDates(lIndex), True)
                If FileExist(strNextBarFile) Then
                    'move file to verify folder
                    If AskBox("w=NOWAIT ; b=Abort ; h=Next Bar Reports ; Date: " & Str(dChkDate)) = "A" Then
                        Exit For
                    End If
                    strFile = App.Path & "\Verify\" & CStr(dChkDate) & ".txt"
                    Name strNextBarFile As strFile
                    lCount = lCount + 1
                    DoEvents
                Else
                    InfBox "i=[] ; h=Verify NBR ; Error at date: " & CStr(dChkDate)
                    Exit For
                End If
            End If
        Next
        
        ' check if orders exist
        'aStrings.FromFile strNextBarFile, False, 4
        'If aStrings.Size > 3 Then NextBarReport = True
    End If
    
    If NextBarMode = eGDNextBarMode_Normal Then
        If aDates.Size > 1 Then
            InfBox CStr(lCount) & " files created in VERIFY folder."
        ElseIf Not FileExist(strNextBarFile) Then
            Err.Raise vbObjectError + 1000, , _
                "Orders report not found." '  To view the Orders report, make sure you check the 'Generate Orders for the next bar' option when testing the strategy."
        Else
            'Show next bar report
            Screen.MousePointer = vbDefault
            NextBarReport = frmNextBar.ShowMe(strNextBarFile, m.Rules)
        End If
    Else
        Select Case NextBarMode
            Case eGDNextBarMode_UseHwnd
                aStrings.FromFile AddSlash(App.Path) & "Trades\RB_" & m.hWnd & ".TXT"
            Case eGDNextBarMode_RunMult
                aStrings.FromFile AddSlash(App.Path) & "Trades\RB-" & Str(m.lSystemNumber) & "-" & Str(m.Bars.Prop(eBARS_SymbolID)) & ".TXT"
        End Select
        
        ''NextBarReport = (aStrings.Size > 3)
        ' Say that the system has orders only if there is a "2" line in the
        ' rule based report that is not false.  11/5/2003 DAJ
        ' Only check the "CP" block. 11/7/2003 DAJ
        NextBarReport = False
        For lIndex = 0 To aStrings.Size - 1
            If UCase(Trim(aStrings(lIndex))) = "0" & vbTab & "CP" Then
                lStart = lIndex + 1
                Exit For
            End If
        Next lIndex
        For lIndex = lStart To aStrings.Size - 1
            If Left(aStrings(lIndex), 3) = "  2" Then
                If UCase(Parse(aStrings(lIndex), vbTab, 2)) <> "FALSE" Then
                    NextBarReport = True
                    Exit For
                End If
            ElseIf Left(aStrings(lIndex), 1) = "0" Then
                Exit For
            End If
        Next lIndex
    End If
    
ErrExit:
    Screen.MousePointer = lScreenPointer
    Exit Function

ErrSection:
    Screen.MousePointer = lScreenPointer
    RaiseError "cSystem.NextBarReport", eGDRaiseError_Raise

End Function

Public Sub TestResults(Optional ByVal bUseHwnd As Boolean = False)
On Error GoTo ErrSection:
    
    Dim strTradeFile As String
    Dim iMarketTraded As Long
    
    'iMarketTraded = GetMarketTraded
    strTradeFile = TradesFile(bUseHwnd)
    If FileExist(strTradeFile) = False Then
        Err.Raise vbObjectError + 1000, , "No Results to Show for this Strategy"
    Else
        'Run the reports
        ''RunStudies strTradeFile, "Futures", kSN_Navigator, m.strDeveloper
        If bUseHwnd Then
            ShowReports m.hWnd
        Else
            ShowReports
        End If
            
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cSystem.TestResults", eGDRaiseError_Raise

End Sub

Public Function TradesFile(Optional ByVal bUseHwnd As Boolean = False) As String
On Error GoTo ErrSection:

    If bUseHwnd Then
        TradesFile = App.Path & "\Trades\S_" & Trim(Str(m.hWnd)) & ".TXT"
    Else
        TradesFile = App.Path & "\Trades\S" & Trim(Str(m.lSystemNumber)) & ".TXT"
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSystem.TradesFile", eGDRaiseError_Raise

End Function

Public Function NextBarFile(Optional ByVal NextBarMode As eGDNextBarMode = eGDNextBarMode_Normal) As String
On Error GoTo ErrSection:
    
    Dim lSymbolID As Long
    Dim strFile As String
    Dim mb As cMemBuffer
    
    lSymbolID = m.Markets("Market1").SymbolID
    
    Select Case NextBarMode
        Case eGDNextBarMode_Normal
            strFile = "NB" & Trim(Str(m.lSystemNumber))
        Case eGDNextBarMode_UseHwnd
            strFile = "NB_" & Trim(Str(m.hWnd))
        Case eGDNextBarMode_RunMult
            strFile = "NB-" & Str(m.lSystemNumber) & "-" & Str(lSymbolID)
            ' TLB 7/17/2012: for baskets, we need the filename be unique based also on custom inputs,
            ' since multiple basket items could be using same symbol and system but with different inputs
            If Len(m.strOverrides) > 0 Then
                ' append Hex of the CRC value for the custom inputs
                Set mb = New cMemBuffer
                mb.Buffer = m.strOverrides
                strFile = strFile & "-" & Hex(gdCalcMemCRC32(mb.MemPtr, mb.Length))
                Set mb = Nothing
            End If
    End Select
    
    If Len(strFile) > 0 Then
        NextBarFile = App.Path & "\Trades\" & strFile & ".TXT"
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSystem.NextBarFile", eGDRaiseError_Raise

End Function

Private Sub Class_Initialize()
On Error GoTo ErrSection:
    
    Set m.Rules = New cRules
    Set m.Markets = New cGdTree
    Set m.Bars = Nothing
    Set m.RptBridges = New cGdTree
    Set m.astrSingleIterationParms = New cGdArray
    
    Set m.tblSystem = New cGdTable
    Set m.tblSystemRules = New cGdTable
    Set m.tblSystemParms = New cGdTable
    Set m.tblRptMarkets = New cGdTable
    Set m.tblRptRules = New cGdTable
    
    With m.tblSystem
        .CreateField eGDARRAY_Longs, SystemField(etblSystem_ID)
        .CreateField eGDARRAY_Strings, SystemField(etblSystem_Name)
        .CreateField eGDARRAY_Strings, SystemField(etblSystem_Handle)
        .CreateField eGDARRAY_Doubles, SystemField(etblSystem_Expenses)
        .CreateField eGDARRAY_Doubles, SystemField(etblSystem_StockExpenses)
        .CreateField eGDARRAY_Longs, SystemField(etblSystem_TradeDepth)
        .CreateField eGDARRAY_TinyInts, SystemField(etblSystem_IfPyramid)
        .CreateField eGDARRAY_TinyInts, SystemField(etblSystem_CanPyramidSameEntry)
        .CreateField eGDARRAY_TinyInts, SystemField(etblSystem_SystemScope)
        .CreateField eGDARRAY_TinyInts, SystemField(etblSystem_DoShuffle)
        .CreateField eGDARRAY_TinyInts, SystemField(etblSystem_AllowReverse)
        .CreateField eGDARRAY_Doubles, SystemField(etblSystem_StockQty)
    End With
    
    With m.tblSystemRules
        .CreateField eGDARRAY_Longs, SystemRuleField(etblSystemRule_ID)
        .CreateField eGDARRAY_Longs, SystemRuleField(etblSystemRule_SysID)
        .CreateField eGDARRAY_Strings, SystemRuleField(etblSystemRule_Name)
        .CreateField eGDARRAY_Strings, SystemRuleField(etblSystemRule_CondCoded)
        .CreateField eGDARRAY_Strings, SystemRuleField(etblSystemRule_PriceCoded)
        .CreateField eGDARRAY_Strings, SystemRuleField(etblSystemRule_LimitCoded)
        .CreateField eGDARRAY_Strings, SystemRuleField(etblSystemRule_OrderPlacement)
        .CreateField eGDARRAY_TinyInts, SystemRuleField(etblSystemRule_BuySell)
        .CreateField eGDARRAY_TinyInts, SystemRuleField(etblSystemRule_Selected)
        .CreateField eGDARRAY_TinyInts, SystemRuleField(etblSystemRule_Alternate)
        .CreateField eGDARRAY_Longs, SystemRuleField(etblSystemRule_Use)
        .CreateField eGDARRAY_TinyInts, SystemRuleField(etblSystemRule_ExitOnEntryBar)
        .CreateField eGDARRAY_TinyInts, SystemRuleField(etblSystemRule_ExitOnEachTrade)
        .CreateField eGDARRAY_Longs, SystemRuleField(etblSystemRule_NumberContracts)
        .CreateField eGDARRAY_TinyInts, SystemRuleField(etblSystemRule_AsPercentPos)
        .CreateField eGDARRAY_Strings, SystemRuleField(etblSystemRule_LinkedRules)
        .CreateField eGDARRAY_Longs, SystemRuleField(etblSystemRule_Sequence)
    End With
    
    With m.tblSystemParms
        .CreateField eGDARRAY_Longs, SystemParmField(etblSystemParm_ID)
        .CreateField eGDARRAY_Strings, SystemParmField(etblSystemParm_Name)
        .CreateField eGDARRAY_Longs, SystemParmField(etblSystemParm_SysID)
        .CreateField eGDARRAY_Longs, SystemParmField(etblSystemParm_RuleID)
        .CreateField eGDARRAY_Longs, SystemParmField(etblSystemParm_TypeID)
        .CreateField eGDARRAY_Strings, SystemParmField(etblSystemParm_Value)
        .CreateField eGDARRAY_Doubles, SystemParmField(etblSystemParm_OptFromValue)
        .CreateField eGDARRAY_Doubles, SystemParmField(etblSystemParm_OptToValue)
        .CreateField eGDARRAY_Doubles, SystemParmField(etblSystemParm_OptStepValue)
        .CreateField eGDARRAY_TinyInts, SystemParmField(etblSystemParm_IfOptimize)
    End With
    
    With m.tblRptRules
        .CreateField eGDARRAY_Longs, RptRulesField(etblRptRules_RuleID)
        .CreateField eGDARRAY_Shorts, RptRulesField(etblRptRules_OrderTypeID)
        .CreateField eGDARRAY_Strings, RptRulesField(etblRptRules_RuleName)
    End With
    
    With m.tblRptMarkets
        .CreateField eGDARRAY_Strings, RptMarketsField(etblRptMarkets_Symbol)
        .CreateField eGDARRAY_Doubles, RptMarketsField(etblRptMarkets_LocalStart)
        .CreateField eGDARRAY_Doubles, RptMarketsField(etblRptMarkets_LocalEnd)
    End With
    
    m.bLinkInputs = True
    m.lLibraryID = kSN_UserLibrary
    m.Expenses = GetIniFileProperty("Expenses", 45#, "Systems", g.strIniFile)
    m.dStockExpenses = GetIniFileProperty("StockExpenses", 15#, "Systems", g.strIniFile)
    m.dForexExpenses = GetIniFileProperty("ForexExpenses", 5#, "Systems", g.strIniFile)
    m.bUseSharesPerTrade = GetIniFileProperty("UseSharesPerTrade", 0, "Systems", g.strIniFile)
    m.lNumShares = GetIniFileProperty("NumShares", 100&, "Systems", g.strIniFile)
    If m.lNumShares = 0 Then m.lNumShares = 100
    m.dDollarsPerTrade = GetIniFileProperty("DollarsPerTrade", 10000#, "Systems", g.strIniFile)
    m.lBarsLoadedBeforeTrading = GetIniFileProperty("BarsLoadedBeforeTrading", 260&, "Systems", g.strIniFile)
    m.lBarsTradedBeforeOrders = GetIniFileProperty("BarsTradedBeforeOrders", 260&, "Systems", g.strIniFile)
    m.lTradeDepth = GetIniFileProperty("TradeDepth", 0&, "Systems", g.strIniFile)
    m.strBarTimeFrame = GetIniFileProperty("BarTimeFrame", "Daily", "Systems", g.strIniFile)
    m.dFromDate = DateSerial(1950, 1, 1)
    m.dToDate = Date
    m.bToEndOfData = True
    m.bAllowReverse = True
    m.bForceLimitThrough = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSystem.Class.Initialize", eGDRaiseError_Raise

End Sub

Private Sub Class_Terminate()
On Error GoTo ErrSection:

    Dim lIndex As Long

    Set m.Rules = Nothing
    Set m.Markets = Nothing
    Set m.Bars = Nothing
    For lIndex = m.RptBridges.Count To 1 Step -1
        Set m.RptBridges(lIndex) = Nothing
    Next lIndex
    Set m.RptBridges = Nothing
    Set m.astrSingleIterationParms = Nothing
    
    Set m.tblSystem = Nothing
    Set m.tblSystemRules = Nothing
    Set m.tblSystemParms = Nothing
    Set m.tblRptRules = Nothing
    Set m.tblRptMarkets = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSystem.Class.Terminate", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadForChart
'' Description: Loads up a system from outside the System Manager
'' Inputs:      System Number, Symbol, Period, Date Range, hWnd of Caller
'' Returns:     Date of "next bar" to trade
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function LoadForChart(ByVal lSystemNumber As Long, Bars As cGdBars, ByVal hWnd As Long, _
            Optional ByVal bGameMode As Boolean = False) As Double
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lNextBar As Long
    Dim dCrossOver As Double
    
    ' Load the appropriate system
    Load lSystemNumber
    
    ' Set the system properties with the values passed in
    m.dFromDate = Bars(eBARS_DateTime, 0)
    m.bToEndOfData = True
    m.hWnd = hWnd
    
    ' Save copy of bars data
    Set m.Bars = Bars.MakeCopy
    For lIndex = m.Bars.Size - 1 To 0 Step -1
        If m.Bars(eBARS_Open, lIndex) <> kNullData Then
            ' see if last bar is "complete" (only if last bar
            ' is before the last daily download date)
            dCrossOver = m.Bars.Prop(eBARS_CrossoverTime)
            If dCrossOver = 0 Then dCrossOver = 1439
            If Bars.Prop(eBARS_PriceHasSettled) And Int(m.Bars(eBARS_DateTime, lIndex)) > LastDailyDownload Then
                dCrossOver = gdFixDateTime(Int(m.Bars(eBARS_DateTime, lIndex)) + dCrossOver / 1440#)
            Else
                dCrossOver = gdFixDateTime(LastDailyDownload + dCrossOver / 1440#)
            End If
            If m.Bars(eBARS_DateTime, lIndex) > dCrossOver Or bGameMode Then
                ' if incomplete, only use the open
                lNextBar = lIndex
                m.Bars(eBARS_Close, lNextBar) = kNullData
                m.Bars(eBARS_High, lNextBar) = kNullData
                m.Bars(eBARS_Low, lNextBar) = kNullData
            Else
                lNextBar = lIndex + 1
                ' if last bar is complete, need to add a forecast bar
                If lNextBar = m.Bars.Size Then m.Bars.AddForecastBars 1
            End If
            ' can chop data arrays after bar to trade
            m.Bars.Size = lNextBar + 1
            Exit For
        ElseIf lIndex = 0 Then
            m.Bars.Size = 0
        End If
    Next
    
    ChangeMarket1 Bars
    
ErrExit:
    ' return date of "next bar" (bar currently being traded)
    LoadForChart = m.Bars(eBARS_DateTime, m.Bars.Size - 1)
    Exit Function
    
ErrSection:
    RaiseError "cSystem.LoadForChart", eGDRaiseError_Raise
    
End Function

#If 0 Then
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadForMultRun
'' Description: Load a system for a multiple run situation
'' Inputs:      System Number, Bars, From Date, To End?, To Date, Overrides
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub LoadForMultRun(ByVal lSystemNumber As Long, Bars As cGdBars, ByVal dFromDate As Double, ByVal bToEnd As Boolean, Optional ByVal dToDate As Double = kNullData, Optional ByVal strOverrides As String = "")
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop

    m.strOverrides = strOverrides

    ' Load the appropriate system
    Load lSystemNumber
    
    m.dFromDate = dFromDate
    
    ' Set the date properties from the Bars
    If Bars.Size > 0 Then
        ''m.dFromDate = Bars(eBARS_DateTime, 0)
        If dToDate = kNullData Then
            m.dToDate = Bars(eBARS_DateTime, Bars.Size - 1)
        Else
            m.dToDate = dToDate
        End If
        m.bToEndOfData = bToEnd
    End If
    
    ' Change the information for all of the occurrences of Market1 with
    ' the Symbol information passed in
    Set m.Bars = Bars.MakeCopy
    m.Bars.AddForecastBars 1
    ChangeMarket1 Bars
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSystem.LoadForMultRun", eGDRaiseError_Raise
    
End Sub
#End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadBarsForMultRun
'' Description: Load a system for a multiple run situation
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub LoadBarsForMultRun(ByVal lSystemNumber As Long, ByVal strSymbol$, ByVal strPeriod$, _
        ByVal dFromDate As Double, ByVal bToEnd As Boolean, ByVal dToDate As Double, _
        Optional ByVal strOverrides As String = "", Optional ByVal bForNextBarReport As Boolean = False, Optional ByVal bAddToStream As Boolean = True)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop

    ' overrides must be set before .Load is called
    m.strOverrides = strOverrides
    
    ' Load the appropriate system (if 0, just keep the same system)
    If lSystemNumber <> 0 Then
        Load lSystemNumber
    End If
    
    m.dFromDate = dFromDate
    m.dToDate = dToDate
    m.bToEndOfData = bToEnd
    
    LoadMarket1Bars strSymbol, strPeriod, bForNextBarReport, bAddToStream
       
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSystem.LoadBarsForMultRun", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadMarket1Bars
'' Description: Load the bars for Market1
'' Inputs:      Symbol, Period, For Next Bar Report?, Add to Stream?
'' Returns:     Bars
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub LoadMarket1Bars(ByVal strSymbol As String, ByVal strPeriod As String, ByVal bForNextBarReport As Boolean, Optional ByVal bAddToStream As Boolean = True)
On Error GoTo ErrSection:

    Dim i&, nNumBars&, nToDate&, nFromDate&, nPeriodicity&
    Dim bExternal As Boolean
    Dim Bars As New cGdBars                 ' Bars to return from the function
    Static nPrevNumBars&, nPrevFromDate&, nPrevToDate&, nPrevPeriodicity&, strPrevSymbol$, nPrevLDD&
    
    If InStr(strSymbol, "|") > 0 Then
        bExternal = True
    Else
        SetBarProperties Bars, strSymbol ' must do this before setting FractZen
    End If
    
    If Not Me.ToEndOfData Then
        nToDate = Me.ToDate
    Else
        nToDate = 0
    End If
    nFromDate = Me.FromDate
    
    If UCase(strPeriod) = "FRACTZEN" And g.FractZen.Allowed And Not bExternal Then
        ' setup for FractZen bars
        g.FractZen.SetFractZen Bars, nToDate
        strPeriod = Bars.Prop(eBARS_PeriodicityStr)
    End If
    nPeriodicity = GetPeriodicity(strPeriod)
    
    If bExternal Then
        ' load data from external file (hard drive)
        Bars.FromFile Parse(strSymbol, "|", 5), Parse(strSymbol, "|", 2), Parse(strSymbol, "|", 1), strPeriod, , 0, nToDate
        strSymbol = Parse(strSymbol, "|", 1)
        GetMarketInfo strSymbol, Bars
    ElseIf Not IsIntraday(nPeriodicity) Then
        ' for non-intraday, just always load from the beginning (for both backtesting and next bar report)
        ' -- unless overridden (FromDate < 0), for special-case efficiency (e.g. SAI Elite report)
        If nFromDate < 0 And nToDate > Abs(nFromDate) Then
            nFromDate = nToDate + nFromDate
        Else
            nFromDate = 0
        End If
        DM_GetBars Bars, strSymbol, nPeriodicity, nFromDate, nToDate
    ElseIf bForNextBarReport Then
        ' for next bar report of intraday data, just load necessary # of bars prior to the report data
        nNumBars = -1.1 * (Me.BarsLoadedBeforeTrading + Me.BarsTradedBeforeOrders)
        DM_GetBars Bars, strSymbol, nPeriodicity, nNumBars, nToDate
    Else ' else for backtesting of intraday data:
        ' calc the necessary # of intraday bars prior to the start date
        nNumBars = -1.1 * Me.BarsLoadedBeforeTrading
        ' TLB 3/24/2015: if we already have this exact data loaded, no need to load it again
        ' (except for FractZen bars)
        If strSymbol <> strPrevSymbol Or nPeriodicity <> nPrevPeriodicity Or nNumBars <> nPrevNumBars _
                Or nFromDate <> nPrevFromDate Or nToDate <> nPrevToDate _
                Or LastDailyDownload <> nPrevLDD Or Bars.Prop(eBARS_FractZen) <> 0 Then
            Set m.PrevBars = Nothing
        End If
        If m.PrevBars Is Nothing Then
            ' first load the necessary # of intraday bars prior to the start date
            DM_GetBars Bars, strSymbol, nPeriodicity, nNumBars, nFromDate - 1
            ' then load the date range to be backtested
            DM_GetBars Bars, strSymbol, nPeriodicity, nFromDate, nToDate, True
            If Bars.Prop(eBARS_FractZen) = 0 Then
                ' save off info for the loaded data (to avoid reloading intraday data if nothing has changed)
                Set m.PrevBars = Bars.MakeCopy
                strPrevSymbol = strSymbol
                nPrevPeriodicity = nPeriodicity
                nPrevNumBars = nNumBars
                nPrevFromDate = nFromDate
                nPrevToDate = nToDate
                nPrevLDD = LastDailyDownload
            End If
        Else
            Set Bars = m.PrevBars.MakeCopy
        End If
    End If
    
    If (bAddToStream = True) And (ToEndOfData = True) And (bExternal = False) Then
        ' append streamed data
        g.RealTime.SpliceBars Bars
    End If
    
    If Bars.Size > 0 Then
        Bars.AddForecastBars 1 ' TLB: I think this should always be done ???
    End If
    Set m.Bars = Bars
    Set Bars = Nothing
    ChangeMarket1 m.Bars
    
    'Me.Bars = Bars

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSystem.LoadMarket1Bars"
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowNextBarReport
'' Description: Show the next bar report using the files that are already there
'' Inputs:      Whether to use hWnd for File Names
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowNextBarReport(Optional ByVal NextBarMode As eGDNextBarMode = eGDNextBarMode_Normal)
On Error GoTo ErrSection:

    Dim strNextBarFile As String        ' Filename of the Next Bar File
    
    ' Get the name of the Next Bar File
    strNextBarFile = NextBarFile(NextBarMode)
    
    ' Try to show the Next Bar Report
    If Not FileExist(strNextBarFile) Then
        Err.Raise vbObjectError + 1000, , "Orders report not found"
    Else
        frmNextBar.ShowMe strNextBarFile
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSystem.ShowNextBarReport", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RunEngine
'' Description: Run the system either in Back Testing or Next Bar mode
'' Inputs:
''     strFillLastBarAtLimitPrice = "" (none), "High", "Low", or "Both" (if limit fills exist at both high/low)
'' Returns:     True if successful, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function RunEngine(ByVal bOrdersNextBar As Boolean, ByVal RunOption As gdRunOption, _
                    ByVal bOptimize As Boolean, ByVal strTradesFile As String, _
                    Optional ByVal bIgnoreNextBarData As Boolean = False, _
                    Optional ByVal bShowReports As Boolean = True, _
                    Optional ByVal lTestRuleID As Long = 0, _
                    Optional aTrades As cGdArray = Nothing, _
                    Optional ByVal dTestToDate As Double = -1, _
                    Optional ByVal dTestFromDate As Double = -1, _
                    Optional ByVal astrSingleIterationParms As cGdArray = Nothing, _
                    Optional ByVal bAutomated As Boolean = False, _
                    Optional strError As String, _
                    Optional ByVal strFillLastBarAtLimitPrice As String = "") As Boolean
On Error GoTo ErrSection:
    
    Dim rc&, i&, iPos As Integer, fh&, iFrom&, d#, iBar&
    Dim ErrMsg$, strTemp$
    Dim dLoadFrom As Date
    Dim dLoadThru As Date
    Dim dStartTrading As Date
    Dim dStopTrading As Date
    Dim lNumRows As Long
    Dim astrParms As New cGdArray
    Dim astrParmValues As New cGdArray
    Dim fCallback As Long
    
    Dim aBars() As cGdBars
    Dim astrBarParms As New cGdArray
    
    Dim astrTrades As New cGdArray
    Dim astrConsolidated As New cGdArray
    Dim astrByRule As New cGdArray
    Dim astrHeader As New cGdArray
    
    Dim strSymbol As String
    Dim strFormat As String
    Dim strSecType As String
    Dim strFileName As String
    Dim strErrMsg As String
    Dim strPath As String
    Dim strPeriod As String
    Dim strSavePeriod As String
    Dim strDefault As String
    Dim iConvFactor As Integer
    Dim dDayMult As Double
    Dim bShowStudies As Boolean
    Dim bToTradeTable As Boolean
    Dim strKey As String
    Dim strReportName  As String
    
    Dim QryDef As QueryDef
    Dim strConsolidatedFile As String
    Dim strByRuleFile As String
    Dim bDMOpen As Boolean
    Dim lStart As Long
    
    Dim lSecurity As Long
    Dim dTime As Double
    
    Dim lLots As Long
    Dim dLots As Double
    Dim bDivide As Boolean
    Dim strForexSymbol As String
    Dim strBaseForex As String
    Dim BaseForex As cGdBars
    Dim MarketOneBars As cGdBars
    Dim bReturn As Boolean
    
    bReturn = True
    ReDim aBars(0) As cGdBars
    
    If bOptimize And FormIsLoaded("frmOptimizer") Then
        If frmOptimizer.InProgress Then
            Select Case frmOptimizer.OptMode
                Case eGDOptMode_Optimization
                    strError = "Cannot optimize this strategy while another optimization is running"
                Case eGDOptMode_MultipleRun, eGDOptMode_StrategyBasket
                    strError = "Cannot optimize this strategy while running a Strategy Basket"
                Case eGDOptMode_GameMode
                    strError = "Cannot optimize this strategy while running a Game Simulation"
                Case eGDOptMode_HighlightBarReport
                    strError = "Cannot optimize this strategy while running a Highlight Bar Report"
            End Select
            
            bReturn = False
            GoTo ErrExit
        End If
    End If
    
    ' If the Bars have been set, but they are empty, error out here...
    If Not m.Bars Is Nothing Then
        If m.Bars.Size = 0 Then
            strError = "No data loaded for " & m.Bars.Prop(eBARS_Symbol)
            bReturn = False
            GoTo ErrExit
        End If
    End If
    
    ' Set up some of the variables
    If dTestToDate = -1 Then
        If m.bToEndOfData Then
#If 1 Then
            ' TLB 9/25/2013: new method to get date of last good bar (ignoring any forecast bars)
            If Not m.Bars Is Nothing Then
                For i = m.Bars.Size - 1 To 0 Step -1
                    If m.Bars(eBARS_Close, i) <> kNullData Then
                        dTestToDate = m.Bars(eBARS_DateTime, i)
                        Exit For
                    End If
                Next
            End If
            If dTestToDate < 0 Then
                dTestToDate = Date + 1
            End If
#Else
            ' OLD CODE: just keep around for awhile for reference sake ...
            If Not m.Bars Is Nothing Then
                If m.Bars.Size = 1 Then
                    dTestToDate = m.Bars(eBARS_DateTime, m.Bars.Size - 1)
                Else
                    dTestToDate = m.Bars(eBARS_DateTime, m.Bars.Size - 2)
                End If
            Else
                dTestToDate = Date + 1
            End If
#End If
        Else
            dTestToDate = m.dToDate
            If Not m.Bars Is Nothing Then
                ' TLB 8/29/2014: for intraday data, add crossover time to the end date
                If m.Bars.IsIntraday Then
                    i = m.Bars.Prop(eBARS_CrossoverTime)
                    If i > 0 Then
                        dTestToDate = Int(dTestToDate) + i / 1440#
                    Else
                        dTestToDate = Int(dTestToDate) + 0.99
                    End If
                End If
            End If
        End If
    End If
    
    ' If we are pyramiding and doing a next bar orders report, we need to use the same
    ' starting date as the backtest because that is the first time we know for sure we
    ' are flat...
    If (dTestFromDate = -1) Or ((bOrdersNextBar = True) And (m.bPyramid = True) And (dTestFromDate <= 0)) Then
        dTestFromDate = m.dFromDate
    End If
    
    m.lTestRuleID = lTestRuleID
    
    ' Initialize the Engine...
    InitEngine True
    
    ' Create Arrays...
    astrBarParms.Create eGDARRAY_Strings
    astrConsolidated.Create eGDARRAY_Strings
    astrByRule.Create eGDARRAY_Strings
    astrTrades.Create eGDARRAY_Strings
    astrParmValues.Create eGDARRAY_Strings
    astrParms.Create eGDARRAY_Strings
    
    ' Get an array of unique market names with indexes into the Securities collection...
    If m.Markets.Count = 0 Then
        strError = "Market1 does not exist"
        bReturn = False
        GoTo ErrExit
    End If
    
    Screen.MousePointer = vbHourglass
    
    ' Walk through the unique markets...
    lNumRows = 0&
    For i = 1 To m.Markets.Count
        ' Set some default information if this is Market1...
        If m.Markets.Item(i).ParmName = "Market1" Then
            strDefault = m.Markets.Item(i).Periodicity
            
            ' TLB 6/11/2012: backtesting is NOT valid for Renko, Kagi, and P&F
            ' (since high/low of bars does not reflect actual trading, which makes backtested fills invalid)
            strTemp = UCase(strDefault)
            If InStr(strTemp, "RENKO") > 0 Or InStr(strTemp, "KAGI") > 0 Or InStr(strTemp, "P&F") > 0 Then
                strError = "Backtesting is not valid for Point & Figure,| Kagi or Renko bars (since these types| of bars do not reflect the actual trades)."
                bReturn = False
                GoTo ErrExit
            End If
            
            ' Calculate dates based on Market1's periodicity settings...
            Select Case UCase(Left(strDefault, 1))
                Case "Y"
                    dDayMult = 366
                Case "Q"
                    dDayMult = 92
                Case "M"
                    dDayMult = 31
                Case "W"
                    dDayMult = 7
                Case Else
                    dDayMult = 1.46
            End Select
            If dTestToDate = 0 Then dTestToDate = Date
            If dTestFromDate = 0 Then
                If bOrdersNextBar Then
                    dTestFromDate = Int(dTestToDate - (m.lBarsTradedBeforeOrders + 1) * dDayMult)
                Else
                    dTestFromDate = DateSerial(1950, 1, 1)
                End If
            End If
            
            ' Get "load from" and "load thru" dates...
            dLoadFrom = Int(dTestFromDate - (m.lBarsLoadedBeforeTrading + 1) * dDayMult)
            dLoadThru = dTestToDate
        End If
        
        lNumRows = lNumRows + 1
        ReDim Preserve aBars(lNumRows - 1) As cGdBars
        
        Set aBars(lNumRows - 1) = New cGdBars
        With m.Markets.Item(i)
            astrBarParms(lNumRows - 1) = .ParmName
            strFormat = .Format
            strSecType = .SecurityType
            If .SymbolID = 0 Then
                strSymbol = .Symbol
            Else
                strSymbol = GetSymbol(.SymbolID)
            End If
            strPath = .Path
            strPeriod = .Periodicity
        End With
        GetMarketInfo strSymbol, aBars(lNumRows - 1)
        'ValidateMarketInfo aBars(lNumRows - 1)

        If Left(strPeriod, 1) = "(" Then strPeriod = strDefault
        
        ' If the period of the system is more than daily, then load all
        ' of the data.  9/9/2002 DAJ
        If GetPeriodicity(strPeriod) > ePRD_Days + 1 Then
            dLoadFrom = 0
        End If
        
        ' Load bars
        If UCase(strFormat) = "CN" Or strFormat = "" Then
            If m.Markets.Item(i).ParmName = "Market1" Then bToTradeTable = True
            
            If m.Markets.Item(i).ParmName = "Market1" And Not m.Bars Is Nothing Then
                Set aBars(lNumRows - 1) = m.Bars
                
                If aBars(lNumRows - 1).Size = 0 Then
                    If Len(strSymbol) = 0 Then
                        strError = "No data available for " & m.Markets.Item(i).ParmName
                    Else
                        strError = "No data available for " & strSymbol
                    End If
                    
                    bReturn = False
                    GoTo ErrExit
                End If
            Else
                iFrom = Int(dLoadFrom)
                strSavePeriod = strPeriod
                If GetPeriodicity(strPeriod) >= ePRD_Days + 1 Then
                    strPeriod = "Daily"
                    ' TLB 12/18/2012: if secondary market is not intraday and primary is intraday,
                    ' then should probably be loading full history for the secondary market
                    If IsIntraday(GetPeriodicity(strDefault)) Then
                        iFrom = 2
                    End If
                End If
                If UCase(m.Markets.Item(i).ParmName) = "UNSPLIT" Then
                    rc = DM_GetBars(aBars(lNumRows - 1), strSymbol, strPeriod, iFrom, dLoadThru, , , True)
                Else
                    rc = DM_GetBars(aBars(lNumRows - 1), strSymbol, strPeriod, iFrom, dLoadThru)
                End If
                If rc <> 0 Then
                    ''g.RealTime.AddTickBuffer aBars(lNumRows - 1)
                    'If g.RealTime.Active And Int(dLoadThru) > LastDailyDownload Then
                    If g.RealTime.Active And (aBars(lNumRows - 1).SessionDateForTime(dLoadThru, False) > LastDailyDownload) Then
                        g.RealTime.SpliceBars aBars(lNumRows - 1)
                    End If
                    If GetPeriodicity(strSavePeriod) > ePRD_Days + 1 Then
                        strPeriod = strSavePeriod
                        aBars(lNumRows - 1).Prop(eBARS_Periodicity) = GetPeriodicity(strPeriod)
                        aBars(lNumRows - 1).BuildBars strPeriod
                    End If
                End If
                
                If rc = 0 Then
                    If Len(strSymbol) = 0 Then
                        strError = "There was an error loading data for " & m.Markets.Item(i).ParmName
                    Else
                        strError = "There was an error loading data for " & strSymbol
                    End If
                    
                    bReturn = False
                    GoTo ErrExit
                ElseIf aBars(lNumRows - 1).Size = 0 Then
                    ' TLB 2/2/2015: make an exception for $CRBX and CRBQ (used in the new CRBdelta function)
                    ' since we're really only using data from one or the other depending on what the date is.
                    If strSymbol = "$CRBX" Or strSymbol = "CRBQ" Then
                        aBars(lNumRows - 1).Size = 1 ' set size to 1 null bar just to avoid engine error
                    Else
                        If Len(strSymbol) = 0 Then
                            strError = "No data available for " & m.Markets.Item(i).ParmName
                        Else
                            strError = "No data available for " & strSymbol
                        End If
                        
                        bReturn = False
                        GoTo ErrExit
                    End If
                End If
            End If
        Else
            rc = aBars(lNumRows - 1).FromFile(strFormat, strPath, strSymbol, strPeriod, strErrMsg, dLoadFrom, dLoadThru)
            
            ' TLB 5/22/2013 #6836: if TickMove/TickValue not set from the file, then try to reload market info for external symbol
            If aBars(lNumRows - 1).TickMove = 0 Or aBars(lNumRows - 1).TickValue = 0 Then
                GetMarketInfo strSymbol, aBars(lNumRows - 1)
            End If
            
            If rc < 0 Then
                If Len(strSymbol) = 0 Then
                    strError = "There was an error loading data for " & m.Markets.Item(i).ParmName
                Else
                    strError = "There was an error loading data for " & strSymbol
                End If
                
                bReturn = False
                GoTo ErrExit
            ElseIf rc = 0 Then
                If Len(strSymbol) = 0 Then
                    strError = "No data available for " & m.Markets.Item(i).ParmName
                Else
                    strError = "No data available for " & strSymbol
                End If
                
                bReturn = False
                GoTo ErrExit
            End If
        End If
        
        If m.Markets.Item(i).ParmName = "Market1" Then
            ' 11/03/2011 DAJ: When the Bars aren't reloaded (e.g. HighlightBar Reporter), we
            ' are ending up double applying multipliers on the Tick Value in the cases below.  To
            ' fix that, load up the bar properties here and use them with the multiplier...
            Set MarketOneBars = New cGdBars
            SetBarProperties MarketOneBars, aBars(lNumRows - 1).Prop(eBARS_Symbol)
            
            ' For Stocks, change the TickValue to the TickValue times the number of shares...
            If aBars(lNumRows - 1).SecurityType = "S" Then
                aBars(lNumRows - 1).Prop(eBARS_TickValue) = MarketOneBars.Prop(eBARS_TickValue) * m.lNumShares
                ' TLB 5/17/2013 (#6835): if using "Dollars Per Trade" for Next Bar Report, then
                ' set the TickValue such that # of shares is based on the last bar's closing price
                If bOrdersNextBar And m.bUseSharesPerTrade = False And m.dDollarsPerTrade > 0 Then
                    ' look for last good price (but must be > 0)
                    For iBar = aBars(lNumRows - 1).Size - 1 To 0 Step -1
                        d = aBars(lNumRows - 1)(eBARS_Close, iBar)
                        If d > 0 Then
                            ' set TickValue = $perTrade / LastPrice * TickMove
                            d = Round(m.dDollarsPerTrade / d) * MarketOneBars.Prop(eBARS_TickMove)
                            aBars(lNumRows - 1).Prop(eBARS_TickValue) = d
                            Exit For
                        End If
                    Next
                End If
            ' 05/13/2010 DAJ: For a Forex symbol, need to multiply the tick value by the number
            ' of lots (usually 10000) for profit calculations to work out correctly.  Additionally,
            ' if the Forex symbol is non-US Dollar based, apply yesterday's close as a constant
            ' conversion rate...
            ElseIf IsForex(aBars(lNumRows - 1).Prop(eBARS_Symbol)) Then
                strForexSymbol = aBars(lNumRows - 1).Prop(eBARS_Symbol)
                strBaseForex = g.Profit.BaseForex(strForexSymbol, lLots, bDivide)
                dLots = CDbl(lLots) * g.Broker.MinimumLotSize(-1&, strForexSymbol)
                ' TLB 7/23/2014: let's make the default be the "normal" forex lot size (instead of the "mini")
                If dLots = 10000 Then
                    dLots = 100000
                End If
                If (Len(strBaseForex) = 0) And (dLots = 1) Then
                    dLots = 10000
                ElseIf (Len(strBaseForex) > 0) Then '(taken out per DaveJ): And (strBaseForex <> strForexSymbol) Then
                    Set BaseForex = New cGdBars
                    If DM_GetBars(BaseForex, strBaseForex, "Daily", LastDailyDownload - 5, , , , , False) Then
                        If BaseForex(eBARS_Close, BaseForex.Size - 1) <> kNullData Then
                            If bDivide Then
                                dLots = dLots / BaseForex(eBARS_Close, BaseForex.Size - 1)
                            Else
                                dLots = dLots * BaseForex(eBARS_Close, BaseForex.Size - 1)
                            End If
                        End If
                    End If
                End If
                aBars(lNumRows - 1).Prop(eBARS_TickValue) = MarketOneBars.Prop(eBARS_TickValue) * dLots
            End If
        End If
        
        ' Compress the bars if we need to
        If (Val(strPeriod) = 0) And (UCase(Left(strPeriod, 1)) <> "D") Then
            ''aBars(lNumRows - 1).BuildBars strPeriod
        End If
        
        ' Make sure got enough data
        If aBars(lNumRows - 1)(eBARS_DateTime, 0) > dLoadFrom Then
            If (lNumRows - 1) = 0 Then '(Market1)
                dLoadFrom = aBars(lNumRows - 1)(eBARS_DateTime, 0)
                If dTestFromDate < dLoadFrom Then
                    dTestFromDate = dLoadFrom
                End If
            ElseIf aBars(lNumRows - 1)(eBARS_DateTime, 0) > dTestFromDate Then
                ' We now want to allow this to flow through no matter if a secondary
                ' market does not have the same amount of data.  11/4/2003 DAJ
                'If m.hWnd = 0 Then
                '    Err.Raise vbObjectError + 1000, , "Not enough data for " & strSymbol
                'End If
            End If
        End If
    Next i
    
    If aBars(0).Prop(eBARS_TickMove) = 0 Or aBars(0).Prop(eBARS_TickValue) = 0 Then
        strError = "The strategy cannot be run because either the Tick Move or the Tick Value for Market1 is zero"
        bReturn = False
        GoTo ErrExit
    End If
        
    ' Fix the Market1 Crossover Time...
    If aBars(0).Prop(eBARS_CrossoverTime) = 0 Then aBars(0).Prop(eBARS_CrossoverTime) = 1439
    
    ' Trading can't go past last good close...
    For i = aBars(0).Size - 1 To 0 Step -1
        If aBars(0)(eBARS_Close, i) <> kNullData Then
            dStopTrading = aBars(0)(eBARS_DateTime, i)
            Exit For
        End If
    Next
    
    ' Fix the test to date to make sure it is not "between" bars...
    For i = aBars(0).Size - 1 To 0 Step -1
        If aBars(0)(eBARS_DateTime, i) = dTestToDate Then
            Exit For
        ElseIf aBars(0)(eBARS_DateTime, i) < dTestToDate Then
            If i < aBars(0).Size - 1 Then
                dTestToDate = aBars(0)(eBARS_DateTime, i + 1)
            End If
            Exit For
        End If
    Next i
    
    If bOrdersNextBar = False Then
        If dStopTrading < dTestToDate Then
            dTestToDate = dStopTrading
        End If
    Else
        Dim dNullValue#
        dNullValue = gdNullValue(aBars(0).ArrayHandle(eBARS_Close))
        If dStopTrading < dTestToDate Then
            ' see if next bar needs to be added
            If dTestToDate > aBars(0)(eBARS_DateTime, aBars(0).Size - 1) Then
                ' Add one bar and assign date
                If aBars(0)(eBARS_Open, aBars(0).Size - 1) = dNullValue And aBars(0)(eBARS_High, aBars(0).Size - 1) = dNullValue And aBars(0)(eBARS_Low, aBars(0).Size - 1) = dNullValue And aBars(0)(eBARS_Close, aBars(0).Size - 1) = dNullValue Then
                    aBars(0)(eBARS_DateTime, aBars(0).Size - 1) = dTestToDate
                Else
                    aBars(0).Size = aBars(0).Size + 1
                    aBars(0)(eBARS_DateTime, aBars(0).Size - 1) = dTestToDate
                End If
            End If
        Else
            Do While aBars(0)(eBARS_DateTime, aBars(0).Size - 1) > dTestToDate
                aBars(0).Size = aBars(0).Size - 1
            Loop
        
            ' Null out the prices on the last bar
            If bIgnoreNextBarData = True Then aBars(0)(eBARS_Open, aBars(0).Size - 1) = dNullValue
            If RunOption <> groCurrentBarOrders Then
                aBars(0)(eBARS_High, aBars(0).Size - 1) = dNullValue
                aBars(0)(eBARS_Low, aBars(0).Size - 1) = dNullValue
                aBars(0)(eBARS_Close, aBars(0).Size - 1) = dNullValue
            End If
        End If
    End If
    
    ' DAJ 08/10/2010: If the last two bars have the same date/time (e.g. for breakout bars), do
    ' an assertion here because that shouldn't happen...
    If IsIDE Then
        ' TLB: for now, exclude TF and things that end 1 minute prior to crossover
        If aBars(0).Prop(eBARS_CrossoverTime) <> aBars(0).Prop(eBARS_EndTime) + 1 Then
            If (aBars(0)(eBARS_DateTime, aBars(0).Size - 1) = aBars(0)(eBARS_DateTime, aBars(0).Size - 2)) Then
                Err.Raise vbObjectError + 1000, "Last two bars have the same date/time"
            End If
        End If
    End If
    
    ' Setup stuff...
    Screen.MousePointer = vbHourglass
    If Len(strTradesFile) > 0 Then
        KillFile strTradesFile
    End If
    astrParms.Clear
    
    ' Build parms in a gdArray (passed to engine)...
    '==============================================
    '0) Function Set Name
    astrParms(0) = ""
    
    strKey = "Software\Genesis Financial Data Services\Navigator Suite\General"
    If GetRegistryValue(rkLocalMachine, strKey, "LoadFunctions", vbUnchecked) = vbChecked Then
        g.bDirtyFunctionLibrary = True '(for now, till engine bug gets fixed)
    End If
    'If g.bDirtyFunctionLibrary Then
    '    rc = LoadEngineFunctions
    '    If rc <> 0 Then
    '        Err.Raise vbObjectError + 1000, , "Error loading engine functions"
    '    End If
    'End If
        
    '1) Up/down method (how to assume whether high or low of bar hits first)
    ' = "OMEGA"   if open > midbar, high hits before low
    ' = "GENESIS" if open > close, high hits before low
    ' = "TICKS"   to use ticks
    ''If GetRegistryValue(rkLocalMachine, strKey, "OptClose", True) Then
    ''    aStrParms(1) = "GENESIS"
    ''Else
    ''    aStrParms(1) = "OMEGA"
    ''End If
    ''astrParms(1) = GetIniFileProperty("TestingMethod", "GENESIS", _
            "Systems", g.strIniFile)
    astrParms(1) = "GENESIS"
    If GetRegistryValue(rkLocalMachine, strKey, "TickDataFill", vbChecked) = vbChecked Then
        If HasModule(aBars(0).SecurityType & "T") Then
            ' drop down to either FullTicks or MinutizedTicks depending on registry settings
            Select Case GetRegistryValue(rkLocalMachine, strKey, "TickFillMode", 0)
            Case 1
                astrParms(1) = "Ticks"
            Case 2
                astrParms(1) = "MinTicks"
            Case Else
                astrParms(1) = "Ticks"
                If aBars(0).Prop(eBARS_PeriodType) >= ePRD_Days Then
                    astrParms(1) = "MinTicks"
                ElseIf aBars(0).Prop(eBARS_PeriodType) = ePRD_Minutes Then
                    If aBars(0).Prop(eBARS_PeriodsPerBar) >= GetRegistryValue(rkLocalMachine, strKey, "TickFillModeMinutes", 30) Then
                        astrParms(1) = "MinTicks"
                    End If
                End If
            End Select
            
            ' TLB 10/19/2010: but just way too many ticks for ES to even bother doing full ticks
            strTemp = aBars(0).Prop(eBARS_Symbol)
            If Left(strTemp, 3) = "ES-" Or Left(strTemp, 4) = "ES1-" Then
                astrParms(1) = "MinTicks"
            End If
            
            ' TLB 10/19/2010: "Precision-tick" is bogging down charted strategies (esp. while
            ' streaming) -- daily charts only update once-a-day, but for intraday charts ...
            If m.hWnd <> 0 And aBars(0).IsIntraday Then
                ' minute-type charts only update once-a-minute at most, so we can probably
                ' just drop down to minutized ticks for those (no need for a 1-minute bar)
                If aBars(0).Prop(eBARS_PeriodType) = ePRD_Minutes And aBars(0).Prop(eBARS_PeriodsPerBar) > 1 Then
                    astrParms(1) = "MinTicks"
                Else ' else just ignore P.T.
                    astrParms(1) = "GENESIS"
                End If
            End If
        End If
    End If
    
    '2) Start Trading Date - YYYYMMDD
    If aBars(0).IsIntraday Then
        astrParms(2) = Format(dTestFromDate - 1, "yyyymmdd") & Format(aBars(0).Prop(eBARS_CrossoverTime) / 1440, "hhmmss")
    Else
        astrParms(2) = Format(dTestFromDate, "yyyymmdd") & "000000"
    End If
    
    '3) Stop Trading Date - YYYYMMDD
    If aBars(0).IsIntraday Then
        If (dTestToDate = Int(dTestToDate)) And (bOrdersNextBar = False) Then
            astrParms(3) = Format(dTestToDate, "yyyymmdd") & _
                    Format(aBars(0).Prop(eBARS_CrossoverTime) / 1440, "hhmmss")
        Else
            astrParms(3) = Format(dTestToDate, "yyyymmdd") & _
                    Format(dTestToDate, "hhmmss")
        End If
    Else
        astrParms(3) = Format(dTestToDate, "yyyymmdd") & "000000"
    End If
    
    '4) Last bar good data - OHLC, O means open is good, H means high is good, etc.
    astrParms(4) = "OHLC"
    'If bOrdersNextBar = True Then
    '    If bIgnoreNextBarData = True Then
    '        astrParms(4) = ""
    '    Else
    '        astrParms(4) = "O"
    '    End If
    'End If

    '5) Run Option - "OPTIMIZE", "ALTERNATE", "BOTH",
    'For single iteration: "PARMRUN"
    If Not bOptimize And Not bOrdersNextBar Then
        RunOption = groSingleIteration
    End If
    If (RunOption = groSingleIteration) Or (RunOption = groCurrentBarOrders) Then
        If Not astrSingleIterationParms Is Nothing Then
            Set m.astrSingleIterationParms = astrSingleIterationParms
        Else
            BuildParms
        End If
    End If
    
    Select Case RunOption
        Case groBoth
            astrParms(5) = "BOTH"
            astrParms(9) = ""
            fCallback = FunctionPtrToLong(AddressOf EngineCallback)
        Case groNextBar
            astrParms(5) = "NEXTBAR"
            astrParms(9) = ""
        Case groOptimize
            astrParms(5) = "OPTIMIZE"
            astrParms(9) = ""
            fCallback = FunctionPtrToLong(AddressOf EngineCallback)
        Case groAlternate
            astrParms(5) = "ALTERNATE"
            astrParms(9) = ""
            fCallback = FunctionPtrToLong(AddressOf EngineCallback)
        Case groSingleIteration
            astrParms(5) = "PARMRUN"
            astrParms(9) = ""
        Case groCurrentBarOrders
            astrParms(5) = "PARMRUN"
            astrParms(9) = "TRUE"
    End Select
    
    '6) Max Rules Off (if Alternating Rules)
    astrParms(6) = CStr(GetMaxAltRulesOff(bOptimize))
    
    '7) Rule ID of Rule to test alone if so desired
    astrParms(7) = "" 'strRuleID

    '10) Fill if Limit price is only touched (vs. must trade through it)
    If m.bForceLimitThrough Then
        astrParms(10) = strFillLastBarAtLimitPrice
    Else
        astrParms(10) = "TRUE"
    End If
    If Len(strFillLastBarAtLimitPrice) > 0 Then
        DebugLog aBars(0).Prop(eBARS_Symbol) & ": FillLastBarAtLimitPrice = " & strFillLastBarAtLimitPrice & ", " & astrParms(10)
    End If

    Dim alBarHandles As New cGdArray
    alBarHandles.Create eGDARRAY_Longs
    m.tblRptMarkets.NumRecords = 0
    
    For i = 0 To UBound(aBars)
        aBars(i).FreeExtra
        alBarHandles(i) = aBars(i).BarsHandle
        
        With m.tblRptMarkets
            .Item(RptMarketsField(etblRptMarkets_Symbol), i) = aBars(i).Prop(eBARS_MarketSymbol)
            .Item(RptMarketsField(etblRptMarkets_LocalStart), i) = aBars(i).Prop(eBARS_StartTime) / 1440
            .Item(RptMarketsField(etblRptMarkets_LocalEnd), i) = aBars(i).Prop(eBARS_EndTime) / 1440
        End With
    Next i
    
    ' Set the Current System to Me if optimizing...
    If bOptimize Then Set g.CurrentSystem = Me
    
    ' Run the system
    If m.astrSingleIterationParms Is Nothing Then
        Set m.astrSingleIterationParms = New cGdArray
    ''Else
        ''m.astrSingleIterationParms.ToFile App.Path & "\PARMS.CHK"
    End If
    
    ' Load the System tables...
    FillSystemTable "", (Len(strForexSymbol) > 0)
    strReportName = FillSystemRulesTable(lTestRuleID)
    FillSystemParmsTable
    If Len(strReportName) = 0 Then
        Select Case m.nRunMode
            Case eGDRunMode_All
                strReportName = m.strSystemName
            Case eGDRunMode_Long
                strReportName = m.strSystemName & " (Long Entries Only)"
            Case eGDRunMode_Short
                strReportName = m.strSystemName & " (Short Entries Only)"
        End Select
    End If
    
    If Not m.Bars Is Nothing Then
        strReportName = strReportName & " (Tested on " & m.Bars.Prop(eBARS_Symbol) & ")"
    End If
                
    ' Put the header information in the three arrays before passing off to the
    ' engine.  The engine will not touch the header line.
    strTemp = BuildHeader(aBars(0), dTestFromDate, dTestToDate, strReportName)
    astrTrades(0) = strTemp
    astrConsolidated(0) = strTemp
    astrByRule(0) = strTemp

'If RunOption = groNextBar Then
If 0 Then
    Dim aProps As New cGdArray
    astrParms.ToFile "StrParms.txt"
    astrBarParms.ToFile "StrBarParms.txt"
    If Not m.astrSingleIterationParms Is Nothing Then
        rc = m.astrSingleIterationParms.Size
    End If
    If fCallback <> 0 Then
        rc = rc
    End If
    For i = 0 To alBarHandles.Size - 1
        gdBarsToFile alBarHandles(i), "AT9", App.Path, "", "", -99, 0, 0
    Next
    For i = 1 To 24
        aProps.Add Str(aBars(0).Prop(i))
    Next
    For i = 101 To 109
        aProps.Add aBars(0).Prop(i)
    Next
    aProps.ToFile "StrProps.txt"
End If

    DebugLog "Start ExecuteSystem:  mode = " & astrParms(5)
    dTime = gdTickCount
    If FileLength(App.Path & "\Provided\SystemAdv.flg") >= 3 Then
If IsIDE And bOrdersNextBar Then
    StatusMsg "OrdersNextBar ... "
End If
        rc = ExecuteSystemAdv(astrParms.ArrayHandle, m.tblSystem.TableHandle, _
            m.tblSystemRules.TableHandle, m.tblSystemParms.TableHandle, astrBarParms.ArrayHandle, _
            alBarHandles.ArrayHandle, astrTrades.ArrayHandle, astrConsolidated.ArrayHandle, _
            astrByRule.ArrayHandle, m.astrSingleIterationParms.ArrayHandle, _
            fCallback, ByVal 0&)
If IsIDE And bOrdersNextBar Then
    StatusMsg "OrdersNextBar = " & Str(Round(gdTickCount - dTime)) & " ms"
End If
        DebugLog "Finish ExecuteSystemADV: " & Format((gdTickCount - dTime) / 1000#, "0.000") & " seconds"
    Else
        rc = ExecuteSystem(astrParms.ArrayHandle, m.tblSystem.TableHandle, _
            m.tblSystemRules.TableHandle, m.tblSystemParms.TableHandle, astrBarParms.ArrayHandle, _
            alBarHandles.ArrayHandle, astrTrades.ArrayHandle, astrConsolidated.ArrayHandle, _
            astrByRule.ArrayHandle, m.astrSingleIterationParms.ArrayHandle, _
            fCallback, ByVal 0&)
        DebugLog "Finish ExecuteSystem: " & Format((gdTickCount - dTime) / 1000#, "0.000") & " seconds"
    End If
    
    If Not bOrdersNextBar Then
        SaveDailyBars aBars(0).Prop(eBARS_Symbol)
    End If
    
    ' Pass back the trades array if an object was passed to hold it
    bShowStudies = True
    If Not aTrades Is Nothing Then
        Set aTrades = astrTrades
        bShowStudies = False
    End If
                
    If bOrdersNextBar = True Then
        bShowStudies = False
    ElseIf bOptimize Then
        If FormIsLoaded("frmOptimizer") Then
            frmOptimizer.TestRuleID = m.lTestRuleID
            frmOptimizer.TestRunMode = m.nRunMode
            
            If frmOptimizer.Visible Then
                bShowStudies = False
            End If
        End If
    End If

    'If next bar file, then do non-optimizing run
    If bOrdersNextBar = True Then
        
        strConsolidatedFile = strTradesFile
        ''strByRuleFile = AddSlash(FilePath(strTradesFile)) & "RB" & CStr(m.lSystemNumber) & ".TXT"
        strByRuleFile = Replace(strTradesFile, "\NB", "\RB")
        
        KillFile strConsolidatedFile
        KillFile strByRuleFile
        
        If rc = 0 Then
            ''astrConsolidated.Add BuildHeader(aBars(0), dTestFromDate, dTestToDate, strReportName), 0
            astrConsolidated.ToFile strConsolidatedFile
            ''astrByRule.Add BuildHeader(aBars(0), dTestFromDate, dTestToDate, strReportName), 0
            astrByRule.ToFile strByRuleFile
        End If
        
    ElseIf bShowStudies Then
        
        KillFile strTradesFile
        
        If rc = 0 Then
            If Parse(astrTrades(0), vbTab, 1) = "L" Or Parse(astrTrades(0), vbTab, 1) = "S" Then
                ''astrTrades.Add BuildHeader(aBars(0), dTestFromDate, dTestToDate, strReportName), 0
            Else
                ''astrTrades(0) = BuildHeader(aBars(0), dTestFromDate, dTestToDate, strReportName)
            End If
            
            astrTrades.ToFile strTradesFile
            'FileCopy "\trades.txt", App.Path & "\Trades\S22.Txt"
            
            ' store in table
            If bToTradeTable Then
                'StoreInTable aBars(0).Prop(eBARS_Symbol), m.lSystemNumber, m.strSystemName, strTradesFile, _
                    aBars(0).Prop(eBARS_PeriodicityStr)
            End If
    
            If FileExist(strTradesFile) And bShowReports Then
                ShowReports -1, strReportName
            End If
        End If
        
    End If
    
    If rc = 0 Then g.bDirtyFunctionLibrary = False

    'Check for an err msg (always appended to end of array)
    ErrMsg = astrParms(astrParms.Size - 1)
    
    '65560 is the message generated by engine when abort occurs
    If rc = kSN_OPTIMIZATION_RUN_ABORTED Then
    ElseIf rc = kSN_MAX_ITERATION_ERROR Then
        strError = "Error " & Str(rc) & ":|This sytem has exceeded the maximum of " & Format(kSN_MAX_ITERATIONS, "#,##0") & " optimization runs.  The strategy test has been aborted."
        bReturn = False
        GoTo ErrExit
    ElseIf rc <> 0 Then
        ' don't show error if problem is just due to not enough data
        If aBars(0).Size > 1 Then
            strError = "Error " & Str(rc) & ":|" & ErrMsg
            bReturn = False
            GoTo ErrExit
        End If
    End If
    
ErrExit:
    Screen.MousePointer = vbDefault
    If (bReturn = False) And (bAutomated = False) And (Len(strError) > 0) Then
        ' TLB 4/29/2016: if a Trades file was passed in, then must be running on a symbol group. In that
        ' case, just clear out the Trades file (except the first line) instead of showing an error message.
        If aTrades Is Nothing Then
            InfBox strError, "!", , "Strategy Test Error"
        Else
            aTrades.Size = 1
        End If
    End If
    RunEngine = bReturn
    
    Set astrParms = Nothing
    Set astrParmValues = Nothing
    For i = 0 To UBound(aBars)
        Set aBars(i) = Nothing
    Next i
    Set astrBarParms = Nothing
    Set astrTrades = Nothing
    Set astrConsolidated = Nothing
    Set astrByRule = Nothing
    Set astrHeader = Nothing
    Set m.astrSingleIterationParms = Nothing
    
    Exit Function
    
ErrSection:
    Screen.MousePointer = vbDefault
    RaiseError "cSystem.RunEngine", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetMaxAltRulesOff
'' Description: Get max # of alternating rules to allow "Off"
'' Inputs:      Whether system is Optimizing
'' Returns:     Max number of Alternating Rules
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetMaxAltRulesOff(ByVal bOptimize As Boolean) As Long
On Error GoTo ErrSection:
    
    Dim lMax As Long                    ' Maximum number of rules to return
    Dim strReturn As String             ' Return from the InfBox
    
    lMax = 12
    If bOptimize Then
        If FileExist(AddSlash(App.Path) & "MaxAlt.ASK") Then
            strReturn = AskBox("i=? ; g=num ; h=Optimizer ; d=12 ; Max number of alternating rules to allow OFF:")
            If strReturn <> "" Then lMax = Val(strReturn)
        End If
    End If
    GetMaxAltRulesOff = lMax
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSystem.GetMaxAltRulesOff", eGDRaiseError_Raise

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowReports
'' Description: Show the reports for the current system
'' Inputs:      Whether to use hWnd or System Number
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowReports(Optional ByVal hWnd As Long = -1, Optional ByVal strReportName As String = "", _
    Optional ByVal bGameMode As Boolean = False)
On Error GoTo ErrSection:
    
    Dim lIndex As Long
    Dim strTradesPath As String
    Dim RptBridge As New cRptBridge

    If bGameMode Then
        If Len(strReportName) > 0 Then
            strTradesPath = strReportName
        Else
            Exit Sub
        End If
    Else
        If hWnd = -1 Then
            strTradesPath = AddSlash(App.Path) & "Trades\S" & Trim(CStr(m.lSystemNumber)) & ".TXT"
        Else
            strTradesPath = AddSlash(App.Path) & "Trades\S_" & Trim(CStr(hWnd)) & ".TXT"
        End If
    End If

    For lIndex = m.RptBridges.Count To 1 Step -1
        If m.RptBridges(lIndex).IsLoaded = False Then
            m.RptBridges.Remove lIndex
        End If
    Next lIndex

    Set RptBridge = New cRptBridge
    With RptBridge
        .AppPath = App.Path
        .AppName = "System Navigator"
        .DB = g.dbNav
        .DefaultBeginBalance = 100000
        If Len(strReportName) > 0 Then
            .PortOrSystemName = strReportName
        ElseIf m.strSystemName <> "" Then
            .PortOrSystemName = m.strSystemName
        Else
            .PortOrSystemName = "New Strategy"
        End If
        .ShowInLocalTime = g.bShowInLocalTimeZone
        
        ' TLB 3/4/2011: more stuff to pass over
        .SetIrxBars GetIrxBarsHandle
        .SetAppBackColor GetAppBackColor
        .AltGridRowColor = ALT_GRID_ROW_COLOR
        
        .ChartHwnd = hWnd
        .Pyramiding = m.bPyramid
        If bGameMode Then
            .ImportTrades strTradesPath, m.lSystemNumber, m.strSystemName
        Else
            .ImportWithHandles strTradesPath, m.lSystemNumber, m.strSystemName, _
                        m.tblRptRules.TableHandle, m.tblRptMarkets.TableHandle
        End If
        .MainForm = frmMain
        .ImageList = frmMain.img16.ListImages
        
        .Help = g.Help
        .HideTdoReports = False
        m.RptBridges.Add RptBridge
        
        .Show True
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cSystem.ShowReports", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ChangeMarket1
'' Description: Change the symbol information for Market1
'' Inputs:      Bars structure with properties set
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ChangeMarket1(Bars As cGdBars)
On Error GoTo ErrSection:

    Dim lRule As Long                   ' Index into a for loop
    Dim lParm As Long                   ' Index into a for loop
    Dim lIndex As Long                  ' Index into a for loop
    Dim strParmName As String           ' Paramater name
    Dim strSymbol As String             ' Symbol for the parameter
    Dim strPeriod As String             ' Period for the parameter

    For lRule = 1 To m.Rules.Count
        With m.Rules.Item(lRule)
            For lParm = 1 To .Inputs.Count
                With .Inputs.Item(lParm)
                    If .ParmTypeID = 5 Then
                        strParmName = .ParmName
                        
                        Select Case UCase(strParmName)
                            Case "MARKET1", "UNSPLIT"
                                .Path = AddSlash(App.Path) & "Data"
                                .Symbol = Bars.Prop(eBARS_Symbol)
                                .MarketSymbol = Bars.Prop(eBARS_MarketSymbol)
                                .Periodicity = Bars.Prop(eBARS_PeriodicityStr)
                                .Format = "CN"
                                .SecurityType = Bars.SecurityType
                                .SecurityName = Bars.Prop(eBARS_Desc)
                                .SymbolID = Bars.Prop(eBARS_SymbolID)
                            
                            Case "DAILY", "WEEKLY", "MONTHLY", "QUARTERLY", "YEARLY"
                                .Path = AddSlash(App.Path) & "Data"
                                .Symbol = Bars.Prop(eBARS_Symbol)
                                .MarketSymbol = Bars.Prop(eBARS_MarketSymbol)
                                .Periodicity = .ParmName
                                .Format = "CN"
                                .SecurityType = Bars.SecurityType
                                .SecurityName = Bars.Prop(eBARS_Desc)
                                .SymbolID = Bars.Prop(eBARS_SymbolID)
                                
                            Case Else
                                If (Left(strParmName, 1) = Chr(34)) And (Right(strParmName, 1) = Chr(34)) Then
                                    strSymbol = Parse(Replace(strParmName, Chr(34), ""), ",", 1)
                                    strPeriod = Parse(Replace(strParmName, Chr(34), ""), ",", 2)
                                    
                                    If (Len(strSymbol) = 0) And (Len(strPeriod) > 0) Then
                                        .Path = AddSlash(App.Path) & "Data"
                                        .Symbol = Bars.Prop(eBARS_Symbol)
                                        .MarketSymbol = Bars.Prop(eBARS_MarketSymbol)
                                        .Periodicity = strPeriod
                                        .Format = "CN"
                                        .SecurityType = Bars.SecurityType
                                        .SecurityName = Bars.Prop(eBARS_Desc)
                                        .SymbolID = Bars.Prop(eBARS_SymbolID)
                                    End If
                                End If
                        
                        End Select
                    End If
                End With
            Next lParm
        End With
    Next lRule
    
    For lIndex = 1 To m.Markets.Count
        With m.Markets.Item(lIndex)
            strParmName = .ParmName
            
            Select Case UCase(strParmName)
                Case "MARKET1", "UNSPLIT"
                    .Path = AddSlash(App.Path) & "Data"
                    .Symbol = Bars.Prop(eBARS_Symbol)
                    .SymbolID = Bars.Prop(eBARS_SymbolID)
                    .MarketSymbol = Bars.Prop(eBARS_MarketSymbol)
                    .Periodicity = Bars.Prop(eBARS_PeriodicityStr)
                    .Format = "CN"
                    .SecurityType = Bars.SecurityType
                    .SecurityName = Bars.Prop(eBARS_Desc)
                
                Case "DAILY", "WEEKLY", "MONTHLY", "QUARTERLY", "YEARLY"
                    .Path = AddSlash(App.Path) & "Data"
                    .Symbol = Bars.Prop(eBARS_Symbol)
                    .SymbolID = Bars.Prop(eBARS_SymbolID)
                    .MarketSymbol = Bars.Prop(eBARS_MarketSymbol)
                    .Periodicity = .ParmName
                    .Format = "CN"
                    .SecurityType = Bars.SecurityType
                    .SecurityName = Bars.Prop(eBARS_Desc)
                    
                Case Else
                    If (Left(strParmName, 1) = Chr(34)) And (Right(strParmName, 1) = Chr(34)) Then
                        strSymbol = Parse(Replace(strParmName, Chr(34), ""), ",", 1)
                        strPeriod = Parse(Replace(strParmName, Chr(34), ""), ",", 2)
                        
                        If (Len(strSymbol) = 0) And (Len(strPeriod) > 0) Then
                            .Path = AddSlash(App.Path) & "Data"
                            .Symbol = Bars.Prop(eBARS_Symbol)
                            .MarketSymbol = Bars.Prop(eBARS_MarketSymbol)
                            .Periodicity = strPeriod
                            .Format = "CN"
                            .SecurityType = Bars.SecurityType
                            .SecurityName = Bars.Prop(eBARS_Desc)
                            .SymbolID = Bars.Prop(eBARS_SymbolID)
                        End If
                    End If
            End Select
        End With
    Next lIndex
    
    ' Also set the BarTimeFrame for the system to the periodicity of Market1...
    m.strBarTimeFrame = Bars.Prop(eBARS_PeriodicityStr)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSystem.ChangeMarket1", eGDRaiseError_Raise
    
End Sub

Private Sub BuildParms()
On Error GoTo ErrSection:
        
    Dim astrTemp As New cGdArray
    Dim lPos As Long
    Dim lIndex As Long
    Dim lRule As Long
    Dim lInput As Long
    Dim bContinue As Boolean
    
    astrTemp.Create eGDARRAY_Strings
    
    Set m.astrSingleIterationParms = New cGdArray
    m.astrSingleIterationParms.Create eGDARRAY_Strings
    
    ' Only include rules that are marked for alternating...
    For lRule = 1 To m.Rules.Count
        With m.Rules.Item(lRule)
            bContinue = False
            Select Case m.nRunMode
                Case eGDRunMode_All
                    If .RuleUse = 1 Or m.lTestRuleID = 0 Or .RuleID = m.lTestRuleID Then
                        bContinue = True
                    End If
                Case eGDRunMode_Long
                    If .RuleUse = 1 Or .BuySell = True Then bContinue = True
                Case eGDRunMode_Short
                    If .RuleUse = 1 Or .BuySell = False Then bContinue = True
            End Select
            
            If bContinue Then
                If .Alternate Then
                    m.astrSingleIterationParms.Add .Name & vbTab & "<alternate>" & vbTab & "false" & vbTab & "false"
                ElseIf .Selected Then
                    For lInput = 1 To .Inputs.Count
                        With .Inputs.Item(lInput)
                            If .ParmTypeID <> kSN_RetBars Then
                                If m.bLinkInputs Then
                                    If Not astrTemp.BinarySearch(.ParmName, lPos, eGdSort_IgnoreCase) Then
                                        m.astrSingleIterationParms.Add "" & vbTab & .ParmName & vbTab & Str(.Value) & vbTab & "true"
                                        astrTemp.Add .ParmName, lPos
                                    End If
                                Else
                                    m.astrSingleIterationParms.Add m.Rules.Item(lRule).Name & vbTab & .ParmName & vbTab & Str(.Value) & vbTab & "true"
                                End If
                            End If
                        End With
                    Next lInput
                End If
            End If
        End With
    Next lRule
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSystem.BuildParms", eGDRaiseError_Raise

End Sub

Private Function BuildHeader(ByVal Bars As cGdBars, ByVal dFromDate As Double, ByVal dToDate As Double, _
                            Optional ByVal strReportName As String = "") As String
On Error GoTo ErrSection:

    Dim dRisk#
    Dim astrHeader As New cGdArray
    
    Set astrHeader = New cGdArray
    astrHeader.Create eGDARRAY_Strings

    astrHeader(TradesHdrField(eTradesHeader_SystemNumber)) = CStr(m.lSystemNumber)
    If m.bHighlightBarReport Then
        astrHeader(TradesHdrField(eTradesHeader_SystemName)) = "Highlight Bar Report"
    ElseIf Len(strReportName) > 0 Then
        astrHeader(TradesHdrField(eTradesHeader_SystemName)) = strReportName
    ElseIf m.lSystemNumber <> 0 Then
        astrHeader(TradesHdrField(eTradesHeader_SystemName)) = m.strSystemName
    Else
        astrHeader(TradesHdrField(eTradesHeader_SystemName)) = "New Strategy"
    End If
    astrHeader(TradesHdrField(eTradesHeader_BarTimeFrame)) = m.strBarTimeFrame
    astrHeader(TradesHdrField(eTradesHeader_StartDate)) = dFromDate
    astrHeader(TradesHdrField(eTradesHeader_EndDate)) = dToDate
    astrHeader(TradesHdrField(eTradesHeader_TotalBars)) = Str(Bars.Size)
    astrHeader(TradesHdrField(eTradesHeader_Expenses)) = Str(m.Expenses)
    astrHeader(TradesHdrField(eTradesHeader_Symbol)) = Bars.Prop(eBARS_Symbol)
    astrHeader(TradesHdrField(eTradesHeader_TickMove)) = Str(Bars.Prop(eBARS_TickMove))
    astrHeader(TradesHdrField(eTradesHeader_TickValue)) = Str(Bars.Prop(eBARS_TickValue))
    astrHeader(TradesHdrField(eTradesHeader_MinMoveInTicks)) = Str(Bars.Prop(eBARS_MinMoveInTicks))
    astrHeader(TradesHdrField(eTradesHeader_Margin)) = Str(Bars.Prop(eBARS_Margin))
    astrHeader(TradesHdrField(eTradesHeader_SecurityType)) = Bars.SecurityType
    astrHeader(TradesHdrField(eTradesHeader_SessionStart)) = Str(Bars.Prop(eBARS_StartTime))
    astrHeader(TradesHdrField(eTradesHeader_SessionEnd)) = Str(Bars.Prop(eBARS_EndTime))
    astrHeader(TradesHdrField(eTradesHeader_TimeZoneInfo)) = Bars.Prop(eBARS_ExchangeTimeZoneInf)
    
    dRisk = GetRiskAmount(True)
    astrHeader(TradesHdrField(eTradesHeader_LongStopLoss)) = Str(Round(dRisk))
    dRisk = GetRiskAmount(False)
    astrHeader(TradesHdrField(eTradesHeader_ShortStopLoss)) = Str(Round(dRisk))
    
    BuildHeader = astrHeader.JoinFields(vbTab)

ErrExit:
    Set astrHeader = Nothing
    Exit Function
    
ErrSection:
    Set astrHeader = Nothing
    RaiseError "cSystem.BuildHeader", eGDRaiseError_Raise
    
End Function

Private Sub FillSystemTable(Optional ByVal strHandle As String = "", Optional ByVal bIsForex As Boolean = False)
On Error GoTo ErrSection:

    With m.tblSystem
        .NumRecords = 0
        .NumRecords = 1
        
        .Num(SystemField(etblSystem_ID), 0) = m.lSystemNumber
        If Len(Trim(m.strSystemName)) > 0 Then
            .Item(SystemField(etblSystem_Name), 0) = m.strSystemName
        Else
            .Item(SystemField(etblSystem_Name), 0) = "New Strategy"
        End If
        .Item(SystemField(etblSystem_Handle), 0) = strHandle
        .Num(SystemField(etblSystem_Expenses), 0) = CDbl(m.Expenses) * 100#
        ' TLB 4/2/2011: the engine is already using "StockExpenses" for both stocks
        ' and indices (including Forex), so for Forex symbols let's just use that field
        ' (but need to divide in half since stocks are used "per transaction").
        If bIsForex Then
            .Num(SystemField(etblSystem_StockExpenses), 0) = CDbl(m.dForexExpenses) * 100# / 2
        Else
            .Num(SystemField(etblSystem_StockExpenses), 0) = CDbl(m.dStockExpenses) * 100#
        End If
        If m.bPyramid Then
            .Num(SystemField(etblSystem_TradeDepth), 0) = m.lTradeDepth
        Else
            .Num(SystemField(etblSystem_TradeDepth), 0) = 0&
        End If
        .Num(SystemField(etblSystem_IfPyramid), 0) = m.bPyramid
        .Num(SystemField(etblSystem_CanPyramidSameEntry), 0) = m.bCanPyramidSameEntry
        .Num(SystemField(etblSystem_SystemScope), 0) = m.bLinkInputs
        If HasLinkedRules Then
            .Num(SystemField(etblSystem_DoShuffle), 0) = 0
        Else
            .Num(SystemField(etblSystem_DoShuffle), 0) = 1
        End If
        .Num(SystemField(etblSystem_AllowReverse), 0) = m.bAllowReverse
        ' TLB 3/14/2011: pass StockQty as either SharesPerTrade or DollarsPerTrade
        If m.bUseSharesPerTrade Then
            .Num(SystemField(etblSystem_StockQty), 0) = -m.lNumShares ' a negative #
        Else
            .Num(SystemField(etblSystem_StockQty), 0) = m.dDollarsPerTrade ' a positive #
        End If
    End With

'FileFromString "Sys.TXT", m.tblSystem.ToString(vbCrLf, vbTab)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSystem.FillSystemTable", eGDRaiseError_Raise
    
End Sub

Private Function FillSystemRulesTable(Optional ByVal lTestRuleEntryID As Long = 0) As String
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    With m.tblSystemRules
        .NumRecords = 0
        .NumRecords = m.Rules.Count
        
        m.tblRptRules.NumRecords = 0
        m.tblRptRules.NumRecords = m.Rules.Count
        
        For lIndex = 1 To m.Rules.Count
            .Num(SystemRuleField(etblSystemRule_ID), lIndex - 1) = m.Rules.Item(lIndex).RuleID
            .Num(SystemRuleField(etblSystemRule_SysID), lIndex - 1) = m.lSystemNumber
            .Item(SystemRuleField(etblSystemRule_Name), lIndex - 1) = m.Rules.Item(lIndex).Name
            .Item(SystemRuleField(etblSystemRule_CondCoded), lIndex - 1) = m.Rules.Item(lIndex).CondCoded
            .Item(SystemRuleField(etblSystemRule_PriceCoded), lIndex - 1) = m.Rules.Item(lIndex).PriceCoded
            .Item(SystemRuleField(etblSystemRule_LimitCoded), lIndex - 1) = m.Rules.Item(lIndex).Price2Coded
            .Item(SystemRuleField(etblSystemRule_OrderPlacement), lIndex - 1) = m.Rules.Item(lIndex).OrderPlacement
            .Num(SystemRuleField(etblSystemRule_BuySell), lIndex - 1) = m.Rules.Item(lIndex).BuySell
            ' if testing an entry rule, ignore all other entries
            If m.Rules.Item(lIndex).RuleUse <> 0 Then
                .Num(SystemRuleField(etblSystemRule_Selected), lIndex - 1) = m.Rules.Item(lIndex).Selected
                .Num(SystemRuleField(etblSystemRule_Alternate), lIndex - 1) = m.Rules.Item(lIndex).Alternate
            ElseIf lTestRuleEntryID = 0 Then
                Select Case m.nRunMode
                    Case eGDRunMode_All
                        .Num(SystemRuleField(etblSystemRule_Selected), lIndex - 1) = m.Rules.Item(lIndex).Selected
                        .Num(SystemRuleField(etblSystemRule_Alternate), lIndex - 1) = m.Rules.Item(lIndex).Alternate
                    Case eGDRunMode_Long
                        .Num(SystemRuleField(etblSystemRule_Selected), lIndex - 1) = m.Rules.Item(lIndex).BuySell And m.Rules.Item(lIndex).Selected
                        .Num(SystemRuleField(etblSystemRule_Alternate), lIndex - 1) = m.Rules.Item(lIndex).BuySell And m.Rules.Item(lIndex).Alternate
                        '.Num(SystemRuleField(etblSystemRule_Alternate), lIndex - 1) = False
                    Case eGDRunMode_Short
                        .Num(SystemRuleField(etblSystemRule_Selected), lIndex - 1) = Not m.Rules.Item(lIndex).BuySell And m.Rules.Item(lIndex).Selected
                        .Num(SystemRuleField(etblSystemRule_Alternate), lIndex - 1) = Not m.Rules.Item(lIndex).BuySell And m.Rules.Item(lIndex).Alternate
                        '.Num(SystemRuleField(etblSystemRule_Alternate), lIndex - 1) = False
                End Select
            ElseIf lTestRuleEntryID = m.Rules.Item(lIndex).RuleID Then
                .Num(SystemRuleField(etblSystemRule_Selected), lIndex - 1) = True
                .Num(SystemRuleField(etblSystemRule_Alternate), lIndex - 1) = False
                ' return name of rule (to use as name of report)
                FillSystemRulesTable = m.Rules.Item(lIndex).Name
            Else
                .Num(SystemRuleField(etblSystemRule_Selected), lIndex - 1) = False
                .Num(SystemRuleField(etblSystemRule_Alternate), lIndex - 1) = False
            End If
            .Num(SystemRuleField(etblSystemRule_Use), lIndex - 1) = m.Rules.Item(lIndex).RuleUse
            .Num(SystemRuleField(etblSystemRule_ExitOnEntryBar), lIndex - 1) = m.Rules.Item(lIndex).ExitOnEntryBar
            .Num(SystemRuleField(etblSystemRule_ExitOnEachTrade), lIndex - 1) = m.Rules.Item(lIndex).SysExitBasedOnEachTrade
            If m.bPyramid Then
                .Num(SystemRuleField(etblSystemRule_NumberContracts), lIndex - 1) = m.Rules.Item(lIndex).SysNumContracts
                .Num(SystemRuleField(etblSystemRule_AsPercentPos), lIndex - 1) = m.Rules.Item(lIndex).SysAsPercentOfPosition
            Else
                .Num(SystemRuleField(etblSystemRule_NumberContracts), lIndex - 1) = 1
                .Num(SystemRuleField(etblSystemRule_AsPercentPos), lIndex - 1) = False
            End If
            .Item(SystemRuleField(etblSystemRule_LinkedRules), lIndex - 1) = m.Rules.Item(lIndex).LinkedRules
            .Num(SystemRuleField(etblSystemRule_Sequence), lIndex - 1) = m.Rules.Item(lIndex).Seq
                        
            m.tblRptRules(RptRulesField(etblRptRules_RuleID), lIndex - 1) = m.Rules.Item(lIndex).RuleID
            m.tblRptRules(RptRulesField(etblRptRules_RuleName), lIndex - 1) = m.Rules.Item(lIndex).Name
            Select Case UCase(m.Rules.Item(lIndex).OrderPlacement)
                Case "MARKET ON CLOSE", "STOP CLOSE ONLY"
                    m.tblRptRules(RptRulesField(etblRptRules_OrderTypeID), lIndex - 1) = 1
                Case Else
                    m.tblRptRules(RptRulesField(etblRptRules_OrderTypeID), lIndex - 1) = 0
            End Select
        Next lIndex
    End With

'FileFromString "SysRules.TXT", m.tblSystemRules.ToString(vbCrLf, vbTab)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSystem.FillSystemRulesTable", eGDRaiseError_Raise
    
End Function

Private Sub FillSystemParmsTable()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lRule As Long
    Dim lItem As Long
    
    With m.tblSystemParms
        .NumRecords = 0
        lItem = 0
        
        For lRule = 1 To m.Rules.Count
            For lIndex = 1 To m.Rules.Item(lRule).Inputs.Count
                If m.Rules.Item(lRule).Inputs.Item(lIndex).ParmTypeID <> 5 Then
                    .Num(SystemParmField(etblSystemParm_ID), lItem) = m.Rules.Item(lRule).Inputs.Item(lIndex).ParmID
                    .Item(SystemParmField(etblSystemParm_Name), lItem) = m.Rules.Item(lRule).Inputs.Item(lIndex).ParmName
                    .Num(SystemParmField(etblSystemParm_SysID), lItem) = m.lSystemNumber
                    .Num(SystemParmField(etblSystemParm_RuleID), lItem) = m.Rules.Item(lRule).Inputs.Item(lIndex).RuleID
                    .Num(SystemParmField(etblSystemParm_TypeID), lItem) = m.Rules.Item(lRule).Inputs.Item(lIndex).ParmTypeID
                    .Item(SystemParmField(etblSystemParm_Value), lItem) = Str(m.Rules.Item(lRule).Inputs.Item(lIndex).Value)
                    .Num(SystemParmField(etblSystemParm_OptFromValue), lItem) = m.Rules.Item(lRule).Inputs.Item(lIndex).OptFromValue
                    .Num(SystemParmField(etblSystemParm_OptToValue), lItem) = m.Rules.Item(lRule).Inputs.Item(lIndex).OptToValue
                    .Num(SystemParmField(etblSystemParm_OptStepValue), lItem) = m.Rules.Item(lRule).Inputs.Item(lIndex).OptStepValue
                    .Num(SystemParmField(etblSystemParm_IfOptimize), lItem) = m.Rules.Item(lRule).Inputs.Item(lIndex).IfOptimize
                    lItem = lItem + 1
                End If
            Next lIndex
        Next lRule
    End With

'FileFromString "SysParms.TXT", m.tblSystemParms.ToString(vbCrLf, vbTab)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSystem.FillSystemParmsTable", eGDRaiseError_Raise
    
End Sub

Friend Sub SetPrivateData(mCopyFrom As mPrivate, Optional ByVal lNewSystemID As Long = 0&)
    
    ' get copy of all private data
    m = mCopyFrom
    m.lSystemNumber = lNewSystemID

    Set m.Rules = mCopyFrom.Rules.MakeCopy(lNewSystemID)
    Set m.Markets = mCopyFrom.Markets.MakeCopy
    If Not mCopyFrom.Bars Is Nothing Then Set m.Bars = mCopyFrom.Bars.MakeCopy
    If Not mCopyFrom.RptBridges Is Nothing Then Set m.RptBridges = mCopyFrom.RptBridges.MakeCopy
    If Not m.astrSingleIterationParms Is Nothing Then
        Set m.astrSingleIterationParms = mCopyFrom.astrSingleIterationParms.MakeCopy
    End If
    
    Set m.tblSystem = mCopyFrom.tblSystem.MakeCopy
    Set m.tblSystemRules = mCopyFrom.tblSystemRules.MakeCopy
    Set m.tblSystemParms = mCopyFrom.tblSystemParms.MakeCopy
    Set m.tblRptRules = mCopyFrom.tblRptRules.MakeCopy
    Set m.tblRptMarkets = mCopyFrom.tblRptMarkets.MakeCopy

End Sub

Public Function MakeCopy(Optional ByVal lNewSystemID As Long = 0&) As cSystem
    
    Dim aCopy As New cSystem
    
    aCopy.SetPrivateData m, lNewSystemID
    Set MakeCopy = aCopy

End Function

Private Function HasLinkedRules() As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long
    
    With m.Rules
        For lIndex = 1 To .Count
            If Len(.Item(lIndex).LinkedRules) > 0 Then
                HasLinkedRules = True
                Exit For
            End If
        Next
    End With

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSystem.HasLinkedRules", eGDRaiseError_Raise

End Function

Public Sub ClearBars()
On Error GoTo ErrSection:

    Set m.Bars = Nothing
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSystem.ClearBars", eGDRaiseError_Raise
    
End Sub

Public Sub ShowMergedReports(ByVal astrTradeFiles As cGdArray, Optional ByVal strReportName As String = "")
On Error GoTo ErrSection:

    Dim lIndex As Long
    Dim RptBridge As New cRptBridge

    For lIndex = m.RptBridges.Count To 1 Step -1
        If m.RptBridges(lIndex).IsLoaded = False Then
            m.RptBridges.Remove lIndex
        End If
    Next lIndex

    If Len(strReportName) = 0 Then
        If Len(m.strSystemName) > 0 Then
            strReportName = m.strSystemName
        Else
            strReportName = "New Strategy"
        End If
    End If

    Set RptBridge = New cRptBridge
    mSysNav.ShowMergedReports RptBridge, strReportName, m.bPyramid, astrTradeFiles.ArrayHandle, m.tblRptRules.TableHandle
    m.RptBridges.Add RptBridge
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSystem.ShowMergedReports", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AutoTradeOrders
'' Description: Generate orders for automated trading
'' Inputs:      Mode, Trades, Next Bar Report, Auto Trading Item ID, Error
''     strFillLastBarAtLimitPrice = "" (none), "High", "Low", or "Both" (if limit fills exist at both high/low)
'' Returns:     True if successful, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AutoTradeOrders(ByVal nMode As eGDAutoTradeOrdersMode, astrTrades As cGdArray, astrNextBar As cGdArray, astrRuleBased As cGdArray, ByVal lAutoTradeItemID As Long, Optional strError As String, _
                    Optional ByVal strFillLastBarAtLimitPrice As String = "") As Boolean
On Error GoTo ErrSection:

    Dim dOrdersDate As Double           ' Date to generate orders for
    Dim strTradesFile As String         ' Trades file
    Dim strNextBarFile As String        ' Name of the next bar order report file
    Dim strRuleBasedFile As String      ' Name of the rule-based next bar order report file
    Dim bReturn As Boolean              ' Return value for the function
    Dim i&, dDate#

    bReturn = False

    ' Figure out the name of the next bar file...
    strNextBarFile = AddSlash(App.Path) & "Trades\NB_ATO_" & Str(lAutoTradeItemID) & ".TXT"
    strRuleBasedFile = Replace(strNextBarFile, "\NB", "\RB")
    strTradesFile = Replace(strNextBarFile, "\NB", "\T")
    
    ' Figure out the date to generate orders for...
    If ((m.Bars.Prop(eBARS_Periodicity) = 0) And (m.Bars.Prop(eBARS_PriceHasSettled) <> 0)) Or (nMode = eGDAutoTradeOrdersMode_OnClose) Then
        dOrdersDate = m.Bars(eBARS_DateTime, m.Bars.Size - 1) + 1
        Do While Not IsWeekday(dOrdersDate)
            dOrdersDate = dOrdersDate + 1
        Loop
    Else
        dOrdersDate = m.Bars(eBARS_DateTime, m.Bars.Size - 1)
        If nMode = eGDAutoTradeOrdersMode_NewBar Then
            m.Bars(eBARS_High, m.Bars.Size - 1) = m.Bars(eBARS_Open, m.Bars.Size - 1)
            m.Bars(eBARS_Low, m.Bars.Size - 1) = m.Bars(eBARS_Open, m.Bars.Size - 1)
            m.Bars(eBARS_Close, m.Bars.Size - 1) = m.Bars(eBARS_Open, m.Bars.Size - 1)
        End If
    End If
    
    If dOrdersDate = kNullData Then
        bReturn = False
        strError = "No data available for " & m.Bars.Prop(eBARS_Symbol)
    Else
        ' Run the engine to generate orders...
        If nMode = eGDAutoTradeOrdersMode_OnClose Then
            ' TLB 6/30/2015: when doing the "on close" check for intraday bars,
            ' we need to make sure the time of the "next bar" is the next session
            ' (to insure "on close" orders and any TradeSense functions that are
            '  checking for the last bar of the day will all work correctly)
            If m.Bars(eBARS_Close, m.Bars.Size - 1) <> kNullData Then
                m.Bars.AddForecastBars 1 ' I believe a forecast bar should exist even for daily bars (?)
            End If
            If m.Bars.IsIntraday And (m.Bars.Size > 0) Then
                ' first find last good data bar
                For i = m.Bars.Size - 1 To 0 Step -1
                    If m.Bars(eBARS_Close, i) <> kNullData Then
                        m.Bars.Size = i + 2 ' (so includes just one forecast bar)
                        ' make sure "next bar" is the next trading session
                        If m.Bars.SessionDate(i + 1) <= m.Bars.SessionDate(i) Then
                            ' just set "next bar" to 1 minute past the start time of the next session
                            ' (this is good enough to make sure all the "On Close" orders will trigger)
                            dDate = m.Bars.SessionDate(i) + 1
                            Do While IsWeekday(dDate) = False
                                dDate = dDate + 1
                            Loop
                            dDate = dDate + (m.Bars.Prop(eBARS_StartTime) + 1) / 1440#
                            If m.Bars.Prop(eBARS_EndTime) < m.Bars.Prop(eBARS_StartTime) Then
                                dDate = dDate - 1 ' (previous evening)
                            End If
                            m.Bars(eBARS_DateTime, i + 1) = dDate
                        End If
                        Exit For
                    End If
                Next
            End If
            
            'bReturn = RunEngine(True, groNextBar, False, strNextBarFile, , False, , astrTrades, dOrdersDate, 0, , True, strError)
            bReturn = RunEngine(False, groSingleIteration, False, strTradesFile, , False, , astrTrades, , , , True, strError, strFillLastBarAtLimitPrice)
        Else
            bReturn = RunEngine(True, groCurrentBarOrders, False, strNextBarFile, , False, , astrTrades, dOrdersDate, 0, , True, strError, strFillLastBarAtLimitPrice)
            
            ' Return the next bar and rule based reports...
            astrNextBar.FromFile strNextBarFile
            astrRuleBased.FromFile strRuleBasedFile
        End If
    End If
    
    AutoTradeOrders = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSystem.AutoTradeOrders", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SecondaryMarkets
'' Description: Compile a list of unique secondary symbol/period pairs
'' Inputs:      None
'' Returns:     Array of Symbol/Period pairs
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SecondaryMarkets(Optional ByVal strDefault As String = "") As cGdArray
On Error GoTo ErrSection:

    Dim astrReturn As New cGdArray      ' Array of symbol/period pairs to return
    Dim lIndex As Long                  ' Index into a for loop
    Dim Market As New cInput            ' Item in the markets collection
    
    ' 1) Figure out the periodicity for Market1...
    If Len(strDefault) = 0 Then
        For lIndex = 1 To m.Markets.Count
            Set Market = m.Markets(lIndex)
            If UCase(Market.ParmName) = "MARKET1" Then
                strDefault = Market.Periodicity
                Exit For
            End If
        Next lIndex
    End If
    
    ' 2) Compile a list of secondary markets replacing (Default) with Market1 periodicity...
    For lIndex = 1 To m.Markets.Count
        Set Market = m.Markets(lIndex)
        If UCase(Market.ParmName) <> "MARKET1" Then
            If UCase(Market.Periodicity) = "(DEFAULT)" Then
                astrReturn.Add Market.Symbol & ";" & strDefault
            Else
                astrReturn.Add Market.Symbol & ";" & Market.Periodicity
            End If
        End If
    Next lIndex
    
    ' 3) Sort the list and get rid of any duplicates...
    astrReturn.Sort eGdSort_DeleteDuplicates
    
    ' 4) Return the array...
    Set SecondaryMarkets = astrReturn

ErrExit:
    Set astrReturn = Nothing
    Exit Function
    
ErrSection:
    Set astrReturn = Nothing
    RaiseError "cSystem.SecondaryMarkets", eGDRaiseError_Raise
    
End Function

' TLB: used for the NavReports
Private Sub SaveDailyBars(ByVal strSymbol$)
On Error GoTo ErrSection:

    Dim strFile$
    Dim Bars As cGdBars
    Static nPrevLDD As Long
    
If Not IsIDE Then Exit Sub
    
    ' only need to do this once for each symbol -- until a new daily download
    If nPrevLDD <> LastDailyDownload Then
        ' so clear out all the old files after a new daily download
        nPrevLDD = LastDailyDownload
        KillFile App.Path & "\Trades\*.BDA"
        'KillFile App.Path & "\Trades\*.GDB"
    End If
    
    strFile = App.Path & "\Trades\" & strSymbol & ".BDA"
    If Not FileExist(strFile) Then
        ' save full history of daily bars for NavReports to use
        Set Bars = New cGdBars
        SetBarProperties Bars, strSymbol, True ' (for this purpose, use default tick move)
        DM_GetBars Bars, strSymbol, "Daily"
        Bars.SerializeDataArrays strFile, True
        'strFile = ReplaceFileExt(strFile, ".GDB")
        'Bars.ToFile "GDB", strFile
        Set Bars = Nothing
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSystem.SaveDailyBars"
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HasOnCloseRules
'' Description: Does the system have On-Close rules?
'' Inputs:      Only if Selected?
'' Returns:     True if has On-Close rules, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function HasOnCloseRules(Optional ByVal bOnlyIfSelected As Boolean = True) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    
    bReturn = False
    For lIndex = 1 To m.Rules.Count
        If InStr(UCase(m.Rules.Item(lIndex).OrderPlacement), "CLOSE") <> 0 Then
            If (m.Rules.Item(lIndex).Selected = True) Or (bOnlyIfSelected = False) Then
                bReturn = True
                Exit For
            End If
        End If
    Next lIndex
    
    HasOnCloseRules = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSystem.HasOnCloseRules"
    
End Function

' Gets the $Risk (for money management purposes) -- looks for the StopLoss-type rules
Private Function GetRiskAmount(bLong As Boolean) As Double
On Error GoTo ErrSection:
    
    Dim iRule&, iInput&, d#, dRisk#, s$
    Dim Rule As cRule
    Dim Parm As cInput
    
    ' TLB 3/15/2015: look for Stop Loss Amount (to base risk on for money management)
    ' (can also look for a Trailing Stop but only if it's checked for "Exit on Entry bar")
    For iRule = 1 To m.Rules.Count
        Set Rule = m.Rules.Item(iRule)
        ' only look at selected (turned on) exit rules that exit on the bar of entry
        If Rule.Selected = True And Rule.RuleUse = 1 And Rule.ExitOnEntryBar = True Then
            ' and only look at Long Exits or Short Exits depending on how called
            If Rule.BuySell <> bLong Then
            'If iDirection = 0 Or (iDirection > 0 And Rule.BuySell = False) Or (iDirection < 0 And Rule.BuySell = True) Then
                '~04011Entry Price ~08001- ~01016Dollars to Price ~16001( ~05016Stop Loss Amount ~17001)
                '~04021Highest High in Trade ~08001- ~01016Dollars to Price ~16001( ~05020Trailing Stop Amount ~17001)
                s = UCase(Rule.Price)
                If InStr(s, UCase("~01016Dollars to Price ~16001( ~05016Stop Loss Amount ~17001)")) > 0 _
                Or InStr(s, UCase("~01016Dollars to Price ~16001( ~05014StopLossAmount ~17001)")) > 0 _
                Or InStr(s, UCase("~01016Dollars to Price ~16001( ~05020Trailing Stop Amount ~17001)")) > 0 _
                Or InStr(s, UCase("~01016Dollars to Price ~16001( ~05018TrailingStopAmount ~17001)")) > 0 Then
                    ' but for risk purposes, need to ignore rules that have a condition using "Bars Since Entry"
                    If InStr(UCase(Rule.Cond), "BARS SINCE ENTRY") = 0 Then
                        For iInput = 1 To Rule.Inputs.Count
                            Set Parm = Rule.Inputs.Item(iInput)
                            s = StripStr(Trim(UCase(Parm.ParmName)), " ")
                            If s = "STOPLOSSAMOUNT" Or s = "TRAILINGSTOPAMOUNT" Then
                                d = Val(Parm.Value)
                                If d > 0 And (d < dRisk Or dRisk = 0) Then
                                    dRisk = d
                                End If
                            End If
                        Next
                    End If
                End If
            End If
        End If
    Next
    Set Parm = Nothing
    Set Rule = Nothing

ErrExit:
    GetRiskAmount = dRisk
    Exit Function
    
ErrSection:
    RaiseError "cSystem.GetRiskAmount"
End Function
