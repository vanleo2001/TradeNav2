VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTrades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cTrades.cls
'' Description: Stores trades and system information
'' Author:      Genesis Financial Data Services
''              425 E Woodmen Rd
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    tblHeader As cGdTable
    tblData As cGdTable
    alSortKey As cGdArray
    
    lLast As Long
    
    astrSignals As cGdArray
    astrSymbols As cGdArray

    ' Used for setting time of trades
    Rules As cGdTable
    aRuleIndex As cGdArray
    aRuleKey As cGdArray
    
    ' Handles into the tables and arrays
    hData As Long
    hHeader As Long
    hSortKey As Long
    hRules As Long
    hRuleIndex As Long
    hRuleKey As Long
    hSymbols As Long
    
    bLoadDone As Boolean
End Type
Private m As mPrivate

Property Get SignalsHandle() As Long
    SignalsHandle = m.astrSignals.ArrayHandle
End Property

'Properties for setting and retrieving header values
Property Get NumHdr(ByVal pRecordNbr, ByVal pField As enth_Trades) As Double
    NumHdr = m.tblHeader.Num(pField, pRecordNbr)
End Property
Property Get ItemHdr(ByVal pRecordNbr, ByVal pField As enth_Trades) As Variant
    ItemHdr = m.tblHeader.Item(pField, pRecordNbr)
End Property
Property Let NumHdr(ByVal pRecordNbr, ByVal pField As enth_Trades, ByVal pValue As Double)
    m.tblHeader.Num(pField, pRecordNbr) = pValue
End Property
Property Let ItemHdr(ByVal pRecordNbr, ByVal pField As enth_Trades, ByVal pValue As Variant)
    m.tblHeader.Item(pField, pRecordNbr) = pValue
End Property
Public Property Get Field(ByVal nField As entd_Trades, Optional ByVal bGetCopy As Boolean = True) As cGdArray
    Set Field = m.tblData.FieldArray(nField, bGetCopy)
End Property
Property Get FieldNumHdr(ByVal pData As enth_Trades) As Long
    FieldNumHdr = pData
End Property
Property Get HdrTableHandle() As Long
    If m.tblHeader Is Nothing Then
        HdrTableHandle = 0
    Else
        HdrTableHandle = m.tblHeader.TableHandle
    End If
End Property
Public Sub CopyHdrFromHandle(ByVal NewTableHandle As Long)
    m.tblHeader.CopyFromHandle NewTableHandle
End Sub
Public Sub SetHdrFromHandle(ByVal TableHandle As Long, pAutoDestroy As Boolean)
    m.tblHeader.SetTableHandle TableHandle, pAutoDestroy
End Sub

'Properties for manipulating Trade detail rows and columns
Property Get FieldNum(ByVal pData As entd_Trades) As Long
    FieldNum = pData
End Property
Property Get FieldName(ByVal pData As entd_Trades) As String
    FieldName = m.tblData.FieldName(pData)
End Property
Property Get FieldType(ByVal pData As entd_Trades) As eGdArray_Type
    FieldType = m.tblData.FieldType(pData)
End Property
Property Get Num(ByVal pRecordNbr, ByVal pField As entd_Trades) As Double
    Num = m.tblData.Num(pField, pRecordNbr)
End Property
Property Get Item(ByVal pRecordNbr, ByVal pField As entd_Trades) As Variant
    Item = m.tblData.Item(pField, pRecordNbr)
End Property
Property Get FieldHandle(ByVal pField As entd_Trades) As Long
    FieldHandle = m.tblData.FieldArrayHandle(pField)
End Property
Property Get NumRecords() As Long
    NumRecords = m.tblData.NumRecords
End Property
Property Let Num(ByVal pRecordNbr, ByVal pField As entd_Trades, ByVal pValue As Double)
    m.tblData.Num(pField, pRecordNbr) = pValue
End Property
Property Let Item(ByVal pRecordNbr, ByVal pField As entd_Trades, ByVal pValue As Variant)
    m.tblData.Item(pField, pRecordNbr) = pValue
End Property
Public Function IsNullAt(ByVal pRecordNbr, ByVal pField As entd_Trades) As Boolean
    IsNullAt = m.tblData.IsNullAt(pField, pRecordNbr)
End Function
Public Sub SetSize(psize As Long)
    Dim X       As Long
    For X = 1 To entd_Cols
        gdReserve FieldHandle(X), psize, 0
        gdSetSize FieldHandle(X), psize, 0
    Next X
End Sub
Public Function CreateIndex(Optional ByVal pFilterByField As entd_Trades = -1) As cGdArray
    Set CreateIndex = m.tblData.CreateIndex(pFilterByField)
End Function
Public Function SortIndex(ByVal aIndex As cGdArray, ByVal nSortByField As entd_Trades, _
        Optional ByVal eSortFlags As eGdSort_Flags = eGdSort_Default) As Long
    SortIndex = m.tblData.SortIndex(aIndex, nSortByField, eSortFlags)
End Function
Property Get DtlTableHandle() As Long
    If m.tblData Is Nothing Then
        DtlTableHandle = 0
    Else
        DtlTableHandle = m.tblData.TableHandle
    End If
End Property

Public Sub CopyTableFromHandle(ByVal NewTableHandle As Long)
    m.tblData.CopyFromHandle NewTableHandle
End Sub
Public Sub SetTableFromHandle(ByVal pTableHandle As Long, pAutoDestroy As Boolean)
    m.tblData.SetTableHandle pTableHandle, pAutoDestroy
End Sub

Public Property Get SignalName(ByVal lRuleID As Long) As String
    
    Dim lIndex As Long              ' Index into the table
        
    If gdBinarySearch(m.hRuleKey, lRuleID, lIndex, eGdSort_Default, 0, gdGetSize(m.hRuleKey)) Then
        SignalName = gdGetTableString(m.hRules, 2, gdGetNum(m.hRuleIndex, lIndex))
    End If
    
End Property

Public Property Get Symbol(ByVal lIndex As Long) As String
    Symbol = gdGetStr(m.astrSymbols.ArrayHandle, lIndex)
End Property

Public Property Get RuleNameHandle() As Long
    RuleNameHandle = m.Rules.FieldArrayHandle(2)
End Property
Public Property Get SymbolHandle() As Long
    SymbolHandle = m.astrSymbols.ArrayHandle
End Property

' Wrappers for getting data from/setting data in the header table...
Public Property Get HeaderNum(ByVal HeaderField As enth_Trades, ByVal lRecord As Long) As Double
    HeaderNum = gdGetTableNum(m.hHeader, HeaderField, lRecord)
End Property
Public Property Let HeaderNum(ByVal HeaderField As enth_Trades, ByVal lRecord As Long, ByVal dValue As Double)
    gdSetTableNum m.hHeader, HeaderField, lRecord, dValue
End Property
Public Property Get HeaderStr(ByVal HeaderField As enth_Trades, ByVal lRecord As Long) As String
    HeaderStr = gdGetTableString(m.hHeader, HeaderField, lRecord)
End Property
Public Property Let HeaderStr(ByVal HeaderField As enth_Trades, ByVal lRecord As Long, ByVal strValue As String)
    gdSetTableStr m.hHeader, HeaderField, lRecord, strValue
End Property
Public Property Get HeaderNumRecords() As Long
    HeaderNumRecords = m.tblHeader.NumRecords
End Property

' Wrappers for getting data from/setting data in the data table...
Public Property Get DataNum(ByVal DataField As entd_Trades, ByVal lRecord As Long) As Double
    DataNum = gdGetTableNum(m.hData, DataField, lRecord)
End Property
Public Property Let DataNum(ByVal DataField As entd_Trades, ByVal lRecord As Long, ByVal dValue As Double)
    gdSetTableNum m.hData, DataField, lRecord, dValue
End Property
Public Property Get DataStr(ByVal DataField As entd_Trades, ByVal lRecord As Long) As String
    DataStr = gdGetTableString(m.hData, DataField, lRecord)
End Property
Public Property Let DataStr(ByVal DataField As entd_Trades, ByVal lRecord As Long, ByVal strValue As String)
    gdSetTableStr m.hData, DataField, lRecord, strValue
End Property
Public Property Get DataNumRecords() As Long
    DataNumRecords = m.tblData.NumRecords
End Property

Public Property Get SortHandle() As Long
    SortHandle = m.hSortKey
End Property
Public Property Let SortHandle(ByVal lHandle As Long)
    m.alSortKey.SetArrayHandle lHandle, False
    m.hSortKey = m.alSortKey.ArrayHandle
End Property


'Import delimited trades text file from System Navigator
Public Sub Import(ByVal strPath As String, ByVal lSystemNumber As Long, ByVal bImportHeaderOnly As Boolean, _
        ByVal dbLocationOfRules As Object, ByVal dbLocationOfMarkets As Object)
On Error GoTo ErrSection:
    
    ' Load tblRules and tblMarkets if not done already...
    If Not m.bLoadDone Then
        LoadLookupTables dbLocationOfRules, dbLocationOfMarkets
        SetLookupHandles
        m.bLoadDone = True
    End If
    
#If 0 Then
    ImportTrades strPath, lSystemNumber, bImportHeaderOnly
#Else
    m.tblHeader.NumRecords = 2

    ImportTradeFile strPath, 1&
    
    Sort
    BuildSignalArray
    CalcAccBal
#End If
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTrades.Import", eGDRaiseError_Raise, g.strAppPath
    
End Sub

Public Sub ImportMultiple(ByVal hFiles As Long, ByVal dbRules As Object, ByVal dbMarkets As Object)
On Error GoTo ErrSection:
    
    Dim astrFiles As cGdArray           ' Array from the given handle

    ' Load tblRules and tblMarkets if not done already...
    If Not m.bLoadDone Then
        LoadLookupTables dbRules, dbMarkets
        SetLookupHandles
        m.bLoadDone = True
    End If
    
    Set astrFiles = New cGdArray
    astrFiles.Create eGDARRAY_Strings
    astrFiles.CopyFromHandle hFiles
    
    m.tblHeader.NumRecords = astrFiles.Size + 1

    ImportMultipleTradeFiles astrFiles
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTrades.ImportMultiple", eGDRaiseError_Raise, g.strAppPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CheckSwitch
'' Description: See if we need to switch this line with the one above it
'' Inputs:      Index of line to check, Header index for that line
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CheckSwitch(ByVal lIndex As Long, ByVal lHeader As Long)
On Error GoTo ErrSection:
    
    Dim bSwitchSortKeys As Boolean      ' Do we need to switch the sort keys?
    Dim bEnsureExitHigher As Boolean    ' Make sure exit is higher
    
    ' Initialize variables...
    bSwitchSortKeys = False
    bEnsureExitHigher = False
       
    'This happens with the following.
    '  7/13/82  Entry   L
    '  7/15/82  Entry   S       <switch
    '  7/15/82  Exit    L       <switch
    
    '  7/13/82  Entry   L
    '  7/15/92  Exit    L
    '  7/15/92  Entry   S
    '  7/15/92  Exit    S       <Ensure sort key higher then previous
       
    ' If Entry/Exit on same day then switch order of Entry/Exit in sorting key...
    If gdGetTableNum(m.hData, entd_TradeDate, lIndex) = gdGetTableNum(m.hData, entd_TradeDate, lIndex - 1) Then
        If gdGetTableNum(m.hData, entd_SignalType, lIndex) <> gdGetTableNum(m.hData, entd_SignalType, lIndex - 1) Then
            If gdGetTableNum(m.hData, entd_Position, lIndex) <> gdGetTableNum(m.hData, entd_Position, lIndex - 1) Then
                bSwitchSortKeys = True
            Else
                bEnsureExitHigher = True
            End If
        End If
    End If
    
    ' Switch sort key of Entry and Exit
    If bSwitchSortKeys Then
        gdSetTableStr m.hData, entd_SortKey, lIndex, Format(gdGetTableNum(m.hData, entd_TradeDate, lIndex), "00000.000000") & gdGetTableString(m.hHeader, enth_Symbol, lHeader) & Str(gdGetTableNum(m.hHeader, enth_SystemNumber, lHeader)) & "1"
        gdSetTableStr m.hData, entd_SortKey, lIndex - 1, Format(gdGetTableNum(m.hData, entd_TradeDate, lIndex - 1), "00000.000000") & gdGetTableString(m.hHeader, enth_Symbol, lHeader) & Str(gdGetTableNum(m.hHeader, enth_SystemNumber, lHeader)) & "0"
    Else
        If bEnsureExitHigher Then
            gdSetTableStr m.hData, entd_SortKey, lIndex, Format(gdGetTableNum(m.hData, entd_TradeDate, lIndex), "00000.000000") & gdGetTableString(m.hHeader, enth_Symbol, lHeader) & Str(gdGetTableNum(m.hHeader, enth_SystemNumber, lHeader)) & "2"
        Else
            'Set sorting key: Trade Date + Symbol + Sys# + Signal type
            gdSetTableStr m.hData, entd_SortKey, lIndex, Format(gdGetTableNum(m.hData, entd_TradeDate, lIndex), "00000.000000") & gdGetTableString(m.hHeader, enth_Symbol, lHeader) & Str(gdGetTableNum(m.hHeader, enth_SystemNumber, lHeader)) & Str(gdGetTableNum(m.hData, entd_SignalType, lIndex))
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTrades.CheckSwitch", eGDRaiseError_Raise, g.strAppPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ConvertDate
'' Description: Convert a date to a date/time
'' Inputs:      Date, Signal Type, Rule ID, Header Index
'' Returns:     Date/Time version of the given date
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ConvertDate(ByVal dDate As Double, ByVal nSignalType As Byte, ByVal lRuleID As Long, ByVal lHeader As Long) As Double
On Error GoTo ErrSection:
    
    Dim lPos As Long                    ' Position of rule in the rules collection
    Dim lOrderType As Long              ' Order type of the rule
    
    ' TLB: need to round to nearest second since numbers from trades file can have partial seconds
    dDate = RoundToSecond(dDate)
    
    If gdGetTableNum(m.hHeader, enth_IntraDaySystem, lHeader) = 0 Then
        If 0 Then 'gdBinarySearch(m.hRuleKey, lRuleID, lPos, eGdSort_Default, 0, gdGetSize(m.hRuleKey)) Then
            lOrderType = gdGetTableNum(m.hRules, 1, gdGetNum(m.hRuleIndex, lPos))
            If (lOrderType = 1) Or (nSignalType = 1) Then
                dDate = dDate + gdGetTableNum(m.hHeader, enth_SessionEnd, lHeader)
            Else
                dDate = dDate + gdGetTableNum(m.hHeader, enth_SessionStart, lHeader)
            End If
        End If
    ElseIf g.bShowInLocalTimeZone Then
        dDate = ConvertTimeZone(dDate, HeaderStr(enth_TimeZoneInfo, lHeader), "")
    End If
    
    ConvertDate = dDate

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTrades.ConvertDate", eGDRaiseError_Raise, g.strAppPath
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BuildSignalArray
'' Description: Build a unique list of signal names
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub BuildSignalArray()
On Error GoTo ErrSection:
    
    Dim hIndex As Long                  ' Handle to the table index array
    Dim alIndex As cGdArray             ' Table index array
    Dim lIndex As Long                  ' Index into a for loop
    Dim hSignal As Long                 ' Handle to the Signals array in the table
    
    If m.tblData.NumRecords - 1 <= 0 Then Exit Sub
    
    Set m.astrSignals = New cGdArray
    m.astrSignals.Create eGDARRAY_Strings
        
    ' Create a Table Index with unique sorted signal names...
    Set alIndex = m.tblData.CreateIndex
    m.tblData.SortIndex alIndex, entd_RuleID, eGdSort_DeleteDuplicates
    hIndex = alIndex.ArrayHandle
    hSignal = FieldHandle(entd_RuleID)

    ' Fill the Signals array...
    For lIndex = 1 To alIndex.Size - 1
        m.astrSignals.Add SignalName(gdGetNum(hSignal, gdGetNum(hIndex, lIndex)))
    Next lIndex
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTrades.BuildSignalArray", eGDRaiseError_Raise, g.strAppPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Sort
'' Description: Sort the trades.  Must be called after the trades are imported
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Sort()
On Error GoTo ErrSection:
    
    ' Set up the Sort Key...
    Set m.alSortKey = m.tblData.CreateIndex
    m.tblData.SortIndex m.alSortKey, FieldNum(entd_SortKey), eGdSort_Stable
    m.hSortKey = m.alSortKey.ArrayHandle
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTrades.Sort", eGDRaiseError_Raise, g.strAppPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CalcAccBal
'' Description: Calculate default values for Tot Profit, dBalance, Cash Allocated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CalcAccBal()
On Error GoTo ErrSection:
    
    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim dBalance As Double              ' Current Balance of the Account
    Dim dTotal As Double                ' Total Profit for this trade
    Dim dAvailableEquity As Double      ' Equity Available at this time
    Dim lDefaultUnits As Long           ' Default units for this system
    
    ' Set initial account balance of trades if blank...
    If m.tblData.NumRecords - 1 > 0 Then
        If gdGetTableNum(m.hData, entd_AccountBalance, 1) <= 0 Then
            dBalance = 0
        Else
            dBalance = gdGetTableNum(m.hData, entd_AccountBalance, 1)
        End If
        dAvailableEquity = dBalance
    End If
        
    ' Set the Available Equity and Account Balance for each transaction...
    For lIndex2 = 1 To m.tblData.NumRecords - 1
        lIndex = gdGetNum(m.hSortKey, lIndex2)
        
        If gdGetTableNum(m.hData, entd_Skip, lIndex) = 0 Then
            If gdGetTableNum(m.hData, entd_SignalType, lIndex) = gEntrySignal Then
                If dAvailableEquity < gdGetTableNum(m.hData, entd_Allocation, lIndex) Then
                    dAvailableEquity = 0
                Else
                    dAvailableEquity = dAvailableEquity - gdGetTableNum(m.hData, entd_Allocation, lIndex)
                End If
            Else
                lDefaultUnits = gdGetTableNum(m.hHeader, enth_DefaultUnits, gdGetTableNum(m.hData, entd_HeaderIndex, lIndex))
                dTotal = (gdGetTableNum(m.hData, entd_Units, lIndex) / lDefaultUnits) * gdGetTableNum(m.hData, entd_Profit, lIndex)
                dBalance = dBalance + dTotal
                dAvailableEquity = dAvailableEquity - gdGetTableNum(m.hData, entd_Allocation, lIndex)
                dAvailableEquity = dBalance
                gdSetTableNum m.hData, entd_TotalProfit, lIndex, dTotal
            End If
            gdSetTableNum m.hData, entd_EquityAvail, lIndex, dAvailableEquity
            gdSetTableNum m.hData, entd_AccountBalance, lIndex, dBalance
        End If
    Next lIndex2
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTrades.CalcAccBal", eGDRaiseError_Raise, g.strAppPath
    
End Sub
           
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SignalTotals
'' Description: Count total entry signals for each trade date/time
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SignalTotals()
On Error GoTo ErrSection:
    
    Dim dDate As Double                 ' Date for the trade
    Dim lCount As Long                  ' Number of occurrences
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim lFirstPos As Long               ' First occurrence
    Dim lIndex As Long                  ' Index into a for loop
    Dim strSymbol As String             ' Symbol of the trade
    
    lFirstPos = 1
    For lIndex = 1 To m.tblData.NumRecords - 1
        ' Set total to zero for exit signals
        If DataNum(entd_SignalType, lIndex) = 1 Then
            DataNum(entd_SignalsTot, lIndex) = 0
        Else
            ' Count total signals for date/time
            dDate = DataNum(entd_TradeDate, lIndex)
            strSymbol = gdGetStr(m.hSymbols, DataNum(entd_SymbolIndex, lIndex))
            lCount = 0
            lFirstPos = lIndex
            Do Until dDate <> DataNum(entd_TradeDate, lIndex) Or DataNum(entd_SignalType, lIndex) = 1 Or _
                    strSymbol <> gdGetStr(m.hSymbols, DataNum(entd_SymbolIndex, lIndex))
                lCount = lCount + 1
                lIndex = lIndex + 1
                If lIndex > NumRecords - 1 Then Exit Do
            Loop
            lIndex = lIndex - 1
            
            ' Go back and set total for each trade with date/time...
            If lCount > 1 Then
                For lIndex2 = lFirstPos To lIndex
                    DataNum(entd_SignalsTot, lIndex2) = lCount
                Next lIndex2
            End If
        End If
    Next lIndex
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTrades.SignalTotals", eGDRaiseError_Raise, g.strAppPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadLookupTables
'' Description: Load Rules from the database into a memory table
'' Inputs:      Rules Database, Markets Database
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadLookupTables(dbRules As Object, dbMarkets As Object)
On Error GoTo ErrSection:
    
    Dim rs As Recordset                 ' Recordset into the database
    Dim lIndex As Long                  ' Counter variable
    
    Set m.Rules = New cGdTable
    m.Rules.CreateField eGDARRAY_Longs, 0, "RuleID"
    m.Rules.CreateField eGDARRAY_Shorts, 1, "OrderTypeID"
    m.Rules.CreateField eGDARRAY_Strings, 2, "RuleName"
        
    ' Load Rules table into gdTable
    Set rs = dbRules.OpenRecordset("SELECT * FROM [tblRules] ORDER BY [Name];", dbOpenSnapshot)
    Do Until rs.EOF
        m.Rules.Num(0, lIndex) = rs!RuleID
        Select Case UCase(rs!OrderPlacement)
            Case "MARKET ON CLOSE", "STOP CLOSE ONLY"
                m.Rules.Num(1, lIndex) = 1
            Case Else
                m.Rules.Num(1, lIndex) = 0
        End Select
        m.Rules.Item(2, lIndex) = rs!Name
        rs.MoveNext
        lIndex = lIndex + 1
    Loop
            
ErrExit:
    Set rs = Nothing
    Exit Sub
    
ErrSection:
    Set rs = Nothing
    RaiseError "cTrades.LoadLookupTables", eGDRaiseError_Raise, g.strAppPath
    
End Sub

Public Function SerializeHdr(ByVal strFile As String, ByVal bPut As Boolean) As Boolean
On Error GoTo ErrSection:

    SerializeHdr = True
    If Not m.tblHeader.Serialize(strFile, bPut) Then
        SerializeHdr = False
    End If
    'FileFromString strFile, m.tblHeader.ToString(vbCrLf, vbTab)
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTrades.SerializeHdr", eGDRaiseError_Raise, g.strAppPath
    
End Function

Public Function SerializeDtl(ByVal strFile As String, ByVal bPut As Boolean) As Boolean
On Error GoTo ErrSection:

    SerializeDtl = True
    If Not m.tblData.Serialize(strFile, bPut) Then
        SerializeDtl = False
    End If
    'FileFromString strFile, m.tblData.ToString(vbCrLf, vbTab)
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTrades.SerializeDtl", eGDRaiseError_Raise, g.strAppPath
    
End Function

Private Sub Class_Initialize()
On Error GoTo ErrSection:
    
    Dim Hdr         As enth_Trades
    Dim Dtl         As entd_Trades
    Dim Y           As Long
    
    'Create gdTable to store header row
    Set m.tblHeader = New cGdTable
    Hdr = enth_SystemNumber
    m.tblHeader.CreateField eGDARRAY_Longs, Hdr, "SystemNumber"
    Hdr = enth_SystemName
    m.tblHeader.CreateField eGDARRAY_Strings, Hdr, "SystemName"
    Hdr = enth_BarTimeFrame
    m.tblHeader.CreateField eGDARRAY_Strings, Hdr, "BarTimeFrame"
    Hdr = enth_IntraDaySystem
    m.tblHeader.CreateField eGDARRAY_TinyInts, Hdr, "IntraDaySystem"
    Hdr = enth_SystemClass
    m.tblHeader.CreateField eGDARRAY_Strings, Hdr, "SystemClass"
    Hdr = enth_StartDate
    m.tblHeader.CreateField eGDARRAY_Doubles, Hdr, "StartDate"
    Hdr = enth_EndDate
    m.tblHeader.CreateField eGDARRAY_Doubles, Hdr, "EndDate"
    Hdr = enth_TotalBars
    m.tblHeader.CreateField eGDARRAY_Longs, Hdr, "TotalBars"
    Hdr = enth_Expenses
    m.tblHeader.CreateField eGDARRAY_Floats, Hdr, "Expenses"
    Hdr = enth_Symbol
    m.tblHeader.CreateField eGDARRAY_Strings, Hdr, "Symbol"
    Hdr = enth_SymbolKey
    m.tblHeader.CreateField eGDARRAY_Strings, Hdr, "SymbolKey"
    Hdr = enth_TickMove
    m.tblHeader.CreateField eGDARRAY_Doubles, Hdr, "TickMove"
    Hdr = enth_TickValue
    m.tblHeader.CreateField eGDARRAY_Doubles, Hdr, "TickValue"
    Hdr = enth_TickMinMove
    m.tblHeader.CreateField eGDARRAY_Doubles, Hdr, "TickMinMove"
    Hdr = enth_Margin
    m.tblHeader.CreateField eGDARRAY_Floats, Hdr, "Margin"
    Hdr = enth_DefaultUnits
    m.tblHeader.CreateField eGDARRAY_Floats, Hdr, "DefaultUnits"
    Hdr = enth_SecurityType
    m.tblHeader.CreateField eGDARRAY_Strings, Hdr, "SecurityType"
    Hdr = enth_SessionStart
    m.tblHeader.CreateField eGDARRAY_Doubles, Hdr, "SessionStart"
    Hdr = enth_SessionEnd
    m.tblHeader.CreateField eGDARRAY_Doubles, Hdr, "SessionEnd"
    Hdr = enth_TimeZoneInfo
    m.tblHeader.CreateField eGDARRAY_Strings, Hdr, "TimeZoneInfo"
    Hdr = enth_LongStopLoss
    m.tblHeader.CreateField eGDARRAY_Strings, Hdr, "LongStopLoss"
    Hdr = enth_ShortStopLoss
    m.tblHeader.CreateField eGDARRAY_Strings, Hdr, "ShortStopLoss"
    
    'Create gdTable to store detail trade data
    Set m.tblData = New cGdTable
    Dtl = entd_TradeNbr
    m.tblData.CreateField eGDARRAY_Longs, Dtl, "TradeNbr"
    Dtl = entd_Position
    m.tblData.CreateField eGDARRAY_TinyInts, Dtl, "Position"
    Dtl = entd_SignalType
    m.tblData.CreateField eGDARRAY_TinyInts, Dtl, "SignalType"
    Dtl = entd_TradeDate
    m.tblData.CreateField eGDARRAY_Doubles, Dtl, "TradeDate"
    Dtl = entd_Price
    m.tblData.CreateField eGDARRAY_Doubles, Dtl, "Price"
    Dtl = entd_RuleID
    m.tblData.CreateField eGDARRAY_Longs, Dtl, "RuleID"
    ''Dtl = entd_Signal
    ''m.tblData.CreateField eGDARRAY_Strings, Dtl, "Signal"
    ''Dtl = entd_EntryRuleID
    ''m.tblData.CreateField eGDARRAY_Longs, Dtl, "EntryRuleID"
    ''Dtl = entd_ExitRuleID
    ''m.tblData.CreateField eGDARRAY_Longs, Dtl, "ExitRuleID"
    Dtl = entd_Units
    m.tblData.CreateField eGDARRAY_Doubles, Dtl, "Units"
    Dtl = entd_NumShares
    m.tblData.CreateField eGDARRAY_Doubles, Dtl, "NumShares"
    Dtl = entd_Profit
    m.tblData.CreateField eGDARRAY_Doubles, Dtl, "Profit"
    Dtl = entd_TotalProfit
    m.tblData.CreateField eGDARRAY_Doubles, Dtl, "TotalProfit"
    Dtl = entd_AccountBalance
    m.tblData.CreateField eGDARRAY_Doubles, Dtl, "AccountBalance"
    Dtl = entd_Equity
    m.tblData.CreateField eGDARRAY_Doubles, Dtl, "Equity"
    Dtl = entd_EquityMA
    m.tblData.CreateField eGDARRAY_Doubles, Dtl, "EquityMA"
    Dtl = entd_SkipEqFilter
    m.tblData.CreateField eGDARRAY_TinyInts, Dtl, "SkipEqFilter"
    Dtl = entd_FilteredEquity
    m.tblData.CreateField eGDARRAY_Doubles, Dtl, "FilteredEquity"
    Dtl = entd_UnfilteredEquity
    m.tblData.CreateField eGDARRAY_Doubles, Dtl, "UnfilteredEquity"
    ''Dtl = entd_TradeSymbol
    ''m.tblData.CreateField eGDARRAY_Strings, Dtl, "TradeSymbol"
    Dtl = entd_SymbolIndex
    m.tblData.CreateField eGDARRAY_Longs, Dtl, "SymbolIndex"
    Dtl = entd_SysNbr
    m.tblData.CreateField eGDARRAY_Longs, Dtl, "SysNbr"
    Dtl = entd_Allocation
    m.tblData.CreateField eGDARRAY_Doubles, Dtl, "Allocation"
    Dtl = entd_MaxProfit
    m.tblData.CreateField eGDARRAY_Doubles, Dtl, "MaxProfit"
    Dtl = entd_MaxLoss
    m.tblData.CreateField eGDARRAY_Doubles, Dtl, "MaxLoss"
    Dtl = entd_BarsInTrade
    m.tblData.CreateField eGDARRAY_Longs, Dtl, "BarsInTrade"
    Dtl = entd_Skip
    m.tblData.CreateField eGDARRAY_TinyInts, Dtl, "Skip"
    Dtl = entd_SkipRpt
    m.tblData.CreateField eGDARRAY_TinyInts, Dtl, "SkipRpt"
    Dtl = entd_TradeDayOfWeek
    m.tblData.CreateField eGDARRAY_TinyInts, Dtl, "TradeDayOfWeek"
    Dtl = entd_TradeDayOfMonth
    m.tblData.CreateField eGDARRAY_TinyInts, Dtl, "TradeDayOfMonth"
    Dtl = entd_TradeDayOfYear
    m.tblData.CreateField eGDARRAY_Longs, Dtl, "TradeDayOfYear"
    Dtl = entd_Rank
    m.tblData.CreateField eGDARRAY_Longs, Dtl, "Rank"
    Dtl = entd_EquityAvail
    m.tblData.CreateField eGDARRAY_Doubles, Dtl, "EquityAvail"
    Dtl = entd_Conflict
    m.tblData.CreateField eGDARRAY_TinyInts, Dtl, "Conflict"
    Dtl = entd_OpenTrade
    m.tblData.CreateField eGDARRAY_TinyInts, Dtl, "OpenTrade"
    Dtl = entd_Link
    m.tblData.CreateField eGDARRAY_Longs, Dtl, "Link"
    Dtl = entd_EntryExitPtr
    m.tblData.CreateField eGDARRAY_Longs, Dtl, "EntryExitPtr"
    Dtl = entd_SortKey
    m.tblData.CreateField eGDARRAY_Strings, Dtl, "SortKey"
    Dtl = entd_TestNumeric
    m.tblData.CreateField eGDARRAY_Doubles, Dtl, "TestNumeric"
    Dtl = entd_TestString
    m.tblData.CreateField eGDARRAY_Strings, Dtl, "TestString"
    Dtl = entd_OpenTradesTotal
    m.tblData.CreateField eGDARRAY_Longs, Dtl, "OpenTradesTotal"
    Dtl = entd_Msg
    m.tblData.CreateField eGDARRAY_Strings, Dtl, "Msg"
    Dtl = entd_SignalsTot
    m.tblData.CreateField eGDARRAY_Longs, Dtl, "SignalsTot"
    Dtl = entd_Show
    m.tblData.CreateField eGDARRAY_TinyInts, Dtl, "Show"
    Dtl = entd_SignalIndex
    m.tblData.CreateField eGDARRAY_Longs, Dtl, "SignalIndex"
    Dtl = entd_HeaderIndex
    m.tblData.CreateField eGDARRAY_Longs, Dtl, "HeaderIndex"
    Dtl = entd_SortKey2
    m.tblData.CreateField eGDARRAY_Strings, Dtl, "SortKey2"
    
    ' Do this for now since everything assumes that these tables are 1-based...
    m.tblData.NumRecords = 1
    m.tblHeader.NumRecords = 1
    
    Set m.astrSignals = New cGdArray
    m.astrSignals.Create eGDARRAY_Strings
    Set m.astrSymbols = New cGdArray
    m.astrSymbols.Create eGDARRAY_Strings
    m.hSymbols = m.astrSymbols.ArrayHandle
    m.lLast = 0
    
    Set m.alSortKey = New cGdArray
    m.alSortKey.Create eGDARRAY_Longs
    m.hSortKey = m.alSortKey.ArrayHandle
    
    m.hHeader = m.tblHeader.TableHandle
    m.hData = m.tblData.TableHandle
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTrades.Initialize", eGDRaiseError_Raise, g.strAppPath
    
End Sub

Private Sub Class_Terminate()
On Error GoTo ErrSection:
    
    Set m.tblHeader = Nothing
    Set m.tblData = Nothing
    Set m.Rules = Nothing
    Set m.astrSignals = Nothing
    Set m.aRuleIndex = Nothing
    Set m.aRuleKey = Nothing
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTrades.Terminate", eGDRaiseError_Raise, g.strAppPath
    
End Sub

Public Sub ImportWithHandles(ByVal strPath As String, ByVal lSystemNumber As Long, _
                    ByVal bImportHeaderOnly As Boolean, ByVal hTblRules As Long, _
                    ByVal hTblMarkets As Long)
On Error GoTo ErrSection:

    ' Load tblRules and tblMarkets if not done already...
    If Not m.bLoadDone Then
        Set m.Rules = New cGdTable
        m.Rules.CopyFromHandle hTblRules
        SetLookupHandles
        m.bLoadDone = True
    End If
    
#If 0 Then
    ImportTrades strPath, lSystemNumber, bImportHeaderOnly
#Else
    m.tblHeader.NumRecords = 2

    ImportTradeFile strPath, 1&
    
    Sort
    BuildSignalArray
    CalcAccBal
#End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTrades.ImportWithHandles", eGDRaiseError_Raise, g.strAppPath

End Sub

Public Sub ImportMultipleWithHandles(ByVal hFiles As Long, ByVal hTblRules As Long)
On Error GoTo ErrSection:

    Dim astrFiles As cGdArray           ' Array from the given handle

    ' Load tblRules and tblMarkets if not done already...
    If Not m.bLoadDone Then
        Set m.Rules = New cGdTable
        m.Rules.CopyFromHandle hTblRules
        SetLookupHandles
        m.bLoadDone = True
    End If
    
    Set astrFiles = New cGdArray
    astrFiles.Create eGDARRAY_Strings
    astrFiles.CopyFromHandle hFiles
    
    m.tblHeader.NumRecords = astrFiles.Size + 1

    ImportMultipleTradeFiles astrFiles
    
#If 0 Then
FileFromString AddSlash(g.strAppPath) & "Header.TXT", m.tblHeader.ToString(vbCrLf, vbTab)
FileFromString AddSlash(g.strAppPath) & "Data.TXT", m.tblData.ToString(vbCrLf, vbTab)

Dim aIndex2 As cGdArray
Dim lIndex As Long

Set aIndex2 = m.tblData.CreateIndex
m.tblData.SortIndex aIndex2, FieldNum(entd_SortKey)
For lIndex = 1 To aIndex2.Size - 1
    FileFromString AddSlash(g.strAppPath) & "Sorted.TXT", m.tblData.GetRecord(aIndex2(lIndex), vbTab), True, True
Next lIndex
#End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTrades.ImportMultipleWithHandles", eGDRaiseError_Raise, g.strAppPath

End Sub

Private Sub SetLookupHandles()
On Error GoTo ErrSection:

    Dim lPos As Long                    ' Position in the Rules table

    ' Need to add a bogus Rule ID (zero) for the "N/A" function...
    If Not m.Rules.FieldArray(2).BinarySearch("N/A", lPos) Then
        m.Rules.FieldArray(0, False).Add 0, lPos
        m.Rules.FieldArray(1, False).Add 0, lPos
        m.Rules.FieldArray(2, False).Add "N/A", lPos
        m.Rules.NumRecords = m.Rules.NumRecords + 1
    End If
    
    ' Need to add a bogus Rule ID (-1) for the "Manual Entry" function...
    If Not m.Rules.FieldArray(2).BinarySearch("Manual Entry", lPos) Then
        m.Rules.FieldArray(0, False).Add -1, lPos
        m.Rules.FieldArray(1, False).Add 0, lPos
        m.Rules.FieldArray(2, False).Add "Manual Entry", lPos
        m.Rules.NumRecords = m.Rules.NumRecords + 1
    End If
    
    ' Need to add a bogus Rule ID (-2) for the "Manual Exit" function...
    If Not m.Rules.FieldArray(2).BinarySearch("Manual Exit", lPos) Then
        m.Rules.FieldArray(0, False).Add -2, lPos
        m.Rules.FieldArray(1, False).Add 0, lPos
        m.Rules.FieldArray(2, False).Add "Manual Exit", lPos
        m.Rules.NumRecords = m.Rules.NumRecords + 1
    End If
    
    ' Setup Rule lookup indexes...
    Set m.aRuleIndex = m.Rules.CreateIndex
    m.Rules.SortIndex m.aRuleIndex, 0
    Set m.aRuleKey = m.Rules.FieldArray(0, True)
    m.aRuleKey.Sort
    m.hRules = m.Rules.TableHandle
    m.hRuleIndex = m.aRuleIndex.ArrayHandle
    m.hRuleKey = m.aRuleKey.ArrayHandle

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTrades.SetLookupHandles", eGDRaiseError_Raise, g.strAppPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ImportMultipleTradeFiles
'' Description: Import multiple trade files into the trades structures
'' Inputs:      Array of Filenames to Import
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ImportMultipleTradeFiles(ByVal astrTradeFiles As cGdArray)
On Error GoTo ErrSection:

    Dim lFile As Long                   ' Index into a for loop

    ' Process each trade file in the array of filenames...
    For lFile = 0 To astrTradeFiles.Size - 1
        ImportTradeFile astrTradeFiles(lFile), lFile + 1&
    Next lFile

    If m.tblData.NumRecords = 0 Then
        Err.Raise gUserErr, , "Reports cannot be viewed because no trades were found."
    End If
    
    If m.tblData.NumRecords = 3 And gdGetTableNum(m.hData, entd_RuleID, 2) = 0 Then
        Err.Raise gUserErr, , "Reports cannot be viewed because there was only one entry with no exits."
    End If
    
    Sort
    BuildSignalArray
    CalcAccBal
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTrades.ImportMultipleTradeFiles", eGDRaiseError_Raise, g.strAppPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ImportTradeFile
'' Description: Import trade file into the trades structures
'' Inputs:      Name of File to Import
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ImportTradeFile(ByVal strFileName As String, ByVal lHeader As Long)
On Error GoTo ErrSection:

    Dim astrTradeFile As New cGdArray   ' Array holding the individual trade file
    Dim astrLine As New cGdArray        ' Line out of the individual trade file
    Dim lTrade As Long                  ' Index into a for loop
    Static lIndex As Long               ' Index into the data table

    ' Create the working arrays...
    astrTradeFile.Create eGDARRAY_Strings
    astrLine.Create eGDARRAY_Strings
    If lIndex = 0& Then lIndex = 1&

    If astrTradeFile.FromFile(strFileName) = True Then
        ' Get the header information out of the first line...
        astrLine.SplitFields astrTradeFile(0), vbTab
        AddHeaderLine astrLine, lHeader
        
        ' Resize the data table appropriately...
        m.tblData.NumRecords = m.tblData.NumRecords + ((astrTradeFile.Size - 1) * 2)
        
        ' Process and store each trade from the file...
        For lTrade = 1 To astrTradeFile.Size - 1
            astrLine.SplitFields astrTradeFile(lTrade), vbTab
            
            AddEntry astrLine, lIndex, lHeader, lTrade
            lIndex = lIndex + 1
            AddExit astrLine, lIndex, lHeader, lTrade
            lIndex = lIndex + 1
        Next lTrade
    End If

ErrExit:
    Set astrTradeFile = Nothing
    Set astrLine = Nothing
    Exit Sub
    
ErrSection:
    Set astrTradeFile = Nothing
    Set astrLine = Nothing
    RaiseError "cTrades.ImportTradeFile", eGDRaiseError_Raise, g.strAppPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddHeaderLine
'' Description: Add a record to the header table with given information
'' Inputs:      Array of header information, Index into the header table
'' Returns:     Record number of the new header
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AddHeaderLine(ByVal astrHeader As cGdArray, ByVal lHeader As Long) As Long
On Error GoTo ErrSection:

    Dim hLine As Long                   ' Array handle for the data array
    Dim strSymbol As String             ' Symbol

    hLine = astrHeader.ArrayHandle

    gdSetTableNum m.hHeader, enth_SystemNumber, lHeader, Val(gdGetStr(hLine, 0))
    gdSetTableStr m.hHeader, enth_SystemName, lHeader, gdGetStr(hLine, 1)
    gdSetTableStr m.hHeader, enth_BarTimeFrame, lHeader, gdGetStr(hLine, 2)
    gdSetTableNum m.hHeader, enth_StartDate, lHeader, Val(gdGetStr(hLine, 3))
    gdSetTableNum m.hHeader, enth_EndDate, lHeader, Val(gdGetStr(hLine, 4))
    gdSetTableNum m.hHeader, enth_TotalBars, lHeader, Val(gdGetStr(hLine, 5))
    gdSetTableNum m.hHeader, enth_Expenses, lHeader, Val(gdGetStr(hLine, 6))
    strSymbol = gdGetStr(hLine, 7)
    gdSetTableStr m.hHeader, enth_Symbol, lHeader, strSymbol
    m.astrSymbols.Add strSymbol
    gdSetTableNum m.hHeader, enth_TickMove, lHeader, Val(gdGetStr(hLine, 8))
    gdSetTableNum m.hHeader, enth_TickValue, lHeader, Val(gdGetStr(hLine, 9))
    gdSetTableNum m.hHeader, enth_TickMinMove, lHeader, Val(gdGetStr(hLine, 10))
    gdSetTableNum m.hHeader, enth_Margin, lHeader, Val(gdGetStr(hLine, 11))
    gdSetTableStr m.hHeader, enth_SecurityType, lHeader, gdGetStr(hLine, 12)
    If Len(gdGetStr(hLine, 13)) > 0 Then
        gdSetTableNum m.hHeader, enth_SessionStart, lHeader, Val(gdGetStr(hLine, 13)) / 1440#
    Else
        gdSetTableNum m.hHeader, enth_SessionStart, lHeader, 0.13333 ' 8:00am
    End If
    If Len(gdGetStr(hLine, 14)) > 0 Then
        gdSetTableNum m.hHeader, enth_SessionEnd, lHeader, Val(gdGetStr(hLine, 14)) / 1440#
    Else
        gdSetTableNum m.hHeader, enth_SessionEnd, lHeader, 0.26667   ' 4:00pm
    End If
    gdSetTableStr m.hHeader, enth_TimeZoneInfo, lHeader, gdGetStr(hLine, 15)
    gdSetTableStr m.hHeader, enth_LongStopLoss, lHeader, gdGetStr(hLine, 16)
    gdSetTableStr m.hHeader, enth_ShortStopLoss, lHeader, gdGetStr(hLine, 17)
    
    Select Case gdGetStr(hLine, 12)
        Case "I"
            If InStr(strSymbol, "-") = 0 Then
                gdSetTableStr m.hHeader, enth_SymbolKey, lHeader, "$"
                gdSetTableStr m.hHeader, enth_SystemClass, lHeader, "Index"
                gdSetTableNum m.hHeader, enth_DefaultUnits, lHeader, 100
            Else
                gdSetTableStr m.hHeader, enth_SymbolKey, lHeader, strSymbol
                gdSetTableStr m.hHeader, enth_SystemClass, lHeader, "Forex"
                gdSetTableNum m.hHeader, enth_DefaultUnits, lHeader, 1
            End If
        
        Case "S"
            gdSetTableStr m.hHeader, enth_SymbolKey, lHeader, "!"
            gdSetTableStr m.hHeader, enth_SystemClass, lHeader, "Equity"
            
            ' Number of shares for stocks = TickValue/TickMove
            If Val(gdGetStr(hLine, 8)) = 0 Then
                gdSetTableNum m.hHeader, enth_DefaultUnits, lHeader, 100
            Else
                gdSetTableNum m.hHeader, enth_DefaultUnits, lHeader, CLng(Val(gdGetStr(hLine, 9)) / Val(gdGetStr(hLine, 8)))
            End If
        
        Case Else
            gdSetTableStr m.hHeader, enth_SymbolKey, lHeader, Parse(strSymbol, "-", 1)
            gdSetTableStr m.hHeader, enth_SystemClass, lHeader, "Future"
            gdSetTableNum m.hHeader, enth_DefaultUnits, lHeader, 1
    End Select
    
    If IsIntraday(GetPeriodicity(gdGetStr(hLine, 2))) Then
        gdSetTableNum m.hHeader, enth_IntraDaySystem, lHeader, 1
    Else
        gdSetTableNum m.hHeader, enth_IntraDaySystem, lHeader, 0
    End If
    
    AddHeaderLine = lHeader

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTrades.AddHeaderLine", eGDRaiseError_Raise, g.strAppPath
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddEntry
'' Description: Add an entry record to the data table with given information
'' Inputs:      Array of data, Data Table Index, Header Index, Link Number
'' Returns:     Record number of the new entry record
''
'' Trade line:  Direction, Entry Time, Entry Price, Entry Rule, Exit Time, Exit Price,
''              Exit Rule, Profit, Max Profit, Max Loss, Bars in Trade, Entry Rule ID,
''              Exit Rule ID, TDOW, TDOM, TDOY, -1 if Low before High or 1 otherwise,
''              Quantity
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AddEntry(ByVal astrData As cGdArray, ByVal lIndex As Long, ByVal lHeader As Long, ByVal lLinkNum As Long) As Long
On Error GoTo ErrSection:

    Dim lRulePos As Long                ' Position of Rule in Signal array
    Dim hLine As Long                   ' Array handle for the data array
    Dim strKey$, dTradeDate#, dEntryDate#

    hLine = astrData.ArrayHandle
    
    gdSetTableNum m.hData, entd_TradeNbr, lIndex, lIndex
    If gdGetStr(hLine, 0) = "L" Or gdGetStr(hLine, 0) = "1" Then
        gdSetTableNum m.hData, entd_Position, lIndex, 1
    Else
        gdSetTableNum m.hData, entd_Position, lIndex, 0
    End If
    gdSetTableNum m.hData, entd_SignalType, lIndex, 0
    
    gdSetTableNum m.hData, entd_TradeDate, lIndex, ConvertDate(Val(gdGetStr(hLine, 1)), 0, Val(gdGetStr(hLine, 11)), lHeader)
    gdSetTableNum m.hData, entd_Price, lIndex, Val(gdGetStr(hLine, 2))
    gdSetTableNum m.hData, entd_RuleID, lIndex, Val(gdGetStr(hLine, 11))
    gdBinarySearch m.hRuleKey, Val(gdGetStr(hLine, 11)), lRulePos, eGdSort_Default, 0, gdGetSize(m.hRuleKey)
    gdSetTableNum m.hData, entd_SignalIndex, lIndex, gdGetNum(m.hRuleIndex, lRulePos)
    gdSetTableNum m.hData, entd_Units, lIndex, gdGetTableNum(m.hHeader, enth_DefaultUnits, lHeader)
    gdSetTableNum m.hData, entd_NumShares, lIndex, Val(gdGetStr(hLine, 18))
    gdSetTableNum m.hData, entd_Profit, lIndex, Val(gdGetStr(hLine, 7))
    gdSetTableNum m.hData, entd_TotalProfit, lIndex, 0
    gdSetTableNum m.hData, entd_AccountBalance, lIndex, 0
    gdSetTableNum m.hData, entd_SymbolIndex, lIndex, lHeader - 1
    gdSetTableNum m.hData, entd_MaxProfit, lIndex, Val(gdGetStr(hLine, 8))
    gdSetTableNum m.hData, entd_MaxLoss, lIndex, Val(gdGetStr(hLine, 9))
    gdSetTableNum m.hData, entd_BarsInTrade, lIndex, Val(gdGetStr(hLine, 10))
    gdSetTableNum m.hData, entd_Skip, lIndex, 0
    gdSetTableNum m.hData, entd_Show, lIndex, 1
    gdSetTableNum m.hData, entd_SkipRpt, lIndex, 0
    gdSetTableNum m.hData, entd_TradeDayOfWeek, lIndex, Val(gdGetStr(hLine, 13))
    gdSetTableNum m.hData, entd_TradeDayOfMonth, lIndex, Val(gdGetStr(hLine, 14))
    gdSetTableNum m.hData, entd_TradeDayOfYear, lIndex, Val(gdGetStr(hLine, 15))
    gdSetTableNum m.hData, entd_Rank, lIndex, 0
    gdSetTableNum m.hData, entd_EquityAvail, lIndex, 0
    gdSetTableNum m.hData, entd_Conflict, lIndex, 0
    gdSetTableNum m.hData, entd_OpenTrade, lIndex, 0
    
    gdSetTableNum m.hData, entd_Link, lIndex, Val(Str(gdGetTableNum(m.hHeader, enth_SystemNumber, lHeader)) & Str(lLinkNum) & "0")
    
    gdSetTableNum m.hData, entd_SysNbr, lIndex, gdGetTableNum(m.hHeader, enth_SystemNumber, lHeader)
    gdSetTableNum m.hData, entd_Allocation, lIndex, 0
    gdSetTableNum m.hData, entd_TestNumeric, lIndex, 0
    gdSetTableNum m.hData, entd_OpenTradesTotal, lIndex, 0
    gdSetTableNum m.hData, entd_SignalsTot, lIndex, 0
    gdSetTableNum m.hData, entd_HeaderIndex, lIndex, lHeader
    
    gdSetTableNum m.hData, entd_EntryExitPtr, lIndex, lIndex + 1
    
If 0 Then
    ' OLD Sort Key = TradeDate + Symbol + Sys# + SignalType
    CheckSwitch lIndex, lHeader
Else
    ' TLB 8/14/2012: NEW SORT KEY = TradeDate + EntryDate + Symbol + Sys# + Seq#
    ' (Note: for an ENTRY, the TradeDate = EntryDate)
    dTradeDate = gdGetTableNum(m.hData, entd_TradeDate, lIndex)
    dEntryDate = dTradeDate
    strKey = Format(dTradeDate, "00000.000000") & ";" & Format(dEntryDate, "00000.000000") & ";" _
        & gdGetTableString(m.hHeader, enth_Symbol, lHeader) & ";" _
        & Str(gdGetTableNum(m.hHeader, enth_SystemNumber, lHeader)) & ";" _
        & Format(lIndex, "000000") '& Str(gdGetTableNum(m.hData, entd_SignalType, lIndex))
    gdSetTableStr m.hData, entd_SortKey, lIndex, strKey
End If
    
    AddEntry = lIndex

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTrades.AddEntry", eGDRaiseError_Raise, g.strAppPath
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddExit
'' Description: Add an exit record to the data table with given information
'' Inputs:      Array of data, Data Table Index, Header Index, Link Number
'' Returns:     Record number of the new exit record
''
'' Trade line:  Direction, Entry Time, Entry Price, Entry Rule, Exit Time, Exit Price,
''              Exit Rule, Profit, Max Profit, Max Loss, Bars in Trade, Entry Rule ID,
''              Exit Rule ID, TDOW, TDOM, TDOY, -1 if Low before High or 1 otherwise,
''              Quantity
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AddExit(ByVal astrData As cGdArray, ByVal lIndex As Long, ByVal lHeader As Long, ByVal lLinkNum As Long) As Long
On Error GoTo ErrSection:

    Dim lRulePos As Long                ' Position of Rule in Signal array
    Dim hLine As Long                   ' Array handle for the data array
    Dim strKey$, dTradeDate#, dEntryDate#

    hLine = astrData.ArrayHandle

    gdSetTableNum m.hData, entd_TradeNbr, lIndex, lIndex
    If gdGetStr(hLine, 0) = "L" Or gdGetStr(hLine, 0) = "1" Then
        gdSetTableNum m.hData, entd_Position, lIndex, 1
    Else
        gdSetTableNum m.hData, entd_Position, lIndex, 0
    End If
    gdSetTableNum m.hData, entd_SignalType, lIndex, 1
    gdSetTableNum m.hData, entd_TradeDate, lIndex, ConvertDate(Val(gdGetStr(hLine, 4)), 1, Val(gdGetStr(hLine, 12)), lHeader)
    gdSetTableNum m.hData, entd_Price, lIndex, Val(gdGetStr(hLine, 5))
    gdSetTableNum m.hData, entd_RuleID, lIndex, Val(gdGetStr(hLine, 12))
    gdBinarySearch m.hRuleKey, Val(gdGetStr(hLine, 12)), lRulePos, eGdSort_Default, 0, gdGetSize(m.hRuleKey)
    gdSetTableNum m.hData, entd_SignalIndex, lIndex, gdGetNum(m.hRuleIndex, lRulePos)
    gdSetTableNum m.hData, entd_Units, lIndex, gdGetTableNum(m.hHeader, enth_DefaultUnits, lHeader)
    gdSetTableNum m.hData, entd_NumShares, lIndex, Val(gdGetStr(hLine, 18))
    gdSetTableNum m.hData, entd_Profit, lIndex, Val(gdGetStr(hLine, 7))
    gdSetTableNum m.hData, entd_TotalProfit, lIndex, 0
    gdSetTableNum m.hData, entd_AccountBalance, lIndex, 0
    gdSetTableNum m.hData, entd_SymbolIndex, lIndex, lHeader - 1
    gdSetTableNum m.hData, entd_MaxProfit, lIndex, Val(gdGetStr(hLine, 8))
    gdSetTableNum m.hData, entd_MaxLoss, lIndex, Val(gdGetStr(hLine, 9))
    gdSetTableNum m.hData, entd_BarsInTrade, lIndex, Val(gdGetStr(hLine, 10))
    gdSetTableNum m.hData, entd_Skip, lIndex, 0
    gdSetTableNum m.hData, entd_Show, lIndex, 1
    gdSetTableNum m.hData, entd_SkipRpt, lIndex, 0
    gdSetTableNum m.hData, entd_TradeDayOfWeek, lIndex, Val(gdGetStr(hLine, 13))
    gdSetTableNum m.hData, entd_TradeDayOfMonth, lIndex, Val(gdGetStr(hLine, 14))
    gdSetTableNum m.hData, entd_TradeDayOfYear, lIndex, Val(gdGetStr(hLine, 15))
    gdSetTableNum m.hData, entd_Rank, lIndex, 0
    gdSetTableNum m.hData, entd_EquityAvail, lIndex, 0
    gdSetTableNum m.hData, entd_Conflict, lIndex, 0
    gdSetTableNum m.hData, entd_OpenTrade, lIndex, 0

    gdSetTableNum m.hData, entd_Link, lIndex, Val(Str(gdGetTableNum(m.hHeader, enth_SystemNumber, lHeader)) & Str(lLinkNum) & "1")
    
    gdSetTableNum m.hData, entd_SysNbr, lIndex, gdGetTableNum(m.hHeader, enth_SystemNumber, lHeader)
    gdSetTableNum m.hData, entd_Allocation, lIndex, 0
    gdSetTableNum m.hData, entd_TestNumeric, lIndex, 0
    gdSetTableNum m.hData, entd_OpenTradesTotal, lIndex, 0
    gdSetTableNum m.hData, entd_SignalsTot, lIndex, 0
    gdSetTableNum m.hData, entd_HeaderIndex, lIndex, lHeader
    
    gdSetTableNum m.hData, entd_EntryExitPtr, lIndex, lIndex - 1
    
If 0 Then
    ' OLD Sort Key = TradeDate + Symbol + Sys# + SignalType
    CheckSwitch lIndex, lHeader
Else
    ' TLB 8/14/2012: NEW SORT KEY = TradeDate + EntryDate + Symbol + Sys# + Seq#
    ' (Note: for an EXIT, add fraction of a second if the trade is still "open" so will sort to bottom)
    dTradeDate = gdGetTableNum(m.hData, entd_TradeDate, lIndex)
    dEntryDate = gdGetTableNum(m.hData, entd_TradeDate, lIndex - 1)
    If gdGetTableNum(m.hData, entd_RuleID, lIndex) = 0 Then
        dTradeDate = dTradeDate + 0.000001
    End If
    strKey = Format(dTradeDate, "00000.000000") & ";" & Format(dEntryDate, "00000.000000") & ";" _
        & gdGetTableString(m.hHeader, enth_Symbol, lHeader) & ";" _
        & Str(gdGetTableNum(m.hHeader, enth_SystemNumber, lHeader)) & ";" _
        & Format(lIndex, "000000") '& Str(gdGetTableNum(m.hData, entd_SignalType, lIndex))
    gdSetTableStr m.hData, entd_SortKey, lIndex, strKey
End If
    
    ' Set the SortKey2 for the entry to the SortKey for the exit and set the SortKey2 for the
    ' exit to a blank string since it is not needed...
    m.tblData.Item(entd_SortKey2, lIndex - 1) = m.tblData.Item(entd_SortKey, lIndex)
    m.tblData.Item(entd_SortKey2, lIndex) = ""
    
    AddExit = lIndex

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTrades.AddExit", eGDRaiseError_Raise, g.strAppPath
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MakeCopy
'' Description: Make a true copy of the trades class
'' Inputs:      None
'' Returns:     Copy of the Trades class
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MakeCopy() As cTrades
On Error GoTo ErrSection:
    
    Dim NewTrades As New cTrades        ' Copy to return
    
    NewTrades.SetPrivateData m
    Set MakeCopy = NewTrades

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTrades.MakeCopy", eGDRaiseError_Raise, g.strAppPath
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetPrivateData
'' Description: Set the private data for a copy of the class
'' Inputs:      mPrivate structure to copy
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub SetPrivateData(mCopyFrom As mPrivate)
On Error GoTo ErrSection:

    m = mCopyFrom
    
    Set m.alSortKey = mCopyFrom.alSortKey.MakeCopy
    m.hSortKey = m.alSortKey.ArrayHandle
    
    Set m.aRuleIndex = mCopyFrom.aRuleIndex.MakeCopy
    m.hRuleIndex = m.aRuleIndex.ArrayHandle
    
    Set m.aRuleKey = mCopyFrom.aRuleKey.MakeCopy
    m.hRuleKey = m.aRuleKey.ArrayHandle
    
    Set m.astrSignals = mCopyFrom.astrSignals.MakeCopy
    
    Set m.astrSymbols = mCopyFrom.astrSymbols.MakeCopy
    m.hSymbols = m.astrSymbols.ArrayHandle
    
    Set m.tblData = mCopyFrom.tblData.MakeCopy
    m.hData = m.tblData.TableHandle
    
    Set m.tblHeader = mCopyFrom.tblHeader.MakeCopy
    m.hHeader = m.tblHeader.TableHandle
    
    Set m.Rules = mCopyFrom.Rules.MakeCopy
    m.hRules = m.Rules.TableHandle

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTrades.SetPrivateData", eGDRaiseError_Raise, g.strAppPath
    
End Sub
