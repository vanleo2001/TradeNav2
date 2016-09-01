VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTodaysFillsUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cTodaysFillsUI.cls
'' Description: Wrapper class for a working orders grid
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 03/15/2010   DAJ         Fixed the grid column information persistence
'' 05/11/2011   DAJ         Utilize CheckSimTradeStatus, EnableCheckStatusMenu
'' 09/23/2011   DAJ         Show date journals form instead of old journals form
'' 10/04/2011   DAJ         Call the ShowJournals function instead of calling the form direct
'' 09/02/2014   DAJ         Move Journal stuff into Journal DLL
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Const kMenuPrefix = "F:"

Private WithEvents m_fgTodaysFills As VSFlexGrid
Attribute m_fgTodaysFills.VB_VarHelpID = -1
Private WithEvents m_tmrMenu As Timer
Attribute m_tmrMenu.VB_VarHelpID = -1
Private WithEvents m_mnuPrint As Menu
Attribute m_mnuPrint.VB_VarHelpID = -1
Private WithEvents m_mnuTradeHistory As Menu
Attribute m_mnuTradeHistory.VB_VarHelpID = -1
Private WithEvents m_mnuSettings As Menu
Attribute m_mnuSettings.VB_VarHelpID = -1
Private WithEvents m_mnuCheckStatus As Menu
Attribute m_mnuCheckStatus.VB_VarHelpID = -1
Private WithEvents m_mnuViewJournals As Menu
Attribute m_mnuViewJournals.VB_VarHelpID = -1
Private WithEvents m_mnuAutoSizeColumns As Menu
Attribute m_mnuAutoSizeColumns.VB_VarHelpID = -1
Private WithEvents m_mnuDefaultColumns As Menu
Attribute m_mnuDefaultColumns.VB_VarHelpID = -1

Private Enum eGDCols
    eGDCol_FillID = 0
    eGDCol_AccountID
    eGDCol_AccountType
    eGDCol_AutoTradeItemID
    eGDCol_SymbolID
    eGDCol_Remove
    
    eGDCol_SessionDate
    eGDCol_Time
    eGdCol_BuySell
    eGdCol_Quantity
    eGDCol_Symbol
    eGDCol_Price
    eGDCol_Account
    eGDCol_AutoTradeItem
    eGDCol_BrokerID
    eGDCol_ClosedProfit
    eGDCol_SessionQty
    eGDCol_SessionProfit
    eGDCol_NumCols
End Enum

Private Type mPrivate
    strSource As String                 ' Name for the calling form
    TCGrid As cTradeConsoleGrid         ' Trade console grid object for generic routines
    tmrRealTime As Timer                ' Realtime timer control
    frm As Form                         ' Form that the grid is on
    SnapshotFills As cGdTree            ' Collection of snapshot order collections by broker
    
    mnuTodaysFills As Object            ' Today's fills menu
End Type
Private m As mPrivate

Private Function GDCol(ByVal nCol As eGDCols) As Long
    GDCol = nCol
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Init
'' Description: Initialize the object
'' Inputs:      Source, Flex Grid Object, Realtime Timer, Form
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Init(ByVal strSource As String, UI As cTodaysFillsControls)
On Error GoTo ErrSection:

    Dim strDisplay As String            ' Display string from INI file

    m.strSource = strSource
    With UI
        Set m.frm = .frm
        Set m_fgTodaysFills = .fgGrid
        Set m.tmrRealTime = .tmrRealTime
        Set m_tmrMenu = .tmrMenu
        
        Set m.mnuTodaysFills = .mnuTodaysFills
        Set m_mnuPrint = .mnuPrint
        Set m_mnuTradeHistory = .mnuTradeHistory
        Set m_mnuSettings = .mnuSettings
        Set m_mnuCheckStatus = .mnuCheckStatus
        Set m_mnuViewJournals = .mnuViewJournals
        Set m_mnuAutoSizeColumns = .mnuAutoSizeColumns
        Set m_mnuDefaultColumns = .mnuDefaultColumns
    End With
    
    Set m.SnapshotFills = New cGdTree
    
    InitFillsGrid
    
    strDisplay = FixDisplay(GetIniFileProperty("FillsDisplay", "", "TTSummary", g.strIniFile))
    
    Set m.TCGrid = New cTradeConsoleGrid
    m.TCGrid.Init m_fgTodaysFills, m.tmrRealTime, strDisplay
    
    LoadFillsGrid
    
    ApplyDisplayString

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTodaysFillsUI.Init"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Update
'' Description: Update the data
'' Inputs:      Broker
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Update(ByVal nBroker As eTT_AccountType)
On Error GoTo ErrSection:

    ' Get the latest collections for the broker...
    Set m.SnapshotFills(nBroker) = g.Broker.SnapshotFillsForBroker(nBroker)
    
    ' Update the orders grid...
    ResetRemoveFillFlag nBroker
    RefreshBrokerFills nBroker
    RemoveFlaggedFills nBroker
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTodaysFillsUI.Update"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ApplyDisplayString
'' Description: Apply the display string from the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ApplyDisplayString()
On Error GoTo ErrSection:

    m.TCGrid.ApplyDisplayString 0&

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTodaysFillsUI.ApplyDisplayString"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenerateReport
'' Description: Set up the print preview form for this grid
'' Inputs:      Arguments passed in from PrintMe
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GenerateReport(ByVal vArgs As Variant)
On Error GoTo ErrSection:

    m.TCGrid.GenerateReport vArgs, "Today's Fills"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTodaysFillsUI.GenerateReport"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SelectedFill
'' Description: Grab the selected fill in the grid
'' Inputs:      None
'' Returns:     Fill (Nothing if not found)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SelectedFill() As cPtFill
On Error GoTo ErrSection:

    Dim Fill As cPtFill                 ' Fill object from the grid

    Set Fill = Nothing
    With m_fgTodaysFills
        If m.TCGrid.ValidRow(.Row) Then
            If TypeOf .RowData(.Row) Is cPtFill Then
                Set Fill = .RowData(.Row)
            End If
        End If
    End With
    
    Set SelectedFill = Fill

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTodaysFillsUI.SelectedFill"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EnableControls
'' Description: Setup the popup menu
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EnableControls()
On Error GoTo ErrSection:

    m_mnuCheckStatus.Enabled = g.Broker.EnableCheckStatusMenu
    m_mnuDefaultColumns.Enabled = (m.TCGrid.Display <> DefaultDisplay)
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTodaysFillsUI.EnableControls"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DefaultDisplay
'' Description: Determine the default display
'' Inputs:      None
'' Returns:     Default Display
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DefaultDisplay() As String
On Error GoTo ErrSection:

    Dim astrDisplay As New cGdArray     ' Array of field display information
    
    astrDisplay.Add "0;Fill ID;1;0;0;0"
    astrDisplay.Add "0;Account ID;1;1;0;0"
    astrDisplay.Add "0;Account Type;1;2;0;0"
    astrDisplay.Add "0;Auto Trade Item ID;1;3;0;0"
    astrDisplay.Add "0;Symbol ID;1;4;0;0"
    astrDisplay.Add "0;Remove;1;5;0;0"
    astrDisplay.Add "0;Session Date;1;6;0;0"
    astrDisplay.Add "1;Time;1;7;0;1"
    astrDisplay.Add "1;Buy/Sell;1;8;0;1"
    astrDisplay.Add "1;Quantity;1;9;0;1"
    astrDisplay.Add "1;Symbol;1;10;0;1"
    astrDisplay.Add "1;Price;1;11;0;1"
    astrDisplay.Add "1;Account;1;12;0;1"
    astrDisplay.Add "1;Auto Trading Item;1;13;0;1"
    astrDisplay.Add "1;Broker ID;1;14;0;1"
    astrDisplay.Add "1;Closed Profit;1;15;0;1"
    astrDisplay.Add "1;Session Quantity;1;16;0;1"
    astrDisplay.Add "1;Session Profit;1;17;0;1"
    
    'DefaultDisplay = "1;Time;1;1,1;Buy/Sell;1;2,1;Quantity;1;3,1;Symbol;1;4,1;Price;1;5,1;Account;1;6,1;Broker ID;1;7"
    DefaultDisplay = "1|" & astrDisplay.JoinFields(",")

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTodaysFillsUI.DefaultDisplay"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FixDisplay
'' Description: Fix the given display string
'' Inputs:      Display
'' Returns:     Fixed Display
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function FixDisplay(ByVal strDisplay As String) As String
On Error GoTo ErrSection:

    If Len(strDisplay) = 0 Then
        strDisplay = DefaultDisplay
    End If
    
    FixDisplay = strDisplay

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTodaysFillsUI.FixDisplay"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillsCol
'' Description: Determine where the given column is really sitting right now
'' Inputs:      Column Number
'' Returns:     Current Position
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function FillsCol(ByVal Col As eGDCols) As Long
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lReturn As Long                 ' Return value for the function
    
    If Col = eGDCol_NumCols Then
        lReturn = Col
    Else
        lReturn = m.TCGrid.ColNumFromColData(Col)
    End If
    
    FillsCol = lReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTodaysFillsUI.FillsCol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ColumnExceptions
'' Description: Build a list of columns not included in the display string
'' Inputs:      None
'' Returns:     Exception String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ColumnExceptions() As String
On Error GoTo ErrSection:

    Dim astrColumns As cGdArray         ' Array of column numbers to join
    
    Set astrColumns = New cGdArray
    astrColumns.Create eGDARRAY_Strings
    
    astrColumns.Add Str(FillsCol(eGDCol_FillID))
    
    ColumnExceptions = "," & astrColumns.JoinFields(",") & ","

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTodaysFillsUI.ColumnExceptions"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitFillsGrid
'' Description: Initialize the fills grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitFillsGrid()
On Error GoTo ErrSection:

    With m_fgTodaysFills
        .Redraw = flexRDNone
        
        .AllowBigSelection = False
        .AllowSelection = False
        .AllowUserResizing = flexResizeColumns
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .BackColorBkg = g.Styler.GetColor(eGrid_Background) 'RH override vbApplicationWorkspacevbButtonFace
        .Editable = flexEDNone
        .ExplorerBar = flexExMove
        .ExtendLastCol = False
        .ScrollBars = flexScrollBarVertical
        .ScrollTrack = True
        .SelectionMode = flexSelectionListBox
        .SheetBorder = RGB(128, 128, 128)
        .HighLight = flexHighlightNever
        
        .Rows = 1
        .FixedRows = 1
        .Cols = GDCol(eGDCol_NumCols)
        .FixedCols = 0
        .FrozenCols = 0
        
        .TextMatrix(0, GDCol(eGDCol_FillID)) = "Fill ID"
        .TextMatrix(0, GDCol(eGDCol_AccountID)) = "Account ID"
        .TextMatrix(0, GDCol(eGDCol_AccountType)) = "Account Type"
        .TextMatrix(0, GDCol(eGDCol_AutoTradeItemID)) = "Auto Trade Item ID"
        .TextMatrix(0, GDCol(eGDCol_SymbolID)) = "Symbol ID"
        .TextMatrix(0, GDCol(eGDCol_Remove)) = "Remove"
        .TextMatrix(0, GDCol(eGDCol_SessionDate)) = "Session Date"
        .TextMatrix(0, GDCol(eGDCol_Time)) = "Time"
        .TextMatrix(0, GDCol(eGdCol_BuySell)) = "Buy/Sell"
        .TextMatrix(0, GDCol(eGdCol_Quantity)) = "Quantity"
        .TextMatrix(0, GDCol(eGDCol_Symbol)) = "Symbol"
        .TextMatrix(0, GDCol(eGDCol_Price)) = "Price"
        .TextMatrix(0, GDCol(eGDCol_Account)) = "Account"
        .TextMatrix(0, GDCol(eGDCol_AutoTradeItem)) = "Auto Trading Item"
        .TextMatrix(0, GDCol(eGDCol_BrokerID)) = "Broker ID"
        .TextMatrix(0, GDCol(eGDCol_ClosedProfit)) = "Closed Profit"
        .TextMatrix(0, GDCol(eGDCol_SessionQty)) = "Session Quantity"
        .TextMatrix(0, GDCol(eGDCol_SessionProfit)) = "Session Profit"
        
        .ColHidden(GDCol(eGDCol_FillID)) = True
        .ColHidden(GDCol(eGDCol_AccountID)) = True
        .ColHidden(GDCol(eGDCol_AccountType)) = True
        .ColHidden(GDCol(eGDCol_AutoTradeItemID)) = True
        .ColHidden(GDCol(eGDCol_SymbolID)) = True
        .ColHidden(GDCol(eGDCol_SessionDate)) = True
        .ColHidden(GDCol(eGDCol_AutoTradeItem)) = True
        .ColHidden(GDCol(eGDCol_Remove)) = True
        
        .ColAlignment(GDCol(eGDCol_Account)) = flexAlignLeftTop
        .ColAlignment(GDCol(eGDCol_BrokerID)) = flexAlignLeftTop
        
        .ColDataType(GDCol(eGDCol_Remove)) = flexDTBoolean
        
        If g.Broker.ShowCents Then
            .ColFormat(GDCol(eGDCol_ClosedProfit)) = "$#,##0.00"
            .ColFormat(GDCol(eGDCol_SessionProfit)) = "$#,##0.00"
        Else
            .ColFormat(GDCol(eGDCol_ClosedProfit)) = "$#,##0"
            .ColFormat(GDCol(eGDCol_SessionProfit)) = "$#,##0"
        End If
        
        .ColFormat(GDCol(eGDCol_Time)) = DateFormat("Format", MM_DD_YYYY, HH_MM_SS, AMPM_UPPER)
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignLeftTop
        
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTodaysFillsUI.InitFillsGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadFillsGrid
'' Description: Load the fills grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadFillsGrid()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim Fill As cPtFill                 ' Fill object
    
    With m_fgTodaysFills
        .Redraw = flexRDNone
        .Rows = .FixedRows
        
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblFills] " & _
                "WHERE tblFills.SessionDate >= " & Str(Date) & ";", dbOpenDynaset)
        Do While Not rs.EOF
            Set Fill = New cPtFill
            If Fill.Load(rs!FillID, rs) Then
                FillToGrid Fill
            End If
            
            rs.MoveNext
        Loop
        
        If .Rows > .FixedRows Then
            .Col = FillsCol(eGDCol_Time)
            .Sort = flexSortGenericAscending
            
            .Row = .Rows - 1
            .RowSel = .Rows - 1
            .ShowCell .Rows - 1, FillsCol(eGDCol_Time)
        End If
        
        m.TCGrid.AutoSizeGrid
        .Redraw = flexRDBuffered
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTodaysFillsUI.LoadFillsGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillToGrid
'' Description: Update the given fill in the grid
'' Inputs:      Fill, Row
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FillToGrid(Fill As cPtFill, Optional ByVal lRow As Long = -1&)
On Error GoTo ErrSection:

    Dim nRedraw As Long                 ' Current state of the grid's redraw
    
    With m_fgTodaysFills
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        If lRow = -1& Then
            .Rows = .Rows + 1&
            lRow = .Rows - 1&
        End If
        
        .RowData(lRow) = Fill
        
        .TextMatrix(lRow, FillsCol(eGDCol_AccountID)) = Str(Fill.AccountID)
        .TextMatrix(lRow, FillsCol(eGDCol_AccountType)) = Str(Fill.Broker)
        .TextMatrix(lRow, FillsCol(eGDCol_AutoTradeItemID)) = Str(Fill.AutoTradingItemID)
        .TextMatrix(lRow, FillsCol(eGDCol_SymbolID)) = Str(Fill.SymbolID)
        
        .TextMatrix(lRow, FillsCol(eGDCol_SessionDate)) = Str(Fill.SessionDate)
        .TextMatrix(lRow, FillsCol(eGDCol_FillID)) = Str(Fill.FillID)
        .TextMatrix(lRow, FillsCol(eGDCol_Time)) = ConvertBrokerDate(Fill.FillDate, Fill.Broker, Fill.Symbol, g.bShowInLocalTimeZone)
        If Fill.Buy = False Then
            .TextMatrix(lRow, FillsCol(eGdCol_BuySell)) = "Sell"
        Else
            .TextMatrix(lRow, FillsCol(eGdCol_BuySell)) = "Buy"
        End If
        .TextMatrix(lRow, FillsCol(eGdCol_Quantity)) = Str(Fill.Quantity)
        .TextMatrix(lRow, FillsCol(eGDCol_Symbol)) = Fill.Symbol
        .TextMatrix(lRow, FillsCol(eGDCol_Price)) = Fill.PriceString
        .TextMatrix(lRow, FillsCol(eGDCol_Account)) = g.Broker.AccountNameForID(Fill.AccountID)
        .TextMatrix(lRow, FillsCol(eGDCol_BrokerID)) = Fill.BrokerID
        .TextMatrix(lRow, FillsCol(eGDCol_ClosedProfit)) = Fill.ClosedProfit
        m.TCGrid.ColorCell lRow, FillsCol(eGDCol_ClosedProfit)
        .TextMatrix(lRow, FillsCol(eGDCol_SessionQty)) = Format(Fill.SessionQuantity, "#,##0")
        .TextMatrix(lRow, FillsCol(eGDCol_SessionProfit)) = Fill.SessionProfit
        m.TCGrid.ColorCell lRow, FillsCol(eGDCol_SessionProfit)
        
        CheckedCell(m_fgTodaysFills, lRow, FillsCol(eGDCol_Remove)) = False
        
        .Redraw = nRedraw
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTodaysFillsUI.FillToGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ResetRemoveFillFlag
'' Description: Set the remove fill flag on all fills for the given broker
'' Inputs:      Broker
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ResetRemoveFillFlag(ByVal nBroker As eTT_AccountType)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    With m_fgTodaysFills
        For lIndex = .FixedRows To .Rows - 1
            If .RowData(lIndex).Broker = nBroker Then
                CheckedCell(m_fgTodaysFills, lIndex, FillsCol(eGDCol_Remove)) = True
            End If
        Next lIndex
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTodaysFillsUI.ResetRemoveFillFlag"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveFlaggedFills
'' Description: Remove rows with the remove flag set for the given broker
'' Inputs:      Broker
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RemoveFlaggedFills(ByVal nBroker As eTT_AccountType)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    With m_fgTodaysFills
        For lIndex = .Rows - 1 To .FixedRows Step -1
            If .RowData(lIndex).Broker = nBroker Then
                If CheckedCell(m_fgTodaysFills, lIndex, FillsCol(eGDCol_Remove)) = True Then
                    .RemoveItem lIndex
                End If
            End If
        Next lIndex
            
        SetBackColors m_fgTodaysFills
        m.TCGrid.AutoSizeGrid
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTodaysFillsUI.RemoveFlaggedFills"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshBrokerFills
'' Description: Refresh fills in the grid for the given broker from the collection
'' Inputs:      Broker
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RefreshBrokerFills(ByVal nBroker As eTT_AccountType)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim lRow As Long                    ' Row in the grid
    Dim Fills As cPtFills               ' Collection of fills for the broker
    
    If Not m.SnapshotFills(nBroker) Is Nothing Then
        Set Fills = m.SnapshotFills(nBroker).MakeCopy
        With m_fgTodaysFills
            For lIndex = 1 To Fills.Count
                lRow = -1&
                For lIndex2 = .FixedRows To .Rows - 1
                    If .RowData(lIndex2).FillID = Fills(lIndex).FillID Then
                        lRow = lIndex2
                        Exit For
                    End If
                Next lIndex2
                    
                FillToGrid Fills(lIndex), lRow
            Next lIndex
        
            If .Rows > .FixedRows Then
                .Col = FillsCol(eGDCol_Time)
                .Sort = flexSortGenericAscending
                
                .Row = .Rows - 1
                .RowSel = .Rows - 1
                .ShowCell .Rows - 1, FillsCol(eGDCol_Time)
            End If
            
            SetBackColors m_fgTodaysFills
            m.TCGrid.AutoSizeGrid
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTodaysFillsUI.RefreshBrokerFills"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    StartMenuTimer
'' Description: Start the menu timer with the given command
'' Inputs:      Command
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub StartMenuTimer(ByVal strCommand As String)
On Error GoTo ErrSection:

    m_tmrMenu.Tag = kMenuPrefix & strCommand
    m_tmrMenu.Enabled = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTodaysFillsUI.StartMenuTimer"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BuildDisplayString
'' Description: Tell the grid class to build the display string
'' Inputs:      Force Sum Calc?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub BuildDisplayString(Optional ByVal bForceSumCalc As Boolean = False)
On Error GoTo ErrSection:

    m.TCGrid.BuildDisplayString ColumnExceptions, bForceSumCalc
    SetIniFileProperty "FillsDisplay", m.TCGrid.Display, "TTSummary", g.strIniFile

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTodaysFillsUI.BuildDisplayString"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_fgTodaysFills_AfterMoveColumn
'' Description: After a user moves a column, resave the display string
'' Inputs:      Column moved, Position moved to
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_fgTodaysFills_AfterMoveColumn(ByVal Col As Long, Position As Long)
On Error GoTo ErrSection:

    If m.frm.Visible Then
        BuildDisplayString
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTodaysFillsUI.m_fgTodaysFills_AfterMoveColumn"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_fgTodaysFills_AfterUserResize
'' Description: Rebuild the display string after a user resize
'' Inputs:      Row and Column of resize
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_fgTodaysFills_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    If m.frm.Visible Then
        BuildDisplayString True
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cTodaysFillsUI.m_fgTodaysFills_AfterUserResize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_fgTodaysFills_BeforeMouseDown
'' Description: Bring up the Pop-Up menu when the user right clicks
'' Inputs:      Button Pressed, Shift/Ctrl/Alt status, Location of Mouse,
''              Whether to Cancel the click
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_fgTodaysFills_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:

    If Button = vbRightButton Then
        m.TCGrid.SelectRow m_fgTodaysFills.MouseRow
        EnableControls
        m.frm.PopupMenu m.mnuTodaysFills
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTodaysFillsUI.m_fgTodaysFills_BeforeMouseDown"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuAutoSizeColumns_Click
'' Description: Allow the user to auto size the columns
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuAutoSizeColumns_Click()
On Error GoTo ErrSection:

    m.TCGrid.AutoSizeGrid True
    BuildDisplayString

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTodaysFillsUI.m_mnuAutoSizeColumns_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuCheckStatus_Click
'' Description: Check the status of any open orders from the trade server
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuCheckStatus_Click()
On Error GoTo ErrSection:

    g.Broker.CheckTradeServerOrders

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTodaysFillsUI.m_mnuCheckStatus_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuDefaultColumns_Click
'' Description: Allow the user to revert back to the default column information
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuDefaultColumns_Click()
On Error GoTo ErrSection:
    
    m.TCGrid.Display = DefaultDisplay
    ApplyDisplayString
    BuildDisplayString

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTodaysFillsUI.m_mnuDefaultColumns"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuPrint_Click
'' Description: Allow the user to print the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuPrint_Click()
On Error GoTo ErrSection:

    m.frm.PrintMe

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTodaysFillsUI.m_mnuPrint_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuSettings_Click
'' Description: Allow the user to view and edit their console settings
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuSettings_Click()
On Error GoTo ErrSection:

    StartMenuTimer "SETTINGS"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTodaysFillsUI.m_mnuSettings_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuTradeHistory_Click
'' Description: Allow the user to view trade history for the given account
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuTradeHistory_Click()
On Error GoTo ErrSection:

    StartMenuTimer "TRADEHISTORY"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTodaysFillsUI.m_mnuTradeHistory_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuViewJournals_Click
'' Description: Allow the user to view their journal entries
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuViewJournals_Click()
On Error GoTo ErrSection:

    StartMenuTimer "JOURNALS"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTodaysFillsUI.m_mnuViewJournals_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_tmrMenu_Timer
'' Description: Perform a specified action when the timer goes off
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_tmrMenu_Timer()
On Error GoTo ErrSection:

    Dim strTag As String                ' Action to perform
    Dim Fill As cPtFill                 ' Fill object

    strTag = m_tmrMenu.Tag
    If Len(strTag) > 2 Then
        If Left(strTag, 2) = kMenuPrefix Then
            strTag = Mid(strTag, 3)
            m_tmrMenu.Tag = ""
            m_tmrMenu.Enabled = False
            
            Select Case UCase(strTag)
                Case "JOURNALS"
                    g.TnJournal.ShowJournals
                    
                Case "SETTINGS"
                    frmTTSummaryCfg.ShowMe
                    
                Case "TRADEHISTORY"
                    Set Fill = SelectedFill
                    If Fill Is Nothing Then
                        frmTTAccounts.ShowMe True
                    Else
                        frmTTPositions.ShowMe Fill.AccountID, Fill.Broker
                    End If
                    
            End Select
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTodaysFillsUI.m_tmrMenu_Timer"
    
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

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTodaysFillsUI.Class_Initialize"

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

    BuildDisplayString

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTodaysFillsUI.Class_Terminate"

End Sub
