VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTnJournal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cTnJournal.cls
'' Description: Application side functions for the Journal DLL
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 10/24/2014   DAJ         Fill Display
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Picture16
'' Description: Grab the appropriate image from the appropriate image list
'' Inputs:      Picture name, Image List
'' Returns:     Image
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Picture16(ByVal strPicture$, Optional ByVal iImageList As Integer = 0) As Object
    Set Picture16 = mMain.Picture16(strPicture, iImageList)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ReportsPicture
'' Description: Grab the image for reports
'' Inputs:      None
'' Returns:     Image
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ReportsPicture() As Object
    If IsWoodiesVersion Then
        Set ReportsPicture = Picture16(ToolbarIcon("ID_TradeFilter"))
    Else
        Set ReportsPicture = Picture16(ToolbarIcon("kPerformance"))
    End If
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ToolbarIcon
'' Description: Return the toolbar icon name for the given ID
'' Inputs:      ID
'' Returns:     Toolbar Icon name
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ToolbarIcon(ByVal strID As String) As String
    ToolbarIcon = mMain.ToolbarIcon(strID)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HasModule
'' Description: Determine if the user has the given module(s)
'' Inputs:      Module(s), Include Source Code?
'' Returns:     True if have module(s), False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function HasModule(ByVal strModules As String, Optional ByVal bIncludeSourceCode As Boolean = False) As Boolean
    HasModule = mMain.HasModule(strModules, bIncludeSourceCode)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CurrentTime
'' Description: If streaming get the feed time, otherwise return now
'' Inputs:      To Time Zone, Symbol, Allow Replay Time?
'' Returns:     Feed Time or Now
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CurrentTime(Optional ByVal strToTimeZone As String = "", Optional ByVal strSymbol As String = "", Optional ByVal bAllowReplayTime As Boolean = False) As Double
    CurrentTime = mDataNav.CurrentTime(strToTimeZone, strSymbol, bAllowReplayTime)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ActiveChartSymbol
'' Description: Get the symbol off of the active chart
'' Inputs:      None
'' Returns:     Symbol off of the active chart ( blank if no active chart )
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ActiveChartSymbol() As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    strReturn = ""
    If Not mMain.ActiveChart Is Nothing Then
        strReturn = mMain.ActiveChart.Chart.Symbol
    End If
    
    ActiveChartSymbol = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTnJournal.ActiveChartSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ActiveChartCaption
'' Description: Build a caption for the active chart
'' Inputs:      None
'' Returns:     Caption for the active chart ( blank if no active chart )
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ActiveChartCaption() As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    strReturn = ""
    If Not mMain.ActiveChart Is Nothing Then
        strReturn = BuildChartCaption(mMain.ActiveChart)
    End If
    
    ActiveChartCaption = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTnJournal.ActiveChartCaption"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadChartsCombo
'' Description: Load a combo box with chart captions
'' Inputs:      Charts Combo, Image Caption
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub LoadChartsCombo(cboCharts As ctlUniComboImageXP, Optional ByVal strImageCaption As String = "")
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim strActiveChart As String        ' Caption for the active chart
    Dim lListIndex As Long              ' List index to select in combo
    Dim strCaption As String            ' Caption for the image

    strActiveChart = ""
    cboCharts.Clear
    lListIndex = -1&
    
    If Not ActiveChart Is Nothing Then
        strActiveChart = BuildChartCaption(ActiveChart)
    End If

    If Len(strImageCaption) > 0 Then
        cboCharts.AddItem strImageCaption
        cboCharts.ItemData(cboCharts.NewIndex) = 0&
        lListIndex = cboCharts.NewIndex
    End If
    
    For lIndex = 0 To Forms.Count - 1
        If IsFrmChart(Forms(lIndex)) = True Then
            strCaption = BuildChartCaption(Forms(lIndex))
            If strCaption <> strImageCaption Then
                cboCharts.AddItem strCaption
                cboCharts.ItemData(cboCharts.NewIndex) = Forms(lIndex).hWnd
                
                If (strCaption = strActiveChart) And (lListIndex = -1&) Then
                    lListIndex = cboCharts.NewIndex
                End If
            End If
        End If
    Next lIndex

    If lListIndex = -1& Then
        cboCharts.ListIndex = 0&
    Else
        cboCharts.ListIndex = lListIndex
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTnJournal.LoadChartsCombo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ChartImageForHwnd
'' Description: Create a chart image for the given hWnd
'' Inputs:      Fill, Chart Caption ( out )
'' Returns:     Filename ( Blank if not found )
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ChartImageForHwnd(ByVal hWnd As Long, Optional strCaption As String) As String
    ChartImageForHwnd = mTradeTracker.ChartImageForHwnd(hWnd, strCaption)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LookupSymbol
'' Description: Allow the user to look up a symbol
'' Inputs:      Key Pressed, Default Symbol
'' Returns:     Symbol selected ( Blank if cancelled out )
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function LookupSymbol(Optional ByVal KeyAscii As Long = 0&, Optional ByVal strDefaultSymbol As String = "") As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim astrSymbol As New cGdArray      ' Array to get lookup symbol from
    
    If KeyAscii = 0& Then
        Set astrSymbol = frmSymbolSelector.ShowMe(strDefaultSymbol, False, True, "Symbol to Buy/Sell", , , True)
    Else
        Set astrSymbol = frmSymbolSelector.ShowMe(Chr(KeyAscii), False, True, "Symbol to Buy/Sell", False, False, True)
    End If
    
    If astrSymbol.Size > 0 Then
        strReturn = astrSymbol(0)
    End If
    
    LookupSymbol = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTnJournal.LookupSymbol", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DoPrintHeader
'' Description: Do the print header for Trade Navigator
'' Inputs:      Font Size, Printer Object
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DoPrintHeader(Optional ByVal nFontSize& = 12, Optional ByVal vp As Object = Nothing)
    mMain.DoPrintHeader nFontSize, vp
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ChangeGridFont
'' Description: Allow the user to change the font for the given grid
'' Inputs:      Grid, Resize Columns?
'' Returns:     True if done, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ChangeGridFont(fg As VSFlexGrid, Optional ByVal bResizeColumns As Boolean = True) As Boolean
    ChangeGridFont = mMain.ChangeGridFont(fg, bResizeColumns)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowTradeFilter
'' Description: Show the trade filter form
'' Inputs:      Account, Symbol, Auto Trade Item ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowTradeFilter(Optional ByVal vAccountNumberOrID As Variant = kNullData, Optional ByVal vSymbolOrSymbolID As Variant = kNullData, Optional ByVal lAutoTradeItemID As Long = kNullData)
    mTradeTracker.ShowTradeFilter vAccountNumberOrID, vSymbolOrSymbolID, lAutoTradeItemID
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderTextForId
'' Description: Get the order text for the given Order ID
'' Inputs:      Order ID, Include Symbol?, Include Account?, Include ID's?
'' Returns:     Order Text
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OrderTextForId(ByVal lOrderID As Long, Optional ByVal bWithSymbol As Boolean = True, Optional ByVal bWithAccount As Boolean = False, Optional ByVal bWithID As Boolean = False) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim Order As cPtOrder               ' Order object
    
    strReturn = ""
    Set Order = New cPtOrder
    If Order.Load(lOrderID) = True Then
        strReturn = Order.OrderText(bWithSymbol, bWithAccount, bWithID)
    End If
    
    OrderTextForId = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTnJournal.OrderTextForId"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OpenRecordset
'' Description: Open a recordset in the Trade Tracker database
'' Inputs:      Query, Type
'' Returns:     Recordset
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OpenRecordset(ByVal strQuery As String, Optional ByVal vType As Variant) As Recordset
    Set OpenRecordset = g.dbPaper.OpenRecordset(strQuery, vType)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetPerformanceReport
'' Description: Get an image of a performance report
'' Inputs:      Capture File ( out ), Current Time ( out )
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetPerformanceReport(strCaptureFile As String, dCurrentTime As Double) As Boolean
On Error GoTo ErrSection:

    Dim strPath As String               ' Path for the saved images
    
    strPath = AddSlash(App.Path) & "SavedImages"

    If DirExist(strPath) = False Then
        MkDir strPath
    End If
    dCurrentTime = CurrentTime
    strCaptureFile = AddSlash(strPath) & Format(dCurrentTime, "YYYYMMDD HHMMSS") & ".RPT"
    
    GetPerformanceReport = frmTradeReportFilter.ShowMe(, False, strCaptureFile)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTnJournal.GetPerformanceReport"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetSymbol
'' Description: Get a symbol for the given Symbol or Symbol ID
'' Inputs:      Symbol or Symbol ID
'' Returns:     Symbol
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetSymbol(ByVal vSymbolOrSymbolID As Variant) As String
    GetSymbol = mDmDll.GetSymbol(vSymbolOrSymbolID)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetSymbolID
'' Description: Get a symbol ID for the given Symbol or Symbol ID
'' Inputs:      Symbol or Symbol ID
'' Returns:     Symbol ID
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetSymbolID(ByVal vSymbolOrSymbolID As Variant) As Long
    GetSymbolID = mDmDll.GetSymbolID(vSymbolOrSymbolID)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderForID
'' Description: Fill an order dictionary for the given ID
'' Inputs:      Order ID
'' Returns:     Order
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OrderForID(ByVal lOrderID As Long) As cBrokerMessage
    Set OrderForID = mTradeTracker.OrderToBrokerMessage(lOrderID)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BuildChartCaption
'' Description: Build an image caption from the given chart form
'' Inputs:      Chart Form
'' Returns:     Image Caption
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function BuildChartCaption(ChartForm As Form) As String
On Error GoTo ErrSection:

    Dim strSymbol As String             ' Symbol for the chart
    Dim strPeriod As String             ' Period for the chart
    Dim strNow As String                ' Current time
    
    strSymbol = ChartForm.Chart.Symbol
    strPeriod = GetPeriodStr(ChartForm.Periodicity)
    strNow = Format(CurrentTime, "YYYY-MM-DD HH:MM:SS")
    
    BuildChartCaption = strSymbol & " " & strPeriod & " " & strNow

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTnJournal.BuildChartCaption"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateJournal
'' Description: Update the given journal entry
'' Inputs:      Journal ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdateJournal(ByVal lJournalID As Long)
On Error GoTo ErrSection:

    If FormIsLoaded("frmTTPositions") Then
        frmTTPositions.UpdateJournal lJournalID
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cTnJournal.UpdateJournal"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OwnerFormIcon
'' Description: Get the owner form icon from the main application
'' Inputs:      None
'' Returns:     Owner Form Icon
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OwnerFormIcon() As StdPicture
    Set OwnerFormIcon = frmMain.Icon
End Function

'******************************************************************************

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowJournals
'' Description: Show the appropriate journals form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowJournals()
On Error GoTo ErrSection:

    If (HasLevel(eTN4_Gold, False) = True) Or (HasModule("DATEJRNL") = True) Then
        g.JournalBridge.ShowDateJournals
    Else
        g.JournalBridge.ShowJournals
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTnJournal.ShowJournals"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowOrderJournal
'' Description: Show the order journal form for the given order
'' Inputs:      Order
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowOrderJournal(ByVal Order As cPtOrder) As Boolean
On Error GoTo ErrSection:

    Dim orderMessage As cBrokerMessage  ' Order as a broker message object
    
    Set orderMessage = mTradeTracker.OrderToBrokerMessage(Order.OrderID, Order)

    ShowOrderJournal = g.JournalBridge.ShowOrderJournal(orderMessage)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTnJournal.ShowOrderJournal"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowOrderJournalForID
'' Description: Show the order journal form for the given order ID
'' Inputs:      Order ID
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowOrderJournalForID(ByVal lOrderID As Long) As Boolean
On Error GoTo ErrSection:

    Dim orderMessage As cBrokerMessage  ' Order as a broker message object
    
    Set orderMessage = mTradeTracker.OrderToBrokerMessage(lOrderID)

    ShowOrderJournalForID = g.JournalBridge.ShowOrderJournal(orderMessage)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTnJournal.ShowOrderJournalForID"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowJournalForTransaction
'' Description: Show the order journal form for the given order
'' Inputs:      Order
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowJournalForTransaction(ByVal Order As cPtOrder) As Boolean
On Error GoTo ErrSection:

    Dim orderMessage As cBrokerMessage  ' Order as a broker message object
    
    Set orderMessage = mTradeTracker.OrderToBrokerMessage(Order.OrderID, Order)

    ShowJournalForTransaction = g.JournalBridge.ShowJournalForTransaction(orderMessage)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTnJournal.ShowJournalForTransaction"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AutoJournalForFill
'' Description: Create a fill journal for the user based on the given fill
'' Inputs:      Fill
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AutoJournalForFill(ByVal Fill As cPtFill)
On Error GoTo ErrSection:

    Dim strChartImage As String         ' Chart image filename
    Dim dCurrentTime As Double          ' Current time
    Dim strChartCaption As String       ' Chart caption
    Dim strAutoExit As String           ' Auto exit name
    Dim strText As String               ' Text for the journal entry
        
    dCurrentTime = CurrentTime(, , True)
    
    strText = mTradeTracker.FillDisplay(Fill, True, True, False, True, False, True, True)
    If g.OrderStrategies.OrderIdExistsInStrategy(Fill.OrderID, strAutoExit) Then
        If Len(strAutoExit) > 0 Then
            strText = strText & ".  Generated from auto exit '" & strAutoExit & "'"
        End If
    End If
        
    If Fill.SymbolID > 0 Then
        strChartImage = ChartImageForFill(Fill, strChartCaption)
    End If
    
    g.JournalBridge.AutoJournalForFill dCurrentTime, strText, strChartCaption, strChartImage

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTnJournal.AutoJournalForFill"
    
End Sub
