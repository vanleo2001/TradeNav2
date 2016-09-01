VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmOptionChain 
   Caption         =   "Form1"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   8325
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   8325
   ShowInTaskbar   =   0   'False
   Begin ActiveToolBars.SSActiveToolBars tbToolbar 
      Left            =   7800
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131083
      ToolBarsCount   =   1
      ToolsCount      =   7
      Tools           =   "frmOptionChain.frx":0000
      ToolBars        =   "frmOptionChain.frx":026A
   End
   Begin VSFlex7LCtl.VSFlexGrid fgUnder 
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   360
      Width           =   2895
      _cx             =   5106
      _cy             =   450
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Timer tmrRealTime 
      Enabled         =   0   'False
      Left            =   7800
      Top             =   5640
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   375
      Left            =   1620
      TabIndex        =   1
      Top             =   5700
      Width           =   5235
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmOptionChain.frx":03E7
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmOptionChain.frx":0413
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmOptionChain.frx":0433
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdSettings 
         Height          =   375
         Left            =   2700
         TabIndex        =   6
         Top             =   0
         Width           =   1215
         _ExtentX        =   0
         _ExtentY        =   0
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmOptionChain.frx":044F
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmOptionChain.frx":0481
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmOptionChain.frx":04A1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdFields 
         Height          =   375
         Left            =   1320
         TabIndex        =   5
         Top             =   0
         Width           =   1275
         _ExtentX        =   0
         _ExtentY        =   0
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmOptionChain.frx":04BD
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmOptionChain.frx":04F9
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmOptionChain.frx":0519
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdClose 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   1215
         _ExtentX        =   0
         _ExtentY        =   0
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmOptionChain.frx":0535
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmOptionChain.frx":0561
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmOptionChain.frx":0581
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdUpdate 
         Height          =   375
         Left            =   4020
         TabIndex        =   2
         Top             =   0
         Width           =   1215
         _ExtentX        =   0
         _ExtentY        =   0
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmOptionChain.frx":059D
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmOptionChain.frx":05CD
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmOptionChain.frx":05ED
         RightToLeft     =   0   'False
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid fg 
      Height          =   4875
      Left            =   0
      TabIndex        =   0
      Top             =   660
      Width           =   8355
      _cx             =   14737
      _cy             =   8599
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   14742776
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmOptionChain.frx":0609
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin HexUniControls.ctlUniLabelXP lblSymbolInfo 
      Height          =   315
      Left            =   0
      Top             =   0
      Width           =   3015
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmOptionChain.frx":06B8
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmOptionChain.frx":06E4
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmOptionChain.frx":0704
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin VB.Menu mnuAdd 
      Caption         =   "Add"
      Begin VB.Menu mnuQuoteList 
         Caption         =   "Add to Quote Board"
      End
      Begin VB.Menu mnuGreeks 
         Caption         =   "Greeks Calculator"
      End
      Begin VB.Menu mnuProbability 
         Caption         =   "Probabilty Calculator"
      End
      Begin VB.Menu mnuBuy 
         Caption         =   "Buy"
      End
      Begin VB.Menu mnuSell 
         Caption         =   "Sell"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChangeSymbol 
         Caption         =   "Change Symbol"
      End
      Begin VB.Menu mnuChangeFields 
         Caption         =   "Fields"
      End
      Begin VB.Menu mnuSettings 
         Caption         =   "Settings"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "Refresh"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChangeFont 
         Caption         =   "Change Font"
      End
   End
End
Attribute VB_Name = "frmOptionChain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmOptionChain.frm
'' Description: Shows the user an option chain for the selected symbol
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' ??/??/??     R Johnson   Created
'' 12/07/00     D Jarmuth   Added formatting/comments
'' 01/02/01     D Jarmuth   Made option chains work with the database and getting
''                          a download when necessary, and added Update button
'' 02/05/2010   DAJ         New functionality for new stock option symbols
'' 02/08/2010   DAJ         Added "Days" into the stable sort
'' 02/11/2010   DAJ         Fixed contract line and symbol column for future options
'' 08/05/2010   DAJ         Show full SO symbol again, separate SO base syms
'' 01/09/2013   DAJ         Use broker view for orders if it is loaded
'' 12/15/2015   DAJ         Fix for regional settings issue in ColorRows
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    lRow As Long                        ' Current row in the grid
    lSymbolID As Long                   ' Symbol ID of the underlying symbol
    lIRXSymbolID As Long                ' Symbol ID for $IRX
    strSymbol As String                 ' Symbol doing the option chain for
    eSecType As eSYM_SecType            ' Security type of underlying symbol
    tblOption As cGdTable               ' Table of option information
    alIndex As cGdArray                 ' Index into the option table
    strDefaultCols As String            ' Default column setup information
    alHeaderRows As cGdArray            ' Array of row numbers for header lines
    dUnderlying As Double               ' Current value of the underlying symbol
    strOCB As String                    ' What to use for option price
    
    bAllowOptionCalc As Boolean         ' Do we allow them to do calculations?
    
    BarsColl As cGdTree                 ' Collection of Bars for real-time
    astrLookup As cGdArray              ' Lookup table for location in grid
    
    bSymbolChanged As Boolean
    bInProgress As Boolean
    strProbCalc As String
End Type
Private m As mPrivate

' NOTE: Whenever you add or remove columns from this list, make sure to update the
' default string(s)!!!!!
Private Enum eGDCols
    eGDCol_CallIndex = 0
    eGDCol_CallSymbol
    eGDCol_CallDate
    eGDCol_CallTime
    eGDCol_CallOpen
    eGDCol_CallHigh
    eGDCol_CallLow
    eGDCol_CallLast
    eGDCol_CallBid
    eGDCol_CallAsk
    eGDCol_CallVol
    eGDCol_CallOI
    eGDCol_CallDaysToExpire
    eGDCol_CallIntrVal
    eGDCol_CallTimeVal
    eGDCol_CallImplVol
    eGDCol_CallDelta
    eGDCol_CallGamma
    eGDCol_CallVega
    eGDCol_CallTheta
    eGDCol_CallRho
    
    eGDCol_Strike
    
    eGDCol_PutIndex
    eGDCol_PutSymbol
    eGDCol_PutDate
    eGDCol_PutTime
    eGDCol_PutOpen
    eGDCol_PutHigh
    eGDCol_PutLow
    eGDCol_PutLast
    eGDCol_PutBid
    eGDCol_PutAsk
    eGDCol_PutVol
    eGDCol_PutOI
    eGDCol_PutDaysToExpire
    eGDCol_PutIntrVal
    eGDCol_PutTimeVal
    eGDCol_PutImplVol
    eGDCol_PutDelta
    eGDCol_PutGamma
    eGDCol_PutVega
    eGDCol_PutTheta
    eGDCol_PutRho
    
    eGDCol_NumCols
End Enum

Private Enum eTblCols
    etblCols_Months = 0
    etblCols_Years
    etblCols_Strikes
    etblCols_Calls
    etblCols_Symbols
    etblCols_Date
    etblCols_DaysLeft
    etblCols_Days
    etblCols_BaseSymbols
    etblCols_NumFields
End Enum

Private Const kAtTheMoneyColor = &HFFFF00
Private Const kNearTheMoneyColor = &HFFFFC0

Private Function TblCol(ByVal Col As eTblCols) As Long
    TblCol = Col
End Function

Private Function GDCol(ByVal Col As eGDCols) As Long
    GDCol = Col
End Function

Public Property Get SymbolID() As Long
    SymbolID = m.lSymbolID
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadGrid
'' Description: Loads the option chain data into the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadGrid(Optional ByVal bForceDownload As Boolean = False)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim strOptType As String            ' Option Type (i.e. Call/Put)
    Dim barsEOD As New cGdBars          ' End of day data for the symbol
    Dim Bars As New cGdBars             ' Data for the underlying symbol
    Dim dDateTime As Double             ' Date and Time of the last bar
    Dim strDate As String               ' Formatted Date of the last bar
    Dim strTime As String               ' Formatted Time of the last bar
    Dim bDistributed As Boolean         ' Did we distribute new information?
    Dim bOnQuoteBoard As Boolean        ' Is there a security on the quote board?
    Dim lRedraw As Long                 ' Current state of the grid's redraw
    Dim strSymbol As String             ' Display version of the symbol
    Dim alIndex As cGdArray             ' Index into the table
    Dim strContract As String           ' Contract we are working on
    Dim dLastStrike As Double           ' Last strike we put in the grid
    Dim strLastContract As String       ' Last contract we put in the grid
    Dim alFields As cGdArray            ' Array of positions of columns
    Dim lDaysToExpire As Long           ' Days left until expiration
    Dim dOptPrice As Double             ' Price of the current option
    Dim dTemp As Double                 ' Temporary variable to use in calculations
    Dim dDelta As Double                ' Delta derived from the calculations
    Dim dVolatility As Double           ' Calculated Implied Volatility
    Dim lExpDate As Long                ' Expiration date for the future option
    Dim lToday As Long                  ' Today's date
    
    ' Need to set up an array of column positions for our columns since the
    ' user can change the order around on us (The original column position is
    ' stored in the ColData of the column)
    Set alFields = New cGdArray
    alFields.Create eGDARRAY_Longs, fg.Cols
    For lIndex = 0 To fg.Cols - 1
        alFields(fg.ColData(lIndex)) = lIndex
    Next lIndex
    
    ' Try to get the option chain data from the database
    If bForceDownload Then
        m.tblOption.NumRecords = 0&
        m.alHeaderRows.Size = 0
        m.alIndex.Size = 0
        m.astrLookup.Size = 0
        m.BarsColl.Clear
    Else
        DM_GetOptChain m.lSymbolID, m.tblOption, barsEOD
    End If

    If g.RealTime.IsServerActive Then frmMain.SuspendNewSymbolCheck 15
        
    ' If the table is still empty, then download the chain...
    If m.tblOption.NumRecords = 0& Then
        bDistributed = GetOptionChain(m.strSymbol)
        DM_GetOptChain m.lSymbolID, m.tblOption, barsEOD
        Caption = "Option Chain for " & m.strSymbol
    End If
    
    ' Get the current interest rate by getting data from $IRX
    m.lIRXSymbolID = g.SymbolPool.SymbolIDforSymbol("$IRX")
    Set Bars = GetBars(m.lIRXSymbolID)
    DisplayIRX Bars
    
    ' Get the bars for the underlying symbol and display it's information...
    'If DM_GetBars(Bars, m.lSymbolID, , LastDailyDownload - 5) Then
    Set Bars = GetBars(m.lSymbolID)
    If Not Bars Is Nothing Then
        If frmQuotes.SymbolExists(m.strSymbol) Then bOnQuoteBoard = True
        DisplayInfo Bars
        m.dUnderlying = Bars(eBARS_Close, Bars.Size - 1)
    End If
    
    lRedraw = fg.Redraw
    
    fg.Rows = fg.FixedRows
    m.alHeaderRows.Size = 0
    
    lToday = Date
    
    With m.tblOption
        Set alIndex = New cGdArray
        alIndex.Create eGDARRAY_Longs
        
        Set m.alIndex = .CreateIndex
        .SortIndex m.alIndex, TblCol(etblCols_Symbols)
        
        ' Create the sorted index on the table by Strike, Month, Year, then Symbol...
        Set alIndex = .CreateIndex
        '.SortIndex alIndex, TblCol(etblCols_Symbols), eGdSort_Stable
        .SortIndex alIndex, TblCol(etblCols_Calls), eGdSort_Stable Or eGdSort_Descending
        .SortIndex alIndex, TblCol(etblCols_Strikes), eGdSort_Stable
        .SortIndex alIndex, TblCol(etblCols_BaseSymbols), eGdSort_Stable
        .SortIndex alIndex, TblCol(etblCols_Days), eGdSort_Stable
        .SortIndex alIndex, TblCol(etblCols_Months), eGdSort_Stable
        .SortIndex alIndex, TblCol(etblCols_Years), eGdSort_Stable
        
        dLastStrike = -1#
        strLastContract = ""
        
        fg.Redraw = flexRDNone
        
        For lIndex = 0 To .NumRecords - 1
            ' Figure out if the option is a call or a put
            If .Item(TblCol(etblCols_Calls), alIndex(lIndex)) = 1 Then
                strOptType = "Call"
            Else
                strOptType = "Put"
            End If

            ' Fix symbol (for backwards-compatibility)
            strSymbol = .Item(TblCol(etblCols_Symbols), alIndex(lIndex))
            If m.eSecType <> eSYMType_Future Then
                If (InStr(strSymbol, " ") = 0) And (Len(strSymbol) >= 3) And (Len(strSymbol) <= 5) Then
                    strSymbol = Left(strSymbol, Len(strSymbol) - 2) & " " & Right(strSymbol, 2)
                End If
            End If
            
            Set Bars = GetBars(strSymbol)
            
            If frmQuotes.SymbolExists(strSymbol) Then bOnQuoteBoard = True
            
            .Item(TblCol(etblCols_Date), alIndex(lIndex)) = barsEOD(eBARS_DateTime, alIndex(lIndex))
            strDate = DateFormat(.Item(TblCol(etblCols_Date), alIndex(lIndex)), MM_DD_YY)
            strTime = DateFormat(barsEOD(eBARS_DateTime, alIndex(lIndex)), NO_DATE, HH_MM, AMPM_LOWER, True)
            If strTime = "12:00 AM" Then strTime = ""
                        
            ' TLB 1/20/2010: migrating to new display format for the "monthly group row"
            If m.eSecType <> eSYMType_Future Then
                'lExpDate = GetDateFromRule(.Item(TblCol(etblCols_Years), alIndex(lIndex)), .Item(TblCol(etblCols_Months), alIndex(lIndex)), "3F") + 1
                'strContract = MonthDisplay(.Item(TblCol(etblCols_Months), alIndex(lIndex))) & " " & _
                        Str(.Item(TblCol(etblCols_Years), alIndex(lIndex)))
                lExpDate = JulFromLong((.Item(TblCol(etblCols_Years), alIndex(lIndex)) * 10000) + (.Item(TblCol(etblCols_Months), alIndex(lIndex)) * 100) + .Item(TblCol(etblCols_Days), alIndex(lIndex)))
                'strContract = m.strSymbol & "  " & Format(lExpDate, "mmm d, yyyy")
                strContract = .Item(TblCol(etblCols_BaseSymbols), alIndex(lIndex)) & "  " & Format(lExpDate, "mmm d, yyyy")
            Else
                strContract = ConvertToTradeSymbol(m.strSymbol, lToday) & " " & MonthName(.Item(TblCol(etblCols_Months), alIndex(lIndex))) & " " & .Item(TblCol(etblCols_Years), alIndex(lIndex))
            End If
                    
            ' If starting a new contract, put a header line out with that contract
            If strContract <> strLastContract Then
                fg.Rows = fg.Rows + 1
                fg.Cell(flexcpText, fg.Rows - 1, 0, fg.Rows - 1, fg.Cols - 1) = strContract
                fg.Cell(flexcpFontBold, fg.Rows - 1, 0, fg.Rows - 1, fg.Cols - 1) = True
                fg.MergeRow(fg.Rows - 1) = True
                strLastContract = strContract
                m.alHeaderRows.Add fg.Rows - 1
                fg.RowData(fg.Rows - 1) = m.alHeaderRows.Size - 1
            End If
            
            ' If a new strike price, then start a new line
            If .Item(TblCol(etblCols_Strikes), alIndex(lIndex)) <> dLastStrike Then
                fg.Rows = fg.Rows + 1
                'If fg.Rows - 1 > fg.BottomRow Then fg.Redraw = flexRDBuffered
                dLastStrike = .Item(TblCol(etblCols_Strikes), alIndex(lIndex))
                fg.TextMatrix(fg.Rows - 1, alFields(eGDCol_Strike)) = Str(RoundToSigDigits(.Item(TblCol(etblCols_Strikes), alIndex(lIndex)), 7))
                fg.MergeRow(fg.Rows - 1) = False
                fg.RowData(fg.Rows - 1) = -1
            Else
                If .Item(TblCol(etblCols_Calls), alIndex(lIndex)) = 1 Then
                    If strSymbol <> fg.TextMatrix(fg.Rows - 1, alFields(eGDCol_CallSymbol)) Then
                        fg.Rows = fg.Rows + 1
                        fg.TextMatrix(fg.Rows - 1, alFields(eGDCol_Strike)) = Str(RoundToSigDigits(.Item(TblCol(etblCols_Strikes), alIndex(lIndex)), 7))
                        fg.MergeRow(fg.Rows - 1) = False
                        fg.RowData(fg.Rows - 1) = -1
                    End If
                End If
            End If
            
            ' Figure out the price to use for calculations based on the user's settings...
            dOptPrice = OptionPrice(barsEOD(eBARS_Bid, alIndex(lIndex)), barsEOD(eBARS_Ask, alIndex(lIndex)), barsEOD(eBARS_Close, alIndex(lIndex)))

            If .Item(TblCol(etblCols_Calls), alIndex(lIndex)) = 1 Then
                m.astrLookup.Add strSymbol & vbTab & Str(fg.Rows - 1) & vbTab & "Call" & vbTab & Str(alIndex(lIndex))
                m.astrLookup.Sort
                
                ' 01/27/2010 DAJ: For phase 1 of the option symbol changes, if the stock option symbol is
                ' greater than 6 characters (new symbology), only display the root symbol...
                If (Len(strSymbol) <= 6) Or (m.eSecType = eSYMType_Future) Then
                    fg.TextMatrix(fg.Rows - 1, alFields(eGDCol_CallSymbol)) = strSymbol
                Else
                    fg.TextMatrix(fg.Rows - 1, alFields(eGDCol_CallSymbol)) = strSymbol ' Parse(strSymbol, " ", 1)
                End If
                
                fg.TextMatrix(fg.Rows - 1, alFields(eGDCol_CallIndex)) = Str(alIndex(lIndex))
                UpdateSymbol strSymbol, Bars
                Calculate strSymbol
            Else
                m.astrLookup.Add strSymbol & vbTab & Str(fg.Rows - 1) & vbTab & "Put" & vbTab & Str(alIndex(lIndex))
                m.astrLookup.Sort
                
                ' 01/27/2010 DAJ: For phase 1 of the option symbol changes, if the stock option symbol is
                ' greater than 6 characters (new symbology), only display the root symbol...
                If (Len(strSymbol) <= 6) Or (m.eSecType = eSYMType_Future) Then
                    fg.TextMatrix(fg.Rows - 1, alFields(eGDCol_PutSymbol)) = strSymbol
                Else
                    fg.TextMatrix(fg.Rows - 1, alFields(eGDCol_PutSymbol)) = strSymbol ' Parse(strSymbol, " ", 1)
                End If
                
                fg.TextMatrix(fg.Rows - 1, alFields(eGDCol_PutIndex)) = Str(alIndex(lIndex))
                UpdateSymbol strSymbol, Bars
                Calculate strSymbol
            End If
        Next lIndex
    End With
    
    If g.RealTime.IsServerActive Then frmMain.SuspendNewSymbolCheck

    If fg.Rows > fg.FixedRows Then
        fg.Cell(flexcpFontBold, fg.FixedRows, alFields(eGDCol_Strike), fg.Rows - 1) = True
        fg.Select fg.FixedRows, GDCol(eGDCol_Strike), fg.Rows - 1
        fg.CellBorder vbBlack, 2, -1, 2, -1, -1, -1
        fg.Select 0, 0
        ColorRows
        ExtendCustomColumn
    End If
    fg.Redraw = lRedraw
       
    If bDistributed Then
        DoEvents
        UpdateVisibleCharts
        If bOnQuoteBoard Then frmQuotes.TotalRefresh True
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptionChain.LoadGrid", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdClose_Click
'' Description: If the user clicks on the Close button, unload the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub cmdClose_Click()
On Error GoTo ErrSection:

    DockState(Me) = eHidden
    frmMain.tbToolbar.Tools("ID_Chain").State = ssUnchecked

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptionChain.cmdClose.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdFields_Click
'' Description: Allow the user to change the fields that are shown
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdFields_Click()
On Error GoTo ErrSection:

    ChangeFields

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptionChain.cmdFields.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdSettings_Click
'' Description: Allow the user to customize some settings
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdSettings_Click()
On Error GoTo ErrSection:

    ChangeSettings

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptionChain.cmdSettings.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdUpdate_Click
'' Description: If the user clicks on the Update button, download the data
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdUpdate_Click()
On Error GoTo ErrSection:

    ' Need to do this for "Default" buttons, since hitting 'Enter'
    ' from another control does not trigger it's 'LostFocus' event.
    MoveFocus cmdUpdate

    If ProcessIsBusy Then Exit Sub

    LoadGrid True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptionChain.cmdUpdate.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fg_AfterRowColChange
'' Description: Enable/Disable controls as the user changes cells in the grid
'' Inputs:      Old Row and Column, New Row and Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fg_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:
    
    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptionChain.fg.AfterRowColChange", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fg_DblClick
'' Description:
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fg_DblClick()
On Error GoTo ErrSection:

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptionChain.fg.DblClick", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fg_KeyDown
'' Description: Pass any keystrokes on to the form
'' Inputs:      Code of the key pressed, Shift/Ctrl/Alt status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fg_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If fgKeyDown(KeyCode, Shift) Then Exit Sub

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptionChain.fg.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fg_MouseDown
'' Description: If the user clicks the right mouse button on the grid, bring
''              up the pop-up menu to allow them to add the symbol to the
''              quote list
'' Inputs:      Mouse Button Pressed, Shift/Ctrl/Alt status, Mouse click location
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fg_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    Dim strSymbol As String             ' Symbol of highlighted section
    Dim lMouseRow As Long               ' Current mouse row of the grid
    Dim lMouseCol As Long               ' Current mouse column of the grid

    With fg
        lMouseRow = .MouseRow
        lMouseCol = .MouseCol
        
        If lMouseRow > .FixedRows - 1 And lMouseRow < .Rows Then
            .Row = lMouseRow
            .RowSel = .Row
                        
            .Col = lMouseCol
            If lMouseCol <= GDCol(eGDCol_Strike) Then
                .Col = GDCol(eGDCol_CallSymbol)
            Else
                .Col = GDCol(eGDCol_PutSymbol)
            End If
            .ColSel = .Col
            
            If Button = vbRightButton Then
                m.lRow = lMouseRow
                'strSymbol = .TextMatrix(.Row, .Col)
                strSymbol = SymbolForRow(.Row, (lMouseCol <= GDCol(eGDCol_Strike)))
                
                If (.RowData(.Row) = -1) And (Len(strSymbol) > 0) Then
                    mnuQuoteList.Caption = "Add " & strSymbol & " to Quote Board"
                    mnuQuoteList.Visible = True
                    mnuGreeks.Caption = "Greeks Calculator for " & strSymbol
                    mnuGreeks.Visible = HasModule("GOPC")
                    mnuProbability.Caption = "Probability Calculator for " & strSymbol
                    mnuProbability.Visible = HasModule("GOPC") And FileExist(m.strProbCalc)
                    mnuBuy.Caption = "Buy " & strSymbol
                    mnuBuy.Visible = True
                    mnuSell.Caption = "Sell " & strSymbol
                    mnuSell.Visible = True
                    mnuSep1.Visible = True
                Else
                    mnuQuoteList.Caption = "Add to Quote Board"
                    mnuQuoteList.Visible = False
                    mnuGreeks.Caption = "Greeks Calculator"
                    mnuGreeks.Visible = HasModule("GOPC")
                    mnuProbability.Caption = "Probability Calculator"
                    mnuProbability.Visible = HasModule("GOPC") And FileExist(m.strProbCalc)
                    mnuBuy.Visible = False
                    mnuSell.Visible = False
                    mnuSep1.Visible = False
                End If
                PopupMenu mnuAdd
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptionChain.fg.MouseDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fg_MouseMove
'' Description: If the mouse is moved over a symbol, show the description as
''              a tool tip
'' Inputs:      Mouse Button Pressed, Shift/Ctrl/Alt status, Location of Mouse
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fg_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    Dim lMouseCol As Long
    Dim lMouseRow As Long
    
    With fg
        lMouseCol = .MouseCol
        lMouseRow = .MouseRow
        
        If lMouseCol = GDCol(eGDCol_CallSymbol) Or lMouseCol = GDCol(eGDCol_PutSymbol) And .MergeRow(lMouseRow) = False Then
            '.ToolTipText = m.BarsColl(.TextMatrix(lMouseRow, lMouseCol)).Prop(eBARS_Desc)
            .ToolTipText = m.BarsColl(SymbolForRow(lMouseRow, (lMouseCol <= GDCol(eGDCol_Strike)))).Prop(eBARS_Desc)
        Else
            .ToolTipText = ""
        End If
    End With

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Activate
'' Description: When the form is activated, reset the toolbar and the window
''              list
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Activate()
On Error GoTo ErrSection:

    ToolbarSync Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptionChain.Form.Activate", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Deactivate
'' Description: Upon deactivation, set this form as the previous form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Deactivate()
On Error GoTo ErrSection:

    SetPrevActiveForm Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptionChain.Form.Deactivate", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_KeyDown
'' Description: Show the help if the user pressed F1
'' Inputs:      Code of the Key Pressed, Shift/Ctrl/Alt status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If KeyCode = vbKeyF1 Then
        KeyCode = 0
        g.Help.ShowF1Help Me
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptionChain.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: When the form is loaded, set the size and position of the form
''              and set up the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Dim w&
    
    g.Styler.StyleForm Me
    
    m.strProbCalc = AddSlash(App.Path) & "Options\ProbCalc.EXE"
    
    mnuAdd.Visible = False
    
    w = 8700
    If frmMain.ScaleWidth < w Then w = frmMain.ScaleWidth
    If w <= 2000 Then w = 2000
    Me.Width = w
    CenterTheForm Me
    Me.Icon = Picture16(ToolbarIcon("ID_Chain"), , True)
    
    With tbToolbar
        .Tools("ID_Close").Picture = Picture16(ToolbarIcon("kCancel"))
        .Tools("ID_Fields").Picture = Picture16(ToolbarIcon("ID_SymbolGrid"))
        .Tools("ID_Refresh").Picture = Picture16(ToolbarIcon("ID_Download"))
        .Tools("ID_Settings").Picture = Picture16(ToolbarIcon("ID_Settings"))
        .Tools("ID_Symbol").Picture = Picture16(ToolbarIcon("ID_Symbol"))
        .Tools("ID_Probability").Picture = Picture16(ToolbarIcon("ID_Criteria"))
        .Tools("ID_Greeks").Picture = Picture16(ToolbarIcon("ID_Criteria"))
    End With
    
    Set m.tblOption = New cGdTable
    With m.tblOption
        .CreateField eGDARRAY_Longs, TblCol(etblCols_Months), "Months"
        .CreateField eGDARRAY_Longs, TblCol(etblCols_Years), "Years"
        .CreateField eGDARRAY_Doubles, TblCol(etblCols_Strikes), "Strikes"
        .CreateField eGDARRAY_Longs, TblCol(etblCols_Calls), "Calls"
        .CreateField eGDARRAY_Strings, TblCol(etblCols_Symbols), "Symbols"
        .CreateField eGDARRAY_Doubles, TblCol(etblCols_Date), "Dates"
        .CreateField eGDARRAY_Longs, TblCol(etblCols_DaysLeft), "DaysLeft", kNullData
        .CreateField eGDARRAY_Longs, TblCol(etblCols_Days), "Days", kNullData
        .CreateField eGDARRAY_Strings, TblCol(etblCols_BaseSymbols), "BaseSyms"
    End With
    
    fg.Clear
    
    Set m.BarsColl = New cGdTree
    Set m.astrLookup = New cGdArray
    m.astrLookup.Create eGDARRAY_Strings
    Set m.alIndex = New cGdArray
    m.alIndex.Create eGDARRAY_Longs
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptionChain.Form.Load", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Month
'' Description: Return the string representation of the month number passed in
'' Inputs:      Month number
'' Returns:     String abbreviation of the month
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function MonthDisplay(iMonthNum As Integer) As String
On Error GoTo ErrSection:

    MonthDisplay = Format(DateSerial(2000, iMonthNum, 1), "mmmm")
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmOptionChain.MonthDisplay", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: If the user clicks on the X, uncheck the option chain toolbar
'' Inputs:      Whether or not to Cancel the Unload, Mode of the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode = 0 Then
        frmMain.tbToolbar.Tools("ID_Chain").State = ssUnchecked
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptionChain.Form.QueryUnload", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Resize
'' Description: When the user resizes the form, resize the controls accordingly
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Resize()
On Error Resume Next
    
    If LimitFormSize(Me, fraButtons.Width, fraButtons.Height * 2) Then Exit Sub

    With lblSymbolInfo
        .Move 0, 120, ScaleWidth
    End With
    
    With fgUnder
        .Move 0, lblSymbolInfo.Top + lblSymbolInfo.Height, ScaleWidth
    End With

    With fraButtons
        .Move (ScaleWidth - .Width) / 2, ScaleHeight - .Height - 120
        .Visible = False
    End With
    
    With fg
        .Move .Left, fgUnder.Top + fgUnder.Height, _
                Me.ScaleWidth - .Left * 2, _
                ScaleHeight - fgUnder.Height - lblSymbolInfo.Height - 120
        .Refresh
        ExtendCustomColumn
    End With
    
    AutoSizeChart
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: When the form is unloaded, update the windows list
'' Inputs:      Whether or not to cancel the unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop

    tmrRealtime.Enabled = False
    
    SetIniFileProperty "OptionChain", FontToString(fg.Font), "Fonts", g.strIniFile
    ToolbarSync Me, False
    frmMain.DockPro.RemoveForm Me.Name
    
    Set m.tblOption = Nothing
    
    'For lIndex = 1 To m.BarsColl.Count
    '    g.RealTime.RemoveTickBuffer m.BarsColl(lIndex)
    'Next lIndex
    Set m.BarsColl = Nothing
    Set m.astrLookup = Nothing
    Set m.alIndex = Nothing
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptionChain.Form.Unload", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DisplayPrice
'' Description: Convert a price into a string to display
'' Inputs:      Price to convert
'' Returns:     String to display
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DisplayPrice(ByVal dPrice As Double, Optional ByVal strIfNull As String = "", Optional ByVal Bars As cGdBars = Nothing) As String
On Error GoTo ErrSection:

    If dPrice = kNullData Then
        DisplayPrice = strIfNull
    ElseIf Not Bars Is Nothing Then
        DisplayPrice = Bars.PriceDisplay(dPrice, True)
    Else
        DisplayPrice = Str(Round(dPrice, 5))
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmOptionChain.DisplayPrice", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DisplayVol
'' Description: Convert a volume into a string to display
'' Inputs:      Volume to convert
'' Returns:     String to display
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DisplayVol(ByVal dVol As Double, Optional ByVal strIfNull As String = "") As String
On Error GoTo ErrSection:

    If dVol = kNullData Then
        DisplayVol = strIfNull
    Else
        DisplayVol = Format(dVol, "#,##0")
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmOptionChain.DisplayVol", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuBuy_Click
'' Description: Allow the user to enter a buy order for the selected symbol
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuBuy_Click()
On Error GoTo ErrSection:
    
    If FormIsLoaded("frmBrokerView") Then
        frmBrokerView.CreateOrderForLot "", SymbolForRow(fg.Row, (fg.Col <= GDCol(eGDCol_Strike))), "Buy"
    Else
        'CreateOrder fg.TextMatrix(fg.Row, fg.Col), , 1, , , "Option Chain"
        CreateOrder SymbolForRow(fg.Row, (fg.Col <= GDCol(eGDCol_Strike))), , 1, , , "Option Chain"
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptionChain.mnuBuy.Click", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuGreeks_Click
'' Description: If the user chooses the Option Calculator menu option, bring
''              the calculator up on the symbol that they are on
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuGreeks_Click()
On Error GoTo ErrSection:

    RunOptionsCalc

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptionChain.mnuGreeks.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuChangeFields_Click
'' Description: Allow the user to change the fields that are shown
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuChangeFields_Click()
On Error GoTo ErrSection:

    ChangeFields

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptionChain.mnuChangeFields.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuChangeFont_Click
'' Description: Change the font of the quotes grid if the user chooses to
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuChangeFont_Click()
On Error GoTo ErrSection:

    ChangeGridFont fg, True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptionChain.mnuChangeFont.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuChangeSymbol_Click
'' Description: Allow the user to get a chain for a different symbol
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuChangeSymbol_Click()
On Error GoTo ErrSection:

    ChangeSymbol

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptionChain.mnuChangeSymbol.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuProbability_Click
'' Description: Allow the user to run the option probability calculator
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuProbability_Click()
On Error GoTo ErrSection:

    RunProbabilityCalc

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptionChain.mnuProbability.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuQuoteList_Click
'' Description: If the user clicks on the Add to Quotes menu option, add the
''              symbol in the current grid row to the quote list
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuQuoteList_Click()
On Error GoTo ErrSection:

    Dim lRow As Long                    ' Row in the grid

    If m.eSecType = eSYMType_Future And (Not DoFutOpts) Then
        InfBox "Future options will be supported on the quote board in a future version.", "!", , "Not yet supported"
        Exit Sub
    End If

    With fg
        lRow = m.lRow
        
        If lRow >= .FixedRows And lRow < .Rows Then
            'frmQuotes.AddSymbol -1&, "Daily", .TextMatrix(lRow, .Col)
            frmQuotes.AddSymbol -1&, "Daily", SymbolForRow(lRow, (.Col <= GDCol(eGDCol_Strike)))
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptionChain.mnuQuoteList.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenerateReport
'' Description: Get the grid ready for printing
'' Inputs:      Arguments passed in from print form
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GenerateReport(ByVal vArgs As Variant)
On Error GoTo ErrSection:

    Dim lRow As Long
    Dim lCol As Long
    Dim strText As String

    With frmPrintPreview.vp
        .StartDoc
        DoPrintHeader
        
        .Font.Size = 12
        .Text = lblSymbolInfo.Caption & vbLf
        .Font.Size = 10
        If frmPrintPreview.GoingToFile Then
            frmPrintPreview.GridToTable fgUnder
        Else
            fgUnder.Redraw = flexRDNone
            fgUnder.BackColorFixed = vbWindowBackground
            .RenderControl = fgUnder.hWnd
            fgUnder.BackColorFixed = Me.BackColor
            fgUnder.Redraw = flexRDBuffered
        End If
        
        If frmPrintPreview.GoingToFile Then
            frmPrintPreview.GridToTable fg
        Else
            .RenderControl = fg.hWnd
        End If
        
        .EndDoc
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptionChain.GenerateReport", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PrintMe
'' Description: Bring up the print preview to print the grid
'' Inputs:      None
'' Returns:     True on success, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function PrintMe() As Boolean
On Error GoTo ErrSection

    PrintMe = frmPrintPreview.ShowMe("CNV OptionChain", frmOptionChain, , , , 0.75, 0.75)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmOptionChain.PrintMe", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitGrid
'' Description: Initialize the Option Chain Grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitGrid()
On Error GoTo ErrSection:

    Dim lRedraw As Long
    Dim strFont As String               ' Font from the INI file
    Dim strCols As String               ' Column show/hidden status from INI
    Dim lIndex As Long

    ' Get the current font from the INI file
    strFont = GetIniFileProperty("OptionChain", "", "Fonts", g.strIniFile)
    
    With fg
        lRedraw = .Redraw
        .Redraw = flexRDNone
        .Clear
        .BackColorAlternate = .BackColor '= ALT_GRID_ROW_COLOR
        .AllowUserResizing = flexResizeColumns
        .ExplorerBar = flexExNone ' = flexExSortShow 'AndMove
        .ScrollTrack = True
        .SelectionMode = flexSelectionListBox
        .SheetBorder = RGB(128, 128, 128)
        .ExtendLastCol = False 'True
        If strFont <> "" Then FontFromString .Font, strFont
        .Editable = flexEDNone
        .MergeCells = flexMergeFree
        
        .Cols = GDCol(eGDCol_NumCols)
        .Rows = 2
        
        .FixedCols = 0
        .FixedRows = 2
        '.FrozenCols = 1
        
        .Cell(flexcpText, 0, 0, 0, GDCol(eGDCol_Strike) - 1) = "CALLS"
        .Cell(flexcpText, 0, GDCol(eGDCol_Strike) + 1, 0, .Cols - 1) = "PUTS"
        .Cell(flexcpText, 0, GDCol(eGDCol_Strike), 1, GDCol(eGDCol_Strike)) = "STRIKE" '"Strike"
        .MergeCol(GDCol(eGDCol_Strike)) = True
        .MergeRow(0) = True
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = True
        
        .TextMatrix(1, GDCol(eGDCol_CallSymbol)) = "Symbol"
        .TextMatrix(1, GDCol(eGDCol_CallDate)) = "Date"
        .TextMatrix(1, GDCol(eGDCol_CallTime)) = "Time"
        .TextMatrix(1, GDCol(eGDCol_CallOpen)) = "Open"
        .TextMatrix(1, GDCol(eGDCol_CallHigh)) = "High"
        .TextMatrix(1, GDCol(eGDCol_CallLow)) = "Low"
        .TextMatrix(1, GDCol(eGDCol_CallLast)) = "Last"
        .TextMatrix(1, GDCol(eGDCol_CallBid)) = "Bid"
        .TextMatrix(1, GDCol(eGDCol_CallAsk)) = "Ask"
        .TextMatrix(1, GDCol(eGDCol_CallVol)) = "Vol"
        .TextMatrix(1, GDCol(eGDCol_CallOI)) = "OI"
        .TextMatrix(1, GDCol(eGDCol_CallDaysToExpire)) = "Days Left"
        .TextMatrix(1, GDCol(eGDCol_CallIntrVal)) = "Intr Val"
        .TextMatrix(1, GDCol(eGDCol_CallTimeVal)) = "Time Val"
        .TextMatrix(1, GDCol(eGDCol_CallImplVol)) = "Impl Vol"
        .TextMatrix(1, GDCol(eGDCol_CallDelta)) = "Delta"
        .TextMatrix(1, GDCol(eGDCol_CallGamma)) = "Gamma"
        .TextMatrix(1, GDCol(eGDCol_CallVega)) = "Vega"
        .TextMatrix(1, GDCol(eGDCol_CallTheta)) = "Theta"
        .TextMatrix(1, GDCol(eGDCol_CallRho)) = "Rho"
                
        .TextMatrix(1, GDCol(eGDCol_PutSymbol)) = "Symbol"
        .TextMatrix(1, GDCol(eGDCol_PutDate)) = "Date"
        .TextMatrix(1, GDCol(eGDCol_PutTime)) = "Time"
        .TextMatrix(1, GDCol(eGDCol_PutOpen)) = "Open"
        .TextMatrix(1, GDCol(eGDCol_PutHigh)) = "High"
        .TextMatrix(1, GDCol(eGDCol_PutLow)) = "Low"
        .TextMatrix(1, GDCol(eGDCol_PutLast)) = "Last"
        .TextMatrix(1, GDCol(eGDCol_PutBid)) = "Bid"
        .TextMatrix(1, GDCol(eGDCol_PutAsk)) = "Ask"
        .TextMatrix(1, GDCol(eGDCol_PutVol)) = "Vol"
        .TextMatrix(1, GDCol(eGDCol_PutOI)) = "OI"
        .TextMatrix(1, GDCol(eGDCol_PutDaysToExpire)) = "Days Left"
        .TextMatrix(1, GDCol(eGDCol_PutIntrVal)) = "Intr Val"
        .TextMatrix(1, GDCol(eGDCol_PutTimeVal)) = "Time Val"
        .TextMatrix(1, GDCol(eGDCol_PutImplVol)) = "Impl Vol"
        .TextMatrix(1, GDCol(eGDCol_PutDelta)) = "Delta"
        .TextMatrix(1, GDCol(eGDCol_PutGamma)) = "Gamma"
        .TextMatrix(1, GDCol(eGDCol_PutVega)) = "Vega"
        .TextMatrix(1, GDCol(eGDCol_PutTheta)) = "Theta"
        .TextMatrix(1, GDCol(eGDCol_PutRho)) = "Rho"
                
        .ColAlignment(GDCol(eGDCol_Strike)) = flexAlignCenterTop
        .ColAlignment(GDCol(eGDCol_CallDate)) = flexAlignCenterTop
        .ColAlignment(GDCol(eGDCol_CallTime)) = flexAlignCenterTop
        .ColAlignment(GDCol(eGDCol_PutDate)) = flexAlignCenterTop
        .ColAlignment(GDCol(eGDCol_PutTime)) = flexAlignCenterTop
        
        .ColHidden(GDCol(eGDCol_CallIndex)) = True
        .ColHidden(GDCol(eGDCol_PutIndex)) = True
        
        For lIndex = 0 To .Cols - 1
            .ColData(lIndex) = lIndex
        Next lIndex
        
        .Cell(flexcpAlignment, 0, 0, 1, .Cols - 1) = flexAlignCenterTop
        .Cell(flexcpAlignment, 0, GDCol(eGDCol_Strike), 1) = flexAlignCenterCenter
        
        FieldOrder = GetIniFileProperty("DisplayFields" & m.eSecType, m.strDefaultCols, "OptionChain", g.strIniFile)
        'ExtendCustomColumn
        .AutoSize 0, .Cols - 1
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptionChain.InitGrid", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RunOptionsCalc
'' Description: Run the options calculator
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RunOptionsCalc()
On Error GoTo ErrSection:

    Dim strStockSym As String           ' Symbol of the underlying stock
    Dim strOptSym As String             ' Symbol of the option
    Dim dStockPrice As Double           ' Price of the underlying stock
    Dim dOptAsk As Double               ' Ask price of the option
    Dim dOptBid As Double               ' Bid price of the option
    Dim dOptLast As Double              ' Last price of the option
    Dim bIsPut As Byte                  ' Put/Call
    Dim nIndex As Long                  ' Index into a for loop
    Dim alFields As cGdArray            ' Array of positions of columns
    
    ' Need to set up an array of column positions for our columns since the
    ' user can change the order around on us (The original column position is
    ' stored in the ColData of the column)
    Set alFields = New cGdArray
    alFields.Create eGDARRAY_Longs, fg.Cols
    For nIndex = 0 To fg.Cols - 1
        alFields(fg.ColData(nIndex)) = nIndex
    Next nIndex

    ' If the user clicked on an invalid row or they are not allowed to use
    ' the calculator, exit the sub
    If fg.Row < fg.FixedRows Then Exit Sub
    'If Not HasModule("GOPC") Then Exit Sub
    If fg.RowData(fg.Row) <> -1 Then Exit Sub
    
    If m.eSecType = eSYMType_Future And (Not DoFutOpts) Then
        InfBox "Cannot perform calculations for future options.", "e", , "Options Calculator"
        Exit Sub
    End If
        
    ' Get the appropriate information from the grid
    strStockSym = m.strSymbol
    dStockPrice = m.dUnderlying
    If fg.Col < GDCol(eGDCol_Strike) Then
        'strOptSym = fg.TextMatrix(fg.Row, alFields(eGDCol_CallSymbol))
        strOptSym = SymbolForRow(fg.Row, True)
        dOptLast = ValOfText(fg.TextMatrix(fg.Row, alFields(eGDCol_CallLast)))
        dOptBid = ValOfText(fg.TextMatrix(fg.Row, alFields(eGDCol_CallBid)))
        dOptAsk = ValOfText(fg.TextMatrix(fg.Row, alFields(eGDCol_CallAsk)))
        bIsPut = 0
        nIndex = CLng(fg.TextMatrix(fg.Row, alFields(eGDCol_CallIndex)))
    Else
        'strOptSym = fg.TextMatrix(fg.Row, alFields(eGDCol_PutSymbol))
        strOptSym = SymbolForRow(fg.Row, False)
        dOptLast = ValOfText(fg.TextMatrix(fg.Row, alFields(eGDCol_PutLast)))
        dOptBid = ValOfText(fg.TextMatrix(fg.Row, alFields(eGDCol_PutBid)))
        dOptAsk = ValOfText(fg.TextMatrix(fg.Row, alFields(eGDCol_PutAsk)))
        bIsPut = 1
        nIndex = CLng(fg.TextMatrix(fg.Row, alFields(eGDCol_PutIndex)))
    End If
    
    With m.tblOption
        frmOptionCalc.ShowMe strStockSym, strOptSym, .Item(TblCol(etblCols_Date), nIndex), _
            .Item(TblCol(etblCols_Months), nIndex), .Item(TblCol(etblCols_Years), nIndex), _
            dStockPrice, .Item(TblCol(etblCols_Strikes), nIndex), dOptAsk, dOptBid, _
            dOptLast, bIsPut
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptionChain.RunOptionsCalc", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Set up and show the form
'' Inputs:      Symbol ID of underlying, Whether to Force a Download
'' Returns:     True if Options Found, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(ByVal lSymbolID&, Optional ByVal bForceDownload As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Set m.alHeaderRows = New cGdArray
    m.alHeaderRows.Create eGDARRAY_Longs
    
    ' Set the caption of the form
    m.strOCB = GetIniFileProperty("OCB", "Avg", "OptionChain", g.strIniFile)
    m.lSymbolID = lSymbolID
    m.strSymbol = g.SymbolPool.SymbolForID(lSymbolID)
    m.eSecType = g.SymbolPool.SecType(g.SymbolPool.PoolRecForSymbolID(lSymbolID))
    Caption = "Option Chain for " & m.strSymbol & "  (data from previous update)"
    
    Select Case m.eSecType
        Case eSYMType_Future
            m.strDefaultCols = "0;-1;,1;-1;Symbol,2;-1;Date,3;-1;Time,4;-1;Open,5;-1;High,6;-1;Low,7;0;Last,8;0;Bid,9;0;Ask,10;0;Vol,11;0;OI,12;-1;Days Left,13;-1;Intr Val,14;-1;Time Val,15;-1;Impl Vol,16;-1;Delta,17;-1;Gamma,18;-1;Vega,19;-1;Theta,20;-1;Rho,21;0;Strike,22;-1;,23;-1;Symbol,24;-1;Date,25;-1;Time,26;-1;Open,27;-1;High,28;-1;Low,29;0;Last,30;0;Bid,31;0;Ask,32;0;Vol,33;0;OI,34;-1;Days Left,35;-1;Intr Val,36;-1;Time Val,37;-1;Impl Vol,38;-1;Delta,39;-1;Gamma,40;-1;Vega,41;-1;Theta,42;-1;Rho"
        Case eSYMType_Stock
            m.strDefaultCols = "0;-1;,1;0;Symbol,2;-1;Date,3;-1;Time,4;-1;Open,5;-1;High,6;-1;Low,7;0;Last,8;0;Bid,9;0;Ask,10;0;Vol,11;0;OI,12;-1;Days Left,13;-1;Intr Val,14;-1;Time Val,15;-1;Impl Vol,16;-1;Delta,17;-1;Gamma,18;-1;Vega,19;-1;Theta,20;-1;Rho,21;0;Strike,22;-1;,23;0;Symbol,24;-1;Date,25;-1;Time,26;-1;Open,27;-1;High,28;-1;Low,29;0;Last,30;0;Bid,31;0;Ask,32;0;Vol,33;0;OI,34;-1;Days Left,35;-1;Intr Val,36;-1;Time Val,37;-1;Impl Vol,38;-1;Delta,39;-1;Gamma,40;-1;Vega,41;-1;Theta,42;-1;Rho"
        Case eSYMType_Index
            m.strDefaultCols = "0;-1;,1;0;Symbol,2;-1;Date,3;-1;Time,4;-1;Open,5;-1;High,6;-1;Low,7;0;Last,8;0;Bid,9;0;Ask,10;0;Vol,11;0;OI,12;-1;Days Left,13;-1;Intr Val,14;-1;Time Val,15;-1;Impl Vol,16;-1;Delta,17;-1;Gamma,18;-1;Vega,19;-1;Theta,20;-1;Rho,21;0;Strike,22;-1;,23;0;Symbol,24;-1;Date,25;-1;Time,26;-1;Open,27;-1;High,28;-1;Low,29;0;Last,30;0;Bid,31;0;Ask,32;0;Vol,33;0;OI,34;-1;Days Left,35;-1;Intr Val,36;-1;Time Val,37;-1;Impl Vol,38;-1;Delta,39;-1;Gamma,40;-1;Vega,41;-1;Theta,42;-1;Rho"
    End Select

    ' Make sure that everything is cleared out...
    m.tblOption.NumRecords = 0&
    m.alHeaderRows.Size = 0
    m.alIndex.Size = 0
    m.astrLookup.Size = 0
    m.BarsColl.Clear
    
    InitUnderGrid
    
    ' Initalize and load the grid
    fg.Redraw = flexRDNone
    InitGrid
    LoadGrid bForceDownload
    fg.Redraw = flexRDBuffered
    
    EnableControls
    
    ' Show the form
    If DockState(Me) = eHidden Then DockState(Me) = eShowAsPrevious
    
    ShowMe = (fg.Rows > fg.FixedRows)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmOptionChain.ShowMe", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DisplayInfo
'' Description: Display the underlying symbol information
'' Inputs:      Bars of Data
'' Returns:     True if Close changed, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DisplayInfo(ByVal Bars As cGdBars) As Boolean
On Error GoTo ErrSection:

    Dim dDateTime As Double             ' Date and Time of the last bar of data
    Dim strDesc As String               ' Description of the security

    DisplayInfo = False

    ' Fill in the Symbol and Description Label...
    strDesc = g.SymbolPool.Desc(g.SymbolPool.PoolRecForSymbolID(m.lSymbolID))
    lblSymbolInfo = m.strSymbol & ": " & strDesc
    
    ' Fill in the Data Label...
    With fgUnder
        .Redraw = flexRDNone
        
        If Bars.Size > 0 Then
            dDateTime = Bars.Item(eBARS_DateTime, Bars.Size - 1)
            dDateTime = dDateTime + (Bars.Prop(eBARS_LastTickTime) / 1440)
            If g.bShowInLocalTimeZone Then
                dDateTime = ConvertTimeZone(dDateTime, Bars.Prop(eBARS_ExchangeTimeZoneInf), "")
            End If
            
            ChangeCell fgUnder, 0, 1, DateFormat(dDateTime, MM_DD_YY, HH_MM_SS, AMPM_LOWER)
            DisplayInfo = ChangeCell(fgUnder, 0, 3, DisplayPrice(Bars(eBARS_Close, Bars.Size - 1), "N/A", Bars))
            ChangeCell fgUnder, 0, 5, DisplayPrice(Bars(eBARS_Bid, Bars.Size - 1), "N/A", Bars)
            ChangeCell fgUnder, 0, 7, DisplayPrice(Bars(eBARS_Ask, Bars.Size - 1), "N/A", Bars)
            
            If m.eSecType = eSYMType_Future Then
                ChangeCell fgUnder, 0, 9, DisplayVol(Bars(eBARS_ContVol, Bars.Size - 1), "N/A")
            Else
                ChangeCell fgUnder, 0, 9, DisplayVol(Bars(eBARS_Vol, Bars.Size - 1), "N/A")
            End If
        Else
            ChangeCell fgUnder, 0, 1, ""
            DisplayInfo = ChangeCell(fgUnder, 0, 3, "")
            ChangeCell fgUnder, 0, 5, ""
            ChangeCell fgUnder, 0, 7, ""
            ChangeCell fgUnder, 0, 9, ""
        End If
        
        .AutoSize 0, .Cols - 1, False
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmOptionChain.DisplayInfo", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ExtendCustomColumn
'' Description: Adjust all column widths to accomodate the custom "extend column"
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ExtendCustomColumn()
On Error GoTo ErrSection:

    Dim lTotal As Long                  ' New width of the extended column
    Dim lIndex As Long                  ' Index into a for loop
    Dim lMinWidth As Long
        
    If fg.Cols < GDCol(eGDCol_NumCols) Then Exit Sub
        
    With fg
        For lIndex = 0 To m.alHeaderRows.Size - 1
            .RowHidden(lIndex) = True
        Next lIndex
        .AutoSize 0, fg.Cols - 1, False, 75
        For lIndex = 0 To m.alHeaderRows.Size - 1
            .RowHidden(lIndex) = False
        Next lIndex
        
        lMinWidth = .ColWidth(GDCol(eGDCol_Strike))
        
        .ColHidden(GDCol(eGDCol_Strike)) = True
        .Redraw = flexRDBuffered '(so .ClientWidth will be correct)
        .Redraw = flexRDNone
        lTotal = 0 * Screen.TwipsPerPixelX
        For lIndex = 0 To .Cols - 1
            If Not .ColHidden(lIndex) Then
                lTotal = lTotal + .ColWidth(lIndex)
            End If
        Next
        lTotal = .ClientWidth - lTotal
        If lTotal > 0 Then
            If lTotal < lMinWidth Then
                .ColWidth(GDCol(eGDCol_Strike)) = lMinWidth
            Else
                .ColWidth(GDCol(eGDCol_Strike)) = lTotal
            End If
        End If
        
        .ColHidden(GDCol(eGDCol_Strike)) = False
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptionChain.ExtendCustomColumn", eGDRaiseError_Raise
        
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ColorRows
'' Description: Color the rows appropriately
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ColorRows(Optional ByVal bForce As Boolean = True)
On Error GoTo ErrSection:

    Dim lRedraw As Long                 ' Current state of the grid's redraw
    Dim lRow As Long                    ' Index into a for loop
    Dim dStrike As Double               ' Strike Price of the current row
    Dim dNextStrike As Double           ' Strike Price of the next row
    Dim lRowData As Long                ' Row Data of the current row
    Dim lLastHeader As Long             ' Last Header row encountered
    Dim lNextHeader As Long             ' Next Header row
    Dim lStart As Long                  ' Starting value for a for loop
    Dim lEnd As Long                    ' Ending value for a for loop
    Dim lTempRow As Long                ' Temporary row
    Dim dTempStrike As Double           ' Temporary strike
    Dim dUnderlying As Double           ' Value of the underlying security
    
    Static sdPrevStrike As Double       ' Strike price less than or equal to underlying
    Static sdNextStrike As Double       ' Strike price greater than or equal to underlying

    dUnderlying = m.BarsColl(Str(m.lSymbolID)).Item(eBARS_Close, m.BarsColl(Str(m.lSymbolID)).Size - 1)
    If Not bForce Then
        If (dUnderlying > sdPrevStrike And dUnderlying < sdNextStrike) Or (dUnderlying = sdPrevStrike And dUnderlying = sdNextStrike) Then
            Exit Sub
        End If
    End If
    
    With fg
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        For lRow = .FixedRows To .Rows - 1
            If .RowData(lRow) >= 0 Then
                lRowData = .RowData(lRow)
                lLastHeader = lRow
                If .RowData(lRow) + 1 < m.alHeaderRows.Size Then
                    lNextHeader = m.alHeaderRows(.RowData(lRow) + 1)
                Else
                    lNextHeader = .Rows
                End If
                If lRow + 1 < .Rows Then
                    dStrike = Val(.TextMatrix(lRow + 1, GDCol(eGDCol_Strike)))
                Else
                    dStrike = 0#
                End If
            Else
                dStrike = Val(.TextMatrix(lRow, GDCol(eGDCol_Strike)))
            End If
            
            If lRow + 1 < .Rows Then
                dNextStrike = Val(.TextMatrix(lRow + 1, GDCol(eGDCol_Strike)))
            Else
                dNextStrike = 0#
            End If
            
            If .RowData(lRow) <> -1 Then ' lRowData Mod 2 <> 0 Then
                fg.Cell(flexcpBackColor, lRow, 0, lRow, fg.Cols - 1) = ALT_GRID_ROW_COLOR
            Else
                fg.Cell(flexcpBackColor, lRow, 0, lRow, fg.Cols - 1) = &H80000005
            End If
            
            If m.eSecType <> eSYMType_Future Or DoFutOpts Then
                If dStrike = dUnderlying Then
                    fg.Cell(flexcpBackColor, lRow, 0, lRow, fg.Cols - 1) = kAtTheMoneyColor
                    sdPrevStrike = dStrike
                    sdNextStrike = dStrike
                ElseIf dStrike < dUnderlying And dNextStrike > dUnderlying Then
                    lTempRow = lRow
                    sdPrevStrike = dStrike
                    sdNextStrike = dNextStrike
                    
                    fg.Cell(flexcpBackColor, lRow, 0, lRow, fg.Cols - 1) = kAtTheMoneyColor
                    Do While Val(.TextMatrix(lTempRow - 1, GDCol(eGDCol_Strike))) = dStrike
                        fg.Cell(flexcpBackColor, lTempRow - 1, 0, lTempRow - 1, fg.Cols - 1) = kAtTheMoneyColor
                        lTempRow = lTempRow - 1
                    Loop

                    For lStart = 1 To 2
                        lTempRow = lTempRow - 1
                        If lTempRow <= lLastHeader Then Exit For
                        fg.Cell(flexcpBackColor, lTempRow, 0, lTempRow, fg.Cols - 1) = kNearTheMoneyColor
                        dTempStrike = Val(.TextMatrix(lTempRow, GDCol(eGDCol_Strike)))
                        Do While Val(.TextMatrix(lTempRow - 1, GDCol(eGDCol_Strike))) = dTempStrike
                            fg.Cell(flexcpBackColor, lTempRow - 1, 0, lTempRow - 1, fg.Cols - 1) = kNearTheMoneyColor
                            lTempRow = lTempRow - 1
                        Loop
                    Next lStart

                    lTempRow = lRow + 1
                    fg.Cell(flexcpBackColor, lTempRow, 0, lTempRow, fg.Cols - 1) = kAtTheMoneyColor
                    dTempStrike = Val(.TextMatrix(lTempRow, GDCol(eGDCol_Strike)))
                    If lTempRow + 1 < .Rows Then
                        Do While Val(.TextMatrix(lTempRow + 1, GDCol(eGDCol_Strike))) = dTempStrike
                            fg.Cell(flexcpBackColor, lTempRow + 1, 0, lTempRow + 1, fg.Cols - 1) = kAtTheMoneyColor
                            lTempRow = lTempRow + 1
                            If lTempRow >= .Rows Then Exit Do
                        Loop
                    End If

                    For lStart = 1 To 2
                        lTempRow = lTempRow + 1
                        If lTempRow >= lNextHeader Then Exit For
                        fg.Cell(flexcpBackColor, lTempRow, 0, lTempRow, fg.Cols - 1) = kNearTheMoneyColor
                        dTempStrike = Val(.TextMatrix(lTempRow, GDCol(eGDCol_Strike)))
                        If lTempRow + 1 < .Rows Then
                            Do While Val(.TextMatrix(lTempRow + 1, GDCol(eGDCol_Strike))) = dTempStrike
                                fg.Cell(flexcpBackColor, lTempRow + 1, 0, lTempRow + 1, fg.Cols - 1) = kNearTheMoneyColor
                                lTempRow = lTempRow + 1
                                If lTempRow >= .Rows - 1 Then Exit Do
                            Loop
                        End If
                    Next lStart

                    lRow = lTempRow
                    If lRow = lNextHeader Then lRow = lRow - 1
                End If
            End If
        Next lRow
        
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptionChain.ColorRows", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FieldOrder.Let
'' Description: Allow for changing of the field order
'' Inputs:      New Field Order
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Let FieldOrder(ByVal strFieldOrder As String)
On Error GoTo ErrSection:

    Dim lCol As Long                    ' Index into a for loop
    Dim lPos As Long                    ' Index into a for loop
    Dim lTo As Long                     ' Where to move the column to
    Dim bHidden As Boolean              ' Is the column hidden?
    
    With fg
        For lCol = 0 To .Cols - 1
            lTo = CLng(Parse(Parse(strFieldOrder, ",", lCol + 1), ";", 1))
            bHidden = CBool(CLng(Parse(Parse(strFieldOrder, ",", lCol + 1), ";", 2)))
        
            For lPos = 0 To .Cols - 1
                If .ColData(lPos) = lCol Then
                    .ColHidden(lPos) = bHidden
                    .ColPosition(lPos) = lTo
                    Exit For
                End If
            Next lPos
        Next lCol
    End With
    
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "frmOptionChain.FieldOrder.Let", eGDRaiseError_Raise
    
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FieldOrder.Get
'' Description: Return the current field order
'' Inputs:      None
'' Returns:     Current Field Order
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get FieldOrder() As String
On Error GoTo ErrSection:

    Dim astrFields As cGdArray
    Dim lCol As Long
    
    Set astrFields = New cGdArray
    astrFields.Create eGDARRAY_Strings, fg.Cols
    
    With fg
        For lCol = 0 To .Cols - 1
            astrFields(.ColData(lCol)) = Str(lCol) & ";" & Str(CLng(.ColHidden(lCol))) & _
                    ";" & .TextMatrix(1, lCol)
        Next lCol
    End With
    
    FieldOrder = astrFields.JoinFields(",")

ErrExit:
    Set astrFields = Nothing
    Exit Property
    
ErrSection:
    Set astrFields = Nothing
    RaiseError "frmOptionChain.FieldOrder.Get", eGDRaiseError_Raise
    
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ChangeFields
'' Description: Allow the user to change the fields that are shown
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ChangeFields()
On Error GoTo ErrSection:

    Dim astrFields As cGdArray
    Dim astrDefaults As cGdArray
    Dim lCol As Long
    Dim lActive As Long
    Dim lIndex As Long
    Dim lIndex2 As Long
    Dim bColHidden As Boolean
    Dim strName As String
    Dim bShow As Boolean
    
    Set astrFields = New cGdArray
    astrFields.Create eGDARRAY_Strings
    Set astrDefaults = New cGdArray
    astrDefaults.Create eGDARRAY_Strings
    
    With fg
        For lCol = 0 To GDCol(eGDCol_Strike) - 1
            If .ColData(lCol) = GDCol(eGDCol_CallIndex) Then
                bShow = False
            ElseIf (.ColData(lCol) >= GDCol(eGDCol_CallDaysToExpire) And m.eSecType = eSYMType_Future And (Not DoFutOpts)) Then
                bShow = False
            Else
                bShow = True
            End If
        
            If .ColHidden(lCol) Then lActive = vbUnchecked Else lActive = vbChecked
            astrFields.Add Str(lActive) & vbTab & .TextMatrix(1, lCol) & vbTab & _
                        Str(.ColData(lCol)) & vbTab & Str(bShow)
                        
            bColHidden = CBool(Parse(Parse(m.strDefaultCols, ",", lCol + 1), ";", 2))
            lIndex = CLng(Parse(Parse(m.strDefaultCols, ",", lCol + 1), ";", 1))
            strName = Parse(Parse(m.strDefaultCols, ",", lCol + 1), ";", 3)
            If .ColData(lCol) = GDCol(eGDCol_CallIndex) Then
                bShow = False
            ElseIf (.ColData(lCol) >= GDCol(eGDCol_CallDaysToExpire) And m.eSecType = eSYMType_Future And (Not DoFutOpts)) Then
                bShow = False
            Else
                bShow = True
            End If
            
            If bColHidden Then lActive = vbUnchecked Else lActive = vbChecked
            astrDefaults(lIndex) = Str(lActive) & vbTab & strName & vbTab & _
                        Str(lCol) & vbTab & Str(bShow)
        Next lCol
    End With
    
    If frmQuoteBoardFields.ShowMe(astrFields, eQbfMode_OptionChain, astrDefaults) Then
        fg.Redraw = flexRDNone
        
        For lIndex = 0 To astrFields.Size - 1
            ' Get information out of the array returned to us...
            bColHidden = (Parse(astrFields(lIndex), vbTab, 1) = "2")
            lCol = CLng(Parse(astrFields(lIndex), vbTab, 3))
            
            ' Move around the appropriate Call column...
            For lIndex2 = 0 To GDCol(eGDCol_Strike) - 1
                If fg.ColData(lIndex2) = lCol Then
                    fg.ColHidden(lIndex2) = bColHidden
                    fg.ColPosition(lIndex2) = lIndex
                    Exit For
                End If
            Next lIndex2
            
            ' Move around the appropriate Put column...
            'lCol = GDCol(eGDCol_NumCols) - lCol - 1
            lCol = lCol + GDCol(eGDCol_Strike) + 1
            For lIndex2 = GDCol(eGDCol_Strike) + 1 To GDCol(eGDCol_NumCols) - 1
                If fg.ColData(lIndex2) = lCol Then
                    fg.ColHidden(lIndex2) = bColHidden
                    'fg.ColPosition(lIndex2) = GDCol(eGDCol_NumCols) - lIndex - 1
                    fg.ColPosition(lIndex2) = lIndex + GDCol(eGDCol_Strike) + 1
                    Exit For
                End If
            Next lIndex2
        Next lIndex
        
        ExtendCustomColumn
        fg.Redraw = flexRDBuffered
    
        SetIniFileProperty "DisplayFields" & m.eSecType, FieldOrder, "OptionChain", g.strIniFile
    End If

ErrExit:
    Set astrFields = Nothing
    Set astrDefaults = Nothing
    Exit Sub
    
ErrSection:
    Set astrFields = Nothing
    Set astrDefaults = Nothing
    RaiseError "frmOptionChain.ChangeFields", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OptionPrice
'' Description: Figure out the current option price for calculations
'' Inputs:      Bid, Ask, Last
'' Returns:     Option Price
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function OptionPrice(ByVal dBid As Double, ByVal dAsk As Double, ByVal dLast As Double) As Double
On Error GoTo ErrSection:

    Dim dTemp As Double
    
    ' First fix the Bid and Ask in case they are Null...
    If dBid = kNullData Or dBid < 0 Then dBid = 0
    If dAsk = kNullData Or dAsk < 0 Then dAsk = 0
    
    Select Case UCase(m.strOCB)
        Case "AVG"
            ' If the Bid and the Ask are both zero, then return zero...
            If dBid = 0 And dAsk = 0 Then
                OptionPrice = 0#
            ' If the Bid and the Ask are both non-zero, take the average of them...
            Else
                OptionPrice = Round((dBid + dAsk) / 2, 5)
            End If
            
        Case "BID"
            OptionPrice = dBid
            
        Case "ASK"
            OptionPrice = dAsk
            
        Case "LAST"
            OptionPrice = dLast
            
    End Select

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmOptionChain.OptionPrice", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ChangeSettings
'' Description: Allow the user to change settings for the option chain
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ChangeSettings()
On Error GoTo ErrSection:

    Dim lShowIRX As Long                ' Do we need to show the IRX columns?

    If frmOCSettings.ShowMe Then
        m.strOCB = GetIniFileProperty("OCB", "Avg", "OptionChain", g.strIniFile)
'        LoadGrid False
        RecalcAll
        lShowIRX = GetIniFileProperty("ShowIRX", vbUnchecked, "OptionChain", g.strIniFile)
        fgUnder.ColHidden(10) = (lShowIRX = vbUnchecked)
        fgUnder.ColHidden(11) = (lShowIRX = vbUnchecked)
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptionChain.ChangeSettings", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuRefresh_Click
'' Description: Allow the user to refresh the data
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuRefresh_Click()
On Error GoTo ErrSection:

    If Not ProcessIsBusy Then LoadGrid True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptionChain.mnuRefresh.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuSell_Click
'' Description: Allow the user to enter a sell order for the selected symbol
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuSell_Click()
On Error GoTo ErrSection:

    If FormIsLoaded("frmBrokerView") Then
        frmBrokerView.CreateOrderForLot "", SymbolForRow(fg.Row, (fg.Col <= GDCol(eGDCol_Strike))), "Sell"
    Else
        'CreateOrder fg.TextMatrix(fg.Row, fg.Col), , 0, , , "Option Chain"
        CreateOrder SymbolForRow(fg.Row, (fg.Col <= GDCol(eGDCol_Strike))), , 0, , , "Option Chain"
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptionChain.mnuSell.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuSettings_Click
'' Description: Allow the user to change settings for the option chain
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuSettings_Click()
On Error GoTo ErrSection:

    ChangeSettings

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptionChain.mnuSettings.Click", eGDRaiseError_Show
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadBars
'' Description: Load the bars for the given symbol or symbol ID
'' Inputs:      Bars to Load, Symbol or Symbol ID to load, Add to Realtime?
'' Returns:     True if DM data, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function LoadBars(Bars As cGdBars, ByVal vSymbolOrSymbolID As Variant, Optional ByVal bAddToRT As Boolean = True) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value from Data Manager load

' TLB 4/2/2008: no more realtime for the old option chain
bAddToRT = False

    Bars.ArrayMask = eBARS_EodBidAsk
    bReturn = DM_GetBars(Bars, vSymbolOrSymbolID, , LastDailyDownload - 5, , , False, , 2)
    If bAddToRT = True And HasModule("OPTNAV") Then
        g.RealTime.AddTickBuffer Bars, False
        g.RealTime.SpliceBars Bars
    End If
    If Bars.Prop(eBARS_IsOption) <> 0 Then
        NumDaysToExpire Bars
    End If
    
    LoadBars = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmOptionChain.LoadBars", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetBars
'' Description: Get the Bars for the given symbol from the collection (add it
''              to the collection if not there)
'' Inputs:      Symbol to get Bars for
'' Returns:     Bars
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetBars(ByVal vSymbolOrSymbolID As Variant, Optional ByVal bAddToRT As Boolean = True) As cGdBars
On Error GoTo ErrSection:

    Dim Bars As New cGdBars             ' Temporary Bars structure
    
    If m.BarsColl.Exists(Str(vSymbolOrSymbolID)) = False Then
        LoadBars Bars, vSymbolOrSymbolID, bAddToRT
        m.BarsColl.Add Bars, Str(vSymbolOrSymbolID)
    End If
    
    Set GetBars = m.BarsColl(Str(vSymbolOrSymbolID))

ErrExit:
    Set Bars = Nothing
    Exit Function
    
ErrSection:
    Set Bars = Nothing
    RaiseError "frmOptionChain.GetBars", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tbToolbar_ToolClick
'' Description: Handle a user selection on the toolbar
'' Inputs:      Tool Clicked
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tbToolbar_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
On Error GoTo ErrSection:

    If Me.Visible Then
        Select Case Tool.ID
            Case "ID_Close"
                DockState(Me) = eHidden
                frmMain.tbToolbar.Tools("ID_Chain").State = ssUnchecked
            
            Case "ID_Fields"
                ChangeFields
                
            Case "ID_Refresh"
                If Not ProcessIsBusy Then LoadGrid True
            
            Case "ID_Settings"
                ChangeSettings
                
            Case "ID_Symbol"
                ChangeSymbol
                
            Case "ID_Greeks"
                RunOptionsCalc
                
            Case "ID_Probability"
                RunProbabilityCalc
            
        End Select
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptionChain.tbToolbar.ToolClick", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tmrRealTime_Timer
'' Description: Do any real time updates that are necessary
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tmrRealTime_Timer()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lSymbolID As Long               ' Symbol ID of the updating symbol
    Dim bNewBar As Boolean              ' Is this a new bar?
    Dim bUnderChanged As Boolean        ' Underlying close changed
    Dim bIRXChanged As Boolean          ' IRX close changed
    Dim lCount As Long                  ' Bars collection count
    
' TLB 4/2/2008: no more realtime for the old option chain
tmrRealtime.Enabled = False
Exit Sub
    
    If Me.Visible Then
        m.bInProgress = True
    
        lCount = m.BarsColl.Count
        For lIndex = 1 To lCount
            ' If the size of the bars collection has changed on us, stop processing...
            If lCount <> m.BarsColl.Count Then Exit For
            If m.bSymbolChanged Then
                m.bSymbolChanged = False
                Exit For
            End If
            
            lSymbolID = m.BarsColl(lIndex).Prop(eBARS_SymbolID)

' TLB 10/19/2007: turn off Greek recalc for now (since will be moving to the new OptNav)
            
            If g.RealTime.UpdateBars(m.BarsColl(lIndex), bNewBar) = True Then
                If bNewBar Then
                    If lSymbolID = 0 Then
                        LoadBars m.BarsColl(lIndex), m.BarsColl(lIndex).Prop(eBARS_Symbol)
                    Else
                        LoadBars m.BarsColl(lIndex), m.BarsColl(lIndex).Prop(eBARS_SymbolID)
                    End If
                End If
            
                If Not m.bSymbolChanged Then
                    Select Case lSymbolID
                        Case m.lSymbolID
                            bUnderChanged = DisplayInfo(m.BarsColl(lIndex))
                            If bUnderChanged Then ColorRows False
                        Case m.lIRXSymbolID
                            bIRXChanged = DisplayIRX(m.BarsColl(lIndex))
                        Case Else
                            UpdateSymbol m.BarsColl(lIndex).Prop(eBARS_Symbol), m.BarsColl(lIndex)
                            ''Calculate m.BarsColl(lIndex).Prop(eBARS_Symbol)
                    End Select
                End If
            ElseIf (lSymbolID = 0) And (bUnderChanged Or bIRXChanged) Then
                ''Calculate m.BarsColl(lIndex).Prop(eBARS_Symbol)
            End If
        Next lIndex
        
        ClearUpdatedColors
    End If
    m.bInProgress = False
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptionChain.tmrRealTime_Timer", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DisplayIRX
'' Description: Display the IRX data in a label
'' Inputs:      Bars with IRX data
'' Returns:     True if Close changes, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DisplayIRX(ByVal Bars As cGdBars) As Boolean
On Error GoTo ErrSection:

    Dim strPrice As String              ' Current price of the $IRX index

    With fgUnder
        .Redraw = flexRDNone
        
        If Bars Is Nothing Then
            strPrice = ""
        ElseIf Bars(eBARS_Close, Bars.Size - 1) = kNullData Then
            strPrice = ""
        Else
            strPrice = Bars.PriceDisplay(Bars(eBARS_Close, Bars.Size - 1), True)
        End If
        
        DisplayIRX = ChangeCell(fgUnder, 0, 11, strPrice)
        
        .AutoSize 0, .Cols - 1, False
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmOptionChain.DisplayIRX", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitUnderGrid
'' Description: Initialize the underlying data grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitUnderGrid()
On Error GoTo ErrSection:
    
    Dim lShowIRX As Long                ' Do we need to show the IRX columns?

    With fgUnder
        .Redraw = flexRDNone
        
        .AllowBigSelection = False
        .AllowSelection = False
        .Appearance = flexFlat
        .BackColorFixed = Me.BackColor
        .BorderStyle = flexBorderNone
        .Editable = flexEDNone
        .ExplorerBar = flexExNone
        .ExtendLastCol = True
        .FocusRect = flexFocusNone
        .GridLinesFixed = flexGridNone
        .ScrollBars = flexScrollBarNone
        .SheetBorder = vbWindowBackground
        
        .Rows = 1
        .FixedRows = 1
        .Cols = 12
        .FixedCols = 0
        
        .TextMatrix(0, 0) = "Last Tick:"
        .TextMatrix(0, 2) = "Last:"
        .TextMatrix(0, 4) = "Bid:"
        .TextMatrix(0, 6) = "Ask:"
        .TextMatrix(0, 8) = "Vol:"
        .TextMatrix(0, 10) = "$IRX:"
        
        lShowIRX = GetIniFileProperty("ShowIRX", vbUnchecked, "OptionChain", g.strIniFile)
        .ColHidden(10) = (lShowIRX = vbUnchecked)
        .ColHidden(11) = (lShowIRX = vbUnchecked)
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignLeftTop
        .AutoSize 0, .Cols - 1, False
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptionChain.InitUnderGrid", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ChangeCell
'' Description: To change text and forecolor of grid cell in grid
'' Inputs:      Grid, Row and Column to change, New Value
'' Returns:     True if value changed, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ChangeCell(Grid As VSFlexGrid, ByVal lRow As Long, ByVal lCol As Long, ByVal strCellText As String) As Boolean
On Error GoTo ErrSection:

    Dim nForeColor As Long              ' Fore color of the cell to update
    Dim dTickCount As Double            ' Current tick count to store in cell
    
    ChangeCell = False
    With Grid
        nForeColor = frmQuotes.UnchColor
        If .TextMatrix(lRow, lCol) <> strCellText Then
            ChangeCell = True
            .TextMatrix(lRow, lCol) = strCellText
            If tmrRealtime.Enabled Then
                nForeColor = frmQuotes.UpdateColor
                
                ' save TickCount for this cell (so will turn back to black after 1 second)
                .Cell(flexcpData, lRow, lCol) = gdTickCount
            End If
        ElseIf tmrRealtime.Enabled Then
            dTickCount = .Cell(flexcpData, lRow, lCol)
            dTickCount = gdTickCount - dTickCount
            If dTickCount >= 0 And dTickCount <= g.nUpdatedColorDuration Then
                nForeColor = frmQuotes.UpdateColor
            End If
        End If
        .Cell(flexcpForeColor, lRow, lCol) = nForeColor
    End With
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmOptionChain.ChangeCell", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ClearUpdatedColors
'' Description: Clear any updated colors in the grids
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ClearUpdatedColors()
On Error GoTo ErrSection:

    Dim lRow As Long                    ' Index into a for loop
    Dim lCol As Long                    ' Index into a for loop
    Dim dTickCount As Double            ' Tick count stored in the cell
    Dim lRedraw As Long                 ' Current state of the grid's redraw

    If g.bUnloading Then Exit Sub
    
    With fgUnder
        lRedraw = .Redraw
        .Redraw = flexRDNone
                
        For lCol = 0 To .Cols - 1
            If tmrRealtime.Enabled Then
                If .Cell(flexcpForeColor, 0, lCol) = frmQuotes.UpdateColor Then
                    dTickCount = .Cell(flexcpData, 0, lCol)
                    dTickCount = gdTickCount - dTickCount
                    
                    If dTickCount >= 0 And dTickCount <= g.nUpdatedColorDuration Then
                        .Cell(flexcpForeColor, 0, lCol) = frmQuotes.UpdateColor
                    Else
                        .Cell(flexcpForeColor, 0, lCol) = frmQuotes.UnchColor
                    End If
                End If
            Else
                .Cell(flexcpForeColor, 0, lCol) = frmQuotes.UnchColor
            End If
        Next lCol
        
        .Redraw = lRedraw
    End With
    
    With fg
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        For lRow = .FixedRows To .Rows - 1
            For lCol = 0 To .Cols - 1
                If tmrRealtime.Enabled Then
                    If .Cell(flexcpForeColor, lRow, lCol) = frmQuotes.UpdateColor Then
                        dTickCount = .Cell(flexcpData, lRow, lCol)
                        dTickCount = gdTickCount - dTickCount
                        
                        If dTickCount >= 0 And dTickCount <= g.nUpdatedColorDuration Then
                            .Cell(flexcpForeColor, lRow, lCol) = frmQuotes.UpdateColor
                        Else
                            .Cell(flexcpForeColor, lRow, lCol) = frmQuotes.UnchColor
                        End If
                    End If
                Else
                    .Cell(flexcpForeColor, lRow, lCol) = frmQuotes.UnchColor
                End If
            Next lCol
        Next lRow
        
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmOptionChain.ClearUpdatedColors", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateSymbol
'' Description: Update the columns for the given symbol and bar data
'' Inputs:      Symbol, Bars
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UpdateSymbol(ByVal strSymbol As String, ByVal Bars As cGdBars)
On Error GoTo ErrSection:

    Dim lRow As Long                    ' Row in the grid for the symbol
    Dim strType As String               ' Call or Put
    Dim lFrom As Long                   ' From value for a for loop
    Dim lTo As Long                     ' To value for a for loop
    Dim lCol As Long                    ' Index into a for loop
    Dim dValue As Double                ' Value to store in the grid
    Dim strValue As String              ' Formatted value to put in the grid
    Dim lTblIndex As Long               ' Index into the information table
    Dim lRedraw As Long                 ' Current state of the grid's redraw
    
    If RowInfo(strSymbol, lRow, strType) Then
        If UCase(strType) = "CALL" Then
            lFrom = 0
            lTo = GDCol(eGDCol_Strike) - 1
        Else
            lFrom = GDCol(eGDCol_Strike) + 1
            lTo = GDCol(eGDCol_NumCols) - 1
        End If
        
        lTblIndex = TblIndexForSymbol(Bars.Prop(eBARS_Symbol))
        
        With fg
            lRedraw = .Redraw
            .Redraw = flexRDNone
            
            For lCol = lFrom To lTo
                strValue = ""
                
                Select Case UCase(.TextMatrix(1, lCol))
                    Case "DATE"
                        dValue = Bars(eBARS_DateTime, Bars.Size - 1)
                        If dValue > 0 Then
                            strValue = DateFormat(dValue, MM_DD_YY)
                        End If
                        ChangeCell fg, lRow, lCol, strValue
                    Case "TIME"
                        If Bars(eBARS_DateTime, Bars.Size - 1) > 0 Then
                            dValue = Bars.Prop(eBARS_LastTickTime) / 1440#
                            If dValue <> 0 Then
                                If g.bShowInLocalTimeZone Then
                                    dValue = ConvertTimeZone(dValue, Bars.Prop(eBARS_ExchangeTimeZoneInf), "")
                                End If
                                strValue = Format(dValue, "hh:mm:ss")
                            End If
                        End If
                        ChangeCell fg, lRow, lCol, strValue
                    Case "OPEN"
                        dValue = Bars(eBARS_Open, Bars.Size - 1)
                        If dValue <> kNullData Then
                            strValue = Bars.PriceDisplay(dValue, True)
                        End If
                        ChangeCell fg, lRow, lCol, strValue
                    Case "HIGH"
                        dValue = Bars(eBARS_High, Bars.Size - 1)
                        If dValue <> kNullData Then
                            strValue = Bars.PriceDisplay(dValue, True)
                        End If
                        ChangeCell fg, lRow, lCol, strValue
                    Case "LOW"
                        dValue = Bars(eBARS_Low, Bars.Size - 1)
                        If dValue <> kNullData Then
                            strValue = Bars.PriceDisplay(dValue, True)
                        End If
                        ChangeCell fg, lRow, lCol, strValue
                    Case "LAST"
                        dValue = Bars(eBARS_Close, Bars.Size - 1)
                        If dValue <> kNullData Then
                            strValue = Bars.PriceDisplay(dValue, True)
                        End If
                        ChangeCell fg, lRow, lCol, strValue
                    Case "BID"
                        dValue = Bars(eBARS_Bid, Bars.Size - 1)
                        If dValue <> kNullData Then
                            strValue = Bars.PriceDisplay(dValue, True)
                        End If
                        ChangeCell fg, lRow, lCol, strValue
                    Case "ASK"
                        dValue = Bars(eBARS_Ask, Bars.Size - 1)
                        If dValue <> kNullData Then
                            strValue = Bars.PriceDisplay(dValue, True)
                        End If
                        ChangeCell fg, lRow, lCol, strValue
                    Case "VOL"
                        dValue = Bars(eBARS_Vol, Bars.Size - 1)
                        If dValue <> kNullData Then
                            strValue = Format(dValue, "#,##0")
                        End If
                        ChangeCell fg, lRow, lCol, strValue
                    Case "OI"
                        dValue = Bars(eBARS_ContOI, Bars.Size - 1)
                        If dValue <> kNullData Then
                            strValue = Format(dValue, "#,##0")
                        End If
                        ChangeCell fg, lRow, lCol, strValue
                    Case "DAYS LEFT"
                        If lTblIndex >= 0 Then
                            dValue = m.tblOption(TblCol(etblCols_DaysLeft), lTblIndex)
                        Else
                            dValue = kNullData
                        End If
                        If dValue <> kNullData Then
                            strValue = Str(dValue)
                        End If
                        ChangeCell fg, lRow, lCol, strValue
                End Select
            Next lCol
            
            .Redraw = lRedraw
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptionChain.UpdateSymbol", eGDRaiseError_Raise
    
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RecalcAll
'' Description: Calculate all Greeks, Volatility etc. for each symbol (All rows in grid)
'' Inputs:      None
'' Returns:     None
''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RecalcAll()
On Error GoTo ErrSection:

    Dim lCount As Long
    Dim lIndex As Long
    
    lCount = m.BarsColl.Count
    For lIndex = 1 To lCount
        Calculate m.BarsColl(lIndex).Prop(eBARS_Symbol)
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptionChain.RecalcAll", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Calculate
'' Description: Calculate the values and the Greeks for a symbol
'' Inputs:      Symbol to calculate
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Calculate(ByVal strSymbol As String)
On Error GoTo ErrSection:

    Dim lRow As Long                    ' Row that the symbol is in
    Dim strType As String               ' Call or Put?
    Dim Bars As New cGdBars             ' Bars for the symbol
    Dim lNumDays As Long                ' Number of days to expiration
    Dim bIsPut As Boolean               ' Is this a put?
    Dim dOptPrice As Double             ' Option "price"
    Dim dUnderlying As Double           ' Current price of the underlying
    Dim dIntRate As Double              ' Current interest rate
    Dim dCarryCost As Double            ' Carry cost depending on security type
    Dim dStrike As Double               ' Strike price for the symbol
    Dim dIntrinsic As Double            ' Intrinsic value of the option
    Dim dTimeValue As Double            ' Time value of the option
    Dim dVolatility As Double           ' Volatility for the option
    Dim dDelta As Double                ' Delta for the option
    Dim lTblIndex As Long               ' Index into the information table
    Dim dGamma As Double                ' Gamma for the option
    Dim dVega As Double                 ' Vega for the option
    Dim dTheta As Double                ' Theta for the option
    Dim dRho As Double                  ' Rho for the option
    Dim lIndex As Long                  ' Index into a for loop
    Dim lFrom As Long                   ' Value to run the for loop from
    Dim lTo As Long                     ' Value to run the for loop to
    Dim strValue As String              ' Value to put into the grid
    Dim lRedraw As Long                 ' Current state of the grid's redraw

    If RowInfo(strSymbol, lRow, strType) Then
        If UCase(strType) = "CALL" Then
            bIsPut = False
            'lTblIndex = CLng(Val(fg.TextMatrix(lRow, GDCol(eGDCol_CallIndex))))
            lFrom = 0&
            lTo = GDCol(eGDCol_Strike) - 1
        Else
            bIsPut = True
            'lTblIndex = CLng(Val(fg.TextMatrix(lRow, GDCol(eGDCol_PutIndex))))
            lFrom = GDCol(eGDCol_Strike) + 1
            lTo = fg.Cols - 1
        End If
        
        lTblIndex = TblIndexForSymbol(strSymbol)
        
        Set Bars = m.BarsColl(strSymbol)
        lNumDays = m.tblOption.Num(TblCol(etblCols_DaysLeft), lTblIndex)
        If lNumDays = kNullData Then
            dIntrinsic = kNullData
            dTimeValue = kNullData
            dVolatility = kNullData
            dDelta = kNullData
            dGamma = kNullData
            dVega = kNullData
            dTheta = kNullData
            dRho = kNullData
        Else
            dOptPrice = OptionPrice(Bars(eBARS_Bid, Bars.Size - 1), Bars(eBARS_Ask, Bars.Size - 1), Bars(eBARS_Close, Bars.Size - 1))
            dUnderlying = m.BarsColl(Str(m.lSymbolID)).Item(eBARS_Close, m.BarsColl(Str(m.lSymbolID)).Size - 1)
            dIntRate = m.BarsColl(Str(m.lIRXSymbolID)).Item(eBARS_Close, m.BarsColl(Str(m.lIRXSymbolID)).Size - 1)
            If dIntRate < 0 Then
                dIntRate = 0
            Else
                dIntRate = dIntRate / 100
            End If
            
            Select Case m.eSecType
                Case eSYMType_Future
                    dCarryCost = 0#
                Case eSYMType_Stock
                    dCarryCost = dIntRate
                Case eSYMType_Index
                    dCarryCost = dIntRate
            End Select
        
            dStrike = Val(fg.TextMatrix(lRow, GDCol(eGDCol_Strike)))
    
            If bIsPut Then
                If dStrike - dUnderlying < 0 Then dIntrinsic = 0 Else dIntrinsic = dStrike - dUnderlying
            Else
                If dUnderlying - dStrike < 0 Then dIntrinsic = 0 Else dIntrinsic = dUnderlying - dStrike
            End If
            
            If dOptPrice - dIntrinsic < 0 Then dTimeValue = 0 Else dTimeValue = dOptPrice - dIntrinsic
            
            dVolatility = Opt_GetVolatility(dOptPrice, dUnderlying, dStrike, lNumDays, dIntRate, dCarryCost, bIsPut, dDelta)
                        
            If dDelta > 2# Or dDelta < -2# Then
                dDelta = kNullData
            End If
            
            dGamma = Opt_Gamma(dUnderlying, dStrike, lNumDays, dVolatility, dIntRate, dCarryCost)
            
            If Round(dVolatility, 5) = 0# Then
                dVega = kNullData
            Else
                dVega = Opt_Vega(dUnderlying, dStrike, lNumDays, dVolatility, dIntRate, dCarryCost)
            End If
            
            dTheta = Opt_Theta(dUnderlying, dStrike, lNumDays, dVolatility, dIntRate, dCarryCost, bIsPut)
            
            dRho = Opt_Rho(dUnderlying, dStrike, lNumDays, dVolatility, dIntRate, dCarryCost, bIsPut)
        End If
        
        With fg
            lRedraw = .Redraw
            .Redraw = flexRDNone
            
            For lIndex = lFrom To lTo
                strValue = ""
                Select Case UCase(.TextMatrix(1, lIndex))
                    Case "INTR VAL"
                        If dIntrinsic <> kNullData Then
                            strValue = Format(dIntrinsic, "##0.00")
                        End If
                        ChangeCell fg, lRow, lIndex, strValue
                    Case "TIME VAL"
                        If dTimeValue <> kNullData Then
                            strValue = Format(dTimeValue, "##0.00")
                        End If
                        ChangeCell fg, lRow, lIndex, strValue
                    Case "IMPL VOL"
                        If dVolatility <> kNullData Then
                            strValue = Format(dVolatility * 100#, "##0.00")
                        End If
                        ChangeCell fg, lRow, lIndex, strValue
                    Case "DELTA"
                        If dDelta <> kNullData Then
                            strValue = Format(dDelta, "0.0000")
                        End If
                        ChangeCell fg, lRow, lIndex, strValue
                    Case "GAMMA"
                        If dGamma <> kNullData Then
                            strValue = Format(dGamma, "0.0000")
                        End If
                        ChangeCell fg, lRow, lIndex, strValue
                    Case "VEGA"
                        If dVega <> kNullData Then
                            strValue = Format(dVega / 100#, "0.0000")
                        End If
                        ChangeCell fg, lRow, lIndex, strValue
                    Case "THETA"
                        If dTheta <> kNullData Then
                            strValue = Format(dTheta / 365#, "0.0000")
                        End If
                        ChangeCell fg, lRow, lIndex, strValue
                    Case "RHO"
                        If dRho <> kNullData Then
                            strValue = Format(dRho / 100#, "0.0000")
                        End If
                        ChangeCell fg, lRow, lIndex, strValue
                End Select
            Next lIndex
            
            .Redraw = lRedraw
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptionChain.Calculate", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RowInfo
'' Description: Given a symbol, determine the row and type of the symbol
'' Inputs:      Symbol, Row, Type, Table Index
'' Returns:     True if Found, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function RowInfo(ByVal strSymbol As String, lRow As Long, strType As String, Optional lTblIndex As Long) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Was the symbol found in the lookup table?
    Dim lPos As Long                    ' Position of the symbol in the lookup table
    
    bReturn = m.astrLookup.BinarySearch(strSymbol & vbTab, lPos, eGdSort_MatchUsingSearchStringLength)
    If bReturn Then
        lRow = CLng(Val(Parse(m.astrLookup(lPos), vbTab, 2)))
        strType = Parse(m.astrLookup(lPos), vbTab, 3)
        lTblIndex = CLng(Val(Parse(m.astrLookup(lPos), vbTab, 4)))
    End If
    
    RowInfo = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmOptionChain.RowInfo", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NumDaysToExpire
'' Description: Number of days left before expiration
'' Inputs:      Bars
'' Returns:     Number of Days
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function NumDaysToExpire(ByVal Bars As cGdBars) As Long
On Error GoTo ErrSection:

    Dim lYear As Long                   ' Year of the option
    Dim lMonth As Long                  ' Month of the option
    Dim dDate As Double                 ' Last date of data for the option
    Dim lReturn As Long                 ' Return value for the function
    Dim lTblIndex As Long               ' Index into the table
    Dim lExpDate As Long                ' Expiration Date for a Future Option
    Dim strSymbol As String             ' Symbol from the Bars
    Dim strExpDate As String            ' Expiration date from the SO symbol

    dDate = Bars(eBARS_DateTime, Bars.Size - 1)
    
    If dDate <> kNullData Then
        If m.eSecType = eSYMType_Stock Or m.eSecType = eSYMType_Index Then
            strSymbol = Bars.Prop(eBARS_Symbol)
            
            If Len(strSymbol) <= 6 Then
                lYear = CLng(Val(Left(Parse(Bars.Prop(eBARS_Desc), " ", 2), 4)))
                lMonth = CLng(Val(Mid(Parse(Bars.Prop(eBARS_Desc), " ", 2), 5, 2)))
                lReturn = GetDateFromRule(lYear, lMonth, "3F") - Int(dDate) + 1&
            Else
                strExpDate = Parse(strSymbol, " ", 2)
                lReturn = JulFromLong(CLng(Val(strExpDate))) - Int(dDate) + 1&
            End If
        ElseIf DoFutOpts And m.eSecType = eSYMType_Future Then
            If SU_GetFutureOptionExp(Bars.Prop(eBARS_Symbol), lExpDate) Then
                lReturn = lExpDate - Int(dDate) + 1
            End If
        End If
    Else
        lReturn = kNullData
    End If
    
    lTblIndex = TblIndexForSymbol(Bars.Prop(eBARS_Symbol))
    If lTblIndex >= 0 Then
        m.tblOption.Num(TblCol(etblCols_DaysLeft), lTblIndex) = lReturn
    End If
    
    NumDaysToExpire = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmOptionChain.NumDaysToExpire", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TblIndexForSymbol
'' Description: Determine the table index for a given symbol
'' Inputs:      Symbol
'' Returns:     Table Index
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function TblIndexForSymbol(ByVal strSymbol As String) As Long
On Error GoTo ErrSection:

    Dim lTblIndex As Long               ' Index into the table
    
    If m.tblOption.SearchAsIndex(m.alIndex, TblCol(etblCols_Symbols), strSymbol, lTblIndex) Then
        TblIndexForSymbol = lTblIndex
    Else
        TblIndexForSymbol = -1&
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmOptionChain.TblIndexForSymbol", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadData
'' Description: Load up all of the data we need and display
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub LoadData()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lSymbolID As Long               ' Symbol ID of the updating symbol
    
    If Me.Visible Then
        If g.RealTime.IsServerActive Then
            Caption = "Option Chain for " & m.strSymbol
        Else
            Caption = "Option Chain for " & m.strSymbol & "  (data from previous update)"
        End If
    
        If g.RealTime.IsServerActive Then frmMain.SuspendNewSymbolCheck 15
        For lIndex = 1 To m.BarsColl.Count
            lSymbolID = m.BarsColl(lIndex).Prop(eBARS_SymbolID)
            If lSymbolID = 0 Then
                LoadBars m.BarsColl(lIndex), m.BarsColl(lIndex).Prop(eBARS_Symbol)
            Else
                LoadBars m.BarsColl(lIndex), m.BarsColl(lIndex).Prop(eBARS_SymbolID)
            End If
            
            Select Case lSymbolID
                Case m.lSymbolID
                    If DisplayInfo(m.BarsColl(lIndex)) Then ColorRows False
                Case m.lIRXSymbolID
                    DisplayIRX m.BarsColl(lIndex)
                Case Else
                    UpdateSymbol m.BarsColl(lIndex).Prop(eBARS_Symbol), m.BarsColl(lIndex)
                    Calculate m.BarsColl(lIndex).Prop(eBARS_Symbol)
            End Select
        Next lIndex
        ExtendCustomColumn
        If g.RealTime.IsServerActive Then frmMain.SuspendNewSymbolCheck
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptionChain.LoadData", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ChangeSymbol
'' Description: Change the underlying symbol
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ChangeSymbol()
On Error GoTo ErrSection:

    Dim astrSymbols As New cGdArray     ' Array of symbols back from symbol selector
    Dim lSymbolID As Long               ' Symbol ID for the symbol selected
    Dim bTimer As Boolean               ' Is the timer on or off?
    Dim lIndex As Long                  ' Index into a for loop
    
    Set astrSymbols = frmSymbolSelector.ShowMe(m.strSymbol, False, True, "Select Symbol for Option Chain")
    If astrSymbols.Size > 0 Then
        Screen.MousePointer = vbHourglass
        bTimer = tmrRealtime.Enabled
        tmrRealtime.Enabled = False
        lSymbolID = g.SymbolPool.SymbolIDforSymbol(astrSymbols(0))
        If (lSymbolID <> 0) And (lSymbolID <> m.lSymbolID) Then
            m.bSymbolChanged = True
            For lIndex = m.BarsColl.Count To 1 Step -1
                m.BarsColl.Remove lIndex
            Next lIndex
            DoEvents
            ShowMe lSymbolID
        End If
        tmrRealtime.Enabled = bTimer
        Screen.MousePointer = vbDefault
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptionChain.ChangeSymbol", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RunProbabilityCalc
'' Description: Run the Monte Carlo Options Probability Calculator
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RunProbabilityCalc()
On Error GoTo ErrSection:

    Dim dStockPrice As Double           ' Current stock price
    Dim lNumDays As Long                ' Number of days until expiration
    Dim dVolatility As Double           ' Volatility
    Dim lStrike As Long                 ' Strike price of current row
    Dim dHighPrice As Double            ' High side price
    Dim dLowPrice As Double             ' Low side price
    Dim dCallAsk As Double              ' Ask price of the call
    Dim dPutAsk As Double               ' Ask price of the put
    Dim dStrike As Double               ' Strike Price
    Dim Bars As New cGdBars             ' Temporary bars object
    Dim lDate As Long                   ' Date in exchange time
    Dim strArgs As String               ' Arguments to the probabilty calculator
    
    If fg.Row >= fg.FixedRows And fg.Row < fg.Rows Then
        dVolatility = Val(fg.TextMatrix(fg.Row, GDCol(eGDCol_CallImplVol)))
        lNumDays = CLng(Val(fg.TextMatrix(fg.Row, GDCol(eGDCol_CallDaysToExpire))))
        dStrike = Val(fg.TextMatrix(fg.Row, GDCol(eGDCol_Strike)))
        
        'Set Bars = m.BarsColl(fg.TextMatrix(fg.Row, GDCol(eGDCol_CallSymbol)))
        Set Bars = m.BarsColl(SymbolForRow(fg.Row, True))
        dCallAsk = Bars(eBARS_Ask, Bars.Size - 1)
        If dCallAsk = kNullData Then dCallAsk = 0#
        'Set Bars = m.BarsColl(fg.TextMatrix(fg.Row, GDCol(eGDCol_PutSymbol)))
        Set Bars = m.BarsColl(SymbolForRow(fg.Row, False))
        dPutAsk = Bars(eBARS_Ask, Bars.Size - 1)
        If dPutAsk = kNullData Then dPutAsk = 0#
        Set Bars = m.BarsColl(Str(m.lSymbolID))
        dStockPrice = Bars(eBARS_Close, Bars.Size - 1)
        If dStockPrice = kNullData Then dStockPrice = 0#
        lDate = ConvertTimeZone(Date, "", Bars.Prop(eBARS_ExchangeTimeZoneInf))
        
        dHighPrice = dStrike + (dCallAsk + dPutAsk)
        dLowPrice = dStrike - (dCallAsk + dPutAsk)
        lNumDays = NumBusinessDays(lDate, lDate + lNumDays)
        
        strArgs = "/P=" & Str(dStockPrice) & " /H=" & Str(dHighPrice) & " /L=" & Str(dLowPrice) & " /D=" & Str(lNumDays) & " /V=" & Str(dVolatility / 100#)
    End If
    
    RunProcess m.strProbCalc, strArgs
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptionChain.RunProbabilityCalc", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EnableControls
'' Description: Enable/Disable controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EnableControls()
On Error GoTo ErrSection:

    Dim bValidSymbolRow As Boolean      ' Is the current row in the grid a valid symbol row?
    Dim bEnableGreeks As Boolean        ' Enable the Greeks Calculator?
    Dim bEnableProb As Boolean          ' Enable the Probability Calculator?

    With tbToolbar
        .Tools("ID_Probability").Visible = HasModule("GOPC") And FileExist(m.strProbCalc)
        .Tools("ID_Greeks").Visible = HasModule("GOPC")
    End With
    
    bValidSymbolRow = False
    If (fg.Row >= fg.FixedRows) And (fg.Row < fg.Rows) Then
        If fg.RowData(fg.Row) = -1 Then bValidSymbolRow = True
    End If
    
    bEnableGreeks = bValidSymbolRow
    
    Enable mnuGreeks, bEnableGreeks
    tbToolbar.Tools("ID_Greeks").Enabled = bEnableGreeks

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptionChain.EnableControls", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SymbolForRow
'' Description: Return the symbol for the given row
'' Inputs:      Row, Put/Call
'' Returns:     Symbol
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SymbolForRow(ByVal lRow As Long, ByVal bCall As Boolean) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim lIndex As Long                  ' Index into the table
    
    If (lRow >= fg.FixedRows) And (lRow < fg.Rows) Then
        If bCall Then
            lIndex = CLng(Val(fg.TextMatrix(lRow, GDCol(eGDCol_CallIndex))))
        Else
            lIndex = CLng(Val(fg.TextMatrix(lRow, GDCol(eGDCol_PutIndex))))
        End If
        
        If (lIndex >= 0) And (lIndex < m.tblOption.NumRecords) Then
            strReturn = m.tblOption(TblCol(etblCols_Symbols), lIndex)
        End If
    End If
    
    SymbolForRow = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmOptionChain.SymbolForRow"
    
End Function

