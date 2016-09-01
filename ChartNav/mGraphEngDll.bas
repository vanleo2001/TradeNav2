Attribute VB_Name = "mGraphEngDll"
Option Explicit

Type chart_win
  paneCount As Long         'number of panes in chart area
  chartFgColor As Long      'colors are COLORREF values
  chartBkColor As Long
  borderFgColor As Long
  borderBkColor As Long
  LogoColor As Long         'for export & print, -1=no logo
  prtOrientation As Long    '0=landscape, 1=portrait
  DateStyle As Long         '0=Feb-99,1=02/19/99
  y_scaleLoc As Long        '0=default=right, 1=left
  y_labelMaxChar As Long    'max number of characters in label
  x_dataPointPix As Long    'pixels per data point
  x_dataPointCount As Long  '# of data points in chartable area
  fWidth As Long            'set by the graphics engine for axis labels
  fHeight As Long           'set by the graphics engine for axis labels
  fSize As Long             'font size used
  fStyle As Long            '0=reg,1=bold,2=italic,3=bold italic
  paneSepWidth As Long      'width of pane separator in pixels
  paneSepHide As Long       '0=default=show separator, 1=hide
  gridMode As Long          '0=vtCoarse,1=vtFine,2=bothCoarse,3=bothFine,4=hzOnly,5=noGrid
  Periodicity As Long       'minutes for intraday bars (e.g. 5-min bar, 15-min bar etc)
  crossOverTime As Double   'for overnight symbols
  hitTestPix As Long        '# of pixels closeness to an item to consider a hit
  glhBars As Long           'handle to gdBars
  splitPaneWidth As Long    'largest width among the various split pane's width (set by graphics engine - do not use)
  scaleArrow As Long        '0:none, 1:point left, 2:point right
  PriceTopMost As Long      '0:default draw order, 1:draw price on top of all indicators
  glhAvailFsize As Long     'available font size
  glhWinRect As Long        '(top,left,bottom,right) of window
  glhChartRect As Long      '(top,left,bottom,right) of chart area
  gdhPrintMargin As Long    '(top,left,bottom,right) printer margins in inches
  gdhDate As Long           'date data for chart
  gshFont As Long           'name of font used
  gshTimeZoneInf As Long    'gdstring holding time zone to display datetime in
  GradientColor As Long     '-1=no gradient
  SeasonalPix As Long       'do not use (reserved for grapheng internal use)
End Type

Type chart_pane
  hObj As Long              'handle to owning object for internal use
  paneId As Long
  paneShow As Long
  indicatorCount As Long
  y_labelAll As Long       '0=label only major ticks in y-scale, 1=label as many as possible, 2=time cluster pane (same behavior as 0)
  y_labelSpace As Long     'height of indicator labels in pixels (set by graphics engine - do not use)
  y_topSepCoord As Long    'coordinate value in pixel set by graphics engine
  y_btmSepCoord As Long    'coordinate value in pixel set by graphics engine
  y_scaleType As Long      '0=default=normal, 1=semi log, 2=square chart, 3=scale to price, 4=log AND log mode draw, 99=reserved for graphics engine
  paneSepHide As Long      '0=default=show separator, 1=hide
  IsPricePane As Long      '0=not price pane, 1=is price pane
  reserved As Long         'usage TBD
  splitPaneColor As Long   'color for vertical split pane
  splitPaneWidth As Long   'vertical split pane's width in pixels (set by graphics engine - do not use)
  splitPaneShow As Long    '0=don't show, 1=show
  gdshSplitPane As Long    'array of strings for vertical split pane (Woodies) - format: text|fontName|fontSize|boldFlag|italicFlag|textColor|textBkColor|y-value
  ptsPerBar As Double      'for squaring chart
  y_base2Move As Double    '0=base 10, >0=use to determin min label move in graphen.dll
  y_scaleMin As Double     'starting data value for y-scale specified by client
  y_scaleMax As Double     'ending data value for y-scale specified by client
  y_adjMin As Double       'adjusted for extra space & labels height (set by graphics engine)
  y_adjMax As Double       'adjusted for extra space & labels height (set by graphics engine)
  y_spAboveRatio As Double
  y_spBelowRatio As Double
  htProportion As Double   'height as proportion of chart area
End Type

Type chart_indicator
  hObj As Long
  paneId As Long
  indicatorId As Long
  indicatorType As Long
  isPriceInd As Long
  altPenStyle As Long       'for split-mode indicators
  altPenSize As Long        'for split-mode indicators
  FillColor As Long         'bollinger bar - open > close
  fillColor2 As Long        'bollinger bar - close >= open
  trueRangeColor As Long
  FillPattern As Long
  fillPct As Long
  nullValStyle As Long
  labelMode As Long         'enum type specifying where to display indicator value
  labelColor As Long        'color for label values in y-scale area
  pnfX As Long              '0=start with "O", 1=start with "X", -999=HawkeyeLevels, >1=bar handle for autoTrendline
  y_baseline As Double
  boxSize As Double         'required for PNF, Renko, also used by highlight markers & profile charts
  reversal As Double        'required for PNF & Kagi, also used by histogram & profile charts
  prevBarHigh As Double     'required for Kagi, also used by split-mode indicators
  prevBarLow As Double      'required for Kagi, also used by split-mode indicators
  TrueRangeFlag As Double   'kNullValue = don't do true range, else = close of bar prior to first bar on chart
  glhImageLoc As Long       'array of longs for highlight markers location (0=above,1=below)
  glhImageDir As Long
  glhImageFill As Long
  glhImageColor As Long
  glhImageSize As Long
  glhPenColor As Long
  glhPenStyle As Long
  glhPenWidth As Long
  glhFlags As Long          '3/9/2010: array of longs currently used only for Wyckoff PNF
  gdhData1 As Long
  gdhData2 As Long
  gdhData3 As Long
  gdhData4 As Long
  gdhYScaleVal As Long
  gdshImage As Long         'array of strings for highlight markers image
End Type

Type chart_annotation
  hObj As Long
  paneId As Long
  annotationId As Long
  PreIndicator As Long
  handleShow As Long
  moveable As Long          '0=user still selecting points, 1=points selection complete
  indicatorId As Long
  showAxes As Long
  showQtrLines As Long
  minorAxisLenData As Long  '0=ratio, 1=points, 2=ticks
  skipHitTest As Long       '0=test for hit, 1=don't test for hit
  reserved As Long
  lenRatio As Double        'length of minor axis as ratio of major axis
  lenPoints As Double       'length of minor axis as points or points per bar of gann lines
  glhAnnType As Long
  glhPenColor As Long
  glhPenStyle As Long
  glhPenSize As Long
  glhUnderline As Long
  glhFillColor As Long
  glhFillPattern As Long
  glhStyle As Long
  glhAlign As Long
  glhJustify As Long
  glhDirection As Long
  glhSize As Long
  glhImageType As Long
  glhPixTop As Long
  glhPixLeft As Long
  glhPixBottom As Long
  glhPixRight As Long
  gdhTop As Long
  gdhLeft As Long
  gdhBottom As Long
  gdhRight As Long
  gdhMisc As Long
  gshText As Long
  gshFont As Long
End Type

'/*
' * The coordFlag field specifies how to interpret the x_value, y_value fields.
' *
' * 0 --> x_value is data value for x-scale, e.g date
' *       y_value is data value for y-scale, e.g. price
' *
' * 1 --> x_value is data point number
' *       y_value is data value for y-scale, e.g. price
' *
' * 2 --> x_value, y_value are screen coordinates (valid only for HITTEST_INFO structure)
' */
Type coordinate_info
  paneId As Long
  x_pixels As Long
  y_pixels As Long
  reserved As Long
  x_value As Double
  y_value As Double
End Type

Type hittest_info
  paneId As Long
  topPaneId As Long
  btmPaneId As Long
  ItemID As Long
  itemType As Long
  annType As Long
  itemIndex As Long
  location As Long
  reserved As Long
  x_pixels As Long
  y_pixels As Long
End Type

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'glhItemShow:  values should be 1 or 0 to show/hide item (default = 0)
'               [0] TPO 1
'               [1] TPO 2  (place holder to match color array)
'               [2] volume (aka vol at price)
'               [3] TPO POC
'               [4] volume POC
'               [5] TPO VA
'               [6] volume VA
'               [7] open/close
'
'glhItemColor   [0] TPO color 1
'               [1] TPO color 2
'               [2] volume color
'               [3] TPO POC color
'               [4] volume POC color
'               [5] TPO VA color
'               [6] volume VA color
'               [7] open/close
'
'gdhItemParm:   [0] ignored - exists to keep aligned with other arrays
'               [1] ignored - exists to keep aligned with other arrays
'               [2] ignored - exists to keep aligned with other arrays
'               [3] ignored - exists to keep aligned with other arrays
'               [4] ignored - exists to keep aligned with other arrays
'               [5] TPO VA percent          'default = 70%
'               [6] volume VA percent       'default = 70%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Type profile_info
    styleTPO As Long        'alphanumeric, classic, blocks (0,1,2)
    ColorScheme As Long     '0=gradient, 1=rainbow, 2=openClose, 3=bidAsk
    RTOnOff As Long         '0=off, 1=on
    reserved3 As Long       'usage TBD
    glhItemShow As Long         'gdArrayL handle for show/hide items
    glhItemColor As Long        'gdArrayL handle for item color
    gdhItemParm As Long         'gdArrayD handle for stats parameters
    gdhXDateCopy As Long        'handle to a copy of aXDate array that gets "chopped off" by grapheng as needed
    gdhreserved1 As Long        'usage TBD: reserved gdArrayD handle
    gdhreserved2 As Long        'usage TBD: reserved gdArrayD handle
End Type

Public Enum ePCStruct           'this is tied to grapheng dll order must be changed both here & in DLL
    ePCStruct_TPO = 0           'gradient color from OR up color
    ePCStruct_TPO_ColorTo
    ePCStruct_Volume
    ePCStruct_TPO_POC
    ePCStruct_Volume_POC
    ePCStruct_TPO_VA
    ePCStruct_Volume_VA
    ePCStruct_Open
    ePCStruct_Close
End Enum

Public Enum eGEQuoteItem
    eGEQuoteItem_Nothing = 0
    eGEQuoteItem_Cell
    eGEQuoteItem_Symbol
    eGEQuoteItem_Graphics
    eGEQuoteItem_lblOpen        'OHLCPD labels
    eGEQuoteItem_lblHigh
    eGEQuoteItem_lblLow
    eGEQuoteItem_lblClose
    eGEQuoteItem_lblPrevClose
    eGEQuoteItem_lblDelta
    eGEQuoteItem_txtOpen        'OHLCPD text
    eGEQuoteItem_txtHigh
    eGEQuoteItem_txtLow
    eGEQuoteItem_txtClose
    eGEQuoteItem_txtPrevClose
    eGEQuoteItem_txtDelta
    eGEQuoteItem_txtDateTime
    eGEQuoteItem_iconBell
End Enum

Public Enum eGEQbExtraInfo
    eGEQbExtraInfo_None = 0
    eGEQbExtraInfo_TimeStamp
    eGEQbExtraInfo_BidAsk
    eGEQbExtraInfo_All
End Enum

Type quote_win
    FgColor As Long
    bkColor As Long
    HighlightColor As Long
    boldColor As Long
    nCompactQB As Long          '0=OHLC label in all cells, 1=OHLC label on left side only
    netUpColor As Long
    netDownColor As Long
    netUnchColor As Long
    thermoOpenColor As Long
    ScrollBarColor As Long      'color for little square where vert & horz scrolls intersect
    nShowExtraInfo As Long      'see eGEQbExtraInfo enum
    nForexUpDownBk As Long      '0=use up/down color for text, 1=use up/down color for background
    qwinStyle As eGDQuoteStyle
    fStyle As Long              'font style: 0=regular,1=bold,2=italic,3=bold italic
    fUnderline As Long          '0=no, 1=yes
    fSize As Long               'font size in points
    fWidth As Long              'font's metric width (set by graphics engine)
    fHeight As Long             'font's metric height (set by graphics engine)
    maxRow As Long              'total number of rows (set by graphics engine)
    maxCol As Long              'total number of columns (set by graphics engine)
    firstVisRow As Long         'first visible row (set to scroll)
    firstVisCol As Long         'first visible col (set to scroll)
    lastVisRow As Long          'last visible row (set by graphics engine)
    lastVisCol As Long          'last visible column (set by graphics engine)
    LastDataRow As Long         'set by graphics engine
    LastDataCol As Long         'set by graphics engine
    glhTextColors As Long       'forex: [0]=outer border,[1]=label border,[2]=middle,[3]=label text
    glhHorzScroll As Long       'gdArray of long (top, left, bottom, right)
    glhVertScroll As Long       'gdArray of long (top, left, bottom, right)
    glhWinDim As Long           'gdArray of long (top, left, bottom, right)
    gshTabAlerts As Long        'gdArray of string for tab alerts
    gshfName As Long            'font name as gdString
End Type

Type quote_cell
    CellID As Long          'unique identifier
    Col As Long             'set by graphics engine
    Row As Long             'set by graphics engine
    SymbolID As Long
    IsHighlighted As Long
    IsBolded As Long
    iBoldTextIdx As Long    'flag for using pen size 2 to "bold" cell's rectangle when price is at high or low
    gshSymbol As Long       'gdString containing symbol name
    gshOpen As Long         'gdString containing formated open pric
    gshHigh As Long
    gshLow As Long
    gshBid As Long
    gshLast As Long
    gshNetChange As Long
    gshTickTime As Long
    glhTextColors As Long   '[0]=symbol,[1]..[6]=OHLCPD,[7]=time stamp
    dOpen As Double
    dHigh As Double
    dLow As Double
    dLast As Double         'current value
    dPrevOpen As Double
    dPrevHigh As Double
    dPrevLow As Double
    dPrevClose As Double
End Type

Type cell_hittest_info
    CellID As Long
    Col As Long
    Row As Long
    Item As eGEQuoteItem
End Type

Type chart_order
    hObj As Long            'to be set by graphics engine only
    nOrderId As Long        'from order object
    nLong As Long           '0=short, 1=long
    nEntry As Long          '0=exit, 1=entry
    nQty As Long            'from order object
    nStatus As Long         'enumerated order status
    nOrderType As Long      'enumerated order type
    nSize As Long           'same usage as trade triangle size in annots (-60 is default)
    nNodeLevel As Long      'to be set by graphics engine only
    nHorzLine As Long       '0=don't draw, 1=draw (extended horizontal line for order)
    nPixTop As Long         'to be set by graphics engine only
    nPixLeft As Long        'these are pix coords for tool tips
    nPixBottom As Long
    nPixRight As Long
    nPixNoseX As Long
    nPixNoseY As Long
    dPrice As Double        'from order object
    glhTriggerPId As Long   'gdarray of long for IDs of parent(s) trigger order(s)
    glhTriggerCId As Long   'gdarray of long for IDs of child(ren) trigger order(s)
    glhCancelId As Long     'gdarray of long of OCO Id(s)
    gshExtra As Long        'array of string: [chart symbol][order symbol][brokerID]
End Type

Type chart_order_spec
    nLongColor As Long
    nShortColor As Long
    nEntryFill As Long      '0=hollow, 1=solid
    nOtherColor As Long     'for non-working orders, eg Parked orders
End Type

Type picbox_struct
    pbWidth As Long
    pbHeight As Long
    iconWidth As Long
    iconHeight As Long
    iLeft As Long
    iTop As Long
End Type

Declare Function GetPixel& Lib "gdi32" (ByVal hDC&, ByVal X&, ByVal Y&)
Declare Function SetPixel& Lib "gdi32" (ByVal hDC&, ByVal X&, ByVal Y&, ByVal Color&)
Declare Function geAnnotCursorPos& Lib "grapheng.dll" (ByVal obj&, ByVal hWnd&, nX&, nY&)
Declare Function geOverlayICO& Lib "grapheng.dll" (ByVal hDC&, ByVal imgBk&, ByVal imgIcon&, ByVal pbWd&, ByVal pbHt&, _
    ByVal icoWd&, ByVal icoHt&, ByVal iTextColor&, ByVal strCaption$, Optional ByVal X& = -1, Optional ByVal Y& = -1)
Declare Function geOverlayICOAll& Lib "grapheng.dll" (ByVal hDC&, ByVal imgBk&, ByVal imgIcon&, ByVal iAlign&, ByVal iTextColor&, _
    ByVal strCaption$, ptrBtnPrev As picbox_struct, ptrBtn As picbox_struct)
    
Declare Sub geShutdownAll Lib "grapheng.dll" ()
Declare Function geInitChart& Lib "grapheng.dll" (winstruct As chart_win, ByVal fnCallback&, ByVal hDC&)
Declare Function geAddItem& Lib "grapheng.dll" (ByVal obj&, ByVal itemType&, Item As Any)
Declare Function geCloseChart& Lib "grapheng.dll" (ByVal obj&)
Declare Function geCalcArcHeight& Lib "grapheng.dll" (AnnotStruct As Any, ByVal dY As Double)
Declare Function geClosestIndicatorIdx& Lib "grapheng.dll" (ByVal obj&, ByVal dX#, ByVal dyPix#, ByVal dxIsDate&, ByVal nIgnoreHzIndicator&, nIndIdx&)
Declare Function geCoordToData& Lib "grapheng.dll" (ByVal obj&, coordInfo As coordinate_info)
Declare Function geDeviceCapsPixX& Lib "grapheng.dll" (ByVal hDC&)
Declare Function geDataToCoord& Lib "grapheng.dll" (ByVal obj&, coordInfo As coordinate_info)
Declare Function geDragSeparator& Lib "grapheng.dll" (ByVal obj&, ByVal hWnd&, ByVal hDC&, ByVal xpixel&, ByVal ypixel&)
Declare Function geDrawSeparator& Lib "grapheng.dll" (ByVal obj&, ByVal hDC&, ByVal paneId&, ByVal sepLocation&, ByVal nHide&)
Declare Function geDrawWindow& Lib "grapheng.dll" (ByVal obj&, ByVal hWnd&, ByVal hDC&)
Declare Function geDrawZoomChart& Lib "grapheng.dll" (ByVal obj&, ByVal hWnd&, ByVal hDC&, ByVal Top&, ByVal Left&, ByVal Bottom&, ByVal Right&)
Declare Function geDrawZoomRect& Lib "grapheng.dll" (ByVal obj&, ByVal hDC&, ByVal Top&, ByVal Left&, ByVal Bottom&, ByVal Right&)
Declare Function geHitTest& Lib "grapheng.dll" (ByVal obj&, hitInfo As hittest_info, ByVal hDC&)
Declare Function geIncDecMinorAxis& Lib "grapheng.dll" (AnnotStruct As Any, ByVal nPixX&, ByVal nPixY&)
Declare Function gePrintChart& Lib "grapheng.dll" (ByVal obj&, ByVal screenDC&, ByVal printerDC&, ByVal bPreview&)
Declare Function geRecalcChart& Lib "grapheng.dll" (ByVal obj&, winstruct As chart_win, ByVal hDC&)
Declare Function geRemoveItem& Lib "grapheng.dll" (ByVal obj&, ByVal itemType&, Item As Any)
Declare Function geRemovePanes& Lib "grapheng.dll" (ByVal obj&)
Declare Function geRepaintTimer& Lib "grapheng.dll" (ByVal obj&, ByVal hWnd&, ByVal hDC&, ByVal nBkColor&, ByVal strText$)
Declare Function geTrendValueY# Lib "grapheng.dll" (ByVal obj&, ByVal dPointNum#)
Declare Function geSaveChart& Lib "grapheng.dll" (ByVal obj&, ByVal hWnd&, ByVal hDC&, ByVal nWidth&, ByVal nHeight&, ByVal nFormat&, ByVal strFile As String, Optional nNewPixX&, Optional nNewDpointsCount&)
Declare Function geUnzoomChart& Lib "grapheng.dll" (ByVal obj&, ByVal hDC&)
Declare Function geZoomPixPerBar& Lib "grapheng.dll" (ByVal obj&, ByVal hDC&, ByVal nPix&, ByVal nAction&)
Declare Function geDrawMarker& Lib "grapheng.dll" (ByVal hDC&, ByVal nPixX&, ByVal nPixY&, ByVal nRadius&, Optional ByVal nPenSize& = 1, Optional ByVal nPenColor& = 0, Optional ByVal nPenStyle& = 0, Optional ByVal nShape& = 0)
Declare Function geGetTextDimension& Lib "grapheng.dll" (ByVal hDC&, AnnotStruct As chart_annotation, ByVal nIdx&, nPixWidth&, nPixHeight&)
Declare Sub geAnnotMove Lib "grapheng.dll" (ByVal obj&, ByVal nAnnotMoving&)
Declare Function geValidTriangle& Lib "grapheng.dll" (ByVal obj&)
'01-16-2008: geSyncCrossHair is obsolete, leave awhile for backwards compatibility then remove
Declare Function geSyncCrossHair& Lib "grapheng.dll" (ByVal obj&, ByVal hWnd&, ByVal hDC&, ByVal dDate#, ByVal dX#, ByVal dY#, ByVal nVert&, ByVal nHorz&)
Declare Function geSyncCrossHairEx& Lib "grapheng.dll" (ByVal obj&, ByVal hWnd&, ByVal hDC&, ByVal dDate#, ByVal dX#, ByVal dY#, ByVal nVert&, ByVal nHorz&, ByVal nDrawFatBar&, Optional ByVal nSplitPane& = 0)
Declare Function geAnnotAtVal& Lib "grapheng.dll" (ByVal obj&, ByVal nAnnotType&, ByVal nPaneID&, ByVal nPix&, ByVal dX#, ByVal dY#, ByVal strReturn As Long)

Declare Sub geAppBkDotsSpec Lib "grapheng.dll" (ByVal nDotSize&, ByVal nDotColor&, ByVal nPixBetweenDots&)
Declare Function geLoadAppBkBitmap& Lib "grapheng.dll" (ByVal hWnd&, ByVal nDotColor&, ByVal nLogoColor&, ByVal nDotSize&, ByVal nLogoSize&, ByVal nPixBetweenDots&, ByVal nPixWidth&, ByVal nPixHeight&, ByVal strFile$)
Declare Function geSnapToDots& Lib "grapheng.dll" (ByVal hbmp&, ByVal hwndDots&, ByVal hwndSnapToDots&)
Declare Sub geSextantFile Lib "grapheng.dll" (ByVal obj&, ByVal strSextantFile As String)
Declare Sub geWaterMarkFile Lib "grapheng.dll" (ByVal obj&, ByVal strWaterMarkFile As String)
Declare Function geHighContrastOn Lib "grapheng.dll" (ByVal hWnd&, ByVal bkColor&, ByVal ForeColor&) As Long

Declare Sub geZoomModeAuto Lib "grapheng.dll" (ByVal obj&)
Declare Sub geZoomModeManual Lib "grapheng.dll" (ByVal obj&)
Declare Sub geSetOrderSpec Lib "grapheng.dll" (ByVal obj&, OrderSpec As chart_order_spec)
Declare Function gePeekClickMsg& Lib "grapheng.dll" (ByVal hWnd&)
Declare Function geWrapText& Lib "grapheng.dll" (ByVal hDC&, ByVal strText As String, ByVal strUseTextWidth As String)
Declare Function geCaptureScreen& Lib "grapheng.dll" (ByVal obj&, ByVal hWnd&, ByVal strText As String, ByVal nLeft&, ByVal nTop&, ByVal nWidth&, ByVal nHeight&, ByVal nChartOnly&)
Declare Function geGetBlueValue& Lib "grapheng.dll" (ByVal nColor&)
Declare Function geGetRedValue& Lib "grapheng.dll" (ByVal nColor&)
Declare Function geGetGreenValue& Lib "grapheng.dll" (ByVal nColor&)
Declare Function geIsCursorInWnd& Lib "grapheng.dll" (ByVal hWnd&, ByRef iCursorX&, ByRef iCursorY&)
Declare Function geIsPointOnScreen& Lib "grapheng.dll" (ByVal nX&, ByVal nY&)
'functions below are for icon palette 'load' icon
Declare Function geDrawText& Lib "grapheng.dll" (ByVal hDC&, ByVal X&, ByVal Y&, ByVal nTextColor&, ByVal strText As String)
Declare Function geTriangleDown& Lib "grapheng.dll" (ByVal hDC&, ByVal nColor&)
Declare Function geTriangleUp& Lib "grapheng.dll" (ByVal hDC&, ByVal nColor&)
Declare Function geDiamond& Lib "grapheng.dll" (ByVal hDC&, ByVal nColor&)
Declare Function geFootprintIcon& Lib "grapheng.dll" (ByVal hDC&, ByVal Icon&, ByVal bkColor&, ByVal IconColor&, ByVal strFile As String)

'functions below are for quote cells
Declare Function geInitQuoteWin& Lib "grapheng.dll" (qwinstruct As quote_win, ByVal hDC&)
Declare Function geCloseQuoteWin& Lib "grapheng.dll" (ByVal obj&)
Declare Function geDrawQuoteCell& Lib "grapheng.dll" (ByVal obj&, CellStruct As Any, ByVal hWnd&, ByVal hDC&)
Declare Function geDrawQuoteWin& Lib "grapheng.dll" (ByVal obj&, ByVal hWnd&, ByVal hDC&)
Declare Function geHitTestCell& Lib "grapheng.dll" (ByVal obj&, ByVal hDC&, cellInfo As cell_hittest_info, ByVal eLastItem As eGEQuoteItem, ByVal xPix&, ByVal yPix&)
Declare Function geUpdateCell& Lib "grapheng.dll" (ByVal obj&, qcell As quote_cell)
Declare Function gePrintQuoteWin& Lib "grapheng.dll" (ByVal obj&, ByVal hWnd&, ByVal hDC&, ByVal nFirstRow&, ByVal nFirstCol&, ByVal nLastRow&, ByVal nLastCol&, ByVal strFileName$)
Declare Function geSaveQuoteWin& Lib "grapheng.dll" (ByVal obj&, ByVal hWnd&, ByVal hDC&, ByVal strFileName$, ByVal strHeader$)

Declare Sub geGetBoldCell Lib "grapheng.dll" (ByVal obj&, nRow&, nCol&)
Declare Sub geSetBoldCell Lib "grapheng.dll" (ByVal obj&, ByVal hWnd&, ByVal hDC&, ByVal nRow&, ByVal nCol&, ByVal nBold&)
Declare Sub geGetHighlightCell Lib "grapheng.dll" (ByVal obj&, nRow&, nCol&)
Declare Sub geSetHighlightCell Lib "grapheng.dll" (ByVal obj&, ByVal hWnd&, ByVal hDC&, ByVal nRow&, ByVal nCol&, ByVal nToggle&)
Declare Function geRecalcCells& Lib "grapheng.dll" (ByVal obj&, ByVal hDC&)

'Bid/Ask implemented in v5.0 build 334 - passing in 1 enables bid/ask on box style QB
Declare Sub geEnableFeature Lib "grapheng.dll" (ByVal obj&, ByVal iFeature&)

'pass profile info struct to the grapheng for profile charts
Declare Sub geSetProfileInfo Lib "grapheng.dll" (ByVal obj&, profileStruct As profile_info)

'functions below are for tick distribution
Declare Function geInitTickObj& Lib "grapheng.dll" ()
Declare Function geInitTickData& Lib "grapheng.dll" (ByVal obj&, ByVal nBarHandle&)
Declare Function geDrawTicks& Lib "grapheng.dll" (ByVal obj&, ByVal hDC&, ByVal dPrice#, ByVal nBarHandle&, ByVal nPenColor&, ByVal nTop&, ByVal nLeft&, ByVal nBottom&, ByVal nRight&)
Declare Function geTickTime& Lib "grapheng.dll" (ByVal obj&, ByVal strTime As String, ByVal nX&, ByVal nY&)
Declare Function geDrawTickTimeScale& Lib "grapheng.dll" (ByVal obj&, ByVal hDC&, ByVal strFontName As String, ByVal nFonSize&, ByVal nTop&, ByVal nLeft&, ByVal nBottom&, ByVal nRight&)
Declare Function geDrawTickTriangle& Lib "grapheng.dll" (ByVal obj&, ByVal hDC&, ByVal nPenColor&, ByVal nBrushColor&, ByVal nUpDown&, ByVal nTop&, ByVal nLeft&, ByVal nBottom&, ByVal nRight&)
Declare Function geDrawTickHistogram& Lib "grapheng.dll" (ByVal obj&, ByVal hWnd&, ByVal hDC&, ByVal aColors&, ByVal aBidAskSizes&, ByVal nSumOfSizes&, ByVal nBackgroundColor&, ByVal nTop&, ByVal nLeft&, ByVal nBottom&, ByVal Right&, ByVal nJustifyFlay&)
Declare Function geTickObjDisplayTimeZone& Lib "grapheng.dll" (ByVal obj&, ByVal strTimeZone$)
Declare Sub geTickLineDirection Lib "grapheng.dll" (ByVal obj&, ByVal nRightToLeft&)

Declare Function geTaskBarNotify& Lib "grapheng.dll" (ByVal hWnd&, ByVal strMsg As String)

Function fnFormatPrice(ByVal dValue#, ByVal hString&, _
    panestruct As chart_pane, ByVal hDC&, ByVal nIndId&) As Long
'On Error GoTo ErrSection:
On Error Resume Next:

    Static iPrevForm As Long

    Dim Chart As cChart, dMax#, szData$, i&, strFmt$, nForSessionDate&
    Dim bAsPercent As Boolean
    Dim Pane As cPane, Ind As cIndicator
    
    If hDC = 0 Then
        ' just do a non-chart specific format -- but use
        ' the specified format if passed
        If gdGetSize(hString) > 0 Then
            gdSetStr hString, 0, Format(dValue, gdGetStr(hString))
        Else
            gdSetStr hString, 0, CStr(dValue)
        End If
        Exit Function
    End If
    
    'initialize return string to blank
    gdSetStr hString, 0, ""
    
    ' abbreviate very large numbers
    dMax = Abs(panestruct.y_scaleMax)
    If dMax < Abs(panestruct.y_scaleMin) Then
        dMax = Abs(panestruct.y_scaleMin)
    End If
    If dMax < Abs(dValue) Then
        dMax = Abs(dValue)
    End If
    If dMax >= 100000# Then  '(so $DJIA prices show normal)
        dValue = Int(dValue)
        If nIndId > 0 Then dMax = Abs(dValue)
        If dValue = 0 Then
            szData = "0"
        ElseIf dMax >= 1000000000000# Then
            szData = Format(dValue / 1000000000000#, "0.0##") & "T"
        ElseIf dMax >= 100000000000# Then
            szData = Format(dValue / 1000000000#, "0.0#") & "B"
        ElseIf dMax >= 1000000000# Then
            szData = Format(dValue / 1000000000#, "0.0##") & "B"
        ElseIf dMax >= 100000000# Then
            szData = Format(dValue / 1000000#, "0.0#") & "M"
        ElseIf dMax >= 1000000# Then
            szData = Format(dValue / 1000000#, "0.0##") & "M"
        Else
            szData = CStr(dValue)
        End If
        gdSetStr hString, 0, Trim(szData)
        Exit Function
    End If
    
    'find the chart object that goes with this form/hdc
    Set Chart = Nothing
    i = iPrevForm
    If i > 0 And i < Forms.Count Then
        If IsFrmChart(Forms(i)) Then
            If Forms(i).pbChart.hDC = hDC Then
                Set Chart = Forms(i).Chart
            End If
        End If
    End If
    If Chart Is Nothing Then
        For i = 0 To Forms.Count - 1
            If IsFrmChart(Forms(i)) Then
                If Forms(i).pbChart.hDC = hDC Then
                    Set Chart = Forms(i).Chart
                    iPrevForm = i
                    Exit For
                End If
            End If
        Next
    End If
    
    'don't have chart object, can't do anything
    If Chart Is Nothing Then Exit Function
    Set Pane = Chart.Tree(panestruct.paneId)
    'don't have pane object, can't do anything
    If Pane Is Nothing Then
        Set Chart = Nothing
        Exit Function
    End If
    
    'if this is seasonal chart with the Trends overlayed then return empty string
    If Chart.TypeOfChart = eTypeChart_Seasonal Then
        If Not Chart.Tree Is Nothing Then
            If TypeOf Chart.Tree(kSeasonalAvgIndKey) Is cIndicator Then
                Set Ind = Chart.Tree(kSeasonalAvgIndKey)
                If Ind.Overlayed Then Exit Function
            End If
        End If
    End If
    
    If Pane.PaneLogFlag = ePANE_LogFlagPercent Or Chart.TypeOfChart = eTypeChart_Seasonal Then
        bAsPercent = True
    End If
    
    If Pane.DisplayFormat = ePANE_PriceFormat And Not bAsPercent Then
        ' TLB 12/4/2009: for now, we'll just use the right-most visible screen date as the
        ' default to determine the price display (e.g. when min move for Bonds changed in 2009)
        i = Chart.LastGoodDataBar(False, True)
        nForSessionDate = Chart.Bars.SessionDate(i)
        If Chart.Tree.NodeLevel(nIndId) > 0 Then Set Ind = Chart.Tree(nIndId)
        If Ind Is Nothing Then
            szData = Chart.PriceDisplay(panestruct.paneId, dValue, nForSessionDate)
        ElseIf Ind.DataType = eINDIC_BarData Then
            'display value appropriate for this specific bar's price data format
            szData = Ind.Bars.PriceDisplay(dValue, True, nForSessionDate)     '4349
        Else
            'display value in format of this chart's primary PRICE bar
            szData = Chart.PriceDisplay(panestruct.paneId, dValue, nForSessionDate)
        End If
    ElseIf Pane.DisplayFormat = ePANE_RoundDecimals Then
        If Pane.DisplayDecimals > 0 Then
            strFmt = "0." & String(Pane.DisplayDecimals, "0")
        Else
            strFmt = "0"
        End If
        szData = Format(dValue, strFmt)
    Else
        If dValue = Int(dValue) Then
            strFmt = "0"
        ElseIf Abs(dValue) > 10 Then
            strFmt = "0.00"
        ElseIf Abs(dValue) > 0.01 Then
            strFmt = "0.0000"
        Else
            strFmt = "0.000000"
        End If
        If bAsPercent And Len(strFmt) > 1 Then
            strFmt = Left(strFmt, Len(strFmt) - 1)
        End If
        szData = Format(dValue, strFmt)
    End If
    
    If bAsPercent Then
        szData = szData & "%"
    End If
            
    gdSetStr hString, 0, Trim(szData)
    
ErrExit:
    Set Chart = Nothing
    Set Pane = Nothing
    Set Ind = Nothing
    Exit Function
    
ErrSection:
    Set Chart = Nothing
    Set Pane = Nothing
    Set Ind = Nothing
    RaiseError "mgraphengDll.fnFormatPrice"
    
End Function
