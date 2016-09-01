VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Begin VB.Form frmBidAskDir 
   Caption         =   "Bid/Ask Directional Analysis"
   ClientHeight    =   6030
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7605
   LinkTopic       =   "Form1"
   ScaleHeight     =   6030
   ScaleWidth      =   7605
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox IconPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   7035
      ScaleHeight     =   33
      ScaleMode       =   0  'User
      ScaleWidth      =   17
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin gdOCX.gdSelectColor gdColor 
      Height          =   375
      Left            =   5040
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      CustomColor     =   255
   End
   Begin VSFlex7LCtl.VSFlexGrid fgBidAskDir 
      Height          =   2055
      Left            =   720
      TabIndex        =   0
      Top             =   840
      Width           =   3855
      _cx             =   6800
      _cy             =   3625
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
   Begin VB.Timer tmr 
      Interval        =   500
      Left            =   5880
      Top             =   2880
   End
   Begin ActiveToolBars.SSActiveToolBars tbToolbar 
      Left            =   5520
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131083
      ToolBarsCount   =   1
      ToolsCount      =   7
      DisplayContextMenu=   0   'False
      Tools           =   "frmBidAskDir.frx":0000
      ToolBars        =   "frmBidAskDir.frx":278A
   End
   Begin VSFlex7LCtl.VSFlexGrid fgSummary 
      Height          =   2055
      Left            =   720
      TabIndex        =   2
      Top             =   3240
      Width           =   3855
      _cx             =   6800
      _cy             =   3625
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
End
Attribute VB_Name = "frmBidAskDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const kCaptionBase = "AMPT Bid/Ask Directional Analysis"

Private Enum eBidAskDir_StatsView
    eTSVView_StatsAll = 0
    eTSVView_StatsSummary
    eTSVView_StatsNone
End Enum

Private Type mPrivate
    Data As cTSVData
    
    eStatsView As eBidAskDir_StatsView
    
    nSymID As Long
    strSym As String
       
    iTickR As Long
    iBigLot As Long
    iHiLightTrades As Long
    
    iTopRow As Long
    iLeftCol As Long
    iLastPriceRow As Long
    iCurrPriceColor As Long
    iPrevPriceSwitchRow As Long     'previously colored & bolded price row for differential switch
    iPriceColLocation As Long       '0=left, 1=right, 2=both
    
    iMouseColDown As Long
    iMouseRowDown As Long
    
    iMouseColDownPrev As Long
    iMouseRowDownPrev As Long
    
    iBoxColor As Long
    iBoxPix As Long
    iBoxFill As Long
    
    iColorAsk As Long
    iColorBid As Long
    iColorMean As Long
    iColorValue As Long
    
    iFgSummaryHeight As Long
    iBlankRows As Long
    
    bInitInprog As Boolean
    bTimerInProg As Boolean
    bReloadData As Boolean
    bEditDrawInProg As Boolean
    
    bKeepAtEnd As Boolean
    bCenterPrice As Boolean

    dLastUpdated As Double
End Type

Private m As mPrivate

Public Sub ShowMe(ByVal strSym$)
On Error GoTo ErrSection:
        
    Dim strText$, bCenter As Boolean
           
    tmr.Enabled = False
    m.bKeepAtEnd = True
    
    InitGrid
    LoadNewSymData strSym

    'Restore/set form size & location
    strText = GetIniFileProperty(Me.Name, "", "Placement", g.strIniFile)
    If strText = "" Then
        CenterTheForm Me
    Else
        SetFormPlacement Me, strText
    End If

    Me.Width = Int(Me.Width / 800) * 800
    ShowForm Me, eForm_Nonmodal, frmMain
        
    bCenter = GetIniFileProperty("BidAskDirCenterPrice", True, "IOAMT", g.strIniFile)
    If bCenter = False Then
        m.bCenterPrice = True   'always center on price when first shown
        FocusGrid
    End If
    
    m.bCenterPrice = bCenter
    tbToolbar.Tools("ID_CenterPrice").State = Abs(m.bCenterPrice)
        
    tmr.Enabled = g.RealTime.Active
    
    Exit Sub
        
ErrSection:
    RaiseError "frmBidAskDir.ShowMe"
    
End Sub

Private Sub InitGrid()
On Error GoTo ErrSection:
        
    m.bInitInprog = True
    
    m.iFgSummaryHeight = 0
    With fgSummary
        .Redraw = flexRDNone
        .FixedCols = 1
        .FixedRows = 0
        .Editable = flexEDNone
        .HighLight = flexHighlightNever
        .ExtendLastCol = True
        .BorderStyle = flexBorderNone
        .ScrollBars = flexScrollBarNone
        
        .Rows = 0
        .Cols = .FixedCols
        
        .Redraw = flexRDBuffered
    End With
    
    With fgBidAskDir
        .Redraw = flexRDNone
        .FixedCols = 1
        .FixedRows = 1
        .Editable = flexEDNone
        .HighLight = flexHighlightNever
        .ExtendLastCol = True
        .BorderStyle = flexBorderNone
        .ScrollBars = flexScrollBarBoth
        
        .Rows = .FixedRows
        .Cols = .FixedCols
        
        .MergeRow(0) = True
        .MergeCells = flexMergeFree
                                        
        .Redraw = flexRDBuffered
    End With
    
    m.bInitInprog = False
    
    Exit Sub
        
ErrSection:
    RaiseError "frmBidAskDir.InitGrid"
    
End Sub


Private Sub fgBidAskDir_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
On Error Resume Next

    Dim iCols&, iCols2&

    If g.bUnloading Then Exit Sub

    If Not m.bInitInprog Then
        m.iTopRow = NewTopRow
        m.iLeftCol = NewLeftCol
        m.bKeepAtEnd = fgBidAskDir.ColIsVisible(fgBidAskDir.Cols - 2)
        fgSummary.LeftCol = NewLeftCol
    End If
        
End Sub

Private Sub fgBidAskDir_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    Dim strText$, strKey$, strIconType$
    Dim iIndex&
    
    If m.bEditDrawInProg Then Exit Sub
    
    If tbToolbar.Tools("ID_Draw").State = ssChecked Then
        With fgBidAskDir
            m.iMouseColDownPrev = m.iMouseColDown
            m.iMouseColDown = .MouseCol
            
            m.iMouseRowDownPrev = m.iMouseRowDown
            m.iMouseRowDown = .MouseRow
            
            If .MouseCol = 0 And InRangeRow(.MouseRow) Then
                If Button = vbLeftButton Then
                    gdColor.Move X, Y
                    gdColor.Visible = True
                Else
                    HandleHighlight True
                End If
            End If
        End With
    ElseIf tbToolbar.Tools("ID_Icons").State = ssChecked Then
        With fgBidAskDir
            If Button = vbLeftButton Then
                If InRangeRow(.Row) And InRangeCol(.Col) Then
                    .Cell(flexcpPicture, .Row, .Col) = frmFootprintIcons.SelectedPic.Picture
                    .Cell(flexcpPictureAlignment, .Row, .Col) = frmFootprintIcons.IconAlign
                    
                    iIndex = frmFootprintIcons.SelectedPicIndex
                    strIconType = frmFootprintIcons.IconTypeStr(iIndex)
                    strKey = strIconType & Str(m.Data.IconCount + 1)
                    strText = "ICON|" & strKey & "|" & strIconType & "|" & .TextMatrix(.Row, 0) & "|" & Str(.Col) & "|" & Str(frmFootprintIcons.IconColor) & "|" & Str(frmFootprintIcons.IconAlign)
                    
                    .Cell(flexcpData, .Row, .Col) = strKey
                    
                    m.Data.AddIcon strText
                End If
            ElseIf InRangeRow(.MouseRow) And InRangeCol(.MouseCol) Then
                .Cell(flexcpPicture, .MouseRow, .MouseCol) = Nothing
                strKey = .Cell(flexcpData, .MouseRow, .MouseCol)
                m.Data.RemoveIcon strKey
            End If
        End With
    End If

End Sub

Private Sub fgBidAskDir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    If tbToolbar.Tools("ID_Draw").State = ssChecked Or tbToolbar.Tools("ID_Icons").State = ssChecked Then
        With fgBidAskDir
            If InRangeRow(.MouseRow) Then
                .MousePointer = flexCustom
                .MouseIcon = Picture16(ToolbarIcon("kPencil"))
            Else
                .MousePointer = flexDefault
            End If
        End With
    Else
        fgBidAskDir.MousePointer = flexDefault
        ResetMouseVar
    End If

End Sub

Private Sub fgBidAskDir_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    
    If tbToolbar.Tools("ID_Draw").State = ssChecked Then
        HandleBox Button
    Else
        ResetMouseVar
    End If

End Sub


Private Sub Form_Load()
On Error GoTo ErrSection:

    Me.Icon = Picture16("kBlank")
    
    g.Styler.StyleForm Me

    m.iTickR = GetIniFileProperty("RevTick", 4, "IOAMT", g.strIniFile)
    m.iBigLot = GetIniFileProperty("BigLot", 500, "IOAMT", g.strIniFile)
    m.iHiLightTrades = GetIniFileProperty("HiLightTrades", 10000, "IOAMT", g.strIniFile)
    m.iBlankRows = GetIniFileProperty("BidAskDirBlankRows", 10, "IOAMT", g.strIniFile)
    m.iColorAsk = GetIniFileProperty("ColorAsk", kAskColor, "IOAMT", g.strIniFile)
    m.iColorBid = GetIniFileProperty("ColorBid", kBidColor, "IOAMT", g.strIniFile)
    m.iColorMean = GetIniFileProperty("ColorMean", vbGreen, "IOAMT", g.strIniFile)
    m.iColorValue = GetIniFileProperty("ColorValue", vbYellow, "IOAMT", g.strIniFile)
    m.eStatsView = GetIniFileProperty("StatsView", eTSVView_StatsSummary, "IOAMT", g.strIniFile)
    m.iPriceColLocation = GetIniFileProperty("PriceColLocation", 0, "IOAMT", g.strIniFile)
    
    With tbToolbar
        .Tools("ID_Symbol").Picture = Picture16(ToolbarIcon("ID_Symbol"))
        .Tools("ID_Close").Picture = Picture16(ToolbarIcon("kCancel"))
        .Tools("ID_Settings").Picture = Picture16(ToolbarIcon("ID_Settings"))   'want new toolbar to use kSettings for consistency
        .Tools("ID_StatsView").ComboBox.ListIndex = m.eStatsView
    End With
        
    m.iBoxColor = vbBlue
    m.iBoxPix = 1
    
    Exit Sub

ErrSection:
    RaiseError "frmBidAskDir.Form_Load"
    
End Sub

Private Sub Form_Resize()
On Error Resume Next
    
    Dim iSummaryHt&, iBidAskHt&, i&

    If Me.WindowState = vbMinimized Then
        Exit Sub
    End If
    
    If fgSummary.Visible Then
        fgSummary.Redraw = flexRDNone
        With fgBidAskDir
            .Redraw = flexRDNone
            iBidAskHt = .Rows * .RowHeight(1)
            If Me.ScaleHeight >= iBidAskHt + m.iFgSummaryHeight Then
                .Move 0, 0, Me.ScaleWidth, iBidAskHt
            ElseIf Me.ScaleHeight - m.iFgSummaryHeight > 0 Then
                .Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - m.iFgSummaryHeight
            Else
                .Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
            End If
            fgSummary.Move 0, .Height, .ClientWidth, m.iFgSummaryHeight
        End With
        
        fgSummary.ColWidth(0) = 1500
        fgBidAskDir.ColWidth(0) = 1500
        fgSummary.ColWidth(fgSummary.Cols - 1) = 1500
        fgBidAskDir.ColWidth(fgBidAskDir.Cols - 1) = 1500
                
        For i = 1 To fgSummary.Cols - 2
            fgSummary.ColWidth(i) = 800
            fgBidAskDir.ColWidth(i) = 800
        Next
        
        fgSummary.LeftCol = fgBidAskDir.LeftCol
        
        fgBidAskDir.Redraw = flexRDBuffered
        fgSummary.Redraw = flexRDBuffered
    Else
        With fgBidAskDir
            .Redraw = flexRDNone
            iBidAskHt = (.Rows + 2) * .RowHeight(1)
            If .MergeRow(.Rows - 1) = True Then
                .Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
            ElseIf Me.ScaleHeight > iBidAskHt Then
                .Move 0, 0, Me.ScaleWidth, iBidAskHt
            Else
                .Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
            End If
            .Redraw = flexRDBuffered
        End With
    End If
        
End Sub

Private Sub UpdatePullback(Optional ByVal bRT As Boolean = False)
On Error GoTo ErrSection:

    Dim Table As cGdTable
    Dim dVolBid#, dVolAsk#
    Dim i&
    
    Dim bAddColRight As Boolean

    Set Table = m.Data.PriceVolPBTable
    
    'precautionary checks
    If Table Is Nothing Then
        Exit Sub
    ElseIf Table.NumRecords <> m.Data.PriceVolTable.NumRecords Then
        Exit Sub
    End If
        
    With fgBidAskDir
        If .TextMatrix(0, .Cols - 2) <> "Pullback" Then
            .Cols = .Cols + 1
            .TextMatrix(0, .Cols - 1) = "Pullback"
            .Cols = .Cols + 1
            .TextMatrix(0, .Cols - 1) = "Pullback"
                        
            .Cell(flexcpAlignment, .FixedRows, .Cols - 1, .Rows - 1, .Cols - 1) = flexAlignLeftCenter
            .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
            
            .Cell(flexcpForeColor, .FixedRows + 1, .Cols - 1, .Rows - 1, .Cols - 1) = vbBlue
            .Cell(flexcpForeColor, .FixedRows + 1, .Cols - 2, .Rows - 1, .Cols - 2) = vbRed
            
            .Cols = .Cols + 1
            bAddColRight = True
        End If
                
        For i = 2 To Table.NumRecords
            If Table(0, i) > 0 Then
                If ValOfText(.TextMatrix(i - .FixedRows, .Cols - 3)) <> Table(0, i) Then
                    .TextMatrix(i - .FixedRows, .Cols - 3) = Table(0, i)
                    If bRT Then
                        .Cell(flexcpFontBold, i - .FixedRows, .Cols - 3) = True
                    End If
                End If
            Else
                If i - .FixedRows < .Rows Then .TextMatrix(i - .FixedRows, .Cols - 3) = ""
            End If
            If Table(1, i) > 0 Then
                If ValOfText(.TextMatrix(i - .FixedRows, .Cols - 2)) <> Table(1, i) Then
                    .TextMatrix(i - .FixedRows, .Cols - 2) = Table(1, i)
                    If bRT Then
                        .Cell(flexcpFontBold, i - .FixedRows, .Cols - 2) = True
                    End If
                End If
            Else
                If i - .FixedRows < .Rows Then .TextMatrix(i - .FixedRows, .Cols - 2) = ""
            End If
            dVolBid = Table(0, i)
            dVolAsk = Table(1, i)
            If dVolBid > dVolAsk Then
                .Cell(flexcpFloodColor, i - .FixedRows, .Cols - 3) = kAskColor
                .Cell(flexcpFloodPercent, i - .FixedRows, .Cols - 3) = (dVolAsk - dVolBid) / (dVolBid + dVolAsk) * 100
            ElseIf dVolBid < dVolAsk Then
                .Cell(flexcpFloodColor, i - .FixedRows, .Cols - 2) = kBidColor
                .Cell(flexcpFloodPercent, i - .FixedRows, .Cols - 2) = (dVolAsk - dVolBid) / (dVolBid + dVolAsk) * 100
            End If
            If i - .FixedRows < .Rows Then
                If .Cell(flexcpBackColor, i - .FixedRows, .Cols - 4) <> vbYellow Then
                    .Cell(flexcpBackColor, i - .FixedRows, .Cols - 3, i - .FixedRows, .Cols - 2) = .Cell(flexcpBackColor, i - .FixedRows, .Cols - 4)
                End If
            End If
        Next
    End With
    
    Set Table = Nothing
    Set Table = m.Data.StatsPBTable
    
    If Not Table Is Nothing Then
        dVolBid = Table(0, 0)
        dVolAsk = Table(1, 0)

        With fgSummary
            .Cols = fgBidAskDir.Cols
            If dVolBid = 0 Then
                .TextMatrix(1, .Cols - 3) = ""
            Else
                .TextMatrix(1, .Cols - 3) = dVolBid
            End If
            If dVolAsk = 0 Then
                .TextMatrix(1, .Cols - 2) = ""
            Else
                .TextMatrix(1, .Cols - 2) = dVolAsk
            End If
    
            If dVolBid > dVolAsk Then
                .Cell(flexcpFloodColor, 1, .Cols - 3) = kAskColor
                .Cell(flexcpFloodPercent, 1, .Cols - 3) = (dVolAsk - dVolBid) / (dVolBid + dVolAsk) * 100
            ElseIf dVolBid < dVolAsk Then
                .Cell(flexcpFloodColor, 1, .Cols - 2) = kBidColor
                .Cell(flexcpFloodPercent, 1, .Cols - 2) = (dVolAsk - dVolBid) / (dVolBid + dVolAsk) * 100
            End If
            '[Differential]
            If Table(0, 1) = 0 Then
                .TextMatrix(2, .Cols - 3) = ""
            Else
                .TextMatrix(2, .Cols - 3) = Table(0, 1)
            End If
            If Table(1, 1) = 0 Then
                .TextMatrix(2, .Cols - 2) = ""
            Else
                .TextMatrix(2, .Cols - 2) = Table(1, 1)
            End If
            '[Big Lots]
            If Table(0, 2) = 0 Then
                .TextMatrix(5, .Cols - 3) = ""
            Else
                .TextMatrix(5, .Cols - 3) = Table(0, 2)
            End If
            If Table(1, 2) = 0 Then
                .TextMatrix(5, .Cols - 2) = ""
            Else
                .TextMatrix(5, .Cols - 2) = Table(1, 2)
            End If
        
'01-09-2007
'this was inntended to make separator rows thicker, but is intermittently broken
'does not seem to make much difference so comment out for now and fix later
'            .Cell(flexcpBackColor, 0, 0, 0, .Cols - 1) = RGB(128, 128, 128)
'            .Cell(flexcpBackColor, 4, 0, 4, .Cols - 1) = RGB(128, 128, 128)
'            .Cell(flexcpBackColor, 7, 0, 7, .Cols - 1) = RGB(128, 128, 128)
        
            .Cell(flexcpForeColor, 0, .Cols - 2, .Rows - 1, .Cols - 2) = vbBlue
            .Cell(flexcpForeColor, 0, .Cols - 3, .Rows - 1, .Cols - 3) = vbRed
            .Cell(flexcpAlignment, 0, .Cols - 2, .Rows - 1, .Cols - 2) = flexAlignLeftCenter
        End With
    End If
    
    If bAddColRight Then
        With fgBidAskDir
            .Cell(flexcpBackColor, 0, .Cols - 1, .Rows - 1, .Cols - 1) = .BackColorFixed
            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .Cols - 1) = .TextMatrix(i, 0)
                If .Cell(flexcpBackColor, i, 0) <> .BackColorFixed Then
                    .Cell(flexcpBackColor, i, .Cols - 1) = .Cell(flexcpBackColor, i, 0)
                End If
                .Cell(flexcpForeColor, i, .Cols - 1) = .Cell(flexcpForeColor, i, 0)
                .Cell(flexcpFontBold, i, .Cols - 1) = .Cell(flexcpFontBold, i, 0)
            Next
            .Cell(flexcpAlignment, 0, .Cols - 1, .Rows - 1, .Cols - 1) = flexAlignLeftCenter
        End With
        With fgSummary
            .Cell(flexcpBackColor, 0, .Cols - 1, .Rows - 1, .Cols - 1) = .BackColorFixed
            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .Cols - 1) = .TextMatrix(i, 0)
                .Cell(flexcpBackColor, i, .Cols - 1) = .Cell(flexcpBackColor, i, 0)
            Next
        End With
    End If
                
    Set Table = Nothing
    
    Exit Sub

ErrSection:
    RaiseError "frmBidAskDir.UpdatePullback"
    
End Sub

Private Sub HighlightTradeSpeed()
On Error GoTo ErrSection:

    Dim aSpeedCells As cGdArray
    Dim i&, iSpeedRow&, iSpeedCol&
    Dim Table As cGdTable
    
    '01-08-2007: per Rasa, don't do this for now
    Exit Sub
    
    Set Table = m.Data.PriceVolTable
    Set aSpeedCells = m.Data.SpeedCells
    
    If Not aSpeedCells Is Nothing Then      'precautionary
        With fgBidAskDir
            For i = 0 To m.Data.SpeedCells.Size - 1
                iSpeedRow = Parse(aSpeedCells(i), ",", 1) - .FixedRows
                iSpeedCol = Parse(aSpeedCells(i), ",", 2)
                If iSpeedRow >= .FixedRows And iSpeedRow < .Rows Then
                    If iSpeedCol >= .FixedCols And iSpeedCol < .Cols Then
                        If Table(0, iSpeedRow + .FixedRows) = ValOfText(.TextMatrix(iSpeedRow, 0)) Then 'precautionary
                            .Cell(flexcpBackColor, iSpeedRow, iSpeedCol) = vbYellow
                        End If
                    End If
                End If
            Next
        End With
    End If
    
    Set aSpeedCells = Nothing
    
    Exit Sub

ErrSection:
    RaiseError "frmBidAskDir.HighlightTradeSpeed"
    
End Sub

Private Sub BoldSwitchPrice()
On Error Resume Next

    'color & bold price where differential changed for pos to neg or vice versa
    'differential = trades@bid - trades@ask = sells - buys
    'differential going from pos->neg = red price (opposite = blue price)
    Dim nColor&, nBkColor&, nRow&, dPriceAtSwitch#
    
    If m.Data.TotalTradesAtBid > m.Data.TotalTradesAtAsk Then
        nColor = vbRed
        nBkColor = kAskColor
    ElseIf m.Data.TotalTradesAtBid < m.Data.TotalTradesAtAsk Then
        nColor = vbBlue
        nBkColor = kBidColor
    Else
        nColor = 0
    End If
    
    dPriceAtSwitch = m.Data.PriceAtSwitch
    nRow = m.Data.TbRowForPrice(dPriceAtSwitch)
    
    With fgBidAskDir
        'clear previously colored & bolded price
        If m.iPrevPriceSwitchRow > .FixedRows And m.iPrevPriceSwitchRow <> m.iLastPriceRow Then
            .Cell(flexcpFontBold, m.iPrevPriceSwitchRow, 0) = False
            .Cell(flexcpForeColor, m.iPrevPriceSwitchRow, 0) = .ForeColor
            .Cell(flexcpBackColor, nRow - 1, 0) = .BackColor
        
            .Cell(flexcpFontBold, m.iPrevPriceSwitchRow, .Cols - 1) = False             'price column on right
            .Cell(flexcpForeColor, m.iPrevPriceSwitchRow, .Cols - 1) = .ForeColor
            .Cell(flexcpBackColor, nRow - 1, .Cols - 1) = .BackColor
        End If
        If nColor > 0 And ValOfText(.TextMatrix(nRow - 1, 0)) = dPriceAtSwitch Then
            .Cell(flexcpFontBold, nRow - 1, 0) = True
            .Cell(flexcpForeColor, nRow - 1, 0) = nColor
            .Cell(flexcpBackColor, nRow - 1, 0) = nBkColor
        
            .Cell(flexcpFontBold, nRow - 1, .Cols - 1) = True
            .Cell(flexcpForeColor, nRow - 1, .Cols - 1) = nColor
            .Cell(flexcpBackColor, nRow - 1, .Cols - 1) = nBkColor
            
            m.iPrevPriceSwitchRow = nRow - 1
        Else
            m.iPrevPriceSwitchRow = 0
        End If
    End With

End Sub

Private Sub LoadGrid()
On Error GoTo ErrSection:

    Dim Table As cGdTable
    Dim Bars As cGdBars
    
    Dim strIcon$, iIconCol&, iIconType&, iIconColor&, iIconAlign&
    Dim i&, j&, k&, iLastCol&, strText$
    Dim dVolBid#, dVolAsk#, dLastPrice#, dDate#
    Dim dPriceMean#, dPriceAbove#, dPriceBelow#, dPrice#
    Dim dHigh#, dLow#
        
    If g.bUnloading Then Exit Sub
    
    Set Table = m.Data.PriceVolTable
    Set Bars = m.Data.TickBars
    
    If Table Is Nothing Or Bars Is Nothing Then
        Exit Sub            'precautionary
    ElseIf Table.NumRecords = 0 Or Bars.Size = 0 Then
        LoadGridNoVol m.strSym
        Exit Sub            'no data for current session
    End If
    
    m.bInitInprog = True
    
    dLastPrice = RoundToMinMove(Bars(eBARS_Close, Bars.Size - 1), Bars.MinMove(m.Data.SessionDate))
    If Not frmQuotes Is Nothing Then
        m.iCurrPriceColor = frmQuotes.UnchColor
    Else
        m.iCurrPriceColor = fgBidAskDir.ForeColor
    End If
    
    With fgBidAskDir
        .Redraw = flexRDNone
                
        .Cols = Table.NumFields
        .Rows = .FixedRows
        iLastCol = Int(.Cols / 3) * 3
        
        'row 0 in table holds up/down flag, row 1 holds date time stamp
        For i = 1 To Table.NumFields - 1 Step 3
            If Table(i, 1) > 0 Then
                dDate = Table(i, 1)
                If g.bShowInLocalTimeZone Then dDate = ConvertTimeZone(dDate, Bars.Prop(eBARS_ExchangeTimeZoneInf), "")
                strText = DateFormat(dDate, NO_DATE, HH_MM_SS)
                .TextMatrix(0, i) = strText
                .TextMatrix(0, i + 1) = strText
            End If
        Next
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        
        m.Data.HighLow dHigh, dLow
        For i = 2 To Table.NumRecords - 1
            .AddItem Table.GetRecord(i, vbTab)
            If Table(0, i) = dHigh Then
                .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, 0) = m.iColorAsk
            ElseIf Table(0, i) = dLow Then
                .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, 0) = m.iColorBid
            End If
            .TextMatrix(.Rows - 1, 0) = Bars.PriceDisplay(Table(0, i))
            If Table(0, i) = dLastPrice Then
                If Not g.RealTime.Active Then .Cell(flexcpFontBold, .Rows - 1, 0) = True
                .Cell(flexcpForeColor, .Rows - 1, 0) = m.iCurrPriceColor
                m.iLastPriceRow = .Rows - 1
            End If
            If m.Data.LineToolColor(i) > 0 Then
                .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = m.Data.LineToolColor(i)
            End If
            dPrice = Table(0, i)
            For j = 3 To .Cols - 1 Step 3
                dVolBid = Table(j - 2, i)
                dVolAsk = Table(j - 1, i)
                If dVolBid > dVolAsk Then
                    .Cell(flexcpForeColor, .Rows - 1, j - 2) = vbRed
                    .Cell(flexcpForeColor, .Rows - 1, j - 1) = RGB(128, 128, 128)
                    .Cell(flexcpFloodColor, .Rows - 1, j - 2) = kAskColor
                    .Cell(flexcpFloodPercent, .Rows - 1, j - 2) = (dVolAsk - dVolBid) / (dVolBid + dVolAsk) * 100
                ElseIf dVolBid < dVolAsk Then
                    .Cell(flexcpForeColor, .Rows - 1, j - 2) = RGB(128, 128, 128)
                    .Cell(flexcpForeColor, .Rows - 1, j - 1) = vbBlue
                    .Cell(flexcpFloodColor, .Rows - 1, j - 1) = kBidColor
                    .Cell(flexcpFloodPercent, .Rows - 1, j - 1) = (dVolAsk - dVolBid) / (dVolBid + dVolAsk) * 100
                End If
                If Table(j - 2, 0) = -1 Then
                    .Cell(flexcpBackColor, 0, j - 2) = kAskColor    'header row
                End If
                dPriceAbove = m.Data.StatsTable(j, 15)
                dPriceMean = m.Data.StatsTable(j - 1, 15)
                dPriceBelow = m.Data.StatsTable(j - 2, 15)
                If j = iLastCol Then
                    j = j
                End If
                If dPriceMean = dPrice Then
                    .Cell(flexcpBackColor, .Rows - 1, j - 1, .Rows - 1, j - 2) = m.iColorMean
                    If j = iLastCol Then .Cell(flexcpBackColor, .Rows - 1, 0) = m.iColorMean
                ElseIf dPriceAbove = dPrice Then
                    .Cell(flexcpBackColor, .Rows - 1, j - 1, .Rows - 1, j - 2) = m.iColorValue
                    'If j = iLastCol Then .Cell(flexcpBackColor, .Rows - 1, 0) = m.iColorValue
                ElseIf dPriceBelow = dPrice Then
                    .Cell(flexcpBackColor, .Rows - 1, j - 1, .Rows - 1, j - 2) = m.iColorValue
                    'If j = iLastCol Then .Cell(flexcpBackColor, .Rows - 1, 0) = m.iColorValue
                End If
                'cumulative to be shown in price column
                If m.Data.PriceCumVAM = dPrice Or m.Data.PriceCumVBM = dPrice Then
                    .Cell(flexcpBackColor, .Rows - 1, 0) = m.iColorValue
                End If
            Next
             'draw icons if any
            For k = 0 To m.Data.IconCount - 1
                strIcon = m.Data.IconString(k)
                If Len(strIcon) > 0 Then
                    'format: ICON|key|icon|price|column|color|alignment (key = icon type concat with number)
                    If Parse(strIcon, "|", 4) = .TextMatrix(.Rows - 1, 0) Then
                        iIconCol = Val(Parse(strIcon, "|", 5))
                        If iIconCol >= .FixedCols And iIconCol < .Cols Then
                            iIconType = frmFootprintIcons.IconTypeNum(Parse(strIcon, "|", 3))
                            iIconColor = Val(Parse(strIcon, "|", 6))
                            iIconAlign = Val(Parse(strIcon, "|", 7))
                            geFootprintIcon IconPic.hDC, iIconType, vbWhite, iIconColor, "FpIcon.bmp"
                            If FileExist("FpIcon.bmp") Then
                                IconPic.Picture = LoadPicture("FpIcon.bmp")
                                .Cell(flexcpPicture, .Rows - 1, iIconCol) = IconPic.Picture
                                .Cell(flexcpPictureAlignment, .Rows - 1, iIconCol) = iIconAlign
                                .Cell(flexcpData, .Rows - 1, iIconCol) = Parse(strIcon, "|", 2)
                            End If
                        End If
                    End If
                End If
            Next
        Next
    End With
                
'highlight speed cells
    HighlightTradeSpeed
        
'per-column sums & stats
    Set Table = Nothing
    Set Table = m.Data.StatsTable
    
    m.iFgSummaryHeight = 0
    With fgSummary
        .Redraw = flexRDNone
        .Rows = .FixedRows
        .Cols = fgBidAskDir.Cols
        
        If Len(Table(0, i)) <> 0 Then
            'add separator at top
            .Rows = .Rows + 1
            .RowHeight(.Rows - 1) = 25
        End If
        
        For i = 0 To Table.NumRecords - 2   'last row holds prices for Mean,VAM,VBM and meant to be hidden
            strText = Table(0, i)
            If Len(strText) = 0 Then
                .Rows = .Rows + 1           'separators (rows 3,6 in table, 4,7 in grid)
                .RowHeight(.Rows - 1) = 25
            ElseIf InStr(strText, "%") <> 0 Then
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 0) = strText
                For j = 1 To Table.NumFields - 1 Step 3
                    If Table(j, i) > 0 Then .TextMatrix(.Rows - 1, j) = Str(Table(j, i)) & "%"
                    If Table(j + 1, i) > 0 Then .TextMatrix(.Rows - 1, j + 1) = Str(Table(j + 1, i)) & "%"
                Next
            Else
                .AddItem Table.GetRecord(i, vbTab)
            End If
        Next
        .Cell(flexcpAlignment, 0, 1, .Rows - 1, .Cols - 1) = flexAlignCenterCenter

        For i = 1 To .Cols - 1 Step 3
            'text color for summary rows: bid cols=red, ask cols=blue
            .Cell(flexcpForeColor, 0, i, .Rows - 1, i) = vbRed
            .Cell(flexcpForeColor, 0, i + 1, .Rows - 1, i + 1) = vbBlue
            'Differential Row: highlight [bid > ask on up bar] OR [bid < ask on down bar]
            'Table: field i = bid, field i+1 = ask, record 0 = "Totals", record 1 = "Differential"
            '01-08-2007: per Rasa, don't do this for now
'            If m.Data.PriceVolTable(i, 0) = 1 Then          'up bar
'                If Table(i, 0) > Table(i + 1, 0) Then
'                    .Cell(flexcpBackColor, 2, i) = vbYellow
'                End If
'            ElseIf m.Data.PriceVolTable(i, 0) = -1 Then     'down bar
'                If Table(i, 0) < Table(i + 1, 0) Then
'                    .Cell(flexcpBackColor, 2, i + 1) = vbYellow
'                End If
'            End If
        Next
        If m.eStatsView = eTSVView_StatsNone Then
            .Visible = False
        Else
            If Not .Visible Then .Visible = True
            For i = 0 To .Rows - 1
                If m.eStatsView = eTSVView_StatsAll Then
                    .RowHidden(i) = False
                ElseIf i = 0 Or i = 1 Or i = 2 Or i = 3 Or i = 12 Or i = 13 Then
                    .RowHidden(i) = False
                Else
                    .RowHidden(i) = True
                End If
                If Not .RowHidden(i) Then m.iFgSummaryHeight = m.iFgSummaryHeight + .RowHeight(i)
            Next
        End If
                
'        .AutoSizeMode = flexAutoSizeColWidth
'        .AutoSize 0
        .Redraw = flexRDBuffered
    End With
        
    UpdatePullback
    
'draw boxes if any (need to do this after adding pullback columns because #of cols change)
    Dim iTop&, iLeft&, iBottom&, iRight&
    Dim iColor&, iPix&, iFill&
    
    For i = 0 To m.Data.BoxCount - 1
        m.Data.BoxRect i + 1, iTop, iLeft, iBottom, iRight
        iTop = iTop - fgBidAskDir.FixedRows
        iBottom = iBottom - fgBidAskDir.FixedRows
        iPix = m.Data.BoxThickness(i + 1)
        iColor = m.Data.BoxColor(i + 1)
        iFill = m.Data.BoxFill(i + 1)
        DrawBox iTop, iLeft, iBottom, iRight, iColor, iPix, iFill
    Next
    
    With fgBidAskDir
        .Redraw = flexRDNone
        .ColWidth(0) = fgSummary.ColWidth(0)
        fgSummary.Cols = .Cols
        For i = 3 To .Cols - 2 Step 3
            .ColWidth(i - 2) = 800
            .ColWidth(i - 1) = 800
            .ColWidth(i) = 800
            .ColHidden(i) = True            'other vol col
            .Cell(flexcpAlignment, .FixedRows, i - 1, .Rows - 1, i - 1) = flexAlignLeftCenter
            
            fgSummary.ColWidth(i - 2) = 800
            fgSummary.ColWidth(i - 1) = 800
            fgSummary.ColWidth(i) = 800
            fgSummary.ColHidden(i) = True
        Next
        If m.iPriceColLocation = 1 Then         '1=price col on right
            .ColHidden(0) = True
            .ColHidden(.Cols - 1) = False
            fgSummary.ColHidden(0) = True
            fgSummary.ColHidden(fgSummary.Cols - 1) = False
        ElseIf m.iPriceColLocation = 2 Then
            .ColHidden(0) = False
            .ColHidden(.Cols - 1) = False
            fgSummary.ColHidden(0) = False
            fgSummary.ColHidden(fgSummary.Cols - 1) = False
        Else
            .ColHidden(0) = False
            .ColHidden(.Cols - 1) = True
            fgSummary.ColHidden(0) = False
            fgSummary.ColHidden(fgSummary.Cols - 1) = True
        End If
        .Redraw = flexRDBuffered
    End With
        
    If m.bKeepAtEnd Then
        Form_Resize             'need to call here to get correct number of rows in grid for centering price
        FocusGrid
    Else
        With fgBidAskDir
            If m.iTopRow > .FixedRows And m.iTopRow < .Rows Then .TopRow = m.iTopRow
            If m.iLeftCol > .FixedCols And m.iLeftCol < .Cols Then .LeftCol = m.iLeftCol
        End With
        Form_Resize
    End If
            
    Set Table = Nothing
    
    BoldSwitchPrice
        
    m.bInitInprog = False
    
    Exit Sub

ErrSection:
    RaiseError "frmBidAskDir.LoadGrid"

End Sub

Private Sub LoadNewSymData(ByVal strSym$)
On Error GoTo ErrSection:

    Dim nSymbolID&, strMsg$, i&
    Dim tbData As cGdTable
                    
    If g.bUnloading Then Exit Sub
        
    If Left(strSym, 1) = "$" Then
        LoadGridNoVol strSym
        Exit Sub
    End If
    
    m.bInitInprog = True
    
    
    tbToolbar.Tools("ID_Draw").State = ssUnchecked
    ResetMouseVar
    
    m.strSym = ""
    m.nSymID = 0
    nSymbolID = g.SymbolPool.SymbolIDforSymbol(strSym)
       
    If nSymbolID > 0 Then
        m.strSym = strSym
        m.nSymID = nSymbolID
        
        m.iTopRow = 0       'reset
        m.iLeftCol = 0
        
        If m.iTickR <= 0 Then m.iTickR = 4
        If m.iBigLot <= 0 Then m.iBigLot = 500
        If m.iHiLightTrades <= 0 Then m.iHiLightTrades = 10000
        
        Set m.Data = New cTSVData
        
        m.Data.ReverseTick = m.iTickR
        m.Data.BigLot = m.iBigLot
        m.Data.HighlightTrades = m.iHiLightTrades
        
        m.Data.ResetTickBars m.strSym, m.nSymID, Nothing
        
        If m.Data.TickBars.Size > 0 Then
            InfBox "Loading data.  Please wait...", , , "Bid/Ask Directional Analysis", True
            m.Data.BlankRows = m.iBlankRows
            m.Data.BuildTables eTSVTb_TbType_Reversal
            LoadGrid
            Me.Caption = kCaptionBase & " for " & m.strSym & " (" & DateFormat(m.Data.SessionDate, MM_DD_YYYY, NO_TIME) & ")"
        Else
            LoadGridNoVol m.strSym
        End If
                        
        InfBox ""
    End If
    
    m.bInitInprog = False
    
    Exit Sub

ErrSection:
    RaiseError "frmBidAskDir.LoadNewSymData"

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Set m.Data = Nothing
    
    frmFootprintIcons.CloseMe Me
    
    'save form size & location
    SetIniFileProperty Me.Name, GetFormPlacement(Me), "Placement", g.strIniFile
    SetIniFileProperty "StatsView", m.eStatsView, "IOAMT", g.strIniFile

End Sub

Private Sub tbToolbar_ComboCloseUp(ByVal Tool As ActiveToolBars.SSTool)
On Error Resume Next

    Dim i&
    
    i = Tool.ComboBox.ListIndex
    
    If i <> m.eStatsView Then
        m.eStatsView = i
        m.iFgSummaryHeight = 0
        If m.eStatsView = eTSVView_StatsNone Then
            fgSummary.Visible = False
        Else
            With fgSummary
                If Not .Visible Then .Visible = True
                For i = 0 To .Rows - 1
                    If m.eStatsView = eTSVView_StatsAll Then
                        .RowHidden(i) = False
                    ElseIf i = 0 Or i = 1 Or i = 2 Or i = 3 Or i = 12 Or i = 13 Then
                        .RowHidden(i) = False
                    Else
                        .RowHidden(i) = True
                    End If
                    If Not .RowHidden(i) Then m.iFgSummaryHeight = m.iFgSummaryHeight + .RowHeight(i)
                Next
            End With
        End If
        Form_Resize         'do not use external FormResize here
    End If

End Sub

Private Sub tbToolbar_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
On Error GoTo ErrSection:
    
    Dim astrSymbols As New cGdArray
    Dim bFocusSave As Boolean
    Dim iNewPriceColLocation As Long
    
    If m.bInitInprog Then Exit Sub
    
    If Tool.ID <> "ID_Draw" Then
        tbToolbar.Tools("ID_Draw").State = ssUnchecked
        ResetMouseVar
    End If
    
    bFocusSave = m.bKeepAtEnd
    m.bKeepAtEnd = True
    
    Select Case Tool.ID
        Case "ID_Icons"
            If Tool.State = ssChecked Then
                frmFootprintIcons.ShowMe Me
            Else
                frmFootprintIcons.CloseMe Me
            End If

        Case "ID_Draw"
            If FormIsLoaded("frmFootprintIcons") Then frmFootprintIcons.CloseMe Me
        
        Case "ID_Symbol"
            Set astrSymbols = frmSymbolSelector.ShowMe("", False)
            If astrSymbols.Size > 0 Then
                m.bInitInprog = True
                With fgBidAskDir
                    .Redraw = flexRDNone
                    .Rows = .FixedRows
                    .Redraw = flexRDBuffered
                End With
                With fgSummary
                    .Redraw = flexRDNone
                    .Rows = .FixedRows
                    .Redraw = flexRDBuffered
                End With
                DoEvents        'to let grid redraw
                m.bInitInprog = False
                LoadNewSymData astrSymbols(0)
                Form_Resize     'do not use external FormResize here
            End If
        
        Case "ID_Settings"
            If m.Data Is Nothing Or fgBidAskDir.MergeRow(fgBidAskDir.FixedRows) = True Then     '5312
                InfBox "Data not available for selected symbol. Please select a different symbol."
            Else
                m.iTickR = m.Data.ReverseTick
                m.iBigLot = m.Data.BigLot
                m.iHiLightTrades = m.Data.HighlightTrades
                iNewPriceColLocation = m.iPriceColLocation
                
                frmBidAskDirCfg.ShowMe m.iTickR, m.iBigLot, m.iHiLightTrades, m.iBlankRows, iNewPriceColLocation
                
                If m.iTickR <> m.Data.ReverseTick Or _
                    m.iBigLot <> m.Data.BigLot Or _
                    m.iHiLightTrades <> m.Data.HighlightTrades Or _
                    m.iBlankRows <> m.Data.BlankRows Or _
                    m.iPriceColLocation <> iNewPriceColLocation Then
                    
                    m.bInitInprog = True
                                   
                    m.Data.ReverseTick = m.iTickR
                    m.Data.BigLot = m.iBigLot
                    m.Data.HighlightTrades = m.iHiLightTrades
                    m.Data.BlankRows = m.iBlankRows
                    m.iPriceColLocation = iNewPriceColLocation
                    
                    m.Data.BuildTables eTSVTb_TbType_Reversal
                    LoadGrid
                   
                    SetIniFileProperty "RevTick", m.iTickR, "IOAMT", g.strIniFile
                    SetIniFileProperty "BigLot", m.iBigLot, "IOAMT", g.strIniFile
                    SetIniFileProperty "HiLightTrades", m.iHiLightTrades, "IOAMT", g.strIniFile
                    SetIniFileProperty "BidAskDirBlankRows", m.iBlankRows, "IOAMT", g.strIniFile
                    SetIniFileProperty "PriceColLocation", m.iPriceColLocation, "IOAMT", g.strIniFile
                    
                    m.iTopRow = 0       'reset to refocus grid
                    m.iLeftCol = 0
                    
                    m.bInitInprog = False
                End If
            End If
        
        Case "ID_CenterPrice"
            If Tool.State = ssChecked Then
                m.bCenterPrice = True
            Else
                m.bCenterPrice = False
            End If
            FocusGrid
            SetIniFileProperty "BidAskDirCenterPrice", m.bCenterPrice, "IOAMT", g.strIniFile
        
        Case "ID_Close"
            Unload Me
    
    End Select
    m.bKeepAtEnd = bFocusSave
    
    Exit Sub
    
ErrSection:
    RaiseError "frmBidAskDir.tbToolbar_ToolClick"
    
End Sub

Private Sub UpdateGridRT(Table As cGdTable)
On Error GoTo ErrSection:

    Dim i&, k&, iCol&, iUpdateColor&, strText$
    Dim dVolBid#, dVolAsk#, dDate#
    Dim dLastPrice#, dGridPrice#, dMinMove#
    Dim dPriceMean#, dPriceAbove#, dPriceBelow#
    Dim dHigh#, dLow#
    
    Dim tbSummaryStats As cGdTable
    Dim Bars As cGdBars

    iCol = m.Data.PriceVolLastDataCol
    
    If Not frmQuotes Is Nothing Then iUpdateColor = frmQuotes.UpdateColor
    
    Set Bars = m.Data.TickBars
    If Not Bars Is Nothing Then
        dMinMove = Bars.MinMove(m.Data.SessionDate)
        dLastPrice = RoundToMinMove(Bars(eBARS_Close, Bars.Size - 1), dMinMove)
    End If
    
    'get current high/low
    m.Data.HighLow dHigh, dLow
    
    With fgBidAskDir
        .Redraw = flexRDNone
        
        'clear previous highlights
        .Cell(flexcpBackColor, .FixedRows, 0, .Rows - 1, 0) = .Cell(flexcpBackColor, 0, 0)
        'clear previously bolded text
        If m.iLastPriceRow > .FixedRows And m.iLastPriceRow < .Rows Then
            .Cell(flexcpFontBold, m.iLastPriceRow, 0) = False
            .Cell(flexcpForeColor, m.iLastPriceRow, 0) = .ForeColor
            .Cell(flexcpFontBold, m.iLastPriceRow, .Cols - 1) = False
            .Cell(flexcpForeColor, m.iLastPriceRow, .Cols - 1) = .ForeColor
        End If
        .Cell(flexcpFontBold, .FixedRows, .LeftCol, .Rows - 1, .Cols - 1) = False
        .Cell(flexcpFloodPercent, .FixedRows, .Cols - 2, .Rows - 1, .Cols - 1) = 0
        For i = 2 To Table.NumRecords - 1
            dGridPrice = ValOfText(.TextMatrix(i - .FixedRows, 0))
            'highlight high/low
            If dGridPrice = dHigh Then
                .Cell(flexcpBackColor, i - .FixedRows, 0, i - .FixedRows, 0) = m.iColorAsk
            ElseIf dGridPrice = dLow Then
                .Cell(flexcpBackColor, i - .FixedRows, 0, i - .FixedRows, 0) = m.iColorBid
            End If
            If Table(iCol, 1) > 0 Then
                dDate = Table(iCol, 1)
                If g.bShowInLocalTimeZone Then dDate = ConvertTimeZone(dDate, m.Data.TickBars.Prop(eBARS_ExchangeTimeZoneInf), "")
                strText = DateFormat(dDate, NO_DATE, HH_MM_SS)
                .TextMatrix(0, iCol) = strText
                .TextMatrix(0, iCol + 1) = strText
            End If
            dVolBid = Table(iCol, i)
            dVolAsk = Table(iCol + 1, i)
            If dLastPrice = RoundToMinMove(Table(eTSVTb_Price, i), dMinMove) Then
                'bold text in price column
                m.iLastPriceRow = i - .FixedRows
                .Cell(flexcpFontBold, m.iLastPriceRow, 0) = True
                .Cell(flexcpForeColor, m.iLastPriceRow, 0) = iUpdateColor
                m.dLastUpdated = gdTickCount
            End If
            If dVolBid > dVolAsk Then
                .Cell(flexcpForeColor, i - .FixedRows, iCol) = vbRed
                .Cell(flexcpForeColor, i - .FixedRows, iCol + 1) = RGB(128, 128, 128)
                .Cell(flexcpFloodColor, i - .FixedRows, iCol) = kAskColor
                .Cell(flexcpFloodPercent, i - .FixedRows, iCol) = (dVolAsk - dVolBid) / (dVolBid + dVolAsk) * 100
            ElseIf dVolBid < dVolAsk Then
                .Cell(flexcpForeColor, i - .FixedRows, iCol) = RGB(128, 128, 128)
                .Cell(flexcpForeColor, i - .FixedRows, iCol + 1) = vbBlue
                .Cell(flexcpFloodColor, i - .FixedRows, iCol + 1) = kBidColor
                .Cell(flexcpFloodPercent, i - .FixedRows, iCol + 1) = (dVolAsk - dVolBid) / (dVolBid + dVolAsk) * 100
            End If
            If dVolBid > 0 Then
                If ValOfText(.TextMatrix(i - .FixedRows, iCol)) <> dVolBid Then
                    .TextMatrix(i - .FixedRows, iCol) = dVolBid
                    .Cell(flexcpFontBold, i - .FixedRows, iCol) = True
                End If
            End If
            If dVolAsk > 0 Then
                If ValOfText(.TextMatrix(i - .FixedRows, iCol + 1)) <> dVolAsk Then
                .TextMatrix(i - .FixedRows, iCol + 1) = dVolAsk
                .Cell(flexcpFontBold, i - .FixedRows, iCol + 1) = True
                End If
            End If
            If m.Data.LineToolColor(i) > 0 Then
                .Cell(flexcpBackColor, i - .FixedRows, iCol, i - .FixedRows, .Cols - 1) = m.Data.LineToolColor(i)
            Else
                .Cell(flexcpBackColor, i - .FixedRows, iCol, i - .FixedRows, .Cols - 1) = .BackColor
            End If
            
            'clear any previous prices that were colored with mean or value area colors
            If .Cell(flexcpBackColor, i - .FixedRows, 0) = m.iColorMean Or .Cell(flexcpBackColor, i - .FixedRows, 0) = m.iColorValue Then
                .Cell(flexcpBackColor, i - .FixedRows, 0) = .Cell(flexcpBackColor, 0, 0)
            End If
            'set back color for price at mean & value areas
            dPriceBelow = m.Data.StatsTable(iCol, 15)
            dPriceMean = m.Data.StatsTable(iCol + 1, 15)
            dPriceAbove = m.Data.StatsTable(iCol + 2, 15)
            If dPriceMean = Table(0, i) Then
                .Cell(flexcpBackColor, i - .FixedRows, iCol, i - .FixedRows, iCol + 1) = m.iColorMean
                .Cell(flexcpBackColor, i - .FixedRows, 0) = m.iColorMean
            ElseIf dPriceAbove = Table(0, i) Or dPriceBelow = Table(0, i) Then
                .Cell(flexcpBackColor, i - .FixedRows, iCol, i - .FixedRows, iCol + 1) = m.iColorValue
                .Cell(flexcpBackColor, i - .FixedRows, 0) = m.iColorValue
            End If
            
            .Cell(flexcpBackColor, i - .FixedRows, .Cols - 1) = .Cell(flexcpBackColor, i - .FixedRows, 0)
            .Cell(flexcpForeColor, i - .FixedRows, .Cols - 1) = .Cell(flexcpForeColor, i - .FixedRows, 0)
            .Cell(flexcpFontBold, i - .FixedRows, .Cols - 1) = .Cell(flexcpFontBold, i - .FixedRows, 0)
            
            'cumulative to be shown in price column
            '            If m.Data.PriceCumVAM = Table(0, i) Or m.Data.PriceCumVBM = Table(0, i) Then
            '                .Cell(flexcpBackColor, .Rows - 1, 0) = m.iColorValue
            '            End If
         Next
        HighlightTradeSpeed
        .Redraw = flexRDBuffered
    End With

    'per-column sums & stats
    Set tbSummaryStats = m.Data.StatsTable
    
    If Not tbSummaryStats Is Nothing Then
        k = tbSummaryStats.NumRecords - 1           'last record holds Mean,VAM,VBM and is meant to be hidden
        With fgSummary
            .Redraw = flexRDNone
            For i = 0 To .Rows - 1
                If i <= k Then
                    strText = tbSummaryStats(0, i)
                    If strText = .TextMatrix(i, 0) Then
                        If tbSummaryStats(iCol, i) = 0 Then
                            .TextMatrix(i, iCol) = ""
                        ElseIf InStr(strText, "%") = 0 Then
                            .TextMatrix(i, iCol) = tbSummaryStats(iCol, i)
                        Else
                            .TextMatrix(i, iCol) = Str(tbSummaryStats(iCol, i)) & "%"
                        End If
                        If tbSummaryStats(iCol + 1, i) = 0 Then
                            .TextMatrix(i, iCol + 1) = ""
                        ElseIf InStr(strText, "%") = 0 Then
                            .TextMatrix(i, iCol + 1) = tbSummaryStats(iCol + 1, i)
                        Else
                            .TextMatrix(i, iCol + 1) = Str(tbSummaryStats(iCol + 1, i)) & "%"
                        End If
                    End If
                End If
            Next
            .Redraw = flexRDBuffered
        End With
    End If
                        
    UpdatePullback True
    RefreshBoxes
    
    Exit Sub

ErrSection:
    RaiseError "frmBidAskDir.UpdateGridRT"

End Sub

Private Sub tmr_Timer()
On Error GoTo ErrSection:
            
    Dim i&, iCol&, strText$
    Dim dVolBid#, dVolAsk#, dDate#, dTickCount#
    Dim Table As cGdTable
    
    Dim bDontCare As Boolean
    Dim bTableRebuild As Boolean
    Dim bChanged As Boolean
    
    If m.bInitInprog Or m.bTimerInProg Or g.bUnloading Then
        Exit Sub
    ElseIf Not g.RealTime.Active Then
        m.bTimerInProg = False
        tmr.Enabled = False
        Exit Sub
    End If
    
    m.bTimerInProg = True
    
    If Not frmQuotes Is Nothing Then
        If m.iCurrPriceColor <> frmQuotes.UnchColor Then
            m.iCurrPriceColor = frmQuotes.UnchColor
            bChanged = True
        End If
    End If
    
    dTickCount = gdTickCount - m.dLastUpdated
    If dTickCount > 0 And dTickCount > g.nUpdatedColorDuration Then
        With fgBidAskDir
            If m.iLastPriceRow > .FixedRows And m.iLastPriceRow < .Rows Then
                .Cell(flexcpForeColor, m.iLastPriceRow, 0) = m.iCurrPriceColor
                .Cell(flexcpForeColor, m.iLastPriceRow, .Cols - 1) = m.iCurrPriceColor
            End If
        End With
    End If
    
    If Not m.Data Is Nothing Then
        If g.RealTime.Active And g.RealTime.FeedTime > 0 Then
            i = m.Data.UpdateDataRT(bDontCare, bDontCare, bTableRebuild)
            Set Table = m.Data.PriceVolTable
            If i = -1 Then
                LoadNewSymData m.strSym     'there was no data when form was opened
            ElseIf bTableRebuild Then
                LoadGrid                    'new high and/or low
            ElseIf i > 0 And Not Table Is Nothing And Not m.bReloadData Then
                If fgBidAskDir.Cols - 3 = Table.NumFields Then
                    'no new column
                    UpdateGridRT Table
                    BoldSwitchPrice
                ElseIf fgBidAskDir.Cols - 3 < Table.NumFields Then
                    LoadGrid
                End If
            ElseIf m.iLastPriceRow > fgBidAskDir.FixedRows And m.iLastPriceRow < fgBidAskDir.Rows Then
                With fgBidAskDir
                    If m.iLastPriceRow > .FixedRows And m.iLastPriceRow < .Rows Then
                        If Not .Cell(flexcpFontBold, m.iLastPriceRow, 0) Then
                            .Cell(flexcpFontBold, m.iLastPriceRow, 0) = True
                        End If
                    End If
                    If bChanged Then .Cell(flexcpForeColor, m.iLastPriceRow, 0) = m.iCurrPriceColor
                End With
                BoldSwitchPrice
            End If
        End If
    End If
    
    If m.bReloadData Then
        m.Data.ResetTickBars m.strSym, m.nSymID, Nothing
        m.Data.BuildTables eTSVTb_TbType_Reversal
        LoadGrid
        Me.Caption = kCaptionBase & " for " & m.strSym & " (" & DateFormat(m.Data.SessionDate, MM_DD_YYYY, NO_TIME) & ")"
    ElseIf m.bCenterPrice And Not bTableRebuild Then
        If Not fgBidAskDir.RowIsVisible(m.iLastPriceRow) Then
            FocusGrid
        End If
    End If
    
    m.bReloadData = False
    
    m.bTimerInProg = False
    
    Exit Sub

ErrSection:
    RaiseError "frmBidAskDir.tmr_Timer"

End Sub

Public Sub RefreshData()
On Error GoTo ErrSection:

    If m.bInitInprog Then Exit Sub
    
    If m.bTimerInProg Then
        m.bReloadData = True
    Else
        tmr.Enabled = False
        m.Data.ResetTickBars m.strSym, m.nSymID, Nothing
        m.Data.BuildTables eTSVTb_TbType_Reversal
        LoadGrid
        Me.Caption = kCaptionBase & " for " & m.strSym & " (" & DateFormat(m.Data.SessionDate, MM_DD_YYYY, NO_TIME) & ")"
        tmr.Enabled = True
    End If
    
    Exit Sub

ErrSection:
    RaiseError "frmBidAskDir.RefreshData"
    
End Sub

Private Sub FocusGrid()
On Error Resume Next

    Dim i&, iShowCol&, iShowRow&

    If fgBidAskDir.Rows <= fgBidAskDir.FixedRows Then Exit Sub

    With fgBidAskDir
        If m.iTopRow < .FixedRows Then m.iTopRow = .FixedRows
        If m.iLeftCol < 0 Then m.iLeftCol = .Cols / 2
        
        If m.bKeepAtEnd Then
            iShowCol = .Cols - 1
        Else
            iShowCol = m.iLeftCol
        End If
        
        If m.bCenterPrice Then
            i = (.BottomRow - .TopRow) / 2
            If m.iLastPriceRow - i > .FixedRows Then
                .TopRow = m.iLastPriceRow - i
                iShowRow = m.iLastPriceRow
            End If
        Else
            iShowRow = m.iTopRow
        End If
        
        .ShowCell iShowRow, iShowCol
        
        m.iTopRow = .TopRow
        m.iLeftCol = .LeftCol
        fgSummary.LeftCol = .LeftCol
        Form_Resize
    End With

End Sub

Private Function InRangeCol(ByVal iCol&) As Boolean
On Error GoTo ErrSection:

    With fgBidAskDir
        If iCol >= .FixedCols And iCol < .Cols Then InRangeCol = True
    End With
    
    Exit Function
    
ErrSection:
    RaiseError "frmBidAskDir.InRangeCol"

End Function

Private Function InRangeRow(ByVal iRow&) As Boolean
On Error GoTo ErrSection:
    
    With fgBidAskDir
        If iRow >= .FixedRows And iRow < .Rows Then InRangeRow = True
    End With
    
    Exit Function
    
ErrSection:
    RaiseError "frmBidAskDir.InRangeRow"

End Function

Private Sub HandleHighlight(ByVal bClear As Boolean)
On Error GoTo ErrSection:
    
    Dim i&, iColor&
        
    If m.bInitInprog Or m.bReloadData Or m.bTimerInProg Or g.bUnloading Then
        ResetMouseVar
        Exit Sub
    End If
    
    m.bEditDrawInProg = True
    
    With fgBidAskDir
        'row 0 in table holds up/down flag, row 1 holds date time stamp
        If InRangeRow(m.iMouseRowDown) Then
            If bClear Then
                iColor = .BackColor
                m.Data.LineToolColor(m.iMouseRowDown + .FixedRows) = 0
            Else
                iColor = gdColor.Color
                m.Data.LineToolColor(m.iMouseRowDown + .FixedRows) = gdColor.Color
            End If
            .Cell(flexcpBackColor, m.iMouseRowDown, .FixedCols, m.iMouseRowDown, .Cols - 1) = iColor
            
            If bClear Then RefreshBoxes

'01-08-2007: Per Rasa, not doing yellow trade speed for now
'            For i = .FixedCols To .Cols - 1
'                If .Cell(flexcpBackColor, m.iMouseRowDown, i) <> vbYellow Then
'                    .Cell(flexcpBackColor, m.iMouseRowDown, i) = iColor
'                End If
'            Next
        End If
    End With
    
    ResetMouseVar
    
    m.bEditDrawInProg = False
    
    Exit Sub

ErrSection:
    RaiseError "frmBidAskDir.HandleHighlight"
    
End Sub

Private Sub RefreshBoxes()
On Error Resume Next

    'refresh filled boxes in case some cells got set to background color due to removal of line tool or RT update
    Dim iTop&, iLeft&, iBottom&, iRight&
    Dim iColor&, iPix&, iFill&, i&
    
    For i = 0 To m.Data.BoxCount - 1
        m.Data.BoxRect i + 1, iTop, iLeft, iBottom, iRight
        iTop = iTop - fgBidAskDir.FixedRows
        iBottom = iBottom - fgBidAskDir.FixedRows
        iPix = m.Data.BoxThickness(i + 1)
        iColor = m.Data.BoxColor(i + 1)
        iFill = m.Data.BoxFill(i + 1)
        If iFill = 1 Then
            DrawBox iTop, iLeft, iBottom, iRight, iColor, iPix, iFill
        End If
    Next

End Sub

Private Sub DrawBox(ByVal iTop&, ByVal iLeft&, ByVal iBottom&, ByVal iRight&, _
    ByVal iColor&, ByVal iPix&, ByVal iFill&)
On Error GoTo ErrSection:
        
    Dim i&, iLineColor&
    Dim iCol&, iRow&
    
    If InRangeRow(iTop) And InRangeRow(iBottom) Then
        If InRangeCol(iLeft) And InRangeCol(iRight) Then
            With fgBidAskDir
                .Select iTop, iLeft, iBottom, iRight
                .CellBorder iColor, iPix, iPix, iPix, iPix, 0, 0
                .Select 0, 0
                If iLeft > iRight Then
                    i = iLeft
                    iLeft = iRight
                    iRight = i
                End If
                If iTop > iBottom Then
                    i = iTop
                    iTop = iBottom
                    iBottom = i
                End If
                For iCol = iLeft To iRight
                    For iRow = iTop To iBottom
                        If iFill = 1 Then
                            If .Cell(flexcpBackColor, iRow, iCol) = .BackColor Or .Cell(flexcpBackColor, iRow, iCol) = 0 Then
                                .Cell(flexcpBackColor, iRow, iCol) = iColor
                            End If
                        ElseIf .Cell(flexcpBackColor, iRow, iCol) = m.iBoxColor Then
                            .Cell(flexcpBackColor, iRow, iCol) = .BackColor     'clear box filled color
                        End If
                    Next
                Next
            End With
        End If
    End If
    
    Exit Sub

ErrSection:
    RaiseError "frmBidAskDir.DrawBox"
    
End Sub

Private Sub HandleBox(Button As Integer)
On Error GoTo ErrSection:

    Dim iTop&, iLeft&, iBottom&, iRight&
    Dim iBoxId&, iColor&, iPix&, iFill&

    If m.bInitInprog Or m.bReloadData Or m.bTimerInProg Or g.bUnloading Then
        ResetMouseVar
        Exit Sub
    End If
    
    m.bEditDrawInProg = True
    
    With fgBidAskDir
        iBoxId = m.Data.BoxId(m.iMouseRowDown + .FixedRows, m.iMouseColDown)
    End With
    
    If iBoxId > 0 Then
        m.Data.BoxRect iBoxId, iTop, iLeft, iBottom, iRight
        iTop = iTop - fgBidAskDir.FixedRows
        iBottom = iBottom - fgBidAskDir.FixedRows
        iColor = m.Data.BoxColor(iBoxId)
        iPix = m.Data.BoxThickness(iBoxId)
        iFill = m.Data.BoxFill(iBoxId)
        
        If Button = vbLeftButton Then
            'edit exisiting box
            If frmPriceVolCfg.ShowBoxSettings(Me, iColor, iPix, iFill) Then
                DrawBox iTop, iLeft, iBottom, iRight, m.iBoxColor, m.iBoxPix, m.iBoxFill
                m.Data.BoxColor(iBoxId) = m.iBoxColor
                m.Data.BoxThickness(iBoxId) = m.iBoxPix
                m.Data.BoxFill(iBoxId) = m.iBoxFill
            End If
        Else
            'clear/delete box
            DrawBox iTop, iLeft, iBottom, iRight, fgBidAskDir.BackColor, 0, 0
            m.Data.RemoveBox iBoxId
        End If
        
        ResetMouseVar
        m.bEditDrawInProg = False
        Exit Sub
    End If
    
    If Button = vbLeftButton Then
        If InRangeRow(m.iMouseRowDown) And InRangeRow(m.iMouseRowDownPrev) Then
            If InRangeCol(m.iMouseColDown) And InRangeCol(m.iMouseColDownPrev) Then
                If frmPriceVolCfg.ShowBoxSettings(Me, m.iBoxColor, m.iBoxPix, m.iBoxFill) Then
                    'add new box
                    With fgBidAskDir
                        DrawBox m.iMouseRowDown, m.iMouseColDown, m.iMouseRowDownPrev, m.iMouseColDownPrev, _
                                m.iBoxColor, m.iBoxPix, m.iBoxFill
                        m.Data.AddBox m.iMouseRowDown + .FixedRows, m.iMouseColDown, _
                            m.iMouseRowDownPrev + .FixedRows, m.iMouseColDownPrev, _
                            m.iBoxColor, m.iBoxPix, 0, m.iBoxFill
                    End With
                End If
                ResetMouseVar
            End If
        End If
    Else
        ResetMouseVar
    End If
    
    m.bEditDrawInProg = False
    
    Exit Sub

ErrSection:
    RaiseError "frmBidAskDir.HandleBox"
    
End Sub

Private Sub ResetMouseVar()
On Error GoTo ErrSection:

    m.iMouseColDown = -1
    m.iMouseRowDown = -1
    
    m.iMouseColDownPrev = -1
    m.iMouseRowDownPrev = -1
    
    Exit Sub

ErrSection:
    RaiseError "frmBidAskDir.ResetMouseVar"
    
End Sub

Private Sub gdColor_Changed()
On Error Resume Next
    
    gdColor.Visible = False
    HandleHighlight False

End Sub

Private Sub gdColor_ColorClicked()
On Error Resume Next

    If gdColor.DropDownVisible Then gdColor.UserControl_Click
    gdColor.Visible = False
    HandleHighlight False

    Exit Sub

ErrSection:
    RaiseError "frmBidAskDir.gdColor_ColorClicked"
    
End Sub

Public Sub BoxSettings(ByVal iColor&, ByVal iPix&, ByVal iFill&)
On Error GoTo ErrSection:

    m.iBoxColor = iColor
    m.iBoxPix = iPix
    m.iBoxFill = iFill
    
    Exit Sub

ErrSection:
    RaiseError "frmBidAskDir.BoxSettings"
    
End Sub

Public Property Get BlankRows() As Long
On Error GoTo ErrSection:

    BlankRows = m.iBlankRows
    
    Exit Property
    
ErrSection:
    RaiseError "frmBidAskDir.BlankRows.Get"
    
End Property

Public Property Let BlankRows(ByVal iRows&)
On Error GoTo ErrSection:

    m.iBlankRows = iRows
    
    Exit Property

ErrSection:
    RaiseError "frmBidAskDir.BlankRows.Let"
    
End Property

Private Sub LoadGridNoVol(ByVal strSym$)
On Error GoTo ErrSection:

    Dim i&
    
    Me.Caption = kCaptionBase & " for " & strSym
    
    fgSummary.Visible = False
    
    With fgBidAskDir
        .Redraw = flexRDNone
        .Cols = 1
        .Rows = 2
        .MergeRow(0) = True
        .MergeRow(1) = True
        .TextMatrix(1, 0) = kFootPrintNoVol
        .Cell(flexcpBackColor, 1) = vbWhite
        .Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
        .Redraw = flexRDBuffered
    End With
        
    Exit Sub
    
ErrSection:
    RaiseError "frmBidAskDir.LoadGridNoVol"

End Sub

Public Sub IconPaletteClose()
On Error Resume Next:
    
    tbToolbar.Tools("ID_Icons").State = ssUnchecked
    fgBidAskDir.MousePointer = flexDefault
    
End Sub

