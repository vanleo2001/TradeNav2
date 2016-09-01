VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Begin VB.Form frmPriceVol 
   Caption         =   "AMPT Volume at Price"
   ClientHeight    =   6750
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10530
   LinkTopic       =   "Form1"
   ScaleHeight     =   6750
   ScaleWidth      =   10530
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox IconPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   9660
      ScaleHeight     =   33
      ScaleMode       =   0  'User
      ScaleWidth      =   17
      TabIndex        =   3
      Top             =   255
      Visible         =   0   'False
      Width           =   255
   End
   Begin gdOCX.gdSelectColor gdColor 
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      Top             =   3240
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      CustomColor     =   255
   End
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3480
      Top             =   2160
   End
   Begin VSFlex7LCtl.VSFlexGrid fgPriceVol 
      Height          =   1815
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   1695
      _cx             =   2990
      _cy             =   3201
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
   Begin ActiveToolBars.SSActiveToolBars tbToolbar 
      Left            =   4200
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131083
      ToolBarsCount   =   1
      ToolsCount      =   8
      DisplayContextMenu=   0   'False
      Tools           =   "frmPriceVol.frx":0000
      ToolBars        =   "frmPriceVol.frx":280E
   End
   Begin VSFlex7LCtl.VSFlexGrid fgSummary 
      Height          =   1815
      Left            =   600
      TabIndex        =   2
      Top             =   2760
      Width           =   1695
      _cx             =   2990
      _cy             =   3201
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
Attribute VB_Name = "frmPriceVol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const kCaptionBase = "AMPT Volume at Price"

Private Enum ePriceVol_View
    eTSVView_VolAll = 0
    eTSVView_VolBidAsk
    eTSVView_VolDelta
    eTSVView_AuctionBar
End Enum

Private Enum ePriceVol_StatsView
    eTSVView_StatsAll = 0
    eTSVView_StatsSummary
    eTSVView_StatsNone
End Enum

Private Enum eSeparator_Offset
    eSepOffset_BidVol = 1
    eSepOffset_AskVol = 2
    eSepOffset_DeltaBidAsk = 3
    eSepOffSet_DeltaCum = 5
    eSepOffset_Percent = 6
    eSepOffSet_VAM = 8
    eSepOffset_Mean = 9
    eSepOffset_VBM = 10
    eSepOffset_Totals = 12
    eSepOffset_AvgVol = 13
    eSepOffset_SHVol = 14
    eSepOffset_Range = 16
    eSepOffset_AvgRange = 17
    eSepOffset_SHRange = 18
End Enum

Private Type mPrivate
    Data As cTSVData
    
    eView As ePriceVol_View
    eStatsView As ePriceVol_StatsView
    
    iAuctionBarPix As Long      'configurable items
    iTrianglePix As Long
    iColorBid As Long
    iColorAsk As Long
    iColorHistogram As Long
    iColorMean As Long
    iColorMode As Long
    iColorTriangle As Long
    iColorUnfairHigh As Long
    iColorUnfairLow As Long
    iColorValue As Long
    
    iBoxColor As Long
    iBoxPix As Long
    iBoxFill As Long
    
    iMouseColDown As Long
    iMouseRowDown As Long
    
    iMouseColDownPrev As Long
    iMouseRowDownPrev As Long
        
    iTopRow As Long
    iLeftCol As Long
    iLastPriceRow As Long
    iBlankRows As Long              'number of extra rows above/below high/low
    iCurrPriceColor As Long
    
    nSymID As Long
    strSym As String
    
    dHistogramMaxVol As Double
    iFgSummaryHeight As Long        'initial height without scroll bar
    
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
    
    tbToolbar.Tools("ID_View").ComboBox.ListIndex = m.eView
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
    
    ShowForm Me, eForm_Nonmodal, frmMain
        
    bCenter = GetIniFileProperty("PriceVolCenterPrice", True, "IOAMT", g.strIniFile)
    If bCenter = False Then
        m.bCenterPrice = True   'always center on price when first shown
        FocusGrid
    End If
    
    m.bCenterPrice = bCenter
    tbToolbar.Tools("ID_AutoCenter").State = Abs(m.bCenterPrice)
        
    tmr.Enabled = g.RealTime.Active
    
    Exit Sub

ErrSection:
    RaiseError "frmPriceVol.ShowMe"

End Sub

Private Sub DrawStatsBlueBox(ByVal iCol&, dBid#, dAsk#, iRow&)
On Error GoTo ErrSection:

    Dim iStatHighVol&

    iStatHighVol = m.Data.StatHighVolAtPrice(iCol)

    If iStatHighVol > 0 And iStatHighVol < dBid + dAsk Then
        With fgPriceVol
        If m.eView = eTSVView_VolAll Then
            .Select iRow, iCol
        Else
            .Select iRow, iCol, iRow, iCol + 1
        End If
        .CellBorder vbBlue, 1, 1, 1, 1, 0, 0
        .Select 0, 0
        End With
    End If
    
    Exit Sub

ErrSection:
    RaiseError "frmPriceVol.DrawStatsBlueBox"

End Sub

Private Sub SetGridRow(ByVal iCol&, ByVal dBid#, ByVal dAsk#, ByVal dOther#, ByVal dRowTotal#, ByVal iRow&)
On Error GoTo ErrSection:

    Dim iFloodPercent&, iLineColor&
    
    If m.eView = eTSVView_AuctionBar Then Exit Sub          'precautionary
        
    With fgPriceVol
        If iRow < .Rows And iCol < .Cols Then
            
            
            If dBid + dAsk > 0 Then iFloodPercent = Int(Abs(dBid - dAsk) / (dBid + dAsk) * 100)
            
            'draw blue box
            DrawStatsBlueBox iCol, dBid, dAsk, iRow
            
            'set data values & flood cell as needed
            If m.eView = eTSVView_VolAll Then
                .TextMatrix(iRow, iCol) = dBid + dAsk + dOther
            ElseIf m.eView = eTSVView_VolDelta Then
                .TextMatrix(iRow, iCol) = dBid + dAsk
                .TextMatrix(iRow, iCol + 1) = Abs(dBid - dAsk)
                If dBid > dAsk Then
                    .Cell(flexcpFloodColor, iRow, iCol + 1) = m.iColorAsk
                    .Cell(flexcpFloodPercent, iRow, iCol + 1) = 100
                ElseIf dBid < dAsk Then
                    .Cell(flexcpFloodColor, iRow, iCol + 1) = m.iColorBid
                    .Cell(flexcpFloodPercent, iRow, iCol + 1) = 100
                End If
            Else
                .TextMatrix(iRow, iCol) = dBid
                .TextMatrix(iRow, iCol + 1) = dAsk
                If dBid > dAsk Then
                    .Cell(flexcpFloodColor, iRow, iCol) = m.iColorAsk
                    .Cell(flexcpFloodPercent, iRow, iCol) = iFloodPercent * -1
                ElseIf dBid < dAsk Then
                    .Cell(flexcpFloodColor, iRow, iCol + 1) = m.iColorBid
                    .Cell(flexcpFloodPercent, iRow, iCol + 1) = iFloodPercent
                End If
            End If
            
            .TextMatrix(iRow, 29) = dRowTotal
            
            iLineColor = m.Data.LineToolColor(iRow - .FixedRows)
            If iLineColor > 0 Then
                .Cell(flexcpBackColor, iRow, 1, iRow, .Cols - 1) = iLineColor
            End If
        End If
    End With
    
    Exit Sub

ErrSection:
    RaiseError "frmPriceVol.SetGridRow"

End Sub

Private Sub ColorGridRow(ByVal eGroup As eTSV_Groups)
On Error GoTo ErrSection:
        
    Dim iMean&, iAbove&, iBelow&
    Dim iColStart&, iColEnd&
    
    If m.eView = eTSVView_AuctionBar And eGroup <> eTSV_Group_Current Then
        Exit Sub          'precautionary
    End If
    
    If eGroup = eTSV_Group_AB Then
        iColStart = eTSVTb_VolBid_A
        iColEnd = eTSVTb_VolAsk_B
    ElseIf eGroup = eTSV_Group_AE Then
        iColStart = eTSVTb_VolBid_C
        iColEnd = eTSVTb_VolAsk_E
    ElseIf eGroup = eTSV_Group_AI Then
        iColStart = eTSVTb_VolBid_F
        iColEnd = eTSVTb_VolAsk_I
    ElseIf eGroup = eTSV_Group_AN Then
        iColStart = eTSVTb_VolBid_J
        iColEnd = eTSVTb_VolAsk_N
    ElseIf eGroup = eTSV_Group_Current Then
        iColStart = eTSVTb_Price
        iColEnd = eTSVTb_Price
    Else
        Exit Sub        'precautionary
    End If
    
    '(calling routine already set grid redraw to none so no need to do it here)
    With fgPriceVol
        m.Data.ColorRowIdx eGroup, True, iMean, iAbove, iBelow
         'clear out previous colored rows
        iMean = iMean + .FixedRows
        iAbove = iAbove + .FixedRows
        iBelow = iBelow + .FixedRows
        
        If iMean > .FixedRows And iMean < .Rows Then
            .Cell(flexcpBackColor, iMean, iColStart, iMean, iColEnd) = .BackColor
        End If
        If iAbove > .FixedRows And iAbove < .Rows Then
            .Cell(flexcpBackColor, iAbove, iColStart, iAbove, iColEnd) = .BackColor
        End If
        If iBelow > .FixedRows And iBelow < .Rows Then
            .Cell(flexcpBackColor, iBelow, iColStart, iBelow, iColEnd) = .BackColor
        End If
        
        m.Data.ColorRowIdx eGroup, False, iMean, iAbove, iBelow
        'color new rows
        iMean = iMean + .FixedRows
        iAbove = iAbove + .FixedRows
        iBelow = iBelow + .FixedRows
        
        If iMean > .FixedRows And iMean < .Rows Then
            .Cell(flexcpBackColor, iMean, iColStart, iMean, iColEnd) = m.iColorMean
        End If
        If iAbove > .FixedRows And iAbove < .Rows Then
            .Cell(flexcpBackColor, iAbove, iColStart, iAbove, iColEnd) = m.iColorValue
        End If
        If iBelow > .FixedRows And iBelow < .Rows Then
            .Cell(flexcpBackColor, iBelow, iColStart, iBelow, iColEnd) = m.iColorValue
        End If
    End With
    
    Exit Sub
    
ErrSection:
    RaiseError "frmPriceVol.ColorGridRow"

End Sub

Private Property Get MergeColSpace(ByVal eCol As eTSV_TbFields) As Long
On Error Resume Next

    If eCol = eTSVTb_VolBid_B Or eCol = eTSVTb_VolAsk_B Or _
       eCol = eTSVTb_VolBid_D Or eCol = eTSVTb_VolAsk_D Or _
       eCol = eTSVTb_VolBid_F Or eCol = eTSVTb_VolAsk_F Or _
       eCol = eTSVTb_VolBid_H Or eCol = eTSVTb_VolAsk_H Or _
       eCol = eTSVTb_VolBid_J Or eCol = eTSVTb_VolAsk_J Or _
       eCol = eTSVTb_VolBid_L Or eCol = eTSVTb_VolAsk_L Or _
       eCol = eTSVTb_VolBid_N Or eCol = eTSVTb_VolAsk_N Then
        
        MergeColSpace = 1
    Else
    
        MergeColSpace = 0
    End If

End Property

Private Sub LoadGridBidAsk(Table As cGdTable, Bars As cGdBars, dLargestVol As Double)
On Error GoTo ErrSection:

    Dim i&, dLastPrice#, dMinMove#
    
    dMinMove = Bars.MinMove(m.Data.SessionDate)
    dLastPrice = RoundToMinMove(Bars(eBARS_Close, Bars.Size - 1), dMinMove)
    
    With fgPriceVol
        For i = 0 To Table.NumRecords - 1
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = Bars.PriceDisplay(Table(eTSVTb_Price, i))
            
            If dLastPrice = RoundToMinMove(Table(eTSVTb_Price, i), dMinMove) Then
                m.iLastPriceRow = .Rows - 1
            End If
            
            If Table(eTSVTb_VolBid_A, i) > 0 Or Table(eTSVTb_VolAsk_A, i) > 0 Or Table(eTSVTb_VolOther_A, i) > 0 Then
                If Table(eTSVTb_VolRow_Total, i) > dLargestVol Then dLargestVol = Table(eTSVTb_VolRow_Total, i)
                SetGridRow eTSVTb_VolBid_A, Table(eTSVTb_VolBid_A, i), Table(eTSVTb_VolAsk_A, i), Table(eTSVTb_VolOther_A, i), Table(eTSVTb_VolRow_Total, i), .Rows - 1
            End If
            
            If Table(eTSVTb_VolBid_B, i) > 0 Or Table(eTSVTb_VolAsk_B, i) Or Table(eTSVTb_VolOther_B, i) > 0 Then
                If Table(eTSVTb_VolRow_Total, i) > dLargestVol Then dLargestVol = Table(eTSVTb_VolRow_Total, i)
                SetGridRow eTSVTb_VolBid_B, Table(eTSVTb_VolBid_B, i), Table(eTSVTb_VolAsk_B, i), Table(eTSVTb_VolOther_B, i), Table(eTSVTb_VolRow_Total, i), .Rows - 1
            End If
            
            If Table(eTSVTb_VolBid_C, i) > 0 Or Table(eTSVTb_VolAsk_C, i) Or Table(eTSVTb_VolOther_C, i) > 0 Then
                If Table(eTSVTb_VolRow_Total, i) > dLargestVol Then dLargestVol = Table(eTSVTb_VolRow_Total, i)
                SetGridRow eTSVTb_VolBid_C, Table(eTSVTb_VolBid_C, i), Table(eTSVTb_VolAsk_C, i), Table(eTSVTb_VolOther_C, i), Table(eTSVTb_VolRow_Total, i), .Rows - 1
            End If
            
            If Table(eTSVTb_VolBid_D, i) > 0 Or Table(eTSVTb_VolAsk_D, i) Or Table(eTSVTb_VolOther_D, i) > 0 Then
                If Table(eTSVTb_VolRow_Total, i) > dLargestVol Then dLargestVol = Table(eTSVTb_VolRow_Total, i)
                SetGridRow eTSVTb_VolBid_D, Table(eTSVTb_VolBid_D, i), Table(eTSVTb_VolAsk_D, i), Table(eTSVTb_VolOther_D, i), Table(eTSVTb_VolRow_Total, i), .Rows - 1
            End If
            
            If Table(eTSVTb_VolBid_E, i) > 0 Or Table(eTSVTb_VolAsk_E, i) Or Table(eTSVTb_VolOther_E, i) > 0 Then
                If Table(eTSVTb_VolRow_Total, i) > dLargestVol Then dLargestVol = Table(eTSVTb_VolRow_Total, i)
                SetGridRow eTSVTb_VolBid_E, Table(eTSVTb_VolBid_E, i), Table(eTSVTb_VolAsk_E, i), Table(eTSVTb_VolOther_E, i), Table(eTSVTb_VolRow_Total, i), .Rows - 1
            End If
            
            If Table(eTSVTb_VolBid_F, i) > 0 Or Table(eTSVTb_VolAsk_F, i) Or Table(eTSVTb_VolOther_F, i) > 0 Then
                If Table(eTSVTb_VolRow_Total, i) > dLargestVol Then dLargestVol = Table(eTSVTb_VolRow_Total, i)
                SetGridRow eTSVTb_VolBid_F, Table(eTSVTb_VolBid_F, i), Table(eTSVTb_VolAsk_F, i), Table(eTSVTb_VolOther_F, i), Table(eTSVTb_VolRow_Total, i), .Rows - 1
            End If
            
            If Table(eTSVTb_VolBid_G, i) > 0 Or Table(eTSVTb_VolAsk_G, i) Or Table(eTSVTb_VolOther_G, i) > 0 Then
                If Table(eTSVTb_VolRow_Total, i) > dLargestVol Then dLargestVol = Table(eTSVTb_VolRow_Total, i)
                SetGridRow eTSVTb_VolBid_G, Table(eTSVTb_VolBid_G, i), Table(eTSVTb_VolAsk_G, i), Table(eTSVTb_VolOther_G, i), Table(eTSVTb_VolRow_Total, i), .Rows - 1
            End If
            
            If Table(eTSVTb_VolBid_H, i) > 0 Or Table(eTSVTb_VolAsk_H, i) Or Table(eTSVTb_VolOther_H, i) > 0 Then
                If Table(eTSVTb_VolRow_Total, i) > dLargestVol Then dLargestVol = Table(eTSVTb_VolRow_Total, i)
                SetGridRow eTSVTb_VolBid_H, Table(eTSVTb_VolBid_H, i), Table(eTSVTb_VolAsk_H, i), Table(eTSVTb_VolOther_H, i), Table(eTSVTb_VolRow_Total, i), .Rows - 1
            End If
            
            If Table(eTSVTb_VolBid_I, i) > 0 Or Table(eTSVTb_VolAsk_I, i) Or Table(eTSVTb_VolOther_I, i) > 0 Then
                If Table(eTSVTb_VolRow_Total, i) > dLargestVol Then dLargestVol = Table(eTSVTb_VolRow_Total, i)
                SetGridRow eTSVTb_VolBid_I, Table(eTSVTb_VolBid_I, i), Table(eTSVTb_VolAsk_I, i), Table(eTSVTb_VolOther_I, i), Table(eTSVTb_VolRow_Total, i), .Rows - 1
            End If
            
            If Table(eTSVTb_VolBid_J, i) > 0 Or Table(eTSVTb_VolAsk_J, i) Or Table(eTSVTb_VolOther_J, i) > 0 Then
                If Table(eTSVTb_VolRow_Total, i) > dLargestVol Then dLargestVol = Table(eTSVTb_VolRow_Total, i)
                SetGridRow eTSVTb_VolBid_J, Table(eTSVTb_VolBid_J, i), Table(eTSVTb_VolAsk_J, i), Table(eTSVTb_VolOther_J, i), Table(eTSVTb_VolRow_Total, i), .Rows - 1
            End If
            
            If Table(eTSVTb_VolBid_K, i) > 0 Or Table(eTSVTb_VolAsk_K, i) Or Table(eTSVTb_VolOther_K, i) > 0 Then
                If Table(eTSVTb_VolRow_Total, i) > dLargestVol Then dLargestVol = Table(eTSVTb_VolRow_Total, i)
                SetGridRow eTSVTb_VolBid_K, Table(eTSVTb_VolBid_K, i), Table(eTSVTb_VolAsk_K, i), Table(eTSVTb_VolOther_K, i), Table(eTSVTb_VolRow_Total, i), .Rows - 1
            End If
            
            If Table(eTSVTb_VolBid_L, i) > 0 Or Table(eTSVTb_VolAsk_L, i) Or Table(eTSVTb_VolOther_L, i) > 0 Then
                If Table(eTSVTb_VolRow_Total, i) > dLargestVol Then dLargestVol = Table(eTSVTb_VolRow_Total, i)
                SetGridRow eTSVTb_VolBid_L, Table(eTSVTb_VolBid_L, i), Table(eTSVTb_VolAsk_L, i), Table(eTSVTb_VolOther_L, i), Table(eTSVTb_VolRow_Total, i), .Rows - 1
            End If
            
            If Table(eTSVTb_VolBid_M, i) > 0 Or Table(eTSVTb_VolAsk_M, i) Or Table(eTSVTb_VolOther_M, i) > 0 Then
                If Table(eTSVTb_VolRow_Total, i) > dLargestVol Then dLargestVol = Table(eTSVTb_VolRow_Total, i)
                SetGridRow eTSVTb_VolBid_M, Table(eTSVTb_VolBid_M, i), Table(eTSVTb_VolAsk_M, i), Table(eTSVTb_VolOther_M, i), Table(eTSVTb_VolRow_Total, i), .Rows - 1
            End If
            
            If Table(eTSVTb_VolBid_N, i) > 0 Or Table(eTSVTb_VolAsk_N, i) Or Table(eTSVTb_VolOther_N, i) > 0 Then
                If Table(eTSVTb_VolRow_Total, i) > dLargestVol Then dLargestVol = Table(eTSVTb_VolRow_Total, i)
                SetGridRow eTSVTb_VolBid_N, Table(eTSVTb_VolBid_N, i), Table(eTSVTb_VolAsk_N, i), Table(eTSVTb_VolOther_N, i), Table(eTSVTb_VolRow_Total, i), .Rows - 1
            End If
        Next
        
'color green/yellow bars
        If m.Data.PriceVolLastDataCol >= eTSVTb_VolBid_B Then ColorGridRow eTSV_Group_AB
        If m.Data.PriceVolLastDataCol >= eTSVTb_VolBid_E Then ColorGridRow eTSV_Group_AE
        If m.Data.PriceVolLastDataCol >= eTSVTb_VolBid_I Then ColorGridRow eTSV_Group_AI
        If m.Data.PriceVolLastDataCol >= eTSVTb_VolBid_N Then ColorGridRow eTSV_Group_AN
    End With
    
    Exit Sub

ErrSection:
    RaiseError "frmPriceVol.LoadGridBidAsk"

End Sub

Private Sub SetAuctionBarCol(tbAuction As cGdTable, ByVal iCol&, ByVal iSepRow&)
On Error GoTo ErrSection:

    Dim eColBid As eTSV_TbFields
    Dim iHigh&, iLow&, iClose&
    Dim iAbove&, iMean&, iBelow&
    Dim iHiestVol&
        
    If tbAuction Is Nothing Then Exit Sub               'precautionary
    
    With fgPriceVol
        iHigh = tbAuction(iCol, eTSVTb_Idx_Hi) + .FixedRows
        iLow = tbAuction(iCol, eTSVTb_Idx_Low) + .FixedRows
        iClose = tbAuction(iCol, eTSVTb_Idx_Close) + .FixedRows
        iAbove = tbAuction(iCol, eTSVTb_Idx_Above) + .FixedRows
        iMean = tbAuction(iCol, eTSVTb_Idx_Mean) + .FixedRows
        iBelow = tbAuction(iCol, eTSVTb_Idx_Below) + .FixedRows
        iHiestVol = tbAuction(iCol, eTSVTb_Idx_HiestVol) + .FixedRows
        
        eColBid = m.Data.TbColVolBid(iCol)
        If eColBid <> -1 And eColBid < .Cols Then
            If iAbove - 1 >= .FixedRows And iAbove - 1 < iSepRow Then
                .Cell(flexcpFloodColor, iHigh, eColBid, iAbove - 1, eColBid) = m.iColorUnfairHigh
                .Cell(flexcpFloodPercent, iHigh, eColBid, iAbove - 1, eColBid) = 0
            End If
            
            If iAbove >= .FixedRows And iAbove < iSepRow Then
                .Cell(flexcpFloodColor, iAbove, eColBid, iMean - 1, eColBid) = m.iColorValue
                .Cell(flexcpFloodPercent, iAbove, eColBid, iMean - 1, eColBid) = 0
            ElseIf iHigh >= .FixedRows And iHigh < iSepRow Then
                .Cell(flexcpFloodColor, iHigh, eColBid, iMean - 1, eColBid) = m.iColorValue
                .Cell(flexcpFloodPercent, iHigh, eColBid, iMean - 1, eColBid) = 0
            End If
            
            If iBelow > .FixedRows And iBelow < iSepRow Then
                .Cell(flexcpFloodColor, iMean + 1, eColBid, iBelow, eColBid) = m.iColorValue
                .Cell(flexcpFloodPercent, iMean + 1, eColBid, iBelow, eColBid) = 0
            ElseIf iLow > .FixedRows And iLow < iSepRow Then
                .Cell(flexcpFloodColor, iMean + 1, eColBid, iLow, eColBid) = m.iColorValue
                .Cell(flexcpFloodPercent, iMean + 1, eColBid, iLow, eColBid) = 0
            End If
            
            If iBelow + 1 >= .FixedRows And iBelow + 1 < iSepRow Then
                .Cell(flexcpFloodColor, iBelow + 1, eColBid, iLow, eColBid) = m.iColorUnfairLow
                .Cell(flexcpFloodPercent, iBelow + 1, eColBid, iLow, eColBid) = 0
            End If
            
            If iHiestVol > .FixedRows And iHiestVol < iSepRow Then
                .Cell(flexcpFloodColor, iHiestVol, eColBid, iHiestVol, eColBid) = m.iColorMode
                .Cell(flexcpFloodPercent, iHiestVol, eColBid, iHiestVol, eColBid) = 0
            End If
            
            If iMean > .FixedRows And iMean < iSepRow Then
                .Cell(flexcpFloodColor, iMean, eColBid, iMean, eColBid) = m.iColorMean
                .Cell(flexcpFloodPercent, iMean, eColBid, iMean, eColBid) = 0
            End If
        End If
    End With
    
    Exit Sub

ErrSection:
    RaiseError "frmPriceVol.SetAuctionBarCol"
    
End Sub

Private Sub LoadGridAuctionBar(Table As cGdTable, Bars As cGdBars, dLargestVol As Double)
On Error GoTo ErrSection:

    Dim tbAuction As cGdTable
    
    Dim iHiestVol&, i&
    
    Set tbAuction = m.Data.AuctionTable
    
    If Table Is Nothing Or Bars Is Nothing Or tbAuction Is Nothing Then
        Exit Sub           'precautionary
    End If
        
    dLargestVol = 0
    With fgPriceVol
        .Redraw = flexRDNone
        .Rows = .FixedRows + Table.NumRecords
        For i = 0 To Table.NumRecords - 1
            If m.Data.LineToolColor(i) > 0 Then
                .Cell(flexcpBackColor, i + .FixedRows, 1, i + .FixedRows, .Cols - 1) = m.Data.LineToolColor(i)
            End If
            iHiestVol = Table(eTSVTb_VolRow_Total, i)
            .TextMatrix(i + .FixedRows, 0) = Bars.PriceDisplay(Table(eTSVTb_Price, i))
            .TextMatrix(i + .FixedRows, 29) = ""
            If iHiestVol > dLargestVol Then
                dLargestVol = iHiestVol
            End If
        Next
        For i = eTSVTb_Col_A To m.Data.StatsLastDataCol
            SetAuctionBarCol tbAuction, i, .Rows
        Next
        SetAuctionBarCol tbAuction, eTSVTb_Col_Total, .Rows
        .Redraw = flexRDBuffered
    End With
    
    Exit Sub

ErrSection:
    RaiseError "frmPriceVol.LoadGridAuctionBar"

End Sub

Private Sub SetSumGridData(ByVal iColStart&, ByVal iColEnd&, Optional ByVal iGridCol As Long = 1)
On Error GoTo ErrSection:

    Dim Table As cGdTable
    Dim Bars As cGdBars
    
    Dim iSpace As Long
    Dim i&, j&
    
    Set Table = m.Data.StatsTable
    Set Bars = m.Data.TickBars
    
    If Table Is Nothing Or Bars Is Nothing Then
        Exit Sub            'precautionary
    ElseIf Table.NumRecords = 0 Or Bars.Size = 0 Then
        Exit Sub            'no data for current session
    End If

    With fgSummary
        .Redraw = flexRDNone
        j = iGridCol       'j & j+1 are merged columns
'per-column sums & stats
        For i = iColStart To iColEnd
            iSpace = MergeColSpace(j)
            If j + 1 > .FixedCols And j + 1 < .Cols Then
                'bid, ask, [+ -]
                .MergeRow(eSepOffset_BidVol) = True
                .TextMatrix(eSepOffset_BidVol, j) = Table(i, eTSVTb_Row_Bid) & Space(iSpace)
                .TextMatrix(eSepOffset_BidVol, j + 1) = .TextMatrix(eSepOffset_BidVol, j)
    
                .TextMatrix(eSepOffset_AskVol, j) = Table(i, eTSVTb_Row_Ask) & Space(iSpace)
                .TextMatrix(eSepOffset_AskVol, j + 1) = .TextMatrix(eSepOffset_AskVol, j)
    
                .TextMatrix(eSepOffset_DeltaBidAsk, j) = Table(i, eTSVTb_Row_Delta) & Space(iSpace)
                .TextMatrix(eSepOffset_DeltaBidAsk, j + 1) = .TextMatrix(eSepOffset_DeltaBidAsk, j)
    
                '[+ -]Cum, %
                .TextMatrix(eSepOffSet_DeltaCum, j) = Table(i, eTSVTb_Row_DeltaCum) & Space(iSpace)
                .TextMatrix(eSepOffSet_DeltaCum, j + 1) = .TextMatrix(eSepOffSet_DeltaCum, j)
    
                .TextMatrix(eSepOffset_Percent, j) = Table(i, eTSVTb_Row_Percent) & "%" & Space(iSpace)
                .TextMatrix(eSepOffset_Percent, j + 1) = .TextMatrix(eSepOffset_Percent, j)
    
                If Table(i, eTSVTb_Row_DeltaCum) > 0 Then
                    'this is cumulative diff(bid, ask)
                    'if > 0 then cumulative bid > cumulative ask --> color with bid color else color with ask color
                    .Cell(flexcpFloodColor, eSepOffSet_DeltaCum, j, eSepOffset_Percent, j + 1) = m.iColorBid
                    .Cell(flexcpFloodPercent, eSepOffSet_DeltaCum, j, eSepOffset_Percent, j + 1) = 100
                ElseIf Table(i, eTSVTb_Row_DeltaCum) < 0 Then
                    .Cell(flexcpFloodColor, eSepOffSet_DeltaCum, j, eSepOffset_Percent, j + 1) = m.iColorAsk
                    .Cell(flexcpFloodPercent, eSepOffSet_DeltaCum, j, eSepOffset_Percent, j + 1) = 100
                End If

                'VAM, Mean, VBM
                .TextMatrix(eSepOffSet_VAM, j) = Table(i, eTSVTb_Row_VAM) & Space(iSpace)
                .TextMatrix(eSepOffSet_VAM, j + 1) = .TextMatrix(eSepOffSet_VAM, j)
    
                .TextMatrix(eSepOffset_Mean, j) = Bars.PriceDisplay(Table(i, eTSVTb_Row_Mean)) & Space(iSpace)
                .TextMatrix(eSepOffset_Mean, j + 1) = .TextMatrix(eSepOffset_Mean, j)
    
                .TextMatrix(eSepOffset_VBM, j) = Table(i, eTSVTb_Row_VBM) & Space(iSpace)
                .TextMatrix(eSepOffset_VBM, j + 1) = .TextMatrix(eSepOffset_VBM, j)
    
                'total vol, avg vol, SH vol
                .TextMatrix(eSepOffset_Totals, j) = Table(i, eTSVTb_Row_VolTotal) & Space(iSpace)
                .TextMatrix(eSepOffset_Totals, j + 1) = .TextMatrix(eSepOffset_Totals, j)
    
                .TextMatrix(eSepOffset_AvgVol, j) = Table(i, eTSVTb_Row_AvgVol) & Space(iSpace)
                .TextMatrix(eSepOffset_AvgVol, j + 1) = .TextMatrix(eSepOffset_AvgVol, j)
    
                .TextMatrix(eSepOffset_SHVol, j) = Table(i, eTSVTb_Row_SHVol) & Space(iSpace)
                .TextMatrix(eSepOffset_SHVol, j + 1) = .TextMatrix(eSepOffset_SHVol, j)
    
                'range, avg range, SH range
                .TextMatrix(eSepOffset_Range, j) = Bars.PriceDisplay(Table(i, eTSVTb_Row_Range)) & Space(iSpace)
                .TextMatrix(eSepOffset_Range, j + 1) = .TextMatrix(eSepOffset_Range, j)
    
                .TextMatrix(eSepOffset_AvgRange, j) = Bars.PriceDisplay(Table(i, eTSVTb_Row_AvgRange)) & Space(iSpace)
                .TextMatrix(eSepOffset_AvgRange, j + 1) = .TextMatrix(eSepOffset_AvgRange, j)
    
                .TextMatrix(eSepOffset_SHRange, j) = Bars.PriceDisplay(Table(i, eTSVTb_Row_SHRange)) & Space(iSpace)
                .TextMatrix(eSepOffset_SHRange, j + 1) = .TextMatrix(eSepOffset_SHRange, j)
            End If
            j = j + 2
        Next
        
'stats for right-most total column
        'bidvol , askvol, [+ -]
        .TextMatrix(eSepOffset_BidVol, 29) = Table(eTSVTb_Col_Total, eTSVTb_Row_Bid)
        .TextMatrix(eSepOffset_AskVol, 29) = Table(eTSVTb_Col_Total, eTSVTb_Row_Ask)
        .TextMatrix(eSepOffset_DeltaBidAsk, 29) = Table(eTSVTb_Col_Total, eTSVTb_Row_Delta)
        '[+ -]Cum, percent
        .TextMatrix(eSepOffSet_DeltaCum, 29) = Table(eTSVTb_Col_Total, eTSVTb_Row_DeltaCum)
        .TextMatrix(eSepOffset_Percent, 29) = Table(eTSVTb_Col_Total, eTSVTb_Row_Percent) & "%"
        If Table(eTSVTb_Col_Total, eTSVTb_Row_DeltaCum) > 0 Then
            'this is cumulative diff(bid, ask)
            'if > 0 then cumulative bid > cumulative ask --> color with bid color else color with ask color
            .Cell(flexcpFloodColor, eSepOffSet_DeltaCum, 29, eSepOffset_Percent, 29) = m.iColorBid
            .Cell(flexcpFloodPercent, eSepOffSet_DeltaCum, 29, eSepOffset_Percent, 29) = 100
        ElseIf Table(eTSVTb_Col_Total, eTSVTb_Row_DeltaCum) < 0 Then
            .Cell(flexcpFloodColor, eSepOffSet_DeltaCum, 29, eSepOffset_Percent, 29) = m.iColorAsk
            .Cell(flexcpFloodPercent, eSepOffSet_DeltaCum, 29, eSepOffset_Percent, 29) = 100
        End If
        'VAM, Mean, VBM
        .TextMatrix(eSepOffSet_VAM, 29) = Table(eTSVTb_Col_Total, eTSVTb_Row_VAM)
        .TextMatrix(eSepOffset_Mean, 29) = Bars.PriceDisplay(Table(eTSVTb_Col_Total, eTSVTb_Row_Mean))
        .TextMatrix(eSepOffset_VBM, 29) = Table(eTSVTb_Col_Total, eTSVTb_Row_VBM)
        'totals, avgVol, SHVol
        .TextMatrix(eSepOffset_Totals, 29) = Table(eTSVTb_Col_Total, eTSVTb_Row_VolTotal)
        .TextMatrix(eSepOffset_AvgVol, 29) = Table(eTSVTb_Col_Total, eTSVTb_Row_AvgVol)
        .TextMatrix(eSepOffset_SHVol, 29) = Table(eTSVTb_Col_Total, eTSVTb_Row_SHVol)
        'range, avgRange, SHRange
        .TextMatrix(eSepOffset_Range, 29) = Bars.PriceDisplay(Table(eTSVTb_Col_Total, eTSVTb_Row_Range))
        .TextMatrix(eSepOffset_AvgRange, 29) = Bars.PriceDisplay(Table(eTSVTb_Col_Total, eTSVTb_Row_AvgRange))
        .TextMatrix(eSepOffset_SHRange, 29) = Bars.PriceDisplay(Table(eTSVTb_Col_Total, eTSVTb_Row_SHRange))
        
        .Redraw = flexRDBuffered
    End With
    
    Set Table = Nothing
    
    Exit Sub

ErrSection:
    RaiseError "frmPriceVol.SetSumGridData"

End Sub

Private Sub LoadGrid()
On Error GoTo ErrSection:

    Dim Table As cGdTable
    Dim Bars As cGdBars
    
    Dim strIcon$, iIconCol&, iIconType&, iIconColor&, iIconAlign&
    Dim dTotal#, dLargestVol#
    Dim dHigh#, dLow#
    Dim i&, j&
        
    Dim bHideCol As Boolean
    
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
    
    If m.eView = eTSVView_VolAll Or m.eView = eTSVView_AuctionBar Then bHideCol = True
    
    'bid cols are odd numbered col beginning with col 1
    'ask cols are even numbered col beginning with col 2
    fgSummary.Redraw = flexRDNone
    fgSummary.ColWidth(0) = 950
    With fgPriceVol
        .Redraw = flexRDNone
        .ColWidth(0) = 950
        For i = 2 To 28 Step 2
            .ColHidden(i) = bHideCol
            .ColWidth(i - 1) = 850
            .ColWidth(i) = 850
            With fgSummary
                .ColHidden(i) = bHideCol
                .ColWidth(i - 1) = 850
                .ColWidth(i) = 850
            End With
        Next
        .ColWidth(29) = 850
        .Rows = .FixedRows
        .Redraw = flexRDBuffered
    End With
    fgSummary.ColWidth(29) = 850
    fgSummary.Redraw = flexRDBuffered
    
    dLargestVol = 0
    m.dLastUpdated = 0
    If Not frmQuotes Is Nothing Then
        m.iCurrPriceColor = frmQuotes.UnchColor
    Else
        m.iCurrPriceColor = fgPriceVol.ForeColor
    End If
    
    With fgPriceVol
        .Redraw = flexRDNone

        If m.eView = eTSVView_AuctionBar Then
            LoadGridAuctionBar Table, Bars, dLargestVol
        Else
            LoadGridBidAsk Table, Bars, dLargestVol
        End If
        
        If Not bHideCol Then
            For i = 2 To 28 Step 2
                .Cell(flexcpAlignment, .FixedRows, i, .Rows - 1, i) = flexAlignLeftCenter
            Next
        End If
                
        'color price column for current values of green, yellow bars & high/low
        ColorGridRow eTSV_Group_Current

'flood histogram
        dLargestVol = dLargestVol * 1.1
        m.Data.HighLow dHigh, dLow
        For i = .FixedRows To .Rows - 1
            dTotal = Table(eTSVTb_VolRow_Total, i - .FixedRows)
            .Cell(flexcpFloodColor, i, 30) = m.iColorHistogram
            
            If dLargestVol > 0 Then
                .Cell(flexcpFloodPercent, i, 30) = (dTotal / dLargestVol * 100)
            Else
                .Cell(flexcpFloodPercent, i, 30) = 0
            End If
            
            If ValOfText(.TextMatrix(i, 0)) = dHigh Then
                .Cell(flexcpBackColor, i, 0, i, 0) = m.iColorAsk
            ElseIf ValOfText(.TextMatrix(i, 0)) = dLow Then
                .Cell(flexcpBackColor, i, 0, i, 0) = m.iColorBid
            End If
            
             'draw icons if any
            For j = 0 To m.Data.IconCount - 1
                strIcon = m.Data.IconString(j)
                If Len(strIcon) > 0 Then
                    'format: ICON|key|icon|price|column|color|alignment (key = icon type concat with number)
                    If Parse(strIcon, "|", 4) = .TextMatrix(i, 0) Then
                        iIconCol = Val(Parse(strIcon, "|", 5))
                        If iIconCol >= .FixedCols And iIconCol < .Cols Then
                            iIconType = frmFootprintIcons.IconTypeNum(Parse(strIcon, "|", 3))
                            iIconColor = Val(Parse(strIcon, "|", 6))
                            iIconAlign = Val(Parse(strIcon, "|", 7))
                            geFootprintIcon IconPic.hDC, iIconType, vbWhite, iIconColor, "FpIcon.bmp"
                            If FileExist("FpIcon.bmp") Then
                                IconPic.Picture = LoadPicture("FpIcon.bmp")
                                .Cell(flexcpPicture, i, iIconCol) = IconPic.Picture
                                .Cell(flexcpPictureAlignment, i, iIconCol) = iIconAlign
                                .Cell(flexcpData, i, iIconCol) = Parse(strIcon, "|", 2)
                            End If
                        End If
                    End If
                End If
            Next
       Next
        m.dHistogramMaxVol = dLargestVol
        
        SetSumGridData eTSVTb_Col_A, m.Data.StatsLastDataCol
                        
        'hide cols that do not yet have data
        fgSummary.Redraw = flexRDNone
        For i = m.Data.PriceVolLastDataCol + 2 To eTSVTb_VolAsk_N
            .ColHidden(i) = True
            fgSummary.ColHidden(i) = True
        Next
        If m.eStatsView = eTSVView_StatsNone Then
            fgSummary.Visible = False
        Else
            m.iFgSummaryHeight = 0
            With fgSummary
                If Not .Visible Then .Visible = True
                For i = 0 To .Rows - 1
                    If m.eStatsView = eTSVView_StatsAll Then
                        .RowHidden(i) = False
                    ElseIf (i >= 0 And i < 6) Or i = 12 Then
                        .RowHidden(i) = False
                    Else
                        .RowHidden(i) = True
                    End If
                    If Not .RowHidden(i) Then m.iFgSummaryHeight = m.iFgSummaryHeight + .RowHeight(i)
                Next
            End With
        End If
        fgSummary.Redraw = flexRDBuffered
                
        'draw boxes if any
        Dim iTop&, iLeft&, iBottom&, iRight&
        Dim iColor&, iPix&, iFill&
        Dim eBoxView As ePriceVol_View
        
        For i = 0 To m.Data.BoxCount - 1
            m.Data.BoxRect i + 1, iTop, iLeft, iBottom, iRight
            iTop = iTop + .FixedRows
            iBottom = iBottom + .FixedRows
            iPix = m.Data.BoxThickness(i + 1)
            iColor = m.Data.BoxColor(i + 1)
            iFill = m.Data.BoxFill(i + 1)
            eBoxView = m.Data.BoxView(i + 1)
            DrawBox iTop, iLeft, iBottom, iRight, iColor, iPix, eBoxView, iFill
        Next
                
        'bold last price
        If Not g.RealTime.Active Then
            If m.iLastPriceRow > .FixedRows And m.iLastPriceRow < .Rows Then
                .Cell(flexcpFontBold, m.iLastPriceRow, 0) = True
                If m.eView = eTSVView_AuctionBar Then
                    .Cell(flexcpForeColor, m.iLastPriceRow, 0) = m.iCurrPriceColor
                Else
                    i = m.Data.PriceVolLastDataCol
                    If i > .FixedCols And i < .Cols Then .Cell(flexcpFontBold, m.iLastPriceRow, i) = True
                    i = i + 1
                    If i > .FixedCols And i < .Cols Then .Cell(flexcpFontBold, m.iLastPriceRow, i) = True
                    .Cell(flexcpForeColor, m.iLastPriceRow, 0, m.iLastPriceRow, i) = m.iCurrPriceColor
                End If
            End If
        End If
        
        .Redraw = flexRDBuffered
    End With
    
    Set Table = Nothing
    Set Bars = Nothing
        
    Form_Resize             'need to call here to get correct number of rows in grid for centering price
    FocusGrid

    m.bInitInprog = False
    
    Exit Sub
    
ErrSection:
    RaiseError "frmPriceVol.LoadGrid"

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
        
        Set m.Data = New cTSVData
        m.Data.ResetTickBars m.strSym, m.nSymID, Nothing, 570
        m.Data.BlankRows = m.iBlankRows
        m.Data.BuildTables eTSVTb_TbType_VolAtPrice
        
        LoadGrid

        Me.Caption = kCaptionBase & " for " & strSym & " (" & DateFormat(m.Data.SessionDate, MM_DD_YYYY, NO_TIME) & ")"
    End If
    m.bInitInprog = False
    
    Exit Sub

ErrSection:
    RaiseError "frmPriceVol.LoadNewSymData"
    
End Sub

Private Sub InitGrid()
On Error GoTo ErrSection:
    
    Dim i&
    
    m.bInitInprog = True
    
    If m.iAuctionBarPix <= 0 Then m.iAuctionBarPix = 16     'precautionary
    If m.iTrianglePix <= 0 Then m.iTrianglePix = 10
    
    With fgSummary
        .Redraw = flexRDNone
        .FixedCols = 0
        .FixedRows = 0
        .Editable = flexEDNone
        .HighLight = flexHighlightNever
        .ExtendLastCol = True
    
        .BorderStyle = flexBorderNone
        .GridLines = flexGridNone
        .ScrollBars = flexScrollBarNone
        .MergeCells = flexMergeFree
        .FrozenCols = 1
        
        .Rows = 19
        .Cols = 31
                
         'top row (this is to have darker line between price & totals)
        .Cell(flexcpBackColor, 0, 0, 0, .Cols - 1) = RGB(128, 128, 128)
        .RowHeight(0) = 25
       
        'bid/ask vol
        .TextMatrix(1, 0) = "Sell Volume"
        .TextMatrix(2, 0) = "Buy Volume"
        .TextMatrix(3, 0) = "[+ -]"
        'separating row
        .Cell(flexcpBackColor, 4, 0, 4, .Cols - 1) = RGB(128, 128, 128)
        .RowHeight(4) = 25
        
        'cumulative bid/ask vol
        .TextMatrix(5, 0) = "[+ -] Cum"
        .TextMatrix(6, 0) = "Percent"
        'separating row
        .Cell(flexcpBackColor, 7, 0, 7, .Cols - 1) = RGB(128, 128, 128)
        .RowHeight(7) = 25
        
        'averages
        .TextMatrix(8, 0) = "VAM"
        .TextMatrix(9, 0) = "Mean"
        .TextMatrix(10, 0) = "VBM"
        'separating row
        .Cell(flexcpBackColor, 11, 0, 11, .Cols - 1) = RGB(128, 128, 128)
        .RowHeight(11) = 25
        
        'totals
        .TextMatrix(12, 0) = "Totals"
        .TextMatrix(13, 0) = "Avg Volume"
        .TextMatrix(14, 0) = "SH Volume"
        .Cell(flexcpBackColor, 15, 0, 15, .Cols - 1) = RGB(128, 128, 128)
        .RowHeight(15) = 25
        
        'ranges
        .TextMatrix(16, 0) = "Range"
        .TextMatrix(17, 0) = "Avg Range"
        .TextMatrix(18, 0) = "SH Range"
        
        .Height = .RowHeight(0)
        .MergeRow(0) = True
        For i = 1 To .Rows - 1
            .Height = .Height + .RowHeight(i)
            .MergeRow(i) = True
        Next
        .Cell(flexcpAlignment, 0, 1, .Rows - 1, 28) = flexAlignCenterCenter
        .Cell(flexcpAlignment, 0, 29, .Rows - 1, 29) = flexAlignRightCenter
        
        m.iFgSummaryHeight = .Height
    End With
    
    With fgPriceVol
        .Redraw = flexRDNone
        .FixedCols = 0
        .FixedRows = 1
        .Editable = flexEDNone
        .HighLight = flexHighlightNever
        .ExtendLastCol = True
        .OwnerDraw = flexODOver
        .BorderStyle = flexBorderNone
        .GridLines = flexGridNone
        .ScrollBars = flexScrollBarBoth
        
        .Rows = .FixedRows
        .Cols = 31
        
        .MergeRow(0) = True
        .MergeCells = flexMergeFree
        .FrozenCols = 1
                                
        'header row 1
        .TextMatrix(0, 0) = "Price"
        
        .TextMatrix(0, 1) = "A"
        .TextMatrix(0, 2) = "A"
        
        .TextMatrix(0, 3) = "B"
        .TextMatrix(0, 4) = "B"
        
        .TextMatrix(0, 5) = "C"
        .TextMatrix(0, 6) = "C"
        
        .TextMatrix(0, 7) = "D"
        .TextMatrix(0, 8) = "D"
        
        .TextMatrix(0, 9) = "E"
        .TextMatrix(0, 10) = "E"
        
        .TextMatrix(0, 11) = "F"
        .TextMatrix(0, 12) = "F"
        
        .TextMatrix(0, 13) = "G"
        .TextMatrix(0, 14) = "G"
        
        .TextMatrix(0, 15) = "H"
        .TextMatrix(0, 16) = "H"
        
        .TextMatrix(0, 17) = "I"
        .TextMatrix(0, 18) = "I"
        
        .TextMatrix(0, 19) = "J"
        .TextMatrix(0, 20) = "J"
        
        .TextMatrix(0, 21) = "K"
        .TextMatrix(0, 22) = "K"
        
        .TextMatrix(0, 23) = "L"
        .TextMatrix(0, 24) = "L"
        
        .TextMatrix(0, 25) = "M"
        .TextMatrix(0, 26) = "M"
        
        .TextMatrix(0, 27) = "N"
        .TextMatrix(0, 28) = "N"
        
        .TextMatrix(0, 29) = "Total"
        .TextMatrix(0, 30) = "Histogram"

        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        .Redraw = flexRDBuffered
    End With
    
    m.bInitInprog = False
    
    Exit Sub
        
ErrSection:
    RaiseError "frmPriceVol.InitGrid"

End Sub

Private Sub fgPriceVol_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
On Error Resume Next

    If g.bUnloading Then Exit Sub

    If Not m.bInitInprog Then
        m.iTopRow = NewTopRow
        m.iLeftCol = NewLeftCol
        m.bKeepAtEnd = fgPriceVol.ColIsVisible(fgPriceVol.Cols - 1)
        fgSummary.LeftCol = NewLeftCol
    End If
    
End Sub

Private Sub fgPriceVol_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
On Error Resume Next

    Dim i&, j&, iMid&, iMax&, iColor&
    Dim tbAuction As cGdTable
    
    If m.bInitInprog Or m.bReloadData Then Exit Sub
    
    If m.eView = eTSVView_AuctionBar Then
        With fgPriceVol
            If .Rows > .FixedRows Then
            iMax = Abs((.Cell(flexcpWidth, .FixedRows, 1) - .Cell(flexcpLeft, .FixedRows, 1)) * 0.7)
            If m.iAuctionBarPix > iMax Then m.iAuctionBarPix = iMax     'precautionary
            
            If Col > 0 And Col < 30 And Row >= .FixedRows And Row < .Rows - 1 Then
                Set tbAuction = m.Data.AuctionTable
                i = m.Data.TbColStats(Col)
                If i <> -1 And Not tbAuction Is Nothing Then
                    iMid = Right - ((Right - Left) / 2)
                    If .Cell(flexcpFloodColor, Row, Col) > 0 Then
                        i = tbAuction(i, eTSVTb_Idx_Close)
                        If i + .FixedRows = Row Then
                            i = iMid + m.iAuctionBarPix / 2 + 2     'start triangle to right of rectangle
                            j = i + m.iTrianglePix
                            i = geDrawTickTriangle(0, hDC, m.iColorTriangle, m.iColorTriangle, 2, Top - 1, i, Bottom + 1, j)
                        End If
                        i = iMid - m.iAuctionBarPix / 2
                        j = iMid + m.iAuctionBarPix / 2
                        iColor = .Cell(flexcpFloodColor, Row, Col)
                        i = geDrawTickTriangle(0, hDC, iColor, iColor, 4, Top, i, Bottom, j)
                    End If
                End If
            End If
            End If
        End With
    End If

End Sub

Private Function InRangeCol(ByVal iCol&) As Boolean
On Error GoTo ErrSection:

    With fgPriceVol
        If iCol > .FixedCols And iCol < .Cols Then InRangeCol = True
    End With
    
    Exit Function
    
ErrSection:
    RaiseError "frmPriceVol.InRangeCol"

End Function

Private Function InRangeRow(ByVal iRow&) As Boolean
On Error GoTo ErrSection:
    
    With fgPriceVol
        If iRow >= .FixedRows And iRow < .Rows - 1 Then InRangeRow = True
    End With
    
    Exit Function
    
ErrSection:
    RaiseError "frmPriceVol.InRangeRow"

End Function

Private Sub DrawBox(ByVal iTop&, ByVal iLeft&, ByVal iBottom&, ByVal iRight&, _
    ByVal iColor&, ByVal iPix&, ByVal eBoxView As ePriceVol_View, ByVal iFill&)
On Error GoTo ErrSection:
    
    Dim iCol&, iRow&
    Dim dBid#, dAsk#
    
    Dim Table As cGdTable
    
    If InRangeRow(iTop) And InRangeRow(iBottom) Then
        If InRangeCol(iLeft) And InRangeCol(iRight) Then
            
            If eBoxView <> m.eView Then
                If eBoxView = eTSVView_AuctionBar Or eBoxView = eTSVView_VolAll Then
                    If m.eView = eTSVView_VolBidAsk Or m.eView = eTSVView_VolDelta Then
                        If iLeft > iRight Then
                            iLeft = iLeft + 1
                        Else
                            iRight = iRight + 1
                        End If
                    End If
                Else
                    If m.eView = eTSVView_AuctionBar Or m.eView = eTSVView_VolAll Then
                        If iLeft Mod 2 = 0 Then iLeft = iLeft - 1
                        If iRight Mod 2 = 0 Then iRight = iRight - 1
                    End If
                End If
            End If
                        
            With fgPriceVol
                .Select iTop, iLeft, iBottom, iRight
                .CellBorder iColor, iPix, iPix, iPix, iPix, 0, 0
                .Select 0, 0
                
                If iLeft > iRight Then
                    iCol = iLeft
                    iLeft = iRight
                    iRight = iCol
                End If
                
                If iTop > iBottom Then
                    iRow = iTop
                    iTop = iBottom
                    iBottom = iRow
                End If
                
                If m.eView = eTSVView_AuctionBar Then
                    For iCol = iLeft To iRight Step 2
                        For iRow = iTop To iBottom
                            If iFill = 1 Then
                                If .Cell(flexcpBackColor, iRow, iCol) = .BackColor Or .Cell(flexcpBackColor, iRow, iCol) = 0 Then
                                    .Cell(flexcpBackColor, iRow, iCol) = iColor
                                    .Cell(flexcpBackColor, iRow, iCol + 1) = iColor
                                End If
                            ElseIf Not IsPreDefinedColor(.Cell(flexcpBackColor, iRow, iCol)) Then
                                If .Cell(flexcpBackColor, iRow, iCol) = m.iBoxColor Then
                                    .Cell(flexcpBackColor, iRow, iCol) = .BackColor     'clear box filled color
                                    .Cell(flexcpBackColor, iRow, iCol + 1) = .BackColor
                                End If
                            End If
                        Next
                    Next
                Else
                    Set Table = m.Data.PriceVolTable
                                        
                    If iLeft Mod 2 = 0 Then iLeft = iLeft - 1
                    
                    For iCol = iLeft To iRight Step 2
                        For iRow = iTop To iBottom
                            dBid = Table(iCol, iRow - .FixedRows)
                            dAsk = Table(iCol + 1, iRow - .FixedRows)
                            DrawStatsBlueBox iCol, dBid, dAsk, iRow
                            If iFill = 1 Then
                                If .Cell(flexcpBackColor, iRow, iCol) = .BackColor Or .Cell(flexcpBackColor, iRow, iCol) = 0 Then
                                    .Cell(flexcpBackColor, iRow, iCol) = iColor
                                    .Cell(flexcpBackColor, iRow, iCol + 1) = iColor
                                End If
                            ElseIf Not IsPreDefinedColor(.Cell(flexcpBackColor, iRow, iCol)) Then
                                If .Cell(flexcpBackColor, iRow, iCol) = m.iBoxColor Then
                                    .Cell(flexcpBackColor, iRow, iCol) = .BackColor     'clear box filled color
                                    .Cell(flexcpBackColor, iRow, iCol + 1) = .BackColor
                                End If
                            End If
                        Next
                    Next
                End If
            End With
        
        End If
    End If
    
    Exit Sub

ErrSection:
    RaiseError "frmPriceVol.DrawBox"

End Sub

Private Sub RefreshBoxes()
On Error Resume Next

    'refresh filled boxes in case some cells got set to background color due to removal of line tool or RT update
    Dim iTop&, iLeft&, iBottom&, iRight&
    Dim iColor&, iPix&, iFill&, i&
    Dim eBoxView As ePriceVol_View
    
    For i = 0 To m.Data.BoxCount - 1
        m.Data.BoxRect i + 1, iTop, iLeft, iBottom, iRight
        iTop = iTop + fgPriceVol.FixedRows
        iBottom = iBottom + fgPriceVol.FixedRows
        iPix = m.Data.BoxThickness(i + 1)
        iColor = m.Data.BoxColor(i + 1)
        iFill = m.Data.BoxFill(i + 1)
        eBoxView = m.Data.BoxView(i + 1)
        If iFill = 1 Then
            DrawBox iTop, iLeft, iBottom, iRight, iColor, iPix, eBoxView, iFill
        End If
    Next

End Sub

Private Sub HandleBox(Button As Integer)
On Error GoTo ErrSection:

    Dim iTop&, iLeft&, iBottom&, iRight&
    Dim iBoxId&, iColor&, iPix&, iFill&
    Dim eBoxView As ePriceVol_View

    If m.bInitInprog Or m.bReloadData Or m.bTimerInProg Or g.bUnloading Then
        ResetMouseVar
        Exit Sub
    End If

    m.bEditDrawInProg = True
    
    With fgPriceVol
        iBoxId = m.Data.BoxId(m.iMouseRowDown - .FixedRows, m.iMouseColDown)
    End With
    
    If iBoxId > 0 Then
        m.Data.BoxRect iBoxId, iTop, iLeft, iBottom, iRight
        iTop = iTop + fgPriceVol.FixedRows
        iBottom = iBottom + fgPriceVol.FixedRows
        iColor = m.Data.BoxColor(iBoxId)
        iPix = m.Data.BoxThickness(iBoxId)
        iFill = m.Data.BoxFill(iBoxId)
        eBoxView = m.Data.BoxView(iBoxId)
        
        If Button = vbLeftButton Then
            If frmPriceVolCfg.ShowBoxSettings(Me, iColor, iPix, iFill) Then
                'editing existing box
                DrawBox iTop, iLeft, iBottom, iRight, m.iBoxColor, m.iBoxPix, eBoxView, m.iBoxFill
                m.Data.BoxColor(iBoxId) = m.iBoxColor
                m.Data.BoxThickness(iBoxId) = m.iBoxPix
                m.Data.BoxFill(iBoxId) = m.iBoxFill
            End If
        Else
            'clear/delete box
            m.iBoxColor = iColor
            DrawBox iTop, iLeft, iBottom, iRight, fgPriceVol.BackColor, 0, eBoxView, 0
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
                    With fgPriceVol
                        'add new box
                        DrawBox m.iMouseRowDown, m.iMouseColDown, m.iMouseRowDownPrev, m.iMouseColDownPrev, _
                                m.iBoxColor, m.iBoxPix, m.eView, m.iBoxFill
                        m.Data.AddBox m.iMouseRowDown - .FixedRows, m.iMouseColDown, _
                            m.iMouseRowDownPrev - .FixedRows, m.iMouseColDownPrev, _
                            m.iBoxColor, m.iBoxPix, m.eView, m.iBoxFill
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
    RaiseError "frmPriceVol.HandleBox"

End Sub

Private Sub fgPriceVol_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
        
    Dim strText$, strKey$, strIconType$
    Dim iIndex&
    
    If m.bInitInprog Or m.bReloadData Or m.bEditDrawInProg Then Exit Sub
        
    If tbToolbar.Tools("ID_Draw").State = ssChecked Then
        With fgPriceVol
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
    ElseIf tbToolbar.Tools("ID_Icons").State Then
        With fgPriceVol
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

Private Sub fgPriceVol_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    If m.bInitInprog Or m.bReloadData Then Exit Sub
    
    If tbToolbar.Tools("ID_Draw").State = ssChecked Or tbToolbar.Tools("ID_Icons").State = ssChecked Then
        With fgPriceVol
            If InRangeRow(.MouseRow) Then
                .MousePointer = flexCustom
                .MouseIcon = Picture16(ToolbarIcon("kPencil"))
            Else
                .MousePointer = flexDefault
            End If
        End With
    Else
        fgPriceVol.MousePointer = flexDefault
        ResetMouseVar
    End If

End Sub

Private Sub fgPriceVol_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    If m.bInitInprog Or m.bReloadData Then Exit Sub
    
    If tbToolbar.Tools("ID_Draw").State = ssChecked Then
        HandleBox Button
    Else
        ResetMouseVar
    End If

End Sub

Private Sub Form_Load()
On Error GoTo ErrSection:
    
    Me.Icon = Picture16(ToolbarIcon("ID_VolumeAtPrice"))
    
    g.Styler.StyleForm Me
    
    m.iAuctionBarPix = GetIniFileProperty("AuctionBarPix", 16, "IOAMT", g.strIniFile)
    m.iTrianglePix = GetIniFileProperty("TrianglePix", 10, "IOAMT", g.strIniFile)
    
    m.iColorAsk = GetIniFileProperty("ColorAsk", kAskColor, "IOAMT", g.strIniFile)
    m.iColorBid = GetIniFileProperty("ColorBid", kBidColor, "IOAMT", g.strIniFile)
    m.iColorHistogram = GetIniFileProperty("ColorHistogram", RGB(125, 255, 255), "IOAMT", g.strIniFile)
    m.iColorMean = GetIniFileProperty("ColorMean", vbGreen, "IOAMT", g.strIniFile)
    m.iColorMode = GetIniFileProperty("ColorMode", RGB(128, 128, 128), "IOAMT", g.strIniFile)
    m.iColorTriangle = GetIniFileProperty("ColorTriangle", RGB(128, 128, 128), "IOAMT", g.strIniFile)
    m.iColorUnfairHigh = GetIniFileProperty("ColorUnfairHigh", vbRed, "IOAMT", g.strIniFile)
    m.iColorUnfairLow = GetIniFileProperty("ColorUnfairLow", vbBlue, "IOAMT", g.strIniFile)
    m.iColorValue = GetIniFileProperty("ColorValue", vbYellow, "IOAMT", g.strIniFile)
    m.iBlankRows = GetIniFileProperty("PriceVolBlankRows", 10, "IOAMT", g.strIniFile)
    m.eStatsView = GetIniFileProperty("PriceVolStatsView", eTSVView_StatsAll, "IOAMT", g.strIniFile)
    
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
    RaiseError "frmPriceVol.Form_Load"

End Sub

Private Sub Form_Resize()
On Error Resume Next

    Dim iPriceVolHt&
        
    If Me.WindowState = vbMinimized Then
        Exit Sub
    End If
        
    If fgSummary.Visible Then
        fgSummary.Redraw = flexRDNone
        With fgPriceVol
            .Redraw = flexRDNone
            If .Rows > .FixedRows Then
                iPriceVolHt = .Rows * .RowHeight(1)
                If Me.ScaleHeight >= iPriceVolHt + m.iFgSummaryHeight Then
                    .Move 0, 0, Me.ScaleWidth, iPriceVolHt
                ElseIf Me.ScaleHeight - m.iFgSummaryHeight > 0 Then
                    .Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - m.iFgSummaryHeight
                Else
                    .Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
                End If
                fgSummary.Move 0, .Height, Me.ScaleWidth, m.iFgSummaryHeight
            Else
                fgSummary.Visible = False
                .Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - m.iFgSummaryHeight
            End If
            .Redraw = flexRDBuffered
        End With
        fgSummary.Redraw = flexRDBuffered
    Else
        With fgPriceVol
            .Redraw = flexRDNone
            If .MergeRow(.Rows - 1) = True Then
                .Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
            ElseIf .Rows > .FixedRows Then
                iPriceVolHt = (.Rows + 2) * .RowHeight(.Rows - 1)
                If Me.ScaleHeight > iPriceVolHt Then
                    .Move 0, 0, Me.ScaleWidth, iPriceVolHt
                Else
                    .Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
                End If
            Else
                .Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
            End If
            .Redraw = flexRDBuffered
        End With
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Set m.Data = Nothing
    
    frmFootprintIcons.CloseMe Me
    
    'save form size & location
    SetIniFileProperty Me.Name, GetFormPlacement(Me), "Placement", g.strIniFile
    SetIniFileProperty "PriceVolStatsView", m.eStatsView, "IOAMT", g.strIniFile

End Sub

Private Function IsPreDefinedColor(ByVal iColor&) As Boolean
On Error GoTo ErrSection:

    If iColor = m.iColorMean Or iColor = m.iColorValue Then
        IsPreDefinedColor = True
    Else
        IsPreDefinedColor = False
    End If
    
    Exit Function

ErrSection:
    RaiseError "frmPriceVol.IsPreDefinedColor"

End Function

Private Sub HandleHighlight(ByVal bClear As Boolean)
On Error GoTo ErrSection:
    
    Dim i&, iColor&
        
    If m.bInitInprog Or m.bReloadData Or m.bTimerInProg Or g.bUnloading Then
        ResetMouseVar
        Exit Sub
    End If
    
    m.bEditDrawInProg = True
    
    With fgPriceVol
        If InRangeRow(m.iMouseRowDown) Then
            If bClear Then
                iColor = .BackColor
                m.Data.LineToolColor(m.iMouseRowDown - .FixedRows) = 0
            Else
                iColor = gdColor.Color
                m.Data.LineToolColor(m.iMouseRowDown - .FixedRows) = gdColor.Color
            End If
            For i = 1 To .Cols - 1
                If Not IsPreDefinedColor(.Cell(flexcpBackColor, m.iMouseRowDown, i)) Then
                    .Cell(flexcpBackColor, m.iMouseRowDown, i) = iColor
                End If
            Next
            If bClear Then RefreshBoxes
        End If
    End With
    
    ResetMouseVar
    
    m.bEditDrawInProg = False
    
    Exit Sub

ErrSection:
    RaiseError "frmPriceVol.HandleHighlight"

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

End Sub

Private Sub tbToolbar_ComboCloseUp(ByVal Tool As ActiveToolBars.SSTool)
On Error GoTo ErrSection:

    Dim i&
    
    i = Tool.ComboBox.ListIndex
    
    If Tool.ID = "ID_View" Then
        If i <> m.eView Then
            m.eView = i
            LoadGrid
            FormResize Me
        End If
    ElseIf i <> m.eStatsView Then
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
                    ElseIf (i >= 0 And i < 6) Or i = 12 Then
                        .RowHidden(i) = False
                    Else
                        .RowHidden(i) = True
                    End If
                    If Not .RowHidden(i) Then m.iFgSummaryHeight = m.iFgSummaryHeight + .RowHeight(i)
                Next
            End With
        End If
        FormResize Me
    End If
    
    Exit Sub

ErrSection:
    RaiseError "frmPriceVol.tbToolbar_ComboCloseUp"

End Sub

Private Sub tbToolbar_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
On Error GoTo ErrSection:

    Dim astrSymbols As New cGdArray
    
    If m.bInitInprog Then Exit Sub
    
    If Tool.ID <> "ID_Draw" Then
        tbToolbar.Tools("ID_Draw").State = ssUnchecked
        ResetMouseVar
    End If
    
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
                With fgPriceVol
                    .Redraw = flexRDNone
                    .Rows = .FixedRows
                    .Redraw = flexRDBuffered
                End With
                With fgSummary
                    .Redraw = flexRDNone
                    .Select 0, 1, .Rows - 1, .Cols - 1
                    .Clear flexClearSelection, flexClearText
                    .Cell(flexcpFloodPercent, 0, 1, .Rows - 1, .Cols - 1) = 0
                    .Redraw = flexRDBuffered
                End With
                DoEvents        'to let grid redraw
                m.bInitInprog = False
                LoadNewSymData astrSymbols(0)
                FormResize Me
            End If
        Case "ID_Settings"
            If frmPriceVolCfg.ShowFormSettings(Me) Then
                m.bInitInprog = True
                                
                m.Data.BlankRows = m.iBlankRows
                LoadGrid
                
                SetIniFileProperty "AuctionBarPix", m.iAuctionBarPix, "IOAMT", g.strIniFile
                SetIniFileProperty "TrianglePix", m.iTrianglePix, "IOAMT", g.strIniFile
                
                SetIniFileProperty "ColorAsk", m.iColorAsk, "IOAMT", g.strIniFile
                SetIniFileProperty "ColorBid", m.iColorBid, "IOAMT", g.strIniFile
                SetIniFileProperty "ColorHistogram", m.iColorHistogram, "IOAMT", g.strIniFile
                SetIniFileProperty "ColorMean", m.iColorMean, "IOAMT", g.strIniFile
                SetIniFileProperty "ColorMode", m.iColorMode, "IOAMT", g.strIniFile
                SetIniFileProperty "ColorTriangle", m.iColorTriangle, "IOAMT", g.strIniFile
                SetIniFileProperty "ColorUnfairHigh", m.iColorUnfairHigh, "IOAMT", g.strIniFile
                SetIniFileProperty "ColorUnfairLow", m.iColorUnfairLow, "IOAMT", g.strIniFile
                SetIniFileProperty "ColorValue", m.iColorValue, "IOAMT", g.strIniFile
                SetIniFileProperty "PriceVolBlankRows", m.iBlankRows, "IOAMT", g.strIniFile
                
                m.bInitInprog = False
            End If
        Case "ID_Close"
            Unload Me
        Case "ID_AutoCenter"
            If Tool.State = ssChecked Then
                m.bCenterPrice = True
            Else
                m.bCenterPrice = False
            End If
            FocusGrid
            SetIniFileProperty "PriceVolCenterPrice", m.bCenterPrice, "IOAMT", g.strIniFile
    End Select
    
    Exit Sub

ErrSection:
    RaiseError "frmPriceVol.tbToolbar_ToolClick"

End Sub

Private Sub UpdateGridRT()
On Error GoTo ErrSection:
    
    Static iPrevCol&, iPrevRow&
    Static bInProgress As Boolean
    
    Dim i&, iCol&, iOtherVolCol&, iSpace&, iUpdateColor&
    Dim dPrice#, dGridPrice#, dLastTradePrice#
    Dim dBid#, dAsk#, dOther#, dMinMove#
    Dim dHigh#, dLow#
    
    Dim Table As cGdTable
    Dim aTbIdx As cGdArray
    Dim Bars As cGdBars
    
    Dim bColUnhide As Boolean
    Dim bPriceFound As Boolean
    Dim bHighFound As Boolean
    Dim bLowFound As Boolean
            
    If bInProgress Or m.bInitInprog Or g.bUnloading Then Exit Sub
            
    Set Table = m.Data.PriceVolTable
    If Table Is Nothing Then Exit Sub           'precautionary
    
    Set aTbIdx = m.Data.PriceVolTableIdx
    If aTbIdx Is Nothing Then                   'precautionary
        Set aTbIdx = Table.CreateSortedIndex(eTSVTb_Price)
    End If
    
    If Table.NumRecords <> aTbIdx.Size Then     'precautionary
        Set Table = Nothing
        Set aTbIdx = Nothing
        Exit Sub
    End If
    
    Set Bars = m.Data.TickBars
    If Not Bars Is Nothing Then
        dMinMove = Bars.MinMove(m.Data.SessionDate)
        dLastTradePrice = RoundToMinMove(Bars(eBARS_Close, m.Data.TickBars.Size - 1), dMinMove)
    End If
    
    bInProgress = True
                        
    iCol = m.Data.PriceVolLastDataCol
        
    If Not frmQuotes Is Nothing Then iUpdateColor = frmQuotes.UpdateColor
    
    'get current high/low
    m.Data.HighLow dHigh, dLow
    
    With fgPriceVol
        .Redraw = flexRDNone
                
        .Cell(flexcpBackColor, .FixedRows, 0, .Rows - 1, 0) = .Cell(flexcpBackColor, 0, 0)  'clear out previous highlights
        If iPrevRow > .FixedRows And iPrevCol > .FixedCols Then
            If iPrevRow < .Rows Then        'Or iPrevCol < .Cols Then
                .Cell(flexcpFontBold, .FixedRows, 0, .Rows - 1, .Cols - 2) = False
                .Cell(flexcpForeColor, iPrevRow, 0, iPrevRow, .Cols - 2) = .ForeColor      'clear update color
            Else
                iPrevRow = 0            'reset
                iPrevCol = 0
                .Cell(flexcpForeColor, .FixedRows, .FixedCols, .Rows - 1, .Cols - 2) = .ForeColor
            End If
        End If
        'unhide column if col now has data
        For i = m.Data.PriceVolLastDataCol To eTSVTb_VolBid_A Step -2
            If i > .FixedCols And i < .Cols - 1 Then
                If .ColHidden(i) Then
                    .ColHidden(i) = False
                    bColUnhide = True
                    fgSummary.ColHidden(i) = False
                    If m.eView <> eTSVView_VolAll And m.eView <> eTSVView_AuctionBar Then
                        .ColHidden(i + 1) = False
                        fgSummary.ColHidden(i + 1) = False
                    End If
                Else
                    Exit For
                End If
            End If
        Next
        If m.eView = eTSVView_AuctionBar Then
            SetAuctionBarCol m.Data.AuctionTable, m.Data.StatsLastDataCol, .Rows - 1
            SetAuctionBarCol m.Data.AuctionTable, eTSVTb_Col_Total, .Rows - 1
            
            bPriceFound = False
            bHighFound = False
            bLowFound = False
            
            For i = .FixedRows To .Rows - 1
                dGridPrice = ValOfPrice(.TextMatrix(i, 0))
                If dMinMove > 0 Then dGridPrice = RoundToMinMove(dGridPrice, dMinMove)
                If dGridPrice = dLastTradePrice Then
                    iPrevRow = i
                    If iPrevRow > .FixedRows And iPrevRow < .Rows Then
                        .Cell(flexcpFontBold, iPrevRow, 0, iPrevRow, 0) = True
                        .Cell(flexcpForeColor, iPrevRow, 0, iPrevRow, 0) = iUpdateColor       'set update color
                        m.dLastUpdated = gdTickCount
                        m.iLastPriceRow = iPrevRow
                    End If
                    bPriceFound = True
                End If
                'highlight high/low
                If dGridPrice = dHigh Then
                    .Cell(flexcpBackColor, i, 0, i, 0) = m.iColorAsk
                    bHighFound = True
                ElseIf dGridPrice = dLow Then
                    .Cell(flexcpBackColor, i, 0, i, 0) = m.iColorBid
                    bLowFound = True
                End If
                If bPriceFound And bHighFound And bLowFound Then Exit For
            Next
        Else
            For i = 0 To Table.NumRecords - 1
                If i + 1 < .Rows Then
                    dPrice = Table(eTSVTb_Price, i)
                    dGridPrice = ValOfPrice(.TextMatrix(i + 1, 0))
                    iOtherVolCol = m.Data.TbColVolOther
                    
                    dBid = Table(iCol, i)
                    dAsk = Table(iCol + 1, i)
                    dOther = Table(iOtherVolCol, i)
                    
                    If dMinMove > 0 Then dGridPrice = RoundToMinMove(dGridPrice, dMinMove)
                    If dPrice = dGridPrice Then
                        If dBid > 0 Or dAsk > 0 Then
                            SetGridRow iCol, dBid, dAsk, dOther, Table(eTSVTb_VolRow_Total, i), i + 1
                            If dPrice = dLastTradePrice Then
                                iPrevRow = i + 1
                                iPrevCol = iCol
                            End If
                        End If
                    End If
                    If m.dHistogramMaxVol < dBid + dAsk Then
                        m.dHistogramMaxVol = (dBid + dAsk) * 1.1
                    End If
                    'highlight high/low
                    If dGridPrice = dHigh Then
                        .Cell(flexcpBackColor, i + 1, 0, i + 1, 0) = m.iColorAsk
                    ElseIf dGridPrice = dLow Then
                        .Cell(flexcpBackColor, i + 1, 0, i + 1, 0) = m.iColorBid
                    End If
                End If
            Next
            
            If iPrevRow > 0 And iPrevCol > 0 Then
                .Cell(flexcpFontBold, iPrevRow, 0) = True
                'aardvark 4089 - GX data has already gone past N column, just update current price
                If iCol + 1 < .Cols Then
                    .Cell(flexcpFontBold, iPrevRow, iCol, iPrevRow, iCol + 1) = True
                    .Cell(flexcpForeColor, iPrevRow, 0, iPrevRow, .Cols - 2) = iUpdateColor       'set update color
                Else
                    .Cell(flexcpForeColor, iPrevRow, 0, iPrevRow, 0) = iUpdateColor       'set update color
                End If
                m.dLastUpdated = gdTickCount
                m.iLastPriceRow = iPrevRow
            End If
            
            'color green/yellow bars
            If m.Data.PriceVolLastDataCol >= eTSVTb_VolBid_N Then
                ColorGridRow eTSV_Group_AN
            ElseIf m.Data.PriceVolLastDataCol >= eTSVTb_VolBid_I Then
                ColorGridRow eTSV_Group_AI
            ElseIf m.Data.PriceVolLastDataCol >= eTSVTb_VolBid_E Then
                ColorGridRow eTSV_Group_AE
            ElseIf m.Data.PriceVolLastDataCol >= eTSVTb_VolBid_B Then
                ColorGridRow eTSV_Group_AB
            End If
        End If
        'flood histogram
        For i = .FixedRows To .Rows - 1
            If m.dHistogramMaxVol > 0 Then
                .Cell(flexcpFloodPercent, i, 30) = Table(eTSVTb_VolRow_Total, i - .FixedRows) / m.dHistogramMaxVol * 100 '(ValOfText(.TextMatrix(i, 29)) / m.dHistogramMaxVol * 100)
            Else
                .Cell(flexcpFloodPercent, i, 30) = 0
            End If
        Next
        'color green/yellow bars (price column)
        ColorGridRow eTSV_Group_Current
        .Redraw = flexRDBuffered
        
    End With
    
    i = m.Data.StatsLastDataCol
    SetSumGridData i, i, iCol
    
    If bColUnhide Then
        'summary grid may now have a scrollbar that did not previously have, which could hide last row
        FormResize Me
    End If
            
    Set Table = Nothing
    Set aTbIdx = Nothing
    Set Bars = Nothing
        
    bInProgress = False
    
    Exit Sub

ErrSection:
    RaiseError "frmPriceVol.UpdateGridRT"

End Sub

Private Sub tmr_Timer()
On Error GoTo ErrSection:
    
    Dim i&, dTickCount#
    
    Dim bNewBarVol As Boolean, bNewBarTrade As Boolean
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
        With fgPriceVol
            If m.iLastPriceRow > .FixedRows And m.iLastPriceRow < .Rows Then
                .Cell(flexcpForeColor, m.iLastPriceRow, 0, m.iLastPriceRow, .Cols - 1) = m.iCurrPriceColor
            End If
        End With
    End If
    
    If Not m.Data Is Nothing Then
        If Not m.Data.PriceVolTable Is Nothing Then
            If g.RealTime.Active And g.RealTime.FeedTime > 0 Then
                i = m.Data.UpdateDataRT(bNewBarVol, bNewBarTrade, bTableRebuild)
                If i = -1 Then
                    LoadNewSymData m.strSym     'there was no data when form was opened
                ElseIf bTableRebuild Then
                    LoadGrid                    'new high and/or low
                ElseIf i > 0 Then
                    UpdateGridRT
                ElseIf m.iLastPriceRow > fgPriceVol.FixedRows And m.iLastPriceRow < fgPriceVol.Rows Then
                    With fgPriceVol
                        If .Cell(flexcpFontBold, m.iLastPriceRow, 0) = False Then
                            .Cell(flexcpFontBold, m.iLastPriceRow, 0) = True
                            If m.eView <> eTSVView_AuctionBar Then
                                i = m.Data.PriceVolLastDataCol
                                If i > .FixedCols And i < .Cols Then
                                    .Cell(flexcpFontBold, m.iLastPriceRow, i) = True
                                    If i + 1 < .Cols Then .Cell(flexcpFontBold, m.iLastPriceRow, i + 1) = True
                                End If
                            End If
                        End If
                        If bChanged Then
                            .Cell(flexcpForeColor, m.iLastPriceRow, 0) = m.iCurrPriceColor
                            If m.eView <> eTSVView_AuctionBar Then
                                i = m.Data.PriceVolLastDataCol
                                If i > .FixedCols And i < .Cols Then
                                    .Cell(flexcpForeColor, m.iLastPriceRow, 0, m.iLastPriceRow, i) = m.iCurrPriceColor
                                    If i + 1 < .Cols Then i = i + 1
                                    .Cell(flexcpForeColor, m.iLastPriceRow, 0, m.iLastPriceRow, i) = m.iCurrPriceColor
                                End If
                            End If
                        End If
                    End With
                End If
            End If
        End If
    End If
    
    If m.bReloadData Then
        m.Data.ResetTickBars m.strSym, m.nSymID, Nothing, 570       'this is minutes from midnight = 09:30 for now
        m.Data.BuildTables eTSVTb_TbType_VolAtPrice
        LoadGrid
    ElseIf m.bCenterPrice And Not bTableRebuild Then
        If m.iLastPriceRow > fgPriceVol.FixedRows And m.iLastPriceRow < fgPriceVol.Rows Then
            If Not fgPriceVol.RowIsVisible(m.iLastPriceRow) Then
                FocusGrid
            End If
        End If
    End If
    
    m.bReloadData = False
    
    m.bTimerInProg = False
    
    Exit Sub
    
ErrSection:
    RaiseError "frmPriceVol.tmr_Timer"

End Sub

Public Sub RefreshData()
On Error GoTo ErrSection:

    If m.bInitInprog Then Exit Sub
    
    If m.bTimerInProg Then
        m.bReloadData = True
    Else
        tmr.Enabled = False
        m.Data.ResetTickBars m.strSym, m.nSymID, Nothing, 570
        m.Data.BuildTables eTSVTb_TbType_VolAtPrice
        LoadGrid
        tmr.Enabled = True
    End If
    
    Exit Sub

ErrSection:
    RaiseError "frmPriceVol.RefreshData"

End Sub

Public Sub GetPixWidth(iAuctionBar&, iTriangle&, iMax&)
On Error GoTo ErrSection:
    
    iAuctionBar = m.iAuctionBarPix
    iTriangle = m.iTrianglePix
    
    With fgPriceVol
        iMax = Abs((.Cell(flexcpWidth, .FixedRows, 1) - .Cell(flexcpLeft, .FixedRows, 1)) * 0.7)
    End With
    
    Exit Sub

ErrSection:
    RaiseError "frmPriceVol.GetPixWidth"

End Sub

Public Sub SetPixWidth(ByVal iAuctionBar&, ByVal iTriangle&)
On Error GoTo ErrSection:
    
    m.iAuctionBarPix = iAuctionBar
    m.iTrianglePix = iTriangle
    
    Exit Sub

ErrSection:
    RaiseError "frmPriceVol.SetPixWidth"

End Sub

Public Sub GetColors(iAsk&, IbID&, iHist&, iMean&, iMode&, iTriangle&, iUnHigh&, iUnLow&, iValue&)
On Error GoTo ErrSection:
    
    iAsk = m.iColorAsk
    IbID = m.iColorBid
    iHist = m.iColorHistogram
    iMean = m.iColorMean
    iMode = m.iColorMode
    iTriangle = m.iColorTriangle
    iUnHigh = m.iColorUnfairHigh
    iUnLow = m.iColorUnfairLow
    iValue = m.iColorValue
    
    Exit Sub

ErrSection:
    RaiseError "frmPriceVol.GetColors"

End Sub

Public Sub SetColors(ByVal iAsk&, ByVal IbID&, ByVal iHist&, ByVal iMean&, _
    ByVal iMode&, ByVal iTriangle&, ByVal iUnHigh&, ByVal iUnLow&, ByVal iValue&)
On Error GoTo ErrSection:
    
    'zero is reserved number in grid colors
    If iAsk = 0 Then iAsk = 1
    If IbID = 0 Then IbID = 1
    If iHist = 0 Then iHist = 1
    If iMean = 0 Then iMean = 1
    If iMode = 0 Then iMode = 1
    If iUnHigh = 0 Then iUnHigh = 1
    If iUnLow = 0 Then iUnLow = 1
    If iValue = 0 Then iValue = 1
    
    m.iColorAsk = iAsk
    m.iColorBid = IbID
    m.iColorHistogram = iHist
    m.iColorMean = iMean
    m.iColorMode = iMode
    m.iColorTriangle = iTriangle
    m.iColorUnfairHigh = iUnHigh
    m.iColorUnfairLow = iUnLow
    m.iColorValue = iValue
    
    Exit Sub

ErrSection:
    RaiseError "frmPriceVol.SetColors"

End Sub

Public Sub BoxSettings(ByVal iColor&, ByVal iPix&, ByVal iFill&)
On Error GoTo ErrSection:

    m.iBoxColor = iColor
    m.iBoxPix = iPix
    m.iBoxFill = iFill
    
    Exit Sub

ErrSection:
    RaiseError "frmPriceVol.BoxSettings"

End Sub

Private Sub ResetMouseVar()
On Error GoTo ErrSection:

    m.iMouseColDown = -1
    m.iMouseRowDown = -1
    
    m.iMouseColDownPrev = -1
    m.iMouseRowDownPrev = -1
    
    Exit Sub

ErrSection:
    RaiseError "frmPriceVol.ResetMouseVar"

End Sub

Private Sub FocusGrid()
On Error Resume Next:

    Dim i&, iShowCol&, iShowRow&
        
    If fgPriceVol.Rows <= fgPriceVol.FixedRows Then Exit Sub

    With fgPriceVol
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

Public Property Get BlankRows() As Long
On Error GoTo ErrSection:

    BlankRows = m.iBlankRows
    
    Exit Property
    
ErrSection:
    RaiseError "frmPriceVol.BlankRows.Get"

End Property

Public Property Let BlankRows(ByVal iRows&)
On Error GoTo ErrSection:

    m.iBlankRows = iRows
    
    Exit Property

ErrSection:
    RaiseError "frmPriceVol.BlankRows.Let"

End Property

Private Sub LoadGridNoVol(ByVal strSym$)
On Error GoTo ErrSection:

    Dim i&
    
    Me.Caption = kCaptionBase & " for " & strSym
    
    fgSummary.Visible = False
    
    With fgPriceVol
        .Redraw = flexRDNone
        .Rows = .FixedRows + 1
        .MergeRow(.Rows - 1) = True
        .TextMatrix(.Rows - 1, 0) = ""
        For i = 1 To .Cols - 1
            .TextMatrix(.Rows - 1, i) = kFootPrintNoVol
        Next
        .Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
        .Redraw = flexRDBuffered
    End With
        
    Exit Sub
    
ErrSection:
    RaiseError "frmPriceVol.LoadGridNoVol"

End Sub

Public Sub IconPaletteClose()

    tbToolbar.Tools("ID_Icons").State = ssUnchecked
    fgPriceVol.MousePointer = flexDefault
    
End Sub

Private Function ValOfPrice(ByVal strPrice$) As Double
On Error GoTo ErrSection:

    Dim dPrice#
    
    If InStr(strPrice, "^") Then
        dPrice = m.Data.TickBars.PriceFromString(strPrice)
    Else
        dPrice = ValOfText(strPrice)
    End If
    
    ValOfPrice = dPrice
    
ErrExit:
    Exit Function

ErrSection:
    RaiseError "frmPriceVol.ValOfPrice"

End Function

