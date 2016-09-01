VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmTimeSales 
   Caption         =   "Time and Sales"
   ClientHeight    =   5685
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   5685
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniTextBoxXP txtMessage 
      Height          =   330
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   4695
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmTimeSales.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   0
      MultiLine       =   0   'False
      Alignment       =   2
      ScrollBars      =   0
      PasswordChar    =   ""
      TrapTab         =   0   'False
      EnableContextMenu=   -1  'True
      RaiseChangeEvent=   -1  'True
      Tip             =   "frmTimeSales.frx":002A
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTimeSales.frx":004A
   End
   Begin HexUniControls.ctlUniFrameWL fraVolFilter 
      Height          =   285
      Left            =   45
      TabIndex        =   1
      Top             =   5085
      Visible         =   0   'False
      Width           =   4545
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
      Caption         =   "frmTimeSales.frx":0066
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmTimeSales.frx":009E
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTimeSales.frx":00BE
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniLabelXP lblVolumeFilterOn 
         Height          =   255
         Left            =   60
         Top             =   45
         Width           =   2760
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
         Caption         =   "frmTimeSales.frx":00DA
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   2
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmTimeSales.frx":011A
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTimeSales.frx":013A
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Interval        =   350
      Left            =   3600
      Top             =   2400
   End
   Begin VSFlex7LCtl.VSFlexGrid fg 
      Height          =   4515
      Left            =   300
      TabIndex        =   0
      Top             =   240
      Width           =   2295
      _cx             =   4048
      _cy             =   7964
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
      Left            =   3540
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131083
      ToolBarsCount   =   1
      ToolsCount      =   6
      DisplayContextMenu=   0   'False
      Tools           =   "frmTimeSales.frx":0156
      ToolBars        =   "frmTimeSales.frx":1C2A
   End
End
Attribute VB_Name = "frmTimeSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum eTSCols
    eCol_Time = 0
    eCol_Price
    eCol_Size
End Enum

Private Enum eTSDisplayStyle
    eStyle_None = 0
    eStyle_TickByTick       '3 columns, Time-Price-Size (no wrap)
    eStyle_MinByMin         '2 columns, Tim-Price (wrap price)
    eStyle_TickBidAsk       '4 columns, Time-Price-Type-Size (no wrap)
    eStyle_Cumulative       '8 columns, Time-Price-Trades-Contracts-AvgTradeSize-LargestTrade-BuyVol-SellVol
End Enum

Private Type mPrivate
    WindowLink As New cWindowLink
    
    fgSource As New cTimeSalesData
    eStyle As eTSDisplayStyle
    
    nFontSize As Long
    nSymID As Long
    nSortOrder As Long
    nUpColor As Long
    nDownColor As Long
    nUpColorBid As Long
    nDownColorBid As Long
    nUpColorAsk As Long
    nDownColorAsk As Long
        
    bStyleChanged As Boolean
    bSessionChanged As Boolean
    
    strFont As String
    strSym As String
    bBold As Boolean
    bItalic As Boolean
    
    nSessionDate As Long
    bSessionCurrent As Boolean
    
    bUpdateInProg As Boolean
        
End Type

Private m As mPrivate

Private Sub fg_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
On Error Resume Next

    '6969
    Dim i&, sizes$
    
    For i = 0 To fg.Cols - 1
        If i = 0 Then
            sizes = Str(fg.ColWidth(i))
        Else
            sizes = sizes & "|" & Str(fg.ColWidth(i))
        End If
    Next

    SetIniFileProperty "GridColSizes", sizes, "Time And Sales", g.strIniFile

End Sub

Private Sub fg_BeforeScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long, Cancel As Boolean)

    GridScrollCheck fg, OldTopRow, OldLeftCol, NewTopRow, NewLeftCol, Cancel

End Sub

Private Sub fg_BeforeSort(ByVal Col As Long, Order As Integer)

    If Col = eCol_Time Then
        m.nSortOrder = Order
    End If

End Sub

Private Sub fg_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 36 Then                    'Home key
        fg.ShowCell fg.FixedRows, 0
        fg.Row = fg.FixedRows
    ElseIf KeyCode = 35 Then                'End key
        fg.ShowCell fg.Rows - 1, 0
        fg.Row = fg.Rows - 1
    ElseIf KeyCode = vbKeyF1 Then
        g.Help.ShowF1Help Nothing
    End If

End Sub

Private Sub Form_Activate()
On Error Resume Next

    TextIncDecRegisterForm Me, False

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTimeSales.Form_Activate"

End Sub

Private Sub Form_Deactivate()
On Error GoTo ErrSection:
    
    TextIncDecUnregisterForm Me

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTimeSales.Form_Deactivate"

End Sub

Private Sub Form_Load()
On Error GoTo ErrSection:

    Dim strText$
    
    g.Styler.StyleForm Me

    Me.Icon = Picture16(ToolbarIcon("ID_TimeSales"), , True)
    m.fgSource.FormTS = Me
    m.nSortOrder = flexSortGenericDescending
    
    m.nSessionDate = 0
    m.bSessionCurrent = True
    'get saved settings from INI file
    m.eStyle = GetIniFileProperty("GridDisplayStyle", eStyle_TickByTick, "Time And Sales", g.strIniFile)
    m.nFontSize = GetIniFileProperty("GridFontSize", 8, "Time And Sales", g.strIniFile)
    m.strFont = GetIniFileProperty("GridFontName", CheckSSFont, "Time And Sales", g.strIniFile)
    m.bBold = GetIniFileProperty("GridFontBold", False, "Time And Sales", g.strIniFile)
    m.bItalic = GetIniFileProperty("GridFontItalic", False, "Time And Sales", g.strIniFile)
    
    m.nUpColor = GetIniFileProperty("GridUpColor", RGB(0, 128, 0), "Time And Sales", g.strIniFile)
    m.nDownColor = GetIniFileProperty("GridDownColor", vbRed, "Time And Sales", g.strIniFile)
    
    m.nUpColorBid = GetIniFileProperty("GridUpColorBid", vbBlue, "Time And Sales", g.strIniFile)
    m.nDownColorBid = GetIniFileProperty("GridDownColorBid", RGB(0, 0, 160), "Time And Sales", g.strIniFile)
    
    m.nUpColorAsk = GetIniFileProperty("GridUpColorAsk", RGB(255, 0, 128), "Time And Sales", g.strIniFile)
    m.nDownColorAsk = GetIniFileProperty("GridDownColorAsk", RGB(128, 0, 128), "Time And Sales", g.strIniFile)
    
    strText = Str(m.nSymID) & "VolFilterMax"
    m.fgSource.VolFilterMax = GetIniFileProperty(strText, 0, "Time And Sales", g.strIniFile)
    
    strText = Str(m.nSymID) & "VolFilterMin"
    m.fgSource.VolFilterMin = GetIniFileProperty(strText, 0, "Time And Sales", g.strIniFile)
    
    strText = "Classic"
    With tbToolbar
        strText = "Classic"
        If g.nTbIconStyle = 1 Then
            If g.nColorTheme = kDarkThemeColor Then
                strText = "Light"
            Else
                strText = "Dark"
            End If
        End If
        .Tools("ID_Symbol").Picture = g.CoreBridge.ImgListToolbarExt(strText, ToolbarIcon("ID_Symbol"), "", 16).ExtractIcon
        .Tools("ID_Settings").Picture = g.CoreBridge.ImgListToolbarExt(strText, ToolbarIcon("ID_Settings"), "", 16).ExtractIcon
        .Tools("ID_TextInc").Picture = g.CoreBridge.ImgListToolbarExt(strText, ToolbarIcon("ID_TextIncrease"), "", 16).ExtractIcon
        .Tools("ID_TextDec").Picture = g.CoreBridge.ImgListToolbarExt(strText, ToolbarIcon("ID_TextDecrease"), "", 16).ExtractIcon
        .Tools("ID_Close").Picture = Picture16(ToolbarIcon("kCancel"))
    End With
    
    InitGrid
    
    'Restore/set form size & location
    strText = GetIniFileProperty(Me.Name, "", "Placement", g.strIniFile)
    If strText = "" Then
        CenterTheForm Me
    Else
        SetFormPlacement Me, strText
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTimeSales.Form.Load", eGDRaiseError_Raise

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    tmr.Enabled = False
    
    TextIncDecUnregisterForm Me
    m.fgSource.TradeProfileUnload = True
    SaveSettings
    Set m.fgSource = Nothing
    If Cancel = 0 Then m.WindowLink.Unhook

End Sub

Private Sub Form_Resize()
On Error Resume Next

    Dim bVolFilterOn As Boolean
    
    If Not m.fgSource Is Nothing Then
        If m.fgSource.VolFilterMax > 0 Or m.fgSource.VolFilterMin > 0 Then bVolFilterOn = True
    End If
    
    With fraVolFilter
        .Visible = bVolFilterOn
        .Move 0, 0, Me.ScaleWidth
        lblVolumeFilterOn.Width = .Width
    End With
    
    With fg
        If bVolFilterOn Then
            .Move 0, _
                  fraVolFilter.Top + fraVolFilter.Height, _
                  Me.ScaleWidth, _
                  Me.ScaleHeight - fraVolFilter.Height
        Else
            .Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
        End If
        If m.eStyle = eStyle_MinByMin Then .AutoSize eCol_Price
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Set m.WindowLink = Nothing
    
End Sub

Private Sub InitGrid()
On Error GoTo ErrSection:

    SetupGrid fg, eGridMode_Grid
    
    Dim i As Long
    Dim colSizes As String
    Dim aSizes As New cGdArray
        
    colSizes = GetIniFileProperty("GridColSizes", "", "Time And Sales", g.strIniFile)
    aSizes.SplitFields colSizes, "|"
    
    With fg
        .Redraw = flexRDNone
        
        .FixedRows = 1
        .FixedCols = 0
        .Editable = flexEDNone
        .ExplorerBar = flexExSort
        .WordWrap = True
        .AutoSizeMode = flexAutoSizeRowHeight
        .Font.Name = m.strFont
        .Font.Size = m.nFontSize
        .Font.Bold = m.bBold
        .FontItalic = m.bItalic
        .FlexDataSource = m.fgSource
        .MergeCells = flexMergeFree
        
        If aSizes.Size > 0 Then         '6969
            .Cols = aSizes.Size
            For i = 0 To .Cols - 1
                .ColWidth(i) = aSizes(i)
            Next
        End If
        
        If m.eStyle = eStyle_MinByMin Then
            .ColAlignment(eCol_Price) = flexAlignLeftCenter
        ElseIf .FontItalic Then
            .ColAlignment(eCol_Price) = flexAlignCenterCenter   'numbers get cut-off when using right alignment
            .ColAlignment(eCol_Size) = flexAlignCenterCenter
        End If
        
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTimeSales.InitGrid", eGDRaiseError_Raise

End Sub

Public Sub ShowMe(ByVal nSymbolID&)
On Error GoTo ErrSection:

    ToggleStartStop True, False

    m.nSymID = nSymbolID
    m.strSym = GetSymbol(nSymbolID)
    m.nSessionDate = 0
    m.bSessionCurrent = True            '4734
    
    SetMyCaption
    
    'JM 12-18-2015: this form has always been shown without alternate grid row color
    '   in Ivory & Classic looks better with just one color
    '   but in dark theme looks better with alternate grid row color
    If g.nColorTheme = kDarkThemeColor Then
        ShowForm Me, False, frmMain, , ALT_GRID_ROW_COLOR
    Else
        ShowForm Me, False, frmMain
    End If
    m.WindowLink.Init Me
    
    ChangeSymbol m.nSymID
        
        
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTimeSales.ShowMe", eGDRaiseError_Raise

End Sub

Public Property Get DisplayStyle() As Long
    DisplayStyle = m.eStyle
End Property

Public Property Get SortOrder() As Long
    SortOrder = m.nSortOrder
End Property

Private Sub tbToolbar_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
On Error Resume Next

    Dim eOldStyle As eTSDisplayStyle
    
    Dim nVolMaxSave&, nVolMinSave&
    Dim bFilterChanged As Boolean
    
    Select Case Tool.ID
        Case "ID_Symbol"
            ToggleStartStop True, False
            ChangeSymbol
        Case "ID_Close"
            Unload Me
        Case "ID_TextInc"
            fg.Font.Size = fg.Font.Size + 1
        Case "ID_TextDec"
            fg.Font.Size = fg.Font.Size - 1
        Case "ID_Settings"
            m.bStyleChanged = False
            m.bSessionChanged = False
            eOldStyle = m.eStyle
            nVolMaxSave = m.fgSource.VolFilterMax
            nVolMinSave = m.fgSource.VolFilterMin
            
            tmr.Enabled = False     '4732
            
            frmTimeSalesCfg.ShowMe Me
            
            If m.fgSource.VolFilterMax <> nVolMaxSave Or m.fgSource.VolFilterMin <> nVolMinSave Then bFilterChanged = True
            
            If m.bStyleChanged Or m.bSessionChanged Or bFilterChanged Then
                If m.eStyle <> eStyle_Cumulative Then ToggleStartStop True, False
                ChangeDisplayStyle bFilterChanged
            End If
        Case "ID_Start"
            ToggleStartStop False, Tool.Visible
            
    End Select

End Sub

Private Sub ChangeDisplayStyle(ByVal bFilterChanged As Boolean)
On Error GoTo ErrSection:

    Dim bUnload As Boolean

    If m.bSessionChanged Or bFilterChanged Then
        With fg
            .Redraw = flexRDNone
            .FlexDataSource = Nothing
            .Rows = .FixedRows
            .Font.Name = m.strFont
            .Font.Size = m.nFontSize
            .Font.Bold = m.bBold
            .FontItalic = m.bItalic
            .Redraw = flexRDBuffered
        End With
        m.fgSource.ClearData
        If Not m.fgSource.NewTradeProfileData(m.nSymID, m.strSym, bFilterChanged) Then bUnload = True
    ElseIf m.eStyle <> eStyle_TickBidAsk Then
        m.fgSource.RemoveBidAskRecords
    End If
    
    If bUnload Then
        Unload Me
    Else
        With fg
            .Redraw = flexRDNone
            .FlexDataSource = Nothing
            .Rows = .FixedRows
            .Font.Name = m.strFont
            .Font.Size = m.nFontSize
            .Font.Bold = m.bBold
            .FontItalic = m.bItalic
            fg.FlexDataSource = m.fgSource
            .Redraw = flexRDBuffered
            If m.eStyle = eStyle_MinByMin Then
                .ColAlignment(eCol_Price) = flexAlignLeftCenter
                .AutoSize eCol_Price
            ElseIf m.bItalic Then
                .ColAlignment(eCol_Price) = flexAlignCenterCenter   'numbers get cut-off when using right alignment
                .ColAlignment(eCol_Size) = flexAlignCenterCenter
            Else
                .ColAlignment(eCol_Price) = flexAlignGeneral
                .ColAlignment(eCol_Size) = flexAlignGeneral
            End If
        End With
        
        SetMyCaption
        Form_Resize
        SaveSettings
            
        If m.bSessionCurrent And g.RealTime.Active Then
            tmr.Enabled = True  '4732
        End If
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTimeSales.ChangeDisplayStyle", eGDRaiseError_Raise

End Sub

Private Sub ChangeSymbol(Optional ByVal nSymID& = 0)
On Error GoTo ErrSection:

    Dim astrSymbols As New cGdArray     ' Symbol(s) back from the symbol selector
    Dim lSymbolID As Long               ' Symbol ID for the symbol selected
    
    astrSymbols.Create eGDARRAY_Strings

    If nSymID = 0 Then
        Set astrSymbols = frmSymbolSelector.ShowMe("", False)
        If astrSymbols.Size > 0 Then
            lSymbolID = g.SymbolPool.SymbolIDforSymbol(astrSymbols(0))
        End If
    Else
        lSymbolID = nSymID
    End If
    If lSymbolID = 0 Then
        Beep
    Else
        tmr.Enabled = False
        m.strSym = GetSymbol(lSymbolID)
        
        With fg
            .Redraw = flexRDNone
            .FlexDataSource = Nothing
            .Rows = .FixedRows
            .Redraw = flexRDBuffered
        End With
        
        If m.bSessionCurrent Then
            m.nSessionDate = 0          'aardvark 4127
        ElseIf nSymID = m.nSymID Then
            Exit Sub                    'aardvark 4733, 4734
        End If
        m.nSymID = lSymbolID

        m.fgSource.ClearData
        If m.fgSource.NewTradeProfileData(m.nSymID, m.strSym) Then
            With fg
                .Redraw = flexRDNone
                .FlexDataSource = Nothing
                .Rows = .FixedRows
                .FlexDataSource = m.fgSource
                .Redraw = flexRDBuffered
                If m.eStyle = eStyle_MinByMin Then .AutoSize eCol_Price
            End With
            SetMyCaption
    
            If g.RealTime.Active Then tmr.Enabled = True
            
            frmMain.SetWindowLink Me
        Else
            Unload Me
        End If
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTimeSales.ChangeSymbol", eGDRaiseError_Raise

End Sub

Private Sub SetMyCaption()
On Error Resume Next

    Dim strName$

    If m.nSessionDate > 0 Then
        strName = m.strSym & " on " & DateFormat(m.nSessionDate)
    Else
        strName = m.strSym
    End If
    
    SetEditorCaption Me, "Time & Sales", strName

End Sub

Private Sub tmr_Timer()
On Error GoTo ErrSection:

    Dim bDontCare As Boolean

    TimerStart "tmrTimeSales.tmr"
    If m.bUpdateInProg Then Exit Sub
    
    If Not g.RealTime.Active Or Not m.bSessionCurrent Then  '4732
        If tbToolbar.Tools("ID_Start").Visible Then ToggleStartStop True, False
        tmr.Enabled = False
        Exit Sub
    End If
    
    If m.eStyle = eStyle_Cumulative Then
        If Not tbToolbar.Tools("ID_Start").Visible Then ToggleStartStop True, True
    ElseIf tbToolbar.Tools("ID_Start").Visible Then
        ToggleStartStop True, False
    End If
        
    m.bUpdateInProg = True
        
    With fg
        If .TopRow = .FixedRows Then
            If m.fgSource.UpdateDataRT(False) Then
                .Redraw = flexRDNone
                .FlexDataSource = m.fgSource
                If m.eStyle = eStyle_MinByMin Then .AutoSize eCol_Price
                .Redraw = flexRDDirect
            End If
        Else
            bDontCare = m.fgSource.UpdateDataRT(True)
        End If
        
        If m.eStyle = eStyle_Cumulative Then
            If tbToolbar.Tools("ID_Start").Name = "Stop" Then
                fg.TextMatrix(0, 6) = "BV=" & m.fgSource.SumBuyVol
                fg.TextMatrix(0, 7) = "SV=" & m.fgSource.SumSellVol
            End If
        End If
    End With
    
    m.bUpdateInProg = False
    TimerEnd "tmrTimeSales.tmr", tmr.Interval

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTimeSales.Timer", eGDRaiseError_Raise

End Sub

Private Sub SaveSettings()
On Error GoTo ErrSection:

    Dim strText$, strVal$
    
    'save settings to INI file
    SetIniFileProperty "GridDisplayStyle", m.eStyle, "Time And Sales", g.strIniFile
    SetIniFileProperty "GridFontSize", m.nFontSize, "Time And Sales", g.strIniFile
    SetIniFileProperty "GridFontName", m.strFont, "Time And Sales", g.strIniFile
    SetIniFileProperty "GridFontBold", m.bBold, "Time And Sales", g.strIniFile
    SetIniFileProperty "GridFontItalic", m.bItalic, "Time And Sales", g.strIniFile
    
    SetIniFileProperty "GridUpColor", m.nUpColor, "Time And Sales", g.strIniFile
    SetIniFileProperty "GridDownColor", m.nDownColor, "Time And Sales", g.strIniFile
    
    SetIniFileProperty "GridUpColorBid", m.nUpColorBid, "Time And Sales", g.strIniFile
    SetIniFileProperty "GridDownColorBid", m.nDownColorBid, "Time And Sales", g.strIniFile
    
    SetIniFileProperty "GridUpColorAsk", m.nUpColorAsk, "Time And Sales", g.strIniFile
    SetIniFileProperty "GridDownColorAsk", m.nDownColorAsk, "Time And Sales", g.strIniFile
    
    strVal = ""
    strText = Str(m.nSymID) & "VolFilterMax"
    If m.fgSource.VolFilterMax > 0 Then
        strVal = Str(m.fgSource.VolFilterMax)
        SetIniFileProperty strText, strVal, "Time And Sales", g.strIniFile
    Else
        'delete the key entirely to keep INI file from getting cluttered with unnecessary keys & values
        WritePrivateProfileString "Time And Sales", strText, Nothing, g.strIniFile
    End If
    
    strVal = ""
    strText = Str(m.nSymID) & "VolFilterMin"
    If m.fgSource.VolFilterMin > 0 Then
        strVal = Str(m.fgSource.VolFilterMin)
        SetIniFileProperty strText, strVal, "Time And Sales", g.strIniFile
    Else
        'delete the key entirely to keep INI file from getting cluttered with unnecessary keys & values
        WritePrivateProfileString "Time And Sales", strText, Nothing, g.strIniFile
    End If
    
    'save form size & location
    SetIniFileProperty Me.Name, GetFormPlacement(Me), "Placement", g.strIniFile

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTimeSales.SaveSettings", eGDRaiseError_Raise

End Sub

Public Sub RefreshData()
On Error Resume Next

    If tmr.Enabled Then Exit Sub
    ChangeSymbol m.nSymID

End Sub

Public Property Get TS_DisplayStyle() As Long
    TS_DisplayStyle = m.eStyle
End Property

Public Property Let TS_DisplayStyle(ByVal nStyle&)
    If m.eStyle <> nStyle Then
        m.bStyleChanged = True
        m.eStyle = nStyle
    End If
End Property

Public Property Get TS_UpColor() As Long
    TS_UpColor = m.nUpColor
End Property

Public Property Let TS_UpColor(ByVal nColor&)
On Error GoTo ErrSection:

    If m.nUpColor <> nColor Then
        m.bStyleChanged = True
        m.nUpColor = nColor
        If Not m.fgSource Is Nothing Then
            m.fgSource.ChangeColor "Trade", nColor, 1
        End If
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "frmTimeSales.TS_UpColor", eGDRaiseError_Raise

End Property

Public Property Get TS_DownColor() As Long
    TS_DownColor = m.nDownColor
End Property

Public Property Let TS_DownColor(ByVal nColor&)
On Error GoTo ErrSection:

    If m.nDownColor <> nColor Then
        m.bStyleChanged = True
        m.nDownColor = nColor
        If Not m.fgSource Is Nothing Then
            m.fgSource.ChangeColor "Trade", nColor, 0
        End If
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "frmTimeSales.TS_DownColor", eGDRaiseError_Raise

End Property

Public Property Get TS_FontItalic() As Boolean
    TS_FontItalic = m.bItalic
End Property

Public Property Let TS_FontItalic(ByVal bItalic As Boolean)
    If m.bItalic <> bItalic Then
        m.bStyleChanged = True
        m.bItalic = bItalic
    End If
End Property

Public Property Get TS_FontBold() As Boolean
    TS_FontBold = m.bBold
End Property

Public Property Let TS_FontBold(ByVal bBold As Boolean)
    If m.bBold <> bBold Then
        m.bStyleChanged = True
        m.bBold = bBold
    End If
End Property

Public Property Get TS_FontName() As String
    TS_FontName = m.strFont
End Property

Public Property Let TS_FontName(ByVal strFont$)
    If m.strFont <> strFont Then
        m.bStyleChanged = True
        m.strFont = strFont
    End If
End Property

Public Property Get TS_FontSize() As Long
    TS_FontSize = m.nFontSize
End Property

Public Property Let TS_FontSize(ByVal nSize&)
    If m.nFontSize <> nSize Then
        m.bStyleChanged = True
        m.nFontSize = nSize
    End If
End Property

Public Property Get TS_SessionCurrent() As Boolean
    TS_SessionCurrent = m.bSessionCurrent
End Property

Public Property Let TS_SessionCurrent(ByVal bCurrent As Boolean)
    m.bSessionCurrent = bCurrent
End Property

Public Property Get TS_SessionDate() As Long
    TS_SessionDate = m.nSessionDate
End Property

Public Property Let TS_SessionDate(ByVal nDate&)
    m.nSessionDate = nDate
End Property

Public Property Let TS_SessionChanged(ByVal bChanged As Boolean)
    m.bSessionChanged = bChanged
End Property

Public Property Get TS_UpColorBid() As Long
    TS_UpColorBid = m.nUpColorBid
End Property

Public Property Let TS_UpColorBid(ByVal nColor&)
On Error GoTo ErrSection:

    If m.nUpColorBid <> nColor Then
        m.bStyleChanged = True
        m.nUpColorBid = nColor
        If Not m.fgSource Is Nothing Then
            m.fgSource.ChangeColor "Bid", nColor, 1
        End If
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "frmTimeSales.TS_UpColorBid", eGDRaiseError_Raise

End Property

Public Property Get TS_DownColorBid() As Long
    TS_DownColorBid = m.nDownColorBid
End Property

Public Property Let TS_DownColorBid(ByVal nColor&)
On Error GoTo ErrSection:

    If m.nDownColorBid <> nColor Then
        m.bStyleChanged = True
        m.nDownColorBid = nColor
        If Not m.fgSource Is Nothing Then
            m.fgSource.ChangeColor "Bid", nColor, 0
        End If
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "frmTimeSales.TS_DownColorBid", eGDRaiseError_Raise

End Property

Public Property Get TS_UpColorAsk() As Long
    TS_UpColorAsk = m.nUpColorAsk
End Property

Public Property Let TS_UpColorAsk(ByVal nColor&)
On Error GoTo ErrSection:

    If m.nUpColorAsk <> nColor Then
        m.bStyleChanged = True
        m.nUpColorAsk = nColor
        If Not m.fgSource Is Nothing Then
            m.fgSource.ChangeColor "Ask", nColor, 1
        End If
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "frmTimeSales.TS_UpColorAsk", eGDRaiseError_Raise

End Property

Public Property Get TS_DownColorAsk() As Long
    TS_DownColorAsk = m.nDownColorAsk
End Property

Public Property Let TS_DownColorAsk(ByVal nColor&)
On Error GoTo ErrSection:

    If m.nDownColorAsk <> nColor Then
        m.bStyleChanged = True
        m.nDownColorAsk = nColor
        If Not m.fgSource Is Nothing Then
            m.fgSource.ChangeColor "Ask", nColor, 0
        End If
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "frmTimeSales.TS_DownColorAsk", eGDRaiseError_Raise

End Property

Public Property Get WindowLink() As cWindowLink
    Set WindowLink = m.WindowLink
End Property

Public Property Get SymbolID() As Long
    SymbolID = m.nSymID
End Property

Public Property Let SymbolID(ByVal nSymbolID As Long)
On Error GoTo ErrSection:
    
    ChangeSymbol nSymbolID
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "frmTimeSales.LetSymbolID", eGDRaiseError_Raise
End Property

Public Property Get TS_VolFilterMax() As Long
    If Not m.fgSource Is Nothing Then TS_VolFilterMax = m.fgSource.VolFilterMax
End Property

Public Property Let TS_VolFilterMax(ByVal nVolMax&)
    m.fgSource.VolFilterMax = nVolMax
End Property

Public Property Get TS_VolFilterMin() As Long
    If Not m.fgSource Is Nothing Then TS_VolFilterMin = m.fgSource.VolFilterMin
End Property

Public Property Let TS_VolFilterMin(ByVal nVolMin&)
    m.fgSource.VolFilterMin = nVolMin
End Property

Public Property Get TS_CanHaveVolFilter() As Boolean
    If Not m.fgSource Is Nothing Then TS_CanHaveVolFilter = m.fgSource.CanHaveVolFilter
End Property

Public Sub UpdateMessage(ByVal strMsg$)
On Error GoTo ErrSection:

    If Len(strMsg) > 0 Then
        txtMessage.Move 0, 0, Me.Width, 250
        txtMessage.Text = strMsg
        txtMessage.Visible = True
    Else
        txtMessage.Visible = False
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTimeSales.UpdateMessage"

End Sub

Private Sub ToggleStartStop(ByVal bReset As Boolean, ByVal bVisible As Boolean)
On Error GoTo ErrSection:

    With tbToolbar.Tools("ID_Start")
        
        If .Name = "Stop" Then
            m.fgSource.SumBuySell = False
            .ChangeAll ssChangeAllName, "Start"
            fg.TextMatrix(0, 6) = "Buy Vol"
            fg.TextMatrix(0, 7) = "Sell Vol"
        ElseIf Not bReset Then
            m.fgSource.SumBuySell = True
            .ChangeAll ssChangeAllName, "Stop"
        End If
        
        .Visible = bVisible
    
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTimeSales.ToggleStartStop"

End Sub

