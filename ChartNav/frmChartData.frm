VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmChartData 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Chart Data"
   ClientHeight    =   4065
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   2550
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   2550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniCheckXP chkAuto 
      Height          =   220
      Left            =   1080
      TabIndex        =   2
      Top             =   0
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   397
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmChartData.frx":0000
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Pressed         =   -1  'True
      Tip             =   "frmChartData.frx":0034
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frmChartData.frx":00CC
      ShowFocus       =   -1  'True
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniCheckXP chkSingleBar 
      Height          =   220
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   397
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmChartData.frx":00E8
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Pressed         =   -1  'True
      Tip             =   "frmChartData.frx":011E
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frmChartData.frx":013E
      ShowFocus       =   -1  'True
      RightToLeft     =   0   'False
   End
   Begin VSFlex7LCtl.VSFlexGrid fg 
      Height          =   3075
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   2295
      _cx             =   4048
      _cy             =   5424
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
   Begin VB.Menu mnuSettings 
      Caption         =   "Settings"
      Begin VB.Menu mnuSingleBar 
         Caption         =   "Single Bar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuMultipleBars 
         Caption         =   "Multiple Bars"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCursor 
         Caption         =   "Sync with Cursor"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuOnClick 
         Caption         =   "Display on Click"
      End
      Begin VB.Menu mnuLastBar 
         Caption         =   "Display the Last Bar"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChangeFont 
         Caption         =   "Change Font"
      End
   End
End
Attribute VB_Name = "frmChartData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const kExtendedCol = 0

Private Type mPrivate
    DataGrid As cChartDataGrid
    strPrevData As String
    nPrevX As Long
    nPrevColWidth As Long
End Type
Private m As mPrivate

Private Sub chkAuto_Click()
On Error GoTo ErrSection:

    g.ChartGlobals.bAutoChartData = -chkAuto

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartData.chkAuto.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub chkSingleBar_Click()
On Error GoTo ErrSection:

    Dim nWidth&
    
    If Not Me.Visible Then Exit Sub

    If Not g.ChartGlobals.bChartDataSingleBar Then
        m.nPrevX = fg.Row - 1
    End If
    g.ChartGlobals.bChartDataSingleBar = -chkSingleBar
    
    ' get current undocked width
    If DockState(Me) = eUndocked Then
        nWidth = Me.Width
    Else
        nWidth = frmMain.DockPro.WidthWhenUnDocked(Me.Name) * Screen.TwipsPerPixelY
    End If
    ' save current undocked width for old mode,
    ' and get previous undocked width for new mode
    If g.ChartGlobals.bChartDataSingleBar Then
        SetIniFileProperty "WidthWhenMultiBar", nWidth, "Charting", g.strIniFile
        nWidth = GetIniFileProperty("WidthWhenSingleBar", 2670, "Charting", g.strIniFile)
    Else
        SetIniFileProperty "WidthWhenSingleBar", nWidth, "Charting", g.strIniFile
        nWidth = GetIniFileProperty("WidthWhenMultiBar", 9000, "Charting", g.strIniFile)
    End If
    ' restore previous undocked width for new mode
    frmMain.DockPro.WidthWhenUnDocked(Me.Name) = nWidth / Screen.TwipsPerPixelY
    If DockState(Me) = eUndocked Then
        Me.Width = nWidth
    End If
    
    InitGrid
    m.strPrevData = ""
    ShowData m.nPrevX
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartData.chkSingleBar.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fg_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    Dim w&, i&
    
    ' if column being resized is the extended column,
    ' then make the next column bigger (instead of adjusting
    ' the extended column)
    If Col >= kExtendedCol And g.ChartGlobals.bChartDataSingleBar Then
        With fg
            .Redraw = flexRDNone
            w = .ColWidth(Col) - m.nPrevColWidth
            For i = Col + 1 To .Cols - 1
                If Not .ColHidden(i) Then
                    .ColWidth(i) = fg.ColWidth(i) - w
                    Exit For
                End If
            Next
            m.nPrevColWidth = 0
            ExtendCustomColumn
            .Redraw = flexRDBuffered
        End With
    Else
        ExtendCustomColumn
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartData.fg.AfterUserResize", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fg_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    ' if column being resized is the extended column, save size
    If Col >= kExtendedCol And g.ChartGlobals.bChartDataSingleBar Then
        m.nPrevColWidth = fg.ColWidth(Col)
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartData.fg.BeforeUserResize", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fg_DblClick()
On Error GoTo ErrSection:

    ExtendCustomColumn

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartData.fg.DblClick", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fg_KeyDown(KeyCode As Integer, Shift As Integer)

    If fgKeyDown(KeyCode, Shift) Then Exit Sub

End Sub

Private Sub fg_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    With fg
        If .MouseRow > .FixedRows - 1 And .MouseRow < .Rows Then
            .Row = .MouseRow
            .RowSel = .Row
            
            If Button = vbRightButton Then
                If chkSingleBar.Value = 0 Then
                    mnuSingleBar.Checked = False
                    mnuMultipleBars.Checked = True
                Else
                    mnuSingleBar.Checked = True
                    mnuMultipleBars.Checked = False
                End If
                If chkAuto.Value = 0 Then
                    mnuCursor.Checked = False
                    mnuOnClick.Checked = True
                Else
                    mnuCursor.Checked = True
                    mnuOnClick.Checked = False
                End If
                PopupMenu mnuSettings
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartData.fg.MouseDown", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fg_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    Dim nRow&, nCol&, strTip$
    
    With fg
        nRow = .MouseRow
        nCol = .MouseCol
        If g.ChartGlobals.bChartDataSingleBar Then
            If nCol = 0 And nRow > 0 Then
                strTip = .TextMatrix(nRow, nCol)
            End If
        Else
            If nRow = 0 Then
                strTip = .TextMatrix(nRow, nCol)
            End If
        End If
        .ToolTipText = strTip
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartData.fg.MouseMove", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub Form_Deactivate()
On Error GoTo ErrSection:

    SetPrevActiveForm Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartData.Form.Deactivate", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If KeyCode = vbKeyF1 Then
        KeyCode = 0
        g.Help.ShowF1Help Me
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartData.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub Form_Load()
On Error GoTo ErrSection:

    Dim i&, strFont$
    
    mnuSettings.Visible = False
    mnuLastBar.Visible = False
    
    Me.Icon = Picture16(ToolbarIcon("ID_ChartData"))
    CenterTheForm Me
    
    g.Styler.StyleForm Me
    
    ' Get font from INI file
    strFont = GetIniFileProperty("ChartData", "", "Fonts", g.strIniFile)
    
    chkAuto.Value = Abs(g.ChartGlobals.bAutoChartData)
    chkSingleBar.Value = Abs(g.ChartGlobals.bChartDataSingleBar)
    
    Set m.DataGrid = New cChartDataGrid
    
    SetupGrid fg, eGridMode_Grid
    
    If strFont <> "" Then
        FontFromString fg.Font, strFont
    End If
       
    InitGrid
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartData.Form.Load", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode = 0 Then
        frmMain.tbToolbar.Tools("ID_ChartData").State = ssUnchecked
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartData.Form.QueryUnload", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub Form_Resize()
On Error Resume Next
    
    With fg
        .Redraw = flexRDNone
        .Move .Left, .Top, Me.ScaleWidth - .Left * 2, Me.ScaleHeight - .Top - .Left
        '.ColWidth(0) = .Width - 60 * 18
        ExtendCustomColumn
        .Redraw = flexRDBuffered
    End With
    
    ' TLB 11/25/2008: disable Single/All checkbox when maximized
    ' (since the auto-resizing will error when form is maximized)
    If Me.WindowState = vbMaximized Then
        chkSingleBar.Enabled = False
    Else
        chkSingleBar.Enabled = True
    End If
    
    AutoSizeChart

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    SetIniFileProperty "ChartData", FontToString(fg.Font), "Fonts", g.strIniFile
    frmMain.DockPro.RemoveForm Me.Name
    Set m.DataGrid = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartData.Form.Unload", eGDRaiseError_Show
    Resume ErrExit

End Sub

Public Sub ShowData(ByVal nXbar&)
On Error GoTo ErrSection:

    Dim i&, nPrevRows&, strCaption$, strData$, strLabels$
    Dim aFields() As String
    Dim frm As Form
    Static strPrevLabels$
    
    If DockState(Me) = eHidden Then Exit Sub
    
    ' get data string for specified bar
    If nXbar < 0 Then m.strPrevData = ""
    Set frm = ActiveChart
    If Not frm Is Nothing Then
        strData = frm.Chart.GetDataWindowLabel(nXbar)
    End If
    If strData = m.strPrevData Then Exit Sub
        
    ' parse data string, and build labels string (for checking)
    aFields = Split(strData, "|")
    For i = 0 To UBound(aFields)
        strLabels = strLabels & Parse(aFields(i), vbTab, 1) & vbTab
    Next
    
    With fg
        .Redraw = flexRDNone
        If g.ChartGlobals.bChartDataSingleBar Then
            ' when showing single bar
            If Len(Trim(strData)) = 0 Then
                .Rows = .FixedRows
                ExtendCustomColumn
            ElseIf strLabels = strPrevLabels _
                And .Rows = .FixedRows + UBound(aFields) Then
                ' same labels and rows, so just replace values
                For i = 1 To UBound(aFields)
                    .TextMatrix(.FixedRows + i - 1, 1) = _
                        Parse(aFields(i), vbTab, 2)
                Next
            Else
                ' rebuild grid with labels and values
                .Rows = .FixedRows
                For i = 1 To UBound(aFields)
                    .AddItem aFields(i)
                Next
                ExtendCustomColumn
            End If
        Else
            ' when showing all bars, must reinit grid
            ' if # of rows or column labels have changed
            If frm Is Nothing Then
                InitGrid
            ElseIf frm.Chart.aXBar.Size <> .Rows - .FixedRows Then
                InitGrid
            ElseIf strLabels <> strPrevLabels Then
                InitGrid
            End If
            ' center and highlight the active bar
            If .Rows > 0 Then
                If nXbar < 0 Then
                    nXbar = .Rows - 2
                End If
                If nXbar >= 0 And nXbar < .Rows - 1 Then
                    .Row = nXbar + 1
                    .TopRow = 0
                    i = .Row - (.BottomRow - .TopRow) \ 2
                    If i < .FixedRows Then i = .FixedRows
                    .TopRow = i
                    .ShowCell .Row, 0
                End If
            End If
        End If
        .Redraw = flexRDBuffered
    End With
    
    ' put symbol in caption
    If Not frm Is Nothing Then
        strCaption = frm.Chart.ChartName
    End If
    If Len(strCaption) = 0 Then strCaption = "Chart Data"
    If Me.Caption <> strCaption Then Me.Caption = strCaption
    
    ' store things for next call
    strPrevLabels = strLabels
    m.strPrevData = strData
    m.nPrevX = nXbar
    Set frm = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartData.ShowData", eGDRaiseError_Raise

End Sub

Public Sub GenerateReport(ByVal vArgs As Variant)
On Error GoTo ErrSection:

    Dim bExtend As Boolean
    Dim alColWidths As New cGdArray
    Dim lIndex As Long
    
    With frmPrintPreview.vp
        .StartDoc
        DoPrintHeader
        
        .Text = "Symbol: " & Trim(Caption) & vbCrLf
        
        alColWidths.Create eGDARRAY_Longs, fg.Cols
        For lIndex = 0 To fg.Cols - 1
            alColWidths(lIndex) = fg.ColWidth(lIndex)
        Next lIndex
        bExtend = fg.ExtendLastCol
        
        If g.ChartGlobals.bChartDataSingleBar Then
            fg.ExtendLastCol = False
            fg.AutoSize 0, fg.Cols - 1
        End If
        
        If frmPrintPreview.GoingToFile Then
            frmPrintPreview.GridToTable fg
        Else
            .RenderControl = fg.hWnd
        End If
        
        fg.ExtendLastCol = bExtend
        
        For lIndex = 0 To fg.Cols - 1
            fg.ColWidth(lIndex) = alColWidths(lIndex)
        Next lIndex
        
        .EndDoc
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartData.GenerateReport", eGDRaiseError_Raise

End Sub

Public Function PrintMe() As Boolean
On Error GoTo ErrSection

    If fg.Rows > 550 Then
        If AskBox("h=Warning ; i=? ; b=+Yes|-No ; Printing this many rows may take a while.||Do you want to continue?") = "N" Then
            Exit Function
        End If
    End If

    PrintMe = frmPrintPreview.ShowMe("CNV ChartData", frmChartData, , , , , , True)
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmChartData.PrintMe", eGDRaiseError_Raise
    
End Function

' adjust all column widths to accomodate the custom "extend column"
Private Sub ExtendCustomColumn()
On Error GoTo ErrSection:

    Dim nTotal&, i&

    If Not g.ChartGlobals.bChartDataSingleBar Then Exit Sub

    With fg
        .ColHidden(kExtendedCol) = True
        .Redraw = flexRDBuffered '(so .ClientWidth will be correct)
        .Redraw = flexRDNone
        nTotal = 0 * Screen.TwipsPerPixelX
        For i = 0 To .Cols - 1
            If Not .ColHidden(i) Then
                nTotal = nTotal + .ColWidth(i)
            End If
        Next
        nTotal = .ClientWidth - nTotal
        If nTotal > 0 Then .ColWidth(kExtendedCol) = nTotal
        .ColHidden(kExtendedCol) = False
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartData.ExtendCustomColumn", eGDRaiseError_Raise

End Sub

Private Sub InitGrid()
On Error GoTo ErrSection:

    With fg
        .Redraw = flexRDNone
        
        .Editable = flexEDNone
        
        If g.ChartGlobals.bChartDataSingleBar Then
            .FlexDataSource = Nothing
            .Cols = 2
            .ColWidth(0) = fg.Width - 60 * 18
            .FixedCols = 0
            .ExplorerBar = flexExNone
            .ColAlignment(0) = flexAlignLeftCenter
            .ColAlignment(1) = flexAlignCenterCenter
            .ExtendLastCol = False
            .FixedRows = 1
            .TextMatrix(0, 0) = "Data"
            .TextMatrix(0, 1) = "Value"
            .Rows = .FixedRows
        Else
            .FlexDataSource = m.DataGrid
            .FixedRows = 1
            If .Cols > 0 Then
                .ColWidth(0) = 60 * 16
                .FrozenCols = 1
                .ColAlignment(-1) = flexAlignCenterCenter
            End If
            .ExtendLastCol = True
            .ExplorerBar = flexExMove
        End If
        
        ExtendCustomColumn
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartData.InitGrid", eGDRaiseError_Raise

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
    RaiseError "frmChartData.mnuChangeFont.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub mnuCursor_Click()
On Error GoTo ErrSection:

    chkAuto.Value = 1

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartData.mnuCursor.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub mnuMultipleBars_Click()
On Error GoTo ErrSection:

    chkSingleBar.Value = 0

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartData.mnuMultipleBars.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub mnuOnClick_Click()
On Error GoTo ErrSection:

    chkAuto.Value = 0

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartData.mnuOnClick.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub mnuSingleBar_Click()
On Error GoTo ErrSection:

    chkSingleBar.Value = 1

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartData.mnuSingleBar.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

