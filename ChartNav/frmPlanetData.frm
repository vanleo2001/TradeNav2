VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmPlanetData 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Planet Data"
   ClientHeight    =   2955
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   2550
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   2550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniButtonImageXP cmdSetBase 
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      Top             =   0
      Width           =   900
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
      Caption         =   "frmPlanetData.frx":0000
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmPlanetData.frx":0032
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmPlanetData.frx":00BA
      RightToLeft     =   0   'False
   End
   Begin VSFlex7LCtl.VSFlexGrid fg 
      Height          =   2715
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   2415
      _cx             =   4260
      _cy             =   4789
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
      Rows            =   11
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
   Begin HexUniControls.ctlUniCheckXP chkAuto 
      Height          =   220
      Left            =   60
      TabIndex        =   1
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
      Caption         =   "frmPlanetData.frx":00D6
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Pressed         =   -1  'True
      Tip             =   "frmPlanetData.frx":010A
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frmPlanetData.frx":01A2
      ShowFocus       =   -1  'True
      RightToLeft     =   0   'False
   End
   Begin VB.Menu mnuSettings 
      Caption         =   "Settings"
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
Attribute VB_Name = "frmPlanetData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type mPrivate
    strPrevData As String
    nPrevX As Long
    nPrevColWidth As Long
    Bars As New cGdBars
    bBaseSet As Boolean
    bShowBenchmark As Boolean

    ' Arrays for the SysNav engine:
    astrParms As New cGdArray       ' Parameters array for the engine
    aExpr As New cGdArray           ' Array of coded text expressions
    astrBarNames As New cGdArray    ' Array of bar names
    aArrayOfBars As New cGdArray    ' Array of bars structures
    aArrayOfResults As New cGdArray ' Array of results
End Type
Private m As mPrivate

Private Sub chkAuto_Click()
On Error GoTo ErrSection:

    g.ChartGlobals.bAutoChartData = -chkAuto

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPlanetData.chkAuto.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub chkAuto_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 Then
        m.bShowBenchmark = Not m.bShowBenchmark
        If Not m.bShowBenchmark Then Me.Caption = "Planet Data"
    End If

End Sub

Private Sub cmdSetBase_Click()
On Error GoTo ErrSection:
    
    SetBase
    MoveFocus fg
    Exit Sub
    
ErrSection:
    RaiseError "frmPlanetData.cmdSetBase", eGDRaiseError_Show
End Sub

Private Sub fg_AfterMoveRow(ByVal Row As Long, Position As Long)

    If Row <> Position Then
        ' clear expressions since the order has changed
        ClearExpr
    End If

End Sub

Private Sub fg_AfterSort(ByVal Col As Long, Order As Integer)

    ' clear expressions since the order has changed
    ClearExpr

End Sub

Private Sub fg_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    ExtendCustomColumn Col

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPlanetData.fg.AfterUserResize", eGDRaiseError_Show
    Resume ErrExit
End Sub

Private Sub fg_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim nRow&, nPos&, nFirstRow&
    
    With fg
        nFirstRow = .FixedRows
        nRow = .MouseRow
        If nRow >= nFirstRow And .Rows > nFirstRow + 1 Then
            .Row = nRow
            .Refresh
            nPos = .DragRow(nRow)
            If nPos <> nRow Then
                Cancel = True
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPlanetData.fg.BeforeMouseDown", eGDRaiseError_Show
    Resume ErrExit
End Sub

Private Sub fg_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    ' save current size of column (in case is after the extended column)
    m.nPrevColWidth = fg.ColWidth(Col)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPlanetData.fg.BeforeUserResize", eGDRaiseError_Show
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
    RaiseError "frmPlanetData.fg.MouseDown", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fg_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    Dim nRow&, nCol&, strTip$
    
  Exit Sub
    
    With fg
        nRow = .MouseRow
        nCol = .MouseCol
        If nCol = 2 Then
            strTip = "Can hit '.' to set the current bar as the 'Base'"
        End If
        .ToolTipText = strTip
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPlanetData.fg.MouseMove", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub Form_Deactivate()
On Error GoTo ErrSection:

    SetPrevActiveForm Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPlanetData.Form.Deactivate", eGDRaiseError_Show
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
    RaiseError "frmPlanetData.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = Asc(".") Then SetBase

End Sub

Private Sub Form_Load()
On Error GoTo ErrSection:

    Dim i&, strFont$
    
    mnuSettings.Visible = False
    mnuLastBar.Visible = False
    
    Me.Icon = Picture16(ToolbarIcon("ID_PlanetData"), , True)
    CenterTheForm Me
    
    g.Styler.StyleForm Me
    
    chkAuto.Value = Abs(g.ChartGlobals.bAutoChartData)
       
    SetupGrid fg, eGridMode_Grid
    
    ' Get font from INI file
    strFont = GetIniFileProperty("PlanetData", "", "Fonts", g.strIniFile)
    If strFont <> "" Then
        FontFromString fg.Font, strFont
    End If
       
    InitGrid
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPlanetData.Form.Load", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode = 0 Then
        frmMain.tbToolbar.Tools("ID_PlanetData").State = ssUnchecked
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPlanetData.Form.QueryUnload", eGDRaiseError_Show
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
    
    cmdSetBase.Left = fg.Left + fg.Width - cmdSetBase.Width - 30
    
    AutoSizeChart

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    Dim i&, strPlanets$
    
    For i = fg.FixedRows To fg.Rows - 1
        strPlanets = strPlanets & fg.TextMatrix(i, 0) & ";"
    Next

    ClearExpr
    SetIniFileProperty "PlanetData", FontToString(fg.Font), "Fonts", g.strIniFile
    SetIniFileProperty "Planets", strPlanets, "PlanetData", g.strIniFile
    frmMain.DockPro.RemoveForm Me.Name

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPlanetData.Form.Unload", eGDRaiseError_Show
    Resume ErrExit

End Sub

Public Sub ShowData(ByVal dDateTime#, Optional Bars As cGdBars = Nothing)
On Error GoTo ErrSection:

    Dim i&, dStartTime#, bSuccess As Boolean
    Static dPrevDateTime#

    If DockState(Me) = eHidden Then Exit Sub
    
    If dDateTime > 0 And Not Bars Is Nothing Then
        dStartTime = gdTickCount
    
        ' only need to set bar properties if symbol has changed
        If m.Bars.Prop(eBARS_SymbolID) <> Bars.Prop(eBARS_SymbolID) Then
            SetBarProperties m.Bars, Bars.Prop(eBARS_SymbolID)
        ElseIf dDateTime = dPrevDateTime Then
            Exit Sub ' if same symbol and date as previous call, don't need to recalc
        End If
        dPrevDateTime = dDateTime
        
        ' set bar period to match
        m.Bars.Prop(eBARS_Periodicity) = Bars.Prop(eBARS_Periodicity)
        
        ' run expressions for specified date
        m.Bars.Size = 1
        m.Bars(eBARS_DateTime, 0) = dDateTime
        bSuccess = RunExpr
    End If
    
    If Not bSuccess Then
        dPrevDateTime = -1
        ' clear grid values if not successful
        With fg
            .Redraw = flexRDNone
            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, 1) = ""
                .TextMatrix(i, 2) = ""
            Next
            .Redraw = flexRDBuffered
        End With
    ElseIf m.bShowBenchmark Then
        Me.Caption = Format(gdTickCount - dStartTime, "#0") & " ms"
    End If
       
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPlanetData.ShowData", eGDRaiseError_Raise

End Sub

Public Sub GenerateReport(ByVal vArgs As Variant)
On Error GoTo ErrSection:

    Dim bExtend As Boolean
    Dim alColWidths As New cGdArray
    Dim lIndex As Long
    
    With frmPrintPreview.vp
        .StartDoc
        DoPrintHeader
        
        '.Text = "Symbol: " & Trim(Caption) & vbCrLf
        
        alColWidths.Create eGDARRAY_Longs, fg.Cols
        For lIndex = 0 To fg.Cols - 1
            alColWidths(lIndex) = fg.ColWidth(lIndex)
        Next lIndex
        bExtend = fg.ExtendLastCol
        
        fg.ExtendLastCol = False
        fg.AutoSize 0, fg.Cols - 1
        
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
    RaiseError "frmPlanetData.GenerateReport", eGDRaiseError_Raise

End Sub

Public Function PrintMe() As Boolean
On Error GoTo ErrSection

    'If fg.Rows > 550 Then
    '    If AskBox("h=Warning ; i=? ; b=+Yes|-No ; Printing this many rows may take a while.||Do you want to continue?") = "N" Then
    '        Exit Function
    '    End If
    'End If

    PrintMe = frmPrintPreview.ShowMe("CNV PlanetData", frmPlanetData, , , , , , True)
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmPlanetData.PrintMe", eGDRaiseError_Raise
    
End Function

' adjust all column widths to accomodate the custom "extend column"
Private Sub ExtendCustomColumn(Optional ByVal nResizeCol As Long = -1)
On Error GoTo ErrSection:
    
    Dim i&, nTotal&, nDiff&, nExtCol&
    
    nExtCol = 0  ' column number of extended column
      
    With fg
        ' if column being resized is after the extended column,
        ' then change the width of the next visible column instead
        If nResizeCol >= nExtCol Then
            .Redraw = flexRDNone
            nDiff = .ColWidth(nResizeCol) - m.nPrevColWidth
            For i = nResizeCol + 1 To .Cols - 1
                If Not .ColHidden(i) Then
                    .ColWidth(i) = .ColWidth(i) - nDiff
                    Exit For
                End If
            Next
            m.nPrevColWidth = 0
        End If
        
        .ColHidden(nExtCol) = True
        .Redraw = flexRDBuffered '(so .ClientWidth will be correct)
        .Redraw = flexRDNone
        nTotal = 0 * Screen.TwipsPerPixelX
        For i = 0 To .Cols - 1
            If Not .ColHidden(i) Then
                nTotal = nTotal + .ColWidth(i)
            End If
        Next
        nTotal = .ClientWidth - nTotal
        If nTotal > 0 Then .ColWidth(nExtCol) = nTotal
        .ColHidden(nExtCol) = False
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPlanetData.ExtendCustomColumn", eGDRaiseError_Raise
End Sub

Private Sub InitGrid()
On Error GoTo ErrSection:

    Dim i&, strPlanets$, aPlanets As New cGdArray
    
    ' get planets
    strPlanets = GetIniFileProperty("Planets", "", "PlanetData", g.strIniFile)
    If Len(Trim(strPlanets)) = 0 Then
        strPlanets = "Sun;Mercury;Venus;Moon;Mars;Jupiter;Saturn;Uranus;Neptune;Pluto;"
    End If
    aPlanets.SplitFields strPlanets, ";"

    With fg
        .Redraw = flexRDNone
        
        SetupGrid fg, eGridMode_Grid
        .Editable = flexEDNone
        .ExplorerBar = flexExMoveRows
        
        .Cols = 3
        .ColWidth(0) = fg.Width - 60 * 18
        .FixedCols = 0
        .ColAlignment(0) = flexAlignLeftCenter
        '.ColAlignment(1) = flexAlignCenterCenter
        .ExtendLastCol = False
        .FixedRows = 1
        
        .TextMatrix(0, 0) = "Planet"
        .TextMatrix(0, 1) = "Longitude"
        .TextMatrix(0, 2) = "from Base"
        .ColDataType(1) = flexDTDouble
        .ColDataType(2) = flexDTDouble
        .ColFormat(1) = "#0.00"
        .ColFormat(2) = "#0.00"
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        
        ' display row for each planet
        .Rows = .FixedRows
        For i = 0 To aPlanets.Size - 1
            strPlanets = Trim(aPlanets(i))
            If Len(strPlanets) > 0 Then
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 0) = strPlanets
                If UCase(strPlanets) = "MOON" Then
                    .Row = .Rows - 1 '(default for highlight)
                End If
            End If
        Next
        
        ExtendCustomColumn
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Set aPlanets = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "frmPlanetData.InitGrid", eGDRaiseError_Raise

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
    RaiseError "frmPlanetData.mnuChangeFont.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub mnuCursor_Click()
On Error GoTo ErrSection:

    chkAuto.Value = 1

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPlanetData.mnuCursor.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub mnuOnClick_Click()
On Error GoTo ErrSection:

    chkAuto.Value = 0

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPlanetData.mnuOnClick.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub ClearExpr()
On Error GoTo ErrSection:
    
    Dim rc&, i&
    
    ' clear expressions array
    m.aExpr.Clear
    
    ' Destroy all the result arrays (array of arrays)
    For i = 0 To m.aArrayOfResults.Size - 1
        gdDestroyArray m.aArrayOfResults(i)
    Next
    m.aArrayOfResults.Clear
    
    ' clear the expression evaluator
    m.astrParms.Clear
    m.astrParms(0) = "frmPlanetData"  ' 0) Expression set name
    SetupExpressions m.astrParms
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPlanetData.ClearExpr", eGDRaiseError_Raise
End Sub

Private Function InitExpr() As Boolean
On Error GoTo ErrSection:

    Dim i&, strText$
    Dim hArray As Long                  ' Array handle
    Dim rc As Long                      ' Return code from function calls
    
    ' first clear everything properly before reinitializing things
    ClearExpr
    
    ' Create the arrays
    m.astrBarNames.Create eGDARRAY_Strings
    m.aArrayOfBars.Create eGDARRAY_Longs
    m.aExpr.Create eGDARRAY_Strings
    m.aArrayOfResults.Create eGDARRAY_Longs
    
    ' Init the array of bars
    m.astrBarNames(0) = "Market1"
    m.aArrayOfBars.Num(0) = m.Bars.BarsHandle
        
    ' Build array of expressions and results
    For i = fg.FixedRows To fg.Rows - 1
        strText = Trim(fg.TextMatrix(i, 0))
        strText = "~01010CalcPlanet ~16001( ~07007Market1 ~22001, ~20004" & Left(strText & Space(4), 4) _
            & " ~22001, ~20001- ~22001, ~130010 ~22001, ~130010 ~22001, ~130010 ~22001, ~130011 ~22001, ~130010 ~17001)"
        m.aExpr.Add strText
        
        hArray = gdCreateArray(eGDARRAY_Doubles, 0)
        m.aArrayOfResults.Add hArray
    Next

    ' Init the expression tree
    If m.aExpr.Size > 0 Then
        InitExpr = SetupExpressions(m.astrParms, m.astrBarNames, m.aExpr)
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmPlanetData.InitExpr", eGDRaiseError_Raise
End Function

Private Function RunExpr() As Boolean
On Error GoTo ErrSection:

    Dim i&, iRow&, dValue#
    Dim hArray As Long                  ' Array handle
    Dim rc As Long                      ' Return code from function calls
    
    If m.aExpr.Size = 0 Then
        ' need to init the expressions
        If Not InitExpr Then Exit Function
    End If
    
    If m.Bars(eBARS_DateTime, 0) > 0 Then
        m.astrParms.Size = 1
        rc = RunExpressions(m.astrParms.ArrayHandle, _
            m.astrBarNames.ArrayHandle, m.aArrayOfBars.ArrayHandle, _
            m.aArrayOfResults.ArrayHandle, ByVal 0&, ByVal 0&)
        If rc = 0 Then
            ' display values
            With fg
                .Redraw = flexRDNone
                For iRow = .FixedRows To .Rows - 1
                    i = iRow - .FixedRows
                    hArray = m.aArrayOfResults.Num(i)
                    dValue = gdGetNum(hArray, 0)
                    .TextMatrix(iRow, 1) = dValue
                    If m.bBaseSet Then
                        .TextMatrix(iRow, 2) = dValue - .RowData(iRow)
                    End If
                Next
                .Redraw = flexRDBuffered
            End With
            RunExpr = True
        End If
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmPlanetData.RunExpr", eGDRaiseError_Raise
End Function

Public Sub SetBase()
On Error GoTo ErrSection:

    Dim i&
    
    With fg
        For i = .FixedRows To .Rows - 1
            .RowData(i) = ValOfText(.TextMatrix(i, 1))
            .TextMatrix(i, 2) = 0
        Next
    End With
    m.bBaseSet = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPlanetData.SetBase", eGDRaiseError_Raise
End Sub

