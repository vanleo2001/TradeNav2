VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmTemplatePage 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   3000
      Top             =   1050
   End
Begin HexUniControls.ctlUniFrameWL fraButtons
VistaStyle      =   0   'False
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      Height          =   250
      Left            =   2160
      TabIndex        =   1
      Top             =   3165
      Width           =   2280
      Begin VB.PictureBox pbLeft 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   250
         Left            =   240
         Picture         =   "frmTemplatePage.frx":0000
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   3
         Top             =   0
         Width           =   250
      End
      Begin VB.PictureBox pbRight 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   250
         Left            =   900
         Picture         =   "frmTemplatePage.frx":038A
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   2
         Top             =   0
         Width           =   250
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid fg 
      Height          =   2535
      Left            =   270
      TabIndex        =   0
      Top             =   135
      Width           =   1920
      _cx             =   3387
      _cy             =   4471
      _ConvInfo       =   1
      Appearance      =   2
      BorderStyle     =   0
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
      GridLines       =   0
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   2
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   0
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
Attribute VB_Name = "frmTemplatePage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type mPrivate
    eFormMode As eTemplateFormMode
    aItems As cGdArray
    
    idxFirstItem As Long            'index into aItems array of first item shown in grid
    idxLastItem As Long             'index in aItems array of last item shown on grid
    idxGridStart As Long            'index of first row in grid to put template/page names into
    iGridCols As Long               'number of columns in grid (either 1 or 2)
        
    iMouseSelRow As Long
    iMouseSelCol As Long
    iMouseRow As Long
    iMouseCol As Long
        
    bApplySel As Boolean
    
    strCurrent As String            'name of current page or template
    strBtnID As String
    frm As Form                     'form that toolbar is on
    
    bClearBtnOnExit As Boolean
    bSkipApply As Boolean
End Type

Private m As mPrivate

Public Function ShowMe(frmSource As Form, ByVal eMode As eTemplateFormMode, ByVal nX&, ByVal nY&, _
    Optional ByVal strID$ = "", Optional ByVal strCaption$ = "")

    Dim pt As POINTAPI
    Dim wp As WINDOWPLACEMENT
    Dim strText$, i&, X&, Y&
    
    Dim iTbWidth&, iDiff&
    
    Dim bShowSCP As Boolean
    Dim bShowPublish As Boolean
    
    fraButtons.Visible = False
    fraButtons.BackColor = vbButtonFace
        
    Set m.frm = frmSource
    
    If m.frm Is Nothing Then Exit Function
    
    If TypeOf m.frm Is frmTbMoreButtons Then
        Set m.frm = frmTbMoreButtons.FormSource
        m.bClearBtnOnExit = False
'        m.bSkipApply = True
    Else
        m.bClearBtnOnExit = True
        m.bSkipApply = False
    End If
    
    If Not IsFrmChart(m.frm) And Not TypeOf m.frm Is frmMain Then
        Exit Function
    End If
        
    If m.frm.Height < 3200 Then
        'this is minimum height needed for menu items and 2 grid rows (1 for name, 1 for buttons)
        strText = "Please increase the height of Trade Navigator's application window and try again."
        InfBox strText, "I", , "Manage templates/pages"
        Exit Function
    End If
    
    m.strBtnID = strID
    m.strCurrent = ""
    m.idxFirstItem = 0
    
    m.iMouseRow = -1
    m.iMouseCol = -1
    m.iMouseSelRow = -1
    m.iMouseCol = -1
    
    m.eFormMode = eMode
    m.bApplySel = False
    
    If TypeOf m.frm Is frmMain Then
        iTbWidth = m.frm.tbToolbar.GetDockWidth(ssDockedRight) * Screen.TwipsPerPixelX
    End If
    
    If m.eFormMode = eMode_Pages Then
        strText = "P"
        If Len(g.strChartPage) = 0 Then
            m.idxGridStart = 4
        Else
            m.idxGridStart = 3
        End If
        Set m.aItems = GetAllowedList(strText)
        
        If Len(g.strChartPage) > 0 Then m.strCurrent = g.strChartPage
    ElseIf eMode = eMode_Templates Then
        strText = "T"
        m.idxGridStart = 2
        Set m.aItems = GetAllowedList(strText)
        If Not ActiveChart.Chart Is Nothing Then m.strCurrent = ActiveChart.Chart.TemplateApplied
    ElseIf eMode = eMode_SecSubCom And Len(strID) > 0 Then
        m.idxGridStart = 0
        Set m.aItems = New cGdArray
        
        ToolbarSectorMenu frmMain.tbToolbar, strID, m.aItems, False, True   'Not (strID = "ID_Components")  - 5564

        If Len(strCaption) > 0 Then
            m.strCurrent = strCaption
        ElseIf Not ActiveChart Is Nothing Then
            m.strCurrent = ActiveChart.Chart.Symbol
        End If
    Else
        Exit Function           'theoretically should never get here
    End If
        
    If m.aItems Is Nothing Then Exit Function
        
    pt.X = nX
    pt.Y = nY
    If TypeOf frmSource Is frmTbMoreButtons Then
        frmMain.ToolBarBtnSizeGet kTbGeneral, X, Y
        pt.Y = pt.Y + Y
        ClientToScreen frmTbMoreButtons.FormSource.hWnd, pt
    Else
        ClientToScreen m.frm.hWnd, pt
    End If
    
    pt.X = pt.X * Screen.TwipsPerPixelX
    pt.Y = pt.Y * Screen.TwipsPerPixelY
    Me.Move pt.X, pt.Y

    Me.Height = m.frm.Height - 1800
    
    i = BestWidth()
    If i <= 0 Then
        Me.Width = 4000       'approx width for 50 chars
    Else
        Me.Width = i
    End If
    
    m.iGridCols = 1
    bShowPublish = HasGold(False) And FileExist(g.strAppPath & "\SCP.flg")
    bShowSCP = bShowPublish Or (HasGold(False) And HasModule("SCP_*"))
    
    With fg
        .BorderStyle = flexBorderNone
        .BackColor = vbButtonFace
        .HighLight = flexHighlightNever
        '.HighLight = flexHighlightAlways
        .SelectionMode = flexSelectionFree
        .MergeCells = flexMergeFree
        .Editable = flexEDNone
        .AllowSelection = False
        .AllowBigSelection = False
        
        .RowHeightMin = fraButtons.Height
        .RowHeightMax = fraButtons.Height
        .Rows = Int(Me.ScaleHeight / .RowHeightMin)
                
        If m.aItems.Size = 0 Then
            .Rows = m.idxGridStart
        ElseIf .Rows - m.idxGridStart < m.aItems.Size And Me.Width * 2 < m.frm.Width - iTbWidth Then
            m.iGridCols = 2
            Me.Width = Me.Width * 2
        End If
        
        .Cols = m.iGridCols
        If m.iGridCols = 1 Then
            .ColWidth(0) = Me.Width
        Else
            .ColWidth(0) = Me.Width / 2
            .ColWidth(1) = Me.Width / 2
        End If
        
        If m.eFormMode = eMode_Templates Then
            If .Cols = 2 Then
                .TextMatrix(0, 0) = "<Manage chart templates>"
                .TextMatrix(0, 1) = "<Manage chart templates>"
                
                .TextMatrix(1, 0) = "<Copy settings to other charts>"
                .TextMatrix(1, 1) = "<Copy settings to other charts>"
                
                .MergeRow(0) = True
                .MergeRow(1) = True
            Else
                .TextMatrix(0, 0) = "<Manage chart templates>"
                .TextMatrix(1, 0) = "<Copy settings to other charts>"
            End If
            
            .Select 1, 0, 1, .Cols - 1
            .CellBorder RGB(188, 188, 188), -1, -1, -1, 1, -1, -1
        ElseIf m.eFormMode = eMode_Pages Then
            If .Cols = 2 Then
                .TextMatrix(0, 0) = "<Manage chart pages>"
                .TextMatrix(0, 1) = "<Manage chart pages>"
                
                .TextMatrix(1, 0) = "<Save chart page>"
                .TextMatrix(1, 1) = "<Save chart page>"
                
                If g.ChartGlobals.bMyPageFeature Then
                    .TextMatrix(2, 0) = "<My Page>"
                    .TextMatrix(2, 1) = "<My Page>"
                Else
                    .TextMatrix(2, 0) = "<Create new chart page>"
                    .TextMatrix(2, 1) = "<Create new chart page>"
                    If bShowSCP Then
                        .TextMatrix(3, 0) = "<Load shared chart page>"
                        .TextMatrix(3, 1) = "<Load shared chart page>"
                        .MergeRow(3) = True
                        m.idxGridStart = m.idxGridStart + 1
                    End If
                    If bShowPublish Then
                        If bShowSCP Then
                            .TextMatrix(4, 0) = "<Publish shared chart page>"
                            .TextMatrix(4, 1) = "<Publish shared chart page>"
                            .MergeRow(4) = True
                        Else
                            .TextMatrix(3, 0) = "<Publish shared chart page>"
                            .TextMatrix(3, 1) = "<Publish shared chart page>"
                            .MergeRow(3) = True
                        End If
                        m.idxGridStart = m.idxGridStart + 1
                    End If
                End If
                
                .MergeRow(0) = True
                .MergeRow(1) = True
                .MergeRow(2) = True
            Else
                .TextMatrix(0, 0) = "<Manage chart pages>"
                .TextMatrix(1, 0) = "<Save chart page>"
                
                If g.ChartGlobals.bMyPageFeature Then
                    .TextMatrix(2, 0) = "<My Page>"
                Else
                    .TextMatrix(2, 0) = "<Create new chart page>"
                    If bShowSCP Then
                        .TextMatrix(3, 0) = "<Load shared chart page>"
                        m.idxGridStart = m.idxGridStart + 1
                        If .Rows < m.idxGridStart Then .Rows = m.idxGridStart   '6923
                    End If
                    If bShowPublish Then
                        If bShowSCP Then
                            .TextMatrix(4, 0) = "<Publish shared chart page>"
                        Else
                            .TextMatrix(3, 0) = "<Publish shared chart page>"
                        End If
                        m.idxGridStart = m.idxGridStart + 1
                        If .Rows < m.idxGridStart Then .Rows = m.idxGridStart   '6923
                    End If
                End If
            End If
            
            If g.ChartGlobals.bMyPageFeature Then
                .RowHidden(0) = True
                .RowHidden(1) = True
            End If
            
            If Len(g.strChartPage) = 0 Then
                If g.ChartGlobals.bMyPageFeature Then
                    m.idxGridStart = 3
                Else
                    .TextMatrix(m.idxGridStart - 1, 0) = "(unnamed)"
                    If g.nColorTheme = kDarkThemeColor Then
                        .Cell(flexcpForeColor, m.idxGridStart - 1, 0) = vbGreen
                    Else
                        .Cell(flexcpForeColor, m.idxGridStart - 1, 0) = vbBlue
                    End If
                    .Cell(flexcpFontBold, m.idxGridStart - 1, 0) = True
                End If
            End If
            
            .Select 2, 0, 2, .Cols - 1
            .CellBorder RGB(188, 188, 188), -1, -1, -1, 1, -1, -1
        End If

        .Select 0, 0

    End With
        
    iDiff = (Me.Left + Me.Width + iTbWidth) - (m.frm.Left + m.frm.Width)
    If iDiff > 0 Then
        Me.Left = Me.Left - iDiff
    End If
    
    FillTemplatePageNames
    
    If fg.Rows <= 5 Then
        If fg.Height < Me.ScaleHeight + 500 Then Me.Height = fg.Height + 100
    End If
    
    pbLeft.BackColor = g.nColorTheme
    pbRight.BackColor = g.nColorTheme
    
    ShowForm Me
        
End Function

Private Sub FillTemplatePageNames()

    Dim i&, j&, iCol&, iRowLast&
    Dim iPageRow&, iPageCol&
    Dim strText$, strSymbol$
    
    Dim bFound As Boolean
    Dim bVisible As Boolean
    
    iPageRow = -1
    iPageCol = -1
    
    iCol = 0
    j = m.idxGridStart
        
    fg.Redraw = flexRDNone
    
    'clear out grid
    With fg
        For i = m.idxGridStart To .Rows - 1
            .TextMatrix(i, 0) = ""
            If .Cols = 2 Then .TextMatrix(i, 1) = ""
        Next
        iRowLast = .Rows    '.Rows - 1
    End With
    
    If g.nColorTheme = kDarkThemeColor Then fg.Cell(flexcpForeColor, 0, 0) = vbWhite
    'JM 12-18-2015: need to call this here because the grids are getting loaded before showing the form
    FixFormControls Me, 0

    
    For i = m.idxFirstItem To m.aItems.Size - 1
        With fg
            .Cell(flexcpForeColor, 0, 0, .Rows - 1, .Cols - 1) = .Cell(flexcpForeColor, 0, 0)
            .Cell(flexcpFontBold, 0, 0, .Rows - 1, .Cols - 1) = False
            If j < .Rows Then
                m.idxLastItem = i
                strText = Parse(m.aItems(i), vbTab, 1)
                .TextMatrix(j, iCol) = strText
                j = j + 1
            ElseIf iCol = 0 Then
                If m.iGridCols = 2 Then
                    j = m.idxGridStart
                    iCol = 1
                    m.idxLastItem = i
                    strText = Parse(m.aItems(i), vbTab, 1)
                    .TextMatrix(j, iCol) = strText
                    j = j + 1
                Else
                    Exit For
                End If
            Else
                Exit For
            End If
            
            'save row & col of current template or page name
            If Not bFound Then
                If Len(m.strCurrent) > 0 Then
                    If m.strBtnID = "ID_Sectors" Or m.strBtnID = "ID_Subsectors" Or m.strBtnID = "ID_Components" Then
                        strSymbol = Parse(strText, ":", 1)
                        If strSymbol = m.strCurrent Then
                            iPageRow = j - 1
                            iPageCol = iCol
                            bFound = True
                        End If
                    ElseIf strText = m.strCurrent Then
                        iPageRow = j - 1
                        iPageCol = iCol
                        bFound = True
                    End If
                Else
                    bFound = True           'so don't keep checking
                End If
            End If
            
            'add hot key number 0-9 to first template names
            If m.eFormMode = eMode_Templates Then
                If i < 10 Then
                    .TextMatrix(j - 1, iCol) = Str(i) & ": " & strText
                End If
            End If
        End With
    Next
    
    If Len(m.strCurrent) = 0 And m.eFormMode = eMode_Pages Then
        If fg.TextMatrix(m.idxGridStart - 1, 0) = "(unnamed)" Then
            iPageRow = m.idxGridStart - 1
        ElseIf g.ChartGlobals.bMyPageFeature Then
            iPageRow = 2
        End If
        iPageCol = 0
    End If
    
    'highlight & bold current template or page name
    With fg
        If iPageRow >= .FixedRows And iPageCol >= .FixedCols Then
            If g.nColorTheme = kDarkThemeColor Then
                .Cell(flexcpForeColor, iPageRow, iPageCol) = vbGreen
            Else
                .Cell(flexcpForeColor, iPageRow, iPageCol) = vbBlue
            End If
            .Cell(flexcpFontBold, iPageRow, iPageCol) = True
        End If
    End With
        
    If m.idxFirstItem > 0 Then
        bVisible = True
        pbLeft.Visible = True
        pbLeft.Enabled = True
    Else
        pbLeft.Visible = False
        pbLeft.Enabled = False
    End If
    
    If m.idxLastItem < m.aItems.Size - 1 Then
        bVisible = True
        pbRight.Visible = True
        pbRight.Enabled = True
    Else
        pbRight.Visible = False
        pbRight.Enabled = False
    End If
            
    With fg
        If j > m.idxLastItem Then .Rows = j
        
        .Move 0, 0, Me.ScaleWidth, .Rows * .RowHeightMin
                
        i = .Cell(flexcpTop, .Rows - 1, 0) + .RowHeightMin
        If i <> .Height Then
            .Height = i
            If bVisible Then
                Me.Height = .Height + fraButtons.Height + Me.Height - Me.ScaleHeight
            Else
                Me.Height = .Height + Me.Height - Me.ScaleHeight
            End If
        End If
    End With
        
    fg.Redraw = flexRDBuffered
    
    'do this last so resize code won't execute until this routine is done
    fraButtons.Visible = bVisible
        
End Sub

Private Sub fg_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If vbKeyReturn = KeyCode Then ApplySelection
        
End Sub

Private Sub fg_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    With fg
        If .MouseRow >= m.idxGridStart And .MouseRow < .Rows Then
            If .MouseCol >= 0 And .MouseCol < .Cols Then
                m.iMouseSelCol = .MouseCol
                m.iMouseSelRow = .MouseRow
            End If
        End If
    End With

End Sub

Private Sub fg_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
            
    Dim iRow&, iCol&
    
    iRow = -1
    iCol = -1
    
    'm.bSkipApply = False
    With fg
        If .MouseCol >= 0 And .MouseCol < .Cols Then
            If .MouseRow >= 0 And .MouseRow < .Rows Then
                iRow = .MouseRow
                iCol = .MouseCol
            End If
        ElseIf .Col >= 0 And .Col < .Cols Then
            If .Row >= 0 And .Row < .Rows Then
                iRow = .Row
                iCol = .Col
            End If
        End If
        
        If iCol >= 0 And iCol < .Cols Then
            If iRow >= 0 And iRow < .Rows Then
                If m.iMouseRow < 0 Then
                    m.iMouseRow = .MouseRow
                    m.iMouseCol = .MouseCol
                Else
                    '.Cell(flexcpBackColor, m.iMouseRow, m.iMouseCol) = vbButtonFace
                    '.Cell(flexcpBackColor, iRow, iCol) = vbHighlight
                    .Row = iRow
                    .Col = iCol
                    .RowSel = iRow
                    .ColSel = iCol
                    m.iMouseRow = iRow
                    m.iMouseCol = iCol
                End If
            End If
        End If
    End With
    
    HandleToolTip X, Y
    
End Sub

Private Sub fg_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If m.iMouseRow < 0 Or m.iMouseSelRow < 0 Or (m.iMouseRow = m.iMouseSelRow And m.iMouseCol = m.iMouseSelCol) Then
        ApplySelection
    ElseIf m.iMouseSelRow > 0 Then
        MoveCell
    Else
        ToolBarNotify
    End If

End Sub

Private Sub ApplySelection()
    
    Dim strName$, nSymbolID&

    If m.bSkipApply Then
        'skip this routine once when this form is invoked from the more buttons
        m.bSkipApply = False
        Exit Sub
    End If

    m.bApplySel = True      'prevent the lost_focus event from notifying the toolbar to unload form
    
    Dim iRow&, iCol&
    
'when up/down arrow keys used, m.iMouseRow & m.iMouseCol will not have valid values
    If fg.Row >= 0 And fg.Row < fg.Rows Then
        iRow = fg.Row
    ElseIf m.iMouseRow >= 0 And m.iMouseRow < fg.Rows Then
        iRow = m.iMouseRow
    Else
        iRow = -1
    End If
    
    If fg.Col >= 0 And fg.Col < fg.Cols Then
        iCol = fg.Col
    ElseIf m.iMouseCol >= 0 And m.iMouseCol < fg.Cols Then
        iCol = m.iMouseCol
    Else
        iCol = -1
    End If
    
    If iRow >= 0 And iRow < fg.Rows Then
        If iCol >= 0 And iCol < fg.Cols Then
            If m.eFormMode = eMode_Pages Then
                strName = fg.TextMatrix(iRow, iCol)
            ElseIf m.eFormMode = eMode_Templates Or m.eFormMode = eMode_SecSubCom Then
                strName = GetName(iRow, iCol)
            End If
        End If
    End If
    
'JM 11-16-2012: Pete ran into an issue while in Europe (Czech Repub I think) with a user groups.
'   the issue has to do with templates near the bottom of the list not loading because the
'   fg.Rows value is less than the actual # of rows in the grid (i.e fg.rows report 50 when
'   there are actually 53 rows. Commenting out the DoEvens appear to fix this for whatever reason.
'   This appears machine specific because I can duplicate it, but Dave (also on XP) cannot.
    
    Me.Hide
'    DoEvents

    Dim bPublish As Boolean
    Dim bLoadShare As Boolean
    
    With fg
        If iRow >= 0 And iRow < .Rows Then
            'set loadshare or publish up front so don't have to keep doing string compares
            If iRow = 3 Or iRow = 4 Then
                If InStr(fg.TextMatrix(iRow, 0), "Load") <> 0 Then
                    bLoadShare = True
                ElseIf InStr(fg.TextMatrix(iRow, 0), "Publish") <> 0 Then
                    bPublish = True
                End If
            End If
            
            If iCol >= 0 And iCol < .Cols Then
                If m.eFormMode = eMode_Templates Then
                    If iRow = 0 Then
                        frmTemplates.ShowMe eMode_Templates, ActiveChart.Chart  'manage templates
                    ElseIf iRow = 1 Then
                        If Not ActiveChart Is Nothing Then
                            CopySettingsToOtherCharts ActiveChart       'copy settings
                        End If
                    ElseIf Not ActiveChart Is Nothing Then
                        If Not ActiveChart.Chart Is Nothing Then
                            If Len(strName) > 0 Then
                                ActiveChart.Chart.TemplateApply strName     'apply template
                            End If
                        End If
                    End If
                ElseIf m.eFormMode = eMode_Pages Then
                    If iRow = 0 Then
                        frmTemplates.ShowMe eMode_Pages, ActiveChart.Chart          'manage pages
                    ElseIf iRow = 1 Then
                        SaveChartPage ""            'save chart page
                    ElseIf iRow = 2 Then
                        LoadChartPage ""            'create new chart page
                        If Not g.ChartGlobals.frmActiveNonDetached Is Nothing Then
                            FormResize frmMain          '6978
                        End If
                    ElseIf iRow = 3 And (bLoadShare Or bPublish) Then
                        If bLoadShare Then
                            DisplaySharedChartPages
                        Else
                            PublishSharedChartPage
                        End If
                    ElseIf iRow = 4 And bPublish Then
                        PublishSharedChartPage
                    ElseIf iCol >= 0 And iCol <= 1 Then
                        LoadChartPage strName       'load chart page
                        If g.ChartGlobals.frmActiveNonDetached Is Nothing Then
                            FormResize frmMain      'theoretically should not get here
                        ElseIf g.ChartGlobals.frmActiveNonDetached.WindowState = vbMaximized Then
                            FormResize frmMain      '5168, 5080
                        Else
                            frmMain.tmrAutoResize.Enabled = True        '5463
                        End If
                    End If
                ElseIf m.eFormMode = eMode_SecSubCom Then
                    If Not ActiveChart Is Nothing Then
                        strName = Parse(strName, ":", 1)
                        nSymbolID = GetSymbolID(strName)
                        If nSymbolID <> 0 Then
                            ActiveChart.Chart.SetSymbol nSymbolID, True
                        Else
                            Beep
                        End If
                    End If
                End If
            End If
        End If
    End With
    
    ToolBarNotify

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    ToolBarNotify
    If Not ActiveChart Is Nothing Then SendMessage ActiveChart.hWnd, WM_ACTIVATE, 1, 0
    
End Sub

Private Sub Form_Resize()
On Error Resume Next:

    If fraButtons.Visible Then
        With fg
            If .RowHeightMin > 0 And .Rows > 0 Then
                fraButtons.Move 0, .Top + .Height, Me.Width
                pbRight.Move fraButtons.Width - pbRight.Width - 180, 5
                pbLeft.Move pbRight.Left - pbLeft.Width - 120
            End If
        End With
    End If

End Sub

Private Sub pbLeft_Click()

    MoveFocus fg
    If tmr.Enabled Then Exit Sub
    
    pbLeft.BorderStyle = 1
    tmr.Enabled = True

End Sub

Private Sub pbRight_Click()
        
    MoveFocus fg
    If tmr.Enabled Then Exit Sub
    
    pbRight.BorderStyle = 1
    tmr.Enabled = True

End Sub

Private Sub tmr_Timer()
        
    Dim i&, strText$
        
    If pbRight.BorderStyle = 1 Then
        pbRight.BorderStyle = 0
        m.idxFirstItem = m.idxFirstItem + (fg.Rows - m.idxGridStart)
    End If
    
    If pbLeft.BorderStyle = 1 Then
        pbLeft.BorderStyle = 0
        m.idxFirstItem = m.idxFirstItem - (fg.Rows - m.idxGridStart)
    End If
    
    FillTemplatePageNames
    Form_Resize
    
    tmr.Enabled = False

End Sub

' Note: X and Y should only be passed in from the mousemove event of the fg grid
Private Sub HandleToolTip(Optional ByVal X As Long = -1, Optional ByVal Y As Long = -1)

    Dim i&, strType$
    Static nRow&, nCol&, nPrevOrderRow&, strTip$
    
    With fg
        ' We only need to do something when the mouse row or column has changed
        ' (when mouse is not over the grid, MouseRow and MouseCol will be -1)
        If .MouseRow <> nRow Or .MouseCol <> nCol Then
            If .MouseCol < 0 Or .MouseRow < 0 Then
                nRow = -1
                nCol = -1
                strTip = ""
            Else
                nRow = .MouseRow
                nCol = .MouseCol
            End If
            .ToolTipText = "" '(to force the tip to move whenever Row or Col has changed)
                        
            If nRow >= .FixedRows And nRow < .Rows Then
                If nCol >= .FixedCols And nCol < .Cols Then
                    If m.eFormMode = eMode_SecSubCom And m.strBtnID <> "ID_Components" Then
                        strTip = g.SymbolPool.Desc(g.SymbolPool.PoolRecForSymbol(.TextMatrix(nRow, nCol)))      '5176
                    Else
                        strTip = .TextMatrix(nRow, nCol)
                    End If
                End If
            End If
        End If
                
        If strTip <> .ToolTipText Then
            .ToolTipText = "" '(to force it to move to correct spot after updating
            .ToolTipText = strTip
        End If
    End With

End Sub

Public Sub ToolBarNotify()

    Dim i&
    Dim aButtons As cGdArray
    Dim oButton As cPicBoxButton

    If Not m.bClearBtnOnExit Then Exit Sub

    If Not m.frm Is Nothing Then
        If TypeOf m.frm Is frmMain Or IsFrmChart(m.frm) Then
            Set aButtons = m.frm.TbButtonsArray(kTbGeneral)
            If Not aButtons Is Nothing Then
                For i = 0 To aButtons.Size - 2
                    Set oButton = aButtons(i)
                    If Not oButton Is Nothing Then
                        If oButton.BtnID = m.strBtnID Then
                            oButton.BtnClearNow m.frm.pbTbBack(oButton.PicboxIndex), aButtons
                            Exit For
                        End If
                    End If
                Next
            End If
        End If
    End If
        

End Sub

Private Sub MoveCell()

    Dim i&, strFrom$, strTo$, strFile$, strText$
    Dim aTemp As New cGdArray, aFields As New cGdArray

    If m.iMouseRow >= m.idxGridStart And m.iMouseRow < fg.Rows Then
        If m.iMouseSelRow >= m.idxGridStart And m.iMouseSelRow < fg.Rows Then
            If m.iMouseRow <> m.iMouseSelRow Or m.iMouseCol <> m.iMouseSelCol Then
                'strFrom = fg.TextMatrix(m.iMouseSelRow, m.iMouseSelCol)
                strFrom = GetName(m.iMouseSelRow, m.iMouseSelCol)
            End If
        End If
    End If
    
    If Len(strFrom) = 0 Then Exit Sub
    
    With fg
        'strTo = .TextMatrix(m.iMouseRow, m.iMouseCol)
        strTo = GetName(m.iMouseRow, m.iMouseCol)
        For i = 0 To m.aItems.Size - 1
            If Parse(m.aItems(i), vbTab, 1) = strFrom Then
                strFrom = m.aItems(i)
                aFields.SplitFields strFrom, vbTab
                Exit For
            End If
        Next
    End With
    
    For i = 0 To m.aItems.Size - 1
        If Parse(m.aItems(i), vbTab, 1) = strTo Then
            aFields(4) = Parse(m.aItems(i), vbTab, 5)
            strFrom = aFields.JoinFields(vbTab)
            aTemp.Add strFrom
            aTemp.Add m.aItems(i)
        ElseIf Parse(m.aItems(i), vbTab, 1) = aFields(0) Then
            'do nothing
        Else
            aTemp.Add m.aItems(i)
        End If
    Next

    m.iMouseSelRow = -1
    m.iMouseSelCol = -1
        
    SaveTempPageList aTemp, m.eFormMode

    Set aTemp = Nothing
    Set m.aItems = Nothing

    If m.eFormMode = eMode_Pages Then
        Set m.aItems = GetAllowedList("P")
    ElseIf m.eFormMode = eMode_Templates Then
        Set m.aItems = GetAllowedList("T")
    End If
        
    If m.aItems Is Nothing Then
        ToolBarNotify
        Exit Sub
    End If

    FillTemplatePageNames
    
    
End Sub

Public Property Get FormMode() As eTemplateFormMode
    FormMode = m.eFormMode
End Property

Private Function GetName(ByVal iRow&, ByVal iCol&) As String

    Dim strName$, strHotKey$
    

    With fg
        If iRow >= .FixedRows And iRow < .Rows Then
            If iCol >= .FixedCols And iCol < .Cols Then
                strName = fg.TextMatrix(iRow, iCol)
                If m.eFormMode = eMode_Templates Then
                    'strip off hot key value (0: thru 9:) if applicable
                    strHotKey = Left(strName, 2)
                    Select Case strHotKey
                        Case "0:", "1:", "2:", "3:", "4:", "5:", "6:", "7:", "8:", "9:"
                            strName = Right(strName, Len(strName) - 3)
                    End Select
                End If
            End If
        End If
    End With
    
    GetName = strName

End Function

Private Function BestWidth() As Long

    Dim i&, j&
    Dim iLongest&, hHandle&
    Dim iMaxWidth&, iWidth&, iTwipsPerChar&

On Error Resume Next

    iTwipsPerChar = Me.TextWidth("b") - 10      'approximate: subtract 10 to account of narrow letters like l,i etc.
    iMaxWidth = iTwipsPerChar * 50
    
    If m.eFormMode = eMode_Templates Then
        iLongest = 30           '<Copy settings to other charts>
    ElseIf m.eFormMode = eMode_Pages Then
        iLongest = 23           '<Create new chart page>
    ElseIf m.eFormMode = eMode_SecSubCom Then
        iLongest = 15
        If m.strBtnID <> "ID_Components" Then iTwipsPerChar = Me.TextWidth("B")
    Else
        Exit Function           'theoretically should never get here
    End If
    
    
    If Not m.aItems Is Nothing Then

        hHandle = m.aItems.ArrayHandle
        
        For i = 0 To m.aItems.Size - 1
            j = Len(Parse(gdGetStr(hHandle, i), vbTab, 1))
            If j > iLongest Then iLongest = j
        Next
        
        iWidth = iLongest * iTwipsPerChar
        If iWidth > iMaxWidth Then iWidth = iMaxWidth
    
        BestWidth = iWidth
        
    End If
        
End Function

