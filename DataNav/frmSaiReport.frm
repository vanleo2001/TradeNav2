VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmSaiReport 
   Caption         =   "SAI Report"
   ClientHeight    =   12495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12330
   LinkTopic       =   "Form1"
   ScaleHeight     =   12495
   ScaleWidth      =   12330
   Begin HexUniControls.ctlUniFrameWL fraDate 
      Height          =   1575
      Left            =   660
      TabIndex        =   1
      Top             =   540
      Visible         =   0   'False
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
      Caption         =   "frmSaiReport.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmSaiReport.frx":004C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmSaiReport.frx":006C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdRunForDate 
         Height          =   375
         Left            =   420
         TabIndex        =   3
         Top             =   960
         Width           =   2175
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
         Caption         =   "frmSaiReport.frx":0088
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmSaiReport.frx":00D4
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmSaiReport.frx":00F4
         RightToLeft     =   0   'False
      End
      Begin gdOCX.gdSelectDate gdReportDate 
         Height          =   315
         Left            =   420
         TabIndex        =   2
         Top             =   420
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         AllowWeekends   =   0   'False
      End
   End
   Begin ActiveToolBars.SSActiveToolBars tbToolbar 
      Left            =   900
      Top             =   3060
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131083
      ToolBarsCount   =   1
      ToolsCount      =   8
      Tools           =   "frmSaiReport.frx":0110
      ToolBars        =   "frmSaiReport.frx":3541
   End
   Begin VSFlex7LCtl.VSFlexGrid fg 
      Height          =   2355
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4515
      _cx             =   7964
      _cy             =   4154
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
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
Attribute VB_Name = "frmSaiReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const kMaxSymbols As Long = 100
Private Const kExtraSpace As Long = -60

Private Type mPrivate
    bForexAllowed As Boolean
    bFuturesAllowed As Boolean
    bStocksAllowed As Boolean
    RowIDs As cGdArray
    dFontSize As Double
    LogoImage As IPictureDisp
    
    nSessionDate As Long
    SymbolIDs As cGdArray ' ID's flagged as negative are not shown in the report
    DefaultSymbolIDs As cGdArray
End Type
Private m As mPrivate

Private Sub cmdRunForDate_Click()
On Error GoTo ErrSection

    If m.nSessionDate <> gdReportDate.Value Or fg.Cols = fg.FixedCols Then
        m.nSessionDate = gdReportDate.Value
        LoadGrid
    End If
    
    fraDate.Visible = False
    fg.Visible = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSaiReport.cmdRunForDate_Click"
End Sub

Private Sub fg_AfterMoveColumn(ByVal Col As Long, Position As Long)
On Error GoTo ErrSection

    Dim i&, j&, nMovedID&, nPriorID&, iMovedTo&

    ' symbol was moved
    nMovedID = GetSymbolID(fg.TextMatrix(0, Position))
    If nMovedID <> 0 Then
        If Position > fg.FixedCols Then
            nPriorID = GetSymbolID(fg.TextMatrix(0, Position - 1))
        End If
        ' find moved symbol
        For i = 0 To m.SymbolIDs.Size - 1
            If Abs(m.SymbolIDs(i)) = nMovedID Then
                m.SymbolIDs.Remove i
                ' find where to move to
                iMovedTo = 0 ' (default = move to the beginning)
                If nPriorID <> 0 Then
                    For j = 0 To m.SymbolIDs.Size - 1
                        If Abs(m.SymbolIDs(j)) = nPriorID Then
                            iMovedTo = j + 1
                            Exit For
                        End If
                    Next
                End If
                m.SymbolIDs.Add nMovedID, iMovedTo
                'frmTest.AddList GetSymbol(nMovedID) & vbTab & Str(iMovedTo)
                Exit For
            End If
        Next
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSaiReport.fg_AfterMoveColumn"
End Sub


Private Sub Form_Load()
On Error GoTo ErrSection

    Dim strText$
    
    g.Styler.StyleForm Me

    strText = GetIniFileProperty("SAI_Report", "", "Placement", g.strIniFile)
    If Len(strText) > 0 Then
        SetFormPlacement Me, strText, "LTHW"
    Else
        CenterTheForm Me
    End If

    Me.Icon = Picture16(ToolbarIcon("kSAI"), , True)
    With tbToolbar
        .Tools("ID_Symbols").Picture = Picture16("kChangeSymbol")
        .Tools("ID_SelectDate").Picture = Picture16("kCalendar")
        .Tools("ID_Print").Picture = Picture16(ToolbarIcon("ID_Print"))
        '.Tools("ID_ZoomIn").Picture = Picture16("kTextIncrease")
        '.Tools("ID_ZoomOut").Picture = Picture16("kTextDecrease")
        .Tools("ID_Close").Picture = Picture16("kCancel")
    
        ' 8/25/2014 per Gary: disable printing
        .Tools("ID_Print").Visible = False
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSaiReport.Form_Load"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    On Error Resume Next
    Dim s$
    s = m.SymbolIDs.JoinFields(",")
    SetIniFileProperty "SymbolIDs", s, "SAI", g.strIniFile
    SetIniFileProperty "FontSize", m.dFontSize, "SAI", g.strIniFile
    SetIniFileProperty "SAI_Report", GetFormPlacement(Me), "Placement", g.strIniFile

End Sub

Private Sub Form_Resize()

    On Error Resume Next
    
    With fg
        .Move .Left, .Top, Me.ScaleWidth - .Left * 2, Me.ScaleHeight - .Top - .Left
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error Resume Next
    
    Set m.RowIDs = Nothing
    Set m.SymbolIDs = Nothing
    Set m.DefaultSymbolIDs = Nothing

End Sub

Public Sub ShowMe()
On Error GoTo ErrSection

    ' see what's allowed based on enablements
    m.bForexAllowed = HasModule("SAI_X")
    m.bFuturesAllowed = HasModule("SAI_F")
    m.bStocksAllowed = HasModule("SAI_S")
    If Not m.bForexAllowed And Not m.bFuturesAllowed And Not m.bStocksAllowed Then Exit Sub

    m.dFontSize = GetIniFileProperty("FontSize", 0, "SAI", g.strIniFile)
    If m.dFontSize < 3 Then m.dFontSize = 9
    
    ' get default date
    m.nSessionDate = 0
    ShowDate
    fraDate.Visible = False
    
    ' show form and load grid using default date
    GetSymbols
    InitGrid
    ShowForm Me, eForm_Nonmodal, frmMain
    LoadGrid

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSaiReport.ShowMe"
End Sub

Private Sub ShowDate()
On Error GoTo ErrSection
    
    Dim nMaxDate&, dGMT#

    ' by default, set MaxDate to next business day after last daily download
    nMaxDate = LastDailyDownload + 1
    Do While Not IsWeekday(nMaxDate)
        nMaxDate = nMaxDate + 1
    Loop
    If m.nSessionDate = 0 Then
        m.nSessionDate = nMaxDate
    End If
    
    ' but if after 10pm GMT on that date, then allow going forward one more business day
    dGMT = ConvertTimeZone(Now, "", "GMT")
    If dGMT > nMaxDate + 22# / 24# Then
        nMaxDate = nMaxDate + 1
        Do While Not IsWeekday(nMaxDate)
            nMaxDate = nMaxDate + 1
        Loop
    End If
    gdReportDate.MaxDate = nMaxDate
    If m.nSessionDate <= nMaxDate Then
        gdReportDate.Value = m.nSessionDate
    Else
        gdReportDate.Value = nMaxDate
    End If
    fraDate.Visible = True
    fg.Visible = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSaiReport.ShowDate"
End Sub

' returns True if this symbol is allowed based on the SAI enablements
Public Function SymbolAllowed(ByVal vSymbol As Variant) As Boolean
On Error GoTo ErrSection

    Dim strSymbol$
    
    strSymbol = GetSymbol(vSymbol)
    If Len(strSymbol) > 0 Then
        Select Case SecurityType(strSymbol)
        Case "F"
            SymbolAllowed = m.bFuturesAllowed
        Case "S"
            SymbolAllowed = m.bStocksAllowed
        Case Else
            If IsForex(strSymbol) Then
                SymbolAllowed = m.bForexAllowed
            ElseIf m.bForexAllowed Or m.bFuturesAllowed Or m.bStocksAllowed Then
                SymbolAllowed = True ' other indices are allowed if any other type is allowed
            End If
        End Select
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmSaiReport.SymbolAllowed"
End Function

Private Sub GetSymbols()
On Error GoTo ErrSection

    Dim i&, nID&, s$
    Dim aSymbols As New cGdArray

    ' setup list of default symbols (based on allowed security types)
    Set m.DefaultSymbolIDs = New cGdArray
    m.DefaultSymbolIDs.Create eGDARRAY_Longs, 0, 0
    s = "ES-067,ZB-067,GC3-067,SI3-067,PL3-067,HG3-067,CL3-067,$AUD-USD,$EUR-JPY,$USD-CAD,$CHF-JPY,$EUR-USD,$GBP-JPY,$GBP-USD,$NZD-USD,$USD-JPY,$USD-CHF,$AUD-JPY,$EUR-GBP,$GBP-CHF,$GBP-CAD,$EUR-AUD,$CAD-JPY,$GBP-AUD,$AUD-CAD,MSFT,JPM"
    aSymbols.SplitFields s, ","
    For i = 0 To aSymbols.Size - 1
        s = aSymbols(i)
        If SymbolAllowed(s) Then
            nID = GetSymbolID(s)
            If nID > 0 Then
                m.DefaultSymbolIDs.Add nID
            End If
        End If
    Next
    
    ' get user's list of symbols
    Set m.SymbolIDs = New cGdArray
    m.SymbolIDs.Create eGDARRAY_Longs, 0, 0
    s = GetIniFileProperty("SymbolIDs", "", "SAI", g.strIniFile)
    aSymbols.SplitFields s, ","
    For i = 0 To aSymbols.Size - 1
        nID = Val(aSymbols(i))
        If SymbolAllowed(Abs(nID)) Then
            m.SymbolIDs.Add nID
        End If
    Next
    
    ' if empty list, set to defaults
    If m.SymbolIDs.Size = 0 Then
        Set m.SymbolIDs = m.DefaultSymbolIDs.MakeCopy
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSaiReport.GetSymbols"
End Sub

' to format list of symbols for Selection dialog
Private Function FormattedSymbolList(SymbolIDs As cGdArray) As cGdArray
On Error GoTo ErrSection

    Dim i&, nID&, strSymbol$, strDesc$
    Dim bActive As Boolean
    Dim aSymbols As New cGdArray
    
    aSymbols.Create eGDARRAY_Strings
    For i = 0 To SymbolIDs.Size - 1
        nID = SymbolIDs(i)
        bActive = (nID > 0) ' ID flagged as negative means to not show in the report
        nID = Abs(nID)
        strSymbol = GetSymbol(nID)
        If SymbolAllowed(strSymbol) Then
            strDesc = g.SymbolPool.Desc(g.SymbolPool.PoolRecForSymbolID(nID))
            aSymbols.Add "1" & vbTab & strSymbol & vbTab & " " & strDesc & vbTab & Str(Abs(bActive))
        End If
    Next
    
    Set FormattedSymbolList = aSymbols

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmSaiReport.FormattedSymbolList"
End Function

' to allow user to select and arrange the symbols (columns)
Private Sub ManageSymbols()
On Error GoTo ErrSection

    Dim i&, s$, nID&, strSymbol$, nActive&
    Dim aSymbols As cGdArray, aDefaultSymbols As cGdArray
    
    ' build the string arrays to hand off to dialog
    Set aSymbols = FormattedSymbolList(m.SymbolIDs)
    Set aDefaultSymbols = FormattedSymbolList(m.DefaultSymbolIDs)
       
    ' portfolio dialog
    If frmQuoteBoardFields.ShowMe(aSymbols, eQbfMode_SaiReport, aDefaultSymbols) Then
        nActive = 0
        m.SymbolIDs.Size = 0
        For i = 0 To aSymbols.Size - 1
            nID = GetSymbolID(Parse(aSymbols(i), vbTab, 2))
            If nID > 0 And nActive < kMaxSymbols Then
                If Val(Parse(aSymbols(i), vbTab, 4)) = 0 Or nActive >= kMaxSymbols Then
                    m.SymbolIDs.Add -nID ' "inactive" = not shown in report
                Else
                    m.SymbolIDs.Add nID ' "active" = show in report
                    nActive = nActive + 1
                End If
            End If
        Next
        
        LoadGrid
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSaiReport.ManageSymbols"
End Sub

Private Sub tbToolbar_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
On Error GoTo ErrSection:
    
    Dim s$
    
    Select Case Tool.ID
    Case "ID_Symbols"
        ManageSymbols
        
    Case "ID_SelectDate"
        If 0 Then
            s = InfBox("Display the SAI Report for:", "?", , "SAI Report Date", , , , , , "d", DateFormat(m.nSessionDate))
            If Len(s) > 0 Then
                If DateOf(s) <> m.nSessionDate Then
                    m.nSessionDate = DateOf(s)
                    LoadGrid
                End If
            End If
        Else
            ShowDate
        End If
            
    Case "ID_Print"
        fraDate.Visible = False
        fg.Visible = True
        PrintMe
        
    Case "ID_ZoomIn"
        ChangeFontSize 1
    
    Case "ID_ZoomOut"
        ChangeFontSize -1
        
    Case "ID_Disclaimer"
        s = "@" & App.Path & "\Info\SAI_Disclaimer.rtf"
        frmMessage.ShowMe "Strategic Analysis Indicator", s ', eModalMessage
                    
    Case "ID_Close"
        Unload Me
        
    Case "ID_UserGuide"
        s = App.Path & "\Info\SAI-manual.pdf"
        If FileExist(s) Then
            RunProcess s
        Else
            InfBox "User Guide not found", "e", , "Strategic Analysis Indicator"
        End If
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSaiReport.tbToolbar_ToolClick"
    Resume ErrExit
End Sub

Public Sub GridTextIncrease()
    ChangeFontSize 1
End Sub

Public Sub GridTextDecrease()
    ChangeFontSize -1
End Sub

Private Sub ChangeFontSize(ByVal dUpDown#)

    Dim iRow&, iCol&, dSize#, dRowHeight#

    If Not IsIDE Then
        On Error Resume Next
    End If
    With fg
        If .Visible Then
            m.dFontSize = m.dFontSize + dUpDown
            If m.dFontSize < 4 Then m.dFontSize = 4
            .Font.Size = m.dFontSize
            dRowHeight = .RowHeight(.Rows - 1) ' height of typical row
    
            ' since some rows have custom font-stuff (e.g. bolding),
            ' we have to reset font size of each cell
            For iRow = 0 To .Rows - 1
                Select Case UCase(Parse(.TextMatrix(iRow, 0), " ", 1))
                Case "DAILY", "MONTHLY", "WEEKLY"
                    dSize = m.dFontSize + 1
                    .RowHeight(iRow) = dRowHeight * 1.25
                Case Else
                    dSize = m.dFontSize
                End Select
                .Cell(flexcpFontSize, iRow, 0, iRow, 0) = dSize
                .Cell(flexcpFontSize, iRow, 1, iRow, .Cols - 1) = m.dFontSize
            Next
            .AutoSize 0, .Cols - 1, , kExtraSpace
        End If
    End With

End Sub


' returns the bar# for the last completed data bar prior to the session date
Private Function GetBarNumberAndFixClose(Bars As cGdBars) As Long
On Error GoTo ErrSection
    
    Dim nBar&, nDate&, nBarEndDate&, strSymbol$, dTime#, i&, dClose#, dStartTC#
    Dim MinuteBars As New cGdBars
    
    If Bars.Size = 0 Or m.nSessionDate <= 0 Then
        nBar = -1
    Else
        ' Get bar# of data completed prior to this session date
        nBar = Bars.FindDateTime(m.nSessionDate) - 1
        Do While nBar >= 0
            If Bars(eBARS_Close, nBar) <> kNullData Then
                Exit Do
            End If
            nBar = nBar - 1
        Loop
        
' TLB: was testing this per Gary, but probably won't need it now?
' (hopefully not, since this would make the report inconsistent with the SAI chart indicators)
#If False Then
        ' for Forex: fix the closing price (set to price at 5pm NY yearround)
        strSymbol = Bars.Prop(eBARS_Symbol)
        If IsForex(strSymbol) Then
            nBarEndDate = Bars(eBARS_DateTime, nBar)
            If nBarEndDate > 0 Then
                dStartTC = gdTickCount
                ' go back until we have a day with data (e.g. if starting from monthly bars)
                For nDate = nBarEndDate To nBarEndDate - 6 Step -1
                    If IsWeekday(nDate) Then
                        DM_GetBars MinuteBars, strSymbol, "60 minute", nDate, nDate
                        If MinuteBars.Size > 0 Then
                            ' now go back until we get the hourly bar at 5pm NY
                            For i = MinuteBars.Size - 1 To 0 Step -1
                                ' convert time to NY and round to the nearest hour
                                dTime = MinuteBars.DateTimeConvert(i, "NY")
                                If Hour(dTime + 0.5 / 24) <= 17 Then
                                    dClose = MinuteBars(eBARS_Close, i)
                                    If dClose > 0 Then
                                        ' replace the Close of the Bars with this price
                                        Bars(eBARS_Close, nBar) = dClose
                                    End If
                                    Exit For
                                End If
                            Next
                            Exit For
                        End If
                    End If
                Next
                If IsIDE Then
                    'frmTest.AddList strSymbol & vbTab & Format(gdTickCount - dStartTC, "####0")
                End If
            End If
        End If
#End If
    End If
    
    GetBarNumberAndFixClose = nBar
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmSaiReport.GetBarNumberAndFixClose"
End Function

' Pass in 0 for BS/SS, or 1-4 for Profit Level
Private Function CalcPL(Bars As cGdBars, ByVal nBar&, ByVal iProfitLevel&, ByVal bSell As Boolean) As Double
On Error GoTo ErrSection
    
    Dim dMult#, dRange#, dClose#
    
    dClose = Bars(eBARS_Close, nBar)
    dRange = Bars(eBARS_High, nBar) - Bars(eBARS_Low, nBar)
    
    If iProfitLevel < 0 Or iProfitLevel > 4 Or dClose = kNullData Or dRange < 0 Then
        CalcPL = 0
    Else
        ' Get the multiplier for this profit level
        'BS/SS = 0.073, PL1 = 0.309, PL2 = 0.545, PL3 = 0.691, PL4 = 0.927
        Select Case iProfitLevel
        Case 0
            dMult = 0.073
        Case 1
            dMult = 0.309
        Case 2
            dMult = 0.545
        Case 3
            dMult = 0.691
        Case 4
            dMult = 0.927
        Case Else
            dMult = 0 ' undefined
        End Select
        If bSell Then
            dMult = -dMult
        End If
        
        CalcPL = dClose + dRange * dMult
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmSaiReport.CalcPL"
End Function

Private Sub InitGrid()
On Error GoTo ErrSection

    Dim s$, strEach$, iRow&, iType&, dRowHeight#
    
    ' init the Row ID's
    'strEach = "RISK ; BUY PL4 ; BUY PL3 ; BUY PL2 ; BUY PL1 ; BUY STOP ; SELL STOP ; SELL PL1 ; SELL PL2; SELL PL3 ; SELL PL4 ; PPC ;"
    strEach = " BUY PL4 ; BUY PL3 ; BUY PL2 ; BUY PL1 ; BUY STOP ; RISK ; SELL STOP ; SELL PL1 ; SELL PL2; SELL PL3 ; SELL PL4 ; PPC ;"
    s = " ; DAILY ;" & strEach & " MTR ; IND ; WEEKLY ;" & strEach & " OTE (weekly) ; MONTHLY ;" & strEach & " OTE (monthly) ;"
    Set m.RowIDs = New cGdArray
    m.RowIDs.SplitFields s, ";"

    With fg
        SetupGrid fg, eGridMode_Grid
        .Font.Size = m.dFontSize
        .SelectionMode = flexSelectionFree
        .AllowSelection = True
        .ExtendLastCol = False
        .ExplorerBar = flexExMove
        '.BackColorFixed = RGB(255, 244, 216)
        '.BackColorFixed = RGB(224, 224, 224)
        .BackColorFixed = &HDDFAF9
        
        
        .Cols = 1
        .FixedCols = 1
        .Rows = m.RowIDs.Size
        dRowHeight = .RowHeight(.Rows - 1) ' height of typical row (e.g. last row)
        For iRow = 0 To m.RowIDs.Size - 1
            s = Trim(m.RowIDs(iRow))
            m.RowIDs(iRow) = s
            
            Select Case UCase(Parse(s, " ", 1))
            Case "DAILY", "MONTHLY", "WEEKLY"
                .Cell(flexcpForeColor, iRow, 0) = vbWhite
                .Cell(flexcpBackColor, iRow, 0) = RGB(1, 1, 1) ' so will be black
                .Cell(flexcpFontSize, iRow, 0) = m.dFontSize + 1
                .RowHeight(iRow) = dRowHeight * 1.25
            Case "BUY"
                .Cell(flexcpForeColor, iRow, 0) = vbBlue
            Case "SELL"
                .Cell(flexcpForeColor, iRow, 0) = vbRed
            Case "OTE"
                s = "OTE"
                .Cell(flexcpBackColor, iRow, 0) = RGB(144, 192, 240)
            End Select
            
            Select Case UCase(Parse(s, " ", 2))
            Case "PL1"
                s = "Profit Level 1"
            Case "PL2"
                s = "Profit Level 2"
            Case "PL3"
                s = "Profit Level 3"
            Case "PL4"
                s = "Profit Level 4"
            End Select
                        
            .TextMatrix(iRow, 0) = s
        Next
        
        .ColAlignment(0) = flexAlignCenterCenter
        .Cell(flexcpFontBold, 0, 0, .Rows - 1, 0) = True
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = True
        
        .AutoSize 0, 0
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSaiReport.InitGrid"
End Sub

Private Sub LoadGrid()
On Error GoTo ErrSection

    Dim i&, iSymbol&, iCol&, iRow&, iLevel&, nBar&, dValue#, dMTR#, iCount&, dDailyClose#, iIndDir&, dBS#, dSS#
    Dim s$, strSymbol$
    Dim bBold As Boolean, bSectionHdr As Boolean
    Dim Bars As cGdBars, Daily As New cGdBars, Weekly As New cGdBars, Monthly As New cGdBars
       
    'm.nSessionDate = gdForDate.Value
    If m.nSessionDate <= 0 Or m.nSessionDate > LastDailyDownload + 4 Then
        m.nSessionDate = LastDailyDownload + 1
    End If
    Do While Not IsWeekday(m.nSessionDate)
        m.nSessionDate = m.nSessionDate + 1
    Loop
    
    Me.Caption = "Strategic Analysis Indicator for " & DateFormat(m.nSessionDate)
    
    fraDate.Visible = False
    fg.Visible = True
    fg.Redraw = flexRDBuffered
    
    iCol = fg.FixedCols
    For iSymbol = 0 To m.SymbolIDs.Size - 1
        ' ID flagged as negative means don't show symbol in report
        If m.SymbolIDs(iSymbol) > 0 Then
            strSymbol = GetSymbol(m.SymbolIDs(iSymbol))
            If Not SymbolAllowed(strSymbol) Then
                strSymbol = ""
            End If
        Else
            strSymbol = ""
        End If
        If Len(strSymbol) > 0 Then
            ' load daily/weekly/monthly bars (but load enough to get a 55-bar moving average)
            Set Daily = New cGdBars
            Set Weekly = New cGdBars
            Set Monthly = New cGdBars
            DM_GetBars Daily, strSymbol, 0, m.nSessionDate - 88, m.nSessionDate + 6
            If g.RealTime.Active And m.nSessionDate > LastDailyDownload Then
                g.RealTime.SpliceBars Daily
            End If
            Daily.AddForecastBars 1
            Weekly.BuildBars "Weekly", Daily.BarsHandle
            Monthly.BuildBars "Monthly", Daily.BarsHandle
            
            ' calc the MTR from a 55-bar MA
            Set Bars = Daily
            nBar = GetBarNumberAndFixClose(Bars)
            iIndDir = 0
            iCount = 0
            dMTR = 0
            For i = nBar To 0 Step -1
                dValue = Bars(eBARS_Close, i)
                If dValue <> kNullData Then
                    dMTR = dMTR + dValue
                    iCount = iCount + 1
                    If iCount >= 55 Then
                        Exit For
                    End If
                End If
            Next
            dDailyClose = Bars(eBARS_Close, nBar)
            If iCount >= 55 Then
                dMTR = dMTR / iCount
                If dMTR > dDailyClose Then
                    iIndDir = 1 ' IND = Buy
                ElseIf dMTR < dDailyClose Then
                    iIndDir = -1 ' IND = Sell
                End If
            Else
                dMTR = kNullData
            End If
            
            With fg
                .Cols = iCol + 1
                .TextMatrix(0, iCol) = strSymbol
                .ColAlignment(iCol) = flexAlignCenterCenter
                
                For iRow = .FixedRows To m.RowIDs.Size - 1
                    bBold = False
                    bSectionHdr = False
                    dValue = kNullData
                    .Cell(flexcpForeColor, iRow, iCol) = .Cell(flexcpForeColor, iRow, 0) ' default
                    
                    s = m.RowIDs(iRow)
                    iLevel = Val(Right(s, 1))
                    Select Case UCase(Parse(s, " ", 1))
                    Case "DAILY"
                        bSectionHdr = True
                        dValue = Bars(eBARS_DateTime, nBar + 1)
                        .TextMatrix(iRow, iCol) = DateFormat(dValue)
                    Case "WEEKLY"
                        bSectionHdr = True
                        Set Bars = Weekly
                        nBar = GetBarNumberAndFixClose(Bars)
                        dValue = Bars(eBARS_DateTime, nBar + 1)
                        ' backup to the previous Monday
                        For i = Int(dValue) To 1 Step -1
                            If Not IsWeekday(i - 1) Then
                                .TextMatrix(iRow, iCol) = DateFormat(i, M_D) & "-" & DateFormat(dValue, M_D)
                                Exit For
                            End If
                        Next
                    Case "MONTHLY"
                        bSectionHdr = True
                        Set Bars = Monthly
                        nBar = GetBarNumberAndFixClose(Bars)
                        dValue = Bars(eBARS_DateTime, nBar + 1)
                        .TextMatrix(iRow, iCol) = DateFormat(dValue, MMM_YY)
                        
                    Case "BUY"
                        dValue = CalcPL(Bars, nBar, iLevel, False)
                        If iLevel = 0 Then bBold = True
                    Case "SELL"
                        dValue = CalcPL(Bars, nBar, iLevel, True)
                        If iLevel = 0 Then bBold = True
                    Case "RISK"
                        dValue = CalcPL(Bars, nBar, iLevel, False) - CalcPL(Bars, nBar, iLevel, True)
                        bBold = True
                    Case "PPC"
                        dValue = Bars(eBARS_Close, nBar)
                    Case "MTR"
                        dValue = dMTR
                    Case "IND"
                        If iIndDir > 0 Then
                            .TextMatrix(iRow, iCol) = "Buy"
                            .Cell(flexcpForeColor, iRow, iCol) = vbBlue
                        ElseIf iIndDir < 0 Then
                            .TextMatrix(iRow, iCol) = "Sell"
                            .Cell(flexcpForeColor, iRow, iCol) = vbRed
                        Else
                            .TextMatrix(iRow, iCol) = ""
                        End If
                    Case "OTE"
                        ' TLB: per our understanding of the formula given to us by Gary ...
                        dBS = CalcPL(Bars, nBar, 0, False) ' Monthly/Weekly Buy Stop
                        dSS = CalcPL(Bars, nBar, 0, True) ' Monthly/Weekly Sell Stop
                        If dDailyClose > dBS Then
                            dValue = dDailyClose - dBS
                            bBold = True
                        ElseIf dDailyClose < dSS Then
                            dValue = dSS - dDailyClose
                            bBold = True
                        Else
                            ' undefined when between BS and SS?
                            dValue = kNullData
                            .TextMatrix(iRow, iCol) = ""
                        End If
                        .Cell(flexcpBackColor, iRow, iCol) = .Cell(flexcpBackColor, iRow, 0)
                    End Select
                    
                    If bSectionHdr Then
                        ' new timeframe section
                        .Cell(flexcpBackColor, iRow, iCol) = .Cell(flexcpBackColor, iRow, 0)
                    ElseIf dValue <> kNullData Then
                        ' display as price
                        .TextMatrix(iRow, iCol) = Bars.PriceDisplay(dValue)
                    End If
                    .Cell(flexcpFontBold, iRow, iCol) = bBold
                Next
            End With
            iCol = iCol + 1
        End If
    Next
    
    With fg
        .Cols = iCol
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = True
        .AutoSize 0, .Cols - 1, , kExtraSpace
        If .Cols > .FixedCols Then
            .ShowCell .FixedRows, .FixedCols
        End If
        .Select 0, 0
    End With
    
    Set Bars = Nothing
    Set Daily = Nothing
    Set Weekly = Nothing
    Set Monthly = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSaiReport.LoadGrid"
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PrintMe
'' Description: Send the grid to the Print Preview
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function PrintMe() As Boolean
On Error GoTo ErrSection

    If m.dFontSize < 7 Then
        PrintMe = frmPrintPreview.ShowMe("CNV SaiReport", frmSaiReport, , 1.6, 0.5, 0.4, 0.3, True, , , True)
    Else
        PrintMe = frmPrintPreview.ShowMe("CNV SaiReport", frmSaiReport, , 1.6, 0.5, 0.4, 0.3, False, , , True)
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmSaiReport.PrintMe"
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenerateReport
'' Description: Callback function for the Print Preview
'' Inputs:      Arguments into the Print Preview
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GenerateReport(ByVal vArgs As Variant)
On Error GoTo ErrSection:

    Dim lRow As Long
    Dim lCol As Long
    Dim lFirstCol&, lLastCol&
    Dim strText As String
    
    ' see if user has selected multiple columns to print (i.e. if not all columns)
    With fg
        If .ColSel > .Col Then
            lFirstCol = .Col
            lLastCol = .ColSel
        ElseIf .ColSel < .Col Then
            lFirstCol = .ColSel
            lLastCol = .Col
        Else ' print all columns
            lFirstCol = 0
            lLastCol = 0
        End If
        If lLastCol > lFirstCol Then
            ' hide the non-selected columns
            For lCol = .FixedCols To .Cols - 1
                If lCol >= lFirstCol And lCol <= lLastCol Then
                    .ColHidden(lCol) = False
                Else
                    .ColHidden(lCol) = True
                End If
            Next
        End If
    End With
    
    strText = App.Path & "\Info\SAI.jpg"
    If FileExist(strText) Then
        Set m.LogoImage = LoadPicture(strText)
    Else
        Set m.LogoImage = Nothing
    End If
    
    With frmPrintPreview.vp
        .Clear
        .StartDoc
        If 0 Then
            DoPrintHeader 8
        Else
            .LineSpacing = 100
            .HdrFontName = fg.Font.Name ' "Times New Roman"
            .HdrFontSize = 10
            strText = "|Trade Navigator" & vbCrLf & "Genesis Financial Technologies - "
            .Header = " "
            '.Header = strText & GetProvidedProperty("Website", , True)
            .Footer = "  Powered by Genesis Financial Technologies - TradeNavigator.com||Page: %d    "
        End If
        
    If 0 Then
        .TextAlign = taCenterMiddle
        .Font.Name = fg.Font.Name ' "Times New Roman"
        .Font.Size = 14
        .Font.Bold = True
        '.FontUnderline = True
        .Text = "Strategic Analysis Indicator Report for " & DateFormat(m.nSessionDate) & vbLf
        .Font.Size = 12
        .FontUnderline = False
        .Font.Bold = False
        .TextAlign = taLeftMiddle
        .Text = vbLf
    End If
              
        
        'fg.ExtendLastCol = False
        If frmPrintPreview.GoingToFile Then
            With fg
                For lRow = 0 To .Rows - 1
                    strText = ""
                    For lCol = 0 To .Cols - 1
                        If Not .ColHidden(lCol) Then
                            strText = strText & .Cell(flexcpTextDisplay, lRow, lCol) & vbTab
                        End If
                    Next lCol
                    strText = Left(strText, Len(strText) - 1) ' strip the trailing tab
                    frmPrintPreview.vp.Text = strText & vbCrLf
                Next lRow
            End With
        Else
            .RenderControl = fg.hWnd
        End If
        'fg.ExtendLastCol = True
        
        strText = FileToString(App.Path & "\Info\SAI_Disclaimer.rtf")
        If Len(strText) > 10 Then
            .NewPage
            .Text = vbLf
            .TextRTF = strText
        End If
        
        .EndDoc
    End With

    If lLastCol > lFirstCol Then
        ' show all columns again
        For lCol = fg.FixedCols To fg.Cols - 1
            fg.ColHidden(lCol) = False
        Next
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSaiReport.GenerateReport", eGDRaiseError_Raise
End Sub

Public Sub AfterHeaderEvent(ByVal vArgs As Variant)
On Error GoTo ErrSection:

    Dim strText As String
    
    With frmPrintPreview.vp
        If Not frmPrintPreview.GoingToFile Then
            If .CurrentPage = 1 Then
                strText = strText
            End If
            
            .CurrentY = "0.5in"
            .TextAlign = taLeftMiddle
            .Font.Name = fg.Font.Name ' "Times New Roman"
                        
            .Font.Bold = False
            .Font.Size = 8
            .Text = "Copyright by Strategic Analysis -- Available by Subscription Only -- NOT For Distribution" & vbLf '& vbLf & vbLf
            
            .Font.Size = 14
            .Font.Bold = True
            .Text = vbLf & "Strategic Analysis Indicator for " & DateFormat(m.nSessionDate) & vbLf
            .Font.Size = 12
            .Font.Bold = False
            .Text = "With 4 pre-defined profit levels!" & vbLf
            '.Font.Size = 8
            '.Text = vbLf & "Copyright by Strategic Analysis (www.strategic-analysis.biz) -- NOT for Distribution" & vbLf
            
            If Not m.LogoImage Is Nothing Then
                .DrawPicture m.LogoImage, "6in", "0.4in", "1in", "1in", vppaZoom ' vppaRightTop
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSaiReport.AfterHeaderEvent", eGDRaiseError_Raise
End Sub


