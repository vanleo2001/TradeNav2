VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmRollsTable 
   Caption         =   "Continuous Contract Rolls"
   ClientHeight    =   5220
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   8070
   Icon            =   "frmRollsTable.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   8070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VSFlex7LCtl.VSFlexGrid fgData 
      Height          =   3615
      Left            =   180
      TabIndex        =   0
      Top             =   780
      Width           =   4455
      _cx             =   7858
      _cy             =   6376
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
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   435
      Left            =   60
      TabIndex        =   1
      Top             =   0
      Width           =   7275
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
      Caption         =   "frmRollsTable.frx":014A
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmRollsTable.frx":017E
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmRollsTable.frx":019E
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniRadioXP opt65 
         Height          =   255
         Left            =   4200
         TabIndex        =   2
         Top             =   120
         Width           =   975
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
         Caption         =   "frmRollsTable.frx":01BA
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmRollsTable.frx":01E8
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmRollsTable.frx":0264
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP opt66 
         Height          =   255
         Left            =   3060
         TabIndex        =   5
         Top             =   120
         Width           =   1095
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
         Caption         =   "frmRollsTable.frx":0280
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmRollsTable.frx":02AE
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmRollsTable.frx":02F4
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP opt67 
         Height          =   255
         Left            =   1920
         TabIndex        =   4
         Top             =   120
         Width           =   1095
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
         Caption         =   "frmRollsTable.frx":0310
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   -1  'True
         Tip             =   "frmRollsTable.frx":033E
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmRollsTable.frx":0386
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniComboImageXP cboSym 
         Height          =   315
         Left            =   720
         TabIndex        =   3
         Top             =   60
         Width           =   975
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         ButtonBackColor =   -2147483633
         ButtonForeColor =   -2147483630
         ButtonStyle     =   -1
         SelectorStyle   =   -1
         SelBackColor    =   -2147483635
         SelForeColor    =   -2147483634
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
         Tip             =   "frmRollsTable.frx":03A2
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmRollsTable.frx":03C2
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblRolling 
         Height          =   315
         Left            =   5400
         Top             =   120
         Width           =   1695
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
         Caption         =   "frmRollsTable.frx":03DE
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmRollsTable.frx":042C
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmRollsTable.frx":044C
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label1 
         Height          =   195
         Left            =   60
         Top             =   120
         Width           =   795
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
         Caption         =   "frmRollsTable.frx":0468
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmRollsTable.frx":0496
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmRollsTable.frx":04B6
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin VB.Menu mnuPref 
      Caption         =   "Preferences"
      Begin VB.Menu mnuSetChart 
         Caption         =   "&Set Active Chart"
      End
      Begin VB.Menu mnuChangeFont 
         Caption         =   "&Change Font"
      End
   End
End
Attribute VB_Name = "frmRollsTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum eGDCols
    eGDCol_Date = 0
    eGDCol_Symbol = 1
    eGDCol_From = 2
    eGDCol_To = 3
    eGDCol_Desc = 4
End Enum

Private Type mPrivate
    astrSymDesc As New cGdArray
    
    bOnlyIfNewRolls As Boolean
End Type
Private m As mPrivate

Private Function GDCol(ByVal Col As eGDCols) As Long
    GDCol = Col
End Function

Private Sub InitCboSym(ByVal strSymSave As String)
On Error GoTo ErrSection:

    Dim aSymbols As cGdArray            ' Array of true/false in grid values
    Dim astrSym As New cGdArray
    Dim strSymBase$, strSymbol$
    Dim i&, dTime#
    
dTime = gdTickCount
    
    cboSym.Clear
    cboSym.AddItem "ALL", 0
    ' Get the field number for the symbol group
    i = g.SymbolPool.FieldNumForID("GRP:ALL FUTURES.GRP")
    Set aSymbols = g.SymbolPool.ArrayTable.FieldArray(i)
    
    astrSym.Size = 0
    m.astrSymDesc.Size = 0
    For i = 0 To aSymbols.Size - 1
        If Abs(aSymbols(i)) = 1 Then
            strSymbol = g.SymbolPool.Symbol(i)
            strSymBase = Parse(strSymbol, "-", 2)
            If "065" = strSymBase Or "066" = strSymBase Or "067" = strSymBase Then
                astrSym.Add Parse(strSymbol, "-", 1)
                m.astrSymDesc.Add strSymbol & vbTab & g.SymbolPool.Desc(i)
            End If
        End If
    Next
    m.astrSymDesc.Sort
    
    astrSym.Sort eGdSort_DeleteDuplicates
    For i = 0 To astrSym.Size - 1
        cboSym.AddItem astrSym(i), i + 1
    Next
    
    cboSym.ListIndex = 0
    
    If strSymSave <> "" Then
        strSymBase = Parse(strSymSave, "-", 1)
        If astrSym.BinarySearch(strSymBase, i) = True Then
            cboSym.ListIndex = i + 1
        End If
        i = Val(Parse(strSymSave, "-", 2))
        If 65 = i Then
            opt65.Value = True
        ElseIf 66 = i Then
            opt66.Value = True
        ElseIf 67 = i Then
            opt67.Value = True
        End If
    End If
        
'dTime = gdTickCount - dTime
'StatusMsg "Init " & Str(dTime)

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmRollsTable.InitCboSym", eGDRaiseError_Raise

End Sub

Private Function GetSelSym() As String
On Error GoTo ErrSection:

    Dim strSym As String
    
    strSym = cboSym.Text
    If opt67.Value = True Then
        strSym = strSym & "-067"
    ElseIf opt66.Value = True Then
        strSym = strSym & "-066"
    ElseIf opt65.Value = True Then
        strSym = strSym & "-065"
    Else
        opt67.Value = True
        strSym = strSym & "-067"
    End If
    
    GetSelSym = strSym
    
ErrExit:
    Exit Function

ErrSection:
    RaiseError "frmRollsTable.GetSelSym", eGDRaiseError_Raise

End Function

Private Function GetSymDesc(ByVal strSym As String) As String
On Error GoTo ErrSection:

    Dim iPos&, s$
    
    m.astrSymDesc.BinarySearch strSym & vbTab, iPos
    s = m.astrSymDesc(iPos)
    If Left(s, Len(strSym) + 1) = strSym & vbTab Then
        GetSymDesc = Parse(s, vbTab, 2)
    End If
    
ErrExit:
    Exit Function

ErrSection:
    RaiseError "frmRollsTable.GetSymDesc", eGDRaiseError_Raise

End Function

Private Function SetGridData(ByVal strSym As String) As Boolean
On Error GoTo ErrSection:

    Dim Table As cGdTable
    Dim rc As Boolean
    Dim i&, j&, k&
    Dim nRedraw$, nSymCnt&
    Dim nRange&, nYear&, nDiff&, dDate#
    Dim strDesc$, strSeries$, strFrom$, strTo$
    Dim dCutoffDate#, dAdjustment#
    Dim dLastDaily As Double
    Dim lIndex As Long

    Screen.MousePointer = vbHourglass
    
    With fgData
        nRedraw = .Redraw
        .Redraw = flexRDNone
        .Rows = fgData.FixedRows
    End With
    
    nSymCnt = 1
    If Parse(strSym, "-", 1) = "ALL" Then
        strSeries = "-" & Parse(strSym, "-", 2)
        nSymCnt = cboSym.ListCount - 1
    End If
    
    'if "ALL" symbols are being displayed then go back only a year (was 6 months)
    If nSymCnt > 1 Then
        dCutoffDate = LastDailyDownload - 380 ' - 183
    End If
    
    BenchMark
    
    For k = 1 To nSymCnt
        If nSymCnt > 1 Then
            strSym = cboSym.List(k) & strSeries
        End If
        Set Table = GetRollsTable(strSym)
        strDesc = GetSymDesc(strSym)
        'rolls table are returned ascending date order
        'we want to display most recent date first
        strFrom = ""
        For i = Table.NumRecords - 1 To 1 Step -1
            ' get date
            dDate = Table(1, i)
            If dDate < dCutoffDate Then Exit For
            ' get symbol for "to" contract
            strTo = ""
            If Len(strFrom) > 0 Then
                strTo = strFrom '(from previous loop)
            Else
                j = Table(0, i)
                strTo = GetSymbol(j)
                If Len(strTo) > 0 Then
                    strTo = Format(Parse(strTo, "-", 2), "@@@@-@@")
                End If
            End If
            dAdjustment = Table(2, i)
            'get previous record to extract the "from" contract
            j = Table(0, i - 1)
            strFrom = GetSymbol(j)
            If Len(strFrom) > 0 Then
                strFrom = Format(Parse(strFrom, "-", 2), "@@@@-@@")
            End If
            'add to grid
            If Len(strFrom) > 0 And Len(strTo) > 0 Then
                fgData.AddItem dDate & vbTab & strSym & vbTab _
                    & strFrom & vbTab & strTo & vbTab & strDesc
            End If
        Next i
    Next k
    
    'BenchMark "Done"
    
    fgData.AutoSize GDCol(eGDCol_Date)
    fgData.Col = GDCol(eGDCol_Date)
    fgData.Sort = flexSortNumericDescending
    
    dLastDaily = LastDailyDownload
        
    SetGridData = False
    If fgData.Rows > fgData.FixedRows Then
        If DateOf(fgData.TextMatrix(fgData.FixedRows, GDCol(eGDCol_Date))) > dLastDaily Then
            For lIndex = fgData.FixedRows To fgData.Rows - 1
                If DateOf(fgData.TextMatrix(lIndex, GDCol(eGDCol_Date))) > dLastDaily Then
                    If g.nColorTheme = kDarkThemeColor Then
                        fgData.Cell(flexcpForeColor, lIndex, 0, lIndex, fgData.Cols - 1) = vbGreen
                    Else
                        fgData.Cell(flexcpForeColor, lIndex, 0, lIndex, fgData.Cols - 1) = vbBlue
                    End If
                Else
                    Exit For
                End If
            Next lIndex
            SetGridData = True
        End If
    End If
    fgData.Redraw = nRedraw
    
    Screen.MousePointer = 0
   
ErrExit:
    Exit Function

ErrSection:
    RaiseError "frmRollsTable.SetGridData", eGDRaiseError_Raise

End Function

Public Sub ShowMe(Optional ByVal bOnlyIfNewRolls As Boolean = False)
On Error GoTo ErrSection:

    Dim bNewRolls As Boolean

    If Not HasModule("F") Then
        cboSym.Clear
    Else
        m.bOnlyIfNewRolls = bOnlyIfNewRolls
        If bOnlyIfNewRolls Then
            cboSym.Text = "ALL"
            opt67.Value = True
        End If
        bNewRolls = SetGridData(GetSelSym())
        lblRolling.Visible = bNewRolls
        If bOnlyIfNewRolls And (Not bNewRolls) Then
            Unload Me
            GoTo ErrExit
        End If
        
        If g.nColorTheme = kDarkThemeColor Then lblRolling.ForeColor = vbGreen
        
        ShowForm Me, , , , ALT_GRID_ROW_COLOR
    End If
    
    If cboSym.ListCount <= 1 Then
        InfBox "No continuous contracts were found| (requires updating 'Futures' data).", "i", , "Continuous Contract Rolls"
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmRollsTable.ShowMe", eGDRaiseError_Raise

End Sub

Private Sub cboSym_Click()
On Error GoTo ErrSection:
    
    If Not Me.Visible Then Exit Sub
    
    'RH commented out cboSym.Refresh
    SetGridData GetSelSym()
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmRollsTable.cboSym.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fgData_DblClick()
On Error GoTo ErrSection:

    With fgData
        If .MouseRow >= .FixedRows And .MouseRow < .Rows Then
            SetActiveChartSymbol .TextMatrix(.MouseRow, GDCol(eGDCol_Symbol))
        End If
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmRollsTable.fgData.DblClick", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fgData_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    With fgData
        If .MouseRow > .FixedRows - 1 And .MouseRow < .Rows Then
            .Row = .MouseRow
            .RowSel = .Row
            
            If Button = vbRightButton Then
                mnuSetChart.Caption = "&Set Active Chart to " & .TextMatrix(.Row, GDCol(eGDCol_Symbol))
                PopupMenu mnuPref
            End If
        End If
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmRollsTable.fgData.MouseDown", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fgData_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    GridTooltip fgData
    
End Sub

Private Sub Form_Deactivate()
On Error GoTo ErrSection:

    SetPrevActiveForm Me

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmRollsTable.Form.Deactivate", eGDRaiseError_Show
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
    RaiseError "frmRollsTable.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub Form_Load()
On Error GoTo ErrSection:
  
    Dim strSaveSym$, strFont$
    
    CenterTheForm Me
    
    g.Styler.StyleForm Me
    
    mnuPref.Visible = False
    
    ' Get font from INI file
    strFont = GetIniFileProperty("RollsTable", "", "Fonts", g.strIniFile)
   
    ' Initialize the grid
    SetupGrid fgData, eGridMode_Grid
    With fgData
        .Redraw = flexRDNone
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .Cols = GDCol(eGDCol_Desc) + 1
        .FixedRows = 1
        .FixedCols = 0
        .TextMatrix(0, GDCol(eGDCol_Date)) = "Date"
        .TextMatrix(0, GDCol(eGDCol_Symbol)) = "Symbol"
        .TextMatrix(0, GDCol(eGDCol_From)) = "From"
        .TextMatrix(0, GDCol(eGDCol_To)) = "To"
        .TextMatrix(0, GDCol(eGDCol_Desc)) = "Description"
        .ColAlignment(GDCol(eGDCol_Date)) = flexAlignCenterCenter
        .ColAlignment(GDCol(eGDCol_Symbol)) = flexAlignCenterCenter
        .ColAlignment(GDCol(eGDCol_From)) = flexAlignCenterCenter
        .ColAlignment(GDCol(eGDCol_To)) = flexAlignCenterCenter
        .ColDataType(GDCol(eGDCol_Date)) = flexDTDouble
        .ColFormat(GDCol(eGDCol_Date)) = DateFormat("Format", MM_DD_YYYY)
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignLeftTop
        
        If strFont <> "" Then
            FontFromString .Font, strFont
        End If
        
        .Editable = flexEDNone
        .Redraw = flexRDBuffered
    End With
    
    ' Get the defaults out of INI file
    strSaveSym = GetIniFileProperty("SymSave", "", "RollsTable", g.strIniFile)
    
    InitCboSym strSaveSym
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmRollsTable.Form.Load", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub Form_Resize()
On Error Resume Next

    If LimitFormSize(Me, fraButtons.Width, fraButtons.Height * 2) Then Exit Sub

    fraButtons.Left = Me.ScaleLeft
    With fgData
        .Move fraButtons.Left, fraButtons.Height, Me.ScaleWidth, Me.ScaleHeight - fraButtons.Height
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    SetIniFileProperty "SymSave", GetSelSym(), "RollsTable", g.strIniFile
    SetIniFileProperty "RollsTable", FontToString(fgData.Font), "Fonts", g.strIniFile

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmRollsTable.Form.Unload", eGDRaiseError_Show
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

    ChangeGridFont fgData, True

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmRollsTable.mnuChangeFont.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub mnuSetChart_Click()
On Error GoTo ErrSection:

    SetActiveChartSymbol fgData.TextMatrix(fgData.Row, GDCol(eGDCol_Symbol))

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmRollsTable.mnuSetChart.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub opt65_Click()
On Error GoTo ErrSection:

    If Me.Visible Then
        SetGridData GetSelSym()
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmRollsTable.opt65.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub opt66_Click()
On Error GoTo ErrSection:

    If Me.Visible Then
        SetGridData GetSelSym()
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmRollsTable.opt66.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub opt67_Click()
On Error GoTo ErrSection:

    If Me.Visible Then
        SetGridData GetSelSym()
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmRollsTable.opt67.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Public Sub GenerateReport(ByVal vArgs As Variant)
On Error GoTo ErrSection:

    With frmPrintPreview.vp
        .StartDoc
        DoPrintHeader
        
        .RenderControl = fgData.hWnd

        .EndDoc
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmRollsTable.GenerateReport", eGDRaiseError_Raise

End Sub

Public Function PrintMe() As Boolean
On Error GoTo ErrSection

    PrintMe = frmPrintPreview.ShowMe("CNV RollsTable", Me)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmRollsTable.PrintMe", eGDRaiseError_Raise
    
End Function

