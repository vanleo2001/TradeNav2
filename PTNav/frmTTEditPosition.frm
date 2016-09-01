VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmTTEditPosition 
   Caption         =   "Form1"
   ClientHeight    =   3825
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   7530
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3825
   ScaleWidth      =   7530
   Begin VB.Frame fraProfit 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   315
      Left            =   3480
      TabIndex        =   21
      Top             =   1860
      Width           =   2175
      Begin VB.TextBox txtProfit 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   45
         Width           =   1035
      End
      Begin VB.Label lblProfit 
         Caption         =   "Closed Profit:"
         Height          =   195
         Left            =   0
         TabIndex        =   22
         Top             =   45
         Width           =   1035
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid fgFills 
      Height          =   795
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   5535
      _cx             =   9763
      _cy             =   1402
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
   Begin VB.Frame fraPriceDisplay 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   795
      Left            =   5940
      TabIndex        =   17
      Top             =   2880
      Width           =   1275
      Begin VB.OptionButton optTradingUnits 
         Caption         =   "Trading Units"
         Height          =   255
         Left            =   0
         TabIndex        =   19
         Top             =   240
         Value           =   -1  'True
         Width           =   1275
      End
      Begin VB.OptionButton optDecimal 
         Caption         =   "Decimal"
         Height          =   255
         Left            =   0
         TabIndex        =   20
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label lblUnits 
         Caption         =   "Display Prices in:"
         Height          =   255
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   1275
      End
   End
   Begin VB.Frame fraMisc 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      Begin VB.TextBox txtSymbol 
         Height          =   315
         Left            =   840
         TabIndex        =   2
         Top             =   0
         Width           =   1155
      End
      Begin VB.CommandButton cmdSymbolLookup 
         Caption         =   "&Lookup"
         Height          =   315
         Left            =   2040
         TabIndex        =   3
         Top             =   0
         Width           =   795
      End
      Begin VB.Label lblPosition 
         Height          =   255
         Left            =   3720
         TabIndex        =   5
         Top             =   30
         Width           =   1755
      End
      Begin VB.Label lblPos 
         Caption         =   "Position:"
         Height          =   255
         Left            =   3000
         TabIndex        =   4
         Top             =   30
         Width           =   795
      End
      Begin VB.Label lblSymbol 
         Caption         =   "Symbol:"
         Height          =   255
         Left            =   0
         TabIndex        =   1
         Top             =   30
         Width           =   795
      End
   End
   Begin VB.Frame fraButtons 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2595
      Left            =   5940
      TabIndex        =   10
      Top             =   120
      Width           =   1275
      Begin VB.CommandButton cmdNewFill 
         Caption         =   "&New Fill"
         Height          =   375
         Left            =   0
         TabIndex        =   14
         Top             =   1380
         Width           =   1275
      End
      Begin VB.CommandButton cmdEditFill 
         Caption         =   "&Edit Fill"
         Height          =   375
         Left            =   0
         TabIndex        =   15
         Top             =   1800
         Width           =   1275
      End
      Begin VB.CommandButton cmdDeleteFill 
         Caption         =   "&Delete Fill"
         Height          =   375
         Left            =   0
         TabIndex        =   16
         Top             =   2220
         Width           =   1275
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   375
         Left            =   0
         TabIndex        =   13
         Top             =   900
         Width           =   1275
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   0
         TabIndex        =   12
         Top             =   420
         Width           =   1275
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&Save"
         Height          =   375
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   1275
      End
   End
   Begin VB.TextBox txtNotes 
      Height          =   855
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   2220
      Width           =   5475
   End
   Begin VB.Label lblFills 
      Caption         =   "Entries and Exits (Fills):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   660
      Width           =   2295
   End
   Begin VB.Label lblNotes 
      Caption         =   "Notes:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   1035
   End
   Begin VB.Menu mnuFills 
      Caption         =   "Fills"
      Begin VB.Menu mnuNewFill 
         Caption         =   "&New Fill"
      End
      Begin VB.Menu mnuEditFill 
         Caption         =   "&Edit Fill"
      End
      Begin VB.Menu mnuDeleteFill 
         Caption         =   "&Delete Fill"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChangeFont 
         Caption         =   "&Change Font"
      End
   End
End
Attribute VB_Name = "frmTTEditPosition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmTTEditPosition
'' Description: Allow the user to edit a position
''
'' Author:      Genesis Financial Data Services
''              425 E Woodmen Rd
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Enum eGDFillCols
    eGDFillCol_Action = 0
    eGDFillCol_PosQuantity
    eGDFillCol_Symbol
    eGDFillCol_Price
    eGDFillCol_Date
    eGDFillCol_Fees
    eGDFillCol_BrokerID
    eGDFillCol_OrderID
    eGDFillCol_FillID
    eGDFillCol_Quantity
    eGDFillCol_NumCols
End Enum

Private Type mPrivate
    bOK As Boolean
    strSecType As String
    lActivityID As Long
    lAccountID As Long
    lAutoTradeItemID As Long
    lPositionID As Long
    strExitPrice As String
    dForceClosedBefore As Double
    lNextFillID As Long
    
    dEntryPrice As Double
    dExitPrice As Double
    
    TradePosition As eTT_Position
    lNetPosition As Long
    bUpdatingFill As Boolean
    bStartNew As Boolean
    
    Position As cPtPosition
    Bars As cGdBars
    
    FillsToDelete As cGdTree
End Type
Private m As mPrivate

Private Function FillCol(ByVal Col As eGDFillCols) As Long
    FillCol = Col
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: If the user clicks Cancel, unload the form without saving the
''              changes
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancel_Click()
On Error GoTo ErrSection:

    m.bOK = False
    Me.Hide

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTTEditPosition.cmdCancel.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdDeleteFill_Click
'' Description: Allow the user to delete the currently selected fill
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdDeleteFill_Click()
On Error GoTo ErrSection:

    DeleteFill

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditPosition.cmdDeleteFill.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdEditFill_Click
'' Description: Allow the user to edit the currently selected fill
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdEditFill_Click()
On Error GoTo ErrSection:

    EditFill

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditPosition.cmdEditFill.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdNewFill_Click
'' Description: Allow the user to enter a new fill for this trade
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdNewFill_Click()
On Error GoTo ErrSection:

    NewFill

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditPosition.cmdNewFill.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOK_Click
'' Description: If the user clicks on OK, allow the ShowMe to save the changes
''              and unload the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lNetPosition As Long            ' Net position for the trade
    Dim strPosition As String           ' Position for the trade (Long, Short)
    Dim strAction As String             ' Action of the current row (Buy, Sell)
    Dim lQuantity As Long               ' Quantity of the current row
    Dim bReverse As Boolean             ' Has the trade reversed?
    Dim dTimeLastFill As Double         ' Time of the last fill

    If Len(Trim(txtSymbol.Text)) = 0 Or Len(Trim(txtSymbol.Text)) > 50 Then
        MoveFocus txtSymbol
        Err.Raise vbObjectError + 1000, , "Symbol must be between 1 and 50 Characters"
    End If
    
    If fgFills.Rows = fgFills.FixedRows Or BlankRow(fgFills.FixedRows) Then
        Err.Raise vbObjectError + 1000, , "There must be at least one Fill"
    End If
    
    With fgFills
        bReverse = False
        lNetPosition = 0&
        Select Case UCase(.TextMatrix(.FixedRows, FillCol(eGDFillCol_Action)))
            Case "BUY"
                strPosition = "LONG"
            Case "SELL"
                strPosition = "SHORT"
        End Select
        
        For lIndex = .FixedRows To .Rows - 1
            If BlankRow(lIndex) = False Then
                strAction = .TextMatrix(lIndex, FillCol(eGDFillCol_Action))
                lQuantity = CLng(ValOfText(.TextMatrix(lIndex, FillCol(eGDFillCol_PosQuantity))))
                
                Select Case UCase(strAction)
                    Case "BUY"
                        lNetPosition = lNetPosition + lQuantity
                        
                    Case "SELL"
                        lNetPosition = lNetPosition - lQuantity
                    
                End Select
                
                Select Case UCase(strPosition)
                    Case "LONG"
                        If lNetPosition < 0 Then bReverse = True
                    Case "SHORT"
                        If lNetPosition > 0 Then bReverse = True
                End Select
                
                dTimeLastFill = .RowData(lIndex).FillDate
                
                If bReverse Then
                    .Row = lIndex
                    .RowSel = lIndex
                    Exit For
                End If
            End If
        Next lIndex
    End With
    
    If bReverse Then
        Err.Raise vbObjectError + 1000, , "Trades cannot reverse position.  Either fix the quantity or break this trade into multiple trades."
    End If
    
    If (lNetPosition <> 0) And (m.dForceClosedBefore > 0) And (dTimeLastFill < m.dForceClosedBefore) Then
        Err.Raise vbObjectError + 1000, , "This trade must be flat before it is saved"
    End If
    
    m.bOK = True
    Me.Hide

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTTEditPosition.cmdOK.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdPrint_Click
'' Description: Allow the user to print the trade information
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdPrint_Click()
On Error GoTo ErrSection:

    PrintMe

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditPosition.cmdPrint.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdSymbolLookup_Click
'' Description: If the user clicks on the Lookup Symbol button, bring up the
''              symbol selector form for them
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdSymbolLookup_Click()
On Error GoTo ErrSection:

    SymbolLookup
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditPosition.cmdSymbolLookup.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgFills_AfterEdit
'' Description: After the user edits a cell, handle some things
'' Inputs:      Row and Column of the edit
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgFills_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    Dim dValue As Double                ' Value of the cell
    Dim lRedraw As Long                 ' Current state of the grid's redraw
    Dim strSymbol As String             ' Symbol being traded

    If Not Me.Visible Then Exit Sub

    With fgFills
        Select Case Col
            Case FillCol(eGDFillCol_Action)
                If Len(.TextMatrix(Row, FillCol(eGDFillCol_PosQuantity))) > 0 Then
                    AddBlankRow
                    SetPositionLabel
                    EnableControls
                End If
                If Len(.TextMatrix(Row, FillCol(eGDFillCol_Action))) > 0 Then
                    If Left(.TextMatrix(Row, FillCol(eGDFillCol_PosQuantity)), 2) = " <" Then
                        If OppositeDirection(Row) Then
                            .TextMatrix(Row, FillCol(eGDFillCol_PosQuantity)) = Parse(lblPosition.Caption, " ", 2)
                        Else
                            .TextMatrix(Row, FillCol(eGDFillCol_PosQuantity)) = "1"
                        End If
                        .TextMatrix(Row, FillCol(eGDFillCol_Date)) = Now
                        If InStr(Trim(txtSymbol.Text), "-0") <> 0 Then
                            .TextMatrix(Row, FillCol(eGDFillCol_Symbol)) = RollSymbolForDate(Trim(txtSymbol.Text), Now)
                        Else
                            .TextMatrix(Row, FillCol(eGDFillCol_Symbol)) = Trim(txtSymbol.Text)
                        End If
                        If DM_GetBars(m.Bars, Trim(txtSymbol.Text), , LastDailyDownload) Then
                            If m.Bars(eBARS_Close, m.Bars.Size - 1) <> -999999# Then
                                .TextMatrix(Row, FillCol(eGDFillCol_Price)) = PriceDisplay(m.Bars(eBARS_Close, m.Bars.Size - 1), Trim(txtSymbol.Text), optTradingUnits)
                            End If
                        End If
                        AutoSize
                    End If
                End If
            
            Case FillCol(eGDFillCol_PosQuantity)
                lRedraw = .Redraw
                .Redraw = flexRDNone
                
                If Len(.TextMatrix(Row, FillCol(eGDFillCol_Action))) > 0 Then
                    AddBlankRow
                    SetPositionLabel
                    EnableControls
                End If
                
                If Not BlankRow(Row) Then
                    .TextMatrix(Row, Col) = Format(ValOfText(.TextMatrix(Row, Col)), "#,##0")
                End If
                AutoSize
                .Redraw = lRedraw
                
            Case FillCol(eGDFillCol_Price)
                lRedraw = .Redraw
                .Redraw = flexRDNone
                
                If Not BlankRow(Row) Then
                    strSymbol = Trim(txtSymbol.Text)
                    dValue = PriceFromDisplay(.TextMatrix(Row, Col), strSymbol)
                    .TextMatrix(Row, Col) = PriceDisplay(dValue, strSymbol, optTradingUnits)
                End If
                
                AutoSize
                .Redraw = lRedraw
                
            Case FillCol(eGDFillCol_Date)
                SortFillGrid
                
            Case FillCol(eGDFillCol_Fees)
                lRedraw = .Redraw
                .Redraw = flexRDNone
                
                If Not BlankRow(Row) Then
                    .TextMatrix(Row, Col) = Format(ValOfText(.TextMatrix(Row, Col)), "$#,##0.00")
                End If
                AutoSize
                .Redraw = lRedraw
                                
        End Select
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditPosition.fgFills.AfterEdit", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgFills_AfterRowColChange
'' Description: After the user moves the cell, put the cell into edit mode
'' Inputs:      Old Row and Column, New Row and Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgFills_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    If Not Me.Visible Then Exit Sub
    
    If NewCol <> FillCol(eGDFillCol_Action) And NewCol <> FillCol(eGDFillCol_Date) Then
        If NewRow >= fgFills.FixedRows And NewRow < fgFills.Rows Then
            fgFills.Select NewRow, NewCol
            fgFills.EditCell
        End If
    End If
    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditPosition.fgFills.AfterRowColChange", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgFills_BeforeEdit
'' Description: Only allow the user to edit certain fields
'' Inputs:      Row and Column of the edit, Whether to Cancel the edit
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgFills_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    If Not Me.Visible Then Exit Sub

    Select Case Col
        Case FillCol(eGDFillCol_OrderID)
            Cancel = True
            
        Case FillCol(eGDFillCol_Action)
            fgFills.ComboList = "Buy|Sell"
            
        Case FillCol(eGDFillCol_Date)
            fgFills.ComboList = "..."
            
        Case Else
            fgFills.ComboList = ""
            
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditPosition.fgFills.BeforeEdit", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgFills_BeforeRowColChange
'' Description: Before the user changes rows, try to save their changes
'' Inputs:      Old Row and Column, New Row and Column, Whether to Cancel move
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgFills_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    If (Not Me.Visible) Or m.bUpdatingFill Or m.bStartNew Then Exit Sub

    If NewRow <> OldRow And OldRow >= fgFills.FixedRows Then
        If Not BlankRow(OldRow) Then
            If UpdateFill(OldRow) = False Then Cancel = True
            EnableControls
        End If
    End If
    
    If m.bStartNew = True Then
        StartNew
        Cancel = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditPosition.fgFills.BeforeRowColChange", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgFills_CellButtonClick
'' Description: Allow the user to easily edit the date
'' Inputs:      Row and Column of the edit
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgFills_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    Dim pt As POINTAPI                  ' Point to display the form
    Dim dDate As Double                 ' Date from the grid
    
    ' Figure out the location to show the Edit Date form...
    pt.X = fgFills.ColPos(Col) / Screen.TwipsPerPixelX
    pt.Y = (fgFills.RowPos(Row) + fgFills.RowHeight(Row)) / Screen.TwipsPerPixelY
    ClientToScreen fgFills.hWnd, pt
    pt.X = pt.X * Screen.TwipsPerPixelX
    pt.Y = pt.Y * Screen.TwipsPerPixelY
    
    ' Get the current date from the grid...
    dDate = DateOf(fgFills.TextMatrix(Row, FillCol(eGDFillCol_Date)))
    
    ' Show the Edit Date form...
    dDate = frmEditDate.ShowMe(pt.X, pt.Y, dDate, Me, , , , , , , , HourMinuteSecond, As12Hour)
    
    ' Redisplay the date to the grid...
    fgFills.TextMatrix(Row, FillCol(eGDFillCol_Date)) = dDate
    AutoSize
    
    SortFillGrid
    SetPositionLabel

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditPosition.fgFills.CellButtonClick", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub fgFills_ChangeEdit()
On Error GoTo ErrSection:

    If fgFills.Col = FillCol(eGDFillCol_Action) Then
        If Len(fgFills.EditText) > 0 Then
            fgFills.Col = FillCol(eGDFillCol_PosQuantity)
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditPosition.fgFills.ChangeEdit", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgFills_ComboCloseUp
'' Description: Happens when the combo drop down is closing up
'' Inputs:      Row and Column of current cell,  Whether to Finish the Edit
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgFills_ComboCloseUp(ByVal Row As Long, ByVal Col As Long, FinishEdit As Boolean)
On Error GoTo ErrSection:

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditPosition.fgFills.ComboCloseUp", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgFills_ComboDropDown
'' Description: Happens when the combo drop down is about to occur
'' Inputs:      Row and Column of current cell
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgFills_ComboDropDown(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditPosition.fgFills.ComboDropDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgFills_DblClick
'' Description: Allow the user to edit a fill by double clicking on it
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgFills_DblClick()
On Error GoTo ErrSection:

    Dim lRow As Long                    ' Current Mouse Row in the grid
    
    With fgFills
        lRow = .MouseRow
        If lRow >= .FixedRows And lRow < .Rows Then
            .Row = lRow
            .RowSel = lRow
            
            EditFill
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditPosition.fgFills.DblClick", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub fgFills_GotFocus()
On Error GoTo ErrSection:

    If Len(Trim(txtSymbol.Text)) = 0 Then
        MoveFocus txtSymbol
        Err.Raise vbObjectError + 1000, , "Please Enter a Symbol"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditPosition.fgFills.GotFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgFills_KeyDown
'' Description: Allow the user to do various things with the keyboard
'' Inputs:      Code of the Key Pressed, Shift/Ctrl/Alt status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgFills_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    fgFills.RowSel = fgFills.Row

    Select Case KeyCode
        Case vbKeyReturn
            EditFill
        
        Case vbKeyInsert
            NewFill
        
        Case vbKeyDelete
            DeleteFill
            
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditPosition.fgFills.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgFills_LostFocus
'' Description: When the control loses focus, update the fill if necessary
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgFills_LostFocus()
On Error GoTo ErrSection:

    If (Not Me.Visible) Or (m.bUpdatingFill) Or (cmdCancel Is Screen.ActiveControl) Or m.bStartNew Then
        Exit Sub
    End If
    
    If fgFills.RowSel >= fgFills.FixedRows Then
        If Not BlankRow(fgFills.RowSel) Then
            If UpdateFill(fgFills.RowSel) = False Then MoveFocus fgFills
            EnableControls
        End If
    End If
    
    If m.bStartNew = True Then StartNew

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditPosition.fgFills.LostFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgFills_MouseDown
'' Description: Show the user a popup menu if they right click on the grid
'' Inputs:      Mouse Button pressed, Shift/Ctrl/Alt status, Mouse Location
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgFills_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    Dim lMouseRow As Long               ' Current Mouse Row in the grid
    Dim lMouseCol As Long               ' Current Mouse Column in the grid
    
    With fgFills
        lMouseRow = .MouseRow
        lMouseCol = .MouseCol
        
        If Button = vbRightButton Then
            If lMouseRow >= .FixedRows And lMouseRow < .Rows Then
                .RowSel = lMouseRow
                If Not .IsSelected(lMouseRow) Then .Row = lMouseRow
            End If
        
            mnuDeleteFill.Enabled = (lMouseRow >= .FixedRows And lMouseRow < .Rows)
            mnuEditFill.Enabled = (lMouseRow >= .FixedRows And lMouseRow < .Rows)
            
            PopupMenu mnuFills
            
            Select Case mnuFills.Tag
                Case "New": mnuNewFill_Click
                Case "Edit": mnuEditFill_Click
                Case "Delete": mnuDeleteFill_Click
            End Select
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditPosition.fgFills.MouseDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgFills_ValidateEdit
'' Description: Make sure that the data in the field is valid
'' Inputs:      Row and Column of the edit, Whether to Cancel the edit
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgFills_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditPosition.fgFills.ValidateEdit", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_KeyDown
'' Description: If the user pressed F1, show the help menu
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
    RaiseError "frmTTEditPosition.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: When the form is loaded, do some initialization
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Dim strFont As String               ' Font information from the ini file
    Dim strTemp As String               ' Entry from the ini file

    Me.Caption = "Edit Trade"
    Me.Icon = Picture16(ToolbarIcon("ID_TradeTracker"))
    
    mnuFills.Visible = False
    
    strFont = GetIniFileProperty("TTEditPosition.Fills", "", "Fonts", g.strIniFile)
    If strFont <> "" Then FontFromString fgFills.Font, strFont

    strTemp = GetIniFileProperty("TTEditPosition", "", "Placement", g.strIniFile)
    If Len(strTemp) = 0 Then
        CenterTheForm Me
    Else
        SetFormPlacement Me, strTemp
    End If
    
    optTradingUnits = GetIniFileProperty("TradingUnits", True, "TTEditPosition", g.strIniFile)
    optDecimal = Not optTradingUnits

    Set m.Bars = New cGdBars
    
    cmdSymbolLookup.ToolTipText = "Lookup a Symbol to Trade"
    
    m.lNextFillID = -99999
    m.bUpdatingFill = False
    
    Set m.FillsToDelete = New cGdTree
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditPosition.Form.Load", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Show the form and allow the user to make changes
'' Inputs:      ID of the Activity in the database
'' Returns:     True if user chose OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(lActivityID As Long, lPositionID As Long, ByVal lAccountID As Long, Optional ByVal dForceClosedBefore As Double = 0#) As Boolean
On Error GoTo ErrSection:

    m.lPositionID = lPositionID
    m.lAccountID = lAccountID
    m.lActivityID = lActivityID
    m.dForceClosedBefore = dForceClosedBefore
    
    InitFillsGrid
    
    If lPositionID = 0& Then
        Me.Caption = "New Trade"
        txtNotes.Text = ""
        m.lAutoTradeItemID = 0&
    Else
        Me.Caption = "Edit Trade #" & Str(m.lPositionID)
        Set m.Position = New cPtPosition
        If m.Position.Load(lPositionID) Then
            m.lAutoTradeItemID = m.Position.AutoTradingItemID
            txtSymbol.Text = m.Position.Symbol
            m.strSecType = m.Position.SecurityType
            txtNotes.Text = m.Position.Notes
            LoadFillsGrid
        End If
    End If
    
    AddBlankRow
    SetPositionLabel
    
    If fgFills.Rows > fgFills.FixedRows Then
        fgFills.RowSel = fgFills.Rows - 1
        fgFills.Row = fgFills.Rows - 1
        fgFills.Col = FillCol(eGDFillCol_Action)
    End If
    
    CalcProfit
    EnableControls
    
    ShowForm Me, True
    
    If m.bOK Then
        Save
        lActivityID = m.lActivityID
        lPositionID = m.lPositionID
    End If
    
ErrExit:
    ShowMe = m.bOK
    Unload Me
    Exit Function

ErrSection:
    ShowMe = False
    Unload Me
    RaiseError "frmTTEditPosition.ShowMe", eGDRaiseError_Raise

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: If the user hits the X, cancel the unload and let the ShowMe
''              take over
'' Inputs:      Whether or not to Cancel the Unload, Mode of the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode = 0 Then
        m.bOK = False
        Cancel = True
        Me.Hide
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTTEditPosition.Form.QueryUnload", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Save
'' Description: Save the trade to the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Save()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim Activity As cPtActivity         ' Activity object
    Dim rs As Recordset                 ' Recordset into the database
    Dim Fill As New cPtFill             ' Fill object
    
    If m.FillsToDelete.Count > 0 Then
        For lIndex = 1 To m.FillsToDelete.Count
            Set Fill = m.FillsToDelete(lIndex)
            If Fill.Quantity = Fill.PosQuantity Then
                Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblFills] WHERE [FillID]=" & Fill.FillID & ";", dbOpenDynaset)
                If Not (rs.BOF And rs.EOF) Then rs.Delete
                
                Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrders] WHERE [OrderID]=" & Fill.OrderID & ";", dbOpenDynaset)
                If Not (rs.BOF And rs.EOF) Then rs.Delete
            Else
                Fill.Quantity = Fill.Quantity - Fill.PosQuantity
                Fill.Save
            End If
        Next lIndex
    End If
    
    If m.lActivityID = 0 Then
        Set Activity = New cPtActivity
        With Activity
            .AccountID = m.lAccountID
            .ActivityType = eTT_ActivityType_Trading
            .StartDate = Now
            .Save
            
            m.lActivityID = .ActivityID
        End With
    End If

    If m.Position Is Nothing Then Set m.Position = New cPtPosition
    With m.Position
        .ActivityID = m.lActivityID
        .Symbol = Trim(txtSymbol.Text)
        .SecurityType = m.strSecType
        .SymbolID = g.SymbolPool.SymbolIDforSymbol(.Symbol)
        .Notes = Trim(txtNotes.Text)
        Select Case Trim(UCase(Parse(lblPosition.Caption, " ", 1)))
            Case "NONE"
                .Position = eTT_Position_None
            Case "LONG"
                .Position = eTT_Position_Long
            Case "SHORT"
                .Position = eTT_Position_Short
        End Select
        .IsOpen = (NetPosition <> 0)
        
        .Fills.Clear
        For lIndex = fgFills.FixedRows To fgFills.Rows - 1
            If Not BlankRow(lIndex) Then
                .Fills(Str(fgFills.RowData(lIndex).FillID)) = fgFills.RowData(lIndex)
            End If
        Next lIndex
        
        .OpenSymbol = .Fills(.Fills.Count).Symbol
        .OpenSymbolID = GetSymbolID(.OpenSymbol)
        
        .CalcPrices
        
        .Save
        
        If Not g.Broker.FillSummary(.AccountID, .SymbolOrSymbolID, .AutoTradingItemID) Is Nothing Then
            g.Broker.FillSummary(.AccountID, .SymbolOrSymbolID, .AutoTradingItemID).RecalculateHistory
        End If
        
        m.lPositionID = .PositionID
    End With
    
    RefreshPosition m.Position

ErrExit:
    Set Activity = Nothing
    Exit Sub

ErrSection:
    Set Activity = Nothing
    RaiseError "frmTTEditPosition.Save", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Resize
'' Description: Move and resize controls as the form is resized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Resize()
On Error Resume Next

    Dim lMinWidth As Long               ' Minimum Form Width
    Dim lMinHeight As Long              ' Minimum Form Height
    
    lMinWidth = fraMisc.Width + fraButtons.Width + (fraMisc.Left * 3)
    lMinHeight = fraButtons.Height + fraPriceDisplay.Height + (fraButtons.Top * 3)
    
    If LimitFormSize(Me, lMinWidth, lMinHeight) Then Exit Sub
    
    With fraButtons
        .Move ScaleWidth - .Width - fraMisc.Left
    End With
    
    With fraPriceDisplay
        .Move fraButtons.Left, fraButtons.Height + (fraButtons.Top * 2)
    End With
    
    With txtNotes
        .Move fraMisc.Left, ScaleHeight - .Height - fraMisc.Top, _
                ScaleWidth - fraButtons.Width - (fraMisc.Left * 3)
    End With

    With lblNotes
        .Move fraMisc.Left, txtNotes.Top - .Height
    End With
    
    With lblFills
        .Move fraMisc.Left, (fraMisc.Height + fraMisc.Top * 2)
    End With
    
    With fgFills
        .Move fraMisc.Left, lblFills.Top + lblFills.Height, txtNotes.Width, _
                lblNotes.Top - fraMisc.Top - lblFills.Top - lblFills.Height
    End With
    
    With fraProfit
        .Move fgFills.Width + fgFills.Left - .Width, fgFills.Top + fgFills.Height
    End With
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: When the form is unloaded, clean up after ourselves
'' Inputs:      Whether or not to Cancel the unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    Set m.Bars = Nothing
    Set m.Position = Nothing
    Set m.FillsToDelete = Nothing
    
    SetIniFileProperty "TTEditPosition.Fills", FontToString(fgFills.Font), "Fonts", g.strIniFile
    SetIniFileProperty "TTEditPosition", GetFormPlacement(Me), "Placement", g.strIniFile
    SetIniFileProperty "TradingUnits", optTradingUnits, "TTEditPosition", g.strIniFile

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTTEditPosition.Form.Unload", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuChangeFont_Click
'' Description: Allow the user to change fonts on the fills grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuChangeFont_Click()
On Error GoTo ErrSection:

    ChangeGridFont fgFills

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditPosition.mnuChangeFont.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuDeleteFill_Click
'' Description: Allow the user to delete the selected fill
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuDeleteFill_Click()
On Error GoTo ErrSection:

    If mnuFills.Tag = "" Then
        mnuFills.Tag = "Delete"
    Else
        DeleteFill
        mnuFills.Tag = ""
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditPosition.mnuDeleteFill.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuEditFill_Click
'' Description: Allow the user to edit the selected fill
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuEditFill_Click()
On Error GoTo ErrSection:

    If mnuFills.Tag = "" Then
        mnuFills.Tag = "Edit"
    Else
        EditFill
        mnuFills.Tag = ""
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditPosition.mnuEditFill.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuNewFill_Click
'' Description: Allow the user to create a new fill
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuNewFill_Click()
On Error GoTo ErrSection:

    If mnuFills.Tag = "" Then
        mnuFills.Tag = "New"
    Else
        NewFill
        mnuFills.Tag = ""
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditPosition.mnuNewFill.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optDecimal_Click
'' Description: If the user clicks on the Decimal option, change the price
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optDecimal_Click()
On Error GoTo ErrSection:

    RedisplayPrices

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTTEditPosition.optDecimal.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optTradingUnits_Click
'' Description: If the user clicks on the Trading Units option, change the price
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optTradingUnits_Click()
On Error GoTo ErrSection:

    RedisplayPrices

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTTEditPosition.optTradingUnits.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SymbolLookup
'' Description: Allow the user to lookup a symbol
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SymbolLookup()
On Error GoTo ErrSection:

    Dim astrSymbols As New cGdArray     ' Symbol that the user selected
    
    astrSymbols.Create eGDARRAY_Strings
    Set astrSymbols = frmSymbolSelector.ShowMe(Trim(txtSymbol.Text), False, , "Select a Symbol to Trade")
    If astrSymbols.Size > 0 Then
        txtSymbol.Text = astrSymbols(0)
        MoveFocus txtSymbol
        m.strSecType = SecType
        ChangeSymbols
    End If
    
    EnableControls

ErrExit:
    Set astrSymbols = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditPosition.SymbolLookup", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EnableControls
'' Description: Enable controls based on the current state of things
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EnableControls()
On Error GoTo ErrSection:

    If Len(Trim(txtSymbol.Text)) = 0 Then
        Disable cmdOK
        Disable cmdPrint
        Disable cmdNewFill
        Disable cmdEditFill
        Disable cmdDeleteFill
        Disable optTradingUnits
        Disable optDecimal
    Else
        Enable cmdOK, fgFills.Rows > fgFills.FixedRows
        Enable cmdPrint, fgFills.Rows > fgFills.FixedRows
        Enable cmdNewFill
        Enable optTradingUnits
        Enable optDecimal
        
        With fgFills
            Enable cmdEditFill, (.Rows > .FixedRows) And (Not BlankRow(.RowSel))
            Enable cmdDeleteFill, (.Rows > .FixedRows) And (Not BlankRow(.RowSel))
        
            .ColHidden(FillCol(eGDFillCol_Symbol)) = (InStr(txtSymbol.Text, "-0") = 0)
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditPosition.EnableControls", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RedisplayPrices
'' Description: Redisplay all prices based on the Trading Units radio button
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RedisplayPrices()
On Error GoTo ErrSection:

    Dim lRedraw As Long                 ' Current state of the grid's redraw
    Dim lRow As Long                    ' Index into a for loop
    Dim dValue As Double                ' Value to redisplay
    Dim strSymbol As String             ' Symbol of the position
    
    strSymbol = Trim(txtSymbol.Text)
    With fgFills
        lRedraw = .Redraw
        .Redraw = flexRDNone
                
        For lRow = .FixedRows To .Rows - 1
            dValue = PriceFromDisplay(.TextMatrix(lRow, FillCol(eGDFillCol_Price)), strSymbol)
            .TextMatrix(lRow, FillCol(eGDFillCol_Price)) = PriceDisplay(dValue, strSymbol, optTradingUnits.Value, True)
        Next lRow
        
        AutoSize
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTTEditPosition.RedisplayPrices", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenerateReport
'' Description: Set up the Print for the position
'' Inputs:      Arguments passed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GenerateReport(ByVal vArgs As Variant)
On Error GoTo ErrSection:
    
    Dim lRow As Long                    ' Index into a for loop
    Dim lCol As Long                    ' Index into a for loop
    Dim strText As String               ' Text to display
    
    With frmPrintPreview.vp
        .StartDoc
        DoPrintHeader
        
        .FontName = "Times New Roman"
        .fontSize = 14
        .fontBold = True
        .TextAlign = taCenterMiddle
        If m.lPositionID = 0 Then
            .Text = "New Trade"
        Else
            .Text = "Trade #" & Str(m.lPositionID)
        End If
        .TextAlign = taLeftMiddle
        .fontBold = False
        
        .Text = vbLf & vbLf
        .Text = "Symbol: " & vbTab & Trim(txtSymbol.Text) & vbLf
        .Text = "Notes: " & vbTab & Trim(txtNotes.Text) & vbLf & vbLf
        
        .Paragraph = ""
        
        .Text = "Fills:" & vbLf
        If BlankRow(fgFills.Rows - 1) Then fgFills.Rows = fgFills.Rows - 1
        If frmPrintPreview.GoingToFile Then
            With fgFills
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
            .RenderControl = fgFills.hWnd
        End If
        AddBlankRow
        
        .EndDoc
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditPosition.GenerateReport", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PrintMe
'' Description: Start the print process for the position
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function PrintMe()
On Error GoTo ErrSection

    PrintMe = frmPrintPreview.ShowMe("CNV TTEditPosition", Me)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTTEditPosition.PrintMe", eGDRaiseError_Raise

End Function

Private Sub txtNotes_GotFocus()
On Error GoTo ErrSection:

    If Len(Trim(txtSymbol.Text)) = 0 Then
        MoveFocus txtSymbol
        Err.Raise vbObjectError + 1000, , "Please Enter a Symbol"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditPosition.txtNotes.GotFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtSymbol_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtSymbol

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditPosition.txtSymbol.GotFocus", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub InitFillsGrid()
On Error GoTo ErrSection:

    Dim lRedraw As Long                 ' Current state of the grid's redraw

    With fgFills
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        .AllowBigSelection = False
        .AllowSelection = False
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .Editable = flexEDKbdMouse ' = flexEDNone
        .ExplorerBar = flexExNone
        .ExtendLastCol = True
        .ScrollTrack = True
        .SelectionMode = flexSelectionListBox
        .SheetBorder = RGB(128, 128, 128)
        .TabBehavior = flexTabCells
        
        .Rows = 1
        .FixedRows = 1
        .Cols = FillCol(eGDFillCol_NumCols)
        .FixedCols = 0
        
        .TextMatrix(0, FillCol(eGDFillCol_Action)) = "Action"
        .TextMatrix(0, FillCol(eGDFillCol_PosQuantity)) = "Quantity"
        .TextMatrix(0, FillCol(eGDFillCol_Symbol)) = "Symbol"
        .TextMatrix(0, FillCol(eGDFillCol_Price)) = "Price"
        .TextMatrix(0, FillCol(eGDFillCol_Date)) = "Date"
        .TextMatrix(0, FillCol(eGDFillCol_Fees)) = "Fees"
        .TextMatrix(0, FillCol(eGDFillCol_BrokerID)) = "Fill ID"
        .TextMatrix(0, FillCol(eGDFillCol_OrderID)) = "Order ID"
        
        .ColDataType(FillCol(eGDFillCol_Date)) = flexDTDate
        .ColFormat(FillCol(eGDFillCol_Date)) = DateFormat("Format", , HH_MM_SS, AMPM_LOWER)
        .ColAlignment(FillCol(eGDFillCol_Date)) = flexAlignCenterTop
        .ColAlignment(FillCol(eGDFillCol_BrokerID)) = flexAlignLeftTop
        
        .ColHidden(FillCol(eGDFillCol_OrderID)) = True
        .ColHidden(FillCol(eGDFillCol_FillID)) = True
        .ColHidden(FillCol(eGDFillCol_Quantity)) = True
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignLeftTop
        .Redraw = lRedraw
    End With
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditPosition.InitFillsGrid", eGDRaiseError_Raise
    
End Sub

Private Sub LoadFillsGrid()
On Error GoTo ErrSection:

    Dim lRedraw As Long
    Dim lFill As Long
    Dim Fill As cPtFill
    
    With fgFills
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        For lFill = 1 To m.Position.Fills.Count
            Set Fill = m.Position.Fills(lFill)
                
            .Rows = .Rows + 1
            FillToGrid .Rows - 1, Fill
        Next lFill
        
        SortFillGrid
        If .Rows > .FixedRows Then
            .RowSel = .FixedRows
            .Row = .FixedRows
            .Col = FillCol(eGDFillCol_Action)
        End If
        
        AutoSize
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditPosition.LoadFillsGrid", eGDRaiseError_Raise
    
End Sub

Private Sub SetPositionLabel()
On Error GoTo ErrSection:

    Dim lQuantity As Long               ' Net Position

    With fgFills
        If .Rows = .FixedRows Then
            lblPosition.Caption = "None"
            m.TradePosition = eTT_Position_None
            m.lNetPosition = 0&
        Else
            lQuantity = NetPosition
            
            Select Case UCase(fgFills.TextMatrix(.FixedRows, FillCol(eGDFillCol_Action)))
                Case "NONE", ""
                    lblPosition.Caption = "None"
                    m.TradePosition = eTT_Position_None
                    m.lNetPosition = 0&
                
                Case "BUY"
                    If lQuantity <> 0 Then
                        lblPosition.Caption = "Long " & Format(lQuantity, "#,##0")
                    Else
                        lblPosition.Caption = "Long"
                    End If
                    m.TradePosition = eTT_Position_Long
                    m.lNetPosition = lQuantity
                
                Case "SELL"
                    If lQuantity <> 0 Then
                        lblPosition.Caption = "Short " & Format(Abs(lQuantity), "#,##0")
                    Else
                        lblPosition.Caption = "Short"
                    End If
                    m.TradePosition = eTT_Position_Short
                    m.lNetPosition = lQuantity
            End Select
        End If
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditPosition.SetPositionLabel", eGDRaiseError_Raise
    
End Sub

Private Sub txtSymbol_LostFocus()
On Error GoTo ErrSection:

    If Len(Trim(txtSymbol.Text)) > 0 Then
        txtSymbol.Text = Trim(UCase(txtSymbol.Text))
        m.strSecType = SecType
        ChangeSymbols
    End If
    
    EnableControls
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditPosition.txtSymbol.LostFocus", eGDRaiseError_Show
    MoveFocus txtSymbol
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NetPosition
'' Description: Determine the net position of the fills in the grid
'' Inputs:      None
'' Returns:     Net Position
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function NetPosition() As Long
On Error GoTo ErrSection:

    Dim lIndex As Long
    Dim lPosition As Long
    
    lPosition = 0&
    With fgFills
        For lIndex = .FixedRows To .Rows - 1
            Select Case UCase(.TextMatrix(lIndex, FillCol(eGDFillCol_Action)))
                Case "BUY"
                    lPosition = lPosition + ValOfText(.TextMatrix(lIndex, FillCol(eGDFillCol_PosQuantity)))
                
                Case "SELL"
                    lPosition = lPosition - ValOfText(.TextMatrix(lIndex, FillCol(eGDFillCol_PosQuantity)))
            End Select
        Next lIndex
    End With
    
    NetPosition = lPosition

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTTEditPosition.NetPosition", eGDRaiseError_Raise
    
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SortFillGrid
'' Description: Sort the fill grid by Date and then by FillID
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SortFillGrid()
On Error GoTo ErrSection:

    Dim lRow As Long                    ' Last row to sort

    With fgFills
        If .Rows > .FixedRows Then
            If BlankRow(.Rows - 1) = True Then
                lRow = .Rows - 2
                .TextMatrix(.Rows - 1, FillCol(eGDFillCol_Date)) = "99999"
            Else
                lRow = .Rows - 1
            End If
            
            If lRow >= .FixedRows Then
                .Select .FixedRows, FillCol(eGDFillCol_FillID), lRow, FillCol(eGDFillCol_FillID)
                .Sort = flexSortGenericAscending
                .Select .FixedRows, FillCol(eGDFillCol_Date), lRow, FillCol(eGDFillCol_Date)
                .Sort = flexSortGenericAscending
            End If
        
            If .TextMatrix(.Rows - 1, FillCol(eGDFillCol_Date)) = "99999" Then .TextMatrix(.Rows - 1, FillCol(eGDFillCol_Date)) = ""
            
            .Row = .Rows - 1
            .RowSel = .Rows - 1
            .Col = FillCol(eGDFillCol_Action)
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditPosition.SortFillGrid", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddBlankRow
'' Description: Add a blank row if applicable
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddBlankRow()
On Error GoTo ErrSection:

    Dim lRedraw As Long                 ' Current state of the grid's redraw
    Dim lNetPos As Long                 ' Net position of the trade so far

    With fgFills
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        lNetPos = NetPosition
        If .Rows = .FixedRows Or lNetPos <> 0 Then
            If Len(.TextMatrix(.Rows - 1, FillCol(eGDFillCol_Action))) > 0 Then
                .Rows = .Rows + 1
                '.TextMatrix(.Rows - 1, FillCol(eGDFillCol_Date)) = Now
                'If InStr(Trim(txtSymbol.Text), "-0") <> 0 Then
                '    .TextMatrix(.Rows - 1, FillCol(eGDFillCol_Symbol)) = RollSymbolForDate(Trim(txtSymbol.Text), Now)
                'Else
                '    .TextMatrix(.Rows - 1, FillCol(eGDFillCol_Symbol)) = Trim(txtSymbol.Text)
                'End If
                .TextMatrix(.Rows - 1, FillCol(eGDFillCol_PosQuantity)) = " <-- Set Action to Buy or Sell for New Fill "
                .MergeCells = flexMergeSpill
                .MergeRow(.Rows - 1) = True
                AutoSize
                
                If .Rows - 1 = .FixedRows Then
                    .RowSel = .Rows - 1
                    .Row = .Rows - 1
                End If
            End If
        ElseIf lNetPos = 0 And .Rows > .FixedRows + 1 Then
            If Len(.TextMatrix(.Rows - 1, FillCol(eGDFillCol_Action))) = 0 Then
                .Rows = .Rows - 1
            End If
        End If
        
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditPosition.AddBlankRow", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NextFillID
'' Description: Return the next temporary Fill ID to use
'' Inputs:      None
'' Returns:     Next Fill ID
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function NextFillID() As Long
On Error GoTo ErrSection:

    NextFillID = m.lNextFillID
    m.lNextFillID = m.lNextFillID + 1

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTTEditPosition.NextFillID", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DeleteFill
'' Description: Allow the user to delete the currently selected fill
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DeleteFill()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim lRow As Long                    ' Row in the grid to delete
    Dim lIndex As Long
    Dim Fill As New cPtFill
    Dim Order As New cPtOrder
    Dim bUpdate As Boolean

    lRow = fgFills.RowSel
    If lRow < fgFills.FixedRows Or lRow >= fgFills.Rows Then Exit Sub
    
    Set Fill = fgFills.RowData(lRow)
    
    If InfBox("Are you sure you want to delete this fill?", "?", "+Yes|-No", "Warning") = "Y" Then
        m.FillsToDelete.Add Fill, Str(Fill.FillID)
        fgFills.RemoveItem lRow
    End If
    
    AddBlankRow
    SetPositionLabel
    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditPosition.DeleteFill", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EditFill
'' Description: Allow the user to edit the currently selected fill
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EditFill()
On Error GoTo ErrSection:

    Dim Fill As New cPtFill             ' Fill to edit
    Dim lRow As Long                    ' Currently selected row in the grid
    
    lRow = fgFills.RowSel
    If lRow < fgFills.FixedRows Or lRow >= fgFills.Rows Then Exit Sub
    If BlankRow(lRow) Then Exit Sub
    
    Set Fill = fgFills.RowData(lRow)
    If frmTTEditFill.ShowMe(Fill, Trim(txtSymbol.Text), m.lAccountID) = True Then
        FillToGrid lRow, Fill
        
        SortFillGrid
        AddBlankRow
        SetPositionLabel
    End If
    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditPosition.EditFill", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NewFill
'' Description: Allow the user to create a new fill for this trade
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub NewFill()
On Error GoTo ErrSection:

    Dim lRedraw As Long                 ' Current state of the grid's redraw
    Dim Fill As New cPtFill             ' Fill object to fill in
    Dim lRow As Long

    If frmTTEditFill.ShowMe(Fill, Trim(txtSymbol.Text), m.lAccountID) Then
        With fgFills
            lRedraw = .Redraw
            .Redraw = flexRDNone
            
            If Not BlankRow(.Rows - 1) Then .Rows = .Rows + 1
            lRow = .Rows - 1
            FillToGrid lRow, Fill
            
            SortFillGrid
            AddBlankRow
            SetPositionLabel
            
            EnableControls
            .Redraw = lRedraw
        End With
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditPosition.NewFill", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillToGrid
'' Description: Fill in the row in the grid with information from the fill
'' Inputs:      Row to edit, Fill, Symbol for the fill
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FillToGrid(ByVal Row As Long, ByVal Fill As cPtFill)
On Error GoTo ErrSection:

    Dim lRedraw As Long                 ' Current state of the grid's redraw
    
    With fgFills
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        If Fill.Buy Then
            .TextMatrix(Row, FillCol(eGDFillCol_Action)) = "Buy"
        Else
            .TextMatrix(Row, FillCol(eGDFillCol_Action)) = "Sell"
        End If
        .TextMatrix(Row, FillCol(eGDFillCol_PosQuantity)) = Format(Fill.PosQuantity, "#,##0")
        If InStr(Trim(txtSymbol.Text), "-0") <> 0 Then
            .TextMatrix(Row, FillCol(eGDFillCol_Symbol)) = RollSymbolForDate(Trim(txtSymbol.Text), Now)
        Else
            .TextMatrix(Row, FillCol(eGDFillCol_Symbol)) = Trim(txtSymbol.Text)
        End If
        .TextMatrix(Row, FillCol(eGDFillCol_Price)) = PriceDisplay(Fill.Price, Trim(txtSymbol.Text), optTradingUnits)
        .TextMatrix(Row, FillCol(eGDFillCol_Date)) = Str(Fill.FillDate)
        .TextMatrix(Row, FillCol(eGDFillCol_Fees)) = Format(Fill.Fees, "$#,##0.00")
        .TextMatrix(Row, FillCol(eGDFillCol_BrokerID)) = Fill.BrokerID
        If Fill.FillID = 0 Then Fill.FillID = NextFillID
        .TextMatrix(Row, FillCol(eGDFillCol_FillID)) = Str(Fill.FillID)
        .TextMatrix(Row, FillCol(eGDFillCol_OrderID)) = Str(Fill.OrderID)
        .TextMatrix(Row, FillCol(eGDFillCol_Quantity)) = Str(Fill.Quantity)
        
        .RowData(Row) = Fill
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditPosition.FillToGrid", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillFromGrid
'' Description: Fill in a Fill structure with the info from a row in the grid
'' Inputs:      Row
'' Returns:     Fill
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function FillFromGrid(ByVal Row As Long) As cPtFill
On Error GoTo ErrSection:

    Dim Fill As cPtFill                 ' Fill to return

    With fgFills
        If VarType(.RowData(Row)) = vbEmpty Then
            Set Fill = New cPtFill
            Fill.FillID = ValOfText(.TextMatrix(Row, FillCol(eGDFillCol_FillID)))
            Fill.OrderID = ValOfText(.TextMatrix(Row, FillCol(eGDFillCol_OrderID)))
            Fill.Symbol = .TextMatrix(Row, FillCol(eGDFillCol_Symbol))
            Fill.SymbolID = g.SymbolPool.SymbolIDforSymbol(Fill.Symbol)
            Fill.AccountID = m.lAccountID
            Fill.AutoTradingItemID = m.lAutoTradeItemID
        Else
            Set Fill = .RowData(Row)
        End If
        
        Fill.Buy = (UCase(.TextMatrix(Row, FillCol(eGDFillCol_Action))) = "BUY")
        Fill.Quantity = ValOfText(.TextMatrix(Row, FillCol(eGDFillCol_Quantity)))
        Fill.PosQuantity = ValOfText(.TextMatrix(Row, FillCol(eGDFillCol_PosQuantity)))
        Fill.Price = PriceFromDisplay(.TextMatrix(Row, FillCol(eGDFillCol_Price)), Trim(txtSymbol.Text))
        Fill.FillDate = DateOf(.TextMatrix(Row, FillCol(eGDFillCol_Date)))
        Fill.Fees = ValOfText(.TextMatrix(Row, FillCol(eGDFillCol_Fees)))
        Fill.BrokerID = .TextMatrix(Row, FillCol(eGDFillCol_BrokerID))
        
        If Fill.Buy Then
            Fill.Position = eTT_Position_Long
        Else
            Fill.Position = eTT_Position_Short
        End If
    End With
    
    Set FillFromGrid = Fill

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTTEditPosition.FillFromGrid", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateFill
'' Description: Update/Create the fill object and the corresponding order
'' Inputs:      Row to Update
'' Returns:     True on success, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function UpdateFill(ByVal Row As Long) As Boolean
On Error GoTo ErrSection:

    Dim bNew As Boolean
    Dim strPosition As String
    Dim lNetPosition As Long
    Dim strSymbol As String
    Dim lQuantity As Long
    Dim lPosQuantity As Long

    If (Row < fgFills.FixedRows) Or (Row >= fgFills.Rows) Then Exit Function

    m.bUpdatingFill = True
    m.bStartNew = False
    
    If Len(fgFills.TextMatrix(Row, FillCol(eGDFillCol_Action))) = 0 Then
        fgFills.Col = FillCol(eGDFillCol_Action)
        InfBox "Please enter in an Action for this fill", , , "Error"
        GoTo ErrExit:
    End If
    
    If PriceFromDisplay(fgFills.TextMatrix(Row, FillCol(eGDFillCol_Price)), Trim(txtSymbol.Text)) <= 0 Then
        fgFills.Col = FillCol(eGDFillCol_Price)
        InfBox "Please enter in a valid price for this fill", , , "Error"
        GoTo ErrExit:
    End If
    
    If ValOfText(fgFills.TextMatrix(Row, FillCol(eGDFillCol_PosQuantity))) <= 0 Then
        fgFills.Col = FillCol(eGDFillCol_PosQuantity)
        InfBox "Please enter in a valid quantity for this fill", , , "Error"
        GoTo ErrExit:
    End If
    
    lNetPosition = NetPosition
    strPosition = Trim(UCase(Parse(lblPosition.Caption, " ", 1)))
    If (strPosition = "SHORT" And lNetPosition > 0) Or (strPosition = "LONG" And lNetPosition < 0) Then
        lQuantity = CLng(ValOfText(fgFills.TextMatrix(Row, FillCol(eGDFillCol_PosQuantity))))
        fgFills.Col = FillCol(eGDFillCol_PosQuantity)
        If InfBox("The quantity on this fill will reverse the trade.  Would you like to start a new trade?", "?", "+Yes|-No", "Warning") = "Y" Then
            fgFills.TextMatrix(Row, FillCol(eGDFillCol_PosQuantity)) = Format(lQuantity - Abs(lNetPosition), "#,##0")
            fgFills.TextMatrix(Row, FillCol(eGDFillCol_Quantity)) = Str(lQuantity)
            m.bStartNew = True
        Else
            fgFills.TextMatrix(Row, FillCol(eGDFillCol_PosQuantity)) = Format(lQuantity - Abs(lNetPosition), "#,##0")
            fgFills.EditCell
            GoTo ErrExit:
        End If
    End If
    
    bNew = False
    If Len(fgFills.TextMatrix(Row, FillCol(eGDFillCol_FillID))) = 0 Then
        fgFills.TextMatrix(Row, FillCol(eGDFillCol_FillID)) = Str(NextFillID)
        bNew = True
    End If
    
    strSymbol = fgFills.TextMatrix(Row, FillCol(eGDFillCol_Symbol))
    If (Len(strSymbol) = 0) Or (InStr(strSymbol, "-0") <> 0) Then
        fgFills.TextMatrix(Row, FillCol(eGDFillCol_Symbol)) = RollSymbolForDate(Trim(txtSymbol.Text), DateOf(fgFills.TextMatrix(Row, FillCol(eGDFillCol_Date))))
    End If
    
    lQuantity = CLng(ValOfText(fgFills.TextMatrix(Row, FillCol(eGDFillCol_Quantity))))
    lPosQuantity = CLng(ValOfText(fgFills.TextMatrix(Row, FillCol(eGDFillCol_PosQuantity))))
    If lPosQuantity > lQuantity Then
        fgFills.TextMatrix(Row, FillCol(eGDFillCol_Quantity)) = Str(lPosQuantity)
    End If
     
    fgFills.RowData(Row) = FillFromGrid(Row)
    
    CalcProfit
    UpdateFill = True

ErrExit:
    m.bUpdatingFill = False
    Exit Function
    
ErrSection:
    m.bUpdatingFill = False
    RaiseError "frmTTEditPosition.UpdateFill", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BlankRow
'' Description: Determines whether the given row is "blank" or not
'' Inputs:      Row to Check
'' Returns:     True if no Action, Quantity, and Price, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function BlankRow(ByVal Row As Long) As Boolean
On Error GoTo ErrSection:

    BlankRow = False
    With fgFills
        If Row >= .FixedRows And Row < .Rows Then
            If Len(.TextMatrix(Row, FillCol(eGDFillCol_Action))) = 0 Then
                'If Len(.TextMatrix(Row, FillCol(eGDFillCol_PosQuantity))) = 0 Then
                    If Len(.TextMatrix(Row, FillCol(eGDFillCol_Price))) = 0 Then
                        BlankRow = True
                    End If
                'End If
            End If
        End If
    End With

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTTEditPosition.BlankRow", eGDRaiseError_Raise
    
End Function

Private Function SecType() As String
On Error GoTo ErrSection:

    Dim nSecType As eSYM_SecType
    
    With g.SymbolPool
        nSecType = .SecType(.PoolRecForSymbol(Trim(txtSymbol.Text)))
    End With
    
    Select Case nSecType
        Case eSYMType_Index
            m.strSecType = "I"
        Case eSYMType_Stock
            m.strSecType = "S"
        Case eSYMType_Future
            m.strSecType = "F"
        Case eSYMType_MutualFund
            m.strSecType = "M"
        Case eSYMType_Forex
            m.strSecType = "FX"
    End Select

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "ftmTTEditPosition.SecType", eGDRaiseError_Raise
    
End Function

Private Sub StartNew()
On Error GoTo ErrSection:

    Dim NewFill As New cPtFill
    Dim lRedraw As Long

    If BlankRow(fgFills.Rows - 1) Then
        Set NewFill = fgFills.RowData(fgFills.Rows - 2)
    Else
        Set NewFill = fgFills.RowData(fgFills.Rows - 1)
    End If
    
    Save
    ''frmTTPositions.RefreshTrade m.lActivityID
    
    With fgFills
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        .Rows = .FixedRows
        .Rows = .Rows + 1
        
        m.lActivityID = 0&
        m.lPositionID = 0&
        Caption = "New Trade"
        Set m.Position = Nothing
        NewFill.PosQuantity = NewFill.Quantity - NewFill.PosQuantity
        FillToGrid .Rows - 1, NewFill
        
        AddBlankRow
        SetPositionLabel
        CalcProfit
        EnableControls
        
        .RowSel = .FixedRows
        .Row = .RowSel
        .Col = FillCol(eGDFillCol_Action)
        
        AutoSize
        .Redraw = lRedraw
    End With
        
    m.bStartNew = False
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditPosition.StartNew", eGDRaiseError_Raise
    
End Sub

Private Sub CalcProfit()
On Error GoTo ErrSection:
    
    Dim strSymbol As String
    Dim lIndex As Long
    Dim Bars As New cGdBars
    Dim dProfit As Double
    Dim dPrice As Double
    Dim dBuyPrice As Double
    Dim dSellPrice As Double
    Dim lQuantity As Long
    Dim lBuyQuantity As Long
    Dim lSellQuantity As Long
    Dim dComms As Double
    Dim dTV As Double
    Dim dTM As Double
    Dim dLastPrice As Double
    
    strSymbol = Trim(txtSymbol.Text)
    If DM_GetBars(Bars, strSymbol, , LastDailyDownload) Then
        dTV = Bars.Prop(eBARS_TickValue)
        dTM = Bars.Prop(eBARS_TickMove)
        dLastPrice = Bars(eBARS_Close, Bars.Size - 1)
        If dLastPrice = gdNullValue(Bars.ArrayHandle(eBARS_Close)) Then dLastPrice = 0
        
        With fgFills
            For lIndex = .FixedRows To .Rows - 1
                If Not BlankRow(lIndex) Then
                    dPrice = PriceFromDisplay(.TextMatrix(lIndex, FillCol(eGDFillCol_Price)), strSymbol)
                    lQuantity = CLng(Val(.TextMatrix(lIndex, FillCol(eGDFillCol_PosQuantity))))
                    
                    If UCase(.TextMatrix(lIndex, FillCol(eGDFillCol_Action))) = "BUY" Then
                        dBuyPrice = dBuyPrice + (dPrice * lQuantity)
                        lBuyQuantity = lBuyQuantity + lQuantity
                    ElseIf UCase(.TextMatrix(lIndex, FillCol(eGDFillCol_Action))) = "SELL" Then
                        dSellPrice = dSellPrice + (dPrice * lQuantity)
                        lSellQuantity = lSellQuantity + lQuantity
                    End If
                    
                    dComms = dComms + ValOfText(.TextMatrix(lIndex, FillCol(eGDFillCol_Fees)))
                End If
            Next lIndex
        End With
        
        If lBuyQuantity <> lSellQuantity Then
            lblProfit = "Open Profit"
        Else
            lblProfit = "Closed Profit"
        End If
                
        If UCase(Parse(lblPosition.Caption, " ", 1)) = "LONG" Then
            dSellPrice = dSellPrice + (dLastPrice * (lBuyQuantity - lSellQuantity))
            lSellQuantity = lBuyQuantity
            dBuyPrice = dBuyPrice / lBuyQuantity
            dSellPrice = dSellPrice / lSellQuantity
            
            dProfit = ((dSellPrice - dBuyPrice) * dTV) / dTM * lBuyQuantity
        Else
            dBuyPrice = dBuyPrice + (dLastPrice * (lSellQuantity - lBuyQuantity))
            lBuyQuantity = lSellQuantity
            dBuyPrice = dBuyPrice / lBuyQuantity
            dSellPrice = dSellPrice / lSellQuantity
            
            dProfit = (-(dBuyPrice - dSellPrice) * dTV) / dTM * lSellQuantity
        End If
        dProfit = dProfit - dComms
        
        If dProfit >= 0 Then
            txtProfit.Text = Format(dProfit, "+$#,##0.00")
            txtProfit.ForeColor = QBColor(2)
        Else
            txtProfit.Text = Format(dProfit, "$#,##0.00")
            txtProfit.ForeColor = vbRed
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditPosition.CalcProfit", eGDRaiseError_Raise
    
End Sub

Private Sub AutoSize()
On Error GoTo ErrSection:

    Dim lRedraw As Long

    With fgFills
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        If Left(.TextMatrix(.Rows - 1, FillCol(eGDFillCol_PosQuantity)), 2) = " <" Then
            .TextMatrix(.Rows - 1, FillCol(eGDFillCol_PosQuantity)) = " <"
        End If
        .AutoSize 0, .Cols - 1, False, 75
        .ColWidth(FillCol(eGDFillCol_Date)) = .ColWidth(FillCol(eGDFillCol_Date)) + 100
        If .TextMatrix(.Rows - 1, FillCol(eGDFillCol_PosQuantity)) = " <" Then
            .TextMatrix(.Rows - 1, FillCol(eGDFillCol_PosQuantity)) = " <-- Set Action to Buy or Sell for New Fill "
        End If
        
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditPosition.AutoSize", eGDRaiseError_Raise
    
End Sub

Private Function OppositeDirection(ByVal lRow As Long) As Boolean
On Error GoTo ErrSection:

    Dim strPosition As String
    
    strPosition = Parse(lblPosition.Caption, " ", 1)
    
    Select Case UCase(strPosition)
        Case "NONE"
            OppositeDirection = False
        Case "LONG"
            OppositeDirection = (UCase(fgFills.TextMatrix(lRow, FillCol(eGDFillCol_Action))) = "SELL")
        Case "SHORT"
            OppositeDirection = (UCase(fgFills.TextMatrix(lRow, FillCol(eGDFillCol_Action))) = "BUY")
    End Select

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTTEditPosition.OppositeDirection", eGDRaiseError_Raise
    
End Function

Private Sub ChangeSymbols()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lRedraw As Long                 ' Current state of the grid's redraw
    Dim strSymbol As String             ' Symbol for the current row
    
    With fgFills
        lRedraw = .Redraw
        .Redraw = flexRDNone
                
        For lIndex = .FixedRows To .Rows - 1
            If Not BlankRow(lIndex) Then
                strSymbol = RollSymbolForDate(Trim(txtSymbol.Text), DateOf(fgFills.TextMatrix(lIndex, FillCol(eGDFillCol_Date))))
                .TextMatrix(lIndex, FillCol(eGDFillCol_Symbol)) = strSymbol
            End If
        Next lIndex
        
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditPosition.ChangeSymbols", eGDRaiseError_Raise
    
End Sub
