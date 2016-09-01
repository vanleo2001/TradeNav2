VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmTTSelectPosition 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraPriceDisplay 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   795
      Left            =   3300
      TabIndex        =   5
      Top             =   1500
      Width           =   1275
      Begin VB.OptionButton optTradingUnits 
         Caption         =   "Trading Units"
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   1275
      End
      Begin VB.OptionButton optDecimal 
         Caption         =   "Decimal"
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label lblUnits 
         Caption         =   "Display Prices in:"
         Height          =   255
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   1275
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid fgPositions 
      Height          =   2895
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2475
      _cx             =   4366
      _cy             =   5106
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
   Begin VB.Frame fraButtons 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1215
      Left            =   3300
      TabIndex        =   0
      Top             =   120
      Width           =   1275
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New"
         Height          =   375
         Left            =   0
         TabIndex        =   3
         Top             =   420
         Width           =   1275
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   0
         TabIndex        =   2
         Top             =   840
         Width           =   1275
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   375
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   1275
      End
   End
End
Attribute VB_Name = "frmTTSelectPosition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum eGDCols
    eGDCol_PositionID = 0
    eGDCol_Symbol
    eGDCol_Date
    eGDCol_Position
    eGDCol_NumCols
End Enum

Private Enum eGDModes
    eGDMode_FromSymbol = 0
    eGDMode_FromArray
    eGDMode_OrderFromSymbol
    eGDMode_Exchanges
    eGDMode_Contracts
    eGDMode_DOM
    eGDMode_Subscribe
End Enum

Private Type mPrivate
    strSymbol As String
    lAccountID As Long
    astrTradeIDs As cGdArray
    Mode As eGDModes

    iOK As Integer
End Type
Private m As mPrivate

Private Function GDCol(ByVal Col As eGDCols) As Long
    GDCol = Col
End Function

Private Sub cmdCancel_Click()
On Error GoTo ErrSection:

    m.iOK = 0
    Me.Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSelectPosition.cmdCancel.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub cmdNew_Click()
On Error GoTo ErrSection:

    m.iOK = 2
    Me.Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSelectPosition.cmdNew.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub cmdOK_Click()
On Error GoTo ErrSection:

    Dim astrSymbols As New cGdArray
    Dim strSymbol As String

    If m.Mode = eGDMode_Subscribe Then
        With fgPositions
            If .RowSel >= .FixedRows And .RowSel < .Rows Then
                strSymbol = .TextMatrix(.RowSel, 0) & " " & .TextMatrix(.RowSel, 1) & " " & .TextMatrix(.RowSel, 2)
                If Len(.TextMatrix(.RowSel, 3)) = 0 Then
                    Set astrSymbols = frmSymbolSelector.ShowMe("", False, False, "Please select Genesis Symol for " & strSymbol, False)
                    If astrSymbols.Size = 0 Then
                        Err.Raise vbObjectError + 1000, , "Please select a genesis symbol for " & strSymbol
                    Else
                        .TextMatrix(.RowSel, 3) = Trim(UCase(astrSymbols(0)))
                        fgPositions_AfterEdit .RowSel, 3
                    End If
                End If
            Else
                Err.Raise vbObjectError + 1000, , "Please select a symbol"
            End If
        End With
    End If

    m.iOK = 1
    Me.Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSelectPosition.cmdOK.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fgPositions_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    Dim strExchange As String           ' PATS Exchange
    Dim strSymbol As String             ' PATS Symbol
    Dim strContract As String           ' PATS Contract
    Dim strGenSym As String             ' Genesis Symbol
    Dim lIndex As Long                  ' Index for a for loop
    Dim lRecord As Long                 ' Record in the table to modify

    If m.Mode = eGDMode_Contracts Or m.Mode = eGDMode_Subscribe Then
#If 0 Then
        If Col = 3 Then
            With fgPositions
                .TextMatrix(Row, 3) = UCase(Trim(.TextMatrix(Row, 3)))
            
                strExchange = .TextMatrix(Row, 0)
                strSymbol = .TextMatrix(Row, 1)
                strContract = .TextMatrix(Row, 2)
                strGenSym = .TextMatrix(Row, 3)
                
                lRecord = -1&
                For lIndex = 0 To g.Pats.Symbols.NumRecords - 1
                    If g.Pats.Symbols(1, lIndex) = strExchange Then
                        If g.Pats.Symbols(2, lIndex) = strSymbol Then
                            If g.Pats.Symbols(3, lIndex) = strContract Then
                                lRecord = lIndex
                            End If
                        End If
                    End If
                Next lIndex
            
                If lRecord = -1& Then
                    If Len(strGenSym) > 0 Then
                        g.Pats.Symbols.AddRecord strGenSym & vbTab & strExchange & vbTab & strSymbol & vbTab & strContract, , vbTab
                    End If
                Else
                    If Len(strGenSym) > 0 Then
                        g.Pats.Symbols(0, lRecord) = strGenSym
                    Else
                        g.Pats.Symbols.RemoveRecords lRecord
                    End If
                End If
            End With
        End If
#End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSelectPosition.fgPositions.AfterEdit", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub fgPositions_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    If m.Mode = eGDMode_Contracts Or m.Mode = eGDMode_Subscribe Then
        If NewCol = 3 Then fgPositions.EditCell
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSelectPosition.fgPositions.AfterRowColChange", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub fgPositions_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    If m.Mode = eGDMode_Contracts Or m.Mode = eGDMode_Subscribe Then
        If Col <> 3 Then
            Cancel = True
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSelectPosition.fgPositions.BeforeEdit", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub Form_Activate()
On Error GoTo ErrSection:

    fraPriceDisplay.Visible = (m.Mode = eGDMode_DOM)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSelectPosition.Form.Activate", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub Form_Load()
On Error GoTo ErrSection:

    Icon = Picture16(ToolbarIcon("ID_TradeTracker"))
    Width = 6400
    CenterTheForm Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSelectPosition.Form.Load", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode <> vbFormCode Then
        If m.Mode = eGDMode_FromSymbol Then
            m.iOK = 1
        Else
            m.iOK = 0
        End If
        Cancel = True
        Me.Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSelectPosition.Form.QueryUnload", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub Form_Resize()
On Error Resume Next

    Dim lMinWidth As Long
    Dim lMinHeight As Long
    
    lMinWidth = fraButtons.Width * 3
    lMinHeight = fraButtons.Height + fraPriceDisplay.Height + (fraButtons.Top * 3)
    If LimitFormSize(Me, lMinWidth, lMinHeight) Then Exit Sub
    
    With fraButtons
        .Move ScaleWidth - .Width - fgPositions.Left, fgPositions.Top
    End With
    
    With fraPriceDisplay
        .Move fraButtons.Left, fraButtons.Height + (fraButtons.Top * 2)
    End With
    
    With fgPositions
        .Move .Left, .Top, fraButtons.Left - (.Left * 2), ScaleHeight - (.Top * 2)
    End With

End Sub

Public Function ShowMe(ByVal strSymbol As String, ByVal lAccountID As Long, Optional ByVal bOrders As Boolean = False) As Long
On Error GoTo ErrSection:

    m.strSymbol = strSymbol
    m.lAccountID = lAccountID
    m.iOK = 2
    
    If bOrders Then
        m.Mode = eGDMode_OrderFromSymbol
        Caption = "Select Order for Fill..."
    Else
        m.Mode = eGDMode_FromSymbol
        Caption = "Select Position for Fill..."
    End If
    
    ' For now, hide the cancel button since it really does not make sense...
    cmdCancel.Visible = False

    InitGrid
    LoadGrid

    If fgPositions.Rows > fgPositions.FixedRows Then
        ShowForm Me, True
    End If
    
    Select Case m.iOK
        Case 0
            ShowMe = -1
        Case 1
            ShowMe = CLng(ValOfText(fgPositions.TextMatrix(fgPositions.RowSel, GDCol(eGDCol_PositionID))))
        Case 2
            ShowMe = 0
    End Select

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmTTSelectPosition.ShowMe", eGDRaiseError_Raise
    
End Function

Private Sub InitGrid()
On Error GoTo ErrSection:

    Dim lRedraw As Long
    
    With fgPositions
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        .AllowBigSelection = False
        .AllowSelection = False
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .Editable = flexEDNone
        .ExplorerBar = flexExSortShow ' = flexExNone
        .ExtendLastCol = True
        .ScrollTrack = True
        .SelectionMode = flexSelectionListBox
        .SheetBorder = RGB(128, 128, 128)
        
        .Rows = 1
        .FixedRows = 1
        .Cols = GDCol(eGDCol_NumCols)
        .FixedCols = 0
        
        .TextMatrix(0, GDCol(eGDCol_PositionID)) = "ID"
        Select Case m.Mode
            Case eGDMode_Exchanges
                .TextMatrix(0, GDCol(eGDCol_Symbol)) = "Exchange"
            Case Else
                .TextMatrix(0, GDCol(eGDCol_Symbol)) = "Symbol"
        End Select
        .TextMatrix(0, GDCol(eGDCol_Date)) = "Date"
        Select Case m.Mode
            Case eGDMode_FromArray, eGDMode_FromSymbol
                .TextMatrix(0, GDCol(eGDCol_Position)) = "Position"
            Case eGDMode_OrderFromSymbol
                .TextMatrix(0, GDCol(eGDCol_Position)) = "Status"
        End Select
        
        If m.Mode = eGDMode_Exchanges Then
            .ColHidden(GDCol(eGDCol_PositionID)) = True
            .ColHidden(GDCol(eGDCol_Position)) = True
            .ColHidden(GDCol(eGDCol_Date)) = True
        End If
        
        .ColDataType(GDCol(eGDCol_Date)) = flexDTDate
        .ColAlignment(GDCol(eGDCol_Date)) = flexAlignCenterTop
        .ColFormat(GDCol(eGDCol_Date)) = DateAndTime("Format")
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignLeftTop
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSelectPosition.InitGrid", eGDRaiseError_Raise
    
End Sub

Private Sub LoadGrid()
On Error GoTo ErrSection:

    Dim lRedraw As Long                 ' Current state of the grid's redraw
    Dim lIndex As Long                  ' Index into a for loop
    Dim Position As New cPtPosition     ' Temporary Position object
    Dim lNet As Long                    ' Quantity open in the position
    Dim rs As Recordset                 ' Recordset into the database
    
    With fgPositions
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        Select Case m.Mode
            Case eGDMode_FromSymbol
                .Rows = .FixedRows
                Set rs = g.dbPaper.OpenRecordset("SELECT tblActivities.AccountID, tblPositions.* " & _
                        "FROM tblActivities INNER JOIN tblPositions ON tblActivities.ActivityID = tblPositions.ActivityID " & _
                        "WHERE tblActivities.AccountID=" & Str(m.lAccountID) & " " & _
                        "AND tblPositions.Open=True AND tblPositions.OpenSymbol='" & m.strSymbol & "';", dbOpenDynaset)
                Do While Not rs.EOF
                    .Rows = .Rows + 1
                    
                    Set Position = New cPtPosition
                    If Position.Load(rs!PositionID) Then
                        .TextMatrix(.Rows - 1, GDCol(eGDCol_PositionID)) = Str(Position.PositionID)
                        .TextMatrix(.Rows - 1, GDCol(eGDCol_Symbol)) = Position.OpenSymbol
                        .TextMatrix(.Rows - 1, GDCol(eGDCol_Date)) = Position.EntryDate
                        lNet = Position.EntryQuantity - Position.ExitQuantity
                        Select Case Position.Position
                            Case eTT_Position_Long
                                .TextMatrix(.Rows - 1, GDCol(eGDCol_Position)) = "Long " & Format(lNet, "#,##0")
                            Case eTT_Position_Short
                                .TextMatrix(.Rows - 1, GDCol(eGDCol_Position)) = "Short " & Format(lNet, "#,##0")
                        End Select
                    End If
                    
                    rs.MoveNext
                Loop
            
            Case eGDMode_OrderFromSymbol
                .Rows = .FixedRows
                Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrders] " & _
                        "WHERE [Symbol]='" & m.strSymbol & "' " & _
                        "AND [Status]<>" & Str(eTT_OrderStatus_Filled) & ";", dbOpenDynaset)
                Do While Not rs.EOF
                    .Rows = .Rows + 1
                    
                    .TextMatrix(.Rows - 1, GDCol(eGDCol_PositionID)) = Str(rs!OrderID)
                    .TextMatrix(.Rows - 1, GDCol(eGDCol_Symbol)) = m.strSymbol
                    .TextMatrix(.Rows - 1, GDCol(eGDCol_Date)) = rs!OrderDate
                    .TextMatrix(.Rows - 1, GDCol(eGDCol_Position)) = OrderStatus(rs!Status)
                    
                    rs.MoveNext
                Loop
                
        End Select
        
        If .Rows > .FixedRows Then
            .Row = .FixedRows
            .RowSel = .FixedRows
        End If
        
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTTSelectPosition.LoadGrid", eGDRaiseError_Raise
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    Set m.astrTradeIDs = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSelectPosition.Form.Unload", eGDRaiseError_Show
    Resume ErrExit

End Sub

Public Function ShowMeFromArray(ByVal astrTradeIDs As cGdArray) As Integer
On Error GoTo ErrSection:

    Set m.astrTradeIDs = astrTradeIDs
    Caption = "Select Trade to Edit..."
    cmdNew.Visible = False
    cmdCancel.Top = cmdNew.Top
    m.Mode = eGDMode_FromArray
    
    InitGrid
    LoadGridFromArray
    
    If fgPositions.Rows > fgPositions.FixedRows Then
        ShowForm Me, True
    End If
    
    Select Case m.iOK
        Case 0
            ShowMeFromArray = -1
        Case 1
            ShowMeFromArray = CLng(ValOfText(fgPositions.TextMatrix(fgPositions.RowSel, GDCol(eGDCol_PositionID))))
        Case 2
            ShowMeFromArray = 0
    End Select

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmTTSelectPosition.ShowMeFromArray", eGDRaiseError_Raise
    
End Function

Private Sub LoadGridFromArray()
On Error GoTo ErrSection:

    Dim lRedraw As Long                 ' Current state of the grid's redraw
    Dim lIndex As Long                  ' Index into a for loop
    Dim Position As New cPtPosition     ' Temporary Position object
    Dim lNet As Long                    ' Quantity open in the position
    Dim strPosition As String           ' Position to display to user
    
    With fgPositions
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        .Rows = .FixedRows
        For lIndex = 0 To m.astrTradeIDs.Size - 1
            Set Position = New cPtPosition
            If Position.Load(CLng(m.astrTradeIDs(lIndex))) Then
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, GDCol(eGDCol_PositionID)) = Str(Position.PositionID)
                .TextMatrix(.Rows - 1, GDCol(eGDCol_Symbol)) = Position.OpenSymbol
                .TextMatrix(.Rows - 1, GDCol(eGDCol_Date)) = Position.EntryDate
                lNet = Position.EntryQuantity - Position.ExitQuantity
                Select Case Position.Position
                    Case eTT_Position_Long
                        strPosition = "Long "
                    Case eTT_Position_Short
                        strPosition = "Short "
                End Select
                If lNet > 0 Then strPosition = strPosition & Format(lNet, "#,##0")
                .TextMatrix(.Rows - 1, GDCol(eGDCol_Position)) = strPosition
            End If
        Next lIndex
        
        If .Rows > .FixedRows Then
            .Row = .FixedRows
            .RowSel = .FixedRows
        End If
        
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSelectPosition.LoadGridFromArray", eGDRaiseError_Raise
    
End Sub

Public Function ShowExchanges(tblData As cGdTable) As Boolean
On Error GoTo ErrSection:

    Dim lRedraw As Long
    Dim lIndex As Long

    cmdNew.Visible = False
    cmdCancel.Top = cmdNew.Top
    m.Mode = eGDMode_Exchanges
    Caption = "Available Exchanges"
    
    With fgPositions
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        .AllowBigSelection = False
        .AllowSelection = False
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .Editable = flexEDNone
        .ExplorerBar = flexExSortShow ' = flexExNone
        .ExtendLastCol = True
        .ScrollTrack = True
        .SelectionMode = flexSelectionListBox
        .SheetBorder = RGB(128, 128, 128)
        
        .Rows = 1
        .FixedRows = 1
        .Cols = 3
        .FixedCols = 0
        
        .TextMatrix(0, 0) = "Exchange"
        .TextMatrix(0, 1) = "Query"
        .TextMatrix(0, 2) = "Amend"
        
        .ColDataType(1) = flexDTBoolean
        .ColDataType(2) = flexDTBoolean
        
        .Rows = .FixedRows
        For lIndex = 0 To tblData.NumRecords - 1
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = tblData(tblData.FieldNum("ExchangeName"), lIndex)
            CheckedCell(fgPositions, .Rows - 1, 1) = (tblData(tblData.FieldNum("QueryEnabled"), lIndex) = "Y")
            CheckedCell(fgPositions, .Rows - 1, 2) = (tblData(tblData.FieldNum("AmendEnabled"), lIndex) = "Y")
        Next lIndex
        
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = lRedraw
    End With
    
    ShowForm Me, True
    
    ShowExchanges = (m.iOK = 1)

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmTTSelectPosition.ShowExchanges", eGDRaiseError_Raise
    
End Function

Public Function ShowContracts(tblData As cGdTable, ByVal bSubscribe As Boolean) As String
On Error GoTo ErrSection:

    Dim lRedraw As Long
    Dim lIndex As Long
    Dim strSymbol As String

    cmdNew.Visible = False
    cmdCancel.Top = cmdNew.Top
    If bSubscribe Then
        m.Mode = eGDMode_Subscribe
    Else
        m.Mode = eGDMode_Contracts
    End If
    Caption = "Available Contracts"
    
    With fgPositions
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        .AllowBigSelection = False
        .AllowSelection = False
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .Editable = flexEDKbdMouse
        .ExplorerBar = flexExSortShow ' = flexExNone
        .ExtendLastCol = True
        .ScrollTrack = True
        .SelectionMode = flexSelectionListBox
        .SheetBorder = RGB(128, 128, 128)
        
        .Rows = 1
        .FixedRows = 1
        .Cols = 4
        .FixedCols = 0
        
        .TextMatrix(0, 0) = "Exchange"
        .TextMatrix(0, 1) = "Security"
        .TextMatrix(0, 2) = "Contract"
        .TextMatrix(0, 3) = "Gen Symbol"
        
        .Rows = .FixedRows
        For lIndex = 0 To tblData.NumRecords - 1
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = tblData(tblData.FieldNum("ExchangeName"), lIndex)
            .TextMatrix(.Rows - 1, 1) = tblData(tblData.FieldNum("SecurityName"), lIndex)
            .TextMatrix(.Rows - 1, 2) = tblData(tblData.FieldNum("ContractDate"), lIndex)
            
            strSymbol = ""
            ''g.Pats.TranslateFromSymbol .TextMatrix(.Rows - 1, 0), .TextMatrix(.Rows - 1, 1), .TextMatrix(.Rows - 1, 2), strSymbol
            .TextMatrix(.Rows - 1, 3) = strSymbol
        Next lIndex
        
        If .Rows > .FixedRows Then
            .Row = .FixedRows
            .RowSel = .FixedRows
        End If
        
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = lRedraw
    End With
    
    ShowForm Me, True
    
    If m.iOK = 1 Then
        If m.Mode = eGDMode_Subscribe Then
            ShowContracts = fgPositions.TextMatrix(fgPositions.RowSel, 3)
        End If
    End If

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmTTSelectPosition.ShowContracts", eGDRaiseError_Raise
    
End Function

Public Function ShowDOM(ByVal strSymbol As String, tblData As cGdTable) As Boolean
On Error GoTo ErrSection:

    Dim lRedraw As Long
    Dim lIndex As Long
    Dim lRow As Long

    cmdNew.Visible = False
    cmdCancel.Top = cmdNew.Top
    m.Mode = eGDMode_DOM
    m.strSymbol = strSymbol
    Caption = "Depth of Market for " & strSymbol
    
    With fgPositions
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        .Clear
        .AllowBigSelection = False
        .AllowSelection = False
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .Editable = flexEDNone
        .ExplorerBar = flexExSortShow ' = flexExNone
        .ExtendLastCol = True
        .ScrollTrack = True
        .SelectionMode = flexSelectionListBox
        .SheetBorder = RGB(128, 128, 128)
        
        .Rows = 2
        .FixedRows = 2
        .Cols = 6
        .FixedCols = 0
        
        .Cell(flexcpText, 0, 0, 0, 2) = "Bid"
        .Cell(flexcpText, 0, 3, 0, 5) = "Ask"
        .TextMatrix(1, 0) = "Price"
        .TextMatrix(1, 1) = "Size"
        .TextMatrix(1, 2) = "Time"
        .TextMatrix(1, 3) = "Price"
        .TextMatrix(1, 4) = "Size"
        .TextMatrix(1, 5) = "Time"
        .MergeCells = flexMergeFixedOnly
        .MergeRow(0) = True
        
        .Rows = 20 + .FixedRows
        For lIndex = 0 To tblData.NumRecords - 1
            If UCase(Left(tblData(tblData.FieldNum("PriceType"), lIndex), 6)) = "BIDDOM" Then
                lRow = CLng(Val(Right(tblData(tblData.FieldNum("PriceType"), lIndex), 2))) + .FixedRows
                .TextMatrix(lRow, 0) = PriceDisplay(tblData(tblData.FieldNum("Price"), lIndex), strSymbol, optTradingUnits)
                .TextMatrix(lRow, 1) = tblData(tblData.FieldNum("Volume"), lIndex)
                .TextMatrix(lRow, 2) = Format(Val(tblData(tblData.FieldNum("Hour"), lIndex)), "00") & ":" & Format(Val(tblData(tblData.FieldNum("Minute"), lIndex)), "00") & ":" & Format(Val(tblData(tblData.FieldNum("Second"), lIndex)), "00")
            ElseIf UCase(Left(tblData(tblData.FieldNum("PriceType"), lIndex), 8)) = "OFFERDOM" Then
                lRow = CLng(Val(Right(tblData(tblData.FieldNum("PriceType"), lIndex), 2))) + .FixedRows
                .TextMatrix(lRow, 3) = PriceDisplay(tblData(tblData.FieldNum("Price"), lIndex), strSymbol, optTradingUnits)
                .TextMatrix(lRow, 4) = tblData(tblData.FieldNum("Volume"), lIndex)
                .TextMatrix(lRow, 5) = Format(Val(tblData(tblData.FieldNum("Hour"), lIndex)), "00") & ":" & Format(Val(tblData(tblData.FieldNum("Minute"), lIndex)), "00") & ":" & Format(Val(tblData(tblData.FieldNum("Second"), lIndex)), "00")
            End If
        Next lIndex
        
        .ColAlignment(2) = flexAlignCenterTop
        .ColAlignment(5) = flexAlignCenterTop
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterTop
        .Cell(flexcpAlignment, 1, 0, 1, .Cols - 1) = flexAlignLeftTop
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = lRedraw
    End With
    
    ShowForm Me, True
    
    ShowDOM = (m.iOK = 1)

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmTTSelectPosition.ShowContracts", eGDRaiseError_Raise
    
End Function

Private Sub RedisplayPrices()
On Error GoTo ErrSection:

    Dim lRedraw As Long
    Dim lIndex As Long
    Dim dValue As Double

    Select Case m.Mode
        Case eGDMode_DOM
            With fgPositions
                lRedraw = .Redraw
                .Redraw = flexRDNone
                
                For lIndex = .FixedRows To .Rows - 1
                    dValue = PriceFromDisplay(.TextMatrix(lIndex, 0), m.strSymbol)
                    .TextMatrix(lIndex, 0) = PriceDisplay(dValue, m.strSymbol, optTradingUnits)
                    dValue = PriceFromDisplay(.TextMatrix(lIndex, 3), m.strSymbol)
                    .TextMatrix(lIndex, 3) = PriceDisplay(dValue, m.strSymbol, optTradingUnits)
                Next lIndex
                
                .Redraw = lRedraw
            End With
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSelectPosition.RedisplayPrices", eGDRaiseError_Raise
    
End Sub

Private Sub optDecimal_Click()
On Error GoTo ErrSection:

    RedisplayPrices
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSelectPosition.optDecimal.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub optTradingUnits_Click()
On Error GoTo ErrSection:

    RedisplayPrices
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSelectPosition.optTradingUnits.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub
