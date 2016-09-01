VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmOrderTracker 
   Caption         =   "Form1"
   ClientHeight    =   3585
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   5370
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   239
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   358
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniButtonImageXP cmdAlert 
      Height          =   315
      Left            =   1080
      TabIndex        =   2
      Top             =   60
      Width           =   855
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
      Caption         =   "frmOrderTracker.frx":0000
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmOrderTracker.frx":002A
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmOrderTracker.frx":004A
      RightToLeft     =   0   'False
   End
   Begin VSFlex7LCtl.VSFlexGrid fgReports 
      Height          =   915
      Left            =   60
      TabIndex        =   1
      Top             =   660
      Visible         =   0   'False
      Width           =   1935
      _cx             =   3413
      _cy             =   1614
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
   Begin HexUniControls.ctlUniButtonImageXP cmdViewReport 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   855
      _ExtentX        =   0
      _ExtentY        =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmOrderTracker.frx":0066
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmOrderTracker.frx":0094
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmOrderTracker.frx":00B4
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblNone 
      Height          =   195
      Left            =   4200
      Top             =   120
      Width           =   435
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
      Caption         =   "frmOrderTracker.frx":00D0
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmOrderTracker.frx":00F8
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmOrderTracker.frx":0118
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP Label2 
      Height          =   195
      Left            =   3240
      Top             =   120
      Width           =   555
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
      Caption         =   "frmOrderTracker.frx":0134
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmOrderTracker.frx":0160
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmOrderTracker.frx":0180
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP Label1 
      Height          =   195
      Left            =   2460
      Top             =   120
      Width           =   375
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
      Caption         =   "frmOrderTracker.frx":019C
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmOrderTracker.frx":01C2
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmOrderTracker.frx":01E2
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblNoOrders 
      Height          =   1815
      Left            =   120
      Top             =   1680
      Width           =   1815
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
      Caption         =   "frmOrderTracker.frx":01FE
      BackColor       =   16777215
      ForeColor       =   4210752
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   1
      AutoSize        =   0   'False
      Tip             =   "frmOrderTracker.frx":028E
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmOrderTracker.frx":02AE
      RightToLeft     =   0   'False
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgNone 
      Height          =   240
      Left            =   3900
      Picture         =   "frmOrderTracker.frx":02CA
      Top             =   90
      Width           =   240
   End
   Begin VB.Image imgBlack 
      Height          =   240
      Left            =   2940
      Picture         =   "frmOrderTracker.frx":0414
      Top             =   90
      Width           =   240
   End
   Begin VB.Image imgRed 
      Height          =   240
      Left            =   2160
      Picture         =   "frmOrderTracker.frx":055E
      Top             =   90
      Width           =   240
   End
   Begin VB.Menu mnuSettings 
      Caption         =   "Settings"
      Begin VB.Menu mnuShowReport 
         Caption         =   "Display Orders for Selected Row"
      End
      Begin VB.Menu mnuDisplayMult 
         Caption         =   "Display Orders for All"
      End
      Begin VB.Menu mnuChangeFont 
         Caption         =   "Change Font"
      End
   End
End
Attribute VB_Name = "frmOrderTracker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmOrderTracker.frm
'' Description: Allow the user to see the systems that have generated next
''              bar reports
''
'' Author:      Genesis Financial Data Services
''              425 E Woodmen Rd
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    lRow As Long
End Type
Private m As mPrivate

Private Enum eGDCols
    eGDCol_Viewed = 0
    eGDCol_Symbol
    eGDCol_SystemName
    eGDCol_Hwnd
    eGDCol_CRCViewed
    eGDCol_Orders
    eGDCol_LastKnown
    eGDCol_NumCols
End Enum

Private Function GDCol(ByVal lColumn As eGDCols) As Long
    GDCol = lColumn
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdAlert_Click
'' Description: Allow the user to configure an alert based on new orders
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdAlert_Click()
On Error GoTo ErrSection:

    frmOrderTrackerCfg.ShowMe

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderTracker.cmdAlert.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdViewReport_Click
'' Description: Allow the user to view the appropriate next bar report
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdViewReport_Click()
On Error GoTo ErrSection:

    ShowMultiple

#If 0 Then
    Dim lRow As Long                    ' Current row selected in the grid

    With fgReports
        If .Rows = .FixedRows + 1 Then
            lRow = .FixedRows
        Else
            lRow = .RowSel
        End If
        If lRow >= .FixedRows And lRow < .Rows Then
            ShowReport lRow
        End If
    End With
#End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "OrderTracker.cmdViewReport.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgReports_DblClick
'' Description: If the user double clicks, show the reports for that row
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgReports_DblClick()
On Error GoTo ErrSection:

    Dim lRow As Long                    ' Current mouse row in the grid
        
    With fgReports
        lRow = .MouseRow
        If lRow >= .FixedRows And lRow < .Rows Then
            If CheckedCell(fgReports, lRow, GDCol(eGDCol_Orders)) Then
                ShowReport lRow
            End If
        End If
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "OrderTracker.fgReports.DblClick", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fgReports_KeyDown(KeyCode As Integer, Shift As Integer)

    If fgKeyDown(KeyCode, Shift) Then Exit Sub

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgReports_KeyPress
'' Description: If the user hits Enter, show the reports for that row
'' Inputs:      Key Pressed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgReports_KeyPress(KeyAscii As Integer)
On Error GoTo ErrSection:

    Dim lRow As Long                    ' Current row in the grid
    
    If KeyAscii = vbKeyReturn Then
        With fgReports
            lRow = .Row
            If lRow >= .FixedRows And lRow < .Rows Then
                If CheckedCell(fgReports, lRow, GDCol(eGDCol_Orders)) Then
                    ShowReport lRow
                End If
            End If
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "OrderTracker.fgReports.KeyPress", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fgReports_LostFocus()
On Error Resume Next

    fgReports.Col = GDCol(eGDCol_Symbol)
    fgReports.ColSel = GDCol(eGDCol_SystemName)

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgReports_MouseDown
'' Description: When the user right clicks on the grid, show the Pop-Up menu
'' Inputs:      Button Pressed, Shift/Ctrl/Alt status, Location of Click
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgReports_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    Dim lRow As Long                    ' Current mouse row in the grid

    If Button = vbRightButton Then
        With fgReports
            lRow = .MouseRow
            If lRow >= .FixedRows And lRow < .Rows Then
                m.lRow = lRow
                .Row = lRow
                .RowSel = lRow
                mnuShowReport.Enabled = True
            Else
                mnuShowReport.Enabled = False
            End If
            
            PopupMenu mnuSettings
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "OrderTracker.fgReports.MouseDown", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fgReports_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
            
    If fgReports.MouseCol = GDCol(eGDCol_Viewed) Then
        GridTooltip fgReports, , "View Status"
    Else
        GridTooltip fgReports
    End If
    
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
    RaiseError "frmOrderTracker.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Place the form and do some initialization
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Dim strFont As String               ' Grid Font from the INI file

    g.Styler.StyleForm Me
    
    mnuSettings.Visible = False
    Me.Icon = Picture16(ToolbarIcon("ID_Orders"))
    
    Caption = "Order Tracker"
    CenterTheForm Me
    
    strFont = GetIniFileProperty("OrderTracker", "", "Fonts", g.strIniFile)
    If strFont <> "" Then
        FontFromString fgReports.Font, strFont
    End If
    
    InitGrid

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "OrderTracker.Form.Load", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: Set the state of the toolbar icon upon exiting
'' Inputs:      Whether to cancel unload, Mode of the unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode = 0 Then
        frmMain.tbToolbar.Tools("ID_Orders").State = ssUnchecked
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderTracker.Form.QueryUnload", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Resize
'' Description: Resize and move the controls on the form as the form is resized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Resize()
On Error Resume Next

    'If LimitFormSize(Me, lblNone.Left + lblNone.Width + cmdViewReport.Left, cmdViewReport.Height * 5) Then
    '    Exit Sub
    'End If

    With cmdViewReport
        '.Move (ScaleWidth - .Width) / 2
    End With
    
    With fgReports
        .Move .Left, cmdViewReport.Height + (cmdViewReport.Top * 2), ScaleWidth - (.Left * 2), _
                ScaleHeight - cmdViewReport.Height - (cmdViewReport.Top * 3)
        lblNoOrders.Move .Left, .Top, .Width, .Height
    End With
    
    AutoSizeChart

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: When the form unloads, save some properties
'' Inputs:      Whether to Cancel the unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    SetIniFileProperty "OrderTracker", FontToString(fgReports.Font), "Fonts", g.strIniFile
    frmMain.DockPro.RemoveForm Me.Name

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "OrderTracker.Form.Unload", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitGrid
'' Description: Initialize the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitGrid()
On Error GoTo ErrSection:

    Dim lRedraw As Long                 ' Current state of the grid's redraw
    
    With fgReports
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        .Editable = flexEDNone
        .ExplorerBar = flexExSortShow
        .AllowBigSelection = False
        .AllowSelection = False
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .ExtendLastCol = True
        .ScrollTrack = True
        .SelectionMode = flexSelectionFree '= flexSelectionListBox
        .SheetBorder = RGB(128, 128, 128)
        
        .Rows = 1
        .FixedRows = 1
        .Cols = GDCol(eGDCol_NumCols)
        .FixedCols = 0
        
        .TextMatrix(0, GDCol(eGDCol_Symbol)) = "Symbol"
        .TextMatrix(0, GDCol(eGDCol_SystemName)) = "Strategy"
        .TextMatrix(0, GDCol(eGDCol_Viewed)) = ""
        
        .ColHidden(GDCol(eGDCol_Hwnd)) = True
        .ColHidden(GDCol(eGDCol_CRCViewed)) = True
        .ColHidden(GDCol(eGDCol_Orders)) = True
        .ColHidden(GDCol(eGDCol_LastKnown)) = True
        
        .ColDataType(GDCol(eGDCol_Orders)) = flexDTBoolean
        
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "OrderTracker.InitGrid", eGDRaiseError_Raise
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ChartForHwnd
'' Description: Returns the correct chart form for the given HWND
'' Inputs:      Hwnd of the form to retrieve
'' Returns:     Chart Form
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ChartForHwnd(ByVal lHwnd As Long) As Form
On Error GoTo ErrSection:

    Dim frm As Form
    
    Set ChartForHwnd = Nothing
    For Each frm In Forms
        'If frm.hWnd = lHwnd And frm.Name = "frmChart" Then         'JM 06-04-2009: original code; leave awhile then remove if all okay
        If frm.hWnd = lHwnd And IsFrmChart(frm) Then
            Set ChartForHwnd = frm
            Exit For
        End If
    Next frm

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "OrderTracker.ChartForHwnd", eGDRaiseError_Raise
    Resume ErrExit

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowReport
'' Description: Show the Next Bar Report for the given row in the grid
'' Inputs:      Current Row in the Grid
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ShowReport(ByVal lRow As Long)
On Error GoTo ErrSection:

    Dim lHwnd As Long                   ' hWnd of Form to look for
    Dim frm As Form                     ' Chart form
    Dim lCRC As Long                    ' CRC of the Next Bar File
    
    PlaySoundFile ' cancel the sound alert
    
    lHwnd = fgReports.TextMatrix(lRow, GDCol(eGDCol_Hwnd))
    Set frm = ChartForHwnd(lHwnd)
    If Not frm Is Nothing Then
        frm.Chart.ShowSystemReport True
        'fgReports.Cell(flexcpPicture, lRow, GDCol(eGDCol_Viewed)) = imgBlack.Picture
        'fgReports.Cell(flexcpPictureAlignment, lRow, GDCol(eGDCol_Viewed)) = flexPicAlignCenterCenter
        'fgReports.TextMatrix(lRow, GDCol(eGDCol_LastKnown)) = "Black"
        'If gdCalcFileCRC32(AddSlash(App.Path) & "Trades\RB_" & lHwnd & ".TXT", lCRC) Then
        '    fgReports.TextMatrix(lRow, GDCol(eGDCol_CRCViewed)) = CStr(lCRC)
        'End If
        'MoveFocus fgReports
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "OrderTracker.ShowReport", eGDRaiseError_Raise
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowOrders
'' Description: Update the grid with the given information
'' Inputs:      Chart hWnd, Symbol being run, System being run
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowOrders(ByVal lHwnd As Long, ByVal strSymbol As String, ByVal strSystem As String, ByVal bOrders As Boolean)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bFound As Boolean               ' Is this entry already in the grid?
    Dim lRow As Long                    ' Row in the grid
    Dim lRedraw As Long                 ' Current state of the grid's redraw
    Dim lCRC As Long                    ' CRC of the Next Bar File
    Dim lPopUp As Long                  ' Does the user want to pop up a message?
    Dim lSound As Long                  ' Does the user want to play a custom sound?
    Dim strMessage As String            ' Message to display
    Dim strSoundFile As String          ' Sound File to play
    Dim bWasRed As Boolean              ' Was the picture already red?
    
    With fgReports
        ' Find the row with the given hWnd
        For lIndex = .FixedRows To .Rows - 1
            If CLng(.TextMatrix(lIndex, GDCol(eGDCol_Hwnd))) = lHwnd Then
                bFound = True
                lRow = lIndex
                Exit For
            End If
        Next lIndex
        
        ' don't need to do anything if passed a blank symbol and hWnd wasn't found
        If bFound Or strSymbol <> "" Then
            lRedraw = .Redraw
            .Redraw = flexRDNone
            
            ' If the symbol is non-blank, then update the row
            If strSymbol <> "" Then
                ' If not found, add a row to the grid
                If Not bFound Then
                    .Rows = .Rows + 1
                    lRow = .Rows - 1
                End If
            
                .TextMatrix(lRow, GDCol(eGDCol_Hwnd)) = Trim(CStr(lHwnd))
                .TextMatrix(lRow, GDCol(eGDCol_Symbol)) = strSymbol
                .TextMatrix(lRow, GDCol(eGDCol_SystemName)) = strSystem
                
                If bOrders Then
                    gdCalcFileCRC32 AddSlash(App.Path) & "Trades\RB_" & lHwnd & ".TXT", lCRC
                    
                    If .TextMatrix(lRow, GDCol(eGDCol_CRCViewed)) = "" Or CLng(ValOfText(.TextMatrix(lRow, GDCol(eGDCol_CRCViewed)))) <> lCRC Then
                        bWasRed = (.TextMatrix(lRow, GDCol(eGDCol_LastKnown)) = "Red")
                        
                        .Cell(flexcpPicture, lRow, GDCol(eGDCol_Viewed)) = imgRed.Picture
                        .TextMatrix(lRow, GDCol(eGDCol_LastKnown)) = "Red"
                        
                        If bFound = True And bWasRed = False Then
                            lPopUp = GetIniFileProperty("PopUp", vbUnchecked, "OrderTracker", g.strIniFile)
                            lSound = GetIniFileProperty("Sound", vbUnchecked, "OrderTracker", g.strIniFile)
                            strMessage = "New orders for" & vbCrLf & "Strategy: " & .TextMatrix(lRow, GDCol(eGDCol_SystemName)) & vbCrLf & "Symbol: " & .TextMatrix(lRow, GDCol(eGDCol_Symbol))
                            strSoundFile = GetIniFileProperty("SoundFile", "", "OrderTracker", g.strIniFile)
                            
                            If lPopUp = vbChecked Then
                                frmAlertPopup.ShowMe eGDAlertMode_NextBarAlert, .TextMatrix(lRow, GDCol(eGDCol_Symbol)), strMessage, lHwnd
                            End If
                            
                            If lSound = vbChecked Then
                                PlaySoundFile strSoundFile
                            End If
                        End If
                    Else
                        .Cell(flexcpPicture, lRow, GDCol(eGDCol_Viewed)) = imgBlack.Picture
                        .TextMatrix(lRow, GDCol(eGDCol_LastKnown)) = "Black"
                    End If
                Else
                    .Cell(flexcpPicture, lRow, GDCol(eGDCol_Viewed)) = imgNone.Picture
                    .TextMatrix(lRow, GDCol(eGDCol_LastKnown)) = "White"
                End If
                .Cell(flexcpPictureAlignment, lRow, GDCol(eGDCol_Viewed)) = flexPicAlignCenterCenter
                
                CheckedCell(fgReports, lRow, GDCol(eGDCol_Orders)) = bOrders
                
            ' Otherwise, if the symbol is blank, then remove the item from the grid
            ElseIf bFound Then
                .RemoveItem lRow
            End If
        
            If .Rows > .FixedRows + 1 And .Row < .FixedRows Then
                .Row = .FixedRows
            Else
                .Row = -1
            End If
            
            If .Rows > .FixedRows Then
                .Visible = True
                cmdViewReport.Enabled = True
                lblNoOrders.Visible = False
            Else
                .Visible = False
                cmdViewReport.Enabled = False
                lblNoOrders.Visible = True
            End If
            
            .AutoSize 0, .Cols - 1, False, 75
            .Redraw = flexRDBuffered
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "OrderTracker.ShowOrders", eGDRaiseError_Raise
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuChangeFont_Click
'' Description: Allow the user to change the font on the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuChangeFont_Click()
On Error GoTo ErrSection:

    ChangeGridFont fgReports, True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "OrderTracker.mnuChangeFont.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub mnuDisplayMult_Click()
On Error GoTo ErrSection:

    ShowMultiple

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderTracker.mnuDisplayMult.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuShowReport_Click
'' Description: Show the report for the entry that the user right clicked on
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuShowReport_Click()
On Error GoTo ErrSection:

    ShowReport m.lRow

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "OrderTracker.mnuShowReport.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMultiple
'' Description: Show all of the orders for next bar on the next bar form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ShowMultiple()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim astrFiles As New cGdArray       ' Next Bar Files to process
    Dim lHwnd As Long                   ' Hwnd from the grid for the chart form
    
    PlaySoundFile ' cancel the sound alert
    
    ' Compile the list of next bar files to display...
    astrFiles.Create eGDARRAY_Strings
    For lIndex = fgReports.FixedRows To fgReports.Rows - 1
        lHwnd = fgReports.TextMatrix(lIndex, GDCol(eGDCol_Hwnd))
        astrFiles.Add AddSlash(App.Path) & "Trades\NB_" & Str(lHwnd) & ".TXT"
    Next lIndex
    
    ' Show the next bar form...
    frmNextBar.ShowMeMult astrFiles, , "Order Tracker"
      
    ' Set the picture and CRC appropriately...
    For lIndex = fgReports.FixedRows To fgReports.Rows - 1
        fgReports.Cell(flexcpPicture, lIndex, GDCol(eGDCol_Viewed)) = imgBlack.Picture
        fgReports.Cell(flexcpPictureAlignment, lIndex, GDCol(eGDCol_Viewed)) = flexPicAlignCenterCenter
        fgReports.TextMatrix(lIndex, GDCol(eGDCol_LastKnown)) = "Black"
        'If gdCalcFileCRC32(AddSlash(App.Path) & "Trades\RB_" & lHwnd & ".TXT", lCRC) Then
        '    fgReports.TextMatrix(lIndex, GDCol(eGDCol_CRCViewed)) = CStr(lCRC)
        'End If
        MoveFocus fgReports
    Next lIndex

ErrExit:
    Set astrFiles = Nothing
    Exit Sub
    
ErrSection:
    Set astrFiles = Nothing
    RaiseError "frmOrderTracker.ShowMultiple", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MarkAsRead
'' Description: Mark the given symbol/strategy as read
'' Inputs:      Hwnd of the Chart
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub MarkAsRead(ByVal lHwnd As Long)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lCRC As Long                    ' CRC of the Next Bar File
    
    With fgReports
        For lIndex = .FixedRows To .Rows - 1
            If .TextMatrix(lIndex, GDCol(eGDCol_Hwnd)) = Str(lHwnd) Then
                .Cell(flexcpPicture, lIndex, GDCol(eGDCol_Viewed)) = imgBlack.Picture
                .Cell(flexcpPictureAlignment, lIndex, GDCol(eGDCol_Viewed)) = flexPicAlignCenterCenter
                .TextMatrix(lIndex, GDCol(eGDCol_LastKnown)) = "Black"
                If gdCalcFileCRC32(AddSlash(App.Path) & "Trades\RB_" & lHwnd & ".TXT", lCRC) Then
                    .TextMatrix(lIndex, GDCol(eGDCol_CRCViewed)) = CStr(lCRC)
                End If
                MoveFocus fgReports
                Exit For
            End If
        Next lIndex
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderTracker.MarkAsRead", eGDRaiseError_Raise
    
End Sub

