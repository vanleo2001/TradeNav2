VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmCompSymbol 
   Caption         =   "Comparison Symbol"
   ClientHeight    =   4950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   5010
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   300
      Top             =   4485
   End
   Begin gdOCX.gdSelectColor gdSelectColor1 
      Height          =   300
      Left            =   4020
      TabIndex        =   6
      Top             =   4425
      Visible         =   0   'False
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   529
      CustomColor     =   255
   End
   Begin HexUniControls.ctlUniFrameWL fraOptButtons 
      Height          =   1455
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   4050
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
      Caption         =   "frmCompSymbol.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmCompSymbol.frx":0020
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmCompSymbol.frx":0040
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniRadioXP optOverlayBars 
         Height          =   285
         Left            =   300
         TabIndex        =   2
         Top             =   285
         Width           =   3690
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
         Caption         =   "frmCompSymbol.frx":005C
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmCompSymbol.frx":00BE
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmCompSymbol.frx":00DE
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optPercentChart 
         Height          =   285
         Left            =   300
         TabIndex        =   5
         Top             =   1095
         Width           =   3690
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
         Caption         =   "frmCompSymbol.frx":00FA
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmCompSymbol.frx":0166
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmCompSymbol.frx":0186
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optOverlay 
         Height          =   285
         Left            =   300
         TabIndex        =   3
         Top             =   555
         Width           =   3690
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
         Caption         =   "frmCompSymbol.frx":01A2
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmCompSymbol.frx":0206
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmCompSymbol.frx":0226
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optNewLinear 
         Height          =   285
         Left            =   300
         TabIndex        =   1
         Top             =   15
         Width           =   3690
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
         Caption         =   "frmCompSymbol.frx":0242
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmCompSymbol.frx":02A6
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmCompSymbol.frx":02C6
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optPercentPane 
         Height          =   285
         Left            =   300
         TabIndex        =   4
         Top             =   825
         Width           =   3690
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
         Caption         =   "frmCompSymbol.frx":02E2
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmCompSymbol.frx":0356
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmCompSymbol.frx":0376
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraCmdButtons 
      Height          =   540
      Left            =   1350
      TabIndex        =   7
      Top             =   4335
      Width           =   2310
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
      Caption         =   "frmCompSymbol.frx":0392
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmCompSymbol.frx":03B2
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmCompSymbol.frx":03D2
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   1230
         TabIndex        =   10
         Top             =   90
         Width           =   915
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
         Caption         =   "frmCompSymbol.frx":03EE
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmCompSymbol.frx":041A
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmCompSymbol.frx":043A
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Default         =   -1  'True
         Height          =   375
         Left            =   165
         TabIndex        =   9
         Top             =   90
         Width           =   915
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
         Caption         =   "frmCompSymbol.frx":0456
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmCompSymbol.frx":047A
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmCompSymbol.frx":049A
         RightToLeft     =   0   'False
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid fgPercentComp 
      Height          =   2610
      Left            =   240
      TabIndex        =   8
      Top             =   1680
      Visible         =   0   'False
      Width           =   4530
      _cx             =   7990
      _cy             =   4604
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
   Begin HexUniControls.ctlUniLabelXP lblSelectSym 
      Height          =   270
      Left            =   480
      Top             =   90
      Width           =   4050
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
      Caption         =   "frmCompSymbol.frx":04B6
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmCompSymbol.frx":0524
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmCompSymbol.frx":0544
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
End
Attribute VB_Name = "frmCompSymbol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type mPrivate
    Chart As cChart
    aNewSymbols As cGdArray
    strSelected As String
    
    bColorChecked As Boolean
    lMouseRow As Long
    lMouseCol As Long
End Type
Private m As mPrivate

Public Function ShowMe(aSymbols As cGdArray, Chart As cChart) As String
On Error GoTo ErrSection:

    Dim i&
    Dim bShowGrid As Boolean
    Dim eCompSymType As eCompSym

    If aSymbols Is Nothing Then Exit Function
    If Chart Is Nothing Then Exit Function
    
    Set m.aNewSymbols = aSymbols
    Set m.Chart = Chart

    InitPercentCompGrid fgPercentComp, aSymbols, fgPercentComp.Height, m.Chart.Symbol

    CenterTheForm Me
    
    eCompSymType = Chart.CompSymType(eCompSym_PercentPane)      '6542
    
    If eCompSymType = eCompSym_PercentPane Then
        Me.optPercentPane = True
        FixControls True
        
'JM 12-12-2011: Not sure why original implementation explicit turned radio buttons on/off.
'   Aardvark 6542 specifies desired behaviors for handling percent change comparison
'
'        fraOptButtons.Visible = False
'        fgPercentComp.Top = fraOptButtons.Top + 75
'        fgPercentComp.Height = fgPercentComp.Height + fraOptButtons.Height
'        lblSelectSym.Caption = "Select symbols for percent change comparison"
    Else
        Me.optOverlay = True
        FixControls False

'        fraOptButtons.Visible = True
'        fraCmdButtons.Top = fraOptButtons.Top + fraOptButtons.Height - 90
'        Me.Height = Me.Height + 60
    End If
    
    ShowForm Me, True
    ShowMe = m.strSelected
    Unload Me

ErrExit:
    Exit Function

ErrSection:
    RaiseError "frmCompSymbol.ShowMe"
    
End Function

Private Sub cmdCancel_Click()
On Error Resume Next

    m.strSelected = ""
    Me.Hide

End Sub

Private Sub cmdOK_Click()
On Error Resume Next
    
    Dim nRec As Long
    Dim aSymbols As New cGdArray
    
    m.strSelected = ""
    If optPercentChart Or optPercentPane Then
        If optPercentChart Then
            m.strSelected = "C"
        Else
            m.strSelected = "P"
        End If
        ParsePercentCompGrid fgPercentComp, m.aNewSymbols, m.Chart.Symbol
        Me.Hide
    Else
        Set aSymbols = frmSymbolSelector.ShowMe("$DJIA", False, True, "Comparison Symbol", True)
        If aSymbols.Size > 0 Then
            nRec = g.SymbolPool.PoolRecForSymbol(aSymbols(0), True)
            ' TLB 4/11/2012: allow if an external symbol (from hard drive)
            If nRec < 0 And InStr(aSymbols(0), "|") = 0 Then
                Beep
            Else
                m.aNewSymbols.Size = 0
                m.aNewSymbols(0) = aSymbols(0)
                If optOverlay Then
                    m.strSelected = "O"
                ElseIf optOverlayBars Then
                    m.strSelected = "B"
                Else
                    m.strSelected = "N"
                End If
                Me.Hide
            End If
        End If
    End If
    
End Sub

Private Sub CheckColorSelect()
On Error GoTo ErrSection
    
    Dim iColor As Long
    
    If gdSelectColor1.Visible Then
        iColor = gdSelectColor1.Color
        If iColor = 0 Then iColor = -1
        With fgPercentComp
            If m.lMouseCol = 2 Then
                If m.lMouseRow >= .FixedRows And m.lMouseRow < .Rows Then
                    .Cell(flexcpBackColor, m.lMouseRow, 2) = iColor
                    .Select m.lMouseRow, m.lMouseCol
                    .Refresh
                    m.bColorChecked = True
                End If
            End If
        End With
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmCompSymbol.CheckColorSelect"
    
End Sub

Private Sub fgPercentComp_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim nRec&, lRow&, lCol&
    Dim aSymbols As New cGdArray
    Dim lColor&
    
    If tmr.Enabled Then
        Cancel = True
        Exit Sub
    End If
    
    With fgPercentComp
        
        lCol = .MouseCol
        lRow = .MouseRow
    
        If lCol = 1 And lRow = .Rows - 1 Then
            Set aSymbols = frmSymbolSelector.ShowMe("$DJIA", False, True, "Comparison Symbol", True)
            If aSymbols.Size > 0 Then
                nRec = g.SymbolPool.PoolRecForSymbol(aSymbols(0), True)
                ' TLB 4/11/2012: allow if an external symbol (from hard drive)
                If nRec < 0 And InStr(aSymbols(0), "|") = 0 Then
                    Beep
                Else
                    lColor = gdSelectColor1.Color
                    If lColor = 0 Then lColor = -1
                    
                    .TextMatrix(lRow, 1) = aSymbols(0)
                    .Cell(flexcpChecked, lRow, 0) = flexChecked
                    .Cell(flexcpPictureAlignment, lRow, 0) = flexAlignCenterCenter
                    .Cell(flexcpBackColor, lRow, 2) = lColor
                    
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 1) = "Click to add..."
                End If
            End If
        End If
    
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmCompSymbol.fgPercentComp_BeforeMouseDown"
    
End Sub

Private Sub fgPercentComp_BeforeScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long, Cancel As Boolean)
    gdSelectColor1.Visible = False
End Sub

Private Sub fgPercentComp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    Dim lCol&, lRow&, lRowsHeight&
    
    
    If tmr.Enabled Then Exit Sub
    
    With fgPercentComp
        lCol = .MouseCol
        lRow = .MouseRow
        
        If lCol = 2 Then
            If lRow >= .FixedRows And lRow < .Rows - 1 Then
                If InStr(.TextMatrix(lRow, 2), "Click") = 0 Then
                    m.lMouseRow = lRow
                    
                    lRowsHeight = .RowHeight(0) * .Rows
                    If lRowsHeight < .ClientHeight Then
                        gdSelectColor1.Move .Left + .Width - .ColWidth(2), .Top + .RowHeight(0) * .MouseRow, .ColWidth(2)
                    Else
                        'adjust for vertical scroll bar
                        gdSelectColor1.Move .Left + .Width - .ColWidth(2), .Top + .RowHeight(0) * (.MouseRow - .TopRow + 1), .ColWidth(2) - 225
                    End If
                    
                    gdSelectColor1.Color = .Cell(flexcpBackColor, lRow, lCol)
                    gdSelectColor1.Visible = True
                    gdSelectColor1.ZOrder
                End If
            End If
        Else
            gdSelectColor1.Visible = False
        End If
    End With

End Sub

Private Sub Form_Load()
On Error GoTo ErrSection:

    Me.Icon = Picture16("kBlank")
    
    g.Styler.StyleForm Me
    
    fgPercentComp.Visible = False
    fraCmdButtons.Top = fgPercentComp.Top - 45
    Me.Height = 2670

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmCompSymbol.Form_Load"
    
End Sub

Private Sub gdSelectColor1_Changed()
On Error Resume Next
    
    CheckColorSelect        'JM 06-11-2010 this event fires in the compiled EXE, but not in the IDE

End Sub

Private Sub gdSelectColor1_DropDown()
On Error Resume Next
    
    m.lMouseCol = 2
    m.bColorChecked = False
    tmr.Enabled = True
    
End Sub

Private Sub optNewLinear_Click()
On Error GoTo ErrSection:

    If Not Me.Visible Then Exit Sub
    If tmr.Enabled Then Exit Sub
    
    FixControls False
    cmdOK_Click
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmCompSymbol.optNewLinear_Click"
    
End Sub

Private Sub optOverlay_Click()
On Error GoTo ErrSection:

    If Not Me.Visible Then Exit Sub
    If tmr.Enabled Then Exit Sub

    FixControls False
    cmdOK_Click

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmCompSymbol.optOverlay_Click"

End Sub

Private Sub optOverlayBars_Click()
On Error GoTo ErrSection:

    If Not Me.Visible Then Exit Sub
    If tmr.Enabled Then Exit Sub

    FixControls False
    cmdOK_Click

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmCompSymbol.optOverlayBars_Click"

End Sub

Private Sub optPercentChart_Click()
On Error GoTo ErrSection:

    If Not Me.Visible Then Exit Sub
    If tmr.Enabled Then Exit Sub
    
    FixControls True

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmCompSymbol.optPercentChart_Click"
    
End Sub

Private Sub optPercentPane_Click()
On Error GoTo ErrSection:

    If Not Me.Visible Then Exit Sub
    If tmr.Enabled Then Exit Sub
    
    FixControls True

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmCompSymbol.optPercentPane_Click"
    
End Sub

Private Sub FixControls(bShowGrid As Boolean)
On Error GoTo ErrSection:
    
    If tmr.Enabled Then Exit Sub

    gdSelectColor1.Visible = False
    
    If bShowGrid Then
        If Not fgPercentComp.Visible Then fgPercentComp.Visible = True
        fgPercentComp.Top = fraOptButtons.Top + fraOptButtons.Height + 30
        fraCmdButtons.Top = fgPercentComp.Top + fgPercentComp.Height - 30
        lblSelectSym.Caption = "Select symbols for percent change comparison"
        Me.Height = 5460
    Else
        If Me.fgPercentComp.Visible Then fgPercentComp.Visible = False
        fraCmdButtons.Top = fraOptButtons.Top + fraOptButtons.Height - 90
        lblSelectSym.Caption = "Select where to place comparison symbol"
        Me.Height = 2730
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmCompSymbol.FixControls"
    
End Sub

Private Sub tmr_Timer()
On Error Resume Next

    If gdSelectColor1.Visible Then
        If Not gdSelectColor1.DropDownVisible Then
            If m.bColorChecked Then
                gdSelectColor1.Visible = False
                tmr.Enabled = False
            Else
                CheckColorSelect
            End If
        End If
    Else
        tmr.Enabled = False
    End If

End Sub

