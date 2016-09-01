VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmWhatsNew 
   Caption         =   "What's New in Trade Navigator?"
   ClientHeight    =   6165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8640
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6165
   ScaleWidth      =   8640
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   2175
      Left            =   7140
      TabIndex        =   3
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
      Caption         =   "frmWhatsNew.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmWhatsNew.frx":002C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmWhatsNew.frx":004C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdUpgrade 
         Height          =   435
         Left            =   120
         TabIndex        =   1
         Top             =   1680
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
         Caption         =   "frmWhatsNew.frx":0068
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmWhatsNew.frx":0098
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmWhatsNew.frx":00B8
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdClose 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   120
         TabIndex        =   4
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
         Caption         =   "frmWhatsNew.frx":00D4
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmWhatsNew.frx":0100
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmWhatsNew.frx":0120
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblUpgrade 
         Height          =   495
         Left            =   60
         Top             =   1260
         Visible         =   0   'False
         Width           =   1020
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
         Caption         =   "frmWhatsNew.frx":013C
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Alignment       =   2
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmWhatsNew.frx":018E
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmWhatsNew.frx":01AE
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblVersion 
         Height          =   675
         Left            =   60
         Top             =   600
         Width           =   1020
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
         Caption         =   "frmWhatsNew.frx":01CA
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   2
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmWhatsNew.frx":022C
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmWhatsNew.frx":024C
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniTextBoxXP txtDesc 
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   4380
      Width           =   5835
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   -1  'True
      Text            =   "frmWhatsNew.frx":0268
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
      MultiLine       =   -1  'True
      Alignment       =   0
      ScrollBars      =   2
      PasswordChar    =   ""
      TrapTab         =   0   'False
      EnableContextMenu=   -1  'True
      RaiseChangeEvent=   -1  'True
      Tip             =   "frmWhatsNew.frx":0288
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmWhatsNew.frx":02A8
   End
   Begin VSFlex7LCtl.VSFlexGrid fg 
      Height          =   2835
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      _cx             =   10186
      _cy             =   5001
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
   Begin HexUniControls.ctlUniLabelXP lblDesc 
      Height          =   255
      Left            =   120
      Top             =   4080
      Width           =   675
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
      Caption         =   "frmWhatsNew.frx":02C4
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmWhatsNew.frx":02F6
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmWhatsNew.frx":0316
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
End
Attribute VB_Name = "frmWhatsNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum eGridColumns
    eCol_Version = 0
    eCol_Build = 1
    eCol_Feature = 2
    eCol_Desc = 3
    eCol_Module = 4
End Enum

Private Sub cmdClose_Click()
On Error GoTo ErrSection:

    Unload Me

ErrExit:
    Exit Sub
ErrSection:
    RaiseError "frmWhatsNew.cmdClose", eGDRaiseError_Show
    Resume ErrExit
End Sub

Private Sub cmdUpgrade_Click()
On Error GoTo ErrSection:

    If ProcessIsBusy Then Exit Sub
            
    cmdUpgrade.Enabled = False
    If FormIsLoaded("frmDownload") Then Unload frmDownload
    Set MsgForm = frmStatus
    frmDownload.optSpecialFile = True
    frmDownload.txtSpecialFile = "Upgrade"
    frmDownload.DownloadData
    Set MsgForm = Nothing
    cmdUpgrade.Enabled = True

ErrExit:
    Exit Sub
ErrSection:
    RaiseError "frmWhatsNew.cmdUpgrade", eGDRaiseError_Show
    Resume ErrExit
End Sub

Private Sub fg_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    Dim strDesc$
    If NewRow >= fg.FixedRows Then
        strDesc = Replace(fg.TextMatrix(NewRow, GridCol(eCol_Desc)), "|", vbCrLf)
    End If
    txtDesc = strDesc

ErrExit:
    Exit Sub
ErrSection:
    RaiseError "frmWhatsNew.fg_AfterRowColChange", eGDRaiseError_Show
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
    RaiseError "frmWhatsNew.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
End Sub

Private Sub Form_Load()
On Error GoTo ErrSection:

    Me.Icon = Picture16(ToolbarIcon("ID_WhatsNew"))
    CenterTheForm Me
    
    g.Styler.StyleForm Me

ErrExit:
    Exit Sub
ErrSection:
    RaiseError "frmWhatsNew.Form_Load", eGDRaiseError_Show
    Resume ErrExit
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    If LimitFormSize(Me, fraButtons.Width + lblDesc.Left + lblDesc.Width, _
            txtDesc.Height + fraButtons.Height + fraButtons.Top + 120) Then Exit Sub
    
    lblDesc.Top = Me.ScaleHeight - txtDesc.Height - lblDesc.Height - 120
    With txtDesc
        .Move .Left, lblDesc.Top + lblDesc.Height, Me.ScaleWidth - .Left * 2, .Height
    End With
    fraButtons.Left = Me.ScaleWidth - fraButtons.Width
    With fg
        .Move .Left, .Top, fraButtons.Left - fg.Left, lblDesc.Top - .Top - 60
    End With

End Sub

Private Function GridCol(ByVal lColumn As eGridColumns) As Long
    GridCol = lColumn
End Function

Public Sub ShowMe()
On Error GoTo ErrSection:

    Dim i&, bNeedsUpgrade As Boolean
    Dim aList As New cGdArray, aFields As New cGdArray

    ' init the grid
    SetupGrid fg, eGridMode_Grid
    With fg
        .Redraw = flexRDNone
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .Cols = 5
        .FixedCols = 0
        .TextMatrix(0, GridCol(eCol_Version)) = "Version"
        .TextMatrix(0, GridCol(eCol_Build)) = "Build"
        .TextMatrix(0, GridCol(eCol_Feature)) = "Feature added"
        .ColAlignment(GridCol(eCol_Version)) = flexAlignCenterCenter
        .ColAlignment(GridCol(eCol_Build)) = flexAlignCenterCenter
        .ColAlignment(GridCol(eCol_Feature)) = flexAlignLeftCenter
        .ColHidden(GridCol(eCol_Desc)) = True
        .ColHidden(GridCol(eCol_Module)) = True
        .ColWidth(GridCol(eCol_Version)) = 900
        .ColWidth(GridCol(eCol_Build)) = 750
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignLeftCenter
        
        ' fill info from the What's New file
        aList.FromFile App.Path & "\Info\WhatsNew.txt"
        .Rows = .FixedRows
        For i = 0 To aList.Size - 1
            aFields.SplitFields aList(i), vbTab
            If aFields.Size >= 3 Then
                If IsDigit(aFields(0), 1) Then
                    If HasModule(aFields(4)) Then
                        If Val(aFields(1)) > App.Revision Then
                            bNeedsUpgrade = True
                            aFields(1) = aFields(1) & "*"
                        End If
                        .AddItem aFields.JoinFields(vbTab)
                        If Right(aFields(1), 1) = "*" Then
                            .Cell(flexcpForeColor, .Rows - 1, GridCol(eCol_Build)) = lblUpgrade.ForeColor
                        End If
                    End If
                End If
            End If
        Next
        .Row = .FixedRows
        
        .Redraw = flexRDBuffered
    End With

    ' version and upgrade message
    ' (don't show upgrade button if just upgraded)
    lblVersion = "You have:  Version " & FormatVersion _
        & "  Build " & App.Revision
    If FileExist(App.Path & "\ftp\upgrd32.exe") Then
        lblUpgrade.Visible = False
        cmdUpgrade.Visible = False
    Else
        lblUpgrade.Visible = bNeedsUpgrade
        cmdUpgrade.Visible = True
    End If
    
    ShowForm Me, False, frmMain, , ALT_GRID_ROW_COLOR

ErrExit:
    Exit Sub
ErrSection:
    RaiseError "frmWhatsNew.ShowMe", eGDRaiseError_Show
    Resume ErrExit
End Sub


