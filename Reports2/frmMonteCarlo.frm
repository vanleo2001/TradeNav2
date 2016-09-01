VERSION 5.00
Object = "{3B008041-905A-11D1-B4AE-444553540000}#1.0#0"; "Vsocx6.ocx"
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMonteCarlo 
   Caption         =   "Monte Carlo simulations to compare Hypothetical Systems"
   ClientHeight    =   8010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13500
   Icon            =   "frmMonteCarlo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8010
   ScaleWidth      =   13500
   StartUpPosition =   3  'Windows Default
Begin HexUniControls.ctlUniFrameWL fraMM
VistaStyle      =   0   'False
      Caption         =   "Money Management comparison (using Average Loss as Risk)"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6735
      Left            =   4560
      TabIndex        =   45
      Top             =   240
      Width           =   8715
Begin HexUniControls.ctlUniFrameWL Frame2
VistaStyle      =   0   'False
         BorderStyle     =   0  'None
         Caption         =   "Money Management comparison"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2400
         TabIndex        =   58
         Top             =   360
         Width           =   5595
Begin HexUniControls.ctlUniComboImageXP cboYears
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3720
Style=1
            TabIndex        =   74
            Top             =   555
            Width           =   735
         End
Begin HexUniControls.ctlUniTextBoxXP txtStartBalance
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1500
            TabIndex        =   10
            Text            =   "$50,000"
            Top             =   555
            Width           =   1095
         End
Begin HexUniControls.ctlUniTextBoxXP txtNumRuns
            Height          =   300
            Left            =   480
            TabIndex        =   9
            Text            =   "1000"
            Top             =   240
            Width           =   735
         End
Begin HexUniControls.ctlUniLabelXP Label13
            BackStyle       =   0  'Transparent
            Caption         =   "For each percentage of account to Risk per trade (base on $AvgLoss):"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   73
            Top             =   0
            Width           =   5295
         End
Begin HexUniControls.ctlUniLabelXP Label14
            BackStyle       =   0  'Transparent
            Caption         =   "simulation starts at"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   63
            Top             =   585
            Width           =   1635
         End
Begin HexUniControls.ctlUniLabelXP Label15
            BackStyle       =   0  'Transparent
            Caption         =   "years."
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4500
            TabIndex        =   62
            Top             =   585
            Width           =   555
         End
Begin HexUniControls.ctlUniLabelXP Label17
            BackStyle       =   0  'Transparent
            Caption         =   "and trades for"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2640
            TabIndex        =   61
            Top             =   585
            Width           =   1095
         End
Begin HexUniControls.ctlUniLabelXP Label22
            BackStyle       =   0  'Transparent
            Caption         =   "Run"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   60
            Top             =   300
            Width           =   735
         End
Begin HexUniControls.ctlUniLabelXP Label24
            BackStyle       =   0  'Transparent
            Caption         =   "simulations of randomly generated trades, where each"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1260
            TabIndex        =   59
            Top             =   300
            Width           =   3915
         End
      End
Begin HexUniControls.ctlUniFrameWL fraColors
VistaStyle      =   0   'False
         BorderStyle     =   0  'None
         Height          =   3615
         Left            =   4920
         TabIndex        =   46
         Top             =   1920
         Width           =   2955
Begin HexUniControls.ctlUniLabelXP lblColor
            Alignment       =   2  'Center
            BackColor       =   &H000000FF&
            Caption         =   "> 90%"
            Height          =   240
            Index           =   9
            Left            =   1020
            TabIndex        =   57
            Top             =   3180
            Width           =   795
         End
Begin HexUniControls.ctlUniLabelXP lblColor
            Alignment       =   2  'Center
            BackColor       =   &H008080FF&
            Caption         =   "80-90%"
            Height          =   240
            Index           =   8
            Left            =   1020
            TabIndex        =   56
            Top             =   2940
            Width           =   795
         End
Begin HexUniControls.ctlUniLabelXP lblColor
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0FF&
            Caption         =   "70-80%"
            Height          =   240
            Index           =   7
            Left            =   1020
            TabIndex        =   55
            Top             =   2700
            Width           =   795
         End
Begin HexUniControls.ctlUniLabelXP lblColor
            Alignment       =   2  'Center
            BackColor       =   &H00C0E0FF&
            Caption         =   "60-70%"
            Height          =   240
            Index           =   6
            Left            =   1020
            TabIndex        =   54
            Top             =   2460
            Width           =   795
         End
Begin HexUniControls.ctlUniLabelXP lblColor
            Alignment       =   2  'Center
            BackColor       =   &H0000FFFF&
            Caption         =   "50-60%"
            Height          =   240
            Index           =   5
            Left            =   1020
            TabIndex        =   53
            Top             =   2220
            Width           =   795
         End
Begin HexUniControls.ctlUniLabelXP lblColor
            Alignment       =   2  'Center
            BackColor       =   &H0080FFFF&
            Caption         =   "40-50%"
            Height          =   240
            Index           =   4
            Left            =   1020
            TabIndex        =   52
            Top             =   1980
            Width           =   795
         End
Begin HexUniControls.ctlUniLabelXP lblColor
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Caption         =   "30-40%"
            Height          =   240
            Index           =   3
            Left            =   1020
            TabIndex        =   51
            Top             =   1740
            Width           =   795
         End
Begin HexUniControls.ctlUniLabelXP lblColor
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFC0&
            Caption         =   "20-30%"
            Height          =   240
            Index           =   2
            Left            =   1020
            TabIndex        =   50
            Top             =   1500
            Width           =   795
         End
Begin HexUniControls.ctlUniLabelXP lblColor
            Alignment       =   2  'Center
            BackColor       =   &H0080FF80&
            Caption         =   "10-20%"
            Height          =   240
            Index           =   1
            Left            =   1020
            TabIndex        =   49
            Top             =   1260
            Width           =   795
         End
Begin HexUniControls.ctlUniLabelXP lblColor
            Alignment       =   2  'Center
            BackColor       =   &H0000FF00&
            Caption         =   "0-10%"
            Height          =   255
            Index           =   0
            Left            =   1020
            TabIndex        =   48
            Top             =   1020
            Width           =   795
         End
Begin HexUniControls.ctlUniLabelXP Label10
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Note: when running the risk analysis, the color of each row will be based on the Average Drawdown%"
            Height          =   795
            Left            =   180
            TabIndex        =   47
            Top             =   300
            Width           =   2595
         End
      End
Begin HexUniControls.ctlUniButtonImageXP cmdRisks
         Caption         =   "&RISK Analysis:"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   360
         TabIndex        =   8
         Top             =   420
         Width           =   1755
      End
      Begin VSFlex7LCtl.VSFlexGrid fgRisks 
         Height          =   6255
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   1380
         Visible         =   0   'False
         Width           =   4005
         _cx             =   7064
         _cy             =   11033
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
         Cols            =   3
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
         Editable        =   2
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
      Begin VSFlex7LCtl.VSFlexGrid fgRisks 
         Height          =   6195
         Index           =   1
         Left            =   4500
         TabIndex        =   12
         Top             =   1380
         Visible         =   0   'False
         Width           =   4005
         _cx             =   7064
         _cy             =   10927
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
         Cols            =   3
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
         Editable        =   2
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
Begin HexUniControls.ctlUniFrameWL fraStrategy
VistaStyle      =   0   'False
      Caption         =   "Performance comparison"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7515
      Left            =   240
      TabIndex        =   13
      Top             =   240
      Width           =   4035
Begin HexUniControls.ctlUniTextBoxXP txtTradesPerYear
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   1
         Left            =   3000
         TabIndex        =   72
         Top             =   780
         Width           =   795
      End
Begin HexUniControls.ctlUniTextBoxXP txtTradesPerYear
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   0
         Left            =   1800
         TabIndex        =   71
         Top             =   780
         Width           =   795
      End
      Begin MSComctlLib.ProgressBar pbRuns 
         Height          =   240
         Left            =   180
         TabIndex        =   64
         Top             =   7020
         Visible         =   0   'False
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   1
      End
Begin HexUniControls.ctlUniButtonImageXP cmdPerformance
         Caption         =   "&Calculate ANNUAL Performance:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   180
         TabIndex        =   7
         Top             =   5040
         Width           =   3675
      End
Begin HexUniControls.ctlUniTextBoxXP txtWinPercent
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   0
         Left            =   1800
         TabIndex        =   0
         Text            =   "80%"
         Top             =   1200
         Width           =   795
      End
Begin HexUniControls.ctlUniTextBoxXP txtAvgWin
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   0
         Left            =   1620
         TabIndex        =   1
         Text            =   "$100.00"
         Top             =   1620
         Width           =   975
      End
Begin HexUniControls.ctlUniTextBoxXP txtAvgLoss
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   0
         Left            =   1620
         TabIndex        =   2
         Text            =   "$180.00"
         Top             =   2040
         Width           =   975
      End
Begin HexUniControls.ctlUniTextBoxXP txtAvgLoss
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   1
         Left            =   2820
         TabIndex        =   5
         Text            =   "$180.00"
         Top             =   2040
         Width           =   975
      End
Begin HexUniControls.ctlUniTextBoxXP txtAvgWin
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   1
         Left            =   2820
         TabIndex        =   4
         Text            =   "$100.00"
         Top             =   1620
         Width           =   975
      End
Begin HexUniControls.ctlUniTextBoxXP txtWinPercent
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   1
         Left            =   3000
         TabIndex        =   3
         Text            =   "50%"
         Top             =   1200
         Width           =   795
      End
Begin HexUniControls.ctlUniCheckXP chkSameAvgTrade
Pressed = 0
         Caption         =   "keep Total $Wins and $Losses equal"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   780
         TabIndex        =   6
         Top             =   2520
         Width           =   3075
      End
Begin HexUniControls.ctlUniLabelXP Label12
         BackStyle       =   0  'Transparent
         Caption         =   "Trades per Year ="
         Height          =   195
         Left            =   180
         TabIndex        =   70
         Top             =   840
         Width           =   1455
      End
Begin HexUniControls.ctlUniLabelXP lblNumSims
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Avg of 10,000 simulations"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   435
         Left            =   2820
         TabIndex        =   68
         Top             =   6960
         Visible         =   0   'False
         Width           =   1035
         WordWrap        =   -1  'True
      End
Begin HexUniControls.ctlUniLabelXP lblCPC
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.66667"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   1
         Left            =   2760
         TabIndex        =   67
         Top             =   3960
         Width           =   975
      End
Begin HexUniControls.ctlUniLabelXP lblCPC
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.66667"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   0
         Left            =   1560
         TabIndex        =   66
         Top             =   3960
         Width           =   975
      End
Begin HexUniControls.ctlUniLabelXP Label11
         BackStyle       =   0  'Transparent
         Caption         =   "CPC Index ="
         Height          =   195
         Left            =   180
         TabIndex        =   65
         Top             =   3960
         Width           =   1215
      End
Begin HexUniControls.ctlUniLabelXP Label8
         BackStyle       =   0  'Transparent
         Caption         =   "="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1620
         TabIndex        =   44
         Top             =   6480
         Width           =   315
         WordWrap        =   -1  'True
      End
Begin HexUniControls.ctlUniLabelXP lblPerfNote
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "(these results are the averages from running 10,000 simulations of randomly generated trades)"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   435
         Left            =   240
         TabIndex        =   43
         Top             =   6960
         Width           =   3555
         WordWrap        =   -1  'True
      End
Begin HexUniControls.ctlUniLabelXP lblAvgRatio
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "123.45"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   1
         Left            =   2760
         TabIndex        =   42
         Top             =   6450
         Width           =   975
      End
Begin HexUniControls.ctlUniLabelXP lblAvgDrawdown
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "$1,000.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   1
         Left            =   2760
         TabIndex        =   41
         Top             =   6060
         Width           =   975
      End
Begin HexUniControls.ctlUniLabelXP lblAvgProfit
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "$1,000.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   1
         Left            =   2760
         TabIndex        =   40
         Top             =   5700
         Width           =   975
      End
Begin HexUniControls.ctlUniLabelXP lblAvgRatio
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "234.77"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   0
         Left            =   1560
         TabIndex        =   39
         Top             =   6450
         Width           =   975
      End
Begin HexUniControls.ctlUniLabelXP lblAvgDrawdown
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "$1,000.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   0
         Left            =   1560
         TabIndex        =   38
         Top             =   6060
         Width           =   975
      End
Begin HexUniControls.ctlUniLabelXP lblAvgProfit
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "$1,000.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   0
         Left            =   1560
         TabIndex        =   37
         Top             =   5700
         Width           =   975
      End
Begin HexUniControls.ctlUniLabelXP lblPerformance
         BackStyle       =   0  'Transparent
         Caption         =   "Average Annual Profit / Drawdown"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         TabIndex        =   36
         Top             =   6420
         Width           =   1575
         WordWrap        =   -1  'True
      End
Begin HexUniControls.ctlUniLabelXP Label19
         BackStyle       =   0  'Transparent
         Caption         =   "Avg Drawdown ="
         Height          =   195
         Left            =   180
         TabIndex        =   35
         Top             =   6060
         Width           =   1335
      End
Begin HexUniControls.ctlUniLabelXP Label18
         BackStyle       =   0  'Transparent
         Caption         =   "Average Profit ="
         Height          =   195
         Left            =   180
         TabIndex        =   34
         Top             =   5700
         Width           =   1215
      End
Begin HexUniControls.ctlUniLabelXP Label4
         BackStyle       =   0  'Transparent
         Caption         =   "Average Trade ="
         Height          =   195
         Left            =   180
         TabIndex        =   33
         Top             =   2880
         Width           =   1215
      End
Begin HexUniControls.ctlUniLabelXP Label5
         BackStyle       =   0  'Transparent
         Caption         =   "Payout Ratio ="
         Height          =   195
         Left            =   180
         TabIndex        =   32
         Top             =   3240
         Width           =   1215
      End
Begin HexUniControls.ctlUniLabelXP Label6
         BackStyle       =   0  'Transparent
         Caption         =   "Profit Factor ="
         Height          =   195
         Left            =   180
         TabIndex        =   31
         Top             =   3600
         Width           =   1215
      End
Begin HexUniControls.ctlUniLabelXP Label7
         BackStyle       =   0  'Transparent
         Caption         =   "Kelly Ratio ="
         Height          =   195
         Left            =   180
         TabIndex        =   30
         Top             =   4680
         Width           =   1215
      End
Begin HexUniControls.ctlUniLabelXP lblAvgTrade
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "$1,000.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   0
         Left            =   1560
         TabIndex        =   29
         Top             =   2880
         Width           =   975
      End
Begin HexUniControls.ctlUniLabelXP lblPayoutRatio
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.66667"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   0
         Left            =   1560
         TabIndex        =   28
         Top             =   3240
         Width           =   975
      End
Begin HexUniControls.ctlUniLabelXP lblProfitFactor
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.66667"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   0
         Left            =   1560
         TabIndex        =   27
         Top             =   3600
         Width           =   975
      End
Begin HexUniControls.ctlUniLabelXP lblKelly
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "66.67%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   0
         Left            =   1560
         TabIndex        =   26
         Top             =   4680
         Width           =   975
      End
Begin HexUniControls.ctlUniLabelXP lblExpectancy
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "66.67%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   0
         Left            =   1560
         TabIndex        =   25
         Top             =   4320
         Width           =   975
      End
Begin HexUniControls.ctlUniLabelXP Label9
         BackStyle       =   0  'Transparent
         Caption         =   "Expectancy ="
         Height          =   195
         Left            =   180
         TabIndex        =   24
         Top             =   4320
         Width           =   1215
      End
Begin HexUniControls.ctlUniLabelXP lblExpectancy
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "66.67%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   1
         Left            =   2760
         TabIndex        =   23
         Top             =   4320
         Width           =   975
      End
Begin HexUniControls.ctlUniLabelXP lblKelly
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "66.67%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   1
         Left            =   2760
         TabIndex        =   22
         Top             =   4680
         Width           =   975
      End
Begin HexUniControls.ctlUniLabelXP lblProfitFactor
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.66667"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   1
         Left            =   2760
         TabIndex        =   21
         Top             =   3600
         Width           =   975
      End
Begin HexUniControls.ctlUniLabelXP lblPayoutRatio
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.66667"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   1
         Left            =   2760
         TabIndex        =   20
         Top             =   3240
         Width           =   975
      End
Begin HexUniControls.ctlUniLabelXP lblAvgTrade
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "$1,000.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   1
         Left            =   2760
         TabIndex        =   19
         Top             =   2880
         Width           =   975
      End
Begin HexUniControls.ctlUniLabelXP Label1
         BackStyle       =   0  'Transparent
         Caption         =   "% Winning trades ="
         Height          =   195
         Left            =   180
         TabIndex        =   18
         Top             =   1260
         Width           =   1455
      End
Begin HexUniControls.ctlUniLabelXP Label2
         BackStyle       =   0  'Transparent
         Caption         =   "Average $ Win ="
         Height          =   195
         Left            =   180
         TabIndex        =   17
         Top             =   1680
         Width           =   1275
      End
Begin HexUniControls.ctlUniLabelXP Label3
         BackStyle       =   0  'Transparent
         Caption         =   "Average $ Loss ="
         Height          =   195
         Left            =   180
         TabIndex        =   16
         Top             =   2100
         Width           =   1335
      End
Begin HexUniControls.ctlUniLabelXP lblSystemB
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "System 'B'"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   2760
         TabIndex        =   15
         Top             =   360
         Width           =   1155
      End
Begin HexUniControls.ctlUniLabelXP lblSystemA
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "System 'A'"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1500
         TabIndex        =   14
         Top             =   360
         Width           =   1155
      End
   End
   Begin vsOcx6LibCtl.vsIndexTab vst 
      Height          =   4215
      Left            =   3720
      TabIndex        =   69
      Top             =   3180
      Visible         =   0   'False
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   7435
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   1
      MousePointer    =   0
      _ConvInfo       =   1
      Version         =   600
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FrontTabColor   =   -2147483633
      BackTabColor    =   -2147483633
      TabOutlineColor =   0
      FrontTabForeColor=   -2147483630
      Caption         =   "Profit / Drawdown Analysis|Risk% Analysis"
      Align           =   0
      Appearance      =   1
      CurrTab         =   1
      FirstTab        =   0
      Style           =   3
      Position        =   0
      AutoSwitch      =   -1  'True
      AutoScroll      =   -1  'True
      TabPreview      =   -1  'True
      ShowFocusRect   =   -1  'True
      TabsPerPage     =   0
      BorderWidth     =   0
      BoldCurrent     =   -1  'True
      DogEars         =   -1  'True
      MultiRow        =   0   'False
      MultiRowOffset  =   200
      CaptionStyle    =   0
      TabHeight       =   0
   End
End
Attribute VB_Name = "frmMonteCarlo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type mPrivate
    aTrades(0 To 1) As cGdArray ' used only if passing actual trades from a backtest

    ' basic stats for systems
    dWinPercent(0 To 1) As Double
    dAvgWin(0 To 1) As Double
    dAvgLoss(0 To 1) As Double
    dAvgTrade(0 To 1) As Double
    dTradesPerYear(0 To 1) As Double
    
    ' parameters for simulations
    nNumRuns As Long
    dStartBalance As Double
    dTargetBalance As Double
    
    bRunning As Boolean ' flag when in progress
End Type
Private m As mPrivate

Private Sub chkSameAvgTrade_Click()
    
    DisplayStats
    
End Sub

Private Sub Form_Activate()

    fgRisks(0).BackColorBkg = Me.BackColor
    fgRisks(1).BackColorBkg = Me.BackColor

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim i&, s$

    ' allow Ctrl-# (where # is 1-9) to specify some predefined values
    If Shift = 2 Then
        i = KeyCode - Asc("0")
        If i >= 0 And i <= 9 Then
            KeyCode = 0
            m.bRunning = True
            m.aTrades(0).Size = 0
            m.aTrades(1).Size = 0
            Select Case i
            Case 1
                chkSameAvgTrade.Value = 1
                txtWinPercent(0) = "90"
                txtAvgWin(0) = "100"
                txtAvgLoss(0) = "225"
                txtWinPercent(1) = "10"
            Case 2
                chkSameAvgTrade.Value = 0
                txtWinPercent(0) = "54.64"
                txtAvgWin(0) = "39.39"
                txtAvgLoss(0) = "10.25"
                txtWinPercent(1) = "87.27"
                txtAvgWin(1) = "31.01"
                txtAvgLoss(1) = "78.98"
            Case 3
                chkSameAvgTrade.Value = 1
                txtWinPercent(0) = "74.5"
                txtAvgWin(0) = "466"
                txtAvgLoss(0) = "693"
                txtWinPercent(1) = ""
            Case 0
                's = "c:\dvlp\genesis\navigator suite\trades\s22.txt"
                s = "c:\dvlp\Batting800.txt"
                Set m.aTrades(0) = LoadTrades(s)
                m.dWinPercent(0) = 0
                s = "c:\dvlp\MarketPulse.txt"
                Set m.aTrades(1) = LoadTrades(s)
                m.dWinPercent(1) = 0
            End Select
            m.bRunning = False
            DisplayStats
            MoveFocus cmdPerformance
        End If
    End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        'DisplayStats
        KeyAscii = 0
        SendKeys vbTab
    End If

End Sub

Private Sub Form_Load()
    
    Dim s$, aTrades As cGdArray
    
    Randomize
    CenterTheForm Me
    
    'RH
     g.Styler.StyleForm Me
     
    'CenterTheControl pbRuns, lblPerfNote
    
    #If STANDALONE_EXE Then
        's = "c:\dvlp\genesis\navigator suite\trades\s22.txt"
        's = "c:\dvlp\Batting800.txt"
        's = "c:\dvlp\MarketPulse.txt"
        Set aTrades = LoadTrades(s)
        
        ShowMe aTrades
    #End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If UnloadMode = 0 Then
        If m.bRunning Then
            m.bRunning = False
            Cancel = True
            Beep
        End If
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set m.aTrades(0) = Nothing
    Set m.aTrades(1) = Nothing

End Sub

Private Sub Form_Resize()

    On Error Resume Next
    
    Dim i&
    i = Me.ScaleHeight - fraMM.Top - fraStrategy.Left
    If i < fraStrategy.Height Then
        i = fraStrategy.Height
    End If
    fraMM.Height = i
    fgRisks(0).Height = fraMM.Height - fgRisks(0).Top - 150
    fgRisks(1).Height = fgRisks(0).Height

End Sub

Public Sub ShowMe(Optional ByVal aTrades As cGdArray, Optional ByVal aTradesCompare As cGdArray, Optional ByVal dAvgTradesPerYear As Double = 0)
            
    Dim i&
            
    cboYears.Clear
    For i = 1 To 30
        cboYears.AddItem Str(i)
    Next
    cboYears.ListIndex = 4
            
    For i = 0 To 1
        If dAvgTradesPerYear >= 1 Then
            m.dTradesPerYear(i) = Round(dAvgTradesPerYear)
        Else
            m.dTradesPerYear(i) = 100 ' default
        End If
        
        m.dWinPercent(i) = 0
        m.dAvgWin(i) = 0
        m.dAvgLoss(i) = 0
        
        ' if passing actual trades from a backtest:
        If i = 0 Then
            Set m.aTrades(i) = aTrades
        Else
            Set m.aTrades(i) = aTradesCompare
        End If
        If m.aTrades(i) Is Nothing Then
            Set m.aTrades(i) = New cGdArray
            m.aTrades(i).Create eGDARRAY_Doubles, 0
        End If
    Next
            
    ' first calculate the initial stats based on the actual trades passed in
    DisplayStats
    
    ' if 2nd set of trades was not passed in then just initially make the stats for B the same as A
    If m.aTrades(1).Size = 0 Then
        txtTradesPerYear(1) = txtTradesPerYear(0)
        txtWinPercent(1) = txtWinPercent(0)
        txtAvgWin(1) = txtAvgWin(0)
        txtAvgLoss(1) = txtAvgLoss(0)
    End If
    
    ' TLB 4/2/2015: due to not having the correct risk/stop-loss values for each trade from a basket of muliple
    ' strategy/symbol combinations, it would be quite inaccurate to continue trying to use the actual trades.
    ' So we will just clear out all the actual trades at this point since from here on out we are only using
    ' hypothetically generated trades (based on the %win and average $win and $loss stats).
    m.aTrades(0).Size = 0
    m.aTrades(1).Size = 0
    
    ' then calc once more (using hypotheticals)
    DisplayStats
    
    #If STANDALONE_EXE = 0 Then
        ShowForm Me
    #End If

End Sub

Private Sub InitGrids()

    Dim iSystem&
    
    For iSystem = 0 To 1
        fgRisks(iSystem).Rows = 2
        SetupGrid fgRisks(iSystem), eGridMode_Grid
        With fgRisks(iSystem)
            .ExtendLastCol = True
            .SelectionMode = flexSelectionFree
            .ExplorerBar = flexExNone
            .BackColorBkg = Me.BackColor
            .ScrollBars = flexScrollBarVertical
            
            .Cols = 4
            .ColHidden(3) = True
'.ScrollBars = flexScrollBarBoth
            .FixedRows = 2
            .Rows = .FixedRows
            
            .Cell(flexcpAlignment, 0, 0, 1, .Cols - 1) = flexAlignCenterCenter
            .MergeCells = flexMergeFixedOnly
            .MergeRow(0) = True
            .MergeCol(0) = True
            .TextMatrix(0, 0) = "Risk%"
            .TextMatrix(1, 0) = "Risk%"
            If iSystem = 0 Then
                .TextMatrix(0, 1) = "System 'A'"
                .Cell(flexcpForeColor, 0, 1, 0, 3) = lblSystemA.ForeColor
            Else
                .TextMatrix(0, 1) = "System 'B'"
                .Cell(flexcpForeColor, 0, 1, 0, 3) = lblSystemB.ForeColor
            End If
            .TextMatrix(0, 2) = .TextMatrix(0, 1)
            .TextMatrix(0, 3) = .TextMatrix(0, 2)
            .Cell(flexcpFontSize, 0, 1, 0, .Cols - 1) = 10
            .RowHeight(0) = Int(.RowHeight(1) * 1.25)
            
            .TextMatrix(1, 1) = "Avg Ending Balance"
            .TextMatrix(1, 2) = "Avg Drawdown%"
            .TextMatrix(1, 3) = "% Bankrupts"
            
            .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = True
            
            .AutoSizeMode = flexAutoSizeColWidth
            .AutoSize 0, .Cols - 1
            
            '.Select 0, 1, 2, 3
            '.CellBorder RGB(0, 0, 128), 2, 2, 2, 2, 0, 0
            '.Select 0, 0
        End With
    Next

End Sub

Private Sub DisplayStats()

    Dim i&, d#, s$, bChanged As Boolean
    Dim iTrade&, iWins&, iLosses&
    
    If m.bRunning Then Exit Sub
    m.bRunning = True

    ' for each of the 2 systems
    For i = 0 To 1
        If m.aTrades(i).Size = 0 Then
            ' fix inputs entered by user
            d = ValOfText(txtWinPercent(i))
            If d < 0 Then
                d = 0
            ElseIf d > 100 Then
                d = 0.99
            ElseIf d >= 1 Then
                d = d / 100
            End If
            If m.dWinPercent(i) <> d Then
                m.dWinPercent(i) = d
                bChanged = True
            End If
            
            d = Round(ValOfText(txtTradesPerYear(i)))
            If d < 1 Then
                d = 100
            End If
            If m.dTradesPerYear(i) <> d Then
                m.dTradesPerYear(i) = d
                bChanged = True
            End If
            txtWinPercent(i).Enabled = True
            
            If i = 1 And chkSameAvgTrade.Value <> 0 Then
                If m.dWinPercent(1) > 0 Then
                    ' make Avg Trades the same for both systems
                    m.dAvgWin(1) = m.dAvgWin(0) * m.dWinPercent(0) / m.dWinPercent(1)
                    m.dAvgLoss(1) = m.dAvgLoss(0) * (1 - m.dWinPercent(0)) / (1 - m.dWinPercent(1))
                    txtAvgWin(i).Visible = True
                    txtAvgLoss(i).Visible = True
                Else
                    txtAvgWin(i).Visible = False
                    txtAvgLoss(i).Visible = False
                End If
                txtAvgWin(i).Enabled = False
                txtAvgLoss(i).Enabled = False
                txtTradesPerYear(i).Enabled = False
                txtTradesPerYear(1).Text = txtTradesPerYear(0).Text
            Else
                d = Abs(ValOfText(txtAvgWin(i)))
                If m.dAvgWin(i) <> d Then
                    m.dAvgWin(i) = d
                    bChanged = True
                End If
                d = Abs(ValOfText(txtAvgLoss(i)))
                If m.dAvgLoss(i) <> d Then
                    m.dAvgLoss(i) = d
                    bChanged = True
                End If
                txtAvgWin(i).Enabled = True
                txtAvgLoss(i).Enabled = True
                txtTradesPerYear(i).Enabled = True
            End If
        ElseIf m.dWinPercent(i) = 0 Then
            ' calc inputs from the actual trades
            bChanged = True
            m.dAvgWin(i) = 0
            m.dAvgLoss(i) = 0
            iWins = 0
            iLosses = 0
            For iTrade = 0 To m.aTrades(i).Size - 1
                d = m.aTrades(i).Num(iTrade)
                If d > 0 Then
                    m.dAvgWin(i) = m.dAvgWin(i) + d
                    iWins = iWins + 1
                Else
                    m.dAvgLoss(i) = m.dAvgLoss(i) + d
                    iLosses = iLosses + 1
                End If
            Next
            m.dWinPercent(i) = CDbl(iWins) / m.aTrades(i).Size
            If iWins > 0 Then
                m.dAvgWin(i) = Abs(m.dAvgWin(i) / iWins)
            End If
            If iLosses > 0 Then
                m.dAvgLoss(i) = Abs(m.dAvgLoss(i) / iLosses)
            End If
            ' and disable inputs
            txtWinPercent(i).Enabled = False
            txtAvgWin(i).Enabled = False
            txtAvgLoss(i).Enabled = False
            txtTradesPerYear(i).Enabled = False
        End If
        m.dAvgTrade(i) = 0
        
        If m.dWinPercent(i) <= 0 Or m.dWinPercent(i) >= 1 Or m.dAvgLoss(i) <= 0 Or m.dAvgWin(i) <= 0 Then
            ' clear stats
            lblAvgTrade(i) = ""
            lblPayoutRatio(i) = ""
            lblProfitFactor(i) = ""
            lblKelly(i) = ""
            lblExpectancy(i) = ""
            lblCPC(i) = ""
            lblAvgProfit(i) = ""
            lblAvgDrawdown(i) = ""
            lblAvgRatio(i) = ""
        Else
            ' display inputs
            txtTradesPerYear(i) = Str(Round(m.dTradesPerYear(i)))
            txtWinPercent(i) = Format(m.dWinPercent(i) * 100, "#0.0#") & "%"
            If m.dAvgWin(i) >= 10000 Then
                txtAvgWin(i) = Format(m.dAvgWin(i), "$#,##0")
            Else
                txtAvgWin(i) = Format(m.dAvgWin(i), "$#,##0.00")
            End If
            If m.dAvgLoss(i) >= 10000 Then
                txtAvgLoss(i) = Format(m.dAvgLoss(i), "$#,##0")
            Else
                txtAvgLoss(i) = Format(m.dAvgLoss(i), "$#,##0.00")
            End If
            
            ' calc and display stats
            m.dAvgTrade(i) = m.dWinPercent(i) * m.dAvgWin(i) - (1 - m.dWinPercent(i)) * m.dAvgLoss(i)
            If m.dAvgTrade(i) >= 10000 Then
                lblAvgTrade(i) = Format(m.dAvgTrade(i), "$#,##0")
            Else
                lblAvgTrade(i) = Format(m.dAvgTrade(i), "$#,##0.00")
            End If
            
            d = m.dAvgWin(i) / m.dAvgLoss(i)
            lblPayoutRatio(i) = Format(d, "#0.0##")
            d = (m.dWinPercent(i) * m.dAvgWin(i)) / ((1 - m.dWinPercent(i)) * m.dAvgLoss(i))
            lblProfitFactor(i) = Format(d, "#0.0##")
            d = (d ^ 2) * (1 - m.dWinPercent(i))
            lblCPC(i) = Format(d, "#0.0##")
            d = (1 + m.dAvgWin(i) / m.dAvgLoss(i)) * m.dWinPercent(i) - 1
            lblExpectancy(i) = Format(d * 100, "#0.0#") & "%"
        
            d = m.dWinPercent(i) - (1 - m.dWinPercent(i)) / (m.dAvgWin(i) / m.dAvgLoss(i))
            If d > 0 Then
                lblKelly(i) = Format(d * 100, "#0.0#") & "%"
            Else
                lblKelly(i) = "0%"
            End If
        End If
    Next
    
    If bChanged Then
        For i = 0 To 1
            fgRisks(i).Visible = False
            lblAvgProfit(i) = ""
            lblAvgDrawdown(i) = ""
            lblAvgRatio(i) = ""
        Next
        fraColors.Visible = True
    End If
    
    If m.aTrades(1).Size = 0 Then
        chkSameAvgTrade.Enabled = True
    Else
        chkSameAvgTrade.Value = 0
        chkSameAvgTrade.Enabled = False
    End If

    m.bRunning = False

End Sub

Private Sub cmdPerformance_Click()
    
    RunPerformance
    
End Sub

Private Sub RunPerformance()

    Dim i&, iSystem&, iMid&, iRun&, dProfit#, dDD#
    Dim dSumProfit(0 To 1) As Double
    Dim dSumDD(0 To 1) As Double
    Dim dMaxDD(0 To 1) As Double

    If m.bRunning Then Exit Sub

    DisplayStats

    m.nNumRuns = 10000
        
    'Rnd -1
    'Randomize 43
    
    m.bRunning = True
    Screen.MousePointer = vbHourglass
    pbRuns.Value = 0
    pbRuns.Visible = True
    pbRuns.ZOrder
    cmdPerformance.Enabled = False
    cmdRisks.Enabled = False
    lblPerfNote.Visible = False
    lblNumSims = ""
    lblNumSims.Visible = True
    
    For iSystem = 0 To 1
        lblAvgProfit(iSystem) = ""
        lblAvgDrawdown(iSystem) = ""
        lblAvgRatio(iSystem) = ""
        dSumProfit(iSystem) = 0
        dSumDD(iSystem) = 0
        dMaxDD(iSystem) = 0
    Next
    Me.Refresh
    
    For iRun = 1 To m.nNumRuns
        For iSystem = 0 To 1
            If m.dAvgTrade(iSystem) > 0 Then
                TestRun iSystem, 0, dProfit, dDD
                dSumProfit(iSystem) = dSumProfit(iSystem) + dProfit
                dSumDD(iSystem) = dSumDD(iSystem) + dDD
                If dDD > dMaxDD(iSystem) Then
                    dMaxDD(iSystem) = dDD
                End If
                If iRun Mod 250 = 0 Or iRun = m.nNumRuns Then
                    ' update progress bar
                    pbRuns.Value = Int(100# * iRun / m.nNumRuns + 0.5)
                    lblNumSims = "Avg of " & Format(iRun, "#,##0") & " simulations"
                    lblAvgProfit(iSystem) = Format(dSumProfit(iSystem) / iRun, "$#,##0")
                    lblAvgDrawdown(iSystem) = Format(dSumDD(iSystem) / iRun, "$#,##0")
                    lblAvgRatio(iSystem) = Format(dSumProfit(iSystem) / dSumDD(iSystem), "#0.00")
                    DoEvents
                End If
                If Not m.bRunning Then Exit For ' user aborting
            End If
        Next
        If Not m.bRunning Then Exit For ' user aborting
    Next
    
    pbRuns.Visible = False
    cmdPerformance.Enabled = True
    cmdRisks.Enabled = True
    lblPerfNote.Visible = True
    lblNumSims.Visible = False
    Screen.MousePointer = vbDefault
    m.bRunning = False
    
End Sub

Private Sub cmdRisks_Click()

    Dim i&, s$, dNetProfit#, dMaxDD#, dAvgProfit#, dAvgMaxDD#, dLargestMaxDD#
    Dim dBalance#, dAvgBalance#, nNumZeros&, dMaxDDP#, dAvgMaxDDP#, dLargestMaxDDP#
    Static strRunCaption$

    If Len(strRunCaption) = 0 Then
        strRunCaption = cmdRisks.Caption
    End If

    ' first check if user is aborting a run which is already in progress
    If m.bRunning Then
        'i = fgRisks.Rows
        'fgRisks.Rows = i + 1
        'fgRisks.TextMatrix(i, 1) = "Aborted"
        fgRisks(0).Visible = False
        fgRisks(1).Visible = False
        fraColors.Visible = True
    Else
        DoEvents ' to allow other LostFocus events to run first
    
        If lblAvgProfit(0) = "" And lblAvgProfit(1) = "" Then
            'RunPerformance
        End If
    
        m.nNumRuns = ValOfText(txtNumRuns)
        If m.nNumRuns < 10 Then
            m.nNumRuns = 1000
            txtNumRuns = "1000"
        End If
        
        m.dStartBalance = ValOfText(txtStartBalance)
        If m.dStartBalance < 10 Then
            m.dStartBalance = 50000
            txtStartBalance = "$50,000"
        End If
        
        m.bRunning = True
        cmdPerformance.Enabled = False
        cmdRisks.Caption = "&Abort"
        RunRisks
    End If
    m.bRunning = False
    cmdPerformance.Enabled = True
    cmdRisks.Caption = strRunCaption

End Sub

Private Sub RunRisks()

    Dim bDone(2) As Boolean
    Dim i&, iRow&, iMid&, iSystem&
    Dim iRisk&, iRun&, iBankrupts&, dAvgEndBalance#, dEndBalance#
    Dim dDDP#, dAvgDDP#, dMaxDDP# ' drawdown percentages
    Dim aEndBalances As New cGdArray, aDrawdowns As New cGdArray
    Dim s$

    InitGrids
    iRow = fgRisks(0).FixedRows
    fgRisks(0).Rows = iRow
    fgRisks(1).Rows = iRow
    fgRisks(1).ZOrder

    ' NOTE: when doing money management analysis, the "normal" averaging doesn't work
    ' very well due to the worst-case scenarios getting stopped out at zero (as opposed
    ' to going negative) -- so the much better "averaging" method is to use the MEDIAN
    ' instead (which requires storing all of them and sorting when done).
    aEndBalances.Create eGDARRAY_Floats, m.nNumRuns, 0
    aDrawdowns.Create eGDARRAY_Floats, m.nNumRuns, 0
    
    For iRisk = 1 To 99
        For iSystem = 0 To 1
            If m.dAvgTrade(iSystem) > 0 And Not bDone(iSystem) Then
                If iRisk = 1 Then
                    fgRisks(iSystem).Visible = True
                    If iSystem = 1 Then
                        'fraColors.Visible = False
                    End If
                    Me.Refresh
                End If
                
                dAvgEndBalance = 0
                dAvgDDP = 0
                dMaxDDP = 0
                iBankrupts = 0
                aEndBalances.Clear False
                aDrawdowns.Clear False
                For iRun = 1 To m.nNumRuns
                    ' run a simulation for this risk%
                    If TestRun(iSystem, iRisk / 100#, dEndBalance, dDDP) = False Then
                        iBankrupts = iBankrupts + 1
                        dEndBalance = 0
                        dDDP = 1 ' bankrupt = 100% drawdown
                    End If
                    aEndBalances.Num(iRun - 1) = dEndBalance
                    aDrawdowns.Num(iRun - 1) = dDDP
                    If dDDP > dMaxDDP Then
                        dMaxDDP = dDDP
                    End If
                    If Not m.bRunning Then Exit For ' user aborting
                Next
                If Not m.bRunning Then Exit For ' user aborting
                
                ' sort results and find the median (do the average of 6 middle spots)
                aEndBalances.Sort
                aDrawdowns.Sort
                dAvgEndBalance = 0
                dAvgDDP = 0
                iMid = Int(m.nNumRuns / 2#)
                For i = iMid - 3 To iMid + 2
                    dAvgEndBalance = dAvgEndBalance + aEndBalances.Num(i)
                    dAvgDDP = dAvgDDP + aDrawdowns.Num(i)
                Next
                dAvgEndBalance = dAvgEndBalance / 6
                dAvgDDP = dAvgDDP / 6
                
                ' output to grid row: AvgEndBalance, AvgDD%, MaxDD%, #Bankrupts
                With fgRisks(iSystem)
                    .MergeCells = flexMergeFree
                    .Rows = iRow + 1
                    .TextMatrix(iRow, 0) = Str(iRisk) & "%"
                    '.Select 0, 1, .Rows - 1, 1
                    '.CellBorder RGB(0, 0, 0), 2, 0, 0, 0, 0, 0
                    '.Select 0, 4, .Rows - 1, 4
                    '.CellBorder RGB(0, 0, 0), 2, 0, 0, 0, 0, 0
                    .Select 0, 0, 0, 0
                    .Cell(flexcpFontBold, iRow, 2, iRow, 2) = True
                    '.Cell(flexcpFontBold, iRow, 5, iRow, 5) = True
                    
                    If dAvgEndBalance = m.dStartBalance And dAvgDDP = 0 Then
                        .MergeRow(iRow) = True
                        s = " starting balance too low"
                        .TextMatrix(iRow, 1) = s
                        .TextMatrix(iRow, 2) = s
                        .TextMatrix(iRow, 3) = ""
                        .Select iRow, 1, iRow, 2
                        .CellAlignment = flexAlignCenterCenter
                        .Cell(flexcpBackColor, iRow, 1, iRow, 3) = RGB(224, 224, 224)
                    Else
If dAvgDDP > 0.98 Then
    i = i
End If
                        .MergeRow(iRow) = False
                        .TextMatrix(iRow, 1) = Format(dAvgEndBalance, "$#,##0")
                        .TextMatrix(iRow, 2) = Format(dAvgDDP, "#0.00%")
                        .TextMatrix(iRow, 3) = Format(CDbl(iBankrupts) / m.nNumRuns, "#0.0%")
                        '.TextMatrix(iRow, 4) = Format(dMaxDDP, "#0.00%")
                        i = Int(dAvgDDP * 10)
                        If i > 9 Then
                            i = 9
                        End If
                        .Cell(flexcpBackColor, iRow, 1, iRow, 3) = lblColor(i).BackColor
                    End If
                End With
                DoEvents
                
                ' don't bother continuing once get past a certain point
                If dAvgEndBalance <= 0 Or dAvgDDP > 0.98 Then
                    bDone(iSystem) = True
                End If
            End If
        Next ' for each system
        
        If bDone(0) And bDone(1) Then
            Exit For
        End If
        iRow = iRow + 1
    Next ' for each risk%

End Sub

Private Function TestRun(ByVal iSystem&, ByVal dRiskPercent#, dBalance#, dMaxDD#) As Boolean

    Dim i&, iTrade&, dRandom#, dPeak#, dTrade#, dNumRisk#, iShuffleTrade&, nNumTrades&
    Dim aShuffled As New cGdArray
    Dim bUseBacktest As Boolean, bUseShuffleMethod As Boolean
        
    If m.aTrades(iSystem).Size > 0 Then
        bUseBacktest = True
'bUseShuffleMethod = True
        If bUseShuffleMethod Then
            Set aShuffled = m.aTrades(iSystem).MakeCopy
            gdShuffle aShuffled.ArrayHandle, 0, aShuffled.Size - 1
        End If
    End If
        
    nNumTrades = m.dTradesPerYear(iSystem)
    If dRiskPercent > 0 Then
        dBalance = m.dStartBalance
        nNumTrades = nNumTrades * Val(cboYears.Text)
    Else
        dBalance = 0
    End If
    dPeak = dBalance
    dMaxDD = 0
    For iTrade = 1 To nNumTrades
        ' random trade
        If Not bUseBacktest Then
            ' randomly determine win or loss
            dRandom = Rnd 'dRandom = gdRandomNumber(0, 999999999) / 1000000000#
            If dRandom < m.dWinPercent(iSystem) Then
                dTrade = m.dAvgWin(iSystem)
            Else
                dTrade = -m.dAvgLoss(iSystem)
            End If
        ElseIf bUseShuffleMethod Then
            ' get a shuffled trade from the backtest
            dTrade = aShuffled.Num(iShuffleTrade)
            iShuffleTrade = iShuffleTrade + 1
            If iShuffleTrade >= aShuffled.Size Then
                iShuffleTrade = 0
            End If
        Else
            ' get a random trade from the backtest
            i = gdRandomNumber(0, m.aTrades(iSystem).Size - 1)
            dTrade = m.aTrades(iSystem).Num(i)
        End If
        
        If dRiskPercent = 0 Then
            dBalance = dBalance + dTrade
            ' calc $MaxDD if 1 contract per trade
            If dBalance > dPeak Then
                dPeak = dBalance
            ElseIf dPeak - dBalance > dMaxDD Then
                dMaxDD = dPeak - dBalance
            End If
        Else
            ' for Money Management analysis, determine
            ' # contracts to risk for this trade
            dNumRisk = Int(dBalance * dRiskPercent / Abs(m.dAvgLoss(iSystem)))
            If dNumRisk < 1 Then dNumRisk = 0 '1
            dBalance = dBalance + dTrade * dNumRisk
            If dBalance <= 0 Then
'If dBalance <= Abs(m.dAvgLoss(iSystem)) Then
                ' BANKRUPT !!
                dBalance = 0
                dMaxDD = 1 ' bankrupt = 100% drawdown
                TestRun = False
                Exit Function
            End If
            
            ' calc MaxDD% if money management
            If dBalance > dPeak Then
                dPeak = dBalance
            ElseIf dPeak > 0 Then
                If (dPeak - dBalance) / dPeak > dMaxDD Then
                    dMaxDD = (dPeak - dBalance) / dPeak
                End If
            End If
            
            ' allow for user to abort a MM run
            If iTrade Mod 1000 = 0 Then
                DoEvents
            End If
            If Not m.bRunning Then Exit For ' user aborting
        End If
    Next
    
    TestRun = True

End Function

Private Sub txtAvgLoss_GotFocus(Index As Integer)
    SelectAll txtAvgLoss(Index)
End Sub

Private Sub txtAvgLoss_LostFocus(Index As Integer)
    DisplayStats
End Sub

Private Sub txtAvgWin_GotFocus(Index As Integer)
    SelectAll txtAvgWin(Index)
End Sub

Private Sub txtAvgWin_LostFocus(Index As Integer)
    DisplayStats
End Sub

Private Sub txtTradesPerYear_GotFocus(Index As Integer)
    SelectAll txtTradesPerYear(Index)
End Sub

Private Sub txtTradesPerYear_LostFocus(Index As Integer)
    DisplayStats
End Sub

Private Sub txtWinPercent_GotFocus(Index As Integer)
    SelectAll txtWinPercent(Index)
End Sub

Private Sub txtWinPercent_LostFocus(Index As Integer)
    DisplayStats
End Sub

Private Function LoadTrades(ByVal strFile$) As cGdArray

    Dim fh%, d#, strLine$, strChk$, bHeaderDone As Boolean
    Dim aTrades As New cGdArray
    
    On Error GoTo ErrExit:
    
    aTrades.Create eGDARRAY_Doubles, 0
    
    If FileExist(strFile) Then
        fh = FreeFile
        Open strFile For Input As #fh
        Do While Not EOF(fh)
            Line Input #fh, strLine
            strChk = Left(strLine, 2)
            If strChk = "S" & vbTab Or strChk = "L" & vbTab Then
                d = Val(Parse(strLine, vbTab, 8))
                aTrades.Add d
            End If
        Loop
        Close #fh
    End If
    
ErrExit:
    Set LoadTrades = aTrades
    Exit Function
    
End Function

