VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{3B008041-905A-11D1-B4AE-444553540000}#1.0#0"; "Vsocx6.ocx"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmCattleAdmin 
   Caption         =   "Form1"
   ClientHeight    =   3765
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   10965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   10965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin vsOcx6LibCtl.vsIndexTab tabAdmin 
      Height          =   3135
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   5530
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      Caption         =   "&Default Lot Columns|&Customer Administration"
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
      Begin HexUniControls.ctlUniFrameWL fraCustomerAdministration 
         Height          =   2760
         Left            =   45
         TabIndex        =   6
         Top             =   330
         Width           =   10605
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
         Caption         =   "frmCattleAdmin.frx":0000
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmCattleAdmin.frx":002C
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmCattleAdmin.frx":004C
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniComboImageXP cboGenesisCustomers 
            Height          =   315
            Left            =   1560
            TabIndex        =   8
            Top             =   60
            Width           =   2055
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
            Tip             =   "frmCattleAdmin.frx":0068
            Sorted          =   -1  'True
            HScroll         =   0   'False
            RoundedBorders  =   -1  'True
            IconDim         =   16
            MousePointer    =   0
            MouseIcon       =   "frmCattleAdmin.frx":0088
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniFrameWL fraCustomerInfo 
            Height          =   2055
            Left            =   0
            TabIndex        =   9
            Top             =   480
            Width           =   3615
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
            Caption         =   "frmCattleAdmin.frx":00A4
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmCattleAdmin.frx":00EC
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmCattleAdmin.frx":010C
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniTextBoxXP txtName 
               Height          =   285
               Left            =   960
               TabIndex        =   13
               Top             =   580
               Width           =   2475
               _ExtentX        =   0
               _ExtentY        =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Enabled         =   -1  'True
               Locked          =   0   'False
               Text            =   "frmCattleAdmin.frx":0128
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
               MultiLine       =   0   'False
               Alignment       =   0
               ScrollBars      =   0
               PasswordChar    =   ""
               TrapTab         =   0   'False
               EnableContextMenu=   -1  'True
               RaiseChangeEvent=   -1  'True
               Tip             =   "frmCattleAdmin.frx":0148
               HideSelection   =   -1  'True
               RightToLeft     =   0   'False
               ManualStart     =   0   'False
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmCattleAdmin.frx":0168
            End
            Begin HexUniControls.ctlUniTextBoxXP txtPassword 
               Height          =   285
               Left            =   960
               TabIndex        =   0
               Top             =   920
               Width           =   2475
               _ExtentX        =   0
               _ExtentY        =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Enabled         =   -1  'True
               Locked          =   0   'False
               Text            =   "frmCattleAdmin.frx":0184
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
               MultiLine       =   0   'False
               Alignment       =   0
               ScrollBars      =   0
               PasswordChar    =   ""
               TrapTab         =   0   'False
               EnableContextMenu=   -1  'True
               RaiseChangeEvent=   -1  'True
               Tip             =   "frmCattleAdmin.frx":01A4
               HideSelection   =   -1  'True
               RightToLeft     =   0   'False
               ManualStart     =   0   'False
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmCattleAdmin.frx":01C4
            End
            Begin HexUniControls.ctlUniComboImageXP cboType 
               Height          =   315
               Left            =   960
               TabIndex        =   7
               Top             =   1260
               Width           =   2475
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
               Tip             =   "frmCattleAdmin.frx":01E0
               Sorted          =   0   'False
               HScroll         =   0   'False
               RoundedBorders  =   -1  'True
               IconDim         =   16
               MousePointer    =   0
               MouseIcon       =   "frmCattleAdmin.frx":0200
               DropDownOnTextClick=   -1  'True
               DropDownWidth   =   -1
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniTextBoxXP txtAccount 
               Height          =   285
               Left            =   960
               TabIndex        =   11
               Top             =   240
               Width           =   2475
               _ExtentX        =   0
               _ExtentY        =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Enabled         =   -1  'True
               Locked          =   0   'False
               Text            =   "frmCattleAdmin.frx":021C
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
               MultiLine       =   0   'False
               Alignment       =   0
               ScrollBars      =   0
               PasswordChar    =   ""
               TrapTab         =   0   'False
               EnableContextMenu=   -1  'True
               RaiseChangeEvent=   -1  'True
               Tip             =   "frmCattleAdmin.frx":023C
               HideSelection   =   -1  'True
               RightToLeft     =   0   'False
               ManualStart     =   0   'False
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmCattleAdmin.frx":025C
            End
            Begin HexUniControls.ctlUniCheckXP chkCanEditLots 
               Height          =   220
               Left            =   120
               TabIndex        =   10
               Top             =   1680
               Width           =   3315
               _ExtentX        =   5847
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
               Caption         =   "frmCattleAdmin.frx":0278
               Enabled         =   -1  'True
               Align           =   0
               CheckBackColor  =   -2147483643
               CheckForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmCattleAdmin.frx":02B4
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmCattleAdmin.frx":02D4
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblName 
               Height          =   195
               Left            =   120
               Top             =   625
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
               Caption         =   "frmCattleAdmin.frx":02F0
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmCattleAdmin.frx":031C
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmCattleAdmin.frx":033C
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblPassword 
               Height          =   195
               Left            =   120
               Top             =   965
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
               Caption         =   "frmCattleAdmin.frx":0358
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmCattleAdmin.frx":038C
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmCattleAdmin.frx":03AC
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblType 
               Height          =   195
               Left            =   120
               Top             =   1320
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
               Caption         =   "frmCattleAdmin.frx":03C8
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmCattleAdmin.frx":03F4
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmCattleAdmin.frx":0414
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblAccount 
               Height          =   195
               Left            =   120
               Top             =   285
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
               Caption         =   "frmCattleAdmin.frx":0430
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmCattleAdmin.frx":0466
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmCattleAdmin.frx":0486
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
         End
         Begin VSFlex7LCtl.VSFlexGrid fgColumns 
            Height          =   2055
            Left            =   6900
            TabIndex        =   12
            Top             =   540
            Width           =   3555
            _cx             =   6271
            _cy             =   3625
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
         Begin VSFlex7LCtl.VSFlexGrid fgYardInfo 
            Height          =   2055
            Left            =   3720
            TabIndex        =   14
            Top             =   540
            Width           =   3015
            _cx             =   5318
            _cy             =   3625
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
         Begin HexUniControls.ctlUniLabelXP lblGenesisCustomers 
            Height          =   195
            Left            =   0
            Top             =   120
            Width           =   1515
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
            Caption         =   "frmCattleAdmin.frx":04A2
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmCattleAdmin.frx":04E8
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmCattleAdmin.frx":0508
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraDefaultLotColumns 
         Height          =   2760
         Left            =   -11250
         TabIndex        =   2
         Top             =   330
         Width           =   10605
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
         Caption         =   "frmCattleAdmin.frx":0524
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmCattleAdmin.frx":0550
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmCattleAdmin.frx":0570
         RightToLeft     =   0   'False
         Begin VSFlex7LCtl.VSFlexGrid fgCustomerTypes 
            Height          =   2055
            Left            =   0
            TabIndex        =   3
            Top             =   120
            Width           =   3555
            _cx             =   6271
            _cy             =   3625
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
         Begin VSFlex7LCtl.VSFlexGrid fgFeedYardSources 
            Height          =   2055
            Left            =   3660
            TabIndex        =   4
            Top             =   120
            Width           =   3555
            _cx             =   6271
            _cy             =   3625
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
         Begin VSFlex7LCtl.VSFlexGrid fgLotColumns 
            Height          =   2055
            Left            =   7320
            TabIndex        =   5
            Top             =   120
            Width           =   3555
            _cx             =   6271
            _cy             =   3625
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
      End
   End
   Begin ActiveToolBars.SSActiveToolBars tbToolbar 
      Left            =   10200
      Top             =   60
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131083
      ToolBarsCount   =   1
      ToolsCount      =   2
      Tools           =   "frmCattleAdmin.frx":058C
      ToolBars        =   "frmCattleAdmin.frx":060E
   End
   Begin HexUniControls.ctlUniLabelXP lblStatus 
      Height          =   195
      Left            =   360
      Top             =   120
      Width           =   1155
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
      Caption         =   "frmCattleAdmin.frx":06A1
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmCattleAdmin.frx":06DB
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmCattleAdmin.frx":06FB
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin VB.Image imgStatus 
      Height          =   195
      Left            =   120
      Picture         =   "frmCattleAdmin.frx":0717
      Top             =   120
      Width           =   195
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "Pop Up"
      Begin VB.Menu mnuSelectAll 
         Caption         =   "&Select All"
      End
      Begin VB.Menu mnuDeselectAll 
         Caption         =   "&Deselect All"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
   End
End
Attribute VB_Name = "frmCattleAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmCattleAdmin.frm
'' Description: Form for allowing user to perform Cattle administration
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 06/11/2012   DAJ         Make Turnkey work with all brokers
'' 06/25/2012   DAJ         Visible Columns mode for Turnkey Admin
'' 09/11/2012   DAJ         Changed icon, added "Feed Yard" customer type
'' 09/14/2012   DAJ         Visible Lot Column by Genesis Customer not Feedyard Customer
'' 09/14/2012   DAJ         Set Dirty flag on select/deselect all
'' 10/22/2012   DAJ         Rename Turnkey to HedgeLinc
'' 01/30/2013   DAJ         Live/Demo/Test modes for Turnkey
'' 11/15/2013   DAJ         Added "Can Edit Lots"; Changed way to get icon/name
'' 12/03/2013   DAJ         Default visible lot columns
'' 12/04/2013   DAJ         Fix for dirty flag
'' 01/03/2014   DAJ         "Either Feedyard" mode
'' 01/31/2014   DAJ         Use KeyValueField for columns instead of ColumnName
'' 03/07/2014   DAJ         Moved into NavCattle.DLL
'' 03/17/2014   DAJ         Renamed Turnkey to Cattle for admin stuff
'' 04/15/2014   DAJ         New lot column administration
'' 04/21/2014   DAJ         Map customers and feedyards back to one enum; Fix for
''                          visible lot column overrides for user
'' 05/14/2014   DAJ         Optionally toggle all customers with toggle of feedyard
'' 05/22/2014   DAJ         Renamed g.Turnkey to g.Cattle
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Enum eGDCols
    eGDCol_Visible = 0
    eGDCol_Name
    eGDCol_ID
    eGDCol_NumCols
End Enum

Private Enum eGDTabs
    eGDTab_LotColumns = 0
    eGDTab_Customers = 1
End Enum

Private Type mPrivate
    nStatus As eGDConnectionStatus      ' Connection status to the Genesis Cattle server
    lSelectedID As Long                 ' Selected ID in the form
    
    iColumnsButton As Integer           ' Last mouse button pressed in columns grid
    iLotColumnsButton As Integer        ' Last mouse button pressed in lot columns grid
    iYardsButton As Integer             ' Last mouse button pressed in feed yards grid
    strCopiedColumns As String          ' Column list that has been copied
    bChangingCustomers As Boolean       ' Are we changing the Genesis customers in the combo?
    
    GenesisCustomers As cGdTree         ' Collection of Genesis customers
    DefaultColumns As cGdTree           ' Collection of default visible lot columns
    VisibleColumns As cGdTree           ' Visible lot columns for Genesis user
End Type
Private m As mPrivate

Private Property Get GDCol(ByVal nCol As eGDCols) As Long
    GDCol = nCol
End Property

Private Property Get GDTab(ByVal nTab As eGDTabs) As Long
    GDTab = nTab
End Property

Private Property Get CurrentTab() As eGDTabs
    CurrentTab = tabAdmin.CurrTab
End Property
Private Property Let CurrentTab(ByVal nTab As eGDTabs)
    tabAdmin.CurrTab = nTab
End Property

Public Property Get Status() As eGDConnectionStatus
    Status = m.nStatus
End Property
Public Property Let Status(ByVal nStatus As eGDConnectionStatus)
On Error GoTo ErrSection:

    Select Case nStatus
        Case eGDConnectionStatus_Disconnected
            imgStatus.Picture = frmCattleAM.imgRed
            lblStatus.Caption = "Disconnected"
            
            If m.nStatus <> eGDConnectionStatus_Disconnected Then
                ClearForm True
            End If
            
        Case eGDConnectionStatus_Disconnecting
            imgStatus.Picture = frmCattleAM.imgYellow
            lblStatus.Caption = "Disconnecting"
            
        Case eGDConnectionStatus_Connecting
            imgStatus.Picture = frmCattleAM.imgYellow
            lblStatus.Caption = "Connecting"
            
        Case eGDConnectionStatus_Connected
            imgStatus.Picture = frmCattleAM.imgGreen
            lblStatus.Caption = "Connected"
            
            If m.nStatus <> eGDConnectionStatus_Connected Then
                If CurrentTab = eGDTab_Customers Then
                    GetGenesisCustomers
                Else
                    GetAllLotColumns
                End If
            End If
                
    End Select

    If nStatus <> m.nStatus Then
        g.Cattle.DumpDebug "Connection status changed to " & lblStatus.Caption
        m.nStatus = nStatus
        
        SetFormCaption
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "frmCattleAdmin.Status.Let"
    
End Property

Private Property Get Dirty() As Boolean
    Dirty = tbToolbar.Tools("ID_Save").Enabled
End Property
Private Property Let Dirty(ByVal bDirty As Boolean)
    tbToolbar.Tools("ID_Save").Enabled = bDirty
End Property

Private Property Get SelectedType(Optional ByVal bConvert As Boolean = False, Optional ByVal lRow As Long = -1&) As eGDCattleCustomerType
On Error GoTo ErrSection:

    Dim nType As eGDCattleCustomerType  ' Type to return

    nType = -1&
    If CurrentTab = eGDTab_Customers Then
        If cboType.ListIndex >= 0 Then
            nType = cboType.ItemData(cboType.ListIndex)
            
            If bConvert Then
                Select Case nType
                    Case eGDCattleCustomerType_CattleNavCustomer
                        nType = eGDCattleCustomerType_TurnkeyCustomer
                    Case eGDCattleCustomerType_CattleNavFeedYard
                        nType = eGDCattleCustomerType_TurnkeyFeedYard
                    Case eGDCattleCustomerType_EitherFeedYard
                        nType = eGDCattleCustomerType_TurnkeyFeedYard
                End Select
            End If
        End If
    Else
        If lRow = -1& Then
            lRow = fgCustomerTypes.Row
        End If
        If ValidGridRow(fgCustomerTypes, lRow) Then
            nType = CLng(Val(fgCustomerTypes.TextMatrix(lRow, 0)))
        End If
    End If
    
    SelectedType = nType
    
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "frmCattleAdmin.SelectedType.Get"
    
End Property

Private Property Get FeedyardFromGrid(ByVal lRow As Long) As cBrokerMessage
On Error GoTo ErrSection:

    Dim FeedYard As cBrokerMessage      ' Feed yard from the grid
    Dim lParent As Long                 ' Parent row in the grid
    
    Set FeedYard = Nothing
    If CurrentTab = eGDTab_Customers Then
        If ValidGridRow(fgYardInfo, lRow) Then
            With fgYardInfo
                If .RowOutlineLevel(lRow) = 1 Then
                    If TypeOf .RowData(lRow) Is cBrokerMessage Then
                        Set FeedYard = .RowData(lRow)
                    End If
                ElseIf .RowOutlineLevel(lRow) = 2 Then
                    lParent = .GetNodeRow(lRow, flexNTParent)
                    If TypeOf .RowData(lParent) Is cBrokerMessage Then
                        Set FeedYard = .RowData(lParent)
                    End If
                End If
            End With
        End If
    End If
    
    Set FeedyardFromGrid = FeedYard

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "frmCattleAdmin.FeedyardFromGrid.Get"
    
End Property

Private Property Get SelectedFeedYard() As cBrokerMessage
    Set SelectedFeedYard = FeedyardFromGrid(fgYardInfo.Row)
End Property

Private Property Get SelectedSource(Optional ByVal lRow As Long = -1&) As eGDCattleFeedYardSource
On Error GoTo ErrSection:

    Dim nType As eGDCattleFeedYardSource ' Feed yard source to return
    Dim FeedYard As cBrokerMessage      ' Feed yard from the grid
    Dim lParent As Long                 ' Parent row in the grid
    
    nType = -1&
    
    If CurrentTab = eGDTab_Customers Then
        If lRow = -1& Then
            Set FeedYard = SelectedFeedYard
        Else
            Set FeedYard = FeedyardFromGrid(lRow)
        End If
        
        If Not FeedYard Is Nothing Then
            nType = CLng(Val(FeedYard("FeedYardSource")))
        End If
    Else
        If lRow = -1& Then
            lRow = fgFeedYardSources.Row
        End If
        
        If ValidGridRow(fgFeedYardSources, lRow) Then
            nType = CLng(Val(fgFeedYardSources.TextMatrix(lRow, 0)))
        End If
    End If
    
    SelectedSource = nType

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "frmCattleAdmin.SelectedSource.Get"
    
End Property

Private Property Get SelectedFeedyardID() As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return type for the function
    Dim FeedYard As cBrokerMessage      ' Feed yard from the grid
    Dim lParent As Long                 ' Parent row in the grid
    
    strReturn = ""
    If CurrentTab = eGDTab_Customers Then
        Set FeedYard = SelectedFeedYard
        If Not FeedYard Is Nothing Then
            strReturn = FeedYard("ID")
        End If
    End If
    
    SelectedFeedyardID = strReturn

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "frmCattleAdmin.SelectedFeedyardID.Get"
    
End Property

Private Property Get DefaultColumnsKey(ByVal strType As String, ByVal strSource As String) As String
    DefaultColumnsKey = strType & "|" & strSource
End Property

Private Property Get DefaultColumns(ByVal strType As String, ByVal strSource As String) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim strKey As String                ' Key into the collection

    strReturn = ""
    strKey = DefaultColumnsKey(strType, strSource)
    
    If Not m.DefaultColumns Is Nothing Then
        If m.DefaultColumns.Exists(strKey) Then
            strReturn = m.DefaultColumns(strKey)
        End If
        
        If Len(strReturn) = 0 Then
            If (strType = "3") And (strSource = "1") Then
                strReturn = DefaultColumns("3", "0")
            ElseIf (strType = "3") And (strSource = "2") Then
                strReturn = DefaultColumns("5", "0")
            End If
        End If
    End If
    
    DefaultColumns = strReturn

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "frmCattleAdmin.DefaultColumns.Get"
    
End Property
Private Property Let DefaultColumns(ByVal strType As String, ByVal strSource As String, ByVal strColumns As String)
On Error GoTo ErrSection:

    Dim strKey As String                ' Key into the collection
    
    If Not m.DefaultColumns Is Nothing Then
        strKey = DefaultColumnsKey(strType, strSource)
        
        If m.DefaultColumns.Exists(strKey) Then
            m.DefaultColumns(strKey) = strColumns
        Else
            m.DefaultColumns.Add strColumns, strKey
        End If
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "frmCattleAdmin.DefaultColumns.Get"
    
End Property

Private Property Get VisibleColumnsKey(ByVal strGenesisID As String, ByVal strFeedyardID As String) As String
    VisibleColumnsKey = strGenesisID & "|" & strFeedyardID
End Property

Private Property Get VisibleColumns(ByVal strGenesisID As String, ByVal strFeedyardID As String) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim strKey As String                ' Key into the collection

    strReturn = ""
    strKey = VisibleColumnsKey(strGenesisID, strFeedyardID)
    
    If Not m.VisibleColumns Is Nothing Then
        If m.VisibleColumns.Exists(strKey) Then
            strReturn = m.VisibleColumns(strKey)
        End If
    End If
    
    VisibleColumns = strReturn

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "frmCattleAdmin.VisibleColumns.Get"
    
End Property
Private Property Let VisibleColumns(ByVal strGenesisID As String, ByVal strFeedyardID As String, ByVal strColumns As String)
On Error GoTo ErrSection:

    Dim strKey As String                ' Key into the collection
    
    If Not m.VisibleColumns Is Nothing Then
        strKey = VisibleColumnsKey(strGenesisID, strFeedyardID)
        
        If m.VisibleColumns.Exists(strKey) Then
            m.VisibleColumns(strKey) = strColumns
        Else
            m.VisibleColumns.Add strColumns, strKey
        End If
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "frmCattleAdmin.VisibleColumns.Get"
    
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Setup and show the form
'' Inputs:      None
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowMe()
On Error GoTo ErrSection:

    SetFormCaption
    
    InitCustomerTypesGrid
    LoadCustomerTypesGrid
    InitFeedYardSourcesGrid
    LoadFeedYardSourcesGrid
    InitColumnsGrid fgLotColumns
    
    InitYardInfoGrid
    InitColumnsGrid fgColumns
    
    If Not g.Cattle Is Nothing Then
        Status = g.Cattle.ConnectionStatus
        
        If Status = eGDConnectionStatus_Disconnected Then
            g.Cattle.Connect True
        End If
    End If

    Dirty = False
    m.bChangingCustomers = False

    ShowForm Me, eForm_Modal, g.frmMain
    
ErrExit:
    Unload Me
    Exit Sub
    
ErrSection:
    Unload Me
    RaiseError "frmCattleAdmin.ShowMe"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Cattle_GenesisCustomer
'' Description: Handle a Genesis customer coming from the cattle server
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Cattle_GenesisCustomer(ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim cattleMessage As cBrokerMessage ' Message object
    Dim strID As String                 ' ID for the Genesis customer
    
    Select Case UCase(Parse(strMessage, vbTab, 1))
        Case "BEGIN"
            cboGenesisCustomers.Clear
        
        Case "END"
            GetAllFeedyards
        
        Case Else
            Set cattleMessage = New cBrokerMessage
            cattleMessage.FromString strMessage
            
            strID = cattleMessage("ID")
            cboGenesisCustomers.AddItem cattleMessage("Account") & " (" & cattleMessage("Name") & ")"
            cboGenesisCustomers.ItemData(cboGenesisCustomers.NewIndex) = CLng(Val(strID))
            
            If m.GenesisCustomers.Exists(strID) Then
                Set m.GenesisCustomers(strID) = cattleMessage
            Else
                m.GenesisCustomers.Add cattleMessage, strID
            End If
            
            If (m.lSelectedID = -1&) And (cattleMessage("Account") = txtAccount.Text) Then
                m.lSelectedID = CLng(Val(strID))
                SelectGenesisCustomer m.lSelectedID
            End If
            
    End Select

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmCattleAdmin.Cattle_GenesisCustomer", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Cattle_Feedyard
'' Description: Handle a feed yard coming from the cattle server
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Cattle_FeedYard(ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim cattleMessage As cBrokerMessage ' Message object
    
    Select Case UCase(Parse(strMessage, vbTab, 1))
        Case "BEGIN"
        
        Case "END"
            GetAllFeedyardCustomers
        
        Case Else
            Set cattleMessage = New cBrokerMessage
            cattleMessage.FromString strMessage
            FeedyardToGrid cattleMessage
            
    End Select

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmCattleAdmin.Cattle_Feedyard", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Cattle_FeedyardCustomer
'' Description: Handle a feed yard customer coming from the cattle server
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Cattle_FeedyardCustomer(ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim cattleMessage As cBrokerMessage ' Message object
    
    Select Case UCase(Parse(strMessage, vbTab, 1))
        Case "BEGIN"
        
        Case "END"
            GetAllLotColumns
        
        Case Else
            Set cattleMessage = New cBrokerMessage
            cattleMessage.FromString strMessage
            FeedyardCustomerToGrid cattleMessage
            
    End Select

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmCattleAdmin.Cattle_FeedyardCustomer", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Cattle_LotColumn
'' Description: Handle a lot column coming from the cattle server
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Cattle_LotColumn(ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim LotColumn As cLotColumn         ' Lot column information
    
    Select Case UCase(Parse(strMessage, vbTab, 1))
        Case "BEGIN"
            fgColumns.Rows = fgColumns.FixedRows
            fgLotColumns.Rows = fgLotColumns.FixedRows
        
        Case "END"
            GetDefaultVisibleLotColumns
        
        Case Else
            Set LotColumn = New cLotColumn
            LotColumn.FromString strMessage
            
            If CurrentTab = eGDTab_Customers Then
                LotColumnToGrid fgColumns, LotColumn
            Else
                LotColumnToGrid fgLotColumns, LotColumn
            End If
            
    End Select

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmCattleAdmin.Cattle_LotColumn", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Cattle_VisibleFeedyard
'' Description: Handle a visible feed yard coming from the cattle server
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Cattle_VisibleFeedyard(ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim cattleMessage As cBrokerMessage ' Message object
    Dim lRow As Long                    ' Row in the grid
    
    Select Case UCase(Parse(strMessage, vbTab, 1))
        Case "BEGIN"
            ClearYardInfoCheckBoxes
        
        Case "END"
            GetVisibleFeedyardCustomers
            
        Case Else
            Set cattleMessage = New cBrokerMessage
            cattleMessage.FromString strMessage
            
            lRow = GridRowForFeedyard(cattleMessage("ID"))
            If lRow > -1& Then
                CheckedCell(fgYardInfo, lRow, GDCol(eGDCol_Visible)) = True
            End If
            
    End Select

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmCattleAdmin.Cattle_VisibleFeedyard", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Cattle_VisibleFeedyardCustomer
'' Description: Handle a visible feed yard customer coming from the cattle server
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Cattle_VisibleFeedyardCustomer(ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim cattleMessage As cBrokerMessage ' Message object
    Dim lRow As Long                    ' Row in the grid
    
    Select Case UCase(Parse(strMessage, vbTab, 1))
        Case "BEGIN"
        
        Case "END"
            GetVisibleLotColumns
            If fgYardInfo.Rows > fgYardInfo.FixedRows Then
                fgYardInfo.Row = fgYardInfo.FixedRows
            End If
        
        Case Else
            Set cattleMessage = New cBrokerMessage
            cattleMessage.FromString strMessage
            
            lRow = GridRowForFeedyardCustomer(cattleMessage("FeedYardID"), cattleMessage("ID"))
            If lRow > -1& Then
                CheckedCell(fgYardInfo, lRow, GDCol(eGDCol_Visible)) = True
            End If
            
    End Select

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmCattleAdmin.Cattle_VisibleFeedyardCustomer", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Cattle_VisibleLotColumn
'' Description: Handle a visible lot column coming from the cattle server
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Cattle_VisibleLotColumn(ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim cattleMessage As cBrokerMessage ' Message object
    
    Select Case UCase(Parse(strMessage, vbTab, 1))
        Case "BEGIN"
            m.VisibleColumns.Clear
        
        Case "END"
            If m.bChangingCustomers = True Then
                Dirty = False
                m.bChangingCustomers = False
            End If
            
        Case Else
            Set cattleMessage = New cBrokerMessage
            cattleMessage.FromString strMessage
            
            VisibleColumns(cattleMessage("GenesisID"), cattleMessage("FeedYardID")) = cattleMessage("LotColumnIds")
            
    End Select

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmCattleAdmin.Cattle_VisibleLotColumn", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Cattle_DefaultVisibleLotColumns
'' Description: Handle default visible lot column info coming from the cattle server
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Cattle_DefaultVisibleLotColumns(ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim cattleMessage As cBrokerMessage ' Message object
    
    Select Case UCase(Parse(strMessage, vbTab, 1))
        Case "BEGIN"
        
        Case "END"
            ChangeVisibleLotColumns
        
        Case Else
            Set cattleMessage = New cBrokerMessage
            cattleMessage.FromString strMessage
            
            DefaultColumns(cattleMessage("Type"), cattleMessage("Source")) = cattleMessage("LotColumnIds")
                        
    End Select

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmCattleAdmin.Cattle_DefaultVisibleLotColumn", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboGenesisCustomers_Click
'' Description: Handle the user changing Genesis customers
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboGenesisCustomers_Click()
On Error GoTo ErrSection:

    Dim lSelectedID As Long             ' Selected ID in the combo
    Dim cattleMessage As cBrokerMessage ' Cattle message for the Genesis customer

    If Visible Then
        If cboGenesisCustomers.ListIndex > -1& Then
            lSelectedID = cboGenesisCustomers.ItemData(cboGenesisCustomers.ListIndex)
            If lSelectedID <> m.lSelectedID Then
                If AskToSave Then
                    m.lSelectedID = lSelectedID
                    
                    m.bChangingCustomers = True
                    Set cattleMessage = m.GenesisCustomers(Str(m.lSelectedID))
                    txtAccount.Text = cattleMessage("Account")
                    txtName.Text = cattleMessage("Name")
                    txtPassword.Text = cattleMessage("Password")
                    SelectType cattleMessage("Type")
                    
                    CheckBoxValue(chkCanEditLots) = g.Cattle.StringToBool(cattleMessage("CanEditLots"))
                                    
                    GetVisibleFeedyards
                Else
                    SelectGenesisCustomer m.lSelectedID
                End If
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleAdmin.cboGenesisCustomers_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboType_Click
'' Description: Set the dirty flag when the contents of the control changes
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboType_Click()
On Error GoTo ErrSection:

    Dirty = True
    
    If Visible Then
        Select Case SelectedType
            Case eGDCattleCustomerType_TurnkeyFeedYard
                CheckBoxValue(chkCanEditLots) = True
                chkCanEditLots.Enabled = True
            Case eGDCattleCustomerType_TurnkeyCustomer
                CheckBoxValue(chkCanEditLots) = False
                chkCanEditLots.Enabled = False
            Case Else
                chkCanEditLots.Enabled = True
        End Select
        
        ChangeVisibleLotColumns
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleAdmin.cboType_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkCanEditLots_Click
'' Description: Set the dirty flag when the contents of the control changes
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkCanEditLots_Click()
On Error GoTo ErrSection:

    Dirty = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleAdmin.chkCanEditLots_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgColumns_BeforeMouseDown
'' Description: Handle the user pressing a mouse button in a cell
'' Inputs:      Button Pressed, Shift/Ctrl/Alt status, Location of Mouse,
''              Whether to Cancel the click
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgColumns_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:

    m.iColumnsButton = Button
    If Button = vbRightButton Then
        mnuPopup.Tag = "Columns"
        
        Enable mnuPaste, (Len(m.strCopiedColumns) > 0)
        
        PopupMenu mnuPopup
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleAdmin.fgColumns_BeforeMouseDown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgColumns_Click
'' Description: Allow the user to toggle the visible column
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgColumns_Click()
On Error GoTo ErrSection:

    Dim lMouseRow As Long               ' Row in the grid that the mouse was clicked in
    Dim lMouseCol As Long               ' Column in the grid that the mouse was clicked in
    
    With fgColumns
        lMouseRow = .MouseRow
        lMouseCol = .MouseCol
        
        If m.iColumnsButton = vbLeftButton Then
            If .Cell(flexcpForeColor, lMouseRow, 0) = vbBlack Then
                CheckedCell(fgColumns, lMouseRow, GDCol(eGDCol_Visible)) = Not CheckedCell(fgColumns, lMouseRow, GDCol(eGDCol_Visible))
                Dirty = True
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleAdmin.fgColumns_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgCustomerTypes_BeforeRowColChange
'' Description: Handle a cell change in the grid
'' Inputs:      Old Row and Column, New Row and Column, Cancel?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgCustomerTypes_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    Cancel = Not ChangeVisibleLotColumns(NewRow)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleAdmin.fgCustomerTypes_BeforeRowColChange"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgFeedYardSources_BeforeRowColChange
'' Description: Handle a cell change in the grid
'' Inputs:      Old Row and Column, New Row and Column, Cancel?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgFeedYardSources_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    Cancel = Not ChangeVisibleLotColumns(, NewRow)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleAdmin.fgFeedYardSources_BeforeRowColChange"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgLotColumns_BeforeMouseDown
'' Description: Handle the user pressing a mouse button in a cell
'' Inputs:      Button Pressed, Shift/Ctrl/Alt status, Location of Mouse,
''              Whether to Cancel the click
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgLotColumns_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:

    m.iLotColumnsButton = Button
    If Button = vbRightButton Then
        mnuPopup.Tag = "Columns"
        
        Enable mnuPaste, (Len(m.strCopiedColumns) > 0)
        
        PopupMenu mnuPopup
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleAdmin.fgLotColumns_BeforeMouseDown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgLotColumns_Click
'' Description: Allow the user to toggle the visible column
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgLotColumns_Click()
On Error GoTo ErrSection:

    Dim lMouseRow As Long               ' Row in the grid that the mouse was clicked in
    Dim lMouseCol As Long               ' Column in the grid that the mouse was clicked in
    
    With fgLotColumns
        lMouseRow = .MouseRow
        lMouseCol = .MouseCol
        
        If m.iLotColumnsButton = vbLeftButton Then
            If .Cell(flexcpForeColor, lMouseRow, 0) = vbBlack Then
                CheckedCell(fgLotColumns, lMouseRow, GDCol(eGDCol_Visible)) = Not CheckedCell(fgLotColumns, lMouseRow, GDCol(eGDCol_Visible))
                Dirty = True
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleAdmin.fgLotColumns_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgYardInfo_AfterRowColChange
'' Description: Handle the user changing cells in the yard info grid
'' Inputs:      Old Row and Column, New Row and Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgYardInfo_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    Dim OldFeedYard As cBrokerMessage   ' Feedyard from the old row
    Dim NewFeedYard As cBrokerMessage   ' Feedyard from the new row
    Dim strColumns As String            ' Columns selected from the grid
    Dim strDefaultColumns As String     ' Default columns

    If NewRow <> OldRow Then
        If ValidGridRow(fgYardInfo, NewRow) Then
            If ValidGridRow(fgYardInfo, OldRow) Then
                UpdateVisibleColumns OldRow
                
'                Set OldFeedYard = FeedyardFromGrid(OldRow)
'                Set NewFeedYard = FeedyardFromGrid(NewRow)
'
'                If OldFeedYard("FeedYardSource") <> NewFeedYard("FeedYardSource") Then
'                    strColumns = BuildVisibleLotColumns
'                    strDefaultColumns = DefaultColumns(Str(SelectedType(True)), OldFeedYard("FeedYardSource"))
'
'                    If strColumns = strDefaultColumns Then
'                        VisibleColumns(Str(m.lSelectedID), OldFeedYard("ID")) = ""
'                    Else
'                        VisibleColumns(Str(m.lSelectedID), OldFeedYard("ID")) = strColumns
'                    End If
'                End If
            End If
            
            ChangeVisibleLotColumns
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleAdmin.fgYardInfo_AfterRowColChange"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgYardInfo_BeforeMouseDown
'' Description: Handle the user pressing a mouse button in a cell
'' Inputs:      Button Pressed, Shift/Ctrl/Alt status, Location of Mouse,
''              Whether to Cancel the click
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgYardInfo_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:

    m.iYardsButton = Button
    If Button = vbRightButton Then
        mnuPopup.Tag = "YardInfo"
        PopupMenu mnuPopup
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleAdmin.fgYardInfo_BeforeMouseDown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgYardInfo_Click
'' Description: Allow the user to toggle the visible column
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgYardInfo_Click()
On Error GoTo ErrSection:

    Dim lMouseRow As Long               ' Row in the grid that the mouse was clicked in
    Dim lMouseCol As Long               ' Column in the grid that the mouse was clicked in
    Dim lChildRow As Long               ' Child row in the grid
    Dim strReturn As String             ' Return from an InfBox
    
    With fgYardInfo
        lMouseRow = .MouseRow
        lMouseCol = .MouseCol
        
        If m.iYardsButton = vbLeftButton Then
            If lMouseCol = GDCol(eGDCol_Visible) Then
                CheckedCell(fgYardInfo, lMouseRow, GDCol(eGDCol_Visible)) = Not CheckedCell(fgYardInfo, lMouseRow, GDCol(eGDCol_Visible))
                
                If .RowOutlineLevel(lMouseRow) = 1 Then
                    If CheckedCell(fgYardInfo, lMouseRow, GDCol(eGDCol_Visible)) = True Then
                        strReturn = InfBox("Do you want to turn all of the customers on for this feed yard as well?", "?", "+Yes|-No", "Confirmation")
                    Else
                        strReturn = "Y"
                    End If
                    
                    If strReturn = "Y" Then
                        lChildRow = .GetNodeRow(lMouseRow, flexNTFirstChild)
                        Do While lChildRow <> -1&
                            CheckedCell(fgYardInfo, lChildRow, GDCol(eGDCol_Visible)) = CheckedCell(fgYardInfo, lMouseRow, GDCol(eGDCol_Visible))
                            
                            lChildRow = .GetNodeRow(lChildRow, flexNTNextSibling)
                        Loop
                    End If
                End If
                
                Dirty = True
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleAdmin.fgYardInfo_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    imgStatus_Click
'' Description: Allow the user to toggle the connection status
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub imgStatus_Click()
On Error GoTo ErrSection:

    ToggleConnection

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleAdmin.imgStatus_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    lblStatus_Click
'' Description: Allow the user to toggle the connection status
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub lblStatus_Click()
On Error GoTo ErrSection:

    ToggleConnection

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleAdmin.lblStatus_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Setup the form when it is loaded
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Icon = g.AppBridge.Picture16(g.Cattle.IconName)
    
    g.Styler.StyleForm Me
    
    PlaceForm Me

    With tbToolbar
        .Tools("ID_New").Picture = g.AppBridge.Picture16(g.AppBridge.ToolbarIcon("kChartNew"))
        .Tools("ID_Save").Picture = g.AppBridge.Picture16(g.AppBridge.ToolbarIcon("kSave"))
    End With
    
    With cboType
        .AddItem "Administrator"
        .ItemData(.NewIndex) = eGDCattleCustomerType_Admin
        .AddItem "Broker"
        .ItemData(.NewIndex) = eGDCattleCustomerType_Broker
        .AddItem "Feed Yard"
        .ItemData(.NewIndex) = eGDCattleCustomerType_TurnkeyFeedYard
        .AddItem "Customer"
        .ItemData(.NewIndex) = eGDCattleCustomerType_TurnkeyCustomer
    End With
    
    m.lSelectedID = -1&
    Set m.GenesisCustomers = New cGdTree
    Set m.DefaultColumns = New cGdTree
    Set m.VisibleColumns = New cGdTree
    
    mnuPopup.Visible = False
    
    Dirty = False
    CurrentTab = eGDTab_LotColumns

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleAdmin.Form_Load"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: Cancel the unload and let the ShowMe handle it
'' Inputs:      Whether to cancel the unload, Mode of the unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode <> vbFormCode Then
        Cancel = True
        
        If AskToSave Then
            Hide
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleAdmin.Form_QueryUnload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Resize
'' Description: Resize and move controls as the form is resized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Resize()
On Error Resume Next

    Dim lMinScaleWidth As Long          ' Minimum scale width
    Dim lMinScaleHeight As Long         ' Minimum scale height
    Dim lSpace As Long                  ' Space between controls
    Dim lTop As Long                    ' Top of the grid
    Dim lGridWidth As Long              ' Grid width
    Dim lLeft As Long                   ' Left of the control

    lSpace = 120
    lMinScaleHeight = 3765
    lMinScaleWidth = 10965

    If Not LimitFormSize(Me, lMinScaleWidth, lMinScaleHeight) Then
        With tabAdmin
            .Move lSpace, .Top, ScaleWidth - (lSpace * 2), ScaleHeight - .Top - lSpace
        End With
        
        lGridWidth = (tabAdmin.ClientWidth - (lSpace * 4)) / 3
        lTop = lSpace
        
        With fgCustomerTypes
            .Move lSpace, lSpace, lGridWidth, tabAdmin.ClientHeight - (lSpace * 2)
        End With
        
        With fgFeedYardSources
            .Move fgCustomerTypes.Width + (lSpace * 2), lSpace, lGridWidth, tabAdmin.ClientHeight - (lSpace * 2)
        End With
        
        With fgLotColumns
            .Move fgFeedYardSources.Left + fgFeedYardSources.Width + lSpace, lSpace, lGridWidth, tabAdmin.ClientHeight - (lSpace * 2)
        End With
        
        lGridWidth = (tabAdmin.ClientWidth - fraCustomerInfo.Width - (lSpace * 4)) / 2
        lTop = fraCustomerInfo.Top
        
        With fgYardInfo
            lLeft = fraCustomerInfo.Left + fraCustomerInfo.Width + lSpace
            .Move lLeft, lTop, lGridWidth, tabAdmin.ClientHeight - lTop - lSpace
        End With
        
        With fgColumns
            lLeft = fgYardInfo.Left + fgYardInfo.Width + lSpace
            .Move lLeft, lTop, lGridWidth, tabAdmin.ClientHeight - lTop - lSpace
        End With
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: Clean up when the form is unloaded
'' Inputs:      Cancel the Unload?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    If (Not g.Cattle Is Nothing) And (Status <> eGDConnectionStatus_Disconnected) Then
        If Not FormIsLoaded("frmLots") Then
            g.Cattle.Disconnect
        End If
    End If

    SaveFormPlacement Me
    
    Set m.GenesisCustomers = Nothing
    Set m.DefaultColumns = Nothing
    Set m.VisibleColumns = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleAdmin.Form_Unload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuCopy_Click
'' Description: Copy the columns
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuCopy_Click()
On Error GoTo ErrSection:

    m.strCopiedColumns = BuildVisibleLotColumns

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleAdmin.mnuCopy_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuDeselectAll_Click
'' Description: Deselect all items in the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuDeselectAll_Click()
On Error GoTo ErrSection:

    If UCase(mnuPopup.Tag) = "YARDINFO" Then
        ClearYardInfoCheckBoxes
    ElseIf UCase(mnuPopup.Tag) = "COLUMNS" Then
        If CurrentTab = eGDTab_Customers Then
            ClearVisibleColumnsCheckBoxes fgColumns
        Else
            ClearVisibleColumnsCheckBoxes fgLotColumns
        End If
    End If
    
    Dirty = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleAdmin.mnuDeselectAll_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuPaste_Click
'' Description: Paste the copied columns
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuPaste_Click()
On Error GoTo ErrSection:

    Dim strColumns As String            ' Current columns

    If Len(m.strCopiedColumns) > 0 Then
        strColumns = BuildVisibleLotColumns
        
        If m.strCopiedColumns <> strColumns Then
            If CurrentTab = eGDTab_Customers Then
                ClearVisibleColumnsCheckBoxes fgColumns
                SetLotColumns fgColumns, m.strCopiedColumns
                Dirty = True
            Else
                ClearVisibleColumnsCheckBoxes fgLotColumns
                SetLotColumns fgLotColumns, m.strCopiedColumns
                Dirty = True
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleAdmin.mnuPaste_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuSelectAll_Click
'' Description: Select all items in the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuSelectAll_Click()
On Error GoTo ErrSection:

    Dim Grid As VSFlexGrid              ' Grid to select all
    Dim lIndex As Long                  ' Index into a for loop

    If UCase(mnuPopup.Tag) = "YARDINFO" Then
        Set Grid = fgYardInfo
    ElseIf UCase(mnuPopup.Tag) = "COLUMNS" Then
        If CurrentTab = eGDTab_Customers Then
            Set Grid = fgColumns
        Else
            Set Grid = fgLotColumns
        End If
    End If
    
    With Grid
        .Redraw = flexRDNone
        
        For lIndex = .FixedRows To .Rows - 1
            CheckedCell(Grid, lIndex, GDCol(eGDCol_Visible)) = True
        Next lIndex
        
        .Redraw = flexRDBuffered
    End With
    
    Dirty = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleAdmin.mnuSelectAll_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tabAdmin_Switch
'' Description: Handle the user switching tabs
'' Inputs:      Old Tab, New Tab, Cancel?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tabAdmin_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
On Error GoTo ErrSection:

    If AskToSave Then
        If NewTab = GDTab(eGDTab_Customers) Then
            GetGenesisCustomers
        Else
            GetAllLotColumns
        End If
    Else
        Cancel = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleAdmin.tabAdmin_Switch"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tbToolbar_ToolClick
'' Description: Handle the user clicking on a tool on the toolbar
'' Inputs:      Tool Clicked
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tbToolbar_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
On Error GoTo ErrSection:

    Select Case UCase(Tool.ID)
        Case "ID_NEW"
            If AskToSave Then
                ClearForm False
            End If
            
        Case "ID_SAVE"
            Save
        
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleAdmin.tbToolbar_ToolClick"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtAccount_Change
'' Description: Set the dirty flag when the contents of the control changes
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtAccount_Change()
On Error GoTo ErrSection:

    Dirty = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleAdmin.txtAccount_Change"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtAccount_GotFocus
'' Description: Select all of the text when the control gets the focus
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtAccount_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtAccount

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleAdmin.txtAccount_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtName_Change
'' Description: Set the dirty flag when the contents of the control changes
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtName_Change()
On Error GoTo ErrSection:

    Dirty = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleAdmin.txtName_Change"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtName_GotFocus
'' Description: Select all of the text when the control gets the focus
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtName_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtName

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleAdmin.txtName_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtPassword_GotFocus
'' Description: Select all of the text when the control gets the focus
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtPassword_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtPassword

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleAdmin.txtPassword_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtPassword_Change
'' Description: Set the dirty flag when the contents of the control changes
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtPassword_Change()
On Error GoTo ErrSection:

    Dirty = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleAdmin.txtPassword_Change"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitCustomerTypesGrid
'' Description: Initialize the customer types grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitCustomerTypesGrid()
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    
    With fgCustomerTypes
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        mGenesis.SetupGrid fgCustomerTypes, eGridMode_List
        
        .FixedRows = 0
        .Rows = 0
        .FixedCols = 0
        .Cols = 2
        
        .ColHidden(0) = True
        
        .Redraw = nRedraw
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleAdmin.InitCustomerTypesGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitFeedYardSourcesGrid
'' Description: Initialize the feedyard sources grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitFeedYardSourcesGrid()
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    
    With fgFeedYardSources
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        mGenesis.SetupGrid fgFeedYardSources, eGridMode_List
        
        .FixedRows = 0
        .Rows = 0
        .FixedCols = 0
        .Cols = 2
        
        .ColHidden(0) = True
        
        .Redraw = nRedraw
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleAdmin.InitFeedYardSourcesGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitYardInfoGrid
'' Description: Initialize the feedyard information grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitYardInfoGrid()
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    
    With fgYardInfo
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        .AllowBigSelection = False
        .AllowSelection = False
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .Editable = flexEDNone
        .ExplorerBar = flexExSortShow
        .ExtendLastCol = True
        .MergeCells = flexMergeNever
        .OutlineBar = flexOutlineBarSimpleLeaf
        .ScrollTrack = True
        .SelectionMode = flexSelectionListBox
        .SheetBorder = RGB(128, 128, 128)
        
        .Rows = 1
        .FixedRows = 1
        .Cols = GDCol(eGDCol_NumCols)
        .FixedCols = 0
        
        .TextMatrix(0, GDCol(eGDCol_Visible)) = "Visible"
        .TextMatrix(0, GDCol(eGDCol_Name)) = "Name"
        .TextMatrix(0, GDCol(eGDCol_ID)) = "ID"
        
        .ColHidden(GDCol(eGDCol_ID)) = True
        
        .ColAlignment(GDCol(eGDCol_Visible)) = flexAlignCenterCenter
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignLeftCenter
        
        .AutoSize 0, .Cols - 1, False, 75
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleAdmin.InitYardInfoGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitColumnsGrid
'' Description: Initialize the visible columns grid
'' Inputs:      Grid
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitColumnsGrid(Grid As VSFlexGrid)
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    
    With Grid
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        .AllowBigSelection = False
        .AllowSelection = False
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .Editable = flexEDNone
        .ExplorerBar = flexExSortShow
        .ExtendLastCol = True
        .MergeCells = flexMergeNever
        .OutlineBar = flexOutlineBarNone
        .ScrollTrack = True
        .SelectionMode = flexSelectionListBox
        .SheetBorder = RGB(128, 128, 128)
        
        .Rows = 1
        .FixedRows = 1
        .Cols = GDCol(eGDCol_NumCols)
        .FixedCols = 0
        
        .TextMatrix(0, GDCol(eGDCol_Visible)) = "Visible"
        .TextMatrix(0, GDCol(eGDCol_Name)) = "Name"
        .TextMatrix(0, GDCol(eGDCol_ID)) = "ID"
        
        .ColHidden(GDCol(eGDCol_ID)) = True
        
        .ColAlignment(GDCol(eGDCol_Visible)) = flexAlignCenterCenter
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignLeftCenter
        
        .AutoSize 0, .Cols - 1, False, 75
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleAdmin.InitColumnsGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadCustomerTypesGrid
'' Description: Load the customer types grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadCustomerTypesGrid()
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    
    With fgCustomerTypes
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        .Rows = 4
        
        .TextMatrix(0, 0) = Str(eGDCattleCustomerType_TurnkeyCustomer)
        .TextMatrix(0, 1) = "Customer"
        .TextMatrix(1, 0) = Str(eGDCattleCustomerType_Broker)
        .TextMatrix(1, 1) = "Broker"
        .TextMatrix(2, 0) = Str(eGDCattleCustomerType_Admin)
        .TextMatrix(2, 1) = "Administrator"
        .TextMatrix(3, 0) = Str(eGDCattleCustomerType_TurnkeyFeedYard)
        .TextMatrix(3, 1) = "Feedyard"
        
        .Row = 0
        
        .Redraw = nRedraw
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleAdmin.LoadCustomerTypesGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadFeedYardSourcesGrid
'' Description: Load the feedyard sources grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadFeedYardSourcesGrid()
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    
    With fgFeedYardSources
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        .Rows = 3
        
        .TextMatrix(0, 0) = Str(eGDCattleFeedYardSource_Turnkey)
        .TextMatrix(0, 1) = "Turnkey"
        .TextMatrix(1, 0) = Str(eGDCattleFeedYardSource_Manual)
        .TextMatrix(1, 1) = "Manual"
        .TextMatrix(2, 0) = Str(eGDCattleFeedYardSource_ViewTrak)
        .TextMatrix(2, 1) = "ViewTrak"
        
        .Row = 0
        
        .Redraw = nRedraw
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleAdmin.LoadFeedYardSourcesGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DumpDebug
'' Description: Dump the message to the log file
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DumpDebug(ByVal strMessage As String)
On Error GoTo ErrSection:

    If Not g.Cattle Is Nothing Then
        g.Cattle.DumpDebug "Cattle Admin: " & strMessage
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleAdmin.DumpDebug"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetGenesisCustomers
'' Description: Get a list of Genesis customers already setup in the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetGenesisCustomers()
On Error GoTo ErrSection:

    If Not g.Cattle Is Nothing Then
        g.Cattle.GetGenesisCustomers
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleAdmin.GetGenesisCustomers"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetAllFeedyards
'' Description: Get a list of all of the feed yards
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetAllFeedyards()
On Error GoTo ErrSection:

    If Not g.Cattle Is Nothing Then
        g.Cattle.GetAllFeedyards
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleAdmin.GetAllFeedyards"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetAllFeedyardCustomers
'' Description: Get a list of all of the customers
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetAllFeedyardCustomers()
On Error GoTo ErrSection:

    If Not g.Cattle Is Nothing Then
        g.Cattle.GetAllCustomers
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleAdmin.GetAllFeedyardCustomers"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetAllLotColumns
'' Description: Get a list of lot columns
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetAllLotColumns()
On Error GoTo ErrSection:

    If Not g.Cattle Is Nothing Then
        g.Cattle.GetAllLotColumnsAdmin
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleAdmin.GetAllLotColumns"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetVisibleFeedyards
'' Description: Get a list of visible feed yards for the Genesis user
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetVisibleFeedyards()
On Error GoTo ErrSection:

    If cboGenesisCustomers.ListIndex > -1& Then
        If Not g.Cattle Is Nothing Then
            g.Cattle.GetVisibleFeedyards Str(cboGenesisCustomers.ItemData(cboGenesisCustomers.ListIndex))
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleAdmin.GetVisibleFeedyards"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetVisibleFeedyardCustomers
'' Description: Get a list of visible feed yard customers for the Genesis user
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetVisibleFeedyardCustomers()
On Error GoTo ErrSection:

    If cboGenesisCustomers.ListIndex > -1& Then
        If Not g.Cattle Is Nothing Then
            g.Cattle.GetVisibleCustomers Str(cboGenesisCustomers.ItemData(cboGenesisCustomers.ListIndex))
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleAdmin.GetVisibleFeedyardCustomers"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetVisibleLotColumns
'' Description: Get a list of visible lot columns for the feedyard customer
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetVisibleLotColumns()
On Error GoTo ErrSection:
    
    If cboGenesisCustomers.ListIndex > -1& Then
        If Not g.Cattle Is Nothing Then
            g.Cattle.GetVisibleLotColumnsAdmin Str(cboGenesisCustomers.ItemData(cboGenesisCustomers.ListIndex))
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleAdmin.GetVisibleLotColumns"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetDefaultVisibleLotColumns
'' Description: Get a list of default visible lot columns per customer type
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetDefaultVisibleLotColumns()
On Error GoTo ErrSection:

    If Not g.Cattle Is Nothing Then
        g.Cattle.GetDefaultVisibleLotColumnsAdmin
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleAdmin.GetDefaultVisibleLotColumns"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GridRowForFeedyard
'' Description: Determine the grid row for the given Feedyard ID
'' Inputs:      Feedyard ID
'' Returns:     Grid Row (-1 if not found)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GridRowForFeedyard(ByVal strID As String) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim lRow As Long                    ' Row in the grid
    
    lReturn = -1&
    With fgYardInfo
        If .Rows > .FixedRows Then
            lRow = .FixedRows
            Do While lRow <> -1&
                If .TextMatrix(lRow, GDCol(eGDCol_ID)) = strID Then
                    lReturn = lRow
                    Exit Do
                End If
                
                lRow = .GetNodeRow(lRow, flexNTNextSibling)
            Loop
        End If
    End With
    
    GridRowForFeedyard = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmCattleAdmin.GridRowForFeedyard"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GridRowForFeedyardCustomer
'' Description: Determine the grid row for the given Feedyard Customer ID
'' Inputs:      Feedyard ID, Feedyard Customer ID, Feedyard Row
'' Returns:     Grid Row (-1 if not found)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GridRowForFeedyardCustomer(ByVal strFeedyardID As String, ByVal strFeedyardCustomerID As String, Optional lFeedYardRow As Long = -1&) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim lRow As Long                    ' Row in the grid
    
    lReturn = -1&
    With fgYardInfo
        If .Rows > .FixedRows Then
            If lFeedYardRow = -1& Then
                lFeedYardRow = GridRowForFeedyard(strFeedyardID)
            End If
            If lFeedYardRow > -1& Then
                lRow = .GetNodeRow(lFeedYardRow, flexNTFirstChild)
                Do While lRow <> -1&
                    If .TextMatrix(lRow, GDCol(eGDCol_ID)) = strFeedyardCustomerID Then
                        lReturn = lRow
                        Exit Do
                    End If
                    
                    lRow = .GetNodeRow(lRow, flexNTNextSibling)
                Loop
            End If
        End If
    End With
    
    GridRowForFeedyardCustomer = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmCattleAdmin.GridRowForFeedyardCustomer"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GridRowForLotColumn
'' Description: Determine the grid row for the given Lot Column ID
'' Inputs:      Grid, Lot Column ID
'' Returns:     Grid Row (-1 if not found)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GridRowForLotColumn(Grid As VSFlexGrid, ByVal strID As String) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim lRow As Long                    ' Row in the grid
    
    lReturn = -1&
    With Grid
        If .Rows > .FixedRows Then
            For lRow = .FixedRows To .Rows - 1
                If .TextMatrix(lRow, GDCol(eGDCol_ID)) = strID Then
                    lReturn = lRow
                    Exit For
                End If
            Next lRow
        End If
    End With
    
    GridRowForLotColumn = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmCattleAdmin.GridRowForLotColumn"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FeedyardToGrid
'' Description: Add or update the given feed yard information in the grid
'' Inputs:      Feedyard Information
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FeedyardToGrid(ByVal cattleMessage As cBrokerMessage)
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw setting for the grid
    Dim lRow As Long                    ' Row in the grid

    With fgYardInfo
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        lRow = GridRowForFeedyard(cattleMessage("ID"))
        If lRow = -1& Then
            .Rows = .Rows + 1
            lRow = .Rows - 1
        End If
        
        .RowData(lRow) = cattleMessage
        .RowOutlineLevel(lRow) = 1
        .IsSubtotal(lRow) = True
        
        CheckedCell(fgYardInfo, lRow, GDCol(eGDCol_Visible)) = False
        .TextMatrix(lRow, GDCol(eGDCol_Name)) = cattleMessage("Name")
        .TextMatrix(lRow, GDCol(eGDCol_ID)) = cattleMessage("ID")
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleAdmin.FeedyardToGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FeedyardCustomerToGrid
'' Description: Add or update the given feed yard customer information in the grid
'' Inputs:      Feedyard Customer Information
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FeedyardCustomerToGrid(ByVal cattleMessage As cBrokerMessage)
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw setting for the grid
    Dim lRow As Long                    ' Row in the grid
    Dim lFeedYardRow As Long            ' Feed yard row
    Dim lLastChild As Long              ' Last child of the feedyard

    lRow = -1&
    lFeedYardRow = -1&
    
    With fgYardInfo
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        lFeedYardRow = GridRowForFeedyard(cattleMessage("FeedYardID"))
        If lFeedYardRow > -1& Then
            If cattleMessage("ID") = "47" Then
                lRow = lRow
            End If
            lRow = GridRowForFeedyardCustomer(cattleMessage("FeedYardID"), cattleMessage("ID"), lFeedYardRow)
            If lRow = -1& Then
                ' Do this before adding the row because if this is the last feed yard, the last row
                ' gets set as a child of the last feedyard even though the RowOutlineLevel is zero...
                lLastChild = .GetNodeRow(lFeedYardRow, flexNTLastChild)
                
                .Rows = .Rows + 1
                
                If lLastChild = -1& Then
                    lRow = lFeedYardRow + 1&
                Else
                    lRow = lLastChild + 1&
                End If
            
                .RowPosition(.Rows - 1) = lRow
            End If
            
            .RowData(lRow) = cattleMessage
            .RowOutlineLevel(lRow) = 2
            .IsSubtotal(lRow) = True
            
            CheckedCell(fgYardInfo, lRow, GDCol(eGDCol_Visible)) = False
            .TextMatrix(lRow, GDCol(eGDCol_Name)) = cattleMessage("Name")
            .TextMatrix(lRow, GDCol(eGDCol_ID)) = cattleMessage("ID")
        End If
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleAdmin.FeedyardCustomerToGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LotColumnToGrid
'' Description: Add or update the given lot column information in the grid
'' Inputs:      Grid, Lot Column Information
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LotColumnToGrid(Grid As VSFlexGrid, ByVal LotColumn As cLotColumn)
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw setting for the grid
    Dim lRow As Long                    ' Row in the grid

    With Grid
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        lRow = GridRowForLotColumn(Grid, Str(LotColumn.ID))
        If lRow = -1& Then
            .Rows = .Rows + 1
            lRow = .Rows - 1
        End If
        
        .RowData(lRow) = LotColumn
        
        If (UCase(LotColumn.KeyValueField) = "FEEDYARDLOTID") Or (UCase(LotColumn.KeyValueField) = "NUMBER") Then
            CheckedCell(Grid, lRow, GDCol(eGDCol_Visible)) = True
            .Cell(flexcpForeColor, lRow, 0, lRow, .Cols - 1) = RGB(128, 128, 128)
        Else
            CheckedCell(Grid, lRow, GDCol(eGDCol_Visible)) = False
            .Cell(flexcpForeColor, lRow, 0, lRow, .Cols - 1) = vbBlack
        End If
        .TextMatrix(lRow, GDCol(eGDCol_Name)) = LotColumn.KeyValueField ' LotColumn.ColumnHeader
        .TextMatrix(lRow, GDCol(eGDCol_ID)) = Str(LotColumn.ID)
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleAdmin.LotColumnToGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ClearYardInfoCheckBoxes
'' Description: Clear all of the check boxes in the feedyard info grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ClearYardInfoCheckBoxes()
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw setting for the grid
    Dim lRow As Long                    ' Row in the grid
    
    With fgYardInfo
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        For lRow = .FixedRows To .Rows - 1
            CheckedCell(fgYardInfo, lRow, GDCol(eGDCol_Visible)) = False
        Next lRow
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleAdmin.ClearYardInfoCheckBoxes"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ClearVisibleColumnsCheckBoxes
'' Description: Clear all of the check boxes in the visible columns grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ClearVisibleColumnsCheckBoxes(Grid As VSFlexGrid)
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw setting for the grid
    Dim lRow As Long                    ' Row in the grid
    
    With Grid
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        For lRow = .FixedRows To .Rows - 1
            If .Cell(flexcpForeColor, lRow, 0) = vbBlack Then
                CheckedCell(Grid, lRow, GDCol(eGDCol_Visible)) = False
            End If
        Next lRow
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleAdmin.ClearVisibleColumnsCheckBoxes"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SelectGenesisCustomer
'' Description: Select the Genesis customer in the combo box
'' Inputs:      Genesis Customer ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SelectGenesisCustomer(ByVal lGenesisCustomerID As Long)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    For lIndex = 0 To cboGenesisCustomers.ListCount - 1
        If cboGenesisCustomers.ItemData(lIndex) = lGenesisCustomerID Then
            cboGenesisCustomers.ListIndex = lIndex
            Exit For
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleAdmin.SelectGenesisCustomer"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SelectType
'' Description: Select the type in the combo box
'' Inputs:      Type
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SelectType(ByVal strType As String)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    Select Case CLng(Val(strType))
        Case eGDCattleCustomerType_TurnkeyCustomer, eGDCattleCustomerType_CattleNavCustomer
            strType = Str(eGDCattleCustomerType_TurnkeyCustomer)
        
        Case eGDCattleCustomerType_TurnkeyFeedYard, eGDCattleCustomerType_CattleNavFeedYard, eGDCattleCustomerType_EitherFeedYard
            strType = Str(eGDCattleCustomerType_TurnkeyFeedYard)
        
    End Select
    
    For lIndex = 0 To cboType.ListCount - 1
        If Str(cboType.ItemData(lIndex)) = strType Then
            cboType.ListIndex = lIndex
            Exit For
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleAdmin.SelectType"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ToggleConnection
'' Description: Toggle the cattle connection
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ToggleConnection()
On Error GoTo ErrSection:

    Dim lTimeOut As Long                ' Timeout waiting for connection to get disconnected

    If Not g.Cattle Is Nothing Then
        If Status = eGDConnectionStatus_Disconnected Then
            g.Cattle.Connect True
        Else
            g.Cattle.Disconnect
            
            lTimeOut = 0&
            Do While (Status <> eGDConnectionStatus_Disconnected) And (lTimeOut < 30&)
                Sleep 1
                lTimeOut = lTimeOut + 1&
            Loop
            
            If Status <> eGDConnectionStatus_Disconnected Then
                Status = eGDConnectionStatus_Disconnected
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleAdmin.ToggleConnection"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ClearForm
'' Description: Clear the controls on the form
'' Inputs:      Clear Grid?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ClearForm(ByVal bClearGrid As Boolean)
On Error GoTo ErrSection:

    If bClearGrid Then
        fgYardInfo.Rows = fgYardInfo.FixedRows
        cboGenesisCustomers.Clear
        fgColumns.Rows = fgColumns.FixedRows
    Else
        ClearYardInfoCheckBoxes
        
        If CurrentTab = eGDTab_Customers Then
            ClearVisibleColumnsCheckBoxes fgColumns
        Else
            ClearVisibleColumnsCheckBoxes fgLotColumns
        End If
    End If
    
    txtAccount.Text = ""
    txtName.Text = ""
    txtPassword.Text = ""
    cboType.ListIndex = -1&
    m.lSelectedID = -1&
    CheckBoxValue(chkCanEditLots) = False
    
    Dirty = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleAdmin.ClearForm"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    VerifyControls
'' Description: Verify that the controls have valid information
'' Inputs:      None
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function VerifyControls() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    If Len(txtAccount.Text) = 0 Then
        MoveFocus txtAccount
        InfBox "Please enter a Genesis account number", "!", , Caption & " Error"
    ElseIf Len(txtName.Text) = 0 Then
        MoveFocus txtName
        InfBox "Please enter the customers name", "!", , Caption & " Error"
    ElseIf Len(txtPassword.Text) = 0 Then
        MoveFocus txtPassword
        InfBox "Please enter the customers Genesis password", "!", , Caption & " Error"
    ElseIf cboType.ListIndex = -1 Then
        MoveFocus cboType
        InfBox "Please specify the customer type", "!", , Caption & " Error"
    Else
        bReturn = True
    End If
    
    VerifyControls = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmCattleAdmin.VerifyControls"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Save
'' Description: Save the information
'' Inputs:      None
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function Save() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    If CurrentTab = eGDTab_Customers Then
        bReturn = SaveCustomer
    Else
        bReturn = SaveLotColumns
    End If
    
    Save = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmCattleAdmin.Save"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SaveCustomer
'' Description: Save the customer information
'' Inputs:      None
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SaveCustomer() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim strGenesisID As String          ' Genesis ID
    Dim strOldGenesisID As String       ' Old Genesis ID
    Dim lTimeOut As Long                ' Timeout value
    Dim cattleMessage As cBrokerMessage ' Turnkey message
    Dim lIndex As Long                  ' Index into a for loop
    Dim strLotColumns As String         ' Visible lot columns
    Dim strKey As String                ' Key into a collection
    Dim strFeedyardID As String         ' Feedyard ID
    
    bReturn = False
    If VerifyControls Then
        UpdateVisibleColumns
        
        If m.lSelectedID >= 0 Then
            strGenesisID = Str(m.lSelectedID)
        Else
            strGenesisID = ""
        End If
        strOldGenesisID = strGenesisID
        
        g.Cattle.UpdateGenesisCustomer strGenesisID, Trim(txtAccount.Text), Trim(txtName.Text), Trim(txtPassword.Text), SelectedType, CheckBoxValue(chkCanEditLots)
        
        If m.lSelectedID = -1& Then
            lTimeOut = 0&
            Do While (m.lSelectedID = -1&) And (lTimeOut < 30&)
                Sleep 1
                lTimeOut = lTimeOut + 1&
            Loop
        End If
        
        If m.lSelectedID >= 0& Then
            If Len(strGenesisID) > 0 Then
                Set cattleMessage = m.GenesisCustomers(strGenesisID)
                cattleMessage("Account") = Trim(txtAccount.Text)
                cattleMessage("Name") = Trim(txtName.Text)
                cattleMessage("Password") = Trim(txtPassword.Text)
                cattleMessage("Type") = Str(SelectedType)
                cattleMessage("CanEditLots") = g.Cattle.BoolToString(CheckBoxValue(chkCanEditLots))
                
                Set m.GenesisCustomers(strGenesisID) = cattleMessage
            Else
                strGenesisID = Str(m.lSelectedID)
            End If
            
            g.Cattle.UpdateVisibleFeedyards strGenesisID, BuildVisibleFeedyards
            g.Cattle.UpdateVisibleCustomers strGenesisID, BuildVisibleFeedyardCustomers
            
            For lIndex = 1 To m.VisibleColumns.Count
                If Len(m.VisibleColumns(lIndex)) > 0 Then
                    strKey = m.VisibleColumns.Key(lIndex)
                    strFeedyardID = Parse(strKey, "|", 2)
                    
                    If Parse(strKey, "|", 1) = strOldGenesisID Then
                        strKey = strGenesisID & "|" & strFeedyardID
                        m.VisibleColumns.Key(lIndex) = strKey
                    End If
                    
                    g.Cattle.UpdateVisibleLotColumns strGenesisID, strFeedyardID, m.VisibleColumns(lIndex)
                End If
            Next lIndex
            
            Dirty = False
            bReturn = True
        End If
    End If
    
    SaveCustomer = bReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmCattleAdmin.SaveCustomer"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SaveLotColumns
'' Description: Save the information
'' Inputs:      None
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SaveLotColumns() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim nType As eGDCattleCustomerType  ' Customer type selected
    Dim nSource As eGDCattleFeedYardSource ' Source selected
    
    bReturn = True
    
    nType = SelectedType
    nSource = SelectedSource
    
    If (nType <> -1&) And (nSource <> -1&) Then
        g.Cattle.UpdateDefaultVisibleLotColumns Str(nType), Str(nSource), BuildVisibleLotColumns
        Dirty = False
    End If
    
    SaveLotColumns = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmCattleAdmin.SaveLotColumns"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AskToSave
'' Description: See if we need to prompt the user to save changes
'' Inputs:      None
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AskToSave() As Boolean
On Error GoTo ErrSection:
    
    Dim bReturn As Boolean              ' Return value for the function
    Dim strResponse As String           ' Response from the InfBox
    
    bReturn = True
    If Dirty Then
        strResponse = InfBox("Do you want to save your changes?", "?", "+Yes|No|-Cancel", Caption)
        Select Case strResponse
            Case "C"
                bReturn = False
            
            Case "Y"
                bReturn = Save
            
            Case "N"
            
        End Select
    End If
    
    AskToSave = bReturn
        
ErrExit:
    Exit Function

ErrSection:
    AskToSave = True
    RaiseError "frmCattleAdmin.AskToSave"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BuildVisibleFeedyards
'' Description: Build a list of the visible feedyards for the customer
'' Inputs:      None
'' Returns:     List of Feedyards
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function BuildVisibleFeedyards() As String
On Error GoTo ErrSection:

    Dim astrReturn As cGdArray          ' Return value for the function
    Dim lRow As Long                    ' Row in the grid
    
    Set astrReturn = New cGdArray
    astrReturn.Create eGDARRAY_Strings
    
    With fgYardInfo
        For lRow = .FixedRows To .Rows - 1
            If .RowOutlineLevel(lRow) = 1 Then
                If CheckedCell(fgYardInfo, lRow, GDCol(eGDCol_Visible)) Then
                    astrReturn.Add .TextMatrix(lRow, GDCol(eGDCol_ID))
                End If
            End If
        Next lRow
    End With
    
    BuildVisibleFeedyards = astrReturn.JoinFields("|")

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmCattleAdmin.BuildVisibleFeedyards"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BuildVisibleFeedyardCustomers
'' Description: Build a list of the visible feedyard customers for the customer
'' Inputs:      None
'' Returns:     List of Feedyard Customers
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function BuildVisibleFeedyardCustomers() As String
On Error GoTo ErrSection:

    Dim astrReturn As cGdArray          ' Return value for the function
    Dim lRow As Long                    ' Row in the grid
    
    Set astrReturn = New cGdArray
    astrReturn.Create eGDARRAY_Strings
    
    With fgYardInfo
        For lRow = .FixedRows To .Rows - 1
            If .RowOutlineLevel(lRow) = 2 Then
                If CheckedCell(fgYardInfo, lRow, GDCol(eGDCol_Visible)) Then
                    astrReturn.Add .TextMatrix(lRow, GDCol(eGDCol_ID))
                End If
            End If
        Next lRow
    End With
    
    BuildVisibleFeedyardCustomers = astrReturn.JoinFields("|")

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmCattleAdmin.BuildVisibleFeedyardCustomers"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BuildVisibleLotColumns
'' Description: Build a list of the visible lot columns for the customer
'' Inputs:      None
'' Returns:     List of Lot Columns
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function BuildVisibleLotColumns() As String
On Error GoTo ErrSection:

    Dim astrReturn As cGdArray          ' Return value for the function
    Dim lRow As Long                    ' Row in the grid
    Dim Grid As VSFlexGrid              ' Grid to pull information from
    
    Set astrReturn = New cGdArray
    astrReturn.Create eGDARRAY_Strings
    
    If CurrentTab = eGDTab_Customers Then
        Set Grid = fgColumns
    Else
        Set Grid = fgLotColumns
    End If
    
    With Grid
        For lRow = .FixedRows To .Rows - 1
            If CheckedCell(Grid, lRow, GDCol(eGDCol_Visible)) Then
                astrReturn.Add .TextMatrix(lRow, GDCol(eGDCol_ID))
            End If
        Next lRow
    End With
    
    BuildVisibleLotColumns = astrReturn.JoinFields(",")

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmCattleAdmin.BuildVisibleLotColumns"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetFormCaption
'' Description: Set the form caption
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetFormCaption()
On Error GoTo ErrSection:

    Dim strMode As String               ' Connection mode
    
    strMode = ""
    If g.Cattle.Mode = "D" Then
        strMode = "Demo "
    ElseIf g.Cattle.Mode = "L" Then
        strMode = "Live "
    ElseIf g.Cattle.Mode = "T" Then
        strMode = "Test "
    End If

    Caption = strMode & g.Cattle.ProductName & " Administration"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleAdmin.SetFormCaption"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetLotColumns
'' Description: Set the visible lot columns
'' Inputs:      Lot Column Ids
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetLotColumns(Grid As VSFlexGrid, ByVal strLotColumnIds As String)
On Error GoTo ErrSection:

    Dim astrLotColumnIds As cGdArray    ' Lot column ID's
    Dim lIndex As Long                  ' Index into the for loop
    Dim lRow As Long                    ' Row in the grid

    Set astrLotColumnIds = New cGdArray
    astrLotColumnIds.SplitFields strLotColumnIds, ","
    
    For lIndex = 0 To astrLotColumnIds.Size - 1
        lRow = GridRowForLotColumn(Grid, astrLotColumnIds(lIndex))
        If lRow > -1& Then
            CheckedCell(Grid, lRow, GDCol(eGDCol_Visible)) = True
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleAdmin.SetLotColumns"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ChangeVisibleLotColumns
'' Description: Change the visible lot columns based on selected type and source
'' Inputs:      Row
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ChangeVisibleLotColumns(Optional ByVal lTypeRow As Long = -1&, Optional ByVal lSourceRow As Long = -1&) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim nType As eGDCattleCustomerType  ' Customer type selected
    Dim nSource As eGDCattleFeedYardSource ' Source selected
    Dim strColumnIds As String          ' Column ID's
    Dim strFeedyardID As String         ' Feedyard ID selected
    
    bReturn = True
    nType = SelectedType(False, lTypeRow)
    nSource = SelectedSource(lSourceRow)
    
    If CurrentTab = eGDTab_Customers Then
        strFeedyardID = SelectedFeedyardID
        If (nType <> -1&) And (nSource <> -1&) And (Len(strFeedyardID) > 0) And (m.lSelectedID <> -1&) Then
            strColumnIds = VisibleColumns(Str(m.lSelectedID), strFeedyardID)
            If Len(strColumnIds) = 0 Then
                strColumnIds = DefaultColumns(Str(nType), Str(nSource))
            End If
            
            If Len(strColumnIds) > 0 Then
                ClearVisibleColumnsCheckBoxes fgColumns
                SetLotColumns fgColumns, strColumnIds
            End If
        End If
    Else
        If (nType <> -1&) And (nSource <> -1&) Then
            If AskToSave Then
                ClearVisibleColumnsCheckBoxes fgLotColumns
                
                strColumnIds = DefaultColumns(Str(nType), Str(nSource))
                If Len(strColumnIds) > 0 Then
                    SetLotColumns fgLotColumns, strColumnIds
                End If
            Else
                bReturn = False
            End If
        End If
    End If
    
    ChangeVisibleLotColumns = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmCattleAdmin.ChangeVisibleLotColumns"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateVisibleLotColumns
'' Description: Update the visible lot columns based on selected type and source
'' Inputs:      Row
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UpdateVisibleColumns(Optional ByVal lRow As Long = -1&)
On Error GoTo ErrSection:

    Dim FeedYard As cBrokerMessage      ' Feedyard object
    Dim strColumns As String            ' Columns selected from the grid
    Dim strDefaultColumns As String     ' Default columns

    If lRow = -1& Then
        lRow = fgYardInfo.Row
    End If
    
    If ValidGridRow(fgYardInfo, lRow) = True Then
        Set FeedYard = FeedyardFromGrid(lRow)
        If Not FeedYard Is Nothing Then
            strColumns = BuildVisibleLotColumns
            strDefaultColumns = DefaultColumns(Str(SelectedType(True)), FeedYard("FeedYardSource"))
            
            If strColumns = strDefaultColumns Then
                VisibleColumns(Str(m.lSelectedID), FeedYard("ID")) = ""
            Else
                VisibleColumns(Str(m.lSelectedID), FeedYard("ID")) = strColumns
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleAdmin.UpdateVisibleColumns"
    
End Sub

