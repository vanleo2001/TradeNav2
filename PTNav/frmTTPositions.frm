VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{3B008041-905A-11D1-B4AE-444553540000}#1.0#0"; "Vsocx6.ocx"
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmTTPositions 
   Caption         =   "Form1"
   ClientHeight    =   6720
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   12720
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   12720
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picPbo 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   10680
      Picture         =   "frmTTPositions.frx":0000
      ScaleHeight     =   255
      ScaleWidth      =   1875
      TabIndex        =   3
      Top             =   6360
      Width           =   1875
   End
   Begin VB.PictureBox picRithmic 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   120
      Picture         =   "frmTTPositions.frx":050E
      ScaleHeight     =   345
      ScaleWidth      =   1995
      TabIndex        =   5
      Top             =   6300
      Width           =   1995
   End
   Begin VB.Timer tmrMenu 
      Left            =   11640
      Top             =   5760
   End
   Begin VB.Timer tmrBrokers 
      Left            =   12120
      Top             =   5760
   End
   Begin vsOcx6LibCtl.vsIndexTab tabPositions 
      Height          =   6075
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   10716
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
      FrontTabForeColor=   -2147483635
      Caption         =   "Acco&unt|&Orders|T&ransactions|&Trades|Po&sitions|&Activity Log"
      Align           =   0
      Appearance      =   1
      CurrTab         =   0
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
      Begin HexUniControls.ctlUniFrameWL fraAccountInfo 
         Height          =   5700
         Left            =   45
         TabIndex        =   1
         Top             =   330
         Width           =   10425
         _ExtentX        =   18389
         _ExtentY        =   10054
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmTTPositions.frx":079D
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmTTPositions.frx":07BD
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTTPositions.frx":07DD
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniFrameWL fraCommission 
            Height          =   915
            Left            =   180
            TabIndex        =   14
            Top             =   2340
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
            Caption         =   "frmTTPositions.frx":07F9
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmTTPositions.frx":0841
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTTPositions.frx":0861
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniTextBoxXP txtStockFees 
               Height          =   285
               Left            =   780
               TabIndex        =   16
               Top             =   240
               Width           =   1035
               _ExtentX        =   0
               _ExtentY        =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Enabled         =   -1  'True
               Locked          =   0   'False
               Text            =   "frmTTPositions.frx":087D
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
               Tip             =   "frmTTPositions.frx":089D
               HideSelection   =   -1  'True
               RightToLeft     =   0   'False
               ManualStart     =   0   'False
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTTPositions.frx":08BD
            End
            Begin HexUniControls.ctlUniTextBoxXP txtFutureFees 
               Height          =   285
               Left            =   780
               TabIndex        =   19
               Top             =   525
               Width           =   1035
               _ExtentX        =   0
               _ExtentY        =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Enabled         =   -1  'True
               Locked          =   0   'False
               Text            =   "frmTTPositions.frx":08D9
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
               Tip             =   "frmTTPositions.frx":08F9
               HideSelection   =   -1  'True
               RightToLeft     =   0   'False
               ManualStart     =   0   'False
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTTPositions.frx":0919
            End
            Begin HexUniControls.ctlUniLabelXP lblStockFees 
               Height          =   255
               Left            =   120
               Top             =   255
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
               Caption         =   "frmTTPositions.frx":0935
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmTTPositions.frx":0963
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTTPositions.frx":0983
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblStockFeesDesc 
               Height          =   255
               Left            =   1920
               Top             =   255
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
               Caption         =   "frmTTPositions.frx":099F
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmTTPositions.frx":09E3
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTTPositions.frx":0A03
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblFutureFeesDesc 
               Height          =   255
               Left            =   1920
               Top             =   540
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
               Caption         =   "frmTTPositions.frx":0A1F
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmTTPositions.frx":0A69
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTTPositions.frx":0A89
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblFutureFees 
               Height          =   255
               Left            =   120
               Top             =   540
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
               Caption         =   "frmTTPositions.frx":0AA5
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmTTPositions.frx":0AD5
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTTPositions.frx":0AF5
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
         End
         Begin HexUniControls.ctlUniFrameWL fraFillMatch 
            Height          =   1335
            Left            =   180
            TabIndex        =   25
            Top             =   4200
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
            Caption         =   "frmTTPositions.frx":0B11
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmTTPositions.frx":0B4B
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTTPositions.frx":0B6B
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniRadioXP optLifo 
               Height          =   255
               Left            =   1740
               TabIndex        =   28
               Top             =   900
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
               Caption         =   "frmTTPositions.frx":0B87
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmTTPositions.frx":0BCD
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmTTPositions.frx":0BED
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniRadioXP optFifo 
               Height          =   255
               Left            =   180
               TabIndex        =   27
               Top             =   900
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
               Caption         =   "frmTTPositions.frx":0C09
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmTTPositions.frx":0C51
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmTTPositions.frx":0C71
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblFillMatch 
               Height          =   615
               Left            =   120
               Top             =   240
               Width           =   3375
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
               Caption         =   "frmTTPositions.frx":0C8D
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmTTPositions.frx":0DA9
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTTPositions.frx":0DC9
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
         End
         Begin HexUniControls.ctlUniFrameWL fraTotals 
            Height          =   3015
            Left            =   4500
            TabIndex        =   29
            Top             =   120
            Width           =   4095
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
            Caption         =   "frmTTPositions.frx":0DE5
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmTTPositions.frx":0E11
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTTPositions.frx":0E31
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniTextBoxXP txtCurrentValue 
               Height          =   285
               Left            =   1620
               TabIndex        =   49
               Top             =   2700
               Width           =   2235
               _ExtentX        =   0
               _ExtentY        =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Enabled         =   -1  'True
               Locked          =   0   'False
               Text            =   "frmTTPositions.frx":0E4D
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
               Tip             =   "frmTTPositions.frx":0E6D
               HideSelection   =   -1  'True
               RightToLeft     =   0   'False
               ManualStart     =   0   'False
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTTPositions.frx":0E8D
            End
            Begin HexUniControls.ctlUniTextBoxXP txtTotalOpenEquity 
               Height          =   285
               Left            =   1620
               TabIndex        =   46
               Top             =   2040
               Width           =   2235
               _ExtentX        =   0
               _ExtentY        =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Enabled         =   -1  'True
               Locked          =   0   'False
               Text            =   "frmTTPositions.frx":0EA9
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
               Tip             =   "frmTTPositions.frx":0EC9
               HideSelection   =   -1  'True
               RightToLeft     =   0   'False
               ManualStart     =   0   'False
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTTPositions.frx":0EE9
            End
            Begin HexUniControls.ctlUniTextBoxXP txtClosedBalance 
               Height          =   285
               Left            =   1620
               TabIndex        =   43
               Top             =   1620
               Width           =   2235
               _ExtentX        =   0
               _ExtentY        =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Enabled         =   -1  'True
               Locked          =   0   'False
               Text            =   "frmTTPositions.frx":0F05
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
               Tip             =   "frmTTPositions.frx":0F25
               HideSelection   =   -1  'True
               RightToLeft     =   0   'False
               ManualStart     =   0   'False
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTTPositions.frx":0F45
            End
            Begin HexUniControls.ctlUniTextBoxXP txtTotalClosedProfit 
               Height          =   285
               Left            =   1620
               TabIndex        =   40
               Top             =   1080
               Width           =   2235
               _ExtentX        =   0
               _ExtentY        =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Enabled         =   -1  'True
               Locked          =   0   'False
               Text            =   "frmTTPositions.frx":0F61
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
               Tip             =   "frmTTPositions.frx":0F81
               HideSelection   =   -1  'True
               RightToLeft     =   0   'False
               ManualStart     =   0   'False
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTTPositions.frx":0FA1
            End
            Begin HexUniControls.ctlUniTextBoxXP txtTotalFees 
               Height          =   285
               Left            =   1620
               TabIndex        =   37
               Top             =   720
               Width           =   2235
               _ExtentX        =   0
               _ExtentY        =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Enabled         =   -1  'True
               Locked          =   0   'False
               Text            =   "frmTTPositions.frx":0FBD
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
               Tip             =   "frmTTPositions.frx":0FDD
               HideSelection   =   -1  'True
               RightToLeft     =   0   'False
               ManualStart     =   0   'False
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTTPositions.frx":0FFD
            End
            Begin HexUniControls.ctlUniTextBoxXP txtTotalAdjustments 
               Height          =   285
               Left            =   1620
               TabIndex        =   34
               Top             =   360
               Width           =   2235
               _ExtentX        =   0
               _ExtentY        =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Enabled         =   -1  'True
               Locked          =   0   'False
               Text            =   "frmTTPositions.frx":1019
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
               Tip             =   "frmTTPositions.frx":1039
               HideSelection   =   -1  'True
               RightToLeft     =   0   'False
               ManualStart     =   0   'False
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTTPositions.frx":1059
            End
            Begin HexUniControls.ctlUniTextBoxXP txtStartingBalance 
               Height          =   285
               Left            =   1620
               TabIndex        =   31
               Top             =   0
               Width           =   2235
               _ExtentX        =   0
               _ExtentY        =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Enabled         =   -1  'True
               Locked          =   0   'False
               Text            =   "frmTTPositions.frx":1075
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
               Tip             =   "frmTTPositions.frx":1095
               HideSelection   =   -1  'True
               RightToLeft     =   0   'False
               ManualStart     =   0   'False
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTTPositions.frx":10B5
            End
            Begin HexUniControls.ctlUniLabelXP lblCurrentValue 
               Height          =   195
               Left            =   240
               Top             =   2760
               Width           =   1335
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
               Caption         =   "frmTTPositions.frx":10D1
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmTTPositions.frx":110D
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTTPositions.frx":112D
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblEqualsCurrentValue 
               Height          =   255
               Left            =   0
               Top             =   2700
               Width           =   195
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
               Caption         =   "frmTTPositions.frx":1149
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmTTPositions.frx":116B
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTTPositions.frx":118B
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblTotalOpenEquity 
               Height          =   195
               Left            =   240
               Top             =   2100
               Width           =   1335
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
               Caption         =   "frmTTPositions.frx":11A7
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmTTPositions.frx":11EB
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTTPositions.frx":120B
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblPlusOpenEquity 
               Height          =   255
               Left            =   0
               Top             =   2040
               Width           =   195
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
               Caption         =   "frmTTPositions.frx":1227
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmTTPositions.frx":1249
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTTPositions.frx":1269
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblClosedBalance 
               Height          =   195
               Left            =   240
               Top             =   1680
               Width           =   1335
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
               Caption         =   "frmTTPositions.frx":1285
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmTTPositions.frx":12C3
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTTPositions.frx":12E3
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblEqualsClosedBalance 
               Height          =   255
               Left            =   0
               Top             =   1620
               Width           =   195
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
               Caption         =   "frmTTPositions.frx":12FF
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmTTPositions.frx":1321
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTTPositions.frx":1341
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblClosedProfit 
               Height          =   195
               Left            =   240
               Top             =   1140
               Width           =   1335
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
               Caption         =   "frmTTPositions.frx":135D
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmTTPositions.frx":13A5
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTTPositions.frx":13C5
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblPlusClosedProfit 
               Height          =   255
               Left            =   0
               Top             =   1080
               Width           =   195
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
               Caption         =   "frmTTPositions.frx":13E1
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmTTPositions.frx":1403
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTTPositions.frx":1423
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblTotalFees 
               Height          =   195
               Left            =   240
               Top             =   780
               Width           =   1335
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
               Caption         =   "frmTTPositions.frx":143F
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmTTPositions.frx":1475
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTTPositions.frx":1495
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblPlusFees 
               Height          =   255
               Left            =   0
               Top             =   720
               Width           =   195
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
               Caption         =   "frmTTPositions.frx":14B1
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmTTPositions.frx":14D3
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTTPositions.frx":14F3
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblTotalAdjustments 
               Height          =   195
               Left            =   240
               Top             =   420
               Width           =   1335
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
               Caption         =   "frmTTPositions.frx":150F
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmTTPositions.frx":1553
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTTPositions.frx":1573
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblPlusAdjustments 
               Height          =   255
               Left            =   0
               Top             =   360
               Width           =   195
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
               Caption         =   "frmTTPositions.frx":158F
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmTTPositions.frx":15B1
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTTPositions.frx":15D1
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblStartingBalance 
               Height          =   255
               Left            =   240
               Top             =   15
               Width           =   1335
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
               Caption         =   "frmTTPositions.frx":15ED
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmTTPositions.frx":1631
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTTPositions.frx":1651
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin VB.Line Line2 
               X1              =   0
               X2              =   3960
               Y1              =   2520
               Y2              =   2520
            End
            Begin VB.Line Line1 
               X1              =   0
               X2              =   3960
               Y1              =   1500
               Y2              =   1500
            End
         End
         Begin HexUniControls.ctlUniFrameWL fraAccountInformation 
            Height          =   3075
            Left            =   180
            TabIndex        =   2
            Top             =   120
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
            Caption         =   "frmTTPositions.frx":166D
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmTTPositions.frx":1699
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTTPositions.frx":16B9
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniTextBoxXP txtAccountName 
               Height          =   285
               Left            =   1380
               TabIndex        =   6
               Top             =   355
               Width           =   2235
               _ExtentX        =   0
               _ExtentY        =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Enabled         =   -1  'True
               Locked          =   0   'False
               Text            =   "frmTTPositions.frx":16D5
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
               Tip             =   "frmTTPositions.frx":16F5
               HideSelection   =   -1  'True
               RightToLeft     =   0   'False
               ManualStart     =   0   'False
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTTPositions.frx":1715
            End
            Begin HexUniControls.ctlUniTextBoxXP txtAccountNumber 
               Height          =   285
               Left            =   1380
               TabIndex        =   4
               Top             =   0
               Width           =   2235
               _ExtentX        =   0
               _ExtentY        =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Enabled         =   -1  'True
               Locked          =   0   'False
               Text            =   "frmTTPositions.frx":1731
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
               Tip             =   "frmTTPositions.frx":1751
               HideSelection   =   -1  'True
               RightToLeft     =   0   'False
               ManualStart     =   0   'False
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTTPositions.frx":1771
            End
            Begin HexUniControls.ctlUniComboImageXP cboAccountType 
               Height          =   315
               Left            =   1380
               TabIndex        =   12
               Top             =   1440
               Width           =   2235
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
               Tip             =   "frmTTPositions.frx":178D
               Sorted          =   0   'False
               HScroll         =   0   'False
               RoundedBorders  =   -1  'True
               IconDim         =   16
               MousePointer    =   0
               MouseIcon       =   "frmTTPositions.frx":17AD
               DropDownOnTextClick=   -1  'True
               DropDownWidth   =   -1
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniTextBoxXP txtBroker 
               Height          =   285
               Left            =   1380
               TabIndex        =   10
               Top             =   1095
               Width           =   2235
               _ExtentX        =   0
               _ExtentY        =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Enabled         =   -1  'True
               Locked          =   0   'False
               Text            =   "frmTTPositions.frx":17C9
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
               Tip             =   "frmTTPositions.frx":17E9
               HideSelection   =   -1  'True
               RightToLeft     =   0   'False
               ManualStart     =   0   'False
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTTPositions.frx":1809
            End
            Begin gdOCX.gdSelectDate gdStartDate 
               Height          =   315
               Left            =   1380
               TabIndex        =   8
               Top             =   705
               Width           =   2235
               _ExtentX        =   3942
               _ExtentY        =   556
            End
            Begin HexUniControls.ctlUniCheckXP chkBrokerOCO 
               Height          =   220
               Left            =   0
               TabIndex        =   13
               Top             =   1860
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
               Caption         =   "frmTTPositions.frx":1825
               Enabled         =   -1  'True
               Align           =   0
               CheckBackColor  =   -2147483643
               CheckForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmTTPositions.frx":1899
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmTTPositions.frx":18B9
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblSimDesc 
               Height          =   435
               Left            =   0
               Top             =   1800
               Width           =   3555
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
               Caption         =   "frmTTPositions.frx":18D5
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmTTPositions.frx":194D
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTTPositions.frx":196D
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblStartDate 
               Height          =   255
               Left            =   0
               Top             =   735
               Width           =   1335
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
               Caption         =   "frmTTPositions.frx":1989
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmTTPositions.frx":19C7
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTTPositions.frx":19E7
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblName 
               Height          =   255
               Left            =   0
               Top             =   370
               Width           =   1335
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
               Caption         =   "frmTTPositions.frx":1A03
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmTTPositions.frx":1A37
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTTPositions.frx":1A57
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblNumber 
               Height          =   255
               Left            =   0
               Top             =   15
               Width           =   1335
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
               Caption         =   "frmTTPositions.frx":1A73
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmTTPositions.frx":1AB3
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTTPositions.frx":1AD3
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblAccountType 
               Height          =   255
               Left            =   0
               Top             =   1470
               Width           =   1335
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
               Caption         =   "frmTTPositions.frx":1AEF
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmTTPositions.frx":1B2B
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTTPositions.frx":1B4B
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblBroker 
               Height          =   255
               Left            =   0
               Top             =   1110
               Width           =   1335
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
               Caption         =   "frmTTPositions.frx":1B67
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmTTPositions.frx":1B97
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTTPositions.frx":1BB7
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
         End
         Begin HexUniControls.ctlUniFrameWL fraSecTypes 
            Height          =   855
            Left            =   180
            TabIndex        =   21
            Top             =   3300
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
            Caption         =   "frmTTPositions.frx":1BD3
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmTTPositions.frx":1C1F
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTTPositions.frx":1C3F
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniCheckXP chkStkOpts 
               Height          =   220
               Left            =   1320
               TabIndex        =   7
               Top             =   540
               Width           =   1155
               _ExtentX        =   2037
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
               Caption         =   "frmTTPositions.frx":1C5B
               Enabled         =   -1  'True
               Align           =   0
               CheckBackColor  =   -2147483643
               CheckForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmTTPositions.frx":1C91
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmTTPositions.frx":1CB1
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniCheckXP chkFutOpts 
               Height          =   220
               Left            =   120
               TabIndex        =   9
               Top             =   540
               Width           =   1155
               _ExtentX        =   2037
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
               Caption         =   "frmTTPositions.frx":1CCD
               Enabled         =   -1  'True
               Align           =   0
               CheckBackColor  =   -2147483643
               CheckForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmTTPositions.frx":1D03
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmTTPositions.frx":1D23
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniCheckXP chkFutures 
               Height          =   220
               Left            =   120
               TabIndex        =   22
               Top             =   240
               Width           =   975
               _ExtentX        =   1720
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
               Caption         =   "frmTTPositions.frx":1D3F
               Enabled         =   -1  'True
               Align           =   0
               CheckBackColor  =   -2147483643
               CheckForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmTTPositions.frx":1D6F
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmTTPositions.frx":1D8F
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniCheckXP chkStocks 
               Height          =   220
               Left            =   1320
               TabIndex        =   23
               Top             =   240
               Width           =   975
               _ExtentX        =   1720
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
               Caption         =   "frmTTPositions.frx":1DAB
               Enabled         =   -1  'True
               Align           =   0
               CheckBackColor  =   -2147483643
               CheckForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmTTPositions.frx":1DD9
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmTTPositions.frx":1DF9
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniCheckXP chkForex 
               Height          =   220
               Left            =   2520
               TabIndex        =   24
               Top             =   240
               Width           =   975
               _ExtentX        =   1720
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
               Caption         =   "frmTTPositions.frx":1E15
               Enabled         =   -1  'True
               Align           =   0
               CheckBackColor  =   -2147483643
               CheckForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmTTPositions.frx":1E41
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmTTPositions.frx":1E61
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraMessages 
         CausesValidation=   0   'False
         Height          =   5700
         Left            =   12360
         TabIndex        =   11
         Top             =   330
         Width           =   10425
         _ExtentX        =   18389
         _ExtentY        =   10054
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmTTPositions.frx":1E7D
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmTTPositions.frx":1E9D
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTTPositions.frx":1EBD
         RightToLeft     =   0   'False
         Begin VSFlex7LCtl.VSFlexGrid fgActivityLog 
            Height          =   2355
            Left            =   120
            TabIndex        =   15
            Top             =   120
            Width           =   4215
            _cx             =   7435
            _cy             =   4154
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
      Begin HexUniControls.ctlUniFrameWL fraAccountStatus 
         Height          =   5700
         Left            =   12060
         TabIndex        =   17
         Top             =   330
         Width           =   10425
         _ExtentX        =   18389
         _ExtentY        =   10054
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmTTPositions.frx":1ED9
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmTTPositions.frx":1EF9
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTTPositions.frx":1F19
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniFrameWL fraPositionFilter 
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   120
            Width           =   7515
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
            Caption         =   "frmTTPositions.frx":1F35
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmTTPositions.frx":1F55
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTTPositions.frx":1F75
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniCheckXP chkShowFlat 
               Height          =   195
               Left            =   0
               TabIndex        =   20
               Top             =   0
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
               Caption         =   "frmTTPositions.frx":1F91
               Enabled         =   -1  'True
               Align           =   0
               CheckBackColor  =   -2147483643
               CheckForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmTTPositions.frx":1FD9
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmTTPositions.frx":1FF9
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
         End
         Begin VSFlex7LCtl.VSFlexGrid fgAccountPositions 
            Height          =   1635
            Left            =   120
            TabIndex        =   26
            Top             =   420
            Width           =   4455
            _cx             =   7858
            _cy             =   2884
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
      Begin HexUniControls.ctlUniFrameWL fraFillHistory 
         Height          =   5700
         Left            =   11460
         TabIndex        =   56
         Top             =   330
         Width           =   10425
         _ExtentX        =   18389
         _ExtentY        =   10054
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmTTPositions.frx":2015
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmTTPositions.frx":2035
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTTPositions.frx":2055
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniCheckXP chkShowJournal 
            Height          =   195
            Left            =   6000
            TabIndex        =   61
            Top             =   120
            Width           =   1335
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
            Caption         =   "frmTTPositions.frx":2071
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmTTPositions.frx":20AD
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmTTPositions.frx":20CD
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkDateRange 
            Height          =   255
            Left            =   120
            TabIndex        =   57
            Top             =   90
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
            Caption         =   "frmTTPositions.frx":20E9
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmTTPositions.frx":211D
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmTTPositions.frx":213D
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin VSFlex7LCtl.VSFlexGrid fgTransactions 
            Height          =   1215
            Left            =   120
            TabIndex        =   62
            Top             =   480
            Width           =   2535
            _cx             =   4471
            _cy             =   2143
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
         Begin gdOCX.gdSelectDate gdFillsFromDate 
            Height          =   315
            Left            =   1140
            TabIndex        =   58
            Top             =   60
            Width           =   2115
            _ExtentX        =   3731
            _ExtentY        =   556
         End
         Begin gdOCX.gdSelectDate gdFillsToDate 
            Height          =   315
            Left            =   3600
            TabIndex        =   60
            Top             =   60
            Width           =   2115
            _ExtentX        =   3731
            _ExtentY        =   556
         End
         Begin HexUniControls.ctlUniLabelXP lblTo 
            Height          =   255
            Left            =   3360
            Top             =   90
            Width           =   195
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
            Caption         =   "frmTTPositions.frx":2159
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmTTPositions.frx":217D
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTTPositions.frx":219D
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraOrderHistory 
         Height          =   5700
         Left            =   11160
         TabIndex        =   50
         Top             =   330
         Width           =   10425
         _ExtentX        =   18389
         _ExtentY        =   10054
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmTTPositions.frx":21B9
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmTTPositions.frx":21D9
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTTPositions.frx":21F9
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniFrameWL fraClosedOrders 
            Height          =   315
            Left            =   120
            TabIndex        =   30
            Top             =   2580
            Width           =   3975
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
            Caption         =   "frmTTPositions.frx":2215
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmTTPositions.frx":2241
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTTPositions.frx":2261
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniTextBoxXP txtNumDays 
               Height          =   285
               Left            =   1860
               TabIndex        =   54
               Top             =   0
               Width           =   435
               _ExtentX        =   0
               _ExtentY        =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Enabled         =   -1  'True
               Locked          =   0   'False
               Text            =   "frmTTPositions.frx":227D
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
               Tip             =   "frmTTPositions.frx":229D
               HideSelection   =   -1  'True
               RightToLeft     =   0   'False
               ManualStart     =   0   'False
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTTPositions.frx":22BD
            End
            Begin HexUniControls.ctlUniLabelXP lblClosedOrders 
               Height          =   255
               Left            =   0
               Top             =   30
               Width           =   2595
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
               Caption         =   "frmTTPositions.frx":22D9
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmTTPositions.frx":2327
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTTPositions.frx":2347
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
         End
         Begin VSFlex7LCtl.VSFlexGrid fgClosedOrders 
            Height          =   1455
            Left            =   120
            TabIndex        =   55
            Top             =   2940
            Width           =   3855
            _cx             =   6800
            _cy             =   2566
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
         Begin VSFlex7LCtl.VSFlexGrid fgOpenOrders 
            Height          =   1815
            Left            =   120
            TabIndex        =   52
            Top             =   420
            Width           =   3675
            _cx             =   6482
            _cy             =   3201
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
         Begin HexUniControls.ctlUniLabelXP lblOpenOrders 
            Height          =   255
            Left            =   120
            Top             =   120
            Width           =   1335
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
            Caption         =   "frmTTPositions.frx":2363
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmTTPositions.frx":239B
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTTPositions.frx":23BB
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraAccountHistory 
         Height          =   5700
         Left            =   11760
         TabIndex        =   63
         Top             =   330
         Width           =   10425
         _ExtentX        =   18389
         _ExtentY        =   10054
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmTTPositions.frx":23D7
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmTTPositions.frx":2403
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTTPositions.frx":2423
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniFrameWL fraTradeFilter 
            Height          =   795
            Left            =   120
            TabIndex        =   64
            Top             =   60
            Width           =   10215
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
            Caption         =   "frmTTPositions.frx":243F
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmTTPositions.frx":246B
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTTPositions.frx":248B
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniButtonImageXP cmdTradesFilter 
               Height          =   315
               Left            =   180
               TabIndex        =   32
               Top             =   300
               Width           =   975
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
               Caption         =   "frmTTPositions.frx":24A7
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               ShowFocus       =   -1  'True
               Tristate        =   0   'False
               Pressed         =   0   'False
               Tip             =   "frmTTPositions.frx":24DB
               Style           =   -1
               RoundedBorders  =   -1  'True
               xTranspColor    =   0
               yTranspColor    =   0
               MousePointer    =   0
               MouseIcon       =   "frmTTPositions.frx":24FB
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblTradeFilter 
               Height          =   375
               Left            =   1320
               Top             =   240
               Width           =   8715
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
               Caption         =   "frmTTPositions.frx":2517
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmTTPositions.frx":254F
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTTPositions.frx":256F
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
         End
         Begin HexUniControls.ctlUniTextBoxXP txtNotes 
            Height          =   915
            Left            =   120
            TabIndex        =   33
            Top             =   3660
            Width           =   5055
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmTTPositions.frx":258B
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
            ScrollBars      =   0
            PasswordChar    =   ""
            TrapTab         =   0   'False
            EnableContextMenu=   -1  'True
            RaiseChangeEvent=   -1  'True
            Tip             =   "frmTTPositions.frx":25AB
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTTPositions.frx":25CB
         End
         Begin VSFlex7LCtl.VSFlexGrid fgTrades 
            Height          =   1695
            Left            =   120
            TabIndex        =   35
            Top             =   1260
            Width           =   2895
            _cx             =   5106
            _cy             =   2990
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
         Begin HexUniControls.ctlUniLabelXP lblOpenEquity 
            Height          =   255
            Left            =   1200
            Top             =   3060
            Width           =   4875
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
            Caption         =   "frmTTPositions.frx":25E7
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmTTPositions.frx":268B
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTTPositions.frx":26AB
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblNotes 
            Height          =   255
            Left            =   120
            Top             =   3360
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
            Caption         =   "frmTTPositions.frx":26C7
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmTTPositions.frx":2725
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTTPositions.frx":2745
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   3975
      Left            =   11040
      TabIndex        =   36
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
      Caption         =   "frmTTPositions.frx":2761
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmTTPositions.frx":278D
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTTPositions.frx":27AD
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdSell 
         Height          =   375
         Left            =   600
         TabIndex        =   38
         Top             =   0
         Width           =   495
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
         Caption         =   "frmTTPositions.frx":27C9
         BackColor       =   10526975
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTTPositions.frx":27F1
         Style           =   1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTTPositions.frx":2811
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdBuy 
         Height          =   375
         Left            =   0
         TabIndex        =   39
         Top             =   0
         Width           =   555
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
         Caption         =   "frmTTPositions.frx":282D
         BackColor       =   16752800
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTTPositions.frx":2853
         Style           =   1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTTPositions.frx":2873
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdCheckStatus 
         Height          =   375
         Left            =   0
         TabIndex        =   41
         Top             =   2220
         Width           =   1515
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
         Caption         =   "frmTTPositions.frx":288F
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTTPositions.frx":28C9
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTTPositions.frx":28E9
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdCancelOrder 
         Height          =   375
         Left            =   0
         TabIndex        =   42
         Top             =   1680
         Width           =   1515
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
         Caption         =   "frmTTPositions.frx":2905
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTTPositions.frx":293F
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTTPositions.frx":295F
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdPrint 
         Height          =   375
         Left            =   0
         TabIndex        =   44
         Top             =   3180
         Width           =   1515
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
         Caption         =   "frmTTPositions.frx":297B
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTTPositions.frx":29A7
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTTPositions.frx":29C7
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdAdjustment 
         Height          =   375
         Left            =   0
         TabIndex        =   45
         Top             =   0
         Width           =   1515
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
         Caption         =   "frmTTPositions.frx":29E3
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTTPositions.frx":2A21
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTTPositions.frx":2A41
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdAccount 
         Height          =   375
         Left            =   0
         TabIndex        =   47
         Top             =   2760
         Width           =   1515
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
         Caption         =   "frmTTPositions.frx":2A5D
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTTPositions.frx":2A9B
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTTPositions.frx":2ABB
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdDelete 
         Height          =   375
         Left            =   0
         TabIndex        =   48
         Top             =   1260
         Width           =   1515
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
         Caption         =   "frmTTPositions.frx":2AD7
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTTPositions.frx":2B05
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTTPositions.frx":2B25
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdEdit 
         Height          =   375
         Left            =   0
         TabIndex        =   51
         Top             =   840
         Width           =   1515
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
         Caption         =   "frmTTPositions.frx":2B41
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTTPositions.frx":2B6B
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTTPositions.frx":2B8B
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdExit 
         Height          =   375
         Left            =   0
         TabIndex        =   53
         Top             =   3600
         Width           =   1515
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
         Caption         =   "frmTTPositions.frx":2BA7
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTTPositions.frx":2BD3
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTTPositions.frx":2BF3
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdNewTrade 
         Height          =   375
         Left            =   0
         TabIndex        =   59
         Top             =   420
         Width           =   1515
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
         Caption         =   "frmTTPositions.frx":2C0F
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTTPositions.frx":2C43
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTTPositions.frx":2C63
         RightToLeft     =   0   'False
      End
   End
   Begin VB.Menu mnuOrders 
      Caption         =   "Orders"
      Begin VB.Menu mnuEditOrder 
         Caption         =   "Edit Order"
      End
      Begin VB.Menu mnuCancelOrder 
         Caption         =   "Cancel Order"
      End
      Begin VB.Menu mnuParkOrder 
         Caption         =   "Park Order"
      End
      Begin VB.Menu mnuSubmitOrder 
         Caption         =   "Submit Order"
      End
      Begin VB.Menu mnuSubmitAll 
         Caption         =   "Submit All Parked Orders"
      End
      Begin VB.Menu mnuOrderHistory 
         Caption         =   "Order History"
      End
      Begin VB.Menu mnuOrderJournal 
         Caption         =   "New Journal for Order"
      End
      Begin VB.Menu mnuOrdersSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOrdersPrint 
         Caption         =   "Print"
      End
      Begin VB.Menu mnuOrdersChangeFont 
         Caption         =   "Change Font"
      End
      Begin VB.Menu mnuViewJournals 
         Caption         =   "View Journals"
      End
   End
   Begin VB.Menu mnuTransactions 
      Caption         =   "Transactions"
      Begin VB.Menu mnuTransactionsNewAdjustment 
         Caption         =   "New Adjustment"
      End
      Begin VB.Menu mnuTransactionsNewFill 
         Caption         =   "New Fill"
      End
      Begin VB.Menu mnuTransactionsEdit 
         Caption         =   "Edit"
      End
      Begin VB.Menu mnuTransactionsDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnuTransactionsSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTransactionsEditJournal 
         Caption         =   "Edit Journal"
      End
      Begin VB.Menu mnuTransactionsDeleteJournal 
         Caption         =   "Delete Journal"
      End
      Begin VB.Menu mnuTransactionsSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTransactionsExport 
         Caption         =   "Export"
      End
      Begin VB.Menu mnuTransactionsPrint 
         Caption         =   "Print"
      End
      Begin VB.Menu mnuTransactionsChangeFont 
         Caption         =   "Change Font"
      End
   End
   Begin VB.Menu mnuTrades 
      Caption         =   "Trades"
      Begin VB.Menu mnuTradesNewAdjustment 
         Caption         =   "New Adjustment"
      End
      Begin VB.Menu mnuTradesSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTradesExport 
         Caption         =   "Export"
      End
      Begin VB.Menu mnuTradesPrint 
         Caption         =   "Print"
      End
      Begin VB.Menu mnuTradesChangeFont 
         Caption         =   "Change Font"
      End
   End
   Begin VB.Menu mnuPositions 
      Caption         =   "Positions"
      Begin VB.Menu mnuPositionsFlatten 
         Caption         =   "Flatten"
      End
      Begin VB.Menu mnuPositionsReverse 
         Caption         =   "Reverse"
      End
      Begin VB.Menu mnuPositionsSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPositionsPrint 
         Caption         =   "Print"
      End
      Begin VB.Menu mnuPositionsChangeFont 
         Caption         =   "Change Font"
      End
   End
   Begin VB.Menu mnuActivityLog 
      Caption         =   "Activity Log"
      Begin VB.Menu mnuActivityLogPrint 
         Caption         =   "Print"
      End
      Begin VB.Menu mnuActivityLogChangeFont 
         Caption         =   "Change Font"
      End
   End
End
Attribute VB_Name = "frmTTPositions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmTTPositions.frm
'' Description: Allows the user to see their trade history for an account
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 02/06/2009   DAJ         Display "Mismatch" for position if in a mismatch
'' 04/23/2009   DAJ         Added journal entries to the transaction grid
'' 04/30/2009   DAJ         Change manual fill from snapshot to history if necessary
'' 06/01/2009   DAJ         Added FO and SO to security type mask for trading
'' 06/02/2009   DAJ         Use Bid/Ask instead of Last for Options P&L
'' 06/23/2009   DAJ         Only reload symbols combo on trades tab if needed
'' 08/21/2009   DAJ         Set UserCancel flag on CancelOrder call
'' 09/01/2009   DAJ         Use new Parked order status
'' 09/23/2009   DAJ         Ensured Transactions and Trades export dates the same
'' 10/07/2009   DAJ         Added support for Linked Orders at broker
'' 12/01/2009   DAJ         Added support for automatic commissions on fills
'' 12/04/2009   DAJ         Changed automatic commissions for stocks
'' 03/11/2010   DAJ         Use global collections and activity log object
'' 04/19/2010   DAJ         Account for fees on the trades tab for calculating balance
'' 04/21/2010   DAJ         Took out flag file check for allowing Broker OCO's
'' 05/25/2010   DAJ         Added checks for g.Broker.Account being Nothing
'' 06/14/2010   DAJ         Allow creation of a new account (#5781)
'' 07/01/2010   DAJ         Mods for the Extreme Charts versions of the software
'' 09/24/2010   DAJ         Added some artwork to be shown for Rithmic
'' 10/05/2010   DAJ         Changed the Rithmic image
'' 10/27/2010   DAJ         More mods to the Rithmic image
'' 06/21/2011   DAJ         Separate out Simulated trading types
'' 06/24/2011   DAJ         Utilize NextAccount functions from simulated objects
'' 07/13/2011   DAJ         If user modifies fill for simulated, call for a refresh
'' 07/15/2011   DAJ         If user modifies snapshot fills, call for a position refresh
'' 07/19/2011   DAJ         Moved new, edit, and delete fill calls to g.Broker
'' 07/28/2011   DAJ         Allow user to create new account from this form
'' 07/28/2011   DAJ         Default new account to SimBroker, give a description for sims
'' 08/15/2011   DAJ         Reload historical fills if created/modified/deleted a historical fill
'' 08/23/2011   DAJ         When user saves account, send it to Option Nav
'' 08/24/2011   DAJ         No longer load historical closed orders
'' 08/25/2011   DAJ         Efficiency tweaks
'' 09/23/2011   DAJ         Show date journals form instead of old journals form
'' 10/04/2011   DAJ         Call the ShowJournals function instead of calling the form direct
'' 10/06/2011   DAJ         Changed default on new simulated account if user has RTG
'' 10/07/2011   DAJ         Don't show SimBroker warning in account combo if not visible yet
'' 10/20/2011   DAJ         Dump BrokerInfo to log
'' 02/20/2013   DAJ         Changed trade filter, utilize settings object
'' 04/15/2013   DAJ         Make sure that the trade report filter defaults to the account shown
'' 06/24/2013   DAJ         Timer Logging
'' 10/16/2013   DAJ         Remove Xpress
'' 02/11/2014   DAJ         Changed "Commissions" frame label to "Commissions and Fees"
'' 09/02/2014   DAJ         Move Journal stuff into Journal DLL
'' 09/04/2014   DAJ         Fix for "Object variable..." error in UpdateJournal
'' 12/10/2014   DAJ         Utilize new DateIsSnapshot routines
'' 03/19/2015   DAJ         Speed enhancements for refreshing prices
'' 05/11/2015   DAJ         Added support for multi-selecting and deleting fills on Transactions tab
'' 05/15/2015   DAJ         Fix for multi-select deleting when only deleting snapshot fills
'' 05/20/2015   DAJ         Allow multiple accounts for the trade report filter
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Const kOpenEquityColor = &HFFFFC0      '&HFFFF80

Private Enum eGDOpenOrdersCols
    eGDOpenOrdersCol_OrderID = 0
    eGDOpenOrdersCol_SymbolID
    eGDOpenOrdersCol_Remove
    
    eGDOpenOrdersCol_Symbol
    eGDOpenOrdersCol_OrderText
    eGDOpenOrdersCol_Cancel
    eGDOpenOrdersCol_NumFilled
    eGDOpenOrdersCol_Status
    eGDOpenOrdersCol_CurrentPrice
    eGDOpenOrdersCol_CurrentBid
    eGDOpenOrdersCol_CurrentAsk
    eGDOpenOrdersCol_Date
    eGDOpenOrdersCol_AutoTradeItem
    eGDOpenOrdersCol_BrokerID
    eGDOpenOrdersCol_LinkStatus
    
    eGDOpenOrdersCol_NumCols
End Enum

Private Enum eGDClosedOrdersCols
    eGDClosedOrdersCol_OrderID = 0
    eGDClosedOrdersCol_SymbolID
    eGDClosedOrdersCol_Remove
    
    eGDClosedOrdersCol_Symbol
    eGDClosedOrdersCol_OrderText
    eGDClosedOrdersCol_NumFilled
    eGDClosedOrdersCol_Status
    eGDClosedOrdersCol_Date
    eGDClosedOrdersCol_AutoTradeItem
    eGDClosedOrdersCol_BrokerID
    eGDClosedOrdersCol_LinkStatus
    
    eGDClosedOrdersCol_NumCols
End Enum

Private Enum eGDTransactionCols
    eGDTransactionCol_Action = 0
    eGDTransactionCol_Quantity
    eGDTransactionCol_Symbol
    eGDTransactionCol_Price
    eGDTransactionCol_Date
    eGDTransactionCol_Fees
    eGDTransactionCol_BrokerID
    eGDTransactionCol_ClosedProfit
    eGDTransactionCol_SessionQuantity
    eGDTransactionCol_SessionProfit
    eGDTransactionCol_Balance
    eGDTransactionCol_Position
    eGDTransactionCol_FillID
    eGDTransactionCol_OrderID
    eGDTransactionCol_PositionID
    eGDTransactionCol_AdjustmentID
    eGDTransactionCol_FillDate
    eGDTransactionCol_Remove
    
    eGDTransactionCol_JournalID
    eGDTransactionCol_EmotionNumber
    eGDTransactionCol_Feelings
    eGDTransactionCol_Reasons
    eGDTransactionCol_Thoughts
    eGDTransactionCol_Note
    
    eGDTransactionCol_NumCols
End Enum

Private Enum eGDTradeCols
    eGDTradeCol_AcctPosID = 0
    eGDTradeCol_Sequence
    
    eGDTradeCol_Symbol
    eGDTradeCol_Position
    eGDTradeCol_Quantity
    eGDTradeCol_EntryDate
    eGDTradeCol_EntryPrice
    eGDTradeCol_EntryRule
    eGDTradeCol_ExitDate
    eGDTradeCol_ExitPrice
    eGDTradeCol_ExitRule
    eGDTradeCol_Profit
    eGDTradeCol_Commission
    eGDTradeCol_Balance
    eGDTradeCol_Flag
    eGDTradeCol_Category
    
    eGDTradeCol_ClosedExitPrice
    eGDTradeCol_ExitQuantity
    eGDTradeCol_Notes
    eGDTradeCol_CategoryID
    eGDTradeCol_Remove
    
    eGDTradeCol_NumCols
End Enum

Private Enum eGDAccountPositionCols
    eGDAccountPositionCol_SymbolID = 0
    eGDAccountPositionCol_Symbol
    eGDAccountPositionCol_AutoTradeItem
    eGDAccountPositionCol_Position
    eGDAccountPositionCol_Quantity
    eGDAccountPositionCol_Flatten
    eGDAccountPositionCol_Reverse
    eGDAccountPositionCol_AvgEntry
    eGDAccountPositionCol_CurrentPrice
    eGDAccountPositionCol_OpenProfit
    eGDAccountPositionCol_OrderStrategy
    eGDAccountPositionCol_LastTraded
    eGDAccountPositionCol_SessionDate
    eGDAccountPositionCol_SessionQuantity
    eGDAccountPositionCol_SessionProfit
    eGDAccountPositionCol_Remove
    eGDAccountPositionCol_NumCols
End Enum

Private Type mPrivate
    lAccountID As Long                  ' ID of the account to edit
    Account As cPtAccount               ' Account object for the account to edit
    
    bLoading As Boolean                 ' Are we currently loading the account?
    strNumDaysSave As String            ' Number of days for closed order filter
    bReload As Boolean                  ' Flag to tell the form to reload all of the grids
    dLastChanged As Double              ' Last changed value from the broker info object
    bClearInfBox As Boolean             ' Do we need to clear the InfBox?
    bEditingSB As Boolean               ' Is the user editing the starting balance?
    
    ActivityLog As cActivityLog         ' Activity log object
    
    HistoricalOrders As cPtOrders       ' Collection of historical orders for this account
    SnapshotOrders As cPtOrders         ' Collection of snapshot orders for this account
    HistoricalFills As cPtFills         ' Collection of historical fills for this account
    SnapshotFills As cPtFills           ' Collection of snapshot fills for this account
    Positions As cAccountPositions      ' Collection of account positions for this account
    Adjustments As cGdTree              ' Collection of adjustments
    
    Journals As cGdTree                 ' Collection of journal entries for this account
    JournalOrderMap As cGdTree          ' Order to Journal map
    TradeFilter As cTradeFilterSettings ' Trade filter settings
    
    TradeRules As cTradeRules           ' Trade rules object
    astrFlags As cGdArray               ' Array of flags
    
    astrTradeSymbols As cGdArray        ' Array of trade symbols
End Type
Private m As mPrivate

Public Property Get AccountID() As Long
    AccountID = m.lAccountID
End Property

Private Function OpenOrdersCol(ByVal Col As eGDOpenOrdersCols) As Long
    OpenOrdersCol = Col
End Function
Private Function ClosedOrdersCol(ByVal Col As eGDClosedOrdersCols) As Long
    ClosedOrdersCol = Col
End Function
Private Function TradeCol(ByVal Col As eGDTradeCols) As Long
    TradeCol = Col
End Function
Private Function TransactionCol(ByVal Col As eGDTransactionCols) As Long
    TransactionCol = Col
End Function
Private Function AccountPosCol(ByVal Col As eGDAccountPositionCols) As Long
    AccountPosCol = Col
End Function
Private Function Tabs(ByVal eTab As eGDTradeTrackerTabs) As Long
    Tabs = eTab
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Set up and show the form
'' Inputs:      Account ID, Account Type, Starting Tab, Trade Filter Settings
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowMe(ByVal lAccountID As Long, Optional ByVal nAccountType As eTT_AccountType = eTT_AccountType_SimBroker, Optional ByVal nStartTab As eGDTradeTrackerTabs = -1&, Optional ByVal TradeFilter As cTradeFilterSettings = Nothing)
On Error GoTo ErrSection:

    Dim UI As cActivityLogControls      ' Activity log controls object

    If Me.Visible Then
        If lAccountID <> m.lAccountID Then
            If AskToSave Then Exit Sub
        Else
            Exit Sub
        End If
    End If

    Screen.MousePointer = vbHourglass

    m.lAccountID = lAccountID
    m.bLoading = True
    
    If (m.lAccountID > 0) And (g.Broker.Account(m.lAccountID) Is Nothing) Then
        InfBox "Error opening account", "!", , "Trade Tracker Error"
        Unload Me
    Else
        If m.lAccountID > 0 Then
            Set m.Account = g.Broker.Account(m.lAccountID).MakeCopy
        Else
            Set m.Account = New cPtAccount
            If (HasModule("RTG") = True) And (nAccountType = eTT_AccountType_SimBroker) Then
                nAccountType = eTT_AccountType_SimStream
            End If
            m.Account.AccountType = nAccountType
            If nAccountType = eTT_AccountType_SimStream Then
                m.Account.AccountNumber = g.SimTradeStream.NextAccount
            ElseIf nAccountType = eTT_AccountType_SimBroker Then
                m.Account.AccountNumber = g.SimTradeTs.NextAccount
            End If
        End If
        
        If Len(m.Account.Name) > 0 Then
            Caption = "Trade Tracker - [" & m.Account.Name & "]"
        ElseIf Len(m.Account.AccountNumber) > 0 Then
            Caption = "Trade Tracker - [" & m.Account.AccountNumber & "]"
        Else
            Caption = "Trade Tracker - [New Account]"
        End If
        
        gdFillsFromDate.MaxDateIsToday = True
        gdFillsToDate.MaxDateIsToday = True
        
        ' Initialize all of the grids on the form...
        InitOpenOrdersGrid
        InitClosedOrdersGrid
        InitTransactionsGrid
        InitPositionsGrid
        InitAccountPositionsGrid
        
        Set UI = New cActivityLogControls
        With UI
            Set .frm = Me
            Set .fgGrid = fgActivityLog
            Set .tmrMenu = tmrMenu
            Set .tmrRealTime = frmTTSummary.tmrRealTime
            Set .mnuActivityLog = mnuActivityLog
            Set .mnuPrint = mnuActivityLogPrint
        End With
        
        Set m.ActivityLog = New cActivityLog
        m.ActivityLog.Init UI, m.lAccountID
        g.ActivityLogs.Add "frmTTPositions", m.ActivityLog
        
        LoadJournals
        
        ' Set the trade filter settings if they were passed in...
        If TradeFilter Is Nothing Then
            LoadTradeFilterSettings
        Else
            Set m.TradeFilter = TradeFilter
        End If
        
        LoadGrids nAccountType
                
        ' If this is a new account, show the account information tab...
        If (m.Account.AccountID = 0&) Then
            tabPositions.CurrTab = Tabs(eGDTradeTrackerTab_Account)
        ElseIf nStartTab = -1& Then
            tabPositions.CurrTab = Tabs(eGDTradeTrackerTab_Orders)
        Else
            tabPositions.CurrTab = nStartTab
        End If
        
        EnableControls
        
        m.bLoading = False
        Screen.MousePointer = vbDefault
        
        tmrBrokers.Enabled = True
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    m.bLoading = False
    Screen.MousePointer = vbDefault
    Unload Me
    RaiseError "frmTTPositions.ShowMe"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PrintMe
'' Description: Allow as outside caller to print the grid information
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function PrintMe()
On Error GoTo ErrSection

    PrintMe = frmPrintPreview.ShowMe("CNV TTPositions", Me, , , , 0.75, 0.75, True)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTTPositions.PrintMe"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenerateReport
'' Description: Set up the print preview form for this grid
'' Inputs:      Arguments passed in from PrintMe
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GenerateReport(ByVal vArgs As Variant)
On Error GoTo ErrSection:
    
    Dim strText As String               ' Temporary string variable
    Dim astrSecTypes As New cGdArray    ' Security type array
    
    With frmPrintPreview.vp
        .StartDoc
        DoPrintHeader
        
        Select Case tabPositions.CurrTab
            Case Tabs(eGDTradeTrackerTab_Account)
                .Font.Name = "Times New Roman"
                .Font.Size = 14
                .Font.Bold = True
                .TextAlign = taCenterMiddle
                .Text = "Account Information for " & g.Broker.AccountNumberForID(m.lAccountID)
                .TextAlign = taLeftMiddle
                .Font.Bold = False
                
                .Paragraph = ""
                .Paragraph = ""
                
                .Text = "Account Number: " & txtAccountNumber.Text & vbLf
                .Text = "Account Nickname: " & txtAccountName.Text & vbLf
                .Text = "Starting Date: " & DateFormat(gdStartDate.Value) & vbLf
                .Text = "Broker: " & txtBroker.Text & vbLf
                
                If chkFutures.Value = vbChecked Then astrSecTypes.Add "Futures"
                If chkStocks.Value = vbChecked Then astrSecTypes.Add "Stocks"
                If chkForex.Value = vbChecked Then astrSecTypes.Add "Forex"
                If chkFutOpts.Value = vbChecked Then astrSecTypes.Add "Future Options"
                If chkStkOpts.Value = vbChecked Then astrSecTypes.Add "Stock Options"
                If astrSecTypes.Size = 0 Then astrSecTypes.Add "None"
                
                .Text = "Security Types Allowed for Trading: " & astrSecTypes.JoinFields(",") & vbLf
                
                .Text = vbLf & vbLf
                
                .Text = "Starting Balance: " & vbTab & vbTab & txtStartingBalance.Text & vbLf
                .Text = "+ Total Adjustments: " & vbTab & txtTotalAdjustments.Text & vbLf
                .Text = "+ Total Fees: " & vbTab & vbTab & txtTotalFees.Text & vbLf
                .Text = "+ Total Closed Profit: " & vbTab & txtTotalClosedProfit.Text & vbLf
                .Text = "___________________________________________" & vbLf
                .Text = "= Closed Balance: " & vbTab & txtClosedBalance.Text & vbLf
                .Text = "+ Total Open Equity: " & vbTab & txtTotalOpenEquity.Text & vbLf
                .Text = "___________________________________________" & vbLf
                .Text = "= Current Value: " & vbTab & vbTab & txtCurrentValue.Text
            
            Case Tabs(eGDTradeTrackerTab_Orders)
                .Font.Name = "Times New Roman"
                .Font.Size = 14
                .Font.Bold = True
                .TextAlign = taCenterMiddle
                strText = "Orders for " & g.Broker.AccountNumberForID(m.lAccountID)
                .Text = strText & vbCrLf
                .TextAlign = taLeftMiddle
                .Font.Bold = False
                
                .Paragraph = ""
                .Paragraph = ""
                .Text = "Open Orders:"
                .Paragraph = ""
                
                If frmPrintPreview.GoingToFile Then
                    frmPrintPreview.GridToTable fgOpenOrders
                Else
                    .RenderControl = fgOpenOrders.hWnd
                End If
                
                .Paragraph = ""
                .Paragraph = ""
                .Text = "Closed Orders for last " & Trim(txtNumDays.Text) & " days:"
                .Paragraph = ""
                
                If frmPrintPreview.GoingToFile Then
                    frmPrintPreview.GridToTable fgClosedOrders
                Else
                    .RenderControl = fgClosedOrders.hWnd
                End If
                
            Case Tabs(eGDTradeTrackerTab_Transactions)
                .Font.Name = "Times New Roman"
                .Font.Size = 14
                .Font.Bold = True
                .TextAlign = taCenterMiddle
                strText = "Transactions for " & g.Broker.AccountNumberForID(m.lAccountID)
                If chkDateRange = vbChecked Then
                    strText = strText & " from " & DateFormat(gdFillsFromDate.Value) & " to " & DateFormat(gdFillsToDate.Value)
                End If
                .Text = strText
                .TextAlign = taLeftMiddle
                .Font.Bold = False
                
                .Paragraph = ""
                .Paragraph = ""
                
                If frmPrintPreview.GoingToFile Then
                    frmPrintPreview.GridToTable fgTransactions
                Else
                    .RenderControl = fgTransactions.hWnd
                End If
                
            
            Case Tabs(eGDTradeTrackerTab_Trades)
                .Font.Name = "Times New Roman"
                .Font.Size = 14
                .Font.Bold = True
                .TextAlign = taCenterMiddle
                .Text = "Trades for " & g.Broker.AccountNumberForID(m.lAccountID)
                .TextAlign = taLeftMiddle
                .Font.Bold = False
                
                .Paragraph = ""
                .Paragraph = ""
                
                If frmPrintPreview.GoingToFile Then
                    frmPrintPreview.GridToTable fgTrades
                Else
                    .RenderControl = fgTrades.hWnd
                End If
                
            Case Tabs(eGDTradeTrackerTab_Positions)
                .Font.Name = "Times New Roman"
                .Font.Size = 14
                .Font.Bold = True
                .TextAlign = taCenterMiddle
                .Text = "Positions for " & g.Broker.AccountNumberForID(m.lAccountID)
                .TextAlign = taLeftMiddle
                .Font.Bold = False
                
                .Paragraph = ""
                .Paragraph = ""
                
                If frmPrintPreview.GoingToFile Then
                    frmPrintPreview.GridToTable fgAccountPositions
                Else
                    .RenderControl = fgAccountPositions.hWnd
                End If
                
            Case Tabs(eGDTradeTrackerTab_ActivityLog)
                .Font.Name = "Times New Roman"
                .Font.Size = 14
                .Font.Bold = True
                .TextAlign = taCenterMiddle
                .Text = "Activity Log for " & g.Broker.AccountNumberForID(m.lAccountID)
                .TextAlign = taLeftMiddle
                .Font.Bold = False
                
                .Paragraph = ""
                .Paragraph = ""
                
                If frmPrintPreview.GoingToFile Then
                    frmPrintPreview.GridToTable fgActivityLog
                Else
                    .RenderControl = fgActivityLog.hWnd
                End If
            
        End Select
        
        .EndDoc
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.GenerateReport"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Reload
'' Description: Set the reload flag if the account is for the given broker
'' Inputs:      Broker
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Reload(Optional ByVal nBroker As eTT_AccountType = -1&)
On Error GoTo ErrSection:

    If (nBroker = -1&) Or (nBroker = AccountType) Then
        m.bReload = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.Reload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshPrices
'' Description: Refresh the prices in appropriate places for the given bars
'' Inputs:      Bars with latest prices
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RefreshPrices(Bars As cGdBars)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lSymbolID As Long               ' Symbol ID from the Bars structure
    Dim strSymbol As String             ' Symbol from the Bars structure
    Dim dPrice As Double                ' Current price from the Bars
    Dim dBid As Double                  ' Current bid from the Bars
    Dim dAsk As Double                  ' Current ask from the Bars
    Dim strValue As String              ' String version of the value
    Dim dEntry As Double                ' Entry price for the position
    Dim lQuantity As Long               ' Quantity entered for the position
    Dim dProfit As Double               ' Profit for the position
    Dim dAvgExit As Double              ' Average exit price from the position
    Dim dDiff As Double                 ' Difference in profit
    Dim bFound As Boolean               ' Symbol was found
    Dim bChanged As Boolean             ' Did the open profit for a trade change?
    Dim dCurrent As Double              ' Current price to use for open profit
    Dim lStartRow As Long               ' Row to start re-calculating balance
    
    If m.bLoading = False Then
        ' Get the information out of the Bars structure once...
        strSymbol = Bars.Prop(eBARS_Symbol)
        lSymbolID = Bars.Prop(eBARS_SymbolID)
        dPrice = Bars(eBARS_Close, Bars.Size - 1)
        dBid = Bars(eBARS_Bid, Bars.Size - 1)
        dAsk = Bars(eBARS_Ask, Bars.Size - 1)
    
        With fgOpenOrders
            For lIndex = .FixedRows To .Rows - 1
                If .RowData(lIndex).SymbolOrSymbolID = Bars.SymbolOrSymbolID Then
                    strValue = ""
                    If dPrice <> kNullData Then strValue = Bars.PriceDisplay(dPrice)
                    ChangeCell fgOpenOrders, lIndex, OpenOrdersCol(eGDOpenOrdersCol_CurrentPrice), strValue
                
                    strValue = ""
                    If dBid <> kNullData Then strValue = Bars.PriceDisplay(dBid)
                    ChangeCell fgOpenOrders, lIndex, OpenOrdersCol(eGDOpenOrdersCol_CurrentBid), strValue
                
                    strValue = ""
                    If dAsk <> kNullData Then strValue = Bars.PriceDisplay(dAsk)
                    ChangeCell fgOpenOrders, lIndex, OpenOrdersCol(eGDOpenOrdersCol_CurrentAsk), strValue
                End If
            Next lIndex
        End With
        
        With fgTrades
            bFound = False
            bChanged = False
            lStartRow = kNullData
            
            For lIndex = .Rows - 1 To .FixedRows + 1 Step -1
                If TypeOf .RowData(lIndex) Is cTradeLine Then
                    If .RowData(lIndex).IsOpen = False Then
                        Exit For
                    ElseIf .RowData(lIndex).SymbolOrSymbolID = Bars.SymbolOrSymbolID Then
                        dCurrent = .RowData(lIndex).CurrentPrice(dPrice, dBid, dAsk)
                        
                        If Bars.Prop(eBARS_LastTickTime) <> 0 Then
                            .TextMatrix(lIndex, TradeCol(eGDTradeCol_ExitDate)) = Bars.LastTickDateTime
                        Else
                            .TextMatrix(lIndex, TradeCol(eGDTradeCol_ExitDate)) = Bars(eBARS_DateTime, Bars.Size - 1)
                        End If
                    
                        strValue = Bars.PriceDisplay(dCurrent)
                        If ChangeCell(fgTrades, lIndex, TradeCol(eGDTradeCol_ExitPrice), strValue) = True Then
                            dProfit = .RowData(lIndex).OpenProfit(dCurrent)
                            
                            .TextMatrix(lIndex, TradeCol(eGDTradeCol_Profit)) = dProfit
                            ColorCell fgTrades, lIndex, TradeCol(eGDTradeCol_Profit)
                            
                            lStartRow = lIndex
                            bChanged = True
                        End If
                        
                        bFound = True
                    End If
                End If
            Next lIndex
            
            If bChanged Then
                CalculateTradeBalance False, lStartRow
            End If
        End With

        With fgAccountPositions
            For lIndex = .FixedRows To .Rows - 1
                If TypeOf .RowData(lIndex) Is cAccountPosition Then
                    If (.RowData(lIndex).SymbolOrSymbolID = Bars.SymbolOrSymbolID) And (.RowData(lIndex).CurrentPositionSnapshot <> 0&) Then
                        strValue = ""
                        dCurrent = .RowData(lIndex).CurrentPrice(dPrice, dBid, dAsk)
                        
                        If dCurrent <> kNullData Then strValue = Bars.PriceDisplay(dCurrent)
                        ChangeCell fgAccountPositions, lIndex, AccountPosCol(eGDAccountPositionCol_CurrentPrice), strValue
                        
                        dProfit = .RowData(lIndex).OpenProfit(dCurrent)
                        .TextMatrix(lIndex, AccountPosCol(eGDAccountPositionCol_OpenProfit)) = dProfit
                        ColorCell fgAccountPositions, lIndex, AccountPosCol(eGDAccountPositionCol_OpenProfit)
                    End If
                End If
            Next lIndex
        End With
        
        If Not g.Broker.Account(m.Account.AccountID) Is Nothing Then
            Set m.Account = g.Broker.Account(m.Account.AccountID).MakeCopy
            RefreshAccountTotals
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.RefreshPrices"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ClearUpdatedColors
'' Description: Clear the updated colors on both grids if necessary
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ClearUpdatedColors()
On Error GoTo ErrSection:

    Dim lRow&, lCol&, dTickCount#, iSaveRedraw%
    Dim bStillColor As Boolean

    If m.bLoading = True Then Exit Sub

    With fgOpenOrders
        iSaveRedraw = .Redraw
        .Redraw = flexRDNone
        For lRow = .FixedRows To .Rows - 1
            If g.bUnloading Then Exit Sub
            If .Cell(flexcpForeColor, lRow, 0) = frmQuotes.UpdateColor Then
                bStillColor = False
                If frmTTSummary.tmrRealTime.Enabled Then
                    For lCol = OpenOrdersCol(eGDOpenOrdersCol_CurrentPrice) To OpenOrdersCol(eGDOpenOrdersCol_CurrentAsk)
                        If .Cell(flexcpForeColor, lRow, lCol) = frmQuotes.UpdateColor Then
                            ' see if has been more than 1 second since colored
                            dTickCount = .Cell(flexcpData, lRow, lCol)
                            dTickCount = gdTickCount - dTickCount
                            If dTickCount >= 0 And dTickCount <= 1000 Then
                                bStillColor = True
                            Else
                                .Cell(flexcpForeColor, lRow, lCol) = frmQuotes.UnchColor
                            End If
                        End If
                    Next lCol
                End If
                
                ' color symbol cell only if a cell was still colored
                If Not bStillColor Then
                    .Cell(flexcpForeColor, lRow, 0) = frmQuotes.UnchColor
                End If
            End If
        Next
        .Redraw = iSaveRedraw
    End With

    With fgTrades
        'iSaveRedraw = .Redraw
        '.Redraw = flexRDNone
        lCol = TradeCol(eGDTradeCol_ExitPrice)
        For lRow = .Rows - 1 To .FixedRows Step -1
            If g.bUnloading Then Exit Sub
            If TypeOf .RowData(lRow) Is cTradeLine Then
                If .RowData(lRow).IsOpen = False Then
                    If .Cell(flexcpForeColor, lRow, 0) = frmQuotes.UpdateColor Then
                        .Cell(flexcpForeColor, lRow, lCol) = frmQuotes.UnchColor
                    Else
                        Exit For
                    End If
                Else
                    If .Cell(flexcpForeColor, lRow, 0) = frmQuotes.UpdateColor Then
                        bStillColor = False
                        If frmTTSummary.tmrRealTime.Enabled Then
                            If .Cell(flexcpForeColor, lRow, lCol) = frmQuotes.UpdateColor Then
                                ' see if has been more than 1 second since colored
                                dTickCount = .Cell(flexcpData, lRow, lCol)
                                dTickCount = gdTickCount - dTickCount
                                If dTickCount >= 0 And dTickCount <= 1000 Then
                                    bStillColor = True
                                Else
                                    .Cell(flexcpForeColor, lRow, lCol) = frmQuotes.UnchColor
                                End If
                            End If
                        End If
                        
                        ' color symbol cell only if a cell was still colored
                        If Not bStillColor Then
                            .Cell(flexcpForeColor, lRow, 0) = frmQuotes.UnchColor
                        End If
                    End If
                End If
            End If
        Next
        '.Redraw = iSaveRedraw
    End With

    With fgAccountPositions
        iSaveRedraw = .Redraw
        .Redraw = flexRDNone
        lCol = AccountPosCol(eGDAccountPositionCol_CurrentPrice)
        For lRow = .FixedRows To .Rows - 1
            If g.bUnloading Then Exit Sub
            If .Cell(flexcpForeColor, lRow, 0) = frmQuotes.UpdateColor Then
                bStillColor = False
                If frmTTSummary.tmrRealTime.Enabled Then
                    If .Cell(flexcpForeColor, lRow, lCol) = frmQuotes.UpdateColor Then
                        ' see if has been more than 1 second since colored
                        dTickCount = .Cell(flexcpData, lRow, lCol)
                        dTickCount = gdTickCount - dTickCount
                        If dTickCount >= 0 And dTickCount <= 1000 Then
                            bStillColor = True
                        Else
                            .Cell(flexcpForeColor, lRow, lCol) = frmQuotes.UnchColor
                        End If
                    End If
                End If
                
                ' color symbol cell only if a cell was still colored
                If Not bStillColor Then
                    .Cell(flexcpForeColor, lRow, 0) = frmQuotes.UnchColor
                End If
            End If
        Next
        .Redraw = iSaveRedraw
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTTPositions.ClearUpdatedColors"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshTradeRules
'' Description: Refresh the trade rules
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RefreshTradeRules()
On Error GoTo ErrSection:

    ' Reload the trade rules collection and combo boxes...
    m.TradeRules.Load
        
    ' Reload the grid in case an abbreviation or something has changed...
    Set m.Positions = g.Broker.FillSummariesForAccount(m.lAccountID).MakeCopy
    LoadTradesGrid
    RefreshTrades

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.RefreshTradeRules"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateJournal
'' Description: Update the given journal entry
'' Inputs:      Journal ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdateJournal(ByVal lJournalID As Long)
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim Journal As cBrokerMessage       ' Journal object
    Dim lIndex As Long                  ' Index into a for loop
    
    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrderJournal] WHERE [JournalID]=" & Str(lJournalID) & ";", dbOpenDynaset)
    If Not (rs.BOF And rs.EOF) Then
        Set Journal = mTradeTracker.RecordsetToBrokerMessage(rs)
        AddJournal Journal
    Else
        Set Journal = Nothing
    End If
    
    With fgTransactions
        For lIndex = .FixedRows To .Rows - 1
            If .TextMatrix(lIndex, TransactionCol(eGDTransactionCol_OrderID)) = Journal("OrderID") Then
                JournalToGrid Journal, lIndex
            End If
        Next lIndex
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.UpdateJournal"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboAccountType_Click
'' Description: Allow the user to change account types
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboAccountType_Click()
On Error GoTo ErrSection:

    Dim nAccountType As eTT_AccountType ' Account type
    
    nAccountType = cboAccountType.ItemData(cboAccountType.ListIndex)
    Select Case nAccountType
        Case eTT_AccountType_SimStream
            chkBrokerOCO.Value = vbUnchecked
        
            If (m.Account.AccountID = 0&) Then
                If (Len(Trim(txtAccountNumber.Text)) = 7) Then
                    If (Left(Trim(UCase(txtAccountNumber.Text)), 3) = "GEN") Then
                        txtAccountNumber.Text = ""
                    End If
                End If
            End If
            
            If Len(txtAccountNumber.Text) = 0 Then
                txtAccountNumber.Text = g.SimTradeStream.NextAccount
            End If
        
        Case eTT_AccountType_SimBroker
            chkBrokerOCO.Value = vbUnchecked
            
            If (m.Account.AccountID = 0&) Then
                If (Len(Trim(txtAccountNumber.Text)) = 7) Then
                    If (Left(Trim(UCase(txtAccountNumber.Text)), 3) = "SIM") Then
                        txtAccountNumber.Text = ""
                    End If
                End If
            End If
            
            If (Len(txtAccountNumber.Text) = 0) Then
                txtAccountNumber.Text = g.SimTradeTs.NextAccount
            End If
            
            If Visible Then
                InfBox "Warning: Fills on the Genesis SimBroker are based on real-time data.  If you are trading off of delayed data, results may not be as desired", "!", , "Account Warning"
            End If
        
        Case eTT_AccountType_SimReplay
        
        Case Else
            If (m.Account.AccountID = 0&) Then
                If (Len(Trim(txtAccountNumber.Text)) = 7) Then
                    If (Left(Trim(UCase(txtAccountNumber.Text)), 3) = "SIM") Or (Left(Trim(UCase(txtAccountNumber.Text)), 3) = "GEN") Then
                        txtAccountNumber.Text = ""
                    End If
                End If
                
                CheckBoxValue(chkBrokerOCO) = g.Broker.BrokerAllowsOCO(cboAccountType.ItemData(cboAccountType.ListIndex))
            End If
    
    End Select

    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.cboAccountType_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkDateRange_Click
'' Description: Allow the user to only show fills for a date range
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkDateRange_Click()
On Error GoTo ErrSection:

    If Visible Then
        FilterTransactions
        EnableControls
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.chkDateRange_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkShowFlat_Click
'' Description: Show/Hide flat positions as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkShowFlat_Click()
On Error GoTo ErrSection:
    
    If Visible Then
        FilterPositions
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.chkShowFlat_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkShowJournal_Click
'' Description: Filter the transactions grid as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkShowJournal_Click()
On Error GoTo ErrSection:

    If Visible = True Then
        FilterTransactions
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.chkShowJournal_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdAccount_Click
'' Description: Allow the user to go off and edit or select another account
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdAccount_Click()
On Error GoTo ErrSection:

    Dim lAccountID As Long              ' Account selected by the user

    lAccountID = frmTTAccounts.ShowFromTrades(m.lAccountID)
    If lAccountID >= 0& Then
        ShowMe lAccountID
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.cmdAccount_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdAdjustment_Click
'' Description: Add a new adjustment to the account
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdAdjustment_Click()
On Error GoTo ErrSection:
    
    Select Case tabPositions.CurrTab
        Case Tabs(eGDTradeTrackerTab_Account)
            g.Broker.BrokerInfo(m.Account.AccountType).SendToLog
        
        Case Tabs(eGDTradeTrackerTab_Orders)
            SubmitAllOrdersFromGrid fgOpenOrders, "Trade Tracker"
            
        Case Tabs(eGDTradeTrackerTab_Transactions)
            NewAdjustment
            
        Case Tabs(eGDTradeTrackerTab_Trades)
            NewAdjustment
            
        Case Tabs(eGDTradeTrackerTab_Positions)
            FlattenPositionFromGrid fgAccountPositions, "Trade Tracker"
            
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.cmdAdjustment_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdBuy_Click
'' Description: Allow the user to enter an order
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdBuy_Click()
On Error GoTo ErrSection:

    CreateOrder "", m.lAccountID, 1, , , "Trade Tracker"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.cmdBuy_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancelOrder_Click
'' Description: Allow the user to cancel the selected order
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancelOrder_Click()
On Error GoTo ErrSection:

    Select Case tabPositions.CurrTab
        Case Tabs(eGDTradeTrackerTab_Account)
            ShowMe 0&, m.Account.AccountType
        
        Case Tabs(eGDTradeTrackerTab_Orders)
            CancelOrderFromGrid fgOpenOrders, "Trade Tracker", True
            
        Case Tabs(eGDTradeTrackerTab_Transactions)
            ExportTransactions
            
    End Select
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.cmdCancelOrder_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCheckStatus_Click
'' Description: Allow the user to check their SimTrade status
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCheckStatus_Click()
On Error GoTo ErrSection:

    g.Broker.CheckTradeServerOrders

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.cmdCheckStatus_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdEdit_Click
'' Description: Edit the currently selected trade
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdEdit_Click()
On Error GoTo ErrSection:

    Select Case tabPositions.CurrTab
        Case Tabs(eGDTradeTrackerTab_Account)
            g.Broker.ShowBrokerConnectionInfo AccountType, False, m.Account.UserName, False
        
        Case Tabs(eGDTradeTrackerTab_Orders)
            EditOrderFromGrid fgOpenOrders, "Trade Tracker"
            
        Case Tabs(eGDTradeTrackerTab_Transactions)
            EditTransactionRow
            
        Case Tabs(eGDTradeTrackerTab_Trades)
            ShowReports
            
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.cmdEdit_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdExit_Click
'' Description: When the user clicks on the Exit button, Unload the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdExit_Click()
On Error GoTo ErrSection:

    If Not AskToSave Then
        Unload Me
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.cmdExit_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdNewTrade_Click
'' Description: If the user clicks on New Trade, allow them to enter in a new
''              trade
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdNewTrade_Click()
On Error GoTo ErrSection:

    Select Case tabPositions.CurrTab
        Case Tabs(eGDTradeTrackerTab_Account)
            SaveAccountInfo
        
        Case Tabs(eGDTradeTrackerTab_Transactions)
            NewFill
            
        Case Tabs(eGDTradeTrackerTab_Trades)
            ExportTrades
        
        Case Tabs(eGDTradeTrackerTab_Positions)
            ReversePositionFromGrid fgAccountPositions, "Trade Tracker"
            
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.cmdNewTrade_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdDelete_Click
'' Description: Remove the currently selected trade
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdDelete_Click()
On Error GoTo ErrSection:

    Select Case tabPositions.CurrTab
        Case Tabs(eGDTradeTrackerTab_Account)
            g.Broker.ShowAccountInfoForm txtAccountNumber.Text
            
        Case Tabs(eGDTradeTrackerTab_Transactions)
            DeleteTransactionRow
            
    End Select
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.cmdDelete_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdPrint_Click
'' Description: Remove the currently selected trade
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
    RaiseError "frmTTPositions.cmdPrint_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdSell_Click
'' Description: Allow the user to enter an order
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdSell_Click()
On Error GoTo ErrSection:

    CreateOrder "", m.lAccountID, 0, , , "Trade Tracker"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.cmdSell_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdTradesFilter_Click
'' Description: Allow the user to change the trades filter
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdTradesFilter_Click()
On Error GoTo ErrSection:

    If frmTradeReportFilter.ShowForSettings(m.TradeFilter) Then
        FilterTrades
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.cmdTradesFilter.Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgAccountPositions_BeforeMouseDown
'' Description: If the user right clicks, show the popup menu
'' Inputs:      Button Pressed, Shift/Ctrl/Alt Status, Mouse Location
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgAccountPositions_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim lMouseRow As Long               ' Current mouse row of the grid
    Dim AcctPos As cAccountPosition     ' Account position from the grid
    
    With fgAccountPositions
        lMouseRow = .MouseRow
        
        If Button = vbRightButton Then
            If (lMouseRow >= .FixedRows) And (lMouseRow < .Rows) Then
                .Row = lMouseRow
                .RowSel = lMouseRow
                
                Set AcctPos = .RowData(.Row)
                mnuPositionsFlatten.Enabled = (AcctPos.CurrentPositionSnapshot <> 0)
                mnuPositionsReverse.Enabled = (AcctPos.CurrentPositionSnapshot <> 0)
            Else
                mnuPositionsFlatten.Enabled = False
                mnuPositionsReverse.Enabled = False
            End If
            
            PopupMenu mnuPositions
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.fgAccountPositions_BeforeMouseDown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgAccountPositions_Click
'' Description: If the user clicks on the Flatten column, flatten the position
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgAccountPositions_Click()
On Error GoTo ErrSection:

    Dim lMouseCol As Long               ' Mouse column in the grid
    Dim lMouseRow As Long               ' Mouse row in the grid
    Dim bMismatch As Boolean            ' Is the symbol currently in a mismatch?
    
    With fgAccountPositions
        lMouseCol = .MouseCol
        lMouseRow = .MouseRow
        
        If (lMouseRow >= .FixedRows) And (lMouseRow < .Rows) Then
            If TypeOf .RowData(lMouseRow) Is cAccountPosition Then
                bMismatch = (g.Broker.PositionMatch(.RowData(lMouseRow).AccountID, .RowData(lMouseRow).SymbolOrSymbolID) = False)
                
                If lMouseCol = AccountPosCol(eGDAccountPositionCol_Flatten) Then
                    If (.RowData(lMouseRow).CurrentPositionSnapshot <> 0&) Or (bMismatch = True) Then
                        FlattenPositionFromGrid fgAccountPositions, "Trade Tracker"
                    End If
                ElseIf lMouseCol = AccountPosCol(eGDAccountPositionCol_Reverse) Then
                    If (.RowData(lMouseRow).CurrentPositionSnapshot <> 0&) Or (bMismatch = True) Then
                        ReversePositionFromGrid fgAccountPositions, "Trade Tracker"
                    End If
                End If
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.fgAccountPositions_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgAccountPositions_MouseMove
'' Description: Show an appropriate tool tip for where the mouse is
'' Inputs:      Button Clicked, Shift/Ctrl/Alt Status, Location of Mouse
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgAccountPositions_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    Dim lMouseCol As Long               ' Mouse column in the grid
    Dim lMouseRow As Long               ' Mouse row in the grid
    
    With fgAccountPositions
        lMouseCol = .MouseCol
        lMouseRow = .MouseRow
        
        .ToolTipText = ""
        If (lMouseRow >= .FixedRows) And (lMouseRow < .Rows) Then
            If lMouseCol = AccountPosCol(eGDAccountPositionCol_Flatten) Then
                .ToolTipText = "Click here to flatten the position"
            ElseIf lMouseCol = AccountPosCol(eGDAccountPositionCol_Reverse) Then
                .ToolTipText = "Click here to reverse the position"
            End If
        End If
    End With

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgAccountPositions_RowColChange
'' Description: Enable/Disable controls as the row or column change
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgAccountPositions_RowColChange()
On Error GoTo ErrSection:

    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.fgAccountPositions_RowColChange"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgClosedOrders_AfterSort
'' Description: After sorting, recolor the background color of each row
'' Inputs:      Column Sorted, Order Sorted
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgClosedOrders_AfterSort(ByVal Col As Long, Order As Integer)
On Error GoTo ErrSection:

    SetBackColors fgClosedOrders

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.fgClosedOrders_AfterSort"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgClosedOrders_BeforeMouseDown
'' Description: If the user right clicks, show the popup menu
'' Inputs:      Button Pressed, Shift/Ctrl/Alt Status, Mouse Location
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgClosedOrders_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim lMouseRow As Long               ' Current mouse row in the grid
    
    With fgClosedOrders
        lMouseRow = .MouseRow
        
        If lMouseRow >= .FixedRows And lMouseRow < .Rows Then
            .Row = lMouseRow
            .RowSel = lMouseRow
            
            If Button = vbRightButton Then
                mnuEditOrder.Enabled = False
                mnuCancelOrder.Enabled = False
                mnuSubmitOrder.Enabled = False
                mnuSubmitAll.Enabled = False
                mnuParkOrder.Enabled = False
                mnuOrderHistory.Enabled = True
                mnuOrderJournal.Enabled = True
                
                mnuOrders.Tag = "Closed"
                PopupMenu mnuOrders
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.fgClosedOrders_BeforeMouseDown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgActivityLog_BeforeMouseDown
'' Description: If the user right clicks, show the popup menu
'' Inputs:      Button Pressed, Shift/Ctrl/Alt Status, Mouse Location
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgActivityLog_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:

    If Button = vbRightButton Then
        PopupMenu mnuActivityLog
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.fgActivityLog_BeforeMouseDown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgAccountPositions_Click
'' Description: If the user clicks on the X, cancel the order
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgOpenOrders_Click()
On Error GoTo ErrSection:

    Dim lMouseRow As Long               ' Current mouse row in the grid
    Dim lMouseCol As Long               ' Current mouse column in the grid
    
    With fgOpenOrders
        lMouseRow = .MouseRow
        lMouseCol = .MouseCol
        
        If (lMouseRow >= .FixedRows) And (lMouseRow < .Rows) Then
            If (lMouseCol = OpenOrdersCol(eGDOpenOrdersCol_Cancel)) Then
                .Row = lMouseRow
                
                CancelOrderFromGrid fgOpenOrders, "Trade Tracker", True
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.fgOpenOrders_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgAccountPositions_DblClick
'' Description: If the user double clicks on the symbol, set the active chart
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgAccountPositions_DblClick()
On Error GoTo ErrSection:

    If fgAccountPositions.MouseCol = AccountPosCol(eGDAccountPositionCol_Symbol) Then
        SetActiveChartSymbol fgAccountPositions.TextMatrix(fgAccountPositions.MouseRow, AccountPosCol(eGDAccountPositionCol_Symbol))
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.fgAccountPositions_DblClick"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgTrades_AfterEdit
'' Description: Fix the entry and exit rules if chosen
'' Inputs:      Row, Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgTrades_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    Dim TradeLine As cTradeLine         ' Trade line object
    Dim lComboData As Long              ' Combo data from the grid
    Dim Fill As cPtFill                 ' Fill object

    Set Fill = New cPtFill

    If Visible Then
        Select Case Col
            Case TradeCol(eGDTradeCol_EntryRule)
                If TypeOf fgTrades.RowData(Row) Is cTradeLine Then
                    Set TradeLine = fgTrades.RowData(Row)
                    lComboData = CLng(Val(fgTrades.ComboData))
                    If InStr(fgTrades.TextMatrix(Row, Col), "-") <> 0 Then
                        fgTrades.TextMatrix(Row, Col) = Parse(fgTrades.TextMatrix(Row, Col), "-", 1)
                    End If
                    If lComboData <> TradeLine.EntryRuleID Then
                        Screen.MousePointer = vbHourglass
                        TradeLine.EntryRuleID = lComboData
                        TradeLine.Save
                        
                        FilterTrades
                        
                        If Fill.Load(TradeLine.EntryFillID) Then
                            Fill.EntryRuleIdCategory = lComboData
                            Fill.Save
                            
                            If g.Broker.DateIsSnapshotForFill(Fill) Then
                                g.Broker.AddFill Fill, False
                            Else
                                g.Broker.RefreshManualFill Fill, , False
                            End If
                        End If
                        Screen.MousePointer = vbDefault
                    End If
                End If
            
            Case TradeCol(eGDTradeCol_ExitRule)
                If TypeOf fgTrades.RowData(Row) Is cTradeLine Then
                    Set TradeLine = fgTrades.RowData(Row)
                    lComboData = CLng(Val(fgTrades.ComboData))
                    If InStr(fgTrades.TextMatrix(Row, Col), "-") <> 0 Then
                        fgTrades.TextMatrix(Row, Col) = Parse(fgTrades.TextMatrix(Row, Col), "-", 1)
                    End If
                    If lComboData <> TradeLine.ExitRuleID Then
                        Screen.MousePointer = vbHourglass
                        TradeLine.ExitRuleID = CLng(Val(fgTrades.ComboData))
                        TradeLine.Save
                        
                        FilterTrades
                        
                        If Fill.Load(TradeLine.ExitFillID) Then
                            Fill.ExitRuleIdCategory = lComboData
                            Fill.Save
                            
                            If g.Broker.DateIsSnapshotForFill(Fill) Then
                                g.Broker.AddFill Fill, False
                            Else
                                g.Broker.RefreshManualFill Fill, , False
                            End If
                        End If
                        Screen.MousePointer = vbDefault
                    End If
                End If
                If InStr(fgTrades.TextMatrix(Row, Col), "-") <> 0 Then
                    fgTrades.TextMatrix(Row, Col) = Parse(fgTrades.TextMatrix(Row, Col), "-", 1)
                End If
                
            Case TradeCol(eGDTradeCol_Flag)
                If TypeOf fgTrades.RowData(Row) Is cTradeLine Then
                    Set TradeLine = fgTrades.RowData(Row)
                    lComboData = CLng(Val(fgTrades.ComboData))
                    If lComboData <> TradeLine.RealSimFlag Then
                        TradeLine.RealSimFlag = lComboData
                        TradeLine.Save
                        
                        FilterTrades
                        
                        If Fill.Load(TradeLine.ExitFillID) Then
                            Fill.RealSimFlagCategory = lComboData
                            Fill.Save
                            
                            If g.Broker.DateIsSnapshotForFill(Fill) Then
                                g.Broker.AddFill Fill
                            Else
                                g.Broker.RefreshManualFill Fill, , False
                            End If
                        End If
                    End If
                End If
            
        End Select
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.fgTrades_AfterEdit"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgTrades_BeforeEdit
'' Description: Only allow the user to edit the trade filter and flag columns
'' Inputs:      Row, Column, Cancel?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgTrades_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim astrFlags As cGdArray           ' Flags array
    Dim TradeLine As cTradeLine         ' Trade Line object

    Select Case Col
        Case TradeCol(eGDTradeCol_EntryRule)
            If TypeOf fgTrades.RowData(Row) Is cTradeLine Then
                Set TradeLine = fgTrades.RowData(Row)
                fgTrades.ComboList = m.TradeRules.EntryComboString(TradeLine.EntryRuleID)
            Else
                fgTrades.ComboList = ""
            End If
        
        Case TradeCol(eGDTradeCol_ExitRule)
            If TypeOf fgTrades.RowData(Row) Is cTradeLine Then
                Set TradeLine = fgTrades.RowData(Row)
                fgTrades.ComboList = m.TradeRules.ExitComboString(TradeLine.ExitRuleID)
            Else
                fgTrades.ComboList = ""
            End If
        
        Case TradeCol(eGDTradeCol_Flag)
            If TypeOf fgTrades.RowData(Row) Is cTradeLine Then
                Set astrFlags = New cGdArray
                
                For lIndex = 0 To m.astrFlags.Size - 1
                    astrFlags(lIndex) = "#" & Str(lIndex) & ";" & m.astrFlags(lIndex)
                Next lIndex
                
                fgTrades.ComboList = astrFlags.JoinFields("|")
            Else
                fgTrades.ComboList = ""
            End If
        
        Case Else
            Cancel = True
            
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.fgTrades_BeforeEdit"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgTrades_ComboCloseUp
'' Description: Tell the control to finish editing so that AfterEdit fires
'' Inputs:      Row, Column, Finish Edit?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgTrades_ComboCloseUp(ByVal Row As Long, ByVal Col As Long, FinishEdit As Boolean)
On Error GoTo ErrSection:

    FinishEdit = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.fgTrades_ComboCloseUp"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgTransactions_BeforeMouseDown
'' Description: If the user right clicks, show the popup menu
'' Inputs:      Button Pressed, Shift/Ctrl/Alt Status, Mouse Location
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgTransactions_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim lMouseRow As Long               ' Current mouse row in the grid
    
    With fgTransactions
        lMouseRow = .MouseRow
        
        If Button = vbRightButton Then
            If (lMouseRow >= .FixedRows + 1) And (lMouseRow < .Rows) Then
                .Row = lMouseRow
                .RowSel = lMouseRow
                
                If TypeOf .RowData(lMouseRow) Is cAccountAdjustment Then
                    mnuTransactionsNewAdjustment.Enabled = True
                    mnuTransactionsNewFill.Enabled = True
                    mnuTransactionsEdit.Caption = "Edit Adjustment"
                    mnuTransactionsEdit.Enabled = True
                    mnuTransactionsDelete.Caption = "Delete Adjustment"
                    mnuTransactionsDelete.Enabled = True
                    
                    mnuTransactionsEditJournal.Enabled = False
                    mnuTransactionsDeleteJournal.Enabled = False
                Else
                    mnuTransactionsNewAdjustment.Enabled = True
                    mnuTransactionsNewFill.Enabled = True
                    mnuTransactionsEdit.Caption = "Edit Fill"
                    mnuTransactionsEdit.Enabled = True
                    mnuTransactionsDelete.Caption = "Delete Fill"
                    mnuTransactionsDelete.Enabled = True
                    
                    If (Len(.TextMatrix(.Row, TransactionCol(eGDTransactionCol_JournalID))) > 0) Then
                        mnuTransactionsEditJournal.Caption = "Edit Journal"
                        mnuTransactionsDeleteJournal.Enabled = True
                    Else
                        mnuTransactionsEditJournal.Caption = "Add Journal"
                        mnuTransactionsDeleteJournal.Enabled = False
                    End If
                End If
            Else
                mnuTransactionsNewAdjustment.Enabled = True
                mnuTransactionsNewFill.Enabled = True
                mnuTransactionsEdit.Caption = "Edit"
                mnuTransactionsEdit.Enabled = False
                mnuTransactionsDelete.Caption = "Delete"
                mnuTransactionsDelete.Enabled = False
                    
                mnuTransactionsEditJournal.Enabled = False
                mnuTransactionsDeleteJournal.Enabled = False
            End If
            
            mnuTransactionsExport.Enabled = VisibleRows(fgTransactions) > 0
            
            PopupMenu mnuTransactions
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.fgTransactions_BeforeMouseDown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgTransactions_DblClick
'' Description: If the user double clicks on the symbol grid, set the active
''              chart to that symbol, otherwise edit the current fill
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgTransactions_DblClick()
On Error GoTo ErrSection:

    Dim lMouseRow As Long               ' Current mouse row from the grid
    Dim lMouseCol As Long               ' Current mouse column from the grid
    
    With fgTransactions
        lMouseRow = .MouseRow
        lMouseCol = .MouseCol
        
        If lMouseRow >= .FixedRows And lMouseRow < .Rows Then
            .Row = lMouseRow
            .RowSel = lMouseRow
            
            If lMouseCol = TransactionCol(eGDTransactionCol_Symbol) Then
                SetActiveChartSymbol .TextMatrix(lMouseRow, TransactionCol(eGDTransactionCol_Symbol))
            Else
                EditTransactionRow
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.fgTransactions_DblClick"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgTransactions_KeyDown
'' Description: If the user hits delete, delete the fill, if the user hits
''              insert then bring up a new trade
'' Inputs:      Code of the Key hit, Shift/Ctrl/Alt status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgTransactions_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value from an InfBox
    
    Select Case KeyCode
        Case vbKeyDelete
            DeleteTransactionRow
        
        Case vbKeyInsert
            strReturn = InfBox("Would you like to insert a new fill or a new adjustment?", "?", "+Fill|Adjustment|-Cancel", "Insert Transaction")
            Select Case strReturn
                Case "F"
                    NewFill
                    
                Case "A"
                    NewAdjustment
                    
            End Select
            
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.fgTransactions_KeyDown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgTransactions_KeyPress
'' Description: If the user hits enter, edit the trade for the selected fill
'' Inputs:      Ascii value of the key pressed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgTransactions_KeyPress(KeyAscii As Integer)
On Error GoTo ErrSection:

    Select Case KeyAscii
        Case vbKeyReturn
            If fgTransactions.Row >= fgTransactions.FixedRows And fgTransactions.Row < fgTransactions.Rows Then
                EditTransactionRow
            End If
            
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.fgTransactions_KeyPress"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgOpenOrders_AfterRowColChange
'' Description: As the user changes entries in the orders grid, load the
''              history grid appropriately
'' Inputs:      Old Row and Column, New Row and Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgOpenOrders_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.fgOpenOrders_AfterRowColChange"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgOpenOrders_BeforeMouseDown
'' Description: If the user right clicks, show the popup menu
'' Inputs:      Button Pressed, Shift/Ctrl/Alt Status, Mouse Location
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgOpenOrders_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim lMouseRow As Long               ' Current mouse row in the grid
    Dim Order As cPtOrder               ' Order object
    
    With fgOpenOrders
        lMouseRow = .MouseRow
        
        If lMouseRow >= .FixedRows And lMouseRow < .Rows Then
            .Row = lMouseRow
            .RowSel = lMouseRow
            
            If Button = vbRightButton Then
                Set Order = .RowData(.Row)
                
                mnuEditOrder.Enabled = True
                mnuCancelOrder.Enabled = True
                mnuSubmitOrder.Enabled = ((Order.Status = eTT_OrderStatus_Open) Or (Order.Status = eTT_OrderStatus_Parked))
                mnuParkOrder.Enabled = IsOpenOrder(Order.Status, False)
                mnuSubmitAll.Enabled = HasParkedOrders
                mnuOrderHistory.Enabled = True
                mnuOrderJournal.Enabled = True
                
                mnuOrders.Tag = "Open"
                PopupMenu mnuOrders
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.fgOpenOrders_BeforeMouseDown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgOpenOrders_DblClick
'' Description: If the user double clicks on the symbol, set the active chart,
''              otherwise edit the selected order
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgOpenOrders_DblClick()
On Error GoTo ErrSection:

    Dim lMouseRow As Long               ' Current mouse row from the grid
    Dim lMouseCol As Long               ' Current mouse column from the grid
    
    With fgOpenOrders
        lMouseRow = .MouseRow
        lMouseCol = .MouseCol
        
        If lMouseRow >= .FixedRows And lMouseRow < .Rows Then
            .Row = lMouseRow
            .RowSel = lMouseRow
            
            If lMouseCol = OpenOrdersCol(eGDOpenOrdersCol_Symbol) Then
                SetActiveChartSymbol .TextMatrix(.Row, OpenOrdersCol(eGDOpenOrdersCol_Symbol))
            Else
                EditOrderFromGrid fgOpenOrders, "Trade Tracker"
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.fgOpenOrders_DblClick"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgOpenOrders_KeyDown
'' Description: Delete the order on Delete hit, New Order on Insert hit
'' Inputs:      Code of the Key hit, Shift/Ctrl/Alt status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgOpenOrders_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    Select Case KeyCode
        Case vbKeyDelete
            If VisibleRows(fgOpenOrders) > 0 Then
                ''DeleteOrder
            End If
            
        Case vbKeyInsert
            'NewOrder
            CreateOrder "", m.lAccountID, , , , "Trade Tracker"
            
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.fgOpenOrders_KeyDown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgOpenOrders_KeyPress
'' Description: If the user hits enter, edit the selected order
'' Inputs:      Ascii Value of the Key Pressed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgOpenOrders_KeyPress(KeyAscii As Integer)
On Error GoTo ErrSection:
    
    Select Case KeyAscii
        Case vbKeyReturn
            EditOrderFromGrid fgOpenOrders, "Trade Tracker"
            
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.fgOpenOrders_KeyPress"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgTrades_AfterRowColChange
'' Description: After the user changes a row, display the notes for the trade
''              that is on the newly selected row
'' Inputs:      Old Row and Column, New Row and Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgTrades_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    If OldRow <> NewRow Then
        With fgTrades
            If NewRow <= .FixedRows Then
                lblNotes.Caption = "Notes for Account:"
            ElseIf TypeOf .RowData(NewRow) Is cAccountAdjustment Then
                lblNotes.Caption = "Notes for " & .TextMatrix(NewRow, TradeCol(eGDTradeCol_Symbol)) & _
                    " on " & .Cell(flexcpTextDisplay, NewRow, TradeCol(eGDTradeCol_ExitDate))
            Else
                lblNotes.Caption = "Notes for " & .Cell(flexcpTextDisplay, NewRow, TradeCol(eGDTradeCol_EntryDate)) & _
                    " trade for " & .TextMatrix(NewRow, TradeCol(eGDTradeCol_Symbol)) & ":"
            End If
            If NewRow >= .FixedRows Then
                txtNotes.Text = .TextMatrix(NewRow, TradeCol(eGDTradeCol_Notes))
            End If
        End With
    End If
    
    If fgTrades.Redraw <> flexRDNone Then EnableControls
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.fgTrades_AfterRowColChange"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgTrades_AfterSort
'' Description: After a sort, change the fixed rows back to 1
'' Inputs:      Column to sort, Order to sort the column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgTrades_AfterSort(ByVal Col As Long, Order As Integer)
On Error GoTo ErrSection:

    fgTrades.Redraw = flexRDBuffered

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.fgTrades_AfterSort"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgTrades_BeforeMouseDown
'' Description: If the user right clicks on the grid, bring up the popup menu
'' Inputs:      Button pressed, Shift/Ctrl/Alt status, Location of Click,
''              Whether to Cancel
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgTrades_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim lMouseRow As Long               ' Current mouse row in the grid
    Dim lMouseCol As Long               ' Current mouse column in the grid
    
    With fgTrades
        lMouseRow = .MouseRow
        lMouseCol = .MouseCol
        
        If Button = vbRightButton Then
            If lMouseRow >= .FixedRows And lMouseRow < .Rows Then
                .RowSel = lMouseRow
                If Not .IsSelected(lMouseRow) Then .Row = lMouseRow
            End If
            
            Enable mnuTradesExport, VisibleRows(fgTrades) > 0
        
            PopupMenu mnuTrades
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.fgTrades_BeforeMouseDown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgTrades_BeforeSort
'' Description: Before a sort, change the fixed rows to 2 so that the first
''              non-fixed row stays where it is regardless of the sort
'' Inputs:      Column to sort, Order to sort the column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgTrades_BeforeSort(ByVal Col As Long, Order As Integer)
On Error GoTo ErrSection:

    fgTrades.Redraw = flexRDNone

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.fgTrades_BeforeSort"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgTrades_DblClick
'' Description: If the user double clicks on the grid, chart the active symbol
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgTrades_DblClick()
On Error GoTo ErrSection:

    Dim lMouseRow As Long               ' Current mouse row in the grid
    Dim lMouseCol As Long               ' Current mouse column in the grid

    With fgTrades
        lMouseRow = .MouseRow
        lMouseCol = .MouseCol
        
        If lMouseRow >= .FixedRows And lMouseRow < .Rows Then
            If TypeOf .RowData(lMouseRow) Is cAccountAdjustment Then
                EditAdjustment .RowData(lMouseRow)
            ElseIf lMouseCol = TradeCol(eGDTradeCol_Symbol) Then
                SetActiveChartSymbol .TextMatrix(.Row, TradeCol(eGDTradeCol_Symbol))
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.fgTrades_DblClick"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgTrades_KeyDown
'' Description: Delete the trade on Delete hit, New Trade on Insert hit
'' Inputs:      Code of the Key hit, Shift/Ctrl/Alt status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgTrades_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If KeyCode = vbKeyDelete Then
        ''DeleteTrade
    ElseIf KeyCode = vbKeyInsert Then
        ''NewTrade
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.fgTrades_KeyDown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgTrades_KeyPress
'' Description: If the user hits enter on the grid, chart the active symbol
'' Inputs:      Key pressed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgTrades_KeyPress(KeyAscii As Integer)
On Error GoTo ErrSection:
    
    If KeyAscii = vbKeyReturn Then
        With fgTrades
            If .Row >= .FixedRows Then
                ''EditTrade
            End If
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.fgTrades_KeyPress"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgTrades_MouseMove
'' Description: Set the tool tip text as appropriate as the user moves the mouse
'' Inputs:      Button pressed, Shift/Ctrl/Alt status, Location of Mouse
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgTrades_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    Dim lMouseRow As Long               ' Current mouse row in the grid
    Dim lMouseCol As Long               ' Current mouse column in the grid
    Dim TradeLine As cTradeLine         ' Trade line object
    Static lLastMouseRow As Long        ' Last mouse row
    Static lLastMouseCol As Long        ' Last mouse col
    
    With fgTrades
        lMouseRow = .MouseRow
        lMouseCol = .MouseCol
        
        If lMouseRow <> lLastMouseRow Or lMouseCol <> lLastMouseCol Then
            lLastMouseRow = lMouseRow
            lLastMouseCol = lMouseCol
            
            If lMouseRow >= 0 And lMouseCol >= 0 Then
                If lMouseRow < .FixedRows And lMouseRow >= 0 Then
                    .ToolTipText = SORT_BY_PREFIX & .TextMatrix(lMouseRow, lMouseCol)
                ElseIf lMouseCol = TradeCol(eGDTradeCol_EntryRule) Then
                    If TypeOf .RowData(lMouseRow) Is cTradeLine Then
                        Set TradeLine = .RowData(lMouseRow)
                        .ToolTipText = m.TradeRules.DescriptionForID(TradeLine.EntryRuleID, eGDTradeRuleType_Entry)
                    End If
                ElseIf lMouseCol = TradeCol(eGDTradeCol_ExitRule) Then
                    If TypeOf .RowData(lMouseRow) Is cTradeLine Then
                        Set TradeLine = .RowData(lMouseRow)
                        .ToolTipText = m.TradeRules.DescriptionForID(TradeLine.ExitRuleID, eGDTradeRuleType_Exit)
                    End If
                Else
                    .ToolTipText = ""
                End If
            Else
                .ToolTipText = ""
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.fgTrades_MouseMove"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgTransactions_RowColChange
'' Description: Enable/Disable controls as the row or column change
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgTransactions_RowColChange()
On Error GoTo ErrSection:

    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.fgTransactions_RowColChange"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Activate
'' Description: Do some things when the form gets the focus
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Activate()
On Error GoTo ErrSection:

    Select Case tabPositions.CurrTab
        Case Tabs(eGDTradeTrackerTab_Trades)
            MoveFocus fgTrades
            
        Case Tabs(eGDTradeTrackerTab_Orders)
            MoveFocus fgOpenOrders
            
        Case Tabs(eGDTradeTrackerTab_Transactions)
            MoveFocus fgTransactions
            
        Case Tabs(eGDTradeTrackerTab_Positions)
            MoveFocus fgAccountPositions
            
    End Select
    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.Form_Activate"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_KeyDown
'' Description: Show the help if the user presses F1 on the form
'' Inputs:      Code of the Key Pressed, Shift/Ctrl/Alt Status
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
    RaiseError "frmTTPositions.Form_KeyDown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: When the form gets loaded, do some initialization
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Dim strFont As String               ' Grid font to use from the INI file
    Dim strTemp As String               ' Temporary string
    
    g.Styler.StyleForm Me

    LoadAccountTypeCombo

    Icon = Picture16(ToolbarIcon("ID_TradeTracker"), , True)
    Caption = "Trade Tracker"
    
    m.bClearInfBox = False
    
    tmrBrokers.Interval = 1000
    tmrBrokers.Enabled = False
    
    tmrMenu.Interval = 100
    tmrMenu.Enabled = False
    
    With txtNotes
        .BackColor = Me.BackColor
        .Locked = True
    End With
    
    mnuOrders.Visible = False
    mnuTransactions.Visible = False
    mnuTrades.Visible = False
    mnuPositions.Visible = False
    mnuActivityLog.Visible = False
    
    strFont = GetIniFileProperty("TTPositions", "", "Fonts", g.strIniFile)
    If strFont <> "" Then FontFromString fgTrades.Font, strFont

    strTemp = GetIniFileProperty("TTPositions", "", "Placement", g.strIniFile)
    If Len(strTemp) = 0 Then
        Width = 11730
        CenterTheForm Me
    Else
        SetFormPlacement Me, strTemp
    End If
    
    ' Closed order filter information...
    txtNumDays = GetIniFileProperty("NumDays", 7, "TTPositions", g.strIniFile)
    
    ' Transactions filter information...
    chkDateRange = GetIniFileProperty("FillsDateRange", vbUnchecked, "TTPositions", g.strIniFile)
    gdFillsFromDate = GetIniFileProperty("FillsFromDate", Date - 30, "TTPositions", g.strIniFile)
    gdFillsToDate = GetIniFileProperty("FillsToDate", Date, "TTPositions", g.strIniFile)
    chkShowJournal = GetIniFileProperty("ShowJournal", vbUnchecked, "TTPositions", g.strIniFile)
    
    Set m.TradeRules = New cTradeRules
    m.TradeRules.Load
    
    Set m.Journals = New cGdTree
    Set m.JournalOrderMap = New cGdTree
    
    Set m.astrFlags = New cGdArray
    m.astrFlags.Create eGDARRAY_Strings
    m.astrFlags.Add "None", 0
    m.astrFlags.Add "Real", 1
    m.astrFlags.Add "Sim", 2
    
    Set m.astrTradeSymbols = New cGdArray
    m.astrTradeSymbols.Create eGDARRAY_Strings
    
    fraTradeFilter.Height = 795 ' 1035
    
    ' Account Positions filter information...
    chkShowFlat.Value = GetIniFileProperty("ShowFlat", vbUnchecked, "TTPositions", g.strIniFile)

    txtTotalAdjustments.Locked = True
    txtTotalAdjustments.BackColor = vbButtonFace
    txtTotalFees.Locked = True
    txtTotalFees.BackColor = vbButtonFace
    txtTotalClosedProfit.Locked = True
    txtTotalClosedProfit.BackColor = vbButtonFace
    txtClosedBalance.Locked = True
    txtClosedBalance.BackColor = vbButtonFace
    txtTotalOpenEquity.Locked = True
    txtTotalOpenEquity.BackColor = vbButtonFace
    txtCurrentValue.Locked = True
    txtCurrentValue.BackColor = vbButtonFace
    
    cmdCheckStatus.ToolTipText = "Connect to Genesis SimTrade Server and check order status"
    
    txtNumDays.Alignment = vbRightJustify
    txtNumDays.Visible = False
    
    ' Default this to False so as to say that the user is not editing the starting balance...
    m.bEditingSB = False
    
    ' DAJ 07/01/2010: Don't show the security type check boxes nor the futures fees
    ' controls if this is running in any Extreme Charts mode...
    If ExtremeCharts = 1 Then
        fraSecTypes.Visible = False
        lblFutureFees.Visible = False
        txtFutureFees.Visible = False
        lblFutureFeesDesc.Visible = False
        
        fraCommission.Height = 615
        fraFillMatch.Top = 3060
    Else
        fraSecTypes.Visible = True
        lblFutureFees.Visible = True
        txtFutureFees.Visible = True
        lblFutureFeesDesc.Visible = True
        
        fraCommission.Height = 915
        fraFillMatch.Top = 4200
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.Form_Load"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: When the user clicks on the "X", close the form
'' Inputs:      Cancel the Unload?, Mode of the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If m.bLoading = True Then
        Cancel = True
    ElseIf UnloadMode <> vbFormCode Then
        Cancel = AskToSave
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.Form_QueryUnload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Resize
'' Description: When the form gets resized, resize/move the controls on the
''              form appropriately
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Resize()
On Error Resume Next

    'If LimitFormSize(Me, fraButtons.Width * 5, fraButtons.Height + txtNotes.Height + (fraButtons.Top * 3) + lblNotes.Height) Then Exit Sub
    'If LimitFormSize(Me, 10515, fraButtons.Height + txtNotes.Height + (fraButtons.Top * 3) + lblNotes.Height) Then Exit Sub
    'If LimitFormSize(Me, 12360, fraButtons.Height + txtNotes.Height + (fraButtons.Top * 3) + lblNotes.Height) Then Exit Sub
    If LimitFormSize(Me, 12360, 7170) Then Exit Sub
    
    With picRithmic
        .Move tabPositions.Left, ScaleHeight - .Height - tabPositions.Top
    End With
    
    With picPbo
        .Move ScaleWidth - .Width - tabPositions.Left, ScaleHeight - .Height - tabPositions.Top
    End With
    
    With tabPositions
        If picRithmic.Visible Then
            .Move .Left, .Top, ScaleWidth - fraButtons.Width - (.Left * 3), _
                    ScaleHeight - picRithmic.Height - (.Top * 3)
        Else
            .Move .Left, .Top, ScaleWidth - fraButtons.Width - (.Left * 3), _
                    ScaleHeight - (.Top * 2)
        End If
        .Refresh
    End With
    
    With fraButtons
        .Move tabPositions.Width + (tabPositions.Left * 2), tabPositions.Top
    End With
    
    With fraTradeFilter
        .Move .Left, .Top, tabPositions.ClientWidth - (.Left * 2)
    End With
    
    With fgTrades
        .Move .Left, (fraTradeFilter.Top * 2) + fraTradeFilter.Height, _
            tabPositions.ClientWidth - (.Left * 2), _
            tabPositions.ClientHeight - txtNotes.Height - (fraTradeFilter.Top * 4) - lblNotes.Height - fraTradeFilter.Height
    End With
    
    With lblOpenEquity
        .Move fgTrades.Left + fgTrades.Width - .Width, _
            30 + fgTrades.Top + fgTrades.Height
    End With

    With lblNotes
        .Move fgTrades.Left, fgTrades.Height + fraTradeFilter.Top + fgTrades.Top, _
            tabPositions.ClientWidth - (.Left * 2)
    End With
    
    With txtNotes
        .Move fgTrades.Left, lblNotes.Top + lblNotes.Height, _
            tabPositions.ClientWidth - (.Left * 2)
    End With
    
    With fgOpenOrders
        .Move .Left, .Top, tabPositions.ClientWidth - (.Left * 2), _
            tabPositions.ClientHeight - .Top - lblClosedOrders.Height - fgClosedOrders.Height - (lblOpenOrders.Top * 2)
    End With
    
    With fraClosedOrders
        .Move .Left, fgOpenOrders.Top + fgOpenOrders.Height + lblOpenOrders.Top
    End With
    
    With fgClosedOrders
        .Move .Left, fraClosedOrders.Top + fraClosedOrders.Height, tabPositions.ClientWidth - (.Left * 2)
    End With
    
    With fgTransactions
        .Move .Left, .Top, tabPositions.ClientWidth - (.Left * 2), tabPositions.ClientHeight - .Top - chkDateRange.Top
    End With
    
    With fgAccountPositions
        .Move .Left, .Top, tabPositions.ClientWidth - (.Left * 2), tabPositions.ClientHeight - .Top - 120
    End With
    
    With fgActivityLog
        .Move .Left, .Top, tabPositions.ClientWidth - (.Left * 2), _
            tabPositions.ClientHeight - (.Top * 2)
    End With
    
    ExtendCustomColumn fgOpenOrders
    ExtendCustomColumn fgClosedOrders
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unloaded
'' Description: Cleanup and save settings when the form gets unloaded
'' Inputs:      Cancel the Unload?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    tmrBrokers.Enabled = False
    tmrMenu.Enabled = False
    
    g.ActivityLogs.Remove "frmTTPositions"

    SetIniFileProperty "TTPositions", FontToString(fgTrades.Font), "Fonts", g.strIniFile
    SetIniFileProperty "TTPositions", GetFormPlacement(Me), "Placement", g.strIniFile

    ' Closed order filter information...
    SetIniFileProperty "NumDays", Val(txtNumDays), "TTPositions", g.strIniFile
    
    ' Transactions filter information...
    SetIniFileProperty "FillsDateRange", chkDateRange, "TTPositions", g.strIniFile
    SetIniFileProperty "FillsFromDate", gdFillsFromDate.Value, "TTPositions", g.strIniFile
    SetIniFileProperty "FillsToDate", gdFillsToDate.Value, "TTPositions", g.strIniFile
    SetIniFileProperty "ShowJournal", chkShowJournal, "TTPositions", g.strIniFile
    
    ' Trades filter information...
    SetIniFileProperty "TradesFilter", m.TradeFilter.ToString(";", "|"), "TTPositions", g.strIniFile
    
    ' Account Positions filter information...
    SetIniFileProperty "ShowFlat", chkShowFlat.Value, "TTPositions", g.strIniFile
    
    Set m.HistoricalOrders = Nothing
    Set m.SnapshotOrders = Nothing
    Set m.HistoricalFills = Nothing
    Set m.SnapshotFills = Nothing
    Set m.Positions = Nothing
    
    Set m.Journals = Nothing
    Set m.JournalOrderMap = Nothing
    
    Set m.astrTradeSymbols = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.Form_Unload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    gdFillsFromDate_Changed
'' Description: When the from date is changed, re-filter the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub gdFillsFromDate_Changed()
On Error GoTo ErrSection:

    If Visible Then
        FilterTransactions
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.gdFillsFromDate_Changed"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    gdFillsToDate_Changed
'' Description: When the to date is changed, re-filter the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub gdFillsToDate_Changed()
On Error GoTo ErrSection:

    If Visible Then
        FilterTransactions
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.gdFillsToDate_Changed"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    gdStartDate_LostFocus
'' Description: Refresh the trade and transaction grids with the new date
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub gdStartDate_LostFocus()
On Error GoTo ErrSection:

    If gdStartDate <> m.Account.StartingDate Then
        If m.lAccountID <> 0& Then SaveAccountInfo
        
        If fgTransactions.Rows > fgTransactions.FixedRows Then
            fgTransactions.TextMatrix(1, TransactionCol(eGDTransactionCol_Date)) = gdStartDate.Value
        End If
    
        If fgTrades.Rows > fgTrades.FixedRows Then
            fgTrades.TextMatrix(1, TradeCol(eGDTradeCol_EntryDate)) = gdStartDate.Value
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.gdStartDate_LostFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuActivityLogChangeFont_Click
'' Description: Allow the user to change a font on the activity log grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuActivityLogChangeFont_Click()
On Error GoTo ErrSection:

    ChangeGridFont fgActivityLog

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.mnuActivityLogChangeFont_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuCancelOrder_Click
'' Description: Allow the user to cancel an open order
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuCancelOrder_Click()
On Error GoTo ErrSection:

    CancelOrderFromGrid fgOpenOrders, "Trade Tracker", True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.mnuCancelOrder_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuEditOrder_Click
'' Description: Allow the user to edit the selected order
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuEditOrder_Click()
On Error GoTo ErrSection:

    EditOrderFromGrid fgOpenOrders, "Trade Tracker"
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.mnuEditOrder_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuOrderHistory_Click
'' Description: Allow the user to view the history for an order
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuOrderHistory_Click()
On Error GoTo ErrSection:

    If mnuOrders.Tag = "Open" Then
        With fgOpenOrders
            If .Row >= .FixedRows And .Row < .Rows Then
                frmOrderHistory.ShowMe .RowData(.Row)
            End If
        End With
    Else
        With fgClosedOrders
            If .Row >= .FixedRows And .Row < .Rows Then
                frmOrderHistory.ShowMe .RowData(.Row)
            End If
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.mnuOrderHistory_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuOrderJournal_Click
'' Description: Allow the user to view the journal for an order
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuOrderJournal_Click()
On Error GoTo ErrSection:

    tmrMenu.Tag = "JOURNAL"
    tmrMenu.Enabled = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.mnuOrderJournal_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuOrdersChangeFont_Click
'' Description: Allow the user to change the font on the orders grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuOrdersChangeFont_Click()
On Error GoTo ErrSection:

    If mnuOrders.Tag = "Open" Then
        ChangeGridFont fgOpenOrders
    Else
        ChangeGridFont fgClosedOrders
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.mnuOrdersChangeFont_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuOrdersPrint_Click
'' Description: Allow the user to print the orders grids
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuOrdersPrint_Click()
On Error GoTo ErrSection:

    PrintMe

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.mnuOrdersPrint_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuParkOrder_Click
'' Description: Allow the user to park the selected order
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuParkOrder_Click()
On Error GoTo ErrSection:

    ParkOrderFromGrid fgOpenOrders, "Trade Tracker"
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.mnuParkOrder_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuPositionsChangeFont_Click
'' Description: Allow the user to change the font on the positions grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuPositionsChangeFont_Click()
On Error GoTo ErrSection:

    ChangeGridFont fgAccountPositions

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.mnuPositionsChangeFont_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuPositionsFlatten_Click
'' Description: Allow the user to flatten the chosen position
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuPositionsFlatten_Click()
On Error GoTo ErrSection:

    FlattenPositionFromGrid fgAccountPositions, "Trade Tracker"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.mnuPositionsFlatten_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuPositionsPrint_Click
'' Description: Allow the user to print the positions grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuPositionsPrint_Click()
On Error GoTo ErrSection:

    PrintMe

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.mnuPositionsPrint_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuPositionsReverse_Click
'' Description: Allow the user to reverse the chosen position
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuPositionsReverse_Click()
On Error GoTo ErrSection:

    ReversePositionFromGrid fgAccountPositions, "Trade Tracker"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.mnuPositionsReverse_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuSubmitAll_Click
'' Description: Allow the user to submit all parked orders
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuSubmitAll_Click()
On Error GoTo ErrSection:

    SubmitAllOrdersFromGrid fgOpenOrders, "Trade Tracker"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.mnuSubmitAll_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuSubmitOrder_Click
'' Description: Allow the user to submit the selected order
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuSubmitOrder_Click()
On Error GoTo ErrSection:

    SubmitOrderFromGrid fgOpenOrders, "Trade Tracker"
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.mnuSubmitOrder_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuTradesChangeFont_Click
'' Description: Allow the user to change the font on the Trades grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuTradesChangeFont_Click()
On Error GoTo ErrSection:

    ChangeGridFont fgTrades

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.mnuTradesChangeFont_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuTradesExport_Click
'' Description: Allow the user to export the trades grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuTradesExport_Click()
On Error GoTo ErrSection:

    ExportTrades

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.mnuTradesExport_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuTradesNewAdjustment_Click
'' Description: Allow the user to enter a new adjustment
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuTradesNewAdjustment_Click()
On Error GoTo ErrSection:

    NewAdjustment

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.mnuTradesNewAdjustment_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuTradesPrint_Click
'' Description: Allow the user to print the trades grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuTradesPrint_Click()
On Error GoTo ErrSection:

    PrintMe

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.mnuTradesPrint_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuTransactionsChangeFont_Click
'' Description: Allow the user to change the font on the transactions grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuTransactionsChangeFont_Click()
On Error GoTo ErrSection:

    ChangeGridFont fgTransactions

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.mnuTransactionsChangeFont_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuTransactionsDelete_Click
'' Description: Allow the user to delete a transaction
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuTransactionsDelete_Click()
On Error GoTo ErrSection:

    DeleteTransactionRow

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.mnuTransactionsDelete_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuTransactionsDeleteJournal_Click
'' Description: Allow the user to delete a journal entry
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuTransactionsDeleteJournal_Click()
On Error GoTo ErrSection:

    Dim lJournalID As Long              ' ID of the Journal to delete
    Dim lIndex As Long                  ' Index into a for loop

    With fgTransactions
        If (.Row >= .FixedRows) And (.Row < .Rows) Then
            lJournalID = CLng(Val(.TextMatrix(.Row, TransactionCol(eGDTransactionCol_JournalID))))
            If lJournalID > 0 Then
                If InfBox("Are you sure you want to delete this journal entry?", "?", "+Yes|-No", "Journal Delete Confirmation") = "Y" Then
                    g.JournalBridge.DeleteOrderJournal lJournalID
                    RemoveJournal lJournalID
                    
                    For lIndex = .FixedRows To .Rows - 1
                        If .TextMatrix(lIndex, TransactionCol(eGDTransactionCol_JournalID)) = Str(lJournalID) Then
                            .TextMatrix(lIndex, TransactionCol(eGDTransactionCol_JournalID)) = ""
                            .TextMatrix(lIndex, TransactionCol(eGDTransactionCol_EmotionNumber)) = ""
                            .TextMatrix(lIndex, TransactionCol(eGDTransactionCol_Feelings)) = ""
                            .TextMatrix(lIndex, TransactionCol(eGDTransactionCol_Reasons)) = ""
                            .TextMatrix(lIndex, TransactionCol(eGDTransactionCol_Thoughts)) = ""
                            .TextMatrix(lIndex, TransactionCol(eGDTransactionCol_Note)) = ""
                        End If
                    Next lIndex
                End If
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.mnuTransactionsDeleteJournal_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuTransactionsEdit_Click
'' Description: Allow the user to edit a transaction
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuTransactionsEdit_Click()
On Error GoTo ErrSection:

    EditTransactionRow

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.mnuTransactionsEdit_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuTransactionsEditJournal_Click
'' Description: Allow the user to edit a journal from the transaction
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuTransactionsEditJournal_Click()
On Error GoTo ErrSection:

    tmrMenu.Tag = "JOURNAL"
    tmrMenu.Enabled = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.mnuTransactionsEditJournal_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuTransactionsExport_Click
'' Description: Allow the user to export the transactions grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuTransactionsExport_Click()
On Error GoTo ErrSection:

    ExportTransactions

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.mnuTransactionsExport_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuTransactionsNewAdjustment_Click
'' Description: Allow the user to create a new adjustment
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuTransactionsNewAdjustment_Click()
On Error GoTo ErrSection:

    NewAdjustment

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.mnuTransactionNewAdjustment_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuTransactionsNewFill_Click
'' Description: Allow the user to enter a new fill
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuTransactionsNewFill_Click()
On Error GoTo ErrSection:

    NewFill

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.mnuTransactionsNewFill_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuTransactionsPrint_Click
'' Description: Allow the user to print the transactions grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuTransactionsPrint_Click()
On Error GoTo ErrSection:

    PrintMe

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.mnuTransactionsPrint_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuViewJournals_Click
'' Description: Allow the user to view their journal entries
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuViewJournals_Click()
On Error GoTo ErrSection:

    tmrMenu.Tag = "JOURNALS"
    tmrMenu.Enabled = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.mnuViewJournals_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optFifo_Click
'' Description: Recalculate stuff based on first-in, first-out if the user
''              changes to this option
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optFifo_Click()
On Error GoTo ErrSection:

    If (Visible = True) And (m.Account.FillMatchMode <> eTT_FillMatchMode_Fifo) Then
        If m.Account.AccountID <> 0& Then
            m.Account.FillMatchMode = eTT_FillMatchMode_Fifo
            m.Account.Save
            
            g.Broker.UpdateAccount m.Account
        End If
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTTPositions.optFifo_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optLifo_Click
'' Description: Recalculate stuff based on last-in, first-out if the user
''              changes to this option
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optLifo_Click()
On Error GoTo ErrSection:

    If (Visible = True) And (m.Account.FillMatchMode <> eTT_FillMatchMode_Lifo) Then
        If m.Account.AccountID <> 0& Then
            m.Account.FillMatchMode = eTT_FillMatchMode_Lifo
            m.Account.Save
            
            g.Broker.UpdateAccount m.Account
        End If
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTTPositions.optLifo_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tabPositions_Click
'' Description: Move the focus and enable controls as applicable
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tabPositions_Click()
On Error GoTo ErrSection:

    Select Case tabPositions.CurrTab
        Case Tabs(eGDTradeTrackerTab_Trades)
            MoveFocus fgTrades
            
        Case Tabs(eGDTradeTrackerTab_Orders)
            MoveFocus fgOpenOrders
            
        Case Tabs(eGDTradeTrackerTab_Transactions)
            MoveFocus fgTransactions
            
        Case Tabs(eGDTradeTrackerTab_Positions)
            MoveFocus fgAccountPositions
            
    End Select
    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.tabPositions_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tabPositions_Switch
'' Description: Event raised when the user switches tabs
'' Inputs:      Old Tab, New Tab, Whether to Cancel switch
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tabPositions_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
On Error GoTo ErrSection:

    If Visible And m.bLoading = False Then
        If OldTab = Tabs(eGDTradeTrackerTab_Account) Then
            MoveFocus txtAccountName
            DoEvents
            
            If m.lAccountID <> 0& Then
                SaveAccountInfo
            Else
                Cancel = AskToSave
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.tabPositions_Switch"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tmrBrokers_Timer
'' Description: Check to see if there is new data available from the broker
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tmrBrokers_Timer()
On Error GoTo ErrSection:

    Dim dLastChanged As Double          ' Last changed value from the broker info object
    Dim bDidSomething As Boolean        ' Did we do something?
    
    TimerStart "frmTTPositions.tmrBrokers"
    bDidSomething = False
    If m.bReload Then
        LoadGrids AccountType
        m.bReload = False
        bDidSomething = True
    Else
        dLastChanged = g.Broker.LastChanged(g.Broker.AccountTypeForID(m.lAccountID))
        If m.dLastChanged < dLastChanged Then
            If Not g.Broker.Account(m.lAccountID) Is Nothing Then
                Set m.Account = g.Broker.Account(m.lAccountID).MakeCopy
                Set m.SnapshotOrders = g.Broker.OrdersForAccount(m.lAccountID).MakeCopy
                Set m.SnapshotFills = g.Broker.FillsForAccount(m.lAccountID).MakeCopy
                Set m.Positions = g.Broker.FillSummariesForAccount(m.lAccountID).MakeCopy
                
                RefreshAccountTotals
                
                ResetOpenOrderRemoveFlags
                RefreshOpenOrders
                RemoveOpenOrders
                
                ResetClosedOrderRemoveFlags
                RefreshClosedOrders
                RemoveClosedOrders
                
                ResetTransactionRemoveFlags
                RefreshTransactions
                RemoveTransactions
                
                ResetTradeRemoveFlags
                RefreshTrades
                RemoveTrades
                
                ResetPositionRemoveFlags
                RefreshAccountPositions
                RemovePositions
                
                m.dLastChanged = dLastChanged
                bDidSomething = True
            End If
        End If
    End If
        
    If (m.bClearInfBox = True) And (bDidSomething = True) Then
        InfBox ""
        m.bClearInfBox = False
    End If
    TimerEnd "frmTTPositions.tmrBrokers", tmrBrokers.Interval

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.tmrBrokers_Timer"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tmrMenu_Timer
'' Description: Perform appropriate menu item
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tmrMenu_Timer()
On Error GoTo ErrSection:

    Dim strTag As String                ' Tag of the timer
    Dim lOrderID As Long                ' Order ID
    Dim Order As cPtOrder               ' Order object
    
    TimerStart "frmTTPositions.tmrMenu"
    strTag = tmrMenu.Tag
    tmrMenu.Tag = ""
    tmrMenu.Enabled = False
    
    Select Case UCase(strTag)
        Case "JOURNAL"
            If tabPositions.CurrTab = Tabs(eGDTradeTrackerTab_Orders) Then
                If mnuOrders.Tag = "Open" Then
                    With fgOpenOrders
                        If (.Row >= .FixedRows) And (.Row < .Rows) Then
                            If TypeOf .RowData(.Row) Is cPtOrder Then
                                g.TnJournal.ShowOrderJournal .RowData(.Row)
                            End If
                        End If
                    End With
                Else
                    With fgClosedOrders
                        If (.Row >= .FixedRows) And (.Row < .Rows) Then
                            If TypeOf .RowData(.Row) Is cPtOrder Then
                                g.TnJournal.ShowOrderJournal .RowData(.Row)
                            End If
                        End If
                    End With
                End If
            ElseIf tabPositions.CurrTab = Tabs(eGDTradeTrackerTab_Transactions) Then
                Set Order = New cPtOrder
                With fgTransactions
                    lOrderID = CLng(Val(.TextMatrix(.Row, TransactionCol(eGDTransactionCol_OrderID))))
                
                    If Order.Load(lOrderID) Then
                        g.TnJournal.ShowJournalForTransaction Order
                    End If
                End With
            End If
        
        
        Case "JOURNALS"
            g.TnJournal.ShowJournals
    
    End Select
    TimerEnd "frmTTPositions.tmrMenu", tmrMenu.Interval

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.tmrMenu_Timer"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtAccountName_GotFocus
'' Description: When the control gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtAccountName_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtAccountName

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.txtAccountName_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtAccountName_LostFocus
'' Description: When the control loses the focus, trim the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtAccountName_LostFocus()
On Error GoTo ErrSection:

    txtAccountName.Text = Trim(txtAccountName.Text)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.txtAccountName_LostFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtAccountNumber_GotFocus
'' Description: When the control gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtAccountNumber_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtAccountNumber

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.txtAccountNumber_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtAccountNumber_LostFocus
'' Description: When the control loses the focus, trim the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtAccountNumber_LostFocus()
On Error GoTo ErrSection:

    txtAccountNumber.Text = Trim(txtAccountNumber.Text)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.txtAccountNumber_LostFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtBroker_GotFocus
'' Description: When the control gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtBroker_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtBroker

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.txtBroker_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtFutureFees_GotFocus
'' Description: When the control gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtFutureFees_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtFutureFees

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.txtFutureFees_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtFutureFees_LostFocus
'' Description: When the control loses the focus, format the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtFutureFees_LostFocus()
On Error GoTo ErrSection:

    txtFutureFees.Text = Format(ValOfText(txtFutureFees.Text), "$#,##0.00")

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.txtFutureFees_LostFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtNumDays_GotFocus
'' Description: When the control gets the focus, select and save the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtNumDays_GotFocus()
On Error GoTo ErrSection:

    m.strNumDaysSave = txtNumDays.Text
    SelectAll txtNumDays

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.txtNumDays_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtNumDays_LostFocus
'' Description: When the control loses the focus, verify the contents
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtNumDays_LostFocus()
On Error GoTo ErrSection:

    txtNumDays.Text = Trim(txtNumDays.Text)

    If IsNumeric(txtNumDays.Text) = False Then
        InfBox "Number of days must be numeric", "!", , "Error"
        txtNumDays.Text = m.strNumDaysSave
        MoveFocus txtNumDays
    ElseIf Val(txtNumDays.Text) < 0 Then
        InfBox "Number of days must be greater than or equal to zero", "!", , "Error"
        txtNumDays.Text = m.strNumDaysSave
        MoveFocus txtNumDays
    ElseIf Len(txtNumDays.Text) = 0 Then
        InfBox "Number of days must be filled in", "!", , "Error"
        txtNumDays.Text = m.strNumDaysSave
        MoveFocus txtNumDays
    ElseIf txtNumDays.Text <> m.strNumDaysSave Then
        FilterClosedOrders
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.txtNumDays_LostFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtStartingBalance_GotFocus
'' Description: When the control gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtStartingBalance_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtStartingBalance
    
    ' Set this flag to True because the user may be going to edit the Starting Balance...
    m.bEditingSB = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.txtStartingBalance_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtStartingBalance_LostFocus
'' Description: When the control loses the focus, format the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtStartingBalance_LostFocus()
On Error GoTo ErrSection:

    txtStartingBalance.Text = Format(ValOfText(txtStartingBalance.Text), "$#,##0.00")
    If m.Account.StartingBalance <> ValOfText(txtStartingBalance.Text) Then
        If m.lAccountID = 0& Then
            m.Account.StartingBalance = ValOfText(txtStartingBalance.Text)
        Else
            SaveAccountInfo
        End If
        
        RefreshAccountTotals
        CalculateTransactionBalance
        CalculateTradeBalance False
        
        g.Broker.UpdateAccount m.Account
    End If
    
    ' Since the control is losing the focus, the user must be done editing the
    ' starting balance...
    m.bEditingSB = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.txtStartingBalance_LostFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtStockFees_GotFocus
'' Description: When the control gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtStockFees_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtStockFees

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.txtStockFees_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtStockFees_LostFocus
'' Description: When the control loses the focus, format the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtStockFees_LostFocus()
On Error GoTo ErrSection:

    txtStockFees.Text = Format(ValOfText(txtStockFees.Text), "$#,##0.00")

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.txtStockFees_LostFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EnableControls
'' Description: Enable/Disable controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EnableControls()
On Error GoTo ErrSection:

    Dim bEnable As Boolean              ' Enable or disable
    Dim nBroker As eTT_AccountType      ' Broker

    If m.Account.AccountID = 0& Then
        tabPositions.TabEnabled(Tabs(eGDTradeTrackerTab_Orders)) = False
        tabPositions.TabEnabled(Tabs(eGDTradeTrackerTab_Transactions)) = False
        tabPositions.TabEnabled(Tabs(eGDTradeTrackerTab_Trades)) = False
        tabPositions.TabEnabled(Tabs(eGDTradeTrackerTab_Positions)) = False
        tabPositions.TabEnabled(Tabs(eGDTradeTrackerTab_ActivityLog)) = False
    Else
        tabPositions.TabEnabled(Tabs(eGDTradeTrackerTab_Orders)) = True
        tabPositions.TabEnabled(Tabs(eGDTradeTrackerTab_Transactions)) = True
        tabPositions.TabEnabled(Tabs(eGDTradeTrackerTab_Trades)) = True
        tabPositions.TabEnabled(Tabs(eGDTradeTrackerTab_Positions)) = True
        tabPositions.TabEnabled(Tabs(eGDTradeTrackerTab_ActivityLog)) = True
    End If

    Select Case tabPositions.CurrTab
        Case Tabs(eGDTradeTrackerTab_Account)
            nBroker = AccountType
            
            Enable lblNumber, (m.Account.AccountID = 0)
            Enable txtAccountNumber, (m.Account.AccountID = 0)
            Enable lblAccountType, (m.Account.AccountID = 0)
            Enable cboAccountType, (m.Account.AccountID = 0)
            
            Select Case nBroker
                Case eTT_AccountType_SimBroker
                    chkBrokerOCO.Visible = False
                    lblSimDesc.Visible = True
                    lblSimDesc.Caption = "Orders will be managed on the Genesis SimBroker server"
                
                Case eTT_AccountType_SimStream
                    chkBrokerOCO.Visible = False
                    lblSimDesc.Visible = True
                    lblSimDesc.Caption = "Orders will be managed in Trade Navigator"
                
                Case Else
                    chkBrokerOCO.Visible = g.Broker.BrokerAllowsOCO(nBroker)
                    lblSimDesc.Visible = False
            
            End Select
            
            cmdNewTrade.Visible = True
            cmdNewTrade.Caption = "Sa&ve"
            cmdNewTrade.Top = 0
            cmdEdit.Visible = g.Broker.IsLiveAccount(nBroker)
            cmdEdit.Enabled = cmdEdit.Visible
            cmdEdit.Caption = "Connection &Info"
            cmdEdit.Top = 420
            cmdDelete.Visible = False '((nBroker = eTT_AccountType_LindWaldock) Or (nBroker = eTT_AccountType_ManExpress))
            cmdDelete.Enabled = (cmdDelete.Visible = True) And (g.Broker.ConnectionStatusForAccount(txtAccountNumber.Text) = eGDConnectionStatus_Connected)
            cmdDelete.Caption = "Vie&w Details"
            cmdDelete.Top = 840
            
            cmdBuy.Visible = False
            cmdSell.Visible = False
            
            cmdAdjustment.Visible = FileExist(AddSlash(App.Path) & "DumpBinfo.FLG")
            cmdAdjustment.Enabled = True
            cmdAdjustment.Caption = "Dump BInfo"
            cmdAdjustment.Top = 1260
            
            cmdCancelOrder.Visible = True
            cmdCancelOrder.Caption = "Ne&w Account"
            cmdCancelOrder.Enabled = True
            
            Enable cmdNewTrade, True
            
            picRithmic.Visible = g.Broker.IsRithmicBroker(m.Account.AccountType)
            picPbo.Visible = g.Broker.IsRithmicBroker(m.Account.AccountType)
            
        Case Tabs(eGDTradeTrackerTab_Orders)
            cmdBuy.Visible = True
            cmdBuy.Caption = "&Buy"
            cmdBuy.Top = 0
            cmdSell.Visible = True
            cmdSell.Caption = "S&ell"
            cmdSell.Top = 0
            cmdEdit.Visible = True
            cmdEdit.Caption = "E&dit Order"
            cmdEdit.Top = 420
            cmdCancelOrder.Visible = True
            cmdCancelOrder.Caption = "&Cancel Order"
            cmdCancelOrder.Top = 840
            cmdAdjustment.Visible = True
            cmdAdjustment.Caption = "Sub&mit Orders"
            cmdAdjustment.Top = 1260
        
            cmdNewTrade.Visible = False
            cmdDelete.Visible = False
            
            Enable cmdBuy, True
            Enable cmdSell, True
            Enable cmdEdit, VisibleRows(fgOpenOrders) > 0
            Enable cmdCancelOrder, VisibleRows(fgOpenOrders) > 0
            Enable cmdAdjustment, HasParkedOrders
            
        Case Tabs(eGDTradeTrackerTab_Transactions)
            cmdAdjustment.Visible = True
            cmdAdjustment.Caption = "&New Adjustment"
            cmdAdjustment.Top = 0
            cmdNewTrade.Visible = True
            cmdNewTrade.Caption = "New &Fill"
            cmdNewTrade.Top = 420
            cmdEdit.Visible = True
            cmdEdit.Top = 840
            cmdDelete.Visible = True
            cmdDelete.Top = 1260
            cmdCancelOrder.Visible = True
            cmdCancelOrder.Caption = "E&xport"
            cmdCancelOrder.Top = 1680
            
            With fgTransactions
                If (.Row >= .FixedRows) And (.Row < .Rows) Then
                    If TypeOf .RowData(.Row) Is cPtFill Then
                        cmdEdit.Caption = "&Edit Fill"
                        cmdEdit.Enabled = True
                        cmdDelete.Caption = "&Delete Fill"
                        cmdDelete.Enabled = True
                    ElseIf TypeOf .RowData(.Row) Is cAccountAdjustment Then
                        cmdEdit.Caption = "&Edit Adjustment"
                        cmdEdit.Enabled = True
                        cmdDelete.Caption = "&Delete Adjustment"
                        cmdDelete.Enabled = True
                    Else
                        cmdEdit.Caption = "&Edit Fill"
                        cmdEdit.Enabled = False
                        cmdDelete.Caption = "&Delete Fill"
                        cmdDelete.Enabled = False
                    End If
                End If
            End With
            
            cmdBuy.Visible = False
            cmdSell.Visible = False
            
            Enable cmdAdjustment, True
            Enable cmdNewTrade, True
            Enable cmdCancelOrder, VisibleRows(fgTransactions) > 0
            
        Case Tabs(eGDTradeTrackerTab_Trades)
            cmdAdjustment.Enabled = True
            cmdAdjustment.Caption = "&New Adjustment"
            cmdAdjustment.Top = 0
            cmdNewTrade.Visible = True
            cmdNewTrade.Caption = "E&xport"
            cmdNewTrade.Top = 420
            cmdEdit.Visible = True
            cmdEdit.Caption = "R&eports"
            cmdEdit.Top = 840
            
            cmdBuy.Visible = False
            cmdSell.Visible = False
            cmdCancelOrder.Visible = False
            cmdDelete.Visible = False
            
            Enable cmdAdjustment, True
            Enable cmdNewTrade, VisibleRows(fgTrades) > 0
            Enable cmdEdit, VisibleRows(fgTrades) > 1
            
        Case Tabs(eGDTradeTrackerTab_Positions)
            cmdAdjustment.Visible = True
            cmdAdjustment.Caption = "&Flatten Position"
            cmdAdjustment.Top = 0
            cmdNewTrade.Visible = True
            cmdNewTrade.Caption = "Re&verse Position"
            cmdNewTrade.Top = 420
            
            cmdBuy.Visible = False
            cmdSell.Visible = False
            cmdEdit.Visible = False
            cmdCancelOrder.Visible = False
            cmdDelete.Visible = False
            
            bEnable = False
            With fgAccountPositions
                If (.Row >= .FixedRows) And (.Row < .Rows) Then
                    bEnable = (.RowData(.Row).CurrentPositionSnapshot <> 0)
                End If
            End With
            
            Enable cmdAdjustment, bEnable
            Enable cmdNewTrade, bEnable
            
        Case Tabs(eGDTradeTrackerTab_ActivityLog)
            cmdBuy.Visible = False
            cmdSell.Visible = False
            cmdAdjustment.Visible = False
            cmdNewTrade.Visible = False
            cmdEdit.Visible = False
            cmdCancelOrder.Visible = False
            cmdDelete.Visible = False
            
    End Select
    
    gdFillsFromDate.Enabled = (chkDateRange.Value = vbChecked)
    gdFillsToDate.Enabled = (chkDateRange.Value = vbChecked)
    
    cmdCheckStatus.Visible = Not g.Broker.IsLiveAccount(AccountType)
    cmdCheckStatus.Enabled = g.Broker.EnableCheckStatusMenu
    cmdPrint.Visible = True
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.EnableControls"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NewFill
'' Description: Allow the user to enter a new fill
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub NewFill()
On Error GoTo ErrSection:

    Dim Fill As New cPtFill             ' Fill object
    
    If g.Broker.CreateNewFill(Fill, "", m.lAccountID, 0&) Then
        If g.Broker.DateIsSnapshotForFill(Fill) = False Then
            LoadHistoricalFills
        End If
        
        m.bClearInfBox = True
        
        m.Account.Load m.Account.AccountID
        RefreshAccountTotals
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.NewFill"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NewAdjustment
'' Description: Allow the user to enter a new adjustment
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub NewAdjustment()
On Error GoTo ErrSection:

    If frmTTEditAdjustment.ShowMe(0&, m.lAccountID) Then
        RefreshAdjustments
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.NewAdjustment"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EditTransactionRow
'' Description: Allow the user to edit either a fill or an adjustment
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EditTransactionRow()
On Error GoTo ErrSection:

    With fgTransactions
        If (.Row >= .FixedRows) And (.Row < .Rows) Then
            If TypeOf .RowData(.Row) Is cPtFill Then
                EditFill .RowData(.Row)
            ElseIf TypeOf .RowData(.Row) Is cAccountAdjustment Then
                EditAdjustment .RowData(.Row)
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.EditTransactionRow"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EditAdjustment
'' Description: Allow the user to edit an existing adjustment
'' Inputs:      Adjustment
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EditAdjustment(Adjustment As cAccountAdjustment)
On Error GoTo ErrSection:

    If frmTTEditAdjustment.ShowMe(Adjustment.AdjustmentID, Adjustment.AccountID) Then
        RefreshAdjustments
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.EditAdjustment"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EditFill
'' Description: Allow the user to edit a fill
'' Inputs:      Fill
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EditFill(Fill As cPtFill)
On Error GoTo ErrSection:

    If g.Broker.ModifyFill(Fill) Then
        If g.Broker.DateIsSnapshotForFill(Fill) = False Then
            LoadHistoricalFills
        End If
        
        m.bClearInfBox = True
        
        m.Account.Load m.Account.AccountID
        RefreshAccountTotals
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.EditFill"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DeleteTransactionRow
'' Description: Allow the user to delete either a fill or an adjustment
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DeleteTransactionRow()
On Error GoTo ErrSection:

    With fgTransactions
        If .SelectedRows = 1 Then
            If (.Row >= .FixedRows) And (.Row < .Rows) Then
                If TypeOf .RowData(.Row) Is cPtFill Then
                    DeleteFill .RowData(.Row)
                ElseIf TypeOf .RowData(.Row) Is cAccountAdjustment Then
                    DeleteAdjustment .RowData(.Row)
                End If
            End If
        ElseIf .SelectedRows > 1 Then
            DeleteTransactionRows
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.DeleteTransactionRow"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DeleteAdjustment
'' Description: Allow the user to delete an existing adjustment
'' Inputs:      Adjustment
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DeleteAdjustment(Adjustment As cAccountAdjustment)
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    
    If InfBox("Are you sure you wish to delete the selected adjustment?", "?", "+Yes|-No", "Adjustment Delete Confirmation") = "Y" Then
        If Adjustment.Delete Then
            RefreshAdjustments
            
            g.Broker.BrokerDebug g.Broker.AccountTypeForID(m.lAccountID), "Adjustment has been manually deleted", True
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.DeleteAdjustment"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DeleteFill
'' Description: Allow the user to delete a fill
'' Inputs:      Fill
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DeleteFill(Fill As cPtFill)
On Error GoTo ErrSection:

    If g.Broker.DeleteFill(Fill, "Trade Tracker") Then
        If g.Broker.DateIsSnapshotForFill(Fill) = False Then
            LoadHistoricalFills
        End If
        
        m.bClearInfBox = True
        
        m.Account.Load m.Account.AccountID
        RefreshAccountTotals
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.DeleteFill"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshAdjustments
'' Description: Reload the adjustments and anything else as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RefreshAdjustments()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim dTotal As Double                ' Total amount of adjustment

    ' Reload all of the adjustments for this account from the database...
    LoadAdjustments
    
    ' Recalculate the total sum of adjustments for this account...
    dTotal = 0#
    For lIndex = 1 To m.Adjustments.Count
        dTotal = dTotal + m.Adjustments(lIndex).Amount
    Next lIndex
    
    ' Update the account with the new total sum of adjustments...
    m.Account.TotalAdjustments = dTotal
    m.Account.Save
    g.Broker.UpdateAccount m.Account
    RefreshAccountTotals
    
    ' Reload the transactions grid...
    LoadTransactionsGrid
    RefreshTransactions
    
    ' Reload the trades grid...
    LoadTradesGrid
    RefreshTrades

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTTPositions.RefreshAdjustments"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HasParkedOrders
'' Description: Are there parked SimTrade orders to submit?
'' Inputs:      None
'' Returns:     True if Parked SimTrade orders exist, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function HasParkedOrders() As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bReturn As Boolean              ' Return value from the function
    
    bReturn = False
    With fgOpenOrders
        For lIndex = .FixedRows To .Rows - 1
            If .RowData(lIndex).Status = eTT_OrderStatus_Parked Then
                bReturn = True
                Exit For
            End If
        Next lIndex
    End With
    
    HasParkedOrders = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTTPositions.HasParkedOrders"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    VisibleRows
'' Description: Count the number of visible rows for the given grid
'' Inputs:      Grid
'' Returns:     Number of Visible Rows
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function VisibleRows(fgGrid As VSFlexGrid) As Long
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lCount As Long                  ' Count of visible rows
    
    With fgGrid
        lCount = 0&
        For lIndex = .FixedRows To .Rows - 1
            If .RowHidden(lIndex) = False Then lCount = lCount + 1
        Next lIndex
    End With
    
    VisibleRows = lCount

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTTPositions.VisibleRows"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ExtendCustomColumn
'' Description: Adjust all column widths to accomodate the custom "extend column"
'' Inputs:      nResizeCol (passed only by the AfterUserResize event)
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ExtendCustomColumn(Grid As VSFlexGrid, Optional ByVal nResizeCol As Long = -1)
On Error GoTo ErrSection:
    
    Dim i&, nTotal&, nDiff&, nExtCol&
    
    nExtCol = OpenOrdersCol(eGDOpenOrdersCol_OrderText)
    
    With Grid
        ' if column being resized is after the extended column,
        ' then change the width of the next visible column instead
        If nResizeCol >= nExtCol Then
            .Redraw = flexRDNone
            'nDiff = .ColWidth(nResizeCol) - m.nPrevColWidth
            For i = nResizeCol + 1 To .Cols - 1
                If Not .ColHidden(i) Then
                    .ColWidth(i) = .ColWidth(i) - nDiff
                    Exit For
                End If
            Next
            'm.nPrevColWidth = 0
        End If
        
        ' size the custom extended column in order to fill the client width
        .ColHidden(nExtCol) = True
        .Redraw = flexRDBuffered '(must do this so .ClientWidth will be correct)
        .Redraw = flexRDNone
        nTotal = 0
        For i = 0 To .Cols - 1
            If Not .ColHidden(i) Then
                nTotal = nTotal + .ColWidth(i)
            End If
        Next
        nTotal = .ClientWidth - nTotal
        'If nTotal > 0 Then .ColWidth(nExtCol) = nTotal
        If nTotal > .ColWidth(nExtCol) Then .ColWidth(nExtCol) = nTotal
        .ColHidden(nExtCol) = False
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
       
ErrSection:
    RaiseError "frmTTPositions.ExtendCustomColumn"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ChangeCell
'' Description: To change text and forecolor of grid cell
'' Inputs:      Grid to change, Row and Column to change, New Value
'' Returns:     True if cell changed, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ChangeCell(Grid As VSFlexGrid, ByVal lRow&, ByVal lCol&, ByVal strCellText$) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim nForeColor As Long              ' Foreground color to color the cell
    Dim dTickCount As Double            ' Current tick count
    
    bReturn = False
    
    With Grid
        nForeColor = frmQuotes.UnchColor
        If .TextMatrix(lRow, lCol) <> strCellText Then
            .TextMatrix(lRow, lCol) = strCellText
            If frmTTSummary.tmrRealTime.Enabled Then
                nForeColor = frmQuotes.UpdateColor
                .Cell(flexcpForeColor, lRow, 0) = frmQuotes.UpdateColor
                .Cell(flexcpData, lRow, lCol) = gdTickCount
            End If
        
            bReturn = True
        ElseIf frmTTSummary.tmrRealTime.Enabled Then
            dTickCount = .Cell(flexcpData, lRow, lCol)
            dTickCount = gdTickCount - dTickCount
            If dTickCount >= 0 And dTickCount <= 1000 Then
                nForeColor = frmQuotes.UpdateColor
            End If
        End If
        
        .Cell(flexcpForeColor, lRow, lCol) = nForeColor
    End With
    
    ChangeCell = bReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTTPositions.ChangeCell"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ColorCell
'' Description: Color the given cell red or green depending on sign
'' Inputs:      Grid, Row and Column to color
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ColorCell(Grid As VSFlexGrid, ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    Dim dValue As Double                ' Value of the cell

    dValue = ValOfText(Grid.TextMatrix(Row, Col))
    
    If dValue < 0 Then
        Grid.Cell(flexcpForeColor, Row, Col) = vbRed
    ElseIf dValue = 0 Then
        Grid.Cell(flexcpForeColor, Row, Col) = vbBlack
    ElseIf g.nColorTheme = kDarkThemeColor And IsGreenRange(QBColor(2), True) Then
        Grid.Cell(flexcpForeColor, Row, Col) = vbGreen
    Else
        Grid.Cell(flexcpForeColor, Row, Col) = QBColor(2)
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.ColorCell"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetAccountTypeCombo
'' Description: Set the account type combo box appropriately
'' Inputs:      Account Type
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetAccountTypeCombo(ByVal nAccountType As eTT_AccountType)
On Error GoTo ErrSection:
    
    Dim lIndex As Long                  ' Index into a for loop
        
    For lIndex = 0 To cboAccountType.ListCount - 1
        If cboAccountType.ItemData(lIndex) = nAccountType Then
            cboAccountType.ListIndex = lIndex
            Exit For
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.SetAccountTypeCombo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadAccountTypeCombo
'' Description: Load up the account type combo box with appropriate account types
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadAccountTypeCombo()
On Error GoTo ErrSection:

    Dim bShowAccountCombo As Boolean    ' Should we show the account combo?

    bShowAccountCombo = g.Broker.LoadBrokerCombo(cboAccountType)
    
    ' Only show the account type combo box if there are non-simulated brokers...
    lblAccountType.Visible = bShowAccountCombo
    cboAccountType.Visible = bShowAccountCombo

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.LoadAccountTypeCombo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AccountType
'' Description: Determine the account type from the combo box
'' Inputs:      None
'' Returns:     Account Type
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AccountType() As eTT_AccountType
On Error GoTo ErrSection:

    Dim nReturn As eTT_AccountType

    nReturn = -1&
    If cboAccountType.ListIndex >= 0 Then
        nReturn = cboAccountType.ItemData(cboAccountType.ListIndex)
    End If
    
    AccountType = nReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTTPositions.AccountType"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SaveAccountInfo
'' Description: Save the account information
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveAccountInfo()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim dOldStarting As Double          ' Old Starting Balance

    If ValOfText(txtStartingBalance.Text) > (2 ^ 31) Then
        MoveFocus txtStartingBalance
        Err.Raise vbObjectError + 1000, , "Invalid Starting Balance"
    End If
    
    If Len(Trim(txtAccountNumber.Text)) = 0 Or Len(Trim(txtAccountNumber.Text)) > 20 Then
        MoveFocus txtAccountNumber
        Err.Raise vbObjectError + 1000, , "Account Number must be Between 1 and 20 Characters"
    End If

    If Len(Trim(txtAccountName.Text)) = 0 Then txtAccountName.Text = txtAccountNumber.Text
    
    If Len(Trim(txtAccountName.Text)) > 50 Then
        MoveFocus txtAccountName
        Err.Raise vbObjectError + 1000, , "Name must be Between 1 and 50 Characters"
    End If
    
    If InStr(txtAccountName.Text, "'") > 0 Then
        MoveFocus txtAccountName
        Err.Raise vbObjectError + 1000, , "Account Name cannot contain an apostrophe"
    End If
    
    If InStr(txtAccountNumber.Text, "'") > 0 Then
        MoveFocus txtAccountNumber
        Err.Raise vbObjectError + 1000, , "Account Number cannot contain an apostrophe"
    End If
    
    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccounts] " & _
            "WHERE [AccountNumber]='" & Trim(txtAccountNumber.Text) & "';", dbOpenDynaset)
    If Not (rs.BOF And rs.EOF) Then
        If rs!AccountID <> m.Account.AccountID Then
            MoveFocus txtAccountNumber
            Err.Raise vbObjectError + 1000, , "Account Number must be unique"
        End If
    End If

    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccounts] " & _
            "WHERE [Name]='" & Trim(txtAccountName.Text) & "';", dbOpenDynaset)
    If Not (rs.BOF And rs.EOF) Then
        If rs!AccountID <> m.Account.AccountID Then
            MoveFocus txtAccountName
            Err.Raise vbObjectError + 1000, , "Account Name must be unique"
        End If
    End If
    
    With m.Account
        dOldStarting = .StartingBalance
        .AccountNumber = Trim(txtAccountNumber.Text)
        .Name = Trim(txtAccountName.Text)
        .StartingBalance = ValOfText(txtStartingBalance.Text)
        .StartingDate = gdStartDate.Value
        .Broker = Trim(txtBroker.Text)
        .Comms = ValOfText(txtFutureFees.Text)
        .StockFees = ValOfText(txtStockFees.Text)
        .AccountType = cboAccountType.ItemData(cboAccountType.ListIndex)
        .HoldOcoAtBroker = CheckBoxValue(chkBrokerOCO)
        .SecTypeMask = SecurityTypeMask
        If optFifo.Value = True Then
            .FillMatchMode = eTT_FillMatchMode_Fifo
        ElseIf optLifo.Value = True Then
            .FillMatchMode = eTT_FillMatchMode_Lifo
        End If
        
        .Save
        
        If Len(.Name) > 0 Then
            Caption = "Trade Tracker - [" & .Name & "]"
        ElseIf Len(.AccountNumber) > 0 Then
            Caption = "Trade Tracker - [" & .AccountNumber & "]"
        Else
            Caption = "Trade Tracker - [New Account]"
        End If
    
        m.lAccountID = .AccountID
    End With
    
    g.Broker.UpdateAccount m.Account
    RefreshAccountCombos
    SendAccountToOptionNav m.Account, False
    
    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.SaveTabInfo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SecurityTypeMask
'' Description: Determine the security type mask from the check boxes
'' Inputs:      None
'' Returns:     Security Type Mask
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SecurityTypeMask() As Long
On Error GoTo ErrSection:

    Dim lSecTypeMask As Long            ' Secutity type mask

    SetBit lSecTypeMask, 1, (chkFutures.Value = vbChecked)
    SetBit lSecTypeMask, 2, (chkStocks.Value = vbChecked)
    SetBit lSecTypeMask, 3, (chkForex.Value = vbChecked)
    SetBit lSecTypeMask, 4, (chkFutOpts.Value = vbChecked)
    SetBit lSecTypeMask, 5, (chkStkOpts.Value = vbChecked)
    
    SecurityTypeMask = lSecTypeMask

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTTPositions.SecurityTypeMask"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AccountIsDirty
'' Description: Determine whether or not the user has made changes to account
'' Inputs:      None
'' Returns:     True if Dirty, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AccountIsDirty() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function

    With m.Account
        If txtAccountNumber.Text <> .AccountNumber Then
            bReturn = True
        ElseIf txtAccountName.Text <> .Name Then
            bReturn = True
        ElseIf ValOfText(txtStartingBalance.Text) <> .StartingBalance Then
            bReturn = True
        ElseIf gdStartDate.Value <> .StartingDate Then
            bReturn = True
        ElseIf txtBroker.Text <> .Broker Then
            bReturn = True
        ElseIf ValOfText(txtFutureFees.Text) <> .Comms Then
            bReturn = True
        ElseIf ValOfText(txtStockFees.Text) <> .StockFees Then
            bReturn = True
        ElseIf cboAccountType.ItemData(cboAccountType.ListIndex) <> .AccountType Then
            bReturn = True
        ElseIf SecurityTypeMask <> .SecTypeMask Then
            bReturn = True
        Else
            bReturn = False
        End If
    End With
    
    AccountIsDirty = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTTPositions.AccountIsDirty"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AskToSave
'' Description: If the account is dirty, ask the user if they wish to save
'' Inputs:      None
'' Returns:     True if user Cancelled the dialog, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AskToSave() As Boolean
On Error GoTo ErrSection:
    
    Dim strResponse As String           ' Response from the dialog
    
    If AccountIsDirty Then
        If WindowState = vbMinimized Then WindowState = vbNormal

        strResponse = InfBox("Do you want to save your changes?||Clicking No will undo any changes you have made to the strategy or any rules in the strategy.", "?", "+Yes|No|-Cancel", Caption)
        Select Case strResponse
            Case "C"
                AskToSave = True
            Case "Y"
                SaveAccountInfo
            Case "N"
                LoadInfoTab
        End Select
    End If
        
ErrExit:
    Exit Function

ErrSection:
    AskToSave = True
    RaiseError "frmTTPositions.AskToSave"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ColorTextBox
'' Description: Color the text box based on the value in it
'' Inputs:      Text Box
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ColorTextBox(tb As ctlUniTextBoxXP) 'RH was TextBox
On Error GoTo ErrSection:

    Select Case ValOfText(tb.Text)
        Case Is > 0
            'some text has white back color, only change fore color for ones that have dark back color
            If g.nColorTheme = kDarkThemeColor And tb.BackColor = Me.BackColor And IsGreenRange(QBColor(2), True) Then
                tb.ForeColor = vbGreen
            Else
                tb.ForeColor = QBColor(2)
            End If
        Case Is < 0
            tb.ForeColor = vbRed
        Case 0
            tb.ForeColor = lblCurrentValue.ForeColor
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.ColorTextBox"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadGrids
'' Description: Load up all of the grids
'' Inputs:      Account Type
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadGrids(Optional ByVal nAccountType As eTT_AccountType = eTT_AccountType_SimStream)
On Error GoTo ErrSection:

    'Dim rsOrders As Recordset           ' Orders from the database
    Dim rsFills As Recordset            ' Fills from the database
    Dim BInfo As cBrokerInfo            ' Broker Info object

    'Set rsOrders = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrders] WHERE [AccountID]=" & Str(m.lAccountID) & ";", dbOpenDynaset)
    Set rsFills = g.dbPaper.OpenRecordset("SELECT * FROM [tblFills] WHERE [AccountID]=" & Str(m.lAccountID) & ";", dbOpenDynaset)
    
    'Set m.HistoricalOrders = New cPtOrders
    'm.HistoricalOrders.LoadOrdersFromRecordset rsOrders, rsFills, False
    
    Set m.HistoricalFills = New cPtFills
    m.HistoricalFills.LoadFillsFromRecordset rsFills, False
    
    Set m.Adjustments = New cGdTree
    LoadAdjustments
    
    Set BInfo = g.Broker.BrokerInfo(m.Account.AccountType)
    If Not BInfo Is Nothing Then
        Set m.SnapshotOrders = BInfo.OrdersForAccount(m.Account.AccountNumber).MakeCopy
        Set m.SnapshotFills = BInfo.FillsForAccount(m.Account.AccountNumber).MakeCopy
        Set m.Positions = BInfo.FillSummariesForAccount(m.lAccountID).MakeCopy
    Else
        Set m.SnapshotOrders = New cPtOrders
        Set m.SnapshotFills = New cPtFills
        Set m.Positions = New cAccountPositions
    End If
        
    ' Load the account information...
    LoadInfoTab nAccountType
    
    If m.bLoading Then
        tabPositions.CurrTab = Tabs(eGDTradeTrackerTab_Account)
        
        EnableControls
        ShowForm Me, eForm_Nonmodal, , , ALT_GRID_ROW_COLOR
    End If
    
    ' Load all of the grids on the form...
    RefreshOpenOrders
    
    'LoadClosedOrdersGrid
    RefreshClosedOrders
    
    LoadTransactionsGrid False
    RefreshTransactions
    
    LoadTradesGrid False
    RefreshTrades
        
    RefreshAccountPositions
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.LoadGrids"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitOpenOrdersGrid
'' Description: Initialize the open orders grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitOpenOrdersGrid()
On Error GoTo ErrSection:

    Dim nRedraw As Long                 ' Current state of the grid's redraw
    
    With fgOpenOrders
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        .AllowBigSelection = False
        .AllowSelection = False
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .BackColorBkg = g.Styler.GetColor(eGrid_Background) 'RH override vbApplicationWorkspacevbButtonFace
        .Editable = flexEDNone
        .ExplorerBar = flexExSortShowAndMove
        .ExtendLastCol = False
        .ScrollBars = flexScrollBarBoth
        .ScrollTrack = True
        .SelectionMode = flexSelectionListBox
        .SheetBorder = RGB(128, 128, 128)
        .HighLight = flexHighlightNever
        
        .Rows = 1
        .FixedRows = 1
        .Cols = OpenOrdersCol(eGDOpenOrdersCol_NumCols)
        .FixedCols = 0
        .FrozenCols = OpenOrdersCol(eGDOpenOrdersCol_OrderText) + 1
        
        .TextMatrix(0, OpenOrdersCol(eGDOpenOrdersCol_OrderID)) = "Order ID"
        .TextMatrix(0, OpenOrdersCol(eGDOpenOrdersCol_SymbolID)) = "Symbol ID"
        .TextMatrix(0, OpenOrdersCol(eGDOpenOrdersCol_Remove)) = "Remove"
        
        .TextMatrix(0, OpenOrdersCol(eGDOpenOrdersCol_Date)) = "Date"
        .TextMatrix(0, OpenOrdersCol(eGDOpenOrdersCol_Symbol)) = "Symbol"
        .TextMatrix(0, OpenOrdersCol(eGDOpenOrdersCol_OrderText)) = "Working Order"
        .Cell(flexcpFontBold, 0, OpenOrdersCol(eGDOpenOrdersCol_OrderText)) = True
        .TextMatrix(0, OpenOrdersCol(eGDOpenOrdersCol_Cancel)) = "X"
        .TextMatrix(0, OpenOrdersCol(eGDOpenOrdersCol_NumFilled)) = "Filled"
        .TextMatrix(0, OpenOrdersCol(eGDOpenOrdersCol_CurrentPrice)) = "Current Price"
        .TextMatrix(0, OpenOrdersCol(eGDOpenOrdersCol_CurrentBid)) = "Bid"
        .TextMatrix(0, OpenOrdersCol(eGDOpenOrdersCol_CurrentAsk)) = "Ask"
        .TextMatrix(0, OpenOrdersCol(eGDOpenOrdersCol_Status)) = "Status"
        .TextMatrix(0, OpenOrdersCol(eGDOpenOrdersCol_AutoTradeItem)) = "Auto Trade Item"
        .TextMatrix(0, OpenOrdersCol(eGDOpenOrdersCol_BrokerID)) = "Broker ID"
        .TextMatrix(0, OpenOrdersCol(eGDOpenOrdersCol_LinkStatus)) = "Link"
        
        .ColHidden(OpenOrdersCol(eGDOpenOrdersCol_OrderID)) = True
        .ColHidden(OpenOrdersCol(eGDOpenOrdersCol_SymbolID)) = True
        .ColHidden(OpenOrdersCol(eGDOpenOrdersCol_Remove)) = True
        
        .ColFormat(OpenOrdersCol(eGDOpenOrdersCol_Date)) = DateAndTime("Format")
        
        .ColDataType(OpenOrdersCol(eGDOpenOrdersCol_Date)) = flexDTDate
        .ColDataType(OpenOrdersCol(eGDOpenOrdersCol_Remove)) = flexDTBoolean
        
        .ColAlignment(OpenOrdersCol(eGDOpenOrdersCol_Date)) = flexAlignCenterTop
        .ColAlignment(OpenOrdersCol(eGDOpenOrdersCol_CurrentPrice)) = flexAlignRightTop
        .ColAlignment(OpenOrdersCol(eGDOpenOrdersCol_CurrentBid)) = flexAlignRightTop
        .ColAlignment(OpenOrdersCol(eGDOpenOrdersCol_CurrentAsk)) = flexAlignRightTop
        .ColAlignment(OpenOrdersCol(eGDOpenOrdersCol_BrokerID)) = flexAlignLeftTop
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignLeftTop
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.InitOpenOrdersGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitClosedOrdersGrid
'' Description: Initialize the closed orders grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitClosedOrdersGrid()
On Error GoTo ErrSection:

    Dim nRedraw As Long                 ' Current state of the grid's redraw
    
    With fgClosedOrders
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        .AllowBigSelection = False
        .AllowSelection = False
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .BackColorBkg = g.Styler.GetColor(eGrid_Background) 'RH override vbApplicationWorkspacevbButtonFace
        .Editable = flexEDNone
        .ExplorerBar = flexExSortShowAndMove
        .ExtendLastCol = False
        .ScrollBars = flexScrollBarBoth
        .ScrollTrack = True
        .SelectionMode = flexSelectionListBox
        .SheetBorder = RGB(128, 128, 128)
        .HighLight = flexHighlightNever
        
        .Rows = 1
        .FixedRows = 1
        .Cols = ClosedOrdersCol(eGDClosedOrdersCol_NumCols)
        .FixedCols = 0
        .FrozenCols = ClosedOrdersCol(eGDClosedOrdersCol_OrderText) + 1
        
        .TextMatrix(0, ClosedOrdersCol(eGDClosedOrdersCol_OrderID)) = "Order ID"
        .TextMatrix(0, ClosedOrdersCol(eGDClosedOrdersCol_SymbolID)) = "Symbol ID"
        .TextMatrix(0, ClosedOrdersCol(eGDClosedOrdersCol_Remove)) = "Remove"
        
        .TextMatrix(0, ClosedOrdersCol(eGDClosedOrdersCol_Date)) = "Date"
        .TextMatrix(0, ClosedOrdersCol(eGDClosedOrdersCol_Symbol)) = "Symbol"
        .TextMatrix(0, ClosedOrdersCol(eGDClosedOrdersCol_OrderText)) = "Closed Order"
        .Cell(flexcpFontBold, 0, ClosedOrdersCol(eGDClosedOrdersCol_OrderText)) = True
        .TextMatrix(0, ClosedOrdersCol(eGDClosedOrdersCol_NumFilled)) = "Filled"
        .TextMatrix(0, ClosedOrdersCol(eGDClosedOrdersCol_Status)) = "Status"
        .TextMatrix(0, ClosedOrdersCol(eGDClosedOrdersCol_AutoTradeItem)) = "Auto Trade Item"
        .TextMatrix(0, ClosedOrdersCol(eGDClosedOrdersCol_BrokerID)) = "Broker ID"
        .TextMatrix(0, ClosedOrdersCol(eGDClosedOrdersCol_LinkStatus)) = "Link"
        
        .ColHidden(ClosedOrdersCol(eGDClosedOrdersCol_OrderID)) = True
        .ColHidden(ClosedOrdersCol(eGDClosedOrdersCol_SymbolID)) = True
        .ColHidden(ClosedOrdersCol(eGDClosedOrdersCol_Remove)) = True
        
        .ColFormat(ClosedOrdersCol(eGDClosedOrdersCol_Date)) = DateAndTime("Format")
        
        .ColDataType(ClosedOrdersCol(eGDClosedOrdersCol_Date)) = flexDTDate
        .ColDataType(ClosedOrdersCol(eGDClosedOrdersCol_Remove)) = flexDTBoolean
        
        .ColAlignment(ClosedOrdersCol(eGDClosedOrdersCol_Date)) = flexAlignCenterTop
        .ColAlignment(ClosedOrdersCol(eGDClosedOrdersCol_BrokerID)) = flexAlignLeftTop
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignLeftTop
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.InitClosedOrdersGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitTransactionGrid
'' Description: Initialize the transactions grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitTransactionsGrid()
On Error GoTo ErrSection:
    
    Dim nRedraw As Long                 ' Current state of the grid's redraw

    With fgTransactions
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        .AllowBigSelection = False
        .AllowSelection = True ' = False
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .Editable = flexEDNone
        .ExplorerBar = flexExNone ' = flexExSortShow ' = flexExNone
        .ExtendLastCol = True
        .ScrollTrack = True
        .SelectionMode = flexSelectionListBox
        .SheetBorder = RGB(128, 128, 128)
        
        .Rows = 1
        .FixedRows = 1
        .Cols = TransactionCol(eGDTransactionCol_NumCols)
        .FixedCols = 0
        
        .MergeCells = flexMergeFree
        
        .TextMatrix(0, TransactionCol(eGDTransactionCol_Action)) = "Action"
        .TextMatrix(0, TransactionCol(eGDTransactionCol_Quantity)) = "Quantity"
        .TextMatrix(0, TransactionCol(eGDTransactionCol_Symbol)) = "Symbol"
        .TextMatrix(0, TransactionCol(eGDTransactionCol_Price)) = "Price"
        .TextMatrix(0, TransactionCol(eGDTransactionCol_Date)) = "Date"
        .TextMatrix(0, TransactionCol(eGDTransactionCol_Fees)) = "Fees"
        .TextMatrix(0, TransactionCol(eGDTransactionCol_BrokerID)) = "Broker ID"
        .TextMatrix(0, TransactionCol(eGDTransactionCol_ClosedProfit)) = "Closed Profit"
        .TextMatrix(0, TransactionCol(eGDTransactionCol_SessionQuantity)) = "Session Qty"
        .TextMatrix(0, TransactionCol(eGDTransactionCol_SessionProfit)) = "Session Profit"
        .TextMatrix(0, TransactionCol(eGDTransactionCol_Balance)) = "Balance"
        .TextMatrix(0, TransactionCol(eGDTransactionCol_Position)) = "Position"
        .TextMatrix(0, TransactionCol(eGDTransactionCol_OrderID)) = "Order ID"
        .TextMatrix(0, TransactionCol(eGDTransactionCol_PositionID)) = "Trade ID"
        .TextMatrix(0, TransactionCol(eGDTransactionCol_FillDate)) = "Fill Date"
        .TextMatrix(0, TransactionCol(eGDTransactionCol_Remove)) = "Remove"
        
        .TextMatrix(0, TransactionCol(eGDTransactionCol_JournalID)) = "JournalID"
        .TextMatrix(0, TransactionCol(eGDTransactionCol_EmotionNumber)) = "Emotion"
        .TextMatrix(0, TransactionCol(eGDTransactionCol_Feelings)) = "Feelings"
        .TextMatrix(0, TransactionCol(eGDTransactionCol_Reasons)) = "Reasons"
        .TextMatrix(0, TransactionCol(eGDTransactionCol_Thoughts)) = "Thoughts"
        .TextMatrix(0, TransactionCol(eGDTransactionCol_Note)) = "Note"
        
        .ColDataType(TransactionCol(eGDTransactionCol_Date)) = flexDTDate
        .ColDataType(TransactionCol(eGDTransactionCol_Remove)) = flexDTBoolean
        
        .ColFormat(TransactionCol(eGDTransactionCol_Date)) = DateFormat("FORMAT", MM_DD_YYYY, HH_MM_SS)
        .ColFormat(TransactionCol(eGDTransactionCol_ClosedProfit)) = "$#,##0.00"
        .ColFormat(TransactionCol(eGDTransactionCol_SessionProfit)) = "$#,##0.00"
        .ColFormat(TransactionCol(eGDTransactionCol_Fees)) = "$#,##0.00"
        .ColFormat(TransactionCol(eGDTransactionCol_Balance)) = "$#,##0.00"
        
        .ColAlignment(TransactionCol(eGDTransactionCol_Date)) = flexAlignCenterTop
        .ColAlignment(TransactionCol(eGDTransactionCol_Price)) = flexAlignRightTop
        .ColAlignment(TransactionCol(eGDTransactionCol_BrokerID)) = flexAlignLeftTop
        
        .ColHidden(TransactionCol(eGDTransactionCol_FillID)) = True
        .ColHidden(TransactionCol(eGDTransactionCol_OrderID)) = True
        .ColHidden(TransactionCol(eGDTransactionCol_PositionID)) = True
        .ColHidden(TransactionCol(eGDTransactionCol_AdjustmentID)) = True
        .ColHidden(TransactionCol(eGDTransactionCol_FillDate)) = True
        .ColHidden(TransactionCol(eGDTransactionCol_Remove)) = True
        
        .ColHidden(TransactionCol(eGDTransactionCol_JournalID)) = True
        
        .ColHidden(TransactionCol(eGDTransactionCol_SessionQuantity)) = True
        .ColHidden(TransactionCol(eGDTransactionCol_SessionProfit)) = True
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignLeftTop
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.InitTransactionsGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitPositionsGrid
'' Description: Initialize the Positions Grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitPositionsGrid()
On Error GoTo ErrSection:

    Dim nRedraw As Long                 ' Current state of the grid's redraw

    With fgTrades
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        .AllowBigSelection = False
        .AllowSelection = False
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .Editable = flexEDKbdMouse
        .ExplorerBar = flexExNone '= flexExSortShow
        .ExtendLastCol = True
        .ScrollTrack = True
        .SelectionMode = flexSelectionFree
        .SheetBorder = RGB(128, 128, 128)
        
        .Rows = 1
        .FixedRows = 1
        .Cols = TradeCol(eGDTradeCol_NumCols)
        .FixedCols = 0
        
        .TextMatrix(0, TradeCol(eGDTradeCol_AcctPosID)) = "Acct Pos ID"
        .TextMatrix(0, TradeCol(eGDTradeCol_Sequence)) = "Sequence"
        .TextMatrix(0, TradeCol(eGDTradeCol_Symbol)) = "Symbol"
        .TextMatrix(0, TradeCol(eGDTradeCol_Position)) = "Pos"
        .TextMatrix(0, TradeCol(eGDTradeCol_Quantity)) = "Quantity"
        .TextMatrix(0, TradeCol(eGDTradeCol_EntryDate)) = "Entry Date"
        .TextMatrix(0, TradeCol(eGDTradeCol_EntryPrice)) = "Entry Price"
        .TextMatrix(0, TradeCol(eGDTradeCol_EntryRule)) = "Rule"
        .TextMatrix(0, TradeCol(eGDTradeCol_ExitDate)) = "Exit Date"
        .TextMatrix(0, TradeCol(eGDTradeCol_ExitPrice)) = "Exit Price"
        .TextMatrix(0, TradeCol(eGDTradeCol_ExitRule)) = "Rule"
        .TextMatrix(0, TradeCol(eGDTradeCol_Profit)) = "Profit/Loss"
        .TextMatrix(0, TradeCol(eGDTradeCol_Commission)) = "Fees"
        .TextMatrix(0, TradeCol(eGDTradeCol_Balance)) = "Acct Balance"
        .TextMatrix(0, TradeCol(eGDTradeCol_Flag)) = "Flag"
        .TextMatrix(0, TradeCol(eGDTradeCol_Category)) = "Category"
        .TextMatrix(0, TradeCol(eGDTradeCol_CategoryID)) = "CategoryID"
        .TextMatrix(0, TradeCol(eGDTradeCol_ClosedExitPrice)) = "Closed Exit Price"
        .TextMatrix(0, TradeCol(eGDTradeCol_ExitQuantity)) = "Exit Quantity"
        .TextMatrix(0, TradeCol(eGDTradeCol_Notes)) = "Notes"
        .TextMatrix(0, TradeCol(eGDTradeCol_Remove)) = "Remove"
        
        .ColHidden(TradeCol(eGDTradeCol_AcctPosID)) = True
        .ColHidden(TradeCol(eGDTradeCol_Sequence)) = True
        .ColHidden(TradeCol(eGDTradeCol_EntryRule)) = False
        .ColHidden(TradeCol(eGDTradeCol_ExitRule)) = False
        .ColHidden(TradeCol(eGDTradeCol_Flag)) = True
        .ColHidden(TradeCol(eGDTradeCol_ClosedExitPrice)) = True
        .ColHidden(TradeCol(eGDTradeCol_ExitQuantity)) = True
        .ColHidden(TradeCol(eGDTradeCol_Notes)) = True
        .ColHidden(TradeCol(eGDTradeCol_Remove)) = True
        .ColHidden(TradeCol(eGDTradeCol_CategoryID)) = True
        
        .ColDataType(TradeCol(eGDTradeCol_EntryDate)) = flexDTDate
        .ColDataType(TradeCol(eGDTradeCol_ExitDate)) = flexDTDate
        .ColDataType(TradeCol(eGDTradeCol_Remove)) = flexDTBoolean
        
        .ColFormat(TradeCol(eGDTradeCol_EntryDate)) = DateFormat("FORMAT", MM_DD_YYYY, HH_MM_SS)
        .ColFormat(TradeCol(eGDTradeCol_ExitDate)) = DateFormat("FORMAT", MM_DD_YYYY, HH_MM_SS)
        .ColFormat(TradeCol(eGDTradeCol_Profit)) = "$#,##0.00"
        .ColFormat(TradeCol(eGDTradeCol_Commission)) = "$#,##0.00"
        .ColFormat(TradeCol(eGDTradeCol_Balance)) = "$#,##0.00"
        
        .ColAlignment(TradeCol(eGDTradeCol_EntryDate)) = flexAlignCenterTop
        .ColAlignment(TradeCol(eGDTradeCol_ExitDate)) = flexAlignCenterTop
        .ColAlignment(TradeCol(eGDTradeCol_EntryPrice)) = flexAlignRightTop
        .ColAlignment(TradeCol(eGDTradeCol_ExitPrice)) = flexAlignRightTop
        .ColAlignment(TradeCol(eGDTradeCol_Profit)) = flexAlignRightTop
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignLeftTop
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.InitPositionsGrid"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitAccountPositionsGrid
'' Description: Initialize the account positions grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitAccountPositionsGrid()
On Error GoTo ErrSection:

    Dim nRedraw As Long                 ' Current state of the grid's redraw

    With fgAccountPositions
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        .AllowBigSelection = False
        .AllowSelection = False
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .Editable = flexEDNone
        .ExplorerBar = flexExNone
        .ExtendLastCol = True
        .OutlineBar = flexOutlineBarSimpleLeaf
        .ScrollTrack = True
        .SelectionMode = flexSelectionListBox
        .SheetBorder = RGB(128, 128, 128)
        
        .Rows = 1
        .FixedRows = 1
        .Cols = AccountPosCol(eGDAccountPositionCol_NumCols)
        .FixedCols = 0
        
        .TextMatrix(0, AccountPosCol(eGDAccountPositionCol_SymbolID)) = "Symbol ID"
        .TextMatrix(0, AccountPosCol(eGDAccountPositionCol_Symbol)) = "Symbol"
        .TextMatrix(0, AccountPosCol(eGDAccountPositionCol_AutoTradeItem)) = "Category"
        .TextMatrix(0, AccountPosCol(eGDAccountPositionCol_Position)) = "Position"
        .TextMatrix(0, AccountPosCol(eGDAccountPositionCol_Quantity)) = "Quantity"
        .TextMatrix(0, AccountPosCol(eGDAccountPositionCol_Flatten)) = "F"
        .TextMatrix(0, AccountPosCol(eGDAccountPositionCol_Reverse)) = "R"
        .TextMatrix(0, AccountPosCol(eGDAccountPositionCol_AvgEntry)) = "Avg Entry"
        .TextMatrix(0, AccountPosCol(eGDAccountPositionCol_CurrentPrice)) = "Current Price"
        .TextMatrix(0, AccountPosCol(eGDAccountPositionCol_OpenProfit)) = "Open Profit"
        .TextMatrix(0, AccountPosCol(eGDAccountPositionCol_OrderStrategy)) = "Exit Strategy"
        .TextMatrix(0, AccountPosCol(eGDAccountPositionCol_LastTraded)) = "Last Traded"
        .TextMatrix(0, AccountPosCol(eGDAccountPositionCol_SessionDate)) = "Session Date"
        .TextMatrix(0, AccountPosCol(eGDAccountPositionCol_SessionQuantity)) = "Session Qty"
        .TextMatrix(0, AccountPosCol(eGDAccountPositionCol_SessionProfit)) = "Session Profit"
        .TextMatrix(0, AccountPosCol(eGDAccountPositionCol_Remove)) = "Remove"
        
        .ColAlignment(AccountPosCol(eGDAccountPositionCol_AvgEntry)) = flexAlignRightTop
        .ColAlignment(AccountPosCol(eGDAccountPositionCol_CurrentPrice)) = flexAlignRightTop
        .ColAlignment(AccountPosCol(eGDAccountPositionCol_LastTraded)) = flexAlignCenterTop
        .ColAlignment(AccountPosCol(eGDAccountPositionCol_SessionDate)) = flexAlignCenterTop
        
        .ColFormat(AccountPosCol(eGDAccountPositionCol_LastTraded)) = DateFormat("FORMAT", MM_DD_YYYY, HH_MM_SS)
        .ColFormat(AccountPosCol(eGDAccountPositionCol_SessionDate)) = DateFormat("FORMAT", MM_DD_YYYY)
        .ColFormat(AccountPosCol(eGDAccountPositionCol_OpenProfit)) = "$#,##0.00"
        .ColFormat(AccountPosCol(eGDAccountPositionCol_SessionProfit)) = "$#,##0.00"
        
        .ColHidden(AccountPosCol(eGDAccountPositionCol_SymbolID)) = True
        .ColHidden(AccountPosCol(eGDAccountPositionCol_Quantity)) = True
        .ColHidden(AccountPosCol(eGDAccountPositionCol_Remove)) = True
        
        .ColDataType(AccountPosCol(eGDAccountPositionCol_Remove)) = flexDTBoolean
        
        .OutlineCol = AccountPosCol(eGDAccountPositionCol_Symbol)
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignLeftTop
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = nRedraw
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.InitAccountPositionsGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadInfoTab
'' Description: Load the account information tab
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadInfoTab(Optional ByVal nAccountType As eTT_AccountType = eTT_AccountType_SimStream)
On Error GoTo ErrSection:

    With m.Account
        If .AccountID <> 0 Then
            txtAccountNumber.Text = .AccountNumber
            txtAccountName.Text = .Name
            gdStartDate.Value = .StartingDate
            txtBroker.Text = .Broker
            txtFutureFees.Text = Format(.Comms, "$#,##0.00")
            txtStockFees.Text = Format(.StockFees, "$#,##0.00")
            SetAccountTypeCombo .AccountType
            
            CheckBoxValue(chkBrokerOCO) = .HoldOcoAtBroker
        
            CheckBoxValue(chkFutures) = GetBit(.SecTypeMask, 1)
            CheckBoxValue(chkStocks) = GetBit(.SecTypeMask, 2)
            CheckBoxValue(chkForex) = GetBit(.SecTypeMask, 3)
            CheckBoxValue(chkFutOpts) = GetBit(.SecTypeMask, 4)
            CheckBoxValue(chkStkOpts) = GetBit(.SecTypeMask, 5)
            
            optFifo.Value = (.FillMatchMode = eTT_FillMatchMode_Fifo)
            optLifo.Value = (.FillMatchMode = eTT_FillMatchMode_Lifo)
        Else
            txtAccountNumber.Text = .AccountNumber
            txtAccountName.Text = ""
            gdStartDate.Value = Now
            txtBroker.Text = ""
            txtFutureFees.Text = ""
            txtStockFees.Text = ""
            SetAccountTypeCombo nAccountType
            CheckBoxValue(chkBrokerOCO) = g.Broker.BrokerAllowsOCO(nAccountType)
            
            If Not g.Broker.IsLiveAccount(nAccountType) Then
                CheckBoxValue(chkFutures) = True
                CheckBoxValue(chkStocks) = True
                CheckBoxValue(chkForex) = True
                CheckBoxValue(chkFutOpts) = True
                CheckBoxValue(chkStkOpts) = True
            Else
                CheckBoxValue(chkFutures) = True
                CheckBoxValue(chkStocks) = False
                CheckBoxValue(chkForex) = False
                CheckBoxValue(chkFutOpts) = False
                CheckBoxValue(chkStkOpts) = False
            End If
            
            optFifo.Value = True
            optLifo.Value = False
        End If
    End With
            
    RefreshAccountTotals
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.LoadInfoTab"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshAccountTotals
'' Description: Refresh the account totals on the info tab
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RefreshAccountTotals()
On Error GoTo ErrSection:

    With m.Account
        ' Don't update this if the user may currently be editing the starting
        ' balance...
        If m.bEditingSB = False Then
            txtStartingBalance.Text = Format(.StartingBalance, "$#,##0.00")
            ColorTextBox txtStartingBalance
        End If
        
        txtTotalAdjustments.Text = Format(.TotalAdjustments, "$#,##0.00")
        ColorTextBox txtTotalAdjustments
        
        txtTotalFees.Text = Format(.TotalFees * -1#, "$#,##0.00")
        ColorTextBox txtTotalFees
        
        txtTotalClosedProfit.Text = Format(.ClosedProfit, "$#,##0.00")
        ColorTextBox txtTotalClosedProfit
        
        txtClosedBalance.Text = Format(.CurrentBalance, "$#,##0.00")
        ColorTextBox txtClosedBalance
        
        txtTotalOpenEquity.Text = Format(.OpenProfit, "$#,##0.00")
        ColorTextBox txtTotalOpenEquity
        
        txtCurrentValue.Text = Format(.CurrentValue, "$#,##0.00")
        ColorTextBox txtCurrentValue
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.RefreshAccountTotals"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ResetOpenOrderRemoveFlags
'' Description: Reset the remove flag on each row of the open orders grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ResetOpenOrderRemoveFlags()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    With fgOpenOrders
        For lIndex = .FixedRows To .Rows - 1
            CheckedCell(fgOpenOrders, lIndex, OpenOrdersCol(eGDOpenOrdersCol_Remove)) = True
        Next lIndex
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.ResetOpenOrderRemoveFlags"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OpenOrderToGrid
'' Description: Refresh the given order to the appropriate line in the grid
'' Inputs:      Order, Row
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub OpenOrderToGrid(Order As cPtOrder, Optional ByVal lRow As Long = -1&)
On Error GoTo ErrSection:

    Dim nRedraw As Long                 ' Current state of the grid's redraw
    Dim Bars As cGdBars                 ' Bars structure to get latest value
    
    With fgOpenOrders
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        If lRow = -1& Then
            .Rows = .Rows + 1
            lRow = .Rows - 1
        End If
        
        .RowData(lRow) = Order

        .TextMatrix(lRow, OpenOrdersCol(eGDOpenOrdersCol_OrderID)) = Str(Order.OrderID)
        .TextMatrix(lRow, OpenOrdersCol(eGDOpenOrdersCol_Date)) = Order.OrderDate
        .TextMatrix(lRow, OpenOrdersCol(eGDOpenOrdersCol_SymbolID)) = Str(Order.SymbolID)
        .TextMatrix(lRow, OpenOrdersCol(eGDOpenOrdersCol_Symbol)) = Order.Symbol
        .TextMatrix(lRow, OpenOrdersCol(eGDOpenOrdersCol_OrderText)) = Order.OrderText(False)
        .TextMatrix(lRow, OpenOrdersCol(eGDOpenOrdersCol_Cancel)) = "X"
        .TextMatrix(lRow, OpenOrdersCol(eGDOpenOrdersCol_NumFilled)) = Str(Order.FillQuantity)
        
        Set Bars = GetTradeBars(Order.SymbolOrSymbolID)
        If Not Bars Is Nothing Then
            If Bars(eBARS_Close, Bars.Size - 1) <> kNullData Then
                .TextMatrix(lRow, OpenOrdersCol(eGDOpenOrdersCol_CurrentPrice)) = Bars.PriceDisplay(Bars(eBARS_Close, Bars.Size - 1))
            Else
                .TextMatrix(lRow, OpenOrdersCol(eGDOpenOrdersCol_CurrentPrice)) = ""
            End If
            If Bars(eBARS_Bid, Bars.Size - 1) <> kNullData Then
                .TextMatrix(lRow, OpenOrdersCol(eGDOpenOrdersCol_CurrentBid)) = Bars.PriceDisplay(Bars(eBARS_Bid, Bars.Size - 1))
            Else
                .TextMatrix(lRow, OpenOrdersCol(eGDOpenOrdersCol_CurrentBid)) = ""
            End If
            If Bars(eBARS_Ask, Bars.Size - 1) <> kNullData Then
                .TextMatrix(lRow, OpenOrdersCol(eGDOpenOrdersCol_CurrentAsk)) = Bars.PriceDisplay(Bars(eBARS_Ask, Bars.Size - 1))
            Else
                .TextMatrix(lRow, OpenOrdersCol(eGDOpenOrdersCol_CurrentAsk)) = ""
            End If
        Else
            .TextMatrix(lRow, OpenOrdersCol(eGDOpenOrdersCol_CurrentPrice)) = ""
            .TextMatrix(lRow, OpenOrdersCol(eGDOpenOrdersCol_CurrentBid)) = ""
            .TextMatrix(lRow, OpenOrdersCol(eGDOpenOrdersCol_CurrentAsk)) = ""
        End If
        
        .TextMatrix(lRow, OpenOrdersCol(eGDOpenOrdersCol_Status)) = OrderStatus(Order.Status)
        .TextMatrix(lRow, OpenOrdersCol(eGDOpenOrdersCol_AutoTradeItem)) = g.TradingItems.NameForID(Order.AutoTradeItemID)
        .TextMatrix(lRow, OpenOrdersCol(eGDOpenOrdersCol_BrokerID)) = Str(Order.BrokerID)
        .TextMatrix(lRow, OpenOrdersCol(eGDOpenOrdersCol_LinkStatus)) = Order.LinkStatus
        
        CheckedCell(fgOpenOrders, lRow, OpenOrdersCol(eGDOpenOrdersCol_Remove)) = False
        
        .Redraw = nRedraw
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.OpenOrderToGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshOpenOrders
'' Description: Refresh the open orders grid from the snapshot orders
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RefreshOpenOrders()
On Error GoTo ErrSection:

    Dim nRedraw As Long                 ' Current state of the grid's redraw
    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim lRow As Long                    ' Row to edit
    
    With fgOpenOrders
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        For lIndex = 1 To m.SnapshotOrders.Count
            If IsOpenOrder(m.SnapshotOrders(lIndex).Status) = True Then
                lRow = -1&
                For lIndex2 = .FixedRows To .Rows - 1
                    If .RowData(lIndex2).OrderID = m.SnapshotOrders(lIndex).OrderID Then
                        lRow = lIndex2
                        Exit For
                    End If
                Next lIndex2
                
                OpenOrderToGrid m.SnapshotOrders(lIndex)
            End If
        Next lIndex
        
        SetBackColors fgOpenOrders
        
        If .Rows > .FixedRows Then
            .Col = OpenOrdersCol(eGDOpenOrdersCol_Date)
            .Sort = flexSortGenericDescending
            
            .Row = .FixedRows
            .RowSel = .FixedRows
        End If
        
        .AutoSize 0, .Cols - 1, False, 75
        ExtendCustomColumn fgOpenOrders
        
        .Redraw = nRedraw
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.RefreshOpenOrders"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveOpenOrders
'' Description: Remove each row of the open orders grid with the remove flag set
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RemoveOpenOrders()
On Error GoTo ErrSection:

    Dim nRedraw As Long                 ' Current state of the grid's redraw
    Dim lIndex As Long                  ' Index into a for loop
    Dim bRemoved As Boolean             ' Did we remove anything?
    
    bRemoved = False
    With fgOpenOrders
        nRedraw = .Redraw
        .Redraw = flexRDNone
    
        For lIndex = .Rows - 1 To .FixedRows Step -1
            If CheckedCell(fgOpenOrders, lIndex, OpenOrdersCol(eGDOpenOrdersCol_Remove)) = True Then
                .RemoveItem lIndex
                bRemoved = True
            End If
        Next lIndex
        
        If bRemoved Then
            SetBackColors fgOpenOrders
            .AutoSize 0, .Cols - 1, False, 75
            ExtendCustomColumn fgOpenOrders
        End If
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.RemoveOpenOrders"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ClosedOrderToGrid
'' Description: Refresh the given order to the appropriate line in the grid
'' Inputs:      Order, Row
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ClosedOrderToGrid(Order As cPtOrder, Optional ByVal lRow As Long = -1&)
On Error GoTo ErrSection:

    Dim nRedraw As Long                 ' Current state of the grid's redraw
    Dim Bars As cGdBars                 ' Bars structure to get latest value
    
    With fgClosedOrders
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        If lRow = -1& Then
            .Rows = .Rows + 1
            lRow = .Rows - 1
        End If
        
        .RowData(lRow) = Order

        .TextMatrix(lRow, ClosedOrdersCol(eGDClosedOrdersCol_OrderID)) = Str(Order.OrderID)
        .TextMatrix(lRow, ClosedOrdersCol(eGDClosedOrdersCol_Date)) = Order.OrderDate
        .TextMatrix(lRow, ClosedOrdersCol(eGDClosedOrdersCol_SymbolID)) = Str(Order.SymbolID)
        .TextMatrix(lRow, ClosedOrdersCol(eGDClosedOrdersCol_Symbol)) = Order.Symbol
        .TextMatrix(lRow, ClosedOrdersCol(eGDClosedOrdersCol_OrderText)) = Order.OrderText(False)
        .TextMatrix(lRow, ClosedOrdersCol(eGDClosedOrdersCol_NumFilled)) = Str(Order.FillQuantity)
        .TextMatrix(lRow, ClosedOrdersCol(eGDClosedOrdersCol_Status)) = OrderStatus(Order.Status)
        .TextMatrix(lRow, ClosedOrdersCol(eGDClosedOrdersCol_AutoTradeItem)) = g.TradingItems.NameForID(Order.AutoTradeItemID)
        .TextMatrix(lRow, ClosedOrdersCol(eGDClosedOrdersCol_BrokerID)) = Str(Order.BrokerID)
        .TextMatrix(lRow, ClosedOrdersCol(eGDClosedOrdersCol_LinkStatus)) = Order.LinkStatus
        
        CheckedCell(fgClosedOrders, lRow, ClosedOrdersCol(eGDClosedOrdersCol_Remove)) = False
        
        .Redraw = nRedraw
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.ClosedOrderToGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadClosedOrdersGrid
'' Description: Load the closed orders grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadClosedOrdersGrid()
On Error GoTo ErrSection:

    Dim nRedraw As Long                 ' Current state of the grid's redraw
    Dim lIndex As Long                  ' Index into a for loop
    
    With fgClosedOrders
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        .Rows = .FixedRows
        For lIndex = 1 To m.HistoricalOrders.Count
            If IsOpenOrder(m.HistoricalOrders(lIndex).Status) = False Then
                ClosedOrderToGrid m.HistoricalOrders(lIndex)
            End If
        Next lIndex
                
        If .Rows > .FixedRows Then
            .Col = ClosedOrdersCol(eGDClosedOrdersCol_Date)
            .Sort = flexSortGenericDescending
            
            .Row = .FixedRows
            .RowSel = .FixedRows
        End If
        
        FilterClosedOrders
        .Redraw = nRedraw
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.LoadClosedOrdersGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ResetClosedOrderRemoveFlags
'' Description: Reset the remove flag on each row of the open orders grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ResetClosedOrderRemoveFlags()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    With fgClosedOrders
        For lIndex = .FixedRows To .Rows - 1
            If .RowData(lIndex).IsSnapshot = True Then
                CheckedCell(fgClosedOrders, lIndex, ClosedOrdersCol(eGDClosedOrdersCol_Remove)) = True
            Else
                CheckedCell(fgClosedOrders, lIndex, ClosedOrdersCol(eGDClosedOrdersCol_Remove)) = False
            End If
        Next lIndex
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.ResetClosedOrderRemoveFlags"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshClosedOrders
'' Description: Refresh the open orders grid from the snapshot orders
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RefreshClosedOrders()
On Error GoTo ErrSection:

    Dim nRedraw As Long                 ' Current state of the grid's redraw
    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim lRow As Long                    ' Row to edit
    
    With fgClosedOrders
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        For lIndex = 1 To m.SnapshotOrders.Count
            If IsOpenOrder(m.SnapshotOrders(lIndex).Status) = False Then
                lRow = -1&
                For lIndex2 = .FixedRows To .Rows - 1
                    If .RowData(lIndex2).OrderID = m.SnapshotOrders(lIndex).OrderID Then
                        lRow = lIndex2
                        Exit For
                    End If
                Next lIndex2
                
                ClosedOrderToGrid m.SnapshotOrders(lIndex), lRow
            End If
        Next lIndex
                
        If .Rows > .FixedRows Then
            .Col = ClosedOrdersCol(eGDClosedOrdersCol_Date)
            .Sort = flexSortGenericDescending
            
            .Row = .FixedRows
            .RowSel = .FixedRows
        End If
        
        FilterClosedOrders
        .Redraw = nRedraw
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.RefreshClosedOrders"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveClosedOrders
'' Description: Remove each row of the open orders grid with the remove flag set
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RemoveClosedOrders()
On Error GoTo ErrSection:

    Dim nRedraw As Long                 ' Current state of the grid's redraw
    Dim lIndex As Long                  ' Index into a for loop
    Dim bRemoved As Boolean             ' Did we remove anything?
    
    bRemoved = False
    With fgClosedOrders
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        For lIndex = .Rows - 1 To .FixedRows Step -1
            If CheckedCell(fgClosedOrders, lIndex, ClosedOrdersCol(eGDClosedOrdersCol_Remove)) = True Then
                .RemoveItem lIndex
                bRemoved = True
            End If
        Next lIndex
        
        If bRemoved Then
            FilterClosedOrders
        End If
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.RemoveClosedOrders"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FilterClosedOrders
'' Description: Filter the orders grid's
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FilterClosedOrders()
On Error GoTo ErrSection:

    Dim nRedraw As Long                 ' Current state of the grid's redraw
    Dim lRow As Long                    ' Index into a for loop
    
    With fgClosedOrders
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        For lRow = .FixedRows To .Rows - 1
            .RowHidden(lRow) = (.RowData(lRow).OrderDate < (Date - (Val(txtNumDays.Text))))
        Next lRow
        
        SetBackColors fgClosedOrders
        
        .AutoSize 0, .Cols - 1, False, 75
        ExtendCustomColumn fgClosedOrders
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.FilterClosedOrders"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillToGrid
'' Description: Update the row in the grid with the given fill
'' Inputs:      Fill, Row
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FillToGrid(Fill As cPtFill, Optional ByVal lRow = -1&)
On Error GoTo ErrSection:

    Dim nRedraw As Long                 ' Current state of the grid's redraw
    Dim Journal As cBrokerMessage       ' Journal object for the order

    With fgTransactions
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        If lRow = -1& Then
            .Rows = .Rows + 1
            lRow = .Rows - 1
        End If
        
        .RowData(lRow) = Fill

        If Fill.Buy = True Then
            .TextMatrix(lRow, TransactionCol(eGDTransactionCol_Action)) = "Buy"
        Else
            .TextMatrix(lRow, TransactionCol(eGDTransactionCol_Action)) = "Sell"
        End If
        .TextMatrix(lRow, TransactionCol(eGDTransactionCol_Quantity)) = Format(Fill.Quantity, "#,##0")
        .TextMatrix(lRow, TransactionCol(eGDTransactionCol_Symbol)) = Fill.Symbol
        .TextMatrix(lRow, TransactionCol(eGDTransactionCol_Price)) = Fill.PriceString
        .TextMatrix(lRow, TransactionCol(eGDTransactionCol_Date)) = CDbl(Fill.FillDate)
        .TextMatrix(lRow, TransactionCol(eGDTransactionCol_Fees)) = Fill.Fees
        .TextMatrix(lRow, TransactionCol(eGDTransactionCol_FillID)) = Str(Fill.FillID)
        .TextMatrix(lRow, TransactionCol(eGDTransactionCol_OrderID)) = Str(Fill.OrderID)
        .TextMatrix(lRow, TransactionCol(eGDTransactionCol_FillDate)) = Str(Fill.FillDate)
        .TextMatrix(lRow, TransactionCol(eGDTransactionCol_BrokerID)) = Fill.BrokerID
        .TextMatrix(lRow, TransactionCol(eGDTransactionCol_ClosedProfit)) = Fill.ClosedProfit
        ColorCell fgTransactions, lRow, TransactionCol(eGDTransactionCol_ClosedProfit)
        .TextMatrix(lRow, TransactionCol(eGDTransactionCol_SessionQuantity)) = Format(Fill.SessionQuantity, "#,##0")
        .TextMatrix(lRow, TransactionCol(eGDTransactionCol_SessionProfit)) = Fill.SessionProfit
        ColorCell fgTransactions, lRow, TransactionCol(eGDTransactionCol_SessionProfit)
        .TextMatrix(lRow, TransactionCol(eGDTransactionCol_Position)) = g.Broker.TextPosition(Fill.CurrentPosition)
        CheckedCell(fgTransactions, lRow, TransactionCol(eGDTransactionCol_Remove)) = False
        
        Set Journal = Nothing
        If m.JournalOrderMap.Exists(Str(Fill.OrderID)) Then
            Set Journal = m.Journals(m.JournalOrderMap(Str(Fill.OrderID)))
        End If
        
        JournalToGrid Journal, lRow
        
        .MergeRow(lRow) = False
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.FillToGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadAdjustments
'' Description: Load the collection of adjustments
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadAdjustments()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim Adj As cAccountAdjustment       ' Account adjustment from the database
    
    m.Adjustments.Clear
    
    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccountAdjustments] " & _
                "WHERE [AccountID]=" & Str(m.Account.AccountID) & ";", dbOpenDynaset)
    Do While Not rs.EOF
        Set Adj = New cAccountAdjustment
        If Adj.Load(rs!AdjustmentID, rs) Then
            m.Adjustments.Add Adj, Str(Adj.AdjustmentID)
        End If
    
        rs.MoveNext
    Loop

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.LoadAdjustments"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AdjustmentToTransactionGrid
'' Description: Update the row in the grid with the given adjustment
'' Inputs:      Adjustment, Row
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AdjustmentToTransactionGrid(Adj As cAccountAdjustment, Optional ByVal lRow As Long = -1&)
On Error GoTo ErrSection:

    Dim nRedraw As Long                 ' Current state of the grid's redraw

    With fgTransactions
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        If lRow = -1& Then
            .Rows = .Rows + 1
            lRow = .Rows - 1
        End If
        
        .RowData(lRow) = Adj
        
        .Cell(flexcpText, lRow, TransactionCol(eGDTransactionCol_Action), lRow, TransactionCol(eGDTransactionCol_Price)) = Adj.Description
        .TextMatrix(lRow, TransactionCol(eGDTransactionCol_Date)) = Adj.AdjustmentTime
        .TextMatrix(lRow, TransactionCol(eGDTransactionCol_Fees)) = Adj.Amount
        ColorCell fgTransactions, lRow, TransactionCol(eGDTransactionCol_Fees)
        .TextMatrix(lRow, TransactionCol(eGDTransactionCol_FillID)) = ""
        .TextMatrix(lRow, TransactionCol(eGDTransactionCol_OrderID)) = ""
        .TextMatrix(lRow, TransactionCol(eGDTransactionCol_PositionID)) = ""
        .TextMatrix(lRow, TransactionCol(eGDTransactionCol_AdjustmentID)) = Str(Adj.AdjustmentID)
        .TextMatrix(lRow, TransactionCol(eGDTransactionCol_FillDate)) = Str(Adj.AdjustmentTime)
        CheckedCell(fgTransactions, lRow, TransactionCol(eGDTransactionCol_Remove)) = False
        
        .MergeRow(lRow) = True
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.AdjustmentToTransactionGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadTransactionsGrid
'' Description: Load the transactions grid
'' Inputs:      Calculate Balance?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadTransactionsGrid(Optional ByVal bCalcBalance As Boolean = True)
On Error GoTo ErrSection:

    Dim lRedraw As Long                 ' Current state of the grid's redraw
    Dim lIndex As Long                  ' Index into a for loop
    Dim rs As Recordset                 ' Recordset into the database
    Dim rs2 As Recordset                ' Recordset into the database
    Dim Fill As New cPtFill             ' Temporary Fill object
    
    With fgTransactions
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        ' Clear out the grid...
        .Rows = .FixedRows
        
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, TransactionCol(eGDTransactionCol_Balance)) = m.Account.StartingBalance
        .TextMatrix(.Rows - 1, TransactionCol(eGDTransactionCol_FillDate)) = CDbl(m.Account.StartingDate)
        
        ' Load up all of the fills for this account...
        For lIndex = 1 To m.HistoricalFills.Count
            FillToGrid m.HistoricalFills(lIndex)
        Next lIndex
        
        ' Load up all of the adjustments for this account...
        For lIndex = 1 To m.Adjustments.Count
            AdjustmentToTransactionGrid m.Adjustments(lIndex)
        Next lIndex
        
        If VisibleRows(fgTransactions) > 1 Then
            .Select .FixedRows + 1, TransactionCol(eGDTransactionCol_BrokerID), .Rows - 1, TransactionCol(eGDTransactionCol_BrokerID)
            .Sort = flexSortGenericAscending
            .Select .FixedRows + 1, TransactionCol(eGDTransactionCol_FillDate), .Rows - 1, TransactionCol(eGDTransactionCol_FillDate)
            .Sort = flexSortGenericAscending
                    
            .Row = .Rows - 1
            .RowSel = .Rows - 1
            .ShowCell .Rows - 1, TransactionCol(eGDTransactionCol_Action)
        End If
        
        If bCalcBalance Then
            CalculateTransactionBalance
            FilterTransactions
        End If
        
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.LoadTransactionsGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FilterTransactions
'' Description: Filter the fills grid on account and date range
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FilterTransactions()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lRedraw As Long                 ' Current state of the grid's redraw
    Dim lAccountID As Long              ' Account ID
    Dim dDate As Double                 ' Date of the fill
    
    With fgTransactions
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        For lIndex = .FixedRows To .Rows - 1
            .RowHidden(lIndex) = False
            
            dDate = .Cell(flexcpValue, lIndex, TransactionCol(eGDTransactionCol_FillDate))
            If (chkDateRange = vbChecked) Then
                If (dDate < gdFillsFromDate.Value) Or (dDate > gdFillsToDate.Value + 0.9999) Then
                    .RowHidden(lIndex) = True
                End If
            End If
        Next lIndex
        
        .ColHidden(TransactionCol(eGDTransactionCol_EmotionNumber)) = (chkShowJournal.Value = vbUnchecked)
        .ColHidden(TransactionCol(eGDTransactionCol_Feelings)) = (chkShowJournal.Value = vbUnchecked)
        .ColHidden(TransactionCol(eGDTransactionCol_Reasons)) = (chkShowJournal.Value = vbUnchecked)
        .ColHidden(TransactionCol(eGDTransactionCol_Thoughts)) = (chkShowJournal.Value = vbUnchecked)
        .ColHidden(TransactionCol(eGDTransactionCol_Note)) = (chkShowJournal.Value = vbUnchecked)
        
        SetBackColors fgTransactions
        
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.FilterTransactions"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CalculateTransactionBalance
'' Description: Calculate the balance column in the fills grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CalculateTransactionBalance()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim dBalance As Double              ' Running balance
    
    With fgTransactions
        .Redraw = flexRDNone
        
        If .Rows = .FixedRows Then .Rows = .Rows + 1
        .TextMatrix(1, TransactionCol(eGDTransactionCol_Balance)) = m.Account.StartingBalance
        
        dBalance = m.Account.StartingBalance
        For lIndex = .FixedRows + 1 To .Rows - 1
            If TypeOf .RowData(lIndex) Is cPtFill Then
                dBalance = dBalance + .RowData(lIndex).ClosedProfit - Abs(.RowData(lIndex).Fees)
            ElseIf TypeOf .RowData(lIndex) Is cAccountAdjustment Then
                dBalance = dBalance + .RowData(lIndex).Amount
            End If
            
            .TextMatrix(lIndex, TransactionCol(eGDTransactionCol_Balance)) = dBalance
            ColorCell fgTransactions, lIndex, TransactionCol(eGDTransactionCol_Balance)
        Next lIndex
        
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.CalculateTransactionBalance"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ResetTransactionRemoveFlags
'' Description: Reset the remove flag on each row of the transactions grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ResetTransactionRemoveFlags()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    With fgTransactions
        For lIndex = .FixedRows To .Rows - 1
            If TypeOf .RowData(lIndex) Is cPtFill Then
                If .RowData(lIndex).IsSnapshot = True Then
                    CheckedCell(fgTransactions, lIndex, TransactionCol(eGDTransactionCol_Remove)) = True
                Else
                    CheckedCell(fgTransactions, lIndex, TransactionCol(eGDTransactionCol_Remove)) = False
                End If
            Else
                CheckedCell(fgTransactions, lIndex, TransactionCol(eGDTransactionCol_Remove)) = False
            End If
        Next lIndex
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.ResetTransactionRemoveFlags"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshTransactions
'' Description: Refresh each of the snapshot fills in the transactions grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RefreshTransactions()
On Error GoTo ErrSection:

    Dim nRedraw As Long                 ' Current state of the grid's redraw
    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim lRow As Long                    ' Row in the grid
    
    With fgTransactions
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        For lIndex = 1 To m.SnapshotFills.Count
            lRow = -1&
            For lIndex2 = .FixedRows To .Rows - 1
                If TypeOf .RowData(lIndex2) Is cPtFill Then
                    If .RowData(lIndex2).FillID = m.SnapshotFills(lIndex).FillID Then
                        lRow = lIndex2
                        Exit For
                    End If
                End If
            Next lIndex2
            
            FillToGrid m.SnapshotFills(lIndex), lRow
        Next lIndex
                
        If VisibleRows(fgTransactions) > 1 Then
            .Select .FixedRows + 1, TransactionCol(eGDTransactionCol_BrokerID), .Rows - 1, TransactionCol(eGDTransactionCol_BrokerID)
            .Sort = flexSortGenericAscending
            .Select .FixedRows + 1, TransactionCol(eGDTransactionCol_FillDate), .Rows - 1, TransactionCol(eGDTransactionCol_FillDate)
            .Sort = flexSortGenericAscending
                    
            .Row = .Rows - 1
            .RowSel = .Rows - 1
            .ShowCell .Rows - 1, TransactionCol(eGDTransactionCol_Action)
        End If
        
        CalculateTransactionBalance
        FilterTransactions
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.RefreshTransactions"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveTransactions
'' Description: Remove each row of the transactions grid with the remove flag set
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RemoveTransactions()
On Error GoTo ErrSection:

    Dim nRedraw As Long                 ' Current state of the grid's redraw
    Dim lIndex As Long                  ' Index into a for loop
    Dim bRemoved As Boolean             ' Did we remove anything?
    
    bRemoved = False
    With fgTransactions
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        For lIndex = .Rows - 1 To .FixedRows Step -1
            If CheckedCell(fgTransactions, lIndex, TransactionCol(eGDTransactionCol_Remove)) = True Then
                .RemoveItem lIndex
                bRemoved = True
            End If
        Next lIndex
        
        If bRemoved Then
            CalculateTransactionBalance
            FilterTransactions
        End If
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.RemoveTransactions"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadTradesGrid
'' Description: Load up the grid with the trades for the current account
'' Inputs:      Calculate Balances?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadTradesGrid(Optional ByVal bCalcBalances As Boolean = True)
On Error GoTo ErrSection:

    Dim nRedraw As Long                 ' Redraw property of the grid
    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim AcctPos As cAccountPosition     ' Account position object
    Dim TradeLines As cTradeLines       ' Trade lines object
    Dim lCount As Long                  ' Item count
    Dim lAutoTradeItemID As Long        ' Automated trading item ID

    With fgTrades
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        .Rows = .FixedRows + 1
        .TextMatrix(.Rows - 1, TradeCol(eGDTradeCol_EntryDate)) = m.Account.StartingDate
        .TextMatrix(.Rows - 1, TradeCol(eGDTradeCol_Balance)) = m.Account.StartingBalance
        .TextMatrix(.Rows - 1, TradeCol(eGDTradeCol_Notes)) = "Account Name: " & m.Account.Name & vbCrLf & _
                    "Account Number: " & m.Account.AccountNumber & vbCrLf & _
                    "Started: " & DateFormat(m.Account.StartingDate)
                    
        For lIndex = 1 To m.Positions.Count
            Set AcctPos = m.Positions(lIndex)
            lAutoTradeItemID = AcctPos.AutoTradeItemID
            
            ' Load the trade lines for the categories, not the total (12/03/2007 DAJ)...
            If lAutoTradeItemID <> -1& Then
                Set TradeLines = AcctPos.TradeLines
                lCount = TradeLines.Count
                
                For lIndex2 = 1 To lCount
                    If TradeLines(lIndex2).IsSnapshot = False Then
                        TradeToGrid TradeLines(lIndex2), lAutoTradeItemID
                    End If
                Next lIndex2
            End If
        Next lIndex
        
        For lIndex = 1 To m.Adjustments.Count
            AdjustmentToTradeGrid m.Adjustments(lIndex)
        Next lIndex
    
        If bCalcBalances Then
            CalculateTradeBalance
        End If
        
        .Redraw = nRedraw
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.LoadTradesGrid"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TradeToGrid
'' Description: Refresh the trades grid with the given trade
'' Inputs:      Trade, Auto Trade Item ID, Row
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TradeToGrid(TradeLine As cTradeLine, ByVal lAtID As Long, Optional ByVal lRow As Long = -1&)
On Error GoTo ErrSection:

    Dim nRedraw As Long                 ' Current state of the grid's redraw
    Dim lColor As Long
    
    If g.nColorTheme = kDarkThemeColor Then
        lColor = RGB(128, 128, 128)
        lblOpenEquity.Caption = "Light gray shading = Open Equity (current value of open positions)"
    Else
        lColor = kOpenEquityColor
        lblOpenEquity.Caption = "Light blue shading = Open Equity (current value of open positions)"
    End If

    With fgTrades
        nRedraw = .Redraw
        .Redraw = flexRDNone
    
        If lRow = -1& Then
            .Rows = .Rows + 1
            lRow = .Rows - 1
        End If
        
        .RowData(lRow) = TradeLine
        
        .TextMatrix(lRow, TradeCol(eGDTradeCol_AcctPosID)) = Str(TradeLine.AccountPositionID)
        .TextMatrix(lRow, TradeCol(eGDTradeCol_Sequence)) = Str(TradeLine.TradeNumber)
        
        .TextMatrix(lRow, TradeCol(eGDTradeCol_Symbol)) = TradeLine.Symbol
        .TextMatrix(lRow, TradeCol(eGDTradeCol_Position)) = TradeLine.DirectionString
        .TextMatrix(lRow, TradeCol(eGDTradeCol_Quantity)) = Str(TradeLine.Quantity)
        .TextMatrix(lRow, TradeCol(eGDTradeCol_EntryDate)) = TradeLine.EntryTime
        .TextMatrix(lRow, TradeCol(eGDTradeCol_EntryPrice)) = TradeLine.EntryPriceString
        .TextMatrix(lRow, TradeCol(eGDTradeCol_EntryRule)) = m.TradeRules.EntryRuleForID(TradeLine.EntryRuleID, True)
        .TextMatrix(lRow, TradeCol(eGDTradeCol_ExitDate)) = TradeLine.ExitTime
        .TextMatrix(lRow, TradeCol(eGDTradeCol_ExitPrice)) = TradeLine.ExitPriceString
        .TextMatrix(lRow, TradeCol(eGDTradeCol_ExitRule)) = m.TradeRules.ExitRuleForID(TradeLine.ExitRuleID, True)
        .TextMatrix(lRow, TradeCol(eGDTradeCol_Profit)) = TradeLine.ClosedProfit
        .TextMatrix(lRow, TradeCol(eGDTradeCol_Flag)) = m.astrFlags(TradeLine.RealSimFlag)

        Select Case lAtID
            Case Is > 0
                .TextMatrix(lRow, TradeCol(eGDTradeCol_Category)) = AutoTradeItemNameForID(lAtID)
            Case 0
                .TextMatrix(lRow, TradeCol(eGDTradeCol_Category)) = "Manual"
            Case -1
                .TextMatrix(lRow, TradeCol(eGDTradeCol_Category)) = "Total"
        End Select
        .TextMatrix(lRow, TradeCol(eGDTradeCol_CategoryID)) = Str(lAtID)
        
        .TextMatrix(lRow, TradeCol(eGDTradeCol_Balance)) = ""
        .TextMatrix(lRow, TradeCol(eGDTradeCol_ClosedExitPrice)) = ""
        .TextMatrix(lRow, TradeCol(eGDTradeCol_ExitQuantity)) = ""
        .TextMatrix(lRow, TradeCol(eGDTradeCol_Notes)) = ""
        .TextMatrix(lRow, TradeCol(eGDTradeCol_Commission)) = TradeLine.Fees
        
        ColorCell fgTrades, lRow, TradeCol(eGDTradeCol_Profit)
        If TradeLine.IsOpen Then
            .Cell(flexcpBackColor, lRow, TradeCol(eGDTradeCol_ExitDate), lRow, TradeCol(eGDTradeCol_Balance)) = lColor
        Else
            .Cell(flexcpBackColor, lRow, TradeCol(eGDTradeCol_ExitDate), lRow, TradeCol(eGDTradeCol_Balance)) = .Cell(flexcpBackColor, lRow, 0)
        End If
        
        CheckedCell(fgTrades, lRow, TradeCol(eGDTradeCol_Remove)) = False
                
        .Redraw = nRedraw
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.TradeToGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AdjustmentToTradeGrid
'' Description: Update the row in the grid with the given adjustment
'' Inputs:      Adjustment, Row
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AdjustmentToTradeGrid(Adj As cAccountAdjustment, Optional ByVal lRow As Long = -1&)
On Error GoTo ErrSection:

    Dim nRedraw As Long                 ' Current state of the grid's redraw

    With fgTrades
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        If lRow = -1& Then
            .Rows = .Rows + 1
            lRow = .Rows - 1
        End If
        
        .RowData(lRow) = Adj
        
        .TextMatrix(lRow, TradeCol(eGDTradeCol_AcctPosID)) = Str(Adj.AdjustmentID)
        .TextMatrix(lRow, TradeCol(eGDTradeCol_Sequence)) = ""
        
        .TextMatrix(lRow, TradeCol(eGDTradeCol_Symbol)) = Adj.Description
        .TextMatrix(lRow, TradeCol(eGDTradeCol_Position)) = ""
        .TextMatrix(lRow, TradeCol(eGDTradeCol_Quantity)) = ""
        .TextMatrix(lRow, TradeCol(eGDTradeCol_EntryDate)) = Adj.AdjustmentTime
        .TextMatrix(lRow, TradeCol(eGDTradeCol_EntryPrice)) = ""
        .TextMatrix(lRow, TradeCol(eGDTradeCol_EntryRule)) = ""
        .TextMatrix(lRow, TradeCol(eGDTradeCol_ExitDate)) = Adj.AdjustmentTime
        .TextMatrix(lRow, TradeCol(eGDTradeCol_ExitPrice)) = ""
        .TextMatrix(lRow, TradeCol(eGDTradeCol_ExitRule)) = ""
        .TextMatrix(lRow, TradeCol(eGDTradeCol_Profit)) = ""
        .TextMatrix(lRow, TradeCol(eGDTradeCol_Commission)) = Adj.Amount
        .TextMatrix(lRow, TradeCol(eGDTradeCol_Flag)) = ""
        ColorCell fgTrades, lRow, TradeCol(eGDTradeCol_Commission)
        .TextMatrix(lRow, TradeCol(eGDTradeCol_Balance)) = ""
        
        .TextMatrix(lRow, TradeCol(eGDTradeCol_ClosedExitPrice)) = ""
        .TextMatrix(lRow, TradeCol(eGDTradeCol_ExitQuantity)) = ""
        .TextMatrix(lRow, TradeCol(eGDTradeCol_Notes)) = Adj.Notes
        
        CheckedCell(fgTrades, lRow, TradeCol(eGDTradeCol_Remove)) = False
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.AdjustmentToTradeGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CalculateTradeBalance
'' Description: Calculate the Balance for the Trades grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CalculateTradeBalance(Optional ByVal bReSort As Boolean = True, Optional ByVal lStartRow = kNullData)
On Error GoTo ErrSection:

    Dim nRedraw As Long                 ' Current state of the grid's redraw
    Dim lIndex As Long                  ' Index into a for loop
    Dim dBalance As Double              ' Cumulative balance
    Dim bHasNonManual As Boolean        ' Is there a non-manual trade line?
    Dim TradeLine As cTradeLine         ' Trade line object
    Dim lColor As Long

    If g.nColorTheme = kDarkThemeColor Then
        lColor = RGB(128, 128, 128)
        lblOpenEquity.Caption = "Light gray shading = Open Equity (current value of open positions)"
    Else
        lColor = kOpenEquityColor
        lblOpenEquity.Caption = "Light blue shading = Open Equity (current value of open positions)"
    End If
    
    With fgTrades
        If .Rows = .FixedRows Then .Rows = .Rows + 1
        If .Cell(flexcpValue, 1, TradeCol(eGDTradeCol_Balance)) <> m.Account.StartingBalance Then
            .TextMatrix(1, TradeCol(eGDTradeCol_Balance)) = m.Account.StartingBalance
        End If
    
        If bReSort Then
            nRedraw = .Redraw
            .Redraw = flexRDNone
            
            ' If we are resorting, we will need to override the start row and do them all...
            lStartRow = kNullData
        
            ' Try to make sure that the open equity rows will sort to the bottom and the
            ' account balance line will sort to the top...
            For lIndex = .FixedRows To .Rows - 1
                If TypeOf .RowData(lIndex) Is cTradeLine Then
                    ' If the trade is open, set the exit date to a large number so that it will sort
                    ' to the bottom -- it will be set back later...
                    If .RowData(lIndex).IsOpen = True Then
                        .TextMatrix(lIndex, TradeCol(eGDTradeCol_ExitDate)) = 99999
                    End If
                ElseIf Not (TypeOf .RowData(lIndex) Is cAccountAdjustment) Then
                    .TextMatrix(lIndex, TradeCol(eGDTradeCol_ExitDate)) = -99999
                End If
            Next lIndex

            ' Sort by entry date and then exit date...
            If .Rows > .FixedRows + 1 Then
                .Select .FixedRows + 1, TradeCol(eGDTradeCol_EntryDate), .Rows - 1, TradeCol(eGDTradeCol_EntryDate)
                .Sort = flexSortGenericAscending
                .Select .FixedRows + 1, TradeCol(eGDTradeCol_ExitDate), .Rows - 1, TradeCol(eGDTradeCol_ExitDate)
                .Sort = flexSortGenericAscending
            End If
            
            ' Reset the exit date on the open equity rows...
            For lIndex = .FixedRows To .Rows - 1
                If TypeOf .RowData(lIndex) Is cTradeLine Then
                    If .RowData(lIndex).IsOpen = True Then
                        If .TextMatrix(lIndex, TradeCol(eGDTradeCol_ExitDate)) = 99999 Then
                            .TextMatrix(lIndex, TradeCol(eGDTradeCol_ExitDate)) = ""
                        End If
                    End If
                ElseIf Not (TypeOf .RowData(lIndex) Is cAccountAdjustment) Then
                    .TextMatrix(lIndex, TradeCol(eGDTradeCol_ExitDate)) = ""
                End If
            Next lIndex
        End If
        
        bHasNonManual = False
        
        ' Accumulate balance...
        If (lStartRow = kNullData) Or (lStartRow <= .FixedRows) Then
            lStartRow = .FixedRows + 1
            dBalance = m.Account.StartingBalance
        Else
            dBalance = .Cell(flexcpValue, lStartRow - 1, TradeCol(eGDTradeCol_Balance))
        End If
    
        For lIndex = lStartRow To .Rows - 1
            If TypeOf .RowData(lIndex) Is cTradeLine Then
                Set TradeLine = .RowData(lIndex)
                
                If TradeLine.IsOpen = False Then
                    dBalance = dBalance + TradeLine.ClosedProfit
                Else
                    dBalance = dBalance + TradeLine.OpenProfit
                End If
                
                If bHasNonManual = False Then
                    If .TextMatrix(lIndex, TradeCol(eGDTradeCol_Category)) <> "Manual" Then
                        bHasNonManual = True
                    End If
                End If
                
                ' 04/19/2010 DAJ: Need to account for the fees right here...
                dBalance = dBalance - TradeLine.Fees
            Else
                dBalance = dBalance + .RowData(lIndex).Amount
            End If
            
            If .Cell(flexcpValue, lIndex, TradeCol(eGDTradeCol_Balance)) <> dBalance Then
                .TextMatrix(lIndex, TradeCol(eGDTradeCol_Balance)) = dBalance
                ColorCell fgTrades, lIndex, TradeCol(eGDTradeCol_Balance)
            End If
        Next lIndex
        
        .ColHidden(TradeCol(eGDTradeCol_Category)) = Not bHasNonManual
        
        If bReSort Then
            If .Rows > .FixedRows Then
                .Select .Rows - 1, 0
                .ShowCell .Rows - 1, TradeCol(eGDTradeCol_Symbol)
            End If
            
            SetBackColors fgTrades
            
            For lIndex = .Rows - 1 To .FixedRows Step -1
                If TypeOf .RowData(lIndex) Is cTradeLine Then
                    If .RowData(lIndex).IsOpen = True Then
                        .Cell(flexcpBackColor, lIndex, TradeCol(eGDTradeCol_ExitDate), lIndex, TradeCol(eGDTradeCol_Balance)) = lColor
                    Else
                        Exit For
                    End If
                End If
            Next lIndex
                    
            ' Resize columns...
            .AutoSize 0, .Cols - 1, False, 75
        
            FilterTrades
            EnableControls
        
            .Redraw = nRedraw
        End If
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.CalculateTradeBalance"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FilterTrades
'' Description: Filter the trades grid based on the users choices
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FilterTrades()
On Error GoTo ErrSection:

    Dim nRedraw As Long                 ' Current state of the grid's redraw
    Dim lIndex As Long                  ' Index into a for loop
    Dim bFiltered As Boolean            ' Do we have a filter in place?
    Dim dDate As Double                 ' Date for the item in the grid
    Dim bHidden As Boolean              ' Was the row hidden?
    Dim lColor As Long

    If g.nColorTheme = kDarkThemeColor Then
        lColor = RGB(128, 128, 128)
        lblOpenEquity.Caption = "Light gray shading = Open Equity (current value of open positions)"
    Else
        lColor = kOpenEquityColor
        lblOpenEquity.Caption = "Light blue shading = Open Equity (current value of open positions)"
    End If
    
    With fgTrades
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        bFiltered = m.TradeFilter.HasFilter(False)
        
        ' Walk through the grid and apply filters if applicable...
        For lIndex = .FixedRows To .Rows - 1
            If lIndex = .FixedRows Then
                ' Hide the original account balance if a filter is in place...
                .RowHidden(lIndex) = bFiltered
            Else
                bHidden = False
                
                If m.TradeFilter.UseDateRange Then
                    If TypeOf .RowData(lIndex) Is cAccountAdjustment Then
                        dDate = .RowData(lIndex).AdjustmentTime
                    ElseIf TypeOf .RowData(lIndex) Is cTradeLine Then
                        dDate = .RowData(lIndex).ExitTime
                    End If
                    
                    bHidden = ((dDate < m.TradeFilter.FromDate) Or (dDate > m.TradeFilter.ToDate + 0.9999))
                End If
                
                If (m.TradeFilter.UseSymbol = True) And (bHidden = False) Then
                    If TypeOf .RowData(lIndex) Is cAccountAdjustment Then
                        bHidden = True
                    ElseIf TypeOf .RowData(lIndex) Is cTradeLine Then
                        bHidden = (.RowData(lIndex).Symbol <> m.TradeFilter.Symbol)
                    End If
                End If
                
                If (m.TradeFilter.Direction <> eGDFilterDirection_All) And (bHidden = False) Then
                    If TypeOf .RowData(lIndex) Is cAccountAdjustment Then
                        bHidden = True
                    ElseIf TypeOf .RowData(lIndex) Is cTradeLine Then
                        If m.TradeFilter.Direction = eGDFilterDirection_Longs Then
                            bHidden = (.RowData(lIndex).Direction = "S")
                        Else
                            bHidden = (.RowData(lIndex).Direction = "L")
                        End If
                    End If
                End If
                
                If (m.TradeFilter.UseEntryRule = True) And (bHidden = False) Then
                    If TypeOf .RowData(lIndex) Is cAccountAdjustment Then
                        bHidden = True
                    ElseIf TypeOf .RowData(lIndex) Is cTradeLine Then
                        bHidden = (.RowData(lIndex).EntryRuleID <> m.TradeFilter.EntryRuleID)
                    End If
                End If
                
                If (m.TradeFilter.UseExitRule = True) And (bHidden = False) Then
                    If TypeOf .RowData(lIndex) Is cAccountAdjustment Then
                        bHidden = True
                    ElseIf TypeOf .RowData(lIndex) Is cTradeLine Then
                        bHidden = (.RowData(lIndex).ExitRuleID <> m.TradeFilter.ExitRuleID)
                    End If
                End If
                
                If (m.TradeFilter.TradeType <> eGDFilterTradeType_All) And (bHidden = False) Then
                    If TypeOf .RowData(lIndex) Is cAccountAdjustment Then
                        bHidden = True
                    ElseIf TypeOf .RowData(lIndex) Is cTradeLine Then
                        bHidden = (.RowData(lIndex).RealSimFlag <> m.TradeFilter.TradeType)
                    End If
                End If
                
                If (m.TradeFilter.UseAutoTrade = True) And (bHidden = False) Then
                    If TypeOf .RowData(lIndex) Is cAccountAdjustment Then
                        bHidden = True
                    ElseIf TypeOf .RowData(lIndex) Is cTradeLine Then
                        bHidden = (.TextMatrix(lIndex, TradeCol(eGDTradeCol_CategoryID)) <> Str(m.TradeFilter.AutoTradeID))
                    End If
                End If
                
                .RowHidden(lIndex) = bHidden
            End If
        Next lIndex
        
        SetBackColors fgTrades
        
        For lIndex = .Rows - 1 To .FixedRows Step -1
            If TypeOf .RowData(lIndex) Is cTradeLine Then
                If .RowData(lIndex).IsOpen = True Then
                    .Cell(flexcpBackColor, lIndex, TradeCol(eGDTradeCol_ExitDate), lIndex, TradeCol(eGDTradeCol_Balance)) = lColor
                Else
                    Exit For
                End If
            End If
        Next lIndex
                
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = nRedraw
    End With

    SetTradesFilterCaption

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.FilterTrades"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ResetTradeRemoveFlags
'' Description: Reset the remove flag on each row of the trades grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ResetTradeRemoveFlags()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    With fgTrades
        For lIndex = .FixedRows To .Rows - 1
            If TypeOf .RowData(lIndex) Is cTradeLine Then
                If .RowData(lIndex).IsSnapshot = True Then
                    CheckedCell(fgTrades, lIndex, TradeCol(eGDTradeCol_Remove)) = True
                Else
                    CheckedCell(fgTrades, lIndex, TradeCol(eGDTradeCol_Remove)) = False
                End If
            Else
                CheckedCell(fgTrades, lIndex, TradeCol(eGDTradeCol_Remove)) = False
            End If
        Next lIndex
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.ResetTradeRemoveFlags"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshTrades
'' Description: Refresh each of the snapshot fills in the Trades grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RefreshTrades()
On Error GoTo ErrSection:

    Dim nRedraw As Long                 ' Current state of the grid's redraw
    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim lIndex3 As Long                 ' Index into a for loop
    Dim lRow As Long                    ' Row in the grid
    
    With fgTrades
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        For lIndex = 1 To m.Positions.Count
            For lIndex2 = 1 To m.Positions(lIndex).TradeLines.Count
                ' Load the trade lines for the categories, not the total (12/03/2007 DAJ)...
                If m.Positions(lIndex).AutoTradeItemID <> -1& Then
                    If m.Positions(lIndex).TradeLines(lIndex2).IsSnapshot Then
                        lRow = -1&
                        For lIndex3 = .FixedRows To .Rows - 1
                            If TypeOf .RowData(lIndex3) Is cTradeLine Then
                                If .RowData(lIndex3).AccountPositionID = m.Positions(lIndex).TradeLines(lIndex2).AccountPositionID Then
                                    If .RowData(lIndex3).TradeNumber = m.Positions(lIndex).TradeLines(lIndex2).TradeNumber Then
                                        lRow = lIndex3
                                        Exit For
                                    End If
                                End If
                            End If
                        Next lIndex3
                        
                        TradeToGrid m.Positions(lIndex).TradeLines(lIndex2), m.Positions(lIndex).AutoTradeItemID, lRow
                    End If
                End If
            Next lIndex2
        Next lIndex
        
        CalculateTradeBalance
        
        .Redraw = nRedraw
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.RefreshTrades"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveTrades
'' Description: Remove each row of the Trades grid with the remove flag set
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RemoveTrades()
On Error GoTo ErrSection:

    Dim nRedraw As Long                 ' Current state of the grid's redraw
    Dim lIndex As Long                  ' Index into a for loop
    Dim bRemoved As Boolean             ' Did we remove anything?
    
    bRemoved = False
    With fgTrades
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        For lIndex = .Rows - 1 To .FixedRows Step -1
            If CheckedCell(fgTrades, lIndex, TradeCol(eGDTradeCol_Remove)) = True Then
                .RemoveItem lIndex
                bRemoved = True
            End If
        Next lIndex
        
        If bRemoved Then
            CalculateTradeBalance
        End If
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.RemoveTrades"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ResetPositionRemoveFlags
'' Description: Reset the remove flag on each row of the account positions grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ResetPositionRemoveFlags()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    With fgAccountPositions
        For lIndex = .FixedRows To .Rows - 1
            CheckedCell(fgAccountPositions, lIndex, AccountPosCol(eGDAccountPositionCol_Remove)) = True
        Next lIndex
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.ResetTransactionRemoveFlags"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshAccountPositions
'' Description: Refresh the account positions from the collection
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RefreshAccountPositions()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim lRow As Long                    ' Row in the grid
    
    With fgAccountPositions
        .Redraw = flexRDNone
        
        For lIndex = 1 To m.Positions.Count
            lRow = -1&
            For lIndex2 = .FixedRows To .Rows - 1
                If .RowData(lIndex2).AccountPositionID = m.Positions(lIndex).AccountPositionID Then
                    lRow = lIndex2
                    Exit For
                End If
            Next lIndex2
            
            AccountPositionToGrid m.Positions(lIndex), lRow
        Next lIndex
        
        FilterPositions
        
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.RefreshAccountPositions"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemovePositions
'' Description: Remove each row of the positions grid with the remove flag set
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RemovePositions()
On Error GoTo ErrSection:

    Dim nRedraw As Long                 ' Current state of the grid's redraw
    Dim lIndex As Long                  ' Index into a for loop
    Dim bRemoved As Boolean             ' Did we remove anything?
    
    bRemoved = False
    With fgAccountPositions
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        For lIndex = .Rows - 1 To .FixedRows Step -1
            If CheckedCell(fgAccountPositions, lIndex, AccountPosCol(eGDAccountPositionCol_Remove)) = True Then
                .RemoveItem lIndex
                bRemoved = True
            End If
        Next lIndex
        
        If bRemoved Then
            FilterPositions
        End If
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.RemoveTransactions"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AccountPositionToGrid
'' Description: Create and/or refrsh a row in the grid for the given account position
'' Inputs:      Account Position, Row
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AccountPositionToGrid(AcctPos As cAccountPosition, Optional ByVal lRow As Long = -1&)
On Error GoTo ErrSection:

    Dim Bars As cGdBars                 ' Bars object
    Dim bMismatch As Boolean            ' Is there currently a position mismatch for the symbol?
    Dim dCurrent As Double              ' Current price for open profit

    With fgAccountPositions
        If lRow = -1& Then
            .Rows = .Rows + 1
            lRow = .Rows - 1
        End If
        
        .RowData(lRow) = AcctPos
        
        bMismatch = (g.Broker.PositionMatch(AcctPos.AccountID, AcctPos.SymbolOrSymbolID) = False)
        
        If bMismatch = False Then
            .TextMatrix(lRow, AccountPosCol(eGDAccountPositionCol_Position)) = AcctPos.CurrentPositionSnapshotString
            .TextMatrix(lRow, AccountPosCol(eGDAccountPositionCol_Quantity)) = Format(Abs(AcctPos.CurrentPositionSnapshot), "#,##0")
        Else
            .TextMatrix(lRow, AccountPosCol(eGDAccountPositionCol_Position)) = "Mismatch"
            .TextMatrix(lRow, AccountPosCol(eGDAccountPositionCol_Quantity)) = ""
        End If
        
        .TextMatrix(lRow, AccountPosCol(eGDAccountPositionCol_Flatten)) = "F"
        .Cell(flexcpFontBold, lRow, AccountPosCol(eGDAccountPositionCol_Flatten)) = True
        .Cell(flexcpBackColor, lRow, AccountPosCol(eGDAccountPositionCol_Flatten)) = vbButtonFace
        .Cell(flexcpAlignment, lRow, AccountPosCol(eGDAccountPositionCol_Flatten)) = flexAlignCenterTop
        If (AcctPos.CurrentPositionSnapshot = 0&) And (bMismatch = False) Then
            .Cell(flexcpForeColor, lRow, AccountPosCol(eGDAccountPositionCol_Flatten)) = vbGrayText
        Else
            .Cell(flexcpForeColor, lRow, AccountPosCol(eGDAccountPositionCol_Flatten)) = vbButtonText
        End If
        
        .TextMatrix(lRow, AccountPosCol(eGDAccountPositionCol_Reverse)) = "R"
        .Cell(flexcpFontBold, lRow, AccountPosCol(eGDAccountPositionCol_Reverse)) = True
        .Cell(flexcpBackColor, lRow, AccountPosCol(eGDAccountPositionCol_Reverse)) = vbButtonFace
        .Cell(flexcpAlignment, lRow, AccountPosCol(eGDAccountPositionCol_Reverse)) = flexAlignCenterTop
        If (AcctPos.AutoTradeItemID <> 0&) Or ((AcctPos.CurrentPositionSnapshot = 0&) And (bMismatch = False)) Then
            .Cell(flexcpForeColor, lRow, AccountPosCol(eGDAccountPositionCol_Reverse)) = vbGrayText
        Else
            .Cell(flexcpForeColor, lRow, AccountPosCol(eGDAccountPositionCol_Reverse)) = vbButtonText
        End If
        
        .TextMatrix(lRow, AccountPosCol(eGDAccountPositionCol_Symbol)) = AcctPos.Symbol
        .TextMatrix(lRow, AccountPosCol(eGDAccountPositionCol_SymbolID)) = Str(AcctPos.SymbolID)
        If AcctPos.AutoTradeItemID = -1& Then
            .TextMatrix(lRow, AccountPosCol(eGDAccountPositionCol_AutoTradeItem)) = "Total"
        ElseIf AcctPos.AutoTradeItemID = 0& Then
            .TextMatrix(lRow, AccountPosCol(eGDAccountPositionCol_AutoTradeItem)) = "Manual"
        Else
            .TextMatrix(lRow, AccountPosCol(eGDAccountPositionCol_AutoTradeItem)) = AutoTradeItemNameForID(AcctPos.AutoTradeItemID)
        End If
        If AcctPos.AutoTradeItemID = 0 Then
            .TextMatrix(lRow, AccountPosCol(eGDAccountPositionCol_OrderStrategy)) = g.OrderStrategies.ExitForAccountAndSymbol(AcctPos.AccountID, AcctPos.SymbolOrSymbolID)
        Else
            .TextMatrix(lRow, AccountPosCol(eGDAccountPositionCol_OrderStrategy)) = ""
        End If
        If AcctPos.LastTradedSnapshot > 0 Then
            .TextMatrix(lRow, AccountPosCol(eGDAccountPositionCol_LastTraded)) = AcctPos.LastTradedSnapshot
        Else
            .TextMatrix(lRow, AccountPosCol(eGDAccountPositionCol_LastTraded)) = ""
        End If
        If AcctPos.SessionDateSnapshot > 0 Then
            .TextMatrix(lRow, AccountPosCol(eGDAccountPositionCol_SessionDate)) = AcctPos.SessionDateSnapshot
            .TextMatrix(lRow, AccountPosCol(eGDAccountPositionCol_SessionQuantity)) = Format(AcctPos.SessionQuantitySnapshot, "#,##0")
            .TextMatrix(lRow, AccountPosCol(eGDAccountPositionCol_SessionProfit)) = Format(AcctPos.SessionProfitSnapshot, "$#,##0.00")
            ColorCell fgAccountPositions, lRow, AccountPosCol(eGDAccountPositionCol_SessionProfit)
        Else
            .TextMatrix(lRow, AccountPosCol(eGDAccountPositionCol_SessionDate)) = ""
            .TextMatrix(lRow, AccountPosCol(eGDAccountPositionCol_SessionQuantity)) = ""
            .TextMatrix(lRow, AccountPosCol(eGDAccountPositionCol_SessionProfit)) = ""
        End If
        
        If (AcctPos.CurrentPositionSnapshot <> 0) And (bMismatch = False) Then
            Set Bars = GetTradeBars(AcctPos.SymbolOrSymbolID)
            If Bars Is Nothing Then
                .TextMatrix(lRow, AccountPosCol(eGDAccountPositionCol_AvgEntry)) = ""
                .TextMatrix(lRow, AccountPosCol(eGDAccountPositionCol_CurrentPrice)) = ""
            Else
                If AcctPos.CurrentPositionSnapshot <> 0 Then
                    .TextMatrix(lRow, AccountPosCol(eGDAccountPositionCol_AvgEntry)) = Bars.PriceDisplay(AcctPos.AverageEntrySnapshot)
                Else
                    .TextMatrix(lRow, AccountPosCol(eGDAccountPositionCol_AvgEntry)) = ""
                End If
                
                dCurrent = AcctPos.CurrentPrice(Bars(eBARS_Close, Bars.Size - 1), Bars(eBARS_Bid, Bars.Size - 1), Bars(eBARS_Ask, Bars.Size - 1))
                If dCurrent <> kNullData Then
                    .TextMatrix(lRow, AccountPosCol(eGDAccountPositionCol_CurrentPrice)) = Bars.PriceDisplay(dCurrent)
                    .TextMatrix(lRow, AccountPosCol(eGDAccountPositionCol_OpenProfit)) = Format(AcctPos.OpenProfit(dCurrent), "$#,##0.00")
                    ColorCell fgAccountPositions, lRow, AccountPosCol(eGDAccountPositionCol_OpenProfit)
                Else
                    .TextMatrix(lRow, AccountPosCol(eGDAccountPositionCol_CurrentPrice)) = ""
                    .TextMatrix(lRow, AccountPosCol(eGDAccountPositionCol_OpenProfit)) = ""
                End If
            End If
        Else
            .TextMatrix(lRow, AccountPosCol(eGDAccountPositionCol_CurrentPrice)) = ""
            .TextMatrix(lRow, AccountPosCol(eGDAccountPositionCol_OpenProfit)) = ""
        End If
        
        CheckedCell(fgAccountPositions, lRow, AccountPosCol(eGDAccountPositionCol_Remove)) = False
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.AccountPositionToGrid"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FilterPositions
'' Description: Filter the account positions grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FilterPositions()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim lRedraw As Long                 ' Current state of the grid's redraw
    Dim bAllFlat As Boolean             ' Are all children flat?
    Dim bMismatch As Boolean            ' Is the symbol currently in a position mismatch?
    
    With fgAccountPositions
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        For lIndex = .FixedRows To .Rows - 1
            .IsSubtotal(lIndex) = True
            If .RowData(lIndex).AutoTradeItemID = -1& Then
                .RowOutlineLevel(lIndex) = 1
            Else
                .RowOutlineLevel(lIndex) = 2
                .TextMatrix(lIndex, AccountPosCol(eGDAccountPositionCol_Symbol)) = ""
            End If
        Next lIndex
        
        For lIndex = .FixedRows To .Rows - 1
            If chkShowFlat.Value = vbChecked Then
                .RowHidden(lIndex) = False
            Else
                If .RowOutlineLevel(lIndex) = 1 Then
                    bAllFlat = True
                    If .GetNodeRow(lIndex, flexNTFirstChild) <> -1& Then
                        For lIndex2 = .GetNodeRow(lIndex, flexNTFirstChild) To .GetNodeRow(lIndex, flexNTLastChild)
                            bMismatch = (g.Broker.PositionMatch(.RowData(lIndex2).AccountID, .RowData(lIndex2).SymbolOrSymbolID) = False)
                        
                            If (.RowData(lIndex2).CurrentPositionSnapshot <> 0&) Or (bMismatch = True) Then
                                bAllFlat = False
                                Exit For
                            End If
                        Next lIndex2
                    End If
                    
                    .RowHidden(lIndex) = bAllFlat
                    If .GetNodeRow(lIndex, flexNTFirstChild) <> -1& Then
                        For lIndex2 = .GetNodeRow(lIndex, flexNTFirstChild) To .GetNodeRow(lIndex, flexNTLastChild)
                            .RowHidden(lIndex2) = bAllFlat
                        Next lIndex2
                    End If
                End If
            End If
        Next lIndex
        
        SetBackColors fgAccountPositions
        If .Rows > .FixedRows Then
            .Cell(flexcpBackColor, .FixedRows, AccountPosCol(eGDAccountPositionCol_Flatten), .Rows - 1, AccountPosCol(eGDAccountPositionCol_Reverse)) = vbButtonFace
        End If
        
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.FilterPositions"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ExportTransactions
'' Description: Export the transactions grid to a delimited-string file
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ExportTransactions()
On Error GoTo ErrSection:

    Dim strFileName As String           ' Filename to export the grid to
    Dim astrFile As New cGdArray        ' Array to dump to the file
    Dim lRow As Long                    ' Index into a for loop
    Dim lCol As Long                    ' Index into a for loop
    Dim astrLine As New cGdArray        ' Array to join together for the line
    Dim strDelim As String              ' Delimiter
    
    strFileName = CommonDialogFile(frmMain.CommonDialog1, True, "Text (Tab delimited) (*.TXT)|*.txt|CSV (Comma delimted) (*.CSV)|*.csv", AddSlash(App.Path) & m.Account.AccountNumber & " Transactions")
    If Len(strFileName) > 0 Then
        If UCase(FileExt(strFileName)) = "CSV" Then
            strDelim = ","
        Else
            strDelim = vbTab
        End If
        
        With fgTransactions
            For lRow = 0 To .Rows - 1
                If .RowHidden(lRow) = False Then
                    astrLine.Clear
                    For lCol = 0 To .Cols - 1
                        If .ColHidden(lCol) = False Then
                            astrLine.Add .TextMatrix(lRow, lCol)
                        End If
                    Next lCol
                    
                    astrFile.Add astrLine.JoinFields(strDelim)
                End If
            Next lRow
        End With
        
        astrFile.ToFile strFileName
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.ExportTransactions"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ExportTrades
'' Description: Export the trades grid to a delimited-string file
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ExportTrades()
On Error GoTo ErrSection:

    Dim strFileName As String           ' Filename to export the grid to
    Dim astrFile As New cGdArray        ' Array to dump to the file
    Dim lRow As Long                    ' Index into a for loop
    Dim lCol As Long                    ' Index into a for loop
    Dim astrLine As New cGdArray        ' Array to join together for the line
    Dim strDelim As String              ' Delimiter
    
    strFileName = CommonDialogFile(frmMain.CommonDialog1, True, "Text (Tab delimited) (*.TXT)|*.txt|CSV (Comma delimted) (*.CSV)|*.csv", AddSlash(App.Path) & m.Account.AccountNumber & " Trades")
    If Len(strFileName) > 0 Then
        If UCase(FileExt(strFileName)) = "CSV" Then
            strDelim = ","
        Else
            strDelim = vbTab
        End If
        
        With fgTrades
            For lRow = 0 To .Rows - 1
                If .RowHidden(lRow) = False Then
                    astrLine.Clear
                    For lCol = 0 To .Cols - 1
                        If .ColHidden(lCol) = False Then
                            astrLine.Add .TextMatrix(lRow, lCol)
                        End If
                    Next lCol
                    
                    astrFile.Add astrLine.JoinFields(strDelim)
                End If
            Next lRow
        End With
        
        astrFile.ToFile strFileName
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.ExportTrades"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowReports
'' Description: Show the reports and send the filters over as defaults
'' Inputs:      None
'' Returns:     None
''
'' Fields:      chkDateRange, From Date, To Date, chkAccount, Account ID, chkSymbol,
''              Symbol, Direction, chkEntryRule, Entry Rule ID, chkExitRule
''              Exit Rule ID, RealSimFlag
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ShowReports()
On Error GoTo ErrSection:

    frmTradeReportFilter.ShowMe m.TradeFilter, False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.ShowReports"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadHistoricalFills
'' Description: Load up the historical fills for the account
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadHistoricalFills()
On Error GoTo ErrSection:

    Dim rsFills As Recordset            ' Recordset into the database

    Set rsFills = g.dbPaper.OpenRecordset("SELECT * FROM [tblFills] WHERE [AccountID]=" & Str(m.lAccountID) & ";", dbOpenDynaset)
    m.HistoricalFills.LoadFillsFromRecordset rsFills, False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.LoadHistoricalFills"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadTradeFilterSettings
'' Description: Load up the trade filter settings from the INI file
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadTradeFilterSettings()
On Error GoTo ErrSection:

    Dim strSettings As String           ' Settings from the INI file

    Set m.TradeFilter = New cTradeFilterSettings

    strSettings = GetIniFileProperty("TradesFilter", "", "TTPositions", g.strIniFile)
    If Len(strSettings) > 0 Then
        m.TradeFilter.FromString strSettings, ";", "|"
    Else
        m.TradeFilter.UseDateRange = (GetIniFileProperty("TradesDateFilter", vbUnchecked, "TTPositions", g.strIniFile) = vbChecked)
        m.TradeFilter.FromDate = GetIniFileProperty("TradesFromDate", Date - 30, "TTPositions", g.strIniFile)
        m.TradeFilter.ToDate = GetIniFileProperty("TradesToDate", Date, "TTPositions", g.strIniFile)
        m.TradeFilter.UseSymbol = GetIniFileProperty("TradesSymbolFilter", vbUnchecked, "TTPositions", g.strIniFile)
        
        If GetIniFileProperty("TradesDirFilter", vbUnchecked, "TTPositions", g.strIniFile) = vbUnchecked Then
            m.TradeFilter.Direction = eGDFilterDirection_All
        Else
            If GetIniFileProperty("TradesDirFilterValue", "Longs Only", "TTPositions", g.strIniFile) = "Longs Only" Then
                m.TradeFilter.Direction = eGDFilterDirection_Longs
            Else
                m.TradeFilter.Direction = eGDFilterDirection_Shorts
            End If
        End If
            
        m.TradeFilter.UseEntryRule = (GetIniFileProperty("TradesEntryRuleFilter", vbUnchecked, "TTPositions", g.strIniFile) = vbChecked)
        m.TradeFilter.EntryRuleID = GetIniFileProperty("TradesEntryRuleFilterValue", 0&, "TTPositions", g.strIniFile)
        m.TradeFilter.UseExitRule = (GetIniFileProperty("TradesExitRuleFilter", vbUnchecked, "TTPositions", g.strIniFile) = vbChecked)
        m.TradeFilter.ExitRuleID = GetIniFileProperty("TradesExitRuleFilterValue", 0&, "TTPositions", g.strIniFile)
        If GetIniFileProperty("TradesFlagFilter", vbUnchecked, "TTPositions", g.strIniFile) = vbUnchecked Then
            m.TradeFilter.TradeType = eGDFilterTradeType_All
        Else
            m.TradeFilter.TradeType = GetIniFileProperty("TradesFlagFilterValue", 0, "TTPositions", g.strIniFile)
        End If
        
        m.TradeFilter.UseAutoTrade = False
        m.TradeFilter.AutoTradeID = 0&
    End If

    ' DAJ 04/15/2013: Make sure that the account is always defaulted to the account being loaded
    ' in this instance of the trade tracker...
    m.TradeFilter.UseAccount = True
    m.TradeFilter.AccountIds.Clear
    m.TradeFilter.AccountIds.Add m.lAccountID

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.LoadTradeFilterSettings"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetTradesFilterCaption
'' Description: Set the trades filter caption
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetTradesFilterCaption()
On Error GoTo ErrSection:

    lblTradeFilter.Caption = "Filter: " & m.TradeFilter.Description(False, "; ", m.TradeRules)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.SetTradesFilterCaption"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadJournals
'' Description: Load up journals for the account
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadJournals()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim Journal As cBrokerMessage       ' Journal entry from the database
    
    m.Journals.Clear
    m.JournalOrderMap.Clear
    
    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrderJournal] WHERE [AccountID]=" & Str(m.lAccountID) & ";", dbOpenDynaset)
    Do While Not rs.EOF
        Set Journal = mTradeTracker.RecordsetToBrokerMessage(rs)
        
        AddJournal Journal
        
        rs.MoveNext
    Loop

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.LoadJournals"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddJournal
'' Description: Add or update the given journal in the collections
'' Inputs:      Journal
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddJournal(Journal As cBrokerMessage)
On Error GoTo ErrSection:

    If m.Journals.Exists(Journal("JournalID")) = True Then
        Set m.Journals(Journal("JournalID")) = Journal
    Else
        m.JournalOrderMap.Add m.Journals.Add(Journal, Journal("JournalID")), Journal("OrderID")
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.AddJournal"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveJournal
'' Description: Remove the journal with the given ID from the collections
'' Inputs:      Journal ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RemoveJournal(ByVal lJournalID As Long)
On Error GoTo ErrSection:

    Dim Journal As cBrokerMessage       ' Journal object

    If m.Journals.Exists(Str(lJournalID)) Then
        Set Journal = m.Journals(Str(lJournalID))
        
        If m.JournalOrderMap.Exists(Journal("OrderID")) Then
            m.JournalOrderMap.Remove Journal("OrderID")
        End If
        
        m.JournalOrderMap.Remove Str(lJournalID)
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.RemoveJournal"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    JournalToGrid
'' Description: Update the grid with the given journal
'' Inputs:      Journal, Row
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub JournalToGrid(ByVal Journal As cBrokerMessage, ByVal lRow As Long)
On Error GoTo ErrSection:

    With fgTransactions
        If Journal Is Nothing Then
            .TextMatrix(lRow, TransactionCol(eGDTransactionCol_JournalID)) = ""
            .TextMatrix(lRow, TransactionCol(eGDTransactionCol_EmotionNumber)) = ""
            .TextMatrix(lRow, TransactionCol(eGDTransactionCol_Feelings)) = ""
            .TextMatrix(lRow, TransactionCol(eGDTransactionCol_Reasons)) = ""
            .TextMatrix(lRow, TransactionCol(eGDTransactionCol_Thoughts)) = ""
            .TextMatrix(lRow, TransactionCol(eGDTransactionCol_Note)) = ""
        Else
            .TextMatrix(lRow, TransactionCol(eGDTransactionCol_JournalID)) = Journal("JournalID")
            If Journal("EmotionNumber") = -1 Then
                .TextMatrix(lRow, TransactionCol(eGDTransactionCol_EmotionNumber)) = ""
            Else
                .TextMatrix(lRow, TransactionCol(eGDTransactionCol_EmotionNumber)) = Journal("EmotionNumber")
            End If
            .TextMatrix(lRow, TransactionCol(eGDTransactionCol_Feelings)) = Journal("Feelings")
            .TextMatrix(lRow, TransactionCol(eGDTransactionCol_Reasons)) = Journal("WhyTrade")
            .TextMatrix(lRow, TransactionCol(eGDTransactionCol_Thoughts)) = Journal("Thoughts")
            .TextMatrix(lRow, TransactionCol(eGDTransactionCol_Note)) = Journal("Note")
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.JournalToGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DeleteTransactionRows
'' Description: Allow the user to delete multiple selected rows in the transactions grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DeleteTransactionRows()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim Fill As cPtFill                 ' Fill object to be deleted
    Dim Adjustment As cAccountAdjustment ' Adjustment to be deleted
    Dim bAdjustmentDeleted As Boolean   ' Has an adjustment been deleted?
    Dim bHistoricalFillDeleted As Boolean ' Has a historical fill been deleted?
    Dim bFillWithFeesDeleted As Boolean ' Was a fill with fees deleted?
    Dim alAutoTradeIds As cGdArray      ' Array of automated trading item ids
    Dim lPos As Long                    ' Position of an item in an array
    Dim FillAccount As cPtAccount       ' Account for the given fill
    Dim strKey As String                ' Key into a collection
    Dim astrAccountSymbol As cGdArray   ' Array of unique account/symbol combinations
    Dim astrAccountSymbolAt As cGdArray ' Array of unique account/symbol/at id combinations
    Dim astrParts As cGdArray           ' Parts of a delimited string

    bAdjustmentDeleted = False
    bHistoricalFillDeleted = False
    bFillWithFeesDeleted = False

    If InfBox("Are you sure that you want to delete the selected items?", "?", "Yes|+-No", "Confirmation") = "Y" Then
        InfBox "Please wait while the historical figures for this account get recalculated", , , "Recalculating History", True
        
        Set alAutoTradeIds = New cGdArray
        alAutoTradeIds.Create eGDARRAY_Longs
        Set astrAccountSymbol = New cGdArray
        astrAccountSymbol.Create eGDARRAY_Strings
        Set astrAccountSymbolAt = New cGdArray
        astrAccountSymbolAt.Create eGDARRAY_Strings
        
        ' Delete the fills and fill up the action collections as appropriate...
        For lIndex = 0 To fgTransactions.SelectedRows - 1
            If TypeOf fgTransactions.RowData(fgTransactions.SelectedRow(lIndex)) Is cAccountAdjustment Then
                Set Adjustment = fgTransactions.RowData(fgTransactions.SelectedRow(lIndex))
                If Adjustment.Delete = True Then
                    bAdjustmentDeleted = True
                    g.Broker.BrokerDebug g.Broker.AccountTypeForID(m.lAccountID), "Adjustment has been manually deleted", True
                End If
                
            ElseIf TypeOf fgTransactions.RowData(fgTransactions.SelectedRow(lIndex)) Is cPtFill Then
                Set Fill = fgTransactions.RowData(fgTransactions.SelectedRow(lIndex))
                Fill.Delete "Manual from Trade Tracker"
                
                strKey = Str(Fill.AccountID) & "|" & Fill.Symbol & "|" & Str(Fill.AutoTradingItemID)
                If g.Broker.DateIsSnapshot(Fill.SessionDate, Fill.Broker) Then
                    g.Broker.RemoveFill Fill, False
                    
                    If astrAccountSymbolAt.BinarySearch(strKey & "|", lPos, eGdSort_MatchUsingSearchStringLength) = False Then
                        astrAccountSymbolAt.Add strKey & "|0", lPos
                    End If
                Else
                    bHistoricalFillDeleted = True
                    
                    If Fill.Fees > 0 Then
                        bFillWithFeesDeleted = True
                    End If
                            
                    If astrAccountSymbolAt.BinarySearch(strKey & "|", lPos, eGdSort_MatchUsingSearchStringLength) = True Then
                        astrAccountSymbolAt(lPos) = strKey & "|1"
                    Else
                        astrAccountSymbolAt.Add strKey & "|1", lPos
                    End If
                End If
                
                If Fill.AutoTradingItemID > 0 Then
                    If alAutoTradeIds.BinarySearch(Fill.AutoTradingItemID, lPos) = False Then
                        alAutoTradeIds.Add Fill.AutoTradingItemID, lPos
                    End If
                End If
                
                strKey = Str(Fill.AccountID) & "|" & Fill.Symbol
                If astrAccountSymbol.BinarySearch(strKey, lPos) = False Then
                    astrAccountSymbol.Add strKey, lPos
                End If
            End If
        Next lIndex
        
        ' If any adjustments were deleted then refresh the adjustments...
        If bAdjustmentDeleted = True Then
            RefreshAdjustments
        End If
        
        ' Recalculate fees for this account in case a fill with fees was deleted...
        If bFillWithFeesDeleted = True Then
            m.Account.RecalculateFees
            m.Account.Save
        End If
        
        ' Rebuild the history for any symbol/account/auto trade id combination for which a non-snapshot
        ' fill was deleted...
        For lIndex = 0 To astrAccountSymbolAt.Size - 1
            Set astrParts = New cGdArray
            astrParts.SplitFields astrAccountSymbolAt(lIndex), "|"
            
            g.Broker.RebuildFillSummaryForSymbol CLng(Val(astrParts(0))), astrParts(1), CLng(Val(astrParts(2))), (astrParts(3) = "1")
        Next lIndex
        
        ' Get the position for any symbol/account combination in a simulated account for which a fill was deleted...
        For lIndex = 0 To astrAccountSymbol.Size - 1
            Set astrParts = New cGdArray
            astrParts.SplitFields astrAccountSymbol(lIndex), "|"
            
            If (g.Broker.IsLiveAccount(m.Account.AccountType) = False) Then
                If (g.Broker.ConnectionStatusForBroker(m.Account.AccountType) = eGDConnectionStatus_Connected) Then
                    Select Case m.Account.AccountType
                        Case eTT_AccountType_SimBroker
                            g.SimTradeTs.GetPositionForSymbol CLng(Val(astrParts(0))), astrParts(1)
                        Case eTT_AccountType_SimStream
                            g.SimTradeStream.GetPositionForSymbol CLng(Val(astrParts(0))), astrParts(1)
                    End Select
                End If
            End If
        Next lIndex
        
        ' Refresh the position for any auto trade item for which a fill was deleted...
        For lIndex = 0 To alAutoTradeIds.Size - 1
            g.TradingItems.RefreshPosition alAutoTradeIds(lIndex)
        Next lIndex
        
        ' If any historical fills were deleted then reload them...
        If bHistoricalFillDeleted = True Then
            LoadHistoricalFills
        End If
        
        m.bClearInfBox = True
        
        m.Account.Load m.Account.AccountID
        RefreshAccountTotals
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.DeleteTransactionRows"
    
End Sub

