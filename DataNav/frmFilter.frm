VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{3B008041-905A-11D1-B4AE-444553540000}#1.0#0"; "Vsocx6.ocx"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmFilter 
   Caption         =   "Filter Editor"
   ClientHeight    =   5655
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   9240
   Icon            =   "frmFilter.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   5655
   ScaleWidth      =   9240
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      ScaleHeight     =   495
      ScaleWidth      =   555
      TabIndex        =   11
      Top             =   5160
      Width           =   555
   End
   Begin HexUniControls.ctlUniFrameWL fraMiddleButtons 
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1860
      Width           =   5865
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
      Caption         =   "frmFilter.frx":058A
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmFilter.frx":05CA
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmFilter.frx":05EA
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdNewItem 
         Height          =   360
         Left            =   4680
         TabIndex        =   6
         Top             =   60
         Width           =   1080
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
         Caption         =   "frmFilter.frx":0606
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmFilter.frx":0640
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmFilter.frx":0660
         RightToLeft     =   0   'False
      End
      Begin VB.PictureBox picAdd 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   120
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   120
         Width           =   240
         Begin VB.Image imgAdd 
            Enabled         =   0   'False
            Height          =   240
            Left            =   -15
            Picture         =   "frmFilter.frx":067C
            Stretch         =   -1  'True
            Top             =   0
            Visible         =   0   'False
            Width           =   240
         End
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdEditItem 
         Height          =   360
         Left            =   3600
         TabIndex        =   5
         Top             =   60
         Width           =   960
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
         Caption         =   "frmFilter.frx":0ABE
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmFilter.frx":0AF2
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmFilter.frx":0B12
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdRemove 
         Height          =   360
         Left            =   1680
         TabIndex        =   4
         Top             =   60
         Width           =   1800
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
         Caption         =   "frmFilter.frx":0B2E
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmFilter.frx":0B72
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmFilter.frx":0B92
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdAdd 
         Height          =   360
         Left            =   60
         TabIndex        =   3
         Top             =   60
         Visible         =   0   'False
         Width           =   1500
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
         Caption         =   "frmFilter.frx":0BAE
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmFilter.frx":0BF4
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmFilter.frx":0C14
         RightToLeft     =   0   'False
      End
      Begin VB.PictureBox picRemove 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   1740
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   120
         Visible         =   0   'False
         Width           =   240
         Begin VB.Image imgRemove 
            Enabled         =   0   'False
            Height          =   240
            Left            =   0
            Picture         =   "frmFilter.frx":0C30
            Stretch         =   -1  'True
            Top             =   0
            Visible         =   0   'False
            Width           =   240
         End
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraEdit 
      Height          =   3720
      Left            =   6000
      TabIndex        =   26
      Top             =   180
      Width           =   3075
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
      Caption         =   "frmFilter.frx":1072
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmFilter.frx":10C0
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmFilter.frx":10E0
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniFrameWL fraValues 
         Height          =   3300
         Left            =   120
         TabIndex        =   27
         Top             =   300
         Width           =   2775
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
         Caption         =   "frmFilter.frx":10FC
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmFilter.frx":112E
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmFilter.frx":114E
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniCheckXP chkInvert 
            Height          =   225
            Left            =   1680
            TabIndex        =   15
            Top             =   2700
            Width           =   915
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
            Caption         =   "frmFilter.frx":116A
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmFilter.frx":119C
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmFilter.frx":125E
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optPercentile 
            Height          =   225
            Left            =   1620
            TabIndex        =   10
            Top             =   540
            Width           =   1125
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
            Caption         =   "frmFilter.frx":127A
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmFilter.frx":12B2
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmFilter.frx":1356
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optValue 
            Height          =   225
            Left            =   1620
            TabIndex        =   9
            Top             =   300
            Width           =   1065
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
            Caption         =   "frmFilter.frx":1372
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   -1  'True
            Tip             =   "frmFilter.frx":13A0
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmFilter.frx":1402
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniTextBoxXP txtLow 
            Height          =   300
            Left            =   1620
            TabIndex        =   14
            Top             =   2100
            Width           =   1080
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   8446207
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmFilter.frx":141E
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
            Alignment       =   2
            ScrollBars      =   0
            PasswordChar    =   ""
            TrapTab         =   0   'False
            EnableContextMenu=   -1  'True
            RaiseChangeEvent=   -1  'True
            Tip             =   "frmFilter.frx":144A
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmFilter.frx":146A
         End
         Begin HexUniControls.ctlUniTextBoxXP txtHigh 
            Height          =   300
            Left            =   1620
            TabIndex        =   12
            Top             =   1200
            Width           =   1080
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   8446207
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmFilter.frx":1486
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
            Alignment       =   2
            ScrollBars      =   0
            PasswordChar    =   ""
            TrapTab         =   0   'False
            EnableContextMenu=   -1  'True
            RaiseChangeEvent=   -1  'True
            Tip             =   "frmFilter.frx":14B4
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmFilter.frx":14D4
         End
         Begin VB.PictureBox sldBar 
            BackColor       =   &H00000040&
            ForeColor       =   &H00FFFFFF&
            Height          =   3105
            Left            =   780
            ScaleHeight     =   3045
            ScaleWidth      =   555
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   90
            Width           =   615
            Begin VB.PictureBox sldHigh 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   60
               Left            =   0
               MouseIcon       =   "frmFilter.frx":14F0
               MousePointer    =   99  'Custom
               ScaleHeight     =   60
               ScaleWidth      =   555
               TabIndex        =   32
               TabStop         =   0   'False
               Top             =   420
               Width           =   555
               Begin VB.Line line3D 
                  BorderColor     =   &H00808080&
                  Index           =   2
                  X1              =   0
                  X2              =   555
                  Y1              =   30
                  Y2              =   30
               End
               Begin VB.Line line3D 
                  BorderColor     =   &H00000000&
                  Index           =   3
                  X1              =   0
                  X2              =   555
                  Y1              =   45
                  Y2              =   45
               End
            End
            Begin VB.PictureBox sldLow 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   60
               Left            =   0
               MouseIcon       =   "frmFilter.frx":1642
               MousePointer    =   99  'Custom
               ScaleHeight     =   60
               ScaleWidth      =   555
               TabIndex        =   33
               TabStop         =   0   'False
               Top             =   2160
               Width           =   555
               Begin VB.Line line3D 
                  BorderColor     =   &H00000000&
                  Index           =   1
                  X1              =   0
                  X2              =   555
                  Y1              =   45
                  Y2              =   45
               End
               Begin VB.Line line3D 
                  BorderColor     =   &H00808080&
                  Index           =   0
                  X1              =   0
                  X2              =   555
                  Y1              =   30
                  Y2              =   30
               End
            End
            Begin vsOcx6LibCtl.vsElastic sldMiddle 
               Height          =   495
               Left            =   0
               TabIndex        =   29
               TabStop         =   0   'False
               Top             =   900
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   873
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
               Appearance      =   0
               MousePointer    =   0
               _ConvInfo       =   1
               Version         =   600
               BackColor       =   16777215
               ForeColor       =   0
               FloodColor      =   192
               ForeColorDisabled=   -2147483631
               Caption         =   "True"
               Align           =   0
               Appearance      =   0
               AutoSizeChildren=   0
               BorderWidth     =   0
               ChildSpacing    =   0
               Splitter        =   0   'False
               FloodDirection  =   0
               FloodPercent    =   0
               CaptionPos      =   4
               WordWrap        =   -1  'True
               MaxChildSize    =   0
               MinChildSize    =   0
               TagWidth        =   0
               TagPosition     =   0
               Style           =   0
               TagSplit        =   0   'False
               PicturePos      =   4
               CaptionStyle    =   0
               ResizeFonts     =   0   'False
               GridRows        =   0
               GridCols        =   0
               _GridInfo       =   ""
            End
            Begin vsOcx6LibCtl.vsElastic sldTop 
               Height          =   450
               Left            =   0
               TabIndex        =   30
               TabStop         =   0   'False
               Top             =   0
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   794
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
               Appearance      =   0
               MousePointer    =   0
               _ConvInfo       =   1
               Version         =   600
               BackColor       =   10543359
               ForeColor       =   0
               FloodColor      =   192
               ForeColorDisabled=   -2147483631
               Caption         =   "True"
               Align           =   0
               Appearance      =   0
               AutoSizeChildren=   0
               BorderWidth     =   0
               ChildSpacing    =   0
               Splitter        =   0   'False
               FloodDirection  =   0
               FloodPercent    =   0
               CaptionPos      =   4
               WordWrap        =   -1  'True
               MaxChildSize    =   0
               MinChildSize    =   0
               TagWidth        =   0
               TagPosition     =   0
               Style           =   0
               TagSplit        =   0   'False
               PicturePos      =   4
               CaptionStyle    =   0
               ResizeFonts     =   0   'False
               GridRows        =   0
               GridCols        =   0
               _GridInfo       =   ""
            End
            Begin vsOcx6LibCtl.vsElastic sldBottom 
               Height          =   495
               Left            =   0
               TabIndex        =   31
               TabStop         =   0   'False
               Top             =   1500
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   873
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
               Appearance      =   0
               MousePointer    =   0
               _ConvInfo       =   1
               Version         =   600
               BackColor       =   16777215
               ForeColor       =   0
               FloodColor      =   192
               ForeColorDisabled=   -2147483631
               Caption         =   "True"
               Align           =   0
               Appearance      =   0
               AutoSizeChildren=   0
               BorderWidth     =   0
               ChildSpacing    =   0
               Splitter        =   0   'False
               FloodDirection  =   0
               FloodPercent    =   0
               CaptionPos      =   4
               WordWrap        =   -1  'True
               MaxChildSize    =   0
               MinChildSize    =   0
               TagWidth        =   0
               TagPosition     =   0
               Style           =   0
               TagSplit        =   0   'False
               PicturePos      =   4
               CaptionStyle    =   0
               ResizeFonts     =   0   'False
               GridRows        =   0
               GridCols        =   0
               _GridInfo       =   ""
            End
         End
         Begin HexUniControls.ctlUniLabelXP lblLow 
            Height          =   225
            Left            =   1680
            Top             =   1860
            WhatsThisHelpID =   960
            Width           =   945
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
            Caption         =   "frmFilter.frx":1794
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   2
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmFilter.frx":17D0
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmFilter.frx":17F0
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblOperator 
            Height          =   195
            Left            =   1680
            Top             =   1590
            Width           =   960
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
            Caption         =   "frmFilter.frx":180C
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   2
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmFilter.frx":1832
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmFilter.frx":1852
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblHigh 
            Height          =   195
            Left            =   1680
            Top             =   960
            Width           =   960
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
            Caption         =   "frmFilter.frx":186E
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   2
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmFilter.frx":18AA
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmFilter.frx":18CA
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblSlider 
            Height          =   195
            Index           =   4
            Left            =   0
            Top             =   3000
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
            Caption         =   "frmFilter.frx":18E6
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   1
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmFilter.frx":190C
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmFilter.frx":192C
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblSlider 
            Height          =   195
            Index           =   3
            Left            =   0
            Top             =   780
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
            Caption         =   "frmFilter.frx":1948
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   1
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmFilter.frx":1978
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmFilter.frx":1998
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblSlider 
            Height          =   195
            Index           =   2
            Left            =   0
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
            Caption         =   "frmFilter.frx":19B4
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   1
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmFilter.frx":19DA
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmFilter.frx":19FA
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblSlider 
            Height          =   195
            Index           =   1
            Left            =   0
            Top             =   300
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
            Caption         =   "frmFilter.frx":1A16
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   1
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmFilter.frx":1A3C
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmFilter.frx":1A5C
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblSlider 
            Height          =   225
            Index           =   0
            Left            =   0
            Top             =   0
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
            Caption         =   "frmFilter.frx":1A78
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   1
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmFilter.frx":1A9E
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmFilter.frx":1ABE
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label5 
            Height          =   195
            Left            =   1980
            Top             =   2880
            Width           =   495
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
            Caption         =   "frmFilter.frx":1ADA
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   2
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmFilter.frx":1B04
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmFilter.frx":1B24
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label4 
            Height          =   195
            Left            =   1620
            Top             =   60
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
            Caption         =   "frmFilter.frx":1B40
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmFilter.frx":1B76
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmFilter.frx":1B96
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraBoolean 
         Height          =   3255
         Left            =   120
         TabIndex        =   0
         Top             =   300
         Visible         =   0   'False
         Width           =   2775
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
         Caption         =   "frmFilter.frx":1BB2
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmFilter.frx":1BE6
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmFilter.frx":1C06
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniRadioXP optTrue 
            Height          =   255
            Left            =   900
            TabIndex        =   8
            Top             =   600
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
            Caption         =   "frmFilter.frx":1C22
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   -1  'True
            Tip             =   "frmFilter.frx":1C4C
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmFilter.frx":1C6C
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optFalse 
            Height          =   255
            Left            =   900
            TabIndex        =   16
            Top             =   900
            Width           =   915
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
            Caption         =   "frmFilter.frx":1C88
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmFilter.frx":1CB4
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmFilter.frx":1CD4
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label3 
            Height          =   495
            Left            =   540
            Top             =   60
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
            Caption         =   "frmFilter.frx":1CF0
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmFilter.frx":1D78
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmFilter.frx":1D98
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniLabelXP lblCondName 
         Height          =   255
         Left            =   0
         Top             =   0
         Width           =   3075
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
         Caption         =   "frmFilter.frx":1DB4
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   2
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmFilter.frx":1E14
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmFilter.frx":1E34
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin VB.PictureBox picLeft 
      Height          =   195
      Left            =   3780
      Picture         =   "frmFilter.frx":1E50
      ScaleHeight     =   135
      ScaleWidth      =   75
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1380
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox picRight 
      Height          =   195
      Left            =   4500
      Picture         =   "frmFilter.frx":1E9C
      ScaleHeight     =   135
      ScaleWidth      =   75
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1380
      Visible         =   0   'False
      Width           =   135
   End
   Begin HexUniControls.ctlUniFrameWL fraCount 
      Height          =   450
      Left            =   5220
      TabIndex        =   22
      Top             =   3960
      Visible         =   0   'False
      Width           =   3315
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
      Caption         =   "frmFilter.frx":1EE8
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmFilter.frx":1F08
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmFilter.frx":1F28
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniRadioXP optNumericCount 
         Height          =   220
         Left            =   1440
         TabIndex        =   24
         Top             =   180
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
         Caption         =   "frmFilter.frx":1F44
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   -1  'True
         Tip             =   "frmFilter.frx":1F72
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmFilter.frx":1F92
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optVisualCount 
         Height          =   220
         Left            =   2400
         TabIndex        =   23
         Top             =   180
         Width           =   795
         _ExtentX        =   1402
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
         Caption         =   "frmFilter.frx":1FAE
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmFilter.frx":1FDA
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmFilter.frx":1FFA
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label2 
         Height          =   195
         Left            =   120
         Top             =   180
         Width           =   1275
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
         Caption         =   "frmFilter.frx":2016
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmFilter.frx":2058
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmFilter.frx":2078
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraUpDown 
      Height          =   345
      Left            =   5700
      TabIndex        =   19
      Top             =   4800
      Visible         =   0   'False
      Width           =   3225
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
      Caption         =   "frmFilter.frx":2094
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmFilter.frx":20B4
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmFilter.frx":20D4
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdDown 
         Height          =   300
         Left            =   2100
         TabIndex        =   21
         Top             =   0
         Width           =   1080
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
         Caption         =   "frmFilter.frx":20F0
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmFilter.frx":2124
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmFilter.frx":2144
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdUp 
         Height          =   300
         Left            =   1140
         TabIndex        =   20
         Top             =   0
         Width           =   900
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
         Caption         =   "frmFilter.frx":2160
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmFilter.frx":2190
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmFilter.frx":21B0
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label1 
         Height          =   255
         Left            =   60
         Top             =   60
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
         Caption         =   "frmFilter.frx":21CC
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmFilter.frx":2202
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmFilter.frx":2222
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid fgUsed 
      Height          =   1635
      Left            =   120
      TabIndex        =   1
      Top             =   180
      Width           =   4980
      _cx             =   8784
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483640
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   8421504
      BackColorAlternate=   16777215
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
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
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
      OleDropMode     =   1
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VSFlex7LCtl.VSFlexGrid fgAvailable 
      Height          =   2355
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Visible         =   0   'False
      Width           =   4965
      _cx             =   8758
      _cy             =   4154
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
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
      MousePointer    =   0
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483640
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorAlternate=   16777215
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
      SelectionMode   =   3
      GridLines       =   4
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   1
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   1
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
      OleDropMode     =   1
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin ActiveToolBars.SSActiveToolBars tbToolbar 
      Left            =   5280
      Top             =   5160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131083
      ToolBarsCount   =   1
      ToolsCount      =   8
      DisplayContextMenu=   0   'False
      Tools           =   "frmFilter.frx":223E
      ToolBars        =   "frmFilter.frx":2509
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUp"
      Begin VB.Menu mnuAddCondition 
         Caption         =   "&Add Condition"
      End
      Begin VB.Menu mnuRemoveCondition 
         Caption         =   "&Remove Condition"
      End
      Begin VB.Menu mnuEditItem 
         Caption         =   "&Edit Item"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMoveUp 
         Caption         =   "Move &Up"
      End
      Begin VB.Menu mnuMoveDown 
         Caption         =   "Move &Down"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChangeFont 
         Caption         =   "&Change Font"
      End
   End
End
Attribute VB_Name = "frmFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Grid columns
Private Const kCondCol = 0
Private Const kNameCol = 1
Private Const kDescCol = 2
Private Const kEnglishCol = 3
Private Const kCountCol = 4
Private Const kExtendedCol = kEnglishCol

Private Enum eCatTableFieldNum
    eTbField_MenuID = 0
    eTbField_MenuEnableFlag
    eTbField_CategoryID
    eTbField_CategoryName
    eTbField_CategoryItem
    eTbField_ObjID
    eTbField_ObjCondition
End Enum

Private Type mPrivate
    strDragSource As String
    vDraggedItems As Variant
    vDraggedRows As Variant

    iSliding As Integer
    nPrevColWidth As Long
    
    nEditedCondRow As Long
    bValueChanged As Boolean
    
    strName As String
    strDesc As String
    
    Filter As cFilter
    
    bModal As Boolean
    bOK As Boolean
    
    tbInfo As cGdTable                  'table holding information about criteria & menu objects
    hCategoryMenu As Long               'handle to primary menu object created on form's load
    aMenuHandles As cGdArray            'array of handles to sub menus
    nMenuSelectID As Long               'ID of menu item selected by user
    bCtrlKey As Boolean
    
    ActiveGrid As VSFlexGrid
    
    nBeginSelRow As Long                'for multi-row select (the .SelectedRow returns 0 when selection mode is flexSelectionByRow
    nEndSelRow As Long
End Type

Private m As mPrivate

Public Property Get ID() As String
    ID = m.Filter.ID
End Property

Private Sub chkInvert_Click()
On Error GoTo ErrSection:

    ConditionEdited

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilter.chkInvert.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdEditItem_Click
'' Description: Allow the user to edit either a criteria or a symbol group
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdEditItem_Click()
On Error GoTo ErrExit:

    EditItem

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilter.cmdEditItem.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdNewItem_Click
'' Description: Allow the user to create a new criteria
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdNewItem_Click()
On Error GoTo ErrSection:

    NewItem

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilter.cmdNewItem.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdRemove_Click
'' Description: Moves an item from the "used" side to the "available" side
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdRemove_Click()
On Error GoTo ErrSection:

'JM 06-21-2010: this button originally has a down-arrow picbox (picRemove) & image (imgRemove)
'   the picbox & image controls are still in the form; they have simply been hidden and sent behind
'   the command button as fix for Aardvark 5714

    RemoveItems

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilter.cmdRemove.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdDown_Click
'' Description: If the user clicks on the down button, move an item in the
''              "used" list down one position in the list
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdDown_Click()
On Error GoTo ErrSection:

    Dim strTemp1 As String              ' Temporary string for swapping
    Dim strTemp2 As String              ' Temporary string for swapping

    ' Make sure we can still move down, then complete the swap
    If (fgUsed.Row + 1) < fgUsed.Rows - 1 Then
        strTemp1 = fgUsed.Text
        fgUsed.Row = fgUsed.Row + 1
        strTemp2 = fgUsed.Text
        fgUsed.Text = strTemp1
        fgUsed.Row = fgUsed.Row - 1
        fgUsed.Text = strTemp2
        fgUsed.Row = fgUsed.Row + 1
        fgUsed.ShowCell fgUsed.Row, fgUsed.Col
    
        RecalcConditions
    End If
    
    ' Enable the Save Button(s)...
    EnableToolbar True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilter.cmdDown.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdAdd_Click
'' Description: If the user clicks on the add button, move the selected item
''              in the "available" list over to the "used" list
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdAdd_Click()
On Error GoTo ErrSection:

    AddItems

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilter.cmdAdd.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Save
'' Description: If the user clicks on the Save button, save the filter
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Save(ByVal strButton As String)
On Error GoTo ErrSection:

    Dim strNewName As String            ' Return from the AskBox
    Dim strText As String               ' Prompt to send to the AskBox
    Dim bSaveAs As Boolean              ' Is this a Save As?
    
    ' Make sure that a Silver user cannot create more Filters by using
    ' the Save As button...
    If strButton = "ID_SaveAs" Then
        If gdNumMatchingFiles(AddSlash(App.Path) & "Custom\Cus0*.FIL") >= 5 Then
            If Not HasGold(True, "Creating more custom Filters") Then
                Exit Sub
            End If
        End If
    End If
    
    ' Handle Rename/Save As
    m.strName = Trim(m.strName)
    If Len(m.strName) = 0 Then
        strText = "Save the current Filter as..."
        strNewName = AskBox("h=Save ; i=? ; g=string ; d=" & m.strName & " ; " & strText)
    ElseIf strButton = "ID_SaveAs" Then
        strText = "Save a copy of the current Filter as..."
        strNewName = AskBox("h=Save As ; i=? ; g=string ; d=" & m.strName & " ; " & strText)
        If Trim(UCase(strNewName)) <> UCase(m.strName) Then
            bSaveAs = True
        End If
    ElseIf strButton = "ID_Rename" Then
        strText = "Rename the current Filter as..."
        strNewName = AskBox("h=Rename ; i=? ; g=string ; d=" & m.strName & " ; " & strText)
    Else
        strNewName = m.strName
    End If
    If Len(Trim(strNewName)) = 0 Then
        Exit Sub 'Err.Raise vbObjectError + 1000, , "You must enter in a name for the filter"
    End If
    m.strName = Trim(strNewName)
    SetEditorCaption Me, "Filter", m.strName
    
    If bSaveAs Then Set m.Filter = m.Filter.MakeCopy
    With m.Filter
        If bSaveAs Then .ID = ""
        .Name = m.strName
        .Desc = m.strDesc
        
        ' Save to file
        .ToFile
        
        ' Add back into pool
        .AddToPool True
    End With
    
    ' Refresh symbol grid dropdown and list
    frmSymbolGrid.RefreshGrid
    
    ' Update the quote board if this is the filter on the filter tab...
    frmQuotes.UpdateFilter "FIL:" & m.Filter.ID
    
    m.bOK = True
    EnableToolbar False
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilter.Save", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdUp_Click
'' Description: If the user clicks on the Up button, move an item in the "used"
''              list up one position in the list
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdUp_Click()
On Error GoTo ErrSection:

    Dim strTemp1 As String              ' Temporary string for swapping
    Dim strTemp2 As String              ' Temporary string for swapping

    ' Make sure we aren't already at the top of the list, otherwise swap rows
    If (fgUsed.Row - 1) >= 0 Then
        strTemp1 = fgUsed.Text
        fgUsed.Row = fgUsed.Row - 1
        strTemp2 = fgUsed.Text
        fgUsed.Text = strTemp1
        fgUsed.Row = fgUsed.Row + 1
        fgUsed.Text = strTemp2
        fgUsed.Row = fgUsed.Row - 1
        fgUsed.ShowCell fgUsed.Row, fgUsed.Col
    
        RecalcConditions
    End If
    
    ' Enable the Save Button(s)...
    EnableToolbar True
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilter.cmdUp.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fgUsed_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    If Not m.bCtrlKey Then
        If OldRow <> NewRow Then
            If m.bValueChanged Then
                fgUsed.Row = OldRow         '6180
                'Need to do DoEvents due to buggy behavior: when change
                'txtHigh or txtLow, then click on a different row in
                'this grid, will lose the change unless do this in order
                'to trigger ConditionEdited prior to ConditionToEditor
                DoEvents
            End If
            ConditionToEditor
            EnableButtons
            
            fgUsed.Select NewRow, kNameCol
            
            ' Took this line out because it was causing some window focus issues in some cases
            ' when two filter windows were open (02/18/2009 DAJ)...
            'MoveFocus fgUsed
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilter.fgUsed.AfterRowColChange", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fgUsed_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:
    
    Dim w&, i&
    ' if column being resized is the extended column,
    ' then make the next column bigger (instead of adjusting
    ' the extended column)
    If Col >= kExtendedCol Then
        With fgUsed
            .Redraw = flexRDNone
            w = .ColWidth(Col) - m.nPrevColWidth
            For i = Col + 1 To .Cols - 1
                If Not .ColHidden(i) Then
                    .ColWidth(i) = .ColWidth(i) - w
                    Exit For
                End If
            Next
            m.nPrevColWidth = 0
            ExtendCustomColumn
            .Redraw = flexRDBuffered
        End With
    Else
        ExtendCustomColumn
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilter.fgUsed.AfterUserResize", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgUsed_BeforeMouseDown
'' Description: Initiate the drag procedure
'' Inputs:      Mouse button pressed, Shift status, X Location, Y Location,
''              Whether or not to cancel the operation
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgUsed_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Increment variable
    Dim lMouseRow As Long               ' Current mouse row
    Dim lMouseCol As Long               ' Current mouse column
    Dim strID As String
    Dim obj As Object

    Dim Point As POINTAPI
    Dim r As Rect, i&
    
    Static slRow As Long                ' Last row that was selected
    
    m.nBeginSelRow = -1
    m.nEditedCondRow = -1
    
    ' Capture the mouse row in case this takes a while...
    lMouseRow = fgUsed.MouseRow
    lMouseCol = fgUsed.MouseCol
    If lMouseRow < fgUsed.FixedRows Or lMouseRow >= fgUsed.Rows Then Exit Sub
    
    m.bCtrlKey = False
    With fgUsed
        If lMouseRow >= .Rows - 1 Then
            fgUsed.Select 0, 0
            If IsMenu(m.hCategoryMenu) Then
                Point.X = X / Screen.TwipsPerPixelX
                Point.Y = Y / Screen.TwipsPerPixelY
                ClientToScreen fgUsed.hWnd, Point
                m.nMenuSelectID = TrackPopupMenu(m.hCategoryMenu, TPM_RETURNCMD, Point.X, Point.Y, 0, fgUsed.hWnd, r)
                AddItems
            End If
        ElseIf Button = vbRightButton Then
            fgUsed.Select lMouseRow, 0
            If .SelectedRows <= 1 Then .Row = lMouseRow
            mnuEditItem.Visible = False
            mnuAddCondition.Visible = False
            mnuRemoveCondition.Visible = True
            mnuMoveUp.Visible = True
            mnuMoveDown.Visible = True
            mnuSep2.Visible = True
            Enable mnuMoveUp, (lMouseRow >= .FixedRows + 1 And lMouseRow < .Rows)
            Enable mnuMoveDown, (lMouseRow >= .FixedRows And lMouseRow < .Rows - 1)
            Set m.ActiveGrid = fgUsed
            PopupMenu mnuPopUp
        ElseIf Shift = 1 Then
            ' The Shift key is being pressed
            ' Select everything in between the last row and the current mouse row
            If slRow >= .FixedRows And slRow < .Rows - 2 Then
                If lMouseRow >= .FixedRows And lMouseRow < .Rows - 2 Then
                    .Select slRow, 0, lMouseRow, 0
                    If slRow < lMouseRow Then
                        m.nBeginSelRow = slRow
                        m.nEndSelRow = lMouseRow
                    Else
                        m.nBeginSelRow = lMouseRow
                        m.nEndSelRow = slRow
                    End If
                End If
            End If
            lblCondName.Visible = False
            fraBoolean.Visible = False
            fraValues.Visible = False
        ElseIf Shift = 2 Then
            m.bCtrlKey = True
            lblCondName.Visible = False
            fraBoolean.Visible = False
            fraValues.Visible = False
        ElseIf lMouseRow >= .FixedRows And lMouseRow < .Rows - 1 Then
            ' No key is being pressed (that we care about)
            fgUsed.Select lMouseRow, 0
            slRow = lMouseRow
            ' Use OLEDrag method to start manual OLE drag operation
            ' this will fire the OLEStartDrag event, which we will use
            ' to fill the DataObject with the data we want to drag.
            .OLEDrag
            ' Tell grid control to ignore mouse movements until the
            ' mouse button goes up again
            Cancel = True
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilter.fgUsed.BeforeMouseDown", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fgUsed_BeforeMoveRow(ByVal Row As Long, Position As Long)
On Error GoTo ErrSection:

    If Position < fgUsed.FixedRows Then Position = fgUsed.FixedRows
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilter.fgUsed.BeforeMoveRow", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fgUsed_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:
    
    ' if column being resized is the extended column, save size
    If Col >= kExtendedCol Then
        m.nPrevColWidth = fgUsed.ColWidth(Col)
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilter.fgUsed.BeforeUserResize", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fgUsed_DblClick()
On Error GoTo ErrSection:

    cmdRemove_Click
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilter.fgUsed.DblClick", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub fgUsed_GotFocus()
On Error GoTo ErrSection:

    EnableButtons
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFilter.fgUsed.GotFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub fgUsed_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:
    
    Dim lRow As Long                    ' Current mouse row in the grid
    
    With fgUsed
        lRow = .MouseRow
        If lRow >= .FixedRows And lRow < .Rows Then
            .ToolTipText = TipStr(.TextMatrix(lRow, kDescCol))
        Else
            .ToolTipText = ""
        End If
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilter.fgUsed.MouseMove", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgUsed_OLEDragDrop
'' Description: Drop information into fgUsed
'' Inputs:      Data to drop, Effects, Mouse button pressed, Shift status, X
''              Location, Y Location
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgUsed_OLEDragDrop(Data As VSFlex7LCtl.VSDataObject, Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
On Error GoTo ErrSection:

    Dim iIndexValue As Integer          ' Row to drop information into
    Dim iIndex As Integer               ' Increment variable

    iIndexValue = fgUsed.MouseRow
    
    If iIndexValue < 0 Then
        iIndexValue = fgUsed.Rows
    ElseIf iIndexValue < fgUsed.FixedRows Then
        iIndexValue = fgUsed.FixedRows
    End If

    ' Store the values and remove rows
    If m.strDragSource = "list" Then
        ' Store dragged item information
        For iIndex = 0 To UBound(m.vDraggedRows) - 1
            fgUsed.Row = m.vDraggedRows(iIndex)
            m.vDraggedItems(iIndex) = WholeGridRow(fgUsed)
            If iIndexValue = m.vDraggedRows(iIndex) Then
                GoTo ExitSub
            End If
            
            ' Enable the Save Button(s)...
            EnableToolbar True
        Next iIndex

        ' Remove dragged rows
        For iIndex = 0 To UBound(m.vDraggedItems) - 1
            fgUsed.RemoveItem m.vDraggedRows(iIndex) - iIndex
            If m.vDraggedRows(iIndex) < iIndexValue Then iIndexValue = iIndexValue - 1
        Next iIndex
    End If

    ' If we don't have a valid column add to the end of the control
    If iIndexValue <= -1 Or iIndexValue >= fgUsed.Rows Then
        ' When doing an additem without a row value do them in incrementing order
        For iIndex = 0 To UBound(m.vDraggedItems) - 1
            fgUsed.AddItem m.vDraggedItems(iIndex)
            fgUsed.IsSelected(fgUsed.Rows - 1) = True
            If iIndex = 0 Then fgUsed.Row = fgUsed.Rows - 1
        Next iIndex
        fgUsed.ShowCell fgUsed.Rows - 1, 0
    Else
        fgUsed.Row = iIndexValue
'        fgUsed.IsSelected(fgUsed.Row) = False

        ' When doing an additem with a row value we need to do last item first
        For iIndex = UBound(m.vDraggedItems) - 1 To 0 Step -1
            fgUsed.AddItem m.vDraggedItems(iIndex), fgUsed.Row
'            fgUsed.IsSelected(fgUsed.Row) = True
        Next iIndex
        fgUsed.ShowCell fgUsed.Row, 0
    End If
    
    'If m_Order = eOrderMode_Alphabetical Then
        'fgUsed.Select fgUsed.FixedRows, 0
        'fgUsed.Sort = flexSortStringNoCaseAscending
    'End If

ExitSub:
    picLeft.Visible = False
    picRight.Visible = False
'    fgAvailable.OLEDropMode = flexOLEDropNone

    ConditionToEditor
    MoveFocus fgUsed
    RecalcConditions

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilter.fgUsed.OLEDragDrop", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgUsed_OLEDragOver
'' Description: What to do while dragging over the fgUsed list
'' Inputs:      Data being dragged, Effects, Mouse button pressed, Shift status
''              X Location, Y Location, Status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgUsed_OLEDragOver(Data As VSFlex7LCtl.VSDataObject, Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, State As Integer)
On Error GoTo ErrSection:

    Dim nMouseRow&
    
    nMouseRow = fgUsed.MouseRow

    'If m_Order <> eOrderMode_Ordered Then Exit Sub

    ' Initialize pictures for drag routine
    If picLeft.Visible = False And fgUsed.Rows > fgUsed.FixedRows + 1 Then
        picLeft.Left = fgUsed.Left - picLeft.Width
        picRight.Left = fgUsed.Left + fgUsed.Width

        picLeft.Visible = True
        picRight.Visible = True
    End If

    ' While dragging move the arrow pictures
    If picLeft.Visible = True Then
        If nMouseRow < 0 Then
            nMouseRow = fgUsed.Rows
        ElseIf nMouseRow < fgUsed.FixedRows Then
            nMouseRow = fgUsed.FixedRows
        End If
        picLeft.Top = (nMouseRow - fgUsed.TopRow + fgUsed.FixedRows) * fgUsed.RowHeight(0) + fgUsed.Top - (picLeft.Height / 2)
        picRight.Top = picLeft.Top
    End If

    ' If leaving control hide arrow pictures
    If State = vbLeave Then
        picLeft.Visible = False
        picRight.Visible = False
        Exit Sub
    End If

    ' Scroll up
    If nMouseRow = fgUsed.TopRow And fgUsed.TopRow <> 0 Then
        fgUsed.TopRow = fgUsed.TopRow - 1
        Exit Sub
    End If

    ' Scroll down
    If fgUsed.Rows > 0 Then
        If nMouseRow = -1 And Y > fgUsed.Top + fgUsed.Height - fgUsed.RowHeight(0) Then
            fgUsed.TopRow = fgUsed.TopRow + 1
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilter.fgUsed.OLEDragOver", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgUsed_OLEStartDrag
'' Description: Start drag from fgUsed
'' Inputs:      Data to drag, Effects allowed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgUsed_OLEStartDrag(Data As VSFlex7LCtl.VSDataObject, AllowedEffects As Long)
On Error GoTo ErrSection:

    Dim iIndex As Integer               ' Increment variable

    m.strDragSource = "list"
'    fgAvailable.OLEDropMode = flexOLEDropManual

    ' Retrieve and store which rows were selected
    If fgUsed.SelectedRows > 1 Then
        ReDim m.vDraggedItems(fgUsed.SelectedRows)
        ReDim m.vDraggedRows(fgUsed.SelectedRows)

        ' Store the rows
        For iIndex = 0 To fgUsed.SelectedRows - 1
            m.vDraggedRows(iIndex) = fgUsed.SelectedRow(iIndex)
        Next iIndex
    Else
        ReDim m.vDraggedItems(1)
        ReDim m.vDraggedRows(1)
        m.vDraggedRows(0) = fgUsed.Row
        m.vDraggedItems(0) = WholeGridRow(fgUsed)
    End If

    ' Set contents of data object for manual drag
    ' (Put this code in so that the program would not crash when the mouse is
    ' moved off the form - 11/10/00 DAJ)
'    Dim s$
'    s = fgAvailable.Clip
'    Data.SetData s, vbCFText

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilter.fgUsed.OLEStartDrag", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Activate
'' Description: When the form is activated, reset the toolbars
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Activate()
On Error GoTo ErrSection:

    Screen.MousePointer = vbDefault
    
    ' Reload the grids in case something has changed...
    InitGrids
    
'    If fgUsed.Rows = fgUsed.FixedRows Then
'        If GetActiveWindow = Me.hWnd Then MoveFocus fgAvailable
'    End If
    EnableButtons
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilter.Form.Activate", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If KeyCode = vbKeyF1 Then
        KeyCode = 0
        g.Help.ShowF1Help Me
    Else
        frmMain.DockPro_ShortcutKeyDown KeyCode, Shift, Me.Name
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilter.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: When the form is loaded, initialize variables
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Dim i&, t&, strTemp$
    Dim strFont As String

    g.Styler.StyleForm Me
    
    ' Initialize form variables
    ReDim m.vDraggedRows(0)
    ReDim m.vDraggedItems(0)
    m.strDragSource = " "

    'RH commented out fraMiddleButtons.BorderStyle = vbBSNone
    'RH commented out fraUpDown.BorderStyle = vbBSNone
    'RH commented out fraValues.BorderStyle = vbBSNone
    'RH commented out fraBoolean.BorderStyle = vbBSNone
    
    With tbToolbar
        .Tools("ID_ShowList").Picture = Picture16(ToolbarIcon("ID_SymbolGrid"))
        .Tools("ID_Description").Picture = Picture16(ToolbarIcon("ID_News"))
        .Tools("ID_Print").Picture = Picture16(ToolbarIcon("ID_Print"))
        .Tools("ID_Toolbox").Picture = Picture16(ToolbarIcon("ID_Toolbox"))
        .Tools("ID_Save").Picture = Picture16(ToolbarIcon("kSave"))
        .Tools("ID_SaveAs").Picture = Picture16(ToolbarIcon("kSaveAs"))
        .Tools("ID_Rename").Picture = Picture16(ToolbarIcon("kRename"))
        .Tools("ID_Close").Picture = Picture16(ToolbarIcon("kCancel"))
    End With

'    fgAvailable.ZOrder
    fraUpDown.ZOrder
    picLeft.ZOrder
    picRight.ZOrder

    ' slider bar
    'sldHigh.MousePointer = vbSizeNS
    sldHigh.Height = 4 * Screen.TwipsPerPixelY
    sldLow.MousePointer = sldHigh.MousePointer
    sldLow.MouseIcon = sldHigh.MouseIcon
    sldLow.BackColor = sldHigh.BackColor
    sldLow.Height = sldHigh.Height
    sldLow.ZOrder
    sldHigh.ZOrder
    ' (size slider so an exact multiple of 100 pixels)
    sldBar.Height = 200 * Screen.TwipsPerPixelY + sldLow.Height _
        + (sldBar.Height - sldBar.ScaleHeight)
    With sldBar
        sldHigh.Width = .ScaleWidth
        sldLow.Width = .ScaleWidth
        sldTop.Width = .ScaleWidth
        sldMiddle.Width = .ScaleWidth
        sldBottom.Width = .ScaleWidth
    End With
    ' shadow lines in the slider bars (sldHigh, sldLow)
    For i = 0 To 3
        line3D(i).X2 = sldBar.ScaleWidth
        If i Mod 2 = 0 Then '(gray line)
            line3D(i).Y1 = 2 * Screen.TwipsPerPixelY
        Else '(black line)
            line3D(i).Y1 = 3 * Screen.TwipsPerPixelY
        End If
        line3D(i).Y2 = line3D(i).Y1
    Next

    ' set position of slider bar labels
    For i = 0 To 4
        With lblSlider(i)
            t = sldBar.Top + (sldBar.ScaleHeight - sldLow.Height) * i / 4 _
                    - lblSlider(0).Height / 2 + Screen.TwipsPerPixelY * 3
            .Move lblSlider(0).Left, t, lblSlider(0).Width, lblSlider(0).Height
        End With
    Next

    sldBar.ForeColor = sldTop.BackColor
    txtHigh.BackColor = sldTop.BackColor
    txtLow.BackColor = sldTop.BackColor

    strTemp = GetIniFileProperty("Filter", "", "Placement", g.strIniFile)
    If strTemp <> "" Then
        SetFormPlacement Me, strTemp
    Else
        CenterTheForm Me
    End If
    
    Me.Icon = Picture16(ToolbarIcon("ID_Filters"), , True)
    
    mnuPopUp.Visible = False
    
    ' Set the grid font from the INI file...
    strFont = GetIniFileProperty("FilterUsed", "", "Fonts", g.strIniFile)
    If strFont <> "" Then FontFromString fgUsed.Font, strFont
'    strFont = GetIniFileProperty("FilterAvailable", "", "Fonts", g.strIniFile)
'    If strFont <> "" Then FontFromString fgAvailable.Font, strFont
    
    m.hCategoryMenu = CreatePopupMenu()
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilter.Form.Load", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode <> vbFormCode Then
        If AskToSave Then
            Cancel = True
        ElseIf m.bModal Then
            Cancel = True
            Me.Hide
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilter.Form.QueryUnload", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub Form_Resize()
On Error Resume Next
    
    Dim w&, h&, n&

    ' set minimum form size
    h = fraEdit.Top + fraEdit.Height
    w = fraEdit.Width + fraMiddleButtons.Width
    If LimitFormSize(Me, w, h) Then Exit Sub

    With fraEdit
        .Left = Me.ScaleWidth - .Width
    End With
    With fgUsed
        .Redraw = flexRDNone
        n = m.Filter.Conditions.Size + 3
        If n < 8 Then n = 8
        n = .RowHeight(0) * n
'        h = (Me.ScaleHeight - .Top) * 0.5
        h = (Me.ScaleHeight - .Top) - fraMiddleButtons.Height - 50
        If n > h Then n = h
        .Move .Left, .Top, fraEdit.Left - .Left, h
        ExtendCustomColumn
        .Redraw = flexRDBuffered
    End With
    With fraMiddleButtons
        n = fgUsed.Left + (fgUsed.Width - .Width) \ 2
        If n < 0 Then n = 0
        .Move n, fgUsed.Top + fgUsed.Height
    End With
'    With fgAvailable
'        n = fraMiddleButtons.Top + fraMiddleButtons.Height
'        .Move .Left, n, fgUsed.Width, Me.ScaleHeight - n - .Left
'    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    Dim i&

    Set m.Filter = Nothing
    Set m.ActiveGrid = Nothing
    SetIniFileProperty "Filter", GetFormPlacement(Me), "Placement", g.strIniFile
    SetIniFileProperty "FilterUsed", FontToString(fgUsed.Font), "Fonts", g.strIniFile
'    SetIniFileProperty "FilterAvailable", FontToString(fgAvailable.Font), "Fonts", g.strIniFile
    
    MenuClear False
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilter.Form.Unload", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub MenuClear(ByVal bReset As Boolean)
On Error GoTo ErrSection:

    Dim i&
    
    If Not m.aMenuHandles Is Nothing Then
        For i = 0 To m.aMenuHandles.Size - 1
            DestroyMenu (m.aMenuHandles(i))
        Next
    End If
    
    DestroyMenu (m.hCategoryMenu)
    Set m.aMenuHandles = Nothing
    m.hCategoryMenu = 0
    
    If bReset Then m.hCategoryMenu = CreatePopupMenu()

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilter.MenuClear"

End Sub

Private Sub imgAdd_Click()
On Error GoTo ErrSection:

    If cmdAdd.Enabled Then cmdAdd_Click
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilter.imgAdd.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub imgRemove_Click()
On Error GoTo ErrSection:

    If cmdRemove.Enabled Then cmdRemove_Click
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilter.imgRemove.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub mnuAddCondition_Click()
On Error GoTo ErrSection:
    
    AddItems

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilter.mnuAddCondition.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub mnuChangeFont_Click()
On Error GoTo ErrSection:

    ChangeGridFont m.ActiveGrid, True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGroup.mnuChangeFont.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub mnuEditItem_Click()
On Error GoTo ErrSection:
    
    EditItem
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilter.mnuEditItem.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub mnuMoveDown_Click()
    
    cmdDown_Click
    
End Sub

Private Sub mnuMoveUp_Click()

    cmdUp_Click

End Sub

Private Sub mnuRemoveCondition_Click()
On Error GoTo ErrSection:

    RemoveItems

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilter.mnuRemoveCondition.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub optFalse_Click()
On Error GoTo ErrSection:

    ConditionEdited
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilter.optFalse.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub optPercentile_Click()
On Error GoTo ErrSection:
    
    'Need to convert current numbers from Values to Percentiles
    Dim nFld&, strFieldID$, i&, nSize&
    Dim aArray As cGdArray
    
    ' get copy of values array
    strFieldID = Trim(fraEdit.Tag)
    strFieldID = "DSV:" & Mid(strFieldID, 5)
    nFld = g.SymbolPool.FieldNumForID(strFieldID)
    If nFld >= 0 Then
        Set aArray = g.SymbolPool.ArrayTable.FieldArray(nFld, True)
        ' sort the copy and resize to get nulls out
        aArray.Sort eGdSort_DeleteNullValues
        nSize = aArray.Size
        If nSize > 0 Then
            If Len(txtLow) > 0 Then
                ' search for current low-value and calc percentile
                aArray.BinarySearch ValOfText(txtLow), i
                i = Int(100# * i / nSize + 0.5)
                txtLow = CStr(i)
            End If
            If Len(txtHigh) > 0 Then
                ' search for current high-value and calc percentile
                aArray.BinarySearch ValOfText(txtHigh), i
                i = Int(100# * i / nSize + 0.5)
                txtHigh = CStr(i)
            End If
        End If
    End If
    
    ConditionEdited
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilter.optPercentile.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub optTrue_Click()
On Error GoTo ErrSection:

    ConditionEdited
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilter.optTrue.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub optValue_Click()
On Error GoTo ErrSection:
    
    'Need to convert current numbers from Percentiles to Values
    Dim nFld&, strFieldID$, i&, nSize&
    Dim aArray As cGdArray
    
    ' get copy of values array
    strFieldID = Trim(fraEdit.Tag)
    strFieldID = "DSV:" & Mid(strFieldID, 5)
    nFld = g.SymbolPool.FieldNumForID(strFieldID)
    If nFld >= 0 Then
        Set aArray = g.SymbolPool.ArrayTable.FieldArray(nFld, True)
        ' sort the copy and resize to get nulls out
        aArray.Sort eGdSort_DeleteNullValues
        nSize = aArray.Size
        If nSize > 0 Then
            If Len(txtLow) > 0 Then
                ' get low-value at given percentile
                i = nSize * ValOfText(txtLow) / 100#
                If i >= nSize Then
                    i = nSize - 1
                ElseIf i < 0 Then
                    i = 0
                End If
                txtLow = SliderToText(aArray.Num(i), False)
            End If
            If Len(txtHigh) > 0 Then
                ' get high-value at given percentile
                i = nSize * ValOfText(txtHigh) / 100#
                If i >= nSize Then
                    i = nSize - 1
                ElseIf i < 0 Then
                    i = 0
                End If
                txtHigh = SliderToText(aArray.Num(i), False)
            End If
        End If
    End If
    
    ConditionEdited
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilter.optValue.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub sldLow_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    m.iSliding = 2
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilter.sldLow.MouseDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub sldLow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    Dim t&
    
    If m.iSliding = 2 Then
        t = sldLow.Top + Y
        If t <= sldHigh.Top + sldHigh.Height Then
            t = sldHigh.Top + sldHigh.Height
            txtLow = ""
        ElseIf t >= sldBar.ScaleHeight - sldLow.Height Then
            t = sldBar.ScaleHeight - sldLow.Height
            txtLow = ""
        Else
            txtLow = SliderToText(t)
        End If
        sldLow.Top = t
        FixSliderRectangles
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilter.sldLow.MouseMove", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub sldLow_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    m.iSliding = 0
    MoveFocus txtLow
    ConditionEdited
    MoveFocus txtLow

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilter.sldLow.MouseUp", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub sldHigh_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    m.iSliding = 1

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilter.sldHigh.MouseDown", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub sldHigh_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    Dim t&
    
    If m.iSliding = 1 Then
        t = sldHigh.Top + Y
        If t <= 0 Then
            t = 0
            txtHigh = ""
        ElseIf t >= sldLow.Top - sldHigh.Height Then
            t = sldLow.Top - sldHigh.Height
            txtHigh = ""
        Else
            txtHigh = SliderToText(t)
        End If
        sldHigh.Top = t
        FixSliderRectangles
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilter.sldHigh.MouseMove", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub sldHigh_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    m.iSliding = 0
    MoveFocus txtHigh
    ConditionEdited
    MoveFocus txtHigh
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFilter.sldHigh.MouseUp", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub tbToolbar_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
On Error GoTo ErrSection:
    
    Dim strKey$
    
    ToggleFocus Me, cmdAdd

    Select Case Tool.ID
        Case "ID_Save", "ID_SaveAs", "ID_Rename"
            Save Tool.ID
        
        Case "ID_Toolbox"
            If Not AskToSave Then
                strKey = m.Filter.ID
                Unload Me
                frmToolbox.ShowMe eTab_Filters, strKey
            End If
        
        Case "ID_Close"
            If Not AskToSave Then
                If m.bModal Then
                    Me.Hide
                Else
                    Unload Me
                End If
            End If
            
        Case "ID_ShowList"
            With frmSymbolGrid.cboList
                strKey = "FIL:" & m.Filter.ID
                If Not m.Filter.IsActive Then
                    InfBox "Only active filters can be displayed.", "i", , "Display List"
                ElseIf .SelectedItem.Key = strKey Then
                    If frmSymbolGrid.Visible Then
                        InfBox "The list is already displayed in the Symbol Grid.", "i", , "Display List"
                    Else
                        DockState(frmSymbolGrid) = eShowAsPrevious
                        MoveFocus Me
                    End If
                Else
                    On Error Resume Next
                    .ComboItems(strKey).Selected = True
                    If .SelectedItem.Key <> strKey Then
                        InfBox "Only active filters can be displayed.", "i", , "Display List"
                    Else
                        frmSymbolGrid.cboList_Click
                        If Not frmSymbolGrid.Visible Then
                            DockState(frmSymbolGrid) = eShowAsPrevious
                        End If
                        MoveFocus Me
                    End If
                End If
            End With
            
        Case "ID_Print"
            PrintMe
            
        Case "ID_Description"
            m.strDesc = frmNotes.ShowMe(m.strDesc, "Description")
            EnableToolbar True
            
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilter.tbToolbar.ToolClick", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtHigh_Change()
    If Me.ActiveControl Is txtHigh Then
        m.bValueChanged = True
    End If
End Sub

Private Sub txtHigh_GotFocus()
On Error GoTo ErrSection:

    With txtHigh
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    EnableButtons
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilter.txtHigh.GotFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtHigh_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:
    
    Dim dVal#, bAdjust As Boolean
    
    If Len(Trim(txtHigh)) > 0 Then
        If KeyCode = vbKeyUp Then
            bAdjust = True
            dVal = Int(ValOfText(txtHigh)) + 1
        ElseIf KeyCode = vbKeyDown Then
            bAdjust = True
            dVal = Int(ValOfText(txtHigh)) - 1
        End If
    ElseIf KeyCode = vbKeyDown And optPercentile Then
        bAdjust = True
        dVal = 99
    End If
    If bAdjust Then
        KeyCode = 0
        txtHigh = CStr(dVal)
        If Len(Trim(txtLow)) > 0 Then
            If ValOfText(txtLow) > dVal Then
                txtLow = CStr(dVal)
            End If
        End If
        ConditionEdited
        MoveFocus txtHigh
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilter.txtHigh.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtHigh_KeyPress(KeyAscii As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If KeyAscii = Asc(vbCrLf) Then
        KeyAscii = 0
        MoveFocus txtLow
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilter.txtHigh.KeyPress", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtHigh_LostFocus()
On Error GoTo ErrSection:

    If Trim(txtHigh) <> "" And ValOfText(txtHigh) < ValOfText(txtLow) Then
        txtLow = ""
    End If
    ConditionEdited
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilter.txtHigh.LostFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtLow_Change()
    If Me.ActiveControl Is txtLow Then
        m.bValueChanged = True
    End If
End Sub

Private Sub txtLow_GotFocus()
On Error GoTo ErrSection:

    With txtLow
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    EnableButtons
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilter.txtLow.GotFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtLow_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:
    
    Dim dVal#, bAdjust As Boolean
    
    If Len(Trim(txtLow)) > 0 Then
        If KeyCode = vbKeyUp Then
            bAdjust = True
            dVal = Int(ValOfText(txtLow)) + 1
        ElseIf KeyCode = vbKeyDown Then
            bAdjust = True
            dVal = Int(ValOfText(txtLow)) - 1
        End If
    ElseIf KeyCode = vbKeyUp And optPercentile Then
        bAdjust = True
        dVal = 1
    End If
    If bAdjust Then
        KeyCode = 0
        txtLow = CStr(dVal)
        If Len(Trim(txtHigh)) > 0 Then
            If ValOfText(txtHigh) < dVal Then
                txtHigh = CStr(dVal)
            End If
        End If
        ConditionEdited
        MoveFocus txtLow
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilter.txtLow.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtLow_KeyPress(KeyAscii As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If KeyAscii = Asc(vbCrLf) Then
        KeyAscii = 0
        MoveFocus txtHigh
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilter.txtLow.KeyPress", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtLow_LostFocus()
On Error GoTo ErrSection:

    If Trim(txtLow) <> "" And ValOfText(txtLow) > ValOfText(txtHigh) Then
        txtHigh = ""
    End If
    ConditionEdited
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilter.txtLow.LostFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub EnableButtons()
On Error GoTo ErrSection:

    Dim strID$, obj As Object, bEnable As Boolean
    
'    With fgAvailable
'        bEnable = False
'        If Me.ActiveControl Is fgAvailable And _
'                .Row >= .FixedRows And .Row < .Rows Then
'            Enable cmdAdd
'            strID = Parse(.TextMatrix(.Row, kCondCol), "|", 1)
'            Set obj = g.SymbolPool.PoolObject(strID)
'            If Not obj Is Nothing Then
'                If obj.Custom Then bEnable = True
'            End If
'            .ShowCell .Row, .Col
'        Else
'            Disable cmdAdd
'        End If
'        Enable cmdEditItem, bEnable
'    End With

    With fgUsed
        If .Row >= .FixedRows And .Row < .Rows - 1 Then
            strID = Parse(.TextMatrix(.Row, kCondCol), "|", 1)
            Set obj = g.SymbolPool.PoolObject(strID)
            If Not obj Is Nothing Then If obj.Custom Then bEnable = True
            Enable cmdEditItem, bEnable
        End If
        If .Rows <= .FixedRows Then
            fraEdit.Enabled = False
            fraEdit.Visible = False
        ElseIf Not fraEdit.Enabled Then
            fraEdit.Enabled = True
            fraEdit.Visible = True
        End If
        Enable cmdUp, (.Row >= .FixedRows + 1 And .Row < .Rows)
        Enable cmdDown, (.Row >= .FixedRows And .Row < .Rows - 1)
        Enable cmdRemove, (.Row >= .FixedRows And .Row < .Rows)
    End With
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFilter.EnableButtons", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub ConditionEdited()
On Error GoTo ErrSection:
        
    Dim strCond$, strOldCond$, nRow&
    Dim strFieldID$, ExprMode As eCOND_ExprMode, dLow#, dHigh#
    
    With Me
        ' make sure a condition is selected
        If m.nEditedCondRow > 0 Then
            nRow = m.nEditedCondRow
        Else
            nRow = fgUsed.Row
        End If
        strFieldID = .fraEdit.Tag
        If Len(strFieldID) > 0 And nRow >= fgUsed.FixedRows And nRow < fgUsed.Rows Then
            ' build new condition from values in editor
            If .fraBoolean.Visible Then
                ExprMode = eCOND_Boolean
                dLow = 0 '(unused)
                If .optFalse = True Then
                    dHigh = 0
                Else
                    dHigh = 1
                End If
            ElseIf .fraValues.Visible Then
                dLow = ValOfText(.txtLow)
                dHigh = ValOfText(.txtHigh)
                If .optPercentile Then
                    ExprMode = ExprMode Or eCOND_Percentiles
                    If dLow <= 0 Then
                        txtLow = ""
                        dLow = 0
                    End If
                    If dHigh >= 100 Then
                        txtHigh = ""
                        dHigh = 100
                    End If
                End If
                If Len(Trim(.txtLow)) > 0 Then
                    ExprMode = ExprMode Or eCOND_LowOp
                End If
                If Len(Trim(.txtHigh)) > 0 Then
                    ExprMode = ExprMode Or eCOND_HighOp
                End If
                If .chkInvert <> 0 Then
                    ExprMode = ExprMode Or eCOND_Invert
                End If
            End If
            strCond = m.Filter.CreateCondition(strFieldID, ExprMode, dLow, dHigh)
                    
            ' now see if condition has changed
            strOldCond = fgUsed.TextMatrix(nRow, kCondCol)
            If UCase(strCond) <> UCase(strOldCond) Then
                ' if so, then replace condition
                fgUsed.TextMatrix(nRow, kCondCol) = strCond
                ' reset editor (to resync: labels, etc)
                ConditionToEditor
                ' then recalc and redisplay the conditions
                RecalcConditions
            End If
        End If
    End With
    
    m.bValueChanged = False
    
    ' Enable the Save Button(s)...
    EnableToolbar True
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilter.ConditionEdited", eGDRaiseError_Raise
    
End Sub

' To recalc and redisplay the conditions and condition counts
Private Sub RecalcConditions()
On Error GoTo ErrSection:
    
    Screen.MousePointer = vbHourglass
    ' get conditions from grid (since order may have changed)
    ConditionsFromGrid
    ' recalc the conditions and counts
    m.Filter.CalcFilter
    ' redisplay the results
    ConditionsToGrid
    
    EnableButtons
    Screen.MousePointer = 0
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilter.RecalcConditions", eGDRaiseError_Raise
    
End Sub

' Loads the condition editor with a condition from the grid
Private Sub ConditionToEditor()
On Error GoTo ErrSection:

    Dim nRow&, strCond$, dMin#, dMax#, nFld&, hArray&
    Dim strFieldID$, ExprMode As eCOND_ExprMode, dLow#, dHigh#
   
    With Me
        .fraEdit.Tag = "" 'for now
        ' get condition from selected row on grid
        With .fgUsed
            If .Row >= .FixedRows And .Row < .Rows Then
                strCond = .TextMatrix(.Row, kCondCol)
                .Row = .Row
            End If
            m.nEditedCondRow = .Row
        End With
        ' parse condition
        nFld = -1
        If m.Filter.ParseCondition(strCond, strFieldID, ExprMode, dLow, dHigh) Then
            nFld = g.SymbolPool.FieldNumForID(strFieldID)
        End If
        If nFld >= 0 Then
            'set controls
            lblCondName = g.SymbolPool.ArrayTable.FieldName(nFld)
            lblCondName.Visible = True
            If ExprMode And eCOND_Boolean Then
                .fraBoolean.Visible = True
                .fraValues.Visible = False
                SetSliderBarLabels 0, 0 '(to inactivate slider bar)
                If dHigh = 0 Then
                    .optFalse = True
                Else
                    .optTrue = True
                End If
            Else
                dMin = 0
                dMax = 100
                If ExprMode And eCOND_Percentiles Then
                    .optPercentile = True
                Else
                    .optValue = True
                    hArray = g.SymbolPool.ArrayTable.FieldArrayHandle(nFld)
                    dMin = gdMinValue(hArray, 0, -1)
                    dMax = gdMaxValue(hArray, 0, -1)
                    '(if it's pretty close to 0-100 range,
                    ' just make it so)
                    If dMin >= 0 And dMin < 10 _
                        And dMax > 90 And dMax <= 100 Then
                            dMin = 0
                            dMax = 100
                    End If
                    If dMin < Fix(dMin) Then
                        dMin = Fix(dMin) - 1
                    End If
                    If dMax > Fix(dMax) Then
                        dMax = Fix(dMax) + 1
                    End If
                End If
                SetSliderBarLabels dMin, dMax
                If ExprMode And eCOND_Invert Then
                    .lblHigh = "&Above ( > ):"
                    .lblOperator = "OR"
                    .lblLow = "&Below ( < ):"
                    .chkInvert = 1
                Else
                    .lblHigh = "&Below ( <= ):"
                    .lblOperator = "and"
                    .lblLow = "&Above ( >= ):"
                    .chkInvert = 0
                End If
                If ExprMode And eCOND_LowOp Then
                    .txtLow = CStr(dLow)
                Else
                    .txtLow = ""
                End If
                If ExprMode And eCOND_HighOp Then
                    .txtHigh = CStr(dHigh)
                Else
                    .txtHigh = ""
                End If
                SetSliderBar
                .fraBoolean.Visible = False
                .fraValues.Visible = True
            End If
        Else
            lblCondName.Visible = False
            .fraBoolean.Visible = False
            .fraValues.Visible = False
        End If
        'store FieldID for reference
        .fraEdit.Tag = strFieldID
    End With

    m.bValueChanged = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilter.ConditionToEditor", eGDRaiseError_Raise

End Sub

' called from cFilter before showing the form
Private Sub InitGrids()
On Error GoTo ErrSection:

    Dim i&, j&, strID$, strCond$
    Dim strIgnore$, strDisabled$
    
    Dim ExprMode As eCOND_ExprMode
    Dim Conditions As cGdArray
    Dim Criteria As cCriteria
    Dim SymbolGroup As cSymbolGroup
    
    Set Conditions = m.Filter.Conditions
    
    If m.tbInfo Is Nothing Then
        Set m.tbInfo = New cGdTable
        m.tbInfo.CreateField eGDARRAY_Longs, eTbField_MenuID, , -1
        m.tbInfo.CreateField eGDARRAY_Longs, eTbField_MenuEnableFlag, , 0
        m.tbInfo.CreateField eGDARRAY_Longs, eTbField_CategoryID, , -1
        m.tbInfo.CreateField eGDARRAY_Strings, eTbField_CategoryName
        m.tbInfo.CreateField eGDARRAY_Strings, eTbField_CategoryItem
        m.tbInfo.CreateField eGDARRAY_Strings, eTbField_ObjID
        m.tbInfo.CreateField eGDARRAY_Strings, eTbField_ObjCondition
    Else
        m.tbInfo.NumRecords = 0
        MenuClear True
    End If

    'get items to ignore in Available grid
    strIgnore = "|GRP:ALL SYMBOLS.GRP|GRP:_FLAGS_.GRP|"
    
    'get items to disabled menu items for
    strDisabled = "|"
    For i = 0 To Conditions.Size - 1
        strID = Parse(Conditions(i), "|", 1)
        strDisabled = strDisabled & UCase(strID) & "|"
    Next
    
    'get all Groups and Criterias from pool
    For Each SymbolGroup In g.SymbolPool.SymbolGroups
        With SymbolGroup
            strID = "GRP:" & .ID
            'unless to be ignored
            If InStr(strIgnore, "|" & UCase(strID) & "|") = 0 And .IsActive = True Then
                'build "default" condition
                strCond = "|1|0|1"
                
                m.tbInfo.AddRecord ""
                j = m.tbInfo.NumRecords - 1
                m.tbInfo(eTbField_CategoryID, j) = 0
                m.tbInfo(eTbField_CategoryName, j) = "Symbol Groups"
                m.tbInfo(eTbField_CategoryItem, j) = .Name
                m.tbInfo(eTbField_ObjID, j) = strID
                m.tbInfo(eTbField_ObjCondition, j) = strCond
                If InStr(strDisabled, "|" & UCase(strID) & "|") <> 0 Then m.tbInfo(eTbField_MenuEnableFlag, j) = MF_GRAYED
            End If
        End With
    Next
    
    For Each Criteria In g.SymbolPool.Criterias
        With Criteria
            strID = "DSV:" & .ID
            'unless to be ignored
            If InStr(strIgnore, "|" & UCase(strID) & "|") = 0 And .IsActive = True Then
                'build "default" condition
                If .IsBoolean Then
                    strCond = "|1|0|1"
                Else
                    ExprMode = eCOND_LowOp Or eCOND_HighOp Or eCOND_Percentiles
                    strCond = "|" & CStr(ExprMode) & "|20|80"
                End If
            
                m.tbInfo.AddRecord ""
                j = m.tbInfo.NumRecords - 1
                m.tbInfo(eTbField_CategoryID, j) = .CategoryID
                m.tbInfo(eTbField_CategoryName, j) = .CategoryName
                m.tbInfo(eTbField_CategoryItem, j) = .Name
                m.tbInfo(eTbField_ObjID, j) = strID
                m.tbInfo(eTbField_ObjCondition, j) = strCond
                If InStr(strDisabled, "|" & UCase(strID) & "|") <> 0 Then m.tbInfo(eTbField_MenuEnableFlag, j) = MF_GRAYED
            End If
        End With
    Next
    LoadCategoryMenu
    
    With fgUsed
        .Redraw = flexRDNone
        .FixedCols = 0
        .FixedRows = 1
        .AllowUserResizing = flexResizeColumns
        .ExplorerBar = flexExNone ' flexExSortShowAndMove
        .ScrollTrack = True
        .SelectionMode = flexSelectionByRow               'flexSelectionListBox
        '.AllowUserFreezing = flexFreezeColumns
        .SheetBorder = RGB(128, 128, 128)
        .GridLinesFixed = flexGridInset
        .BackColorBkg = g.Styler.GetColor(eGrid_Background) 'RH override vbApplicationWorkspaceRGB(128, 128, 128)
    
        'Columns
        .Cols = kCountCol + 1
        .ColHidden(kCondCol) = True
        .ColHidden(kNameCol) = True
        .ColHidden(kDescCol) = True
        .Rows = .FixedRows
        .TextMatrix(0, kEnglishCol) = "Filter Conditions"
        .TextMatrix(0, kCountCol) = "# Symbols"
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        '.Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = True
        .AutoSize kEnglishCol, kCountCol, True
    
        'load Rows
        ConditionsToGrid
        
        If .Rows > .FixedRows Then
            .Row = .FixedRows
        End If
        
        .Redraw = flexRDBuffered
    End With
    
ErrExit:
    Set Conditions = Nothing
    Exit Sub
    
ErrSection:
    Set Conditions = Nothing
    RaiseError "frmFilter.InitGrids", eGDRaiseError_Raise

End Sub

Private Sub ConditionsFromGrid()
On Error GoTo ErrSection:

    Dim i&, nRow&, strCond$
    
    With fgUsed
        m.Filter.Conditions.Clear
        For nRow = .FixedRows To .Rows - 1
            strCond = .TextMatrix(nRow, kCondCol)
            If InStr(strCond, "|") > 0 Then
                m.Filter.Conditions.Add strCond
            End If
        Next
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilter.ConditionsFromGrid", eGDRaiseError_Raise

End Sub

Private Sub ConditionsToGrid()
On Error GoTo ErrSection:

    Dim i&, nRow&, strText$, nFld&
    Dim strFieldID$, ExprMode As eCOND_ExprMode, dLow#, dHigh#
    Dim Conditions As cGdArray
    Dim obj As Object
    
    Set Conditions = m.Filter.Conditions
    With fgUsed
        ' for each condition
        .Redraw = flexRDNone
        .Rows = .FixedRows + Conditions.Size
        For i = 0 To Conditions.Size - 1
            'put expression
            nRow = .FixedRows + i
            .TextMatrix(nRow, kCondCol) = Conditions(i)
            
            'build display string
            m.Filter.ParseCondition Conditions(i), strFieldID, ExprMode, dLow, dHigh
            nFld = m.Filter.CondFields(i)
            strText = ""
            If nFld > 0 Then
                strText = g.SymbolPool.ArrayTable.FieldName(nFld)
                .TextMatrix(nRow, kNameCol) = strText
                strText = strText & " "
                If ExprMode And eCOND_Boolean Then
                    If dHigh = 0 Then
                        strText = strText & " = False"
                    Else
                        strText = strText & " = True"
                    End If
                Else
                    If ExprMode And eCOND_LowOp Then
                        If ExprMode And eCOND_Invert Then
                            strText = strText & " < " & CStr(dLow)
                        Else
                            strText = strText & " >= " & CStr(dLow)
                        End If
                        If ExprMode And eCOND_Percentiles Then
                            strText = strText & "%"
                        End If
                    End If
                    If ExprMode And eCOND_HighOp Then
                        If ExprMode And eCOND_Invert Then
                            If ExprMode And eCOND_LowOp Then
                                strText = strText & " OR "
                            End If
                            strText = strText & " > " & CStr(dHigh)
                        Else
                            If ExprMode And eCOND_LowOp Then
                                strText = strText & " and "
                            End If
                            strText = strText & " <= " & CStr(dHigh)
                        End If
                        If ExprMode And eCOND_Percentiles Then
                            strText = strText & "%"
                        End If
                    End If
                End If
            End If
            .TextMatrix(nRow, kEnglishCol) = strText
            
            Select Case UCase(Left(strFieldID, 2))
            Case "GR"
                .Cell(flexcpPicture, .FixedRows + i, kEnglishCol) = Picture16(ToolbarIcon("ID_SymbolGroups"))
            Case "DS"
                .Cell(flexcpPicture, .FixedRows + i, kEnglishCol) = Picture16(ToolbarIcon("ID_Criteria"))
            End Select
            
            ' show counts
            strText = CStr(m.Filter.CondCounts(i))
            If ValOfText(strText) < 0 Then
                strText = ""
            End If
            .TextMatrix(nRow, kCountCol) = strText
            
            ' put desc
            strText = Parse(Conditions(i), "|", 1)
            Set obj = g.SymbolPool.PoolObject(strText)
            If Not obj Is Nothing Then
                .TextMatrix(nRow, kDescCol) = obj.Desc
                .TextMatrix(nRow, kNameCol) = obj.Name
            End If
        Next
        Set obj = Nothing
        
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, kEnglishCol) = "Click to add ..."
        
        .RowHeight(-1) = -1
        If .RowHeight(0) < 18 * Screen.TwipsPerPixelY Then
            .RowHeight(-1) = 18 * Screen.TwipsPerPixelY
        End If
        
        ' size column widths
        '.AutoSize kEnglishCol
        'If .ColWidth(kEnglishCol) < .Width / 2 Then
        '    .ColWidth(kEnglishCol) = .Width / 2 - Screen.TwipsPerPixelX
        'End If
        ExtendCustomColumn
        
        .Redraw = flexRDBuffered
    End With
    
ErrExit:
    Set Conditions = Nothing
    Exit Sub
    
ErrSection:
    Set Conditions = Nothing
    RaiseError "frmFilter.ConditionsToGrid", eGDRaiseError_Raise

End Sub

' To fix the slider rectangles (sldTop, sldMiddle, sldBottom)
' after the sliders (sldHigh, sldLow) have been set.
Private Sub FixSliderRectangles()
On Error GoTo ErrSection:

    Dim h&
    ' size top area
    h = sldHigh.Top
    If h <= 0 Then
        sldTop.Visible = False
    Else
        sldTop.Move 0, 0, sldBar.ScaleWidth, h
        If Not sldTop.Visible Then sldTop.Visible = True
    End If
    ' size middle area
    h = sldLow.Top - sldHigh.Top - sldHigh.Height
    If h <= 0 Then
        sldMiddle.Visible = False
    Else
        sldMiddle.Move 0, sldHigh.Top + sldHigh.Height, sldBar.ScaleWidth, h
        If Not sldMiddle.Visible Then sldMiddle.Visible = True
    End If
    ' size bottom area
    h = sldBar.ScaleHeight - sldLow.Top - sldLow.Height
    If h <= 0 Then
        sldBottom.Visible = False
    Else
        sldBottom.Move 0, sldLow.Top + sldLow.Height, sldBar.ScaleWidth, h
        If Not sldBottom.Visible Then sldBottom.Visible = True
    End If
    
    If sldHigh.Top > sldLow.Top Then
        ' if Invalid
        If sldMiddle.BackColor <> vbRed Then
            sldTop.BackColor = vbRed
            sldMiddle.BackColor = vbRed
            sldBottom.BackColor = vbRed
            sldTop.ForeColor = vbWhite
            sldMiddle.ForeColor = vbWhite
            sldBottom.ForeColor = vbWhite
            sldTop.Caption = "Invalid"
            sldMiddle.Caption = "Invalid"
            sldBottom.Caption = "Invalid"
        End If
    ElseIf chkInvert Then
        ' if Inverted
        If sldMiddle.BackColor <> sldBar.BackColor Then
            sldTop.BackColor = sldBar.ForeColor
            sldMiddle.BackColor = sldBar.BackColor
            sldBottom.BackColor = sldBar.ForeColor
            sldTop.ForeColor = vbBlack
            sldMiddle.ForeColor = vbWhite
            sldBottom.ForeColor = vbBlack
            sldTop.Caption = "True"
            sldMiddle.Caption = "False"
            sldBottom.Caption = "True"
        End If
    Else
        ' if Normal
        If sldTop.BackColor <> sldBar.BackColor Then
            sldTop.BackColor = sldBar.BackColor
            sldMiddle.BackColor = sldBar.ForeColor
            sldBottom.BackColor = sldBar.BackColor
            sldTop.ForeColor = vbWhite
            sldMiddle.ForeColor = vbBlack
            sldBottom.ForeColor = vbWhite
            sldTop.Caption = "False"
            sldMiddle.Caption = "True"
            sldBottom.Caption = "False"
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilter.FixSliderRectangles", eGDRaiseError_Raise
    
End Sub

' Set text for the labels of the slider bar
Private Sub SetSliderBarLabels(ByVal dMin#, ByVal dMax#)
On Error GoTo ErrSection:
       
    Dim i&, t&, strText$
    ' store min, max
    If dMin = dMax Then 'Exit Sub
        'unknown, so just default
        dMin = 0
        dMax = 100
    End If
    lblSlider(0).Tag = CStr(dMax)
    lblSlider(4).Tag = CStr(dMin)
    ' set labels
    For i = 0 To 4
        strText = SliderToText(dMax - (dMax - dMin) * i / 4, False)
        If optPercentile Then strText = strText & "%"
        lblSlider(i).Caption = strText
    Next

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilter.SetSliderBarLabels", eGDRaiseError_Raise

End Sub

' Converts a slider position to a formatted text value
Private Function SliderToText(ByVal dVal#, Optional ByVal bConvertFromTwips As Boolean = True) As String
On Error GoTo ErrSection:
    
    Dim dMax#, dMin#, nTotal&, strText$, strFormat$
    dMin = ValOfText(lblSlider(4).Tag)
    dMax = ValOfText(lblSlider(0).Tag)
    If dMax - dMin > 100000000 Then
        strFormat = "M"
    ElseIf dMax - dMin > 100000 Then
        strFormat = "K"
    ElseIf dMax - dMin >= 100 Or dMax = dMin Then
        strFormat = "0"
    ElseIf dMax - dMin >= 1 Then
        strFormat = "0.0#"
    Else
        strFormat = "0.0###"
    End If
    If bConvertFromTwips Then
        nTotal = sldBar.ScaleHeight - sldLow.Height
        If dVal >= 0 And dVal <= nTotal And nTotal > 0 Then
            dVal = dMax - (dMax - dMin) * dVal / nTotal
        Else
            dVal = kNullData
        End If
    End If
    If dVal = kNullData Then
        strText = ""
    ElseIf strFormat = "M" Then
        strText = Format(dVal / 1000000#, "0") & " M"
    ElseIf strFormat = "K" Then
        strText = Format(dVal / 1000#, "0") & " K"
    Else
        strText = Format(dVal, strFormat)
    End If
    
    SliderToText = strText
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmFilter.SliderToText", eGDRaiseError_Raise
    
End Function

' Sets the slider bar from the values in txtLow and txtHigh
Private Sub SetSliderBar()
On Error GoTo ErrSection:

    Dim dHigh#, dLow#, dMin#, dMax#, nTotal&, strName$
    
    ' get min and max
    dMin = ValOfText(lblSlider(4).Tag)
    dMax = ValOfText(lblSlider(0).Tag)
    If dMin = dMax Then Exit Sub '(invalid)
    ' get high and low values
    If Trim(txtHigh) = "" Then
        dHigh = dMax
    Else
        dHigh = ValOfText(txtHigh)
    End If
    If Trim(txtLow) = "" Then
        dLow = dMin
    Else
        dLow = ValOfText(txtLow)
    End If
    nTotal = sldBar.ScaleHeight - sldLow.Height
    
    ' convert to twips
    dHigh = nTotal * (dMax - dHigh) / (dMax - dMin)
    dLow = nTotal * (dMax - dLow) / (dMax - dMin)
    ' round to nearest pixel
    dHigh = Int(dHigh / Screen.TwipsPerPixelY + 0.5) * Screen.TwipsPerPixelY
    dLow = Int(dLow / Screen.TwipsPerPixelY + 0.5) * Screen.TwipsPerPixelY
    ' make sure in range
    If dHigh < 0 Then
        dHigh = 0
    ElseIf dHigh > nTotal Then
        dHigh = nTotal
    End If
    If dLow < 0 Then
        dLow = 0
    ElseIf dLow > nTotal Then
        dLow = nTotal
    End If
    ' set sliders and rectangles
    sldHigh.Top = dHigh
    sldLow.Top = dLow
    FixSliderRectangles

    ' change blank text fields to min/max
    If Not Me.ActiveControl Is Nothing Then
        strName = Me.ActiveControl.Name
    End If
    If Trim(txtHigh) = "" Then
        'txtHigh = SliderToText(sldHigh.Top)
        'MoveFocus txtHigh
    ElseIf Trim(txtLow) = "" Then
        'txtLow = SliderToText(sldLow.Top)
        'MoveFocus txtLow
    ElseIf UCase(Left(strName, 3)) = "SLD" Then
        'focus should NOT be in the slider bar (but sometimes it
        'ends up there even with TabStop=False -- I don't know why!)
        MoveFocus Me.fgUsed 'txtHigh
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilter.SetSliderBar", eGDRaiseError_Raise

End Sub

' adjust all column widths to accomodate the custom "extend column"
Private Sub ExtendCustomColumn()
On Error GoTo ErrSection:

    Dim nTotal&, i&

    With fgUsed
        .ColHidden(kExtendedCol) = True
        .Redraw = flexRDBuffered '(so .ClientWidth will be correct)
        .Redraw = flexRDNone
        nTotal = 0 * Screen.TwipsPerPixelX
        For i = 0 To .Cols - 1
            If Not .ColHidden(i) Then
                nTotal = nTotal + .ColWidth(i)
            End If
        Next
        nTotal = .ClientWidth - nTotal
        If nTotal > 0 Then .ColWidth(kExtendedCol) = nTotal
        .ColHidden(kExtendedCol) = False
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".ExtendCustomColumn", eGDRaiseError_Raise

End Sub

Public Function ShowMe(ByVal strPath As String, ByVal strID As String, Optional ByVal bModal As Boolean = False) As String
On Error GoTo ErrSection:

    Set m.Filter = New cFilter
    m.bModal = bModal

    If Len(strID) > 0 Then
        If Not m.Filter.FromFile(strPath, strID) Then
            Err.Raise vbObjectError + 1000, , strID & " could not be loaded"
        End If
    
        m.Filter.CalcFilter
    End If

    With m.Filter
        m.strName = .Name
        m.strDesc = .Desc
        If .CountMode = 1 Then optVisualCount.Value = True
        InitGrids
    End With
        
    SetEditorCaption Me, "Filter", m.strName
    EnableToolbar False
    m.bOK = False
    ShowForm Me, bModal, frmMain
    If bModal Then
        If m.bOK Then ShowMe = m.Filter.ID
        Unload Me
    End If

    picAdd.BackColor = g.nColorTheme
    Picture1.BackColor = g.nColorTheme

ErrExit:
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmFilter.ShowMe", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EnableToolbar
'' Description: Enable/Disable the controls on the toolbar appropriately
'' Inputs:      Whether to Enable or Disable
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EnableToolbar(ByVal bEnableSave As Boolean)
On Error GoTo ErrSection:

    With tbToolbar
        .Tools("ID_Toolbox").Enabled = Not m.bModal
        .Tools("ID_Save").Enabled = bEnableSave
        .Tools("ID_SaveAs").Enabled = (Trim(m.strName) <> "")
        .Tools("ID_Rename").Enabled = (Trim(m.strName) <> "")
        .Tools("ID_ShowList").Enabled = Not bEnableSave
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".EnableToolbar", eGDRaiseError_Raise
    
End Sub

' Returns True if Cancelled
Public Function AskToSave() As Boolean
On Error GoTo ErrSection:
    
    Dim strResponse As String
    
    If tbToolbar.Tools("ID_Save").Enabled Then
        If WindowState = vbMinimized Then WindowState = vbNormal
    
        strResponse = InfBox("Do you want to save your changes?", "?", "+Yes|No|-Cancel", Caption)
        Select Case strResponse
            Case "C"
                AskToSave = True
            Case "Y"
                Save "ID_Save"
        End Select
    End If
        
ErrExit:
    Exit Function

ErrSection:
    AskToSave = True
    RaiseError Me.Name & ".AskToSave"

End Function

Public Sub PrintMe()
On Error GoTo ErrSection:

    frmPrintPreview.ShowMe "CNV Filters", Me, 0

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilter.PrintMe", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenerateReport
'' Description: Callback function for the Print Preview
'' Inputs:      Variant set of arguments from the Print Preview control
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GenerateReport(ByVal vArgs As Variant)
On Error GoTo ErrSection:

    Dim lRow As Long
    Dim lCol As Long
    Dim strText As String

    With frmPrintPreview.vp
        .StartDoc
        
        ' Header and Footer
        DoPrintHeader
        
        .Font.Name = "Times New Roman"
        .Font.Bold = True
        .Font.Size = 14
        .FontUnderline = True
        .Text = vbLf & "Filter:"
        .FontUnderline = False
        .Text = "    " & Trim(m.strName) & vbLf
        .Font.Size = 12
        .Font.Bold = False
        .Text = "Description: " & Trim(m.strDesc) & vbLf
        .Text = vbLf & vbLf
        
        If frmPrintPreview.GoingToFile Then
            With fgUsed
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
            .RenderControl = fgUsed.hWnd
        End If
        
        .EndDoc
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilter.GenerateReport", eGDRaiseError_Raise

End Sub

Private Sub AddItems()
On Error GoTo ErrSection:

    Dim i&
    Dim strData As String
    Dim aTbIdx As cGdArray
    Dim ItemInfo As MENUITEMINFO
    
    If m.nMenuSelectID > 0 Then
        Set aTbIdx = m.tbInfo.CreateSortedIndex(eTbField_MenuID)
        If Not aTbIdx Is Nothing Then
            If m.tbInfo.SearchAsIndex(aTbIdx, eTbField_MenuID, m.nMenuSelectID, i) Then
                i = aTbIdx(i)
                'add new item to filter grid
                strData = m.tbInfo(eTbField_ObjID, i) & m.tbInfo(eTbField_ObjCondition, i)
                fgUsed.AddItem strData
                
                ItemInfo.cbSize = Len(ItemInfo)
                ItemInfo.fMask = MIIM_ID
                'locate menu item & gray it out
                If GetMenuItemInfo(m.hCategoryMenu, m.nMenuSelectID, 0, ItemInfo) <> 0 Then
                    If ItemInfo.wID = m.nMenuSelectID Then
                        i = EnableMenuItem(m.hCategoryMenu, m.nMenuSelectID, MF_GRAYED)
                        m.tbInfo(eTbField_MenuEnableFlag, aTbIdx(i)) = MF_GRAYED
                    End If
                End If
                
                MoveFocus fgUsed
                RecalcConditions
                fgUsed.Select fgUsed.Rows - 2, 0
                fgUsed.Redraw = flexRDBuffered
            
                ' Enable the Save Button(s)...
                EnableToolbar True
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilter.AddItems", eGDRaiseError_Raise
    
End Sub

Private Sub RemoveItems()
On Error GoTo ErrSection:

    Dim i&, j&, strID$
    Dim lNumSelected As Long            ' Number of rows that are selected
    Dim lMenuID As Long                 ' menu ID of selected item
    
    Dim aRemove As New cGdArray
    Dim aTbIdx As cGdArray
    Dim ItemInfo As MENUITEMINFO
    
    With fgUsed
        ' Set up the array of selected rows in the used grid
        lNumSelected = .SelectedRows
        
        If lNumSelected = 0 Then
            If m.nBeginSelRow <= 0 Or m.nEndSelRow <= 0 Then
                If .Row >= .FixedRows And .Row < .Rows - 1 Then
                    aRemove(0) = .Row
                    lNumSelected = 1
                End If
            Else
                j = 0
                For i = m.nBeginSelRow - 1 To m.nEndSelRow - 1
                    aRemove(j) = i + 1
                    j = j + 1
                Next
                lNumSelected = aRemove.Size
            End If
        Else
            For i = 0 To lNumSelected - 1
                aRemove(i) = .SelectedRow(i)
            Next
        End If
    
    End With
    
    m.nBeginSelRow = -1     'reset
    m.nEndSelRow = -1
            
    With fgUsed
        .Redraw = flexRDNone
        ' Walk through the selected rows and move them to the available grid
        For i = lNumSelected - 1 To 0 Step -1
            strID = Parse(.TextMatrix(aRemove(i), kCondCol), "|", 1)
            If Len(strID) > 0 Then
                'remove row from grid
                .RemoveItem aRemove(i)
                
                'locate menu item & enable it if found
                Set aTbIdx = m.tbInfo.CreateSortedIndex(eTbField_ObjID)
                If Not aTbIdx Is Nothing Then
                    If m.tbInfo.SearchAsIndex(aTbIdx, eTbField_ObjID, strID, j) Then
                        j = aTbIdx(j)
                        lMenuID = m.tbInfo(eTbField_MenuID, j)
                        
                        ItemInfo.cbSize = Len(ItemInfo)
                        ItemInfo.fMask = MIIM_ID
                        If GetMenuItemInfo(m.hCategoryMenu, lMenuID, 0, ItemInfo) Then
                            EnableMenuItem m.hCategoryMenu, lMenuID, 0
                            m.tbInfo(eTbField_MenuEnableFlag, j) = 0
                        End If
                    End If
                End If
                
                ' Enable the Save button(s)
                'EnableToolbar True
            End If
        Next
        .Redraw = flexRDBuffered
    End With

    If aRemove.Size > 0 Then
        MoveFocus fgUsed
        RecalcConditions
        ConditionToEditor
        MoveFocus fgUsed
    
        ' Enable the Save button(s)
        EnableToolbar True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilter.RemoveItems", eGDRaiseError_Raise
    
End Sub

Private Sub EditItem()
On Error GoTo ErrSection:

    Dim strID As String                 ' ID of the item to edit
    Dim strType As String               ' Type of object to edit
    Dim obj As Object                   ' The object to edit
    Dim frm As Form                     ' Editor to bring it up in
    
    With fgUsed
        If .Row >= .FixedRows And .Row < .Rows Then
            strID = Parse(.TextMatrix(.Row, kCondCol), "|", 1)
            Set obj = g.SymbolPool.PoolObject(strID)
            If Not obj Is Nothing Then
                If obj.Custom Then
                    strType = UCase(Left(strID, 3))
                    strID = Mid(strID, 5)
                End If
            End If
        End If
    End With
    Select Case strType
        Case "DSV":
            Set frm = New frmCriteria
            frm.ShowMe AddSlash(App.Path) & "Custom\", strID

        Case "GRP":
            Set frm = New frmSymbolGroup
            frm.ShowMe AddSlash(App.Path) & "Custom\", strID
            
        Case Else:
            Beep
            
    End Select

ErrExit:
    Set frm = Nothing
    Exit Sub
    
ErrSection:
    Set frm = Nothing
    RaiseError "frmFilter.EditItem", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NewItem
'' Description: Allow the user to create a new filter criteria
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub NewItem()
On Error GoTo ErrSection:

    Dim frm As frmCriteria              ' Form to bring up the new critera in
    
    Set frm = New frmCriteria
    frm.ShowMe AddSlash(App.Path) & "Custom\", ""

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilter.NewItem", eGDRaiseError_Raise
    
End Sub

Private Sub LoadCategoryMenu()
On Error GoTo ErrSection:

    Dim rc&, i&, j&
    Dim iHandle&, hMenu&, iFlag&
    Dim CategoryID&, PrevCategoryID&, MenuID&
    Dim strName As String
    
    Dim aTbIndex As cGdArray

    If IsMenu(m.hCategoryMenu) = 0 Then Exit Sub
    If GetMenuItemCount(m.hCategoryMenu) > 0 Then Exit Sub
    
    If m.tbInfo Is Nothing Then Exit Sub
    If m.tbInfo.NumRecords = 0 Then Exit Sub
    
    'sort by category name then criteria name
    Set aTbIndex = m.tbInfo.CreateSortedIndex(eTbField_CategoryName, eGdSort_Default, eTbField_CategoryItem, eGdSort_Default)
    If aTbIndex Is Nothing Then Exit Sub
    
    If m.aMenuHandles Is Nothing Then Set m.aMenuHandles = New cGdArray
    
    MenuID = 0
    iHandle = -1
    m.aMenuHandles.Size = 0
    PrevCategoryID = -1
    
    For i = 0 To aTbIndex.Size - 1
        j = aTbIndex(i)
        CategoryID = m.tbInfo(eTbField_CategoryID, j)
        If CategoryID <> PrevCategoryID Then
            hMenu = CreatePopupMenu()
            If IsMenu(hMenu) Then
                iHandle = iHandle + 1
                m.aMenuHandles(iHandle) = hMenu
                
                strName = m.tbInfo(eTbField_CategoryName, j)
                rc = InsertMenu(m.hCategoryMenu, MenuID, MF_STRING Or MF_POPUP, hMenu, strName)
                
                If rc <> 0 Then
                    m.tbInfo(eTbField_MenuID, j) = MenuID
                    MenuID = MenuID + 1
                
                    iFlag = m.tbInfo(eTbField_MenuEnableFlag, j)
                    strName = m.tbInfo(eTbField_CategoryItem, j)
                    rc = InsertMenu(hMenu, 0, MF_STRING Or iFlag, MenuID, strName)
                    If rc <> 0 Then m.tbInfo(eTbField_MenuID, j) = MenuID
                    MenuID = MenuID + 1
                End If
                
                PrevCategoryID = CategoryID
            Else
                hMenu = 0
            End If
        ElseIf hMenu <> 0 Then
            iFlag = m.tbInfo(eTbField_MenuEnableFlag, j)
            strName = m.tbInfo(eTbField_CategoryItem, j)
            rc = InsertMenu(hMenu, 0, MF_STRING Or iFlag, MenuID, strName)
            If rc <> 0 Then m.tbInfo(eTbField_MenuID, j) = MenuID
            MenuID = MenuID + 1
        End If
    Next
    
    Set aTbIndex = Nothing
    Set aTbIndex = m.tbInfo.CreateSortedIndex(eTbField_MenuID)
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilter.LoadCategoryMenu"
    
End Sub

