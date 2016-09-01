VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{3B008041-905A-11D1-B4AE-444553540000}#1.0#0"; "Vsocx6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLibrary 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Library Wizard"
   ClientHeight    =   6420
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   9960
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   9960
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2460
      Top             =   5820
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1740
      Top             =   5760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   23
      ImageHeight     =   22
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary.frx":0000
            Key             =   "kSave"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary.frx":0670
            Key             =   "kLeave"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary.frx":0970
            Key             =   "kDot"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary.frx":0B74
            Key             =   "kPackage"
         EndProperty
      EndProperty
   End
   Begin ActiveToolBars.SSActiveToolBars tbToolbar 
      Left            =   3180
      Top             =   5880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131083
      ToolBarsCount   =   1
      ToolsCount      =   6
      DisplayContextMenu=   0   'False
      Tools           =   "frmLibrary.frx":0ED8
      ToolBars        =   "frmLibrary.frx":10DD
   End
   Begin VB.Frame fraWizardToolbar 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   435
      Left            =   180
      TabIndex        =   112
      Top             =   5820
      Width           =   9615
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Next >>"
         Height          =   405
         Left            =   6555
         TabIndex        =   53
         Top             =   30
         Width           =   1200
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "<< &Back"
         Height          =   405
         Left            =   5265
         TabIndex        =   52
         Top             =   30
         Width           =   1200
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   420
         Left            =   0
         TabIndex        =   51
         Top             =   0
         Width           =   1200
      End
      Begin VB.Label lblPage 
         Caption         =   "Page 1 of 5"
         Height          =   240
         Left            =   7995
         TabIndex        =   113
         Top             =   105
         Width           =   915
      End
   End
   Begin vsOcx6LibCtl.vsIndexTab tabLibrary 
      Height          =   5535
      Left            =   120
      TabIndex        =   54
      Top             =   120
      Width           =   9690
      _ExtentX        =   17092
      _ExtentY        =   9763
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
      Caption         =   "Description|Author|Library Permissions|Item Permissions|&Definition|Library &Permission|Library &Items|Include &Files"
      Align           =   0
      Appearance      =   1
      CurrTab         =   5
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
      Begin VB.Frame fraIncludeFiles 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   5160
         Left            =   10635
         TabIndex        =   116
         Top             =   330
         Width           =   9600
         Begin VB.Frame fraFileButtons 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   3375
            Left            =   8040
            TabIndex        =   118
            Top             =   240
            Width           =   1335
            Begin VB.CommandButton cmdRemoveFile 
               Caption         =   "Remove File"
               Height          =   435
               Left            =   0
               TabIndex        =   120
               Top             =   480
               Width           =   1335
            End
            Begin VB.CommandButton cmdAddFile 
               Caption         =   "Add File"
               Height          =   435
               Left            =   0
               TabIndex        =   119
               Top             =   0
               Width           =   1335
            End
         End
         Begin VSFlex7LCtl.VSFlexGrid fgFiles 
            Height          =   4875
            Left            =   120
            TabIndex        =   117
            Top             =   120
            Width           =   7695
            _cx             =   13573
            _cy             =   8599
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
      Begin VB.Frame fraAuthor 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   5160
         Left            =   -11145
         TabIndex        =   93
         Top             =   330
         Width           =   9600
         Begin VB.TextBox txtAuthor 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   1575
            TabIndex        =   3
            Top             =   1545
            Width           =   4455
         End
         Begin VB.TextBox txtEMail 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   285
            Index           =   0
            Left            =   1575
            TabIndex        =   5
            Top             =   2685
            Width           =   4455
         End
         Begin VB.TextBox txtPhoneNumber 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   1575
            TabIndex        =   4
            Top             =   2115
            Width           =   2115
         End
         Begin VB.TextBox txtWebSite 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   285
            Index           =   0
            Left            =   1575
            TabIndex        =   6
            Top             =   3330
            Width           =   4455
         End
         Begin VB.CommandButton cmdLookup 
            Caption         =   "..."
            Height          =   285
            Index           =   0
            Left            =   8280
            TabIndex        =   8
            Top             =   4185
            Width           =   345
         End
         Begin VB.TextBox txtFileName 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   1575
            TabIndex        =   7
            Top             =   4185
            Width           =   6690
         End
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   645
            Index           =   1
            Left            =   120
            Picture         =   "frmLibrary.frx":123D
            ScaleHeight     =   645
            ScaleWidth      =   630
            TabIndex        =   94
            TabStop         =   0   'False
            Top             =   240
            Width           =   630
         End
         Begin VB.Label Label1 
            Caption         =   "Enter the name of the developer of this library"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   1575
            TabIndex        =   101
            Top             =   1305
            Width           =   4020
         End
         Begin VB.Label Label1 
            Caption         =   "Enter the E-Mail address of the developer (optional)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   1575
            TabIndex        =   100
            Top             =   2475
            Width           =   5565
         End
         Begin VB.Label Label1 
            Caption         =   "Enter the phone number of the developer (optional)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   1575
            TabIndex        =   99
            Top             =   1875
            Width           =   6450
         End
         Begin VB.Label Label1 
            Caption         =   "Enter the Web site (optional)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   6
            Left            =   1590
            TabIndex        =   98
            Top             =   3105
            Width           =   3645
         End
         Begin VB.Label Label1 
            Caption         =   "Enter the path of a rich text file that contains detailed information about your library, your company, etc."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   7
            Left            =   1575
            TabIndex        =   97
            Top             =   3765
            Width           =   6450
         End
         Begin VB.Label Label4 
            Caption         =   "Author/Developer Information"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   1140
            TabIndex        =   96
            Top             =   180
            Width           =   5355
         End
         Begin VB.Label Label1 
            Caption         =   $"frmLibrary.frx":1E0F
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   9
            Left            =   1140
            TabIndex        =   95
            Top             =   570
            Width           =   7725
         End
      End
      Begin VB.Frame fraLibPermissions1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   5160
         Left            =   -10845
         TabIndex        =   87
         Top             =   330
         Width           =   9600
         Begin VB.Frame fLibraryPermissions 
            Caption         =   "Library Definition Permissions"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1740
            Index           =   1
            Left            =   1260
            TabIndex        =   89
            Top             =   1740
            Width           =   6975
            Begin VB.OptionButton LibAccess 
               Caption         =   "RESTRICTED permission (cannot edit or view without a password)"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   1
               Left            =   360
               TabIndex        =   10
               Top             =   1020
               Value           =   -1  'True
               Width           =   5550
            End
            Begin VB.OptionButton LibAccess 
               Caption         =   "FULL permissions (can edit and view item)"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   360
               TabIndex        =   9
               Top             =   420
               Width           =   5460
            End
         End
         Begin VB.CheckBox chkCannotDelete 
            Caption         =   "Check this box to prevent this library from being deleted"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   1440
            TabIndex        =   12
            Top             =   4320
            Visible         =   0   'False
            Width           =   6525
         End
         Begin VB.TextBox txtPassword 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   1425
            TabIndex        =   11
            Top             =   3825
            Width           =   1965
         End
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   600
            Index           =   3
            Left            =   240
            Picture         =   "frmLibrary.frx":1EC7
            ScaleHeight     =   600
            ScaleWidth      =   555
            TabIndex        =   88
            TabStop         =   0   'False
            Top             =   270
            Width           =   555
         End
         Begin VB.Label Label3 
            Caption         =   $"frmLibrary.frx":2269
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   945
            Index           =   3
            Left            =   1200
            TabIndex        =   92
            Top             =   645
            Width           =   7050
         End
         Begin VB.Label Label3 
            Caption         =   "Enter a password (if RESTRICTED chosen)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   0
            Left            =   3540
            TabIndex        =   91
            Top             =   3885
            Width           =   4080
         End
         Begin VB.Label Label4 
            Caption         =   "Library Permissions"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   2
            Left            =   1200
            TabIndex        =   90
            Top             =   210
            Width           =   5355
         End
      End
      Begin VB.Frame fraItemPremissions 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   5160
         Left            =   -10545
         TabIndex        =   77
         Top             =   330
         Width           =   9600
         Begin VB.Frame fItemPermissions 
            Caption         =   "Default Item Permissions"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2325
            Left            =   1290
            TabIndex        =   80
            Top             =   1155
            Width           =   6465
            Begin VB.TextBox txtDefaultPassword 
               Height          =   315
               Left            =   570
               TabIndex        =   17
               Top             =   1860
               Width           =   2490
            End
            Begin VB.OptionButton optItemPermission 
               Caption         =   "NO ACCESS permission (no access to item.  It does not show up in menus)"
               Height          =   360
               Index           =   3
               Left            =   285
               TabIndex        =   16
               Top             =   1380
               Width           =   5715
            End
            Begin VB.OptionButton optItemPermission 
               Caption         =   "RESTRICTED permission (cannot edit or view without a password)"
               Height          =   360
               Index           =   2
               Left            =   270
               TabIndex        =   15
               Top             =   1005
               Width           =   5430
            End
            Begin VB.OptionButton optItemPermission 
               Caption         =   "PARTIAL permission (can view but not edit without a password)"
               Height          =   285
               Index           =   1
               Left            =   270
               TabIndex        =   14
               Top             =   675
               Width           =   5595
            End
            Begin VB.OptionButton optItemPermission 
               Caption         =   "FULL permissions (can edit and view item)"
               Height          =   360
               Index           =   0
               Left            =   270
               TabIndex        =   13
               Top             =   285
               Value           =   -1  'True
               Width           =   5745
            End
            Begin VB.Label Label6 
               Caption         =   "Default password"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3135
               TabIndex        =   81
               Top             =   1905
               Width           =   1425
            End
         End
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   600
            Index           =   2
            Left            =   240
            Picture         =   "frmLibrary.frx":23A8
            ScaleHeight     =   600
            ScaleWidth      =   555
            TabIndex        =   79
            TabStop         =   0   'False
            Top             =   240
            Width           =   555
         End
         Begin VB.Frame fraLibraryType 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   255
            Index           =   1
            Left            =   1740
            TabIndex        =   78
            Top             =   4740
            Width           =   3735
            Begin VB.OptionButton optComDLL 
               Caption         =   "COM/VB DLL"
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   20
               Top             =   0
               Value           =   -1  'True
               Width           =   1335
            End
            Begin VB.OptionButton optStandardDLL 
               Caption         =   "Standard (C/C++) DLL"
               Height          =   255
               Index           =   0
               Left            =   1440
               TabIndex        =   21
               Top             =   0
               Width           =   1935
            End
         End
         Begin VB.CheckBox chkDLLRequired 
            Caption         =   "One or more functions in this library reside in a DLL"
            Height          =   255
            Index           =   0
            Left            =   1260
            TabIndex        =   18
            Top             =   4050
            Width           =   3975
         End
         Begin VB.TextBox txtDLLName 
            Height          =   330
            Index           =   0
            Left            =   2715
            TabIndex        =   19
            Top             =   4365
            Width           =   2430
         End
         Begin VB.TextBox txtSecurityCode 
            Height          =   330
            Index           =   0
            Left            =   6660
            TabIndex        =   22
            Top             =   4380
            Width           =   1575
         End
         Begin VB.Label Label3 
            Caption         =   $"frmLibrary.frx":274A
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   4
            Left            =   1260
            TabIndex        =   86
            Top             =   3600
            Visible         =   0   'False
            Width           =   6450
         End
         Begin VB.Label lblDLLName 
            Caption         =   "DLL Name:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   1755
            TabIndex        =   85
            Top             =   4380
            Width           =   855
         End
         Begin VB.Label Label3 
            Caption         =   $"frmLibrary.frx":27E8
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Index           =   5
            Left            =   1200
            TabIndex        =   84
            Top             =   555
            Width           =   7410
         End
         Begin VB.Label Label4 
            Caption         =   "Library Item Permissions"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   1
            Left            =   1200
            TabIndex        =   83
            Top             =   180
            Width           =   5355
         End
         Begin VB.Label lblSecurityCode 
            Caption         =   "Security Code:"
            Height          =   255
            Index           =   0
            Left            =   5475
            TabIndex        =   82
            Top             =   4410
            Width           =   1095
         End
      End
      Begin VB.Frame fraDefinition 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   5160
         Left            =   -10245
         TabIndex        =   61
         Top             =   330
         Width           =   9600
         Begin VB.Frame Frame2 
            Caption         =   "Description"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2190
            Left            =   240
            TabIndex        =   71
            Top             =   180
            Width           =   8745
            Begin MSComCtl2.UpDown udVersion 
               Height          =   285
               Left            =   4351
               TabIndex        =   121
               Top             =   1395
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   503
               _Version        =   393216
               BuddyControl    =   "txtVersion(1)"
               BuddyDispid     =   196657
               BuddyIndex      =   1
               OrigLeft        =   4620
               OrigTop         =   1410
               OrigRight       =   4860
               OrigBottom      =   1665
               SyncBuddy       =   -1  'True
               BuddyProperty   =   65547
               Enabled         =   -1  'True
            End
            Begin VB.TextBox txtDescription 
               Height          =   690
               Index           =   1
               Left            =   2160
               MultiLine       =   -1  'True
               TabIndex        =   24
               Top             =   630
               Width           =   6120
            End
            Begin VB.TextBox txtName 
               Height          =   285
               Index           =   1
               Left            =   2160
               TabIndex        =   23
               Top             =   285
               Width           =   3750
            End
            Begin VB.TextBox txtVersion 
               Height          =   285
               Index           =   1
               Left            =   2160
               TabIndex        =   25
               Top             =   1395
               Width           =   2190
            End
            Begin VB.TextBox txtFileName 
               Height          =   285
               Index           =   1
               Left            =   2160
               TabIndex        =   26
               Top             =   1785
               Width           =   6075
            End
            Begin VB.CommandButton cmdLookup 
               Caption         =   "..."
               Height          =   285
               Index           =   1
               Left            =   8310
               TabIndex        =   27
               Top             =   1800
               Width           =   345
            End
            Begin VB.PictureBox Picture2 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   5
               Left            =   270
               Picture         =   "frmLibrary.frx":28A0
               ScaleHeight     =   480
               ScaleWidth      =   510
               TabIndex        =   72
               TabStop         =   0   'False
               Top             =   360
               Width           =   510
            End
            Begin VB.Label Label5 
               Caption         =   "Short Description"
               Height          =   390
               Index           =   1
               Left            =   990
               TabIndex        =   76
               Top             =   660
               Width           =   1095
            End
            Begin VB.Label Label5 
               Caption         =   "Library Name"
               Height          =   270
               Index           =   0
               Left            =   990
               TabIndex        =   75
               Top             =   315
               Width           =   1095
            End
            Begin VB.Label Label5 
               Caption         =   "Version"
               Height          =   255
               Index           =   6
               Left            =   1005
               TabIndex        =   74
               Top             =   1395
               Width           =   1065
            End
            Begin VB.Label Label5 
               Caption         =   "Info File Name"
               Height          =   255
               Index           =   8
               Left            =   1005
               TabIndex        =   73
               Top             =   1785
               Width           =   1065
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Developer"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2415
            Left            =   240
            TabIndex        =   62
            Top             =   2550
            Width           =   8745
            Begin VB.Frame fraLibraryType 
               BorderStyle     =   0  'None
               Caption         =   "Frame1"
               Height          =   255
               Index           =   0
               Left            =   1440
               TabIndex        =   64
               Top             =   2040
               Width           =   3735
               Begin VB.OptionButton optComDLL 
                  Caption         =   "COM/VB DLL"
                  Height          =   255
                  Index           =   1
                  Left            =   0
                  TabIndex        =   34
                  Top             =   0
                  Value           =   -1  'True
                  Width           =   1335
               End
               Begin VB.OptionButton optStandardDLL 
                  Caption         =   "Standard (C/C++) DLL"
                  Height          =   255
                  Index           =   1
                  Left            =   1440
                  TabIndex        =   35
                  Top             =   0
                  Width           =   1935
               End
            End
            Begin VB.TextBox txtPhoneNumber 
               Height          =   285
               Index           =   1
               Left            =   6120
               TabIndex        =   29
               Top             =   285
               Width           =   1770
            End
            Begin VB.TextBox txtAuthor 
               Height          =   285
               Index           =   1
               Left            =   2160
               TabIndex        =   28
               Top             =   285
               Width           =   2475
            End
            Begin VB.TextBox txtWebSite 
               Height          =   285
               Index           =   1
               Left            =   2160
               TabIndex        =   31
               Top             =   990
               Width           =   4470
            End
            Begin VB.TextBox txtEMail 
               Height          =   285
               Index           =   1
               Left            =   2160
               TabIndex        =   30
               Top             =   645
               Width           =   4455
            End
            Begin VB.PictureBox Picture2 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   645
               Index           =   4
               Left            =   210
               Picture         =   "frmLibrary.frx":2BA2
               ScaleHeight     =   645
               ScaleWidth      =   630
               TabIndex        =   63
               TabStop         =   0   'False
               Top             =   300
               Width           =   630
            End
            Begin VB.CheckBox chkDLLRequired 
               Caption         =   "One or more functions in this library reside in a DLL"
               Height          =   255
               Index           =   1
               Left            =   960
               TabIndex        =   32
               Top             =   1440
               Width           =   4095
            End
            Begin VB.TextBox txtDLLName 
               Height          =   285
               Index           =   1
               Left            =   2400
               TabIndex        =   33
               Top             =   1725
               Width           =   2460
            End
            Begin VB.TextBox txtSecurityCode 
               Height          =   330
               Index           =   1
               Left            =   6360
               TabIndex        =   36
               Top             =   1680
               Width           =   1575
            End
            Begin VB.Label Label5 
               Caption         =   "Phone Number"
               Height          =   255
               Index           =   3
               Left            =   4905
               TabIndex        =   70
               Top             =   300
               Width           =   1095
            End
            Begin VB.Label Label5 
               Caption         =   "Author"
               Height          =   255
               Index           =   2
               Left            =   960
               TabIndex        =   69
               Top             =   300
               Width           =   1095
            End
            Begin VB.Label Label5 
               Caption         =   "WebSite"
               Height          =   255
               Index           =   5
               Left            =   945
               TabIndex        =   68
               Top             =   1050
               Width           =   1095
            End
            Begin VB.Label Label5 
               Caption         =   "eMail"
               Height          =   255
               Index           =   4
               Left            =   960
               TabIndex        =   67
               Top             =   675
               Width           =   1095
            End
            Begin VB.Label lblDLLName 
               Caption         =   "DLL Name:"
               Height          =   255
               Index           =   1
               Left            =   1440
               TabIndex        =   66
               Top             =   1740
               Width           =   855
            End
            Begin VB.Label lblSecurityCode 
               Caption         =   "Security Code:"
               Height          =   255
               Index           =   1
               Left            =   5160
               TabIndex        =   65
               Top             =   1725
               Width           =   1095
            End
         End
      End
      Begin VB.Frame fraLibPermissions2 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   5160
         Left            =   45
         TabIndex        =   57
         Top             =   330
         Width           =   9600
         Begin VB.TextBox txtRequiredMod 
            Height          =   285
            Left            =   1860
            TabIndex        =   42
            Top             =   3195
            Width           =   930
         End
         Begin VB.CheckBox chkCannotDelete 
            Caption         =   "Check this box to prevent this library from being deleted."
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   43
            Top             =   4680
            Visible         =   0   'False
            Width           =   6120
         End
         Begin VB.Frame fLibraryPermissions 
            Caption         =   "Library Definition Permissions"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1950
            Index           =   0
            Left            =   360
            TabIndex        =   58
            Top             =   300
            Width           =   8235
            Begin VB.OptionButton LibAccess 
               Caption         =   "RESTRICTED permission (cannot edit or view without a password)"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   3
               Left            =   1095
               TabIndex        =   38
               Top             =   1290
               Width           =   5550
            End
            Begin VB.PictureBox Picture2 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   600
               Index           =   6
               Left            =   195
               Picture         =   "frmLibrary.frx":3774
               ScaleHeight     =   600
               ScaleWidth      =   555
               TabIndex        =   59
               TabStop         =   0   'False
               Top             =   240
               Width           =   555
            End
            Begin VB.OptionButton LibAccess 
               Caption         =   "FULL permissions (can edit and view item)"
               Height          =   255
               Index           =   2
               Left            =   1095
               TabIndex        =   37
               Top             =   780
               Value           =   -1  'True
               Width           =   7005
            End
            Begin VB.Label Label7 
               Caption         =   "This option controls how the library definition is accessed (not the items in the library)"
               Height          =   405
               Left            =   1095
               TabIndex        =   60
               Top             =   300
               Width           =   6780
            End
         End
         Begin VB.TextBox txtPassword 
            Height          =   315
            Index           =   1
            Left            =   1035
            TabIndex        =   39
            Top             =   2670
            Width           =   1965
         End
         Begin VB.Label lblRequiredMod 
            Caption         =   "Required Module:"
            Height          =   195
            Left            =   420
            TabIndex        =   41
            Top             =   3240
            Width           =   1395
         End
         Begin VB.Label Label3 
            Caption         =   "Enter a password (If RESTRICTED access selected)"
            Height          =   255
            Index           =   7
            Left            =   3105
            TabIndex        =   40
            Top             =   2730
            Width           =   4740
         End
      End
      Begin VB.Frame fraLibItems 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   5160
         Left            =   10335
         TabIndex        =   55
         Top             =   330
         Width           =   9600
         Begin VB.Frame fraItemButtons 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   3495
            Left            =   8100
            TabIndex        =   114
            Top             =   180
            Width           =   1275
            Begin VB.TextBox txtOwners 
               Height          =   285
               Left            =   0
               TabIndex        =   49
               Top             =   1500
               Width           =   1215
            End
            Begin VB.CheckBox chkShadow 
               Caption         =   "Shado&w"
               Height          =   255
               Left            =   0
               TabIndex        =   47
               Top             =   960
               Width           =   1215
            End
            Begin VB.CheckBox chkShowLocal 
               Caption         =   "Show &Local Rules"
               Height          =   435
               Left            =   0
               TabIndex        =   50
               Top             =   3060
               Width           =   1215
            End
            Begin VB.CommandButton cmdRemove 
               Caption         =   "&Remove"
               Height          =   390
               Left            =   0
               TabIndex        =   46
               Top             =   420
               Width           =   1275
            End
            Begin VB.CommandButton cmdAdd 
               Caption         =   "&Add"
               Height          =   390
               Left            =   0
               TabIndex        =   45
               Top             =   0
               Width           =   1275
            End
            Begin VB.Label lblOwners 
               Caption         =   "&Owners:"
               Height          =   255
               Left            =   0
               TabIndex        =   48
               Top             =   1260
               Width           =   1035
            End
         End
         Begin VSFlex7LCtl.VSFlexGrid vsItems 
            Height          =   3315
            Left            =   120
            TabIndex        =   44
            Top             =   420
            Width           =   7860
            _cx             =   13864
            _cy             =   5847
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
            Rows            =   1
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   ""
            ScrollTrack     =   -1  'True
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
         Begin RichTextLib.RichTextBox txtPreview 
            Height          =   1155
            Left            =   120
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   3840
            Width           =   9300
            _ExtentX        =   16404
            _ExtentY        =   2037
            _Version        =   393217
            BackColor       =   -2147483648
            ReadOnly        =   -1  'True
            TextRTF         =   $"frmLibrary.frx":3B16
         End
         Begin VB.Label lblGridNotes 
            Caption         =   "Items in Red need to be resaved before Exporting the Library"
            Height          =   315
            Left            =   120
            TabIndex        =   115
            Top             =   120
            Width           =   7755
         End
      End
      Begin vsOcx6LibCtl.vsElastic vsElastic2 
         Height          =   5160
         Left            =   -11445
         TabIndex        =   102
         TabStop         =   0   'False
         Top             =   330
         Width           =   9600
         _ExtentX        =   16933
         _ExtentY        =   9102
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
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FloodColor      =   192
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         Appearance      =   0
         AutoSizeChildren=   0
         BorderWidth     =   6
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   1
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
         Begin VB.TextBox txtName 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   1770
            TabIndex        =   0
            Top             =   2025
            Width           =   3750
         End
         Begin VB.TextBox txtDescription 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Index           =   0
            Left            =   1800
            TabIndex        =   2
            Top             =   3780
            Width           =   6780
         End
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   690
            Index           =   0
            Left            =   405
            Picture         =   "frmLibrary.frx":3B98
            ScaleHeight     =   690
            ScaleWidth      =   660
            TabIndex        =   106
            TabStop         =   0   'False
            Top             =   255
            Width           =   660
         End
         Begin VB.TextBox txtVersion 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   1800
            TabIndex        =   1
            Top             =   2895
            Width           =   2715
         End
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   330
            Index           =   0
            Left            =   1410
            ScaleHeight     =   330
            ScaleWidth      =   315
            TabIndex        =   105
            Top             =   1800
            Width           =   315
         End
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   330
            Index           =   1
            Left            =   1395
            ScaleHeight     =   330
            ScaleWidth      =   315
            TabIndex        =   104
            Top             =   2670
            Width           =   315
         End
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   330
            Index           =   2
            Left            =   1410
            ScaleHeight     =   330
            ScaleWidth      =   315
            TabIndex        =   103
            Top             =   3540
            Width           =   315
         End
         Begin VB.Label Label1 
            Caption         =   "Enter a name that identifies your library"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   1785
            TabIndex        =   111
            Top             =   1785
            Width           =   5565
         End
         Begin VB.Label Label1 
            Caption         =   $"frmLibrary.frx":4412
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1020
            Index           =   1
            Left            =   1350
            TabIndex        =   110
            Top             =   660
            Width           =   7050
         End
         Begin VB.Label Label1 
            Caption         =   "Enter a short summary describing the contents of your library"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   1815
            TabIndex        =   109
            Top             =   3525
            Width           =   5220
         End
         Begin VB.Label Label2 
            Caption         =   "Library Wizard"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1365
            TabIndex        =   108
            Top             =   210
            Width           =   3360
         End
         Begin VB.Label Label1 
            Caption         =   "Enter the version number of your library"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   8
            Left            =   1800
            TabIndex        =   107
            Top             =   2640
            Width           =   5400
         End
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUp"
      Begin VB.Menu mnuAdd 
         Caption         =   "&Add"
      End
      Begin VB.Menu mnuRemove 
         Caption         =   "&Remove"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChangeFont 
         Caption         =   "&Change Font"
      End
   End
End
Attribute VB_Name = "frmLibrary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmLibrary.frm
'' Description: Form to allow user to enter/modify information about a library
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 04/03/2013   DAJ         Move Strategy Baskets into database
'' 05/01/2013   DAJ         Shadow Trading
'' 05/24/2013   DAJ         Don't allow library to be non-Restricted if Shadow
'' 05/28/2013   DAJ         Fixes for verifying password, have "Close" and "X" work the same
'' 07/23/2013   DAJ         Don't allow a basket to be a shadow item if it contains a filter
'' 08/13/2013   DAJ         Fix for shadow column not having check box after add item
'' 04/01/2014   DAJ         If user saves from the UI, clear the expiration date
'' 10/16/2014   DAJ         Replaced File System Object calls
'' 10/20/2014   DAJ         Replaced frmBrowseFolders
'' 04/28/2015   DAJ         Fixed issue with getting file size and date when saving files
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Option Compare Text

Private Type mPrivate
    Library As cLibrary
    lLibraryID As Long
    nMode As eGDLibraryMode
    vRetVal As Variant
    nDefaultItemSecurity As Byte
    strDefaultItemPassword As String
    astrRemoved As cGdArray
    astrItems As cGdArray
    bAscending As Boolean
    lColumn As Long
    lRowHeight As Long
    bSaved As Boolean
    
    strSavePassword As String
End Type
Private m As mPrivate

Private Enum eGDTabs
    eGDTab_Description = 0
    eGDTab_Author = 1
    eGDTab_LibPermissions = 2
    eGDTab_ItemPermissions = 3
    eGDTab_Summary = 4
    eGDTab_LibPermissions2 = 5
    eGDTab_Items = 6
    eGDTab_IncludeFiles = 7
End Enum

'Controls mode (normal or wizard)
Private Enum eGDLibraryMode
    eGDLibMode_WizardMode = 0
    eGDLibMode_NormalMode = 1
End Enum

'Error codeds from cLibrary Validate method
Private Const kErrLibName = 1
Private Const kErrLibDesc = 2
Private Const kErrVersion = 3
Private Const kErrAuthor = 4
Private Const kErrLibPassword = 5
Private Const kErrNoDLLName = 7
Private Const kErrFileName = 8
Private Const kErrEMail = 9
Private Const kErrWebSite = 10

Private Enum eGDCols
    eGDCol_Select = 0
    eGDCol_Name = 1
    eGDCol_ItemType = 2
    eGDCol_LibraryName = 3
    eGDCol_ItemTypeCat = 4
    eGDCol_LastModified = 5
    eGDCol_Preview = 6
    eGDCol_ID = 7
    eGDCol_SecurityLevel = 8
    eGDCol_Password = 9
    eGDCol_CannotDelete = 10
    eGDCol_SystemNumber = 11
    eGDCol_Reverify = 12
    eGDCol_Shadow = 13
    eGDCol_RequiredMod = 14
    eGDCol_NumCols = 15
End Enum

Private Enum eGDFileCols
    eFGCol_FileName = 0
    eFGCol_FileSize = 1
    eFGCol_FileDate = 2
    eFGCol_Flags = 3
    eFGCol_FileInfo = 4
    eFGCol_OnlyNewer = 5
    eFGCol_ReadOnly = 6
End Enum
Private Const kFileGridCols = 7

Private Function Tabs(ByVal lTab As eGDTabs) As Long
    Tabs = lTab
End Function
Private Function GDCol(ByVal lColumn As eGDCols) As Long
    GDCol = lColumn
End Function
Private Function FGCol(ByVal lColumn As eGDFileCols) As Long
    FGCol = lColumn
End Function

Public Property Get LibraryID() As Long
    LibraryID = m.lLibraryID
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Add
'' Description: Start the wizard for a new library
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Add()
On Error GoTo ErrSection:
    
    Dim X As Integer
    
    m.vRetVal = LockWindowUpdate(Me.hWnd)
    
    Set m.Library = New cLibrary
    m.lLibraryID = 0
    
    'This initializes the Items grid and enables the showing of permissions.
    ''Set m.LibItems = New cItems
    ''With m.LibItems
    ''    .PreviewRtf = txtPreview
    ''    .Items = vsItems
    ''    .ShowPermissions = True
    ''    .LibraryID = 0
    ''    .Load
    ''End With
        
    'Format the form...
    m.nMode = eGDLibMode_WizardMode
    tabLibrary.Appearance = apFlat
    tabLibrary.Caption = "-Description|-Author|-Library Permissions|-Item Permissions|-&Definition|-Library Permission|-Library Items"
    tabLibrary.TabOutlineColor = BackColor
    tabLibrary.CurrTab = Tabs(eGDTab_Description)
    For X = 0 To 6
        tabLibrary.TabVisible(X) = False
    Next X
    tabLibrary.TabVisible(Tabs(eGDTab_Description)) = True
    
    cmdBack.Enabled = False
    
    MoveFocus txtName(m.nMode)
    Caption = "Library Wizard"
    
    fraWizardToolbar.Visible = True
    tbToolbar.Visible = False
    
    m.nDefaultItemSecurity = 0
    
    txtAuthor(m.nMode).Text = GetIniFileProperty("Author", "", "Library", g.strIniFile)
    chkDLLRequired(m.nMode).Value = vbUnchecked
    lblDLLName(m.nMode).Enabled = False
    txtDLLName(m.nMode).Enabled = False
    lblSecurityCode(m.nMode).Enabled = False
    txtSecurityCode(m.nMode).Enabled = False
    
    InitGrid
    InitFilesGrid
    tabLibrary.TabStop = False
    
ErrExit:
    m.vRetVal = LockWindowUpdate(0)
    Exit Sub

ErrSection:
    RaiseError "frmLibrary.Add", eGDRaiseError_Raise, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadRec
'' Description: Load the information for the given library
'' Inputs:      Library ID to load
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub LoadRec(ByVal pLibraryID As Long)
On Error GoTo ErrSection:
    
    Dim X       As Integer
    
    m.vRetVal = LockWindowUpdate(Me.hWnd)
    
    Set m.Library = New cLibrary
    m.lLibraryID = pLibraryID
    With m.Library
        .LibraryID = pLibraryID
        .Load
    End With
    
''    Set m.LibItems = New cItems
''    With m.LibItems
''        .PreviewRtf = txtPreview
''        .Items = vsItems
''        .ShowPermissions = True
''        .LibraryID = pLibraryID
''        .Load
''    End With
    With vsItems
        .Redraw = flexRDNone
        InitGrid
        LoadGrid
        .Redraw = flexRDBuffered
    End With
    
    With fgFiles
        .Redraw = flexRDNone
        InitFilesGrid
        LoadFilesGrid
        .Redraw = flexRDBuffered
    End With
    
    'Format the form
    m.nMode = eGDLibMode_NormalMode
    With tabLibrary
        .TabVisible(Tabs(eGDTab_Description)) = False
        .TabVisible(Tabs(eGDTab_Author)) = False
        .TabVisible(Tabs(eGDTab_LibPermissions)) = False
        .TabVisible(Tabs(eGDTab_ItemPermissions)) = False
    End With
    
    'Move class fields to form
    With m.Library
        txtName(m.nMode).Text = Str(.LibraryName)
        txtDescription(m.nMode).Text = Str(.LibraryDesc)
        udVersion.Min = CLng(Val(.LastExported))
        udVersion.Max = 200000000
        txtVersion(m.nMode).Text = Str(.Version)
        txtVersion(m.nMode).Locked = True
        txtAuthor(m.nMode).Text = Str(.Author)
        txtPhoneNumber(m.nMode).Text = Str(.Phone)
        txtEMail(m.nMode).Text = Str(.EMail)
        txtWebSite(m.nMode).Text = Str(.WebSite)
        txtDLLName(m.nMode).Text = Str(.DLLName)
        txtSecurityCode(m.nMode).Text = Trim(Str(.DLLSecurityCode))
        If txtDLLName(m.nMode).Text = "" Then
            chkDLLRequired(m.nMode).Value = vbUnchecked
            txtDLLName(m.nMode).Enabled = False
            lblDLLName(m.nMode).Enabled = False
            txtSecurityCode(m.nMode).Enabled = False
            lblSecurityCode(m.nMode).Enabled = False
        Else
            chkDLLRequired(m.nMode).Value = vbChecked
            txtDLLName(m.nMode).Enabled = True
            lblDLLName(m.nMode).Enabled = True
            txtSecurityCode(m.nMode).Enabled = True
            lblSecurityCode(m.nMode).Enabled = True
        End If
        txtFileName(m.nMode).Text = Str(.RtfFileName)
        If .LibraryType = 1 Then
            optComDLL(m.nMode).Value = True
        Else
            optStandardDLL(m.nMode).Value = True
        End If
        
        'Format permissions
        If .SecurityLevel = 0 Then
            LibAccess(m.nMode * 2).Value = 1
            LibAccess((m.nMode * 2) + 1).Value = 0
        Else
            LibAccess(m.nMode * 2).Value = 0
            LibAccess((m.nMode * 2) + 1).Value = 1
        End If
        
        txtPassword(m.nMode).Text = Str(.Password)
        If .CannotDelete Then chkCannotDelete(m.nMode).Value = 1 Else chkCannotDelete(m.nMode).Value = 0
        
        txtRequiredMod.Text = FixRequiredMod(.RequiredMod)
        CheckBoxValue(chkShadow) = .IsGuru
        txtOwners.Text = .Owners
        
        ToggleShadowColumns
    End With
    
    tabLibrary.CurrTab = Tabs(eGDTab_Summary)
    MoveFocus txtName(m.nMode)
    SetEditorCaption Me, "Library", m.Library.LibraryName
    
    tbToolbar.Visible = True
    fraWizardToolbar.Visible = False
    
    EnableSave False
    tabLibrary.Align = asNone
    tabLibrary.TabStop = True
    
    m.nDefaultItemSecurity = 0
    
ErrExit:
    m.vRetVal = LockWindowUpdate(0)
    Exit Sub

ErrSection:
    RaiseError "frmLibrary.LoadRec", eGDRaiseError_Raise, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkShadow_Click
'' Description: Show/Hide the shadow columns on the grid as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkShadow_Click()
On Error GoTo ErrSection:

    If Visible Then
        ToggleShadowColumns
        EnableSave True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.chkShadow_Click", eGDRaiseError_Show, g.strAppPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkShowLocal_Click
'' Description: Show/Hide the Local Rules in the Grid as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkShowLocal_Click()
On Error GoTo ErrSection:

    ShowLocalRules

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.chkShowLocal.Click", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdAdd_Click
'' Description: Allow the user to add user library items to the library
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdAdd_Click()
On Error GoTo ErrSection:

    AddItem

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.cmdAdd.Click", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdAddFile_Click
'' Description: Add file(s) to the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdAddFile_Click()
On Error GoTo ErrSection:
    
    AddFile
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.cmdAddFile.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: Cancel the wizard
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancel_Click()
On Error GoTo ErrSection:

    Me.Hide
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.cmdCancel.Click", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdLookup_Clck
'' Description: Allow the user to choose a filename for library description
'' Inputs:      Which button was pressed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdLookup_Click(Index As Integer)
On Error GoTo ErrSection:
    
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName <> "" Then
        txtFileName(m.nMode).Text = CommonDialog1.FileName
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmLibrary.cmdLookup.Click", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkCannotDelete_Click
'' Description: Enable the Save button upon change
'' Inputs:      Which check box was checked
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkCannotDelete_Click(Index As Integer)
On Error GoTo ErrSection:

    EnableSave True
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.chkCannotDelete.Click", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkDLLRequired
'' Description: Enable/Disable controls based on this check box
'' Inputs:      Which check box was checked
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkDLLRequired_Click(Index As Integer)
On Error GoTo ErrSection:

    EnableSave True
    txtDLLName(Index).Enabled = (chkDLLRequired(Index).Value = vbChecked)
    lblDLLName(Index).Enabled = (chkDLLRequired(Index).Value = vbChecked)
    txtSecurityCode(Index).Enabled = (chkDLLRequired(Index).Value = vbChecked)
    lblSecurityCode(Index).Enabled = (chkDLLRequired(Index).Value = vbChecked)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.chkDLLRequired.Click", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdRemove_Click
'' Description: Remove the item from the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdRemove_Click()
On Error GoTo ErrSection:

    RemoveItem

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.cmdRemove.Click", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdRemoveFile_Click
'' Description: Remove file(s) from the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdRemoveFile_Click()
On Error GoTo ErrSection:

    RemoveFile

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.cmdRemoveFile.Click", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgFiles_AfterEdit
'' Description: Update the Flags column with what the user chose
'' Inputs:      Row and Column of Cell to Edit
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgFiles_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    Dim lTemp As Long                   ' Temporary variable

    lTemp = CLng(ValOfText(fgFiles.TextMatrix(Row, FGCol(eFGCol_Flags))))
    Select Case Col
        Case FGCol(eFGCol_OnlyNewer)
            SetBit lTemp, 1, CheckedCell(fgFiles, Row, Col)
        Case FGCol(eFGCol_ReadOnly)
            SetBit lTemp, 2, CheckedCell(fgFiles, Row, Col)
    End Select
    fgFiles.TextMatrix(Row, FGCol(eFGCol_Flags)) = Str(lTemp)
    
    EnableSave True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.fgFiles.AfterEdit", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgFiles_BeforeEdit
'' Description: Only allow the user to edit certain columns
'' Inputs:      Row and Column of Cell to Edit, Whether to Cancel the Edit
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgFiles_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    If Col <> FGCol(eFGCol_OnlyNewer) And Col <> FGCol(eFGCol_ReadOnly) Then
        Cancel = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.fgFiles.BeforeEdit", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

Private Sub fgFiles_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If KeyCode = vbKeyDelete Then
        RemoveFile
    ElseIf KeyCode = vbKeyInsert Then
        AddFile
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.fgFiles.KeyDown", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit
    
End Sub

Private Sub fgFiles_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    Dim lMouseRow As Long
    Dim lMouseCol As Long
    
    With fgFiles
        lMouseRow = .MouseRow
        lMouseCol = .MouseCol
        
        If Button = vbRightButton Then
            If lMouseRow >= .FixedRows And lMouseRow < .Rows Then
                .RowSel = lMouseRow
                If Not .IsSelected(lMouseRow) Then .Row = lMouseRow
            End If
            
            mnuRemove.Enabled = (lMouseRow >= .FixedRows And lMouseRow < .Rows)
            
            PopupMenu mnuPopUp
            If mnuPopUp.Tag = "Add" Then AddFile
            mnuPopUp.Tag = ""
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.fgFiles.MouseDown", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit
    
End Sub

Private Sub fgFiles_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    Dim lMouseRow As Long
    Dim lMouseCol As Long
    
    With fgFiles
        lMouseRow = .MouseRow
        lMouseCol = .MouseCol
        
        If lMouseRow < .FixedRows And lMouseRow >= 0 Then
            .ToolTipText = "Sort By: " & Trim(.TextMatrix(lMouseRow, lMouseCol))
        Else
            .ToolTipText = ""
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.fgFiles.MouseMove", eGDRaiseError_Show, g.strAppPath
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
    RaiseError "frmLibrary.Form.KeyDown", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Initalize the controls and the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:
    
    Dim strPlacement As String
    Dim strFont As String
    
    Icon = Picture16("kLibrary")
    
    With tbToolbar
        .Tools("ID_Save").Picture = Picture16("kSave")
        .Tools("ID_SaveAs").Picture = Picture16("kSaveAs")
        .Tools("ID_Rename").Picture = Picture16("kRename")
        .Tools("ID_Export").Picture = Picture16("kExport")
        .Tools("ID_Toolbox").Picture = Picture16("kTools")
        .Tools("ID_Cancel").Picture = Picture16("kCancel")
    End With

    CenterTheForm Me
    strPlacement = GetIniFileProperty("Lib", "", "Placement", AddSlash(g.strAppPath) & "ChartNavigator.INI")
    If strPlacement <> "" Then SetFormPlacement Me, strPlacement, "LHT"
    
    Set m.astrRemoved = New cGdArray
    m.astrRemoved.Create eGDARRAY_Strings
    Set m.astrItems = New cGdArray
    m.astrItems.Create eGDARRAY_Strings
    
    chkShowLocal = GetIniFileProperty("Lib_ShowLocal", vbChecked, "Library", AddSlash(g.strAppPath) & "ChartNavigator.INI")
    
    tabLibrary.DogEars = False
    
    ' Only show the Required Module stuff if this is a Genesis user...
    lblRequiredMod.Visible = FileExist("C:\Common\Files32.EXE")
    txtRequiredMod.Visible = FileExist("C:\Common\Files32.EXE")
    fraLibraryType(0).Visible = FileExist("C:\Common\Files32.EXE")
    fraLibraryType(1).Visible = FileExist("C:\Common\Files32.EXE")
    
    ' On a new Library, force the Version to be 1...
    txtVersion(0).Text = 1
    txtVersion(0).Locked = True
    
    mnuPopUp.Visible = False
    strFont = GetIniFileProperty("LibraryItems", "", "Fonts", g.strIniFile)
    If strFont <> "" Then FontFromString vsItems.Font, strFont
    strFont = GetIniFileProperty("LibraryFiles", "", "Fonts", g.strIniFile)
    If strFont <> "" Then FontFromString fgFiles.Font, strFont
    
    ' Hide the Show Local check box and turn it off (DAJ: 01/05/2004)...
    chkShowLocal.Value = vbUnchecked
    chkShowLocal.Visible = False
    
    CheckBoxValue(chkShadow) = False
    chkShadow.Visible = g.bShowShadow
    lblOwners.Visible = g.bShowShadow
    txtOwners.Visible = g.bShowShadow
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmLibrary.Form.Load", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LibraryExists
'' Description: Check to see if the library name already exists in the database
'' Inputs:      Library Name to check
'' Returns:     True if exists, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function LibraryExists(ByVal strLibraryName As String) As Boolean
On Error GoTo ErrSection:

    Dim rs As Recordset
    Dim QryDef As QueryDef
    
    LibraryExists = False
    Set QryDef = g.dbNav.QueryDefs("qryLibraryIDFromName")
    QryDef.Parameters(0).Value = strLibraryName
    Set rs = QryDef.OpenRecordset
    
    LibraryExists = rs.RecordCount
        
ErrExit:
    Set rs = Nothing
    Set QryDef = Nothing
    Exit Function

ErrSection:
    RaiseError "frmLibrary.LibraryExists", eGDRaiseError_Raise, g.strAppPath
    Resume ErrExit:

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: If the user hits the 'X', unload the form
'' Inputs:      Whether to Cancel the Unload, Mode of the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode <> vbFormCode Then
        Cancel = True
        CloseForm
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.Form.QueryUnload", eGDRaiseError_Show, g.strAppPath
    HandleSaveError
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Resize
'' Description: Resize and move controls as the form resizes
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Resize()
On Error Resume Next

    With tabLibrary
        If m.nMode = eGDLibMode_NormalMode Then
            .Move .Left, .Top, ScaleWidth - (.Left * 2), ScaleHeight - (.Top * 2)
        Else
            .Move .Left, .Top, ScaleWidth - (.Left * 2), _
                    ScaleHeight - fraWizardToolbar.Height - (.Top * 3)
        End If
        .Refresh
    End With
    
    With vsItems
        .Move .Left, .Top, tabLibrary.ClientWidth - fraItemButtons.Width - (.Left * 3), _
                    tabLibrary.ClientHeight - txtPreview.Height - lblGridNotes.Height - (lblGridNotes.Top * 4)
    End With
    
    With txtPreview
        .Move .Left, tabLibrary.ClientHeight - txtPreview.Height - lblGridNotes.Top, _
                    tabLibrary.ClientWidth - (.Left * 2)
    End With
    
    With fraItemButtons
        .Move tabLibrary.ClientWidth - vsItems.Left - .Width, vsItems.Top
    End With

    With fraWizardToolbar
        .Move (ScaleWidth / 2) - (.Width / 2), tabLibrary.Height + (tabLibrary.Top * 2)
    End With
    
    With fraFileButtons
        .Move tabLibrary.ClientWidth - .Width - fgFiles.Left, fgFiles.Top
    End With
    
    With fgFiles
        .Move .Left, .Top, tabLibrary.ClientWidth - fraFileButtons.Width - (.Left * 3), _
                    tabLibrary.ClientHeight - (.Top * 2)
    End With
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: Ask if the user wants to save upon unload
'' Inputs:      Whether to Cancel the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:
            
    Set m.astrRemoved = Nothing
    Set m.astrItems = Nothing
    
    SetIniFileProperty "Lib_ShowLocal", chkShowLocal, "Library", g.strIniFile
    SetIniFileProperty "Lib", GetFormPlacement(Me), "Placement", g.strIniFile
    SetIniFileProperty "LibraryItems", FontToString(vsItems.Font), "Fonts", g.strIniFile
    SetIniFileProperty "LibraryFiles", FontToString(fgFiles.Font), "Fonts", g.strIniFile
            
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmLibrary.Form.Unload", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit:

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LibAccess_Click
'' Description: Enable/Disable controls based on this
'' Inputs:      Which one was clicked
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LibAccess_Click(Index As Integer)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bNoItemsRestricted As Boolean   ' Are there any restricted items?

    If Index = 0 Or Index = 2 Then
        txtPassword(m.nMode).Enabled = False
        
        vsItems.ColHidden(GDCol(eGDCol_SecurityLevel)) = True
        vsItems.ColHidden(GDCol(eGDCol_Password)) = True
    Else
        txtPassword(m.nMode).Enabled = True
    
        With vsItems
            bNoItemsRestricted = True
            For lIndex = .FixedRows To .Rows - 1
                If .TextMatrix(lIndex, GDCol(eGDCol_SecurityLevel)) <> SecurityDesc(0) Then
                    bNoItemsRestricted = False
                    Exit For
                End If
            Next lIndex
            
            If bNoItemsRestricted = True And .Rows > .FixedRows Then
                If InfBox("Would you like to restrict all of the items in the library as well?", "?", "+Yes|-No", "Confirmation") = "Y" Then
                    For lIndex = .FixedRows To .Rows - 1
                        .TextMatrix(lIndex, GDCol(eGDCol_SecurityLevel)) = SecurityDesc(1)
                        .TextMatrix(lIndex, GDCol(eGDCol_Password)) = txtPassword(m.nMode).Text
                    Next lIndex
                End If
            End If
        End With
        
        vsItems.ColHidden(GDCol(eGDCol_SecurityLevel)) = False
        vsItems.ColHidden(GDCol(eGDCol_Password)) = False
    End If
    
    EnableSave True
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.LibAccess.Click", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit
    
End Sub

Private Sub mnuAdd_Click()
On Error GoTo ErrSection:

    mnuPopUp.Tag = "Add"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.mnuAdd.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub mnuChangeFont_Click()
On Error GoTo ErrSection:

    Select Case tabLibrary.CurrTab
        Case Tabs(eGDTab_Items)
            ChangeGridFont vsItems
        Case Tabs(eGDTab_IncludeFiles)
            ChangeGridFont fgFiles
    End Select
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.mnuChangeFont.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub mnuRemove_Click()
On Error GoTo ErrSection:

    Select Case tabLibrary.CurrTab
        Case Tabs(eGDTab_Items)
            RemoveItem
        Case Tabs(eGDTab_IncludeFiles)
            RemoveFile
    End Select
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.mnuRemove.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optComDLL_Click
'' Description: Enable/Disable controls based on this
'' Inputs:      Which one was pressed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optComDLL_Click(Index As Integer)
On Error GoTo ErrSection:

    EnableSave True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.optComDLL.Click", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optItemPermission_Click
'' Description: Enable/Disable controls based on this
'' Inputs:      Which one was pressed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optItemPermission_Click(Index As Integer)
On Error GoTo ErrSection:

    If Index > 0 Then
        txtDefaultPassword.Enabled = True
    Else
        txtDefaultPassword.Enabled = False
        m.strDefaultItemPassword = ""
    End If
    m.nDefaultItemSecurity = Index
    EnableSave True
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.optItemPermission.Click", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optStandardDLL_Click
'' Description: Enable/Disable controls based on this
'' Inputs:      Which one was clicked
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optStandardDLL_Click(Index As Integer)
On Error GoTo ErrSection:

    EnableSave True
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.optStandardDLL.Click", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tabLibrary_Switch
'' Description: When the user switches tabs, set the focus to the appropriate
''              control
'' Inputs:      Old and New tabs, Whether to Cancel Switch
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tabLibrary_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
On Error GoTo ErrSection:

    If Me.Visible And tabLibrary.Visible Then
        Cancel = False
        
        If OldTab = Tabs(eGDTab_LibPermissions2) Then
            If LibAccess((m.nMode * 2) + 1).Value = True Then
                If Len(Trim(txtPassword(m.nMode).Text)) = 0 Then
                    MoveFocus txtPassword(m.nMode)
                    InfBox "RESTRICTED Libraries must have a password", "!", , "Library Error"
                    txtPassword(m.nMode).Text = m.strSavePassword
                    
                    Cancel = True
                End If
            End If
        End If
        
        If Cancel = False Then
            Select Case NewTab
                Case Tabs(eGDTab_Author)
                    ''txtAuthor(0).SetFocus
                    MoveFocus txtAuthor(0)
                Case Tabs(eGDTab_Description)
                    ''txtName(0).SetFocus
                    MoveFocus txtName(0)
                Case Tabs(eGDTab_ItemPermissions)
                    MoveFocus optItemPermission(0)
                Case Tabs(eGDTab_Items)
                    MoveFocus vsItems
                Case Tabs(eGDTab_LibPermissions)
                    MoveFocus LibAccess(2)
                Case Tabs(eGDTab_LibPermissions2)
                    MoveFocus LibAccess(0)
                Case Tabs(eGDTab_Summary)
                    ''txtName(1).SetFocus
                    MoveFocus txtName(1)
            End Select
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.tabLibrary.Switch", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tbToolbar_ToolClick
'' Description: Execute the button that the user clicked on
'' Inputs:      Tool that the user clicked on
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tbToolbar_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
On Error GoTo ErrSection:

    Dim bSave As Boolean                ' Is the user trying to save the library?

    bSave = False
    ToggleFocus Me, tabLibrary

    Select Case Tool.ID
        Case "ID_Save", "ID_SaveAs", "ID_Rename"
            bSave = True
            Save Tool.ID
        
        Case "ID_Export"
            frmLibraryPackager.ShowMe m.lLibraryID
            If g.bReload Then
                LoadGrid
            End If
        
        Case "ID_Toolbox"
        
        Case "ID_Cancel"
            bSave = True
            CloseForm
            
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.tbToolbar.ToolClick", eGDRaiseError_Show, g.strAppPath
    If bSave = True Then
        HandleSaveError
    End If
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtAuthor_Change
'' Description: Enable/Disable controls based on this
'' Inputs:      Which one was changed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtAuthor_Change(Index As Integer)
On Error GoTo ErrSection:

    EnableSave True
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.txtAuthor.Change", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EnableSave
'' Description: Enable/Disable save button
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EnableSave(ByVal bEnable As Boolean)
On Error GoTo ErrSection:

    If tbToolbar.Visible Then
        tbToolbar.Tools("ID_Save").Enabled = bEnable
        tbToolbar.Tools("ID_Export").Enabled = Not bEnable
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.EnableSave", eGDRaiseError_Raise, g.strAppPath
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtAuthor_GotFocus
'' Description: Upon getting the focus, select the text
'' Inputs:      Which one has the focus
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtAuthor_GotFocus(Index As Integer)
On Error GoTo ErrSection:

    SelectAll txtAuthor(Index)
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.txtAuthor.GotFocus", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtDefaultPassword_Change
'' Description: Enable/Disable controls based on this
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtDefaultPassword_Change()
On Error GoTo ErrSection:

    m.strDefaultItemPassword = txtDefaultPassword.Text
    EnableSave True
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.txtDefaultPassword.Change", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtDefaultPassword_GotFocus
'' Description: Upon getting the focus, select the current text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtDefaultPassword_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtDefaultPassword
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.txtDefaultPassword.GotFocus", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtDescription_Change
'' Description: Enable/Disable controls based on this
'' Inputs:      Which one was clicked
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtDescription_Change(Index As Integer)
On Error GoTo ErrSection:

    EnableSave True
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.txtDescription.Change", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtDescription_GotFocus
'' Description: Upon getting the focus, select the current text
'' Inputs:      Which one has focus
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtDescription_GotFocus(Index As Integer)
On Error GoTo ErrSection:

    SelectAll txtDescription(Index)
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.txtDescription.GotFocus", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtDLLName_Change
'' Description: Enable/Disable controls based on this
'' Inputs:      Which one was clicked
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtDLLName_Change(Index As Integer)
On Error GoTo ErrSection:

    EnableSave True
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.txtDLLName.Change", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtDLLName_GotFocus
'' Description: Upon getting the focus, select the existing text
'' Inputs:      Which one has the focus
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtDLLName_GotFocus(Index As Integer)
On Error GoTo ErrSection:

    SelectAll txtDLLName(Index)
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.txtDLLName.GotFocus", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtDLLName_LostFocus
'' Description: Add a DLL extension if they didn't put one there
'' Inputs:      Which one was changed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtDLLName_LostFocus(Index As Integer)
On Error GoTo ErrSection:

    If FileExt(txtDLLName(Index).Text) = "" Then
        txtDLLName(Index).Text = FileBase(txtDLLName(Index).Text) & ".DLL"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.txtDLLName.LostFocus", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtEMail_Change
'' Description: Enable/Disable controls based on this
'' Inputs:      Which one was changed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtEMail_Change(Index As Integer)
On Error GoTo ErrSection:

    EnableSave True
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.txtEMail.Change", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtEMail_GotFocus
'' Description: Upon getting the focus, select the current text
'' Inputs:      Which one has the focus
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtEMail_GotFocus(Index As Integer)
On Error GoTo ErrSection:

    SelectAll txtEMail(Index)
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmLibrary.txtEMail.GotFocus", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtFileName_Change
'' Description: Enable/Disable controls based on this
'' Inputs:      Which one was changed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtFileName_Change(Index As Integer)
On Error GoTo ErrSection:

    EnableSave True
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.txtFileName.Change", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtFileName_GotFocus
'' Description: Upon getting the focus, select the current text
'' Inputs:      Which one has the focus
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtFileName_GotFocus(Index As Integer)
On Error GoTo ErrSection:

    SelectAll txtFileName(Index)
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.txtFileName.GotFocus", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtName_Change
'' Description: Enable/Disable controls based on this
'' Inputs:      Which one was changed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtName_Change(Index As Integer)
On Error GoTo ErrSection:

    EnableSave True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.txtName.Change", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtName_GotFocus
'' Description: Upon getting the focus, select the current text
'' Inputs:      Which one has the focus
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtName_GotFocus(Index As Integer)
On Error GoTo ErrSection:

    SelectAll txtName(Index)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.txtName.GotFocus", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtOwners_Change
'' Description: Enable/Disable controls based on this
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtOwners_Change()
On Error GoTo ErrSection:

    EnableSave True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.txtOwners_Change", , g.strAppPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtOwners_GotFocus
'' Description: Upon getting the focus, select the current text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtOwners_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtOwners

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.txtOwners_GotFocus", , g.strAppPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtPassword_Change
'' Description: Enable/Disable controls based on this
'' Inputs:      Which one was changed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtPassword_Change(Index As Integer)
On Error GoTo ErrSection:

    EnableSave True
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.txtPassword.Change", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtPassword_GotFocus
'' Description: Upon getting the focus, select the current text
'' Inputs:      Which one has the focus
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtPassword_GotFocus(Index As Integer)
On Error GoTo ErrSection:

    SelectAll txtPassword(Index)
    m.strSavePassword = txtPassword(Index).Text

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.txtPassword.GotFocus", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtPassword_LostFocus
'' Description: If the password has changed, change any item passwords that
''              were the same as the old one
'' Inputs:      Which one has the focus
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtPassword_LostFocus(Index As Integer)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    If ActiveControl.Name <> "cmdCancel" Then
'        If LibAccess((m.nMode * 2) + 1).Value = True Then
'            If Len(Trim(txtPassword(Index).Text)) = 0 Then
'                MoveFocus txtPassword(lIndex)
'                'InfBox "RESTRICTED Libraries must have a password", "!", , "Library Error"
'                txtPassword(Index).Text = m.strSavePassword
'                Err.Raise vbObjectError + 1000, , "RESTRICTED Libraries must have a password"
'                GoTo ErrExit
'            End If
'        End If
        
        If (LibAccess(m.nMode * 2).Value = True) Or (Len(Trim(txtPassword(Index).Text)) > 0) Then
            If Trim(txtPassword(Index).Text) <> Trim(m.strSavePassword) Then
                With vsItems
                    For lIndex = .FixedRows To .Rows - 1
                        If .TextMatrix(lIndex, GDCol(eGDCol_Password)) = Trim(m.strSavePassword) Then
                            .TextMatrix(lIndex, GDCol(eGDCol_Password)) = Trim(txtPassword(Index).Text)
                        End If
                    Next lIndex
                End With
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.txtPassword.LostFocus", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtPhoneNumber_Change
'' Description: Enable/Disable controls based on this
'' Inputs:      Which one was changed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtPhoneNumber_Change(Index As Integer)
On Error GoTo ErrSection:

    EnableSave True
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.txtPassword.Change", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtPhoneNumber_GotFocus
'' Description: Upon getting the focus, select the current text
'' Inputs:      Which one has the focus
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtPhoneNumber_GotFocus(Index As Integer)
On Error GoTo ErrSection:

    SelectAll txtPhoneNumber(Index)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.txtPhoneNumber.GotFocus", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtRequiredMod_Change
'' Description: Enable/Disable controls based on this
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtRequiredMod_Change()
On Error GoTo ErrSection:

    EnableSave True
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.txtRequiredMod.Change", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtSecurityCode_Change
'' Description: Enable/Disable controls based on this
'' Inputs:      Which one was changed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtSecurityCode_Change(Index As Integer)
On Error GoTo ErrSection:

    EnableSave True
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.txtSecurityCode.Change", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtVersion_Change
'' Description: Enable/Disable controls based on this
'' Inputs:      Which one was changed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtVersion_Change(Index As Integer)
On Error GoTo ErrSection:

    EnableSave True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.txtVersion.Change", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtVersion_GotFocus
'' Description: Upon getting the focus, select the current text
'' Inputs:      Which one has the focus
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtVersion_GotFocus(Index As Integer)
On Error GoTo ErrSection:

    SelectAll txtVersion(Index)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.txtVersion.GotFocus", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtWebSite_Change
'' Description: Enable/Disable controls based on this
'' Inputs:      Which one was changed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtWebSite_Change(Index As Integer)
On Error GoTo ErrSection:

    EnableSave True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.txtWebSite.Change", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtWebSite_GotFocus
'' Description: Upon getting the focus, select the current text
'' Inputs:      Which one has the focus
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtWebSite_GotFocus(Index As Integer)
On Error GoTo ErrSection:

    SelectAll txtWebSite(Index)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.txtWebSite.GotFocus", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    vsItems_AfterCollapse
'' Description: After an expand or collapse, reset the back colors
'' Inputs:      Row expanded or collapsed, Whether it was expanded or collapsed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub vsItems_AfterCollapse(ByVal Row As Long, ByVal State As Integer)
On Error GoTo ErrSection:

    SetBackColors

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.vsItems.AfterCollapse", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    vsItems_AfterRowColChange
'' Description: After a row/column change, set the edit cell
'' Inputs:      Old Row and Column, New Row and Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub vsItems_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    ItemPreview
    
    With vsItems
        If (NewCol <> GDCol(eGDCol_CannotDelete)) And (NewCol <> GDCol(eGDCol_Shadow)) Then
            .EditCell
        End If
        
        If .TextMatrix(NewRow, GDCol(eGDCol_ItemType)) = "Local Rule" Then
            cmdRemove.Enabled = False
        Else
            cmdRemove.Enabled = True
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.vsItems.AfterRowColChange", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit
    
End Sub

Private Sub vsItems_AfterSort(ByVal Col As Long, Order As Integer)
On Error GoTo ErrSection:

    SetBackColors

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.vsItems.AfterSort", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

Private Sub vsItems_BeforeSort(ByVal Col As Long, Order As Integer)

    ''If Col = GDCol(eGDCol_LastModified) Then
    ''    Order = Order + 2
    ''End If

End Sub

Private Sub vsItems_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)

#If 0 Then
    With vsItems
        If .RowOutlineLevel(Row1) = .RowOutlineLevel(Row2) Then
            If .TextMatrix(Row1, m.lColumn) < .TextMatrix(Row2, m.lColumn) Then
                Cmp = -1
            ElseIf .TextMatrix(Row1, m.lColumn) > .TextMatrix(Row2, m.lColumn) Then
                Cmp = 1
            Else
                Cmp = 0
            End If
        End If
    End With
#End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    vsItems_GotFocus
'' Description: Set the edit cell
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub vsItems_GotFocus()
On Error GoTo ErrSection:

    If (vsItems.Col <> GDCol(eGDCol_CannotDelete)) And (vsItems.Col <> GDCol(eGDCol_Shadow)) Then
        vsItems.EditCell
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.vsItems.GotFocus", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdNext_Click
'' Description: Move to the next wizard screen
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdNext_Click()
On Error GoTo ErrSection:

    'Validate current tab before continuing
    If ValidTab(tabLibrary.CurrTab) Then
        Enable cmdBack, True
   
        'Finish button pressed
        If tabLibrary.CurrTab = Tabs(eGDTab_Items) Then
            Save "ID_Save"
            Me.Hide
        Else
            'Skip past tabs used only in Edit mode
            If tabLibrary.CurrTab + 1 = Tabs(eGDTab_ItemPermissions) Then
                With tabLibrary
                    .TabVisible(.CurrTab) = False
                    .CurrTab = Tabs(eGDTab_Items)
                    If vsItems.Rows = 1 Then
                        DoEvents
                        cmdAdd_Click
                    End If
                End With
                cmdNext.Caption = "Finished"
                cmdNext.FontBold = True
            Else
                With tabLibrary
                    .TabVisible(.CurrTab) = False
                    .CurrTab = .CurrTab + 1
                End With
            End If
        
            AssignPageNumber
            SetTheFocus
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.cmdNext.Click", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdBack_Click
'' Description: Move to the previous wizard screen
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdBack_Click()
On Error GoTo ErrSection:

    Enable cmdNext, True
    
    If tabLibrary.CurrTab = Tabs(eGDTab_Items) Then
        cmdNext.Caption = "Next >>"
        cmdNext.FontBold = False
        With tabLibrary
            .TabVisible(Tabs(eGDTab_Items)) = False
            .CurrTab = Tabs(eGDTab_ItemPermissions) - 1
        End With
    Else
        If tabLibrary.CurrTab - 1 = Tabs(eGDTab_Description) Then
            Enable cmdBack, False
        End If
        With tabLibrary
            .TabVisible(.CurrTab) = False
            .CurrTab = .CurrTab - 1
        End With
    End If
    
    AssignPageNumber
    SetTheFocus

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.cmdBack.Click", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AssignPageNumber
'' Description: Assign the appropriate page number
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AssignPageNumber()
On Error GoTo ErrSection:
    
    Select Case tabLibrary.CurrTab
        Case Tabs(eGDTab_Description): lblPage.Caption = "Page 1 of 5"
        Case Tabs(eGDTab_Author): lblPage.Caption = "Page 2 of 5"
        Case Tabs(eGDTab_LibPermissions): lblPage.Caption = "Page 3 of 5"
        Case Tabs(eGDTab_ItemPermissions): lblPage.Caption = "Page 4 of 5"
        Case Tabs(eGDTab_Summary):
        Case Tabs(eGDTab_Items): lblPage.Caption = "Page 5 of 5"
    End Select
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.AssignPageNumber", eGDRaiseError_Raise, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetTheFocus
'' Description: Set the focus to the appropriate tab
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetTheFocus()
On Error GoTo ErrSection:
    
    Select Case tabLibrary.CurrTab
        Case Tabs(eGDTab_Description): MoveFocus txtName(m.nMode)
        Case Tabs(eGDTab_Author): MoveFocus txtAuthor(m.nMode)
        Case Tabs(eGDTab_LibPermissions): MoveFocus fLibraryPermissions(m.nMode)
        Case Tabs(eGDTab_ItemPermissions): MoveFocus fItemPermissions
        Case Tabs(eGDTab_Summary): MoveFocus txtName(m.nMode)
        Case Tabs(eGDTab_LibPermissions2): MoveFocus fLibraryPermissions(m.nMode)
        Case Tabs(eGDTab_Items): MoveFocus vsItems
    End Select
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.SetTheFocus", eGDRaiseError_Raise, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Valid tab
'' Description: What is the valid tab now?
'' Inputs:      Tab Number
'' Returns:     True if valid, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ValidTab(ByVal lTabNumber As Long) As Boolean
On Error GoTo ErrSection:

    Select Case lTabNumber
        Case Tabs(eGDTab_Description): ValidTab = vTabDesc
        Case Tabs(eGDTab_Author): ValidTab = vTabAuthor
        Case Tabs(eGDTab_LibPermissions): ValidTab = vTabLibPermissions
        Case Tabs(eGDTab_ItemPermissions): ValidTab = vTabItemPermissions
        Case Else: ValidTab = True
    End Select
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmLibrary.ValidTab", eGDRaiseError_Raise, g.strAppPath
    Resume ErrExit

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    vTabDesc
'' Description: Figure out if current tab is valid
'' Inputs:      None
'' Returns:     True if valid, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function vTabDesc() As Boolean
On Error GoTo ErrSection:

    vTabDesc = True
    ValidateFields

ErrExit:
    Exit Function
    
ErrSection:
    If Err.Number < 0 Then
        Select Case m.Library.ErrNbr
            'Error in library name
            Case kErrLibName
                vTabDesc = False
                ''MoveFocus txtName(m.nMode)
                txtName(m.nMode).SetFocus
                
            'Error in Library Description
            Case kErrLibDesc
                vTabDesc = False
                MoveFocus txtDescription(m.nMode)
                
            'Error in Library Description
            Case kErrVersion
                vTabDesc = False
                MoveFocus txtVersion(m.nMode)
            
            Case Else
                Resume ErrExit
                
        End Select
    End If
    RaiseError "frmLibrary.vTabDesc", eGDRaiseError_Raise, g.strAppPath
    Resume ErrExit

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    vTabAuthor
'' Description: Figure out if current tab is valid
'' Inputs:      None
'' Returns:     True if valid, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function vTabAuthor() As Boolean
On Error GoTo ErrSection:

    vTabAuthor = True
    ValidateFields

ErrExit:
    Exit Function

ErrSection:
    If Err.Number < 0 Then
        Select Case m.Library.ErrNbr
        
            'Error in library name
            Case kErrAuthor
                vTabAuthor = False
                MoveFocus txtAuthor(m.nMode)
        
            'Error in eMail name
            Case kErrEMail
                vTabAuthor = False
                MoveFocus txtEMail(m.nMode)
        
            'Error in WebSite name
            Case kErrWebSite
                vTabAuthor = False
                MoveFocus txtWebSite(m.nMode)
        
            'Error in rtf Filename
            Case kErrFileName
                vTabAuthor = False
                MoveFocus txtFileName(m.nMode)
                
            Case Else
                Resume ErrExit
                
        End Select
    End If
    RaiseError "frmLibrary.vTabAuthor", eGDRaiseError_Raise, g.strAppPath
    Resume ErrExit

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    vTabLibPermissions
'' Description: Figure out if current tab is valid
'' Inputs:      None
'' Returns:     True if valid, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function vTabLibPermissions() As Boolean
On Error GoTo ErrSection:

    vTabLibPermissions = True
    'ValidateFields

    If LibAccess((m.nMode * 2) + 1).Value = True Then
        If Len(Trim(txtPassword(m.nMode).Text)) = 0 Then
            InfBox "RESTRICTED Libraries must have a password", "!", , "Library Error"
            MoveFocus txtPassword(m.nMode)
            vTabLibPermissions = False
        End If
    End If
    
ErrExit:
    Exit Function

ErrSection:
    If Err.Number < 0 Then
        Select Case m.Library.ErrNbr
        
            'Error in password name
            Case kErrLibPassword
                vTabLibPermissions = False
                MoveFocus txtPassword(m.nMode)
                
            Case Else
                Resume ErrExit
                
        End Select
    End If
    RaiseError "frmLibrary.vTabLibPermissions", eGDRaiseError_Raise, g.strAppPath
    Resume ErrExit

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    vTabItemPermissions
'' Description: Figure out if current tab is valid
'' Inputs:      None
'' Returns:     True if valid, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function vTabItemPermissions() As Boolean
On Error GoTo ErrSection:

    vTabItemPermissions = True
    
    If m.nDefaultItemSecurity > 0 Then
        If Len(m.strDefaultItemPassword) < 5 Or Len(m.strDefaultItemPassword) > 10 Then
            InfBox "Passwords must be 5 to 10 characters in length.", "i", , "Error"
            vTabItemPermissions = False
            MoveFocus txtDefaultPassword
            Exit Function
        End If
    End If
    
    ValidateFields

ErrExit:
    Exit Function

ErrSection:
    If Err.Number < 0 Then
        Select Case m.Library.ErrNbr
        
            'Error DLL name required for builtin functions
            Case kErrNoDLLName
                vTabItemPermissions = False
                MoveFocus txtDLLName(m.nMode)
                
            Case Else
                Resume ErrExit
                
        End Select
    End If
    RaiseError "frmLibrary.vTabItemPermissions", eGDRaiseError_Raise, g.strAppPath
    Resume ErrExit
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ValidateFields
'' Description: Validate the fields
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ValidateFields()
On Error GoTo ErrSection:

    With m.Library
        .LibraryName = Trim(txtName(m.nMode).Text)
        .LibraryDesc = Trim(txtDescription(m.nMode).Text)
        .Version = Trim(txtVersion(m.nMode).Text)
        .Author = Trim(txtAuthor(m.nMode).Text)
        .Phone = Trim(txtPhoneNumber(m.nMode).Text)
        .EMail = Trim(txtEMail(m.nMode).Text)
        .WebSite = Trim(txtWebSite(m.nMode).Text)
        If chkDLLRequired(m.nMode).Value = vbUnchecked Then
            txtDLLName(m.nMode).Text = ""
            'txtSecurityCode(m.nMode).Text = "0" ' TLB 5/16/2012: no longer do this
        End If
        .DLLName = Trim(txtDLLName(m.nMode).Text)
        .DLLSecurityCode = CLng(ValOfText(txtSecurityCode(m.nMode).Text))
        If optComDLL(m.nMode).Value = True Then .LibraryType = 1 Else .LibraryType = 0
        .RtfFileName = Trim(txtFileName(m.nMode).Text)
        .Items = vsItems
        .Password = Trim(txtPassword(m.nMode).Text)
        If chkCannotDelete(m.nMode).Value Then .CannotDelete = True Else .CannotDelete = False
        
        If LibAccess(m.nMode * 2).Value Then .SecurityLevel = 0
        If LibAccess((m.nMode * 2) + 1).Value Then .SecurityLevel = 2
        
        .RequiredMod = Trim(txtRequiredMod.Text)
        .IsGuru = CheckBoxValue(chkShadow)
        .Owners = txtOwners.Text

        .Validate
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmLibrary.ValidateFields", eGDRaiseError_Raise, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    vsItems_BeforeEdit
'' Description: Only allow the user to edit certain columns
'' Inputs:      Row and Column of cell to edit, Whether to cancel the edit
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub vsItems_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim strMainType As String           ' Main type for the row

    If Col < GDCol(eGDCol_SecurityLevel) Then
        Cancel = True
    Else
        strMainType = MainTypeForRow(Row)
        vsItems.ComboList = ""
        
        'Only allow changes to Permission related fields
        Select Case Col
            Case GDCol(eGDCol_SecurityLevel)
                If Row >= vsItems.FixedRows Then
                    If UCase(vsItems.TextMatrix(Row, GDCol(eGDCol_ItemType))) = "FUNCTION" Then
                        vsItems.ComboList = "|#0;" & SecurityDesc(0) & _
                                     "|#1;" & SecurityDesc(1) & _
                                     "|#2;" & SecurityDesc(2) & _
                                     "|#3;" & SecurityDesc(3)
                    Else
                        vsItems.ComboList = "|#0;" & SecurityDesc(0) & _
                                     "|#1;" & SecurityDesc(1) & _
                                     "|#2;" & SecurityDesc(2)
                    End If
                End If
                
            Case GDCol(eGDCol_Password), GDCol(eGDCol_CannotDelete)
                Cancel = False
                
            Case GDCol(eGDCol_Shadow)
                Cancel = (UCase(strMainType) <> "STRATEGY") And (UCase(strMainType) <> "BASKET")
            
            Case GDCol(eGDCol_RequiredMod)
                Cancel = (UCase(strMainType) <> "BASKET")
                
            Case Else
                Cancel = True
        
        End Select
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmLibrary.vsItems.BeforeEdit", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit:

End Sub

Private Sub vsItems_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If KeyCode = vbKeyDelete Then
        RemoveItem
    ElseIf KeyCode = vbKeyInsert Then
        AddItem
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.vsItems.KeyDown", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit
    
End Sub

Private Sub vsItems_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    Dim lMouseRow As Long
    Dim lMouseCol As Long
    
    With vsItems
        lMouseRow = .MouseRow
        lMouseCol = .MouseCol
        
        If Button = vbRightButton Then
            If lMouseRow >= .FixedRows And lMouseRow < .Rows Then
                .RowSel = lMouseRow
                If Not .IsSelected(lMouseRow) Then .Row = lMouseRow
            End If
            
            mnuRemove.Enabled = (lMouseRow >= .FixedRows And lMouseRow < .Rows)
            
            PopupMenu mnuPopUp
            
            If mnuPopUp.Tag = "Add" Then AddItem
            mnuPopUp.Tag = ""
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.vsItems.MouseDown", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit
    
End Sub

Private Sub vsItems_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    Dim lMouseRow As Long
    Dim lMouseCol As Long
    
    With vsItems
        lMouseRow = .MouseRow
        lMouseCol = .MouseCol
        
        If lMouseRow < .FixedRows And lMouseRow >= 0 Then
            .ToolTipText = "Sort By: " & Trim(.TextMatrix(lMouseRow, lMouseCol))
        Else
            .ToolTipText = ""
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.vsItems.MouseMove", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    vsItems_ValidateEdit
'' Description: Validate what the user entered
'' Inputs:      Row and Column of the cell being edited, Whether to Cancel
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub vsItems_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim iLen As Integer                 ' Length of a string
    Dim lStrategyBasketID As Long       ' Strategy Basket ID
    
    Select Case Col
        Case GDCol(eGDCol_Password)
            iLen = Len(vsItems.EditText)
            If iLen > 0 And (iLen < 5 Or iLen > 10) Then
                Cancel = True
                Err.Raise vbObjectError + 1000, , "Passwords must be 5 to 10 characters in length."
            End If
    
        Case GDCol(eGDCol_SecurityLevel)
            Select Case vsItems.EditText
                Case "Can Edit/Can View"
                Case "No Edit/Can View"
                Case "No Edit/No View"
                Case "No Access"
                Case Else
                    Cancel = True
                    Err.Raise vbObjectError + 1000, , "Please select a valid security level."
            End Select
            
        Case GDCol(eGDCol_Shadow)
            If UCase(MainTypeForRow(Row)) = "BASKET" Then
                lStrategyBasketID = CLng(Val(vsItems.TextMatrix(Row, GDCol(eGDCol_ID))))
                If StrategyBasketHasFilter(lStrategyBasketID) = True Then
                    Cancel = True
                    Err.Raise vbObjectError + 1000, , "A strategy basket with a filter cannot be a Shadow item"
                End If
            End If
            
    End Select
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmLibrary.vsItems.ValidateEdit", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit:

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    vsItems_AfterEdit
'' Description: Enable the Save button after editing
'' Inputs:      Row and Column of the cell edited
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub vsItems_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    EnableSave True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.vsItems.AfterEdit", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    vsItems_BeforeScrollTip
'' Description: Update the scroll tip when appropriate
'' Inputs:      Row to show the ToolTip for
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub vsItems_BeforeScrollTip(ByVal Row As Long)
On Error GoTo ErrSection:

    With vsItems
        .ScrollTipText = .TextMatrix(Row, 0)
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.vsItems.BeforeScrollTip", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Save
'' Description: Validate and save the library
'' Inputs:      Button pressed (Mode to Save)
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Save(ByVal strButton As String, Optional ByVal bShowMsg As Boolean = True)
On Error GoTo ErrSection:
    
    Dim strText As String               ' Text to show in a message box
    Dim strReturn As String             ' Return from the message box
    Dim bSaveAs As Boolean              ' Are we in SaveAs mode?
    
#If 0 Then
    If strButton = "ID_SaveAs" Then
        strText = "Save a copy of the current Library as..."
        
        strReturn = AskBox("h=Save As ; i=? ; g=string ; d=" & Trim(m.strName) & " ; " & strText)
        If strReturn = "" Then Err.Raise vbObjectError + 1000, , "You must enter in a name for the library"
        If Trim(UCase(strReturn)) <> Trim(UCase(m.strName)) Then
            m.strName = Trim(strReturn)
            bSaveAs = True
        End If
    ElseIf strButton = "ID_Rename" Or Len(m.strName) = 0 Then
        If Trim(m.strName) <> "" Then
            strText = "Rename the current Library as..."
        Else
            strText = "Save the current Library as..."
        End If
        
        strReturn = AskBox("h=Rename ; i=? ; g=string ; d=" & Trim(m.strName) & " ; " & strText)
        If strReturn = "" Then Err.Raise vbObjectError + 1000, , "You must enter in a name for the library"
        m.strName = Trim(strReturn)
    End If
#End If
    
    ' Ask the user if they wish to bump the version number...
    If CLng(ValOfText(txtVersion(m.nMode).Text)) = m.Library.Version Then
        If m.Library.Version = m.Library.LastExported Then
            strReturn = InfBox("It is recommended that you increase the version number when " & _
                    "you make changes to a library.||Would you like to change the version to " & _
                    m.Library.Version + 1 & "?|", , "+Yes|-No", _
                    "Confirmation")
            If strReturn = "Y" Then txtVersion(m.nMode).Text = CStr(m.Library.Version + 1&)
        End If
    End If
    
    Screen.MousePointer = vbHourglass
    
    ValidateFields
    m.Library.SaveFromUI
    
''    m.LibItems.Author = m.Library.Author
''    m.LibItems.LibraryID = m.Library.LibraryID
''    m.LibItems.Save
    m.lLibraryID = m.Library.LibraryID
    SaveItems
    SaveFiles
    
    If bShowMsg Then InfBox "Library: " & m.Library.LibraryName & " saved successfully. ", "i", , "Confirmation"
    EnableSave False
    g.bChanged = True
    m.bSaved = True
    
    SetIniFileProperty "Author", m.Library.Author, "Library", g.strIniFile
    
ErrExit:
    Screen.MousePointer = vbDefault
    Exit Sub

ErrSection:
    RaiseError "frmLibrary.Save", eGDRaiseError_Raise, g.strAppPath
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

    With vsItems
        lRedraw = .Redraw
        .Redraw = flexRDNone

        .Clear
        .Editable = flexEDKbdMouse
        .ExplorerBar = flexExSortShow
        .ExtendLastCol = True
        .SelectionMode = flexSelectionListBox
        .AllowUserResizing = flexResizeColumns
        .AllowSelection = True
        .AutoSearch = flexSearchFromTop
        ''.OutlineBar = flexOutlineBarSimpleLeaf
        ''.OutlineCol = GDCol(eGDCol_Name)
        .WordWrap = True
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .ScrollBars = flexScrollBarBoth
        .ScrollTrack = True
        .FixedCols = 0
        .FixedRows = 1
        .Rows = 1
        .Cols = GDCol(eGDCol_NumCols)
        .FrozenCols = 2                 ' Freeze the Name & Select columns
        m.lRowHeight = .RowHeight(0)
        .RowHeightMax = .Height - (.RowHeight(0) * 2)
        
        .TextMatrix(0, GDCol(eGDCol_Select)) = "Select"
        .TextMatrix(0, GDCol(eGDCol_Name)) = "Name"
        .TextMatrix(0, GDCol(eGDCol_ItemType)) = "Item"
        .TextMatrix(0, GDCol(eGDCol_LibraryName)) = "Library"
        .TextMatrix(0, GDCol(eGDCol_ItemTypeCat)) = "Item Type"
        .TextMatrix(0, GDCol(eGDCol_LastModified)) = "Last Modified"
        .TextMatrix(0, GDCol(eGDCol_SecurityLevel)) = "Permission"
        .TextMatrix(0, GDCol(eGDCol_Password)) = "Password"
        .TextMatrix(0, GDCol(eGDCol_CannotDelete)) = "Can't Delete"
        .TextMatrix(0, GDCol(eGDCol_Shadow)) = "Shadow"
        .TextMatrix(0, GDCol(eGDCol_RequiredMod)) = "Module"
        
        .ColHidden(GDCol(eGDCol_Preview)) = True
        .ColHidden(GDCol(eGDCol_ID)) = True
        .ColHidden(GDCol(eGDCol_SystemNumber)) = True
        .ColHidden(GDCol(eGDCol_Select)) = True
        .ColHidden(GDCol(eGDCol_LibraryName)) = True
        .ColHidden(GDCol(eGDCol_Reverify)) = True
        .ColHidden(GDCol(eGDCol_Password)) = (m.Library.SecurityLevel = 0)
        .ColHidden(GDCol(eGDCol_SecurityLevel)) = (m.Library.SecurityLevel = 0)
        .ColHidden(GDCol(eGDCol_Shadow)) = True
        .ColHidden(GDCol(eGDCol_RequiredMod)) = True
        
        .ColDataType(GDCol(eGDCol_Select)) = flexDTBoolean
        .ColDataType(GDCol(eGDCol_CannotDelete)) = flexDTBoolean
        .ColDataType(GDCol(eGDCol_Reverify)) = flexDTBoolean
        
        .ColFormat(GDCol(eGDCol_LastModified)) = DateAndTime("Format")
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignLeftTop
        .AutoSize 0, .Cols - 1, False, 75
        
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.InitGrid", eGDRaiseError_Raise, g.strAppPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadGrid
'' Description: Load the items grid with the items from the library
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadGrid()
On Error GoTo ErrSection:
    
    Dim lRedraw As Long                 ' Current state of the grid's redraw
    Dim rs As Recordset                 ' Recordset from the database
    Dim rs2 As Recordset                ' Secondary recordset from the database
    Dim strItemTypeCat As String        ' Item type category
    Dim strPreview As String            ' Preview string
    
    With vsItems
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        .Rows = .FixedRows
        
        'Load Systems
        If g.CalledFrom = SystemNavigator Then
            Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblSystems] " & _
                "WHERE [LibraryID]=" & Str(m.lLibraryID) & ";", dbOpenSnapshot)
            If Not (rs.BOF And rs.EOF) Then rs.MoveFirst
            Do Until rs.EOF
                If rs!CheckSum = BuildCheckSum(rs, "tblSystems") Then
                    AddRow rs!SystemName, "Strategy" & vbTab & "0", CDbl(rs!LastModified), "N/A", _
                        rs!Notes, rs!SystemNumber, _
                        rs!SecurityLevel, DecryptField(rs!Password), rs!CannotDelete, rs!Reverify, _
                        rs!IsGuru
                        
                    Set rs2 = g.dbNav.OpenRecordset("SELECT * FROM [tblRules] " & _
                        "WHERE [SystemNumber]=" & Str(rs!SystemNumber) & " " & _
                        "ORDER BY [Name];", dbOpenSnapshot)
                    If Not (rs2.BOF And rs2.EOF) Then rs2.MoveFirst
                    Do Until rs2.EOF
                        If rs2!CheckSum = BuildCheckSum(rs2, "tblRules") Then
                            If rs2!BuySell = True Then
                                If rs2!RuleType = 0 Then
                                    strItemTypeCat = "Long Entry"
                                Else
                                    strItemTypeCat = "Short Exit"
                                End If
                            Else
                                If rs2!RuleType = 0 Then
                                    strItemTypeCat = "Short Entry"
                                Else
                                    strItemTypeCat = "Long Exit"
                                End If
                            End If
                            AddRow rs2!Name, "Local Rule" & vbTab & rs2!SystemNumber, CDbl(rs2!LastModified), strItemTypeCat, _
                                DecryptField(rs2!PreviewRTF), rs2!RuleID, _
                                rs2!SecurityLevel, DecryptField(rs2!Password), rs2!CannotDelete, rs2!Reverify
                        End If
                        
                        rs2.MoveNext
                    Loop
                End If
                
                rs.MoveNext
            Loop
        Else
            'Load Portfolios for Portfolio Navigator
            Set rs = g.dbNav.OpenRecordset("Select * from [tblPortfolios] Where " & _
                "[LibraryID]=" & Str(m.lLibraryID) & _
                " Order by [LibraryID];", dbOpenSnapshot)
            Do Until rs.EOF
                AddRow rs!PortfolioName, "Portfolio" & vbTab & "0", CDbl(rs!LastModified), rs!PortfolioClass, _
                    rs!Notes, rs!PortfolioNumber, _
                    rs!SecurityLevel, rs!Password, rs!CannotDelete, rs!Reverify
                rs.MoveNext
            Loop
            
            'Load Models
            Set rs = g.dbNav.OpenRecordset("Select * from [tblModels] Where " & _
                "[LibraryID]=" & Str(m.lLibraryID) & _
                " Order by [LibraryID];", dbOpenSnapshot)
            Do Until rs.EOF
                AddRow rs!ModelName, "Model" & vbTab & "0", CDbl(rs!LastModified), rs!SecurityType, _
                    rs!Notes, rs!ModelNumber, _
                    rs!SecurityLevel, rs!Password, rs!CannotDelete, rs!Reverify
                rs.MoveNext
            Loop
        
        End If
            
        ' Load all Functions
        Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblFunctions] " & _
                "WHERE [LibraryID]=" & Str(m.lLibraryID) & ";", dbOpenSnapshot)
        If Not (rs.BOF And rs.EOF) Then rs.MoveFirst
        Do Until rs.EOF
            If rs!CheckSum = BuildCheckSum(rs, "tblFunctions") Then
                strPreview = "Usage: " & rs!TradeSenseUsage & vbCrLf & _
                         "Description: " & rs!Description
    
                AddRow rs!FunctionName, "Function" & vbTab & "0", _
                    CDbl(rs!LastModified), ImplementationTypeDesc(rs!ImplementationTypeID), _
                    strPreview, rs!FunctionID, _
                    rs!SecurityLevel, DecryptField(rs!Password), rs!CannotDelete, rs!Reverify
            End If
            
            rs.MoveNext
        Loop

        ' Load all Rules
        Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblRules] " & _
            "WHERE [LibraryID]=" & Str(m.lLibraryID) & " AND [SystemNumber]=0;", dbOpenSnapshot)
        If Not (rs.BOF And rs.EOF) Then rs.MoveFirst
        Do Until rs.EOF
            If rs!CheckSum = BuildCheckSum(rs, "tblRules") Then
                If g.CalledFrom = SystemNavigator Then
                    If rs!BuySell = True Then
                        If rs!RuleType = 0 Then
                            strItemTypeCat = "Long Entry"
                        Else
                            strItemTypeCat = "Short Exit"
                        End If
                    Else
                        If rs!RuleType = 0 Then
                            strItemTypeCat = "Short Entry"
                        Else
                            strItemTypeCat = "Long Exit"
                        End If
                    End If
                    AddRow rs!Name, "Shared Rule" & vbTab & rs!SystemNumber, CDbl(rs!LastModified), strItemTypeCat, _
                        DecryptField(rs!PreviewRTF), rs!RuleID, _
                        rs!SecurityLevel, DecryptField(rs!Password), rs!CannotDelete, rs!Reverify
                Else
                    'Portfolio Navigator tblRules
                    AddRow rs!RuleName, "Rule" & vbTab & "0", CDbl(rs!LastModified), rs!RuleClass, _
                        DecryptField(rs!Preview), rs!RuleID, _
                        rs!SecurityLevel, DecryptField(rs!Password), rs!CannotDelete, rs!Reverify
                End If
            End If
            
            rs.MoveNext
        Loop
        
        'Load Strategy Baskets...
        If g.CalledFrom = SystemNavigator Then
            Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblStrategyBaskets] WHERE [LibraryID]=" & Str(m.lLibraryID) & ";", dbOpenSnapshot)
            Do Until rs.EOF
                If rs!CheckSum = BuildCheckSum(rs, "tblStrategyBaskets") Then
                    AddRow rs!Name, "Basket" & vbTab & "0", CDbl(rs!LastModified), "N/A", _
                        rs!Description, rs!StrategyBasketID, _
                        rs!SecurityLevel, DecryptField(rs!Password), rs!CannotDelete, False, _
                        rs!IsGuru, rs!RequiredMod
                End If
                
                rs.MoveNext
            Loop
        End If
        
        ''.Outline 0
        
        ShowLocalRules
        If .Rows > 1 Then
            .Col = GDCol(eGDCol_Name)
            .Sort = flexSortGenericAscending
            .Row = 1
            .RowSel = 1
            ItemPreview
        End If
    
        .AutoSize 0, .Cols - 1, False, 75
        .AutoSize GDCol(eGDCol_SecurityLevel), , , 150
        .Redraw = lRedraw
    End With
    
ErrExit:
    Set rs = Nothing
    Exit Sub

ErrSection:
    Screen.MousePointer = vbDefault
    RaiseError "frmLibrary.LoadGrid", eGDRaiseError_Raise, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddRow
'' Description: Add a row to the grid
'' Inputs:      Items to add to each column of the new row
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddRow(pName As Variant, pItemType As Variant, pLastMod As Variant, _
    pItemTypeCat As Variant, pPreview As Variant, pID As Variant, _
    pSecurityLevel As Variant, pPassword As Variant, pCannotDelete As Variant, _
    ByVal bReverify As Boolean, Optional ByVal bIsGuru As Boolean = False, Optional ByVal strModule As String = "")
On Error GoTo ErrSection:

    Dim lRedraw As Long                 ' Current state of the grid's redraw
    Dim lSystemNumber As Long           ' System number of a rule
    Dim strItemType As String           ' Type of the item being added
    Dim lIndex As Long                  ' Index into a for loop
    Dim lRow As Long                    ' Current row
    Dim strLine As String               ' Line to add to the grid
    Dim strPreview As String
    
    strItemType = Parse(Str(pItemType), vbTab, 1)
    lSystemNumber = CLng(Val(Parse(Str(pItemType), vbTab, 2)))
    
    strPreview = Replace(Replace(pPreview, vbCrLf, "||"), Chr(9), Chr(1))
    
    With vsItems
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        ''If lSystemNumber <> 0 Then
        ''    .RowOutlineLevel(lRow) = 1
        ''Else
        ''    .RowOutlineLevel(lRow) = 0
        ''End If
        ''.IsSubtotal(lRow) = True
        
        If lSystemNumber = 0 Then
            .Rows = .Rows + 1
            lRow = .Rows - 1
            .TextMatrix(lRow, GDCol(eGDCol_Name)) = pName
            .TextMatrix(lRow, GDCol(eGDCol_ItemType)) = strItemType
            If IsNull(pItemTypeCat) Then .TextMatrix(lRow, GDCol(eGDCol_ItemTypeCat)) = "" Else .TextMatrix(lRow, GDCol(eGDCol_ItemTypeCat)) = pItemTypeCat
            .TextMatrix(lRow, GDCol(eGDCol_LastModified)) = Str(pLastMod) 'DateFormat(pLastMod) & " " & Format(pLastMod, "hh:mm:ss AM/PM")
            .TextMatrix(lRow, GDCol(eGDCol_Preview)) = strPreview
            .TextMatrix(lRow, GDCol(eGDCol_ID)) = Str(pID)
            .TextMatrix(lRow, GDCol(eGDCol_SecurityLevel)) = Str(pSecurityLevel)
            .TextMatrix(lRow, GDCol(eGDCol_Password)) = NullChk(pPassword)
            .TextMatrix(lRow, GDCol(eGDCol_CannotDelete)) = Str(pCannotDelete)
            .TextMatrix(lRow, GDCol(eGDCol_SecurityLevel)) = SecurityDesc(pSecurityLevel)
            .TextMatrix(lRow, GDCol(eGDCol_SystemNumber)) = Str(lSystemNumber)
            CheckedCell(vsItems, lRow, GDCol(eGDCol_Reverify)) = bReverify
            
            If (UCase(strItemType) = "STRATEGY") Or (UCase(strItemType) = "BASKET") Then
                CheckedCell(vsItems, lRow, GDCol(eGDCol_Shadow)) = bIsGuru
                .Cell(flexcpPictureAlignment, lRow, GDCol(eGDCol_Shadow)) = flexAlignCenterTop
                .TextMatrix(lRow, GDCol(eGDCol_RequiredMod)) = strModule
            End If
            
        Else
            lRow = .Rows - 1
            AddLocalRule lRow, GDCol(eGDCol_Name), "    " & pName
            AddLocalRule lRow, GDCol(eGDCol_ItemType), strItemType
            AddLocalRule lRow, GDCol(eGDCol_ItemTypeCat), NullChk(pItemTypeCat)
            ''AddLocalRule lRow, GDCol(eGDCol_LastModified), Str(pLastMod) 'DateFormat(pLastMod) & " " & Format(pLastMod, "hh:mm:ss AM/PM")
            ''AddLocalRule lRow, GDCol(eGDCol_Preview), pPreview
            AddLocalRule lRow, GDCol(eGDCol_ID), Str(pID)
            ''AddLocalRule lRow, GDCol(eGDCol_SecurityLevel), pSecurityLevel
            ''AddLocalRule lRow, GDCol(eGDCol_Password), NullChk(pPassword)
            ''AddLocalRule lRow, GDCol(eGDCol_CannotDelete), pCannotDelete
            ''AddLocalRule lRow, GDCol(eGDCol_SecurityLevel), SecurityDesc(pSecurityLevel)
            AddLocalRule lRow, GDCol(eGDCol_SystemNumber), Str(lSystemNumber)
            ''.RowHeight(lRow) = .RowHeight(lRow) + m.lRowHeight
        End If
        
        AddItemToArray lRow
        
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.AddRow", eGDRaiseError_Raise, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ItemPreview
'' Description: Show the preview for the selected item in the RTF box
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ItemPreview()
On Error GoTo ErrSection:
    
    Dim lLength As Long
    Dim Rule As Object
    Dim strPreview As String
    
    txtPreview.Text = ""
    Set Rule = CreateObject(g.strCommonDLL & "cRule")
    
    strPreview = vsItems.TextMatrix(vsItems.Row, GDCol(eGDCol_Preview))
    strPreview = Replace(Replace(strPreview, "||", vbCrLf), Chr(1), Chr(9))
        
    'Continue if NOT Rich text
    If InStr(vsItems.TextMatrix(vsItems.Row, GDCol(eGDCol_ItemType)), "Rule") Then
        txtPreview.TextRTF = Rule.GetRTF(strPreview)
    Else
        With txtPreview
            If Len(strPreview) = 0 Then
                .Text = "No Description"
                .SelStart = 0
                .SelLength = Len(.Text)
                .SelColor = vbBlack
                .SelBold = False
                .SelItalic = False
                .SelLength = 0
            Else
                .Text = strPreview
                lLength = InStr(.Text, Chr(13)) - InStr(.Text, ": ") - 2
                If lLength > 0 Then
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .SelColor = vbBlack
                    .SelItalic = False
                    .SelStart = InStr(.Text, ": ") + 1
                    .SelLength = InStr(.Text, Chr(13)) - InStr(.Text, ": ") - 2
                    If .SelLength > 0 Then
                        .SelBold = True
                    Else
                        .SelBold = False
                    End If
                    .SelLength = 0
                Else
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .SelColor = vbBlack
                    .SelBold = False
                    .SelItalic = False
                    .SelLength = 0
                End If
            End If
        End With
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmLibrary.ItemPreview", eGDRaiseError_Raise, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SaveItems
'' Description: Save the items in the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveItems()
On Error GoTo ErrSection:
    
    Dim rs As Recordset                 ' Recordset into the database
    Dim lIndex As Long                  ' Index into a for loop
    Dim strTable As String              ' Table in the database for the item type
    Dim strID As String                 ' ID Field in the table
    Dim astrItemType As New cGdArray    ' Item Type array for the current row
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim bBuildCheckSums As Boolean      ' Rebuild the checksums?
          
    ' First move all of the Removed Items into the User Library
    RemoveItems
    
    ' Second, reassign permissions to each item currently in grid
    With vsItems
        Set astrItemType = New cGdArray
        astrItemType.Create eGDARRAY_Strings
        For lIndex = .FixedRows To .Rows - 1
            astrItemType.SplitFields .TextMatrix(lIndex, GDCol(eGDCol_ItemType)), vbLf
            For lIndex2 = 0 To astrItemType.Size - 1
                Select Case UCase(astrItemType(lIndex2))
                    Case "STRATEGY"
                        strTable = "tblSystems"
                        strID = "SystemNumber"
                        bBuildCheckSums = True
                    Case "RULE", "LOCAL RULE", "SHARED RULE"
                        strTable = "tblRules"
                        strID = "RuleID"
                        bBuildCheckSums = True
                    Case "FUNCTION"
                        strTable = "tblFunctions"
                        strID = "FunctionID"
                        bBuildCheckSums = True
                    Case "PORTFOLIO"
                        strTable = "tblPortfolios"
                        strID = "PortfolioNumber"
                        bBuildCheckSums = False
                    Case "MODEL"
                        strTable = "tblModels"
                        strID = "ModelNumber"
                        bBuildCheckSums = False
                    Case "BASKET"
                        strTable = "tblStrategyBaskets"
                        strID = "StrategyBasketID"
                        bBuildCheckSums = True
                
                End Select
            
                Set rs = g.dbNav.OpenRecordset("SELECT * FROM [" & strTable & "] " & _
                        "WHERE [" & strID & "] = " & Parse(.TextMatrix(lIndex, GDCol(eGDCol_ID)), vbCrLf, lIndex2 + 1) & ";", dbOpenDynaset)
                If Not rs.EOF Then
                    rs.Edit
                    
                    'Default the developer name
                    If UCase(astrItemType(lIndex2)) = "STRATEGY" Then
                        rs!Developer = m.Library.Author
                        rs!IsGuru = CheckedCell(vsItems, lIndex, GDCol(eGDCol_Shadow))
                    ElseIf UCase(astrItemType(lIndex2)) = "BASKET" Then
                        rs!IsGuru = CheckedCell(vsItems, lIndex, GDCol(eGDCol_Shadow))
                        rs!RequiredMod = .TextMatrix(lIndex, GDCol(eGDCol_RequiredMod))
                    End If
                    
                    If m.Library.SecurityLevel = 0 Then
                        rs!SecurityLevel = 0
                        rs!Password = ""
                    Else
                        rs!SecurityLevel = SecurityLevelID(.TextMatrix(lIndex, GDCol(eGDCol_SecurityLevel)))
                        rs!Password = .TextMatrix(lIndex, GDCol(eGDCol_Password))
                    End If
                    rs!CannotDelete = .TextMatrix(lIndex, GDCol(eGDCol_CannotDelete))
                    rs!LibraryID = m.lLibraryID
                    If bBuildCheckSums = True Then
                        EncryptField rs!Password, NullChk(rs!Password)
                        rs!CheckSum = BuildCheckSum(rs, strTable)
                    End If
                    rs.Update
                End If
            Next lIndex2
        Next lIndex
    End With
    
ErrExit:
    Set rs = Nothing
    Exit Sub

ErrSection:
    RaiseError "frmLibrary.SaveItems", eGDRaiseError_Raise, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SecurityLevelID
'' Description: Return the security level for the given text
'' Inputs:      Security Level text from the grid
'' Returns:     Security Level
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SecurityLevelID(ByVal strDescription As String) As Byte
On Error GoTo ErrSection:
    
    Select Case strDescription
        Case "Can Edit/Can View": SecurityLevelID = 0
        Case "No Edit/Can View": SecurityLevelID = 1
        Case "No Edit/No View": SecurityLevelID = 2
        Case "No Access": SecurityLevelID = 3
    End Select

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmLibrary.SecurityLevelID", eGDRaiseError_Raise, g.strAppPath
    Resume ErrExit

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveItem
'' Description: Remove the items from the library (set to User Library)
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RemoveItems()
On Error GoTo ErrSection:
    
    Dim rs As Recordset                 ' Recordset into the database
    Dim strTable As String              ' Table of the database to query
    Dim strID As String                 ' ID Field in the database table
    Dim strItemType As String           ' Item Type of the item to remove
    Dim strItemID As String             ' ID of the item to remove
    Dim lIndex As Long                  ' Index into a for loop
    
    For lIndex = 0 To m.astrRemoved.Size - 1
        strItemType = Parse(m.astrRemoved(lIndex), vbTab, 1)
        strItemID = Parse(m.astrRemoved(lIndex), vbTab, 2)
    
        Select Case UCase(strItemType)
            Case "STRATEGY"
                strTable = "tblSystems"
                strID = "SystemNumber"
            Case "RULE", "LOCAL RULE", "SHARED RULE"
                strTable = "tblRules"
                strID = "RuleID"
            Case "FUNCTION"
                strTable = "tblFunctions"
                strID = "FunctionID"
            Case "PORTFOLIO"
                strTable = "tblPortfolios"
                strID = "PortfolioNumber"
            Case "MODEL"
                strTable = "tblModels"
                strID = "ModelNumber"
            Case "BASKET"
                strTable = "tblStrategyBaskets"
                strID = "StrategyBasketID"
        End Select
    
        Set rs = g.dbNav.OpenRecordset("SELECT * FROM [" & strTable & "] " & _
                "WHERE [" & strID & "] = " & strItemID & ";", dbOpenDynaset)
        Do Until rs.EOF
            rs.Edit
            rs!LibraryID = kUserLibrary
            If strTable = "tblSystems" Or strTable = "tblRules" Or strTable = "tblFunctions" Or strTable = "tblStrategyBaskets" Then
                rs!SecurityLevel = 0
                rs!CannotDelete = False
                rs!Password = ""
                rs!CheckSum = BuildCheckSum(rs, strTable)
            End If
            rs.Update
            
            rs.MoveNext
        Loop
    Next lIndex
    
ErrExit:
    Set rs = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.RemoveItems", eGDRaiseError_Raise, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddItemToArray
'' Description: Add an item to the Items array
'' Inputs:      Row in the grid to add
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddItemToArray(ByVal lRow As Long)
On Error GoTo ErrSection:

    Dim strSearch As String             ' Unique key to search for in the array
    Dim lPos As Long                    ' Position in the array
    Dim astrItemTypes As New cGdArray   ' Item Types in this cell
    Dim astrIDs As New cGdArray         ' IDs in this cell
    Dim lIndex As Long                  ' Index into a for loop

    astrItemTypes.Create eGDARRAY_Strings
    astrIDs.Create eGDARRAY_Strings

    astrItemTypes.SplitFields vsItems.TextMatrix(lRow, GDCol(eGDCol_ItemType)), vbLf
    astrIDs.SplitFields vsItems.TextMatrix(lRow, GDCol(eGDCol_ID)), vbLf

    For lIndex = 0 To astrItemTypes.Size - 1
        strSearch = astrItemTypes(lIndex) & vbTab & astrIDs(lIndex)
        If Not m.astrItems.BinarySearch(strSearch, lPos) Then
            m.astrItems.Add strSearch, lPos
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.AddItemToArray", eGDRaiseError_Raise, g.strAppPath
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetBackColors
'' Description: Set the background color of the rows appropriately
'' Inputs:      Grid that is currently active
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetBackColors()
On Error GoTo ErrSection:

    Dim bAlt As Boolean                 ' Is this an alternate row?
    Dim lRow As Long                    ' Index into a for loop
    Dim lRedraw As Long                 ' Current state of the redraw
    
    With vsItems
        lRedraw = .Redraw
        .Redraw = flexRDNone
        For lRow = .FixedRows To .Rows - 1
            If .RowHidden(lRow) = False Then
                If Not bAlt Then
                    .Cell(flexcpBackColor, lRow, 0, lRow, .Cols - 1) = .BackColor
                Else
                    .Cell(flexcpBackColor, lRow, 0, lRow, .Cols - 1) = .BackColorAlternate
                End If
                bAlt = Not bAlt
            End If
        Next lRow
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmLibrary.SetBackColors", eGDRaiseError_Raise, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Initialize and show the form
'' Inputs:      Library ID (None for Wizard)
'' Returns:     True if Saved, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(Optional ByVal lLibraryID As Long = 0) As Boolean
On Error GoTo ErrSection:

    ' Hide these for now since they don't apply to libraries...
    tbToolbar.Tools("ID_SaveAs").Visible = False
    tbToolbar.Tools("ID_Rename").Visible = False
    tbToolbar.Tools("ID_Toolbox").Visible = False

    If lLibraryID = 0 Then
        Add
    Else
        LoadRec lLibraryID
    End If
    
    If m.nMode <> eGDLibMode_WizardMode Then
        ScaleWidth = tabLibrary.Width + (tabLibrary.Left * 2)
        ScaleHeight = tabLibrary.Height + (tabLibrary.Top * 2)
    Else
        ScaleWidth = tabLibrary.Width + (tabLibrary.Left * 2)
        ScaleHeight = fraWizardToolbar.Height + tabLibrary.Height + (tabLibrary.Top * 3)
    End If
    CenterTheForm Me
    
    m.bSaved = False
    ShowForm Me, eForm_Modal, g.frmOwner
    ShowMe = m.bSaved Or g.bReload
    
ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmLibrary.ShowMe", eGDRaiseError_Raise, g.strAppPath

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddLocalRule
'' Description: Add a local rule to the same row as the appropriate system
'' Inputs:      Row and Column to add to, Item to add
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddLocalRule(ByVal lRow As Long, ByVal lCol As Long, ByVal strItem As String)
On Error GoTo ErrSection:

    With vsItems
        .TextMatrix(lRow, lCol) = .TextMatrix(lRow, lCol) & vbCrLf & strItem
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.AddLocalRule", eGDRaiseError_Raise, g.strAppPath
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowLocalRules
'' Description: Show/Hide the local rules as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ShowLocalRules()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lRedraw As Long                 ' Current state of the grid's redraw

    With vsItems
        lRedraw = .Redraw
        .Redraw = flexRDNone
        For lIndex = .FixedRows To .Rows - 1
            If chkShowLocal = vbChecked Then
                .RowHeight(lIndex) = RowHeight(Me, .Cell(flexcpFont, lIndex, GDCol(eGDCol_Name)), .TextMatrix(lIndex, GDCol(eGDCol_Name))) + 50
                If InStr(.TextMatrix(lIndex, GDCol(eGDCol_Name)), vbCrLf) Then .RowHeight(lIndex) = .RowHeight(lIndex) + 50
            Else
                .RowHeight(lIndex) = .RowHeight(0)
            End If
            
            If CheckedCell(vsItems, lIndex, GDCol(eGDCol_Reverify)) Then
                .Cell(flexcpForeColor, lIndex, GDCol(eGDCol_Name)) = vbRed
            Else
                .Cell(flexcpForeColor, lIndex, GDCol(eGDCol_Name)) = .Cell(flexcpForeColor, 0, 0)
            End If
        Next lIndex
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.ShowLocalRules", eGDRaiseError_Raise, g.strAppPath
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitFilesGrid
'' Description: Initialize the files grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitFilesGrid()
On Error GoTo ErrSection:

    Dim lRedraw As Long                 ' Current state of the grid's redraw

    With fgFiles
        lRedraw = .Redraw
        .Redraw = flexRDNone

        .Clear
        .Editable = flexEDKbdMouse
        .ExplorerBar = flexExSortShow
        .ExtendLastCol = True
        .SelectionMode = flexSelectionListBox
        .AllowUserResizing = flexResizeColumns
        .AllowSelection = True
        .AutoSearch = flexSearchFromTop
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .ScrollBars = flexScrollBarBoth
        .ScrollTrack = True
        .FixedCols = 0
        .FixedRows = 1
        .Rows = 1
        .Cols = kFileGridCols
        
        .TextMatrix(0, FGCol(eFGCol_FileName)) = "File Name"
        .TextMatrix(0, FGCol(eFGCol_FileDate)) = "File Date"
        .TextMatrix(0, FGCol(eFGCol_FileSize)) = "File Size"
        .TextMatrix(0, FGCol(eFGCol_OnlyNewer)) = "Only If Newer"
        .TextMatrix(0, FGCol(eFGCol_ReadOnly)) = "Read Only"
        
        .ColHidden(FGCol(eFGCol_Flags)) = True
        .ColHidden(FGCol(eFGCol_FileInfo)) = True
        
        .ColDataType(FGCol(eFGCol_OnlyNewer)) = flexDTBoolean
        .ColDataType(FGCol(eFGCol_ReadOnly)) = flexDTBoolean
        
        .ColFormat(FGCol(eFGCol_FileDate)) = DateFormat("format") & " HH:MM AM/PM"
        .ColFormat(FGCol(eFGCol_FileSize)) = "#,##0"
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignLeftTop
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.InitFilesGrid", eGDRaiseError_Raise, g.strAppPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadFilesGrid
'' Description: Load the files grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadFilesGrid()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset out of the database
    Dim lRedraw As Long                 ' Current state of the grid's redraw
    Dim lRow As Long                    ' Current row in the grid
    
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblFiles] WHERE [LibraryID]=" & Str(m.lLibraryID) & ";", dbOpenDynaset)
    Do While Not rs.EOF
        With fgFiles
            lRedraw = .Redraw
            .Redraw = flexRDNone
            
            .Rows = .Rows + 1
            lRow = .Rows - 1
            
            .TextMatrix(lRow, FGCol(eFGCol_FileName)) = rs!FileName
            .TextMatrix(lRow, FGCol(eFGCol_FileSize)) = rs!FileSize
            .TextMatrix(lRow, FGCol(eFGCol_FileDate)) = rs!FileDate
            .TextMatrix(lRow, FGCol(eFGCol_Flags)) = rs!Flags
            .TextMatrix(lRow, FGCol(eFGCol_FileInfo)) = rs!FileInfo
            
            CheckedCell(fgFiles, lRow, FGCol(eFGCol_OnlyNewer)) = (rs!Flags And 1)
            CheckedCell(fgFiles, lRow, FGCol(eFGCol_ReadOnly)) = (rs!Flags And 2)
                
            .AutoSize 0, .Cols - 1, False, 75
            .Redraw = lRedraw
        End With
        rs.MoveNext
    Loop

ErrExit:
    Set rs = Nothing
    Exit Sub
    
ErrSection:
    Set rs = Nothing
    RaiseError "frmLibrary.LoadFilesGrid", eGDRaiseError_Raise, g.strAppPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SaveFiles
'' Description: Save the file information to the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveFiles()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim lIndex As Long                  ' Index into a for loop
    Dim bFound As Boolean               ' Was the file found in the grid?
    Dim bRemoved As Boolean             ' Have any files been removed from the library?
    Dim strFileName As String           ' Path and filename for the file to save

    ' Get the file information out of the database for the Library ID
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblFiles] " & _
                "WHERE [LibraryID]=" & Str(m.lLibraryID) & " " & _
                "ORDER BY [FileName];", dbOpenDynaset)
    
    With fgFiles
        ' Delete all of the file information for removed files
        Do While Not rs.EOF
            bFound = False
            For lIndex = .FixedRows To .Rows - 1
                If rs!FileName = .TextMatrix(lIndex, FGCol(eFGCol_FileName)) Then
                    bFound = True
                    Exit For
                End If
            Next lIndex
            
            If Not bFound Then rs.Delete
            rs.MoveNext
        Loop
    
        ' Insert all of the file information from the grid
        bRemoved = False
        If Not rs.EOF Then rs.MoveFirst
        For lIndex = .Rows - 1 To .FixedRows Step -1
            strFileName = AddSlash(g.strAppPath) & .TextMatrix(lIndex, FGCol(eFGCol_FileName))
            
            If FileExist(strFileName) Then
                rs.FindFirst "[FileName]='" & .TextMatrix(lIndex, FGCol(eFGCol_FileName)) & "'"
                If rs.NoMatch Then
                    rs.AddNew
                Else
                    rs.Edit
                End If
                rs!LibraryID = m.lLibraryID
                rs!FileName = .TextMatrix(lIndex, FGCol(eFGCol_FileName))
                rs!FileSize = mGenesis.FileLength(strFileName)
                rs!FileDate = mGenesis.FileDate(strFileName)
                rs!Flags = CLng(ValOfText(.TextMatrix(lIndex, FGCol(eFGCol_Flags))))
                rs!FileInfo = .TextMatrix(lIndex, FGCol(eFGCol_FileInfo))
                rs.Update
            Else
                bRemoved = True
                rs.FindFirst "[FileName]='" & .TextMatrix(lIndex, FGCol(eFGCol_FileName)) & "'"
                If rs.NoMatch = False Then
                    rs.Delete
                End If
                .RemoveItem lIndex
            End If
        Next lIndex
    End With
    
    If bRemoved Then
        InfBox "One or more files have been removed from your library because they no longer exist", "i", , "Library Save Warning"
    End If

ErrExit:
    Set rs = Nothing
    Exit Sub
    
ErrSection:
    Set rs = Nothing
    RaiseError "frmLibrary.SaveFiles", eGDRaiseError_Raise, g.strAppPath
    
End Sub

Private Sub AddItem()
On Error GoTo ErrSection:

    Dim astrItems As New cGdArray       ' Items added from the add form
    Dim lRedraw As Long                 ' Current state of the grid's redraw
    Dim lIndex As Long                  ' Index into a for loop
    Dim strItemType As String           ' Item type for the item
    
    Set astrItems = frmLibraryAddItem.ShowMe(eGDLibAddItemsMode_All, 0, "", m.astrItems)
    If astrItems.Size > 0 Then
        With vsItems
            lRedraw = .Redraw
            .Redraw = flexRDNone
            
            For lIndex = 0 To astrItems.Size - 1
                .AddItem astrItems(lIndex)
                
                ' Set the default security stuff
                If LibAccess(m.nMode * 2).Value = True Then
                    .TextMatrix(.Rows - 1, GDCol(eGDCol_SecurityLevel)) = SecurityDesc(0)
                    .TextMatrix(.Rows - 1, GDCol(eGDCol_Password)) = ""
                Else
                    .TextMatrix(.Rows - 1, GDCol(eGDCol_SecurityLevel)) = SecurityDesc(1)
                    .TextMatrix(.Rows - 1, GDCol(eGDCol_Password)) = txtPassword(m.nMode).Text
                End If
                
                strItemType = .TextMatrix(.Rows - 1, GDCol(eGDCol_ItemType))
                If (UCase(strItemType) = "STRATEGY") Or (UCase(strItemType) = "BASKET") Then
                    CheckedCell(vsItems, .Rows - 1, GDCol(eGDCol_Shadow)) = False
                    .Cell(flexcpPictureAlignment, .Rows - 1, GDCol(eGDCol_Shadow)) = flexAlignCenterTop
                    .TextMatrix(.Rows - 1, GDCol(eGDCol_RequiredMod)) = ""
                End If
                
                AddItemToArray .Rows - 1
            Next lIndex
            
            ShowLocalRules
            
            .AutoSize 0, .Cols - 1, False, 75
            .AutoSize GDCol(eGDCol_SecurityLevel), , , 150
            .Redraw = lRedraw
            
            EnableSave True
        End With
    End If

ErrExit:
    Set astrItems = Nothing
    Exit Sub
    
ErrSection:
    Set astrItems = Nothing
    RaiseError "frmLibrary.AddItem", eGDRaiseError_Raise, g.strAppPath
    
End Sub

Private Sub RemoveItem()
On Error GoTo ErrSection:

    Dim lRedraw As Long                 ' Current state of the grid's redraw
    Dim lIndex As Long                  ' Index into a for loop
    Dim lID As Long                     ' ID of the item being deleted
    Dim strRemoved As String            ' String to add to the removed array
    Dim lPos As Long                    ' Position in the items array
    Dim astrItemTypes As New cGdArray   ' Item Types in this cell
    Dim astrIDs As New cGdArray         ' IDs in this cell
    Dim lRow As Long                    ' Index into a for loop

    ' Create the arrays
    astrItemTypes.Create eGDARRAY_Strings
    astrIDs.Create eGDARRAY_Strings

    With vsItems
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        ' Walk through the selected rows
        For lRow = .SelectedRows - 1 To 0 Step -1
            ' Make sure that the row is valid
            If .SelectedRow(lRow) >= .FixedRows And .SelectedRow(lRow) < .Rows Then
                ' Split the lines of the row into arrays
                astrItemTypes.SplitFields .TextMatrix(.SelectedRow(lRow), GDCol(eGDCol_ItemType)), vbLf
                astrIDs.SplitFields .TextMatrix(.SelectedRow(lRow), GDCol(eGDCol_ID)), vbLf
                
                ' Remove items from the Items array and add to the Removed array
                For lIndex = 0 To astrItemTypes.Size - 1
                    strRemoved = astrItemTypes(lIndex) & vbTab & astrIDs(lIndex)
                    If m.astrRemoved.BinarySearch(strRemoved, lPos) = False Then
                        m.astrRemoved.Add strRemoved, lPos
                    End If
                    If m.astrItems.BinarySearch(strRemoved, lPos) = True Then
                        m.astrItems.Remove lPos
                    End If
                Next lIndex
                
                ' Remove the item from the grid
                .RemoveItem .SelectedRow(lRow)
            End If
        Next lRow
        
        ' Select the first row
        If .Rows > .FixedRows Then
            .Row = .FixedRows
            .RowSel = .FixedRows
        End If
        
        .Redraw = lRedraw
    End With
    
    EnableSave True

ErrExit:
    Set astrItemTypes = Nothing
    Set astrIDs = Nothing
    Exit Sub
    
ErrSection:
    Set astrItemTypes = Nothing
    Set astrIDs = Nothing
    RaiseError "frmLibrary.RemoveItem", eGDRaiseError_Raise, g.strAppPath
    
End Sub

Private Sub AddFile()
On Error GoTo ErrSection:

    Dim strFiles As String              ' List of files to add to the grid
    Dim lRedraw As Long                 ' Current state of the grid's redraw
    Dim astrFiles As cGdArray           ' List of files that were selected
    Dim lIndex As Long                  ' Index into a for loop
    Dim lRow As Long                    ' Current row in the grid
    Dim dFileDate As Double             ' Date of the current file
    Dim strFileName As String           ' Filename to add
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim strDisplay As String            ' Filename to display

    strFiles = mGenesis.CommonDialogFile(Me.CommonDialog1, False, , g.strAppPath, "Select File(s)", , ",")
    If strFiles <> "" Then
        With fgFiles
            lRedraw = .Redraw
            .Redraw = flexRDNone
            
            Set astrFiles = New cGdArray
            astrFiles.Create eGDARRAY_Strings
            astrFiles.SplitFields strFiles, ","
            For lIndex = 0 To astrFiles.Size - 1
                If InStr(astrFiles(lIndex), "'") <> 0 Then
                    strFileName = StripStr(astrFiles(lIndex), Chr(34))
                    strFileName = Replace(strFileName, AddSlash(g.strAppPath), "")
                    InfBox "The following file cannot be included because it has an apostrophe in the file name.  Please rename it if you wish to include it in this library.||" & strFileName & "|", "!", , "Include File Error"
                ElseIf Not IsCustomFile(astrFiles(lIndex)) Then
                    .Rows = .Rows + 1
                    lRow = .Rows - 1
                    strFileName = StripStr(astrFiles(lIndex), Chr(34))
                    
                    strDisplay = Replace(strFileName, AddSlash(g.strAppPath), "")
                    For lIndex2 = .FixedRows To .Rows - 1
                        If .TextMatrix(lIndex2, FGCol(eFGCol_FileName)) = strDisplay Then
                            lRow = lIndex2
                            Exit For
                        End If
                    Next lIndex2
                    
                    .TextMatrix(lRow, FGCol(eFGCol_FileName)) = strDisplay
                    .TextMatrix(lRow, FGCol(eFGCol_FileSize)) = FileLength(strFileName) ' F.Size
                    .TextMatrix(lRow, FGCol(eFGCol_FileDate)) = FileDate(strFileName) ' F.DateLastModified
                Else
                    strFileName = StripStr(astrFiles(lIndex), Chr(34))
                    strFileName = Replace(strFileName, AddSlash(g.strAppPath), "")
                    InfBox "The following file cannot be included because it has a default file name.  Please rename the file if you wish to include it in this library.||" & strFileName & "|", "!", , "Include File Error"
                End If
            Next lIndex
            
            .AutoSize 0, .Cols - 1, False, 75
            .Redraw = lRedraw
            
            EnableSave True
        End With
    End If

ErrExit:
    Set astrFiles = Nothing
    Exit Sub
    
ErrSection:
    Set astrFiles = Nothing
    RaiseError "frmLibrary.AddFile", eGDRaiseError_Raise, g.strAppPath
    
End Sub

Private Sub RemoveFile()
On Error GoTo ErrSection:

    Dim lRedraw As Long                 ' Current state of the grid's redraw
    Dim lIndex As Long                  ' Index into a for loop

    With fgFiles
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        For lIndex = .SelectedRows - 1 To 0 Step -1
            .RemoveItem .SelectedRow(lIndex)
        Next lIndex
        
        .Redraw = lRedraw
        EnableSave True
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.RemoveFile", eGDRaiseError_Raise, g.strAppPath
    
End Sub

Private Function IsCustomFile(ByVal strFileName As String) As Boolean
On Error GoTo ErrSection:

    Dim strFileBase As String           ' Base of the filename passed in
    
    IsCustomFile = False
    strFileBase = FileBase(strFileName)
    If Left(strFileBase, 3) = "CUS" Then
        If IsNumeric(Right(strFileBase, 5)) = True Then
            IsCustomFile = True
        End If
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmLibrary.IsCustomFile", eGDRaiseError_Raise, g.strAppPath
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMeCreate
'' Description: Create a library from a name and a list of items
'' Inputs:      Name, Author, Description, Permissions, Password, Items, ID
'' Returns:     True if Saved, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMeCreate(ByVal strName As String, ByVal strAuthor As String, ByVal strDescription As String, ByVal bFull As Boolean, ByVal strPassword As String, ByVal strItems As String, lLibraryID As Long) As Boolean
On Error GoTo ErrSection:

    ' Hide these for now since they don't apply to libraries...
    tbToolbar.Tools("ID_SaveAs").Visible = False
    tbToolbar.Tools("ID_Rename").Visible = False
    tbToolbar.Tools("ID_Toolbox").Visible = False
    
    ScaleWidth = tabLibrary.Width + (tabLibrary.Left * 2)
    ScaleHeight = tabLibrary.Height + (tabLibrary.Top * 2)
    CenterTheForm Me
    
    Set m.Library = New cLibrary
    m.lLibraryID = 0
    
    With fgFiles
        .Redraw = flexRDNone
        InitFilesGrid
        .Redraw = flexRDBuffered
    End With
    
    m.nMode = eGDLibMode_NormalMode
    With tabLibrary
        .TabVisible(Tabs(eGDTab_Description)) = False
        .TabVisible(Tabs(eGDTab_Author)) = False
        .TabVisible(Tabs(eGDTab_LibPermissions)) = False
        .TabVisible(Tabs(eGDTab_ItemPermissions)) = False
    End With

    strName = StripStr(strName, ":\/*?|><" & Chr(34))
    SetEditorCaption Me, "Library", strName
    txtName(m.nMode).Text = strName
    txtDescription(m.nMode).Text = strDescription
    udVersion.Min = 1&
    udVersion.Max = 200000000
    txtVersion(m.nMode).Text = "1"
    txtVersion(m.nMode).Locked = True
    txtAuthor(m.nMode).Text = strAuthor
    txtPhoneNumber(m.nMode).Text = ""
    txtEMail(m.nMode).Text = ""
    txtWebSite(m.nMode).Text = ""
    txtDLLName(m.nMode).Text = ""
    txtSecurityCode(m.nMode).Text = ""
    chkDLLRequired(m.nMode).Value = vbUnchecked
    txtDLLName(m.nMode).Enabled = False
    lblDLLName(m.nMode).Enabled = False
    txtSecurityCode(m.nMode).Enabled = False
    lblSecurityCode(m.nMode).Enabled = False
    txtFileName(m.nMode).Text = ""
    If bFull Then
        LibAccess(m.nMode * 2).Value = 1
        LibAccess((m.nMode * 2) + 1).Value = 0
    Else
        LibAccess(m.nMode * 2).Value = 0
        LibAccess((m.nMode * 2) + 1).Value = 1
    End If
    txtPassword(m.nMode).Text = strPassword
    chkCannotDelete(m.nMode).Value = vbUnchecked
    txtRequiredMod.Text = ""
    
    tabLibrary.CurrTab = Tabs(eGDTab_Items)
    tbToolbar.Visible = True
    fraWizardToolbar.Visible = False
    
    EnableSave True
    tabLibrary.Align = asNone
    tabLibrary.TabStop = True
    
    m.nDefaultItemSecurity = 0
    
    vsItems.Redraw = flexRDNone
    InitGrid
    AddPassedItems strItems
    vsItems.Redraw = flexRDBuffered
    
    Save "ID_Save", False
    
    If tbToolbar.Tools("ID_Save").Enabled = True Then
        ShowForm Me, eForm_Modal, g.frmOwner
    End If
    
    lLibraryID = m.lLibraryID
    ShowMeCreate = m.bSaved Or g.bReload

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmLibrary.ShowMeCreate", eGDRaiseError_Raise, g.strAppPath
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddPassedItems
'' Description: Add the items that were passed in to the items grid
'' Inputs:      String of Items
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddPassedItems(ByVal strItems As String)
On Error GoTo ErrSection:

    Dim lRedraw As Long                 ' Current state of the grid's redraw
    Dim astrItems As New cGdArray       ' Array of items to add to the grid
    Dim strType As String               ' Type of item
    Dim strID As String                 ' ID of the item
    Dim lIndex As Long                  ' Index into a for loop
    Dim rs As Recordset                 ' Recordset into the database
    Dim rs2 As Recordset                ' Recordset into the database
    Dim strPreview As String            ' Preview for the item
    Dim strItemTypeCat As String        ' Item type category
    Dim astrDepends As New cGdArray     ' Array of dependencies
    Dim lPos As Long                    ' Position where the item should be
    Dim nSecurityLevel As Byte          ' Security level for the item
    Dim strPassword As String           ' Password for the item

    ' Initialize variables...
    astrDepends.Create eGDARRAY_Strings
    astrItems.Create eGDARRAY_Strings
    astrItems.SplitFields strItems, "|"
    astrItems.Sort eGdSort_DeleteDuplicates
    
    If LibAccess(m.nMode * 2).Value = True Then
        nSecurityLevel = 0
        strPassword = ""
    Else
        nSecurityLevel = 2
        strPassword = Trim(txtPassword(m.nMode).Text)
    End If
    
    ' Determine other user library dependencies...
    For lIndex = 0 To astrItems.Size - 1
        strType = Parse(astrItems(lIndex), ",", 1)
        strID = Parse(astrItems(lIndex), ",", 2)
        
        Select Case UCase(strType)
            Case "FUNCTION"
                FunctionDependencies astrDepends, Val(strID), eGDDependencyFilter_UserLibraryOnly
            
            Case "RULE"
                RuleDependencies astrDepends, Val(strID), eGDDependencyFilter_UserLibraryOnly
                
            Case "SYSTEM"
                SystemDependencies astrDepends, Val(strID), False, eGDDependencyFilter_UserLibraryOnly
                
            Case "BASKET"
                BasketDependencies astrDepends, Val(strID), False, eGDDependencyFilter_UserLibraryOnly
        
        End Select
    Next lIndex
    
    ' Add the user library dependencies to the items array...
    For lIndex = 0 To astrDepends.Size - 1
        strType = Parse(astrDepends(lIndex), vbTab, 3)
        strID = Parse(astrDepends(lIndex), vbTab, 1)
        
        If astrItems.BinarySearch(strType & "," & strID, lPos) = False Then
            astrItems.Add strType & "," & strID, lPos
        End If
    Next lIndex
    
    lRedraw = vsItems.Redraw
    vsItems.Redraw = flexRDNone
    
    For lIndex = 0 To astrItems.Size - 1
        strType = Parse(astrItems(lIndex), ",", 1)
        strID = Parse(astrItems(lIndex), ",", 2)
        
        Select Case UCase(strType)
            Case "FUNCTION"
                Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblFunctions] " & _
                            "WHERE [FunctionID]=" & strID & ";", dbOpenDynaset)
                If Not (rs.BOF And rs.EOF) Then
                    If rs!CheckSum = BuildCheckSum(rs, "tblFunctions") Then
                        strPreview = "Usage: " & rs!TradeSenseUsage & vbCrLf & _
                                 "Description: " & rs!Description
            
                        AddRow rs!FunctionName, "Function" & vbTab & "0", _
                            CDbl(rs!LastModified), ImplementationTypeDesc(rs!ImplementationTypeID), _
                            strPreview, rs!FunctionID, _
                            nSecurityLevel, strPassword, rs!CannotDelete, rs!Reverify
                    End If
                End If
                
            Case "RULE"
                Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblRules] " & _
                            "WHERE [RuleID]=" & strID & ";", dbOpenDynaset)
                If Not (rs.BOF And rs.EOF) Then
                    If rs!CheckSum = BuildCheckSum(rs, "tblRules") Then
                        If rs!BuySell = True Then
                            If rs!RuleType = 0 Then
                                strItemTypeCat = "Long Entry"
                            Else
                                strItemTypeCat = "Short Exit"
                            End If
                        Else
                            If rs!RuleType = 0 Then
                                strItemTypeCat = "Short Entry"
                            Else
                                strItemTypeCat = "Long Exit"
                            End If
                        End If
                        AddRow rs!Name, "Shared Rule" & vbTab & rs!SystemNumber, CDbl(rs!LastModified), strItemTypeCat, _
                            DecryptField(rs!PreviewRTF), rs!RuleID, _
                            nSecurityLevel, strPassword, rs!CannotDelete, rs!Reverify
                    End If
                End If
                
            Case "SYSTEM"
                Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblSystems] " & _
                            "WHERE [SystemNumber]=" & strID & ";", dbOpenDynaset)
                If Not (rs.BOF And rs.EOF) Then
                    If rs!CheckSum = BuildCheckSum(rs, "tblSystems") Then
                        AddRow rs!SystemName, "Strategy" & vbTab & "0", CDbl(rs!LastModified), "N/A", _
                            rs!Notes, rs!SystemNumber, _
                            nSecurityLevel, strPassword, rs!CannotDelete, rs!Reverify
                            
                        Set rs2 = g.dbNav.OpenRecordset("SELECT * FROM [tblRules] " & _
                            "WHERE [SystemNumber]=" & Str(rs!SystemNumber) & " " & _
                            "ORDER BY [Name];", dbOpenSnapshot)
                        If Not (rs2.BOF And rs2.EOF) Then rs2.MoveFirst
                        Do Until rs2.EOF
                            If rs2!CheckSum = BuildCheckSum(rs2, "tblRules") Then
                                If rs2!BuySell = True Then
                                    If rs2!RuleType = 0 Then
                                        strItemTypeCat = "Long Entry"
                                    Else
                                        strItemTypeCat = "Short Exit"
                                    End If
                                Else
                                    If rs2!RuleType = 0 Then
                                        strItemTypeCat = "Short Entry"
                                    Else
                                        strItemTypeCat = "Long Exit"
                                    End If
                                End If
                                AddRow rs2!Name, "Local Rule" & vbTab & rs2!SystemNumber, CDbl(rs2!LastModified), strItemTypeCat, _
                                    DecryptField(rs2!PreviewRTF), rs2!RuleID, _
                                    nSecurityLevel, strPassword, rs2!CannotDelete, rs2!Reverify
                            End If
                            
                            rs2.MoveNext
                        Loop
                    End If
                End If
                
            Case "BASKET"
                Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblStrategyBaskets] WHERE [StrategyBasketID]=" & strID & ";", dbOpenDynaset)
                If Not (rs.BOF And rs.EOF) Then
                    If rs!CheckSum = BuildCheckSum(rs, "tblStrategyBaskets") Then
                        AddRow rs!Name, "Basket" & vbTab & "0", CDbl(rs!LastModified), "N/A", _
                            rs!Description, rs!StrategyBasketID, _
                            nSecurityLevel, strPassword, rs!CannotDelete, rs!Reverify
                    End If
                End If
                
        End Select
    Next lIndex
    
    vsItems.AutoSize 0, vsItems.Cols - 1, False, 75
    vsItems.Redraw = lRedraw

ErrExit:
    Set astrItems = Nothing
    Exit Sub
    
ErrSection:
    Set astrItems = Nothing
    RaiseError "frmLibrary.AddPassedItems", eGDRaiseError_Raise, g.strAppPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ToggleShadowColumns
'' Description: Show/Hide the shadow columns in the grid as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ToggleShadowColumns()
On Error GoTo ErrSection:

    Dim bShowShadow As Boolean          ' Show the shadow columns?

    bShowShadow = ((g.bShowShadow = True) And (CheckBoxValue(chkShadow) = True))
    
    With vsItems
        .ColHidden(GDCol(eGDCol_Shadow)) = Not bShowShadow
        .ColHidden(GDCol(eGDCol_RequiredMod)) = Not bShowShadow
    End With
    
    If bShowShadow Then
        If Len(Trim(txtOwners.Text)) = 0 Then
            txtOwners.Text = Str(g.lLCD)
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.ToggleShadowColumns", , g.strAppPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MainTypeForRow
'' Description: Determine the main type for the given row
'' Inputs:      Row
'' Returns:     Main Type
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function MainTypeForRow(ByVal lRow As Long) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim astrItemType As cGdArray        ' Item Type array for the current row
    
    strReturn = ""
    With vsItems
        If (lRow >= .FixedRows) And (lRow < .Rows) Then
            Set astrItemType = New cGdArray
            astrItemType.SplitFields .TextMatrix(lRow, GDCol(eGDCol_ItemType)), vbLf
            
            If astrItemType.Size > 0 Then
                strReturn = astrItemType(0)
            End If
        End If
    End With
    
    MainTypeForRow = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmLibrary.MainTypeForRow", , g.strAppPath
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CloseForm
'' Description: Attempt to close the form if we can
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CloseForm()
On Error GoTo ErrSection:

    Dim strResult As String             ' Result back from the InfBox

    If tbToolbar.Tools("ID_Save").Enabled Then
        strResult = InfBox("Do you want to save the changes?", "?", "+Yes|No|-Cancel", "Confirmation")
        Select Case UCase(strResult)
            Case "Y"
                Save "ID_Save"
                Me.Hide
            Case "N"
                Me.Hide
        End Select
    Else
        Me.Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.CloseForm", , g.strAppPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleSaveError
'' Description: Handle an error that happened because of a Save
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HandleSaveError()
On Error GoTo ErrSection:

    If (m.Library.ErrNbr = 11) Or (m.Library.ErrNbr = 12) Then
        If m.nMode = eGDLibMode_WizardMode Then
            tabLibrary.CurrTab = Tabs(eGDTab_LibPermissions)
            LibAccess(1).Value = True
            MoveFocus txtPassword(0)
        Else
            tabLibrary.CurrTab = Tabs(eGDTab_LibPermissions2)
            LibAccess(3).Value = True
            MoveFocus txtPassword(1)
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibrary.HandleSaveError", , g.strAppPath
    
End Sub
