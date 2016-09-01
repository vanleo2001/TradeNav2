VERSION 5.00
Object = "{A8561640-E93C-11D3-AC3B-CE6078F7B616}#1.0#0"; "VSPRINT7.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.0#0"; "HexUniControls42.ocx"
Begin VB.Form frmPrintPreview 
   Caption         =   "To Printer"
   ClientHeight    =   6375
   ClientLeft      =   2070
   ClientTop       =   1605
   ClientWidth     =   9450
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPrintPreview.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6375
   ScaleWidth      =   9450
   ShowInTaskbar   =   0   'False
   Begin HexUniControls.ctlUniFrameWL fraPages 
      Height          =   1275
      Left            =   60
      TabIndex        =   23
      Top             =   2940
      Width           =   3165
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
      Caption         =   "frmPrintPreview.frx":014A
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmPrintPreview.frx":017A
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmPrintPreview.frx":019A
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdScale 
         Height          =   250
         Index           =   3
         Left            =   2475
         TabIndex        =   29
         Top             =   240
         Width           =   250
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
         Caption         =   "frmPrintPreview.frx":01B6
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmPrintPreview.frx":01DA
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmPrintPreview.frx":01FA
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdScale 
         Height          =   250
         Index           =   2
         Left            =   2220
         TabIndex        =   28
         Top             =   240
         Width           =   250
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
         Caption         =   "frmPrintPreview.frx":0216
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmPrintPreview.frx":0238
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmPrintPreview.frx":0258
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdScale 
         Height          =   250
         Index           =   1
         Left            =   1575
         TabIndex        =   26
         Top             =   240
         Width           =   250
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
         Caption         =   "frmPrintPreview.frx":0274
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmPrintPreview.frx":0296
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmPrintPreview.frx":02B6
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdScale 
         Height          =   250
         Index           =   0
         Left            =   1320
         TabIndex        =   25
         Top             =   240
         Width           =   250
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
         Caption         =   "frmPrintPreview.frx":02D2
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmPrintPreview.frx":02F6
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmPrintPreview.frx":0316
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniTextBoxXP txtNumCopies 
         Height          =   285
         Left            =   1800
         TabIndex        =   35
         Top             =   885
         Width           =   915
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmPrintPreview.frx":0332
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         Tip             =   "frmPrintPreview.frx":0352
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPrintPreview.frx":0372
      End
      Begin HexUniControls.ctlUniTextBoxXP txtToPage 
         Height          =   285
         Left            =   2220
         TabIndex        =   33
         Top             =   540
         Width           =   495
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmPrintPreview.frx":038E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         Tip             =   "frmPrintPreview.frx":03AE
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPrintPreview.frx":03CE
      End
      Begin HexUniControls.ctlUniTextBoxXP txtFromPage 
         Height          =   285
         Left            =   1320
         TabIndex        =   31
         Top             =   540
         Width           =   495
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmPrintPreview.frx":03EA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         Tip             =   "frmPrintPreview.frx":040A
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPrintPreview.frx":042A
      End
      Begin HexUniControls.ctlUniLabelXP lblScaleAmount 
         Height          =   165
         Left            =   1800
         Top             =   285
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
         Caption         =   "frmPrintPreview.frx":0446
         BackColor       =   14737632
         ForeColor       =   -2147483630
         Alignment       =   2
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmPrintPreview.frx":046C
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPrintPreview.frx":048C
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblScale 
         Height          =   195
         Left            =   150
         Top             =   270
         Width           =   435
         _ExtentX        =   794
         _ExtentY        =   344
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmPrintPreview.frx":04A8
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   1
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   -1  'True
         Tip             =   "frmPrintPreview.frx":04D4
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPrintPreview.frx":04F4
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label4 
         Height          =   255
         Left            =   240
         Top             =   900
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
         Caption         =   "frmPrintPreview.frx":0510
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmPrintPreview.frx":0552
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPrintPreview.frx":0572
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label3 
         Height          =   285
         Left            =   1980
         Top             =   540
         Width           =   135
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
         Caption         =   "frmPrintPreview.frx":058E
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmPrintPreview.frx":05B2
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPrintPreview.frx":05D2
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label2 
         Height          =   285
         Left            =   240
         Top             =   540
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
         Caption         =   "frmPrintPreview.frx":05EE
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmPrintPreview.frx":0626
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPrintPreview.frx":0646
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL Frame2 
      Height          =   1350
      Left            =   60
      TabIndex        =   36
      Top             =   4320
      Width           =   3165
      _ExtentX        =   5583
      _ExtentY        =   2381
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmPrintPreview.frx":0662
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmPrintPreview.frx":0698
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmPrintPreview.frx":06B8
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniRadioXP optClipboard 
         Height          =   220
         Left            =   1860
         TabIndex        =   39
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
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
         Caption         =   "frmPrintPreview.frx":06D4
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmPrintPreview.frx":0706
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmPrintPreview.frx":0726
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optFile 
         Height          =   220
         Left            =   1140
         TabIndex        =   38
         Top             =   240
         Width           =   675
         _ExtentX        =   1191
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
         Caption         =   "frmPrintPreview.frx":0742
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmPrintPreview.frx":076A
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmPrintPreview.frx":078A
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optPrinter 
         Height          =   225
         Left            =   180
         TabIndex        =   37
         Top             =   240
         Width           =   870
         _ExtentX        =   1535
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
         Caption         =   "frmPrintPreview.frx":07A6
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   -1  'True
         Tip             =   "frmPrintPreview.frx":07D4
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmPrintPreview.frx":07F4
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniComboImageXP cmbDevice 
         Height          =   315
         Left            =   60
         TabIndex        =   40
         Top             =   540
         Width           =   3060
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
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Tip             =   "frmPrintPreview.frx":0810
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmPrintPreview.frx":0830
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP btnCustom 
         Height          =   315
         Left            =   540
         TabIndex        =   41
         Top             =   900
         Width           =   2115
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
         Caption         =   "frmPrintPreview.frx":084C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmPrintPreview.frx":089A
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmPrintPreview.frx":08BA
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
      Cancel          =   -1  'True
      Height          =   405
      Left            =   1320
      TabIndex        =   1
      Top             =   5820
      Width           =   1020
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
      Caption         =   "frmPrintPreview.frx":08D6
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmPrintPreview.frx":0902
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmPrintPreview.frx":0922
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdOK 
      Height          =   405
      Left            =   120
      TabIndex        =   42
      Top             =   5820
      Width           =   1020
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
      Caption         =   "frmPrintPreview.frx":093E
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmPrintPreview.frx":0968
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmPrintPreview.frx":0988
      RightToLeft     =   0   'False
   End
   Begin MSComDlg.CommonDialog cdlgExportFile 
      Left            =   2580
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VSPrinter7LibCtl.VSPrinter vp 
      Height          =   5295
      Left            =   3300
      TabIndex        =   4
      Top             =   300
      Width           =   6135
      _cx             =   10821
      _cy             =   9340
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      MousePointer    =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty HdrFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _ConvInfo       =   1
      AutoRTF         =   -1  'True
      Preview         =   -1  'True
      DefaultDevice   =   0   'False
      PhysicalPage    =   -1  'True
      AbortWindow     =   -1  'True
      AbortWindowPos  =   0
      AbortCaption    =   "Printing..."
      AbortTextButton =   "Cancel"
      AbortTextDevice =   "on the %s on %s"
      AbortTextPage   =   "Now printing Page %d of"
      FileName        =   ""
      MarginLeft      =   1440
      MarginTop       =   1440
      MarginRight     =   1440
      MarginBottom    =   1440
      MarginHeader    =   0
      MarginFooter    =   0
      IndentLeft      =   0
      IndentRight     =   0
      IndentFirst     =   0
      IndentTab       =   720
      SpaceBefore     =   0
      SpaceAfter      =   0
      LineSpacing     =   100
      Columns         =   1
      ColumnSpacing   =   180
      ShowGuides      =   2
      LargeChangeHorz =   300
      LargeChangeVert =   300
      SmallChangeHorz =   30
      SmallChangeVert =   30
      Track           =   0   'False
      ProportionalBars=   -1  'True
      Zoom            =   30.0189393939394
      ZoomMode        =   3
      ZoomMax         =   400
      ZoomMin         =   10
      ZoomStep        =   10
      EmptyColor      =   -2147483636
      TextColor       =   0
      HdrColor        =   0
      BrushColor      =   0
      BrushStyle      =   0
      PenColor        =   0
      PenStyle        =   0
      PenWidth        =   0
      PageBorder      =   0
      Header          =   ""
      Footer          =   ""
      TableSep        =   "|;"
      TableBorder     =   7
      TablePen        =   0
      TablePenLR      =   0
      TablePenTB      =   0
      NavBar          =   0
      NavBarColor     =   -2147483633
      ExportFormat    =   0
      URL             =   ""
      Navigation      =   3
      NavBarMenuText  =   "Whole &Page|Page &Width|&Two Pages|Thumb&nail"
   End
   Begin VB.PictureBox picPrinterControlBar 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3540
      ScaleHeight     =   375
      ScaleWidth      =   4710
      TabIndex        =   7
      Top             =   5760
      Width           =   4710
      Begin HexUniControls.ctlUniButtonImageXP cmdPage 
         Height          =   250
         Index           =   0
         Left            =   660
         TabIndex        =   10
         Top             =   60
         Width           =   250
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
         Caption         =   "frmPrintPreview.frx":09A4
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmPrintPreview.frx":09C8
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmPrintPreview.frx":09E8
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdPage 
         Height          =   250
         Index           =   1
         Left            =   915
         TabIndex        =   15
         Top             =   60
         Width           =   250
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
         Caption         =   "frmPrintPreview.frx":0A04
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmPrintPreview.frx":0A26
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmPrintPreview.frx":0A46
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdPage 
         Height          =   250
         Index           =   2
         Left            =   1560
         TabIndex        =   18
         Top             =   60
         Width           =   250
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
         Caption         =   "frmPrintPreview.frx":0A62
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmPrintPreview.frx":0A84
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmPrintPreview.frx":0AA4
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdPage 
         Height          =   250
         Index           =   3
         Left            =   1815
         TabIndex        =   24
         Top             =   60
         Width           =   250
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
         Caption         =   "frmPrintPreview.frx":0AC0
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmPrintPreview.frx":0AE4
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmPrintPreview.frx":0B04
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdZoom 
         Height          =   250
         Index           =   0
         Left            =   3000
         TabIndex        =   27
         Top             =   60
         Width           =   250
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
         Caption         =   "frmPrintPreview.frx":0B20
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmPrintPreview.frx":0B44
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmPrintPreview.frx":0B64
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdZoom 
         Height          =   250
         Index           =   1
         Left            =   3255
         TabIndex        =   30
         Top             =   60
         Width           =   250
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
         Caption         =   "frmPrintPreview.frx":0B80
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmPrintPreview.frx":0BA2
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmPrintPreview.frx":0BC2
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdZoom 
         Height          =   250
         Index           =   2
         Left            =   3900
         TabIndex        =   32
         Top             =   60
         Width           =   250
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
         Caption         =   "frmPrintPreview.frx":0BDE
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmPrintPreview.frx":0C00
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmPrintPreview.frx":0C20
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdZoom 
         Height          =   250
         Index           =   3
         Left            =   4155
         TabIndex        =   34
         Top             =   60
         Width           =   250
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
         Caption         =   "frmPrintPreview.frx":0C3C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmPrintPreview.frx":0C60
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmPrintPreview.frx":0C80
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label1 
         Height          =   165
         Index           =   12
         Left            =   -60
         Top             =   105
         Width           =   360
         _ExtentX        =   741
         _ExtentY        =   344
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmPrintPreview.frx":0C9C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   1
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   -1  'True
         Tip             =   "frmPrintPreview.frx":0CC6
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPrintPreview.frx":0CE6
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label1 
         Height          =   165
         Index           =   3
         Left            =   2370
         Top             =   105
         Width           =   420
         _ExtentX        =   794
         _ExtentY        =   344
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmPrintPreview.frx":0D02
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   1
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   -1  'True
         Tip             =   "frmPrintPreview.frx":0D2C
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPrintPreview.frx":0D4C
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblPage 
         Height          =   165
         Left            =   1140
         Top             =   105
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
         Caption         =   "frmPrintPreview.frx":0D68
         BackColor       =   14737632
         ForeColor       =   -2147483630
         Alignment       =   2
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmPrintPreview.frx":0D8A
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPrintPreview.frx":0DAA
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblZoom 
         Height          =   165
         Left            =   3480
         Top             =   105
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
         Caption         =   "frmPrintPreview.frx":0DC6
         BackColor       =   14737632
         ForeColor       =   -2147483630
         Alignment       =   2
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmPrintPreview.frx":0DEC
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPrintPreview.frx":0E0C
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL Frame1 
      Height          =   1635
      Index           =   2
      Left            =   60
      TabIndex        =   13
      Top             =   1200
      Width           =   3165
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
      Caption         =   "frmPrintPreview.frx":0E28
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmPrintPreview.frx":0E5C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmPrintPreview.frx":0E7C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniRadioXP opOrient 
         Height          =   315
         Index           =   0
         Left            =   900
         TabIndex        =   21
         Top             =   1140
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
         Caption         =   "frmPrintPreview.frx":0E98
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   -1  'True
         Tip             =   "frmPrintPreview.frx":0ECA
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmPrintPreview.frx":0EEA
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP opOrient 
         Height          =   315
         Index           =   1
         Left            =   1920
         TabIndex        =   22
         Top             =   1140
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
         Caption         =   "frmPrintPreview.frx":0F06
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmPrintPreview.frx":0F3A
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmPrintPreview.frx":0F5A
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin VB.VScrollBar scrlPaperSize 
         Height          =   300
         Index           =   1
         Left            =   2760
         Min             =   -32767
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   727
         Width           =   195
      End
      Begin VB.VScrollBar scrlPaperSize 
         Height          =   300
         Index           =   0
         Left            =   1200
         Min             =   -32767
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   727
         Width           =   195
      End
      Begin HexUniControls.ctlUniTextBoxXP txtPaperSize 
         Height          =   285
         Index           =   1
         Left            =   2220
         TabIndex        =   19
         Top             =   735
         Width           =   555
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmPrintPreview.frx":0F76
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         Tip             =   "frmPrintPreview.frx":0F9C
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPrintPreview.frx":0FBC
      End
      Begin HexUniControls.ctlUniTextBoxXP txtPaperSize 
         Height          =   300
         Index           =   0
         Left            =   660
         TabIndex        =   16
         Top             =   727
         Width           =   555
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmPrintPreview.frx":0FD8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         Tip             =   "frmPrintPreview.frx":1000
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPrintPreview.frx":1020
      End
      Begin HexUniControls.ctlUniComboImageXP cmbPaperSizes 
         Height          =   315
         Left            =   180
         TabIndex        =   14
         Top             =   300
         Width           =   2820
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
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Tip             =   "frmPrintPreview.frx":103C
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmPrintPreview.frx":105C
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label1 
         Height          =   195
         Index           =   1
         Left            =   1680
         Top             =   780
         Width           =   465
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
         Caption         =   "frmPrintPreview.frx":1078
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   -1  'True
         Tip             =   "frmPrintPreview.frx":10A6
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPrintPreview.frx":10C6
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label1 
         Height          =   195
         Index           =   0
         Left            =   180
         Top             =   780
         Width           =   420
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
         Caption         =   "frmPrintPreview.frx":10E2
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   -1  'True
         Tip             =   "frmPrintPreview.frx":110E
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPrintPreview.frx":112E
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin VB.Image imgOrient 
         Height          =   480
         Index           =   0
         Left            =   300
         Picture         =   "frmPrintPreview.frx":114A
         Top             =   1080
         Width           =   480
      End
      Begin VB.Image imgOrient 
         Height          =   480
         Index           =   1
         Left            =   300
         Picture         =   "frmPrintPreview.frx":1454
         Top             =   1080
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin HexUniControls.ctlUniFrameWL Frame1 
      Height          =   975
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   3165
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
      Caption         =   "frmPrintPreview.frx":175E
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmPrintPreview.frx":178C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmPrintPreview.frx":17AC
      RightToLeft     =   0   'False
      Begin VB.VScrollBar scrlMargin 
         Height          =   300
         Index           =   2
         Left            =   2760
         Min             =   -32767
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   600
         Width           =   195
      End
      Begin HexUniControls.ctlUniTextBoxXP txtMargin 
         Height          =   285
         Index           =   2
         Left            =   2220
         TabIndex        =   11
         Top             =   615
         Width           =   555
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmPrintPreview.frx":17C8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         Tip             =   "frmPrintPreview.frx":17EC
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPrintPreview.frx":180C
      End
      Begin VB.VScrollBar scrlMargin 
         Height          =   300
         Index           =   3
         Left            =   2760
         Min             =   -32767
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   180
         Width           =   195
      End
      Begin HexUniControls.ctlUniTextBoxXP txtMargin 
         Height          =   285
         Index           =   3
         Left            =   2220
         TabIndex        =   5
         Top             =   195
         Width           =   555
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmPrintPreview.frx":1828
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         Tip             =   "frmPrintPreview.frx":184C
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPrintPreview.frx":186C
      End
      Begin VB.VScrollBar scrlMargin 
         Height          =   300
         Index           =   0
         Left            =   1260
         Min             =   -32767
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   600
         Width           =   195
      End
      Begin HexUniControls.ctlUniTextBoxXP txtMargin 
         Height          =   285
         Index           =   0
         Left            =   720
         TabIndex        =   8
         Top             =   615
         Width           =   555
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmPrintPreview.frx":1888
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         Tip             =   "frmPrintPreview.frx":18AC
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPrintPreview.frx":18CC
      End
      Begin VB.VScrollBar scrlMargin 
         Height          =   300
         Index           =   1
         Left            =   1260
         Min             =   -32767
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   180
         Width           =   195
      End
      Begin HexUniControls.ctlUniTextBoxXP txtMargin 
         Height          =   285
         Index           =   1
         Left            =   720
         TabIndex        =   2
         Top             =   195
         Width           =   555
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmPrintPreview.frx":18E8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         Tip             =   "frmPrintPreview.frx":190C
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPrintPreview.frx":192C
      End
      Begin HexUniControls.ctlUniLabelXP Label1 
         Height          =   195
         Index           =   8
         Left            =   300
         Top             =   240
         Width           =   285
         _ExtentX        =   476
         _ExtentY        =   344
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmPrintPreview.frx":1948
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   -1  'True
         Tip             =   "frmPrintPreview.frx":1972
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPrintPreview.frx":1992
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label1 
         Height          =   195
         Index           =   7
         Left            =   1740
         Top             =   240
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
         Caption         =   "frmPrintPreview.frx":19AE
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   -1  'True
         Tip             =   "frmPrintPreview.frx":19DA
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPrintPreview.frx":19FA
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label1 
         Height          =   195
         Index           =   6
         Left            =   315
         Top             =   660
         Width           =   270
         _ExtentX        =   503
         _ExtentY        =   344
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmPrintPreview.frx":1A16
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   -1  'True
         Tip             =   "frmPrintPreview.frx":1A3E
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPrintPreview.frx":1A5E
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label1 
         Height          =   195
         Index           =   9
         Left            =   1620
         Top             =   660
         Width           =   510
         _ExtentX        =   873
         _ExtentY        =   344
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmPrintPreview.frx":1A7A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   -1  'True
         Tip             =   "frmPrintPreview.frx":1AA8
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPrintPreview.frx":1AC8
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniLabelXP Label1 
      Height          =   195
      Index           =   4
      Left            =   3360
      Top             =   75
      Width           =   570
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
      Caption         =   "frmPrintPreview.frx":1AE4
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   -1  'True
      Tip             =   "frmPrintPreview.frx":1B12
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmPrintPreview.frx":1B32
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
End
Attribute VB_Name = "frmPrintPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmPrintPreview.frm
'' Description: Generic Print Preview form that allows the user to change
''              different print settings before printing
''
'' Author:      Genesis Financial Data Services
''              425 Woodmen Rd
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' ??/??/??     M Thorne    Created
'' 02/28/2001   DAJ         Made generic
'' 10/29/2012   DAJ         Added tree stuff to GridToTable
'' 10/14/2013   DAJ         Added scale controls; override for Turnkey
'' 12/04/2013   DAJ         Override for Turnkey Reports
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Public Enum ePrintToFile
    ePrintToFile_None = 0
    ePrintToFile_Rtf = 1
    ePrintToFile_Html = 2
    ePrintToFile_Both = 3
    ePrintToFile_Image = 4
End Enum

Private Type mPrivate
    objReport As Object
    vArgs As Variant
    strName As String
    dMargin(0 To 3) As Double
    bLandscape As Boolean
    bOverride As Boolean
    ToFileOptions As ePrintToFile
    bGoingToFile As Boolean
    bCustomFrmPrint As Boolean
    bCallAfterHeaderEvent As Boolean
End Type
Private m As mPrivate

Private PaperName(68)       As String

Public Property Let CustomPrint(ByVal bValue As Boolean)
    m.bCustomFrmPrint = bValue
End Property

Public Property Get GoingToFile() As Boolean
    GoingToFile = m.bGoingToFile
End Property
Public Property Let GoingToFile(ByVal bValue As Boolean)
    m.bGoingToFile = bValue
End Property

Sub ClearPreview()
    
    With vp
        .ShowGuides = False
        .StartDoc
        .KillDoc
    End With

End Sub

Function ToInches(ByVal v As Variant) As String
    
    ' convert from twips
    v = Format(v / 1440, "#0.##")
    
    ' trim string
    If Right(v, 1) = "." Then v = Left(v, Len(v) - 1)
    
    ' return measurement
    ToInches = v & """"

End Function

Sub UpdatePreview(Optional ByVal bOnlyIfVisible As Boolean = True, _
                    Optional ByVal bPrintingToFile As Boolean = False)
    
    Dim i%, s$
    
    If bOnlyIfVisible Then
        If Not Me.Visible Then Exit Sub
    End If
    
    ' redraw
    DoEvents
    
    With vp
    
        ' load device list
        If cmbDevice.ListCount = 0 Then
            For i = 0 To .NDevices - 1
                cmbDevice.AddItem .Devices(i)
            Next
        End If
        
        ' select current device
        If .Device <> cmbDevice Then
            For i = 0 To .NDevices - 1
                If .Devices(i) = .Device Then
                    cmbDevice.ListIndex = i
                    Exit For
                End If
            Next
            
            ' show list of papers for this device
            cmbPaperSizes.Clear
            For i = 1 To 68
                If PaperName(i) <> "" And .PaperSizes(i) = True Then
                    cmbPaperSizes.AddItem PaperName(i)
                    cmbPaperSizes.ItemData(cmbPaperSizes.NewIndex) = i
                End If
            Next
            If .PaperSizes(256) = True Then
                cmbPaperSizes.AddItem "Custom"
                cmbPaperSizes.ItemData(cmbPaperSizes.NewIndex) = 256
            End If
            For i = 0 To cmbPaperSizes.ListCount - 1
                If .PaperSize = cmbPaperSizes.ItemData(i) Then
                    cmbPaperSizes.ListIndex = i
                    Exit For
                End If
            Next
        End If
    
        ' show orientations
        i = .Orientation
        opOrient(i).Value = True
        imgOrient(i).Visible = True
        imgOrient(1 - i).Visible = False
        
        .ScaleOutput = Val(lblScaleAmount.Caption)
        .Zoom = Val(lblZoom.Caption)
        
        ' show margins
        If txtMargin(0) <> ToInches(.MarginTop) Then txtMargin(0) = ToInches(.MarginTop)
        If txtMargin(1) <> ToInches(.MarginLeft) Then txtMargin(1) = ToInches(.MarginLeft)
        If txtMargin(2) <> ToInches(.MarginBottom) Then txtMargin(2) = ToInches(.MarginBottom)
        If txtMargin(3) <> ToInches(.MarginRight) Then txtMargin(3) = ToInches(.MarginRight)
        
        ' select paper size
        If cmbPaperSizes.ListIndex > -1 Then
            If .PaperSize <> cmbPaperSizes.ItemData(cmbPaperSizes.ListIndex) Then
                .PaperSize = cmbPaperSizes.ItemData(cmbPaperSizes.ListIndex)
            End If
            i = .PaperSize
            txtPaperSize(0).Enabled = (i = 256)
            txtPaperSize(1).Enabled = (i = 256)
            scrlPaperSize(0).Enabled = (i = 256)
            scrlPaperSize(1).Enabled = (i = 256)
        End If
        
        ' show paper sizes
        .PhysicalPage = True
        If txtPaperSize(0) <> ToInches(.PageWidth) Then _
            txtPaperSize(0) = ToInches(.PageWidth)
        If txtPaperSize(1) <> ToInches(.PageHeight) Then _
            txtPaperSize(1) = ToInches(.PageHeight)
        .PhysicalPage = False
        
        ' update the preview
        DoEvents
        m.bGoingToFile = bPrintingToFile
        m.objReport.GenerateReport m.vArgs
        
        txtFromPage.Text = "1"
        txtToPage.Text = .PageCount
        
        'THIS SECTION MOVES THE DOCUMENT TO THE PRINTER FOR PRINTING...
        'With cPaperReport
        '    .vsPrt = vp
        '    .Run
        'End With
        
        's = "Page preview: device is '" & .Device & "'"
        's = s & ", driver is '" & .Driver & "'"
        's = s & ", port is '" & .Port & "'"
        's = s & ", paper is #" & .PaperSize
        '.StartDoc
        '.DrawRectangle .MarginLeft - 100, .MarginTop, _
        '    .MarginLeft - 144, .MarginTop + 1440
        'While .CurrentY < .PageHeight - 3 * .MarginBottom Or _
        '      .CurrentColumn < .Columns - 1 And .CurrentPage = 1
        '    vp = s
        'Wend
    End With

End Sub

Private Sub btnCustom_Click()
On Error GoTo ErrSection:
    
    With vp
        .PrintDialog False
        Select Case .Error
            Case 3: Err.Raise vbObjectError + 1000, , _
                "Can't access the printer"
            Case 4: Err.Raise vbObjectError + 1000, , _
                "The printer is not responding"
            Case 6: Err.Raise vbObjectError + 1000, , _
                "A print job is already printing"
            Case 7: Err.Raise vbObjectError + 1000, , _
                "The current printer settings are not compatible with the printer"
        End Select
    End With
    
ErrExit:
    Exit Sub

ErrSection:
    ShowMsg
    Resume ErrExit

End Sub

' select a new printer if this one is different
Private Sub cmbDevice_Click()
On Error GoTo ErrSection:

    Screen.MousePointer = vbHourglass
    If vp.Device <> cmbDevice Then
        ClearPreview
        vp.Device = cmbDevice
        UpdatePreview
    End If
    
ErrExit:
    Screen.MousePointer = vbDefault
    Exit Sub

ErrSection:
    ShowMsg
    Resume ErrExit

End Sub

' select a new paper size if this one is different
Private Sub cmbPaperSizes_Click()
On Error GoTo ErrSection:
    
    Dim i%
    
    Screen.MousePointer = vbHourglass
    With cmbPaperSizes
        i = .ItemData(.ListIndex)
        If vp.PaperSize <> i Then
            ClearPreview
            vp.PaperSize = i
            txtPaperSize(0).Enabled = (i = 256)
            txtPaperSize(1).Enabled = (i = 256)
            scrlPaperSize(0).Enabled = (i = 256)
            scrlPaperSize(1).Enabled = (i = 256)
            UpdatePreview
        End If
    End With
    
ErrExit:
    Screen.MousePointer = vbDefault
    Exit Sub

ErrSection:
    ShowMsg
    Resume ErrExit

End Sub

Private Sub cmdCancel_Click()
On Error GoTo ErrSection:

    Unload Me
    
ErrExit:
    Exit Sub

ErrSection:
    ShowMsg
    Resume ErrExit

End Sub

Private Sub cmdOK_Click()
On Error GoTo ErrSection:
    
    Dim lIndex As Long
    Dim strRTF As String
    Dim strFilter As String
    
    If m.bCustomFrmPrint = True Then
        Select Case True
            Case optPrinter = True
                For lIndex = 1 To CLng(ValOfText(txtNumCopies.Text))
                    m.objReport.PrintReport cmbDevice.Text  'name of selected printer
                Next lIndex
            Case optFile = True
                m.objReport.PrintReport 1               'saving / printing to file
            Case optClipBoard = True
                m.objReport.PrintReport 2
        End Select
        cmdCancel_Click
        Exit Sub
    End If
        
    Screen.MousePointer = vbHourglass
    
    Select Case True
        Case optPrinter = True
            vp.Preview = False
            For lIndex = 1 To CLng(ValOfText(txtNumCopies.Text))
                vp.PrintDoc False, CLng(ValOfText(txtFromPage.Text)), CLng(ValOfText(txtToPage.Text))
            Next lIndex
            
        Case optFile = True
            vp.Preview = True
            Select Case m.ToFileOptions
                Case ePrintToFile_Both
                    strFilter = "RTF Files (*.rtf)|*.rtf|HTML Files (*.htm)|*.htm"
                Case ePrintToFile_Html
                    strFilter = "HTML Files (*.htm)|*.htm"
                Case ePrintToFile_Rtf
                    strFilter = "RTF Files (*.rtf)|*.rtf"
                Case ePrintToFile_Image
                    strFilter = "Bitmap Files (*.bmp)|*.bmp|PNG Files (*.png)|*.png"
            End Select
            vp.ExportFile = CommonDialogFile(Me.cdlgExportFile, True, strFilter)
            If Right(UCase(vp.ExportFile), 4) = ".RTF" Then
                vp.ExportFormat = vpxRTF
            Else
                vp.ExportFormat = vpxPlainHTML
            End If
            UpdatePreview , True
            vp.ExportFile = ""
        
        Case optClipBoard = True
            Clipboard.Clear
            vp.Preview = True
            If m.ToFileOptions = ePrintToFile_Image Then
                vp.ExportFile = "BOXQBTOCLIPBOARD"
                UpdatePreview , True
                vp.ExportFile = ""
            Else
                vp.ExportFile = AddSlash(App.Path) & "TempClip.RTF"
                vp.ExportFormat = vpxRTF
                UpdatePreview , True
                vp.ExportFile = ""
                strRTF = FileToString(AddSlash(App.Path) & "TempClip.RTF")
                Clipboard.SetText strRTF, vbCFRTF
                KillFile AddSlash(App.Path) & "TempClip.RTF"
            End If
            InfBox "You can now paste the report into |another application by selecting |'Edit-Paste'  (or hit 'Ctrl-V').", "i"
    End Select
    
    Unload Me
    
ErrExit:
    Screen.MousePointer = vbDefault
    Exit Sub

ErrSection:
    ShowMsg
    Resume ErrExit

End Sub

Private Sub cmdPage_Click(Index As Integer)
On Error GoTo ErrSection:

    Screen.MousePointer = vbHourglass
    Select Case Index
        Case 0
            vp.PreviewPage = 1
        Case 1
            vp.PreviewPage = vp.PreviewPage - 1
        Case 2
            vp.PreviewPage = vp.PreviewPage + 1
        Case 3
            vp.PreviewPage = vp.PageCount
    End Select
    UpdateControlBar
    
ErrExit:
    Screen.MousePointer = vbDefault
    Exit Sub

ErrSection:
    ShowMsg
    Resume ErrExit

End Sub

Private Sub cmdScale_Click(Index As Integer)
On Error GoTo ErrSection:
    
    Dim lScale As Long                  ' Scale amount

    lScale = CLng(Val(lblScaleAmount.Caption))

    Screen.MousePointer = vbHourglass
    Select Case Index
        Case 0
            lScale = 1
        Case 1
            lScale = lScale - 1
        Case 2
            lScale = lScale + 1
        Case 3
            lScale = 100
    End Select
    
    lblScaleAmount.Caption = Str(lScale)
    vp.ScaleOutput = lScale
    UpdateControlBar
    
ErrExit:
    Screen.MousePointer = vbDefault
    Exit Sub

ErrSection:
    ShowMsg
    Resume ErrExit

End Sub

Private Sub cmdZoom_Click(Index As Integer)
On Error GoTo ErrSection:
    
    Dim z%

    Screen.MousePointer = vbHourglass
    Select Case Index
        Case 0
            z = vp.ZoomMin
        Case 1
            z = vp.Zoom - vp.ZoomStep
        Case 2
            z = vp.Zoom + vp.ZoomStep
        Case 3
            z = vp.ZoomMax
    End Select
    If z < vp.ZoomMin Then z = vp.ZoomMin
    If z > vp.ZoomMax Then z = vp.ZoomMax
    vp.Zoom = z
    UpdateControlBar
    
ErrExit:
    Screen.MousePointer = vbDefault
    Exit Sub

ErrSection:
    ShowMsg
    Resume ErrExit

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

#If TRADENAV_EXE Then
    If KeyCode = vbKeyF1 Then
        KeyCode = 0
        g.Help.ShowF1Help Me
    End If
#End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPrintPreview.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub Form_Load()
On Error GoTo ErrSection:

    Dim strText$, strKey$
    
    g.Styler.StyleForm Me

    strKey = "Software\Genesis Financial Data Services\PrintSetup\" & m.strName
    strText = GetRegistryValue(rkLocalMachine, strKey, "Placement", "")
    If strText = "" Then
        CenterTheForm Me
    Else
        SetFormPlacement Me, strText, "LHTW"
    End If

    Screen.MousePointer = vbHourglass
    
    ' define paper names
    PaperName(0) = ""
    PaperName(1) = "Letter 8 1/2 x 11 in"
    PaperName(2) = "Letter Small 8 1/2 x 11 in"
    PaperName(3) = "Tabloid 11 x 17 in"
    PaperName(4) = "Ledger 17 x 11 in"
    PaperName(5) = "Legal 8 1/2 x 14 in"
    PaperName(6) = "Statement 5 1/2 x 8 1/2 in"
    PaperName(7) = "Executive 7 1/4 x 10 1/2 in"
    PaperName(8) = "A3 297 x 420 mm"
    PaperName(9) = "A4 210 x 297 mm"
    PaperName(10) = "A4 Small 210 x 297 mm"
    PaperName(11) = "A5 148 x 210 mm"
    PaperName(12) = "B4 (JIS) 250 x 354"
    PaperName(13) = "B5 (JIS) 182 x 257 mm"
    PaperName(14) = "Folio 8 1/2 x 13 in"
    PaperName(15) = "Quarto 215 x 275 mm"
    PaperName(16) = "10x14 in"
    PaperName(17) = "11x17 in"
    PaperName(18) = "Note 8 1/2 x 11 in"
    PaperName(19) = "Envelope #9 3 7/8 x 8 7/8"
    PaperName(20) = "Envelope #10 4 1/8 x 9 1/2"
    PaperName(21) = "Envelope #11 4 1/2 x 10 3/8"
    PaperName(22) = "Envelope #12 4 \276 x 11"
    PaperName(23) = "Envelope #14 5 x 11 1/2"
    PaperName(24) = "C size sheet"
    PaperName(25) = "D size sheet"
    PaperName(26) = "E size sheet"
    PaperName(27) = "Envelope DL 110 x 220mm"
    PaperName(28) = "Envelope C5 162 x 229 mm"
    PaperName(29) = "Envelope C3  324 x 458 mm"
    PaperName(30) = "Envelope C4  229 x 324 mm"
    PaperName(31) = "Envelope C6  114 x 162 mm"
    PaperName(32) = "Envelope C65 114 x 229 mm"
    PaperName(33) = "Envelope B4  250 x 353 mm"
    PaperName(34) = "Envelope B5  176 x 250 mm"
    PaperName(35) = "Envelope B6  176 x 125 mm"
    PaperName(36) = "Envelope 110 x 230 mm"
    PaperName(37) = "Envelope Monarch 3.875 x 7.5 in"
    PaperName(38) = "6 3/4 Envelope 3 5/8 x 6 1/2 in"
    PaperName(39) = "US Std Fanfold 14 7/8 x 11 in"
    PaperName(40) = "German Std Fanfold 8 1/2 x 12 in"
    PaperName(41) = "German Legal Fanfold 8 1/2 x 13 in"
    PaperName(42) = "B4 (ISO) 250 x 353 mm"
    PaperName(43) = "Japanese Postcard 100 x 148 mm"
    PaperName(44) = "9 x 11 in"
    PaperName(45) = "10 x 11 in"
    PaperName(46) = "15 x 11 in"
    PaperName(47) = "Envelope Invite 220 x 220 mm"
    PaperName(48) = "" ' RESERVED--DO NOT USE
    PaperName(49) = "" ' RESERVED--DO NOT USE
    PaperName(50) = "Letter Extra 9 \275 x 12 in"
    PaperName(51) = "Legal Extra 9 \275 x 15 in"
    PaperName(52) = "Tabloid Extra 11.69 x 18 in"
    PaperName(53) = "A4 Extra 9.27 x 12.69 in"
    PaperName(54) = "Letter Transverse 8 \275 x 11 in"
    PaperName(55) = "A4 Transverse 210 x 297 mm"
    PaperName(56) = "Letter Extra Transverse 9\275 x 12 in"
    PaperName(57) = "SuperA/SuperA/A4 227 x 356 mm"
    PaperName(58) = "SuperB/SuperB/A3 305 x 487 mm"
    PaperName(59) = "Letter Plus 8.5 x 12.69 in"
    PaperName(60) = "A4 Plus 210 x 330 mm"
    PaperName(61) = "A5 Transverse 148 x 210 mm"
    PaperName(62) = "B5 (JIS) Transverse 182 x 257 mm"
    PaperName(63) = "A3 Extra 322 x 445 mm"
    PaperName(64) = "A5 Extra 174 x 235 mm"
    PaperName(65) = "B5 (ISO) Extra 201 x 276 mm"
    PaperName(66) = "A2 420 x 594 mm"
    PaperName(67) = "A3 Transverse 297 x 420 mm"
    PaperName(68) = "A3 Extra Transverse 322 x 445 mm"
    
    vp.MarginTop = m.dMargin(0) * 1440
    vp.MarginLeft = m.dMargin(1) * 1440
    vp.MarginBottom = m.dMargin(2) * 1440
    vp.MarginRight = m.dMargin(3) * 1440
    
    ' save scrollbar values
    Dim i%
    For i = 0 To 3
        txtMargin(i).Text = CStr(m.dMargin(i))
        
        If m.bOverride = False Then
            scrlMargin(i).Value = GetRegistryValue(rkLocalMachine, strKey, "Margin" & Str(i), scrlMargin(i).Value)
        End If
        scrlMargin(i).Tag = scrlMargin(i).Value
    Next
    For i = 0 To 1
        scrlPaperSize(i).Value = GetRegistryValue(rkLocalMachine, strKey, "PaperSize", scrlPaperSize(i).Value)
        scrlPaperSize(i).Tag = scrlPaperSize(i).Value
    Next
    
    If m.bOverride = True Then
        opOrient(1).Value = m.bLandscape
    Else
        opOrient(1).Value = GetRegistryValue(rkLocalMachine, strKey, "Landscape", m.bLandscape)
    End If
    
    txtNumCopies.Text = 1
    
    strText = GetRegistryValue(rkLocalMachine, strKey, "Destination", "0")
    Select Case strText
        Case "0"
            optPrinter = True
        Case "1"
            optFile = True
        Case "2"
            optClipBoard = True
    End Select
    
    If m.bOverride Then
        lblScaleAmount.Caption = "100"
    Else
        lblScaleAmount.Caption = GetRegistryValue(rkLocalMachine, strKey, "Scale", "100")
    End If
    
    ' show the form
    vp.Preview = True
    UpdatePreview False
        
ErrExit:
    Screen.MousePointer = vbDefault
    Exit Sub

ErrSection:
    ShowMsg
    Resume ErrExit

End Sub

Private Sub Form_Resize()
On Error Resume Next

    Dim lMinScaleWidth As Long          ' Minimum scale width
    Dim lMinScaleHeight As Long         ' Minimum scale height

    lMinScaleWidth = (Frame1(0).Width * 3) + (Frame1(0).Left * 3)
    lMinScaleHeight = cmdOK.Top + cmdOK.Height + Frame1(0).Top

    If Not LimitFormSize(Me, lMinScaleWidth, lMinScaleHeight) Then
        With vp
            .Move .Left, .Top, Me.ScaleWidth - .Left - btnCustom.Left, Me.ScaleHeight - .Top - picPrinterControlBar.Height
        End With
        
        With picPrinterControlBar
            .Move ((vp.Width / 2) + vp.Left) - (.Width / 2), vp.Top + vp.Height
        End With
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Dim strKey$

    ' store scrollbar values
    Dim i%
    strKey = "Software\Genesis Financial Data Services\PrintSetup\" & m.strName
    For i = 0 To 3
        SetRegistryValue rkLocalMachine, strKey, "Margin" & Str(i), scrlMargin(i).Value
    Next
    For i = 0 To 1
        SetRegistryValue rkLocalMachine, strKey, "PaperSize", scrlPaperSize(i).Value
    Next
    
    SetRegistryValue rkLocalMachine, strKey, "Landscape", CInt(opOrient(1).Value)

    SetRegistryValue rkLocalMachine, strKey, "Placement", GetFormPlacement(Me)
    
    Select Case True
        Case optPrinter
            SetRegistryValue rkLocalMachine, strKey, "Destination", "0"
        Case optFile
            SetRegistryValue rkLocalMachine, strKey, "Destination", "1"
        Case optClipBoard
            SetRegistryValue rkLocalMachine, strKey, "Destination", "2"
    End Select
    
    SetRegistryValue rkLocalMachine, strKey, "Scale", lblScaleAmount.Caption

    If m.bCustomFrmPrint = True Then
        m.objReport.EndReport m.vArgs
    End If

End Sub

Private Sub lblPage_Click()
On Error GoTo ErrSection:
    
    Dim p&

    p = vp.PageCount \ 2
    If p < 1 Then p = 1
    p = ValOfText(InfBox("Go to Page ...", "?", , Me.Caption, , , , , , "s", CStr(p)))
    If p > 0 Then
        Screen.MousePointer = vbHourglass
        If p > vp.PageCount Then p = vp.PageCount
        vp.PreviewPage = p
        UpdateControlBar
    End If
    
ErrExit:
    Screen.MousePointer = vbDefault
    Exit Sub

ErrSection:
    ShowMsg
    Resume ErrExit

End Sub

Private Sub lblScaleAmount_Click()
On Error GoTo ErrSection:
    
    Dim lScale As Long                  ' Scale amount

    lScale = ValOfText(InfBox("Zoom percentage ...", "?", , Me.Caption, , , , , , "s", "100"))
    If lScale > 0 Then
        Screen.MousePointer = vbHourglass
        
        If lScale > 100 Then lScale = 100
        
        lblScaleAmount.Caption = Str(lScale)
        vp.ScaleOutput = lScale
        UpdateControlBar
    End If
    
ErrExit:
    Screen.MousePointer = vbDefault
    Exit Sub

ErrSection:
    ShowMsg
    Resume ErrExit

End Sub

Private Sub lblZoom_Click()
On Error GoTo ErrSection:
    
    Dim z&

    z = ValOfText(InfBox("Zoom percentage ...", "?", , Me.Caption, , , , , , "s", "100"))
    If z > 0 Then
        Screen.MousePointer = vbHourglass
        If z < vp.ZoomMin Then z = vp.ZoomMin
        If z > vp.ZoomMax Then z = vp.ZoomMax
        vp.Zoom = z
        UpdateControlBar
    End If
    
ErrExit:
    Screen.MousePointer = vbDefault
    Exit Sub

ErrSection:
    ShowMsg
    Resume ErrExit

End Sub

' select a new orientation if this one is different
Private Sub opOrient_Click(Index As Integer)
On Error GoTo ErrSection:

    Screen.MousePointer = vbHourglass
    If vp.Orientation <> Index Then
        'ClearPreview
        vp.Orientation = Index
        UpdatePreview
    End If
        
ErrExit:
    Screen.MousePointer = vbDefault
    Exit Sub

ErrSection:
    ShowMsg
    Resume ErrExit

End Sub

Private Sub optClipBoard_Click()
On Error GoTo ErrSection:

    cmbDevice.Enabled = False
    btnCustom.Enabled = False
    
ErrExit:
    Exit Sub
    
ErrSection:
    ShowMsg
    Resume ErrExit
    
End Sub

Private Sub optFile_Click()
On Error GoTo ErrSection:

    cmbDevice.Enabled = False
    btnCustom.Enabled = False
    
ErrExit:
    Exit Sub
    
ErrSection:
    ShowMsg
    Resume ErrExit
    
End Sub

Private Sub optPrinter_Click()
On Error GoTo ErrSection:

    cmbDevice.Enabled = True
    btnCustom.Enabled = True
    
ErrExit:
    Exit Sub
    
ErrSection:
    ShowMsg
    Resume ErrExit
    
End Sub

'Apply change to margin
Private Sub scrlMargin_Change(Index As Integer)
On Error GoTo ErrSection:

    Screen.MousePointer = vbHourglass
    With scrlMargin(Index)
        If Val(.Value) < Val(.Tag) Then
            txtMargin(Index).Text = CStr(Val(txtMargin(Index).Text) + 0.1)
        Else
            txtMargin(Index).Text = CStr(Val(txtMargin(Index).Text) - 0.1)
        End If
        .Tag = .Value
        txtMargin_LostFocus (Index)
    End With
        
ErrExit:
    Screen.MousePointer = vbDefault
    Exit Sub

ErrSection:
    ShowMsg
    Resume ErrExit

End Sub

'Apply change to paper size
Private Sub scrlPaperSize_Change(Index As Integer)
On Error GoTo ErrSection:

    Screen.MousePointer = vbHourglass
    With scrlPaperSize(Index)
        If Val(.Value) < Val(.Tag) Then
            txtPaperSize(Index).Text = CStr(Val(txtPaperSize(Index).Text) + 0.1)
        Else
            txtPaperSize(Index).Text = CStr(Val(txtPaperSize(Index).Text) - 0.1)
        End If
        .Tag = .Value
        txtPaperSize_LostFocus (Index)
    End With
        
ErrExit:
    Screen.MousePointer = vbDefault
    Exit Sub

ErrSection:
    ShowMsg
    Resume ErrExit

End Sub

Private Sub txtFromPage_LostFocus()

    If CLng(ValOfText(txtToPage.Text)) < CLng(ValOfText(txtFromPage.Text)) Then
        txtFromPage.Text = txtToPage.Text
    End If

    If CLng(ValOfText(txtFromPage.Text)) < 1 Or CLng(ValOfText(txtFromPage.Text)) > vp.PageCount Then
        txtFromPage.Text = "1"
    End If
    
End Sub

Private Sub txtMargin_GotFocus(Index As Integer)
    
    SelectAll txtMargin(Index)

End Sub

Private Sub txtMargin_KeyPress(Index As Integer, KeyAscii As Integer, Shift As Integer)
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtMargin_LostFocus (Index)
    End If

End Sub

' apply new margin
Private Sub txtMargin_LostFocus(Index As Integer)
    
    Dim v
    
    With txtMargin(Index)
        v = ValOfText(.Text) * 1440
        If v < 0 Then v = 0
        Select Case Index
            Case 0:
                If vp.MarginTop = v Then Exit Sub
                vp.MarginTop = v
            Case 1:
                If vp.MarginLeft = v Then Exit Sub
                vp.MarginLeft = v
            Case 2:
                If vp.MarginBottom = v Then Exit Sub
                vp.MarginBottom = v
            Case 3:
                If vp.MarginRight = v Then Exit Sub
                vp.MarginRight = v
        End Select
    End With
    UpdatePreview

End Sub

Private Sub txtNumCopies_LostFocus()

    If ValOfText(txtNumCopies.Text) <= 0 Then
        txtNumCopies.Text = "1"
    End If

End Sub

Private Sub txtPaperSize_GotFocus(Index As Integer)
    
    SelectAll txtPaperSize(Index)

End Sub

Private Sub txtPaperSize_KeyPress(Index As Integer, KeyAscii As Integer, Shift As Integer)
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtPaperSize_LostFocus (Index)
    End If

End Sub

Private Sub txtPaperSize_LostFocus(Index As Integer)
    
    Dim v
    
    v = ValOfText(txtPaperSize(Index)) * 1440
    If v < 0 Then v = 0
    ClearPreview
    If Index = 0 Then
        vp.PaperWidth = v
    Else
        vp.PaperHeight = v
    End If
    UpdatePreview

End Sub

Sub UpdateControlBar()
    
    ' update labels
    lblPage = Format(vp.PreviewPage)
    lblZoom = CStr(vp.Zoom) & "%"
        
    ' enable/disable buttons
    cmdPage(0).Enabled = (vp.PreviewPage > 1)
    cmdPage(1).Enabled = (vp.PreviewPage > 1)
    cmdPage(2).Enabled = (vp.PreviewPage < vp.PageCount)
    cmdPage(3).Enabled = (vp.PreviewPage < vp.PageCount)
    cmdZoom(0).Enabled = (vp.Zoom > vp.ZoomMin)
    cmdZoom(1).Enabled = (vp.Zoom > vp.ZoomMin)
    cmdZoom(2).Enabled = (vp.Zoom < vp.ZoomMax)
    cmdZoom(3).Enabled = (vp.Zoom < vp.ZoomMax)
    
    ' variable zoom rate
    If vp.Zoom >= 50 Then
        vp.ZoomStep = 10
    Else
        vp.ZoomStep = 5
    End If

End Sub

Public Function ShowMe(ByVal strName As String, ByVal objReport As Object, _
                    Optional ByVal vArgs As Variant = 0, Optional dTopMargin# = 1, _
                    Optional ByVal dBottomMargin# = 1, Optional ByVal dLeftMargin# = 1, _
                    Optional ByVal dRightMargin# = 1, Optional ByVal bLandscape As Boolean = False, _
                    Optional ByVal bOverride As Boolean = False, _
                    Optional ByVal ToFileOptions As ePrintToFile = ePrintToFile_Both, _
                    Optional ByVal bCallAfterHeaderEvent As Boolean = False) As Boolean

    ' Need to make sure to clear this...
    m.bCustomFrmPrint = False
       
    Set m.objReport = objReport
    m.vArgs = vArgs
    m.strName = strName

    m.dMargin(0) = dTopMargin
    m.dMargin(1) = dLeftMargin
    m.dMargin(2) = dBottomMargin
    m.dMargin(3) = dRightMargin
    
    m.bLandscape = bLandscape
    
    m.ToFileOptions = ToFileOptions
    If ToFileOptions = ePrintToFile_None Then
        optFile.Visible = False
        optClipBoard.Visible = False
    End If
    
    If (UCase(strName) = "TURNKEY") Or (UCase(strName) = "TURNKEY REPORT") Then
        m.bOverride = bOverride
    Else
        m.bOverride = True
    End If
    
    ' TLB 5/16/2012: to call the "AfterHeaderEvent" for custom page headers (e.g. with a logo, etc)
    m.bCallAfterHeaderEvent = bCallAfterHeaderEvent

    ShowForm Me, True
    
End Function

'Private Sub txtScale_LostFocus()
'
'    Dim lScale As Long
'
'    lScale = CLng(Val(txtScale.Text))
'    If lScale <= 0 Then
'        lScale = 1
'        txtScale.Text = "1"
'    ElseIf lScale > 100 Then
'        lScale = 100
'        txtScale.Text = "100"
'    End If
'
'    vp.ScaleOutput = lScale
'    UpdateControlBar
'
'End Sub

Private Sub txtToPage_LostFocus()

    If CLng(ValOfText(txtToPage.Text)) > vp.PageCount Then
        txtToPage.Text = CStr(vp.PageCount)
    End If
    
    If CLng(ValOfText(txtToPage.Text)) < CLng(ValOfText(txtFromPage.Text)) Then
        txtToPage.Text = txtFromPage.Text
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GridToTable
'' Description: Convert a FlexGrid to a table
'' Inputs:      Flex Grid, Show Tree?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GridToTable(fg As VSFlexGrid, Optional ByVal bShowTree As Boolean = True)

    Dim lIndex As Long                  ' Index into a for loop
    Dim lRows As Long                   ' Number of visible rows
    Dim lRow As Long                    ' Index into a for loop
    Dim lCol As Long                    ' Index into a for loop
    Dim lGridCol As Long                ' Grid column
    Dim strText As String               ' Text to display in the cell
    Dim bChecked As Boolean
    
    ' Figure out the number of visible rows in the grid...
    lRows = 0
    For lIndex = 0 To fg.Rows - 1
        If fg.RowHidden(lIndex) = False Then lRows = lRows + 1
    Next lIndex
    
    With vp
        .StartTable
        
        ' Initialize the table...
        .TableCell(tcCols) = 0
        .TableCell(tcRows) = 0
                
        ' Set up the visible columns of the grid in the table...
        For lIndex = 0 To fg.Cols - 1
            If fg.ColHidden(lIndex) = False Then
                .TableCell(tcCols) = .TableCell(tcCols) + 1
                .TableCell(tcColWidth, , .TableCell(tcCols)) = fg.ColWidth(lIndex) + 100
                .TableCell(tcColData, , .TableCell(tcCols)) = lIndex
            End If
        Next lIndex
        
        ' Walk through the visible rows in the grid and fill in the table...
        For lRow = 0 To fg.Rows - 1
            If fg.RowHidden(lRow) = False Then
                .TableCell(tcRows) = .TableCell(tcRows) + 1
                
                For lCol = 1 To .TableCell(tcCols)
                    lGridCol = .TableCell(tcColData, , lCol)
                    
                    strText = ""
                    If (bShowTree = True) And (lCol = 1) And (fg.OutlineBar <> flexOutlineBarNone) Then
                        For lIndex = 1 To fg.RowOutlineLevel(lRow) - 1
                            strText = strText & "|    "
                        Next lIndex
                        If fg.RowOutlineLevel(lRow) > 0 Then
                            strText = strText & "|-- "
                        End If
                    End If
                    
                    If (fg.ColDataType(lGridCol) = flexDTBoolean) And (lRow > 0) Then
                        ' TLB: had to replace the call to CheckedCell since frmPrintPreview is used
                        ' in other generic projects that don't have access to that routine.
                        'strText = strText & CStr(CheckedCell(fg, lRow, lGridCol))
                        bChecked = (fg.Cell(flexcpChecked, lRow, lGridCol) = flexChecked)
                        strText = strText & CStr(bChecked)
                    Else
                        strText = strText & fg.Cell(flexcpTextDisplay, lRow, lGridCol)
                    End If
                
                    .TableCell(tcText, .TableCell(tcRows), lCol) = strText
                Next lCol
            End If
        Next lRow
        
        ' Set some global properties for the table...
        .TableCell(tcFont, 1, 1, .TableCell(tcRows), .TableCell(tcCols)) = fg.Font
        .TableCell(tcAlign, 1, 1, .TableCell(tcRows), .TableCell(tcCols)) = taLeftMiddle
    
        If fg.GridLines = flexGridNone Then
            .TableBorder = tbNone
        Else
            .TableBorder = tbAll
        End If
        
        .EndTable
    End With

End Sub

Private Sub vp_AfterHeader()

    If m.bCallAfterHeaderEvent Then
        m.objReport.AfterHeaderEvent m.vArgs
    End If

End Sub







