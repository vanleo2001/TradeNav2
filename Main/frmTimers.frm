VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmTimers 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Refresh Rates"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5100
   Icon            =   "frmTimers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   60
      Top             =   60
   End
   Begin HexUniControls.ctlUniFrameWL Frame1 
      Height          =   4215
      Left            =   1620
      TabIndex        =   0
      Top             =   180
      Width           =   3435
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
      Caption         =   "frmTimers.frx":0442
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmTimers.frx":046E
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTimers.frx":048E
      RightToLeft     =   0   'False
      Begin MSComctlLib.Slider sldTime 
         Height          =   435
         Index           =   0
         Left            =   0
         TabIndex        =   1
         Top             =   300
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   767
         _Version        =   393216
         LargeChange     =   1
         Max             =   9
         TextPosition    =   1
      End
      Begin MSComctlLib.Slider sldTime 
         Height          =   435
         Index           =   1
         Left            =   0
         TabIndex        =   2
         Top             =   780
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   767
         _Version        =   393216
         LargeChange     =   1
         Max             =   9
      End
      Begin MSComctlLib.Slider sldTime 
         Height          =   435
         Index           =   2
         Left            =   0
         TabIndex        =   3
         Top             =   1260
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   767
         _Version        =   393216
         LargeChange     =   1
         Max             =   9
      End
      Begin MSComctlLib.Slider sldTime 
         Height          =   435
         Index           =   3
         Left            =   0
         TabIndex        =   4
         Top             =   1740
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   767
         _Version        =   393216
         LargeChange     =   1
         Max             =   9
      End
      Begin MSComctlLib.Slider sldTime 
         Height          =   435
         Index           =   4
         Left            =   0
         TabIndex        =   5
         Top             =   2220
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   767
         _Version        =   393216
         LargeChange     =   1
         Max             =   9
      End
      Begin MSComctlLib.Slider sldTime 
         Height          =   435
         Index           =   5
         Left            =   0
         TabIndex        =   6
         Top             =   2700
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   767
         _Version        =   393216
         LargeChange     =   1
         Max             =   9
      End
      Begin MSComctlLib.Slider sldTime 
         Height          =   435
         Index           =   6
         Left            =   0
         TabIndex        =   7
         Top             =   3180
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   767
         _Version        =   393216
         LargeChange     =   1
         Max             =   9
      End
      Begin MSComctlLib.Slider sldTime 
         Height          =   435
         Index           =   7
         Left            =   0
         TabIndex        =   8
         Top             =   3660
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   767
         _Version        =   393216
         LargeChange     =   1
         Max             =   9
      End
      Begin HexUniControls.ctlUniLabelXP lblTime 
         Height          =   255
         Index           =   9
         Left            =   3000
         Tag             =   "60000"
         Top             =   0
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
         Caption         =   "frmTimers.frx":04AA
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   2
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmTimers.frx":04CE
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTimers.frx":04EE
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblTime 
         Height          =   255
         Index           =   8
         Left            =   2640
         Tag             =   "30000"
         Top             =   0
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
         Caption         =   "frmTimers.frx":050A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   2
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmTimers.frx":052E
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTimers.frx":054E
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblTime 
         Height          =   255
         Index           =   7
         Left            =   2280
         Tag             =   "10000"
         Top             =   0
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
         Caption         =   "frmTimers.frx":056A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   2
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmTimers.frx":058E
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTimers.frx":05AE
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblTime 
         Height          =   255
         Index           =   6
         Left            =   1980
         Tag             =   "5000"
         Top             =   0
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
         Caption         =   "frmTimers.frx":05CA
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   2
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmTimers.frx":05EC
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTimers.frx":060C
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblTime 
         Height          =   255
         Index           =   5
         Left            =   1680
         Tag             =   "2000"
         Top             =   0
         Width           =   315
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
         Caption         =   "frmTimers.frx":0628
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   2
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmTimers.frx":064A
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTimers.frx":066A
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblTime 
         Height          =   255
         Index           =   4
         Left            =   1380
         Tag             =   "1000"
         Top             =   0
         Width           =   315
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
         Caption         =   "frmTimers.frx":0686
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   2
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmTimers.frx":06A8
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTimers.frx":06C8
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblTime 
         Height          =   255
         Index           =   3
         Left            =   1020
         Tag             =   "500"
         Top             =   0
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
         Caption         =   "frmTimers.frx":06E4
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   2
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmTimers.frx":070C
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTimers.frx":072C
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblTime 
         Height          =   255
         Index           =   2
         Left            =   660
         Tag             =   "250"
         Top             =   0
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
         Caption         =   "frmTimers.frx":0748
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   2
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmTimers.frx":076E
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTimers.frx":078E
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblTime 
         Height          =   255
         Index           =   1
         Left            =   360
         Tag             =   "125"
         Top             =   0
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
         Caption         =   "frmTimers.frx":07AA
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   2
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmTimers.frx":07D0
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTimers.frx":07F0
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblTime 
         Height          =   255
         Index           =   0
         Left            =   0
         Tag             =   "62"
         Top             =   0
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
         Caption         =   "frmTimers.frx":080C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   2
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmTimers.frx":0834
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTimers.frx":0854
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniLabelXP lblName 
      Height          =   255
      Index           =   7
      Left            =   0
      Top             =   3900
      Width           =   1575
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
      Caption         =   "frmTimers.frx":0870
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   1
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmTimers.frx":08AA
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTimers.frx":08CA
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblName 
      Height          =   255
      Index           =   6
      Left            =   0
      Top             =   3420
      Width           =   1575
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
      Caption         =   "frmTimers.frx":08E6
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   1
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmTimers.frx":0926
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTimers.frx":0946
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblName 
      Height          =   255
      Index           =   5
      Left            =   0
      Top             =   2940
      Width           =   1575
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
      Caption         =   "frmTimers.frx":0962
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   1
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmTimers.frx":099C
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTimers.frx":09BC
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblName 
      Height          =   255
      Index           =   4
      Left            =   0
      Top             =   2460
      Width           =   1575
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
      Caption         =   "frmTimers.frx":09D8
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   1
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmTimers.frx":0A12
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTimers.frx":0A32
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblName 
      Height          =   255
      Index           =   3
      Left            =   0
      Top             =   1980
      Width           =   1575
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
      Caption         =   "frmTimers.frx":0A4E
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   1
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmTimers.frx":0A88
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTimers.frx":0AA8
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblName 
      Height          =   255
      Index           =   2
      Left            =   0
      Top             =   1500
      Width           =   1575
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
      Caption         =   "frmTimers.frx":0AC4
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   1
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmTimers.frx":0AFE
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTimers.frx":0B1E
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblName 
      Height          =   255
      Index           =   1
      Left            =   0
      Top             =   1020
      Width           =   1575
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
      Caption         =   "frmTimers.frx":0B3A
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   1
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmTimers.frx":0B74
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTimers.frx":0B94
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP Label1 
      Height          =   255
      Left            =   240
      Top             =   165
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
      Caption         =   "frmTimers.frx":0BB0
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   1
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmTimers.frx":0BE2
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTimers.frx":0C02
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblName 
      Height          =   255
      Index           =   0
      Left            =   0
      Top             =   540
      Width           =   1575
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
      Caption         =   "frmTimers.frx":0C1E
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   1
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmTimers.frx":0C58
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTimers.frx":0C78
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
End
Attribute VB_Name = "frmTimers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type mPrivate
    bSettingSliders As Boolean
End Type
Private m As mPrivate

Private Sub Form_Load()

    Dim i&, strText$
    
    g.Styler.StyleForm Me
    
    strText = GetIniFileProperty("TimersPlacement", "", "Forms", g.strIniFile)
    If strText = "" Then
        CenterTheForm Me
    Else
        SetFormPlacement Me, strText, "LT" '"LTHW"
    End If
    
    For i = lblName.LBound To lblName.UBound
        lblName(i).Caption = ""
        sldTime(i).Visible = False
    Next

End Sub

Private Sub Form_Unload(Cancel As Integer)

    tmr.Enabled = False
    SetIniFileProperty "TimersPlacement", GetFormPlacement(Me), "Forms", g.strIniFile

End Sub

Public Sub ShowMe()

    SetAllSliders
    tmr.Enabled = True
    ShowForm Me, eForm_Nonmodal, frmMain

End Sub

Private Sub SetSlider(ByVal iIndex&, ByVal strName$, ByVal iInterval&, ByVal bEnabled As Boolean)

    Dim i&, iPos&
    lblName(iIndex).Caption = strName & ":"
    If iInterval = 0 Then bEnabled = False
    If iInterval >= 0 Then
        For i = 0 To 9
            iPos = i
            If iInterval <= Val(lblTime(i).Tag) Then Exit For
        Next
        If sldTime(iIndex).Value <> iPos Then
            sldTime(iIndex).Value = iPos
        End If
    End If
If iIndex = 7 Then bEnabled = True
    If sldTime(iIndex).Visible <> bEnabled Then
        sldTime(iIndex).Visible = bEnabled
    End If

End Sub

Private Sub SetAllSliders()

    Dim i&, iInterval&
    Dim frmActive As frmChart

    ' 16th, 8th, 4th, Half, 1, 2, 5, 10, 30, 60
    ' GenesisRT, Charts, QuoteBoard, TradeConsole, PriceLadders
    m.bSettingSliders = True
    
    SetSlider 0, "GenesisRT", g.RealTime.RtInterval, g.RealTime.ActiveRTG
    
    Set frmActive = ActiveChart
    If frmActive Is Nothing Then
        SetSlider 1, "Charts", 250, False
    Else
        SetSlider 1, "Charts", frmActive.tmr.Interval, frmActive.tmr.Enabled
        Set frmActive = Nothing
    End If
    
    SetSlider 2, "QuoteBoard", frmQuotes.tmrRealtime.Interval, frmQuotes.tmrRealtime.Enabled
       
    SetSlider 3, "MainForm", frmMain.tmrMain.Interval, frmMain.tmrMain.Enabled
    
    iInterval = 0
    For i = Forms.Count - 1 To 0 Step -1
        If TypeOf Forms(i) Is frmTickDistribution Then
            iInterval = Forms(i).tmr.Interval
            If Not Forms(i).tmr.Enabled Then iInterval = -iInterval
            Exit For
        End If
    Next
    SetSlider 4, "Price Ladders", Abs(iInterval), (iInterval > 0)
    
    SetSlider 5, "TC Realtime", frmTTSummary.tmrRealtime.Interval, frmTTSummary.tmrRealtime.Enabled
    SetSlider 6, "TC Brokers", frmTTSummary.tmrBrokers.Interval, frmTTSummary.tmrBrokers.Enabled
    SetSlider 7, "Broker Msg", frmOnlineBroker.tmrMessages.Interval, frmOnlineBroker.tmrMessages.Enabled
    
    m.bSettingSliders = False

End Sub

Private Sub sldTime_Change(Index As Integer)

    Dim i&, iInterval&

    If m.bSettingSliders Then Exit Sub
    
    m.bSettingSliders = True
    i = sldTime(Index).Value
    iInterval = Val(lblTime(i).Tag)
    
    Select Case Index
    Case 0
        If g.RealTime.ActiveRTG Then
            If iInterval >= 60000 Then
                iInterval = 0
            End If
            g.RealTime.RtInterval = iInterval
        End If
    
    Case 1
        For i = 0 To Forms.Count - 1
            If TypeOf Forms(i) Is frmChart Then
                Forms(i).tmr.Interval = iInterval
            End If
        Next
    
    Case 2
        frmQuotes.tmrRealtime.Interval = iInterval
            
    Case 3
        frmMain.tmrMain.Interval = iInterval
        
    Case 4
        For i = 0 To Forms.Count - 1
            If TypeOf Forms(i) Is frmTickDistribution Then
                Forms(i).tmr.Interval = iInterval
            End If
        Next
        
    Case 5
        frmTTSummary.tmrRealtime.Interval = iInterval
    
    Case 6
        frmTTSummary.tmrBrokers.Interval = iInterval
    
    Case 7
        frmOnlineBroker.tmrMessages.Interval = iInterval
    
    End Select
    
    m.bSettingSliders = False

End Sub

Private Sub tmr_Timer()

    If g.bUnloading Or g.bStarting Then Exit Sub
    If m.bSettingSliders Or MouseIsPressed Then Exit Sub
    
    SetAllSliders

End Sub

