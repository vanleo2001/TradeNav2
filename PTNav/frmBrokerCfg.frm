VERSION 5.00
Begin VB.Form frmBrokerCfg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   12315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12315
   ScaleWidth      =   14085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraTT 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1935
      Left            =   10020
      TabIndex        =   110
      Top             =   1320
      Width           =   3675
      Begin VB.Frame fraTtUserInfo 
         Caption         =   "User Information"
         Height          =   1095
         Left            =   0
         TabIndex        =   111
         Top             =   0
         Width           =   3615
         Begin VB.TextBox txtTtPassword 
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   1020
            PasswordChar    =   "*"
            TabIndex        =   113
            Top             =   240
            Width           =   2415
         End
         Begin VB.Label lblTtPassword 
            Caption         =   "Pass&word:"
            Height          =   255
            Left            =   120
            TabIndex        =   112
            Top             =   270
            Width           =   735
         End
      End
      Begin VB.Frame fraTtServerInfo 
         Caption         =   "Server Information"
         Height          =   735
         Left            =   0
         TabIndex        =   114
         Top             =   1200
         Width           =   3615
         Begin VB.TextBox txtTtIp 
            Height          =   285
            Left            =   480
            TabIndex        =   116
            Top             =   285
            Width           =   1515
         End
         Begin VB.TextBox txtTtPort 
            Height          =   315
            Left            =   2640
            TabIndex        =   118
            Top             =   270
            Width           =   795
         End
         Begin VB.Label lblTtIp 
            Caption         =   "&IP:"
            Height          =   255
            Left            =   180
            TabIndex        =   115
            Top             =   300
            Width           =   255
         End
         Begin VB.Label lblTtPort 
            Caption         =   "&Port:"
            Height          =   255
            Left            =   2160
            TabIndex        =   117
            Top             =   300
            Width           =   375
         End
      End
   End
   Begin VB.Frame fraFXCM 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2355
      Left            =   6180
      TabIndex        =   101
      Top             =   9600
      Width           =   3675
      Begin VB.Frame fraFxcmUserInfo 
         Caption         =   "User Information"
         Height          =   1095
         Left            =   0
         TabIndex        =   102
         Top             =   0
         Width           =   3615
         Begin VB.TextBox txtFxcmPassword 
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   1020
            PasswordChar    =   "*"
            TabIndex        =   106
            Top             =   600
            Width           =   2415
         End
         Begin VB.TextBox txtFxcmUserID 
            Height          =   315
            Left            =   1020
            TabIndex        =   104
            Top             =   240
            Width           =   2415
         End
         Begin VB.Label lblFxcmPassword 
            Caption         =   "Pass&word:"
            Height          =   255
            Left            =   120
            TabIndex        =   105
            Top             =   660
            Width           =   855
         End
         Begin VB.Label lblFxcmUserID 
            Caption         =   "&User ID:"
            Height          =   255
            Left            =   120
            TabIndex        =   103
            Top             =   270
            Width           =   615
         End
      End
      Begin VB.Frame fraFxcmServerInfo 
         Caption         =   "Server Information"
         Height          =   1095
         Left            =   0
         TabIndex        =   107
         Top             =   1200
         Width           =   3615
         Begin VB.TextBox txtFxcmUrl 
            Height          =   285
            Left            =   660
            TabIndex        =   109
            Top             =   285
            Width           =   2835
         End
         Begin VB.Label lblFxcmUrl 
            Caption         =   "U&RL:"
            Height          =   255
            Left            =   180
            TabIndex        =   108
            Top             =   300
            Width           =   435
         End
      End
   End
   Begin VB.Frame fraPfg 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2355
      Left            =   2280
      TabIndex        =   58
      Top             =   9660
      Width           =   3675
      Begin VB.Frame fraPfgServerInfo 
         Caption         =   "Server Information"
         Height          =   1095
         Left            =   0
         TabIndex        =   64
         Top             =   1200
         Width           =   3615
         Begin VB.TextBox txtPfgIP 
            Height          =   285
            Left            =   660
            TabIndex        =   66
            Top             =   285
            Width           =   2835
         End
         Begin VB.TextBox txtPfgPort 
            Height          =   315
            Left            =   660
            TabIndex        =   68
            Top             =   660
            Width           =   795
         End
         Begin VB.Label lblPfgIP 
            Caption         =   "&IP:"
            Height          =   255
            Left            =   180
            TabIndex        =   65
            Top             =   300
            Width           =   255
         End
         Begin VB.Label lblPfgPort 
            Caption         =   "&Port:"
            Height          =   255
            Left            =   180
            TabIndex        =   67
            Top             =   690
            Width           =   375
         End
      End
      Begin VB.Frame fraPfgUserInfo 
         Caption         =   "User Information"
         Height          =   1095
         Left            =   0
         TabIndex        =   59
         Top             =   0
         Width           =   3615
         Begin VB.TextBox txtPfgUserID 
            Height          =   315
            Left            =   1140
            TabIndex        =   61
            Top             =   240
            Width           =   2295
         End
         Begin VB.TextBox txtPfgPassword 
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   1140
            PasswordChar    =   "*"
            TabIndex        =   63
            Top             =   600
            Width           =   2295
         End
         Begin VB.Label lblPfgUserID 
            Caption         =   "&Account:"
            Height          =   255
            Left            =   120
            TabIndex        =   60
            Top             =   270
            Width           =   675
         End
         Begin VB.Label lblPfgPassword 
            Caption         =   "Access &Key:"
            Height          =   255
            Left            =   120
            TabIndex        =   62
            Top             =   660
            Width           =   975
         End
      End
   End
   Begin VB.Frame fraTransact 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3675
      Left            =   6180
      TabIndex        =   80
      Top             =   3780
      Width           =   3675
      Begin VB.TextBox txtTransactSymbols 
         Height          =   915
         Left            =   60
         MultiLine       =   -1  'True
         TabIndex        =   91
         Top             =   2700
         Width           =   3555
      End
      Begin VB.Frame fraTransactServerInfo 
         Caption         =   "Server Information"
         Height          =   735
         Left            =   0
         TabIndex        =   87
         Top             =   1620
         Width           =   3615
         Begin VB.TextBox txtTransactIP 
            Height          =   285
            Left            =   480
            TabIndex        =   89
            Top             =   285
            Width           =   2955
         End
         Begin VB.Label lblTransactIP 
            Caption         =   "&IP:"
            Height          =   255
            Left            =   180
            TabIndex        =   88
            Top             =   300
            Width           =   255
         End
      End
      Begin VB.Frame fraTransactUserInfo 
         Caption         =   "User Information"
         Height          =   1575
         Left            =   0
         TabIndex        =   81
         Top             =   0
         Width           =   3615
         Begin VB.TextBox txtTransactUserID 
            Height          =   315
            Left            =   1020
            TabIndex        =   84
            Top             =   720
            Width           =   2415
         End
         Begin VB.TextBox txtTransactPassword 
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   1020
            PasswordChar    =   "*"
            TabIndex        =   86
            Top             =   1080
            Width           =   2415
         End
         Begin VB.Label lblTransActInfo 
            Caption         =   "Please enter your broker-provided user name and  password in the appropriate places below:"
            Height          =   435
            Left            =   120
            TabIndex        =   82
            Top             =   240
            Width           =   3375
         End
         Begin VB.Label lblTransactUserID 
            Caption         =   "&User Name:"
            Height          =   255
            Left            =   120
            TabIndex        =   83
            Top             =   750
            Width           =   915
         End
         Begin VB.Label lblTransactPassword 
            Caption         =   "Pass&word:"
            Height          =   255
            Left            =   120
            TabIndex        =   85
            Top             =   1140
            Width           =   855
         End
      End
      Begin VB.Label lblTransactSymbols 
         Caption         =   "Authorized Symbols from TransAct:"
         Height          =   195
         Left            =   60
         TabIndex        =   90
         Top             =   2460
         Width           =   2475
      End
   End
   Begin VB.Frame fraAlaron 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2355
      Left            =   6180
      TabIndex        =   69
      Top             =   1320
      Width           =   3675
      Begin VB.Frame fraAlaronUserInfo 
         Caption         =   "User Information"
         Height          =   1095
         Left            =   0
         TabIndex        =   70
         Top             =   0
         Width           =   3615
         Begin VB.TextBox txtAlaronPassword 
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   1020
            PasswordChar    =   "*"
            TabIndex        =   74
            Top             =   630
            Width           =   2415
         End
         Begin VB.TextBox txtAlaronUserID 
            Height          =   315
            Left            =   1020
            TabIndex        =   72
            Top             =   240
            Width           =   2415
         End
         Begin VB.Label lblAlaronPassword 
            Caption         =   "Pass&word:"
            Height          =   255
            Left            =   120
            TabIndex        =   73
            Top             =   660
            Width           =   855
         End
         Begin VB.Label lblAlaronUserID 
            Caption         =   "&User ID:"
            Height          =   255
            Left            =   120
            TabIndex        =   71
            Top             =   270
            Width           =   615
         End
      End
      Begin VB.Frame fraAlaronServerInfo 
         Caption         =   "Server Information"
         Height          =   1095
         Left            =   0
         TabIndex        =   75
         Top             =   1200
         Width           =   3615
         Begin VB.TextBox txtAlaronPort 
            Height          =   315
            Left            =   660
            TabIndex        =   79
            Top             =   660
            Width           =   795
         End
         Begin VB.TextBox txtAlaronIP 
            Height          =   285
            Left            =   660
            TabIndex        =   77
            Top             =   285
            Width           =   2835
         End
         Begin VB.Label lblAlaronPort 
            Caption         =   "&Port:"
            Height          =   255
            Left            =   180
            TabIndex        =   78
            Top             =   690
            Width           =   375
         End
         Begin VB.Label lblAlaronIP 
            Caption         =   "&IP:"
            Height          =   255
            Left            =   180
            TabIndex        =   76
            Top             =   300
            Width           =   255
         End
      End
   End
   Begin VB.ListBox lstBrokers 
      Height          =   2205
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Frame fraPats 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2295
      Left            =   2280
      TabIndex        =   46
      Top             =   7200
      Width           =   3675
      Begin VB.Frame fraPatsUserInfo 
         Caption         =   "User Information"
         Height          =   1095
         Left            =   0
         TabIndex        =   47
         Top             =   0
         Width           =   3615
         Begin VB.TextBox txtPatsPassword 
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   1020
            PasswordChar    =   "*"
            TabIndex        =   51
            Top             =   630
            Width           =   2415
         End
         Begin VB.TextBox txtPatsUserID 
            Height          =   315
            Left            =   1020
            TabIndex        =   49
            Top             =   240
            Width           =   2415
         End
         Begin VB.Label lblPatsPassword 
            Caption         =   "Pass&word:"
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   660
            Width           =   855
         End
         Begin VB.Label lblPatsUserID 
            Caption         =   "&User ID:"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   270
            Width           =   615
         End
      End
      Begin VB.Frame fraPatsServer 
         Caption         =   "Server Information"
         Height          =   1095
         Left            =   0
         TabIndex        =   52
         Top             =   1200
         Width           =   3615
         Begin VB.CheckBox chkSuperTAS 
            Caption         =   "Use &SuperTAS"
            Height          =   255
            Left            =   480
            TabIndex        =   57
            Top             =   720
            Width           =   1935
         End
         Begin VB.TextBox txtPatsHostIP 
            Height          =   285
            Left            =   480
            TabIndex        =   54
            Top             =   285
            Width           =   1515
         End
         Begin VB.TextBox txtPatsHostPort 
            Height          =   315
            Left            =   2640
            TabIndex        =   56
            Top             =   270
            Width           =   795
         End
         Begin VB.Label lblPatsHostIP 
            Caption         =   "&IP:"
            Height          =   255
            Left            =   180
            TabIndex        =   53
            Top             =   300
            Width           =   255
         End
         Begin VB.Label lblPatsHostPort 
            Caption         =   "&Port:"
            Height          =   255
            Left            =   2160
            TabIndex        =   55
            Top             =   300
            Width           =   375
         End
      End
   End
   Begin VB.Frame fraLW 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2415
      Left            =   2280
      TabIndex        =   31
      Top             =   4620
      Width           =   3675
      Begin VB.TextBox txtLWFirm 
         Height          =   315
         Left            =   540
         TabIndex        =   43
         Top             =   2040
         Width           =   975
      End
      Begin VB.Frame fraLindWaldockUserInfo 
         Caption         =   "User Information"
         Height          =   1095
         Left            =   0
         TabIndex        =   32
         Top             =   0
         Width           =   3615
         Begin VB.TextBox txtLWPassword 
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   1020
            PasswordChar    =   "*"
            TabIndex        =   36
            Top             =   630
            Width           =   2415
         End
         Begin VB.TextBox txtLWUserID 
            Height          =   315
            Left            =   1020
            TabIndex        =   34
            Top             =   240
            Width           =   2415
         End
         Begin VB.Label lblLWPassword 
            Caption         =   "Pass&word:"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   660
            Width           =   855
         End
         Begin VB.Label lblLWUserID 
            Caption         =   "&User ID:"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   270
            Width           =   615
         End
      End
      Begin VB.Frame fraLindWaldockServer 
         Caption         =   "Server Information"
         Height          =   735
         Left            =   0
         TabIndex        =   37
         Top             =   1200
         Width           =   3615
         Begin VB.TextBox txtLWHost 
            Height          =   285
            Left            =   480
            TabIndex        =   39
            Top             =   285
            Width           =   1515
         End
         Begin VB.TextBox txtLWPort 
            Height          =   315
            Left            =   2640
            TabIndex        =   41
            Top             =   270
            Width           =   795
         End
         Begin VB.Label lblLWHost 
            Caption         =   "&IP:"
            Height          =   255
            Left            =   180
            TabIndex        =   38
            Top             =   300
            Width           =   255
         End
         Begin VB.Label lblLWPort 
            Caption         =   "&Port:"
            Height          =   255
            Left            =   2160
            TabIndex        =   40
            Top             =   300
            Width           =   375
         End
      End
      Begin VB.TextBox txtLWTestNum 
         Height          =   315
         Left            =   2700
         TabIndex        =   45
         Top             =   2040
         Width           =   915
      End
      Begin VB.Label lblLWFirm 
         Caption         =   "&Firm:"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   2070
         Width           =   435
      End
      Begin VB.Label lblLWTestNum 
         Caption         =   "&Subsystem:"
         Height          =   255
         Left            =   1740
         TabIndex        =   44
         Top             =   2070
         Width           =   855
      End
   End
   Begin VB.Frame fraIntBrokers 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1935
      Left            =   6180
      TabIndex        =   92
      Top             =   7560
      Width           =   3675
      Begin VB.Frame fraIntBrokersServer 
         Caption         =   "Server Information"
         Height          =   735
         Left            =   0
         TabIndex        =   96
         Top             =   1200
         Width           =   3615
         Begin VB.TextBox txtPort 
            Height          =   315
            Left            =   2640
            TabIndex        =   100
            Top             =   270
            Width           =   795
         End
         Begin VB.TextBox txtHostIP 
            Height          =   285
            Left            =   480
            TabIndex        =   98
            Top             =   285
            Width           =   1515
         End
         Begin VB.Label lblPort 
            Caption         =   "&Port:"
            Height          =   255
            Left            =   2160
            TabIndex        =   99
            Top             =   300
            Width           =   375
         End
         Begin VB.Label lblHostIP 
            Caption         =   "&IP:"
            Height          =   255
            Left            =   180
            TabIndex        =   97
            Top             =   300
            Width           =   255
         End
      End
      Begin VB.Frame fraIntBrokersUserInfo 
         Caption         =   "User Information"
         Height          =   1095
         Left            =   0
         TabIndex        =   93
         Top             =   0
         Width           =   3615
         Begin VB.TextBox txtClientID 
            Height          =   315
            Left            =   1020
            TabIndex        =   95
            Top             =   240
            Width           =   2415
         End
         Begin VB.Label lblClientID 
            Caption         =   "C&lient ID:"
            Height          =   255
            Left            =   120
            TabIndex        =   94
            Top             =   270
            Width           =   735
         End
      End
   End
   Begin VB.Frame fraPhoton 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4335
      Left            =   2280
      TabIndex        =   1
      Top             =   120
      Width           =   3675
      Begin VB.Frame fraServer 
         Caption         =   "Server Information"
         Height          =   1395
         Left            =   0
         TabIndex        =   7
         Top             =   1200
         Width           =   3615
         Begin VB.TextBox txtPhotonEuropePort 
            Height          =   315
            Left            =   2640
            TabIndex        =   17
            Top             =   990
            Width           =   795
         End
         Begin VB.TextBox txtPhotonEuropeIP 
            Height          =   285
            Left            =   480
            TabIndex        =   15
            Top             =   1005
            Width           =   1515
         End
         Begin VB.TextBox txtPhotonChicagoIP 
            Height          =   285
            Left            =   480
            TabIndex        =   10
            Top             =   465
            Width           =   1515
         End
         Begin VB.TextBox txtPhotonChicagoPort 
            Height          =   315
            Left            =   2640
            TabIndex        =   12
            Top             =   450
            Width           =   795
         End
         Begin VB.Label lblPhotonLiffeEurex 
            Caption         =   "Eurex/Liffe:"
            Height          =   195
            Left            =   180
            TabIndex        =   13
            Top             =   780
            Width           =   915
         End
         Begin VB.Label lblPhotonEuropePort 
            Caption         =   "Por&t:"
            Height          =   255
            Left            =   2160
            TabIndex        =   16
            Top             =   1020
            Width           =   375
         End
         Begin VB.Label lblPhotonEuropeIP 
            Caption         =   "I&P:"
            Height          =   255
            Left            =   180
            TabIndex        =   14
            Top             =   1020
            Width           =   255
         End
         Begin VB.Label lblPhotonCmeCbot 
            Caption         =   "CME/CBOT:"
            Height          =   195
            Left            =   180
            TabIndex        =   8
            Top             =   240
            Width           =   915
         End
         Begin VB.Label lblPhotonChicagoIP 
            Caption         =   "&IP:"
            Height          =   255
            Left            =   180
            TabIndex        =   9
            Top             =   480
            Width           =   255
         End
         Begin VB.Label lblPhotonChicagoPort 
            Caption         =   "Po&rt:"
            Height          =   255
            Left            =   2160
            TabIndex        =   11
            Top             =   480
            Width           =   375
         End
      End
      Begin VB.Frame fraSubnodes 
         Caption         =   "SubNode Information"
         Height          =   1575
         Left            =   0
         TabIndex        =   18
         Top             =   2700
         Width           =   3615
         Begin VB.TextBox txtInterval 
            Height          =   315
            Left            =   2640
            TabIndex        =   30
            Top             =   1140
            Width           =   795
         End
         Begin VB.TextBox txtEurex 
            Height          =   315
            Left            =   2640
            TabIndex        =   26
            Top             =   720
            Width           =   795
         End
         Begin VB.TextBox txtECBOT 
            Height          =   315
            Left            =   2640
            TabIndex        =   22
            Top             =   300
            Width           =   795
         End
         Begin VB.TextBox txtData 
            Height          =   315
            Left            =   840
            TabIndex        =   28
            Top             =   1140
            Width           =   795
         End
         Begin VB.TextBox txtLiffe 
            Height          =   315
            Left            =   840
            TabIndex        =   24
            Top             =   720
            Width           =   795
         End
         Begin VB.TextBox txtCME 
            Height          =   315
            Left            =   840
            TabIndex        =   20
            Top             =   300
            Width           =   795
         End
         Begin VB.CheckBox chkData 
            Caption         =   "&Data"
            Height          =   195
            Left            =   180
            TabIndex        =   27
            Top             =   1200
            Width           =   675
         End
         Begin VB.CheckBox chkLiffe 
            Caption         =   "&Liffe"
            Height          =   195
            Left            =   180
            TabIndex        =   23
            Top             =   780
            Width           =   675
         End
         Begin VB.CheckBox chkCME 
            Caption         =   "C&ME"
            Height          =   195
            Left            =   180
            TabIndex        =   19
            Top             =   360
            Width           =   675
         End
         Begin VB.CheckBox chkEurex 
            Caption         =   "&Eurex"
            Height          =   195
            Left            =   1860
            TabIndex        =   25
            Top             =   780
            Width           =   735
         End
         Begin VB.CheckBox chkCBOT 
            Caption         =   "C&BOT"
            Height          =   195
            Left            =   1860
            TabIndex        =   21
            Top             =   360
            Width           =   735
         End
         Begin VB.Label lblInterval 
            Caption         =   "I&nterval:"
            Height          =   255
            Left            =   2040
            TabIndex        =   29
            Top             =   1170
            Width           =   615
         End
      End
      Begin VB.Frame fraUserInformation 
         Caption         =   "User Information"
         Height          =   1155
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   3615
         Begin VB.TextBox txtUserID 
            Height          =   315
            Left            =   1020
            TabIndex        =   4
            Top             =   240
            Width           =   2415
         End
         Begin VB.TextBox txtPassword 
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   1020
            PasswordChar    =   "*"
            TabIndex        =   6
            Top             =   630
            Width           =   2415
         End
         Begin VB.Label lblUserID 
            Caption         =   "&User ID:"
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   270
            Width           =   615
         End
         Begin VB.Label lblPassword 
            Caption         =   "Pass&word:"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   660
            Width           =   855
         End
      End
   End
   Begin VB.Frame fraButtons 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1035
      Left            =   6120
      TabIndex        =   119
      Top             =   120
      Width           =   1215
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   495
         Left            =   0
         TabIndex        =   121
         Top             =   540
         Width           =   1215
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   495
         Left            =   0
         TabIndex        =   120
         Top             =   0
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmBrokerCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmBrokerCfg.frm
'' Description: Allow the user to configure their Photon connection settings
''
'' Author:      Genesis Financial Data Services
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 12/10/2010   DAJ         Changed over to the IsBrokerUser function
'' 08/25/2011   DAJ         Removed references to TT
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    bOK As Boolean                      ' Did the user click OK?
    bStarting As Boolean                ' Starting the ShowMe
    nAcctType As eTT_AccountType        ' Account type passed in
    
    strPatsInfo As String               ' PATS information
    strManLondonInfo As String          ' Man Financial - London information
    strManChicagoInfo As String         ' Man Financial - Chicago information
End Type
Private m As mPrivate

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Initialize and show the form
'' Inputs:      Mode
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(Optional ByVal nStart As eTT_AccountType = -1&, Optional ByVal bReconnect As Boolean = True) As Boolean
On Error GoTo ErrSection:

    Dim strIntBrokersIni As String      ' Interactive Brokers ini file
    Dim strPatsIni As String            ' Pats ini file
    Dim strPhotonIni As String          ' Photon ini file
    Dim strLindWaldockIni As String     ' LindWaldock ini file
    Dim strManLondonIni As String       ' Man Financial - London ini file
    Dim strManChicagoIni As String      ' Man Financial - Chicago ini file
    Dim strAlaronIni As String          ' Alaron ini file
    Dim strTransactIni As String        ' TransAct ini file
    Dim strPfgIni As String             ' PFG ini file
    Dim strFxcmIni As String            ' FXCM ini file
    Dim strTtIni As String              ' TT ini file
    Dim lTimeOut As Long                ' Timeout value
    Dim strPassword As String           ' Password from the ini file
    Dim bVisible As Boolean
    
    m.bStarting = True
    m.nAcctType = nStart
    
    strIntBrokersIni = AddSlash(App.Path) & "IntBrokers.INI"
    strPatsIni = AddSlash(App.Path) & "Pats.INI"
    strPhotonIni = AddSlash(App.Path) & "Photon.INI"
    strLindWaldockIni = AddSlash(App.Path) & "LindWaldock.INI"
    strManLondonIni = AddSlash(App.Path) & "ManLondon.INI"
    strManChicagoIni = AddSlash(App.Path) & "ManChicago.INI"
    strAlaronIni = AddSlash(App.Path) & "Alrn.INI"
    strTransactIni = AddSlash(App.Path) & "TransAct.INI"
    strPfgIni = AddSlash(App.Path) & "Pfg.INI"
    strFxcmIni = AddSlash(App.Path) & "Fxcm.INI"
    strTtIni = AddSlash(App.Path) & "TradingTechnologies.INI"
    
    If g.Broker.IsBrokerUser(eTT_AccountType_Photon) And DoBroker(eTT_AccountType_Photon) Then
        txtUserID.Text = GetIniFileProperty("UserID", "", "User", strPhotonIni)
        txtPassword.Text = GetIniFileProperty("Password", "", "User", strPhotonIni)
        If Len(txtPassword.Text) > 0 Then
            SetIniFileProperty "Password", "", "User", strPhotonIni
            SetIniFileProperty "Password2", EncryptToHex(txtPassword.Text), "User", strPhotonIni
        Else
            txtPassword.Text = DecryptFromHex(GetIniFileProperty("Password2", "", "User", strPhotonIni))
        End If
        txtPhotonChicagoIP.Text = GetIniFileProperty("ServerIP", "198.63.193.40", "Server", strPhotonIni)
        txtPhotonChicagoPort.Text = GetIniFileProperty("ServerPort", "9030", "Server", strPhotonIni)
        txtPhotonEuropeIP.Text = GetIniFileProperty("ServerIPEurope", "209.120.215.41", "Server", strPhotonIni)
        txtPhotonEuropePort.Text = GetIniFileProperty("ServerPortEurope", "9030", "Server", strPhotonIni)
        chkData.Value = GetIniFileProperty("Market", vbChecked, "Connect", strPhotonIni)
        txtData.Text = GetIniFileProperty("Market", "7247", "SubNodes", strPhotonIni)
        chkCME.Value = GetIniFileProperty("CME", vbChecked, "Connect", strPhotonIni)
        txtCME.Text = GetIniFileProperty("CME", "18249", "SubNodes", strPhotonIni)
        chkCBOT.Value = GetIniFileProperty("eCBOT", vbChecked, "Connect", strPhotonIni)
        txtECBOT.Text = GetIniFileProperty("eCBOT", "18217", "SubNodes", strPhotonIni)
        chkEurex.Value = GetIniFileProperty("Eurex", vbUnchecked, "Connect", strPhotonIni)
        txtEurex.Text = GetIniFileProperty("Eurex", "18263", "SubNodes", strPhotonIni)
        chkLiffe.Value = GetIniFileProperty("Liffe", vbUnchecked, "Connect", strPhotonIni)
        txtLiffe.Text = GetIniFileProperty("Liffe", "12874", "SubNodes", strPhotonIni)
        txtInterval.Text = GetIniFileProperty("Interval", 2000&, "Simulate", strPhotonIni)
                
        'chkData.Visible = False
        'txtData.Visible = False
        bVisible = False
        If g.nReplaySession > 0 Then bVisible = True
        lblInterval.Visible = bVisible
        txtInterval.Visible = bVisible
    End If
            
    If g.Broker.IsBrokerUser(eTT_AccountType_IntBrokers) And DoBroker(eTT_AccountType_IntBrokers) Then
        txtHostIP.Text = GetIniFileProperty("HostIP", "", "Connection", strIntBrokersIni)
        txtPort.Text = GetIniFileProperty("Port", "", "Connection", strIntBrokersIni)
        txtClientID.Text = GetIniFileProperty("ClientID", "", "User", strIntBrokersIni)
    End If
            
    If g.Broker.IsBrokerUser(eTT_AccountType_LindWaldock) And DoBroker(eTT_AccountType_LindWaldock) Then
        txtLWUserID.Text = GetIniFileProperty("UserID", "", "User", strLindWaldockIni)
        txtLWPassword.Text = GetIniFileProperty("Password", "", "User", strLindWaldockIni)
        If Len(txtLWPassword.Text) > 0 Then
            SetIniFileProperty "Password", "", "User", strLindWaldockIni
            SetIniFileProperty "Password2", EncryptToHex(txtLWPassword.Text), "User", strLindWaldockIni
        Else
            txtLWPassword.Text = DecryptFromHex(GetIniFileProperty("Password2", "", "User", strLindWaldockIni))
        End If
        txtLWHost.Text = GetIniFileProperty("Host", "", "User", strLindWaldockIni)
        txtLWPort.Text = GetIniFileProperty("Port", "", "User", strLindWaldockIni)
        txtLWFirm.Text = GetIniFileProperty("Firm", "", "User", strLindWaldockIni)
        txtLWTestNum.Text = GetIniFileProperty("TestNum", "", "User", strLindWaldockIni)
    End If
            
    If g.Broker.IsBrokerUser(eTT_AccountType_PATS) And DoBroker(eTT_AccountType_PATS) Then
        m.strPatsInfo = GetIniFileProperty("UserID", "", "User", strPatsIni) & ";"
        strPassword = GetIniFileProperty("Password", "", "User", strPatsIni)
        If Len(strPassword) > 0 Then
            SetIniFileProperty "Password", "", "User", strPatsIni
            SetIniFileProperty "Password2", EncryptToHex(strPassword), "User", strPatsIni
        Else
            strPassword = DecryptFromHex(GetIniFileProperty("Password2", "", "User", strPatsIni))
        End If
        m.strPatsInfo = m.strPatsInfo & strPassword & ";"
        m.strPatsInfo = m.strPatsInfo & GetIniFileProperty("HostIP", "", "Start", strPatsIni) & ";"
        m.strPatsInfo = m.strPatsInfo & GetIniFileProperty("HostPort", "", "Start", strPatsIni) & ";"
        m.strPatsInfo = m.strPatsInfo & GetIniFileProperty("UseSuperTAS", "N", "Start", strPatsIni)
    End If
    
    If g.Broker.IsBrokerUser(eTT_AccountType_ManLondon) And DoBroker(eTT_AccountType_ManLondon) Then
        m.strManLondonInfo = GetIniFileProperty("UserID", "", "User", strManLondonIni) & ";"
        strPassword = GetIniFileProperty("Password", "", "User", strManLondonIni)
        If Len(strPassword) > 0 Then
            SetIniFileProperty "Password", "", "User", strManLondonIni
            SetIniFileProperty "Password2", EncryptToHex(strPassword), "User", strManLondonIni
        Else
            strPassword = DecryptFromHex(GetIniFileProperty("Password2", "", "User", strManLondonIni))
        End If
        m.strManLondonInfo = m.strManLondonInfo & strPassword & ";"
        m.strManLondonInfo = m.strManLondonInfo & GetIniFileProperty("HostIP", "", "Start", strManLondonIni) & ";"
        m.strManLondonInfo = m.strManLondonInfo & GetIniFileProperty("HostPort", "", "Start", strManLondonIni) & ";"
        m.strManLondonInfo = m.strManLondonInfo & GetIniFileProperty("UseSuperTAS", "N", "Start", strManLondonIni)
    End If
    
    If g.Broker.IsBrokerUser(eTT_AccountType_ManChicago) And DoBroker(eTT_AccountType_ManChicago) Then
        m.strManChicagoInfo = GetIniFileProperty("UserID", "", "User", strManChicagoIni) & ";"
        strPassword = GetIniFileProperty("Password", "", "User", strManChicagoIni)
        If Len(strPassword) > 0 Then
            SetIniFileProperty "Password", "", "User", strManChicagoIni
            SetIniFileProperty "Password2", EncryptToHex(strPassword), "User", strManChicagoIni
        Else
            strPassword = DecryptFromHex(GetIniFileProperty("Password2", "", "User", strManChicagoIni))
        End If
        m.strManChicagoInfo = m.strManChicagoInfo & strPassword & ";"
        m.strManChicagoInfo = m.strManChicagoInfo & GetIniFileProperty("HostIP", "", "Start", strManChicagoIni) & ";"
        m.strManChicagoInfo = m.strManChicagoInfo & GetIniFileProperty("HostPort", "", "Start", strManChicagoIni) & ";"
        m.strManChicagoInfo = m.strManChicagoInfo & GetIniFileProperty("UseSuperTAS", "N", "Start", strManChicagoIni)
    End If
    
    If g.Broker.IsBrokerUser(eTT_AccountType_Alaron) And DoBroker(eTT_AccountType_Alaron) Then
        txtAlaronUserID.Text = GetIniFileProperty("UserID", "", "User", strAlaronIni)
        txtAlaronPassword.Text = GetIniFileProperty("Password", "", "User", strAlaronIni)
        If Len(txtAlaronPassword.Text) > 0 Then
            SetIniFileProperty "Password", "", "User", strAlaronIni
            SetIniFileProperty "Password2", EncryptToHex(txtAlaronPassword.Text), "User", strAlaronIni
        Else
            txtAlaronPassword.Text = DecryptFromHex(GetIniFileProperty("Password2", "", "User", strAlaronIni))
        End If
        txtAlaronIP.Text = GetIniFileProperty("IP", "", "User", strAlaronIni)
        txtAlaronPort.Text = GetIniFileProperty("Port", "", "User", strAlaronIni)
        
        If Len(Trim(txtAlaronIP.Text)) = 0 Then
            txtAlaronIP.Text = GetIniFileProperty("Live", "", "IP", AddSlash(App.Path) & "Provided\AlrIps.INI")
            txtAlaronPort.Text = GetIniFileProperty("Live", "", "Port", AddSlash(App.Path) & "Provided\AlrIps.INI")
        End If
    End If
    
    If g.Broker.IsBrokerUser(eTT_AccountType_TransAct) And DoBroker(eTT_AccountType_TransAct) Then
        txtTransactUserID.Text = GetIniFileProperty("UserID", "", "User", strTransactIni)
        txtTransactPassword.Text = GetIniFileProperty("Password", "", "User", strTransactIni)
        If Len(txtTransactPassword.Text) > 0 Then
            SetIniFileProperty "Password", "", "User", strTransactIni
            SetIniFileProperty "Password2", EncryptToHex(txtTransactPassword.Text), "User", strTransactIni
        Else
            txtTransactPassword.Text = DecryptFromHex(GetIniFileProperty("Password2", "", "User", strTransactIni))
        End If
        txtTransactIP.Text = GetIniFileProperty("HostIP", "", "User", strTransactIni)
        If Len(txtTransactIP.Text) = 0 Then
            txtTransactIP.Text = GetIniFileProperty("IP", "mt13.yorkba.com", "User", strTransactIni)
            If txtTransactIP.Text = "mt17.yorkba.com" Then
                txtTransactIP.Text = "mt13.yorkba.com"
            End If
        End If
        
        FillTransActSymbols
    End If
    
    If g.Broker.IsBrokerUser(eTT_AccountType_PFG) And DoBroker(eTT_AccountType_PFG) Then
        txtPfgUserID.Text = GetIniFileProperty("UserID", "", "User", strPfgIni)
        txtPfgPassword.Text = DecryptFromHex(GetIniFileProperty("Password", "", "User", strPfgIni))
        txtPfgIP.Text = GetIniFileProperty("HostIP", "", "User", strPfgIni)
        txtPfgPort.Text = GetIniFileProperty("HostPort", "", "User", strPfgIni)
        
        If Len(Trim(txtPfgIP.Text)) = 0 Then
            If (Len(Trim(txtPfgUserID.Text)) = 0) Or (Left(UCase(txtPfgUserID.Text), 1) = "D") Then
                txtPfgIP.Text = "12.36.73.195"
                txtPfgPort.Text = "2000"
            Else
                txtPfgIP.Text = "12.36.73.184"
                txtPfgPort.Text = "2000"
            End If
        End If
    End If
    
    If g.Broker.IsBrokerUser(eTT_AccountType_FXCM) And DoBroker(eTT_AccountType_FXCM) Then
        txtFxcmUserID.Text = GetIniFileProperty("UserID", "", "User", strFxcmIni)
        txtFxcmPassword.Text = DecryptFromHex(GetIniFileProperty("Password", "", "User", strFxcmIni))
        txtFxcmUrl.Text = GetIniFileProperty("URL", "", "User", strFxcmIni)
    End If
    
    SelectBroker nStart
    lstBrokers_Click
    If m.nAcctType > -1& Then
        lstBrokers.Visible = False
        Form_Resize
        Caption = lstBrokers.List(lstBrokers.ListIndex) & " Connection Configuration"
    End If

    SizeForm
    m.bStarting = False
    ShowForm Me, eForm_Modal, frmMain
    
    If m.bOK Then
        Select Case UCase(lstBrokers.Text)
            Case "PATS", "MAN FINANCIAL - LONDON", "MAN FINANCIAL - CHICAGO"
                ControlsToPatsString UCase(lstBrokers.Text)
        End Select
        
        If g.Broker.IsBrokerUser(eTT_AccountType_Photon) And DoBroker(eTT_AccountType_Photon) = True Then
            SetIniFileProperty "UserID", txtUserID.Text, "User", strPhotonIni
            SetIniFileProperty "Password2", EncryptToHex(txtPassword.Text), "User", strPhotonIni
            SetIniFileProperty "ServerIP", txtPhotonChicagoIP.Text, "Server", strPhotonIni
            SetIniFileProperty "ServerPort", txtPhotonChicagoPort.Text, "Server", strPhotonIni
            SetIniFileProperty "ServerIPEurope", txtPhotonEuropeIP.Text, "Server", strPhotonIni
            SetIniFileProperty "ServerPortEurope", txtPhotonEuropePort.Text, "Server", strPhotonIni
            SetIniFileProperty "Market", chkData.Value, "Connect", strPhotonIni
            SetIniFileProperty "Market", txtData.Text, "SubNodes", strPhotonIni
            SetIniFileProperty "CME", chkCME.Value, "Connect", strPhotonIni
            SetIniFileProperty "CME", txtCME.Text, "SubNodes", strPhotonIni
            SetIniFileProperty "eCBOT", chkCBOT.Value, "Connect", strPhotonIni
            SetIniFileProperty "eCBOT", txtECBOT.Text, "SubNodes", strPhotonIni
            SetIniFileProperty "Eurex", chkEurex.Value, "Connect", strPhotonIni
            SetIniFileProperty "Eurex", txtEurex.Text, "SubNodes", strPhotonIni
            SetIniFileProperty "Liffe", chkLiffe.Value, "Connect", strPhotonIni
            SetIniFileProperty "Liffe", txtLiffe.Text, "SubNodes", strPhotonIni
            SetIniFileProperty "Interval", txtInterval.Text, "Simulate", strPhotonIni
                    
            If (Not g.Photon Is Nothing) Then
                DoEvents
                g.Photon.RefreshConnections bReconnect
            End If
        End If
            
        If g.Broker.IsBrokerUser(eTT_AccountType_IntBrokers) And DoBroker(eTT_AccountType_IntBrokers) = True Then
            SetIniFileProperty "HostIP", txtHostIP.Text, "Connection", strIntBrokersIni
            SetIniFileProperty "Port", txtPort.Text, "Connection", strIntBrokersIni
            SetIniFileProperty "ClientID", txtClientID.Text, "User", strIntBrokersIni
        End If
                
        If g.Broker.IsBrokerUser(eTT_AccountType_LindWaldock) And DoBroker(eTT_AccountType_LindWaldock) = True Then
            SetIniFileProperty "UserID", txtLWUserID.Text, "User", strLindWaldockIni
            SetIniFileProperty "Password2", EncryptToHex(txtLWPassword.Text), "User", strLindWaldockIni
            SetIniFileProperty "Host", txtLWHost.Text, "User", strLindWaldockIni
            SetIniFileProperty "Port", txtLWPort.Text, "User", strLindWaldockIni
            SetIniFileProperty "Firm", txtLWFirm.Text, "User", strLindWaldockIni
            SetIniFileProperty "TestNum", txtLWTestNum.Text, "User", strLindWaldockIni
        
            If bReconnect Then
                lTimeOut = 0&
                
                g.LindWaldock.Disconnect
                Do While (g.LindWaldock.ConnectionStatus <> eGDConnectionStatus_Disconnected) And (lTimeOut < 30&)
                    Sleep 1
                    lTimeOut = lTimeOut + 1&
                Loop
                If LiveTradingAllowed(eTT_AccountType_LindWaldock) Then
                    If g.LindWaldock.ConnectionStatus = eGDConnectionStatus_Disconnected Then
                        g.LindWaldock.Connect
                    End If
                End If
            End If
        End If
                
        If g.Broker.IsBrokerUser(eTT_AccountType_PATS) And DoBroker(eTT_AccountType_PATS) = True Then
            SetIniFileProperty "UserID", Parse(m.strPatsInfo, ";", 1), "User", strPatsIni
            SetIniFileProperty "Password2", EncryptToHex(Parse(m.strPatsInfo, ";", 2)), "User", strPatsIni
            SetIniFileProperty "HostIP", Parse(m.strPatsInfo, ";", 3), "Start", strPatsIni
            SetIniFileProperty "HostPort", Parse(m.strPatsInfo, ";", 4), "Start", strPatsIni
            SetIniFileProperty "UseSuperTAS", Parse(m.strPatsInfo, ";", 5), "Start", strPatsIni
        End If
    
        If g.Broker.IsBrokerUser(eTT_AccountType_ManLondon) And DoBroker(eTT_AccountType_ManLondon) = True Then
            SetIniFileProperty "UserID", Parse(m.strManLondonInfo, ";", 1), "User", strManLondonIni
            SetIniFileProperty "Password2", EncryptToHex(Parse(m.strManLondonInfo, ";", 2)), "User", strManLondonIni
            SetIniFileProperty "HostIP", Parse(m.strManLondonInfo, ";", 3), "Start", strManLondonIni
            SetIniFileProperty "HostPort", Parse(m.strManLondonInfo, ";", 4), "Start", strManLondonIni
            SetIniFileProperty "UseSuperTAS", Parse(m.strManLondonInfo, ";", 5), "Start", strManLondonIni
        End If
    
        If g.Broker.IsBrokerUser(eTT_AccountType_ManChicago) And DoBroker(eTT_AccountType_ManChicago) = True Then
            SetIniFileProperty "UserID", Parse(m.strManChicagoInfo, ";", 1), "User", strManChicagoIni
            SetIniFileProperty "Password2", EncryptToHex(Parse(m.strManChicagoInfo, ";", 2)), "User", strManChicagoIni
            SetIniFileProperty "HostIP", Parse(m.strManChicagoInfo, ";", 3), "Start", strManChicagoIni
            SetIniFileProperty "HostPort", Parse(m.strManChicagoInfo, ";", 4), "Start", strManChicagoIni
            SetIniFileProperty "UseSuperTAS", Parse(m.strManChicagoInfo, ";", 5), "Start", strManChicagoIni
        End If
        
        If g.Broker.IsBrokerUser(eTT_AccountType_Alaron) And DoBroker(eTT_AccountType_Alaron) = True Then
            SetIniFileProperty "UserID", Trim(txtAlaronUserID.Text), "User", strAlaronIni
            g.Alaron.UserName = Trim(txtAlaronUserID.Text)
            SetIniFileProperty "Password2", EncryptToHex(Trim(txtAlaronPassword.Text)), "User", strAlaronIni
            g.Alaron.Password = Trim(txtAlaronPassword.Text)
            SetIniFileProperty "IP", Trim(txtAlaronIP.Text), "User", strAlaronIni
            g.Alaron.IPAddress = Trim(txtAlaronIP.Text)
            SetIniFileProperty "Port", Trim(txtAlaronPort.Text), "User", strAlaronIni
            g.Alaron.Port = Trim(txtAlaronPort.Text)
            
            If bReconnect Then
                lTimeOut = 0&
                g.Alaron.Disconnect
                Do While (g.Alaron.ConnectionStatus <> eGDConnectionStatus_Disconnected) And (lTimeOut < 30&)
                    Sleep 1
                    lTimeOut = lTimeOut + 1&
                Loop
                If LiveTradingAllowed(eTT_AccountType_Alaron) Then
                    If g.Alaron.ConnectionStatus = eGDConnectionStatus_Disconnected Then
                        g.Alaron.Connect
                    End If
                End If
            End If
        End If
        
        If g.Broker.IsBrokerUser(eTT_AccountType_TransAct) And DoBroker(eTT_AccountType_TransAct) = True Then
            SetIniFileProperty "UserID", Trim(txtTransactUserID.Text), "User", strTransactIni
            g.Transact.UserName = Trim(txtTransactUserID.Text)
            SetIniFileProperty "Password2", EncryptToHex(Trim(txtTransactPassword.Text)), "User", strTransactIni
            g.Transact.Password = Trim(txtTransactPassword.Text)
            SetIniFileProperty "HostIP", Trim(txtTransactIP.Text), "User", strTransactIni
            g.Transact.IPAddress = Trim(txtTransactIP.Text)
            
            If bReconnect Then
                lTimeOut = 0&
                
                InfBox "Disconnecting from TransAct servers.|Please wait...", , , "TransAct Disconnect", True
                g.Transact.Disconnect "", False, "Login Information Changed"
                'Do While (g.Transact.ConnectionStatus <> eGDConnectionStatus_Disconnected) And (lTimeOut < 30&)
                Do While (g.Transact.AppLoaded = True) And (lTimeOut < 30&)
                    Sleep 1
                    lTimeOut = lTimeOut + 1&
                Loop
                If LiveTradingAllowed(eTT_AccountType_TransAct) Then
                    g.Transact.Account = ""
                    If g.Transact.ConnectionStatus = eGDConnectionStatus_Disconnected Then
                        InfBox "Connecting to TransAct servers.|Please wait...", , , "TransAct Connect", True
                        g.Transact.Connect
                    End If
                End If
                InfBox ""
            End If
        End If
        
        If g.Broker.IsBrokerUser(eTT_AccountType_PFG) And DoBroker(eTT_AccountType_PFG) Then
            SetIniFileProperty "UserID", Trim(txtPfgUserID.Text), "User", strPfgIni
            g.PFG.UserName = Trim(txtPfgUserID.Text)
            SetIniFileProperty "Password", EncryptToHex(Trim(txtPfgPassword.Text)), "User", strPfgIni
            g.PFG.Password = Trim(txtPfgPassword.Text)
            SetIniFileProperty "HostIP", Trim(txtPfgIP.Text), "User", strPfgIni
            g.PFG.HostIP = Trim(txtPfgIP.Text)
            SetIniFileProperty "HostPort", Trim(txtPfgPort.Text), "User", strPfgIni
            g.PFG.HostPort = Trim(txtPfgPort.Text)
            
            If bReconnect Then
                lTimeOut = 0&
                
                If (g.PFG.ConnectionStatus <> eGDConnectionStatus_Disconnected) Then
                    InfBox "Disconnecting from PFG servers.|Please wait...", , , "PFG Disconnect", True
                    g.PFG.Disconnect False, "Changing User Information"
                    Do While (g.PFG.ConnectionStatus <> eGDConnectionStatus_Disconnected) And (lTimeOut < 30&)
                        Sleep 1
                        lTimeOut = lTimeOut + 1&
                    Loop
                End If
                If LiveTradingAllowed(eTT_AccountType_PFG) Then
                    If g.PFG.ConnectionStatus = eGDConnectionStatus_Disconnected Then
                        g.PFG.Connect
                        InfBox "Connecting to PFG servers.|Please wait...", , , "PFG Connect", True
                    End If
                End If
                InfBox ""
            End If
        End If
        
        If g.Broker.IsBrokerUser(eTT_AccountType_FXCM) And DoBroker(eTT_AccountType_FXCM) Then
            SetIniFileProperty "UserID", Trim(txtFxcmUserID.Text), "User", strFxcmIni
            g.FXCM.UserName = Trim(txtFxcmUserID.Text)
            SetIniFileProperty "Password", EncryptToHex(Trim(txtFxcmPassword.Text)), "User", strFxcmIni
            g.FXCM.Password = Trim(txtFxcmPassword.Text)
            SetIniFileProperty "URL", Trim(txtFxcmUrl.Text), "User", strFxcmIni
            g.FXCM.URL = Trim(txtFxcmUrl.Text)
            
            If bReconnect Then
                lTimeOut = 0&
                
                If g.FXCM.ConnectionStatus <> eGDConnectionStatus_Disconnected Then
                    InfBox "Disconnecting from FXCM servers|Please wait...", , , "FXCM Disconnect", True
                    g.FXCM.Disconnect False, "Changing User Information"
                    Do While (g.FXCM.ConnectionStatus <> eGDConnectionStatus_Disconnected) And (lTimeOut < 30&)
                        Sleep 1
                        lTimeOut = lTimeOut + 1&
                    Loop
                End If
                
                If LiveTradingAllowed(eTT_AccountType_FXCM) Then
                    If g.FXCM.ConnectionStatus = eGDConnectionStatus_Disconnected Then
                        g.FXCM.Connect
                        InfBox "Connecting to FXCM servers|Please wait...", , , "FXCM Connect", True
                    End If
                End If
                InfBox ""
            End If
        End If
    End If
        
    ShowMe = m.bOK

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmBrokerCfg.ShowMe", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: Set the flag to allow ShowMe to unload
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancel_Click()
On Error GoTo ErrSection:

    m.bOK = False
    Me.Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerCfg.cmdCancel.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOK_Click
'' Description: Set the flag to allow ShowMe to save information before unload
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
On Error GoTo ErrSection:

    MoveFocus cmdOK
    
    Select Case m.nAcctType
        Case eTT_AccountType_PFG
            If Left(txtPfgUserID.Text, 1) <> "D" Then
                If Not FileExist(AddSlash(App.Path) & "PfgLive.FLG") Then
                    MoveFocus txtPfgUserID
                    InfBox "You are not currently authorized to trade a live PFG account with Trade Navigator", "!", , "PFG Account Error"
                    GoTo ErrExit
                End If
            End If
    
        Case eTT_AccountType_TransAct
            If (UCase(Trim(txtTransactUserID.Text)) = UCase(g.Transact.SimUserUserName)) Then
                MoveFocus txtTransactUserID
                Err.Raise vbObjectError + 1000, , "Please enter in your broker-provided TransAct User Name.  Please call your broker for more information."
                GoTo ErrExit
            End If
            If UCase(Trim(txtTransactIP.Text)) = "SIMMT.YORKBA.COM" Then
                MoveFocus txtTransactUserID
                Err.Raise vbObjectError + 1000, , "Please enter in the live TransAct IP address"
                GoTo ErrExit
            End If
            If IsDigit(Left(Trim(txtTransactUserID.Text), 1)) = True Then
                MoveFocus txtTransactUserID
                Err.Raise vbObjectError + 1000, , "TransAct User Name cannot start with a digit.  Please contact your broker for more information."
                GoTo ErrExit
            End If
    End Select

    m.bOK = True
    Me.Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerCfg.cmdOK.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Activate
'' Description: When the form is activated, show the right stuff for the mode
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Activate()
On Error GoTo ErrSection:

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerCfg.Form.Activate", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Initialize the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Icon = Picture16(ToolbarIcon("ID_TradeTracker"))
    CenterTheForm Me
    cmdCancel.Cancel = True
    
    If g.Broker.IsBrokerUser(eTT_AccountType_IntBrokers) Then
        lstBrokers.AddItem "Interactive Brokers"
        lstBrokers.ItemData(lstBrokers.NewIndex) = eTT_AccountType_IntBrokers
    End If
    If g.Broker.IsBrokerUser(eTT_AccountType_PATS) Then
        lstBrokers.AddItem "PATS"
        lstBrokers.ItemData(lstBrokers.NewIndex) = eTT_AccountType_PATS
    End If
    If g.Broker.IsBrokerUser(eTT_AccountType_Photon) Then
        lstBrokers.AddItem "Photon"
        lstBrokers.ItemData(lstBrokers.NewIndex) = eTT_AccountType_Photon
    End If
    If g.Broker.IsBrokerUser(eTT_AccountType_LindWaldock) Then
        lstBrokers.AddItem g.Broker.BrokerName(eTT_AccountType_LindWaldock)
        lstBrokers.ItemData(lstBrokers.NewIndex) = eTT_AccountType_LindWaldock
    End If
    If g.Broker.IsBrokerUser(eTT_AccountType_ManLondon) Then
        lstBrokers.AddItem "Man Financial - London"
        lstBrokers.ItemData(lstBrokers.NewIndex) = eTT_AccountType_ManLondon
    End If
    If g.Broker.IsBrokerUser(eTT_AccountType_ManChicago) Then
        lstBrokers.AddItem "Man Financial - Chicago"
        lstBrokers.ItemData(lstBrokers.NewIndex) = eTT_AccountType_ManChicago
    End If
    If g.Broker.IsBrokerUser(eTT_AccountType_Alaron) Then
        lstBrokers.AddItem "Alaron"
        lstBrokers.ItemData(lstBrokers.NewIndex) = eTT_AccountType_Alaron
    End If
    If g.Broker.IsBrokerUser(eTT_AccountType_TransAct) Then
        lstBrokers.AddItem "TransAct"
        lstBrokers.ItemData(lstBrokers.NewIndex) = eTT_AccountType_TransAct
    End If
    If g.Broker.IsBrokerUser(eTT_AccountType_PFG) Then
        lstBrokers.AddItem "PFG"
        lstBrokers.ItemData(lstBrokers.NewIndex) = eTT_AccountType_PFG
    End If
    If g.Broker.IsBrokerUser(eTT_AccountType_FXCM) Then
        lstBrokers.AddItem "FXCM"
        lstBrokers.ItemData(lstBrokers.NewIndex) = eTT_AccountType_FXCM
    End If
    
    If lstBrokers.ListCount = 1 Then
        lstBrokers.Visible = False
        Caption = lstBrokers.List(0) & " Configuration"
    Else
        Caption = "On-Line Broker Configuration"
    End If
    
    txtTransactSymbols.Enabled = True
    txtTransactSymbols.BackColor = cmdOK.BackColor
    txtTransactSymbols.Locked = True
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerCfg.Form.Load", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: Make sure that we close down correctly
'' Inputs:      Whether to Cancel the Unload, Mode of the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode <> vbFormCode Then
        m.bOK = False
        Cancel = True
        Me.Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerCfg.Form.QueryUnload", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Resize
'' Description: Move and size controls as the form is resized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Resize()
On Error Resume Next

    Dim lLeft As Long                   ' Left of the frame
    
    If (lstBrokers.ListCount > 1) And (m.nAcctType = -1&) Then
        lLeft = lstBrokers.Width + 120
    Else
        lLeft = 60
    End If

    With lstBrokers
        .Move 60, 60, .Width, ScaleHeight - 120
    End With
    
    With fraIntBrokers
        .Move lLeft, 60
    End With
    
    With fraPats
        .Move lLeft, 60
    End With
    
    With fraPhoton
        .Move lLeft, 60
    End With
    
    With fraLW
        .Move lLeft, 60
    End With
    
    With fraAlaron
        .Move lLeft, 60
    End With
    
    With fraTransact
        .Move lLeft, 60
    End With
    
    With fraPfg
        .Move lLeft, 60
    End With
    
    fraFXCM.Move lLeft, 60
    
    With fraTT
        .Move lLeft, 60
    End With

    With fraButtons
        .Move ScaleWidth - .Width - 60, 60
    End With
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    lstBrokers_Click
'' Description: Show the appropriate frame based on the broker the user chose
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub lstBrokers_Click()
On Error GoTo ErrSection:

    Static strLastBroker As String      ' Last broker selected

    If m.bStarting = True Then strLastBroker = ""

    Select Case UCase(strLastBroker)
        Case "PATS", "MAN FINANCIAL - LONDON", "MAN FINANCIAL - CHICAGO"
            ControlsToPatsString UCase(strLastBroker)
    End Select
    
    Select Case UCase(lstBrokers.Text)
        Case "INTERACTIVE BROKERS"
            fraIntBrokers.Top = fraButtons.Top
            fraPhoton.Visible = False
            fraIntBrokers.Visible = True
            fraLW.Visible = False
            fraPats.Visible = False
            fraAlaron.Visible = False
            fraTransact.Visible = False
            fraPfg.Visible = False
            fraFXCM.Visible = False
            fraTT.Visible = False
            
        Case "PATS", "MAN FINANCIAL - LONDON", "MAN FINANCIAL - CHICAGO"
            fraPats.Top = fraButtons.Top
            fraPhoton.Visible = False
            fraIntBrokers.Visible = False
            fraLW.Visible = False
            fraPats.Visible = True
            PatsStringToControls UCase(lstBrokers.Text)
            fraAlaron.Visible = False
            fraTransact.Visible = False
            fraPfg.Visible = False
            fraFXCM.Visible = False
            fraTT.Visible = False
        
        Case "PHOTON"
            fraPhoton.Top = fraButtons.Top
            fraPhoton.Visible = True
            fraIntBrokers.Visible = False
            fraLW.Visible = False
            fraPats.Visible = False
            fraAlaron.Visible = False
            fraTransact.Visible = False
            fraPfg.Visible = False
            fraFXCM.Visible = False
            fraTT.Visible = False
        
        Case UCase(g.Broker.BrokerName(eTT_AccountType_LindWaldock))
            fraLW.Top = fraButtons.Top
            fraPhoton.Visible = False
            fraIntBrokers.Visible = False
            fraLW.Visible = True
            fraPats.Visible = False
            fraAlaron.Visible = False
            fraTransact.Visible = False
            fraPfg.Visible = False
            fraFXCM.Visible = False
            fraTT.Visible = False
            
        Case "ALARON"
            fraAlaron.Top = fraButtons.Top
            fraPhoton.Visible = False
            fraIntBrokers.Visible = False
            fraLW.Visible = False
            fraPats.Visible = False
            fraAlaron.Visible = True
            fraTransact.Visible = False
            fraPfg.Visible = False
            fraFXCM.Visible = False
            fraTT.Visible = False
        
        Case "TRANSACT"
            fraTransact.Top = fraButtons.Top
            fraPhoton.Visible = False
            fraIntBrokers.Visible = False
            fraLW.Visible = False
            fraPats.Visible = False
            fraAlaron.Visible = False
            fraTransact.Visible = True
            fraPfg.Visible = False
            fraFXCM.Visible = False
            fraTT.Visible = False
            
        Case "PFG"
            fraPfg.Top = fraButtons.Top
            fraPhoton.Visible = False
            fraIntBrokers.Visible = False
            fraLW.Visible = False
            fraPats.Visible = False
            fraAlaron.Visible = False
            fraTransact.Visible = False
            fraPfg.Visible = True
            fraFXCM.Visible = False
            fraTT.Visible = False
    
        Case "FXCM"
            fraFXCM.Top = fraButtons.Top
            fraPhoton.Visible = False
            fraIntBrokers.Visible = False
            fraLW.Visible = False
            fraPats.Visible = False
            fraAlaron.Visible = False
            fraTransact.Visible = False
            fraPfg.Visible = False
            fraFXCM.Visible = True
            fraTT.Visible = False
    
        Case "TRADING TECHNOLOGIES"
            fraFXCM.Top = fraButtons.Top
            fraPhoton.Visible = False
            fraIntBrokers.Visible = False
            fraLW.Visible = False
            fraPats.Visible = False
            fraAlaron.Visible = False
            fraTransact.Visible = False
            fraPfg.Visible = False
            fraFXCM.Visible = False
            fraTT.Visible = True
    
    End Select
    
    strLastBroker = UCase(lstBrokers.Text)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerCfg.lstBrokers.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtAlaronIP_GotFocus
'' Description: When the control gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtAlaronIP_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtAlaronIP

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerCfg.txtAlaronIP_GotFocus"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtAlaronIP_LostFocus
'' Description: When the control loses the focus, revert to default if blank
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtAlaronIP_LostFocus()
On Error GoTo ErrSection:

    If Len(Trim(txtAlaronIP.Text)) = 0 Then
        txtAlaronIP.Text = GetIniFileProperty("Live", "", "IP", AddSlash(App.Path) & "Provided\AlrIps.INI")
        txtAlaronPort.Text = GetIniFileProperty("Live", "", "Port", AddSlash(App.Path) & "Provided\AlrIps.INI")
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerCfg.txtAlaronIP_LostFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtAlaronPassword_GotFocus
'' Description: When the control gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtAlaronPassword_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtAlaronPassword

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerCfg.txtAlaronPassword_GotFocus"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtAlaronUserPort_GotFocus
'' Description: When the control gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtAlaronPort_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtAlaronPort

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerCfg.txtAlaronPort_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtAlaronUserID_GotFocus
'' Description: When the control gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtAlaronUserID_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtAlaronUserID

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerCfg.txtAlaronUserID_GotFocus"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtClientID_GotFocus
'' Description: When the control gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtClientID_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtClientID

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerCfg.txtClientID.GotFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtHostIP_GotFocus
'' Description: When the control gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtHostIP_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtHostIP

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerCfg.txtHostIP.GotFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtCME_GotFocus
'' Description: When the control gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtCME_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtCME

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerCfg.txtCME.GotFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtData_GotFocus
'' Description: When the control gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtData_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtData

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerCfg.txtData.GotFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtECBOT_GotFocus
'' Description: When the control gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtECBOT_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtECBOT

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerCfg.txtECBOT.GotFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtEurex_GotFocus
'' Description: When the control gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtEurex_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtEurex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerCfg.txtEurex.GotFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtFxcmPassword_GotFocus
'' Description: When the control gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtFxcmPassword_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtFxcmPassword

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerCfg.txtFxcmPassword_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtFxcmURL_GotFocus
'' Description: When the control gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtFxcmURL_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtFxcmUrl

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerCfg.txtFxcmURL_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtFxcmUserID_GotFocus
'' Description: When the control gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtFxcmUserID_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtFxcmUserID

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerCfg.txtFxcmUserID_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtLiffe_GotFocus
'' Description: When the control gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtLiffe_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtLiffe

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerCfg.txtLiffe.GotFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtLWFirm_GotFocus
'' Description: When the control gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtLWFirm_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtLWFirm

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerCfg.txtLWFirm.GotFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtLWHost_GotFocus
'' Description: When the control gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtLWHost_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtLWHost

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerCfg.txtLWHost.GotFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtLWPassword_GotFocus
'' Description: When the control gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtLWPassword_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtLWPassword

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerCfg.txtLWPassword.GotFocus", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtLWPort_GotFocus
'' Description: When the control gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtLWPort_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtLWPort

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerCfg.txtLWPort.GotFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtLWTestNum_GotFocus
'' Description: When the control gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtLWTestNum_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtLWTestNum

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerCfg.txtLWTestNum.GotFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtLWUserID_GotFocus
'' Description: When the control gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtLWUserID_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtLWUserID

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerCfg.txtLWUSerID.GotFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtPassword_GotFocus
'' Description: When the control gets the focus, select all of the text
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
    RaiseError "frmBrokerCfg.txtPassword.GotFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtPatsHostIP_GotFocus
'' Description: When the control gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtPatsHostIP_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtPatsHostIP

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerCfg.txtPatsHostIP.GotFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtPatsHostPort_GotFocus
'' Description: When the control gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtPatsHostPort_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtPatsHostPort

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerCfg.txtPatsHostPort.GotFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtPatsPassword_GotFocus
'' Description: When the control gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtPatsPassword_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtPatsPassword

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerCfg.txtPatsPassword.GotFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtPatsUserID_GotFocus
'' Description: When the control gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtPatsUserID_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtPatsUserID

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerCfg.txtPatsUserID.GotFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtPfgIP_GotFocus
'' Description: When the control gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtPfgIP_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtPfgIP

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerCfg.txtPfgIP_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtPfgPassword_GotFocus
'' Description: When the control gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtPfgPassword_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtPfgPassword

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerCfg.txtPfgPassword_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtPfgPort_GotFocus
'' Description: When the control gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtPfgPort_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtPfgPort

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerCfg.txtPfgPort_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtPfgUserID_GotFocus
'' Description: When the control gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtPfgUserID_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtPfgUserID

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerCfg.txtPfgUserID_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtPfgUserID_LostFocus
'' Description: When the control loses the focus, verify the information
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtPfgUserID_LostFocus()
On Error GoTo ErrSection:

    txtPfgUserID.Text = UCase(txtPfgUserID.Text)

    If Left(txtPfgUserID.Text, 1) <> "D" Then
        If (txtPfgIP.Text = "12.36.73.195") Or (txtPfgIP.Text = "12.36.73.39") Then
            txtPfgIP.Text = "12.36.73.184"
        End If
    Else
        If (txtPfgIP.Text = "12.36.73.184") Or (txtPfgIP.Text = "12.36.73.39") Then
            txtPfgIP.Text = "12.36.73.195"
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerCfg.txtPfgUserID_LostFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtPort_GotFocus
'' Description: When the control gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtPort_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtPort

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerCfg.txtPort.GotFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtPhotonChicagoIP_GotFocus
'' Description: When the control gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtPhotonChicagoIP_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtPhotonChicagoIP

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerCfg.txtPhotonChicagoIP.GotFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtPhotonChicagoPort_GotFocus
'' Description: When the control gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtPhotonChicagoPort_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtPhotonChicagoPort

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerCfg.txtPhotonChicagoPort.GotFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtPhotonEuropeIP_GotFocus
'' Description: When the control gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtPhotonEuropeIP_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtPhotonEuropeIP

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerCfg.txtPhotonEuropeIP.GotFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtPhotonEuropePort_GotFocus
'' Description: When the control gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtPhotonEuropePort_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtPhotonEuropePort

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerCfg.txtPhotonEuropePort.GotFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtTransactIP_GotFocus
'' Description: When the control gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtTransactIP_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtTransactIP

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerCfg.txtTransactIP_GotFocus"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtTransactPassword_GotFocus
'' Description: When the control gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtTransactPassword_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtTransactPassword

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerCfg.txtTransactPassword_GotFocus"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtTransactUserID_GotFocus
'' Description: When the control gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtTransactUserID_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtTransactUserID

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerCfg.txtTransactUserID_GotFocus"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtTransactUserID_LostFocus
'' Description: When the control loses the focus, modify the IP
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtTransactUserID_LostFocus()
On Error GoTo ErrSection:

    If UCase(Trim(txtTransactUserID.Text)) <> UCase(g.Transact.SimUserUserName) Then
        If UCase(Trim(txtTransactIP.Text)) = "SIMMT.YORKBA.COM" Then
            txtTransactIP.Text = "mt13.yorkba.com"
        End If
    Else
        If UCase(Trim(txtTransactPassword.Text)) <> UCase(g.Transact.SimUserPassword) Then
            txtTransactPassword.Text = g.Transact.SimUserPassword
        End If
        If UCase(Trim(txtTransactIP.Text)) <> "SIMMT.YORKBA.COM" Then
            txtTransactIP.Text = "simmt.yorkba.com"
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerCfg.txtTransactUserID_LostFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtTtIp_GotFocus
'' Description: When the control gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtTtIp_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtTtIp

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerCfg.txtTtIp_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtTtPassword_GotFocus
'' Description: When the control gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtTtPassword_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtTtPassword

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerCfg.txtTtPassword_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtTtPort_GotFocus
'' Description: When the control gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtTtPort_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtTtPort

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerCfg.txtTtPort_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtUserID_GotFocus
'' Description: When the control gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtUserID_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtUserID

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerCfg.txtUserID.GotFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SizeForm
'' Description: Size the form according the mode
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SizeForm()
On Error GoTo ErrSection:

    Dim lScaleHeightDiff As Long        ' Form Height - Scale Height
    Dim lScaleWidthDiff As Long         ' Form Width - Scale Width
    Dim lHeight As Long                 ' Height to set the form to
    Dim lWidth As Long                  ' Width to set the form to

    lScaleHeightDiff = Height - ScaleHeight
    lScaleWidthDiff = Width - ScaleWidth
    
    lHeight = fraPhoton.Height + lScaleHeightDiff + 120
    
    If (lstBrokers.ListCount = 1) Or (m.nAcctType > -1) Then
        lWidth = fraPhoton.Width + fraButtons.Width + lScaleWidthDiff + 180
    Else
        lWidth = lstBrokers.Width + fraPhoton.Width + fraButtons.Width + lScaleWidthDiff + 240
    End If
    
    Move Left, Top, lWidth, lHeight
    CenterTheForm Me
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerCfg.SizeForm", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PatsStringToControls
'' Description: Set the controls according to the string
'' Inputs:      Broker
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PatsStringToControls(ByVal strBroker As String)
On Error GoTo ErrSection:

    Dim astrInfo As New cGdArray        ' Information to send to the controls
    
    Select Case UCase(strBroker)
        Case "PATS"
            astrInfo.SplitFields m.strPatsInfo, ";"
        Case "MAN FINANCIAL - LONDON"
            astrInfo.SplitFields m.strManLondonInfo, ";"
        Case "MAN FINANCIAL - CHICAGO"
            astrInfo.SplitFields m.strManChicagoInfo, ";"
    End Select
    
    txtPatsUserID.Text = astrInfo(0)
    txtPatsPassword.Text = astrInfo(1)
    txtPatsHostIP.Text = astrInfo(2)
    txtPatsHostPort.Text = astrInfo(3)
    If astrInfo(4) = "Y" Then
        chkSuperTAS.Value = vbChecked
    Else
        chkSuperTAS.Value = vbUnchecked
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerCfg.PatsStringToControls", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ControlsToPatsString
'' Description: Save the controls to the string
'' Inputs:      Broker
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ControlsToPatsString(ByVal strBroker As String)
On Error GoTo ErrSection:

    Dim astrInfo As New cGdArray        ' Information to send to the controls
    
    astrInfo(0) = Trim(txtPatsUserID.Text)
    astrInfo(1) = Trim(txtPatsPassword.Text)
    astrInfo(2) = Trim(txtPatsHostIP.Text)
    astrInfo(3) = Trim(txtPatsHostPort.Text)
    If chkSuperTAS.Value = vbChecked Then
        astrInfo(4) = "Y"
    Else
        astrInfo(4) = "N"
    End If
    
    Select Case UCase(strBroker)
        Case "PATS"
            m.strPatsInfo = astrInfo.JoinFields(";")
        Case "MAN FINANCIAL - LONDON"
            m.strManLondonInfo = astrInfo.JoinFields(";")
        Case "MAN FINANCIAL - CHICAGO"
            m.strManChicagoInfo = astrInfo.JoinFields(";")
    End Select
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerCfg.ControlsToPatsString", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SelectBroker
'' Description: Select a broker from the list box
'' Inputs:      Account Type
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SelectBroker(ByVal nAccountType As eTT_AccountType)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    lstBrokers.ListIndex = 0
    For lIndex = 0 To lstBrokers.ListCount - 1
        If lstBrokers.ItemData(lIndex) = nAccountType Then
            lstBrokers.ListIndex = lIndex
            Exit For
        End If
    Next lIndex

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmBrokerCfg.SelectBroker"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillTransActSymbols
'' Description: Fill in the TransAct symbols text box
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FillTransActSymbols()
On Error GoTo ErrSection:

    Dim astrSymbols As New cGdArray     ' List of symbols to display
    Dim lIndex As Long                  ' Index into a for loop
    Dim strSymbols As String            ' String to display
    
    astrSymbols.Serialize AddSlash(App.Path) & "Provided\LKSyms.TRN", False
    If astrSymbols.Size > 0 Then
        strSymbols = astrSymbols.JoinFields(",")
        strSymbols = Replace(strSymbols, "-", "")
        txtTransactSymbols.Text = strSymbols
    ElseIf FileExist(AddSlash(App.Path) & "Provided\LKSyms.TRN") Then
        txtTransactSymbols.Text = "You are not activated for any markets.  Please call TransAct to get enabled for symbols you wish to trade."
    Else
        txtTransactSymbols.Text = ""
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerCfg.FillTransActSymbols"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DoBroker
'' Description: Do we want to handle this broker?
'' Inputs:      Account Type
'' Returns:     True if handle, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DoBroker(ByVal nBroker As eTT_AccountType) As Boolean
On Error GoTo ErrSection:

    If (m.nAcctType = -1&) Or (m.nAcctType = nBroker) Then
        DoBroker = True
    Else
        DoBroker = False
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmBrokerCfg.DoBroker"
    
End Function
