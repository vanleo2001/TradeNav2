VERSION 5.00
Begin VB.Form frmRegFiles 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Verifying Installed Files"
   ClientHeight    =   795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   795
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
Begin HexUniControls.ctlUniLabelXP Label2
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(this process may take up to a minute)"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   5295
   End
Begin HexUniControls.ctlUniLabelXP Label1
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Please wait while verifying the installed program files ..."
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
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5175
   End
End
Attribute VB_Name = "frmRegFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


