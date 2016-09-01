VERSION 5.00
Object = "{74115649-BA79-4D7E-AEF6-1B3465B81D22}#1.0#0"; "ActivePhoton.OCX"
Begin VB.Form frmPhoton 
   Caption         =   "Form1"
   ClientHeight    =   1200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4170
   LinkTopic       =   "Form1"
   ScaleHeight     =   1200
   ScaleWidth      =   4170
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraPhoton 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   975
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      Begin ACTIVEPHOTONLib.ActivePhoton PhotonCME 
         Height          =   900
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   900
         _Version        =   65536
         _ExtentX        =   1588
         _ExtentY        =   1588
         _StockProps     =   0
      End
      Begin ACTIVEPHOTONLib.ActivePhoton PhotonLiffe 
         Height          =   900
         Left            =   2880
         TabIndex        =   3
         Top             =   0
         Width           =   900
         _Version        =   65536
         _ExtentX        =   1588
         _ExtentY        =   1588
         _StockProps     =   0
      End
      Begin ACTIVEPHOTONLib.ActivePhoton PhotonEurex 
         Height          =   900
         Left            =   1920
         TabIndex        =   2
         Top             =   0
         Width           =   900
         _Version        =   65536
         _ExtentX        =   1588
         _ExtentY        =   1588
         _StockProps     =   0
      End
      Begin ACTIVEPHOTONLib.ActivePhoton PhotonCBOT 
         Height          =   900
         Left            =   960
         TabIndex        =   1
         Top             =   0
         Width           =   900
         _Version        =   65536
         _ExtentX        =   1588
         _ExtentY        =   1588
         _StockProps     =   0
      End
   End
End
Attribute VB_Name = "frmPhoton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

