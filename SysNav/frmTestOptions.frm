VERSION 5.00
Begin VB.Form frmTestOptions 
   Caption         =   "Test Strategy"
   ClientHeight    =   2715
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3615
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2715
   ScaleWidth      =   3615
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Assume hit High before Low of bar ..."
      Height          =   855
      Left            =   360
      TabIndex        =   4
      Top             =   1080
      Width           =   2895
      Begin VB.OptionButton optOmegaMethod 
         Caption         =   "if Open > Midpoint of bar"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   480
         Width           =   2175
      End
      Begin VB.OptionButton optGenesisMethod 
         Caption         =   "if Open > Close of bar"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   2160
      Width           =   975
   End
   Begin VB.CheckBox chkNextBarReport 
      Caption         =   "&Generate Orders for the Next Bar"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   2775
   End
   Begin VB.CheckBox chkLastBarComplete 
      Caption         =   "&Last Bar of data is Complete"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Value           =   1  'Checked
      Width           =   2535
   End
End
Attribute VB_Name = "frmTestOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
On Error GoTo ErrSection:

    cmdCancel.Tag = "CANCEL"
    Me.Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTestOptions.cmdCancel.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub cmdOK_Click()
On Error GoTo ErrSection:
    
    cmdCancel.Tag = ""
    Me.Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTestOptions.cmdOK.Cancel", eGDRaiseError_Show
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
    RaiseError "frmTestOptions.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub Form_Load()
On Error GoTo ErrSection:

    CenterTheForm Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTestOptions.Form.Load", eGDRaiseError_Show
    Resume ErrExit

End Sub
