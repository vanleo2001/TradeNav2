VERSION 5.00
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmChartPixPerBar 
   BorderStyle     =   0  'None
   ClientHeight    =   1275
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   2550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   1200
      Left            =   38
      ScaleHeight     =   1140
      ScaleWidth      =   2415
      TabIndex        =   0
      Top             =   37
      Width           =   2475
      Begin HexUniControls.ctlUniButtonImageXP cmdRestore 
         Height          =   375
         Left            =   165
         TabIndex        =   2
         Top             =   680
         Width           =   2055
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
         Caption         =   "frmChartPixPerBar.frx":0000
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmChartPixPerBar.frx":0048
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmChartPixPerBar.frx":0068
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdSet 
         Height          =   375
         Left            =   165
         TabIndex        =   1
         Top             =   240
         Width           =   2055
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
         Caption         =   "frmChartPixPerBar.frx":0084
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmChartPixPerBar.frx":00C6
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmChartPixPerBar.frx":00E6
         RightToLeft     =   0   'False
      End
   End
End
Attribute VB_Name = "frmChartPixPerBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type mPrivate
    frmCaller As Form
End Type

Private m As mPrivate


Public Sub ShowMe(frm As Form)
On Error GoTo ErrSection:
    
    Dim pt As POINTAPI
    Dim X As Long, Y As Long
    
    If Not frm Is Nothing Then
        Set m.frmCaller = frm
        If GetCursorPos(pt) <> 0 Then
            X = (pt.X * Screen.TwipsPerPixelX) - Me.Width - 100
            Y = (pt.Y * Screen.TwipsPerPixelY) - Me.Height - 100
            
            Me.Move X, Y
            
            Picture1.BackColor = g.nColorTheme
            Picture1.ForeColor = vbRed
            Picture1.Font.Bold = True
            Picture1.Font.Size = 8
            Picture1.CurrentX = Picture1.Left + 2200
            Picture1.CurrentY = 0
            Picture1.Print "X"
            
            ShowForm Me, eForm_Modal
        Else
            'theoretically should never get here, but this message will help track the problem
            StatusMsg "GetCursorPos failed."
            Unload Me
        End If
    Else
        Unload Me
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmChartPixPerBar.ShowMe"
    
End Sub

Private Sub cmdRestore_Click()
On Error GoTo ErrSection:

    Dim i&, j&

    If Not m.frmCaller Is Nothing Then
        If Not m.frmCaller.Chart Is Nothing Then
            i = m.frmCaller.Chart.PixelsPerBar
            j = m.frmCaller.Chart.DefaultPixelsPerBar(Me)
            If i <> j Then m.frmCaller.Chart.GenerateChart eRedo1_Scrolled
        End If
    End If
    
    Set m.frmCaller = Nothing
    
    Unload Me

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmChartPixPerBar.cmdRestore_Click"
    
End Sub

Private Sub cmdSet_Click()
On Error GoTo ErrSection:

    Dim i&
    
    If Not m.frmCaller Is Nothing Then
        If Not m.frmCaller.Chart Is Nothing Then
            i = m.frmCaller.Chart.PixelsPerBar
            m.frmCaller.Chart.DefaultPixelsPerBar(Me) = i
        End If
    End If
    
    Set m.frmCaller = Nothing
    
    Unload Me

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmChartPixPerBar.cmdSet_Click"
    
End Sub

Private Sub Form_Load()
    g.Styler.StyleForm Me
End Sub

Private Sub Picture1_Click()
    Unload Me
End Sub

