VERSION 5.00
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmTradeDiff 
   Caption         =   "Trade differences"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7995
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4050
   ScaleWidth      =   7995
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniButtonImageXP cmdNotepad 
      Height          =   372
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1092
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
      Caption         =   "TradeDiff.frx":0000
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "TradeDiff.frx":0036
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "TradeDiff.frx":0056
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdClose 
      Cancel          =   -1  'True
      Height          =   372
      Left            =   7200
      TabIndex        =   2
      Top             =   840
      Width           =   852
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
      Caption         =   "TradeDiff.frx":0072
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "TradeDiff.frx":009E
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "TradeDiff.frx":00BE
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniTextBoxXP txtTradeStation 
      Height          =   288
      Left            =   3240
      TabIndex        =   1
      Top             =   60
      Width           =   4092
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "TradeDiff.frx":00DA
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
      Tip             =   "TradeDiff.frx":00FA
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "TradeDiff.frx":011A
   End
   Begin HexUniControls.ctlUniTextBoxXP txtSysNav 
      Height          =   288
      Left            =   3240
      TabIndex        =   3
      Top             =   420
      Width           =   4092
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "TradeDiff.frx":0136
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
      Tip             =   "TradeDiff.frx":0156
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "TradeDiff.frx":0176
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdDiff 
      Default         =   -1  'True
      Height          =   372
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   1092
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
      Caption         =   "TradeDiff.frx":0192
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "TradeDiff.frx":01C4
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "TradeDiff.frx":01E4
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniListBoxXP lst 
      Height          =   2760
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   6972
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
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
      TrapTab         =   0   'False
      Tip             =   "TradeDiff.frx":0200
      MultiSelect     =   0
      Sorted          =   0   'False
      HScroll         =   0   'False
      SelBackColor    =   -2147483635
      SelForeColor    =   -2147483634
      RoundedBorders  =   0   'False
      SelectorStyle   =   -1
      MousePointer    =   0
      MouseIcon       =   "TradeDiff.frx":0220
      ManualStart     =   0   'False
      Columns         =   0
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP Label2 
      Height          =   252
      Left            =   1320
      Top             =   120
      Width           =   1812
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
      Caption         =   "TradeDiff.frx":023C
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   0
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "TradeDiff.frx":0290
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "TradeDiff.frx":02B0
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP Label1 
      Height          =   252
      Left            =   1320
      Top             =   480
      Width           =   1812
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
      Caption         =   "TradeDiff.frx":02CC
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   0
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "TradeDiff.frx":031E
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "TradeDiff.frx":033E
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
End
Attribute VB_Name = "frmTradeDiff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
On Error GoTo ErrSection:

    Me.Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeDiff.cmdClose.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub cmdDiff_Click()
On Error GoTo ErrSection:

    Dim strTSfile$, strSNfile$

    strTSfile = Trim(txtTradeStation)
    If Not FileExist(strTSfile) Then
        InfBox "i=e ; TradeStation file does not exist:|" & strTSfile
        Exit Sub
    End If
    
    strSNfile = Trim(txtSysNav)
    If Not FileExist(strTSfile) Then
        InfBox "i=e ; System Navigator file does not exist:|" & strSNfile
        Exit Sub
    End If

    TradeDiff lst, strTSfile, strSNfile
    Enable cmdNotepad

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeDiff.cmdDiff.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub cmdNotepad_Click()
On Error GoTo ErrSection:

    Dim strOutFile$
    
    strOutFile = AddSlash(App.Path) + "TrdDiff.txt"

    If ListToFile(lst, strOutFile) Then
        Shell "Notepad.exe " & strOutFile, vbNormalFocus
        Me.Hide
    Else
        InfBox "i=e ; Could not create " & strOutFile
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeDiff.cmdNotepad.Click", eGDRaiseError_Show
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
    RaiseError "frmTradeDiff.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub Form_Load()
On Error GoTo ErrSection:

    g.Styler.StyleForm Me
    
    cmdClose.Left = -cmdClose.Width - 500
        
    txtTradeStation = GetIniFileProperty("TSfile", "", "Misc", g.strIniFile)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeDiff.Form.Load", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode = 0 Then
        Cancel = True
        Me.Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeDiff.Form.QueryUnload", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub Form_Resize()
On Error Resume Next

    lst.Width = Me.ScaleWidth - lst.Left * 2
    lst.Height = Me.ScaleHeight - lst.Left - lst.Top

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    SetIniFileProperty "TSfile", txtTradeStation, "Misc", g.strIniFile

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeDiff.Form.Unload", eGDRaiseError_Show
    Resume ErrExit

End Sub

