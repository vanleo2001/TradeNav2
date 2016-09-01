VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmFunctionInfo 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Function Information"
   ClientHeight    =   2415
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   5535
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Height          =   315
      Left            =   2040
      TabIndex        =   1
      Top             =   2040
      Width           =   975
   End
   Begin RichTextLib.RichTextBox txtDesc 
      Height          =   1845
      Left            =   75
      TabIndex        =   0
      Top             =   90
      Width           =   5265
      _ExtentX        =   9287
      _ExtentY        =   3254
      _Version        =   393217
      BackColor       =   -2147483633
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmFunctionInfo.frx":0000
   End
End
Attribute VB_Name = "frmFunctionInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type mPrivate
    lTextWidthMax As Long
End Type
Private m As mPrivate

Private Const kLengthOfLongestArg = 5

'tjr 2/03 - should be obsolete function
Public Sub ShowFunctionInfo(strDesc As String, bRTF As Boolean)
On Error GoTo ErrSection:
    Dim lWidth As Long
    Dim lHeight As Long
    
    'Format and display function
    If bRTF Then
        txtDesc.TextRtf = strDesc
    Else
        txtDesc.Text = strDesc
    End If
        
    'Set form Font size/bold to "seed" the TextWidth method...
    With txtDesc
        FontSize = .Font.Size
        FontName = .Font.Name
        FontBold = True
            
        'Add rows to RTF if max width exceeded...
        If Me.TextWidth(.Text) > m.lTextWidthMax Then
            .Text = .Text & Chr(13) & Chr(10)
            lWidth = m.lTextWidthMax
        Else
            'Adjust RTF box width shorter than max...
            lWidth = Me.TextWidth(.Text) + _
                (Screen.TwipsPerPixelX * kLengthOfLongestArg)
        End If
        
        'Calculate height of rtf box...
        lHeight = TextHeight(.Text) + (Screen.TwipsPerPixelY * 2)
        .Move 0, 0, lWidth, lHeight
        
        'Set size of form to fully contain the Rtf box
        Me.Width = .Width + (Me.Width - Me.ScaleWidth)
        Me.Height = .Height + (Me.Height - Me.ScaleHeight)
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "TSOCX.frmFunctionInfo.ShowFunctionInfo", eGDRaiseError_Raise, g.strAppPath
    
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Deactivate()
    Unload Me
End Sub

Private Sub Form_Load()
    m.lTextWidthMax = txtDesc.Width
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Dim btnLft As Integer, btnTop As Integer
    Const C_BMARGIN = 50
    'there is a minimum
    If LimitFormSize(Me, 3100, 1800) Then Exit Sub
    
    btnLft = ScaleWidth / 2 - cmdOK.Width / 2
    btnTop = ScaleHeight - (cmdOK.Height + C_BMARGIN)
    'want the ok button on bottom
    cmdOK.Move btnLft, btnTop
    'want rtf to take up the rest
    txtDesc.Move 0, 0, ScaleWidth, (btnTop - C_BMARGIN)
End Sub
