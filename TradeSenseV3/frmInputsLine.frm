VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmInputsLine 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7875
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   465
   ScaleWidth      =   7875
   ShowInTaskbar   =   0   'False
   Begin RichTextLib.RichTextBox ParmText 
      Height          =   240
      Left            =   45
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   75
      Width           =   7665
      _ExtentX        =   13520
      _ExtentY        =   423
      _Version        =   393217
      BackColor       =   12648447
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      TextRTF         =   $"frmInputsLine.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmInputsLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type mPrivate
    lTextWidthMax As Long
End Type
Private m As mPrivate

Public Sub SizeParmLine(ByVal strText As String, ByVal lLengthOfLongestArg As Long)
On Error GoTo ErrSection:
    
    Dim lWidth As Long
    Dim lHeight As Long
    Const C_MARGIN = 50
    With ParmText
        'Simulate text entered into RTF box...
        .Text = " " & strText & " "
        
        'Set form Font size/bold to "seed" the TextWidth method...
        FontSize = .Font.Size
        FontName = .Font.Name
        FontBold = True
        
        'Add rows to RTF if max width exceeded...
        If Me.TextWidth(.Text) > m.lTextWidthMax Then
            .Text = .Text & Chr(13) & Chr(10)
            lWidth = m.lTextWidthMax
        Else
            'Adjust RTF box width shorter than max...
            lWidth = Me.TextWidth(.Text) + (Screen.TwipsPerPixelX * lLengthOfLongestArg)
        End If
        
        'Calculate height of rtf box...
        lHeight = TextHeight(.Text) + (Screen.TwipsPerPixelY * 2) + C_MARGIN
        .Move 0, 0, lWidth, lHeight
        
        'Set size of form to fully contain the Rtf box
        Me.Width = .Width + (Me.Width - Me.ScaleWidth)
        Me.Height = .Height + (Me.Height - Me.ScaleHeight)
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "TSOCX.frmInputsLine.SizeParmLine", eGDRaiseError_Raise, g.strAppPath
        
End Sub

Private Sub Form_Load()
On Error GoTo ErrSection:

    m.lTextWidthMax = ParmText.Width

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "TSOCX.frmInputsLine.Form.Load", eGDRaiseError_Show, g.strAppPath

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If Me.Visible Then
        Me.Hide
        Cancel = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "TSOCX.frmInputsLine.Form.QueryUnload", eGDRaiseError_Show, g.strAppPath

End Sub
