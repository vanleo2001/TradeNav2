VERSION 5.00
Object = "{74416953-E3C7-11D3-BB97-00600842D31C}#1.0#0"; "DiagramCtl.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmDgm 
   Caption         =   "frmDgm"
   ClientHeight    =   6135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7680
   Icon            =   "frmDgm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   7680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   1035
      Left            =   5580
      ScaleHeight     =   975
      ScaleWidth      =   1635
      TabIndex        =   4
      Top             =   4320
      Visible         =   0   'False
      Width           =   1695
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   5400
      Width           =   6015
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
      Caption         =   "frmDgm.frx":030A
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmDgm.frx":032A
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmDgm.frx":034A
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdClose 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   960
         TabIndex        =   3
         Top             =   60
         Width           =   795
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
         Caption         =   "frmDgm.frx":0366
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmDgm.frx":0392
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmDgm.frx":03B2
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdPrint 
         Height          =   375
         Left            =   0
         TabIndex        =   2
         Top             =   60
         Width           =   795
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
         Caption         =   "frmDgm.frx":03CE
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmDgm.frx":03FA
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmDgm.frx":041A
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label1 
         Height          =   435
         Left            =   1920
         Top             =   60
         Width           =   3795
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
         Caption         =   "frmDgm.frx":0436
         BackColor       =   -2147483633
         ForeColor       =   8388608
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmDgm.frx":0530
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmDgm.frx":0550
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin DIAGRAMCTLLib.DiagramCtl dgm 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7455
      _Version        =   65536
      _ExtentX        =   13150
      _ExtentY        =   9128
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmDgm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmDgm.frm
'' Description: Shows the diagram of the rule
''
'' Author:      Genesis Financial Data Services
''              425 E Woodmen Rd
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    strRuleName As String               ' Name of the rule to diagram
    strRuleText As String               ' Text of the rule to diagram
End Type
Private m As mPrivate

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdClose_Click
'' Description: Close the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdClose_Click()
On Error GoTo ErrSection:

    Me.Hide
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDgm.Close", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdPrint_Click
'' Description: Call the print preview form to be able to print the diagram
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdPrint_Click()
On Error GoTo ErrSection:

    frmPrintPreview.ShowMe "SNV Diagram", frmDgm, 0, , , , , , , ePrintToFile_None

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmDgm.cmdPrint.Click", eGDRaiseError_Show
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
    RaiseError "frmDgm.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Initialize the controls and size/place the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Dim strText As String               ' Text of placement out of the ini file

    Me.Icon = Picture16(ToolbarIcon("kDiagram"))
    
    g.Styler.StyleForm Me
    
    strText = GetIniFileProperty("Dgm", "", "Placement", g.strIniFile)
    If strText = "" Then
        CenterTheForm Me
    Else
        SetFormPlacement Me, strText, "LHTW"
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDgm.Form.Load", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: If the user clicks on the 'X', close the form
'' Inputs:      Whether to Cancel the unload, Mode of the unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode = 0 Then
        Cancel = True
        Me.Hide
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDgm.Form.QueryUnload", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Resize
'' Description: Resize and move the controls as the form is resized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Resize()
On Error Resume Next

    If LimitFormSize(Me, 2000, 1500) Then Exit Sub

    fraButtons.Top = Me.ScaleHeight - fraButtons.Height
    With dgm
        .Move .Left, .Top, Me.ScaleWidth - .Left * 2, fraButtons.Top - .Top
    End With

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: Save the form placement on unload
'' Inputs:      Whether to Cancel the unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    SetIniFileProperty "Dgm", GetFormPlacement(Me), "Placement", g.strIniFile
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDgm.Form.Unload", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenerateReport
'' Description: Callback function for the Print Preview form
'' Inputs:      Arguments
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GenerateReport(ByVal vArgs As Variant)
On Error GoTo ErrSection:

    Dim lAvailWidth As Long             ' Available width on the page
    Dim lAvailHeight As Long            ' Available height on the page
    Dim lPicHeight As Long              ' Picture height
    Dim lPicWidth As Long               ' Picture width
    Dim lResize As Long                 ' Resize amount for the picture
    Dim lSaveFontSize As Long           ' Font Size before printing
    
    With frmPrintPreview.vp
        .StartDoc
        
        ' change font size for printing
        ' (avoid flicker by locking window update)
        LockWindowUpdate Me.hWnd
        lSaveFontSize = dgm.Font.Size
        'dgm.Font.Name = "Times New Roman"
        dgm.Font.Size = 8
        dgm.RedrawDC = .hDC '(trigger to save bitmap)
        dgm.Font.Size = lSaveFontSize
        LockWindowUpdate 0
        Picture1.Picture = LoadPicture(AddSlash(App.Path) & "DiagramCtl.BMP")
        
        ' Set the header and the footer
        DoPrintHeader
        
        .Font.Name = "Times New Roman"
        '.TextAlign = taCenterTop
        .Font.Bold = True
        .Font.Size = 14
        .FontUnderline = True
        .Text = "Rule:"
        .FontUnderline = False
        .Text = "    " & m.strRuleName & vbLf
        .Font.Bold = False
        .Font.Size = 10
        .TextAlign = taLeftTop
        ''.Text = m.strRuleText & vbLf
        
        ' Redraw the diagram on the given device context
        lAvailWidth = .PageWidth - .MarginLeft - .MarginRight
        lAvailHeight = .PageHeight - .CurrentY - .MarginBottom
        .CalcPicture = Picture1.Picture
        lPicHeight = .Y2 - .Y1
        lPicWidth = .X2 - .X1
        
        .X1 = 0
        .Y1 = 0
        
        If lPicHeight > lAvailHeight And lPicWidth > lAvailWidth Then
            .DrawPicture Picture1.Picture, _
                 .MarginLeft, .CurrentY, lAvailWidth, lAvailHeight, vppaZoom
        ElseIf lPicHeight > lAvailHeight Then
            lResize = (lPicWidth * (1 - (lAvailHeight / lPicHeight))) / 2
            .DrawPicture Picture1.Picture, _
                 .MarginLeft - lResize, .CurrentY, lPicWidth, lAvailHeight, vppaZoom
        ElseIf lPicWidth > lAvailWidth Then
            lResize = (lPicHeight * (1 - (lAvailWidth / lPicWidth))) / 2
            .DrawPicture Picture1.Picture, _
                 .MarginLeft, .CurrentY - lResize, lAvailWidth, lPicHeight, vppaZoom
        Else
            .DrawPicture Picture1.Picture, _
                 .MarginLeft, .CurrentY, lPicWidth, lPicHeight, vppaZoom
        End If
        
        .EndDoc
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDgm.GenerateReport", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Initialize and show the form
'' Inputs:      Text to diagram, Name of the rule to diagram
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowMe(ByVal strText As String, ByVal strRuleName As String)
On Error GoTo ErrSection:
    
    m.strRuleName = strRuleName
    m.strRuleText = strText
    Caption = "Diagram:  " & strRuleName
    dgm.CodedText = strText
    ShowForm Me, True
    
ErrExit:
    Unload Me
    Exit Sub
    
ErrSection:
    RaiseError "frmDgm.ShowMe", eGDRaiseError_Raise

End Sub

