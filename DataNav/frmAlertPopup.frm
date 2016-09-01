VERSION 5.00
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmAlertPopup 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3150
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   3150
   Begin HexUniControls.ctlUniFrameWL fraOrderAlert 
      Height          =   375
      Left            =   188
      TabIndex        =   5
      Top             =   2040
      Width           =   2775
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
      Caption         =   "frmAlertPopup.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmAlertPopup.frx":002C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmAlertPopup.frx":004C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdEditOrder 
         Height          =   375
         Left            =   960
         TabIndex        =   0
         Top             =   0
         Width           =   855
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
         Caption         =   "frmAlertPopup.frx":0068
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmAlertPopup.frx":0092
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmAlertPopup.frx":00B2
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdSubmitOrder 
         Default         =   -1  'True
         Height          =   375
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   855
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
         Caption         =   "frmAlertPopup.frx":00CE
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmAlertPopup.frx":00FC
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmAlertPopup.frx":011C
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdCancelOrder 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   1920
         TabIndex        =   6
         Top             =   0
         Width           =   855
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
         Caption         =   "frmAlertPopup.frx":0138
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmAlertPopup.frx":0166
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmAlertPopup.frx":0186
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdOK 
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   1080
      Width           =   975
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
      Caption         =   "frmAlertPopup.frx":01A2
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmAlertPopup.frx":01C6
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmAlertPopup.frx":01E6
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniFrameWL fraNextBarButtons 
      Height          =   375
      Left            =   548
      TabIndex        =   2
      Top             =   1560
      Width           =   2055
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
      Caption         =   "frmAlertPopup.frx":0202
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmAlertPopup.frx":022E
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmAlertPopup.frx":024E
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Height          =   375
         Left            =   1080
         TabIndex        =   4
         Top             =   0
         Width           =   975
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
         Caption         =   "frmAlertPopup.frx":026A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmAlertPopup.frx":0298
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmAlertPopup.frx":02B8
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdDisplay 
         Height          =   375
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   975
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
         Caption         =   "frmAlertPopup.frx":02D4
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmAlertPopup.frx":0304
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmAlertPopup.frx":0324
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniLabelXP lblAlert 
      Height          =   195
      Left            =   1245
      Top             =   180
      Width           =   645
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
      Caption         =   "frmAlertPopup.frx":0340
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   0
      BorderStyle     =   0
      AutoSize        =   -1  'True
      Tip             =   "frmAlertPopup.frx":036C
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmAlertPopup.frx":038C
      RightToLeft     =   0   'False
      WordWrap        =   -1  'True
   End
   Begin VB.Image icoAlert 
      Height          =   240
      Left            =   60
      Picture         =   "frmAlertPopup.frx":03A8
      Top             =   1080
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmAlertPopup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmAlertPopup.frm
'' Description: Displays an alert to the user
''
'' Author:      Genesis Financial Data Services
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Public Enum eGD_AlertPopupMode
    eGDAlertMode_QuoteBoardAlert = 0
    eGDAlertMode_NextBarAlert
    eGDAlertMode_OrderStatusAlert
    eGDAlertMode_PlaceOrderAlert
    eGDAlertMode_MessageBox
End Enum

Private Type mPrivate
    nMode As eGD_AlertPopupMode
    hWnd As Long                        ' hWnd of chart calling the alert
    
    strAlert As String                  ' Alert string passed in
    strSymbol As String                 ' Symbol passed in
    
    Alert As cAlert                     ' Alert object if order was triggered by an alert
End Type
Private m As mPrivate

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Setup and show the form
'' Inputs:      Mode, Symbol, Alert String, Hwnd of Chart
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowMe(ByVal Mode As eGD_AlertPopupMode, ByVal strSymbol As String, ByVal strAlert As String, _
    Optional ByVal hWnd As Long = 0&, Optional ByVal nOrderConfirm As Long = 1&, _
    Optional Alert As cAlert = Nothing)
On Error GoTo ErrSection:

    Dim lLabelWidth As Long             ' What the label width should be

    m.nMode = Mode
    m.strSymbol = strSymbol
    m.strAlert = strAlert
    Set m.Alert = Alert
    
    lblAlert.AutoSize = False
    lblAlert.WordWrap = False
    lblAlert.Font.Bold = True
    
    Select Case Mode
        Case eGDAlertMode_QuoteBoardAlert
            Caption = "Alert for " & strSymbol
            cmdOK.Visible = True
            fraNextBarButtons.Visible = False
            fraOrderAlert.Visible = False
            lblAlert.Alignment = vbCenter
            lblAlert.Caption = strAlert
        
        Case eGDAlertMode_NextBarAlert
            Caption = "Orders Alert for " & strSymbol
            cmdOK.Visible = False
            fraNextBarButtons.Visible = True
            fraNextBarButtons.Top = cmdOK.Top
            fraOrderAlert.Visible = False
            lblAlert.Alignment = vbCenter
            lblAlert.Caption = strAlert
        
        Case eGDAlertMode_OrderStatusAlert
            Caption = strSymbol
            cmdOK.Visible = True
            fraNextBarButtons.Visible = False
            fraOrderAlert.Visible = False
            lblAlert.Alignment = vbLeftJustify
            lblAlert.Caption = strAlert
            
            lLabelWidth = LabelWidth(strAlert)
            If lLabelWidth > lblAlert.Width Then
                Width = (lLabelWidth + 120) + (Width - ScaleWidth)
                lblAlert.Width = lLabelWidth
                cmdOK.Left = (Me.ScaleWidth - cmdOK.Width) / 2
            End If
        
        Case eGDAlertMode_PlaceOrderAlert
            Caption = "Alert for " & strSymbol
            cmdOK.Visible = False
            fraNextBarButtons.Visible = False
            fraOrderAlert.Visible = True
            lblAlert.Alignment = vbLeftJustify
            fraOrderAlert.Top = cmdOK.Top
            lblAlert.Caption = OrderToCaption
            
            strAlert = Parse(lblAlert.Caption, vbCrLf, 1)
            If Len(Parse(lblAlert.Caption, vbCrLf, 2)) > Len(strAlert) Then
                strAlert = Parse(lblAlert.Caption, vbCrLf, 2)
            End If
            If Len(Parse(lblAlert.Caption, vbCrLf, 3)) > Len(strAlert) Then
                strAlert = Parse(lblAlert.Caption, vbCrLf, 3)
            End If
            
            lLabelWidth = LabelWidth(strAlert)
            If lLabelWidth > lblAlert.Width Then
                If lLabelWidth > fraOrderAlert.Width Then
                    Width = (lLabelWidth + 120) + (Width - ScaleWidth)
                Else
                    Width = (fraOrderAlert.Width + 120) + (Width - ScaleWidth)
                End If
                lblAlert.Width = lLabelWidth
                cmdOK.Left = (Me.ScaleWidth - cmdOK.Width) / 2
            End If
            If Not m.Alert Is Nothing Then
                If m.Alert.AlertType = eGDAlertType_Annot Then
                    If Not m.Alert.Annotation Is Nothing Then
                        If Not m.Alert.Annotation.AnnotChart Is Nothing Then
                            If m.Alert.Annotation.AnnotChart.IsInWhatIfMode Then
                                cmdSubmitOrder.Enabled = False      '4293
                                cmdEditOrder.Enabled = False
                                nOrderConfirm = 1
                            End If
                        End If
                    End If
                ElseIf m.Alert.AlertType = eGDAlertType_Chart Then
                    If Not m.Alert.Indicator Is Nothing Then
                        If Not m.Alert.Indicator.IndChart Is Nothing Then
                            If m.Alert.Indicator.IndChart.IsInWhatIfMode Then
                                cmdSubmitOrder.Enabled = False
                                nOrderConfirm = 1
                            End If
                        End If
                    End If
                End If
            End If
    End Select
    
    If (Mode = eGDAlertMode_PlaceOrderAlert) And (nOrderConfirm = 0&) Then
        SubmitOrder
    Else
        Height = 1905
    
        m.hWnd = hWnd
        Me.Refresh
    
        SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_SHOWWINDOW
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertPopup.ShowMe"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMessageBox
'' Description: Setup and show the form as a message box
'' Inputs:      Message, Caption, Text Alignment
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowMessageBox(ByVal strMessage As String, ByVal strCaption As String, Optional ByVal nTextAlignment As AlignmentConstants = vbLeftJustify, _
            Optional ByVal bBoldFont As Boolean = False, Optional ByVal nBackColor As Long = vbButtonFace)
On Error GoTo ErrSection:
    
    Dim dHeightDiff As Double
    
    m.nMode = eGDAlertMode_MessageBox
    
    Caption = strCaption
    Icon = Picture16("kBlank")
    BackColor = nBackColor
    
    cmdOK.Visible = True
    fraNextBarButtons.Visible = False
    fraOrderAlert.Visible = False
    
    lblAlert.AutoSize = True
    lblAlert.WordWrap = True
    lblAlert.Width = 4560 - 360
    lblAlert.Font.Bold = bBoldFont
    lblAlert.Alignment = nTextAlignment
    lblAlert.Caption = Replace(strMessage, "|", Chr(13))
    lblAlert.Refresh
    
    dHeightDiff = Height - ScaleHeight
    Move Left, Top, 4560, lblAlert.Height + cmdOK.Height + 540 + dHeightDiff
    
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_SHOWWINDOW

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertPopup.ShowMessageBox"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: When the user clicks Cancel, close the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancel_Click()
On Error GoTo ErrSection:

    Unload Me
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertPopup.cmdCancel_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancelOrder_Click
'' Description: Close the alert form without submitting the order
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancelOrder_Click()
On Error GoTo ErrSection:

    Unload Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertPopup.cmdCancelOrder_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdDisplay_Click
'' Description: When the user clicks Display, show the next bar report for
''              the correct chart
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdDisplay_Click()
On Error GoTo ErrSection:

    Dim frm As Form                     ' Temporary form object
    Dim bFound As Boolean               ' Was the correct form found?

    ' Look for the chart with the correct hWnd
    For Each frm In Forms
        If frm.hWnd = m.hWnd Then
            bFound = True
            Me.Hide
            frm.Chart.ShowSystemReport True
            Exit For
        End If
    Next frm
    
    ' If the chart was not found, notify the user
    If Not bFound Then
        Me.Hide
        Err.Raise vbObjectError + 1000, , "The Chart for this Next Bar Report is no longer open"
    End If
    
ErrExit:
    Unload Me
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertPopup.cmdDisplay_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdEditOrder_Click
'' Description: Allow the user to edit the order
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdEditOrder_Click()
On Error GoTo ErrSection:

    Unload Me
    EditOrder OrderTextToOrder, True
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertPopup.cmdEditOrder_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOK_Click
'' Description: When the user clicks OK, close the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
On Error GoTo ErrSection:

    Unload Me
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertPopup.cmdOK_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdSubmitOrder_Click
'' Description: Allow the user to submit the order that they entered
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdSubmitOrder_Click()
On Error GoTo ErrSection:

    SubmitOrder

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertPopup.cmdSubmitOrder_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_KeyDown
'' Description: Allow the user to view help on the form by pressing F1
'' Inputs:      Code of Key Pressed, Shift/Ctrl/Alt status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If KeyCode = vbKeyF1 Then
        KeyCode = 0
        g.Help.ShowF1Help Me
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertPopup.Form_KeyDown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: When the form is loaded, center it
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Dim i&, iMaxTop&, iLeft&

    g.Styler.StyleForm Me
    
    'CenterTheForm Me
    
    Me.Icon = Picture16(ToolbarIcon("ID_Alerts"), , True)
    Me.fraNextBarButtons.BackColor = Me.BackColor
    Me.fraOrderAlert.BackColor = Me.BackColor
    
    ' default: place at 1/4 from left, 1/4 from top
    'iLeft = Screen.Width / 4
    'iMaxTop = Screen.Height / 4
    iLeft = 240
    iMaxTop = 240
    
    ' get lowest cascaded alerts form that currently exists
    For i = 0 To Forms.Count - 1
        If TypeOf Forms(i) Is frmAlertPopup Then
            If Forms(i).hWnd <> Me.hWnd Then
                If Forms(i).Top > iMaxTop Then iMaxTop = Forms(i).Top
            End If
        End If
    Next
    
    ' keep alert to left of center so no chance of covering
    ' up a centered modal dialog box (which then makes it
    ' seem to the user like the program is locked up!),
    ' but cascade them down the left side (so can see multiple)
    iMaxTop = iMaxTop + Me.Height - Me.ScaleHeight
    If iMaxTop >= Screen.Height - Me.Height Then
        iMaxTop = 0 '(but if below screen, just put at top)
    End If
    Me.Move iLeft, iMaxTop
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertPopup.Form_Load"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Resize
'' Description: Resize and move controls as the form is resized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Resize()
On Error Resume Next

    With lblAlert
        If m.nMode = eGDAlertMode_PlaceOrderAlert Then
            .Move 60, 60, ScaleWidth - 120, ScaleHeight - fraOrderAlert.Height - (60 * 3)
        Else
            .Move 60, 60, ScaleWidth - 120
        End If
    End With
    
    With cmdOK
        .Move ((ScaleWidth / 2) - (.Width / 2)), ScaleHeight - .Height - 60
    End With

    With fraNextBarButtons
        .Move ((ScaleWidth / 2) - (.Width / 2)), ScaleHeight - .Height - 60
    End With

    With fraOrderAlert
        .Move ((ScaleWidth / 2) - (.Width / 2)), ScaleHeight - .Height - 60
    End With

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: Cleanup after ourseleves when the form is unloading
'' Inputs:      Whether to Cancel the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    PlaySoundFile ' to cancel sound

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertPopup.Form_Unload"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LabelWidth
'' Description: Determine the width of the label by the given text
'' Inputs:      Text to calculate for
'' Returns:     What the label width should be
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function LabelWidth(ByVal strText As String) As Single
On Error GoTo ErrSection:

    Dim FormFont As StdFont             ' Current font of the form
    
    Set FormFont = Me.Font
    Set Me.Font = lblAlert.Font
    LabelWidth = Me.TextWidth(strText)
    Set Me.Font = FormFont

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmAlertPopup.LabelWidth"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderToCaption
'' Description: Make an english string out of the given order string
'' Inputs:      Order Text
'' Returns:     English String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function OrderToCaption() As String
On Error GoTo ErrSection:

    Dim astrOrder As New cGdArray       ' Order broken out into an array
    Dim strReturn As String             ' String to return from the function
    Dim strSymbol As String             ' Symbol to work with
    
    astrOrder.Create eGDARRAY_Strings
    astrOrder.SplitFields m.strAlert, ","
    
    strSymbol = RollSymbolForDate(m.strSymbol, Date)
    
    strReturn = astrOrder(0) & " " & Format(astrOrder(1), "#,##0") & " "
    strReturn = strReturn & strSymbol & vbCrLf
    Select Case ValOfText(astrOrder(2))
        Case eTT_OrderType_Market
            strReturn = strReturn & "At MARKET" & vbCrLf
        Case eTT_OrderType_Stop
            strReturn = strReturn & "At " & PriceDisplay(ValOfText(astrOrder(3)), strSymbol) & " STOP" & vbCrLf
        Case eTT_OrderType_Limit
            strReturn = strReturn & "At " & PriceDisplay(ValOfText(astrOrder(4)), strSymbol) & " LIMIT" & vbCrLf
        Case eTT_OrderType_StopWithLimit
            strReturn = strReturn & "At " & PriceDisplay(ValOfText(astrOrder(3)), strSymbol) & " STOP with " & PriceDisplay(ValOfText(astrOrder(4)), strSymbol) & " LIMIT" & vbCrLf
    End Select
    strReturn = strReturn & "In Account: " & g.Broker.AccountNameForID(CLng(Val(astrOrder(5))))
    If ValOfText(astrOrder(6)) = 0 Then
        strReturn = strReturn & " GTC"
    ElseIf ValOfText(astrOrder(6)) > 0 Then
        strReturn = strReturn & " GTD: " & Format(ValOfText(astrOrder(6)), DateFormat("Format", MM_DD_YYYY))
    End If
    
    OrderToCaption = strReturn

ErrExit:
    Set astrOrder = Nothing
    Exit Function
    
ErrSection:
    Set astrOrder = Nothing
    RaiseError "frmAlertPopup.OrderToCaption"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderTextToOrder
'' Description: Convert the order text to an order object
'' Inputs:      None
'' Returns:     Order object
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function OrderTextToOrder() As cPtOrder
On Error GoTo ErrSection:

    Dim Order As New cPtOrder           ' Order object to fill with information
    Dim astrOrder As New cGdArray       ' Order broken out into an array
    
    astrOrder.Create eGDARRAY_Strings
    astrOrder.SplitFields m.strAlert, ","
    
    With Order
        .Buy = (UCase(astrOrder(0)) = "BUY")
        .Quantity = ValOfText(astrOrder(1))
        .OrderType = ValOfText(astrOrder(2))
        .StopPrice = ValOfText(astrOrder(3))
        .LimitPrice = ValOfText(astrOrder(4))
        .AccountID = ValOfText(astrOrder(5))
        .Expiration = ValOfText(astrOrder(6))
        .SymbolOrSymbolID = RollSymbolForDate(m.strSymbol, Date)
    End With
    
    Set OrderTextToOrder = Order

ErrExit:
    Set Order = Nothing
    Set astrOrder = Nothing
    Exit Function
    
ErrSection:
    Set Order = Nothing
    Set astrOrder = Nothing
    RaiseError "frmAlertPopup.OrderTextToOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SubmitOrder
'' Description: Allow the user to submit the order that they entered
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SubmitOrder()
On Error GoTo ErrSection:

    Dim Order As New cPtOrder           ' Order to submit
    
    Set Order = OrderTextToOrder
    Order.GenesisOrderID = NextGenesisOrderID(g.Broker.AccountNumberForID(Order.AccountID))
    Order.Save
    If Not m.Alert Is Nothing Then m.Alert.OrderSubmitted = Order.OrderID
    
    Unload Me
    
    g.Broker.BrokerDebug g.Broker.AccountTypeForID(Order.AccountID), "Creating Order from Alert: " & Order.OrderText, True
    mTradeTracker.SubmitOrder Order

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertPopup.SubmitOrder"
    
End Sub

