VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmHumeOrder 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Place Order"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7335
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraButtons 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   375
      Left            =   2880
      TabIndex        =   27
      Top             =   7080
      Width           =   3255
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   1680
         TabIndex        =   29
         Top             =   0
         Width           =   1575
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&Submit"
         Height          =   375
         Left            =   0
         TabIndex        =   28
         Top             =   0
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDE6D7&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1815
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   7095
      Begin VB.Label lblSellQuantity2 
         Alignment       =   2  'Center
         BackColor       =   &H00FDE6D7&
         Caption         =   "Q2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6480
         TabIndex        =   26
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label lblSellQuantity1 
         Alignment       =   2  'Center
         BackColor       =   &H00FDE6D7&
         Caption         =   "Q1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6480
         TabIndex        =   25
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label lblSellPrice2 
         Alignment       =   2  'Center
         BackColor       =   &H00FDE6D7&
         Caption         =   "Price2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5400
         TabIndex        =   24
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lblSellPrice1 
         Alignment       =   2  'Center
         BackColor       =   &H00FDE6D7&
         Caption         =   "Price1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5400
         TabIndex        =   23
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label lblBuyQuantity2 
         Alignment       =   2  'Center
         BackColor       =   &H00FDE6D7&
         Caption         =   "Q2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   22
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label lblBuyQuantity1 
         Alignment       =   2  'Center
         BackColor       =   &H00FDE6D7&
         Caption         =   "Q1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   21
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label lblBuyPrice2 
         Alignment       =   2  'Center
         BackColor       =   &H00FDE6D7&
         Caption         =   "Price2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   20
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label lblBuyPrice1 
         Alignment       =   2  'Center
         BackColor       =   &H00FDE6D7&
         Caption         =   "Price1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   19
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lblBuySymbol2 
         BackColor       =   &H00FDE6D7&
         Caption         =   "Symbol2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00FDE6D7&
         Caption         =   "Price"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5520
         TabIndex        =   17
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblBuySymbol1 
         BackColor       =   &H00FDE6D7&
         Caption         =   "Symbol1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00FDE6D7&
         Caption         =   "Quantity"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6360
         TabIndex        =   15
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblSellSymbol2 
         BackColor       =   &H00FDE6D7&
         Caption         =   "Symbol2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3720
         TabIndex        =   14
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label lblSellSymbol1 
         BackColor       =   &H00FDE6D7&
         Caption         =   "Symbol1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3720
         TabIndex        =   13
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FDE6D7&
         Caption         =   "Symbol"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3720
         TabIndex        =   12
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FDE6D7&
         Caption         =   "Quantity"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   11
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00FDE6D7&
         Caption         =   "Price"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   10
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FDE6D7&
         Caption         =   "Symbol"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FDE6D7&
         Caption         =   "SELL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4080
         TabIndex        =   8
         Top             =   120
         Width           =   2655
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FDE6D7&
         Caption         =   "BUY"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   7
         Top             =   120
         Width           =   2775
      End
      Begin VB.Line Line12 
         X1              =   3600
         X2              =   3600
         Y1              =   0
         Y2              =   1800
      End
      Begin VB.Line Line11 
         X1              =   7080
         X2              =   7080
         Y1              =   0
         Y2              =   1800
      End
      Begin VB.Line Line10 
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   1800
      End
      Begin VB.Line Line9 
         X1              =   0
         X2              =   7080
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line Line8 
         X1              =   0
         X2              =   7080
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line7 
         X1              =   0
         X2              =   7080
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close Window"
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   7080
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      FillColor       =   &H80000009&
      Height          =   1575
      Left            =   120
      Picture         =   "frmHumeOrder.frx":0000
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   469
      TabIndex        =   3
      Top             =   120
      Width           =   7095
   End
   Begin RichTextLib.RichTextBox rtbSample 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   4080
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   3201
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"frmHumeOrder.frx":DAE7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblWarning 
      Alignment       =   2  'Center
      Caption         =   $"frmHumeOrder.frx":DB6F
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   0
      TabIndex        =   5
      Top             =   6360
      Width           =   7335
   End
   Begin VB.Label lblSpecificOrders 
      Alignment       =   2  'Center
      Caption         =   "For Specific Orders Call - Fox Investments at (888) 281-9566"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   6000
      Width           =   6855
   End
   Begin VB.Label Label16 
      Caption         =   "Sample Broker Instructions:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   3720
      Width           =   3015
   End
End
Attribute VB_Name = "frmHumeOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmHumeOrder.frm
'' Description: Shows a sample order placement form for the order which the
''              user just entered
''
'' Author:      Genesis Financial Data Services
''              425 Woodmen Rd
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date      Author      Description
'' 03/09/01  D Jarmuth   Created
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    bOK As Boolean                      ' Did the user click on OK?
End Type
Private m As mPrivate

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel
'' Description: If the user clicks on the Cancel button, unload the form
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
    RaiseError "frmHumeOrder.cmdCancel.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdClose_Click
'' Description: If the user clicks on the Close Window button, unload the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdClose_Click()
On Error GoTo ErrSection:

    m.bOK = False
    Me.Hide
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmHumeOrder.cmdClose.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOK_Click
'' Description: Unload the form if the user clicks on OK
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
On Error GoTo ErrSection:

    m.bOK = True
    Me.Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmHumeOrder.cmdOK.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_KeyDown
'' Description: Show the help form if the user presses on F1
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
    RaiseError "frmHumeOrder.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: When the form is loaded, initialize the controls
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:
    
    CenterTheForm Me
'    Me.Icon = Picture16(ToolbarIcon("kHume"))
    Me.Icon = Picture16(ToolbarIcon("ID_TradeTracker"))
    
    lblBuySymbol1.Caption = ""
    lblBuySymbol2.Caption = ""
    lblSellSymbol1.Caption = ""
    lblSellSymbol2.Caption = ""
    
    lblBuyPrice1.Caption = ""
    lblBuyPrice2.Caption = ""
    lblSellPrice1.Caption = ""
    lblSellPrice2.Caption = ""
    
    lblBuyQuantity1.Caption = ""
    lblBuyQuantity2.Caption = ""
    lblSellQuantity1.Caption = ""
    lblSellQuantity2.Caption = ""
    
    rtbSample.Text = ""

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmHumeOrder.Form.Load", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Before the form is shown, fill in the controls with the values
''              that are passed in
'' Inputs:      Symbol1, Price1, Quantity1, Whether or not symbol1 is a buy,
''              Symbol2, Price2, Quantity2, Whether or not symbol2 is a buy,
''              order text
'' Returns:     True
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(ByVal pstrSymbol1$, ByVal pdPrice1$, ByVal plQuantity1&, _
                    ByVal pbBuySell1 As Boolean, Optional ByVal pstrSymbol2$ = "", _
                    Optional ByVal pdPrice2# = 0#, Optional ByVal plQuantity2& = 0&, _
                    Optional ByVal pbBuySell2 As Boolean = True, Optional ByVal pstrText$ = "") As Boolean
On Error GoTo ErrSection:
                    
    fraButtons.Visible = False
    cmdClose.Visible = True
                    
    lblSpecificOrders.Caption = "For Specific Orders, Call Your Broker"
                    
    If pbBuySell1 = True Then
        lblBuySymbol1.Caption = pstrSymbol1
        lblBuyPrice1.Caption = pdPrice1
        lblBuyQuantity1.Caption = Format(plQuantity1, "#,##0")
    Else
        lblSellSymbol1.Caption = pstrSymbol1
        lblSellPrice1.Caption = pdPrice1
        lblSellQuantity1.Caption = Format(plQuantity1, "#,##0")
    End If
    
    If pstrSymbol2 <> "" Then
        If pbBuySell2 = True Then
            If pbBuySell1 = True Then
                lblBuySymbol2.Caption = pstrSymbol2
                lblBuyPrice2.Caption = Format(pdPrice2, "$#,##0.00")
                lblBuyQuantity2.Caption = Format(plQuantity2, "#,##0")
            Else
                lblBuySymbol1.Caption = pstrSymbol2
                lblBuyPrice1.Caption = Format(pdPrice2, "$#,##0.00")
                lblBuyQuantity1.Caption = Format(plQuantity2, "#,##0")
            End If
        Else
            If pbBuySell1 = True Then
                lblSellSymbol1.Caption = pstrSymbol2
                lblSellPrice1.Caption = Format(pdPrice2, "$#,##0.00")
                lblSellQuantity1.Caption = Format(plQuantity2, "#,##0")
            Else
                lblSellSymbol2.Caption = pstrSymbol2
                lblSellPrice2.Caption = Format(pdPrice2, "$#,##0.00")
                lblSellQuantity2.Caption = Format(plQuantity2, "#,##0")
            End If
        End If
    End If
    
    rtbSample.Text = pstrText
    
    ShowForm Me, eForm_Modal, frmMain
    
    ShowMe = True
                    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmHumeOrder.ShowMe", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: Hide the form and let ShowMe unload it (if user clicked X)
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
    RaiseError "frmHumeOrder.Form.QueryUnload", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Resize
'' Description: Move and resize controls as the form is resized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Resize()
On Error Resume Next

    With cmdClose
        .Move (ScaleWidth / 2) - (.Width / 2)
    End With

    With fraButtons
        .Move (ScaleWidth / 2) - (.Width / 2)
    End With

End Sub
