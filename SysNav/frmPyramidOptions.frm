VERSION 5.00
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmPyramidOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pyramid Options"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7455
   Icon            =   "frmPyramidOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   7455
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   1095
      Left            =   6060
      TabIndex        =   1
      Top             =   180
      Width           =   1215
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
      Caption         =   "frmPyramidOptions.frx":014A
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmPyramidOptions.frx":0176
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmPyramidOptions.frx":0196
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Cancel          =   -1  'True
         Height          =   495
         Left            =   0
         TabIndex        =   10
         Top             =   600
         Width           =   1215
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
         Caption         =   "frmPyramidOptions.frx":01B2
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmPyramidOptions.frx":01E0
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmPyramidOptions.frx":0200
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Default         =   -1  'True
         Height          =   495
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   1215
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
         Caption         =   "frmPyramidOptions.frx":021C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmPyramidOptions.frx":0242
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmPyramidOptions.frx":0262
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraPyramidEnter 
      Height          =   855
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   5775
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
      Caption         =   "frmPyramidOptions.frx":027E
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmPyramidOptions.frx":02C0
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmPyramidOptions.frx":02E0
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniTextBoxXP txtContractsToEnter 
         Height          =   285
         Left            =   1620
         TabIndex        =   2
         Top             =   345
         Width           =   855
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmPyramidOptions.frx":02FC
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
         Tip             =   "frmPyramidOptions.frx":031C
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPyramidOptions.frx":033C
      End
      Begin HexUniControls.ctlUniLabelXP lblEnterText 
         Height          =   255
         Left            =   2640
         Top             =   360
         Width           =   2235
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
         Caption         =   "frmPyramidOptions.frx":0358
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmPyramidOptions.frx":03A0
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPyramidOptions.frx":03C0
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblContractsToEnter 
         Height          =   255
         Left            =   300
         Top             =   360
         Width           =   1275
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
         Caption         =   "frmPyramidOptions.frx":03DC
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmPyramidOptions.frx":041C
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPyramidOptions.frx":043C
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraPyramidExit 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   1140
      Width           =   5775
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
      Caption         =   "frmPyramidOptions.frx":0458
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmPyramidOptions.frx":0492
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmPyramidOptions.frx":04B2
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniRadioXP optTrade 
         Height          =   255
         Left            =   720
         TabIndex        =   8
         Top             =   1800
         Width           =   975
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
         Caption         =   "frmPyramidOptions.frx":04CE
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmPyramidOptions.frx":04F8
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmPyramidOptions.frx":0518
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniFrameWL fraPosition 
         Height          =   675
         Left            =   1080
         TabIndex        =   11
         Top             =   1080
         Width           =   4515
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
         Caption         =   "frmPyramidOptions.frx":0534
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmPyramidOptions.frx":0560
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPyramidOptions.frx":0580
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniTextBoxXP txtPercent 
            Height          =   285
            Left            =   1800
            TabIndex        =   7
            Top             =   360
            Width           =   855
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmPyramidOptions.frx":059C
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
            Tip             =   "frmPyramidOptions.frx":05BC
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmPyramidOptions.frx":05DC
         End
         Begin HexUniControls.ctlUniRadioXP optPercent 
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   360
            Width           =   1935
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
            Caption         =   "frmPyramidOptions.frx":05F8
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmPyramidOptions.frx":063E
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmPyramidOptions.frx":065E
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniTextBoxXP txtContractsToExit 
            Height          =   285
            Left            =   1800
            TabIndex        =   5
            Top             =   0
            Width           =   855
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmPyramidOptions.frx":067A
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
            Tip             =   "frmPyramidOptions.frx":069A
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmPyramidOptions.frx":06BA
         End
         Begin HexUniControls.ctlUniRadioXP optNumContracts 
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   15
            Width           =   1455
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
            Caption         =   "frmPyramidOptions.frx":06D6
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmPyramidOptions.frx":0714
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmPyramidOptions.frx":0734
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblExitText 
            Height          =   255
            Left            =   2760
            Top             =   15
            Width           =   1700
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
            Caption         =   "frmPyramidOptions.frx":0750
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmPyramidOptions.frx":0798
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmPyramidOptions.frx":07B8
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniRadioXP optPosition 
         Height          =   255
         Left            =   720
         TabIndex        =   3
         Top             =   840
         Width           =   975
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
         Caption         =   "frmPyramidOptions.frx":07D4
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmPyramidOptions.frx":0804
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmPyramidOptions.frx":0824
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblBasedOn 
         Height          =   255
         Left            =   240
         Top             =   480
         Width           =   1095
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
         Caption         =   "frmPyramidOptions.frx":0840
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmPyramidOptions.frx":087C
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPyramidOptions.frx":089C
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
End
Attribute VB_Name = "frmPyramidOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmPyramidOptions.frm
'' Description: Form to allow users to enter pyrading information for a rule
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Date         Author      Description
'' 04/17/2009   DAJ         Don't allow percent value lower than 5
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    bOK As Boolean                      ' Did the user OK the dialog?
    lUnits As Long                      ' Base units for the strategy
    strUnits As String                  ' Base units for the strategy
End Type
Private m As mPrivate

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Setup and show the form
'' Inputs:      Pyramid Information
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(ByVal bEnter As Boolean, bPosition As Boolean, _
        bPercent As Boolean, lNumContracts As Long, ByVal strRuleName As String, _
        ByVal lNumUnits As Long, ByVal strUnits As String) As Boolean
On Error GoTo ErrSection:

    If bEnter = True Then
        optPosition.Enabled = False
        optTrade.Enabled = False
        optNumContracts.Enabled = False
        optPercent.Enabled = False
        txtContractsToExit.Enabled = False
        txtPercent.Enabled = False
        lblExitText.Enabled = False
        lblContractsToEnter.Enabled = True
        txtContractsToEnter.Enabled = True
        lblEnterText.Enabled = True
        
        txtContractsToEnter.Text = Trim(CStr(lNumContracts))
        txtContractsToExit.Text = Trim(CStr(lNumContracts))
    Else
        optPosition.Enabled = True
        optTrade.Enabled = True
        optNumContracts.Enabled = True
        optPercent.Enabled = True
        txtContractsToExit.Enabled = True
        lblExitText.Enabled = True
        txtPercent.Enabled = True
        lblContractsToEnter.Enabled = False
        txtContractsToEnter.Enabled = False
        lblEnterText.Enabled = False
        
        ' For now the trade option is disabled
        If bPosition = False And optTrade.Visible = False Then
            bPosition = True
            bPercent = False
        End If
        
        If bPosition = True Then
            optPosition.Value = True
            If bPercent = True Then
                optPercent = True
                txtPercent.Text = Trim(CStr(lNumContracts))
                txtContractsToExit.Enabled = False
            Else
                optNumContracts.Value = True
                txtContractsToExit.Text = Trim(CStr(lNumContracts))
                txtContractsToEnter.Text = Trim(CStr(lNumContracts))
                txtPercent.Enabled = False
            End If
        Else
            optTrade.Value = True
        End If
    End If
    
    m.lUnits = lNumUnits
    m.strUnits = strUnits
    
    SetLabels
    
    SetEditorCaption Me, "Pyramiding Options", strRuleName
    ShowForm Me, eForm_Modal, frmMain
    
    If m.bOK = True Then
        If bEnter = True Then
            lNumContracts = CLng(ValOfText(txtContractsToEnter.Text))
            bPosition = True
            bPercent = False
        Else
            bPosition = optPosition.Value
            If bPosition Then
                bPercent = optPercent.Value
                If bPercent Then
                    lNumContracts = CLng(ValOfText(txtPercent.Text))
                Else
                    lNumContracts = CLng(ValOfText(txtContractsToExit.Text))
                End If
            Else
                lNumContracts = 0&
            End If
        End If
    End If
    
    ShowMe = m.bOK

ErrExit:
    Unload Me
    Exit Function

ErrSection:
    Unload Me
    RaiseError "frmPyramidOptions.ShowMe"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: User cancelled the dialog
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
    RaiseError "frmPyramidOptions.cmdCancel_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOK_Click
'' Description: User OK'd the dialog
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
On Error GoTo ErrSection:

    MoveFocus cmdOK

    ' If we have an entry, make sure that the user is entering at least one contract...
    If txtContractsToEnter.Enabled Then
        If ValOfText(txtContractsToEnter.Text) < 1 Then
            MoveFocus txtContractsToEnter
            Err.Raise vbObjectError + 1000, , "Must enter at least one contract"
        End If
        
    ' Otherwise make sure that the user is exiting at least a portion of the position...
    Else
        If optPercent.Value = True Then
            If ValOfText(txtPercent.Text) < kSN_MIN_ASPERCENT Or ValOfText(txtPercent.Text) > 100 Then
                MoveFocus txtPercent
                Err.Raise vbObjectError + 1000, , "Percent values can only be between " & Str(kSN_MIN_ASPERCENT) & " and 100 percent"
            ElseIf CLng(ValOfText(txtPercent.Text)) = 1 Then
                optNumContracts.Value = True
                txtContractsToExit.Text = "1"
            End If
        ElseIf optNumContracts.Value = True Then
            If ValOfText(txtContractsToExit.Text) < 1 Then
                MoveFocus txtContractsToExit
                Err.Raise vbObjectError + 1000, , "Must exit at least one contract"
            End If
        End If
    End If
    
    m.bOK = True
    Me.Hide

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmPyramidOptions.cmdOK_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Activate
'' Description: Set the focus to a control when the form gets the focus
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Activate()
On Error GoTo ErrSection:

    If txtContractsToEnter.Enabled Then
        MoveFocus txtContractsToEnter
    ElseIf txtContractsToExit.Enabled Then
        MoveFocus txtContractsToExit
    ElseIf txtPercent.Enabled Then
        MoveFocus txtPercent
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPyramidOptions.Form_Activate"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_KeyDown
'' Description: Show Help if user presses F1
'' Inputs:      Code of Key Pressed, Shift/Ctrl/Alt Status
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
    RaiseError "frmPyramidOptions.Form_KeyDown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Initialize things when the form is loaded
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    ' Hide this option for now
    optTrade.Visible = False
    CenterTheForm Me
    
    g.Styler.StyleForm Me

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmPyramidOptions.Form_Load"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: If user clicks on X, hide the form and let ShowMe unload it
'' Inputs:      Cancel the Unload?, Mode of the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode = 0 Then
        m.bOK = False
        Cancel = True
        Me.Hide
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmPyramidOptions.Form_QueryUnload"

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

    Dim lWidth As Long, lHeight As Long
    
    lWidth = fraPyramidEnter.Width + fraButtons.Width + (fraPyramidEnter.Left * 3)
    lHeight = fraPyramidEnter.Height + fraPyramidExit.Height + (fraPyramidEnter.Top * 3)
    
    If LimitFormSize(Me, lWidth, lHeight) Then Exit Sub
    
    With fraButtons
        .Move fraPyramidEnter.Width + fraPyramidEnter.Left * 2, fraPyramidEnter.Top
    End With
    
    With fraPyramidExit
        .Move fraPyramidEnter.Left, fraPyramidEnter.Height + fraPyramidEnter.Top * 2
    End With

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtContractsToEnter_Change
'' Description: Set the labels as the value in the control changes
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtContractsToEnter_Change()
On Error GoTo ErrSection:

    SetLabels
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPyramidOptions.txtContractsToEnter_Change"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optNumContracts_Click
'' Description: Fix other controls when this one changes
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optNumContracts_Click()
On Error GoTo ErrSection:

    txtPercent.Enabled = False
    txtContractsToExit.Enabled = True

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmPyramidOptions.optNumContracts_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optPercent_Click
'' Description: Fix other controls when this one changes
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optPercent_Click()
On Error GoTo ErrSection:

    txtPercent.Enabled = True
    txtContractsToExit.Enabled = False

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmPyramidOptions.optPercent_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optPosition_Click
'' Description: Fix other controls when this one changes
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optPosition_Click()
On Error GoTo ErrSection:
    
    optPercent.Enabled = True
    optNumContracts.Enabled = True
    txtPercent.Enabled = True
    txtContractsToExit.Enabled = True

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmPyramidOptions.optPosition_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optTrade_Click
'' Description: Fix other controls when this one changes
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optTrade_Click()
On Error GoTo ErrSection:

    optPercent.Enabled = False
    optNumContracts.Enabled = False
    txtPercent.Enabled = False
    txtContractsToExit.Enabled = False

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmPyramidOptions.optTrade_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtContractsToEnter_GotFocus
'' Description: Select all of the text when the control gets the focus
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtContractsToEnter_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtContractsToEnter

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmPyramidOptions.txtContractsToEnter_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtContractsToExit_Change
'' Description: Set the labels as the value of the text changes
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtContractsToExit_Change()
On Error GoTo ErrSection:

    SetLabels

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPyramidOptions.txtContractsToExit_Change"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtContractsToExit_GotFocus
'' Description: Select all of the text when the control gets the focus
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtContractsToExit_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtContractsToExit

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmPyramidOptions.txtContractsToExit_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtPercent_GotFocus
'' Description: Select all of the text when the control gets the focus
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtPercent_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtPercent

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmPyramidOptions.txtPercent_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetLabels
'' Description: Set the labels based on the selected controls
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetLabels()
On Error GoTo ErrSection:

    If m.lUnits = 1& Then
        lblEnterText.Caption = m.strUnits
        lblExitText.Caption = m.strUnits
    Else
        lblEnterText.Caption = " X " & CStr(m.lUnits) & " = " & _
                Format(CLng(ValOfText(txtContractsToEnter.Text)) * m.lUnits, "#,##0") & " " & m.strUnits
        lblExitText.Caption = " X " & CStr(m.lUnits) & " = " & _
                Format(CLng(ValOfText(txtContractsToExit.Text)) * m.lUnits, "#,##0") & " " & m.strUnits
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPyramidOptions.SetLabels"
    
End Sub

