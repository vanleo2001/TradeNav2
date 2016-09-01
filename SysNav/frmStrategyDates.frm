VERSION 5.00
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmStrategyDates 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   3690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   435
      Left            =   518
      TabIndex        =   7
      Top             =   1500
      Width           =   2655
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
      Caption         =   "frmStrategyDates.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmStrategyDates.frx":002C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmStrategyDates.frx":004C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Height          =   435
         Left            =   0
         TabIndex        =   0
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
         Caption         =   "frmStrategyDates.frx":0068
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmStrategyDates.frx":008E
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmStrategyDates.frx":00AE
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Height          =   435
         Left            =   1440
         TabIndex        =   3
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
         Caption         =   "frmStrategyDates.frx":00CA
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmStrategyDates.frx":00F8
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmStrategyDates.frx":0118
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraToDate 
      Height          =   675
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   3195
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
      Caption         =   "frmStrategyDates.frx":0134
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmStrategyDates.frx":0160
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmStrategyDates.frx":0180
      RightToLeft     =   0   'False
      Begin gdOCX.gdSelectDate gdToDate 
         Height          =   315
         Left            =   900
         TabIndex        =   5
         Top             =   0
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
      End
      Begin HexUniControls.ctlUniRadioXP optToEnd 
         Height          =   255
         Left            =   600
         TabIndex        =   6
         Top             =   360
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
         Caption         =   "frmStrategyDates.frx":019C
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   -1  'True
         Tip             =   "frmStrategyDates.frx":01D8
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmStrategyDates.frx":01F8
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optToDate 
         Height          =   255
         Left            =   600
         TabIndex        =   4
         Top             =   30
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
         Caption         =   "frmStrategyDates.frx":0214
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmStrategyDates.frx":0246
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmStrategyDates.frx":0266
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblTo 
         Height          =   255
         Left            =   0
         Top             =   60
         Width           =   495
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
         Caption         =   "frmStrategyDates.frx":0282
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmStrategyDates.frx":02A8
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmStrategyDates.frx":02C8
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin gdOCX.gdSelectDate gdFromDate 
      Height          =   315
      Left            =   840
      TabIndex        =   1
      Top             =   180
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   556
   End
   Begin HexUniControls.ctlUniLabelXP lblFrom 
      Height          =   255
      Left            =   240
      Top             =   210
      Width           =   495
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
      Caption         =   "frmStrategyDates.frx":02E4
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmStrategyDates.frx":030E
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmStrategyDates.frx":032E
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
End
Attribute VB_Name = "frmStrategyDates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmStrategyDates.frm
'' Description: Allow the user to set dates to run a strategy
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 08/02/2013   DAJ         Force the 'ToDate' to be after the 'FromDate'
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    bOK As Boolean                      ' Did the user OK the dialog?
End Type
Private m As mPrivate

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Initialize and show the Form
'' Inputs:      From Date, To Date, To End of Data
'' Returns:     True on OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(dFromDate As Double, dToDate As Double, bToEndOfData As Boolean) As Boolean
On Error GoTo ErrSection:

    gdFromDate.Value = dFromDate
    gdToDate.Value = dToDate
    optToEnd.Value = bToEndOfData
    optToDate.Value = Not bToEndOfData

    ShowForm Me, eForm_Modal, frmMain
    
    If m.bOK = True Then
        dFromDate = gdFromDate.Value
        dToDate = gdToDate.Value
        bToEndOfData = optToEnd.Value
    End If

    ShowMe = m.bOK

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmStrategyDates.ShowMe"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: Handle the user clicking on the Cancel button
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancel_Click()
On Error GoTo ErrSection:

    m.bOK = False
    Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStrategyDates.cmdCancel_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOK_Click
'' Description: Handle the user clicking on the OK button
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
On Error GoTo ErrSection:

    If (optToDate.Value = True) And (gdToDate.Value <= gdFromDate.Value) Then
        InfBox "The to date must be after the from date", "!", , Caption & " Error"
        MoveFocus gdToDate
    Else
        m.bOK = True
        Hide
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStrategyDates.cmdOK_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Initialize the form when it is loaded
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Caption = "Strategy Dates"
    Icon = Picture16("kSystem")
    CenterTheForm Me
    
    g.Styler.StyleForm Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStrategyDates.Form_Load"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: Handle the user clicking on the X
'' Inputs:      Cancel Unload?, Mode of the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode <> vbFormCode Then
        m.bOK = False
        Cancel = True
        Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStrategyDates.Form_QueryUnload"
    
End Sub

