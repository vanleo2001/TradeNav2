VERSION 5.00
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmAskDate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Date"
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3615
   ControlBox      =   0   'False
   Icon            =   "frmAskDate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   3615
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL Frame1 
      Height          =   1035
      Left            =   180
      TabIndex        =   3
      Top             =   540
      Width           =   3255
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
      Caption         =   "frmAskDate.frx":000C
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmAskDate.frx":002C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmAskDate.frx":004C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Default         =   -1  'True
         Height          =   375
         Left            =   660
         TabIndex        =   1
         Top             =   600
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
         Caption         =   "frmAskDate.frx":0068
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmAskDate.frx":008E
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmAskDate.frx":00AE
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   1740
         TabIndex        =   2
         Top             =   600
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
         Caption         =   "frmAskDate.frx":00CA
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmAskDate.frx":00F8
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmAskDate.frx":0118
         RightToLeft     =   0   'False
      End
      Begin gdOCX.gdSelectDate gdSelectDate1 
         Height          =   345
         Left            =   480
         TabIndex        =   0
         Top             =   60
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   609
      End
   End
   Begin HexUniControls.ctlUniLabelXP lblMessage 
      Height          =   255
      Left            =   285
      Top             =   120
      Width           =   3045
      _ExtentX        =   5371
      _ExtentY        =   344
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmAskDate.frx":0134
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   -1  'True
      Tip             =   "frmAskDate.frx":0162
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmAskDate.frx":0182
      RightToLeft     =   0   'False
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmAskDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmAskDate.frm
'' Description: Queries user for a date
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 02/02/2015   MJM         Initial code
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private Type mPrivate
    bOK As Boolean                      ' Did the user click on the OK button?
End Type
Private m As mPrivate

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: When the form is called from someone else, do some initialization,
''              show the form, then do some afterwards processing
'' Inputs:      Default date as double
'' Returns:     Selected date as double
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(ByVal strMsg As String, Optional ByVal dDefaultDate As Double = 0, _
        Optional ByVal bAllowWeekends As Boolean = True, Optional ByVal dMinDate As Double = 0, Optional ByVal dMaxDate As Double = 0) As Double
On Error GoTo ErrSection:
    
    Dim dDateSelected As Double, nAddHeight&

    lblMessage.Caption = Replace(strMsg, "|", vbCrLf)
    If dDefaultDate <= 0 Then
        dDefaultDate = Date
    End If
    
    With gdSelectDate1
        .Value = dDefaultDate
        .AllowWeekends = bAllowWeekends
        If dMinDate > 0 Then
            .MinDate = dMinDate
        End If
        If dMaxDate > 0 Then
            .MaxDate = dMaxDate
        End If
    End With

    ' Show the form
    nAddHeight = lblMessage.Height - 255
    If nAddHeight > 0 Then
        Frame1.Top = Frame1.Top + nAddHeight
        Me.Height = Me.Height + nAddHeight
    End If
    CenterTheForm Me
    ShowForm Me, eForm_ActModal, frmMain
    
    ' If the user pressed OK, save then return date
    If m.bOK Then
        dDateSelected = gdSelectDate1.Value
    End If

    ' Return the OK status and unload the form
    ShowMe = dDateSelected

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmAskDate.ShowMe", eGDRaiseError_Raise


End Function

Private Sub cmdCancel_Click()
On Error GoTo ErrSection:

    Me.Hide

ErrExit:
    Exit Sub
    
ErrSection:
    Unload Me
    RaiseError "frmAskDate.cmdCancel_Click"

End Sub

Private Sub cmdOK_Click()
On Error GoTo ErrSection:
    
    m.bOK = True
    Me.Hide

ErrExit:
    Exit Sub
    
ErrSection:
    Unload Me
    RaiseError "frmAskDate.cmdOK_Click"

End Sub


Private Sub Form_Load()
    g.Styler.StyleForm Me
End Sub
