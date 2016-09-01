VERSION 5.00
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmEditDate 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   480
   ScaleWidth      =   3900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   315
      Left            =   2340
      TabIndex        =   1
      Top             =   60
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
      Caption         =   "frmEditDate.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmEditDate.frx":002C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmEditDate.frx":004C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Default         =   -1  'True
         Height          =   315
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   555
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
         Caption         =   "frmEditDate.frx":0068
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmEditDate.frx":008E
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmEditDate.frx":00AE
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Cancel          =   -1  'True
         Height          =   315
         Left            =   660
         TabIndex        =   2
         Top             =   0
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
         Caption         =   "frmEditDate.frx":00CA
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmEditDate.frx":00F8
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmEditDate.frx":0118
         RightToLeft     =   0   'False
      End
   End
   Begin gdOCX.gdSelectDate gdDate 
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   556
      ShowPM          =   2
   End
End
Attribute VB_Name = "frmEditDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmEditDate.frm
'' Description: Allows the user to edit the date with the Genesis Date Control
''
'' Author:      Genesis Financial Data Services
''              425 E Woodmen Rd
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    bOK As Boolean
End Type
Private m As mPrivate

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Initialize and show the form
'' Inputs:      Left and Top Location of the Form, Starting value for Date
'' Returns:     Date passed if Cancelled, Date selected otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(ByVal dLeft As Double, ByVal dTop As Double, ByVal dDefault As Double, _
            Optional OwnerForm As Form = Nothing, Optional dMaxDate As Double = 401768, _
            Optional bMaxDateIsToday As Boolean = False, Optional dMinDate As Double = 0, _
            Optional bMinDateIsToday As Boolean = False, Optional bShowCalendar As Boolean = True, _
            Optional iShowDate As eDateDisplays = YearMonthDay, Optional bShowDOW As Boolean = True, _
            Optional iShowTime As eTimeDisplays = NoTime, Optional iShowPM As eAmPmDisplays = Default12or24, _
            Optional bAllowWeekends As Boolean = True, Optional bCancelled As Boolean) As Double
On Error GoTo ErrSection:

    Move dLeft, dTop
    
    With gdDate
        .AllowWeekends = bAllowWeekends
        .MaxDate = dMaxDate
        .MaxDateIsToday = bMaxDateIsToday
        .MinDate = dMinDate
        .MinDateIsToday = bMinDateIsToday
        .ShowCalendar = bShowCalendar
        .ShowDate = iShowDate
        .ShowDayOfWeek = bShowDOW
        .ShowPM = iShowPM
        .ShowTime = iShowTime
        .Value = dDefault
    End With
    
    If iShowTime = NoTime Then
        Width = 3900
    Else
        Width = 4830
    End If
    
    If Not OwnerForm Is Nothing Then
        Show vbModal, OwnerForm
    Else
        Show vbModal
    End If
    
    bCancelled = Not m.bOK
    If m.bOK Then ShowMe = gdDate.Value Else ShowMe = dDefault
    
ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    RaiseError "frmEditDate.ShowMe", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: Unload the form without saving
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancel_Click()
On Error GoTo ErrSection:

    MoveFocus cmdCancel
    m.bOK = False
    Me.Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditDate.cmdCancel.Click", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOK_Click
'' Description: Unload the form and return the users value
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
On Error GoTo ErrSection:

    MoveFocus cmdOK
    m.bOK = True
    Me.Hide
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditDate.cmdOK.Click", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

Private Sub Form_Load()
    g.Styler.StyleForm Me
    
End Sub

Private Sub Form_Resize()
On Error Resume Next

    With fraButtons
        .Move ScaleWidth - .Width - gdDate.Left
    End With
    
    With gdDate
        .Move .Left, .Top, ScaleWidth - fraButtons.Width - (.Left * 3)
    End With

End Sub


