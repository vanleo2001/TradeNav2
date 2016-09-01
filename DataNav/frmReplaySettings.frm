VERSION 5.00
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.16#0"; "gdOCX.ocx"
Begin VB.Form frmReplaySettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Stream Replay Date/Time"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   3930
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstSessions 
      Height          =   3375
      Left            =   180
      TabIndex        =   5
      Top             =   1140
      Width           =   2175
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   435
      Left            =   2700
      TabIndex        =   7
      Top             =   1560
      Width           =   915
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Start"
      Default         =   -1  'True
      Height          =   435
      Left            =   2700
      TabIndex        =   6
      Top             =   960
      Width           =   915
   End
   Begin gdOCX.gdSelectDate dtDate 
      Height          =   315
      Left            =   180
      TabIndex        =   1
      Top             =   360
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   556
      AllowWeekends   =   0   'False
      MaxDate         =   39042
      MaxDateIsToday  =   -1  'True
      Value           =   39029
   End
   Begin gdOCX.gdSelectDate dtTime 
      Height          =   315
      Left            =   2580
      TabIndex        =   3
      Top             =   360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      ShowDayOfWeek   =   0   'False
      ShowDate        =   0
      ShowTime        =   2
      MinDate         =   0
      MaxDate         =   0.99999
      Value           =   0.506944444444444
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Start Time"
      Height          =   255
      Left            =   2580
      TabIndex        =   2
      Top             =   120
      Width           =   1155
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Trading Session:"
      Height          =   255
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   1995
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sessions already downloaded:"
      Height          =   255
      Left            =   180
      TabIndex        =   4
      Top             =   840
      Width           =   2235
   End
End
Attribute VB_Name = "frmReplaySettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type mPrivate
    dDate As Double
End Type
Private m As mPrivate

Private Sub cmdCancel_Click()

    m.dDate = 0
    Unload Me

End Sub

Private Sub cmdOK_Click()

    m.dDate = 0
    Unload Me

End Sub

Private Sub Form_Load()

    CenterTheForm Me

End Sub

Public Function ShowMe(Optional ByVal dDate# = 0) As Double
On Error GoTo ErrSection:

    Dim i&, d#, s$
    Dim aFiles As New cGdArray

    m.dDate = dDate
    If dDate > 0 Then dtDate.Value = dDate

    lstSessions.Clear
    aFiles.GetMatchingFiles App.Path & "\RTS\*.rts", False
    aFiles.Sort eGdSort_IgnoreCase Or eGdSort_Descending
    For i = 0 To aFiles.Size - 1
        s = Parse(aFiles(i), ".", 1)
        If Len(s) >= 8 Then
            d = Val(s)
            d = DateSerial(Int(d / 10000), Int(d / 100) Mod 100, d Mod 100)
            s = DateFormat(d) & Format(d, "  ddd")
            lstSessions.AddItem s
            lstSessions.ItemData(lstSessions.ListCount - 1) = d
        End If
    Next

    ShowForm Me, eForm_Modal

ErrExit:
    ShowMe = m.dDate
    Exit Function
    
ErrSection:
    RaiseError "frmReplaySettings.ShowMe"
End Function

Private Sub lstSessions_Click()

    Dim d#
    
    d = lstSessions.ListIndex
    If d >= 0 And d < lstSessions.ListCount Then
        d = lstSessions.ItemData(d)
        If d > 0 And d < 999999 Then
            dtDate.Value = d
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReplaySettings.lstSessions_Click"
End Sub
