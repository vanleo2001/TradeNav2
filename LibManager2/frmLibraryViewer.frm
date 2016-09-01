VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmLibraryViewer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Library"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9705
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   9705
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraDeveloperInfo 
      Caption         =   "Developer Information"
      Height          =   1995
      Left            =   2160
      TabIndex        =   2
      Top             =   4080
      Width           =   7395
      Begin VB.Label lblAuthor 
         Caption         =   "Author"
         Height          =   270
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   300
         Width           =   900
      End
      Begin VB.Label lblWebSite 
         Caption         =   "Web Site"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   708
         Width           =   900
      End
      Begin VB.Label lblEMail 
         Caption         =   "E-Mail"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   1108
         Width           =   900
      End
      Begin VB.Label lblPhoneNumber 
         Caption         =   "Phone Number"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   1508
         Width           =   1200
      End
      Begin VB.Label lblAuthor 
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Index           =   1
         Left            =   1740
         TabIndex        =   6
         Top             =   300
         Width           =   2505
      End
      Begin VB.Label lblWebSite 
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Index           =   1
         Left            =   1740
         TabIndex        =   5
         Top             =   700
         Width           =   2505
      End
      Begin VB.Label lblEMail 
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Index           =   1
         Left            =   1740
         TabIndex        =   4
         Top             =   1100
         Width           =   2505
      End
      Begin VB.Label lblPhoneNumber 
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Index           =   1
         Left            =   1740
         TabIndex        =   3
         Top             =   1500
         Width           =   2505
      End
   End
   Begin VB.Frame fraLibraryInfo 
      Caption         =   "Library Information"
      Height          =   3735
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   7395
      Begin RichTextLib.RichTextBox lblDetail 
         Height          =   1320
         Left            =   1140
         TabIndex        =   17
         Top             =   2100
         Width           =   6075
         _ExtentX        =   10716
         _ExtentY        =   2328
         _Version        =   393217
         BackColor       =   -2147483648
         ReadOnly        =   -1  'True
         TextRTF         =   $"frmLibraryViewer.frx":0000
      End
      Begin VB.Label lblLibrary 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   1
         Left            =   1140
         TabIndex        =   20
         Top             =   300
         Width           =   3540
      End
      Begin VB.Label lblLibrary 
         Caption         =   "Library"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   19
         Top             =   300
         Width           =   855
      End
      Begin VB.Label lblDetails 
         Caption         =   "Details"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label lblVersion 
         Caption         =   "Version"
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   16
         Top             =   1620
         Width           =   900
      End
      Begin VB.Label lblLastModified 
         Caption         =   "Last Modified"
         Height          =   285
         Index           =   0
         Left            =   2880
         TabIndex        =   15
         Top             =   1620
         Width           =   1110
      End
      Begin VB.Label lblVersion 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   1
         Left            =   1140
         TabIndex        =   14
         Top             =   1620
         Width           =   1620
      End
      Begin VB.Label lblLastModified 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   1
         Left            =   4080
         TabIndex        =   13
         Top             =   1620
         Width           =   2160
      End
      Begin VB.Label lblSummary 
         Caption         =   "Summary"
         Height          =   270
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   720
         Width           =   900
      End
      Begin VB.Label lblSummary 
         BorderStyle     =   1  'Fixed Single
         Height          =   780
         Index           =   1
         Left            =   1140
         TabIndex        =   11
         Top             =   720
         Width           =   6075
      End
   End
   Begin VB.PictureBox picLibrary 
      BorderStyle     =   0  'None
      Height          =   3600
      Left            =   120
      Picture         =   "frmLibraryViewer.frx":00D5
      ScaleHeight     =   3600
      ScaleWidth      =   1965
      TabIndex        =   0
      Top             =   120
      Width           =   1965
   End
End
Attribute VB_Name = "frmLibraryViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmLibraryViewer.frm
'' Description: Allow the user to view information on a library
''
'' Author:      Genesis Financial Data Services
''              425 E Woodmen Rd
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If KeyCode = vbKeyF1 Then
        KeyCode = 0
        g.Help.ShowF1Help Me
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibraryViewer.Form.KeyDown", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Initialize and the form and the controls
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:
    
    Icon = Picture16("kLibrary")
    CenterTheForm Me

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmLibraryViewer.Form.Load", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Initialize and show the form
'' Inputs:      Library ID to load
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowMe(ByVal lID As Long)
On Error GoTo ErrSection:

    Dim Library As cLibrary             ' Temporary library object
    
    Set Library = New cLibrary
    With Library
        .LibraryID = lID
        .Load
        
        lblLibrary(1).Caption = .LibraryName
        lblSummary(1).Caption = .LibraryDesc
        lblVersion(1).Caption = Str(.Version)
        lblLastModified(1).Caption = .LastModified
        lblAuthor(1).Caption = .Author
        lblWebSite(1).Caption = .WebSite
        lblEMail(1).Caption = .EMail
        lblPhoneNumber(1).Caption = .Phone
        lblDetail.FileName = .RtfFileName
        
        SetCaption
    End With
    ShowForm Me, True, g.frmOwner
    Unload Me

ErrExit:
    Set Library = Nothing
    Exit Sub
    
ErrSection:
    If Err.Number = 75 Then
        lblDetail.Text = "A library summary document was not found"
        Resume Next
    Else
        RaiseError "frmLibraryViewer.ShowMe", eGDRaiseError_Show, g.strAppPath
        Resume ErrExit
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetCaption
'' Description: Set the caption appropriately
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetCaption()
On Error GoTo ErrSection:
    
    If InStr(Me.Caption, " - ") = 0 Then
        Me.Caption = Me.Caption + " - [" & lblLibrary(1).Caption & "]"
    Else
        Me.Caption = Mid(Me.Caption, 1, InStr(Me.Caption, " - ") - 1) _
            & " - [" & lblLibrary(1).Caption & "]"
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    Err.Raise Err.Number, "frmLibraryViewer.SetCaption", Err.Description, g.strAppPath
    Resume ErrExit

End Sub


