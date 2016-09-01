VERSION 5.00
Object = "{3B008041-905A-11D1-B4AE-444553540000}#1.0#0"; "Vsocx6.ocx"
Begin VB.Form frmDataInstall 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Data Install"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6660
   Icon            =   "frmDataInstall.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   6660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrUnload 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3240
      Top             =   3240
   End
   Begin VB.TextBox txtHwnd 
      Height          =   285
      Left            =   3960
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   3240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame fraSpace 
      Height          =   705
      Left            =   240
      TabIndex        =   13
      Top             =   2580
      Width           =   2595
      Begin VB.Label lblSpaceAvailable 
         Alignment       =   1  'Right Justify
         Caption         =   "1234 MB"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   1500
         TabIndex        =   17
         Top             =   420
         Width           =   915
      End
      Begin VB.Label Label8 
         Caption         =   "Space Available:"
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   120
         TabIndex        =   16
         Top             =   420
         Width           =   1335
      End
      Begin VB.Label lblSpaceRequired 
         Alignment       =   1  'Right Justify
         Caption         =   "450 MB"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   1500
         TabIndex        =   15
         Top             =   180
         Width           =   915
      End
      Begin VB.Label Label5 
         Caption         =   "Space Required:"
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   120
         TabIndex        =   14
         Top             =   180
         Width           =   1335
      End
   End
   Begin VB.Frame fraPaths 
      Caption         =   "Installation Directories"
      Height          =   975
      Left            =   240
      TabIndex        =   7
      Top             =   3720
      Visible         =   0   'False
      Width           =   5295
      Begin VB.CommandButton cmdTo 
         Caption         =   "To:"
         Height          =   270
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton cmdFrom 
         Caption         =   "From:"
         Height          =   270
         Left            =   120
         TabIndex        =   8
         Top             =   270
         Width           =   615
      End
      Begin VB.Label lblTo 
         Caption         =   "C:\Program Files\Genesis\Navigator Suite\DataHistory\"
         Height          =   195
         Left            =   840
         TabIndex        =   20
         Top             =   630
         Width           =   4275
      End
      Begin VB.Label lblFrom 
         Caption         =   "D:\DATA\"
         Height          =   195
         Left            =   840
         TabIndex        =   9
         Top             =   300
         Width           =   4155
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   435
      Left            =   5160
      TabIndex        =   6
      Top             =   2760
      Width           =   1035
   End
   Begin VB.CommandButton cmdInstall 
      Caption         =   "Install &Data"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3720
      TabIndex        =   5
      Top             =   2760
      Width           =   1275
   End
   Begin VB.Frame fraAmount 
      Caption         =   "Amount to Hard Drive"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   2880
      TabIndex        =   1
      Top             =   240
      Width           =   3435
      Begin VB.OptionButton optAll 
         Caption         =   "&Entire history"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   4
         Top             =   1620
         Width           =   1575
      End
      Begin VB.OptionButton optYear 
         Caption         =   "&One year of data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   3
         Top             =   1140
         Width           =   1815
      End
      Begin VB.OptionButton optMinimal 
         Caption         =   "Mi&nimal amount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   2
         Top             =   660
         Width           =   1695
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         Caption         =   "Please choose amount of End of Day data to install."
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "(faster Criteria, CD required)"
         Height          =   195
         Left            =   480
         TabIndex        =   12
         Top             =   1365
         Width           =   2355
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "(slower Criteria, CD required)"
         Height          =   195
         Left            =   480
         TabIndex        =   11
         Top             =   885
         Width           =   2355
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "(CD not required to run program)"
         Height          =   195
         Left            =   480
         TabIndex        =   10
         Top             =   1845
         Width           =   2355
      End
   End
   Begin VB.Frame fraDataTypes 
      Caption         =   "Data Types to Install"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2415
      Begin VB.Frame fraEodData 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1335
         Left            =   600
         TabIndex        =   23
         Top             =   540
         Width           =   1575
         Begin VB.CheckBox chkIdxEod 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   0
            TabIndex        =   27
            Top             =   367
            Value           =   1  'Checked
            Width           =   255
         End
         Begin VB.CheckBox chkStkEod 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   0
            TabIndex        =   26
            Top             =   727
            Value           =   1  'Checked
            Width           =   255
         End
         Begin VB.CheckBox chkFutEod 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   0
            TabIndex        =   25
            Top             =   7
            Value           =   1  'Checked
            Width           =   255
         End
         Begin VB.CheckBox chkMutEod 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   0
            TabIndex        =   24
            Top             =   1087
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label lblIndexes 
            Caption         =   "Indexes"
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
            Left            =   360
            TabIndex        =   30
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label lblFutures 
            Caption         =   "Futures"
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
            Left            =   360
            TabIndex        =   31
            Top             =   0
            Width           =   1095
         End
         Begin VB.Label lblStocks 
            Caption         =   "Stocks"
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
            Left            =   360
            TabIndex        =   29
            Top             =   720
            Width           =   735
         End
         Begin VB.Label lblMutFunds 
            Caption         =   "Mutual Funds"
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
            Left            =   360
            TabIndex        =   28
            Top             =   1080
            Visible         =   0   'False
            Width           =   1215
         End
      End
      Begin VB.Frame fraTickData 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1335
         Left            =   240
         TabIndex        =   32
         Top             =   240
         Width           =   735
         Begin VB.CheckBox chkFutTick 
            Height          =   240
            Left            =   60
            TabIndex        =   35
            Top             =   300
            Width           =   255
         End
         Begin VB.CheckBox chkIdxTick 
            Height          =   240
            Left            =   60
            TabIndex        =   34
            Top             =   660
            Width           =   255
         End
         Begin VB.CheckBox chkStkTick 
            Height          =   240
            Left            =   60
            TabIndex        =   33
            Top             =   1020
            Width           =   255
         End
         Begin VB.Label lblEOD 
            Caption         =   "EOD"
            Height          =   255
            Left            =   360
            TabIndex        =   37
            Top             =   0
            Width           =   375
         End
         Begin VB.Label lblTick 
            Caption         =   "Tick"
            Height          =   255
            Left            =   0
            TabIndex        =   36
            Top             =   0
            Width           =   375
         End
      End
   End
   Begin vsOcx6LibCtl.vsElastic vsProgress 
      Height          =   315
      Left            =   180
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   2820
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   3
      MousePointer    =   0
      _ConvInfo       =   1
      Version         =   600
      BackColor       =   8421504
      ForeColor       =   16777215
      FloodColor      =   16711680
      ForeColorDisabled=   -2147483631
      Caption         =   "50%"
      Align           =   0
      Appearance      =   3
      AutoSizeChildren=   0
      BorderWidth     =   2
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   1
      FloodPercent    =   50
      CaptionPos      =   4
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   0   'False
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   0
      GridCols        =   0
      _GridInfo       =   ""
   End
End
Attribute VB_Name = "frmDataInstall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmDataInstall.frm
'' Description: Performs a data install according to what the user chooses
''
'' Author:      Genesis Financial Data Services
''              425 Woodmen Rd
''              Colorado Springs, CO 80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date      Author      Description
''
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private Enum eInstallDataType
    eInstallDataType_EOD = 0
    eInstallDataType_Tick
End Enum

Private Type mPrivate
    bTick As Boolean                    ' Does the CD have Tick Data?
    lHandle As Long                     ' Handle from the install
    astrFileSizes As cGdArray           ' File size array from the file
    strInstall As String                ' Install string from the CD
    strDefault As String                ' Defaults string from the CD
    strDataServID As String             ' Data Service ID from the CD (if there)
    strPassword As String               ' Password from the CD (if there)
    bMultipleDisks As Boolean           ' true if data on multiple disks
    strStarterDate As String            ' Date of the Starter.GZP
End Type
Private m As mPrivate

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkFutEod_Click
'' Description: If the user clicks on the Futures End of day, figure out the
''              file size requirements.  If the user turns the end of day
''              off then make sure that the tick is off also.
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkFutEod_Click()
On Error GoTo ErrSection:

    If chkFutEod.Value = vbUnchecked And chkFutTick.Value = vbChecked Then
        chkFutTick.Value = vbUnchecked
    Else
        EnableControls
        FigureSize
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDataInstall.chkFutEod.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkFutTick_Click
'' Description: If the user clicks on the Futures Tick, figure out the file
''              size requirements.  If the user turns it on, make sure that
''              the end of day is also on
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkFutTick_Click()
On Error GoTo ErrSection:

    If chkFutEod.Value = vbUnchecked And chkFutTick.Value = vbChecked Then
        chkFutEod.Value = vbChecked
    Else
        EnableControls
        FigureSize
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDataInstall.chkFutTick.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkIdxEod_Click
'' Description: If the user clicks on the Indexes End of day, figure out the
''              file size requirements.  If the user turns the end of day
''              off then make sure that the tick is off also.
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkIdxEod_Click()
On Error GoTo ErrSection:

    If chkIdxEod.Value = vbUnchecked And chkIdxTick.Value = vbChecked Then
        chkIdxTick.Value = vbUnchecked
    Else
        EnableControls
        FigureSize
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDataInstall.chkIdxEod.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkIdxTick_Click
'' Description: If the user clicks on the Index Tick, figure out the file
''              size requirements.  If the user turns it on, make sure that
''              the end of day is also on
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkIdxTick_Click()
On Error GoTo ErrSection:

    If chkIdxEod.Value = vbUnchecked And chkIdxTick.Value = vbChecked Then
        chkIdxEod.Value = vbChecked
    Else
        EnableControls
        FigureSize
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDataInstall.chkIdxTick.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkMutEod_Click
'' Description: If the user clicks on the Mutual Funds check box, figure out the
''              file size requirements
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkMutEod_Click()
On Error GoTo ErrSection:

    EnableControls
    FigureSize

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDataInstall.chkMutEod.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkStkEod_Click
'' Description: If the user clicks on the Stock End of day, figure out the
''              file size requirements.  If the user turns the end of day
''              off then make sure that the tick is off also.
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkStkEod_Click()
On Error GoTo ErrSection:

    If chkStkEod.Value = vbUnchecked And chkStkTick.Value = vbChecked Then
        chkStkTick.Value = vbUnchecked
    Else
        EnableControls
        FigureSize
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDataInstall.chkStkEod.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkStkTick_Click
'' Description: If the user clicks on the Stock Tick, figure out the file
''              size requirements.  If the user turns it on, make sure that
''              the end of day is also on
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkStkTick_Click()
On Error GoTo ErrSection:

    If chkStkEod.Value = vbUnchecked And chkStkTick.Value = vbChecked Then
        chkStkEod.Value = vbChecked
    Else
        EnableControls
        FigureSize
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDataInstall.chkStkTick.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdFrom_Click
'' Description: If the user clicks on the From button, bring up a Windows
''              common Open dialog to allow them to browse for the Starter.GZP
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdFrom_Click()
On Error GoTo ErrSection:

    Dim strFile As String               ' Filename and path of the Starter.GZP
    Dim strTemp As String               ' Temporary string variable
    Dim d As Double                     ' Date of the last CD
    Dim strCdInDrive As String          ' CD drive where Genesis CD is in
    Dim strPath As String               ' Path of the CD drive
    
    strCdInDrive = GenesisCDInDrive
    If Len(strCdInDrive) > 0 Then strPath = strCdInDrive & ":\Data\Starter.GZP" Else strPath = ""
    
    strFile = CommonDialogFile(frmMain.CommonDialog1, False, "STARTER.GZP|STARTER.GZP", strPath, "Where is the data to install ...", cdlOFNFileMustExist Or cdlOFNHideReadOnly)
    
    ' make sure not older than CD data
    strTemp = GetCDDataInf
    d = Val(Parse(strTemp, vbTab, 2))
    If FileDate(strFile) < d - 1# / 1400# Then
        InfBox "Can't install data older than the|most recent installation CD.", "e", , "Error"
        strFile = ""
    End If
    CheckInstallData strFile

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDataInstall.cmdFrom.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: If the user clicks on cancel, either unload the form or cancel
''              the install
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancel_Click()
On Error GoTo ErrSection:

    ' If an install hasn't started yet, unload the form
    If cmdInstall.Visible Then
        Unload Me
        
    ' Otherwise cancel the download, then unload the form
    Else
        Screen.MousePointer = vbNormal
        If DM_DistCancel(g.DMS, m.lHandle) <> 0 Then
            InfBox "Data Install Cancelled", "i", , "Data Installation"
        End If
        
        KillFile DataPath & "*.*"
        Unload Me
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDataInstall.cmdCancel.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EnableControls
'' Description: If the optAll option is on, then turn on the lblTo, otherwise
''              turn the lblTo off
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EnableControls()
On Error GoTo ErrSection:

    If optAll Then
        Enable cmdTo
        lblTo.Visible = True
    Else
        Disable cmdTo
        lblTo.Visible = False
    End If
    
    If optAll + optMinimal + optYear <> 0 And Len(cmdInstall.Tag) > 0 Then
    'If Len(cmdInstall.Tag) > 0 Then
        Enable cmdInstall
    Else
        Disable cmdInstall
    End If
    
    If m.bTick And (chkFutTick = vbChecked Or chkStkTick = vbChecked Or chkIdxTick = vbChecked) Then
        If Left(lblInfo.Caption, 4) <> "Full" Then
            lblInfo.Caption = "Full history of End of Day will be installed.  Please choose amount of tick data to install."
            
            optYear.Caption = "&6 months of data"
        End If
    Else
        If Left(lblInfo.Caption, 4) = "Full" Then
            lblInfo.Caption = "Please choose amount of End of Day data to install."
            
            optYear.Caption = "&One year of data"
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDataInstall.EnableControls", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdInstall_Click
'' Description: If the user clicks on the Install button, start the install
''              using the options that they chose
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdInstall_Click()
On Error GoTo ErrSection:

    Dim strAuthString As String         ' Authorization string for install
    Dim aStrings As New cGdArray        ' Temp array to go to a file
    Dim strDataPath$
       
    ' Need to do this for "Default" buttons, since hitting 'Enter'
    ' from another control does not trigger it's 'LostFocus' event.
    MoveFocus cmdInstall

    If optAll + optYear + optMinimal = 0 Then
        Beep
        InfBox "Please select an option for|'Amount to Hard Drive'", "e", , "Data Install"
        Exit Sub
    End If
    
    ' Hide the controls that are no longer needed
    fraPaths.Enabled = False
    fraDataTypes.Enabled = False
    fraAmount.Enabled = False
    fraSpace.Visible = False
    cmdInstall.Visible = False
    
    Me.Caption = "Installing Data ..."
    Screen.MousePointer = vbHourglass
    
    ' Show the progress bar
    vsProgress.Visible = True
    vsProgress.FloodPercent = 0
    vsProgress.Caption = "Initializing data install ..."
    
    ' Build the authorization string
    strAuthString = BuildAuthorization
    
    ' Get the data manager ready for the install
    DM_Init False
    g.Universe.CloseDb
    
    strDataPath = DataPath
    
    ClearReadOnlyFlags AddSlash(strDataPath) & "*.*"
    
    ' Wipe out the temporary "Old" folder
    KillFile strDataPath & "Old\*.*"
    
    ' Clean out the Data directory and unzip the Starter.GZP into it
    KillFile AddSlash(App.Path) & "SYMPOOL.MEM"
    KillFile strDataPath & "*.*"
    m.strStarterDate = Str(CDbl(FileDate(AddSlash(lblFrom) & "Starter.GZP")))
    ZipExecute "U", AddSlash(lblFrom.Caption) & "STARTER.GZP", strDataPath
    If Not FileExist(strDataPath & "SYMBOLS.DBF") Then
        'zip file must have been corrupted or something!
        InfBox "CD data starter files could not be unzipped.", "e", , "Error"
        KillFile strDataPath & "*.*"
        Unload Me
        Exit Sub
    End If
    
    ' Start the install
    aStrings(0) = lblFrom.Caption
    aStrings(1) = strAuthString
    aStrings(2) = Str(txtHwnd.hWnd)
    aStrings.ToFile App.Path & "\Chk\Install.chk"
    vsProgress.Caption = "0%"
    m.lHandle = DM_Install(strDataPath, lblFrom.Caption, _
            strDataPath & "SETUP.SDM", _
            strAuthString & "|" & m.strInstall, txtHwnd.hWnd, Me.hWnd)
                            
    ' If the install failed, warn the user and clean out the data directory
    If m.lHandle = 0 Then
        InfBox "Error Installing Data", "e", , "Error"
        KillFile strDataPath & "*.*"
        Unload Me
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDataInstall.cmdInstall.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Activate
'' Description: When the form is activated for the first time, set defaults
''              if this is a Hume install, and click Install for them
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Activate()
On Error GoTo ErrSection:

    Static bAlreadyDone As Boolean
    
    If Not bAlreadyDone Then
        bAlreadyDone = True
        'check for first-time data install
        If g.SymbolPool.NumRecords = 0 Then
' Commented out 11/2/2001 by DAJ for Hume People
'            If FileExist(App.Path & "\Hume.mod") Then
'                chkFutEod = 1
'                chkIdxEod = 0
'                chkStkEod = 0
'                chkMutEod = 0
'                optAll = True
'                EnableControls
'                Me.Refresh
'                If cmdInstall.Enabled Then cmdInstall_Click
'            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDataInstall.Form.Activate", eGDRaiseError_Show
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
    RaiseError "frmDataInstall.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: When the form is loaded, center it and initialize some controls
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Dim strStarter$
    Dim strTemp As String
   
    CenterTheForm Me
    
    Me.Height = fraSpace.Top + fraSpace.Height + 600
   
    SetToPath AddSlash(App.Path) & "Data"
    
    lblFutures.ToolTipText = "Futures (Commodities) - individual and continuous contracts"
    lblStocks.ToolTipText = "History for over 10,000 stock symbols"
    lblIndexes.ToolTipText = "History for common Market Indices"
    lblMutFunds.ToolTipText = "History for Mutual Funds"
    
    chkFutTick.ToolTipText = "Tick by Tick Data for Futures"
    chkStkTick.ToolTipText = "Tick by Tick Data for Stocks"
    chkIdxTick.ToolTipText = "Tick by Tick Data for Indices"
    chkFutEod.ToolTipText = "End of Day Data for Futures"
    chkStkEod.ToolTipText = "End of Day Data for Stocks"
    chkIdxEod.ToolTipText = "End of Day Data for Indices"
    chkMutEod.ToolTipText = "End of Day Data for Mutual Funds"
    
    strStarter = Parse(GetCDDataInf, vbTab, 1)
    If Len(Trim(strStarter)) = 0 Then
        cmdFrom_Click
'        If Len(lblFrom.Caption) = 0 Then Unload Me
    Else
        CheckInstallData strStarter
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDataInstall.Form.Load", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: When the form is unloaded, if there are no records in the
''              symbol pool, warn the user
'' Inputs:      Whether or not to cancel the unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    Dim strTemp As String               ' Temporary string variable
    
    If g.SymbolPool.NumRecords = 0 Then
        If Me.Visible Then Me.Hide
        strTemp = vbCrLf & "Data could not be found." & vbCrLf & vbCrLf _
            & "Please select 'File', 'Install Data' from the menu" & vbCrLf _
            & "in order to find the data to install."
        frmMessage.ShowMe "Message", strTemp, eNormalMessage, True
    End If
    
    Set m.astrFileSizes = Nothing

    ' need this so last events still have an existing window to return to
    DoEvents

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDataInstall.Form.Unload", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Label5_Click
'' Description: If the user clicks on the space required label, bring up the
''              Paths frame
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Label5_Click()
On Error GoTo ErrSection:

    Me.Height = fraPaths.Height + fraPaths.Top + 600
    fraPaths.Visible = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDataInstall.Label5.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optAll_Click
'' Description: If the user clicks on the All option, enable the other controls
''              as necessary and figure out the file size requirements
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optAll_Click()
On Error GoTo ErrSection:

    EnableControls
    FigureSize
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDataInstall.optAll.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optMinimal_Click
'' Description: If the user clicks on the Minimal option, enable the other
''              controls as necessary and figure out the file size requirements
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optMinimal_Click()
On Error GoTo ErrSection:
    
    EnableControls
    FigureSize
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDataInstall.optMinimal.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optYear_Click
'' Description: If the user clicks on the Year option, enable the other controls
''              as necessary and figure out the file size requirements
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optYear_Click()
On Error GoTo ErrSection:
    
    EnableControls
    FigureSize

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDataInstall.optYear.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BuildAuthorization
'' Description: Builds the installation authorization string according to the
''              options the user chose
'' Inputs:      None
'' Returns:     Authorization String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function BuildAuthorization() As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return string
    Dim strInstallFile As String
    
    strReturn = ""
    g.strInstalledString = ""
    
    ' look for special override file
    strInstallFile = App.Path & "\Install.ovr"
    If FileExist(strInstallFile) Then
        BuildAuthorization = FileToString(strInstallFile, , True)
        Exit Function
    End If
    
    If chkFutEod.Value = vbChecked Then
        strReturn = strReturn & "F,E," & GetNumMonths(eInstallDataType_EOD) & "|"
        g.strInstalledString = g.strInstalledString & ",F,"
    End If
        
    If chkStkEod.Value = vbChecked Then
        strReturn = strReturn & "S,E," & GetNumMonths(eInstallDataType_EOD) & "|"
        g.strInstalledString = g.strInstalledString & ",S,"
    End If
        
    If chkMutEod.Value = vbChecked Then
        strReturn = strReturn & "M,E," & GetNumMonths(eInstallDataType_EOD) & "|"
        g.strInstalledString = g.strInstalledString & ",M,"
    End If
        
    If chkIdxEod.Value = vbChecked Then
        strReturn = strReturn & "I,E," & GetNumMonths(eInstallDataType_EOD) & "|"
        g.strInstalledString = g.strInstalledString & ",I,"
    End If
    
    If chkFutTick.Value = vbChecked Then
        strReturn = strReturn & "F,T," & GetNumMonths(eInstallDataType_Tick) & "|"
        g.strInstalledString = g.strInstalledString & ",FT,"
    End If
        
    If chkStkTick.Value = vbChecked Then
        strReturn = strReturn & "S,T," & GetNumMonths(eInstallDataType_Tick) & "|"
        g.strInstalledString = g.strInstalledString & ",ST,"
    End If
    
    If chkIdxTick.Value = vbChecked Then
        strReturn = strReturn & "I,T," & GetNumMonths(eInstallDataType_Tick) & "|"
        g.strInstalledString = g.strInstalledString & ",IT,"
    End If
        
    If Right(strReturn, 1) = "|" Then
        strReturn = Mid(strReturn, 1, Len(strReturn) - 1)
    End If
    
    BuildAuthorization = strReturn
        
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmDataInstall.BuildAuthorization", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetNumMonths
'' Description: Get the number of months to pass in the authorization string
''              according to the option that the user chose
'' Inputs:      Type of data (Tick or EOD)
'' Returns:     0 for minimal install, -12 for full install, 12 for one year
''   (full install: EOD = -12 for faster scans, TICK = -2 for faster charts)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetNumMonths(ByVal eType As eInstallDataType) As String
On Error GoTo ErrSection:

    If chkFutTick Or chkStkTick Or chkIdxTick Then
        If eType = eInstallDataType_Tick Or m.bMultipleDisks = False Then
            If optAll.Value = True Then
                If eType = eInstallDataType_Tick Then
                    GetNumMonths = "-2"
                Else
                    GetNumMonths = "-12"
                End If
            ElseIf optMinimal.Value = True Then
                GetNumMonths = "0"
            ElseIf eType = eInstallDataType_Tick Then
                GetNumMonths = "6"
            Else
                GetNumMonths = "12"
            End If
        Else
            GetNumMonths = "-12"
        End If
    Else
        If eType = eInstallDataType_Tick Then
            GetNumMonths = ""
        Else
            If optAll.Value = True Then
                GetNumMonths = "-12"
            ElseIf optMinimal.Value = True Then
                GetNumMonths = "0"
            Else
                GetNumMonths = "12"
            End If
        End If
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmDataInstall.GetNumMonths", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tmrUnload_Timer
'' Description: When the timer is enabled, unload the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tmrUnload_Timer()
On Error GoTo ErrSection:

    ' we can't unload the form until after the Change event for the text
    ' box has finished (otherwise a crash), so we'll do it from here
    tmrUnload.Enabled = False
    Unload Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDataInstall.tmrUnload.Timer", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtHwnd_Change
'' Description: When a new message comes from the install, update the progress
''              bar or set off the finish install stuff
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtHwnd_Change()
On Error GoTo ErrSection:

    Dim iPercent&

    Select Case Val(Parse(txtHwnd.Text, vbTab, 6))
        Case Is = eDMStatus_Process
            iPercent = Val(Parse(txtHwnd.Text, vbTab, 2))
            vsProgress.FloodPercent = iPercent
            vsProgress.Caption = Format(iPercent, "0") & "%"
            vsProgress.Refresh
        Case Is = eDMStatus_Done
            FinishInstall
            tmrUnload.Enabled = True
        Case Is = eDMStatus_Err
            Screen.MousePointer = vbNormal
            InfBox "Error Installing Data", "e", , "Error"
            KillFile DataPath & "*.*"
            g.strInstalledString = ""
            tmrUnload.Enabled = True
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDataInstall.txtHwnd.Change", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FinishInstall
'' Description: Reload the data manager, the universe, the symbol grid, and
''              any graphs that might be visible
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FinishInstall()
On Error GoTo ErrSection:

    Dim strInf$, strStarter$, strTemp$, i&
    Dim bPurchaseOK As Boolean
    Dim fh As Integer
    Dim frmActive As Form

    cmdCancel.Enabled = False
    Sleep 0.5
    
    ' store Starter.GZP info (format date same way InstallShield does)
    strStarter = AddSlash(lblFrom) & "Starter.GZP"
    'strInf = strStarter & vbTab & Str(CDbl(FileDate(strStarter)))
    strInf = strStarter & vbTab & m.strStarterDate
    FileFromString DataPath & "Starter.INF", strInf
    
    DM_Close g.DMS
    DM_Init True
    g.Universe.OpenDb
    
    vsProgress.Caption = "Reloading symbols ..."
    vsProgress.Refresh
    g.SymbolPool.Load False
    
    Me.Top = -Me.Height - 1000
    DoEvents
    Me.Hide
    
    frmSymbolGrid.InitForm
    If (DockState(frmSymbolGrid) = eHidden) And IsExtremeVersion Then
        DockState(frmSymbolGrid) = eShowAsPrevious
    End If
    
    Set frmActive = ActiveChart
    If frmActive Is Nothing Then
        Set frmActive = New frmChart
        frmActive.Chart.SetSymbol g.SymbolPool.SymbolIDforSymbol("$DJIA")
        frmActive.WindowState = 2
        frmActive.Show
    End If
    
    frmSymbolGrid.RefreshGrid '.fg.Refresh
    UpdateVisibleCharts
    frmSymbolGrid.ShowInitialSymbol
    
    frmQuotes.LoadTable
    frmQuotes.TotalRefresh True
    If DockState(frmQuotes) = eHidden And Not IsExtremeVersion Then
        DockState(frmQuotes) = eShowAsPrevious
    End If
        
    Screen.MousePointer = vbNormal
    InfBox "Data installation completed successfully", "i", , "Finished"
    
    ' Put out flag file to tell us we need to ask the user to activate
    FileFromString App.Path & "\Install.flg", m.strInstall
    
    ' If we are installing the "old way", then we need to ask for activation...
    'If Not FileExist(App.Path & "\Provided\Install.CFG") Then
        AskForActivate
    'End If
    
    'NOTE: we CANNOT unload it here or it will crash!
    'Unload Me
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDataInstall.FinishInstall", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FigureSize
'' Description: Figures out how much disk space the user needs to install
''              what they have chosen on the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FigureSize()
On Error GoTo ErrSection:

    Dim lSize As Long                   ' Size required
        
    ' Get the size for the overhead stuff
    lSize = GetSize("*")
    
    ' Add the size for the futures if necessary
    If chkFutEod.Value = vbChecked Then
        lSize = lSize + GetSize("F,E," & GetNumMonths(eInstallDataType_EOD))
    End If
    
    ' Add the size for the indexes if necessary
    If chkIdxEod.Value = vbChecked Then
        lSize = lSize + GetSize("I,E," & GetNumMonths(eInstallDataType_EOD))
    End If
    
    ' Add the size for the stocks if necessary
    If chkStkEod.Value = vbChecked Then
        lSize = lSize + GetSize("S,E," & GetNumMonths(eInstallDataType_EOD))
    End If
    
    ' Add the size for the futures ticks if necessary
    If chkFutTick.Value = vbChecked Then
        lSize = lSize + GetSize("F,T," & GetNumMonths(eInstallDataType_Tick))
    End If
    
    ' Add the size for the index ticks if necessary
    If chkIdxTick.Value = vbChecked Then
        lSize = lSize + GetSize("I,T," & GetNumMonths(eInstallDataType_Tick))
    End If
    
    ' Add the size for the stock ticks if necessary
    If chkStkTick.Value = vbChecked Then
        lSize = lSize + GetSize("S,T," & GetNumMonths(eInstallDataType_Tick))
    End If
    
    ' Add the size for the mutual funds if necesary
    If chkMutEod.Value = vbChecked Then
        lSize = lSize + GetSize("M,E," & GetNumMonths(eInstallDataType_EOD))
    End If
    
    ' Set the space required label
    lblSpaceRequired.Caption = Format(lSize / 1000000, "#,##0") & " MB"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDataInstall.FigureSize", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetSize
'' Description: Gets the file size from the file size array for the given
''              security type/data type/number of months to install
'' Inputs:      Entry to process
'' Returns:     Size for the entry
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetSize(ByVal pstrEntry As String) As Long
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index for a for loop
    
    ' Initialize the size
    GetSize = 0&
    
    ' Walk through the array and try to find the entry passed in
    For lIndex = 0 To m.astrFileSizes.Size - 1
        If InStr(m.astrFileSizes(lIndex), pstrEntry) Then
            GetSize = Val(Parse(m.astrFileSizes(lIndex), vbTab, 2))
            Exit For
        End If
    Next lIndex
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmDataInstall.GetSize", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CheckInstallData
'' Description: Set up the install according to the starter file
'' Inputs:      Starter
'' Returns:
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CheckInstallData(ByVal strStarter$)
On Error GoTo ErrSection:

    Dim strInf As String
    Dim strTemp As String
    Dim dInstallDataDate As Double
    
    If Len(strStarter) > 0 Then
        ' Don't install data older than what was on CD
        strInf = GetCDDataInf
        dInstallDataDate = Val(Parse(strInf, vbTab, 2))
        If UCase(Right(strStarter, 4)) <> ".GZP" Then
            strStarter = AddSlash(strStarter) & "Starter.GZP"
        End If
        Do While FileDate(strStarter) < dInstallDataDate - 1# / 1400#
            strTemp = "Updated CD data could not be found.|Please insert the installation CD from Genesis."
            If AskBox("i=[] ; h=Data CD Setup ; b=+Retry|-Abort ; " & strTemp) = "A" Then
                strStarter = ""
                'If FileExist("k:\common\files.exe") Then
                    'Unload Me
                'End If
                Exit Do
            End If
        Loop
    End If
    
    
    Set m.astrFileSizes = New cGdArray
    m.astrFileSizes.Create eGDARRAY_Strings
    If FileExist(strStarter) Then
        cmdInstall.Tag = strStarter
        lblFrom = AddSlash(FilePath(strStarter))
        m.astrFileSizes.FromFile AddSlash(FilePath(strStarter)) & "Files.TXT"
        fraDataTypes.Enabled = True
        fraAmount.Enabled = True
        FigureSize
    Else
        cmdInstall.Tag = ""
        lblFrom = ""
        lblSpaceRequired = ""
        fraDataTypes.Enabled = False
        fraAmount.Enabled = False
    End If
    
    ReadInfoFile
    
    If Not m.bTick Then
        fraTickData.Visible = False
        fraEodData.Left = fraTickData.Left
        lblInfo.Caption = "Please select amount of data to install."
        If IsExtremeVersion And (chkFutEod.Value = 0) Then
            lblFutures.Visible = False
            chkFutEod.Visible = False
        Else
            lblFutures.Visible = True
            chkFutEod.Visible = True
        End If
    Else
        fraTickData.Visible = True
        fraEodData.Left = 600
        lblInfo.Caption = "Please select amount of End of Day data to install."
        lblFutures.Visible = True
        chkFutEod.Visible = True
    End If
    EnableControls

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmDataInstall.CheckInstallData", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetToPath
'' Description: Get Drive and Space information
'' Inputs:      Path to install to
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetToPath(ByVal strPath$)
On Error Resume Next

    Dim dMB#, drv
    Dim fs As New FileSystemObject

    Set drv = fs.GetDrive(fs.GetDriveName(strPath))
    dMB = drv.AvailableSpace / (1024# * 1024#)
    lblSpaceAvailable.Caption = CStr(Int(dMB)) & " MB"
    lblTo.Caption = AddSlash(strPath)

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ReadInfoFile
'' Description: Read in the information from the info file off of the CD and
''              store the information in the appropriate member variables
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ReadInfoFile()
On Error GoTo ErrSection:

    Dim fh As Integer                   ' File handle to the info file
    Dim strBuffer As String             ' Buffer from the info file
    Dim strUserName As String           ' User name to put into the registry
    Dim strPassword As String           ' Password to put into the registry
    Dim strKey As String                ' Key into the registry
    Dim strTemp As String               ' Temporary string
    
    m.bMultipleDisks = False
    If FileExist(AddSlash(lblFrom.Caption) & "FileInfo.TXT") Then
        fh = FreeFile
        Open AddSlash(lblFrom.Caption) & "FileInfo.TXT" For Input As #fh
        Do While Not EOF(fh)
            Line Input #fh, strBuffer
            Select Case UCase(Parse(strBuffer, "=", 1))
                Case "AVAILABLE":
                    m.strInstall = Parse(strBuffer, "=", 2)
                Case "DEFAULT":
                    m.strDefault = Parse(strBuffer, "=", 2)
                Case "DATASERVICE":
                    m.strDataServID = Parse(strBuffer, "=", 2)
                Case "PASSWORD":
                    m.strPassword = Parse(strBuffer, "=", 2)
                Case "HASTICK":
                    m.bTick = CBool(Parse(strBuffer, "=", 2))
                Case "NUMDATA":
                    m.bMultipleDisks = True
            End Select
        Loop
        Close #fh
        
        ' Save the user name and password if there isn't one there already
        If m.strDataServID <> "" Then
            If RI_GetDataServiceID = 0 Then
                RI_SetDataServiceID Val(m.strDataServID)
            End If
        End If
        If m.strPassword <> "" Then
            If RI_GetUserPassword = "" Then
                RI_SetUserPassword m.strPassword
            End If
        End If
    Else
        m.strInstall = "*F:0-0,*S:0-0,*I:0-0,*SO:0-0,*IO:0-0"
    End If

    m.strInstall = m.strInstall & "," & DM_GetPurchased(False)
    If m.strInstall <> "" Then
        ' Disable appropriate items according to the available string
        chkFutEod.Enabled = InStr(m.strInstall, "*F:")
        chkStkEod.Enabled = InStr(m.strInstall, "*S:")
        chkIdxEod.Enabled = InStr(m.strInstall, "*I:")
        chkMutEod.Enabled = InStr(m.strInstall, "*M:")
        chkFutTick.Enabled = InStr(m.strInstall, "*FT:")
        chkStkTick.Enabled = InStr(m.strInstall, "*ST:")
        chkIdxTick.Enabled = InStr(m.strInstall, "*IT:")
    End If
        
    If m.strDefault <> "" Then
        ' Set up defaults from the defaults string
        strTemp = UCase("," & Parse(m.strDefault, ";", 1) & ",")
        If InStr(strTemp, ",F,") <> 0 Then chkFutEod.Value = vbChecked Else chkFutEod.Value = vbUnchecked
        If InStr(strTemp, ",S,") <> 0 Then chkStkEod.Value = vbChecked Else chkStkEod.Value = vbUnchecked
        If InStr(strTemp, ",I,") <> 0 Then chkIdxEod.Value = vbChecked Else chkIdxEod.Value = vbUnchecked
        If InStr(strTemp, ",M,") <> 0 Then chkMutEod.Value = vbChecked Else chkMutEod.Value = vbUnchecked
        If InStr(strTemp, ",FT,") <> 0 Then chkFutTick.Value = vbChecked Else chkFutTick.Value = vbUnchecked
        If InStr(strTemp, ",ST,") <> 0 Then chkStkTick.Value = vbChecked Else chkStkTick.Value = vbUnchecked
        If InStr(strTemp, ",IT,") <> 0 Then chkIdxTick.Value = vbChecked Else chkIdxTick.Value = vbUnchecked
        Select Case Val(Parse(m.strDefault, ";", 2))
            Case -12:
                optAll.Value = True
            Case 0:
                optMinimal.Value = True
            Case 12, 6:
                optYear.Value = True
        End Select
    End If
      
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDataInstall.ReadInfoFile", eGDRaiseError_Raise
    
End Sub
