VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmIbDataPull 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Interactive Brokers Data Pull"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniTextBoxXP txtOutputDirectory 
      Height          =   285
      Left            =   1500
      TabIndex        =   7
      Top             =   1500
      Width           =   4755
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmIbDataPull.frx":0000
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
      Tip             =   "frmIbDataPull.frx":0020
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmIbDataPull.frx":0040
   End
   Begin MSComctlLib.ProgressBar pbStatus 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   3600
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   495
      Left            =   1920
      TabIndex        =   2
      Top             =   2520
      Width           =   2535
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
      Caption         =   "frmIbDataPull.frx":005C
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmIbDataPull.frx":0088
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmIbDataPull.frx":00A8
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdStop 
         Height          =   495
         Left            =   1320
         TabIndex        =   6
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
         Caption         =   "frmIbDataPull.frx":00C4
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmIbDataPull.frx":00F8
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmIbDataPull.frx":0118
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdStart 
         Height          =   495
         Left            =   0
         TabIndex        =   10
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
         Caption         =   "frmIbDataPull.frx":0134
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmIbDataPull.frx":016A
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmIbDataPull.frx":018A
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniTextBoxXP txtOutputFile 
      Height          =   285
      Left            =   1500
      TabIndex        =   3
      Top             =   840
      Width           =   4755
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmIbDataPull.frx":01A6
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
      Tip             =   "frmIbDataPull.frx":01C6
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmIbDataPull.frx":01E6
   End
   Begin HexUniControls.ctlUniTextBoxXP txtSymbolList 
      Height          =   285
      Left            =   1500
      TabIndex        =   1
      Top             =   210
      Width           =   4755
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmIbDataPull.frx":0202
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
      Tip             =   "frmIbDataPull.frx":0222
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmIbDataPull.frx":0242
   End
   Begin HexUniControls.ctlUniFrameWL fraIntraday 
      Height          =   315
      Left            =   240
      TabIndex        =   8
      Top             =   1980
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
      Caption         =   "frmIbDataPull.frx":025E
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmIbDataPull.frx":028A
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmIbDataPull.frx":02AA
      RightToLeft     =   0   'False
      Begin gdOCX.gdSelectDate gdStartDate 
         Height          =   315
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   556
         ShowTime        =   2
      End
      Begin gdOCX.gdSelectDate gdEndDate 
         Height          =   315
         Left            =   3180
         TabIndex        =   11
         Top             =   0
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   556
         ShowTime        =   2
      End
      Begin HexUniControls.ctlUniLabelXP lblTo 
         Height          =   195
         Left            =   2880
         Top             =   60
         Width           =   195
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
         Caption         =   "frmIbDataPull.frx":02C6
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmIbDataPull.frx":02EA
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmIbDataPull.frx":030A
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniRadioXP optIntraday 
      Height          =   220
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   397
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmIbDataPull.frx":0326
      Enabled         =   -1  'True
      Align           =   0
      RadioBackColor  =   -2147483643
      RadioForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Pressed         =   0   'False
      Tip             =   "frmIbDataPull.frx":0358
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frmIbDataPull.frx":0378
      ShowFocus       =   -1  'True
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniRadioXP optDaily 
      Height          =   220
      Left            =   240
      TabIndex        =   4
      Top             =   660
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   397
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmIbDataPull.frx":0394
      Enabled         =   -1  'True
      Align           =   0
      RadioBackColor  =   -2147483643
      RadioForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Pressed         =   0   'False
      Tip             =   "frmIbDataPull.frx":03C0
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frmIbDataPull.frx":03E0
      ShowFocus       =   -1  'True
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblOutputDirectory 
      Height          =   195
      Left            =   240
      Top             =   1530
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
      Caption         =   "frmIbDataPull.frx":03FC
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmIbDataPull.frx":0440
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmIbDataPull.frx":0460
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblStatus 
      Height          =   255
      Left            =   240
      Top             =   3300
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
      Caption         =   "frmIbDataPull.frx":047C
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmIbDataPull.frx":049C
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmIbDataPull.frx":04BC
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblOutputFile 
      Height          =   195
      Left            =   240
      Top             =   870
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
      Caption         =   "frmIbDataPull.frx":04D8
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmIbDataPull.frx":0512
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmIbDataPull.frx":0532
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblSymbolList 
      Height          =   195
      Left            =   240
      Top             =   240
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
      Caption         =   "frmIbDataPull.frx":054E
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmIbDataPull.frx":0588
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmIbDataPull.frx":05A8
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
End
Attribute VB_Name = "frmIbDataPull"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmIbDataPull.frm
'' Description: Allows pulling data from the Traders Workstation software
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 05/20/2014   DAJ         Extension on intraday output file; Fix for persisting EndDate
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    astrRequests As cGdArray            ' List of data requests to make to the server
    lRequestIndex As Long               ' Current index into the request list
    bStopRequest As Boolean             ' Stop the request?
    astrData As cGdArray                ' Array of data information
    strRequestSymbol As String          ' Symbol being requested
End Type
Private m As mPrivate

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Setup and show the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowMe()
On Error GoTo ErrSection:

    LoadControls
    ShowForm Me, eForm_Modal, frmMain

ErrExit:
    Unload Me
    Exit Sub
    
ErrSection:
    Unload Me
    RaiseError "frmIbDataPull.ShowMe"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IB_Data
'' Description: Data message coming back from Interactive Brokers
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub IB_Data(ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim astrIbData As cGdArray
    Dim astrOutData As cGdArray
    
    Set astrIbData = New cGdArray
    astrIbData.SplitFields strMessage, vbTab
    
    If UCase(astrIbData(1)) = "END" Then
        m.lRequestIndex = m.lRequestIndex + 1&
        If (m.lRequestIndex < m.astrRequests.Size) And (m.bStopRequest = False) Then
            pbStatus.Value = m.lRequestIndex
            pbStatus.Refresh
            
            Sleep 10
            
            SendNextRequest
        Else
            If optIntraday.Value = True Then
                lblStatus.Caption = "Building output file for " & m.strRequestSymbol
            Else
                lblStatus.Caption = "Building output file"
            End If
            lblStatus.Refresh
            
            OutputData
            
            lblStatus.Caption = "Done"
            lblStatus.Refresh
            
            EnableControls False
        End If
    Else
        Set astrOutData = New cGdArray
        astrOutData.Create eGDARRAY_Strings, 10
        
        astrOutData(0) = m.strRequestSymbol
        astrOutData(1) = astrIbData(1)
        astrOutData(2) = astrIbData(2)
        astrOutData(3) = astrIbData(3)
        astrOutData(4) = astrIbData(4)
        astrOutData(5) = astrIbData(5)
        
        If IsForex(m.strRequestSymbol) Then
            astrOutData(6) = "0"
        Else
            astrOutData(6) = astrIbData(6)
        End If
        
        astrOutData(7) = "0"
        astrOutData(8) = "0"
        astrOutData(9) = "0"
        
        m.astrData.Add astrOutData.JoinFields(vbTab)
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmIbDataPull.IB_Data", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdStart_Click
'' Description: Start downloading the data
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdStart_Click()
On Error GoTo ErrSection:

    If Validate Then
        SaveControls
        
        BuildRequests
        m.astrData.Clear
        
        If m.astrRequests.Size > 0 Then
            EnableControls True
            m.bStopRequest = False
            
            pbStatus.Min = 0
            pbStatus.Max = m.astrRequests.Size - 1
            pbStatus.Value = 0
                        
            m.lRequestIndex = 0&
            SendNextRequest
        End If
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmIbDataPull.cmdStart_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdStop_Click
'' Description: Stop downloading the data
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdStop_Click()
On Error GoTo ErrSection:

    m.bStopRequest = True
    cmdStop.Enabled = False

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmIbDataPull.cmdStop_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Setup the form when it is loaded
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Set m.astrRequests = New cGdArray
    m.astrRequests.Create eGDARRAY_Strings
    
    Set m.astrData = New cGdArray
    m.astrData.Create eGDARRAY_Strings
    
    optDaily.Value = True

    g.Styler.StyleForm Me
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmIbDataPull.Form_Load"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Validate
'' Description: Verify the controls and the broker
'' Inputs:      None
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function Validate() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    
    If g.IntBroker Is Nothing Then
        InfBox "Interactive Brokers object is not initialized", "!", , "Error"
    ElseIf g.Broker.ConnectionStatusForBroker(eTT_AccountType_IntBrokers) <> eGDConnectionStatus_Connected Then
        InfBox "Not connected to Interactive Brokers", "!", , "Error"
    ElseIf Len(Trim(txtSymbolList.Text)) = 0 Then
        InfBox "Please enter in a symbol list file", "!", , "Error"
        MoveFocus txtSymbolList
    ElseIf Not FileExist(Trim(txtSymbolList.Text)) Then
        InfBox "List file does not exist.  Please enter in a valid list file", "!", , "Error"
        MoveFocus txtSymbolList
    ElseIf (optDaily.Value = True) And (Len(Trim(txtOutputFile.Text)) = 0) Then
        InfBox "Please enter in an output file", "!", , "Error"
        MoveFocus txtOutputFile
    ElseIf (optIntraday.Value = True) And (Len(Trim(txtOutputDirectory.Text)) = 0) Then
        InfBox "Please enter in an output directory", "!", , "Error"
        MoveFocus txtOutputDirectory
    ElseIf (optIntraday.Value = True) And (Not DirExist(Trim(txtOutputDirectory.Text))) Then
        InfBox "Output directory does not exist.  Please enter in a valid output directory", "!", , "Error"
        MoveFocus txtOutputDirectory
    Else
        bReturn = True
    End If
    
    Validate = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmIbDataPull.Validate"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BuildRequests
'' Description: Build the request list from the controls on the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub BuildRequests()
On Error GoTo ErrSection:

    Dim astrSymbolList As cGdArray      ' List of symbols for which to request data
    Dim lIndex As Long                  ' Index into a for loop
    Dim dStartDate As Double            ' Starting date for the request
    Dim dEndDate As Double              ' Ending date for the request
    Dim strDuration As String           ' Duration
    Dim strWhatToShow As String         ' What to show
    Dim lNumDays As Long                ' Number of days
    Dim dDate As Double                 ' Date/Time for the request

    m.astrRequests.Clear
    
    Set astrSymbolList = New cGdArray
    astrSymbolList.FromFile Trim(txtSymbolList.Text)
    
    dStartDate = gdStartDate.Value
    dEndDate = gdEndDate.Value
    
    ' Symbol, End Date, Duration, Bar Size, What to Show
    If optDaily.Value = True Then
        lNumDays = CLng(dEndDate - dStartDate)
        Select Case lNumDays
            Case Is > 180
                strDuration = "1 Y"
            Case Is > 90
                strDuration = "6 M"
            Case Is > 30
                strDuration = "3 M"
            Case Is > 5
                strDuration = "1 M"
            Case Is > 2
                strDuration = "1 W"
            Case Is > 1
                strDuration = "2 D"
            Case Else
                strDuration = "1 D"
        End Select
        
        For lIndex = 0 To astrSymbolList.Size - 1
            If IsForex(astrSymbolList(lIndex)) Then
                strWhatToShow = "BID"
            Else
                strWhatToShow = "TRADES"
            End If
            
            m.astrRequests.Add astrSymbolList(lIndex) & vbTab & Format(dEndDate, "yyyyMMdd hh:mm:ss") & vbTab & strDuration & vbTab & "1 day" & vbTab & strWhatToShow
        Next lIndex
    ElseIf optIntraday.Value = True Then
        For lIndex = 0 To astrSymbolList.Size - 1
            If IsForex(astrSymbolList(lIndex)) Then
                strWhatToShow = "BID"
            Else
                strWhatToShow = "TRADES"
            End If
            strDuration = "1800 S"
            
            dDate = dStartDate
            Do While dDate <= dEndDate
                m.astrRequests.Add astrSymbolList(lIndex) & vbTab & Format(dDate, "yyyyMMdd hh:mm:ss") & vbTab & strDuration & vbTab & "1 secs" & vbTab & strWhatToShow
                dDate = RoundToMinute(dDate + (30# / 1440#))
            Loop
        Next lIndex
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmIbDataPull.BuildRequests"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SendNextRequest
'' Description: Send the next request to Interactive Brokers
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SendNextRequest()
On Error GoTo ErrSection:

    Dim astrRequest As cGdArray         ' Array of request arguments
    
    If (m.lRequestIndex >= 0) And (m.lRequestIndex < m.astrRequests.Size) Then
        Set astrRequest = New cGdArray
        astrRequest.SplitFields m.astrRequests(m.lRequestIndex), vbTab
        
        If Not g.IntBroker Is Nothing Then
            If m.strRequestSymbol <> astrRequest(0) Then
                If (optIntraday.Value = True) And (Len(m.strRequestSymbol) > 0) Then
                    lblStatus.Caption = "Building output file for " & m.strRequestSymbol
                    lblStatus.Refresh
                    
                    OutputData
                    m.astrData.Clear
                End If
                
                m.strRequestSymbol = astrRequest(0)
            End If
            
            lblStatus.Caption = "Requesting: " & astrRequest.JoinFields(",")
            lblStatus.Refresh
            
            g.IntBroker.RequestHistory astrRequest(0), astrRequest(1), astrRequest(2), astrRequest(3), astrRequest(4)
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmIbDataPull.SendNextRequest"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OutputData
'' Description: Output the data to the output file
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub OutputData()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim astrOutputFile As cGdArray      ' Output file
    Dim astrDataLine As cGdArray        ' Data line split out into fields
    Dim astrOutputLine As cGdArray      ' Output line split out into fields
    Dim strSymbol As String             ' Symbol

    Set astrOutputFile = New cGdArray
    astrOutputFile.Create eGDARRAY_Strings

    If optDaily.Value = True Then
        strSymbol = ""
        
        For lIndex = 0 To m.astrData.Size - 1
            Set astrDataLine = New cGdArray
            astrDataLine.SplitFields m.astrData(lIndex), vbTab
            
            If astrDataLine(0) <> strSymbol Then
                strSymbol = astrDataLine(0)
                astrOutputFile.Add strSymbol
            End If
            
            Set astrOutputLine = astrDataLine.MakeCopy
            astrOutputLine.Remove 0
            
            astrOutputFile.Add astrOutputLine.JoinFields(vbTab)
        Next lIndex
        
        astrOutputFile.ToFile Trim(txtOutputFile.Text)
    ElseIf optIntraday.Value = True Then
        m.astrData.ToFile AddSlash(Trim(txtOutputDirectory.Text)) & StripStr(m.strRequestSymbol, "$") & ".BAR"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmIbDataPull.OutputData"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EnableControls
'' Description: Enable controls based on if we are running or not
'' Inputs:      Running?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EnableControls(ByVal bRunning As Boolean)
On Error GoTo ErrSection:

    Enable lblSymbolList, Not bRunning
    Enable txtSymbolList, Not bRunning
    Enable optDaily, Not bRunning
    Enable lblOutputFile, Not bRunning
    Enable txtOutputFile, Not bRunning
    Enable optIntraday, Not bRunning
    Enable lblOutputDirectory, Not bRunning
    Enable txtOutputDirectory, Not bRunning
    Enable gdStartDate, Not bRunning
    Enable gdEndDate, Not bRunning
    Enable cmdStart, Not bRunning
    Enable cmdStop, bRunning

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmIbDataPull.EnableControls"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadControls
'' Description: Load control values from the last known set
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadControls()
On Error GoTo ErrSection:

    txtSymbolList.Text = GetIniFileProperty("SymbolList", "", "IbDataPull", g.strIniFile)
    optDaily.Value = GetIniFileProperty("Daily", True, "IbDataPull", g.strIniFile)
    txtOutputFile.Text = GetIniFileProperty("OutputFile", "", "IbDataPull", g.strIniFile)
    optIntraday.Value = GetIniFileProperty("Intraday", False, "IbDataPull", g.strIniFile)
    txtOutputDirectory.Text = GetIniFileProperty("OutputDirectory", "", "IbDataPull", g.strIniFile)
    gdStartDate.Value = GetIniFileProperty("StartDate", Now, "IbDataPull", g.strIniFile)
    gdEndDate.Value = GetIniFileProperty("EndDate", Now, "IbDataPull", g.strIniFile)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmIbDataPull.LoadControls"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SaveControls
'' Description: Save the control values
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveControls()
On Error GoTo ErrSection:

    SetIniFileProperty "SymbolList", txtSymbolList.Text, "IbDataPull", g.strIniFile
    SetIniFileProperty "Daily", optDaily.Value, "IbDataPull", g.strIniFile
    SetIniFileProperty "OutputFile", txtOutputFile.Text, "IbDataPull", g.strIniFile
    SetIniFileProperty "Intraday", optIntraday.Value, "IbDataPull", g.strIniFile
    SetIniFileProperty "OutputDirectory", txtOutputDirectory.Text, "IbDataPull", g.strIniFile
    SetIniFileProperty "StartDate", gdStartDate.Value, "IbDataPull", g.strIniFile
    SetIniFileProperty "EndDate", gdEndDate.Value, "IbDataPull", g.strIniFile

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmIbDataPull.SaveControls"
    
End Sub


