VERSION 5.00
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmDownload 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Update Data"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4785
   Icon            =   "frmDownload.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   Begin HexUniControls.ctlUniButtonImageXP cmdConnections 
      Height          =   525
      Left            =   3540
      TabIndex        =   1
      Top             =   2760
      Width           =   1155
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
      Caption         =   "frmDownload.frx":038A
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmDownload.frx":03D0
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmDownload.frx":03F0
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdDataTypes 
      Height          =   435
      Left            =   3540
      TabIndex        =   3
      Top             =   1260
      Width           =   1155
      _ExtentX        =   0
      _ExtentY        =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmDownload.frx":040C
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmDownload.frx":0440
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmDownload.frx":0460
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniCheckXP chkSkipDownload 
      Height          =   435
      Left            =   3540
      TabIndex        =   6
      Top             =   1800
      Visible         =   0   'False
      Width           =   1155
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
      Caption         =   "frmDownload.frx":047C
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Pressed         =   0   'False
      Tip             =   "frmDownload.frx":04B6
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frmDownload.frx":04D6
      ShowFocus       =   -1  'True
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniFrameWL fraData 
      Height          =   3615
      Left            =   120
      TabIndex        =   8
      Top             =   60
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
      Caption         =   "frmDownload.frx":04F2
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmDownload.frx":053E
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmDownload.frx":055E
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniTextBoxXP txtSpecialFile 
         Height          =   285
         Left            =   420
         TabIndex        =   10
         Top             =   3210
         Width           =   2595
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   0   'False
         Locked          =   0   'False
         Text            =   "frmDownload.frx":057A
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
         Tip             =   "frmDownload.frx":05A8
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmDownload.frx":05C8
      End
      Begin HexUniControls.ctlUniRadioXP optQBR 
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   2220
         Width           =   2415
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
         Caption         =   "frmDownload.frx":05E4
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmDownload.frx":062C
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmDownload.frx":064C
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optCurrentSession 
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   2895
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
         Caption         =   "frmDownload.frx":0668
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmDownload.frx":06BE
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmDownload.frx":06DE
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin gdOCX.gdSelectDate dtpToDate 
         Height          =   315
         Left            =   900
         TabIndex        =   4
         Top             =   960
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         AllowWeekends   =   0   'False
         MaxDate         =   42605
         MaxDateIsToday  =   -1  'True
         Value           =   37015
      End
      Begin gdOCX.gdSelectDate dtpFromDate 
         Height          =   315
         Left            =   900
         TabIndex        =   2
         Top             =   600
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         AllowWeekends   =   0   'False
         MaxDate         =   42605
         MaxDateIsToday  =   -1  'True
         Value           =   37015
      End
      Begin HexUniControls.ctlUniRadioXP optDaily 
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   300
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
         Caption         =   "frmDownload.frx":06FA
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   -1  'True
         Tip             =   "frmDownload.frx":0744
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmDownload.frx":0764
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optSpecialFile 
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   2970
         Width           =   2595
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
         Caption         =   "frmDownload.frx":0780
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmDownload.frx":07CE
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmDownload.frx":07EE
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label2 
         Height          =   435
         Left            =   420
         Top             =   2460
         Width           =   2745
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
         Caption         =   "frmDownload.frx":080A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmDownload.frx":08BE
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmDownload.frx":08DE
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label1 
         Height          =   435
         Left            =   420
         Top             =   1680
         Width           =   2595
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
         Caption         =   "frmDownload.frx":08FA
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmDownload.frx":09A2
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmDownload.frx":09C2
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblTo 
         Height          =   255
         Left            =   540
         Top             =   1020
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
         Caption         =   "frmDownload.frx":09DE
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmDownload.frx":0A02
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmDownload.frx":0A22
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblFrom 
         Height          =   255
         Left            =   420
         Top             =   660
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
         Caption         =   "frmDownload.frx":0A3E
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmDownload.frx":0A66
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmDownload.frx":0A86
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
      Cancel          =   -1  'True
      Height          =   435
      Left            =   3540
      TabIndex        =   12
      Top             =   660
      Width           =   1155
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
      Caption         =   "frmDownload.frx":0AA2
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmDownload.frx":0AD0
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmDownload.frx":0AF0
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdDownload 
      Default         =   -1  'True
      Height          =   435
      Left            =   3540
      TabIndex        =   11
      Top             =   120
      Width           =   1155
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
      Caption         =   "frmDownload.frx":0B0C
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmDownload.frx":0B38
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmDownload.frx":0B58
      RightToLeft     =   0   'False
   End
End
Attribute VB_Name = "frmDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmDownload.frm
'' Description: Form to handle the downloading of data
''
'' Author:      Genesis Financial Data Services
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modfication History:
'' Date      Author         Description
'' ??/??/??     R Johnson   Created
'' 12/08/2000   DAJ         Added comments/formatting
'' 04/25/2011   DAJ         Flatten expired simulated positions after a daily download
'' 06/21/2011   DAJ         Separate out Simulated trading types
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Public bDownloadDone As Boolean

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: If the user clicks on the cancel button, unload the form
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
    RaiseError "frmDownload.cmdCancel_Click"
    Resume ErrExit
        
End Sub

Private Sub cmdConnections_Click()
On Error GoTo ErrSection:

    frmHTTP.ShowMe

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConfig.cmdConnections_Click"
    Resume ErrExit
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdDataTypes_Click
'' Description: Allow the user to select data types to filter
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdDataTypes_Click()
On Error GoTo ErrSection:

    frmFilterDownload.ShowMe

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDownload.cmdDataTypes_Click"
    Resume ErrExit
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Download
'' Description: If the user clicks on the download button, download the data
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DownloadData()
On Error GoTo ErrSection:

    Dim strFileName As String           ' File name of the file to download
    Dim nStartDate As Date              ' Starting date of the download
    Dim nEndDate As Date                ' Ending date of the download
    Dim nDate As Date                   ' Index for a for loop
    Dim lIndex As Long                  ' Index for a for loop
    Dim lValid As Long                  ' Return from ZipExecute
    Dim lDate As Long
    Dim State As DM_Status              ' State of the data manager
    Dim strTemp As String               ' Temporary string variable
    Dim bReload As Boolean              ' Do the symbols need to be reloaded?
    Dim bDoFinalUpdate As Boolean
    Dim bRunProc As Boolean             ' Return from Run Process
    Dim SymbolGroup As cSymbolGroup     ' SymbolGroup
    Dim bSuccess As Boolean
    Dim strKey As String
    Dim bFilesDistributed As Boolean
    Dim strFormat As String
    Dim aStrings As New cGdArray
    Dim strAnswer As String
    Dim i&, s$
    Dim bIncludeSnapshot As Boolean
    Dim aRequest As New cGdArray
    Dim aFiles As New cGdArray
    Dim QuoteList As New cSymbolGroup
    Dim strSecType As String
    Dim strSymbol As String
    Dim iExistingDays As Long
    Dim iNewDays As Long
    Dim bDownloadNow As Boolean
    Dim aDates As New cGdArray
    
    ' check for quote board refresh
    If optQBR.Value Then
        Me.Hide
        'g.RealTime.RefreshSymbolList True
        frmQuotes.RefreshQB
        Unload Me
        Exit Sub
    End If
    
    ' don't allow re-entering while in the middle of downloading
    Static bInProgress As Boolean
    If bInProgress Then Exit Sub
    bInProgress = True
    frmStatus.IsBusy = True
   
    'Bring up filter download form if user has never seen it
    If optDaily Then 'And cmdDataTypes.Visible Then
        i = True ' do a hidden save
        If Len(GetIniFileProperty("DownloadExclude", "", "General", g.strIniFile)) < 1 Then
            If Len(GetIniFileProperty("DownloadInclude", "", "General", g.strIniFile)) < 1 Then
                'If ExtremeCharts = 1 Or HasModule("TRANS") Or HasModule("PHOTON") Then
                '    frmFilterDownload.ShowMe True
                'Else
                    'frmFilterDownload.ShowMe
                    i = False ' don't hide it
                'End If
            End If
        End If
        frmFilterDownload.ShowMe i
    End If

    DebugLog "Auto Daily Update Disabled"
    g.dNextDownloadTry = 0

    If dtpToDate < dtpFromDate Then dtpToDate = dtpFromDate
    
    strKey = "Software\Genesis Financial Data Services\Navigator Suite\General"

    ' Bring up the download status form
    Me.Hide
    bDownloadDone = False
    frmStatus.Status = eStatus_Initialized

    If Not optDaily Then
        frmStatus.AddDetail "Building Request"
    End If
    DoEvents
    MoveFocus ActiveChart 'move focus back to chart
        
    ' Make sure all our universe stuff has been flushed
    ' before doing a distribution (separate thread).
    SyncCodebaseCaches
    ''TblFlush
    
    ' clear any read-only flags
    ClearReadOnlyFlags App.Path & "\Info\*.*"
    ClearReadOnlyFlags App.Path & "\ftp\*.*"
    ClearReadOnlyFlags App.Path & "\ftp\backup\*.*"
    ClearReadOnlyFlags App.Path & "\ftp\dist\*.*"
    
    If Dir(App.Path & "\ftp\*.*") <> "" Then
        KillFile AddSlash(App.Path) & "Ftp\*.*", True
    End If
    
    ' reset GenTick export list
    If FileExist(App.Path & "\GenTick.EXP") Then
        FileFromString App.Path & "\Data\TickDist.Lst", " ", True
    End If
    
    bIncludeSnapshot = False
    aDates.Create eGDARRAY_Longs
    Select Case True
    
        Case optDaily.Value
            ' show daily news for Better Trades (any BT user even if Gold or Plat)
            strTemp = GetProvidedProperty("WebNews", , True)
            If Len(strTemp) > 0 And Len(InternetBrowser) > 0 Then
                ' TLB 3/5/2014: but only show it once per day (in this case, we just want to round to the nearest NY midnight)
                nDate = Round(ConvertTimeZone(Now, "", "NY"))
                If nDate > GetIniFileProperty("WebNewsDisplayed", 0, "", g.strIniFile) Then
                    SetIniFileProperty "WebNewsDisplayed", nDate, "", g.strIniFile
                    RunProcess InternetBrowser, Chr(34) & strTemp & Chr(34)
                End If
            End If
        
            ' Get the date range of the dates to download from the user interface
            frmStatus.SetTitle "End-of-day Updating"
            nStartDate = dtpFromDate.Value
            nEndDate = dtpToDate.Value
            If chkSkipDownload = 0 Then
                iNewDays = 0
                iExistingDays = 0
                aRequest.Clear
                For nDate = nStartDate To nEndDate
                    If IsWeekday(nDate) Then
                        aDates.Add nDate
                    
                        ' Write a new header line (%DAILY DOWNLOAD) for each date
                        aRequest.Add "%DAILY DOWNLOAD"
                        
                        ' add exclude/include modules
                        strTemp = Trim(GetIniFileProperty("DownloadExclude", "", "General", g.strIniFile))
                        If Len(strTemp) > 0 Then
                            aRequest.Add "+EXCLUDE:" & strTemp
                        End If
                        strTemp = Trim(GetIniFileProperty("DownloadInclude", "", "General", g.strIniFile))
                        If Len(strTemp) > 0 Then
                            aRequest.Add "+INCLUDE:" & strTemp
                        End If
                        
                        ' flag for if the installed database is the newer full-tick type
                        If IsFullTickDB Then
                            aRequest.Add "+FULLTICK:TRUE"
                        Else
                            aRequest.Add "+FULLTICK:FALSE"
                        End If
                        
                        ' Write the date line
                        aRequest.Add "+DATE:" & ConvertDate(nDate)
                        
                        ' Write information for each file that we have in the backup directory
                        ' for this date
                        aFiles.GetMatchingFiles App.Path & "\ftp\backup\" & ConvertDate(nDate) & "*.GZP"
                        If aFiles.Size > 0 Then
                            For i = 0 To aFiles.Size - 1
                                aRequest.Add "+CRC:" & FileCrcString(aFiles(i))
                            Next
                            iExistingDays = iExistingDays + 1
                        Else
                            iNewDays = iNewDays + 1
                        End If
                    
                        ' see if aborting
                        If frmStatus.Status >= eStatus_Aborting And frmStatus.Status <= eStatus_Error Then
                            Exit For
                        End If
                        
                        ' See if should request the data so-far:
                        ' - if last day we need to do it now
                        If nDate = nEndDate Then
                            bDownloadNow = True
                        ' - but if only one more day after this, we'll wait for the last day
                        ElseIf nDate = nEndDate - 1 Or (nDate Mod 7 = 6 And nEndDate - nDate <= 3) Then
                            bDownloadNow = False
                        ' - else see if enough days for this chunk
                        ElseIf iNewDays >= 2 Or iNewDays + iExistingDays >= 5 Then
                            bDownloadNow = True
                        Else ' - otherwise wait
                            bDownloadNow = False
                        End If
                        If bDownloadNow Then
                            ' Request and download data
                            If nDate <> nEndDate Then
                                '(show details if will be more than one download)
                                frmStatus.ShowDetails True
                            End If
                            frmStatus.AddDetail "Requesting data through " & DateFormat(nDate, MM_DD_YY)
                            bSuccess = FtpRequest(aRequest, , , True) ' skip symbol reload since will be done later
                            ' Backup daily download file(s)
                            For i = 0 To aDates.Size - 1
                                strFileName = App.Path & "\ftp\" & ConvertDate(aDates(i)) & "*.GZP"
                                If FileExist(strFileName) Then
                                    FileCopy strFileName, AddSlash(App.Path) & "Ftp\Backup\", True
                                    KillFile strFileName
                                End If
                            Next
                            ' reset things
                            iNewDays = 0
                            iExistingDays = 0
                            aRequest.Clear
                            aDates.Clear
                            If frmStatus.Status <> eStatus_Completed Then
                                Exit For
                            End If
                        End If
                    End If
                Next nDate
            End If
          
        Case optCurrentSession.Value
            frmStatus.SetTitle "Current Session Daily Bars"
            If ScansEnabled Then
                strAnswer = AskBox("h=Criteria ; b=+Yes|-No ; i=? ; Do you want to recalculate criteria and filters based on the current session update data?")
                If strAnswer = "Y" Then
                    bReload = True
                    bIncludeSnapshot = True
                End If
            End If
            frmStatus.Status = eStatus_Completed '(to skip download)
            bSuccess = True
    
        Case optSpecialFile.Value
            frmStatus.SetTitle "Downloading Special File"
            strFileName = FixSpecialFilename
            If UCase(strFileName) = "UPGRADE.GZP" Then
                ''strFilename = "UPGRD" & Format(App.Major * 100 + App.Minor, "000") & ".GZP"
                strFileName = "UPGRD" & Str(App.Major) & Str(App.Minor) & ".GZP"
                txtSpecialFile = strFileName
            ElseIf UCase(strFileName) = "BETA.GZP" Then
                strFileName = "BETA" & Str(App.Major) & Str(App.Minor) & ".GZP"
                txtSpecialFile = strFileName
            End If
            ' see if specifying a local path
            If InStr(strFileName, "\") > 0 Then
                If FileExist(strFileName) Then
                    FileCopy strFileName, App.Path & "\ftp\", True
                End If
                strFileName = FileBase(strFileName) & "." & FileExt(strFileName)
                If FileExist(App.Path & "\ftp\" & strFileName) Then
                    txtSpecialFile = strFileName
                    frmStatus.AddDetail txtSpecialFile & " copied"
                    frmStatus.Status = eStatus_Completed '(to skip download)
                    bSuccess = True
                Else
                    frmStatus.AddDetail txtSpecialFile & " does not exist"
                    frmStatus.Status = eStatus_Error '(to abort entire process)
                End If
            ElseIf UCase(strFileName) = "TEST.GZP" Or UCase(strFileName) = "GETCODES.GZP" Then
                KillFile App.Path & "\ftp\backup\" & strFileName
            End If
            aRequest.Add "<" & FileCrcString(App.Path & "\ftp\backup\" & strFileName)
            
    End Select
    
    ' Request the data
    If frmStatus.Status < eStatus_Aborting Then
        If chkSkipDownload Then
            bSuccess = True
            frmStatus.Status = eStatus_Completed
        ElseIf Not optDaily And Not optCurrentSession Then
            bSuccess = FtpRequest(aRequest)
            aRequest.Clear
        End If
    End If

'If optDaily Then frmStatus.Status = eStatus_Aborted

    If frmStatus.Status = eStatus_Aborted Then
        '(nothing more to do)
    ElseIf frmStatus.Status = eStatus_Error Or bSuccess = False Then
        frmStatus.AddDetail "ERROR downloading data"
    ElseIf frmStatus.Status = eStatus_Completed Then
        bDownloadDone = True
        
        ' Begin distribution
        Select Case True
            ' Do special file download for the end of day files
            Case optDaily.Value
                ' Unzip files from the backup folder into the distribution folder
                For nDate = nStartDate To nEndDate
                    
                    If IsWeekday(nDate) Then
          
                        ' Clean out the distribution directory
                        If Dir(App.Path & "\ftp\dist\*.*") <> "" Then
                            KillFile AddSlash(App.Path) & "Ftp\Dist\*.*", True
                        End If
          
                        ' Unzip any gzp files for this date that are in the ftp\backup\ directory
                        aFiles.GetMatchingFiles App.Path & "\ftp\backup\" & ConvertDate(nDate) & "*.GZP"
                        If aFiles.Size > 0 Then
                            For i = 0 To aFiles.Size - 1
                                ' Unzip the file into the distribution directory
                                lValid = ZipExecute("U", aFiles(i), App.Path & "\ftp\dist", "", False, False)
                                
                                ' Process special download files
                                CheckForSpecialDownloadFiles App.Path & "\ftp\dist"
                            Next
                        
                            ' distribute them
                            frmStatus.AddDetail "Starting Distribution for " & nDate
                            If DM_DistribData(App.Path & "\ftp\dist\") = True Then
                                strKey = "Software\Genesis Financial Data Services\Navigator Suite\General"
                                SetRegistryValue rkLocalMachine, strKey, "SessionUpdate", False
                            
                                bFilesDistributed = True
                                'if successfully distributed, save date to file (if newer than existing)
                                lDate = JulToLong(nDate, True)
                                strTemp = DataPath & "LastDown.txt"
                                If lDate > Val(FileToString(strTemp, 250)) Then
                                    FileFromString strTemp, Trim(Str(lDate)), True, False
                                End If
                                
                                ' delete COT report if distributed COT data
                                ' (so will force it to recalculate the report)
                                If FileExist(App.Path & "\ftp\dist\??????CT.*") Then
                                    KillFile App.Path & "\CotData.GRD"
                                End If
                            End If
                            bReload = True
                        Else
                            frmStatus.AddDetail "Files do not exist for " & nDate
                            If nDate <> nEndDate Or Not bFilesDistributed Then
                                frmStatus.Status = eStatus_Error
                            End If
                        End If
          
                        If frmStatus.Status = eStatus_Aborted Or frmStatus.Status = eStatus_Error Then
                            Exit For
                        End If
                    End If
                Next nDate
            
            Case optCurrentSession.Value
                If g.RealTime.RefreshSymbolList(2, False, True) Then
                    strKey = "Software\Genesis Financial Data Services\Navigator Suite\General"
                    SetRegistryValue rkLocalMachine, strKey, "SessionUpdate", True
                    bDoFinalUpdate = True
                Else
                    bReload = False
                    bDoFinalUpdate = False
                End If
                
            Case optSpecialFile.Value
        
                frmStatus.Status = eStatus_Running
                bRunProc = False
                strFileName = FixSpecialFilename
                If FileExist(AddSlash(App.Path) & "Ftp\" & strFileName) Then ' Or FileExist(AddSlash(App.Path) & "Ftp\Backup\" & Trim(txtSpecialFile.Text)) Then
                    FileCopy AddSlash(App.Path) & "Ftp\" & strFileName, AddSlash(App.Path) & "Ftp\Backup\", True
                    lValid = ZipExecute("U", App.Path & "\ftp\" & strFileName, App.Path & "\ftp", "", False, False)
                ElseIf FileExist(AddSlash(App.Path) & "Ftp\Backup\" & strFileName) Then
                    lValid = ZipExecute("U", App.Path & "\ftp\backup\" & strFileName, App.Path & "\ftp", "", False, False)
                Else
                    frmStatus.AddDetail strFileName & " not found"
                    frmStatus.Status = eStatus_Error
                End If
                           
                If frmStatus.Status <> eStatus_Error Then
                    ' If found run upgrade
                    If FileExist(AddSlash(App.Path) & "Ftp\Upgrd32.exe") Then
                        ' Backup the upgrade file
                        FileCopy AddSlash(App.Path) & "Ftp\Upgrd32.exe", AddSlash(App.Path) & "Ftp\Backup\Upgrd32.exe", True
                        
                        ' TLB 1/29/2014: see if this has a newer (or older!) build of TradeNav
                        s = FileVersion(AddSlash(App.Path) & "Ftp\NavSuite.exe")
                        i = Val(Parse(s, ".", 4))
                        If i > 0 And i < App.Revision Then
                            InfBox "The version of the program you are currently on is newer than what was just downloaded.", "e", , "Upgrade Error"
                        Else
                            If i > App.Revision Then
                                s = "Do you wish to upgrade now?||From: version " & Str(App.Major) & "." & Str(App.Minor) & " build " & Str(App.Revision) _
                                    & "| To:  version " & Parse(s, ".", 1) & "." & Parse(s, ".", 2) & " build " & Parse(s, ".", 4) & "|"
                            Else
                                s = "Do you wish to upgrade now?"
                            End If
                            If AskBox("i=? ; h=Program Upgrade ; b=+Yes|-No ; " & s) = "Y" Then
                                g.bSkipMdbCompact = True '(skip compacting since Upgrade process may try to kill this process if not finished in time)
                                g.strRunWhenExit = Chr(34) & App.Path & "\ftp\upgrd32.exe" & Chr(34) & " " & Chr(34) & App.Path & Chr(34) & "'" & Chr(34) & App.EXEName & Chr(34)
                                frmStatus.Status = eStatus_Completed
                                frmStatus.IsBusy = False
                                bInProgress = False
                                Screen.MousePointer = vbHourglass
                                frmMain.tmrMain.Tag = "QUIT"
                                ''Unload frmMain
                                Exit Sub
                            End If
                        End If
                    End If
                      
                    ' Process special download files
                    CheckForSpecialDownloadFiles App.Path & "\ftp"
                    frmStatus.Status = eStatus_Completed
                      
                    ' If found a ctl file do distribution
                    If FileExist(AddSlash(App.Path) & "Ftp\" & ReplaceFileExt(txtSpecialFile.Text, "ctl")) Then
                        ' Unzip files into the distribution directory
                        frmStatus.Status = eStatus_Running
                    
                        ' Clean out distribution directory
                        If FileExist(AddSlash(App.Path) & "Ftp\Dist\*.*") = True Then
                            KillFile AddSlash(App.Path) & "Ftp\Dist\*.*", True
                        End If
                    
                        If FileExist(AddSlash(App.Path) & "Ftp\" & ReplaceFileExt(txtSpecialFile.Text, "ctl")) Then
                            ' Call Karl's Data Manager
                            DM_DistribData App.Path & "\ftp\"
                              
                            '?? do we need to reload, or just do final updating?
                            bReload = True
                        End If
                    End If
                End If
        End Select

        If frmStatus.Status = eStatus_Completed Then
            If bReload Then bDoFinalUpdate = True
            If Not bDoFinalUpdate Then
                ' Just force visible charts to update
                UpdateVisibleCharts
            Else
                If Not optCurrentSession Then
                    frmStatus.AddDetail "Final Updating"
                    DM_DistribData ""
                End If
                               
                If bReload = True And frmStatus.Status <> eStatus_Aborted Then
                    ' reload symbols
                    frmStatus.UpdateProgress "Reloading Symbols ..."
                    frmStatus.AddDetail "Reloading Symbols"
                    frmStatus.Status = eStatus_Running
'DebugLog Str(ActiveChart.SymbolID) & " A"
                    g.SymbolPool.Load False '(will also set criteria dirty)
                    DoEvents '(to allow aborting to show up)
                    If frmStatus.Status = eStatus_Aborting Then
                        frmStatus.Status = eStatus_Aborted
                    End If
                    frmSymbolGrid.RefreshGrid
                
                    If optDaily Then
                        ' make this call just clear calcs for recent data (in case fixes came in daily download)
                        g.FractZen.GetFractZenRange ""
                    End If
                End If
                
                ' Update old GenTick files (can do this before recalc criteria)
                UpdateGenTick
                
                ' if real-time is active during a daily download, refresh the real-time data
                ' (since symbols could have rolled, etc.)
                If Not optCurrentSession Then
                    If optDaily And g.RealTime.Active Then
                        g.RealTime.RefreshSymbolList True, True
                    Else
                        ' otherwise, just reload the data on all the forms
                        frmStatus.AddDetail "Reloading Data"
                        frmStatus.UpdateProgress "Reloading Data"
                        g.RealTime.RefreshAllFormData True
                    End If
                    frmStatus.IsBusy = True
                End If
                    
                ' After a daily download, show any new rolls that may have occurred...
                If optDaily And HasModule("F") And frmStatus.Status <> eStatus_Aborted Then
                    frmRollsTable.ShowMe True
                End If
                
                ' After a daily download, expire any parked or trigger pending orders...
                ' DAJ 04/25/2011: ...and also expire any positions that are no longer valid...
                If optDaily Then
                    ExpireNonSubmittedOrders
                    g.Broker.FlattenExpiredPositions
                    
                    ' TLB 11/7/2011: and show warning message if close to expiration of data subscription
                    ExpiringDataPkgWarning False
                    
                    ' TLB 11/18/2013: special Sector Tree export (primarily for Greg S.)
                    If FileExist(App.Path & "\SectorExport.flg") Then
                        frmSectorTree.ShowMe "", True
                    End If
                End If
                
                ' Recalculate the criteria
                If bReload = True And frmStatus.Status <> eStatus_Aborted Then
                    g.SymbolPool.RecalcDirtyCriteria bIncludeSnapshot
                End If
            
                ' Export data (must do this AFTER criteria recalc so filtered groups will be updated)
                If optDaily = True And frmStatus.Status <> eStatus_Aborted _
                        And (HasGold(False, , False) Or HasModule("VPT,DTP,EXP")) Then
                    ExportData
                End If
                
                ' If running as ImageServer, then shut down after daily download to let it restart (due to memory leak)
                If optDaily And FileLength(App.Path & "\ImageServer.flg") > 10 Then
                    g.bSkipMdbCompact = True '(skip compacting since Upgrade process may try to kill this process if not finished in time)
                    ''g.strRunWhenExit = Chr(34) & App.Path & "\ftp\upgrd32.exe" & Chr(34) & " " & Chr(34) & App.Path & Chr(34) & "'" & Chr(34) & App.EXEName & Chr(34)
                    frmStatus.Status = eStatus_Completed
                    frmStatus.IsBusy = False
                    bInProgress = False
                    Screen.MousePointer = vbHourglass
                    frmMain.tmrMain.Tag = "QUIT"
                    ''Unload frmMain
                    Exit Sub
                End If
            End If
            
            If frmStatus.Status <> eStatus_Aborted And frmStatus.Status <> eStatus_Error Then
                frmStatus.Status = eStatus_Completed
                frmStatus.AddDetail "Finished"
                
                ' if daily downloads are "caught up"
                If optDaily And Not NeedDailyUpdate Then
                    SetupBrokerLayout
                    If HasModule("RTG,RTE") And Not g.RealTime.Active Then
                        StatusMsg "To start realtime streaming, click the traffic light on the toolbar.", -1
                    End If
                    DoPFCheck
                    
                    ' TLB 2/25/2013: save some things now (in case they had changed)
                    If FormIsLoaded("frmQuotes") Then
                        frmQuotes.SaveAllQbSettings
                    End If
                    SaveCharts
                    SaveVisibleForms
                    SaveChartGlobals
                End If
            End If
        End If
    End If
        
    CalcNextTryTime True
    
    ' Clean up after ourselves
    '2/7/01: NO LONGER -- in case a batch file is still running from this path
    ''If Dir(App.Path & "\ftp\*.*") <> "" Then KillFile App.Path & "\ftp\*.*", True
       
ErrExit:
    frmStatus.IsBusy = False
    bInProgress = False
    Unload Me
    Exit Sub
    
ErrSection:
    If frmStatus.Status = eStatus_Running Then frmStatus.Status = eStatus_Aborted
    frmStatus.IsBusy = False
    bInProgress = False
    RaiseError "frmDownload.DownloadData"
    Resume ErrExit
        
End Sub

Private Sub cmdDownload_Click()
On Error GoTo ErrSection:

    ' Need to do this for "Default" buttons, since hitting 'Enter'
    ' from another control does not trigger it's 'LostFocus' event.
    MoveFocus cmdDownload
    DoEvents

    DownloadData

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDownload.cmdDownload_Click"
    Resume ErrExit
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Activate
'' Description: When the form is activated, reset the toolbar
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Activate()
On Error GoTo ErrSection:

    EnableButtons

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDownload.Form_Activate"
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
    RaiseError "frmDownload.Form_KeyDown"
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: When the form is loaded, do some initialization
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()

    Dim wday%, d As Date
    Dim lLastDownload As Long           ' Last download for a symbol

    'can't afford to have an error in this routine,
    'else user couldn't even download a fix!
    On Error Resume Next

    CenterTheForm Me
    
    g.Styler.StyleForm Me
    
    Me.Icon = Picture16(ToolbarIcon("ID_Download"), , True)

    lLastDownload = LastDailyDownload
    If lLastDownload <= 0 Then
        ' If no data, don't bother doing a daily download!
        optDaily.Enabled = False
        optCurrentSession.Enabled = False
        optSpecialFile = True
    Else
        optDaily.Enabled = True
        optCurrentSession.Enabled = True
        With dtpToDate
            .AllowWeekends = False
            .MaxDateIsToday = True
            .MinDate = lLastDownload
            .Value = Int(ConvertTimeZone(Now, "", "NY") - 0.5)
        End With
        With dtpFromDate
            .AllowWeekends = False
            ' set max to next weekday after last download date
            Select Case Weekday(lLastDownload)
                Case vbFriday
                    .MaxDate = lLastDownload + 3
                Case vbSaturday
                    .MaxDate = lLastDownload + 2
                Case Else
                    .MaxDate = lLastDownload + 1
            End Select
            If .MaxDate > Date Then .MaxDate = Date
            .Value = .MaxDate
        End With
    End If
    
    dtpToDate.Enabled = False
       
    If ExtremeCharts = 1 And Not HasModule("F") And Not HasModule("IT") And Not HasModule("ST") Then
        cmdDataTypes.Visible = True 'False
    Else
        cmdDataTypes.Visible = True
    End If
       
       
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Resize
'' Description: When the form gets resized, call refresh to make the painting
''              faster
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Resize()
On Error Resume Next

    If Me.Visible Then
        Me.Refresh 'to paint fast
    End If
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    lblTo_Click
'' Description: If the user clicks on the To label, toggle the ToDate control
''              and the visibility of the Skip Download check box
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub lblTo_Click()
On Error GoTo ErrSection:

    dtpToDate.Enabled = Not dtpToDate.Enabled
    chkSkipDownload.Visible = dtpToDate.Enabled
    cmdDataTypes.Visible = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDownload.lblTo_Click"
    Resume ErrExit
        
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optCurrentSession_Click
'' Description: If they try to download a Current Session Update after 4:30
''              more than one business day since they have done a daily
''              download, try to get them to do a daily download first
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optCurrentSession_Click()
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return from an InfBox
    Dim dNextMarketClosed As Double                 ' Temporary variable

    EnableButtons

    If Not Me.Visible Then Exit Sub
       
    ' If the current time in New York is after 7:00pm more than one business
    ' day since the last daily download, ask if they would like to do a Daily
    ' Download before doing a Current Session Update
    dNextMarketClosed = LastDailyDownload + 19# / 24#
    Do
        dNextMarketClosed = dNextMarketClosed + 1#
    Loop While Not IsWeekday(dNextMarketClosed)
    
    ' See if current time in NY is past market close
    If ConvertTimeZone(Now) > dNextMarketClosed Then
        strReturn = InfBox("We recommend that you do a End-of-Day Update before doing a Current Session Update.||Would you like to do this now?|", "?", "+Yes|-No", "Warning")
        If strReturn = "Y" Then
            optDaily.Value = True
            DownloadData
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDownload.optCurrentSession_Click"
    Resume ErrExit
        
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optDaily_Click
'' Description: If the daily option is clicked, change the focus to an
''              appropriate control
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optDaily_Click()
On Error GoTo ErrSection:

    EnableButtons
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDownload.optDaily_Click"
    Resume ErrExit
        
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optSpecialFile_Click
'' Description: If the Special File option is clicked, change the focus to an
''              appropriate control
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optSpecialFile_Click()
On Error GoTo ErrSection:

    EnableButtons
    SelectAll txtSpecialFile
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDownload.optSpecialFile_Click"
    Resume ErrExit
        
End Sub

Private Sub txtSpecialFile_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtSpecialFile

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDownload.txtSpecialFile_GotFocus"
    Resume ErrExit
        
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtSpecialFile_LostFocus
'' Description: If the Special File text box loses focus, append a ".gzp" to
''              the end of the filename if there is no decimal point there.
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtSpecialFile_LostFocus()
On Error GoTo ErrSection:

    FixSpecialFilename
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDownload.txtSpecialFile_LostFocus"
    Resume ErrExit
        
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EnableButtons
'' Description: Enable/Disable the appropriate controls under certain conditions
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EnableButtons()
On Error GoTo ErrSection:

    If optDaily And Not optDaily.Enabled Then
        optSpecialFile = True
    End If
    
    If optDaily Then
        Enable dtpFromDate
        MoveFocus dtpFromDate
        Disable txtSpecialFile
        txtSpecialFile.BackColor = fraData.BackColor
    ElseIf Me.optCurrentSession Then
        Disable txtSpecialFile
        txtSpecialFile.BackColor = fraData.BackColor
        Disable dtpFromDate
        Disable dtpToDate
    Else 'special file
        txtSpecialFile.BackColor = &H80000005
        Enable txtSpecialFile
        MoveFocus txtSpecialFile
        txtSpecialFile.SelStart = 0
        txtSpecialFile.SelLength = Len(txtSpecialFile.Text)
        Disable dtpFromDate
        Disable dtpToDate
    End If
    
If Not FileExist(g.strAppPath & "\Provided\DownloadFilter.cfg") Then
    cmdDataTypes.Visible = False
Else
    cmdDataTypes.Enabled = optDaily
End If

    optQBR.Enabled = Not g.RealTime.SalmonIsRunning

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDownload.EnableButtons"
        
End Sub

Private Function FixSpecialFilename() As String
On Error GoTo ErrSection:

    Dim strFileName$
    
    ' fix name of file (strip out whitespace, quotes, etc. -- e.g. things
    ' that might get cut-and-paste from an email)
    strFileName = StripStr(txtSpecialFile.Text, " '`<>" & Chr(34) & vbTab)
    If Len(strFileName) > 0 Then
        If InStr(Right(strFileName, 4), ".") = 0 Then
            strFileName = strFileName & ".GZP"
        End If
    End If
    If strFileName <> txtSpecialFile.Text Then
        txtSpecialFile.Text = strFileName
    End If
    FixSpecialFilename = strFileName

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmDownload.FixSpecialFilename"
End Function

