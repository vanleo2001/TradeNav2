VERSION 5.00
Object = "{3B008041-905A-11D1-B4AE-444553540000}#1.0#0"; "Vsocx6.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.Ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmMain 
   Caption         =   "Archive"
   ClientHeight    =   3705
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5010
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   5010
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4380
      Top             =   3060
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin HexUniControls.ctlUniFrameWL fraPath 
      Height          =   705
      Left            =   120
      TabIndex        =   6
      Top             =   2220
      Width           =   4755
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
      Caption         =   "frmMain.frx":0442
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmMain.frx":0496
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmMain.frx":04B6
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdBrowse 
         Height          =   330
         Left            =   120
         TabIndex        =   7
         Top             =   270
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
         Caption         =   "frmMain.frx":04D2
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmMain.frx":0502
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmMain.frx":0522
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniTextBoxXP txtPath 
         Height          =   285
         Left            =   1020
         TabIndex        =   8
         Top             =   300
         Width           =   3615
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmMain.frx":053E
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
         Tip             =   "frmMain.frx":055E
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmMain.frx":057E
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraAction 
      Height          =   1935
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   4755
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
      Caption         =   "frmMain.frx":059A
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmMain.frx":05CE
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmMain.frx":05EE
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniRadioXP optUndoRestore 
         Height          =   220
         Left            =   2640
         TabIndex        =   4
         Top             =   300
         Width           =   1695
         _ExtentX        =   2990
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
         Caption         =   "frmMain.frx":060A
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmMain.frx":0644
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmMain.frx":0664
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optBackup 
         Height          =   220
         Left            =   240
         TabIndex        =   2
         Top             =   300
         Width           =   1035
         _ExtentX        =   1826
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
         Caption         =   "frmMain.frx":0680
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmMain.frx":06AE
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmMain.frx":06CE
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optRestore 
         Height          =   220
         Left            =   1440
         TabIndex        =   3
         Top             =   300
         Width           =   1095
         _ExtentX        =   1931
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
         Caption         =   "frmMain.frx":06EA
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmMain.frx":071A
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmMain.frx":073A
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin RichTextLib.RichTextBox rtfHelp 
         Height          =   1215
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   4515
         _ExtentX        =   7964
         _ExtentY        =   2143
         _Version        =   393217
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"frmMain.frx":0756
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   615
      Left            =   780
      TabIndex        =   10
      Top             =   3000
      Width           =   3375
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
      Caption         =   "frmMain.frx":07D8
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmMain.frx":07F8
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmMain.frx":0818
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdExit 
         Height          =   495
         Left            =   2280
         TabIndex        =   1
         Top             =   60
         Width           =   915
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
         Caption         =   "frmMain.frx":0834
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmMain.frx":085E
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmMain.frx":087E
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Height          =   495
         Left            =   180
         TabIndex        =   0
         Top             =   60
         Width           =   1755
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
         Caption         =   "frmMain.frx":089A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmMain.frx":08D4
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmMain.frx":08F4
         RightToLeft     =   0   'False
      End
   End
   Begin vsOcx6LibCtl.vsElastic vseMessage 
      Height          =   360
      Left            =   120
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3180
      Visible         =   0   'False
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   635
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
      BackColor       =   7368816
      ForeColor       =   54000
      FloodColor      =   5177367
      ForeColorDisabled=   -2147483631
      Caption         =   "Message ..."
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
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmMain.frm
'' Description: Main form for the Trade Navigator Archive project
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 02/13/2009   DAJ         Created
'' 06/11/2009   DAJ         If user path doesn't exist, prompt to create
'' 08/19/2009   DAJ         Changed command line, name backup file with date,
''                          allow user to choose restore file
'' 09/02/2009   DAJ         Changed default file name to include century
'' 09/09/2009   DAJ         On Restore, Error on Customer ID, Warn on Data Service
'' 09/11/2009   DAJ         Took out "AND False" when doing a restore
'' 06/29/2010   DAJ         Don't do customer ID check if Genesis employee
'' 10/21/2014   DAJ         Replaced usage of the old Browse Folders form
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    strSettingsFile As String           ' Settings information INI File
    strRegistryFile As String           ' File of registry information
    
    RegistryItems As cRegistryItems     ' Collection of registry items
    FileItems As cFileItems             ' Collection of file items
    
    bInProcess As Boolean               ' In process of backup or restore?
End Type
Private m As mPrivate

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdBrowse_Click
'' Description: Allow the user to browse for a path
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdBrowse_Click()
On Error GoTo ErrSection:

    Dim strPath As String               ' Path returned from dialog

    If optRestore.Value = True Then
        strPath = mGenesis.CommonDialogFile(Me.CommonDialog1, False, "Archive Files (*.TNA)|*.TNA|All Files|*.*", Trim(txtPath.Text), "Select Archive File")
    Else
        strPath = mGenesis.BrowseForFolder(Trim(txtPath.Text), "Select Path for Archive File")
    End If
    
    If Len(strPath) > 0 Then
        txtPath.Text = Trim(strPath)
    End If
    
    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMain.cmdBrowse_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdExit_Click
'' Description: Allow the user to exit the application
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdExit_Click()
On Error GoTo ErrSection:

    Unload Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMain.cmdExit_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOK_Click
'' Description: Save settings and start the backup or restore process
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
On Error GoTo ErrSection:

    If KillProcess("NavSuite", True) = 0 Then
        SaveSettings
        
        If optBackup.Value = True Then
            Backup
        Else
            Restore
        End If
        
        If InfBox("Would you like to restart " & g.CommandLine.Caption & " now?", "?", "+Yes|-No", "Restart " & g.CommandLine.Caption) = "Y" Then
            RunProcess AddSlash(App.Path) & "NavSuite.EXE"
            Unload Me
        Else
            EnableControls
            Unload Me
        End If
    Else
        Select Case True
            Case optBackup
                ShutdownMsg "before doing a backup", "Backup Error"
            Case optRestore
                ShutdownMsg "before doing a restore", "Restore Error"
            Case optUndoRestore
                ShutdownMsg "before doing an undo restore", "Undo Restore Error"
        End Select
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMain.cmdOK_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Activate
'' Description: The first time the form is activated, check for automatic action
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Activate()
On Error GoTo ErrSection:

    Static bAlreadyHere As Boolean      ' Have we already been through here?
    
    If bAlreadyHere = False Then
        bAlreadyHere = True
        
        Select Case g.CommandLine.Action
            Case 0
                Disable optBackup
                Disable optRestore
                Disable optUndoRestore
                
                Backup g.CommandLine.ArchiveFile
                Unload Me
                
            Case 1
                Disable optBackup
                Disable optRestore
                Disable optUndoRestore
                
                Restore AddSlash(g.CommandLine.ArchivePath) & g.CommandLine.ArchiveFile
                Unload Me
        
        End Select
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMain.Form_Activate"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Set up the form when it is initially loaded
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:
    
    Dim strPlacement As String          ' Placement of the form

    m.bInProcess = False

    ' Set the caption for the main window...
    Caption = g.CommandLine.Caption & " Archive"
    
    'RH
    g.Styler.StyleForm Me
    
    
    ' Get and apply the placement for the main window...
    strPlacement = GetIniFileProperty("Main", "", "Placement", g.strIniFile)
    If Len(strPlacement) = 0 Then
        Move Left, Top, 8200, 5685
        CenterTheForm Me
    Else
        SetFormPlacement Me, strPlacement
    End If
    
    ' Load the settings from the INI file...
    LoadSettings
    
    ' Clear out the message control...
    UpdateMessage 0, ""
    
    m.strSettingsFile = AddSlash(App.Path) & kArchiveSettings
    m.strRegistryFile = AddSlash(App.Path) & kRegistryFile
    
    Set m.RegistryItems = New cRegistryItems
    Set m.FileItems = New cFileItems
    
    EnableControls
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMain.Form_Load"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: Cancel the unload if in process
'' Inputs:      Cancel Unload?, Mode of the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If (m.bInProcess = True) And (UnloadMode <> vbFormCode) Then
        Cancel = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMain.Form_QueryUnload"
    
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

    Dim nLeft&
    
    If Not LimitFormSize(Me, 4800, 3600) Then
        nLeft = fraAction.Left
        
        With fraButtons
            .Move (Me.ScaleWidth - .Width) / 2, Me.ScaleHeight - .Height - 60
            vseMessage.Move nLeft, .Top + (.Height - vseMessage.Height) / 2, Me.ScaleWidth - nLeft * 2
        End With
        
        With fraPath
            .Move fraAction.Left, fraButtons.Top - .Height - 120, Me.ScaleWidth - fraAction.Left * 2
            txtPath.Width = .Width - txtPath.Left - nLeft
        End With
        
        With fraAction
            .Move nLeft, .Top, Me.ScaleWidth - nLeft * 2, fraPath.Top - .Top * 2
            rtfHelp.Move nLeft, rtfHelp.Top, .Width - nLeft * 2, .Height - rtfHelp.Top - nLeft
        End With
    End If
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: Save settings and clean up when the form is unloaded
'' Inputs:      Cancel the Unload?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    Set m.RegistryItems = Nothing
    Set m.FileItems = Nothing

    SetIniFileProperty "Main", GetFormPlacement(Me), "Placement", g.strIniFile

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMain.Form_Unload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optBackup_Click
'' Description: Enable/Disable controls when the user changes this one
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optBackup_Click()
On Error GoTo ErrSection:

    Dim strMessage As String            ' Message to display in RTF box

    SetPathText
    EnableControls
    
    strMessage = FileToString(AddSlash(App.Path) & kBackupInfo)
    rtfHelp.TextRTF = strMessage
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMain.optBackup_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optRestore_Click
'' Description: Enable/Disable controls when the user changes this one
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optRestore_Click()
On Error GoTo ErrSection:

    Dim strMessage As String            ' Message to display in RTF box

    SetPathText
    EnableControls
    
    strMessage = FileToString(AddSlash(App.Path) & kRestoreInfo)
    rtfHelp.TextRTF = strMessage
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMain.optRestore_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optUndoRestore_Click
'' Description: Enable/Disable controls when the user changes this one
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optUndoRestore_Click()
On Error GoTo ErrSection:

    Dim strMessage As String            ' Message to display in RTF box

    SetPathText
    EnableControls
    
    strMessage = FileToString(AddSlash(App.Path) & kUndoRestoreInfo)
    rtfHelp.TextRTF = strMessage
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMain.optUndoRestore_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtPath_Change
'' Description: As the user changes the path, enable/disable controls
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtPath_Change()
On Error GoTo ErrSection:

    If Visible Then
        EnableControls
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMain.txtPath_Change"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateMessage
'' Description: Update the elastic control as appropriate
'' Inputs:      Percent Done, Message, Log Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UpdateMessage(ByVal lPercent As Long, Optional ByVal strMessage As String = "", Optional ByVal strLogMessage As String = "")
On Error GoTo ErrSection:

    Static strPrevMessage As String     ' Previous message
    
    With vseMessage
        If fraButtons.Visible Then
            fraButtons.Visible = False
            .Visible = True
        End If
        
        If lPercent >= 0 Then
            .FloodPercent = lPercent
        End If
        
        If Len(strMessage) = 0 Then
            strMessage = strPrevMessage
        Else
            strPrevMessage = strMessage
        End If
        
        .Caption = strMessage
        
        DoEvents
    End With
    
    If Len(strLogMessage) > 0 Then
        LogToFile strLogMessage
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMain.UpdateMessage"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SaveSettings
'' Description: Save the user interface settings
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveSettings()
On Error GoTo ErrSection:

    If optBackup.Value = True Then
        SetIniFileProperty "Action", 0&, "Main", g.strIniFile
    ElseIf optRestore.Value = True Then
        SetIniFileProperty "Action", 1&, "Main", g.strIniFile
    End If
    
    If optBackup.Value = True Then
        SetIniFileProperty "Path", txtPath.Text, "Main", g.strIniFile
    Else
        SetIniFileProperty "Path", FilePath(txtPath.Text), "Main", g.strIniFile
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMain.SaveSettings"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadSettings
'' Description: Load the user interface settings
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadSettings()
On Error GoTo ErrSection:

    Dim lAction As Long                 ' Initial value for action
    
    ' Set the action according the the last one selected...
    lAction = GetIniFileProperty("Action", 0&, "Main", g.strIniFile)
    If lAction = 0& Then
        optBackup.Value = True
        optRestore.Value = False
    Else
        optBackup.Value = False
        optRestore.Value = True
    End If
    
    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMain.LoadSettings"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Backup
'' Description: Backup files as per the settings
'' Inputs:      Archive File
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Backup(Optional ByVal strArchiveFile As String = "")
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lCount As Long                  ' Number of objects
    Dim strPath As String               ' Path for the archived files
    Dim InfoFile As cInfoFile           ' Information file object
    Dim strBackupFile As String         ' Name of the backup file
    Dim strUserFile As String           ' Name of the backup file with user's path
    Dim strPrevFile As String           ' Name of the previous backup file
    Dim strNewFile As String            ' Name of the new backup file
    Dim bContinue As Boolean            ' Do we want to continue the backup?
    Dim strLastDir As String            ' Last directory in the path

    If KillProcess("NavSuite", True) = 0 Then
        m.bInProcess = True
        EnableControls
        
        bContinue = True
        strPath = FilePath(txtPath.Text)
        If DirExist(strPath) = 0 Then
            bContinue = False
            strLastDir = LastDirInPath(strPath)
            If InfBox("The directory '" & strLastDir & "' does not exist.  Would you like to create it?", "?", "+Yes|-No", "Confirmation") = "Y" Then
                If MakeDir(strPath) = False Then
                    InfBox "Error creating directory", "!", , "Backup Error"
                Else
                    bContinue = True
                End If
            End If
        End If
        
        If bContinue Then
            m.FileItems.FromIniFile m.strSettingsFile
            lCount = m.FileItems.Count + 2
            
            If Len(strArchiveFile) = 0 Then
                strArchiveFile = Format(Now, "YYYYMMDD_HHMMSS") & ".TNA"
            End If
            
            If optBackup.Value = True Then
                strBackupFile = AddSlash(App.Path) & "Archive\" & strArchiveFile
                strPrevFile = Replace(AddSlash(App.Path) & "Archive\" & strArchiveFile, ".TNA", ".OLD")
                strNewFile = Replace(AddSlash(App.Path) & "Archive\" & strArchiveFile, ".TNA", ".NEW")
                strUserFile = strPath & strArchiveFile
            Else
                strBackupFile = strPath & kUndoRestoreFile
                strPrevFile = Replace(strBackupFile, ".TNA", ".OLD")
                strNewFile = Replace(strBackupFile, ".TNA", ".NEW")
            End If
            
            ' Backup the information out of the registry...
            UpdateMessage 0&, "Saving Registry Information ...", "Retrieving and saving registry information"
            m.RegistryItems.FromIniFile m.strSettingsFile
            m.RegistryItems.ToRegFile m.strRegistryFile
            
            ' Walk through each of the file groups and back them up...
            For lIndex = 1 To m.FileItems.Count
                UpdateMessage (lIndex / lCount) * 100, "Backing up " & m.FileItems(lIndex).GroupName & " ...", m.FileItems(lIndex).LogMessage(True)
                m.FileItems(lIndex).ZipUpFiles AddSlash(App.Path) & "Archive"
            Next lIndex
            
            ' Backup the existing archive file (if exists) and create new one...
            UpdateMessage ((lCount - 1) / lCount) * 100, "Finishing Backup ...", "Preparing " & strUserFile
            KillFile strPrevFile
            RenameFile strBackupFile, strPrevFile
            ZipExecute "Z", strNewFile, AddSlash(App.Path), "Archive\*.GZP"
            KillFile AddSlash(App.Path) & "Archive\*.GZP"
            RenameFile strNewFile, strBackupFile
            
            ' Create and add the info file to the big zip file...
            Set InfoFile = New cInfoFile
            InfoFile.GetInformation
            InfoFile.ToFile AddSlash(App.Path) & "Archive\Archive.INF"
            ZipExecute "Z", strBackupFile, AddSlash(App.Path), "Archive\Archive.INF"
            KillFile AddSlash(App.Path) & "Archive\Archive.INF"
            
            ' If the user specified a different path than our default, copy the GZP over there...
            If optBackup.Value = True Then
                If UCase(strBackupFile) <> UCase(strUserFile) Then
                    FileCopy strBackupFile, strUserFile, True
                End If
            End If
            
            If optBackup.Value = True Then
                SetIniFileProperty "LastBackupFile", strArchiveFile, "Main", g.strIniFile
            End If
            
            UpdateMessage 100&, "Files Backed Up", "Done Backing Up"
        End If
    Else
        ShutdownMsg " before creating an archive", "Backup Error"
    End If
    
    m.bInProcess = False
    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    m.bInProcess = False
    RaiseError "frmMain.Backup"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Restore
'' Description: Restore files as per the settings
'' Inputs:      Archive File
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Restore(Optional ByVal strArchiveFile As String = "")
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lCount As Long                  ' Number of objects
    Dim InfoFile As cInfoFile           ' Information file object
    Dim ThisInfo As cInfoFile           ' Information for this machine
    Dim strRestoreFile As String        ' Restore file
    Dim strUserFile As String           ' User selected file to restore
    Dim strReturn As String             ' Return from an InfBox

    If optRestore Then
        If Len(strArchiveFile) = 0 Then
            strUserFile = Trim(txtPath.Text)
        Else
            strUserFile = strArchiveFile
        End If
    Else
        strUserFile = Trim(txtPath.Text)
    End If
    
    strRestoreFile = AddSlash(App.Path) & "Archive\" & FileBase(strUserFile) & "." & FileExt(strUserFile)
    
    If KillProcess("NavSuite", True) = 0 Then
        ' First do a backup in case something goes wrong...
        If optRestore Then
            Backup
        End If
    
        If FileExist(strUserFile) Then
            m.bInProcess = True
            EnableControls
            Set InfoFile = New cInfoFile
            
            ' If the selected file was in a different path, copy it over first...
            If UCase(strRestoreFile) <> UCase(strUserFile) Then
                FileCopy strUserFile, strRestoreFile, True
            End If
            
            ZipExecute "U", strRestoreFile, AddSlash(App.Path) & "Archive", "Arch*.INF"
            If InfoFile.FromFile(AddSlash(App.Path) & "Archive\Archive.INF") Then
                Set ThisInfo = New cInfoFile
                ThisInfo.GetInformation
                
                If ThisInfo.TradeNavBuild < InfoFile.TradeNavBuild Then
                    InfBox "Error: This archive file is too new for this version of " & g.CommandLine.Caption, "!", , "Restore Error"
                    LogToFile "Error: This archive file is too new for this version of " & g.CommandLine.Caption
                ElseIf (ThisInfo.CustomerNumber <> InfoFile.CustomerNumber) And (Not FileExist("C:\Common\Files.EXE")) Then
                    InfBox "Error: the Customer ID is not the same as the archive file", "!", , "Restore Error"
                    LogToFile "Error: the Customer ID is not the same as the archive file"
                ElseIf ThisInfo.ArchiveVersion < InfoFile.ArchiveVersion Then
                    InfBox "Error: This archive file is too new for this version of the Archive program", "!", , "Restore Error"
                    LogToFile "Error: This archive file is too new for this version of the Archive program"
                Else
                    strReturn = "Y"
                    If ThisInfo.DataService <> InfoFile.DataService Then
                        strReturn = InfBox("You have chosen to restore a file that was backed up under a different data service ID.||Do you want to continue?|", "?", "+Yes|-No", "Restore Confirmation")
                    End If
                    
                    If strReturn = "Y" Then
                        m.FileItems.FromIniFile m.strSettingsFile
                        lCount = m.FileItems.Count + 2
                        
                        UpdateMessage 0&, "Unzipping Archive File ...", "Unzipping Archive.GZP to " & AddSlash(App.Path) & "Archive"
                        ZipExecute "U", strRestoreFile, AddSlash(App.Path) & "Archive"
                        
                        For lIndex = 1 To m.FileItems.Count
                            UpdateMessage (lIndex / lCount) * 100, "Restoring " & m.FileItems(lIndex).GroupName & " ...", m.FileItems(lIndex).LogMessage(False)
                            m.FileItems(lIndex).ExtractFiles AddSlash(App.Path) & "Archive"
                        Next lIndex
                        
                        UpdateMessage ((lCount - 1) / lCount) * 100, "Restoring Registry Information ...", "Restoring registry information"
                        m.RegistryItems.FromRegFile m.strRegistryFile
                        m.RegistryItems.ToRegistry
                        
                        ' Delete the SymPool.MEM so that Trade Navigator will regenerate it...
                        KillFile AddSlash(App.Path) & "SymPool.MEM"
                        
                        UpdateMessage 100&, "Files Restored", "Done Restoring"
                    Else
                        LogToFile "User chose not to restore because data service was different"
                    End If
                End If
                
                KillFile AddSlash(App.Path) & "Archive\Archive.INF"
            Else
                InfBox "Error: Archive information file does not exist", "!", , "Restore Error"
                LogToFile "Error: Archive information file does not exist"
            End If
        Else
            InfBox "Error: There is no archive file to restore", "!", , "Restore Error"
            LogToFile "Error: There is no archive file to restore"
        End If
    Else
        ShutdownMsg " before restoring an archive", "Restore Error"
    End If

    m.bInProcess = False
    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    m.bInProcess = False
    RaiseError "frmMain.Restore"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EnableControls
'' Description: Enable/Disable controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EnableControls()
On Error GoTo ErrSection:

    Dim bOneOptionChosen As Boolean     ' Is one of the options chosen?
    Dim strButton$
    
    If optUndoRestore Then
        strButton = "Undo Restore"
    ElseIf optRestore.Value Then
        strButton = "&Start Restore"
    Else
        strButton = "&Start Backup"
    End If
    If cmdOK.Caption <> strButton Then cmdOK.Caption = strButton

    bOneOptionChosen = ((optBackup.Value = True) Or (optRestore.Value = True) Or (optUndoRestore.Value = True))

    Enable optUndoRestore, FileDate(AddSlash(App.Path) & "Archive\" & kUndoRestoreFile) > FileDate(AddSlash(App.Path) & "Archive\" & kArchiveFile)
    
    Enable cmdOK, (bOneOptionChosen And (Len(Trim(txtPath.Text)) > 0)) And (m.bInProcess = False)
    Enable cmdExit, Not m.bInProcess
    Enable cmdBrowse, (Not optUndoRestore.Value) And (Not m.bInProcess)
    Enable txtPath, (Not optUndoRestore.Value)
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMain.EnableControls"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetPathText
'' Description: Set the path text based on the option selected
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetPathText()
On Error GoTo ErrSection:

    Dim strPath As String               ' Path to store in the text box
    Dim strLastBackup As String         ' Last backup file

    ' Set the default path...
    strPath = g.CommandLine.ArchivePath
    
    If optRestore.Value = True Then
        strLastBackup = GetIniFileProperty("LastBackupFile", "", "Main", g.strIniFile)
        If Len(strLastBackup) > 0 Then
            strPath = AddSlash(strPath) & strLastBackup
        Else
            strPath = AddSlash(strPath) & kArchiveFile
        End If
    ElseIf optUndoRestore.Value = True Then
        strPath = AddSlash(App.Path) & "Archive\" & kUndoRestoreFile
    End If
    
    txtPath.Text = strPath
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMain.SetPathText"

End Sub

