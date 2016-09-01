VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmLibraryImport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import a Library"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9075
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   9075
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7905
      Top             =   3735
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   435
      Left            =   7740
      TabIndex        =   18
      Top             =   4305
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   4155
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   8835
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   8100
         TabIndex        =   6
         Top             =   1575
         Width           =   360
      End
      Begin VB.TextBox txtPath 
         Height          =   285
         Left            =   2415
         TabIndex        =   5
         Top             =   1560
         Width           =   5640
      End
      Begin VB.TextBox Version 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3675
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   2505
         Width           =   2775
      End
      Begin VB.TextBox LastModified 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3675
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   2880
         Width           =   2775
      End
      Begin VB.TextBox Author 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3675
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   3255
         Width           =   4095
      End
      Begin VB.TextBox LibName 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3675
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   2100
         Width           =   4095
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   2595
         Left            =   120
         Picture         =   "frmLibraryImport.frx":0000
         ScaleHeight     =   2595
         ScaleWidth      =   1875
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   180
         Width           =   1875
      End
      Begin VB.Label Label5 
         Caption         =   "Select the file name of the library to Import"
         Height          =   195
         Index           =   1
         Left            =   2415
         TabIndex        =   4
         Top             =   1350
         Width           =   4140
      End
      Begin VB.Label Label3 
         Caption         =   "Version"
         Height          =   195
         Index           =   1
         Left            =   2355
         TabIndex        =   9
         Top             =   2565
         Width           =   1035
      End
      Begin VB.Label Label3 
         Caption         =   "Last Modified"
         Height          =   195
         Index           =   3
         Left            =   2355
         TabIndex        =   11
         Top             =   2940
         Width           =   1035
      End
      Begin VB.Label Label3 
         Caption         =   "Author"
         Height          =   195
         Index           =   4
         Left            =   2370
         TabIndex        =   13
         Top             =   3285
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Library Name"
         Height          =   195
         Index           =   0
         Left            =   2355
         TabIndex        =   7
         Top             =   2145
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Import a Library"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   2160
         TabIndex        =   2
         Top             =   300
         Width           =   3615
      End
      Begin VB.Label Label1 
         Caption         =   $"frmLibraryImport.frx":7F7E
         Height          =   450
         Left            =   2400
         TabIndex        =   3
         Top             =   720
         Width           =   6255
      End
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "Import"
      Default         =   -1  'True
      Height          =   435
      Left            =   6540
      TabIndex        =   17
      Top             =   4305
      Width           =   1095
   End
   Begin VSFlex7LCtl.VSFlexGrid vsStatus 
      Height          =   315
      Left            =   120
      TabIndex        =   15
      Top             =   4380
      Visible         =   0   'False
      Width           =   1515
      _cx             =   2672
      _cy             =   556
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
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
      MousePointer    =   0
      BackColor       =   8421504
      ForeColor       =   16777215
      BackColorFixed  =   8421504
      ForeColorFixed  =   -2147483630
      BackColorSel    =   8421504
      ForeColorSel    =   16777215
      BackColorBkg    =   -2147483636
      BackColorAlternate=   8421504
      GridColor       =   -2147483639
      GridColorFixed  =   -2147483639
      TreeColor       =   -2147483632
      FloodColor      =   16711680
      SheetBorder     =   -2147483642
      FocusRect       =   0
      HighLight       =   0
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   0
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   0
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   8421504
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Label lblStatusMsg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1740
      TabIndex        =   16
      Top             =   4440
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   4620
   End
End
Attribute VB_Name = "frmLibraryImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmLibraryImport.frm
'' Description: Allow the user to import a library
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 10/20/2014   DAJ         Replaced frmBrowseFolders
'' 10/21/2014   DAJ         Fixed filter for file lookup dialog
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Option Compare Text

Private Type mPrivate
    Library As cLibrary                 ' Library object
    bOK As Boolean
End Type
Private m As mPrivate

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Initialize and show the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe() As Boolean
On Error GoTo ErrSection:
    
    MoveFocus txtPath
    ShowForm Me, eForm_Modal, g.frmOwner
    ShowMe = m.bOK

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmLibraryImport.ShowMe", , g.strAppPath

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
    
    m.bOK = False
    Me.Hide

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmLibraryImport.cmdCancel_Click", , g.strAppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdImport_Click
'' Description: Start the importing
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdImport_Click()
On Error GoTo ErrSection:
    
    ' Anytime that a button is set as the default button, we need to make sure to
    ' move the focus to the button and perform a DoEvents so that any appropriate
    ' LostFocus events occur first...
    MoveFocus cmdImport
    DoEvents
    
    Screen.MousePointer = vbHourglass
    If Not m.Library Is Nothing Then
        With m.Library
            .StatusBar = vsStatus
            .Path = txtPath.Text
            .StatusMsg = lblStatusMsg
            .Import
        End With
    Else
        Err.Raise vbObjectError + 1000, , "Library not Initialized"
    End If
    
    cmdImport.Enabled = False
    m.bOK = True
    Me.Hide
    
ErrExit:
    Screen.MousePointer = vbDefault
    Exit Sub

ErrSection:
    Screen.MousePointer = vbDefault
    RaiseError "frmLibraryImport.cmdImport_Click", , g.strAppPath
    If FileExist(AddSlash(g.strAppPath) & "AutoImport.CFG") Then
        m.bOK = False
        Me.Hide
    End If
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdBrowse_Click
'' Description: Allow the user to browse for a library to import
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdBrowse_Click()
On Error GoTo ErrSection:

    Dim strReturn As String             ' Path returned from the browser
    Dim strInitialPath As String        ' Initial path for the search
    
    If Len(txtPath.Text) = 0 Then
        strInitialPath = AddSlash(g.strAppPath) & "Libraries"
    Else
        strInitialPath = txtPath.Text
    End If
        
    strReturn = mGenesis.CommonDialogFile(Me.CommonDialog1, False, "Library Files (*.GLB)|*.GLB|All Files|*.*", strInitialPath, "Select Library(s) to Import", , ",")
    If Len(strReturn) > 0 Then
        txtPath.Text = strReturn
        SetLibraryInfo
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmLibraryImport.cmdBrowse_Click", , g.strAppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_KeyDown
'' Description: If the user hits the F1 key, show the help
'' Inputs:      Code of key pressed, Shift/Ctrl/Alt Status
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
    RaiseError "frmLibraryImport.Form_KeyDown", , g.strAppPath
    
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
    RaiseError "frmLibraryImport.Form_Load", , g.strAppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: Unload the form if the user presses the 'X'
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode <> vbFormCode Then
        m.bOK = False
        Cancel = True
        Me.Hide
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibraryImport.Form_QueryUnload", , g.strAppPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: Clean up on form unload
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    Set m.Library = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibraryImport.Form_Unload", , g.strAppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtPath_GotFocus
'' Description: When the control gets the focus, select all the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtPath_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtPath
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibraryImport.txtPath_GotFocus", , g.strAppPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtPath_LostFocus
'' Description: When the path loses the focus, verify the information
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtPath_LostFocus()
On Error GoTo ErrSection:
    
    Dim astrFiles As cGdArray           ' Array of files in the current path
    Dim lIndex As Long                  ' Index into a for loop
    
    If Len(txtPath.Text) > 0 Then
        Set astrFiles = New cGdArray
        astrFiles.Create eGDARRAY_Strings
        
        astrFiles.SplitFields Trim(txtPath.Text), ","
        For lIndex = 0 To astrFiles.Size - 1
            astrFiles(lIndex) = StripStr(astrFiles(lIndex), Chr(34))
            If FileExt(astrFiles(lIndex)) <> "GLB" Then astrFiles(lIndex) = astrFiles(lIndex) & ".GLB"
            If InStr(astrFiles(lIndex), "\") = 0 Then astrFiles(lIndex) = AddSlash(g.strAppPath) & "Libraries\" & astrFiles(lIndex)
            If astrFiles.Size > 0 Then astrFiles(lIndex) = Chr(34) & astrFiles(lIndex) & Chr(34)
        Next lIndex
    
        txtPath.Text = astrFiles.JoinFields(",")
    
        SetLibraryInfo
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibraryImport.txtPath_LostFocus", , g.strAppPath
    
    txtPath.Text = ""
    SetLibraryInfo
    MoveFocus txtPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetLibraryInfo
'' Description: Set the library information from the given file
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetLibraryInfo()
On Error GoTo ErrSection:

    If VerifyLibraryInfo(False) Then
        ' Show library information...
        Set m.Library = New cLibrary
        With m.Library.PackagedFile
            .Path = StripStr(Parse(txtPath.Text, ",", 1), Chr(34))
            .GetPackagedLibraryInfo
            LibName.Text = .PackName
            Version.Text = Str(.PackVersion)
            LastModified.Text = DateFormat(.PackLastModified, MM_DD_YYYY, HH_MM_SS)
            Author.Text = .PackAuthor
        End With
    Else
        Set m.Library = New cLibrary
        LibName.Text = ""
        Version.Text = ""
        LastModified.Text = ""
        Author.Text = ""
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibraryImport.SetLibraryInfo", , g.strAppPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    VerifyLibraryInfo
'' Description: Verify the library information from the given file
'' Inputs:      Raise Error?
'' Returns:     True if verified, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function VerifyLibraryInfo(Optional ByVal bRaiseError As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function

    bReturn = False
    If Len(txtPath.Text) > 0 Then
        If FileExist(txtPath.Text) Then
            bReturn = True
        Else
            If bRaiseError Then
                MoveFocus txtPath
                Err.Raise vbObjectError + 1000, , "Specified Library does not exist"
            End If
        End If
    Else
        If bRaiseError Then
            MoveFocus txtPath
            Err.Raise vbObjectError + 1000, , "Please specify a Library to Import"
        End If
    End If
    
    VerifyLibraryInfo = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmLibraryImport.VerifyLibraryInfo", , g.strAppPath
    
End Function
