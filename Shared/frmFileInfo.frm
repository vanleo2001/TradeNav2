VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmFileInfo 
   Caption         =   "File Information"
   ClientHeight    =   4485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5790
   Icon            =   "frmFileInfo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   5790
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   435
      Left            =   1328
      TabIndex        =   1
      Top             =   3720
      Width           =   3135
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
      Caption         =   "frmFileInfo.frx":0442
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmFileInfo.frx":046E
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmFileInfo.frx":048E
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdSave 
         Height          =   435
         Left            =   1920
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
         Caption         =   "frmFileInfo.frx":04AA
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmFileInfo.frx":04E4
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmFileInfo.frx":0504
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Cancel          =   -1  'True
         Default         =   -1  'True
         Height          =   435
         Left            =   0
         TabIndex        =   2
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
         Caption         =   "frmFileInfo.frx":0520
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmFileInfo.frx":0546
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmFileInfo.frx":0566
         RightToLeft     =   0   'False
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid fgFileInfo 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      _cx             =   9763
      _cy             =   5953
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
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
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
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
End
Attribute VB_Name = "frmFileInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmFileInfo.frm
'' Description: Reads in a file (FileInfo.LST) which contains the files that
''              the program depends on, gets information on each of those files,
''              and displays the information in a grid
''
'' Author:      Genesis Financial Data Services
''              425 E Woodmen Rd
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date      Author      Description
'' 06/29/01  D Jarmuth   Created
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private Const C_NAME = 0
Private Const C_PATH = 1
Private Const C_SIZE = 2
Private Const C_DATETIME = 3
Private Const C_VERSION = 4
Private Const NUMCOLS = 5

Private Type mPrivate
    strFile As String
    bNeedVersionInfo As Boolean
End Type
Private m As mPrivate

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOK_Click
'' Description: If the user clicks on the OK button, unload the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()

    Unload Me

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Functions:   cmdSave_Click
'' Description: When the user clicks on the Save to File button, save the
''              information on the grid to a file
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdSave_Click()

    Dim fhOutput As Integer             ' File handle to the output file
    Dim lRow As Long                    ' Index into a for loop
    Dim lCol As Long                    ' Index into a for loop
    Dim strFile$
    
    strFile = Left(m.strFile, Len(m.strFile) - 3) & "TXT"
    fhOutput = FreeFile
    Open strFile For Output As #fhOutput
    
    With fgFileInfo
        For lRow = .FixedRows To .Rows - .FixedRows
            For lCol = 0 To NUMCOLS - 1
                Print #fhOutput, .Cell(flexcpText, lRow, lCol);
                If lCol < NUMCOLS - 1 Then Print #fhOutput, vbTab;
            Next lCol
            
            Print #fhOutput,
        Next lRow
    End With
    
    Close #fhOutput
    
    InfBox "h=Success ; i=i ; " & strFile & " successfully created"

End Sub

Private Sub Form_Activate()

    Dim strCaption$

    ' Get version info after rest of grid has shown
    ' (since it takes so long!)
    If m.bNeedVersionInfo Then
        m.bNeedVersionInfo = False
        strCaption = Me.Caption
        Me.Caption = strCaption & "  -->  GETTING VERSION INFO ..."
        DoEvents
        GetVersionInfo
        Me.Caption = strCaption
        DoEvents
        cmdOK.Enabled = True
        cmdSave.Enabled = True
    End If
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

#If TRADENAV_EXE Then
    If KeyCode = vbKeyF1 Then
        KeyCode = 0
        g.Help.ShowF1Help Me
    End If
#End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFileInfo.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: When the form is loaded, initialize the grid and fill it in
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()

    Dim strFileName As String           ' Filename to search on
    Dim fhInput As Integer              ' File handle to the input file
    Dim strBuffer As String             ' Buffer from the input file
    Dim strTemp As String               ' Temporary string variable

    ' Center the form on the screen
    Me.Move Me.Left, Me.Top, 9540, 6855
    CenterTheForm Me
    
    g.Styler.StyleForm Me

    ' Initialize the grid
    With fgFileInfo
        .Redraw = flexRDNone
        .ExplorerBar = flexExSortShowAndMove
        .SelectionMode = flexSelectionListBox
        .Editable = flexEDNone
        .ExtendLastCol = True
        .SheetBorder = RGB(128, 128, 128)
        .AllowUserResizing = flexResizeColumns
        
        .Rows = 1
        .FixedRows = 1
        .Cols = NUMCOLS
        .FixedCols = 0
        
        .Cell(flexcpText, 0, C_NAME) = "Name"
        .Cell(flexcpText, 0, C_PATH) = "Path"
        .Cell(flexcpText, 0, C_SIZE) = "Size"
        .Cell(flexcpText, 0, C_DATETIME) = "Date/Time"
        .Cell(flexcpText, 0, C_VERSION) = "Version"
        
        .ColDataType(C_DATETIME) = flexDTDate
        .ColFormat(C_DATETIME) = DateAndTime("Format")
        .ColAlignment(C_DATETIME) = flexAlignLeftCenter
        
        If FileExist(m.strFile) Then
            Screen.MousePointer = vbHourglass
            fhInput = FreeFile
            Open m.strFile For Input As #fhInput
            Do While Not EOF(fhInput)
                Line Input #fhInput, strBuffer
                Select Case Left(strBuffer, 2)
                    Case "@\"           ' Application Path
                        strFileName = AddSlash(App.Path) & Right(strBuffer, Len(strBuffer) - 2)
                    Case "$\"           ' System Path
                        strFileName = WinSysPath & Right(strBuffer, Len(strBuffer) - 2)
                    Case "!\"           ' Windows Path
                        strFileName = WindowsPath & Right(strBuffer, Len(strBuffer) - 2)
                    Case "&\"           ' Shared Self-Reg Path
                        strFileName = App.Path & "\..\SharedSelfReg\" & Right(strBuffer, Len(strBuffer) - 2)
                End Select
                
                strTemp = Dir(strFileName)
                Do While strTemp <> ""
                    AddToGrid AddSlash(FilePath(strFileName)) & strTemp
                    strTemp = Dir
                Loop
            Loop
            Close #fhInput
            Screen.MousePointer = vbNormal
        End If
        
        .Select 1, 0
        .Sort = flexSortGenericAscending
        .Select 1, 0
        .AutoSize 0, NUMCOLS - 1
        .Redraw = flexRDBuffered
    End With
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If UnloadMode = 0 Then
        If cmdOK.Enabled = False Then
            Cancel = True
            Beep
        End If
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Resize
'' Description: As the form is resized, resize the grid to go with it
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Resize()

    If LimitFormSize(Me, fraButtons.Width * 1.1, cmdOK.Height * 5) Then Exit Sub
    
    fgFileInfo.Width = Me.ScaleWidth - fgFileInfo.Left * 2
    fgFileInfo.Height = Me.ScaleHeight - cmdOK.Height - fgFileInfo.Top * 3
    
    fraButtons.Top = fgFileInfo.Height + fgFileInfo.Top * 2
    fraButtons.Left = Me.ScaleWidth / 2 - fraButtons.Width / 2

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Show the form
'' Inputs:      None
'' Returns:     True on success, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe() As Boolean

    ' look in INFO folder first, then in APP folder
    m.strFile = App.Path & "\Info\FileInfo.LST"
    If Not FileExist(m.strFile) Then
        m.strFile = App.Path & "\FileInfo.LST"
    End If
    If Not FileExist(m.strFile) Then
        InfBox "h=Error ; i=! ; Could not find FileInfo.LST"
        ShowMe = False
        Exit Function
    End If
    
    m.bNeedVersionInfo = True
    cmdOK.Enabled = False
    cmdSave.Enabled = False
    ShowForm Me, True
    
    ShowMe = True

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddToGrid
'' Description: Adds a file to the grid
'' Inputs:      Filename with path
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddToGrid(ByVal pstrFilename As String)

    Dim strFile As String               ' Local copy of the filename

    With fgFileInfo
        strFile = FileNameDisplay(FullPathName(pstrFilename))
        
        .Rows = .Rows + 1
        .Cell(flexcpText, .Rows - 1, C_NAME) = FileBase(strFile) & "." & FileExt(strFile)
        .Cell(flexcpText, .Rows - 1, C_PATH) = FilePath(strFile)
        .Cell(flexcpText, .Rows - 1, C_SIZE) = CStr(FileLength(strFile))
        .Cell(flexcpText, .Rows - 1, C_DATETIME) = FileDate(strFile)
        '.Cell(flexcpText, .Rows - 1, C_VERSION) = FileVersion(strFile)
    End With

End Sub

' Get version info for each file
' (done after rest of grid filled in since takes so much time)
Private Sub GetVersionInfo()
    
    Dim iRow&, strFile$
    
    With fgFileInfo
        .Redraw = flexRDNone
        Screen.MousePointer = vbHourglass
        For iRow = .FixedRows To .Rows - 1
            strFile = AddSlash(.TextMatrix(iRow, C_PATH)) & .TextMatrix(iRow, C_NAME)
            .Cell(flexcpText, iRow, C_VERSION) = FileVersion(strFile)
        Next
        Screen.MousePointer = vbNormal
        .Redraw = flexRDBuffered
    End With
    
End Sub

