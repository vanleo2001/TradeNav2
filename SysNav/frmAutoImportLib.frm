VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmAutoImportLib 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Auto Import..."
   ClientHeight    =   1305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7320
   ControlBox      =   0   'False
   Icon            =   "frmAutoImportLib.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   7320
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniButtonImageXP cmdCorner 
      Height          =   375
      Left            =   6195
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   960
      Visible         =   0   'False
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
      Caption         =   "frmAutoImportLib.frx":0A02
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmAutoImportLib.frx":0A2E
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmAutoImportLib.frx":0A4E
      RightToLeft     =   0   'False
   End
   Begin VSFlex7LCtl.VSFlexGrid vsStatus 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   705
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
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   8421504
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483639
      GridColorFixed  =   -2147483639
      TreeColor       =   -2147483632
      FloodColor      =   16711680
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
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
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin HexUniControls.ctlUniLabelXP lblHeader 
      Height          =   285
      Left            =   120
      Top             =   240
      Width           =   6960
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
      Caption         =   "frmAutoImportLib.frx":0A6A
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmAutoImportLib.frx":0B12
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmAutoImportLib.frx":0B32
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblStatusMsg 
      Height          =   285
      Left            =   1725
      Top             =   720
      Visible         =   0   'False
      Width           =   5415
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
      Caption         =   "frmAutoImportLib.frx":0B4E
      BackColor       =   -2147483633
      ForeColor       =   -2147483640
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmAutoImportLib.frx":0B70
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmAutoImportLib.frx":0B90
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
End
Attribute VB_Name = "frmAutoImportLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmAutoImportLib.frm
'' Description: Shows the user status information while an auto import of
''              library(s) is done
''
'' Author:      Genesis Financial Data Services
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' ??/??/??     M Thorne    Created
'' 04/17/2009   DAJ         Fix pyramiding information after auto import
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Start the auto importing
'' Inputs:      File to import
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowMe(strFile As String)
On Error GoTo ErrSection:

    Dim LibExport As cPackagedFile      ' Packaged file to import
    Dim LibMgrBridge As cLibManagerBridge ' Bridge to the library manager
    Dim fhCfgFile As Integer            ' File handle to the configuration file
    Dim strBuffer As String             ' Buffer read in from the config file
    Dim strAddDeleteMode As String      ' e.g. "-" means delete
    Dim bLibExists As Boolean           ' Does the library exist?
    Dim bTablesImported As Boolean      ' Were any tables imported?
    Dim astrLibraries As New cGdArray   ' Array of libraries to import
    Dim lIndex As Long                  ' Index for a for loop
    Dim strSecondLib As String          ' Secondary library

    ShowForm Me

    Set LibExport = New cPackagedFile

    Set LibMgrBridge = New cLibManagerBridge
    With LibMgrBridge
        .CalledFrom = SystemNavigator
        .AppPath = App.Path
        .dbNavPassword = DbPassword
        .dbNavRef = g.dbNav
        .ImageList = frmMain.img16.ListImages
        .CustomerID = g.lLCD
        .Help = g.Help
        .HonestDate = RI_HonestDate
    End With

    With LibExport
        .StatusBar = vsStatus
        .StatusMsg = lblStatusMsg
        '.Path = App.Path & "\Genesis System Functions.txt"
        
        ' Default the tables imported to false
        bTablesImported = False
        astrLibraries.Create eGDARRAY_Strings
        
        ' If a config file was passed in, read through it and process each entry
        If UCase(Right(strFile, 4)) = ".CFG" Then
            
            ' Open the config file
            fhCfgFile = FreeFile
            Open strFile For Input As #fhCfgFile
            
            ' Walk through the config file processing each entry
            Do While Not EOF(fhCfgFile)
                ' Get the filename and mode flag
                Line Input #fhCfgFile, strBuffer
                strAddDeleteMode = UCase(Parse(strBuffer, ",", 1))
                strSecondLib = Parse(strBuffer, ",", 3)
                strBuffer = Parse(strBuffer, ",", 2)
                .Path = App.Path & "\TempLib\" & strBuffer
                If Len(strBuffer) > 0 Then
                    If strAddDeleteMode = "-" Then
                        ' delete this library
                        astrLibraries.Add "-" & .Path
                    ElseIf FileExist(.Path) Then
                        ' Get the library information out of the file
                        .GetPackagedLibraryInfo
                        If HasModule(.PackRequiredMod) Or IsIDE Then
                            astrLibraries.Add .Path
                        Else
                            astrLibraries.Add "-" & .Path
                        End If
                    End If
                End If
            Loop
            
            ' Close the config file
            Close #fhCfgFile
        
        ' A text file was passed in, so just process that one file
        Else
            astrLibraries.Add strFile
        End If
        
        .ImportTables astrLibraries.ArrayHandle, lblHeader
        
        ' Clean up any bogus pyramid information (but don't reload the rules table because
        ' that will happen a little later anyway)...
        FixPyramidInfo False
        
        g.bDirtyLibrariesMDB = True
    End With

   
NormalExit:
    Set LibExport = Nothing
    
ErrExit:
    Unload Me
    Exit Sub

ErrSection:
    RaiseError "frmAutoImportLib.Run", eGDRaiseError_Show
    If LibExport Is Nothing Then
        Resume ErrExit
    Else
        Resume NormalExit
    End If

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
    RaiseError "frmAutoImportLib.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: When the form is loaded, size and center the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:
    
    ReSizeMDIChildForm Me, cmdCorner
    CenterTheForm Me
    Screen.MousePointer = vbDefault
    
    g.Styler.StyleForm Me
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmAutoImportLib.Form.Load", eGDRaiseError_Show
    Resume ErrExit

End Sub

