VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmDataMan 
   Caption         =   "Browse for Data Files"
   ClientHeight    =   4725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11190
   Icon            =   "DataMan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4725
   ScaleWidth      =   11190
   StartUpPosition =   2  'CenterScreen
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   495
      Left            =   6660
      TabIndex        =   10
      Top             =   3900
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
      Caption         =   "DataMan.frx":030A
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "DataMan.frx":0336
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "DataMan.frx":0356
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Cancel          =   -1  'True
         Height          =   495
         Left            =   1320
         TabIndex        =   7
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
         Caption         =   "DataMan.frx":0372
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "DataMan.frx":03A0
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "DataMan.frx":03C0
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Height          =   495
         Left            =   0
         TabIndex        =   11
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
         Caption         =   "DataMan.frx":03DC
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "DataMan.frx":0402
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "DataMan.frx":0422
         RightToLeft     =   0   'False
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid fgSecurities 
      Height          =   3495
      Left            =   4800
      TabIndex        =   9
      Top             =   180
      Width           =   5655
      _cx             =   9975
      _cy             =   6165
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
      ScrollTrack     =   0   'False
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
   Begin HexUniControls.ctlUniFrameWL fraHardDrive 
      Height          =   3315
      Left            =   100
      TabIndex        =   1
      Top             =   100
      Width           =   4455
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
      Caption         =   "DataMan.frx":043E
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "DataMan.frx":045E
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "DataMan.frx":047E
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniCheckXP chkGDB 
         Height          =   255
         Left            =   3060
         TabIndex        =   8
         Top             =   3000
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
         Caption         =   "DataMan.frx":049A
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   -1  'True
         Tip             =   "DataMan.frx":04D0
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "DataMan.frx":04F0
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniFileBoxXP dirPathTree 
         Height          =   2565
         Left            =   600
         TabIndex        =   6
         Top             =   360
         Width           =   3855
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         SelBackColor    =   -2147483635
         SelForeColor    =   -2147483634
         RoundedBorders  =   0   'False
         SelectorStyle   =   -1
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
         Tip             =   "DataMan.frx":050C
         Path            =   ""
         Pattern         =   "*.*"
         PatternAlsoForDirs=   0   'False
         ReadOnly        =   -1  'True
         System          =   0   'False
         Hidden          =   0   'False
         PermitNavigation=   -1  'True
         MultiSelect     =   0
         HScroll         =   -1  'True
         ShowFullPath    =   0   'False
         DisplayMode     =   2
         MousePointer    =   0
         MouseIcon       =   "DataMan.frx":052C
      End
      Begin HexUniControls.ctlUniDriveBoxXP drvDrive 
         Height          =   315
         Left            =   600
         TabIndex        =   5
         Top             =   0
         Width           =   3855
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         ButtonBackColor =   0
         ButtonForeColor =   0
         ButtonStyle     =   -1
         SelectorStyle   =   -1
         SelBackColor    =   -2147483635
         SelForeColor    =   -2147483634
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
         Tip             =   "DataMan.frx":0548
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "DataMan.frx":0568
         DropDownWidth   =   -1
      End
      Begin HexUniControls.ctlUniCheckXP chkCSI 
         Height          =   255
         Left            =   660
         TabIndex        =   4
         Top             =   3000
         Width           =   615
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
         Caption         =   "DataMan.frx":0584
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   -1  'True
         Tip             =   "DataMan.frx":05AA
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "DataMan.frx":05CA
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkMetaStock 
         Height          =   255
         Left            =   1680
         TabIndex        =   3
         Top             =   3000
         Width           =   1095
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
         Caption         =   "DataMan.frx":05E6
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   -1  'True
         Tip             =   "DataMan.frx":0618
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "DataMan.frx":0638
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkGenTick 
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   2760
         Visible         =   0   'False
         Width           =   975
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
         Caption         =   "DataMan.frx":0654
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "DataMan.frx":0682
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "DataMan.frx":06A2
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label1 
         Height          =   255
         Left            =   0
         Top             =   0
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
         Caption         =   "DataMan.frx":06BE
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "DataMan.frx":06EA
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "DataMan.frx":070A
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label2 
         Height          =   255
         Left            =   0
         Top             =   360
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
         Caption         =   "DataMan.frx":0726
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "DataMan.frx":0750
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "DataMan.frx":0770
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdMarket 
      Height          =   375
      Left            =   9480
      TabIndex        =   0
      Top             =   4080
      Visible         =   0   'False
      Width           =   1095
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
      Caption         =   "DataMan.frx":078C
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "DataMan.frx":07C4
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "DataMan.frx":07E4
      RightToLeft     =   0   'False
   End
End
Attribute VB_Name = "frmDataMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        DataMan.FRM
'' Description: Routines for handling a data manager to select a CSI or MS7
''              file from a directory
''
'' Author:      Genesis Financial Data Services
''              425 E Woodmen Rd
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date      Author      Description
'' 08/19/99  D Jarmuth   Created
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Const kExtendedCol = 1

Private Enum eGDCols
    eGDCol_Symbol = 0
    eGDCol_Name
    eGDCol_Period
    eGDCol_SecType
    eGDCol_Format
    eGDCol_File
    eGDCol_NumCols
End Enum

Private Type mPrivate
    bOK As Boolean
    lSortedCol As Long
    lSortOrder As Long
    lTreeIndex As Long                  ' Index in the tree
    lColumn As Long                     ' Last column selected in list view
    
    nPrevColWidth As Long               ' Used for Custom Extended Column
End Type
Private m As mPrivate

' Private structures
Private Type CSImaster_rec              ' CSI QMaster record
    csiNum As String * 4
    Name As String * 20
    FileType As String * 1
    DelivMonth As String * 2
    DelivYear As String * 2
    ConvFactor As String * 2
    PricingUnit As String * 5
    Symbol2 As String * 2
    SecFlag As String * 1
    OptionFlag As String * 1
    StrikingPrice As String * 5
    Symbol6 As String * 6
    Deleted As String * 1
    Century As String * 2 'Reserved As String * 2
    DataHdr As String * 7
    DispFactor As String * 2
    ExpByte As String * 1
End Type

Private Type MSmaster_rec               ' MetaStock Master Record
    FileNum As String * 1
    'FileType As Integer
    FileType As String * 1
    FileType2 As String * 1  ' do this because of OmniTrader!
    RecLength As String * 1
    NumRFields As String * 1
    reserved As String * 1 'NumIFields As String * 1
    Century As String * 1 'NumBFields As String * 1
    Name As String * 16
    reserved1 As Integer
    BeginDate As Single
    LastDate As Single
    DataFormat As String * 1
    IntraTime As Integer
    Symbol As String * 14
    reserved2 As String * 1
    Flag As String * 1
    reserved3 As String * 1
End Type

Private Function GDCol(ByVal lColumn As eGDCols) As Long
    GDCol = lColumn
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: If the user hits cancel, clear out the public data members
''              and return
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
    RaiseError "frmDataMan.cmdCancel.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdMarket_Click
'' Description: If the user hits the Market Information button, show them the
''              Market Information form for the selected symbol
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdMarket_Click()
On Error GoTo ErrSection:
    
    With fgSecurities
        frmMarkets.ShowMe "*" & .TextMatrix(.SelectedRow(0), GDCol(eGDCol_Symbol)), .TextMatrix(.SelectedRow(0), GDCol(eGDCol_Name))
    End With
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmDataMan.cmdMarket.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOK_Click
'' Description: If the user hits the OK button, copy the necessary information
''              into the public data members and exit the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
On Error GoTo ErrSection:
    
    DoCheck

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDataMan.cmdOK.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    dirPathTree_KeyUp
'' Description: When the user presses a key on the tree, if the user pressed
''              ENTER then open up the tree at that level, otherwise call
''              the PathChanged function to see if we can fill the list view
'' Inputs:      What key was hit, shift/ctrl/alt status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub dirPathTree_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    ' If the user pressed enter, expand the tree at that point
    If (KeyCode = vbKeyReturn) Then
        If (Right(dirPathTree.List(dirPathTree.ListIndex), 1) <> "\") Then
            dirPathTree.Path = dirPathTree.List(dirPathTree.ListIndex) + "\"
        End If
    End If
    
    PathChanged False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDataMan.dirPathTree.KeyUp", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    dirPathTree_MouseUp
'' Description: When the user clicks on a different item in the tree, see
''              if we can fill the list view or not
'' Inputs:      Which mouse button was hit, shift/ctrl/alt status, location
''              mouse was hit at
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub dirPathTree_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    PathChanged False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDataMan.dirPathTree.MouseUp", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    drvDrive_Change
'' Description: When the user changes the drive in the drive control, change
''              the path in the directory control
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub drvDrive_Change()
On Error GoTo ErrSection:

    dirPathTree.Path = drvDrive.Drive
    PathChanged False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDataMan.drvDrive.Change", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgSecurities_AfterRowColChange
'' Description: After a user moves rows in the grid, make sure the new row is
''              selected
'' Inputs:      Old Row and Column, New Row and Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgSecurities_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    With fgSecurities
        'TLB: only change if not already correct,
        'otherwise causes recursion that can overflow stack!
        If .RowSel <> NewRow Then .RowSel = NewRow
        If .Row <> NewRow Then .Row = NewRow
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDataMan.fgSecurities.AfterRowColChange", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgSecurities_AfterSort
'' Description: After a user sorts the grid, make sure to show the selected
''              row and save off the new sort information
'' Inputs:      Column sorted, Order sorted in
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgSecurities_AfterSort(ByVal Col As Long, Order As Integer)
On Error GoTo ErrSection:

    With fgSecurities
        .ShowCell .SelectedRow(0), GDCol(eGDCol_Symbol)
        CenterSelection
    End With
    m.lSortedCol = Col
    m.lSortOrder = Order

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDataMan.fgSecurities.AfterSort", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgSecurities_BeforeScrollTip
'' Description: Show the user the symbol of the first visible row as they are
''              scrolling
'' Inputs:      First Visible Row
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgSecurities_BeforeScrollTip(ByVal Row As Long)
On Error GoTo ErrSection:

    With fgSecurities
        .ScrollTipText = .TextMatrix(Row, GDCol(eGDCol_Symbol))
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDataMan.fgSecurities.BeforeScrollTip", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgSecurities_DblClick
'' Description: If a user double clicks on an entry in the grid, click on the
''              OK button for them
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgSecurities_DblClick()
On Error GoTo ErrSection:

    DoCheck
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDataMan.fgSecurities.DblClick", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgSecurities_AfterSort
'' Description: If the user hits Enter on a symbol in the grid, click on the
''              OK button for them
'' Inputs:      Key Pressed, Shift/Ctrl/Alt status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgSecurities_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If KeyCode = vbKeyReturn Then
        DoCheck
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDataMan.fgSecurities.KeyUp", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Activate
'' Desciption:  Each time that the form is activated, set the focus to the
''              securities list
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Activate()
On Error GoTo ErrSection:

    MoveFocus fgSecurities
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmDataMan.Form.Activate", eGDRaiseError_Show
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
    RaiseError "frmDataMan.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: When the form loads, set up the column headers on the list
''              view control and set up some variables
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:
    
    g.Styler.StyleForm Me
    
    Dim strPlacement As String
    
    If FileExist("C:\Common\Files.EXE") Then cmdMarket.Visible = True
    
    strPlacement = GetIniFileProperty("DataMan", "", "Placement", g.strIniFile)
    If strPlacement <> "" Then SetFormPlacement Me, strPlacement, "LT"
    CenterTheForm Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDataMan.Form.Load", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PathChanged
'' Description: When the user moves in the tree, we need to try to load data
''              from the new directory into the list view control
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PathChanged(ByVal bForce As Boolean)
On Error GoTo ErrSection:

    Dim iCountCSI As Integer            ' Number of records in the Qmaster
    Dim iCountMS As Integer             ' Number of records in the Master
    Dim iCountGT As Integer             ' Number of GT files in directory
    Dim iCountGDB As Integer            ' Number of GDB files in directory
    Dim strSymbol As String             ' Current symbol
    Dim strPath As String
        
    ' If we have in fact moved, try to load new data
    If (dirPathTree.ListIndex <> m.lTreeIndex) Or (bForce = True) Then
        With fgSecurities
            .Redraw = flexRDNone
            If .SelectedRow(0) >= .FixedRows Then strSymbol = .TextMatrix(.SelectedRow(0), GDCol(eGDCol_Symbol))
            .Rows = .FixedRows

            iCountCSI = 0
            iCountMS = 0
            iCountGT = 0
        
            strPath = AddSlash(dirPathTree.List(dirPathTree.ListIndex))
        
            If chkCSI.Value = vbChecked Then
                'iCountCSI = LoadQmaster(dirPathTree.List(dirPathTree.ListIndex))
                iCountCSI = LoadMasterFile(strPath, "CSI")
            End If
        
            If chkMetaStock.Value = vbChecked Then
                'iCountMS = LoadMSmaster(dirPathTree.List(dirPathTree.ListIndex))
                iCountMS = LoadMasterFile(strPath, "MS")
            End If
        
            If chkGenTick.Value = vbChecked Then
                iCountGT = LoadGTFiles(strPath)
            End If
            
            If chkGDB.Value = vbChecked Then
                iCountGDB = LoadGDBFiles(strPath)
            End If
        
            If iCountCSI + iCountMS + iCountGT + iCountGDB > 0 Then
                cmdOK.Enabled = True
                cmdMarket.Enabled = True
            Else
                cmdOK.Enabled = False
                cmdMarket.Enabled = False
            End If
            
            .ColSort(m.lSortedCol) = m.lSortOrder
            .Sort = flexSortUseColSort
            SelectSymbol strSymbol, True
            
            ExtendCustomColumn
            .Redraw = flexRDBuffered
        End With
        
        m.lTreeIndex = dirPathTree.ListIndex
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmDataMan.PathChanged", eGDRaiseError_Raise
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadMasterFile
'' Description: Loads information from the Qmaster/Master/Xmaster in the given path
'' Inputs:      Path to find the master file
'' Returns:     Number of records in the master file
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function LoadMasterFile(ByVal strPath As String, ByVal strFormat As String) As Long
On Error GoTo ErrSection:
    
    Dim i&, strSecType$, iSecType&
    Dim tblMatches As cGdTable
    
    Set tblMatches = GetMasterFileMatches(strPath, strFormat)
    ' - returns table: 0=Symbol, 1=Period (d/w/m/##), 2=Desc, 3=Filename, 4=Format, 5=File_Num, 6=ConvFact, 7=SecFlag
    For i = 0 To tblMatches.NumRecords - 1
        iSecType = tblMatches(7, i)
        If iSecType > 0 Then
            strSecType = Chr(iSecType)
        Else
            strSecType = ""
        End If
        AddSymbolToList tblMatches(0, i), strSecType, tblMatches(1, i), tblMatches(2, i), tblMatches(4, i), tblMatches(3, 1)
    Next
    ExtendCustomColumn
    LoadMasterFile = tblMatches.NumRecords
    
ErrExit:
    Set tblMatches = Nothing
    Exit Function

ErrSection:
    RaiseError "frmDataMan.LoadMasterFile", eGDRaiseError_Raise
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadQmaster
'' Description: Loads information from the Qmaster in the given path
'' Inputs:      Path to find the Qmaster
'' Returns:     Number of records in the Qmaster
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function LoadQmaster(ByVal strPath As String)
On Error GoTo ErrSection:

    Dim iCount As Integer               ' Number of records in Qmaster
    Dim strFileName As String           ' Full path and filename of Qmaster
    Dim fh As Integer                   ' File handle of the Qmaster file
    Dim bAdd As Integer                 ' Do we want to add this record?
    Dim iRec As Integer                 ' Counter for a for loop
    Dim strTemp As String               ' Temporary holder of the symbol
    Dim strSecType As String            ' Security Type in the current record
    Dim strSymbol As String             ' Symbol in the current record
    Dim CSI As CSImaster_rec            ' Record out of the Qmaster
    Dim lRedraw As Long                 ' Current state of the grid redraw
    
    ' Reset counter variable
    iCount = 0
    
    ' Setup the Qmater filename and full path
    strFileName = AddSlash(strPath) + "QMASTER."
    
    ' If the Qmaster exists, read through it
    If Dir(strFileName) <> "" Then
        lRedraw = fgSecurities.Redraw
        fgSecurities.Redraw = flexRDNone
        
        ' Try to open the Qmaster file
        fh = FreeFile
        Open strFileName For Random As #fh Len = Len(CSI)
        
        ' Walk through all of the records in the Qmaster
        For iRec = 1 To LOF(fh) \ Len(CSI)
            Get #fh, iRec, CSI
            FixCsiRec CSI   '1/26/98

            ' Fix some old symbols that need to be changed
            If CSI.Symbol2 = "EK" Then
                If InStr(CSI.Name, "(") = 0 Then
                    CSI.Symbol2 = "LR"
                    CSI.Name = "EUROMARK (LIFFE)"
                    CSI.csiNum = "182"
                    Put #fh, iRec, CSI
                End If
            End If
            If CSI.Symbol2 = "SG" Then
                If InStr(UCase(CSI.Name), "SUGAR") > 0 Then
                    CSI.Symbol2 = "SE"
                    Put #fh, iRec, CSI
                End If
            End If

            ' If we have hit the last record, exit the for loop
            If CSI.csiNum = "9999" Then Exit For
            
            ' Check for deleted/bad record
            If Asc(CSI.csiNum) < 32 Or Asc(CSI.csiNum) > 64 Then CSI.Deleted = "1"
            bAdd = True
            
            ' If we have a deleted record, don't add it
            If CSI.Deleted = "1" Then
                CSI.csiNum = Space(30)
                CSI.Name = Space(30)
                CSI.Symbol2 = Space(30)
                CSI.Symbol6 = Space(30)
                bAdd = False
            End If
            
            ' If we want to add the record, get the information out of it
            If bAdd Then
                CSI.Symbol2 = UCase(CSI.Symbol2)
                CSI.Symbol6 = UCase(CSI.Symbol6)

                strSecType = CSI.SecFlag
                If strSecType = "C" Or strSecType = "F" Then
                    strSecType = "F"
                Else
                    strTemp = Trim(CSI.Symbol6)
                    If Left(strTemp, 1) = "$" Then
                        strSecType = "I"
                    ElseIf Len(strTemp) = 5 And Mid(strTemp, 5, 1) = "X" And InStr(strTemp, "_") = 0 And InStr(strTemp, " ") = 0 And Left(strTemp, 1) <> "$" Then
                        strSecType = "M"
                    Else
                        strSecType = "S"
                    End If
                End If
                If CSI.OptionFlag <> "N" And CSI.OptionFlag <> " " Then
                    strSecType = strSecType + "O"
                End If

                ' Is the symbol bogus?
                If Left(CSI.Symbol6, 1) = "@" And Right(CSI.Symbol6, 3) = "996" And Len(Trim(CSI.Symbol2)) > 0 Then
                    CSI.Symbol6 = " "
                End If

                If Asc(Left(CSI.Symbol6, 1)) <= 32 Or CSI.Symbol6 = "999999" Then
                    If Left(strSecType, 1) = "F" Then
                        strSymbol = Trim(CSI.Symbol2) + "-" + CSI.DelivYear + CSI.DelivMonth
                    Else
                        strSymbol = CSI.Symbol2
                    End If
                ElseIf Left(CSI.Symbol6, 1) = "@" Then
                    strSymbol = Trim(Mid(CSI.Symbol6, 2))
                    If Mid(strSymbol, 2, 1) = " " Then strSymbol = Trim(Left(strSymbol, 1) + Mid(strSymbol, 3))
                Else
                    strSymbol = CSI.Symbol6
                End If
                strSymbol = Trim(strSymbol)
                If Len(strSymbol) <= 3 And Left(strSecType, 1) = "F" Then
                    strSymbol = Trim(strSymbol + "-" + CSI.DelivYear + CSI.DelivMonth)
                End If
                If Left(strSecType, 1) = "F" And Len(strSymbol) >= 5 Then
                    strTemp = Mid(strSymbol, Len(strSymbol) - 4, 1)
                    If strTemp <> "-" Then  'And strTemp >= "0" And strTemp <= "9" Then
                        strSymbol = Trim(Left(strSymbol, Len(strSymbol) - 4)) + "-" + Trim(Mid(strSymbol, Len(strSymbol) - 3))
                    End If
                End If
                If Left(strSecType, 1) = "I" And Left(strSymbol, 1) >= "A" Then strSymbol = "$" + strSymbol

                If Len(Trim(CSI.Name)) = 0 Then CSI.Name = strSymbol
                
                ' Add the record to the list view component
                AddSymbolToList strSymbol, strSecType, CSI.FileType, CSI.Name, "CSI", "F" & Format(iRec, "000") & ".DTA"
                
                ' Increment the count
                iCount = iCount + 1
            End If
        Next 'iRec
        Close #fh
        
        ExtendCustomColumn
        fgSecurities.Redraw = lRedraw
    End If
    LoadQmaster = iCount
    
ErrExit:
    Exit Function

ErrSection:
    LoadQmaster = -1
    RaiseError "frmDataMan.LoadQmaster", eGDRaiseError_Raise

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FixCsiRec
'' Description: Fixes some potential problems in a Qmaster record
'' Inputs:      Qmaster record
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FixCsiRec(CSI As CSImaster_rec)
On Error GoTo ErrSection:

    If Asc(CSI.FileType) <= 32 Then CSI.FileType = "D"

    CSI.SecFlag = UCase(CSI.SecFlag)
    If CSI.SecFlag = "F" Then CSI.SecFlag = "C"
    'If csi.SecFlag <> "C" Then csi.SecFlag = "S"

    CSI.OptionFlag = UCase(CSI.OptionFlag)
    If CSI.OptionFlag <> "P" And CSI.OptionFlag <> "C" Then
        CSI.OptionFlag = "N"
        CSI.StrikingPrice = " "
    End If

    If Asc(CSI.Symbol2) <= 32 Then CSI.Symbol2 = "  "
    If CSI.SecFlag = "C" And Len(Trim(CSI.Symbol2)) > 0 Then
        CSI.Symbol6 = CSI.Symbol2 + Space(4)
    ElseIf Asc(CSI.Symbol6) <= 32 Then
        CSI.Symbol6 = Space(6)
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDataMan.FixCsiRec", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadMSmaster
'' Description: Load records from a MetaStock master record
'' Inputs:      Path of the MetaStock master record
'' Returns:     Number of records in the master
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function LoadMSmaster(ByVal strPath$)
On Error GoTo ErrSection:

    Dim fh%                             ' File handle of the master file
    Dim i%                              ' Place holder for strings
    Dim iCount%                         ' Number of records in the master
    Dim bAdd%                           ' Do we want to add this record?
    Dim iRec%                           ' Counter for a for loop
    Dim strTemp$                        ' Temporary symbol holder
    Dim strFileName$                    ' Filename and path of the master
    Dim strSecType$                     ' Security Type in the record
    Dim strSymbol$                      ' Symbol in the record
    Dim ms As MSmaster_rec              ' Metastock master record
    Dim lRedraw As Long                 ' Current state of the grid redraw
    
    ' Reset the counter
    iCount = 0
    
    ' Set up the filename and full path of the master
    If (Right(strPath, 1) <> "\") Then strPath = strPath + "\"
    strFileName = strPath + "MASTER."
    
    ' If the master file exists, read through it
    If Dir(strFileName) <> "" Then
        lRedraw = fgSecurities.Redraw
        fgSecurities.Redraw = flexRDNone
        
        ' Try to open the master file
        fh = FreeFile
        Open strFileName For Random As #fh Len = Len(ms)
        
        ' Walk through all of the records in the master file
        For iRec = 2 To LOF(fh) \ Len(ms)
            Get #fh, iRec, ms

            ' Change old symbols
            If Left(ms.Symbol, 3) = "@EK" Then
                If InStr(ms.Name, "(") = 0 And InStr(ms.Symbol, "227") > 0 Then
                    ms.Symbol = "@LR" + Mid(ms.Symbol, 4, 6) + "182" + Mid(ms.Symbol, 13)
                    ms.Name = "EUROMARK (LIFFE)"
                    Put #fh, iRec, ms
                End If
            End If
            If Left(ms.Symbol, 3) = "@SG" Then
                If InStr(UCase(ms.Name), "SUGAR") > 0 Then
                    ms.Symbol = "@SE" + Mid(ms.Symbol, 4)
                    Put #fh, iRec, ms
                End If
            End If
            
            ' Clear NULL in case it got there accidentally
            i = InStr(ms.Name, Chr(0))
            If i > 0 Then
                ms.Name = Left(ms.Name, i - 1)
                Put #fh, iRec, ms
            End If
            i = InStr(ms.Symbol, Chr(0))
            If i > 0 Then
                ms.Symbol = Left(ms.Symbol, i - 1)
                Put #fh, iRec, ms
            End If

            bAdd = True
            If Asc(ms.FileNum) = 0 Or (Asc(ms.FileType) <> 101 And Asc(ms.FileType2) <> 101) Then
                bAdd = False
                ms.FileNum = Chr(0)
                ms.Symbol = ""
                ms.Name = ""
                ms.BeginDate = 0
                ms.LastDate = 0
                On Error Resume Next
                On Error GoTo ErrSection
            End If
            If bAdd Then
                strSymbol = Trim(Left(UCase(ms.Symbol), 8))

                If InStr(Left(strSymbol, 4), "-") Then
                    strSecType = "F"
                ElseIf Left(strSymbol, 1) = "@" Then
                    strSecType = "F"
                    strSymbol = Mid(strSymbol, 2)
                    If AreDigits(strSymbol, 3, 4) And Len(strSymbol) = 7 Then strSymbol = Left(strSymbol, 6)
                    If Left(strSecType, 1) = "F" And Len(strSymbol) >= 5 Then
                        strTemp = Mid(strSymbol, Len(strSymbol) - 4, 1)
                        If strTemp <> "-" Then  'And strTemp >= "0" And strTemp <= "9" Then
                            strSymbol = Trim(Left(strSymbol, Len(strSymbol) - 4)) + "-" + Trim(Mid(strSymbol, Len(strSymbol) - 3))
                        End If
                    End If
                Else
                    strSecType = "S"
                    If Left(strSymbol, 1) = "*" Or Left(strSymbol, 1) = "$" Then
                        strSecType = "I"
                    ElseIf Len(Trim(strSymbol)) = 5 And Mid(strSymbol, 5, 1) = "X" And InStr(strSymbol, "_") = 0 And InStr(strSymbol, " ") = 0 Then
                        strSecType = "M"
                    End If
                    If strSecType = "S" And (InStr(strPath, "IDX\") > 0 Or InStr(strPath, "INDEXES\") > 0 Or InStr(strPath, "INDICES\") > 0) Then strSecType = "I"
                End If
                If Left(strSecType, 1) = "I" And Left(strSymbol, 1) >= "A" Then strSymbol = "$" + strSymbol
                If InStr(strSymbol, " ") > 0 Then strSecType = strSecType + "O"
                strTemp = ms.Name
                FixNullTermStr strTemp
                AddSymbolToList strSymbol, strSecType, ms.FileType, strTemp, "MS7", "F" & Trim(Str(ms.FileNum)) & ".DAT"
                iCount = iCount + 1
            End If
        Next 'iRec
        Close #fh
        
        ExtendCustomColumn
        fgSecurities.Redraw = lRedraw
    End If
    LoadMSmaster = iCount
    
ErrExit:
    Exit Function

ErrSection:
    LoadMSmaster = -1
    RaiseError "frmDataMan.LoadMSMaster", eGDRaiseError_Raise

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FixNullTermStr
'' Description: If the string is null terminated, cut off the string at the
''              null
'' Inputs:      The string to change
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FixNullTermStr(strChange As String)
On Error GoTo ErrSection:

    Dim i As Integer
    
    i = InStr(strChange, Chr(0))
    If i > 0 Then strChange = Left$(strChange, i - 1)
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDataMan.FixNullTermStr", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AreDigits
'' Description: Tells if the characters from the starting point to the
''              starting point + length are digits
'' Inputs:      String to check, Starting point, Length to check
'' Returns:     True if they are digits, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AreDigits%(strCheck$, ByVal iStart%, ByVal iLength%)
On Error GoTo ErrSection:
    
    Dim iIndex As Integer               ' Index for a for loop
    Dim iReturn As Integer              ' Return value from the function
    Dim strLetter As String             ' Current letter looking at

    iReturn = True
    
    ' Make adjustment to the length if necessary
    If iStart + iLength - 1 > Len(strCheck) Then
        iLength = Len(strCheck) - iStart + 1
    End If
    
    ' Walk through the string iLength characters from the start position
    For iIndex = iStart To iStart + iLength - 1
        strLetter = Mid$(strCheck, iIndex, 1)
        
        ' If the current character is not a digit, return false
        If strLetter < "0" Or strLetter > "9" Then
            iReturn = False
            Exit For
        End If
    Next 'i

    AreDigits = iReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmDataMan.AreDigits", eGDRaiseError_Raise

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddSymbolToList
'' Description: Adds a symbol to the list view control list
'' Inputs:      Symbol, Security Type, Name, and Format
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddSymbolToList(ByVal strSymbol$, ByVal strSecType$, ByVal strPeriod$, ByVal strName$, ByVal strFormat$, ByVal strFile$)
On Error GoTo ErrSection:
    
    Dim lRedraw As Long                 ' Current state of the redraw of the grid
    
    With fgSecurities
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        .Rows = .Rows + 1
        'If Left(strSymbol, 1) <> "*" Then strSymbol = "*" & strSymbol
        .TextMatrix(.Rows - 1, GDCol(eGDCol_Symbol)) = strSymbol
        Select Case UCase(Left(strPeriod, 1))
            Case "D"
                .TextMatrix(.Rows - 1, GDCol(eGDCol_Period)) = "Daily"
            Case "W"
                .TextMatrix(.Rows - 1, GDCol(eGDCol_Period)) = "Weekly"
            Case "M"
                .TextMatrix(.Rows - 1, GDCol(eGDCol_Period)) = "Monthly"
            Case "Q"
                .TextMatrix(.Rows - 1, GDCol(eGDCol_Period)) = "Quarterly"
            Case "Y"
                .TextMatrix(.Rows - 1, GDCol(eGDCol_Period)) = "Yearly"
            Case Else
                .TextMatrix(.Rows - 1, GDCol(eGDCol_Period)) = strPeriod & " Min"
        End Select
        .TextMatrix(.Rows - 1, GDCol(eGDCol_Name)) = strName
        .TextMatrix(.Rows - 1, GDCol(eGDCol_SecType)) = strSecType
        .TextMatrix(.Rows - 1, GDCol(eGDCol_Format)) = strFormat
        .TextMatrix(.Rows - 1, GDCol(eGDCol_File)) = strFile
        
        .Redraw = lRedraw
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDataMan.AddSymbolToList", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: If the user clicks on the X in the control menu, treat it like
''              a cancel
'' Inputs:      Whether or not to cancel the close, Mode of the unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode = 0 Then
        m.bOK = False
        Cancel = True
        Me.Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDataMan.Form.QueryUnload", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Resize
'' Description: Move and size the controls as the form gets resized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Resize()
On Error Resume Next

    If LimitFormSize(Me, fraHardDrive.Width * 2, fraHardDrive.Height + (fraHardDrive.Top * 2)) Then Exit Sub

    With fgSecurities
        .Move fraHardDrive.Width + (fraHardDrive.Left * 2), fraHardDrive.Top, _
                ScaleWidth - fraHardDrive.Width - (fraHardDrive.Left * 3), _
                ScaleHeight - fraButtons.Height - (fraHardDrive.Top * 3)
    End With

    With fraButtons
        .Move ((fgSecurities.Width - .Width) / 2) + fgSecurities.Left, _
                ScaleHeight - .Height - fraHardDrive.Top
    End With
    
    ExtendCustomColumn
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: When the form is unloaded, if the data manager is open, close it
'' Inputs:      Whether or not to cancel the unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    Dim strTemp As String               ' Temporary string for holding information

    strTemp = dirPathTree.List(dirPathTree.ListIndex) & "|"
    strTemp = strTemp & fgSecurities.TextMatrix(fgSecurities.Row, GDCol(eGDCol_Symbol)) & "|"
    strTemp = strTemp & CStr(chkCSI.Value) & "|"
    strTemp = strTemp & CStr(chkMetaStock.Value) & "|"
    strTemp = strTemp & CStr(chkGenTick.Value) & "|"
    strTemp = strTemp & CStr(chkGDB.Value) & "|"

    SetIniFileProperty "DataMan", GetFormPlacement(Me), "Placement", g.strIniFile
    SetIniFileProperty "LastSelection", strTemp, "DataMan", g.strIniFile
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDataMan.Form.Unload", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadGTFiles
'' Description: Load information about the GT files in a given path
'' Inputs:      Path to load GT files from
'' Returns:     Number of files found
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function LoadGTFiles(strPath As String) As Long
On Error GoTo ErrSection:

    Dim astrFiles() As String           ' GenTick files in a directory
    Dim lCount As Long                  ' Number of files matched
    Dim lIndex As Long                  ' Index into a for loop
    Dim strSymbol As String             ' Symbol from the GenTick file
    Dim strSecType As String            ' Security Type from the GenTick file
    
    ' Find out how many GenTick files there are in the path
    lCount = GetMatchingFiles(astrFiles, strPath, "*.GT")
    
    ' Walk through the files adding the information to the grid
    For lIndex = 1 To lCount
        strSymbol = UCase(Mid(astrFiles(lIndex), 1, InStr(UCase(astrFiles(lIndex)), ".GT") - 1))
        
        If Mid(strSymbol, 1, 1) = "$" Then
            strSecType = "I"
        ElseIf InStr(strSymbol, "_") > 0 Or InStr(strSymbol, "-") > 0 Then
            If Len(strSymbol) = 8 Then
                strSecType = "S"
            Else
                strSecType = "F"
            End If
        Else
            strSecType = "S"
        End If
        AddSymbolToList strSymbol, strSecType, "", "", "GT", astrFiles(lIndex)
    Next lIndex
    
    ExtendCustomColumn
    
    ' Return the number of files matched
    LoadGTFiles = lCount

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmDataMan.LoadGTFiles", eGDRaiseError_Raise

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadGDBFiles
'' Description: Load information about the GDB files in a given path
'' Inputs:      Path to load GDB files from
'' Returns:     Number of files found
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function LoadGDBFiles(strPath As String) As Long
On Error GoTo ErrSection:

    Dim astrFiles() As String           ' GenTick files in a directory
    Dim lCount As Long                  ' Number of files matched
    Dim lIndex As Long                  ' Index into a for loop
    Dim strSymbol As String             ' Symbol from the GenTick file
    Dim strSecType As String            ' Security Type from the GenTick file
    Dim strLine$, strPeriod$, strName$
    Dim aFlds As New cGdArray
    
    ' Find out how many GDB files there are in the path
    lCount = GetMatchingFiles(astrFiles, strPath, "*.GDB")
    
    ' Walk through the files adding the information to the grid
    For lIndex = 1 To lCount
        ' BarsSize, SymbolID, Symbol, BarPeriod, MinMoveInTicks, TickMove, TickValue, CrossoverTime, StartTime, EndTime, ExchTZ
        ' 11715  11936   IBM Daily   1   0.01    0.01    990 570 960 NY
        ' 1962-01-02 7.71    7.71    7.63    7.63    390000  -999999 -999999 -999999
        
        strLine = FileToString(AddSlash(strPath) & astrFiles(lIndex), 999, True)
        If InStr(strLine, vbTab) > 0 Then
            aFlds.SplitFields strLine, vbTab
        Else
            aFlds.SplitFields strLine
        End If
        strSymbol = Trim(aFlds(2))
        If Len(strSymbol) = 0 Then
            strSymbol = Trim(aFlds(0))
        End If
        strPeriod = aFlds(3)
        If Len(strPeriod) = 0 Then strPeriod = "Daily"
        strName = astrFiles(lIndex)
        
        If Mid(strSymbol, 1, 1) = "$" Then
            strSecType = "I"
        ElseIf InStr(strSymbol, "_") > 0 Or InStr(strSymbol, "-") > 0 Then
            If Len(strSymbol) = 8 Then
                strSecType = "S"
            Else
                strSecType = "F"
            End If
        Else
            strSecType = "S"
        End If
        AddSymbolToList strSymbol, strSecType, strPeriod, strName, "GDB", astrFiles(lIndex)
    Next lIndex
    
    ExtendCustomColumn
    
    ' Return the number of files matched
    LoadGDBFiles = lCount

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmDataMan.LoadGDBFiles", eGDRaiseError_Raise

End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitGrid
'' Description: Initialize the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitGrid()
On Error GoTo ErrSection:

    Dim lRedraw As Long                 ' Current state of the grid's redraw

    With fgSecurities
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        .AllowBigSelection = False
        .AllowSelection = False
        .AllowUserResizing = flexResizeColumns
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .Editable = flexEDNone
        .ExplorerBar = flexExSortShow
        .ExtendLastCol = True
        .SelectionMode = flexSelectionListBox
        .ScrollTrack = True
        .ScrollTips = True
        .SheetBorder = RGB(128, 128, 128)
        
        .AutoSearch = flexSearchFromTop
        .AutoSearchDelay = 2
        
        .FixedRows = 1
        .FixedCols = 0
        .Rows = 1
        .Cols = GDCol(eGDCol_NumCols)
        
        .TextMatrix(0, GDCol(eGDCol_Symbol)) = "Symbol"
        .TextMatrix(0, GDCol(eGDCol_Period)) = "Period"
        .TextMatrix(0, GDCol(eGDCol_Name)) = "Name"
        .TextMatrix(0, GDCol(eGDCol_SecType)) = "Sec Type"
        .TextMatrix(0, GDCol(eGDCol_Format)) = "Format"
        
        .ColHidden(GDCol(eGDCol_SecType)) = True
        .ColHidden(GDCol(eGDCol_File)) = True
        
        .Redraw = lRedraw
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDataMan.InitGrid", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ExtendCustomColumn
'' Description: Adjust all column widths to accomodate the custom "extend column"
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ExtendCustomColumn(Optional ByVal nResizeCol As Long = -1)
On Error GoTo ErrSection:
    
    Dim i&, nTotal&, nDiff&
    
    With fgSecurities
        ' if column being resized is after the extended column,
        ' then change the width of the next visible column instead
        If nResizeCol >= kExtendedCol Then
            .Redraw = flexRDNone
            nDiff = .ColWidth(nResizeCol) - m.nPrevColWidth
            For i = nResizeCol + 1 To .Cols - 1
                If Not .ColHidden(i) Then
                    .ColWidth(i) = .ColWidth(i) - nDiff
                    Exit For
                End If
            Next
            m.nPrevColWidth = 0
        End If
        
        ' size the custom extended column in order to fill the client width
        .ColHidden(kExtendedCol) = True
        .Redraw = flexRDBuffered '(must do this so .ClientWidth will be correct)
        .Redraw = flexRDNone
        nTotal = 0
        For i = 0 To .Cols - 1
            If Not .ColHidden(i) Then
                nTotal = nTotal + .ColWidth(i)
            End If
        Next
        nTotal = .ClientWidth - nTotal
        If nTotal > 0 Then .ColWidth(kExtendedCol) = nTotal
        .ColHidden(kExtendedCol) = False
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDataMan.ExtendCustomColumn", eGDRaiseError_Raise
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Initialize and show the form
'' Inputs:      Pipe-delimited string of where to start
'' Returns:     Pipe-delimited string of the selection (or blank string if
''              Cancelled)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(Optional ByVal strStart As String = "") As String
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim SymInf As vbSymbolInfo          ' Information about a symbol
    Dim strInfo As String               ' Information out of the ini file
    Dim strSymbol As String             ' Symbol to attemt to select

    Screen.MousePointer = vbHourglass

    ' Initialize the grid
    InitGrid
      
    ' Initialize some variables
    m.lTreeIndex = -1
    m.lColumn = -1
    m.lSortedCol = GDCol(eGDCol_Symbol)
    m.lSortOrder = flexSortStringAscending
    dirPathTree.Path = "C:\"
    
    ' Get the default information out of the ini file
    strInfo = GetIniFileProperty("LastSelection", "", "DataMan", g.strIniFile)
    
    ' Set the default path
    strSymbol = ""
    If DirExist("C:\GD") Then
        dirPathTree.Path = "C:\GD\"
    Else
        dirPathTree.Path = "C:\"
    End If
    If strInfo <> "" Then
        If DirExist(Parse(strInfo, "|", 1)) Then
            dirPathTree.Path = Parse(strInfo, "|", 1)
        End If
        
        strSymbol = Parse(strInfo, "|", 2)
        chkCSI.Value = Val(Parse(strInfo, "|", 3))
        chkMetaStock.Value = Val(Parse(strInfo, "|", 4))
        chkGenTick.Value = Val(Parse(strInfo, "|", 5))
        If Parse(strInfo, "|", 6) = "0" Then
            chkGDB.Value = 0
        Else
            chkGDB.Value = 1
        End If
    End If
    If strStart <> "" Then
        If DirExist(Parse(strStart, "|", 1)) Then
            dirPathTree.Path = Parse(strStart, "|", 1)
        End If
        strSymbol = Parse(strStart, "|", 2)
    End If
    drvDrive.Drive = dirPathTree.Path
    
    PathChanged True
    fgSecurities.ColHidden(GDCol(eGDCol_Format)) = False
    
    ' Select the given symbol (if possible)
    If strSymbol <> "" Then SelectSymbol strSymbol, True
    
    Screen.MousePointer = vbDefault
    
    ' Only show the market button if in the IDE...
    cmdMarket.Visible = IsIDE
    
    ' Show the form
    ShowForm Me, True
    
    ' If the user hit OK, return the appropriate symbol information
    If m.bOK Then
        With fgSecurities
            ShowMe = "*" & .TextMatrix(.Row, GDCol(eGDCol_Symbol)) & "|"
            Select Case .TextMatrix(.Row, GDCol(eGDCol_Format))
                Case "CSI"
                    'ShowMe = ShowMe & AddSlash(dirPathTree.List(dirPathTree.ListIndex)) & "QMASTER."
                Case "MS7"
                    'ShowMe = ShowMe & AddSlash(dirPathTree.List(dirPathTree.ListIndex)) & "MASTER."
                Case Else
                    'ShowMe = ShowMe & AddSlash(dirPathTree.List(dirPathTree.ListIndex))
            End Select
ShowMe = ShowMe & AddSlash(dirPathTree.List(dirPathTree.ListIndex)) & Trim(.TextMatrix(.Row, GDCol(eGDCol_File)))
            ShowMe = ShowMe & "|" & Trim(.TextMatrix(.Row, GDCol(eGDCol_Name))) _
                & "|" & .TextMatrix(.Row, GDCol(eGDCol_SecType)) _
                & "|" & .TextMatrix(.Row, GDCol(eGDCol_Format)) & "|D"
        End With
    Else
        ShowMe = ""
    End If
    
ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmDataMan.ShowMe", eGDRaiseError_Raise

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SelectSymbol
'' Description: Select a symbol in the grid and try to center it
'' Inputs:      Symbol to select, Whether or not to select the first entry in
''              the grid if the given symbol was not found
'' Returns:     True if symbol was found, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SelectSymbol(ByVal strSymbol As String, ByVal bSelectFirstOnFailure As Boolean) As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lRedraw As Long                 ' Current state of the grid redraw

    With fgSecurities
        lRedraw = .Redraw
        .Redraw = flexRDNone
        For lIndex = .FixedRows To .Rows - 1
            If UCase(Left(.TextMatrix(lIndex, GDCol(eGDCol_Symbol)), Len(strSymbol))) = UCase(strSymbol) Then
                .Row = lIndex
                .RowSel = lIndex
                .ShowCell lIndex, GDCol(eGDCol_Symbol)
                CenterSelection
                SelectSymbol = True
                Exit For
            End If
        Next lIndex
        
        If bSelectFirstOnFailure And SelectSymbol = False Then
            If .Rows > .FixedRows Then
                .Row = .FixedRows
                .RowSel = .FixedRows
                .ShowCell .FixedRows, GDCol(eGDCol_Symbol)
            End If
        End If
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmDataMan.SelectSymbol", eGDRaiseError_Raise
    Resume ErrExit

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CenterSelection
'' Description: Center the selection on the grid vertically if possible
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CenterSelection()
On Error GoTo ErrSection:

    Dim lRedraw As Long                 ' Current state of the redraw
    Dim lTopRow As Long                 ' New Top Row

    With fgSecurities
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        lTopRow = .SelectedRow(0) - ((.BottomRow - .TopRow) / 2)
        If lTopRow >= .FixedRows Then .TopRow = lTopRow
        
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDataMan.CenterSelection", eGDRaiseError_Raise
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DoCheck
'' Description: Check to see if market information exists for the symbol
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DoCheck()
On Error GoTo ErrSection:

    Dim Bars As New cGdBars             ' Temporary bars structure
    Dim strSymbol As String             ' Symbol to attemt to select
    Dim strDesc As String               ' Description for the symbol

    strSymbol = "*" & fgSecurities.TextMatrix(fgSecurities.Row, GDCol(eGDCol_Symbol))
    strDesc = Trim(fgSecurities.TextMatrix(fgSecurities.Row, GDCol(eGDCol_Name)))
    
    SetBarProperties Bars, strSymbol
    
    If Bars.Prop(eBARS_TickMove) = 0 Or Bars.Prop(eBARS_TickValue) = 0 Then
        If frmMarkets.ShowMe(strSymbol, strDesc) = False Then GoTo ErrExit
    End If
    
    m.bOK = True
    Me.Hide
    
ErrExit:
    Set Bars = Nothing
    Exit Sub
    
ErrSection:
    Set Bars = Nothing
    RaiseError "frmDataMan.DoCheck", eGDRaiseError_Raise
    
End Sub

