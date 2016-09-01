VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmTradeSenseOrderGroups 
   Caption         =   "Form1"
   ClientHeight    =   4350
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   5055
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniTextBoxXP txtPreview 
      Height          =   555
      Left            =   120
      TabIndex        =   1
      Top             =   1860
      Width           =   2955
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmTradeSenseOrderGroups.frx":0000
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
      MultiLine       =   -1  'True
      Alignment       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      TrapTab         =   0   'False
      EnableContextMenu=   -1  'True
      RaiseChangeEvent=   -1  'True
      Tip             =   "frmTradeSenseOrderGroups.frx":0020
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTradeSenseOrderGroups.frx":0040
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   4155
      Left            =   3240
      TabIndex        =   2
      Top             =   120
      Width           =   1635
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
      Caption         =   "frmTradeSenseOrderGroups.frx":005C
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmTradeSenseOrderGroups.frx":0088
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTradeSenseOrderGroups.frx":00A8
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniFrameWL fraFavorites 
         Height          =   735
         Left            =   15
         TabIndex        =   9
         Top             =   3360
         Width           =   1605
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
         Caption         =   "frmTradeSenseOrderGroups.frx":00C4
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmTradeSenseOrderGroups.frx":00F6
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTradeSenseOrderGroups.frx":0116
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniButtonImageXP cmdTSO1 
            Height          =   375
            Left            =   105
            TabIndex        =   13
            Top             =   240
            Width           =   305
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
            Caption         =   "frmTradeSenseOrderGroups.frx":0132
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmTradeSenseOrderGroups.frx":0154
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmTradeSenseOrderGroups.frx":0174
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdTSO2 
            Height          =   375
            Left            =   465
            TabIndex        =   12
            Top             =   240
            Width           =   305
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
            Caption         =   "frmTradeSenseOrderGroups.frx":0190
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmTradeSenseOrderGroups.frx":01B2
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmTradeSenseOrderGroups.frx":01D2
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdTSO3 
            Height          =   375
            Left            =   825
            TabIndex        =   11
            Top             =   240
            Width           =   305
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
            Caption         =   "frmTradeSenseOrderGroups.frx":01EE
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmTradeSenseOrderGroups.frx":0210
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmTradeSenseOrderGroups.frx":0230
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdTSO4 
            Height          =   375
            Left            =   1185
            TabIndex        =   10
            Top             =   240
            Width           =   305
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
            Caption         =   "frmTradeSenseOrderGroups.frx":024C
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmTradeSenseOrderGroups.frx":026E
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmTradeSenseOrderGroups.frx":028E
            RightToLeft     =   0   'False
         End
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdClose 
         Height          =   495
         Left            =   210
         TabIndex        =   8
         Top             =   2700
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
         Caption         =   "frmTradeSenseOrderGroups.frx":02AA
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTradeSenseOrderGroups.frx":02D6
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTradeSenseOrderGroups.frx":02F6
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdPark 
         Height          =   495
         Left            =   210
         TabIndex        =   7
         Top             =   2160
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
         Caption         =   "frmTradeSenseOrderGroups.frx":0312
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTradeSenseOrderGroups.frx":033C
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTradeSenseOrderGroups.frx":035C
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdSubmit 
         Height          =   495
         Left            =   210
         TabIndex        =   6
         Top             =   1620
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
         Caption         =   "frmTradeSenseOrderGroups.frx":0378
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTradeSenseOrderGroups.frx":03A6
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTradeSenseOrderGroups.frx":03C6
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdDelete 
         Height          =   495
         Left            =   210
         TabIndex        =   5
         Top             =   1080
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
         Caption         =   "frmTradeSenseOrderGroups.frx":03E2
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTradeSenseOrderGroups.frx":0410
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTradeSenseOrderGroups.frx":0430
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdEdit 
         Height          =   495
         Left            =   210
         TabIndex        =   4
         Top             =   540
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
         Caption         =   "frmTradeSenseOrderGroups.frx":044C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTradeSenseOrderGroups.frx":0476
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTradeSenseOrderGroups.frx":0496
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdNew 
         Height          =   495
         Left            =   210
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
         Caption         =   "frmTradeSenseOrderGroups.frx":04B2
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTradeSenseOrderGroups.frx":04DA
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTradeSenseOrderGroups.frx":04FA
         RightToLeft     =   0   'False
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid fgGroups 
      Height          =   1515
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2955
      _cx             =   5212
      _cy             =   2672
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
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUp"
      Begin VB.Menu mnuNew 
         Caption         =   "New"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Edit"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnuSubmit 
         Caption         =   "Submit"
      End
      Begin VB.Menu mnuPark 
         Caption         =   "Park"
      End
   End
End
Attribute VB_Name = "frmTradeSenseOrderGroups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmTradeSenseOrderGroups.frm
'' Description: Form that handles Trade Sense order group management
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 06/15/2010   DAJ         Changed the form icon
'' 07/15/2010   DAJ         Added capabilities for inputs
'' 08/11/2010   DAJ         Sorted grid and only reload collection if save as/rename
'' 08/23/2010   DAJ         Added required module flag for TradeSense orders/groups
'' 09/16/2010   DAJ         Don't add new group to collection if not saved
'' 09/29/2010   DAJ         Removed reference to global order confirmation flag
'' 09/29/2010   DAJ         After editing, reload group to get new information
'' 10/05/2010   DAJ         Reload group before editing/submitting/parking
'' 10/20/2010   DAJ         Continuous Loop
'' 11/16/2010   DAJ         Added allow manual submission flag
'' 03/31/2011   DAJ         Fix for deleting a provided order group (#6232)
'' 05/18/2011   DAJ         Added custom start/stop time for Market1
'' 09/26/2011   DAJ         Allow manual submission of Daniel Code TSOG with enablement
'' 06/08/2012   MJM         Added ability to assign a group to a favorite button
'' 10/03/2012   DAJ         Lot size for forex symbols in TradeSense order groups
'' 04/16/2014   DAJ         Fix for submitting TradeSense order group via favorites ( Pete Laverde )
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Enum eGDCols
    eGDCol_Favorite = 0
    eGDCol_Name
    eGDCol_NumCols
End Enum

Private Type mPrivate
    TsoGroups As cTradeSenseOrderGroups ' Trade Sense order groups
    
    bUseGivenInfo As Boolean            ' Use the info given in the ShowMe?
    strSymbol As String                 ' Symbol to use
    lAccountID As Long                  ' Account to use
    lQuantity As Long                   ' Quantity to use
    astrInputs As cGdArray              ' Array of input information to use
End Type
Private m As mPrivate

Private Function GDCol(ByVal nCol As eGDCols) As Long
    GDCol = nCol
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Load controls and show the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowMe()
On Error GoTo ErrSection:

    Set m.TsoGroups = New cTradeSenseOrderGroups
    m.TsoGroups.Load
    
    m.bUseGivenInfo = False
    m.strSymbol = ""
    m.lAccountID = 0&
    m.lQuantity = kNullData

    InitGrid
    LoadGrid
    
    EnableControls

    ShowForm Me, eForm_Modal, frmMain, , ALT_GRID_ROW_COLOR

ErrExit:
    Unload Me
    Exit Sub
    
ErrSection:
    Unload Me
    RaiseError "frmTradeSenseOrderGroups.ShowMe"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMeWithInfo
'' Description: Load controls and show the form
'' Inputs:      Symbol, Account, Quantity
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowMeWithInfo(ByVal strSymbol As String, ByVal lAccountID As Long, ByVal lQuantity As Long)
On Error GoTo ErrSection:

    Set m.TsoGroups = New cTradeSenseOrderGroups
    m.TsoGroups.Load
    
    m.bUseGivenInfo = True
    m.strSymbol = strSymbol
    m.lAccountID = lAccountID
    m.lQuantity = lQuantity

    InitGrid
    LoadGrid
    
    EnableControls

    ShowForm Me, eForm_Modal, frmMain

ErrExit:
    Unload Me
    Exit Sub
    
ErrSection:
    Unload Me
    RaiseError "frmTradeSenseOrderGroups.ShowMeWithInfo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleTradeSenseWrapper
'' Description: A wrapper to allow charts to submit order groups via favorite button
'' Inputs:      Order group object, symbol, trade account ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub HandleTradeSenseWrapper(tsoGrp As cTradeSenseOrderGroup, ByVal strSymbol$, ByVal lTradeAccountID&)
On Error GoTo ErrSection:
    
    If Not tsoGrp Is Nothing Then
        m.strSymbol = strSymbol
        m.lAccountID = lTradeAccountID
        HandleTradeSenseGroup True, tsoGrp
    End If
    
ErrExit:
    Unload Me
    Exit Sub
    
ErrSection:
    Unload Me
    RaiseError "frmTradeSenseOrderGroups.HandleTradeSenseWrapper"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdClose_Click
'' Description: Allow the ShowMe routine to unload the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdClose_Click()
On Error GoTo ErrSection:

    Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroups.cmdClose_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdDelete_Click
'' Description: Allow the user to delete a Trade Sense order group
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdDelete_Click()
On Error GoTo ErrSection:

    DeleteTradeSenseGroup

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroups.cmdDelete_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdEdit_Click
'' Description: Allow the user to edit a Trade Sense order group
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdEdit_Click()
On Error GoTo ErrSection:

    EditTradeSenseGroup

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroups.cmdEdit_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdNew_Click
'' Description: Allow the user to create a new Trade Sense order group
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdNew_Click()
On Error GoTo ErrSection:

    NewTradeSenseGroup

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroups.cmdNew_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdPark_Click
'' Description: Allow the user to park a Trade Sense order group
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdPark_Click()
On Error GoTo ErrSection:

    HandleTradeSenseGroup False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroups.cmdPark_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdSubmit_Click
'' Description: Allow the user to submit a Trade Sense order group
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdSubmit_Click()
On Error GoTo ErrSection:

    HandleTradeSenseGroup True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroups.cmdSubmit_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdTSO1_MouseUp
'' Description: Allow the user to assign or clear an order group in global favorite group array index 0
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdTSO1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    If Button = vbRightButton Then
        ClearFavorites 0
    Else
        AssignFavorites 0
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroups.cmdTSO1_MouseUp"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdTSO2_MouseUp
'' Description: Allow the user to assign or clear an order group in global favorite group array index 1
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdTSO2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    If Button = vbRightButton Then
        ClearFavorites 1
    Else
        AssignFavorites 1
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroups.cmdTSO2_MouseUp"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdTSO3_MouseUp
'' Description: Allow the user to assign or clear an order group in global favorite group array index 2
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdTSO3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    If Button = vbRightButton Then
        ClearFavorites 2
    Else
        AssignFavorites 2
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroups.cmdTSO3_MouseUp"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdTSO4_MouseUp
'' Description: Allow the user to assign or clear an order group in global favorite group array index 3
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdTSO4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    If Button = vbRightButton Then
        ClearFavorites 3
    Else
        AssignFavorites 3
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroups.cmdTSO4_MouseUp"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgGroups_AfterRowColChange
'' Description: After the user changes cells, enable/disable controls
'' Inputs:      Old Row, Old Column, New Row, New Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgGroups_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    Dim tsoGrp As cTradeSenseOrderGroup ' Trade Sense order group object
    
    Set tsoGrp = SelectedGroup
    If Not tsoGrp Is Nothing Then
        txtPreview.Text = tsoGrp.Description
    Else
        txtPreview.Text = ""
    End If

    EnableControls
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroups.fgGroups_AfterRowColChange"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgGroups_BeforeMouseDown
'' Description: Show popup menu on a right click
'' Inputs:      Mouse Button, Shift/Ctrl/Alt Status, Mouse Location, Cancel?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgGroups_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:

    fgGroups.Row = fgGroups.MouseRow
    
    If Button = vbRightButton Then
        PopupMenu mnuPopUp
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroups.fgGroups_BeforeMouseDown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgGroups_DblClick
'' Description: Allow the user to edit a custom group upon double click
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgGroups_DblClick()
On Error GoTo ErrSection:

    fgGroups.Row = fgGroups.MouseRow
    
    If ValidRowSelected Then
        EditTradeSenseGroup
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroups.fgGroups_DblClick"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgGroups_KeyDown
'' Description: Allow user to create a new or delete an existing group
'' Inputs:      Key Code, Shift/Ctrl/Alt status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgGroups_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If KeyCode = vbKeyInsert Then
        NewTradeSenseGroup
    ElseIf KeyCode = vbKeyDelete Then
        If ValidRowSelected Then
            DeleteTradeSenseGroup
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroups.fgGroups_KeyDown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgGroups_KeyPress
'' Description: Allow the user to edit an existing group
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgGroups_KeyPress(KeyAscii As Integer)
On Error GoTo ErrSection:

    If KeyAscii = vbKeyReturn Then
        If ValidRowSelected Then
            EditTradeSenseGroup
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroups.fgGroups_KeyPress"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Initialize form when it is loaded
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Dim strPlacement As String          ' Form placement from the INI file
    
    g.Styler.StyleForm Me
    
    strPlacement = GetIniFileProperty("frmTradeSenseOrderGroups", "", "Placement", g.strIniFile)
    If Len(strPlacement) = 0 Then
        CenterTheForm Me
    Else
        SetFormPlacement Me, strPlacement
    End If
    
    txtPreview.Locked = True
    txtPreview.BackColor = &H80000000
    
    Caption = "Trade Sense Order Groups"
    Icon = Picture16(ToolbarIcon("kTradeSenseOrders"))
    
    Set m.astrInputs = New cGdArray
    
    mnuPopUp.Visible = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroups.Form_Load"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: If the user clicks on the X, allow ShowMe to unload the form
'' Inputs:      Cancel the Unload?, Mode of the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode <> vbFormCode Then
        Cancel = True
        Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroups.Form_QueryUnload"
    
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

    Dim lMinScaleWidth As Long          ' Minimum scale width for the form
    Dim lMinScaleHeight As Long         ' Minimum scale height for the form
    Dim lVertSpace As Long              ' Vertical space between controls
    Dim lHorzSpace As Long              ' Horizontal space between controls
    
    lVertSpace = 120
    lHorzSpace = 120
    
    lMinScaleWidth = (fraButtons.Width * 3) + (lHorzSpace * 3)
    lMinScaleHeight = fraButtons.Height + (lVertSpace * 2)
    
    If LimitFormSize(Me, lMinScaleWidth, lMinScaleHeight) = False Then
        With fraButtons
            .Move ScaleWidth - .Width - lHorzSpace, lVertSpace
        End With
        
        With txtPreview
            .Move lHorzSpace, ScaleHeight - .Height - lVertSpace, ScaleWidth - fraButtons.Width - (lHorzSpace * 3)
        End With
        
        With fgGroups
            .Move lHorzSpace, fraButtons.Top, txtPreview.Width, ScaleHeight - txtPreview.Height - (lVertSpace * 3)
        End With
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: Clean up when the form is unloaded
'' Inputs:      Cancel the Unload?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    SetIniFileProperty "frmTradeSenseOrderGroups", GetFormPlacement(Me), "Placement", g.strIniFile
    Set m.TsoGroups = Nothing
    Set m.astrInputs = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroups.Form_Unload"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuDelete_Click
'' Description: Allow the user to delete a Trade Sense order group
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuDelete_Click()
On Error GoTo ErrSection:

    DeleteTradeSenseGroup

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroups.mnuDelete_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuEdit_Click
'' Description: Allow the user to edit a Trade Sense order group
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuEdit_Click()
On Error GoTo ErrSection:

    EditTradeSenseGroup

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroups.mnuEdit_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuNew_Click
'' Description: Allow the user to create a new Trade Sense order group
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuNew_Click()
On Error GoTo ErrSection:

    NewTradeSenseGroup

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroups.mnuNew_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuPark_Click
'' Description: Allow the user to park a Trade Sense order group
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuPark_Click()
On Error GoTo ErrSection:

    HandleTradeSenseGroup False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroups.mnuPark_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuSubmit_Click
'' Description: Allow the user to submit a Trade Sense order group
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuSubmit_Click()
On Error GoTo ErrSection:

    HandleTradeSenseGroup True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroups.mnuSubmit_Click"
    
End Sub

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

    With fgGroups
        .Redraw = flexRDNone
        
        .Rows = 0
        .FixedRows = 0
        .Cols = GDCol(eGDCol_NumCols)
        .FixedCols = 0
        
        SetupGrid fgGroups, eGridMode_List
        
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .ColWidth(GDCol(eGDCol_Favorite)) = 305
        .ExtendLastCol = True
        
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroups.InitGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadGrid
'' Description: Load the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadGrid()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim tsoGrp As cTradeSenseOrderGroup ' Trade Sense order group object
    Dim strSelected As String           ' Item selected
    Dim strSelectedID As String         ' ID of the selected item
    Dim lRowToSelect As Long            ' Row to select
    Dim lFavoriteIndex As Long          ' Index into the TradeSense order group favorites
    
    With fgGroups
        .Redraw = flexRDNone
        
        strSelectedID = ""
        lRowToSelect = -1&
        
        If (.Row >= .FixedRows) And (.Row < .Rows) Then
            If TypeOf .RowData(.Row) Is cTradeSenseOrderGroup Then
                Set tsoGrp = .RowData(.Row)
                strSelectedID = tsoGrp.ID
            End If
        End If
        
        .Rows = .FixedRows
        For lIndex = 1 To m.TsoGroups.Count
            Set tsoGrp = m.TsoGroups(lIndex)
            If HasModule(tsoGrp.RequiredMod) Then
                .Rows = .Rows + 1
                .RowData(.Rows - 1) = tsoGrp
                .TextMatrix(.Rows - 1, GDCol(eGDCol_Name)) = tsoGrp.Name
                
                lFavoriteIndex = TsogFavorite(tsoGrp)
                If lFavoriteIndex = -1& Then
                    .TextMatrix(.Rows - 1, GDCol(eGDCol_Favorite)) = ""
                Else
                    .TextMatrix(.Rows - 1, GDCol(eGDCol_Favorite)) = Str(lFavoriteIndex + 1)
                End If
            End If
        Next lIndex
        
        .Col = GDCol(eGDCol_Name)
        .Sort = flexSortGenericAscending
        
        If Len(strSelectedID) > 0 Then
            For lIndex = .FixedRows To .Rows - 1
                Set tsoGrp = .RowData(lIndex)
                If tsoGrp.ID = strSelectedID Then
                    lRowToSelect = lIndex
                    Exit For
                End If
            Next lIndex
        End If
        
        ' Select the first row if there are rows in the grid...
        If (Len(strSelectedID) > 0) And (lRowToSelect > .FixedRows) Then
            .Row = lRowToSelect
        ElseIf .Rows > .FixedRows Then
            .Row = .FixedRows
        End If
        
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroups.LoadGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EnableControls
'' Description: Enable/Disable the controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EnableControls()
On Error GoTo ErrSection:

    Dim bValidRowSelected As Boolean    ' Is there a valid row selected in the grid?
    Dim tsoGrp As cTradeSenseOrderGroup ' Currently selected group
    Dim bAllowManual As Boolean         ' Allow manual submission of the group?
    Dim bAllowEditOrDelete As Boolean   ' Allow edit or delete of the group?
    
    bAllowManual = True
    bValidRowSelected = ValidRowSelected
    Set tsoGrp = SelectedGroup
    If Not tsoGrp Is Nothing Then
        bAllowManual = tsoGrp.AllowManualSubmission
        If bAllowManual = False Then
            bAllowManual = AllowDanielCodeManual(tsoGrp)
        End If
    End If
    bAllowEditOrDelete = AllowEditOrDelete(tsoGrp)
    
    Enable cmdEdit, bValidRowSelected And bAllowEditOrDelete
    Enable mnuEdit, bValidRowSelected And bAllowEditOrDelete
    Enable cmdDelete, bValidRowSelected And bAllowEditOrDelete
    Enable mnuDelete, bValidRowSelected And bAllowEditOrDelete
    Enable cmdSubmit, bValidRowSelected And bAllowManual
    Enable mnuSubmit, bValidRowSelected And bAllowManual
    Enable cmdPark, bValidRowSelected And bAllowManual
    Enable mnuPark, bValidRowSelected And bAllowManual

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroups.EnableControls"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ValidRowSelected
'' Description: Is a valid row selected in the grid?
'' Inputs:      None
'' Returns:     True if valid row selected, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ValidRowSelected() As Boolean
On Error GoTo ErrSection:

    ValidRowSelected = ((fgGroups.Row >= fgGroups.FixedRows) And (fgGroups.Row < fgGroups.Rows))

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroups.ValidRowSelected"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SelectedGroup
'' Description: Selected group in the grid
'' Inputs:      None
'' Returns:     Selected Group (Nothing if not valid)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SelectedGroup() As cTradeSenseOrderGroup
On Error GoTo ErrSection:

    Dim tsoGrp As cTradeSenseOrderGroup ' Return value for the function
    
    Set tsoGrp = Nothing
    If ValidRowSelected Then
        If TypeOf fgGroups.RowData(fgGroups.Row) Is cTradeSenseOrderGroup Then
            Set tsoGrp = fgGroups.RowData(fgGroups.Row)
        End If
    End If
    
    Set SelectedGroup = tsoGrp

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroups.SelectedGroup"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowGroupForm
'' Description: Show the Trade Sense order group form
'' Inputs:      Trade Sense Order Group
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ShowGroupForm(tsoGroup As cTradeSenseOrderGroup)
On Error GoTo ErrSection:

    Dim bReload As Boolean              ' Do we need to reload things?
    
    tsoGroup.Reload
    frmTradeSenseOrderGroup.ShowMe tsoGroup, bReload
    
    If bReload Then
        m.TsoGroups.Load
    ElseIf Len(tsoGroup.Name) > 0 Then
        tsoGroup.Reload
        m.TsoGroups(tsoGroup.Name) = tsoGroup
    End If
    
    LoadGrid

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroups.ShowGroupForm"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NewTradeSenseGroup
'' Description: Allow the user to create a new Trade Sense group
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub NewTradeSenseGroup()
On Error GoTo ErrSection:

    Dim tsoGrp As cTradeSenseOrderGroup ' Trade Sense order group object
    
    Set tsoGrp = New cTradeSenseOrderGroup
    ShowGroupForm tsoGrp

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroups.NewTradeSenseGroup"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EditTradeSenseGroup
'' Description: Allow the user to edit an existing Trade Sense group
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EditTradeSenseGroup()
On Error GoTo ErrSection:

    Dim tsoGrp As cTradeSenseOrderGroup ' Trade Sense order group object
    
    Set tsoGrp = SelectedGroup
    If Not tsoGrp Is Nothing Then
        If AllowEditOrDelete(tsoGrp) Then
            ShowGroupForm tsoGrp
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroups.EditTradeSenseGroup"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DeleteTradeSenseGroup
'' Description: Allow the user to delete an existing Trade Sense group
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DeleteTradeSenseGroup()
On Error GoTo ErrSection:

    Dim tsoGrp As cTradeSenseOrderGroup ' Trade Sense order group object
    
    Set tsoGrp = SelectedGroup
    If Not tsoGrp Is Nothing Then
        If AllowEditOrDelete(tsoGrp) Then
            If InfBox("Are you sure that you want to delete '" & tsoGrp.Name & "'?", "?", "+Yes|-No", "Delete Confirmation") = "Y" Then
                ClearFavorites TsogFavorite(tsoGrp)
            
                KillFile tsoGrp.FileName
                m.TsoGroups.Remove tsoGrp.Name
                LoadGrid
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroups.DeleteTradeSenseGroup"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleTradeSenseGroup
'' Description: Allow the user to either submit or park an existing TradeSense group
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HandleTradeSenseGroup(ByVal bSubmit As Boolean, Optional tsoGrpSelected As cTradeSenseOrderGroup = Nothing)
On Error GoTo ErrSection:

    Dim tsoGrp As cTradeSenseOrderGroup ' Trade Sense order group object
    Dim lLotSize As Long                ' Lot size
    Dim Inputs As cTradeSenseOrderInputs ' Inputs collection
    Dim bContinue As Boolean            ' Submit or park the order group?
    Dim bLoop As Boolean                ' Continuously loop?
    Dim dLoopExp As Double              ' Loop expiration
    Dim dCustomStartTime As Double      ' Custom start time
    Dim dCustomStopTime As Double       ' Custom stop time
    
    If tsoGrpSelected Is Nothing Then
        Set tsoGrp = SelectedGroup
    Else
        Set tsoGrp = tsoGrpSelected
    End If
    
    If Not tsoGrp Is Nothing Then
        bContinue = True
        tsoGrp.Reload
        
        dCustomStartTime = kNullData
        dCustomStopTime = kNullData
        
        If (m.bUseGivenInfo = False) Or (InfoInvalid = True) Then
            Set Inputs = New cTradeSenseOrderInputs
            Inputs.ForGroups = True
            
            bContinue = frmActiveTsOrderGroup.ShowMe(tsoGrp, m.strSymbol, m.lAccountID, m.lQuantity, lLotSize, Inputs, bLoop, dLoopExp, dCustomStartTime, dCustomStopTime)
        End If
        
        If bContinue Then
            tsoGrp.SetInputValues Inputs
            If bSubmit Then
                g.TsoGroups.SubmitGroup tsoGrp, m.strSymbol, m.lAccountID, m.lQuantity, bLoop, dLoopExp, dCustomStartTime, dCustomStopTime, lLotSize
            Else
                g.TsoGroups.ParkGroup tsoGrp, m.strSymbol, m.lAccountID, m.lQuantity, bLoop, dLoopExp, dCustomStartTime, dCustomStopTime, lLotSize
            End If
            
            Hide
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroups.HandleTradeSenseGroup"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InfoInvalid
'' Description: Was the info given in the ShowMeWithInfo invalid?
'' Inputs:      None
'' Returns:     True if Invalid, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function InfoInvalid() As Boolean
On Error GoTo ErrSection:

    InfoInvalid = (Len(m.strSymbol) = 0) Or (m.lAccountID <= 0&) Or (m.lQuantity <= 0&)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroups.InfoInvalid"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AllowEditOrDelete
'' Description: Should we allow editing or deleting the given order group?
'' Inputs:      TradeSense Order Group
'' Returns:     True if allow, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AllowEditOrDelete(tsoGrp As cTradeSenseOrderGroup) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    If Not tsoGrp Is Nothing Then
        bReturn = (tsoGrp.Custom Or FileExist("C:\Common\Files.EXE"))
    End If
    
    AllowEditOrDelete = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroups.AllowEditOrDelete"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AllowDanielCodeManual
'' Description: Should we allow a Daniel Code order group to be manually submitted?
'' Inputs:      TradeSense Order Group
'' Returns:     True if allow, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AllowDanielCodeManual(tsoGrp As cTradeSenseOrderGroup) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    If InStr("," & UCase(tsoGrp.RequiredMod) & ",", ",DCPLUS,") <> 0 Then
        If HasModule("DCPLUS") And HasModule("DCMANUAL") Then
            bReturn = True
        End If
    End If
    
    AllowDanielCodeManual = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroups.AllowDanielCodeManual"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AssignFavorites
'' Description: Assign the selected TSO group as a global Favorite TSO group
'' Inputs:      Index number for global array holding Favorite TSO groups (there are only 4 items in global array)
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AssignFavorites(ByVal lFavoriteIndex As Long)
On Error GoTo ErrSection:

    Dim bClearOnly As Boolean           ' Clear the assignment only?
    Dim tsoGrp As cTradeSenseOrderGroup ' Selected TradeSense order group
    
    If (lFavoriteIndex >= 0) And (lFavoriteIndex < 4) Then
        With fgGroups
            If ValidGridRow(fgGroups) Then
                If TypeOf .RowData(.Row) Is cTradeSenseOrderGroup Then
                    Set tsoGrp = .RowData(.Row)
                    If Not tsoGrp Is Nothing Then
                        If tsoGrp.AllowManualSubmission Then
                            bClearOnly = False
                    
                            If Len(.TextMatrix(.Row, GDCol(eGDCol_Favorite))) > 0 Then
                                If Val(.TextMatrix(.Row, GDCol(eGDCol_Favorite))) = (lFavoriteIndex + 1) Then
                                    If InfBox("Clear assignment for button '" & Str(lFavoriteIndex + 1) & "'?", "?", "Yes|No") = "Y" Then
                                        bClearOnly = True
                                    End If
                                Else
                                    ' The TradeSense order group on this row was previously assigned to
                                    ' a DIFFERENT favorite index, so clear that one...
                                    ClearFavorites Int(Val(.TextMatrix(.Row, GDCol(eGDCol_Favorite)))) - 1
                                End If
                            End If
                            
                            ' Clear any TradeSense order group that was assigned to THIS favorite
                            ' index before assigning new one...
                            ClearFavorites lFavoriteIndex
                            
                            If bClearOnly = False Then
                                g.ChartGlobals.astrTsogFavorites(lFavoriteIndex) = TsogFavoriteString(tsoGrp)
                                .TextMatrix(.Row, GDCol(eGDCol_Favorite)) = Str(lFavoriteIndex + 1)
                            End If
                        Else
                            InfBox "Manual submission is not allowed for this TradeSense order group, therefore you|cannot assign it to favorites", "!", , "Error"
                        End If
                    End If
                End If
            End If
        End With
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroups.AssignFavorites"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ClearFavorites
'' Description: Clear global favorite either by specified index or if matches passed in TSO group
'' Inputs:      Index number for global array holding Favorite TSO groups (there are only 4 items in global array)
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ClearFavorites(ByVal lFavoriteIndex&)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop

    If (lFavoriteIndex >= 0) And (lFavoriteIndex < 4) Then
        g.ChartGlobals.astrTsogFavorites(lFavoriteIndex) = ""
        
        With fgGroups
            For lIndex = .FixedRows To .Rows - 1
                If .TextMatrix(lIndex, GDCol(eGDCol_Favorite)) = Str(lFavoriteIndex + 1) Then
                    .TextMatrix(lIndex, GDCol(eGDCol_Favorite)) = ""
                    Exit For
                End If
            Next lIndex
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroups.ClearFavorites"
    
End Sub

