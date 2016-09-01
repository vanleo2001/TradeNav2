VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmReassignFills 
   Caption         =   "Form1"
   ClientHeight    =   4395
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   ScaleHeight     =   4395
   ScaleWidth      =   6855
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   3615
      Left            =   2820
      TabIndex        =   3
      Top             =   360
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
      Caption         =   "frmReassignFills.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmReassignFills.frx":002C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmReassignFills.frx":004C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   0
         TabIndex        =   0
         Top             =   3240
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
         Caption         =   "frmReassignFills.frx":0068
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmReassignFills.frx":0096
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmReassignFills.frx":00B6
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Default         =   -1  'True
         Height          =   375
         Left            =   0
         TabIndex        =   9
         Top             =   2820
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
         Caption         =   "frmReassignFills.frx":00D2
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmReassignFills.frx":00F8
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmReassignFills.frx":0118
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdDeleteFill 
         Height          =   375
         Left            =   0
         TabIndex        =   8
         Top             =   2040
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
         Caption         =   "frmReassignFills.frx":0134
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmReassignFills.frx":016C
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmReassignFills.frx":018C
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdEditFill 
         Height          =   375
         Left            =   0
         TabIndex        =   7
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
         Caption         =   "frmReassignFills.frx":01A8
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmReassignFills.frx":01DC
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmReassignFills.frx":01FC
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdNewFill 
         Height          =   375
         Left            =   0
         TabIndex        =   6
         Top             =   1200
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
         Caption         =   "frmReassignFills.frx":0218
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmReassignFills.frx":024A
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmReassignFills.frx":026A
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdToManual 
         Height          =   375
         Left            =   0
         TabIndex        =   5
         Top             =   420
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
         Caption         =   "frmReassignFills.frx":0286
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmReassignFills.frx":02BE
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmReassignFills.frx":02DE
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdToAuto 
         Height          =   375
         Left            =   0
         TabIndex        =   4
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
         Caption         =   "frmReassignFills.frx":02FA
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmReassignFills.frx":032E
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmReassignFills.frx":034E
         RightToLeft     =   0   'False
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid fgAutoFills 
      Height          =   3615
      Left            =   4140
      TabIndex        =   2
      Top             =   360
      Width           =   2535
      _cx             =   4471
      _cy             =   6376
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
      OleDropMode     =   1
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VSFlex7LCtl.VSFlexGrid fgManualFills 
      Height          =   3615
      Left            =   180
      TabIndex        =   1
      Top             =   360
      Width           =   2535
      _cx             =   4471
      _cy             =   6376
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
      OleDropMode     =   1
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin HexUniControls.ctlUniLabelXP lblAutoPosition 
      Height          =   195
      Left            =   4140
      Top             =   4020
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
      Caption         =   "frmReassignFills.frx":036A
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmReassignFills.frx":039C
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmReassignFills.frx":03BC
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblManualPosition 
      Height          =   195
      Left            =   180
      Top             =   4020
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
      Caption         =   "frmReassignFills.frx":03D8
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmReassignFills.frx":040A
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmReassignFills.frx":042A
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblAutoFills 
      Height          =   195
      Left            =   4140
      Top             =   120
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
      Caption         =   "frmReassignFills.frx":0446
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmReassignFills.frx":0496
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmReassignFills.frx":04B6
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblManualFills 
      Height          =   195
      Left            =   180
      Top             =   120
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
      Caption         =   "frmReassignFills.frx":04D2
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmReassignFills.frx":050C
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmReassignFills.frx":052C
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "PopUp"
      Begin VB.Menu mnuToManual 
         Caption         =   "Move To &Manual"
      End
      Begin VB.Menu mnuToAuto 
         Caption         =   "Move to &Auto"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNewFill 
         Caption         =   "&New Fill"
      End
      Begin VB.Menu mnuEditFill 
         Caption         =   "&Edit Fill"
      End
      Begin VB.Menu mnuDeleteFill 
         Caption         =   "&Delete Fill"
      End
   End
End
Attribute VB_Name = "frmReassignFills"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmReassignFills.frm
'' Description: Allow the user to reassign fills between manual and auto trade item
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 10/21/2011   DAJ         Set focus to auto grid up front for default on "New" button
'' 10/24/2014   DAJ         Include all fills for auto trade item; all fills for base
''                          symbol when continuous; Fill Display
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    bOK As Boolean                      ' Did the user click the OK button?
    bDirty As Boolean                   ' Has the user changed something?
    
    TradeItem As cAutoTradeItem         ' Automated trading item for this form
    SymbolOrSymbolID As Variant         ' Symbol or Symbol ID for the auto trade item
    
    strSelectedGrid As String           ' Grid that last had the focus
    strDragSource As String             ' Grid that started the drag
End Type
Private m As mPrivate

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Setup and show the form
'' Inputs:      Auto Trade Item
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(ByVal TradeItem As cAutoTradeItem, Optional bDirty As Boolean) As Boolean
On Error GoTo ErrSection:

    Caption = "Reassign Fills for " & TradeItem.Symbol & " in " & g.Broker.AccountNumberForID(TradeItem.AccountID)
    
    Set m.TradeItem = TradeItem
    m.SymbolOrSymbolID = ConvertToTradeSymbol(m.TradeItem.SymbolOrSymbolID, CurrentTime("", m.TradeItem.Symbol, True))
    
    InitGrid fgManualFills
    LoadManualFillsGrid
    
    InitGrid fgAutoFills
    LoadAutoFillsGrid
    
    m.bDirty = False
    EnableControls
    
    ShowForm Me, eForm_Modal, frmMain
    
    bDirty = m.bDirty
    If m.bOK Then
        ReassignFills
    End If
    
    ShowMe = m.bOK

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmReassignFills.ShowMe"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: User has chosen to cancel out of the dialog
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancel_Click()
On Error GoTo ErrSection:

    m.bOK = False
    Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReassignFills.cmdCancel_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdDeleteFill_Click
'' Description: User has chosen to delete an existing fill
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdDeleteFill_Click()
On Error GoTo ErrSection:

    DeleteFill

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReassignFills.cmdDeleteFill_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdEditFill_Click
'' Description: User has chosen to edit an existing fill
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdEditFill_Click()
On Error GoTo ErrSection:

    EditFill
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReassignFills.cmdEditFill_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdNewFill_Click
'' Description: User has chosen to create a new fill
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdNewFill_Click()
On Error GoTo ErrSection:

    NewFill

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReassignFills.cmdNewFill_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOK_Click
'' Description: User has chosen to OK the dialog
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
On Error GoTo ErrSection:

    m.bOK = True
    Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReassignFills.cmdOK_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdToAuto_Click
'' Description: Move the selected manual fills to the auto grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdToAuto_Click()
On Error GoTo ErrSection:

    MoveToAuto

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReassignFills.cmdToAuto_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdToManual_Click
'' Description: Move the selected auto fills to the manual grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdToManual_Click()
On Error GoTo ErrSection:

    MoveToManual

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReassignFills.cmdToManual_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgAutoFills_AfterRowColChange
'' Description: Event fired when the user changes the row or column in the grid
'' Inputs:      Old Row and Column, New Row and Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgAutoFills_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReassignFills.fgAutoFills_AfterRowColChange"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgAutoFills_BeforeMouseDown
'' Description: Event fired when the user presses a mouse button on the grid
'' Inputs:      Button Pressed, Shift/Ctrl/Alt Status, Mouse Location, Cancel?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgAutoFills_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Increment variable
    Dim lMouseRow As Long               ' Current mouse row
    
    Static slRow As Long                ' Last row that was selected
    Static slXXX As Long

    With fgAutoFills
        If Button = vbRightButton Then
            .RowSel = .MouseRow
            
            EnableControls
            mnuToManual.Visible = True
            mnuToAuto.Visible = False
            PopupMenu mnuPopUp
        Else
            If (.MouseRow >= .FixedRows) And (.MouseRow < .Rows) Then
                ' Capture the mouse row in case this takes a while...
                lMouseRow = .MouseRow
        
                ' The Shift key is being pressed
                If Shift And vbShiftMask Then
                    slXXX = 0
                    
                    ' If the Control key is not down, clear the current selection and
                    ' start over
                    If (Shift And vbCtrlMask) = 0 Then
                        .Row = lMouseRow
                    End If
                    
                    ' Select everything in between the last row and the current mouse row
                    If slRow < lMouseRow Then
                        For lIndex = slRow To lMouseRow
                            .IsSelected(lIndex) = True
                        Next lIndex
                    ElseIf slRow > lMouseRow Then
                        For lIndex = slRow To lMouseRow Step -1
                            .IsSelected(lIndex) = True
                        Next lIndex
                    Else
                        .IsSelected(lMouseRow) = True
                    End If
                    
                ' The Control key is being pressed, but not the Shift key
                ElseIf Shift And vbCtrlMask Then
                    slXXX = 0
                    
                    ' Toggle the selection of the row being clicked on
                    .IsSelected(lMouseRow) = Not .IsSelected(lMouseRow)
                
                ' No key is being pressed (that we care about)
                Else
                    ' If the current row is not selected or it has been clicked twice in
                    ' a row, then clear out the current selection and start over
                    If .IsSelected(lMouseRow) = False Or lMouseRow = slXXX Then
                        .Row = lMouseRow
                        slXXX = 0
                    Else
                        slXXX = lMouseRow
                    End If
                    
                    .IsSelected(lMouseRow) = True
                End If
                
                ' If the Shift key was not being pressed, then change the last saved row
                If (Shift And vbShiftMask) = 0 Then
                    If .SelectedRows > 0 Then
                        slRow = lMouseRow
                    Else
                        slRow = 0&
                    End If
                End If
                
                ' Use OLEDrag method to start manual OLE drag operation
                ' this will fire the OLEStartDrag event, which we will use
                ' to fill the DataObject with the data we want to drag.
                .OLEDrag
        
                ' Tell grid control to ignore mouse movements until the
                ' mouse button goes up again
                Cancel = True
                
                EnableControls
            End If
        End If
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReassignFills.fgAutoFills_BeforeMouseDown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgAutoFills_Compare
'' Description: Event fired to custom compare rows for sort purposes
'' Inputs:      Row1, Row2, Compare Value
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgAutoFills_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)
On Error GoTo ErrSection:

    Cmp = Compare(fgAutoFills, Row1, Row2)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReassignFills.fgAutoFills_Compare"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgAutoFills_DblClick
'' Description: Event fired when user double clicks on the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgAutoFills_DblClick()
On Error GoTo ErrSection:

    Dim lMouseRow As Long               ' Row that the mouse is in
    
    With fgAutoFills
        lMouseRow = .MouseRow
        If (lMouseRow >= .FixedRows) And (lMouseRow < .Rows) Then
            .Row = lMouseRow
            .RowSel = lMouseRow
            
            MoveToManual
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReassignFills.fgAutoFills_DblClick"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgAutoFills_GotFocus
'' Description: Event fired when the grid gets the focus
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgAutoFills_GotFocus()
On Error GoTo ErrSection:

    m.strSelectedGrid = "Auto"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReassignFills.fgAutoFills_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgAutoFills_OLEDragDrop
'' Description: Event fired when items are dropped onto the grid
'' Inputs:      Data, Effect, Mouse Button, Shift/Ctrl/Alt Status, Mouse Location
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgAutoFills_OLEDragDrop(Data As VSFlex7LCtl.VSDataObject, Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
On Error GoTo ErrSection:

    If m.strDragSource = "Manual" Then
        MoveToAuto
        fgManualFills.OLEDropMode = flexOLEDropNone
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReassignFills.fgAutoFills_OleDragDrop"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgAutoFills_OLEStartDrag
'' Description: Begin Drag Procedure
'' Inputs:      Data to drag, Allowed effects of the drag
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgAutoFills_OLEStartDrag(Data As VSFlex7LCtl.VSDataObject, AllowedEffects As Long)
On Error GoTo ErrSection:

    m.strDragSource = "Auto"
    fgManualFills.OLEDropMode = flexOLEDropManual
    Data.SetData fgAutoFills.Clip, vbCFText

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReassignFills.fgAutoFills_OLEStartDrag"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgManualFills_AfterRowColChange
'' Description: Event fired when the user changes the row or column in the grid
'' Inputs:      Old Row and Column, New Row and Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgManualFills_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReassignFills.fgManualFills_AfterRowColChange"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgManualFills_BeforeMouseDown
'' Description: Event fired when the user presses a mouse button on the grid
'' Inputs:      Button Pressed, Shift/Ctrl/Alt Status, Mouse Location, Cancel?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgManualFills_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Increment variable
    Dim lMouseRow As Long               ' Current mouse row
    
    Static slRow As Long                ' Last row that was selected
    Static slXXX As Long

    With fgManualFills
        If Button = vbRightButton Then
            .Row = .MouseRow
            
            EnableControls
            mnuToManual.Visible = False
            mnuToAuto.Visible = True
            PopupMenu mnuPopUp
        Else
            If (.MouseRow >= .FixedRows) And (.MouseRow < .Rows) Then
                ' Capture the mouse row in case this takes a while...
                lMouseRow = .MouseRow
        
                ' The Shift key is being pressed
                If Shift And vbShiftMask Then
                    slXXX = 0
                    
                    ' If the Control key is not down, clear the current selection and
                    ' start over
                    If (Shift And vbCtrlMask) = 0 Then
                        .Row = lMouseRow
                    End If
                    
                    ' Select everything in between the last row and the current mouse row
                    If slRow < lMouseRow Then
                        For lIndex = slRow To lMouseRow
                            .IsSelected(lIndex) = True
                        Next lIndex
                    ElseIf slRow > lMouseRow Then
                        For lIndex = slRow To lMouseRow Step -1
                            .IsSelected(lIndex) = True
                        Next lIndex
                    Else
                        .IsSelected(lMouseRow) = True
                    End If
                    
                ' The Control key is being pressed, but not the Shift key
                ElseIf Shift And vbCtrlMask Then
                    slXXX = 0
                    
                    ' Toggle the selection of the row being clicked on
                    .IsSelected(lMouseRow) = Not .IsSelected(lMouseRow)
                
                ' No key is being pressed (that we care about)
                Else
                    ' If the current row is not selected or it has been clicked twice in
                    ' a row, then clear out the current selection and start over
                    If .IsSelected(lMouseRow) = False Or lMouseRow = slXXX Then
                        .Row = lMouseRow
                        slXXX = 0
                    Else
                        slXXX = lMouseRow
                    End If
                    
                    .IsSelected(lMouseRow) = True
                End If
                
                ' If the Shift key was not being pressed, then change the last saved row
                If (Shift And vbShiftMask) = 0 Then
                    If .SelectedRows > 0 Then
                        slRow = lMouseRow
                    Else
                        slRow = 0&
                    End If
                End If
                
                ' Use OLEDrag method to start manual OLE drag operation
                ' this will fire the OLEStartDrag event, which we will use
                ' to fill the DataObject with the data we want to drag.
                .OLEDrag
        
                ' Tell grid control to ignore mouse movements until the
                ' mouse button goes up again
                Cancel = True
                
                EnableControls
            End If
        End If
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReassignFills.fgManualFills_BeforeMouseDown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgManualFills_Compare
'' Description: Event fired to custom compare rows for sort purposes
'' Inputs:      Row1, Row2, Compare Value
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgManualFills_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)
On Error GoTo ErrSection:

    Cmp = Compare(fgManualFills, Row1, Row2)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReassignFills.fgManualFills_Compare"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgManualFills_DblClick
'' Description: Event fired when user double clicks on the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgManualFills_DblClick()
On Error GoTo ErrSection:

    Dim lMouseRow As Long               ' Row that the mouse is in
    
    With fgManualFills
        lMouseRow = .MouseRow
        If (lMouseRow >= .FixedRows) And (lMouseRow < .Rows) Then
            .Row = lMouseRow
            .RowSel = lMouseRow
            
            MoveToAuto
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReassignFills.fgAutoFills_DblClick"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgManualFills_GotFocus
'' Description: Event fired when the grid gets the focus
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgManualFills_GotFocus()
On Error GoTo ErrSection:

    m.strSelectedGrid = "Manual"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReassignFills.fgManualFills_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgManualFills_OLEDragDrop
'' Description: Event fired when items are dropped onto the grid
'' Inputs:      Data, Effect, Mouse Button, Shift/Ctrl/Alt Status, Mouse Location
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgManualFills_OLEDragDrop(Data As VSFlex7LCtl.VSDataObject, Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
On Error GoTo ErrSection:

    If m.strDragSource = "Auto" Then
        MoveToManual
        fgAutoFills.OLEDropMode = flexOLEDropNone
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReassignFills.fgManualFills_OLEDragDrop"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgManualFills_OLEStartDrag
'' Description: Begin Drag Procedure
'' Inputs:      Data to drag, Allowed effects of the drag
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgManualFills_OLEStartDrag(Data As VSFlex7LCtl.VSDataObject, AllowedEffects As Long)
On Error GoTo ErrSection:

    m.strDragSource = "Manual"
    fgAutoFills.OLEDropMode = flexOLEDropManual
    Data.SetData fgManualFills.Clip, vbCFText

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReassignFills.fgManualFills_OLEStartDrag"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Activate
'' Description: Handle the form getting the focus
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Activate()
On Error GoTo ErrSection:

    Static bAlreadyDone As Boolean      ' Have we already done the "do once" code?
    
    If bAlreadyDone = False Then
        ' 10/21/2011 DAJ: Set the automated trading fills grid as the selected grid so that if the
        ' user clicks on the "New Fill" button right off the bat, the category gets defaulted
        ' to the automated trading item, not manual...
        MoveFocus fgAutoFills
        
        bAlreadyDone = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReassignFills.Form_Activate"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Setup and intialize the form upon loading
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Dim strPlacement As String          ' Last known placement of the form

    strPlacement = GetIniFileProperty("ReassignFills", "", "Placement", g.strIniFile)
    If Len(strPlacement) = 0 Then
        CenterTheForm Me
    Else
        SetFormPlacement Me, strPlacement, "LHTW"
    End If
    Icon = Picture16("kBlank")
    
    g.Styler.StyleForm Me
    
    m.strDragSource = ""
    m.strSelectedGrid = ""
    
    mnuPopUp.Visible = False
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReassignFills.Form_Load"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: Handle the form being unloaded
'' Inputs:      Cancel the Unload?, Mode of the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode <> vbFormCode Then
        m.bOK = False
        Cancel = True
        Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReassignFills.Form_QueryUnload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Resize
'' Description: Move and resize controls as the form is resized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Resize()
On Error Resume Next

    Dim lLeftOver As Long               ' Leftover space for the grids

    If Not LimitFormSize(Me, 6855, 4170) Then
        lLeftOver = ScaleWidth - fraButtons.Width - (120 * 4)
        
        With fgManualFills
            .Move 120, .Top, lLeftOver / 2, ScaleHeight - .Top - lblManualPosition.Height - 120
        End With
        With lblManualPosition
            .Move 120, fgManualFills.Top + fgManualFills.Height, fgManualFills.Width
        End With
        With fraButtons
            .Move fgManualFills.Width + (120 * 2), fgManualFills.Top
        End With
        With lblAutoFills
            .Move fraButtons.Left + fraButtons.Width + 120
        End With
        With fgAutoFills
            .Move lblAutoFills.Left, fgManualFills.Top, fgManualFills.Width, fgManualFills.Height
        End With
        With lblAutoPosition
            .Move lblAutoFills.Left, lblManualPosition.Top, fgAutoFills.Width
        End With
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: Save settings and clean up upon unloading
'' Inputs:      Whether to Cancel the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    SetIniFileProperty "ReassignFills", GetFormPlacement(Me), "Placement", g.strIniFile

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReassignFills.Form_Unload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuDeleteFill_Click
'' Description: User has chosen to delete an existing fill
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuDeleteFill_Click()
On Error GoTo ErrSection:

    DeleteFill

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReassignFills.mnuDeleteFill_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuEditFill_Click
'' Description: User has chosen to edit an existing fill
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuEditFill_Click()
On Error GoTo ErrSection:

    EditFill

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReassignFills.mnuEditFill_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuNewFill_Click
'' Description: User has chosen to create a new fill
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuNewFill_Click()
On Error GoTo ErrSection:

    NewFill

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReassignFills.mnuNewFill_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuToAuto_Click
'' Description: User has chosen to move fills from manual to automated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuToAuto_Click()
On Error GoTo ErrSection:

    MoveToAuto

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReassignFills.mnuToAuto_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuToManual_Click
'' Description: User has chosen to move fills from automated to manual
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuToManual_Click()
On Error GoTo ErrSection:

    MoveToManual

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReassignFills.mnuToManual_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitGrid
'' Description: Initialize the given grid
'' Inputs:      Grid
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitGrid(fgGrid As VSFlexGrid)
On Error GoTo ErrSection:

    With fgGrid
        .Redraw = flexRDNone
        
        .FixedCols = 0
        .Cols = 1
        .Rows = 0
        SetupGrid fgGrid, eGridMode_List
        
        .AllowBigSelection = True
        .AllowSelection = True
        .Editable = flexEDNone
        .HighLight = flexHighlightAlways
        .OLEDropMode = flexOLEDropManual
        .ScrollTrack = True
        .SelectionMode = flexSelectionListBox
        
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReassignFills.InitGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadManualFillsGrid
'' Description: Load the manual fills grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadManualFillsGrid()
On Error GoTo ErrSection:

    Dim Fills As cPtFills               ' Collection of fills
    Dim lIndex As Long                  ' Index into a for loop
    
    Set Fills = New cPtFills
    'Fills.LoadFillsForSymbol m.TradeItem.AccountID, m.SymbolOrSymbolID, 0&
    Fills.LoadFillsForSymbol m.TradeItem.AccountID, m.TradeItem.SymbolOrSymbolID, 0&
    
    With fgManualFills
        .Redraw = flexRDNone
        
        For lIndex = 1 To Fills.Count
            FillToGrid Fills(lIndex), fgManualFills
        Next lIndex
        
        .Sort = flexSortCustom
        
        .Redraw = flexRDBuffered
    End With
    
    lblManualPosition.Caption = "Position: " & PositionToString(CalculatePosition(fgManualFills))

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReassignFills.LoadManualFillsGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadAutoFillsGrid
'' Description: Load the automated trading fills grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadAutoFillsGrid()
On Error GoTo ErrSection:

    Dim Fills As cPtFills               ' Collection of fills
    Dim lIndex As Long                  ' Index into a for loop
    
    Set Fills = New cPtFills
    'Fills.LoadFillsForSymbol m.TradeItem.AccountID, m.SymbolOrSymbolID, m.TradeItem.AutoTradeItemID
    Fills.LoadFillsForAutoTradeItem m.TradeItem.AutoTradeItemID
    
    With fgAutoFills
        .Redraw = flexRDNone
        
        For lIndex = 1 To Fills.Count
            FillToGrid Fills(lIndex), fgAutoFills
        Next lIndex
        
        .Sort = flexSortCustom
        
        .Redraw = flexRDBuffered
    End With
    
    lblAutoPosition.Caption = "Position: " & PositionToString(CalculatePosition(fgAutoFills))

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReassignFills.LoadAutoFillsGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillToGrid
'' Description: Add the given fill to the given grid
'' Inputs:      Fill, Grid
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FillToGrid(ByVal Fill As cPtFill, fgGrid As VSFlexGrid, Optional ByVal lRow As Long = -1&)
On Error GoTo ErrSection:

    Dim bIncludeSymbol As Boolean       ' Include the symbol?

    With fgGrid
        If lRow = -1& Then
            .Rows = .Rows + 1
            lRow = .Rows - 1
        End If
        
        .RowData(lRow) = Fill
        
        bIncludeSymbol = (InStr(m.TradeItem.Symbol, "-0") > 0)
        .TextMatrix(lRow, 0) = mTradeTracker.FillDisplay(Fill, bIncludeSymbol, True, False, False, False, False, False) & " (" & DateFormat(Fill.FillDate, MM_DD_YYYY, HH_MM_SS) & ")"
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReassignFills.FillToGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ChangeFillInGrid
'' Description: Change the given fill in the given grid
'' Inputs:      Fill, Grid
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ChangeFillInGrid(ByVal Fill As cPtFill, fgGrid As VSFlexGrid)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim GridFill As cPtFill             ' Fill in the grid

    With fgGrid
        For lIndex = .FixedRows To .Rows - 1
            If TypeOf .RowData(lIndex) Is cPtFill Then
                Set GridFill = .RowData(lIndex)
                If GridFill.FillID = Fill.FillID Then
                    FillToGrid Fill, fgGrid, lIndex
                    Exit For
                End If
            End If
        Next lIndex
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReassignFills.ChangeFillInGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveFillFromGrid
'' Description: Remove the given fill from the given grid
'' Inputs:      Fill, Grid
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RemoveFillFromGrid(ByVal Fill As cPtFill, fgGrid As VSFlexGrid)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim GridFill As cPtFill             ' Fill in the grid

    With fgGrid
        For lIndex = .FixedRows To .Rows - 1
            If TypeOf .RowData(lIndex) Is cPtFill Then
                Set GridFill = .RowData(lIndex)
                If GridFill.FillID = Fill.FillID Then
                    fgGrid.RemoveItem lIndex
                    Exit For
                End If
            End If
        Next lIndex
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReassignFills.RemoveFillFromGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MoveSelectedFills
'' Description: Move the selected fills from source grid to destination grid
'' Inputs:      Source Grid, Destination Grid
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub MoveSelectedFills(fgSource As VSFlexGrid, fgDestination As VSFlexGrid)
On Error GoTo ErrSection:

    Dim FillsToMove As cGdTree          ' Fills to move to the other grid
    Dim lIndex As Long                  ' Index into a for loop
    
    Set FillsToMove = New cGdTree
    With fgSource
        .Redraw = flexRDNone
        
        For lIndex = 0 To .SelectedRows - 1
            FillsToMove.Add .RowData(.SelectedRow(lIndex))
        Next lIndex
        
        For lIndex = .SelectedRows - 1 To 0 Step -1
            .RemoveItem .SelectedRow(lIndex)
        Next lIndex
        
        .Redraw = flexRDBuffered
    End With
    
    With fgDestination
        .Redraw = flexRDNone
        
        For lIndex = 1 To FillsToMove.Count
            FillToGrid FillsToMove(lIndex), fgDestination
        Next lIndex
        
        .Sort = flexSortCustom
        
        .Redraw = flexRDBuffered
    End With
    
    lblManualPosition.Caption = "Position: " & PositionToString(CalculatePosition(fgManualFills))
    lblAutoPosition.Caption = "Position: " & PositionToString(CalculatePosition(fgAutoFills))
    m.bDirty = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReassignFills.MoveSelectedFills"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MoveToAuto
'' Description: Move the selected manual fills to the auto grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub MoveToAuto()
On Error GoTo ErrSection:

    MoveSelectedFills fgManualFills, fgAutoFills

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReassignFills.MoveToAuto"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MoveToManual
'' Description: Move the selected auto fills to the manual grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub MoveToManual()
On Error GoTo ErrSection:

    MoveSelectedFills fgAutoFills, fgManualFills

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReassignFills.MoveToManual"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Compare
'' Description: Compare the two rows of the grid to determine sort order
'' Inputs:      Grid, Row1, Row2
'' Returns:     -1 if Row1 before Row2, Zero if the same, 1 if Row1 after Row2
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function Compare(fgGrid As VSFlexGrid, ByVal lRow1 As Long, ByVal lRow2 As Long) As Integer
On Error GoTo ErrSection:

    Dim iReturn As Integer              ' Return value for the function
    Dim Fill1 As cPtFill                ' First fill to compare
    Dim Fill2 As cPtFill                ' Second fill to compare

    With fgGrid
        Set Fill1 = .RowData(lRow1)
        Set Fill2 = .RowData(lRow2)
        
        If Fill1.FillDate < Fill2.FillDate Then
            iReturn = -1
        ElseIf Fill1.FillDate > Fill2.FillDate Then
            iReturn = 1
        ElseIf Fill1.FillID < Fill2.FillID Then
            iReturn = -1
        ElseIf Fill1.FillID > Fill2.FillID Then
            iReturn = 1
        Else
            iReturn = 0
        End If
    End With
    
    Compare = iReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmReassignFills.Compare"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CalculatePosition
'' Description: Calculate the position from the fills in the given grid
'' Inputs:      Grid
'' Returns:     Position
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CalculatePosition(fgGrid As VSFlexGrid) As Long
On Error GoTo ErrSection:
    
    Dim lReturn As Long                 ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim Fill As cPtFill                 ' Fill object
    
    lReturn = 0&
    With fgGrid
        For lIndex = .FixedRows To .Rows - 1
            If TypeOf .RowData(lIndex) Is cPtFill Then
                Set Fill = .RowData(lIndex)
                If Fill.Buy Then
                    lReturn = lReturn + Fill.Quantity
                Else
                    lReturn = lReturn - Fill.Quantity
                End If
            End If
        Next lIndex
    End With
    
    CalculatePosition = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmReassignFills.CalculatePosition"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NewFill
'' Description: Allow the user to create a new fill
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub NewFill()
On Error GoTo ErrSection:

    Dim Fill As cPtFill                 ' Fill object
    Dim lAutoTradeID As Long            ' Automated trading item ID
    
    If m.strSelectedGrid = "Auto" Then
        lAutoTradeID = m.TradeItem.AutoTradeItemID
    Else
        lAutoTradeID = 0&
    End If
    
    Set Fill = New cPtFill
    If g.Broker.CreateNewFill(Fill, GetSymbol(m.SymbolOrSymbolID), m.TradeItem.AccountID, lAutoTradeID) Then
        If Fill.AutoTradingItemID = 0& Then
            With fgManualFills
                .Redraw = flexRDNone
                
                FillToGrid Fill, fgManualFills
                .Sort = flexSortCustom
                
                .Redraw = flexRDBuffered
            End With
            
            lblManualPosition.Caption = "Position: " & PositionToString(CalculatePosition(fgManualFills))
        ElseIf Fill.AutoTradingItemID = m.TradeItem.AutoTradeItemID Then
            With fgAutoFills
                .Redraw = flexRDNone
                
                FillToGrid Fill, fgAutoFills
                .Sort = flexSortCustom
                
                .Redraw = flexRDBuffered
            End With
            
            lblAutoPosition.Caption = "Position: " & PositionToString(CalculatePosition(fgAutoFills))
        End If
        
        m.bDirty = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReassignFills.NewFill"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EditFill
'' Description: Allow the user to edit an existing fill
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EditFill()
On Error GoTo ErrSection:

    Dim Fill As cPtFill                 ' Fill object
    Dim lOldTradeItemID As Long         ' Old trading item ID

    Set Fill = Nothing
    If m.strSelectedGrid = "Auto" Then
        If fgAutoFills.SelectedRows > 0 Then
            Set Fill = fgAutoFills.RowData(fgAutoFills.SelectedRow(0))
        End If
    Else
        If fgManualFills.SelectedRows > 0 Then
            Set Fill = fgManualFills.RowData(fgManualFills.SelectedRow(0))
        End If
    End If
    
    If Not Fill Is Nothing Then
        lOldTradeItemID = Fill.AutoTradingItemID
        If g.Broker.ModifyFill(Fill) Then
            If Fill.AutoTradingItemID = 0& Then
                With fgManualFills
                    .Redraw = flexRDNone
                    
                    If lOldTradeItemID = Fill.AutoTradingItemID Then
                        ChangeFillInGrid Fill, fgManualFills
                    Else
                        RemoveFillFromGrid Fill, fgAutoFills
                        FillToGrid Fill, fgManualFills
                    End If
                    
                    .Sort = flexSortCustom
                    .Redraw = flexRDBuffered
                End With
            ElseIf Fill.AutoTradingItemID = m.TradeItem.AutoTradeItemID Then
                With fgAutoFills
                    .Redraw = flexRDNone
                    
                    If lOldTradeItemID = Fill.AutoTradingItemID Then
                        ChangeFillInGrid Fill, fgAutoFills
                    Else
                        RemoveFillFromGrid Fill, fgManualFills
                        FillToGrid Fill, fgAutoFills
                    End If
                    
                    .Sort = flexSortCustom
                    .Redraw = flexRDBuffered
                End With
            Else
                If lOldTradeItemID = 0& Then
                    RemoveFillFromGrid Fill, fgManualFills
                Else
                    RemoveFillFromGrid Fill, fgAutoFills
                End If
            End If
        
            lblManualPosition.Caption = "Position: " & PositionToString(CalculatePosition(fgManualFills))
            lblAutoPosition.Caption = "Position: " & PositionToString(CalculatePosition(fgAutoFills))
            m.bDirty = True
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReassignFills.EditFill"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DeleteFill
'' Description: Allow the user to delete an existing fill
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DeleteFill()
On Error GoTo ErrSection:

    Dim Fill As cPtFill                 ' Fill object

    Set Fill = Nothing
    If m.strSelectedGrid = "Auto" Then
        If fgAutoFills.SelectedRows > 0 Then
            Set Fill = fgAutoFills.RowData(fgAutoFills.SelectedRow(0))
        End If
    Else
        If fgManualFills.SelectedRows > 0 Then
            Set Fill = fgManualFills.RowData(fgManualFills.SelectedRow(0))
        End If
    End If
    
    If Not Fill Is Nothing Then
        If g.Broker.DeleteFill(Fill, "Reassign Fills") Then
            If Fill.AutoTradingItemID = 0& Then
                RemoveFillFromGrid Fill, fgManualFills
                lblManualPosition.Caption = "Position: " & PositionToString(CalculatePosition(fgManualFills))
            ElseIf Fill.AutoTradingItemID = m.TradeItem.AutoTradeItemID Then
                RemoveFillFromGrid Fill, fgAutoFills
                lblAutoPosition.Caption = "Position: " & PositionToString(CalculatePosition(fgAutoFills))
            End If
            
            m.bDirty = True
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReassignFills.DeleteFill"
    
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

    Dim bManualFillsSelected As Boolean ' Are there manual fills selected?
    Dim bAutoFillsSelected As Boolean   ' Are there automated fills selected?

    bManualFillsSelected = (fgManualFills.SelectedRows > 0)
    bAutoFillsSelected = (fgAutoFills.SelectedRows > 0)

    Enable cmdToAuto, bManualFillsSelected
    Enable mnuToAuto, bManualFillsSelected
    Enable cmdToManual, bAutoFillsSelected
    Enable mnuToManual, bAutoFillsSelected

    If m.strSelectedGrid = "Auto" Then
        Enable cmdEditFill, bAutoFillsSelected
        Enable mnuEditFill, bAutoFillsSelected
        Enable cmdDeleteFill, bAutoFillsSelected
        Enable mnuDeleteFill, bAutoFillsSelected
    Else
        Enable cmdEditFill, bManualFillsSelected
        Enable mnuEditFill, bManualFillsSelected
        Enable cmdDeleteFill, bManualFillsSelected
        Enable mnuDeleteFill, bManualFillsSelected
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReassignFills.EnableControls"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ReassignFills
'' Description: Reassign fills as necessary
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ReassignFills()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim Fill As cPtFill                 ' Fill from the grid
    Dim bFillReassigned As Boolean      ' Has a fill been reassigned?
    
    With fgManualFills
        For lIndex = .FixedRows To .Rows - 1
            If TypeOf .RowData(lIndex) Is cPtFill Then
                Set Fill = .RowData(lIndex)
                If Fill.AutoTradingItemID <> 0& Then
                    Fill.AutoTradingItemID = 0&
                    Fill.Save
                    bFillReassigned = True
                End If
            End If
        Next lIndex
    End With

    With fgAutoFills
        For lIndex = .FixedRows To .Rows - 1
            If TypeOf .RowData(lIndex) Is cPtFill Then
                Set Fill = .RowData(lIndex)
                If Fill.AutoTradingItemID = 0& Then
                    Fill.AutoTradingItemID = m.TradeItem.AutoTradeItemID
                    Fill.Save
                    bFillReassigned = True
                End If
            End If
        Next lIndex
    End With
    
    If bFillReassigned = True Then
        g.Broker.RebuildFillSummaryForSymbol m.TradeItem.AccountID, m.SymbolOrSymbolID, 0&, True
        g.Broker.RebuildFillSummaryForSymbol m.TradeItem.AccountID, m.SymbolOrSymbolID, m.TradeItem.AutoTradeItemID, True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReassignFills.ReassignFills"
    
End Sub

