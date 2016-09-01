VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmPositionFix 
   Caption         =   "Form1"
   ClientHeight    =   5400
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   5625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   495
      Left            =   1620
      TabIndex        =   7
      Top             =   4380
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
      Caption         =   "frmPositionFix.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmPositionFix.frx":002C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmPositionFix.frx":004C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Height          =   495
         Left            =   1380
         TabIndex        =   0
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
         Caption         =   "frmPositionFix.frx":0068
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmPositionFix.frx":0096
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmPositionFix.frx":00B6
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Height          =   495
         Left            =   0
         TabIndex        =   1
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
         Caption         =   "frmPositionFix.frx":00D2
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmPositionFix.frx":00F8
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmPositionFix.frx":0118
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraFillButtons 
      Height          =   1695
      Left            =   4320
      TabIndex        =   3
      Top             =   1440
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
      Caption         =   "frmPositionFix.frx":0134
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmPositionFix.frx":0160
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmPositionFix.frx":0180
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdRemove 
         Height          =   495
         Left            =   0
         TabIndex        =   6
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
         Caption         =   "frmPositionFix.frx":019C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmPositionFix.frx":01D4
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmPositionFix.frx":01F4
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdEdit 
         Height          =   495
         Left            =   0
         TabIndex        =   5
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
         Caption         =   "frmPositionFix.frx":0210
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmPositionFix.frx":0244
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmPositionFix.frx":0264
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdAdd 
         Height          =   495
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
         Caption         =   "frmPositionFix.frx":0280
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmPositionFix.frx":02B2
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmPositionFix.frx":02D2
         RightToLeft     =   0   'False
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid fgFills 
      Height          =   1695
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   4095
      _cx             =   7223
      _cy             =   2990
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
   Begin HexUniControls.ctlUniLabelXP lblProblem 
      Height          =   435
      Left            =   120
      Top             =   120
      Width           =   5355
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
      Caption         =   "frmPositionFix.frx":02EE
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmPositionFix.frx":0444
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmPositionFix.frx":0464
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblWarning 
      Height          =   435
      Left            =   420
      Top             =   3660
      Width           =   5055
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
      Caption         =   "frmPositionFix.frx":0480
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmPositionFix.frx":0586
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmPositionFix.frx":05A6
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblCurrentPositionValue 
      Height          =   195
      Left            =   2100
      Top             =   3240
      Width           =   855
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
      Caption         =   "frmPositionFix.frx":05C2
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmPositionFix.frx":05F6
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmPositionFix.frx":0616
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblCurrentPosition 
      Height          =   195
      Left            =   180
      Top             =   3240
      Width           =   1875
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
      Caption         =   "frmPositionFix.frx":0632
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmPositionFix.frx":0684
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmPositionFix.frx":06A4
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblAverageEntryValue 
      Height          =   195
      Left            =   4320
      Top             =   3240
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
      Caption         =   "frmPositionFix.frx":06C0
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmPositionFix.frx":06FE
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmPositionFix.frx":071E
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblAverageEntry 
      Height          =   195
      Left            =   3180
      Top             =   3240
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
      Caption         =   "frmPositionFix.frx":073A
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmPositionFix.frx":0776
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmPositionFix.frx":0796
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblInstructions 
      Height          =   435
      Left            =   120
      Top             =   600
      Width           =   5355
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
      Caption         =   "frmPositionFix.frx":07B2
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmPositionFix.frx":093A
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmPositionFix.frx":095A
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblFills 
      Height          =   255
      Left            =   120
      Top             =   1140
      Width           =   5355
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
      Caption         =   "frmPositionFix.frx":0976
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmPositionFix.frx":0A2C
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmPositionFix.frx":0A4C
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUp"
      Begin VB.Menu mnuAddFill 
         Caption         =   "Add Fill"
      End
      Begin VB.Menu mnuEditFill 
         Caption         =   "Edit Fill"
      End
      Begin VB.Menu mnuRemoveFill 
         Caption         =   "Remove Fill"
      End
   End
End
Attribute VB_Name = "frmPositionFix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmPositionFix.frm
'' Description: Allow the user to fix their position by changing fills
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 07/10/2014   DAJ         Ensure order snapshot flag is the same as the fill
''                          snapshot flag after fill edit
'' 10/24/2014   DAJ         Fill Display
'' 12/10/2014   DAJ         Utilize new DateIsSnapshot routines
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Enum eGDCols
    eGDCol_DateTime = 0
    eGdCol_BuySell
    eGdCol_Quantity
    eGDCol_Price
    eGDCol_Position
    eGDCol_BrokerFillID
    eGDCol_NumCols
End Enum

Private Type mPrivate
    bOK As Boolean                      ' Did the user click on OK?
    
    FillSummary As cAccountPosition     ' Account position information
    lAccountID As Long                  ' Account ID
    strAccountNumber As String          ' Account number
    lSymbolID As Long                   ' Symbol ID
    strSymbol As String                 ' Symbol
    lBrokerPosition As Long             ' Broker Reported Position
    lBrokerCarried As Long              ' Broker Reported Carried Position
    Fills As cPtFills                   ' Snapshot fills
    nFillMatchMode As eTT_FillMatchMode ' Fill match mode for this account
    Bars As cGdBars                     ' Bars object
    
    ToSave As cGdTree                   ' Collection of fills to save upon exit
    ToRemove As cGdTree                 ' Collection of fills to remove upon exit
    
    lPosition As Long                   ' Current calculated position
End Type
Private m As mPrivate

Private Function GDCol(ByVal nCol As eGDCols) As Long
    GDCol = nCol
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Setup and show the form
'' Inputs:      Fill Summary, Broker Position, Broker Carried
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(ByVal FillSummary As cAccountPosition, ByVal lBrokerPosition As Long, ByVal lBrokerCarried As Long) As Boolean
On Error GoTo ErrSection:

    Dim lDifference As Long             ' Difference between positions
    Dim strBrokerPosition As String     ' String representation of broker position
    Dim Account As cPtAccount           ' Account object

    ' Calculate the difference in position...
    lDifference = lBrokerPosition - FillSummary.CurrentPositionSnapshot
    
    ' Only need to show the form if there is a difference in the position...
    If lDifference <> 0 Then
        ' Set the module level variables from what was passed in...
        Set m.FillSummary = FillSummary
        m.lAccountID = FillSummary.AccountID
        m.strAccountNumber = g.Broker.AccountNumberForID(FillSummary.AccountID)
        m.lSymbolID = FillSummary.SymbolID
        m.strSymbol = FillSummary.Symbol
        m.lBrokerCarried = lBrokerCarried
        m.lBrokerPosition = lBrokerPosition
        strBrokerPosition = g.Broker.TextPosition(lBrokerPosition)
        m.lPosition = FillSummary.CurrentPositionSnapshot
        
        Set Account = New cPtAccount
        If Account.Load(m.lAccountID) Then
            m.nFillMatchMode = Account.FillMatchMode
        End If
        
        Set m.Bars = New cGdBars
        SetBarProperties m.Bars, m.strSymbol
        
        Set m.Fills = New cPtFills
        m.Fills.LoadFillsForSymbol m.strAccountNumber, m.strSymbol, -1&
        
        Set m.ToSave = New cGdTree
        Set m.ToRemove = New cGdTree
        
        ' Set the caption based on the information...
        Caption = "Synchronization for " & m.strSymbol & " in account " & m.strAccountNumber
        
        ' Change the information labels as appropriate...
        lblProblem.Caption = Replace(lblProblem.Caption, "<Broker>", g.Broker.BrokerName(FillSummary.Broker))
        lblProblem.Caption = Replace(lblProblem.Caption, "<BrokerPos>", UCase(g.Broker.TextPosition(lBrokerCarried)))
        lblProblem.Caption = Replace(lblProblem.Caption, "<symbol>", m.strSymbol)
        lblProblem.Caption = Replace(lblProblem.Caption, "<account>", m.strAccountNumber)
        lblProblem.Caption = Replace(lblProblem.Caption, "<TnPos>", UCase(g.Broker.TextPosition(FillSummary.CurrentPosition)))
        lblFills.Caption = Replace(lblFills.Caption, "<account>", m.strAccountNumber)
        lblFills.Caption = Replace(lblFills.Caption, "<symbol>", m.strSymbol)
        lblInstructions.Caption = Replace(lblInstructions.Caption, "<quantity>", Str(Abs(lDifference)))
        Select Case lDifference
            Case 1
                lblInstructions.Caption = Replace(lblInstructions.Caption, "<direction>", "BUY")
            Case -1
                lblInstructions.Caption = Replace(lblInstructions.Caption, "<direction>", "SELL")
            Case Is > 0
                lblInstructions.Caption = Replace(lblInstructions.Caption, "<direction>", "BUYS")
            Case Is < 0
                lblInstructions.Caption = Replace(lblInstructions.Caption, "<direction>", "SELLS")
        End Select
        lblInstructions.Caption = Replace(lblInstructions.Caption, "<position>", UCase(strBrokerPosition))
        lblCurrentPositionValue.Caption = UCase(FillSummary.CurrentPositionSnapshotString)
        lblAverageEntryValue.Caption = FillSummary.AverageEntrySnapshotString
        lblWarning.Caption = UCase(Replace(lblWarning.Caption, "<position>", UCase(strBrokerPosition)))
        
        InitGrid
        LoadGrid
        EnableControls
    
        ShowForm Me, eForm_ActModal, frmMain, , ALT_GRID_ROW_COLOR
        
        If m.bOK Then
            Save
        End If
    Else
        m.bOK = True
    End If
    
    ShowMe = m.bOK

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmPositionFix.ShowMe"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdAdd_Click
'' Description: Allow the user to add a new fill
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdAdd_Click()
On Error GoTo ErrSection:

    AddFill

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPositionFix.cmdAdd_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: Allow the ShowMe to exit the form
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
    RaiseError "frmPositionFix.cmdCancel_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdEdit_Click
'' Description: Allow the user to edit an existing fill
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdEdit_Click()
On Error GoTo ErrSection:

    EditFill

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPositionFix.cmdEdit_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOK_Click
'' Description: Allow the ShowMe to exit the form
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
    RaiseError "frmPositionFix.cmdOK_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdRemove_Click
'' Description: Allow the user to remove a fill
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdRemove_Click()
On Error GoTo ErrSection:

    RemoveFill

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPositionFix.cmdRemove_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgFills_AfterRowColChange
'' Description: After a cell change, enable/disable controls
'' Inputs:      Old Row and Column, New Row and Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgFills_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPositionFix.fgFills_AfterRowColChange"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgFills_BeforeMouseDown
'' Description: If the user right clicks, show the popup menu
'' Inputs:      Button Pressed, Shift/Ctrl/Alt Status, Mouse Location
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgFills_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim lMouseRow As Long               ' Current mouse row in the grid
    Dim Fill As cPtFill                 ' Fill object
    
    If Button = vbRightButton Then
        With fgFills
            lMouseRow = .MouseRow
        
            If ValidRow(lMouseRow) = False Then
                mnuAddFill.Enabled = True
                mnuEditFill.Enabled = False
                mnuRemoveFill.Enabled = False
            Else
                .Row = lMouseRow
                .RowSel = lMouseRow
                
                Set Fill = .RowData(lMouseRow)
                mnuAddFill.Enabled = True
                mnuEditFill.Enabled = Not (Fill.IsSnapshot = True And Fill.IsManual = False)
                mnuRemoveFill.Enabled = Not (Fill.IsSnapshot = True And Fill.IsManual = False)
            End If
            
            PopupMenu mnuPopUp
        End With
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPositionFix.fgFills_BeforeMouseDown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgFills_DblClick
'' Description: Allow the user to edit a fill by double clicking on it
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgFills_DblClick()
On Error GoTo ErrSection:

    Dim lMouseRow As Long               ' Current mouse row in the grid
    Dim Fill As cPtFill                 ' Fill object

    With fgFills
        lMouseRow = .MouseRow
        
        If ValidRow(lMouseRow) Then
            .Row = lMouseRow
            .RowSel = lMouseRow
            
            Set Fill = .RowData(.Row)
            If (Fill.IsSnapshot = False) Or (Fill.IsManual = True) Then
                EditFill
            End If
        End If
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPositionFix.fgFills_DblClick"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgFills_KeyDown
'' Description: Insert or Remove a fill if insert or delete key is pressed
'' Inputs:      Code of Key Pressed, Shift/Ctrl/Alt Status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgFills_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    Dim Fill As cPtFill                 ' Fill object
    
    Select Case KeyCode
        Case vbKeyInsert
            AddFill
        
        Case vbKeyDelete
            If ValidRow(fgFills.Row) Then
                Set Fill = fgFills.RowData(fgFills.Row)
                If (Fill.IsSnapshot = False) Or (Fill.IsManual = True) Then
                    RemoveFill
                End If
            End If
            
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPositionFix.fgFills_KeyDown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Initialize the form when it is loaded
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Icon = Picture16("kBlank")
    CenterTheForm Me
    
    g.Styler.StyleForm Me
    
    ' Don't show the pop-up menu unless user right-clicks in the grid...
    mnuPopUp.Visible = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPositionFix.Form_Load"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: If the user closes the form with the 'X', allow ShowMe to exit
'' Inputs:      Cancel the Unload?, Mode of the unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode <> vbFormCode Then
        Cancel = True
        m.bOK = False
        Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPositionFix.Form_QueryUnload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Resize
'' Description: Move and resize the controls as the form is resized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Resize()
On Error Resume Next

    If Not LimitFormSize(Me, 6600, 4665) Then
        With lblProblem
            .Move 60, .Top, ScaleWidth - 120, .Height
        End With
        
        With lblFills
            .Move 60, .Top, ScaleWidth - 120, .Height
        End With
    
        With lblInstructions
            .Move 60, .Top, ScaleWidth - 120, .Height
        End With
        
        With fgFills
            .Move 60, .Top, ScaleWidth - fraFillButtons.Width - 180, ScaleHeight - .Top - lblAverageEntry.Height - lblWarning.Height - fraButtons.Height - 360
        End With
        
        With fraFillButtons
            .Move ScaleWidth - .Width - 60, fgFills.Top, .Width, .Height
        End With
        
        With lblAverageEntry
            .Move .Left, fgFills.Top + fgFills.Height + 60, .Width, .Height
        End With
    
        With lblAverageEntryValue
            .Move .Left, fgFills.Top + fgFills.Height + 60, .Width, .Height
        End With
        
        With lblCurrentPosition
            .Move .Left, fgFills.Top + fgFills.Height + 60, .Width, .Height
        End With
    
        With lblCurrentPositionValue
            .Move .Left, fgFills.Top + fgFills.Height + 60, .Width, .Height
        End With
        
        With lblWarning
            .Move 60, lblAverageEntry.Top + lblAverageEntry.Height + 120, ScaleWidth - 120, .Height
        End With
        
        With fraButtons
            .Move (ScaleWidth - .Width) / 2, lblWarning.Top + lblWarning.Height + 120, .Width, .Height
        End With
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuAddFill_Click
'' Description: Allow the user to add a new fill
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuAddFill_Click()
On Error GoTo ErrSection:

    AddFill

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPositionFix.mnuAddFill_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuEditFill_Click
'' Description: Allow the user to edit an existing fill
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
    RaiseError "frmPositionFix.mnuEditFill_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuRemoveFill_Click
'' Description: Allow the user to remove an existing fill
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuRemoveFill_Click()
On Error GoTo ErrSection:

    RemoveFill

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPositionFix.mnuRemoveFill_Click"

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

    With fgFills
        .Redraw = flexRDNone
        
        .AllowBigSelection = False
        .AllowSelection = False
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .Editable = flexEDNone
        .ExplorerBar = flexExNone
        .ExtendLastCol = True
        .HighLight = flexHighlightNever
        .ScrollTrack = True
        .SelectionMode = flexSelectionListBox
        .SheetBorder = RGB(128, 128, 128)
        
        .FixedRows = 1
        .Rows = .FixedRows
        .FixedCols = 0
        .Cols = GDCol(eGDCol_NumCols)
        
        .TextMatrix(0, GDCol(eGDCol_DateTime)) = "Time"
        .TextMatrix(0, GDCol(eGdCol_BuySell)) = "B/S"
        .TextMatrix(0, GDCol(eGdCol_Quantity)) = "Quantity"
        .TextMatrix(0, GDCol(eGDCol_Price)) = "Price"
        .TextMatrix(0, GDCol(eGDCol_Position)) = "Pos"
        .TextMatrix(0, GDCol(eGDCol_BrokerFillID)) = "Fill ID"
        
        .ColFormat(GDCol(eGDCol_DateTime)) = DateFormat("FORMAT", MM_DD_YYYY, HH_MM_SS, AMPM_UPPER)

        .ColHidden(GDCol(eGDCol_BrokerFillID)) = True
        
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPositionFix.InitGrid"
    
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
    
    With fgFills
        .Redraw = flexRDNone
        
        ' First, make sure that the grid is cleared out...
        .Rows = .FixedRows
        
        ' Now load the fills into the grid...
        For lIndex = 1 To m.Fills.Count
            .Rows = .Rows + 1
            FillToGrid m.Fills(lIndex), .Rows - 1
        Next lIndex
        
        SortGrid
        Recalculate
        
        .AutoSize 0, .Cols - 1, False, 75
        
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPositionFix.LoadGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillToGrid
'' Description: Add the given fill to the grid at the given row
'' Inputs:      Fill, Row
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FillToGrid(ByVal Fill As cPtFill, Optional ByVal lRow As Long = -1&)
On Error GoTo ErrSection:

    With fgFills
        If ValidRow(lRow) = False Then
            .Rows = .Rows + 1
            lRow = .Rows - 1
        End If
    
        .RowData(lRow) = Fill
        .TextMatrix(lRow, GDCol(eGDCol_DateTime)) = Fill.FillDate
        If Fill.Buy Then
            .TextMatrix(lRow, GDCol(eGdCol_BuySell)) = "Buy"
        Else
            .TextMatrix(lRow, GDCol(eGdCol_BuySell)) = "Sell"
        End If
        .TextMatrix(lRow, GDCol(eGdCol_Quantity)) = Str(Fill.Quantity)
        .TextMatrix(lRow, GDCol(eGDCol_Price)) = Fill.PriceString
        .TextMatrix(lRow, GDCol(eGDCol_BrokerFillID)) = Fill.BrokerID
        
        If (Fill.IsSnapshot = True) And (Fill.IsManual = False) Then
            .Cell(flexcpForeColor, lRow, 0, lRow, .Cols - 1) = RGB(128, 128, 128)
        Else
            .Cell(flexcpForeColor, lRow, 0, lRow, .Cols - 1) = .Cell(flexcpForeColor, 0, 0)
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPositionFix.FillToGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddFill
'' Description: Allow the user to add a new fill
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddFill()
On Error GoTo ErrSection:

    Dim Fill As cPtFill                 ' New fill
    
    Set Fill = New cPtFill
    
    If frmTTEditFill.ShowMe(Fill, m.strSymbol, m.lAccountID, True) Then
        Fill.FillID = -1&
        Fill.IsSnapshot = g.Broker.DateIsSnapshotForFill(Fill)
        FillToGrid Fill
        
        SortGrid
        Recalculate
        EnableControls
    
        If m.ToSave.Exists(Fill.BrokerID) = False Then
            m.ToSave.Add Fill, Fill.BrokerID
        Else
            m.ToSave(Fill.BrokerID) = Fill
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPositionFix.AddFill"
    
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

    With fgFills
        If ValidRow(.Row) Then
            Set Fill = .RowData(.Row)
            If (Fill.IsSnapshot = False) Or (Fill.IsManual = True) Then
                If frmTTEditFill.ShowMe(Fill, m.strSymbol, m.lAccountID) Then
                    Fill.IsSnapshot = g.Broker.DateIsSnapshotForFill(Fill)
                    FillToGrid Fill, .Row
                    
                    SortGrid
                    Recalculate
                    EnableControls
                
                    If m.ToSave.Exists(Fill.BrokerID) = False Then
                        m.ToSave.Add Fill, Fill.BrokerID
                    Else
                        m.ToSave(Fill.BrokerID) = Fill
                    End If
                End If
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPositionFix.EditFill"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveFill
'' Description: Allow the user to remove a fill
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RemoveFill()
On Error GoTo ErrSection:

    Dim Fill As cPtFill                 ' Fill object

    With fgFills
        If ValidRow(.Row) Then
            Set Fill = .RowData(.Row)
            If (Fill.IsSnapshot = False) Or (Fill.IsManual = True) Then
                If InfBox("Are you sure you want to remove fill #" & Fill.BrokerID & "||" & mTradeTracker.FillDisplay(Fill) & "?", "?", "+Yes|-No", "Fill Remove Confirmation") = "Y" Then
                    .RemoveItem .Row
                    
                    Recalculate
                    EnableControls
                    
                    If m.ToRemove.Exists(Fill.BrokerID) = False Then
                        m.ToRemove.Add Fill, Fill.BrokerID
                    End If
                End If
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPositionFix.RemoveFill"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ValidRow
'' Description: Is the given row a valid row in the grid?
'' Inputs:      Row
'' Returns:     True if valid, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ValidRow(ByVal lRow As Long) As Boolean
On Error GoTo ErrSection:

    ValidRow = ((lRow >= fgFills.FixedRows) And (lRow < fgFills.Rows))

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmPositionFix.ValidRow"
    
End Function

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

    Dim Fill As cPtFill                 ' Fill object

    With fgFills
        If ValidRow(.Row) Then
            Set Fill = .RowData(.Row)
            Enable cmdEdit, ((Fill.IsSnapshot = False) Or (Fill.IsManual = True))
            Enable cmdRemove, ((Fill.IsSnapshot = False) Or (Fill.IsManual = True))
        Else
            Enable cmdEdit, False
            Enable cmdRemove, False
        End If
    End With
    
    Enable cmdOK, (m.lPosition = m.lBrokerPosition)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPositionFix.EnableControls"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SortGrid
'' Description: Sort the grid by Fill Date then by Broker Fill ID
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SortGrid()
On Error GoTo ErrSection:

    With fgFills
        .Col = GDCol(eGDCol_BrokerFillID)
        .Sort = flexSortGenericAscending
        .Col = GDCol(eGDCol_DateTime)
        .Sort = flexSortGenericAscending
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPositionFix.SortGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Recalculate
'' Description: Recalculate the position and average entry information
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Recalculate()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim Fill As cPtFill                 ' Fill object
    Dim lPosition As Long               ' Position based on this fill
    Dim adEntries As cGdArray           ' Array of entry values
    
    lPosition = 0&
    Set adEntries = New cGdArray
    adEntries.Create eGDARRAY_Doubles
    
    With fgFills
        For lIndex = .FixedRows To .Rows - 1
            Set Fill = .RowData(lIndex)
            If Fill.Buy Then
                If lPosition >= 0 Then
                    For lIndex2 = 1 To Fill.Quantity
                        adEntries.Add Fill.Price
                    Next lIndex2
                ElseIf m.nFillMatchMode = eTT_FillMatchMode_Fifo Then
                    adEntries.Remove 0, Fill.Quantity
                Else
                    adEntries.Remove adEntries.Size - Fill.Quantity - 1, Fill.Quantity
                End If
                
                lPosition = lPosition + Fill.Quantity
            Else
                If lPosition <= 0 Then
                    For lIndex2 = 1 To Fill.Quantity
                        adEntries.Add Fill.Price
                    Next lIndex2
                ElseIf m.nFillMatchMode = eTT_FillMatchMode_Fifo Then
                    adEntries.Remove 0, Fill.Quantity
                Else
                    adEntries.Remove adEntries.Size - Fill.Quantity - 1, Fill.Quantity
                End If
            
                lPosition = lPosition - Fill.Quantity
            End If
            .TextMatrix(lIndex, GDCol(eGDCol_Position)) = g.Broker.TextPosition(lPosition)
        Next lIndex
    End With
    
    m.lPosition = lPosition
    lblCurrentPositionValue.Caption = UCase(g.Broker.TextPosition(m.lPosition))
    
    If lPosition = 0& Then
        lblAverageEntryValue.Caption = ""
    Else
        lblAverageEntryValue.Caption = m.Bars.PriceDisplay(adEntries.CalcStatistic(eGdStat_Average))
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPositionFix.Recalculate"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Save
'' Description: Save the changes the user has done
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Save()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim Fill As cPtFill                 ' Fill object
    Dim bRebuildHistory As Boolean      ' Rebuild the history?
    Dim Order As cPtOrder               ' Order object
    
    bRebuildHistory = False
    For lIndex = 1 To m.ToRemove.Count
        Set Fill = m.ToRemove(lIndex)
        Fill.Delete
        If g.Broker.DateIsSnapshotForFill(Fill) Then
            g.Broker.RemoveFill Fill, False
        Else
            bRebuildHistory = True
        End If
    Next lIndex
    
    For lIndex = 1 To m.ToSave.Count
        Set Fill = m.ToSave(lIndex)
        Fill.Save
        
        ' DAJ 07/10/2014: If the fill date changed on a simulated fill that has an order such that the
        ' snapshot flag changed on the fill, make sure that the snapshot flag gets changed appropriately
        ' on the order as well.  Tim ran into a case where he changed a snapshot fill to history, but the
        ' order stayed snapshot which caused the fill to be counted as both history and snapshot...
        If (Fill.OrderID > 0) And (g.Broker.IsLiveAccount(Fill.Broker) = False) Then
            Set Order = New cPtOrder
            If Order.Load(Fill.OrderID) Then
                If Order.IsSnapshot <> Fill.IsSnapshot Then
                    If Order.IsSnapshot = True Then
                        g.Broker.BrokerDebug Fill.Broker, "Order '" & Order.OrderText(True, True, True) & "' has been changed from snapshot to history because fill just changed"
                    Else
                        g.Broker.BrokerDebug Fill.Broker, "Order '" & Order.OrderText(True, True, True) & "' has been changed from history to snapshot because fill just changed"
                    End If
                    
                    Order.IsSnapshot = Fill.IsSnapshot
                    Order.Save
                End If
            End If
        End If
        
        If g.Broker.DateIsSnapshotForFill(Fill) Then
            g.Broker.AddFill Fill, False
        Else
            bRebuildHistory = True
        End If
    Next lIndex
    
    g.Broker.BrokerInfo(g.Broker.AccountTypeForID(m.lAccountID)).RebuildFillSummaries m.strAccountNumber, m.strSymbol, bRebuildHistory

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPositionFix.Save"
    
End Sub

